"""
prg_opf.constraints — OPF constraint formulation
==================================================
Builds all constraints and the objective function for the linearized
Power Router Grid optimal power flow problem.
"""

from pyomo.environ import Constraint, ConstraintList, Objective, minimize


# ── Port loss constraints (Big-M decomposition) ─────────────────────────

def add_port_loss_constraints(model):
    """P = P_POS - P_NEG decomposition with linearized converter losses."""

    def port_loss_rule1(model, pr, port):
        return model.P[port] == model.P_POS[port] - model.P_NEG[port]
    model.port_loss_rule1 = Constraint(model.PR_PORT, rule=port_loss_rule1)

    def port_loss_rule2(model, pr, port):
        return model.P_POS[port] <= model.y[port] * model.M
    model.port_loss_rule2 = Constraint(model.PR_PORT, rule=port_loss_rule2)

    def port_loss_rule3(model, pr, port):
        return model.P_NEG[port] <= (1 - model.y[port]) * model.M
    model.port_loss_rule3 = Constraint(model.PR_PORT, rule=port_loss_rule3)

    def port_loss_rule4(model, pr, port):
        return model.P_LOSS_POS[port] == model.y[port] * model.port_loss_c0[pr, port] + model.P_POS[port] * model.port_loss_c1[pr, port]
    model.port_loss_rule4 = Constraint(model.PR_PORT, rule=port_loss_rule4)

    def port_loss_rule5(model, pr, port):
        return model.P_LOSS_NEG[port] == (1 - model.y[port]) * model.port_loss_c0[pr, port] + model.P_NEG[port] * model.port_loss_c1[pr, port]
    model.port_loss_rule5 = Constraint(model.PR_PORT, rule=port_loss_rule5)

    def port_loss_rule6(model, pr, port):
        return model.P_LOSS[port] == model.P_LOSS_POS[port] + model.P_LOSS_NEG[port]
    model.port_loss_rule6 = Constraint(model.PR_PORT, rule=port_loss_rule6)


# ── Power flow & voltage constraints (AC lines) ─────────────────────────

def add_power_flow_constraints(model):
    """Active/reactive power flow, voltage balance, and SOCP relaxation."""

    def active_power_balance_rule(model, pr):
        ports = [p for p in model.PR_PORT if p[0] == pr]
        return sum(model.P[p[1]] for p in ports) == -sum(model.P_LOSS[p[1]] for p in ports)
    model.active_power_balance_rule = Constraint(model.PR, rule=active_power_balance_rule)

    def active_power_flow_rule(model, port1, port2):
        return -model.P[port2] == model.P[port1] - model.A[port1, port2] * model.line_R[port1, port2]
    model.active_power_flow_constraint = Constraint(model.LINES, rule=active_power_flow_rule)

    def reactive_power_flow_rule(model, port1, port2):
        return -model.Q[port2] == model.Q[port1] - model.A[port1, port2] * model.line_X[port1, port2]
    model.reactive_power_flow_constraint = Constraint(model.LINES, rule=reactive_power_flow_rule)

    def voltage_balance_rule(model, port1, port2):
        return model.V[port2] == model.V[port1] - \
            2 * (model.P[port1] * model.line_R[port1, port2] + model.Q[port1] * model.line_X[port1, port2]) \
            + model.A[port1, port2] * (model.line_R[port1, port2] ** 2 + model.line_X[port1, port2] ** 2)
    model.voltage_balance_constraint = Constraint(model.LINES, rule=voltage_balance_rule)

    # SOCP relaxation — makes the problem convex
    def current_power_relation(model, port1, port2):
        return (2 * model.P[port1]) ** 2 + (2 * model.Q[port1]) ** 2 + \
            (model.A[port1, port2] - model.V[port1]) ** 2 - \
            (model.A[port1, port2] + model.V[port1]) ** 2 <= 0
    model.current_power_relation = Constraint(model.LINES, rule=current_power_relation)

    def voltage_control_rule(model, pr, port):
        return model.V[port] == model.v_port_voltage[pr, port]
    model.voltage_control_rule = Constraint(model.V_PORT, rule=voltage_control_rule)


# ── Terminal port constraints ────────────────────────────────────────────

def add_terminal_constraints(model, enable_constraints):
    """Fix P and Q at terminal ports (on PRs and on buses)."""

    if enable_constraints['terminal_pr']:
        def terminal_port_P_rule(model, pr, port):
            return model.P[port] == model.terminal_port_P[pr, port] * model.S_ref
        model.terminal_port_P_rule = Constraint(model.TERMINAL_PORT, rule=terminal_port_P_rule)

        def terminal_port_Q_rule(model, pr, port):
            return model.Q[port] == model.terminal_port_Q[pr, port] * model.S_ref
        model.terminal_port_Q_rule = Constraint(model.TERMINAL_PORT, rule=terminal_port_Q_rule)

    if enable_constraints['terminal_bus']:
        def terminal_port_P_bus_rule(model, pr, port):
            return model.P[port] == model.terminal_port_bus_P[pr, port] * model.S_ref
        model.terminal_port_P_bus_rule = Constraint(model.TERMINAL_PORT_BUS, rule=terminal_port_P_bus_rule)

        def terminal_port_Q_bus_rule(model, pr, port):
            return model.Q[port] == model.terminal_port_bus_Q[pr, port] * model.S_ref
        model.terminal_port_Q_bus_rule = Constraint(model.TERMINAL_PORT_BUS, rule=terminal_port_Q_bus_rule)


# ── Reactive power forcing (Q = 0 on PQ / slack / ext_grid ports) ───────

def _compute_q_exempt(input_data):
    """
    Identify receiving-end ports that must be exempted from Q=0 forcing
    to avoid over-constraining lines where both ends would have Q=0
    (which would force A=0, making the line dead).
    """
    q_zero = set()
    for _, p in input_data['sets']['PQ_PORT']:
        q_zero.add(p)
    for _, p in input_data['sets']['EXT_GRID']:
        q_zero.add(p)
    for _, p in input_data['sets']['SLACK_PORT']:
        q_zero.add(p)
    for (pr, p), q_val in input_data.get('terminal_port_q', {}).items():
        if q_val == 0:
            q_zero.add(p)

    q_exempt = set()
    all_lines = list(input_data['ac_lines'].keys()) + list(input_data['dc_lines'].keys())
    for p1, p2 in all_lines:
        if p1 in q_zero and p2 in q_zero:
            q_exempt.add(p2)
    return q_exempt


def add_reactive_constraints(model, q_exempt):
    """Force Q=0 on PQ, slack, and ext_grid ports (with exemptions)."""

    def reactive_power_flow_PQ(model, pr, p):
        if p in q_exempt:
            return Constraint.Skip
        return model.Q[p] == 0
    model.reactive_power_flow_PQ = Constraint(model.PQ_PORT, rule=reactive_power_flow_PQ)

    def reactive_power_flow_forced(model):
        return [model.Q[port] == 0 for (pr, port) in model.SLACK_PORT]
    model.reactive_power_flow_forced = ConstraintList(rule=reactive_power_flow_forced)

    def reactive_power_flow_ext_grid(model, pr, p):
        if p in q_exempt:
            return Constraint.Skip
        return model.Q[p] == 0
    model.reactive_power_ext_grid = Constraint(model.EXT_GRID, rule=reactive_power_flow_ext_grid)


# ── Bus constraints ──────────────────────────────────────────────────────

def add_bus_constraints(model):
    """Power balance, voltage equality, and zero-loss at bus ports."""

    def active_power_balance_bus_rule(model, bus):
        ports = [p for p in model.BUS_PORT if p[0] == bus]
        return sum(model.P[p[1]] for p in ports) == 0
    model.active_power_balance_bus_rule = Constraint(model.BUS, rule=active_power_balance_bus_rule)

    def reactive_power_balance_bus_rule(model, bus):
        ports = [p for p in model.BUS_PORT if p[0] == bus]
        return sum(model.Q[p[1]] for p in ports) == 0
    model.reactive_power_balance_bus_rule = Constraint(model.BUS, rule=reactive_power_balance_bus_rule)

    def voltage_bus_rule(model, bus, port1, port2):
        return model.V[port1] == model.V[port2]
    model.voltage_bus_constraint = Constraint(model.BUS_PAIRS, rule=voltage_bus_rule)

    def port_loss_rule_bus(model, bus, port):
        return model.P_LOSS[port] == 0
    model.port_loss_rule_bus = Constraint(model.BUS_PORT, rule=port_loss_rule_bus)

    def port_loss_rule_bus2(model, bus, port):
        return model.P_LOSS_POS[port] == model.P_LOSS_NEG[port]
    model.port_loss_rule_bus2 = Constraint(model.BUS_PORT, rule=port_loss_rule_bus2)

    def port_loss_rule_bus3(model, bus, port):
        return model.P_POS[port] == model.P_NEG[port]
    model.port_loss_rule_bus3 = Constraint(model.BUS_PORT, rule=port_loss_rule_bus3)

    def port_loss_rule_bus4(model, bus, port):
        return model.P_LOSS_POS[port] == 0
    model.port_loss_rule_bus4 = Constraint(model.BUS_PORT, rule=port_loss_rule_bus4)

    def port_loss_rule_bus5(model, bus, port):
        return model.P_POS[port] == 0
    model.port_loss_rule_bus5 = Constraint(model.BUS_PORT, rule=port_loss_rule_bus5)


# ── DC line constraints ──────────────────────────────────────────────────

def add_dc_constraints(model):
    """DC line power flow and voltage balance (only if DC lines exist)."""

    # IMPORTANT: V for DC ports is in kV (not kV²)
    def active_power_flow_rule_DC1(model, port1, port2):
        return model.P[port1] == model.V[port1] * model.A_DC[port1, port2]
    model.active_power_flow_rule_DC1 = Constraint(model.DC_LINES, rule=active_power_flow_rule_DC1)

    def active_power_flow_rule_DC2(model, port1, port2):
        return model.P[port2] == -(model.V[port2] * model.A_DC[port1, port2])
    model.active_power_flow_rule_DC2 = Constraint(model.DC_LINES, rule=active_power_flow_rule_DC2)

    def reactive_power_flow_rule_DC(model, port1, port2):
        return model.Q[port1] == model.Q[port2]
    model.reactive_power_flow_rule_DC = Constraint(model.DC_LINES, rule=reactive_power_flow_rule_DC)

    def voltage_balance_rule_DC(model, port1, port2):
        return model.V[port2] == model.V[port1] - \
            model.dc_line_R[port1, port2] * model.A_DC[port1, port2]
    model.voltage_balance_rule_DC = Constraint(model.DC_LINES, rule=voltage_balance_rule_DC)


# ── Objective function ───────────────────────────────────────────────────

def add_objective(model, has_dc):
    """Minimize total losses (AC line + converter + DC line)."""

    if has_dc:
        def obj_rule(model):
            return (
                sum(model.line_R[p1, p2] * model.A[p1, p2]
                    for (p1, p2) in model.LINES)
                + sum(model.P_LOSS[port[1]] for port in model.PR_PORT)
                + sum(model.dc_line_R[p1, p2] * model.A_DC[p1, p2] ** 2
                      for (p1, p2) in model.DC_LINES)
            )
        model.obj = Objective(rule=obj_rule, sense=minimize)
    else:
        def obj_rule(model):
            return (
                sum(model.line_R[p1, p2] * model.A[p1, p2]
                    for (p1, p2) in model.LINES)
                + sum(model.P_LOSS[port[1]] for port in model.PR_PORT)
            )
        model.obj = Objective(rule=obj_rule, sense=minimize)


# ── Orchestrator ─────────────────────────────────────────────────────────

def build_formulation(model, enable_constraints, input_data):
    """
    Build the complete OPF formulation by calling all constraint builders.

    Parameters
    ----------
    model : pyomo.environ.AbstractModel
        A model with sets, parameters and variables already declared.
    enable_constraints : dict
        Keys: 'terminal_pr', 'terminal_bus', 'dc_lines' (booleans).
    input_data : dict
        The parsed input data dictionary (used for Q-exempt computation).

    Returns
    -------
    model : the same model object, with all constraints attached.
    """
    q_exempt = _compute_q_exempt(input_data)

    add_port_loss_constraints(model)
    add_power_flow_constraints(model)
    add_terminal_constraints(model, enable_constraints)
    add_reactive_constraints(model, q_exempt)
    add_bus_constraints(model)

    if enable_constraints['dc_lines']:
        add_dc_constraints(model)

    add_objective(model, enable_constraints['dc_lines'])

    return model
