"""
prg_opf.model — Pyomo model construction
==========================================
Defines sets, parameters, and decision variables for the linearized OPF.
"""

from pyomo.environ import (
    AbstractModel, Set, Param, Var, NonNegativeReals, Binary,
)


def define_sets(model, input_data):
    """Declare all Pyomo sets from the parsed input data."""

    sets = input_data['sets']

    model.PR = Set(initialize=sets['PR'])
    model.PORT = Set(initialize=sets['PORT'])
    model.PR_PORT = Set(dimen=2, within=model.PR * model.PORT, initialize=sets['PR_PORT'])
    model.SLACK_PORT = Set(dimen=2, within=model.PR_PORT, initialize=sets['SLACK_PORT'])
    model.PQ_PORT = Set(dimen=2, within=model.PR_PORT, initialize=sets['PQ_PORT'])
    model.EXT_GRID = Set(dimen=2, within=model.PR_PORT, initialize=sets['EXT_GRID'])
    model.BUS = Set(initialize=sets['BUS'])
    model.BUS_PORT = Set(dimen=2, within=model.BUS * model.PORT, initialize=sets['BUS_PORT'])

    # Auxiliary set: all port-pairs within each bus (for voltage equality)
    def bus_pairs_init(model):
        pairs = []
        for bus in model.BUS:
            ports = [p[1] for p in model.BUS_PORT if p[0] == bus]
            for i in range(len(ports)):
                for j in range(i + 1, len(ports)):
                    pairs.append((bus, ports[i], ports[j]))
        return pairs

    model.BUS_PAIRS = Set(initialize=bus_pairs_init, dimen=3)

    return model


def define_parameters(model, input_data):
    """Declare all model parameters and line sets."""

    params = input_data['params']

    # AC line parameters
    model.LINES = Set(initialize=list(input_data['ac_lines'].keys()))
    model.line_R = Param(model.LINES, initialize={k: v['R'] for k, v in input_data['ac_lines'].items()})
    model.line_X = Param(model.LINES, initialize={k: v['X'] for k, v in input_data['ac_lines'].items()})
    model.line_smax = Param(model.LINES, initialize={k: v['Smax'] for k, v in input_data['ac_lines'].items()})

    # DC line parameters
    has_dc = bool(input_data['dc_lines'])
    if has_dc:
        model.DC_LINES = Set(initialize=list(input_data['dc_lines'].keys()))
        model.dc_line_R = Param(model.DC_LINES, initialize={k: v['R'] for k, v in input_data['dc_lines'].items()})
        model.dc_line_X = Param(model.DC_LINES, initialize={k: v['X'] for k, v in input_data['dc_lines'].items()})
        model.dc_line_smax = Param(model.DC_LINES, initialize={k: v['Smax'] for k, v in input_data['dc_lines'].items()})

    # Terminal ports on PRs
    has_terminal_pr = bool(input_data['terminal_ports'])
    if has_terminal_pr:
        model.TERMINAL_PORT = Set(dimen=2, within=model.PR_PORT,
                                  initialize=input_data['terminal_ports'])
        model.terminal_port_P = Param(model.TERMINAL_PORT,
                                      initialize=input_data['terminal_port_p'])
        model.terminal_port_Q = Param(model.TERMINAL_PORT,
                                      initialize=input_data['terminal_port_q'])

    # Terminal ports on Buses
    has_terminal_bus = bool(input_data['terminal_ports_bus'])
    if has_terminal_bus:
        model.TERMINAL_PORT_BUS = Set(dimen=2, within=model.BUS_PORT,
                                      initialize=input_data['terminal_ports_bus'])
        model.terminal_port_bus_P = Param(model.TERMINAL_PORT_BUS,
                                          initialize=input_data['terminal_port_bus_p'])
        model.terminal_port_bus_Q = Param(model.TERMINAL_PORT_BUS,
                                          initialize=input_data['terminal_port_bus_q'])

    # Voltage-controlled ports
    model.V_PORT = Set(dimen=2, within=model.PR_PORT, initialize=input_data['v_ports'])
    model.v_port_voltage = Param(model.V_PORT, initialize=input_data['v_port_values'])

    # Optional fixed setpoints for PQ ports
    has_pq_p_set = bool(input_data['pq_port_p_setpoints'])
    if has_pq_p_set:
        model.PQ_PORT_P_SET = Set(
            dimen=2,
            within=model.PQ_PORT,
            initialize=list(input_data['pq_port_p_setpoints'].keys()),
        )
        model.pq_port_P_setpoint = Param(
            model.PQ_PORT_P_SET,
            initialize=input_data['pq_port_p_setpoints'],
        )

    has_pq_q_set = bool(input_data['pq_port_q_setpoints'])
    if has_pq_q_set:
        model.PQ_PORT_Q_SET = Set(
            dimen=2,
            within=model.PQ_PORT,
            initialize=list(input_data['pq_port_q_setpoints'].keys()),
        )
        model.pq_port_Q_setpoint = Param(
            model.PQ_PORT_Q_SET,
            initialize=input_data['pq_port_q_setpoints'],
        )

    # Base parameters (voltage = V² as part of the linearization)
    model.V_ref = Param(initialize=params['Vbase_squared'])
    model.S_ref = Param(initialize=params['Sbase'])

    # Auto-scale BigM: M only needs to exceed the maximum |P| at any port.
    # A value much larger than the power scale degrades Gurobi numerics.
    power_magnitudes = (
        [abs(v) for v in input_data.get('terminal_port_p', {}).values()]
        + [abs(v) for v in input_data.get('terminal_port_bus_p', {}).values()]
        + [abs(v) for v in input_data['pq_port_p_setpoints'].values()]
        + [v['Smax'] for v in input_data['ac_lines'].values()]
        + [v['Smax'] for v in input_data['dc_lines'].values()]
    )
    max_power = max(power_magnitudes) if power_magnitudes else 1.0
    data_M = max(max_power * 10, 100)          # 10× margin, floor 100
    user_M = params['BigM']
    if user_M > data_M:
        print(f'NOTE: BigM={user_M:.0e} >> power scale (~{max_power:.1f}). '
              f'Auto-capping to {data_M:.0f} for numerical stability.')
        params['BigM'] = data_M
    model.M = Param(initialize=params['BigM'])

    # Per-port loss coefficients
    # Negative c0 values make P_LOSS infeasible at P=0 (any port type),
    # so clamp them to 0 for a physically consistent nonnegative loss model.
    c0_raw = input_data['port_loss_c0']
    c0_clamped = {k: max(0.0, float(v)) for k, v in c0_raw.items()}
    n_clamped = sum(1 for k in c0_raw if c0_raw[k] < 0)
    if n_clamped > 0:
        print(f'NOTE: Clamped {n_clamped} negative port_loss_c0 value(s) to 0.0 '
              'to avoid infeasibility at P=0.')
    model.port_loss_c0 = Param(model.PR_PORT, initialize=c0_clamped)
    model.port_loss_c1 = Param(model.PR_PORT, initialize=input_data['port_loss_c1'])

    # Ports with a forced P = 0 setpoint: the Big-M loss model is bypassed for
    # these ports (c0 < 0 would make P_LOSS_POS or P_LOSS_NEG go negative).
    zero_p_ports = set()
    for (_, port), val in input_data['pq_port_p_setpoints'].items():
        if val == 0.0:
            zero_p_ports.add(port)
    for (_, port), val in input_data.get('terminal_port_p', {}).items():
        if val == 0.0:
            zero_p_ports.add(port)
    for (_, port), val in input_data.get('terminal_port_bus_p', {}).items():
        if val == 0.0:
            zero_p_ports.add(port)
    zero_p_ports = sorted(zero_p_ports)
    has_zero_p_ports = bool(zero_p_ports)
    if has_zero_p_ports:
        model.ZERO_P_PORT = Set(within=model.PORT, initialize=zero_p_ports)

    enable_constraints = {
        'terminal_pr': has_terminal_pr,
        'terminal_bus': has_terminal_bus,
        'dc_lines': has_dc,
        'pq_p_set': has_pq_p_set,
        'pq_q_set': has_pq_q_set,
        'zero_p_ports': has_zero_p_ports,
    }

    return model, enable_constraints


def define_variables(model, enable_constraints):
    """Declare all decision variables."""

    # Active power per port
    model.P = Var(model.PORT)
    model.P_POS = Var(model.PORT, within=NonNegativeReals)
    model.P_NEG = Var(model.PORT, within=NonNegativeReals)

    # Reactive power per port
    model.Q = Var(model.PORT)

    # AC line current (squared)
    model.A = Var(model.LINES, within=NonNegativeReals)

    # DC line current
    if enable_constraints['dc_lines']:
        model.A_DC = Var(model.DC_LINES)

    # Port losses
    model.P_LOSS = Var(model.PORT, within=NonNegativeReals)
    model.P_LOSS_POS = Var(model.PORT, within=NonNegativeReals)
    model.P_LOSS_NEG = Var(model.PORT, within=NonNegativeReals)

    # Voltage (squared for AC, linear for DC)
    model.V = Var(model.PORT, within=NonNegativeReals)

    # Auxiliary variables for SOC reformulation of A*V ≥ P²+Q²
    # w = (A + V)/2,  z = (A - V)/2  →  A*V = w² - z²
    # Then: P²+Q²+z² ≤ w²  is a standard second-order cone.
    model.soc_w = Var(model.LINES, within=NonNegativeReals)
    model.soc_z = Var(model.LINES)

    # Binary variable for Big-M loss decomposition
    model.y = Var(model.PORT, within=Binary)

    return model
