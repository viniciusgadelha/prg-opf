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


def define_parameters(model, input_data, K=1):
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

    # Base parameters (voltage = V² as part of the linearization)
    model.V_ref = Param(initialize=params['Vbase_squared'])
    model.S_ref = Param(initialize=params['Sbase'])
    model.M = Param(initialize=params['BigM'])

    # Loss coefficients
    model.loss_c0 = Param(initialize=params['loss_c0'])
    model.loss_c1 = Param(initialize=params['loss_c1'])

    # Selectable factor for power router losses scaling
    model.K = Param(initialize=K)

    enable_constraints = {
        'terminal_pr': has_terminal_pr,
        'terminal_bus': has_terminal_bus,
        'dc_lines': has_dc,
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

    # Binary variable for Big-M loss decomposition
    model.y = Var(model.PORT, within=Binary)

    return model
