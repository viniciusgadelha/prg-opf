"""
prg_opf.io — Input data loading
================================
Reads a unified Excel workbook and returns the complete data dictionary
consumed by the model builder, constraint formulation, and plotter.
"""

import pandas as pd


def load_input_excel(filepath):
    """
    Load a unified input Excel file and return all data needed by the PRG optimization.

    Expected sheets:
        - "Power Routers": columns [PR, Port, Type, V_setpoint, P_setpoint, Q_setpoint]
            Type can be (case-insensitive):
                slack,ext_grid  – slack port connected to external grid (requires V_setpoint)
                slack           – slack port connected to lines (no V_setpoint)
                pq              – power-control port
                v-f             – voltage-frequency control port (requires V_setpoint)
                terminal        – terminal port with load/generation (requires P_setpoint, Q_setpoint)
            V_setpoint: voltage squared [kV²] for voltage-controlled ports (leave blank otherwise)
            P_setpoint, Q_setpoint: active/reactive power [MW/MVAR] for terminal ports (leave blank otherwise)

        - "Buses": columns [Bus, Port, P_setpoint, Q_setpoint]
            P_setpoint, Q_setpoint: active/reactive power for loads at bus (leave blank if none)

        - "AC Lines": columns [From_Port, To_Port, R, X, Smax]

        - "DC Lines": columns [From_Port, To_Port, R, X, Smax]  (can be empty or absent)

        - "Parameters" (optional): columns [name, value]
            Supported: Sbase, Vbase_squared, loss_c0, loss_c1, BigM

    Returns a dict with all parsed data ready for the optimization model.
    """
    sheets = pd.ExcelFile(filepath).sheet_names

    # --- Power Routers sheet ---
    pr_df = pd.read_excel(filepath, sheet_name='Power Routers')
    pr_df.columns = [c.strip() for c in pr_df.columns]

    # --- Buses sheet ---
    bus_df = pd.read_excel(filepath, sheet_name='Buses')
    bus_df.columns = [c.strip() for c in bus_df.columns]

    # --- AC Lines sheet ---
    ac_lines_df = pd.read_excel(filepath, sheet_name='AC Lines')
    ac_lines_df.columns = [c.strip() for c in ac_lines_df.columns]

    # --- DC Lines sheet (optional) ---
    dc_lines_df = pd.DataFrame()
    if 'DC Lines' in sheets:
        dc_lines_df = pd.read_excel(filepath, sheet_name='DC Lines')
        dc_lines_df.columns = [c.strip() for c in dc_lines_df.columns]
        dc_lines_df = dc_lines_df.dropna(how='all')

    # --- Parameters sheet (optional) ---
    params = {
        'Sbase': 1,
        'Vbase_squared': 36,
        'loss_c0': -5.82e-5,
        'loss_c1': 0.00154,
        'BigM': 999999999,
    }
    if 'Parameters' in sheets:
        params_df = pd.read_excel(filepath, sheet_name='Parameters')
        params_df.columns = [c.strip() for c in params_df.columns]
        for _, row in params_df.iterrows():
            name = str(row['name']).strip()
            if name in params:
                params[name] = float(row['value'])

    # --- Derive sets from Power Routers sheet ---
    pr_list = sorted(pr_df['PR'].unique().tolist())

    # All ports: PR ports + bus ports
    pr_ports = pr_df['Port'].unique().tolist()
    bus_ports = bus_df['Port'].unique().tolist()
    all_ports = sorted(set(pr_ports + bus_ports))

    pr_port_pairs = list(zip(pr_df['PR'].astype(int), pr_df['Port'].astype(int)))

    # Classify port types (Type can be comma-separated, e.g. 'slack, ext_grid')
    pr_df['Type'] = pr_df['Type'].str.strip().str.lower()
    slack_ports = [(int(r['PR']), int(r['Port'])) for _, r in pr_df.iterrows() if 'slack' in r['Type']]
    pq_ports = [(int(r['PR']), int(r['Port'])) for _, r in pr_df.iterrows() if 'pq' in r['Type']]
    ext_grid_ports = [(int(r['PR']), int(r['Port'])) for _, r in pr_df.iterrows() if 'ext_grid' in r['Type']]

    # Buses
    bus_list = sorted(bus_df['Bus'].unique().tolist(), key=str)
    bus_port_pairs = list(zip(bus_df['Bus'], bus_df['Port'].astype(int)))

    # Voltage-controlled ports (non-empty V_setpoint in PR sheet)
    v_ports = []
    v_port_values = {}
    for _, r in pr_df.iterrows():
        if pd.notna(r.get('V_setpoint')) and r['V_setpoint'] != '':
            pr_id = int(r['PR'])
            port_id = int(r['Port'])
            v_ports.append((pr_id, port_id))
            v_port_values[(pr_id, port_id)] = float(r['V_setpoint'])

    # Terminal ports on PRs (Type == 'terminal')
    terminal_ports = []
    terminal_port_p = {}
    terminal_port_q = {}
    for _, r in pr_df.iterrows():
        if 'terminal' in r['Type']:
            pr_id = int(r['PR'])
            port_id = int(r['Port'])
            terminal_ports.append((pr_id, port_id))
            terminal_port_p[(pr_id, port_id)] = float(r['P_setpoint']) if pd.notna(r.get('P_setpoint')) else 0.0
            terminal_port_q[(pr_id, port_id)] = float(r.get('Q_setpoint', 0)) if pd.notna(r.get('Q_setpoint')) else 0.0

    # Terminal ports on Buses (non-empty P_setpoint)
    terminal_ports_bus = []
    terminal_port_bus_p = {}
    terminal_port_bus_q = {}
    for _, r in bus_df.iterrows():
        if pd.notna(r.get('P_setpoint')) and r['P_setpoint'] != '':
            bus_id = r['Bus']
            port_id = int(r['Port'])
            terminal_ports_bus.append((bus_id, port_id))
            terminal_port_bus_p[(bus_id, port_id)] = float(r['P_setpoint'])
            terminal_port_bus_q[(bus_id, port_id)] = float(r.get('Q_setpoint', 0)) if pd.notna(r.get('Q_setpoint')) else 0.0

    # AC Lines
    ac_lines = {}
    for _, r in ac_lines_df.iterrows():
        key = (int(r['From_Port']), int(r['To_Port']))
        ac_lines[key] = {'R': float(r['R']), 'X': float(r['X']), 'Smax': float(r['Smax'])}

    # DC Lines
    dc_lines = {}
    if not dc_lines_df.empty:
        for _, r in dc_lines_df.iterrows():
            key = (int(r['From_Port']), int(r['To_Port']))
            dc_lines[key] = {'R': float(r['R']), 'X': float(r['X']), 'Smax': float(r['Smax'])}

    return {
        'sets': {
            'PR': pr_list,
            'PORT': all_ports,
            'PR_PORT': pr_port_pairs,
            'SLACK_PORT': slack_ports,
            'PQ_PORT': pq_ports,
            'EXT_GRID': ext_grid_ports,
            'BUS': bus_list,
            'BUS_PORT': bus_port_pairs,
        },
        'v_ports': v_ports,
        'v_port_values': v_port_values,
        'terminal_ports': terminal_ports,
        'terminal_port_p': terminal_port_p,
        'terminal_port_q': terminal_port_q,
        'terminal_ports_bus': terminal_ports_bus,
        'terminal_port_bus_p': terminal_port_bus_p,
        'terminal_port_bus_q': terminal_port_bus_q,
        'ac_lines': ac_lines,
        'dc_lines': dc_lines,
        'params': params,
    }
