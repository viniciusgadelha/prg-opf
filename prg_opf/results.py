"""
prg_opf.results — Results extraction and export
=================================================
Converts Pyomo solution variables into DataFrames and writes them
to an Excel workbook.
"""

import math
import os

import pandas as pd
from pyomo.environ import value


def _pyomo_var_to_df(var, col_name='solution'):
    """Convert a Pyomo indexed variable/constraint to a two-column DataFrame."""
    data = [(k, value(var[k])) for k in var]
    if data:
        return pd.DataFrame(data, columns=['index', col_name])
    return pd.DataFrame(columns=['index', col_name])


def export_results(solution, input_data=None, path=None, label=''):
    """
    Export optimisation results to an Excel workbook with *nodes* and *lines* sheets.

    Parameters
    ----------
    solution : pyomo.core.ConcreteModel
        The solved instance returned by ``run_optimization``.
    input_data : dict or None
        Parsed input data (currently unused, reserved for future enrichment).
    path : str or None
        Directory to write the results file into. Defaults to ``<cwd>/results/``.
    label : str
        Optional suffix appended to the filename (e.g. hour index for timeseries).

    Returns
    -------
    output_file : str
        Path to the written Excel file.
    """
    if path is None:
        path = os.path.join(os.getcwd(), 'results', '')

    print('\n####################################################')
    print('Saving and exporting results...')
    print('####################################################')

    objective = solution.obj()
    has_dc = hasattr(solution, 'A_DC')

    # Identify DC ports (voltage in kV, not kV²)
    dc_ports = set()
    if has_dc:
        for (p1, p2) in solution.DC_LINES:
            dc_ports.add(p1)
            dc_ports.add(p2)

    # --- Nodes sheet ---
    voltage_df = _pyomo_var_to_df(solution.V, 'V [kV]')
    # Convert AC voltages from kV² to kV (skip DC ports)
    for i, row in voltage_df.iterrows():
        try:
            port_id = int(row['index'])
        except (ValueError, TypeError):
            continue
        if port_id not in dc_ports:
            voltage_df.at[i, 'V [kV]'] = math.sqrt(max(row['V [kV]'], 0))

    active_powers_df = _pyomo_var_to_df(solution.P, 'P [MW]')
    active_powers_df.loc[len(active_powers_df)] = ['Objective', objective]
    react_powers_df = _pyomo_var_to_df(solution.Q, 'Q [MVAR]')
    losses_df = _pyomo_var_to_df(solution.P_LOSS, 'P_LOSS')
    p_loss_pos_df = _pyomo_var_to_df(solution.P_LOSS_POS, 'P_LOSS_POS')
    p_loss_neg_df = _pyomo_var_to_df(solution.P_LOSS_NEG, 'P_LOSS_NEG')

    # Coerce index columns to string for consistent merge
    for df in [voltage_df, active_powers_df, react_powers_df,
               losses_df, p_loss_pos_df, p_loss_neg_df]:
        df['index'] = df['index'].astype(str)

    nodes_df = voltage_df
    for df in [active_powers_df, react_powers_df, losses_df, p_loss_pos_df, p_loss_neg_df]:
        nodes_df = nodes_df.merge(df, on='index', how='outer')

    # --- Lines sheet ---
    current_df = _pyomo_var_to_df(solution.A, 'I [kA]')
    for i, row in current_df.iterrows():
        current_df.at[i, 'I [kA]'] = math.sqrt(max(row['I [kA]'], 0))

    dc_current_df = (
        _pyomo_var_to_df(solution.A_DC, 'I_DC [kA]')
        if has_dc
        else pd.DataFrame(columns=['index', 'I_DC [kA]'])
    )
    relaxation_df = _pyomo_var_to_df(solution.current_power_relation, 'relaxation')

    for df in [current_df, dc_current_df, relaxation_df]:
        df['index'] = df['index'].astype(str)

    lines_df = current_df
    for df in [dc_current_df, relaxation_df]:
        lines_df = lines_df.merge(df, on='index', how='outer')

    # Write Excel
    output_file = os.path.join(path, f'optimization_results_{label}.xlsx')
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        nodes_df.to_excel(writer, sheet_name='nodes', index=False)
        lines_df.to_excel(writer, sheet_name='lines', index=False)

    # Summary
    print(f'\n--- Results Summary ---')
    print(f'Objective (total losses): {objective:.6f}')
    print(f'Results saved to: {output_file}')
    print('End of optimization!')

    return output_file
