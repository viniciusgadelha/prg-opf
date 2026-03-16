"""
prg_opf.sensitivity — Sensitivity analysis across timesteps
============================================================
Reads a sensitivity-input Excel file (with a ``Timestep`` column in the
Ports and Lines sheets) and the base ``input.xlsx``, runs the OPF for every
timestep, and writes per-port / per-line results plus a summary sheet
into ``sens_results_.xlsx``.
"""

from __future__ import annotations

import copy
import math
import os
import time

import pandas as pd
from pyomo.environ import AbstractModel, value

from prg_opf.io import load_input_excel
from prg_opf.model import define_sets, define_parameters, define_variables
from prg_opf.constraints import build_formulation
from prg_opf.solver import run_optimization
from prg_opf.results import _get_dc_voltage_ports


# ─── helpers ──────────────────────────────────────────────────────────────

def _override_dict_by_port(d: dict, port_id: int, val):
    """Set *val* for the first key ``(owner, port_id)`` found in *d*."""
    for key in d:
        if key[1] == port_id:
            d[key] = val
            return


def _override_or_add_bus_terminal(data: dict, port_id: int, pq: str, val):
    """Override — or insert — a bus-terminal P/Q entry."""
    dkey = 'terminal_port_bus_p' if pq == 'P' else 'terminal_port_bus_q'
    target = data[dkey]
    for key in target:
        if key[1] == port_id:
            target[key] = val
            return
    # Not found → add, looking up the bus that owns this port
    for bus_id, pid in data['sets']['BUS_PORT']:
        if int(pid) == port_id:
            target[(bus_id, port_id)] = val
            if (bus_id, port_id) not in data['terminal_ports_bus']:
                data['terminal_ports_bus'].append((bus_id, port_id))
            return


def _override_v_setpoint(data: dict, port_id: int, val):
    """Override voltage setpoint for any port that has one, or add it."""
    for key in data['v_port_values']:
        if key[1] == port_id:
            data['v_port_values'][key] = val
            return
    # Port not yet voltage-controlled → find its PR/bus owner and add
    for owner, pid in data['sets']['PR_PORT']:
        if int(pid) == port_id:
            data['v_port_values'][(owner, port_id)] = val
            if (owner, port_id) not in data['v_ports']:
                data['v_ports'].append((owner, port_id))
            return


def _apply_timestep_overrides(base_data: dict,
                              sens_ports_df: pd.DataFrame,
                              sens_lines_df: pd.DataFrame | None,
                              timestep) -> dict:
    """Deep-copy *base_data* and apply the sens_input overrides for *timestep*."""
    data = copy.deepcopy(base_data)

    # ── Port overrides ────────────────────────────────────────────────
    ts_rows = sens_ports_df[sens_ports_df['Timestep'] == timestep]

    for _, row in ts_rows.iterrows():
        port_id = int(row['Port'])
        ptype = str(row['Type']).strip().lower()

        # Voltage setpoint (any type can carry one)
        if pd.notna(row.get('V_setpoint')):
            _override_v_setpoint(data, port_id, float(row['V_setpoint']))

        if ptype == 'v-f':
            if pd.notna(row.get('Q_setpoint')):
                _override_dict_by_port(data.get('pq_port_q_setpoints', {}),
                                       port_id, float(row['Q_setpoint']))

        elif ptype == 'terminal_bus':
            if pd.notna(row.get('P_setpoint')):
                _override_or_add_bus_terminal(data, port_id, 'P',
                                              float(row['P_setpoint']))
            if pd.notna(row.get('Q_setpoint')):
                _override_or_add_bus_terminal(data, port_id, 'Q',
                                              float(row['Q_setpoint']))

        elif ptype == 'terminal':
            if pd.notna(row.get('P_setpoint')):
                _override_dict_by_port(data['terminal_port_p'], port_id,
                                       float(row['P_setpoint']))
            if pd.notna(row.get('Q_setpoint')):
                _override_dict_by_port(data['terminal_port_q'], port_id,
                                       float(row['Q_setpoint']))

        elif ptype == 'pq':
            if pd.notna(row.get('P_setpoint')):
                _override_dict_by_port(data['pq_port_p_setpoints'], port_id,
                                       float(row['P_setpoint']))
            if pd.notna(row.get('Q_setpoint')):
                _override_dict_by_port(data['pq_port_q_setpoints'], port_id,
                                       float(row['Q_setpoint']))

    # ── Line parameter overrides ──────────────────────────────────────
    if sens_lines_df is not None and not sens_lines_df.empty:
        lt_rows = sens_lines_df[sens_lines_df['Timestep'] == timestep]
        for _, row in lt_rows.iterrows():
            port_str = str(row['Port']).strip()
            parts = port_str.split(',')
            p1, p2 = int(parts[0].strip()), int(parts[1].strip())
            key = (p1, p2)
            for lines_dict in (data['ac_lines'], data['dc_lines']):
                if key in lines_dict:
                    if pd.notna(row.get('R')):
                        lines_dict[key]['R'] = float(row['R'])
                    if pd.notna(row.get('X')):
                        lines_dict[key]['X'] = float(row['X'])
                    if pd.notna(row.get('Smax')):
                        lines_dict[key]['Smax'] = float(row['Smax'])

    return data


# ─── main entry point ────────────────────────────────────────────────────

def run_sensitivity(base_input_file: str,
                    sens_input_file: str,
                    output_file: str,
                    solver: str = 'gurobi',
                    time_limit: int = 600,
                    verbose: bool = True):
    """
    Run the OPF for every timestep defined in *sens_input_file*, collecting
    V, P, Q, P_LOSS and line-loss results into *output_file*.

    Parameters
    ----------
    base_input_file : path to the standard ``input.xlsx`` (full topology).
    sens_input_file : path to ``sens_input.xlsx`` (Timestep-varying setpoints).
    output_file     : path to the template ``sens_results_.xlsx`` (overwritten).
    solver          : solver name passed to ``run_optimization``.
    time_limit      : per-timestep solver time limit (seconds).
    verbose         : print progress per timestep.

    Returns
    -------
    output_file : str
    """
    t0 = time.time()

    # 1. Load base input (full topology + default setpoints)
    base_data = load_input_excel(base_input_file)

    # 2. Sensitivity overrides (time-varying)
    sens_xls = pd.ExcelFile(sens_input_file)

    sens_ports_df = pd.read_excel(sens_xls, sheet_name='Ports')
    sens_ports_df.columns = [c.strip() for c in sens_ports_df.columns]

    sens_lines_df = None
    if 'Lines' in sens_xls.sheet_names:
        sens_lines_df = pd.read_excel(sens_xls, sheet_name='Lines')
        sens_lines_df.columns = [c.strip() for c in sens_lines_df.columns]

    timesteps = sorted(sens_ports_df['Timestep'].unique())

    # 3. Read template to discover expected port and line indices
    tpl = pd.ExcelFile(output_file)
    v_tpl = pd.read_excel(tpl, sheet_name='V (kV)')
    port_list = v_tpl['Port/Timestep'].tolist()

    # Line indices: all AC + DC lines from base topology
    # (ensures DC line losses are written, not just template lines)
    line_tuples = list(base_data['ac_lines'].keys()) + list(base_data['dc_lines'].keys())
    line_keys = [f"{p1},{p2}" for (p1, p2) in line_tuples]

    # 4. Prepare empty result containers (Port/Timestep + one col per timestep)
    ts_strs = [str(t) for t in timesteps]

    def _empty_df(index_col, index_vals):
        df = pd.DataFrame({index_col: index_vals})
        for c in ts_strs:
            df[c] = float('nan')
        return df

    v_df     = _empty_df('Port/Timestep', port_list)
    p_df     = _empty_df('Port/Timestep', port_list)
    q_df     = _empty_df('Port/Timestep', port_list)
    ploss_df = _empty_df('Port/Timestep', port_list)
    lloss_df = _empty_df('Port/Timestep', line_keys)

    dc_voltage_ports = _get_dc_voltage_ports(base_data)

    # Summary accumulators (one row per timestep)
    summary_rows = []

    # 5. Solve per timestep
    for t_idx, t in enumerate(timesteps):
        if verbose:
            print(f'\n{"=" * 60}')
            print(f'  Sensitivity timestep {t}  ({t_idx + 1}/{len(timesteps)})')
            print(f'{"=" * 60}')

        input_data = _apply_timestep_overrides(base_data, sens_ports_df,
                                               sens_lines_df, t)

        model = AbstractModel()
        model = define_sets(model, input_data)
        model, enable = define_parameters(model, input_data)
        model = define_variables(model, enable)
        model = build_formulation(model, enable, input_data)
        solution = run_optimization(model, solver=solver,
                                    time_limit=time_limit,
                                    verbose=verbose)

        tc = str(t)
        has_dc = hasattr(solution, 'A_DC')
        objective_mw = value(solution.obj)

        # Per-port results
        total_port_loss_kw = 0.0
        for port in port_list:
            pi = int(port)
            try:
                v_raw = value(solution.V[pi])
                if pi not in dc_voltage_ports:
                    v_raw = math.sqrt(max(v_raw, 0))
                v_df.loc[v_df['Port/Timestep'] == port, tc] = v_raw
            except (KeyError, ValueError):
                pass
            try:
                p_df.loc[p_df['Port/Timestep'] == port, tc] = value(solution.P[pi])
            except (KeyError, ValueError):
                pass
            try:
                q_df.loc[q_df['Port/Timestep'] == port, tc] = value(solution.Q[pi])
            except (KeyError, ValueError):
                pass
            try:
                ploss_val = value(solution.P_LOSS[pi])
                ploss_df.loc[ploss_df['Port/Timestep'] == port, tc] = ploss_val * 1000
                total_port_loss_kw += abs(ploss_val) * 1000
            except (KeyError, ValueError):
                pass

        # Per-line losses (separate AC and DC totals)
        total_ac_line_loss_kw = 0.0
        total_dc_line_loss_kw = 0.0
        for li, (p1, p2) in enumerate(line_tuples):
            try:
                if (p1, p2) in solution.A:
                    loss = (abs(value(solution.A[p1, p2]))
                            * value(solution.line_R[p1, p2]) * 1000)
                    total_ac_line_loss_kw += loss
                elif has_dc and (p1, p2) in solution.A_DC:
                    loss = (value(solution.dc_line_R[p1, p2])
                            * abs(value(solution.A_DC[p1, p2])) ** 2 * 1000)
                    total_dc_line_loss_kw += loss
                else:
                    continue
                lloss_df.loc[lloss_df['Port/Timestep'] == line_keys[li], tc] = loss
            except (KeyError, ValueError):
                pass

        # Also compute full AC / DC line losses (not just template lines)
        full_ac_loss_kw = 0.0
        for (lp1, lp2) in solution.A:
            try:
                full_ac_loss_kw += abs(value(solution.A[lp1, lp2])) * value(solution.line_R[lp1, lp2]) * 1000
            except (KeyError, ValueError):
                pass
        full_dc_loss_kw = 0.0
        if has_dc:
            for (lp1, lp2) in solution.A_DC:
                try:
                    full_dc_loss_kw += value(solution.dc_line_R[lp1, lp2]) * abs(value(solution.A_DC[lp1, lp2])) ** 2 * 1000
                except (KeyError, ValueError):
                    pass

        # Total injected power = sum of absolute P at ext-grid / slack ports
        total_gen_kw = 0.0
        for (_, pid) in input_data['sets'].get('EXT_GRID', []):
            try:
                total_gen_kw += abs(value(solution.P[int(pid)])) * 1000
            except (KeyError, ValueError):
                pass
        # Fallback: also include terminal port loads as absolute injection
        total_load_kw = 0.0
        for (_, pid) in input_data.get('terminal_ports', []):
            try:
                total_load_kw += abs(value(solution.P[int(pid)])) * 1000
            except (KeyError, ValueError):
                pass
        for (_, pid) in input_data.get('terminal_ports_bus', []):
            try:
                total_load_kw += abs(value(solution.P[int(pid)])) * 1000
            except (KeyError, ValueError):
                pass

        total_loss_kw = objective_mw * 1000
        ref_power_kw = total_gen_kw if total_gen_kw > 0 else total_load_kw
        pct = lambda x: (x / ref_power_kw * 100) if ref_power_kw > 1e-9 else 0.0

        summary_rows.append({
            'Timestep': t,
            'Objective (MW)': objective_mw,
            'Total Loss (kW)': total_loss_kw,
            'AC Line Loss (kW)': full_ac_loss_kw,
            'DC Line Loss (kW)': full_dc_loss_kw,
            'Port Loss (kW)': total_port_loss_kw,
            'AC Line Loss (%)': pct(full_ac_loss_kw),
            'DC Line Loss (%)': pct(full_dc_loss_kw),
            'Port Loss (%)': pct(total_port_loss_kw),
            'Total Loss (%)': pct(total_loss_kw),
            'Ref Power (kW)': ref_power_kw,
        })

        if verbose:
            r = summary_rows[-1]
            print(f'  Objective     : {objective_mw:.6f} MW')
            print(f'  AC line loss  : {full_ac_loss_kw:.2f} kW  ({r["AC Line Loss (%)"]:.2f}%)')
            print(f'  DC line loss  : {full_dc_loss_kw:.2f} kW  ({r["DC Line Loss (%)"]:.2f}%)')
            print(f'  Port loss     : {total_port_loss_kw:.2f} kW  ({r["Port Loss (%)"]:.2f}%)')
            print(f'  Total loss    : {total_loss_kw:.2f} kW  ({r["Total Loss (%)"]:.2f}%)')

    # 6. Build summary DataFrame
    summary_df = pd.DataFrame(summary_rows)

    # 7. Write results
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        v_df.to_excel(writer, sheet_name='V (kV)', index=False)
        p_df.to_excel(writer, sheet_name='P (MW)', index=False)
        q_df.to_excel(writer, sheet_name='Q (MVAR)', index=False)
        ploss_df.to_excel(writer, sheet_name='P_LOSS (kW)', index=False)
        lloss_df.to_excel(writer, sheet_name='L_LOSS (kW)', index=False)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

    elapsed = time.time() - t0
    if verbose:
        print(f'\n{"=" * 60}')
        print(f'Sensitivity analysis complete — {len(timesteps)} timesteps')
        print(f'Results saved to: {output_file}')
        print(f'Total time: {elapsed:.1f} s')
        print(f'{"=" * 60}')

    return output_file
