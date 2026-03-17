"""
prg_opf.slack_tree — Slack-tree enumeration and analysis
=========================================================
Reads the PRG input topology, builds a PR-level graph, enumerates all
spanning trees rooted at the external-grid PR, assigns port roles per tree,
plots them interactively (Plotly) and runs the OPF for each configuration.

Three rules must hold for every valid slack tree:
  1. Every PR has at least one slack port.
  2. Every AC line is connected to at least one v-f port.
  3. The slack ports form a connected path (tree) to the ext_grid.
"""

from __future__ import annotations

import copy
import math
import os
import time

import networkx as nx
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pyomo.environ import AbstractModel, value

from prg_opf.io import load_input_excel
from prg_opf.model import define_sets, define_parameters, define_variables
from prg_opf.constraints import build_formulation
from prg_opf.solver import run_optimization
from prg_opf.results import _get_dc_voltage_ports, _safe_value


# ─── Graph construction ──────────────────────────────────────────────────

def _build_port_to_pr(input_data: dict) -> dict[int, int]:
    """Map each port_id → PR_id from the PR_PORT set."""
    return {int(port): int(pr) for pr, port in input_data['sets']['PR_PORT']}


def _build_pr_graph(input_data: dict):
    """
    Build an undirected PR-level graph from the AC Lines sheet.

    Returns
    -------
    G : nx.Graph
        Nodes are PR ids.  Each edge carries attributes:
          from_port, to_port  – the AC line's port pair
    ext_grid_pr : int
        The PR that contains the ext_grid port.
    port_to_pr : dict[int, int]
    """
    port_to_pr = _build_port_to_pr(input_data)

    # Bus ports bridge two PRs — map bus ports to the PR reached through the bus
    bus_port_to_pr: dict[int, int] = {}
    bus_to_ports: dict = {}
    for bus, port in input_data['sets']['BUS_PORT']:
        bus_to_ports.setdefault(bus, set()).add(int(port))
    # For each bus, find which PR each bus-port connects to
    # A bus-port might itself be a PR port, or connect to a PR port via a line
    for bus, ports in bus_to_ports.items():
        for bp in ports:
            if bp in port_to_pr:
                bus_port_to_pr[bp] = port_to_pr[bp]

    def _port_owner(port_id: int) -> int | None:
        """Find the PR that owns a port (directly or via bus)."""
        if port_id in port_to_pr:
            return port_to_pr[port_id]
        if port_id in bus_port_to_pr:
            return bus_port_to_pr[port_id]
        # Walk through bus connections
        for bus, ports in bus_to_ports.items():
            if port_id in ports:
                for bp in ports:
                    if bp in port_to_pr:
                        return port_to_pr[bp]
        return None

    G = nx.Graph()
    G.add_nodes_from(input_data['sets']['PR'])

    # Track which ports form each edge (PR_from, PR_to) → list of (from_port, to_port)
    edge_lines: dict[tuple, list] = {}
    for (fp, tp) in input_data['ac_lines']:
        pr_from = _port_owner(fp)
        pr_to = _port_owner(tp)
        if pr_from is None or pr_to is None or pr_from == pr_to:
            continue
        key = (min(pr_from, pr_to), max(pr_from, pr_to))
        edge_lines.setdefault(key, []).append((fp, tp))

    for (pr_a, pr_b), lines in edge_lines.items():
        G.add_edge(pr_a, pr_b, ac_lines=lines)

    # Identify ext_grid PR
    ext_grid_pr = None
    for pr, port in input_data['sets']['EXT_GRID']:
        ext_grid_pr = int(pr)
        break

    return G, ext_grid_pr, port_to_pr


# ─── Spanning-tree enumeration (adapted from old code) ───────────────────

def _expand(G: nx.Graph, explored_nodes: frozenset, explored_edges: frozenset):
    """One BFS expansion step — yields all single-node extensions."""
    for v in explored_nodes:
        for u in nx.neighbors(G, v):
            if u not in explored_nodes:
                yield (
                    explored_nodes | frozenset([u]),
                    explored_edges | frozenset([(u, v), (v, u)]),
                )


def find_all_spanning_trees(G: nx.Graph, root: int) -> list[nx.Graph]:
    """
    Enumerate all spanning trees of an undirected graph by BFS expansion.

    Parameters
    ----------
    G    : undirected NetworkX graph (PR-level)
    root : ext_grid PR node

    Returns
    -------
    list of undirected NetworkX graphs, one per spanning tree
    """
    solutions = [(frozenset([root]), frozenset())]
    for _ in range(G.number_of_nodes() - 1):
        next_solutions = set()
        for nodes, edges in solutions:
            for expanded in _expand(G, nodes, edges):
                next_solutions.add(expanded)
        solutions = list(next_solutions)
    return [nx.from_edgelist(edges) for (nodes, edges) in solutions]


def _orient_tree(tree: nx.Graph, root: int) -> nx.DiGraph:
    """
    Orient an undirected spanning tree so that every non-root node has
    exactly one incoming edge (from its parent toward the root).

    The edge direction encodes: source → target means "source is closer to
    root" (parent → child).  The child side receives a slack port; the
    parent side receives a v-f port.  But from the *slack-tree connectivity*
    perspective: the child port is slack and the parent port is v-f.
    """
    DG = nx.DiGraph()
    DG.add_nodes_from(tree.nodes())
    visited = {root}
    queue = [root]
    while queue:
        node = queue.pop(0)
        for nbr in tree.neighbors(node):
            if nbr not in visited:
                DG.add_edge(node, nbr)  # parent → child
                visited.add(nbr)
                queue.append(nbr)
    return DG


def _find_pq_lines(G_full: nx.Graph, tree: nx.DiGraph) -> list[tuple]:
    """Return edges in the full graph that are NOT in the spanning tree."""
    tree_edges_undirected = set()
    for u, v in tree.edges():
        tree_edges_undirected.add((min(u, v), max(u, v)))
    pq_edges = []
    for u, v in G_full.edges():
        key = (min(u, v), max(u, v))
        if key not in tree_edges_undirected:
            pq_edges.append((u, v))
    return pq_edges


# ─── Port-role assignment ────────────────────────────────────────────────

def _assign_port_roles(input_data: dict,
                       G_full: nx.Graph,
                       directed_tree: nx.DiGraph,
                       pq_edges: list[tuple],
                       port_to_pr: dict[int, int],
                       ext_grid_pr: int) -> dict:
    """
    Deep-copy input_data and reassign port types for a given spanning tree.

    For each AC line (between two PRs):
      - If the line is in the spanning tree (slack line):
          * The port on the child PR side → slack
          * The port on the parent PR side → v-f (with V_setpoint)
      - If the line is NOT in the spanning tree (PQ line):
          * Both ports → pq

    Terminal ports, ext_grid ports, and bus-terminal ports are unchanged.
    """
    data = copy.deepcopy(input_data)

    # Build lookup: which ports are terminals / ext_grid (immutable)
    immutable_ports = set()
    for pr, port in data['sets']['EXT_GRID']:
        immutable_ports.add(int(port))
    for pr, port in data.get('terminal_ports', []):
        immutable_ports.add(int(port))
    # Bus-terminal ports are identified by port_types
    for port_id, ptype in data['port_types'].items():
        if 'terminal' in ptype:
            immutable_ports.add(int(port_id))

    # Build directed tree edge set for quick lookup
    tree_edge_set = set()
    for u, v in directed_tree.edges():
        tree_edge_set.add((min(u, v), max(u, v)))

    pq_edge_set = set()
    for u, v in pq_edges:
        pq_edge_set.add((min(u, v), max(u, v)))

    # Compute depth of each PR in the directed tree (for PQ line v-f assignment)
    depth = {ext_grid_pr: 0}
    queue = [ext_grid_pr]
    while queue:
        node = queue.pop(0)
        for nbr in directed_tree.successors(node):
            if nbr not in depth:
                depth[nbr] = depth[node] + 1
                queue.append(nbr)

    # Collect the new port assignments
    new_slack_ports = []
    new_pq_ports = []
    new_v_ports = []
    new_v_values = {}
    new_pq_p_setpoints = {}
    new_pq_q_setpoints = {}
    new_port_types = dict(data['port_types'])

    # Keep ext_grid port as-is
    for pr, port in data['sets']['EXT_GRID']:
        new_slack_ports.append((int(pr), int(port)))

    # Keep terminal ports as-is
    for pr, port in data.get('terminal_ports', []):
        new_port_types[int(port)] = 'terminal'

    # Get a reference V_setpoint from the base data (use ext_grid's value)
    ref_v = None
    for (pr, port), v_val in data['v_port_values'].items():
        if (pr, port) in [(p, q) for p, q in data['sets']['EXT_GRID']]:
            ref_v = v_val
            break
    if ref_v is None:
        # Fallback: use any existing v_port_value
        for key, v_val in data['v_port_values'].items():
            ref_v = v_val
            break
    if ref_v is None:
        ref_v = 10000  # default

    # Process each AC line
    for (fp, tp), line_params in data['ac_lines'].items():
        fp, tp = int(fp), int(tp)

        # Find the PRs these ports belong to
        pr_fp = port_to_pr.get(fp)
        pr_tp = port_to_pr.get(tp)

        # If ports go through buses, find the PR endpoint
        if pr_fp is None:
            pr_fp = _find_pr_through_bus(fp, data)
        if pr_tp is None:
            pr_tp = _find_pr_through_bus(tp, data)

        if pr_fp is None or pr_tp is None:
            continue

        edge_key = (min(pr_fp, pr_tp), max(pr_fp, pr_tp))

        if edge_key in tree_edge_set:
            # This is a SLACK line (in the spanning tree)
            # Determine parent/child: parent→child in directed_tree
            if directed_tree.has_edge(pr_fp, pr_tp):
                parent_pr, child_pr = pr_fp, pr_tp
                parent_port, child_port = fp, tp
            elif directed_tree.has_edge(pr_tp, pr_fp):
                parent_pr, child_pr = pr_tp, pr_fp
                parent_port, child_port = tp, fp
            else:
                continue

            # Child-side port → slack (unless immutable)
            if child_port not in immutable_ports:
                pr_of_child = port_to_pr.get(child_port, child_pr)
                new_slack_ports.append((pr_of_child, child_port))
                new_port_types[child_port] = 'slack'

            # Parent-side port → v-f with V_setpoint (unless immutable)
            if parent_port not in immutable_ports:
                pr_of_parent = port_to_pr.get(parent_port, parent_pr)
                new_v_ports.append((pr_of_parent, parent_port))
                new_v_values[(pr_of_parent, parent_port)] = ref_v
                new_port_types[parent_port] = 'v-f'

        elif edge_key in pq_edge_set:
            # This is a PQ line (not in the spanning tree)
            # Rule 2: at least one port must be v-f.
            # The port on the PR closer to root (smaller depth) → v-f
            # The other port → pq
            depth_fp = depth.get(pr_fp, 999)
            depth_tp = depth.get(pr_tp, 999)
            if depth_fp <= depth_tp:
                vf_port, pq_port = fp, tp
                vf_pr, pq_pr = pr_fp, pr_tp
            else:
                vf_port, pq_port = tp, fp
                vf_pr, pq_pr = pr_tp, pr_fp

            if vf_port not in immutable_ports:
                pr_of_vf = port_to_pr.get(vf_port, vf_pr)
                new_v_ports.append((pr_of_vf, vf_port))
                new_v_values[(pr_of_vf, vf_port)] = ref_v
                new_port_types[vf_port] = 'v-f'

            if pq_port not in immutable_ports:
                pr_of_pq = port_to_pr.get(pq_port, pq_pr)
                new_pq_ports.append((pr_of_pq, pq_port))
                new_port_types[pq_port] = 'pq'

    # Now update the data dict
    data['sets']['SLACK_PORT'] = list(set(new_slack_ports))
    data['sets']['PQ_PORT'] = list(set(new_pq_ports))
    data['v_ports'] = list(set(new_v_ports))
    data['v_port_values'] = new_v_values
    data['pq_port_p_setpoints'] = new_pq_p_setpoints
    data['pq_port_q_setpoints'] = new_pq_q_setpoints
    data['port_types'] = new_port_types

    return data


def _find_pr_through_bus(port_id: int, data: dict) -> int | None:
    """Find the PR that a bus port connects to, traversing bus topology."""
    port_to_pr = {int(p): int(pr) for pr, p in data['sets']['PR_PORT']}
    if port_id in port_to_pr:
        return port_to_pr[port_id]
    bus_to_ports: dict = {}
    for bus, port in data['sets']['BUS_PORT']:
        bus_to_ports.setdefault(bus, set()).add(int(port))
    for bus, ports in bus_to_ports.items():
        if port_id in ports:
            for bp in ports:
                if bp in port_to_pr:
                    return port_to_pr[bp]
    return None


# ─── Validation ──────────────────────────────────────────────────────────

def _validate_tree_config(data: dict, tree_idx: int) -> list[str]:
    """Check the three rules. Return list of warning strings (empty = OK)."""
    warnings = []

    # Rule 1: Every PR has ≥1 slack port
    slack_by_pr: dict[int, list] = {}
    for pr, port in data['sets']['SLACK_PORT']:
        slack_by_pr.setdefault(pr, []).append(port)
    for pr in data['sets']['PR']:
        if pr not in slack_by_pr:
            warnings.append(f'  ST {tree_idx}: PR {pr} has no slack port')

    # Rule 2: Every AC line is connected to ≥1 v-f port
    v_port_ids = set(port for _, port in data['v_ports'])
    for (fp, tp) in data['ac_lines']:
        fp, tp = int(fp), int(tp)
        # Check if either endpoint (or its bus-connected PR port) is v-f
        if fp not in v_port_ids and tp not in v_port_ids:
            # Also check through buses
            fp_pr_port = _find_vf_through_bus(fp, data, v_port_ids)
            tp_pr_port = _find_vf_through_bus(tp, data, v_port_ids)
            if not fp_pr_port and not tp_pr_port:
                warnings.append(
                    f'  ST {tree_idx}: AC line ({fp},{tp}) has no v-f port')

    return warnings


def _find_vf_through_bus(port_id: int, data: dict, v_port_ids: set) -> bool:
    """Check if a port connects to a v-f port through bus topology."""
    bus_to_ports: dict = {}
    for bus, port in data['sets']['BUS_PORT']:
        bus_to_ports.setdefault(bus, set()).add(int(port))
    for bus, ports in bus_to_ports.items():
        if port_id in ports:
            for bp in ports:
                if bp in v_port_ids:
                    return True
    return False


# ─── Plotting ────────────────────────────────────────────────────────────

def _subplot_grid(n: int) -> tuple[int, int]:
    """Return (rows, cols) for a subplot grid that fits n panels."""
    cols = math.ceil(math.sqrt(n))
    rows = math.ceil(n / cols)
    return rows, cols


def plot_slack_trees(G_full: nx.Graph,
                     directed_trees: list[nx.DiGraph],
                     pq_lines_list: list[list],
                     ext_grid_pr: int,
                     save_path: str = 'results/slack_trees_interactive.html'):
    """
    Plot all spanning trees in a single interactive Plotly figure.

    Red solid directed  = slack-tree edges.
    Green dashed        = PQ (controllable) lines.
    Blue node           = ext_grid PR.
    Orange node         = regular PR.
    """
    n = len(directed_trees)
    rows, cols = _subplot_grid(n)
    fig = make_subplots(
        rows=rows, cols=cols,
        subplot_titles=[f'ST {i+1}  |  PQ lines: {len(pq)}'
                        for i, pq in enumerate(pq_lines_list)],
        horizontal_spacing=0.05,
        vertical_spacing=0.08,
    )

    pos = nx.spring_layout(G_full, seed=42)

    for idx, (tree, pq_lines) in enumerate(zip(directed_trees, pq_lines_list)):
        r = idx // cols + 1
        c = idx % cols + 1

        pq_set = set()
        for u, v in pq_lines:
            pq_set.add((min(u, v), max(u, v)))

        # Draw slack-tree edges (red, with arrows via annotations)
        for u, v in tree.edges():
            x0, y0 = pos[u]
            x1, y1 = pos[v]
            fig.add_trace(go.Scatter(
                x=[x0, x1, None], y=[y0, y1, None],
                mode='lines',
                line=dict(color='red', width=3),
                hoverinfo='text',
                text=f'Slack: PR {u} -> PR {v}',
                showlegend=False,
            ), row=r, col=c)
            # Arrow head
            fig.add_annotation(
                ax=x0, ay=y0, x=x1, y=y1,
                xref=f'x{idx+1}' if idx > 0 else 'x',
                yref=f'y{idx+1}' if idx > 0 else 'y',
                axref=f'x{idx+1}' if idx > 0 else 'x',
                ayref=f'y{idx+1}' if idx > 0 else 'y',
                showarrow=True, arrowhead=3, arrowsize=1.5,
                arrowcolor='red', arrowwidth=2,
            )

        # Draw PQ lines (green dashed)
        for u, v in pq_lines:
            x0, y0 = pos[u]
            x1, y1 = pos[v]
            fig.add_trace(go.Scatter(
                x=[x0, x1, None], y=[y0, y1, None],
                mode='lines',
                line=dict(color='green', width=2, dash='dash'),
                hoverinfo='text',
                text=f'PQ line: PR {u} - PR {v}',
                showlegend=False,
            ), row=r, col=c)

        # Draw nodes
        for node in G_full.nodes():
            x, y = pos[node]
            is_ext = (node == ext_grid_pr)
            color = 'steelblue' if is_ext else 'orange'
            label = f'PR {node} (ext_grid)' if is_ext else f'PR {node}'
            fig.add_trace(go.Scatter(
                x=[x], y=[y],
                mode='markers+text',
                marker=dict(size=30, color=color, line=dict(color='black', width=2)),
                text=str(node),
                textposition='middle center',
                textfont=dict(size=12, color='white', family='Arial Black'),
                hovertext=label,
                hoverinfo='text',
                showlegend=False,
            ), row=r, col=c)

    fig.update_layout(
        title=dict(text=f'PRG — All Spanning Trees ({n} trees)',
                   font=dict(size=16)),
        height=max(400, rows * 350),
        width=max(600, cols * 400),
        showlegend=False,
        plot_bgcolor='white',
    )

    # Hide axes
    for i in range(1, n + 1):
        ax_suffix = str(i) if i > 1 else ''
        fig.update_layout(**{
            f'xaxis{ax_suffix}': dict(showgrid=False, zeroline=False,
                                       showticklabels=False),
            f'yaxis{ax_suffix}': dict(showgrid=False, zeroline=False,
                                       showticklabels=False),
        })

    os.makedirs(os.path.dirname(save_path) or '.', exist_ok=True)
    fig.write_html(save_path)
    print(f'Slack-tree plot saved to: {save_path}')
    return fig


# ─── OPF runner for all slack trees ──────────────────────────────────────

def run_slack_tree_analysis(base_input_file: str,
                            output_file: str = 'results/sens_results_.xlsx',
                            solver: str = 'gurobi',
                            time_limit: int = 600,
                            plot: bool = True,
                            plot_path: str = 'results/slack_trees_interactive.html',
                            verbose: bool = True):
    """
    Full slack-tree pipeline:
      1. Load base topology.
      2. Build PR-level graph and enumerate all spanning trees.
      3. For each tree, assign port roles and solve OPF.
      4. Collect results in the same Excel format as sensitivity analysis.
      5. Plot all trees interactively.

    Parameters
    ----------
    base_input_file : path to the standard input.xlsx
    output_file     : path for the results Excel (same format as sens_results_.xlsx)
    solver          : solver name (default: gurobi)
    time_limit      : per-tree solver time limit (seconds)
    plot            : generate Plotly HTML plot
    plot_path       : path for the HTML plot file
    verbose         : print progress

    Returns
    -------
    output_file : str
    """
    t0 = time.time()

    # 1. Load base topology
    base_data = load_input_excel(base_input_file)

    # 2. Build PR graph and enumerate spanning trees
    G_full, ext_grid_pr, port_to_pr = _build_pr_graph(base_data)
    if ext_grid_pr is None:
        raise ValueError('No ext_grid port found in the input data.')

    print(f'PR graph: {G_full.number_of_nodes()} nodes, '
          f'{G_full.number_of_edges()} edges')

    spanning_trees_undirected = find_all_spanning_trees(G_full, root=ext_grid_pr)
    print(f'Found {len(spanning_trees_undirected)} spanning trees')

    # 3. Orient each tree and detect PQ lines
    directed_trees: list[nx.DiGraph] = []
    pq_lines_list: list[list] = []

    for tree in spanning_trees_undirected:
        dt = _orient_tree(tree, ext_grid_pr)
        pq = _find_pq_lines(G_full, dt)
        directed_trees.append(dt)
        pq_lines_list.append(pq)

    # 4. Plot
    if plot:
        plot_slack_trees(G_full, directed_trees, pq_lines_list,
                         ext_grid_pr, save_path=plot_path)

    # 5. For each tree, assign port roles, validate, solve OPF, collect results
    n_trees = len(directed_trees)

    # Prepare port/line index lists (same as sensitivity format)
    port_list = sorted(set(
        int(port) for _, port in base_data['sets']['PR_PORT']
    ))
    line_tuples = list(base_data['ac_lines'].keys()) + list(base_data['dc_lines'].keys())
    line_keys = [f'{p1},{p2}' for (p1, p2) in line_tuples]

    # Use integer timestep labels (1, 2, 3, ...) matching sensitivity format
    ts_labels = [str(i + 1) for i in range(n_trees)]

    def _empty_df(index_col, index_vals):
        df = pd.DataFrame({index_col: index_vals})
        for c in ts_labels:
            df[c] = float('nan')
        return df

    v_df     = _empty_df('Port/Timestep', port_list)
    p_df     = _empty_df('Port/Timestep', port_list)
    q_df     = _empty_df('Port/Timestep', port_list)
    ploss_df = _empty_df('Port/Timestep', port_list)
    lloss_df = _empty_df('Port/Timestep', line_keys)

    dc_voltage_ports = _get_dc_voltage_ports(base_data)
    summary_rows = []
    tree_data_list = []  # per-tree input_data dicts (for plotting)

    for t_idx, (dt, pq_lines) in enumerate(zip(directed_trees, pq_lines_list)):
        st_label = ts_labels[t_idx]

        if verbose:
            print(f'\n{"=" * 60}')
            print(f'  Slack Tree {t_idx + 1}/{n_trees}  ({st_label})')
            tree_edges = [(u, v) for u, v in dt.edges()]
            print(f'  Tree edges: {tree_edges}')
            print(f'  PQ lines:   {pq_lines}')
            print(f'{"=" * 60}')

        # Assign port roles for this tree
        tree_data = _assign_port_roles(
            base_data, G_full, dt, pq_lines, port_to_pr, ext_grid_pr)
        tree_data_list.append(tree_data)

        # Validate
        warnings = _validate_tree_config(tree_data, t_idx + 1)
        if warnings:
            for w in warnings:
                print(f'  WARNING: {w}')

        if verbose:
            _print_port_summary(tree_data, t_idx + 1)

        # Build and solve model
        model = AbstractModel()
        model = define_sets(model, tree_data)
        model, enable = define_parameters(model, tree_data)
        model = define_variables(model, enable)
        model = build_formulation(model, enable, tree_data)

        try:
            solution = run_optimization(model, solver=solver,
                                        time_limit=time_limit,
                                        verbose=verbose,
                                        raise_on_fail=False)
        except Exception as e:
            print(f'  ERROR solving tree {t_idx + 1}: {e}')
            summary_rows.append({
                'Timestep': int(st_label),
                'Tree Edges': str([(u, v) for u, v in dt.edges()]),
                'PQ Lines': str(pq_lines),
                'Termination': 'error', 'MIP Gap': None,
                'Objective (MW)': float('nan'), 'Total Loss (kW)': float('nan'),
                'AC Line Loss (kW)': float('nan'), 'DC Line Loss (kW)': float('nan'),
                'Port Loss (kW)': float('nan'), 'AC Line Loss (%)': float('nan'),
                'DC Line Loss (%)': float('nan'), 'Port Loss (%)': float('nan'),
                'Total Loss (%)': float('nan'), 'Ref Power (kW)': float('nan'),
            })
            continue

        solve_status = getattr(solution, '_solve_status', 'unknown')
        solve_gap = getattr(solution, '_solve_gap', None)
        _feasible = solve_status in ('optimal', 'locallyOptimal', 'feasible', 'other')

        if not _feasible:
            if verbose:
                print(f'  Tree {t_idx + 1}: {solve_status} -- skipping result extraction')
            summary_rows.append({
                'Timestep': int(st_label),
                'Tree Edges': str([(u, v) for u, v in dt.edges()]),
                'PQ Lines': str(pq_lines),
                'Termination': solve_status, 'MIP Gap': solve_gap,
                'Objective (MW)': float('nan'), 'Total Loss (kW)': float('nan'),
                'AC Line Loss (kW)': float('nan'), 'DC Line Loss (kW)': float('nan'),
                'Port Loss (kW)': float('nan'), 'AC Line Loss (%)': float('nan'),
                'DC Line Loss (%)': float('nan'), 'Port Loss (%)': float('nan'),
                'Total Loss (%)': float('nan'), 'Ref Power (kW)': float('nan'),
            })
            continue

        # Extract results
        has_dc = hasattr(solution, 'A_DC')
        objective_mw = _safe_value(solution.obj, default=float('nan'))

        total_port_loss_kw = 0.0
        for port in port_list:
            pi = int(port)
            try:
                v_raw = value(solution.V[pi])
                if pi not in dc_voltage_ports:
                    v_raw = math.sqrt(max(v_raw, 0))
                v_df.loc[v_df['Port/Timestep'] == port, st_label] = v_raw
            except (KeyError, ValueError):
                pass
            try:
                p_df.loc[p_df['Port/Timestep'] == port, st_label] = value(solution.P[pi])
            except (KeyError, ValueError):
                pass
            try:
                q_df.loc[q_df['Port/Timestep'] == port, st_label] = value(solution.Q[pi])
            except (KeyError, ValueError):
                pass
            try:
                ploss_val = value(solution.P_LOSS[pi])
                ploss_df.loc[ploss_df['Port/Timestep'] == port, st_label] = ploss_val * 1000
                total_port_loss_kw += abs(ploss_val) * 1000
            except (KeyError, ValueError):
                pass

        # Line losses
        full_ac_loss_kw = 0.0
        full_dc_loss_kw = 0.0
        for li, (p1, p2) in enumerate(line_tuples):
            try:
                if (p1, p2) in solution.A:
                    loss = abs(value(solution.A[p1, p2])) * value(solution.line_R[p1, p2]) * 1000
                    full_ac_loss_kw += loss
                elif has_dc and (p1, p2) in solution.A_DC:
                    loss = value(solution.dc_line_R[p1, p2]) * abs(value(solution.A_DC[p1, p2])) ** 2 * 1000
                    full_dc_loss_kw += loss
                else:
                    continue
                lloss_df.loc[lloss_df['Port/Timestep'] == line_keys[li], st_label] = loss
            except (KeyError, ValueError):
                pass

        # Also accumulate full AC/DC losses from all indices
        _full_ac = 0.0
        for (lp1, lp2) in solution.A:
            try:
                _full_ac += abs(value(solution.A[lp1, lp2])) * value(solution.line_R[lp1, lp2]) * 1000
            except (KeyError, ValueError):
                pass
        _full_dc = 0.0
        if has_dc:
            for (lp1, lp2) in solution.A_DC:
                try:
                    _full_dc += value(solution.dc_line_R[lp1, lp2]) * abs(value(solution.A_DC[lp1, lp2])) ** 2 * 1000
                except (KeyError, ValueError):
                    pass

        total_gen_kw = 0.0
        for (_, pid) in tree_data['sets'].get('EXT_GRID', []):
            try:
                total_gen_kw += abs(value(solution.P[int(pid)])) * 1000
            except (KeyError, ValueError):
                pass
        total_load_kw = 0.0
        for (_, pid) in tree_data.get('terminal_ports', []):
            try:
                total_load_kw += abs(value(solution.P[int(pid)])) * 1000
            except (KeyError, ValueError):
                pass
        for (_, pid) in tree_data.get('terminal_ports_bus', []):
            try:
                total_load_kw += abs(value(solution.P[int(pid)])) * 1000
            except (KeyError, ValueError):
                pass

        total_loss_kw = objective_mw * 1000
        ref_power_kw = total_gen_kw if total_gen_kw > 0 else total_load_kw
        pct = lambda x: (x / ref_power_kw * 100) if ref_power_kw > 1e-9 else 0.0

        summary_rows.append({
            'Timestep': int(st_label),
            'Tree Edges': str([(u, v) for u, v in dt.edges()]),
            'PQ Lines': str(pq_lines),
            'Termination': solve_status,
            'MIP Gap': solve_gap,
            'Objective (MW)': objective_mw,
            'Total Loss (kW)': total_loss_kw,
            'AC Line Loss (kW)': _full_ac,
            'DC Line Loss (kW)': _full_dc,
            'Port Loss (kW)': total_port_loss_kw,
            'AC Line Loss (%)': pct(_full_ac),
            'DC Line Loss (%)': pct(_full_dc),
            'Port Loss (%)': pct(total_port_loss_kw),
            'Total Loss (%)': pct(total_loss_kw),
            'Ref Power (kW)': ref_power_kw,
        })

        if verbose:
            r = summary_rows[-1]
            print(f'  Objective     : {objective_mw:.6f} MW')
            print(f'  AC line loss  : {_full_ac:.2f} kW  ({r["AC Line Loss (%)"]:.2f}%)')
            print(f'  DC line loss  : {_full_dc:.2f} kW  ({r["DC Line Loss (%)"]:.2f}%)')
            print(f'  Port loss     : {total_port_loss_kw:.2f} kW  ({r["Port Loss (%)"]:.2f}%)')
            print(f'  Total loss    : {total_loss_kw:.2f} kW  ({r["Total Loss (%)"]:.2f}%)')

    # 6. Build summary
    summary_df = pd.DataFrame(summary_rows)

    # 7. Write results
    os.makedirs(os.path.dirname(output_file) or '.', exist_ok=True)
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
        print(f'Slack-tree analysis complete — {n_trees} trees')
        print(f'Results saved to: {output_file}')
        print(f'Total time: {elapsed:.1f} s')
        print(f'{"=" * 60}')

    return output_file, tree_data_list


# ─── Helpers ─────────────────────────────────────────────────────────────

def _print_port_summary(data: dict, tree_idx: int):
    """Print a compact summary of port assignments for a tree."""
    print(f'  Port roles for ST {tree_idx}:')
    slack_ports = {p for _, p in data['sets']['SLACK_PORT']}
    pq_ports = {p for _, p in data['sets']['PQ_PORT']}
    v_ports = {p for _, p in data['v_ports']}
    ext_ports = {p for _, p in data['sets']['EXT_GRID']}
    term_ports = {p for _, p in data.get('terminal_ports', [])}

    for pr in sorted(data['sets']['PR']):
        pr_ports = [p for _pr, p in data['sets']['PR_PORT'] if _pr == pr]
        parts = []
        for p in sorted(pr_ports):
            if p in ext_ports:
                parts.append(f'{p}(ext_grid)')
            elif p in term_ports:
                parts.append(f'{p}(terminal)')
            elif p in slack_ports:
                parts.append(f'{p}(slack)')
            elif p in v_ports:
                parts.append(f'{p}(v-f)')
            elif p in pq_ports:
                parts.append(f'{p}(pq)')
            else:
                parts.append(f'{p}(?)')
        print(f'    PR {pr}: {", ".join(parts)}')
