"""
PRG Interactive Topological Plotter
====================================
Engineering-paper style schematic of a Power Router Grid, rendered with
Plotly for interactive hover / zoom.

Layout strategy
---------------
The plotter analyses the input topology (PRs, buses, lines) and builds a
deterministic grid layout so that:

  * PRs are placed on a horizontal centre line with generous spacing.
  * AC buses form one ring (above), DC buses form another ring (below).
  * Lines are routed orthogonally (Manhattan style).
  * Ext-grid symbols, load arrows, and port circles are placed around
    the edges of each PR / bus.

Colour conventions
------------------
AC (6 kV)  : black   solid lines, black   busbars
DC (10 kV) : #40A9F1 solid lines, #40A9F1 busbars

Port types
----------
  Slack           -> red    filled circle
  Ext Grid        -> red    filled circle  (+ crosshatched square nearby)
  Voltage control -> blue   filled circle
  PQ (power)      -> green  filled circle
  Terminal        -> purple filled circle
  Internal / Bus  -> white  filled circle  with black outline

Line types
----------
  Controllable    -> green  (connected to PQ port)
  Slack           -> red    (connected to slack port)

Usage::

    from prg_plot_interactive import plot_prg_interactive

    plot_prg_interactive('data/cs2/input.xlsx')
    plot_prg_interactive('data/cs2/input.xlsx',
                         results_file='results/optimization_results_.xlsx')
"""

from __future__ import annotations

import math
import os
from collections import defaultdict

import numpy as np
import pandas as pd
import plotly.graph_objects as go

from prg_opf.io import load_input_excel

# ═══════════════════════════════════════════════════════════════════════════
#  Style constants
# ═══════════════════════════════════════════════════════════════════════════
FONT_FAMILY = "Times New Roman, serif"

C = {
    "bg":            "#F0F0F0",
    "grid":          "#DCDCDC",
    "paper":         "#FFFFFF",
    # PR boxes
    "pr_fill":       "white",
    "pr_border":     "black",
    # Busbars
    "bus_ac":        "black",
    "bus_dc":        "#40A9F1",
    # Lines
    "ac_line":       "black",
    "dc_line":       "#40A9F1",
    # Port fills
    "port_slack":    "#D42B17",
    "port_extgrid":  "#D42B17",
    "port_voltage":  "#2B6BD4",
    "port_pq":       "#1FA035",
    "port_internal": "white",
    "port_bus":      "white",
    "port_terminal":   "#8B00D4",
    "port_generation": "#FFD700",
    "port_load":       "#96789e",
    # Port outline (always black)
    "port_outline":  "black",
    # Flow arrows
    "arrow_ac":      "#D42B17",
    "arrow_dc":      "#2B6BD4",
    # Text
    "text":          "black",
    "text_light":    "#555555",
    # Ext-grid cross-hatch box
    "extgrid_fill":  "rgba(230,230,230,0.6)",
    "extgrid_line":  "black",
    # Load arrows
    "load_ac":       "black",
    "load_dc":       "#40A9F1",
    # Legend panel bg
    "legend_bg":     "rgba(255,255,255,0.95)",
}

# Sizing
PR_W = 0.8     # full width  of a PR box  (scaled down 50%)
PR_H = 0.6     # full height of a PR box  (scaled down 50%)
BUS_W = 0.55   # full width  of a bus box (slightly smaller than PR)
BUS_H = 0.40   # full height of a bus box
BUS_LEN = BUS_W / 2  # half-width used for port placement at edges
PORT_R = 19    # marker size for all ports (uniform)
EXTGRID_SZ = 0.32  # half-side of ext-grid crosshatch square


# ═══════════════════════════════════════════════════════════════════════════
#  Helpers
# ═══════════════════════════════════════════════════════════════════════════
def _classify_port(port_id: int, data: dict) -> str:
    """Return 'slack', 'ext_grid', 'pq', 'terminal', or 'internal'."""
    for _, p in data["sets"]["SLACK_PORT"]:
        if p == port_id:
            return "slack"
    for _, p in data["sets"]["PQ_PORT"]:
        if p == port_id:
            return "pq"
    for _, p in data["sets"]["EXT_GRID"]:
        if p == port_id:
            return "ext_grid"
    for _, p in data.get("terminal_ports", []):
        if p == port_id:
            return "terminal"
    return "internal"


def _is_voltage_controlled(port_id: int, data: dict) -> bool:
    """True when the port has a V setpoint (appears in v_ports)."""
    return any(p == port_id for _, p in data["v_ports"])


def _port_color(port_id: int, data: dict) -> str:
    """Determine fill colour of a port circle."""
    ptype = _classify_port(port_id, data)
    if ptype in ("slack", "ext_grid"):
        return C["port_slack"]
    if ptype == "terminal":
        return C["port_terminal"]
    if _is_voltage_controlled(port_id, data):
        return C["port_voltage"]
    if ptype == "pq":
        return C["port_pq"]
    return C["port_internal"]


def _angle_to_face(angle: float) -> str:
    """Map an angle (radians) to the nearest box face."""
    a = angle % (2 * math.pi)
    if a > math.pi:
        a -= 2 * math.pi
    if -math.pi / 4 <= a < math.pi / 4:
        return "right"
    if math.pi / 4 <= a < 3 * math.pi / 4:
        return "up"
    if -3 * math.pi / 4 <= a < -math.pi / 4:
        return "down"
    return "left"


def _spread_on_face(pos: dict, port_list: list[int], ax: float, ay: float,
                    ddx: float, ddy: float, sp: float = 0.35):
    """Spread ports evenly along a box face."""
    n = len(port_list)
    for i, port in enumerate(port_list):
        off = (i - (n - 1) / 2) * sp if n > 1 else 0
        pos[f"P{port}"] = (ax + off * ddx, ay + off * ddy)


def _load_results(results_file: str) -> dict | None:
    results: dict = {}
    # Read nodes sheet (V, P, Q, P_LOSS)
    try:
        ndf = pd.read_excel(results_file, sheet_name="nodes")
        for col, key in [("V [kV]", "V"), ("P [MW]", "P"), ("Q [MVAR]", "Q"),
                         ("P_LOSS", "P_LOSS")]:
            if col in ndf.columns:
                results[key] = {}
                for _, row in ndf.iterrows():
                    try:
                        results[key][int(row["index"])] = float(row[col])
                    except (ValueError, TypeError):
                        pass
    except Exception:
        pass
    # Read lines sheet (I)
    try:
        ldf = pd.read_excel(results_file, sheet_name="lines")
        if "I [kA]" in ldf.columns:
            results["I"] = {}
            for _, row in ldf.iterrows():
                idx = row["index"]
                if isinstance(idx, str) and "," in idx:
                    parts = idx.strip("()").split(",")
                    results["I"][(int(parts[0].strip()),
                                  int(parts[1].strip()))] = float(row["I [kA]"])
    except Exception:
        pass
    return results if results else None


def _build_line_types(data: dict) -> dict:
    """
    Classify each line as 'controllable' or 'slack' by tracing the full
    connected path through any intermediate bus ports to find the two terminal
    PR port endpoints, then checking the type of those endpoints.

    Rules:
    - Bus ports are transparent routing nodes; they are never the true path
      endpoint.  Bus-to-bus chains are traversed completely.
    - For each edge (p1, p2) the two terminal PR ports are found by BFS from
      each side, passing only through bus ports, until a PR port is reached.
    - A terminal PR port is 'controllable' when its PR has at least one PQ
      port (all ports of a PR are internally connected, so a PQ port anywhere
      on the PR makes every line on that PR controllable).
    - If either terminal is controllable → whole path is controllable.
    - Otherwise → slack.
    """
    sets = data["sets"]
    pq_port_ids  = set(p for _, p in sets["PQ_PORT"])
    all_pr_ports = set(p for _, p in sets["PR_PORT"])
    all_bus_ports = set(p for _, p in sets["BUS_PORT"])

    # PR-level: controllable if ANY port on that PR is PQ
    port_to_pr: dict = {port: pr for pr, port in sets["PR_PORT"]}
    pr_to_ports: dict = {}
    for pr, port in sets["PR_PORT"]:
        pr_to_ports.setdefault(pr, set()).add(port)
    pr_is_ctrl: dict = {
        pr: bool(ports & pq_port_ids)
        for pr, ports in pr_to_ports.items()
    }

    # Full adjacency graph (bidirectional) — lines + intra-bus connections
    all_lines = list(data["ac_lines"].keys()) + list(data["dc_lines"].keys())
    adj: dict = {}
    for (p1, p2) in all_lines:
        adj.setdefault(p1, set()).add(p2)
        adj.setdefault(p2, set()).add(p1)
    # All ports on the same bus are internally connected (busbar)
    bus_to_ports: dict = {}
    for bus, port in sets["BUS_PORT"]:
        bus_to_ports.setdefault(bus, set()).add(port)
    for bus_ports in bus_to_ports.values():
        for pa in bus_ports:
            for pb in bus_ports:
                if pa != pb:
                    adj.setdefault(pa, set()).add(pb)

    def find_pr_endpoint(start, blocked):
        """BFS from `start` (not going back through `blocked`) across bus
        ports until the first PR port is reached.  If `start` is already a
        PR port, return it immediately."""
        if start in all_pr_ports:
            return start
        visited = {blocked, start}
        queue = [start]
        while queue:
            node = queue.pop(0)
            for nbr in adj.get(node, set()):
                if nbr in visited:
                    continue
                if nbr in all_pr_ports:
                    return nbr
                if nbr in all_bus_ports:
                    visited.add(nbr)
                    queue.append(nbr)
        return None  # isolated bus segment with no PR endpoint

    def endpoint_is_ctrl(ep):
        if ep is None:
            return False
        pr = port_to_pr.get(ep)
        return pr_is_ctrl.get(pr, False) if pr is not None else False

    line_types: dict = {}
    for (p1, p2) in all_lines:
        ep1 = find_pr_endpoint(p1, p2)
        ep2 = find_pr_endpoint(p2, p1)
        controllable = endpoint_is_ctrl(ep1) or endpoint_is_ctrl(ep2)
        line_types[(p1, p2)] = "controllable" if controllable else "slack"
    return line_types


# ═══════════════════════════════════════════════════════════════════════════
#  Deterministic grid layout engine
# ═══════════════════════════════════════════════════════════════════════════

def _build_layout_linear(data: dict) -> dict:
    """
    Linear layout: PRs on a horizontal centre line, AC buses above,
    DC buses below.  Used when n_pr <= 2.
    """
    sets = data["sets"]
    ac_lines = data["ac_lines"]
    dc_lines = data["dc_lines"]

    # ── adjacency maps ────────────────────────────────────────────────
    pr_port_map: dict[int, int] = {}
    bus_port_map: dict[int, str] = {}
    for pr, port in sets["PR_PORT"]:
        pr_port_map[port] = pr
    for bus, port in sets["BUS_PORT"]:
        bus_port_map[port] = bus

    all_line_keys = list(ac_lines.keys()) + list(dc_lines.keys())

    pr_bus_edges: list[dict] = []
    bus_bus_edges: list[dict] = []
    for p1, p2 in all_line_keys:
        lt = "dc" if (p1, p2) in dc_lines else "ac"
        if p1 in pr_port_map and p2 in bus_port_map:
            pr_bus_edges.append(dict(pr=pr_port_map[p1], bus=bus_port_map[p2],
                                     pr_port=p1, bus_port=p2, line_type=lt))
        elif p2 in pr_port_map and p1 in bus_port_map:
            pr_bus_edges.append(dict(pr=pr_port_map[p2], bus=bus_port_map[p1],
                                     pr_port=p2, bus_port=p1, line_type=lt))
        elif p1 in bus_port_map and p2 in bus_port_map:
            bus_bus_edges.append(dict(bus1=bus_port_map[p1],
                                      bus2=bus_port_map[p2],
                                      port1=p1, port2=p2, line_type=lt))

    # ── classify each bus as AC or DC ─────────────────────────────────
    bus_line_type: dict[str, str] = {}
    for e in pr_bus_edges:
        bus_line_type.setdefault(e["bus"], e["line_type"])
    for e in bus_bus_edges:
        bus_line_type.setdefault(e["bus1"], e["line_type"])
        bus_line_type.setdefault(e["bus2"], e["line_type"])

    ac_buses = sorted([b for b, t in bus_line_type.items() if t == "ac"])
    dc_buses = sorted([b for b, t in bus_line_type.items() if t == "dc"])

    # ── place PRs on a horizontal centre line ─────────────────────────
    n_pr = len(sets["PR"])
    pr_spacing = 5.5
    pr_y = 0.0
    pos: dict[str, tuple[float, float]] = {}
    pr_xs: dict[int, float] = {}
    for i, pr in enumerate(sorted(sets["PR"])):
        x = (i - (n_pr - 1) / 2) * pr_spacing
        pos[f"PR{pr}"] = (x, pr_y)
        pr_xs[pr] = x

    # ── group buses by owning PR ──────────────────────────────────────
    pr_ac_buses: dict[int, list[str]] = defaultdict(list)
    pr_dc_buses: dict[int, list[str]] = defaultdict(list)
    for e in pr_bus_edges:
        target = pr_ac_buses if e["line_type"] == "ac" else pr_dc_buses
        if e["bus"] not in target[e["pr"]]:
            target[e["pr"]].append(e["bus"])

    assigned = set()
    for pr in sets["PR"]:
        assigned.update(pr_ac_buses.get(pr, []))
        assigned.update(pr_dc_buses.get(pr, []))

    bus_neighbours: dict[str, list[str]] = defaultdict(list)
    for e in bus_bus_edges:
        bus_neighbours[e["bus1"]].append(e["bus2"])
        bus_neighbours[e["bus2"]].append(e["bus1"])

    for bus in ac_buses + dc_buses:
        if bus in assigned:
            continue
        vis = {bus}
        queue = [bus]
        found_pr = None
        while queue and found_pr is None:
            cur = queue.pop(0)
            for nb in bus_neighbours.get(cur, []):
                if nb in assigned:
                    for pr in sets["PR"]:
                        if nb in pr_ac_buses.get(pr, []) + pr_dc_buses.get(pr, []):
                            found_pr = pr
                            break
                    break
                if nb not in vis:
                    vis.add(nb)
                    queue.append(nb)
        if found_pr is not None:
            lt = bus_line_type.get(bus, "ac")
            (pr_ac_buses if lt == "ac" else pr_dc_buses)[found_pr].append(bus)
            assigned.add(bus)

    # ── lay out bus rows (chain-aware ordering) ───────────────────────
    ac_row_y = pr_y + 2.6
    dc_row_y = pr_y - 2.6
    bus_spacing = 1.8

    def _build_chains(all_buses_list):
        """Find connected components (chains) among buses."""
        bus_set = set(all_buses_list)
        adj: dict[str, list[str]] = defaultdict(list)
        for e in bus_bus_edges:
            if e["bus1"] in bus_set and e["bus2"] in bus_set:
                adj[e["bus1"]].append(e["bus2"])
                adj[e["bus2"]].append(e["bus1"])
        visited = set()
        chains: list[list[str]] = []
        for b in all_buses_list:
            if b in visited:
                continue
            # walk the chain from an endpoint
            endpoints = []
            # BFS to collect component
            comp = []
            q = [b]
            visited.add(b)
            while q:
                cur = q.pop(0)
                comp.append(cur)
                for nb in adj.get(cur, []):
                    if nb not in visited:
                        visited.add(nb)
                        q.append(nb)
            # order the component as a chain
            if len(comp) <= 1:
                chains.append(comp)
                continue
            endpoints = [c for c in comp if len(adj.get(c, [])) <= 1]
            start = endpoints[0] if endpoints else comp[0]
            chain = [start]
            vis2 = {start}
            while True:
                ext = False
                for nb in adj.get(chain[-1], []):
                    if nb in bus_set and nb not in vis2:
                        chain.append(nb)
                        vis2.add(nb)
                        ext = True
                        break
                if not ext:
                    break
            for c in comp:
                if c not in vis2:
                    chain.append(c)
            chains.append(chain)
        return chains

    def _chain_anchor_x(chain, pr_buses_map):
        """Average x of the PRs that connect to any bus in this chain."""
        xs = []
        for b in chain:
            for pr, buses in pr_buses_map.items():
                if b in buses:
                    xs.append(pr_xs[pr])
        return sum(xs) / len(xs) if xs else 0.0

    def _layout_bus_row(pr_buses_map, row_y, all_buses_list):
        chains = _build_chains(all_buses_list)
        # sort chains left-to-right by their PR anchor
        chain_data = []
        for chain in chains:
            ax = _chain_anchor_x(chain, pr_buses_map)
            # orient chain: left-end should connect to leftmost PR
            if len(chain) >= 2:
                left_x = _chain_anchor_x([chain[0]], pr_buses_map)
                right_x = _chain_anchor_x([chain[-1]], pr_buses_map)
                if left_x > right_x:
                    chain = list(reversed(chain))
            chain_data.append((ax, chain))
        chain_data.sort(key=lambda x: x[0])

        ordered = []
        for _, chain in chain_data:
            ordered.extend(chain)
        n = len(ordered)
        if n == 0:
            return
        total_w = (n - 1) * bus_spacing
        sx = -total_w / 2.0
        for i, bus in enumerate(ordered):
            pos[f"BUS_{bus}"] = (sx + i * bus_spacing, row_y)

    _layout_bus_row(pr_ac_buses, ac_row_y, ac_buses)
    _layout_bus_row(pr_dc_buses, dc_row_y, dc_buses)

    # ── helper: find the target position for a port via its line ──────
    def _port_target_pos(port):
        """Return (x,y) of the component at the other end of this port's line."""
        for p1, p2 in all_line_keys:
            if p1 == port:
                other = p2
            elif p2 == port:
                other = p1
            else:
                continue
            if other in pr_port_map:
                return pos.get(f"PR{pr_port_map[other]}")
            if other in bus_port_map:
                return pos.get(f"BUS_{bus_port_map[other]}")
        return None

    # ── place PR ports (sorted by target x on each edge) ─────────────
    for pr in sets["PR"]:
        cx, cy = pos[f"PR{pr}"]
        ports = [p for pr_, p in sets["PR_PORT"] if pr_ == pr]

        port_target_bus: dict[int, str] = {}
        for e in pr_bus_edges:
            if e["pr"] == pr:
                port_target_bus[e["pr_port"]] = e["bus"]

        up_ports, down_ports, left_ports, right_ports, free_ports = (
            [], [], [], [], [])

        for port in ports:
            tgt_bus = port_target_bus.get(port)
            if tgt_bus:
                bus_key = f"BUS_{tgt_bus}"
                if bus_key in pos:
                    _, by = pos[bus_key]
                    (up_ports if by > cy else down_ports).append(port)
                    continue
            ptype = _classify_port(port, data)
            if ptype in ("slack", "ext_grid"):
                (left_ports if not left_ports else right_ports).append(port)
            else:
                free_ports.append(port)

        for port in free_ports:
            counts = [len(up_ports), len(down_ports),
                      len(left_ports), len(right_ports)]
            mi = counts.index(min(counts))
            [up_ports, down_ports, left_ports, right_ports][mi].append(port)

        # Sort horizontal edges by target x so lines don't cross
        def _target_x(port):
            tp = _port_target_pos(port)
            return tp[0] if tp else 0.0

        up_ports.sort(key=_target_x)
        down_ports.sort(key=_target_x)

        hw, hh = PR_W / 2, PR_H / 2

        def _spread(port_list, ax, ay, dx, dy, sp=0.40):
            n = len(port_list)
            for i, port in enumerate(port_list):
                off = (i - (n - 1) / 2) * sp if n > 1 else 0
                pos[f"P{port}"] = (ax + off * dx, ay + off * dy)

        _spread(up_ports,    cx, cy + hh, 1, 0)
        _spread(down_ports,  cx, cy - hh, 1, 0)
        _spread(left_ports,  cx - hw, cy, 0, 1)
        _spread(right_ports, cx + hw, cy, 0, 1)

        # Ext-grid symbols pushed further outward
        for port in ports:
            if _classify_port(port, data) == "ext_grid" and f"P{port}" in pos:
                px, py = pos[f"P{port}"]
                if abs(px - cx) > abs(py - cy):
                    dx_sign = -1 if px < cx else 1
                    pos[f"EXT_{port}"] = (px + dx_sign * 0.65, py)
                else:
                    dy_sign = -1 if py < cy else 1
                    pos[f"EXT_{port}"] = (px, py + dy_sign * 0.65)

    # ── place bus ports (face-based, on box outline) ──────────────────
    for bus in sets["BUS"]:
        bk = f"BUS_{bus}"
        if bk not in pos:
            continue
        bx, by = pos[bk]
        ports = [p for b, p in sets["BUS_PORT"] if b == bus]
        face_ports: dict[str, list[int]] = {
            "up": [], "down": [], "left": [], "right": []}
        for port in ports:
            tp = _port_target_pos(port)
            if tp is not None:
                a = math.atan2(tp[1] - by, tp[0] - bx)
                face_ports[_angle_to_face(a)].append(port)
            else:
                mf = min(face_ports, key=lambda f: len(face_ports[f]))
                face_ports[mf].append(port)
        for axis in ("up", "down"):
            face_ports[axis].sort(
                key=lambda p: (_port_target_pos(p) or (0, 0))[0])
        for axis in ("left", "right"):
            face_ports[axis].sort(
                key=lambda p: (_port_target_pos(p) or (0, 0))[1])
        hw, hh = BUS_W / 2, BUS_H / 2
        _spread_on_face(pos, face_ports["up"],    bx, by + hh, 1, 0, 0.30)
        _spread_on_face(pos, face_ports["down"],  bx, by - hh, 1, 0, 0.30)
        _spread_on_face(pos, face_ports["left"],  bx - hw, by, 0, 1, 0.30)
        _spread_on_face(pos, face_ports["right"], bx + hw, by, 0, 1, 0.30)

    return pos


# ═══════════════════════════════════════════════════════════════════════════
#  Polygon layout engine  (3-6 PRs)
# ═══════════════════════════════════════════════════════════════════════════

def _build_layout_polygon(data: dict) -> tuple[dict, dict]:
    """
    Polygon layout: PRs at regular-polygon vertices, buses along edges.
    Returns (pos, bus_angles).
    """
    sets = data["sets"]
    ac_lines = data["ac_lines"]
    dc_lines = data["dc_lines"]

    pr_port_map: dict[int, int] = {}
    bus_port_map: dict[int, str] = {}
    for pr, port in sets["PR_PORT"]:
        pr_port_map[port] = pr
    for bus, port in sets["BUS_PORT"]:
        bus_port_map[port] = bus

    all_line_keys = list(ac_lines.keys()) + list(dc_lines.keys())

    pr_bus_edges: list[dict] = []
    bus_bus_edges: list[dict] = []
    for p1, p2 in all_line_keys:
        lt = "dc" if (p1, p2) in dc_lines else "ac"
        if p1 in pr_port_map and p2 in bus_port_map:
            pr_bus_edges.append(dict(pr=pr_port_map[p1], bus=bus_port_map[p2],
                                     pr_port=p1, bus_port=p2, line_type=lt))
        elif p2 in pr_port_map and p1 in bus_port_map:
            pr_bus_edges.append(dict(pr=pr_port_map[p2], bus=bus_port_map[p1],
                                     pr_port=p2, bus_port=p1, line_type=lt))
        elif p1 in bus_port_map and p2 in bus_port_map:
            bus_bus_edges.append(dict(bus1=bus_port_map[p1],
                                      bus2=bus_port_map[p2],
                                      port1=p1, port2=p2, line_type=lt))

    bus_line_type: dict[str, str] = {}
    for e in pr_bus_edges:
        bus_line_type.setdefault(e["bus"], e["line_type"])
    for e in bus_bus_edges:
        bus_line_type.setdefault(e["bus1"], e["line_type"])
        bus_line_type.setdefault(e["bus2"], e["line_type"])

    # ── place PRs at regular polygon vertices ─────────────────────
    n_pr = len(sets["PR"])
    sorted_prs = sorted(sets["PR"])
    pos: dict[str, tuple[float, float]] = {}
    bus_angles: dict[str, float] = {}

    edge_len = 5.5
    R = edge_len / (2 * math.sin(math.pi / n_pr))
    for i, pr in enumerate(sorted_prs):
        theta = math.pi / 2 - 2 * math.pi * i / n_pr
        pos[f"PR{pr}"] = (R * math.cos(theta), R * math.sin(theta))

    # ── find bus chains (connected components) ────────────────────
    bus_adj: dict[str, set[str]] = defaultdict(set)
    for e in bus_bus_edges:
        bus_adj[e["bus1"]].add(e["bus2"])
        bus_adj[e["bus2"]].add(e["bus1"])

    all_buses = set(bus_line_type.keys())
    visited: set[str] = set()
    chains: list[list[str]] = []
    for bus in sorted(all_buses):
        if bus in visited:
            continue
        comp: list[str] = []
        q = [bus]
        visited.add(bus)
        while q:
            cur = q.pop(0)
            comp.append(cur)
            for nb in bus_adj.get(cur, set()):
                if nb not in visited:
                    visited.add(nb)
                    q.append(nb)
        chains.append(comp)

    def _order_chain(chain):
        if len(chain) <= 1:
            return chain
        chain_set = set(chain)
        endpoints = [b for b in chain
                     if len(bus_adj.get(b, set()) & chain_set) <= 1]
        start = endpoints[0] if endpoints else chain[0]
        ordered = [start]
        vis2 = {start}
        while True:
            ext = False
            for nb in bus_adj.get(ordered[-1], set()):
                if nb in chain_set and nb not in vis2:
                    ordered.append(nb)
                    vis2.add(nb)
                    ext = True
                    break
            if not ext:
                break
        return ordered

    connections: list[dict] = []
    for chain in chains:
        chain_set = set(chain)
        prs: set[int] = set()
        lt = None
        for e in pr_bus_edges:
            if e["bus"] in chain_set:
                prs.add(e["pr"])
                if lt is None:
                    lt = e["line_type"]
        if lt is None:
            lt = bus_line_type.get(chain[0], "ac")
        ordered = _order_chain(chain)
        pr_list = sorted(prs)
        if len(pr_list) >= 2:
            first_prs = {e["pr"] for e in pr_bus_edges
                         if e["bus"] == ordered[0]}
            if pr_list[0] not in first_prs and pr_list[1] in first_prs:
                ordered = list(reversed(ordered))
        connections.append({"prs": pr_list, "buses": ordered,
                            "line_type": lt})

    # ── place buses: AC on outer edges, DC on inner polygon ───────
    BUS_OFFSET = 1.3
    edge_conns: dict[tuple, list[dict]] = defaultdict(list)
    for conn in connections:
        if len(conn["prs"]) >= 2:
            key = (conn["prs"][0], conn["prs"][1])
        elif len(conn["prs"]) == 1:
            key = (conn["prs"][0], None)
        else:
            continue
        edge_conns[key].append(conn)

    inner_buses_info: list[tuple[str, float]] = []  # (bus, ideal_angle)

    for (pr_a, pr_b), conns in edge_conns.items():
        if pr_a is None:
            continue
        pa_x, pa_y = pos[f"PR{pr_a}"]
        if pr_b is not None:
            pb_x, pb_y = pos[f"PR{pr_b}"]
        else:
            dist = math.hypot(pa_x, pa_y)
            if dist < 1e-6:
                pb_x, pb_y = pa_x + edge_len, pa_y
            else:
                pb_x = pa_x + edge_len * pa_x / dist
                pb_y = pa_y + edge_len * pa_y / dist

        dx, dy = pb_x - pa_x, pb_y - pa_y
        elen = math.hypot(dx, dy)
        if elen < 1e-6:
            continue
        ux, uy = dx / elen, dy / elen
        nx, ny = -uy, ux
        mx, my = (pa_x + pb_x) / 2, (pa_y + pb_y) / 2
        if mx * nx + my * ny < 0:
            nx, ny = -nx, -ny

        bus_angle = math.atan2(uy, ux)
        for conn in conns:
            buses = conn["buses"]
            n_buses = len(buses)
            if conn["line_type"] == "ac":
                # AC buses → outer edges (existing behaviour)
                for bi, bus in enumerate(buses):
                    t = (bi + 1) / (n_buses + 1)
                    pos[f"BUS_{bus}"] = (
                        pa_x + t * dx + BUS_OFFSET * nx,
                        pa_y + t * dy + BUS_OFFSET * ny,
                    )
                    bus_angles[bus] = bus_angle
            else:
                # DC buses → collect for inner polygon
                for bi, bus in enumerate(buses):
                    t = (bi + 1) / (n_buses + 1)
                    ix = pa_x + t * dx
                    iy = pa_y + t * dy
                    inner_buses_info.append((bus, math.atan2(iy, ix)))

    # ── place DC (inner) buses on a concentric inner polygon ──────
    n_inner = len(inner_buses_info)
    if n_inner > 0:
        R_inner = R * 0.38
        shift = math.pi / 4          # 45° rotation from outer polygon
        start_angle = math.pi / 2 + shift
        vertex_angles = [
            start_angle - 2 * math.pi * k / n_inner
            for k in range(n_inner)
        ]
        # normalise to [-pi, pi]
        vertex_angles = [
            ((a + math.pi) % (2 * math.pi)) - math.pi
            for a in vertex_angles
        ]

        # sort both lists by angle and find the rotation that
        # minimises total angular mismatch
        inner_buses_info.sort(key=lambda x: x[1])
        va_order = sorted(range(n_inner), key=lambda i: vertex_angles[i])

        def _ang_dist(a1: float, a2: float) -> float:
            d = abs(a1 - a2) % (2 * math.pi)
            return min(d, 2 * math.pi - d)

        best_rot, best_err = 0, float("inf")
        for r in range(n_inner):
            err = sum(
                _ang_dist(inner_buses_info[i][1],
                          vertex_angles[va_order[(i + r) % n_inner]])
                for i in range(n_inner)
            )
            if err < best_err:
                best_err = err
                best_rot = r

        for i, (bus, _) in enumerate(inner_buses_info):
            vi = va_order[(i + best_rot) % n_inner]
            theta = vertex_angles[vi]
            pos[f"BUS_{bus}"] = (
                R_inner * math.cos(theta),
                R_inner * math.sin(theta),
            )
            bus_angles[bus] = theta

    for bus in sets["BUS"]:
        if f"BUS_{bus}" not in pos:
            pos[f"BUS_{bus}"] = (0.0, 0.0)
            bus_angles.setdefault(bus, 0.0)

    # ── helper: target position ───────────────────────────────────
    def _port_target_pos(port):
        for p1, p2 in all_line_keys:
            if p1 == port:
                other = p2
            elif p2 == port:
                other = p1
            else:
                continue
            if other in pr_port_map:
                return pos.get(f"PR{pr_port_map[other]}")
            if other in bus_port_map:
                return pos.get(f"BUS_{bus_port_map[other]}")
        return None

    # ── place PR ports (angle-based face assignment) ──────────────
    for pr in sets["PR"]:
        cx, cy = pos[f"PR{pr}"]
        ports = [p for pr_, p in sets["PR_PORT"] if pr_ == pr]
        face_ports: dict[str, list[int]] = {
            "up": [], "down": [], "left": [], "right": []}
        free_ports: list[int] = []

        for port in ports:
            tp = _port_target_pos(port)
            if tp is not None:
                a = math.atan2(tp[1] - cy, tp[0] - cx)
                face_ports[_angle_to_face(a)].append(port)
            else:
                ptype = _classify_port(port, data)
                if ptype in ("slack", "ext_grid"):
                    out_a = (math.atan2(cy, cx)
                             if math.hypot(cx, cy) > 1e-6 else math.pi)
                    face_ports[_angle_to_face(out_a)].append(port)
                else:
                    free_ports.append(port)

        for port in free_ports:
            min_face = min(face_ports, key=lambda f: len(face_ports[f]))
            face_ports[min_face].append(port)

        face_ports["up"].sort(
            key=lambda p: (_port_target_pos(p) or (0, 0))[0])
        face_ports["down"].sort(
            key=lambda p: (_port_target_pos(p) or (0, 0))[0])
        face_ports["left"].sort(
            key=lambda p: (_port_target_pos(p) or (0, 0))[1])
        face_ports["right"].sort(
            key=lambda p: (_port_target_pos(p) or (0, 0))[1])

        hw, hh = PR_W / 2, PR_H / 2

        def _spread(port_list, ax, ay, ddx, ddy, sp=0.40):
            n = len(port_list)
            for i, port in enumerate(port_list):
                off = (i - (n - 1) / 2) * sp if n > 1 else 0
                pos[f"P{port}"] = (ax + off * ddx, ay + off * ddy)

        _spread(face_ports["up"],    cx, cy + hh, 1, 0)
        _spread(face_ports["down"],  cx, cy - hh, 1, 0)
        _spread(face_ports["left"],  cx - hw, cy, 0, 1)
        _spread(face_ports["right"], cx + hw, cy, 0, 1)

        for port in ports:
            if _classify_port(port, data) == "ext_grid" and f"P{port}" in pos:
                px, py = pos[f"P{port}"]
                if abs(px - cx) > abs(py - cy):
                    dx_sign = -1 if px < cx else 1
                    pos[f"EXT_{port}"] = (px + dx_sign * 0.65, py)
                else:
                    dy_sign = -1 if py < cy else 1
                    pos[f"EXT_{port}"] = (px, py + dy_sign * 0.65)

    # ── place bus ports (face-based, on box outline) ──────────────
    for bus in sets["BUS"]:
        bk = f"BUS_{bus}"
        if bk not in pos:
            continue
        bx, by = pos[bk]
        ports = [p for b, p in sets["BUS_PORT"] if b == bus]
        face_ports: dict[str, list[int]] = {
            "up": [], "down": [], "left": [], "right": []}
        for port in ports:
            tp = _port_target_pos(port)
            if tp is not None:
                a = math.atan2(tp[1] - by, tp[0] - bx)
                face_ports[_angle_to_face(a)].append(port)
            else:
                mf = min(face_ports, key=lambda f: len(face_ports[f]))
                face_ports[mf].append(port)
        for axis in ("up", "down"):
            face_ports[axis].sort(
                key=lambda p: (_port_target_pos(p) or (0, 0))[0])
        for axis in ("left", "right"):
            face_ports[axis].sort(
                key=lambda p: (_port_target_pos(p) or (0, 0))[1])
        hw, hh = BUS_W / 2, BUS_H / 2
        _spread_on_face(pos, face_ports["up"],    bx, by + hh, 1, 0, 0.30)
        _spread_on_face(pos, face_ports["down"],  bx, by - hh, 1, 0, 0.30)
        _spread_on_face(pos, face_ports["left"],  bx - hw, by, 0, 1, 0.30)
        _spread_on_face(pos, face_ports["right"], bx + hw, by, 0, 1, 0.30)

    return pos, bus_angles


def _build_layout(data: dict) -> tuple[dict, dict]:
    """
    Dispatch to linear (n_pr <= 2) or polygon (3-6 PRs) layout.
    Returns (pos, bus_angles).
    """
    n_pr = len(data["sets"]["PR"])
    if n_pr > 6:
        raise ValueError(f"Maximum 6 PRs supported, got {n_pr}")
    if n_pr <= 2:
        return _build_layout_linear(data), {}
    return _build_layout_polygon(data)


# ═══════════════════════════════════════════════════════════════════════════
#  Orthogonal (Manhattan) line routing with channel separation
# ═══════════════════════════════════════════════════════════════════════════

def _assign_routing_channels(pos: dict, data: dict):
    """
    Pre-compute a y-channel for every line that needs an L/Z-bend,
    so that parallel horizontal segments don't overlap.

    Returns a dict  (p1, p2) -> channel_y.
    """
    ac_lines = data["ac_lines"]
    dc_lines = data["dc_lines"]
    all_keys = list(ac_lines.keys()) + list(dc_lines.keys())

    # separate lines that are purely horizontal / vertical from those
    # needing a bend (different x AND different y)
    bend_lines_up: list[tuple] = []   # lines going into AC bus row (y>0)
    bend_lines_dn: list[tuple] = []   # lines going into DC bus row (y<0)

    for p1, p2 in all_keys:
        k1, k2 = f"P{p1}", f"P{p2}"
        if k1 not in pos or k2 not in pos:
            continue
        x1, y1 = pos[k1]
        x2, y2 = pos[k2]
        if abs(x1 - x2) < 1e-9 or abs(y1 - y2) < 1e-9:
            continue  # straight line, no channel needed
        mid_y = (y1 + y2) / 2
        if mid_y > 0:
            bend_lines_up.append((p1, p2, x1, y1, x2, y2))
        else:
            bend_lines_dn.append((p1, p2, x1, y1, x2, y2))

    channels: dict[tuple, float] = {}

    def _assign(lines, y_lo, y_hi):
        """Distribute channel y-values between y_lo and y_hi."""
        if not lines:
            return
        # sort by horizontal span length (longest in the middle)
        lines_sorted = sorted(lines,
                              key=lambda t: abs(t[4] - t[2]),
                              reverse=True)
        n = len(lines_sorted)
        for i, (p1, p2, *_) in enumerate(lines_sorted):
            t = (i + 1) / (n + 1)
            channels[(p1, p2)] = y_lo + (y_hi - y_lo) * t

    # AC bend lines: channels between PR top (PR_H/2) and AC bus row
    pr_top = PR_H / 2 + 0.15
    ac_bot = 2.6 - 0.15
    _assign(bend_lines_up, pr_top, ac_bot)

    pr_bot = -(PR_H / 2 + 0.15)
    dc_top = -(2.6 - 0.15)
    _assign(bend_lines_dn, dc_top, pr_bot)

    return channels


def _ortho_route(x1, y1, x2, y2, channel_y=None):
    """
    Manhattan path from (x1, y1) to (x2, y2).

    If *channel_y* is given, the horizontal segment runs at that y;
    otherwise falls back to midpoint.
    """
    if abs(x1 - x2) < 1e-9 or abs(y1 - y2) < 1e-9:
        return [x1, x2], [y1, y2]
    cy = channel_y if channel_y is not None else (y1 + y2) / 2
    return [x1, x1, x2, x2], [y1, cy, cy, y2]


# ═══════════════════════════════════════════════════════════════════════════
#  Drawing functions
# ═══════════════════════════════════════════════════════════════════════════

def _draw_pr_boxes(fig: go.Figure, pos: dict, data: dict):
    """White rectangles with black border, labelled PR 1, PR 2 ..."""
    for pr in data["sets"]["PR"]:
        cx, cy = pos[f"PR{pr}"]
        hw, hh = PR_W / 2, PR_H / 2
        fig.add_shape(
            type="rect", x0=cx - hw, y0=cy - hh, x1=cx + hw, y1=cy + hh,
            fillcolor=C["pr_fill"],
            line=dict(color=C["pr_border"], width=2.5),
            layer="below",
        )
        fig.add_annotation(
            x=cx, y=cy,
            text=f"<b>PR {pr}</b>",
            showarrow=False,
            font=dict(size=16, color=C["text"], family=FONT_FAMILY),
        )


def _draw_extgrid_symbols(fig: go.Figure, pos: dict, data: dict):
    """Cross-hatched squares next to ext-grid ports."""
    for _, port in data["sets"]["EXT_GRID"]:
        ek = f"EXT_{port}"
        if ek not in pos:
            continue
        ex, ey = pos[ek]
        s = EXTGRID_SZ

        fig.add_shape(
            type="rect", x0=ex - s, y0=ey - s, x1=ex + s, y1=ey + s,
            fillcolor=C["extgrid_fill"],
            line=dict(color=C["extgrid_line"], width=1.5),
            layer="above",
        )
        n_hatch = 4
        for i in range(n_hatch):
            t = (i + 0.5) / n_hatch
            fig.add_shape(
                type="line",
                x0=ex - s + 2 * s * t, y0=ey - s,
                x1=ex - s,              y1=ey - s + 2 * s * t,
                line=dict(color=C["extgrid_line"], width=0.7),
                layer="above",
            )
            fig.add_shape(
                type="line",
                x0=ex + s, y0=ey + s - 2 * s * t,
                x1=ex + s - 2 * s * t, y1=ey + s,
                line=dict(color=C["extgrid_line"], width=0.7),
                layer="above",
            )

        fig.add_annotation(
            x=ex, y=ey - s - 0.18,
            text="<b>Ext Grid</b>",
            showarrow=False,
            font=dict(size=8, color=C["text"], family=FONT_FAMILY),
        )
        pk = f"P{port}"
        if pk in pos:
            px, py = pos[pk]
            fig.add_shape(
                type="line", x0=px, y0=py, x1=ex, y1=ey,
                line=dict(color=C["extgrid_line"], width=1.5, dash="dot"),
            )


def _draw_busbars(fig: go.Figure, pos: dict, data: dict,
                  bus_angles: dict | None = None):
    """Rectangles for each bus, colour coded AC / DC. Angle-aware."""
    sets = data["sets"]
    ac_lines = data["ac_lines"]
    dc_lines = data["dc_lines"]

    bus_port_map: dict[int, str] = {}
    for bus, port in sets["BUS_PORT"]:
        bus_port_map[port] = bus
    bus_lt: dict[str, str] = {}
    for p1, p2 in list(ac_lines.keys()) + list(dc_lines.keys()):
        lt = "dc" if (p1, p2) in dc_lines else "ac"
        for p in (p1, p2):
            if p in bus_port_map:
                bus_lt.setdefault(bus_port_map[p], lt)

    ba = bus_angles or {}
    legend_added = {"ac": False, "dc": False}

    for bus in sets["BUS"]:
        bk = f"BUS_{bus}"
        if bk not in pos:
            continue
        bx, by = pos[bk]
        lt = bus_lt.get(bus, "ac")
        colour = C["bus_ac"] if lt == "ac" else C["bus_dc"]

        hw, hh = BUS_W / 2, BUS_H / 2
        fig.add_shape(
            type="rect",
            x0=bx - hw, y0=by - hh, x1=bx + hw, y1=by + hh,
            fillcolor="white",
            line=dict(color=colour, width=2),
            layer="below",
        )
        fig.add_annotation(
            x=bx, y=by,
            text=f"<b>{bus}</b>",
            showarrow=False,
            font=dict(size=9, color=colour, family=FONT_FAMILY),
        )
        if not legend_added[lt]:
            lname = "AC Bus (6 kV)" if lt == "ac" else "DC Bus (10 kV)"
            fig.add_trace(go.Scatter(
                x=[None], y=[None], mode="markers",
                marker=dict(symbol="square", size=14, color="white",
                            line=dict(color=colour, width=2)),
                name=lname, showlegend=True,
            ))
            legend_added[lt] = True


def _draw_lines(fig: go.Figure, pos: dict, data: dict,
                results: dict | None,
                channels: dict | None = None,
                straight: bool = False):
    """AC and DC transmission lines (straight or Manhattan routing)."""

    # Classify lines using bus-aware topology analysis
    line_types = _build_line_types(data)
    C_line = {"controllable": C["port_pq"], "slack": C["port_slack"]}

    def _line_type(p1, p2):
        return line_types.get((p1, p2), "slack")

    # Legend entries: AC (solid) and DC (dashed) for each line type
    legend_added = {"controllable_ac": False, "controllable_dc": False,
                    "slack_ac": False, "slack_dc": False}

    ch = channels or {}

    def _one(p1, p2, vals, lt):
        k1, k2 = f"P{p1}", f"P{p2}"
        if k1 not in pos or k2 not in pos:
            return
        x1, y1 = pos[k1]
        x2, y2 = pos[k2]
        ltype = _line_type(p1, p2)
        colour = C_line[ltype]
        dash = "dash" if lt == "dc" else "solid"
        if straight:
            xs, ys = [x1, x2], [y1, y2]
        else:
            cy = ch.get((p1, p2), ch.get((p2, p1)))
            xs, ys = _ortho_route(x1, y1, x2, y2, channel_y=cy)
        hover = (f"<b>{'AC' if lt == 'ac' else 'DC'} Line: {p1} - {p2}</b>"
                 f"<br>Type: {ltype.capitalize()}"
                 f"<br>R={vals['R']:.6f}, X={vals['X']:.6f}"
                 f"<br>Smax={vals['Smax']:.2f}")
        if results:
            iv = results.get("I", {}).get((p1, p2))
            if iv is not None:
                hover += f"<br>I = {iv:.4f} kA"
            pv = results.get("P", {}).get(p1)
            if pv is not None:
                hover += f"<br>P_send = {pv:.4f} MW"
        legend_key = f"{ltype}_{lt}"
        show = not legend_added[legend_key]
        lname = f"{'Controllable' if ltype == 'controllable' else 'Slack'} Line ({'DC' if lt == 'dc' else 'AC'})"
        fig.add_trace(go.Scatter(
            x=xs + [None], y=ys + [None], mode="lines",
            line=dict(color=colour, width=2, dash=dash),
            hovertext=hover, hoverinfo="text",
            name=lname, showlegend=show, legendgroup=lname,
        ))
        if show:
            legend_added[legend_key] = True

    for (p1, p2), v in data["ac_lines"].items():
        _one(p1, p2, v, "ac")
    for (p1, p2), v in data["dc_lines"].items():
        _one(p1, p2, v, "dc")


def _draw_ports(fig: go.Figure, pos: dict, data: dict,
                results: dict | None):
    """Circular port markers, colour coded by type."""
    sets = data["sets"]
    all_pr_ports = set(p for _, p in sets["PR_PORT"])
    all_bus_ports = set(p for _, p in sets["BUS_PORT"])

    groups = {
        "Slack Port":       {"color": C["port_slack"],      "ports": [], "sz": PORT_R, "txtcol": "white"},
        "Ext Grid Port":    {"color": C["port_extgrid"],    "ports": [], "sz": PORT_R, "txtcol": "white"},
        "Voltage Ctrl":     {"color": C["port_voltage"],    "ports": [], "sz": PORT_R, "txtcol": "white"},
        "PQ Port":          {"color": C["port_pq"],         "ports": [], "sz": PORT_R, "txtcol": "white"},
        "Generation Port":  {"color": C["port_generation"], "ports": [], "sz": PORT_R, "txtcol": "black"},
        "Load Port":        {"color": C["port_load"],       "ports": [], "sz": PORT_R, "txtcol": "white"},
        "Terminal Port":    {"color": C["port_terminal"],   "ports": [], "sz": PORT_R, "txtcol": "white"},
        "Internal Port":    {"color": C["port_internal"],   "ports": [], "sz": PORT_R, "txtcol": "white"},
        "Bus Port":         {"color": C["port_bus"],        "ports": [], "sz": PORT_R, "txtcol": C["text"]},
    }

    # Build port_id → P_setpoint map from input data (PR terminal ports + bus terminal ports)
    p_set_map = {}
    for (_, port_id), pval in data.get("terminal_port_p", {}).items():
        p_set_map[port_id] = pval
    for (_, port_id), pval in data.get("terminal_port_bus_p", {}).items():
        p_set_map[port_id] = pval

    for port in sets["PORT"]:
        pk = f"P{port}"
        if pk not in pos:
            continue
        p_set = p_set_map.get(port)
        if p_set is not None and p_set != 0:
            if p_set < 0:
                groups["Generation Port"]["ports"].append(port)
            else:
                groups["Load Port"]["ports"].append(port)
        elif port in all_pr_ports:
            ptype = _classify_port(port, data)
            if ptype == "slack":
                if any(p == port for _, p in sets["EXT_GRID"]):
                    groups["Ext Grid Port"]["ports"].append(port)
                else:
                    groups["Slack Port"]["ports"].append(port)
            elif ptype == "ext_grid":
                groups["Ext Grid Port"]["ports"].append(port)
            elif ptype == "pq":
                if _is_voltage_controlled(port, data):
                    groups["Voltage Ctrl"]["ports"].append(port)
                else:
                    groups["PQ Port"]["ports"].append(port)
            elif ptype == "terminal":
                groups["Terminal Port"]["ports"].append(port)
            else:  # internal
                if _is_voltage_controlled(port, data):
                    groups["Voltage Ctrl"]["ports"].append(port)
                else:
                    groups["Internal Port"]["ports"].append(port)
        elif port in all_bus_ports:
            groups["Bus Port"]["ports"].append(port)

    for gname, g in groups.items():
        if not g["ports"]:
            continue
        xs, ys, texts, hovers = [], [], [], []
        for port in g["ports"]:
            px, py = pos[f"P{port}"]
            xs.append(px)
            ys.append(py)
            texts.append(f"<b>{port}</b>")
            parent = ""
            if port in all_pr_ports:
                parent = f"PR {next(pr for pr, p in sets['PR_PORT'] if p == port)}"
            elif port in all_bus_ports:
                parent = f"Bus {next(b for b, p in sets['BUS_PORT'] if p == port)}"
            h = f"<b>Port {port}</b><br>Type: {gname}<br>Parent: {parent}"
            p_set = p_set_map.get(port)
            if p_set is not None and p_set != 0:
                h += f"<br>P_setpoint = {p_set:.4f} MW"
            if results:
                pv = results.get("P", {}).get(port)
                qv = results.get("Q", {}).get(port)
                vv = results.get("V", {}).get(port)
                lv = results.get("P_LOSS", {}).get(port)
                if pv is not None:
                    h += f"<br>P = {pv:.4f} MW"
                if qv is not None:
                    h += f"<br>Q = {qv:.4f} MVAR"
                if vv is not None:
                    h += f"<br>V = {vv:.4f} kV"
                if lv is not None:
                    h += f"<br>P_LOSS = {lv * 1000:.4f} kW"
            hovers.append(h)

        fig.add_trace(go.Scatter(
            x=xs, y=ys, mode="markers+text",
            marker=dict(symbol="circle", size=g["sz"],
                        color=g["color"],
                        line=dict(color=C["port_outline"], width=1.5)),
            text=texts,
            textposition="middle center",
            textfont=dict(size=10, color=g["txtcol"], family=FONT_FAMILY),
            hovertext=hovers, hoverinfo="text",
            name=gname, showlegend=True,
        ))


def _draw_flow_arrows(fig: go.Figure, pos: dict, data: dict,
                      results: dict,
                      channels: dict | None = None,
                      straight: bool = False):
    """Small power flow arrows at the midpoint of each line, coloured by line type."""
    p_vals = results.get("P", {})
    line_types = _build_line_types(data)
    all_lines = list(data["ac_lines"].keys()) + list(data["dc_lines"].keys())
    C_line = {"controllable": C["port_pq"], "slack": C["port_slack"]}

    ch = channels or {}

    for (p1, p2) in all_lines:
        k1, k2 = f"P{p1}", f"P{p2}"
        if k1 not in pos or k2 not in pos:
            continue
        pv = p_vals.get(p1, 0)
        if abs(pv) < 1e-6:
            continue

        x1, y1 = pos[k1]
        x2, y2 = pos[k2]
        ltype = line_types.get((p1, p2), "slack")
        colour = C_line[ltype]

        if straight:
            xs, ys = [x1, x2], [y1, y2]
        else:
            cy = ch.get((p1, p2), ch.get((p2, p1)))
            xs, ys = _ortho_route(x1, y1, x2, y2, channel_y=cy)

        # Find longest segment and place arrow at its midpoint
        best_i, best_len = 0, 0
        for i in range(len(xs) - 1):
            sl = math.hypot(xs[i + 1] - xs[i], ys[i + 1] - ys[i])
            if sl > best_len:
                best_len = sl
                best_i = i
        if best_len < 0.2:
            continue

        sx, sy = xs[best_i], ys[best_i]
        ex, ey = xs[best_i + 1], ys[best_i + 1]
        mx, my = (sx + ex) / 2, (sy + ey) / 2
        dx, dy = ex - sx, ey - sy
        seg_l = math.hypot(dx, dy)
        ux, uy = dx / seg_l, dy / seg_l
        ah = min(0.12, seg_l * 0.2)

        if pv > 0:
            ax_pt, ay_pt = mx - ux * ah, my - uy * ah
            end_x, end_y = mx + ux * ah, my + uy * ah
        else:
            ax_pt, ay_pt = mx + ux * ah, my + uy * ah
            end_x, end_y = mx - ux * ah, my - uy * ah

        fig.add_annotation(
            x=end_x, y=end_y, ax=ax_pt, ay=ay_pt,
            xref="x", yref="y", axref="x", ayref="y",
            showarrow=True,
            arrowhead=3, arrowsize=1.0, arrowwidth=1.8,
            arrowcolor=colour,
        )


# ═══════════════════════════════════════════════════════════════════════════
#  Main entry point
# ═══════════════════════════════════════════════════════════════════════════

def plot_prg_interactive(
    input_file: str,
    results_file: str | None = None,
    save_path: str | None = None,
    show: bool = True,
    title: str | None = None,
) -> go.Figure:
    """
    Draw an interactive, engineering-paper-style PRG topology with Plotly.

    Parameters
    ----------
    input_file   : Path to unified ``input.xlsx``.
    results_file : Optional path to ``optimization_results_.xlsx``.
    save_path    : If given, save the figure (HTML or image).
    show         : Open in browser.
    title        : Custom title string.

    Returns
    -------
    fig : plotly.graph_objects.Figure
    """
    data = load_input_excel(input_file)
    pos, bus_angles = _build_layout(data)
    results = (_load_results(results_file)
               if results_file and os.path.exists(results_file) else None)

    fig = go.Figure()

    # draw order: back-to-front
    _draw_pr_boxes(fig, pos, data)
    _draw_busbars(fig, pos, data, bus_angles)
    if bus_angles:
        _draw_lines(fig, pos, data, results, straight=True)
    else:
        channels = _assign_routing_channels(pos, data)
        _draw_lines(fig, pos, data, results, channels=channels)
    _draw_extgrid_symbols(fig, pos, data)
    _draw_ports(fig, pos, data, results)
    if results:
        if bus_angles:
            _draw_flow_arrows(fig, pos, data, results, straight=True)
        else:
            _draw_flow_arrows(fig, pos, data, results, channels=channels)

    # ── figure layout ─────────────────────────────────────────────────
    fig.update_layout(
        title=dict(
            text=title or "Power Router Grid  -  Topology & Results",
            font=dict(size=18, color=C["text"], family=FONT_FAMILY),
            x=0.42,
        ),
        plot_bgcolor=C["bg"],
        paper_bgcolor=C["paper"],
        showlegend=True,
        legend=dict(
            x=1.005, y=1, xanchor="left", yanchor="top",
            bgcolor=C["legend_bg"],
            bordercolor="rgb(180,180,180)", borderwidth=1,
            font=dict(size=11, family=FONT_FAMILY, color=C["text"]),
            itemsizing="constant",
            title=dict(text="<b>Legend</b>",
                       font=dict(size=13, family=FONT_FAMILY)),
            tracegroupgap=4,
        ),
        xaxis=dict(
            showgrid=False, zeroline=False, showticklabels=False,
            scaleanchor="y", scaleratio=1,
        ),
        yaxis=dict(
            showgrid=False, zeroline=False, showticklabels=False,
        ),
        hovermode="closest",
        margin=dict(l=15, r=15, t=55, b=15),
        width=1600, height=760,
    )

    # ── save / show ───────────────────────────────────────────────────
    if save_path:
        if save_path.endswith(".html"):
            html_str = fig.to_html()
            with open(save_path, "w", encoding="utf-8") as fh:
                fh.write(html_str)
        else:
            fig.write_image(save_path, scale=2)
        print(f"Plot saved to {save_path}")

    if show:
        fig.show()

    return fig


# ═══════════════════════════════════════════════════════════════════════════
#  CLI
# ═══════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    import sys
    input_file = sys.argv[1] if len(sys.argv) > 1 else "data/cs2/input.xlsx"
    results_file = sys.argv[2] if len(sys.argv) > 2 else None
    plot_prg_interactive(input_file, results_file=results_file,
                         save_path="results/prg_topology_interactive.html")
