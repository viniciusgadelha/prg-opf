# PRG-OPF — Power Router Grid Optimal Power Flow

A Pyomo-based linearised OPF solver for **Power Router Grids (PRGs)**.
The model determines optimal active/reactive power dispatch and converter
loss allocation across a meshed AC/DC grid containing Power Routers.

## Features

- **MILP / MIQCP** formulation with Big-M loss decomposition and SOCP relaxation
- Gurobi solver integration (configurable time-limit and non-convexity handling)
- **Per-port loss coefficients** for detailed converter modeling (c0, c1 per Power Router port)
- Unified Excel-based input format (Power Routers, Buses, AC Lines, DC Lines, Parameters)
- Automatic results export to Excel (nodes & lines sheets)
- Interactive topology visualisation with Plotly

## Installation

```bash
# Create / activate your environment (conda example)
conda activate fever

# Install dependencies
pip install -r requirements.txt
```

> **Note:** A licensed [Gurobi](https://www.gurobi.com/) installation must be available
> on `PATH`.  Academic licenses are free.

## Quick Start

```bash
# Run with the default case study
python main.py

# Specify input and output directory
# You can use the three examples present in data/examples/ to try out
python main.py -i data/examples/3PR-setup.xlsx -o results/

# Skip the interactive plot
python main.py --no-plot
```

### CLI Options

| Flag | Default | Description |
|------|---------|-------------|
| `-i`, `--input` | `data/input.xlsx` | Path to the unified Excel input file |
| `-o`, `--output` | `results/` | Directory for result files |
| `--solver` | `gurobi` | Solver name |
| `--time-limit` | `600` | Solver wall-clock limit (seconds) |
| `--plot` / `--no-plot` | `--plot` | Generate interactive topology HTML |

## Project Structure

```
main.py                  ← entry point (CLI)
prg_opf/
    __init__.py
    io.py                ← Excel input parser
    model.py             ← Pyomo sets, parameters, variables
    constraints.py       ← OPF constraints & objective
    solver.py            ← solver configuration & execution
    results.py           ← results extraction & Excel export
    plotting.py          ← interactive Plotly topology
    mmc/
        __init__.py
        parameters.py    ← MMC converter parameter calculations
        losses.py        ← MMC converter loss model
data/                    ← input workbooks
results/                 ← solver output (git-ignored)
archive/                 ← legacy/deprecated scripts
```

## Input Format

The solver expects a single Excel workbook with these sheets:

| Sheet | Key Columns |
|-------|-------------|
| **Power Routers** | `PR`, `Port`, `Type` (slack, ext_grid, pq, v-f, terminal), `V_setpoint`, `P_setpoint`, `Q_setpoint`, `c0`, `c1` |
| **Buses** | `Bus`, `Port`, `P_setpoint`, `Q_setpoint` |
| **AC Lines** | `From_Port`, `To_Port`, `R`, `X`, `Smax` |
| **DC Lines** | `From_Port`, `To_Port`, `R`, `X`, `Smax` |
| **Parameters** | `Sbase`, `Vbase_squared`, `BigM`, `loss_c0`, `loss_c1` |

### Loss Coefficients

Loss coefficients (`c0`, `c1`) can be specified **per-port** in the Power Routers sheet.
If not provided, the solver falls back to the global defaults from the Parameters sheet:
- `c0`: constant loss term (default: `-5.82e-5`)
- `c1`: linear loss term (default: `1.54e-3`)

This allows modeling ports with different converter characteristics (e.g., MMC vs. traditional converters).

See `data/input.xlsx` or the examples in `data/examples/` for working cases.

## License

Internal / research use.