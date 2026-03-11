# Archive

This directory contains **outdated scripts** preserved for reference only.

These scripts use an older API (e.g. `get_dict()`, `DataPortal`-based data loading,
old function signatures) and **cannot run** with the current codebase without
significant modifications.

They are kept here so historical approaches and analysis logic can be reviewed
if needed. They are not maintained and should not be imported.

## Contents

| File | Original purpose |
|------|-----------------|
| `sensitivity.py` | Multi-tree sensitivity analysis (PQ port / load variation) |
| `sensitivity_loss_factor.py` | Sensitivity sweep over loss factor K |
| `sensitivity_terminal_ports.py` | Sensitivity sweep over terminal port power |
| `timeseries_IEEE-9bus.py` | 24-hour timeseries OPF with demand/solar/wind profiles |
| `pandapower_test.py` | PandaPower 3-bus network comparison |
| `generate_mmc_model.py` | ML surrogate model training for MMC losses |
