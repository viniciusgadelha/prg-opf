"""
main.py — Entry point for the PRG Optimal Power Flow solver.

Usage examples:
    python main.py
    python main.py -i data/examples/ieee14.xlsx
    python main.py --no-plot
"""

import argparse
import time

from pyomo.environ import AbstractModel

from prg_opf.io import load_input_excel
from prg_opf.model import define_sets, define_parameters, define_variables
from prg_opf.constraints import build_formulation
from prg_opf.solver import run_optimization
from prg_opf.results import export_results


def main():
    parser = argparse.ArgumentParser(
        description='Power Router Grid — Linearised Optimal Power Flow',
    )
    parser.add_argument(
        '-i', '--input', default='data/input.xlsx',
        help='Path to the unified Excel input file (default: data/input.xlsx)',
    )
    parser.add_argument(
        '-o', '--output', default='results/',
        help='Directory for result files (default: results/)',
    )
    parser.add_argument(
        '--solver', default='gurobi',
        help='Solver name (default: gurobi)',
    )
    parser.add_argument(
        '--time-limit', type=int, default=600,
        help='Solver time limit in seconds (default: 600)',
    )
    parser.add_argument(
        '--plot', action='store_true', default=True,
        help='Generate interactive topology plot (default: True)',
    )
    parser.add_argument(
        '--no-plot', action='store_false', dest='plot',
        help='Skip interactive topology plot',
    )
    args = parser.parse_args()

    start_time = time.time()

    # 1. Load input data
    input_data = load_input_excel(args.input)

    # 2. Build abstract model
    print('Initializing model and loading data...')
    model = AbstractModel()
    model = define_sets(model, input_data)
    model, enable_constraints = define_parameters(model, input_data)
    model = define_variables(model, enable_constraints)

    # 3. Build OPF formulation (constraints + objective)
    model = build_formulation(model, enable_constraints, input_data)

    # 4. Solve
    solution = run_optimization(
        model,
        solver=args.solver,
        time_limit=args.time_limit,
        verbose=True,
    )
    elapsed = time.time() - start_time
    print(f'\n--- Time elapsed: {elapsed:.2f} seconds ---')

    # 5. Export results
    results_file = export_results(solution, input_data=input_data, path=args.output)

    # 6. Optionally plot
    if args.plot:
        from prg_opf.plotting import plot_prg_interactive
        plot_prg_interactive(
            args.input,
            results_file=results_file,
            save_path='results/prg_topology_interactive.html',
        )


if __name__ == '__main__':
    main()
