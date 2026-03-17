"""
prg_opf.solver — Solver configuration and execution
=====================================================
Configures and runs the Gurobi solver on the Pyomo AbstractModel instance.
"""

import os

from pyomo.environ import Constraint, SolverFactory, Var
from pyomo.opt import TerminationCondition


def run_optimization(model, solver='gurobi', time_limit=600, verbose=True,
                     raise_on_fail=True):
    """
    Create a concrete instance and solve the OPF problem.

    Parameters
    ----------
    model : pyomo.environ.AbstractModel
        Fully constructed abstract model with all sets, parameters,
        variables, constraints, and objective.
    solver : str
        Solver name (default: 'gurobi').
    time_limit : int
        Maximum solver wall-clock time in seconds.
    verbose : bool
        If True, print model summary and full solver output.
    raise_on_fail : bool
        If True (default), raise RuntimeError on solver failure.
        If False, return the instance with status metadata attached.

    Returns
    -------
    instance : pyomo.core.ConcreteModel
        The solved concrete model instance.  Always carries:
        ``_solve_status`` (str), ``_solve_gap`` (float|None),
        ``_solve_time`` (float|None).
    """
    instance = model.create_instance()

    if verbose:
        print('\n####################################################')
        print('Model summary:')
        print('####################################################')
        print('\nInstance is constructed:', instance.is_constructed())
        instance.pprint()

    num_constraints = sum(1 for _ in instance.component_data_objects(Constraint, active=True))
    num_variables = sum(1 for _ in instance.component_data_objects(Var, active=True))
    print(f"Number of constraints: {num_constraints}")
    print(f"Number of variables: {num_variables}")

    # Configure solver
    log_path = os.path.join(os.getcwd(), 'gurobi.log')
    opt = SolverFactory(solver)
    opt.options['TimeLimit'] = time_limit
    opt.options['NonConvex'] = 2
    opt.options['BarHomogeneous'] = 1
    opt.options['Aggregate'] = 0        # keep SOC auxiliary vars intact
    opt.options['NumericFocus'] = 2     # careful numerics for node relaxations
    opt.options['NodeMethod'] = 2       # barrier at B&B nodes for stability
    opt.options['LogFile'] = log_path

    print('\n####################################################')
    print('Starting optimization process...')
    print('####################################################')

    results = opt.solve(instance, tee=verbose, symbolic_solver_labels=True, keepfiles=True)

    # ── Store solver metadata on instance ─────────────────────────────
    tc = results.solver.termination_condition
    instance._solve_status = str(tc)

    try:
        ub = float(results.problem[0].upper_bound)
        lb = float(results.problem[0].lower_bound)
        instance._solve_gap = (abs(ub - lb) / max(abs(ub), 1e-10)
                                if abs(ub) > 1e-10 else 0.0)
    except (IndexError, AttributeError, TypeError, ValueError):
        instance._solve_gap = None

    try:
        instance._solve_time = float(results.solver.wallclock_time)
    except (AttributeError, TypeError, ValueError):
        instance._solve_time = None

    # Check termination status
    if tc not in (TerminationCondition.optimal, TerminationCondition.locallyOptimal,
                  TerminationCondition.feasible, TerminationCondition.other):
        if raise_on_fail:
            raise RuntimeError(
                f"Solver failed -- termination condition: {tc}\n"
                f"The model may be infeasible. Check input data and constraints."
            )
        print(f'\nWARNING: Solver did not find a feasible solution -- {tc}')
        return instance

    if tc == TerminationCondition.other:
        print('\nWARNING: Sub-optimal solution found (numerical difficulties).'
              '\n         Some results may be NaN.')
    print('\nMODEL SUCCESSFULLY SOLVED!')

    if verbose and tc != TerminationCondition.other:
        print('\n####################################################')
        print('Showing results..')
        print('####################################################')
        instance.display()

    return instance
