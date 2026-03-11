"""
prg_opf.solver — Solver configuration and execution
=====================================================
Configures and runs the Gurobi solver on the Pyomo AbstractModel instance.
"""

import os

from pyomo.environ import Constraint, SolverFactory, Var
from pyomo.opt import TerminationCondition


def run_optimization(model, solver='gurobi', time_limit=600, verbose=True):
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

    Returns
    -------
    instance : pyomo.core.ConcreteModel
        The solved concrete model instance.
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
    opt.options['LogFile'] = log_path

    print('\n####################################################')
    print('Starting optimization process...')
    print('####################################################')

    results = opt.solve(instance, tee=verbose, symbolic_solver_labels=True, keepfiles=True)

    # Check termination status
    tc = results.solver.termination_condition
    if tc not in (TerminationCondition.optimal, TerminationCondition.locallyOptimal,
                  TerminationCondition.feasible, TerminationCondition.other):
        raise RuntimeError(
            f"Solver failed — termination condition: {tc}\n"
            f"The model may be infeasible. Check input data and constraints."
        )
    if tc == TerminationCondition.other:
        print('\nWARNING: Sub-optimal solution found (numerical difficulties).')
    print('\nMODEL SUCCESSFULLY SOLVED!')

    if verbose:
        print('\n####################################################')
        print('Showing results..')
        print('####################################################')
        instance.display()

    return instance
