from prg_opt import *
import os
import re

def assign_profile(profiles_df, terminal_bus_power, bus_power_default, h):

    for idx in terminal_bus_power.index:
        if idx <3 :
            terminal_bus_power.loc[idx, 'p_mw'] = bus_power_default.loc[idx, 'p_mw']*profiles_df.loc[h, 'demand']

        if idx == 5:
            terminal_bus_power.loc[idx, 'p_mw'] = bus_power_default.loc[idx, 'p_mw'] * profiles_df.loc[h, 'solar']

        if idx == 8:
            terminal_bus_power.loc[idx, 'p_mw'] = bus_power_default.loc[idx, 'p_mw'] * profiles_df.loc[h, 'wind']
        if idx == 3 or idx == 4 or idx == 6 or idx == 7:
            terminal_bus_power.loc[idx, 'p_mw'] = bus_power_default.loc[idx, 'p_mw'] * profiles_df.loc[h, 'demand']

    return terminal_bus_power


def extract_results(case):
    df = pd.read_excel(os.getcwd() + case + 'timeseries/extract_df.xlsx')

    for idx in df.index:
        for h in np.arange(0, 24):
            opt_results = pd.read_excel(os.getcwd() + case + 'timeseries/optimization_results_' + str(h) + '.xlsx', index_col='index', sheet_name='active_powers')
            opt_results.dropna(inplace=True)
            port_list = [int(s) for s in re.findall(r'\b\d+\b', df.loc[idx, 'index'])]
            value = abs(opt_results.loc[port_list, 'P [MW]'].sum()*1000)
            df.loc[idx, 'loss_' + str(h)] = value
    df.to_excel('test.xlsx')
    return


if __name__ == "__main__":

    start_time = time.time()

    case = '/data/cs2/'
    dc = True

    # extract_results(case)

    input_dict = get_dict(case)

    case = case[1:len(case)]

    print('Initializing model and loading data...')

    profiles_df = pd.read_excel(case + 'profiles.xlsx', index_col=0)
    bus_power_default = pd.read_csv(case + 'terminal_ports_bus.csv')
    for h in np.arange(0, 24):

        terminal_bus_power = pd.read_csv(case + 'terminal_ports_bus.csv')

        terminal_bus_power = assign_profile(profiles_df, terminal_bus_power, bus_power_default, h)

        terminal_bus_power.to_csv(case + 'terminal_ports_bus.csv', index_label=False, index=False)

        model = AbstractModel()
        data = DataPortal()

        model = define_sets(model, data, case)

        model, enable_constraints = define_parameters_variables(model, data, case, K=1)

        model = prg_opf_formulation(model, enable_constraints)

        solution = run_optimization(model, data)

        export_results(solution, h=str(h))
    print("\n--- Time elapsed: %s seconds ---" % (time.time() - start_time))