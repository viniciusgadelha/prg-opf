from pyomo.environ import *
import pandas as pd
import numpy as np
import os
from prg_opt import *


def run_slack_trees(cs=3, sensitivity=False, load_variation=False, add_power=0):
    count = 0

    if cs == 3:
        cs = '/data/case_study_3/'
    else:
        cs = '/data/case_study_4/'

    for folder in os.listdir(os.getcwd() + cs):

        path = os.getcwd() + cs + folder + '/'

        model = AbstractModel()
        data = DataPortal()

        define_sets(model, data, path)

        if sensitivity:

            define_parameters_variables(model, data, path)

            pq_port = pd.read_csv(path + 'pq_ports.csv')
            opt_powers = pd.read_excel(os.getcwd() + '/results/st_summary_base.xlsx', sheet_name='active_powers',
                                       index_col=0)
            # next line we select which PQ port we will be iterating
            value = opt_powers.loc[opt_powers['index'] == str(tuple(pq_port.values[1]))]
            # value = value.iloc[0, int(folder[-1])] + add_power
            value.iloc[0, 1] = value.iloc[0, int(folder[-1])] + add_power  # folder-1 used to select the correct ST
            value = {value.iloc[0, 0]: value.iloc[0, 1]}  # cuts all other values and takes simply the port ID and power

            prg_opf_formulation(model, controlled=True, control_p=value)

        elif load_variation:

            pq_port = pd.read_csv(path + 'pq_ports.csv')
            opt_powers = pd.read_excel(os.getcwd() + '/results/st_summary_base.xlsx', sheet_name='active_powers',
                                       index_col=0)
            # next line we select which PQ port we will be FIXING
            value = opt_powers.loc[opt_powers['index'] == str(tuple(pq_port.values[0]))]  # Fixes the PQ port power
            value = {value.iloc[0, 0]: value.iloc[0, 1]}  # cuts all other values and takes simply the port ID and power

            terminal_port = pd.read_csv(path + 'terminal_ports_power.csv')
            terminal_port.loc[0, 'p_mw'] = terminal_port.loc[0, 'p_mw'] + add_power  # increment of power to the terminal port
            terminal_port.to_csv(path + 'terminal_ports_power_variation.csv', index=False)
            define_parameters_variables(model, data, path, load_variation=True)

            prg_opf_formulation(model, controlled=True, control_p=value)

        else:
            define_parameters_variables(model, data, path)
            prg_opf_formulation(model)

        solution = run_optimization(model, data)

        export_results(solution, path)

        if count == 0:
            df_v = pd.read_excel(path + 'optimization_results.xlsx', sheet_name='voltage', index_col=0)
            df_v.dropna(inplace=True)

            df_p = pd.read_excel(path + 'optimization_results.xlsx', sheet_name='active_powers', index_col=0)
            df_p.dropna(inplace=True)

            df_q = pd.read_excel(path + 'optimization_results.xlsx', sheet_name='reactive_powers', index_col=0)
            df_q.dropna(inplace=True)

            df_i = pd.read_excel(path + 'optimization_results.xlsx', sheet_name='current', index_col=0)
            df_i.dropna(inplace=True)

            df_r = pd.read_excel(path + 'optimization_results.xlsx', sheet_name='relaxation', index_col=0)
            df_r.dropna(inplace=True)
        else:
            df_v_add = pd.read_excel(path + 'optimization_results.xlsx', sheet_name='voltage', index_col=0)
            df_v_add.dropna(inplace=True)
            df_v = pd.concat([df_v, df_v_add.iloc[:, 1]], axis=1)

            df_p_add = pd.read_excel(path + 'optimization_results.xlsx', sheet_name='active_powers', index_col=0)
            df_p_add.dropna(inplace=True)
            df_p = pd.concat([df_p, df_p_add.iloc[:, 1]], axis=1)

            df_q_add = pd.read_excel(path + 'optimization_results.xlsx', sheet_name='reactive_powers', index_col=0)
            df_q_add.dropna(inplace=True)
            df_q = pd.concat([df_q, df_q_add.iloc[:, 1]], axis=1)

            df_i_add = pd.read_excel(path + 'optimization_results.xlsx', sheet_name='current', index_col=0)
            df_i_add.dropna(inplace=True)

            df_i = pd.concat([df_i, df_i_add.iloc[:, 1]], axis=1)

            df_r_add = pd.read_excel(path + 'optimization_results.xlsx', sheet_name='relaxation', index_col=0)
            df_r_add.dropna(inplace=True)
            df_r = pd.concat([df_r, df_r_add.iloc[:, 1]], axis=1)
        count += 1

    writer = pd.ExcelWriter(os.getcwd() + '/results/st_summary.xlsx', engine='xlsxwriter')

    # Write each dataframe to a different sheet in the workbook
    df_v.to_excel(writer, sheet_name='voltage')
    df_p.to_excel(writer, sheet_name='active_powers')
    df_q.to_excel(writer, sheet_name='reactive_powers')
    df_i.to_excel(writer, sheet_name='current')
    df_r.to_excel(writer, sheet_name='relaxation')

    # Save the workbook and close the writer
    writer.close()

    return


if __name__ == "__main__":

    ######### Sensitivity analysis by changing the output of pq ports  #########

    # run_slack_trees(cs=3)
    #
    # sens_df = pd.read_excel(os.getcwd() + '/results/st_summary.xlsx', sheet_name='active_powers', index_col=0)
    # sens_df.drop(sens_df.index.to_list()[1:], axis=0, inplace=True)

    # for p in np.arange(-20, 20.5, 0.5):
    #     run_slack_trees(sensitivity=True, add_power=p)
    #     sens_df_add = pd.read_excel(os.getcwd() + '/results/st_summary.xlsx', sheet_name='active_powers', index_col=0)
    #     sens_df_add.drop(sens_df_add.index.to_list()[1:], axis=0, inplace=True)
    #     sens_df_add.iloc[0, 0] = str(p)
    #     sens_df = pd.concat([sens_df, sens_df_add])
    # sens_df.to_excel('results/sensitivity_analysis_pq.xlsx')

    ############### Sensitivity analysis by changing the demand of terminal ports ############

    run_slack_trees(cs=4)

    sens_df = pd.read_excel(os.getcwd() + '/results/st_summary.xlsx', sheet_name='active_powers', index_col=0)
    sens_df.drop(sens_df.index.to_list()[0:-1], axis=0, inplace=True)
    for p in np.arange(-20, 20.5, 0.5):
        run_slack_trees(cs=4, sensitivity=False, load_variation=True, add_power=p)
        sens_df_add = pd.read_excel(os.getcwd() + '/results/st_summary.xlsx', sheet_name='active_powers', index_col=0)
        sens_df_add.drop(sens_df_add.index.to_list()[0:-1], axis=0, inplace=True)
        sens_df_add.iloc[0, 0] = str(p)
        sens_df = pd.concat([sens_df, sens_df_add])

    sens_df.to_excel('results/sensitivity_analysis_demand.xlsx')
