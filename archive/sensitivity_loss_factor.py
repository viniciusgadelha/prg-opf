import pandas as pd

from prg_opt import *


if __name__ == "__main__":

    case = '/data/test/'
    dc = False
    count = 0
    input_dict = get_dict(case)

    case = case[1:len(case)]

    print('Initializing model and loading data...')

    model = AbstractModel()
    data = DataPortal()

    model = define_sets(model, data, case)

    for k in np.arange(1, 100.5, 0.5):

        model = define_parameters_variables(model, data, case, DC=dc, K=k)

        model = prg_opf_formulation(model, DC=dc)

        solution = run_optimization(model, data)

        export_results(solution)
        if count == 0:
            df_v = pd.read_excel(os.getcwd() + '/results/optimization_results.xlsx', sheet_name='voltage', index_col=0)
            df_v.dropna(inplace=True)
            df_v.rename(columns={'V [kV^2]': 'V [kV^2], K=' + str(k)}, inplace=True)

            df_p = pd.read_excel(os.getcwd() + '/results/optimization_results.xlsx', sheet_name='active_powers', index_col=0)
            df_p.dropna(inplace=True)
            df_p.rename(columns={'P [MW]': 'P [MW], K=' + str(k)}, inplace=True)

            df_q = pd.read_excel(os.getcwd() + '/results/optimization_results.xlsx', sheet_name='reactive_powers', index_col=0)
            df_q.dropna(inplace=True)
            df_p.rename(columns={'Q [MVAR]': 'Q [MVAR], K=' + str(k)}, inplace=True)

            df_i = pd.read_excel(os.getcwd() + '/results/optimization_results.xlsx', sheet_name='current', index_col=0)
            df_i.dropna(inplace=True)
            df_p.rename(columns={'I [kA^2]': 'I [kA^2], K=' + str(k)}, inplace=True)

            df_r = pd.read_excel(os.getcwd() + '/results/optimization_results.xlsx', sheet_name='relaxation', index_col=0)
            df_r.dropna(inplace=True)

            df_losses = pd.read_excel(os.getcwd() + '/results/optimization_results.xlsx', sheet_name='losses', index_col=0)
            df_losses.dropna(inplace=True)
            df_p.rename(columns={'solution': 'Loss [MW], K=' + str(k)}, inplace=True)

            # here we can add the ports we want to evaluate the power, or the sum (losses).
            watch_list = list(df_i['index'])
            watch_list.extend(['(3, 4)', '(1)', '(2)'])
            sens_df = pd.DataFrame(index=[k], columns=[watch_list])
            for column in sens_df:
                if column[0].__len__() > 3:
                    sens_df[column] = abs(df_p.iloc[int(column[0][1]), 1] + df_p.iloc[int(column[0][4]), 1]) * 1000
                else:
                    sens_df[column] = df_p.iloc[int(column[0][1]), 1] * 1000


        else:
            df_v_add = pd.read_excel(os.getcwd() + '/results/optimization_results.xlsx', sheet_name='voltage', index_col=0)
            df_v_add.dropna(inplace=True)
            df_v_add.rename(columns={'V [kV^2]': 'V [kV^2], K=' + str(k)}, inplace=True)
            df_v = pd.concat([df_v, df_v_add.iloc[:, 1]], axis=1)

            df_p_add = pd.read_excel(os.getcwd() + '/results/optimization_results.xlsx', sheet_name='active_powers', index_col=0)
            df_p_add.dropna(inplace=True)
            df_p_add.rename(columns={'P [MW]': 'P [MW], K=' + str(k)}, inplace=True)
            df_p = pd.concat([df_p, df_p_add.iloc[:, 1]], axis=1)

            df_q_add = pd.read_excel(os.getcwd() + '/results/optimization_results.xlsx', sheet_name='reactive_powers', index_col=0)
            df_q_add.dropna(inplace=True)
            df_q_add.rename(columns={'Q [MVAR]': 'Q [MVAR], K=' + str(k)}, inplace=True)
            df_q = pd.concat([df_q, df_q_add.iloc[:, 1]], axis=1)

            df_i_add = pd.read_excel(os.getcwd() + '/results/optimization_results.xlsx', sheet_name='current', index_col=0)
            df_i_add.dropna(inplace=True)
            df_i_add.rename(columns={'I [kA^2]': 'I [kA^2], K=' + str(k)}, inplace=True)

            df_i = pd.concat([df_i, df_i_add.iloc[:, 1]], axis=1)

            df_r_add = pd.read_excel(os.getcwd() + '/results/optimization_results.xlsx', sheet_name='relaxation', index_col=0)
            df_r_add.dropna(inplace=True)
            df_r = pd.concat([df_r, df_r_add.iloc[:, 1]], axis=1)

            df_losses_add = pd.read_excel(os.getcwd() + '/results/optimization_results.xlsx', sheet_name='losses', index_col=0)
            df_losses_add.dropna(inplace=True)
            df_losses_add.rename(columns={'solution': 'Loss [MW], K=' + str(k)}, inplace=True)
            df_losses = pd.concat([df_losses, df_losses_add.iloc[:, 1]], axis=1)

            sens_df_add = pd.DataFrame(index=[k], columns=[watch_list])
            for column in sens_df_add:
                if column[0].__len__() > 3:
                    sens_df_add[column] = abs(df_p_add.iloc[int(column[0][1]), 1] + df_p_add.iloc[int(column[0][4]), 1]) * 1000
                else:
                    sens_df_add[column] = df_p_add.iloc[int(column[0][1]), 1] * 1000
            sens_df = pd.concat([sens_df, sens_df_add])
        count += 1

    writer = pd.ExcelWriter(os.getcwd() + '/results/loss_summary.xlsx', engine='xlsxwriter')

    # Write each dataframe to a different sheet in the workbook
    df_v.to_excel(writer, sheet_name='voltage')
    df_p.to_excel(writer, sheet_name='active_powers')
    df_q.to_excel(writer, sheet_name='reactive_powers')
    df_i.to_excel(writer, sheet_name='current')
    df_r.to_excel(writer, sheet_name='relaxation')
    df_losses.to_excel(writer, sheet_name='Port Losses')
    sens_df.to_excel(writer, sheet_name='Sensitivity')
    # Save the workbook and close the writer
    writer.close()


