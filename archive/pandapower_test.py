import pandapower as pp
import pandas as pd
from pandapower.plotting import *
import math


def run_pf(p, q):
    net = pp.create_empty_network()

    pp.create_bus(net, 6, 1, geodata=(0, 0))
    pp.create_bus(net, 6, 2, geodata=(1, -0.5))
    pp.create_bus(net, 6, 3, geodata=(0, -1))

    #
    pp.create_ext_grid(net, 0)
    #

    pp.create_load(net, 2, p, q)

    pp.create_line_from_parameters(net, 0, 1, 5, 0.194, 0.315, 0.001, 0.35)
    pp.create_line_from_parameters(net, 1, 2, 5, 0.194, 0.315, 0.001, 0.35)
    pp.create_line_from_parameters(net, 0, 2, 10, 0.194, 0.315, 0.001, 0.35)


    pp.runpp(net)

    # print(net.res_bus)
    # print('----------------------------------')
    # print(net.res_line)

    # pf_res_plotly(net)

    return net.res_bus, net.res_line
if __name__ == "__main__":

    results_df = pd.DataFrame(index=np.arange(1, 0.7, 0.01), columns=['Line01', 'Line12', 'Line02', 'Bus0', 'Bus1', 'Bus2'])

    for p in np.arange(0.5, 3.5, 0.5):
        for pf in np.arange(1, 0.7, -0.01):
            q = p*math.tan(math.acos(pf))
            bus_df, line_df = run_pf(p, q)
            results_df.loc[pf, 'Line01'] = line_df.loc[0, 'pl_mw']
            results_df.loc[pf, 'Line12'] = line_df.loc[1, 'pl_mw']
            results_df.loc[pf, 'Line02'] = line_df.loc[2, 'pl_mw']
            results_df.loc[pf, 'Bus0'] = bus_df.loc[0, 'vm_pu']
            results_df.loc[pf, 'Bus1'] = bus_df.loc[1, 'vm_pu']
            results_df.loc[pf, 'Bus2'] = bus_df.loc[2, 'vm_pu']
        print(p)
        results_df.to_excel('results_pp_' + str(p) + '.xlsx')

