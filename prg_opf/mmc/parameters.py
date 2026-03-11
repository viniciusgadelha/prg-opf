"""
prg_opf.mmc.parameters — Marqueadt MMC model parameters
=========================================================
Computes modulation index, arm currents, conduction/switching loss
parameters for modular multilevel converters.
"""

import numpy as np


def parameters(k, i_ac, v_dc, n, u_sm, phi):
    m = 2 / (k * np.cos(phi))
    i_dc = i_ac * (3 * np.pi) / (4 * m)
    x = (1 - m ** -2) ** (3 / 2)
    b = v_dc / (2 * n * u_sm)
    i_equi_p = (i_dc / 3) * (m + 1) * (np.pi / 4)  # arm (T2 and D1)
    i_equi_n = (i_dc / 3) * (m - 1) * (np.pi / 4)  # arm (T1 and D2)

    return {
        "m": m,
        "i_dc": i_dc,
        "i_ac": i_ac,
        "u_sm": u_sm,
        "x": x,
        "b": b,
        "i_equi_p": i_equi_p,
        "i_equi_n": i_equi_n,
    }


def calculate_currents(b, x, i_ac, m, i_dc):
    i_t1_d1 = 1 / 4 * b * x * i_ac
    i_t2 = 1 / 4 * (1 - b * x) * i_ac + 1 / 6 * i_dc * (1 + 1 / (3 * m))
    i_d2 = 1 / 4 * (1 - b * x) * i_ac - 1 / 6 * i_dc * (1 - 1 / (3 * m))
    return i_t1_d1, i_t2, i_d2


def calculate_conduction_losses(voltage, current, i_equi):
    voltage_calculated = third_order_approximation(voltage, current)
    resistance_calculated = voltage_calculated / current
    p_cond = voltage_calculated * current + resistance_calculated * current * i_equi
    return p_cond


def calculate_average_current(m, i_dc):
    i_avg_t2_d1 = (i_dc / 6) * (2 * m / np.pi + 1 + 1 / (3 * m))
    i_avg_t1_d2 = (i_dc / 6) * (2 * m / np.pi - 1 + 1 / (3 * m))
    return i_avg_t2_d1, i_avg_t1_d2


def max_modulation_index(dc_voltage, ac_voltage_rms):
    return ac_voltage_rms * np.sqrt(2) / dc_voltage


def third_order_approximation(coefficients, i_rms):
    """Polynomial approximation: a + b*x + c*x² + d*x³."""
    if isinstance(coefficients, float):
        return coefficients
    return (coefficients[0] + coefficients[1] * i_rms
            + coefficients[2] * i_rms ** 2 + coefficients[3] * i_rms ** 3)


def calculate_switching_losses(voltage_parameters, f_switch, u_device, u_nominal, i_avg, i_nominal):
    calculated_voltage = third_order_approximation(voltage_parameters, i_avg)
    return calculated_voltage * f_switch * (u_device / u_nominal) * (i_avg / i_nominal)
