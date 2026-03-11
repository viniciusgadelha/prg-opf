"""
prg_opf.mmc.losses — MMC conduction and switching loss calculation
===================================================================
Calculates total converter losses for a Modular Multilevel Converter (MMC)
based on IGBT/diode datasheet characteristics and the Marqueadt model.
"""

import math

import numpy as np

from prg_opf.mmc.parameters import (
    calculate_average_current,
    calculate_conduction_losses,
    calculate_currents,
    calculate_switching_losses,
    max_modulation_index,
    parameters,
)


def get_converter_characteristics(i_rms, u_rms, power_factor, freq):
    """Return IGBT, diode characteristics and device parameters."""

    # Infineon FZ1500R33HE3
    igbt_characteristic = {
        "e_on": [298.1e-3, 1.6744e-3, -4.2714e-7, 2.0444e-10],
        "e_off": [150.238e-3, 1.1877e-3, 9.0476e-8, 3.333e-12],
        "u_nominal": 3300,
        "v_ceo": [1.0115, 0.0021, -6.3e-7, 1.0518e-10],
        "i_nominal": 1500,
        "r_t": 12.6e-2,
    }

    diode_characteristic = {
        "e_rec": [348.8095e-3, 1.3496e-3, -3.8e-7, 3.7778e-11],
        "u_nominal": 3300,
        "v_do": [0.8625, 0.0021, -6.7063e-7, 1.0029e-10],
        "i_nominal": 1500,
        "r_d": 5.6e-2,
    }

    device_parameters = {
        "f_switch": freq,
        "u_dc": 10000,
        "apparent_power": u_rms * i_rms,
        "u_rms": u_rms,
        "power_factor": power_factor,
        "Number of submodules": 10,
    }

    return igbt_characteristic, diode_characteristic, device_parameters


def get_model_parameters(device):
    """Compute Marqueadt model parameters from device specification."""
    submodule_voltage = device["u_dc"] / device["Number of submodules"]
    k = max_modulation_index(device["u_dc"], device["u_rms"])
    i_ac = device["apparent_power"] / (device["u_rms"] * np.sqrt(3))
    return parameters(k, i_ac, device["u_dc"], device["Number of submodules"],
                      submodule_voltage, np.arccos(device["power_factor"]))


def cond_losses(marqueadt, igbt_characteristic, diode_characteristic):
    """Calculate three-phase conduction losses."""
    i_t1_d1, i_t2, i_d2 = calculate_currents(
        marqueadt['b'], marqueadt['x'], marqueadt['i_ac'],
        marqueadt['m'], marqueadt['i_dc'])

    t2_cond = calculate_conduction_losses(igbt_characteristic["v_ceo"], i_t2, marqueadt['i_equi_p'])
    d1_cond = calculate_conduction_losses(diode_characteristic["v_do"], i_t1_d1, marqueadt['i_equi_p'])
    t1_cond = calculate_conduction_losses(igbt_characteristic["v_ceo"], i_t1_d1, marqueadt['i_equi_n'])
    d2_cond = calculate_conduction_losses(diode_characteristic["v_do"], i_d2, marqueadt['i_equi_n'])

    igbt_cond = t1_cond + t2_cond
    diode_cond = d1_cond + d2_cond
    return igbt_cond + diode_cond, igbt_cond, diode_cond


def switch_losses(marqueadt, igbt_characteristic, diode_characteristic, device_parameters):
    """Calculate three-phase switching losses."""
    i_avg_t2_d1, i_avg_t1_d2 = calculate_average_current(marqueadt['m'], marqueadt['i_dc'])

    t1_on = calculate_switching_losses(
        igbt_characteristic["e_on"], device_parameters["f_switch"],
        marqueadt['u_sm'], igbt_characteristic["u_nominal"],
        i_avg_t1_d2, igbt_characteristic["i_nominal"])
    t1_off = calculate_switching_losses(
        igbt_characteristic["e_off"], device_parameters["f_switch"],
        marqueadt['u_sm'], igbt_characteristic["u_nominal"],
        i_avg_t1_d2, igbt_characteristic["i_nominal"])

    t2_on = calculate_switching_losses(
        igbt_characteristic["e_on"], device_parameters["f_switch"],
        marqueadt['u_sm'], igbt_characteristic["u_nominal"],
        i_avg_t2_d1, igbt_characteristic["i_nominal"])
    t2_off = calculate_switching_losses(
        igbt_characteristic["e_off"], device_parameters["f_switch"],
        marqueadt['u_sm'], igbt_characteristic["u_nominal"],
        i_avg_t2_d1, igbt_characteristic["i_nominal"])

    d1_sw = calculate_switching_losses(
        diode_characteristic["e_rec"], device_parameters["f_switch"],
        marqueadt['u_sm'], diode_characteristic["u_nominal"],
        i_avg_t2_d1, diode_characteristic["i_nominal"])
    d2_sw = calculate_switching_losses(
        diode_characteristic["e_rec"], device_parameters["f_switch"],
        marqueadt['u_sm'], diode_characteristic["u_nominal"],
        i_avg_t1_d2, diode_characteristic["i_nominal"])

    igbt_sw = (t1_on + t1_off) + (t2_on + t2_off)
    diode_sw = d1_sw + d2_sw
    return igbt_sw + diode_sw, igbt_sw, diode_sw


def calc_mmc_losses(p_mw=1, u_rms=36, power_factor=1, freq=200):
    """
    Calculate total MMC converter losses in MW.

    Parameters
    ----------
    p_mw : float
        Active power in MW.
    u_rms : float
        Voltage squared [kV²] (as used in the OPF linearization).
    power_factor : float
        Power factor (0 to 1).
    freq : float
        Switching frequency in Hz.

    Returns
    -------
    float
        Total losses (conduction + switching) in MW.
    """
    if p_mw == 0:
        return 0

    u_rms = math.sqrt(u_rms) * 1000  # convert kV² → V
    i_rms = abs(p_mw) * 1_000_000 / u_rms

    igbt, diode, device = get_converter_characteristics(i_rms, u_rms, power_factor, freq)
    marqueadt = get_model_parameters(device)
    total_cond, _, _ = cond_losses(marqueadt, igbt, diode)
    total_switch, _, _ = switch_losses(marqueadt, igbt, diode, device)

    # Three phases
    total_cond *= 3
    total_switch *= 3

    # Filter and cooling loss coefficients
    k1, k2 = 1.3, 1.3
    total_cond *= k1 * k2
    total_switch *= k1 * k2

    return (total_cond + total_switch) / 1_000_000  # W → MW


if __name__ == "__main__":
    total = calc_mmc_losses(1, 36, 1, freq=500)
    print(f"Total MMC losses: {total:.6f} MW")
