"""T–V・Ps のサンプル計算。モデルは外部から渡され、Excelを直接扱わない。"""

from __future__ import annotations

import numpy as np
import pandas as pd

from utils.solvers import EquationSolver

G = 9.80665


def calculate_tv_ps(
    aero,
    propulsion,
    *,
    mass_kg: float,
    wing_area_m2: float,
    rho_kg_m3: float,
    speed_of_sound_m_s: float,
    altitude_m: float,
    velocities_m_s,
) -> pd.DataFrame:
    """L=Wで迎角を決め、Treq≈DとPsを計算する簡易定常モデル。

    本例は Cm=0（縦トリム）、尾翼・舵角・推力方向・迎角による
    抵抗補正等を考慮していない。実用段階では完全なトリム計算に置換する。
    """
    numbers = [mass_kg, wing_area_m2, rho_kg_m3, speed_of_sound_m_s]
    if not all(np.isfinite(numbers)) or any(x <= 0 for x in numbers):
        raise ValueError("質量、翼面積、空気密度、音速は正の有限数が必要です")
    if not np.isfinite(altitude_m):
        raise ValueError("高度は有限数が必要です")
    velocity_array = np.atleast_1d(np.asarray(velocities_m_s, dtype=float))
    if velocity_array.ndim != 1 or len(velocity_array) == 0 or not np.isfinite(velocity_array).all() or np.any(velocity_array <= 0):
        raise ValueError("速度は正の有限数からなる1次元配列を指定してください")

    weight = mass_kg * G
    solver = EquationSolver(residual_tol=1e-8)
    output = []
    for velocity in velocity_array:
        velocity = float(velocity)
        mach = velocity / speed_of_sound_m_s
        q = 0.5 * rho_kg_m3 * velocity**2
        cl_req = weight / (q * wing_area_m2)
        row = {
            "V": velocity, "Mach": mach, "CL_required": cl_req,
            "alpha": np.nan, "CD": np.nan,
            "T_required": np.nan, "T_available": np.nan, "Ps": np.nan,
            "status": "",
        }

        if not aero.supports_mach(mach) or not propulsion.supports(velocity=velocity, altitude=altitude_m):
            row["status"] = "OUT_OF_MAP"
            output.append(row)
            continue

        def cl_residual(alpha: float) -> float:
            return aero.coefficients(alpha=alpha, mach=mach)["CL"] - cl_req

        solution = solver.solve(cl_residual, bracket=aero.alpha_bounds)
        if not solution.converged:
            row["status"] = "NO_LIFT_BALANCE"
            output.append(row)
            continue

        alpha = float(solution.x)
        coeffs = aero.coefficients(alpha=alpha, mach=mach)
        cd = float(coeffs["CD"])
        thrust_available = propulsion.thrust(velocity=velocity, altitude=altitude_m)
        if not np.isfinite(cd) or not np.isfinite(thrust_available):
            row["status"] = "MISSING_DATA"
            output.append(row)
            continue

        drag = q * wing_area_m2 * cd
        ps = (thrust_available - drag) * velocity / weight
        row.update(alpha=alpha, CD=cd, T_required=drag, T_available=thrust_available, Ps=ps, status="OK")
        output.append(row)

    return pd.DataFrame(output)
