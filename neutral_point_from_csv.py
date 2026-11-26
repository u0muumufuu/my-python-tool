"""マッハ数別の空力係数CSVから、局所中立点と代表中立点を計算する。

入力CSVの例（1行目は読み飛ばす。迎角の単位は度）:
    aircraft_Mach0.3_aero.csv
    ここは任意の説明文
    alpha,CA,CN,Cm
    3,0.026,0.395,0.0242
    4,0.024,0.480,0.0156
    ...

位置座標: 機首原点、x=機尾方向正、y=右翼方向正、z=上方向正
空力係数: CA=機尾方向正、CN=上方向正、Cm=機首上げ正

実行: python neutral_point_from_csv.py
"""
from pathlib import Path
import re

import numpy as np
import pandas as pd


# ============================================================================
# 1. 設定：変更するのは基本的にここだけ
# ============================================================================

INPUT_DIR = Path(__file__).resolve().parent / "aero_csv"
OUTPUT_LOCAL_CSV = Path(__file__).with_name("neutral_point_result.csv")
OUTPUT_REP_CSV = Path(__file__).with_name("neutral_point_representative.csv")

# 入力Cmのモーメント基準点 [m]
X_REF = 0.25
Y_REF = 0.0
Z_REF = 0.0

# 中立点を評価する高さ・左右位置 [m]（計算するのはx座標）
Z_NP = 0.05
Y_NP = Y_REF

# MAC前縁のx座標、MAC長、およびCmを無次元化した基準長 [m]
X_MAC_LE = 0.0
C_MAC = 1.0
C_CM = C_MAC

# 代表中立点を求める直線近似の迎角範囲 [deg]
FIT_ALPHA_MIN_DEG = -2.0
FIT_ALPHA_MAX_DEG = 8.0

# CNの迎角微係数が小さすぎる場合は中立点をNaNとする [1/rad]
MIN_ABS_CN_ALPHA = 1e-3


# ============================================================================
# 2. CSVの読み込み
# ============================================================================

# 例: aircraft_Mach0.3_aero.csv → Mach=0.3
MACH_PATTERN = re.compile(r"(?:^|[_-])Mach(\d+(?:\.\d+)?)(?=[_.-]|$)", re.I)


def mach_from_filename(filename: str) -> float:
    """CSVのファイル名に含まれるMach値を取得する。"""
    match = MACH_PATTERN.search(Path(filename).stem)
    if match is None:
        raise ValueError(f"ファイル名からMach値を取得できません: {filename}")
    return float(match.group(1))


def load_mach_csvs(folder: Path) -> pd.DataFrame:
    """各MachのCSVを読み込み、1つの表にまとめる。

    必須列は alpha, Cm と CA,CN（またはCX,CZ）。
    alphaのみ対応し、alpha_degやAoAは受け付けない。
    CSVの1行目を飛ばし、2行目を列名として使用する。
    """
    if not folder.is_dir():
        raise FileNotFoundError(f"入力フォルダがありません: {folder}")

    frames = []
    for path in sorted(folder.glob("*.csv")):
        if MACH_PATTERN.search(path.stem) is None:
            continue  # Machを含まないCSVは対象外

        # skiprows=1: 1行目を読み飛ばし、2行目をヘッダーにする。
        df = pd.read_csv(path, skiprows=1, encoding="utf-8-sig")
        df.columns = df.columns.str.strip()  # 列名の前後の空白のみ削除

        if not {"alpha", "Cm"}.issubset(df.columns):
            raise ValueError(f"{path.name}: 'alpha' と 'Cm' 列が必要です")

        # CA,CNがない場合は CX,CZ を符号反転して変換する。
        if not {"CA", "CN"}.issubset(df.columns):
            if not {"CX", "CZ"}.issubset(df.columns):
                raise ValueError(f"{path.name}: CA,CN または CX,CZ 列が必要です")
            df["CA"] = -pd.to_numeric(df["CX"], errors="raise")
            df["CN"] = -pd.to_numeric(df["CZ"], errors="raise")

        # ファイル名からMachを設定し、必要な列を数値に変換する。
        df["Mach"] = mach_from_filename(path.name)
        for col in ["alpha", "CA", "CN", "Cm"] + (["Re"] if "Re" in df else []):
            df[col] = pd.to_numeric(df[col], errors="raise")
        if not np.isfinite(df[["alpha", "CA", "CN", "Cm"]].to_numpy()).all():
            raise ValueError(f"{path.name}: 空力データに欠損値または無限大があります")
        if "Re" in df and not np.isfinite(df["Re"].to_numpy()).all():
            raise ValueError(f"{path.name}: Reに欠損値または無限大があります")

        df["source_file"] = path.name
        frames.append(df[["Mach", "alpha", "CA", "CN", "Cm"]
                         + (["Re"] if "Re" in df else []) + ["source_file"]])

    if not frames:
        raise FileNotFoundError(f"{folder}: Machを含むCSVが見つかりません")
    if any("Re" in f for f in frames) and not all("Re" in f for f in frames):
        raise ValueError("Re列を使う場合は、すべてのCSVにRe列が必要です")
    return pd.concat(frames, ignore_index=True)


# ============================================================================
# 3. 局所中立点：各迎角における微係数から計算
# ============================================================================

def calculate_neutral_points(data: pd.DataFrame) -> pd.DataFrame:
    """各Mach・迎角に対して局所中立点を求める。

    Cm(x,z) = Cm_ref + (x-x_ref)*CN/C_CM - (z-z_ref)*CA/C_CM
    中立点で Cm_alpha=0 と置くと、
    x_NP = x_ref + [(z_NP-z_ref)*CA_alpha - C_CM*Cm_alpha]/CN_alpha
    """
    if C_MAC <= 0 or C_CM <= 0:
        raise ValueError("C_MACとC_CMは正の値にしてください")

    group_cols = ["Mach"] + (["Re"] if "Re" in data else [])
    if data.duplicated(group_cols + ["alpha"]).any():
        raise ValueError("同じMach・Re・迎角のデータが重複しています")

    results = []
    for _, group in data.sort_values(group_cols + ["alpha"]).groupby(group_cols):
        g = group.copy()
        if len(g) < 3:
            raise ValueError("各Mach・Reで迎角データが3点以上必要です")

        # 迎角を度からradへ変換。刻みが不等間隔でも差分可能。
        alpha_rad = np.deg2rad(g["alpha"].to_numpy())

        # 内部点：2次精度の中心差分
        # 最初と最後：3点を使う2次精度の前進・後退差分
        ca_a = np.gradient(g["CA"].to_numpy(), alpha_rad, edge_order=2)
        cn_a = np.gradient(g["CN"].to_numpy(), alpha_rad, edge_order=2)
        cm_a = np.gradient(g["Cm"].to_numpy(), alpha_rad, edge_order=2)

        with np.errstate(divide="ignore", invalid="ignore"):
            x_np = X_REF + ((Z_NP - Z_REF)*ca_a - C_CM*cm_a)/cn_a
        x_np[np.abs(cn_a) < MIN_ABS_CN_ALPHA] = np.nan

        g["CA_alpha_per_rad"] = ca_a
        g["CN_alpha_per_rad"] = cn_a
        g["Cm_alpha_per_rad"] = cm_a
        g["x_np_m"] = x_np
        g["y_np_m"] = Y_NP  # y,zは指定値。空力係数から求めていない。
        g["z_np_m"] = Z_NP
        g["h_np"] = (x_np - X_MAC_LE)/C_MAC
        g["h_np_percent"] = g["h_np"]*100
        results.append(g)

    return pd.concat(results, ignore_index=True)


# ============================================================================
# 4. 代表中立点：指定迎角範囲の直線近似からMachごとに計算
# ============================================================================

def calculate_representative_neutral_points(data: pd.DataFrame) -> pd.DataFrame:
    """空力係数を迎角に対して直線近似し、Machごとの代表中立点を求める。"""
    group_cols = ["Mach"] + (["Re"] if "Re" in data else [])
    results = []
    for _, group in data.groupby(group_cols):
        g = group[group["alpha"].between(FIT_ALPHA_MIN_DEG, FIT_ALPHA_MAX_DEG)]
        if len(g) < 3:
            raise ValueError(f"Mach={group['Mach'].iloc[0]}: 直線近似範囲に迎角3点以上が必要です")

        alpha_rad = np.deg2rad(g["alpha"].to_numpy())
        # 最小二乗直線の傾きが迎角微係数（1/rad）となる。
        ca_a = np.polyfit(alpha_rad, g["CA"], 1)[0]
        cn_a = np.polyfit(alpha_rad, g["CN"], 1)[0]
        cm_a = np.polyfit(alpha_rad, g["Cm"], 1)[0]

        x_np = (np.nan if abs(cn_a) < MIN_ABS_CN_ALPHA else
                X_REF + ((Z_NP - Z_REF)*ca_a - C_CM*cm_a)/cn_a)
        h_np = (x_np - X_MAC_LE)/C_MAC
        row = {col: g[col].iloc[0] for col in group_cols}
        row.update({
            "fit_alpha_min_deg": FIT_ALPHA_MIN_DEG,
            "fit_alpha_max_deg": FIT_ALPHA_MAX_DEG,
            "n_fit_points": len(g),
            "CA_alpha_per_rad": ca_a,
            "CN_alpha_per_rad": cn_a,
            "Cm_alpha_per_rad": cm_a,
            "x_np_m": x_np,
            "y_np_m": Y_NP,
            "z_np_m": Z_NP,
            "h_np": h_np,
            "h_np_percent": h_np*100,
        })
        results.append(row)

    return pd.DataFrame(results)


# ============================================================================
# 5. 実行
# ============================================================================

def main() -> None:
    data = load_mach_csvs(INPUT_DIR)
    local = calculate_neutral_points(data)
    representative = calculate_representative_neutral_points(data)
    local.to_csv(OUTPUT_LOCAL_CSV, index=False, encoding="utf-8-sig")
    representative.to_csv(OUTPUT_REP_CSV, index=False, encoding="utf-8-sig")
    print(f"局所中立点: {OUTPUT_LOCAL_CSV}")
    print(f"代表中立点: {OUTPUT_REP_CSV}")
    print(representative[["Mach", "x_np_m", "h_np_percent"]].to_string(index=False))


if __name__ == "__main__":
    main()
