"""マッハ数別の空力係数CSVから、局所中立点と代表中立点を計算し、Excelに出力する。

入力CSVの例（1行目は読み飛ばす。迎角の単位は度）:
    aircraft_Mach0.3_aero.csv
    ここは任意の説明文
    alpha,CA,CN,Cm
    3,0.026,0.395,0.0242
    4,0.024,0.480,0.0156
    ...

位置座標: 機首原点、x=機尾方向正、y=右翼方向正、z=上方向正
空力係数: CA=機尾方向正、CN=上方向正、Cm=機首上げ正

必要ライブラリ: numpy, pandas, xlwings（デスクトップ版Excelが必要）
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
# 同名のExcelが存在する場合は更新、なければ新規作成する。
OUTPUT_EXCEL = Path(__file__).with_name("neutral_point_result.xlsx")

# 出力先シート名。既存ブックのシート名に合わせて変更してもよい。
# 各シートのA1から始まる表を上書きする（グラフ、書式、他シートは削除しない）。
SHEET_NAMES = ("x_NP", "h_NP", "x_NP_over_L", "NP_representative")

# 入力Cmのモーメント基準点 [m]
X_REF = 0.25
Y_REF = 0.0
Z_REF = 0.0

# 中立点を評価する高さ・左右位置 [m]（計算するのはx座標）
Z_NP = 0.05
Y_NP = Y_REF

# 25% MACのx座標とMAC長 [m]
X_MAC_25 = 0.25
C_MAC = 1.0
C_CM = C_MAC  # Cmの無次元化に用いた基準長 [m]

# 胴体長 [m]：以下はサンプル値。実機の値に変更する。
L_FUSELAGE = 5.0

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
        for col in ["alpha", "CA", "CN", "Cm"]:
            df[col] = pd.to_numeric(df[col], errors="raise")
        if not np.isfinite(df[["alpha", "CA", "CN", "Cm"]].to_numpy()).all():
            raise ValueError(f"{path.name}: 空力データに欠損値または無限大があります")

        df["source_file"] = path.name
        frames.append(df[["Mach", "alpha", "CA", "CN", "Cm", "source_file"]])

    if not frames:
        raise FileNotFoundError(f"{folder}: Machを含むCSVが見つかりません")
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
    if not all(np.isfinite(v) and v > 0 for v in (C_MAC, C_CM, L_FUSELAGE)):
        raise ValueError("C_MAC、C_CM、L_FUSELAGEは有限の正の値にしてください")

    if data.duplicated(["Mach", "alpha"]).any():
        raise ValueError("同じMach・迎角のデータが重複しています")

    results = []
    for _, group in data.sort_values(["Mach", "alpha"]).groupby("Mach"):
        g = group.copy()
        if len(g) < 3:
            raise ValueError("各Machで迎角データが3点以上必要です")

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
        # MAC前縁からの位置をMAC長で割る。25% MACを起点としても同じ結果。
        g["h_np"] = 0.25 + (x_np - X_MAC_25)/C_MAC
        # 機首原点のx座標を胴体長で無次元化する。
        g["x_np_over_L"] = x_np / L_FUSELAGE
        g["h_np_percent"] = g["h_np"]*100
        results.append(g)

    return pd.concat(results, ignore_index=True)


# ============================================================================
# 4. 代表中立点：指定迎角範囲の直線近似からMachごとに計算
# ============================================================================

def calculate_representative_neutral_points(data: pd.DataFrame) -> pd.DataFrame:
    """空力係数を迎角に対して直線近似し、Machごとの代表中立点を求める。"""
    results = []
    for mach, group in data.groupby("Mach"):
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
        h_np = 0.25 + (x_np - X_MAC_25)/C_MAC
        row = {"Mach": mach}
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
            "x_np_over_L": x_np / L_FUSELAGE,
            "h_np_percent": h_np*100,
        })
        results.append(row)

    return pd.DataFrame(results)


# ============================================================================
# 5. Excel出力：横持ち形式に変換して4シートに保存
# ============================================================================

def make_excel_tables(local: pd.DataFrame, representative: pd.DataFrame) -> dict:
    """迎角を行、Machを列にしたグラフ用の表を作る。"""
    tables = {}
    for name, column in zip(SHEET_NAMES[:3], ["x_np_m", "h_np", "x_np_over_L"]):
        # pivot: 縦持ちの計算結果を横持ちに変換する。
        table = local.pivot(index="alpha", columns="Mach", values=column)
        table = table.sort_index().sort_index(axis=1)
        table.columns = [f"Mach{m:g}" for m in table.columns]
        tables[name] = table.reset_index()

    # 代表中立点はMachごとに1行。3種類の位置と近似条件を併記する。
    tables[SHEET_NAMES[3]] = representative[[
        "Mach", "x_np_m", "h_np", "x_np_over_L", "h_np_percent",
        "fit_alpha_min_deg", "fit_alpha_max_deg", "n_fit_points",
        "CA_alpha_per_rad", "CN_alpha_per_rad", "Cm_alpha_per_rad",
    ]].sort_values("Mach")
    return tables


def save_excel(tables: dict, path: Path) -> None:
    """Excelを新規作成、または既存の同名シートを更新する。"""
    import xlwings as xw  # 計算だけのテストではExcelを必要としない。

    path = Path(path).resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    # xw.Bookは既存ブックが開いていれば接続し、なければ開く。
    # 新規ブックの場合は初期シートを再利用する。
    if path.exists():
        book = xw.Book(str(path))
    else:
        book = xw.Book()
        book.sheets[0].name = next(iter(tables))

    for name, table in tables.items():
        if name in [s.name for s in book.sheets]:
            sheet = book.sheets[name]
        else:
            sheet = book.sheets.add(name, after=book.sheets[-1])

        # A1から始まる旧データ表だけを消去する。
        # clear_contentsはセルの書式やグラフを削除しない。
        # 古いMach列や迎角行が残らないようにしてから書き込む。
        # この領域は出力専用にする（隣接する手入力データにも注意）。
        sheet.range("A1").current_region.clear_contents()

        # NaN（算出不能や欠測点）はExcelでは空白セルにする。
        # 列名も含めて一括で書き込み、余計なDataFrameのindexは出力しない。
        values = table.astype(object).where(pd.notna(table), None).values.tolist()
        sheet.range("A1").value = [table.columns.tolist()] + values

    # 既存ファイルはそのまま保存。新規ファイルには指定パスで名前を付ける。
    if path.exists():
        book.save()
    else:
        book.save(str(path))
    # 結果をすぐ確認できるよう、Excelは閉じずに残す。


# ============================================================================
# 6. 実行
# ============================================================================

def main() -> None:
    data = load_mach_csvs(INPUT_DIR)
    local = calculate_neutral_points(data)
    representative = calculate_representative_neutral_points(data)
    tables = make_excel_tables(local, representative)
    save_excel(tables, OUTPUT_EXCEL)
    print(f"Excel出力完了: {OUTPUT_EXCEL}")
    print(representative[["Mach", "x_np_m", "h_np", "x_np_over_L"]].to_string(index=False))


if __name__ == "__main__":
    main()
