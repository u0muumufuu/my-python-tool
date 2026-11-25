"""マッハ数・迎角ごとの空力係数CSVから、縦方向の局所中立点を求める。

【このプログラムで行うこと】
  1. CSVに記録された CA, CN, Cm（または CX, CZ, Cm）を読み込む。
  2. 同じMach（Re列があれば同じRe）のデータを迎角順に並べる。
  3. 迎角に対する各空力係数の傾き（微係数）を差分で計算する。
  4. 指定した高さ z_NP で、Cm_alpha = 0 となる前後位置 x_NP を求める。
  5. x_NP [m] と、MAC前縁から測った中立点位置 [%MAC] をCSVに保存する。

【位置座標】
  原点：機首（位置はすべて[m]）
  +x：機尾方向、+y：右翼方向、+z：上方向

【空力係数の符号】
  CA：機尾方向が正（CA = -CX）
  CN：上方向が正（CN = -CZ）
  Cm：機首上げが正（ピッチモーメント）
  CX：機首方向が正、CZ：下方向が正

【入力CSVに必要な列】
  Mach, alpha_deg, Cm, CA, CN
  または Mach, alpha_deg, Cm, CX, CZ
  ※ Re列があれば、MachとReの組合せごとに別々に計算する。
  ※ alpha_deg は「度」。内部ではradに変えて微分する。

【実行方法】
  pip install numpy pandas
  python neutral_point_from_csv.py

ファイル名・基準点・MAC情報などは、下の「設定」だけを書き換える。
"""
from pathlib import Path

import numpy as np
import pandas as pd


# ============================================================================
# 1. 設定：通常はこの範囲だけ編集すればよい
# ============================================================================

# 入力・出力CSV。初期設定では、このPythonファイルと同じフォルダに置く。
INPUT_CSV = Path(__file__).with_name("neutral_point_sample.csv")
OUTPUT_CSV = Path(__file__).with_name("neutral_point_result.csv")

# 【Cmを計算したモーメント基準点（reference point）】単位はm。
# CFD・風洞試験の出力Cmが「どの点周りか」を必ず確認する。
# 例：MAC前縁がx=0、MAC長が1mなら、x=0.25 は25% MACに相当する。
X_REF = 0.25
Y_REF = 0.0
Z_REF = 0.0

# 【中立点の高さ・左右位置】単位はm。
# Cm, CA, CNの迎角変化だけでは中立点の高さは一意に決まらない。
# そのため、zを固定し、その高さでのx_NPだけを計算する。
# y_NPもCmだけでは決められないので、値を指定して出力する。
Z_NP = 0.05
Y_NP = Y_REF

# 【MACに関する情報】単位はm。
# X_MAC_LE：平均空力翼弦（MAC）の「前縁」のx座標。翼根前縁とは限らない。
# C_MAC：平均空力翼弦長。この値を使って x_NP を%MACに変換する。
# C_CM：入力Cmを無次元化したときの基準長。通常は C_MAC と同じ。
X_MAC_LE = 0.0
C_MAC = 1.0
C_CM = C_MAC

# 【迎角に対する差分方法】
# 内部点は中心差分、両端は3点を使う2次精度の前進・後退差分を使用する。
# 迎角が不等間隔でも np.gradient が実際の格子間隔を考慮する。
# この方法は滑らかなデータに向くが、端点に大きな数値ノイズがある場合は注意。
# 迎角は各Mach（ReがあればMach・Re）の組合せごとに3点以上必要。

# CNの迎角微係数がほぼゼロの場合は、式の分母が小さくなって不安定。
# その点の中立点は NaN（計算不能扱い）にする。単位は 1/rad。
MIN_ABS_CN_ALPHA = 1e-3

# ============================================================================


def calculate_neutral_points(df: pd.DataFrame) -> pd.DataFrame:
    """入力の空力係数テーブルから、各Mach・迎角における局所中立点を求める。

    入力：pandas.DataFrame（上記のCSV列を持つ）
    出力：入力列に、迎角微係数・中立点座標・%MAC列を追加したDataFrame

    使用する中立点の式（位置座標は+x機尾、+z上）：

        Cm(x,z) = Cm(ref) + (x-x_ref)*CN/C_CM - (z-z_ref)*CA/C_CM

    中立点では、MachとReなどを固定して迎角を変えたとき、
    その点周りのピッチモーメントの傾きが0になる：

        dCm(x_NP,z_NP)/dalpha = 0

    よって、

        x_NP = x_ref + [ (z_NP-z_ref)*CA_alpha
                            - C_CM*Cm_alpha(ref) ] / CN_alpha

    h_NP = (x_NP - x_MAC_LE) / C_MAC
    [%MAC] = 100 * h_NP

    ※ 出てくるのは「指定したz_NPにおける局所中立点の前後位置」。
      動安定を評価したり、失速付近の非線形性を保証したりするものではない。
    """

    # ------------------------------------------------------------------------
    # 2. CSVの列名と設定値を確認する
    # ------------------------------------------------------------------------
    required = {"Mach", "alpha_deg", "Cm"}
    if not required.issubset(df.columns):
        raise ValueError(f"CSVに必要な列がありません: {sorted(required - set(df.columns))}")

    # CA/CN形式とCX/CZ形式のどちらでも入力可能。
    has_ca_cn = {"CA", "CN"}.issubset(df.columns)
    has_cx_cz = {"CX", "CZ"}.issubset(df.columns)
    if not (has_ca_cn or has_cx_cz):
        raise ValueError("CSVに CA,CN または CX,CZ の2列が必要です")

    # MAC長やCm基準長が0以下では、%MAC・中立点の計算ができない。
    if C_MAC <= 0 or C_CM <= 0:
        raise ValueError("C_MAC と C_CM は正の値にしてください")

    # Machごと（Re列があればMachとReごと）に迎角微分する。
    # 同じ組合せで迎角が重複すると差分の間隔が0になるので禁止する。
    group_cols = ["Mach"] + (["Re"] if "Re" in df.columns else [])
    if df.duplicated(group_cols + ["alpha_deg"]).any():
        raise ValueError("Mach・Re・迎角が同じデータが重複しています")

    # ------------------------------------------------------------------------
    # 3. 空力係数の符号規約を CA・CN に統一する
    # ------------------------------------------------------------------------
    # 入力のDataFrameを直接書き換えないようにコピーする。
    data = df.copy()
    if has_ca_cn:
        # 両方の形式が同時に存在するときは、符号が整合するか確認する。
        if has_cx_cz and not (
            np.allclose(data["CA"], -data["CX"])
            and np.allclose(data["CN"], -data["CZ"])
        ):
            raise ValueError("CA=-CX, CN=-CZ の対応が一致していません")
    else:
        # CX（機首正）→CA（機尾正）、CZ（下正）→CN（上正）。
        data["CA"] = -data["CX"]
        data["CN"] = -data["CZ"]

    # 迎角の並び順がCSVでバラバラでも正しく微分できるようソートする。
    data = data.sort_values(group_cols + ["alpha_deg"])
    results = []

    # ------------------------------------------------------------------------
    # 4. Mach（必要ならRe）を固定して、迎角方向に微分する
    # ------------------------------------------------------------------------
    for _, g in data.groupby(group_cols, sort=False):
        g = g.copy()

        # 3点を使う端点の2次精度差分には、各グループ3点以上必要。
        if len(g) < 3:
            raise ValueError(f"迎角点が不足しています: {g[group_cols].iloc[0].to_dict()}")

        # np.gradient には迎角の「実際の格子点」を渡す。
        # 角度はdeg→radに変換するため、微係数の単位はすべて 1/rad。
        alpha = np.deg2rad(g["alpha_deg"].to_numpy(dtype=float))

        # 内部点：前後の迎角から中心差分。
        # 最初・最後：それぞれ3点で計算する2次精度の前進・後退差分。
        # 例）迎角が4,5,6°なら、5°での傾きは概ね[f(6°)-f(4°)]/2°。
        ca_a = np.gradient(g["CA"].to_numpy(dtype=float), alpha, edge_order=2)
        cn_a = np.gradient(g["CN"].to_numpy(dtype=float), alpha, edge_order=2)
        cm_a = np.gradient(g["Cm"].to_numpy(dtype=float), alpha, edge_order=2)

        # --------------------------------------------------------------------
        # 5. 中立点条件 Cm_alpha(x_NP, z_NP) = 0 を解く
        # --------------------------------------------------------------------
        # 高さZ_NPは冒頭設定値で固定する。
        # x_NP = x_ref + [(z_NP-z_ref)*CA_alpha - C_CM*Cm_alpha]/CN_alpha
        with np.errstate(divide="ignore", invalid="ignore"):
            x_np = X_REF + ((Z_NP - Z_REF) * ca_a - C_CM * cm_a) / cn_a

        # CN_alphaがほぼ0では、計算結果が極端な値になりやすいため無効とする。
        x_np[np.abs(cn_a) < MIN_ABS_CN_ALPHA] = np.nan

        # --------------------------------------------------------------------
        # 6. 結果をテーブルに追加する
        # --------------------------------------------------------------------
        # 微係数も保存しておけば、結果の検算・プロットがしやすい。
        g["CA_alpha_per_rad"] = ca_a
        g["CN_alpha_per_rad"] = cn_a
        g["Cm_alpha_per_rad"] = cm_a

        # 中立点の絶対座標 [m]（機首原点、x機尾正、y右翼正、z上正）。
        g["x_np_m"] = x_np
        g["y_np_m"] = Y_NP  # 指定値。Cmのみから求めた値ではない。
        g["z_np_m"] = Z_NP  # 指定値。

        # MAC前縁を0%、そこから機尾方向にC_MAC進んだ点を100%とする。
        g["h_np"] = (x_np - X_MAC_LE) / C_MAC
        g["h_np_percent"] = g["h_np"] * 100
        results.append(g)

    # Mach/Reごとの結果を縦につなげて一つのCSV用テーブルにする。
    return pd.concat(results, ignore_index=True)


# ============================================================================
# 7. 実行部分（関数を別のPythonからimportしても、勝手には実行されない）
# ============================================================================
if __name__ == "__main__":
    # CSVを読み込み、中立点を計算し、結果を保存する。
    result = calculate_neutral_points(pd.read_csv(INPUT_CSV))

    # utf-8-sig にすると、Excelなどでも日本語列名が文字化けしにくい。
    result.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")

    print(f"保存しました: {OUTPUT_CSV}")
    # 主要な列だけコンソールにも表示する。詳細結果は出力CSVを参照。
    print(result[["Mach", "alpha_deg", "x_np_m", "h_np_percent"]].to_string(index=False))
