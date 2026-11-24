"""年ごとのZIP(年.zip、中身は年.csv)から、各月の最高水位・最低水位を集計する(計算部分)。

GUIからは monthly_hl_gui.py 経由で使う。コマンドラインからも使える:
    python monthly_hl.py ZIPを入れたフォルダ 出力.xlsx
"""
import sys
import zipfile
from pathlib import Path

import pandas as pd

# 既定で集計に使う属性(0:有効のみ)。凡例: 0:有効 1:欠測 2:推定値 3:閉局 8:工事中水位
DEFAULT_ATTR = ("0",)


def read_one(src, name: str) -> pd.DataFrame:
    """CSV1個を読み込む。属性は先頭0が消えないよう文字列で読む。文字コードはcp932(Shift_JIS)。"""
    df = pd.read_csv(src, encoding="cp932", dtype={"属性": str, "暫定フラグ": str})
    df["元ファイル"] = name
    return df


def load(folder: Path) -> tuple[pd.DataFrame, int]:
    """フォルダ内のZIP(と直下のCSV)を読み込んで1つの表にする。戻り値: (表, 読み込んだCSV数)"""
    frames = []
    # ZIPは展開せず、中のCSVをメモリ上で直接読む(ディスクを汚さず、不正なパスの書き出しも防げる)
    for z in sorted(folder.glob("*.zip")):
        with zipfile.ZipFile(z) as zf:
            for member in zf.namelist():
                if member.lower().endswith(".csv"):
                    with zf.open(member) as fp:
                        frames.append(read_one(fp, f"{z.name}:{member}"))
    # 展開済みのCSVがフォルダ直下にあればそれも読む
    for f in sorted(folder.glob("*.csv")):
        frames.append(read_one(f, f.name))
    if not frames:
        raise FileNotFoundError(f"ZIP/CSVが見つかりません: {folder}")
    return pd.concat(frames, ignore_index=True), len(frames)


def aggregate(folder: Path, valid_attr=DEFAULT_ATTR) -> tuple[pd.DataFrame, list[str]]:
    """月別の最高・最低水位を集計する。戻り値: (結果の表, 確認用メッセージのリスト)"""
    messages = []
    df, n_files = load(folder)
    messages.append(f"読み込んだCSV数: {n_files}")

    n0 = len(df)
    df = df.drop_duplicates(subset=["日時", "データ", "属性"])
    if len(df) != n0:
        messages.append(f"重複行を{n0 - len(df)}件除外しました")

    # このデータは日の終わりを「24:00」と書くため、通常の日時変換は失敗する。
    # 日付部分 + 時刻(時間・分)の足し算で変換し、24:00は翌日00:00として扱う。
    df["日時表記"] = df["日時"].str.strip()
    date_part = pd.to_datetime(df["日時表記"].str[:10], format="%Y/%m/%d")
    hh = df["日時表記"].str[11:13].astype(int)
    mm = df["日時表記"].str[14:16].astype(int)
    df["日時"] = date_part + pd.to_timedelta(hh, unit="h") + pd.to_timedelta(mm, unit="m")
    df["属性"] = df["属性"].fillna("").str.strip()

    # --- データ品質の確認(除外前の件数) ---
    counts = df["属性"].value_counts()
    messages.append("属性ごとの件数: " + ", ".join(f"{k}={v}" for k, v in counts.items()))
    messages.append(f"-9999の件数: {int((df['データ'] == -9999).sum())}")

    # --- 無効値を除外 ---
    ok = df[df["属性"].isin(valid_attr) & (df["データ"] != -9999)].copy()
    if ok.empty:
        raise ValueError("集計できる有効データがありません(属性の選択を確認してください)")

    # 24:00(=翌日00:00)は前日の終わりなので、1分戻した時刻で月を判定する
    ok["年月"] = (ok["日時"] - pd.Timedelta(minutes=1)).dt.to_period("M")

    rows = []
    for ym, g in ok.groupby("年月"):
        hi = g.loc[g["データ"].idxmax()]
        lo = g.loc[g["データ"].idxmin()]
        rows.append({
            "年月": str(ym),
            "最高水位(m)": hi["データ"], "最高の日時": hi["日時表記"],
            "最低水位(m)": lo["データ"], "最低の日時": lo["日時表記"],
            "有効データ数": len(g),
            "暫定値を含む": "*" in set(g["暫定フラグ"].fillna("")),
        })
    result = pd.DataFrame(rows)

    # 有効データ数が少ない月に目印を付ける(欠測が多い月・途中の月の見落とし防止)
    result["注意"] = ""
    result.loc[result["有効データ数"] < result["有効データ数"].max() * 0.9, "注意"] = "データ数少"
    return result, messages


def save_excel(result: pd.DataFrame, out: Path) -> None:
    """結果をExcelに保存する。"""
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        result.to_excel(xw, sheet_name="月別最高最低", index=False)
        ws = xw.sheets["月別最高最低"]
        for col, w in zip("ABCDEFGH", (10, 14, 18, 14, 18, 14, 14, 12)):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A2"


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    res, msgs = aggregate(Path(sys.argv[1]))
    print("\n".join(msgs))
    print(res.to_string(index=False))
    save_excel(res, Path(sys.argv[2]))
    print("出力:", sys.argv[2])
