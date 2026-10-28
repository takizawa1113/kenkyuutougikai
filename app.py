"""
研究棟議会 発表割り当てサポート（叩き版）

起動: streamlit run app.py
データ: Excel（シート「発表履歴」「研究テーマ」）または CSV 2つを読み込み、
        編集後に Excel / CSV でダウンロードする。コード内にダミーデータは持たない。
"""
import io
import re
from datetime import date

import pandas as pd
import streamlit as st

st.set_page_config(page_title="研究棟議会 発表割り当て", layout="wide")

HIST_SHEET = "発表履歴"
THEME_SHEET = "研究テーマ"
MONTH_RE = re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月")


# ---------- ユーティリティ ----------
def parse_month(label: str):
    """'2026年10月' -> (2026, 10)。解釈できなければ None"""
    m = MONTH_RE.search(str(label))
    return (int(m.group(1)), int(m.group(2))) if m else None


def month_label(y: int, m: int) -> str:
    return f"{y}年{m}月"


def month_index(ym) -> int:
    return ym[0] * 12 + ym[1] - 1


def add_months(ym, n):
    i = month_index(ym) + n
    return (i // 12, i % 12 + 1)


def is_blank(v) -> bool:
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() in ("", "nan", "None")


def clean(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    df = df.astype(object).where(pd.notna(df), "")
    first = df.columns[0]
    df = df.rename(columns={first: "メンバー"})
    df["メンバー"] = df["メンバー"].astype(str).str.strip()
    df = df[df["メンバー"] != ""]
    return df.reset_index(drop=True)


def sort_month_columns(df: pd.DataFrame) -> pd.DataFrame:
    months = [c for c in df.columns if c != "メンバー" and parse_month(c)]
    others = [c for c in df.columns if c != "メンバー" and not parse_month(c)]
    months.sort(key=lambda c: month_index(parse_month(c)))
    return df[["メンバー"] + months + others]


def sync_members(hist: pd.DataFrame, themes: pd.DataFrame):
    """どちらかの表にしかいないメンバーを、もう片方に空行で追加"""
    all_members = list(dict.fromkeys(list(themes["メンバー"]) + list(hist["メンバー"])))
    def ensure(df):
        missing = [m for m in all_members if m not in set(df["メンバー"])]
        if missing:
            add = pd.DataFrame({c: [""] * len(missing) for c in df.columns})
            add["メンバー"] = missing
            df = pd.concat([df, add], ignore_index=True)
        return df
    return ensure(hist), ensure(themes)


def load_data(hist: pd.DataFrame, themes: pd.DataFrame):
    hist, themes = clean(hist), clean(themes)
    hist, themes = sync_members(hist, themes)
    st.session_state.hist = sort_month_columns(hist)
    st.session_state.themes = themes
    st.session_state.loaded = True
    st.session_state.ver = st.session_state.get("ver", 0) + 1


# ---------- リコメンドロジック ----------
def build_recommendations(hist: pd.DataFrame, themes: pd.DataFrame, target_ym, min_gap: int):
    month_cols = [c for c in hist.columns if parse_month(c)]
    rows = []
    for _, r in themes.iterrows():
        member = r["メンバー"]
        h = hist[hist["メンバー"] == member]
        h = h.iloc[0] if len(h) else pd.Series(dtype=object)

        # 対象月より前の発表を抽出
        past = []  # (ym, 発表内容)
        for c in month_cols:
            ym = parse_month(c)
            if month_index(ym) < month_index(target_ym) and not is_blank(h.get(c, "")):
                past.append((ym, str(h[c])))
        past.sort(key=lambda x: month_index(x[0]))

        already_assigned = not is_blank(h.get(month_label(*target_ym), ""))

        if past:
            last_ym, last_title = past[-1]
            gap = month_index(target_ym) - month_index(last_ym)
            last_str = f"{month_label(*last_ym)}（{last_title}）"
        else:
            gap, last_str = None, "発表実績なし"

        # テーマ案: 未発表のテーマを優先、全部発表済みなら一番昔に発表したもの
        member_themes = [str(r[c]).strip() for c in themes.columns if c != "メンバー" and not is_blank(r[c])]
        theme_last = {}
        for t in member_themes:
            hits = [ym for ym, title in past if t in title]
            theme_last[t] = max(hits, key=month_index) if hits else None
        unpresented = [t for t in member_themes if theme_last[t] is None]
        if unpresented:
            suggestion, reason = unpresented[0], "未発表テーマ"
        elif member_themes:
            suggestion = min(member_themes, key=lambda t: month_index(theme_last[t]))
            reason = f"前回 {month_label(*theme_last[suggestion])} に発表"
        else:
            suggestion, reason = "", "テーマ未登録"

        rows.append({
            "メンバー": member,
            "前回発表": last_str,
            "空き(月)": gap,
            "発表回数": len(past),
            "テーマ案": suggestion,
            "テーマ案の根拠": reason,
            "対象月に割当済": "✓" if already_assigned else "",
            "_sort": 10_000 if gap is None else gap,
        })

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df = df[df["_sort"] >= min_gap]
    # 割当済みは後ろ、空きが長い順、発表回数が少ない順
    df = df.sort_values(["対象月に割当済", "_sort", "発表回数"], ascending=[True, False, True])
    return df.drop(columns="_sort").reset_index(drop=True)


# ---------- 出力 ----------
def to_excel_bytes(hist, themes) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        hist.to_excel(w, sheet_name=HIST_SHEET, index=False)
        themes.to_excel(w, sheet_name=THEME_SHEET, index=False)
    return buf.getvalue()


def to_csv_bytes(df) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")  # Excelで文字化けしないようBOM付き


def read_csv_any(f) -> pd.DataFrame:
    raw = f.getvalue()
    for enc in ("utf-8-sig", "cp932"):
        try:
            return pd.read_csv(io.BytesIO(raw), encoding=enc, dtype=str)
        except UnicodeDecodeError:
            continue
    raise ValueError("CSVの文字コードを判別できませんでした")


# ---------- サイドバー: データ入出力 ----------
st.sidebar.header("① データ読み込み")
src = st.sidebar.radio("形式", ["Excel（2シート）", "CSV（2ファイル）"], horizontal=True)

def try_load(key, loader):
    """新しいファイルがアップロードされた時だけ自動で読み込む。
    「アップロード時の状態に戻す」を押した時は同じファイルでも読み直す。"""
    force = st.session_state.pop("force_reload", False)
    if key and (force or st.session_state.get("last_file_key") != key):
        st.session_state.last_file_key = key
        try:
            loader()
            st.session_state.load_msg = st.session_state.pop("swap_msg", "ファイルを読み込みました")
        except Exception as e:
            st.session_state.loaded = False
            st.sidebar.error(f"読み込み失敗: {e}")


def count_month_cols(df):
    return sum(1 for c in df.columns if parse_month(c))


def load_checked(hist, themes):
    """発表履歴と研究テーマの取り違えを検出してから読み込む"""
    h, t = count_month_cols(hist), count_month_cols(themes)
    if t > 0 and h == 0:
        hist, themes = themes, hist
        st.session_state.swap_msg = "発表履歴と研究テーマが逆だったため、入れ替えて読み込みました"
    elif t > 0:
        raise ValueError("研究テーマ側に「◯年◯月」の列があります。発表履歴のファイルを2回入れていないか確認してください")
    load_data(hist, themes)


def load_excel(up):
    sheets = pd.read_excel(up, sheet_name=None, dtype=str)
    if HIST_SHEET not in sheets or THEME_SHEET not in sheets:
        raise ValueError(f"シート名は「{HIST_SHEET}」「{THEME_SHEET}」にしてください（見つかったシート: {list(sheets)}）")
    load_checked(sheets[HIST_SHEET], sheets[THEME_SHEET])


if src.startswith("Excel"):
    up = st.sidebar.file_uploader(f"シート「{HIST_SHEET}」「{THEME_SHEET}」を含む .xlsx", type=["xlsx"])
    try_load(up.file_id if up else None, lambda: load_excel(up))
else:
    up_h = st.sidebar.file_uploader("発表履歴の .csv", type=["csv"], key="csv_h")
    up_t = st.sidebar.file_uploader("研究テーマの .csv", type=["csv"], key="csv_t")
    if (up_h is None) != (up_t is None):
        st.sidebar.info("2つ目のCSVもアップロードしてください")
    key = f"{up_h.file_id}|{up_t.file_id}" if up_h and up_t else None
    try_load(key, lambda: load_checked(read_csv_any(up_h), read_csv_any(up_t)))

if st.session_state.get("load_msg"):
    st.sidebar.success(st.session_state.pop("load_msg"))

if st.session_state.get("last_file_key") and st.sidebar.button("アップロード時の状態に戻す"):
    st.session_state.force_reload = True
    st.rerun()

if st.sidebar.button("空の状態から始める"):
    load_data(pd.DataFrame({"メンバー": []}), pd.DataFrame({"メンバー": [], "テーマ①": [], "テーマ②": [], "テーマ③": []}))

# ---------- メイン ----------
st.title("研究棟議会 発表割り当てサポート")

if not st.session_state.get("loaded"):
    st.info("左のサイドバーで Excel / CSV をアップロードすると、その内容で始まります。データなしで試す場合は「空の状態から始める」を押してください。")
    st.stop()

tab_rec, tab_hist, tab_theme = st.tabs(["🎯 リコメンド", "📅 発表履歴", "🧪 研究テーマ"])
ver = st.session_state.get("ver", 0)


CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"


def next_theme_col(df):
    for ch in CIRCLED:
        if f"テーマ{ch}" not in df.columns:
            return f"テーマ{ch}"
    return f"テーマ{len(df.columns)}"


def delete_columns_ui(df, label, key):
    """列を選んで削除する。入力済みのセルがある列は確認チェックを必須にする"""
    cols = [c for c in df.columns if c != "メンバー"]
    with st.expander(f"{label}の列を削除"):
        targets = st.multiselect("削除する列", cols, key=f"del_{key}_{ver}")
        filled = sum(int((df[c].astype(str).str.strip() != "").sum()) for c in targets)
        ok = True
        if filled:
            ok = st.checkbox(f"入力済みのセルが {filled} 件あります。消えてもよい", key=f"delok_{key}_{ver}")
        if st.button("選択した列を削除", key=f"delbtn_{key}", disabled=not targets or not ok):
            return df.drop(columns=targets)
    return None


def commit(hist, themes):
    """ボタン操作による変更を確定し、エディタを作り直す"""
    hist, themes = sync_members(clean(hist), clean(themes))
    st.session_state.hist = sort_month_columns(hist)
    st.session_state.themes = themes
    st.session_state.ver = st.session_state.get("ver", 0) + 1
    st.rerun()


# --- 発表履歴 ---
with tab_hist:
    st.caption("セルを直接編集できます。行の追加はテーブル下端の「＋」から。")
    c1, c2 = st.columns([1, 4])
    with c2:
        hist = clean(st.data_editor(st.session_state.hist, num_rows="dynamic", width="stretch",
                                    hide_index=True, key=f"hist_editor_{ver}"))

# --- 研究テーマ ---
with tab_theme:
    st.caption("その時点での各メンバーの研究テーマを入力。行の追加はテーブル下端の「＋」から。")
    themes = clean(st.data_editor(st.session_state.themes, num_rows="dynamic", width="stretch",
                                  hide_index=True, key=f"theme_editor_{ver}"))
    if st.button("テーマ列を追加"):
        themes[next_theme_col(themes)] = ""
        commit(hist, themes)
    new_t = delete_columns_ui(themes, "研究テーマ", "theme")
    if new_t is not None:
        commit(hist, new_t)

hist, themes = sync_members(hist, themes)

with tab_hist:
    with c1:
        today = date.today()
        y = st.number_input("年", 2000, 2100, today.year, key="add_y")
        m = st.number_input("月", 1, 12, today.month, key="add_m")
        if st.button("この月の列を追加"):
            lbl = month_label(int(y), int(m))
            if lbl in hist.columns:
                st.warning(f"{lbl} は既にあります")
            else:
                hist[lbl] = ""
                commit(hist, themes)
        new_h = delete_columns_ui(hist, "発表履歴", "hist")
        if new_h is not None:
            commit(new_h, themes)

# --- リコメンド ---
with tab_rec:
    today = date.today()
    nxt = add_months((today.year, today.month), 1)
    c1, c2, c3, c4 = st.columns(4)
    ty = c1.number_input("対象 年", 2000, 2100, nxt[0])
    tm = c2.number_input("対象 月", 1, 12, nxt[1])
    min_gap = c3.number_input("前回から最低何か月空ける？", 0, 36, 0)
    top_n = c4.number_input("表示人数", 1, 50, 5)
    target = (int(ty), int(tm))

    rec = build_recommendations(hist, themes, target, int(min_gap))
    st.subheader(f"{month_label(*target)} の発表候補")
    if rec.empty:
        st.warning("条件に合うメンバーがいません。")
    else:
        show = rec.head(int(top_n)).copy()
        show["空き(月)"] = show["空き(月)"].apply(lambda v: "初回" if pd.isna(v) else str(int(v)))
        st.dataframe(show, width="stretch", hide_index=True)
        st.caption("並び順: 対象月に未割当 → 前回発表からの空きが長い → 発表回数が少ない。"
                   "テーマ案は未発表のテーマを優先（発表履歴にテーマ名が含まれていれば発表済みと判定）。")

        st.markdown("#### 割り当てを履歴に反映")
        c1, c2, c3 = st.columns([1, 2, 1])
        who = c1.selectbox("メンバー", list(rec["メンバー"]))
        default_title = rec.loc[rec["メンバー"] == who, "テーマ案"].iloc[0]
        title = c2.text_input("発表テーマ", value=f"{default_title}について" if default_title else "",
                              key=f"title_{who}_{target}")
        c3.write("")
        if c3.button("割り当てる", type="primary"):
            lbl = month_label(*target)
            if lbl not in hist.columns:
                hist[lbl] = ""
            hist.loc[hist["メンバー"] == who, lbl] = title
            st.session_state.flash = f"{who} さんを {lbl} に割り当てました（「発表履歴」タブに反映済み）"
            commit(hist, themes)
    if st.session_state.get("flash"):
        st.success(st.session_state.pop("flash"))

# ---------- サイドバー: 保存 ----------
st.sidebar.header("② 保存（ダウンロード）")
stamp = date.today().strftime("%Y%m%d")
hist_out = sort_month_columns(hist)
st.sidebar.download_button("Excelで保存", to_excel_bytes(hist_out, themes),
                           file_name=f"研究棟議会_{stamp}.xlsx", width="stretch")
st.sidebar.download_button("発表履歴.csv", to_csv_bytes(hist_out),
                           file_name=f"発表履歴_{stamp}.csv", width="stretch")
st.sidebar.download_button("研究テーマ.csv", to_csv_bytes(themes),
                           file_name=f"研究テーマ_{stamp}.csv", width="stretch")
st.sidebar.caption("※ブラウザを閉じると編集内容は消えます。終わる前に保存してください。")
