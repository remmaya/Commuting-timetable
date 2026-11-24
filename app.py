"""通勤時刻表（Streamlit）

起動: streamlit run app.py

表示・スワイプ・地点タップ・残り時間の処理はすべてブラウザ側（timetable.html）で行う。
Streamlit のウィジェットを使わないので再実行（rerun）が起きず、手動で動かした表示位置が
勝手に戻ることがない。時刻表データは data/ フォルダの CSV（README.md 参照）。
"""
import json
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

from timetable_loader import TimetableError, load_all

st.set_page_config(page_title="通勤時刻表", page_icon="🚃", layout="centered",
                   initial_sidebar_state="collapsed")

st.markdown(
    """
    <style>
      header[data-testid="stHeader"], #MainMenu, footer { display: none !important; }
      .block-container, [data-testid="stMainBlockContainer"] {
        padding: 0.5rem 0.4rem 0 0.4rem !important; max-width: 540px !important;
      }
      iframe { display: block; }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data
def load(mtimes):  # データファイルを書き換えたら（更新日時が変われば）読み直す
    return load_all()


data_dir = Path(__file__).parent / "data"
mtimes = tuple(sorted((p.name, p.stat().st_mtime) for p in data_dir.iterdir()
                      if p.suffix.lower() in (".csv", ".xlsx", ".xlsm")))
try:
    data = load(mtimes)
except TimetableError as e:
    st.error(f"時刻表データの読み込みに失敗しました\n\n{e}")
    st.stop()

if data["warnings"]:
    with st.expander(f"⚠️ 時刻表データの確認（{len(data['warnings'])}件）"):
        for w in data["warnings"]:
            st.write("・" + w)

html = (Path(__file__).parent / "timetable.html").read_text(encoding="utf-8")
html = html.replace("/*__DATA__*/null", json.dumps(data, ensure_ascii=False))


# 高さ（timetable.html の CSS 変数と合わせる）
def table_height(d):
    h = 36  # 見出し行
    for r in d["rows"]:
        if r["type"] == "leg":
            h += 30
        elif len(r["times"]) == 1:
            h += 48
        else:
            h += 32 + 44
    return h


height = 52 + 46 + 80 + max(table_height(d) for d in data["directions"]) + 2 + 50 + 64
components.html(html, height=height, scrolling=False)
