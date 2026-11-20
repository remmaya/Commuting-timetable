"""通勤時刻表 プロトタイプ（Streamlit）

起動: streamlit run app.py

表示・スワイプ・地点タップの処理はすべてブラウザ側（timetable.html の JavaScript）で
行う。Streamlit のウィジェットを使わないので再実行（rerun）が起きず、
手動で動かした表示位置が勝手に戻ることがない。
"""
import json
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

from timetable_data import build_timetable

st.set_page_config(page_title="通勤時刻表", page_icon="🚃", layout="centered",
                   initial_sidebar_state="collapsed")

# スマホ縦画面向けに Streamlit の余白・ヘッダーを詰める
st.markdown(
    """
    <style>
      header[data-testid="stHeader"], #MainMenu, footer { display: none !important; }
      .block-container, [data-testid="stMainBlockContainer"] {
        padding: 0.4rem 0.4rem 0 0.4rem !important; max-width: 540px !important;
      }
      iframe { display: block; }
    </style>
    """,
    unsafe_allow_html=True,
)

data = build_timetable()
html = (Path(__file__).parent / "timetable.html").read_text(encoding="utf-8")
html = html.replace("/*__DATA__*/null", json.dumps(data, ensure_ascii=False))

# 高さ = 上部バー + 見出し + 時刻行×34px + 区間行×22px + 下部
n_time = sum(len(r["times"]) for r in data["rows"] if r["type"] == "stop")
n_leg = sum(1 for r in data["rows"] if r["type"] == "leg")
height = 44 + 30 + n_time * 34 + n_leg * 22 + 60

components.html(html, height=height, scrolling=False)
