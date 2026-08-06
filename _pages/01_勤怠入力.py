"""
01_勤怠入力.py

【概要】
一般ユーザー用の勤怠入力・閲覧画面（SC-02）。
実際の入力・プレビュー機能はDay4・Day5で実装する。
このファイル自体はDay1時点では「画面に到達できること」と
「権限チェックが機能すること」の確認が目的。

session_service.require_general_user() を先頭で呼ぶことで、
- 未ログインならログイン画面へ誘導（st.stop）
- adminがこの画面に迷い込んだ場合も案内して停止（st.stop）
という二重の安全策を、画面ごとに個別実装しなくても統一的に効かせている。
"""

import streamlit as st

from services import session_service

user = session_service.require_general_user()

st.title("勤怠入力・閲覧")
st.write(f"{user.display_name} さんの勤怠を入力・閲覧します。")
st.info("入力・プレビュー機能は Day4・Day5 で実装予定です。")