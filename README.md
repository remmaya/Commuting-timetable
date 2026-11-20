# 通勤時刻表 プロトタイプ

## 起動
```
pip install -r requirements.txt
streamlit run app.py --server.address 0.0.0.0
```
同じWi-Fiに繋いだAndroidのChromeで `http://<PCのIPアドレス>:8501` を開く。

## ファイル
- `app.py` … Streamlit本体（HTMLにデータを埋め込んで表示するだけ）
- `timetable.html` … 画面・スワイプ・地点タップの処理（HTML/CSS/JS）
- `timetable_data.py` … 仮の時刻表データ（平日 1限）

## 操作
- 起動時：日本時間の現在時刻以降で最初の「自宅発」のスジを中央に表示
- 地点名タップ：その地点の「発」（なければ「着」）が現在時刻以降の最初のスジを中央へ
- 左右スワイプ／◀▶：スジを1本ずつ移動（時間経過で勝手に動くことはない）
- テスト時刻：入力すると、次に地点をタップしたときの「現在時刻」として使われる
