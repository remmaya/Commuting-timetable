"""仮の時刻表データ（平日 1限）。

Excelの表をそのまま写したプロトタイプ用データ。
将来はExcel読み込みに置き換える想定なので、build_timetable() が返す
辞書の形（rows / trips）だけを画面側との約束事にしている。
"""

# ---- 地点（縦方向）と区間の定義 ------------------------------------------
# type="stop": 地点。times に "arr"(着) / "dep"(発) を持つ
# type="leg" : 地点間の移動。kind で表示の仕方を変える
ROWS = [
    {"type": "stop", "id": "home",     "name": "自宅",   "times": ["dep"]},
    {"type": "leg",  "id": "walk_home", "kind": "plain"},
    {"type": "stop", "id": "akabane",  "name": "赤羽駅", "times": ["arr", "dep"]},
    {"type": "leg",  "id": "jr",        "kind": "plain"},
    {"type": "stop", "id": "shinjuku", "name": "新宿駅", "times": ["arr", "dep"]},
    {"type": "leg",  "id": "keio",      "kind": "train"},
    {"type": "stop", "id": "chofu",    "name": "調布駅", "times": ["arr", "dep"]},
    {"type": "leg",  "id": "bus",       "kind": "bus"},
    {"type": "stop", "id": "busstop",  "name": "バス停", "times": ["arr"]},
    {"type": "leg",  "id": "walk_school", "kind": "plain"},
    {"type": "stop", "id": "school",   "name": "学校",   "times": ["arr"]},
]

# 列車種別・バス系統の表示色
TRAIN_TYPES = {"特": "#d93025", "急": "#188038", "区": "#0b8a7a"}
BUS_ROUTES = {"K": "#7b1fa2", "O": "#d93025", "H": "#e37400", "M": "#188038"}

# ---- スジ（横方向）---------------------------------------------------------
HOME_DEP    = "6:24 6:32 6:36 6:40 6:45 6:49 6:52 6:55 7:00 7:05 7:08 7:12 7:16 7:20 7:24 7:28 7:31".split()
AKABANE_ARR = "6:39 6:47 6:51 6:55 7:00 7:04 7:07 7:10 7:15 7:20 7:23 7:27 7:31 7:35 7:39 7:43 7:46".split()
AKABANE_DEP = "6:41 6:49 6:53 6:57 7:02 7:06 7:09 7:12 7:17 7:22 7:25 7:29 7:33 7:37 7:41 7:45 7:48".split()
SHINJUKU_ARR = "6:54 7:03 7:08 7:12 7:18 7:21 7:24 7:27 7:33 7:37 7:40 7:45 7:48 7:52 7:57 8:00 8:02".split()

# 新宿→調布の列車。(新宿発, 種別, 調布着, 何本のスジがこの列車に合流するか)
KEIO_TRAINS = [
    ("7:03", "特", "7:23", 1),
    ("7:15", "特", "7:36", 2),
    ("7:21", "急", "7:44", 2),
    ("7:29", "区", "7:53", 2),
    ("7:37", "特", "7:55", 1),
    ("7:42", "区", "8:11", 1),
    ("7:47", "特", "8:03", 2),
    ("7:53", "特", "8:12", 2),
    ("8:01", "特", "8:23", 2),
    ("8:10", "特", "8:33", 1),
    ("8:16", "急", "8:39", 1),
]

# 調布→バス停→学校。(調布発, 系統, バス停着, 学校着) / 接続なしは None
BUSES = [
    ("7:39", "K", "7:54", "8:01"),
    ("7:40", "O", "7:55", "8:02"),
    ("7:41", "H", "7:53", "7:58"),
    ("7:50", "O", "8:05", "8:12"),
    ("7:56", "M", "8:08", "8:13"),
    ("7:57", "K", "8:12", "8:19"),
    None,
    ("8:01", "O", "8:16", "8:23"),
    None,
    ("8:12", "O", "8:27", "8:34"),
    ("8:18", "K", "8:33", "8:40"),
    ("8:21", "H", "8:33", "8:38"),
    ("8:23", "O", "8:38", "8:45"),
    ("8:34", "M", "8:46", "8:51"),
    ("8:35", "O", "8:50", "8:57"),
    ("8:36", "K", "8:51", "8:58"),
    None,
]


def build_timetable() -> dict:
    keio_per_trip = []
    for dep, ttype, arr, n in KEIO_TRAINS:
        keio_per_trip += [{"dep": dep, "type": ttype, "arr": arr, "train": f"keio-{dep}"}] * n
    assert len(keio_per_trip) == len(HOME_DEP) == len(BUSES)

    trips = []
    for i in range(len(HOME_DEP)):
        k = keio_per_trip[i]
        bus = BUSES[i]
        trips.append({
            "no": i + 1,
            "times": {
                "home.dep": HOME_DEP[i],
                "akabane.arr": AKABANE_ARR[i],
                "akabane.dep": AKABANE_DEP[i],
                "shinjuku.arr": SHINJUKU_ARR[i],
                "shinjuku.dep": k["dep"],
                "chofu.arr": k["arr"],
                "chofu.dep": bus[0] if bus else None,
                "busstop.arr": bus[2] if bus else None,
                "school.arr": bus[3] if bus else None,
            },
            "legs": {
                "keio": {"label": k["type"], "color": TRAIN_TYPES[k["type"]]},
                "bus": {"label": bus[1], "color": BUS_ROUTES[bus[1]]} if bus else None,
            },
            # 同じ値を持つ隣のスジと「合流」表示にする行と、その識別子
            "joins": {
                "shinjuku.dep": k["train"],
                "keio": k["train"],
                "chofu.arr": k["train"],
            },
        })

    return {"title": "平日 1限", "rows": ROWS, "trips": trips}
