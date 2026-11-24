"""区間ごとの時刻表（バス・京王など）から、乗り継ぎを自動で組み立てる（復路用）。

Excel のシート構成（return_timetable.xlsx）
  設定          学校→各バス停の徒歩、調布の乗り換え時間
  バス所要時間   乗るバス停・時間帯ごとの調布駅までの所要時間
  バス          乗るバス停・系統・発
  京王          調布発・種別・新宿の駅（新宿/京王 → 京王新宿、それ以外 → 新線新宿）・新宿着
いずれも1行目はメモ、2行目は見出し、3行目からデータ。
"""
from __future__ import annotations

from pathlib import Path

from timetable_loader import TimetableError, _col, _min, _norm_time, _read_xlsx

STOP_SHORT = {"調布中学校": "調中", "大沢四丁目": "大沢"}
STATIONS = {"京王": "京王新宿", "新線": "新線新宿"}
TYPE_SHORT = {"特急": "特", "急行": "急", "快速": "快", "区急": "区", "区間急行": "区", "各停": "各", "準特急": "準"}


def _sheet(path: Path, name: str) -> list[tuple[int, list[str]]]:
    grid = _read_xlsx(path, name)
    return [(i + 1, r) for i, r in enumerate(grid[2:], start=2) if any(c.strip() for c in r)]


def _fm(m: int) -> str:
    return f"{m // 60}:{m % 60:02d}"


def _station(s: str) -> str:
    s = s.strip()
    if s in ("新宿", "京王", "京王新宿"):
        return "京王"
    return "新線"   # 本八幡・大島・新線新宿 など都営新宿線直通


def load_segments(path: Path, colors: dict[str, str], name: str) -> dict:
    fn = path.name
    # 設定
    walk, transfer = {}, None
    for li, r in _sheet(path, "設定"):
        item, val = r[0].strip(), (r[1] if len(r) > 1 else "").strip()
        if not val:
            continue
        try:
            v = int(float(val))
        except ValueError:
            raise TimetableError(f"{fn}「設定」{li}行目: 値「{val}」が数字ではありません")
        if "乗り換え" in item:
            transfer = v
        elif "→" in item:
            stop = item.split("→")[1].replace("徒歩", "").strip()
            walk[stop] = v
    if transfer is None:
        raise TimetableError(f"{fn}「設定」: 調布の乗り換え時間がありません")

    # バス所要時間
    ride = []
    for li, r in _sheet(path, "バス所要時間"):
        r = r + [""] * 4
        where = f"{fn}「バス所要時間」{li}行目"
        ride.append((r[0].strip(), _min(_norm_time(r[1], where)), _min(_norm_time(r[2], where)), int(float(r[3]))))

    def ride_min(stop, dep):
        for s, a, b, m in ride:
            if s == stop and a <= dep <= b:
                return m
        raise TimetableError(f"{fn}「バス所要時間」: {stop} {_fm(dep)}発の所要時間がありません")

    # バス
    buses = []
    for li, r in _sheet(path, "バス"):
        r = r + [""] * 3
        stop, route, dep = r[0].strip(), r[1].strip(), _norm_time(r[2], f"{fn}「バス」{li}行目 C列")
        if not dep:
            continue
        if stop not in walk:
            raise TimetableError(f"{fn}「バス」{li}行目: 「{stop}」までの徒歩時間が「設定」にありません")
        d = _min(dep)
        buses.append({"stop": stop, "route": route, "dep": d, "arr": d + ride_min(stop, d)})

    # 京王
    trains = []
    for li, r in _sheet(path, "京王"):
        r = r + [""] * 4
        where = f"{fn}「京王」{li}行目"
        dep = _norm_time(r[0], where + " A列")
        arr = _norm_time(r[3], where + " D列")
        if not dep or not arr:      # 新宿着が空欄の列車は候補にしない
            continue
        if _min(arr) <= _min(dep):
            raise TimetableError(f"{where}: 新宿着 {arr} が調布発 {dep} より前です")
        trains.append({"dep": _min(dep), "type": r[1].strip(), "st": _station(r[2]), "arr": _min(arr)})

    rows = [
        {"type": "stop", "id": "s0", "name": "学校", "times": ["dep"]},
        {"type": "leg", "id": "l1"},
        {"type": "stop", "id": "s2", "name": "バス停", "times": ["dep"]},
        {"type": "leg", "id": "l3"},
        {"type": "stop", "id": "s4", "name": "調布駅", "times": ["arr", "dep"]},
        {"type": "leg", "id": "l5"},
        {"type": "stop", "id": "s6", "name": "新宿駅", "times": ["arr"]},
    ]
    gray = "#80868b"
    trips = []
    for b in buses:
        ok = [t for t in trains if t["dep"] >= b["arr"] + transfer]
        best = {}
        for st in STATIONS:
            c = [t for t in ok if t["st"] == st]
            if c:   # 新宿に一番早く着く列車。同着なら調布での待ちが短い方
                best[st] = min(c, key=lambda t: (t["arr"], -t["dep"]))
        if not best:
            continue
        overall = min(best.values(), key=lambda t: (t["arr"], -t["dep"]))
        for st, t in best.items():
            short = TYPE_SHORT.get(t["type"], t["type"][:1])
            trips.append({
                "times": {"s0.dep": _fm(b["dep"] - walk[b["stop"]]), "s2.dep": _fm(b["dep"]),
                          "s4.arr": _fm(b["arr"]), "s4.dep": _fm(t["dep"]), "s6.arr": _fm(t["arr"])},
                "legs": {"l1": {"label": STOP_SHORT.get(b["stop"], b["stop"][:2]), "color": gray},
                         "l3": {"label": b["route"], "color": colors.get(b["route"], gray)},
                         "l5": {"label": short, "color": colors.get(short, gray)}},
                "marks": {"s6.arr": "京" if st == "京王" else "新"},
                "station": st,
                "best": t is overall,
                "src": f"{fn} バス {b['stop']} {_fm(b['dep'])}",
            })
    trips.sort(key=lambda t: (_min(t["times"]["s0.dep"]), _min(t["times"]["s6.arr"])))
    for i, t in enumerate(trips):
        t["no"] = i + 1
    return {"name": name, "rows": rows, "trips": trips, "goals": [],
            "stations": [{"key": "", "name": "どちらでも"}] + [{"key": k, "name": v} for k, v in STATIONS.items()]}
