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

from timetable_loader import TimetableError, _col, _min, _norm_time, _read_xlsx, find_link, load_links

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
    links = load_links()
    # 設定
    walk, transfer, sz_transfer_cfg = {}, None, None
    for li, r in _sheet(path, "設定"):
        item, val = r[0].strip(), (r[1] if len(r) > 1 else "").strip()
        if not val:
            continue
        try:
            v = int(float(val))
        except ValueError:
            raise TimetableError(f"{fn}「設定」{li}行目: 値「{val}」が数字ではありません")
        if "笹塚" in item and "乗り換え" in item:
            sz_transfer_cfg = v
        elif "乗り換え" in item:
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
        sz = _norm_time(r[4], where + " E列") if len(r) > 4 else None
        trains.append({"dep": _min(dep), "type": r[1].strip(), "st": _station(r[2]), "arr": _min(arr),
                       "sz": _min(sz) if sz else None})

    # 笹塚（新線新宿方面）。シートがなければ笹塚乗換は作らない
    sz_trains, sz_transfer = [], 1
    import openpyxl
    if "笹塚" in openpyxl.load_workbook(path, read_only=True).sheetnames:
        for li, r in _sheet(path, "笹塚"):
            r = r + [""] * 3
            where = f"{fn}「笹塚」{li}行目"
            d, a = _norm_time(r[0], where + " A列"), _norm_time(r[1], where + " B列")
            if d and a:
                sz_trains.append({"dep": _min(d), "arr": _min(a), "type": r[2].strip()})

    if sz_transfer_cfg is not None:
        sz_transfer = sz_transfer_cfg
    rows = [
        {"type": "stop", "id": "s0", "name": "学校", "times": ["dep"]},
        {"type": "leg", "id": "l1"},
        {"type": "stop", "id": "s2", "name": "バス停", "times": ["dep"]},
        {"type": "leg", "id": "l3"},
        {"type": "stop", "id": "s4", "name": "調布駅", "times": ["arr", "dep"]},
        {"type": "leg", "id": "l5"},
        {"type": "stop", "id": "s6", "name": "新宿駅", "times": ["arr"]},
    ]
    rows_sz = rows[:5] + [
        {"type": "leg", "id": "l5"},
        {"type": "stop", "id": "sz", "name": "笹塚", "times": ["arr", "dep"]},
        {"type": "leg", "id": "lz"},
        rows[6],
    ]
    gray = "#80868b"
    key = lambda t: (t["arr"], -t["dep"])   # 新宿に一番早く着く。同着なら調布での待ちが短い方

    def make(b, t, st, via=None):
        short = TYPE_SHORT.get(t["type"], t["type"][:1])
        trip = {
            "times": {"s0.dep": _fm(b["dep"] - walk[b["stop"]]), "s2.dep": _fm(b["dep"]),
                      "s4.arr": _fm(b["arr"]), "s4.dep": _fm(t["dep"])},
            "legs": {"l1": {"label": STOP_SHORT.get(b["stop"], b["stop"][:2]), "color": gray},
                     "l3": {"label": b["route"], "color": colors.get(b["route"], gray),
                            "url": find_link(links, b["stop"], b["route"])},
                     "l5": {"label": short, "color": colors.get(short, gray)}},
            "marks": {"s6.arr": "京" if st == "京王" else "新"},
            "src": f"{fn} バス {b['stop']} {_fm(b['dep'])}",
            "bus": f"{b['stop']}|{b['route']}|{b['dep']}",
        }
        if via:   # 笹塚で新線に乗り換え
            trip["times"].update({"sz.arr": _fm(t["sz"]), "sz.dep": _fm(via["dep"]), "s6.arr": _fm(via["arr"])})
            vs = TYPE_SHORT.get(via["type"], via["type"][:1]) if via["type"] else None
            trip["legs"]["lz"] = {"label": vs, "color": colors.get(vs, gray)} if vs else None
            trip["prefer"] = True   # 地点をタップしたとき、同じバスの中ではこのスジ（先着）を中央に
        else:
            trip["times"]["s6.arr"] = _fm(t["arr"])
        return trip

    views = {"": [], "京王": [], "新線": []}
    for b in buses:
        ok = [t for t in trains if t["dep"] >= b["arr"] + transfer]
        best = {st: min((t for t in ok if t["st"] == st), key=key, default=None) for st in STATIONS}
        cands = [t for t in best.values() if t]
        if not cands:
            continue
        overall = min(cands, key=key)
        views[""].append(make(b, overall, overall["st"]))
        if best["京王"]:
            views["京王"].append(make(b, best["京王"], "京王"))
        if best["新線"]:
            views["新線"].append(make(b, best["新線"], "新線"))
        # 特急で笹塚へ行き、新線に乗り換える（直通より早く着くときだけ）
        if sz_trains:
            opts = []
            for t in ok:
                if t["sz"] is None:
                    continue
                v = min((x for x in sz_trains if x["dep"] >= t["sz"] + sz_transfer), key=key, default=None)
                if v:
                    opts.append((v["arr"], -t["dep"], t, v))
            if opts:
                arr, _, t, v = min(opts, key=lambda o: o[:2])
                if not best["新線"] or arr < best["新線"]["arr"]:
                    views["新線"].append(make(b, t, "新線", via=v))

    # 列車基準のスジ（調布駅・笹塚をタップしたとき用）：バスに関係なく、調布を出る列車を全部並べる
    def make_train(t, st, via=None):
        short = TYPE_SHORT.get(t["type"], t["type"][:1])
        trip = {"times": {"s4.dep": _fm(t["dep"])},
                "legs": {"l5": {"label": short, "color": colors.get(short, gray)}},
                "marks": {"s6.arr": "京" if st == "京王" else "新"},
                "src": f"{fn} 京王 {_fm(t['dep'])}", "train": True}
        if via:
            vs = TYPE_SHORT.get(via["type"], via["type"][:1]) if via["type"] else None
            trip["times"].update({"sz.arr": _fm(t["sz"]), "sz.dep": _fm(via["dep"]), "s6.arr": _fm(via["arr"])})
            trip["legs"]["lz"] = {"label": vs, "color": colors.get(vs, gray)} if vs else None
        else:
            trip["times"]["s6.arr"] = _fm(t["arr"])
        return trip

    tviews = {"": [], "京王": [], "新線": []}
    for t in trains:
        tviews[""].append(make_train(t, t["st"]))
        tviews[t["st"]].append(make_train(t, t["st"]))
        if sz_trains and t["sz"] is not None:   # 特急 → 笹塚で新線に乗り換え
            v = min((x for x in sz_trains if x["dep"] >= t["sz"] + sz_transfer), key=key, default=None)
            if v:
                tviews["新線"].append(make_train(t, "新線", via=v))

    from timetable_loader import _compute_joins
    tout = {}
    for k, ts in tviews.items():
        ts.sort(key=lambda t: (_min(t["times"]["s4.dep"]), _min(t["times"]["s6.arr"])))
        r = rows_sz if (k == "新線" and sz_trains) else rows
        _compute_joins(r, ts)
        tout[k] = {"rows": r, "trips": ts}

    # 笹塚をタップしたとき用：笹塚から新線新宿へ向かう列車を全部
    sz_view = None
    if sz_trains:
        sz_view = [{"times": {"sz.dep": _fm(x["dep"]), "s6.arr": _fm(x["arr"])},
                    "legs": {"lz": {"label": TYPE_SHORT.get(x["type"], x["type"][:1]),
                                    "color": colors.get(TYPE_SHORT.get(x["type"], x["type"][:1]), gray)} if x["type"] else None},
                    "marks": {"s6.arr": "新"}, "train": True, "mode": "笹塚からの列車", "src": f"{fn} 笹塚"}
                   for x in sorted(sz_trains, key=lambda x: x["dep"])]
        _compute_joins(rows_sz, sz_view)

    out = {}
    for k, ts in views.items():
        ts.sort(key=lambda t: (_min(t["times"]["s0.dep"]), _min(t["times"]["s6.arr"])))
        r = rows_sz if (k == "新線" and sz_trains) else rows
        _compute_joins(r, ts)
        for t in ts:   # 同じバスに乗るスジは、学校を出るところからバスを降りるまで合流
            for k2 in ("s0.dep", "l1", "s2.dep", "l3", "s4.arr"):
                t["joins"][k2] = t["bus"]
        for i, t in enumerate(ts):
            t["no"] = i + 1
        out[k] = {"rows": r, "trips": ts}
    return {"name": name, "rows": rows, "trips": out[""]["trips"], "views": out, "trainViews": tout, "szView": {"trips": sz_view} if sz_view else None, "goals": [],
            "busStops": [w for w in walk if any(b["stop"] == w for b in buses)],   # 設定シートの順
            "stations": [{"key": "", "name": "どちらでも"}] + [{"key": k, "name": v} for k, v in STATIONS.items()]}
