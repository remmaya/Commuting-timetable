"""往路：区間ごとの時刻表（JR・京王・バス）から乗り継ぎを自動で組み立てる。

Excel のシート構成（outbound_timetable.xlsx）。いずれも1行目はメモ、2行目は見出し、3行目からデータ。
  設定          自宅→赤羽駅、新宿の乗り換え、調布の乗り換え、各バス停→学校の徒歩
  バス所要時間   降りるバス停・調布駅を出る時間帯ごとの所要時間
  系統          系統ごとの降りるバス停
  JR            赤羽発・新宿着
  京王          新宿発・種別・調布着
  バス          系統・調布発

組み立て方
  JR 1本ごとに、学校に一番早く着く行き方を探す。同じバスに乗れる京王が複数あれば早い方
  （遅い方はあとのJRのスジとして隣に出る）。降りるバス停が違う行き方が同着なら両方出す。
  あとから出る列車に抜かれる列車は候補にしない。ただし新宿で最初に来る列車がそれで、
  乗ると学校に遅れて着くなら ⚠ 付きで出す（うっかり乗ってしまう失敗を見えるようにするため）。
"""
from __future__ import annotations

from pathlib import Path

from segment_loader import TYPE_SHORT, _fm, _sheet
from timetable_loader import TimetableError, _compute_joins, _min, _norm_time, find_link, load_links

TYPE_FULL = {"特": "特急", "急": "急行", "区": "区間急行", "快": "快速", "各": "各停"}


def _int(v, where):
    try:
        return int(float(v))
    except ValueError:
        raise TimetableError(f"{where}: 値「{v}」が数字ではありません")


def load_outbound(path: Path, colors: dict[str, str], name: str) -> dict:
    fn = path.name
    links = load_links()
    gray = "#80868b"

    # 設定
    home = sh_tr = ch_tr = None
    walk = {}
    for li, r in _sheet(path, "設定"):
        item, val = r[0].strip(), (r[1] if len(r) > 1 else "").strip()
        if not val:
            continue
        v = _int(val, f"{fn}「設定」{li}行目")
        if "自宅" in item:
            home = v
        elif "新宿" in item and "乗り換え" in item:
            sh_tr = v
        elif "調布" in item and "乗り換え" in item:
            ch_tr = v
        elif "→学校" in item:
            walk[item.split("→")[0].strip()] = v
    for label, v in (("自宅→赤羽駅", home), ("新宿の乗り換え", sh_tr), ("調布の乗り換え", ch_tr)):
        if v is None:
            raise TimetableError(f"{fn}「設定」: {label}の値がありません")

    ride = []
    for li, r in _sheet(path, "バス所要時間"):
        r = r + [""] * 4
        w = f"{fn}「バス所要時間」{li}行目"
        ride.append((r[0].strip(), _min(_norm_time(r[1], w)), _min(_norm_time(r[2], w)), _int(r[3], w)))

    def ride_min(stop, dep):
        for s, a, b, m in ride:
            if s == stop and a <= dep <= b:
                return m
        raise TimetableError(f"{fn}「バス所要時間」: {stop} 行き {_fm(dep)}発の所要時間がありません")

    route_stop = {}
    for li, r in _sheet(path, "系統"):
        r = r + [""] * 2
        if r[0].strip():
            route_stop[r[0].strip()] = r[1].strip()

    jr = []
    for li, r in _sheet(path, "JR"):
        r = r + [""] * 2
        w = f"{fn}「JR」{li}行目"
        d, a = _norm_time(r[0], w + " A列"), _norm_time(r[1], w + " B列")
        if d and a:
            jr.append({"dep": _min(d), "arr": _min(a)})

    keio = []
    for li, r in _sheet(path, "京王"):
        r = r + [""] * 3
        w = f"{fn}「京王」{li}行目"
        d, a = _norm_time(r[0], w + " A列"), _norm_time(r[2], w + " C列")
        if d and a:
            keio.append({"dep": _min(d), "type": r[1].strip(), "arr": _min(a)})

    buses = []
    for li, r in _sheet(path, "バス"):
        r = r + [""] * 2
        w = f"{fn}「バス」{li}行目"
        route, d = r[0].strip(), _norm_time(r[1], w + " B列")
        if not d:
            continue
        stop = route_stop.get(route)
        if not stop:
            raise TimetableError(f"{w}: 系統「{route}」の降りるバス停が「系統」シートにありません")
        if stop not in walk:
            raise TimetableError(f"{w}: 「{stop}→学校」の徒歩時間が「設定」にありません")
        dep = _min(d)
        arr = dep + ride_min(stop, dep)
        buses.append({"route": route, "stop": stop, "dep": dep, "arr": arr, "school": arr + walk[stop]})

    # 抜かれる列車：あとから新宿を出る列車が、調布に先に（または同時に）着く
    for t in keio:
        faster = [u for u in keio if u["dep"] > t["dep"] and u["arr"] <= t["arr"]]
        t["overtaken_by"] = min(faster, key=lambda u: u["dep"]) if faster else None

    def best_buses(k):
        ok = [b for b in buses if b["dep"] >= k["arr"] + ch_tr]
        if not ok:
            return []
        best = min(b["school"] for b in ok)
        return [b for b in ok if b["school"] == best]

    rows = [
        {"type": "stop", "id": "s0", "name": "自宅", "times": ["dep"]},
        {"type": "leg", "id": "l1"},
        {"type": "stop", "id": "s2", "name": "赤羽駅", "times": ["dep"]},
        {"type": "leg", "id": "l3"},
        {"type": "stop", "id": "s4", "name": "新宿駅", "times": ["arr", "dep"]},
        {"type": "leg", "id": "l5"},
        {"type": "stop", "id": "s6", "name": "調布駅", "times": ["arr", "dep"]},
        {"type": "leg", "id": "l7"},
        {"type": "stop", "id": "s8", "name": "バス停", "times": ["arr"]},
        {"type": "leg", "id": "l9"},
        {"type": "stop", "id": "s10", "name": "学校", "times": ["arr"]},
    ]

    def make(j, k, b, warn=None):
        short = TYPE_SHORT.get(k["type"], k["type"][:1])
        trip = {
            "times": {"s0.dep": _fm(j["dep"] - home), "s2.dep": _fm(j["dep"]), "s4.arr": _fm(j["arr"]),
                      "s4.dep": _fm(k["dep"]), "s6.arr": _fm(k["arr"]),
                      "s6.dep": _fm(b["dep"]) if b else None, "s8.arr": _fm(b["arr"]) if b else None,
                      "s10.arr": _fm(b["school"]) if b else None},
            "legs": {"l5": {"label": short, "color": colors.get(short, gray)},
                     "l7": {"label": b["route"], "color": colors.get(b["route"], gray),
                            "url": find_link(links, "調布駅", b["route"])} if b else None},
            "jr": j["dep"],
            "src": f"{fn} JR {_fm(j['dep'])}",
        }
        if warn:
            u = warn
            trip["legs"]["l5"]["warn"] = True
            trip["warn"] = f"⚠ {_fm(u['dep'])}発の{TYPE_FULL.get(TYPE_SHORT.get(u['type'], ''), u['type'])}が先に調布に着きます"
        return trip

    trips = []
    for j in jr:
        ok = [k for k in keio if k["dep"] >= j["arr"] + sh_tr]
        if not ok:
            continue
        # 抜かれる列車は候補にしない（同じ時刻に着くとしても、あとの速い列車に乗ればよい）
        opts = [(k, b) for k in ok if not k["overtaken_by"] for b in best_buses(k)]
        if opts:
            best = min(b["school"] for k, b in opts)
            chosen = [(k, b) for k, b in opts if b["school"] == best]
            # 同じバスに乗れる京王が複数あるときは、早い方（調布で余裕を持てる方）だけ。
            # 遅い方の列車は、あとのJRのスジに乗る列車として隣に出てくる
            by_bus = {}
            for k, b in chosen:
                by_bus.setdefault(id(b), []).append((k, b))
            chosen = [min(group, key=lambda kb: kb[0]["dep"]) for group in by_bus.values()]
        else:   # 学校まで行けるバスがない時間帯は、一番早く着く列車まで
            k = min(ok, key=lambda k: k["arr"])
            chosen = [(k, None)]
        for k, b in chosen:
            trips.append(make(j, k, b))
        # 新宿で最初に来る列車が抜かれる列車で、乗ると学校に遅れて着くなら ⚠ 付きで出す
        first = min(ok, key=lambda k: k["dep"])
        if first["overtaken_by"]:
            bs = best_buses(first)
            late = not bs or (opts and bs[0]["school"] > best)
            if late:
                trips.append(make(j, first, bs[0] if bs else None, warn=first["overtaken_by"]))

    trips.sort(key=lambda t: (_min(t["times"]["s0.dep"]), _min(t["times"]["s4.dep"]),
                              _min(t["times"]["s6.dep"]) if t["times"]["s6.dep"] else 9999))
    _compute_joins(rows, trips)
    for t in trips:   # 同じJRに乗るスジは、自宅から新宿着まで合流
        for key in ("s0.dep", "l1", "s2.dep", "l3", "s4.arr"):
            t["joins"][key] = f"jr{t['jr']}"
    for i, t in enumerate(trips):
        t["no"] = i + 1
    return {"name": name, "rows": rows, "trips": trips, "goals": [], "auto": True}
