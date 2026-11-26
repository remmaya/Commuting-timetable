"""時刻表の読み込み。

data/timetables.csv … 方向・目標・ファイル・シート・締切。同じ方向のシートは1本の時刻表にまとめる
data/colors.csv     … 種別・系統の表示色
data/*.xlsx / *.csv  … 時刻表本体（縦：地点、横：スジ）。Excelはシートをそのまま読む

CSV の書き方は README.md を参照。
"""
from __future__ import annotations

import csv
import io
import re
from pathlib import Path

DATA_DIR = Path(__file__).parent / "data"
KINDS = {"発": "dep", "着": "arr"}
NONE_MARKS = {"", "-", "－", "—", "ー", "×"}
SAME_AS_LEFT = {"〃", "=", "＝"}
DEFAULT_COLOR = "#5f6368"


class TimetableError(Exception):
    pass


def _read_csv(path: Path) -> list[list[str]]:
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "cp932"):   # Excelの「CSV UTF-8」と通常のCSV(Shift_JIS)の両方に対応
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise TimetableError(f"{path.name}: 文字コードを判別できません")
    rows = []
    for row in csv.reader(io.StringIO(text)):
        cells = [c.strip() for c in row]
        # コメント行は空行として残す（エラー表示の行番号をファイルと一致させるため）
        rows.append([] if cells and cells[0].startswith("#") else cells)
    return rows


def _col(i: int) -> str:
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _read_xlsx(path: Path, sheet: str) -> list[list[str]]:
    import datetime as dt
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True)       # 数式は計算結果の値を読む
    if sheet not in wb.sheetnames:
        raise TimetableError(f"{path.name}: シート「{sheet}」がありません（あるのは {', '.join(wb.sheetnames)}）")
    ws = wb[sheet]
    formulas = openpyxl.load_workbook(path)[sheet]

    def fmt(v, coord):
        if v is None:
            f = formulas[coord].value
            if isinstance(f, str) and f.startswith("="):
                # 計算結果が保存されていない（Excel以外で保存した等）
                raise TimetableError(f"{path.name}「{sheet}」{coord}: 数式の計算結果がありません。"
                                     "Excelで開いて保存し直してください")
            return ""
        if isinstance(v, dt.datetime):
            v = v.time()
        if isinstance(v, dt.time):
            secs = v.hour * 3600 + v.minute * 60 + v.second + v.microsecond / 1e6
        elif isinstance(v, (int, float)) and 0 <= v < 2:   # 日付の小数（1日=1.0）
            secs = v * 86400
        else:
            return str(v).strip()
        m = int(secs + 30) // 60                             # 数式の誤差を分単位に丸める
        return f"{m // 60}:{m % 60:02d}"

    grid = [[fmt(c.value, c.coordinate) for c in row] for row in ws.iter_rows()]
    # 結合セルは左上の値を範囲全体に広げる（複数のスジが同じ列車・バスに乗る区間）
    for rng in ws.merged_cells.ranges:
        v = grid[rng.min_row - 1][rng.min_col - 1]
        for r in range(rng.min_row - 1, rng.max_row):
            for c in range(rng.min_col - 1, rng.max_col):
                grid[r][c] = v
    return grid


def _norm_time(s: str, where: str) -> str | None:
    s = s.strip().replace("：", ":")
    if s in NONE_MARKS:
        return None
    m = re.fullmatch(r"(\d{1,2}):(\d{2})(?::\d{2})?", s)
    if not m:
        raise TimetableError(f"{where}: 時刻「{s}」を読めません（例 7:05）")
    h, mi = int(m[1]), int(m[2])
    if h > 29 or mi > 59:
        raise TimetableError(f"{where}: 時刻「{s}」が範囲外です")
    return f"{h}:{mi:02d}"


def _norm_label(s: str) -> str | None:
    s = re.sub(r"[↓→⇒\s]", "", s)   # Excelからコピーした「↓特」もそのまま使える
    return None if s in NONE_MARKS else s


def load_colors() -> dict[str, str]:
    path = DATA_DIR / "colors.csv"
    colors = {}
    if path.exists():
        for row in _read_csv(path)[1:]:
            if len(row) >= 2 and row[0]:
                colors[row[0]] = row[1]
    return colors


def load_links() -> list[tuple[str, str, str]]:
    """data/links.csv：乗るバス停・系統ごとの接近情報ページ（系統の欄を空にすると、そのバス停の全系統）。"""
    path = DATA_DIR / "links.csv"
    out = []
    if path.exists():
        for row in _read_csv(path)[1:]:
            row = row + ["", "", ""]
            if row[2].startswith("http"):
                out.append((row[0], row[1], row[2]))
    return out


def find_link(links, stop: str, route: str) -> str | None:
    for s, r, u in links:
        if s == stop and r == route:
            return u
    for s, r, u in links:
        if s == stop and not r:
            return u
    for s, r, u in links:
        if not s and r == route:
            return u
    return None


def load_timetable(path: Path, colors: dict[str, str], sheet: str = "") -> dict:
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        table = _read_xlsx(path, sheet)
        path = Path(f"{path.name}「{sheet}」")
    else:
        table = _read_csv(path)
    filled = [i for i, r in enumerate(table) if any(r)]
    if len(filled) < 2:
        raise TimetableError(f"{path.name}: データがありません")
    first = filled[0] + 1                # 最初の行は見出し（スジ番号など）
    body = table[first:]
    n = max(len(r) for r in body) - 2
    # 末尾の空列を除く
    while n > 0 and all(len(r) <= n + 1 or r[n + 1] in NONE_MARKS for r in body):
        n -= 1
    if n <= 0:
        raise TimetableError(f"{path.name}: スジ（3列目以降）がありません")

    def cells(row, line):
        vals = (row[2:] + [""] * n)[:n]
        out = []
        for j, v in enumerate(vals):
            if v in SAME_AS_LEFT:
                if j == 0:
                    raise TimetableError(f"{path.name} {line}行目: 最初の列に「〃」は使えません")
                v = out[-1] or ""
            out.append(v)
        return out

    rows, trips = [], [{"no": j + 1, "times": {}, "legs": {}, "src": f"{path.name} {_col(j + 2)}列"}
                       for j in range(n)]
    stop = None
    for li, row in enumerate(body, start=first + 1):
        row = row + ["", ""]
        place, kind = row[0], row[1]
        if not any(row):
            continue
        if kind in KINDS:
            k = KINDS[kind]
            # 地点名が空欄なら直前の地点の続き（Excelの結合セル対策）
            continuing = stop is not None and rows[-1] is stop and (not place or place == stop["name"])
            if not continuing:
                if not place:
                    raise TimetableError(f"{path.name} {li}行目: 地点名が空欄です")
                if rows and rows[-1]["type"] == "stop":   # 地点が続くときは間に矢印行を補う
                    rows.append({"type": "leg", "id": f"l{len(rows)}"})
                stop = {"type": "stop", "id": f"s{len(rows)}", "name": place, "times": []}
                rows.append(stop)
            if k in stop["times"]:
                raise TimetableError(f"{path.name} {li}行目: {stop['name']}の「{kind}」が重複しています")
            stop["times"].append(k)
            for j, v in enumerate(cells(row, li)):
                trips[j]["times"][f"{stop['id']}.{k}"] = _norm_time(v, f"{path.name} {li}行目 {_col(j + 2)}列")
        else:
            # 発/着以外の行 = 区間（種別・系統・矢印）
            if not rows:
                raise TimetableError(f"{path.name} {li}行目: 最初の行は地点（発/着）にしてください")
            board = rows[-1]["name"] if rows[-1]["type"] == "stop" else ""   # この区間で乗る地点
            leg = {"type": "leg", "id": f"l{len(rows)}"}
            rows.append(leg)
            links = load_links()
            stop = None
            for j, v in enumerate(cells(row, li)):
                lab = _norm_label(v)
                trips[j]["legs"][leg["id"]] = (
                    {"label": lab, "color": colors.get(lab, DEFAULT_COLOR),
                     "url": find_link(links, board, lab)} if lab else None)

    if rows[0]["type"] != "stop" or rows[-1]["type"] != "stop":
        raise TimetableError(f"{path.name}: 最初と最後の行は地点（発/着）にしてください")
    # 着→発の順に並べる
    for r in rows:
        if r["type"] == "stop":
            r["times"].sort(key=lambda k: 0 if k == "arr" else 1)

    return {"rows": rows, "trips": trips}


def _compute_joins(rows, trips):
    """隣のスジと「前の地点の発・区間・次の地点の着」がすべて同じなら同じ便＝合流とみなす。"""
    for t in trips:
        t["joins"] = {}
    for i, r in enumerate(rows):
        if r["type"] != "leg":
            continue
        a, b = rows[i - 1], rows[i + 1]
        # 前の地点に「発」がない（バス停→学校の徒歩など）ときは、着が合流済みならその続きとして扱う
        ka = f"{a['id']}.dep" if "dep" in a["times"] else f"{a['id']}.arr"
        kb = f"{b['id']}.arr" if "arr" in b["times"] else None
        if not kb:
            continue
        for t in trips:
            dep, arr = t["times"].get(ka), t["times"].get(kb)
            if not (dep and arr):
                continue
            if ka.endswith(".arr"):
                if ka not in t["joins"]:
                    continue
                dep = t["joins"][ka]
            leg = t["legs"].get(r["id"])
            sig = f"{dep}|{leg['label'] if leg else ''}|{arr}"
            t["joins"][ka] = t["joins"].get(ka, sig) if ka.endswith(".arr") else sig
            t["joins"][r["id"]] = t["joins"][kb] = sig


def _shape(tt):
    return [(r["type"], r.get("name"), tuple(r.get("times", ()))) for r in tt["rows"]]


def _compatible(part, full):
    """part の書いてある値がすべて full と一致するか（途中までしか書いていないスジの判定）。"""
    for k, v in part["times"].items():
        if v is not None and full["times"].get(k) != v:
            return False
    for k, v in part["legs"].items():
        if v is not None and (full["legs"].get(k) or {}).get("label") != v["label"]:
            return False
    return True


def _merge_trips(rows, trips):
    """複数シートのスジを1本にまとめる。
    ・すべて同じスジは1本に
    ・途中までしか書いていないスジは、同じ内容で最後まで書いてあるスジがあれば除く
    ・バスの選び方などが違うスジは両方残す
    """
    keys = [f"{r['id']}.{k}" for r in rows if r["type"] == "stop" for k in r["times"]]
    last = keys[-1]
    sig = lambda t: (tuple(t["times"].get(k) for k in keys),
                     tuple(sorted((k, v["label"]) for k, v in t["legs"].items() if v)))
    _inherit_continuation(rows, trips)
    seen, uniq = set(), []
    for t in trips:
        if sig(t) not in seen:
            seen.add(sig(t))
            uniq.append(t)
    full = [t for t in uniq if t["times"].get(last)]
    kept = [t for t in uniq if t["times"].get(last)
            or not any(_compatible(t, f) for f in full if f is not t)]

    def order(t):
        return [(0, _min(t["times"][k])) if t["times"].get(k) else (1, 0) for k in keys]
    kept.sort(key=order)
    for i, t in enumerate(kept):
        t["no"] = i + 1
    return kept


def _inherit_continuation(rows, trips):
    """途中で書くのをやめているスジ（Excelで結合セルの片側だけ先が書いてある等）は、
    最後に書いてある地点まで同じ便に乗るほかのスジから、その先を引き継ぐ。
    候補が複数あるときは、その先で一番早く出る便を選ぶ。"""
    seq = []   # 行の順に並べた (種類, キー)
    for r in rows:
        if r["type"] == "stop":
            seq += [("t", f"{r['id']}.{k}") for k in r["times"]]
        else:
            seq.append(("l", r["id"]))
    tkeys = [k for typ, k in seq if typ == "t"]

    def lastpos(t):
        return max((i for i, (typ, k) in enumerate(seq) if typ == "t" and t["times"].get(k)), default=-1)

    for t in trips:
        p = lastpos(t)
        if p < 0 or seq[p][1] == tkeys[-1]:
            continue
        # 最後に書いてある地点へ同じ便で着いているか（前の地点の発・区間・その地点の時刻が同じ）
        upto = [x for x in seq[:p + 1]]
        tail = upto[-3:] if len(upto) >= 3 else upto
        def same_vehicle(o):
            for typ, k in tail:
                a = t["times"].get(k) if typ == "t" else (t["legs"].get(k) or {}).get("label")
                b = o["times"].get(k) if typ == "t" else (o["legs"].get(k) or {}).get("label")
                if a != b:
                    return False
            return True
        cands = [o for o in trips if o is not t and lastpos(o) > p and same_vehicle(o)]
        nxt = next((k for typ, k in seq[p + 1:] if typ == "t"), None)
        cands = [o for o in cands if o["times"].get(nxt)]
        if not cands:
            continue
        src = min(cands, key=lambda o: _min(o["times"][nxt]))
        for typ, k in seq[p + 1:]:
            if typ == "t":
                t["times"][k] = src["times"].get(k)
            else:
                t["legs"][k] = src["legs"].get(k)


def _min(s):
    h, m = s.split(":")
    return int(h) * 60 + int(m)


def load_all() -> dict:
    """timetables.csv の「方向」ごとに、シートを1本の時刻表にまとめて返す。"""
    colors = load_colors()
    index = _read_csv(DATA_DIR / "timetables.csv")
    filled = [i for i, r in enumerate(index) if any(r)]
    directions: dict[str, dict] = {}
    for li, row in enumerate(index[filled[0] + 1:] if filled else [], start=(filled[0] + 2) if filled else 2):
        row = row + [""] * 5
        dname, goal, file, sheet, deadline = row[:5]
        if not dname:
            continue
        path = DATA_DIR / file
        if not path.exists():
            raise TimetableError(f"timetables.csv {li}行目: {file} が見つかりません")
        if sheet == "区間ごと":   # 区間ごとの時刻表から乗り継ぎを自動で組み立てる
            from segment_loader import load_segments
            if dname in directions:
                raise TimetableError(f"timetables.csv {li}行目: 「区間ごと」の方向は1行にしてください")
            directions[dname] = load_segments(path, colors, dname)
            continue
        if path.suffix.lower() in (".xlsx", ".xlsm") and not sheet:
            raise TimetableError(f"timetables.csv {li}行目: Excelファイルにはシート名が必要です")
        tt = load_timetable(path, colors, sheet)
        d = directions.setdefault(dname, {"name": dname, "rows": tt["rows"], "trips": [], "goals": []})
        if _shape(tt) != [(r["type"], r.get("name"), tuple(r.get("times", ()))) for r in d["rows"]]:
            raise TimetableError(f"timetables.csv {li}行目: {file}「{sheet}」の地点の並びが、"
                                 f"同じ「{dname}」のほかのシートと違います")
        d["trips"] += tt["trips"]
        if goal:
            d["goals"].append({"name": goal,
                               "deadline": _norm_time(deadline, f"timetables.csv {li}行目 締切") if deadline else None})
    if not directions:
        raise TimetableError("timetables.csv に時刻表が登録されていません")
    warnings = []
    for d in directions.values():
        if "stations" in d:   # 区間ごと：乗り継ぎと合流は segment_loader で計算済み
            continue
        d["trips"] = _merge_trips(d["rows"], d["trips"])
        _compute_joins(d["rows"], d["trips"])
        warnings += _check_order(d)
    return {"directions": list(directions.values()), "warnings": warnings}


def _check_order(d):
    """時刻が前の地点より早くなっているところを探す（入力ミスや結合セルの範囲違いの発見用）。"""
    labels = [(f"{r['id']}.{k}", f"{r['name']}{'着' if k == 'arr' else '発'}")
              for r in d["rows"] if r["type"] == "stop" for k in r["times"]]
    out = []
    for t in d["trips"]:
        prev = None
        for key, name in labels:
            v = t["times"].get(key)
            if not v:
                continue
            if prev and _min(v) < _min(prev[1]):
                out.append(f"{t['src']}（自宅 {t['times'][labels[0][0]]} 発のスジ）: "
                           f"{prev[0]} {prev[1]} のあとに {name} {v} になっています")
            prev = (name, v)
    return out
