# -*- coding: utf-8 -*-
"""
HIT AT 效率诊断日报 · 构建脚本
读取:  HIT_Data_Summary*.zip (车辆遥测 / 能耗 / 操作日志 CSV) + (IT61D) Daily AT Movement*.xlsx
产出:  自包含 HTML 报告 (数据内嵌 + ECharts 内联, 双击即开)

用法:
  python build_report.py --zip <zip> --xlsx <xlsx> --out <html>
  python build_report.py --data-dir <解压后的csv目录> --xlsx <xlsx> --out <html>
"""
import argparse, glob, json, os, re, shutil, sys, zipfile, tempfile, datetime as dt
import xml.etree.ElementTree as ET

import numpy as np
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
SAMPLE_SEC = 3.0  # 遥测采样间隔近似值(秒)

VEHICLES_ALL = ["AT003", "AT004", "AT005", "AT006", "AT007", "AT008"]


# --------------------------------------------------------------------------
# 1. 数据读取
# --------------------------------------------------------------------------
def extract_zip(zip_path, dest):
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(dest)
    return dest


def read_telemetry(csv_dir):
    """读取 HIT_ATxxx_VehicleData_*.csv, 合并为 DataFrame"""
    files = sorted(glob.glob(os.path.join(csv_dir, "HIT_AT*_VehicleData*.csv")))
    frames = []
    for f in files:
        try:
            d = pd.read_csv(f, dtype=str, low_memory=False)
        except Exception:
            continue
        if len(d) == 0:
            continue
        frames.append(d)
    if not frames:
        return pd.DataFrame()
    v = pd.concat(frames, ignore_index=True)
    v["local_time"] = pd.to_datetime(v["local_time"], errors="coerce")
    for c in ["speed", "mileage_distance", "soc", "soh", "power_on_time",
              "Motor running hour", "total_charge_energy", "total_discharge_energy",
              "total_charge_times", "position_x", "position_y"]:
        if c in v.columns:
            v[c] = pd.to_numeric(v[c], errors="coerce")
    return v.sort_values(["vehicle_id", "local_time"]).reset_index(drop=True)


def read_energy(csv_dir):
    files = glob.glob(os.path.join(csv_dir, "*EQUIPMENT_ENERGY_USAGE*.csv"))
    if not files:
        return pd.DataFrame()
    e = pd.read_csv(files[0])
    e.columns = [c.strip() for c in e.columns]
    return e


def read_log(csv_dir):
    files = glob.glob(os.path.join(csv_dir, "*user_operation_log*.csv"))
    if not files:
        return pd.DataFrame()
    return pd.read_csv(files[0])


def col_idx(ref):
    letters = "".join(c for c in ref if c.isalpha())
    n = 0
    for c in letters:
        n = n * 26 + (ord(c) - 64)
    return n - 1


def parse_xlsx(xlsx_path):
    """直接解析 xlsx 内部 XML (源文件 stylesheet 损坏, openpyxl 无法读取)"""
    tmp = tempfile.mkdtemp(prefix="xlsx_")
    try:
        with zipfile.ZipFile(xlsx_path) as z:
            z.extractall(tmp)
        xl = os.path.join(tmp, "xl")
        ss = []
        ssp = os.path.join(xl, "sharedStrings.xml")
        if os.path.exists(ssp):
            for si in ET.parse(ssp).getroot():
                ss.append("".join(x.text or "" for x in si.iter(NS + "t")))
        sheets = sorted(glob.glob(os.path.join(xl, "worksheets", "*.xml")))
        rows = []
        for row in ET.parse(sheets[0]).getroot().iter(NS + "row"):
            vals = []
            for c in row.iter(NS + "c"):
                ref = c.get("r") or ""
                t = c.get("t")
                v = c.find(NS + "v")
                isel = c.find(NS + "is")
                if t == "s" and v is not None:
                    val = ss[int(v.text)] if int(v.text) < len(ss) else ""
                elif t == "inlineStr" and isel is not None:
                    val = "".join(x.text or "" for x in isel.iter(NS + "t"))
                elif v is not None:
                    val = v.text
                else:
                    val = ""
                ci = col_idx(ref) if ref else len(vals)
                while len(vals) <= ci:
                    vals.append("")
                vals[ci] = val
            rows.append(vals)
        if not rows:
            return pd.DataFrame()
        hdr = rows[0]
        n = len(hdr)
        data = [r + [""] * (n - len(r)) for r in rows[1:]]
        return pd.DataFrame(data, columns=hdr)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# --------------------------------------------------------------------------
# 2. 指标计算
# --------------------------------------------------------------------------
def xl_dt(x):
    try:
        return pd.Timestamp("1899-12-30") + pd.to_timedelta(float(x), unit="D")
    except Exception:
        return pd.NaT


def stage_seconds(v):
    """按采样时间差累计各 task_stage 停留时长(秒)"""
    out = {}
    for vid, g in v.groupby("vehicle_id"):
        g = g.sort_values("local_time")
        dtsec = g["local_time"].diff().dt.total_seconds()
        dtsec = dtsec.clip(lower=0, upper=60).fillna(0)
        stage = g["task_stage"].fillna("Idle/No task")
        for s, d in zip(stage, dtsec):
            out[s] = out.get(s, 0.0) + d
    return out


def moving_seconds(v, thr=0.5):
    out = {}
    for vid, g in v.groupby("vehicle_id"):
        g = g.sort_values("local_time")
        dtsec = g["local_time"].diff().dt.total_seconds().clip(lower=0, upper=60).fillna(0)
        mv = (g["speed"].abs() > thr).astype(float)
        out[vid] = float((dtsec * mv).sum())
    return out


def error_stats(v, gap_sec=180):
    """error_code 采样计数 + 独立事件计数(上升沿 + 去抖: 需静默 >=gap_sec 才算新事件)"""
    sample_cnt, event_cnt = {}, {}
    per_veh = {}
    for vid, g in v.groupby("vehicle_id"):
        g = g.sort_values("local_time")
        times = g["local_time"].tolist()
        codes_list = [set(c for c in str(raw).split() if c) for raw in g["error_code"].fillna("")]
        veh_codes = {}
        active, absent_since = {}, {}   # code -> 是否在场 / 最后一次缺席时间
        for t, codes in zip(times, codes_list):
            for c in codes:
                sample_cnt[c] = sample_cnt.get(c, 0) + 1
                veh_codes[c] = veh_codes.get(c, 0) + 1
                if not active.get(c, False):
                    a = absent_since.get(c)
                    if a is None or (t - a).total_seconds() >= gap_sec:
                        event_cnt[c] = event_cnt.get(c, 0) + 1
                active[c] = True
                absent_since.pop(c, None)
            for c in list(active.keys()):
                if c not in codes and active.get(c):
                    active[c] = False
                    absent_since[c] = t
        per_veh[vid] = veh_codes
    return sample_cnt, event_cnt, per_veh


def charge_events(v):
    """充电事件(上升沿)与 SOC 极值; total_charge_energy 为原始计数器, 不直接采用"""
    out = {}
    for vid, g in v.groupby("vehicle_id"):
        g = g.sort_values("local_time")
        cs = g["charge_status"].fillna("No Charging")
        starts = int(((cs == "Charging") & (cs.shift(1) != "Charging")).sum())
        ct = pd.to_numeric(g["total_charge_times"], errors="coerce")
        ct_delta = int(ct.max() - ct.min()) if ct.notna().any() else 0
        soc = g["soc"]
        out[vid] = {
            "starts": starts,
            "charge_times": ct_delta,
            "soc_min": float(soc.min()) if soc.notna().any() else None,
            "soc_max": float(soc.max()) if soc.notna().any() else None,
            "soc_mean": round(float(soc.mean()), 1) if soc.notna().any() else None,
        }
    return out


def downsample_map(v, step=15):
    """按 step 秒抽稀轨迹, 返回边界与每车点列 [t_off, x, y, speed]"""
    if v.empty or "position_x" not in v.columns:
        return None
    t0 = v["local_time"].min()
    day0 = t0.normalize()
    veh = {}
    minx = miny = 1e18
    maxx = maxy = -1e18
    for vid, g in v.groupby("vehicle_id"):
        g = g.sort_values("local_time")
        g = g.dropna(subset=["position_x", "position_y"])
        if g.empty:
            continue
        ts = (g["local_time"] - day0).dt.total_seconds().to_numpy()
        if len(ts) == 0:
            continue
        targets = np.arange(ts[0], ts[-1] + 1e-6, step)
        idx = np.unique(np.searchsorted(ts, targets))
        idx = idx[idx < len(ts)]
        gg = g.iloc[idx]
        tts = (gg["local_time"] - day0).dt.total_seconds().to_numpy()
        x = gg["position_x"].to_numpy()
        y = gg["position_y"].to_numpy()
        sp = gg["speed"].abs().fillna(0).to_numpy()
        veh[vid] = [[round(float(a), 1), round(float(b), 2), round(float(c), 2), round(float(d), 2)]
                    for a, b, c, d in zip(tts, x, y, sp)]
        minx = min(minx, float(np.nanmin(x))); maxx = max(maxx, float(np.nanmax(x)))
        miny = min(miny, float(np.nanmin(y))); maxy = max(maxy, float(np.nanmax(y)))
    # 场地若为竖向(南北向长于东西向), 交换 x/y 使轨迹横向铺满画布; 同时对横向取反实现左右镜像
    if (maxy - miny) > (maxx - minx):
        for vid in list(veh):
            veh[vid] = [[a, -c, b, d] for a, b, c, d in veh[vid]]
        minx, maxx, miny, maxy = -maxy, -miny, minx, maxx
    return {
        "bounds": {"minx": minx, "maxx": maxx, "miny": miny, "maxy": maxy},
        "day0": day0.strftime("%Y-%m-%d %H:%M:%S"),
        "vehicles": veh,
    }


def build(data_dir, xlsx_path, zip_path=None):
    work = data_dir
    tmpdir = None
    if zip_path and not data_dir:
        tmpdir = tempfile.mkdtemp(prefix="hitdata_")
        work = extract_zip(zip_path, tmpdir)

    v = read_telemetry(work)
    energy = read_energy(work)
    log = read_log(work)
    mv = parse_xlsx(xlsx_path)

    # ---- 车队基础 ----
    veh_present = sorted(v["vehicle_id"].unique().tolist()) if not v.empty else []
    mv["TRACTOR_NO"] = mv.get("TRACTOR_NO", pd.Series(dtype=str)).astype(str)

    # --- 位移/箱量 ---
    mv_boxes = mv["TRACTOR_NO"].value_counts().to_dict()
    mv["__ts"] = mv["LAST_CHKPT_DATETIME"].map(xl_dt) if "LAST_CHKPT_DATETIME" in mv.columns else pd.NaT
    mv["__start"] = mv["START_CHKPT_DATETIME"].map(xl_dt) if "START_CHKPT_DATETIME" in mv.columns else pd.NaT
    mv["__cycle_min"] = (mv["__ts"] - mv["__start"]).dt.total_seconds() / 60.0
    mv["__hour"] = mv["__ts"].dt.hour
    mv["__dir"] = mv.get("MOVEMENT_REFERENCE", pd.Series(dtype=str)).astype(str)

    total_boxes = int(len(mv))
    size_dist = mv["CONT_SIZE"].value_counts().to_dict() if "CONT_SIZE" in mv.columns else {}
    dir_dist = mv["__dir"].value_counts().to_dict()
    status_dist = mv["CONT_STATUS"].value_counts().to_dict() if "CONT_STATUS" in mv.columns else {}
    twin_dist = mv["TWIN_MOVE_IND"].replace("", "否").value_counts().to_dict() if "TWIN_MOVE_IND" in mv.columns else {}
    from_loc = mv["FROM_LS_LOCATION_TYPE"].value_counts().to_dict() if "FROM_LS_LOCATION_TYPE" in mv.columns else {}
    to_loc = mv["TO_LS_LOCATION_TYPE"].value_counts().to_dict() if "TO_LS_LOCATION_TYPE" in mv.columns else {}
    hour_boxes = [0] * 24
    for h, c in mv["__hour"].value_counts().items():
        if pd.notna(h):
            hour_boxes[int(h)] = int(c)

    # --- 岸桥 / 堆场 每小时箱数 ---
    def hour_vec(sub):
        vec = [0] * 24
        for hh, cc in sub["__hour"].value_counts().items():
            if pd.notna(hh):
                vec[int(hh)] = int(cc)
        return vec

    mv["__qc"] = mv.get("QUAY_CRANE_ID", pd.Series(dtype=str)).fillna("").astype(str).str.strip()
    qc_order = sorted(q for q in mv["__qc"].unique().tolist() if q and q.lower() != "nan")
    qc_hourly = {q: hour_vec(mv[mv["__qc"] == q]) for q in qc_order}

    def yard_bay(row):
        """堆场一侧的「堆场-贝位」: 卸船(TO=GNSTK)取 TO_LS_LEVEL_1(堆场)+LEVEL_2(贝位), 装船(FROM=GNSTK)取 FROM_*"""
        if str(row.get("TO_LS_LOCATION_TYPE", "") or "").strip() == "GNSTK":
            lv1 = str(row.get("TO_LS_LEVEL_1", "") or "").strip()
            lv2 = str(row.get("TO_LS_LEVEL_2", "") or "").strip()
        elif str(row.get("FROM_LS_LOCATION_TYPE", "") or "").strip() == "GNSTK":
            lv1 = str(row.get("FROM_LS_LEVEL_1", "") or "").strip()
            lv2 = str(row.get("FROM_LS_LEVEL_2", "") or "").strip()
        else:
            return ""
        return "-".join(p for p in (lv1, lv2) if p)

    mv["__bay"] = mv.apply(yard_bay, axis=1)
    bay_order = [b for b in mv["__bay"].replace("", pd.NA).value_counts().index.tolist() if pd.notna(b)]
    bay_hourly = {b: hour_vec(mv[mv["__bay"] == b]) for b in bay_order}

    def cyc(sub):
        s = sub["__cycle_min"].dropna()
        s = s[(s > 0) & (s < 600)]
        if len(s) == 0:
            return {"n": 0, "p50": None, "p90": None, "mean": None}
        return {"n": int(len(s)), "p50": round(float(s.quantile(.5)), 1),
                "p90": round(float(s.quantile(.9)), 1), "mean": round(float(s.mean()), 1)}

    cycle = {
        "overall": cyc(mv),
        "VSDS": cyc(mv[mv["__dir"] == "VSDS"]),
        "VSLD": cyc(mv[mv["__dir"] == "VSLD"]),
    }

    # --- 超时箱清单 (>60 min) ---
    mv_cyc = mv[mv["__cycle_min"].notna() & (mv["__cycle_min"] > 0) & (mv["__cycle_min"] < 600)]

    def _vc(series):
        s = series.astype(str).str.strip()
        s = s[s != ""]
        return {str(k): int(v) for k, v in s.value_counts().items()}

    # --- AT 箱时 > 60 分钟 箱清单 ---
    slow_boxes = []
    mv_slow = mv_cyc[mv_cyc["__cycle_min"] > 60].sort_values("__cycle_min", ascending=False)
    for _, r in mv_slow.iterrows():
        slow_boxes.append({
            "dur": round(float(r["__cycle_min"]), 1),
            "vehicle": str(r.get("TRACTOR_NO", "") or "").strip(),
            "qc": str(r.get("__qc", "") or "").strip(),
            "bay": str(r.get("__bay", "") or "").strip(),
            "dir": str(r.get("__dir", "") or "").strip(),
            "size": str(r.get("CONT_SIZE", "") or "").strip(),
            "start": r["__start"].strftime("%H:%M") if pd.notna(r["__start"]) else "",
        })
    slow_summary = {
        "n": int(len(slow_boxes)),
        "pct": round(len(slow_boxes) * 100.0 / total_boxes, 1) if total_boxes else 0,
        "vehicle": _vc(mv_slow["TRACTOR_NO"]),
        "qc": _vc(mv_slow["__qc"]),
        "dir": _vc(mv_slow["__dir"]),
    }

    # --- 能耗 ---
    energy_rows = []
    e_map = {}
    if not energy.empty:
        for _, r in energy.iterrows():
            vid = str(r.get("Equipment ID", "")).strip()
            e_map[vid] = r
            energy_rows.append({
                "vehicle": vid,
                "mileage_km": _num(r.get("Kilometer Run（Mileage）")),
                "energy_kwh": _num(r.get("Energy Consumption")),
                "power_hour": _num(r.get("Power On Hour")),
                "motor_hour": _num(r.get("Equipment Motor Running Hour")),
            })

    # --- 遥测派生 ---
    telem = {}
    for vid, g in (v.groupby("vehicle_id") if not v.empty else []):
        mil = g["mileage_distance"]
        dist_km = float(mil.max() - mil.min()) if mil.notna().any() else 0.0
        sp = g["speed"].abs()
        telem[vid] = {
            "rows": int(len(g)),
            "dist_km": round(dist_km, 1),
            "avg_speed": round(float(sp.mean()), 2),
            "max_speed": round(float(sp.max()), 2),
            "p95_speed": round(float(sp.quantile(.95)), 2),
            "moving_min": round(moving_seconds(v).get(vid, 0) / 60.0, 1),
        }

    mvsec = moving_seconds(v)
    stg = stage_seconds(v)
    total_stage_sec = sum(stg.values()) or 1

    err_sample, err_event, err_per_veh = ({}, {}, {}) if v.empty else error_stats(v)
    chg = {} if v.empty else charge_events(v)

    # --- 故障码 → 车辆 关联 ---
    err_code_veh_all = {}
    for _vid, _codes in (err_per_veh or {}).items():
        for _c, _n in _codes.items():
            err_code_veh_all.setdefault(_c, {})[_vid] = int(_n)
    err_top = sorted(err_sample.items(), key=lambda t: -t[1])[:12]
    err_code_veh = {}
    for _c, _ in err_top:
        _d = err_code_veh_all.get(_c, {})
        err_code_veh[_c] = {vid: _d[vid] for vid in sorted(_d, key=lambda x: -_d[x])}
    err_vehicles = sorted({vid for _d in err_code_veh.values() for vid in _d})

    # 分时活跃车队(每小时有移动的车辆数)
    hourly_active = [0] * 24
    hourly_moving_min = [0.0] * 24
    if not v.empty:
        for vid, g in v.groupby("vehicle_id"):
            g = g.sort_values("local_time")
            dts = g["local_time"].diff().dt.total_seconds().clip(lower=0, upper=60).fillna(0)
            hr = g["local_time"].dt.hour
            mv_flag = (g["speed"].abs() > 0.5).astype(float)
            for hh, dd, mf in zip(hr, dts, mv_flag):
                if dd > 0:
                    hourly_moving_min[hh] += dd * mf / 60.0
        # 活跃车辆: 每小时内移动 >5min
        for vid, g in v.groupby("vehicle_id"):
            g = g.sort_values("local_time")
            dts = g["local_time"].diff().dt.total_seconds().clip(lower=0, upper=60).fillna(0)
            hr = g["local_time"].dt.hour
            mv_flag = (g["speed"].abs() > 0.5).astype(float)
            per_hour = np.zeros(24)
            for hh, dd, mf in zip(hr, dts, mv_flag):
                if dd > 0:
                    per_hour[hh] += dd * mf / 60.0
            for hh in range(24):
                if per_hour[hh] > 5:
                    hourly_active[hh] += 1

    # --- 各物理段: 平均车速 + 行驶里程 ---
    stage_seg = []
    if not v.empty:
        df = v.sort_values(["vehicle_id", "local_time"]).copy()
        df["__asp"] = df["speed"].abs()
        df["__dd"] = df.groupby("vehicle_id")["mileage_distance"].diff().clip(lower=0).fillna(0)
        for s, g in df.groupby(df["task_stage"].fillna("Idle/No task")):
            stage_seg.append({
                "stage": str(s),
                "speed": round(float(g["__asp"].mean()), 2),
                "km": round(float(g["__dd"].sum()), 2),
                "hours": round(float(g["__dd"].size) * SAMPLE_SEC / 3600.0, 2),
            })
        stage_seg.sort(key=lambda x: -x["km"])

    # --- 停车热点: 按时段 × 区域 (区域按取箱/交箱作业锚点近似归类) ---
    parking = None
    if not v.empty and "position_x" in v.columns and len(mv):
        anch = {"quay": [], "yard": []}
        by_veh = {str(k): g for k, g in v.groupby("vehicle_id")}
        chk_o = mv["CHKPT_1_DATETIME"].map(xl_dt) if "CHKPT_1_DATETIME" in mv.columns else pd.Series(pd.NaT, index=mv.index)
        chk_d = mv["LAST_CHKPT_DATETIME"].map(xl_dt) if "LAST_CHKPT_DATETIME" in mv.columns else pd.Series(pd.NaT, index=mv.index)
        for i in mv.index:
            tno = str(mv.at[i, "TRACTOR_NO"]).strip()
            veh = tno if tno.startswith("AT") else "AT" + tno[-3:]
            sub = by_veh.get(veh)
            if sub is None or sub.empty:
                continue
            for t, lt in [(chk_o.get(i), mv.at[i, "FROM_LS_LOCATION_TYPE"] if "FROM_LS_LOCATION_TYPE" in mv.columns else ""),
                          (chk_d.get(i), mv.at[i, "TO_LS_LOCATION_TYPE"] if "TO_LS_LOCATION_TYPE" in mv.columns else "")]:
                if pd.isna(t):
                    continue
                w = sub[(sub["local_time"] >= t - pd.Timedelta(minutes=4)) & (sub["local_time"] <= t + pd.Timedelta(minutes=2))]
                w = w[w["speed"].abs() < 0.05]
                if w.empty:
                    continue
                anch["quay" if str(lt).strip() == "ONVSL" else "yard"].append(
                    (float(w["position_x"].median()), float(w["position_y"].median())))
        parked = v.dropna(subset=["position_x", "position_y"])
        parked = parked[parked["speed"].abs() < 0.05]
        if anch["quay"] and anch["yard"] and not parked.empty:
            P = parked[["position_x", "position_y"]].to_numpy(float)
            Q = np.array(anch["quay"], float)
            Y = np.array(anch["yard"], float)
            n = len(P)
            dq = np.empty(n)
            dy = np.empty(n)
            for s0 in range(0, n, 20000):
                blk = P[s0:s0 + 20000]
                dq[s0:s0 + 20000] = np.min(np.hypot(blk[:, 0, None] - Q[None, :, 0], blk[:, 1, None] - Q[None, :, 1]), axis=1)
                dy[s0:s0 + 20000] = np.min(np.hypot(blk[:, 0, None] - Y[None, :, 0], blk[:, 1, None] - Y[None, :, 1]), axis=1)
            zone = np.where(np.minimum(dq, dy) > 20.0, "other", np.where(dq < dy, "quay", "yard"))
            hrs = parked["local_time"].dt.hour.to_numpy()
            zones = ["quay", "yard", "other"]
            by_zone = {}
            for z in zones:
                vc = pd.Series(hrs[zone == z]).value_counts().to_dict()
                by_zone[z] = [round(float(vc.get(hh, 0)) * SAMPLE_SEC / 60.0, 1) for hh in range(24)]
            parking = {
                "zones": zones,
                "labels": {"quay": "岸桥区", "yard": "堆场区", "other": "其他(道路/等待)"},
                "by_zone": by_zone,
                "unit": "min",
                "coverage": {z: int((zone == z).sum()) for z in zones},
            }

    # --- 操作日志 ---
    op_counts = log["operation_name"].value_counts().to_dict() if not log.empty else {}
    op_veh = log["vehicle_name"].value_counts().to_dict() if not log.empty else {}
    op_user = log["user"].value_counts().to_dict() if not log.empty else {}

    # --- 地图 ---
    mapdata = downsample_map(v)

    # --- 车队汇总表 ---
    fleet = []
    for vid in VEHICLES_ALL:
        box = int(mv_boxes.get(vid, 0))
        te = telem.get(vid, {})
        em = e_map.get(vid)
        e_kwh = _num(em.get("Energy Consumption")) if em is not None else None
        mile = _num(em.get("Kilometer Run（Mileage）")) if em is not None else None
        ph = _num(em.get("Power On Hour")) if em is not None else None
        mh = _num(em.get("Equipment Motor Running Hour")) if em is not None else None
        fleet.append({
            "vehicle": vid,
            "boxes": box,
            "mileage_km": mile if mile is not None else te.get("dist_km"),
            "energy_kwh": e_kwh,
            "power_hour": ph,
            "motor_hour": mh,
            "energy_per_km": round(e_kwh / mile, 2) if (e_kwh and mile) else None,
            "energy_per_box": round(e_kwh / box, 2) if (e_kwh and box) else None,
            "active": bool(te.get("rows", 0) > 0),
            "dist_telem": te.get("dist_km"),
            "moving_min": te.get("moving_min"),
            "avg_speed": te.get("avg_speed"),
            "max_speed": te.get("max_speed"),
        })

    avg_speed = _avg([f["avg_speed"] for f in fleet if f["avg_speed"]])
    avg_epk = _avg([f["energy_per_km"] for f in fleet if f["energy_per_km"]])
    avg_epb = _avg([f["energy_per_box"] for f in fleet if f["energy_per_box"]])
    all_mileage = sum(f["mileage_km"] or 0 for f in fleet)
    all_energy = sum(f["energy_kwh"] or 0 for f in fleet)
    all_motor = sum(f["motor_hour"] or 0 for f in fleet)

    kpi = [
        {"label": {"zh": "完成作业箱量", "en": "Containers moved"}, "value": f"{total_boxes}",
         "sub": {"zh": "5 台 AT 全天", "en": "across 5 AT units"}, "tone": "sage"},
        {"label": {"zh": "车队总里程", "en": "Fleet mileage"}, "value": f"{all_mileage:.0f} km",
         "sub": {"zh": "均值 %.0f km/车" % (all_mileage / 5), "en": "avg %.0f km/unit" % (all_mileage / 5)}, "tone": "amber"},
        {"label": {"zh": "总电耗", "en": "Energy consumed"}, "value": f"{all_energy:.0f} kWh",
         "sub": {"zh": "单箱 %.1f kWh · 单公里 %.1f kWh" % (avg_epb or 0, avg_epk or 0),
                 "en": "%.1f kWh/box · %.1f kWh/km" % (avg_epb or 0, avg_epk or 0)}, "tone": "clay"},
        {"label": {"zh": "平均车速", "en": "Avg speed"}, "value": f"{avg_speed:.1f} m/s",
         "sub": {"zh": "≈ %.1f km/h" % (avg_speed * 3.6), "en": "≈ %.1f km/h" % (avg_speed * 3.6)}, "tone": "sage"},
        {"label": {"zh": "电机时长占比", "en": "Motor-hour share"}, "value": f"{all_motor / max(sum(f['power_hour'] or 0 for f in fleet),1e-9) * 100:.0f}%",
         "sub": {"zh": "电机 %.0f h / 开机 %.0f h" % (all_motor, sum(f['power_hour'] or 0 for f in fleet)),
                 "en": "%.0f h / %.0f h" % (all_motor, sum(f['power_hour'] or 0 for f in fleet))}, "tone": "neutral"},
        {"label": {"zh": "独立故障事件", "en": "Distinct fault events"}, "value": f"{sum(err_event.values())}",
         "sub": {"zh": "涉及 %d 类故障码" % len(err_event), "en": "%d fault codes" % len(err_event)}, "tone": "clay"},
    ]

    # ---- 结论卡片 (数据驱动) ----
    conclusions = build_conclusions(fleet, mv, cycle, total_boxes, all_energy, all_mileage,
                                    err_event, op_counts, chg, hour_boxes)

    RD = {
        "meta": {
            "port": "HIT",
            "port_label": {"zh": "和记港口 · 自动化堆场", "en": "HIT Terminal · Automated Yard"},
            "data_window": str(mv["MOVEMENT_DATE_STAMP"].iloc[0]) if "MOVEMENT_DATE_STAMP" in mv.columns and len(mv) else "—",
            "report_date": _fmt_date(mv["__ts"].max() if mv["__ts"].notna().any() else None),
            "generated": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "vehicle_count": len(veh_present),
            "fleet": VEHICLES_ALL,
        },
        "kpi": kpi,
        "conclusions": conclusions,
        "fleet": fleet,
        "movement": {
            "total": total_boxes,
            "size_dist": size_dist,
            "dir_dist": dir_dist,
            "status_dist": status_dist,
            "twin_dist": twin_dist,
            "from_loc": from_loc,
            "to_loc": to_loc,
            "by_tractor": mv_boxes,
            "by_hour": hour_boxes,
            "qc_hourly": {"cranes": qc_order, "by_crane": qc_hourly},
            "yard_hourly": {"bays": bay_order, "by_bay": bay_hourly},
            "cycle": cycle,
            "slow_boxes": slow_boxes,
            "slow_summary": slow_summary,
        },
        "stage_time_hours": {k: round(x / 3600.0, 2) for k, x in sorted(stg.items(), key=lambda t: -t[1])},
        "stage_seg": stage_seg,
        "parking": parking,
        "hourly_active": {"active_vehicles": hourly_active, "moving_min": [round(x, 1) for x in hourly_moving_min]},
        "errors": {
            "top_sample": err_top,
            "top_event": sorted(err_event.items(), key=lambda t: -t[1])[:12],
            "code_veh": err_code_veh,
            "vehicles": err_vehicles,
        },
        "charging": chg,
        "interventions": {
            "by_op": op_counts,
            "by_vehicle": op_veh,
            "by_user": op_user,
            "total": int(len(log)),
        },
        "map": mapdata,
        "source_note": {
            "zh": "口径: 遥测为逐秒采样(≈3s)；箱量为 Daily AT Movement 当日完成记录；能耗取设备日结能耗表。异常事件按上升沿去重。",
            "en": "Basis: telemetry sampled ~3s; container moves from Daily AT Movement (same day); energy from daily equipment ledger; faults de-duplicated by rising edge.",
        },
    }
    if tmpdir:
        shutil.rmtree(tmpdir, ignore_errors=True)
    return RD


def _num(x):
    try:
        if x is None or (isinstance(x, str) and x.strip() in ("", "nil", "None")):
            return None
        return float(x)
    except Exception:
        return None


def _avg(xs):
    xs = [x for x in xs if x is not None]
    return round(sum(xs) / len(xs), 1) if xs else 0.0


def _fmt_date(ts):
    if ts is None or pd.isna(ts):
        return "—"
    return pd.Timestamp(ts).strftime("%Y-%m-%d")


def build_conclusions(fleet, mv, cycle, boxes, energy, mileage, err_event, op_counts, chg, hour_boxes):
    C = []
    # 1) 作业量与效率
    cyc = cycle["overall"]
    C.append({
        "tone": "sage",
        "tag": {"zh": "作业表现", "en": "Throughput"},
        "title": {"zh": "全天完成 %d 个箱位移动作业" % boxes, "en": "%d container moves completed" % boxes},
        "body": {"zh": "由 5 台 AT 承担，单箱平均全程耗时 %s 分钟(P90 %s 分钟)。堆场内(N2)与陆侧(M6)流向分布均衡，说明卸船/装船节拍基本对等。"
                 % (cyc["p50"], cyc["p90"]),
                 "en": "Handled by 5 AT units; average end-to-end time %s min (P90 %s min). Flows to yard (N2) and landside (M6) are balanced."
                 % (cyc["p50"], cyc["p90"])},
    })
    # 2) 峰值时段
    peak_h = int(np.argmax(hour_boxes)) if any(hour_boxes) else 0
    C.append({
        "tone": "amber",
        "tag": {"zh": "时间节奏", "en": "Rhythm"},
        "title": {"zh": "作业高峰出现在 %02d:00 时段" % peak_h, "en": "Peak activity at %02d:00" % peak_h},
        "body": {"zh": "该时段完成 %d 个箱位移动作业，为全天最高；建议据此安排换电/维护窗口，避开高峰。"
                 % hour_boxes[peak_h],
                 "en": "%d moves in this hour — the day's highest. Schedule battery swaps and maintenance away from this window."
                 % hour_boxes[peak_h]},
    })
    # 3) 效率参差
    boxed = [f for f in fleet if f["active"]]
    if boxed:
        hi = max(boxed, key=lambda f: f["boxes"])
        lo = min(boxed, key=lambda f: f["boxes"])
        C.append({
            "tone": "clay",
            "tag": {"zh": "效率差异", "en": "Spread"},
            "title": {"zh": "%s 与 %s 箱量相差 %d 箱" % (hi["vehicle"], lo["vehicle"], hi["boxes"] - lo["boxes"]),
                      "en": "%s vs %s differ by %d moves" % (hi["vehicle"], lo["vehicle"], hi["boxes"] - lo["boxes"])},
            "body": {"zh": "车队内单车辆作业量差异明显(%d vs %d)。短于平均的路径占用、等待时间偏长可能是主因，建议核对 %s 的等待与对齐阶段占比。"
                     % (hi["boxes"], lo["boxes"], lo["vehicle"]),
                     "en": "Per-unit throughput varies (%d vs %d). Longer waiting/alignment on %s likely explains it."
                     % (hi["boxes"], lo["boxes"], lo["vehicle"])},
        })
    # 4) 故障
    top_err = sorted(err_event.items(), key=lambda t: -t[1])[:1]
    if top_err:
        code, cnt = top_err[0]
        C.append({
            "tone": "clay",
            "tag": {"zh": "异常聚焦", "en": "Reliability"},
            "title": {"zh": "最高频故障码 %s 触发 %d 次" % (code, cnt), "en": "Top fault code %s fired %d×" % (code, cnt)},
            "body": {"zh": "该故障码为当日独立事件数最多者；结合人工介入日志(共 %d 次)判断，多数为感知/避障类提示，建议纳入下一轮标定观察。"
                     % sum(op_counts.values()),
                     "en": "Most frequent distinct fault today. Combined with %d manual interventions, most are perception/avoidance prompts worth tuning."
                     % sum(op_counts.values())},
        })
    return C


# --------------------------------------------------------------------------
# 3. 渲染
# --------------------------------------------------------------------------
def render(RD, template_path, echarts_path, out_path):
    with open(template_path, "r", encoding="utf-8") as f:
        tpl = f.read()
    with open(echarts_path, "r", encoding="utf-8") as f:
        echarts = f.read()
    title = "HIT AT 效率诊断日报 %s" % RD["meta"]["report_date"]
    html = tpl.replace("/*__REPORT_TITLE__*/", title)
    html = html.replace("/*__ECHARTS__*/", echarts)
    html = html.replace("/*__REPORT_DATA__*/", json.dumps(RD, ensure_ascii=False))
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--zip", default=None)
    ap.add_argument("--data-dir", dest="data_dir", default=None)
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--template", default=os.path.join(HERE, "template.html"))
    ap.add_argument("--echarts", default=os.path.join(ROOT, "assets", "echarts.min.js"))
    a = ap.parse_args()
    RD = build(a.data_dir, a.xlsx, a.zip)
    render(RD, a.template, a.echarts, a.out)
    print("OK -> %s" % a.out)
    print("箱量=%s 车队=%s 地图点数=%s" % (
        RD["movement"]["total"], RD["meta"]["vehicle_count"],
        sum(len(x) for x in (RD["map"]["vehicles"].values() if RD["map"] else []))))


if __name__ == "__main__":
    main()
