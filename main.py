import flet as ft
import openpyxl
import os
import sys
import ctypes
import shutil
import threading
import time
import asyncio
import ssl
import csv
import json
import urllib.request
import urllib.error
from datetime import datetime
import traceback

if os.name == "nt":
    from ctypes import wintypes
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
else:
    wintypes = None
    WNDENUMPROC = None


def _platform_name(page):
    """取小写平台标识：优先嵌入环境变量 FLET_PLATFORM，其次会话属性；缺失/非法返回空串"""
    raw = os.environ.get("FLET_PLATFORM")
    if not raw:
        try:
            raw = page._get_attr("platform")
        except Exception:
            raw = None
    if not raw:
        try:
            raw = getattr(getattr(page, "platform", None), "value", None)
        except Exception:
            raw = None
    if raw is None:
        return ""
    return str(raw).strip().lower()


def _is_web(page):
    try:
        return bool(page.web)
    except Exception:
        return False


def _is_desktop(page):
    """是否桌面端（Windows/macOS/Linux）；Web、iOS/Android 及平台未知的嵌入环境均按非桌面处理"""
    if _is_web(page):
        return False
    p = _platform_name(page)
    if p in ("ios", "android"):
        return False
    if p in ("windows", "macos", "linux"):
        return True
    return False


def _is_mobile(page):
    """是否手机端（iOS/Android 嵌入环境）；Web 与桌面均按非手机处理"""
    return not _is_web(page) and not _is_desktop(page)


def _write_diag(page):
    """真机诊断：把平台判定各来源与资源存在性写入 diag.log（私有目录 + 外部应用目录）"""
    try:
        lines = []

        def add(k, v):
            lines.append(f"{k} = {v!r}")

        add("VERSION", globals().get("VERSION"))
        add("time", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        add("sys.platform", sys.platform)
        add("sys.executable", sys.executable)
        add("sys.version", sys.version)
        add("cwd", os.getcwd())
        add("RES_DIR", RES_DIR)
        add("WORK_DIR", WORK_DIR)
        add("env.FLET_PLATFORM", os.environ.get("FLET_PLATFORM"))
        for k in sorted(os.environ):
            if any(t in k.upper() for t in ("PLATFORM", "FLUTTER", "ANDROID", "FLET_")):
                add("env." + k, os.environ[k])
        raw_plat = raw_web = plat_exc = None
        try:
            raw_plat = page._get_attr("platform")
        except Exception as e:
            plat_exc = repr(e)
        try:
            raw_web = page._get_attr("web")
        except Exception as e:
            plat_exc = repr(e)
        add("page.attr.platform(raw)", raw_plat)
        add("page.attr.web(raw)", raw_web)
        add("page.web", page.web)
        try:
            add("page.platform.enum", repr(getattr(getattr(page, "platform", None), "value", None)))
        except Exception as e:
            add("page.platform.enum", "EXC " + repr(e))
        add("platform_exc", plat_exc)
        add("_platform_name", _platform_name(page))
        add("_is_web", _is_web(page))
        add("_is_desktop", _is_desktop(page))
        add("_is_mobile", _is_mobile(page))
        for f in ("cover.jpg", "brand_logo.png", "fault_icon.png", "select_icon.png",
                  "repl_icon.png", "logo.png", "servo_data.xlsx", "fault_db.json"):
            p = os.path.join(RES_DIR, f)
            exists = os.path.exists(p)
            sz = os.path.getsize(p) if exists else None
            add("res:" + f, f"exists={exists} size={sz}")
        text = "\n".join(lines) + "\n"
        try:
            os.makedirs(WORK_DIR, exist_ok=True)
            with open(os.path.join(WORK_DIR, "diag.log"), "w", encoding="utf-8") as fp:
                fp.write(text)
        except Exception:
            pass
        for d in ("/storage/emulated/0/Android/data/com.jwy.servo.servo_tool/files",
                  "/sdcard/Android/data/com.jwy.servo.servo_tool/files"):
            try:
                os.makedirs(d, exist_ok=True)
                with open(os.path.join(d, "diag.log"), "w", encoding="utf-8") as fp:
                    fp.write(text)
            except Exception:
                pass
    except Exception:
        pass


def _copy_file(src, dst):
    """跨平台复制文件内容；不复制权限/时间戳/扩展属性，规避安卓私有目录 xattr 权限报错"""
    try:
        shutil.copyfile(src, dst)
    except OSError:
        with open(src, "rb") as f, open(dst, "wb") as g:
            g.write(f.read())


def _res_dir():
    """打包成 exe 后资源(logo/表格)在 _MEIPASS；直接运行则在源码目录"""
    if hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()


def _work_dir():
    """导出 BOM 的落盘目录：exe 所在目录 / 源码目录"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()


RES_DIR = _res_dir()
WORK_DIR = _work_dir()
RES_DATA = os.path.join(RES_DIR, "servo_data.xlsx")
DATA_FILE = os.path.join(WORK_DIR, "servo_data.xlsx")

# ---------- 纠错 / 版本恢复：原始基线 + 历次快照 + 台账 ----------
ORIGINAL_FILE = os.path.join(WORK_DIR, "servo_data.original.xlsx")
FIX_DIR = os.path.join(WORK_DIR, "纠错记录")
SNAP_DIR = os.path.join(FIX_DIR, "snapshots")
FIX_LOG = os.path.join(FIX_DIR, "纠错台账.xlsx")

# ---------- 配置 / 日志 / 导出记录 ----------
SETTINGS_FILE = os.path.join(WORK_DIR, "settings.json")
LOG_DIR = os.path.join(WORK_DIR, "logs")
LOG_FILE = os.path.join(LOG_DIR, "运行日志.log")
EXPORT_DIR = os.path.join(WORK_DIR, "清单文件夹")
EXPORT_LOG = os.path.join(EXPORT_DIR, "导出记录.json")
REPL_EXPORT_DIR = os.path.join(WORK_DIR, "替换清单文件夹")
REPL_EXPORT_LOG = os.path.join(REPL_EXPORT_DIR, "导出记录.json")
DATA_EXPORT_DIR = os.path.join(WORK_DIR, "数据导出")
COVER_IMG = os.path.join(RES_DIR, "cover.jpg")
BRAND_LOGO = os.path.join(RES_DIR, "brand_logo.png")
FAULT_ICON = os.path.join(RES_DIR, "fault_icon.png")
SELECT_ICON = os.path.join(RES_DIR, "select_icon.png")
REPL_ICON = os.path.join(RES_DIR, "repl_icon.png")
FAULT_DB_FILE = os.path.join(RES_DIR, "fault_db.json")
# 故障码库纠错：可写副本 + 出厂原始基线 + 历次快照 + 台账（与选型纠错同一目录体系）
FAULT_DB_WRITABLE = os.path.join(WORK_DIR, "fault_db.json")
FAULT_DB_ORIGINAL = os.path.join(WORK_DIR, "fault_db.original.json")
FAULT_SNAP_DIR = os.path.join(FIX_DIR, "故障码快照")
FAULT_FIX_LOG = os.path.join(FIX_DIR, "故障码纠错台账.json")
JWY_WEBSITE = "https://sz-jwy.com/"

# 其他端版本下载直链（EasyLink 短链，点击在系统浏览器打开落地页下载）
DOWNLOAD_LINKS = [
    {"label": "安卓全能版", "desc": "选型匹配 · 产品替换 · 故障码查询 · BOM 导出，功能最全", "url": "https://easylink.cc/2qxmw"},
    {"label": "安卓轻量故障版", "desc": "仅故障码 / 警告码查询，体积小、安装快", "url": "https://easylink.cc/d4prkq"},
    {"label": "苹果轻量故障版", "desc": "iPhone 端故障码 / 警告码查询", "url": "https://easylink.cc/lryzuu"},
]

# ---------- 自定义主题：用户导入的背景图素材（可写目录，随程序走） ----------
CUSTOM_DIR = os.path.join(WORK_DIR, "custom_assets")
CUSTOM_IMG_KEYS = ("bg", "cover", "card", "dialog")

# ---------- 故障码 / 警告码查询：系列清单 + 码表（外置 fault_db.json，按手册逐条补录） ----------
# 结构：{系列: {码: {"title": 名称, "level": 级别, "reset": 能否复位,
#                  "mechanism": 故障机理, "rows": [[故障原因, 确认方法, 处理措施], ...]}}}
FAULT_SERIES = ["SV630N", "SV630P", "SV660N", "SV660P", "SV670N", "SV670P",
                "SV680N", "SV680P", "SV600EN", "SV600EP", "SV605N", "SV605P", "SV635N"]


def _load_fault_db(path=None):
    path = path or FAULT_DB_WRITABLE
    for cand in (path, FAULT_DB_FILE):
        try:
            if os.path.exists(cand):
                with open(cand, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
    return {}


def _ensure_fault_db():
    """打包后内置资源只读：首次把内置故障库复制到可写目录，并建立出厂原始基线"""
    os.makedirs(WORK_DIR, exist_ok=True)
    if not os.path.exists(FAULT_DB_WRITABLE):
        if os.path.exists(FAULT_DB_FILE):
            _copy_file(FAULT_DB_FILE, FAULT_DB_WRITABLE)
            _log("从内置资源复制故障码库到可写目录")
        elif os.path.exists(FAULT_DB_ORIGINAL):
            _copy_file(FAULT_DB_ORIGINAL, FAULT_DB_WRITABLE)
    if not os.path.exists(FAULT_DB_ORIGINAL):
        src = FAULT_DB_WRITABLE if os.path.exists(FAULT_DB_WRITABLE) else FAULT_DB_FILE
        if src and os.path.exists(src):
            _copy_file(src, FAULT_DB_ORIGINAL)


def _dump_fault_db():
    with open(FAULT_DB_WRITABLE, "w", encoding="utf-8") as f:
        json.dump(FAULT_DB, f, ensure_ascii=False, indent=2)


def _next_fault_snapshot_index():
    if not os.path.isdir(FAULT_SNAP_DIR):
        return 1
    n = sum(1 for f in os.listdir(FAULT_SNAP_DIR)
            if f.startswith("故障纠错_") and f.endswith(".json"))
    return n + 1


def _read_fault_fix_log():
    if not os.path.exists(FAULT_FIX_LOG):
        return []
    try:
        with open(FAULT_FIX_LOG, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            return data
    except Exception:
        pass
    return []


def _write_fault_fix_log(entries):
    os.makedirs(FIX_DIR, exist_ok=True)
    with open(FAULT_FIX_LOG, "w", encoding="utf-8") as f:
        json.dump(entries, f, ensure_ascii=False, indent=2)


FAULT_XLSX_HEADER = ["系列", "故障码", "故障名称", "级别", "能否复位", "故障机理",
                      "故障原因", "确认方法", "处理措施"]


FAULT_DB = _load_fault_db()

DEFAULT_SETTINGS = {
    "default_voltage": "AC220",
    "default_depth": "3.0",
    "default_high_flex": True,
    "watermark": "姓许的苏剑",
    "bom_format": "xlsx",
    "bom_title": "伺服电机与驱动器配置清单",
    "bom_company": "",
    "bom_footer": "",
    "col_type": True,
    "col_model": True,
    "col_qty": True,
    "col_note": True,
    "data_version": "",
    "data_date": "",
    "theme_mode": "",
    "font_family": "",
    "custom_palette": {},
    "custom_bg_image": "",
    "custom_cover_image": "",
    "custom_card_image": "",
    "custom_dialog_image": "",
    "last_fault_series": "",
    "last_fault_code": "",
    "ai_provider": "zhipu",
    "ai_base_url": "",
    "ai_model": "",
    "ai_api_key": "",
}


def _load_settings():
    cfg = dict(DEFAULT_SETTINGS)
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            if isinstance(saved, dict):
                cfg.update({k: v for k, v in saved.items() if k in DEFAULT_SETTINGS})
    except Exception:
        pass
    cfg["watermark"] = "姓许的苏剑"
    return cfg


def _save_settings(cfg):
    os.makedirs(WORK_DIR, exist_ok=True)
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ---------- AI 故障诊断：云端兼容接口 + 本地故障库召回 ----------
AI_PROVIDERS = {
    "zhipu": {"label": "智谱 GLM（免费）", "base_url": "https://open.bigmodel.cn/api/paas/v4", "model": "glm-4.7-flash"},
    "dashscope": {"label": "阿里云百炼 Qwen（免费额度）", "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1", "model": "qwen3-flash"},
    "siliconflow": {"label": "硅基流动（免费额度）", "base_url": "https://api.siliconflow.cn/v1", "model": "Qwen2.5-7B-Instruct"},
    "custom": {"label": "自定义兼容端点", "base_url": "", "model": ""},
}

AI_SYSTEM_PROMPT = (
    "你是汇川伺服售后诊断助手。只能依据下面提供的本地故障库条目回答，不得编造条目之外的型号、参数、故障码或处理措施；"
    "若条目与问题无关或库内没有对应信息，直接说明“本地库内未收录对应故障”。"
    "回答分两部分：① 原因判断；② 排查步骤（按条目里的原因、确认方法、处理措施整理，逐条列出）。语言简洁专业。"
)


def _fault_recall_text(rec):
    parts = [str(rec.get("title") or ""), str(rec.get("mechanism") or "")]
    for r in rec.get("rows", []):
        parts.extend(str(x or "") for x in r)
    return " ".join(parts).lower()


def _ai_recall(query, top=5):
    q = (query or "").lower().strip()
    kws = [k for k in re.split(r"[\s,，。、/;；:：()（）]+", q) if len(k) >= 2]
    scored = []
    for s, db in FAULT_DB.items():
        for c, rec in db.items():
            text = _fault_recall_text(rec)
            score = sum(1 for k in kws if k in text) if kws else 0
            if q and q in c.lower():
                score += 5
            if score > 0:
                scored.append((score, s, c, rec))
    scored.sort(key=lambda x: -x[0])
    return scored[:top]


def _ai_format_entries(hits):
    blocks = []
    for _, s, c, rec in hits:
        lines = [f"系列：{s}　故障码：{c}　名称：{rec.get('title') or ''}"]
        if rec.get("level"):
            lines.append(f"级别：{rec['level']}")
        if rec.get("reset"):
            lines.append(f"能否复位：{rec['reset']}")
        if rec.get("mechanism"):
            lines.append(f"机理：{rec['mechanism']}")
        for i, r in enumerate(rec.get("rows", []) or [], 1):
            r = (r + ["", "", ""])[:3]
            lines.append(f"排除{i}：原因 {r[0]}｜确认 {r[1]}｜措施 {r[2]}")
        blocks.append("；".join(lines))
    return "\n\n".join(blocks)


def _ai_chat(cfg, messages, ssl_ctx, timeout=30):
    base = (cfg.get("ai_base_url") or "").rstrip("/")
    if not base:
        raise RuntimeError("未配置接口地址")
    payload = {
        "model": cfg.get("ai_model") or "",
        "messages": messages,
        "temperature": 0.3,
        "stream": False,
    }
    req = urllib.request.Request(
        base + "/chat/completions",
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={"Authorization": f"Bearer {cfg.get('ai_api_key') or ''}", "Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ssl_ctx) as r:
            data = json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as he:
        detail = ""
        try:
            body = he.read().decode("utf-8", errors="ignore")
            j = json.loads(body)
            err = j.get("error")
            detail = err.get("message") if isinstance(err, dict) else j.get("message")
            detail = detail or body[:200]
        except Exception:
            pass
        raise RuntimeError(f"HTTP {he.code} {he.reason}" + (f"：{detail}" if detail else ""))
    return data["choices"][0]["message"]["content"]


def _log(event, level="INFO"):
    """运行日志：任何异常都不允许影响主流程"""
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} [{level}] {event}\n"
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass


def _fmt_size(path):
    try:
        b = os.path.getsize(path)
        return f"{b / 1024 / 1024:.1f} MB" if b >= 1024 * 1024 else f"{b / 1024:.0f} KB"
    except Exception:
        return ""


def _fmt_mtime(path):
    try:
        return datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M")
    except Exception:
        return ""

FIX_FIELDS = [
    ("驱动器型号", "drive_model"),
    ("动力线", "power_line_base"),
    ("编码线", "encoder_line_base"),
    ("CN1 接头", "cn1"),
    ("电池套件", "battery"),
    ("通讯类型", "comm"),
    ("编码器短码", "encoder_short"),
    ("编码器形式", "encoder_type"),
    ("免电池磁编", "magnet_encoder"),
    ("需电池", "battery_need"),
    ("是否抱闸", "has_brake"),
    ("是否油封", "has_oilseal"),
    ("法兰尺寸", "flange_size"),
    ("备注", "remark"),
]
FIX_LOG_HEADER = ["序号", "时间", "电机型号", "驱动器", "字段", "改前值", "改后值", "备注", "快照文件"]


def _ensure_runtime_data():
    """打包后内置资源只读：首次把内置表复制到可写的程序目录"""
    os.makedirs(WORK_DIR, exist_ok=True)
    if not os.path.exists(DATA_FILE):
        if os.path.exists(RES_DATA):
            _copy_file(RES_DATA, DATA_FILE)
            _log("从内置资源复制数据库到可写目录")
        elif os.path.exists(ORIGINAL_FILE):
            _copy_file(ORIGINAL_FILE, DATA_FILE)
            _log("出厂基线缺失，从原始基线恢复数据库")


def _ensure_original():
    """首次运行把当前表建立为出厂原始基线，之后永不被纠错覆盖"""
    if os.path.exists(ORIGINAL_FILE):
        return
    src = DATA_FILE if os.path.exists(DATA_FILE) else (RES_DATA if os.path.exists(RES_DATA) else None)
    if src:
        _copy_file(src, ORIGINAL_FILE)


def _next_snapshot_index():
    if not os.path.isdir(SNAP_DIR):
        return 1
    n = sum(1 for f in os.listdir(SNAP_DIR)
            if f.startswith("纠错_") and f.endswith(".xlsx"))
    return n + 1


def _read_fix_log():
    if not os.path.exists(FIX_LOG):
        return []
    wb = openpyxl.load_workbook(FIX_LOG)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(h) for h in rows[0]]
    return [dict(zip(header, r)) for r in rows[1:] if r and any(v is not None for v in r)]


def _write_fix_log(entries):
    os.makedirs(FIX_DIR, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "纠错台账"
    ws.append(FIX_LOG_HEADER)
    for e in entries:
        ws.append([e.get(h) for h in FIX_LOG_HEADER])
    wb.save(FIX_LOG)

# ---------- 窗口图标：用 logo.png 生成 app.ico，运行时替换标题栏/任务栏默认图标 ----------
IMAGE_ICON = 1
LR_LOADFROMFILE = 0x0010
WM_SETICON = 0x0080
ICON_SMALL = 0
ICON_BIG = 1


def _find_logo_path():
    for d in (WORK_DIR, RES_DIR):
        for ext in (".png", ".jpg", ".jpeg", ".webp", ".bmp"):
            fp = os.path.join(d, "logo" + ext)
            if os.path.exists(fp):
                return fp
    return None


def _ensure_app_ico():
    """把 logo 转成多尺寸 app.ico，返回 ico 路径；失败返回 None"""
    ico_path = os.path.join(WORK_DIR, "app.ico")
    if os.path.exists(ico_path):
        return ico_path
    logo = _find_logo_path()
    if not logo:
        return None
    try:
        from PIL import Image
        img = Image.open(logo).convert("RGBA")
        img.save(ico_path, format="ICO",
                 sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
        return ico_path
    except Exception:
        return None


def _apply_window_icon(title):
    """后台线程：找到本程序顶层窗口，替换标题栏与任务栏图标"""
    ico_path = _ensure_app_ico()
    if not ico_path:
        return

    def _worker():
        user32 = ctypes.windll.user32
        user32.LoadImageW.restype = wintypes.HICON
        user32.LoadImageW.argtypes = [
            wintypes.HINSTANCE, wintypes.LPCWSTR, wintypes.UINT,
            ctypes.c_int, ctypes.c_int, wintypes.UINT,
        ]
        h_small = user32.LoadImageW(None, ico_path, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
        h_big = user32.LoadImageW(None, ico_path, IMAGE_ICON, 32, 32, LR_LOADFROMFILE)
        if not h_small and not h_big:
            return

        for i in range(8):
            if i > 0:
                time.sleep(0.2)
            target = []

            @WNDENUMPROC
            def _enum(hwnd, _):
                if user32.IsWindowVisible(hwnd):
                    n = user32.GetWindowTextLengthW(hwnd)
                    if n:
                        buf = ctypes.create_unicode_buffer(n + 1)
                        user32.GetWindowTextW(hwnd, buf, n + 1)
                        if title and title in buf.value:
                            target.append(hwnd)
                return True

            user32.EnumWindows(_enum, 0)
            if not target:
                continue
            hwnd = target[-1]
            if h_small:
                user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, h_small)
            if h_big:
                user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, h_big)
            return

    threading.Thread(target=_worker, daemon=True).start()

from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

HEADER = [
    "series_code", "inertia", "capacity", "power_text", "power_code",
    "rated_speed", "speed_code", "voltage", "encoder_type", "encoder_short",
    "encoder_tag", "magnet_encoder", "battery_need", "shaft_type", "shaft_code",
    "brake_oilseal_code", "has_brake", "has_oilseal", "comm", "motor_model",
    "drive_model", "power_line_base", "encoder_line_base", "cn1", "battery",
    "remark", "flange_size"
]

# 驱动器系列 -> (通讯, 编码器) 自动匹配
# 驱动器系列 -> (通讯, 编码器) 自动匹配
# 630N/660N 为 EtherCAT 总线型；本套 630N 数据全部为 T3 = 18bit 多圈
SERIES_PRESET = {
    "SV630N": ("EtherCAT", "全部"),
    "SV660N": ("EtherCAT", "全部"),
    "SV670N": ("EtherCAT", "全部"),
    "SV680N": ("EtherCAT", "全部"),
    "SV600EN": ("EtherCAT", "全部"),
    "SV605N": ("EtherCAT", "全部"),
    "SV605P": ("脉冲", "全部"),
    "SV600EP": ("脉冲", "全部"),
    "SV630P": ("脉冲", "全部"),
    "SV660P": ("脉冲", "全部"),
    "SV670P": ("脉冲", "全部"),
    "SV680P": ("脉冲", "全部"),
}

import re

# 线缆型号统一格式：S6-L-B107-3.0-T / S6-L-P124-5.0
# 前缀(族) - 长度 - 高柔标记，长度与是否高柔由界面选择决定
LINE_RE = re.compile(r"^(S6-L-[A-Z]+\d+)(?:-\d+\.\d+)?(?:-T)?$")


def build_line(code, depth, high_flex):
    """把表里的线缆型号改写成用户选择的长度/高柔版本"""
    if not code:
        return ""
    code = str(code).strip()
    m = LINE_RE.match(code)
    if not m:
        return code
    base = m.group(1)
    out = f"{base}-{depth}"
    if high_flex:
        out += "-T"
    return out


VERSION = "3.0.6"
LATEST_DATE = "2026-11-25"

# ---------- 深浅色主题调色板 ----------
# 多套外观主题；控件颜色统一从 PALETTES[mode] 取，避免深色系统下深底深字
PALETTES = {
    "light": {
        "bg": "#F4F6FA", "card": "#FFFFFF", "card_alt": "#F7FAFD", "card_green": "#F4F8F2",
        "border": "#E3E9F2", "border_green": "#DCEBD6",
        "text": "#1F2733", "title": "#123A5A", "sub": "#5A6675", "muted": "#7B8794",
        "body2": "#3B4754", "link": "#3B76B5",
        "divider": "#EDF1F6", "chip": "#DCE6F1", "header_bg": "#F0F4F9",
        "tag_warn_bg": "#FFF3E0", "tag_ok_bg": "#E9F7EE", "row_alt": "#F7FAFD",
        "primary": "#2E6DA4", "success": "#1E7A3C", "danger": "#B23A3A",
        "gray": "#5A6B7B", "warn": "#B26A00",
        "banner": "#123A5A", "banner_sub": "#A9C4DE",
    },
    "dark": {
        "bg": "#1B2029", "card": "#272E39", "card_alt": "#222933", "card_green": "#243029",
        "border": "#39424F", "border_green": "#39483E",
        "text": "#E6EAF0", "title": "#93C5EA", "sub": "#C3CDD8", "muted": "#95A2B1",
        "body2": "#D2D9E2", "link": "#82B4E6",
        "divider": "#313A46", "chip": "#31404F", "header_bg": "#313A46",
        "tag_warn_bg": "#473826", "tag_ok_bg": "#243029", "row_alt": "#222933",
        "primary": "#4E8FC9", "success": "#3FA568", "danger": "#D66666",
        "gray": "#5A6B7B", "warn": "#E0A95E",
        "banner": "#123A5A", "banner_sub": "#A9C4DE",
    },
    "brand": {
        "bg": "#E8F0FB", "card": "#FFFFFF", "card_alt": "#EFF6FE", "card_green": "#F4F8F2",
        "border": "#C7DBF3", "border_green": "#DCEBD6",
        "text": "#18283C", "title": "#0F3D8C", "sub": "#425C7C", "muted": "#6E8199",
        "body2": "#2E425C", "link": "#2456A6",
        "divider": "#D8E6F6", "chip": "#D5E6FA", "header_bg": "#D8E8FA",
        "tag_warn_bg": "#FDEEDD", "tag_ok_bg": "#E9F7EE", "row_alt": "#EFF6FE",
        "primary": "#2456A6", "success": "#1E7A3C", "danger": "#B23A3A",
        "gray": "#54677C", "warn": "#EE8A12",
        "banner": "#0F3D8C", "banner_sub": "#CFE0F6",
    },
    "mint": {
        "bg": "#EEF5EF", "card": "#FFFFFF", "card_alt": "#F2F8F2", "card_green": "#F4F8F2",
        "border": "#D7E8D9", "border_green": "#DCEBD6",
        "text": "#1E2A22", "title": "#1F6B3B", "sub": "#4E6356", "muted": "#71857A",
        "body2": "#33463A", "link": "#2E8A52",
        "divider": "#E4F0E5", "chip": "#DCEFE0", "header_bg": "#E6F2E8",
        "tag_warn_bg": "#FFF3E0", "tag_ok_bg": "#E9F7EE", "row_alt": "#F2F8F2",
        "primary": "#2E8A52", "success": "#1E7A3C", "danger": "#B23A3A",
        "gray": "#5A6B60", "warn": "#C77A17",
        "banner": "#1F6B3B", "banner_sub": "#C4E4CF",
    },
    "night": {
        "bg": "#10151F", "card": "#1A2230", "card_alt": "#161D29", "card_green": "#1B2A22",
        "border": "#2C3748", "border_green": "#31453A",
        "text": "#E8EEF6", "title": "#7FB8F0", "sub": "#C2CEDD", "muted": "#93A3B6",
        "body2": "#D3DDE9", "link": "#8FBEF2",
        "divider": "#232D3C", "chip": "#263244", "header_bg": "#232D3C",
        "tag_warn_bg": "#4A3A24", "tag_ok_bg": "#1B2A22", "row_alt": "#161D29",
        "primary": "#5B9BD9", "success": "#3FA568", "danger": "#D66666",
        "gray": "#5A6B7B", "warn": "#E0A95E",
        "banner": "#16324F", "banner_sub": "#A9C4DE",
    },
    "sepia": {
        "bg": "#F6F1E7", "card": "#FFFFFF", "card_alt": "#FAF5EC", "card_green": "#F7F4EA",
        "border": "#E9DEC9", "border_green": "#E4E6CF",
        "text": "#2C261B", "title": "#7A4E1D", "sub": "#6B5C43", "muted": "#8A7A5E",
        "body2": "#463B29", "link": "#A06A24",
        "divider": "#F0E8D8", "chip": "#ECDFC6", "header_bg": "#F1E9D8",
        "tag_warn_bg": "#FDEEDD", "tag_ok_bg": "#EEF5E6", "row_alt": "#FAF5EC",
        "primary": "#B07424", "success": "#5E8A34", "danger": "#B23A3A",
        "gray": "#6B5C43", "warn": "#C77A17",
        "banner": "#7A4E1D", "banner_sub": "#EBD9BC",
    },
}

# 圆润可爱风统一圆角档位：主卡片 / 次级面板 / 输入与下拉 / 小徽章
R_CARD = 16
R_PANEL = 12
R_FIELD = 13
R_BADGE = 10
R_BTN = 12

# 主题元数据：图标 / 名称 / 描述，三处选择入口统一从这里生成
THEME_META = {
    "light":  {"icon": "☀️", "label": "浅色模式", "desc": "明亮清爽，适合白天使用"},
    "dark":   {"icon": "🌙", "label": "深色模式", "desc": "深色背景，夜间更护眼"},
    "brand":  {"icon": "🔵", "label": "品牌蓝",   "desc": "金伍源品牌蓝主题"},
    "mint":   {"icon": "🌿", "label": "护眼绿",   "desc": "柔和薄荷绿，长时间使用护眼"},
    "night":  {"icon": "🌌", "label": "暗夜蓝",   "desc": "深蓝黑底，高对比夜间模式"},
    "sepia":  {"icon": "📜", "label": "暖米白",   "desc": "米黄复古色调，柔和不刺眼"},
    "custom": {"icon": "🎨", "label": "自定义",   "desc": "自选配色与背景图片"},
}
THEME_MODES = tuple(THEME_META.keys())
PRESET_MODES = tuple(m for m in THEME_MODES if m != "custom")
DARK_MODES = ("dark", "night")

# 彩蛋隐藏主题：常规选择入口不列出，需触发暗号后本次临时套用
HIDDEN_MODES = ("brand",)

def _visible_modes(modes):
    return tuple(m for m in modes if m not in HIDDEN_MODES)

FONT_OPTIONS = [
    ("system", "默认\n（系统字体）"),
    ("Microsoft YaHei", "微软雅黑"),
    ("DengXian", "等线"),
    ("SimHei", "黑体"),
    ("SimSun", "宋体"),
    ("KaiTi", "楷体"),
    ("FangSong", "仿宋"),
    ("Consolas", "Consolas（等宽）"),
]


def _font_to_store(v):
    return "" if v in ("", "system") else v


def _font_to_ui(v):
    return "system" if v in ("", "system") else v

# ---------- 自定义主题：配色归一 / 派生调色板 / 背景图路径 ----------
def _norm_color(v, fallback=""):
    """归一为 #RRGGBB；非法则回退 fallback。"""
    if not v:
        return fallback
    s = str(v).strip().lstrip("#")
    if len(s) == 3:
        s = "".join(ch * 2 for ch in s)
    if len(s) == 6 and all(ch in "0123456789abcdefABCDEF" for ch in s):
        return "#" + s.upper()
    return fallback

def _custom_img_path(cfg, key):
    """返回自定义背景图绝对路径；未配置或文件缺失则 None。key: bg/cover/card/dialog。"""
    name = cfg.get(f"custom_{key}_image", "")
    if not name:
        return None
    p = os.path.join(CUSTOM_DIR, os.path.basename(name))
    return p if os.path.exists(p) else None

def _build_custom_palette(cfg):
    """以浅色模板为底，覆盖用户自选的 5 项核心颜色，其余派生 / 沿用模板。"""
    base = dict(PALETTES["light"])
    cp = cfg.get("custom_palette", {}) or {}
    overrides = {}
    for k in ("primary", "bg", "card", "title", "text"):
        col = _norm_color(cp.get(k, ""), "")
        if col:
            overrides[k] = col
    base.update(overrides)
    if "primary" in overrides:
        base["link"] = overrides["primary"]
    if "title" in overrides:
        base["banner"] = overrides["title"]
    if "text" in overrides:
        base["body2"] = overrides["text"]
    if "card" in overrides:
        base["card_alt"] = overrides["card"]
    return base

def _resolve_palette(mode, cfg):
    if mode == "custom":
        return _build_custom_palette(cfg)
    return PALETTES.get(mode, PALETTES["light"])

# 按中版本（主.次）分组的整体总结，各小版本明细收进折叠层
CHANGELOG_SUMMARY = {
    "3.0": "进入 3.0 阶段：新增「产品替换」对照表，按客户原品牌型号逐项对照汇川选型，可从已选搭配带入汇川型号或手填兜底，导出五列替换清单（类型 / 初始型号 / 汇川对应型号 / 数量 / 备注）；2.x 历史版本在更新日志中按阶段折叠归档。",
    "2.0–2.3": "选型基础能力与早期系列录入：纠错与三级版本恢复、线缆查询、手册联网下载、BOM 导出与按日期归档；接入 SV600EN/EP、SV605N/P 经济型系列，并按手册复核 660 档位纠错、更新出厂基线。",
    "2.4–2.7": "主页框架、故障码库建设与外观主题：新增应用主页启动器与导航互通；录齐 12 个系列故障码并接入纠错恢复；扩充至六套预设主题；清洗故障库错位标点与折行排版，补录 SV635N。",
    "2.8–2.12": "自定义外观、数据导入导出与检索体验：自定义配色与背景图背板；选型库 / 故障码库文件双向导入导出；故障码在线纠错；关键词检索、跨系列查码；滚动提速、回到顶部、上次查询、一键回原与使用说明等体验优化。",
    "1.0": "初版桌面选型工具：本地数据库多条件联合筛选、智能推荐与 BOM 导出。",
}

CHANGELOG = [
    ("v3.0.6", [
        "主页新增「其他端版本下载」卡片：弹窗列出安卓全能版、安卓轻量故障版、苹果轻量故障版，点击「打开」在系统浏览器跳转对应落地页下载。",
    ]),
    ("v3.0.5", [
        "修复安卓端免责声明弹窗反馈邮箱在窄屏横向溢出弹窗边界：移动端标签与邮箱改为上下竖排，桌面端仍保持一行横排。",
    ]),
    ("v3.0.4", [
        "修复安卓端选完主题后进入主界面报错白屏：下拉框条目高度 item_height 原为 44，低于打包版 Flet 要求的下限 48.0 触发断言，导致主界面构建中断；已改为 48。",
    ]),
    ("v3.0.3", [
        "修复安卓端首次选完主题后主界面白屏：主题选择改回与已有配置启动一致的同步重建链路，去掉异步延迟任务；根节点与主页容器补满页 expand 约束，避免移动端初始尺寸塌缩；主界面构建异常直接上屏并写入 error.log，便于定位。",
    ]),
    ("v3.0.2", [
        "修复安卓端首次主题选择页卡片区域渲染成灰色占位块、卡片图标与文字不显示的问题：去掉移动端卡片 expand 平分布局，改为单列拉伸铺满宽度，滚动列表限定高度，避免无界滚动约束丢失导致控件未渲染。",
    ]),
    ("v3.0.1", [
        "修复安卓端首次启动选择主题后卡在选择页、无法进入主界面：主题选择改为关闭弹窗后异步延迟重建页面，与设置内切换主题链路一致，避免弹窗回调内同步清空控件树导致移动端刷新丢失。",
        "修复窄屏主题选择卡片单列偏窄、右侧留白及底部卡片被裁切：移动端卡片自适应平分可用宽度，弹窗内容开启滚动。",
    ]),
    ("v3.0.0", [
        "新增「产品替换」选型入口：面向客户用汇川产品替换原品牌产品的场景，按方案分组对照电机、驱动器、动力线、编码线；CN1 接口与电池型号各套相同时合并为共用组、数量合计，不同则归入各自方案组。初始型号列照抄客户铭牌，汇川型号可从勾选搭配一键带入，也可手填兜底。",
        "选型页勾选搭配后可直接「生成替换对照清单」带入替换页；替换清单导出五列对照表（类型 / 初始型号 / 汇川对应型号 / 数量 / 备注），分组标题跨列合并，沿用现有 BOM 的黄底表头、边框列宽、按日期归档与导出记录机制，支持 Excel 与 CSV。",
        "更新日志按中版本折叠归档：2.x 历史版本分组收起，组内保留整体总结，明细收进可展开折叠层。",
        "主页「产品替换选型」卡片图标由交换 emoji 更换为同款小羊捧双向交换图标的自定义图标（repl_icon.png），与选型、故障卡片的小羊形象统一。",
        "替换清单与选型 BOM 分文件夹归档：替换对照表存入独立「替换清单文件夹」并单独记导出记录，可在替换页「打开替换清单」「导出记录」中查阅；替换页同样提供「❓ 使用说明」，首进自动弹出。选型、故障、替换三页说明文案按操作流程重写。",
    ]),
    ("v2.12.5", [
        "选型页、故障查询页首次进入时各弹出一次使用说明，讲解筛选 / 查询操作要点；两页按钮行新增「❓ 使用说明」按钮，可随时回看。",
        "加固窗口最小化恢复：从最小化还原时若窗口坐标仍是屏外残留，自动拉回屏幕中央并刷新一次渲染；修复切换窗口 / 打开文件夹时线程输入附着在异常路径下未释放的残留隐患。",
    ]),
    ("v2.12.4", [
        "选型页右下角新增悬浮「回到顶部」按钮：浏览到列表下方时点击即可平滑滚动回页面顶部，无需长距离上滑；主页与故障查询页不显示该按钮。",
    ]),
    ("v2.12.3", [
        "优化选型结果列表滚动流畅度：结果分批加载，每批 30 条，首屏只渲染当前批次型号卡片，不再一次性铺开全部条目；保留原有页面滚动与滚动条用于判断当前位置，底部「加载更多」可追加后续型号，条目较多时上下滑动更顺滑。",
    ]),
    ("v2.12.2", [
        "清洗 SV680N / SV680P 的 EEE0.0（龙门配对轴故障）机理字段：剔除录入时串入的整份安全模块故障附录，机理恢复为手册原文“龙门配对轴故障”；三列排除表、级别与能否复位保持不变。",
    ]),
    ("v2.12.1", [
        "界面字体选择入口从选型设置弹窗移除，字体统一在封面「外观主题」弹窗中切换；已选字体的持久化与全局生效保持不变。",
    ]),
    ("v2.12.0", [
        "故障码查询新增关键词检索：输入非故障码的现象词（如 过载、编码器断线），在故障名称、机理与原因 / 确认方法 / 处理措施中模糊匹配；系列下拉可选「全部」跨库检索，结果卡片标注所属系列。",
        "系列下拉新增「全部」选项；按码查询在跨系列时同样列出同名 / 同主码条目。",
    ]),
    ("v2.11.1", [
        "故障码上次查询改为本地持久化：最近一次查询的驱动器系列与故障码写入配置，关闭软件后再次打开仍自动回填，可直接查询或点「上次查询」调出。",
        "故障码查询区布局调整：查询、上次查询、纠错记录按钮移到输入框下方单独一行，不再挤占故障码输入框宽度。",
    ]),
    ("v2.11.0", [
        "选型工具新增「↺ 一键回原」：把电压、功率、转速、惯量、法兰、驱动器、通讯、编码器、电池 / 磁编、抱闸、油封、线缆长度与高柔线等条件一次性恢复为默认值，并清空已勾选搭配后重新出结果。",
        "故障码 / 警告码查询新增「🕘 上次查询」按钮：自动记住最近一次查询的驱动器系列与故障码，一键回填并重新检索；尚无记录时按钮置灰。",
    ]),
    ("v2.10.7", [
        "清洗故障机理字段中混入的手册页脚页码与“故障处理”栏目标题残留（如“-542-故障处理”“-607-”），共规整 89 条；故障原因、确认方法、处理措施三列内容保持不变。",
    ]),
    ("v2.10.6", [
        "重制主页封面横幅：改为与小羊图标统一的扁平卡通插画，呈现西装小羊调试伺服电机与驱动器时被电到的幽默场景（眩晕眼、卷毛炸起、黄色闪电与电火花）；画面按窄条横幅做超宽构图、主体收在中部水平带并预留上下安全留白，左侧保留标题留白，替换原写实科技风封面图。",
    ]),
    ("v2.10.5", [
        "新增界面字体选择：设置弹窗与封面外观主题弹窗以同款小方块卡片切换默认 / 微软雅黑 / 等线 / 黑体 / 宋体 / 楷体 / 仿宋 / Consolas（等宽）字体，方块内用对应字体预览名称；选择持久化到本地配置，全局界面随即生效，下次启动保持上次所选。",
    ]),
    ("v2.10.4", [
        "对照 SV680N 手册原文修正 EE17 系列故障原因的 OCR 折行错位：还原被挪动的逗号、顿号位置，补齐英文参数词之间的顿号分隔（EE17.0 六条及 EE17.1 / 17.2 / 17.5 / 17.6，共 10 格）；确认方法、处理措施与其余条目保持不变。",
    ]),
    ("v2.10.3", [
        "清理故障库三列文本末尾残留的顿号、逗号、分号等折行标点（共 365 格），句子中间错位的标点保留不动；查询展示时再做一次尾标点兜底规整，避免后续纠错或导入再次带入尾巴。",
    ]),
    ("v2.10.2", [
        "故障排除表恢复老版单格罗列显示：同一故障码下多条原因合并在一个原因格内逐条换行罗列，确认方法、处理措施整格共用并自动去重；不再把原因逐行拆分、不再显示序号列，原因格内容更饱满，对应关系沿用手册原版排版。",
        "数据库仍保留清洗后的干净逐行结构用于纠错与导入导出，查询展示时按整条故障码合并呈现，两套结构互不影响。",
    ]),
    ("v2.10.1", [
        "清洗故障码 / 警告码库的错位重复数据：去除残留的「●」符号；合并同一故障码内干净版与符号错位版重复成对的原因行；把挤在同一格的多条原因拆回逐行，并继承对应确认方法与处理措施；记录内归一化等价的共用方法 / 措施统一为信息完整的版本。",
        "故障排除表界面加固：每行左侧新增序号列，原因 / 确认方法 / 处理措施三列文字顶部对齐，原因列优先加宽，空单元格以浅灰「—」占位，逐行对应关系更直观，不再因文本长短不一而高低错位、底色断开。",
        "修复排除表在交叉轴拉伸布局下整表高度塌陷、查询结果整片空白的问题：行背景改为行级容器铺底，三列随内容自适应高度，不再依赖交叉轴拉伸。",
        "少数手册原文为并列结构、库内方法 / 措施无法唯一对应的条目暂保留空缺，待后续结合手册人工核对补全。",
    ]),
    ("v2.10.0", [
        "选型库与故障码库纠错记录弹窗新增「📤 导出文件」「📥 导入文件」：可把当前数据库导出为文件离线批量修改文本，改完再导回软件整库覆盖，方便后续维护文本版本。",
        "选型库直接导出 / 导入原 xlsx；故障码库导出为中文表头扁平表格（系列、故障码、名称、级别、能否复位、机理、原因、确认方法、措施），导入时按系列 + 故障码自动聚合还原排除表结构。",
        "整库导入前自动保存一条快照并记入纠错台账，可在记录里撤销本次导入、恢复到导入前；导入文件缺表或缺关键列会校验拦截，不会覆盖当前库。",
    ]),
    ("v2.9.0", [
        "故障码 / 警告码查询接入纠错能力：在任意故障码卡片上点击「✏️ 纠错」，可修改故障名称、级别、能否复位、故障机理与三列故障排除表，并可填写纠错备注。",
        "故障码纠错记录与版本恢复与选型工具一致：纠错前自动保存整库快照，台账逐条记录系列、故障码、字段、改前改后值与备注；支持撤销最近一次纠错、任选历史快照恢复、单条删除记录。",
        "新增恢复出厂原始故障库（二次确认撤销全部纠错）与清空全部纠错记录（二次确认，不回滚当前故障库、不动出厂基线）；纠错数据落到程序目录可写副本，内置只读资源不受影响。",
    ]),
    ("v2.8.2", [
        "品牌蓝主题封面横幅左上角的金伍源 JWY logo 支持点击：点击后在系统默认浏览器打开公司官网（https://sz-jwy.com/），鼠标悬停显示「访问金伍源官网」提示。",
    ]),
    ("v2.8.1", [
        "修复保存自定义外观或切换主题后界面卡死、按钮点不动的问题：弹窗关闭瞬间整树重建会让旧弹窗路由的灰色遮罩残留在窗口上并拦截点击；改为先关闭弹窗、等待弹窗路由回收完成后再重建界面，预设主题与自定义主题切换、保存应用、恢复默认各链路均不再残留遮罩。",
    ]),
    ("v2.8.0", [
        "新增「自定义」外观主题：在六套预设之外追加一套可自定义入口；核心主色、背景色、卡片色、标题色、正文色五项可自选，提供常用色块点选与十六进制色值输入，边框、分隔线、链接、横幅等其余颜色由所选配色自动派生。",
        "支持导入照片作为四类背板：整窗铺底背景、主页封面横幅、卡片背板、弹窗背板；图片复制到程序目录 custom_assets 下随包便携，不存绝对路径，可单独移除或一键恢复默认；背板做了淡化处理以保证文字可读。",
    ]),
    ("v2.7.3", [
        "故障排除表的多条确认方法 / 处理措施按编号拆分为逐行显示：原粘连在同一段的「1. …… 2. …… 3. ……」顺序编号项各自换行，编号统一为「N. 」格式；功能码 / 位号、电压数值与取值说明（如 0-/1-/2-）不受影响。",
    ]),
    ("v2.7.2", [
        "清洗故障码 / 警告码库中 PDF 折行抽取残留的错位标点：去除词中间被折行拆开后误插入的多余逗号、句号与句点，折叠 `，，`、`、、`、`，.`、`。.` 等标点堆叠；正常分句的逗号、顿号、句号以及功能码 / 位号内的句点（如 H01.72、607D01h）均保留不动。",
    ]),
    ("v2.7.1", [
        "故障码 / 警告码查询补充 SV635N 手册数据：按用户手册一览表与详细说明共录入 107 条（104 条警告 / 故障码 + 3 条内部故障），每条含名称、级别、能否复位、故障机理与三列排除表（原因 / 确认方法 / 处理措施）；覆盖过流过压、过载过热、编码器、STO 安全状态、飞车与通信类码型。",
    ]),
    ("v2.7.0", [
        "外观主题扩充至六套：新增「护眼绿」「暗夜蓝」「暖米白」三套预设；主题选择弹窗、首次启动引导与设置下拉统一按主题清单生成，卡片两行排列。",
    ]),
    ("v2.6.2", [
        "新增「品牌蓝」外观主题：按金伍源 logo / 招牌采样品牌深蓝与招牌蓝为主色、品牌橙作点缀；封面顶部为带 JWY 品牌标识的横幅卡片，配橙色装饰线；封面主题弹窗与选型设置下拉均可切换。",
    ]),
    ("v2.6.1", [
        "封面主页右下角新增「🎨 外观主题」独立入口，浅色 / 深色一张卡片即可切换，不再只藏在选型页设置弹窗里。",
        "主题切换弹窗的两张卡片收窄为按内容自适应的小卡片，不再纵向撑满整个弹窗。",
        "修复选型设置弹窗最上方两个下拉（默认供电电压、默认线缆长度）标题文字被弹窗上沿裁掉一半的问题。",
    ]),
    ("v2.6.0", [
        "新增浅色 / 深色外观主题：首次启动在免责声明之前先选择浅色或深色模式，选择结果自动保存；之后可在选型页设置弹窗中随时切换，切换后界面即时重建。",
        "适配深色模式整套配色：弹窗、卡片、文字、边框、分隔线、表格、输入框、标签与状态色均按深色调色板显示；修复系统深色模式下免责声明深底深字、几乎不可读的问题。",
    ]),
    ("v2.5.12", [
        "主页「伺服选型工具」卡片图标更换为抱伺服驱动器的小羊图标（select_icon.png）；全局 logo 与各页页眉图标保持不变。",
    ]),
    ("v2.5.11", [
        "主页故障码卡片图标更换为无水印的小羊警告图标（替换上一版带水印素材）。",
    ]),
    ("v2.5.10", [
        "主页「故障码 / 警告码查询」卡片图标由警告 emoji 更换为自定义小羊警告图标（fault_icon.png）。",
    ]),
    ("v2.5.9", [
        "更新日志改在应用主页（封面）查看，选型工具内不再放置日志入口。",
        "日志按中版本（如 v2.5）折叠分组：组内给出整体总结，各小版本明细收进可展开的折叠层。",
        "仅最新版本标注更新日期，历史版本不再逐条显示日期。",
    ]),
    ("v2.5.8", [
        "修复故障码查询「故障排除」表长文本被截断的问题：原 DataTable 行高不随多行内容撑开，确认方法 / 处理措施列折行文字被裁；改为自适应行高的三列布局，长文本完整换行显示。",
    ]),
    ("v2.5.7", [
        "手册下载列表补充 SV660N/P、SV670N/P、SV680N/P、SV600EN/EP 共 8 条用户手册直链，与原有选型手册并存；SV605N、SV605P/C 两条已存在，按重复跳过。",
    ]),
    ("v2.5.6", [
        "故障码查询支持主码前缀匹配：只输主码（如 E108）即可列出该码下全部子码（E108.0 / E108.1 …）；输入完整码仍精确匹配单条。",
    ]),
    ("v2.5.5", [
        "故障码 / 警告码查询补充 SV600EN、SV600EP 手册数据：共录入 SV600EN 129 条、SV600EP 129 条，含名称、级别、能否复位、故障机理与三列排除表；EN 型含 EtherCAT 总线码，EP 型含 CANlink 与脉冲型码。至此 12 个系列码表全部录入完成。",
    ]),
    ("v2.5.4", [
        "故障码 / 警告码查询补充 SV605N、SV605P 手册数据：共录入 SV605N 129 条、SV605P 129 条，含名称、级别、能否复位、故障机理与三列排除表；N 型含 EtherCAT 总线码，P 型（P/C 合册，整册录入）含脉冲与 CANopen/CANlink 总线码。",
    ]),
    ("v2.5.3", [
        "故障码 / 警告码查询补充 SV680N、SV680P 手册数据：共录入 SV680N 136 条、SV680P 158 条，含名称、级别、能否复位、故障机理与三列排除表；含 FSoE 安全通信码（EE17.x）与龙门配对轴故障（EEE0.0）等码型。",
    ]),
    ("v2.5.2", [
        "故障码 / 警告码查询补充 SV670N、SV670P 手册数据：共录入 SV670N 179 条、SV670P 162 条，含名称、级别、能否复位、故障机理与三列排除表；其中 EA41.0（转矩波动补偿失败）、E956.0（工艺段正向超程报警）、E958.0（工艺段反向超程报警）手册未附机理 / 排除表，按原文留空。",
    ]),
    ("v2.5.1", [
        "故障码 / 警告码查询补充 SV660N、SV660P 手册数据：共录入 SV660N 91 条、SV660P 90 条，含名称、级别、能否复位、故障机理与三列排除表；其中 E994.0（CANlink站号冲突）、ED03.0（CANlink掉线）手册未附机理 / 排除表，按原文留空。",
    ]),
    ("v2.5.0", [
        "故障码 / 警告码查询接入 SV630N、SV630P 手册数据：解析用户手册一览表与详细说明，共录入 SV630N 142 条、SV630P 89 条，每条含故障名称、级别、能否复位、故障机理与三列故障排除表（原因 / 确认方法 / 处理措施）。",
        "故障码库外置为 fault_db.json，随程序目录打包；其他系列（SV660/670/680、SV600EN/EP、SV605N/P）占位待后续按手册补录。",
    ]),
    ("v2.4.0", [
        "新增应用主页（启动器）：打开软件先进入工具集主页，封面横幅 + 应用卡片；当前收录「伺服选型工具」，点击卡片或「进入」即进入现有选型界面，以后可在主页继续追加其他功能入口。",
        "选型界面顶部新增「← 主页」返回按钮，可随时回到应用主页；主页与选型页切换时保留下拉条件、勾选与结果状态，不重新加载。",
        "新增主页封面图 cover.jpg（工业蓝伺服主题横幅），随程序目录打包。",
    ]),
    ("v2.3.3", [
        "数据清理：按 SV660 标准编码器选型一览表复核全部光编（A 后缀）机型，280 条档位与手册一致；删除 1 条两头不靠的冗余错录 MS1H1-15C30CD-A331R（N / P 共 2 条，380V 误配 220V 档 S012）；同型 220V 机型 MS1H1-15C30CB 与 380V 机型 MS1H2-15C30CD 正确记录均已在库，未受影响。",
        "出厂基线同步更新为清理后的当前数据库。",
    ]),
    ("v2.3.2", [
        "数据纠错：按 SV660 选型一览表，修正 0.2kW 免电池编码器机型（MS1H1-20B / MS1H4-20B，N / P 各两型共 8 条）推荐驱动器由 S2R8I 改为 S1R6I，配套手册编号更正为 00002；0.4kW（40B）档位维持 S2R8I 不变。",
        "出厂基线更新：将当前数据库（含 600EN/EP、605N/P 及本次纠错）另存为新的出厂原始基线，「恢复出厂原始」不再回退掉新系列数据；旧基线已留存为 servo_data.original.legacy.xlsx。",
    ]),
    ("v2.3.1", [
        "新增 SV605P 经济型脉冲伺服系列：与 SV605N 同平台，录入 MSAH3 / MSAH4 共 10 款机型、20 条搭配（非抱闸 / 带抱闸各一），覆盖 0.1～1.8kW、AC220 / AC380、脉冲指令通讯。",
        "脉冲型 CN1 端子为 DB44（S6-C8），电池套件 S6-C4A，编码器沿用 T2（18 位多圈绝对值），电机后缀 T231R / T234R。",
        "配套线缆与总线型一致：H3 动力线 M111/B111、编码线 P12L；H4 前出线动力线 M107/B107、编码线 P124，后出线线号在备注中注明。",
        "手册列表新增 SV605P/C 系列伺服用户手册联网下载。",
    ]),
    ("v2.3.0", [
        "新增 SV605N 经济型总线伺服系列：录入 MSAH3 / MSAH4 共 10 款机型、20 条搭配（非抱闸 / 带抱闸各一），覆盖 0.1～1.8kW、AC220 / AC380、EtherCAT 通讯。",
        "编码器下拉新增 T2（18 位多圈绝对值）；605N 电机后缀 T231R / T234R，CN1 端子 DB15（S6-C6），电池套件 S6-C4A。",
        "配套线缆：H3 动力线 M111/B111、编码线 P12L；H4 前出线动力线 M107/B107、编码线 P124，后出线线号在备注中注明。",
        "手册列表新增 SV605N 系列伺服用户手册联网下载。",
    ]),
    ("v2.2.1", [
        "新增 SV600EP 经济型脉冲伺服系列：与 SV600EN 同平台，录入 MSAH3 / MSAH4 共 10 款机型、20 条搭配（非抱闸 / 带抱闸各一），覆盖 0.1～1.8kW、AC220 / AC380、脉冲指令通讯。",
        "驱动器系列下拉、通讯与编码器联动预设同步接入 SV600EP；脉冲型 CN1 端子为 DB44（S6-C8），电池套件 S6-C4A，编码器 C2（18 位多圈绝对值）。",
        "配套线缆与总线型一致：H3 动力线 M111/B111、编码线 P12L；H4 前出线动力线 M107/B107、编码线 P124，后出线线号在备注中注明。",
        "手册列表新增 SV600EP 系列伺服选型手册联网下载。",
    ]),
    ("v2.2.0", [
        "新增 SV600EN 经济型总线伺服系列：录入 MSAH3 / MSAH4 共 10 款机型、20 条搭配（非抱闸 / 带抱闸各一），覆盖 0.1～1.8kW、AC220 / AC380、EtherCAT 通讯。",
        "驱动器系列下拉、通讯与编码器联动预设同步接入 SV600EN；编码器新增 C2（18 位多圈绝对值）选项。",
        "配套线缆按手册分档：H3 动力线 M111/B111、编码线 P12L；H4 前出线动力线 M107/B107、编码线 P124，后出线线号在备注中注明。",
        "手册列表新增 SV600EN 系列伺服选型手册联网下载；驱动器系列匹配由固定 6 位前缀改为前缀匹配，兼容 SV600EN 等 7 字符系列标识。",
    ]),
    ("v2.1.2", [
        "结果区新增「✖ 取消勾选」按钮，一键清空所有已勾选的搭配。",
        "导出清单按日期归档：自动存入「清单文件夹/YYYY-MM-DD/」子文件夹，导出记录显示所属日期目录。",
    ]),
    ("v2.1.1", [
        "水印署名固定为「姓许的苏剑」，设置中不再提供署名修改入口。",
        "多套搭配导出改为一套一张独立 BOM 小表，纵向排列在同一文件内，不再按同型号合并累加数量。",
        "两套及以上搭配自动比对差异，在每套小表下方注明与其他套取值不同的字段（抱闸、动力线、电机型号等）。",
        "导出记录支持单条删除，以及清空全部记录（二次确认）。",
    ]),
    ("v2.1.0", [
        "空结果诊断：无匹配时逐步筛选定位把结果筛没的瓶颈条件，给出相邻功率/法兰档位与「放宽为全部」等建议。",
        "多套搭配合并 BOM：结果卡片可勾选，一键把勾选搭配汇总到一份清单，同型号数量累加、备注合并去重。",
        "BOM 模板可配置：设置内可调导出格式（Excel/CSV）、清单标题、公司抬头、落款与导出列（类型/型号/数量/备注）。",
        "导出记录：每次导出记录时间、文件、格式、套数与选型条件，弹窗可查阅。",
        "CSV 格式可选：除 xlsx 外支持导出 UTF-8（带 BOM）CSV，便于导入 ERP。",
        "异常兜底与运行日志：数据读不出、文件占用、下载/导出/纠错失败均给出明确提示，并落 logs/运行日志.log。",
        "只读保护加固：内置数据库只读，纠错与导出落在程序目录可写副本；基线损坏时自动兜底，启动不崩溃。",
        "设置项独立保存：默认电压/线长/高柔、水印署名、BOM 模板与数据版本写入 settings.json，不写死在代码。",
        "手册缓存说明：已下载手册在列表显示本地大小与修改时间。",
        "数据版本标记：页脚显示数据依据的手册版本与日期（可在设置填写）。",
    ]),
    ("v2.0.4", [
        "纠错功能：结果卡片新增「✏️ 纠错」，可直接修改驱动器、动力线、编码线、CN1、电池、通讯、编码器、抱闸、油封、法兰、备注等业务字段并写回表格。",
        "每次纠错前自动保存整表快照，并在「纠错台账」记录序号、时间、电机、驱动器、字段、改前值、改后值与备注。",
        "三级版本恢复：撤销最近一次、任选历史快照恢复到对应时点、一键恢复出厂原始（需二次确认）；出厂基线首次运行自动建立，不被纠错覆盖。",
        "纠错记录支持单条删除（无其他记录引用的快照一并清理）与清空全部记录（二次确认，不回滚当前表格、不影响出厂基线）。",
    ]),
    ("v2.0.3", [
        "线缆型号查询：输入电机型号（忽略大小写），选择长度与是否高柔，查出对应动力线、编码线；跨驱动器系列分行列出。",
        "选型手册联网下载：内置 10 份官方手册直链，一键下载到本地「选型手册」文件夹，下载时显示进度条与百分比，已下载可直接打开，失败可重试。",
        "已下载的手册支持一键移除本地文件，需要时可重新下载。",
        "抱闸由型号后缀自动判定，线缆查询移除抱闸筛选，改为结果备注显示。",
        "修复手册、清单文件夹打开后窗口未自动置顶的问题。",
    ]),
    ("v2.0.2", [
        "新增「打开清单文件夹」按钮；BOM 统一存入「清单文件夹」，同名自动加序号不覆盖。",
        "新增选型手册列表弹窗，可直接打开本地已有手册。",
        "惯量 & 容量由复选框改为下拉选项，节省界面空间。",
        "修复页面无法滚动、结果区被挤压的布局问题；界面卡片化美化，控件宽度统一。",
        "修复弹窗残留、偶发灰屏及控件未加入页面即更新的报错。",
    ]),
    ("v2.0.1", [
        "启动免责声明弹窗：个人学习交流、完全免费、禁止售卖转售，附反馈邮箱。",
        "接入伺服数据库，按电压/功率/转速/惯量/法兰/驱动器/通讯/编码器/抱闸/油封等条件联合匹配。",
        "智能推荐：无完全符合型号时，自动给出最接近的搭配并标注差异。",
        "BOM 一键导出 Excel 采购清单（类型/型号/数量/备注）。",
        "页眉自动加载 logo，界面署名水印「姓许的苏剑」。",
        "补录 SV670N / SV670P 系列；670 油封口径统一为「有油封」。",
    ]),
    ("v1.0（早期版本）", [
        "初版桌面选型工具：Python + Flet 0.22.0，本地 openpyxl 读写数据库（未采用被策略阻止的 pandas）。",
        "数据表按铭牌段位建字段：惯量/容量/功率/转速/电压/编码器/轴码/抱闸油封/通讯/型号/驱动器/线缆/CN1/电池。",
        "多条件联合筛选：电压、功率、转速、惯量、法兰、驱动器、通讯、编码器、免电池/光编、抱闸、油封、高柔线。",
        "逐步录入 SV630N/P、SV660N/P、SV680N/P；动力线 M=无抱闸、B=带抱闸，编码线按惯量族分档；CN1 总线 S6-C6、脉冲 S6-C8。",
        "初版 BOM 导出：类型/型号/数量/备注，表头黄底加粗、细边框、列宽适配，重名自动加序号。",
        "编码器短码纠正为 A6/S6/A3/T3/U2，后缀第二位 3 为轴码；区分光编/免电池磁编及是否配电池 S6-C4A。",
        "联动修正：切换通讯/编码器/电池不再清空驱动器系列；630N 全系 18bit；筛空自动回落全部。",
        "电机系列定名 MS1-R（铭牌订货型号仍写作 MS1Hx 开头）。",
        "缺数据库文件自动建空表兜底；修复 Flet 0.22 页面宽度赋值报错。",
    ]),
]


def _unique_path(base_dir, name):
    """同名文件已存在时自动加序号，避免覆盖上一次导出的清单"""
    path = os.path.join(base_dir, name)
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(name)
    i = 2
    while os.path.exists(os.path.join(base_dir, f"{stem}({i}){ext}")):
        i += 1
    return os.path.join(base_dir, f"{stem}({i}){ext}")


def _item_bom_rows(it, depth, high_flex):
    """单条搭配展开为 [类型,型号,数量,备注] 行列表"""
    motor = str(it.get("motor_model") or "")
    drive = str(it.get("drive_model") or "")

    kw = 0.0
    try:
        kw = float("".join(c for c in str(it.get("power_text") or "") if c.isdigit() or c == "."))
    except Exception:
        kw = 0.0
    watt = f"{kw * 1000:g}W"

    enc_type = str(it.get("encoder_type") or "")
    enc_bits = ""
    for part in enc_type.replace("位", "bit").split():
        if part.lower().endswith("bit"):
            enc_bits = part.lower().replace("bit", "位")
            break
    enc_note = f"{enc_bits}编码器" if enc_bits else (str(it.get("encoder_short") or "") + "编码器")

    brake = "带抱闸" if str(it.get("has_brake")) == "是" else "无抱闸"
    flange = str(it.get("flange_size") or "")
    frame = flange.split("×")[0].replace("mm", "").strip()
    frame_note = f"{frame}基座" if frame else "基座待补"

    comm = str(it.get("comm") or "")
    drive_note = "总线型" if comm == "EtherCAT" else "脉冲型"

    power_cable = build_line(it.get("power_line_base"), depth, high_flex)
    encoder_cable = build_line(it.get("encoder_line_base"), depth, high_flex)
    try:
        depth_txt = f"{float(depth):g}"
    except Exception:
        depth_txt = str(depth)
    cable_note = f"{depth_txt}米， {'高柔' if high_flex else '普通'}"
    cn1 = str(it.get("cn1") or "")
    battery = str(it.get("battery") or "")

    rows = [
        ["伺服电机", motor, 1, f"{watt}， {enc_note}， {brake}， {frame_note}"],
        ["伺服驱动器", drive, 1, drive_note],
        ["动力线型号", power_cable, 1, f"{cable_note}， {brake}"],
        ["编码器型号", encoder_cable, 1, cable_note],
        ["CN1接口", cn1, 1, "正负限， io接口"],
        ["绝对值电池", battery, 1, "绝对值模式选配"],
    ]
    return [r for r in rows if str(r[1]).strip() not in ("", "None", "nan")]


def _compare_fields(it):
    brake = "带抱闸" if str(it.get("has_brake")) == "是" else "无抱闸"
    oil = "有油封" if str(it.get("has_oilseal")) == "是" else "无油封"
    return {
        "电机型号": str(it.get("motor_model") or "待补"),
        "驱动器": str(it.get("drive_model") or "待补"),
        "功率": str(it.get("power_text") or "待补"),
        "编码器": str(it.get("encoder_short") or "待补"),
        "抱闸": brake,
        "油封": oil,
        "法兰": str(it.get("flange_size") or "待补"),
        "通讯": str(it.get("comm") or "待补"),
        "动力线": str(it.get("power_line_base") or "待补"),
        "编码线": str(it.get("encoder_line_base") or "待补"),
        "电池": str(it.get("battery") or "待补"),
    }


def _diff_notes(items):
    """多套搭配时，逐套生成与其他套取值不同的字段说明"""
    if len(items) < 2:
        return [""] * len(items)
    fields = [_compare_fields(it) for it in items]
    notes = []
    for i, f in enumerate(fields):
        diffs = []
        for k, v in f.items():
            others = [fields[j][k] for j in range(len(fields)) if j != i]
            if any(o != v for o in others):
                other = others[0] if len(others) == 1 else "其他套"
                diffs.append(f"{k}：{v}（{other}）")
        notes.append("；".join(diffs))
    return notes


def _active_columns(cfg):
    cols = []
    if cfg.get("col_type", True):
        cols.append(("type", "类型"))
    if cfg.get("col_model", True):
        cols.append(("model", "型号"))
    if cfg.get("col_qty", True):
        cols.append(("qty", "数量"))
    if cfg.get("col_note", True):
        cols.append(("note", "备注"))
    return cols


def _row_values(row, cols):
    mapping = {"type": row[0], "model": row[1], "qty": row[2], "note": row[3]}
    return [mapping[k] for k, _ in cols]


def _write_bom_xlsx(path, groups, cfg, cols, conditions_txt):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BOM清单"
    head_fill = PatternFill("solid", start_color="FFFFFF00", end_color="FFFFFF00")
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    thin = Side(style="thin", color="FF000000")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    ncol = max(len(cols), 1)

    def _merge_line(row, text, font, align=center):
        ws.cell(row, 1, text)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncol)
        ws.cell(row, 1).font = font
        ws.cell(row, 1).alignment = align

    r = 1
    _merge_line(r, cfg.get("bom_title") or "伺服电机与驱动器配置清单", Font(bold=True, size=14))
    r += 1
    if cfg.get("bom_company"):
        _merge_line(r, cfg["bom_company"], Font(size=11), left)
        r += 1
    if conditions_txt:
        _merge_line(r, conditions_txt, Font(size=9, color="808080"), left)
        r += 1

    for gi, g in enumerate(groups, start=1):
        _merge_line(r, f"搭配 {gi}：{g['title']}", Font(bold=True, size=12, color="123A5A"), left)
        r += 1
        for ci, (_, title) in enumerate(cols, start=1):
            c = ws.cell(r, ci, title)
            c.font = Font(bold=True, size=12)
            c.fill = head_fill
            c.alignment = center
            c.border = box
        r += 1
        for line in g["rows"]:
            for ci, val in enumerate(_row_values(line, cols), start=1):
                c = ws.cell(r, ci, val)
                c.alignment = center
                c.border = box
            r += 1
        if g["diff"]:
            _merge_line(r, "区别：" + g["diff"], Font(size=9, color="B26A00"), left)
            r += 1
        r += 1

    if cfg.get("bom_footer"):
        ws.cell(r, 1, cfg["bom_footer"]).font = Font(size=10)

    widths = {"type": 16, "model": 30, "qty": 8, "note": 44}
    for ci, (k, _) in enumerate(cols, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = widths.get(k, 18)
    wb.save(path)


def _write_bom_csv(path, groups, cfg, cols, conditions_txt):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([cfg.get("bom_title") or "伺服电机与驱动器配置清单"])
        if cfg.get("bom_company"):
            w.writerow([cfg["bom_company"]])
        if conditions_txt:
            w.writerow([conditions_txt])
        for gi, g in enumerate(groups, start=1):
            w.writerow([f"搭配 {gi}：{g['title']}"])
            w.writerow([t for _, t in cols])
            for line in g["rows"]:
                w.writerow(_row_values(line, cols))
            if g["diff"]:
                w.writerow(["区别：" + g["diff"]])
            w.writerow([])
        if cfg.get("bom_footer"):
            w.writerow([cfg["bom_footer"]])


# ---------- 产品替换对照表导出（五列：类型 / 初始型号 / 汇川对应型号 / 数量 / 备注） ----------
REPL_HEADERS = ["类型", "初始型号", "汇川对应型号", "数量", "备注"]
REPL_BODY_TYPES = ["伺服电机", "伺服驱动器", "动力线型号", "编码器型号"]
REPL_ACC_TYPES = ["CN1接口", "绝对值电池"]


def _write_repl_xlsx(path, groups, cfg, conditions_txt):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "替换清单"
    head_fill = PatternFill("solid", start_color="FFFFFF00", end_color="FFFFFF00")
    group_fill = PatternFill("solid", start_color="FFFFF2CC", end_color="FFFFF2CC")
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    thin = Side(style="thin", color="FF000000")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    ncol = len(REPL_HEADERS)

    def _merge_line(row, text, font, fill=None, align=center):
        ws.cell(row, 1, text)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncol)
        c = ws.cell(row, 1)
        c.font = font
        c.alignment = align
        if fill:
            c.fill = fill

    r = 1
    _merge_line(r, cfg.get("bom_title") or "伺服电机与驱动器配置清单", Font(bold=True, size=14))
    r += 1
    if cfg.get("bom_company"):
        _merge_line(r, cfg["bom_company"], Font(size=11), align=left)
        r += 1
    if conditions_txt:
        _merge_line(r, conditions_txt, Font(size=9, color="808080"), align=left)
        r += 1
    for ci, h in enumerate(REPL_HEADERS, start=1):
        c = ws.cell(r, ci, h)
        c.font = Font(bold=True, size=12)
        c.fill = head_fill
        c.alignment = center
        c.border = box
    r += 1
    for g in groups:
        if g.get("title"):
            _merge_line(r, g["title"], Font(bold=True, size=12), group_fill, left)
            r += 1
        for line in g["rows"]:
            for ci, val in enumerate(line, start=1):
                c = ws.cell(r, ci, val)
                c.alignment = center if ci in (1, 4) else left
                c.border = box
            r += 1
    if cfg.get("bom_footer"):
        ws.cell(r, 1, cfg["bom_footer"]).font = Font(size=10)
    widths = [16, 26, 26, 8, 40]
    for ci, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = w
    wb.save(path)


def _write_repl_csv(path, groups, cfg, conditions_txt):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([cfg.get("bom_title") or "伺服电机与驱动器配置清单"])
        if cfg.get("bom_company"):
            w.writerow([cfg["bom_company"]])
        if conditions_txt:
            w.writerow([conditions_txt])
        w.writerow(REPL_HEADERS)
        for g in groups:
            if g.get("title"):
                w.writerow([g["title"]])
            for line in g["rows"]:
                w.writerow(line)
        if cfg.get("bom_footer"):
            w.writerow([cfg["bom_footer"]])


def export_repl(groups, cfg, conditions_txt=""):
    fmt = (cfg.get("bom_format") or "xlsx").lower()
    day = datetime.now().strftime("%Y-%m-%d")
    day_dir = os.path.join(REPL_EXPORT_DIR, day)
    os.makedirs(day_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y.%m.%d_%H%M%S")
    ext = "csv" if fmt == "csv" else "xlsx"
    path = _unique_path(day_dir, f"替换清单_{stamp}.{ext}")
    if fmt == "csv":
        _write_repl_csv(path, groups, cfg, conditions_txt)
    else:
        _write_repl_xlsx(path, groups, cfg, conditions_txt)
    n = sum(len(g["rows"]) for g in groups)
    _append_export_log(path, fmt, n, conditions_txt or "产品替换对照表",
                       base_dir=REPL_EXPORT_DIR, log_path=REPL_EXPORT_LOG)
    _log(f"导出替换清单：{os.path.basename(path)}（{n} 行）")
    return path, n


def _read_export_log(log_path=None):
    log_path = log_path or EXPORT_LOG
    try:
        if os.path.exists(log_path):
            with open(log_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, list) else []
    except Exception as ex:
        _log(f"导出记录读取失败：{ex}", "ERROR")
    return []


def _write_export_log(entries, log_path=None):
    log_path = log_path or EXPORT_LOG
    os.makedirs(os.path.dirname(log_path), exist_ok=True)
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(entries[-200:], f, ensure_ascii=False, indent=2)


def _append_export_log(path, fmt, count, conditions_txt, base_dir=None, log_path=None):
    base_dir = base_dir or EXPORT_DIR
    log_path = log_path or EXPORT_LOG
    try:
        entries = _read_export_log(log_path)
        try:
            rel = os.path.relpath(os.path.dirname(path), base_dir)
            sub = "" if rel == "." else rel.replace(os.sep, "/")
        except Exception:
            sub = ""
        entries.append({
            "时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "目录": sub,
            "文件": os.path.basename(path),
            "格式": fmt,
            "套数": count,
            "条件": conditions_txt,
        })
        _write_export_log(entries, log_path)
    except Exception as ex:
        _log(f"导出记录写入失败：{ex}", "ERROR")


def export_bom(items, depth, high_flex, cfg, base_dir, conditions_txt="", name_hint=None):
    """items: 单条 dict 或 dict 列表；返回 (文件绝对路径, 套数)"""
    if isinstance(items, dict):
        items = [items]
    diffs = _diff_notes(items)
    groups = []
    for i, it in enumerate(items):
        title = str(it.get("motor_model") or f"搭配{i + 1}")
        groups.append({
            "title": title,
            "rows": _item_bom_rows(it, depth, high_flex),
            "diff": diffs[i],
        })
    cols = _active_columns(cfg)
    fmt = (cfg.get("bom_format") or "xlsx").lower()
    day = datetime.now().strftime("%Y-%m-%d")
    day_dir = os.path.join(EXPORT_DIR, day)
    os.makedirs(day_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y.%m.%d_%H%M%S")
    stem = f"选型BOM_{name_hint}" if name_hint else f"选型BOM_{stamp}"
    ext = "csv" if fmt == "csv" else "xlsx"
    path = _unique_path(day_dir, f"{stem}.{ext}")
    if fmt == "csv":
        _write_bom_csv(path, groups, cfg, cols, conditions_txt)
    else:
        _write_bom_xlsx(path, groups, cfg, cols, conditions_txt)
    _append_export_log(path, fmt, len(items), conditions_txt)
    total_rows = sum(len(g["rows"]) for g in groups)
    _log(f"导出 BOM：{os.path.basename(path)}（{len(items)} 套，{total_rows} 行）")
    return path, len(items)


def create_empty_excel(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "data"
    ws.append(HEADER)
    wb.save(path)

def load_excel(path):
    if not os.path.exists(path):
        create_empty_excel(path)
    try:
        wb = openpyxl.load_workbook(path)
        ws = wb["data"]
        rows = list(ws.iter_rows(values_only=True))
        header = rows[0]
        data = []
        for r in rows[1:]:
            if not r or not any(r):
                continue
            data.append(dict(zip(header, r)))
        return data
    except Exception as ex:
        _log(f"数据库读取失败，启用空表兜底：{ex}", "ERROR")
        create_empty_excel(path)
        return []

def _app(page, cfg, data, base_dir, C, mode, rebuild=None, show_license=True):
    web = page.web
    is_mobile = _is_mobile(page)

    def _w(desktop_width):
        """窄屏不写死宽度、跟随父容器拉伸；桌面/网页保留固定宽"""
        return None if is_mobile else desktop_width

    def _asset(name):
        return name if web else os.path.join(RES_DIR, name)

    # ---------- 圆润可爱风统一样式工厂：配色仍取自当前主题，只统一圆角 / 柔底 / 细边框 ----------
    def _dd(**kw):
        kw.setdefault("fill_color", C["card_alt"])
        kw.setdefault("bgcolor", C["card"])
        kw.setdefault("border_color", C["border"])
        kw.setdefault("focused_border_color", C["primary"])
        kw.setdefault("border_radius", R_FIELD)
        kw.setdefault("icon_enabled_color", C["muted"])
        kw.setdefault("content_padding", ft.padding.symmetric(horizontal=14, vertical=10))
        kw.setdefault("item_height", 48)
        return ft.Dropdown(**kw)

    def _tf(**kw):
        kw.setdefault("bgcolor", C["card_alt"])
        kw.setdefault("border_color", C["border"])
        kw.setdefault("focused_border_color", C["primary"])
        kw.setdefault("border_radius", R_FIELD)
        kw.setdefault("content_padding", ft.padding.symmetric(horizontal=14, vertical=10))
        return ft.TextField(**kw)

    def _btn_shape():
        return ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=R_BTN))

    def _eb(text=None, **kw):
        if text is not None:
            kw.setdefault("text", text)
        kw.setdefault("style", _btn_shape())
        return ft.ElevatedButton(**kw)

    def _tb(text=None, **kw):
        if text is not None:
            kw.setdefault("text", text)
        kw.setdefault("style", _btn_shape())
        return ft.TextButton(**kw)

    def _cb(**kw):
        kw.setdefault("shape", ft.RoundedRectangleBorder(radius=6))
        return ft.Checkbox(**kw)

    # ---------- 启动免责声明弹窗 ----------
    def _close_license(e):
        license_dlg.open = False
        page.update()

    license_dlg = ft.AlertDialog(
        open=True,
        title=ft.Text("使用须知 · 免责声明", size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
        content=ft.Column([
            ft.Text(
                "本工具为个人学习、设计交流用途，完全免费；严禁任何形式的商业售卖、"
                "二次转售、付费打包或未经授权的公开分发。",
                size=12.5, color=C["text"]),
            ft.Text(
                "工具给出的电机、驱动器、动力线、编码线、接头及电池等配套型号，以及"
                "故障码 / 警告码对应的原因、确认方法与处理措施，仅作选型与排查参考；"
                "正式采购、接线、调试、投产及现场排障前，请以汇川官方最新手册资料及"
                "现场技术人员确认为准。",
                size=12.5, color=C["text"]),
            ft.Text(
                "数据依据公开手册人工整理，产品可能随版本升级调整，型号、参数与码表"
                "如有出入请以官方资料为准；导出的 BOM 清单、线缆与故障记录仅供内部"
                "核对，不构成采购或工程承诺；下载的手册版权归原厂所有，仅供学习参考。",
                size=12.5, color=C["text"]),
            ft.Text(
                "因使用本工具选型、查询或导出内容所引发的任何直接或间接损失，作者不"
                "承担相应责任；如发现型号或码表错漏，欢迎通过下方邮箱反馈勘误。",
                size=12.5, color=C["text"]),
            (
                ft.Row([
                    ft.Text("问题反馈 / 型号勘误邮箱：", size=12.5, color=C["text"]),
                    ft.Text("1220313772@qq.com", size=12.5, weight=ft.FontWeight.BOLD, color=C["link"]),
                ], spacing=4)
                if not is_mobile else
                ft.Column([
                    ft.Text("问题反馈 / 型号勘误邮箱：", size=12.5, color=C["text"]),
                    ft.Text("1220313772@qq.com", size=12.5, weight=ft.FontWeight.BOLD, color=C["link"]),
                ], spacing=4, tight=True)
            ),
        ], spacing=12, tight=True, width=_w(480)),
        actions=[
            _tb("我已知晓，进入工具", on_click=_close_license),
        ],
        actions_alignment=ft.MainAxisAlignment.END,
    )

    # ---------- 使用说明弹窗：本次运行内首次进入各弹一次；按钮可随时回看 ----------
    _hint_shown = {"selector": False, "fault": False, "repl": False}
    _hint_content = {
        "selector": ("伺服选型工具 · 使用说明", [
            "① 选条件：顶部按供电电压、额定功率、额定转速、惯量容量、法兰尺寸、驱动器系列等组合筛选；任一条件选「全部」即不限制该项，条件越具体结果越集中；没有型号完全符合时，会自动给出最接近的推荐搭配。",
            "② 看结果：结果分批加载，滑到底部点「加载更多」查看其余型号；卡片上可查配套线缆、下载对应手册；右下角悬浮箭头可快速回到页面顶部。",
            "③ 勾选导出：勾选需要的型号卡片后，按底部已选数量合并 / 分表导出 BOM 清单，支持 CSV 与模板配置；历史文件在「导出记录」中查阅或删除，「打开选型清单」直接进入归档文件夹。",
            "④ 纠错回退：型号配套如有错漏，可在卡片上「纠错」修改；保存前自动存快照，可在「纠错记录」里撤销、恢复历史快照或恢复出厂原始基线；「一键回原」清空全部筛选条件并恢复默认。",
        ]),
        "fault": ("故障码 / 警告码查询 · 使用说明", [
            "① 查码：先在顶部选择伺服系列（如 SV680N），再输入故障码或警告码精确查询；也可以直接输入关键词，在故障名称、故障机理与排除措施中做模糊检索。",
            "② 看结果：卡片显示故障名称、级别、能否复位、故障机理，以及「故障原因 — 确认方法 — 处理措施」三列排查表，按表逐项确认即可定位问题。",
            "③ 记录与纠错：「上一次查询」一键调取上次的系列与故障码，重启后仍保留；条目如有错漏可在卡片上「✏️ 纠错」修改，快照可在「🕘 纠错记录」里撤销、任选恢复或回到出厂原始。",
            "④ 提示：故障码结论仅供排查参考，现场接线、复位与处理请以汇川官方手册及实际情况为准。",
        ]),
        "repl": ("产品替换对照表 · 使用说明", [
            "① 用途：客户用汇川产品替换原品牌产品时，按方案逐项对照——「初始型号」照抄客户铭牌，「汇川对应型号」按选型结果填入，数量与备注按需补充。",
            "② 带入搭配：在选型页选好搭配并勾选后，点「📋 生成替换对照清单」按方案分组带入汇川型号；也可以在本页直接手填任一列。",
            "③ 分组规则：多套方案按功率 / 通讯分组；CN1 接口、电池型号跨套相同时并入共用「接口和电池」组并合计数量，型号不同则归入各自方案，不建共用组；单套方案不显示分组标题。",
            "④ 导出与清空：「⬇ 导出替换清单」生成五列对照表（类型 / 初始型号 / 汇川对应型号 / 数量 / 备注），存入「替换清单文件夹」并记入导出记录；「🧹 清空汇川型号」清掉汇川型号与备注、保留原型号和数量；「↺ 重置全部」重建空白表。",
        ]),
    }

    def _close_hint(dlg):
        dlg.open = False
        page.update()

    def _open_hint_dialog(kind):
        title, points = _hint_content[kind]
        body = [ft.Text(t, size=12.5, color=C["text"]) for t in points]
        dlg = ft.AlertDialog(
            open=True,
            title=ft.Text(title, size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
            content=ft.Column(body, spacing=10, tight=True, width=_w(460),
                              scroll=ft.ScrollMode.AUTO, height=380),
            actions=[_tb("我知道了", on_click=lambda e: _close_hint(dlg))],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        page.dialog = dlg
        page.update()

    def _show_hint(kind):
        if _hint_shown[kind]:
            return
        _hint_shown[kind] = True
        _open_hint_dialog(kind)

    # ---------- 应用主页 / 选型页 视图切换（控件常驻闭包，切换不丢状态） ----------
    home_view = None

    def _wrap_bg(view):
        bg = _custom_img_path(cfg, "bg") if mode == "custom" else None
        if not bg:
            return view
        return ft.Container(
            view,
            image_src=bg,
            image_fit=ft.ImageFit.COVER,
            image_opacity=0.35,
            expand=True,
        )

    def _show(view):
        page.controls.clear()
        page.add(ft.SafeArea(_wrap_bg(view), top=True, bottom=True, left=True, right=True, expand=True))
        page.update()

    def _go_home(e):
        if home_view is not None:
            page.floating_action_button = None
            _show(home_view)

    def opt(*vals):
        return [ft.dropdown.Option("全部")] + [ft.dropdown.Option(v) for v in vals]

    volt_drop = _dd(label="供电电压", options=opt("AC220", "AC380"),
                            value=cfg.get("default_voltage", "AC220"))
    def _kw(s):
        """从 '1.3kW' 这类文本里取出数值，用于功率比较"""
        try:
            return float("".join(c for c in str(s) if c.isdigit() or c == "."))
        except Exception:
            return 0.0

    def _flange_mm(s):
        return _kw(str(s).split("×")[0])

    power_vals = sorted({str(d.get("power_text")).strip() for d in data if str(d.get("power_text") or "").strip()},
                        key=_kw)
    power_drop = _dd(
        label="额定功率",
        options=opt(*power_vals),
        value="全部"
    )
    speed_drop = _dd(
        label="额定转速",
        options=opt("1500rpm", "2000rpm", "3000rpm"),
        value="全部"
    )
    series_drop = _dd(
        label="惯量 & 容量",
        options=opt(
            "H1｜低惯量-小容量",
            "H2｜低惯量-中容量",
            "H3｜中惯量-中容量",
            "H4｜中惯量-小容量"
        ),
        value="全部"
    )
    flange_vals = sorted(
        {str(d.get("flange_size")).strip() for d in data if str(d.get("flange_size") or "").strip()},
        key=lambda s: int("".join(ch for ch in s.split("×")[0] if ch.isdigit()) or 0)
    )
    flange_drop = _dd(
        label="法兰尺寸（机座）",
        options=opt(*flange_vals),
        value="全部"
    )
    driver_drop = _dd(
        label="驱动器系列",
        options=[
            ft.dropdown.Option("全部"),
            ft.dropdown.Option("SV630N", "SV630N（总线）"),
            ft.dropdown.Option("SV660N", "SV660N（总线）"),
            ft.dropdown.Option("SV670N", "SV670N（总线）"),
            ft.dropdown.Option("SV680N", "SV680N（总线）"),
            ft.dropdown.Option("SV600EN", "SV600EN（总线）"),
            ft.dropdown.Option("SV605N", "SV605N（总线）"),
            ft.dropdown.Option("SV605P", "SV605P（脉冲）"),
            ft.dropdown.Option("SV600EP", "SV600EP（脉冲）"),
            ft.dropdown.Option("SV630P", "SV630P（脉冲）"),
            ft.dropdown.Option("SV660P", "SV660P（脉冲）"),
            ft.dropdown.Option("SV670P", "SV670P（脉冲）"),
            ft.dropdown.Option("SV680P", "SV680P（脉冲）"),
        ],
        value="全部"
    )
    comm_drop = _dd(label="通讯类型", options=opt("EtherCAT", "脉冲"), value="全部")
    enc_drop = _dd(
        label="编码器（铭牌⑥位）",
        options=[
            ft.dropdown.Option("全部"),
            ft.dropdown.Option("A6", "A6｜26位多圈绝对值"),
            ft.dropdown.Option("S6", "S6｜安全型26位多圈"),
            ft.dropdown.Option("A3", "A3｜23位多圈绝对值"),
            ft.dropdown.Option("T3", "T3｜18位多圈绝对值"),
            ft.dropdown.Option("T2", "T2｜18位多圈绝对值"),
            ft.dropdown.Option("C2", "C2｜18位多圈绝对值"),
            ft.dropdown.Option("U2", "U2｜20位单圈（25机座）"),
        ],
        value="全部"
    )
    magnet_drop = _dd(
        label="编码器形式",
        options=[
            ft.dropdown.Option("全部"),
            ft.dropdown.Option("是", "免电池磁编（不配电池）"),
            ft.dropdown.Option("否", "传统光编（需电池）"),
        ],
        value="否"
    )
    battery_drop = _dd(
        label="绝对值电池",
        options=[
            ft.dropdown.Option("全部"),
            ft.dropdown.Option("需要", "需要电池（S6-C4A）"),
            ft.dropdown.Option("免电池", "免电池（不配电池）"),
        ],
        value="需要"
    )
    depth_drop = _dd(
        label="线缆长度",
        options=[ft.dropdown.Option(v, t) for v, t in [
            ("3.0", "3 m（-3.0）"),
            ("5.0", "5 m（-5.0）"),
            ("10.0", "10 m（-10.0）"),
        ]],
        value=cfg.get("default_depth", "3.0")
    )
    flex_check = _cb(label="高柔线（-T）", value=bool(cfg.get("default_high_flex", True)))
    brake_check = _cb(label="需要抱闸", value=False)
    oilseal_check = _cb(label="带油封", value=True)

    def _snapshot_windows():
        """采集当前可见、带标题的顶层窗口句柄集合，用于打开外部程序后识别新窗口"""
        if not _is_desktop(page):
            return set()
        user32 = ctypes.windll.user32
        snaps = set()

        @WNDENUMPROC
        def _enum(hwnd, _):
            if user32.IsWindowVisible(hwnd):
                n = user32.GetWindowTextLengthW(hwnd)
                if n:
                    snaps.add(hwnd)
            return True

        user32.EnumWindows(_enum, 0)
        return snaps

    def _activate_window(keyword, before=None, delay=0.6):
        """打开外部文件/文件夹后，延时把对应新窗口拉到最前并激活"""
        if not _is_desktop(page):
            return
        before = before or set()

        def _worker():
            time.sleep(delay)
            try:
                user32 = ctypes.windll.user32
                kernel32 = ctypes.windll.kernel32
                matched, news = [], []

                @WNDENUMPROC
                def _enum(hwnd, _):
                    if user32.IsWindowVisible(hwnd):
                        n = user32.GetWindowTextLengthW(hwnd)
                        if n:
                            buf = ctypes.create_unicode_buffer(n + 1)
                            user32.GetWindowTextW(hwnd, buf, n + 1)
                            if keyword and keyword in buf.value:
                                matched.append(hwnd)
                            if hwnd not in before:
                                news.append(hwnd)
                    return True

                user32.EnumWindows(_enum, 0)
                target = matched[-1] if matched else (news[-1] if news else None)
                if not target:
                    return
                fg = user32.GetForegroundWindow()
                fg_tid = user32.GetWindowThreadProcessId(fg, None)
                cur_tid = kernel32.GetCurrentThreadId()
                attached = False
                try:
                    user32.AttachThreadInput(cur_tid, fg_tid, True)
                    attached = True
                    user32.ShowWindow(target, 9)
                    user32.BringWindowToTop(target)
                    user32.SetActiveWindow(target)
                    user32.SetForegroundWindow(target)
                finally:
                    if attached:
                        user32.AttachThreadInput(cur_tid, fg_tid, False)
            except Exception:
                pass

        threading.Thread(target=_worker, daemon=True).start()

    def _close_dlg(dlg):
        dlg.open = False
        page.update()

    _dlg_img = _custom_img_path(cfg, "dialog") if mode == "custom" else None
    _dlg_bg = {}
    if _dlg_img:
        _dlg_bg = {"image_src": _dlg_img, "image_fit": ft.ImageFit.COVER, "image_opacity": 0.12}

    def _open_local_folder(folder):
        if not _is_desktop(page):
            page.show_snack_bar(ft.SnackBar(
                content=ft.Text("当前环境无法直接打开本地文件夹，请在桌面版查看导出清单")))
            return
        os.makedirs(folder, exist_ok=True)
        os.startfile(folder)

        def _bring_front():
            time.sleep(0.5)
            try:
                user32 = ctypes.windll.user32
                kernel32 = ctypes.windll.kernel32
                target = os.path.basename(folder)
                matches = []

                @WNDENUMPROC
                def _enum(hwnd, _):
                    if user32.IsWindowVisible(hwnd):
                        n = user32.GetWindowTextLengthW(hwnd)
                        if n:
                            buf = ctypes.create_unicode_buffer(n + 1)
                            user32.GetWindowTextW(hwnd, buf, n + 1)
                            cls = ctypes.create_unicode_buffer(256)
                            user32.GetClassNameW(hwnd, cls, 256)
                            if cls.value in ("CabinetWClass", "ExploreWClass") and target in buf.value:
                                matches.append(hwnd)
                    return True

                user32.EnumWindows(_enum, 0)
                if not matches:
                    return
                hwnd = matches[-1]
                fg = user32.GetForegroundWindow()
                fg_tid = user32.GetWindowThreadProcessId(fg, None)
                cur_tid = kernel32.GetCurrentThreadId()
                attached = False
                try:
                    user32.AttachThreadInput(cur_tid, fg_tid, True)
                    attached = True
                    user32.ShowWindow(hwnd, 9)
                    user32.BringWindowToTop(hwnd)
                    user32.SetActiveWindow(hwnd)
                    user32.SetForegroundWindow(hwnd)
                finally:
                    if attached:
                        user32.AttachThreadInput(cur_tid, fg_tid, False)
            except Exception:
                pass

        threading.Thread(target=_bring_front, daemon=True).start()

    def open_bom_folder(e):
        _open_local_folder(EXPORT_DIR)

    def open_repl_folder(e):
        _open_local_folder(REPL_EXPORT_DIR)

    btn_open_folder = _eb(
        "📂 打开选型清单",
        height=34,
        bgcolor=C["primary"],
        color="white",
        on_click=open_bom_folder,
    )

    btn_open_repl_folder = _eb(
        "📂 打开替换清单",
        height=34,
        bgcolor=C["primary"],
        color="white",
        on_click=open_repl_folder,
    )

    # 选型手册清单：(显示名, 下载直链, 本地存盘文件名)
    MANUALS = [
        ("SV630P 系列伺服用户手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00006294_PDF_CN/D00/SV630P%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E7%94%A8%E6%88%B7%E6%89%8B%E5%86%8C-CN-D00.PDF",
         "SV630P系列伺服用户手册-CN-D00.PDF"),
        ("SV630N 系列伺服用户手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00006295_PDF_CN/D00/SV630N%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E7%94%A8%E6%88%B7%E6%89%8B%E5%86%8C-CN-D00.PDF",
         "SV630N系列伺服用户手册-CN-D00.PDF"),
        ("SV660N 系列伺服选型手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00003006_PDF_CN/D01/SV660N%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E9%80%89%E5%9E%8B%E6%89%8B%E5%86%8C-CN-D01.PDF",
         "SV660N系列伺服选型手册-CN-D01.PDF"),
        ("SV660N 系列伺服用户手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00005512_PDF_CN/D01/SV660N%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E7%94%A8%E6%88%B7%E6%89%8B%E5%86%8C-CN-D01.PDF",
         "SV660N系列伺服用户手册-CN-D01.PDF"),
        ("SV660P 系列伺服选型手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00003068_PDF_CN/D01/SV660P%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E9%80%89%E5%9E%8B%E6%89%8B%E5%86%8C-CN-D01.PDF",
         "SV660P系列伺服选型手册-CN-D01.PDF"),
        ("SV660P 系列伺服用户手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00005513_PDF_CN/D01/SV660P%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E7%94%A8%E6%88%B7%E6%89%8B%E5%86%8C-CN-D01.PDF",
         "SV660P系列伺服用户手册-CN-D01.PDF"),
        ("SV670N 系列伺服选型手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00005327_PDF_CN/B00/SV670N%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E9%80%89%E5%9E%8B%E6%89%8B%E5%86%8C-CN-B00.PDF",
         "SV670N系列伺服选型手册-CN-B00.PDF"),
        ("SV670N 系列伺服用户手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00005527_PDF_CN/B00/SV670N%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E7%94%A8%E6%88%B7%E6%89%8B%E5%86%8C-CN-B00.PDF",
         "SV670N系列伺服用户手册-CN-B00.PDF"),
        ("SV670P 系列伺服选型手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00005165_PDF_CN/B00/SV670P%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E9%80%89%E5%9E%8B%E6%89%8B%E5%86%8C-CN-B00.PDF",
         "SV670P系列伺服选型手册-CN-B00.PDF"),
        ("SV670P 系列伺服用户手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00005526_PDF_CN/B00/SV670P%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E7%94%A8%E6%88%B7%E6%89%8B%E5%86%8C-CN-B00.PDF",
         "SV670P系列伺服用户手册-CN-B00.PDF"),
        ("SV680N 系列伺服选型手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00003312_PDF_CN/C00/SV680N%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E9%80%89%E5%9E%8B%E6%89%8B%E5%86%8C-CN-C00.PDF",
         "SV680N系列伺服选型手册-CN-C00.PDF"),
        ("SV680N 系列伺服用户手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00005405_PDF_CN/B00/SV680N%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E7%94%A8%E6%88%B7%E6%89%8B%E5%86%8C-CN-B00.PDF",
         "SV680N系列伺服用户手册-CN-B00.PDF"),
        ("SV680P 系列伺服选型手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00003105_PDF_CN/C00/SV680P%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E9%80%89%E5%9E%8B%E6%89%8B%E5%86%8C-CN-C00.PDF",
         "SV680P系列伺服选型手册-CN-C00.PDF"),
        ("SV680P 系列伺服用户手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00005388_PDF_CN/B00/SV680P%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E7%94%A8%E6%88%B7%E6%89%8B%E5%86%8C-CN-B00.PDF",
         "SV680P系列伺服用户手册-CN-B00.PDF"),
        ("SV600EN 系列伺服选型手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00009079_PDF_CN/B00/SV600EN%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E9%80%89%E5%9E%8B%E6%89%8B%E5%86%8C-CN-B00.PDF",
         "SV600EN系列伺服选型手册-CN-B00.PDF"),
        ("SV600EN 系列伺服用户手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00009078_PDF_CN/B00/SV600EN%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E7%94%A8%E6%88%B7%E6%89%8B%E5%86%8C-CN-B00.PDF",
         "SV600EN系列伺服用户手册-CN-B00.PDF"),
        ("SV600EP 系列伺服选型手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00009094_PDF_CN/B00/SV600EP%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E9%80%89%E5%9E%8B%E6%89%8B%E5%86%8C-CN-B00.PDF",
         "SV600EP系列伺服选型手册-CN-B00.PDF"),
        ("SV600EP 系列伺服用户手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00009075_PDF_CN/B00/SV600EP%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E7%94%A8%E6%88%B7%E6%89%8B%E5%86%8C-CN-B00.PDF",
         "SV600EP系列伺服用户手册-CN-B00.PDF"),
        ("SV605N 系列伺服用户手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00017524_PDF_CN/B00/SV605N%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E7%94%A8%E6%88%B7%E6%89%8B%E5%86%8C-CN-B00.PDF",
         "SV605N系列伺服用户手册-CN-B00.PDF"),
        ("SV605P/C 系列伺服用户手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00009745_PDF_CN/B00/SV605P_C%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E7%94%A8%E6%88%B7%E6%89%8B%E5%86%8C-CN-B00.PDF",
         "SV605P_C系列伺服用户手册-CN-B00.PDF"),
        ("SV680LN 直驱电机伺服选型手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00004890_PDF_CN/A05/SV680LN%E7%9B%B4%E9%A9%B1%E7%94%B5%E6%9C%BA%E4%BC%BA%E6%9C%8D%E9%80%89%E5%9E%8B%E6%89%8B%E5%86%8C-CN-A05.PDF",
         "SV680LN直驱电机伺服选型手册-CN-A05.PDF"),
        ("MS1-R 系列伺服电机选型手册",
         "https://newweb.inovance.com/owfile/ProdDoc/SC/PS00004605_PDF_CN/B05/MS1-R%E7%B3%BB%E5%88%97%E4%BC%BA%E6%9C%8D%E7%94%B5%E6%9C%BA%E9%80%89%E5%9E%8B%E6%89%8B%E5%86%8C-CN-B05.PDF",
         "MS1-R系列伺服电机选型手册-CN-B05.PDF"),
    ]

    _ssl_ctx = ssl.create_default_context()

    def _download_manual(url, dest, btn, state, rem_btn, prog):
        tmp = dest + ".part"
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=30, context=_ssl_ctx) as r:
                total = r.headers.get("Content-Length")
                total = int(total) if total and str(total).isdigit() else None
                downloaded = 0
                last_update = 0
                prog.visible = True
                prog.value = 0
                page.update()
                with open(tmp, "wb") as f:
                    while True:
                        chunk = r.read(64 * 1024)
                        if not chunk:
                            break
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total:
                            ratio = min(downloaded / total, 1.0)
                            prog.value = ratio
                            btn.text = f"⬇ {int(ratio * 100)}%"
                        else:
                            prog.value = None
                        if downloaded - last_update >= 256 * 1024 or downloaded == total:
                            page.update()
                            last_update = downloaded
            os.replace(tmp, dest)
            state["ready"] = True
            btn.text = "📂 打开"
            btn.bgcolor = C["success"]
            btn.disabled = False
            rem_btn.visible = True
            prog.visible = False
            _log(f"手册下载完成：{os.path.basename(dest)}")
        except Exception as ex:
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except Exception:
                pass
            state["ready"] = False
            btn.text = "⬇ 重试"
            btn.disabled = False
            prog.visible = False
            _log(f"手册下载失败：{os.path.basename(dest)}（{ex}）", "ERROR")
            tip_text.value = f"手册下载失败：{os.path.basename(dest)}（{ex}）"
            tip_text.color = C["danger"]
        page.update()

    def open_manual_list(e):
        manual_dir = os.path.join(base_dir, "选型手册")
        os.makedirs(manual_dir, exist_ok=True)
        rows = []

        for title, url, filename in MANUALS:
            dest = os.path.join(manual_dir, filename)
            state = {"ready": os.path.exists(dest)}
            btn = _eb(
                "📂 打开" if state["ready"] else "⬇ 下载",
                height=30,
                bgcolor=C["success"] if state["ready"] else C["primary"],
                color="white",
            )
            rem_btn = _eb(
                "🗑 移除",
                height=30,
                bgcolor=C["danger"],
                color="white",
                visible=state["ready"] and _is_desktop(page),
            )
            prog = ft.ProgressBar(
                value=0, visible=False, bar_height=6,
                color=C["primary"], bgcolor=C["chip"], width=_w(420),
            )

            def _on_click(ev, dest=dest, url=url, filename=filename, btn=btn, state=state, rem_btn=rem_btn, prog=prog):
                if not _is_desktop(page):
                    page.launch_url(url)
                    return
                if state["ready"]:
                    manual_dlg.open = False
                    page.update()
                    before = _snapshot_windows()
                    try:
                        os.startfile(dest)
                        _activate_window(os.path.splitext(filename)[0], before, delay=0.7)
                    except Exception as ex:
                        tip_text.value = f"打开手册失败：{ex}"
                        tip_text.color = C["danger"]
                        page.update()
                    return
                btn.disabled = True
                btn.text = "下载中…"
                page.update()
                threading.Thread(
                    target=_download_manual,
                    args=(url, dest, btn, state, rem_btn, prog),
                    daemon=True,
                ).start()

            def _on_remove(ev, dest=dest, filename=filename, btn=btn, state=state, rem_btn=rem_btn):
                try:
                    if os.path.exists(dest):
                        os.remove(dest)
                    part = dest + ".part"
                    if os.path.exists(part):
                        os.remove(part)
                    state["ready"] = False
                    btn.text = "⬇ 下载"
                    btn.bgcolor = C["primary"]
                    btn.disabled = False
                    rem_btn.visible = False
                    tip_text.value = f"已移除本地手册：{filename}"
                    tip_text.color = C["success"]
                    _log(f"移除本地手册：{filename}")
                except Exception as ex:
                    _log(f"手册移除失败：{filename}（{ex}）", "ERROR")
                    tip_text.value = f"移除失败：{filename}（{ex}）"
                    tip_text.color = C["danger"]
                page.update()

            btn.on_click = _on_click
            rem_btn.on_click = _on_remove
            meta = ""
            if state["ready"]:
                meta = f"{_fmt_size(dest)} · {_fmt_mtime(dest)}"
            rows.append(ft.Column([
                ft.Row([
                    ft.Text(title, size=12.5, color=C["text"], expand=True),
                    btn,
                    rem_btn,
                ], spacing=10),
                ft.Text(meta, size=10.5, color=C["muted"]) if meta else ft.Container(),
                prog,
            ], spacing=4))

        content = ft.Column(rows, spacing=8, scroll=ft.ScrollMode.AUTO, width=_w(440), height=400)

        def _close_manual(ev):
            manual_dlg.open = False
            page.update()

        manual_dlg = ft.AlertDialog(
            open=True,
            title=ft.Text("选型手册（点下载自动保存到本地）", size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
            content=content,
            actions=[_tb("关闭", on_click=_close_manual)],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        page.dialog = manual_dlg
        manual_dlg.open = True
        page.update()

    # ---------- 线缆型号查询弹窗 ----------
    cable_motor = _tf(
        label="电机型号（忽略大小写）",
        hint_text="如 MS1H3-13C15CB-A634R",
        width=_w(320),
    )
    cable_depth = _dd(
        label="线缆长度",
        options=[ft.dropdown.Option(v, t) for v, t in [
            ("3.0", "3 m（-3.0）"),
            ("5.0", "5 m（-5.0）"),
            ("10.0", "10 m（-10.0）"),
        ]],
        value="3.0",
        width=_w(150),
    )
    cable_flex = _cb(label="高柔线（-T）", value=True)
    cable_result = ft.Column(spacing=8, scroll=ft.ScrollMode.AUTO, height=260, width=_w(460))

    def _cable_card(d, depth, flex):
        motor = str(d.get("motor_model") or "")
        drive = str(d.get("drive_model") or "")
        power = build_line(d.get("power_line_base"), depth, flex)
        enc = build_line(d.get("encoder_line_base"), depth, flex)
        brake_txt = "带抱闸" if str(d.get("has_brake")) == "是" else "无抱闸"
        flex_txt = "高柔" if flex else "普通"
        return ft.Container(
            ft.Column([
                ft.Row([
                    ft.Text(motor, size=12.5, weight=ft.FontWeight.BOLD, color=C["text"]),
                    ft.Text(drive, size=11, color=C["muted"]),
                ], spacing=8),
                ft.Row([
                    ft.Text("动力线：", size=12, color=C["text"]),
                    ft.Text(power, size=12.5, weight=ft.FontWeight.BOLD, color=C["success"]),
                ], spacing=6),
                ft.Row([
                    ft.Text("编码线：", size=12, color=C["text"]),
                    ft.Text(enc, size=12.5, weight=ft.FontWeight.BOLD, color=C["success"]),
                ], spacing=6),
                ft.Text(f"{brake_txt} · {depth}米 · {flex_txt}", size=11, color=C["muted"]),
            ], spacing=4),
            bgcolor=C["card_green"],
            border_radius=R_PANEL,
            padding=12,
            border=ft.border.all(1, C["border_green"]),
        )

    def on_cable_search(e):
        q = (cable_motor.value or "").strip().upper().replace(" ", "")
        cable_result.controls.clear()
        if not q:
            cable_result.controls.append(ft.Text("请先输入电机型号", size=12.5, color=C["danger"]))
            page.update()
            return
        depth = cable_depth.value
        flex = bool(cable_flex.value)
        hits, seen = [], set()
        for d in data:
            m = str(d.get("motor_model") or "").strip().upper()
            if not (q == m or q in m):
                continue
            key = (m, str(d.get("drive_model") or ""))
            if key in seen:
                continue
            seen.add(key)
            hits.append((q == m, d))
        hits.sort(key=lambda x: (not x[0], str(x[1].get("drive_model") or "")))
        if not hits:
            cable_result.controls.append(ft.Text(
                "未在数据库找到该型号，请核对铭牌后重试；型号如有出入可发邮件反馈勘误。",
                size=12.5, color=C["danger"]))
        else:
            for _exact, d in hits:
                cable_result.controls.append(_cable_card(d, depth, flex))
        page.update()

    def open_cable_query(e):
        page.dialog = cable_dlg
        cable_dlg.open = True
        page.update()

    def _close_cable(ev):
        cable_dlg.open = False
        page.update()

    cable_dlg = ft.AlertDialog(
        open=False,
        title=ft.Text("线缆型号查询", size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
        content=ft.Column([
            ft.Row([cable_motor], wrap=True),
            ft.Row([cable_depth, cable_flex], spacing=16, wrap=True),
            ft.Row([
                _eb("🔍 查询线缆", bgcolor=C["primary"], color="white",
                                  height=32, on_click=on_cable_search),
            ], alignment=ft.MainAxisAlignment.END),
            ft.Divider(height=8, color=C["divider"]),
            cable_result,
        ], spacing=10, tight=True, width=_w(480)),
        actions=[_tb("关闭", on_click=_close_cable)],
        actions_alignment=ft.MainAxisAlignment.END,
    )

    btn_cable = _eb(
        "🔌 线缆查询",
        height=34,
        bgcolor=C["success"],
        color="white",
        on_click=open_cable_query,
    )

    btn_hint_sel = _eb(
        "❓ 使用说明",
        height=34,
        bgcolor=C["gray"],
        color="white",
        on_click=lambda e: _open_hint_dialog("selector"),
    )

    # ---------- 纠错 + 版本恢复 ----------
    fix_target = {"it": None, "note": None}
    fix_fields_ctrl = {}

    def _open_fix(e, it):
        fix_target["it"] = it
        fix_fields_ctrl.clear()
        field_rows = []
        for label, col in FIX_FIELDS:
            val = "" if it.get(col) is None else str(it.get(col))
            tf = _tf(label=label, value=val, height=38,
                              content_padding=8, width=_w(440))
            fix_fields_ctrl[col] = tf
            field_rows.append(tf)
        fix_note = _tf(label="纠错备注（可选）", value="", height=38,
                                content_padding=8, width=_w(440))
        fix_target["note"] = fix_note
        fix_dlg.content = ft.Column(
            [ft.Text(f"电机：{it.get('motor_model')}　驱动器：{it.get('drive_model')}",
                     size=12.5, weight=ft.FontWeight.BOLD, color=C["title"]),
             *field_rows, fix_note],
            spacing=10, scroll=ft.ScrollMode.AUTO, width=_w(460), height=460)
        page.dialog = fix_dlg
        fix_dlg.open = True
        page.update()

    def _close_fix(ev):
        fix_dlg.open = False
        page.update()

    def _save_fix(ev):
        it = fix_target["it"]
        if not it:
            return
        motor = str(it.get("motor_model") or "").strip()
        drive = str(it.get("drive_model") or "").strip()
        changes = []
        for label, col in FIX_FIELDS:
            old = "" if it.get(col) is None else str(it.get(col)).strip()
            new = (fix_fields_ctrl[col].value or "").strip()
            if new != old:
                changes.append((col, label, old, new))
        note = (fix_target["note"].value or "").strip()
        if not changes:
            tip_text.value = "没有需要保存的改动"
            tip_text.color = C["warn"]
            page.update()
            return
        try:
            os.makedirs(SNAP_DIR, exist_ok=True)
            idx = _next_snapshot_index()
            snap_path = os.path.join(SNAP_DIR, f"纠错_{idx:03d}.xlsx")
            _copy_file(DATA_FILE, snap_path)

            wb = openpyxl.load_workbook(DATA_FILE)
            ws = wb["data"]
            header = [c.value for c in ws[1]]
            mi = header.index("motor_model")
            di = header.index("drive_model")
            row_idx = None
            for r in range(2, ws.max_row + 1):
                if str(ws.cell(r, mi + 1).value or "").strip() == motor and \
                   str(ws.cell(r, di + 1).value or "").strip() == drive:
                    row_idx = r
                    break
            if row_idx is None:
                tip_text.value = "未在表格中定位到该记录，未保存"
                tip_text.color = C["danger"]
                page.update()
                return
            for col, label, old, new in changes:
                ci = header.index(col) + 1
                ws.cell(row_idx, ci).value = new if new else None
            wb.save(DATA_FILE)

            entries = _read_fix_log()
            seq = len(entries) + 1
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for col, label, old, new in changes:
                entries.append({
                    "序号": seq, "时间": now, "电机型号": motor, "驱动器": drive,
                    "字段": label, "改前值": old, "改后值": new,
                    "备注": note, "快照文件": os.path.basename(snap_path),
                })
            _write_fix_log(entries)

            for d in data:
                if str(d.get("motor_model") or "").strip() == motor and \
                   str(d.get("drive_model") or "").strip() == drive:
                    for col, label, old, new in changes:
                        d[col] = new if new else None
                    break

            fix_dlg.open = False
            tip_text.value = f"已保存纠错：{motor}（{len(changes)} 项）"
            tip_text.color = C["success"]
            _log(f"纠错保存：{motor} {drive}，{len(changes)} 项改动")
            refresh(None)
        except Exception as ex:
            _log(f"纠错保存失败：{ex}", "ERROR")
            tip_text.value = f"纠错保存失败：{ex}"
            tip_text.color = C["danger"]
            page.update()

    fix_dlg = ft.AlertDialog(
        open=False,
        title=ft.Text("✏️ 纠错搭配", size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
        content=ft.Column([ft.Text("")], width=_w(460), height=460),
        actions=[
            _tb("取消", on_click=_close_fix),
            _tb("保存纠错", on_click=_save_fix),
        ],
        actions_alignment=ft.MainAxisAlignment.END,
    )

    def _reload_data():
        data[:] = load_excel(DATA_FILE)

    def _restore_snapshot(snap_file):
        if not snap_file:
            return
        snap_path = os.path.join(SNAP_DIR, snap_file)
        if not os.path.exists(snap_path):
            tip_text.value = "快照文件不存在，无法恢复"
            tip_text.color = C["danger"]
            page.update()
            return
        _copy_file(snap_path, DATA_FILE)
        _reload_data()
        hist_dlg.open = False
        tip_text.value = f"已恢复到快照：{snap_file}"
        tip_text.color = C["success"]
        _log(f"恢复快照：{snap_file}")
        refresh(None)

    orig_confirm = {"armed": False}
    orig_btn = _eb(
        "♻ 恢复出厂原始", height=34, bgcolor=C["danger"], color="white")

    def _restore_original(ev):
        if not os.path.exists(ORIGINAL_FILE):
            tip_text.value = "未找到出厂原始基线"
            tip_text.color = C["danger"]
            page.update()
            return
        if not orig_confirm["armed"]:
            orig_confirm["armed"] = True
            orig_btn.text = "再次点击确认（撤销全部纠错）"
            orig_btn.color = C["danger"]
            page.update()
            return
        _copy_file(ORIGINAL_FILE, DATA_FILE)
        _reload_data()
        orig_confirm["armed"] = False
        orig_btn.text = "♻ 恢复出厂原始"
        orig_btn.color = "white"
        hist_dlg.open = False
        tip_text.value = "已恢复为出厂原始表格（全部纠错已撤销）"
        tip_text.color = C["success"]
        _log("恢复出厂原始表格（撤销全部纠错）")
        refresh(None)

    orig_btn.on_click = _restore_original

    def _delete_entry(seq, snap_file):
        entries = _read_fix_log()
        kept = [e for e in entries if str(e.get("序号")) != str(seq)]
        if snap_file and not any(str(e.get("快照文件") or "") == snap_file for e in kept):
            sp = os.path.join(SNAP_DIR, snap_file)
            if os.path.exists(sp):
                os.remove(sp)
        _write_fix_log(kept)
        tip_text.value = f"已删除第 {seq} 条纠错记录"
        tip_text.color = C["success"]
        _log(f"删除纠错记录 #{seq}")
        open_history(None)

    clear_confirm = {"armed": False}
    clear_btn = _eb("🗑 清空全部记录", height=34, bgcolor=C["danger"], color="white")

    def _clear_all_entries(ev):
        if not clear_confirm["armed"]:
            clear_confirm["armed"] = True
            clear_btn.text = "再次点击确认清空（不回滚当前表格）"
            clear_btn.color = C["danger"]
            page.update()
            return
        _write_fix_log([])
        if os.path.isdir(SNAP_DIR):
            for f in os.listdir(SNAP_DIR):
                if f.endswith(".xlsx"):
                    os.remove(os.path.join(SNAP_DIR, f))
        clear_confirm["armed"] = False
        clear_btn.text = "🗑 清空全部记录"
        clear_btn.color = "white"
        tip_text.value = "已清空全部纠错记录（当前表格与出厂基线不受影响）"
        _log("清空全部纠错记录")
        tip_text.color = C["success"]
        open_history(None)

    clear_btn.on_click = _clear_all_entries

    # ---------- 选型库 文件导出 / 导入（整库覆盖，导入前自动存快照可回滚） ----------
    imp_target = {"kind": "selector"}

    def _export_selector_db(ev):
        if not _is_desktop(page):
            tip_text.value = "当前环境暂不支持导出文件，请使用桌面版导出"
            tip_text.color = C["warn"]
            page.update()
            return
        try:
            os.makedirs(DATA_EXPORT_DIR, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out = os.path.join(DATA_EXPORT_DIR, f"选型库_{stamp}.xlsx")
            _copy_file(DATA_FILE, out)
            tip_text.value = f"已导出选型库：{os.path.basename(out)}"
            tip_text.color = C["success"]
            _log(f"导出选型库：{out}")
        except Exception as ex:
            _log(f"导出选型库失败：{ex}", "ERROR")
            tip_text.value = f"导出失败：{ex}"
            tip_text.color = C["danger"]
        page.update()

    def _import_selector_db(path):
        if not path or not os.path.exists(path):
            return
        try:
            wb = openpyxl.load_workbook(path)
            if "data" not in wb.sheetnames:
                tip_text.value = "导入文件缺少 data 工作表，未导入"
                tip_text.color = C["danger"]
                page.update()
                return
            ws = wb["data"]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                tip_text.value = "导入文件为空，未导入"
                tip_text.color = C["danger"]
                page.update()
                return
            header = [str(h) if h is not None else "" for h in rows[0]]
            if "motor_model" not in header or "drive_model" not in header:
                tip_text.value = "导入文件表头缺少 motor_model / drive_model，未导入"
                tip_text.color = C["danger"]
                page.update()
                return
            os.makedirs(SNAP_DIR, exist_ok=True)
            idx = _next_snapshot_index()
            snap_path = os.path.join(SNAP_DIR, f"纠错_{idx:03d}.xlsx")
            _copy_file(DATA_FILE, snap_path)
            _copy_file(path, DATA_FILE)
            data[:] = load_excel(DATA_FILE)
            entries = _read_fix_log()
            seq = len(entries) + 1
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            entries.append({
                "序号": seq, "时间": now, "电机型号": "（整库导入）", "驱动器": "",
                "字段": "整库导入", "改前值": "导入前整库", "改后值": os.path.basename(path),
                "备注": "", "快照文件": os.path.basename(snap_path),
            })
            _write_fix_log(entries)
            tip_text.value = f"已导入选型库：{len(data)} 条（导入前已存快照）"
            tip_text.color = C["success"]
            _log(f"整库导入选型库：{os.path.basename(path)}，{len(data)} 条")
            refresh(None)
            open_history(None)
        except Exception as ex:
            _log(f"导入选型库失败：{ex}", "ERROR")
            tip_text.value = f"导入失败：{ex}"
            tip_text.color = C["danger"]
            page.update()

    def _on_pick_import(res):
        if not res.files:
            return
        path = res.files[0].path
        if imp_target["kind"] == "fault":
            _import_fault_db(path)
        else:
            _import_selector_db(path)

    def _import_sel_click(ev):
        if not _is_desktop(page):
            tip_text.value = "当前环境暂不支持导入文件，请在桌面版导入"
            tip_text.color = C["warn"]
            page.update()
            return
        imp_target["kind"] = "selector"
        imp_picker.pick_files(dialog_title="选择要导入的选型库 xlsx",
                              allowed_extensions=["xlsx"])

    def _import_fault_click(ev):
        if not _is_desktop(page):
            fault_tip.value = "当前环境暂不支持导入文件，请在桌面版导入"
            fault_tip.color = C["warn"]
            page.update()
            return
        imp_target["kind"] = "fault"
        imp_picker.pick_files(dialog_title="选择要导入的故障码库 xlsx",
                              allowed_extensions=["xlsx"])

    imp_picker = ft.FilePicker(on_result=_on_pick_import)
    page.overlay.append(imp_picker)

    def open_history(e):
        entries = _read_fix_log()
        orig_confirm["armed"] = False
        orig_btn.text = "♻ 恢复出厂原始"
        orig_btn.color = "white"
        clear_confirm["armed"] = False
        clear_btn.text = "🗑 清空全部记录"
        clear_btn.color = "white"
        hist_rows = []
        if not entries:
            hist_rows.append(ft.Text("暂无纠错记录", size=12.5, color=C["muted"]))
        for en in entries:
            snap = en.get("快照文件") or ""
            seq = en.get("序号")
            hist_rows.append(ft.Row([
                ft.Column([
                    ft.Text(f"#{seq}　{en.get('时间')}", size=11, color=C["muted"]),
                    ft.Text(f"{en.get('电机型号')}　{en.get('驱动器')}", size=12, color=C["text"]),
                    ft.Text(f"{en.get('字段')}：{en.get('改前值') or '空'} → {en.get('改后值') or '空'}",
                            size=11.5, color=C["link"]),
                ], spacing=2, expand=True),
                _eb("恢复此版本", height=30, bgcolor=C["gray"], color="white",
                                  on_click=lambda ev, s=snap: _restore_snapshot(s)),
                _eb("🗑", height=30, bgcolor=C["danger"], color="white",
                                  on_click=lambda ev, q=seq, fn=snap: _delete_entry(q, fn)),
            ], spacing=10))
        last = entries[-1] if entries else None
        undo_btn = _eb(
            "↩ 撤销最近一次", height=34, bgcolor=C["primary"], color="white",
            disabled=last is None,
            on_click=lambda ev: _restore_snapshot(last.get("快照文件") if last else ""))
        exp_sel_btn = _eb("📤 导出文件", height=34, bgcolor=C["gray"], color="white",
                                         on_click=_export_selector_db)
        imp_sel_btn = _eb(
            "📥 导入文件", height=34, bgcolor=C["gray"], color="white",
            on_click=_import_sel_click)
        hist_dlg.content = ft.Column(
            [ft.Row([undo_btn, orig_btn, clear_btn], spacing=12, wrap=True),
             ft.Row([exp_sel_btn, imp_sel_btn], spacing=12, wrap=True),
             ft.Divider(height=8, color=C["divider"]),
             *hist_rows],
            spacing=10, scroll=ft.ScrollMode.AUTO, width=_w(480), height=440)
        page.dialog = hist_dlg
        hist_dlg.open = True
        page.update()

    def _close_hist(ev):
        hist_dlg.open = False
        page.update()

    hist_dlg = ft.AlertDialog(
        open=False,
        title=ft.Text("🕘 纠错记录与版本恢复", size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
        content=ft.Column([ft.Text("")], width=_w(480), height=440),
        actions=[_tb("关闭", on_click=_close_hist)],
        actions_alignment=ft.MainAxisAlignment.END,
    )

    btn_hist = _eb(
        "🕘 纠错记录",
        height=34,
        bgcolor=C["gray"],
        color="white",
        on_click=open_history,
    )

    # ---------- 更新日志弹窗 ----------
    def open_changelog(e):
        page.dialog = log_dlg
        log_dlg.open = True
        page.update()

    def _close_log(ev):
        log_dlg.open = False
        page.update()

    def _ver_num(title):
        m = re.search(r"(\d+(?:\.\d+){1,2})", title)
        return m.group(1) if m else title

    def _group_key(ver):
        parts = ver.split(".")
        if parts and parts[0] == "2":
            minor = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
            if minor <= 3:
                return "2.0–2.3"
            if minor <= 7:
                return "2.4–2.7"
            return "2.8–2.12"
        return ".".join(parts[:2])

    log_groups = []
    _group_map = {}
    for title, items in CHANGELOG:
        ver = _ver_num(title)
        gk = _group_key(ver)
        if gk not in _group_map:
            _group_map[gk] = []
            log_groups.append((gk, _group_map[gk]))
        _group_map[gk].append((title, ver, items))

    def _version_block(title, ver, items):
        label = f"{title}（{LATEST_DATE}）" if ver == VERSION else title
        return ft.Column([
            ft.Text(label, size=12.5, weight=ft.FontWeight.BOLD, color=C["title"]),
            ft.Column([ft.Text("• " + it, size=12, color=C["text"]) for it in items], spacing=3),
        ], spacing=5)

    log_tiles = []
    for idx, (gk, entries) in enumerate(log_groups):
        summary = CHANGELOG_SUMMARY.get(gk, "")
        log_tiles.append(ft.ExpansionTile(
            title=ft.Text(f"v{gk}", size=14, weight=ft.FontWeight.BOLD, color=C["title"]),
            subtitle=ft.Text(summary, size=11.5, color=C["sub"]) if summary else None,
            initially_expanded=(idx == 0),
            controls_padding=ft.padding.symmetric(horizontal=14, vertical=6),
            controls=[_version_block(t, v, it) for t, v, it in entries],
        ))

    log_dlg = ft.AlertDialog(
        open=False,
        title=ft.Text(f"更新日志　v{VERSION}", size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
        content=ft.Column(log_tiles, spacing=8, scroll=ft.ScrollMode.AUTO, width=_w(520), height=440),
        actions=[_tb("关闭", on_click=_close_log)],
        actions_alignment=ft.MainAxisAlignment.END,
    )

    btn_log = _eb(
        "📝 更新日志",
        height=34,
        bgcolor=C["gray"],
        color="white",
        on_click=open_changelog,
    )

    btn_manual = _eb(
        "📖 选型手册",
        height=34,
        bgcolor=C["banner"],
        color="white",
        on_click=open_manual_list,
    )

    # ---------- 封面外观主题切换弹窗 ----------
    _font_values = {v for v, _ in FONT_OPTIONS}
    _raw_font = cfg.get("font_family", "")
    _cur_font = _font_to_ui(_raw_font) if _raw_font in _font_values or _raw_font in ("", "system") else "system"

    def _pick_font(k):
        if k in _font_values:
            cfg["font_family"] = _font_to_store(k)
            try:
                _save_settings(cfg)
                _log(f"切换界面字体：{k or '默认'}")
            except Exception as ex:
                _log(f"字体保存失败：{ex}", "ERROR")
            theme_dlg.open = False
            if rebuild:
                rebuild(mode)

    def _font_tile(key, label):
        active = _cur_font == key
        ff = None if key == "system" else key
        return ft.Container(
            ft.Column([
                ft.Text(label, size=13.5, weight=ft.FontWeight.BOLD, font_family=ff,
                        text_align=ft.TextAlign.CENTER, color=C["text"]),
                ft.Text("当前使用中" if active else "点击切换", size=11,
                        color=C["muted"], text_align=ft.TextAlign.CENTER),
            ], spacing=6, horizontal_alignment=ft.CrossAxisAlignment.CENTER,
               alignment=ft.MainAxisAlignment.CENTER),
            bgcolor=C["card_alt"],
            border_radius=R_CARD,
            padding=10,
            width=108,
            height=76,
            border=ft.border.all(2 if active else 1, C["primary"] if active else C["border"]),
            on_click=lambda e, k=key: _pick_font(k),
            ink=True,
        )

    def _theme_tile(m):
        meta = THEME_META[m]
        active = mode == m
        return ft.Container(
            ft.Column([
                ft.Text(meta["icon"], size=26, text_align=ft.TextAlign.CENTER),
                ft.Text(meta["label"], size=13.5, weight=ft.FontWeight.BOLD,
                        text_align=ft.TextAlign.CENTER, color=C["text"]),
                ft.Text("当前使用中" if active else "点击切换", size=11,
                         color=C["muted"], text_align=ft.TextAlign.CENTER),
            ], spacing=6, horizontal_alignment=ft.CrossAxisAlignment.CENTER,
               alignment=ft.MainAxisAlignment.CENTER),
            bgcolor=C["card_alt"],
            border_radius=R_CARD,
            padding=14,
            width=136,
            height=120,
            border=ft.border.all(2 if active else 1, C["primary"] if active else C["border"]),
            on_click=lambda e: _pick_theme(m),
            ink=True,
        )

    def _pick_theme(m):
        if m == "custom":
            open_custom(None)
            return
        cfg["theme_mode"] = m
        try:
            _save_settings(cfg)
            _log(f"切换外观主题：{m}")
        except Exception as ex:
            _log(f"主题保存失败：{ex}", "ERROR")
        theme_dlg.open = False
        page.update()
        if rebuild:
            rebuild(m)

    theme_dlg = ft.AlertDialog(
        open=False,
        title=ft.Text("🎨 外观主题", size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
        content=ft.Container(
            ft.Column([
                ft.Row([_theme_tile(m) for m in _visible_modes(THEME_MODES)[:3]], spacing=14, wrap=True,
                       alignment=ft.MainAxisAlignment.CENTER),
                ft.Row([_theme_tile(m) for m in _visible_modes(THEME_MODES)[3:]], spacing=14, wrap=True,
                       alignment=ft.MainAxisAlignment.CENTER),
                ft.Divider(height=1, color=C["divider"]),
                ft.Text("界面字体", size=12.5, weight=ft.FontWeight.BOLD, color=C["title"]),
                ft.Row([_font_tile(k, lab) for k, lab in FONT_OPTIONS[:4]], spacing=12,
                       alignment=ft.MainAxisAlignment.CENTER, wrap=True),
                ft.Row([_font_tile(k, lab) for k, lab in FONT_OPTIONS[4:]], spacing=12,
                       alignment=ft.MainAxisAlignment.CENTER, wrap=True),
            ], spacing=14, horizontal_alignment=ft.CrossAxisAlignment.CENTER,
               width=_w(500), scroll=ft.ScrollMode.AUTO, height=440),
            **_dlg_bg,
        ),
    )

    def open_theme(e):
        page.dialog = theme_dlg
        theme_dlg.open = True
        page.update()

    btn_theme = _eb(
        "🎨 外观主题",
        height=34,
        bgcolor=C["gray"],
        color="white",
        on_click=open_theme,
    )

    # ---------- 自定义外观面板：5 项核心配色 + 四类背景图 ----------
    os.makedirs(CUSTOM_DIR, exist_ok=True)

    COLOR_FIELDS_DEF = [
        ("primary", "主色（按钮 / 链接 / 强调）",
         ["#1565C0", "#2E7D32", "#C62828", "#6A1B9A", "#EF6C00", "#00838F", "#37474F", "#000000"]),
        ("bg", "背景色（整窗底色）",
         ["#FFFFFF", "#F5F5F5", "#F3E5D8", "#E8F0E8", "#E3F2FD", "#121212", "#1E1E2E", "#263238"]),
        ("card", "卡片色（面板底色）",
         ["#FFFFFF", "#FAFAFA", "#FFF7ED", "#F1F8E9", "#ECEFF1", "#2C2C34", "#37474F", "#1E1E2E"]),
        ("title", "标题色",
         ["#1A1A1A", "#0D47A1", "#1B5E20", "#B71C1C", "#4A148C", "#FFFFFF", "#37474F", "#EF6C00"]),
        ("text", "正文色",
         ["#333333", "#555555", "#1A1A1A", "#263238", "#FFFFFF", "#424242", "#0D47A1", "#1B5E20"]),
    ]
    IMG_SLOTS = [
        ("bg", "整窗铺底背景"),
        ("cover", "主页封面横幅"),
        ("card", "卡片背板"),
        ("dialog", "弹窗背板"),
    ]

    cp_cur = cfg.get("custom_palette", {}) or {}
    c_color_fields = {}
    c_img_labels = {}
    _pick_target = {"key": None}

    def _img_status(key):
        p = _custom_img_path(cfg, key)
        return "已导入：" + os.path.basename(p) if p else "未设置（使用主题默认）"

    def _refresh_img_labels():
        for key, lab in c_img_labels.items():
            lab.value = _img_status(key)
            lab.update()

    def _on_pick_custom_img(e):
        key = _pick_target["key"]
        if not key or not e.files:
            return
        src = e.files[0].path
        if not src:
            return
        ext = os.path.splitext(src)[1].lower() or ".png"
        if ext not in (".jpg", ".jpeg", ".png", ".bmp", ".webp"):
            ext = ".png"
        dst_name = f"custom_{key}{ext}"
        dst = os.path.join(CUSTOM_DIR, dst_name)
        try:
            _copy_file(src, dst)
            cfg[f"custom_{key}_image"] = dst_name
            _log(f"自定义背景图已导入：{key} -> {dst_name}")
        except Exception as ex:
            _log(f"背景图导入失败：{ex}", "ERROR")
        _refresh_img_labels()
        page.update()

    cpicker = ft.FilePicker(on_result=_on_pick_custom_img)
    page.overlay.append(cpicker)

    def _swatch_row(colors, field):
        def _pick_swatch(ev):
            field.value = ev.control.data
            field.update()
        return ft.Row([
            ft.Container(width=26, height=26, border_radius=7, bgcolor=c,
                         border=ft.border.all(1, C["border"]), data=c,
                         on_click=_pick_swatch, ink=True)
            for c in colors
        ], spacing=6, wrap=True)

    color_rows = []
    for key, label, colors in COLOR_FIELDS_DEF:
        fld = _tf(label=label, value=_norm_color(cp_cur.get(key, ""), ""), width=_w(150), height=40)
        c_color_fields[key] = fld
        color_rows.append(ft.Row([fld, _swatch_row(colors, fld)],
                                 spacing=12, wrap=True, vertical_alignment=ft.CrossAxisAlignment.CENTER))

    img_rows = []
    for key, label in IMG_SLOTS:
        lab = ft.Text(_img_status(key), size=11.5, color=C["muted"])
        c_img_labels[key] = lab

        def _open_pick(ev, k=key):
            if not _is_desktop(page):
                page.show_snack_bar(ft.SnackBar(
                    content=ft.Text("当前环境暂不支持自定义背景图，请在桌面版设置")))
                return
            _pick_target["key"] = k
            cpicker.pick_files(dialog_title=f"选择{label}图片",
                               allowed_extensions=["jpg", "jpeg", "png", "bmp", "webp"])

        def _clear_img(ev, k=key):
            p = _custom_img_path(cfg, k)
            if p:
                try:
                    os.remove(p)
                except Exception:
                    pass
            cfg[f"custom_{k}_image"] = ""
            _refresh_img_labels()
            page.update()

        img_rows.append(ft.Row([
            ft.Text(label, size=12.5, width=110, color=C["text"]),
            ft.Column([lab], spacing=2, expand=True),
            _tb("导入", on_click=_open_pick),
            _tb("移除", on_click=_clear_img),
        ], spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER))

    def _save_custom(ev):
        new_pal = {}
        for key, _, _ in COLOR_FIELDS_DEF:
            col = _norm_color(c_color_fields[key].value, "")
            if col:
                new_pal[key] = col
        cfg["custom_palette"] = new_pal
        cfg["theme_mode"] = "custom"
        try:
            _save_settings(cfg)
            _log("自定义外观已保存并应用")
        except Exception as ex:
            _log(f"自定义外观保存失败：{ex}", "ERROR")
        if rebuild:
            custom_dlg.open = False
            page.update()
            rebuild("custom")

    def _reset_custom(ev):
        cfg["custom_palette"] = {}
        for key, _ in IMG_SLOTS:
            p = _custom_img_path(cfg, key)
            if p:
                try:
                    os.remove(p)
                except Exception:
                    pass
            cfg[f"custom_{key}_image"] = ""
        try:
            _save_settings(cfg)
        except Exception:
            pass
        if rebuild:
            custom_dlg.open = False
            page.update()
            rebuild("custom")

    custom_dlg = ft.AlertDialog(
        open=False,
        title=ft.Text("🎨 自定义外观", size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
        content=ft.Container(
            ft.Column(color_rows + [ft.Divider(color=C["divider"])] + img_rows,
                      spacing=12, width=_w(520), scroll=ft.ScrollMode.AUTO),
            padding=ft.padding.only(top=6, left=4, right=4),
            **_dlg_bg,
        ),
        actions=[
            _tb("恢复默认", on_click=_reset_custom),
            _tb("取消", on_click=lambda ev: _close_dlg(custom_dlg)),
            _tb("保存并应用", on_click=_save_custom),
        ],
        actions_alignment=ft.MainAxisAlignment.END,
    )

    def open_custom(ev):
        page.dialog = custom_dlg
        custom_dlg.open = True
        page.update()

    # ---------- 设置弹窗：默认值 / BOM 模板 / 导出列 / 数据版本 ----------
    s_volt = _dd(label="默认供电电压", options=opt("AC220", "AC380"),
                         value=cfg.get("default_voltage", "AC220"), width=_w(205))
    s_depth = _dd(
        label="默认线缆长度", width=_w(205),
        options=[ft.dropdown.Option(v, t) for v, t in [
            ("3.0", "3 m（-3.0）"), ("5.0", "5 m（-5.0）"), ("10.0", "10 m（-10.0）")]],
        value=cfg.get("default_depth", "3.0"))
    s_flex = _cb(label="默认高柔线（-T）", value=bool(cfg.get("default_high_flex", True)))
    s_format = _dd(
        label="BOM 导出格式", width=_w(205),
        options=[ft.dropdown.Option("xlsx", "Excel（.xlsx）"), ft.dropdown.Option("csv", "CSV（.csv）")],
        value=cfg.get("bom_format", "xlsx"))
    s_title = _tf(label="清单标题", value=cfg.get("bom_title", ""), width=_w(205))
    s_company = _tf(label="公司抬头（可留空）", value=cfg.get("bom_company", ""), width=_w(205))
    s_footer = _tf(label="落款（可留空）", value=cfg.get("bom_footer", ""), width=_w(205))
    s_col_type = _cb(label="类型", value=bool(cfg.get("col_type", True)))
    s_col_model = _cb(label="型号", value=bool(cfg.get("col_model", True)))
    s_col_qty = _cb(label="数量", value=bool(cfg.get("col_qty", True)))
    s_col_note = _cb(label="备注", value=bool(cfg.get("col_note", True)))
    s_data_ver = _tf(label="数据版本（手册版本号，可留空）", value=cfg.get("data_version", ""), width=_w(205))
    s_data_date = _tf(label="数据日期（YYYY-MM-DD，可留空）", value=cfg.get("data_date", ""), width=_w(205))
    _cur_mode = cfg.get("theme_mode") if cfg.get("theme_mode") in _visible_modes(THEME_MODES) else "light"
    s_theme = _dd(
        label="外观主题", width=_w(205),
        options=[ft.dropdown.Option(m, THEME_META[m]["label"]) for m in _visible_modes(THEME_MODES)],
        value=_cur_mode)

    def _save_settings_dlg(ev):
        cfg["default_voltage"] = s_volt.value
        cfg["default_depth"] = s_depth.value
        cfg["default_high_flex"] = s_flex.value
        cfg["watermark"] = "姓许的苏剑"
        cfg["bom_format"] = s_format.value
        cfg["bom_title"] = s_title.value.strip() or "伺服电机与驱动器配置清单"
        cfg["bom_company"] = s_company.value.strip()
        cfg["bom_footer"] = s_footer.value.strip()
        cfg["col_type"] = s_col_type.value
        cfg["col_model"] = s_col_model.value
        cfg["col_qty"] = s_col_qty.value
        cfg["col_note"] = s_col_note.value
        cfg["data_version"] = s_data_ver.value.strip()
        cfg["data_date"] = s_data_date.value.strip()
        new_mode = s_theme.value if s_theme.value in _visible_modes(THEME_MODES) else _cur_mode
        theme_changed = new_mode != _cur_mode
        cfg["theme_mode"] = new_mode
        try:
            _save_settings(cfg)
            _log("设置已保存")
        except Exception as ex:
            _log(f"设置保存失败：{ex}", "ERROR")
        if theme_changed and rebuild:
            settings_dlg.open = False
            rebuild(new_mode)
            return
        volt_drop.value = cfg["default_voltage"]
        depth_drop.value = cfg["default_depth"]
        flex_check.value = cfg["default_high_flex"]
        footer_text.value = _footer_text()
        settings_dlg.open = False
        refresh(None)
        page.update()

    settings_dlg = ft.AlertDialog(
        open=False,
        title=ft.Text("⚙️ 设置", size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
        content=ft.Container(
            ft.Column([
                ft.Row([s_volt, s_depth], spacing=12, wrap=True),
                s_flex,
                ft.Row([s_format, s_theme], spacing=12, wrap=True),
                ft.Row([s_title, s_company], spacing=12, wrap=True),
                s_footer,
                ft.Text("导出列：", size=12, color=C["muted"]),
                ft.Row([s_col_type, s_col_model, s_col_qty, s_col_note], spacing=16, wrap=True),
                ft.Row([s_data_ver, s_data_date], spacing=12, wrap=True),
            ], spacing=10, width=_w(480)),
            padding=ft.padding.only(top=6, left=4, right=4),
            **_dlg_bg,
        ),
        actions=[
            _tb("取消", on_click=lambda ev: _close_dlg(settings_dlg)),
            _tb("保存", on_click=_save_settings_dlg),
        ],
        actions_alignment=ft.MainAxisAlignment.END,
    )

    def open_settings(e):
        page.dialog = settings_dlg
        settings_dlg.open = True
        page.update()

    btn_settings = _eb(
        "⚙️ 设置",
        height=34,
        bgcolor=C["gray"],
        color="white",
        on_click=open_settings,
    )

    # ---------- 导出记录弹窗 ----------
    btn_clear_log = _tb("清空全部", on_click=None)
    btn_open_log_folder = _tb("打开清单文件夹", on_click=None)
    export_log_dlg = ft.AlertDialog(
        open=False,
        title=ft.Text("📋 导出记录", size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
        content=ft.Column([ft.Text("")], width=_w(480), height=440),
        actions=[
            btn_open_log_folder,
            btn_clear_log,
            _tb("关闭", on_click=lambda ev: _close_dlg(export_log_dlg)),
        ],
        actions_alignment=ft.MainAxisAlignment.END,
    )
    clear_armed = {"on": False}
    log_scope = {"v": "bom"}
    _EXP_SCOPE = {
        "bom": {"log": EXPORT_LOG, "title": "📋 选型清单导出记录", "folder": open_bom_folder, "folder_label": "打开选型清单"},
        "repl": {"log": REPL_EXPORT_LOG, "title": "📋 替换清单导出记录", "folder": open_repl_folder, "folder_label": "打开替换清单"},
    }

    def _delete_export_entry(idx):
        sc = _EXP_SCOPE[log_scope["v"]]
        entries = _read_export_log(sc["log"])
        removed = ""
        if 0 <= idx < len(entries):
            removed = str(entries[idx].get("文件", ""))
            entries.pop(idx)
            _write_export_log(entries, sc["log"])
            _log(f"删除导出记录：{removed}")
        open_export_log(None)

    def _clear_export_log(ev):
        if not clear_armed["on"]:
            clear_armed["on"] = True
            btn_clear_log.text = "再次点击确认清空"
            btn_clear_log.update()
            return
        _write_export_log([], _EXP_SCOPE[log_scope["v"]]["log"])
        clear_armed["on"] = False
        _log("清空全部导出记录")
        open_export_log(None)

    btn_clear_log.on_click = _clear_export_log
    btn_open_log_folder.on_click = lambda e: _EXP_SCOPE[log_scope["v"]]["folder"](e)

    def open_export_log(e, scope="bom"):
        log_scope["v"] = scope
        sc = _EXP_SCOPE[scope]
        clear_armed["on"] = False
        btn_clear_log.text = "清空全部"
        btn_open_log_folder.text = sc["folder_label"]
        export_log_dlg.title = ft.Text(sc["title"], size=16, weight=ft.FontWeight.BOLD, color=C["title"])
        entries = _read_export_log(sc["log"])
        rows = []
        if not entries:
            rows.append(ft.Text("暂无导出记录", size=12.5, color=C["muted"]))
        for idx, en in reversed(list(enumerate(entries))):
            sub = en.get("目录", "") or ""
            fname = en.get("文件", "")
            file_label = f"{sub}/{fname}" if sub else fname
            rows.append(ft.Column([
                ft.Row([
                    ft.Text(en.get("时间", ""), size=11, color=C["muted"]),
                    ft.Text(en.get("格式", ""), size=11, color=C["link"]),
                    ft.Text(f"{en.get('套数', 0)} 套", size=11, color=C["muted"]),
                    ft.IconButton(icon=ft.icons.DELETE_OUTLINE, icon_size=16, tooltip="删除该记录",
                                  on_click=lambda ev, i=idx: _delete_export_entry(i)),
                 ], spacing=10, alignment=ft.MainAxisAlignment.SPACE_BETWEEN, wrap=True),
                ft.Text(file_label, size=12, color=C["text"]),
                ft.Text(en.get("条件", "") or "无条件", size=11, color=C["muted"]),
            ], spacing=2))
        export_log_dlg.content = ft.Column(rows, spacing=10, scroll=ft.ScrollMode.AUTO, width=_w(480), height=440)
        page.dialog = export_log_dlg
        export_log_dlg.open = True
        page.update()

    btn_export_log = _eb(
        "📋 导出记录",
        height=34,
        bgcolor=C["gray"],
        color="white",
        on_click=lambda e: open_export_log(e, "bom"),
    )

    result_view = ft.Column(spacing=10)
    more_box = ft.Container()
    count_text = ft.Text("", size=13, color=C["muted"])
    tip_text = ft.Text("", size=12, color=C["success"])
    state = {"count": 0, "res": [], "shown": 0}
    BATCH = 30
    selected = {}

    def _sel_key(it):
        return f"{it.get('motor_model')}|{it.get('drive_model')}"

    def _conditions_summary():
        parts = []
        if volt_drop.value != "全部":
            parts.append(f"电压{volt_drop.value}")
        if power_drop.value != "全部":
            parts.append(f"功率{power_drop.value}")
        if speed_drop.value != "全部":
            parts.append(f"转速{speed_drop.value}")
        if series_drop.value != "全部":
            parts.append(f"惯量{series_drop.value.split('｜')[0]}")
        if flange_drop.value != "全部":
            parts.append(f"法兰{flange_drop.value}")
        if driver_drop.value != "全部":
            parts.append(f"驱动{driver_drop.value}")
        if comm_drop.value != "全部":
            parts.append(f"通讯{comm_drop.value}")
        if enc_drop.value != "全部":
            parts.append(f"编码器{enc_drop.value}")
        if battery_drop.value != "全部":
            parts.append(f"电池{battery_drop.value}")
        if brake_check.value:
            parts.append("带抱闸")
        if oilseal_check.value:
            parts.append("带油封")
        return "；".join(parts)

    def _refresh_sel_btn():
        btn_export_sel.text = f"⬇ 导出勾选的 BOM（{len(selected)}）"
        page.update()

    def _toggle_select(it, checked):
        k = _sel_key(it)
        if checked:
            selected[k] = it
        else:
            selected.pop(k, None)
        _refresh_sel_btn()

    def on_export(e, it):
        if not _is_desktop(page):
            tip_text.value = "当前环境暂不支持导出文件，请使用桌面版导出"
            tip_text.color = C["warn"]
            page.update()
            return
        try:
            path, n = export_bom(it, depth_drop.value, flex_check.value, cfg, base_dir,
                                 conditions_txt=_conditions_summary(),
                                 name_hint=str(it.get("motor_model") or "清单"))
            tip_text.value = f"已导出：{path}"
            tip_text.color = C["success"]
        except Exception as ex:
            _log(f"导出失败：{ex}", "ERROR")
            tip_text.value = f"导出失败：{ex}"
            tip_text.color = C["danger"]
        page.update()

    def on_export_selected(e):
        if not _is_desktop(page):
            tip_text.value = "当前环境暂不支持导出文件，请使用桌面版导出"
            tip_text.color = C["warn"]
            page.update()
            return
        items = list(selected.values())
        if not items:
            tip_text.value = "请先勾选要导出的搭配"
            tip_text.color = C["warn"]
            page.update()
            return
        try:
            path, n = export_bom(items, depth_drop.value, flex_check.value, cfg, base_dir,
                                conditions_txt=_conditions_summary())
            tip_text.value = f"已导出 {n} 套搭配：{path}"
            tip_text.color = C["success"]
        except Exception as ex:
            _log(f"合并导出失败：{ex}", "ERROR")
            tip_text.value = f"导出失败：{ex}"
            tip_text.color = C["danger"]
        page.update()

    btn_export_sel = _eb(
        "⬇ 导出勾选的 BOM（0）",
        height=34,
        bgcolor=C["success"],
        color="white",
        on_click=on_export_selected,
    )

    def on_clear_selection(e):
        selected.clear()
        refresh(None)
        _refresh_sel_btn()

    btn_clear_sel = _eb(
        "✖ 取消勾选",
        height=34,
        bgcolor=C["gray"],
        color="white",
        on_click=on_clear_selection,
    )

    def on_goto_repl(e):
        items = list(selected.values()) if selected else (state.get("res") or [])
        if items:
            repl_groups[:] = _repl_build_groups(items)
            _rebuild_repl_list()
            repl_tip.value = f"已带入 {len(items)} 套选型搭配；初始型号列请照抄客户铭牌。"
            repl_tip.color = C["success"]
        else:
            repl_tip.value = "暂无已选搭配，可在替换页手填汇川型号。"
            repl_tip.color = C["warn"]
        _open_repl(e)

    btn_to_repl = _eb(
        "📋 生成替换对照清单",
        height=34,
        bgcolor=C["primary"],
        color="white",
        on_click=on_goto_repl,
    )

    def on_reset_selector(e):
        volt_drop.value = cfg.get("default_voltage", "AC220")
        power_drop.value = "全部"
        speed_drop.value = "全部"
        series_drop.value = "全部"
        flange_drop.value = "全部"
        driver_drop.value = "全部"
        comm_drop.value = "全部"
        enc_drop.value = "全部"
        magnet_drop.value = "否"
        battery_drop.value = "需要"
        depth_drop.value = cfg.get("default_depth", "3.0")
        flex_check.value = bool(cfg.get("default_high_flex", True))
        brake_check.value = False
        oilseal_check.value = True
        selected.clear()
        refresh(None)
        _refresh_sel_btn()
        tip_text.value = "已恢复为默认选型条件"
        tip_text.color = C["muted"]
        page.update()

    btn_reset = _eb(
        "↺ 一键回原",
        height=34,
        bgcolor=C["gray"],
        color="white",
        on_click=on_reset_selector,
    )

    def diff_score(d):
        """给一条记录打“接近程度”分：越小越接近用户条件。返回 (分数, 差异说明)"""
        if volt_drop.value != "全部" and d.get("voltage") != volt_drop.value:
            return None
        pen, notes = 0.0, []
        if power_drop.value != "全部":
            req, act = _kw(power_drop.value), _kw(d.get("power_text"))
            if act < req - 1e-9:
                pen += 3000 + (req - act) / max(req, 1e-9) * 1000
                notes.append(f"功率{act:g}kW低于要求{req:g}kW")
            elif act > req + 1e-9:
                pen += (act - req) / max(req, 1e-9) * 300
                notes.append(f"功率 {req:g}kW→{act:g}kW（偏大可用）")
        if flange_drop.value != "全部":
            req, act = _flange_mm(flange_drop.value), _flange_mm(d.get("flange_size"))
            if abs(act - req) > 1e-9:
                pen += 400 + abs(act - req) / 20.0 * 80
                notes.append(f"法兰 {flange_drop.value}→{d.get('flange_size') or '待补'}")
        if speed_drop.value != "全部" and d.get("rated_speed") != speed_drop.value:
            pen += 200
            notes.append(f"转速 {speed_drop.value}→{d.get('rated_speed')}")
        if series_drop.value != "全部":
            pick = series_drop.value.split("｜")[0].strip()
            if d.get("series_code") != pick:
                pen += 150
                notes.append(f"惯量 {pick}→{d.get('series_code')}")
        if driver_drop.value != "全部" and not str(d.get("drive_model") or "").startswith(driver_drop.value):
            pen += 500
            notes.append(f"驱动 {driver_drop.value}→{str(d.get('drive_model') or '')}")
        if comm_drop.value != "全部" and d.get("comm") != comm_drop.value:
            pen += 400
            notes.append(f"通讯 {comm_drop.value}→{d.get('comm')}")
        if enc_drop.value != "全部" and d.get("encoder_short") != enc_drop.value:
            pen += 400
            notes.append(f"编码器 {enc_drop.value}→{d.get('encoder_short')}")
        if magnet_drop.value != "全部" and d.get("magnet_encoder") != magnet_drop.value:
            pen += 250
            notes.append("编码器形式不同")
        bat = str(d.get("battery") or "")
        if battery_drop.value == "需要" and bat != "S6-C4A":
            pen += 250
            notes.append("无电池套件")
        if battery_drop.value == "免电池" and not bat.startswith("无"):
            pen += 250
            notes.append("需配电池")
        if brake_check.value and d.get("has_brake") != "是":
            pen += 900
            notes.append("无抱闸（要求带抱闸）")
        if not brake_check.value and d.get("has_brake") == "是":
            pen += 200
            notes.append("带抱闸（要求不带）")
        if oilseal_check.value and d.get("has_oilseal") != "是":
            pen += 120
            notes.append("无油封（要求带油封）")
        if not oilseal_check.value and d.get("has_oilseal") == "是":
            pen += 120
            notes.append("带油封（要求不带）")
        return pen, notes

    def recommend(limit=5):
        """无完全匹配时，按“最接近可用”的程度给出推荐搭配"""
        scored, overflow = [], []
        for d in data:
            r = diff_score(d)
            if r is None:
                continue
            pen, notes = r
            (scored if pen < 3000 else overflow).append((pen, notes, d))
        scored.sort(key=lambda x: x[0])
        if not scored:
            scored = sorted(overflow, key=lambda x: x[0])
        picked, seen = [], set()
        for pen, notes, d in scored:
            key = str(d.get("motor_model"))
            if key in seen:
                continue
            seen.add(key)
            picked.append((pen, notes, d))
            if len(picked) >= limit:
                break
        return picked

    def _diagnose_empty():
        """逐步过滤，定位把候选筛没的瓶颈条件并给放宽建议"""
        pool = list(data)
        steps = []

        def step(label, value, pred, suggest):
            nonlocal pool
            if value == "全部":
                return
            before = len(pool)
            pool = [d for d in pool if pred(d)]
            steps.append((label, str(value), before, len(pool), suggest))

        step("供电电压", volt_drop.value, lambda d: d.get("voltage") == volt_drop.value,
             "另一档 AC220 / AC380")
        step("额定功率", power_drop.value, lambda d: _kw(d.get("power_text")) == _kw(power_drop.value), None)
        step("额定转速", speed_drop.value, lambda d: d.get("rated_speed") == speed_drop.value,
             "1500 / 2000 / 3000rpm")
        if series_drop.value != "全部":
            pick = series_drop.value.split("｜")[0].strip()
            step("惯量&容量", series_drop.value, lambda d: d.get("series_code") == pick, "放宽为「全部」")
        step("法兰尺寸", flange_drop.value, lambda d: str(d.get("flange_size") or "").strip() == flange_drop.value, None)
        if driver_drop.value != "全部":
            step("驱动器系列", driver_drop.value,
                 lambda d: str(d.get("drive_model") or "").startswith(driver_drop.value), "放宽为「全部」")
        step("通讯类型", comm_drop.value, lambda d: d.get("comm") == comm_drop.value, "放宽为「全部」")
        step("编码器", enc_drop.value, lambda d: d.get("encoder_short") == enc_drop.value, "放宽为「全部」")
        if magnet_drop.value != "全部":
            step("编码器形式", magnet_drop.value, lambda d: d.get("magnet_encoder") == magnet_drop.value, "放宽为「全部」")
        if battery_drop.value == "需要":
            step("电池", battery_drop.value, lambda d: str(d.get("battery") or "") == "S6-C4A", "放宽为「全部」")
        elif battery_drop.value == "免电池":
            step("电池", battery_drop.value, lambda d: str(d.get("battery") or "").startswith("无"), "放宽为「全部」")
        if brake_check.value:
            step("抱闸", "需要抱闸", lambda d: d.get("has_brake") == "是", "取消「需要抱闸」勾选")
        if oilseal_check.value:
            step("油封", "带油封", lambda d: d.get("has_oilseal") == "是", "取消「带油封」勾选")

        bad = [s for s in steps if s[3] == 0]
        if not bad:
            return "未找到完全匹配，可适当放宽筛选条件。"
        label, value, before, after, suggest = bad[0]
        if label == "额定功率":
            req = _kw(power_drop.value)
            lower = [p for p in power_vals if _kw(p) < req]
            upper = [p for p in power_vals if _kw(p) > req]
            cand = ([lower[-1]] if lower else []) + ([upper[0]] if upper else [])
            sug = "可试 " + " / ".join(cand) if cand else "放宽为「全部」"
        elif label == "法兰尺寸":
            sug = "库内法兰：" + " / ".join(flange_vals[:6])
        else:
            sug = suggest or "放宽为「全部」"
        return f"最可能是「{label}：{value}」把结果筛没了；建议：{sug}。"

    def card_of(it, diff=None):
        enc_short = it.get("encoder_short") or "待补"
        enc_type = it.get("encoder_type") or ""
        magnet = "免电池磁编" if it.get("magnet_encoder") == "是" else "传统光编"
        power = it.get("power_text") or "待补"
        speed = it.get("rated_speed") or "待补"
        flange = it.get("flange_size") or "待补"
        power_line = build_line(it.get("power_line_base"), depth_drop.value, flex_check.value) or "待补"
        enc_line = build_line(it.get("encoder_line_base"), depth_drop.value, flex_check.value) or "待补"

        tag = ft.Container(
            ft.Text("推荐搭配" if diff else "完全匹配", size=10.5,
                    color=C["warn"] if diff else C["success"]),
            bgcolor=C["tag_warn_bg"] if diff else C["tag_ok_bg"],
            border_radius=R_BADGE,
            padding=ft.padding.symmetric(horizontal=9, vertical=3),
        )

        def kv(k, v):
            return ft.Row([
                ft.Text(k, size=12, color=C["muted"], width=64),
                ft.Text(str(v), size=12, color=C["text"], expand=True),
            ], spacing=4)

        left = ft.Column([
            kv("驱动器", it.get("drive_model")),
            kv("法兰", flange),
            kv("功率", power),
            kv("编码器", f"{enc_short}｜{enc_type}"),
        ], spacing=4, expand=True)
        right = ft.Column([
            kv("转速", speed),
            kv("动力线", power_line),
            kv("编码线", enc_line),
            kv("CN1", it.get("cn1")),
        ], spacing=4, expand=True)

        head_rows = [
            ft.Row([
                ft.Text(it.get("motor_model") or "待补", size=14.5,
                        weight=ft.FontWeight.BOLD, color=C["title"], expand=True),
                tag,
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
        ]
        if diff:
            head_rows.append(ft.Text(f"差异：{diff}", size=11.5, color=C["warn"], italic=True))

        return ft.Container(
            ft.Column([
                *head_rows,
                ft.Divider(height=8, color=C["divider"]),
                ft.Row([left, right], spacing=18),
                ft.Row([
                    _cb(
                        value=selected.get(_sel_key(it), False),
                        on_change=lambda ev, x=it: _toggle_select(x, ev.control.value),
                    ),
                    _eb(
                        "⬇ 导出这份 BOM 清单",
                        height=34,
                        bgcolor=C["success"],
                        color="white",
                        on_click=lambda e, x=it: on_export(e, x),
                    ),
                    _eb(
                        "✏️ 纠错",
                        height=34,
                        bgcolor=C["primary"],
                        color="white",
                        on_click=lambda e, x=it: _open_fix(e, x),
                    ),
                    ft.Text(cfg.get("watermark", "姓许的苏剑"), size=10.5, color=C["muted"]),
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN, wrap=True),
            ], spacing=6),
            bgcolor=C["card"],
            border=ft.border.all(1, C["border"]),
            border_radius=R_CARD,
            padding=14,
        )

    def refresh(e):
        res = []
        for d in data:
            if volt_drop.value != "全部" and d.get("voltage") != volt_drop.value:
                continue
            if power_drop.value != "全部" and _kw(d.get("power_text")) != _kw(power_drop.value):
                continue
            if speed_drop.value != "全部" and d.get("rated_speed") != speed_drop.value:
                continue
            if series_drop.value != "全部":
                pick = series_drop.value.split("｜")[0].strip()
                if d.get("series_code") != pick:
                    continue
            if flange_drop.value != "全部" and str(d.get("flange_size") or "").strip() != flange_drop.value:
                continue
            if driver_drop.value != "全部":
                drive_model = str(d.get("drive_model") or "").strip()
                if not drive_model.startswith(driver_drop.value):
                    continue
            if comm_drop.value != "全部" and d.get("comm") != comm_drop.value:
                continue
            if enc_drop.value != "全部" and d.get("encoder_short") != enc_drop.value:
                continue
            if magnet_drop.value != "全部" and d.get("magnet_encoder") != magnet_drop.value:
                continue
            bat = str(d.get("battery") or "")
            if battery_drop.value == "需要" and bat != "S6-C4A":
                continue
            if battery_drop.value == "免电池" and not bat.startswith("无"):
                continue
            if brake_check.value and d.get("has_brake") != "是":
                continue
            if not brake_check.value and d.get("has_brake") == "是":
                continue
            if oilseal_check.value and d.get("has_oilseal") != "是":
                continue
            if not oilseal_check.value and d.get("has_oilseal") == "是":
                continue
            res.append(d)

        state["count"] = len(res)
        state["res"] = res
        state["shown"] = min(BATCH, len(res))
        result_view.controls.clear()
        more_box.content = None
        if res:
            count_text.value = f"匹配 {len(res)} 条"
            count_text.color = C["muted"]
            for it in res[:state["shown"]]:
                result_view.controls.append(card_of(it))
            _refresh_more_btn()
        else:
            diag = _diagnose_empty()
            rec = recommend()
            if not rec:
                count_text.value = "匹配 0 条"
                result_view.controls.append(ft.Text(diag, color=C["danger"]))
            else:
                count_text.value = f"完全匹配 0 条 · 已按最接近条件推荐 {len(rec)} 款"
                count_text.color = C["warn"]
                result_view.controls.append(ft.Text(diag, color=C["warn"]))
                result_view.controls.append(ft.Text(
                    "没有完全符合的型号，以下为最接近且可用的搭配（按接近程度排序）：", color=C["warn"]))
                for pen, notes, it in rec:
                    result_view.controls.append(card_of(it, "；".join(notes) if notes else "无差异"))
        page.update()

    def _load_more(e):
        res = state["res"]
        nxt = min(state["shown"] + BATCH, len(res))
        for it in res[state["shown"]:nxt]:
            result_view.controls.append(card_of(it))
        state["shown"] = nxt
        _refresh_more_btn()
        page.update()

    def _refresh_more_btn():
        res = state["res"]
        if state["shown"] < len(res):
            more_box.content = ft.Row(
                [_eb(
                    f"⬇ 加载更多（剩余 {len(res) - state['shown']} 条）",
                    height=34, bgcolor=C["chip"], color=C["title"],
                    on_click=_load_more)],
                alignment=ft.MainAxisAlignment.CENTER)
        else:
            more_box.content = None

    def _is_empty():
        return state["count"] == 0

    def on_driver_change(e):
        v = driver_drop.value
        if v != "全部" and v in SERIES_PRESET:
            comm, enc = SERIES_PRESET[v]
            comm_drop.value = comm
            enc_drop.value = enc
            refresh(e)
            # 预设编码器若把结果筛空，说明该系列不是这个编码器，自动回落为“全部”
            if _is_empty():
                enc_drop.value = "全部"
                refresh(e)
            # 仍为空：该系列可能只有免电池磁编机型，放开编码器形式与电池限制
            if _is_empty():
                magnet_drop.value = "全部"
                battery_drop.value = "全部"
                refresh(e)
            return
        refresh(e)

    def on_manual_change(e):
        # 手动改通讯/编码器时不再清空驱动器系列，保留用户已选的系列（如 SV660N + 传统光编）
        refresh(e)

    def on_magnet_change(e):
        # 编码器形式 ↔ 绝对值电池 联动：磁编=不配电池，光编=需电池；驱动器系列保持不变
        battery_drop.value = {"是": "免电池", "否": "需要"}.get(magnet_drop.value, "全部")
        refresh(e)

    def on_battery_change(e):
        magnet_drop.value = {"免电池": "是", "需要": "否"}.get(battery_drop.value, "全部")
        refresh(e)

    driver_drop.on_change = on_driver_change
    comm_drop.on_change = on_manual_change
    enc_drop.on_change = on_manual_change
    magnet_drop.on_change = on_magnet_change
    battery_drop.on_change = on_battery_change
    for ctrl in [volt_drop, power_drop, speed_drop, series_drop, flange_drop,
                 depth_drop, brake_check, oilseal_check, flex_check]:
        ctrl.on_change = refresh

    # ---------- 界面美化：统一控件宽度 + 卡片化布局 ----------
    for c in [volt_drop, power_drop, speed_drop, series_drop, flange_drop,
              driver_drop, comm_drop, enc_drop, magnet_drop, battery_drop, depth_drop]:
        c.width = _w(205)

    def section_title(txt):
        return ft.Text(txt, size=13.5, weight=ft.FontWeight.BOLD, color=C["title"])

    # ---------- LOGO：把 logo.png / logo.jpg 放在 main.py 同目录即可自动显示 ----------
    def find_logo():
        for d in (base_dir, RES_DIR):
            for ext in (".png", ".jpg", ".jpeg", ".webp", ".bmp"):
                fp = os.path.join(d, "logo" + ext)
                if os.path.exists(fp):
                    return _asset("logo" + ext) if web else fp
        return None

    logo_path = find_logo()
    if logo_path:
        logo_widget = ft.Image(src=logo_path, width=40, height=40, fit=ft.ImageFit.CONTAIN)
    else:
        logo_widget = ft.Container(
            ft.Text("MS", size=15, weight=ft.FontWeight.BOLD, color="white"),
            width=40, height=40, bgcolor=C["primary"], border_radius=R_BADGE,
            alignment=ft.alignment.center,
        )

    header = ft.Container(
        ft.Row([
            ft.Container(
                logo_widget,
                bgcolor=C["card"],
                border_radius=R_BADGE,
                padding=3,
            ),
            ft.Column([
                ft.Text("汇川 MS1-R 伺服选型工具", size=21, weight=ft.FontWeight.BOLD, color="white"),
                ft.Text(f"SV600EN/EP、SV605N/P / SV630 / SV660 / SV670 / SV680　·　数据库 {len(data)} 条　·　惯量 / 法兰 / 编码器联合筛选　·　v{VERSION}",
                        size=11.5, color=C["banner_sub"]),
            ], spacing=2, expand=True),
            _eb(content=ft.Text("主页", weight=ft.FontWeight.BOLD, color=C["title"]),
                               height=34, bgcolor=C["card"], on_click=_go_home),
        ], spacing=14, alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
        bgcolor=C["banner"],
        padding=ft.padding.symmetric(horizontal=20, vertical=15),
    )

    filter_card = ft.Container(
        ft.Column([
            section_title("①  选型条件"),
            ft.Row([volt_drop, power_drop, speed_drop, series_drop], spacing=12, wrap=True),
            ft.Row([flange_drop, driver_drop, comm_drop, enc_drop], spacing=12, wrap=True),
            ft.Row([magnet_drop, battery_drop, depth_drop], spacing=12, wrap=True),
            ft.Row([brake_check, oilseal_check, flex_check], spacing=12, wrap=True),
            ft.Row(
                [
                    btn_settings,
                    ft.Row([btn_reset, btn_cable, btn_manual, btn_export_log, btn_hist, btn_open_folder, btn_hint_sel],
                           spacing=12, wrap=True),
                ],
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                wrap=True,
            ),
            ft.Divider(height=10, color=C["divider"]),
            ft.Text("提示：惯量 & 容量、法兰尺寸默认「全部」；没有任何型号完全符合时，会自动给出最接近的推荐搭配。",
                    size=11.5, color=C["muted"]),
        ], spacing=10),
        bgcolor=C["card"],
        border_radius=R_CARD,
        padding=16,
        border=ft.border.all(1, C["border"]),
    )

    result_head = ft.Row([
        section_title("②  选型结果"),
        ft.Container(
            count_text,
            bgcolor=C["chip"],
            border_radius=10,
            padding=ft.padding.symmetric(horizontal=10, vertical=3),
        ),
        ft.Row([btn_clear_sel, btn_export_sel, btn_to_repl], spacing=10, wrap=True),
     ], spacing=10, alignment=ft.MainAxisAlignment.SPACE_BETWEEN, wrap=True)

    action_hint = ft.Text("选好型号后，点结果卡片最下方的【⬇ 导出这份 BOM 清单】即可生成 Excel 采购清单",
                          size=11.5, color=C["link"])

    def _footer_text():
        tag = ""
        if cfg.get("data_version") or cfg.get("data_date"):
            tag = f"　·　数据版本 {cfg.get('data_version') or '—'}（{cfg.get('data_date') or '—'}）"
        return f"{cfg.get('watermark', '姓许的苏剑')}　·　汇川 MS1-R 伺服选型工具　·　数据依据 MS1-R 系列伺服电机选型手册{tag}"

    footer_text = ft.Text(_footer_text(), size=11, color=C["muted"])
    footer = ft.Container(
        footer_text,
        alignment=ft.alignment.center,
        padding=ft.padding.symmetric(vertical=12),
    )

    page.bgcolor = C["bg"]
    page.padding = 0

    def _back_to_top(e):
        page.scroll_to(offset=0, duration=300, curve=ft.AnimationCurve.EASE_IN_OUT)

    back_top_btn = ft.FloatingActionButton(
        icon=ft.icons.ARROW_UPWARD,
        tooltip="回到顶部",
        mini=True,
        bgcolor=C["primary"],
        foreground_color="white",
        on_click=_back_to_top,
    )

    selector_view = ft.Column([
        header,
        ft.Container(
            ft.Column([
                filter_card,
                result_head,
                action_hint,
                tip_text,
                result_view,
                more_box,
            ], spacing=12),
            padding=ft.padding.all(18),
        ),
        footer,
    ], spacing=0, tight=True)

    # ---------- 故障码 / 警告码查询视图 ----------
    def _open_fault(e):
        page.floating_action_button = back_top_btn
        _show(fault_view)
        _show_hint("fault")

    def _fault_card(series, code, rec):
        fault_rows = rec.get("rows", [])
        meta = [series]
        if rec.get("level"):
            meta.append(f"级别 {rec['level']}")
        if rec.get("reset"):
            meta.append(f"可复位：{rec['reset']}")
        children = [
            ft.Row([
                ft.Text(f"{code}：{rec.get('title', '')}", size=15,
                        weight=ft.FontWeight.BOLD, color=C["title"], expand=True),
                ft.Text("　".join(meta), size=11.5, color=C["muted"]),
            ], spacing=10),
        ]
        if rec.get("mechanism"):
            children.append(ft.Column([
                ft.Text("故障机理", size=12.5, weight=ft.FontWeight.BOLD, color=C["title"]),
                ft.Text(rec.get("mechanism", ""), size=12.5, color=C["body2"]),
            ], spacing=4))
        if fault_rows:
            def _mcell(text, bold=False, color=C["text"], weight=1):
                t = str(text or "").strip()
                shown = t if t else "—"
                shown_color = color if t else C["muted"]
                return ft.Container(
                    ft.Text(shown, size=12, weight=ft.FontWeight.BOLD if bold else None,
                            color=shown_color),
                    padding=ft.padding.symmetric(horizontal=10, vertical=8),
                    expand=weight,
                    alignment=ft.alignment.top_left,
                )
            def _join(items):
                tail = '、，,；;：: \t'
                seen, out = set(), []
                for it in items:
                    s = str(it or "").strip().rstrip(tail).strip()
                    if s and s not in seen:
                        seen.add(s)
                        out.append(s)
                return "\n".join(out)
            reasons = [r[0] for r in fault_rows if str(r[0] or "").strip()]
            methods = [r[1] for r in fault_rows if str(r[1] or "").strip()]
            measures = [r[2] for r in fault_rows if str(r[2] or "").strip()]
            head_row = ft.Container(
                ft.Row([
                    _mcell("故障原因", bold=True, color=C["title"], weight=3),
                    _mcell("确认方法", bold=True, color=C["title"], weight=2),
                    _mcell("处理措施", bold=True, color=C["title"], weight=2),
                ], spacing=1, vertical_alignment=ft.CrossAxisAlignment.START),
                bgcolor=C["header_bg"],
            )
            body_row = ft.Container(
                ft.Row([
                    _mcell(_join(reasons), weight=3),
                    _mcell(_join(methods), weight=2),
                    _mcell(_join(measures), weight=2),
                ], spacing=1, vertical_alignment=ft.CrossAxisAlignment.START),
                bgcolor=C["card"],
            )
            children.append(ft.Column([
                ft.Text("故障排除", size=12.5, weight=ft.FontWeight.BOLD, color=C["title"]),
                ft.Container(
                    ft.Column([head_row, body_row], spacing=1),
                    border=ft.border.all(1, C["border"]), border_radius=R_PANEL,
                    clip_behavior=ft.ClipBehavior.HARD_EDGE,
                ),
            ], spacing=4))
        children.append(ft.Row([
            ft.Text(cfg.get("watermark", "姓许的苏剑"), size=10.5, color=C["muted"]),
            _eb("✏️ 纠错", height=32, bgcolor=C["primary"], color="white",
                              on_click=lambda e, s=series, c=code, r=rec: _open_fault_fix(e, s, c, r)),
        ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN, wrap=True))
        return ft.Column(children, spacing=14)

    def _code_sort_key(code):
        head, _, sub = code.partition(".")
        nums = [int(x) for x in re.findall(r"\d+", head)]
        return (nums, sub)

    def _iter_faults(series):
        if series == "全部":
            for s, db in FAULT_DB.items():
                for c, rec in db.items():
                    yield s, c, rec
        else:
            for c, rec in FAULT_DB.get(series, {}).items():
                yield series, c, rec

    def _fault_search_text(rec):
        parts = [str(rec.get("title") or ""), str(rec.get("mechanism") or "")]
        for r in rec.get("rows", []):
            parts.extend(str(x or "") for x in r)
        return " ".join(parts).lower()

    def _fault_card_box(s, c, rec):
        return ft.Container(
            _fault_card(s, c, rec),
            bgcolor=C["card_alt"], border_radius=R_CARD, padding=12,
            border=ft.border.all(1, C["border"]),
        )

    def _render_by_code(series, q, raw):
        exact = [(s, c, rec) for s, c, rec in _iter_faults(series) if c == q]
        if exact:
            return ft.Column([
                ft.Text(f"匹配到 {len(exact)} 条：",
                        size=12.5, weight=ft.FontWeight.BOLD, color=C["title"]),
                ft.Column([_fault_card_box(s, c, rec) for s, c, rec in exact], spacing=12),
            ], spacing=14)
        base = q.partition(".")[0]
        prefix = [(s, c, rec) for s, c, rec in _iter_faults(series)
                  if c.partition(".")[0] == base]
        prefix.sort(key=lambda x: (x[0], _code_sort_key(x[1])))
        scope = "全部系列" if series == "全部" else series
        if not prefix:
            return ft.Text(f"未收录 {scope} 的 {raw or '—'} 该码；对应手册说明待补充。",
                           size=12.5, color=C["danger"])
        return ft.Column([
            ft.Text(f"匹配到 {len(prefix)} 条子码（{base}）：",
                    size=12.5, weight=ft.FontWeight.BOLD, color=C["title"]),
            ft.Column([_fault_card_box(s, c, rec) for s, c, rec in prefix], spacing=12),
        ], spacing=14)

    def _render_fault(series, code):
        raw = (code or "").strip()
        q = raw.upper().replace(" ", "").rstrip(".")
        if not q:
            return ft.Text("请输入报警码 / 故障码，或输入现象关键词检索。", size=12.5, color=C["muted"])
        if re.match(r"^E[A-Z]?\d", q):
            return _render_by_code(series, q, raw)
        kw = q.lower()
        hits = [(s, c, rec) for s, c, rec in _iter_faults(series) if kw in _fault_search_text(rec)]
        hits.sort(key=lambda x: (x[0], _code_sort_key(x[1])))
        scope = "所选系列" if series != "全部" else "全部系列"
        if not hits:
            return ft.Text(f"在{scope}未检索到含「{raw}」的故障 / 警告码。",
                           size=12.5, color=C["danger"])
        return ft.Column([
            ft.Text(f"关键词「{raw}」在{scope}匹配到 {len(hits)} 条：",
                    size=12.5, weight=ft.FontWeight.BOLD, color=C["title"]),
            ft.Column([_fault_card_box(s, c, rec) for s, c, rec in hits], spacing=12),
        ], spacing=14)

    # ---------- 故障码纠错 + 版本恢复（与选型纠错机制一致） ----------
    FAULT_FIX_FIELDS = [
        ("故障名称", "title"),
        ("级别", "level"),
        ("能否复位", "reset"),
        ("故障机理", "mechanism"),
    ]

    def _rows_to_text(rows):
        return "\n".join("｜".join((r + ["", "", ""])[:3]) for r in (rows or []))

    def _parse_rows_text(text):
        rows = []
        for line in (text or "").splitlines():
            line = line.strip()
            if not line:
                continue
            parts = [p.strip() for p in re.split(r"[｜|]", line)]
            while len(parts) < 3:
                parts.append("")
            rows.append(parts[:3])
        return rows

    ftgt = {"series": None, "code": None, "rec": None, "note": None}
    ffix_ctrl = {}

    def _open_fault_fix(e, series, code, rec):
        ftgt["series"] = series
        ftgt["code"] = code
        ftgt["rec"] = rec
        ffix_ctrl.clear()
        field_rows = []
        for label, col in FAULT_FIX_FIELDS:
            val = "" if rec.get(col) is None else str(rec.get(col))
            tf = _tf(label=label, value=val, height=38,
                              content_padding=8, width=_w(440),
                              multiline=(col == "mechanism"), min_lines=1, max_lines=4)
            ffix_ctrl[col] = tf
            field_rows.append(tf)
        rows_tf = _tf(label="故障排除表（每行：原因｜确认方法｜处理措施）",
                               value=_rows_to_text(rec.get("rows", [])),
                               height=120, content_padding=8, width=_w(440),
                               multiline=True, min_lines=3, max_lines=8)
        ffix_ctrl["rows"] = rows_tf
        ffix_note = _tf(label="纠错备注（可选）", value="", height=38,
                                 content_padding=8, width=_w(440))
        ftgt["note"] = ffix_note
        ffix_dlg.content = ft.Column(
            [ft.Text(f"{series}　{code}：{rec.get('title', '')}",
                     size=12.5, weight=ft.FontWeight.BOLD, color=C["title"]),
             *field_rows, rows_tf, ffix_note],
            spacing=10, scroll=ft.ScrollMode.AUTO, width=_w(460), height=480)
        page.dialog = ffix_dlg
        ffix_dlg.open = True
        page.update()

    def _close_fault_fix(ev):
        ffix_dlg.open = False
        page.update()

    def _save_fault_fix(ev):
        rec = ftgt["rec"]
        series = ftgt["series"]
        code = ftgt["code"]
        if not rec:
            return
        changes = []
        for label, col in FAULT_FIX_FIELDS:
            old = "" if rec.get(col) is None else str(rec.get(col)).strip()
            new = (ffix_ctrl[col].value or "").strip()
            if new != old:
                changes.append((col, label, old, new))
        old_rows = rec.get("rows", []) or []
        new_rows = _parse_rows_text(ffix_ctrl["rows"].value)
        if _rows_to_text(old_rows) != _rows_to_text(new_rows):
            changes.append(("rows", "故障排除表", _rows_to_text(old_rows), _rows_to_text(new_rows)))
        note = (ftgt["note"].value or "").strip()
        if not changes:
            fault_tip.value = "没有需要保存的改动"
            fault_tip.color = C["warn"]
            page.update()
            return
        try:
            os.makedirs(FAULT_SNAP_DIR, exist_ok=True)
            idx = _next_fault_snapshot_index()
            snap_path = os.path.join(FAULT_SNAP_DIR, f"故障纠错_{idx:03d}.json")
            with open(snap_path, "w", encoding="utf-8") as f:
                json.dump(FAULT_DB, f, ensure_ascii=False, indent=2)

            for col, label, old, new in changes:
                if col == "rows":
                    rec["rows"] = new_rows
                else:
                    rec[col] = new if new else None
            _dump_fault_db()

            entries = _read_fault_fix_log()
            seq = len(entries) + 1
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for col, label, old, new in changes:
                entries.append({
                    "序号": seq, "时间": now, "系列": series, "故障码": code,
                    "字段": label, "改前值": old, "改后值": new,
                    "备注": note, "快照文件": os.path.basename(snap_path),
                })
            _write_fault_fix_log(entries)

            ffix_dlg.open = False
            fault_tip.value = f"已保存纠错：{code}（{len(changes)} 项）"
            fault_tip.color = C["success"]
            _log(f"故障码纠错保存：{series} {code}，{len(changes)} 项改动")
            _query_fault(None)
        except Exception as ex:
            _log(f"故障码纠错保存失败：{ex}", "ERROR")
            fault_tip.value = f"纠错保存失败：{ex}"
            fault_tip.color = C["danger"]
            page.update()

    ffix_dlg = ft.AlertDialog(
        open=False,
        title=ft.Text("✏️ 纠错故障码", size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
        content=ft.Column([ft.Text("")], width=_w(460), height=480),
        actions=[
            _tb("取消", on_click=_close_fault_fix),
            _tb("保存纠错", on_click=_save_fault_fix),
        ],
        actions_alignment=ft.MainAxisAlignment.END,
    )

    def _reload_fault_db():
        FAULT_DB.clear()
        FAULT_DB.update(_load_fault_db())

    def _restore_fault_snapshot(snap_file):
        if not snap_file:
            return
        snap_path = os.path.join(FAULT_SNAP_DIR, snap_file)
        if not os.path.exists(snap_path):
            fault_tip.value = "快照文件不存在，无法恢复"
            fault_tip.color = C["danger"]
            page.update()
            return
        _copy_file(snap_path, FAULT_DB_WRITABLE)
        _reload_fault_db()
        fhist_dlg.open = False
        fault_tip.value = f"已恢复到快照：{snap_file}"
        fault_tip.color = C["success"]
        _log(f"恢复故障码快照：{snap_file}")
        _query_fault(None)

    forig_confirm = {"armed": False}
    forig_btn = _eb("♻ 恢复出厂原始", height=34, bgcolor=C["danger"], color="white")

    def _restore_fault_original(ev):
        if not os.path.exists(FAULT_DB_ORIGINAL):
            fault_tip.value = "未找到出厂原始基线"
            fault_tip.color = C["danger"]
            page.update()
            return
        if not forig_confirm["armed"]:
            forig_confirm["armed"] = True
            forig_btn.text = "再次点击确认（撤销全部纠错）"
            forig_btn.color = C["danger"]
            page.update()
            return
        _copy_file(FAULT_DB_ORIGINAL, FAULT_DB_WRITABLE)
        _reload_fault_db()
        forig_confirm["armed"] = False
        forig_btn.text = "♻ 恢复出厂原始"
        forig_btn.color = "white"
        fhist_dlg.open = False
        fault_tip.value = "已恢复为出厂原始故障库（全部纠错已撤销）"
        fault_tip.color = C["success"]
        _log("恢复出厂原始故障库（撤销全部纠错）")
        _query_fault(None)

    forig_btn.on_click = _restore_fault_original

    def _delete_fault_entry(seq, snap_file):
        entries = _read_fault_fix_log()
        kept = [e for e in entries if str(e.get("序号")) != str(seq)]
        if snap_file and not any(str(e.get("快照文件") or "") == snap_file for e in kept):
            sp = os.path.join(FAULT_SNAP_DIR, snap_file)
            if os.path.exists(sp):
                os.remove(sp)
        _write_fault_fix_log(kept)
        fault_tip.value = f"已删除第 {seq} 条纠错记录"
        fault_tip.color = C["success"]
        _log(f"删除故障码纠错记录 #{seq}")
        open_fault_history(None)

    fclear_confirm = {"armed": False}
    fclear_btn = _eb("🗑 清空全部记录", height=34, bgcolor=C["danger"], color="white")

    def _clear_fault_entries(ev):
        if not fclear_confirm["armed"]:
            fclear_confirm["armed"] = True
            fclear_btn.text = "再次点击确认清空（不回滚当前故障库）"
            fclear_btn.color = C["danger"]
            page.update()
            return
        _write_fault_fix_log([])
        if os.path.isdir(FAULT_SNAP_DIR):
            for f in os.listdir(FAULT_SNAP_DIR):
                if f.endswith(".json"):
                    os.remove(os.path.join(FAULT_SNAP_DIR, f))
        fclear_confirm["armed"] = False
        fclear_btn.text = "🗑 清空全部记录"
        fclear_btn.color = "white"
        fault_tip.value = "已清空全部纠错记录（当前故障库与出厂基线不受影响）"
        _log("清空全部故障码纠错记录")
        fault_tip.color = C["success"]
        open_fault_history(None)

    fclear_btn.on_click = _clear_fault_entries

    # ---------- 故障码库 文件导出 / 导入（整库覆盖，导入前自动存快照可回滚） ----------
    def _export_fault_db(ev):
        if not _is_desktop(page):
            fault_tip.value = "当前环境暂不支持导出文件，请使用桌面版导出"
            fault_tip.color = C["warn"]
            page.update()
            return
        try:
            os.makedirs(DATA_EXPORT_DIR, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out = os.path.join(DATA_EXPORT_DIR, f"故障码库_{stamp}.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "fault"
            ws.append(FAULT_XLSX_HEADER)
            for series, codes in FAULT_DB.items():
                for code, rec in codes.items():
                    rows = rec.get("rows", []) or []
                    if not rows:
                        rows = [["", "", ""]]
                    for r in rows:
                        r3 = (list(r) + ["", "", ""])[:3]
                        ws.append([
                            series, code,
                            rec.get("title", ""), rec.get("level", ""),
                            rec.get("reset", ""), rec.get("mechanism", ""),
                            r3[0], r3[1], r3[2],
                        ])
            wb.save(out)
            fault_tip.value = f"已导出故障码库：{os.path.basename(out)}"
            fault_tip.color = C["success"]
            _log(f"导出故障码库：{out}")
        except Exception as ex:
            _log(f"导出故障码库失败：{ex}", "ERROR")
            fault_tip.value = f"导出失败：{ex}"
            fault_tip.color = C["danger"]
        page.update()

    def _import_fault_db(path):
        if not path or not os.path.exists(path):
            return
        try:
            wb = openpyxl.load_workbook(path)
            ws = wb["fault"] if "fault" in wb.sheetnames else wb.active
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                fault_tip.value = "导入文件为空，未导入"
                fault_tip.color = C["danger"]
                page.update()
                return
            header = [str(h) if h is not None else "" for h in rows[0]]

            def _col(name):
                return header.index(name) if name in header else -1

            ci_s, ci_c = _col("系列"), _col("故障码")
            if ci_s < 0 or ci_c < 0:
                fault_tip.value = "导入文件表头缺少「系列」或「故障码」列，未导入"
                fault_tip.color = C["danger"]
                page.update()
                return
            ci_t, ci_l, ci_r, ci_m = _col("故障名称"), _col("级别"), _col("能否复位"), _col("故障机理")
            ci_a, ci_b, ci_d = _col("故障原因"), _col("确认方法"), _col("处理措施")

            merged = {}
            for r in rows[1:]:
                if not r or not any(v is not None for v in r):
                    continue

                def _g(i):
                    return "" if i < 0 or i >= len(r) or r[i] is None else str(r[i]).strip()

                series = _g(ci_s)
                code = _g(ci_c)
                if not series or not code:
                    continue
                rec = merged.setdefault((series, code), {
                    "title": "", "level": "", "reset": "", "mechanism": "", "rows": []})
                if not rec["title"]:
                    rec["title"] = _g(ci_t)
                if not rec["level"]:
                    rec["level"] = _g(ci_l)
                if not rec["reset"]:
                    rec["reset"] = _g(ci_r)
                if not rec["mechanism"]:
                    rec["mechanism"] = _g(ci_m)
                a, b, d = _g(ci_a), _g(ci_b), _g(ci_d)
                if a or b or d:
                    rec["rows"].append([a, b, d])
            if not merged:
                fault_tip.value = "导入文件中没有有效故障码记录，未导入"
                fault_tip.color = C["danger"]
                page.update()
                return

            new_db = {}
            for (series, code), rec in merged.items():
                new_db.setdefault(series, {})[code] = rec

            os.makedirs(FAULT_SNAP_DIR, exist_ok=True)
            idx = _next_fault_snapshot_index()
            snap_path = os.path.join(FAULT_SNAP_DIR, f"故障纠错_{idx:03d}.json")
            with open(snap_path, "w", encoding="utf-8") as f:
                json.dump(FAULT_DB, f, ensure_ascii=False, indent=2)

            FAULT_DB.clear()
            FAULT_DB.update(new_db)
            _dump_fault_db()

            entries = _read_fault_fix_log()
            seq = len(entries) + 1
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            entries.append({
                "序号": seq, "时间": now, "系列": "（整库导入）", "故障码": "",
                "字段": "整库导入", "改前值": "导入前整库", "改后值": os.path.basename(path),
                "备注": "", "快照文件": os.path.basename(snap_path),
            })
            _write_fault_fix_log(entries)

            fault_tip.value = f"已导入故障码库：{len(merged)} 条（导入前已存快照）"
            fault_tip.color = C["success"]
            _log(f"整库导入故障码库：{os.path.basename(path)}，{len(merged)} 条")
            _query_fault(None)
            open_fault_history(None)
        except Exception as ex:
            _log(f"导入故障码库失败：{ex}", "ERROR")
            fault_tip.value = f"导入失败：{ex}"
            fault_tip.color = C["danger"]
            page.update()

    def open_fault_history(e):
        entries = _read_fault_fix_log()
        forig_confirm["armed"] = False
        forig_btn.text = "♻ 恢复出厂原始"
        forig_btn.color = "white"
        fclear_confirm["armed"] = False
        fclear_btn.text = "🗑 清空全部记录"
        fclear_btn.color = "white"
        hist_rows = []
        if not entries:
            hist_rows.append(ft.Text("暂无纠错记录", size=12.5, color=C["muted"]))
        for en in entries:
            snap = en.get("快照文件") or ""
            seq = en.get("序号")
            hist_rows.append(ft.Row([
                ft.Column([
                    ft.Text(f"#{seq}　{en.get('时间')}", size=11, color=C["muted"]),
                    ft.Text(f"{en.get('系列')}　{en.get('故障码')}", size=12, color=C["text"]),
                    ft.Text(f"{en.get('字段')}：{en.get('改前值') or '空'} → {en.get('改后值') or '空'}",
                            size=11.5, color=C["link"]),
                ], spacing=2, expand=True),
                _eb("恢复此版本", height=30, bgcolor=C["gray"], color="white",
                                  on_click=lambda ev, s=snap: _restore_fault_snapshot(s)),
                _eb("🗑", height=30, bgcolor=C["danger"], color="white",
                                  on_click=lambda ev, q=seq, fn=snap: _delete_fault_entry(q, fn)),
            ], spacing=10))
        last = entries[-1] if entries else None
        fundo_btn = _eb(
            "↩ 撤销最近一次", height=34, bgcolor=C["primary"], color="white",
            disabled=last is None,
            on_click=lambda ev: _restore_fault_snapshot(last.get("快照文件") if last else ""))
        exp_fault_btn = _eb("📤 导出文件", height=34, bgcolor=C["gray"], color="white",
                                          on_click=_export_fault_db)
        imp_fault_btn = _eb(
            "📥 导入文件", height=34, bgcolor=C["gray"], color="white",
            on_click=_import_fault_click)
        fhist_dlg.content = ft.Column(
            [ft.Row([fundo_btn, forig_btn, fclear_btn], spacing=12, wrap=True),
             ft.Row([exp_fault_btn, imp_fault_btn], spacing=12, wrap=True),
             ft.Divider(height=8, color=C["divider"]),
             *hist_rows],
            spacing=10, scroll=ft.ScrollMode.AUTO, width=_w(480), height=440)
        page.dialog = fhist_dlg
        fhist_dlg.open = True
        page.update()

    def _close_fault_hist(ev):
        fhist_dlg.open = False
        page.update()

    fhist_dlg = ft.AlertDialog(
        open=False,
        title=ft.Text("🕘 故障码纠错记录与版本恢复", size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
        content=ft.Column([ft.Text("")], width=_w(480), height=440),
        actions=[_tb("关闭", on_click=_close_fault_hist)],
        actions_alignment=ft.MainAxisAlignment.END,
    )

    btn_fault_hist = _eb(
        "🕘 纠错记录", height=34, bgcolor=C["gray"], color="white",
        on_click=open_fault_history,
    )

    fault_series = _dd(
        label="驱动器系列",
        options=[ft.dropdown.Option("全部")] + [ft.dropdown.Option(v) for v in FAULT_SERIES],
        value=FAULT_SERIES[0],
        width=_w(180),
    )
    fault_code = _tf(
        label="报警码 / 故障码",
        hint_text="输故障码（如 E108.0 / E108），或输现象关键词（如 过载、编码器断线）",
        on_submit=lambda e: _query_fault(None),
    )
    fault_result = ft.Container(
        ft.Text("选择驱动器系列，输入故障码或现象关键词，点击【查询】查看机理与排除方法。",
                size=12, color=C["muted"]),
        bgcolor=C["card"], border_radius=R_CARD, padding=16,
        border=ft.border.all(1, C["border"]),
    )
    fault_tip = ft.Text("", size=12, color=C["muted"])

    _last_series = cfg.get("last_fault_series", "")
    _last_code = (cfg.get("last_fault_code", "") or "").strip().upper()
    if _last_series in FAULT_SERIES or _last_series == "全部":
        fault_series.value = _last_series
    if _last_code:
        fault_code.value = _last_code
    last_fault = (fault_series.value, _last_code) if _last_code else None

    def _query_fault(e):
        nonlocal last_fault
        series = fault_series.value
        code = (fault_code.value or "").strip().upper()
        if code:
            last_fault = (series, code)
            cfg["last_fault_series"] = series
            cfg["last_fault_code"] = code
            _save_settings(cfg)
            btn_last_fault.disabled = False
            btn_last_fault.text = f"🕘 上次：{code}"
        fault_result.content = _render_fault(series, code)
        fault_result.update()
        btn_last_fault.update()

    def _repeat_last_fault(e):
        if not last_fault:
            fault_tip.value = "还没有查询记录"
            fault_tip.color = C["muted"]
            fault_tip.update()
            return
        series, code = last_fault
        fault_series.value = series
        fault_code.value = code
        _query_fault(None)

    fault_query_btn = _eb("🔍 查询", height=34, bgcolor=C["primary"], color="white",
                                        on_click=_query_fault)
    btn_last_fault = _eb(
        f"🕘 上次：{last_fault[1]}" if last_fault else "🕘 上次查询",
        height=34, bgcolor=C["gray"], color="white",
        disabled=last_fault is None, on_click=_repeat_last_fault)

    btn_hint_fault = _eb(
        "❓ 使用说明",
        height=34, bgcolor=C["gray"], color="white",
        on_click=lambda e: _open_hint_dialog("fault"))

    # ---------- AI 故障诊断面板 ----------
    ai_input = _tf(
        label="描述现象或输入故障码",
        hint_text="如：电机过载报警、编码器断线、E108.0",
        multiline=True, min_lines=2, max_lines=3, expand=True,
    )
    ai_answer = ft.Text("填写现象或故障码后点「提问」，AI 依据本地故障库给出原因判断与排查步骤。",
                        size=12.5, color=C["muted"])
    ai_base = _tf(label="接口地址（Base URL）", width=_w(300), text_size=12.5)
    ai_model = _tf(label="模型名", width=_w(300), text_size=12.5)
    ai_key = _tf(label="API Key", password=True, width=_w(300), text_size=12.5)
    ai_provider_drop = _dd(
        label="渠道", width=_w(300),
        options=[ft.dropdown.Option(k, v["label"]) for k, v in AI_PROVIDERS.items()],
        value=cfg.get("ai_provider", "zhipu"),
        on_change=lambda e: _ai_provider_changed(),
    )

    def _ai_provider_changed():
        prov = AI_PROVIDERS.get(ai_provider_drop.value, {})
        ai_base.value = prov.get("base_url", "")
        ai_model.value = prov.get("model", "")
        page.update()

    def _ai_save_config(e):
        cfg["ai_provider"] = ai_provider_drop.value
        cfg["ai_base_url"] = ai_base.value.strip()
        cfg["ai_model"] = ai_model.value.strip()
        cfg["ai_api_key"] = ai_key.value.strip()
        try:
            _save_settings(cfg)
            _log("AI 接口配置已保存")
        except Exception as ex:
            _log(f"AI 配置保存失败：{ex}", "ERROR")
        ai_cfg_dlg.open = False
        page.update()

    def _ai_cfg_cancel(e):
        ai_cfg_dlg.open = False
        page.update()

    def _ai_worker(q):
        try:
            if not (cfg.get("ai_base_url") and cfg.get("ai_model") and cfg.get("ai_api_key")):
                ai_answer.value = "尚未配置 AI 接口：点「⚙ 接口配置」选择渠道并填入 API Key。"
                return
            hits = _ai_recall(q)
            if not hits:
                ai_answer.value = "本地故障库未检索到相关条目，AI 无法据此作答。"
                return
            entries = _ai_format_entries(hits)
            messages = [
                {"role": "system", "content": AI_SYSTEM_PROMPT},
                {"role": "user", "content": f"【本地故障库条目】\n{entries}\n\n【用户问题】{q}"},
            ]
            ans = _ai_chat(cfg, messages, _ssl_ctx)
            ai_answer.value = ans or "AI 未返回有效内容。"
        except Exception as ex:
            msg = str(ex)
            if "429" in msg:
                ai_answer.value = "触发平台限流或免费额度已用完（429）：请稍等一分钟再试；若仍报错，到模型平台后台查看额度 / 余额。"
            else:
                ai_answer.value = f"请求失败：{msg}（未配置接口、网络不通或模型返回异常）"
        finally:
            try:
                ai_answer.update()
            except Exception:
                pass

    def _ai_ask(e):
        q = (ai_input.value or "").strip()
        if not q:
            ai_answer.value = "请先描述现象或输入故障码。"
            ai_answer.update()
            return
        ai_answer.value = "正在检索本地故障库并请求 AI 分析…"
        ai_answer.update()
        threading.Thread(target=_ai_worker, args=(q,), daemon=True).start()

    def _ai_open_config(e):
        ai_provider_drop.value = cfg.get("ai_provider", "zhipu")
        ai_base.value = cfg.get("ai_base_url", "")
        ai_model.value = cfg.get("ai_model", "")
        ai_key.value = cfg.get("ai_api_key", "")
        page.dialog = ai_cfg_dlg
        ai_cfg_dlg.open = True
        page.update()

    def _ai_close(e):
        ai_dlg.open = False
        page.update()

    ai_cfg_dlg = ft.AlertDialog(
        open=False,
        title=ft.Text("⚙ AI 接口配置", size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
        content=ft.Container(
            ft.Column([
                ai_provider_drop,
                ai_base,
                ai_model,
                ai_key,
                ft.Text("Key 仅保存在本地配置文件；提问内容与召回条目会发送至所选模型服务商。",
                        size=10.5, color=C["muted"]),
            ], spacing=12, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            **_dlg_bg,
        ),
        actions=[
            _tb("取消", on_click=_ai_cfg_cancel),
            _eb("保存", bgcolor=C["primary"], color="white", on_click=_ai_save_config),
        ],
        actions_alignment=ft.MainAxisAlignment.END,
    )

    ai_dlg = ft.AlertDialog(
        open=False,
        title=ft.Text("🤖 AI 故障诊断", size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
        content=ft.Container(
            ft.Column([
                ai_input,
                ft.Row([
                    _eb("提问", bgcolor=C["primary"], color="white", on_click=_ai_ask),
                    _eb("⚙ 接口配置", bgcolor=C["gray"], color="white", on_click=_ai_open_config),
                ], spacing=12, alignment=ft.MainAxisAlignment.END, wrap=True),
                ft.Container(
                    ft.Column([ai_answer], scroll=ft.ScrollMode.AUTO),
                    bgcolor=C["card_alt"], border_radius=R_PANEL, padding=12, height=260,
                ),
                ft.Text("AI 结论仅供参考，具体以官方手册与现场排查为准。", size=10.5, color=C["muted"]),
            ], spacing=12, horizontal_alignment=ft.CrossAxisAlignment.CENTER, width=_w(480)),
            **_dlg_bg,
        ),
        actions=[_tb("关闭", on_click=_ai_close)],
        actions_alignment=ft.MainAxisAlignment.END,
    )

    DEV_PASS = "000000"

    def _dev_close(dlg):
        dlg.open = False
        page.update()

    def _ai_open_panel(e):
        dev_pass = _tf(label="管理者验证码", password=True, width=_w(220), text_size=13)
        dev_tip = ft.Text("", size=11, color=C["danger"])

        def _dev_try(ev):
            if (dev_pass.value or "").strip() == DEV_PASS:
                dlg.open = False
                page.dialog = ai_dlg
                ai_dlg.open = True
                page.update()
            else:
                dev_tip.value = "验证码不正确"
                page.update()

        dlg = ft.AlertDialog(
            open=True,
            title=ft.Text("🤖 AI 诊断", size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
            content=ft.Container(
                ft.Column([
                    ft.Text("该功能正在开发中，暂未开放使用。", size=13, color=C["text"]),
                    ft.Text("持有管理者验证码可提前试用。", size=12, color=C["muted"]),
                    dev_pass,
                    dev_tip,
                    ft.Row([
                        _tb("关闭", on_click=lambda ev: _dev_close(dlg)),
                        _eb("试用", bgcolor=C["primary"], color="white", on_click=_dev_try),
                    ], alignment=ft.MainAxisAlignment.END, spacing=12, wrap=True),
                ], spacing=12, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                **_dlg_bg,
            ),
        )
        page.dialog = dlg
        page.update()

    btn_ai = _eb(
        "🤖 AI 诊断", height=34, bgcolor=C["primary"], color="white", on_click=_ai_open_panel)

    fault_header = ft.Container(
        ft.Row([
            ft.Container(logo_widget, bgcolor=C["card"], border_radius=R_BADGE, padding=3),
            ft.Column([
                ft.Text("故障码 / 警告码查询", size=21, weight=ft.FontWeight.BOLD, color="white"),
                ft.Text("按驱动器系列查询报警码、故障码的故障机理与排除方法",
                        size=11.5, color=C["banner_sub"]),
            ], spacing=2, expand=True),
            _eb(content=ft.Text("主页", weight=ft.FontWeight.BOLD, color=C["title"]),
                               height=34, bgcolor=C["card"], on_click=_go_home),
        ], spacing=14, alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
        bgcolor=C["banner"],
        padding=ft.padding.symmetric(horizontal=20, vertical=15),
    )

    fault_view = ft.Column([
        fault_header,
        ft.Container(
            ft.Column([
                (ft.Row([fault_series, fault_code], spacing=12,
                        vertical_alignment=ft.CrossAxisAlignment.END)
                 if not is_mobile else
                 ft.Column([fault_series, fault_code], spacing=12)),
                ft.Row([fault_query_btn, btn_last_fault, btn_fault_hist, btn_hint_fault, btn_ai],
                       spacing=12, wrap=True),
                fault_result,
                fault_tip,
            ], spacing=12),
            padding=ft.padding.all(18),
        ),
    ], spacing=0, tight=True)

    # ---------- 产品替换对照表（一期：手工对照，可从已选搭配带入汇川型号） ----------
    # ---------- 产品替换对照表（按方案分组；接口/电池相同合并、不同归入各自套） ----------
    def _repl_row_widget(_type, inov="", note="", qty="1", orig=""):
        return {
            "type": _type,
            "orig": _tf(value=orig, hint_text="客户原型号（照抄铭牌）", expand=True),
            "inov": _tf(value=inov, hint_text="汇川对应型号", expand=True),
            "qty": _tf(value=qty, width=_w(70), text_size=12.5),
            "note": _tf(value=note, width=_w(190), text_size=12.5),
        }

    def _repl_group_title(it):
        try:
            kw = float("".join(c for c in str(it.get("power_text") or "") if c.isdigit() or c == "."))
            watt = f"{kw * 1000:g}W"
        except Exception:
            watt = str(it.get("power_text") or "").strip()
        comm = "伺服总线" if str(it.get("comm") or "") == "EtherCAT" else "脉冲"
        return f"{watt}{comm}" if watt else comm

    def _repl_empty_group():
        return {"title": "方案一",
                "rows": [_repl_row_widget(t) for t in REPL_BODY_TYPES + REPL_ACC_TYPES]}

    repl_groups = [_repl_empty_group()]
    repl_list_col = ft.Column(spacing=10, tight=True)
    repl_tip = ft.Text("", size=12, color=C["muted"])

    def _repl_show_title(g):
        # 多套分组时显示各组标题；单套平铺不显示方案标题；共用「接口和电池」组标题始终保留
        return len(repl_groups) > 1 or g["title"] == "接口和电池"

    def _rebuild_repl_list():
        repl_list_col.controls = [_repl_group_card(g, _repl_show_title(g)) for g in repl_groups]

    def _repl_build_groups(items):
        groups = []
        acc_pool = {}
        for it in items:
            bom = _item_bom_rows(it, depth_drop.value, flex_check.value)
            by_type = {r[0]: r for r in bom}
            g = {"title": _repl_group_title(it), "rows": []}
            for t in REPL_BODY_TYPES:
                r = by_type.get(t)
                if r:
                    g["rows"].append(_repl_row_widget(t, inov=r[1], note=r[3]))
            groups.append(g)
            gi = len(groups) - 1
            for t in REPL_ACC_TYPES:
                r = by_type.get(t)
                if not r:
                    continue
                model = r[1]
                if model not in acc_pool:
                    acc_pool[model] = {"type": t, "note": r[3], "count": 0, "groups": []}
                acc_pool[model]["count"] += 1
                acc_pool[model]["groups"].append(gi)
        shared = []
        for model, info in acc_pool.items():
            if len(info["groups"]) > 1:
                shared.append(_repl_row_widget(info["type"], inov=model,
                                                note=info["note"], qty=str(info["count"])))
            else:
                groups[info["groups"][0]]["rows"].append(
                    _repl_row_widget(info["type"], inov=model, note=info["note"]))
        if shared:
            groups.append({"title": "接口和电池", "rows": shared})
        return groups

    repl_tip = ft.Text("", size=12, color=C["muted"])

    def on_repl_fill(e):
        items = list(selected.values()) if selected else (state.get("res") or [])
        if not items:
            repl_tip.value = "暂无已选搭配：请先到「伺服选型工具」选好搭配并勾选，或在下方表内手填汇川型号。"
            repl_tip.color = C["warn"]
            page.update()
            return
        repl_groups[:] = _repl_build_groups(items)
        _rebuild_repl_list()
        repl_tip.value = f"已带入 {len(items)} 套选型搭配；初始型号列请照抄客户铭牌。"
        repl_tip.color = C["success"]
        page.update()

    def on_repl_clear_inov(e):
        for g in repl_groups:
            for w in g["rows"]:
                w["inov"].value = ""
                w["note"].value = ""
        repl_tip.value = "已清空汇川对应型号与备注，初始型号与数量保留。"
        repl_tip.color = C["muted"]
        page.update()

    def on_repl_reset(e):
        repl_groups[:] = [_repl_empty_group()]
        _rebuild_repl_list()
        repl_tip.value = "对照表已重置。"
        repl_tip.color = C["muted"]
        page.update()

    def _repl_collect_groups():
        out = []
        for g in repl_groups:
            rows = []
            for w in g["rows"]:
                orig = (w["orig"].value or "").strip()
                inov = (w["inov"].value or "").strip()
                if not orig and not inov:
                    continue
                qty = (w["qty"].value or "").strip() or "1"
                note = (w["note"].value or "").strip()
                rows.append([w["type"], orig, inov, qty, note])
            if rows:
                show = len(repl_groups) > 1 or g["title"] == "接口和电池"
                out.append({"title": g["title"] if show else "", "rows": rows})
        return out

    def on_export_repl(e):
        if not _is_desktop(page):
            repl_tip.value = "当前环境暂不支持导出文件，请使用桌面版导出"
            repl_tip.color = C["warn"]
            page.update()
            return
        groups = _repl_collect_groups()
        if not groups:
            repl_tip.value = "请先填写对照内容（初始型号或汇川型号至少填一项）"
            repl_tip.color = C["warn"]
            page.update()
            return
        try:
            path, n = export_repl(groups, cfg, conditions_txt="产品替换对照表")
            repl_tip.value = f"已导出 {n} 行对照清单：{path}"
            repl_tip.color = C["success"]
        except Exception as ex:
            _log(f"导出替换清单失败：{ex}", "ERROR")
            repl_tip.value = f"导出失败：{ex}"
            repl_tip.color = C["danger"]
        page.update()

    def _repl_row_card(w):
        type_label = ft.Container(
            ft.Text(w["type"], size=12.5, weight=ft.FontWeight.BOLD, color=C["title"]),
            width=_w(96),
        )
        if is_mobile:
            return ft.Container(
                ft.Column([type_label, ft.Row([w["orig"], w["qty"]], spacing=10),
                           w["inov"], w["note"]], spacing=8),
                bgcolor=C["card"], border_radius=R_CARD, padding=12,
                border=ft.border.all(1, C["border"]),
            )
        return ft.Container(
            ft.Row([type_label, w["orig"], w["inov"], w["qty"], w["note"]],
                   spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            bgcolor=C["card"], border_radius=R_CARD, padding=12,
            border=ft.border.all(1, C["border"]),
        )

    def _repl_group_card(g, show_title=True):
        head = []
        if show_title:
            head.append(ft.Container(
                ft.Text(g["title"], size=13, weight=ft.FontWeight.BOLD, color=C["title"]),
                bgcolor=C["card_alt"], border_radius=R_BADGE,
                padding=ft.padding.symmetric(horizontal=10, vertical=6),
            ))
        return ft.Container(
            ft.Column(head + [ft.Column([_repl_row_card(w) for w in g["rows"]], spacing=10)], spacing=10),
            border=ft.border.all(1, C["border"]), border_radius=R_CARD, padding=12,
        )

    repl_header = ft.Container(
        ft.Row([
            ft.Container(logo_widget, bgcolor=C["card"], border_radius=R_BADGE, padding=3),
            ft.Column([
                ft.Text("产品替换选型", size=21, weight=ft.FontWeight.BOLD, color="white"),
                ft.Text("客户原品牌型号对照汇川选型，逐项匹配电机 / 驱动器 / 线缆 / 接头 / 电池并导出清单",
                        size=11.5, color=C["banner_sub"]),
            ], spacing=2, expand=True),
            _eb(content=ft.Text("主页", weight=ft.FontWeight.BOLD, color=C["title"]),
                height=34, bgcolor=C["card"], on_click=_go_home),
        ], spacing=14, alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
        bgcolor=C["banner"],
        padding=ft.padding.symmetric(horizontal=20, vertical=15),
    )

    repl_head_row = ft.Row([
        ft.Container(ft.Text("类型", size=11.5, weight=ft.FontWeight.BOLD, color=C["sub"]), width=_w(96)),
        ft.Container(ft.Text("初始型号（客户原品牌）", size=11.5, weight=ft.FontWeight.BOLD, color=C["sub"]), expand=True),
        ft.Container(ft.Text("汇川对应型号", size=11.5, weight=ft.FontWeight.BOLD, color=C["sub"]), expand=True),
        ft.Container(ft.Text("数量", size=11.5, weight=ft.FontWeight.BOLD, color=C["sub"]), width=_w(70)),
        ft.Container(ft.Text("备注", size=11.5, weight=ft.FontWeight.BOLD, color=C["sub"]), width=_w(190)),
    ], spacing=10)

    _rebuild_repl_list()

    repl_view = ft.Column([
        repl_header,
        ft.Container(
            ft.Column([
                ft.Text("填写说明：在「伺服选型工具」选好搭配并勾选后点「带入选型搭配」按方案分组带入；初始型号列照抄客户铭牌，也可直接手填任一列。",
                        size=12, color=C["muted"]),
                repl_head_row,
                repl_list_col,
                ft.Row([
                    _eb("🔄 带入选型搭配", height=34, bgcolor=C["primary"], color="white", on_click=on_repl_fill),
                    _eb("⬇ 导出替换清单", height=34, bgcolor=C["success"], color="white", on_click=on_export_repl),
                    _eb("📂 打开替换清单", height=34, bgcolor=C["primary"], color="white", on_click=open_repl_folder),
                    _eb("📋 导出记录", height=34, bgcolor=C["gray"], color="white",
                        on_click=lambda e: open_export_log(e, "repl")),
                    _eb("🧹 清空汇川型号", height=34, bgcolor=C["gray"], color="white", on_click=on_repl_clear_inov),
                    _eb("❓ 使用说明", height=34, bgcolor=C["gray"], color="white",
                        on_click=lambda e: _open_hint_dialog("repl")),
                    _eb("↺ 重置全部", height=34, bgcolor=C["gray"], color="white", on_click=on_repl_reset),
                ], spacing=12, wrap=True),
                repl_tip,
            ], spacing=12),
            padding=ft.padding.all(18),
        ),
    ], spacing=0, tight=True)

    def _open_repl(e):
        page.floating_action_button = back_top_btn
        _show(repl_view)
        _show_hint("repl")

    # ---------- 隐藏主题彩蛋：版本号连点 7 次，暗号 JWY 本次套用品牌蓝 ----------
    _egg_clicks = {"n": 0}
    egg_input = _tf(label="输入暗号", password=True, width=_w(220), text_size=13)
    egg_tip = ft.Text("", size=11, color=C["danger"])

    def _egg_open_dlg():
        egg_input.value = ""
        egg_tip.value = ""
        page.dialog = egg_dlg
        egg_dlg.open = True
        page.update()

    def _egg_submit(e):
        if (egg_input.value or "").strip().upper() == "JWY":
            egg_dlg.open = False
            page.update()
            _log("彩蛋解锁：品牌蓝主题（本次）")
            if rebuild:
                rebuild("brand")
        else:
            egg_tip.value = "暗号不正确，再试一次"
            page.update()

    def _egg_cancel(e):
        egg_dlg.open = False
        page.update()

    egg_dlg = ft.AlertDialog(
        open=False,
        title=ft.Text("🔒 隐藏主题", size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
        content=ft.Container(
            ft.Column([
                ft.Text("检测到隐藏入口，请输入暗号解锁。", size=12.5, color=C["text"]),
                egg_input,
                egg_tip,
                ft.Row([
                    _tb("取消", on_click=_egg_cancel),
                    _eb("解锁", bgcolor=C["primary"], color="white", on_click=_egg_submit),
                ], alignment=ft.MainAxisAlignment.END, spacing=12),
            ], spacing=12, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            **_dlg_bg,
        ),
    )

    def _egg_click(e):
        _egg_clicks["n"] += 1
        if _egg_clicks["n"] >= 7:
            _egg_clicks["n"] = 0
            _egg_open_dlg()

    # ---------- 应用主页（启动器）：以后新增功能在此追加卡片 ----------
    def _open_selector(e):
        page.floating_action_button = back_top_btn
        _show(selector_view)
        _show_hint("selector")

    cover_container = None
    if mode == "brand" and not _is_mobile(page) and (web or os.path.exists(BRAND_LOGO)):
        cover_container = ft.Container(
            ft.Column([
                ft.Row([
                    ft.Container(
                        ft.Image(src=_asset("brand_logo.png"), height=50, fit=ft.ImageFit.CONTAIN),
                        on_click=lambda e: page.launch_url(JWY_WEBSITE),
                        tooltip="访问金伍源官网",
                        ink=True,
                    ),
                    ft.Column([
                        ft.Text("汇川伺服工具集", size=18, weight=ft.FontWeight.BOLD, color=C["title"]),
                        ft.Text("选型匹配 · 产品替换 · 故障码查询 · BOM 清单导出", size=12, color=C["sub"]),
                    ], spacing=5, expand=True),
                ], spacing=18, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                ft.Container(width=56, height=4, border_radius=2, bgcolor=C["warn"]),
            ], spacing=12),
            bgcolor=C["card"],
            border_radius=R_CARD,
            padding=20,
            border=ft.border.all(1, C["border"]),
        )
    else:
        _cover_custom = _custom_img_path(cfg, "cover") if mode == "custom" else None
        _cover_src = _cover_custom if _cover_custom else _asset("cover.jpg")
        if not _is_mobile(page) and (_cover_custom or web or os.path.exists(_cover_src)):
            cover_container = ft.Container(
                ft.Image(src=_cover_src, height=190, fit=ft.ImageFit.COVER, expand=True),
                border_radius=R_CARD,
                clip_behavior=ft.ClipBehavior.HARD_EDGE,
            )

    _card_img = _custom_img_path(cfg, "card") if mode == "custom" else None
    _card_bg = {}
    if _card_img:
        _card_bg = {"image_src": _card_img, "image_fit": ft.ImageFit.COVER, "image_opacity": 0.16}

    select_icon_widget = (
        ft.Image(src=_asset("select_icon.png"), width=40, height=40, fit=ft.ImageFit.CONTAIN)
        if web or os.path.exists(SELECT_ICON)
        else logo_widget
    )
    sel_card = ft.Container(
        ft.Row([
            ft.Container(select_icon_widget, bgcolor=C["card"], border_radius=R_BADGE, padding=3),
            ft.Column([
                ft.Text("伺服选型工具", size=18, weight=ft.FontWeight.BOLD, color=C["title"]),
                ft.Text("电机 / 驱动器 / 动力线 / 编码线联合匹配，支持 BOM 导出、手册下载、纠错与版本恢复",
                        size=12.5, color=C["sub"]),
            ], spacing=6, expand=True),
            _tb("进入 →", on_click=_open_selector),
        ], spacing=16),
        bgcolor=C["card"],
        border_radius=R_CARD,
        padding=12 if is_mobile else 18,
        border=ft.border.all(1, C["border"]),
        on_click=_open_selector,
        ink=True,
        **_card_bg,
    )

    fault_icon_widget = (
        ft.Image(src=_asset("fault_icon.png"), width=40, height=40, fit=ft.ImageFit.CONTAIN)
        if web or os.path.exists(FAULT_ICON)
        else ft.Text("⚠️", size=22, text_align=ft.TextAlign.CENTER)
    )
    fault_card = ft.Container(
        ft.Row([
            ft.Container(fault_icon_widget, bgcolor=C["card"], border_radius=R_BADGE, padding=3),
            ft.Column([
                ft.Text("故障码 / 警告码查询", size=18, weight=ft.FontWeight.BOLD, color=C["title"]),
                ft.Text("按驱动器系列查询报警码、故障码的故障机理与排除方法",
                        size=12.5, color=C["sub"]),
            ], spacing=6, expand=True),
            _tb("进入 →", on_click=_open_fault),
        ], spacing=16),
        bgcolor=C["card"],
        border_radius=R_CARD,
        padding=12 if is_mobile else 18,
        border=ft.border.all(1, C["border"]),
        on_click=_open_fault,
        ink=True,
        **_card_bg,
    )

    repl_icon_widget = (
        ft.Image(src=_asset("repl_icon.png"), width=40, height=40, fit=ft.ImageFit.CONTAIN)
        if web or os.path.exists(REPL_ICON)
        else ft.Text("🔄", size=22, text_align=ft.TextAlign.CENTER)
    )
    repl_card = ft.Container(
        ft.Row([
            ft.Container(repl_icon_widget,
                         bgcolor=C["card"], border_radius=R_BADGE, padding=3),
            ft.Column([
                ft.Text("产品替换选型", size=18, weight=ft.FontWeight.BOLD, color=C["title"]),
                ft.Text("客户原品牌型号逐项对照汇川选型，导出五列替换清单",
                        size=12.5, color=C["sub"]),
            ], spacing=6, expand=True),
            _tb("进入 →", on_click=_open_repl),
        ], spacing=16),
        bgcolor=C["card"],
        border_radius=R_CARD,
        padding=12 if is_mobile else 18,
        border=ft.border.all(1, C["border"]),
        on_click=_open_repl,
        ink=True,
        **_card_bg,
    )

    # ---------- 多端版本下载弹窗 ----------
    def _dl_row(item):
        return ft.Row([
            ft.Column([
                ft.Text(item["label"], size=14, weight=ft.FontWeight.BOLD, color=C["title"]),
                ft.Text(item["desc"], size=11, color=C["sub"]),
            ], spacing=4, expand=True),
            _tb("打开 ↗", on_click=lambda e, u=item["url"]: page.launch_url(u)),
        ], spacing=12)

    dl_content = ft.Column(
        [_dl_row(it) for it in DOWNLOAD_LINKS],
        spacing=14, scroll=ft.ScrollMode.AUTO, width=_w(460),
    )

    def _close_dl(ev):
        dl_dlg.open = False
        page.update()

    dl_dlg = ft.AlertDialog(
        open=False,
        title=ft.Text("其他端版本下载", size=16, weight=ft.FontWeight.BOLD, color=C["title"]),
        content=dl_content,
        actions=[_tb("关闭", on_click=_close_dl)],
        actions_alignment=ft.MainAxisAlignment.END,
    )

    def _open_downloads(e):
        dl_dlg.open = True
        page.dialog = dl_dlg
        page.update()

    dl_card = ft.Container(
        ft.Row([
            ft.Container(ft.Text("📱", size=22, text_align=ft.TextAlign.CENTER),
                         bgcolor=C["card"], border_radius=R_BADGE, padding=3),
            ft.Column([
                ft.Text("其他端版本下载", size=18, weight=ft.FontWeight.BOLD, color=C["title"]),
                ft.Text("安卓全能版 / 安卓轻量故障版 / 苹果轻量故障版，点击跳转下载",
                        size=12.5, color=C["sub"]),
            ], spacing=6, expand=True),
            _tb("获取 →", on_click=_open_downloads),
        ], spacing=16),
        bgcolor=C["card"],
        border_radius=R_CARD,
        padding=12 if is_mobile else 18,
        border=ft.border.all(1, C["border"]),
        on_click=_open_downloads,
        ink=True,
        **_card_bg,
    )


    home_children = []
    if cover_container is not None:
        home_children.append(cover_container)
    _title_children = []
    if mode != "brand":
        _title_children.append(ft.Text("汇川伺服工具集", size=24, weight=ft.FontWeight.BOLD, color=C["title"]))
    _title_children.append(ft.Container(
        ft.Text(f"选择下方应用进入　·　v{VERSION}　·　{cfg.get('watermark', '姓许的苏剑')}",
                size=12.5, color=C["muted"]),
        on_click=_egg_click, ink=True, padding=4,
    ))
    home_children += [
        ft.Column(_title_children, spacing=6),
        sel_card,
        fault_card,
        repl_card,
        dl_card,
        ft.Row([btn_log, btn_theme], spacing=12,
               alignment=ft.MainAxisAlignment.SPACE_BETWEEN if is_mobile else ft.MainAxisAlignment.END),
    ]
    home_view = ft.Container(
        ft.Column(
            home_children,
            spacing=18,
            scroll=ft.ScrollMode.AUTO,
            horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
            tight=True,
        ),
        padding=14 if is_mobile else 24,
        expand=True,
    )

    _show(home_view)
    refresh(None)

    # 主界面挂载完成后再弹出免责声明，避免“控件未加入页面即更新”报错
    if show_license:
        page.dialog = license_dlg
        license_dlg.open = True
    page.update()

    # 替换标题栏/任务栏默认图标为本地 logo（仅桌面端）
    if _is_desktop(page):
        _apply_window_icon(page.title)


def _main_inner(page: ft.Page):
    _write_diag(page)
    page.title = "汇川 MS1-R 伺服选型工具"
    page.scroll = ft.ScrollMode.ADAPTIVE

    def _on_window_event(e):
        # 从最小化恢复时：若坐标仍是最小化残留（落在屏外），拉回屏幕中央；并刷新一次渲染
        try:
            evt = getattr(e, "data", "")
            evt_name = evt.strip().lower() if isinstance(evt, str) else ""
            if evt_name == "restore":
                left = page.window_left
                top = page.window_top
                if left is None or top is None or left < -1000 or top < -1000:
                    page.window_center()
                page.update()
        except Exception:
            pass

    if _is_desktop(page):
        page.on_window_event = _on_window_event

    _ensure_runtime_data()
    _ensure_fault_db()
    FAULT_DB.clear()
    FAULT_DB.update(_load_fault_db())
    cfg = _load_settings()
    try:
        data = load_excel(DATA_FILE)
        _log(f"数据库加载完成：{len(data)} 条")
    except Exception as ex:
        _log(f"数据库加载异常：{ex}", "ERROR")
        data = []
    _ensure_original()
    base_dir = WORK_DIR

    def apply_theme(m):
        pal = _resolve_palette(m, cfg)
        _ff = cfg.get("font_family") or None
        page.theme_mode = ft.ThemeMode.DARK if m in DARK_MODES else ft.ThemeMode.LIGHT
        page.theme = ft.Theme(color_scheme_seed=pal["primary"], font_family=_ff)
        page.dark_theme = ft.Theme(color_scheme_seed=PALETTES["dark"]["primary"], font_family=_ff)
        page.bgcolor = pal["bg"]

    def _rebuild(m):
        if page.dialog is not None:
            page.dialog.open = False
        page.update()
        page.run_task(_delayed_rebuild, m)

    async def _delayed_rebuild(m):
        await asyncio.sleep(0.7)
        apply_theme(m)
        page.clean()
        _app(page, cfg, data, base_dir, _resolve_palette(m, cfg), m, rebuild=_rebuild, show_license=False)

    def _start_app(m):
        apply_theme(m)
        _app(page, cfg, data, base_dir, _resolve_palette(m, cfg), m, rebuild=_rebuild, show_license=True)

    saved_mode = cfg.get("theme_mode")
    if saved_mode not in THEME_MODES:
        saved_mode = ""

    if saved_mode:
        _start_app(saved_mode)
        return

    # ---------- 首次启动：免责声明之前先选择浅色 / 深色主题 ----------
    page.theme_mode = ft.ThemeMode.LIGHT
    page.theme = ft.Theme(color_scheme_seed=PALETTES["light"]["primary"])
    page.bgcolor = PALETTES["light"]["bg"]

    def _pick(m):
        cfg["theme_mode"] = m
        try:
            _save_settings(cfg)
            _log(f"首次选择外观主题：{m}")
        except Exception as ex:
            _log(f"主题保存失败：{ex}", "ERROR")
        try:
            if page.dialog is not None:
                page.dialog.open = False
            page.update()
            _start_app(m)
        except Exception:
            err = traceback.format_exc()
            try:
                os.makedirs(WORK_DIR, exist_ok=True)
                with open(os.path.join(WORK_DIR, "error.log"), "w", encoding="utf-8") as f:
                    f.write(err)
            except Exception:
                pass
            try:
                page.clean()
                page.add(ft.Container(
                    ft.Column([
                        ft.Text("进入工具出错", size=18, weight=ft.FontWeight.BOLD, color="red"),
                        ft.Text(err, size=11, color="#333333", selectable=True),
                    ], scroll=ft.ScrollMode.AUTO, spacing=10),
                    padding=16, expand=True,
                ))
                page.update()
            except Exception:
                pass

    def _theme_card(m):
        meta = THEME_META[m]
        mobile = _is_mobile(page)
        return ft.Container(
            ft.Column([
                ft.Text(meta["icon"], size=26, text_align=ft.TextAlign.CENTER),
                ft.Text(meta["label"], size=14, weight=ft.FontWeight.BOLD, text_align=ft.TextAlign.CENTER),
                ft.Text(meta["desc"], size=10.5, color="#5A6675", text_align=ft.TextAlign.CENTER),
            ], spacing=7, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            bgcolor="white", border_radius=R_CARD, padding=14,
            width=None if mobile else 150,
            border=ft.border.all(1, "#E3E9F2"), on_click=lambda e: _pick(m), ink=True,
        )

    _theme_cards = [_theme_card(m) for m in _visible_modes(PRESET_MODES)]
    if _is_mobile(page):
        _theme_body = ft.Column(_theme_cards, spacing=14,
                                horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
                                scroll=ft.ScrollMode.AUTO, height=420)
    else:
        _theme_body = ft.Column([
            ft.Row(_theme_cards[:3], spacing=14, wrap=True, alignment=ft.MainAxisAlignment.CENTER),
            ft.Row(_theme_cards[3:], spacing=14, wrap=True, alignment=ft.MainAxisAlignment.CENTER),
        ], spacing=14, horizontal_alignment=ft.CrossAxisAlignment.CENTER)

    pick_dlg = ft.AlertDialog(
        open=True,
        title=ft.Text("选择外观主题", size=16, weight=ft.FontWeight.BOLD, color="#123A5A"),
        content=ft.Column([
            ft.Text("首次使用请选择一套界面主题；之后可在封面「🎨 外观主题」中随时切换。",
                    size=12.5, color="#1F2733"),
            _theme_body,
        ], spacing=16, tight=True, width=None if _is_mobile(page) else 520),
    )
    page.dialog = pick_dlg
    page.update()


def main(page: ft.Page):
    try:
        _main_inner(page)
    except Exception:
        err = traceback.format_exc()
        try:
            os.makedirs(WORK_DIR, exist_ok=True)
            with open(os.path.join(WORK_DIR, "error.log"), "w", encoding="utf-8") as f:
                f.write(err)
        except Exception:
            pass
        try:
            page.clean()
            page.add(ft.Container(
                ft.Column([
                    ft.Text("启动出错", size=18, weight=ft.FontWeight.BOLD, color="red"),
                    ft.Text(err, size=11, color="#333333", selectable=True),
                ], scroll=ft.ScrollMode.AUTO, spacing=10),
                padding=16, expand=True,
            ))
            page.update()
        except Exception:
            pass


ft.app(target=main, assets_dir="assets")
