#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发送 TG Matrix 用量日报：正文短提示 + PDF 附件。

  python tools/send_usage_digest.py
  python tools/send_usage_digest.py --force
  python tools/send_usage_digest.py --day 2026-09-09 --force
"""

from __future__ import annotations

import argparse
import io
import json
import math
import os
import re
import smtplib
import ssl
import struct
import sys
import zlib
from datetime import datetime, timedelta, timezone
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

_MYT = timezone(timedelta(hours=8))
ROOT = Path(__file__).resolve().parents[1]
JSON_DIR = ROOT / "Json"
FONT_CANDIDATES = [
    Path(r"C:\Windows\Fonts\msyh.ttc"),
    Path(r"C:\Windows\Fonts\simhei.ttf"),
    Path(r"C:\Windows\Fonts\simsun.ttc"),
    Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
    Path("/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc"),
    Path("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"),
    Path("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"),
]


def _env(name: str, default: str = "") -> str:
    return str(os.environ.get(name) or default).strip()


def _load_json(name: str) -> Dict[str, Any]:
    p = JSON_DIR / name
    if not p.exists():
        raise SystemExit(f"缺少 {p}")
    return json.loads(p.read_text(encoding="utf-8-sig"))


def _mail_cfg() -> Dict[str, Any]:
    if _env("DIGEST_SMTP_PASSWORD") or _env("DIGEST_ADMIN_PASSWORD"):
        to_raw = _env("DIGEST_TO_EMAILS")
        return {
            "admin_email": _env("DIGEST_ADMIN_EMAIL"),
            "admin_password": _env("DIGEST_ADMIN_PASSWORD"),
            "smtp_host": _env("DIGEST_SMTP_HOST") or "smtp.gmail.com",
            "smtp_port": int(_env("DIGEST_SMTP_PORT") or "587"),
            "smtp_user": _env("DIGEST_SMTP_USER"),
            "smtp_password": _env("DIGEST_SMTP_PASSWORD"),
            "from_email": _env("DIGEST_FROM_EMAIL") or _env("DIGEST_SMTP_USER"),
            "to_emails": [x.strip() for x in to_raw.split(",") if x.strip()],
        }
    return _load_json("usage_mail_config.json")


def _client():
    try:
        from supabase import create_client
    except ImportError as e:
        raise SystemExit("请先: python -m pip install supabase") from e
    url = _env("DIGEST_SUPABASE_URL")
    key = _env("DIGEST_SUPABASE_ANON_KEY")
    if not url or not key:
        cfg = _load_json("supabase_config.json")
        url = (cfg.get("url") or "").strip()
        key = (cfg.get("anon_key") or "").strip()
    if not url or not key:
        raise SystemExit("需要 DIGEST_SUPABASE_URL / DIGEST_SUPABASE_ANON_KEY 或 Json/supabase_config.json")
    return create_client(url, key)


def _sign_in(client, mail_cfg: Dict[str, Any]) -> None:
    email = (mail_cfg.get("admin_email") or "").strip()
    password = mail_cfg.get("admin_password") or ""
    if not email or not password:
        raise SystemExit("usage_mail_config.json 需要 admin_email / admin_password")
    try:
        res = client.auth.sign_in_with_password({"email": email, "password": password})
    except Exception as e:
        raise SystemExit(f"Matrix 管理员登录失败（{email}）\n{e}") from e
    if not getattr(res, "user", None):
        raise SystemExit("管理员登录失败")


def _digest_day(*, in_progress: bool) -> str:
    now = datetime.now(_MYT)
    if in_progress:
        return (now - timedelta(hours=9)).date().isoformat()
    if now.hour >= 9:
        return (now.date() - timedelta(days=1)).isoformat()
    return (now.date() - timedelta(days=2)).isoformat()


def _window(day: str) -> Tuple[datetime, datetime]:
    d = datetime.fromisoformat(day).date()
    start = datetime(d.year, d.month, d.day, 9, 0, 0, tzinfo=_MYT)
    return start, start + timedelta(hours=24)


def _as_dict(extra: Any) -> Dict[str, Any]:
    if isinstance(extra, dict):
        return extra
    if isinstance(extra, str):
        try:
            v = json.loads(extra)
            return v if isinstance(v, dict) else {}
        except Exception:
            return {}
    return {}


def _jint(d: Dict[str, Any], *keys: str) -> int:
    cur: Any = d
    for k in keys:
        if not isinstance(cur, dict):
            return 0
        cur = cur.get(k)
    try:
        return int(cur or 0)
    except (TypeError, ValueError):
        return 0


def _extra_acc(extra: Dict[str, Any]) -> Dict[str, Any]:
    acc = extra.get("accounts")
    return acc if isinstance(acc, dict) else {}


def _extra_sessions(extra: Dict[str, Any]) -> int:
    acc = _extra_acc(extra)
    return max(
        _jint(acc, "sessions"),
        _jint(extra, "sessions"),
        _jint(acc, "listed"),
        _jint(extra, "listed"),
        _jint(acc, "inventory"),
        _jint(extra, "inventory"),
    )


def _extra_score(extra: Dict[str, Any]) -> Tuple[int, int, int]:
    v = 1 if str(extra.get("hb_v") or "") == "2" else 0
    return (v, _extra_sessions(extra), len(_names_from_extra(extra)))


def _parse_ts(v: Any) -> Optional[datetime]:
    if not v:
        return None
    s = str(v).replace("Z", "+00:00")
    try:
        dt = datetime.fromisoformat(s)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(_MYT)
    except Exception:
        return None


def _natural_m(name: str) -> Tuple:
    m = re.match(r"^m(\d+)$", (name or "").strip(), re.I)
    if m:
        return (0, int(m.group(1)))
    return (1, (name or "").lower())


def _inventory(logged_names: set) -> List[Dict[str, Any]]:
    """本机 Sessions / JSON / 组队 / BanID 拼出小号名单。"""
    by: Dict[str, Dict[str, Any]] = {}

    def row(name: str) -> Dict[str, Any]:
        n = (name or "").strip()
        r = by.setdefault(
            n,
            {"name": n, "phone": "", "team": "", "risk": False, "ban": False, "logged": False},
        )
        return r

    try:
        raw = json.loads((JSON_DIR / "multi_accounts.json").read_text(encoding="utf-8-sig"))
        for a in raw.get("accounts") or []:
            if not isinstance(a, dict):
                continue
            n = str(a.get("name") or "").strip()
            if not n:
                continue
            r = row(n)
            r["phone"] = str(a.get("phone") or r["phone"])
            r["risk"] = bool(a.get("risk_flagged") or r["risk"])
    except Exception:
        pass
    p = ROOT / "sessions" / "multi"
    if p.is_dir():
        for f in p.glob("*.session"):
            stem = f.stem  # m001_1539...
            name = stem.split("_", 1)[0]
            r = row(name)
            if "_" in stem and not r["phone"]:
                r["phone"] = stem.split("_", 1)[1]
    try:
        raw = json.loads((JSON_DIR / "teams.json").read_text(encoding="utf-8-sig"))
        for t in raw.get("teams") or []:
            tname = str(t.get("name") or "")
            for m in t.get("members") or []:
                n = str(m or "").strip()
                if n:
                    r = row(n)
                    r["team"] = tname or r["team"]
    except Exception:
        pass
    bp = ROOT / "BanID"
    if bp.is_dir():
        for d in bp.iterdir():
            if not d.is_dir():
                continue
            meta = d / "ban_info.json"
            name = d.name
            if meta.exists():
                try:
                    info = json.loads(meta.read_text(encoding="utf-8"))
                    name = str(info.get("name") or name)
                except Exception:
                    pass
            r = row(name)
            r["ban"] = True
    logged = {str(x).strip() for x in (logged_names or set()) if str(x).strip()}
    for r in by.values():
        r["logged"] = r["name"] in logged
    return sorted(by.values(), key=lambda r: _natural_m(r["name"]))


def _names_from_extra(extra: Dict[str, Any]) -> List[str]:
    raw = extra.get("names_online")
    if not raw:
        acc = extra.get("accounts")
        raw = acc.get("names_online") if isinstance(acc, dict) else None
    out: List[str] = []
    if isinstance(raw, list):
        for n in raw:
            s = str(n or "").strip()
            if s:
                out.append(s)
    return out


def _hex_rgb(h: str) -> Tuple[int, int, int]:
    s = (h or "#94a3b8").strip().lstrip("#")
    if len(s) == 3:
        s = "".join(c * 2 for c in s)
    if len(s) < 6:
        return (148, 163, 184)
    return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))


def _png_rgba(width: int, height: int, rows: List[List[Tuple[int, int, int, int]]]) -> bytes:
    def chunk(tag: bytes, data: bytes) -> bytes:
        crc = zlib.crc32(tag + data) & 0xFFFFFFFF
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", crc)

    raw = b"".join(b"\x00" + bytes(b for px in row for b in px) for row in rows)
    ihdr = struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", zlib.compress(raw, 9))
        + chunk(b"IEND", b"")
    )


def _render_pie_png(parts: List[Tuple[Tuple[int, int, int], int]], size: int = 220) -> bytes:
    items = [(rgb, n) for rgb, n in parts if n > 0]
    bg = (255, 255, 255, 255)
    if not items:
        return _png_rgba(size, size, [[bg] * size for _ in range(size)])
    total = float(sum(n for _, n in items)) or 1.0
    cx = cy = (size - 1) / 2.0
    radius = size / 2.0 - 2.0
    two_pi = math.pi * 2.0
    rows: List[List[Tuple[int, int, int, int]]] = []
    for y in range(size):
        row: List[Tuple[int, int, int, int]] = []
        for x in range(size):
            dx = x - cx
            dy = y - cy
            dist = math.hypot(dx, dy)
            if dist > radius:
                row.append(bg)
                continue
            frac = (math.atan2(dx, -dy) % two_pi) / two_pi
            color = items[-1][0]
            acc = 0.0
            for rgb, n in items:
                acc += n / total
                if frac <= acc + 1e-9:
                    color = rgb
                    break
            row.append((color[0], color[1], color[2], 255))
        rows.append(row)
    return _png_rgba(size, size, rows)


def _fetch_rows(client, start: datetime, end: datetime) -> Tuple[List[dict], Dict[str, dict]]:
    ev = (
        client.table("usage_events")
        .select("*")
        .gte("occurred_at", start.isoformat())
        .lt("occurred_at", end.isoformat())
        .order("occurred_at")
        .limit(5000)
        .execute()
    )
    events = list(getattr(ev, "data", None) or [])
    pf = client.table("profiles").select("id,email,display_name,plan,role").execute()
    profiles = {str(r.get("id")): r for r in (getattr(pf, "data", None) or [])}
    return events, profiles


_FEAT_ZH = {
    "duo": "一对一",
    "multi": "一对多",
    "group": "群聊",
    "outreach": "私信来电",
    "invite": "拉群",
    "profile": "编辑用户",
    "member_backup": "成员备份",
}


def _build_stats(day: str, events: List[dict], profiles: Dict[str, dict]) -> Dict[str, Any]:
    start, end = _window(day)
    by_user: Dict[str, Dict[str, Any]] = {}
    task_log: List[Dict[str, Any]] = []
    for e in events:
        uid = str(e.get("user_id") or "")
        if not uid:
            continue
        u = by_user.setdefault(
            uid,
            {
                "user_id": uid,
                "names": set(),
                "last_extra": {},
                "best_extra": {},
                "heartbeats": 0,
                "busy": 0,
                "task_starts": 0,
                "task_finishes": 0,
                "features": {},
                "first": None,
                "last": None,
                "peak": 0,
            },
        )
        extra = _as_dict(e.get("extra"))
        kind = str(e.get("kind") or "")
        src = str(extra.get("source") or "")
        feat = str(e.get("feature") or "")
        ts = _parse_ts(e.get("occurred_at"))
        if ts:
            u["first"] = ts if u["first"] is None else min(u["first"], ts)
            u["last"] = ts if u["last"] is None else max(u["last"], ts)
        if kind == "heartbeat" and src != "login":
            u["heartbeats"] += 1
            u["last_extra"] = extra
            if _extra_score(extra) >= _extra_score(u.get("best_extra") or {}):
                u["best_extra"] = extra
            running = extra.get("running")
            if isinstance(running, list) and running:
                u["busy"] += 1
            try:
                u["peak"] = max(u["peak"], int(e.get("accounts_online") or 0))
            except (TypeError, ValueError):
                pass
        if kind == "task_start":
            u["task_starts"] += 1
            u["features"][feat] = u["features"].get(feat, 0) + 1
        if kind == "task_finish":
            u["task_finishes"] += 1
        if kind in ("task_start", "task_finish"):
            p0 = profiles.get(uid) or {}
            dur_s = "—"
            if kind == "task_finish":
                try:
                    sec = int(extra.get("duration_sec") or 0)
                except (TypeError, ValueError):
                    sec = 0
                if sec > 0:
                    dur_s = f"{sec // 60} 分" if sec >= 60 else f"{sec} 秒"
            try:
                nacc = int(e.get("accounts_online") or 0)
            except (TypeError, ValueError):
                nacc = 0
            task_log.append(
                {
                    "time": ts.strftime("%H:%M:%S") if ts else "",
                    "display_name": p0.get("display_name") or "",
                    "email": p0.get("email") or uid,
                    "action": "开始" if kind == "task_start" else "结束",
                    "feature": _FEAT_ZH.get(feat, feat or "其他"),
                    "label": extra.get("label") or e.get("task_id") or "",
                    "accounts": nacc,
                    "duration": dur_s,
                }
            )
        for n in _names_from_extra(extra):
            u["names"].add(n)

    users = []
    tot_new = tot_risk = tot_ban = tot_login = tot_sess = 0
    for uid, u in by_user.items():
        extra = u.get("best_extra") or u.get("last_extra") or {}
        acc = _extra_acc(extra)
        sess = _extra_sessions(extra)
        peak = int(u.get("peak") or 0)
        if sess <= 0 and peak > 0:
            sess = peak
        risk = max(_jint(acc, "risk"), _jint(extra, "risk"))
        ban = max(_jint(acc, "ban"), _jint(extra, "ban_count"), _jint(acc, "frozen"), _jint(extra, "frozen"))
        new = max(_jint(acc, "new_today"), _jint(extra, "new_today"))
        logged = len(u["names"])
        if logged <= 0 and peak > 0:
            logged = peak
        if sess > 0:
            logged = min(logged, sess)
        p = profiles.get(uid) or {}
        hours = 0.0
        if u["first"] and u["last"]:
            hours = round((u["last"] - u["first"]).total_seconds() / 3600.0 + 0.25, 1)
        feat = u["features"]
        row = {
            "email": p.get("email") or uid,
            "display_name": p.get("display_name") or "",
            "plan": p.get("plan") or "",
            "new": new,
            "risk": risk,
            "ban": ban,
            "logged": logged,
            "sessions": sess,
            "heartbeats": u["heartbeats"],
            "busy": u["busy"],
            "task_starts": u["task_starts"],
            "task_finishes": u["task_finishes"],
            "duo": feat.get("duo", 0),
            "multi": feat.get("multi", 0),
            "group": feat.get("group", 0),
            "outreach": feat.get("outreach", 0),
            "invite": feat.get("invite", 0),
            "features": feat,
            "teams": extra.get("teams") or acc.get("teams") or [],
            "first": u["first"].strftime("%H:%M") if u["first"] else "—",
            "last": u["last"].strftime("%H:%M") if u["last"] else "—",
            "hours": hours,
            "peak": peak,
            "names": u["names"],
        }
        users.append(row)
        tot_new += new
        tot_risk += risk
        tot_ban += ban
        tot_login += logged
        tot_sess += sess
    users.sort(key=lambda r: str(r["email"]))

    return {
        "day": day,
        "start": start,
        "end": end,
        "users": users,
        "n_users": len(users),
        "new": tot_new,
        "risk": tot_risk,
        "ban": tot_ban,
        "logged": tot_login,
        "sessions": tot_sess,
        "n_events": len(events),
        "task_log": task_log[:120],
    }


def _cjk_font() -> str:
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
    except ImportError as e:
        raise SystemExit("请先: python -m pip install reportlab") from e
    for p in FONT_CANDIDATES:
        if not p.exists():
            continue
        name = "DigestCJK"
        try:
            if p.suffix.lower() == ".ttc":
                pdfmetrics.registerFont(TTFont(name, str(p), subfontIndex=0))
            else:
                pdfmetrics.registerFont(TTFont(name, str(p)))
            return name
        except Exception:
            continue
    return "Helvetica"


def _write_pdf(stats: Dict[str, Any]) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    font = _cjk_font()
    styles = getSampleStyleSheet()
    title = ParagraphStyle("t", parent=styles["Title"], fontName=font, fontSize=18, leading=24, alignment=TA_CENTER)
    h2 = ParagraphStyle("h2", parent=styles["Heading2"], fontName=font, fontSize=13, leading=18, spaceBefore=10)
    body = ParagraphStyle("b", parent=styles["Normal"], fontName=font, fontSize=10, leading=14, alignment=TA_LEFT)
    small = ParagraphStyle("s", parent=styles["Normal"], fontName=font, fontSize=8, leading=11, textColor=colors.HexColor("#64748b"))
    cell = ParagraphStyle("c", parent=styles["Normal"], fontName=font, fontSize=8, leading=11)

    def P(text: Any, st=cell) -> Paragraph:
        return Paragraph(str(text if text is not None else ""), st)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=16 * mm,
        rightMargin=16 * mm,
        topMargin=14 * mm,
        bottomMargin=14 * mm,
        title=f"TG Matrix 日报 {stats['day']}",
    )
    start: datetime = stats["start"]
    end: datetime = stats["end"]
    story: List[Any] = []
    story.append(P(f"TG Matrix 日报 {stats['day']}", title))
    story.append(P(
        f"窗口 {start.strftime('%Y-%m-%d %H:%M')} ～ {end.strftime('%Y-%m-%d %H:%M')}（马来西亚）",
        small,
    ))
    story.append(Spacer(1, 8))

    kpis = [
        ["今日新增", "风控", "冻结/隔离", "已登录", "总账号"],
        [str(stats["new"]), str(stats["risk"]), str(stats["ban"]), str(stats["logged"]), str(stats["sessions"])],
    ]
    kt = Table(kpis, colWidths=[32 * mm] * 5)
    kt.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("FONTSIZE", (0, 1), (-1, 1), 16),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#dbeafe")),
        ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#fef3c7")),
        ("BACKGROUND", (2, 0), (2, 0), colors.HexColor("#fee2e2")),
        ("BACKGROUND", (3, 0), (3, 0), colors.HexColor("#dcfce7")),
        ("BACKGROUND", (4, 0), (4, 0), colors.HexColor("#f1f5f9")),
        ("BOX", (0, 0), (-1, -1), 0.4, colors.HexColor("#cbd5e1")),
        ("INNERGRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#e2e8f0")),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(kt)
    story.append(P(
        "顶栏与饼图 = 下面各用户明细逐行相加。已登录 = 该用户窗口内去重小号（无名单则用该用户心跳峰值）。"
        "总账号 = 该用户心跳上报的 session 数（没有则暂用该用户峰值）。不是发信电脑上的文件夹。",
        small,
    ))
    story.append(Spacer(1, 10))

    logged = int(stats["logged"] or 0)
    sess = int(stats["sessions"] or 0)
    n_new = int(stats["new"] or 0)
    n_risk = int(stats["risk"] or 0)
    n_ban = int(stats["ban"] or 0)
    # 圆饼 = 总账号。蓝/黄/红优先占格（与已登录可能重叠），剩余再分已登录、未登录。
    remain = max(sess, 0)
    def _take(n: int) -> int:
        nonlocal remain
        k = min(max(int(n or 0), 0), remain)
        remain -= k
        return k

    s_new = _take(n_new)
    s_risk = _take(n_risk)
    s_ban = _take(n_ban)
    s_logged = _take(logged)
    s_off = remain
    story.append(P("全站汇总", h2))
    if sess <= 0 and logged <= 0 and n_new <= 0 and n_risk <= 0 and n_ban <= 0:
        story.append(P("登录区暂无数据：心跳里还没有 session 计数。请重启后再发。", body))
    else:
        png = _render_pie_png(
            [
                (_hex_rgb("#2563eb"), s_new),
                (_hex_rgb("#d97706"), s_risk),
                (_hex_rgb("#dc2626"), s_ban),
                (_hex_rgb("#16a34a"), s_logged),
                (_hex_rgb("#94a3b8"), s_off),
            ],
            size=240,
        )
        pie_io = io.BytesIO(png)
        pie_io.name = "pie.png"
        img = Image(pie_io, width=52 * mm, height=52 * mm)
        legend = Table(
            [
                [P('<font color="#2563eb">■</font> 今日新增'), P(str(n_new))],
                [P('<font color="#d97706">■</font> 风控'), P(str(n_risk))],
                [P('<font color="#dc2626">■</font> 冻结/隔离'), P(str(n_ban))],
                [P('<font color="#16a34a">■</font> 已登录'), P(str(logged))],
                [P('<font color="#94a3b8">■</font> 未登录'), P(str(max(sess - logged, 0)))],
                [P("圆饼 = 总账号"), P(str(sess))],
            ],
            colWidths=[36 * mm, 22 * mm],
        )
        legend.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), font),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        wrap = Table([[img, legend]], colWidths=[58 * mm, 70 * mm])
        wrap.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
        story.append(wrap)
        story.append(P("圆饼合计 = 总账号。蓝/黄/红按五格数值从总账号里着色；绿=已登录；灰=其余未登录。", small))

    story.append(P("各用户明细", h2))
    head = [P("<b>显示名</b>"), P("<b>邮箱</b>"), P("<b>新增</b>"), P("<b>风控</b>"),
            P("<b>冻结</b>"), P("<b>已登录</b>"), P("<b>总账号</b>")]
    rows = [head]
    for u in stats["users"]:
        dn = u["display_name"] or str(u["email"]).split("@")[0]
        rows.append([
            P(dn), P(u["email"]), P(u["new"]), P(u["risk"]),
            P(u["ban"]), P(u["logged"]), P(u["sessions"]),
        ])
    if len(rows) == 1:
        rows.append([P("（本窗口无心跳）"), P(""), P(""), P(""), P(""), P(""), P("")])
    ut = Table(rows, colWidths=[28 * mm, 48 * mm, 16 * mm, 16 * mm, 16 * mm, 18 * mm, 18 * mm], repeatRows=1)
    ut.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e2e8f0")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#94a3b8")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (2, 1), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(ut)

    story.append(P("任务用量", h2))
    thead = [P("<b>显示名</b>"), P("<b>套餐</b>"), P("<b>首次</b>"), P("<b>末次</b>"),
             P("<b>约小时</b>"), P("<b>峰值</b>"), P("<b>心跳</b>"), P("<b>忙碌</b>"),
             P("<b>一对多</b>"), P("<b>群聊</b>"), P("<b>私信</b>"), P("<b>拉群</b>"), P("<b>任务</b>")]
    trows = [thead]
    for u in stats["users"]:
        dn = u["display_name"] or str(u["email"]).split("@")[0]
        trows.append([
            P(dn), P(u.get("plan") or ""), P(u.get("first") or "—"), P(u.get("last") or "—"),
            P(u.get("hours") or 0), P(u.get("peak") or 0), P(u.get("heartbeats") or 0), P(u.get("busy") or 0),
            P(u.get("multi") or 0), P(u.get("group") or 0), P(u.get("outreach") or 0), P(u.get("invite") or 0),
            P(f"{u.get('task_starts') or 0}/{u.get('task_finishes') or 0}"),
        ])
    if len(trows) == 1:
        trows.append([P("—")] * 13)
    tw = [22*mm, 14*mm, 14*mm, 14*mm, 14*mm, 12*mm, 12*mm, 12*mm, 14*mm, 12*mm, 12*mm, 12*mm, 16*mm]
    tt_task = Table(trows, colWidths=tw, repeatRows=1)
    tt_task.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e2e8f0")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#94a3b8")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]))
    story.append(tt_task)

    story.append(P("当天任务流水", h2))
    story.append(P("最多列出 120 条。耗时仅结束行有。开始行的「小号」是任务登记人数，不是此刻在线。", small))
    lhead = [P("<b>时间</b>"), P("<b>显示名</b>"), P("<b>邮箱</b>"), P("<b>动作</b>"),
             P("<b>功能</b>"), P("<b>说明</b>"), P("<b>小号</b>"), P("<b>耗时</b>")]
    lrows = [lhead]
    for t in stats.get("task_log") or []:
        dn = t.get("display_name") or str(t.get("email") or "").split("@")[0]
        lrows.append([
            P(t.get("time") or ""),
            P(dn),
            P(t.get("email") or ""),
            P(t.get("action") or ""),
            P(t.get("feature") or ""),
            P(t.get("label") or ""),
            P(t.get("accounts") if t.get("accounts") is not None else "—"),
            P(t.get("duration") or "—"),
        ])
    if len(lrows) == 1:
        lrows.append([P("当天没有任务起止记录"), P(""), P(""), P(""), P(""), P(""), P(""), P("")])
    lt = Table(
        lrows,
        colWidths=[20*mm, 28*mm, 42*mm, 12*mm, 16*mm, 28*mm, 14*mm, 14*mm],
        repeatRows=1,
    )
    lt.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e2e8f0")),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#94a3b8")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (3, 1), (4, -1), "CENTER"),
        ("ALIGN", (6, 1), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
    ]))
    story.append(lt)

    team_rows = [[P("<b>用户</b>"), P("<b>组队</b>"), P("<b>当前任务</b>"), P("<b>成员</b>"), P("<b>当时在线</b>")]]
    for u in stats["users"]:
        teams = u.get("teams") or []
        if not isinstance(teams, list):
            continue
        for t in teams:
            if not isinstance(t, dict):
                continue
            team_rows.append([
                P(u["email"]),
                P(t.get("name") or ""),
                P(t.get("task") or "待机"),
                P(t.get("members") or 0),
                P(t.get("online") or 0),
            ])
    if len(team_rows) > 1:
        story.append(P("组队登录区", h2))
        tt = Table(team_rows, colWidths=[48 * mm, 28 * mm, 32 * mm, 18 * mm, 22 * mm], repeatRows=1)
        tt.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), font),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e2e8f0")),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#94a3b8")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(tt)

    story.append(Spacer(1, 12))
    story.append(P(f"事件条数 {stats['n_events']} · 用户 {stats['n_users']}", small))
    doc.build(story)
    return buf.getvalue()


def _send_mail(mail_cfg: Dict[str, Any], to_emails: List[str], subject: str, body: str, pdf: bytes, pdf_name: str) -> None:
    host = (mail_cfg.get("smtp_host") or "smtp.gmail.com").strip()
    port = int(mail_cfg.get("smtp_port") or 587)
    user = (mail_cfg.get("smtp_user") or "").strip()
    password = (mail_cfg.get("smtp_password") or "").replace(" ", "")
    from_email = (mail_cfg.get("from_email") or user).strip()
    if not user or not password:
        raise SystemExit("usage_mail_config.json 需要 smtp_user / smtp_password")
    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"] = from_email
    msg["To"] = ", ".join(to_emails)
    msg.attach(MIMEText(body, "plain", "utf-8"))
    att = MIMEApplication(pdf, _subtype="pdf")
    att.add_header("Content-Disposition", "attachment", filename=pdf_name)
    msg.attach(att)
    ctx = ssl.create_default_context()
    with smtplib.SMTP(host, port, timeout=40) as smtp:
        smtp.ehlo()
        smtp.starttls(context=ctx)
        smtp.login(user, password)
        smtp.sendmail(from_email, to_emails, msg.as_string())


def _settings(client) -> Dict[str, Any]:
    res = (
        client.table("usage_report_settings")
        .select("admin_emails,enabled,last_sent_day")
        .eq("id", 1)
        .limit(1)
        .execute()
    )
    rows = getattr(res, "data", None) or []
    return rows[0] if rows else {"admin_emails": [], "enabled": True}


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="发送 TG Matrix 用量日报（PDF 附件）")
    ap.add_argument("--day", help="报表日 YYYY-MM-DD（09:00～次日 08:59）")
    ap.add_argument("--dry-run", action="store_true", help="只生成 PDF，不发信")
    ap.add_argument("--force", action="store_true", help="忽略已发过；无 --day 时发当前未结束窗口")
    args = ap.parse_args(argv)

    mail_cfg = _mail_cfg()
    client = _client()
    _sign_in(client, mail_cfg)

    day = args.day or _digest_day(in_progress=bool(args.force or args.dry_run))
    start, end = _window(day)
    print(f"报表日 {day}")
    print(f"窗口 {start.isoformat()} ～ {end.isoformat()}")

    st = _settings(client)
    if not bool(st.get("enabled", True)):
        print("usage_report_settings.enabled=false，跳过")
        return 0
    emails = [str(x).strip() for x in (st.get("admin_emails") or []) if str(x).strip()]
    for x in mail_cfg.get("to_emails") or []:
        s = str(x).strip()
        if s and s not in emails:
            emails.append(s)
    if not emails:
        print("没有收件人")
        return 1

    if not args.force and not args.dry_run:
        last = st.get("last_sent_day")
        if last and str(last) == day:
            print(f"last_sent_day={last} 已发过，跳过（加 --force 可重发）")
            return 0

    events, profiles = _fetch_rows(client, start, end)
    print(f"事件 {len(events)} 条")
    stats = _build_stats(day, events, profiles)
    print(
        f"用户 {stats['n_users']} · 新增 {stats['new']} · 风控 {stats['risk']} · "
        f"冻结 {stats['ban']} · 已登录 {stats['logged']} · 总账号 {stats['sessions']}"
    )
    pdf = _write_pdf(stats)
    pdf_name = f"TG_Matrix_日报_{day}.pdf"
    out_dir = ROOT / "release"
    out_dir.mkdir(exist_ok=True)
    local = out_dir / pdf_name
    local.write_bytes(pdf)
    print(f"PDF {local} ({len(pdf)} bytes)")

    if args.dry_run:
        return 0

    stamp = datetime.now(_MYT).strftime("%H:%M")
    subject = f"TG Matrix 日报 {day} · PDF {stamp}"
    body = (
        f"TG Matrix 日报 {day}\n\n"
        f"完整数据请查看附件：{pdf_name}\n\n"
        f"窗口：{start.strftime('%Y-%m-%d %H:%M')} ～ {end.strftime('%Y-%m-%d %H:%M')}（马来西亚）\n"
        f"今日新增 {stats['new']} · 风控 {stats['risk']} · 冻结/隔离 {stats['ban']} · "
        f"已登录 {stats['logged']} · 总账号 {stats['sessions']}\n"
    )
    _send_mail(mail_cfg, emails, subject, body, pdf, pdf_name)
    try:
        client.rpc("usage_mark_digest_sent", {"p_day": day}).execute()
    except Exception as e:
        print(f"邮件已发，标记 last_sent_day 失败: {e}")
    print(f"已发送 → {emails}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
