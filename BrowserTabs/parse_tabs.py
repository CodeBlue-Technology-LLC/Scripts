#!/usr/bin/env python3
"""
parse_tabs.py  -  Decode browser tab/history data collected from an RDS host.

Reads a staging folder produced by Collect-BrowserTabs.ps1 and writes a single
consolidated CSV of what each user had open / recently visited.

Layout expected under --input:
    <staging>\<user>\firefox\<profile>_recovery.jsonlz4
    <staging>\<user>\firefox\<profile>_sessionstore.jsonlz4   (optional)
    <staging>\<user>\chrome\Current Session
    <staging>\<user>\chrome\Current Tabs
    <staging>\<user>\chrome\History            (sqlite, fallback)
    <staging>\<user>\edge\Current Session
    <staging>\<user>\edge\Current Tabs
    <staging>\<user>\edge\History

Firefox current tabs are reliable. Chrome/Edge "current tabs" come from the
SNSS session files and are BEST-EFFORT (undocumented binary format). The
History SQLite DB is the reliable fallback for those two - it gives visited
URLs (with timestamps), not live tab state.

Usage:
    pip install lz4
    python parse_tabs.py --input .\staging --output browser_tabs.csv

Only stdlib + lz4 required. If lz4 is missing, Firefox parsing falls back to a
pure-python LZ4 decoder (slower but no dependency).
"""

import argparse
import csv
import glob
import json
import os
import re
import sqlite3
import struct
import sys
from datetime import datetime, timezone, timedelta

# ---------- Firefox: mozLz4 ----------

MOZ_MAGIC = b"mozLz40\0"

def _lz4_decompress(payload, decompressed_size):
    """Try the lz4 library; fall back to a pure-python LZ4 block decoder."""
    try:
        import lz4.block
        return lz4.block.decompress(payload, uncompressed_size=decompressed_size)
    except ImportError:
        return _lz4_block_pure(payload, decompressed_size)

def _lz4_block_pure(src, out_size):
    """Minimal pure-python LZ4 block decompressor (no dependencies)."""
    out = bytearray()
    i = 0
    n = len(src)
    while i < n:
        token = src[i]; i += 1
        lit = token >> 4
        if lit == 15:
            while True:
                b = src[i]; i += 1
                lit += b
                if b != 255:
                    break
        out += src[i:i+lit]; i += lit
        if i >= n:
            break
        offset = src[i] | (src[i+1] << 8); i += 2
        match = token & 0x0F
        if match == 15:
            while True:
                b = src[i]; i += 1
                match += b
                if b != 255:
                    break
        match += 4
        start = len(out) - offset
        for j in range(match):
            out.append(out[start + j])
    return bytes(out[:out_size]) if out_size else bytes(out)

def read_mozlz4(path):
    with open(path, "rb") as f:
        magic = f.read(8)
        if magic != MOZ_MAGIC:
            raise ValueError(f"not a mozLz4 file (bad magic): {path}")
        size = struct.unpack("<I", f.read(4))[0]
        payload = f.read()
    return _lz4_decompress(payload, size)

def parse_firefox(path):
    """Return list of dicts: {window, tab_index, title, url, active}."""
    raw = read_mozlz4(path)
    data = json.loads(raw.decode("utf-8", errors="replace"))
    rows = []
    for wi, window in enumerate(data.get("windows", []), start=1):
        for ti, tab in enumerate(window.get("tabs", []), start=1):
            entries = tab.get("entries", [])
            if not entries:
                continue
            idx = tab.get("index", len(entries)) - 1
            idx = max(0, min(idx, len(entries) - 1))
            cur = entries[idx]
            rows.append({
                "window": wi,
                "tab_index": ti,
                "title": cur.get("title", ""),
                "url": cur.get("url", ""),
                "active": tab.get("selected") is True,
            })
    return rows

# ---------- Chrome / Edge: SNSS session reconstruction ----------
#
# SNSS is an append-only command log. To get *currently open* tabs we must
# replay the commands, not scrape URLs (scraping picks up closed tabs and
# back/forward history). We parse the Session_ file's command stream and honor
# tab-open / tab-closed / window-closed, then resolve each surviving tab to its
# selected navigation entry.
#
# Command IDs (components/sessions/core/session_service_commands.cc; stable):
CMD_SET_TAB_WINDOW              = 0   # struct: int32 window_id, int32 tab_id
CMD_UPDATE_TAB_NAVIGATION      = 6   # pickle: tab_id, index, url(str), title(str16)
CMD_SET_SELECTED_NAV_INDEX     = 7   # struct: int32 tab_id, int32 index
CMD_TAB_CLOSED                 = 16  # struct: int32 tab_id (+padding, int64 time)
CMD_WINDOW_CLOSED              = 17  # struct: int32 window_id (+ int64 time)

def _read_update_tab_nav(payload):
    """Parse a kCommandUpdateTabNavigation pickle payload -> (tab_id, index, url, title)."""
    p = payload
    off = 4  # skip 4-byte Pickle header (payload length)
    def rint():
        nonlocal off
        v = struct.unpack_from("<i", p, off)[0]; off += 4; return v
    def rstr():
        nonlocal off
        n = struct.unpack_from("<i", p, off)[0]; off += 4
        if n < 0 or off + n > len(p):
            raise ValueError("bad string length")
        s = p[off:off+n].decode("utf-8", "replace"); off += n
        off += (-n) % 4  # Pickle pads to 4-byte alignment
        return s
    def rstr16():
        nonlocal off
        n = struct.unpack_from("<i", p, off)[0]; off += 4
        b = n * 2
        if n < 0 or off + b > len(p):
            raise ValueError("bad string16 length")
        s = p[off:off+b].decode("utf-16-le", "replace"); off += b
        off += (-b) % 4
        return s
    tab_id = rint()
    index = rint()
    url = rstr()
    title = rstr16()
    return tab_id, index, url, title

def reconstruct_session(path):
    """Replay an SNSS Session_ file and return currently-open tabs in order."""
    try:
        with open(path, "rb") as f:
            data = f.read()
    except OSError:
        return []
    n = len(data)
    i = 8 if data[:4] == b"SNSS" else 4  # magic + int32 version, else just version

    tabs = {}            # tab_id -> {navs, sel, window, order, closed}
    windows_closed = set()
    order = [0]

    def get_tab(t):
        tb = tabs.get(t)
        if tb is None:
            tb = {"navs": {}, "sel": None, "window": None, "order": None, "closed": False}
            tabs[t] = tb
        if tb["order"] is None:
            tb["order"] = order[0]; order[0] += 1
        return tb

    while i + 2 <= n:
        size = struct.unpack_from("<H", data, i)[0]; i += 2
        if size == 0 or i + size > n:
            break
        cmd = data[i:i+size]; i += size
        cid = cmd[0]; payload = cmd[1:]
        try:
            if cid == CMD_SET_TAB_WINDOW and len(payload) >= 8:
                w, t = struct.unpack_from("<ii", payload, 0)
                get_tab(t)["window"] = w
            elif cid == CMD_UPDATE_TAB_NAVIGATION:
                t, idx, url, title = _read_update_tab_nav(payload)
                get_tab(t)["navs"][idx] = (url, title)
            elif cid == CMD_SET_SELECTED_NAV_INDEX and len(payload) >= 8:
                t, idx = struct.unpack_from("<ii", payload, 0)
                get_tab(t)["sel"] = idx
            elif cid == CMD_TAB_CLOSED and len(payload) >= 4:
                t = struct.unpack_from("<i", payload, 0)[0]
                if t in tabs:
                    tabs[t]["closed"] = True
            elif cid == CMD_WINDOW_CLOSED and len(payload) >= 4:
                windows_closed.add(struct.unpack_from("<i", payload, 0)[0])
        except (struct.error, ValueError, UnicodeDecodeError):
            continue

    result = []
    for t, tb in tabs.items():
        if tb["closed"] or tb["window"] in windows_closed or not tb["navs"]:
            continue
        sel = tb["sel"]
        if sel is None or sel not in tb["navs"]:
            sel = max(tb["navs"].keys())
        url, title = tb["navs"][sel]
        if not url:
            continue
        result.append({"window": tb["window"] if tb["window"] is not None else 0,
                       "order": tb["order"], "url": url, "title": title})
    result.sort(key=lambda r: (r["window"], r["order"]))
    return result

# ---------- Chrome / Edge: History SQLite (reliable fallback) ----------

CHROME_EPOCH = datetime(1601, 1, 1, tzinfo=timezone.utc)

def chrome_time(microseconds):
    if not microseconds:
        return ""
    try:
        return (CHROME_EPOCH + timedelta(microseconds=microseconds)).strftime("%Y-%m-%d %H:%M:%S")
    except OverflowError:
        return ""

def parse_history(path, days=7, limit=500):
    """Return recent visits from a copied Chrome/Edge History DB."""
    rows = []
    uri = f"file:{path}?mode=ro&immutable=1"
    try:
        con = sqlite3.connect(uri, uri=True)
    except sqlite3.Error as e:
        return [{"_error": f"open failed: {e}"}]
    try:
        cur = con.cursor()
        cutoff = 0
        if days:
            dt = datetime.now(timezone.utc) - timedelta(days=days)
            cutoff = int((dt - CHROME_EPOCH).total_seconds() * 1_000_000)
        cur.execute("""
            SELECT v.visit_time, u.url, u.title
            FROM visits v JOIN urls u ON u.id = v.url
            WHERE v.visit_time > ?
            ORDER BY v.visit_time DESC
            LIMIT ?
        """, (cutoff, limit))
        for vt, url, title in cur.fetchall():
            rows.append({"visit_time": chrome_time(vt), "url": url, "title": title or ""})
    except sqlite3.Error as e:
        rows.append({"_error": f"query failed: {e}"})
    finally:
        con.close()
    return rows

def build_title_map(history_path):
    """url -> title lookup from the History DB, to enrich SNSS tabs (no titles)."""
    m = {}
    uri = f"file:{history_path}?mode=ro&immutable=1"
    try:
        con = sqlite3.connect(uri, uri=True)
        for url, title in con.execute("SELECT url, title FROM urls WHERE title != ''"):
            m[url] = title
            m[url.rstrip("/")] = title
        con.close()
    except sqlite3.Error:
        pass
    return m

# ---------- Driver ----------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="staging folder from the PS collector")
    ap.add_argument("--output", default="browser_tabs.csv")
    ap.add_argument("--history-days", type=int, default=7,
                    help="how far back to pull Chrome/Edge history (0 = all)")
    args = ap.parse_args()

    out_rows = []

    for user_dir in sorted(glob.glob(os.path.join(args.input, "*"))):
        if not os.path.isdir(user_dir):
            continue
        user = os.path.basename(user_dir)

        # --- Firefox ---
        for ff in glob.glob(os.path.join(user_dir, "firefox", "*.jsonlz4")):
            profile = os.path.basename(ff).split("_")[0]
            try:
                for r in parse_firefox(ff):
                    out_rows.append({
                        "user": user, "browser": "firefox", "source": "live_tabs",
                        "profile": profile, "window": r["window"], "position": r["tab_index"],
                        "active": r["active"], "timestamp": "",
                        "title": r["title"], "url": r["url"],
                    })
            except Exception as e:
                out_rows.append({"user": user, "browser": "firefox", "source": "ERROR",
                                 "profile": profile, "title": str(e), "url": ff})

        # --- Chrome / Edge live tabs (best-effort SNSS) ---
        for br in ("chrome", "edge"):
            bdir = os.path.join(user_dir, br)
            if not os.path.isdir(bdir):
                continue

            hp = os.path.join(bdir, "History")
            title_map = build_title_map(hp) if os.path.exists(hp) else {}

            # Currently-open tabs come from the "Session" file (full session
            # state). NOTE: "Tabs_*" / "Current Tabs" is the tab-RESTORE log
            # (recently CLOSED tabs) - do NOT use it for open-tab counts.
            # Newest Session file wins.
            snss_path, snss_label = None, None
            for pat, label in (("Current Session", "Current Session"),
                               ("Session_*", "Session")):
                matches = [p for p in glob.glob(os.path.join(bdir, pat)) if os.path.isfile(p)]
                if matches:
                    snss_path = max(matches, key=os.path.getmtime)
                    snss_label = label
                    break

            if snss_path:
                for i, tab in enumerate(reconstruct_session(snss_path), start=1):
                    title = tab["title"] or title_map.get(tab["url"].rstrip("/"),
                                                           title_map.get(tab["url"], ""))
                    out_rows.append({
                        "user": user, "browser": br,
                        "source": f"live_tabs({snss_label})",
                        "profile": "", "window": tab["window"], "position": i, "active": "",
                        "timestamp": "", "title": title, "url": tab["url"],
                    })

            # --- History fallback ---
            if os.path.exists(hp):
                for r in parse_history(hp, days=args.history_days):
                    if "_error" in r:
                        out_rows.append({"user": user, "browser": br, "source": "history_ERROR",
                                         "title": r["_error"], "url": hp})
                        continue
                    out_rows.append({
                        "user": user, "browser": br, "source": "history",
                        "profile": "", "window": "", "position": "", "active": "",
                        "timestamp": r["visit_time"], "title": r["title"], "url": r["url"],
                    })

    cols = ["user", "browser", "source", "profile", "window", "position",
            "active", "timestamp", "title", "url"]
    with open(args.output, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        for r in out_rows:
            w.writerow(r)

    print(f"Wrote {len(out_rows)} rows to {args.output}")

if __name__ == "__main__":
    main()