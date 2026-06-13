import sys, os, subprocess

# ── venv bootstrap (cross-platform) ──────────────────────────────────────────
# macOS/Linux: .venv/lib/pythonX.Y/site-packages
# Windows:     .venv/Lib/site-packages
_base = os.path.dirname(os.path.abspath(__file__))
if sys.platform == "win32":
    _venv = os.path.join(_base, ".venv", "Lib", "site-packages")
else:
    _ver  = f"python{sys.version_info.major}.{sys.version_info.minor}"
    _venv = os.path.join(_base, ".venv", "lib", _ver, "site-packages")
if os.path.exists(_venv) and _venv not in sys.path:
    sys.path.insert(0, _venv)

import csv, math, time, threading, socket
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog

try:
    import openpyxl
    HAS_XLSX = True
except ImportError:
    HAS_XLSX = False

# ── cross-platform ping ───────────────────────────────────────────────────────

def _subprocess_ping(ip):
    """Fallback ping using the OS ping binary — works on both Windows and macOS/Linux."""
    try:
        if sys.platform == "win32":
            cmd = ["ping", "-n", "1", "-w", "800", ip]
        else:
            cmd = ["ping", "-c", "1", "-W", "800", "-t", "1", ip]
        r = subprocess.run(cmd, stdout=subprocess.PIPE,
                           stderr=subprocess.DEVNULL, timeout=3)
        if r.returncode == 0:
            out = r.stdout.decode(errors="ignore")
            for token in out.split():
                t = token.lower().replace("time=", "").replace("time<", "") \
                         .replace("ms", "").rstrip("<")
                try:
                    return float(t)
                except ValueError:
                    pass
            return 1.0
    except Exception:
        pass
    return None

try:
    from ping3 import ping as _ping3
    def do_ping(ip, timeout=0.8):
        try:
            r = _ping3(ip, timeout=timeout)
            return round(r * 1000, 1) if r else None
        except PermissionError:
            # Windows requires Administrator for raw sockets — fall back gracefully
            return _subprocess_ping(ip)
        except Exception:
            return _subprocess_ping(ip)
except ImportError:
    def do_ping(ip, timeout=0.8):
        return _subprocess_ping(ip)

# ── config ───────────────────────────────────────────────────────────────────

def resource(name):
    return os.path.join(getattr(sys, "_MEIPASS", os.path.abspath(".")), name)

CSV_FILE    = resource("coop_locations_large.csv")
MAX_THREADS = 60

# ── colour palette ───────────────────────────────────────────────────────────

BG_MAP    = "#f0f4fa"
BG_PANEL  = "#f8f9fc"
BG_WIDGET = "#ffffff"
FG_TEXT   = "#1a1f36"
FG_SUB    = "#6b7280"
ACCENT    = "#2563eb"
SEP_COL   = "#e5e7eb"

C_ONLINE_CORE  = "#16a34a"
C_ONLINE_MID   = "#4ade80"
C_ONLINE_GLOW  = "#dcfce7"
C_OFFLINE_CORE = "#dc2626"
C_OFFLINE_GLOW = "#fee2e2"
C_SOURCE_CORE  = "#2563eb"
C_SOURCE_MID   = "#60a5fa"
C_SOURCE_GLOW  = "#dbeafe"
C_LINE_ON      = "#bbf7d0"
C_LINE_OFF     = "#fecaca"

DEPT_COLORS = [
    "#3b82f6","#8b5cf6","#f97316","#06b6d4",
    "#eab308","#ef4444","#22c55e","#6b7280",
    "#ec4899","#0ea5e9",
]

LOCK_ON_BG  = "#fef3c7"
LOCK_ON_FG  = "#92400e"
LOCK_OFF_BG = "#f3f4f6"
LOCK_OFF_FG = "#374151"

# ── colour helpers ────────────────────────────────────────────────────────────

def hex_lighter(h, amt=0.55):
    h = h.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    r = min(255, int(r + (255 - r) * amt))
    g = min(255, int(g + (255 - g) * amt))
    b = min(255, int(b + (255 - b) * amt))
    return f"#{r:02x}{g:02x}{b:02x}"

def hex_darker(h, amt=0.25):
    h = h.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    r = max(0, int(r * (1 - amt)))
    g = max(0, int(g * (1 - amt)))
    b = max(0, int(b * (1 - amt)))
    return f"#{r:02x}{g:02x}{b:02x}"

# ── data helpers ──────────────────────────────────────────────────────────────

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80)); ip = s.getsockname()[0]; s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def load_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls") and HAS_XLSX:
        return _from_xlsx(path)
    return _from_csv(path)

def _from_csv(path):
    rows = []
    if not os.path.exists(path):
        return rows
    with open(path, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            ip = r.get("IP", "").strip()
            if ip:
                rows.append({"IP": ip,
                             "Name": r.get("Name", "").strip(),
                             "Department": r.get("Department", "Unknown").strip(),
                             "Location": r.get("Location", "").strip()})
    return rows

def _from_xlsx(path):
    rows = []
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    hdrs = [str(c.value or "").strip() for c in next(ws.iter_rows(1, 1))]
    col  = {h: i for i, h in enumerate(hdrs)}
    for row in ws.iter_rows(min_row=2, values_only=True):
        ip = str(row[col.get("IP", 0)] or "").strip()
        if ip:
            rows.append({"IP": ip,
                         "Name": str(row[col.get("Name", 1)] or "").strip(),
                         "Department": str(row[col.get("Department", 2)] or "Unknown").strip(),
                         "Location": str(row[col.get("Location", 3)] or "").strip() if "Location" in col else ""})
    return rows

def sunflower(n, cx, cy, spacing=28):
    golden = (1 + math.sqrt(5)) / 2
    pts = []
    for i in range(n):
        r     = spacing * math.sqrt(i + 0.5)
        theta = 2 * math.pi * i / golden ** 2
        pts.append((cx + r * math.cos(theta), cy + r * math.sin(theta)))
    return pts

# ── tooltip ───────────────────────────────────────────────────────────────────

class Tooltip:
    def __init__(self, canvas):
        self._cv = canvas
        self._win = None

    def show(self, text, rx, ry):
        self.hide()
        self._win = tw = tk.Toplevel(self._cv)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{rx + 16}+{ry - 12}")
        outer = tk.Frame(tw, bg="#1e293b", padx=1, pady=1)
        outer.pack()
        tk.Label(outer, text=text, justify="left",
                 bg="#1e293b", fg="#f8fafc",
                 padx=10, pady=8, font=("Arial", 10),
                 relief="flat").pack()

    def hide(self):
        if self._win:
            try: self._win.destroy()
            except Exception: pass
            self._win = None

# ── map canvas ────────────────────────────────────────────────────────────────

class MapCanvas(tk.Canvas):
    R_SRC = 22
    R_DEV = 10

    def __init__(self, parent):
        super().__init__(parent, bg=BG_MAP, highlightthickness=0)
        self._zoom     = 1.0
        self._pan      = [0.0, 0.0]
        self._drag_xy  = None
        self._drag_ip  = None
        self._drag_off = (0, 0)
        self._locked   = False
        self._nodes    = {}
        self._hovered  = None
        self._tip      = Tooltip(self)

        self.bind("<ButtonPress-1>",   self._b1_down)
        self.bind("<B1-Motion>",       self._b1_drag)
        self.bind("<ButtonRelease-1>", self._b1_up)
        self.bind("<MouseWheel>",      self._wheel)
        self.bind("<Button-4>",        self._wheel)
        self.bind("<Button-5>",        self._wheel)
        self.bind("<Configure>",       lambda _: self._redraw())
        self.bind("<Motion>",          self._mouse_move)
        self.bind("<Leave>",           lambda _: self._tip.hide())

    # ── lock ──────────────────────────────────────────────────────────────────

    def set_locked(self, locked):
        self._locked = locked
        self.configure(cursor="hand2" if locked else "fleur")

    # ── coords ────────────────────────────────────────────────────────────────

    def _tc(self, lx, ly):
        w = self.winfo_width()  or 900
        h = self.winfo_height() or 600
        return (w/2 + (lx + self._pan[0]) * self._zoom,
                h/2 + (ly + self._pan[1]) * self._zoom)

    def _tl(self, cx, cy):
        w = self.winfo_width()  or 900
        h = self.winfo_height() or 600
        return ((cx - w/2) / self._zoom - self._pan[0],
                (cy - h/2) / self._zoom - self._pan[1])

    def _hit(self, ex, ey):
        hits = self.find_overlapping(ex - 8, ey - 8, ex + 8, ey + 8)
        for ip, d in self._nodes.items():
            if d.get("oval_id") in hits:
                return ip
        return None

    # ── interaction ───────────────────────────────────────────────────────────

    def _b1_down(self, e):
        self._tip.hide()
        ip = self._hit(e.x, e.y)
        if self._locked and ip:
            self._drag_ip  = ip
            self._drag_off = (e.x, e.y)
            self._drag_xy  = None
        elif not self._locked:
            self._drag_ip = None
            self._drag_xy = (e.x, e.y)

    def _b1_drag(self, e):
        if self._locked and self._drag_ip:
            dx = (e.x - self._drag_off[0]) / self._zoom
            dy = (e.y - self._drag_off[1]) / self._zoom
            self._nodes[self._drag_ip]["lx"] += dx
            self._nodes[self._drag_ip]["ly"] += dy
            self._drag_off = (e.x, e.y)
            self._redraw()
        elif not self._locked and self._drag_xy:
            dx = (e.x - self._drag_xy[0]) / self._zoom
            dy = (e.y - self._drag_xy[1]) / self._zoom
            self._pan[0] += dx
            self._pan[1] += dy
            self._drag_xy = (e.x, e.y)
            self._redraw()

    def _b1_up(self, e):
        self._drag_ip = None
        self._drag_xy = None

    def _wheel(self, e):
        delta = getattr(e, "delta", 0) or (120 if e.num == 4 else -120)
        f = 1.15 if delta > 0 else 1 / 1.15
        lx, ly = self._tl(e.x, e.y)
        self._zoom *= f
        w = self.winfo_width()  or 900
        h = self.winfo_height() or 600
        self._pan[0] = (e.x - w/2) / self._zoom - lx
        self._pan[1] = (e.y - h/2) / self._zoom - ly
        self._redraw()

    def fit_view(self):
        if not self._nodes: return
        xs = [d["lx"] for d in self._nodes.values()]
        ys = [d["ly"] for d in self._nodes.values()]
        self._pan = [-(min(xs)+max(xs))/2, -(min(ys)+max(ys))/2]
        w = self.winfo_width()  or 900
        h = self.winfo_height() or 600
        self._zoom = min(w / max(max(xs)-min(xs)+240, 1),
                         h / max(max(ys)-min(ys)+240, 1), 1.8)
        self._redraw()

    def _mouse_move(self, e):
        ip = self._hit(e.x, e.y)
        if ip != self._hovered:
            self._hovered = ip
            if ip:
                d   = self._nodes[ip]
                lat = d.get("latency")
                tip = (f"IP:        {ip}\n"
                       f"Name:      {d['name']}\n"
                       f"Dept:      {d['dept']}\n"
                       f"Location:  {d.get('location') or '—'}\n"
                       f"Status:    {'Online  ' + str(lat) + ' ms' if lat else 'Offline'}")
                self._tip.show(tip, self.winfo_rootx()+e.x, self.winfo_rooty()+e.y)
            else:
                self._tip.hide()

    # ── node data ─────────────────────────────────────────────────────────────

    def set_nodes(self, node_list, preserved_positions=None):
        old = {ip: (d["lx"], d["ly"]) for ip, d in self._nodes.items()}
        if preserved_positions:
            old.update(preserved_positions)
        self._nodes = {}
        for d in node_list:
            ip = d["ip"]
            lx, ly = old.get(ip, (d["lx"], d["ly"]))
            self._nodes[ip] = {
                "lx": lx, "ly": ly,
                "fill": d["fill"],
                "name": d["name"],
                "dept": d["dept"],
                "location": d.get("location", ""),
                "is_source": d.get("is_source", False),
                "latency": None,
                "oval_id": None,
            }
        self._redraw()

    def update_status(self, ip, latency):
        if ip in self._nodes:
            self._nodes[ip]["latency"] = latency

    # ── drawing ───────────────────────────────────────────────────────────────

    def _redraw(self):
        self.delete("all")
        if not self._nodes: return
        w = self.winfo_width()  or 900
        h = self.winfo_height() or 600

        # 1 ── dot grid background
        step = max(30, int(45 * self._zoom))
        ox   = int((self._pan[0] * self._zoom + w/2) % step)
        oy   = int((self._pan[1] * self._zoom + h/2) % step)
        for gx in range(ox - step, w + step, step):
            for gy in range(oy - step, h + step, step):
                self.create_oval(gx-1, gy-1, gx+1, gy+1,
                                 fill="#d1d9e6", outline="")

        # 2 ── source ripple rings (decorative circles around source)
        sx, sy = self._tc(0, 0)
        for ring_r in [60, 110, 170]:
            rr = ring_r * self._zoom
            self.create_oval(sx-rr, sy-rr, sx+rr, sy+rr,
                             outline="#c7d2e8", width=1, dash=(4, 6), fill="")

        # 3 ── cluster halos
        dept_nodes: dict[str, list] = {}
        for ip, d in self._nodes.items():
            if not d["is_source"]:
                dept_nodes.setdefault(d["dept"], []).append(d)

        dept_fill_cache = {}
        for dept, dnodes in dept_nodes.items():
            cx_l = sum(n["lx"] for n in dnodes) / len(dnodes)
            cy_l = sum(n["ly"] for n in dnodes) / len(dnodes)
            max_r = max((math.sqrt((n["lx"]-cx_l)**2 + (n["ly"]-cy_l)**2)
                         for n in dnodes), default=0)
            halo_r = (max_r + 30) * self._zoom
            hcx, hcy = self._tc(cx_l, cy_l)
            fill_c = dnodes[0]["fill"]
            dept_fill_cache[dept] = fill_c
            light_c = hex_lighter(fill_c, 0.80)
            border_c = hex_lighter(fill_c, 0.55)
            self.create_oval(hcx - halo_r, hcy - halo_r,
                             hcx + halo_r, hcy + halo_r,
                             fill=light_c, outline=border_c,
                             width=max(1, self._zoom * 1.2))

            # dept badge at halo top
            bx, by = hcx, hcy - halo_r - 4
            self._badge(bx, by, dept, fill_c, hex_darker(fill_c, 0.3))

        # 4 ── connection lines (source → each device)
        for ip, d in self._nodes.items():
            if d["is_source"]: continue
            lat = d.get("latency")
            tx, ty = self._tc(d["lx"], d["ly"])
            col  = C_LINE_ON  if lat is not None else C_LINE_OFF
            dash = ()         if lat is not None else (3, 8)
            self.create_line(sx, sy, tx, ty,
                             fill=col, width=max(1, self._zoom * 0.8),
                             dash=dash)

        # 5 ── nodes
        for ip, d in self._nodes.items():
            cx, cy = self._tc(d["lx"], d["ly"])
            self._draw_node(cx, cy, ip, d)

        # 6 ── node labels (visible when zoomed in)
        if self._zoom > 1.3:
            for ip, d in self._nodes.items():
                if d["is_source"]: continue
                cx, cy = self._tc(d["lx"], d["ly"])
                r = self.R_DEV * self._zoom
                self._node_label(cx, cy + r + 4, ip, d)

        # 7 ── legend card
        self._legend(w - 196, 14)

        # 8 ── lock banner
        if self._locked:
            self.create_rectangle(0, 0, w, 28,
                                  fill="#fef9c3", outline="#fde047", width=1)
            self.create_text(w//2, 14,
                             text="🔒  LOCK MODE ON  —  drag any node to reposition it",
                             fill="#92400e", font=("Arial", 10, "bold"))

    # ── rich node drawing ─────────────────────────────────────────────────────

    def _draw_node(self, cx, cy, ip, d):
        lat    = d.get("latency")
        is_src = d["is_source"]
        z      = self._zoom
        r      = (self.R_SRC if is_src else self.R_DEV) * z
        r      = max(r, 4)

        if is_src:
            # 3-layer source node
            self.create_oval(cx-(r+14*z), cy-(r+14*z),
                             cx+(r+14*z), cy+(r+14*z),
                             fill=C_SOURCE_GLOW, outline="#93c5fd",
                             width=max(1, z))
            self.create_oval(cx-(r+7*z), cy-(r+7*z),
                             cx+(r+7*z), cy+(r+7*z),
                             fill="#bfdbfe", outline=C_SOURCE_MID,
                             width=max(1.5, z*1.5))
            oid = self.create_oval(cx-r, cy-r, cx+r, cy+r,
                                   fill=C_SOURCE_CORE, outline="#1d4ed8",
                                   width=max(2, z*2))
            d["oval_id"] = oid
            # inner white shine
            self.create_oval(cx-(r*0.3), cy-(r*0.55),
                             cx+(r*0.3), cy-(r*0.05),
                             fill="white", outline="", width=0)
            # labels
            self.create_text(cx, cy - 2,
                             text="YOU", fill="white",
                             font=("Arial", max(7, int(9*z)), "bold"))
            self.create_text(cx, cy + r + 6*z,
                             text=ip, fill="#1e40af",
                             font=("Arial", max(6, int(7*z)), "bold"))
        else:
            fill  = d["fill"]
            if lat is not None:
                # ── ONLINE ── layered glow
                self.create_oval(cx-(r+9*z), cy-(r+9*z),
                                 cx+(r+9*z), cy+(r+9*z),
                                 fill=C_ONLINE_GLOW, outline=C_ONLINE_MID,
                                 width=max(1, z * 0.8))
                self.create_oval(cx-(r+4*z), cy-(r+4*z),
                                 cx+(r+4*z), cy+(r+4*z),
                                 fill=hex_lighter(fill, 0.6),
                                 outline=C_ONLINE_MID,
                                 width=max(1.2, z * 1.2))
                oid = self.create_oval(cx-r, cy-r, cx+r, cy+r,
                                       fill=fill, outline=C_ONLINE_CORE,
                                       width=max(2, z*2))
                # latency micro-badge (only if space)
                if z > 1.1:
                    ms = f"{lat}ms"
                    bw = len(ms) * max(4, int(5.5 * z))
                    bh = max(10, int(12 * z))
                    self.create_rectangle(cx - bw/2, cy - r - bh - 2*z,
                                          cx + bw/2, cy - r - 2*z,
                                          fill=C_ONLINE_CORE, outline="",
                                          width=0)
                    self.create_text(cx, cy - r - bh/2 - 2*z,
                                     text=ms, fill="white",
                                     font=("Arial", max(5, int(6*z)), "bold"))
            else:
                # ── OFFLINE ── muted glow
                self.create_oval(cx-(r+6*z), cy-(r+6*z),
                                 cx+(r+6*z), cy+(r+6*z),
                                 fill=C_OFFLINE_GLOW, outline="#fca5a5",
                                 width=max(1, z * 0.8))
                oid = self.create_oval(cx-r, cy-r, cx+r, cy+r,
                                       fill="#e5e7eb", outline=C_OFFLINE_CORE,
                                       width=max(2, z*2))
                # X mark
                o = r * 0.38
                self.create_line(cx-o, cy-o, cx+o, cy+o,
                                 fill=C_OFFLINE_CORE,
                                 width=max(1.5, z*1.5), capstyle="round")
                self.create_line(cx+o, cy-o, cx-o, cy+o,
                                 fill=C_OFFLINE_CORE,
                                 width=max(1.5, z*1.5), capstyle="round")

            d["oval_id"] = oid

            # white shine spot (online only)
            if lat is not None:
                self.create_oval(cx-(r*0.28), cy-(r*0.55),
                                 cx+(r*0.28), cy-(r*0.05),
                                 fill="white", outline="", width=0)

            # drag handle dot when locked
            if self._locked:
                self.create_oval(cx-2.5, cy-2.5, cx+2.5, cy+2.5,
                                 fill="white", outline=FG_SUB, width=1)

    def _node_label(self, cx, by, ip, d):
        """Pill label below a device node."""
        z    = self._zoom
        text = d["name"][:16] if z > 1.8 else ip
        fs   = max(6, int(7.5 * z))
        font = ("Arial", fs)
        # measure roughly
        tw = len(text) * fs * 0.62
        th = fs + 6
        px, py = 4, 2
        self.create_rectangle(cx - tw/2 - px, by,
                               cx + tw/2 + px, by + th + py*2,
                               fill="white", outline=SEP_COL, width=1)
        self.create_text(cx, by + th/2 + py,
                         text=text, fill=FG_TEXT, font=font)

    def _badge(self, cx, by, text, bg, fg):
        """Department name badge (pill at top of cluster halo)."""
        z    = self._zoom
        fs   = max(7, int(8 * z))
        font = ("Arial", fs, "bold")
        tw   = len(text) * fs * 0.68
        th   = fs + 6
        px, py = 6, 3
        self.create_rectangle(cx - tw/2 - px, by - th - py,
                               cx + tw/2 + px, by,
                               fill=bg, outline=hex_darker(bg, 0.15), width=1)
        self.create_text(cx, by - th/2 - py/2,
                         text=text, fill="white", font=font)

    def _legend(self, x, y):
        """Styled legend card."""
        pad = 12; lh = 24; bw = 182; bh = pad*2 + 16 + lh*3
        # shadow
        self.create_rectangle(x+3, y+3, x+bw+3, y+bh+3,
                              fill="#c8d0dc", outline="")
        # card
        self.create_rectangle(x, y, x+bw, y+bh,
                              fill="white", outline=SEP_COL, width=1)
        self.create_text(x+pad, y+pad,
                         text="LEGEND", fill=ACCENT,
                         font=("Arial", 9, "bold"), anchor="nw")
        items = [
            (C_ONLINE_CORE,  C_ONLINE_GLOW,  "Online"),
            (C_OFFLINE_CORE, C_OFFLINE_GLOW, "Offline"),
            (C_SOURCE_CORE,  C_SOURCE_GLOW,  "Source  (your machine)"),
        ]
        for i, (core, glow, lbl) in enumerate(items):
            oy = y + pad + 16 + lh * i
            self.create_oval(x+pad,     oy+3,  x+pad+16, oy+19,
                             fill=glow, outline=core, width=2)
            self.create_oval(x+pad+3,   oy+6,  x+pad+13, oy+16,
                             fill=core, outline="")
            self.create_text(x+pad+22, oy+4,
                             text=lbl, fill=FG_TEXT,
                             font=("Arial", 9), anchor="nw")

# ── main application ──────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Network Map Monitor")
        self.geometry("1380x800")
        self.minsize(900, 560)
        self.configure(bg=BG_PANEL)

        self._source_ip = get_local_ip()
        self._systems   = {}
        self._devices   = []
        self._executor  = ThreadPoolExecutor(max_workers=MAX_THREADS)
        self._search    = tk.StringVar()
        self._dept_var  = tk.StringVar(value="All Departments")
        self._locked    = False
        self._scanning  = False

        self._search.trace_add("write",   lambda *_: self._refresh_table())
        self._dept_var.trace_add("write", lambda *_: self._refresh_table())

        self._build_ui()
        self._load(CSV_FILE)
        self._schedule_scan()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", background=BG_PANEL, foreground=FG_TEXT,
                        fieldbackground=BG_WIDGET, font=("Arial", 11))
        style.configure("Treeview", background=BG_WIDGET, foreground=FG_TEXT,
                        fieldbackground=BG_WIDGET, rowheight=22)
        style.configure("Treeview.Heading", background="#e5e7eb",
                        foreground=FG_SUB, font=("Arial", 10, "bold"))
        style.map("Treeview", background=[("selected", "#dbeafe")])
        style.configure("TEntry",    fieldbackground=BG_WIDGET, foreground=FG_TEXT)
        style.configure("TCombobox", fieldbackground=BG_WIDGET, foreground=FG_TEXT)
        style.configure("TButton",   background="#e5e7eb", foreground=FG_TEXT, padding=(6,5))
        style.map("TButton", background=[("active", "#d1d5db")])

        left = tk.Frame(self, bg=BG_PANEL, width=295)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)

        def lbl(parent, text, color=FG_TEXT, size=11, bold=False):
            return tk.Label(parent, text=text, bg=BG_PANEL, fg=color,
                            font=("Arial", size, "bold" if bold else "normal"),
                            anchor="w", padx=12)

        lbl(left, "Network Map Monitor", ACCENT, 14, bold=True).pack(fill="x", pady=(12,2))
        self._lbl_src = lbl(left, f"Source IP:  {self._source_ip}", FG_SUB, 10)
        self._lbl_src.pack(fill="x")
        ttk.Separator(left).pack(fill="x", pady=8, padx=8)
        self._lbl_stats = lbl(left, "Loading…", FG_TEXT, 11, bold=True)
        self._lbl_stats.pack(fill="x")
        self._lbl_time = lbl(left, "", FG_SUB, 9)
        self._lbl_time.pack(fill="x", pady=(0,6))

        tk.Label(left, text="Search", bg=BG_PANEL, fg=FG_SUB,
                 font=("Arial",9), anchor="w", padx=12).pack(fill="x")
        ttk.Entry(left, textvariable=self._search).pack(fill="x", padx=10, pady=(0,6))
        tk.Label(left, text="Department", bg=BG_PANEL, fg=FG_SUB,
                 font=("Arial",9), anchor="w", padx=12).pack(fill="x")
        self._dept_cb = ttk.Combobox(left, textvariable=self._dept_var, state="readonly")
        self._dept_cb.pack(fill="x", padx=10, pady=(0,8))

        for txt, cmd in [("⟳  Scan Now", self._scan),
                          ("📂  Load CSV / Excel", self._load_dialog),
                          ("💾  Export CSV", self._export),
                          ("⊡  Fit Map", lambda: self._map.fit_view())]:
            ttk.Button(left, text=txt, command=cmd).pack(fill="x", padx=10, pady=2)

        self._lock_btn = tk.Button(
            left, text="🔓  Lock Map  (OFF)",
            bg=LOCK_OFF_BG, fg=LOCK_OFF_FG,
            font=("Arial", 11, "bold"),
            relief="flat", bd=0, padx=10, pady=6,
            cursor="hand2", command=self._toggle_lock)
        self._lock_btn.pack(fill="x", padx=10, pady=(6,2))

        ttk.Separator(left).pack(fill="x", pady=8, padx=8)
        tk.Label(left, text="Double-click row → jump to node",
                 bg=BG_PANEL, fg=FG_SUB, font=("Arial",9),
                 anchor="w", padx=12).pack(fill="x", pady=(0,4))

        cols = ("IP","Name","Latency","Status")
        self._tree = ttk.Treeview(left, columns=cols, show="headings", selectmode="browse")
        for c in cols:
            self._tree.heading(c, text=c)
            self._tree.column(c, width=60, anchor="center")
        self._tree.column("IP",   width=110, anchor="w")
        self._tree.column("Name", width=100, anchor="w")
        self._tree.tag_configure("online",  background="#dcfce7", foreground="#166534")
        self._tree.tag_configure("offline", background="#fee2e2", foreground="#991b1b")
        sb = ttk.Scrollbar(left, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=sb.set)
        self._tree.pack(side="left", fill="both", expand=True, padx=(10,0), pady=(0,10))
        sb.pack(side="left", fill="y", pady=(0,10), padx=(0,6))
        self._tree.bind("<Double-1>", self._jump_to)

        self._map = MapCanvas(self)
        self._map.pack(side="left", fill="both", expand=True)

    # ── lock ──────────────────────────────────────────────────────────────────

    def _toggle_lock(self):
        self._locked = not self._locked
        self._map.set_locked(self._locked)
        if self._locked:
            self._lock_btn.config(text="🔒  Lock Map  (ON)",
                                  bg=LOCK_ON_BG, fg=LOCK_ON_FG)
        else:
            self._lock_btn.config(text="🔓  Lock Map  (OFF)",
                                  bg=LOCK_OFF_BG, fg=LOCK_OFF_FG)
        self._map._redraw()

    # ── data ──────────────────────────────────────────────────────────────────

    def _load_dialog(self):
        ft = [("CSV","*.csv")]
        if HAS_XLSX: ft.append(("Excel","*.xlsx *.xls"))
        ft.append(("All","*.*"))
        path = filedialog.askopenfilename(title="Load Device File", filetypes=ft)
        if path: self._load(path)

    def _load(self, path):
        if not path or not os.path.exists(path): return
        self._devices = load_file(path)
        self._build_map()
        self._scan()

    def _build_map(self):
        preserved = {ip: (d["lx"], d["ly"]) for ip, d in self._map._nodes.items()}
        self._systems = {}
        depts      = sorted({d["Department"] for d in self._devices})
        dept_color = {d: DEPT_COLORS[i % len(DEPT_COLORS)] for i, d in enumerate(depts)}

        self._dept_cb["values"] = ["All Departments"] + depts
        self._dept_var.set("All Departments")

        n_depts = len(depts)
        RING    = max(260, n_depts * 60)
        node_list = [{"ip": self._source_ip, "name":"You","dept":"Source",
                      "location":"This machine","fill":C_SOURCE_CORE,
                      "lx":0,"ly":0,"is_source":True}]
        self._systems[self._source_ip] = {"name":"You","dept":"Source",
                                           "location":"","latency":1.0}

        dept_groups = {}
        for dev in self._devices:
            dept_groups.setdefault(dev["Department"], []).append(dev)

        for di, (dept, devs) in enumerate(dept_groups.items()):
            angle = 2*math.pi*di/n_depts - math.pi/2
            dcx   = RING * math.cos(angle)
            dcy   = RING * math.sin(angle)
            n     = len(devs)
            sp    = max(16, min(28, 160/math.sqrt(n+1)))
            for dev, (nx, ny) in zip(devs, sunflower(n, dcx, dcy, sp)):
                ip = dev["IP"]
                self._systems[ip] = {"name":dev["Name"],"dept":dept,
                                      "location":dev["Location"],"latency":None}
                node_list.append({"ip":ip,"name":dev["Name"],"dept":dept,
                                  "location":dev["Location"],
                                  "fill":dept_color[dept],
                                  "lx":nx,"ly":ny,"is_source":False})

        self._map.set_nodes(node_list, preserved_positions=preserved)
        self.after(150, self._map.fit_view)

    # ── ping ──────────────────────────────────────────────────────────────────

    def _schedule_scan(self):
        self._scan()
        self.after(5000, self._schedule_scan)

    def _scan(self):
        if self._scanning: return
        self._scanning = True
        threading.Thread(target=self._ping_all, daemon=True).start()

    def _ping_all(self):
        ips = [ip for ip in self._systems if ip != self._source_ip]
        futures = [self._executor.submit(lambda i: (i, do_ping(i)), ip) for ip in ips]
        for f in futures:
            try:
                ip, lat = f.result()
                self._systems[ip]["latency"] = lat
            except Exception: pass
        self._scanning = False
        self.after(0, self._on_done)

    def _on_done(self):
        for ip, d in self._systems.items():
            self._map.update_status(ip, d["latency"])
        self._map._redraw()
        online = sum(1 for ip,d in self._systems.items()
                     if ip != self._source_ip and d["latency"] is not None)
        total  = len(self._systems) - 1
        self._lbl_stats.config(
            text=f"Online: {online}   Offline: {total-online}   Total: {total}")
        self._lbl_time.config(
            text=f"Last scan: {datetime.now().strftime('%H:%M:%S')}")
        self._refresh_table()

    # ── table ─────────────────────────────────────────────────────────────────

    def _refresh_table(self):
        search = self._search.get().lower()
        dept_f = self._dept_var.get()
        self._tree.delete(*self._tree.get_children())
        for ip, d in self._systems.items():
            if ip == self._source_ip: continue
            if dept_f != "All Departments" and d["dept"] != dept_f: continue
            if search and not any(search in v.lower()
                                  for v in [ip, d["name"], d["dept"], d["location"]]):
                continue
            lat = d["latency"]
            self._tree.insert("","end", iid=ip,
                              values=(ip, d["name"],
                                      f"{lat} ms" if lat else "Timeout",
                                      "Online" if lat else "Offline"),
                              tags=("online" if lat else "offline",))

    def _jump_to(self, event):
        sel = self._tree.selection()
        if not sel: return
        nd = self._map._nodes.get(sel[0])
        if not nd: return
        self._map._pan  = [-nd["lx"], -nd["ly"]]
        self._map._zoom = max(self._map._zoom, 2.0)
        self._map._redraw()

    def _export(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV","*.csv")],
            initialfile="network_status.csv")
        if not path: return
        with open(path, "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(["IP","Name","Department","Location","Latency","Status"])
            for ip, d in self._systems.items():
                if ip == self._source_ip: continue
                lat = d["latency"]
                w.writerow([ip, d["name"], d["dept"], d["location"],
                            f"{lat} ms" if lat else "Timeout",
                            "Online" if lat else "Offline"])

if __name__ == "__main__":
    app = App()
    app.mainloop()
