# org_chart_app.py
# -*- coding: utf-8 -*-
"""
Gelişmiş Masaüstü Organizasyon Şeması
- Büyütülmüş toolbar
- Excel/CSV seçerek yükleme + (varsa) Sürükle-Bırak
- BG/Şerit renklendirme, Arama (merkeze kaydırma), Zoom
- PNG/PDF dışa aktarım

Gerekli: pandas, pillow, openpyxl
Opsiyonel DnD: tkinterdnd2
"""

import os, sys, json
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from io import BytesIO
from textwrap import wrap

# ---- Opsiyonel DnD ----
DnDEnabled = False
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DnDEnabled = True
except Exception:
    DnDEnabled = False

APP_TITLE = "Organizasyon Şeması"
REQ_COLS = ["Kullanıcı Adı","Ad Soyad","Departman","Pozisyon","Üst Kademe","Mail"]

# Kart metrikleri
BOX_W, BOX_H = 300, 88
BOX_PAD_X = 44
ROW_GAP, COL_GAP = 90, 42

CANVAS_BG   = (248, 250, 252)
BOX_BG      = (255, 255, 255)
BOX_BORDER  = (39, 94, 254)
LINE_COLOR  = (120, 144, 156)
TEXT_COLOR  = (26, 32, 44)
SUBTEXT     = (88, 96, 108)
HILIGHT     = (220, 53, 69)

FONT_TITLE_SIZE, FONT_SUB_SIZE = 17, 13
DEFAULT_SETTINGS_FILE = os.path.join(os.path.expanduser("~"), ".org_chart_settings.json")

@dataclass
class Person:
    username: str
    full_name: str
    dept: str
    title: str
    manager: Optional[str]
    mail: str

# ---------------- Yardımcılar ----------------
def ensure_font(size: int) -> ImageFont.FreeTypeFont:
    candidates = [
        "DejaVuSans.ttf",
        "C:\\Windows\\Fonts\\DejaVuSans.ttf",
        "C:\\Windows\\Fonts\\arial.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for p in candidates:
        try: return ImageFont.truetype(p, size)
        except: pass
    return ImageFont.load_default()

def load_table(path: str) -> pd.DataFrame:
    p = path.lower()
    df = pd.read_excel(path) if p.endswith((".xlsx",".xls")) else pd.read_csv(path)
    m = {}
    for c in df.columns:
        k = c.strip().lower()
        if "kullanıcı" in k and "adı" in k: m[c] = "Kullanıcı Adı"
        elif k in ["ad soyad","ad-soyad","ad_soyad","isim","ad","soyad","adi soyadi","adı soyadı"]: m[c] = "Ad Soyad"
        elif "departman" in k or "department" in k: m[c] = "Departman"
        elif "pozisy" in k or "görev" in k or "gorev" in k or "title" in k or "ünvan" in k or "unvan" in k: m[c] = "Pozisyon"
        elif "üst" in k or "ust" in k or "manager" in k or "yönetici" in k or "yonetici" in k or "amiri" in k: m[c] = "Üst Kademe"
        elif "mail" in k or "e-posta" in k or "email" in k: m[c] = "Mail"
    df = df.rename(columns=m)
    missing = [c for c in REQ_COLS if c not in df.columns]
    if missing: raise ValueError(f"Eksik kolonlar: {missing}\nMevcut: {list(df.columns)}")
    for c in REQ_COLS:
        df[c] = df[c].astype(str).str.strip().replace({"nan":None,"None":None,"":None})
    return df[REQ_COLS].copy()

def build_people(df: pd.DataFrame, dept: Optional[str]=None) -> Dict[str,Person]:
    if dept and dept != "Tümü": df = df.loc[df["Departman"]==dept].copy()
    out: Dict[str,Person] = {}
    for _,r in df.iterrows():
        u = (r["Kullanıcı Adı"] or "").strip()
        if not u or u.lower()=="none": continue
        out[u] = Person(
            username=u,
            full_name=(r["Ad Soyad"] or "").strip(),
            dept=(r["Departman"] or "").strip(),
            title=(r["Pozisyon"] or "").strip(),
            manager=((r["Üst Kademe"] or "").strip() or None),
            mail=(r["Mail"] or "").strip()
        )
    return out

def build_tree(people: Dict[str,Person]):
    children = {u:[] for u in people}
    has_mgr = set()
    for u,p in people.items():
        if p.manager and p.manager in people:
            children[p.manager].append(u); has_mgr.add(u)
    for k in children: children[k].sort(key=lambda x: people[x].full_name.lower())
    roots = [u for u in people if u not in has_mgr]; roots.sort(key=lambda x: people[x].full_name.lower())
    return children, roots

def compute_layout(children: Dict[str,List[str]], root: str):
    width = {}
    def w(u):
        if not children.get(u): width[u]=1; return 1
        total = sum(w(v) for v in children[u]); width[u]=max(1,total); return width[u]
    w(root)
    pos = {}
    def place(u, depth, x_left):
        myw = width[u]; x_center = x_left + myw//2; pos[u]=(x_center, depth)
        nxt = x_left
        for v in children.get(u, []): nxt = place(v, depth+1, nxt)
        return x_left + myw
    place(root, 0, 0)
    return pos

# -------------- Çizim --------------
def rel_luma(rgb):
    r,g,b = [x/255.0 for x in rgb]
    def adj(v): return v/12.92 if v<=0.03928 else ((v+0.055)/1.055)**2.4
    r,g,b = adj(r), adj(g), adj(b)
    return 0.2126*r + 0.7152*g + 0.0722*b

def text_color_for(bg_rgb): return (255,255,255) if rel_luma(bg_rgb) < 0.55 else (26,32,44)
def darker(rgb, f=0.80): return tuple(max(0, int(c*f)) for c in rgb)

def draw_chart(people, children, roots, start_from, scale, title_colors, show_dept, show_mail, highlight_user, color_style, add_legend=True):
    """
    Şemayı çizer ve (img, bboxes) döner.
    bboxes: {username: (x0,y0,x1,y1)} piksel koordinatları (render ölçeğinde)
    """
    if not people:
        return Image.new("RGB",(900,600), CANVAS_BG), {}

    draw_roots = [start_from] if (start_from and start_from in people) else roots
    if not draw_roots: draw_roots = list(people.keys())[:1]

    BW, BH = int(BOX_W*scale), int(BOX_H*scale)
    PX, RG, CG = int(BOX_PAD_X*scale), int(ROW_GAP*scale), int(COL_GAP*scale)
    FTS, FSS = max(11,int(FONT_TITLE_SIZE*scale)), max(10,int(FONT_SUB_SIZE*scale))
    font_title, font_sub = ensure_font(FTS), ensure_font(FSS)

    layouts = []
    total_cols, max_depth = 0, 0
    for r in draw_roots:
        layout = compute_layout(children, r)
        cols  = max([x for x,_ in layout.values()]+[0]) + 1
        depth = max([y for _,y in layout.values()]+[0]) + 1
        total_cols += cols; max_depth = max(max_depth, depth)
        layouts.append(layout)

    # Legend (tekilleştirilmiş)
    legend_w = 0; legend_items=[]
    if add_legend and title_colors:
        uniq_titles = sorted({p.title for p in people.values() if title_colors.get(p.title)})
        legend_items = [(t, title_colors[t]) for t in uniq_titles]
        if legend_items: legend_w = int(260*scale)

    img_w = PX*2 + total_cols*BW + (total_cols-1)*CG + legend_w
    img_h = PX*2 + max_depth*BH + (max_depth-1)*RG
    img = Image.new("RGB",(img_w,img_h), CANVAS_BG)
    d = ImageDraw.Draw(img)

    bboxes: Dict[str, Tuple[int,int,int,int]] = {}

    def draw_box(xc, yc, person: Person, highlight: bool):
        x0 = PX + xc*(BW+CG); y0 = PX + yc*(BH+RG)
        x1, y1 = x0+BW, y0+BH
        bboxes[person.username] = (x0,y0,x1,y1)
        radius = int(18*scale)

        pos_col = title_colors.get(person.title)
        if color_style == "bg" and pos_col:
            fill_col = pos_col
            border_col = darker(pos_col, 0.75)
            txt_main = text_color_for(fill_col)
            txt_sub  = text_color_for(tuple(min(255,int(c*1.1)) for c in fill_col))
            d.rounded_rectangle([x0+3,y0+3,x1+3,y1+3], radius=radius, fill=darker(fill_col,0.85))
            d.rounded_rectangle([x0,y0,x1,y1], radius=radius, fill=fill_col, outline=border_col, width=max(1,int(2*scale)))
        else:
            d.rounded_rectangle([x0+3,y0+3,x1+3,y1+3], radius=radius, fill=(225,230,240))
            d.rounded_rectangle([x0,y0,x1,y1], radius=radius, fill=BOX_BG, outline=BOX_BORDER, width=max(1,int(2*scale)))
            txt_main, txt_sub = TEXT_COLOR, SUBTEXT
            if pos_col and color_style == "stripe":
                sw = int(16*scale)
                d.rounded_rectangle([x0,y0,x0+sw,y1], radius=radius, fill=pos_col, outline=pos_col)

        if highlight:
            d.rounded_rectangle([x0-2,y0-2,x1+2,y1+2], radius=radius, outline=HILIGHT, width=max(2,int(3*scale)))

        pad = int(12*scale)
        name = person.full_name or person.username
        line2 = f"{person.title}" + (f"  |  {person.dept}" if (show_dept and person.dept) else "")
        line3 = person.mail if (show_mail and person.mail) else ""

        max_name = 30 if scale>=1.0 else 24
        max_sub  = 46 if scale>=1.0 else 36
        name_wrapped   = "\n".join(wrap(name, max_name)) if len(name) > max_name else name
        line2_wrapped  = "\n".join(wrap(line2, max_sub)) if len(line2) > max_sub else line2

        d.text((x0+pad, y0+int(10*scale)),  name_wrapped,  fill=txt_main, font=font_title)
        d.text((x0+pad, y0+int(36*scale)), line2_wrapped, fill=txt_sub,  font=font_sub)
        if line3: d.text((x0+pad, y0+int(58*scale)), line3, fill=txt_sub, font=font_sub)

    xoff = 0
    for layout in layouts:
        centers = {}
        for u,(cx,cy) in layout.items():
            x0 = PX + (xoff+cx)*(BW+CG); y0 = PX + cy*(BH+RG)
            centers[u] = (x0 + BW//2, y0 + BH//2)
        # bağlantılar
        for u,(cx,cy) in layout.items():
            for v in children.get(u, []):
                if v not in layout: continue
                x1,y1 = centers[u]; xv,yv = centers[v]
                top_child_y = PX + layout[v][1]*(BH+RG)
                bottom_parent_y = PX + layout[u][1]*(BH+RG) + BH
                mid_y = (bottom_parent_y + top_child_y)//2
                d.line([(x1,bottom_parent_y),(x1,mid_y)], fill=LINE_COLOR, width=max(1,int(2*scale)))
                d.line([(x1,mid_y),(xv,mid_y)],         fill=LINE_COLOR, width=max(1,int(2*scale)))
                d.line([(xv,mid_y),(xv,top_child_y)],   fill=LINE_COLOR, width=max(1,int(2*scale)))
        # kutular
        for u,(cx,cy) in layout.items():
            draw_box(xoff+cx, cy, people[u], highlight=(highlight_user and u==highlight_user))
        used = max([x for x,_ in layout.values()]+[0]) + 1
        xoff += used

    # Legend (tekilleştirilmiş kullan)
    if legend_w and legend_items:
        dfont = ensure_font(max(11,int(15*scale)))
        sfont = ensure_font(max(10,int(12*scale)))
        items = legend_items
        lx0 = img_w - legend_w + int(12*scale)
        ly0 = int(16*scale); lx1 = img_w - int(12*scale)
        ly1 = ly0 + int((len(items)*24 + 56)*scale)
        d.rounded_rectangle([lx0,ly0,lx1,ly1], radius=int(12*scale), fill=(255,255,255), outline=(220,225,235))
        d.text((lx0+int(12*scale), ly0+int(12*scale)), "Pozisyon Renkleri", fill=(40,40,40), font=dfont)
        y = ly0 + int(40*scale)
        for title, col in items:
            sz = int(18*scale)
            d.rounded_rectangle([lx0+int(12*scale), y, lx0+int(12*scale)+sz, y+sz], radius=int(6*scale), fill=col, outline=col)
            d.text((lx0+int(12*scale)+sz+int(10*scale), y-2), title, fill=(60,60,60), font=sfont)
            y += int(24*scale)

    return img, bboxes

def save_png(img: Image.Image, path: str): img.save(path, "PNG", optimize=True)
def save_pdf(img: Image.Image, path: str): img.convert("RGB").save(path, "PDF", resolution=300.0)

# -------------- UI --------------
class OrgChartApp((TkinterDnD.Tk if DnDEnabled else tk.Tk)):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        try:
            if sys.platform.startswith("win"):
                from ctypes import windll
                windll.shcore.SetProcessDpiAwareness(1)
            self.tk.call('tk','scaling', 1.4)
        except: pass
        self.geometry("1360x820"); self.minsize(1180,680)

        # ttk style
        default_font = ("Segoe UI", 12)
        small_font   = ("Segoe UI", 11)
        self.option_add("*Font", default_font)
        style = ttk.Style(self)
        try: style.theme_use("clam")
        except: pass
        style.configure("TButton", padding=(14,8), font=default_font)
        style.configure("TLabel",  font=default_font)
        style.configure("TEntry",  padding=6, font=default_font)
        style.configure("TCombobox", padding=6, font=default_font)
        style.configure("Horizontal.TScale", sliderlength=28, troughrelief="flat")
        style.configure("Treeview.Heading", font=("Segoe UI", 12, "bold"))
        style.configure("Treeview", rowheight=28, font=small_font)

        # durum
        self.df: Optional[pd.DataFrame] = None
        self.people: Dict[str,Person] = {}
        self.children: Dict[str,List[str]] = {}
        self.roots: List[str] = []
        self.render_img: Optional[Image.Image] = None
        self.preview_img: Optional[Image.Image] = None
        self.render_tk = None
        self.scale = 1.0
        self.highlight_user: Optional[str] = None
        self.show_dept, self.show_mail = True, True
        self.color_style = "bg"
        self.var_fast = tk.BooleanVar(value=True)
        self.node_bboxes: Dict[str, Tuple[int,int,int,int]] = {}  # render ölçeğinde

        # ayarlar (renkler)
        self.settings_path = DEFAULT_SETTINGS_FILE
        self.settings = self._load_settings()
        self.title_colors = {k:tuple(v) for k,v in self.settings.get("title_colors", {}).items()}

        self._build_ui()

        # DnD bağlama
        if DnDEnabled:
            for w in (self, self.canvas, self.right_panel):
                try:
                    w.drop_target_register(DND_FILES)
                    w.dnd_bind("<<Drop>>", self._on_drop)
                except Exception:
                    pass
            self.status.configure(text="Sürükle-bırak aktif: Excel/CSV dosyasını pencereye bırakın.")
        else:
            self.status.configure(text="Sürükle-bırak pasif (tkinterdnd2 kurulu değil). Dosya seçerek yükleyin.")

    # ---------- UI ----------
    def _build_ui(self):
        top = ttk.Frame(self); top.pack(side=tk.TOP, fill=tk.X, padx=12, pady=8)

        ttk.Button(top, text="Excel/CSV Yükle", command=self.on_load).pack(side=tk.LEFT, padx=(0,8))

        ttk.Label(top, text="Departman:").pack(side=tk.LEFT, padx=(10,4))
        self.cmb_dept = ttk.Combobox(top, width=26, state="disabled"); self.cmb_dept.pack(side=tk.LEFT)
        self.cmb_dept.bind("<<ComboboxSelected>>", lambda e: self.on_relayout())

        ttk.Label(top, text="Başlangıç (Yönetici):").pack(side=tk.LEFT, padx=(14,4))
        self.cmb_root = ttk.Combobox(top, width=30, state="disabled"); self.cmb_root.pack(side=tk.LEFT)

        self.btn_build = ttk.Button(top, text="Şemayı Oluştur", command=self.on_relayout, state="disabled")
        self.btn_build.pack(side=tk.LEFT, padx=10)

        ttk.Label(top, text="Yakınlaştır:").pack(side=tk.LEFT, padx=(14,6))
        self.zoom = ttk.Scale(top, from_=50, to=200, orient="horizontal", command=self.on_zoom, length=180)
        self.zoom.set(100); self.zoom.configure(state="disabled"); self.zoom.pack(side=tk.LEFT, padx=(0,10))

        ttk.Checkbutton(top, text="Hızlı ön-izleme", variable=self.var_fast, command=self.on_zoom).pack(side=tk.LEFT, padx=(0,12))

        ttk.Label(top, text="Ara:").pack(side=tk.LEFT, padx=(8,4))
        #self.ent_search = ttk.Entry(top, width=20); self.ent_search.pack(side=tk.LEFT)
        #ttk.Button(top, text="Bul", command=self.on_search).pack(side=tk.LEFT, padx=(6,4))
        ttk.Button(top, text="Vurguyu Temizle", command=self.clear_highlight).pack(side=tk.LEFT, padx=(4,10))

        self.btn_export_png = ttk.Button(top, text="PNG Dışa Aktar", command=lambda: self.on_export("png"), state="disabled")
        self.btn_export_pdf = ttk.Button(top, text="PDF Dışa Aktar", command=lambda: self.on_export("pdf"), state="disabled")
        self.btn_export_png.pack(side=tk.RIGHT, padx=(6,0)); self.btn_export_pdf.pack(side=tk.RIGHT, padx=(6,0))

        # gövde
        body = ttk.Frame(self); body.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=12, pady=(0,10))
        left = ttk.Frame(body); left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(left, bg="#f6f8fa", highlightthickness=0)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.vbar = ttk.Scrollbar(left, orient=tk.VERTICAL, command=self.canvas.yview); self.vbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.hbar = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.canvas.xview); self.hbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.configure(yscrollcommand=self.vbar.set, xscrollcommand=self.hbar.set)

        self.right_panel = ttk.Frame(body, width=340); self.right_panel.pack(side=tk.RIGHT, fill=tk.Y); self.right_panel.pack_propagate(False)
        ttk.Label(self.right_panel, text="Pozisyon Renkleri", font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=6, pady=(4,8))
        self.tree_colors = ttk.Treeview(self.right_panel, columns=("poz","renk"), show="headings", height=16)
        self.tree_colors.heading("poz", text="Pozisyon")
        self.tree_colors.heading("renk", text="Renk")
        self.tree_colors.column("poz", width=210, anchor="w")
        self.tree_colors.column("renk", width=100, anchor="center")
        self.tree_colors.pack(fill=tk.BOTH, padx=6, pady=(0,8), expand=False)
        ttk.Label(self.right_panel, text="İpucu: Renk hücresine çift tıklayın.", foreground="#666").pack(anchor="w", padx=8, pady=(0,10))

        b = ttk.Frame(self.right_panel); b.pack(fill=tk.X, padx=6, pady=(4,10))
        ttk.Button(b, text="Renkleri Kaydet", command=self.save_color_settings).pack(side=tk.LEFT, padx=(0,6))
        ttk.Button(b, text="Renkleri Yükle", command=self.load_color_settings).pack(side=tk.LEFT, padx=6)
        ttk.Button(b, text="Temizle", command=self.clear_colors).pack(side=tk.LEFT, padx=6)

        ttk.Separator(self.right_panel, orient="horizontal").pack(fill=tk.X, padx=6, pady=8)
        self.var_show_dept = tk.BooleanVar(value=True)
        self.var_show_mail = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.right_panel, text="Departman satırını göster", variable=self.var_show_dept, command=self.on_redraw).pack(anchor="w", padx=8)
        ttk.Checkbutton(self.right_panel, text="Mail satırını göster",       variable=self.var_show_mail, command=self.on_redraw).pack(anchor="w", padx=8)

        ttk.Label(self.right_panel, text="Renk uygulama stili").pack(anchor="w", padx=8, pady=(10,2))
        self.cmb_style = ttk.Combobox(self.right_panel, values=["Arka plan","Sol şerit"], state="readonly", width=16)
        self.cmb_style.set("Arka plan"); self.cmb_style.pack(anchor="w", padx=8)
        self.cmb_style.bind("<<ComboboxSelected>>", lambda e: self.on_redraw())

        self.status = ttk.Label(self, text="Hazır", anchor="w")
        self.status.pack(side=tk.BOTTOM, fill=tk.X, padx=12, pady=(0,10))

        self.tree_colors.bind("<Double-1>", self.on_color_pick)

    # ---------- DnD ----------
    @staticmethod
    def _parse_dnd_path(event_data: str) -> Optional[str]:
        if not event_data: return None
        if event_data.startswith("{") and event_data.endswith("}"):
            # tek dosya süslü
            return event_data.strip("{}")
        # birden çok bırakma: {a} {b}
        parts=[]
        cur=""; inb=False
        for ch in event_data:
            if ch=="{": inb=True; cur=""; continue
            if ch=="}": inb=False; parts.append(cur); cur=""; continue
            if inb: cur+=ch
        if parts: return parts[0]
        return event_data.strip()

    def _on_drop(self, event):
        path = self._parse_dnd_path(event.data)
        if not path: return
        if not os.path.isfile(path):
            messagebox.showwarning("Uyarı", "Bıraktığınız öğe bir dosya değil."); return
        if not path.lower().endswith((".xlsx",".xls",".csv")):
            messagebox.showwarning("Uyarı", "Lütfen Excel/CSV dosyası bırakın."); return
        self._load_path(path)

    # ---------- Olaylar ----------
    def on_load(self):
        path = filedialog.askopenfilename(title="Excel/CSV seçin", filetypes=[("Excel/CSV","*.xlsx;*.xls;*.csv"),("Tüm Dosyalar","*.*")])
        if not path: return
        self._load_path(path)

    def _load_path(self, path: str):
        try:
            self.df = load_table(path)
        except Exception as e:
            messagebox.showerror("Hata", f"Tablo okunamadı:\n{e}"); return

        depts = ["Tümü"] + sorted([x for x in self.df["Departman"].dropna().unique()])
        self.cmb_dept.configure(values=depts, state="readonly"); self.cmb_dept.set("Tümü")

        users = ["Otomatik (Kökler)"] + [f'{u} — {n}' for u,n in zip(self.df["Kullanıcı Adı"], self.df["Ad Soyad"])]
        self.cmb_root.configure(values=users, state="readonly"); self.cmb_root.set("Otomatik (Kökler)")

        self.btn_build.configure(state="normal")
        self.cmb_dept.configure(state="readonly"); self.cmb_root.configure(state="readonly")
        self.zoom.configure(state="normal")

        self.refresh_position_list()
        self.on_relayout()
        self.status.configure(text=f"Yüklendi: {path} | {len(self.df)} satır")

    def on_relayout(self):
        if self.df is None: return
        dept = self.cmb_dept.get(); dept = None if (not dept or dept=="Tümü") else dept
        self.people = build_people(self.df, dept)
        if not self.people:
            messagebox.showwarning("Uyarı","Filtreye uyan kayıt bulunamadı."); return
        self.children, self.roots = build_tree(self.people)

        start_from=None
        root_sel = self.cmb_root.get()
        if root_sel and root_sel!="Otomatik (Kökler)":
            sf = root_sel.split(" — ")[0].strip()
            if sf in self.people: start_from = sf

        self.show_dept = self.var_show_dept.get()
        self.show_mail = self.var_show_mail.get()
        self.color_style = "bg" if self.cmb_style.get()=="Arka plan" else "stripe"

        self.render_img, self.node_bboxes = draw_chart(
            self.people, self.children, self.roots, start_from,
            scale=self.scale, title_colors=self.title_colors,
            show_dept=self.show_dept, show_mail=self.show_mail,
            highlight_user=self.highlight_user, color_style=self.color_style,
            add_legend=True
        )
        self.preview_img = self.render_img
        self._paint_preview()

        self.btn_export_png.configure(state="normal")
        self.btn_export_pdf.configure(state="normal")

        info = f"Şema: {len(self.people)} | Kök: {len(self.roots)} | Ölçek: {int(self.scale*100)}%"
        if start_from: info += f" | Başlangıç: {start_from}"
        if dept: info += f" | Departman: {dept}"
        self.status.configure(text=info)

    def on_redraw(self):
        if self.df is None or not self.people: return
        self.on_relayout()

    def on_zoom(self, _=None):
        if self.df is None: return
        v = max(50, min(200, int(float(self.zoom.get()))))
        new_scale = v/100.0
        if self.var_fast.get() and self.render_img is not None:
            self.scale = new_scale
            new_w = max(1,int(self.render_img.width * self.scale))
            new_h = max(1,int(self.render_img.height * self.scale))
            self.preview_img = self.render_img.resize((new_w,new_h), Image.BILINEAR)
            self._paint_preview()
        else:
            self.scale = new_scale
            self.on_relayout()

    def _paint_preview(self):
        if self.preview_img is None: return
        tkimg = self._pil_to_tk(self.preview_img); self.render_tk = tkimg
        self.canvas.delete("all")
        self.canvas.create_image(0,0, anchor="nw", image=self.render_tk)
        self.canvas.config(scrollregion=(0,0,self.preview_img.width, self.preview_img.height))

    def _center_on_user(self, username: str):
        """Bulunduğunda kişiyi ekrana getir (merkezle)."""
        if not self.preview_img or username not in self.node_bboxes:
            return
        # bbox render ölçeğinde; önizleme ölçeğine çevir
        ratio_x = self.preview_img.width  / self.render_img.width
        ratio_y = self.preview_img.height / self.render_img.height
        x0,y0,x1,y1 = self.node_bboxes[username]
        cx = ((x0+x1)//2) * ratio_x
        cy = ((y0+y1)//2) * ratio_y

        view_w = max(1, self.canvas.winfo_width())
        view_h = max(1, self.canvas.winfo_height())
        # hedef sol-üst
        target_x = max(0, cx - view_w/2)
        target_y = max(0, cy - view_h/2)

        frac_x = target_x / max(1, self.preview_img.width  - view_w)
        frac_y = target_y / max(1, self.preview_img.height - view_h)
        self.canvas.xview_moveto(min(1.0, max(0.0, frac_x)))
        self.canvas.yview_moveto(min(1.0, max(0.0, frac_y)))

    def on_search(self):
        q = (self.ent_search.get() or "").strip().lower()
        if not q or not self.people: return
        found = None
        for u,p in self.people.items():
            if q in u.lower() or q in (p.full_name or "").lower():
                found = u; break
        if not found:
            messagebox.showinfo("Bilgi","Eşleşme bulunamadı."); return
        self.highlight_user = found
        self.on_relayout()
        self.after(50, lambda: self._center_on_user(found))  # görüntü çizildikten hemen sonra merkeze al

    def clear_highlight(self):
        self.highlight_user=None; self.on_relayout()

    def on_export(self, fmt):
        if self.render_img is None: return
        if fmt=="png":
            path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG","*.png")], title="PNG olarak kaydet")
            if not path: return
            try: save_png(self.render_img, path); messagebox.showinfo("Kaydedildi", f"PNG: {path}")
            except Exception as e: messagebox.showerror("Hata", f"Kaydedilemedi:\n{e}")
        else:
            path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF","*.pdf")], title="PDF olarak kaydet")
            if not path: return
            try: save_pdf(self.render_img, path); messagebox.showinfo("Kaydedildi", f"PDF: {path}")
            except Exception as e: messagebox.showerror("Hata", f"Kaydedilemedi:\n{e}")

    # ---- Renk paneli ----
    def refresh_position_list(self):
        self.tree_colors.delete(*self.tree_colors.get_children())
        if self.df is None: return
        titles = sorted({(t or "").strip() for t in self.df["Pozisyon"].dropna().unique()})
        for t in titles:
            hexv = self._rgb_to_hex(self.title_colors.get(t)) if self.title_colors.get(t) else ""
            self.tree_colors.insert("", "end", values=(t, hexv))

    def on_color_pick(self, _):
        sel = self.tree_colors.selection()
        if not sel: return
        t = self.tree_colors.item(sel[0], "values")[0]
        initial = self.title_colors.get(t, (39,94,254))
        col = colorchooser.askcolor(color=initial, title=f"Renk seç — {t}")
        if not col or not col[0]: return
        r,g,b = [int(x) for x in col[0]]
        self.title_colors[t] = (r,g,b)
        self.tree_colors.set(sel[0], column="renk", value=self._rgb_to_hex((r,g,b)))
        self.on_redraw()

    def save_color_settings(self):
        self._save_settings({"title_colors": {k:list(v) for k,v in self.title_colors.items()}})
        messagebox.showinfo("Bilgi", f"Renkler kaydedildi:\n{self.settings_path}")

    def load_color_settings(self):
        self.settings = self._load_settings()
        self.title_colors = {k:tuple(v) for k,v in self.settings.get("title_colors", {}).items()}
        self.refresh_position_list(); self.on_redraw()
        messagebox.showinfo("Bilgi","Kaydedilmiş renkler yüklendi.")

    def clear_colors(self):
        self.title_colors = {}; self.refresh_position_list(); self.on_redraw()

    # ---- Toolkit ----
    def _load_settings(self):
        if os.path.exists(self.settings_path):
            try:
                with open(self.settings_path,"r",encoding="utf-8") as f: return json.load(f)
            except: return {}
        return {}
    def _save_settings(self, data: dict):
        try:
            with open(self.settings_path,"w",encoding="utf-8") as f: json.dump(data,f,ensure_ascii=False,indent=2)
        except Exception as e:
            messagebox.showwarning("Uyarı", f"Ayarlar kaydedilemedi:\n{e}")

    @staticmethod
    def _pil_to_tk(img: Image.Image):
        b = BytesIO(); img.save(b, "PNG"); return tk.PhotoImage(data=b.getvalue())

    @staticmethod
    def _rgb_to_hex(rgb: Optional[Tuple[int,int,int]]) -> str:
        if not rgb: return ""
        r,g,b = rgb; return f"#{r:02x}{g:02x}{b:02x}"

# -------------- Çalıştır --------------
def main():
    app = OrgChartApp()
    app.mainloop()

if __name__ == "__main__":
    main()
