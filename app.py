import base64
import json
import os
import re
import sys
import tempfile
import threading
import time
import webbrowser
import tkinter as tk
from datetime import datetime
from io import BytesIO
from tkinter import messagebox, ttk

import barcode
import qrcode
import requests
import win32print
import win32ui
from barcode.writer import ImageWriter
from PIL import Image, ImageDraw, ImageFont, ImageTk, ImageWin

# Selenium & Webdriver Manager
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

try:
    from webdriver_manager.chrome import ChromeDriverManager
    HAS_WD_MANAGER = True
except ImportError:
    HAS_WD_MANAGER = False


# ============================================================
# 0. CẤU HÌNH Đường DẪN FILE HỆ THỐNG
# ============================================================

def app_data_path(filename):
    if getattr(sys, "frozen", False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)


# ============================================================
# 1. CẤU HÌNH & DICTIONARY WMS
# ============================================================

URL_WMS = "https://wms.ssc.shopee.vn/v2/returninbound/arrving"
BASE_URL = "https://wms.ssc.shopee.vn"
ENDPOINT_SEARCH_ASN = "/api/apps/process/returninbound/riasn/search_asn"
ENDPOINT_SEARCH_ASN_SKU = "/api/apps/process/returninbound/riasn/search_asn_sku"
ENDPOINT_SEARCH_ASN_DETAIL = "/api/apps/process/returninbound/riasn/search_asn_detail"

DEFAULT_USER = "797060"
DEFAULT_PASS = "Saigon@123"
DEFAULT_STATION = "1"

LABEL_WIDTH_IN = 4.00
LABEL_HEIGHT_IN = 2.00

HARDCODED_X_CSRFTOKEN = "JGcx4YRxWPtLsgzz7fa25i1Vc7J2HBeP"

INBOUND_TYPE_MAP = {
    1: "Failed Delivery", "1": "Failed Delivery",
    2: "Buyer Return", "2": "Buyer Return",
}

FOLLOW_UP_MAP = {
    1: "Return to Buyer", "1": "Return to Buyer",
    2: "Return to Seller", "2": "Return to Seller",
    3: "Return to Seller [Non-SBS]", "3": "Return to Seller [Non-SBS]",
    4: "Receive", "4": "Receive",
    5: "Liquidate/Dispose", "5": "Liquidate/Dispose",
    6: "Pending Seller Feedback", "6": "Pending Seller Feedback",
    7: "Shortage Return", "7": "Shortage Return",
    8: "Liquidate", "8": "Liquidate",
    9: "Dispose", "9": "Dispose",
    10: "Return to Seller [JIT]", "10": "Return to Seller [JIT]",
}


def normalize_cb_option(value):
    if value is None or isinstance(value, bool):
        return "-"
    if isinstance(value, (int, float)):
        return "CB" if value == 1 else ("Local" if value == 0 else "-")
    value_str = str(value).strip().lower()
    if not value_str or value_str in {"null", "none", "-", "nan"}:
        return "-"
    if value_str in {"cb", "oversea", "overseas", "1", "1.0"}:
        return "CB"
    if value_str in {"local", "0", "0.0"}:
        return "Local"
    return "-"


def format_currency(amount):
    if amount is None or amount == "-":
        return "-"
    try:
        val = float(amount)
        if val <= 0:
            return "-"
        return f"{val:,.0f} đ".replace(",", ".")
    except Exception:
        return "-"


def _decode_json_string(value):
    if not isinstance(value, str):
        return None
    text = value.strip()
    if not ((text.startswith("{") and text.endswith("}")) or (text.startswith("[") and text.endswith("]"))):
        return None
    try:
        return json.loads(text)
    except Exception:
        return None


# ============================================================
# 2. TẠO TEM & IN MÁY IN
# ============================================================

def generate_label_image(print_value, sub_info="", followup_info="", inbound_type="", cb_option="-", total_price="-"):
    dpi = 300
    width_px = int(LABEL_WIDTH_IN * dpi)
    height_px = int(LABEL_HEIGHT_IN * dpi)

    img = Image.new("RGB", (width_px, height_px), "white")
    draw = ImageDraw.Draw(img)

    LEFT, RIGHT = 30, 30
    usable_width = width_px - LEFT - RIGHT

    def load_font(size, bold=False):
        candidates = ["arialbd.ttf", "segoeuib.ttf"] if bold else ["arial.ttf", "segoeui.ttf"]
        for path in candidates:
            try:
                return ImageFont.truetype(path, int(size))
            except Exception:
                continue
        return ImageFont.load_default()

    def text_size(text, font):
        box = draw.textbbox((0, 0), str(text), font=font)
        return (box[2] - box[0], box[3] - box[1])

    def fit_font(text, max_width, max_size, min_size=12, bold=True):
        text = str(text or "").strip()
        if not text:
            return load_font(max_size, bold)
        size = max_size
        while size > min_size:
            font = load_font(size, bold)
            w, _ = text_size(text, font)
            if w <= max_width:
                return font
            size -= 1
        return load_font(min_size, bold)

    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    font_time = load_font(20, False)
    draw.text((LEFT, 15), f"Thời gian: {current_time}", fill="black", font=font_time)

    if total_price and total_price != "-":
        price_text = f"Giá: {total_price}"
        font_price = load_font(26, True)
        pw, _ = text_size(price_text, font_price)
        draw.text((width_px - RIGHT - pw, 12), price_text, fill="black", font=font_price)

    meta_info_parts = []
    if sub_info: meta_info_parts.append(sub_info)
    if inbound_type and inbound_type != "-": meta_info_parts.append(f"Loại: {inbound_type}")
    if cb_option and cb_option != "-": meta_info_parts.append(f"[{cb_option}]")

    combined_meta = "  |  ".join(meta_info_parts)
    if combined_meta:
        font_meta = fit_font(combined_meta, usable_width, max_size=28, min_size=16, bold=True)
        draw.text((LEFT, 45), combined_meta, fill="black", font=font_meta)

    draw.line((LEFT, 85, width_px - RIGHT, 85), fill="black", width=3)

    y_center = 95
    primary_code = str(print_value or "").split(",")[0].strip()
    primary_code_clean = re.sub(r"[^A-Za-z0-9\-_]", "", primary_code) or "UNKNOWN"

    qr_size = 210
    barcode_width = usable_width - qr_size - 40
    barcode_height = 140

    try:
        code128 = barcode.get_barcode_class("code128")
        bc_instance = code128(primary_code_clean, writer=ImageWriter())
        bc_path = os.path.join(tempfile.gettempdir(), f"bc_{primary_code_clean}")
        bc_file = bc_instance.save(bc_path, options={"write_text": False, "module_height": 18, "quiet_zone": 1})
        with Image.open(bc_file) as raw:
            bc_img = raw.convert("RGB").resize((barcode_width, barcode_height), Image.Resampling.LANCZOS)
        img.paste(bc_img, (LEFT, y_center))
    except Exception:
        pass

    try:
        qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=8, border=1)
        qr.add_data(str(print_value))
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB").resize((qr_size, qr_size), Image.Resampling.NEAREST)
        img.paste(qr_img, (width_px - RIGHT - qr_size, y_center))
    except Exception:
        pass

    y_code_text = y_center + barcode_height + 5
    code_text = str(print_value or "").strip()
    font_code = fit_font(code_text, barcode_width, max_size=36, min_size=20, bold=True)
    cw, _ = text_size(code_text, font_code)
    draw.text((LEFT + (barcode_width - cw) // 2, y_code_text), code_text, fill="black", font=font_code)

    y_followup = 355
    draw.line((LEFT, y_followup, width_px - RIGHT, y_followup), fill="black", width=2)

    follow_text = str(followup_info or "RECEIVE").strip()
    font_follow = fit_font(follow_text, usable_width - 20, max_size=110, min_size=40, bold=True)
    fw, fh = text_size(follow_text, font_follow)

    draw.text(((width_px - fw) // 2, y_followup + (230 - fh) // 2 - 10), follow_text, fill="black", font=font_follow)
    draw.rectangle((4, 4, width_px - 5, height_px - 5), outline="black", width=3)

    output_path = os.path.join(tempfile.gettempdir(), "label_4x2_to_print.png")
    img.save(output_path, dpi=(dpi, dpi))
    return output_path


def print_label_direct(image_path):
    try:
        printer_name = win32print.GetDefaultPrinter()
        if not printer_name:
            messagebox.showerror("Lỗi Máy In", "Chưa chọn Máy in mặc định!")
            return
        img = Image.open(image_path)
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(printer_name)
        printable_area = (hDC.GetDeviceCaps(8), hDC.GetDeviceCaps(10))
        hDC.StartDoc("WMS_Print_Label")
        hDC.StartPage()
        dib = ImageWin.Dib(img)
        dib.draw(hDC.GetHandleOutput(), (0, 0, printable_area[0], printable_area[1]))
        hDC.EndPage()
        hDC.EndDoc()
        hDC.DeleteDC()
    except Exception as e:
        messagebox.showerror("Lỗi In Tem", f"Không thể gửi lệnh in tới máy in!\n\nChi tiết: {e}")


# ============================================================
# 3. TRÍCH XUẤT NÂNG CAO (SKU NAME, IMAGE, QC IMAGES, TOT. PRICE)
# ============================================================

def extract_wms_details_deep(data_obj):
    skus, follow_ups, cb_options, inbound_types = set(), set(), set(), set()
    sku_names, sku_images, qc_images = set(), set(), set()
    total_price_sum = 0.0

    sku_keys = {"sku_id", "item_sku_id", "sku", "seller_sku", "sku_code", "msku_id", "product_sku_id"}
    followup_keys = {"follow_up", "follow_up_type", "followup", "followup_type", "follow_up_status", "action_type"}
    qty_keys = {"qty", "quantity", "amount", "total_qty", "item_qty", "sku_qty", "expected_qty"}
    price_keys = {"total_price", "price", "item_price", "order_price", "unit_price", "goods_price", "original_price"}

    def parse_number(val):
        if val is None: return 0.0
        try:
            clean_str = re.sub(r"[^\d.]", "", str(val).strip())
            return float(clean_str) if clean_str else 0.0
        except Exception:
            return 0.0

    def recurse(obj):
        nonlocal total_price_sum
        if isinstance(obj, dict):
            dict_price = 0.0
            dict_qty = 1.0

            for k, v in obj.items():
                k_lower = str(k).lower().strip()
                
                if k_lower in sku_keys and v and str(v).strip() != "0":
                    skus.add(str(v).strip())

                if k_lower in {"sku_name", "item_name", "product_name"} and v:
                    val_s = str(v).strip()
                    if val_s: sku_names.add(val_s)

                if k_lower in {"image", "image_url", "sku_image", "product_image"} and v:
                    val_s = str(v).strip()
                    if val_s.startswith("http"): sku_images.add(val_s)

                if k_lower in {"qc_image_list", "qc_images", "qc_media_info"} and v:
                    if isinstance(v, list):
                        for img_item in v:
                            if isinstance(img_item, str) and img_item.startswith("http"):
                                qc_images.add(img_item)
                            elif isinstance(img_item, dict) and img_item.get("url"):
                                url_str = str(img_item.get("url")).strip()
                                if url_str.startswith("http"): qc_images.add(url_str)

                if k_lower in followup_keys and v:
                    val_str = str(v).strip()
                    follow_ups.add(FOLLOW_UP_MAP.get(v) or FOLLOW_UP_MAP.get(val_str) or val_str)
                
                if k_lower in {"cb_option", "cb_flag", "is_cb", "cb"}:
                    cb_options.add(normalize_cb_option(v))
                
                if k_lower in {"inbound_type", "inboundtype", "return_inbound_type"} and v:
                    val_str = str(v).strip()
                    inbound_types.add(INBOUND_TYPE_MAP.get(v) or INBOUND_TYPE_MAP.get(val_str) or val_str)

                if k_lower in price_keys and v is not None:
                    p = parse_number(v)
                    if p > dict_price: dict_price = p

                if k_lower in qty_keys and v is not None:
                    q = parse_number(v)
                    if q > 0: dict_qty = q

                if isinstance(v, (dict, list)):
                    recurse(v)
                elif isinstance(v, str):
                    nested = _decode_json_string(v)
                    if nested:
                        recurse(nested)

            if dict_price > 0:
                total_price_sum += (dict_price * dict_qty)

        elif isinstance(obj, list):
            for element in obj:
                recurse(element)

    recurse(data_obj)

    return (
        ", ".join(skus) if skus else "-",
        ", ".join(follow_ups) if follow_ups else "-",
        ", ".join(inbound_types) if inbound_types else "-",
        "CB" if "CB" in cb_options else ("Local" if "Local" in cb_options else "-"),
        total_price_sum,
        " | ".join(sku_names) if sku_names else "-",
        list(sku_images),
        list(qc_images),
    )


# ============================================================
# 4. BUTTON WIDGET
# ============================================================

class AnimatedButton(tk.Button):
    def __init__(self, parent, text, command, bg_color, hover_color, fg_color="#FFFFFF", font=("Segoe UI Semibold", 10), pad_x=16, pad_y=8, **kwargs):
        super().__init__(
            parent, text=text, command=command, font=font,
            bg=bg_color, fg=fg_color, activebackground=hover_color, activeforeground=fg_color,
            bd=0, relief="flat", cursor="hand2", padx=pad_x, pady=pad_y, **kwargs
        )
        self.bg_color = bg_color
        self.hover_color = hover_color
        self.bind("<Enter>", lambda e: self.config(bg=self.hover_color))
        self.bind("<Leave>", lambda e: self.config(bg=self.bg_color))


# ============================================================
# 5. APPLICATION CHÍNH
# ============================================================

class ScanReturnTNApp:

    def __init__(self, root):
        self.root = root

        self.root.title("WMS SSC - System Check & Auto Print")
        
        # MỞ RỘNG GIAO DIỆN KHỞI TẠO MẶC ĐỊNH & PHÓNG TO TOÀN MÀN HÌNH
        self.root.geometry("1680x950")
        self.root.minsize(1280, 750)
        try:
            self.root.state("zoomed")  # Phóng to toàn màn hình trên Windows
        except Exception:
            pass
            
        self.root.configure(bg="#0F172A")

        self.session = requests.Session()
        self.cookie_val = ""
        self.csrf_token = ""
        self.all_records = []
        self.photo_cache = []

        self.setup_theme_styles()
        self.create_layout()

    def setup_theme_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.colors = {
            "bg": "#0F172A",
            "card_bg": "#1E293B",
            "card_border": "#334155",
            "text": "#F8FAFC",
            "muted": "#94A3B8",
            "primary": "#0068FF",
            "primary_hover": "#0052CC",
            "success": "#22C55E",
            "success_hover": "#16A34A",
            "warning": "#F59E0B",
            "danger": "#EF4444",
            "danger_hover": "#DC2626",
            "accent": "#6366F1",
            "accent_hover": "#4F46E5",
        }

        self.style.configure(
            "Treeview",
            background="#1E293B",
            foreground="#F8FAFC",
            fieldbackground="#1E293B",
            rowheight=40,
            borderwidth=0,
            font=("Segoe UI", 10)
        )

        self.style.configure(
            "Treeview.Heading",
            background="#334155",
            foreground="#FFFFFF",
            borderwidth=1,
            relief="flat",
            font=("Segoe UI Bold", 10),
            padding=(6, 12)
        )

        self.style.map("Treeview", background=[("selected", "#0068FF")], foreground=[("selected", "#FFFFFF")])

    def _create_stat_card(self, parent, title, value, icon, accent_color):
        card = tk.Frame(parent, bg=self.colors["card_bg"], highlightbackground=self.colors["card_border"], highlightthickness=1)
        
        icon_box = tk.Frame(card, bg=accent_color, width=44, height=44)
        icon_box.pack(side="left", padx=14, pady=12)
        icon_box.pack_propagate(False)
        tk.Label(icon_box, text=icon, bg=accent_color, fg="#FFFFFF", font=("Segoe UI", 14, "bold")).pack(expand=True)

        info = tk.Frame(card, bg=self.colors["card_bg"])
        info.pack(side="left", fill="both", expand=True, pady=10)
        
        tk.Label(info, text=title, bg=self.colors["card_bg"], fg=self.colors["muted"], font=("Segoe UI", 8, "bold")).pack(anchor="w")
        val_lbl = tk.Label(info, text=value, bg=self.colors["card_bg"], fg=self.colors["text"], font=("Segoe UI", 16, "bold"))
        val_lbl.pack(anchor="w")

        return card, val_lbl

    def create_layout(self):
        # 1. HEADER
        header = tk.Frame(self.root, bg=self.colors["card_bg"], height=65)
        header.pack(fill="x")
        header.pack_propagate(False)

        brand = tk.Frame(header, bg=self.colors["card_bg"])
        brand.pack(side="left", padx=24, pady=10)

        logo = tk.Frame(brand, bg=self.colors["primary"], width=38, height=38)
        logo.pack(side="left", padx=(0, 12))
        logo.pack_propagate(False)
        tk.Label(logo, text="Z", bg=self.colors["primary"], fg="#FFFFFF", font=("Segoe UI", 16, "bold")).pack(expand=True)

        title_box = tk.Frame(brand, bg=self.colors["card_bg"])
        title_box.pack(side="left")
        tk.Label(title_box, text="WMS SYSTEM CHECK", bg=self.colors["card_bg"], fg=self.colors["text"], font=("Segoe UI", 14, "bold")).pack(anchor="w")
        tk.Label(title_box, text="AUTO PRINT & DATA MONITORING", bg=self.colors["card_bg"], fg=self.colors["muted"], font=("Segoe UI", 8, "bold")).pack(anchor="w")

        # 2. KHU VỰC ĐĂNG NHẬP WMS
        login_card = tk.Frame(self.root, bg=self.colors["card_bg"], highlightbackground=self.colors["card_border"], highlightthickness=1)
        login_card.pack(fill="x", padx=24, pady=(10, 4))

        login_frame = tk.Frame(login_card, bg=self.colors["card_bg"])
        login_frame.pack(fill="x", padx=16, pady=8)

        tk.Label(login_frame, text="🔑 ĐĂNG NHẬP WMS:", bg=self.colors["card_bg"], fg=self.colors["text"], font=("Segoe UI", 10, "bold")).pack(side="left", padx=(0, 10))

        tk.Label(login_frame, text="Mã NV:", bg=self.colors["card_bg"], fg=self.colors["muted"], font=("Segoe UI", 9)).pack(side="left", padx=(0, 4))
        self.entry_user = tk.Entry(login_frame, font=("Segoe UI", 10), bg="#0F172A", fg="#F8FAFC", insertbackground="#FFFFFF", bd=1, relief="solid", width=14)
        self.entry_user.pack(side="left", padx=(0, 10), ipady=3)
        self.entry_user.insert(0, DEFAULT_USER)

        tk.Label(login_frame, text="Mật Khẩu:", bg=self.colors["card_bg"], fg=self.colors["muted"], font=("Segoe UI", 9)).pack(side="left", padx=(0, 4))
        self.entry_pass = tk.Entry(login_frame, font=("Segoe UI", 10), show="*", bg="#0F172A", fg="#F8FAFC", insertbackground="#FFFFFF", bd=1, relief="solid", width=14)
        self.entry_pass.pack(side="left", padx=(0, 10), ipady=3)
        self.entry_pass.insert(0, DEFAULT_PASS)

        tk.Label(login_frame, text="Station:", bg=self.colors["card_bg"], fg=self.colors["muted"], font=("Segoe UI", 9)).pack(side="left", padx=(0, 4))
        self.entry_station = tk.Entry(login_frame, font=("Segoe UI", 10), bg="#0F172A", fg="#F8FAFC", insertbackground="#FFFFFF", bd=1, relief="solid", width=8)
        self.entry_station.pack(side="left", padx=(0, 15), ipady=3)
        self.entry_station.insert(0, DEFAULT_STATION)

        self.btn_login_wms = AnimatedButton(
            login_frame, "⚡ KẾT NỐI ĐĂNG NHẬP WMS", self.login_wms_thread,
            self.colors["success"], self.colors["success_hover"], font=("Segoe UI Bold", 9), pad_x=14, pad_y=4
        )
        self.btn_login_wms.pack(side="right")

        # 3. STAT CARDS
        dash = tk.Frame(self.root, bg=self.colors["bg"])
        dash.pack(fill="x", padx=24, pady=(10, 8))

        c1, self.stat_total = self._create_stat_card(dash, "TỔNG SỐ LẦN QUÉT", "0", "📊", self.colors["primary"])
        c1.pack(side="left", fill="x", expand=True, padx=(0, 8))

        c2, self.stat_success = self._create_stat_card(dash, "THÀNH CÔNG", "0", "✓", self.colors["success"])
        c2.pack(side="left", fill="x", expand=True, padx=8)

        c3, self.stat_notfound = self._create_stat_card(dash, "KHÔNG TÌM THẤY", "0", "🔍", self.colors["warning"])
        c3.pack(side="left", fill="x", expand=True, padx=8)

        c4, self.stat_error = self._create_stat_card(dash, "LỖI HỆ THỐNG", "0", "⚠", self.colors["danger"])
        c4.pack(side="left", fill="x", expand=True, padx=(8, 0))

        # 4. INPUT AREA
        input_card = tk.Frame(self.root, bg=self.colors["card_bg"], highlightbackground=self.colors["card_border"], highlightthickness=1)
        input_card.pack(fill="x", padx=24, pady=(0, 10))

        input_top = tk.Frame(input_card, bg=self.colors["card_bg"])
        input_top.pack(fill="x", padx=20, pady=(12, 4))

        tk.Label(input_top, text="🎯 NHẬP / QUÉT MÃ MÁY QUÉT", bg=self.colors["card_bg"], fg=self.colors["text"], font=("Segoe UI", 11, "bold")).pack(side="left")

        self.auto_print_var = tk.BooleanVar(value=True)
        auto_chk = tk.Checkbutton(
            input_top, text="Tự động gửi lệnh in khi tìm thấy data", variable=self.auto_print_var,
            bg=self.colors["card_bg"], fg=self.colors["success"], activebackground=self.colors["card_bg"],
            activeforeground=self.colors["success"], selectcolor=self.colors["card_bg"],
            font=("Segoe UI", 9, "bold"), cursor="hand2"
        )
        auto_chk.pack(side="right")

        input_row = tk.Frame(input_card, bg=self.colors["card_bg"])
        input_row.pack(fill="x", padx=20, pady=(4, 16))

        self.entry_input_code = tk.Entry(
            input_row, font=("Segoe UI", 14, "bold"), bg="#0F172A", fg="#F8FAFC",
            insertbackground="#FFFFFF", bd=1, relief="solid", highlightcolor=self.colors["primary"]
        )
        self.entry_input_code.pack(side="left", fill="x", expand=True, ipady=8, padx=(0, 12))
        self.entry_input_code.bind("<Return>", self.start_search)

        self.btn_search = AnimatedButton(input_row, "🔎 TRA CỨU & IN TEM", self.start_search, self.colors["success"], self.colors["success_hover"], font=("Segoe UI Bold", 11), pad_y=10)
        self.btn_search.pack(side="right")

        # STATUS BAR
        status_card = tk.Frame(self.root, bg=self.colors["bg"])
        status_card.pack(fill="x", padx=24, pady=(0, 8))

        self.status_dot = tk.Label(status_card, text="●", bg=self.colors["bg"], fg=self.colors["primary"], font=("Segoe UI", 12))
        self.status_dot.pack(side="left", padx=(2, 6))

        self.status_label = tk.Label(status_card, text="Vui lòng bấm 'KẾT NỐI ĐĂNG NHẬP WMS' để bắt đầu...", bg=self.colors["bg"], fg=self.colors["muted"], font=("Segoe UI", 9, "bold"), anchor="w")
        self.status_label.pack(side="left", fill="x", expand=True)

        # 5. TABLE HISTORY (TỐI ƯU CỘT VỚI KÍCH THƯỚC SIÊU RỘNG)
        table_card = tk.Frame(self.root, bg=self.colors["card_bg"], highlightbackground=self.colors["card_border"], highlightthickness=1)
        table_card.pack(fill="both", expand=True, padx=24, pady=(0, 10))

        table_header = tk.Frame(table_card, bg=self.colors["card_bg"])
        table_header.pack(fill="x", padx=18, pady=(10, 6))

        tk.Label(table_header, text="LỊCH SỬ DỮ LIỆU HIỂN THỊ (Click đúp dòng để xem chi tiết ảnh QC / Ảnh sản phẩm)", bg=self.colors["card_bg"], fg=self.colors["text"], font=("Segoe UI", 11, "bold")).pack(side="left")

        table_frame = tk.Frame(table_card, bg=self.colors["card_bg"])
        table_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        columns = ("time", "input_code", "found_data", "sku_name", "total_price", "sku_id", "follow_up", "inbound_type", "cb_option", "has_image", "has_qc_image", "status")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", style="Treeview")

        headings = {
            "time": "THỜI GIAN",
            "input_code": "MÃ QUÉT",
            "found_data": "DỮ LIỆU TÌM THẤY",
            "sku_name": "TÊN SP (SKU NAME)",
            "total_price": "TỔNG GIÁ TIỀN",
            "sku_id": "SKU ID",
            "follow_up": "FOLLOW UP",
            "inbound_type": "INBOUND TYPE",
            "cb_option": "CB OPTION",
            "has_image": "ẢNH SP",
            "has_qc_image": "ẢNH QC",
            "status": "TRẠNG THÁI"
        }

        col_widths = {
            "time": 110,
            "input_code": 170,
            "found_data": 170,
            "sku_name": 380,
            "total_price": 120,
            "sku_id": 200,
            "follow_up": 160,
            "inbound_type": 130,
            "cb_option": 90,
            "has_image": 85,
            "has_qc_image": 85,
            "status": 110
        }

        for col in columns:
            self.tree.heading(col, text=headings[col], anchor="center")
            self.tree.column(col, anchor="center", width=col_widths.get(col, 120), minwidth=70, stretch=True)

        self.tree.tag_configure("success", background="#15803D", foreground="#FFFFFF")
        self.tree.tag_configure("notfound", background="#B45309", foreground="#FFFFFF")
        self.tree.tag_configure("error", background="#B91C1C", foreground="#FFFFFF")

        scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self.on_tree_double_click)

        # FOOTER
        footer = tk.Frame(self.root, bg=self.colors["bg"], height=40)
        footer.pack(fill="x", padx=24, pady=(0, 10))
        footer.pack_propagate(False)

        AnimatedButton(footer, "🧹 Xóa Lịch Sử", self.clear_screen_history, "#334155", "#475569", pad_y=4).pack(side="left")
        AnimatedButton(footer, "✕ Thoát Application", self.root.destroy, self.colors["danger"], self.colors["danger_hover"], pad_y=4).pack(side="right")

        self.entry_input_code.focus_set()

    def _update_stats(self):
        total = len(self.all_records)
        success = sum(1 for r in self.all_records if r["status"] == "Thành công")
        notfound = sum(1 for r in self.all_records if r["status"] == "Không tìm thấy")
        errors = total - success - notfound

        self.stat_total.config(text=str(total))
        self.stat_success.config(text=str(success))
        self.stat_notfound.config(text=str(notfound))
        self.stat_error.config(text=str(errors))

    def set_status(self, message, color):
        self.status_label.config(text=message, fg=color)
        self.status_dot.config(fg=color)

    # ============================================================
    # TÍNH NĂNG ĐĂNG NHẬP WMS BẰNG SELENIUM
    # ============================================================

    def _clean_wdm_locks(self):
        try:
            wdm_dir = os.path.expanduser("~/.wdm")
            if os.path.exists(wdm_dir):
                for root, dirs, files in os.walk(wdm_dir):
                    for f in files:
                        if f.endswith(".lock") or "lock" in f:
                            try:
                                os.remove(os.path.join(root, f))
                            except Exception:
                                pass
        except Exception:
            pass

    def login_wms_thread(self):
        user = self.entry_user.get().strip()
        password = self.entry_pass.get().strip()
        station_code = self.entry_station.get().strip() or "1"

        if not user or not password:
            self.set_status("⚠️ Vui lòng nhập đầy đủ Tài khoản và Mật khẩu WMS!", self.colors["warning"])
            return

        self.btn_login_wms.config(state="disabled")
        self.set_status(f"Đang kết nối đăng nhập WMS (NV {user} - Station {station_code})...", self.colors["primary"])
        threading.Thread(target=self._worker_login_wms, args=(user, password, station_code), daemon=True).start()

    def _worker_login_wms(self, user, password, station_code):
        self._clean_wdm_locks()

        chrome_options = Options()
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")

        driver = None
        try:
            try:
                driver = webdriver.Chrome(options=chrome_options)
            except Exception:
                if HAS_WD_MANAGER:
                    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
                else:
                    raise

            wait = WebDriverWait(driver, 30)

            self.root.after(0, lambda: self.set_status("Đang mở trang đăng nhập WMS...", self.colors["primary"]))
            driver.get(URL_WMS)

            user_input = wait.until(EC.presence_of_element_located((
                By.XPATH, '//input[@id="warehouse-management-system-authKey"] | //input[@placeholder="Email/Username"] | //input[contains(@class, "el-input__inner")]'
            )))
            user_input.clear()
            user_input.send_keys(user)

            pass_input = driver.find_element(
                By.XPATH, '//input[@id="warehouse-management-system-password"] | //input[@type="password"]'
            )
            pass_input.clear()
            pass_input.send_keys(password)
            time.sleep(1)
            pass_input.send_keys(Keys.ENTER)

            self.root.after(0, lambda: self.set_status(f"Đang chọn Station WMS ({station_code})...", self.colors["primary"]))
            time.sleep(4)

            try:
                station_box = wait.until(EC.element_to_be_clickable((
                    By.XPATH, '//span[contains(@class, "station")]//input | //*[@id="app"]//section[1]//input'
                )))
                station_box.click()
                time.sleep(1)

                station_elements = driver.find_elements(
                    By.XPATH, f'//li[contains(text(), "{station_code}")] | //ul/li[contains(., "{station_code}")]'
                )
                if station_elements:
                    station_elements[0].click()
                else:
                    if station_code.isdigit():
                        idx = int(station_code)
                        driver.find_element(By.XPATH, f'//ul/li[{idx}]').click()
                time.sleep(2)
            except Exception as st_err:
                print(f"Lưu ý chọn station: {st_err}")

            cookies = driver.get_cookies()
            self.cookie_val = "; ".join([f"{c['name']}={c['value']}" for c in cookies])
            
            for c in cookies:
                if c['name'] == 'csrftoken':
                    self.csrf_token = c['value']
                    break

            def _on_login_success():
                self.set_status(f"✅ Đăng nhập WMS thành công! (NV: {user} | Station: {station_code})", self.colors["success"])
                self.btn_login_wms.config(state="normal")

            self.root.after(0, _on_login_success)

        except Exception as e:
            err_msg = str(e)
            if "wdm-lock" in err_msg or "lock" in err_msg:
                friendly_msg = "Xử lý xóa file lock tự động. Vui lòng bấm đăng nhập lại!"
            else:
                friendly_msg = f"Lỗi đăng nhập: {err_msg[:80]}"

            def _show_err():
                self.set_status(f"❌ {friendly_msg}", self.colors["danger"])
                self.btn_login_wms.config(state="normal")

            self.root.after(0, _show_err)
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass

    # ============================================================
    # TRA CỨU API & XỬ LÝ DỮ LIỆU
    # ============================================================

    def start_search(self, event=None):
        input_code = self.entry_input_code.get().strip().upper()
        if not input_code:
            self.set_status("Vui lòng quét hoặc nhập mã!", self.colors["danger"])
            return

        if not self.cookie_val:
            self.set_status("⚠️ Chưa đăng nhập WMS! Vui lòng kết nối trước.", self.colors["warning"])
            return

        self.btn_search.config(state="disabled")
        self.set_status(f"Đang tra cứu dữ liệu [{input_code}]...", self.colors["warning"])
        threading.Thread(target=self.process_search_and_print, args=(input_code,), daemon=True).start()

    def fetch_wms_api(self, search_key, mode):
        csrf_token = self.csrf_token if self.csrf_token else HARDCODED_X_CSRFTOKEN

        headers = {
            "accept": "application/json, text/plain, */*",
            "cookie": self.cookie_val,
            "referer": "https://wms.ssc.shopee.vn/v2/returninbound/order",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/152.0.0.0 Safari/537.36",
            "x-cctv-tenant-id": "WMS",
            "x-csrftoken": csrf_token,
            "wms-cid": "VN",
        }
        params = {"sort_field": "", "sort_type": 1, "pageno": 1, "count": 20}
        params["return_tn" if mode == "RETURN_TN" else "order_no"] = search_key

        response = self.session.get(f"{BASE_URL}{ENDPOINT_SEARCH_ASN}", headers=headers, params=params, timeout=25)
        response.raise_for_status()
        return response.json()

    def fetch_wms_sku_api(self, inbound_id):
        csrf_token = self.csrf_token if self.csrf_token else HARDCODED_X_CSRFTOKEN

        headers = {
            "accept": "application/json, text/plain, */*",
            "cookie": self.cookie_val,
            "referer": f"https://wms.ssc.shopee.vn/v2/returninbound/order/detail/{inbound_id}",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/152.0.0.0 Safari/537.36",
            "x-cctv-tenant-id": "WMS",
            "x-csrftoken": csrf_token,
            "wms-cid": "VN",
        }
        params = {"pageno": 1, "count": 50, "inbound_id": str(inbound_id)}

        response = self.session.get(f"{BASE_URL}{ENDPOINT_SEARCH_ASN_SKU}", headers=headers, params=params, timeout=25)
        response.raise_for_status()
        return response.json()

    def fetch_wms_asn_detail_api(self, inbound_id):
        csrf_token = self.csrf_token if self.csrf_token else HARDCODED_X_CSRFTOKEN

        headers = {
            "accept": "application/json, text/plain, */*",
            "cookie": self.cookie_val,
            "referer": f"https://wms.ssc.shopee.vn/v2/returninbound/order/detail/{inbound_id}",
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/152.0.0.0 Safari/537.36",
            "x-cctv-tenant-id": "WMS",
            "x-csrftoken": csrf_token,
            "wms-cid": "VN",
        }
        params = {"inbound_id": str(inbound_id)}

        response = self.session.get(f"{BASE_URL}{ENDPOINT_SEARCH_ASN_DETAIL}", headers=headers, params=params, timeout=25)
        response.raise_for_status()
        return response.json()

    def process_search_and_print(self, input_code):
        sku_id_str, follow_up_str, inbound_type_str, cb_option_str, total_price_str = "-", "-", "-", "-", "-"
        sku_name_str = "-"
        all_sku_images, all_qc_images = [], []

        try:
            detected_mode = "RETURN_TN"
            try:
                api_result = self.fetch_wms_api(input_code, "RETURN_TN")
                data = api_result.get("data") or {}
                records = data.get("list") or data.get("items") or []
            except Exception:
                records = []

            if not records:
                detected_mode = "ORDER_NO"
                try:
                    api_result = self.fetch_wms_api(input_code, "ORDER_NO")
                    data = api_result.get("data") or {}
                    records = data.get("list") or data.get("items") or []
                except Exception:
                    records = []

            if records:
                status = "Thành công"
                color = self.colors["success"]

                _, fu_1, in_1, cb_1, pr_1, name_1, img_1, qc_1 = extract_wms_details_deep(api_result)

                inbound_ids = set()
                all_cb_list = [cb_1] if cb_1 != "-" else []
                all_sku_name_list = [name_1] if name_1 != "-" else []
                all_sku_images.extend(img_1)
                all_qc_images.extend(qc_1)

                grand_total_price = pr_1 if isinstance(pr_1, (int, float)) else 0.0

                for item in records:
                    if isinstance(item, dict):
                        mapped_cb = normalize_cb_option(item.get("cb_flag") or item.get("cb_option"))
                        if mapped_cb != "-":
                            all_cb_list.append(mapped_cb)

                        for key in ["inbound_id", "id", "riasn_id", "return_inbound_id"]:
                            val = item.get(key)
                            if val:
                                inbound_ids.add(str(val))

                all_sku_list = []
                all_follow_list = [fu_1] if fu_1 != "-" else []
                all_inbound_list = [in_1] if in_1 != "-" else []

                for ib_id in inbound_ids:
                    try:
                        asn_detail_res = self.fetch_wms_asn_detail_api(ib_id)
                        _, f_dt, i_dt, cb_dt, pr_dt, n_dt, img_dt, qc_dt = extract_wms_details_deep(asn_detail_res)
                        if f_dt != "-": all_follow_list.append(f_dt)
                        if i_dt != "-": all_inbound_list.append(i_dt)
                        if cb_dt != "-": all_cb_list.append(cb_dt)
                        if n_dt != "-": all_sku_name_list.append(n_dt)
                        all_sku_images.extend(img_dt)
                        all_qc_images.extend(qc_dt)
                        if isinstance(pr_dt, (int, float)): grand_total_price += pr_dt
                    except Exception:
                        pass

                    try:
                        sku_api_res = self.fetch_wms_sku_api(ib_id)
                        s_res, f_res, i_res, cb_res, pr_res, n_res, img_res, qc_res = extract_wms_details_deep(sku_api_res)

                        if s_res != "-": all_sku_list.append(s_res)
                        if f_res != "-": all_follow_list.append(f_res)
                        if i_res != "-": all_inbound_list.append(i_res)
                        if cb_res != "-": all_cb_list.append(cb_res)
                        if n_res != "-": all_sku_name_list.append(n_res)
                        all_sku_images.extend(img_res)
                        all_qc_images.extend(qc_res)
                        if isinstance(pr_res, (int, float)): grand_total_price += pr_res
                    except Exception:
                        pass

                sku_id_str = ", ".join(list(dict.fromkeys(all_sku_list))) if all_sku_list else "-"
                follow_up_str = ", ".join(list(dict.fromkeys(all_follow_list))) if all_follow_list else "-"
                inbound_type_str = ", ".join(list(dict.fromkeys(all_inbound_list))) if all_inbound_list else "-"
                sku_name_str = " | ".join(list(dict.fromkeys(all_sku_name_list))) if all_sku_name_list else "-"

                all_sku_images = list(dict.fromkeys(all_sku_images))
                all_qc_images = list(dict.fromkeys(all_qc_images))

                total_price_str = format_currency(grand_total_price)

                unique_cb_values = list(dict.fromkeys(all_cb_list))
                cb_option_str = "CB" if "CB" in unique_cb_values else ("Local" if "Local" in unique_cb_values else "-")

                if detected_mode == "RETURN_TN":
                    all_orders = [o for item in records if isinstance(item, dict) for o in (item.get("order_no_list") or [])]
                    found_data = ", ".join(list(dict.fromkeys(all_orders))) if all_orders else input_code
                    sub_info = f"Return TN: {input_code}"
                else:
                    all_rtn = [item.get("return_tn") for item in records if isinstance(item, dict) and item.get("return_tn")]
                    found_data = ", ".join(list(dict.fromkeys(all_rtn))) if all_rtn else input_code
                    sub_info = f"Order: {input_code}"

                if self.auto_print_var.get():
                    label_img = generate_label_image(
                        print_value=found_data,
                        sub_info=sub_info,
                        followup_info=follow_up_str,
                        inbound_type=inbound_type_str,
                        cb_option=cb_option_str,
                        total_price=total_price_str
                    )
                    self.root.after(0, print_label_direct, label_img)

            else:
                status = "Không tìm thấy"
                found_data = "-"
                color = self.colors["warning"]

        except Exception as error:
            status = "Lỗi API"
            found_data = str(error)
            color = self.colors["danger"]

        self.root.after(
            0, self.update_result, input_code, found_data, total_price_str, sku_id_str, 
            follow_up_str, inbound_type_str, cb_option_str, sku_name_str, all_sku_images, all_qc_images, status, color
        )

    def update_result(self, input_code, found_data, total_price, sku_id, follow_up, inbound_type, cb_option, sku_name, sku_images, qc_images, status, color):
        current_time = datetime.now().strftime("%H:%M:%S %d-%m")
        has_image_str = f"🖼️ ({len(sku_images)})" if sku_images else "-"
        has_qc_str = f"📸 ({len(qc_images)})" if qc_images else "-"

        record_data = {
            "time": current_time,
            "input_code": input_code,
            "found_data": found_data,
            "sku_name": sku_name,
            "total_price": total_price,
            "sku_id": sku_id,
            "follow_up": follow_up,
            "inbound_type": inbound_type,
            "cb_option": cb_option,
            "has_image": has_image_str,
            "has_qc_image": has_qc_str,
            "status": status,
            "sku_images": sku_images,
            "qc_images": qc_images
        }

        self.all_records.append(record_data)

        tag = "success" if status == "Thành công" else ("notfound" if status == "Không tìm thấy" else "error")
        tree_values = (
            current_time, input_code, found_data, sku_name, total_price, 
            sku_id, follow_up, inbound_type, cb_option, has_image_str, has_qc_str, status
        )
        item_id = self.tree.insert("", 0, values=tree_values, tags=(tag,))
        
        record_data["tree_id"] = item_id

        self._update_stats()
        self.set_status(f"Mã: {input_code} | Trạng thái: {status}", color)
        self.entry_input_code.delete(0, tk.END)
        self.entry_input_code.focus_set()
        self.btn_search.config(state="normal")

    def clear_screen_history(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.all_records.clear()
        self.photo_cache.clear()
        self._update_stats()
        self.set_status("Đã xóa sạch dữ liệu lịch sử hiển thị.", self.colors["primary"])

    # ============================================================
    # POPUP CỬA SỔ HIỂN THỊ CHI TIẾT ẢNH SP & ẢNH QC
    # ============================================================

    def on_tree_double_click(self, event):
        selected_item = self.tree.selection()
        if not selected_item:
            return

        item_id = selected_item[0]
        match_record = next((r for r in self.all_records if r.get("tree_id") == item_id), None)
        
        if not match_record:
            return

        self.show_image_gallery_popup(match_record)

    def show_image_gallery_popup(self, record):
        popup = tk.Toplevel(self.root)
        popup.title(f"Chi Tiết Dữ Liệu & Hình Ảnh - [{record['input_code']}]")
        
        popup.geometry("1200x800")
        popup.configure(bg="#0F172A")
        popup.transient(self.root)

        top_info = tk.Frame(popup, bg="#1E293B", padx=20, pady=12)
        top_info.pack(fill="x")

        tk.Label(top_info, text=f"📦 Mã Quét: {record['input_code']}", font=("Segoe UI", 13, "bold"), fg="#38BDF8", bg="#1E293B").pack(anchor="w")
        tk.Label(top_info, text=f"🏷️ Tên SKU: {record['sku_name']}", font=("Segoe UI", 11), fg="#F8FAFC", bg="#1E293B", wraplength=1150, justify="left").pack(anchor="w", pady=(3, 0))

        canvas_frame = tk.Frame(popup, bg="#0F172A")
        canvas_frame.pack(fill="both", expand=True, padx=15, pady=15)

        canvas = tk.Canvas(canvas_frame, bg="#0F172A", highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scroll_content = tk.Frame(canvas, bg="#0F172A")

        scroll_content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_content, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        sku_imgs = record.get("sku_images", [])
        qc_imgs = record.get("qc_images", [])

        if not sku_imgs and not qc_imgs:
            tk.Label(scroll_content, text="🚫 Không tìm thấy liên kết hình ảnh nào cho mã này.", font=("Segoe UI", 13), fg="#94A3B8", bg="#0F172A").pack(pady=50)
            return

        def download_and_render_images():
            headers = {"cookie": self.cookie_val, "user-agent": "Mozilla/5.0"}

            # Section 1: Ảnh Sản Phẩm
            if sku_imgs:
                lbl_sec1 = tk.Label(scroll_content, text="🖼️ HÌNH ẢNH SẢN PHẨM (SKU IMAGE)", font=("Segoe UI", 12, "bold"), fg="#22C55E", bg="#0F172A")
                lbl_sec1.pack(anchor="w", pady=(10, 8))
                
                grid_sku = tk.Frame(scroll_content, bg="#0F172A")
                grid_sku.pack(fill="x", pady=5)

                for idx, img_url in enumerate(sku_imgs):
                    try:
                        resp = requests.get(img_url, headers=headers, timeout=10)
                        if resp.status_code == 200:
                            pil_img = Image.open(BytesIO(resp.content)).resize((240, 240), Image.Resampling.LANCZOS)
                            photo = ImageTk.PhotoImage(pil_img)
                            self.photo_cache.append(photo)

                            btn_img = tk.Button(grid_sku, image=photo, bg="#1E293B", activebackground="#334155", bd=1, relief="solid", command=lambda u=img_url: webbrowser.open(u))
                            btn_img.grid(row=idx // 4, column=idx % 4, padx=10, pady=10)
                    except Exception:
                        pass

            # Section 2: Ảnh QC / Bằng chứng
            if qc_imgs:
                lbl_sec2 = tk.Label(scroll_content, text="📸 HÌNH ẢNH KIỂM HÀNG / QC (QC IMAGE LIST)", font=("Segoe UI", 12, "bold"), fg="#F59E0B", bg="#0F172A")
                lbl_sec2.pack(anchor="w", pady=(25, 8))

                grid_qc = tk.Frame(scroll_content, bg="#0F172A")
                grid_qc.pack(fill="x", pady=5)

                for idx, img_url in enumerate(qc_imgs):
                    try:
                        resp = requests.get(img_url, headers=headers, timeout=10)
                        if resp.status_code == 200:
                            pil_img = Image.open(BytesIO(resp.content)).resize((240, 240), Image.Resampling.LANCZOS)
                            photo = ImageTk.PhotoImage(pil_img)
                            self.photo_cache.append(photo)

                            box = tk.Frame(grid_qc, bg="#1E293B", bd=1, relief="solid")
                            box.grid(row=idx // 4, column=idx % 4, padx=10, pady=10)

                            btn_img = tk.Button(box, image=photo, bg="#1E293B", activebackground="#334155", bd=0, command=lambda u=img_url: webbrowser.open(u))
                            btn_img.pack()
                            
                            tk.Label(box, text=f"Ảnh QC #{idx+1} (Click để mở Web)", font=("Segoe UI", 9), fg="#94A3B8", bg="#1E293B").pack(pady=4)
                    except Exception:
                        pass

        threading.Thread(target=download_and_render_images, daemon=True).start()


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = ScanReturnTNApp(root)
    root.mainloop()
