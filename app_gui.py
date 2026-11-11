#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GIAO DIỆN TRỰC QUAN (DESKTOP GUI) - PHẦN MỀM TỰ ĐỘNG ĐẶT TÊN FILE PDF HỒ SƠ
Tính năng chính:
  1. Tải lên (Upload) file Excel danh mục đặt tên mới bất kỳ (.xlsx) hoặc chỉnh sửa trực tiếp trên bảng.
  2. Quét OCR & Đối chiếu tự động (Dry-Run) tốc độ cao nhờ SQLite Cache (chỉ mất vài giây khi đổi Excel mới).
  3. Xem trước trực tiếp trang PDF, xem nội dung OCR và sửa tên file dự kiến ngay trên giao diện.
  4. Thực thi đổi tên hàng loạt an toàn 2 bước (Two-Phase Rename) và Hoàn tác (Rollback) 1 chạm.
"""

import os
import sys
import time
import shutil
import pathlib
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pymupdf
from PIL import Image, ImageTk

import core_engine as ce


class PDFRenameApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Phần mềm Tự động Nhận diện OCR & Đặt tên File PDF Hồ sơ (Hỗ trợ Tải lên Danh mục Excel)")
        self.geometry("1440x860")
        self.minsize(1180, 720)
        self.configure(bg="#F3F4F6")

        # Cấu hình & Dữ liệu
        self.cfg = ce.load_config()
        self.catalog = ce.reload_global_catalog(pathlib.Path(self.cfg['excel_catalog_path']))
        self.all_records = []
        self.filtered_records = []
        self.current_page = 1
        self.page_size = 200
        self.selected_record = None

        # Trạng thái xem trước PDF
        self.preview_pdf_path = None
        self.preview_page_idx = 0
        self.preview_total_pages = 0
        self.preview_photo = None

        # Tiến trình chạy ngầm
        self.worker_thread = None
        self.stop_event = threading.Event()

        self._setup_styles()
        self._build_ui()
        self._refresh_catalog_ui()
        self._load_initial_records()

    # ==========================================================================
    # THIẾT LẬP GIAO DIỆN & STYLE
    # ==========================================================================

    def _setup_styles(self):
        style = ttk.Style(self)
        if 'clam' in style.theme_names():
            style.theme_use('clam')

        style.configure("TFrame", background="#F3F4F6")
        style.configure("Card.TFrame", background="#FFFFFF", relief="flat")
        style.configure("TLabel", background="#F3F4F6", foreground="#1F2937", font=("Segoe UI", 10))
        style.configure("Card.TLabel", background="#FFFFFF", foreground="#1F2937", font=("Segoe UI", 10))
        style.configure("Header.TLabel", background="#1E3A8A", foreground="#FFFFFF", font=("Segoe UI", 14, "bold"))
        style.configure("SubHeader.TLabel", background="#1E3A8A", foreground="#DBEAFE", font=("Segoe UI", 9))

        style.configure("TNotebook", background="#F3F4F6", borderwidth=0)
        style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=(16, 8), background="#E5E7EB", foreground="#374151")
        style.map("TNotebook.Tab", background=[("selected", "#1D4ED8")], foreground=[("selected", "#FFFFFF")])

        style.configure("Treeview", font=("Segoe UI", 9), rowheight=26, background="#FFFFFF", fieldbackground="#FFFFFF")
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#E5E7EB", foreground="#111827", padding=5)
        style.map("Treeview", background=[("selected", "#DBEAFE")], foreground=[("selected", "#1E3A8A")])

        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), background="#2563EB", foreground="#FFFFFF", padding=(12, 6))
        style.map("Primary.TButton", background=[("active", "#1D4ED8"), ("disabled", "#9CA3AF")])

        style.configure("Success.TButton", font=("Segoe UI", 10, "bold"), background="#059669", foreground="#FFFFFF", padding=(12, 6))
        style.map("Success.TButton", background=[("active", "#047857"), ("disabled", "#9CA3AF")])

        style.configure("Warning.TButton", font=("Segoe UI", 10, "bold"), background="#D97706", foreground="#FFFFFF", padding=(12, 6))
        style.map("Warning.TButton", background=[("active", "#B45309"), ("disabled", "#9CA3AF")])

        style.configure("Danger.TButton", font=("Segoe UI", 10, "bold"), background="#DC2626", foreground="#FFFFFF", padding=(12, 6))
        style.map("Danger.TButton", background=[("active", "#B91C1C"), ("disabled", "#9CA3AF")])

        style.configure("Secondary.TButton", font=("Segoe UI", 9), background="#FFFFFF", foreground="#1F2937", padding=(10, 5))
        style.map("Secondary.TButton", background=[("active", "#E5E7EB")])

    def _build_ui(self):
        # 1. BANNER ĐẦU TRANG
        banner = tk.Frame(self, bg="#1E3A8A", height=62)
        banner.pack(fill="x", side="top")
        banner.pack_propagate(False)

        title_box = tk.Frame(banner, bg="#1E3A8A")
        title_box.pack(side="left", padx=18, pady=8)
        ttk.Label(
            title_box,
            text="🗂️ HỆ THỐNG TỰ ĐỘNG NHẬN DIỆN OCR & ĐẶT TÊN FILE PDF HỒ SƠ",
            style="Header.TLabel"
        ).pack(anchor="w")
        ttk.Label(
            title_box,
            text="Hỗ trợ tải lên danh mục Excel đặt tên tùy chỉnh • Xem trước PDF trực quan • Đổi tên 2 bước an toàn & Hoàn tác",
            style="SubHeader.TLabel"
        ).pack(anchor="w")

        btn_banner_right = tk.Frame(banner, bg="#1E3A8A")
        btn_banner_right.pack(side="right", padx=18)
        ttk.Button(
            btn_banner_right,
            text="📂 Mở thư mục Báo cáo Excel",
            style="Secondary.TButton",
            command=self._open_report_folder
        ).pack(side="right", padx=4)

        # 2. KHUNG CẤU HÌNH & TẢI LÊN EXCEL
        cfg_panel = tk.Frame(self, bg="#FFFFFF", bd=1, relief="solid")
        cfg_panel.pack(fill="x", padx=12, pady=(10, 6))

        # Dòng 1: Thư mục dữ liệu & Số luồng
        row1 = tk.Frame(cfg_panel, bg="#FFFFFF")
        row1.pack(fill="x", padx=12, pady=(8, 4))

        tk.Label(row1, text="📁 Thư mục hồ sơ PDF:", bg="#FFFFFF", fg="#111827", font=("Segoe UI", 9, "bold"), width=22, anchor="w").pack(side="left")
        self.var_root_dir = tk.StringVar(value=self.cfg.get('root_dir', str(ce._DEFAULT_ROOT_DIR)))
        tk.Entry(row1, textvariable=self.var_root_dir, font=("Segoe UI", 9), bg="#F9FAFB", relief="solid", bd=1).pack(side="left", fill="x", expand=True, padx=(0, 6), ipady=3)
        ttk.Button(row1, text="Chọn thư mục...", style="Secondary.TButton", command=self._choose_root_dir).pack(side="left", padx=(0, 14))

        tk.Label(row1, text="Số luồng OCR:", bg="#FFFFFF", fg="#374151", font=("Segoe UI", 9)).pack(side="left", padx=(4, 4))
        self.var_workers = tk.IntVar(value=int(self.cfg.get('workers', 10)))
        ttk.Spinbox(row1, from_=1, to=32, textvariable=self.var_workers, width=5, font=("Segoe UI", 9)).pack(side="left", padx=(0, 14))

        self.var_allow_outside = tk.BooleanVar(value=bool(self.cfg.get('allow_outside_excel', True)))
        ttk.Checkbutton(
            row1,
            text="Đặt tên theo tiêu đề trang đầu nếu ngoài danh mục Excel",
            variable=self.var_allow_outside,
            command=self._save_current_config
        ).pack(side="left")

        # Dòng 2: File Excel danh mục đặt tên + Nút Tải lên
        row2 = tk.Frame(cfg_panel, bg="#FFFFFF")
        row2.pack(fill="x", padx=12, pady=(4, 8))

        tk.Label(row2, text="📊 File Excel danh mục tên:", bg="#FFFFFF", fg="#111827", font=("Segoe UI", 9, "bold"), width=22, anchor="w").pack(side="left")
        self.var_excel_path = tk.StringVar(value=self.cfg.get('excel_catalog_path', ''))
        tk.Entry(row2, textvariable=self.var_excel_path, font=("Segoe UI", 9, "bold"), fg="#1E40AF", bg="#EFF6FF", relief="solid", bd=1, state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 6), ipady=3)

        self.lbl_catalog_badge = tk.Label(
            row2,
            text="0 mục",
            bg="#DBEAFE",
            fg="#1E3A8A",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=3
        )
        self.lbl_catalog_badge.pack(side="left", padx=(0, 8))

        ttk.Button(
            row2,
            text="📤 Tải lên File Excel danh mục mới (.xlsx)...",
            style="Primary.TButton",
            command=self._upload_new_excel_catalog
        ).pack(side="left", padx=(0, 6))

        ttk.Button(
            row2,
            text="📥 Tải file Excel mẫu",
            style="Secondary.TButton",
            command=self._export_template_excel
        ).pack(side="left")

        # 3. THANH ĐIỀU KHIỂN HÀNH ĐỘNG & THẺ THỐNG KÊ NHANH
        action_bar = tk.Frame(self, bg="#F3F4F6")
        action_bar.pack(fill="x", padx=12, pady=(2, 6))

        # Các nút tác vụ chính bên trái
        btn_group = tk.Frame(action_bar, bg="#F3F4F6")
        btn_group.pack(side="left")

        self.btn_dry_run = ttk.Button(
            btn_group,
            text="🔍 1. Quét & Dự kiến đổi tên (Dry-Run)",
            style="Primary.TButton",
            command=self._start_dry_run
        )
        self.btn_dry_run.pack(side="left", padx=(0, 6))

        self.btn_stop = ttk.Button(
            btn_group,
            text="⏹ Dừng",
            style="Danger.TButton",
            command=self._stop_dry_run,
            state="disabled"
        )
        self.btn_stop.pack(side="left", padx=(0, 12))

        self.btn_execute = ttk.Button(
            btn_group,
            text="✅ 2. Thực thi đổi tên thật",
            style="Success.TButton",
            command=self._execute_rename
        )
        self.btn_execute.pack(side="left", padx=(0, 6))

        self.btn_rollback = ttk.Button(
            btn_group,
            text="⏪ 3. Hoàn tác (Rollback)",
            style="Warning.TButton",
            command=self._rollback_rename
        )
        self.btn_rollback.pack(side="left", padx=(0, 6))

        # Các thẻ thống kê bên phải
        cards_frame = tk.Frame(action_bar, bg="#F3F4F6")
        cards_frame.pack(side="right")

        self.lbl_stat_total = self._create_stat_badge(cards_frame, "Tổng số PDF", "0", "#1F2937", "#E5E7EB")
        self.lbl_stat_excel = self._create_stat_badge(cards_frame, "Khớp Excel", "0", "#065F46", "#D1FAE5")
        self.lbl_stat_outside = self._create_stat_badge(cards_frame, "Ngoài Excel", "0", "#1E40AF", "#DBEAFE")
        self.lbl_stat_dup = self._create_stat_badge(cards_frame, "Trùng (.1, .2)", "0", "#5B21B6", "#EDE9FE")
        self.lbl_stat_check = self._create_stat_badge(cards_frame, "Cần kiểm tra", "0", "#991B1B", "#FEE2E2")

        # 4. THANH TIẾN TRÌNH (PROGRESS BAR)
        prog_frame = tk.Frame(self, bg="#F3F4F6")
        prog_frame.pack(fill="x", padx=12, pady=(0, 6))

        self.var_progress = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(prog_frame, variable=self.var_progress, maximum=100.0)
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.var_status_msg = tk.StringVar(value="Sẵn sàng.")
        tk.Label(prog_frame, textvariable=self.var_status_msg, bg="#F3F4F6", fg="#374151", font=("Segoe UI", 9, "italic")).pack(side="right")

        # 5. NOTEBOOK 3 TAB CHÍNH
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=(0, 10))

        self.tab_files = ttk.Frame(self.notebook)
        self.tab_catalog = ttk.Frame(self.notebook)
        self.tab_stats = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_files, text="  📋 1. Danh sách File PDF & Xem trước trực quan  ")
        self.notebook.add(self.tab_catalog, text="  📊 2. Quản lý Danh mục Excel đặt tên (Tải lên / Chỉnh sửa)  ")
        self.notebook.add(self.tab_stats, text="  📈 3. Thống kê theo Đơn vị & Báo cáo  ")

        self._build_tab_files()
        self._build_tab_catalog()
        self._build_tab_stats()

    def _create_stat_badge(self, parent, title, val, fg_color, bg_color):
        box = tk.Frame(parent, bg=bg_color, padx=10, pady=3, bd=1, relief="solid")
        box.pack(side="left", padx=3)
        tk.Label(box, text=f"{title}: ", bg=bg_color, fg=fg_color, font=("Segoe UI", 9)).pack(side="left")
        lbl_val = tk.Label(box, text=val, bg=bg_color, fg=fg_color, font=("Segoe UI", 10, "bold"))
        lbl_val.pack(side="left")
        return lbl_val

    # ==========================================================================
    # TAB 1: DANH SÁCH FILE PDF & XEM TRƯỚC TRỰC QUAN
    # ==========================================================================

    def _build_tab_files(self):
        # Thanh bộ lọc & tìm kiếm
        filter_bar = tk.Frame(self.tab_files, bg="#FFFFFF", bd=1, relief="solid")
        filter_bar.pack(fill="x", padx=6, pady=6)

        tk.Label(filter_bar, text="Lọc trạng thái:", bg="#FFFFFF", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(10, 4), pady=6)
        self.var_filter_status = tk.StringVar(value="ALL")
        cb_status = ttk.Combobox(
            filter_bar,
            textvariable=self.var_filter_status,
            state="readonly",
            width=24,
            values=[
                "ALL - Tất cả file",
                "EXCEL - Khớp danh mục Excel",
                "OUTSIDE - Ngoài danh mục Excel",
                "CHECK - Cần kiểm tra thủ công",
                "DUP - Có trùng lặp (.1, .2...)"
            ]
        )
        cb_status.current(0)
        cb_status.pack(side="left", padx=4)
        cb_status.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())

        tk.Label(filter_bar, text="Phòng ban:", bg="#FFFFFF", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(12, 4))
        self.var_filter_dept = tk.StringVar(value="Tất cả")
        self.cb_dept = ttk.Combobox(filter_bar, textvariable=self.var_filter_dept, state="readonly", width=22, values=["Tất cả"])
        self.cb_dept.pack(side="left", padx=4)
        self.cb_dept.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())

        tk.Label(filter_bar, text="🔎 Tìm kiếm:", bg="#FFFFFF", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(12, 4))
        self.var_search = tk.StringVar()
        ent_search = tk.Entry(filter_bar, textvariable=self.var_search, font=("Segoe UI", 9), width=28, bg="#F9FAFB", relief="solid", bd=1)
        ent_search.pack(side="left", padx=4, ipady=2)
        ent_search.bind("<Return>", lambda e: self._apply_filters())
        ttk.Button(filter_bar, text="Lọc", style="Secondary.TButton", command=self._apply_filters).pack(side="left", padx=3)
        ttk.Button(filter_bar, text="Xóa lọc", style="Secondary.TButton", command=self._reset_filters).pack(side="left", padx=3)

        # Điều khiển phân trang
        page_bar = tk.Frame(filter_bar, bg="#FFFFFF")
        page_bar.pack(side="right", padx=10)
        ttk.Button(page_bar, text="◀ Trước", style="Secondary.TButton", command=self._prev_page).pack(side="left", padx=2)
        self.lbl_page_info = tk.Label(page_bar, text="Trang 1/1 (0 dòng)", bg="#FFFFFF", font=("Segoe UI", 9, "bold"))
        self.lbl_page_info.pack(side="left", padx=6)
        ttk.Button(page_bar, text="Sau ▶", style="Secondary.TButton", command=self._next_page).pack(side="left", padx=2)

        # PanedWindow chia đôi: Trái = Bảng danh sách, Phải = Xem trước PDF & Chỉnh sửa tên
        paned = ttk.PanedWindow(self.tab_files, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        # --- Khung bên trái: Bảng Treeview ---
        left_frame = tk.Frame(paned, bg="#FFFFFF", bd=1, relief="solid")
        paned.add(left_frame, weight=3)

        cols = ("stt", "dept", "person", "old_name", "status", "confidence", "new_name", "reason")
        self.tree_files = ttk.Treeview(left_frame, columns=cols, show="headings", selectmode="browse")
        self.tree_files.heading("stt", text="STT")
        self.tree_files.heading("dept", text="Phòng ban")
        self.tree_files.heading("person", text="Họ tên hồ sơ")
        self.tree_files.heading("old_name", text="Tên file gốc")
        self.tree_files.heading("status", text="Phân loại")
        self.tree_files.heading("confidence", text="Độ tin cậy")
        self.tree_files.heading("new_name", text="Tên file mới dự kiến")
        self.tree_files.heading("reason", text="Lý do / Ghi chú đối chiếu")

        self.tree_files.column("stt", width=48, anchor="center", stretch=False)
        self.tree_files.column("dept", width=125, anchor="w", stretch=False)
        self.tree_files.column("person", width=140, anchor="w", stretch=False)
        self.tree_files.column("old_name", width=180, anchor="w")
        self.tree_files.column("status", width=95, anchor="center", stretch=False)
        self.tree_files.column("confidence", width=75, anchor="center", stretch=False)
        self.tree_files.column("new_name", width=260, anchor="w")
        self.tree_files.column("reason", width=220, anchor="w")

        self.tree_files.tag_configure("EXCEL", background="#F0FDF4", foreground="#065F46")
        self.tree_files.tag_configure("OUTSIDE", background="#EFF6FF", foreground="#1E40AF")
        self.tree_files.tag_configure("CHECK", background="#FEF2F2", foreground="#991B1B")

        vsb = ttk.Scrollbar(left_frame, orient="vertical", command=self.tree_files.yview)
        hsb = ttk.Scrollbar(left_frame, orient="horizontal", command=self.tree_files.xview)
        self.tree_files.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree_files.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        left_frame.grid_rowconfigure(0, weight=1)
        left_frame.grid_columnconfigure(0, weight=1)

        self.tree_files.bind("<<TreeviewSelect>>", self._on_select_file_row)
        self.tree_files.bind("<Double-1>", lambda e: self._open_selected_pdf_external())

        # --- Khung bên phải: Xem trước trang PDF & Chỉnh sửa tên trực tiếp ---
        right_frame = tk.Frame(paned, bg="#FFFFFF", bd=1, relief="solid", width=430)
        paned.add(right_frame, weight=2)

        # Khung gán/sửa tên nhanh
        edit_box = tk.LabelFrame(right_frame, text=" ✏️ Chỉnh sửa / Gán tên chuẩn cho file đang chọn ", bg="#FFFFFF", fg="#1E3A8A", font=("Segoe UI", 9, "bold"), padx=8, pady=6)
        edit_box.pack(fill="x", padx=8, pady=(8, 4))

        self.lbl_selected_info = tk.Label(edit_box, text="Chưa chọn file nào", bg="#FFFFFF", fg="#374151", font=("Segoe UI", 9), anchor="w", justify="left", wraplength=390)
        self.lbl_selected_info.pack(fill="x", pady=(0, 4))

        tk.Label(edit_box, text="Chọn nhanh từ danh mục Excel:", bg="#FFFFFF", font=("Segoe UI", 9)).pack(anchor="w")
        self.var_quick_catalog = tk.StringVar()
        self.cb_quick_catalog = ttk.Combobox(edit_box, textvariable=self.var_quick_catalog, font=("Segoe UI", 9))
        self.cb_quick_catalog.pack(fill="x", pady=2)
        self.cb_quick_catalog.bind("<<ComboboxSelected>>", self._on_pick_quick_catalog)
        self.cb_quick_catalog.bind("<KeyRelease>", self._on_filter_quick_catalog)

        tk.Label(edit_box, text="Tên file mới dự kiến (có thể gõ sửa trực tiếp):", bg="#FFFFFF", font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 0))
        edit_row = tk.Frame(edit_box, bg="#FFFFFF")
        edit_row.pack(fill="x", pady=2)
        self.var_edit_new_name = tk.StringVar()
        tk.Entry(edit_row, textvariable=self.var_edit_new_name, font=("Segoe UI", 9, "bold"), fg="#065F46", bg="#F0FDF4", relief="solid", bd=1).pack(side="left", fill="x", expand=True, ipady=3, padx=(0, 6))
        ttk.Button(edit_row, text="💾 Áp dụng tên này", style="Success.TButton", command=self._apply_manual_new_name).pack(side="right")

        # Nội dung OCR
        ocr_box = tk.LabelFrame(right_frame, text=" 📄 Văn bản OCR nhận diện được ", bg="#FFFFFF", fg="#374151", font=("Segoe UI", 9, "bold"), padx=6, pady=4)
        ocr_box.pack(fill="x", padx=8, pady=4)
        self.txt_ocr_preview = tk.Text(ocr_box, height=4, font=("Consolas", 8), bg="#F9FAFB", fg="#111827", wrap="word", relief="flat")
        self.txt_ocr_preview.pack(fill="x")

        # Khung hình ảnh trang PDF
        preview_header = tk.Frame(right_frame, bg="#FFFFFF")
        preview_header.pack(fill="x", padx=8, pady=(4, 2))
        tk.Label(preview_header, text="👁️ Xem trước trang PDF:", bg="#FFFFFF", font=("Segoe UI", 9, "bold")).pack(side="left")

        ttk.Button(preview_header, text="📂 Mở thư mục", style="Secondary.TButton", command=self._open_selected_folder).pack(side="right", padx=2)
        ttk.Button(preview_header, text="🔍 Mở PDF", style="Secondary.TButton", command=self._open_selected_pdf_external).pack(side="right", padx=2)
        ttk.Button(preview_header, text="▶", style="Secondary.TButton", width=3, command=self._next_pdf_page).pack(side="right", padx=1)
        self.lbl_pdf_page = tk.Label(preview_header, text="0/0", bg="#FFFFFF", font=("Segoe UI", 9, "bold"))
        self.lbl_pdf_page.pack(side="right", padx=4)
        ttk.Button(preview_header, text="◀", style="Secondary.TButton", width=3, command=self._prev_pdf_page).pack(side="right", padx=1)

        self.canvas_preview = tk.Canvas(right_frame, bg="#E5E7EB", highlightthickness=1, highlightbackground="#D1D5DB")
        self.canvas_preview.pack(fill="both", expand=True, padx=8, pady=(2, 8))
        self.canvas_preview.bind("<Configure>", lambda e: self._render_current_pdf_page())

    # ==========================================================================
    # TAB 2: QUẢN LÝ DANH MỤC EXCEL ĐẶT TÊN (UPLOAD / CHỈNH SỬA TRỰC TIẾP)
    # ==========================================================================

    def _build_tab_catalog(self):
        top_bar = tk.Frame(self.tab_catalog, bg="#FFFFFF", bd=1, relief="solid")
        top_bar.pack(fill="x", padx=6, pady=6)

        ttk.Button(
            top_bar,
            text="📤 Tải lên File Excel danh mục mới (.xlsx)...",
            style="Primary.TButton",
            command=self._upload_new_excel_catalog
        ).pack(side="left", padx=(10, 6), pady=8)

        ttk.Button(
            top_bar,
            text="➕ Thêm mục đặt tên mới",
            style="Success.TButton",
            command=self._add_catalog_item_dialog
        ).pack(side="left", padx=4, pady=8)

        ttk.Button(
            top_bar,
            text="✏️ Sửa mục đang chọn",
            style="Secondary.TButton",
            command=self._edit_selected_catalog_item
        ).pack(side="left", padx=4, pady=8)

        ttk.Button(
            top_bar,
            text="🗑️ Xóa mục đang chọn",
            style="Danger.TButton",
            command=self._delete_selected_catalog_item
        ).pack(side="left", padx=4, pady=8)

        ttk.Button(
            top_bar,
            text="📥 Xuất danh mục ra file Excel...",
            style="Secondary.TButton",
            command=self._export_template_excel
        ).pack(side="right", padx=10, pady=8)

        ttk.Button(
            top_bar,
            text="🔄 Áp dụng danh mục này cho toàn bộ kết quả Dry-Run",
            style="Warning.TButton",
            command=self._start_dry_run
        ).pack(side="right", padx=4, pady=8)

        # Hướng dẫn cấu trúc Excel
        help_banner = tk.Frame(self.tab_catalog, bg="#EFF6FF", bd=1, relief="solid")
        help_banner.pack(fill="x", padx=6, pady=(0, 6))
        tk.Label(
            help_banner,
            text=(
                "💡 Hướng dẫn tải lên file Excel đặt tên: Bạn có thể tải lên bất kỳ file .xlsx nào có 1, 2 hoặc 3 cột:\n"
                "   • Cột 1: Tên đặt File không dấu an toàn (VD: '01. Ly lich cua nguoi vao Dang') — nếu bạn chỉ có cột có dấu, phần mềm sẽ tự động bỏ dấu giúp bạn.\n"
                "   • Cột 2: Tên tài liệu có dấu gốc (VD: 'Lý lịch của người vào Đảng').\n"
                "   • Cột 3 (Tùy chọn): Từ khóa nhận diện bổ sung cách nhau bởi dấu phẩy (VD: 'mau 2 knd, ly lich nguoi xin vao dang')."
            ),
            bg="#EFF6FF",
            fg="#1E3A8A",
            font=("Segoe UI", 9),
            justify="left",
            anchor="w",
            padx=10,
            pady=6
        ).pack(fill="x")

        # Bảng danh mục Excel
        cat_frame = tk.Frame(self.tab_catalog, bg="#FFFFFF", bd=1, relief="solid")
        cat_frame.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        c_cols = ("id", "safe_name", "orig_name", "norm_title", "keywords")
        self.tree_catalog = ttk.Treeview(cat_frame, columns=c_cols, show="headings", selectmode="browse")
        self.tree_catalog.heading("id", text="Mã / STT")
        self.tree_catalog.heading("safe_name", text="Tên đặt File/Folder chuẩn (Không dấu)")
        self.tree_catalog.heading("orig_name", text="Tên tài liệu (Có dấu gốc)")
        self.tree_catalog.heading("norm_title", text="Chuỗi chuẩn hóa đối chiếu OCR")
        self.tree_catalog.heading("keywords", text="Từ khóa nhận diện bổ sung (Tùy chọn)")

        self.tree_catalog.column("id", width=80, anchor="center", stretch=False)
        self.tree_catalog.column("safe_name", width=360, anchor="w")
        self.tree_catalog.column("orig_name", width=360, anchor="w")
        self.tree_catalog.column("norm_title", width=260, anchor="w")
        self.tree_catalog.column("keywords", width=240, anchor="w")

        vsb_c = ttk.Scrollbar(cat_frame, orient="vertical", command=self.tree_catalog.yview)
        self.tree_catalog.configure(yscrollcommand=vsb_c.set)
        self.tree_catalog.pack(side="left", fill="both", expand=True)
        vsb_c.pack(side="right", fill="y")

        self.tree_catalog.bind("<Double-1>", lambda e: self._edit_selected_catalog_item())

    # ==========================================================================
    # TAB 3: THỐNG KÊ THEO PHÒNG BAN & BÁO CÁO
    # ==========================================================================

    def _build_tab_stats(self):
        top_btns = tk.Frame(self.tab_stats, bg="#FFFFFF", bd=1, relief="solid")
        top_btns.pack(fill="x", padx=6, pady=6)

        tk.Label(top_btns, text="Mở nhanh các file báo cáo Excel xuất ra:", bg="#FFFFFF", font=("Segoe UI", 9, "bold")).pack(side="left", padx=12, pady=8)
        ttk.Button(top_btns, text="📗 Danh_sach_du_kien_doi_ten.xlsx", style="Secondary.TButton", command=lambda: self._open_report_file("Danh_sach_du_kien_doi_ten.xlsx")).pack(side="left", padx=4)
        ttk.Button(top_btns, text="📕 Danh_sach_can_kiem_tra.xlsx", style="Secondary.TButton", command=lambda: self._open_report_file("Danh_sach_can_kiem_tra.xlsx")).pack(side="left", padx=4)
        ttk.Button(top_btns, text="📘 Thong_ke.xlsx", style="Secondary.TButton", command=lambda: self._open_report_file("Thong_ke.xlsx")).pack(side="left", padx=4)
        ttk.Button(top_btns, text="📙 Nhat_ky_doi_ten.xlsx", style="Secondary.TButton", command=lambda: self._open_report_file("Nhat_ky_doi_ten.xlsx")).pack(side="left", padx=4)

        stats_frame = tk.Frame(self.tab_stats, bg="#FFFFFF", bd=1, relief="solid")
        stats_frame.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        s_cols = ("dept", "total", "recognized", "excel", "outside", "dup", "check", "rate")
        self.tree_stats = ttk.Treeview(stats_frame, columns=s_cols, show="headings")
        self.tree_stats.heading("dept", text="Phòng ban / Đơn vị")
        self.tree_stats.heading("total", text="Tổng số PDF")
        self.tree_stats.heading("recognized", text="Tổng nhận diện được")
        self.tree_stats.heading("excel", text="Khớp trong Excel")
        self.tree_stats.heading("outside", text="Ngoài Excel (Theo tiêu đề)")
        self.tree_stats.heading("dup", text="Trùng lặp (.1, .2...)")
        self.tree_stats.heading("check", text="Cần kiểm tra")
        self.tree_stats.heading("rate", text="Tỷ lệ nhận diện (%)")

        self.tree_stats.column("dept", width=260, anchor="w")
        for c in ("total", "recognized", "excel", "outside", "dup", "check", "rate"):
            self.tree_stats.column(c, width=140, anchor="center")

        self.tree_stats.pack(fill="both", expand=True)

    # ==========================================================================
    # XỬ LÝ DỮ LIỆU DANH MỤC EXCEL (UPLOAD / THÊM / SỬA / XÓA)
    # ==========================================================================

    def _refresh_catalog_ui(self):
        self.lbl_catalog_badge.config(text=f"{len(self.catalog):,} mục đặt tên")
        for item in self.tree_catalog.get_children():
            self.tree_catalog.delete(item)

        combo_values = []
        for cid in sorted(self.catalog.keys()):
            it = self.catalog[cid]
            self.tree_catalog.insert(
                "",
                "end",
                iid=str(cid),
                values=(
                    cid,
                    it.get('safe_name', ''),
                    it.get('orig_name', ''),
                    it.get('norm_title', ''),
                    it.get('keywords_raw', '')
                )
            )
            combo_values.append(it.get('safe_name', ''))

        self._all_catalog_names = combo_values
        self.cb_quick_catalog['values'] = combo_values

    def _upload_new_excel_catalog(self):
        filepath = filedialog.askopenfilename(
            title="Chọn file Excel (.xlsx) danh mục đặt tên mới",
            filetypes=[("Excel Files (*.xlsx)", "*.xlsx"), ("All Files (*.*)", "*.*")]
        )
        if not filepath:
            return

        src_path = pathlib.Path(filepath)
        try:
            new_cat = ce.load_excel_catalog(src_path)
            if not new_cat:
                messagebox.showerror("Lỗi dữ liệu Excel", "File Excel bạn chọn không có dòng dữ liệu hợp lệ nào!")
                return

            # Sao chép vào thư mục data của project để lưu trữ lâu dài
            ce.DATA_DIR.mkdir(parents=True, exist_ok=True)
            dest_path = ce.DATA_DIR / src_path.name
            if src_path.resolve() != dest_path.resolve():
                shutil.copy2(src_path, dest_path)

            # Cập nhật chuẩn hóa vào cả file mặc định và config
            self.cfg['excel_catalog_path'] = str(dest_path)
            ce.save_config(self.cfg)
            self.var_excel_path.set(str(dest_path))

            self.catalog = ce.reload_global_catalog(dest_path)
            self._refresh_catalog_ui()

            ask_reapply = messagebox.askyesno(
                "Tải lên Danh mục Excel thành công",
                f"Đã tải lên thành công file:\n{src_path.name}\n\n"
                f"Tổng số mục đặt tên nạp được: {len(self.catalog):,} mục.\n\n"
                f"Bạn có muốn chạy đối chiếu lại ngay lập tức với danh mục Excel mới này không?\n"
                f"(Nhờ bộ nhớ đệm OCR, việc cập nhật lại tên mới chỉ mất vài giây!)"
            )
            if ask_reapply:
                self._start_dry_run()

        except Exception as e:
            messagebox.showerror("Lỗi đọc file Excel", f"Không thể đọc file Excel:\n{e}")

    def _export_template_excel(self):
        save_path = filedialog.asksaveasfilename(
            title="Lưu file Excel danh mục đặt tên",
            defaultextension=".xlsx",
            initialfile="Danh_muc_tai_lieu_dat_ten.xlsx",
            filetypes=[("Excel Files (*.xlsx)", "*.xlsx")]
        )
        if not save_path:
            return
        try:
            ce.save_catalog_to_xlsx(self.catalog, pathlib.Path(save_path))
            messagebox.showinfo("Xuất file thành công", f"Đã xuất danh mục ({len(self.catalog)} mục) ra file:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Lỗi xuất Excel", str(e))

    def _add_catalog_item_dialog(self):
        self._open_catalog_editor_modal(mode="add")

    def _edit_selected_catalog_item(self):
        sel = self.tree_catalog.selection()
        if not sel:
            messagebox.showwarning("Chưa chọn mục", "Vui lòng chọn 1 dòng trong bảng danh mục để sửa.")
            return
        cid = int(sel[0])
        self._open_catalog_editor_modal(mode="edit", cid=cid)

    def _delete_selected_catalog_item(self):
        sel = self.tree_catalog.selection()
        if not sel:
            messagebox.showwarning("Chưa chọn mục", "Vui lòng chọn 1 dòng trong bảng danh mục để xóa.")
            return
        cid = int(sel[0])
        it = self.catalog.get(cid)
        if not it:
            return
        if messagebox.askyesno("Xác nhận xóa", f"Bạn có chắc muốn xóa mục [{cid}] '{it['safe_name']}' khỏi danh mục?"):
            del self.catalog[cid]
            ce.save_catalog_to_xlsx(self.catalog, pathlib.Path(self.var_excel_path.get()))
            self._refresh_catalog_ui()

    def _open_catalog_editor_modal(self, mode="add", cid=None):
        win = tk.Toplevel(self)
        win.title("Thêm mục đặt tên mới" if mode == "add" else f"Chỉnh sửa mục #{cid}")
        win.geometry("560x320")
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()
        win.configure(bg="#FFFFFF")

        item = self.catalog.get(cid, {}) if mode == "edit" else {}
        next_id = (max(self.catalog.keys()) + 1) if self.catalog else 1

        frm = tk.Frame(win, bg="#FFFFFF", padx=18, pady=14)
        frm.pack(fill="both", expand=True)

        tk.Label(frm, text="Mã / Số thứ tự:", bg="#FFFFFF", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", pady=6)
        var_id = tk.StringVar(value=str(cid if mode == "edit" else next_id))
        ent_id = tk.Entry(frm, textvariable=var_id, font=("Segoe UI", 9), width=15, state="readonly" if mode == "edit" else "normal")
        ent_id.grid(row=0, column=1, sticky="w", pady=6)

        tk.Label(frm, text="Tên tài liệu (Có dấu gốc):", bg="#FFFFFF", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w", pady=6)
        var_orig = tk.StringVar(value=item.get('orig_name', ''))
        ent_orig = tk.Entry(frm, textvariable=var_orig, font=("Segoe UI", 9), width=48)
        ent_orig.grid(row=1, column=1, sticky="ew", pady=6)

        tk.Label(frm, text="Tên đặt File (Không dấu):", bg="#FFFFFF", font=("Segoe UI", 9, "bold")).grid(row=2, column=0, sticky="w", pady=6)
        var_safe = tk.StringVar(value=item.get('safe_name', f"{next_id:02d}. " if mode == "add" else ""))
        ent_safe = tk.Entry(frm, textvariable=var_safe, font=("Segoe UI", 9, "bold"), fg="#1E40AF", width=48)
        ent_safe.grid(row=2, column=1, sticky="ew", pady=6)

        def auto_fill_safe(*_):
            if mode == "add":
                raw_o = var_orig.get().strip()
                num_s = var_id.get().strip()
                prefix = f"{int(num_s):02d}. " if num_s.isdigit() else ""
                if raw_o:
                    var_safe.set(ce.sanitize_windows_filename(prefix + ce.remove_accents(raw_o)))

        ent_orig.bind("<KeyRelease>", auto_fill_safe)

        tk.Label(frm, text="Từ khóa nhận diện (Tùy chọn):", bg="#FFFFFF", font=("Segoe UI", 9, "bold")).grid(row=3, column=0, sticky="w", pady=6)
        var_kw = tk.StringVar(value=item.get('keywords_raw', ''))
        tk.Entry(frm, textvariable=var_kw, font=("Segoe UI", 9), width=48).grid(row=3, column=1, sticky="ew", pady=6)
        tk.Label(frm, text="(Các từ khóa cách nhau bởi dấu phẩy, VD: mau 2 knd, don xin)", bg="#FFFFFF", fg="#6B7280", font=("Segoe UI", 8)).grid(row=4, column=1, sticky="w")

        def save_modal():
            try:
                new_cid = int(var_id.get().strip())
            except ValueError:
                messagebox.showerror("Lỗi", "Mã / Số thứ tự phải là số nguyên!", parent=win)
                return
            safe_s = ce.sanitize_windows_filename(var_safe.get().strip())
            orig_s = var_orig.get().strip() or safe_s
            kw_s = var_kw.get().strip()
            if not safe_s:
                messagebox.showerror("Lỗi", "Tên đặt File không được để trống!", parent=win)
                return

            clean_title = ce.re.sub(r'^\d+\.\s*', '', safe_s)
            self.catalog[new_cid] = {
                'id': new_cid,
                'safe_name': safe_s,
                'orig_name': orig_s,
                'keywords': [ce.norm_text(k) for k in ce.re.split(r'[,;|]', kw_s) if ce.norm_text(k)],
                'keywords_raw': kw_s,
                'norm_title': ce.norm_text(clean_title),
                'compact_title': ce.norm_text(clean_title).replace(' ', '')
            }
            ce.CATALOG = self.catalog
            ce.save_catalog_to_xlsx(self.catalog, pathlib.Path(self.var_excel_path.get()))
            self._refresh_catalog_ui()
            win.destroy()

        btn_box = tk.Frame(frm, bg="#FFFFFF")
        btn_box.grid(row=5, column=0, columnspan=2, pady=16)
        ttk.Button(btn_box, text="💾 Lưu vào Danh mục Excel", style="Success.TButton", command=save_modal).pack(side="left", padx=6)
        ttk.Button(btn_box, text="Hủy", style="Secondary.TButton", command=win.destroy).pack(side="left", padx=6)

    # ==========================================================================
    # QUÉT OCR & DỰ KIẾN ĐỔI TÊN (DRY-RUN) + BỘ LỌC BẢNG
    # ==========================================================================

    def _load_initial_records(self):
        report_dir = pathlib.Path(self.cfg.get('report_dir', str(ce._DEFAULT_REPORT_DIR)))
        root_dir = pathlib.Path(self.var_root_dir.get())
        records = ce.load_saved_ui_records(report_dir, root_dir)
        if records:
            self.all_records = records
            self._update_summary_badges_and_stats()
            self._populate_dept_filter()
            self._apply_filters()
            self.var_status_msg.set(f"Đã tải {len(records):,} kết quả từ lần chạy gần nhất.")

    def _choose_root_dir(self):
        initial_dir = pathlib.Path(self.var_root_dir.get())
        if not initial_dir.is_dir():
            initial_dir = pathlib.Path.home()
        d = filedialog.askdirectory(title="Chọn thư mục gốc chứa hồ sơ PDF", initialdir=str(initial_dir))
        if d:
            self.var_root_dir.set(d)
            self._save_current_config()

    def _save_current_config(self):
        self.cfg['root_dir'] = self.var_root_dir.get().strip()
        self.cfg['excel_catalog_path'] = self.var_excel_path.get().strip()
        self.cfg['workers'] = int(self.var_workers.get())
        self.cfg['allow_outside_excel'] = bool(self.var_allow_outside.get())
        ce.save_config(self.cfg)

    def _start_dry_run(self):
        if self.worker_thread and self.worker_thread.is_alive():
            return
        self._save_current_config()
        root_dir = pathlib.Path(self.var_root_dir.get().strip())
        excel_path = pathlib.Path(self.var_excel_path.get().strip())
        report_dir = pathlib.Path(self.cfg.get('report_dir', str(root_dir / '_BAO_CAO_DOI_TEN')))

        if not root_dir.exists():
            messagebox.showerror("Lỗi đường dẫn", f"Thư mục hồ sơ không tồn tại:\n{root_dir}")
            return
        if not excel_path.exists():
            messagebox.showerror("Lỗi file Excel", f"File Excel danh mục không tồn tại:\n{excel_path}")
            return

        tess_dir = pathlib.Path(self.cfg.get('tess_dir', ''))
        tessdata_dir = pathlib.Path(self.cfg.get('tessdata_dir', ''))
        if not (tess_dir / 'libtesseract-5.dll').is_file() or not (tessdata_dir / 'vie.traineddata').is_file():
            messagebox.showerror(
                "Chưa cấu hình OCR",
                "Không tìm thấy Tesseract hoặc dữ liệu tiếng Việt. Hãy đặt TESSERACT_DIR tới thư mục Tesseract rồi khởi động lại ứng dụng."
            )
            return

        self.stop_event.clear()
        self.btn_dry_run.config(state="disabled")
        self.btn_execute.config(state="disabled")
        self.btn_rollback.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.var_progress.set(2.0)
        self.var_status_msg.set("Đang khởi động quét & đối chiếu...")

        def progress_cb(stage, current, total, msg):
            pct = (current * 100.0 / total) if total > 0 else 0.0
            self.after(0, lambda: self._update_progress_ui(pct, msg))

        def run_task():
            try:
                res = ce.run_dry_run_engine(
                    root_dir=root_dir,
                    excel_path=excel_path,
                    report_dir=report_dir,
                    tess_dir=self.cfg.get('tess_dir', ce.DEFAULT_CONFIG['tess_dir']),
                    tessdata_dir=self.cfg.get('tessdata_dir', ce.DEFAULT_CONFIG['tessdata_dir']),
                    workers=int(self.var_workers.get()),
                    allow_outside=bool(self.var_allow_outside.get()),
                    progress_callback=progress_cb,
                    stop_event=self.stop_event
                )
                self.after(0, lambda: self._on_dry_run_finished(res, None))
            except Exception as e:
                self.after(0, lambda: self._on_dry_run_finished(None, e))

        self.worker_thread = threading.Thread(target=run_task, daemon=True)
        self.worker_thread.start()

    def _stop_dry_run(self):
        self.stop_event.set()
        self.var_status_msg.set("Đang dừng tiến trình...")

    def _update_progress_ui(self, pct, msg):
        self.var_progress.set(pct)
        self.var_status_msg.set(msg)

    def _on_dry_run_finished(self, res, err):
        self.btn_dry_run.config(state="normal")
        self.btn_execute.config(state="normal")
        self.btn_rollback.config(state="normal")
        self.btn_stop.config(state="disabled")

        if err:
            messagebox.showerror("Lỗi trong quá trình quét", str(err))
            self.var_status_msg.set(f"Lỗi: {err}")
            return

        if not res or res.get('cancelled'):
            self.var_status_msg.set("Đã dừng quá trình quét theo yêu cầu.")
            return

        self.all_records = res.get('records', [])
        self._update_summary_badges_and_stats()
        self._populate_dept_filter()
        self._apply_filters()
        self.notebook.select(self.tab_files)

        messagebox.showinfo(
            "Hoàn tất Dự kiến đổi tên (Dry-Run)",
            f"Đã hoàn tất quét & đối chiếu trong {res['elapsed']:.1f} giây!\n\n"
            f"• Tổng số file PDF: {res['total_pdfs']:,}\n"
            f"• Khớp danh mục Excel: {res['excel_matched']:,}\n"
            f"• Ngoài danh mục Excel: {res['outside_excel']:,}\n"
            f"• Trùng lặp đánh số (.1, .2): {res['duplicates']:,}\n"
            f"• Cần kiểm tra: {res['need_check']:,}\n\n"
            f"Báo cáo Excel đã được xuất tại: {res['report_dir']}"
        )

    def _populate_dept_filter(self):
        depts = sorted({r['dept'] for r in self.all_records if r.get('dept')})
        self.cb_dept['values'] = ["Tất cả"] + depts
        if self.var_filter_dept.get() not in self.cb_dept['values']:
            self.var_filter_dept.set("Tất cả")

    def _update_summary_badges_and_stats(self):
        total = len(self.all_records)
        excel_c = sum(1 for r in self.all_records if r['status'] == 'EXCEL' and r['new_name'] != "(CẦN KIỂM TRA - CHƯA ĐỔI TÊN)")
        out_c = sum(1 for r in self.all_records if r['status'] == 'OUTSIDE' and r['new_name'] != "(CẦN KIỂM TRA - CHƯA ĐỔI TÊN)")
        dup_c = sum(1 for r in self.all_records if r.get('is_dup') == 'Có')
        chk_c = total - excel_c - out_c

        self.lbl_stat_total.config(text=f"{total:,}")
        self.lbl_stat_excel.config(text=f"{excel_c:,}")
        self.lbl_stat_outside.config(text=f"{out_c:,}")
        self.lbl_stat_dup.config(text=f"{dup_c:,}")
        self.lbl_stat_check.config(text=f"{chk_c:,}")

        # Cập nhật bảng thống kê theo phòng ban ở Tab 3
        for item in self.tree_stats.get_children():
            self.tree_stats.delete(item)

        dept_map = ce.defaultdict(lambda: {'total': 0, 'excel': 0, 'outside': 0, 'dup': 0, 'check': 0})
        for r in self.all_records:
            d = r['dept']
            dept_map[d]['total'] += 1
            if r.get('is_dup') == 'Có':
                dept_map[d]['dup'] += 1
            if r['status'] == 'EXCEL' and r['new_name'] != "(CẦN KIỂM TRA - CHƯA ĐỔI TÊN)":
                dept_map[d]['excel'] += 1
            elif r['status'] == 'OUTSIDE' and r['new_name'] != "(CẦN KIỂM TRA - CHƯA ĐỔI TÊN)":
                dept_map[d]['outside'] += 1
            else:
                dept_map[d]['check'] += 1

        for d_name, st in sorted(dept_map.items()):
            rec_cnt = st['excel'] + st['outside']
            rate = (rec_cnt * 100.0 / st['total']) if st['total'] > 0 else 0.0
            self.tree_stats.insert(
                "",
                "end",
                values=(
                    d_name,
                    f"{st['total']:,}",
                    f"{rec_cnt:,}",
                    f"{st['excel']:,}",
                    f"{st['outside']:,}",
                    f"{st['dup']:,}",
                    f"{st['check']:,}",
                    f"{rate:.1f}%"
                )
            )

    def _reset_filters(self):
        self.var_filter_status.set("ALL - Tất cả file")
        self.var_filter_dept.set("Tất cả")
        self.var_search.set("")
        self._apply_filters()

    def _apply_filters(self):
        st_raw = self.var_filter_status.get().split(" - ")[0].strip()
        dept_f = self.var_filter_dept.get().strip()
        q = ce.norm_text(self.var_search.get())

        filtered = []
        for idx, r in enumerate(self.all_records):
            if st_raw == "EXCEL" and r['status'] != 'EXCEL':
                continue
            elif st_raw == "OUTSIDE" and r['status'] != 'OUTSIDE':
                continue
            elif st_raw == "CHECK" and r['status'] != 'CHECK':
                continue
            elif st_raw == "DUP" and r.get('is_dup') != 'Có':
                continue

            if dept_f and dept_f != "Tất cả" and r['dept'] != dept_f:
                continue

            if q:
                haystack = ce.norm_text(f"{r['person']} {r['old_name']} {r['new_name']} {r['reason']} {r.get('ocr_text', '')}")
                if q not in haystack:
                    continue

            filtered.append((idx, r))

        self.filtered_records = filtered
        self.current_page = 1
        self._render_files_page()

    def _render_files_page(self):
        for item in self.tree_files.get_children():
            self.tree_files.delete(item)

        total_items = len(self.filtered_records)
        total_pages = max(1, (total_items + self.page_size - 1) // self.page_size)
        if self.current_page > total_pages:
            self.current_page = total_pages

        start_i = (self.current_page - 1) * self.page_size
        end_i = min(total_items, start_i + self.page_size)

        self.lbl_page_info.config(text=f"Trang {self.current_page}/{total_pages} ({total_items:,} file)")

        status_label_map = {
            'EXCEL': 'Khớp Excel',
            'OUTSIDE': 'Ngoài Excel',
            'CHECK': 'Cần kiểm tra'
        }

        for display_num, (orig_idx, r) in enumerate(self.filtered_records[start_i:end_i], start=start_i + 1):
            st_code = r.get('status', 'CHECK')
            self.tree_files.insert(
                "",
                "end",
                iid=str(orig_idx),
                tags=(st_code,),
                values=(
                    display_num,
                    r.get('dept', ''),
                    r.get('person', ''),
                    r.get('old_name', ''),
                    status_label_map.get(st_code, st_code),
                    f"{r.get('confidence', 0):.0f}%",
                    r.get('new_name', ''),
                    r.get('reason', '')
                )
            )

    def _prev_page(self):
        if self.current_page > 1:
            self.current_page -= 1
            self._render_files_page()

    def _next_page(self):
        total_pages = max(1, (len(self.filtered_records) + self.page_size - 1) // self.page_size)
        if self.current_page < total_pages:
            self.current_page += 1
            self._render_files_page()

    # ==========================================================================
    # XEM TRƯỚC TRANG PDF & CHỈNH SỬA TÊN FILE TRỰC TIẾP
    # ==========================================================================

    def _on_select_file_row(self, _event=None):
        sel = self.tree_files.selection()
        if not sel:
            return
        orig_idx = int(sel[0])
        if not (0 <= orig_idx < len(self.all_records)):
            return

        rec = self.all_records[orig_idx]
        self.selected_record = (orig_idx, rec)

        self.lbl_selected_info.config(
            text=f"👤 {rec['person']}  |  📂 {rec['rel_dir']}\n📄 File gốc: {rec['old_name']}"
        )
        new_n = rec.get('new_name', '')
        self.var_edit_new_name.set("" if new_n == "(CẦN KIỂM TRA - CHƯA ĐỔI TÊN)" else new_n)

        self.txt_ocr_preview.delete("1.0", "end")
        self.txt_ocr_preview.insert("1.0", rec.get('ocr_text', '(Không có dữ liệu OCR)'))

        # Tìm đường dẫn file PDF thực tế trên đĩa
        root_dir = pathlib.Path(self.var_root_dir.get().strip())
        candidate_old = root_dir / rec['rel_dir'] / rec['old_name']
        candidate_new = root_dir / rec['rel_dir'] / rec['new_name'] if rec.get('new_name') else candidate_old
        if candidate_old.exists():
            self.preview_pdf_path = candidate_old
        elif candidate_new.exists():
            self.preview_pdf_path = candidate_new
        else:
            self.preview_pdf_path = candidate_old

        self.preview_page_idx = 0
        self._render_current_pdf_page()

    def _render_current_pdf_page(self):
        self.canvas_preview.delete("all")
        if not self.preview_pdf_path or not self.preview_pdf_path.exists():
            self.lbl_pdf_page.config(text="0/0")
            self.canvas_preview.create_text(
                200, 160,
                text="Không tìm thấy file PDF trên đĩa\n(Có thể file đã được đổi tên hoặc di chuyển)",
                fill="#6B7280",
                font=("Segoe UI", 9),
                justify="center"
            )
            return

        try:
            doc = pymupdf.open(self.preview_pdf_path)
            self.preview_total_pages = len(doc)
            if self.preview_total_pages == 0:
                doc.close()
                return
            if self.preview_page_idx >= self.preview_total_pages:
                self.preview_page_idx = 0

            self.lbl_pdf_page.config(text=f"{self.preview_page_idx + 1}/{self.preview_total_pages}")
            page = doc[self.preview_page_idx]
            pix = page.get_pixmap(dpi=110)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc.close()

            cw = max(240, self.canvas_preview.winfo_width() - 16)
            ch = max(260, self.canvas_preview.winfo_height() - 16)
            img.thumbnail((cw, ch), Image.Resampling.LANCZOS)

            self.preview_photo = ImageTk.PhotoImage(img)
            self.canvas_preview.create_image(cw // 2 + 8, ch // 2 + 8, image=self.preview_photo, anchor="center")
        except Exception as e:
            self.canvas_preview.create_text(180, 120, text=f"Lỗi đọc PDF: {e}", fill="#DC2626")

    def _prev_pdf_page(self):
        if self.preview_total_pages > 0 and self.preview_page_idx > 0:
            self.preview_page_idx -= 1
            self._render_current_pdf_page()

    def _next_pdf_page(self):
        if self.preview_total_pages > 0 and self.preview_page_idx + 1 < self.preview_total_pages:
            self.preview_page_idx += 1
            self._render_current_pdf_page()

    def _on_pick_quick_catalog(self, _event=None):
        chosen = self.var_quick_catalog.get().strip()
        if chosen:
            if not chosen.lower().endswith('.pdf'):
                chosen += ".pdf"
            self.var_edit_new_name.set(chosen)

    def _on_filter_quick_catalog(self, _event=None):
        q = ce.norm_text(self.var_quick_catalog.get())
        if not q:
            self.cb_quick_catalog['values'] = getattr(self, '_all_catalog_names', [])
        else:
            matches = [n for n in getattr(self, '_all_catalog_names', []) if q in ce.norm_text(n)]
            self.cb_quick_catalog['values'] = matches

    def _apply_manual_new_name(self):
        if not self.selected_record:
            messagebox.showwarning("Chưa chọn file", "Vui lòng chọn 1 file PDF trong danh sách bên trái trước.")
            return
        raw_new = self.var_edit_new_name.get().strip()
        if not raw_new:
            messagebox.showwarning("Tên trống", "Vui lòng nhập hoặc chọn tên file mới.")
            return

        if not raw_new.lower().endswith('.pdf'):
            raw_new += ".pdf"
        safe_new = ce.sanitize_windows_filename(raw_new[:-4]) + ".pdf"

        orig_idx, rec = self.selected_record
        rec['new_name'] = safe_new
        rec['status'] = 'EXCEL'
        rec['confidence'] = 100.0
        rec['reason'] = "Người dùng gán/chỉnh sửa trực tiếp trên giao diện"

        # Lưu lại vào kế hoạch rename_plan.json và báo cáo Excel
        report_dir = pathlib.Path(self.cfg.get('report_dir', str(ce._DEFAULT_REPORT_DIR)))
        ce.save_reports_and_plan(report_dir, self.all_records)
        self._update_summary_badges_and_stats()
        self._render_files_page()
        self.tree_files.selection_set(str(orig_idx))
        self.var_status_msg.set(f"Đã cập nhật tên mới: {safe_new}")

    # ==========================================================================
    # THỰC THI ĐỔI TÊN THẬT & HOÀN TÁC (ROLLBACK)
    # ==========================================================================

    def _execute_rename(self):
        root_dir = pathlib.Path(self.var_root_dir.get().strip())
        report_dir = pathlib.Path(self.cfg.get('report_dir', str(root_dir / '_BAO_CAO_DOI_TEN')))
        plan_file = report_dir / 'rename_plan.json'
        if not plan_file.exists():
            messagebox.showwarning("Chưa có kế hoạch", "Vui lòng bấm '1. Quét & Dự kiến đổi tên (Dry-Run)' trước khi thực thi đổi tên thật!")
            return

        ready_cnt = sum(1 for r in self.all_records if r.get('new_name') and r['new_name'] != "(CẦN KIỂM TRA - CHƯA ĐỔI TÊN)" and r['old_name'] != r['new_name'])
        if not messagebox.askyesno(
            "Xác nhận Thực thi đổi tên thật",
            f"Hệ thống chuẩn bị đổi tên an toàn (2 bước) cho {ready_cnt:,} file PDF trong:\n{root_dir}\n\n"
            f"• Các file nằm trong danh sách 'Cần kiểm tra' sẽ được giữ nguyên tên cũ.\n"
            f"• Nhật ký đổi tên được lưu đầy đủ để có thể Hoàn tác (Rollback) bất cứ lúc nào.\n\n"
            f"Bạn có chắc chắn muốn thực hiện ngay không?"
        ):
            return

        try:
            res = ce.execute_rename_engine(root_dir, report_dir)
            messagebox.showinfo(
                "Đổi tên hoàn tất",
                f"Đã thực thi đổi tên thành công!\n\n"
                f"• Số file đã đổi sang tên mới: {res['renamed_count']:,} file\n"
                f"• Số file đã đúng tên chuẩn từ trước: {res['already_named_count']:,} file\n"
                f"• Số file bỏ qua (tránh xung đột): {res['skipped_count']:,} file\n\n"
                f"Nhật ký đã lưu tại: {report_dir / 'Nhat_ky_doi_ten.xlsx'}"
            )
            self.var_status_msg.set(f"Đã đổi tên thành công {res['renamed_count']:,} file.")
        except Exception as e:
            messagebox.showerror("Lỗi thực thi đổi tên", str(e))

    def _rollback_rename(self):
        root_dir = pathlib.Path(self.var_root_dir.get().strip())
        report_dir = pathlib.Path(self.cfg.get('report_dir', str(root_dir / '_BAO_CAO_DOI_TEN')))
        if not messagebox.askyesno(
            "Xác nhận Hoàn tác (Rollback)",
            "Bạn có muốn khôi phục lại tên file ban đầu cho các file đã đổi tên gần nhất từ nhật ký không?"
        ):
            return
        try:
            res = ce.rollback_rename_engine(report_dir)
            messagebox.showinfo(
                "Hoàn tác thành công",
                f"Đã khôi phục thành công {res['restored']:,}/{res['total']:,} file về tên gốc ban đầu!"
            )
            self.var_status_msg.set(f"Đã hoàn tác {res['restored']:,} file về tên cũ.")
        except Exception as e:
            messagebox.showerror("Lỗi hoàn tác", str(e))

    # ==========================================================================
    # TIỆN ÍCH MỞ FILE & THƯ MỤC
    # ==========================================================================

    def _open_selected_pdf_external(self):
        if self.preview_pdf_path and self.preview_pdf_path.exists():
            os.startfile(str(self.preview_pdf_path))
        else:
            messagebox.showwarning("Không tìm thấy file", "File PDF không tồn tại ở đường dẫn hiện tại.")

    def _open_selected_folder(self):
        if self.preview_pdf_path and self.preview_pdf_path.parent.exists():
            os.startfile(str(self.preview_pdf_path.parent))

    def _open_report_folder(self):
        report_dir = pathlib.Path(self.cfg.get('report_dir', str(ce._DEFAULT_REPORT_DIR)))
        report_dir.mkdir(parents=True, exist_ok=True)
        os.startfile(str(report_dir))

    def _open_report_file(self, filename: str):
        report_dir = pathlib.Path(self.cfg.get('report_dir', str(ce._DEFAULT_REPORT_DIR)))
        f = report_dir / filename
        if f.exists():
            os.startfile(str(f))
        else:
            messagebox.showwarning("Chưa có báo cáo", f"File báo cáo '{filename}' chưa tồn tại. Hãy chạy Dry-Run hoặc Thực thi trước.")


if __name__ == '__main__':
    import multiprocessing
    multiprocessing.freeze_support()
    app = PDFRenameApp()
    app.mainloop()
