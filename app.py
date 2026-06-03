import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from PIL import Image, ImageTk, ImageDraw
import requests

# 程式版本與更新設定
VERSION = "V1.0.0"
# 請自行替換為實際的 GitHub repository API URL
UPDATE_URL = "https://api.github.com/repos/ChenYuChunEric/RenameandCut_Pic/releases/latest"

# 匯入核心處理器
from processor import PhotoProcessor

# 全域未捕獲異常處理，防止默默閃退
def show_exception_and_exit(exc_type, exc_value, exc_tb):
    err_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    try:
        with open("error_log.txt", "w", encoding="utf-8") as f:
            f.write(err_msg)
    except Exception:
        pass
    
    # 彈出視窗
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "程式發生嚴重錯誤 (直接閃退防護)", 
        f"程式因未預期錯誤而崩潰，即將關閉。\n\n詳細錯誤資訊已被寫入至同目錄下的 error_log.txt。\n\n詳細資訊：\n{exc_value}"
    )
    sys.exit(1)

sys.excepthook = show_exception_and_exit

# 設定 CustomTkinter 外觀
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class PhotoProcessorApp(ctk.CTk):
    def report_callback_exception(self, exc, val, tb):
        """
        捕獲 Tkinter 事件循環中的所有回呼異常，防止閃退，並彈出錯誤視窗。
        """
        err_msg = "".join(traceback.format_exception(exc, val, tb))
        try:
            with open("error_log.txt", "a", encoding="utf-8") as f:
                f.write(f"\n\n--- Tkinter Callback Error ({sys.platform}) ---\n{err_msg}")
        except Exception:
            pass
        messagebox.showerror(
            "執行錯誤", 
            f"執行操作時發生錯誤，程式已攔截以防止閃退。\n\n錯誤訊息：\n{str(val)}\n\n詳細錯誤追蹤已被記錄於同目錄下的 error_log.txt"
        )

    def __init__(self):
        super().__init__()
        
        # 視窗基本設定
        self.title(f"新生照片批次處理與更名工具 - {VERSION}")
        self.geometry("1100x750")
        self.minsize(1000, 700)
        
        # 資料與狀態變數
        self.excel_path = ""
        self.source_dir = ""
        self.output_dir = ""
        
        self.excel_headers = []
        self.id_column_var = tk.StringVar()
        self.student_id_column_var = tk.StringVar()
        self.student_mapping = {}
        
        self.photos = []          # 儲存照片資訊的清單
        self.selected_idx = -1    # 當前選取的照片索引
        
        # 預設通用影像處理參數
        self.target_width = 354   # 預設 2 吋證件照寬 (3:4)
        self.target_height = 472  # 預設 2 吋證件照高
        
        # 界面元件變數
        self.photo_widgets = []
        
        # 滑鼠拖曳狀態變數
        self.drag_start_y = 0
        self.preview_ratio = 1.0
        self.orig_height = 1.0
        
        # 初始化介面佈局
        self._create_widgets()
        # 啟動版本更新檢查（非阻塞）
        threading.Thread(target=self._check_for_updates, daemon=True).start()
        
    def _check_for_updates(self):
        try:
            response = requests.get(UPDATE_URL, timeout=5)
            if response.status_code == 200:
                data = response.json()
                latest_version = data.get("tag_name")
                if latest_version and latest_version != VERSION:
                    print(f"發現新版本: {latest_version}")
        except Exception:
            pass

    def _create_widgets(self):
        # 設定主網格佈局 (1 row, 2 columns)
        self.grid_columnconfigure(0, weight=0, minsize=400) # 左側控制面板固定寬度
        self.grid_columnconfigure(1, weight=1)              # 右側工作區自適應
        self.grid_rowconfigure(0, weight=1)
        
        # =====================================================================
        # 左側控制面板 (Control Panel)
        # =====================================================================
        self.sidebar = ctk.CTkFrame(self, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        
        # 標題
        title_label = ctk.CTkLabel(self.sidebar, text="照片批次處理系統", font=ctk.CTkFont(family="Microsoft JhengHei", size=20, weight="bold"))
        title_label.pack(padx=20, pady=(20, 10), anchor="w")
        
        # 捲動容器 (放置所有設定)
        settings_scroll = ctk.CTkScrollableFrame(self.sidebar, fg_color="transparent")
        settings_scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        # ---- 區塊一：Excel 對照設定 ----
        excel_group = ctk.CTkFrame(settings_scroll)
        excel_group.pack(fill="x", pady=5)
        
        excel_title = ctk.CTkLabel(excel_group, text="1. 對照表設定 (Excel/CSV)", font=ctk.CTkFont(family="Microsoft JhengHei", size=14, weight="bold"))
        excel_title.pack(padx=10, pady=(10, 5), anchor="w")
        
        self.excel_btn = ctk.CTkButton(excel_group, text="選擇對照檔案 (Excel/CSV)", command=self._select_excel_file)
        self.excel_btn.pack(padx=10, pady=5, fill="x")
        
        self.excel_path_label = ctk.CTkLabel(excel_group, text="未選擇檔案", wraplength=360, anchor="w", justify="left", text_color="gray")
        self.excel_path_label.pack(padx=10, pady=2, fill="x")
        
        # 欄位選擇 (預設隱藏，讀取 Excel 後顯示)
        self.column_select_frame = ctk.CTkFrame(excel_group, fg_color="transparent")
        self.column_select_frame.pack(fill="x", padx=10, pady=5)
        
        id_lbl = ctk.CTkLabel(self.column_select_frame, text="身分證欄位:")
        id_lbl.grid(row=0, column=0, sticky="w", pady=2)
        self.id_combobox = ctk.CTkComboBox(self.column_select_frame, variable=self.id_column_var, state="disabled", command=self._on_column_changed)
        self.id_combobox.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=2)
        
        sid_lbl = ctk.CTkLabel(self.column_select_frame, text="學號對應欄位:")
        sid_lbl.grid(row=1, column=0, sticky="w", pady=2)
        self.sid_combobox = ctk.CTkComboBox(self.column_select_frame, variable=self.student_id_column_var, state="disabled", command=self._on_column_changed)
        self.sid_combobox.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=2)
        self.column_select_frame.grid_columnconfigure(1, weight=1)
        
        # ---- 區塊二：來源照片資料夾 ----
        source_group = ctk.CTkFrame(settings_scroll)
        source_group.pack(fill="x", pady=5)
        
        source_title = ctk.CTkLabel(source_group, text="2. 照片來源路徑", font=ctk.CTkFont(family="Microsoft JhengHei", size=14, weight="bold"))
        source_title.pack(padx=10, pady=(10, 5), anchor="w")
        
        self.source_btn = ctk.CTkButton(source_group, text="選擇照片資料夾", command=self._select_source_dir)
        self.source_btn.pack(padx=10, pady=5, fill="x")
        
        self.source_path_label = ctk.CTkLabel(source_group, text="未選擇資料夾", wraplength=360, anchor="w", justify="left", text_color="gray")
        self.source_path_label.pack(padx=10, pady=2, fill="x")
        
        # ---- 區塊三：處理與輸出設定 ----
        output_group = ctk.CTkFrame(settings_scroll)
        output_group.pack(fill="x", pady=5)
        
        output_title = ctk.CTkLabel(output_group, text="3. 規格與輸出設定", font=ctk.CTkFont(family="Microsoft JhengHei", size=14, weight="bold"))
        output_title.pack(padx=10, pady=(10, 5), anchor="w")
        
        # 尺寸規格快速選擇
        size_lbl = ctk.CTkLabel(output_group, text="輸出尺寸規格:")
        size_lbl.pack(padx=10, pady=(5, 0), anchor="w")
        self.size_preset_combobox = ctk.CTkComboBox(
            output_group, 
            values=["2 吋證件照 (354 x 472 px)", "1 吋證件照 (283 x 378 px)", "自訂尺寸"], 
            command=self._on_preset_size_changed
        )
        self.size_preset_combobox.pack(padx=10, pady=5, fill="x")
        
        # 寬度與高度設定 (可手動輸入)
        self.size_input_frame = ctk.CTkFrame(output_group, fg_color="transparent")
        self.size_input_frame.pack(fill="x", padx=10, pady=5)
        
        w_lbl = ctk.CTkLabel(self.size_input_frame, text="寬 (px):")
        w_lbl.grid(row=0, column=0, sticky="w")
        self.w_entry = ctk.CTkEntry(self.size_input_frame, width=80)
        self.w_entry.insert(0, str(self.target_width))
        self.w_entry.grid(row=0, column=1, sticky="w", padx=(5, 15))
        self.w_entry.bind("<FocusOut>", self._on_manual_size_changed)
        self.w_entry.bind("<Return>", self._on_manual_size_changed)
        
        h_lbl = ctk.CTkLabel(self.size_input_frame, text="高 (px):")
        h_lbl.grid(row=0, column=2, sticky="w")
        self.h_entry = ctk.CTkEntry(self.size_input_frame, width=80)
        self.h_entry.insert(0, str(self.target_height))
        self.h_entry.grid(row=0, column=3, sticky="w", padx=5)
        self.h_entry.bind("<FocusOut>", self._on_manual_size_changed)
        self.h_entry.bind("<Return>", self._on_manual_size_changed)
        
        # 輸出格式
        format_lbl = ctk.CTkLabel(output_group, text="輸出相片格式:")
        format_lbl.pack(padx=10, pady=(5, 0), anchor="w")
        self.format_combobox = ctk.CTkComboBox(output_group, values=["JPG (.jpg)", "PNG (.png)"])
        self.format_combobox.set("JPG (.jpg)")
        self.format_combobox.pack(padx=10, pady=5, fill="x")
        
        # 輸出目標資料夾
        dest_lbl = ctk.CTkLabel(output_group, text="輸出資料夾:")
        dest_lbl.pack(padx=10, pady=(5, 0), anchor="w")
        self.dest_btn = ctk.CTkButton(output_group, text="選擇輸出資料夾", fg_color="gray50", hover_color="gray40", command=self._select_output_dir)
        self.dest_btn.pack(padx=10, pady=5, fill="x")
        self.dest_path_label = ctk.CTkLabel(output_group, text="未選擇資料夾 (預設為來源資料夾下的 output)", wraplength=360, anchor="w", justify="left", text_color="gray")
        self.dest_path_label.pack(padx=10, pady=2, fill="x")
        
        # ---- 區塊四：批次處理控制 ----
        action_group = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        action_group.pack(fill="x", side="bottom", padx=20, pady=20)
        
        self.progress_bar = ctk.CTkProgressBar(action_group)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", pady=5)
        
        self.progress_label = ctk.CTkLabel(action_group, text="準備就緒", text_color="gray")
        self.progress_label.pack(pady=2)
        
        self.run_btn = ctk.CTkButton(
            action_group, 
            text="開始批次處理", 
            fg_color="#2ECC71", 
            hover_color="#27AE60", 
            font=ctk.CTkFont(family="Microsoft JhengHei", size=16, weight="bold"),
            command=self._start_batch_processing
        )
        self.run_btn.pack(fill="x", pady=5, ipady=5)
        
        # =====================================================================
        # 右側工作區 (Work Area)
        # =====================================================================
        self.work_area = ctk.CTkFrame(self, fg_color="transparent")
        self.work_area.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        self.work_area.grid_columnconfigure(0, weight=1)
        self.work_area.grid_rowconfigure(0, weight=4) # 上半部照片列表
        self.work_area.grid_rowconfigure(1, weight=5) # 下半部預覽與微調
        
        # ---- 右上：照片列表 ----
        list_container = ctk.CTkFrame(self.work_area)
        list_container.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        list_header_lbl = ctk.CTkLabel(list_container, text="照片清單與學號對照狀態", font=ctk.CTkFont(family="Microsoft JhengHei", size=14, weight="bold"))
        list_header_lbl.pack(padx=15, pady=(10, 5), anchor="w")
        
        # 欄位標頭
        header_frame = ctk.CTkFrame(list_container, height=30, fg_color="gray25")
        header_frame.pack(fill="x", padx=10, pady=2)
        
        h_name = ctk.CTkLabel(header_frame, text="原始照片檔名 (身分證)", width=180, anchor="w", font=ctk.CTkFont(weight="bold"))
        h_name.pack(side="left", padx=10)
        h_sid = ctk.CTkLabel(header_frame, text="對照學號", width=120, anchor="w", font=ctk.CTkFont(weight="bold"))
        h_sid.pack(side="left", padx=10)
        h_status = ctk.CTkLabel(header_frame, text="狀態 / 裁切設定", anchor="w", font=ctk.CTkFont(weight="bold"))
        h_status.pack(side="left", padx=10, fill="x", expand=True)
        
        # 捲動的清單容器
        self.photo_list_frame = ctk.CTkScrollableFrame(list_container, fg_color="transparent")
        self.photo_list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # ---- 右下：即時對比預覽與微調 ----
        preview_container = ctk.CTkFrame(self.work_area)
        preview_container.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        
        preview_header_lbl = ctk.CTkLabel(preview_container, text="相片處理即時預覽與裁切框微調", font=ctk.CTkFont(family="Microsoft JhengHei", size=14, weight="bold"))
        preview_header_lbl.pack(padx=15, pady=(10, 5), anchor="w")
        
        # 預覽區域佈局
        preview_split = ctk.CTkFrame(preview_container, fg_color="transparent")
        preview_split.pack(fill="both", expand=True, padx=10)
        preview_split.grid_columnconfigure(0, weight=1) # 原始預覽
        preview_split.grid_columnconfigure(1, weight=1) # 裁切後預覽
        preview_split.grid_rowconfigure(0, weight=1)
        
        # 左：原始照片 (帶有裁切指示線)
        self.orig_preview_box = ctk.CTkFrame(preview_split, border_width=1, border_color="gray30")
        self.orig_preview_box.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.orig_preview_title = ctk.CTkLabel(self.orig_preview_box, text="原始照片與裁切區域指示", text_color="gray70")
        self.orig_preview_title.pack(pady=2)
        self.orig_img_label = ctk.CTkLabel(self.orig_preview_box, text="請先載入照片並於上方清單點選預覽", text_color="gray50")
        self.orig_img_label.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 綁定滑鼠拖曳裁切框事件
        self.orig_img_label.bind("<Button-1>", self._on_orig_img_press)
        self.orig_img_label.bind("<B1-Motion>", self._on_orig_img_drag)
        self.orig_img_label.bind("<ButtonRelease-1>", self._on_orig_img_release)
        
        # 右：裁切後預覽
        self.dest_preview_box = ctk.CTkFrame(preview_split, border_width=1, border_color="gray30")
        self.dest_preview_box.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        self.dest_preview_title = ctk.CTkLabel(self.dest_preview_box, text="裁切與轉檔後效果預覽", text_color="gray70")
        self.dest_preview_title.pack(pady=2)
        self.dest_img_label = ctk.CTkLabel(self.dest_preview_box, text="預覽區域", text_color="gray50")
        self.dest_img_label.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 微調控制面板 (滑桿)
        control_panel = ctk.CTkFrame(preview_container, height=80, fg_color="transparent")
        control_panel.pack(fill="x", padx=15, pady=(5, 10))
        control_panel.grid_columnconfigure(0, weight=1)
        control_panel.grid_columnconfigure(1, weight=1)
        control_panel.grid_rowconfigure(0, weight=1)
        control_panel.grid_rowconfigure(1, weight=1)
        
        # 滑桿 1：垂直偏移量 (Vertical Offset)
        v_frame = ctk.CTkFrame(control_panel, fg_color="transparent")
        v_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=2)
        v_lbl = ctk.CTkLabel(v_frame, text="垂直位移 (上下):", width=100, anchor="w")
        v_lbl.pack(side="left")
        self.v_offset_slider = ctk.CTkSlider(v_frame, from_=-0.5, to=0.5, number_of_steps=100, command=self._on_slider_changed)
        self.v_offset_slider.set(0.0)
        self.v_offset_slider.pack(side="left", fill="x", expand=True, padx=5)
        self.v_val_label = ctk.CTkLabel(v_frame, text="0.00", width=40)
        self.v_val_label.pack(side="right")
        
        # 滑桿 2：裁切比例 / 縮放 (Zoom)
        s_frame = ctk.CTkFrame(control_panel, fg_color="transparent")
        s_frame.grid(row=0, column=1, sticky="ew", padx=10, pady=2)
        s_lbl = ctk.CTkLabel(s_frame, text="裁切框大小 (縮放):", width=110, anchor="w")
        s_lbl.pack(side="left")
        self.scale_slider = ctk.CTkSlider(s_frame, from_=0.1, to=1.0, number_of_steps=90, command=self._on_slider_changed)
        self.scale_slider.set(1.0)
        self.scale_slider.pack(side="left", fill="x", expand=True, padx=5)
        self.s_val_label = ctk.CTkLabel(s_frame, text="100%", width=40)
        self.s_val_label.pack(side="right")
        
        # 微調快捷鍵與套用
        btn_frame = ctk.CTkFrame(control_panel, fg_color="transparent")
        btn_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(5, 0))
        
        self.preset_center_btn = ctk.CTkButton(btn_frame, text="置中對齊 (0.0)", width=100, fg_color="gray30", hover_color="gray40", command=self._set_preset_center)
        self.preset_center_btn.pack(side="left", padx=5)
        
        self.preset_top_btn = ctk.CTkButton(btn_frame, text="偏上對齊 (0.15)", width=110, fg_color="gray30", hover_color="gray40", command=self._set_preset_top)
        self.preset_top_btn.pack(side="left", padx=5)
        
        self.apply_all_btn = ctk.CTkButton(btn_frame, text="將此裁切設定套用至全部照片", width=200, fg_color="#3498DB", hover_color="#2980B9", command=self._apply_settings_to_all)
        self.apply_all_btn.pack(side="right", padx=5)
        
        self.reset_btn = ctk.CTkButton(btn_frame, text="重設此張微調", width=100, fg_color="gray30", hover_color="gray40", command=self._reset_current_photo_settings)
        self.reset_btn.pack(side="right", padx=5)

    # =====================================================================
    # 事件處理與交互邏輯
    # =====================================================================
    def _check_for_updates(self):
        """向 GitHub 查詢最新 Release，若比當前 VERSION 更新則提示使用者"""
        try:
            resp = requests.get(UPDATE_URL, timeout=5)
            if resp.status_code == 200:
                data = resp.json()
                latest_tag = data.get("tag_name", "")
                if latest_tag:
                    # 移除前導的 'v' 或 'V'
                    cur = VERSION.lstrip('vV')
                    latest = latest_tag.lstrip('vV')
                    # 簡易版本比較（僅支援 X.Y.Z 格式）
                    def to_tuple(v):
                        return tuple(int(part) for part in v.split('.') if part.isdigit())
                    if to_tuple(latest) > to_tuple(cur):
                        messagebox.showinfo(
                            "版本更新",
                            f"有新版本可用: {latest_tag}\n目前版本: {VERSION}\n請前往 GitHub 下載最新版本。"
                        )
        except Exception as e:
            # 若檢查失敗，僅記錄錯誤，不打擾使用者
            print(f"更新檢查失敗: {e}")

    def _select_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="選擇對照資料 (Excel/CSV)",
            filetypes=[("對照檔案", "*.xlsx;*.xls;*.csv"), ("All Files", "*.*")]
        )
        if not file_path:
            return
            
        self.excel_path = file_path
        self.excel_path_label.configure(text=os.path.basename(file_path), text_color="white")
        
        try:
            # 讀取 Excel 標頭
            self.excel_headers = PhotoProcessor.get_excel_headers(self.excel_path)
            
            # 更新下拉選單
            self.id_combobox.configure(values=self.excel_headers, state="normal")
            self.sid_combobox.configure(values=self.excel_headers, state="normal")
            
            # 智慧預測欄位：尋找包含「身分證」、「證號」、「ID」、「學號」、「NO」等關鍵字的欄位
            predicted_id = ""
            predicted_sid = ""
            for h in self.excel_headers:
                h_lower = h.lower()
                if "身分證" in h or "證號" in h or "id" in h_lower or "identity" in h_lower:
                    predicted_id = h
                elif "學號" in h or "student" in h_lower or "sid" in h_lower or "no" in h_lower:
                    predicted_sid = h
            
            if predicted_id:
                self.id_combobox.set(predicted_id)
            elif self.excel_headers:
                self.id_combobox.set(self.excel_headers[0])
                
            if predicted_sid:
                self.sid_combobox.set(predicted_sid)
            elif len(self.excel_headers) > 1:
                self.sid_combobox.set(self.excel_headers[1])
            elif self.excel_headers:
                self.sid_combobox.set(self.excel_headers[0])
                
            self._load_excel_data()
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取 Excel 失敗:\n{str(e)}")
            self.excel_path = ""
            self.excel_path_label.configure(text="讀取失敗", text_color="red")
            
    def _load_excel_data(self):
        id_col = self.id_column_var.get()
        sid_col = self.student_id_column_var.get()
        
        if not self.excel_path or not id_col or not sid_col:
            return
            
        try:
            self.student_mapping = PhotoProcessor.load_excel_mapping(self.excel_path, id_col, sid_col)
            self._update_photo_match_status()
            messagebox.showinfo("成功", f"成功載入對照表！\n共載入 {len(self.student_mapping)} 筆對應資料。")
        except Exception as e:
            messagebox.showerror("錯誤", f"載入對照表資料失敗:\n{str(e)}")

    def _on_column_changed(self, event=None):
        self._load_excel_data()

    def _select_source_dir(self):
        dir_path = filedialog.askdirectory(title="選擇新生照片來源資料夾")
        if not dir_path:
            return
            
        self.source_dir = dir_path
        self.source_path_label.configure(text=dir_path, text_color="white")
        
        # 自動填入預設的輸出資料夾 (來源路徑下的 output)
        if not self.output_dir:
            default_out = os.path.join(dir_path, "output")
            self.dest_path_label.configure(text=f"預設: {default_out}", text_color="gray")
            
        try:
            raw_photos = PhotoProcessor.scan_source_directory(self.source_dir)
            
            # 初始化照片的裁切參數與狀態
            self.photos = []
            for p in raw_photos:
                self.photos.append({
                    'filename': p['filename'],
                    'filepath': p['filepath'],
                    'id': p['id'],
                    'student_id': None,
                    'status': '待處理',
                    'scale': 1.0,
                    'vertical_offset': 0.0,
                })
                
            self._update_photo_match_status()
            self._refresh_photo_list_ui()
            
            if self.photos:
                self._select_photo(0) # 預設選取第一張
            else:
                self.selected_idx = -1
                self.orig_img_label.configure(text="此資料夾內沒有支援的相片檔案 (.jpg, .jpeg, .png, .bmp, .webp, .tiff)", image=None)
                self.dest_img_label.configure(text="預覽區域", image=None)
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取相片資料夾失敗:\n{str(e)}")

    def _select_output_dir(self):
        dir_path = filedialog.askdirectory(title="選擇輸出儲存資料夾")
        if not dir_path:
            return
        self.output_dir = dir_path
        self.dest_path_label.configure(text=dir_path, text_color="white")

    def _update_photo_match_status(self):
        """
        比對照片檔名 (身分證字號) 與 Excel 中的學號對照表。
        """
        if not self.photos:
            return
            
        for p in self.photos:
            photo_id = p['id']
            if self.student_mapping:
                if photo_id in self.student_mapping:
                    p['student_id'] = self.student_mapping[photo_id]
                    p['status'] = '已對照'
                else:
                    p['student_id'] = None
                    p['status'] = '無對照學號'
            else:
                p['student_id'] = None
                p['status'] = '待處理 (未匯入 Excel)'

    def _refresh_photo_list_ui(self):
        # 清除舊的 widget
        for w in self.photo_widgets:
            w.destroy()
        self.photo_widgets = []
        
        if not self.photos:
            empty_lbl = ctk.CTkLabel(self.photo_list_frame, text="無相片資料，請點擊左側「選擇照片資料夾」載入。")
            empty_lbl.pack(pady=20)
            self.photo_widgets.append(empty_lbl)
            return
            
        # 繪製清單
        for idx, p in enumerate(self.photos):
            # 行容器 (Frame)
            row_bg = "gray20" if idx != self.selected_idx else "#1F6AA5"
            row_frame = ctk.CTkFrame(self.photo_list_frame, height=35, fg_color=row_bg, corner_radius=4)
            row_frame.pack(fill="x", pady=2, padx=5)
            
            # 使用 bind 來使整行卡片點擊皆可觸發選取
            row_frame.bind("<Button-1>", lambda event, i=idx: self._select_photo(i))
            
            # 檔名欄
            lbl_name = ctk.CTkLabel(row_frame, text=p['filename'], width=180, anchor="w")
            lbl_name.pack(side="left", padx=10)
            lbl_name.bind("<Button-1>", lambda event, i=idx: self._select_photo(i))
            
            # 學號欄
            sid_str = p['student_id'] if p['student_id'] else "---"
            lbl_sid = ctk.CTkLabel(row_frame, text=sid_str, width=120, text_color="gray85" if p['student_id'] else "gray50", anchor="w")
            lbl_sid.pack(side="left", padx=10)
            lbl_sid.bind("<Button-1>", lambda event, i=idx: self._select_photo(i))
            
            # 狀態與微調參數欄
            status_text = p['status']
            if p['scale'] != 1.0 or p['vertical_offset'] != 0.0:
                status_text += f" (已微調: 縮放 {int(p['scale']*100)}%, 偏移 {p['vertical_offset']:.2f})"
                
            status_color = "gray"
            if p['status'] == '已對照':
                status_color = "#2ECC71" # 綠色
            elif p['status'] in ('無對照學號', '失敗'):
                status_color = "#E74C3C" # 紅色
            elif p['status'] == '成功':
                status_color = "#3498DB" # 藍色
                
            lbl_status = ctk.CTkLabel(row_frame, text=status_text, text_color=status_color, anchor="w")
            lbl_status.pack(side="left", padx=10, fill="x", expand=True)
            lbl_status.bind("<Button-1>", lambda event, i=idx: self._select_photo(i))
            
            self.photo_widgets.append(row_frame)

    def _select_photo(self, idx):
        if not (0 <= idx < len(self.photos)):
            return
            
        # 更新選取索引
        old_selected = self.selected_idx
        self.selected_idx = idx
        
        # 刷新原選中列與新選中列的背景顏色，避免重建整個 UI
        if old_selected != -1 and old_selected < len(self.photo_widgets):
            try:
                self.photo_widgets[old_selected].configure(fg_color="gray20")
            except Exception:
                pass
        if self.selected_idx < len(self.photo_widgets):
            try:
                self.photo_widgets[self.selected_idx].configure(fg_color="#1F6AA5")
            except Exception:
                pass
        
        # 載入此相片的裁切參數至滑桿
        p = self.photos[idx]
        self.v_offset_slider.set(p['vertical_offset'])
        self.scale_slider.set(p['scale'])
        self.v_val_label.configure(text=f"{p['vertical_offset']:.2f}")
        self.s_val_label.configure(text=f"{int(p['scale']*100)}%")
        
        # 繪製預覽圖
        self._update_preview()

    def _on_slider_changed(self, value=None):
        if self.selected_idx == -1:
            return
            
        # 取得滑桿數值並更新到當前選取照片
        v_offset = self.v_offset_slider.get()
        scale = self.scale_slider.get()
        
        # 四捨五入以優化顯示與計算
        v_offset = round(v_offset, 2)
        scale = round(scale, 2)
        
        self.v_val_label.configure(text=f"{v_offset:.2f}")
        self.s_val_label.configure(text=f"{int(scale*100)}%")
        
        # 暫存到照片設定中
        p = self.photos[self.selected_idx]
        p['vertical_offset'] = v_offset
        p['scale'] = scale
        
        # 即時更新預覽
        self._update_preview()
        
        # 更新列表中的文字狀態描述，避免整個重繪
        # 我們直接找到該行的 label 進行更新
        status_text = p['status']
        if p['scale'] != 1.0 or p['vertical_offset'] != 0.0:
            status_text += f" (已微調: 縮放 {int(p['scale']*100)}%, 偏移 {p['vertical_offset']:.2f})"
        
        # row_frame.winfo_children() 的第三個元件是 lbl_status
        if self.selected_idx < len(self.photo_widgets):
            try:
                lbl_status = self.photo_widgets[self.selected_idx].winfo_children()[2]
                lbl_status.configure(text=status_text)
            except Exception:
                pass

    def _set_preset_center(self):
        self.v_offset_slider.set(0.0)
        self._on_slider_changed()
        
    def _set_preset_top(self):
        self.v_offset_slider.set(0.15) # 偏上 0.15 是最適合大部分新生照片頭像對齊的通用參數
        self._on_slider_changed()

    def _reset_current_photo_settings(self):
        self.v_offset_slider.set(0.0)
        self.scale_slider.set(1.0)
        self._on_slider_changed()

    def _apply_settings_to_all(self):
        if self.selected_idx == -1:
            return
        
        curr_v = round(self.v_offset_slider.get(), 2)
        curr_s = round(self.scale_slider.get(), 2)
        
        if not messagebox.askyesno("確認", f"您確定要把目前的裁切參數套用到所有照片嗎？\n(垂直位移: {curr_v:.2f}, 縮放比: {int(curr_s*100)}%)"):
            return
            
        for p in self.photos:
            p['vertical_offset'] = curr_v
            p['scale'] = curr_s
            
        self._refresh_photo_list_ui()
        messagebox.showinfo("成功", "已將目前的裁切參數套用至所有照片。")

    def _on_preset_size_changed(self, choice):
        if choice.startswith("2 吋"):
            self.target_width = 354
            self.target_height = 472
        elif choice.startswith("1 吋"):
            self.target_width = 283
            self.target_height = 378
        else:
            # 自訂尺寸，允許修改 Entry
            pass
            
        self.w_entry.delete(0, tk.END)
        self.w_entry.insert(0, str(self.target_width))
        self.h_entry.delete(0, tk.END)
        self.h_entry.insert(0, str(self.target_height))
        
        self._update_preview()

    def _on_manual_size_changed(self, event=None):
        try:
            w = int(self.w_entry.get().strip())
            h = int(self.h_entry.get().strip())
            if w <= 0 or h <= 0:
                raise ValueError
            self.target_width = w
            self.target_height = h
            self.size_preset_combobox.set("自訂尺寸")
            self._update_preview()
        except ValueError:
            # 復原為原數值
            self.w_entry.delete(0, tk.END)
            self.w_entry.insert(0, str(self.target_width))
            self.h_entry.delete(0, tk.END)
            self.h_entry.insert(0, str(self.target_height))

    # =====================================================================
    # 預覽渲染繪製
    # =====================================================================
    def _update_preview(self):
        if self.selected_idx == -1 or not self.photos:
            return
            
        p = self.photos[self.selected_idx]
        src_path = p['filepath']
        
        # 確保讀取正常
        if not os.path.exists(src_path):
            return
            
        try:
            # 取得 UI 顯示框的預計寬高限制 (約 250 x 250)
            max_disp_w = 260
            max_disp_h = 260
            
            with Image.open(src_path) as img:
                # 1. 修正旋轉
                img = PhotoProcessor._correct_image_orientation(img)
                orig_w, orig_h = img.size
                
                # 2. 計算在原圖上的實際裁切框座標
                v_offset = p['vertical_offset']
                scale = p['scale']
                
                x1, y1, x2, y2 = PhotoProcessor.calculate_crop_box(
                    orig_w, orig_h, self.target_width, self.target_height, scale, v_offset
                )
                
                # 3. 渲染左側原始照片預覽 (包含紅色的裁切指示框線)
                # 等比例縮放至適合預覽框的大小
                ratio = min(max_disp_w / orig_w, max_disp_h / orig_h)
                self.preview_ratio = ratio
                self.orig_height = orig_h
                disp_w = int(orig_w * ratio)
                disp_h = int(orig_h * ratio)
                
                preview_orig = img.resize((disp_w, disp_h), Image.Resampling.BILINEAR)
                draw = ImageDraw.Draw(preview_orig)
                
                # 將原圖裁切框座標映射至顯示尺寸上
                px1 = x1 * ratio
                py1 = y1 * ratio
                px2 = x2 * ratio
                py2 = y2 * ratio
                
                # 畫上紅色裁切框與半透明遮罩
                # 畫出紅色框線 (外框 2px)
                draw.rectangle([px1, py1, px2, py2], outline="#E74C3C", width=2)
                
                # 將 PIL Image 轉換為 Tkinter PhotoImage 顯示
                orig_tk_img = ImageTk.PhotoImage(preview_orig)
                self.orig_img_label.configure(image=orig_tk_img, text="")
                self.orig_img_label.image = orig_tk_img  # 保留參考防止被垃圾回收
                
                # 4. 渲染右側裁切後的結果預覽
                # 先取得高品質裁切 resized 後的圖片 (指定寬高)
                processed_pil = PhotoProcessor.process_image(
                    src_path, 
                    dest_path=None, 
                    target_width=self.target_width, 
                    target_height=self.target_height, 
                    scale=scale, 
                    vertical_offset=v_offset
                )
                
                # 縮放到預覽框大小以便顯示
                dest_ratio = min(max_disp_w / self.target_width, max_disp_h / self.target_height)
                dest_disp_w = int(self.target_width * dest_ratio)
                dest_disp_h = int(self.target_height * dest_ratio)
                
                preview_dest = processed_pil.resize((dest_disp_w, dest_disp_h), Image.Resampling.BILINEAR)
                
                dest_tk_img = ImageTk.PhotoImage(preview_dest)
                self.dest_img_label.configure(image=dest_tk_img, text="")
                self.dest_img_label.image = dest_tk_img
                
                self.orig_preview_title.configure(text=f"原始尺寸: {orig_w}x{orig_h}")
                self.dest_preview_title.configure(text=f"輸出規格: {self.target_width}x{self.target_height}")
                
        except Exception as e:
            self.orig_img_label.configure(text=f"載入預覽失敗:\n{str(e)}", image=None)
            self.dest_img_label.configure(text="載入預覽失敗", image=None)

    # =====================================================================
    # 批次處理邏輯 (非同步執行)
    # =====================================================================
    def _start_batch_processing(self):
        if not self.photos:
            messagebox.showwarning("警告", "請先選擇照片資料夾載入照片。")
            return
            
        # 決定輸出路徑
        out_dir = self.output_dir
        if not out_dir:
            out_dir = os.path.join(self.source_dir, "output")
            self.dest_path_label.configure(text=f"預設: {out_dir}", text_color="gray")
            
        # 檢查是否選擇了欄位
        if self.excel_path and (not self.id_column_var.get() or not self.student_id_column_var.get()):
            messagebox.showwarning("警告", "請選擇 Excel 中對應的身分證和學號欄位名稱。")
            return
            
        # 確認是否執行
        confirmed = messagebox.askyesno(
            "確認執行", 
            f"準備開始批次處理！\n"
            f"照片數量: {len(self.photos)} 張\n"
            f"輸出資料夾: {out_dir}\n"
            f"規格尺寸: {self.target_width} x {self.target_height} px\n\n"
            f"是否開始處理？"
        )
        if not confirmed:
            return
            
        # 禁用執行按鈕避免重複點擊
        self.run_btn.configure(state="disabled", text="處理中...")
        self.excel_btn.configure(state="disabled")
        self.source_btn.configure(state="disabled")
        self.dest_btn.configure(state="disabled")
        
        # 啟動背景線程進行批次處理，避免 GUI 卡死
        thread = threading.Thread(target=self._run_batch, args=(out_dir,))
        thread.daemon = True
        thread.start()

    def _run_batch(self, out_dir):
        try:
            total = len(self.photos)
            success_count = 0
            skip_count = 0
            fail_count = 0
            
            # 解析輸出格式與副檔名
            fmt_choice = self.format_combobox.get()
            out_ext = ".jpg" if "JPG" in fmt_choice else ".png"
            out_format = "JPEG" if "JPG" in fmt_choice else "PNG"
            
            for idx, p in enumerate(self.photos):
                self.progress_label.configure(text=f"正在處理: {p['filename']} ({idx+1}/{total})")
                self.progress_bar.set((idx + 1) / total)
                
                # 計算更名後的檔名
                # 規則：如果有對照到學號，以學號命名。否則，以原檔名（身分證字號）命名，或是依據需求跳過。
                dest_filename = ""
                if p['student_id']:
                    dest_filename = p['student_id'] + out_ext
                    p['status'] = '已對照'
                else:
                    # 無對應學號時的處理：以身分證字號原名輸出，並加 _unmatched 標記
                    dest_filename = p['id'] + "_未對照" + out_ext
                    p['status'] = '無對照學號'
                    skip_count += 1
                
                dest_filepath = os.path.join(out_dir, dest_filename)
                
                try:
                    # 執行圖片轉檔與裁切
                    PhotoProcessor.process_image(
                        src_path=p['filepath'],
                        dest_path=dest_filepath,
                        target_width=self.target_width,
                        target_height=self.target_height,
                        scale=p['scale'],
                        vertical_offset=p['vertical_offset'],
                        output_format=out_format
                    )
                    p['status'] = '成功'
                    success_count += 1
                except Exception as e:
                    print(f"處理出錯 {p['filename']}: {str(e)}", file=sys.stderr)
                    p['status'] = '失敗'
                    fail_count += 1
                    
            # 處理完成，回到主線程更新 UI
            self.after(0, lambda: self._on_batch_complete(success_count, skip_count, fail_count, out_dir))
        except Exception as e:
            # 攔截背景線程錯誤，回到主線程報告，防止 GUI 鎖死
            self.after(0, lambda err=e: self._on_thread_error(err))

    def _on_thread_error(self, error):
        self.run_btn.configure(state="normal", text="開始批次處理")
        self.excel_btn.configure(state="normal")
        self.source_btn.configure(state="normal")
        self.dest_btn.configure(state="normal")
        self.progress_label.configure(text="處理發生錯誤")
        messagebox.showerror("背景處理錯誤", f"批次處理執行緒發生嚴重錯誤：\n{str(error)}")

    def _on_batch_complete(self, success, skipped, failed, out_dir):
        # 恢復按鈕狀態
        self.run_btn.configure(state="normal", text="開始批次處理")
        self.excel_btn.configure(state="normal")
        self.source_btn.configure(state="normal")
        self.dest_btn.configure(state="normal")
        
        self.progress_bar.set(1.0)
        self.progress_label.configure(text="處理完成")
        
        # 重新整理清單顯示最新處理狀態
        self._refresh_photo_list_ui()
        
        msg = f"批次處理完成！\n\n" \
              f"● 成功匯出: {success} 張\n" \
              f"● 無學號對照 (使用原名標記匯出): {skipped} 張\n" \
              f"● 失敗: {failed} 張\n\n" \
              f"輸出路徑: {out_dir}"
              
        if failed > 0:
            messagebox.showwarning("完成 (部分失敗)", msg)
        else:
            messagebox.showinfo("完成", msg)

    def _on_orig_img_press(self, event):
        """
        記錄滑鼠點下時的 y 座標。
        """
        if self.selected_idx == -1 or not self.photos:
            return
        self.drag_start_y = event.y

    def _on_orig_img_drag(self, event):
        """
        處理滑鼠拖曳，計算 delta 位移，依比例轉換為 vertical_offset，並連動更新滑桿與預覽。
        """
        if self.selected_idx == -1 or not self.photos:
            return
            
        p = self.photos[self.selected_idx]
        dy = event.y - self.drag_start_y
        
        if dy == 0:
            return
            
        # 將顯示位移 dy 映射到原始圖片的像素高度
        dy_orig = dy / self.preview_ratio
        
        # 垂直偏移量計算 (dy > 0 代表滑鼠往下拖曳，裁切框往下移，在公式中 vertical_offset 減小)
        d_offset = - (dy_orig / self.orig_height)
        new_offset = p['vertical_offset'] + d_offset
        new_offset = max(-0.5, min(0.5, new_offset))
        
        # 四捨五入避免小數點精度抖動
        new_offset = round(new_offset, 2)
        
        if new_offset != p['vertical_offset']:
            p['vertical_offset'] = new_offset
            self.v_offset_slider.set(new_offset)
            self.v_val_label.configure(text=f"{new_offset:.2f}")
            self._update_preview()
            
            # 更新列表顯示
            status_text = p['status']
            if p['scale'] != 1.0 or p['vertical_offset'] != 0.0:
                status_text += f" (已微調: 縮放 {int(p['scale']*100)}%, 偏移 {p['vertical_offset']:.2f})"
            if self.selected_idx < len(self.photo_widgets):
                try:
                    lbl_status = self.photo_widgets[self.selected_idx].winfo_children()[2]
                    lbl_status.configure(text=status_text)
                except Exception:
                    pass
                    
        # 更新拖曳起點為當前滑鼠位置
        self.drag_start_y = event.y

    def _on_orig_img_release(self, event):
        pass

if __name__ == "__main__":
    app = PhotoProcessorApp()
    app.mainloop()
