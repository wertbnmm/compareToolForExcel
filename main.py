"""Excel 智慧比對工具 - CustomTkinter 圖形介面主程式。

負責使用者互動（拖曳/選擇檔案、選擇比對模式、顯示執行結果），
實際的比對邏輯委派給 comparelist / compare 模組，匯出報表則交給 resultExport 模組。
"""
import sys
from tkinter import filedialog
from typing import Any, Callable, Dict, Optional

import customtkinter as ctk
import pandas as pd
from tkinterdnd2 import DND_FILES, TkinterDnD

# 導入你寫好的 comparelist 模組
import comparelist
from resultExport import export_records_report

# 設定外觀主題
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class App(ctk.CTk, TkinterDnD.DnDWrapper):

  def __init__(self):
    super().__init__()
    self.TkdndVersion = TkinterDnD._require(self)

# --- 新增：用來攔截 print 並寫入介面 Textbox 的輔助類別 ---
class TextRedirector:
  """將 print() 輸出導向 CTkTextbox，讓 Console 訊息顯示在圖形介面上。"""

  def __init__(self, text_widget: ctk.CTkTextbox) -> None:
    self.text_widget = text_widget

  def write(self, str_text: str) -> None:
    self.text_widget.configure(state="normal")  # 允許寫入
    self.text_widget.insert("end", str_text)  # 插入文字
    self.text_widget.see("end")  # 自動捲動到最底部
    self.text_widget.configure(state="disabled")  # 設為唯讀防手殘修改
    self.text_widget.update_idletasks()

  def flush(self) -> None:
    pass  # 為了相容檔案串流介面，保留空方法
  
class ExcelComparatorApp(App):

  def __init__(self):
    super().__init__()

    self.title("Excel 智慧比對工具 (支援拖曳)")
    self.geometry("540x520")
    self.resizable(False, False)

    self.old_file_path: Optional[str] = None
    self.new_file_path: Optional[str] = None

    # 主標題
    self.title_label = ctk.CTkLabel(
        self, text="📊 專案 Excel 快速比對工具", font=("Microsoft JhengHei", 20, "bold")
    )
    self.title_label.pack(pady=15)

    # --- 舊專案拖曳與選擇區 ---
    self.frame_old = ctk.CTkFrame(
        self, fg_color=("#e9ecef", "#2b2b2b"), height=75, corner_radius=10
    )
    self.frame_old.pack(fill="x", padx=30, pady=8)
    self.frame_old.pack_propagate(False)

    self.frame_old.grid_columnconfigure(0, weight=1)
    self.frame_old.grid_columnconfigure(1, weight=0)

    self.lbl_old = ctk.CTkLabel(
        self.frame_old,
        text="📁 請將【舊專案】Excel 拖曳至此\n或點擊右側按鈕選擇",
        font=("Microsoft JhengHei", 12),
        text_color="gray",
        anchor="w",
        justify="left",
    )
    self.lbl_old.grid(row=0, column=0, sticky="ew", padx=(20, 10), pady=12)

    self.btn_old = ctk.CTkButton(
        self.frame_old,
        text="選擇舊檔",
        command=self.select_old_file,
        font=("Microsoft JhengHei", 12),
        width=90,
        height=32,
    )
    self.btn_old.grid(row=0, column=1, sticky="e", padx=(0, 20), pady=12)

    self.frame_old.drop_target_register(DND_FILES)
    self.frame_old.dnd_bind("<<Drop>>", self.drop_old_file)

    # --- 新專案拖曳與選擇區 ---
    self.frame_new = ctk.CTkFrame(
        self, fg_color=("#e9ecef", "#2b2b2b"), height=75, corner_radius=10
    )
    self.frame_new.pack(fill="x", padx=30, pady=8)
    self.frame_new.pack_propagate(False)

    self.frame_new.grid_columnconfigure(0, weight=1)
    self.frame_new.grid_columnconfigure(1, weight=0)

    self.lbl_new = ctk.CTkLabel(
        self.frame_new,
        text="📁 請將【新專案】Excel 拖曳至此\n或點擊右側按鈕選擇",
        font=("Microsoft JhengHei", 12),
        text_color="gray",
        anchor="w",
        justify="left",
    )
    self.lbl_new.grid(row=0, column=0, sticky="ew", padx=(20, 10), pady=12)

    self.btn_new = ctk.CTkButton(
        self.frame_new,
        text="選擇新檔",
        command=self.select_new_file,
        font=("Microsoft JhengHei", 12),
        width=90,
        height=32,
    )
    self.btn_new.grid(row=0, column=1, sticky="e", padx=(0, 20), pady=12)

    self.frame_new.drop_target_register(DND_FILES)
    self.frame_new.dnd_bind("<<Drop>>", self.drop_new_file)

    # --- 比對方法選擇區 (下拉選單) ---
    self.frame_method = ctk.CTkFrame(
        self, fg_color=("#e9ecef", "#2b2b2b"), height=65, corner_radius=10
    )
    self.frame_method.pack(fill="x", padx=30, pady=8)
    self.frame_method.pack_propagate(False)

    self.frame_method.grid_columnconfigure(0, weight=1)
    self.frame_method.grid_columnconfigure(1, weight=0)

    self.lbl_method = ctk.CTkLabel(
        self.frame_method,
        text="⚙️ 選擇比對模式：",
        font=("Microsoft JhengHei", 12),
        anchor="w",
    )
    self.lbl_method.grid(row=0, column=0, sticky="ew", padx=(20, 10), pady=15)

    # 取得 comparelist 裡面的方法字典
    self.methods_dict: Dict[str, Callable] = comparelist.get_comparison_methods()
    method_names = (
        list(self.methods_dict.keys())
        if self.methods_dict
        else ["無可用比對方法"]
    )

    self.option_method = ctk.CTkOptionMenu(
        self.frame_method,
        values=method_names,
        font=("Microsoft JhengHei", 12),
        width=160,
    )
    self.option_method.grid(row=0, column=1, sticky="e", padx=(0, 20), pady=15)
    
    if method_names:
      self.option_method.set(method_names[0])

    # --- 執行按鈕 ---
    self.btn_run = ctk.CTkButton(
        self,
        text="開始執行比對",
        command=self.run_comparison,
        font=("Microsoft JhengHei", 15, "bold"),
        fg_color="#2b8a3e",
        hover_color="#237032",
        width=240,
        height=42,
        corner_radius=10,
    )
    self.btn_run.pack(pady=15)

    # --- 即時 Console 輸出視窗 (CTkTextbox) ---
    self.console_box = ctk.CTkTextbox(
        self,
        height=150,
        corner_radius=6,
        font=("Consolas", 10),
    )
    self.console_box.pack(fill="x", padx=20, pady=(0, 10))
    self.console_box.configure(state="disabled")

    sys.stdout = TextRedirector(self.console_box)
    print("✨ 系統初始化完成，隨時可以開始比對！\n")

  def select_old_file(self) -> None:
    """透過檔案選擇對話框載入舊專案 Excel。"""
    path = filedialog.askopenfilename(
        title="選擇舊專案檔案", filetypes=[("Excel files", "*.xls *.xlsx")]
    )
    if path:
      self.set_old_file(path)

  def select_new_file(self) -> None:
    """透過檔案選擇對話框載入新專案 Excel。"""
    path = filedialog.askopenfilename(
        title="選擇新專案檔案", filetypes=[("Excel files", "*.xlsx")]
    )
    if path:
      self.set_new_file(path)

  def drop_old_file(self, event: Any) -> None:
    """處理拖曳檔案至「舊專案」拖曳區的事件。"""
    path = event.data.strip("{}")
    if path.endswith((".xls", ".xlsx")):
      self.set_old_file(path)

  def drop_new_file(self, event: Any) -> None:
    """處理拖曳檔案至「新專案」拖曳區的事件。"""
    path = event.data.strip("{}")
    if path.endswith(".xlsx"):
      self.set_new_file(path)

  def set_old_file(self, path: str) -> None:
    """記錄舊專案檔案路徑，並更新畫面上的提示標籤。"""
    self.old_file_path = path
    filename = path.split("/")[-1].split("\\")[-1]
    self.lbl_old.configure(
        text=f"✅ 已載入舊專案：\n{filename}", text_color="#2b8a3e"
    )

  def set_new_file(self, path: str) -> None:
    """記錄新專案檔案路徑，並更新畫面上的提示標籤。"""
    self.new_file_path = path
    filename = path.split("/")[-1].split("\\")[-1]
    self.lbl_new.configure(
        text=f"✅ 已載入新專案：\n{filename}", text_color="#2b8a3e"
    )

  def run_comparison(self) -> None:
    """依照目前選定的比對模式執行比對，並將差異資料匯出成 Markdown 報表。"""
    if not self.old_file_path or not self.new_file_path:
      print("請先完整選擇舊專案與新專案檔案！")
      return

    selected_method_name = self.option_method.get()

    try:
      # 取得對應的比較函式並統一呼叫介面
      compare_func = self.methods_dict.get(selected_method_name)
      if compare_func:
        print(f"正在執行比對模式：{selected_method_name}")
        result = compare_func(self.old_file_path, self.new_file_path)
        print(result)
      else:
        print(f"錯誤：找不到對應的比對方法 '{selected_method_name}'")

    except Exception as e:
      print(f"執行過程發生錯誤: {e}")


if __name__ == "__main__":
  app = ExcelComparatorApp()
  app.mainloop()