from tkinter import filedialog
import customtkinter as ctk
import pandas as pd
from tkinterdnd2 import DND_FILES, TkinterDnD

# 導入你寫好的 comparelist 模組
import comparelist

# 設定外觀主題
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class App(ctk.CTk, TkinterDnD.DnDWrapper):

  def __init__(self):
    super().__init__()
    self.TkdndVersion = TkinterDnD._require(self)


class ExcelComparatorApp(App):

  def __init__(self):
    super().__init__()

    self.title("Excel 智慧比對工具 (支援拖曳)")
    self.geometry("540x520")
    self.resizable(False, False)

    self.old_file_path = None
    self.new_file_path = None

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
    self.methods_dict = comparelist.get_comparison_methods()
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

  def select_old_file(self):
    path = filedialog.askopenfilename(
        title="選擇舊專案檔案", filetypes=[("Excel files", "*.xls *.xlsx")]
    )
    if path:
      self.set_old_file(path)

  def select_new_file(self):
    path = filedialog.askopenfilename(
        title="選擇新專案檔案", filetypes=[("Excel files", "*.xlsx")]
    )
    if path:
      self.set_new_file(path)

  def drop_old_file(self, event):
    path = event.data.strip("{}")
    if path.endswith((".xls", ".xlsx")):
      self.set_old_file(path)

  def drop_new_file(self, event):
    path = event.data.strip("{}")
    if path.endswith(".xlsx"):
      self.set_new_file(path)

  def set_old_file(self, path):
    self.old_file_path = path
    filename = path.split("/")[-1].split("\\")[-1]
    self.lbl_old.configure(
        text=f"✅ 已載入舊專案：\n{filename}", text_color="#2b8a3e"
    )

  def set_new_file(self, path):
    self.new_file_path = path
    filename = path.split("/")[-1].split("\\")[-1]
    self.lbl_new.configure(
        text=f"✅ 已載入新專案：\n{filename}", text_color="#2b8a3e"
    )

  def run_comparison(self):
    if not self.old_file_path or not self.new_file_path:
      print("請先完整選擇舊專案與新專案檔案！")
      return

    selected_method_name = self.option_method.get()

    try:
      # 取得對應的比較函式並統一呼叫介面
      compare_func = self.methods_dict.get(selected_method_name)
      if compare_func:
        print(f"正在執行比對模式：{selected_method_name}")
        result, old_missing, new_missing = compare_func(self.old_file_path, self.new_file_path)
        print(result)
      else:
        print(f"錯誤：找不到對應的比對方法 '{selected_method_name}'")

    except Exception as e:
      print(f"執行過程發生錯誤: {e}")


if __name__ == "__main__":
  app = ExcelComparatorApp()
  app.mainloop()