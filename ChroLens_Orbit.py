# row 0: 時間輸入框（含關於按鈕）
# row 1: 星期選擇下拉選單
# row 2: 啟動檔案路徑（含瀏覽按鈕）
# row 3: 功能按鈕（建立排程、刪除、保留）
# row 4: 已建立排程 標題
# row 5: 已建立排程清單
# row 6: 保留的排程 標題 + 復原按鈕
# row 7: 保留的排程清單

import tkinter as tk
import ttkbootstrap as tb
from ttkbootstrap import Style
from ttkbootstrap import ttk
from ttkbootstrap.dialogs import Messagebox
from tkinter import filedialog
import subprocess
import datetime
import os
import sys
import shlex
import json
import hashlib
import tempfile

SAVE_FILE = "schedules.json"

class SchedulerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ChroLens_Orbit")
        self.schedules = []
        self.saved_schedules = []

        self.style = Style("darkly")

        # 主框架
        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.pack(fill="both", expand=True)

        # 左側：動態日誌
        self.log_frame = ttk.Frame(self.main_frame)
        self.log_frame.grid(row=0, column=0, sticky="nswe", padx=(0, 10))
        self.log_label = ttk.Label(self.log_frame, text="動態日誌", anchor="w")
        self.log_label.pack(anchor="w")
        self.log_text = tk.Text(self.log_frame, width=30, height=25, state="disabled")
        self.log_text.pack(fill="both", expand=True)

        # 右側：功能區
        self.right_frame = ttk.Frame(self.main_frame)
        self.right_frame.grid(row=0, column=1, sticky="nswe")

        # ====== row 6~7: 保留/復原 + 保留排程清單 ======
        # 合併保留標題、復原按鈕、保留清單到同一個細框
        self.saved_outer_frame = ttk.Frame(self.right_frame, borderwidth=0.5, relief="solid", padding=(4, 2))
        self.saved_outer_frame.grid(row=6, column=0, columnspan=5, rowspan=2, sticky="we", pady=(0, 4))
        self.saved_label = ttk.Label(self.saved_outer_frame, text="保留", anchor="w")
        self.saved_label.grid(row=0, column=0, sticky="w")
        self.restore_btn = ttk.Button(self.saved_outer_frame, text="復原", command=self.restore_schedule)
        self.restore_btn.grid(row=0, column=1, sticky="e", padx=(8, 0))
        self.saved_outer_frame.columnconfigure(0, weight=1)
        self.saved_outer_frame.columnconfigure(1, weight=0)
        # 保留排程清單直接放進同一個框
        self.saved_listbox = tk.Listbox(self.saved_outer_frame, width=60, height=5, selectmode="extended")
        self.saved_listbox.grid(row=1, column=0, columnspan=2, pady=5, sticky="we")

        # ====== row 0: 執行時間、時間輸入框、DAY、星期選擇下拉選單、...、關於 ======
        self.time_outer_frame = ttk.Frame(self.right_frame, borderwidth=0.5, relief="solid", padding=(4,2))
        self.time_outer_frame.grid(row=0, column=0, columnspan=5, sticky="we", pady=(0,4))
        self.time_label = ttk.Label(self.time_outer_frame, text="執行時間：", anchor="w")
        self.time_label.grid(row=0, column=0, sticky="w")

        # --- 時間輸入框改為兩個下拉選單（只能選擇，不能輸入） ---
        self.hour_var = tk.StringVar()
        self.minute_var = tk.StringVar()

        hour_values = [f"{i:02d}" for i in range(0, 24)]
        minute_values = [f"{i:02d}" for i in range(0, 60)]

        self.hour_combobox = ttk.Combobox(
            self.time_outer_frame, width=3, textvariable=self.hour_var, values=hour_values, state="readonly", justify="center", font=("Consolas", 11)
        )
        self.hour_combobox.grid(row=0, column=1, sticky="w")
        self.hour_combobox.set("18")  # 預設18

        ttk.Label(self.time_outer_frame, text=":").grid(row=0, column=2, sticky="w")

        self.minute_combobox = ttk.Combobox(
            self.time_outer_frame, width=3, textvariable=self.minute_var, values=minute_values, state="readonly", justify="center", font=("Consolas", 11)
        )
        self.minute_combobox.grid(row=0, column=3, sticky="w")
        self.minute_combobox.set("00")  # 預設00

        # Day和時間輸入框之間增加空白
        self.time_outer_frame.grid_columnconfigure(4, minsize=60)  # 約6個字元寬度

        self.day_label = ttk.Label(self.time_outer_frame, text="DAY", anchor="w")
        self.day_label.grid(row=0, column=5, sticky="w")
        self.day_combobox = ttk.Combobox(self.time_outer_frame, width=5, state="readonly")
        self.day_combobox['values'] = [str(i) for i in range(0, 8)]
        self.day_combobox.current(0)
        self.day_combobox.grid(row=0, column=6, sticky="w")
        self.about_btn = ttk.Button(self.time_outer_frame, text="關於", command=self.show_about_dialog)
        self.about_btn.grid(row=0, column=7, sticky="e", padx=(8, 0))
        self.time_outer_frame.columnconfigure(0, weight=0)
        self.time_outer_frame.columnconfigure(1, weight=0)
        self.time_outer_frame.columnconfigure(2, weight=0)
        self.time_outer_frame.columnconfigure(3, weight=0)
        self.time_outer_frame.columnconfigure(4, weight=0)
        self.time_outer_frame.columnconfigure(5, weight=1)
        self.time_outer_frame.columnconfigure(6, weight=0)
        self.time_outer_frame.columnconfigure(7, weight=0)

        # ====== row 2~5: 執行中/建立排程/刪除/保留/排程清單 ======
        self.schedule_outer_frame = ttk.Frame(self.right_frame, borderwidth=0.5, relief="solid", padding=(4, 2))
        self.schedule_outer_frame.grid(row=2, column=0, columnspan=5, rowspan=4, sticky="we", pady=(0, 4))
        # 標題
        self.schedule_label = ttk.Label(self.schedule_outer_frame, text="執行中", anchor="w")
        self.schedule_label.grid(row=0, column=0, sticky="w")
        # 功能按鈕
        self.add_btn = ttk.Button(self.schedule_outer_frame, text="建立排程", command=self.add_and_create_task)
        self.add_btn.grid(row=0, column=1, sticky="e", padx=(8, 0))
        self.delete_btn = ttk.Button(self.schedule_outer_frame, text="刪除", command=self.delete_schedule)
        self.delete_btn.grid(row=0, column=2, sticky="e", padx=(8, 0))
        self.save_btn = ttk.Button(self.schedule_outer_frame, text="保留", command=self.save_schedule)
        self.save_btn.grid(row=0, column=3, sticky="e", padx=(8, 0))
        # 讓按鈕區塊自動伸縮
        self.schedule_outer_frame.columnconfigure(0, weight=1)
        self.schedule_outer_frame.columnconfigure(1, weight=0)
        self.schedule_outer_frame.columnconfigure(2, weight=0)
        self.schedule_outer_frame.columnconfigure(3, weight=0)
        # 排程清單
        self.schedule_listbox = tk.Listbox(self.schedule_outer_frame, width=60, height=10, selectmode="extended")
        self.schedule_listbox.grid(row=1, column=0, columnspan=4, pady=5, sticky="we")

        # ====== row 1: 檔案路徑、路徑顯示框、空、瀏覽、空... ======
        self.path_outer_frame = ttk.Frame(self.right_frame, borderwidth=0.5, relief="solid", padding=(4,2))
        self.path_outer_frame.grid(row=1, column=0, columnspan=5, sticky="we", pady=(0,4))
        self.path_label = ttk.Label(self.path_outer_frame, text="檔案路徑：", anchor="w")
        self.path_label.grid(row=0, column=0, sticky="w")
        self.path_entry = ttk.Entry(self.path_outer_frame, width=30)
        self.path_entry.grid(row=0, column=1, columnspan=2, sticky="we")
        # 讓"瀏覽"按鈕對齊上一行"關於"按鈕
        self.browse_btn = ttk.Button(self.path_outer_frame, text="瀏覽", command=self.browse_file)
        self.browse_btn.grid(row=0, column=4, sticky="e", padx=(8, 0))
        self.path_outer_frame.columnconfigure(0, weight=0)
        self.path_outer_frame.columnconfigure(1, weight=1)
        self.path_outer_frame.columnconfigure(2, weight=0)
        self.path_outer_frame.columnconfigure(3, weight=0)
        self.path_outer_frame.columnconfigure(4, weight=1)

        # 讓所有column都能自動伸縮（可選）
        for i in range(10):
            self.right_frame.columnconfigure(i, weight=1)

        self.load_data()

    def log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"{datetime.datetime.now().strftime('%H:%M:%S')} {msg}\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def browse_file(self):
        path = filedialog.askopenfilename(
            filetypes=[
                ("可執行檔案", "*.exe *.lnk *.bat *.py"),
                ("所有檔案", "*.*")
            ]
        )
        if path:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, path)

    def add_and_create_task(self):
        hour_str = self.hour_var.get()
        minute_str = self.minute_var.get()
        day_str = self.day_combobox.get()
        path = self.path_entry.get()
        try:
            hour = int(hour_str) if hour_str.isdigit() else 18
            minute = int(minute_str) if minute_str.isdigit() else 0
            day = int(day_str)
            if not (0 <= hour <= 23 and 0 <= minute <= 59):
                raise ValueError("請輸入正確的時間")
            if not path:
                raise ValueError("請選擇啟動檔案路徑")
            sched = {"hour": hour, "minute": minute, "day": day, "path": path}
            # 建立到 Windows 工作排程器
            self.create_windows_task(sched)
            self.schedules.append(sched)
            self.schedule_listbox.insert(tk.END, f"DAY {day} {hour:02d}:{minute:02d} 執行 {path}")
            self.save_data()
            self.log(f"建立排程並同步到工作排程器：DAY {day} {hour:02d}:{minute:02d} 執行 {path}")
        except Exception as e:
            Messagebox.show_error(str(e))

    def delete_schedule(self):
        idxs = self.schedule_listbox.curselection()
        if not idxs:
            Messagebox.show_error("請選擇要刪除的排程")
            return
        # 反向刪除，避免 index 變動
        for index in reversed(idxs):
            sched = self.schedules.pop(index)
            self.schedule_listbox.delete(index)
            self.delete_windows_task(sched)
            self.log(f"刪除排程並同步移除工作排程器：DAY {sched['day']} {sched['hour']:02d}:{sched['minute']:02d} 執行 {sched['path']}")
        self.save_data()

    def save_schedule(self):
        idxs = self.schedule_listbox.curselection()
        if not idxs:
            Messagebox.show_error("請選擇要保留的排程")
            return
        for index in reversed(idxs):
            sched = self.schedules.pop(index)
            self.schedule_listbox.delete(index)
            self.delete_windows_task(sched)
            name = f"DAY {sched['day']} {sched['hour']:02d}:{sched['minute']:02d} 執行 {sched['path']}"
            self.saved_schedules.append(sched)
            self.saved_listbox.insert(tk.END, name)
            self.log(f"保留排程並同步移除工作排程器：{name}")
        self.save_data()

    def restore_schedule(self):
        idxs = self.saved_listbox.curselection()
        if not idxs:
            Messagebox.show_error("請選擇要復原的排程")
            return
        for index in reversed(idxs):
            sched = self.saved_schedules.pop(index)
            self.saved_listbox.delete(index)
            # 建立到 Windows 工作排程器
            self.create_windows_task(sched)
            self.schedules.append(sched)
            self.schedule_listbox.insert(
                tk.END, f"DAY {sched['day']} {sched['hour']:02d}:{sched['minute']:02d} 執行 {sched['path']}"
            )
            self.log(f"復原排程並同步到工作排程器：DAY {sched['day']} {sched['hour']:02d}:{sched['minute']:02d} 執行 {sched['path']}")
        self.save_data()

    def get_task_name(self, sched):
        # 使用 md5 產生固定 hash
        path_hash = hashlib.md5(sched['path'].encode('utf-8')).hexdigest()[:8]
        return f"ChroLens_{sched['day']}_{sched['hour']:02d}{sched['minute']:02d}_{path_hash}"

    def create_windows_task(self, sched):
        task_name = self.get_task_name(sched)
        time_str = f"{sched['hour']:02d}:{sched['minute']:02d}"
        if sched["day"] == 0:
            schedule_type = "/SC DAILY"
            day_part = ""
        else:
            schedule_type = "/SC WEEKLY"
            week_map = ["MON","TUE","WED","THU","FRI","SAT","SUN"]
            day_part = f"/D {week_map[sched['day']-1]}"
        path = sched["path"]
        ext = os.path.splitext(path)[1].lower()
        if ext == ".py":
            cmd = f'"{sys.executable}" "{path}"'
            tr_path = cmd
        elif ext == ".lnk":
            # 建立一個暫存 bat 檔來啟動捷徑，並強制 chcp 65001
            bat_dir = os.path.join(tempfile.gettempdir(), "ChroLensOrbit_bat")
            os.makedirs(bat_dir, exist_ok=True)
            bat_name = self.get_task_name(sched) + ".bat"
            bat_path = os.path.join(bat_dir, bat_name)
            with open(bat_path, "w", encoding="utf-8") as f:
                f.write(f'chcp 65001\nstart "" "{path}"\n')
            tr_path = f'"{bat_path}"'
        else:
            tr_path = f'"{path}"'
        schtasks_cmd = f'schtasks /Create /TN "{task_name}" {schedule_type} {day_part} /TR {tr_path} /ST {time_str} /F'
        result = subprocess.run(schtasks_cmd, shell=True, capture_output=True, text=True)
        if result.returncode == 0:
            self.log(f"已建立到 Windows 工作排程器：{task_name}")
        else:
            raise Exception(result.stderr)

    def delete_windows_task(self, sched):
        task_name = self.get_task_name(sched)
        schtasks_cmd = f'schtasks /Delete /TN "{task_name}" /F'
        result = subprocess.run(schtasks_cmd, shell=True, capture_output=True, text=True)
        if result.returncode == 0:
            self.log(f"已從 Windows 工作排程器刪除：{task_name}")
        else:
            self.log(f"刪除工作排程器失敗：{result.stderr.strip()} (任務名稱：{task_name})")

    def save_data(self):
        data = {
            "schedules": self.schedules,
            "saved_schedules": self.saved_schedules
        }
        with open(SAVE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def load_data(self):
        if os.path.exists(SAVE_FILE):
            with open(SAVE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                self.schedules = data.get("schedules", [])
                self.saved_schedules = data.get("saved_schedules", [])
            for sched in self.schedules:
                self.schedule_listbox.insert(
                    tk.END, f"DAY {sched['day']} {sched['hour']:02d}:{sched['minute']:02d} 執行 {sched['path']}"
                )
            for sched in self.saved_schedules:
                name = f"DAY {sched['day']} {sched['hour']:02d}:{sched['minute']:02d} 執行 {sched['path']}"
                self.saved_listbox.insert(tk.END, name)

    def on_close(self):
        self.save_data()
        self.root.destroy()

    def show_about_dialog(self):
        about_win = tb.Toplevel(self.root)
        about_win.title("關於 ChroLens_Mimic")
        about_win.geometry("450x300")
        about_win.resizable(False, False)
        about_win.grab_set()
        # 置中顯示
        self.root.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 175
        y = self.root.winfo_y() + 80
        about_win.geometry(f"+{x}+{y}")

        # 設定icon與主程式相同
        try:
            import sys, os
            if getattr(sys, 'frozen', False):
                icon_path = os.path.join(sys._MEIPASS, "觸手眼鏡貓.ico")
            else:
                icon_path = "觸手眼鏡貓.ico"
            about_win.iconbitmap(icon_path)
        except Exception as e:
            print(f"無法設定 about 視窗 icon: {e}")

        frm = tb.Frame(about_win, padding=20)
        frm.pack(fill="both", expand=True)

        tb.Label(frm, text="ChroLens_Orbit\n簡化Windows工作排程的動作\n更簡易的設定排程", font=("Microsoft JhengHei", 11,)).pack(anchor="w", pady=(0, 6))
        link = tk.Label(frm, text="ChroLens_模擬器討論區", font=("Microsoft JhengHei", 10, "underline"), fg="#5865F2", cursor="hand2")
        link.pack(anchor="w")
        link.bind("<Button-1>", lambda e: os.startfile("https://discord.gg/72Kbs4WPPn"))
        github = tk.Label(frm, text="查看更多工具(巴哈)", font=("Microsoft JhengHei", 10, "underline"), fg="#24292f", cursor="hand2")
        github.pack(anchor="w", pady=(8, 0))
        github.bind("<Button-1>", lambda e: os.startfile("https://home.gamer.com.tw/profile/index_creation.php?owner=umiwued&folder=523848"))
        tb.Label(frm, text="Creat By Lucienwooo", font=("Microsoft JhengHei", 11,)).pack(anchor="w", pady=(0, 6))
        tb.Button(frm, text="關閉", command=about_win.destroy, width=8, bootstyle=tb.SECONDARY).pack(anchor="e", pady=(16, 0))

if __name__ == "__main__":
    root = tk.Tk()
    app = SchedulerApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()