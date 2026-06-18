"""
点名器 v2.0 — 随机点名工具
功能：读取 Excel 名单，4 种点名模式，点名历史

模式说明：
  ① 自由点名  — 可重复，每次随机抽取（带滚动动画）
  ② 不重复点名 — 所有人点完才重复（带滚动动画）
  ③ 分组抽选  — 一次抽 N 人，适合分组活动
  ④ 顺序轮流  — 打乱顺序后轮流点名，公平公正
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import random
import os
import sys

try:
    import openpyxl
except ImportError:
    print("缺少 openpyxl 库，请运行: pip install openpyxl")
    sys.exit(1)

# ============================================================
# 主题配色
# ============================================================
COLORS = {
    "bg": "#f0f4f8",
    "card": "#ffffff",
    "accent": "#4a90d9",
    "accent_hover": "#357abd",
    "text": "#2c3e50",
    "text_sub": "#6b7280",
    "text_muted": "#9ca3af",
    "success": "#22c55e",
    "danger": "#ef4444",
    "warning": "#f59e0b",
    "border": "#d1d5db",
    "highlight": "#e8f4fd",
}

# ============================================================
# 字体（按系统自动回退）
# ============================================================
if sys.platform == "win32":
    FONT_FAMILY = ("Microsoft YaHei",)
else:
    FONT_FAMILY = ("PingFang SC", "Microsoft YaHei", "Noto Sans CJK SC", "Heiti SC")


def font(size, bold=False):
    return (*FONT_FAMILY, size, "bold" if bold else "normal")


# ============================================================
# 模式定义
# ============================================================
MODES = [
    ("free", "🎲  自由点名", "可重复点名，每次随机抽取"),
    ("exhaust", "📋  不重复点名", "所有人点完才可重复"),
    ("group", "👥  分组抽选", "一次性抽取多人，适合分组活动"),
    ("rotation", "🔄  顺序轮流", "打乱顺序后轮流点名，公平公正"),
]


class NamePicker:
    def __init__(self, root):
        self.root = root
        self.root.title("🎯 随机点名器 v2.0")
        self.root.geometry("780x680")
        self.root.minsize(620, 560)
        self.root.configure(bg=COLORS["bg"])

        # ---- 数据 ----
        self.students = []          # [(学号, 姓名)]
        self.names = []             # [姓名]
        self.remaining = []         # 不重复 / 分组 模式的剩余池
        self.rotation_list = []     # 顺序模式的队列
        self.rotation_idx = 0       # 顺序模式指针

        self.current_result = None  # 当前显示的结果
        self.is_running = False
        self.mode_key = "free"
        self.group_size = tk.IntVar(value=3)

        self.history = []           # 点名记录
        self.total_picks = 0        # 总点名次数（含重复）
        self.round_num = 1          # 第几轮（不重复模式用）

        self.after_id = None
        self.roll_interval = 80     # 滚动间隔 ms

        # ---- 界面 ----
        self.setup_ui()
        self.bind_shortcuts()

    # ============================================================
    # UI 构建
    # ============================================================
    def setup_ui(self):
        # ===== 标题 =====
        title_frame = tk.Frame(self.root, bg=COLORS["bg"])
        title_frame.pack(fill="x", padx=24, pady=(16, 4))
        tk.Label(title_frame, text="🎯 随机点名器",
                 font=font(22, bold=True), fg=COLORS["text"],
                 bg=COLORS["bg"]).pack(side="left")

        # ===== 设置区域（卡片） =====
        config = tk.Frame(self.root, bg=COLORS["card"],
                          highlightbackground=COLORS["border"],
                          highlightthickness=1, padx=16, pady=10)
        config.pack(fill="x", padx=20, pady=(4, 8))

        # 行1：文件选择
        row1 = tk.Frame(config, bg=COLORS["card"])
        row1.pack(fill="x", pady=2)
        tk.Label(row1, text="📂 数据文件", font=font(11),
                 bg=COLORS["card"], fg=COLORS["text"]).pack(side="left")
        self.file_label = tk.Label(row1, text="未选择文件",
                                   font=font(10), fg=COLORS["text_muted"],
                                   bg=COLORS["card"])
        self.file_label.pack(side="left", padx=(10, 0), fill="x", expand=True)
        tk.Button(row1, text="选择 Excel 文件", font=font(10),
                  bg=COLORS["accent"], fg="white", relief="flat",
                  activebackground=COLORS["accent_hover"],
                  activeforeground="white", padx=12, pady=2,
                  cursor="hand2", command=self.select_file).pack(side="right")

        # 行2：模式选择
        row2 = tk.Frame(config, bg=COLORS["card"])
        row2.pack(fill="x", pady=(6, 2))

        # 模式下拉
        self.mode_var = tk.StringVar(value=MODES[0][1])
        self.mode_combo = ttk.Combobox(row2, textvariable=self.mode_var,
                                       values=[m[1] for m in MODES],
                                       state="readonly", width=18,
                                       font=font(10))
        self.mode_combo.pack(side="left")
        self.mode_combo.bind("<<ComboboxSelected>>", self.on_mode_change)

        # 模式说明
        self.mode_desc = tk.Label(row2, text=MODES[0][2],
                                  font=font(9), fg=COLORS["text_sub"],
                                  bg=COLORS["card"])
        self.mode_desc.pack(side="left", padx=(10, 0))

        # 分组大小（仅分组模式显示）
        self.group_frame = tk.Frame(row2, bg=COLORS["card"])
        tk.Label(self.group_frame, text="每组", font=font(10),
                 bg=COLORS["card"], fg=COLORS["text"]).pack(side="left")
        self.group_spin = tk.Spinbox(self.group_frame, from_=2, to=10,
                                     textvariable=self.group_size, width=3,
                                     font=font(10), relief="solid",
                                     bd=1, justify="center")
        self.group_spin.pack(side="left", padx=4)
        tk.Label(self.group_frame, text="人", font=font(10),
                 bg=COLORS["card"], fg=COLORS["text"]).pack(side="left")
        self.group_frame.pack_forget()

        # ===== 统计栏 =====
        self.stats_bar = tk.Frame(self.root, bg=COLORS["bg"])
        self.stats_bar.pack(fill="x", padx=24, pady=(2, 0))
        self.stats_label = tk.Label(self.stats_bar, text="请先加载学生名单",
                                    font=font(10), fg=COLORS["text_muted"],
                                    bg=COLORS["bg"])
        self.stats_label.pack(side="left")

        # ===== 中央显示区（卡片） =====
        display_container = tk.Frame(self.root, bg=COLORS["bg"])
        display_container.pack(fill="both", expand=True, padx=20, pady=8)

        self.display_card = tk.Frame(display_container, bg=COLORS["card"],
                                     highlightbackground=COLORS["border"],
                                     highlightthickness=1)
        self.display_card.pack(fill="both", expand=True)

        # 主名字（大号）
        self.name_label = tk.Label(self.display_card,
                                   text="📂 请选择数据文件",
                                   font=font(52, bold=True),
                                   fg=COLORS["text_muted"],
                                   bg=COLORS["card"])
        self.name_label.pack(fill="both", expand=True, pady=(50, 4))

        # 副信息（学号 / 状态）
        self.sub_info = tk.Label(self.display_card, text="",
                                 font=font(14), fg=COLORS["text_sub"],
                                 bg=COLORS["card"])
        self.sub_info.pack(pady=(0, 50))

        # 分组模式：多人结果展示区（平时隐藏）
        self.group_result_frame = tk.Frame(self.display_card,
                                           bg=COLORS["card"])
        self.group_result_frame.pack_forget()

        # ===== 按钮栏 =====
        btn_bar = tk.Frame(self.root, bg=COLORS["bg"])
        btn_bar.pack(pady=(8, 4))

        self.btn_start = tk.Button(
            btn_bar, text="▶  开始", font=font(13, bold=True),
            bg=COLORS["success"], fg="white", relief="flat",
            activebackground="#16a34a", activeforeground="white",
            padx=36, pady=8, cursor="hand2", state="disabled",
            command=self.start
        )
        self.btn_start.pack(side="left", padx=8)

        self.btn_stop = tk.Button(
            btn_bar, text="⏹  停止", font=font(13, bold=True),
            bg=COLORS["danger"], fg="white", relief="flat",
            activebackground="#dc2626", activeforeground="white",
            padx=36, pady=8, cursor="hand2",
            command=self.stop
        )
        self.btn_stop.pack(side="left", padx=8)

        # 快捷提示
        tip_frame = tk.Frame(self.root, bg=COLORS["bg"])
        tip_frame.pack(pady=(0, 4))
        tk.Label(tip_frame, text="Enter / Space 快速开始/停止",
                 font=font(9), fg=COLORS["text_muted"],
                 bg=COLORS["bg"]).pack()

        # ===== 历史记录（底部） =====
        history_frame = tk.Frame(self.root, bg=COLORS["card"],
                                 highlightbackground=COLORS["border"],
                                 highlightthickness=1)
        history_frame.pack(fill="x", padx=20, pady=(2, 14), ipady=2)

        header_frame = tk.Frame(history_frame, bg=COLORS["card"])
        header_frame.pack(fill="x", padx=12, pady=(6, 2))
        tk.Label(header_frame, text="📝 点名记录", font=font(11, bold=True),
                 bg=COLORS["card"], fg=COLORS["text"]).pack(side="left")
        tk.Button(header_frame, text="清空", font=font(9),
                  bg=COLORS["bg"], fg=COLORS["text_sub"], relief="flat",
                  activebackground=COLORS["border"], padx=8, cursor="hand2",
                  command=self.clear_history).pack(side="right")

        self.history_text = tk.Text(history_frame, height=3, width=1,
                                    font=font(10), fg=COLORS["text_sub"],
                                    bg=COLORS["card"], relief="flat",
                                    state="disabled", wrap="none",
                                    borderwidth=0)
        self.history_text.pack(fill="x", padx=12, pady=(0, 6))

        # 水平滚动
        h_scroll = tk.Scrollbar(history_frame, orient="horizontal",
                                command=self.history_text.xview)
        self.history_text.configure(xscrollcommand=h_scroll.set)
        h_scroll.pack(fill="x", padx=12, pady=(0, 6))

        self.update_history_display()

    # ============================================================
    # 快捷键
    # ============================================================
    def bind_shortcuts(self):
        self.root.bind("<Return>", lambda e: self.toggle_run())
        self.root.bind("<space>", lambda e: self.toggle_run())

    def toggle_run(self):
        if not self.names:
            return
        if self.mode_key in ("free", "exhaust"):
            if self.is_running:
                self.stop()
            else:
                self.start()
        elif self.mode_key == "group":
            self.pick_group()
        elif self.mode_key == "rotation":
            self.next_rotation()

    # ============================================================
    # 文件选择
    # ============================================================
    def select_file(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 数据文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if not path:
            return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            self.students.clear()
            self.names.clear()
            for row in ws.iter_rows(values_only=True):
                if all(v is None for v in row):
                    continue
                if len(row) >= 2 and row[0] is not None and row[1] is not None:
                    sid = str(row[0]).strip()
                    name = str(row[1]).strip()
                    if name and name != "姓名":
                        self.students.append((sid, name))
                        self.names.append(name)
            if not self.names:
                messagebox.showerror("错误", "未找到有效数据。\n"
                                     "请确保 Excel 第一列为学号，第二列为姓名。")
                return
            fname = os.path.basename(path)
            self.file_label.config(text=f"✅ {fname}（{len(self.names)} 人）",
                                   fg=COLORS["text"])
            self.reset_state()
            self.btn_start.config(state="normal")
            self.stats_label.config(text=f"已加载 {len(self.names)} 名学生",
                                    fg=COLORS["success"])
            self.name_label.config(text="准备就绪", fg=COLORS["text"])
            self.sub_info.config(text="点击「开始」点名")
        except Exception as e:
            messagebox.showerror("读取失败", f"无法读取文件：\n{str(e)}")

    # ============================================================
    # 状态重置
    # ============================================================
    def reset_state(self):
        """切换文件或模式时重置所有状态"""
        if self.after_id:
            self.root.after_cancel(self.after_id)
            self.after_id = None
        self.is_running = False
        self.btn_start.config(state="normal" if self.names else "disabled",
                              bg=COLORS["success"])

        self.remaining = self.names.copy()
        random.shuffle(self.remaining)

        self.rotation_list = self.names.copy()
        random.shuffle(self.rotation_list)
        self.rotation_idx = 0

        self.round_num = 1
        self.current_result = None

    def on_mode_change(self, event=None):
        raw = self.mode_var.get()
        for key, label, desc in MODES:
            if label == raw:
                self.mode_key = key
                self.mode_desc.config(text=desc)
                break
        # 分组模式显示人数选择
        if self.mode_key == "group":
            self.group_frame.pack(side="left", padx=(14, 0))
        else:
            self.group_frame.pack_forget()

        self.reset_state()

        if self.names:
            n = len(self.names)
            self.name_label.config(text="准备就绪", fg=COLORS["text"])
            if self.mode_key == "exhaust":
                self.stats_label.config(text=f"不重复点名 | 剩余 {n} 人",
                                        fg=COLORS["text_sub"])
            elif self.mode_key == "rotation":
                self.stats_label.config(text=f"顺序轮流 | 共 {n} 人",
                                        fg=COLORS["text_sub"])
            elif self.mode_key == "group":
                self.stats_label.config(text=f"分组抽选 | 共 {n} 人",
                                        fg=COLORS["text_sub"])
            else:
                self.stats_label.config(text=f"自由点名 | 共 {n} 人",
                                        fg=COLORS["text_sub"])
            self.sub_info.config(text="")

    # ============================================================
    # 开始 / 停止（自由 + 不重复模式）
    # ============================================================
    def start(self):
        if not self.names or self.is_running:
            return

        # 不重复模式：检查是否点完
        if self.mode_key == "exhaust" and not self.remaining:
            if messagebox.askyesno("本轮结束",
                                   "所有同学都已被点到！\n是否开始新一轮？"):
                self.remaining = self.names.copy()
                random.shuffle(self.remaining)
                self.round_num += 1
                self.stats_label.config(
                    text=f"第 {self.round_num} 轮开始 | 剩余 {len(self.remaining)} 人",
                    fg=COLORS["warning"])
                self.name_label.config(text="🔄 新一轮")
                self.root.after(500, self.begin_roll)
            return

        if self.mode_key == "free" and not self.remaining:
            self.remaining = self.names.copy()

        self.begin_roll()

    def begin_roll(self):
        self.is_running = True
        self.btn_start.config(state="disabled", bg=COLORS["text_muted"])
        self.name_label.config(fg=COLORS["accent"])
        self.roll()

    def roll(self):
        if not self.is_running:
            return
        if self.mode_key == "exhaust" and not self.remaining:
            self.stop_roll()
            self.name_label.config(text="✅ 全部点完", fg=COLORS["text"])
            return
        student = random.choice(self.students)
        self.current_result = student
        sid, name = student
        self.name_label.config(text=name)
        self.sub_info.config(text=f"学号：{sid}")
        self.after_id = self.root.after(self.roll_interval, self.roll)

    def stop(self):
        if not self.is_running:
            return
        self.stop_roll()

        if not self.current_result:
            return

        sid, name = self.current_result

        # 不重复模式：从剩余中移除
        if self.mode_key == "exhaust" and name in self.remaining:
            self.remaining.remove(name)

        # 记录
        self.add_history(name, sid)

        # 高亮
        self.name_label.config(text=name, fg=COLORS["success"])
        self.sub_info.config(text=f"✅ 选中：{name}（{sid}）")

        # 更新统计
        self.update_stats()

    def stop_roll(self):
        self.is_running = False
        self.btn_start.config(state="normal", bg=COLORS["success"])
        if self.after_id:
            self.root.after_cancel(self.after_id)
            self.after_id = None

    # ============================================================
    # 分组抽选
    # ============================================================
    def pick_group(self):
        if not self.names:
            return
        if not hasattr(self, '_remaining_group'):
            self.reset_group_pool()

        available = self._remaining_group
        if not available:
            if messagebox.askyesno("本轮结束",
                                   "所有同学都已被分组！\n是否重新开始？"):
                self.reset_group_pool()
                available = self._remaining_group
            else:
                return

        k = min(self.group_size.get(), len(available))
        picked = random.sample(available, k)
        for p in picked:
            self._remaining_group.remove(p)

        # 显示
        def make_display(students, grp_size):
            if grp_size == 1:
                s = students[0]
                return s[1], f"学号：{s[0]}"
            names = "\n".join(f"  {s[1]}　{s[0]}" for s in students)
            return f"👥 共 {grp_size} 人", names

        main_text, sub_text = make_display(picked, k)

        # 用大号字体显示
        if k == 1:
            self.name_label.config(text=main_text, font=font(52, bold=True))
            self.sub_info.config(text=sub_text)
        else:
            self.name_label.config(text=main_text, font=font(28, bold=True))
            self.sub_info.config(text=sub_text)

        self.current_result = picked

        # 记录
        for s in picked:
            self.add_history(s[1], s[0])

        self.update_stats()

    def reset_group_pool(self):
        pool = self.names.copy()
        random.shuffle(pool)
        self._remaining_group = pool

    # ============================================================
    # 顺序轮流
    # ============================================================
    def next_rotation(self):
        if not self.names:
            return
        if not self.rotation_list:
            self.rotation_list = self.names.copy()
            random.shuffle(self.rotation_list)

        if self.rotation_idx >= len(self.rotation_list):
            # 一轮结束，重新打乱
            self.rotation_list = self.names.copy()
            random.shuffle(self.rotation_list)
            self.rotation_idx = 0

        name = self.rotation_list[self.rotation_idx]
        # 找学号
        sid = ""
        for s_id, s_name in self.students:
            if s_name == name:
                sid = s_id
                break

        self.rotation_idx += 1

        self.name_label.config(text=name, fg=COLORS["success"])
        self.sub_info.config(text=f"顺序：{self.rotation_idx}/{len(self.rotation_list)}")

        self.add_history(name, sid)
        self.update_stats()

    # ============================================================
    # 历史记录
    # ============================================================
    def add_history(self, name, sid=None):
        self.history.append((name, sid or ""))
        if len(self.history) > 100:
            self.history.pop(0)
        self.update_history_display()

    def clear_history(self):
        self.history.clear()
        self.update_history_display()

    def update_history_display(self):
        self.history_text.config(state="normal")
        self.history_text.delete("1.0", "end")
        if not self.history:
            self.history_text.insert("end", "（暂无记录）")
        else:
            tags = reversed(self.history[-50:])
            parts = []
            for i, (name, sid) in enumerate(tags):
                suffix = f" ({sid})" if sid else ""
                parts.append(f"{len(self.history) - i}. {name}{suffix}")
            self.history_text.insert("end", "  │  ".join(parts))
        self.history_text.config(state="disabled")

    # ============================================================
    # 统计更新
    # ============================================================
    def update_stats(self):
        total = len(self.names)
        self.total_picks = len(self.history)
        parts = []

        if self.mode_key == "free":
            parts.append(f"🎲 自由点名 | 共 {total} 人")
        elif self.mode_key == "exhaust":
            remain = len(self.remaining)
            parts.append(f"📋 不重复点名 | 第 {self.round_num} 轮")
            parts.append(f"剩余 {remain}/{total}")
        elif self.mode_key == "group":
            n = getattr(self, '_remaining_group', self.names.copy())
            remain = len(n)
            parts.append(f"👥 分组抽选 | 剩余 {remain}/{total}")
        elif self.mode_key == "rotation":
            parts.append(f"🔄 顺序轮流 | 进度 {self.rotation_idx}/{total}")

        parts.append(f"已点 {self.total_picks} 次")
        self.stats_label.config(text="  |  ".join(parts), fg=COLORS["text_sub"])

    # ============================================================
    # 启动入口
    # ============================================================
    def run(self):
        self.root.mainloop()


# ============================================================
# 主程序
# ============================================================
def main():
    root = tk.Tk()
    root.configure(bg=COLORS["bg"])
    app = NamePicker(root)
    app.run()


if __name__ == "__main__":
    main()
