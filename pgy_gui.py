#!/usr/bin/env python3
from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from pgy_opener import TOOL_VERSION, load_creators, open_creator_tabs

INITIAL_WINDOW_SIZE = "1100x600"
MINIMUM_WINDOW_SIZE = (1000, 540)


class PgyOpenerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"蒲公英博主打开工具 v{TOOL_VERSION}")
        self.root.geometry(INITIAL_WINDOW_SIZE)
        self.root.minsize(*MINIMUM_WINDOW_SIZE)

        self.messages: queue.Queue[tuple[str, str]] = queue.Queue()
        self.worker: threading.Thread | None = None

        self.excel_path = tk.StringVar()
        self.start = tk.StringVar(value="1")
        self.count = tk.StringVar(value="10")
        self.window_size = tk.StringVar(value="10")
        self.delay = tk.StringVar(value="0.3")
        self.dry_run = tk.BooleanVar(value=False)

        self._build_ui()
        self.root.after(100, self._drain_messages)

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=18)
        outer.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(outer, text="小红书蒲公英博主详情页批量打开", font=("", 18, "bold"))
        title.pack(anchor=tk.W)

        subtitle = ttk.Label(outer, text="选择 Excel 后点击开始。默认每 10 个链接新开一个浏览器窗口。")
        subtitle.pack(anchor=tk.W, pady=(4, 16))

        file_row = ttk.Frame(outer)
        file_row.pack(fill=tk.X)
        ttk.Label(file_row, text="Excel 文件").pack(anchor=tk.W)
        file_input_row = ttk.Frame(file_row)
        file_input_row.pack(fill=tk.X, pady=(6, 0))
        ttk.Entry(file_input_row, textvariable=self.excel_path).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(file_input_row, text="选择...", command=self.choose_file).pack(side=tk.LEFT, padx=(8, 0))

        options = ttk.LabelFrame(outer, text="打开设置", padding=12)
        options.pack(fill=tk.X, pady=16)

        self._number_field(options, "从第几个有效博主开始", self.start, 0)
        self._number_field(options, "打开几个", self.count, 1)
        self._number_field(options, "每几个新开一个窗口", self.window_size, 2)
        self._number_field(options, "打开间隔秒数", self.delay, 3)
        ttk.Checkbutton(options, text="只预览链接，不打开浏览器", variable=self.dry_run).grid(
            row=1, column=0, columnspan=2, sticky=tk.W, pady=(10, 0)
        )
        for column in (1, 3, 5, 7):
            options.columnconfigure(column, weight=1)

        actions = ttk.Frame(outer)
        actions.pack(fill=tk.X, pady=(0, 12))
        self.start_button = ttk.Button(actions, text="开始打开", command=self.start_run)
        self.start_button.pack(side=tk.LEFT)
        ttk.Button(actions, text="清空日志", command=self.clear_log).pack(side=tk.LEFT, padx=(8, 0))

        log_frame = ttk.LabelFrame(outer, text="运行日志", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log = tk.Text(log_frame, height=14, wrap=tk.WORD, state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        self.log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _number_field(self, parent: ttk.Frame, label: str, variable: tk.StringVar, column_pair: int) -> None:
        row = 0
        label_column = column_pair * 2
        entry_column = label_column + 1
        ttk.Label(parent, text=label).grid(row=row, column=label_column, sticky=tk.W, padx=(0, 8))
        ttk.Entry(parent, width=10, textvariable=variable).grid(row=row, column=entry_column, sticky=tk.W, padx=(0, 18))

    def choose_file(self) -> None:
        path = filedialog.askopenfilename(
            title="请选择包含小红书博主信息的 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.excel_path.set(path)

    def start_run(self) -> None:
        if self.worker and self.worker.is_alive():
            return

        try:
            options = self._read_options()
        except ValueError as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        self.start_button.configure(state=tk.DISABLED)
        self._append_log(f"蒲公英博主打开工具 v{TOOL_VERSION}")
        self._append_log(f"开始处理 Excel：从第 {options['start']} 个有效博主开始，最多打开 {options['count']} 个。")

        self.worker = threading.Thread(target=self._run_worker, args=(options,), daemon=True)
        self.worker.start()

    def _read_options(self) -> dict[str, object]:
        excel = self.excel_path.get().strip()
        if not excel:
            raise ValueError("请先选择 Excel 文件。")
        if not Path(excel).exists():
            raise ValueError("Excel 文件不存在，请重新选择。")

        start = _positive_int(self.start.get(), "从第几个有效博主开始")
        count = _positive_int(self.count.get(), "打开几个")
        window_size = _positive_int(self.window_size.get(), "每几个新开一个窗口")
        delay = _non_negative_float(self.delay.get(), "打开间隔秒数")

        return {
            "excel": excel,
            "start": start,
            "count": count,
            "window_size": window_size,
            "delay": delay,
            "dry_run": self.dry_run.get(),
        }

    def _run_worker(self, options: dict[str, object]) -> None:
        try:
            creators = load_creators(
                options["excel"],
                start=options["start"],
                count=options["count"],
                progress=lambda message: self.messages.put(("log", message)),
            )
            open_creator_tabs(
                creators,
                delay=options["delay"],
                dry_run=options["dry_run"],
                window_size=options["window_size"],
                progress=lambda message: self.messages.put(("log", message)),
            )

            failed = [creator for creator in creators if creator.error]
            if failed:
                self.messages.put(("log", ""))
                self.messages.put(("log", "以下博主未能打开："))
                for creator in failed:
                    self.messages.put(("log", f"- 第 {creator.row_number} 行 {creator.name}: {creator.error}"))
            self.messages.put(("done", "执行完成。"))
        except Exception as exc:  # noqa: BLE001 - show friendly GUI error.
            self.messages.put(("error", str(exc)))

    def _drain_messages(self) -> None:
        while True:
            try:
                kind, message = self.messages.get_nowait()
            except queue.Empty:
                break

            if kind == "log":
                self._append_log(message)
            elif kind == "done":
                self._append_log(message)
                self.start_button.configure(state=tk.NORMAL)
                messagebox.showinfo("完成", message)
            elif kind == "error":
                self._append_log(f"运行失败：{message}")
                self.start_button.configure(state=tk.NORMAL)
                messagebox.showerror("运行失败", message)

        self.root.after(100, self._drain_messages)

    def _append_log(self, message: str) -> None:
        self.log.configure(state=tk.NORMAL)
        self.log.insert(tk.END, message + "\n")
        self.log.see(tk.END)
        self.log.configure(state=tk.DISABLED)

    def clear_log(self) -> None:
        self.log.configure(state=tk.NORMAL)
        self.log.delete("1.0", tk.END)
        self.log.configure(state=tk.DISABLED)


def _positive_int(value: str, label: str) -> int:
    try:
        parsed = int(value)
    except ValueError as exc:
        raise ValueError(f"{label} 必须是整数。") from exc
    if parsed < 1:
        raise ValueError(f"{label} 必须大于等于 1。")
    return parsed


def _non_negative_float(value: str, label: str) -> float:
    try:
        parsed = float(value)
    except ValueError as exc:
        raise ValueError(f"{label} 必须是数字。") from exc
    if parsed < 0:
        raise ValueError(f"{label} 不能小于 0。")
    return parsed


def main() -> None:
    root = tk.Tk()
    app = PgyOpenerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
