# -*- coding: utf-8 -*-
"""수량산출서 출력 자동화 도구 - GUI (단일 PDF 병합판, Tkinter 내장).

직원용: 수량산출서 '루트 폴더' 선택 → [PDF 생성 시작] → 간지 포함 단일 PDF + 오류 검토.
처리는 toolruntime.py 를 별도 프로세스로 실행하고 로그를 실시간 표시한다.
"""
from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from config import load_settings  # noqa: E402


class App:
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self._last_out = None
        root.title("수량산출서 출력 자동화 도구")
        root.geometry("780x600")
        root.minsize(700, 520)
        pad = {"padx": 8, "pady": 4}

        tk.Label(root, text="수량산출서 출력 자동화 도구",
                 font=("맑은 고딕", 15, "bold")).pack(anchor="w", padx=12, pady=(12, 2))
        tk.Label(root, fg="#555",
                 text="루트 폴더의 모든 시트를 폴더·파일 순서로 간지 포함 단일 PDF로 병합합니다."
                 ).pack(anchor="w", padx=12)

        frm = tk.LabelFrame(root, text=" 수량산출서 루트 폴더 ")
        frm.pack(fill="x", padx=12, pady=(10, 4))
        self.path_var = tk.StringVar()
        tk.Entry(frm, textvariable=self.path_var).grid(row=0, column=0, sticky="we", **pad)
        tk.Button(frm, text="폴더 선택", width=12,
                  command=self.pick_folder).grid(row=0, column=1, **pad)
        frm.columnconfigure(0, weight=1)

        frm2 = tk.LabelFrame(root, text=" 출력 PDF (비우면 선택 폴더\\수량산출서 output.pdf) ")
        frm2.pack(fill="x", padx=12, pady=4)
        self.out_var = tk.StringVar()
        tk.Entry(frm2, textvariable=self.out_var).grid(row=0, column=0, sticky="we", **pad)
        tk.Button(frm2, text="선택", width=12,
                  command=self.pick_out).grid(row=0, column=1, **pad)
        frm2.columnconfigure(0, weight=1)

        frm3 = tk.Frame(root)
        frm3.pack(fill="x", padx=12, pady=4)
        self.backup = tk.BooleanVar(value=True)
        tk.Checkbutton(frm3, text="원본 백업(권장)", variable=self.backup).pack(side="left")
        tk.Label(frm3, fg="#555",
                 text="오류 검토는 프로그램 내부 로직으로 수행합니다.").pack(side="left", padx=10)

        frm4 = tk.Frame(root)
        frm4.pack(fill="x", padx=12, pady=(6, 2))
        self.run_btn = tk.Button(frm4, text="PDF 생성 시작", height=2, bg="#2d7",
                                 fg="white", font=("맑은 고딕", 11, "bold"),
                                 command=self.start)
        self.run_btn.pack(side="left")
        self.review_btn = tk.Button(frm4, text="출력물 오류 검토", height=2,
                                    state="disabled", command=self.start_review)
        self.review_btn.pack(side="left", padx=8)
        self.open_btn = tk.Button(frm4, text="출력 폴더 열기", height=2,
                                  state="disabled", command=self.open_output)
        self.open_btn.pack(side="left", padx=8)
        self.prog = ttk.Progressbar(frm4, mode="indeterminate")
        self.prog.pack(side="left", fill="x", expand=True, padx=8)
        self._last_pdf = None

        self.log = scrolledtext.ScrolledText(root, height=15, font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        self._log("준비 완료. 수량산출서 '루트 폴더'를 선택한 뒤 [PDF 생성 시작]을 누르세요.\n")

        self.root.after(120, self._drain)

    # --- 실행 ---
    def pick_folder(self):
        p = filedialog.askdirectory(title="수량산출서 루트 폴더 선택")
        if p:
            self.path_var.set(p)

    def pick_out(self):
        p = filedialog.asksaveasfilename(title="출력 PDF 저장 위치",
                                         defaultextension=".pdf",
                                         filetypes=[("PDF", "*.pdf")])
        if p:
            self.out_var.set(p)

    def _log(self, text):
        self.log.insert("end", text)
        self.log.see("end")

    def _expected_pdf(self, target):
        if self.out_var.get().strip():
            return self._normalize_pdf_path(self.out_var.get().strip())
        # toolruntime 과 동일하게 설정값에서 출력 폴더명/접미사를 읽어 경로를 맞춘다.
        cfg = load_settings(target)
        return os.path.join(target, cfg.get("output_pdf_name", "수량산출서 output.pdf"))

    def _normalize_pdf_path(self, path):
        path = os.path.abspath(path)
        if os.path.splitext(path)[1].lower() != ".pdf":
            path += ".pdf"
        return path

    def start(self):
        target = self.path_var.get().strip()
        if not target or not os.path.isdir(target):
            messagebox.showwarning("경고", "수량산출서 '루트 폴더'를 선택하세요.")
            return
        if not self.backup.get():
            if not messagebox.askyesno(
                    "주의", "원본 백업 없이 진행하면 원본 파일이 변경됩니다.\n계속할까요?"):
                return
        # PDF 생성만 수행 (오류 검토는 완료 후 '출력물 오류 검토' 버튼으로 선택 실행)
        cmd = [sys.executable, "-u", os.path.join(HERE, "toolruntime.py"), target]
        if self.out_var.get().strip():
            cmd += ["--out", self._normalize_pdf_path(self.out_var.get().strip())]
        if not self.backup.get():
            cmd += ["--no-backup"]
        self._last_pdf = self._expected_pdf(target)
        self._last_out = os.path.dirname(self._last_pdf)
        self._busy("PDF 생성 시작...\n\n")
        threading.Thread(target=self._worker, args=(cmd, "generate"), daemon=True).start()

    def start_review(self):
        if not self._last_pdf or not os.path.isfile(self._last_pdf):
            messagebox.showinfo("안내", "먼저 PDF를 생성하세요.")
            return
        cmd = [sys.executable, "-u", os.path.join(HERE, "toolruntime.py"),
               "--review", self._last_pdf]
        self._busy("출력물 오류 검토 시작...\n\n")
        threading.Thread(target=self._worker, args=(cmd, "review"), daemon=True).start()

    def _busy(self, msg):
        self.run_btn.config(state="disabled")
        self.review_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.prog.start(12)
        self.log.delete("1.0", "end")
        self._log(msg)

    def _worker(self, cmd, mode):
        try:
            env = dict(os.environ)
            env["PYTHONUTF8"] = "1"
            proc = subprocess.Popen(cmd, stdout=subprocess.PIPE,
                                    stderr=subprocess.STDOUT, text=True,
                                    encoding="utf-8", errors="replace", bufsize=1, env=env)
            for line in proc.stdout:
                self.q.put(("log", line))
            proc.wait()
            self.q.put(("done", (mode, proc.returncode)))
        except Exception as e:
            self.q.put(("log", f"[오류] {e}\n"))
            self.q.put(("done", (mode, -1)))

    def _drain(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self._log(payload)
                elif kind == "done":
                    mode, code = payload
                    self.prog.stop()
                    self.run_btn.config(state="normal")
                    pdf_ok = bool(self._last_pdf and os.path.isfile(self._last_pdf))
                    out_ok = bool(self._last_out and os.path.isdir(self._last_out))
                    self.open_btn.config(state="normal" if out_ok else "disabled")
                    self.review_btn.config(state="normal" if pdf_ok else "disabled")
                    if code == 0 and mode == "generate":
                        self._log("\n=== PDF 생성 완료 ===\n")
                        if pdf_ok:
                            self._log("필요하면 [출력물 오류 검토]를 눌러 추가 검토하세요.\n")
                    elif code == 0 and mode == "review":
                        self._log("\n=== 출력물 오류 검토 완료 ===\n")
                    else:
                        self._log(f"\n=== 종료(코드 {code}) ===\n")
        except queue.Empty:
            pass
        self.root.after(120, self._drain)

    def open_output(self):
        if self._last_out and os.path.isdir(self._last_out):
            os.startfile(self._last_out)
        else:
            messagebox.showinfo("안내", "아직 출력 폴더가 없습니다.")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
