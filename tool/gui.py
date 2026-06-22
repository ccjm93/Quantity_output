# -*- coding: utf-8 -*-
"""수량산출서 출력 자동화 도구 - GUI (단일 PDF 병합판, Tkinter 내장).

직원용: 수량산출서 '루트 폴더' 선택 → [PDF 생성 시작] → 간지 포함 단일 PDF + AI 사후검토.
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

# API 키는 이 프로그램 폴더의 .env 에만 저장(프로그램 삭제 시 키도 사라짐).
ENV_PATH = os.path.join(HERE, ".env")
KEY_NAME = "GEMINI_API_KEY"


def read_api_key() -> str:
    if not os.path.isfile(ENV_PATH):
        return ""
    try:
        with open(ENV_PATH, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith(KEY_NAME + "="):
                    return line.split("=", 1)[1].strip().strip('"').strip("'")
    except Exception:
        pass
    return ""


def _read_env_lines() -> list:
    if not os.path.isfile(ENV_PATH):
        return []
    try:
        with open(ENV_PATH, "r", encoding="utf-8") as f:
            return f.read().splitlines()
    except Exception:
        return []


def save_api_key(key: str) -> None:
    key = (key or "").strip()
    lines = _read_env_lines()
    out, found = [], False
    for line in lines:
        if line.strip().startswith(KEY_NAME + "="):
            out.append(f"{KEY_NAME}={key}")
            found = True
        else:
            out.append(line)
    if not found:
        out.append(f"{KEY_NAME}={key}")
    with open(ENV_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(out) + "\n")
    if key:
        os.environ[KEY_NAME] = key
    else:
        os.environ.pop(KEY_NAME, None)


def delete_api_key() -> None:
    lines = _read_env_lines()
    out = [ln for ln in lines if not ln.strip().startswith(KEY_NAME + "=")]
    try:
        with open(ENV_PATH, "w", encoding="utf-8") as f:
            f.write(("\n".join(out) + "\n") if out else "")
    except Exception:
        pass
    os.environ.pop(KEY_NAME, None)


def _ai_status():
    # 무거운 toolruntime(→win32com/excel_app) 대신 .env 키만 환경에 반영하고
    # review 만 가볍게 사용한다. 실패 시 라벨에 긴 트레이스가 노출되지 않도록 요약한다.
    try:
        key = read_api_key()
        if key and not os.environ.get(KEY_NAME):
            os.environ[KEY_NAME] = key
        import review
        return review.is_available(False)
    except Exception as e:
        return False, type(e).__name__


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

        frm2 = tk.LabelFrame(root, text=" 출력 PDF (비우면 루트\\_output\\<폴더명>_통합.pdf) ")
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
        self.ai_lbl = tk.Label(frm3, text="")
        self.ai_lbl.pack(side="left", padx=10)
        tk.Button(frm3, text="API 키 설정",
                  command=self.open_apikey_dialog).pack(side="right", padx=6)
        self.refresh_ai_status()

        frm4 = tk.Frame(root)
        frm4.pack(fill="x", padx=12, pady=(6, 2))
        self.run_btn = tk.Button(frm4, text="PDF 생성 시작", height=2, bg="#2d7",
                                 fg="white", font=("맑은 고딕", 11, "bold"),
                                 command=self.start)
        self.run_btn.pack(side="left")
        self.review_btn = tk.Button(frm4, text="AI 출력물 검토", height=2,
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
        self.root.after(500, self._first_run_prompt)

    # --- API 키 ---
    def refresh_ai_status(self):
        ok, reason = _ai_status()
        self.ai_lbl.config(text=("AI: 사용 가능" if ok else f"AI: 미설정 ({reason})"),
                           fg=("#197" if ok else "#a60"))

    def _first_run_prompt(self):
        if read_api_key():
            return
        msg = ("이 프로그램은 키 없이도 PDF 생성이 정상 동작합니다.\n\n"
               "최종 PDF의 'AI 사후검토(이상 탐지)'를 쓰려면 '개인' Google AI Studio "
               "API 키가 필요합니다. 키는 이 PC의 프로그램 폴더에만 저장되며 다른 사람과 "
               "공유되지 않습니다.\n\n지금 개인 API 키를 입력하시겠습니까?")
        if messagebox.askyesno("개인 API 키 입력 (최초 실행)", msg):
            self.open_apikey_dialog()

    def show_apikey_help(self):
        try:
            import toolruntime
            messagebox.showinfo("API 키 발급 방법", toolruntime.APIKEY_HELP)
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def open_apikey_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("API 키 설정")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)
        tk.Label(dlg, text="Google AI Studio API 키",
                 font=("맑은 고딕", 11, "bold")).pack(anchor="w", padx=14, pady=(12, 2))
        tk.Label(dlg, fg="#555", justify="left",
                 text=("최종 PDF 사후검토 품질 향상용(선택).\n"
                       "키는 이 프로그램 폴더(tool\\.env)에만 저장되며,\n"
                       "프로그램을 삭제하면 키도 함께 사라집니다.")
                 ).pack(anchor="w", padx=14)
        key_var = tk.StringVar(value=read_api_key())
        show_var = tk.BooleanVar(value=False)
        row = tk.Frame(dlg)
        row.pack(fill="x", padx=14, pady=(10, 2))
        ent = tk.Entry(row, textvariable=key_var, width=46, show="*")
        ent.pack(side="left", fill="x", expand=True)

        def toggle():
            ent.config(show="" if show_var.get() else "*")
        tk.Checkbutton(dlg, text="키 표시", variable=show_var,
                       command=toggle).pack(anchor="w", padx=12)

        btns = tk.Frame(dlg)
        btns.pack(fill="x", padx=14, pady=12)

        def do_save():
            k = key_var.get().strip()
            if not k:
                messagebox.showwarning("경고", "키를 입력하세요.", parent=dlg)
                return
            save_api_key(k)
            self.refresh_ai_status()
            messagebox.showinfo("완료", "API 키가 저장되었습니다.", parent=dlg)
            dlg.destroy()

        def do_delete():
            if not read_api_key():
                messagebox.showinfo("안내", "저장된 키가 없습니다.", parent=dlg)
                return
            if messagebox.askyesno("확인", "저장된 API 키를 삭제할까요?", parent=dlg):
                delete_api_key()
                self.refresh_ai_status()
                dlg.destroy()

        tk.Button(btns, text="저장", width=10, bg="#2d7", fg="white",
                  command=do_save).pack(side="left")
        tk.Button(btns, text="키 삭제", width=10, command=do_delete).pack(side="left", padx=6)
        tk.Button(btns, text="발급 방법", width=10,
                  command=self.show_apikey_help).pack(side="left")
        tk.Button(btns, text="닫기", width=8, command=dlg.destroy).pack(side="right")
        ent.focus_set()

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
            return os.path.abspath(self.out_var.get().strip())
        # toolruntime 과 동일하게 설정값에서 출력 폴더명/접미사를 읽어 경로를 맞춘다.
        cfg = load_settings(target)
        name = os.path.basename(target.rstrip("\\/")) + cfg["merged_suffix"]
        return os.path.join(target, cfg["output_dir_name"], name)

    def start(self):
        target = self.path_var.get().strip()
        if not target or not os.path.isdir(target):
            messagebox.showwarning("경고", "수량산출서 '루트 폴더'를 선택하세요.")
            return
        if not self.backup.get():
            if not messagebox.askyesno(
                    "주의", "원본 백업 없이 진행하면 원본 파일이 변경됩니다.\n계속할까요?"):
                return
        # PDF 생성만 수행 (AI 검토는 완료 후 'AI 출력물 검토' 버튼으로 선택 실행)
        cmd = [sys.executable, "-u", os.path.join(HERE, "toolruntime.py"), target]
        if self.out_var.get().strip():
            cmd += ["--out", self.out_var.get().strip()]
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
        ok, reason = _ai_status()
        if not ok and not messagebox.askyesno(
                "안내", f"AI 키가 없어 규칙 기반 검토만 됩니다 ({reason}).\n계속할까요?"):
            return
        cmd = [sys.executable, "-u", os.path.join(HERE, "toolruntime.py"),
               "--review", self._last_pdf]
        self._busy("AI 출력물 검토 시작...\n\n")
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
                            self._log("필요하면 [AI 출력물 검토]를 눌러 추가 검토하세요.\n")
                    elif code == 0 and mode == "review":
                        self._log("\n=== AI 출력물 검토 완료 ===\n")
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
