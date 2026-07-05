# -*- coding: utf-8 -*-
"""Excel COM 인스턴스 수명 관리.

- 사용자의 기존 Excel과 분리된 전용 인스턴스(DispatchEx)를 띄운다.
- 매크로 비활성, 경고/이벤트 차단 등 배치 처리용 전역 옵션을 1회 설정.
- 예외가 나도 반드시 Quit + 고아 프로세스 정리.
"""
from __future__ import annotations

import ctypes
import gc
import subprocess

import pythoncom
import win32com.client as win32

from config import MSO_AUTOMATION_SECURITY_FORCE_DISABLE


def _excel_pids() -> set[int]:
    """현재 살아있는 EXCEL.EXE PID 집합 (best-effort)."""
    try:
        out = subprocess.check_output(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            stderr=subprocess.DEVNULL, text=True,
        )
    except Exception:
        return set()
    pids = set()
    for line in out.splitlines():
        parts = [p.strip('"') for p in line.split('","')]
        if len(parts) >= 2 and parts[1].strip('"').isdigit():
            pids.add(int(parts[1].strip('"')))
    return pids


class ExcelApp:
    """with 블록으로 사용. 우리가 생성한 인스턴스만 정리한다."""

    def __init__(self):
        self.app = None
        self._own_pids: set[int] = set()

    def __enter__(self) -> "ExcelApp":
        pythoncom.CoInitialize()
        before = _excel_pids()
        self.app = win32.DispatchEx("Excel.Application")
        # 자기 인스턴스의 PID를 창 핸들에서 직접 얻는다(동시에 뜬 다른 Excel 오인 방지).
        pid = self._pid_from_hwnd()
        if pid:
            self._own_pids = {pid}
        else:
            self._own_pids = _excel_pids() - before  # 차선책: 새로 뜬 프로세스 차집합

        app = self.app
        app.Visible = False
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        app.EnableEvents = False
        app.AskToUpdateLinks = False
        try:
            app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        except Exception:
            pass
        # 계산 모드는 자동(기본값) 유지 — 열 때 재계산되어 최신 값으로 출력된다.
        return self

    def _pid_from_hwnd(self):
        try:
            hwnd = int(self.app.Hwnd)
            pid = ctypes.c_ulong(0)
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            return int(pid.value) or None
        except Exception:
            return None

    def __exit__(self, exc_type, exc, tb):
        try:
            if self.app is not None:
                try:
                    self.app.DisplayAlerts = False
                except Exception:
                    pass
                self.app.Quit()
        except Exception:
            pass
        finally:
            self.app = None
            gc.collect()
            # Quit 후에도 남아있는 우리 PID만 강제 종료
            still = _excel_pids() & self._own_pids
            for pid in still:
                try:
                    subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                                   stdout=subprocess.DEVNULL,
                                   stderr=subprocess.DEVNULL)
                except Exception:
                    pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        return False  # 예외 전파
