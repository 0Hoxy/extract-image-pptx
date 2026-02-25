import threading
import time
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk
from pathlib import Path

from src.extractor import extract_images_from_pptx


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PPTX 이미지 추출기")
        self.geometry("600x500")
        self.resizable(False, False)

        self._build_ui()

    def _build_ui(self):
        # --- PPTX 파일 선택 ---
        frame_file = ttk.LabelFrame(self, text="PPTX 파일", padding=10)
        frame_file.pack(fill="x", padx=15, pady=(15, 5))

        self.var_pptx = tk.StringVar()
        ttk.Entry(frame_file, textvariable=self.var_pptx, state="readonly").pack(side="left", fill="x", expand=True)
        ttk.Button(frame_file, text="찾아보기", command=self._browse_pptx).pack(side="right", padx=(10, 0))

        # --- 출력 폴더 선택 ---
        frame_output = ttk.LabelFrame(self, text="출력 폴더", padding=10)
        frame_output.pack(fill="x", padx=15, pady=5)

        self.var_output = tk.StringVar(value="./output")
        ttk.Entry(frame_output, textvariable=self.var_output, state="readonly").pack(side="left", fill="x", expand=True)
        ttk.Button(frame_output, text="찾아보기", command=self._browse_output).pack(side="right", padx=(10, 0))

        # --- 실행 버튼 ---
        self.btn_run = ttk.Button(self, text="추출 시작", command=self._run)
        self.btn_run.pack(pady=10)

        # --- 진행률 ---
        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x", padx=15)

        # --- 로그 영역 ---
        frame_log = ttk.LabelFrame(self, text="로그", padding=10)
        frame_log.pack(fill="both", expand=True, padx=15, pady=(5, 15))

        self.log = scrolledtext.ScrolledText(frame_log, height=12, state="disabled", font=("Courier", 10))
        self.log.pack(fill="both", expand=True)

    def _browse_pptx(self):
        path = filedialog.askopenfilename(
            title="PPTX 파일 선택",
            filetypes=[("PowerPoint 파일", "*.pptx"), ("모든 파일", "*.*")],
        )
        if path:
            self.var_pptx.set(path)

    def _browse_output(self):
        path = filedialog.askdirectory(title="출력 폴더 선택")
        if path:
            self.var_output.set(path)

    def _log(self, msg: str):
        """스레드 안전한 로그 출력."""
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _run(self):
        pptx_path = self.var_pptx.get()
        output_dir = self.var_output.get()

        if not pptx_path:
            self._log("[오류] PPTX 파일을 선택해주세요.")
            return

        if not Path(pptx_path).exists():
            self._log(f"[오류] 파일을 찾을 수 없습니다: {pptx_path}")
            return

        self.btn_run.configure(state="disabled")
        self.progress["value"] = 0

        # 별도 스레드에서 추출 실행 (GUI 프리징 방지)
        thread = threading.Thread(target=self._extract, args=(pptx_path, output_dir), daemon=True)
        thread.start()

    def _extract(self, pptx_path: str, output_dir: str):
        import sys

        start = time.time()

        # print 출력 리다이렉트
        original_stdout = sys.stdout
        original_stderr = sys.stderr

        class ThreadSafeLog:
            def __init__(self, app):
                self.app = app

            def write(self, msg):
                if msg.strip():
                    self.app.after(0, self.app._log, msg.strip())

            def flush(self):
                pass

        sys.stdout = ThreadSafeLog(self)
        sys.stderr = ThreadSafeLog(self)

        try:
            extract_images_from_pptx(pptx_path, output_dir)
            elapsed = time.time() - start
            self.after(0, self._log, f"\n소요 시간: {elapsed:.2f}초")
        except Exception as e:
            self.after(0, self._log, f"[오류] {e}")
        finally:
            sys.stdout = original_stdout
            sys.stderr = original_stderr
            self.after(0, lambda: self.btn_run.configure(state="normal"))
            self.after(0, lambda: self.progress.configure(value=100))


if __name__ == "__main__":
    app = App()
    app.mainloop()
