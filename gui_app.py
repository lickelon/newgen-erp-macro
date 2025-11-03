"""
부양가족 대량 입력 GUI 애플리케이션
CustomTkinter 기반
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import queue
import sys
from pathlib import Path
from bulk_dependent_input import BulkDependentInput
# CustomTkinter 설정
ctk.set_appearance_mode("dark")  # "light", "dark", "system"
ctk.set_default_color_theme("blue")  # "blue", "green", "dark-blue"


class StopInfoWindow(ctk.CTkToplevel):
    """중지 안내 플로팅 창 (클릭 불가, 안내만)"""

    def __init__(self, parent):
        super().__init__(parent)

        # 창 설정
        self.title("중지 방법")
        self.geometry("280x140")
        self.resizable(False, False)

        # 항상 위에 표시
        self.attributes('-topmost', True)

        # 창 투명도
        self.attributes('-alpha', 0.9)

        # 창 닫기 버튼 비활성화
        self.protocol("WM_DELETE_WINDOW", lambda: None)

        # 내용
        label1 = ctk.CTkLabel(
            self,
            text="⚡ 자동화 실행 중",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#2ecc71"
        )
        label1.pack(pady=(15, 5))

        label2 = ctk.CTkLabel(
            self,
            text="중지하려면:",
            font=ctk.CTkFont(size=12)
        )
        label2.pack(pady=(10, 5))

        label3 = ctk.CTkLabel(
            self,
            text="Pause 키를 3번 누르세요",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color="#e74c3c"
        )
        label3.pack(pady=5)

        label4 = ctk.CTkLabel(
            self,
            text="(2초 이내에)",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        )
        label4.pack(pady=(0, 10))

        # 드래그 가능하게
        for widget in [self, label1, label2, label3, label4]:
            widget.bind('<Button-1>', self.start_move)
            widget.bind('<B1-Motion>', self.on_move)

        self._drag_start_x = 0
        self._drag_start_y = 0

    def start_move(self, event):
        self._drag_start_x = event.x
        self._drag_start_y = event.y

    def on_move(self, event):
        x = self.winfo_x() + (event.x - self._drag_start_x)
        y = self.winfo_y() + (event.y - self._drag_start_y)
        self.geometry(f"+{x}+{y}")


class LogRedirector:
    """로그를 GUI 텍스트박스로 리디렉션"""

    def __init__(self, text_widget, queue):
        self.text_widget = text_widget
        self.queue = queue

    def write(self, text):
        if text.strip():  # 빈 줄 제외
            self.queue.put(text)

    def flush(self):
        pass


class BulkInputGUI(ctk.CTk):
    """부양가족 대량 입력 GUI"""

    def __init__(self):
        super().__init__()

        # 윈도우 설정
        self.title("부양가족 대량 입력 자동화")
        self.geometry("900x700")

        # 변수
        self.csv_path = ctk.StringVar()
        self.employee_count = ctk.StringVar()
        self.global_delay = ctk.StringVar(value="1.0")
        self.dry_run = ctk.BooleanVar(value=False)
        self.is_running = False
        self.log_queue = queue.Queue()
        self.bulk_automation = None  # BulkDependentInput 인스턴스
        self.stop_window = None  # 중지 전용 플로팅 창

        # UI 생성
        self.create_widgets()

        # 로그 업데이트 타이머
        self.update_log()

        # 종료 시 핫키 해제
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        """UI 위젯 생성"""

        # ===== 헤더 =====
        header = ctk.CTkLabel(
            self,
            text="🤖 부양가족 대량 입력 자동화",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        header.pack(pady=20)

        # 중지 안내
        hotkey_label = ctk.CTkLabel(
            self,
            text="💡 실행 중 'Pause 키 3번'으로 중지 (안내 창 표시됨)",
            font=ctk.CTkFont(size=12),
            text_color="orange"
        )
        hotkey_label.pack(pady=(0, 10))

        # ===== 파일 선택 프레임 =====
        file_frame = ctk.CTkFrame(self)
        file_frame.pack(padx=20, pady=10, fill="x")

        ctk.CTkLabel(
            file_frame,
            text="CSV 파일:",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(side="left", padx=10, pady=10)

        self.file_entry = ctk.CTkEntry(
            file_frame,
            textvariable=self.csv_path,
            width=400,
            placeholder_text="CSV 파일을 선택하세요..."
        )
        self.file_entry.pack(side="left", padx=10, pady=10, fill="x", expand=True)

        self.browse_btn = ctk.CTkButton(
            file_frame,
            text="찾아보기",
            command=self.browse_file,
            width=100
        )
        self.browse_btn.pack(side="left", padx=10, pady=10)

        # ===== 옵션 프레임 =====
        options_frame = ctk.CTkFrame(self)
        options_frame.pack(padx=20, pady=10, fill="x")

        # 사원 수 입력
        count_frame = ctk.CTkFrame(options_frame)
        count_frame.pack(side="left", padx=10, pady=10)

        ctk.CTkLabel(
            count_frame,
            text="처리할 사원 수:",
            font=ctk.CTkFont(size=13)
        ).pack(side="left", padx=5)

        self.count_entry = ctk.CTkEntry(
            count_frame,
            textvariable=self.employee_count,
            width=100,
            placeholder_text="전체"
        )
        self.count_entry.pack(side="left", padx=5)

        ctk.CTkLabel(
            count_frame,
            text="(비어있으면 전체 처리)",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        ).pack(side="left", padx=5)

        # Delay 설정
        delay_frame = ctk.CTkFrame(options_frame)
        delay_frame.pack(side="left", padx=10, pady=10)

        ctk.CTkLabel(
            delay_frame,
            text="입력 속도:",
            font=ctk.CTkFont(size=13)
        ).pack(side="left", padx=5)

        self.delay_entry = ctk.CTkEntry(
            delay_frame,
            textvariable=self.global_delay,
            width=60,
            placeholder_text="1.0"
        )
        self.delay_entry.pack(side="left", padx=5)

        ctk.CTkLabel(
            delay_frame,
            text="(0.5~2.0배)",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        ).pack(side="left", padx=5)

        # Dry run 체크박스
        self.dry_run_check = ctk.CTkCheckBox(
            options_frame,
            text="Dry Run (실제 입력 안함)",
            variable=self.dry_run,
            font=ctk.CTkFont(size=13)
        )
        self.dry_run_check.pack(side="left", padx=20, pady=10)

        # ===== 실행 버튼 =====
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(padx=20, pady=10, fill="x")

        self.start_btn = ctk.CTkButton(
            button_frame,
            text="▶ 시작",
            command=self.start_automation,
            font=ctk.CTkFont(size=16, weight="bold"),
            height=40,
            fg_color="#2ecc71",
            hover_color="#27ae60"
        )
        self.start_btn.pack(side="left", padx=10, pady=10, fill="x", expand=True)

        self.stop_btn = ctk.CTkButton(
            button_frame,
            text="■ 중지",
            command=self.stop_automation,
            font=ctk.CTkFont(size=16, weight="bold"),
            height=40,
            fg_color="#e74c3c",
            hover_color="#c0392b",
            state="disabled"
        )
        self.stop_btn.pack(side="left", padx=10, pady=10, fill="x", expand=True)

        # ===== 진행 상태 =====
        progress_frame = ctk.CTkFrame(self)
        progress_frame.pack(padx=20, pady=10, fill="x")

        self.progress_label = ctk.CTkLabel(
            progress_frame,
            text="대기 중...",
            font=ctk.CTkFont(size=13)
        )
        self.progress_label.pack(pady=5)

        self.progress_bar = ctk.CTkProgressBar(
            progress_frame,
            width=800
        )
        self.progress_bar.pack(pady=5)
        self.progress_bar.set(0)

        # ===== 로그 출력 =====
        log_frame = ctk.CTkFrame(self)
        log_frame.pack(padx=20, pady=10, fill="both", expand=True)

        ctk.CTkLabel(
            log_frame,
            text="📋 실행 로그",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=5)

        self.log_text = ctk.CTkTextbox(
            log_frame,
            font=ctk.CTkFont(family="Consolas", size=11),
            wrap="word"
        )
        self.log_text.pack(padx=10, pady=10, fill="both", expand=True)

    def browse_file(self):
        """CSV 파일 선택"""
        filename = filedialog.askopenfilename(
            title="CSV 파일 선택",
            filetypes=[("CSV 파일", "*.csv"), ("모든 파일", "*.*")]
        )
        if filename:
            self.csv_path.set(filename)
            self.log(f"파일 선택: {filename}")

    def log(self, message):
        """로그 메시지 추가"""
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")

    def update_log(self):
        """큐에서 로그 메시지 가져와서 표시"""
        try:
            # 배치로 처리 (최대 50개)
            count = 0
            while count < 50:
                message = self.log_queue.get_nowait()
                self.log(message.strip())
                count += 1
        except queue.Empty:
            pass

        # 300ms마다 업데이트 (최적화: 100ms → 300ms)
        self.after(300, self.update_log)

    def start_automation(self):
        """자동화 시작"""
        # 유효성 검사
        csv_file = self.csv_path.get()
        if not csv_file:
            messagebox.showerror("오류", "CSV 파일을 선택하세요.")
            return

        if not Path(csv_file).exists():
            messagebox.showerror("오류", f"파일을 찾을 수 없습니다:\n{csv_file}")
            return

        # 사원 수 파싱
        count_str = self.employee_count.get().strip()
        count = None
        if count_str:
            try:
                count = int(count_str)
                if count <= 0:
                    messagebox.showerror("오류", "사원 수는 양수여야 합니다.")
                    return
            except ValueError:
                messagebox.showerror("오류", "사원 수는 숫자여야 합니다.")
                return

        # Delay 파싱
        delay_str = self.global_delay.get().strip()
        delay = 1.0
        if delay_str:
            try:
                delay = float(delay_str)
                if delay < 0.5 or delay > 2.0:
                    messagebox.showerror("오류", "입력 속도는 0.5~2.0 범위여야 합니다.")
                    return
            except ValueError:
                messagebox.showerror("오류", "입력 속도는 숫자여야 합니다.")
                return

        # UI 상태 변경
        self.is_running = True
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.browse_btn.configure(state="disabled")
        self.count_entry.configure(state="disabled")
        self.delay_entry.configure(state="disabled")
        self.dry_run_check.configure(state="disabled")

        self.log_text.delete("1.0", "end")
        self.log("=" * 50)
        self.log("자동화 시작")
        self.log("=" * 50)
        self.log("💡 중지하려면: Pause 키를 3번 누르세요 (2초 이내)")
        self.log("=" * 50)
        self.progress_label.configure(text="실행 중...")
        self.progress_bar.set(0)

        # 중지 안내 플로팅 창 열기
        self.stop_window = StopInfoWindow(self)
        self.log("✓ 안내 창 열림 (항상 위에 표시)")

        # 백그라운드 스레드에서 실행
        thread = threading.Thread(
            target=self.run_automation,
            args=(csv_file, count, delay, self.dry_run.get()),
            daemon=True
        )
        thread.start()

    def run_automation(self, csv_file, count, delay, dry_run):
        """백그라운드에서 자동화 실행"""
        try:
            # stdout 리디렉션
            original_stdout = sys.stdout
            sys.stdout = LogRedirector(self.log_text, self.log_queue)

            # BulkDependentInput 실행 (verbose=False로 DEBUG 로그 끄기)
            self.bulk_automation = BulkDependentInput(csv_file, verbose=False, global_delay=delay)
            result = self.bulk_automation.run(count=count, dry_run=dry_run)

            # 리소스 정리 (keyboard 후크 해제는 run()에서 이미 처리됨)

            # 결과 표시
            if not dry_run and result:
                self.log_queue.put("\n" + "=" * 50)
                self.log_queue.put("✅ 완료!")
                self.log_queue.put(f"처리: {result['processed']}명")
                self.log_queue.put(f"성공: {result['success']}명")
                self.log_queue.put(f"건너뜀: {result['skipped']}명")
                self.log_queue.put(f"실패: {result['failed']}명")
                self.log_queue.put(f"입력 부양가족: {result['total_dependents']}명")
                self.log_queue.put("=" * 50)

            # stdout 복원
            sys.stdout = original_stdout

            # 성공 완료
            self.after(0, lambda: self.on_automation_complete(True))

        except Exception as e:
            # stdout 복원
            sys.stdout = original_stdout

            # 에러 메시지 추출
            error_message = str(e)

            self.log_queue.put(f"\n❌ 오류 발생: {error_message}")
            import traceback
            self.log_queue.put(traceback.format_exc())

            # 리소스 정리
            if self.bulk_automation:
                try:
                    self.bulk_automation.cleanup()
                except:
                    pass

            # 실패 완료 + 에러 메시지 전달
            self.after(0, lambda: self.on_automation_complete(False, error_message))

    def on_automation_complete(self, success, error_message=None):
        """자동화 완료 후 처리"""
        # 안내 창 닫기
        if self.stop_window is not None:
            try:
                self.stop_window.destroy()
                self.stop_window = None
                self.log("✓ 안내 창 닫힘")
            except Exception as e:
                self.log(f"⚠️ 안내 창 닫기 실패: {e}")

        self.is_running = False
        self.bulk_automation = None  # 인스턴스 리셋
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.browse_btn.configure(state="normal")
        self.count_entry.configure(state="normal")
        self.delay_entry.configure(state="normal")
        self.dry_run_check.configure(state="normal")

        if success:
            self.progress_label.configure(text="✅ 완료!")
            self.progress_bar.set(1.0)
        else:
            self.progress_label.configure(text="❌ 오류 발생")
            self.progress_bar.set(0)

            # 에러 메시지가 있으면 별도 알림창으로 표시
            if error_message:
                messagebox.showerror("오류", error_message)

    def _do_stop(self):
        """실제 중지 처리 (GUI 스레드에서 실행)"""
        if self.bulk_automation and self.is_running:
            self.log("🛑 중지 요청됨 (Pause 키 3번)")
            self.bulk_automation.stop()
            self.stop_btn.configure(state="disabled")  # 중복 클릭 방지

            # 리소스 정리
            try:
                self.bulk_automation.cleanup()
            except:
                pass

            # 안내 창도 닫기
            if self.stop_window is not None:
                try:
                    self.stop_window.destroy()
                    self.stop_window = None
                except:
                    pass

    def stop_automation(self):
        """자동화 중지 (버튼 클릭)"""
        self._do_stop()

    def on_closing(self):
        """윈도우 종료 시 정리"""
        # 실행 중이면 경고
        if self.is_running:
            if messagebox.askokcancel("종료", "자동화가 실행 중입니다. 종료하시겠습니까?"):
                if self.bulk_automation:
                    self.bulk_automation.stop()
                    # 리소스 정리
                    try:
                        self.bulk_automation.cleanup()
                    except:
                        pass
                # 안내 창 닫기
                if self.stop_window is not None:
                    try:
                        self.stop_window.destroy()
                    except:
                        pass
                self.destroy()
        else:
            self.destroy()


def main():
    """메인 함수"""
    app = BulkInputGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
