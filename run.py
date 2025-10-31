import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # Windows 8.1+
except AttributeError:
    try:
        ctypes.windll.user32.SetProcessDPIAware(True)  # Windows Vista+
    except AttributeError:
        pass
# ==========================================

import tkinter as tk
from tkinter import ttk, messagebox
from PIL import ImageGrab, Image
import pytesseract
import requests
import threading
import subprocess
import sys
import os

# ===== 配置 (请修改为您自己的) =====
AI_API_KEY = "sk-4ldUznwMK9DXqSXtnBHIc2zOuKPh66ue5C4LINF5naQWcpTI"
AI_API_URL = "https://api.moonshot.cn/v1/chat/completions"
AI_MODEL = "kimi-latest"
OCR_LANG = 'chi_sim+eng'
TIME_LIMIT_SECONDS = 3


# ===== 截图遮罩窗口 (无需修改) =====
class CaptureOverlay:
    """
    创建一个全屏半透明窗口，用于选择截图区域。
    [必须在主线程中创建和运行]
    """

    def __init__(self, root):
        self.root = root
        self.overlay = tk.Toplevel(root)
        self.overlay.attributes('-fullscreen', True)
        self.overlay.attributes('-alpha', 0.2)  # 半透明
        self.overlay.attributes('-topmost', True)  # 始终置顶

        self.canvas = tk.Canvas(self.overlay, cursor="crosshair", bg="gray")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.start_x = None
        self.start_y = None
        self.rect = None
        self.selection_box = None

        # 绑定事件
        self.canvas.bind("<ButtonPress-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_move)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.overlay.bind("<Escape>", self.cancel)

    def on_mouse_down(self, event):
        self.start_x = event.x
        self.start_y = event.y
        if self.rect:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y,
            outline="blue", width=2, dash=(4, 2)
        )

    def on_mouse_move(self, event):
        cur_x, cur_y = event.x, event.y
        self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)

    def on_mouse_up(self, event):
        end_x, end_y = event.x, event.y
        x1 = min(self.start_x, end_x)
        y1 = min(self.start_y, end_y)
        x2 = max(self.start_x, end_x)
        y2 = max(self.start_y, end_y)
        screen_x1 = self.overlay.winfo_rootx() + x1
        screen_y1 = self.overlay.winfo_rooty() + y1
        screen_x2 = self.overlay.winfo_rootx() + x2
        screen_y2 = self.overlay.winfo_rooty() + y2

        if (x2 - x1) < 10 or (y2 - y1) < 10:
            self.selection_box = None
        else:
            self.selection_box = (screen_x1, screen_y1, screen_x2, screen_y2)
        self.overlay.destroy()

    def cancel(self, event=None):
        self.selection_box = None
        self.overlay.destroy()

    def get_selection(self):
        """
        [阻塞主线程] 显示窗口并等待用户选择，返回 BBox 或 None。
        """
        self.root.wait_window(self.overlay)
        return self.selection_box


# ===== 主应用窗口  =====

class OcrHelperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("识图解题助手")
        self.root.geometry("340x400+100+100")
        self.root.attributes('-topmost', True)

        # --- 顶部控制栏 ---
        control_frame = ttk.Frame(root, padding=10)
        control_frame.pack(fill='x')
        ttk.Label(control_frame, text="识题助手 (Kimi)").pack(side='left', expand=True)
        self.start_button = ttk.Button(
            control_frame,
            text="开始截图",
            command=self.start_capture
        )
        self.start_button.pack(side='right')

        # --- 状态栏 ---
        self.status_var = tk.StringVar()
        self.status_var.set("状态：等待操作")
        status_label = ttk.Label(
            root, textvariable=self.status_var, padding=(10, 0, 10, 5)
        )
        status_label.pack(fill='x')

        # --- 结果显示区 ---
        text_frame = ttk.Frame(root, padding=(10, 0, 10, 10))
        text_frame.pack(fill='both', expand=True)
        self.answer_text = tk.Text(
            text_frame, wrap='word', height=10, font=("system-ui", 10),
            bg="#f0f0f0", borderwidth=0, padx=8, pady=8
        )
        scrollbar = ttk.Scrollbar(text_frame, command=self.answer_text.yview)
        self.answer_text['yscrollcommand'] = scrollbar.set
        scrollbar.pack(side='right', fill='y')
        self.answer_text.pack(side='left', fill='both', expand=True)
        self.answer_text.tag_config("ocr_header", foreground="#666", font=("system-ui", 10, "bold"))
        self.answer_text.tag_config("ai_header", foreground="#005a9c", font=("system-ui", 10, "bold"))
        self.answer_text.tag_config("content", foreground="#000")

    def set_status(self, text):
        """[安全] 安全地从任何线程更新状态栏"""
        self.root.after(0, lambda: self.status_var.set(text))

    def show_answer(self, ocr_text, ai_answer):
        """[安全] 安全地从任何线程更新结果文本框"""

        def update():
            self.answer_text.delete(1.0, tk.END)
            self.answer_text.insert(tk.END, "--- OCR 结果 ---\n", "ocr_header")
            self.answer_text.insert(tk.END, ocr_text + "\n\n", "content")
            self.answer_text.insert(tk.END, "--- AI 回答 ---\n", "ai_header")
            self.answer_text.insert(tk.END, ai_answer, "content")

        self.root.after(0, update)

    def start_capture(self):
        """
        [主线程] 1. 开始截图流程。
        所有 GUI 操作都在这个函数中完成。
        """
        try:
            # 隐藏主窗口，禁用按钮
            self.root.withdraw()
            self.start_button.config(state='disabled')
            self.set_status("状态：请选择截图区域 (Esc取消)")

            # [主线程 阻塞] 创建并等待截图遮罩
            overlay = CaptureOverlay(self.root)
            bbox = overlay.get_selection()  # 这会阻塞主线程，直到截图完成或取消

            # [主线程] 截图结束，显示主窗口
            self.root.deiconify()

            if bbox is None:
                self.set_status("状态：已取消")
                self.start_button.config(state='normal')
                return

            # [主线程] 截图成功
            self.set_status("状态：正在截图...")

            # 延迟50ms执行，确保遮罩完全消失，GUI刷新
            self.root.after(50, self.start_ocr_workflow, bbox)

        except Exception as e:
            # 捕获截图遮罩期间的意外错误
            self._handle_error(e)

    def start_ocr_workflow(self, bbox):
        """
        [主线程] 2. 截取图像，然后启动工作线程进行OCR和AI。
        """
        try:
            # [主线程] 执行截图 (很快)
            image = ImageGrab.grab(bbox=bbox, all_screens=True)

            self.set_status(f"状态：OCR识别中 (Lang: {OCR_LANG})...")

            # [工作线程] 启动一个新线程来处理【耗时】的OCR和AI
            thread = threading.Thread(
                target=self.run_ocr_and_ai,
                args=(image,)  # 将图像数据传递给新线程
            )
            thread.start()

        except Exception as e:
            # 捕获 ImageGrab 期间的错误
            self._handle_error(e)

    def run_ocr_and_ai(self, image):
        """
        [工作线程] 3. 执行【耗时】的OCR和AI请求。
        !!! 此函数中【严禁】直接操作 GUI (tk) !!!
        """
        ocr_text = "(OCR失败)"  # 预设值以便错误处理
        try:
            # ！！！这是调用 Tesseract 引擎的核心 ！！！
            ocr_text = pytesseract.image_to_string(image, lang=OCR_LANG)

            if not ocr_text or ocr_text.strip() == "":
                self.set_status("状态：OCR识别失败或无文字")
                self.show_answer("（未识别到文字）", "")
                return  # 结束工作线程

            # [安全] 通知主线程更新GUI
            self.set_status("状态：OCR完成，正在请求AI...")
            self.show_answer(ocr_text, "（AI正在思考...）")

            # [工作线程] 4. 请求 AI (耗时)
            ai_answer = self.ask_ai(ocr_text)

            # [安全] 通知主线程更新最终结果
            self.set_status("状态：完成")
            self.show_answer(ocr_text, ai_answer)

        except Exception as e:
            # [安全] 捕获 OCR 或 AI 期间的错误，通知主线程显示
            self._handle_error_threaded(e, ocr_text)

        finally:
            # [安全] 无论成功与否，最后都通知主线程恢复按钮
            self.root.after(0, lambda: self.start_button.config(state='normal'))

    def _handle_error(self, e):
        """ [主线程] 统一的GUI错误处理 """
        error_msg = f"发生错误: {str(e)}"
        print(f"[Error] {error_msg}")
        import traceback
        traceback.print_exc()

        self.set_status(error_msg)
        self.show_answer("(未开始)", f"--- 错误 ---\n{str(e)}")

        if self.root.state() == 'withdrawn':
            self.root.deiconify()
        self.start_button.config(state='normal')

    def _handle_error_threaded(self, e, ocr_text):
        """ [工作线程] 安全的错误处理，用于在主线程更新GUI """
        error_msg = f"发生错误: {str(e)}"
        print(f"[Error] {error_msg}")
        import traceback
        traceback.print_exc()

        # 使用 self.root.after(0, ...) 来确保GUI操作在主线程
        self.root.after(0, self.set_status, error_msg)
        self.root.after(0, self.show_answer, ocr_text, f"--- 错误 ---\n{str(e)}")
        self.root.after(0, lambda: {
            self.root.deiconify() if self.root.state() == 'withdrawn' else None
        })

    def ask_ai(self, ocr_text):
        """
        调用 Kimi AI (无需修改)
        """
        prompt = f"""
你是一名高效选择题助手。
你最多只能思考{TIME_LIMIT_SECONDS}秒，必须在{TIME_LIMIT_SECONDS}秒内给出答案。
---
这是从图片中OCR识别出的文本：
---
{ocr_text}
---
注意：
- 识别结果可能有错字或格式混乱，请尽力理解。
- 选项可能用 A/B/C/D、1/2/3/4、①②③④，或没有编号。
- 如果没有编号，请输出完整选项文本作为答案。
- 输出格式固定为：【答案】xxx
仅输出该行，勿额外解释或多行。
"""
        headers = {
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {AI_API_KEY}'
        }
        body = {
            "model": AI_MODEL,
            "messages": [
                {"role": "system", "content": "你是识图答题助手，只返回【答案】行。"},
                {"role": "user", "content": prompt}
            ],
            "max_tokens": 1000,
            "temperature": 0.0
        }
        try:
            resp = requests.post(AI_API_URL, headers=headers, json=body, timeout=30)
            if resp.status_code == 401:
                return "AI错误：API Key无效或过期，请检查配置。"
            resp.raise_for_status()
            j = resp.json()
            content = j.get("choices", [{}])[0].get("message", {}).get("content", "")
            return content.strip() or "（AI未返回有效答案）"
        except requests.exceptions.RequestException as e:
            return f"AI请求失败：{str(e)}"
        except Exception as e:
            return f"AI解析失败：{str(e)}"


# ===== 启动检查 (无需修改) =====

def check_tesseract_installed():
    """
    检查 Tesseract OCR 引擎是否已安装在系统中。
    """
    try:
        subprocess.run(
            ["tesseract", "--version"],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        )
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        return False


# ===== 主程序入口 (无需修改) =====
if __name__ == "__main__":
    root = tk.Tk()
    app = OcrHelperApp(root)
    root.mainloop()