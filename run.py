import ctypes

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # Windows 8.1+
except AttributeError:
    try:
        ctypes.windll.user32.SetProcessDPIAware(True)  # Windows Vista+
    except AttributeError:
        pass

import tkinter as tk
from tkinter import ttk, messagebox
from PIL import ImageGrab, Image
import pytesseract
import requests
import threading
import subprocess
import sys
import os

# ==========================================
# ===== 总配置 (Remote / Ollama 切换) =====
# ==========================================

# ===== 1. 总开关 =====
# 在这里选择要使用的 AI 后端:
# - "remote" : 使用远程 (OpenAI 兼容) API
# - "ollama" : 使用本地 Ollama 服务
AI_BACKEND = "ollama"  # <-- ！！！您只需要修改这一行！！！

# ===== 2. Remote (远程) 配置 =====
REMOTE_API_KEY = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"  # 在这里填写您的 Remote API Key
REMOTE_API_URL = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"  # OpenAI 兼容 URL
REMOTE_MODEL = "xxxxxxxxxx"                           # 模型名称

# ===== 3. Ollama (本地) 配置 =====
OLLAMA_API_URL = "http://127.0.0.1:11434/api/chat"
OLLAMA_MODEL = "xxxxxxxxxx"                           # 模型名称

# ===== 4. 通用配置 =====
OCR_LANG = 'chi_sim+eng'
TIME_LIMIT_SECONDS = 3
REQUEST_TIMEOUT_SECONDS = 60


# ==========================================
# ===== 截图遮罩窗口  =====
# ==========================================
class CaptureOverlay:
    """
    创建一个全屏半透明窗口，用于选择截图区域。
    [必须在主线程中创建和运行]
    """

    def __init__(self, root):
        self.root = root
        self.overlay = tk.Toplevel(root)
        self.overlay.attributes('-fullscreen', True)
        self.overlay.attributes('-alpha', 0.2)
        self.overlay.attributes('-topmost', True)
        self.canvas = tk.Canvas(self.overlay, cursor="crosshair", bg="gray")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.start_x = None
        self.start_y = None
        self.rect = None
        self.selection_box = None
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
        self.root.wait_window(self.overlay)
        return self.selection_box


# ==========================================
# ===== 主应用窗口 (重构 AI 调用逻辑) =====
# ==========================================

class OcrHelperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("识图解题助手")
        self.root.geometry("340x400+100+100")
        self.root.attributes('-topmost', True)
        # --- 顶部控制栏 ---
        control_frame = ttk.Frame(root, padding=10)
        control_frame.pack(fill='x')

        # 标题现在会显示当前用的是 Remote 还是 Ollama
        backend_name = "Remote (远程)" if AI_BACKEND == "remote" else f"Ollama ({OLLAMA_MODEL.split(':')[0]}) (本地)"

        self.start_button = ttk.Button(
            control_frame,
            text="开始截图",
            command=self.start_capture
        )
        self.start_button.pack(side='right', padx=(0, 2))
        ttk.Label(control_frame, text=f"识题助手 ({backend_name})").pack(side='left', fill='x', expand=True,padx=(2, 0))
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
        self.root.after(0, lambda: self.status_var.set(text))

    def show_answer(self, ocr_text, ai_answer):
        def update():
            self.answer_text.delete(1.0, tk.END)
            self.answer_text.insert(tk.END, "--- OCR 结果 ---\n", "ocr_header")
            self.answer_text.insert(tk.END, ocr_text + "\n\n", "content")
            self.answer_text.insert(tk.END, "--- AI 回答 ---\n", "ai_header")
            self.answer_text.insert(tk.END, ai_answer, "content")

        self.root.after(0, update)

    def start_capture(self):
        try:
            self.root.withdraw()
            self.start_button.config(state='disabled')
            self.set_status("状态：请选择截图区域 (Esc取消)")
            overlay = CaptureOverlay(self.root)
            bbox = overlay.get_selection()
            self.root.deiconify()
            if bbox is None:
                self.set_status("状态：已取消")
                self.start_button.config(state='normal')
                return
            self.set_status("状态：正在截图...")
            self.root.after(50, self.start_ocr_workflow, bbox)
        except Exception as e:
            self._handle_error(e)

    def start_ocr_workflow(self, bbox):
        try:
            image = ImageGrab.grab(bbox=bbox, all_screens=True)
            self.set_status(f"状态：OCR识别中 (Lang: {OCR_LANG})...")
            thread = threading.Thread(
                target=self.run_ocr_and_ai,
                args=(image,)
            )
            thread.start()
        except Exception as e:
            self._handle_error(e)

    def run_ocr_and_ai(self, image):
        ocr_text = "(OCR失败)"
        try:
            ocr_text = pytesseract.image_to_string(image, lang=OCR_LANG)
            if not ocr_text or ocr_text.strip() == "":
                self.set_status("状态：OCR识别失败或无文字")
                self.show_answer("（未识别到文字）", "")
                return
            self.set_status("状态：OCR完成，正在请求AI...")
            self.show_answer(ocr_text, "（AI正在思考...）")

            # --- AI 调用路由 ---
            ai_answer = self.ask_ai(ocr_text)
            # ------------------------------------

            self.set_status("状态：完成")
            self.show_answer(ocr_text, ai_answer)
        except Exception as e:
            self._handle_error_threaded(e, ocr_text)
        finally:
            self.root.after(0, lambda: self.start_button.config(state='normal'))

    def _handle_error(self, e):
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
        error_msg = f"发生错误: {str(e)}"
        print(f"[Error] {error_msg}")
        import traceback
        traceback.print_exc()
        self.root.after(0, self.set_status, error_msg)
        self.root.after(0, self.show_answer, ocr_text, f"--- 错误 ---\n{str(e)}")
        self.root.after(0, lambda: {
            self.root.deiconify() if self.root.state() == 'withdrawn' else None
        })

    # ==========================================
    # =====  AI 路由及函数 =====
    # ==========================================

    def ask_ai(self, ocr_text):
        """
        [AI 路由函数]
        根据顶部的 AI_BACKEND 配置, 调用 Remote 或 Ollama。
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
        try:
            # --- 路由开始 ---
            if AI_BACKEND == "remote":
                return self._ask_remote(prompt) 
            elif AI_BACKEND == "ollama":
                return self._ask_ollama(prompt)
            # --- 路由结束 ---
            else:
                return f"配置错误: 未知的 AI_BACKEND: '{AI_BACKEND}'"

        except requests.exceptions.ConnectionError as e:
            # 捕获 Ollama 连接失败
            if AI_BACKEND == "ollama":
                return "Ollama连接失败：请确认 Ollama 服务正在运行"
            return f"AI连接失败：{str(e)}"
        except requests.exceptions.Timeout:
            return f"AI请求超时 (超过 {REQUEST_TIMEOUT_SECONDS} 秒)"
        except requests.exceptions.RequestException as e:
            # 捕获 Remote 401 或 Ollama 404
            if AI_BACKEND == "remote" and e.response and e.response.status_code == 401:
                return "Remote 错误：API Key无效或过期，请检查配置。" 
            if AI_BACKEND == "ollama" and e.response and e.response.status_code == 404:
                return f"Ollama模型未找到: 请确认模型 '{OLLAMA_MODEL}' 已下载"
            return f"AI请求失败：{str(e)}"
        except Exception as e:
            return f"AI解析失败：{str(e)}"

    def _ask_remote(self, prompt_text):
        """ [AI 助手] 调用 Remote (OpenAI 兼容) API """ 
        headers = {
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {REMOTE_API_KEY}' 
        }
        body = {
            "model": REMOTE_MODEL, 
            "messages": [
                {"role": "system", "content": "你是识图答题助手，只返回【答案】行。"},
                {"role": "user", "content": prompt_text}
            ],
            "max_tokens": 1000,
            "temperature": 0.0
        }
        resp = requests.post(
            REMOTE_API_URL,  
            headers=headers,
            json=body,
            timeout=REQUEST_TIMEOUT_SECONDS
        )
        resp.raise_for_status()  # 抛出 4xx/5xx 错误, 由 ask_ai 捕获
        j = resp.json()
        content = j.get("choices", [{}])[0].get("message", {}).get("content", "")
        return content.strip() or "（Remote未返回有效答案）"  

    def _ask_ollama(self, prompt_text):
        """ [AI 助手] 调用本地 Ollama API """
        headers = {
            'Content-Type': 'application/json',
        }
        body = {
            "model": OLLAMA_MODEL,
            "messages": [
                {"role": "system", "content": "你是识图答题助手，只返回【答案】行。"},
                {"role": "user", "content": prompt_text}
            ],
            "options": {
                "temperature": 0.0
            },
            "stream": False
        }
        resp = requests.post(
            OLLAMA_API_URL,
            headers=headers,
            json=body,
            timeout=REQUEST_TIMEOUT_SECONDS
        )
        resp.raise_for_status()  # 抛出 4xx/5xx 错误, 由 ask_ai 捕获
        j = resp.json()
        content = j.get("message", {}).get("content", "")
        return content.strip() or "（Ollama未返回有效答案）"


# ==========================================
# ===== 启动检查 =====
# ==========================================

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


# ===== 主程序入口  =====
if __name__ == "__main__":

    root = tk.Tk()
    app = OcrHelperApp(root)
    root.mainloop()