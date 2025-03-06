import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
import pandas as pd
import requests
import openai
from openai import OpenAI
#import json
import os
#import openpyxl
import threading


client = OpenAI(
    api_key="YOUR_KEY", # 在这里将 MOONSHOT_API_KEY 替换为你从 Kimi 开放平台申请的 API Key
    base_url="YOUR_URL",
)

def read_excel(file_path):
    return pd.read_excel(file_path)


def get_model_response(prompt):
    #openai.default_headers = {"x-foo": "true"}

    completion = client.chat.completions.create(
        model="YOUR_LLM",
        messages=[
            {
                "role": "user",
                "content": prompt,
            },
        ],
    )

    try:
        answer = completion.choices[0].message.content
       # response = requests.request("POST", url, headers=headers, data=payload, timeout=30)  # 增加超时时间为30秒
        return answer
    except requests.exceptions.Timeout:
        print("请求超时，正在尝试重新连接...")
        # 这里可以添加重试逻辑或其他处理
    except requests.exceptions.ConnectionError:
        print("网络连接错误")

def analyze_documents(file_path, columns_to_use, requirement, save_path, update_index_callback):
    df = read_excel(file_path)
    # 使用“模型回答”作为存放模型回复的列（避免与“备注”混淆）
    df['模型回答'] = ''

    try:
        # 遍历每一行数据
        for index, row in df.iterrows():
            try:
                # 拼接用户选定列的数据和需求构造 prompt
                prompt = ' '.join([f"{col}: {row[col]}" for col in columns_to_use if pd.notna(row[col])])
                response = get_model_response(prompt + requirement)
                df.at[index, '模型回答'] = response  # 保存模型回答
                update_index_callback(os.path.basename(file_path), index)
            except Exception as e:
                # 当处理单行数据时发生错误时，打印错误并中断当前文件的后续处理
                print(f"处理文件 {file_path} 第 {index} 行时出错: {e}")
                break  # 出错后退出循环，以便保存已处理的数据
    except Exception as e:
        # 捕获文件级别的意外错误
        print(f"处理文件 {file_path} 时出现意外错误: {e}")
    finally:
        # 无论是否出错，都保存当前处理的数据
        save_file_path = os.path.join(save_path, 'processed_' + os.path.basename(file_path))
        df.to_excel(save_file_path, index=False)
        print(f"已将处理后的数据保存至 {save_file_path}")


def process_all_excels(directory, columns_to_use, requirement, save_path, update_index_callback):
    for filename in os.listdir(directory):
        # 检查文件是不是Excel文件且不是示例文件
        if (filename.endswith('.xlsx') or filename.endswith('.xls')) and filename != "示例文件.xlsx":
            file_path = os.path.join(directory, filename)
            analyze_documents(file_path, columns_to_use, requirement, save_path, update_index_callback)
            print(f"Processed {filename}")

def create_menu(master):
    menubar = tk.Menu(master)
    master.config(menu=menubar)

    # 帮助菜单
    help_menu = tk.Menu(menubar, tearoff=0)
    help_menu.add_command(label="注意事项", command=open_help)
    help_menu.add_command(label="关于", command=open_about)
    menubar.add_cascade(label="About", menu=help_menu)

def open_help():
    # 显示帮助文档
    messagebox.showinfo("注意事项", "1.目前调用的是百度云端的文心一言2.0大模型，百度云对算力收费，0.008元/每千token（大约750词)\n"
                                    "2.确保输入文件夹中的Excel文件格式正确且无损坏\n"
                                    "3.请确保网络连接正常，以便与百度AI接口通信")

def open_about():
    # 显示关于对话框
    messagebox.showinfo("关于", "AlexAna：智能文档分析软件\ndeveloped by CAI HONG\n版本 1.0")

class LargeInputDialog(tk.Toplevel):
    def __init__(self, parent, title="输入要求", prompt="请输入您的要求:"):
        super().__init__(parent)
        self.title(title)
        self.result = None

        tk.Label(self, text=prompt).pack(padx=10, pady=10)

        self.text_input = tk.Text(self, width=60, height=10)
        self.text_input.pack(padx=10, pady=10)

        button_frame = tk.Frame(self)
        button_frame.pack(padx=10, pady=10)

        tk.Button(button_frame, text="确定", command=self.on_ok).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="取消", command=self.on_cancel).pack(side=tk.RIGHT, padx=5)

        self.transient(parent)  # 设置为临时窗口
        self.wait_visibility()  # 等待窗口出现
        self.grab_set()  # 模态对话框
        self.wait_window()  # 等待对话框关闭

    def on_ok(self):
        self.result = self.text_input.get("1.0", tk.END).strip()
        self.destroy()

    def on_cancel(self):
        self.destroy()

class DocumentAnalyzerUI:
    def __init__(self, master):
        self.master = master  # Add this line to set self.master
        master.title("LLM智能文档分析软件demo")
        master.geometry("1000x600")  # 设置窗口大小

        # 使用 PanedWindow 分割窗口为主功能区和调试区
        paned_window = tk.PanedWindow(master, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True)


        main_frame = tk.Frame(paned_window)
        paned_window.add(main_frame, width=600)  # 设置主功能区宽度

        # 调试区框架
        debug_frame = tk.Frame(paned_window)
        paned_window.add(debug_frame, width=400)  # 设置调试区宽度

        # 在main_frame中添加原有的UI组件
        tk.Label(main_frame, text="选择输入文件夹:请确保文件夹中的excel文件以xlsx结尾").grid(row=0, column=0)
        self.input_dir_button = tk.Button(main_frame, text="选择文件夹", command=self.select_input_directory)
        self.input_dir_button.grid(row=1, column=0)

        tk.Label(main_frame, text="选择输出文件夹:存放处理后文件").grid(row=2, column=0)
        self.output_dir_button = tk.Button(main_frame, text="选择文件夹", command=self.select_output_directory)
        self.output_dir_button.grid(row=3, column=0)

        self.columns_label = tk.Label(main_frame, text="从列表中选择一个或多个列，这些列的数据将用于生成AI模型的请求")
        self.columns_label.grid(row=4, column=0)

        self.columns_listbox = tk.Listbox(main_frame, selectmode=tk.MULTIPLE)
        self.columns_listbox.grid(row=5, column=0)

        tk.Label(main_frame, text="请尽可能精准地描述您的需求").grid(row=6, column=0)
        self.requirement_button = tk.Button(main_frame, text="输入问题", command=self.enter_requirement)
        self.requirement_button.grid(row=7, column=0)

        tk.Label(main_frame, text="准备就绪后，点击“开始处理”运行程序").grid(row=8, column=0)
        self.process_button = tk.Button(main_frame, text="开始处理", command=self.process_files)
        self.process_button.grid(row=9, column=0)

        self.status_label = tk.Label(main_frame, text="未开始处理")
        self.status_label.grid(row=10, column=0)

        # 在debug_frame中添加调试区的UI组件
        tk.Label(debug_frame, text="提问:").grid(row=0, column=0)
        self.question_text = tk.Text(debug_frame, height=10, width=40)
        self.question_text.grid(row=1, column=0)

        tk.Label(debug_frame, text="回答:").grid(row=4, column=0)
        self.answer_text = tk.Text(debug_frame, height=10, width=40, state=tk.DISABLED)
        self.answer_text.grid(row=5, column=0)

        self.ask_button = tk.Button(debug_frame, text="获取模型回答", command=self.get_model_answer)
        self.ask_button.grid(row=2, column=0)

        # 初始化其他成员变量
        self.input_directory = ""
        self.output_directory = ""
        self.columns = []
        self.requirement = ""


    def select_input_directory(self):
        self.input_directory = filedialog.askdirectory()
        if self.input_directory:  # 如果选择了文件夹
            self.update_columns_listbox()  # 更新列名列表框
        self.update_status()

    def select_output_directory(self):
        self.output_directory = filedialog.askdirectory()
        self.update_status()

    def update_index(self, filename, index):
        if self.master.winfo_exists():  # 检查窗口是否存在
            self.status_label.config(text=f"正在处理文件《{filename}》的第{index}行")
            self.master.update_idletasks()

    def enter_requirement(self):
        dialog = LargeInputDialog(self.master)
        if dialog.result:
            self.requirement = dialog.result
            self.update_status()

    def update_columns_listbox(self):
        # 清空现有的列名列表
        self.columns_listbox.delete(0, tk.END)

        # 检查文件夹是否存在并获取第一个Excel文件
        if self.input_directory and os.path.exists(self.input_directory):
            for filename in os.listdir(self.input_directory):
                if filename.endswith('.xlsx') or filename.endswith('.xls'):
                    file_path = os.path.join(self.input_directory, filename)
                    try:
                        df = pd.read_excel(file_path)
                        self.columns = df.columns.tolist()
                        for col in self.columns:
                            self.columns_listbox.insert(tk.END, col)
                        break  # 找到第一个文件后就退出循环
                    except Exception as e:
                        messagebox.showwarning("警告", f"无法读取文件 '{filename}': {e}")
                        break
        else:
            messagebox.showwarning("警告", "输入文件夹未选择或不存在")

    def process_files(self):
        if not self.input_directory or not self.output_directory or not self.requirement:
            messagebox.showwarning("警告", "请确保所有选项都已设置")
            return

        # 获取用户选择的列名
        selected_columns = [self.columns_listbox.get(idx) for idx in self.columns_listbox.curselection()]

        if not selected_columns:
            messagebox.showwarning("警告", "请至少选择一个列名")
            return
        # 使用线程来处理文件
        processing_thread = threading.Thread(
            target=self.run_processing,
            args=(self.input_directory, selected_columns, self.requirement, self.output_directory)
        )
        processing_thread.start()

    def run_processing(self, directory, columns_to_use, requirement, save_path):
        process_all_excels(directory, columns_to_use, requirement, save_path, self.update_index)
        self.master.after(0, self.processing_complete)  # 在主线程中调用处理完成的方法
    def processing_complete(self):
        messagebox.showinfo("完成", "文件处理完成")
    def update_status(self):
        status = f"输入路径: {self.input_directory}\n输出路径: {self.output_directory}\n要求: {self.requirement}"
        self.columns_label.config(text="从列表中选择一个或多个列，这些列的数据将用于生成AI模型的请求")
        # 获取用户选择的列名
        selected_indices = self.columns_listbox.curselection()
        selected_columns = [self.columns_listbox.get(i) for i in selected_indices]

        if selected_columns:
            status += f"\n已选择的列名: {', '.join(selected_columns)}"
        else:
            status += "\n未选择列名"

        messagebox.showinfo("当前状态", status)
    def get_model_answer(self):
        question = self.question_text.get("1.0", tk.END).strip()
        if question:
            answer = get_model_response(question)
            print(answer)  # 打印 answer 变量的值
            self.answer_text.config(state=tk.NORMAL)
            self.answer_text.delete("1.0", tk.END)
            self.answer_text.insert(tk.END, answer)
            self.answer_text.config(state=tk.DISABLED)

root = tk.Tk()
app = DocumentAnalyzerUI(root)
root.mainloop()
