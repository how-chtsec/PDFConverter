import os
import tkinter as tk
import configparser
import pythoncom
import win32com.client
import threading
import time
import queue
from tkinter import filedialog
from time import strftime
import datetime

class App(tk.Frame):
    global is_quit
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        
        self.pack()

        # 設定檔案名稱
        self.config_file = 'config.ini'

        # 讀取上次設定的路徑
        self.config = configparser.ConfigParser()
        if os.path.exists(self.config_file):
            self.config.read(self.config_file)
            self.input_path = self.config['PATHS']['input']
            self.output_path = self.config['PATHS']['output']
        else:
            self.input_path = ''
            self.output_path = ''
        
        # 第一次使用
        if 'PATHS' not in self.config:
            self.config['PATHS'] = {}
            self.config['PATHS']['output'] = ''

        # 建立 GUI
        self.create_widgets()

        # 關閉視窗事件
        self.master.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # 轉換器worker
        self.converters = []
        
        # 變數
        self.total_jobs = 0
        self.file_queue = queue.Queue()
        self.threads = 1 # 好像沒比較快

    def create_widgets(self):
        self.master.geometry("700x500")
        
        font_style12 = ("Arial", 12)

        # 輸入路徑的標籤、輸入框、按鈕
        self.input_label = tk.Label(self.master, text='輸入路徑', font=font_style12)
        self.input_label.place(x=48, y=50)
        self.input_path_var = tk.StringVar(value=self.input_path)
        self.input_entry = tk.Entry(self.master, textvariable=self.input_path_var, width=43, font=font_style12)
        self.input_entry.place(x=140, y=50)
        self.input_button = tk.Button(self.master, text='選擇輸入資料夾', command=self.choose_input_path, font=font_style12)
        self.input_button.place(x=550, y=45)

        # 輸出路徑的標籤、輸入框、按鈕
        self.output_label = tk.Label(self.master, text='輸出路徑：', font=font_style12)
        self.output_label.place(x=48, y=100)
        self.output_path_var = tk.StringVar(value=self.output_path)
        self.output_entry = tk.Entry(self.master, textvariable=self.output_path_var, width=43, font=font_style12)
        self.output_entry.place(x=140, y=100)
        self.output_button = tk.Button(self.master, text='選擇輸出資料夾', command=self.choose_output_path, font=font_style12)
        self.output_button.place(x=550, y=95)

        # 執行按鈕
        self.execute_button = tk.Button(self.master, text='執行',command=self.execute, font=("Arial", 20))
        self.execute_button.place(relx=0.5, y=180, anchor='center')

        # 進度標籤
        self.msgString = tk.StringVar()
        self.msgString.set('進度')
        self.msgLabel = tk.Label(self.master, textvariable=self.msgString, height=3, fg="black", font=font_style12)
        self.msgLabel.place(x=50, y=200)

        # 錯誤標籤
        self.errorString = tk.StringVar()
        self.errorString.set('錯誤次數: 0')
        self.msgLabel = tk.Label(self.master, textvariable=self.errorString, height=3, fg="black", font=font_style12)
        self.msgLabel.place(x=550, y=200)

        # 訊息框
        self.progress_text_box = tk.Text(self.master, height=15, width=85)
        self.progress_text_box.place(x=50, y=250)

        # scrollbar
        scrollbar = tk.Scrollbar(self.master, orient=tk.VERTICAL,command=self.progress_text_box.yview_scroll)
        scrollbar.place(x=650, y=250, height=200)
        self.progress_text_box.config(yscrollcommand=scrollbar.set)

        # 執行時間標籤
        self.timeString = tk.StringVar()
        self.timeString.set('時間')
        self.msgLabel = tk.Label(self.master, textvariable=self.timeString, height=1, fg="black", font=font_style12)
        self.msgLabel.place(x=50, y=455)

    def choose_input_path(self):
        path = filedialog.askdirectory(initialdir=self.input_entry.get())
        if path:
            self.input_path_var.set(path)
            self.config['PATHS']['input'] = path
            with open(self.config_file, 'w') as configfile:
                self.config.write(configfile)

    def choose_output_path(self):
        path = filedialog.askdirectory(initialdir=self.output_entry.get())
        if path:
            self.output_path_var.set(path)
            self.config['PATHS']['output'] = path
            with open(self.config_file, 'w') as configfile:
                self.config.write(configfile)

    def list_files(self, root_dir):
        # 回傳相對路徑
        file_set = []
        for dir_, _, files in os.walk(root_dir):
            for file_name in files:
                rel_dir = os.path.relpath(dir_, root_dir)
                rel_file = os.path.join(rel_dir, file_name)

                # 將相對路徑 '.\' 移除
                if rel_file.startswith('.\\'):
                    rel_file = rel_file[2:]
                
                # 略過 . or ~ 開頭的隱藏檔
                if os.path.basename(rel_file).startswith('.') or os.path.basename(rel_file).startswith('~'):
                    continue

                print('rel_file', rel_file)

                file_set.append(rel_file)

        return file_set

    def execute(self):
        global is_quit
        global finished_tasks
        global error_count

        if self.execute_button.cget('text') == '執行':
            self.start_time = datetime.datetime.now()
            is_quit = False
            finished_tasks = 0
            error_count = 0

            self.execute_button.configure(text='停止', fg='red')
            self.input_button.config(state="disable") 
            self.output_button.config(state="disable") 
            
            self.progress_text_box.delete('1.0', tk.END)
            input_folder = os.path.normpath(self.input_entry.get())
            output_folder = os.path.normpath(self.output_entry.get())
            
            for file in self.list_files(input_folder):
                if file.endswith('.docx') or file.endswith('.doc') or file.endswith('.xls'):
                    output_file = os.path.join(output_folder, '.'.join(file.split('.')[:-1]) + '.pdf')
                    input_file = os.path.join(input_folder, file)
                    print('%s > %s' % (input_file, output_file))
                    
                    if not os.path.isdir(os.path.dirname(output_file)):
                        os.makedirs(os.path.dirname(output_file))

                    self.file_queue.put((input_file, output_file))

            # 檔案(任務)總數
            self.total_jobs = self.file_queue.qsize()

            # 開始轉檔
            self.converters = []
            for i in range(self.threads):
                converter = Converter(self.file_queue, progress_text_box=self.progress_text_box)
                converter.start_conversion()
                self.converters.append(converter)

            # 監控轉檔進度
            thread = threading.Thread(target=self.refresh_progress)
            thread.daemon = True
            thread.start()
        elif self.execute_button.cget('text') == '停止':
            is_quit = True
            self.progress_text_box.insert('end', '停止中...等待剩餘任務完成...\n')
            self.progress_text_box.insert('end', '執行緒' + str([_.ident for _ in self.converters]) + '停止中...\n')
            self.progress_text_box.see('end')
            thread = threading.Thread(target=self.wait_all_tasks_quit)
            thread.daemon = True
            thread.start()
 
    def refresh_progress(self):
        global lock
        global error_count
        # queue 還有東西
        while not self.file_queue.empty() or not self.all_tasks_finished():
            # 每隔1秒檢查一次子執行緒是否已經完成
            self.msgString.set('進度: 執行中 (%s)' % self.get_progress())
            self.errorString.set('錯誤次數: % d' % error_count)
            time.sleep(1)

        self.msgString.set('進度: 完成 (%s)' % self.get_progress())
        self.errorString.set('錯誤次數: % d' % error_count)
        self.progress_text_box.insert('end', '完成')
        self.progress_text_box.see('end')
        self.timeString.set('時間: %s' % str(datetime.datetime.now() - self.start_time))
        self.execute_button.configure(text='執行', fg='black')
        self.input_button.config(state="active") 
        self.output_button.config(state="active")
    
    def get_progress(self):
        global finished_tasks
        return '%s/%s' % (finished_tasks, self.total_jobs)

    def wait_all_tasks_quit(self):
        while True:
            # queue is empty and all threads are stop
            if self.file_queue.empty() and self.all_tasks_finished():
                break
            time.sleep(1)
        self.progress_text_box.insert('end', '全部任務已停止\n')
        self.progress_text_box.see('end')

    def wait_to_close_windows(self):
        while True:
            # queue is empty and all threads are stop
            if self.file_queue.empty() and self.all_tasks_finished():
                self.master.destroy()
                break
            time.sleep(1)

    def all_tasks_finished(self):
        # return True if all tasks are stop
        # print('all_tasks_finished', any([_.is_alive() for _ in self.converters]), [_.is_alive() for _ in self.converters])
        return not any([_.is_alive() for _ in self.converters])

    def on_close(self):
        global is_quit

        # 還有任務未完成時
        if not self.file_queue.empty() or not self.all_tasks_finished():
            result = tk.messagebox.askyesno("警告", "還有任務在執行中，確定要關閉嗎？")
            if result:
                is_quit = True
                self.progress_text_box.insert('end', '關閉中...等待剩餘任務完成...\n')
                self.progress_text_box.insert('end', '執行緒' + str([_.ident for _ in self.converters]) + '停止中...\n')
                self.progress_text_box.see('end')

                # 等待當前任務完成後關閉
                thread = threading.Thread(target=self.wait_to_close_windows)
                thread.daemon = True
                thread.start()
        else:
            self.master.destroy()

class Converter(threading.Thread):
    def __init__(self, file_queue, progress_text_box=None, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.file_queue = file_queue
        self.progress_text_box = progress_text_box

    def start_conversion(self):
        self.daemon = True
        self.start()

    def run(self):
        self.convert_to_pdf()

    def convert_to_pdf(self):
        global is_quit
        global lock
        global error_count
        global finished_tasks

        pythoncom.CoInitialize()

        while not self.file_queue.empty() and not is_quit:
            input_file, output_file = self.file_queue.get()
            try:
                print(input_file, output_file)
                if input_file.split('.'[-1]) == 'xls':
                    word = win32com.client.DispatchEx("Excel.Application")
                else:
                    word = win32com.client.DispatchEx("Word.Application")

                msg = "--- %s -- docx  -> pdf %s" % (strftime("%H:%M:%S"), os.path.relpath(output_file))
                if self.progress_text_box is not None:
                    self.progress_text_box.insert('end', msg + '\n')
                    self.progress_text_box.see('end')
                
                print(msg) 
                doc = word.Documents.Open(input_file)
                    
                doc.SaveAs(output_file, FileFormat = 17)
                doc.Close()
            except Exception as e:
                
                print(e)
                error_msg = 'Error: %s, Path: %s\n' % (str(e), input_file)
                if self.progress_text_box is not None:
                    self.progress_text_box.insert('end', error_msg)
                    self.progress_text_box.see('end')
                error_count += 1
                lock.acquire()
                with open('error.log', 'a') as f:
                    f.write(error_msg + '\n')
                lock.release()
            finally:
                word.Quit()
                finished_tasks += 1
        
        if is_quit:
            msg = 'Thread %s quit' % threading.get_ident()
            self.file_queue.queue.clear()
            print(msg)
            self.progress_text_box.insert('end', msg + '\n')
            self.progress_text_box.see('end')

# Global variable
is_quit = False
is_pause = False
lock = threading.Lock()
error_count = 0
finished_tasks = 0

# Start GUI
root = tk.Tk()
root.title("PDF Converter (doc, docx, xls)")
app = App(master=root)
app.mainloop()