import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import requests
from PIL import Image, ImageTk
from io import BytesIO
from functools import partial

class ImageApp:
    def __init__(self, root):
        self.root = root
        self.root.title("图片查看器")
        self.root.geometry('1000x1150')

        # 设置主题
        style = ttk.Style()
        style.theme_use('clam')

        # 文本框
        self.label = tk.Label(root, text="选择一个Excel文件加载图片", font=("Helvetica", 16), bg="#f0f0f0", fg="#333333")
        self.label.pack(pady=20)

        # 按钮
        self.load_button = tk.Button(root, text="加载Excel文件", command=self.load_excel, font=("Helvetica", 14), bg="#4CAF50", fg="white", activebackground="#45a049")
        self.load_button.pack(pady=10)

        # 创建 Canvas 和滚动条
        self.canvas = tk.Canvas(root, bg="#ffffff")
        self.v_scrollbar = tk.Scrollbar(root, orient="vertical", command=self.canvas.yview, width=20)
        self.h_scrollbar = tk.Scrollbar(root, orient="horizontal", command=self.canvas.xview, width=20)
        self.scrollable_frame = tk.Frame(self.canvas, bg="#ffffff")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.v_scrollbar.pack(side="right", fill="y")
        self.h_scrollbar.pack(side="bottom", fill="x")

        self.canvas.bind("<MouseWheel>", self._on_mouse_wheel)
        self.scrollable_frame.bind("<Enter>", self._bind_to_mousewheel)
        self.scrollable_frame.bind("<Leave>", self._unbind_to_mousewheel)

        self.root.bind("<Left>", self._on_left_arrow)
        self.root.bind("<Right>", self._on_right_arrow)

        self.current_page = 0
        self.images_per_page = 5

        style.configure("Treeview.Heading", font=("Segoe UI", 14, "bold"), background="#e1e1e1", foreground="#333333")
        style.configure("Treeview", font=("Segoe UI", 12), rowheight=35, background="#ffffff", foreground="#333333")

        self.center_window()

    def _on_mouse_wheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        if self.canvas.yview()[1] == 1.0:
            self.current_page += 1
            self.show_page()

    def _bind_to_mousewheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mouse_wheel)

    def _unbind_to_mousewheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")

    def _on_left_arrow(self, event):
        self.canvas.xview_scroll(-1, "units")

    def _on_right_arrow(self, event):
        self.canvas.xview_scroll(1, "units")

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            df = pd.read_excel(file_path)

            # 自动识别包含 URL 的列
            url_column = None
            for column in df.columns:
                if df[column].astype(str).str.startswith(('http://', 'https://')).any():
                    url_column = column
                    break

            if url_column is None:
                messagebox.showerror("错误", "未找到包含 URL 的列")
                return

            self.urls = df[url_column].tolist()
            self.attributes = df.drop(columns=[url_column]).to_dict(orient='records')
            self.columns = df.columns.tolist()
            self.columns.remove(url_column)
            self.current_page = 0
            self.show_page()

    def show_page(self):
        start_index = self.current_page * self.images_per_page
        end_index = min(start_index + self.images_per_page, len(self.urls))

        for i in range(start_index, end_index):
            url = self.urls[i]
            attributes = self.attributes[i]

            response = requests.get(url)
            img = Image.open(BytesIO(response.content))
            img_res = img.resize((200, 200), Image.LANCZOS)
            img_tk = ImageTk.PhotoImage(img_res)

            frame = tk.Frame(self.scrollable_frame, bd=2, relief="groove", padx=10, pady=10, bg="#f9f9f9")
            frame.grid(row=i, column=0, padx=10, pady=10, sticky="w")

            img_label = tk.Label(frame, image=img_tk, bg="#f9f9f9")
            img_label.image = img_tk
            img_label.pack(side="left")

            img_label.bind("<Button-1>", partial(self.show_large_image, img))

            tree = ttk.Treeview(frame, columns=self.columns, show='headings', height=5)
            tree.pack(side="left", padx=10)
            for col in self.columns:
                tree.heading(col, text=col)
                tree.column(col, width=150, anchor='center')
            tree.insert('', 'end', values=[attributes[col] for col in self.columns])

    def show_large_image(self, img, event):
        top = tk.Toplevel(self.root)
        top.title("放大图片")

        img_large = img.resize((800, 800), Image.LANCZOS)
        img_large_tk = ImageTk.PhotoImage(img_large)

        label = tk.Label(top, image=img_large_tk)
        label.image = img_large_tk
        label.pack()

        top.update_idletasks()
        width = top.winfo_width()
        height = top.winfo_height()
        x = (top.winfo_screenwidth() // 2) - (width // 2)
        y = (top.winfo_screenheight() // 2) - (height // 2)
        top.geometry('{}x{}+{}+{}'.format(width, height, x, y))

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry('{}x{}+{}+{}'.format(width, height, x, y))

root = tk.Tk()
app = ImageApp(root)
root.mainloop()
