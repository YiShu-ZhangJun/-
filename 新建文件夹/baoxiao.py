ABOUT_TEXT = """
这是一个用于生成报销单的程序。            \n\n
作者：Neeko                               \n
版本：1.0.4                              \n
时间：2024-10-26                     \n
联系方式：小破站：妮蔻大王现在可帅气惹（全平台同名）           \n
功能：可以添加路径并且在填写内容后自动生成对应表格                \n
计划：                    \n
1、希望加上识别发票的功能，之后可以依据发票自动填入            \n
2、自动保存，软件会将现在的数据保存。下一次打开后可以恢复上一次数据         \n
3、模版            \n
4、编辑表格界面的数据，进行微调（或删除）              \n
5、自动生成发票打印文件，并且可以识别打印机开始打印相关数据       \n
6、新手教程与粘贴方式   \n
7、自动打包 \n
    

"""
from PIL import Image, ImageTk
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook
import datetime
import shutil
import sys

def resource_path(relative_path):
# """ 获取资源文件的路径，适用于打包后的应用 """
    return os.path.join(os.path.abspath("."), relative_path)


class ReimbursementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("报销程序")

        # 设置背景图
        bg_path = resource_path('background.png')
        self.background_image = ImageTk.PhotoImage(Image.open(bg_path))
        self.background_label = tk.Label(root, image=self.background_image)
        self.background_label.place(relwidth=1, relheight=1)

        # 创建菜单栏
        menu_bar = tk.Menu(root)
        root.config(menu=menu_bar)

        # 创建“文件”菜单
        file_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="选择路径", command=self.browse)
        file_menu.add_command(label="添加", command=self.add_entry)
        file_menu.add_command(label="生成", command=self.generate_report)

        # 创建“关于”菜单
        about_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="关于本软件", menu=about_menu)
        about_menu.add_command(label="关于", command=self.show_about)

        self.label = tk.Label(root, text="选择存放报销单的路径（默认：桌面）：")
        self.label.pack(padx=10, pady=10)

        self.path_entry = tk.Entry(root, width=50)
        self.path_entry.pack(padx=10, pady=10)
        self.path_entry.insert(0, os.path.join(os.path.expanduser("~"), "Desktop"))  # 默认路径设为桌面

        self.browse_button = tk.Button(root, text="浏览", command=self.browse)
        self.browse_button.pack(pady=10)

        self.generate_button = tk.Button(root, text="生成报销单", command=self.generate_report)
        self.generate_button.pack(pady=10)

        columns = ["序号", "报销类别", "代码", "月", "日", "报销人", "客户名称", "地点", "招待对象及电话", "公司随行人员", "招待人数", "人民币", "备注"]
        self.table = ttk.Treeview(root, columns=columns, show="headings")
        self.table.pack(padx=10, pady=10)

        for col in columns:
            self.table.heading(col, text=col)

        self.add_button = tk.Button(root, text="添加", command=self.add_entry)
        self.add_button.pack(pady=10)


    def browse(self):
        path = filedialog.askdirectory()
        if path:
            self.path_entry.delete(0, tk.END)  # 确保没有重复路径
            self.path_entry.insert(0, path)

    def generate_report(self):
        save_path = self.path_entry.get()
        if not save_path:
            messagebox.showerror("错误", "请选择存放报销单的路径")
            return

        try:
            template_path = os.path.join(os.path.dirname(__file__), "报销单.xlsx")
            report_path = os.path.join(save_path, "报销单_生成.xlsx")

            # 检查是否存在同名文件，若存在则追加标号
            counter = 1
            base_report_path = report_path
            while os.path.exists(report_path):
                report_path = base_report_path.replace(".xlsx", f"_{counter}.xlsx")
                counter += 1

            shutil.copy(template_path, report_path)  # 复制模板文件

            workbook = load_workbook(report_path)
            sheet = workbook.active
            
            # 假设数据从第11行开始写入
            start_row = 11
            for idx, item in enumerate(self.table.get_children()):
                values = self.table.item(item, "values")
                for col_num, value in enumerate(values, 1):
                    cell = chr(65 + col_num - 1) + str(start_row + idx)
                    sheet[cell] = value

            workbook.save(report_path)
            messagebox.showinfo("成功", "报销单生成成功！")
        except Exception as e:
            messagebox.showerror("错误", f"生成报销单时出错：{str(e)}")

    def add_entry(self):
        entry_window = tk.Toplevel(self.root)
        entry_window.title("添加报销条目")

        labels = ["报销类别", "代码", "月", "日", "报销人", "客户名称", "地点", "招待对象及电话", "公司随行人员", "招待人数", "人民币", "备注"]
        entries = {}
        row = 0

        for label in labels:
            tk.Label(entry_window, text=label).grid(row=row, column=0, padx=10, pady=5)
            if label == "报销类别":
                cb = ttk.Combobox(entry_window, values=["招待费", "交通费", "出差费", "通信费", "办公费", "研发费"])
                cb.grid(row=row, column=1, padx=10, pady=5)
                cb.bind("<<ComboboxSelected>>", lambda event: self.update_code(entries))
                entries[label] = cb
            elif label == "月":
                cb = ttk.Combobox(entry_window, values=[str(i) for i in range(1, 13)])
                cb.set(datetime.datetime.now().month)
                cb.grid(row=row, column=1, padx=10, pady=5)
                entries[label] = cb
            elif label == "日":
                cb = ttk.Combobox(entry_window, values=[str(i) for i in range(1, 32)])
                cb.set(datetime.datetime.now().day)
                cb.grid(row=row, column=1, padx=10, pady=5)
                entries[label] = cb
            else:
                entry = tk.Entry(entry_window)
                entry.grid(row=row, column=1, padx=10, pady=5)
                entries[label] = entry
            row += 1

        def save_entry():
            values = [self.table.get_children().__len__() + 1]
            for label in labels:
                values.append(entries[label].get())
            self.table.insert("", "end", values=values)
            entry_window.destroy()

        tk.Button(entry_window, text="保存", command=save_entry).grid(row=row, column=0, columnspan=2, pady=10)

    def update_code(self, entries):
        category = entries["报销类别"].get()
        code_mapping = {"招待费": "A", "交通费": "B", "出差费": "C", "通信费": "D", "办公费": "E", "研发费": "F"}
        if category in code_mapping:
            code_entry = entries["代码"]
            code_entry.delete(0, tk.END)
            code_entry.insert(0, code_mapping[category])

    def show_about(self):
        messagebox.showinfo("关于本软件", ABOUT_TEXT)

if __name__ == "__main__":
    root = tk.Tk()
    icon_path = resource_path('Neeko.ico')  # 确保图标文件在当前目录下
    root.iconbitmap(icon_path)
    app = ReimbursementApp(root)
    root.mainloop()
