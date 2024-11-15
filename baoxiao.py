ABOUT_TEXT = """
这是一个用于生成报销单的程序。            \n\n
作者：Neeko                               \n
版本：1.0.9                              \n
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
目前：可以实现移入文件进行pdf合并，但是行程单存在问题。需要添加新的功能，对每个图片进行手动裁剪。（在合并前）
    

"""
from PIL import Image, ImageTk
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook
import datetime
import shutil
import sys
import tkinterdnd2 as tkdnd
from tkinterdnd2 import TkinterDnD, DND_FILES  # 导入 TkinterDnD 和 DND_FILES
import shutil


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

        # +++++++++++++++++++++  选择路径部分 ++++++++++++++++++++++
        self.label = tk.Label(root, text="选择存放报销单的路径（默认：桌面）：")
        self.label.grid(row=0, column=0,columnspan=2, pady=10, sticky="w") 

        self.path_entry = tk.Entry(root, width=50)
        self.path_entry.grid(row=0, column=2,columnspan=4, pady=10,sticky="ew")
        self.path_entry.insert(0, os.path.join(os.path.expanduser("~"), "Desktop"))  # 默认路径设为桌面

        self.browse_button = tk.Button(root, text="浏览", command=self.browse)
        self.browse_button.grid(row=0, column=6, padx=10, pady=10)

        self.generate_button = tk.Button(root, text="生成报销单", command=self.generate_report)
        self.generate_button.grid(row=0, column=7, padx=10, pady=10)

        # +++++++++++++++++++++  选择表格部分 ++++++++++++++++++++++  
        columns = ["序号", "报销类别", "代码", "月", "日", "报销人", "客户名称", "地点", "招待对象及电话", "公司随行人员", "招待人数", "人民币", "备注"]
        self.table = ttk.Treeview(root, columns=columns, show="headings")
        # 设置表格列宽
        self.table.column("序号", width=50)
        self.table.column("报销类别", width=100)
        self.table.column("代码", width=50)
        self.table.column("月", width=50)
        self.table.column("日", width=50)
        self.table.column("报销人", width=100)
        self.table.column("客户名称", width=150)
        self.table.column("地点", width=100)
        self.table.column("招待对象及电话", width=150)
        self.table.column("公司随行人员", width=100)
        self.table.column("招待人数", width=100)
        self.table.column("人民币", width=100)
        self.table.column("备注", width=150)

        self.table.grid(row=1, column=0, columnspan=10, pady=10)

        for col in columns:
            self.table.heading(col, text=col)
        # ++++++++++++++++++++++++++++++++ 表格操作功能按键 ++++++++++++++++++++++++++++++

        self.button_frame = tk.Frame(root)      #   按键垂直排列功能
        self.button_frame.grid(row=2, column=3, rowspan=2, padx=10, pady=10, sticky="n")
                               
        self.add_button = tk.Button(self.button_frame, text="添加报销项目", command=self.add_entry)
        # self.add_button.grid(row=2, column=3, padx=10, pady=10)
        self.add_button.pack(pady=5)

        self.coming_soon_button = tk.Button(self.button_frame, text="删除该报销项目", command=self.show_coming_soon)
        #self.coming_soon_button.grid(row=2, column=4, columnspan=1, pady=10)
        self.coming_soon_button.pack(pady=5)



        # +++++++++++++++++++++  选择发票导入部分 ++++++++++++++++++++++
        #   拖拽框
        self.dnd_label = tk.Label(root, text="将文件拖动到此处", bg="lightgrey", width=40, height=10)
        self.dnd_label.grid(row=2, column=0, padx=10, pady=10,columnspan=2)
        self.dnd_label.drop_target_register(tkdnd.DND_FILES)
        self.dnd_label.dnd_bind('<<Drop>>', self.on_drop)

        self.files = []

        # 添加合并按钮
        self.merge_button = tk.Button(root, text="开始发票合并", command=self.merge_files)
        self.merge_button.grid(row=2, column=2, padx=10, pady=10)
    #++++++++++++++++++++++++++++++++++++++++++++++++++++
    def browse(self):
        path = filedialog.askdirectory()
        if path:
            self.path_entry.delete(0, tk.END)  # 确保没有重复路径
            self.path_entry.insert(0, path)

    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        target_path = self.path_entry.get()
        if not target_path:
            messagebox.showerror("错误", "请选择存放文件的路径")
            return

        # 创建原始发票文件夹
        self.original_invoices_path = os.path.join(target_path, "原始发票")
        os.makedirs(self.original_invoices_path, exist_ok=True)

        # 获取当前已有的文件数量，用于命名新的文件
        existing_files = [f for f in os.listdir(self.original_invoices_path) if f.endswith('.pdf')]
        start_index = len(existing_files) + 1

        for idx, file in enumerate(files, start=start_index):
            new_file_name = f"{idx}.pdf"
            new_file_path = os.path.join(self.original_invoices_path, new_file_name)
            shutil.copy(file, new_file_path)
            self.files.append(new_file_path)
            self.dnd_label.config(text="\n".join(self.files))





    def merge_files(self):
        target_path = self.path_entry.get()
        merge_output_path = os.path.join(target_path, "可打印发票")
        os.makedirs(merge_output_path, exist_ok=True)

        # 调用拼接程序，并传递原始发票文件夹路径和合并输出路径
        merge_script_path = './pingjie.py'  # 相对路径
        command = f'python {merge_script_path} "{self.original_invoices_path}" "{merge_output_path}"'
        os.system(command)
        
        messagebox.showinfo("成功", "文件已合并并保存到 '可打印发票' 文件夹中")


    def show_coming_soon(self): 
        messagebox.showinfo("功能敬请期待", "该功能尚在开发中，敬请期待！")


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

#+++++++++++++++++++++++++++++++++ 添加功能部分 ++++++++++++++++++++++++++++++++++++++++++++++++++
    def add_entry(self):
        entry_window = tk.Toplevel(self.root)
        entry_window.title("添加报销条目")
        #   ++++++++++++++++++++++ 模版部分
        self.template_frame = tk.Frame(entry_window, bg="lightgrey") 
        self.template_frame.grid(row=0, column=4, rowspan=6, padx=10, pady=10, sticky="nsew") 

        self.template_label = tk.Label(self.template_frame, text="模板名称：") 
        self.template_label.grid(row=0, column=0, padx=10, pady=5) 

        self.template_name_entry = tk.Entry(self.template_frame) 
        self.template_name_entry.grid(row=1, column=0, padx=10, pady=5) 
        
        self.save_template_button = tk.Button(self.template_frame, text="保存模板", command=self.save_template) 
        self.save_template_button.grid(row=2, column=0, padx=10, pady=5) 
        
        self.template_listbox = tk.Listbox(self.template_frame) # 确保正确初始化 
        self.template_listbox.grid(row=3, column=0, padx=10, pady=5) 
        
        self.load_template_button = tk.Button(self.template_frame, text="使用模板", command=self.load_template) 
        self.load_template_button.grid(row=4, column=0, padx=10, pady=5) 
        
        self.templates = {} # 用于存储模板


        labels = ["报销类别", "代码", "月", "日", "报销人", "客户名称", "地点", "招待对象及电话", "公司随行人员", "招待人数", "人民币", "备注"]
        entries = {}
        row = 0

        for label in labels:                #   遍历
            tk.Label(entry_window, text=label).grid(row=row, column=0, padx=10, pady=5)
            if label == "报销类别":
                cb = ttk.Combobox(entry_window, values=["招待午餐费","招待晚餐费","招待娱乐费", "交通巴士费", "交通的士费", "交通过桥/过路费", "交通停车费", "交通油费", "出差机票费", "出差车船费", "出差住宿费", "出差餐费", "出差其他费", "通信费", "办公费", "研发费"])
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
        tk.Button(entry_window, text="添加到列表中", command=save_entry).grid(row=row, column=0, columnspan=2, pady=10)

    def save_template(self): 
        name = self.template_name_entry.get() 
        if name: 
            template_data = [] 
            for label in self.table.get_children(): 
                values = self.table.item(label, "values") 
                template_data.append(values) 
            self.templates[name] = template_data 
            self.template_listbox.insert(tk.END, name) 
            messagebox.showinfo("成功", f"模板 '{name}' 已保存") 
        else: 
            messagebox.showerror("错误", "请填写模板名称") 

    def load_template(self): 
        selected_template = self.template_listbox.curselection() 
        if selected_template: 
            name = self.template_listbox.get(selected_template) 
            template_data = self.templates[name] 
            for item in self.table.get_children(): 
                self.table.delete(item) 
            for values in template_data: 
                self.table.insert("", "end", values=values) 
            messagebox.showinfo("成功", f"模板 '{name}' 内容已填入") 
        else: 
            messagebox.showerror("错误", "请选择一个模板")

    def update_code(self, entries):
        category = entries["报销类别"].get()
        code_mapping = {
            "招待午餐费": "A1","招待晚餐费": "A2","招待娱乐费": "A3", 
            "交通巴士费": "B1", "交通的士费": "B2", "交通过桥/过路费": "B3", "交通停车费": "B4", "交通油费": "B5", 
            "出差机票费": "C1", "出差车船费": "C2", "出差住宿费": "C3", "出差餐费": "C4", "出差其他费": "C5", 
            "通信费": "D", "办公费": "E", "研发费": "F"
            }
        if category in code_mapping:
            code_entry = entries["代码"]
            code_entry.delete(0, tk.END) 
            code_entry.insert(0, code_mapping[category])

    def show_about(self):
        messagebox.showinfo("关于本软件", ABOUT_TEXT)

if __name__ == "__main__":
    # root = TkinterDnD.Tk()
    # icon_path = resource_path('Neeko.ico')  # 确保图标文件在当前目录下
    # root.iconbitmap(icon_path)
    
    # 修改工作目录到脚本所在目录 
    os.chdir(os.path.dirname(os.path.abspath(__file__))) 
    
    root = TkinterDnD.Tk() 
    icon_path = resource_path('Neeko.ico') # 确保图标文件在当前目录下 
    # 检查图标文件是否存在 
    if os.path.exists(icon_path): 
        root.iconbitmap(icon_path) 
    else: 
        print(f"图标文件不存在: {icon_path}")

    app = ReimbursementApp(root)
    root.mainloop()
