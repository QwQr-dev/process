# -*- coding: utf-8 -*-
#For Windows and only Windows supported
import os
import re
import sys
import tkinter as tk
from getpass import getuser
from winshell import shortcut
from wmi import WMI 
from tkinter import messagebox,Menu,filedialog,ttk,font,colorchooser
from tkinter.font import Font
import time
import ctypes
import psutil
import platform
import shutil
import keyboard
import subprocess
import logging
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception as e:
        messagebox.showerror('process',f'发生错误：{e}')
        return False
    
def restart_program():
    '''重启程序'''
    command = f'''
@echo off
taskkill /f /im "{os.path.basename(sys.argv[0])}" >nul 2>&1
timeout /t 5 
start "" "{s + os.path.basename(sys.argv[0])}"
exit /b
'''
    
    command_2 = f'''
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c {s}restart.bat", 0, True
Set WshShell = Nothing
'''
    
    a = messagebox.askquestion('process','你确定要重启程序吗？')
    if a == 'yes':
        # 通过调用vbs文件（vbs文件主要是隐藏bat文件在运行后打开的命令提示窗口）来运行bat文件，以达到重启的目的
        with open(f'{s}restart.bat','w',encoding='utf-8') as f:
            f.write(command)
        with open(f'{s}restart.vbs','w',encoding='utf-8') as f:
            f.write(command_2)
        subprocess.Popen(['cmd','/c','start','',f'{s}restart.vbs'])
    

def is_valid_path_syntax(path):
    """
    验证 Windows 路径语法，包含以下检查：
    1. 绝对路径检查
    2. 非法字符检查
    3. 保留名称检查（包括各级目录）
    4. 路径长度限制
    5. 文件名/目录名格式
    6. 特殊结尾检查（空格/点结尾）
    7. 驱动器名/UNC路径格式验证
    """
    # 空路径检查
    if not path.strip():
        messagebox.showerror('process','路径不能为空')
        return False

    try:
        path = os.path.normpath(path)
    except ValueError as e:
        messagebox.showerror('process',f'发生错误：{e}')
        return False

    # 必须为绝对路径
    if not os.path.isabs(path):
        messagebox.showerror('process','提供的路径必须为绝对路径')
        return False

    # 路径长度检查（260字符限制）
    if len(path) > 259:
        messagebox.showerror('process','路径长度最大限制为260字符')
        return False

    # 分解驱动器/UNC路径
    drive, tail = os.path.splitdrive(path)
    is_unc = path.startswith("\\\\")

    # UNC路径验证
    if is_unc:
        # 分割UNC路径为组成部分（排除空字符串）
        unc_parts = [p for p in path.split("\\")[2:] if p]
        # 必须包含至少服务器名和共享名
        if len(unc_parts) < 2 or not unc_parts[0] or not unc_parts[1]:
            messagebox.showerror('process','路径不为UNC路径基础格式')
            return False
    # 驱动器路径验证
    elif drive:
        # 驱动器格式必须为单字母加冒号
        if not re.fullmatch(r"[A-Za-z]:", drive):
            messagebox.showerror('process','驱动器格式不正确')
            return False
    # 非UNC非驱动器的绝对路径（理论上不会出现，因为os.path.isabs已过滤）
    else:
        messagebox.showerror('process','驱动器格式不正确')
        return False

    # 分割路径为层级结构（防止死循环）
    parts = []
    while True:
        new_tail, part = os.path.split(tail)
        if new_tail == tail:  # 无法继续分割时退出
            break
        tail = new_tail
        if part:
            parts.insert(0, part)

    # 添加根标识符到层级列表
    if drive:
        parts.insert(0, drive)
    elif is_unc:
        parts.insert(0, "\\\\")

    # 保留名称集合
    RESERVED_NAMES = {
        'CON', 'PRN', 'AUX', 'NUL',
        'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
        'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
    }

    # 非法字符集合（排除路径分隔符）
    INVALID_CHARS = {'<', '>', ':', '"', '|', '?', '*'}

    for part in parts:
        # 跳过根标识符（如 C: 或 \\）
        if part in (drive, "\\\\"):
            continue

        # 根目录特殊处理（如 C:\ 分割后的空部分）
        if part == os.sep or not part:
            continue

        # 非法字符检查
        if any(c in INVALID_CHARS for c in part):
            messagebox.showerror('process',f'文件夹名或文件名不得含有以下字符：{INVALID_CHARS}')
            return False

        # 保留名称检查（不区分扩展名）
        name_Windows = part.split('.')[0].upper()
        if name_Windows in RESERVED_NAMES:
            messagebox.showerror('process',f'文件夹名或文件名不得含有以下字符：{RESERVED_NAMES}')
            return False

        # 特殊结尾检查（空格或点）
        if part[-1] in (' ', '.'):
            messagebox.showerror('process','文件扩展名不得只含有“.”或空格')
            return False

        # 全空格或空名称检查
        if part.strip() == '':
            messagebox.showerror('process','文件名不能为空')
            return False

        # 单级长度限制（255字符）
        if len(part) > 255:
            messagebox.showerror('process','字符串长度最大为255字符')
            return False
    return path

def safe_create_directory(directory_total):
    """
    安全创建目录的增强方法，包含以下改进：
    1. 使用Python原生目录创建方法
    2. 完善的错误处理机制
    3. 路径规范化处理
    4. 用户友好提示
    5. 日志记录
    """
    try:
        # 路径规范化处理（重要安全步骤）
        normalized_path = os.path.normpath(directory_total)
        
        # 先验证路径语法有效性（使用之前完善的路径验证函数）
        if not is_valid_path_syntax(normalized_path):
            raise ValueError(f"无效的路径格式: {normalized_path}")

        # 使用更安全的makedirs方法（自动创建父目录）
        os.makedirs(normalized_path, exist_ok=True)
        logging.info(f"成功创建或确认目录存在: {normalized_path}")
        
    except FileExistsError:
        # 当exist_ok=False时可能触发（虽然当前参数设置为True）
        logging.debug(f"目录已存在: {normalized_path}")
    except PermissionError as e:
        logging.error(f"权限不足无法创建目录: {normalized_path}")
        messagebox.showerror("process", f"没有权限创建目录:\n{normalized_path}")
        return False
    except OSError as e:
        logging.error(f"创建目录失败: {e}")
        messagebox.showerror("process", 
            f"无法创建目录:\n{normalized_path}\n"
            f"错误类型: {type(e).__name__}\n"
            f"详细信息: {str(e)}")
        return False
    except Exception as e:
        logging.critical(f"未知错误: {e}")
        messagebox.showerror("process", 
            f"发生意外错误:\n{str(e)}")
        return False
    return True

def other_options():
    '''其他选项'''
    def change_directory_total():
        '''更改默认路径'''
        try:
            file_path = filedialog.askdirectory(title='更改默认路径')
        except Exception as e:
            messagebox.showerror('process',f'发生错误：{e}')
        else:
            if file_path:
                file_path = file_path.split('/') # 根据'/'分割字符串
                file_path = '\\'.join(file_path)   # 重新连接字符串
                with open(f'{s}directory','w',encoding='utf-8') as f:
                    f.write(file_path)
                messagebox.showinfo('process',f'路径已更改为（重启后生效）：\n{file_path}')
        return

    def restore_change_directory_total():
        '''还原默认路径'''
        self = messagebox.askquestion('process','你确定要还原为默认路径吗？')
        if self == 'yes':
            with open(f'{s}directory','w',encoding='utf-8') as f:
                f.write('')
            messagebox.showinfo('process',f'已还原为默认路径（重启后生效）：\n C:\\Users\\{username}\\AppData\\processdata')
        return
    
    class WordOptions:
        '''字体设置'''
        def __init__(self):
            pass

        @staticmethod
        def update_font_size(val):
            text_font.configure(size=int(float(val)))

        @staticmethod
        def change_color(color_type):
            global change_word_color
            global change_background_color

            color = colorchooser.askcolor()[1]
            if color:
                other_text.config(**{color_type: color})
            if color_type == 'fg':
                if color == '':
                    return
                change_word_color = color
                return change_word_color
            elif color_type == 'bg':
                if color == '':
                    return
                change_background_color = color
                return change_background_color
            return

        @staticmethod
        def update_font(event):
            selected_font = font_combo.get()
            size_font = int(size_slider.get())
            try:
                # 尝试应用字体
                text_font.config(family=selected_font,size=size_font)
            except tk.TclError:
                messagebox.showerror("错误", f"字体 {selected_font} 不可用！")
        
        @staticmethod
        def save_update_font():
            a = messagebox.askquestion('process','你确定要更改吗？')
            if a == 'yes':
                selected_font = font_combo.get()
                size_font = int(size_slider.get())
                with open(f'{s}wordsetting.txt','w',encoding='utf-8') as f:
                    f.write(f'wordtype:{selected_font}\n')
                    f.write(f'wordsize:{size_font}\n')
                    f.write(f'wordcolor:{change_word_color}\n')
                    f.write(f'backgroundcolor:{change_background_color}')
                messagebox.showinfo('process','更改成功（重启后生效）')

        @staticmethod
        def re_update_font():
            a = messagebox.askquestion('process','你确定要重置吗？')
            if a == 'yes':
                with open(f'{s}wordsetting.txt','w',encoding='utf-8') as f:
                    f.write('wordtype:Microsoft YaHei UI\n')
                    f.write(f'wordsize:11\n')
                    f.write(f'wordcolor:#000000\n')
                    f.write(f'backgroundcolor:#ffffff')
                messagebox.showinfo('process','重置成功（重启后生效）')

    try:
        with open(f'{s}word_context.txt','r',encoding='utf-8') as f:
            context = f.read()
    except Exception:
        context = ''

    other_options_window = tk.Toplevel(Windows)
    other_options_window.title("其他选项")
    other_options_window.geometry("960x600")
    other_options_window.transient(Windows)  # 设置副窗口与主窗口的关联
    other_options_window.grab_set()  # 捕获所有事件到副窗口

    """
    text_frame = tk.Frame(other_options_window,height=400,width=200)
    text_frame.pack_propagate(False)  # 禁止自动调整大小
    text.place(x=10)
    """

    other_text = tk.Text(other_options_window,height=100,width=60,font=text_font)
    other_text.place(x=0)
    other_text.insert(tk.END,context)
    color_type = 'bg'
    other_text.config(**{color_type: backgroundcolor})
    color_type = 'fg'
    other_text.config(**{color_type: wordcolor})

    directory_total_label = ttk.Label(other_options_window)
    directory_total_label.place(x=550,y=0)
    directory_total_label.config(text=f'当前的默认路径为：\n{directory_total}')

    change_directory_total_button = ttk.Button(other_options_window,text='更改默认路径',command=change_directory_total)
    change_directory_total_button.place(x=550,y=45)
    restore_change_directory_total_button = ttk.Button(other_options_window,text='还原为默认路径',command=restore_change_directory_total)
    restore_change_directory_total_button.place(x=650,y=45)
    restarting_button = ttk.Button(other_options_window,text='重新启动应用',command=restart_program)
    restarting_button.place(x=760,y=45)

    word_label = ttk.Label(other_options_window)
    word_label.place(x=550,y=90)
    word_label.config(text='字体设置')
    
    # 获取系统字体列表（按字母排序）
    system_fonts = sorted(font.families())

    # 创建带搜索功能的 Combobox
    font_combo = ttk.Combobox(
        other_options_window,
        values=system_fonts,
        width=40,
        font=(wordtype, wordsize)
    )
    font_combo.place(x=550,y=120)
    font_combo.bind("<<ComboboxSelected>>", WordOptions.update_font)        

    # 字体大小调节滑块
    size_control = ttk.Frame(other_options_window)
    size_control.place(x=550,y=150)
    ttk.Label(size_control, text="字体大小:").pack(side=tk.LEFT)
    size_slider = ttk.Scale(
        size_control,
        from_=1,
        to=20,
        command=WordOptions.update_font_size,
        orient=tk.HORIZONTAL
    )
    size_slider.set(wordsize)
    size_slider.pack(side=tk.RIGHT)

    # 颜色选择按钮
    ttk.Button(other_options_window, text="文字颜色", command=lambda: WordOptions.change_color("fg")).place(x=550,y=200)
    ttk.Button(other_options_window, text="背景颜色", command=lambda: WordOptions.change_color("bg")).place(x=650,y=200)

    save_word_change_button = ttk.Button(other_options_window,text='保存更改',command=WordOptions.save_update_font)
    save_word_change_button.place(x=550,y=230)
    re_word_change_button = ttk.Button(other_options_window,text='重置',command=WordOptions.re_update_font)
    re_word_change_button.place(x=650,y=230)

    return
    

def _directory_total():
    '''自定义路径，默认路径为：
    'C:\\Users\\username\\AppData\\processdata'

    （其中username为用户名）
    '''
    
    directory_total = f'C:\\Users\\{username}\\AppData\\processdata'
    if os.path.isfile(f'{s}directory'):
        with open(f'{s}directory','r',encoding='utf-8') as f:
            file_content = f.read()
            if file_content == '': # 如果未提供路径，则返回默认路径
                return directory_total
            else:
                file_content = is_valid_path_syntax(file_content)
                if file_content == False: # 如果提供的路径格式不正确，则返回False
                    messagebox.showinfo('process',f'将使用默认路径：\n{directory_total}')
                    return directory_total # 返回False后，将返回默认路径，否则返回提供的路径
                return file_content
    else:
        # 文件不存在时，创建一份空文件，并返回默认路径
        with open(f'{s}directory','w',encoding='utf-8') as f:
            pass 
        return directory_total

def about_sub_window():
    '''创建关于界面'''
    def _secrets():
        directory_secrets = f'{s}process.sys'
        if os.path.isfile(f'{s}process'):
            pass
        else:
            with open(f'{s}process','w',encoding='utf-8') as f:
                pass

        with open(f'{s}process','r',encoding='utf-8') as f:
            file = f.read()
        file = s + file
        if is_valid_path_syntax(file) == False: # 如果创建的目录不合法，则返回False
            os._exit(0)
        if os.path.isfile(directory_secrets):
                if file == '': # 如果文件中没有任何内容，则使用给予的统一文件名
                    os.rename(directory_secrets,f'{s}all.mp4')
                    subprocess.Popen(['cmd','/c','start','/max','',f'{s}all.mp4'])
                    return
                else:
                    os.rename(directory_secrets,file)
                    subprocess.Popen(['cmd','/c','start','/max','',file])
                    return
        else:
            # 如果文件被解析，则直接使用解析的文件
            if os.path.isfile(file): 
                subprocess.Popen(['cmd','/c','start','/max','',file])
            elif os.path.isfile(f'{s}all.mp4'):
                subprocess.Popen(['cmd','/c','start','/max','',f'{s}all.mp4'])
            return
        
    content = f'''                                   关于
作者在家闲的没事瞎做的
源代码将会在Github上开源
网址：

https://github.com/QwQr-dev/process

其中，

banned_processes.txt的默认路径为（可更改）：
C:\\Users\\username\\AppData\\processdata\\banned_processes.txt（其中username为用户名）

warmed_processes.txt的默认路径为（可更改）：
C:\\Users\\username\\AppData\\processdata\\warmed_processes.txt（其中username为用户名）

注意：

文件夹名或文件名不得含有以下字符：'CON', 'PRN', 'AUX', 'NUL',
                                'COM1', 'COM2', 'COM3', 'COM4',
                                'COM5', 'COM6', 'COM7', 'COM8', 
                                'COM9','LPT1', 'LPT2', 'LPT3',
                                'LPT4', 'LPT5', 'LPT6', 'LPT7', 
                                'LPT8', 'LPT9'

以及'<', '>', ':', '"', '|', '?', '*'


qWq
'''

    sub_window = tk.Toplevel(Windows)
    sub_window.title("关于")
    sub_window.geometry("800x480")
    sub_window.transient(Windows)  # 设置副窗口与主窗口的关联
    sub_window.grab_set()  # 捕获所有事件到副窗口

    about_text = tk.Text(sub_window,height=100,font=text_font)
    color_type = 'fg'
    about_text.config(**{color_type: wordcolor})
    color_type = 'bg'
    about_text.config(**{color_type: backgroundcolor})
    about_text.pack(side='left',expand=True)
    # 创建一个滚动条，并设置其command选项为text的yview方法
    about_text.insert(tk.END,content)
    about_text.config(state=tk.DISABLED)  # 设置为只读模式
    scrollbar = tk.Scrollbar(sub_window, command=text.yview)
    scrollbar.pack(side=tk.LEFT, fill=tk.Y)
    # 设置text小部件的yscrollcommand选项为scrollbar的set方法
    about_text.configure(yscrollcommand=scrollbar.set)

    secrets_button = ttk.Button(sub_window,text='更多',width=15,command=_secrets)
    secrets_button.pack(side='top')
    close_button = ttk.Button(sub_window,text='关闭',width=15,command=sub_window.destroy)
    close_button.pack(side='top')
    return

def information():
    '''获取设备信息'''
    system_information = f'{platform.platform()}[{platform.machine()}{str(platform.architecture())}]'  # 获取系统信息
    processors = wmi_obj.Win32_Processor()
    gpus = wmi_obj.Win32_VideoController()
    # 使用 psutil 获取实时使用情况
    mem = psutil.virtual_memory()
    physical_disks = wmi_obj.Win32_DiskDrive()
    computer_info = wmi_obj.Win32_ComputerSystem()[0]
    with open(directory_information,'w',encoding='utf-8') as file:
        print(f"计算机名: {computer_info.Name}",file=file)
        file.write(f"用户名：{username}\n")
        file.write(f'系统信息：{system_information}\n')
        for processor in processors:
            print("处理器ID：", processor.ProcessorID,file=file)
            print("处理器型号：", processor.Name,file=file)
            print("处理器架构：", processor.Architecture,file=file)
            print("处理器核心数：", processor.NumberOfCores,file=file)
        for gpu in gpus:
            print(f"GPU型号： {gpu.Name}",file=file)
            print(f"驱动版本：{gpu.DriverVersion}",file=file)
            print(f"GPU显存： {-int(gpu.AdapterRAM) / 1024**2:.2f} MB",file=file)
        print(f'内存总信息：',file=file)
        print(f"└[实时使用]",file=file)
        print(f"   总内存: {mem.total / (1024**3):.2f} GB",file=file)
        print(f"   已用内存: {mem.used / (1024**3):.2f} GB",file=file)
        print(f"   使用率: {mem.percent}%",file=file)
        # 使用 WMI 获取物理内存硬件信息
        mem_modules = wmi_obj.Win32_PhysicalMemory()
        total_gb = sum(int(mem.Capacity) for mem in mem_modules) / (1024**3)
        print(f"  [硬件信息]",file=file)
        print(f"   总容量: {total_gb:.2f} GB",file=file)
        for idx, mem in enumerate(mem_modules):
            print(f"   内存条 {idx + 1}:",file=file)
            print(f"   └厂商: {mem.Manufacturer}",file=file)
            print(f"     型号: {mem.PartNumber.strip()}",file=file)
            print(f"     频率: {mem.Speed} MHz",file=file)
        # 物理磁盘信息
        for p_disk in physical_disks:
            print(f"物理磁盘: {p_disk.Caption}",file=file)
            print(f"接口类型: {p_disk.InterfaceType}",file=file)
            print(f"总容量: {int(p_disk.Size) / (1024**3):.2f} GB",file=file)
            # 关联逻辑分区
            partitions = p_disk.associators("Win32_DiskDriveToDiskPartition")
            for partition in partitions:
                logical_disks = partition.associators("Win32_LogicalDiskToPartition")
                for l_disk in logical_disks:
                    print(f"└逻辑分区: {l_disk.Caption}",file=file)
                    print(f"  文件系统: {l_disk.FileSystem}",file=file)
                    print(f"  总容量: {int(l_disk.Size) / (1024**3):.2f} GB",file=file)
                    print(f"  剩余容量: {int(l_disk.FreeSpace) / (1024**3):.2f} GB",file=file)
    with open(directory_information,'r',encoding='utf-8') as f:
        return f.read()
    
def computer_information():
    def re_computer_information():
        '''更新设备信息'''
        self = messagebox.askquestion('设备基本信息','是否要更新设备信息？')
        if self == 'yes':
            sub_text.config(state=tk.NORMAL)
            sub_text.delete('1.0',tk.END)
            sub_text.insert(tk.END,information())
            sub_text.config(state=tk.DISABLED)
            messagebox.showinfo('设备基本信息','已更新设备信息')
        else:
            messagebox.showinfo('设备基本信息','已取消操作')
        return
    
    def save_computer_information():
        '''保存设备信息'''
        file_path = filedialog.asksaveasfilename(defaultextension='.txt',
                                            filetypes=[('Text files', '*.txt')],
                                            title='设备基本信息')
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(information())
                messagebox.showinfo("设备基本信息", f"已保存到 {file_path}")
            except Exception as e:
                messagebox.showerror("设备基本信息", f"无法保存文件：{e}")
        return
        
    sub_window = tk.Toplevel(Windows)
    sub_window.title("设备基本信息")
    sub_window.geometry("860x480")
    sub_window.resizable()
    sub_window.transient(Windows)  # 设置副窗口与主窗口的关联
    sub_window.grab_set()  # 捕获所有事件到副窗口

    sub_text = tk.Text(sub_window,height=20, width=80,font=text_font)
    sub_text.pack(side="left", fill=tk.BOTH)
    sub_text.insert(tk.END,information())
    sub_text.config(state=tk.DISABLED)  # 设置为只读模式
    # 创建一个滚动条，并设置其command选项为text的yview方法
    scrollbar = tk.Scrollbar(sub_window, command=text.yview)
    scrollbar.pack(side=tk.LEFT, fill=tk.Y)
    # 设置text小部件的yscrollcommand选项为scrollbar的set方法
    sub_text.configure(yscrollcommand=scrollbar.set)

    button1 = ttk.Button(sub_window,text='保存',width=15,command=save_computer_information)
    button1.pack(side='top')
    update_button = ttk.Button(sub_window, text="更新设备信息",width=15,command=re_computer_information)
    update_button.pack(side='top')
    button2 = ttk.Button(sub_window,text='关闭',width=15,command=sub_window.destroy)
    button2.pack(side='top')
    return

def new_file():
    '''创建一份新的文件，并返回文件路径'''
    global public_directory
    content = text.get('1.0',tk.END)
    try:
        with open(public_directory,'r',encoding='utf-8') as f:
            file = f.read()
            if content != f'{file}\n':
                self = messagebox.askquestion()
                if self == 'yes':
                    save_file()
                return
    except FileNotFoundError:   # 忽略路径不存在时的情况
        pass

    with open(directory_new_file,'w',encoding='utf-8') as f:
        pass
    text.insert(tk.END,'')
    text.delete('1.0',tk.END)
    text_file = directory_new_file
    file_name_label.config(text=f'已打开文件：\n{text_file.replace(directory_total + '\\','',1)}')
    text.config(state=tk.NORMAL)
    public_directory = directory_new_file
    return public_directory

def open_file():
    '''打开一份文件，并返回文件路径'''
    global public_directory
    file_path = filedialog.askopenfilename(defaultextension='.txt',
                                        filetypes=[('Text files', '*.txt')],
                                        title='打开文件')
    if file_path:
        try:
            with open(file_path,'r',encoding='utf-8') as f:
                file_content = f.read()
        except Exception as e:
                messagebox.showerror("process", f"无法打开文件：{e}")
        else:
            text.config(state=tk.NORMAL)
            text.delete('1.0',tk.END)
            text.insert(tk.END,file_content)   
            file_name_label.config(text=f'已打开文件：\n{file_path}')
            public_directory = file_path   
        return public_directory

def save_file():
    '''保存文件，如果要保存的文件为新文件，则返回文件路径'''
    global public_directory
    content = text.get('1.0',tk.END)
    if public_directory == directory_new_file:
        file_path = filedialog.asksaveasfilename(defaultextension='.txt',
                                            filetypes=[('Text files', '*.txt')],
                                            title='保存文件')
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                messagebox.showinfo('process', f"已保存到 {file_path}")
            except Exception as e:
                messagebox.showerror("process", f"无法保存文件：{e}")
            else:
                file_name_label.config(text=f'已打开文件：\n{file_path}')
                public_directory = file_path
            return public_directory
        
    if public_directory:
        with open(public_directory,'w',encoding='utf-8') as f:
            f.write(content)
        return
    
    if content == '\n':
        messagebox.showwarning('process','没有文件可保存')
        return
    
def save_as_file():
    '''将文件另存为新的文件'''
    content = text.get('1.0',tk.END)
    if public_directory == directory_new_file:
        messagebox.showerror('process','请先保存文件')
        return
    
    if content == '\n':
        while True: # 当文件路径存在时，跳出循环，并保存文件
            if public_directory:
                break
            else:
                messagebox.showwarning('process','没有文件可保存')
                return
            
    file_path = filedialog.asksaveasfilename(defaultextension='.txt',
                                            filetypes=[('Text files', '*.txt')],
                                            title='另存为')
    if file_path:
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            messagebox.showinfo("process", f"已保存到 {file_path}")
        except Exception as e:
            messagebox.showerror("process", f"无法保存文件：{e}")

def close_file():
    '''关闭文件'''
    global public_directory
    content = text.get('1.0',tk.END)
    try:
        with open(public_directory,'r',encoding='utf-8') as f:
            file_content = f.read()
    except FileNotFoundError:
        messagebox.showwarning('process','没有可关闭的文件')
        return

    if content != f'{file_content}\n':
        if public_directory == directory_new_file:
            self = messagebox.askyesnocancel('process','你想在关闭文件前保存文件吗？')
            if self == True:
                file_path = filedialog.asksaveasfilename(defaultextension='.txt',
                                            filetypes=[('Text files', '*.txt')],
                                            title='保存文件')
                if file_path:
                    try:
                        with open(file_path,'w',encoding='utf-8') as f:
                            f.write(content)
                    except Exception as e:
                        messagebox.showerror("process", f"无法保存文件：{e}")
                        return
                    else:
                        text.delete('1.0',tk.END)
                        text.config(state=tk.DISABLED)
                        file_name_label.config(text='')
                        public_directory = ''
                        messagebox.showinfo("process", f"已保存到 {file_path}")
                        return public_directory
                return None
            elif self == False:
                text.delete('1.0',tk.END)
                text.config(state=tk.DISABLED)
                file_name_label.config(text='')
                public_directory = ''
                return public_directory
            return None
        
        if public_directory:
            self = messagebox.askyesnocancel('process','你想在关闭文件前保存文件吗？')
            if self == True:
                with open(public_directory,'w',encoding='utf-8') as f:
                    f.write(content)
                text.delete('1.0',tk.END)
                text.config(state=tk.DISABLED)
                file_name_label.config(text='')
                public_directory = ''
                return public_directory
            elif self == False:
                text.delete('1.0',tk.END)
                text.config(state=tk.DISABLED)
                file_name_label.config(text='')
                public_directory = ''
                return public_directory
            return
    else:
        with open(public_directory,'w',encoding='utf-8') as f:
            f.write(content)
        text.delete('1.0',tk.END)
        text.config(state=tk.DISABLED)
        file_name_label.config(text='')
        public_directory = ''
        return public_directory
    
def undo():
    '''撤销'''
    try:
        text.edit_undo()
    except tk.TclError:
        messagebox.showwarning('process','无修改历史记录')

def redo():
    '''重做'''
    if public_directory == '':
        messagebox.showwarning('process','没有需要重做的文件')
        return
    self = messagebox.askquestion('process','你确定要重做吗？（无法撤销）')
    if self == 'yes':
        text.delete('1.0',tk.END)
        with open(public_directory,'w',encoding='utf-8') as f:
            f.write('')

def cleaned_directory():
    '''获取自身文件路径'''
    text = os.path.abspath(sys.argv[0]) # 获得当前的文件路径
    text = text.split('/') # 根据'/'分割字符串
    text = '\\'.join(text)   # 重新连接字符串
    cleaned_directory = text.rstrip(os.path.basename(sys.argv[0])) # 去掉文件名，获得干净的文件路径
    if is_valid_path_syntax(cleaned_directory) == False:
        sys.exit()
    return cleaned_directory

def temporary_operation_cess():
    if os.path.isfile(s + 'cess.exe'):
        command = f'''
  
'''
        
        with open(f'{s}start.bat','w',encoding='utf-8') as f:
            f.write(command)
        subprocess.Popen(['cmd','/c','start','',f'{s}start.bat'])
    else:
        messagebox.showerror('process',f'缺少以下文件：{s + 'cess.exe'}')

def temporary_operation_cessw():
    if os.path.isfile(s + 'cessw.exe'):
        command = f'''
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c {s}cessw.bat", 0, True
        '''

        command_2 = f'''
@echo off
start "" "{s}cessw.exe"
exit /b
        '''

        with open(f'{s}cessw.bat','w',encoding='utf-8') as file:
            file.write(command_2)
        with open(f'{s}cessw.vbs','w',encoding='utf-8') as f:
            f.write(command)
        subprocess.Popen(['cmd','/c','',f'{s}cessw.vbs'])
    else:
        messagebox.showerror('process',f'缺少以下文件：{s + 'cessw.exe'}')

def shortcutcess():
    def _shortcutcess(_target: str,
                        _shortcut_name: str,
                        _icon: str,
                        _description: str
                        ):
        '''
        创建快捷方式

        _target:提供需要生成快捷方式的程序（程序的名称）

        _shortcut_name:生成快捷方式后所提供的命名（快捷方式的名称）

        _icon:（可选）生成快捷方式后所引用的图标（必须为ico格式文件，并且图片尺寸不能太大）

        _description:（可选）快捷方式的描述
        '''

        # 设置快捷方式的目标路径
        target_path = s + _target
        # 设置快捷方式的名称
        shortcut_name = s + _shortcut_name
        # 设置快捷方式的图标路径
        icon_path = s + _icon
        # 创建快捷方式
        try:
            with shortcut(shortcut_name) as link:
                link.path = target_path
                link.description = _description
                if icon_path:
                    icon_index = 0
                    link.icon_location = (icon_path,icon_index)
                    time.sleep(3)
        except PermissionError:
            messagebox.showerror('process',f'没有足够的权限来访问以下目录：\n{s}')

    _shortcutcess('process.exe','process.lnk','icon-console.ico',"This is a useless application.")
    _shortcutcess('cessw.exe','cessw.lnk','icon-console.ico',"This is a useless application.")

def self_starting():
    '''设置开机自启动'''
    if os.path.isfile(f'{s}cessw.lnk') and os.path.isfile(f'{s}process.lnk'):
        if os.path.isfile(f'{directory_startup}\\cessw.lnk') and os.path.isfile(f'{directory_startup}\\process.lnk'):
            messagebox.showinfo('process','已设置开机自启动，无需操作')
        else:
            a = messagebox.askquestion('process','是否要添加开机自启动？')
            if a == 'yes':
                shutil.copy(f'{s}cessw.lnk',directory_startup)
                shutil.copy(f'{s}process.lnk',directory_startup)
                messagebox.showinfo('process','下次开机程序(cess.exe和process.exe)将会自启动')
    else:
        shortcutcess()
        if os.path.isfile(f'{directory_startup}\\cessw.lnk') and os.path.isfile(f'{directory_startup}\\process.lnk'):
            messagebox.showinfo('process','已设置开机自启动，无需操作')
        else:
            a = messagebox.askquestion('process','是否要添加开机自启动？')
            if a == 'yes':
                shutil.copy(f'{s}cessw.lnk',directory_startup)
                shutil.copy(f'{s}process.lnk',directory_startup)
                messagebox.showinfo('process','下次开机程序(cess.exe和process.exe)将会自启动')

def close_self_starting():
    '''关闭开机自启动'''
    if os.path.isfile(directory_startup + '\\cessw.lnk') and os.path.isfile(directory_startup + '\\process.lnk'):
        a = messagebox.askquestion('process','是否要关闭开机自启动？')
        if a == 'yes':
            os.remove(directory_startup + '\\cessw.lnk')
            os.remove(directory_startup + '\\process.lnk')
            messagebox.showinfo('process','下次开机程序(cess.exe和process.exe)不会自启动')
    else:
        messagebox.showinfo('process','已关闭开机自启动，无需操作')

def only_self_starting_cess():
    '''设置开机自启动(仅cessw.exe)'''
    if os.path.isfile(f'{s}cessw.exe'):
        self = messagebox.askquestion('process','是否要添加开机自启动(仅cessw.exe)？')
        if self == 'yes':
            if os.path.isfile(f'{s}cess.lnk'):
                if os.path.isfile(f'{directory_startup}\\cessw.lnk'):
                    messagebox.showinfo('process','已设置开机自启动(仅cessw.exe)，无需操作')
                else:
                    shutil.copy(f'{s}cessw.lnk',directory_startup)
                    messagebox.showinfo('process','添加成功，下次开机将会自启动cessw.exe')
            else:
                if os.path.isfile(f'{directory_startup}\\cessw.lnk'):
                    messagebox.showinfo('process','已设置开机自启动(仅cessw.exe)，无需操作')
                else:
                    shortcutcess()
                    shutil.copy(f'{s}cessw.lnk',directory_startup)
                    messagebox.showinfo('process','添加成功，下次开机将会自启动cessw.exe')
    else:
        messagebox.showerror('process',f'缺少以下文件：{s + 'cessw.exe'}')

def only_close_self_starting_cess():
    '''关闭开机自启动(仅cessw.exe)'''
    if os.path.isfile(f'{directory_startup}\\cessw.lnk'):
        self = messagebox.askquestion('process','是否要关闭开机自启动(仅cessw.exe)？')
        if self == 'yes':
            if os.path.isfile(f'{directory_startup}\\cessw.lnk'):
                os.remove(f'{directory_startup}\\cessw.lnk')
                messagebox.showinfo('process','下次开机程序(仅cessw.exe)不会自启动')
    else:
        messagebox.showinfo('process','已关闭开机自启动(仅cessw.exe)，无需操作')

def secrets():
    self = messagebox.askquestion('process','你确定吗？')
    if self == 'yes':
        os.system(r'shutdown /s /t 20')
        messagebox.showwarning('process','都说了不要点，这下好了吧，设备将在20秒后关机')
    else:
        messagebox.showinfo('process','已取消操作')

def open_banned_processes():
    global public_directory
    text.delete('1.0',tk.END)
    try:
        file = open(directory_banned_processes,'r',encoding='utf-8')
        file_path = file.read()
    except PermissionError:
        messagebox.showerror('process','没有足够的权限访问该文件')
        return
    except FileNotFoundError:
        messagebox.showerror('process','文件不存在')
        self = messagebox.askquestion('process','是否新建banned_processes.txt？')
        if self == 'yes':
            with open(directory_banned_processes,'w',encoding='utf-8') as f:
                pass
            text.config(state=tk.NORMAL)
            file_name_label.config(text=f'已打开文件：\n{directory_banned_processes}')
            public_directory = directory_banned_processes
        return public_directory
    except Exception as e:
            messagebox.showerror("process", f"无法打开文件：{e}")
    else:
        text.config(state=tk.NORMAL)
        text.insert(tk.END,file_path)
        file_name_label.config(text=f'已打开文件：\n{directory_banned_processes}')
        public_directory = directory_banned_processes
        return public_directory

def open_warmed_processes():
    global public_directory
    text.delete('1.0',tk.END)
    try:
        file = open(directory_warmed_processes,'r',encoding='utf-8')
        file_path = file.read()
    except PermissionError:
        messagebox.showerror('process','没有足够的权限访问该文件')
        return
    except FileNotFoundError:
        messagebox.showerror('process','文件不存在')
        self = messagebox.askquestion('process','是否新建warmed_processes.txt？')
        if self == 'yes':
            with open(directory_warmed_processes,'w',encoding='utf-8') as f:
                pass
            text.config(state=tk.NORMAL)
            file_name_label.config(text=f'已打开文件：\n{directory_warmed_processes}')
            public_directory = directory_warmed_processes
        return public_directory
    except Exception as e:
        messagebox.showerror("process", f"无法打开文件：{e}")
    else:
        text.config(state=tk.NORMAL)
        text.insert(tk.END,file_path)  
        file_name_label.config(text=f'已打开文件：\n{directory_warmed_processes}')
        public_directory = directory_warmed_processes
        return public_directory

def open_Microsoft_keyboard():
    subprocess.Popen(['osk.exe'],shell=True)

def _open_Microsoft_keyboard():
    subprocess.Popen(['C:\\Program Files\\Common Files\\microsoft shared\\ink\\TabTip.exe'],shell=True)
    try:
        subprocess.run(['cmd','/c','taskkill /f /im TabTip.exe'])
    except Exception as e:
        messagebox.showerror('process',f'发生错误：{e}')

def close_temporary_operation_cessw():
    try:
        subprocess.run(['cmd','/c','taskkill /f /im cessw.exe'])
    except subprocess.CalledProcessError as e:
        messagebox.showerror('process',f'发生错误：{e}')

def Windows_quit():
    '''退出程序'''
    global public_directory
    content = text.get('1.0',tk.END)
    try:
        with open(public_directory,'r',encoding='utf-8') as f:
            file = f.read()
    except FileNotFoundError:
        Windows.quit()
    else:
        if content != f'{file}\n':
            if close_file() == None:
                return
            Windows.quit()
        else:
            Windows.quit()

def cmd():
    '''打开命令提示符'''
    os.system(r'start cmd.exe')

def powershell():
    '''打开PowerShell'''
    os.system(r'start powershell.exe')

def main():
    '''定义主函数'''
    def _keyboard():
        """监听组合键"""
        keyboard.add_hotkey('ctrl+n',new_file)
        keyboard.add_hotkey('ctrl+o',open_file)
        keyboard.add_hotkey('ctrl+s',save_file)
        keyboard.add_hotkey('ctrl+shift+s',save_as_file)
        keyboard.add_hotkey('ctrl+w',close_file)
        keyboard.add_hotkey('ctrl+z',undo)
        keyboard.add_hotkey('ctrl+y',redo)
        
    def cut_text():
        text.event_generate("<<Cut>>")

    def copy_text():
        text.event_generate("<<Copy>>")

    def paste_text():
        text.event_generate("<<Paste>>")
    
    def open_app():
        file_path = filedialog.askopenfilename(defaultextension='',
                                        filetypes=[('应用程序', '*.exe'),('快捷方式','*.lnk')],
                                        title='打开一个应用')
        if file_path:
            try:
                subprocess.run(['cmd','/c','start','',file_path])
            except Exception as e:
                messagebox.showerror('process',f'无法打开文件：{e}')

    # 创建一个顶层菜单
    menu_bar = Menu(Windows)
    # 创建一个文件菜单，并添加命令
    file_menu = Menu(menu_bar, tearoff=0)
    file_menu.add_command(label="新建文件(Ctrl+N)", command=new_file)
    file_menu.add_command(label="打开文件(Ctrl+O)", command=open_file)
    file_menu.add_command(label="保存文件(Ctrl+S)", command=save_file)
    file_menu.add_command(label='另存为(Ctrl+Shift+S)',command=save_as_file)
    file_menu.add_command(label='关闭文件(Ctrl+W)',command=close_file)
    file_menu.add_separator()  # 添加分隔线
    file_menu.add_command(label='打开banned_processes.txt',command=open_banned_processes)
    file_menu.add_command(label='打开warmed_processes.txt',command=open_warmed_processes)
    menu_bar.add_cascade(label="文件", menu=file_menu)  # 将文件菜单添加到菜单栏

    # 创建一个编辑菜单，并添加命令
    edit_menu = Menu(menu_bar, tearoff=0)
    edit_menu.add_command(label='剪切(Ctrl+X)',command=cut_text)
    edit_menu.add_command(label='复制(Ctrl+C)',command=copy_text)
    edit_menu.add_command(label='粘贴(Ctrl+V)',command=paste_text)
    edit_menu.add_command(label="撤销(Ctrl+Z)", command=undo)
    edit_menu.add_command(label="重做(Ctrl+Y)", command=redo)
    menu_bar.add_cascade(label="编辑", menu=edit_menu)  # 将编辑菜单添加到菜单栏

    # 创建一个帮助菜单，并添加命令
    help_menu = Menu(menu_bar, tearoff=0)
    help_menu.add_command(label='设备基本信息',command=computer_information)
    help_menu.add_command(label='不要点击！！！',command=secrets)
    help_menu.add_command(label='打开command(CMD)',command=cmd)
    help_menu.add_command(label='打开shell(PowerShell)',command=powershell)
    help_menu.add_command(label="关于", command=about_sub_window)
    menu_bar.add_cascade(label="帮助", menu=help_menu)  # 将帮助菜单添加到菜单栏

    open_keybord = Menu(menu_bar,tearoff=0)
    open_keybord.add_command(label='系统内置(osk)',command=open_Microsoft_keyboard)
    open_keybord.add_command(label='系统内置(TabTip)',command=_open_Microsoft_keyboard)
    menu_bar.add_cascade(label="软键盘",menu=open_keybord)

    menu_bar.add_cascade(label='其他选项',command=other_options)

    menu_bar.add_cascade(label="退出", command=Windows_quit)  # 退出程序
    # 显示菜单栏
    Windows.config(menu=menu_bar)

    _keyboard()

    text.pack(side="left", fill=tk.BOTH,expand=True)
    scrollbar = tk.Scrollbar(Windows, command=text.yview)
    scrollbar.pack(side=tk.LEFT, fill=tk.Y)
    # 设置text小部件的yscrollcommand选项为scrollbar的set方法
    text.configure(yscrollcommand=scrollbar.set)

    temporary_operation_cess_button = ttk.Button(Windows, text="临时运行cess.exe",width=50,command=temporary_operation_cess)
    temporary_operation_cess_button.pack(side="top")
    temporary_operation_cessw_button = ttk.Button(Windows, text="临时运行cessw.exe",width=50,command=temporary_operation_cessw)
    temporary_operation_cessw_button.pack(side="top")
    close_temporary_operation_cessw_button = ttk.Button(Windows,text="关闭cessw.exe",width=50,command=close_temporary_operation_cessw)
    close_temporary_operation_cessw_button.pack(side='top')
    self_starting_cessw_button = ttk.Button(Windows, text="开机自启动(仅cessw.exe)",width=50,command=only_self_starting_cess)
    self_starting_cessw_button.pack(side="top")
    close_self_starting_cessw_button = ttk.Button(Windows, text="关闭开机自启动(仅cessw.exe)",width=50,command=only_close_self_starting_cess)
    close_self_starting_cessw_button.pack(side="top")
    self_starting_all_button = ttk.Button(Windows,text='开机自启动',width=50,command=self_starting)
    self_starting_all_button.pack(side="top")
    close_self_starting_all_button = ttk.Button(Windows,text='关闭开机自启动',width=50,command=close_self_starting)
    close_self_starting_all_button.pack(side="top")
    open_app_button = ttk.Button(Windows,text='打开一个应用',width=50,command=open_app)
    open_app_button.pack(side='top')
        
wmi_obj = WMI()
username = getuser()    # 获取当前用户的用户名
s = cleaned_directory()
directory_total = _directory_total()
if safe_create_directory(directory_total) == False:
    os._exit(0)

directory_banned_processes = f'{directory_total}\\banned_processes.txt'
directory_warmed_processes = f'{directory_total}\\warmed_processes.txt'
directory_information = f'{directory_total}\\information'
directory_startup = f'C:\\Users\\{username}\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup'
directory_new_file = f'{directory_total}\\Untitled-1'
public_directory = ''
change_word_color = ''
change_background_color = ''

if __name__ == '__main__':
    if is_admin():
        # 已经是管理员，执行你的代码
        Windows = tk.Tk()
        Windows.geometry('1200x600')
        Windows.title("process")

        # 初始化变量（可选默认值）
        wordtype = "Microsoft YaHei UI"    # 默认字体
        wordsize = 11       # 默认字号
        wordcolor = "#000000"  # 默认文本框字体颜色
        backgroundcolor = "#ffffff"  # 默认文本框背景颜色

        # 定义解析规则（键名: 转换函数）
        parsers = {
            'wordtype': str,
            'wordsize': int,
            'wordcolor': str,
            'backgroundcolor': str
        }

        try:
            with open(f'{s}wordsetting.txt', 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue  # 跳过空行
                    
                    # 使用正则匹配键值对（支持中英文冒号和空格）
                    match = re.match(r'^\s*(\w+)\s*[:：]\s*(.+?)\s*$', line)
                    if match:
                        key = match.group(1)
                        value_str = match.group(2)
                        
                        if key in parsers:
                            try:
                                # 根据键名转换数据类型
                                value = parsers[key](value_str)
                                # 更新变量
                                if key == 'wordtype':
                                    wordtype = value
                                elif key == 'wordsize':
                                    wordsize = value
                                elif key == 'wordcolor':
                                    wordcolor = value
                                elif key == 'backgroundcolor':
                                    backgroundcolor = value
                            except ValueError:
                                messagebox.showerror('process',f"格式错误：无法将 '{value_str}' 转换为 {parsers[key].__name__}（行内容：'{line}'）")
                        else:
                            messagebox.showerror('process',f"未知键名：'{key}'（行内容：'{line}'）")
                    else:
                        pass
        except FileNotFoundError:
            messagebox.showerror('process',"错误：未找到配置文件 'wordsetting.txt'，将使用默认字体")
            with open(f'{s}wordsetting.txt','w',encoding='utf-8') as f:
                pass
        except Exception as e:
            messagebox.showerror('process',f'发生其他错误：\n{e}')

        text_font = Font(family=wordtype, size=wordsize)
        text = tk.Text(Windows,undo=True,maxundo=500,height=20, width=100,state=tk.DISABLED,font=text_font)
        color_type = 'fg'
        text.config(**{color_type: wordcolor})
        color_type = 'bg'
        text.config(**{color_type: backgroundcolor})

        main()
        Windows.protocol("WM_DELETE_WINDOW", Windows_quit)

        file_name_label = ttk.Label(Windows)
        file_name_label.pack(side='top')
        Windows.mainloop()

    else:
        # 不是管理员，请求管理员权限
        if ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable,f'"{os.path.abspath(sys.argv[0])}"', None, 0) != 0: 
            pass
