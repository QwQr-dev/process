#coding = 'utf-8'
#For Windows and only Windows supported
import os
import re
import sys
import time
import ctypes
import getpass
import subprocess
from tkinter import messagebox
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception as e:
        messagebox.showerror('process',f'发生错误：{e}')
        return False
if is_admin():
    # 已经是管理员，执行你的代码

    def cleaned_directory():
        '''获取自身文件路径'''
        text = os.path.abspath(sys.argv[0]) # 获得当前的文件路径
        text = text.split('/') # 根据'/'分割字符串
        text = '\\'.join(text)   # 重新连接字符串
        cleaned_directory = text.rstrip(os.path.basename(sys.argv[0])) # 去掉文件名，获得干净的文件路径
        if is_valid_path_syntax(cleaned_directory) == False:
            exit()
        return cleaned_directory
    
    def is_valid_path_syntax(path):
        """验证 Windows 路径语法，包含以下检查：
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
    
    def _directory_total():
        '''自定义路径，默认路径为：
        'C:\\Users\\username\\AppData\\processdata'

        （其中username为用户名）
        '''
        directory_total = f'C:\\Users\\{username}\\AppData\\processdata'
        if os.path.isfile(f'{self_directory}directory'):
            with open(f'{self_directory}directory','r',encoding='utf-8') as f:
                file_content = f.read()
                if file_content == '': # 如果未提供路径，则返回默认路径
                    return directory_total
                else:
                    file_content = is_valid_path_syntax(file_content)
                    if file_content == False: # 如果提供的路径格式不正确，则返回False
                        os._exit(0)
                    return file_content
        else:
            #文件不存在时直接退出
            os._exit(0)
    
    def read_file_content(content):
        try:
            with open(f'{directory_total}\\{content}','r',encoding='utf-8') as f:
                content = f.read()
                content = content.split(',')
        except Exception:
            content = ''
        return content
    
    def main():
        banned_processes_content = read_file_content('banned_processes.txt')
        warned_processes_content = read_file_content('warmed_processes.txt')
        if '' in warned_processes_content:
            pass
        else:
            for s in banned_processes_content:
                if s in warned_processes_content:
                    print('进程名单中含有',s,'，软件将在5秒后自动关闭',sep='')
                    time.sleep(5)
                    os._exit(0)

        if banned_processes_content == '':
            os._exit(0)

        while True:
            try:
                for c in banned_processes_content:
                    time.sleep(0.3)
                    subprocess.Popen(['cmd','/c','taskkill /f /im',c])  # 终止进程
                    time.sleep(0.3)     
            except KeyboardInterrupt:
                pass

    if __name__ == '__main__':
        username = getpass.getuser()    # 获取当前用户的用户名
        self_directory = cleaned_directory()
        directory_total = _directory_total()
        main()
else:
    # 不是管理员，请求管理员权限
    if ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable,f'"{os.path.abspath(sys.argv[0])}"', None, 1) != 0:
        pass