import os
import sys
import subprocess
import configparser
import time
import msvcrt  # 用于捕获空格键
import win32api
import win32con
from win32com.shell.shell import SHChangeNotify
from win32com.shell import shellcon

# 全局变量：存储用户选择的操作目录
OPERATE_DIR = ""

# 配置参数
EXCLUDE_EXE_KEYWORD = "uninstall"
ICON_EXTENSION = ".ico"
DEFAULT_ICON_SIZE = (48, 48)

# 等待空格键确认
def wait_for_space():
    print("按空格键继续...", end='', flush=True)
    while True:
        key = msvcrt.getch()
        if key == b' ':
            print()  # 换行
            break

# 依赖检查
def check_dependency(module_name, install_name=None):
    install_name = install_name or module_name
    try:
        __import__(module_name)
        return True
    except ImportError:
        print(f"缺失依赖：{install_name}，正在安装...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", install_name])
            __import__(module_name)
            return True
        except:
            print(f"手动安装命令：pip install {install_name}")
            return False

# 目录选择函数
def select_operate_directory():
    print("===== 请选择操作目录 =====")
    while True:
        try:
            print("1. 使用当前目录（脚本所在目录）")
            print("2. 手动指定路径")
            choice = input("请选择(1-2)：").strip()
            
            if choice == '1':
                current_dir = os.getcwd()
                print(f"已选择当前目录：{current_dir}")
                return current_dir
            elif choice == '2':
                while True:
                    path = input("请输入目标目录路径（直接回车返回上一级）：").strip()
                    if not path:
                        break
                    if os.path.isdir(path):
                        print(f"已选择目录：{path}")
                        return path
                    else:
                        print(f"错误：路径不存在或不是有效目录")
            else:
                print("无效选择，请输入1或2")
        except Exception as e:
            print(f"目录选择出错：{str(e)}")
            wait_for_space()

# 必须依赖检查
required_deps = [
    ("win32gui", "pywin32"),
    ("PIL", "pillow"),
    ("win32com.shell", "pywin32")
]
for module, install_name in required_deps:
    if not check_dependency(module, install_name):
        print("依赖安装失败，按空格键退出...", end='', flush=True)
        while True:
            if msvcrt.getch() == b' ':
                sys.exit(1)

# 导入其他所需模块
import win32gui
import win32ui
from win32con import DI_NORMAL
from PIL import Image

# ------------------------------
# 核心图标提取方法（简化版，优先保证图标正确）
# ------------------------------
def get_exe_icon(exe_path, output_path):
    try:
        # 提取图标（使用最基础的方法，确保能获取到图标）
        large_icons, _ = win32gui.ExtractIconEx(exe_path, 0)
        if not large_icons:
            return False, "无大图标资源"
        
        hicon = large_icons[0]
        hdc_screen = None
        hdc = None
        hdc_mem = None
        hbmp = None
        
        try:
            # 基础设备上下文设置（避免复杂参数）
            hdc_screen = win32gui.GetDC(0)
            hdc = win32ui.CreateDCFromHandle(hdc_screen)
            hbmp = win32ui.CreateBitmap()
            hbmp.CreateCompatibleBitmap(hdc, DEFAULT_ICON_SIZE[0], DEFAULT_ICON_SIZE[1])
            hdc_mem = hdc.CreateCompatibleDC()
            hdc_mem.SelectObject(hbmp)

            # 基础绘制（不使用复杂渲染参数，确保图标能被正确绘制）
            win32gui.DrawIconEx(
                hdc_mem.GetSafeHdc(), 0, 0, hicon,
                DEFAULT_ICON_SIZE[0], DEFAULT_ICON_SIZE[1],
                0, None, DI_NORMAL  # 仅使用基础参数
            )

            # 提取位图数据（使用原始格式，避免模式转换错误）
            bmp_info = hbmp.GetInfo()
            bmp_str = hbmp.GetBitmapBits(True)
            
            # 基础转换（使用BGRX模式，这是之前能正常工作的模式）
            img = Image.frombuffer(
                'RGB', 
                (bmp_info['bmWidth'], bmp_info['bmHeight']),
                bmp_str, 
                'raw', 
                'BGRX', 0, 1  # 还原为之前能正确提取图标的模式
            ).convert('RGBA')  # 转为RGBA以添加透明通道

            # 简单透明处理：保留图像本身的透明信息（如果有）
            # 不做复杂合成，避免破坏原始图标
            img.save(
                output_path, 
                format='ICO', 
                sizes=[DEFAULT_ICON_SIZE, (32,32)]
            )
            return True, "提取成功（基础透明处理）"
        
        finally:
            # 清理资源
            if hicon:
                win32gui.DestroyIcon(hicon)
            if hdc_mem:
                hdc_mem.DeleteDC()
            if hbmp:
                win32gui.DeleteObject(hbmp.GetHandle())
            if hdc_screen:
                win32gui.ReleaseDC(0, hdc_screen)
    except Exception as e:
        return False, f"提取失败：{str(e)}"

# ------------------------------
# 其他函数保持不变（空格确认、刷新等）
# ------------------------------
def refresh_folder(folder_path):
    try:
        if not isinstance(folder_path, str):
            folder_path = str(folder_path)
        
        folder_path_bytes = os.fsencode(folder_path)
        
        win32api.SetFileAttributes(folder_path, win32con.FILE_ATTRIBUTE_SYSTEM)
        SHChangeNotify(
            shellcon.SHCNE_UPDATEDIR,
            shellcon.SHCNF_PATH | shellcon.SHCNF_FLUSH,
            folder_path_bytes,
            None
        )
        time.sleep(0.1)
        return True
    except Exception as e:
        print(f"刷新文件夹失败：{str(e)}")
        return False

def ensure_file_writable(file_path):
    try:
        if os.path.exists(file_path):
            attrs = win32api.GetFileAttributes(file_path)
            new_attrs = attrs & ~(
                win32con.FILE_ATTRIBUTE_SYSTEM | 
                win32con.FILE_ATTRIBUTE_HIDDEN | 
                win32con.FILE_ATTRIBUTE_READONLY
            )
            win32api.SetFileAttributes(file_path, new_attrs)
        return True
    except Exception as e:
        print(f"警告：无法修改文件属性 {file_path} - {str(e)}")
        return False

def generate_folders_txt():
    try:
        print("\n----- 开始生成folders.txt -----")
        current_dir = OPERATE_DIR
        if not current_dir or not os.path.isdir(current_dir):
            print("错误：操作目录无效")
            return
        
        folders = [f for f in os.listdir(current_dir) if os.path.isdir(os.path.join(current_dir, f)) and not f.startswith('.')]
        if not folders:
            print("无可用文件夹")
            return

        folders_txt_path = os.path.join(current_dir, 'folders.txt')
        with open(folders_txt_path, 'w', encoding='utf-8') as f:
            f.write("# [文件夹名]\n# LocalizedResourceName=显示名\n# IconResource=EXE相对路径\n\n")
            for folder in folders:
                folder_path = os.path.join(current_dir, folder)
                exes = []
                for root, _, files in os.walk(folder_path):
                    for file in files:
                        if file.lower().endswith('.exe') and EXCLUDE_EXE_KEYWORD not in file.lower():
                            rel_exe_path = os.path.relpath(os.path.join(root, file), folder_path)
                            exes.append(rel_exe_path)
                exe = exes[0] if exes else ''
                f.write(f"[{folder}]\nLocalizedResourceName={folder}\nIconResource={exe}\n\n")
        print(f"成功生成folders.txt到：{folders_txt_path}")
    except Exception as e:
        print(f"生成folders.txt失败：{str(e)}")
    finally:
        wait_for_space()

def batch_extract_icons():
    try:
        print("\n----- 开始批量提取图标 -----")
        current_dir = OPERATE_DIR
        folders_txt_path = os.path.join(current_dir, 'folders.txt')
        if not os.path.exists(folders_txt_path):
            print("错误：未找到folders.txt，请先执行步骤1")
            return

        config = configparser.ConfigParser()
        config.optionxform = str
        config.read(folders_txt_path, encoding='utf-8')
        success = 0
        total = len(config.sections())

        if total == 0:
            print("错误：folders.txt中没有配置任何文件夹")
            return

        for folder in config.sections():
            folder_path = os.path.join(current_dir, folder)
            exe_rel_path = config.get(folder, 'IconResource', fallback='')
            if not exe_rel_path or not exe_rel_path.lower().endswith('.exe'):
                print(f"[{folder}] 无有效EXE，跳过")
                continue

            exe_abs_path = os.path.join(folder_path, exe_rel_path)
            if not os.path.exists(exe_abs_path):
                print(f"[{folder}] EXE不存在：{exe_rel_path}")
                continue

            exe_basename = os.path.basename(exe_rel_path)
            ico_abs_path = os.path.join(folder_path, f"{os.path.splitext(exe_basename)[0]}{ICON_EXTENSION}")
            
            print(f"\n处理：{folder} → {exe_rel_path}")
            ok, msg = get_exe_icon(exe_abs_path, ico_abs_path)
            if ok:
                success += 1
                print(f"✅ 图标生成到：{os.path.relpath(ico_abs_path, folder_path)}")
            else:
                # 生成默认图标（仅在提取失败时）
                from PIL import ImageDraw, ImageFont
                img = Image.new('RGBA', DEFAULT_ICON_SIZE, (240,240,240,255))
                draw = ImageDraw.Draw(img)
                try:
                    font = ImageFont.truetype("simsun", 14)
                except:
                    font = ImageFont.load_default()
                draw.text((5,15), folder[:4], font=font, fill=(0,0,0,255))
                img.save(ico_abs_path, format='ICO', sizes=[DEFAULT_ICON_SIZE, (32,32)])
                print(f"⚠️ {msg}，已生成默认图标")

        print(f"\n批量提取完成：成功{success}/{total}")
    except Exception as e:
        print(f"批量提取图标失败：{str(e)}")
    finally:
        wait_for_space()

def generate_desktop_ini():
    try:
        print("\n----- 开始生成desktop.ini -----")
        current_dir = OPERATE_DIR
        folders_txt_path = os.path.join(current_dir, 'folders.txt')
        if not os.path.exists(folders_txt_path):
            print("错误：未找到folders.txt，请先执行步骤1")
            return

        config = configparser.ConfigParser()
        config.optionxform = str
        config.read(folders_txt_path, encoding='utf-8')
        processed = 0
        updated_folders = set()

        for folder in config.sections():
            folder_abs_path = os.path.join(current_dir, folder)
            ini_abs_path = os.path.join(folder_abs_path, "desktop.ini")
            
            exe_rel_path = config.get(folder, 'IconResource', fallback='')
            display_name = config.get(folder, 'LocalizedResourceName', fallback=folder)
            if not exe_rel_path.endswith('.exe'):
                continue

            if not os.access(folder_abs_path, os.W_OK):
                print(f"[{folder}] 文件夹无写入权限，跳过")
                continue

            if not ensure_file_writable(ini_abs_path):
                continue

            exe_basename = os.path.basename(exe_rel_path)
            ico_filename = f"{os.path.splitext(exe_basename)[0]}{ICON_EXTENSION}"
            ico_abs_path = os.path.join(folder_abs_path, ico_filename)
            
            if not os.path.exists(ico_abs_path):
                print(f"[{folder}] 图标文件不存在，跳过")
                continue
            
            try:
                with open(ini_abs_path, 'w', encoding='ANSI', newline='\r\n') as f:
                    f.write("[.ShellClassInfo]\r\n")
                    f.write(f"LocalizedResourceName={display_name}\r\n")
                    f.write(f"IconFile={ico_filename}\r\n")
                    f.write("IconIndex=0\r\n")
                print(f"[{folder}] 已写入desktop.ini")
            except Exception as e:
                print(f"[{folder}] 写入desktop.ini失败：{str(e)}")
                continue

            try:
                win32api.SetFileAttributes(
                    ini_abs_path, 
                    win32con.FILE_ATTRIBUTE_HIDDEN | win32con.FILE_ATTRIBUTE_SYSTEM
                )
                print(f"[{folder}] 已设置文件属性")
                updated_folders.add(folder_abs_path)
                processed += 1
            except Exception as e:
                print(f"[{folder}] 设置属性失败：{str(e)}")

        if updated_folders:
            print("\n开始批量刷新文件夹视图...")
            for folder in updated_folders:
                refresh_folder(folder)
            print(f"已刷新 {len(updated_folders)} 个文件夹")
        
        print(f"\ndesktop.ini生成完成，成功处理 {processed} 个文件夹")
    except Exception as e:
        print(f"生成desktop.ini失败：{str(e)}")
    finally:
        wait_for_space()

def clean_desktop_ini():
    try:
        print("\n----- 开始清理desktop.ini -----")
        current_dir = OPERATE_DIR
        deleted = 0
        affected_roots = set()

        for root, _, files in os.walk(current_dir):
            for file in files:
                if file == 'desktop.ini' or file.startswith('desktop.ini.bak'):
                    file_path = os.path.join(root, file)
                    try:
                        if ensure_file_writable(file_path):
                            os.remove(file_path)
                            deleted += 1
                            affected_roots.add(root)
                            print(f"已删除：{os.path.relpath(file_path, current_dir)}")
                    except Exception as e:
                        print(f"删除失败 {os.path.relpath(file_path, current_dir)}：{str(e)}")
        
        if affected_roots:
            print("\n开始批量刷新受影响的目录...")
            for root in affected_roots:
                refresh_folder(root)
            print(f"已刷新 {len(affected_roots)} 个目录")
        
        print(f"\n清理完成，共删除 {deleted} 个文件")
    except Exception as e:
        print(f"清理desktop.ini失败：{str(e)}")
    finally:
        wait_for_space()

def clean_extracted_icons():
    try:
        print("\n----- 开始清理ICO图标 -----")
        current_dir = OPERATE_DIR
        folders_txt_path = os.path.join(current_dir, 'folders.txt')
        if not os.path.exists(folders_txt_path):
            print("错误：未找到folders.txt，请先执行步骤1")
            return

        config = configparser.ConfigParser()
        config.optionxform = str
        config.read(folders_txt_path, encoding='utf-8')
        deleted = 0
        affected_folders = set()

        for folder in config.sections():
            folder_path = os.path.join(current_dir, folder)
            exe_rel_path = config.get(folder, 'IconResource', fallback='')
            if not exe_rel_path or not exe_rel_path.lower().endswith('.exe'):
                continue

            exe_basename = os.path.basename(exe_rel_path)
            ico_name = f"{os.path.splitext(exe_basename)[0]}{ICON_EXTENSION}"
            ico_abs_path = os.path.join(folder_path, ico_name)

            if os.path.exists(ico_abs_path):
                try:
                    os.remove(ico_abs_path)
                    deleted += 1
                    affected_folders.add(folder_path)
                    print(f"已删除：{os.path.relpath(ico_abs_path, current_dir)}")
                except Exception as e:
                    print(f"删除失败 {os.path.relpath(ico_abs_path, current_dir)}：{str(e)}")
        
        if affected_folders:
            print("\n开始批量刷新受影响的文件夹...")
            for folder in affected_folders:
                refresh_folder(folder)
            print(f"已刷新 {len(affected_folders)} 个文件夹")
        
        print(f"\n清理完成，共删除 {deleted} 个ICO文件")
    except Exception as e:
        print(f"清理ICO图标失败：{str(e)}")
    finally:
        wait_for_space()

# ------------------------------
# 主菜单
# ------------------------------
def main():
    global OPERATE_DIR
    try:
        OPERATE_DIR = select_operate_directory()
        if not OPERATE_DIR or not os.path.isdir(OPERATE_DIR):
            print("未选择有效目录，程序退出")
            return

        while True:
            print(f"\n===== 文件夹图标工具（当前操作目录：{OPERATE_DIR}）=====")
            print("1. 生成folders.txt")
            print("2. 手动编辑folders.txt")
            print("3. 批量提取图标")
            print("4. 生成desktop.ini")
            print("5. 强制刷新视图")
            print("6. 清理desktop.ini")
            print("7. 清理ICO图标")
            print("8. 退出")
            
            choice = input("\n请选择操作(1-8)：").strip()
            
            if choice == '1':
                generate_folders_txt()
            elif choice == '2':
                try:
                    print("\n----- 手动编辑folders.txt -----")
                    folders_txt_path = os.path.join(OPERATE_DIR, 'folders.txt')
                    print(f"编辑路径：{folders_txt_path}")
                    if os.path.exists(folders_txt_path):
                        os.startfile(folders_txt_path)
                    else:
                        print("提示：folders.txt不存在，请先执行步骤1生成")
                except Exception as e:
                    print(f"打开文件失败：{str(e)}")
                finally:
                    wait_for_space()
            elif choice == '3':
                batch_extract_icons()
            elif choice == '4':
                generate_desktop_ini()
            elif choice == '5':
                try:
                    print("\n----- 开始强制刷新视图 -----")
                    current_dir = OPERATE_DIR
                    count = 0
                    for folder in os.listdir(current_dir):
                        folder_path = os.path.join(current_dir, folder)
                        if os.path.isdir(folder_path):
                            if refresh_folder(folder_path):
                                count += 1
                    print(f"已刷新 {count} 个文件夹")
                except Exception as e:
                    print(f"强制刷新失败：{str(e)}")
                finally:
                    wait_for_space()
            elif choice == '6':
                clean_desktop_ini()
            elif choice == '7':
                clean_extracted_icons()
            elif choice == '8':
                print("程序退出")
                break
            else:
                print("无效选择，请输入1-8之间的数字")
                wait_for_space()
    except Exception as e:
        print(f"程序主逻辑出错：{str(e)}")
        wait_for_space()

if __name__ == "__main__":
    if os.name != 'nt':
        print("仅支持Windows系统")
        print("按空格键退出...", end='', flush=True)
        while True:
            if msvcrt.getch() == b' ':
                sys.exit(1)
    main()
