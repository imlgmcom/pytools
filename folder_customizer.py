import os
import configparser
import ctypes
import traceback
import shutil
from datetime import datetime

def refresh_folder_view():
    """刷新资源管理器视图"""
    try:
        print("正在刷新文件夹视图...")
        
        # 使用系统API刷新视图
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0x0000, None, None)
        
        print("文件夹视图已刷新，更改应该立即可见")
        return True
    except Exception as e:
        print(f"刷新文件夹视图时出错: {str(e)}")
        return False

def get_desktop_ini_info(folder_path):
    """读取文件夹中已有的desktop.ini信息"""
    desktop_ini_path = os.path.join(folder_path, 'desktop.ini')
    name = ""
    icon = ""
    
    if os.path.exists(desktop_ini_path):
        try:
            with open(desktop_ini_path, 'r', encoding='ANSI') as f:
                content = f.read()
            
            if '[.ShellClassInfo]' in content:
                lines = content.split('\n')
                for line in lines:
                    line = line.strip()
                    if line.startswith('LocalizedResourceName='):
                        name = line.split('=', 1)[1].strip()
                    elif line.startswith('IconResource='):
                        icon = line.split('=', 1)[1].strip()
            return name, icon
        except Exception as e:
            print(f"读取 {os.path.basename(folder_path)} 的desktop.ini时出错: {str(e)}")
    
    return name, icon

def generate_folders_txt():
    """生成初始的folders.txt文件"""
    try:
        current_dir = os.getcwd()
        folders = [f for f in os.listdir(current_dir) if os.path.isdir(os.path.join(current_dir, f)) and not f.startswith('.')]
        
        if not folders:
            print("当前目录下没有找到文件夹")
            return
        
        with open('folders.txt', 'w', encoding='utf-8') as f:
            f.write(f"# 文件夹配置文件 - 生成于 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("# 格式：\n")
            f.write("# [文件夹实际名称]\n")
            f.write("# LocalizedResourceName=显示名称\n")
            f.write("# IconResource=图标路径（可相对或绝对路径）\n\n")
            
            for folder in folders:
                folder_path = os.path.join(current_dir, folder)
                name, icon = get_desktop_ini_info(folder_path)
                
                if not name:
                    name = folder
                
                f.write(f'[{folder}]\n')
                f.write(f'LocalizedResourceName={name}\n')
                f.write(f'IconResource={icon}\n\n')
        
        print(f"成功生成folders.txt文件，包含{len(folders)}个文件夹配置")
    except Exception as e:
        print(f"生成folders.txt时出错: {str(e)}")
        traceback.print_exc()

def generate_desktop_ini(skip_existing=True, auto_refresh=True):
    """根据folders.txt生成desktop.ini文件（存在则跳过）"""
    try:
        if not os.path.exists('folders.txt'):
            print("错误：未找到folders.txt文件，请先生成该文件并编辑")
            return
        
        config = configparser.ConfigParser()
        config.optionxform = str  # 保持键名大小写
        read_ok = config.read('folders.txt', encoding='utf-8')
        if not read_ok:
            print("错误：无法读取folders.txt文件，请检查文件是否存在且格式正确")
            return
        
        sections = config.sections()
        if not sections:
            print("错误：folders.txt文件中未找到任何文件夹配置")
            return
        
        print(f"开始生成desktop.ini文件，共发现{len(sections)}个文件夹配置")
        current_dir = os.getcwd()
        processed = 0
        skipped = 0
        
        for folder in sections:
            folder_path = os.path.join(current_dir, folder)
            
            if not os.path.isdir(folder_path):
                print(f"警告：文件夹 '{folder}' 不存在，已跳过")
                continue
            
            desktop_ini_path = os.path.join(folder_path, 'desktop.ini')
            exists = os.path.exists(desktop_ini_path)
            
            # 如果存在且需要跳过，则不处理
            if exists and skip_existing:
                skipped += 1
                continue
            
            # 获取配置值
            name = config.get(folder, 'LocalizedResourceName', fallback='').strip()
            icon = config.get(folder, 'IconResource', fallback='').strip()
            
            # 构建配置内容
            content = "[.ShellClassInfo]\r\n"
            if name:
                content += f"LocalizedResourceName={name}\r\n"
            if icon:
                content += f"IconResource={icon}\r\n"
            
            # 如已存在，先创建备份
            if exists:
                backup_path = f"{desktop_ini_path}.bak.{datetime.now().strftime('%Y%m%d%H%M%S')}"
                try:
                    shutil.copy2(desktop_ini_path, backup_path)
                    print(f"  已为 '{folder}' 创建备份: {os.path.basename(backup_path)}")
                except Exception as e:
                    print(f"  为 '{folder}' 创建备份失败: {str(e)}")
            
            # 写入配置文件
            try:
                with open(desktop_ini_path, 'w', encoding='ANSI', newline='\r\n') as f:
                    f.write(content)
                
                # 设置文件属性
                if os.name == 'nt':
                    FILE_ATTRIBUTE_HIDDEN = 0x02
                    FILE_ATTRIBUTE_SYSTEM = 0x04
                    ctypes.windll.kernel32.SetFileAttributesW(
                        desktop_ini_path, 
                        FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM
                    )
                    # 设置文件夹为系统属性
                    current_attr = ctypes.windll.kernel32.GetFileAttributesW(folder_path)
                    if current_attr != 0xFFFFFFFF:
                        new_attr = current_attr | FILE_ATTRIBUTE_SYSTEM
                        ctypes.windll.kernel32.SetFileAttributesW(folder_path, new_attr)
                
                processed += 1
                print(f"  已处理: {folder}")
            except Exception as e:
                print(f"  处理 '{folder}' 时出错: {str(e)}")
                traceback.print_exc()
        
        print(f"\n处理完成：成功生成{processed}个，跳过{skipped}个（已存在）")
        
        # 自动刷新视图
        if auto_refresh and (processed > 0):
            refresh_folder_view()
        else:
            print("提示：可使用选项4手动刷新文件夹视图")
        
    except Exception as e:
        print(f"执行过程中发生错误: {str(e)}")
        traceback.print_exc()

def clean_desktop_ini():
    """清理所有desktop.ini及相关备份文件"""
    try:
        current_dir = os.getcwd()
        folders = [f for f in os.listdir(current_dir) if os.path.isdir(os.path.join(current_dir, f)) and not f.startswith('.')]
        
        if not folders:
            print("当前目录下没有找到文件夹")
            return
        
        print(f"开始清理desktop.ini文件，共发现{len(folders)}个文件夹")
        removed_files = 0
        
        for folder in folders:
            folder_path = os.path.join(current_dir, folder)
            
            # 清理desktop.ini
            desktop_ini_path = os.path.join(folder_path, 'desktop.ini')
            if os.path.exists(desktop_ini_path):
                try:
                    # 先移除系统和隐藏属性
                    if os.name == 'nt':
                        ctypes.windll.kernel32.SetFileAttributesW(
                            desktop_ini_path, 
                            0x80  # 正常文件属性
                        )
                    os.remove(desktop_ini_path)
                    removed_files += 1
                    print(f"  已删除: {folder}/desktop.ini")
                except Exception as e:
                    print(f"  删除 '{folder}/desktop.ini' 失败: {str(e)}")
            
            # 清理所有备份文件
            for file in os.listdir(folder_path):
                if file.startswith('desktop.ini.bak'):
                    file_path = os.path.join(folder_path, file)
                    try:
                        os.remove(file_path)
                        removed_files += 1
                        print(f"  已删除: {folder}/{file}")
                    except Exception as e:
                        print(f"  删除 '{folder}/{file}' 失败: {str(e)}")
            
            # 移除文件夹的系统属性
            if os.name == 'nt':
                try:
                    current_attr = ctypes.windll.kernel32.GetFileAttributesW(folder_path)
                    if current_attr != 0xFFFFFFFF:
                        # 清除系统属性但保留其他属性
                        new_attr = current_attr & ~0x04  # 0x04是系统属性
                        ctypes.windll.kernel32.SetFileAttributesW(folder_path, new_attr)
                except Exception as e:
                    print(f"  重置 '{folder}' 属性失败: {str(e)}")
        
        print(f"\n清理完成，共删除{removed_files}个文件")
        refresh_folder_view()  # 清理后自动刷新
        
    except Exception as e:
        print(f"清理过程中发生错误: {str(e)}")
        traceback.print_exc()

def main():
    while True:  # 使用循环保持程序运行，直到用户选择退出
        print("\n===== 文件夹批量自定义工具 =====")
        print("1. 生成/更新folders.txt文件")
        print("2. 生成desktop.ini (存在则跳过)")
        print("3. 清理所有desktop.ini及备份文件")
        print("4. 刷新文件夹视图（立即显示更改）")
        print("5. 退出")
        
        try:
            choice = input("请选择操作 (1-5): ")
            
            if choice == '1':
                generate_folders_txt()
            elif choice == '2':
                generate_desktop_ini(skip_existing=True, auto_refresh=True)
            elif choice == '3':
                confirm = input("确定要删除所有desktop.ini及备份文件吗？(y/n): ")
                if confirm.lower() == 'y':
                    clean_desktop_ini()
                else:
                    print("已取消清理操作")
            elif choice == '4':
                refresh_folder_view()
            elif choice == '5':
                print("程序已退出")
                break  # 退出循环，结束程序
            else:
                print("无效的选择，请重试")
                
            # 等待用户确认后返回主菜单
            input("\n按回车键返回主菜单...")
        except Exception as e:
            print(f"程序执行出错: {str(e)}")
            traceback.print_exc()
            input("按回车键返回主菜单...")

if __name__ == "__main__":
    if os.name != 'nt':
        print("错误：此工具仅适用于Windows系统")
        input("按回车键退出...")
    else:
        main()
    