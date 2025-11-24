import os
import shutil
from PIL import Image
import sys

# 支持的格式定义
SUPPORTED_FORMATS = {
    'png': ('PNG', '.png'),
    'webp': ('WEBP', '.webp'), 
    'jpg': ('JPEG', '.jpg'),
    'jpeg': ('JPEG', '.jpg'),
    'jpe': ('JPEG', '.jpg'),
    'tif': ('TIFF', '.tif'),
    'tiff': ('TIFF', '.tiff'),
    'bmp': ('BMP', '.bmp')
}

def get_storage_option():
    """获取用户选择的存储方式"""
    print("\n请选择存储方式:")
    print("1. 使用默认文件夹 (input_images 和 converted_images)")
    print("2. 自定义输入和输出路径")
    
    while True:
        choice = input("请选择 (1 或 2): ").strip()
        if choice == '1':
            return 'default'
        elif choice == '2':
            return 'custom'
        else:
            print("无效选择，请输入 1 或 2")

def create_default_folders():
    """创建默认的输入和输出文件夹"""
    input_folder = "input_images"
    output_folder = "converted_images"
    
    if not os.path.exists(input_folder):
        os.makedirs(input_folder)
        print(f"已创建输入文件夹: {input_folder}")
        print(f"请将要转换的图片放入 '{input_folder}' 文件夹中")
    else:
        print(f"输入文件夹已存在: {input_folder}")
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"已创建输出文件夹: {output_folder}")
    else:
        print(f"输出文件夹已存在: {output_folder}")
    
    return input_folder, output_folder

def select_folder_dialog(title):
    """使用tkinter弹出文件夹选择对话框"""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        print("错误: 无法导入tkinter，请确保已安装tkinter")
        return None
    
    # 创建隐藏的根窗口
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    root.attributes('-topmost', True)  # 确保对话框在最前面
    
    # 弹出文件夹选择对话框
    folder_path = filedialog.askdirectory(title=title)
    
    # 销毁根窗口
    root.destroy()
    
    return folder_path

def get_custom_paths():
    """获取用户自定义的输入和输出路径（使用图形界面）"""
    print("\n将弹出文件夹选择对话框...")
    
    # 选择输入文件夹
    print("请选择输入文件夹...")
    input_folder = select_folder_dialog("选择输入文件夹（包含待转换图片）")
    
    if not input_folder:
        print("未选择输入文件夹，使用默认路径")
        input_folder = "input_images"
        if not os.path.exists(input_folder):
            os.makedirs(input_folder)
    
    print(f"输入文件夹: {input_folder}")
    
    # 选择输出文件夹
    print("\n请选择输出文件夹...")
    output_folder = select_folder_dialog("选择输出文件夹（保存转换后的图片）")
    
    if not output_folder:
        print("未选择输出文件夹，使用默认路径")
        output_folder = "converted_images"
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
    
    print(f"输出文件夹: {output_folder}")
    
    # 检查输入和输出路径是否相同
    if os.path.abspath(input_folder) == os.path.abspath(output_folder):
        print("警告：输入和输出路径相同，这可能导致文件覆盖！")
        confirm = input("是否继续? (Y/N): ").strip().lower()
        if confirm not in ['y', 'yes']:
            return get_custom_paths()  # 重新选择
    
    return input_folder, output_folder

def setup_folders():
    """设置输入和输出文件夹"""
    storage_option = get_storage_option()
    
    if storage_option == 'default':
        return create_default_folders()
    else:
        return get_custom_paths()

def get_target_format():
    """获取用户选择的目标格式"""
    print("\n支持的输出格式: PNG, WEBP, JPG, JPEG, JPE, TIF, TIFF, BMP")
    
    while True:
        target_format = input("请输入目标格式 (例如: png, jpg, webp): ").lower().strip()
        
        if target_format in SUPPORTED_FORMATS:
            return target_format
        else:
            print("不支持的格式，请重新输入！")

def convert_image(input_path, output_path, target_format):
    """转换单张图片"""
    try:
        with Image.open(input_path) as img:
            pil_format, _ = SUPPORTED_FORMATS[target_format]
            
            # 处理RGB转换（对于不支持透明通道的格式）
            if target_format in ['jpg', 'jpeg', 'jpe', 'bmp'] and img.mode in ('RGBA', 'LA', 'P'):
                # 创建白色背景的RGB图像
                if img.mode == 'P' and 'transparency' in img.info:
                    img = img.convert('RGBA')
                rgb_img = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'RGBA':
                    rgb_img.paste(img, mask=img.split()[-1])
                else:
                    rgb_img.paste(img)
                img = rgb_img
            
            # 保存图片
            save_kwargs = {}
            if target_format == 'webp':
                save_kwargs['quality'] = 80  # 默认质量
            elif target_format in ['jpg', 'jpeg', 'jpe']:
                save_kwargs['quality'] = 95  # JPEG质量
            
            img.save(output_path, format=pil_format, **save_kwargs)
            return True, None
            
    except Exception as e:
        return False, str(e)

def process_images(input_folder, output_folder, target_format):
    """批量处理图片"""
    processed_count = 0
    skipped_count = 0
    error_count = 0
    
    print(f"\n开始处理图片...")
    print(f"输入文件夹: {input_folder}")
    print(f"输出文件夹: {output_folder}")
    print(f"目标格式: {target_format}")
    print("-" * 50)
    
    # 获取输入文件夹中的所有文件
    try:
        files = [f for f in os.listdir(input_folder) if os.path.isfile(os.path.join(input_folder, f))]
    except Exception as e:
        print(f"无法读取输入文件夹: {e}")
        return False, 0, 0, 0
    
    if not files:
        print("输入文件夹中没有找到任何文件！")
        return False, 0, 0, 0
    
    _, target_extension = SUPPORTED_FORMATS[target_format]
    
    for i, filename in enumerate(files, 1):
        input_path = os.path.join(input_folder, filename)
        
        # 获取文件扩展名（不含点）
        file_ext = os.path.splitext(filename)[1].lower().lstrip('.')
        
        # 检查是否为支持的图片格式
        if not file_ext or file_ext not in SUPPORTED_FORMATS:
            print(f"[{i}/{len(files)}] 跳过不支持的文件: {filename}")
            skipped_count += 1
            continue
        
        # 生成输出文件名
        name_without_ext = os.path.splitext(filename)[0]
        output_filename = f"{name_without_ext}{target_extension}"
        output_path = os.path.join(output_folder, output_filename)
        
        # 显示进度
        print(f"[{i}/{len(files)}] 处理: {filename}", end="")
        
        # 如果源格式与目标格式相同，直接复制
        if file_ext == target_format:
            try:
                shutil.copy2(input_path, output_path)
                print(f" -> 复制 (格式相同)")
                skipped_count += 1
            except Exception as e:
                print(f" -> 复制失败: {e}")
                error_count += 1
            continue
        
        # 转换图片
        success, error_msg = convert_image(input_path, output_path, target_format)
        
        if success:
            print(f" -> 转换成功")
            processed_count += 1
        else:
            print(f" -> 转换失败: {error_msg}")
            error_count += 1
    
    # 输出统计信息
    print("-" * 50)
    print(f"处理完成！")
    print(f"成功转换: {processed_count} 张")
    print(f"跳过处理: {skipped_count} 张")
    print(f"转换失败: {error_count} 张")
    print(f"输出文件夹: {output_folder}")
    
    return True, processed_count, skipped_count, error_count

def ask_continue():
    """询问用户是否继续转换"""
    while True:
        response = input("\n是否继续转换? (Y/N->是/否): ").strip().lower()
        if response in ['y', 'yes', '']:
            return True
        elif response in ['n', 'no']:
            return False
        else:
            print("请输入 Y(是) 或 N(否)")

def clear_input_folder(input_folder):
    """清空输入文件夹"""
    try:
        files_removed = 0
        for filename in os.listdir(input_folder):
            file_path = os.path.join(input_folder, filename)
            if os.path.isfile(file_path):
                os.remove(file_path)
                files_removed += 1
        print(f"已清空输入文件夹: {input_folder} (删除了 {files_removed} 个文件)")
        return True
    except Exception as e:
        print(f"清空输入文件夹时出错: {e}")
        return False

def change_folders():
    """询问用户是否要更改文件夹设置"""
    while True:
        change = input("\n是否要更改输入/输出文件夹设置? (Y/N): ").strip().lower()
        if change in ['y', 'yes']:
            return True
        elif change in ['n', 'no']:
            return False
        else:
            print("请输入 Y(是) 或 N(否)")

def main():
    """主函数"""
    print("=" * 60)
    print("           图片格式批量转换工具")
    print("=" * 60)
    
    try:
        # 设置文件夹
        input_folder, output_folder = setup_folders()
        
        # 主循环
        while True:
            # 检查输入文件夹是否有文件
            try:
                files = [f for f in os.listdir(input_folder) if os.path.isfile(os.path.join(input_folder, f))]
            except Exception as e:
                print(f"无法访问输入文件夹: {e}")
                if change_folders():
                    input_folder, output_folder = setup_folders()
                    continue
                else:
                    break
                    
            if not files:
                print(f"\n请在 '{input_folder}' 文件夹中放入要转换的图片")
                input("放置完成后，按 Enter 键继续...")
                
                # 再次检查
                files = [f for f in os.listdir(input_folder) if os.path.isfile(os.path.join(input_folder, f))]
                if not files:
                    print("仍然没有找到图片文件。")
                    if not ask_continue():
                        break
                    continue
            
            # 获取目标格式
            target_format = get_target_format()
            
            # 处理图片
            success, processed_count, skipped_count, error_count = process_images(
                input_folder, output_folder, target_format)
            
            # 询问是否继续
            if not ask_continue():
                break
                
            # 询问是否清空输入文件夹
            if files:  # 只有在有文件时才询问
                clear_input = input("\n是否清空输入文件夹以准备下一批图片? (Y/N): ").strip().lower()
                if clear_input in ['y', 'yes']:
                    clear_input_folder(input_folder)
            
            # 询问是否要更改文件夹设置
            if change_folders():
                input_folder, output_folder = setup_folders()
        
        print(f"\n感谢使用图片格式批量转换工具！")
        
    except KeyboardInterrupt:
        print("\n\n用户中断操作")
    except Exception as e:
        print(f"\n发生错误: {e}")
        import traceback
        traceback.print_exc()
    
    input("\n按 Enter 键退出...")

if __name__ == "__main__":
    main()