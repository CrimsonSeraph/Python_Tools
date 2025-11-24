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

def create_folders():
    """创建输入和输出文件夹"""
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
    files = [f for f in os.listdir(input_folder) if os.path.isfile(os.path.join(input_folder, f))]
    
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
            shutil.copy2(input_path, output_path)
            print(f" -> 复制 (格式相同)")
            skipped_count += 1
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
        for filename in os.listdir(input_folder):
            file_path = os.path.join(input_folder, filename)
            if os.path.isfile(file_path):
                os.remove(file_path)
        print(f"已清空输入文件夹: {input_folder}")
        return True
    except Exception as e:
        print(f"清空输入文件夹时出错: {e}")
        return False

def main():
    """主函数"""
    print("=" * 60)
    print("           图片格式批量转换工具")
    print("=" * 60)
    
    try:
        # 创建文件夹
        input_folder, output_folder = create_folders()
        
        # 主循环
        while True:
            # 检查输入文件夹是否有文件
            files = [f for f in os.listdir(input_folder) if os.path.isfile(os.path.join(input_folder, f))]
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
            clear_input = input("\n是否清空输入文件夹以准备下一批图片? (Y/N): ").strip().lower()
            if clear_input in ['y', 'yes']:
                clear_input_folder(input_folder)
        
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