import os
import sys
import zipfile
import rarfile
import py7zr
import tempfile
import shutil
from pathlib import Path
from typing import List, Dict, Optional, Tuple
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import logging

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class BatchExtractor:
    def __init__(self):
        self.supported_formats = {'.zip', '.rar', '.7z', '.001', '.z01'}
        self.multi_part_extensions = {'.001', '.z01', '.r00', '.7z.001'}
        self.password_cache = {}
        self.root = None
        self.program_dir = Path(__file__).parent
        
    def initialize_ui(self):
        """初始化UI（仅在需要时）"""
        if not self.root:
            self.root = tk.Tk()
            self.root.withdraw()
    
    def close_ui(self):
        """关闭UI"""
        if self.root:
            self.root.destroy()
            self.root = None
    
    def console_input(self, prompt: str, valid_options: List[str] = None) -> str:
        """控制台输入处理"""
        while True:
            user_input = input(prompt).strip().lower()
            if not valid_options or user_input in valid_options:
                return user_input
            else:
                print(f"无效输入，请选择: {', '.join(valid_options)}")
    
    def create_temp_folder(self, folder_type: str) -> Path:
        """在程序目录创建临时文件夹"""
        temp_dir = self.program_dir / f"batch_extract_{folder_type}_temp"
        temp_dir.mkdir(exist_ok=True)
        return temp_dir
    
    def select_input_folder(self) -> Optional[Path]:
        """选择输入文件夹"""
        print("\n" + "="*50)
        print("选择输入文件夹")
        print("="*50)
        
        while True:
            choice = self.console_input(
                "请选择输入文件夹方式:\n"
                "  Y - 选择现有文件夹\n"
                "  N - 在程序目录创建新文件夹\n"
                "  quit - 退出程序\n"
                "请输入选择 (Y/N/quit): ",
                ['y', 'n', 'quit']
            )
            
            if choice == 'quit':
                return None
            elif choice == 'y':
                self.initialize_ui()
                folder_path = filedialog.askdirectory(
                    title="选择包含压缩包的文件夹",
                    parent=self.root
                )
                if folder_path:
                    return Path(folder_path)
                else:
                    print("未选择文件夹，请重试")
            else:  # choice == 'n'
                temp_dir = self.create_temp_folder("input")
                print(f"已在程序目录创建输入文件夹: {temp_dir}")
                print("请将压缩包放入此文件夹，然后按回车键继续...")
                input()
                return temp_dir
    
    def select_output_folder(self) -> Optional[Path]:
        """选择输出文件夹"""
        print("\n" + "="*50)
        print("选择输出文件夹")
        print("="*50)
        
        while True:
            choice = self.console_input(
                "请选择输出文件夹方式:\n"
                "  Y - 选择现有文件夹\n"
                "  N - 在程序目录创建新文件夹\n"
                "  quit - 退出程序\n"
                "请输入选择 (Y/N/quit): ",
                ['y', 'n', 'quit']
            )
            
            if choice == 'quit':
                return None
            elif choice == 'y':
                self.initialize_ui()
                folder_path = filedialog.askdirectory(
                    title="选择解压输出文件夹",
                    parent=self.root
                )
                if folder_path:
                    return Path(folder_path)
                else:
                    print("未选择文件夹，请重试")
            else:  # choice == 'n'
                temp_dir = self.create_temp_folder("output")
                print(f"已在程序目录创建输出文件夹: {temp_dir}")
                return temp_dir
    
    def find_archive_files(self, input_folder: Path) -> List[Path]:
        """查找所有支持的压缩文件"""
        archive_files = []
        for file_path in input_folder.iterdir():
            if file_path.is_file():
                suffix = file_path.suffix.lower()
                if suffix in self.supported_formats:
                    archive_files.append(file_path)
        return sorted(archive_files)
    
    def detect_multi_part_archives(self, archive_files: List[Path]) -> Dict[str, List[Path]]:
        """检测并分组分段压缩包"""
        multi_part_groups = {}
        single_files = []
        
        for file_path in archive_files:
            stem = file_path.stem.lower()
            suffix = file_path.suffix.lower()
            
            # 检查是否是分段文件
            is_multi_part = False
            for ext in self.multi_part_extensions:
                if str(file_path).lower().endswith(ext):
                    is_multi_part = True
                    base_name = stem
                    # 处理类似 .7z.001 的情况
                    if suffix == '.001' and '.' in stem:
                        base_name = stem.rsplit('.', 1)[0]
                    
                    if base_name not in multi_part_groups:
                        multi_part_groups[base_name] = []
                    multi_part_groups[base_name].append(file_path)
                    break
            
            if not is_multi_part:
                single_files.append(file_path)
        
        # 对分段文件组进行排序
        for base_name in multi_part_groups:
            multi_part_groups[base_name].sort()
        
        return multi_part_groups, single_files
    
    def verify_archive(self, archive_path: Path, password: Optional[str] = None) -> Tuple[bool, str]:
        """验证压缩包是否完整"""
        suffix = archive_path.suffix.lower()
        
        try:
            if suffix == '.zip':
                with zipfile.ZipFile(archive_path, 'r') as zipf:
                    if password:
                        zipf.setpassword(password.encode('utf-8'))
                    # 测试读取第一个文件（不实际解压）
                    test_file = zipf.infolist()[0] if zipf.infolist() else None
                    if test_file:
                        zipf.read(test_file.filename)
                return True, "正常"
                
            elif suffix == '.rar':
                try:
                    with rarfile.RarFile(archive_path, 'r') as rarf:
                        if password:
                            rarf.setpassword(password)
                        # 测试读取第一个文件
                        test_file = rarf.infolist()[0] if rarf.infolist() else None
                        if test_file:
                            rarf.read(test_file)
                    return True, "正常"
                except rarfile.NeedFirstVolume:
                    return True, "需要其他分卷"
                except rarfile.BadRarFile as e:
                    return False, f"损坏的RAR文件: {str(e)}"
                
            elif suffix in ['.7z', '.001']:
                try:
                    with py7zr.SevenZipFile(archive_path, 'r', password=password) as szf:
                        # 测试读取文件列表
                        files = szf.getnames()
                    return True, "正常"
                except Exception as e:
                    return False, f"7z文件错误: {str(e)}"
                    
            else:
                return False, f"不支持的文件格式: {suffix}"
                
        except Exception as e:
            error_msg = str(e)
            if "password" in error_msg.lower() or "encrypted" in error_msg.lower():
                return True, "需要密码"
            else:
                return False, f"验证失败: {error_msg}"
    
    def verify_all_archives(self, multi_part_groups: Dict[str, List[Path]], 
                          single_files: List[Path]) -> Tuple[List[Path], List[Tuple[Path, str]]]:
        """验证所有压缩包"""
        valid_files = []
        problematic_files = []
        
        print("\n开始验证压缩包完整性...")
        
        # 验证单个文件
        for file_path in single_files:
            print(f"验证文件: {file_path.name}", end="")
            is_valid, message = self.verify_archive(file_path)
            if is_valid:
                valid_files.append(file_path)
                print(f" - ✓ {message}")
            else:
                problematic_files.append((file_path, message))
                print(f" - ✗ {message}")
        
        # 验证分段压缩包
        for base_name, file_group in multi_part_groups.items():
            print(f"验证分段压缩包: {base_name} (共{len(file_group)}个文件)", end="")
            main_file = file_group[0]
            is_valid, message = self.verify_archive(main_file)
            if is_valid:
                valid_files.extend(file_group)
                print(f" - ✓ {message}")
            else:
                for file_path in file_group:
                    problematic_files.append((file_path, f"分段压缩包错误: {message}"))
                print(f" - ✗ {message}")
        
        return valid_files, problematic_files
    
    def get_password(self, archive_path: Path) -> Optional[str]:
        """获取密码输入"""
        archive_name = archive_path.name
        
        if archive_path in self.password_cache:
            return self.password_cache[archive_path]
        
        while True:
            print(f"\n文件 {archive_name} 需要密码")
            choice = self.console_input(
                "请选择:\n"
                "  Y - 输入密码\n"
                "  N - 跳过此文件\n"
                "  quit - 终止当前解压任务\n"
                "请输入选择 (Y/N/quit): ",
                ['y', 'n', 'quit']
            )
            
            if choice == 'quit':
                return None
            elif choice == 'n':
                return "SKIP"
            else:  # choice == 'y'
                password = input("请输入密码: ")
                if password:
                    self.password_cache[archive_path] = password
                    return password
                else:
                    print("密码不能为空！")
    
    def extract_zip(self, archive_path: Path, output_path: Path, password: Optional[str] = None) -> bool:
        """解压ZIP文件"""
        try:
            with zipfile.ZipFile(archive_path, 'r') as zipf:
                if password:
                    zipf.setpassword(password.encode('utf-8'))
                zipf.extractall(output_path)
            return True
        except Exception as e:
            logger.error(f"解压ZIP失败 {archive_path}: {str(e)}")
            return False
    
    def extract_rar(self, archive_path: Path, output_path: Path, password: Optional[str] = None) -> bool:
        """解压RAR文件"""
        try:
            with rarfile.RarFile(archive_path, 'r') as rarf:
                if password:
                    rarf.setpassword(password)
                rarf.extractall(output_path)
            return True
        except Exception as e:
            logger.error(f"解压RAR失败 {archive_path}: {str(e)}")
            return False
    
    def extract_7z(self, archive_path: Path, output_path: Path, password: Optional[str] = None) -> bool:
        """解压7z文件"""
        try:
            with py7zr.SevenZipFile(archive_path, 'r', password=password) as szf:
                szf.extractall(output_path)
            return True
        except Exception as e:
            logger.error(f"解压7z失败 {archive_path}: {str(e)}")
            return False
    
    def extract_archive(self, archive_path: Path, output_folder: Path) -> bool:
        """解压单个压缩包"""
        suffix = archive_path.suffix.lower()
        archive_name = archive_path.stem
        
        # 为分段压缩包创建统一的输出目录
        if suffix in self.multi_part_extensions:
            # 获取基础名称（去除分段编号）
            base_name = archive_name
            if '.' in archive_name and suffix == '.001':
                base_name = archive_name.rsplit('.', 1)[0]
            output_path = output_folder / base_name
        else:
            output_path = output_folder / archive_name
        
        output_path.mkdir(parents=True, exist_ok=True)
        
        # 尝试无密码解压
        success = False
        password_attempts = 0
        max_attempts = 3
        
        while not success and password_attempts < max_attempts:
            password = None
            if password_attempts == 0:
                # 第一次尝试无密码
                print(f"尝试无密码解压: {archive_path.name}")
            else:
                # 需要密码
                password = self.get_password(archive_path)
                if password == "SKIP":
                    print(f"跳过文件: {archive_path.name}")
                    return False
                elif password is None:
                    print("用户取消操作")
                    return False
            
            # 根据文件格式选择解压方法
            if suffix == '.zip':
                success = self.extract_zip(archive_path, output_path, password)
            elif suffix == '.rar':
                success = self.extract_rar(archive_path, output_path, password)
            elif suffix in ['.7z', '.001']:
                success = self.extract_7z(archive_path, output_path, password)
            else:
                print(f"不支持的格式: {suffix}")
                return False
            
            if not success and password:
                password_attempts += 1
                if password_attempts < max_attempts:
                    print("密码错误，请重新输入")
                else:
                    print(f"密码错误次数过多，跳过文件: {archive_path.name}")
        
        if success:
            print(f"成功解压: {archive_path.name} -> {output_path}")
        else:
            print(f"解压失败: {archive_path.name}")
        
        return success
    
    def process_archives(self, input_folder: Path, output_folder: Path):
        """处理所有压缩包"""
        # 查找所有压缩文件
        archive_files = self.find_archive_files(input_folder)
        if not archive_files:
            print("在输入文件夹中未找到支持的压缩文件")
            return
        
        print(f"\n找到 {len(archive_files)} 个压缩文件")
        
        # 检测分段压缩包
        multi_part_groups, single_files = self.detect_multi_part_archives(archive_files)
        
        if multi_part_groups:
            print(f"检测到 {len(multi_part_groups)} 个分段压缩包")
        
        # 验证所有压缩包
        valid_files, problematic_files = self.verify_all_archives(multi_part_groups, single_files)
        
        # 显示问题文件
        if problematic_files:
            print("\n以下压缩包有问题，将被跳过:")
            for file_path, reason in problematic_files:
                print(f"  • {file_path.name}: {reason}")
        
        if not valid_files:
            print("没有有效的压缩包可以解压")
            return
        
        # 开始解压
        print(f"\n开始解压 {len(valid_files)} 个压缩包...")
        
        success_count = 0
        failed_count = 0
        
        # 处理分段压缩包（只需要处理主文件）
        processed_multi_part = set()
        for base_name, file_group in multi_part_groups.items():
            if file_group[0] in valid_files:
                main_file = file_group[0]
                if self.extract_archive(main_file, output_folder):
                    success_count += 1
                    processed_multi_part.add(base_name)
                else:
                    failed_count += 1
        
        # 处理单个文件
        for file_path in single_files:
            # 检查是否属于已处理的分段压缩包
            is_part_of_processed = False
            for base_name in processed_multi_part:
                if file_path.stem.lower().startswith(base_name.lower()):
                    is_part_of_processed = True
                    break
            
            if not is_part_of_processed and file_path in valid_files:
                if self.extract_archive(file_path, output_folder):
                    success_count += 1
                else:
                    failed_count += 1
        
        # 显示结果
        print(f"\n" + "="*50)
        print("解压完成!")
        print(f"成功: {success_count}")
        print(f"失败: {failed_count}")
        print(f"输出目录: {output_folder}")
        print("="*50)
    
    def run_extraction_cycle(self) -> bool:
        """运行一次解压循环"""
        try:
            # 选择输入文件夹
            input_folder = self.select_input_folder()
            if not input_folder:
                print("用户取消选择输入文件夹")
                return True  # 继续循环
            
            # 选择输出文件夹
            output_folder = self.select_output_folder()
            if not output_folder:
                print("用户取消选择输出文件夹")
                return True  # 继续循环
            
            # 处理压缩包
            self.process_archives(input_folder, output_folder)
                    
            return True
            
        except Exception as e:
            print(f"程序运行出错: {str(e)}")
            return True  # 即使出错也继续循环
    
    def run(self):
        """运行主程序"""
        try:
            # 显示欢迎信息
            print("\n" + "="*60)
            print("欢迎使用批量解压工具!")
            print("="*60)
            print("本工具支持:")
            print("  • ZIP、RAR、7z格式")
            print("  • 分段压缩包")
            print("  • 密码保护文件")
            print("  • 批量处理")
            print("="*60)
            
            # 主循环
            while True:
                # 运行一次解压流程
                self.run_extraction_cycle()
                
                # 询问是否继续
                print("\n" + "="*50)
                choice = self.console_input(
                    "是否要继续进行批量解压？\n"
                    "  Y - 继续解压\n"
                    "  N - 退出程序\n"
                    "请输入选择 (Y/N): ",
                    ['y', 'n']
                )
                
                if choice == 'n':
                    break
            
            print("\n感谢使用批量解压工具！")
            input("\n按 Enter 键退出...")
            
        except Exception as e:
            print(f"程序运行出错: {str(e)}")
        
        finally:
            self.close_ui()

def main():
    """主函数"""
    # 检查必要的库
    try:
        import rarfile
        import py7zr
    except ImportError as e:
        print("缺少必要的库，请安装：")
        print("pip install rarfile py7zr")
        print(f"错误详情: {e}")
        return
    
    # 运行提取器
    extractor = BatchExtractor()
    extractor.run()

if __name__ == "__main__":
    main()