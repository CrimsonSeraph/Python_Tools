# Word文档转换为JS/HTML文件的工具
# 需要安装依赖: pip install python-docx

import os
import json
import re
from docx import Document
from pathlib import Path

class WordToFileConverter:
    """
    Word文档转文件转换器
    """
    
    def __init__(self):
        # 中文字号与磅值对应关系
        self.size_mapping = {
            42: 'initial',        # 初号
            36: 'small-initial',  # 小初
            26: 'one',            # 一号
            24: 'small-one',      # 小一
            22: 'two',            # 二号
            18: 'small-two',      # 小二
            16: 'three',          # 三号
            15: 'small-three',    # 小三
            14: 'four',           # 四号
            12: 'small-four',     # 小四
            10.5: 'five',         # 五号
            9: 'small-five',      # 小五
            7.5: 'six',           # 六号
            6.5: 'small-six',     # 小六
            5.5: 'seven',         # 七号
            5: 'eight'            # 八号
        }
        
        # 对齐方式映射
        self.alignment_map = {
            0: 'left',    # 左对齐
            1: 'center',  # 居中  
            2: 'right',   # 右对齐
            3: 'justify'  # 两端对齐
        }

    def select_folder_dialog(self, title):
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

    def select_file_dialog(self, title, filetypes):
        """使用tkinter弹出文件选择对话框"""
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
        
        # 弹出文件选择对话框
        file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        
        # 销毁根窗口
        root.destroy()
        
        return file_path

    def pt_to_class_name(self, font_size_pt):
        """
        将Word字体大小(磅)转换为CSS类名
        """
        if font_size_pt is None:
            return "normal"
        
        # 尝试精确匹配
        if font_size_pt in self.size_mapping:
            return self.size_mapping[font_size_pt]
        
        # 容差匹配（允许0.5磅误差）
        for standard_pt, class_name in self.size_mapping.items():
            if abs(font_size_pt - standard_pt) < 0.5:
                return class_name
        
        # 处理数字和半数字情况
        if isinstance(font_size_pt, (int, float)):
            # 检查是否为半磅值
            if font_size_pt % 1 == 0.5:
                integer_part = int(font_size_pt)
                return f"num-half-{integer_part}"
            else:
                # 整数磅值
                integer_part = int(round(font_size_pt))
                return f"num-{integer_part}"
        
        # 默认情况
        return "normal"

    def get_run_styles(self, run):
        """
        获取run的样式信息
        """
        styles = []
        
        # 粗体
        if run.bold:
            styles.append("font-weight:bold")
        
        # 斜体  
        if run.italic:
            styles.append("font-style:italic")
        
        # 下划线
        if run.underline:
            styles.append("text-decoration:underline")
        
        # 字体颜色
        if run.font.color and run.font.color.rgb:
            try:
                color = run.font.color.rgb
                # 将颜色转换为16进制
                if hasattr(color, '__iter__'):
                    # 如果是RGB元组
                    color_hex = ''.join(f'{c:02x}' for c in color)
                else:
                    # 如果是整数
                    color_hex = f'{color:06x}'
                styles.append(f"color:#{color_hex}")
            except:
                pass  # 颜色解析失败时忽略
        
        # 字体名称
        if run.font.name:
            styles.append(f"font-family:'{run.font.name}'")
        
        return styles

    def process_paragraph(self, paragraph):
        """
        处理单个段落，返回HTML字符串
        """
        html_parts = []
        
        # 获取段落对齐方式
        alignment = 'left'
        if paragraph.alignment is not None:
            alignment = self.alignment_map.get(paragraph.alignment, 'left')
        
        for run in paragraph.runs:
            text = run.text
            if not text.strip():
                continue
                
            # 获取字体大小
            font_size_pt = None
            if run.font.size:
                font_size_pt = run.font.size.pt
            
            # 获取CSS类名
            css_class = self.pt_to_class_name(font_size_pt)
            
            # 获取内联样式
            styles = self.get_run_styles(run)
            style_attr = f' style="{"; ".join(styles)}"' if styles else ""
            
            # 构建HTML元素
            if css_class != "normal" or style_attr:
                class_attr = f' class="{css_class}"' if css_class != "normal" else ""
                html_parts.append(f'<span{class_attr}{style_attr}>{text}</span>')
            else:
                html_parts.append(text)
        
        if html_parts:
            # 合并连续的文本
            combined_html = ''.join(html_parts)
            return f'<p class="align-{alignment}">{combined_html}</p>'
        return ""

    def convert_word_to_file(self, docx_path, output_type='js', output_dir=None):
        """
        将Word文档转换为指定类型的文件
        output_type: 'js' 或 'html'
        """
        try:
            # 将路径转换为Path对象
            input_path = Path(docx_path)
            
            # 检查文件是否存在
            if not input_path.exists():
                print(f"错误: 文件不存在: {input_path}")
                return False
                
            # 使用Path对象的suffix属性检查文件格式
            if input_path.suffix.lower() not in ['.docx']:
                print(f"错误: 不支持的文件格式: {input_path}")
                return False
            
            print(f"正在处理: {input_path.name}")
            
            # 读取Word文档
            doc = Document(input_path)
            
            # 处理所有段落
            html_content = []
            for i, paragraph in enumerate(doc.paragraphs):
                if paragraph.text.strip():  # 跳过空段落
                    html_para = self.process_paragraph(paragraph)
                    if html_para:
                        html_content.append(html_para)
            
            # 组合所有HTML内容，每行一个段落
            full_html = '\n'.join(html_content)
            
            # 确定输出路径
            if output_dir is None:
                output_dir = Path.cwd()
            else:
                output_dir = Path(output_dir)
                output_dir.mkdir(exist_ok=True)
            
            # 生成输出文件名
            output_filename = input_path.stem + f'.{output_type}'
            output_path = output_dir / output_filename
            
            # 根据输出类型生成不同的内容
            if output_type == 'js':
                # 生成JS文件内容
                file_content = f"""// 自动生成的文档内容
const CONTENT_DATA = {{
    content: `{full_html}`
}};

// 导出供其他模块使用
if (typeof module !== 'undefined' && module.exports) {{
    module.exports = CONTENT_DATA;
}}

// 自动插入内容到页面
if (typeof document !== 'undefined') {{
    document.addEventListener('DOMContentLoaded', function() {{
        const contentElement = document.getElementById('word-content');
        if (contentElement) {{
            contentElement.innerHTML = CONTENT_DATA.content;
        }}
    }});
}}
"""
            elif output_type == 'html':
                # 生成HTML文件内容
                file_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{input_path.stem}</title>
    <link rel="stylesheet" href="word-styles.css">
    <style>
        body {{
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
            font-family: Arial, sans-serif;
        }}
        .container {{
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            padding: 30px;
            margin: 0 auto;
            max-width: 900px;
        }}
        .header {{
            text-align: center;
            color: #333;
            border-bottom: 2px solid #eee;
            padding-bottom: 10px;
            margin-bottom: 20px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>{input_path.stem}</h1>
        </div>
        <div id="word-content" class="content-container">
{full_html}
        </div>
    </div>
</body>
</html>
"""
            else:
                print(f"错误: 不支持的输出类型: {output_type}")
                return False
            
            # 写入文件
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(file_content)
            
            # 确保CSS文件存在
            self.ensure_css_file(output_dir)
            
            print(f"✓ 成功转换: {input_path.name} -> {output_path}")
            return True
            
        except Exception as e:
            print(f"✗ 转换失败 {docx_path}: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

    def ensure_css_file(self, output_dir):
        """
        确保CSS文件存在
        """
        css_path = Path(output_dir) / 'word-styles.css'
        if not css_path.exists():
            self.generate_css_file(output_dir)

    def generate_css_file(self, output_dir=None):
        """
        生成CSS文件
        """
        if output_dir is None:
            output_dir = Path.cwd()
        
        css_content = """/* Word文档字体大小样式定义 */
/* 中文字号 */
.initial { font-size: 42pt; }        /* 初号 */
.small-initial { font-size: 36pt; }  /* 小初 */
.one { font-size: 26pt; }            /* 一号 */
.small-one { font-size: 24pt; }      /* 小一 */
.two { font-size: 22pt; }            /* 二号 */
.small-two { font-size: 18pt; }      /* 小二 */
.three { font-size: 16pt; }          /* 三号 */
.small-three { font-size: 15pt; }    /* 小三 */
.four { font-size: 14pt; }           /* 四号 */
.small-four { font-size: 12pt; }     /* 小四 */
.five { font-size: 10.5pt; }         /* 五号 */
.small-five { font-size: 9pt; }      /* 小五 */
.six { font-size: 7.5pt; }           /* 六号 */
.small-six { font-size: 6.5pt; }     /* 小六 */
.seven { font-size: 5.5pt; }         /* 七号 */
.eight { font-size: 5pt; }           /* 八号 */

/* 数字字号 */
.num-6 { font-size: 6pt; }
.num-7 { font-size: 7pt; }
.num-8 { font-size: 8pt; }
.num-9 { font-size: 9pt; }
.num-10 { font-size: 10pt; }
.num-11 { font-size: 11pt; }
.num-12 { font-size: 12pt; }
.num-13 { font-size: 13pt; }
.num-14 { font-size: 14pt; }
.num-15 { font-size: 15pt; }
.num-16 { font-size: 16pt; }
.num-18 { font-size: 18pt; }
.num-20 { font-size: 20pt; }
.num-22 { font-size: 22pt; }
.num-24 { font-size: 24pt; }
.num-26 { font-size: 26pt; }
.num-28 { font-size: 28pt; }
.num-30 { font-size: 30pt; }
.num-32 { font-size: 32pt; }
.num-36 { font-size: 36pt; }
.num-40 { font-size: 40pt; }
.num-48 { font-size: 48pt; }
.num-56 { font-size: 56pt; }
.num-72 { font-size: 72pt; }

/* 半数字字号 */
.num-half-8 { font-size: 8.5pt; }
.num-half-9 { font-size: 9.5pt; }
.num-half-10 { font-size: 10.5pt; }
.num-half-11 { font-size: 11.5pt; }
.num-half-12 { font-size: 12.5pt; }
.num-half-14 { font-size: 14.5pt; }
.num-half-16 { font-size: 16.5pt; }
.num-half-18 { font-size: 18.5pt; }

/* 段落对齐 */
.align-left { text-align: left; }
.align-center { text-align: center; }
.align-right { text-align: right; }
.align-justify { text-align: justify; }

/* 段落基础样式 */
p {
    line-height: 1.6;
    margin: 12px 0;
    padding: 4px 0;
    page-break-inside: avoid;
}

/* 文本样式 */
span {
    display: inline;
    white-space: pre-wrap;
}

/* 页面容器 */
.content-container {
    max-width: 800px;
    margin: 0 auto;
    padding: 20px;
    font-family: 'Microsoft YaHei', 'SimSun', serif;
    background: white;
    line-height: 1.6;
}
"""
        
        css_path = output_dir / 'word-styles.css'
        with open(css_path, 'w', encoding='utf-8') as f:
            f.write(css_content)
        
        print(f"✓ CSS文件已生成: {css_path}")

    def batch_convert_word_files(self, input_dir=None, output_type='js', output_dir=None):
        """
        批量转换目录下的所有Word文档
        """
        if input_dir is None:
            input_dir = Path.cwd()
        else:
            input_dir = Path(input_dir)
        
        if not input_dir.exists():
            print(f"错误: 输入目录不存在: {input_dir}")
            return
        
        # 如果未指定输出目录，使用输入目录
        if output_dir is None:
            output_dir = input_dir
        else:
            output_dir = Path(output_dir)
            output_dir.mkdir(exist_ok=True)
        
        # 支持的Word文档扩展名
        word_extensions = ['.docx']
        
        converted_count = 0
        total_count = 0
        
        print(f"开始在目录 {input_dir} 中查找Word文档...")
        print(f"输出目录: {output_dir}")
        
        # 确保CSS文件存在
        self.ensure_css_file(output_dir)
        
        for file_path in input_dir.iterdir():
            if file_path.is_file() and file_path.suffix.lower() in word_extensions:
                total_count += 1
                if self.convert_word_to_file(file_path, output_type, output_dir):
                    converted_count += 1
        
        print(f"\n转换完成: {converted_count}/{total_count} 个文件成功转换")
        print(f"输出目录: {output_dir}")
        return converted_count

    def generate_sample_files(self, output_dir=None):
        """
        生成示例文件
        """
        if output_dir is None:
            output_dir = Path.cwd()
        
        # 生成CSS文件
        self.generate_css_file(output_dir)
        
        print(f"✓ 示例文件已生成在: {output_dir}")
        print("\n使用说明:")
        print("1. 生成的word-styles.css包含所有字体样式定义")
        print("2. JS文件需要通过HTML引入并调用CONTENT_DATA.content")
        print("3. HTML文件可以直接在浏览器中打开查看")

def main():
    """
    主函数
    """
    # 安装依赖提示
    try:
        from docx import Document
    except ImportError:
        print("请先安装 python-docx: pip install python-docx")
        exit(1)
    
    print("Word文档转文件转换器")
    print("=" * 50)
    
    converter = WordToFileConverter()
    
    while True:
        print("\n请选择操作:")
        print("1. 批量转换Word文档")
        print("2. 转换单个Word文档")
        print("3. 生成CSS文件")
        print("4. 退出")
        
        choice = input("请输入选择 (1-4): ").strip()
        
        if choice == '1':
            print("\n请选择输入目录...")
            input_dir = converter.select_folder_dialog("选择包含Word文档的目录")
            
            if not input_dir:
                print("未选择目录，操作取消")
                continue
                
            print(f"已选择输入目录: {input_dir}")
            
            # 选择输出目录
            print("\n请选择输出目录...")
            output_dir = converter.select_folder_dialog("选择输出目录")
            
            if not output_dir:
                output_dir = input_dir  # 如果未选择输出目录，使用输入目录
                print(f"使用输入目录作为输出目录: {input_dir}")
            else:
                print(f"已选择输出目录: {output_dir}")
            
            print("\n请选择输出格式:")
            print("1. JS文件")
            print("2. HTML文件")
            format_choice = input("请输入选择 (1-2): ").strip()
            
            output_type = 'js' if format_choice == '1' else 'html'
            
            print(f"\n开始批量转换...")
            converter.batch_convert_word_files(input_dir, output_type, output_dir)
            
        elif choice == '2':
            print("\n请选择Word文档...")
            file_path = converter.select_file_dialog(
                "选择Word文档", 
                [("Word文档", "*.docx"), ("所有文件", "*.*")]
            )
            
            if not file_path:
                print("未选择文件，操作取消")
                continue
                
            print(f"已选择文件: {file_path}")
            
            # 选择输出目录
            print("\n请选择输出目录...")
            output_dir = converter.select_folder_dialog("选择输出目录")
            
            if not output_dir:
                output_dir = os.path.dirname(file_path)  # 如果未选择输出目录，使用文件所在目录
                print(f"使用文件所在目录作为输出目录: {output_dir}")
            else:
                print(f"已选择输出目录: {output_dir}")
            
            print("\n请选择输出格式:")
            print("1. JS文件")
            print("2. HTML文件")
            format_choice = input("请输入选择 (1-2): ").strip()
            
            output_type = 'js' if format_choice == '1' else 'html'
            
            print(f"\n开始转换...")
            converter.convert_word_to_file(file_path, output_type, output_dir)
                
        elif choice == '3':
            print("\n请选择CSS文件输出目录...")
            output_dir = converter.select_folder_dialog("选择CSS文件输出目录")
            
            if not output_dir:
                output_dir = None
                print("使用当前目录")
            else:
                print(f"已选择输出目录: {output_dir}")
                
            converter.generate_css_file(output_dir)
            
        elif choice == '4':
            print("再见！")
            input("\n按 Enter 键退出...")
            break
            
        else:
            print("无效选择，请重新输入")

if __name__ == "__main__":
    main()