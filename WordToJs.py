### Word文档转换为JS文件的工具（改进版）
## 需要安装依赖: pip install python-docx

import os
import json
import re
from docx import Document
from pathlib import Path

class WordToJSConverter:
    """
    Word文档转JS转换器
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

    def pt_to_class_name(self, font_size_pt):
        """
        将Word字体大小(磅)转换为CSS类名（改进版）
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
        处理单个段落，返回HTML字符串（改进版）
        """
        html_parts = []
        
        # 获取段落对齐方式
        alignment = 'left'
        if paragraph.alignment is not None:
            alignment = self.alignment_map.get(paragraph.alignment, 'left')
        
        for run in paragraph.runs:
            text = run.text.strip()
            if not text:
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
            combined_html = ' '.join(html_parts)
            return f'<p class="align-{alignment}">{combined_html}</p>'
        return ""

    def convert_word_to_js(self, docx_path, output_dir=None):
        """
        将Word文档转换为JS文件（改进版）
        """
        try:
            # 检查文件是否存在
            if not os.path.exists(docx_path):
                print(f"错误: 文件不存在: {docx_path}")
                return False
                
            # 检查文件格式
            if not docx_path.lower().endswith('.docx'):
                print(f"错误: 不支持的文件格式: {docx_path}")
                return False
            
            print(f"正在处理: {Path(docx_path).name}")
            
            # 读取Word文档
            doc = Document(docx_path)
            
            # 处理所有段落
            html_content = []
            for i, paragraph in enumerate(doc.paragraphs):
                if paragraph.text.strip():  # 跳过空段落
                    html_para = self.process_paragraph(paragraph)
                    if html_para:
                        html_content.append(html_para)
            
            # 组合所有HTML内容
            full_html = '\n        '.join(html_content)
            
            # 使用JSON安全地转义HTML内容
            escaped_html = json.dumps(full_html)[1:-1]  # 移除外层的引号
            
            # 生成JS文件内容（改进版，避免模板字符串问题）
            js_content = f"""// 自动生成的文档内容
const CONTENT_DATA = {{
    content: `{escaped_html}`
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
            # 确定输出路径
            if output_dir is None:
                output_dir = Path.cwd()
            else:
                output_dir = Path(output_dir)
                output_dir.mkdir(exist_ok=True)
            
            # 生成输出文件名
            input_path = Path(docx_path)
            output_filename = input_path.stem + '.js'
            output_path = output_dir / output_filename
            
            # 写入JS文件
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(js_content)
            
            print(f"✓ 成功转换: {input_path.name} -> {output_path.name}")
            return True
            
        except Exception as e:
            print(f"✗ 转换失败 {docx_path}: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

    def batch_convert_word_files(self, directory=None):
        """
        批量转换目录下的所有Word文档
        """
        if directory is None:
            directory = Path.cwd()
        else:
            directory = Path(directory)
        
        if not directory.exists():
            print(f"错误: 目录不存在: {directory}")
            return
        
        # 支持的Word文档扩展名
        word_extensions = ['.docx']  # 主要支持.docx格式
        
        converted_count = 0
        total_count = 0
        
        print(f"开始在目录 {directory} 中查找Word文档...")
        
        for file_path in directory.iterdir():
            if file_path.is_file() and file_path.suffix.lower() in word_extensions:
                total_count += 1
                if self.convert_word_to_js(file_path, directory):
                    converted_count += 1
        
        print(f"\n转换完成: {converted_count}/{total_count} 个文件成功转换")
        return converted_count

    def generate_sample_files(self, output_dir=None):
        """
        生成示例CSS文件和HTML文件
        """
        if output_dir is None:
            output_dir = Path.cwd()
        
        # 生成CSS文件
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
    margin: 8px 0;
    padding: 4px 0;
}

/* 文本样式 */
span {
    display: inline;
}

/* 页面容器 */
.content-container {
    max-width: 800px;
    margin: 0 auto;
    padding: 20px;
    font-family: 'Microsoft YaHei', 'SimSun', serif;
    background: white;
}
"""
        
        css_path = output_dir / 'word-styles.css'
        with open(css_path, 'w', encoding='utf-8') as f:
            f.write(css_content)
        
        # 生成HTML示例文件
        html_content = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Word文档内容展示</title>
    <link rel="stylesheet" href="word-styles.css">
    <style>
        body {
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
            font-family: Arial, sans-serif;
        }
        .container {
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            padding: 30px;
            margin: 0 auto;
            max-width: 900px;
        }
        h1 {
            text-align: center;
            color: #333;
            border-bottom: 2px solid #eee;
            padding-bottom: 10px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>文档内容展示</h1>
        <div id="word-content" class="content-container">
            <!-- 内容将通过JS自动填充 -->
        </div>
    </div>
    
    <!-- 引入生成的JS文件（请将example.js替换为实际的文件名） -->
    <script src="example.js"></script>
</body>
</html>
"""
        
        html_path = output_dir / 'example.html'
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"✓ 示例CSS文件已生成: {css_path}")
        print(f"✓ 示例HTML文件已生成: {html_path}")
        print("\n使用说明:")
        print("1. 将生成的JS文件重命名为example.html中引用的文件名")
        print("2. 用浏览器打开example.html查看效果")

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
    
    print("Word文档转JS转换器（改进版）")
    print("=" * 50)
    
    converter = WordToJSConverter()
    
    while True:
        print("\n请选择操作:")
        print("1. 批量转换当前目录下的Word文档")
        print("2. 转换指定Word文档")
        print("3. 生成示例文件")
        print("4. 退出")
        
        choice = input("请输入选择 (1-4): ").strip()
        
        if choice == '1':
            directory = input("请输入目录路径（直接回车使用当前目录）: ").strip()
            if not directory:
                directory = None
            converter.batch_convert_word_files(directory)
            
        elif choice == '2':
            file_path = input("请输入Word文档路径: ").strip()
            if file_path:
                output_dir = input("请输入输出目录（直接回车使用当前目录）: ").strip()
                if not output_dir:
                    output_dir = None
                converter.convert_word_to_js(file_path, output_dir)
            else:
                print("文件路径不能为空")
                
        elif choice == '3':
            output_dir = input("请输入输出目录（直接回车使用当前目录）: ").strip()
            if not output_dir:
                output_dir = None
            converter.generate_sample_files(output_dir)
            
        elif choice == '4':
            print("再见！")
            break
            
        else:
            print("无效选择，请重新输入")

if __name__ == "__main__":
    main()