# Python Tools  
一系列通过Python实现的小工具集合  

首先请确保支持python运行环境，并安装所需依赖库。 

## 📋 工具列表  

### Word_To_Html_Or_Js - Word文档转HTML/JS、CSS文件  
**文件**: `Word_To_Html_Or_Js.py`  

#### 功能描述  
将Word文档(.docx格式)转换为HTML(.html)/JavaScript(.js)和CSS(.css)文件，便于在网页中使用Word文档内容。  

#### 注意事项  
-  仅支持 .docx 格式  
-  生成后请手动检查和调整样式以确保符合需求

#### 安装额外依赖库  

-  安装额外依赖库:    
```bash
pip install python-docx
```

- 相关示例  
![Word_To_Html_Or_Js.py示例图](images/Word_To_Html_Or_Js.png)  

### Picture_Batch_Conv - 图片批量格式转换工具  
**文件**: `Picture_Batch_Conv.py`  

#### 功能描述  
支持多种图片格式之间的批量转换，包括PNG、WEBP、JPG、JPEG、JPE、TIF、TIFF、BMP等格式。 

#### 注意事项  
- 转换为JPG等格式时自动填充白色背景  
- 输入文件夹中不支持的文件会被跳过并提示  

#### 安装额外依赖库  
    
```bash
pip install Pillow
```  

- 相关示例  
![Picture_Batch_Conv.py示例图](images/Picture_Batch_Conv.png)  

### Batch_Decompress - 批量解压工具  
**文件**: `Batch_Decompress.py`  

#### 功能描述  
批量解压多种格式的压缩文件，支持ZIP、RAR、7z等常见格式，能够自动识别分段压缩包并合并解压。  

#### 注意事项  
##### 系统要求  

- Windows: 需要安装WinRAR或unrar工具  

- macOS: brew install unrar  

- Linux: sudo apt-get install unrar (Ubuntu/Debian)  

#### 安装额外依赖库  
    
```bash
pip install rarfile py7zr
```  

- 相关示例  
![Batch_Decompress.py示例图](images/Batch_Decompress.png)

### PathFinder - 跨平台文件/目录路径查找工具  
**文件**: `PathFinder.py`  

#### 功能描述  
提供统一接口在 Windows、macOS、Linux 上查找指定的文件或目录，特别支持通过 Steam 库定位游戏安装路径，也支持手动指定目录进行搜索并返回匹配路径。

- **双模式**：`steam`（自动检测 Steam 库并查找游戏）和 `manual`（直接指定路径或在目录内搜索）。
- **交互支持**：可弹出图形化对话框（基于 `tkinter`）让用户选择匹配项或输入路径；也支持非交互模式（`interactive=False`）直接返回第一个匹配或 `None`。
- **灵活返回**：可返回文件路径或目录路径（通过 `target_type` 参数指定）。

#### 注意事项  
- 交互模式依赖 `tkinter`（通常随 Python 自带），若不可用则自动回退控制台交互。
- 当 `manual_path` 为空字符串且 `interactive=True` 时，会弹出文件/目录选择对话框（根据 `target_type`）。
- 无需安装第三方库，仅使用 Python 标准库。

#### 安装额外依赖库  
无需额外依赖（可选：若需图形化文件选择，确保 `tkinter` 可用）。

- 相关示例  
（暂无示例图）
