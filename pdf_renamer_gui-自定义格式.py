import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from pathlib import Path
import PyPDF2
import pdfplumber


# --------------------------
# PDF 智能提取函数（保持不变）
# --------------------------

def extract_year_from_text(text):
    year_patterns = [
        r'\b(19[0-9]{2}|20[0-2][0-9])\b',
        r'\((\d{4})\)',
        r'\b(\d{4})\s*[,-]?\s*(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b',
        r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s*[,-]?\s*(\d{4})\b',
    ]
    for pattern in year_patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        if matches:
            for match in matches:
                if isinstance(match, tuple):
                    match = match[0]
                year = int(match)
                if 1900 <= year <= 2030:
                    return str(year)
    return None


def extract_year_from_pdf(pdf_path):
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            metadata = pdf_reader.metadata
            if metadata:
                for field in ['/CreationDate', '/ModDate']:
                    if field in metadata:
                        date_str = metadata[field]
                        year_match = re.search(r'D:(\d{4})', date_str)
                        if year_match:
                            year = year_match.group(1)
                            if 1900 <= int(year) <= 2030:
                                return year
        with pdfplumber.open(pdf_path) as pdf:
            for page_num in range(min(3, len(pdf.pages))):
                page = pdf.pages[page_num]
                text = page.extract_text()
                if text:
                    year = extract_year_from_text(text)
                    if year:
                        return year
                    try:
                        top_region = page.within_bbox((0, 0, page.width, page.height * 0.2))
                        top_text = top_region.extract_text()
                        if top_text:
                            year = extract_year_from_text(top_text)
                            if year:
                                return year
                        bottom_region = page.within_bbox((0, page.height * 0.8, page.width, page.height))
                        bottom_text = bottom_region.extract_text()
                        if bottom_text:
                            year = extract_year_from_text(bottom_text)
                            if year:
                                return year
                    except:
                        pass
    except Exception as e:
        pass
    return None


def extract_title_with_pypdf2(pdf_path):
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            metadata = pdf_reader.metadata
            if metadata and '/Title' in metadata:
                title = metadata['/Title']
                if title and title.strip():
                    return title.strip()
    except Exception as e:
        pass
    return None


def extract_title_with_pdfplumber(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            first_page = pdf.pages[0]
            text = first_page.extract_text()
            if not text:
                return None
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            potential_titles = []
            for i, line in enumerate(lines[:10]):
                if (len(line) > 10 and len(line) < 200 and
                        not re.search(r'abstract|introduction|references|page|\d{1,2}\s*$', line.lower()) and
                        not re.search(r'^[0-9\s\.\-]*$', line)):
                    potential_titles.append((i, line))
            if potential_titles:
                potential_titles.sort(key=lambda x: x[0])
                return potential_titles[0][1]
    except Exception as e:
        pass
    return None


def extract_title_advanced(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            first_page = pdf.pages[0]
            text = first_page.extract_text()
            if not text:
                return None
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            excluded_keywords = [
                'abstract', 'introduction', 'keywords', 'reference',
                'journal', 'vol', 'volume', 'pp', 'page', 'doi',
                'proceedings', 'conference', 'university', 'department'
            ]
            for i, line in enumerate(lines[:15]):
                line_lower = line.lower()
                if (len(line) < 10 or len(line) > 250 or
                        any(keyword in line_lower for keyword in excluded_keywords) or
                        re.search(r'^\d{1,4}\s*$', line) or
                        re.search(r'^[ivxlc]+$', line, re.IGNORECASE) or
                        re.search(r'^[a-z]\s*$', line) or
                        re.search(r'\.{3,}', line) or
                        line.count('.') > 5):
                    continue
                if (re.search(r'[A-Z]', line) and
                        line.count('.') <= 3 and
                        not line.endswith('.') and
                        not line.startswith('Received') and
                        not line.startswith('Copyright')):
                    if ',' in line and len(line.split(',')) <= 3:
                        continue
                    return line
    except Exception as e:
        pass
    return None


def sanitize_filename(title):
    if not title:
        return None
    illegal_chars = r'[<>:"/\\|?*]'
    title = re.sub(illegal_chars, '', title)
    title = re.sub(r'\s+', ' ', title).strip()
    if len(title) > 120:
        title = title[:120] + "..."
    return title


# --------------------------
# 支持自定义格式的新重命名函数
# --------------------------

def rename_pdf_files_custom_format(folder_path, format_template="{title}.pdf", dry_run=True, log_callback=None, progress_callback=None):
    folder = Path(folder_path)
    pdf_files = list(folder.glob("*.pdf"))

    if not pdf_files:
        if log_callback:
            log_callback("未找到PDF文件")
        return 0, [], ["未找到PDF文件"]

    if log_callback:
        log_callback(f"找到 {len(pdf_files)} 个PDF文件，使用格式模板: {format_template}")

    renamed_count = 0
    failed_files = []
    total_files = len(pdf_files)

    for idx, pdf_file in enumerate(pdf_files):
        if progress_callback:
            progress_callback(idx + 1, total_files, pdf_file.name)

        if log_callback:
            log_callback(f"\n处理文件: {pdf_file.name}")

        title = None
        methods = [
            ("元数据提取", extract_title_with_pypdf2),
            ("内容分析", extract_title_with_pdfplumber),
            ("智能识别", extract_title_advanced)
        ]

        for method_name, method_func in methods:
            title = method_func(pdf_file)
            if title:
                if log_callback:
                    log_callback(f"  {method_name}成功: {title[:80]}...")
                break
            else:
                if log_callback:
                    log_callback(f"  {method_name}失败")

        if not title:
            if log_callback:
                log_callback(f"  无法提取标题，跳过此文件")
            failed_files.append(pdf_file.name)
            continue

        clean_title = sanitize_filename(title)
        if not clean_title:
            if log_callback:
                log_callback(f"  标题清理失败，跳过此文件")
            failed_files.append(pdf_file.name)
            continue

        year = extract_year_from_pdf(pdf_file)
        if year:
            if log_callback:
                log_callback(f"  识别到年份: {year}")
        else:
            if log_callback:
                log_callback(f"  未识别到年份，使用'未知年份'")
            year = "未知年份"

        # 替换占位符
        new_filename = format_template.replace('{year}', year).replace('{title}', clean_title)

        new_filepath = pdf_file.parent / new_filename

        counter = 1
        original_new_filepath = new_filepath
        while new_filepath.exists() and new_filepath != pdf_file:
            name_part = os.path.splitext(new_filename)[0]
            ext = os.path.splitext(new_filename)[1]
            new_filename = f"{name_part}_{counter}{ext}"
            new_filepath = pdf_file.parent / new_filename
            counter += 1

        if dry_run:
            if log_callback:
                log_callback(f"  预览重命名: {pdf_file.name} -> {new_filename}")
        else:
            try:
                pdf_file.rename(new_filepath)
                if log_callback:
                    log_callback(f"  成功重命名: {new_filename}")
                renamed_count += 1
            except Exception as e:
                if log_callback:
                    log_callback(f"  重命名失败: {e}")
                failed_files.append(pdf_file.name)

    if log_callback:
        log_callback(f"\n{'=' * 50}")
        log_callback(f"处理完成!")
        if dry_run:
            log_callback(f"预览模式 - 将重命名 {renamed_count} 个文件")
        else:
            log_callback(f"成功重命名 {renamed_count} 个文件")
        if failed_files:
            log_callback(f"失败文件 ({len(failed_files)} 个):")
            for f in failed_files:
                log_callback(f"  - {f}")
        else:
            log_callback("没有失败文件。")
    
    if progress_callback:
        progress_callback(total_files, total_files, "完成")
    
    return renamed_count, failed_files, None


# --------------------------
# GUI 部分（带下拉选择模板 + 进度条）
# --------------------------

class PDFRenamerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("文献PDF 智能重命名工具 - 自定义格式")
        self.root.geometry("800x750")

        # 主框架
        main_frame = ttk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 选择文件夹
        ttk.Label(main_frame, text="选择需要重新命名的文献文件夹:").pack(anchor=tk.W, pady=(0, 5))
        folder_frame = ttk.Frame(main_frame)
        folder_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.folder_path = tk.StringVar()
        ttk.Entry(folder_frame, textvariable=self.folder_path, width=60).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(folder_frame, text="浏览", command=self.browse_folder).pack(side=tk.LEFT)

        # 模式：预览模式
        self.dry_run = tk.BooleanVar(value=True)
        ttk.Checkbutton(main_frame, text="预览模式（不实际重命名，只显示结果）", variable=self.dry_run).pack(anchor=tk.W, pady=(0, 15))

        # 文件名格式模板选择
        ttk.Label(main_frame, text="文件名格式模板:").pack(anchor=tk.W, pady=(0, 5))
        
        # 模板选择变量
        self.use_preset_format = tk.BooleanVar(value=True)
        self.format_template = tk.StringVar(value="{title}.pdf")

        format_frame = ttk.Frame(main_frame)
        format_frame.pack(fill=tk.X, pady=(0, 10))

        # 单选按钮：使用预设 or 自定义
        ttk.Radiobutton(format_frame, text="使用预设模板", variable=self.use_preset_format, 
                        value=True, command=self.on_format_type_change).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(format_frame, text="自定义格式", variable=self.use_preset_format, 
                        value=False, command=self.on_format_type_change).pack(side=tk.LEFT, padx=(0, 10))

        # 下拉选择框（预设模板）
        self.format_combobox = ttk.Combobox(format_frame, textvariable=self.format_template,
                                            values=[
                                                "{title}.pdf",                     # 仅标题
                                                "{year}_{title}.pdf",              # 年份_标题
                                                "{title}_{year}.pdf",              # 标题_年份
                                                "({year})_{title}.pdf",            # (年份)_标题
                                                "{title}-{year}.pdf",              # 标题-年份
                                                "{year}-{title}.pdf",              # 年份-标题
                                            ],
                                            state="readonly", width=30)
        self.format_combobox.pack(side=tk.LEFT, padx=(10, 5))

        # 自定义输入框
        self.custom_format_entry = ttk.Entry(format_frame, textvariable=self.format_template, width=30)

        # 提示
        ttk.Label(format_frame, text="(可用变量: {year}=年份, {title}=标题)", foreground="gray").pack(side=tk.LEFT, padx=(10, 0))

        # 进度条
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill=tk.X, expand=True)
        
        self.progress_label = ttk.Label(progress_frame, text="准备就绪")
        self.progress_label.pack(anchor=tk.W, pady=(5, 0))

        # 日志显示区域
        ttk.Label(main_frame, text="运行日志:").pack(anchor=tk.W, pady=(10, 5))
        
        log_frame = ttk.Frame(main_frame)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=20)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 开始处理按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(15, 0))
        
        self.start_button = ttk.Button(button_frame, text="开始处理", command=self.start_rename, style="Accent.TButton")
        self.start_button.pack()

        # 初始化界面
        self.on_format_type_change()

    def on_format_type_change(self):
        if self.use_preset_format.get():  # 使用预设
            self.format_combobox.pack(side=tk.LEFT, padx=(10, 5))
            self.custom_format_entry.pack_forget()
        else:  # 使用自定义
            self.format_combobox.pack_forget()
            self.custom_format_entry.pack(side=tk.LEFT, padx=(10, 5))

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def update_progress(self, current, total, filename):
        progress_percent = (current / total) * 100 if total > 0 else 0
        self.progress['value'] = progress_percent
        self.progress_label.config(text=f"正在处理: {current}/{total} - {filename}")
        self.root.update_idletasks()

    def start_rename(self):
        folder = self.folder_path.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("错误", "请先选择一个有效的 PDF 文件夹！")
            return

        fmt = self.format_template.get()
        self.log("=" * 50)
        self.log(f"开始处理文件夹: {folder}")
        self.log(f"模式: {'预览模式（不实际重命名）' if self.dry_run.get() else '实际执行模式'}")
        self.log(f"使用文件名格式模板: {fmt}")
        self.log("")

        # 禁用开始按钮，防止重复点击
        self.start_button.config(state="disabled")
        
        thread = threading.Thread(target=self.run_rename, args=(folder, fmt), daemon=True)
        thread.start()

    def run_rename(self, folder, fmt):
        try:
            rename_pdf_files_custom_format(
                folder_path=folder,
                format_template=fmt,
                dry_run=self.dry_run.get(),
                log_callback=self.log,
                progress_callback=self.update_progress
            )
        except Exception as e:
            self.log(f"处理过程中发生错误: {str(e)}")
        finally:
            # 重新启用开始按钮
            self.root.after(0, lambda: self.start_button.config(state="normal"))


# --------------------------
# 启动 GUI
# --------------------------

if __name__ == "__main__":
    root = tk.Tk()
    
    # 设置主题样式
    style = ttk.Style()
    style.theme_use('clam')
    
    # 创建强调按钮样式
    style.configure("Accent.TButton", foreground="white", background="#007acc")
    
    app = PDFRenamerGUI(root)
    root.mainloop()