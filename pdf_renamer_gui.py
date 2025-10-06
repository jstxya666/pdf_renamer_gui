import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from pathlib import Path
import PyPDF2
import pdfplumber


# --------------------------
# PDF 智能处理函数（原逻辑，保持不变）
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
# 专为 GUI 设计的 PDF 重命名函数（不使用 print，收集日志）
# --------------------------


def rename_pdf_files_for_gui(folder_path, dry_run=True, log_callback=None):
    folder = Path(folder_path)
    pdf_files = list(folder.glob("*.pdf"))

    if not pdf_files:
        if log_callback:
            log_callback("未找到PDF文件")
        return 0, [], ["未找到PDF文件"]

    if log_callback:
        log_callback(f"找到 {len(pdf_files)} 个PDF文件")

    renamed_count = 0
    failed_files = []

    for pdf_file in pdf_files:
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

        new_filename = f"{clean_title}.pdf"
        new_filepath = pdf_file.parent / new_filename

        counter = 1
        original_new_filepath = new_filepath
        while new_filepath.exists() and new_filepath != pdf_file:
            new_filename = f"{clean_title}_{counter}.pdf"
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
        if log_callback:
            log_callback(f"\n{'=' * 50}")
            log_callback(f"处理完成!")
            if dry_run:
                log_callback(f"预览模式 - 实际会重命名 {renamed_count} 个文件")
            else:
                log_callback(f"成功重命名 {renamed_count} 个文件")
            if failed_files:
                log_callback(f"失败文件 ({len(failed_files)} 个):")
                for f in failed_files:
                    log_callback(f"  - {f}")
            else:
                if log_callback:
                    log_callback("没有失败文件。")
    return renamed_count, failed_files, None


# --------------------------
# GUI 部分
# --------------------------

class PDFRenamerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("文献PDF智能重命名工具")
        self.root.geometry("700x600")

        # 选择文件夹
        tk.Label(root, text="选择需要重命名的文献文件夹:").pack(pady=5)
        self.folder_path = tk.StringVar()
        tk.Entry(root, textvariable=self.folder_path, width=80).pack(pady=5)
        tk.Button(root, text="浏览文件夹", command=self.browse_folder).pack(pady=5)

        # 模式：预览模式复选框
        self.dry_run = tk.BooleanVar(value=True)
        tk.Checkbutton(root, text=" 预览模式（不实际重命名，只显示结果）", variable=self.dry_run).pack(pady=5)

        # 日志显示区域
        tk.Label(root, text="运行日志:").pack(anchor="w", padx=10)
        self.log_text = scrolledtext.ScrolledText(root, height=25, width=85)
        self.log_text.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

        # 开始处理按钮
        tk.Button(root, text="🚀 开始处理", command=self.start_rename, bg="#4CAF50", fg="white", font=("Arial", 12)).pack(pady=10)

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def start_rename(self):
        folder = self.folder_path.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("错误", "请先选择一个有效的 PDF 文件夹！")
            return

        self.log("=" * 50)
        self.log(f"开始处理文件夹: {folder}")
        self.log(f"模式: {'预览模式（不实际重命名）' if self.dry_run.get() else '实际执行模式'}")
        self.log("")

        # 在新线程中运行
        thread = threading.Thread(target=self.run_rename, args=(folder,), daemon=True)
        thread.start()

    def run_rename(self, folder):
        rename_pdf_files_for_gui(
            folder_path=folder,
            dry_run=self.dry_run.get(),
            log_callback=self.log
        )


# --------------------------
# 启动 GUI
# --------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFRenamerGUI(root)
    root.mainloop()