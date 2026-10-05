# -*- coding: utf-8 -*-
"""Step2: 精简阅读页签 + _reader_open 仅PDF + __init__ 去 IPC/全局键绑定 + 文件类型仅PDF。"""
import io

PATH = r"excel_pdf_sync-260729-usefull.py"
with io.open(PATH, "r", encoding="utf-8") as f:
    lines = f.readlines()

def find_idx(substr, start=0):
    idx = -1
    for i in range(start, len(lines)):
        if substr in lines[i]:
            if idx >= 0:
                raise ValueError("start anchor non-unique: " + substr)
            idx = i
    if idx < 0:
        raise ValueError("anchor not found: " + substr)
    return idx

def delete_between(start_anchor, end_anchor):
    """删除 [start_anchor 所在行, end_anchor 所在行)，end 保留。"""
    global lines
    s = find_idx(start_anchor)
    e = find_idx(end_anchor, s + 1)
    del lines[s:e]

def delete_lines_containing(substr):
    global lines
    lines = [l for l in lines if substr not in l]

def replace_method_body(def_anchor, next_anchor, new_body):
    global lines
    d = find_idx(def_anchor)
    n = find_idx(next_anchor, d + 1)
    nb = new_body.splitlines(keepends=True)
    lines[d + 1:n] = nb

# ═════════════ 1. 删除阅读页签左侧搜索区块 ═════════
delete_between("# ── 常用文件（默认收起", "# 文件夹选择")
delete_between("# ── Everything 全局搜索", "# ── 文件内容搜索")
delete_between("# ── 文件内容搜索", "# ── 标记搜索（标签/笔记")
delete_between("# ── 标记搜索（标签/笔记", "# ── 网络 PDF 搜索（TI 官方直链")
delete_between("# ── 网络 PDF 搜索（TI 官方直链", "# 搜索过滤")

# ═════════════ 2. 删除阅读页签导航栏无用按钮 ═════════
delete_between("# ── 放大镜控件", "self._reader_fit_btn = RoundedButton")          # 放大镜 + 弹窗模式
delete_between("self._reader_region_label_btn = RoundedButton", "self._reader_dual_btn = RoundedButton")  # 位置标注 + EPUB样式
delete_between("self._reader_dual_btn = RoundedButton", "# ── 滚轮横滚模式控件")     # 多开
delete_between("# ── 滚轮横滚模式控件", "# 颜色主题选择按钮")                        # 横滚
delete_between("# Edge浏览器切换按钮", "self._reader_ai_btn = RoundedButton")        # Edge按钮
delete_between("self._reader_ai_btn = RoundedButton", "self._reader_toc_all_btn = RoundedButton")  # AI按钮
delete_between("self._reader_toc_all_btn = RoundedButton", "# 始终可见的文件栏开关")   # ≡目录/≡匹配/联动/同步/快捷键/跨程序/工作区

# ═════════════ 3. 删除第 3/4 阅读区 ═════════
delete_lines_containing("_build_reader_aux_panel(reader_view_paned, 3)")
delete_lines_containing("_build_reader_aux_panel(reader_view_paned, 4)")

# ═════════════ 4. 删除 Edge 面板 + AI 面板 + 嵌入状态 + 搜索历史 ═════════
delete_between("# ── Edge浏览器面板（PanedWindow右窗格，默认隐藏）──", "self.after_idle(lambda: self._reader_set_active_slot(1))")

# ═════════════ 5. __init__ 去 IPC / 全局键绑定 ═════════
delete_between("# ── 多实例联动：跨程序发送", "self._build_ui()")
delete_between("# 启动收件箱轮询", "# 阅读窗口最大化状态")

# ═════════════ 6. READER_SUPPORTED_EXTS 仅 PDF ═════════
for i, l in enumerate(lines):
    if l.startswith("READER_SUPPORTED_EXTS ="):
        lines[i] = 'READER_SUPPORTED_EXTS = (".pdf",)\n'
        break

# ═════════════ 7. _reader_open 仅 PDF ═════════
replace_method_body(
    "def _reader_open(self, full_path, rel_path):",
    "def _reader_open_office_file",
    '''        # ── 重入保护 + 文件存在性检查 ──
        if self._reader_opening:
            return
        if not full_path or not os.path.isfile(full_path):
            messagebox.showerror("打开失败", f"文件不存在或路径无效:\\n{full_path}")
            return
        self._reader_save_resume()
        self._record_open(full_path)
        try:
            self._reader_opening = True
            self._reader_doc_generation += 1
            generation = self._reader_doc_generation
            if self._reader_lazy_job:
                try:
                    self.after_cancel(self._reader_lazy_job)
                except Exception:
                    pass
                self._reader_lazy_job = None
            if self._reader_scroll_anim_job:
                try:
                    self.after_cancel(self._reader_scroll_anim_job)
                except Exception:
                    pass
                self._reader_scroll_anim_job = None
                self._reader_scroll_velocity = 0.0
            if self._reader_pdf:
                self._reader_pdf.close()
                self._reader_pdf = None
            self._reader_load_toc()
            self._reader_page_cache.clear()
            self._reader_text_cache = None
            self._reader_word_index = None
            self._reader_yoffset_estimated = False
            self._reader_file_type = ""
            self._reader_match_lb.delete(0, tk.END)
            self._reader_match_pages = []
            self._reader_match_counts = {}
            self._reader_refresh_main_match_flow()
            self._reader_match_count_var.set("")
            self._reader_preview_search.set("")
            self._reader_canvas.grid()

            ext = os.path.splitext(full_path)[1].lower()
            if ext != ".pdf":
                messagebox.showerror("打开失败", f"仅支持 PDF 文件:\\n{rel_path}")
                return
            self._reader_file_type = ext
            self._reader_current_path = full_path
            self._reader_active_file_path = full_path
            doc = fitz.open(full_path)
            self._reader_apply_pdf_doc(doc, rel_path, icon="📄", generation=generation)
        except Exception as e:
            messagebox.showerror("打开失败", f"无法打开文件: {e}")
            self._reader_pdf = None
            self._reader_load_toc()
        finally:
            self._reader_opening = False
''')

with io.open(PATH, "w", encoding="utf-8", newline="") as f:
    f.writelines(lines)

print("done, total lines:", len(lines))
