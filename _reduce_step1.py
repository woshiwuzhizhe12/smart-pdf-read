# -*- coding: utf-8 -*-
"""Step1: 重写 6 个入口方法。splice 必须按原始行号严格降序执行。"""
import io

PATH = r"excel_pdf_sync-260729-usefull.py"

with io.open(PATH, "r", encoding="utf-8") as f:
    lines = f.readlines()

def splice(def_line, end_line, new_text):
    """def_line/end_line 1-indexed：保留 def_line 这一行，替换 [def_line+1 .. end_line]。"""
    global lines
    new_lines = new_text.splitlines(keepends=True)
    lines = lines[:def_line] + new_lines + lines[end_line:]

# ══ 按降序执行 ══

# 1. on_close: def 26792, 尾 26900
splice(26792, 26900, '''        self._reader_doc_generation += 1
        for job_name in ("_reader_lazy_job", "_reader_scroll_job", "_reader_render_job",
                         "_reader_zoom_job", "_reader_scroll_anim_job"):
            job = getattr(self, job_name, None)
            if job:
                try:
                    self.after_cancel(job)
                except Exception:
                    pass
                setattr(self, job_name, None)
        self._reader_save_resume()
        for slot in self._reader_aux:
            self._reader_save_aux_resume(slot)
        self._save_config()
        if self._reader_pdf:
            self._reader_pdf.close()
        for state in self._reader_aux.values():
            for key in ("resize_job", "scroll_anim_job"):
                if state.get(key):
                    try:
                        self.after_cancel(state[key])
                    except Exception:
                        pass
            if state.get("doc"):
                try:
                    state["doc"].close()
                except Exception:
                    pass
        for image in self._diff_images:
            try:
                image.close()
            except Exception:
                pass
        self._diff_images.clear()
        self._reader_page_cache.clear()
        self.destroy()
''')

# 2. _shortcut_registry: def 20102, 尾 20134
splice(20102, 20134, '''        """快捷键注册表：返回 [(id, 说明, 默认绑定序列, 处理函数), ...]。"""
        reg = [
            ("focus_search", "聚焦搜索框", ["<Control-f>", "<Control-F>"], self._shortcut_focus_search),
            ("add_bookmark", "添加标记", ["<Control-b>", "<Control-B>"], self._shortcut_add_bookmark),
            ("toggle_bookmarks", "显示/隐藏标记面板", ["<Control-Shift-B>", "<Control-Shift-b>"], self._shortcut_toggle_bookmarks),
            ("page_prev", "上一页", ["<Prior>", "<Left>"], lambda e: self._shortcut_reader_page(-1)),
            ("page_next", "下一页", ["<Next>", "<Right>"], lambda e: self._shortcut_reader_page(1)),
            ("first_page", "跳到第一页", ["<Home>"], self._shortcut_reader_first_page),
            ("last_page", "跳到最后一页", ["<End>"], self._shortcut_reader_last_page),
            ("zoom_in", "放大", ["<Control-equal>", "<Control-plus>"], lambda e: self._shortcut_reader_zoom(1.15)),
            ("zoom_out", "缩小", ["<Control-minus>"], lambda e: self._shortcut_reader_zoom(1 / 1.15)),
            ("zoom_reset", "恢复适应宽度", ["<Control-Key-0>"], self._shortcut_reader_zoom_reset),
            ("toggle_crop", "开启/关闭智能裁边", ["<Control-e>", "<Control-E>"], self._shortcut_toggle_crop),
            ("toggle_fit_width", "开启/关闭适应宽度", ["<Control-w>", "<Control-W>"], self._shortcut_toggle_fit_width),
        ]
        for s in (1, 2):
            reg.append((f"slot_{s}", f"切换阅读区·PDF{s}",
                        [f"<Alt-Key-{s}>"],
                        lambda e, s=s: self._shortcut_set_active_slot(s)))
        return reg
''')

# 3. _setup_keyboard_shortcuts: def 20076, 尾 20100
splice(20076, 20100, '''        """应用内快捷键：翻页 / 缩放 / 聚焦搜索 / 添加标记。"""
        self._shortcut_bound_seqs = set()
        self._shortcut_apply_bindings()
''')

# 4. _save_config: def 7494, 尾 7589
splice(7494, 7589, '''        import json
        try:
            with open(self._config_file, "w", encoding="utf-8") as f:
                json.dump({
                    "pdf_folder": self._reader_dir,
                    "reader_highlight_color": self._reader_highlight_hex.get(),
                    "reader_open_counts": self._open_counts,
                    "reader_resume": self._reader_resume,
                    "reader_annotations": self._reader_annotations,
                    "reader_page_marks": self._reader_page_marks,
                    "reader_text_marks": self._reader_text_marks,
                    "reader_text_hl": self._reader_text_hl,
                    "reader_theme": self._reader_theme.get(),
                    "reader_crop_margins": self._reader_crop_margins_mode,
                    "reader_fit_width": self._reader_fit_width,
                    "reader_scroll_lines": self._reader_scroll_lines,
                    "reader_body_crop_regions": {
                        k: list(v) for k, v in self._reader_body_crop_regions.items()},
                    "excel_search_search_path": self._excel_search_search_path,
                    "excel_search_recent_paths": self._excel_search_recent_paths,
                    "excel_search_recursive": self._excel_search_recursive.get(),
                    "shortcut_bindings": getattr(self, "_shortcut_bindings", {}),
                }, f, indent=2)
        except Exception:
            pass
''')

# 5. _load_config: def 7042, 尾 7493
splice(7042, 7493, '''        import json
        try:
            with open(self._config_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            saved_shortcuts = data.get("shortcut_bindings", {})
            if isinstance(saved_shortcuts, dict):
                for sid, bl in saved_shortcuts.items():
                    if isinstance(bl, list):
                        seqs = [str(b) for b in bl if isinstance(b, str) and b]
                        if seqs:
                            self._shortcut_bindings[str(sid)] = seqs
            self._reader_dir = data.get("pdf_folder", "")
            saved_highlight = str(data.get("reader_highlight_color", HIGHLIGHT)).lower()
            if re.fullmatch(r"#[0-9a-f]{6}", saved_highlight):
                self._reader_highlight_hex.set(saved_highlight)
                self._reader_highlight_color = tuple(
                    int(saved_highlight[i:i + 2], 16) for i in (1, 3, 5))
                if hasattr(self, "_reader_highlight_btn"):
                    self._reader_highlight_btn.config(bg=saved_highlight, fg="#111111")
            open_counts = data.get("reader_open_counts", {})
            if isinstance(open_counts, dict):
                cleaned_counts = {}
                for k, v in open_counts.items():
                    try:
                        n = int(v)
                        if n > 0:
                            cleaned_counts[str(k)] = n
                    except (TypeError, ValueError):
                        continue
                self._open_counts = cleaned_counts
            resume = data.get("reader_resume", {})
            if isinstance(resume, dict):
                self._reader_resume = {
                    str(k): dict(v) for k, v in resume.items()
                    if isinstance(v, dict)
                }
            annotations = data.get("reader_annotations", {})
            if isinstance(annotations, dict):
                cleaned = {}
                for path, items in annotations.items():
                    if not isinstance(items, list):
                        continue
                    valid_items = []
                    for it in items:
                        if not isinstance(it, dict):
                            continue
                        try:
                            page = int(it.get("page", 0))
                        except (TypeError, ValueError):
                            continue
                        valid_items.append({
                            "page": page,
                            "label": str(it.get("label", "")),
                            "note": str(it.get("note", "")),
                        })
                    if valid_items:
                        cleaned[str(path)] = valid_items
                self._reader_annotations = cleaned

            marks = data.get("reader_page_marks", {})
            if isinstance(marks, dict):
                cleaned_marks = {}
                for path, by_page in marks.items():
                    if not isinstance(by_page, dict):
                        continue
                    cleaned_pages = {}
                    for page_key, rect_list in by_page.items():
                        if not isinstance(rect_list, list):
                            continue
                        valid_rects = []
                        for it in rect_list:
                            if not isinstance(it, dict):
                                continue
                            rect = it.get("rect")
                            if (isinstance(rect, list) and len(rect) == 4 and
                                    all(isinstance(v, (int, float)) for v in rect)):
                                valid_rects.append({
                                    "rect": [float(v) for v in rect],
                                    "color": str(it.get("color", "#ffdc00")),
                                })
                        if valid_rects:
                            cleaned_pages[str(page_key)] = valid_rects
                    if cleaned_pages:
                        cleaned_marks[str(path)] = cleaned_pages
                self._reader_page_marks = cleaned_marks

            text_marks = data.get("reader_text_marks", {})
            if isinstance(text_marks, dict):
                cleaned_tm = {}
                for path, by_page in text_marks.items():
                    if not isinstance(by_page, dict):
                        continue
                    cleaned_pages = {}
                    for page_key, tm_list in by_page.items():
                        if not isinstance(tm_list, list):
                            continue
                        valid = []
                        for it in tm_list:
                            if not isinstance(it, dict):
                                continue
                            x = it.get("x"); y = it.get("y")
                            if not (isinstance(x, (int, float)) and isinstance(y, (int, float))):
                                continue
                            valid.append({
                                "x": float(x), "y": float(y),
                                "text": str(it.get("text", "")),
                                "color": str(it.get("color", "#ffdc00")),
                            })
                        if valid:
                            cleaned_pages[str(page_key)] = valid
                    if cleaned_pages:
                        cleaned_tm[str(path)] = cleaned_pages
                self._reader_text_marks = cleaned_tm

            text_hl = data.get("reader_text_hl", {})
            if isinstance(text_hl, dict):
                cleaned_hl = {}
                for path, by_page in text_hl.items():
                    if not isinstance(by_page, dict):
                        continue
                    cleaned_pages = {}
                    for page_key, hl_list in by_page.items():
                        if not isinstance(hl_list, list):
                            continue
                        valid = []
                        for it in hl_list:
                            if not isinstance(it, dict):
                                continue
                            rect = it.get("rect")
                            if not (isinstance(rect, list) and len(rect) == 4 and
                                    all(isinstance(v, (int, float)) for v in rect)):
                                continue
                            valid.append({
                                "rect": [float(v) for v in rect],
                                "text": str(it.get("text", "")),
                                "color": str(it.get("color", "#ff5050")),
                            })
                        if valid:
                            cleaned_pages[str(page_key)] = valid
                    if cleaned_pages:
                        cleaned_hl[str(path)] = cleaned_pages
                self._reader_text_hl = cleaned_hl
            # ── 阅读器设置 ──
            saved_reader_theme = data.get("reader_theme", "normal")
            if saved_reader_theme in ("normal", "dark", "green", "sepia", "black_white"):
                self._reader_theme.set(saved_reader_theme)
                theme_bg = self._reader_themes[saved_reader_theme].get("canvas_bg", CANVAS)
                if hasattr(self, "_reader_canvas"):
                    self._reader_canvas.config(bg=theme_bg)
            saved_crop = data.get("reader_crop_margins", False)
            if isinstance(saved_crop, bool):
                self._reader_crop_margins_mode = saved_crop
            saved_fit = data.get("reader_fit_width", True)
            if isinstance(saved_fit, bool):
                self._reader_fit_width = saved_fit
            saved_scroll = data.get("reader_scroll_lines", 3)
            if isinstance(saved_scroll, int) and 1 <= saved_scroll <= 20:
                self._reader_scroll_lines = saved_scroll
                if hasattr(self, "_reader_speed_combo_var"):
                    self._reader_sync_speed_combo()
            # 手动框选正文区
            saved_body = data.get("reader_body_crop_regions", {})
            if isinstance(saved_body, dict):
                self._reader_body_crop_regions = {}
                for k, v in saved_body.items():
                    if isinstance(v, (list, tuple)) and len(v) == 4:
                        try:
                            self._reader_body_crop_regions[str(k)] = tuple(float(x) for x in v)
                        except (TypeError, ValueError):
                            pass
            # Excel 搜索页签
            _esr = data.get("excel_search_search_path")
            if isinstance(_esr, str) and _esr:
                self._excel_search_search_path = _esr
            _esrp = data.get("excel_search_recent_paths")
            if isinstance(_esrp, list):
                self._excel_search_recent_paths = [p for p in _esrp if isinstance(p, str)][:10]
            if isinstance(data.get("excel_search_recursive"), bool):
                self._excel_search_recursive.set(data["excel_search_recursive"])
        except Exception:
            self._reader_dir = ""
''')

# 6. _build_ui: def 1547, 尾 1780
splice(1547, 1780, '''        # ── 主体区（Notebook分页）──
        self._body_frame = tk.Frame(self, bg=BG)
        self._body_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)

        self._notebook = ttk.Notebook(self._body_frame)
        self._notebook.pack(fill=tk.BOTH, expand=True)

        # === 页签1: 文件阅读 ===
        self._tab_reader = tk.Frame(self._notebook, bg=BG)
        self._notebook.add(self._tab_reader, text="📂 文件阅读")
        self._build_pdf_reader_tab()

        # === 页签2: PDF差异对比 ===
        self._tab_pdf_diff = tk.Frame(self._notebook, bg=BG)
        self._notebook.add(self._tab_pdf_diff, text="🔍 PDF差异对比")
        self._build_pdf_diff_tab()

        # === 页签3: Excel 搜索 ===
        self._tab_excel_search = tk.Frame(self._notebook, bg=BG)
        self._notebook.add(self._tab_excel_search, text="📊 Excel 搜索")
        self._build_excel_search_tab()

        # ── 页签文字配色 ──
        tab_colors = {
            self._tab_reader:       "#e0a83e",  # 琥珀金 — 文件阅读
            self._tab_pdf_diff:     "#e0607a",  # 玫红色 — PDF差异对比
            self._tab_excel_search: "#4cc38a",  # 翠绿色 — Excel 搜索
        }
        for tab_widget, color in tab_colors.items():
            try:
                tab_id = str(tab_widget)
                self._notebook.tk.call(
                    self._notebook._w, "tab", tab_id, "-foreground", color)
            except Exception:
                pass

        # === 配置持久化 ===
        self._shortcut_bindings = {}
        self._config_file = os.path.join(os.path.dirname(__file__), "config.json")
        self._load_config()
        # 默认选中「文件阅读」页签
        self.after_idle(lambda: self._notebook.select(self._tab_reader))

        # ── 同步已加载的配置到按钮状态 ──
        self._apply_reader_settings_to_ui()
        self._reader_refresh_bookmark_list()

        # === 应用内快捷键（翻页/缩放/聚焦搜索）===
        self._setup_keyboard_shortcuts()
''')

with io.open(PATH, "w", encoding="utf-8", newline="") as f:
    f.writelines(lines)

print("done, total lines:", len(lines))
