# -*- coding: utf-8 -*-
"""Step3: 清理事件处理器中的已删功能引用。"""
import io

PATH = r"excel_pdf_sync-260729-usefull.py"
with io.open(PATH, "r", encoding="utf-8") as f:
    lines = f.readlines()

def find_idx(substr, start=0):
    for i in range(start, len(lines)):
        if substr in lines[i]:
            return i
    raise ValueError("anchor not found: " + substr)

def replace_method_body(def_anchor, next_anchor, new_body):
    global lines
    d = find_idx(def_anchor)
    n = find_idx(next_anchor, d + 1)
    nb = new_body.splitlines(keepends=True)
    lines[d + 1:n] = nb

def delete_lines_containing(substr):
    global lines
    lines = [l for l in lines if substr not in l]

def delete_multiline(exact_text):
    global lines
    full = "".join(lines)
    if exact_text not in full:
        raise ValueError("multiline not found:\n" + exact_text)
    full = full.replace(exact_text, "", 1)
    lines = full.splitlines(keepends=True)

# 1. 主阅读区点击：仅保留「取词 + 高亮收藏」
replace_method_body(
    "def _on_reader_canvas_click(self, event):",
    "def _reader_search_and_show",
    '''        """点击预览文字 → 高亮收藏模式时持久高亮该词。"""
        self._reader_set_active_slot(1)
        if not self._reader_pdf:
            self._reader_canvas.scan_mark(event.x, event.y)
            return
        zoom = self.PDF_DPI / 72.0 * self._reader_scale
        canvas_x = self._reader_canvas.canvasx(event.x)
        canvas_y = self._reader_canvas.canvasy(event.y)
        yoffsets = self._reader_page_yoffsets
        if not yoffsets:
            self._reader_canvas.scan_mark(event.x, event.y)
            return
        import bisect
        idx = bisect.bisect_right(yoffsets, canvas_y) - 1
        target_page = idx if 0 <= idx < self._reader_pdf.page_count else None
        if target_page is None:
            self._reader_canvas.scan_mark(event.x, event.y)
            return
        bbox = self._reader_canvas.bbox(f"page_{target_page}")
        x_offset = bbox[0] if bbox else 0
        crop_x, crop_y = (
            self._reader_page_crop_origins[target_page]
            if target_page < len(self._reader_page_crop_origins) else (0.0, 0.0))
        pdf_x = (canvas_x - x_offset) / zoom + crop_x
        pdf_y = (canvas_y - yoffsets[target_page]) / zoom + crop_y
        self._reader_page = target_page
        self._reader_update_active_page_info()
        self._reader_highlight_thumb()
        page = self._reader_pdf[target_page]
        clicked_word = None
        clicked_rect = None
        cache = self._reader_text_cache
        if cache and cache[target_page] is not None:
            for w in cache[target_page]:
                if w[0] <= pdf_x <= w[2] and w[1] <= pdf_y <= w[3]:
                    clicked_word = self._clean_word(w[4])
                    clicked_rect = [w[0], w[1], w[2], w[3]]
                    break
        else:
            for w in page.get_text("words"):
                if w[0] <= pdf_x <= w[2] and w[1] <= pdf_y <= w[3]:
                    clicked_word = self._clean_word(w[4])
                    clicked_rect = [w[0], w[1], w[2], w[3]]
                    break
        if not clicked_word or len(clicked_word) < 1:
            self._reader_canvas.scan_mark(event.x, event.y)
            return
        self._reader_last_clicked_word = clicked_word
        if self._reader_text_hl_mode and clicked_rect:
            self._reader_save_text_hl(target_page, clicked_rect, clicked_word)
''')

# 2. 辅助阅读区点击：仅取词
replace_method_body(
    "def _reader_aux_canvas_click(self, slot, event):",
    "def _reader_capture_merged_geometry",
    '''        state = self._reader_aux.get(slot)
        if not state or not state["doc"]:
            return
        self._reader_set_active_slot(slot)
        canvas = state["canvas"]
        canvas_y = canvas.canvasy(event.y)
        import bisect
        page_idx = bisect.bisect_right(state["page_yoffsets"], canvas_y) - 1
        if not 0 <= page_idx < state["doc"].page_count:
            canvas.scan_mark(event.x, event.y)
            return
        page_x = state["page_xoffsets"].get(page_idx, 0)
        crop_x, crop_y = state.get("page_crop_origins", {}).get(page_idx, (0.0, 0.0))
        pdf_x = (canvas.canvasx(event.x) - page_x) / state["zoom"] + crop_x
        pdf_y = (canvas_y - state["page_yoffsets"][page_idx]) / state["zoom"] + crop_y
        state["page"] = page_idx
        state["page_var"].set(f"{page_idx + 1} / {state['doc'].page_count}")
        self._reader_update_active_page_info()
        clicked_word = None
        words = self._reader_aux_get_words(slot)[page_idx]
        for x0, y0, x1, y1, word in words:
            if x0 <= pdf_x <= x1 and y0 <= pdf_y <= y1:
                clicked_word = word
                break
        if clicked_word:
            self._reader_last_clicked_word = clicked_word
        else:
            canvas.scan_mark(event.x, event.y)
''')

# 3. 双击：去掉圆盘菜单
replace_method_body(
    "def _reader_on_canvas_double(self, slot, event):",
    "def _reader_on_canvas_right",
    '''        self._reader_cancel_pending_click()
''')

# 4. 右键：仅保留标注/文本标记自身的右键判断，去掉圆盘菜单
replace_method_body(
    "def _reader_on_canvas_right(self, slot, event):",
    "def _reader_show_radial_menu",
    '''        self._reader_cancel_pending_click()
        canvas = self._reader_canvas if slot == 1 else self._reader_aux[slot]["canvas"]
        try:
            current = canvas.find_withtag("current")
            if current:
                tags = canvas.gettags(current[0])
                if any(t in ("reader_annotation", "reader_text_mark", "reader_text_hl") for t in tags):
                    return
        except Exception:
            pass
''')

# 5. 去掉文本/代码预览的点击取词绑定（多行）
delete_multiline('''        self._reader_text_preview.bind(
            "<Button-1>",
            lambda e: self._reader_text_click(self._reader_text_preview, e, slot=1),
            add="+")
''')
delete_multiline('''        self._reader_code_text.bind(
            "<Button-1>",
            lambda e: self._reader_text_click(self._reader_code_text, e, slot=1),
            add="+")
''')

# 6. 去掉内容搜索跨文件提示绑定（单行）
delete_lines_containing("_show_anytxt_hits_popup")

# 7. 右键菜单：去掉「阅读区 3 打开」（多行）和「添加到常用文件」（单行）
delete_multiline('''        self._reader_ctx_menu.add_command(
            label="▥ 在阅读区 3 打开",
            command=lambda: self._reader_dual_open_selected(3))
''')
delete_lines_containing("_reader_ctx_add_favorite")

with io.open(PATH, "w", encoding="utf-8", newline="") as f:
    f.writelines(lines)

print("done, total lines:", len(lines))
