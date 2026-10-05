# -*- coding: utf-8 -*-
"""Step4: 引用闭包删除孤立方法（用 ast 精确获取方法行范围）。"""
import io, re, ast

PATH = r"excel_pdf_sync-260729-usefull.py"
with io.open(PATH, "r", encoding="utf-8") as f:
    src = f.read()
lines = src.splitlines(keepends=True)

tree = ast.parse(src)

# 找到 ExcelPdfSync 类
cls_node = None
for node in tree.body:
    if isinstance(node, ast.ClassDef) and node.name == "ExcelPdfSync":
        cls_node = node
        break
assert cls_node is not None, "class not found"

# 收集类内方法 (name -> (lineno, end_lineno, 源码))
methods = {}
for node in cls_node.body:
    if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
        methods[node.name] = (node.lineno, node.end_lineno)

name_set = set(methods)
print(f"总方法数: {len(methods)}")

def method_src(name):
    a, b = methods[name]
    return "".join(lines[a - 1:b])  # 0-indexed slice

def refs_of(name):
    body = method_src(name)
    refs = set(re.findall(r"self\.([A-Za-z_][A-Za-z0-9_]*)", body))
    return {r for r in refs if r in name_set}

seeds = {
    "__init__", "on_close", "_apply_app_icon", "_build_ui",
    "_build_pdf_reader_tab", "_build_pdf_diff_tab", "_build_excel_search_tab",
    "_build_reader_aux_panel", "_load_config", "_save_config",
    "_setup_keyboard_shortcuts", "_shortcut_registry", "_shortcut_apply_bindings",
    "_apply_reader_settings_to_ui",
}

live = set(seeds)
stack = list(seeds)
while stack:
    m = stack.pop()
    for r in refs_of(m):
        if r not in live:
            live.add(r)
            stack.append(r)

dead = sorted([m for m in name_set if m not in live])
print(f"live: {len(live)}, dead: {len(dead)}")

# 删除 dead 方法（按起始行号从后往前），范围向上扩展到装饰器行
def true_start(a):
    # a 是 1-indexed 的 def 行号；向上吞掉 @decorator 行
    while a > 1 and lines[a - 2].lstrip().startswith("@"):
        a -= 1
    return a

dead_ranges = sorted([(true_start(methods[m][0]), methods[m][1]) for m in dead], key=lambda x: -x[0])
for a, b in dead_ranges:
    del lines[a - 1:b]

with io.open(PATH, "w", encoding="utf-8", newline="") as f:
    f.writelines(lines)

print("done, total lines:", len(lines))
