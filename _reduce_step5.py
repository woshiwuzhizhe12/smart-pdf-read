# -*- coding: utf-8 -*-
"""Step5: 删除未被引用的模块级顶层定义（死常量/死类/死函数）。"""
import io, ast

PATH = r"excel_pdf_sync-260729-usefull.py"
with io.open(PATH, "r", encoding="utf-8") as f:
    src = f.read()
lines = src.splitlines(keepends=True)
tree = ast.parse(src)

# 1. 收集所有被 Load 引用的名字
loaded = set()
for node in ast.walk(tree):
    if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Load):
        loaded.add(node.id)
    elif isinstance(node, ast.Attribute):
        pass  # 属性访问不用处理，Name 已覆盖 base 引用

# 2. 收集顶层定义（排除 ExcelPdfSync 主类本身）
top_defs = []  # (name, start_lineno, end_lineno)
for node in tree.body:
    if isinstance(node, ast.ClassDef):
        if node.name != "ExcelPdfSync":
            top_defs.append((node.name, node.lineno, node.end_lineno))
    elif isinstance(node, ast.FunctionDef):
        top_defs.append((node.name, node.lineno, node.end_lineno))
    elif isinstance(node, ast.Assign):
        for t in node.targets:
            if isinstance(t, ast.Name):
                top_defs.append((t.id, node.lineno, node.end_lineno))
    elif isinstance(node, ast.AnnAssign):
        if isinstance(node.target, ast.Name):
            top_defs.append((node.target.id, node.lineno, node.end_lineno))

# 3. 找死定义
def true_start(a):
    while a > 1 and lines[a - 2].lstrip().startswith("@"):
        a -= 1
    return a

dead = []
for name, a, b in top_defs:
    if name not in loaded:
        dead.append((name, true_start(a), b))

# 打印并删除
dead.sort(key=lambda x: -x[1])
print("顶层定义总数:", len(top_defs), "死定义数:", len(dead))
for name, a, b in sorted(dead, key=lambda x: x[1]):
    print(f"  DEAD {name}: {a}-{b}")
for name, a, b in dead:
    del lines[a - 1:b]

with io.open(PATH, "w", encoding="utf-8", newline="") as f:
    f.writelines(lines)
print("done, total lines:", len(lines))
