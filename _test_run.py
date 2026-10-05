# -*- coding: utf-8 -*-
"""启动测试：构造 ExcelPdfSync，2 秒后自动关闭，捕获异常。"""
import importlib.util, os, sys, traceback

here = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("eps", os.path.join(here, "excel_pdf_sync-260729-usefull.py"))
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

try:
    app = mod.ExcelPdfSync()
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.after(2500, app.destroy)
    app.mainloop()
    print("OK: 程序启动并正常退出")
except Exception:
    traceback.print_exc()
    print("FAILED")
    sys.exit(1)
