# Smart PDF Read

面向 Windows 的本地 PDF/Excel 阅读与检索效率工具（Tkinter 单文件程序）。

一个界面里完成「读 PDF、查内容、比对位置、联动 Excel、AI 辅助阅读」，适合需要频繁翻阅技术文档 / 数据手册 / 论文的工程师。

## 功能

- **PDF 阅读**：PyMuPDF 渲染，缩放 / 翻页 / 目录导航，多开窗口
- **全文检索**：本机 PDF 全文搜索 + 网络搜索
- **Excel 处理**：openpyxl / xlrd 读写，位号匹配比对（PDF↔Excel 位置对齐）
- **AI 阅读助手**：OpenAI 兼容接口（自填 API key），默认 `gpt-4o-mini`
- **内容/文件名搜索增强**（可选，需本机装对应工具）：
  - AnyTXT Searcher —— 内容级检索，走本机 `http://127.0.0.1:9920`
  - Everything —— 文件名检索（自动探测 `Everything64.dll`）
- **EPUB / HTML 阅读**：EbookLib / tkinterweb
- **个性化**：多主题、搜索历史、联动搜索匹配列

## 运行

```bash
# 推荐用 Python 启动器
py smart_pdf_read.py
```

首次运行会自动生成 `config.json`（保存收藏路径、主题等，已加入 `.gitignore`，不进版本库）。

## 依赖

```bash
pip install -r requirements.txt
```

> 仅支持 Windows（用到 `pywin32`、`ctypes` 加载 Everything DLL、Tkinter）。

## 可选外部工具

| 工具 | 用途 | 说明 |
|------|------|------|
| [Everything](https://www.voidtools.com/) | 文件名搜索 | 非必需，缺失时对应功能不可用 |
| [AnyTXT Searcher](https://anytxt.net/) | 内容搜索 | 非必需，需本机开启 9920 端口服务 |
| OpenAI 兼容 API key | AI 阅读助手 | 非必需，界面上自填 |

## 目录结构

```
smart_pdf_read.py   # 主程序（单文件）
requirements.txt    # pip 依赖
config.json         # 运行生成，本地配置（git 忽略）
```

## 许可

暂未指定 License，保留所有权利。
