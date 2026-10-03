# 项目技术概览

## 背景

`mail4.py` 是一个命令行批量的面试通知发送脚本，存在三个主要限制：收件列（姓名、日期、时段等）写死在代码里、HTML 模板内嵌在代码中改起来麻烦、SMTP 配置硬编码。本项目将其改造为图形界面工具 `mail_gui.py`，核心目标是：**模板完全自定义 + 任意 Excel 列可引用 + 打包成免安装的 exe**。

## 整体架构

`mail_gui.py` 单文件实现，约 800 行，内部分为四层：

```
┌─────────────────────────────────────────────┐
│  界面层 (MailApp)                             │
│  tkinter/ttk 四个区块：收件数据 → 发件设置 →   │
│  邮件模板 → 发送控制 + 日志                    │
├─────────────────────────────────────────────┤
│  数据层        │ 模板引擎      │ 发送层        │
│  read_table() │ render_      │ connect_     │
│  xlsx/xls/csv │ template()   │ smtp()       │
│  → list[dict] │ {列名:格式}   │ build_       │
│               │              │ message()    │
└─────────────────────────────────────────────┘
        配置持久化：mail_config.json（exe 同目录）
```

### 数据层

`read_table(path)` 按扩展名分发到三个读取器，统一返回 `(headers, rows)`，其中 `rows` 是 `list[dict]`（列名 → 原始值）：

- **.xlsx/.xlsm**：openpyxl（`data_only=True` 取公式计算值，`read_only=True` 省内存）
- **.xls**：xlrd（仅它支持老格式）；日期单元格按 `ctype == XL_CELL_DATE` 用 `xldate_as_datetime` 转换，纯时间列会带上 1899 基准日期，渲染时需特判
- **.csv**：标准库 csv，编码依次尝试 utf-8-sig / gbk / utf-8

刻意**不使用 pandas**：仅为读表引入 pandas + numpy 会让 exe 膨胀数百 MB，openpyxl + xlrd 合计仅几 MB。

### 模板引擎

占位符语法 `{列名}` / `{列名:格式}`，由一个正则驱动：

```python
PLACEHOLDER_RE = re.compile(r"\{([^{}\n:]+)(?::([^{}\n]+))?\}")
```

- 值格式化在 `fmt_value()`：datetime → 默认 `%Y-%m-%d`（1899 基准日期且带时间则视为纯时间，默认 `%H:%M`）；`datetime.time` → `%H:%M`；整值浮点（Excel 常见）→ 转整数；带格式串时走 Python 的 `format()`/`strftime()`
- **未识别的占位符原样保留**。这是有意设计：正文是 HTML，`<style>` 块里的 `{font-size:12px}` 不会被误替换，只会触发发送前的校验提醒，用户可选择继续
- 发送前校验：提取模板中全部占位符与表头比对，未识别的弹窗列出，由用户决定是否继续

### 界面层

- 「可插入的列」按钮由实际表头动态生成，点击把 `{列名}` 插入最近聚焦的主题框或正文框（通过 `FocusIn` 事件记录目标控件）
- 高分屏适配：启动时 `SetProcessDpiAwareness(1)`，窗口尺寸按 `winfo_fpixels('1i') / 96` 的缩放比例计算，避免缩放屏上内容溢出
- 收件邮箱列不要求固定命名，自动匹配列名含「邮箱 / 邮件 / email / mail」的列

### 并发模型

发送在后台线程执行，与 UI 线程只通过一个 `queue.Queue` 通信：

- 工作线程只做计算和 `put(("log"|"progress"|"done", ...))`，**不触碰任何控件**
- UI 线程 `after(80, _poll_log)` 轮询队列，统一刷新日志（带红/绿着色）、进度条和完成弹窗
- 停止用 `threading.Event`，发送间隔的 sleep 拆成 0.1 秒小片以便及时响应
- 单条收件失败不中断整批，逐行 try/except 汇总到失败名单（可导出 CSV，utf-8-sig 保证 Excel 打开不乱码）

每次点「开始发送」都会**重新读取数据文件**，用户改完 Excel 无需重启。

### 配置持久化

`mail_config.json` 存于 exe 同目录（`sys.frozen` 时取 `sys.executable` 所在目录）。保存主题、模板、SMTP、间隔等全部设置；**授权码默认不保存**，勾选「记住授权码」才写入。

## 打包策略

`build_exe.bat` 用 PyInstaller `--onefile --windowed` 打包，关键点是**排除清单**：

本机 Python 环境装了大量科学计算/ML 包，openpyxl 内部有 `from PIL import ...`、`from pandas import ...` 这类可选导入（try/except 包裹），静态分析会把 PIL、pandas 拉进来，进而连锁引入 numpy、torch、transformers，首次打包 exe 达 311MB。逐层排查 xref 确认链路后，加 `--exclude-module torch/pandas/numpy/PIL/...` 清单，这些导入在运行时抛 ImportError 后被原有 try/except 吞掉，不影响功能——exe 降到 19MB。

冻结环境的功能验证用过一个临时控制台版 exe 执行核心断言（读表/渲染/构建邮件），确认可选导入缺失路径无碍后才定稿。

## 文件清单

| 文件 | 角色 |
|---|---|
| `源代码/mail_gui.py` | 主程序，包含全部四层实现 |
| `build_exe.bat` | 一键打包（含排除清单） |
| `dist/邮件群发助手.exe` | 交付物 |
| `dist/mail_config.json` | 运行时自动生成的配置 |
| `面试通知示例.xlsx` | 演示数据（含日期、时间单元格，用于验证格式化） |
| `使用说明.md` | 面向使用者的操作手册 |

## 已知限制

- .xls 依赖 xlrd 读取，建议用户另存为 .xlsx
- 占位符不支持嵌套花括号（HTML 内联样式不受影响，`<style>` 块会触发「未识别占位符」提醒）
- 列按钮每行固定排 7 个，列特别多时需拉宽窗口
- 界面语言为中文，字体优先微软雅黑
