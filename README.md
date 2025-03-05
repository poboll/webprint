# WebPrint

基于 Flask + Python + win32print 的简易网页打印服务

## 项目简介

`WebPrint` 是一个可在 Windows 环境下，通过 Flask 提供 Web 界面上传 `doc` / `docx` / `pdf` 文件，并调用本地安装的 **HP LaserJet P1007** 打印机（也可自行改成其他打印机）进行打印的简易项目。

本项目适用于以下场景：
- 办公室环境：无需专业文档管理，只想简单通过局域网网页上传并自动打印
- 自助打印：在局域网中让员工/客户直接上传文档并打印

> **注意**：只能在 **Windows** 环境下使用，因为依赖 `pywin32` 库（win32api、win32print 等）进行打印操作。

## 功能特性

1. 上传文件格式支持：`doc`, `docx`, `pdf`  
2. 自动调用指定打印机（如 `HP LaserJet P1007`）打印  
3. 使用 `Flask` 提供 Web 界面，简单易用  
4. 基本安全处理，包括文件名过滤、上传限制等  

## 运行环境

- Windows (具有安装并可正常使用的打印机，如 HP LaserJet P1007)
- Python 3.7+（测试环境为 Python 3.8 / 3.9 / 3.10 均可）
- 需要安装以下依赖：
  - `flask`
  - `pywin32` (win32api, win32print)
  - （可根据需要安装 `werkzeug` 等最新版本，一般 `pip install Flask pywin32` 即可）

## 安装与使用

1. **克隆本项目**（或者下载 ZIP 解压）
   ```bash
   git clone https://github.com/poboll/webprint.git
   cd webprint
   ```

2. **安装依赖**
   ```bash
   pip install -r requirements.txt
   ```
   如果你没有 `requirements.txt`，可自行安装：
   ```bash
   pip install flask pywin32
   ```

3. **在 Windows 系统中确保安装了你要使用的打印机驱动**  
   - 并在脚本中修改 `print_file_hp_p1007` 函数，改成你的打印机名字或打印机匹配关键字。

4. **启动服务**
   ```bash
   python app.py
   ```
   默认会在 `http://0.0.0.0:7001` 监听，你也可以在 `app.run()` 中修改端口。

5. **访问 Web 界面**  
   - 在浏览器中输入 `http://你的服务器IP:7001/` 即可访问上传页面
   - 选择要打印的 `doc/docx/pdf` 文件后提交

## 注意事项

- 如果你在调用 `ShellExecute(printto, ...)` 时报错：`pywintypes.error: (31, 'ShellExecute', '连到系统上的设备没有发挥作用。')`，通常是因为你没有给文件类型配置默认打开程序。  
  - 例如 `.pdf` 必须与你的 PDF 阅读器（Adobe 或其他）建立默认关联，`.docx` 必须和 Office Word 建立关联，否则无法执行 `printto`。

- 打印机名称匹配失败的话，可以在函数 `print_file_hp_p1007` 里把关键字改成你自己的打印机名字（如 `"HP LaserJet P1108"`），或者不做大小写匹配。

- **安全**：本项目仅做简单演示，若要用于生产环境，需进行访问权限控制、上传大小限制、文件类型更严格校验等措施。

## 文件结构

```
webprint/
  ├── app.py                 # Flask 主程序
  ├── templates/
  │    └── index.html        # 上传界面
  ├── temple/                # 临时存放上传的文件
  ├── requirements.txt       # Python依赖
  └── README.md              # 说明文档
```

## 常见问题

1. **为什么打印机不执行任务？**  
   - 确认 `win32print.SetDefaultPrinter` 设置正确的打印机  
   - 确认本机 Office / PDF 阅读器能默认打开对应文件  
   - 确认 Windows 打印服务正常启用

2. **如何换别的打印机？**  
   - 修改 `if 'HP LASERJET P1007' in prn.upper():` 那一行即可。

3. **如何提供外网访问？**  
   - 在 `app.run()` 使用 `host='0.0.0.0'` 并做好路由器端口映射或云服务器安全组开放端口即可。

---

## 三、简单原理说明

- **Flask 部分**  
  通过一个简单的表单（`index.html`）上传文件到后端 (`upload_print_file`)，后端会先做安全校验（检查文件后缀、生成安全文件名等），再调用 `print_file_hp_p1007` 打印。

- **win32print & win32api**  
  - `win32print.EnumPrinters` 用于获取本地打印机列表  
  - `win32print.SetDefaultPrinter` 设置默认打印机  
  - `win32api.ShellExecute(…, “printto”, …)` 可调用已关联应用的 “打印到” 功能，必须指定打印机名称。  
  - 前提：Windows 对应文件类型 (doc, docx, pdf) 有默认关联应用才能正常 “printto”

- **安全注意**  
  项目中通过 `secure_filename`、时间戳重命名等方式，防止恶意上传脚本并执行等安全风险。但**仍不建议公开给 Internet**，最好只在内网使用并配合身份鉴权。
