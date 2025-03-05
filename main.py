# -*- coding: utf-8 -*-
"""
基于 Flask + Python + win32print 的简易网络打印服务示例
适用于 Windows 系统下，通过调用系统的 ShellExecute("printto") 函数，将
doc / docx / pdf 文件发送到指定打印机 (如 HP LaserJet P1007) 进行打印。

主要流程：
1. 用户在浏览器上传文件（仅支持 doc、docx、pdf）
2. 服务器后台保存到指定文件夹（并进行文件名安全过滤）
3. 调用 print_file_hp_p1007() 函数执行打印动作
4. 如果检测到指定打印机，则成功发送打印任务
5. 如果未检测到对应打印机，则返回错误提示

使用说明：
- 需要在 Windows 环境下运行，Python 版本 3.7+；
- 必须安装 pywin32: pip install pywin32
- doc/docx/pdf 文件需要在系统中有默认关联的打开程序，否则无法执行“printto”。
- 若要在局域网其他电脑访问，请将 host 设置为 '0.0.0.0' 并在路由器/防火墙中开放端口。
"""

import os
import time
import win32api
import win32print
from flask import Flask, request, render_template
from werkzeug.utils import secure_filename

# 允许上传的后缀（不带“.”）
ALLOW_EXTENSIONS = {'doc', 'docx', 'pdf'}

app = Flask(__name__)

# 上传文件保存目录（这里示例命名为 temple，可自行修改）
UPLOAD_FOLDER = 'temple'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def print_file_hp_p1007(file_path: str) -> bool:
    """
    使用 HP LaserJet P1007 打印给定的 file_path 文件。
    1. 列出所有本地打印机，匹配名称中含“HP LaserJet P1007”的打印机。
    2. 如果找到，就将它设为默认打印机。
    3. 调用 ShellExecute("printto") 执行打印动作。
    4. 若成功发送打印任务，返回 True；如果找不到打印机或调用失败，返回 False。
    """
    # 列出本地所有打印机信息
    printers_info = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
    # p[2] 是打印机名称
    printers = [p[2] for p in printers_info]

    # 在打印机列表中搜索包含“HP LaserJet P1007”的名称（不区分大小写）
    target_printer = None
    for prn in printers:
        if 'HP LASERJET P1007' in prn.upper():
            target_printer = prn
            break

    if not target_printer:
        # 未找到目标打印机，直接返回 False
        return False

    # 设置此打印机为默认打印机（通常为了保证 printto 时使用该打印机）
    win32print.SetDefaultPrinter(target_printer)

    try:
        # 使用 "printto" 动作，将文件发送到指定打印机
        # doc/docx/pdf 必须在系统中有默认关联的打开程序（如 Word / Acrobat）。
        win32api.ShellExecute(
            0,
            "printto",              # 动作改为 printto
            file_path,              # 要打印的文件
            f'"{target_printer}"',  # 目标打印机名称，带双引号可以避免空格冲突
            ".",
            0
        )
        return True
    except Exception as e:
        print(f"打印异常: {e}")
        return False

@app.route('/', methods=['GET', 'POST'])
def upload_print_file():
    """
    Flask 路由：根路径（'/'）
    - 如果是 GET 请求，直接渲染 index.html，让用户看到上传界面
    - 如果是 POST 请求，则获取上传文件，进行安全检查、保存文件并调用打印机
    """
    msg = ""
    if request.method == 'POST':
        file = request.files.get('file')
        if not file:
            msg = "没有检测到文件，请重新选择要上传打印的文件。"
            return render_template('index.html', msg=msg)

        # 用户上传的原始文件名（可能含中文）
        original_name = file.filename
        if not original_name:
            msg = "文件名为空，请重新选择文件。"
            return render_template('index.html', msg=msg)

        # Step1: 使用 secure_filename 做初步安全过滤
        #        例如 "数字图像处理考查内容及评分标准.docx" 可能会被转成 "docx"
        safe_name = secure_filename(original_name)

        # Step2: 分离文件主干和后缀，ext 含 '.' 前缀，如 '.docx'
        root, ext = os.path.splitext(safe_name)
        ext = ext.lower()  # 统一转小写，便于后续判断

        # Step3: 判断后缀是否在允许列表
        if ext.lstrip('.') not in ALLOW_EXTENSIONS:
            msg = "文件类型不允许，仅支持 doc / docx / pdf！"
            return render_template('index.html', msg=msg)

        # Step4: 构造“时间戳 + 主文件名 + 后缀”保证唯一性
        timestamp_str = time.strftime('%Y%m%d%H%M%S')
        # 再次用 secure_filename 对原文件名的“主干”部分进行过滤，防止恶意字符
        original_main_name = original_name.rsplit('.', 1)[0]  # 去掉最后一个点及后缀
        safe_original_main_name = secure_filename(original_main_name) or "file"

        final_filename = f"{timestamp_str}_{safe_original_main_name}{ext}"

        # Step5: 将文件保存到本地指定目录
        save_path = os.path.join(app.config['UPLOAD_FOLDER'], final_filename)
        file.save(save_path)

        # Step6: 调用打印机函数
        success = print_file_hp_p1007(save_path)
        if success:
            msg = "已发送打印任务，请耐心等待打印完成。"
        else:
            msg = "未检测到 [HP LaserJet P1007] 打印机或系统无法执行打印，请联系管理员。"

    # GET / POST 最后都渲染 index.html
    return render_template('index.html', msg=msg)

if __name__ == "__main__":
    """
    启动 Flask 服务
    - host='0.0.0.0' 表示外部网络也可访问（需配合端口映射或防火墙设置）
    - port=7001 端口可根据需要修改
    """
    app.run(host='0.0.0.0', port=7001)
