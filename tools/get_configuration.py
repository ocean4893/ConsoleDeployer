"""
配置备份工具
传入串口对象或SSH shell对象，执行 display current-configuration 并保存到文件
"""

import os
import re
import time
from datetime import datetime
#调用方式：from tools.get_configuration import *


def clean_output(text):
    """
    清洗设备回显内容：
    1. 删除 ---- More ---- 及其前后的换行/空格/回车
    2. 将连续的空行（含仅空格的行）压缩为单个空行
    3. 去除首尾空白
    """
    # 删除 ---- More ---- 及其前后的换行/回车/空格
    # 注意：设备回显中 ---- More ---- 后面可能跟 \r\r 空格 而不是 \n
    text = re.sub(r'[\r\n\s]*---- More ----[\r\n\s]*', '\n', text)
    # 将包含空格的空白行也视为空行，连续空白行压缩为单个空行
    text = re.sub(r'(\n[ \t]*\n)[ \t]*\n', r'\1', text)
    # 再多做几轮确保完全压缩
    while re.search(r'\n[ \t]*\n[ \t]*\n', text):
        text = re.sub(r'(\n[ \t]*\n)[ \t]*\n', r'\1', text)
    # 确保 # 前面有换行（如果 # 前不是空行则补一个）
    text = re.sub(r'([^\n])\n#', r'\1\n\n#', text)
    return text.strip()

def get_current_configuration(conn, device_name='H3C'+datetime.now().strftime("%Y-%m-%d"), config_dir="log"):
    """
    从设备获取配置并保存到文件

    参数:
        conn: 连接对象，支持两种类型：
              - serial.Serial 实例（串口）
              - paramiko.Channel 实例（SSH shell）
        device_name: 设备名称
        config_dir: 保存目录，默认为 log

    返回:
        保存的文件路径，失败返回 None
    """
    # 生成保存路径: log/{device_name}/{日期}/{时分}/display current-configuration.yml
    date_str = datetime.now().strftime("%Y-%m-%d")
    time_str = datetime.now().strftime("%H-%M")
    save_dir = os.path.join(config_dir, device_name, date_str, time_str)
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    second_str = datetime.now().strftime("%H-%M-%S")
    filename = f"display current-configuration_{second_str}.yml"
    filepath = os.path.join(save_dir, filename)

    # 判断连接类型：有 reset_input_buffer 是串口，否则是 SSH shell
    is_serial = hasattr(conn, 'reset_input_buffer')

    def write_cmd(cmd):
        if is_serial:
            conn.reset_input_buffer()
            conn.write(f"{cmd}\r\n".encode())
        else:
            conn.send(f"{cmd}\n")

    def read_data(buf_size=4096):
        if is_serial:
            return conn.read(buf_size).decode('utf-8', errors='ignore')
        else:
            return conn.recv(buf_size).decode('utf-8', errors='ignore')

    try:
        # 发送 display current-configuration 命令
        write_cmd("display current-configuration")
        time.sleep(0.3)

        # 循环读取回显，直到命令执行完毕
        all_output = ""
        while True:
            chunk = read_data()
            if not chunk:
                break
            all_output += chunk
            # 遇到 ---- More ---- 分页，发送空格翻页
            if '---- More ----' in chunk:
                write_cmd(" ")
                time.sleep(0.01)
                continue
            # 如果回显末尾出现提示符（如 < 、 [ ），说明命令执行完毕
            if any(marker in chunk for marker in ['<', ']>', ']']):
                break

        # 清洗回显内容
        cleaned = clean_output(all_output)

        # 写入文件
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(f"# 设备: {device_name}\n")
            f.write(f"# 备份时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"# 命令: display current-configuration\n")
            f.write("#" + "=" * 60 + "\n\n")
            f.write(cleaned)


        return (f"✅ 配置已保存至: {filepath},from device: {device_name},---命令: display current-configuration")

    except Exception as e:
        print(f"❌ 获取配置失败: {e}")
        return None
#==========================================================================

def get_save_configuration(conn, device_name="Device"+datetime.now().strftime("%Y-%m-%d"), config_dir="log"):
    """
    从设备获取配置并保存到文件

    参数:
        conn: 连接对象，支持两种类型：
              - serial.Serial 实例（串口）
              - paramiko.Channel 实例（SSH shell）
        device_name: 设备名称
        config_dir: 保存目录，默认为 log

    返回:
        保存的文件路径，失败返回 None
    """
    # 生成保存路径: log/{device_name}/{日期}/{时分}/display save.yml
    date_str = datetime.now().strftime("%Y-%m-%d")
    time_str = datetime.now().strftime("%H-%M")
    save_dir = os.path.join(config_dir, device_name, date_str, time_str)
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    second_str = datetime.now().strftime("%H-%M-%S")
    filename = f"display save_{second_str}.yml"
    filepath = os.path.join(save_dir, filename)

    # 判断连接类型：有 reset_input_buffer 是串口，否则是 SSH shell
    is_serial = hasattr(conn, 'reset_input_buffer')
    def write_cmd(cmd):
        if is_serial:
            conn.reset_input_buffer()
            conn.write(f"{cmd}\r\n".encode())
        else:
            conn.send(f"{cmd}\n")

    def read_data(buf_size=4096):
        if is_serial:
            return conn.read(buf_size).decode('utf-8', errors='ignore')
        else:
            return conn.recv(buf_size).decode('utf-8', errors='ignore')

    try:
        # 发送 display save 命令
        write_cmd("display save")
        time.sleep(0.3)

        # 循环读取回显，直到命令执行完毕
        all_output = ""
        while True:
            chunk = read_data()
            if not chunk:
                break
            all_output += chunk
            # 遇到 ---- More ---- 分页，发送空格翻页
            if '---- More ----' in chunk:
                write_cmd(" ")
                time.sleep(0.01)
                continue
            # 如果回显末尾出现提示符（如 < 、 [ ），说明命令执行完毕
            if any(marker in chunk for marker in ['<', ']>', ']']):
                break

        # 清洗回显内容
        cleaned = clean_output(all_output)

        # 写入文件
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(f"# 设备: {device_name}\n")
            f.write(f"# 备份时间: {datetime.now().strftime('%Y-%m-%d')}\n")
            f.write(f"# 命令: display save\n")
            f.write("#" + "=" * 60 + "\n\n")
            f.write(cleaned)

        return (f"✅ 配置已保存至: {filepath},from device: {device_name},---命令: display save")

    except Exception as e:
        print(f"❌ 获取配置失败: {e}")
        return None
