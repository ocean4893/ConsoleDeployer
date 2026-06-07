"""
设备信息获取工具
传入一个 paramiko Channel（shell）对象，自动发送命令获取设备信息，
返回包含设备各项信息的字典。

调用方式：from tools.get_device_info import get_device_info
"""

import os
import re
import time
from datetime import datetime


def _send_command(shell, cmd, wait=0.3):
    """发送命令并接收回显"""
    shell.send(f"{cmd}\n")
    time.sleep(wait)
    result = shell.recv(65535).decode("utf-8", errors="ignore")
    return result


def _clean_output(text):
    """清洗回显：去除 More 分页、退格符、控制字符"""
    # 去除 ---- More ---- 及其前后空白
    text = re.sub(r'[\r\n\s]*---- More ----[\r\n\s]*', '\n', text)
    # 去除退格符及其前面的字符（如果有）
    text = re.sub(r'.[\b]', '', text)
    # 去除 \r
    text = text.replace('\r', '')
    return text.strip()


def _save_info_to_file(info):
    """将设备信息保存到 log/{设备名}/{日期}/info.md"""
    device_name = info.get("device_name", "UnknownDevice")
    # 清理设备名中的非法文件名字符
    device_name = re.sub(r'[<>:"/\\|?*]', '_', device_name)
    today = datetime.now().strftime("%Y-%m-%d")
    log_dir = os.path.join("log", device_name, today)
    os.makedirs(log_dir, exist_ok=True)
    filepath = os.path.join(log_dir, "info.md")

    lines = [
        f"# 设备信息 - {device_name}",
        f"",
        f"| 字段 | 值 |",
        f"|------|-----|",
    ]
    for key in ["device_name", "model", "software_version", "hardware_version",
                 "serial_number", "mac_address", "uptime", "cpu_usage",
                 "memory_usage"]:
        val = info.get(key, "N/A")
        lines.append(f"| {key} | {val} |")

    # 设备描述单独放，内容可能较长
    if "description" in info:
        lines.append("")
        lines.append("## 设备描述")
        lines.append("```")
        lines.append(info["description"])
        lines.append("```")

    content = "\n".join(lines)
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(content)
        print(f"  📁 设备信息已保存至: {filepath}")
    except Exception as e:
        print(f"[ERROR] 保存设备信息文件失败: {e}")


def get_device_info(shell, version_wait=1.0, device_wait=0.5,printflag=1):
    """
    从设备获取基本信息，返回字典。

    参数:
        shell: paramiko.Channel 对象（ssh_client.invoke_shell() 的返回值）
        version_wait: display version 命令的等待时间（秒），默认 1.0
        device_wait: 其他命令的等待时间（秒），默认 0.5

    返回:
        dict，包含以下键（可能缺失部分字段）：
            - device_name: 设备名称（sysname）
            - model: 设备型号
            - software_version: 软件版本
            - hardware_version: 硬件版本
            - serial_number: 序列号（SN）
            - mac_address: MAC 地址
            - uptime: 运行时间
            - cpu_usage: CPU 利用率（最近5秒）
            - memory_usage: 内存利用率
            - description: 设备描述信息

    用法示例:
        ssh_client = paramiko.SSHClient()
        ssh_client.connect(hostname=ip, port=22, username=user, password=pwd)
        shell = ssh_client.invoke_shell()
        time.sleep(0.5)
        shell.recv(65535)  # 清空缓冲区

        info = get_device_info(shell)
        print(info["device_name"], info["serial_number"])
    """
    info = {}
    print(f"\n🔍 正在获取设备信息...")
    # ==================== 1. 获取设备名称 ====================
    try:
        result = _send_command(shell, "display current-configuration | include sysname", device_wait)
        # 排除回显中命令本身的 "sysname" 字样，只匹配配置行
        # 在回显中，配置行通常是 "sysname SwitchHCL-3" 这样的格式
        lines = result.split("\n")
        device_name = None
        for line in lines:
            # 跳过包含命令本身的提示行（如 <Switch>display ... | include sysname）
            if "include sysname" in line or "| include" in line:
                continue
            m = re.search(r'sysname\s+(\S+)', line)
            if m:
                device_name = m.group(1)
                break
        if device_name:
            info["device_name"] = device_name
        else:
            info["device_name"] = f"[未匹配到sysname] 原始回显: {result}"
    except Exception as e:
        print(f"[ERROR] 获取设备名称失败: {e}")
        input("按回车键继续...")

    # ==================== 2. 获取版本信息 ====================
    try:
        result = _send_command(shell, "display version", version_wait)
        cleaned = _clean_output(result)

        # H3C 版本格式: "H3C S6850 uptime is ..." 或 "H3C Comware Software, Version ..."
        # 优先从 "H3C <型号> uptime" 格式提取型号
        m = re.search(r'H3C\s+(\S+?)\s+uptime', cleaned, re.IGNORECASE)
        if m:
            info["model"] = m.group(1)
        else:
            # 回退到 "H3C Comware Software" 格式
            m = re.search(r'H3C\s+([\w-]+)\s+[Ss]oftware', cleaned)
            if m:
                info["model"] = m.group(1)
            else:
                info["model"] = f"[未匹配到型号] 原始回显: {cleaned[:200]}"

        # 软件版本
        m = re.search(r'[Ss]oftware,\s*(.+?)(?:,|$)', cleaned)
        if m:
            info["software_version"] = m.group(1).strip()
        else:
            info["software_version"] = f"[未匹配到软件版本] 原始回显: {cleaned[:200]}"

        # 硬件版本
        m = re.search(r'[Hh]ardware\s+[Vv]ersion[：:\s]*(\S+)', cleaned)
        if m:
            info["hardware_version"] = m.group(1)
        else:
            info["hardware_version"] = f"[未匹配到硬件版本] 原始回显: {cleaned[:200]}"

        # 序列号 SN
        m = re.search(r'(?:SN\s*[：:]\s*|Serial Number\s*[：:]\s*)(\S+)', cleaned)
        if m:
            info["serial_number"] = m.group(1)
        else:
            info["serial_number"] = f"[未匹配到序列号] 原始回显: {cleaned[:200]}"

        # 运行时间 uptime
        m = re.search(r'(?:uptime|Up time|运行时间)[：:\s]*(.+?)(?:\n|$)', cleaned, re.IGNORECASE)
        if m:
            info["uptime"] = m.group(1).strip()
        else:
            info["uptime"] = f"[未匹配到运行时间] 原始回显: {cleaned[:200]}"

        # MAC 地址
        m = re.search(r'[Mm][Aa][Cc]\s*[Aa]ddress\s*[：:\s]*(\S+)', cleaned)
        if m:
            info["mac_address"] = m.group(1)
        else:
            info["mac_address"] = f"[未匹配到MAC地址] 原始回显: {cleaned[:200]}"
    except Exception as e:
        print(f"[ERROR] 获取版本信息失败: {e}")
        input("按回车键继续...")

    # ==================== 3. 获取设备信息（含序列号、MAC） ====================
    try:
        result = _send_command(shell, "display device manuinfo", device_wait)
        cleaned = _clean_output(result)

        # 如果上一步没获取到序列号，从这里再尝试
        if "serial_number" not in info:
            m = re.search(r'(?:SN\s*[：:]\s*|Serial Number\s*[：:]\s*)(\S+)', cleaned)
            if m:
                info["serial_number"] = m.group(1)
            else:
                info["serial_number"] = f"[未匹配到序列号] 原始回显: {cleaned[:200]}"

        if "mac_address" not in info:
            m = re.search(r'[Mm][Aa][Cc]\s*[Aa]ddress\s*[：:\s]*(\S+)', cleaned)
            if m:
                info["mac_address"] = m.group(1)
            else:
                info["mac_address"] = f"[未匹配到MAC地址] 原始回显: {cleaned[:200]}"
    except Exception as e:
        print(f"[ERROR] 获取设备详细信息失败: {e}")
        input("按回车键继续...")

    # ==================== 4. 获取 CPU 和内存信息 ====================
    try:
        result = _send_command(shell, "display cpu-usage", device_wait)
        cleaned = _clean_output(result)
        # CPU 利用率: "5% in last 5 seconds"
        m = re.search(r'(\d+)%\s*in\s+last\s+5\s+seconds', cleaned)
        if m:
            info["cpu_usage"] = f"{m.group(1)}%"
        else:
            info["cpu_usage"] = f"[未匹配到CPU使用率] 原始回显: {cleaned[:200]}"
    except Exception as e:
        print(f"[ERROR] 获取CPU使用率失败: {e}")
        input("按回车键继续...")

    try:
        result = _send_command(shell, "display memory", device_wait)
        cleaned = _clean_output(result)
        # 内存利用率: "Memory usage: 23.5%"
        m = re.search(r'[Mm]emory\s+[Uu]sage[：:\s]*([\d.]+)%', cleaned)
        if m:
            info["memory_usage"] = f"{m.group(1)}%"
        else:
            info["memory_usage"] = f"[未匹配到内存使用率] 原始回显: {cleaned[:200]}"
    except Exception as e:
        print(f"[ERROR] 获取内存使用率失败: {e}")
        input("按回车键继续...")

    # ==================== 5. 获取设备描述 ====================
    try:
        result = _send_command(shell, "display device", device_wait)
        cleaned = _clean_output(result)
        # 提取设备描述部分
        info["description"] = cleaned[:500]  # 截取前500字符
    except Exception as e:
        print(f"[ERROR] 获取设备描述失败: {e}")
        input("按回车键继续...")
    # ==================== 6. 保存信息到文件 ====================
    _save_info_to_file(info)

    if printflag:
        print("\n📋 设备核心信息:")
        for key in info:
            print(f"  {key}: {info.get(key, 'N/A')}")
        input("\n按回车继续...")

    return info

