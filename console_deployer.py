"""
Console 串口自动化配置下发工具
===============================
功能：通过 Console 串口直连设备，实现自动化配置下发。
流程：监听连接 → 识别设备名 → 匹配 Excel → 配置下发 → 返回监听（循环）

依赖：pyserial, pandas, openpyxl
"""

import serial
import serial.tools.list_ports
import time
import re
import pandas as pd
import os
from datetime import datetime
from tools.Toolkit import col_name_to_index
from tools.get_configuration import get_current_configuration, get_save_configuration



# ============================================================
# 调试开关
# ============================================================
PRINT_DEBUG = 1           # 是否打印调试信息（1=开启，0=关闭）
SAVE_FLAG_DEBUG = 0       # 仅供调试：0=跳过保存检测直接等待断开，1=正常保存检测

# ============================================================
# 串口配置
# ============================================================
COM_PORT = 'COM7'           # 串口号
BAUDRATE = 9600             # 波特率
CONNECT_TIMEOUT = 30        # 连接检测超时（分钟）
COMMAND_DELAY = 0.3         # 每条命令下发后的等待秒数
PROBE_INTERVAL = 0.3        # 回车探测间隔（秒）



# ============================================================
# 日志相关
# ============================================================
LOG_DIR = "log"
log_timestamp = ""


def debug_print(message):
    """调试输出，由 PRINT_DEBUG 开关控制"""
    if PRINT_DEBUG:
        print(f"[DEBUG] {message}")


# =====================================================================
# 输入层
# =====================================================================
def input_layer():
    """
    输入层：获取用户输入的 Excel 文件路径、工作模式、起始行号、配置确认开关。

    两种工作模式:
      1) 匹配模式 — 以 MATCH_COLUMN 列（代码顶部配置）为匹配键查找 Excel 行，无需起始行号
      2) 顺序模式 — 纯按 Excel 行号顺序逐行读取，需要输入起始行号

    返回: (excel_path, mode, start_row, print_config_flag)
          mode: 'match' 或 'sequence'
          start_row: 顺序模式时有效，匹配模式为 None
    """
    print("=" * 50)
    print("       Console 串口自动化配置下发工具")
    print("=" * 50)

    # 1. Excel 文件路径
    while True:
        excel_path = input("请拖入 Excel 文件：").strip()
        excel_path = re.sub(r"[&\"]", "", excel_path).strip()
        if not excel_path:
            print("❌ 路径不能为空，请重新输入")
            continue
        if not os.path.isfile(excel_path):
            print(f"❌ 文件不存在: {excel_path}")
            continue
        break

    # 2. 选择工作模式
    while True:
        mode_input = input("选择工作模式（1=匹配模式/2=顺序模式）：").strip()
        if mode_input == '1':
            mode = 'match'
            break
        elif mode_input == '2':
            mode = 'sequence'
            break
        else:
            print("❌ 请输入 1（匹配模式）或 2（顺序模式）")

    start_row = None

    if mode == 'sequence':
        # 顺序模式：需要起始行号
        while True:
            start_row_input = input("数据起始行号：").strip()
            if not start_row_input.isdigit() or int(start_row_input) < 1:
                print("❌ 请输入有效的正整数！")
                continue
            start_row = int(start_row_input)
            break

    # 3. 配置确认开关
    while True:
        flag_input = input("下发前确认配置（1是/0否）：").strip()
        if flag_input not in ('0', '1'):
            print("❌ 请输入 1（是）或 0（否）")
            continue
        print_config_flag = int(flag_input)
        break

    return excel_path, mode, start_row, print_config_flag


# =====================================================================
# 计算层
# =====================================================================
def calculation_layer(excel_path, mode, start_row=None):
    """
    计算层：读取 Excel，根据工作模式逐行解析，为每台设备生成配置模板列表。

    两种模式:
      - match（匹配模式）: 读取整个 Excel，以 E 列（设备名）建立索引。
        后续下发时根据实际连接的设备名查找对应配置。不需要起始行号。
      - sequence（顺序模式）: 从 start_row 开始逐行读取，按行号顺序生成配置列表。

    参数:
        excel_path:    Excel 文件路径
        mode:          'match' 或 'sequence'
        start_row:     顺序模式时有效，数据起始行号（1-based）

    返回:
        list[dict] — 设备配置列表，每个元素:
            {
                "device_name": str,
                "row": int,
                "raw_data": dict,
                "templates": list[list[str]]
            }
    """
    print("\n[计算层]")
    print(f"📂 正在读取 Excel: {excel_path}")
    print(f"📋 工作模式: {'匹配模式' if mode == 'match' else '顺序模式'}")

#+++++++++++++++++++++++配置定义区域+++++++++++++++++++++++++++++++++++++++++++
#+++++++++++++++++++++++配置定义区域+++++++++++++++++++++++++++++++++++++++++++
    DEVICE_NAME_COL = 'e'
    CABLE_VLAN_COL = 'f'
    more_cols = ['g', 'h', 'i', 'j']  # 在这里定义更多设备属性吧！！
    COMPLETE_MARKER_COL = 'g'  # 完成标记列（下发完成后写入 TRUE）
#+++++++++++++++++++++++配置定义区域+++++++++++++++++++++++++++++++++++++++++++
#+++++++++++++++++++++++配置定义区域+++++++++++++++++++++++++++++++++++++++++++

    if mode == 'sequence':
        # ---- 顺序模式：从指定行开始逐行读取 ----
        df = pd.read_excel(excel_path, header=None, skiprows=start_row - 1, sheet_name=0)
        print(f"📋 共读取到 {len(df)} 行数据（从第 {start_row} 行开始）")

        device_configs = []
        for i in range(len(df)):
            try:
                device_name = df.iloc[i, col_name_to_index(DEVICE_NAME_COL)]
                if pd.isna(device_name):
                    continue
                device_name = str(device_name).strip()
                if not device_name:
                    continue

                raw_data = {
                    "device_name": device_name,
                    "cable_vlan": df.iloc[i, col_name_to_index(CABLE_VLAN_COL)],
                    "deploy_status": df.iloc[i, col_name_to_index(COMPLETE_MARKER_COL)],
                }

                templates = generate_templates(raw_data)

                device_config = {
                    "device_name": device_name,
                    "row": start_row + i,
                    "excel_path": excel_path,
                    "raw_data": raw_data,
                    "templates": templates,
                    'deploy_status': raw_data["deploy_status"],
                    'complete_marker_col': COMPLETE_MARKER_COL
                }
                device_configs.append(device_config)
                print(f"  设备 {len(device_configs)}: {device_name} → {len(templates)} 个模板")

            except Exception as e:
                print(f"  ⚠️ 第 {start_row + i} 行数据读取失败: {e}")
                continue

    else:
        # ---- 匹配模式：读取整个 Excel，以 DEVICE_NAME_COL 列为设备名建立索引 ----
        df = pd.read_excel(excel_path, header=None, sheet_name=0)
        name_col_idx = col_name_to_index(DEVICE_NAME_COL)
        print(f"📋 共扫描 {len(df)} 行数据，匹配列: {DEVICE_NAME_COL.upper()}")

        device_configs = []
        for i in range(len(df)):
            try:
                device_name = df.iloc[i, name_col_idx]
                if pd.isna(device_name):
                    continue
                device_name = str(device_name).strip()
                if not device_name:
                    continue

                raw_data = {
                    "device_name": device_name,
                    "cable_vlan": df.iloc[i, col_name_to_index(CABLE_VLAN_COL)],
                    "deploy_status": df.iloc[i, col_name_to_index(COMPLETE_MARKER_COL)],
                }

                templates = generate_templates(raw_data)

                device_config = {
                    "device_name": device_name,
                    "row": i + 1,  # 1-based 行号
                    "excel_path": excel_path,
                    "raw_data": raw_data,
                    'deploy_status': raw_data["deploy_status"],
                    "templates": templates,
                    'complete_marker_col': COMPLETE_MARKER_COL
                }
                device_configs.append(device_config)
                print(f"  设备 {len(device_configs)}: {device_name}（Excel 第 {i+1} 行）→ {len(templates)} 个模板")

            except Exception as e:
                print(f"  ⚠️ 第 {i+1} 行数据读取失败: {e}")
                continue

    print(f"✅ 共加载 {len(device_configs)} 台设备配置\n")
    return device_configs


def generate_templates(device_data):
    """
    根据设备数据生成配置模板列表。
    每个模板最后一个元素为保存命令的完成检测标记，倒数第二个为保存命令本身。

    参数:
        device_data: dict，包含设备原始数据

    返回:
        list[list[str]] — 模板列表
    """
    cable_vlan = device_data["cable_vlan"]
    # 处理可能的 NaN
    try:
        vlan_id = int(float(cable_vlan))
    except (ValueError, TypeError):
        vlan_id = ""

    templates = [
        # 模板 1：配置 VLAN 和接口
        [
            'sys',
            f'vlan {vlan_id}',
            'quit',
            'save force',
            'The current configuration is saved to the active main board successfully',#串口中，似乎无法强制断开来重启，所以只能通过检测保存完成的提示来确认保存成功
        ],
        [
            'sys',
            f'vlan {vlan_id+1}',
            '#model_2',
            'quit',
            'save force',
            'Validating file. Please wait...',
        ],
    ]

    return templates


# =====================================================================
# 下发层
# =====================================================================
def deploy_layer(device_configs, print_config_flag):
    """
    下发层：核心循环，管理串口生命周期和配置下发。
    全程单线程，循环：监听 → 匹配 → 下发 → 返回监听
    """
    print("[下发层 - 监听模式]")

    while True:
        # ---- 阶段 1: 监听连接 + 提取设备名 ----
        serial_conn, device_name = wait_for_device(COM_PORT, BAUDRATE, CONNECT_TIMEOUT)
        if serial_conn is None or device_name is None:
            continue  # 超时后继续等待

        print(f"\n[设备识别] 📋 提取到设备名: {device_name}")

        # ---- 阶段 2: 匹配 Excel ----
        matched_device = match_device_in_excel(device_name, device_configs)
        if matched_device is None:
            print(f"⚠️ 设备 {device_name} 未在 Excel 中找到匹配项，跳过下发")
            write_log(f"未匹配设备: {device_name}")
            close_serial(serial_conn)
            continue

        # ---- 阶段 2.5: 检查完成标记 ----
        if not check_status_col(matched_device):
            # 用户选择跳过，直接进入等待断开
            print("⏳ 请断开当前设备的 Console 线...")
            print("🔍 监听串口数据，等待断开信号...")
            wait_for_disconnect(serial_conn)
            close_serial(serial_conn)
            print("🔄 连接断开，返回监听状态...")
            continue

        # ---- 阶段 3: 配置下发 ----
        print(f"\n📡 开始为 {device_name} 下发配置...")
        deploy_to_device(serial_conn, matched_device, print_config_flag)

        # ---- 阶段 4: 完成，等待断开 ----
        print(f"\n✅ {device_name} 配置下发完成")
        write_complete(device_name)
        write_status_col(matched_device)  
        print("⏳ 请断开当前设备的 Console 线...")
        print("🔍 监听串口数据，等待断开信号...")

        # 等待串口断开（工程师拔线后串口自动不可用）
        wait_for_disconnect(serial_conn)
        close_serial(serial_conn)
        print("🔄 连接断开，返回监听状态...")


# ------------------------------------------------------------------
# 2.3.1 监听连接
# ------------------------------------------------------------------
def wait_for_device(com_port, baudrate, timeout_minutes):
    """
    监听指定 COM 口，持续发送回车直到检测到设备提示符。

    检测模式：正则 r'<([^>]+)>' 匹配到即认为设备已就绪，
    同时从尖括号中提取设备名（如 <H3C> → H3C, <SW1> → SW1）。

    返回:
        (serial.Serial, str) — (串口对象, 设备名)，超时返回 (None, None)
    """
    print(f"\n🔍 正在监听 {com_port}，等待设备连接...")
    start_time = time.time()
    timeout_seconds = timeout_minutes * 60

    # 尝试打开串口
    serial_conn = None
    while serial_conn is None:
        try:
            serial_conn = serial.Serial(
                port=com_port,
                baudrate=baudrate,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=1,
            )
            print(f"✅ 串口 {com_port} 打开成功")
        except serial.SerialException as e:
            elapsed = time.time() - start_time
            if elapsed > timeout_seconds:
                print(f"⏱️ 等待超时（{timeout_minutes}分钟），继续监听...")
                start_time = time.time()  # 重置计时器
            else:
                print(f"⏳ 串口 {com_port} 无法打开 ({e})，{PROBE_INTERVAL}秒后重试...")
                time.sleep(PROBE_INTERVAL)

    # 循环发送回车，检测设备提示符
    # 支持两种提示符格式: <H3C> 或 [SW1]
    prompt_pattern = re.compile(r'[<\[][A-Za-z0-9_-]+[>\]]')
    name_extract_pattern = re.compile(r'[<\[]([A-Za-z0-9_-]+)[>\]]')
    # 累积缓冲区，防止提示符被分多次读取截断
    accumulated = ""
    send_count = 0
    while True:
        try:
            # 发送回车
            send_count += 1

            # 写前状态
            in_waiting_before = serial_conn.in_waiting
            out_waiting_before = serial_conn.out_waiting

            serial_conn.write(b'\r\n')

            # 写后状态
            out_waiting_after = serial_conn.out_waiting

            time.sleep(PROBE_INTERVAL)

            # 读取串口缓冲区所有可用数据
            total_bytes = 0
            chunks_info = []
            while serial_conn.in_waiting:
                chunk = serial_conn.read(serial_conn.in_waiting)
                total_bytes += len(chunk)
                text = chunk.decode('utf-8', errors='ignore')
                accumulated += text
                chunks_info.append(f"{len(chunk)}B:{repr(text[:100])}")

            # 累积缓冲区长度
            acc_len = len(accumulated)
            in_waiting_after = serial_conn.in_waiting

            # 打印完整探测信息
            debug_print(
                f"⏎ 探测#{send_count} | "
                f"写前in={in_waiting_before} 写后out={out_waiting_after} | "
                f"in_waiting={in_waiting_after} | "
                f"累积={acc_len}字符 | "
                f"数据={'✅' if chunks_info else '❌'} "
                f"{'|'.join(chunks_info) if chunks_info else '(无)'}"
            )

            # 在累积数据中搜索提示符
            match = prompt_pattern.search(accumulated)
            if match:
                prompt_str = match.group()
                name_match = name_extract_pattern.search(prompt_str)
                device_name = name_match.group(1).strip() if name_match else prompt_str.strip('<>[]')
                print(f"✅ 检测到设备提示符: {prompt_str} → 设备名: {device_name}")
                debug_print(f"累积缓冲区总长度: {acc_len} 字符")
                return serial_conn, device_name

            # 超时检测
            elapsed = time.time() - start_time
            if elapsed > timeout_seconds:
                print(f"⏱️ 连接检测超时（{timeout_minutes}分钟），重新监听...")
                close_serial(serial_conn)
                return None, None

        except serial.SerialException as e:
            print(f"❌ 串口通信异常: {e}")
            close_serial(serial_conn)
            return None, None
        except Exception as e:
            print(f"❌ 未知异常: {e}")
            close_serial(serial_conn)
            return None, None


# ------------------------------------------------------------------
# 2.3.2 匹配 Excel
# ------------------------------------------------------------------
def match_device_in_excel(device_name, device_configs):
    """
    在设备配置列表中按设备名匹配。
    支持精确匹配和前缀匹配（如设备提示符 [SW1-vlan101] 可匹配 Excel 中的 SW1）。
    优先精确匹配，失败后尝试前缀匹配。

    返回:
        dict — 匹配到的设备配置，未匹配返回 None
    """
    print("\n[匹配 Excel]")
    print(f"🔎 在 Excel 中查找 {device_name}...")

    # 1. 精确匹配
    for config in device_configs:
        if config["device_name"] == device_name:
            print(f"✅ 精确匹配成功！")
            return config

    # 2. 前缀匹配：设备名以 Excel 中的名称为前缀（如 SW1-vlan101 → SW1）
    for config in device_configs:
        excel_name = config["device_name"]
        if device_name.startswith(excel_name):
            print(f"✅ 前缀匹配成功！({device_name} → {excel_name})")
            return config

    print(f"❌ 未找到匹配项")
    return None


# ------------------------------------------------------------------
# 2.3.3 配置下发
# ------------------------------------------------------------------
def deploy_to_device(serial_conn, device, print_config_flag):
    """
    向设备下发配置。

    流程：
    1. 备份当前配置和保存配置
    2. 逐模板执行命令
    3. 保存并检测完成
    """
    device_name = device["device_name"]
    templates = device["templates"]
    total_templates = len(templates)

    print(f"\n[配置下发] 设备: {device_name}, 共 {total_templates} 个模板")

    # 如果开启了配置确认，先打印配置
    if print_config_flag == 1:
        print_config_preview(device)

    # 备份配置（可选）
    
    if print_config_flag == 1:
        try:
            back_up = input("是否备份当前配置？（1是/0否）：").strip()
            if back_up == '1':
                print("📦 正在备份当前配置...")
                result = get_current_configuration(serial_conn, device_name, config_dir=LOG_DIR)
                print(f"   {result}")
                print("📦 正在备份保存配置...")
                result = get_save_configuration(serial_conn, device_name, config_dir=LOG_DIR)
                print(f"   {result}")
        except Exception as e:
            print(f"   ⚠️ 备份配置失败: {e}")
    else:
        try:
            print("📦 正在备份当前配置...")
            result = get_current_configuration(serial_conn, device_name, config_dir=LOG_DIR)
            print(f"   {result}")
            print("📦 正在备份保存配置...")
            result = get_save_configuration(serial_conn, device_name, config_dir=LOG_DIR)
            print(f"   {result}")
        except Exception as e:
            print(f"   ⚠️ 备份配置失败: {e}")

    # 逐模板下发
    for idx, command_book in enumerate(templates, start=1):
        print(f"\n  ── 模板 {idx}/{total_templates} ──")
        success = execute_one_template(serial_conn, device_name, command_book, idx)
        if not success:
            print(f"  ⚠️ 模板 {idx} 执行异常，继续下一个模板...")
            write_failed(device_name, f"模板 {idx} 执行异常")

    print(f"\n  ✅ {device_name} 配置完成")


def execute_one_template(serial_conn, device_name, command_book, template_index):
    """
    执行单个模板的所有命令。

    模板结构: [...配置命令..., 保存命令, 完成检测标记]
    """
    # 配置命令（除最后两个元素）
    config_cmds = command_book[:-2]

    for cmd in config_cmds:
        try:
            serial_conn.reset_input_buffer()
            debug_print(f"➡️ 发送命令: {cmd}")
            serial_conn.write(f"{cmd}\r\n".encode('utf-8'))
            time.sleep(COMMAND_DELAY)

            # 读取串口缓冲区所有可用数据
            output = ""
            while serial_conn.in_waiting:
                chunk = serial_conn.read(serial_conn.in_waiting)
                text = chunk.decode('utf-8', errors='ignore')
                output += text
                debug_print(f"📥 回显 ({len(chunk)} 字节): {repr(text[:200])}")

            # 打印执行结果
            result_preview = output.strip()[:100] if output.strip() else "(无回显)"
            print(f"    ✅ {cmd}")
            write_log(f"[{device_name}] [模板{template_index}] 执行: {cmd}")
            write_log(f"[{device_name}] [模板{template_index}] 返回: {result_preview}")

        except Exception as e:
            print(f"    ❌ {cmd} — 失败: {e}")
            write_log(f"[{device_name}] [模板{template_index}] 命令失败: {cmd}, 错误: {e}")
            return False

    # 执行保存命令
    if SAVE_FLAG_DEBUG:
        save_cmd = command_book[-2]
        save_marker = command_book[-1]
        try:
            print(f"    💾 保存中...")
            debug_print(f"➡️ 发送保存命令: {save_cmd}")
            serial_conn.reset_input_buffer()
            serial_conn.write(f"{save_cmd}\r\n".encode('utf-8'))
            time.sleep(1)

            # 循环检测保存完成标记
            while True:
                output = ""
                while serial_conn.in_waiting:
                    chunk = serial_conn.read(serial_conn.in_waiting)
                    text = chunk.decode('utf-8', errors='ignore')
                    output += text
                    debug_print(f"📥 保存回显 ({len(chunk)} 字节): {repr(text[:200])}")

                if save_marker in output:
                    print(f"    ✅ 保存完成")
                    write_log(f"[{device_name}] [模板{template_index}] 保存成功")
                    break

                if re.search(r'[<\[][A-Za-z0-9_-]+[>\]]', output):
                    # 出现提示符但未检测到标记，可能保存已完成
                    print(f"    ✅ 保存完成（检测到提示符）")
                    write_log(f"[{device_name}] [模板{template_index}] 保存成功（提示符确认）")
                    break

                time.sleep(0.5)

        except Exception as e:
            print(f"    ❌ 保存失败: {e}")
            write_log(f"[{device_name}] [模板{template_index}] 保存失败: {e}")
            return False

    return True


def print_config_preview(device):
    """打印配置预览并请求用户确认"""
    print(f"\n{'=' * 50}")
    print("                   配置预览")
    print(f"{'=' * 50}")
    print(f"设备名: {device['device_name']}")
    print(f"Excel 行号: {device['row']}")
    print(f"原始数据: {device['raw_data']}")
    print(f"\n即将下发的配置:")
    for i, cmds in enumerate(device["templates"], start=1):
        print(f"\n  --- 模板 {i} ---")
        for cmd in cmds:
            print(f"    {cmd}")
    print(f"\n{'=' * 50}")

    while True:
        confirm = input("确认无误后按回车继续，输入 quit 退出: ").strip()
        if confirm.lower() == 'quit':
            exit()
        break


def deploy_layer_sequence(device_configs, print_config_flag):
    """
    顺序模式下发层：按 Excel 行号顺序依次为每台设备下发配置。
    每台设备下发前等待工程师连接并识别设备名，然后匹配配置并下发。
    适合设备连接顺序与 Excel 顺序一致的场景。
    """
    print("[下发层 - 顺序模式]")
    total = len(device_configs)
    print(f"📋 共 {total} 台设备待下发")

    for idx, device in enumerate(device_configs, start=1):
        device_name = device["device_name"]
        print(f"\n{'='*50}")
        print(f"  第 {idx}/{total} 台: {device_name}")
        print(f"{'='*50}")

        # ---- 等待设备连接 ----
        print(f"\n🔍 等待 {device_name} 连接...（请连接该设备的 Console 线）")
        serial_conn, detected_name = wait_for_device(COM_PORT, BAUDRATE, CONNECT_TIMEOUT)
        if serial_conn is None:
            print("⚠️ 连接设备失败，跳至下一台...")
            write_failed(device_name, "连接失败")
            continue

        print(f"\n[设备识别] 📋 检测到设备: {detected_name}")

        # ---- 检查完成标记 ----
        if not check_status_col(device):
            # 用户选择跳过，直接进入等待断开
            print("⏳ 请断开当前设备的 Console 线...")
            print("🔍 监听串口数据，等待断开信号...")
            wait_for_disconnect(serial_conn)
            close_serial(serial_conn)
            print("🔄 连接断开，准备下一台设备...")
            continue

        # ---- 配置下发 ----
        print(f"\n📡 开始为 {detected_name} 下发配置...")
        deploy_to_device(serial_conn, device, print_config_flag)

        # ---- 完成，等待断开 ----
        print(f"\n✅ {detected_name} 配置下发完成")
        write_complete(detected_name)
        write_status_col(device) #写入完成标记
        print("⏳ 请断开当前设备的 Console 线...")
        print("🔍 监听串口数据，等待断开信号...")

        wait_for_disconnect(serial_conn)
        close_serial(serial_conn)
        print("🔄 连接断开，准备下一台设备...")

    print(f"\n{'='*50}")
    print(f"  🎉 所有 {total} 台设备配置下发完成！")
    print(f"{'='*50}")


# ------------------------------------------------------------------
# 2.3.4 等待断开 + 关闭串口
# ------------------------------------------------------------------
def wait_for_disconnect(serial_conn):
    """
    等待工程师物理断开 Console 线。

    检测策略:
      1. 硬件信号检测 — DCD/CTS/DSR 任一丢失即判定断开（毫秒级）
      2. write() + in_waiting 联合检测 — 发 \\n 探测，
         连续多次无回显 + write() 无异常时，用 read(1) 确认断开
      3. 串口异常捕获 — SerialException / OSError

    判定断开条件（满足任一即可）:
      - 硬件信号丢失（DCD/CTS/DSR=False）
      - 连续 N 次探测无回显，且 read(1) 超时返回空字节
      - write() 抛出 SerialException/OSError

    注意: 设备配置完成后处于静默等待状态，不发数据是正常的，
          因此不能单靠 read() 超时判定断开，必须结合 write() 行为。
    """
    debug_print(f"🔄 进入等待断开循环，串口状态: is_open={serial_conn.is_open}")

    # ---- 检测各硬件信号是否可用 ----
    hw_available = {}
    for name, method in [("DCD", serial_conn.getCD),
                         ("CTS", serial_conn.getCTS),
                         ("DSR", serial_conn.getDSR)]:
        try:
            val = method()
            hw_available[name] = True
            debug_print(f"📊 初始 {name}={val}")
        except Exception as e:
            hw_available[name] = False
            debug_print(f"⚠️ {name} 不可用 ({e})")

    # ---- 探测参数 ----
    probe_count = 0
    silent_count = 0            # 连续无回显次数
    SILENT_THRESHOLD = 2        # 连续 2 次无回显即判定断开（约 0.9 秒）
    PROBE_INTERVAL_DISCONNECT = 0.2

    try:
        while True:
            if not serial_conn.is_open:
                print("🔌 检测到串口已关闭")
                return

            # ---- 方案 1: 硬件信号检测（毫秒级） ----
            hw_debug = []
            for name, method in [("DCD", serial_conn.getCD),
                                 ("CTS", serial_conn.getCTS),
                                 ("DSR", serial_conn.getDSR)]:
                if hw_available.get(name):
                    try:
                        val = method()
                        hw_debug.append(f"{name}={val}")
                        if not val:
                            print(f"🔌 检测到 {name} 信号丢失 → 物理断开")
                            return
                    except Exception as e:
                        hw_debug.append(f"{name}=ERR({e})")

            # ---- 方案 2: 探测 + 回显检测 ----
            probe_count += 1

            # 写前状态
            in_waiting_before = serial_conn.in_waiting
            out_waiting_before = serial_conn.out_waiting

            serial_conn.write(b'\n')

            # 写后状态
            out_waiting_after = serial_conn.out_waiting

            time.sleep(PROBE_INTERVAL_DISCONNECT)

            # 读取回显
            data_found = False
            total_bytes = 0
            chunks_info = []
            while serial_conn.in_waiting:
                chunk = serial_conn.read(serial_conn.in_waiting)
                total_bytes += len(chunk)
                text = chunk.decode('utf-8', errors='ignore')
                chunks_info.append(f"{len(chunk)}B:{repr(text[:100])}")
                data_found = True

            # 打印完整探测信息
            hw_str = " | ".join(hw_debug) if hw_debug else "(无硬件信号)"
            in_waiting_after = serial_conn.in_waiting
            debug_print(
                f"⏎ 探测#{probe_count} | "
                f"写前in={in_waiting_before} 写后out={out_waiting_after} | "
                f"in_waiting={in_waiting_after} | "
                f"回显={'✅' if data_found else '❌'} "
                f"{'|'.join(chunks_info) if chunks_info else '(无)'} | "
                f"静默计数={silent_count+1 if not data_found else 0} | "
                f"{hw_str}"
            )

            if data_found:
                silent_count = 0
            else:
                silent_count += 1

                # 连续多次无回显 → 用 read(1) 确认是否真断开
                if silent_count >= SILENT_THRESHOLD:
                    debug_print(f"🔍 连续 {silent_count} 次无回显，执行 read(1) 确认...")
                    original_to = serial_conn.timeout
                    serial_conn.timeout = 0.5
                    leftover = serial_conn.read(1)
                    serial_conn.timeout = original_to
                    debug_print(f"🔍 read(1) 确认结果: {repr(leftover)}")
                    if leftover == b'':
                        print("🔌 连续无回显 + read() 超时 → 判定为物理断开")
                        return
                    else:
                        debug_print(f"📥 read(1) 仍有数据: {repr(leftover)}")
                        silent_count = 0  # 还有数据，重置计数器

    except (serial.SerialException, OSError) as e:
        print(f"🔌 检测到串口断开: {e}")
        return


def close_serial(serial_conn):
    """安全关闭串口连接"""
    try:
        if serial_conn and serial_conn.is_open:
            serial_conn.close()
            print("🔌 串口已关闭")
    except Exception as e:
        print(f"⚠️ 关闭串口时出现异常: {e}")


# ------------------------------------------------------------------
# 2.3.6 完成标记管理
# ------------------------------------------------------------------
# 内存缓存：记录已下发完成的设备行号，避免重复读写 Excel
_completed_rows_cache = set()

def check_status_col(device):
    """
    - 未完成 → 返回 True（正常下发）
    - 已完成 → 询问用户：y=继续下发（覆盖），n=跳过进入等待断开
    """
    row = device["row"]
    device_name = device["device_name"]

    # ---- 查内存缓存 ----
    if row in _completed_rows_cache:
        print(f"\n⚠️ 设备 {device_name}（Excel 第 {row} 行）已完成下发（内存缓存）")
        while True:
            choice = input("  输入 y 重新下发（覆盖标记），输入 n 跳过（进入等待断开）：").strip().lower()
            if choice == 'y':
                print("  ✅ 继续下发配置...")
                return True
            elif choice == 'n':
                print("  ⏭️ 跳过下发，进入等待断开...")
                return False
            else:
                print("  ❌ 请输入 y 或 n")

    deploy_status = device['deploy_status']
    try:
        is_marked = str(deploy_status).strip().upper() == "TRUE"
    except Exception:
        is_marked = False

    if is_marked:
        print(f"\n⚠️ 设备 {device_name}（Excel 第 {row} 行）的完成标记列已有 TRUE")
        while True:
            choice = input("  输入 y 重新下发（覆盖标记），输入 n 跳过（进入等待断开）：").strip().lower()
            if choice == 'y':
                print("  ✅ 继续下发配置...")
                return True
            elif choice == 'n':
                print("  ⏭️ 跳过下发，进入等待断开...")
                return False
            else:
                print("  ❌ 请输入 y 或 n")
    else:
        return True


def write_status_col(device):
    """在设备所在 Excel 行的完成标记列写入 TRUE，并加入内存缓存"""
    import openpyxl

    excel_path = device["excel_path"]
    row = device["row"]
    marker_col = device["complete_marker_col"]
    status_col_idx = col_name_to_index(marker_col)  # 0-based

    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        ws.cell(row=row, column=status_col_idx + 1, value="TRUE")
        wb.save(excel_path)
        wb.close()
        print(f"  📝 已在 Excel 第 {row} 行 {marker_col.upper()} 列写入 TRUE")
    except Exception as e:
        print(f"  ⚠️ 写入 Excel 失败: {e}")
    _completed_rows_cache.add(row)


# =====================================================================
# 日志系统
# =====================================================================
def init_log():
    """初始化日志时间戳和目录"""
    global log_timestamp
    log_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    os.makedirs(LOG_DIR, exist_ok=True)


def write_log(message):
    """写入执行日志"""
    global log_timestamp
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        with open(f"{LOG_DIR}/ConsoleDeploy_{log_timestamp}.txt", "a", encoding="utf-8") as f:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{timestamp}] {message}\n")
    except Exception as e:
        print(f"⚠️ 写入日志失败: {e}")


def write_complete(device_name):
    """写入完成记录"""
    global log_timestamp
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        with open(f"{LOG_DIR}/Complete_{log_timestamp}.txt", "a", encoding="utf-8") as f:
            f.write(f"{device_name}  √\n")
    except Exception as e:
        print(f"⚠️ 写入完成记录失败: {e}")


def write_failed(device_name, reason=""):
    """写入失败记录"""
    global log_timestamp
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        with open(f"{LOG_DIR}/Failed_{log_timestamp}.txt", "a", encoding="utf-8") as f:
            f.write(f"{device_name}  ×  {reason}\n")
    except Exception as e:
        print(f"⚠️ 写入失败记录失败: {e}")


# =====================================================================
# 程序入口
# =====================================================================
if __name__ == "__main__":
    # ---- 调试变量警告 ----
    if PRINT_DEBUG or SAVE_FLAG_DEBUG==0:
        print("⚠️  =============================================")
        print("⚠️  请注意，有调试变量被修改不合理！")
        if PRINT_DEBUG: print(f"⚠️     PRINT_DEBUG    = {PRINT_DEBUG}")
        if SAVE_FLAG_DEBUG==0: print(f"⚠️     SAVE_FLAG_DEBUG = {SAVE_FLAG_DEBUG}")
        print("⚠️  该设置会影响真实环境，请慎重！")
        print("⚠️  =============================================\n")

    # 初始化日志
    init_log()

    # 输入层
    excel_path, mode, start_row, print_config_flag = input_layer()

    # 计算层
    device_configs = calculation_layer(excel_path, mode, start_row)

    if not device_configs:
        print("⚠️ 没有读取到任何设备配置，程序退出。")
        input("按任意键退出...")
        exit()

    if mode == 'sequence':
        # 顺序模式：直接按顺序下发，不经过匹配流程
        print(f"\n{'='*50}")
        print("      顺序模式 — 按 Excel 行号顺序下发")
        print(f"{'='*50}")
        try:
            deploy_layer_sequence(device_configs, print_config_flag)
        except KeyboardInterrupt:
            print("\n\n🛑 用户中断，程序退出。")
        except Exception as e:
            print(f"\n❌ 程序异常: {e}")
            import traceback
            traceback.print_exc()
            input("按任意键退出...")
    else:
        # 匹配模式：监听 → 匹配 → 下发 → 返回监听（循环）
        print(f"\n{'='*50}")
        print("      匹配模式 — 监听设备连接，按设备名匹配下发")
        print(f"{'='*50}")
        try:
            deploy_layer(device_configs, print_config_flag)
        except KeyboardInterrupt:
            print("\n\n🛑 用户中断，程序退出。")
        except Exception as e:
            print(f"\n❌ 程序异常: {e}")
            import traceback
            traceback.print_exc()
            input("按任意键退出...")
