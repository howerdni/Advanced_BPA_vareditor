import streamlit as st
import pandas as pd
import unicodedata
import os
import io
import ast
from datetime import datetime
from cryptography.fernet import Fernet
import importlib.util
import sys
import json

# Decrypt and load BPA_models
def load_encrypted_module():
    try:
        with open('key.txt', 'rb') as f:
            key = f.read()
        cipher = Fernet(key)
        with open('BPA_models.encrypted', 'rb') as f:
            encrypted = f.read()
        code = cipher.decrypt(encrypted).decode('utf-8')
        spec = importlib.util.spec_from_loader('BPA_models', loader=None)
        module = importlib.util.module_from_spec(spec)
        sys.modules['BPA_models'] = module
        exec(code, module.__dict__)
        return module
    except Exception as e:
        st.error(f"无法解密 BPA_models.encrypted: {e}")
        raise

BPA_models = load_encrypted_module()
BCard = BPA_models.BCard
BQCard = BPA_models.BQCard
LCard = BPA_models.LCard
create_T2 = BPA_models.create_T2
create_T3 = BPA_models.create_T3

# Embedded line_parameters.json content
LINE_PARAMETERS = {
    "4-630": {"r_km": 0.013, "x_km": 0.267, "B1_km": 4.24, "Irate": 4000},
    "4-400": {"r_km": 0.02, "x_km": 0.28, "B1_km": 4.125, "Irate": 3000},
    "4-300-2": {"r_km": 0.028, "x_km": 0.282, "B1_km": 4.09, "Irate": 2256},
    "4-300-1": {"r_km": 0.033, "x_km": 0.27, "B1_km": 3.51, "Irate": 2256},
    "4-720": {"r_km": 0.011, "x_km": 0.241, "B1_km": 4.69, "Irate": 4240},
    "4-240": {"r_km": 0.021, "x_km": 0.208, "B1_km": 5.71, "Irate": 2400},
    "1-400": {"r_km": 0.08, "x_km": 0.41, "B1_km": 2.787, "Irate": 660},
    "2-300": {"r_km": 0.0541, "x_km": 0.31, "B1_km": 3.62, "Irate": 1128},
    "2-400": {"r_km": 0.04, "x_km": 0.2975, "B1_km": 3.81, "Irate": 1365},
    "2-630": {"r_km": 0.0304, "x_km": 0.277, "B1_km": 4.03, "Irate": 1890},
    "X-1000": {"r_km": 0.0232, "x_km": 0.16, "B1_km": 51.83, "Irate": 750},
    "X-2000": {"r_km": 0.0201, "x_km": 0.1420, "B1_km": 68.17, "Irate": 1500},
    "X-2500": {"r_km": 0.018, "x_km": 0.14, "B1_km": 73.51, "Irate": 1890},
    "8-630": {"r_km": 0.0065, "x_km": 0.2538, "B1_km": 4.51, "Irate": 8000}
}

def create_line(V=525., L=1, N1='', N2='', T='4-630', P=1):
    try:
        voltage = float(V)
        length = float(L)
        name1 = str(N1)
        name2 = str(N2)
        conductor_type = str(T)
        par = int(P)
    except (ValueError, TypeError) as e:
        raise ValueError(f"无效的输入参数: {e}")

    parameters = LINE_PARAMETERS
    if conductor_type not in parameters:
        raise ValueError(f"不支持的线路型号: {conductor_type}")

    params = parameters[conductor_type]
    try:
        r_km = float(params["r_km"])
        x_km = float(params["x_km"])
        B1_km = float(params["B1_km"])
        Irate = float(params["Irate"])
    except (KeyError, ValueError) as e:
        raise ValueError(f"无效的线路参数内容: {e}")

    factor1 = 1 / ((voltage ** 2) / 100)
    factor2 = (voltage ** 2) / 100 / 1000000

    r_pu = r_km * length * factor1
    x_pu = x_km * length * factor1
    b2_pu = B1_km * length * factor2 / 2

    strLCard0 = f"L     {name1:<8}100. {name2:<8}100.1            0.01              1   投运"
    lcard_instance = LCard(strLCard0)
    lcard_instance._r_pu = lcard_instance._format_value(f"{r_pu:.10f}".lstrip('0'), 6, 5)
    lcard_instance._x_pu = lcard_instance._format_value(f"{x_pu:.10f}".lstrip('0'), 6, 5)
    lcard_instance._b2_pu = lcard_instance._format_value(f"{b2_pu:.10f}".lstrip('0'), 6, 5)
    lcard_instance._length = lcard_instance._format_value(length, 4, 1)
    lcard_instance.bus_name1 = name1
    lcard_instance.bus_name2 = name2
    lcard_instance.vol_rank1 = voltage
    lcard_instance.vol_rank2 = voltage
    lcard_instance.parallel = par
    lcard_instance.length = length
    lcard_instance.irate = Irate

    return lcard_instance

def _format_string(value, length):
    def char_width(char):
        ea_width = unicodedata.east_asian_width(char)
        return 2 if ea_width in ('F', 'W', 'A') else 1

    def string_width(string):
        return sum(char_width(c) for c in string)

    formatted_value = ''
    current_width = 0
    for char in str(value):
        if string_width(formatted_value) >= length:
            break
        w = char_width(char)
        if current_width + w > length:
            break
        formatted_value += char
        current_width += w

    while string_width(formatted_value) < length:
        formatted_value += ' '
    return formatted_value

class DATModifierApp:
    def __init__(self):
        self.logs = []
        self.uploaded_files = []
        self.b_parameters = {
            "dist": "str",
            "owner": "str",
            "mw_load": "num",
            "mvar_load": "num",
            "shunt_mw": "num",
            "shunt_var": "num",
            "capacity": "num",
            "pout": "num",
            "qout": "num",
        }
        self.bq_parameters = {
            "dist": "str",
            "owner": "str",
            "mw_load": "num",
            "mvar_load": "num",
            "shunt_mw": "num",
            "shunt_var": "num",
            "capacity": "num",
            "pout": "num",
            "qout": "num",
            "qout_min": "num",
            "vmax_pu": "num",
        }

    def log(self, msg, level="INFO"):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] [{level}] {msg}"
        self.logs.append(log_message)
        try:
            with open("operation_log.txt", "a", encoding='utf-8') as log_file:
                log_file.write(log_message + "\n")
        except Exception as e:
            print(f"无法写入日志文件: {e}")

    def log_file_upload(self, file):
        if file is not None:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            file_size = len(file.getvalue()) / 1024  # Size in KB
            log_message = f"[{timestamp}] [UPLOAD] File uploaded: {file.name}, Size: {file_size:.2f} KB"
            self.logs.append(log_message)
            self.uploaded_files.append({
                "name": file.name,
                "size_kb": file_size,
                "timestamp": timestamp
            })
            try:
                with open("operation_log.txt", "a", encoding='utf-8') as log_file:
                    log_file.write(log_message + "\n")
            except Exception as e:
                print(f"无法写入日志文件: {e}")

    def read_and_parse_dat(self, file_content):
        self.log("读取文件内容")
        try:
            lines = file_content.decode('gbk', errors='ignore').splitlines()
        except Exception as e:
            st.error(f"无法读取文件: {e}")
            self.log(f"错误: 无法读取文件: {e}", level="ERROR")
            return None, None

        original_lines = []
        categorized_objects = {'B': [], 'BQ': []}

        for idx, line in enumerate(lines):
            line_stripped = line.rstrip('\n')
            if line_stripped.startswith("B "):
                try:
                    b_obj = BCard(line_stripped, idx)
                    categorized_objects['B'].append(b_obj)
                    original_lines.append(b_obj)
                except Exception as e:
                    self.log(f"错误: 解析 B卡失败 (行 {idx}): {e}", level="ERROR")
            elif line_stripped.startswith("BQ"):
                try:
                    bq_obj = BQCard(line_stripped, idx)
                    categorized_objects['BQ'].append(bq_obj)
                    original_lines.append(bq_obj)
                except Exception as e:
                    self.log(f"错误: 解析 BQ卡失败 (行 {idx}): {e}", level="ERROR")
            else:
                original_lines.append(line_stripped)

        self.log("文件解析完成。")
        return original_lines, categorized_objects

    def write_back_dat(self, original_lines):
        output = io.StringIO()
        for item in original_lines:
            if hasattr(item, 'gen') and callable(item.gen):
                output.write(item.gen() + '\n')
            else:
                output.write(item + '\n')
        return output.getvalue().encode('gbk')

    def modify_b_cards(self, categorized_objects, dist_f, owner_f, vol_f, modifications):
        b_list = categorized_objects.get('B', [])
        filtered = []

        user_vol = None
        if vol_f:
            try:
                user_vol = float(vol_f)
            except ValueError:
                self.log(f"警告: B卡电压 '{vol_f}' 非法，忽略电压筛选", level="WARNING")
                user_vol = None

        dist_list = None
        if dist_f:
            dist_list = [item.strip() for item in dist_f.split(',') if item.strip()]
            if not dist_list:
                self.log(f"警告: B卡分区 '{dist_f}' 格式非法，无有效值", level="WARNING")

        owner_list = None
        if owner_f:
            owner_list = [item.strip() for item in owner_f.split(',') if item.strip()]
            if not owner_list:
                self.log(f"警告: B卡所有者 '{owner_f}' 格式非法，无有效值", level="WARNING")

        for b_obj in b_list:
            if dist_list and b_obj.dist.strip() not in dist_list:
                continue
            if owner_list and b_obj.owner.strip() not in owner_list:
                continue
            if user_vol is not None:
                b_vol = getattr(b_obj, "vol_rank", "0")
                try:
                    b_vol_f = float(b_vol)
                    if abs(b_vol_f - user_vol) >= 0.1:
                        continue
                except ValueError:
                    continue
            filtered.append(b_obj)

        self.log(f"B卡符合条件: {len(filtered)}")

        for b_obj in filtered:
            bus_name = getattr(b_obj, 'bus_name', 'NoBusName?')
            for param, mod in modifications.items():
                if not mod['apply']:
                    continue
                old_val = getattr(b_obj, param, "0")
                ptype = self.b_parameters[param]
                if ptype == "str":
                    new_val = mod['value']
                    setattr(b_obj, param, new_val)
                    self.log(f"B卡 [BusName={bus_name}]: {param} 由 '{old_val}' -> '{new_val}'")
                else:
                    try:
                        old_val_num = float(old_val)
                        if mod['method'] == "set":
                            new_num = float(mod['value'])
                            setattr(b_obj, param, f"{new_num:.2f}")
                            self.log(f"B卡 [BusName={bus_name}]: {param} 设为 {new_num}")
                        else:
                            coeff = float(mod['value'])
                            new_num = old_val_num * coeff
                            setattr(b_obj, param, f"{new_num:.2f}")
                            self.log(f"B卡 [BusName={bus_name}]: {param} 由 {old_val_num}×{coeff}={new_num}")
                    except ValueError:
                        self.log(f"错误: 无法将 B卡 [BusName={bus_name}] 的 {param}='{old_val}' 转为浮点数", level="ERROR")

    def modify_bq_cards(self, categorized_objs, dist_f, owner_f, vol_f, min_cap, max_cap, modifications):
        bq_list = categorized_objs.get('BQ', [])
        filtered = []

        user_vol = None
        if vol_f:
            try:
                user_vol = float(vol_f)
            except ValueError:
                self.log(f"警告: BQ卡电压'{vol_f}'非法，忽略电压筛选", level="WARNING")
                user_vol = None

        user_cap_min = None
        user_cap_max = None
        if min_cap:
            try:
                user_cap_min = float(min_cap)
            except ValueError:
                self.log(f"警告: BQ卡 capacity下限'{min_cap}'非法, 不做下限限制", level="WARNING")
                user_cap_min = None
        if max_cap:
            try:
                user_cap_max = float(max_cap)
            except ValueError:
                self.log(f"警告: BQ卡 capacity上限'{max_cap}'非法, 不做上限限制", level="WARNING")
                user_cap_max = None

        dist_list = None
        if dist_f:
            dist_list = [item.strip() for item in dist_f.split(',') if item.strip()]
            if not dist_list:
                self.log(f"警告: BQ卡分区 '{dist_f}' 格式非法，无有效值", level="WARNING")

        owner_list = None
        if owner_f:
            owner_list = [item.strip() for item in owner_f.split(',') if item.strip()]
            if not owner_list:
                self.log(f"警告: BQ卡所有者 '{owner_f}' 格式非法，无有效值", level="WARNING")

        for bq_obj in bq_list:
            if dist_list and bq_obj.dist.strip() not in dist_list:
                continue
            if owner_list and bq_obj.owner.strip() not in owner_list:
                continue
            if user_vol is not None:
                try:
                    bqv = float(getattr(bq_obj, "vol_rank", "0"))
                    if abs(bqv - user_vol) >= 0.1:
                        continue
                except ValueError:
                    continue

            try:
                obj_cap = float(getattr(bq_obj, "capacity", "0"))
                if user_cap_min is not None and obj_cap < user_cap_min:
                    continue
                if user_cap_max is not None and obj_cap > user_cap_max:
                    continue
            except ValueError:
                continue

            filtered.append(bq_obj)

        self.log(f"BQ卡符合条件: {len(filtered)}")

        for bq_obj in filtered:
            bus_name = getattr(bq_obj, 'bus_name', 'NoBusName?')
            for param, mod in modifications.items():
                if not mod['apply']:
                    continue
                old_val = getattr(bq_obj, param, "0")
                ptype = self.bq_parameters[param]
                if ptype == "str":
                    new_val = mod['value']
                    setattr(bq_obj, param, new_val)
                    self.log(f"BQ卡 [BusName={bus_name}]: {param} 由'{old_val}'->'{new_val}'")
                else:
                    try:
                        old_val_num = float(old_val)
                        if param in ["pout", "mw_load"] and mod['method'] == "pct":
                            percent = float(mod['value'])
                            capacity = float(getattr(bq_obj, "capacity", "0"))
                            new_f = capacity * percent
                            setattr(bq_obj, param, f"{new_f:.2f}")
                            self.log(f"BQ卡 [BusName={bus_name}]: {param} 从{old_val_num}设为容量 {capacity} × {percent} = {new_f}")
                        elif mod['method'] == "set":
                            new_f = float(mod['value'])
                            setattr(bq_obj, param, f"{new_f:.2f}")
                            self.log(f"BQ卡 [BusName={bus_name}]: {param} 从{old_val_num}设为 {new_f}")
                        elif mod['method'] == "mul":
                            coeff = float(mod['value'])
                            new_f = old_val_num * coeff
                            setattr(bq_obj, param, f"{new_f:.2f}")
                            self.log(f"BQ卡 [BusName={bus_name}]: {param} 由 {old_val_num}×{coeff}={new_f}")
                    except ValueError:
                        self.log(f"错误: 无法把 BQ卡 [BusName={bus_name}] 的 {param}='{old_val}' 转为float", level="ERROR")

    def create_b_tab(self):
        st.subheader("文件选择 (B卡)")
        b_input_file = st.file_uploader("上传输入.dat文件 (B卡)", type=["dat"], key="b_input")
        if b_input_file:
            self.log_file_upload(b_input_file)
        b_output_filename = st.text_input("输出.dat文件名", value="modified_b.dat", key="b_output_filename")

        st.subheader("B卡筛选条件")
        col1, col2, col3 = st.columns(3)
        with col1:
            b_dist = st.text_input("分区(dist, 用逗号分隔, 如 C1,D1)", value="C1,D1", key="b_dist")
        with col2:
            b_owner = st.text_input("所有者(owner, 用逗号分隔, 如 苏,锡)", value="苏,锡", key="b_owner")
        with col3:
            b_vol = st.text_input("电压(vol_rank)", value="", key="b_vol_rank")

        st.subheader("B卡修改字段")
        modifications = {}
        for param, ptype in self.b_parameters.items():
            with st.expander(f"修改 {param}"):
                apply = st.checkbox(f"启用 {param} 修改", key=f"b_{param}_apply")
                if apply:
                    if ptype == "str":
                        method = "set"
                        value = st.text_input(f"{param} 新值", key=f"b_{param}_value")
                    else:
                        method = st.radio(
                            f"{param} 修改方式",
                            ["设值", "乘系数"],
                            key=f"b_{param}_method",
                            format_func=lambda x: x
                        )
                        method = "set" if method == "设值" else "mul"
                        if method == "set":
                            value = st.text_input(f"{param} 新值", key=f"b_{param}_value")
                        else:
                            value = st.text_input(f"{param} 系数", key=f"b_{param}_coeff")
                    modifications[param] = {'apply': True, 'method': method, 'value': value}
                else:
                    modifications[param] = {'apply': False, 'method': None, 'value': None}

        if st.button("执行修改 (B卡)", key="b_execute", type="primary"):
            if not b_input_file:
                st.warning("请选择输入的 .dat 文件。")
                return
            if not b_output_filename:
                st.warning("请指定输出文件名。")
                return

            self.log("开始处理 B卡...")
            original_lines, categorized = self.read_and_parse_dat(b_input_file.read())
            if original_lines is None or categorized is None:
                return
            self.modify_b_cards(categorized, b_dist, b_owner, b_vol, modifications)
            output_data = self.write_back_dat(original_lines)
            st.download_button(
                label="下载修改后的文件",
                data=output_data,
                file_name=b_output_filename,
                mime="application/octet-stream",
                key="b_download"
            )
            self.log(f"修改完成，准备下载: {b_output_filename}")

    def create_bq_tab(self):
        st.subheader("文件选择 (BQ卡)")
        bq_input_file = st.file_uploader("上传输入.dat文件 (BQ卡)", type=["dat"], key="bq_input")
        if bq_input_file:
            self.log_file_upload(bq_input_file)
        bq_output_filename = st.text_input("输出.dat文件名", value="modified_bq.dat", key="bq_output_filename")

        st.subheader("BQ卡筛选条件")
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            bq_dist = st.text_input("分区(dist, 用逗号分隔, 如 C1,D1)", value="C1,D1", key="bq_dist")
        with col2:
            bq_owner = st.text_input("所有者(owner, 用逗号分隔, 如 苏,锡)", value="苏,锡", key="bq_owner")
        with col3:
            bq_vol = st.text_input("电压(vol_rank)", value="", key="bq_vol_rank")
        with col4:
            bq_min_cap = st.text_input("容量下限", value="", key="bq_min_cap")
        with col5:
            bq_max_cap = st.text_input("容量上限", value="", key="bq_max_cap")

        st.subheader("BQ卡修改字段")
        modifications = {}
        for param, ptype in self.bq_parameters.items():
            with st.expander(f"修改 {param}"):
                apply = st.checkbox(f"启用 {param} 修改", key=f"bq_{param}_apply")
                if apply:
                    if ptype == "str":
                        method = "set"
                        value = st.text_input(f"{param} 新值", key=f"bq_{param}_value")
                    else:
                        options = ["设值", "乘系数"]
                        if param in ["pout", "mw_load"]:
                            options.append("设为容量*系数")
                        method = st.radio(
                            f"{param} 修改方式",
                            options,
                            key=f"bq_{param}_method",
                            format_func=lambda x: x
                        )
                        if method == "设值":
                            method = "set"
                            value = st.text_input(f"{param} 新值", key=f"bq_{param}_value")
                        elif method == "乘系数":
                            method = "mul"
                            value = st.text_input(f"{param} 系数", key=f"bq_{param}_coeff")
                        else:
                            method = "pct"
                            value = st.text_input(f"{param} 容量百分比 (如 0.5 表示 50%)", key=f"bq_{param}_pct")
                    modifications[param] = {'apply': True, 'method': method, 'value': value}
                else:
                    modifications[param] = {'apply': False, 'method': None, 'value': None}

        if st.button("执行修改 (BQ卡)", key="bq_execute", type="primary"):
            if not bq_input_file:
                st.warning("请选择输入的 .dat 文件。")
                return
            if not bq_output_filename:
                st.warning("请指定输出文件名。")
                return

            self.log("开始处理 BQ卡...")
            original_lines, categorized = self.read_and_parse_dat(bq_input_file.read())
            if original_lines is None
