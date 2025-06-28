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
create_line = BPA_models.create_line
create_T2 = BPA_models.create_T2
create_T3 = BPA_models.create_T3

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
            if original_lines is None or categorized is None:
                return
            self.modify_bq_cards(categorized, bq_dist, bq_owner, bq_vol, bq_min_cap, bq_max_cap, modifications)
            output_data = self.write_back_dat(original_lines)
            st.download_button(
                label="下载修改后的文件",
                data=output_data,
                file_name=bq_output_filename,
                mime="application/octet-stream",
                key="bq_download"
            )
            self.log(f"修改完成，准备下载: {b_output_filename}")

    def create_l_tab(self):
        st.subheader("生成 L卡 参数")
        col1, col2 = st.columns(2)
        with col1:
            l_v = st.text_input("电压等级 (V)", value="", key="l_voltage")
            l_length = st.text_input("线路长度 (km)", value="", key="l_length")
            l_n1 = st.text_input("首端名字 (N1)", value="", key="l_n1")
        with col2:
            l_n2 = st.text_input("末端名字 (N2)", value="", key="l_n2")
            l_t = st.text_input("型号 (T)", value="", key="l_type")
            l_p = st.text_input("并列号 (P)", value="", key="l_parallel")

        if 'l_output' not in st.session_state:
            st.session_state.l_output = ""

        if st.button("生成 L卡", key="l_generate", type="primary"):
            try:
                V = float(l_v.strip())
            except ValueError:
                st.error("电压等级 (V) 必须是一个有效的数字。")
                self.log("错误: 电压等级 (V) 输入无效。", level="ERROR")
                return

            try:
                L = float(l_length.strip())
            except ValueError:
                st.error("线路长度 (km) 必须是一个有效的数字。")
                self.log("错误: 线路长度 (km) 输入无效。", level="ERROR")
                return

            N1, N2, T, P = l_n1.strip(), l_n2.strip(), l_t.strip(), l_p.strip()
            if not all([N1, N2, T, P]):
                st.error("请确保所有字段均已填写。")
                self.log("错误: 缺少必要的输入字段。", level="ERROR")
                return

            try:
                P = int(P)
            except ValueError:
                st.error("并列号 (P) 必须是一个有效的整数。")
                self.log("错误: 并列号 (P) 输入无效。", level="ERROR")
                return

            try:
                l_card = create_line(V=V, L=L, N1=N1, N2=N2, T=T, P=P)
                l_card_str = l_card.gen() + '\n'
                if not l_card_str:
                    raise ValueError("生成的 L卡 内容为空。")
                st.session_state.l_output += l_card_str
                self.log(f"L卡已生成: \n{l_card_str.rstrip()}")
            except Exception as e:
                st.error(f"生成 L卡 失败: {e}")
                self.log(f"错误: 生成 L卡 失败: {e}", level="ERROR")

        if st.button("清空输出", key="l_clear"):
            st.session_state.l_output = ""
            self.log("L卡输出区域已清空。")

        st.subheader("生成的 L卡")
        st.text_area("L卡输出", value=st.session_state.l_output, height=200, key="l_output")

    def create_t3_tab(self):
        st.subheader("生成 三卷变 参数")
        col1, col2 = st.columns(2)
        with col1:
            t3_V1 = st.text_input("电压等级 V1", value="525.0", key="t3_v1")
            t3_V2 = st.text_input("电压等级 V2", value="230.0", key="t3_v2")
            t3_V3 = st.text_input("电压等级 V3", value="37.0", key="t3_v3")
            t3_VB = st.text_input("电压等级 VB", value="1.0", key="t3_vb")
            t3_N1 = st.text_input("节点名称 N1", value="", key="t3_n1")
            t3_N2 = st.text_input("节点名称 N2", value="", key="t3_n2")
        with col2:
            t3_N3 = st.text_input("节点名称 N3", value="", key="t3_n3")
            t3_NB = st.text_input("节点名称 NB", value="", key="t3_nb")
            t3_Owner = st.text_input("所有者 Owner", value="", key="t3_owner")
            t3_cap = st.text_input("容量 cap", value="1000.0", key="t3_cap")
            t3_x_pu = st.text_input("电抗参数 x_pu1, x_pu2, x_pu3", value="13,64,44", key="t3_x_pu")
            t3_tap_vol = st.text_input("抽头电压 tap_vol1, tap_vol2, tap_vol3", value="525.0,230.0,37.0", key="t3_tap_vol")

        if 't3_output' not in st.session_state:
            st.session_state.t3_output = ""

        if st.button("生成 三卷变", key="t3_generate", type="primary"):
            try:
                V1 = float(t3_V1.strip())
                V2 = float(t3_V2.strip())
                V3 = float(t3_V3.strip())
                VB = float(t3_VB.strip())
            except ValueError:
                st.error("电压等级 (V1, V2, V3, VB) 必须是有效的数字。")
                self.log("错误: 电压等级 (V1, V2, V3, VB) 输入无效。", level="ERROR")
                return

            N1, N2, N3, NB, Owner = t3_N1.strip(), t3_N2.strip(), t3_N3.strip(), t3_NB.strip(), t3_Owner.strip()
            cap, x_pu, tap_vol = t3_cap.strip(), t3_x_pu.strip(), t3_tap_vol.strip()

            if not all([N1, N2, N3, NB, Owner, cap, x_pu, tap_vol]):
                st.error("请确保所有字段均已填写。")
                self.log("错误: 缺少必要的输入字段。", level="ERROR")
                return

            try:
                cap = float(cap)
            except ValueError:
                st.error("容量 (cap) 必须是一个有效的数字。")
                self.log("错误: 容量 (cap) 输入无效。", level="ERROR")
                return

            try:
                x_pu_list = [float(x.strip()) for x in x_pu.split(',')]
                if len(x_pu_list) != 3:
                    raise ValueError("x_pu 必须包含三个数值，以逗号分隔。")
            except ValueError as e:
                st.error(f"电抗参数 x_pu 无效: {e}")
                self.log(f"错误: 电抗参数 x_pu 无效: {e}", level="ERROR")
                return

            try:
                tap_vol_list = [float(vol.strip()) for vol in tap_vol.split(',')]
                if len(tap_vol_list) != 3:
                    raise ValueError("tap_vol 必须包含三个数值，以逗号分隔。")
            except ValueError as e:
                st.error(f"抽头电压 tap_vol 无效: {e}")
                self.log(f"错误: 抽头电压 tap_vol 无效: {e}", level="ERROR")
                return

            try:
                t3_card = create_T3(
                    V1=V1, V2=V2, V3=V3, VB=VB, N1=N1, N2=N2, N3=N3, NB=NB,
                    Owner=Owner, cap=cap, x_pu=x_pu_list,
                    tap_vol1=tap_vol_list[0], tap_vol2=tap_vol_list[1], tap_vol3=tap_vol_list[2]
                )
                st.session_state.t3_output += t3_card + '\n'
                self.log(f"三卷变已生成: \n{t3_card.rstrip()}")
            except Exception as e:
                st.error(f"生成三卷变失败: {e}")
                self.log(f"错误: 生成三卷变失败: {e}", level="ERROR")

        if st.button("清空输出", key="t3_clear"):
            st.session_state.t3_output = ""
            self.log("三卷变输出区域已清空。")

        st.subheader("生成的 三卷变")
        st.text_area("三卷变输出", value=st.session_state.t3_output, height=200, key="t3_output")

    def create_t2_tab(self):
        st.subheader("生成 两卷变T卡 参数")
        col1, col2 = st.columns(2)
        with col1:
            t2_V1 = st.text_input("电压等级 V1", value="525.0", key="t2_v1")
            t2_V2 = st.text_input("电压等级 V2", value="230.0", key="t2_v2")
            t2_N1 = st.text_input("节点名称 N1", value="", key="t2_n1")
            t2_N2 = st.text_input("节点名称 N2", value="", key="t2_n2")
        with col2:
            t2_Owner = st.text_input("所有者 Owner", value="", key="t2_owner")
            t2_cap = st.text_input("容量 cap", value="1000.0", key="t2_cap")
            t2_x_pu = st.text_input("电抗参数 x_pu", value="13", key="t2_x_pu")
            t2_tap_vol = st.text_input("抽头电压 tap_vol1, tap_vol2", value="525.0,230.0", key="t2_tap_vol")

        if 't2_output' not in st.session_state:
            st.session_state.t2_output = ""

        if st.button("生成 两卷变T卡", key="t2_generate", type="primary"):
            try:
                V1 = float(t2_V1.strip())
                V2 = float(t2_V2.strip())
            except ValueError:
                st.error("电压等级 (V1, V2) 必须是有效的数字。")
                self.log("错误: 电压等级 (V1, V2) 输入无效。", level="ERROR")
                return

            N1, N2, Owner = t2_N1.strip(), t2_N2.strip(), t2_Owner.strip()
            cap, x_pu, tap_vol = t2_cap.strip(), t2_x_pu.strip(), t2_tap_vol.strip()

            if not all([N1, N2, Owner, cap, x_pu, tap_vol]):
                st.error("请确保所有字段均已填写。")
                self.log("错误: 缺少必要的输入字段。", level="ERROR")
                return

            try:
                cap = float(cap)
            except ValueError:
                st.error("容量 (cap) 必须是一个有效的数字。")
                self.log("错误: 容量 (cap) 输入无效。", level="ERROR")
                return

            try:
                x_pu_val = float(x_pu)
            except ValueError:
                st.error("电抗参数 x_pu 必须是一个有效的数字。")
                self.log("错误: 电抗参数 x_pu 输入无效。", level="ERROR")
                return

            try:
                tap_vol_list = [float(vol.strip()) for vol in tap_vol.split(',')]
                if len(tap_vol_list) != 2:
                    raise ValueError("tap_vol 必须包含两个数值，以逗号分隔。")
            except ValueError as e:
                st.error(f"抽头电压 tap_vol 无效: {e}")
                self.log(f"错误: 抽头电压 tap_vol 无效: {e}", level="ERROR")
                return

            try:
                t2_card = create_T2(
                    V1=V1, V2=V2, N1=N1, N2=N2, Owner=Owner, cap=cap,
                    x_pu=x_pu_val, tap_vol1=tap_vol_list[0], tap_vol2=tap_vol_list[1]
                )
                st.session_state.t2_output += t2_card + '\n'
                self.log(f"两卷变T卡已生成: \n{t2_card.rstrip()}")
            except Exception as e:
                st.error(f"生成 两卷变T卡 失败: {e}")
                self.log(f"错误: 生成 两卷变T卡 失败: {e}", level="ERROR")

        if st.button("清空输出", key="t2_clear"):
            st.session_state.t2_output = ""
            self.log("两卷变T卡输出区域已清空。")

        st.subheader("生成的 两卷变T卡")
        st.text_area("两卷变T卡输出", value=st.session_state.t2_output, height=200, key="t2_output")

def main():
    st.set_page_config(page_title="PSD-BPA DAT文件批量修改生成工具", layout="wide")
    st.title("PSD-BPA DAT文件批量修改生成工具")
    st.markdown("**使用条款**: 本应用仅限授权用户使用。请勿上传敏感数据。")

    app = DATModifierApp()

    tabs = st.tabs(["B卡修改", "BQ卡修改", "L卡生成", "三卷变T*3卡生成", "两卷变T卡生成"])
    with tabs[0]:
        app.create_b_tab()
    with tabs[1]:
        app.create_bq_tab()
    with tabs[2]:
        app.create_l_tab()
    with tabs[3]:
        app.create_t3_tab()
    with tabs[4]:
        app.create_t2_tab()

    with st.expander("查看日志"):
        st.text_area("操作日志 (可滚动查看，不可编辑)", value="\n".join(app.logs), height=200, key="log_output")

if __name__ == "__main__":
    main()
