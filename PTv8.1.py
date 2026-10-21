"""
診所智慧排班系統 (Clinic Roster Master v4.7 專業圖示旗艦版)
功能特點：
1. 精準時段配置：平日 (W1-W5) 精確 1個A1、1個C1；週六 (W6) 精確 2個A1、1個A。
2. Excel 匯出完整還原示範檔 5 列結構：包含當日時數、班別代碼、以及三時段圖示 (☼:加長班, ○:一般班, -:無班/休假, 特殊任務文字)。
3. 人力不足高韌性排班：缺工時照常生成最優班表，並於「⚠️ 診所缺額」列直覺標記。
4. 預假支援連續區間語法 (如 11-14, 18-23)。
5. 固定任務僅作為保留時段，不計工時且不與個人排休互斥。
6. 視窗下方動態即時顯示「當月總人力時數需求」與「當月同仁最低工時總和」。

執行方式：streamlit run app.py
"""

import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import datetime
import io
import re
from typing import List, Dict, Set, Tuple
from dataclasses import dataclass
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. 班別與資料模型定義
# ==========================================
@dataclass
class ShiftDef:
    code: str         # 班別代碼: A, A1, B, C, C1
    name: str         # 班別名稱
    time_range: str   # 時段文字: 08:30-12:30, 08:30-13:30 等
    slot: str         # 對應大時段: AM, PM, NIGHT
    hours: float      # 計算工時 (hr)
    is_default: bool = False

@dataclass
class StaffInput:
    id: str
    name: str
    target_hours: int
    min_hours: int
    max_hours: int
    daily_max_hours: int = 10
    max_ac_shifts: int = 4         # 每位員工月度純 AC 班上限 (兼職/半職設為 0)
    max_c_shifts: int = 10         # 每位員工月度 C 班 (晚班) 上限 (無法上晚班者設為 0)
    is_senior: bool = False
    unavail_str: str = ""          # 預假: "5, 11-14, 18-23, 12-C"
    blocked_weekday_str: str = ""  # 星期限制: "W1-B, W2-BC" 或 "all except: W2-B, W3-A"
    fixed_or_only_str: str = ""    # ALL EXCEPT 或 固定任務 (如 17-B:內訓, W1-A:公司)

# ==========================================
# 2. 字串解析輔助函式
# ==========================================
def parse_clinic_closures(closure_str: str) -> Set[Tuple[int, str]]:
    closed = set()
    if not closure_str or pd.isna(closure_str):
        return closed
    tokens = [t.strip().upper() for t in str(closure_str).split(",") if t.strip()]
    for token in tokens:
        if "-" in token:
            parts = token.split("-")
            try:
                day = int(parts[0])
                slots_part = parts[1]
                if "A" in slots_part: closed.add((day, "AM"))
                if "B" in slots_part: closed.add((day, "PM"))
                if "C" in slots_part: closed.add((day, "NIGHT"))
            except ValueError:
                pass
        else:
            try:
                day = int(token)
                closed.add((day, "AM"))
                closed.add((day, "PM"))
                closed.add((day, "NIGHT"))
            except ValueError:
                pass
    return closed

def parse_code_to_slots(code_str: str, shift_map: Dict[str, ShiftDef]) -> List[str]:
    clean_code = code_str.strip().upper()
    if clean_code in ("AM", "PM", "NIGHT"):
        return [clean_code]
    if clean_code in shift_map:
        return [shift_map[clean_code].slot]

    slots = set()
    if "A" in clean_code: slots.add("AM")
    if "B" in clean_code: slots.add("PM")
    if "C" in clean_code: slots.add("NIGHT")

    return list(slots) if slots else [clean_code]

def parse_unavailabilities(unavail_str: str, shift_map: Dict[str, ShiftDef]) -> Set[Tuple[int, str]]:
    res = set()
    if not unavail_str or pd.isna(unavail_str):
        return res
    tokens = [t.strip() for t in str(unavail_str).split(",") if t.strip()]
    for token in tokens:
        m_range = re.match(r"^(\d+)\s*-\s*(\d+)(?:-([A-Za-z0-9]+))?$", token)
        if m_range:
            d_start = int(m_range.group(1))
            d_end = int(m_range.group(2))
            slot_part = m_range.group(3)
            slots_to_block = ["AM", "PM", "NIGHT"]
            if slot_part:
                clean_slot = slot_part.strip().upper()
                temp_slots = set()
                if "A" in clean_slot: temp_slots.add("AM")
                if "B" in clean_slot: temp_slots.add("PM")
                if "C" in clean_slot: temp_slots.add("NIGHT")
                if temp_slots:
                    slots_to_block = list(temp_slots)

            for d in range(min(d_start, d_end), max(d_start, d_end) + 1):
                for sl in slots_to_block:
                    res.add((d, sl))
            continue

        if "-" in token:
            parts = token.split("-")
            try:
                day = int(parts[0])
                slots = parse_code_to_slots(parts[1], shift_map)
                for s in slots:
                    res.add((day, s))
            except ValueError:
                pass
        else:
            try:
                day = int(token)
                res.add((day, "AM"))
                res.add((day, "PM"))
                res.add((day, "NIGHT"))
            except ValueError:
                pass
    return res

def parse_blocked_weekdays(blocked_str: str, shift_map: Dict[str, ShiftDef]) -> Set[Tuple[int, str]]:
    res = set()
    if not blocked_str or pd.isna(blocked_str):
        return res
    text = str(blocked_str).strip()

    if text.upper().startswith("ALL EXCEPT:"):
        only_part = text[len("ALL EXCEPT:"):].strip()
        allowed = set()
        tokens = [t.strip().upper() for t in only_part.split(",") if t.strip()]
        for token in tokens:
            m = re.match(r"W([1-7])(-([A-Z0-9]+))?", token)
            if m:
                w_idx = int(m.group(1)) - 1
                c_part = m.group(3) if m.group(3) else "ALL"
                if c_part == "ALL":
                    allowed.add((w_idx, "AM")); allowed.add((w_idx, "PM")); allowed.add((w_idx, "NIGHT"))
                else:
                    slots = parse_code_to_slots(c_part, shift_map)
                    for s in slots:
                        allowed.add((w_idx, s))
        
        for w in range(7):
            for s in ["AM", "PM", "NIGHT"]:
                if (w, s) not in allowed:
                    res.add((w, s))
        return res

    tokens = [t.strip().upper() for t in text.split(",") if t.strip()]
    for token in tokens:
        match = re.match(r"W([1-7])(-([A-Z0-9]+))?", token)
        if match:
            w_idx = int(match.group(1)) - 1
            code_part = match.group(3)
            if code_part:
                slots = parse_code_to_slots(code_part, shift_map)
                for s in slots:
                    res.add((w_idx, s))
            else:
                res.add((w_idx, "AM")); res.add((w_idx, "PM")); res.add((w_idx, "NIGHT"))
    return res

def parse_fixed_and_only(fixed_str: str, all_days: List[int], year: int, month: int, shift_map: Dict[str, ShiftDef]):
    fixed_tasks = {}
    extra_unavails = set()
    if not fixed_str or pd.isna(fixed_str):
        return fixed_tasks, extra_unavails

    text = str(fixed_str).strip()
    if text.upper().startswith("ALL EXCEPT:"):
        only_part = text[len("ALL EXCEPT:"):].strip()
        allowed_slots = set()
        tokens = [t.strip() for t in only_part.split(",") if t.strip()]
        for token in tokens:
            if "-" in token:
                parts = token.split("-")
                try:
                    d = int(parts[0])
                    slots = parse_code_to_slots(parts[1], shift_map)
                    for s in slots:
                        allowed_slots.add((d, s))
                except ValueError:
                    pass
        for d in all_days:
            for s in ["AM", "PM", "NIGHT"]:
                if (d, s) not in allowed_slots:
                    extra_unavails.add((d, s))
        return fixed_tasks, extra_unavails

    tokens = [t.strip() for t in text.split(",") if t.strip()]
    for token in tokens:
        if ":" in token:
            key_part, task = token.split(":", 1)
            parts = key_part.strip().split("-")
            if len(parts) == 2:
                prefix = parts[0].strip().upper()
                raw_code = parts[1].strip().upper()
                slots = parse_code_to_slots(raw_code, shift_map)
                slot = slots[0] if slots else raw_code

                if prefix.startswith("W"):
                    try:
                        target_w = int(prefix[1:]) - 1
                        for d in all_days:
                            if datetime.date(year, month, d).weekday() == target_w:
                                fixed_tasks[(d, slot)] = (raw_code, task.strip())
                    except ValueError:
                        pass
                else:
                    try:
                        day = int(prefix)
                        fixed_tasks[(day, slot)] = (raw_code, task.strip())
                    except ValueError:
                        pass
    return fixed_tasks, extra_unavails

# ==========================================
# 3. 排班運籌求解引擎 (OR-Tools CP-SAT)
# ==========================================
class AdvancedClinicScheduler:
    def __init__(self, year: int, month: int, staff_data: List[StaffInput], shift_defs: List[ShiftDef], clinic_closures: Set[Tuple[int, str]]):
        self.year = year
        self.month = month
        self.staff_data = staff_data
        self.shift_defs = shift_defs
        self.shift_map = {s.code: s for s in shift_defs}
        self.clinic_closures = clinic_closures

        if month == 12:
            num_days = 31
        else:
            num_days = (datetime.date(year, month + 1, 1) - datetime.date(year, month, 1)).days
        self.days = list(range(1, num_days + 1))
        self.slots = ["AM", "PM", "NIGHT"]

    def _get_slot_demand(self, day: int) -> Dict[str, int]:
        weekday = datetime.date(self.year, self.month, day).weekday()
        if weekday < 5:
            base = {"AM": 3, "PM": 2, "NIGHT": 3}
        elif weekday == 5:
            base = {"AM": 3, "PM": 0, "NIGHT": 0}
        else:
            base = {"AM": 0, "PM": 0, "NIGHT": 0}

        for slot in self.slots:
            if (day, slot) in self.clinic_closures:
                base[slot] = 0
        return base

    def solve(self):
        model = cp_model.CpModel()
        x = {}
        for s in self.staff_data:
            for d in self.days:
                for slot in self.slots:
                    x[(s.id, d, slot)] = model.NewBoolVar(f"x_{s.id}_d{d}_{slot}")

        is_A1 = {}
        is_C1 = {}
        for s in self.staff_data:
            for d in self.days:
                is_A1[(s.id, d)] = model.NewBoolVar(f"isA1_{s.id}_d{d}")
                is_C1[(s.id, d)] = model.NewBoolVar(f"isC1_{s.id}_d{d}")
                model.Add(is_A1[(s.id, d)] <= x[(s.id, d, "AM")])
                model.Add(is_C1[(s.id, d)] <= x[(s.id, d, "NIGHT")])
                # 防呆限制：午班B (13:30-17:30) 與 提早晚班C1 (16:30-21:30) 重疊，禁止 BC1 組合
                model.Add(x[(s.id, d, "PM")] + is_C1[(s.id, d)] <= 1)

        # 出勤天數變數 (用於集中排班)
        worked_day = {}
        for s in self.staff_data:
            for d in self.days:
                worked_day[(s.id, d)] = model.NewBoolVar(f"worked_{s.id}_d{d}")
                for slot in self.slots:
                    model.Add(x[(s.id, d, slot)] <= worked_day[(s.id, d)])
                model.Add(
                    worked_day[(s.id, d)]
                    <= sum(x[(s.id, d, slot)] for slot in self.slots)
                )

        # 純 AC 班定義（AM=1, NIGHT=1 且 PM=0，ABC全日班不算）
        is_pure_ac = {}
        for s in self.staff_data:
            for d in self.days:
                is_pure_ac[(s.id, d)] = model.NewBoolVar(f"pure_ac_{s.id}_d{d}")
                model.Add(is_pure_ac[(s.id, d)] <= x[(s.id, d, "AM")])
                model.Add(is_pure_ac[(s.id, d)] <= x[(s.id, d, "NIGHT")])
                model.Add(is_pure_ac[(s.id, d)] <= 1 - x[(s.id, d, "PM")])
                model.Add(
                    is_pure_ac[(s.id, d)]
                    >= x[(s.id, d, "AM")]
                    + x[(s.id, d, "NIGHT")]
                    - x[(s.id, d, "PM")]
                    - 1
                )

        # ----------------------------------------------------
        # 1. 班別人力短缺補償 (Slack Variables)
        # ----------------------------------------------------
        shortage = {}
        for d in self.days:
            demand = self._get_slot_demand(d)
            for slot in self.slots:
                shortage[(d, slot)] = model.NewIntVar(0, demand[slot], f"short_{d}_{slot}")
                model.Add(
                    sum(x[(s.id, d, slot)] for s in self.staff_data)
                    + shortage[(d, slot)]
                    == demand[slot]
                )

        # ----------------------------------------------------
        # 修正規則：
        # 平日 (W1-W5)：精準只需要 1 個 A1 班與 1 個 C1 班
        # 週六 (W6)：早診 3 人精準為 2 個 A1 班與 1 個 A 班
        # ----------------------------------------------------
        shortage_A1 = {}
        shortage_C1 = {}
        for d in self.days:
            demand = self._get_slot_demand(d)
            weekday = datetime.date(self.year, self.month, d).weekday()

            if weekday < 5 and demand["AM"] > 0:
                # 平日 AM：剛好 1 個 A1 (其餘 2 人為 A)
                shortage_A1[d] = model.NewIntVar(0, 1, f"short_A1_{d}")
                model.Add(sum(is_A1[(s.id, d)] for s in self.staff_data) + shortage_A1[d] == 1)
            elif weekday == 5 and demand["AM"] > 0:
                # 週六 AM：剛好 2 個 A1 (其餘 1 人為 A，共 3 人)
                shortage_A1[d] = model.NewIntVar(0, 2, f"short_A1_{d}")
                model.Add(sum(is_A1[(s.id, d)] for s in self.staff_data) + shortage_A1[d] == 2)
            else:
                shortage_A1[d] = model.NewIntVar(0, 0, f"short_A1_{d}")

            if weekday < 5 and demand["NIGHT"] > 0:
                # 平日 NIGHT：剛好 1 個 C1 (其餘 2 人為 C)
                shortage_C1[d] = model.NewIntVar(0, 1, f"short_C1_{d}")
                model.Add(sum(is_C1[(s.id, d)] for s in self.staff_data) + shortage_C1[d] == 1)
            else:
                shortage_C1[d] = model.NewIntVar(0, 0, f"short_C1_{d}")

        senior_staff = [s for s in self.staff_data if s.is_senior]
        shortage_senior_A = {}
        shortage_senior_C = {}
        for d in self.days:
            demand = self._get_slot_demand(d)
            if demand["AM"] > 0 and senior_staff:
                shortage_senior_A[d] = model.NewIntVar(0, 1, f"short_senior_A_{d}")
                model.Add(sum(x[(s.id, d, "AM")] for s in senior_staff) + shortage_senior_A[d] >= 1)
            else:
                shortage_senior_A[d] = model.NewIntVar(0, 0, f"short_senior_A_{d}")

            if demand["NIGHT"] > 0 and senior_staff:
                shortage_senior_C[d] = model.NewIntVar(0, 1, f"short_senior_C_{d}")
                model.Add(sum(x[(s.id, d, "NIGHT")] for s in senior_staff) + shortage_senior_C[d] >= 1)
            else:
                shortage_senior_C[d] = model.NewIntVar(0, 0, f"short_senior_C_{d}")

        # 班別數量硬約束 (AC 班上限 & C 班上限)
        for s in self.staff_data:
            model.Add(sum(is_pure_ac[(s.id, d)] for d in self.days) <= s.max_ac_shifts)
            model.Add(sum(x[(s.id, d, "NIGHT")] for d in self.days) <= s.max_c_shifts)

        # 每週純 AC 班頻率壓制 (每週最多 1 次)
        weeks_dict = {}
        for d in self.days:
            dt = datetime.date(self.year, self.month, d)
            w_num = dt.isocalendar()[1]
            weeks_dict.setdefault(w_num, []).append(d)

        excess_ac_weekly_vars = []
        for s in self.staff_data:
            for w_num, w_days in weeks_dict.items():
                total_ac_this_week = sum(is_pure_ac[(s.id, d)] for d in w_days)
                excess = model.NewIntVar(0, len(w_days), f"excess_ac_w_{s.id}_w{w_num}")
                model.Add(excess >= total_ac_this_week - 1)
                excess_ac_weekly_vars.append(excess)

        # 每週 Pattern 慣性
        pattern_vars = {}
        pattern_devs = []
        for s in self.staff_data:
            for w_idx in range(6):  # 0=週一 ... 5=週六
                for slot in self.slots:
                    p_var = model.NewBoolVar(f"pat_{s.id}_w{w_idx}_{slot}")
                    pattern_vars[(s.id, w_idx, slot)] = p_var

                    matching_days = [
                        d for d in self.days
                        if datetime.date(self.year, self.month, d).weekday() == w_idx
                    ]
                    for d in matching_days:
                        diff = model.NewIntVar(0, 1, f"patdiff_{s.id}_d{d}_{slot}")
                        model.Add(diff >= x[(s.id, d, slot)] - p_var)
                        model.Add(diff >= p_var - x[(s.id, d, slot)])
                        pattern_devs.append(diff)

        hours_A = self.shift_map["A"].hours if "A" in self.shift_map else 4.0
        hours_A1 = self.shift_map["A1"].hours if "A1" in self.shift_map else 5.0
        hours_B = self.shift_map["B"].hours if "B" in self.shift_map else 4.0
        hours_C = self.shift_map["C"].hours if "C" in self.shift_map else 4.0
        hours_C1 = self.shift_map["C1"].hours if "C1" in self.shift_map else 5.0

        hA_val = int(hours_A * 10)
        diff_A1_val = int((hours_A1 - hours_A) * 10)
        hB_val = int(hours_B * 10)
        hC_val = int(hours_C * 10)
        diff_C1_val = int((hours_C1 - hours_C) * 10)

        deviations = []
        min_hour_shortage_vars = []

        for s in self.staff_data:
            unavails = parse_unavailabilities(s.unavail_str, self.shift_map)
            blocked_weekdays = parse_blocked_weekdays(s.blocked_weekday_str, self.shift_map)
            fixed_tasks, only_unavails = parse_fixed_and_only(
                s.fixed_or_only_str, self.days, self.year, self.month, self.shift_map
            )
            all_blocked = unavails.union(only_unavails)
            daily_max_val = int(s.daily_max_hours * 10)

            month_hour_terms = []

            for d in self.days:
                weekday = datetime.date(self.year, self.month, d).weekday()
                daily_hour_terms = []

                for slot in self.slots:
                    if (d, slot) in fixed_tasks:
                        model.Add(x[(s.id, d, slot)] == 0)
                    else:
                        if (d, slot) in all_blocked:
                            model.Add(x[(s.id, d, slot)] == 0)
                        if (weekday, slot) in blocked_weekdays:
                            model.Add(x[(s.id, d, slot)] == 0)

                    if slot == "AM":
                        daily_hour_terms.append(x[(s.id, d, "AM")] * hA_val + is_A1[(s.id, d)] * diff_A1_val)
                    elif slot == "PM":
                        daily_hour_terms.append(x[(s.id, d, "PM")] * hB_val)
                    elif slot == "NIGHT":
                        daily_hour_terms.append(x[(s.id, d, "NIGHT")] * hC_val + is_C1[(s.id, d)] * diff_C1_val)

                # 僅以單日上限工時約束 (允許排入 ABC 全日班 4+4+4=12hr)
                if daily_hour_terms:
                    model.Add(sum(daily_hour_terms) <= daily_max_val)
                    month_hour_terms.extend(daily_hour_terms)

            month_hours_expr = sum(month_hour_terms) if month_hour_terms else 0

            min_slack = model.NewIntVar(0, int(s.min_hours * 10), f"min_slack_{s.id}")
            min_hour_shortage_vars.append(min_slack)
            model.Add(month_hours_expr + min_slack >= int(s.min_hours * 10))
            model.Add(month_hours_expr <= int(s.max_hours * 10))

            dev = model.NewIntVar(0, 2000, f"dev_{s.id}")
            model.AddAbsEquality(dev, month_hours_expr - int(s.target_hours * 10))
            deviations.append(dev)

        # 最佳化目標
        total_shortage_penalty = (
            sum(shortage[(d, slot)] * 20000 for d in self.days for slot in self.slots)
            + sum(shortage_A1[d] * 10000 for d in self.days)
            + sum(shortage_C1[d] * 10000 for d in self.days)
            + sum(shortage_senior_A[d] * 5000 for d in self.days)
            + sum(shortage_senior_C[d] * 5000 for d in self.days)
        )
        total_min_hour_penalty = sum(min_hour_shortage_vars) * 500
        total_excess_ac_weekly_penalty = sum(excess_ac_weekly_vars) * 500
        total_pattern_penalty = sum(pattern_devs) * 25
        total_working_days_penalty = sum(worked_day[(s.id, d)] for s in self.staff_data for d in self.days) * 80

        model.Minimize(
            total_shortage_penalty
            + total_min_hour_penalty
            + total_excess_ac_weekly_penalty
            + total_pattern_penalty
            + total_working_days_penalty
            + sum(deviations)
        )

        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 25.0
        status = solver.Solve(model)

        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            return self._format_results(
                solver, x, is_A1, is_C1, is_pure_ac, shortage, shortage_A1, shortage_C1, shortage_senior_A, shortage_senior_C
            )
        return None, None

    def _format_results(self, solver, x, is_A1, is_C1, is_pure_ac, shortage, shortage_A1, shortage_C1, shortage_senior_A, shortage_senior_C):
        rows = []
        total_days_in_month = len(self.days)

        # Row 0: 置頂「星期別」專用列
        weekday_row = {"員工編號": "", "姓名": "星期", "資深": "", "上限": ""}
        for d in self.days:
            wk_str = ["一", "二", "三", "四", "五", "六", "日"][datetime.date(self.year, self.month, d).weekday()]
            weekday_row[f"{d}日"] = f"週{wk_str}"
        weekday_row["實排時數"] = ""
        weekday_row["目標時數"] = ""
        weekday_row["差額"] = ""
        weekday_row["上班天數"] = ""
        weekday_row["Off天數"] = ""
        weekday_row["純AC次數"] = ""
        weekday_row["C班次數"] = ""
        rows.append(weekday_row)

        # Row 1: 置頂「診所缺額標示」專用列
        shortage_summary_row = {"員工編號": "", "姓名": "⚠️ 診所缺額", "資深": "", "上限": ""}
        for d in self.days:
            day_shortages = []
            for slot, s_name in [("AM", "早"), ("PM", "午"), ("NIGHT", "晚")]:
                cnt = solver.Value(shortage[(d, slot)])
                if cnt > 0:
                    day_shortages.append(f"缺{s_name}{cnt}人")

            if solver.Value(shortage_A1[d]) > 0:
                day_shortages.append("缺A1")
            if solver.Value(shortage_C1[d]) > 0:
                day_shortages.append("缺C1")
            if solver.Value(shortage_senior_A[d]) > 0 or solver.Value(shortage_senior_C[d]) > 0:
                day_shortages.append("缺資深")

            if day_shortages:
                shortage_summary_row[f"{d}日"] = "[" + ",".join(day_shortages) + "]"
            else:
                shortage_summary_row[f"{d}日"] = "-"

        shortage_summary_row["實排時數"] = ""
        shortage_summary_row["目標時數"] = ""
        shortage_summary_row["差額"] = ""
        shortage_summary_row["上班天數"] = ""
        shortage_summary_row["Off天數"] = ""
        shortage_summary_row["純AC次數"] = ""
        shortage_summary_row["C班次數"] = ""
        rows.append(shortage_summary_row)

        # 員工排班明細
        for s in self.staff_data:
            fixed_tasks, _ = parse_fixed_and_only(s.fixed_or_only_str, self.days, self.year, self.month, self.shift_map)
            row = {"員工編號": s.id, "姓名": s.name, "資深": "是" if s.is_senior else "否", "上限": f"{s.daily_max_hours}h"}
            tot_hrs = 0.0
            worked_days_count = 0

            for d in self.days:
                active_codes = []
                is_working_today = False

                # AM 時段
                if (d, "AM") in fixed_tasks:
                    is_working_today = True
                    raw_c, task = fixed_tasks[(d, "AM")]
                    active_codes.append(f"({raw_c}:{task})")
                elif solver.Value(x[(s.id, d, "AM")]) == 1:
                    is_working_today = True
                    if solver.Value(is_A1[(s.id, d)]) == 1:
                        active_codes.append("A1")
                        tot_hrs += self.shift_map.get("A1", ShiftDef("","","", "", 5.0)).hours
                    else:
                        active_codes.append("A")
                        tot_hrs += self.shift_map.get("A", ShiftDef("","","", "", 4.0)).hours

                # PM 時段
                if (d, "PM") in fixed_tasks:
                    is_working_today = True
                    raw_c, task = fixed_tasks[(d, "PM")]
                    active_codes.append(f"({raw_c}:{task})")
                elif solver.Value(x[(s.id, d, "PM")]) == 1:
                    is_working_today = True
                    active_codes.append("B")
                    tot_hrs += self.shift_map.get("B", ShiftDef("","","", "", 4.0)).hours

                # NIGHT 時段
                if (d, "NIGHT") in fixed_tasks:
                    is_working_today = True
                    raw_c, task = fixed_tasks[(d, "NIGHT")]
                    active_codes.append(f"({raw_c}:{task})")
                elif solver.Value(x[(s.id, d, "NIGHT")]) == 1:
                    is_working_today = True
                    if solver.Value(is_C1[(s.id, d)]) == 1:
                        active_codes.append("C1")
                        tot_hrs += self.shift_map.get("C1", ShiftDef("","","", "", 5.0)).hours
                    else:
                        active_codes.append("C")
                        tot_hrs += self.shift_map.get("C", ShiftDef("","","", "", 4.0)).hours

                if not active_codes:
                    row[f"{d}日"] = "off"
                else:
                    worked_days_count += 1
                    if any("(" in c for c in active_codes):
                        row[f"{d}日"] = "+".join(active_codes)
                    elif all(len(c) <= 2 for c in active_codes):
                        row[f"{d}日"] = "".join(active_codes)
                    else:
                        row[f"{d}日"] = "+".join(active_codes)

            pure_ac_count = sum(1 for d in self.days if solver.Value(is_pure_ac[(s.id, d)]) == 1)
            c_shift_count = sum(1 for d in self.days if solver.Value(x[(s.id, d, "NIGHT")]) == 1)
            off_days_count = total_days_in_month - worked_days_count

            row["實排時數"] = round(tot_hrs, 1)
            row["目標時數"] = s.target_hours
            row["差額"] = round(tot_hrs - s.target_hours, 1)
            row["上班天數"] = worked_days_count
            row["Off天數"] = off_days_count
            row["純AC次數"] = pure_ac_count
            row["C班次數"] = c_shift_count
            rows.append(row)

        df_schedule = pd.DataFrame(rows)

        # 彙整缺工預警表格
        shortage_rows = []
        for d in self.days:
            wk_str = ["一", "二", "三", "四", "五", "六", "日"][datetime.date(self.year, self.month, d).weekday()]
            d_str = f"{self.month}/{d} (週{wk_str})"

            for slot in self.slots:
                cnt = solver.Value(shortage[(d, slot)])
                if cnt > 0:
                    shortage_rows.append({"日期": d_str, "類別": f"{slot} 人力不足", "缺少數量": cnt, "詳細說明": f"該診次缺額 {cnt} 人 (需兼職/臨時支援)"})

            if solver.Value(shortage_A1[d]) > 0:
                shortage_rows.append({"日期": d_str, "類別": "A1 人數不足", "缺少數量": solver.Value(shortage_A1[d]), "詳細說明": f"12:30-13:30 尚缺 {solver.Value(shortage_A1[d])} 人延後下班"})
            if solver.Value(shortage_C1[d]) > 0:
                shortage_rows.append({"日期": d_str, "類別": "C1 人數不足", "缺少數量": 1, "詳細說明": "16:30-21:30 晚診未排入 C1 提早班"})
            if solver.Value(shortage_senior_A[d]) > 0:
                shortage_rows.append({"日期": d_str, "類別": "A班缺資深人員", "缺少數量": 1, "詳細說明": "當日早診無資深人員在場"})
            if solver.Value(shortage_senior_C[d]) > 0:
                shortage_rows.append({"日期": d_str, "類別": "C班缺資深人員", "缺少數量": 1, "詳細說明": "當日晚診無資深人員在場"})

        return df_schedule, pd.DataFrame(shortage_rows)

# ==========================================
# 4. 衝突診斷與數據統計模組
# ==========================================
class ConflictDiagnosticAnalyzer:
    def __init__(self, year: int, month: int, staff_data: List[StaffInput], shift_defs: List[ShiftDef], clinic_closures: Set[Tuple[int, str]]):
        self.year = year
        self.month = month
        self.staff_data = staff_data
        self.shift_defs = shift_defs
        self.shift_map = {s.code: s for s in shift_defs}
        self.clinic_closures = clinic_closures

        if month == 12:
            num_days = 31
        else:
            num_days = (datetime.date(year, month + 1, 1) - datetime.date(year, month, 1)).days
        self.days = list(range(1, num_days + 1))
        self.slots = ["AM", "PM", "NIGHT"]

    def get_summary_statistics(self) -> Tuple[float, float]:
        """計算 (當月總人力時數需求, 當月同仁最低工時總和)"""
        total_demand_hours = 0.0
        for d in self.days:
            dt = datetime.date(self.year, self.month, d)
            w = dt.weekday()
            if w == 6:
                continue
            elif w == 5:
                # 週六：2個A1 (5h) + 1個A (4h) = 14h
                if (d, "AM") not in self.clinic_closures:
                    total_demand_hours += 14.0
            else:
                # 平日：1個A1(5h)+2個A(8h)=13h, 2個B=8h, 1個C1(5h)+2個C(8h)=13h -> 34h
                if (d, "AM") not in self.clinic_closures: total_demand_hours += 13.0
                if (d, "PM") not in self.clinic_closures: total_demand_hours += 8.0
                if (d, "NIGHT") not in self.clinic_closures: total_demand_hours += 13.0

        total_min_hours = sum(s.min_hours for s in self.staff_data)
        return total_demand_hours, float(total_min_hours)

# ==========================================
# 5. 專業 Excel 美化匯出模組 (包含 5 列示範結構與圖示)
# ==========================================
def export_to_excel(year: int, month: int, df_schedule: pd.DataFrame, df_shortage: pd.DataFrame, staff_inputs: List[StaffInput]) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{year}{month:02d}"

    font_header_title = Font(name="新細明體", size=10, bold=True)
    font_day_header = Font(name="新細明體", size=9)
    font_bold = Font(name="新細明體", size=9, bold=True)
    font_body = Font(name="新細明體", size=9)
    font_symbol = Font(name="新細明體", size=9)
    font_shortage = Font(name="微軟正黑體", size=9, bold=True, color="C00000")

    fill_shortage = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
    fill_weekend = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")
    fill_header = PatternFill(start_color="E6EEF8", end_color="E6EEF8", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )

    # 取得天數清單
    days_cols = [c for c in df_schedule.columns if "日" in str(c)]
    days = [int(str(c).replace("日", "")) for c in days_cols]

    # ------------------------------------------
    # 頂部第 1~3 列：日期、星期、工作日代號
    # ------------------------------------------
    ws.cell(row=1, column=1, value=str(year)).font = font_header_title
    ws.cell(row=1, column=2, value=float(month)).font = font_header_title
    col_idx = 3
    for d in days:
        dt = datetime.date(year, month, d)
        wk = dt.weekday()
        wk_str = ["一", "二", "三", "四", "五", "六", "日"][wk]
        w_code = "W" if wk < 5 else ("S" if wk == 5 else "H")

        c1 = ws.cell(row=1, column=col_idx, value=d)
        c2 = ws.cell(row=2, column=col_idx, value=wk_str)
        c3 = ws.cell(row=3, column=col_idx, value=w_code)

        for c in (c1, c2, c3):
            c.font = font_day_header
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin_border
            if wk >= 5: c.fill = fill_weekend
            else: c.fill = fill_header
        col_idx += 1

    # 右側統計大表頭
    stat_start_col = col_idx
    ws.cell(row=1, column=stat_start_col, value="應上").font = font_bold
    ws.cell(row=1, column=stat_start_col+1, value=160).font = font_bold
    ws.cell(row=2, column=stat_start_col, value="應休").font = font_bold
    ws.cell(row=2, column=stat_start_col+1, value=len([d for d in days if datetime.date(year, month, d).weekday() >= 5])).font = font_bold

    # ------------------------------------------
    # 第 4 列：診所缺額標記列
    # ------------------------------------------
    cur_row = 4
    shortage_row_df = df_schedule[df_schedule["姓名"] == "⚠️ 診所缺額"]
    if not shortage_row_df.empty:
        r_data = shortage_row_df.iloc[0]
        ws.cell(row=cur_row, column=1, value="⚠️ 缺額").font = font_shortage
        ws.cell(row=cur_row, column=2, value="診所人力").font = font_shortage
        c_idx = 3
        for d in days:
            val = str(r_data[f"{d}日"])
            cell = ws.cell(row=cur_row, column=c_idx, value=val)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
            if val != "-":
                cell.font = font_shortage
                cell.fill = fill_shortage
            else:
                cell.font = font_body
            c_idx += 1
        cur_row += 1

    # ------------------------------------------
    # 第 5 列起：每位同仁標準 5 列展開 (含時數、代碼、三時段圖示與右側統計)
    # ------------------------------------------
    staff_df = df_schedule[(df_schedule["姓名"] != "星期") & (df_schedule["姓名"] != "⚠️ 診所缺額")]
    staff_input_map = {s.name: s for s in staff_inputs}

    for _, s_row in staff_df.iterrows():
        name = str(s_row["姓名"])
        s_obj = staff_input_map.get(name, None)
        emp_id = s_obj.id if s_obj else str(s_row["員工編號"])
        fixed_tasks, _ = parse_fixed_and_only(s_obj.fixed_or_only_str, days, year, month, {}) if s_obj else ({}, set())

        r1 = cur_row
        r2 = cur_row + 1
        r3 = cur_row + 2
        r4 = cur_row + 3
        r5 = cur_row + 4

        # Col 1: 姓名 (置於 r1 與 r2)
        ws.cell(row=r1, column=1, value=name).font = font_bold
        ws.cell(row=r2, column=1, value=name).font = font_bold
        for r_sub in range(r1, r1 + 5):
            ws.cell(row=r_sub, column=1).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=r_sub, column=1).border = thin_border

        # Col 2: 代號與三個時段標題
        ws.cell(row=r1, column=2, value=emp_id).font = font_bold
        ws.cell(row=r2, column=2, value=emp_id).font = font_bold
        ws.cell(row=r3, column=2, value="0830-1230").font = font_body
        ws.cell(row=r4, column=2, value="1330-1730").font = font_body
        ws.cell(row=r5, column=2, value="1730-2130").font = font_body
        for r_sub in range(r1, r1 + 5):
            ws.cell(row=r_sub, column=2).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=r_sub, column=2).border = thin_border

        # 填入每日資料 (Col 3 ~ 3+len(days)-1)
        allday_count = 0
        c_idx = 3
        for d in days:
            shift_text = str(s_row[f"{d}日"])
            is_off = (shift_text == "off")

            # 計算當日工時與代碼
            daily_hrs = 0.0
            am_sym = "-"
            pm_sym = "-"
            night_sym = "-"

            # AM 判斷
            if (d, "AM") in fixed_tasks:
                raw_c, task = fixed_tasks[(d, "AM")]
                am_sym = task
            elif "A1" in shift_text:
                am_sym = "☼"
                daily_hrs += 5.0
            elif "A" in shift_text and "ALL EXCEPT" not in shift_text:
                am_sym = "○"
                daily_hrs += 4.0

            # PM 判斷
            if (d, "PM") in fixed_tasks:
                raw_c, task = fixed_tasks[(d, "PM")]
                pm_sym = task
            elif "B" in shift_text:
                pm_sym = "○"
                daily_hrs += 4.0

            # NIGHT 判斷
            if (d, "NIGHT") in fixed_tasks:
                raw_c, task = fixed_tasks[(d, "NIGHT")]
                night_sym = task
            elif "C1" in shift_text:
                night_sym = "☼"
                daily_hrs += 5.0
            elif "C" in shift_text and "ALL EXCEPT" not in shift_text:
                night_sym = "○"
                daily_hrs += 4.0

            # 班別代碼修飾 (A1 轉為 A*, C1 轉為 C*)
            display_code = shift_text.replace("A1", "A*").replace("C1", "C*")
            if is_off:
                display_code = ""

            if "A" in shift_text and "B" in shift_text and "C" in shift_text:
                allday_count += 1

            # 填入 5 列單元格
            c_r1 = ws.cell(row=r1, column=c_idx, value=int(daily_hrs) if daily_hrs > 0 else (None if is_off else 0))
            c_r2 = ws.cell(row=r2, column=c_idx, value=display_code if display_code else None)
            c_r3 = ws.cell(row=r3, column=c_idx, value=am_sym if not is_off else "-")
            c_r4 = ws.cell(row=r4, column=c_idx, value=pm_sym if not is_off else "-")
            c_r5 = ws.cell(row=r5, column=c_idx, value=night_sym if not is_off else "-")

            dt = datetime.date(year, month, d)
            is_wkend = (dt.weekday() >= 5)

            for c_cell in (c_r1, c_r2, c_r3, c_r4, c_r5):
                c_cell.font = font_symbol
                c_cell.alignment = Alignment(horizontal="center", vertical="center")
                c_cell.border = thin_border
                if is_wkend:
                    c_cell.fill = fill_weekend

            c_idx += 1

        # ------------------------------------------
        # 右側標準統計欄位 (精確還原示範檔)
        # ------------------------------------------
        # Row 1 (r1): 上月, 固定排班
        ws.cell(row=r1, column=stat_start_col, value="上月").font = font_body
        ws.cell(row=r1, column=stat_start_col+1, value=0.0).font = font_bold
        ws.cell(row=r1, column=stat_start_col+2, value="").font = font_body
        ws.cell(row=r1, column=stat_start_col+3, value="固定排班").font = font_body

        # Row 2 (r2): 應上, 應上工時, Off, Off天數
        ws.cell(row=r2, column=stat_start_col, value="應上").font = font_body
        ws.cell(row=r2, column=stat_start_col+1, value=float(s_row["目標時數"])).font = font_bold
        ws.cell(row=r2, column=stat_start_col+2, value="Off").font = font_body
        ws.cell(row=r2, column=stat_start_col+3, value=int(s_row["Off天數"])).font = font_bold

        # Row 3 (r3): 本月, 實排工時, AC, AC次數
        ws.cell(row=r3, column=stat_start_col, value="本月").font = font_body
        ws.cell(row=r3, column=stat_start_col+1, value=float(s_row["實排時數"])).font = font_bold
        ws.cell(row=r3, column=stat_start_col+2, value="AC").font = font_body
        ws.cell(row=r3, column=stat_start_col+3, value=int(s_row["純AC次數"])).font = font_bold

        # Row 4 (r4): 結餘, 差額時數, Allday, 全日班次數
        ws.cell(row=r4, column=stat_start_col, value="結餘").font = font_body
        ws.cell(row=r4, column=stat_start_col+1, value=float(s_row["差額"])).font = font_bold
        ws.cell(row=r4, column=stat_start_col+2, value="Allday").font = font_body
        ws.cell(row=r4, column=stat_start_col+3, value=allday_count).font = font_bold

        # Row 5 (r5): 特休, 0, C班統計, C班次數
        ws.cell(row=r5, column=stat_start_col, value="特休").font = font_body
        ws.cell(row=r5, column=stat_start_col+1, value=0.0).font = font_bold
        ws.cell(row=r5, column=stat_start_col+2, value="C班統計").font = font_body
        ws.cell(row=r5, column=stat_start_col+3, value=int(s_row["C班次數"])).font = font_bold

        for r_sub in range(r1, r1 + 5):
            for sc in range(stat_start_col, stat_start_col + 4):
                ws.cell(row=r_sub, column=sc).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(row=r_sub, column=sc).border = thin_border

        cur_row += 5

    # 自動調整欄寬
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 11
    for col_i in range(3, stat_start_col):
        col_letter = get_column_letter(col_i)
        ws.column_dimensions[col_letter].width = 4.2
    ws.column_dimensions[get_column_letter(stat_start_col)].width = 7.0
    ws.column_dimensions[get_column_letter(stat_start_col+1)].width = 7.5
    ws.column_dimensions[get_column_letter(stat_start_col+2)].width = 8.5
    ws.column_dimensions[get_column_letter(stat_start_col+3)].width = 9.0

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# 6. Streamlit 網頁主程式
# ==========================================
def main():
    st.set_page_config(page_title="診所自動排班系統 v4.7", page_icon="🏥", layout="wide")
    st.title("🏥 診所自動排班系統 (v4.7 專業圖示旗艦版)")

    # 側邊欄設定
    st.sidebar.header("🗓️ 排班月份與公休設定")
    year = st.sidebar.number_input("年份 (西元)", min_value=2024, max_value=2035, value=2026)
    month = st.sidebar.number_input("月份", min_value=1, max_value=12, value=11)

    clinic_closure_str = st.sidebar.text_input(
        "🏥 診所公休 / 停診時段",
        value="",
        help="輸入全診公休或特定時段停診。例如：'28' (28日全休), '28-ABC, 17-B' (17日下午停診)。"
    )
    clinic_closures = parse_clinic_closures(clinic_closure_str)

    st.sidebar.markdown("---")
    st.sidebar.markdown(
        "**💡 快捷規則說明：**\n"
        "- **時段配置**：平日 1個A1、1個C1；週六 2個A1、1個A。\n"
        "- **Excel 匯出**：生成標準 5 列圖示 (☼ 加長班, ○ 一般班, - 休假)。\n"
        "- **連續休假區間**：預假可直接填 `11-14`、`18-23`。"
    )

    # 區塊 1: 班別設定
    st.subheader("⚙️ 1. 班別設定 (包含 A1 / C1 特殊班別)")
    default_shifts = [
        {"班別代碼": "A",  "班別名稱": "早班",             "時段範圍": "08:30-12:30", "對應時段": "AM",    "計算工時": 4.0, "預設班別": True},
        {"班別代碼": "A1", "班別名稱": "早班延後(涵蓋13:30)", "時段範圍": "08:30-13:30", "對應時段": "AM",    "計算工時": 5.0, "預設班別": False},
        {"班別代碼": "B",  "班別名稱": "午班",             "時段範圍": "13:30-17:30", "對應時段": "PM",    "計算工時": 4.0, "預設班別": True},
        {"班別代碼": "C",  "班別名稱": "晚班",             "時段範圍": "17:30-21:30", "對應時段": "NIGHT", "計算工時": 4.0, "預設班別": True},
        {"班別代碼": "C1", "班別名稱": "晚班提早(涵蓋14:30)", "時段範圍": "16:30-21:30", "對應時段": "NIGHT", "計算工時": 5.0, "預設班別": False},
    ]
    edited_shifts = st.data_editor(pd.DataFrame(default_shifts), num_rows="dynamic", use_container_width=True)

    shift_defs = []
    for _, r in edited_shifts.iterrows():
        if pd.notna(r["班別代碼"]):
            shift_defs.append(ShiftDef(
                code=str(r["班別代碼"]).strip().upper(),
                name=str(r["班別名稱"]),
                time_range=str(r["時段範圍"]),
                slot=str(r["對應時段"]).strip().upper(),
                hours=float(r["計算工時"]),
                is_default=bool(r["預設班別"])
            ))

    st.markdown("---")

    # 區塊 2: 員工名單與個人排班條件
    st.subheader("👥 2. 員工名單與個人排班條件")

    default_staff = [
        {"員工編號": "PT001", "姓名": "羅昌智", "資深同仁": True,  "每日最高時數": 10, "AC班上限": 5, "C班上限": 9, "目標時數": 120, "最低時數": 115, "最高時數": 140, "預假日期": "", "特定星期不排": "W1-B, W2-BC, W3-BC, W4-AB, W5-BC", "固定任務 / 兼職僅排指定洞": "17-B:內訓"},
        {"員工編號": "PT002", "姓名": "林姵君", "資深同仁": True,  "每日最高時數": 10, "AC班上限": 4, "C班上限": 9,  "目標時數": 120, "最低時數": 52,  "最高時數": 140, "預假日期": "", "特定星期不排": "W4-C", "固定任務 / 兼職僅排指定洞": "17-B:內訓, W1-A:公司"},
        {"員工編號": "PT003", "姓名": "丁雅婷", "資深同仁": True,  "每日最高時數": 10, "AC班上限": 4, "C班上限": 9, "目標時數": 120, "最低時數": 115, "最高時數": 140, "預假日期": "5, 6", "特定星期不排": "W3-C, W4-AB", "固定任務 / 兼職僅排指定洞": "17-B:內訓"},
        {"員工編號": "PT004", "姓名": "江明蒨", "資深同仁": False, "每日最高時數": 10, "AC班上限": 4, "C班上限": 12, "目標時數": 168, "最低時數": 165, "最高時數": 175, "預假日期": "", "特定星期不排": "", "固定任務 / 兼職僅排指定洞": "17-B:內訓"},
        {"員工編號": "PT005", "姓名": "周佩璇", "資深同仁": False, "每日最高時數": 10, "AC班上限": 0, "C班上限": 20,  "目標時數": 84,  "最低時數": 84,  "最高時數": 100, "預假日期": "2, 18", "特定星期不排": "W2-A, W4, W6", "固定任務 / 兼職僅排指定洞": "17-B:內訓"},
        {"員工編號": "PT006", "姓名": "尚恩典", "資深同仁": False, "每日最高時數": 10, "AC班上限": 0, "C班上限": 20,  "目標時數": 84,  "最低時數": 75,  "最高時數": 100, "預假日期": "11-13", "特定星期不排": "W1, W2, W3", "固定任務 / 兼職僅排指定洞": "17-B:內訓"},
        {"員工編號": "PT007", "姓名": "姚欣妤", "資深同仁": False, "每日最高時數": 9,  "AC班上限": 0, "C班上限": 20,  "目標時數": 84,  "最低時數": 75,  "最高時數": 100, "預假日期": "", "特定星期不排": "W1-BC, W2-BC, W3-BC, W4-BC", "固定任務 / 兼職僅排指定洞": "17-B:內訓"},
        {"員工編號": "PT008", "姓名": "張又千", "資深同仁": False, "每日最高時數": 5,  "AC班上限": 0, "C班上限": 20,  "目標時數": 5,   "最低時數": 4,   "最高時數": 5,   "預假日期": "", "特定星期不排": "", "固定任務 / 兼職僅排指定洞": "all except: 10-A"},
        {"員工編號": "PT009", "姓名": "吳汶儒", "資深同仁": False, "每日最高時數": 5,  "AC班上限": 0, "C班上限": 20,  "目標時數": 10,  "最低時數": 4,   "最高時數": 30,  "預假日期": "", "特定星期不排": "all except: W2-B, W3-A", "固定任務 / 兼職僅排指定洞": ""},
        {"員工編號": "PT010", "姓名": "施芸安", "資深同仁": False, "每日最高時數": 10, "AC班上限": 0, "C班上限": 20,  "目標時數": 50,  "最低時數": 45,  "最高時數": 75,  "預假日期": "12, 13-AB, 14, 17-AC, 18-23, 24-AB, 27-30", "特定星期不排": "W1-BC, W3-A, W4-A", "固定任務 / 兼職僅排指定洞": ""},
    ]

    edited_staff = st.data_editor(
        pd.DataFrame(default_staff),
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "資深同仁": st.column_config.CheckboxColumn("資深同仁", default=False),
            "每日最高時數": st.column_config.NumberColumn("單日上限(hr)", min_value=4, max_value=12, step=1),
            "AC班上限": st.column_config.NumberColumn("AC班上限", min_value=0, max_value=15, step=1, help="半職或兼職建議設為 0 以強烈避開花班"),
            "C班上限": st.column_config.NumberColumn("C班上限", min_value=0, max_value=25, step=1, help="無法上晚班者請設為 0")
        }
    )

    current_staff_inputs = []
    for _, r in edited_staff.iterrows():
        if pd.notna(r["姓名"]) and str(r["姓名"]).strip():
            current_staff_inputs.append(StaffInput(
                id=str(r["員工編號"]),
                name=str(r["姓名"]).strip(),
                is_senior=bool(r["資深同仁"]),
                daily_max_hours=int(r["每日最高時數"]),
                max_ac_shifts=int(r.get("AC班上限", 4)),
                max_c_shifts=int(r.get("C班上限", 10)),
                target_hours=int(r["目標時數"]),
                min_hours=int(r["最低時數"]),
                max_hours=int(r["最高時數"]),
                unavail_str=str(r.get("預假日期", "")),
                blocked_weekday_str=str(r.get("特定星期不排", "")),
                fixed_or_only_str=str(r.get("固定任務 / 兼職僅排指定洞", ""))
            ))

    analyzer = ConflictDiagnosticAnalyzer(
        year=year,
        month=month,
        staff_data=current_staff_inputs,
        shift_defs=shift_defs,
        clinic_closures=clinic_closures
    )
    demand_hrs, min_hrs = analyzer.get_summary_statistics()

    # 區塊 3: 求解與下載
    if st.button("🚀 開始自動排班 (求解班表)", type="primary", key="btn_run_scheduler"):
        with st.spinner("高韌性排班引擎運算中 (含週六 2個A1/1個A 與 精確圖示輸出)..."):
            scheduler = AdvancedClinicScheduler(
                year=year,
                month=month,
                staff_data=current_staff_inputs,
                shift_defs=shift_defs,
                clinic_closures=clinic_closures
            )
            df_schedule, df_shortage = scheduler.solve()

        if df_schedule is not None:
            if not df_shortage.empty:
                st.warning("⚠️ 排班已順利生成！系統偵測到下列時段出現人力短缺，已在班表「⚠️ 診所缺額」列清楚標示：")
                st.dataframe(df_shortage, use_container_width=True)
            else:
                st.success(f"🎉 {year} 年 {month} 月 班表生成成功！本月所有診次人力 100% 滿載！")

            st.subheader("📋 班表預覽 (含 星期別、診所缺額標記列、Off天數、純AC次數 與 C班次數)")
            st.dataframe(df_schedule, use_container_width=True)

            excel_buf = export_to_excel(year, month, df_schedule, df_shortage, current_staff_inputs)
            st.download_button(
                label="📥 下載完整班表 Excel 檔案 (.xlsx - 包含 5 列示範結構與 ☼/○ 圖示)",
                data=excel_buf,
                file_name=f"診所排班表_{year}_{month:02d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_download_excel"
            )

    st.markdown("---")

    # ==========================================
    # 區塊 4: 視窗下方常駐顯示「數據指標」與「通用規則列表」
    # ==========================================
    st.subheader("📊 當月總工時數據指標 (即時動態計算)")
    c1, c2, c3 = st.columns(3)
    c1.metric(
        "🏥 當月總人力時數需求",
        f"{demand_hrs:.1f} 小時",
        help="扣除診所公休與週日休診後，平日(AM 13h, PM 8h, NIGHT 13h)與週六(AM 14h)所需之總工時"
    )
    c2.metric(
        "👥 當月同仁最低工時總和",
        f"{min_hrs:.1f} 小時",
        help="員工名單中所有同仁設定之「最低時數」總和"
    )
    diff_hrs = demand_hrs - min_hrs
    c3.metric(
        "⚖️ 工時供需餘裕",
        f"{diff_hrs:+.1f} 小時",
        delta="工時充裕" if diff_hrs >= 0 else "工時透支 (需求不足)",
        delta_color="normal" if diff_hrs >= 0 else "inverse"
    )

    if diff_hrs < 0:
        st.warning(f"💡 提醒：同仁最低工時總和超過了診所總門診需求 {abs(diff_hrs):.1f} 小時，系統會自動在缺工標記列與時數差額中呈現平衡。")

    st.subheader("📜 診所自動排班通用規則列表 (Global System Rules)")
    with st.expander("點擊展開檢視完整系統通用排班規範", expanded=True):
        st.markdown("""
        1. **時段與精準班別配置 (Hard Constraints)**：
           - **週一至週五 (平日)**：
             - 早診 (AM)：配置 **3 人**，精準為 **1 個 A1 班 (5hr)** 與 **2 個 A 班 (4hr)**。
             - 午診 (PM)：配置 **2 人**，標準 B 班 (4hr)。
             - 晚診 (NIGHT)：配置 **3 人**，精準為 **1 個 C1 班 (5hr)** 與 **2 個 C 班 (4hr)**。
           - **週六**：營業至 13:30，早診配置 **3 人**，精準為 **2 個 A1 班 (5hr)** 與 **1 個 A 班 (4hr)**。
           - **週日**：全日診所固定休診。
        2. **Excel 匯出標準圖示 (同仁直觀辨識)**：
           - 每位同仁於 Excel 中依序展開為 5 列：
             1. **當日總工時** (4, 5, 8, 9, 10 等)。
             2. **班別代碼** (A, A*, AB, BC, C*, A*C*, ABC 等)。
             3. **0830-1230 早診圖示**：加長班顯示 **`☼`**，一般班顯示 **`○`**，沒上班顯示 **`-`** (外派顯示「公司」)。
             4. **1330-1730 午診圖示**：上班顯示 **`○`**，沒上班顯示 **`-`**。
             5. **1730-2130 晚診圖示**：加長班顯示 **`☼`**，一般班顯示 **`○`**，沒上班顯示 **`-`** (特殊任務顯示「團課」)。
        3. **高韌性排班原則 (Resilient Scheduling)**：
           - **缺工照常排班**：若遇同仁集中請假或個別診次人力不足，系統**絕對不會中斷報錯**，而是會用現有人力排出全局最優班表，並在「⚠️ 診所缺額」列清楚標示缺額。
        4. **每週固定 Pattern 規律排班**：
           - 演算法在每週的同一星期幾維持穩定的出勤與休假慣性（例如固定週二早、週四晚），方便同仁安排常態私人生活。
        5. **固定任務與公休規範**：
           - **固定任務 (如 17-B:內訓, W1-A:公司)**：僅作為「保留時段」，**不計入門診排班工時**，亦不衝突同仁個人的預假或星期限制，班表格式化顯示為 `(代碼:名稱)`。
           - **連續預假語法**：支援填寫 `11-14`、`18-23` 自動連續排休多日。
           - **診所公休**：側邊欄輸入停診時段（如 `28-ABC, 17-B`），系統自動將門診需求歸零，同仁自動 Off。
        """)

if __name__ == "__main__":
    main()