# -*- coding: utf-8 -*-
"""
校区学员情况管理 Skill - 公共工具函数模块

提供字段映射、冲突格式化、日期处理、Excel样式等公共功能。
被 export_student_xlsx.py / merge_xlsx.py / write_tags.py / generate_html_report.py 共用。

依赖：openpyxl>=3.1.2
"""

import json
import os
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

# ============================================================
# 字段定义：三层架构字段清单
# ============================================================

# A. 基础标识字段（顾问+老师共录）
A_FIELDS: List[str] = ["姓名", "昵称", "年级", "年龄", "所在校区"]

# B. 家庭背景与规划字段（顾问主录）
B_FIELDS: List[str] = ["家长职业", "单位性质", "家庭结构", "教育氛围",
                       "居住小区", "家长规划目标", "对AI认知度"]

# C. 学生在校情况字段（老师主录）
C_FIELDS: List[str] = ["学校名称", "成绩水平", "性格特点",
                       "兴趣偏好", "课堂表现", "同伴关系"]

# D. 在读课程与成果字段（老师主录）
D_FIELDS: List[str] = ["已报名课程", "在读课程", "学习时长",
                       "作品成果", "续费历史"]

# AI补齐层字段
AI_FIELDS: List[str] = ["小区房价段", "住户画像", "周边学校",
                        "学校层次", "是否科技特色校", "周边竞品",
                        "家庭消费力评估"]

# 决策标签字段
TAG_FIELDS: List[str] = ["支付力", "续费风险", "转介绍潜力",
                         "跟进优先级", "推荐产品方向"]

# 顾问版Excel列顺序（A + B）
CONSULTANT_COLUMNS: List[str] = A_FIELDS + B_FIELDS

# 老师版Excel列顺序（A + C + D）
TEACHER_COLUMNS: List[str] = A_FIELDS + C_FIELDS + D_FIELDS

# 汇总表主表列顺序（全部字段）
SUMMARY_COLUMNS: List[str] = [
    "校区", "姓名", "年级", "年龄",
    # B家庭背景
    "家长职业", "单位性质", "家庭结构", "教育氛围", "居住小区",
    "家长规划目标", "对AI认知度",
    # C在校情况
    "学校名称", "成绩水平", "性格特点", "兴趣偏好", "课堂表现", "同伴关系",
    # D课程成果
    "已报名课程", "在读课程", "学习时长", "作品成果", "续费历史",
    # 合并生成
    "冲突标注",
    # AI补齐
    "学校层次(科技特色)", "小区房价段", "住户画像", "周边竞品", "家庭消费力",
    # 决策标签
    "支付力", "续费风险", "转介绍潜力", "跟进优先级", "推荐产品方向",
]

# 冲突清单sheet列
CONFLICT_COLUMNS: List[str] = ["姓名", "年级", "字段名", "顾问值", "老师值", "状态", "备注"]

# 变更记录sheet列
CHANGELOG_COLUMNS: List[str] = ["姓名", "变更时间", "变更字段", "旧值", "新值", "变更来源"]

# 采集完成率sheet列
COMPLETION_COLUMNS: List[str] = ["校区", "采集人", "角色", "名单总数", "已采数", "完成率"]


# ============================================================
# 字段分组映射：字段名 → 所属分组
# ============================================================

FIELD_GROUP_MAP: Dict[str, str] = {}
for f in A_FIELDS:
    FIELD_GROUP_MAP[f] = "A基础标识"
for f in B_FIELDS:
    FIELD_GROUP_MAP[f] = "B家庭背景"
for f in C_FIELDS:
    FIELD_GROUP_MAP[f] = "C在校情况"
for f in D_FIELDS:
    FIELD_GROUP_MAP[f] = "D课程成果"


# ============================================================
# JSON缓存相关函数
# ============================================================

def get_timestamp() -> str:
    """获取当前时间的ISO 8601格式字符串。

    Returns:
        ISO 8601格式时间戳，如 "2026-06-22T10:30:00"
    """
    return datetime.now().strftime("%Y-%m-%dT%H:%M:%S")


def load_json(file_path: str) -> Dict[str, Any]:
    """加载JSON文件。

    Args:
        file_path: JSON文件路径

    Returns:
        解析后的字典；文件不存在或解析失败时返回空字典
    """
    if not os.path.exists(file_path):
        return {}
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError) as e:
        print(f"[警告] 加载JSON失败 {file_path}: {e}")
        return {}


def save_json(file_path: str, data: Dict[str, Any]) -> bool:
    """保存数据为JSON文件（UTF-8编码，缩进2空格）。

    Args:
        file_path: 输出文件路径
        data: 要保存的字典数据

    Returns:
        保存成功返回True，失败返回False
    """
    try:
        os.makedirs(os.path.dirname(file_path), exist_ok=True) if os.path.dirname(file_path) else None
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except IOError as e:
        print(f"[错误] 保存JSON失败 {file_path}: {e}")
        return False


def create_empty_student_record(name: str, campus: str, role: str) -> Dict[str, Any]:
    """创建空的学员档案记录。

    Args:
        name: 学生姓名
        campus: 所在校区
        role: 采集人角色（顾问/老师）

    Returns:
        包含所有字段（空值）的学员档案字典
    """
    record: Dict[str, Any] = {
        "姓名": name,
        "昵称": "",
        "年级": "",
        "年龄": "",
        "所在校区": campus,
        "家庭背景": {f: "" for f in B_FIELDS},
        "在校情况": {f: "" for f in C_FIELDS},
        "课程成果": {f: "" for f in D_FIELDS},
        "AI补齐": None,
        "决策标签": None,
        "采集人角色": role,
        "采集时间": "",
        "采集状态": "未采",
    }
    return record


def create_collection_cache(campus: str, role: str, collector_name: str) -> Dict[str, Any]:
    """创建空的采集端缓存结构。

    Args:
        campus: 校区名
        role: 采集人角色
        collector_name: 采集人姓名

    Returns:
        采集端缓存字典
    """
    return {
        "$schema": "collection_cache_v1",
        "校区": campus,
        "采集人角色": role,
        "采集人姓名": collector_name,
        "名单": [],
        "已采集记录": [],
        "最后更新时间": get_timestamp(),
    }


# ============================================================
# 冲突处理函数
# ============================================================

def format_conflict_cell(consultant_value: str, teacher_value: str) -> str:
    """格式化冲突字段为Excel单元格内双值格式。

    A+C双保留方案：同字段不同来源都保留，标注来源。

    Args:
        consultant_value: 顾问录入的值
        teacher_value: 老师录入的值

    Returns:
        格式化后的字符串，如 "顾问认为:xxx / 老师反馈:yyy [待核实]"
    """
    c_val = consultant_value.strip() if consultant_value else ""
    t_val = teacher_value.strip() if teacher_value else ""
    return f"顾问认为:{c_val} / 老师反馈:{t_val} [待核实]"


def create_conflict_record(name: str, grade: str, field: str,
                           consultant_value: str, teacher_value: str) -> Dict[str, Any]:
    """创建冲突记录。

    Args:
        name: 学生姓名
        grade: 年级
        field: 冲突字段名
        consultant_value: 顾问值
        teacher_value: 老师值

    Returns:
        冲突记录字典
    """
    return {
        "姓名": name,
        "年级": grade,
        "字段名": field,
        "顾问值": consultant_value,
        "老师值": teacher_value,
        "状态": "待核实",
        "备注": "",
    }


def detect_conflict(val1: str, val2: str) -> bool:
    """检测两个字段值是否冲突。

    两个值都非空且不相等时判定为冲突。

    Args:
        val1: 第一个值
        val2: 第二个值

    Returns:
        冲突返回True，否则False
    """
    v1 = (val1 or "").strip()
    v2 = (val2 or "").strip()
    if not v1 or not v2:
        return False
    return v1 != v2


# ============================================================
# 完成率计算函数
# ============================================================

def calculate_completion(cache: Dict[str, Any]) -> Dict[str, Any]:
    """计算采集完成率。

    Args:
        cache: 采集端缓存字典

    Returns:
        包含名单总数、已采数、完成率的字典
    """
    total = len(cache.get("名单", []))
    collected = len(cache.get("已采集记录", []))
    rate = f"{collected}/{total} ({(collected / total * 100):.0f}%)" if total > 0 else "0/0 (0%)"
    return {
        "名单总数": total,
        "已采数": collected,
        "完成率": rate,
    }


# ============================================================
# 姓名匹配函数
# ============================================================

def get_match_key(name: str, grade: str) -> Tuple[str, str]:
    """生成学生匹配键（姓名+年级）。

    唯一标识=中文正式名（完整含姓），重名用年级对齐。

    Args:
        name: 学生姓名
        grade: 年级

    Returns:
        (姓名, 年级) 元组作为匹配键
    """
    return (name.strip() if name else "", grade.strip() if grade else "")


def find_student_in_list(records: List[Dict[str, Any]],
                         name: str, grade: str) -> Optional[Dict[str, Any]]:
    """在记录列表中按姓名+年级查找学生。

    Args:
        records: 学生记录列表
        name: 要查找的姓名
        grade: 要查找的年级

    Returns:
        找到的记录字典，未找到返回None
    """
    target_key = get_match_key(name, grade)
    for record in records:
        rec_key = get_match_key(
            record.get("姓名", ""),
            record.get("年级", ""),
        )
        if rec_key == target_key:
            return record
    return None


# ============================================================
# Excel样式函数（依赖openpyxl）
# ============================================================

def get_excel_styles():
    """获取Excel样式对象。

    延迟导入openpyxl，避免未安装时模块加载失败。

    Returns:
        包含header_fill, header_font, border等样式对象的字典
    """
    try:
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    except ImportError:
        print("[错误] 未安装openpyxl，请执行: pip install openpyxl>=3.1.2")
        raise

    # 表头样式：加粗 + 浅蓝背景 + 居中
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 数据样式
    data_font = Font(name="微软雅黑", size=10)
    data_align = Alignment(vertical="center", wrap_text=True)

    # 边框
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    # 决策标签条件格式颜色
    high_risk_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # 红
    medium_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")     # 黄
    low_risk_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")   # 绿

    return {
        "header_fill": header_fill,
        "header_font": header_font,
        "header_align": header_align,
        "data_font": data_font,
        "data_align": data_align,
        "thin_border": thin_border,
        "high_risk_fill": high_risk_fill,
        "medium_fill": medium_fill,
        "low_risk_fill": low_risk_fill,
    }


def apply_header_style(ws, row: int = 1, col_count: int = 0):
    """应用表头样式到指定工作表。

    Args:
        ws: openpyxl工作表对象
        row: 表头所在行号
        col_count: 列数
    """
    styles = get_excel_styles()
    for col in range(1, col_count + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = styles["header_fill"]
        cell.font = styles["header_font"]
        cell.alignment = styles["header_align"]
        cell.border = styles["thin_border"]


def apply_tag_conditional_format(ws, col_idx: int, row_count: int):
    """对决策标签列应用条件格式（红/黄/绿）。

    Args:
        ws: 工作表对象
        col_idx: 决策标签列序号（1-based）
        row_count: 数据行数
    """
    from openpyxl.formatting.rule import CellIsRule
    styles = get_excel_styles()

    cell_range = f"{ws.cell(row=2, column=col_idx).coordinate}:{ws.cell(row=row_count + 1, column=col_idx).coordinate}"

    # 高=红，中=黄，低=绿
    ws.conditional_formatting.add(cell_range, CellIsRule(
        operator="equal", formula=['"高"'], fill=styles["high_risk_fill"]
    ))
    ws.conditional_formatting.add(cell_range, CellIsRule(
        operator="equal", formula=['"中"'], fill=styles["medium_fill"]
    ))
    ws.conditional_formatting.add(cell_range, CellIsRule(
        operator="equal", formula=['"低"'], fill=styles["low_risk_fill"]
    ))


# ============================================================
# 决策标签推算辅助函数
# ============================================================

def parse_learning_months(learning_duration: str) -> int:
    """从学习时长字符串中解析月份数。

    Args:
        learning_duration: 学习时长字符串，如"8个月""半年""1年"

    Returns:
        月份数（整数），无法解析返回0
    """
    if not learning_duration:
        return 0
    text = learning_duration.strip()
    # 提取数字
    import re
    numbers = re.findall(r'\d+', text)
    if numbers:
        n = int(numbers[0])
        if "年" in text:
            return n * 12
        if "半" in text:
            return 6
        return n
    if "半" in text:
        return 6
    return 0


def parse_renewal_count(renewal_history: str) -> int:
    """从续费历史字符串中解析续费次数。

    Args:
        renewal_history: 续费历史字符串，如"续费1次""续过2次""没续过"

    Returns:
        续费次数（整数），无法解析返回0
    """
    if not renewal_history:
        return 0
    text = renewal_history.strip()
    import re
    numbers = re.findall(r'\d+', text)
    if numbers:
        return int(numbers[0])
    if "没" in text or "未" in text or "无" in text:
        return 0
    if "续" in text:
        return 1  # "续过"但无具体数字，默认1次
    return 0


def calculate_payment_level(occupation: str, housing_price: str,
                            consumption: str, thresholds: Dict[str, str]) -> str:
    """根据配置阈值推算支付力等级。

    阈值为自然语言描述，这里做关键词匹配。
    workbuddy的AI能力可进一步精确理解这些自然语言规则。

    Args:
        occupation: 家长职业
        housing_price: 小区房价段（AI补齐）
        consumption: 家庭消费力评估（AI补齐）
        thresholds: 支付力阈值配置 {"高": "条件", "中": "条件", "低": "条件"}

    Returns:
        支付力等级："高"/"中"/"低"
    """
    occ = (occupation or "").strip()
    price = (housing_price or "").strip()
    cons = (consumption or "").strip()
    combined = f"{occ} {price} {cons}"

    # 高支付力信号
    high_signals = ["高端", "别墅", "豪宅", "高薪", "高管", "老板", "做生意",
                    "体制内", "公务员", "医生", "律师", "金融", "10万", "8万", "高"]
    # 低支付力信号
    low_signals = ["回迁", "老小区", "城中村", "不稳定", "打工", "临时", "低收入", "低"]

    high_count = sum(1 for s in high_signals if s in combined)
    low_count = sum(1 for s in low_signals if s in combined)

    if high_count >= 1:
        return "高"
    if low_count >= 1:
        return "低"
    return "中"


def calculate_renewal_risk(learning_months: int, class_performance: str,
                           renewal_count: int, thresholds: Dict[str, str]) -> str:
    """根据配置阈值推算续费风险。

    Args:
        learning_months: 学习月份数
        class_performance: 课堂表现
        renewal_count: 续费次数
        thresholds: 续费风险阈值配置

    Returns:
        续费风险："高"/"中"/"低"
    """
    perf = (class_performance or "").strip()

    # 高风险：学习时间短 + 表现下滑 + 没续费
    if learning_months < 3 and renewal_count == 0:
        if any(w in perf for w in ["走神", "下滑", "不专注", "差", "一般"]):
            return "高"
        return "中"

    # 低风险：学习时间长 + 表现稳定 + 已续费
    if learning_months > 6 and renewal_count >= 1:
        if any(w in perf for w in ["专注", "稳定", "好", "积极"]):
            return "低"
        return "中"

    # 中间状态
    if learning_months >= 3 and learning_months <= 6:
        return "中"

    return "中"


def calculate_referral_potential(family_structure: str, plan_goal: str,
                                 ai_awareness: str, thresholds: Dict[str, str]) -> str:
    """根据配置阈值推算转介绍潜力。

    Args:
        family_structure: 家庭结构
        plan_goal: 家长规划目标
        ai_awareness: 对AI认知度
        thresholds: 转介绍潜力阈值配置

    Returns:
        转介绍潜力："高"/"中"/"低"
    """
    score = 0
    struct = (family_structure or "").strip()
    plan = (plan_goal or "").strip()
    ai = (ai_awareness or "").strip()

    # 三代同堂 +1
    if "三代" in struct or "老人" in struct or "爷爷" in struct or "奶奶" in struct:
        score += 1
    # 规划目标明确 +1
    if plan and plan != "不知道" and plan != "不清楚":
        score += 1
    # 对AI认知度高 +1
    if "高" in ai:
        score += 1

    if score >= 2:
        return "高"
    if score == 1:
        return "中"
    return "低"


def calculate_priority(payment: str, risk: str, referral: str,
                       priority_logic: str) -> int:
    """根据配置权重计算跟进优先级（1-5星）。

    Args:
        payment: 支付力等级
        risk: 续费风险
        referral: 转介绍潜力
        priority_logic: 跟进优先级排序逻辑（自然语言描述）

    Returns:
        优先级星级（1-5）
    """
    # 默认权重：支付力0.4 + 续费风险0.3 + 转介绍潜力0.3
    level_map = {"高": 3, "中": 2, "低": 1}

    p_score = level_map.get(payment, 2)
    r_score = level_map.get(risk, 2)
    ref_score = level_map.get(referral, 2)

    # 续费风险高=需要重点跟进（反向）
    risk_weight = 3 if risk == "高" else (2 if risk == "中" else 1)

    weighted = p_score * 0.4 + risk_weight * 0.3 + ref_score * 0.3

    if weighted >= 2.5:
        return 5
    if weighted >= 2.2:
        return 4
    if weighted >= 1.8:
        return 3
    if weighted >= 1.5:
        return 2
    return 1


def match_recommendation(student: Dict[str, Any],
                         rules: List[Dict[str, Any]]) -> Tuple[str, str]:
    """根据推荐规则匹配学生适合的产品方向。

    Args:
        student: 学生记录字典（含所有字段）
        rules: 推荐规则列表

    Returns:
        (推荐课程, 理由) 元组，无匹配返回 ("待配置", "")
    """
    if not rules:
        return ("待配置", "")

    payment = student.get("支付力", "")
    grade = student.get("年级", "")
    interest = student.get("兴趣偏好", "")
    school_tech = student.get("是否科技特色校", "")

    for rule in rules:
        conditions = rule.get("条件", {})
        matched = True

        if "支付力" in conditions and conditions["支付力"]:
            if payment != conditions["支付力"]:
                matched = False

        if matched and "年级" in conditions and conditions["年级"]:
            grade_cond = conditions["年级"]
            if grade_cond == "1-3年级":
                if not any(g in grade for g in ["1年级", "2年级", "3年级"]):
                    matched = False
            elif grade_cond == "4-6年级":
                if not any(g in grade for g in ["4年级", "5年级", "6年级"]):
                    matched = False

        if matched and "兴趣" in conditions and conditions["兴趣"]:
            if "科技" in conditions["兴趣"]:
                if not any(w in interest for w in ["科技", "编程", "电脑", "机器人", "乐高"]):
                    matched = False

        if matched and "学校科技特色" in conditions:
            if conditions["学校科技特色"] is True:
                if "科技特色" not in school_tech and "是" not in school_tech:
                    matched = False

        if matched:
            return (rule.get("推荐课程", "待配置"), rule.get("理由", ""))

    return ("待配置", "")


# ============================================================
# 脚本结果输出
# ============================================================

def print_script_result(success: bool, message: str, **stats) -> None:
    """输出脚本执行结果（JSON格式到控制台）。

    Args:
        success: 是否成功
        message: 结果消息
        **stats: 统计信息键值对
    """
    result = {
        "success": success,
        "message": message,
    }
    result.update(stats)
    print(json.dumps(result, ensure_ascii=False, indent=2))


# ============================================================
# 模块自测
# ============================================================

if __name__ == "__main__":
    # 测试冲突格式化
    print("=== 冲突格式化测试 ===")
    print(format_conflict_cell("专注", "容易走神"))

    # 测试完成率计算
    print("\n=== 完成率计算测试 ===")
    cache = create_collection_cache("示例校区", "老师", "王老师")
    cache["名单"] = ["张小明", "李小红", "王小刚"]
    cache["已采集记录"] = [{"姓名": "张小明"}]
    print(calculate_completion(cache))

    # 测试学习时长解析
    print("\n=== 学习时长解析测试 ===")
    print(parse_learning_months("8个月"))  # 8
    print(parse_learning_months("半年"))   # 6
    print(parse_learning_months("1年"))    # 12

    # 测试续费次数解析
    print("\n=== 续费次数解析测试 ===")
    print(parse_renewal_count("续费1次"))  # 1
    print(parse_renewal_count("没续过"))   # 0

    # 测试支付力推算
    print("\n=== 支付力推算测试 ===")
    thresholds = {"高": "高端小区", "中": "普通小区", "低": "回迁房"}
    print(calculate_payment_level("做生意", "8-10万/㎡", "高", thresholds))  # 高

    print("\n[utils.py] 自测完成")
