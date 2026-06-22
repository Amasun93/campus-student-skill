# -*- coding: utf-8 -*-
"""
校区学员情况管理 Skill - 决策标签推算脚本

功能：读取汇总表 + 校区配置.json → 按配置阈值计算5个决策标签
     → 写入汇总表（条件格式红黄绿）→ 无匹配规则标"待配置"

决策标签：
1. 支付力等级（高/中/低）- 从配置读支付力阈值
2. 续费风险（高/中/低）- 从配置读续费风险阈值
3. 转介绍潜力（高/中/低）- 从配置读转介绍潜力阈值
4. 跟进优先级（1-5星）- 从配置读权重逻辑
5. 推荐产品方向 - 从配置读推荐规则匹配

用法：
    python scripts/write_tags.py --input <汇总表.xlsx> --output <带标签汇总表.xlsx> --config <校区配置.json>

依赖：openpyxl>=3.1.2
"""

import argparse
import json
import os
import sys
from typing import Any, Dict, List, Optional, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from utils import (
    A_FIELDS, B_FIELDS, C_FIELDS, D_FIELDS,
    SUMMARY_COLUMNS, AI_FIELDS, TAG_FIELDS,
    get_excel_styles, apply_header_style, apply_tag_conditional_format,
    parse_learning_months, parse_renewal_count,
    calculate_payment_level, calculate_renewal_risk,
    calculate_referral_potential, calculate_priority,
    match_recommendation, print_script_result,
)


def read_summary_xlsx(file_path: str) -> List[Dict[str, Any]]:
    """读取汇总表.xlsx，解析为学生记录列表（扁平结构）。

    Args:
        file_path: 汇总表xlsx路径

    Returns:
        学生记录列表，每条含所有汇总表列字段
    """
    from openpyxl import load_workbook

    wb = load_workbook(file_path, data_only=True)
    if "学员汇总" not in wb.sheetnames:
        print(f"[错误] 汇总表中未找到'学员汇总'sheet")
        return []

    ws = wb["学员汇总"]
    headers = []
    for cell in ws[1]:
        headers.append(str(cell.value) if cell.value else "")

    students: List[Dict[str, Any]] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue

        student: Dict[str, Any] = {}
        for idx, header in enumerate(headers):
            student[header] = str(row[idx]) if idx < len(row) and row[idx] is not None else ""
        students.append(student)

    return students


def calculate_tags_for_student(student: Dict[str, Any],
                               config: Dict[str, Any]) -> Dict[str, Any]:
    """为单个学生计算5个决策标签。

    Args:
        student: 学生记录（扁平结构，含所有字段）
        config: 校区配置字典

    Returns:
        包含5个决策标签的字典
    """
    thresholds = config.get("决策标签阈值", {})
    rules = config.get("推荐规则", [])

    # 提取字段值
    occupation = student.get("家长职业", "")
    housing_price = student.get("小区房价段", "")
    consumption = student.get("家庭消费力", "")
    learning_duration = student.get("学习时长", "")
    class_performance = student.get("课堂表现", "")
    renewal_history = student.get("续费历史", "")
    family_structure = student.get("家庭结构", "")
    plan_goal = student.get("家长规划目标", "")
    ai_awareness = student.get("对AI认知度", "")

    # 1. 支付力等级
    payment_thresholds = thresholds.get("支付力", {})
    payment = calculate_payment_level(occupation, housing_price, consumption, payment_thresholds)

    # 2. 续费风险
    risk_thresholds = thresholds.get("续费风险", {})
    learning_months = parse_learning_months(learning_duration)
    renewal_count = parse_renewal_count(renewal_history)
    risk = calculate_renewal_risk(learning_months, class_performance, renewal_count, risk_thresholds)

    # 3. 转介绍潜力
    referral_thresholds = thresholds.get("转介绍潜力", {})
    referral = calculate_referral_potential(family_structure, plan_goal, ai_awareness, referral_thresholds)

    # 4. 跟进优先级
    priority_logic = thresholds.get("跟进优先级", "")
    priority = calculate_priority(payment, risk, referral, priority_logic)

    # 5. 推荐产品方向
    # 构建匹配用的学生字典
    match_student = dict(student)
    match_student["支付力"] = payment
    match_student["是否科技特色校"] = student.get("学校层次(科技特色)", "")
    recommended, reason = match_recommendation(match_student, rules)

    return {
        "支付力": payment,
        "续费风险": risk,
        "转介绍潜力": referral,
        "跟进优先级": str(priority),
        "推荐产品方向": recommended,
        "推荐理由": reason,
    }


def write_tags_xlsx(students: List[Dict[str, Any]],
                    tags_list: List[Dict[str, Any]],
                    template_path: str,
                    output_path: str) -> bool:
    """将带决策标签的数据写入汇总表.xlsx。

    Args:
        students: 学生记录列表
        tags_list: 对应的决策标签列表
        template_path: 原汇总表路径（用于复制其他sheet）
        output_path: 输出文件路径

    Returns:
        写入成功返回True
    """
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    # 基于原汇总表复制
    from openpyxl import Workbook
    wb = load_workbook(template_path)

    # 获取学员汇总sheet
    if "学员汇总" in wb.sheetnames:
        ws = wb["学员汇总"]
    else:
        ws = wb.active

    # 找到决策标签列的索引
    headers = []
    for cell in ws[1]:
        headers.append(str(cell.value) if cell.value else "")

    tag_col_map = {}
    for tag in TAG_FIELDS:
        if tag in headers:
            tag_col_map[tag] = headers.index(tag) + 1  # 1-based
        elif tag == "推荐产品方向" and "推荐产品方向" in headers:
            tag_col_map[tag] = headers.index("推荐产品方向") + 1

    # 写入决策标签
    for row_idx, (student, tags) in enumerate(zip(students, tags_list), 2):
        if "支付力" in tag_col_map:
            ws.cell(row=row_idx, column=tag_col_map["支付力"], value=tags["支付力"])
        if "续费风险" in tag_col_map:
            ws.cell(row=row_idx, column=tag_col_map["续费风险"], value=tags["续费风险"])
        if "转介绍潜力" in tag_col_map:
            ws.cell(row=row_idx, column=tag_col_map["转介绍潜力"], value=tags["转介绍潜力"])
        if "跟进优先级" in tag_col_map:
            stars = "⭐" * int(tags["跟进优先级"]) if tags["跟进优先级"].isdigit() else tags["跟进优先级"]
            ws.cell(row=row_idx, column=tag_col_map["跟进优先级"], value=stars)
        if "推荐产品方向" in tag_col_map:
            ws.cell(row=row_idx, column=tag_col_map["推荐产品方向"], value=tags["推荐产品方向"])

    # 应用条件格式
    for tag_col_name in ["支付力", "续费风险", "转介绍潜力"]:
        if tag_col_name in tag_col_map:
            apply_tag_conditional_format(ws, tag_col_map[tag_col_name], len(students))

    # 保存
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True) if os.path.dirname(output_path) else None
    wb.save(output_path)
    return True


def main():
    """主函数：解析参数并执行决策标签推算。"""
    parser = argparse.ArgumentParser(description="决策标签推算：读汇总表+配置→计算标签→写入汇总表")
    parser.add_argument("--input", required=True, help="输入汇总表xlsx路径")
    parser.add_argument("--output", required=True, help="输出带标签汇总表xlsx路径")
    parser.add_argument("--config", required=True, help="校区配置.json路径")
    args = parser.parse_args()

    # 检查文件
    if not os.path.exists(args.input):
        print_script_result(False, f"汇总表不存在: {args.input}")
        sys.exit(1)
    if not os.path.exists(args.config):
        print_script_result(False, f"配置文件不存在: {args.config}")
        sys.exit(1)

    # 加载配置
    with open(args.config, "r", encoding="utf-8") as f:
        config = json.load(f)

    # 读取汇总表
    students = read_summary_xlsx(args.input)
    if not students:
        print_script_result(False, "汇总表为空或读取失败")
        sys.exit(1)

    # 计算决策标签
    tags_list: List[Dict[str, Any]] = []
    for student in students:
        tags = calculate_tags_for_student(student, config)
        tags_list.append(tags)

    # 统计
    high_payment = sum(1 for t in tags_list if t["支付力"] == "高")
    high_risk = sum(1 for t in tags_list if t["续费风险"] == "高")
    high_referral = sum(1 for t in tags_list if t["转介绍潜力"] == "高")
    pending_config = sum(1 for t in tags_list if t["推荐产品方向"] == "待配置")

    # 写入汇总表
    try:
        write_tags_xlsx(students, tags_list, args.input, args.output)
        print_script_result(
            True,
            f"标签推算成功：{args.output}",
            总人数=len(students),
            高净值人数=high_payment,
            高续费风险人数=high_risk,
            高转介绍潜力人数=high_referral,
            待配置推荐人数=pending_config,
        )
    except Exception as e:
        print_script_result(False, f"标签推算异常: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
