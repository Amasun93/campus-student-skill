# -*- coding: utf-8 -*-
"""
校区学员情况管理 Skill - HTML报告生成脚本

功能：读取带标签汇总表.xlsx → 计算统计数据 → 序列化JSON → 注入HTML模板
     → 输出单文件校区分析报告.html（三页结构）

三页结构：
- 页1：全员表格+筛选（按校区/年级/支付力/风险筛选排序）
- 页2：校区分析报告（高净值占比/小区分布/学校层次分布/决策标签分布）
- 页3：产品推荐画像（推荐课程分布+百分比+画像筛选）

用法：
    python scripts/generate_html_report.py --input <带标签汇总表.xlsx> --output <报告.html> --config <校区配置.json>

依赖：openpyxl>=3.1.2
"""

import argparse
import json
import os
import sys
from collections import Counter
from typing import Any, Dict, List

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from utils import (
    SUMMARY_COLUMNS, get_timestamp, print_script_result,
)


def read_tagged_xlsx(file_path: str) -> List[Dict[str, Any]]:
    """读取带标签汇总表.xlsx，解析为学生记录列表。

    Args:
        file_path: 汇总表xlsx路径

    Returns:
        学生记录列表
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
            val = str(row[idx]) if idx < len(row) and row[idx] is not None else ""
            student[header] = val
        students.append(student)

    return students


def parse_priority_to_int(priority_str: str) -> int:
    """将跟进优先级字符串解析为整数（⭐符号计数）。

    Args:
        priority_str: 优先级字符串，如"⭐⭐⭐⭐"或"4"

    Returns:
        星级数（1-5）
    """
    if not priority_str:
        return 0
    # 数字符串中的⭐数量
    star_count = priority_str.count("⭐")
    if star_count > 0:
        return star_count
    # 尝试解析数字
    try:
        return int(priority_str)
    except ValueError:
        return 0


def build_report_data(students: List[Dict[str, Any]], campus: str) -> Dict[str, Any]:
    """构建HTML报告数据结构。

    Args:
        students: 学生记录列表
        campus: 校区名

    Returns:
        报告数据字典（含学员列表和统计信息）
    """
    total = len(students)

    # 构建学员列表（精简字段用于前端展示）
    student_list: List[Dict[str, Any]] = []
    for s in students:
        student_list.append({
            "姓名": s.get("姓名", ""),
            "年级": s.get("年级", ""),
            "校区": s.get("校区", campus),
            "居住小区": s.get("居住小区", ""),
            "学校层次": s.get("学校层次(科技特色)", ""),
            "支付力": s.get("支付力", ""),
            "续费风险": s.get("续费风险", ""),
            "转介绍潜力": s.get("转介绍潜力", ""),
            "跟进优先级": parse_priority_to_int(s.get("跟进优先级", "")),
            "推荐方向": s.get("推荐产品方向", ""),
            "推荐理由": s.get("推荐理由", "") if "推荐理由" in s else "",
        })

    # 统计计算
    payment_counter = Counter(s.get("支付力", "") for s in students)
    risk_counter = Counter(s.get("续费风险", "") for s in students)
    referral_counter = Counter(s.get("转介绍潜力", "") for s in students)
    recommend_counter = Counter(s.get("推荐产品方向", "") for s in students)

    # 高净值占比
    high_payment_count = payment_counter.get("高", 0)
    high_payment_pct = f"{(high_payment_count / total * 100):.0f}%" if total > 0 else "0%"

    # 续费高风险占比
    high_risk_count = risk_counter.get("高", 0)
    high_risk_pct = f"{(high_risk_count / total * 100):.0f}%" if total > 0 else "0%"

    # 高转介绍潜力占比
    high_referral_count = referral_counter.get("高", 0)
    high_referral_pct = f"{(high_referral_count / total * 100):.0f}%" if total > 0 else "0%"

    # 家长来源小区分布
    community_counter = Counter()
    for s in students:
        community = s.get("居住小区", "").strip()
        if community and community != "不知道" and community != "不清楚":
            community_counter[community] += 1
    community_dist = [{"小区": k, "人数": v} for k, v in community_counter.most_common(10)]

    # 学校层次分布
    school_counter = Counter()
    for s in students:
        school_level = s.get("学校层次(科技特色)", "").strip()
        if not school_level:
            school_level = "未补齐"
        if "科技特色" in school_level:
            school_level = "科技特色校"
        elif "重点" in school_level:
            school_level = "重点校"
        elif "普通" in school_level:
            school_level = "普通校"
        school_counter[school_level] += 1
    school_dist = [{"层次": k, "人数": v} for k, v in school_counter.most_common()]

    # 推荐分布
    recommend_dist = []
    for course, count in recommend_counter.most_common():
        pct = f"{(count / total * 100):.0f}%" if total > 0 else "0%"
        recommend_dist.append({"课程": course if course else "待配置", "人数": count, "占比": pct})

    # 决策标签分布
    tag_dist = []
    for level in ["高", "中", "低"]:
        tag_dist.append({"标签": f"支付力-{level}", "人数": payment_counter.get(level, 0)})
        tag_dist.append({"标签": f"续费风险-{level}", "人数": risk_counter.get(level, 0)})
        tag_dist.append({"标签": f"转介绍潜力-{level}", "人数": referral_counter.get(level, 0)})

    # 组装报告数据
    report_data = {
        "生成时间": get_timestamp(),
        "校区": campus,
        "学员列表": student_list,
        "统计": {
            "总人数": total,
            "高净值占比": high_payment_pct,
            "续费高风险占比": high_risk_pct,
            "高转介绍潜力占比": high_referral_pct,
            "家长来源小区分布": community_dist,
            "学校层次分布": school_dist,
            "推荐分布": recommend_dist,
            "决策标签分布": tag_dist,
        }
    }

    return report_data


def inject_data_to_html(template_path: str, report_data: Dict[str, Any],
                        output_path: str) -> bool:
    """将报告数据注入HTML模板，生成单文件报告。

    Args:
        template_path: HTML模板文件路径
        report_data: 报告数据字典
        output_path: 输出HTML文件路径

    Returns:
        生成成功返回True
    """
    # 读取模板
    with open(template_path, "r", encoding="utf-8") as f:
        html_content = f.read()

    # 将数据序列化为JSON
    data_json = json.dumps(report_data, ensure_ascii=False, indent=2)

    # 替换模板中的REPORT_DATA占位
    # 模板中有: const REPORT_DATA = { ... };
    # 我们用实际数据替换，使用贪婪匹配处理嵌套花括号
    import re
    pattern = r'const REPORT_DATA\s*=\s*\{.*?\};'
    replacement = f'const REPORT_DATA = {data_json};'
    html_content = re.sub(pattern, replacement, html_content, count=1, flags=re.DOTALL)

    # 写入输出文件
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True) if os.path.dirname(output_path) else None
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    return True


def main():
    """主函数：解析参数并生成HTML报告。"""
    parser = argparse.ArgumentParser(description="生成HTML单文件报告（三页结构）")
    parser.add_argument("--input", required=True, help="带标签汇总表xlsx路径")
    parser.add_argument("--output", required=True, help="输出HTML报告路径")
    parser.add_argument("--config", help="校区配置.json路径（读取报告偏好，可选）")
    args = parser.parse_args()

    # 检查文件
    if not os.path.exists(args.input):
        print_script_result(False, f"汇总表不存在: {args.input}")
        sys.exit(1)

    # 读取配置（可选）
    campus = ""
    if args.config and os.path.exists(args.config):
        with open(args.config, "r", encoding="utf-8") as f:
            config = json.load(f)
        campus = config.get("校区", "")

    # 读取汇总表
    students = read_tagged_xlsx(args.input)
    if not students:
        print_script_result(False, "汇总表为空或读取失败")
        sys.exit(1)

    # 如果未从配置获取校区，从数据中取
    if not campus and students:
        campus = students[0].get("校区", "未知校区")

    # 构建报告数据
    report_data = build_report_data(students, campus)

    # 查找HTML模板
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    template_path = os.path.join(project_root, "templates", "html_report_template.html")

    if not os.path.exists(template_path):
        print_script_result(False, f"HTML模板不存在: {template_path}")
        sys.exit(1)

    # 注入数据生成报告
    try:
        inject_data_to_html(template_path, report_data, args.output)
        print_script_result(
            True,
            f"HTML报告生成成功：{args.output}",
            校区=campus,
            总人数=report_data["统计"]["总人数"],
            高净值占比=report_data["统计"]["高净值占比"],
            续费高风险占比=report_data["统计"]["续费高风险占比"],
        )
    except Exception as e:
        print_script_result(False, f"HTML报告生成异常: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
