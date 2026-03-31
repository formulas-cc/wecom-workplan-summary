#!/usr/bin/env python3
"""
团队工作计划汇总脚本
用法: python3 summary.py [周度/月度] [周数] [年月]
"""

import json
import subprocess
import re
from collections import defaultdict
from datetime import datetime, timedelta

# 智能表格配置
DOCID = "dcrZNwuyF7QzK5GW4oQ9B3Y3i6Vdc_RzIxAc1zWMMVr9K4EWCEEza2Ea1XGGuQumBW2IKp9XR6au-lDtMi--j27A"
SHEET_ID = "q979lj"

# 管理层角色
ROLE_MAP = {
    "张鹏乐": ("月度目标制定", 0),
    "王紫龙": ("项目总监", 1),
    "付岩": ("技术总监", 2),
}


def call_smartsheet_get_records():
    """调用 wecom_mcp 获取智能表格记录"""
    cmd = [
        "wecom_mcp", "call", "doc", "smartsheet_get_records",
        json.dumps({"docid": DOCID, "sheet_id": SHEET_ID})
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise Exception(f"MCP调用失败: {result.stderr}")

    # 解析 JSON 输出
    try:
        data = json.loads(result.stdout)
        if data.get("errcode") != 0:
            raise Exception(f"API错误: {data.get('errmsg')}")
        return data.get("data", {}).get("records", [])
    except json.JSONDecodeError:
        raise Exception(f"返回数据解析失败: {result.stdout}")


def get_week_range(week_num, year=2026):
    """获取指定周数的日期范围（周一到周五）"""
    # 找到该年的第1周周一
    jan_1 = datetime(year, 1, 1)
    # 第1周的周一
    first_monday = jan_1 + timedelta(days=(7 - jan_1.weekday()) % 7)
    # 第N周的周一
    target_monday = first_monday + timedelta(weeks=week_num - 1)

    dates = []
    for i in range(5):  # 周一到周五
        d = target_monday + timedelta(days=i)
        dates.append(d)

    return dates


def get_month_range(year, month):
    """获取指定月份的日期范围"""
    from calendar import monthrange
    start = datetime(year, month, 1)
    _, last_day = monthrange(year, month)
    end = datetime(year, month, last_day)

    dates = []
    d = start
    while d <= end:
        if d.weekday() < 5:  # 周一到周五
            dates.append(d)
        d += timedelta(days=1)

    return dates


def parse_date(date_str):
    """解析日期字符串"""
    # 尝试多种格式
    formats = [
        "%Y年%m月%d日",
        "%Y-%m-%d",
        "%Y/%m/%d",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None


def filter_by_range(records, date_range, date_field="日期"):
    """按日期范围筛选记录"""
    filtered = defaultdict(lambda: {'岗位': '', '计划': []})

    for record in records:
        # 获取字段值
        fields = record.get("fields", {})

        # 解析日期
        date_val = fields.get(date_field, "")
        if isinstance(date_val, list):
            date_val = date_val[0].get("text", "") if date_val else ""
        elif isinstance(date_val, dict):
            date_val = date_val.get("text", "")

        parsed_date = parse_date(str(date_val))
        if not parsed_date:
            continue

        # 检查是否在日期范围内
        if parsed_date not in date_range:
            continue

        # 解析姓名（从 _id 或其他字段）
        name = record.get("record_id", "未知")
        # 尝试从字段获取姓名
        name_field = fields.get("姓名") or fields.get("成员") or fields.get("提交人")
        if name_field:
            if isinstance(name_field, list):
                name = name_field[0].get("text", name) if name_field else name
            elif isinstance(name_field, dict):
                name = name_field.get("text", name)

        # 解析计划内容
        plan_field = fields.get("今日计划", "")
        if isinstance(plan_field, list):
            plan = plan_field[0].get("text", "") if plan_field else ""
        elif isinstance(plan_field, dict):
            plan = plan_field.get("text", "")
        else:
            plan = str(plan_field)

        if plan:
            filtered[name]['计划'].append(plan)

        # 解析岗位
        position_field = fields.get("岗位") or fields.get("职位")
        if position_field and not filtered[name]['岗位']:
            if isinstance(position_field, list):
                filtered[name]['岗位'] = position_field[0].get("text", "") if position_field else ""
            elif isinstance(position_field, dict):
                filtered[name]['岗位'] = position_field.get("text", "")
            else:
                filtered[name]['岗位'] = str(position_field)

    return filtered


def format_output(people, date_range, mode="周度"):
    """格式化输出"""
    # 排序
    def sort_key(item):
        name = item[0]
        if name in ROLE_MAP:
            return (ROLE_MAP[name][1], name)
        return (99, name)

    filled = []
    empty = []

    for name, info in sorted(people.items(), key=sort_key):
        plans = info['计划']
        if plans:
            filled.append((name, info, plans))
        else:
            empty.append((name, info))

    # 构建输出
    lines = []

    if mode == "周度":
        week_num = 13  # TODO: 根据日期计算
        start_str = date_range[0].strftime("%m/%d")
        end_str = date_range[-1].strftime("%m/%d")
        lines.append(f"## 📅 第{week_num}周工作周报（{start_str}-{end_str}）\n")
    else:
        month = date_range[0].month
        year = date_range[0].year
        lines.append(f"## 📅 {year}年{month}月 月度工作汇总\n")

    lines.append("---\n")

    for name, info, plans in filled:
        role = info['岗位']
        if name in ROLE_MAP:
            role = ROLE_MAP[name][0]

        lines.append(f"\n### {name} · {role}")
        lines.append(f"**本{mode}**：")
        for p in plans:
            p_clean = p.replace('\n', ' ').strip()
            if len(p_clean) > 55:
                p_clean = p_clean[:52] + "..."
            lines.append(f"  • {p_clean}")

        lines.append(f"> 🔹 目标对齐：任务推进中")
        lines.append(f"> 💡 建议：持续推进，注意进度同步")

    if empty:
        mode_label = "周" if mode == "周度" else "月"
        lines.append(f"\n### ⚠️ 本{mode_label}无记录")
        for name, info in empty:
            lines.append(f"- {name}（{info['岗位']}）")

    last_date = date_range[-1].strftime("%Y年%m月%d日")
    lines.append(f"\n---\n*注：数据仅到{last_date}*")

    return "\n".join(lines)


def main():
    import sys

    mode = "周度"
    week_num = 13
    year = 2026
    month = 3

    # 解析参数
    if len(sys.argv) > 1:
        if sys.argv[1] in ["周度", "月度"]:
            mode = sys.argv[1]

    if len(sys.argv) > 2:
        if mode == "周度":
            week_num = int(sys.argv[2])
        else:
            month = int(sys.argv[2])

    if len(sys.argv) > 3:
        year = int(sys.argv[3])

    # 计算日期范围
    if mode == "周度":
        date_range = get_week_range(week_num, year)
    else:
        date_range = get_month_range(year, month)

    if not date_range:
        print(f"错误：无效的日期范围")
        return

    try:
        # 读取数据
        print("正在读取智能表格数据...")
        records = call_smartsheet_get_records()
        print(f"共获取 {len(records)} 条记录")

        # 筛选数据
        people = filter_by_range(records, date_range)
        print(f"筛选后 {len(people)} 人有记录")

        # 生成报告
        output = format_output(people, date_range, mode)
        print("\n" + output)

    except Exception as e:
        print(f"错误：{e}")
        print("请确认：")
        print("1. MCP 工具已正确配置")
        print("2. docid 和 sheet_id 正确")


if __name__ == "__main__":
    main()
