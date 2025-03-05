#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
课表转ICS工具
将文本格式的课表信息转换为ICS格式的日历文件
支持自动提取课程名称、教师信息、地点、时间等关键信息
"""

import re
import json
from datetime import datetime, timedelta
import pytz
from ics import Calendar, Event
from ics.alarm import DisplayAlarm

# 节次时间映射表 - 使用阿拉伯数字作为键
# 更新为用户提供的时间段
class_time_map = {
    1: {"start": "08:00", "end": "08:45"},
    2: {"start": "08:55", "end": "09:30"},
    3: {"start": "09:40", "end": "10:25"},
    4: {"start": "10:35", "end": "11:10"},
    5: {"start": "12:40", "end": "13:25"},
    6: {"start": "13:35", "end": "14:10"},
    7: {"start": "14:20", "end": "15:05"},
    8: {"start": "15:15", "end": "15:50"},
    9: {"start": "17:00", "end": "17:45"},
    10: {"start": "17:55", "end": "18:40"},
    11: {"start": "18:50", "end": "19:35"},
    12: {"start": "19:45", "end": "20:10"},
}

# 合并节次的时间映射表 - 用于直接获取合并节次的开始和结束时间
merged_time_map = {
    "1-2": {"start": "08:00", "end": "09:30"},
    "3-4": {"start": "09:40", "end": "11:10"},
    "5-6": {"start": "12:40", "end": "14:10"},
    "7-8": {"start": "14:20", "end": "15:50"},
    "9-12": {"start": "17:00", "end": "20:10"},
}

# 显示用的节次名称
slot_names = {
    1: "第一节", 
    2: "第二节", 
    3: "第三节", 
    4: "第四节", 
    5: "第五节",
    6: "第六节", 
    7: "第七节", 
    8: "第八节", 
    9: "第九节", 
    10: "第十节", 
    11: "第十一节", 
    12: "第十二节"
}

# 星期映射 - 用于计算日期偏移量
weekday_map = {
    "星期一": 0,
    "星期二": 1,
    "星期三": 2,
    "星期四": 3,
    "星期五": 4,
    "星期六": 5,
    "星期日": 6,
}

# 中文星期显示 - 用于格式化输出
weekday_display = {
    0: "星期一",
    1: "星期二",
    2: "星期三",
    3: "星期四",
    4: "星期五",
    5: "星期六",
    6: "星期日",
}

# 中国时区
local_timezone = pytz.timezone('Asia/Shanghai')

# 中文数字到阿拉伯数字的映射 - 用于解析节次信息
cn_num_map = {
    "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
    "六": 6, "七": 7, "八": 8, "九": 9, "十": 10,
    "十一": 11, "十二": 12
}

def extract_course_name(text_block):
    """
    从课程文本块中提取课程名称
    
    Args:
        text_block: 包含课程信息的文本块
        
    Returns:
        str: 提取出的课程名称，如无法提取则返回"未知课程"
    """
    lines = text_block.strip().split('\n')
    if not lines:
        return "未知课程"
    
    # 课程名称通常是第一行，如"AI+商品信息采编"或"Python Django框架开发实践"
    # 排除第一行如果它是"课程信息"等表头
    first_line = lines[0].strip()
    if first_line in ["课程信息", "教学班", "时间地点人员", "人数", "教学材料"]:
        if len(lines) > 1:
            return lines[1].strip()
        return "未知课程"
    
    return first_line

def parse_weeks(week_str):
    """
    解析周次信息，如'1~8周'转换为[(1, 8)]
    
    Args:
        week_str: 包含周次信息的字符串，如"1~8周"、"9~14周"或"10周"
        
    Returns:
        list: 周次范围列表，每个元素为(start_week, end_week)的元组
    """
    result = []
    # 匹配形如"1~8周"或"9~14周"的字符串
    pattern = r'(\d+)~(\d+)周'
    matches = re.finditer(pattern, week_str)
    
    for match in matches:
        start_week = int(match.group(1))
        end_week = int(match.group(2))
        result.append((start_week, end_week))
    
    # 处理单周情况，如"10周"
    single_week_pattern = r'(\d+)周(?!\s*~)'
    single_matches = re.finditer(single_week_pattern, week_str)
    for match in single_matches:
        week_num = int(match.group(1))
        result.append((week_num, week_num))
    
    return result

def parse_time_slots(time_str):
    """
    解析节次信息，如'第一节~第二节'转换为[(1, 2)]
    
    Args:
        time_str: 包含节次信息的字符串，如"第一节~第二节"或"第九节~十二节"
        
    Returns:
        list: 节次范围列表，每个元素为(start_slot, end_slot)的元组
    """
    result = []
    # 匹配形如"第一节~第二节"或"第九节~十二节"的字符串
    pattern = r'第([一二三四五六七八九十]+)节~(?:第)?([一二三四五六七八九十]+)节'
    matches = re.finditer(pattern, time_str)
    
    for match in matches:
        start_slot = match.group(1)
        end_slot = match.group(2)
        
        # 转换中文数字到阿拉伯数字
        start_num = _convert_cn_num(start_slot)
        end_num = _convert_cn_num(end_slot)
        
        result.append((start_num, end_num))
    
    return result

def _convert_cn_num(cn_num):
    """
    将中文数字转换为阿拉伯数字
    
    Args:
        cn_num: 中文数字字符串，如"一"、"十一"
        
    Returns:
        int: 对应的阿拉伯数字
    """
    if cn_num in cn_num_map:
        return cn_num_map[cn_num]
    
    # 处理"十一"、"十二"这样的特殊情况
    if "十" in cn_num:
        if len(cn_num) > 1:  # 十几
            second_char = cn_num[1]
            if second_char in cn_num_map:
                return 10 + cn_num_map[second_char]
            else:
                return 10
        else:
            return 10
    
    return 1  # 默认值

def parse_course_info(text):
    """
    解析课程信息，提取课程名称、周次、星期、节次、教室等关键信息
    
    Args:
        text: 包含所有课程信息的文本
        
    Returns:
        list: 课程信息列表，每个元素为一个课程字典
    """
    courses = []
    
    # 分割课程块 - 每个课程以课程名称开始，到下一个课程名称前结束
    course_blocks = []
    current_block = []
    course_started = False
    
    lines = text.strip().split('\n')
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
        
        # 识别课程开始 - 检查下一行是否包含课程代码模式
        if i < len(lines) - 1 and re.search(r'\d{3}[A-Z]\d{4}', lines[i+1]):
            if current_block:
                course_blocks.append('\n'.join(current_block))
            current_block = [line]
            course_started = True
        elif course_started:
            current_block.append(line)
    
    # 添加最后一个课程块
    if current_block:
        course_blocks.append('\n'.join(current_block))
    
    # 处理每个课程块
    for block in course_blocks:
        # 提取课程名称
        course_name = extract_course_name(block)
        
        # 创建课程对象
        course = {
            "name": course_name,
            "time_locations": []
        }
        
        # 解析时间地点信息 - 支持两种时间格式
        time_loc_pattern = r'((?:\d+~\d+|\d+)周)\s+(星期[一二三四五六日])\s+((?:第.*?节~.*?节)|(?:\d{2}:\d{2}~\d{2}:\d{2}))\s+(金海\s*\w+)\s+([^;\n]+)'
        for match in re.finditer(time_loc_pattern, block):
            weeks_str = match.group(1)
            weekday = match.group(2)
            time_str = match.group(3)
            location = match.group(4).replace(' ', '')
            teacher = match.group(5).strip().rstrip(';')
            
            # 解析周次
            weeks = parse_weeks(weeks_str)
            
            # 判断时间格式并解析
            if '第' in time_str:
                # 传统节次格式
                time_slots = parse_time_slots(time_str)
                if time_slots:
                    first_slot = time_slots[0][0]
                    last_slot = time_slots[-1][1]
                    slot_key = f"{first_slot}-{last_slot}"
                    if slot_key in merged_time_map:
                        start_time = merged_time_map[slot_key]["start"]
                        end_time = merged_time_map[slot_key]["end"]
                    else:
                        start_time = class_time_map[first_slot]["start"]
                        end_time = class_time_map[last_slot]["end"]
            else:
                # 直接时间格式 (如 17:40~19:55)
                start_time, end_time = time_str.split('~')
                time_slots = None
            
            time_location = {
                "weeks": weeks,
                "weekday": weekday,
                "time_slots": time_slots,
                "start_time": start_time,
                "end_time": end_time,
                "time_str": f"{start_time}-{end_time}",
                "location": location,
                "teacher": teacher
            }
            
            course["time_locations"].append(time_location)
        
        # 只添加有时间地点信息的课程
        if course["time_locations"]:
            courses.append(course)
    
    return courses

def calculate_total_weeks(time_locations):
    """
    计算课程总周数
    
    Args:
        time_locations: 课程时间地点信息列表
        
    Returns:
        int: 总周数
    """
    if not time_locations:
        return 0
    
    # 创建一个集合来存储所有周次
    all_weeks = set()
    
    for tl in time_locations:
        for start_week, end_week in tl["weeks"]:
            # 将该周次范围内的所有周次添加到集合中
            all_weeks.update(range(start_week, end_week + 1))
    
    return len(all_weeks)

def get_time_range(time_slots):
    """
    获取课程时间范围，合并连续的时间段
    
    Args:
        time_slots: 节次信息列表，如[(1, 2)]
        
    Returns:
        dict: 包含开始和结束时间的字典
    """
    if not time_slots:
        return None
    
    first_slot = time_slots[0][0]
    last_slot = time_slots[-1][1]
    
    # 尝试使用合并节次时间
    slot_key = f"{first_slot}-{last_slot}"
    if slot_key in merged_time_map:
        return merged_time_map[slot_key]
    
    # 如果没有合并节次的映射，使用单节次时间
    return {
        "start": class_time_map[first_slot]["start"], 
        "end": class_time_map[last_slot]["end"]
    }

def parse_classroom_location(location):
    """
    解析教室位置信息，提取楼号、楼层和教室号
    
    Args:
        location: 原始地点字符串，如"金海9408"
        
    Returns:
        str: 格式化后的地址
    """
    # 移除"金海"和空格
    location = location.replace("金海", "").strip()
    
    # 如果长度不是4或5，返回原始地址
    if len(location) not in [4, 5]:
        return f"上海杉达学院金海校区 {location}"
    
    try:
        if len(location) == 4:
            building = location[0]     # 楼号
            floor = location[1]        # 楼层
            room = location[2:]        # 教室号
        else:  # 长度为5，处理如12510这样的情况
            building = location[0:2]   # 楼号（如12）
            floor = location[2]        # 楼层
            room = location[3:]        # 教室号
        
        # 楼栋名称映射
        building_names = {
            "9": "胜祥商学院楼",
            "8": "科学实验楼",
            "12": "综合实验实训大楼"
        }
        
        # 处理1-6号楼
        if building in ["1", "2", "3", "4", "5", "6"]:
            return f"上海杉达学院金海校区教学楼{building}号楼 {floor}楼{room}教室"
        # 处理其他特殊命名的楼
        elif building in building_names:
            return f"上海杉达学院金海校区{building_names[building]} {floor}楼{room}教室"
        else:
            return f"上海杉达学院金海校区{building}号楼 {floor}楼{room}教室"
    except:
        return f"上海杉达学院金海校区 {location}"

def generate_events(courses, semester_start_date):
    """
    生成日历事件
    
    Args:
        courses: 课程信息列表
        semester_start_date: 学期开始日期
        
    Returns:
        Calendar: 包含所有课程事件的日历对象
    """
    calendar = Calendar()
    
    # 确保学期开始日期是一个datetime对象
    if isinstance(semester_start_date, str):
        semester_start_date = datetime.strptime(semester_start_date, '%Y-%m-%d')
    
    # 确定第一周的周一日期
    days_offset = semester_start_date.weekday()  # 0是周一，6是周日
    first_monday = semester_start_date - timedelta(days=days_offset)
    
    # 用于记录已添加事件的集合，避免重复
    # 格式: (课程名, 教师, 周次, 星期, 时间段)
    added_events = set()
    
    # 地址映射字典 - 用于扩展简短地点为完整地址
    location_map = {
        "金海": "上海杉达学院-金海校区"
    }
    
    # 创建一个字典来存储每天的课程时间段
    # 格式: {(周数, 星期): {时间段}}
    daily_slots = {}
    
    # 第一遍遍历：收集每天的课程时间段
    for course in courses:
        for time_location in course["time_locations"]:
            weekday_offset = weekday_map[time_location["weekday"]]
            
            # 获取课程的时间段
            start_hour, start_minute = map(int, time_location["start_time"].split(':'))
            
            # 判断属于哪个时间段（1-2节，3-4节，5-6节，7-8节，9-12节）
            time_slot = None
            if start_hour == 8:  # 1-2节
                time_slot = "1-2"
            elif start_hour == 9:  # 3-4节
                time_slot = "3-4"
            elif start_hour == 12:  # 5-6节
                time_slot = "5-6"
            elif start_hour == 14:  # 7-8节
                time_slot = "7-8"
            elif start_hour == 17:  # 9-12节
                time_slot = "9-12"
            
            # 处理每个周次范围
            for week_range in time_location["weeks"]:
                start_week, end_week = week_range
                for week_num in range(start_week, end_week + 1):
                    day_key = (week_num, weekday_offset)
                    if day_key not in daily_slots:
                        daily_slots[day_key] = set()
                    daily_slots[day_key].add(time_slot)
    
    # 第二遍遍历：生成事件并根据规则添加提醒
    for course in courses:
        for time_location in course["time_locations"]:
            weekday_offset = weekday_map[time_location["weekday"]]
            
            # 处理每个周次范围
            for week_range in time_location["weeks"]:
                start_week, end_week = week_range
                
                # 处理每一周
                for week_num in range(start_week, end_week + 1):
                    # 创建事件唯一标识
                    event_key = (
                        course['name'],
                        time_location['teacher'],
                        week_num,
                        time_location["weekday"],
                        time_location["start_time"],
                        time_location["end_time"]
                    )
                    
                    # 如果事件已添加过，则跳过
                    if event_key in added_events:
                        continue
                    
                    # 记录事件已添加
                    added_events.add(event_key)
                    
                    # 计算具体日期：第一周周一 + (当前周-1)*7天 + 当前星期偏移
                    course_date = first_monday + timedelta(days=(week_num-1)*7 + weekday_offset)
                    
                    # 创建开始和结束时间的datetime对象
                    start_hour, start_minute = map(int, time_location["start_time"].split(':'))
                    end_hour, end_minute = map(int, time_location["end_time"].split(':'))
                    
                    # 创建带时区的datetime对象
                    start_datetime = local_timezone.localize(
                        course_date.replace(hour=start_hour, minute=start_minute, second=0)
                    )
                    end_datetime = local_timezone.localize(
                        course_date.replace(hour=end_hour, minute=end_minute, second=0)
                    )
                    
                    # 创建事件
                    event = Event()
                    event.name = f"{course['name']}（{time_location['teacher']}）"
                    event.begin = start_datetime
                    event.end = end_datetime
                    
                    # 判断是否需要添加提醒
                    day_key = (week_num, weekday_offset)
                    day_slots = daily_slots.get(day_key, set())
                    
                    # 获取当前课程的时间段
                    current_slot = None
                    if start_hour == 8:
                        current_slot = "1-2"
                    elif start_hour == 9:
                        current_slot = "3-4"
                    elif start_hour == 12:
                        current_slot = "5-6"
                    elif start_hour == 14:
                        current_slot = "7-8"
                    elif start_hour == 17:
                        current_slot = "9-12"
                    
                    # 判断是否需要添加提醒
                    need_alarm = True
                    if current_slot == "3-4" and "1-2" in day_slots:
                        need_alarm = False
                    elif current_slot == "7-8" and "5-6" in day_slots:
                        need_alarm = False
                    
                    # 如果需要提醒，添加VALARM组件
                    if need_alarm:
                        alarm = DisplayAlarm(trigger=timedelta(minutes=-20))
                        event.alarms.append(alarm)
                    
                    # 处理地点信息 - 解析教室号码
                    original_location = time_location["location"]
                    full_location = parse_classroom_location(original_location)
                    event.location = full_location
                    
                    # 添加描述 - 只包含关键信息
                    description = [
                        f"周次: 第{week_num}周",
                        f"时间: {time_location['weekday']} {time_location['time_str']}"
                    ]
                    event.description = "\n".join(filter(None, description))
                    
                    calendar.events.add(event)
    
    return calendar

def format_course_summary(course):
    """
    格式化课程摘要信息 - 只包含关键字段
    
    Args:
        course: 课程信息字典
        
    Returns:
        dict: 格式化后的课程摘要
    """
    teachers = set(tl['teacher'] for tl in course['time_locations'])
    
    # 计算总周数
    all_weeks = set()
    for tl in course['time_locations']:
        for start_week, end_week in tl['weeks']:
            all_weeks.update(range(start_week, end_week + 1))
    total_weeks = len(all_weeks)
    
    # 课程时间地点摘要
    time_loc_summary = []
    for tl in course['time_locations']:
        weeks_str = ", ".join([f"{start}-{end}周" for start, end in tl['weeks']])
        time_loc_summary.append(f"{weeks_str} {tl['weekday']} {tl['time_str']} {tl['location']}")
    
    summary = {
        "name": course['name'],
        "teachers": ", ".join(teachers),
        "total_weeks": total_weeks,
        "time_locations": time_loc_summary
    }
    
    return summary

def main():
    """主函数 - 程序入口点"""
    print("=" * 50)
    print("课表转ICS工具")
    print("=" * 50)
    
    try:
        # 读取课表文本
        with open('kebiao.txt', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 解析课程信息
        courses = parse_course_info(content)
        
        # 提示用户输入学期开始日期
        print("请输入学期开始日期（格式：YYYY-MM-DD）：")
        semester_start = input().strip()
        
        try:
            # 解析日期
            semester_start_date = datetime.strptime(semester_start, '%Y-%m-%d')
            
            # 生成日历事件
            calendar = generate_events(courses, semester_start_date)
            
            # 输出ICS文件
            with open('schedule.ics', 'w', encoding='utf-8') as f:
                f.write(str(calendar))
            
            print(f"\n成功生成ICS文件：schedule.ics，包含 {len(calendar.events)} 个课程事件")
            
            # 打印已解析的课程信息供用户确认
            print("\n已解析的课程信息：")
            for i, course in enumerate(courses, 1):
                summary = format_course_summary(course)
                print(f"{i}. {summary['name']} - 教师: {summary['teachers']} - 共{summary['total_weeks']}周")
                for j, tl_str in enumerate(summary['time_locations'], 1):
                    print(f"   上课{j}: {tl_str}")
                print()
            
            # 将课程数据结构保存到JSON文件 - 只包含关键信息
            course_data = [format_course_summary(course) for course in courses]
            with open('course_data.json', 'w', encoding='utf-8') as f:
                json.dump(course_data, f, ensure_ascii=False, indent=2)
            print(f"课程数据已保存到 course_data.json")
            
        except ValueError:
            print("日期格式错误，请使用YYYY-MM-DD格式")
    except FileNotFoundError:
        print("错误：找不到课表文件kebiao.txt")
        print("请确保文件存在并放置在正确的目录中")
    except Exception as e:
        print(f"发生错误：{e}")
    
    print("\n" + "=" * 50)

if __name__ == "__main__":
    main() 