import pandas as pd
import numpy as np
import os
from datetime import datetime, timedelta
import re

def create_output_folder():
    """创建输出文件夹"""
    output_folder = "个人日程表"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"创建输出文件夹: {output_folder}")
    return output_folder

def read_schedule_data():
    """读取总表数据"""
    if os.path.exists('总表.xlsx'):
        df = pd.read_excel('总表.xlsx', index_col=0)
        return df
    else:
        raise FileNotFoundError("未找到总表.xlsx文件")

def extract_all_personnel_names(zdf):
    """从总表中提取所有人员名字"""
    personnel_names = set()
    
    for date in zdf.index:
        for village in zdf.columns:
            cell_value = zdf.loc[date, village]
            if pd.notna(cell_value) and cell_value != '':
                names_in_cell = parse_cell_content(cell_value)
                for name in names_in_cell:
                    personnel_names.add(name)
    
    return personnel_names

def parse_cell_content(cell_value):
    """解析单元格内容，提取人员名字"""
    if pd.isna(cell_value) or cell_value == '':
        return []
    
    # 按逗号分割
    names = str(cell_value).split(',')
    # 去除空白字符
    names = [name.strip() for name in names if name.strip()]
    
    return names

def find_person_schedule(person_name, zdf):
    """查找某个人的日程安排"""
    schedule = []
    
    for date in zdf.index:
        for village in zdf.columns:
            cell_value = zdf.loc[date, village]
            if pd.notna(cell_value) and cell_value != '':
                names_in_cell = parse_cell_content(cell_value)
                
                if person_name in names_in_cell:
                    # 找到该人员，记录日程
                    other_names = [name for name in names_in_cell if name != person_name]
                    schedule.append({
                        '日期': date,
                        '村庄': village,
                        '同行人员': ','.join(other_names) if other_names else '无'
                    })
    
    # 按日期排序
    schedule.sort(key=lambda x: datetime.strptime(x['日期'], '%Y-%m-%d'))
    
    return schedule

def create_person_excel(person_name, schedule, output_folder):
    """为个人创建Excel表格"""
    if not schedule:
        return
    
    # 创建DataFrame
    df = pd.DataFrame(schedule)
    
    # 格式化日期显示
    df['日期'] = pd.to_datetime(df['日期']).dt.strftime('%Y年%m月%d日')
    
    # 创建Excel文件
    filename = os.path.join(output_folder, f"{person_name}.xlsx")
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=person_name, index=False)
        
        # 获取工作表
        worksheet = writer.sheets[person_name]
        
        # 设置列宽
        worksheet.column_dimensions['A'].width = 15
        worksheet.column_dimensions['B'].width = 15
        worksheet.column_dimensions['C'].width = 30
        
        # 设置标题
        worksheet.insert_rows(1)
        worksheet['A1'] = person_name
        worksheet['A1'].font = worksheet['A1'].font.copy(bold=True, size=16)
        worksheet.merge_cells('A1:C1')
    
    print(f"已生成 {filename}")

def create_calendar_reminder(person_name, schedule, output_folder):
    """创建日程提醒文件（ICS格式）"""
    if not schedule:
        return
    
    ics_content = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//村庄访问调度//NONSGML Event//CN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH"
    ]
    
    for item in schedule:
        date_str = item['日期']
        village = item['村庄']
        companions = item['同行人员']
        
        # 解析日期
        event_date = datetime.strptime(date_str, '%Y-%m-%d')
        
        # 设置提醒时间为当天早上7点
        reminder_time = event_date.replace(hour=7, minute=0, second=0)
        
        # 格式化时间为UTC格式
        dtstart = reminder_time.strftime('%Y%m%dT%H%M%S')
        dtend = (reminder_time + timedelta(hours=1)).strftime('%Y%m%dT%H%M%S')
        
        # 创建事件
        event = [
            "BEGIN:VEVENT",
            f"DTSTART:{dtstart}",
            f"DTEND:{dtend}",
            f"SUMMARY:看望{village}",
            f"DESCRIPTION:看望{village}\\n同行人员: {companions}",
            f"UID:{person_name}-{date_str}-{village}@village-schedule",
            "END:VEVENT"
        ]
        
        ics_content.extend(event)
    
    ics_content.append("END:VCALENDAR")
    
    # 保存ICS文件
    filename = os.path.join(output_folder, f"{person_name}_日程提醒.ics")
    with open(filename, 'w', encoding='utf-8') as f:
        f.write('\n'.join(ics_content))
    
    print(f"已生成 {filename}")

def main():
    """主程序"""
    try:
        print("开始生成个人日程表...")
        
        # 创建输出文件夹
        output_folder = create_output_folder()
        
        # 读取总表数据
        zdf = read_schedule_data()
        print(f"读取总表数据: {zdf.shape[0]} 行 × {zdf.shape[1]} 列")
        
        # 从总表中提取所有人员名字
        all_personnel = extract_all_personnel_names(zdf)
        print(f"从总表中提取到 {len(all_personnel)} 个人员名字")
        
        # 为每个人员生成日程
        generated_count = 0
        for person_name in all_personnel:
            schedule = find_person_schedule(person_name, zdf)
            if schedule:
                create_person_excel(person_name, schedule, output_folder)
                create_calendar_reminder(person_name, schedule, output_folder)
                generated_count += 1
        
        print(f"\n已为 {generated_count} 个人员生成日程表")
        print(f"所有文件已保存到 '{output_folder}' 文件夹中")
        
        print("\n程序执行完成！")
        print("生成的文件包括:")
        print("1. 个人Excel表格（人名.xlsx）")
        print("2. 个人日程提醒文件（人名_日程提醒.ics）")
        print(f"3. 所有文件位于 '{output_folder}' 文件夹中")
        
    except Exception as e:
        print(f"程序执行出错: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()