import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta
import os

def get_next_quarter_saturdays():
    """获取下一个季度的所有星期六日期"""
    now = datetime.now()
    
    # 确定下一个季度的开始年份和月份
    current_month = now.month
    current_year = now.year
    
    if current_month <= 3:
        next_quarter_start = 4
        quarter_year = current_year
    elif current_month <= 6:
        next_quarter_start = 7
        quarter_year = current_year
    elif current_month <= 9:
        next_quarter_start = 10
        quarter_year = current_year
    else:
        next_quarter_start = 1
        quarter_year = current_year + 1
    
    # 下一个季度的第一天
    quarter_start = datetime(quarter_year, next_quarter_start, 1)
    
    # 计算季度结束日期
    if next_quarter_start == 1:
        quarter_end = datetime(quarter_year, 3, 31)
    elif next_quarter_start == 4:
        quarter_end = datetime(quarter_year, 6, 30)
    elif next_quarter_start == 7:
        quarter_end = datetime(quarter_year, 9, 30)
    else:  # next_quarter_start == 10
        quarter_end = datetime(quarter_year, 12, 31)
    
    # 找到第一个星期六
    days_to_saturday = (5 - quarter_start.weekday()) % 7
    first_saturday = quarter_start + timedelta(days=days_to_saturday)
    
    # 生成该季度的所有星期六
    saturdays = []
    current_saturday = first_saturday
    
    while current_saturday <= quarter_end:
        saturdays.append(current_saturday.strftime('%Y-%m-%d'))
        current_saturday += timedelta(days=7)
    
    return saturdays

def read_village_data():
    """读取村庄数据"""
    if os.path.exists('村.csv'):
        df = pd.read_csv('村.csv', encoding='utf-8')
    elif os.path.exists('村.xlsx'):
        df = pd.read_excel('村.xlsx')
    else:
        raise FileNotFoundError("未找到村.csv或村.xlsx文件")
    
    return df

def read_personnel_data():
    """读取人员数据"""
    if os.path.exists('看望人员.csv'):
        df = pd.read_csv('看望人员.csv', encoding='utf-8')
    elif os.path.exists('看望人员.xlsx'):
        df = pd.read_excel('看望人员.xlsx')
    else:
        raise FileNotFoundError("未找到看望人员.csv或看望人员.xlsx文件")
    
    return df

def create_empty_dataframes(village_df, saturdays):
    """创建两个空的DataFrame"""
    village_names = village_df['村名'].tolist()
    
    zdf1 = pd.DataFrame(index=saturdays, columns=village_names)
    zdf2 = pd.DataFrame(index=saturdays, columns=village_names)
    
    return zdf1, zdf2, village_names

def fill_zdf1_with_village_people(zdf1, village_df, village_names):
    """用村庄接待人员填充zdf1"""
    for village in village_names:
        # 获取该村的接待人员
        village_row = village_df[village_df['村名'] == village]
        if not village_row.empty:
            # 提取该村的所有接待人员（去除空值）
            people = []
            for col in village_row.columns[1:]:  # 跳过村名列
                person = village_row[col].iloc[0]
                if pd.notna(person) and person != '':
                    people.append(person)
            
            # 如果有接待人员，循环填充该村的所有日期
            if people:
                for i, date in enumerate(zdf1.index):
                    zdf1.loc[date, village] = people[i % len(people)]
    
    return zdf1

def fill_zdf2_with_visit_people_relaxed(zdf2, kwdf, village_names):
    """用看望人员填充zdf2 - 宽松模式，允许重复访问但尽量隔开时间"""
    a = kwdf['计划出访次数'].max()
    num_villages = len(village_names)
    num_dates = len(zdf2.index)
    
    # 创建工作副本
    kwdf_work = kwdf.copy()
    
    # 为每个村建立访问记录，记录每个团队的最后访问时间
    village_visit_record = {}
    for village in village_names:
        village_visit_record[village] = {}
    
    # 按顺序填充每个时间点
    for date_idx, date in enumerate(zdf2.index):
        # 获取当前日期已分配的团队，避免同一天重复
        daily_assigned_teams = set()
        
        for village in village_names:
            # 找到可用的团队（还有出访次数的）
            available_teams = kwdf_work[kwdf_work['计划出访次数'] > 0]
            
            if available_teams.empty:
                continue
            
            # 计算每个团队的适合度得分
            team_scores = []
            for idx, row in available_teams.iterrows():
                team = f"{row['人员1']},{row['人员2']}"
                
                # 如果这个团队今天已经被分配过，跳过
                if team in daily_assigned_teams:
                    continue
                
                score = calculate_team_score(team, village, date_idx, village_visit_record, num_dates)
                team_scores.append((idx, team, score))
            
            if not team_scores:
                continue
            
            # 按得分排序，选择最适合的团队
            team_scores.sort(key=lambda x: x[2], reverse=True)
            
            # 选择得分最高的团队
            selected_idx, selected_team, selected_score = team_scores[0]
            
            # 分配团队
            zdf2.loc[date, village] = selected_team
            daily_assigned_teams.add(selected_team)
            
            # 更新出访次数
            kwdf_work.loc[selected_idx, '计划出访次数'] -= 1
            
            # 更新访问记录
            village_visit_record[village][selected_team] = date_idx
    
    return zdf2, a, num_villages

def calculate_team_score(team, village, current_date_idx, village_visit_record, num_dates):
    """计算团队在特定村庄和日期的适合度得分"""
    score = 100  # 基础分数
    
    # 检查该团队是否之前访问过这个村庄
    if team in village_visit_record[village]:
        last_visit_idx = village_visit_record[village][team]
        time_gap = current_date_idx - last_visit_idx
        
        # 时间间隔越大，得分越高
        if time_gap == 1:
            score -= 80  # 连续访问，大幅降分
        elif time_gap == 2:
            score -= 50  # 隔一周，中等降分
        elif time_gap == 3:
            score -= 20  # 隔两周，小幅降分
        else:
            score += 10  # 间隔足够长，加分
    else:
        # 首次访问该村庄，加分
        score += 20
    
    # 根据总体时间分布调整得分
    # 如果是季度开始或结束，稍微降低得分，让访问更均匀
    if current_date_idx < num_dates * 0.2 or current_date_idx > num_dates * 0.8:
        score -= 5
    
    # 添加随机因子，避免过度确定性
    score += random.randint(-10, 10)
    
    return score

def optimize_schedule_relaxed(zdf2, max_iterations=500):
    """优化调度，减少时间间隔过短的情况"""
    
    for iteration in range(max_iterations):
        improved = False
        
        # 找到所有可能的改进机会
        improvement_opportunities = find_improvement_opportunities(zdf2)
        
        if not improvement_opportunities:
            break
        
        # 随机选择一个改进机会
        opportunity = random.choice(improvement_opportunities)
        
        # 尝试改进
        if apply_improvement(zdf2, opportunity):
            improved = True
        
        if not improved:
            break
    
    return zdf2

def find_improvement_opportunities(zdf2):
    """找到可以改进的机会（时间间隔过短的情况）"""
    opportunities = []
    
    # 检查每个村庄的访问模式
    for village_idx, village in enumerate(zdf2.columns):
        team_visits = {}
        
        # 收集每个团队的访问时间
        for date_idx, date in enumerate(zdf2.index):
            team = zdf2.loc[date, village]
            if pd.notna(team) and team != '':
                if team not in team_visits:
                    team_visits[team] = []
                team_visits[team].append(date_idx)
        
        # 检查是否有时间间隔过短的情况
        for team, visits in team_visits.items():
            if len(visits) > 1:
                visits.sort()
                for i in range(len(visits) - 1):
                    if visits[i+1] - visits[i] <= 2:  # 间隔小于等于2周
                        opportunities.append({
                            'type': 'short_interval',
                            'village': village,
                            'team': team,
                            'date1': visits[i],
                            'date2': visits[i+1],
                            'interval': visits[i+1] - visits[i]
                        })
    
    return opportunities

def apply_improvement(zdf2, opportunity):
    """应用改进方案"""
    if opportunity['type'] == 'short_interval':
        return improve_short_interval(zdf2, opportunity)
    
    return False

def improve_short_interval(zdf2, opportunity):
    """改进时间间隔过短的情况"""
    village = opportunity['village']
    team = opportunity['team']
    date1_idx = opportunity['date1']
    date2_idx = opportunity['date2']
    
    date1 = zdf2.index[date1_idx]
    date2 = zdf2.index[date2_idx]
    
    # 尝试将第二次访问移动到其他时间
    for try_date_idx, try_date in enumerate(zdf2.index):
        if try_date_idx == date1_idx or try_date_idx == date2_idx:
            continue
        
        # 检查是否可以交换
        if can_move_team_to_date(zdf2, team, village, date2, try_date):
            # 执行移动
            zdf2.loc[try_date, village] = team
            zdf2.loc[date2, village] = None
            return True
    
    return False

def can_move_team_to_date(zdf2, team, village, from_date, to_date):
    """检查团队是否可以移动到新的日期"""
    # 检查目标日期是否已经有安排
    if pd.notna(zdf2.loc[to_date, village]) and zdf2.loc[to_date, village] != '':
        return False
    
    # 检查该团队在目标日期是否已经安排了其他村庄
    for other_village in zdf2.columns:
        if other_village != village:
            if pd.notna(zdf2.loc[to_date, other_village]) and zdf2.loc[to_date, other_village] == team:
                return False
    
    return True

def merge_dataframes(zdf1, zdf2):
    """合并两个DataFrame，把接待人员放在看望人员后面"""
    zdf = pd.DataFrame(index=zdf1.index, columns=zdf1.columns)
    
    for i in range(len(zdf1.index)):
        for j in range(len(zdf1.columns)):
            reception_person = zdf1.iloc[i, j] if pd.notna(zdf1.iloc[i, j]) else ""
            visit_people = zdf2.iloc[i, j] if pd.notna(zdf2.iloc[i, j]) else ""
            
            if visit_people and reception_person:
                zdf.iloc[i, j] = f"{visit_people},{reception_person}"
            elif visit_people:
                zdf.iloc[i, j] = visit_people
            elif reception_person:
                zdf.iloc[i, j] = reception_person
    
    return zdf

def analyze_schedule_quality(zdf2):
    """分析调度质量"""
    print("\n=== 调度质量分析 ===")
    
    total_short_intervals = 0
    total_visits = 0
    
    for village in zdf2.columns:
        team_visits = {}
        village_visits = 0
        
        # 收集每个团队的访问时间
        for date_idx, date in enumerate(zdf2.index):
            team = zdf2.loc[date, village]
            if pd.notna(team) and team != '':
                village_visits += 1
                total_visits += 1
                if team not in team_visits:
                    team_visits[team] = []
                team_visits[team].append(date_idx)
        
        # 分析该村庄的访问模式
        short_intervals = 0
        for team, visits in team_visits.items():
            if len(visits) > 1:
                visits.sort()
                for i in range(len(visits) - 1):
                    interval = visits[i+1] - visits[i]
                    if interval <= 2:
                        short_intervals += 1
        
        total_short_intervals += short_intervals
        
        if village_visits > 0:
            print(f"{village}: {village_visits}次访问, {short_intervals}次短间隔")
    
    print(f"\n总计: {total_visits}次访问, {total_short_intervals}次短间隔")
    if total_visits > 0:
        print(f"短间隔比例: {total_short_intervals/total_visits*100:.1f}%")

def main():
    """主程序"""
    try:
        # 获取下一个季度的星期六
        saturdays = get_next_quarter_saturdays()
        print(f"下一个季度的星期六日期: {saturdays}")
        
        # 读取村庄数据
        village_df = read_village_data()
        print(f"读取到 {len(village_df)} 个村庄")
        
        # 读取人员数据
        kwdf = read_personnel_data()
        print(f"读取到 {len(kwdf)} 组看望人员")
        
        # 创建空的DataFrame
        zdf1, zdf2, village_names = create_empty_dataframes(village_df, saturdays)
        print(f"创建了 {len(saturdays)} 行 × {len(village_names)} 列的DataFrame")
        
        # 填充zdf1
        zdf1 = fill_zdf1_with_village_people(zdf1, village_df, village_names)
        print("已填充接待人员到zdf1")
        
        # 填充zdf2 - 宽松模式
        zdf2, a, num_villages = fill_zdf2_with_visit_people_relaxed(zdf2, kwdf, village_names)
        print(f"已填充看望人员到zdf2（宽松模式）, 最大出访次数a={a}, 村庄数量={num_villages}")
        
        # 优化调度
        zdf2 = optimize_schedule_relaxed(zdf2)
        print("已完成调度优化")
        
        # 分析调度质量
        analyze_schedule_quality(zdf2)
        
        # 合并DataFrame
        zdf = merge_dataframes(zdf1, zdf2)
        print("已合并DataFrame")
        
        # 保存为Excel文件
        zdf.to_excel('总表.xlsx', engine='openpyxl')
        print("已保存为总表.xlsx")
        
        # 显示结果预览
        print("\n结果预览:")
        print(zdf.head())
        
    except Exception as e:
        print(f"程序执行出错: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()