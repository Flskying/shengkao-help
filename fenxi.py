import os
import pandas as pd
import matplotlib.pyplot as plt
import re
from sklearn.linear_model import LinearRegression

# 指定要分析的单位和岗位信息
target_info = [
    ('太原日报社-太原日报社', '管理1'),
    ('太原市发展和改革委员会-太原市粮食技工学校', '专技1'),
    ('太原市教育局-太原市财贸学校', '管理1')
]

# 定义 data 文件夹路径
data_folder = 'data'

# 获取 data 文件夹下所有的 xlsx 文件
xlsx_files = [f for f in os.listdir(data_folder) if f.endswith('.xlsx')]

# 按日期排序文件
xlsx_files.sort()

# 初始化一个空的 DataFrame 用于存储分析结果
analysis_result = pd.DataFrame(columns=['日期', '招聘单位', '岗位类型', '填报信息人数', '初审通过人数', '缴费人数'])

# 遍历每个文件
for file in xlsx_files:
    # 提取日期信息
    date_match = re.search(r'(\d{1,2}\.\d{1,2})', file)
    if date_match:
        date = date_match.group(1)
    else:
        date = '未知日期'
    # 读取文件，指定表头位于第二行
    file_path = os.path.join(data_folder, file)
    df = pd.read_excel(file_path, header=2)

    # 打印文件的列名，方便检查
    print(f"文件 {file} 的列名: {df.columns}")

    # 检查文件是否为空
    if df.empty:
        print(f"文件 {file} 为空，跳过。")
        continue

    new_rows = []
    # 假设列索引 0 为招聘单位，列索引 1 为岗位类型，列索引 3 为填报信息人数，列索引 4 为初审通过人数，列索引 5 为缴费人数
    for index, row in df.iterrows():
        unit = row[0]
        position = row[1]
        total_applicants = row[3]
        passed_applicants = row[4]
        paid_applicants = row[5]

        for target_unit, target_position in target_info:
            if unit == target_unit and position == target_position:
                # 构建新行数据
                new_row = {
                    '日期': date,
                    '招聘单位': unit,
                    '岗位类型': position,
                    '填报信息人数': total_applicants,
                    '初审通过人数': passed_applicants,
                    '缴费人数': paid_applicants
                }
                new_rows.append(new_row)

    # 使用 concat 方法添加新行
    if new_rows:
        new_df = pd.DataFrame(new_rows)
        analysis_result = pd.concat([analysis_result, new_df], ignore_index=True)

# 输出分析结果
print(analysis_result)

# 保存分析结果到一个新的 Excel 文件
analysis_result.to_excel('报名情况分析结果.xlsx', index=False)

# 保存分析结果到一个新的 CSV 文件
analysis_result.to_csv('报名情况分析结果.csv', index=False)

# 设置图片清晰度
plt.rcParams['figure.dpi'] = 300

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

# 创建 fenxi 目录（如果不存在）
if not os.path.exists('fenxi'):
    os.makedirs('fenxi')

# 绘制折线图
for position in analysis_result['岗位类型'].unique():
    position_data = analysis_result[analysis_result['岗位类型'] == position]
    plt.figure(figsize=(10, 6))
    plt.plot(position_data['日期'], position_data['填报信息人数'], label='填报信息人数')
    plt.plot(position_data['日期'], position_data['初审通过人数'], label='初审通过人数')
    plt.plot(position_data['日期'], position_data['缴费人数'], label='缴费人数')
    plt.title(f'{position} 岗位报名情况随时间变化')
    plt.xlabel('日期')
    plt.ylabel('人数')
    plt.legend()
    plt.grid(True)
    plt.xticks(rotation=45)
    plt.tight_layout()
    # 保存折线图到 fenxi 目录
    plt.savefig(os.path.join('fenxi', f'{position}_报名情况随时间变化.png'))
    plt.close()

# 绘制柱状图对比同一日期下两个岗位的报名情况
dates = analysis_result['日期'].unique()
for date in dates:
    date_data = analysis_result[analysis_result['日期'] == date]
    positions = date_data['岗位类型']
    applicants = date_data[['填报信息人数', '初审通过人数', '缴费人数']]

    x = range(len(positions))
    width = 0.2

    fig, ax = plt.subplots(figsize=(10, 6))
    rects1 = ax.bar([i - width for i in x], applicants['填报信息人数'], width, label='填报信息人数')
    rects2 = ax.bar(x, applicants['初审通过人数'], width, label='初审通过人数')
    rects3 = ax.bar([i + width for i in x], applicants['缴费人数'], width, label='缴费人数')

    ax.set_ylabel('人数')
    ax.set_title(f'{date} 不同岗位报名情况对比')
    ax.set_xticks(x)
    ax.set_xticklabels(positions)
    ax.legend()

    def autolabel(rects):
        for rect in rects:
            height = rect.get_height()
            ax.annotate('{}'.format(height),
                        xy=(rect.get_x() + rect.get_width() / 2, height),
                        xytext=(0, 3),  # 3 points vertical offset
                        textcoords="offset points",
                        ha='center', va='bottom')

    autolabel(rects1)
    autolabel(rects2)
    autolabel(rects3)

    fig.tight_layout()
    # 保存柱状图到 fenxi 目录
    plt.savefig(os.path.join('fenxi', f'{date}_不同岗位报名情况对比.png'))
    plt.close()

# 预测功能
# 假设我们要预测未来 3 天的报名情况
future_days = 3

for position in analysis_result['岗位类型'].unique():
    position_data = analysis_result[analysis_result['岗位类型'] == position]

    # 转换日期为数字以便进行线性回归
    position_data['日期数字'] = pd.to_numeric(position_data['日期'].str.replace('.', ''))

    for column in ['填报信息人数', '初审通过人数', '缴费人数']:
        X = position_data[['日期数字']]
        y = position_data[column]

        # 创建线性回归模型
        model = LinearRegression()
        model.fit(X, y)

        # 获取最后一个日期的数字
        last_date_num = position_data['日期数字'].iloc[-1]

        # 生成未来日期的数字
        future_dates_num = [last_date_num + i for i in range(1, future_days + 1)]
        future_dates_num = pd.DataFrame(future_dates_num, columns=['日期数字'])

        # 进行预测
        predictions = model.predict(future_dates_num)

        # 打印预测结果
        print(f'{position} 岗位 {column} 未来 {future_days} 天的预测结果:')
        for i, pred in enumerate(predictions):
            future_date_str = str(int(future_dates_num.iloc[i]['日期数字']))
            future_date_str = f'{future_date_str[:-2]}.{future_date_str[-2:]}'
            print(f'日期: {future_date_str}, 预测人数: {pred:.2f}')

        # 绘制包含预测结果的折线图
        plt.figure(figsize=(10, 6))
        plt.plot(position_data['日期'], position_data[column], label='实际人数')
        future_dates = [str(int(num)).zfill(4) for num in future_dates_num['日期数字']]
        future_dates = [f'{date[:-2]}.{date[-2:]}' for date in future_dates]
        plt.plot(future_dates, predictions, label='预测人数', linestyle='--')
        plt.title(f'{position} 岗位 {column} 报名情况及预测')
        plt.xlabel('日期')
        plt.ylabel('人数')
        plt.legend()
        plt.grid(True)
        plt.xticks(rotation=45)
        plt.tight_layout()
        # 保存包含预测结果的折线图到 fenxi 目录
        plt.savefig(os.path.join('fenxi', f'{position}_{column}_报名情况及预测.png'))
        plt.close()
