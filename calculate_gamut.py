import numpy as np
import pandas as pd
from shapely.geometry import Polygon
import matplotlib.pyplot as plt

# ======================== DCI-P3 标准色域坐标转换 ========================
dci_p3_xy = {'R': (0.68, 0.32), 'G': (0.265, 0.69), 'B': (0.15, 0.06)}
def xy_to_uv_prime(x, y):
    denom = -2 * x + 12 * y + 3
    return (4 * x / denom, 9 * y / denom)

# 创建DCI-P3标准多边形（自动处理顶点顺序）
dci_p3_uv = [xy_to_uv_prime(*dci_p3_xy[key]) for key in ['R', 'G', 'B']]
dci_p3_poly = Polygon(dci_p3_uv)

# ======================== 几何计算函数优化 ========================
def calculate_metrics(target_points):
    try:
        target_poly = Polygon(target_points)
        
        # 计算覆盖率（交集面积）
        intersection = dci_p3_poly.intersection(target_poly)
        coverage = (intersection.area / dci_p3_poly.area) * 100 if dci_p3_poly.area > 0 else 0
        
        # 计算面积比率
        area_ratio = (target_poly.area / dci_p3_poly.area) * 100 if dci_p3_poly.area > 0 else 0
        
        return round(coverage, 2), round(area_ratio, 2)
    except:
        return 0.0, 0.0
    
# ======================== 绘制灰阶和色域范围的两张子图 ========================
def plot_gamut_coverage(data):
    fig, axes = plt.subplots(1, 2, figsize=(12, 6))
    
    # 绘制灰阶图
    axes[0].plot(data['Grayscale'], data['Gamut_Coverage'], 'o-')
    axes[0].set_xlabel('Grayscale')
    axes[0].set_ylabel('Gamut Coverage (%)')
    axes[0].set_title('Gamut Coverage')
    
    # 绘制色域图
    axes[1].plot(data['Grayscale'], data['Gamut_Ratio'], 'o-')
    axes[1].set_xlabel('Grayscale')
    axes[1].set_ylabel('Gamut Ratio (%)')
    axes[1].set_title('Gamut Ratio')
    
    plt.tight_layout()
    plt.savefig('Gamut_Coverage.png')
    # plt.show()

# ======================== 主程序 ========================
if __name__ == "__main__":
    filename = input('请输入数据文件名（含扩展名）: ')
    data = pd.read_excel(filename)
    
    # 初始化结果列
    data['Gamut_Coverage'] = 0.0
    data['Gamut_Ratio'] = 0.0
    
    # 批量处理数据
    for idx, row in data.iterrows():
        target_points = [
            (row['Red_u'], row['Red_v']),
            (row['Green_u'], row['Green_v']),
            (row['Blue_u'], row['Blue_v'])
        ]
        
        coverage, ratio = calculate_metrics(target_points)
        data.at[idx, 'Gamut_Coverage'] = coverage
        data.at[idx, 'Gamut_Ratio'] = ratio
    
    # 保存结果
    output_filename = filename.replace('.xlsx', '_result.xlsx')
    data.to_excel(output_filename, index=False)

    # 绘制结果图
    plot_gamut_coverage(data)

    print(f"处理完成！结果已保存至 {output_filename}")