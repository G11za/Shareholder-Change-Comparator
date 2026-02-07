import pandas as pd
import sys
import re
import os


def clean_names_to_set(cell_value):
    """
    清洗单元格数据：去除非法字符，提取人名并转为集合。
    """
    if pd.isna(cell_value) or str(cell_value).strip() == "":
        return set()

    # 使用正则表达式将所有非中文字符、非英文字符的标点替换为空格
    # 这样可以一次性处理：、，, \n \t 以及各种特殊符号
    text = str(cell_value)
    # 仅保留汉字、字母、数字（以防有英文名股东）
    clean_text = re.sub(r'[^\w\u4e00-\u9fa5]', ' ', text)

    # 切分并剔除空格，返回集合
    return {name.strip() for name in clean_text.split() if name.strip()}


def run_comparison(input_path, output_path):
    """
    执行对比任务：自动获取前四列中的后三列进行比对
    """
    print(f"[*] 正在加载文件: {input_path}")

    try:
        # 自动识别引擎读取 Excel
        df = pd.read_excel(input_path)
    except Exception as e:
        print(f"[!] 读取失败: {e}")
        return

    # 检查列数是否满足要求（至少4列：公司名 + 3个周期）
    if df.shape[1] < 4:
        print(f"[!] 错误：表格列数不足 4 列，当前仅有 {df.shape[1]} 列。")
        return

    # 获取参与对比的列名（用于控制台展示）
    compare_cols = df.columns[1:4].tolist()
    print(f"[*] 检测到对比周期: {compare_cols}")

    change_flags = []

    for i in range(len(df)):
        # 动态获取第 2, 3, 4 列数据 (索引为 1, 2, 3)
        data_1 = clean_names_to_set(df.iloc[i, 1])
        data_2 = clean_names_to_set(df.iloc[i, 2])
        data_3 = clean_names_to_set(df.iloc[i, 3])

        # 逻辑判断：只要第一阶段或第二阶段有变化，标记为 1
        if (data_1 != data_2) or (data_2 != data_3):
            change_flags.append(1)
        else:
            change_flags.append(0)

    # 将结果插入到原表最后
    df['has_change'] = change_flags

    try:
        df.to_excel(output_path, index=False)
        print(f"[+] 分析完成！结果已保存至: {output_path}")
        # 打印简单摘要
        print(f"统计：共有 {sum(change_flags)} 家公司发生变动。")
    except Exception as e:
        print(f"[!] 写入文件失败: {e}")


if __name__ == "__main__":
    # 配置区
    INPUT_FILE = "name.xlsx"
    OUTPUT_FILE = "shareholder_analysis_result.xlsx"

    # 环境依赖检查
    try:
        import openpyxl
    except ImportError:
        print("[!] 运行失败：请先安装依赖库 'pip install openpyxl'")
        sys.exit(1)

    if os.path.exists(INPUT_FILE):
        run_comparison(INPUT_FILE, OUTPUT_FILE)
    else:
        print(f"[!] 未找到输入文件 '{INPUT_FILE}'，请确认文件是否存在于当前目录下。")