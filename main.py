import sys

import pandas as pd
from pathlib import Path

# ==========================================
# 请在这里修改你的配置
# ==========================================
# 兼容打包后和源码运行两种模式
if getattr(sys, 'frozen', False):
    # 打包成 .app 或 .exe 后，用可执行文件所在目录
    SCRIPT_DIR = Path(sys.executable).parent
else:
    # 直接在源代码运行时，用脚本文件所在目录
    SCRIPT_DIR = Path(__file__).parent

EXCEL_FOLDER = SCRIPT_DIR / "test" # 直接使用脚本所在目录作为搜索文件夹
# 比如：EXCEL_FOLDER = Path("C:/Users/你的名字/Desktop/销售报表")
OUTPUT_FILE = "result.xlsx"  # 最后生成的文件名
HEADER_ROW = 0  # 0代表第一行是表头，如果没表头就写None


# ==========================================

def main():
    # 1. 列出文件夹下所有Excel文件
    file_list = list(EXCEL_FOLDER.glob("*.xlsx")) + list(EXCEL_FOLDER.glob("*.xls"))

    if not file_list:
        print("❌ 没找到任何Excel文件，请检查文件夹路径~")
        return

    print(f"✅ 找到 {len(file_list)} 个文件，开始合并...")
    all_data = []  # 用来装所有数据

    for file in file_list:
        try:
            # 2. 读取每个Excel的第一个sheet
            df = pd.read_excel(file, header=HEADER_ROW)
            # 3. 可以选择加一列“来源文件”，方便溯源
            df["来源文件"] = file.name
            all_data.append(df)
            print(f"  -> 已读取: {file.name}")
        except Exception as e:
            print(f"  ❌ 读取失败: {file.name}，原因: {e}")

    if all_data:
        # 4. 纵向拼接所有表格
        merged_df = pd.concat(all_data, ignore_index=True)
        # 5. 保存结果
        merged_df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
        print(f"🎉 大功告成！已为你合并了 {len(all_data)} 个文件，结果保存在 >>> {OUTPUT_FILE}")
        print(f"    总行数（含表头）: {merged_df.shape[0]}")
    else:
        print("❌ 没有成功读取到任何数据。")


if __name__ == "__main__":
    main()
    input("按回车键退出...")  # 加上这行
