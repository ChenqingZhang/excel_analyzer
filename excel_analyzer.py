import pandas as pd
import os
import sys

def main():
    print("=" * 50)
    print("    Excel对比字段分析工具 - Windows版本")
    print("=" * 50)
    
    # 检查当前目录下的Excel文件
    excel_files = []
    for file in os.listdir('.'):
        if file.lower().endswith(('.xlsx', '.xls')):
            excel_files.append(file)
    
    if not excel_files:
        print("❌ 在当前目录未找到Excel文件")
        print("请将Excel文件放在与此程序相同的目录下")
        input("按回车键退出...")
        return
    
    # 显示找到的Excel文件
    print("找到以下Excel文件:")
    for i, file in enumerate(excel_files, 1):
        print(f"  {i}. {file}")
    
    # 选择文件
    try:
        choice = int(input("\n请选择要分析的文件编号: ")) - 1
        if choice < 0 or choice >= len(excel_files):
            print("❌ 无效的选择")
            return
        
        excel_file = excel_files[choice]
        print(f"📊 正在分析: {excel_file}")
        
    except ValueError:
        print("❌ 请输入有效的数字")
        return
    
    try:
        # 读取Excel文件
        df = pd.read_excel(excel_file)
        print(f"✅ 成功读取文件，共 {len(df)} 行数据")
        
        # 查找包含"对比"的列
        comparison_columns = []
        for col in df.columns:
            if '对比' in str(col):
                comparison_columns.append(col)
        
        if not comparison_columns:
            print("❌ 未找到包含'对比'字段的列")
            input("按回车键退出...")
            return
        
        print(f"\n🎯 找到 {len(comparison_columns)} 个包含'对比'的列:")
        for col in comparison_columns:
            print(f"   - {col}")
        
        # 分析每个对比列
        print("\n" + "=" * 60)
        print("分析结果:")
        print("=" * 60)
        
        results = []
        total_fail_count = 0
        total_record_count = 0
        
        for column in comparison_columns:
            # 统计不通过数量
            if df[column].dtype == 'object':  # 文本类型
                fail_conditions = ['不通过', '失败', '不合格', '未通过', 'Fail', 'FAIL']
                fail_count = df[column].astype(str).str.contains(
                    '|'.join(fail_conditions), case=False, na=False
                ).sum()
            else:  # 数值类型
                fail_count = (df[column] == 0).sum()
            
            total_count = len(df[column].dropna())  # 非空记录数
            fail_percentage = (fail_count / total_count * 100) if total_count > 0 else 0
            
            total_fail_count += fail_count
            total_record_count += total_count
            
            print(f"\n📋 列名: {column}")
            print(f"   ❌ 不通过数量: {fail_count}/{total_count}")
            print(f"   📊 不通过率: {fail_percentage:.2f}%")
            print(f"   ✅ 通过率: {100-fail_percentage:.2f}%")
            
            results.append({
                '列名': column,
                '不通过数量': fail_count,
                '总记录数': total_count,
                '不通过率': f"{fail_percentage:.2f}%",
                '通过率': f"{100-fail_percentage:.2f}%"
            })
        
        # 汇总信息
        if total_record_count > 0:
            overall_percentage = (total_fail_count / total_record_count) * 100
            print("\n" + "=" * 60)
            print("📈 汇总信息:")
            print(f"   总不通过记录: {total_fail_count}/{total_record_count}")
            print(f"   平均不通过率: {overall_percentage:.2f}%")
            print("=" * 60)
        
        # 保存结果到文本文件
        output_file = os.path.splitext(excel_file)[0] + "_分析结果.txt"
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("Excel对比字段分析结果\\n")
            f.write("=" * 50 + "\\n")
            for result in results:
                f.write(f"列名: {result['列名']}\\n")
                f.write(f"  不通过数量: {result['不通过数量']}/{result['总记录数']}\\n")
                f.write(f"  不通过率: {result['不通过率']}\\n")
                f.write(f"  通过率: {result['通过率']}\\n")
                f.write("-" * 40 + "\\n")
        
        print(f"\n💾 分析结果已保存到: {output_file}")
        
    except Exception as e:
        print(f"❌ 分析过程中出现错误: {str(e)}")
    
    print("\n🎉 分析完成！")
    input("按回车键退出...")

if __name__ == "__main__":
    main()