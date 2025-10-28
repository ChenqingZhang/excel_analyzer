import pandas as pd
import os
import sys

def main():
    print("=" * 50)
    print("    Excel对比字段分析工具")
    print("=" * 50)
    
    # 获取当前目录
    current_dir = os.getcwd()
    print(f"当前目录: {current_dir}")
    
    # 查找Excel文件
    excel_files = []
    for file in os.listdir('.'):
        if file.lower().endswith(('.xlsx', '.xls')):
            excel_files.append(file)
    
    if not excel_files:
        print("\n❌ 未找到Excel文件")
        print("请将Excel文件放在此程序相同目录下")
        input("\n按回车键退出...")
        return
    
    print(f"\n📁 找到 {len(excel_files)} 个Excel文件:")
    for i, file in enumerate(excel_files, 1):
        print(f"  {i}. {file}")
    
    # 选择文件
    try:
        choice = int(input("\n请选择要分析的文件编号: ")) - 1
        if choice < 0 or choice >= len(excel_files):
            print("❌ 无效的选择")
            input("按回车键退出...")
            return
        
        excel_file = excel_files[choice]
        print(f"\n📊 正在分析: {excel_file}")
        
    except ValueError:
        print("❌ 请输入有效的数字")
        input("按回车键退出...")
        return
    
    try:
        # 读取Excel
        df = pd.read_excel(excel_file)
        print(f"✅ 成功读取，共 {len(df)} 行 {len(df.columns)} 列")
        
        # 查找对比字段
        comparison_cols = [col for col in df.columns if '对比' in str(col)]
        
        if not comparison_cols:
            print("❌ 未找到包含'对比'的列")
            print("可用列名:")
            for col in df.columns:
                print(f"  - {col}")
            input("\n按回车键退出...")
            return
        
        print(f"\n🎯 找到 {len(comparison_cols)} 个对比字段:")
        for col in comparison_cols:
            print(f"  - {col}")
        
        # 分析结果
        print("\n" + "=" * 60)
        print("分析结果:")
        print("=" * 60)
        
        results = []
        total_fails = 0
        total_records = 0
        
        for col in comparison_cols:
            # 统计不通过
            if df[col].dtype == 'object':
                fail_keywords = ['不通过', '失败', '不合格', '未通过']
                fail_count = df[col].astype(str).str.contains(
                    '|'.join(fail_keywords), case=False, na=False
                ).sum()
            else:
                fail_count = (df[col] == 0).sum()
            
            total_count = df[col].notna().sum()
            fail_rate = (fail_count / total_count * 100) if total_count > 0 else 0
            
            total_fails += fail_count
            total_records += total_count
            
            print(f"\n📋 {col}:")
            print(f"   ❌ 不通过: {fail_count}/{total_count}")
            print(f"   📊 不通过率: {fail_rate:.2f}%")
            print(f"   ✅ 通过率: {100-fail_rate:.2f}%")
            
            results.append({
                '字段': col,
                '不通过数': fail_count,
                '总记录': total_count,
                '不通过率': f"{fail_rate:.2f}%"
            })
        
        # 汇总
        if total_records > 0:
            overall_rate = (total_fails / total_records) * 100
            print("\n" + "=" * 60)
            print("📈 汇总统计:")
            print(f"   总不通过记录: {total_fails}/{total_records}")
            print(f"   平均不通过率: {overall_rate:.2f}%")
            print("=" * 60)
        
        # 保存结果
        output_file = os.path.splitext(excel_file)[0] + "_分析结果.txt"
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("Excel对比字段分析报告\\n")
            f.write("=" * 50 + "\\n\\n")
            for result in results:
                f.write(f"字段: {result['字段']}\\n")
                f.write(f"  不通过数: {result['不通过数']}/{result['总记录']}\\n")
                f.write(f"  不通过率: {result['不通过率']}\\n\\n")
            f.write(f"总不通过率: {overall_rate:.2f}%\\n")
        
        print(f"\n💾 详细结果已保存到: {output_file}")
        
    except Exception as e:
        print(f"❌ 错误: {str(e)}")
    
    print("\n🎉 分析完成!")
    input("按回车键退出...")

if __name__ == "__main__":
    main()
