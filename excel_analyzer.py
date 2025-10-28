import pandas as pd
import os
import sys
from datetime import datetime

def main():
    print("=" * 50)
    print("    Excel比对字段分析工具")
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
        
        # 查找比对字段（修改为"比对"而不是"对比"）
        comparison_cols = [col for col in df.columns if '比对' in str(col)]
        
        if not comparison_cols:
            print("❌ 未找到包含'比对'的列")
            print("可用列名:")
            for col in df.columns:
                print(f"  - {col}")
            input("\n按回车键退出...")
            return
        
        print(f"\n🎯 找到 {len(comparison_cols)} 个比对字段:")
        for col in comparison_cols:
            print(f"  - {col}")
        
        # 分析结果
        print("\n" + "=" * 60)
        print("分析结果:")
        print("=" * 60)
        
        results = []
        total_fails = 0
        total_records = 0
        
        # 创建原始数据的副本，用于添加结果列
        result_df = df.copy()
        
        for col in comparison_cols:
            # 定义不通过的关键词
            fail_keywords = ['不通过', '失败', '不合格', '未通过', '不匹配', '不一致']
            
            # 统计不通过数量
            if df[col].dtype == 'object':
                # 创建布尔掩码，标记不通过的记录
                fail_mask = df[col].astype(str).str.contains(
                    '|'.join(fail_keywords), case=False, na=False
                )
                fail_count = fail_mask.sum()
                
                # 在结果DataFrame中添加状态列
                result_df[f'{col}_状态'] = '通过'
                result_df.loc[fail_mask, f'{col}_状态'] = '不通过'
                
                # 记录不通过的具体值
                fail_values = df.loc[fail_mask, col].unique()
                
            else:
                # 数值类型，假设0表示不通过
                fail_mask = (df[col] == 0)
                fail_count = fail_mask.sum()
                result_df[f'{col}_状态'] = '通过'
                result_df.loc[fail_mask, f'{col}_状态'] = '不通过'
                fail_values = [0] if fail_count > 0 else []
            
            total_count = df[col].notna().sum()
            fail_rate = (fail_count / total_count * 100) if total_count > 0 else 0
            pass_rate = 100 - fail_rate
            
            total_fails += fail_count
            total_records += total_count
            
            print(f"\n📋 {col}:")
            print(f"   ❌ 不通过: {fail_count}/{total_count}")
            print(f"   📊 不通过率: {fail_rate:.2f}%")
            print(f"   ✅ 通过率: {pass_rate:.2f}%")
            
            if fail_count > 0:
                print(f"   🔍 不通过的值: {', '.join(map(str, fail_values[:5]))}")  # 显示前5个不通过的值
                if len(fail_values) > 5:
                    print(f"      ... 还有 {len(fail_values) - 5} 个其他值")
            
            results.append({
                '比对字段': col,
                '不通过数量': fail_count,
                '通过数量': total_count - fail_count,
                '总记录数': total_count,
                '不通过率': f"{fail_rate:.2f}%",
                '通过率': f"{pass_rate:.2f}%"
            })
        
        # 创建统计汇总表
        summary_df = pd.DataFrame(results)
        
        # 汇总统计
        if total_records > 0:
            overall_fail_rate = (total_fails / total_records) * 100
            overall_pass_rate = 100 - overall_fail_rate
            
            print("\n" + "=" * 60)
            print("📈 汇总统计:")
            print(f"   总不通过记录: {total_fails}/{total_records}")
            print(f"   平均不通过率: {overall_fail_rate:.2f}%")
            print(f"   平均通过率: {overall_pass_rate:.2f}%")
            print("=" * 60)
        
        # 保存结果到新的Excel文件
        output_file = os.path.splitext(excel_file)[0] + "_详细分析结果.xlsx"
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Sheet1: 原始数据（带状态列）
            result_df.to_excel(writer, sheet_name='原始数据带状态', index=False)
            
            # Sheet2: 统计汇总
            summary_df.to_excel(writer, sheet_name='统计汇总', index=False)
            
            # Sheet3: 详细分析
            detailed_analysis = []
            for col in comparison_cols:
                if df[col].dtype == 'object':
                    value_counts = df[col].value_counts()
                    for value, count in value_counts.items():
                        status = '不通过' if any(keyword in str(value) for keyword in fail_keywords) else '通过'
                        detailed_analysis.append({
                            '比对字段': col,
                            '具体值': value,
                            '出现次数': count,
                            '状态': status
                        })
            
            if detailed_analysis:
                detailed_df = pd.DataFrame(detailed_analysis)
                detailed_df.to_excel(writer, sheet_name='详细分析', index=False)
            
            # Sheet4: 通过率统计
            pass_rate_summary = summary_df[['比对字段', '通过数量', '不通过数量', '通过率', '不通过率']].copy()
            pass_rate_summary = pass_rate_summary.sort_values('通过率', ascending=True)
            pass_rate_summary.to_excel(writer, sheet_name='通过率排名', index=False)
        
        print(f"\n💾 详细分析结果已保存到: {output_file}")
        print("包含以下工作表:")
        print("  - 原始数据带状态: 原始数据，每列添加了状态列")
        print("  - 统计汇总: 各比对字段的通过/不通过统计")
        print("  - 通过率排名: 按通过率排序的字段排名")
        if detailed_analysis:
            print("  - 详细分析: 每个具体值的状态分析")
        
        # 显示通过率最低的字段
        if not summary_df.empty:
            worst_field = summary_df.loc[summary_df['不通过数量'].idxmax()]
            best_field = summary_df.loc[summary_df['通过数量'].idxmax()]
            
            print(f"\n⚠️  需要关注的字段:")
            print(f"   通过率最低: {worst_field['比对字段']} ({worst_field['不通过率']})")
            print(f"   通过率最高: {best_field['比对字段']} ({best_field['通过率']})")
        
    except Exception as e:
        print(f"❌ 错误: {str(e)}")
        import traceback
        traceback.print_exc()
    
    print("\n🎉 分析完成!")
    input("按回车键退出...")

if __name__ == "__main__":
    main()
