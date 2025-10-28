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
        
        # 查找比对字段
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
        all_fail_records = []  # 存储所有不通过的记录
        
        # 定义不通过的关键词
        fail_keywords = ['不通过', '失败', '不合格', '未通过', '不匹配', '不一致']
        
        for col in comparison_cols:
            # 统计不通过数量
            if df[col].dtype == 'object':
                # 创建布尔掩码，标记不通过的记录
                fail_mask = df[col].astype(str).str.contains(
                    '|'.join(fail_keywords), case=False, na=False
                )
                fail_count = fail_mask.sum()
                
                # 获取不通过的记录
                fail_records = df[fail_mask].copy()
                if not fail_records.empty:
                    # 添加标识列，说明为什么不通过
                    fail_records['不通过原因'] = f"字段[{col}]值不通过"
                    all_fail_records.append(fail_records)
                
            else:
                # 数值类型，假设0表示不通过
                fail_mask = (df[col] == 0)
                fail_count = fail_mask.sum()
                
                # 获取不通过的记录
                fail_records = df[fail_mask].copy()
                if not fail_records.empty:
                    fail_records['不通过原因'] = f"字段[{col}]值为0"
                    all_fail_records.append(fail_records)
            
            total_count = df[col].notna().sum()
            fail_rate = (fail_count / total_count * 100) if total_count > 0 else 0
            pass_rate = 100 - fail_rate
            
            total_fails += fail_count
            total_records += total_count
            
            print(f"\n📋 {col}:")
            print(f"   ❌ 不通过: {fail_count}/{total_count}")
            print(f"   📊 不通过率: {fail_rate:.2f}%")
            print(f"   ✅ 通过率: {pass_rate:.2f}%")
            
            # 显示不通过的具体值样例
            if fail_count > 0:
                fail_values = df.loc[fail_mask, col].dropna().unique()
                sample_values = fail_values[:3]  # 显示前3个不通过的值
                print(f"   🔍 不通过值样例: {', '.join(map(str, sample_values))}")
                if len(fail_values) > 3:
                    print(f"      ... 还有 {len(fail_values) - 3} 个其他值")
            
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
        
        # 合并所有不通过记录
        combined_fail_records = pd.DataFrame()
        if all_fail_records:
            combined_fail_records = pd.concat(all_fail_records, ignore_index=True)
            # 去除重复记录（同一条记录可能因为多个字段不通过而被多次记录）
            combined_fail_records = combined_fail_records.drop_duplicates()
        
        # 汇总统计
        if total_records > 0:
            overall_fail_rate = (total_fails / total_records) * 100
            overall_pass_rate = 100 - overall_fail_rate
            
            print("\n" + "=" * 60)
            print("📈 汇总统计:")
            print(f"   总不通过记录: {total_fails}/{total_records}")
            print(f"   平均不通过率: {overall_fail_rate:.2f}%")
            print(f"   平均通过率: {overall_pass_rate:.2f}%")
            print(f"   不通过记录数: {len(combined_fail_records)} 条")
            print("=" * 60)
        
        # 保存结果到新的Excel文件
        output_file = os.path.splitext(excel_file)[0] + "_分析报告.xlsx"
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Sheet1: 原始数据（保持不变）
            df.to_excel(writer, sheet_name='原始数据', index=False)
            
            # Sheet2: 统计汇总
            summary_df.to_excel(writer, sheet_name='统计汇总', index=False)
            
            # Sheet3: 不通过记录
            if not combined_fail_records.empty:
                combined_fail_records.to_excel(writer, sheet_name='不通过记录', index=False)
            else:
                # 如果没有不通过记录，创建一个空表
                pd.DataFrame({'说明': ['没有不通过记录']}).to_excel(writer, sheet_name='不通过记录', index=False)
            
            # Sheet4: 通过率排名
            pass_rate_summary = summary_df[['比对字段', '通过数量', '不通过数量', '通过率', '不通过率']].copy()
            # 将百分比转换为数值用于排序
            pass_rate_summary['通过率数值'] = pass_rate_summary['通过率'].str.rstrip('%').astype(float)
            pass_rate_summary = pass_rate_summary.sort_values('通过率数值', ascending=True)
            pass_rate_summary.drop('通过率数值', axis=1, inplace=True)
            pass_rate_summary.to_excel(writer, sheet_name='通过率排名', index=False)
            
            # Sheet5: 字段详情分析
            field_details = []
            for col in comparison_cols:
                if df[col].dtype == 'object':
                    value_counts = df[col].value_counts().head(10)  # 只显示前10个最常见的值
                    for value, count in value_counts.items():
                        status = '不通过' if any(keyword in str(value) for keyword in fail_keywords) else '通过'
                        field_details.append({
                            '比对字段': col,
                            '具体值': value,
                            '出现次数': count,
                            '占比': f"{(count/len(df))*100:.1f}%",
                            '状态': status
                        })
            
            if field_details:
                field_details_df = pd.DataFrame(field_details)
                field_details_df.to_excel(writer, sheet_name='字段详情', index=False)
        
        print(f"\n💾 分析报告已保存到: {output_file}")
        print("包含以下工作表:")
        print("  - 原始数据: 完整的原始数据（未修改）")
        print("  - 统计汇总: 各比对字段的通过/不通过统计")
        print("  - 不通过记录: 所有不通过的记录及其原因")
        print("  - 通过率排名: 按通过率排序的字段排名")
        print("  - 字段详情: 各字段具体值的分布情况")
        
        # 显示关键信息
        if not summary_df.empty:
            worst_field = summary_df.loc[summary_df['不通过数量'].idxmax()]
            best_field = summary_df.loc[summary_df['通过数量'].idxmax()]
            
            print(f"\n⚠️  关键发现:")
            print(f"   需要重点关注的字段: {worst_field['比对字段']} (不通过率: {worst_field['不通过率']})")
            print(f"   表现最佳的字段: {best_field['比对字段']} (通过率: {best_field['通过率']})")
            
            if not combined_fail_records.empty:
                print(f"   共发现 {len(combined_fail_records)} 条不通过记录，详见'不通过记录'工作表")
        
    except Exception as e:
        print(f"❌ 错误: {str(e)}")
        import traceback
        traceback.print_exc()
    
    print("\n🎉 分析完成!")
    input("按回车键退出...")

if __name__ == "__main__":
    main()
