import pandas as pd
import os
import sys
from datetime import datetime

def ask_for_detailed_analysis():
    """询问用户是否需要详细原因分析"""
    try:
        import tkinter as tk
        from tkinter import messagebox
        
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        
        response = messagebox.askyesno(
            "分析选项",
            "是否需要进行详细原因分析？\n\n" +
            "✅ 详细分析：分析每条不通过记录的具体原因（速度较慢）\n" +
            "❌ 基础分析：只统计通过率和不通过数量（速度较快）\n\n" +
            "建议：\n" +
            "- 数据量小或需要详细原因时选择【是】\n" +
            "- 数据量大或只需概览时选择【否】"
        )
        
        root.destroy()
        return response
        
    except Exception:
        # 如果GUI不可用，使用命令行询问
        print("\n" + "=" * 60)
        print("分析选项")
        print("=" * 60)
        print("请选择分析模式：")
        print("1. 详细分析 - 分析每条不通过记录的具体原因（速度较慢）")
        print("2. 基础分析 - 只统计通过率和不通过数量（速度较快）")
        
        while True:
            try:
                choice = input("\n请选择 (1/2): ").strip()
                if choice == '1':
                    return True
                elif choice == '2':
                    return False
                else:
                    print("❌ 请输入 1 或 2")
            except KeyboardInterrupt:
                print("\n用户取消操作")
                sys.exit(0)

def basic_analysis(df, comparison_cols):
    """基础分析 - 只统计数量，不分析具体原因"""
    results = []
    total_fails = 0
    total_records = 0
    
    # 定义不通过的关键词
    fail_keywords = ['不通过', '失败', '不合格', '未通过', '不匹配', '不一致']
    
    print("\n🔍 正在进行基础分析...")
    
    for col in comparison_cols:
        # 统计不通过数量
        if df[col].dtype == 'object':
            fail_mask = df[col].astype(str).str.contains(
                '|'.join(fail_keywords), case=False, na=False
            )
            fail_count = fail_mask.sum()
        else:
            fail_mask = (df[col] == 0)
            fail_count = fail_mask.sum()
        
        total_count = df[col].notna().sum()
        fail_rate = (fail_count / total_count * 100) if total_count > 0 else 0
        pass_rate = 100 - fail_rate
        
        total_fails += fail_count
        total_records += total_count
        
        print(f"📋 {col}: {fail_count}/{total_count} 不通过 ({fail_rate:.2f}%)")
        
        results.append({
            '比对字段': col,
            '不通过数量': fail_count,
            '通过数量': total_count - fail_count,
            '总记录数': total_count,
            '不通过率': f"{fail_rate:.2f}%",
            '通过率': f"{pass_rate:.2f}%"
        })
    
    # 基础分析模式下返回空的 detailed_analysis
    detailed_analysis = []
    
    return results, total_fails, total_records, detailed_analysis

def detailed_analysis(df, comparison_cols):
    """详细分析 - 分析每条记录的具体原因"""
    results = []
    detailed_analysis = []
    total_fails = 0
    total_records = 0
    
    # 定义不通过的关键词
    fail_keywords = ['不通过', '失败', '不合格', '未通过', '不匹配', '不一致']
    
    print("\n🔍 正在进行详细原因分析，请稍候...")
    
    for i, col in enumerate(comparison_cols):
        print(f"分析进度: {i+1}/{len(comparison_cols)} - {col}")
        
        # 统计不通过数量
        if df[col].dtype == 'object':
            fail_mask = df[col].astype(str).str.contains(
                '|'.join(fail_keywords), case=False, na=False
            )
            fail_count = fail_mask.sum()
        else:
            fail_mask = (df[col] == 0)
            fail_count = fail_mask.sum()
        
        total_count = df[col].notna().sum()
        fail_rate = (fail_count / total_count * 100) if total_count > 0 else 0
        pass_rate = 100 - fail_rate
        
        total_fails += fail_count
        total_records += total_count
        
        # 基础统计
        results.append({
            '比对字段': col,
            '不通过数量': fail_count,
            '通过数量': total_count - fail_count,
            '总记录数': total_count,
            '不通过率': f"{fail_rate:.2f}%",
            '通过率': f"{pass_rate:.2f}%"
        })
        
        # 详细原因分析
        if fail_count > 0:
            fail_records = df[fail_mask]
            
            # 分析字段类型和对应的原因
            if '_比对' in col:
                # 这是比对结果字段，分析对应的新旧值字段
                field_prefix = col.replace('_比对', '')
                new_col = f'新值_{field_prefix}'
                old_col = f'旧值_{field_prefix}'
                
                if new_col in df.columns and old_col in df.columns:
                    # 分析新旧值差异
                    for idx, row in fail_records.iterrows():
                        new_val = row[new_col]
                        old_val = row[old_col]
                        
                        reason = "比对不通过"
                        if pd.isna(new_val) and not pd.isna(old_val):
                            reason = "新值为空，旧值有数据"
                        elif not pd.isna(new_val) and pd.isna(old_val):
                            reason = "旧值为空，新值有数据"
                        elif pd.isna(new_val) and pd.isna(old_val):
                            reason = "新旧值都为空"
                        elif str(new_val) != str(old_val):
                            reason = f"数值不一致: 新值={new_val}, 旧值={old_val}"
                        else:
                            reason = "标记为不通过但数值相同"
                        
                        detailed_analysis.append({
                            '比对字段': col,
                            '记录行号': idx + 2,
                            '新值': new_val,
                            '旧值': old_val,
                            '比对结果': row[col],
                            '不通过原因': reason
                        })
                else:
                    # 没有找到对应的新旧值字段
                    for idx, row in fail_records.iterrows():
                        detailed_analysis.append({
                            '比对字段': col,
                            '记录行号': idx + 2,
                            '新值': 'N/A',
                            '旧值': 'N/A',
                            '比对结果': row[col],
                            '不通过原因': '无对应新旧值字段'
                        })
            else:
                # 非比对字段的不通过原因分析
                for idx, row in fail_records.iterrows():
                    cell_value = row[col]
                    if pd.isna(cell_value):
                        reason = "字段值为空"
                    elif any(keyword in str(cell_value).lower() for keyword in ['fail', '失败']):
                        reason = "标记为失败"
                    elif any(keyword in str(cell_value) for keyword in ['不通过', '不合格']):
                        reason = "标记为不通过"
                    elif any(keyword in str(cell_value) for keyword in ['不匹配', '不一致']):
                        reason = "标记为不匹配"
                    else:
                        reason = "其他不通过原因"
                    
                    detailed_analysis.append({
                        '比对字段': col,
                        '记录行号': idx + 2,
                        '字段值': cell_value,
                        '不通过原因': reason
                    })
    
    return results, total_fails, total_records, detailed_analysis

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
        comparison_cols = []
        for col in df.columns:
            col_str = str(col)
            if any(pattern in col_str for pattern in ['_比对', '比对结果', '对比结果', '比对']):
                comparison_cols.append(col)
        
        if not comparison_cols:
            print("❌ 未找到比对字段")
            print("可用列名:")
            for col in df.columns:
                print(f"  - {col}")
            input("\n按回车键退出...")
            return
        
        print(f"\n🎯 找到 {len(comparison_cols)} 个比对字段:")
        for col in comparison_cols:
            print(f"  - {col}")
        
        # 询问用户是否需要详细分析
        need_detailed = ask_for_detailed_analysis()
        
        # 初始化变量
        results = []
        total_fails = 0
        total_records = 0
        detailed_analysis = []  # 确保变量被初始化
        
        # 根据用户选择执行不同的分析
        if need_detailed:
            print("\n🎯 已选择【详细分析】模式")
            results, total_fails, total_records, detailed_analysis = detailed_analysis(df, comparison_cols)
            analysis_mode = "详细"
        else:
            print("\n🎯 已选择【基础分析】模式")
            results, total_fails, total_records, detailed_analysis = basic_analysis(df, comparison_cols)
            analysis_mode = "基础"
        
        # 创建统计汇总表
        summary_df = pd.DataFrame(results)
        
        # 汇总统计
        if total_records > 0:
            overall_fail_rate = (total_fails / total_records) * 100
            overall_pass_rate = 100 - overall_fail_rate
            
            print("\n" + "=" * 60)
            print("📈 汇总统计:")
            print(f"   总不通过记录数: {total_fails}")
            print(f"   总记录数: {total_records}")
            print(f"   平均不通过率: {overall_fail_rate:.2f}%")
            print(f"   平均通过率: {overall_pass_rate:.2f}%")
            if need_detailed and detailed_analysis:
                print(f"   详细不通过记录: {len(detailed_analysis)} 条")
            print("=" * 60)
        
        # 保存结果到新的Excel文件
        output_file = os.path.splitext(excel_file)[0] + f"_{analysis_mode}分析报告.xlsx"
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Sheet1: 统计汇总
            summary_df.to_excel(writer, sheet_name='统计汇总', index=False)
            
            # Sheet2: 详细原因分析（仅详细模式）
            if need_detailed and detailed_analysis:
                detailed_df = pd.DataFrame(detailed_analysis)
                detailed_df.to_excel(writer, sheet_name='详细原因', index=False)
            elif need_detailed:
                # 如果选择了详细分析但没有不通过记录，创建空表
                pd.DataFrame({'说明': ['没有不通过记录']}).to_excel(writer, sheet_name='详细原因', index=False)
            
            # Sheet3: 原因统计（仅详细模式）
            if need_detailed and detailed_analysis:
                reason_stats = []
                for col in comparison_cols:
                    field_reasons = [item for item in detailed_analysis if item['比对字段'] == col]
                    if field_reasons:
                        reason_counts = {}
                        for item in field_reasons:
                            reason = item['不通过原因']
                            reason_counts[reason] = reason_counts.get(reason, 0) + 1
                        
                        for reason, count in reason_counts.items():
                            reason_stats.append({
                                '比对字段': col,
                                '不通过原因': reason,
                                '出现次数': count,
                                '占比': f"{(count/len(field_reasons))*100:.1f}%"
                            })
                
                if reason_stats:
                    reason_stats_df = pd.DataFrame(reason_stats)
                    reason_stats_df.to_excel(writer, sheet_name='原因统计', index=False)
                else:
                    pd.DataFrame({'说明': ['没有不通过记录']}).to_excel(writer, sheet_name='原因统计', index=False)
            
            # Sheet4: 通过率排名
            pass_rate_summary = summary_df[['比对字段', '通过数量', '不通过数量', '通过率', '不通过率']].copy()
            pass_rate_summary['通过率数值'] = pass_rate_summary['通过率'].str.rstrip('%').astype(float)
            pass_rate_summary = pass_rate_summary.sort_values('通过率数值', ascending=True)
            pass_rate_summary.drop('通过率数值', axis=1, inplace=True)
            pass_rate_summary.to_excel(writer, sheet_name='通过率排名', index=False)
        
        print(f"\n💾 {analysis_mode}分析报告已保存到: {output_file}")
        print("包含以下工作表:")
        print("  - 统计汇总: 各比对字段的基础统计")
        print("  - 通过率排名: 字段通过率排序")
        if need_detailed:
            print("  - 详细原因: 每条不通过记录的具体原因")
            print("  - 原因统计: 各种不通过原因的统计")
        
        # 显示关键发现
        if not summary_df.empty:
            worst_field = summary_df.loc[summary_df['不通过数量'].idxmax()]
            best_field = summary_df.loc[summary_df['通过数量'].idxmax()]
            
            print(f"\n⚠️  关键发现:")
            print(f"   问题最多的字段: {worst_field['比对字段']} (不通过率: {worst_field['不通过率']})")
            print(f"   表现最佳的字段: {best_field['比对字段']} (通过率: {best_field['通过率']})")
            
            if need_detailed and detailed_analysis:
                # 统计最主要的不通过原因
                from collections import Counter
                all_reasons = [item['不通过原因'] for item in detailed_analysis]
                top_reasons = Counter(all_reasons).most_common(3)
                print(f"\n🔍 主要不通过原因:")
                for reason, count in top_reasons:
                    print(f"   - {reason}: {count}次")
        
        print(f"\n⏱️  分析模式: {analysis_mode}分析")
        print(f"📊 数据规模: {len(df)} 行 × {len(df.columns)} 列")
        print(f"🎯 分析字段: {len(comparison_cols)} 个比对字段")
        
    except Exception as e:
        print(f"❌ 错误: {str(e)}")
        import traceback
        traceback.print_exc()
    
    print("\n🎉 分析完成!")
    input("按回车键退出...")

if __name__ == "__main__":
    main()
