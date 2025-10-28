# -*- coding: utf-8 -*-
"""
Spyder 编辑器

extract top dish based on 推荐数
"""

import streamlit as st
import pandas as pd
import io


def read_preview_excel(uploaded_file, rows=5):
    """
    读取Excel文件的前几行，仅获取列名及前几行数据供用户选择。
    """
    df = pd.read_excel(uploaded_file, header=0, nrows=rows)
    return df


def read_full_excel(uploaded_file, column_names, footer_keywords=['Total', '应用的筛选器', '总计', '合计', '汇总']):
    """
    读取完整的Excel文件并处理底部信息
    """
    df = pd.read_excel(uploaded_file, header=0)
    
    # 查找数据结束位置（找到包含"Total"或"应用的筛选器"的行）
    end_row = None
    for i, row in df.iterrows():
        # 检查当前行是否包含任何底部关键词
        row_contains_footer = False
        for keyword in footer_keywords:
            if row.astype(str).str.contains(keyword, case=False, na=False).any():
                end_row = i
                row_contains_footer = True
                st.info(f"检测到底部信息行（包含 '{keyword}'），将在第 {i} 行处截断")
                break
        if row_contains_footer:
            break
    
    if end_row is None:
        st.info("未检测到底部信息行，使用全部数据")
        return df[column_names]
    else:
        # 只读取到底部信息之前的行
        st.info(f"已自动去除底部 {len(df) - end_row} 行信息")
        return df[column_names].iloc[:end_row]


def extract_top(df, keyword_column, index_column, extract_keyword_list, top_number):
    """
    提取TOP菜品
    """
    df_keyword = df[df[keyword_column].isin(extract_keyword_list)]
    
    if df_keyword.empty:
        return df_keyword
    
    # 排序
    df_keyword = df_keyword.sort_values([keyword_column, index_column], ascending=[True, False]).copy().reset_index(drop=True)
    df_keyword['rank_by_keyword'] = df_keyword.groupby(keyword_column).cumcount() + 1
    
    top_df_keyword = df_keyword[df_keyword['rank_by_keyword'] <= top_number]
    
    return top_df_keyword


def main():
    st.title("📊 菜品数据分析工具")
    st.write("上传Excel文件，提取指定风味/原料等关键词的TOP菜品")
    
    # 文件上传
    uploaded_file = st.file_uploader(
        "选择Excel文件，上传菜品数据。请看右侧的注意事项（鼠标移到？处）。", 
        type=["xlsx", "xls"],
        help="请确保上传的文件是Excel文件，菜品名称的列为dish, 品牌为brand，并且只有一个季度的数据。"
    )
    
    if uploaded_file is not None:
        try:
            # 读取前5行数据，仅供用户选择列名
            with st.spinner("正在读取Excel文件前5行..."):
                df_preview = read_preview_excel(uploaded_file, rows=5)
            
            st.success(f"文件读取成功！数据形状: {df_preview.shape[0]} 行 × {df_preview.shape[1]} 列")
            
            # 显示数据预览
            st.subheader("数据预览（前5行）")
            st.dataframe(df_preview)
            
            # 显示所有列名，供用户选择
            st.subheader("选择列")
            col1, col2 = st.columns(2)
            
            with col1:
                keyword_column = st.selectbox(
                    "选择关键词列",
                    options=list(df_preview.columns),
                    help="选择包含菜品关键词的列（如：品类、类型等）"
                )
            
            with col2:
                index_column = st.selectbox(
                    "选择推荐数列",
                    options=list(df_preview.columns),
                    help="选择用于排序的数值列（如：推荐数等）"
                )
            
            # 输入TOP数量
            top_number = st.number_input(
                "提取TOP菜品数量",
                min_value=1,
                max_value=1000,
                value=10,
                step=1,
                help="请输入正整数，表示每个关键词提取的TOP数量"
            )
            
            # 输入关键词
            st.subheader("提取关键词")
            st.write("请输入要提取的关键词，每个关键词一行：")
            
            keywords_input = st.text_area(
                "关键词列表",
                value="奶茶\n红糖\n茉莉\n咖啡\n鲜奶",
                height=150,
                help="每行输入一个关键词，将根据这些关键词提取TOP菜品"
            )
            
            # 处理关键词输入
            extract_keyword_list = [k.strip() for k in keywords_input.split('\n') if k.strip()]
            
            st.write("检测到的关键词:", extract_keyword_list)
            
            # 执行提取
            if st.button("提取TOP菜品", type="primary"):
                if not extract_keyword_list:
                    st.error("请至少输入一个关键词！")
                    return
                
                if keyword_column == index_column:
                    st.error("关键词列和推荐数列不能相同！")
                    return
                
                # 读取完整Excel数据
                with st.spinner("正在读取完整数据..."):
                    try:
                        full_df = read_full_excel(uploaded_file, column_names=[keyword_column, index_column, 'dish', 'brand'])
                        
                        result_df = extract_top(
                            df=full_df,
                            keyword_column=keyword_column,
                            index_column=index_column,
                            extract_keyword_list=extract_keyword_list,
                            top_number=top_number
                        )
                        
                        if result_df.empty:
                            st.warning("没有找到匹配关键词的数据！")
                        else:
                            st.success(f"提取完成！共找到 {len(result_df)} 条记录")
                            
                            # 显示结果
                            st.subheader("提取结果")
                            st.dataframe(result_df)
                            
                            # 统计信息
                            st.subheader("统计信息")
                            keyword_counts = result_df[keyword_column].value_counts()
                            st.write("各关键词提取数量:")
                            st.dataframe(keyword_counts)
                            
                            # 下载结果
                            st.subheader("下载结果")
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                result_df.to_excel(writer, sheet_name='TOP菜品', index=False)
                                keyword_counts_df = keyword_counts.reset_index()
                                keyword_counts_df.columns = ['关键词', '数量']
                                keyword_counts_df.to_excel(writer, sheet_name='统计', index=False)
                            
                            output.seek(0)
                            st.download_button(
                                label="下载Excel结果",
                                data=output,
                                file_name="TOP菜品_结果.xlsx",
                                mime="application/vnd.ms-excel"
                            )
                            
                    except Exception as e:
                        st.error(f"提取过程中出错: {e}")
                        
        except Exception as e:
            st.error(f"读取文件时出错: {e}")

if __name__ == "__main__":
    main()
#streamlit.io.cloud

