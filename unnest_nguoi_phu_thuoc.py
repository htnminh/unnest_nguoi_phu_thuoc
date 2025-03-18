import streamlit as st
import pandas as pd
# import pprint
import io


def df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()


def main():
    st.set_page_config(layout="wide")
    col1, col2 = st.columns(2)

    col1.markdown('### Tool chuyển đổi người phụ thuộc thành nhiều hàng')

    uploaded_file = col1.file_uploader('Chọn file excel (.xlsx, .xls)', type=['xlsx', 'xls'])

    if uploaded_file is not None:
        col1.write('Đang đọc dữ liệu...')

        try:
            df = pd.read_excel(uploaded_file)
        except Exception as e:
            col1.error(f'Lỗi đọc file: {e}')

        else:
            col1.info(f'Dữ liệu tải lên có **{df.shape[0]}** hàng và **{df.shape[1]}** cột. Dữ liệu 5 hàng đầu:')
            col1.dataframe(df.head(5))

            dependent_columns = [col for col in df.columns if 'của tất cả người phụ thuộc' in col]
            if dependent_columns:
                col2.markdown(f'Đang xử lý **{len(dependent_columns)}** cột chứa thông tin người phụ thuộc...')

                new_df = pd.DataFrame(columns=df.columns)

                for _, row in df.iterrows():
                    row_need_split = True
                    split_col_tups = []  # a list of form (col_index, split_col_list)

                    for col_index, col in enumerate(df.columns):
                        if col in dependent_columns:
                            split_col_list = list(map(str.strip, str(row[col]).split('|')))
                            split_col_tups.append((col_index, split_col_list))

                            if len(split_col_list) <= 1:
                                new_df = pd.concat([
                                    new_df, pd.DataFrame([row.copy()])
                                    ], ignore_index=True)
                                row_need_split = False
                                break
                    
                    # pprint.pprint(split_col_tups)
                    if row_need_split:
                        temp_df = pd.DataFrame([row.copy()] * len(split_col_list), )
                        # print(temp_df)
                        for _col_index, _split_col_list in split_col_tups:
                            # print(_col_index, _split_col_list)
                            for i, _split_col in enumerate(_split_col_list):
                                # print(i, _split_col)
                                temp_df.iloc[i, _col_index] = _split_col
                        new_df = pd.concat([new_df, temp_df], ignore_index=True)

            new_df = new_df.rename(columns=lambda x: str(x).replace('của tất cả người phụ thuộc', 'người phụ thuộc'))

            col2.success(f'Chuyển đổi thành công. Dữ liệu mới có **{new_df.shape[0]}** hàng và **{new_df.shape[1]}** cột.')
            col2.download_button(
                label="\u2193 Tải xuống file kết quả (.xlsx)",
                data=df_to_excel(new_df),
                file_name=f"{uploaded_file.name}_unnested.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            col2.write('Dữ liệu sau chuyển đổi (10 hàng đầu):')
            col2.dataframe(new_df.head(10))

            

if __name__ == '__main__':
    main()
