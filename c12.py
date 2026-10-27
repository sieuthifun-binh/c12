import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io

st.set_page_config(page_title="Hệ Thống Quản Lý & Tự Động Hóa Thu Nộp - Nợ Đơn Vị", layout="wide")

st.title("📊 HỆ THỐNG QUẢN LÝ THU NỘP & NỢ ĐƠN VỊ (QUẢN LÝ DỮ LIỆU ĐỘNG)")
st.markdown("---")

uploaded_file = st.file_uploader("Tải lên file Excel dữ liệu gốc (Theo doi tinh hinh thu nop don vi 2025.xls)", type=["xls", "xlsx"])

if uploaded_file is not None:
    # Đọc toàn bộ các sheet từ file Excel gốc vào session_state
    excel_file = pd.ExcelFile(uploaded_file)
    sheet_names = excel_file.sheet_names
    
    if 'dfs' not in st.session_state:
        st.session_state['dfs'] = {sheet: pd.read_excel(uploaded_file, sheet_name=sheet) for sheet in sheet_names}
    
    dfs = st.session_state['dfs']
    
    if 'Cập nhật' in dfs:
        df_capnhat = dfs['Cập nhật']
        
        # Sidebar điều hướng
        st.sidebar.header("🗂️ Chức năng hệ thống")
        menu = st.sidebar.radio(
            "Chọn tác vụ:", 
            [
                "📊 Dashboard & Trực quan", 
                "🎯 Xem Chi Tiết Theo Mã Đơn Vị", 
                "📋 Tra cứu danh sách tổng hợp", 
                "➕ Thêm / Cập nhật dòng dữ liệu"
            ]
        )
        
        # -------------------------------------------------------------
        # TÁC VỤ 1: DASHBOARD TỔNG QUAN
        # -------------------------------------------------------------
        if menu == "📊 Dashboard & Trực quan":
            st.subheader("Tổng quan tình hình thu nộp và nợ đọng toàn hệ thống")
            
            total_units = len(df_capnhat)
            total_labor = df_capnhat['sld'].sum() if 'sld' in df_capnhat else 0
            total_pst = df_capnhat['pst'].sum() if 'pst' in df_capnhat else 0
            total_debt = df_capnhat['tongThieu'].sum() if 'tongThieu' in df_capnhat else 0
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Tổng Số Đơn Vị", f"{total_units:,}")
            col2.metric("Tổng Lao Động", f"{total_labor:,.0f}")
            col3.metric("Tổng Phát Sinh Thu", f"{total_pst:,.0f} đ")
            col4.metric("Tổng Số Tiền Nợ", f"{total_debt:,.0f} đ")
            
            st.markdown("---")
            col_a, col_b = st.columns(2)
            
            with col_a:
                st.markdown("#### Top 10 Đơn Vị Nợ Cao Nhất")
                if 'tongThieu' in df_capnhat and 'tendvi' in df_capnhat:
                    top10 = df_capnhat.sort_values(by='tongThieu', ascending=False).head(10)
                    fig, ax = plt.subplots(figsize=(8, 6))
                    ax.barh(top10['tendvi'].astype(str).str[:30], top10['tongThieu'] / 1e6, color='#e74c3c')
                    ax.set_xlabel('Tổng nợ (Triệu VNĐ)')
                    ax.invert_yaxis()
                    st.pyplot(fig)
                    
            with col_b:
                st.markdown("#### Tổng Nợ Theo Khối Đơn Vị")
                if 'makhoi' in df_capnhat and 'tongThieu' in df_capnhat:
                    block_grp = df_capnhat.groupby('makhoi')['tongThieu'].sum().reset_index().sort_values(by='tongThieu', ascending=False).head(10)
                    fig2, ax2 = plt.subplots(figsize=(8, 6))
                    ax2.bar(block_grp['makhoi'].astype(str), block_grp['tongThieu'] / 1e6, color='#3498db')
                    ax2.set_ylabel('Tổng nợ (Tr VNĐ)')
                    ax2.set_xlabel('Khối đơn vị')
                    st.pyplot(fig2)

        # -------------------------------------------------------------
        # TÁC VỤ 2: XEM CHI TIẾT THEO TỪNG MÃ ĐƠN VỊ (MỚI BỔ SUNG)
        # -------------------------------------------------------------
        elif menu == "🎯 Xem Chi Tiết Theo Mã Đơn Vị":
            st.subheader("🎯 Tra cứu Hồ sơ Chi tiết Đơn vị")
            
            # Tạo danh sách Mã đơn vị kèm Tên đơn vị để tìm kiếm / chọn nhanh
            unit_options = df_capnhat['madvi'].astype(str) + " - " + df_capnhat['tendvi'].astype(str)
            selected_option = st.selectbox("👉 Chọn hoặc Gõ tìm kiếm Mã / Tên Đơn Vị:", unit_options)
            
            if selected_option:
                selected_madvi = selected_option.split(" - ")[0]
                unit_data = df_capnhat[df_capnhat['madvi'].astype(str) == selected_madvi].iloc[0]
                
                # Khung thông tin tổng quan
                st.markdown(f"### 🏢 {unit_data.get('tendvi', 'N/A')}")
                st.info(f"**Mã đơn vị:** {unit_data.get('madvi', 'N/A')} | **Khối:** {unit_data.get('makhoi', 'N/A')} | **Kỳ theo dõi:** {unit_data.get('thang', 'N/A')}")
                
                # Thẻ Chỉ số chính của Đơn vị
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Số lao động", f"{unit_data.get('sld', 0):,.0f} người")
                c2.metric("Phát sinh thu (PST)", f"{unit_data.get('pst', 0):,.0f} đ")
                c3.metric("Tổng thiếu / Nợ", f"{unit_data.get('tongThieu', 0):,.0f} đ")
                c4.metric("Tỷ lệ nợ BHXH", f"{unit_data.get('tyleno', 0) * 100:.1f}%")
                
                st.markdown("---")
                
                # Phân chia thông tin chi tiết thành các nhóm Tab
                tab_info, tab_financial, tab_all = st.tabs(["📌 Thông tin hành chính & Liên hệ", "💰 Chi tiết Thu Nộp & Nợ", "📄 Tất cả 230 Cột Dữ Liệu"])
                
                with tab_info:
                    col_i1, col_i2 = st.columns(2)
                    with col_i1:
                        st.write(f"**Địa chỉ:** {unit_data.get('diachi', 'Chưa có thông tin')}")
                        st.write(f"**Điện thoại liên hệ:** {unit_data.get('dienthoai', 'Chưa có')}")
                        st.write(f"**Người liên hệ:** {unit_data.get('nguoilh', 'Chưa có')}")
                    with col_i2:
                        st.write(f"**Chuyên quản phụ trách:** {unit_data.get('ma_quan', 'Chưa có')}")
                        st.write(f"**Email / Chuyên quản:** {unit_data.get('dsnv', 'Chưa có')}")
                        st.write(f"**Ghi chú đơn vị:** {unit_data.get('ghichu', 'Không có')}")
                        
                with tab_financial:
                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        st.markdown("##### 📥 Các khoản Phát sinh & Dư nợ")
                        st.write(f"- **Phát sinh BHXH:** {unit_data.get('_ps_bh', 0):,.0f} đ")
                        st.write(f"- **Phát sinh BHYT:** {unit_data.get('_ps_yt', 0):,.0f} đ")
                        st.write(f"- **Phát sinh BHTN:** {unit_data.get('_ps_tn', 0):,.0f} đ")
                        st.write(f"- **Số tiền đầu kỳ:** {unit_data.get('_tien_dk', 0):,.0f} đ")
                    with col_f2:
                        st.markdown("##### 📤 Số tiền Đã Nộp & Cuối kỳ")
                        st.write(f"- **Số tiền đã nộp (UNC):** {unit_data.get('_tien_unc', 0):,.0f} đ")
                        st.write(f"- **Số tiền cuối kỳ:** {unit_data.get('_tien_ck', 0):,.0f} đ")
                        st.write(f"- **Tỷ lệ nợ BHYT:** {unit_data.get('tylenoyt', 0) * 100:.1f}%")

                with tab_all:
                    st.markdown("##### 🔍 Toàn bộ thuộc tính dữ liệu của đơn vị này (Hơn 200 thuộc tính):")
                    df_single = pd.DataFrame([unit_data])
                    st.dataframe(df_single.T.rename(columns={0: "Giá trị"}), use_container_width=True)

        # -------------------------------------------------------------
        # TÁC VỤ 3: TRA CỨU DANH SÁCH TỔNG HỢP
        # -------------------------------------------------------------
        elif menu == "📋 Tra cứu danh sách tổng hợp":
            st.subheader("Tra cứu chi tiết trên tất cả các dòng dữ liệu")
            search_term = st.text_input("Tìm kiếm theo Tên đơn vị hoặc Mã đơn vị:")
            if search_term:
                filtered = df_capnhat[df_capnhat['tendvi'].astype(str).str.contains(search_term, case=False, na=False) | 
                                      df_capnhat['madvi'].astype(str).str.contains(search_term, case=False, na=False)]
            else:
                filtered = df_capnhat
                
            st.write(f"Hệ thống đang hiển thị **{len(filtered)}** dòng dữ liệu khớp điều kiện.")
            st.dataframe(filtered, use_container_width=True)

        # -------------------------------------------------------------
        # TÁC VỤ 4: THÊM / CẬP NHẬT DỮ LIỆU
        # -------------------------------------------------------------
        elif menu == "➕ Thêm / Cập nhật dòng dữ liệu":
            st.subheader("Thêm mới hoặc cập nhật thông tin đơn vị")
            
            with st.form("add_row_form"):
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    new_madvi = st.text_input("Mã đơn vị (madvi) *bắt buộc*")
                    new_tendvi = st.text_input("Tên đơn vị (tendvi)")
                    new_makhoi = st.text_input("Mã khối (makhoi)")
                    new_diachi = st.text_input("Địa chỉ (diachi)")
                with col_f2:
                    new_sld = st.number_input("Số lao động (sld)", min_value=0, value=1)
                    new_pst = st.number_input("Phát sinh thu (pst)", min_value=0.0, value=0.0)
                    new_tongThieu = st.number_input("Tổng nợ thiếu (tongThieu)", min_value=0.0, value=0.0)
                    new_thang = st.text_input("Kỳ tháng (thang, VD: 202609)", value="202609")
                
                submitted = st.form_submit_button("💾 Lưu dòng dữ liệu")
                
                if submitted:
                    if not new_madvi:
                        st.error("Vui lòng nhập Mã đơn vị!")
                    else:
                        if new_madvi in df_capnhat['madvi'].astype(str).values:
                            idx = df_capnhat[df_capnhat['madvi'].astype(str) == new_madvi].index[0]
                            df_capnhat.loc[idx, 'tendvi'] = new_tendvi if new_tendvi else df_capnhat.loc[idx, 'tendvi']
                            df_capnhat.loc[idx, 'makhoi'] = new_makhoi if new_makhoi else df_capnhat.loc[idx, 'makhoi']
                            df_capnhat.loc[idx, 'sld'] = new_sld
                            df_capnhat.loc[idx, 'pst'] = new_pst
                            df_capnhat.loc[idx, 'tongThieu'] = new_tongThieu
                            st.success(f"Đã cập nhật thông tin cho đơn vị mã: {new_madvi}")
                        else:
                            new_row = {col: pd.NA for col in df_capnhat.columns}
                            new_row['madvi'] = new_madvi
                            new_row['tendvi'] = new_tendvi
                            new_row['makhoi'] = new_makhoi
                            new_row['diachi'] = new_diachi
                            new_row['sld'] = new_sld
                            new_row['pst'] = new_pst
                            new_row['tongThieu'] = new_tongThieu
                            new_row['thang'] = new_thang
                            
                            df_capnhat = pd.concat([df_capnhat, pd.DataFrame([new_row])], ignore_index=True)
                            st.session_state['dfs']['Cập nhật'] = df_capnhat
                            st.success(f"✅ Đã thêm thành công đơn vị mới! Tổng dòng: {len(df_capnhat):,}")

        st.markdown("---")
        # Nút xuất File Excel hoàn chỉnh
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for s_name, s_df in st.session_state['dfs'].items():
                s_df.to_excel(writer, sheet_name=s_name, index=False)
        processed_data = output.getvalue()

        st.download_button(
            label="📥 Tải xuống File Excel hoàn chỉnh (Full dữ liệu)",
            data=processed_data,
            file_name='Theo_doi_tinh_hinh_thu_nop_don_vi_2025_full_updated.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
else:
    st.info("👆 Vui lòng tải lên file Excel gốc của anh/chị ở khung trên để khởi động ứng dụng.")
