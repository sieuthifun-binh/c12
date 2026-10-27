import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io

st.set_page_config(page_title="Hệ Thống Phân Tích Tự Động Thu Nộp & Nợ Đơn Vị", layout="wide")

# Cấu hình giao diện tiếng Việt & font chữ
plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['axes.unicode_minus'] = False

st.title("📊 HỆ THỐNG PHÂN TÍCH & QUẢN LÝ THU NỘP - NỢ ĐƠN VỊ (TỰ ĐỘNG HÓA)")
st.markdown("---")

# 1. KHU VỰC NẠP DỮ LIỆU GỐC
uploaded_file = st.file_uploader("📤 Tải lên file Excel dữ liệu gốc hàng tháng (ví dụ: .xls hoặc .xlsx)", type=["xls", "xlsx"])

if uploaded_file is not None:
    try:
        # Đọc toàn bộ các sheet có trong file Excel trực tiếp (không dùng cache lỗi file handle)
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        # Đọc sheet 'Cập nhật' làm master data, nếu không có lấy sheet đầu tiên
        target_sheet = 'Cập nhật' if 'Cập nhật' in sheet_names else sheet_names[0]
        df_master = pd.read_excel(uploaded_file, sheet_name=target_sheet)
        
        # Chuẩn hóa tên cột (loại bỏ khoảng trắng thừa nếu có)
        df_master.columns = df_master.columns.astype(str).str.strip()
        
        st.sidebar.success(f"Đọc file thành công! Các sheet phát hiện: {', '.join(sheet_names)}")
        
        # Sidebar điều hướng chức năng
        menu = st.sidebar.radio(
            "⚙️ Chọn Chức Năng Phân Tích:", 
            [
                "📈 Tổng Quan Dashboard & Biểu Đồ", 
                "🎯 Tra Cứu Chi Tiết Từng Đơn Vị", 
                "📋 Xem Bảng Dữ Liệu Gốc & Xuất File"
            ]
        )
        
        # -------------------------------------------------------------
        # TÁC VỤ 1: DASHBOARD & BIỂU ĐỒ TRỰC QUAN TỰ ĐỘNG
        # -------------------------------------------------------------
        if menu == "📈 Tổng Quan Dashboard & Biểu Đồ":
            st.subheader("📈 Báo Cáo Phân Tích Tổng Quan Tự Động")
            
            total_units = len(df_master)
            total_labor = df_master['sld'].sum() if 'sld' in df_master else 0
            total_pst = df_master['pst'].sum() if 'pst' in df_master else 0
            total_debt = df_master['tongThieu'].sum() if 'tongThieu' in df_master else 0
            
            # Hiển thị thẻ KPI
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Tổng Số Đơn Vị", f"{total_units:,}")
            c2.metric("Tổng Số Lao Động", f"{total_labor:,.0f}")
            c3.metric("Tổng Phát Sinh Thu", f"{total_pst:,.0f} đ")
            c4.metric("Tổng Số Tiền Nợ", f"{total_debt:,.0f} đ")
            
            st.markdown("---")
            
            # Vẽ biểu đồ phân tích trực quan
            col_a, col_b = st.columns(2)
            
            with col_a:
                st.markdown("#### 🔴 Top 10 Đơn Vị Nợ Cao Nhất")
                if 'tongThieu' in df_master and 'tendvi' in df_master:
                    top10 = df_master.sort_values(by='tongThieu', ascending=False).head(10)
                    fig, ax = plt.subplots(figsize=(8, 6))
                    ax.barh(top10['tendvi'].astype(str).str[:32], top10['tongThieu'] / 1e6, color='#e74c3c')
                    ax.set_xlabel('Tổng nợ (Triệu VNĐ)', fontweight='bold')
                    ax.set_title('Top 10 Đơn vị cần đôn đốc thu nợ', fontsize=12, fontweight='bold', color='#c0392b')
                    ax.invert_yaxis()
                    st.pyplot(fig)
                else:
                    st.warning("File không đủ cột 'tongThieu' hoặc 'tendvi' để vẽ biểu đồ nợ.")
                    
            with col_b:
                st.markdown("#### 🔵 Tổng Nợ Theo Khối Đơn Vị")
                if 'makhoi' in df_master and 'tongThieu' in df_master:
                    block_grp = df_master.groupby('makhoi')['tongThieu'].sum().reset_index().sort_values(by='tongThieu', ascending=False).head(10)
                    fig2, ax2 = plt.subplots(figsize=(8, 6))
                    ax2.bar(block_grp['makhoi'].astype(str), block_grp['tongThieu'] / 1e6, color='#3498db')
                    ax2.set_ylabel('Tổng nợ (Tr VNĐ)', fontweight='bold')
                    ax2.set_xlabel('Khối đơn vị', fontweight='bold')
                    ax2.set_title('Phân bổ nợ theo Khối', fontsize=12, fontweight='bold', color='#2980b9')
                    st.pyplot(fig2)
                else:
                    st.warning("File không đủ cột 'makhoi' hoặc 'tongThieu'.")

        # -------------------------------------------------------------
        # TÁC VỤ 2: TRA CỨU HỒ SƠ CHI TIẾT TỪNG ĐƠN VỊ
        # -------------------------------------------------------------
        elif menu == "🎯 Tra Cứu Chi Tiết Từng Đơn Vị":
            st.subheader("🎯 Tra Cứu Hồ Sơ & Chỉ Số Chi Tiết Đơn Vị")
            
            if 'madvi' in df_master and 'tendvi' in df_master:
                unit_options = df_master['madvi'].astype(str) + " - " + df_master['tendvi'].astype(str)
                selected_option = st.selectbox("🔍 Gõ hoặc chọn Mã / Tên Đơn Vị cần phân tích:", unit_options)
                
                if selected_option:
                    selected_madvi = selected_option.split(" - ")[0]
                    unit_row = df_master[df_master['madvi'].astype(str) == selected_madvi].iloc[0]
                    
                    st.markdown(f"### 🏢 {unit_row.get('tendvi', 'N/A')}")
                    st.info(f"**Mã đơn vị:** {unit_row.get('madvi', 'N/A')} | **Khối:** {unit_row.get('makhoi', 'N/A')} | **Kỳ tháng:** {unit_row.get('thang', 'N/A')}")
                    
                    sc1, sc2, sc3, sc4 = st.columns(4)
                    sc1.metric("Số lao động", f"{unit_row.get('sld', 0):,.0f}")
                    sc2.metric("Phát sinh thu", f"{unit_row.get('pst', 0):,.0f} đ")
                    sc3.metric("Tổng nợ", f"{unit_row.get('tongThieu', 0):,.0f} đ")
                    sc4.metric("Tỷ lệ nợ", f"{unit_row.get('tyleno', 0) * 100:.1f}%")
                    
                    st.markdown("---")
                    st.markdown("#### 📋 Thông tin chi tiết toàn bộ các cột dữ liệu của đơn vị:")
                    df_single_unit = pd.DataFrame([unit_row]).T.reset_index()
                    df_single_unit.columns = ["Tên thuộc tính (Cột)", "Giá trị dữ liệu"]
                    st.dataframe(df_single_unit, use_container_width=True)
            else:
                st.error("Không tìm thấy cột 'madvi' hoặc 'tendvi' trong file dữ liệu.")

        # -------------------------------------------------------------
        # TÁC VỤ 3: XEM BẢNG GỐC VÀ XUẤT FILE
        # -------------------------------------------------------------
        elif menu == "📋 Xem Bảng Dữ Liệu Gốc & Xuất File":
            st.subheader("📋 Kho Dữ Liệu Gốc & Bộ Lọc Nâng Cao")
            
            search_key = st.text_input("🔎 Tìm kiếm nhanh theo Tên hoặc Mã đơn vị trong bảng:")
            if search_key:
                filtered_df = df_master[df_master['tendvi'].astype(str).str.contains(search_key, case=False, na=False) | 
                                        df_master['madvi'].astype(str).str.contains(search_key, case=False, na=False)]
            else:
                filtered_df = df_master
                
            st.write(f"Đang hiển thị **{len(filtered_df):,}** bản ghi khớp với dữ liệu nạp lên.")
            st.dataframe(filtered_df, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, sheet_name='Ket_qua_phan_tich', index=False)
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Tải xuống Báo Cáo Phân Tích (Excel)",
                data=processed_data,
                file_name='Bao_cao_phan_tich_thu_nop.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

    except Exception as e:
        st.error(f"⚠️ Đã có lỗi xảy ra khi đọc file dữ liệu: {e}")
        st.info("Vui lòng đảm bảo file Excel tải lên đúng định dạng tiêu chuẩn.")
else:
    st.info("👆 Vui lòng bấm nút **Browse files** ở trên để tải file Excel dữ liệu thô lên. Hệ thống sẽ tự động phân tích và trực quan hóa ngay lập tức!")
