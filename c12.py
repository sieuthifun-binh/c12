import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io
import zipfile
from datetime import datetime
from docx import Document
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

st.set_page_config(page_title="Hệ Thống Quản Lý & Tác Nghiệp Thu Nộp Thông Minh", layout="wide")

# Cấu hình font tiếng Việt cho biểu đồ
plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['axes.unicode_minus'] = False

st.title("🚀 HỆ THỐNG QUẢN LÝ THU NỘP & TÁC NGHIỆP THÔNG MINH (ENTERPRISE EDITION)")
st.markdown("---")

# Khởi tạo bộ nhớ tạm cho CRM Activity Log nếu chưa có
if 'crm_logs' not in st.session_state:
    st.session_state['crm_logs'] = pd.DataFrame(columns=['Thoi_Gian', 'Ma_Dvi', 'Ten_Dvi', 'Hanh_Dong', 'Noi_Dung', 'Can_Bo'])

# 1. KHU VỰC NẠP DỮ LIỆU
st.sidebar.header("📁 1. Nạp Dữ Liệu Hệ Thống")
uploaded_file = st.sidebar.file_uploader("Tải lên file Excel tháng hiện tại (.xls, .xlsx)", type=["xls", "xlsx"])
uploaded_file_prev = st.sidebar.file_uploader("Tải lên file Excel tháng trước (Tùy chọn - so sánh đa kỳ)", type=["xls", "xlsx"])

if uploaded_file is not None:
    try:
        # Đọc file hiện tại
        xls = pd.ExcelFile(uploaded_file)
        target_sheet = 'Cập nhật' if 'Cập nhật' in xls.sheet_names else xls.sheet_names[0]
        df_master = pd.read_excel(uploaded_file, sheet_name=target_sheet)
        df_master.columns = df_master.columns.astype(str).str.strip()
        
        # Đọc file tháng trước nếu có
        df_prev = None
        if uploaded_file_prev is not None:
            xls_prev = pd.ExcelFile(uploaded_file_prev)
            target_sheet_prev = 'Cập nhật' if 'Cập nhật' in xls_prev.sheet_names else xls_prev.sheet_names[0]
            df_prev = pd.read_excel(uploaded_file_prev, sheet_name=target_sheet_prev)
            df_prev.columns = df_prev.columns.astype(str).str.strip()

        # THUẬT TOÁN CHẤM ĐIỂM RỦI RO THÔNG MINH KẾT HỢP CỘT `tyleno` (Số tháng nợ)
        def ai_risk_scoring(row):
            tyleno = row.get('tyleno', 0)
            debt = row.get('tongThieu', 0)
            
            if pd.isna(tyleno):
                tyleno = 0
            
            # Quy chuẩn nghiệp vụ theo số tháng nợ (tyleno)
            if tyleno >= 3.0 or (pd.isna(tyleno) and debt > 100000000):
                return "🔴 Nguy cơ nợ kéo dài (>= 3 tháng)"
            elif 1.0 <= tyleno < 3.0:
                return "🟡 Nợ đọng (Từ 1 đến < 3 tháng)"
            elif 0.0 < tyleno < 1.0:
                return "🟠 Chậm đóng (< 1 tháng)"
            else:
                return "🟢 An toàn / Không nợ"
        
        df_master['Muc_Do_Rui_Ro'] = df_master.apply(ai_risk_scoring, axis=1)

        # Menu điều hướng các tính năng
        menu = st.sidebar.selectbox(
            "⚙️ Chọn Chức Nặng Tác Nghiệp:", 
            [
                "📊 1. Dashboard & Cảnh Báo Rủi Ro (Theo Số Tháng Nợ)", 
                "🤖 2. Trợ Lý AI Tra Cứu Ngôn Ngữ Tự Nhiên", 
                "👤 3. Quản Lý Chuyên Quản & CRM Tác Nghiệp", 
                "📧 4. Gửi Email Tự Động Qua Gmail API/SMTP", 
                "📈 5. Phân Tích Biến Động Đa Kỳ (MoM)", 
                "📝 6. Soạn & Tải Công Văn Word Hàng Loạt", 
                "📋 7. Kho Dữ Liệu Gốc & Xuất Báo Cáo"
            ]
        )
        
        # -------------------------------------------------------------
        # TÍNH NĂNG 1: DASHBOARD & CẢNH BÁO RỦI RO (TÍCH HỢP _TIEN_UNC & TYLENO)
        # -------------------------------------------------------------
        if menu == "📊 1. Dashboard & Cảnh Báo Rủi Ro (Theo Số Tháng Nợ)":
            st.subheader("📊 Tổng Quan & Phân Tích Dòng Tiền, Số Tháng Nợ (`tyleno`) & Đã Nộp (`_tien_unc`)")
            
            total_units = len(df_master)
            total_labor = df_master['sld'].sum() if 'sld' in df_master else 0
            total_pst = df_master['pst'].sum() if 'pst' in df_master else 0
            total_unc = df_master['_tien_unc'].sum() if '_tien_unc' in df_master else 0
            total_debt = df_master['tongThieu'].sum() if 'tongThieu' in df_master else 0
            
            # Hiển thị các chỉ số KPI Tổng quan (Bổ sung thêm Số đã thu _tien_unc)
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Tổng Đơn Vị", f"{total_units:,}")
            c2.metric("Tổng Lao Động", f"{total_labor:,.0f}")
            c3.metric("Tổng Phát Sinh", f"{total_pst:,.0f} đ")
            c4.metric("💰 Đã Thu (_tien_unc)", f"{total_unc:,.0f} đ")
            c5.metric("⚠️ Tổng Tiền Nợ", f"{total_debt:,.0f} đ")
            
            st.markdown("---")
            
            # Phân tách dữ liệu theo các nhóm rủi ro chuẩn số tháng nợ
            df_long_debt = df_master[df_master['Muc_Do_Rui_Ro'] == "🔴 Nguy cơ nợ kéo dài (>= 3 tháng)"]
            df_medium_debt = df_master[df_master['Muc_Do_Rui_Ro'] == "🟡 Nợ đọng (Từ 1 đến < 3 tháng)"]
            df_short_debt = df_master[df_master['Muc_Do_Rui_Ro'] == "🟠 Chậm đóng (< 1 tháng)"]
            df_safe = df_master[df_master['Muc_Do_Rui_Ro'] == "🟢 An toàn / Không nợ"]

            # Hiển thị thẻ Card thống kê số lượng, tổng tiền nợ và tổng số đã thu từng nhóm
            rc1, rc2, rc3, rc4 = st.columns(4)
            rc1.metric("🔴 Nợ kéo dài (>=3T)", f"{len(df_long_debt):,} đơn vị", f"Nợ: {df_long_debt['tongThieu'].sum():,.0f} đ")
            rc2.metric("🟡 Nợ đọng (1-3T)", f"{len(df_medium_debt):,} đơn vị", f"Nợ: {df_medium_debt['tongThieu'].sum():,.0f} đ")
            rc3.metric("🟠 Chậm đóng (<1T)", f"{len(df_short_debt):,} đơn vị", f"Nợ: {df_short_debt['tongThieu'].sum():,.0f} đ")
            rc4.metric("🟢 An toàn", f"{len(df_safe):,} đơn vị", f"Đã nộp: {df_safe['_tien_unc'].sum():,.0f} đ")
            
            st.markdown("---")
            st.markdown("### 📂 Bấm chọn các Tab bên dưới để xem chi tiết danh sách đơn vị theo số tháng nợ (bao gồm cột số đã nộp `_tien_unc`):")
            
            # Tạo các Tab trực quan cho từng nhóm
            tab_long, tab_med, tab_short, tab_safe_t = st.tabs([
                f"🔴 Nợ kéo dài >=3 tháng ({len(df_long_debt)})", 
                f"🟡 Nợ đọng 1-3 tháng ({len(df_medium_debt)})", 
                f"🟠 Chậm đóng <1 tháng ({len(df_short_debt)})",
                f"🟢 An toàn ({len(df_safe)})"
            ])
            
            # Đảm bảo hiển thị cột _tien_unc cùng các cột trọng tâm
            cols_show = [c for c in ['madvi', 'tendvi', 'makhoi', 'sld', 'pst', '_tien_unc', 'tongThieu', 'tyleno', 'diachi'] if c in df_master.columns]

            with tab_long:
                st.markdown("##### 🚨 Danh sách đơn vị nợ kéo dài từ **3 tháng trở lên** (`tyleno >= 3.0`) - Cần đôn đốc pháp lý/thanh tra gấp")
                if len(df_long_debt) > 0:
                    st.dataframe(df_long_debt[cols_show], use_container_width=True)
                    out_l = io.BytesIO()
                    with pd.ExcelWriter(out_l, engine='openpyxl') as writer:
                        df_long_debt.to_excel(writer, sheet_name='No_keo_dai', index=False)
                    st.download_button("📥 Tải danh sách Nợ kéo dài (Excel)", out_l.getvalue(), "Danh_sach_no_keo_dai.xlsx", key="dl_l")
                else:
                    st.info("Không có đơn vị nào trong nhóm này.")

            with tab_med:
                st.markdown("##### ⚠️ Danh sách đơn vị nợ đọng từ **1 đến dưới 3 tháng** (`1.0 <= tyleno < 3.0`)")
                if len(df_medium_debt) > 0:
                    st.dataframe(df_medium_debt[cols_show], use_container_width=True)
                    out_m = io.BytesIO()
                    with pd.ExcelWriter(out_m, engine='openpyxl') as writer:
                        df_medium_debt.to_excel(writer, sheet_name='No_dong', index=False)
                    st.download_button("📥 Tải danh sách Nợ đọng (Excel)", out_m.getvalue(), "Danh_sach_no_dong.xlsx", key="dl_m")
                else:
                    st.info("Không có đơn vị nào trong nhóm này.")

            with tab_short:
                st.markdown("##### 🔸 Danh sách đơn vị chậm đóng dưới **1 tháng** (`tyleno < 1.0`)")
                if len(df_short_debt) > 0:
                    st.dataframe(df_short_debt[cols_show], use_container_width=True)
                    out_s = io.BytesIO()
                    with pd.ExcelWriter(out_s, engine='openpyxl') as writer:
                        df_short_debt.to_excel(writer, sheet_name='Cham_dong', index=False)
                    st.download_button("📥 Tải danh sách Chậm đóng (Excel)", out_s.getvalue(), "Danh_sach_cham_dong.xlsx", key="dl_s")
                else:
                    st.info("Không có đơn vị nào trong nhóm này.")

            with tab_safe_t:
                st.markdown("##### ✅ Danh sách đơn vị hoàn thành nghĩa vụ (`tyleno = 0`)")
                if len(df_safe) > 0:
                    st.dataframe(df_safe[cols_show], use_container_width=True)
                    out_saf = io.BytesIO()
                    with pd.ExcelWriter(out_saf, engine='openpyxl') as writer:
                        df_safe.to_excel(writer, sheet_name='An_toan', index=False)
                    st.download_button("📥 Tải danh sách An toàn (Excel)", out_saf.getvalue(), "Danh_sach_an_toan.xlsx", key="dl_saf")
                else:
                    st.info("Không có đơn vị nào trong nhóm này.")

            st.markdown("---")
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("#### 🔴 Top 10 Đơn Vị Có Số Tháng Nợ (`tyleno`) Cao Nhất")
                top10_tyleno = df_master.sort_values(by='tyleno', ascending=False).head(10)
                fig, ax = plt.subplots(figsize=(8, 6))
                ax.barh(top10_tyleno['tendvi'].astype(str).str[:30], top10_tyleno['tyleno'], color='#e74c3c')
                ax.set_xlabel('Số tháng nợ quy đổi (Tỷ lệ nợ)', fontweight='bold')
                ax.invert_yaxis()
                st.pyplot(fig)
            with col_b:
                st.markdown("#### 🔵 Tỷ Trọng Phân Bổ Các Nhóm Số Tháng Nợ")
                risk_counts = df_master['Muc_Do_Rui_Ro'].value_counts()
                fig2, ax2 = plt.subplots(figsize=(7, 7))
                ax2.pie(risk_counts.values, labels=risk_counts.index, autopct='%1.1f%%', colors=['#2ecc71', '#f39c12', '#e67e22', '#e74c3c'], startangle=140)
                st.pyplot(fig2)

        # -------------------------------------------------------------
        # TÍNH NĂNG 2: TRỢ LÝ AI TRA CỨU NGÔN NGỮ TỰ NHIÊN
        # -------------------------------------------------------------
        elif menu == "🤖 2. Trợ Lý AI Tra Cứu Ngôn Ngữ Tự Nhiên":
            st.subheader("🤖 Trợ Lý Ảo Truy Vấn Dữ Liệu Bằng Ngôn Ngữ Tự Nhiên")
            st.info("💡 **Ví dụ câu lệnh:** Nhập `nợ > 3` (tìm đơn vị nợ từ 3 tháng trở lên) hoặc `đã nộp > 1000000`.")
            
            query = st.text_input("💬 Nhập yêu cầu tìm kiếm của bạn:")
            if query:
                try:
                    query_lower = query.lower()
                    if "nợ >=" in query_lower or "nợ >" in query_lower:
                        val = float(query_lower.replace("nợ >=", "").replace("nợ >", "").strip())
                        result_df = df_master[df_master['tyleno'] >= val]
                    elif "đã nộp >" in query_lower:
                        val = float(query_lower.split("đã nộp >")[1].strip())
                        result_df = df_master[df_master['_tien_unc'] > val]
                    elif "lao động >" in query_lower:
                        val = float(query_lower.split("lao động >")[1].strip())
                        result_df = df_master[df_master['sld'] > val]
                    else:
                        result_df = df_master[df_master['tendvi'].astype(str).str.contains(query, case=False, na=False)]
                    
                    st.success(f"Tìm thấy **{len(result_df):,}** đơn vị phù hợp:")
                    st.dataframe(result_df[['madvi', 'tendvi', 'sld', '_tien_unc', 'tongThieu', 'tyleno']], use_container_width=True)
                except Exception as ex:
                    st.warning(f"Không hiểu câu lệnh. Vui lòng thử lại theo cú pháp mẫu. (Lỗi: {ex})")

        # -------------------------------------------------------------
        # T
