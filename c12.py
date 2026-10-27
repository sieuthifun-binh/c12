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
        # TÍNH NĂNG 3: QUẢN LÝ CHUYÊN QUẢN & CRM TÁC NGHIỆP
        # -------------------------------------------------------------
        elif menu == "👤 3. Quản Lý Chuyên Quản & CRM Tác Nghiệp":
            st.subheader("👤 Nhật Ký CRM & Theo Dõi Tác Nghiệp Chuyên Quản")
            
            quan_col = 'dsnv' if 'dsnv' in df_master.columns else ('ma_quan' if 'ma_quan' in df_master.columns else None)
            
            if quan_col:
                selected_quan = st.selectbox("👉 Chọn Cán bộ Chuyên quản:", df_master[quan_col].dropna().astype(str).unique())
                df_quan = df_master[df_master[quan_col].astype(str) == selected_quan]
                
                st.markdown(f"#### Danh sách đơn vị do **{selected_quan}** phụ trách ({len(df_quan)} đơn vị)")
                st.dataframe(df_quan[['madvi', 'tendvi', 'sld', '_tien_unc', 'tongThieu', 'tyleno']], use_container_width=True)
                
                st.markdown("---")
                st.markdown("### 📝 Sổ Tay Ghi Nhận Lịch Sử Tác Nghiệp (CRM Log)")
                with st.form("crm_form"):
                    c_madvi = st.selectbox("Chọn mã đơn vị làm việc:", df_quan['madvi'].astype(str).unique())
                    c_action = st.selectbox("Hình thức tác nghiệp:", ["Gọi điện nhắc nợ", "Gửi công văn", "Gặp trực tiếp doanh nghiệp", "Hứa hẹn lịch nộp"])
                    c_note = st.text_area("Nội dung trao đổi / Ghi chú chi tiết:")
                    c_staff = st.text_input("Tên cán bộ thực hiện:", value=selected_quan)
                    
                    submitted_crm = st.form_submit_button("💾 Lưu Nhật Ký Tác Nghiệp")
                    if submitted_crm:
                        unit_name = df_quan[df_quan['madvi'].astype(str) == c_madvi]['tendvi'].values[0]
                        new_log = pd.DataFrame([{
                            'Thoi_Gian': datetime.now().strftime("%Y-%m-%d %H:%M"),
                            'Ma_Dvi': c_madvi,
                            'Ten_Dvi': unit_name,
                            'Hanh_Dong': c_action,
                            'Noi_Dung': c_note,
                            'Can_Bo': c_staff
                        }])
                        st.session_state['crm_logs'] = pd.concat([st.session_state['crm_logs'], new_log], ignore_index=True)
                        st.success("Đã lưu nhật ký CRM thành công!")
                
                if len(st.session_state['crm_logs']) > 0:
                    st.markdown("#### 📋 Lịch sử tác nghiệp đã lưu trong hệ thống:")
                    st.dataframe(st.session_state['crm_logs'], use_container_width=True)

        # -------------------------------------------------------------
        # TÍNH NĂNG 4: GỬI EMAIL TỰ ĐỘNG QUA GMAIL API/SMTP
        # -------------------------------------------------------------
        elif menu == "📧 4. Gửi Email Tự Động Qua Gmail API/SMTP":
            st.subheader("📧 Hệ Thống Tự Động Gửi Email Nhắc Nợ Hàng Loạt Qua Gmail")
            st.markdown("---")
            
            st.warning("⚠️ **Lưu ý cấu hình Gmail:** Sử dụng **Mật khẩu ứng dụng (App Password)** trong tài khoản Google của bạn.")
            
            with st.form("gmail_form"):
                col_g1, col_g2 = st.columns(2)
                with col_g1:
                    sender_email = st.text_input("Email người gửi:", placeholder="your_email@gmail.com")
                    sender_password = st.text_input("Mật khẩu ứng dụng Gmail (16 ký tự):", type="password")
                with col_g2:
                    email_subject = st.text_input("Tiêu đề Email:", value="THÔNG BÁO: Đôn đốc nộp tiền BHXH, BHYT, BHTN")
                    target_unit_email = st.selectbox("Chọn đơn vị nhận email:", df_master['madvi'].astype(str) + " - " + df_master['tendvi'].astype(str))
                
                email_body = st.text_area("Nội dung Email nhắc nợ:", value="Kính gửi Quý đơn vị,\n\nQua hệ thống theo dõi thu nộp, đơn vị hiện đang có số tháng nợ vượt mức cho phép. Đề nghị đơn vị khẩn trương thanh toán để đảm bảo quyền lợi người lao động.\n\nTrân trọng!")
                
                send_btn = st.form_submit_button("🚀 Gửi Email Tự Động Ngay")
                
                if send_btn:
                    if not sender_email or not sender_password:
                        st.error("Vui lòng điền đầy đủ thông tin tài khoản gửi!")
                    else:
                        try:
                            selected_madvi = target_unit_email.split(" - ")[0]
                            row_u = df_master[df_master['madvi'].astype(str) == selected_madvi].iloc[0]
                            recipient_email = row_u.get('dsnv', 'recipient@example.com')
                            if "@" not in str(recipient_email):
                                recipient_email = "recipient@example.com"
                                
                            msg = MIMEMultipart()
                            msg['From'] = sender_email
                            msg['To'] = recipient_email
                            msg['Subject'] = email_subject
                            msg.attach(MIMEText(email_body, 'plain', 'utf-8'))
                            
                            server = smtplib.SMTP('smtp.gmail.com', 587)
                            server.starttls()
                            server.login(sender_email, sender_password)
                            server.sendmail(sender_email, recipient_email, msg.as_string())
                            server.quit()
                            
                            st.success(f"✅ Đã gửi email thành công đến đơn vị: {row_u.get('tendvi')}!")
                        except Exception as e:
                            st.error(f"❌ Gửi email thất bại: {e}")

        # -------------------------------------------------------------
        # TÍNH NĂNG 5: PHÂN TÍCH BIẾN ĐỘNG ĐA KỲ (MoM)
        # -------------------------------------------------------------
        elif menu == "📈 5. Phân Tích Biến Động Đa Kỳ (MoM)":
            st.subheader("📈 Phân Tích Biến Động Số Tháng Nợ & Số Đã Nộp Giữa 2 Kỳ")
            
            if df_prev is not None:
                st.success("Đã nhận diện dữ liệu tháng trước!")
                merged = pd.merge(df_master, df_prev, on='madvi', suffixes=('_nay', '_truoc'))
                if 'tyleno_nay' in merged.columns and 'tyleno_truoc' in merged.columns:
                    merged['Chenh_Lech_So_Thang_No'] = merged['tyleno_nay'] - merged['tyleno_truoc']
                    
                    st.markdown("#### Top đơn vị có số tháng nợ tăng mạnh nhất so với tháng trước:")
                    top_inc = merged.sort_values(by='Chenh_Lech_So_Thang_No', ascending=False).head(10)
                    st.dataframe(top_inc[['madvi', 'tendvi_nay', 'tyleno_truoc', 'tyleno_nay', 'Chenh_Lech_So_Thang_No']], use_container_width=True)
                else:
                    st.warning("Không tìm thấy cột 'tyleno' ở cả hai file để so sánh.")
            else:
                st.info("ℹ️ Vui lòng tải file Excel tháng trước ở thanh bên trái để kích hoạt.")

        # -------------------------------------------------------------
        # TÍNH NĂNG 6: SOẠN & TẢI CÔNG VĂN WORD HÀNG LOẠT
        # -------------------------------------------------------------
        elif menu == "📝 6. Soạn & Tải Công Văn Word Hàng Loạt":
            st.subheader("📝 Tự Động Hóa Soạn Thảo Công Văn Nhắc Nợ (.docx)")
            
            template_file = st.file_uploader("📤 Tải lên file Word mẫu gốc (.docx)", type=["docx"])
            
            if template_file is not None:
                target_unit = st.selectbox("Chọn đơn vị cần tạo công văn:", df_master['madvi'].astype(str) + " - " + df_master['tendvi'].astype(str))
                
                if st.button("⚙️ Tạo Công Văn Cho Đơn Vị"):
                    selected_madvi = target_unit.split(" - ")[0]
                    row_data = df_master[df_master['madvi'].astype(str) == selected_madvi].iloc[0]
                    
                    data_dict = {
                        "tendvi": str(row_data.get('tendvi', '')),
                        "madvi": str(row_data.get('madvi', '')),
                        "diachi": str(row_data.get('diachi', '')),
                        "sld": f"{row_data.get('sld', 0):,.0f}",
                        "pst": f"{row_data.get('pst', 0):,.0f}",
                        "_tien_unc": f"{row_data.get('_tien_unc', 0):,.0f}",
                        "tongThieu": f"{row_data.get('tongThieu', 0):,.0f}",
                        "tyleno": f"{row_data.get('tyleno', 0):.2f}",
                        "thang": str(row_data.get('thang', ''))
                    }
                    
                    doc = Document(template_file)
                    for p in doc.paragraphs:
                        for key, val in data_dict.items():
                            if f"{{{{{key}}}}}" in p.text:
                                p.text = p.text.replace(f"{{{{{key}}}}}", val)
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for key, val in data_dict.items():
                                    if f"{{{{{key}}}}}" in cell.text:
                                        cell.text = cell.text.replace(f"{{{{{key}}}}}", val)
                                        
                    doc_io = io.BytesIO()
                    doc.save(doc_io)
                    doc_io.seek(0)
                    
                    st.success("Đã tạo công văn thành công!")
                    st.download_button("📥 Tải xuống Công Văn (.docx)", doc_io, f"Cong_van_{selected_madvi}.docx")

        # -------------------------------------------------------------
        # TÍNH NĂNG 7: KHO DỮ LIỆU GỐC & XUẤT BÁO CÁO
        # -------------------------------------------------------------
        elif menu == "📋 7. Kho Dữ Liệu Gốc & Xuất Báo Cáo":
            st.subheader("📋 Kho Dữ Liệu Gốc & Xuất Báo Cáo Tổng Hợp")
            
            search_key = st.text_input("🔎 Tìm kiếm nhanh theo Tên hoặc Mã đơn vị:")
            if search_key:
                filtered_df = df_master[df_master['tendvi'].astype(str).str.contains(search_key, case=False, na=False) | 
                                        df_master['madvi'].astype(str).str.contains(search_key, case=False, na=False)]
            else:
                filtered_df = df_master
                
            st.dataframe(filtered_df, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, sheet_name='Ket_qua_tong_hop', index=False)
            
            st.download_button("📥 Tải xuống Báo Cáo Excel Tổng Hợp", output.getvalue(), "Bao_cao_tong_hop.xlsx")

    except Exception as e:
        st.error(f"⚠️ Đã có lỗi xảy ra khi xử lý file: {e}")
else:
    st.info("👆 Vui lòng tải lên file Excel dữ liệu thô ở thanh bên trái để khởi động toàn bộ hệ thống thông minh.")
