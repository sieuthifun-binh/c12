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
from email.mime.base import MIMEBase
from email import encoders

st.set_page_config(page_title="Hệ Thống Quản Lý & Tác Nghiệp Thu Nộp Thông Minh", layout="wide")

# Cấu hình font tiếng Việt cho biểu đồ
plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['axes.unicode_minus'] = False

st.title("🚀 HỆ THỐNG QUẢN LÝ THU NỘP & TÁC NGHIỆP THÔNG MINH (ENTERPRISE EDITION)")
st.markdown("---")

# Khởi tạo bộ nhớ tạm cho CRM Activity Log nếu chưa có
if 'crm_logs' not in st.session_state:
    st.session_state['crm_logs'] = pd.DataFrame(columns=['Thoi_Gian', 'Ma_Dvi', 'Ten_Dvi', 'Hanh_Dong', 'Noi_Dung', 'Can_Bo'])

# 1. KHU VỰC NẠP DỮ LIỆU (Hỗ trợ đa kỳ để so sánh)
st.sidebar.header("📁 1. Nạp Dữ Liệu Hệ Thống")
uploaded_file = st.sidebar.file_uploader("Tải lên file Excel tháng hiện tại (.xls, .xlsx)", type=["xls", "xlsx"])
uploaded_file_prev = st.sidebar.file_uploader("Tải lên file Excel tháng trước (Tùy chọn - để so sánh biến động)", type=["xls", "xlsx"])

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

        # Menu điều hướng 7 tính năng
        menu = st.sidebar.selectbox(
            "⚙️ Chọn Chức Nặng Tác Nghiệp:", 
            [
                "📊 1. Dashboard & Dự Báo Rủi Ro AI", 
                "🤖 2. Trợ Lý AI Tra Cứu Ngôn Ngữ Tự Nhiên", 
                "👤 3. Quản Lý Chuyên Quản & CRM Tác Nghiệp", 
                "📧 4. Gửi Email Tự Động Qua Gmail API/SMTP", 
                "📈 5. Phân Tích Biến Động Đa Kỳ (MoM)", 
                "📝 6. Soạn & Tải Công Văn Word Hàng Loạt", 
                "📋 7. Kho Dữ Liệu Gốc & Xuất Báo Cáo"
            ]
        )
        
        # -------------------------------------------------------------
        # TÍNH NĂNG 1: DASHBOARD & DỰ BÁO RỦI RO AI
        # -------------------------------------------------------------
        if menu == "📊 1. Dashboard & Dự Báo Rủi Ro AI":
            st.subheader("📊 Tổng Quan & Chấm Điểm Dự Báo Rủi Ro Tự Động")
            
            total_units = len(df_master)
            total_labor = df_master['sld'].sum() if 'sld' in df_master else 0
            total_pst = df_master['pst'].sum() if 'pst' in df_master else 0
            total_debt = df_master['tongThieu'].sum() if 'tongThieu' in df_master else 0
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Tổng Số Đơn Vị", f"{total_units:,}")
            c2.metric("Tổng Số Lao Động", f"{total_labor:,.0f}")
            c3.metric("Tổng Phát Sinh Thu", f"{total_pst:,.0f} đ")
            c4.metric("Tổng Số Tiền Nợ", f"{total_debt:,.0f} đ")
            
            st.markdown("---")
            
            # Gán nhãn rủi ro thông minh
            def ai_risk_scoring(row):
                debt = row.get('tongThieu', 0)
                labor = row.get('sld', 0)
                if pd.isna(debt) or debt <= 0:
                    return "🟢 An toàn"
                elif debt < 50000000:
                    return "🟡 Cảnh báo nhắc nhở"
                else:
                    return "🔴 Nguy cơ nợ xấu cao"
            
            df_master['Muc_Do_Rui_Ro'] = df_master.apply(ai_risk_scoring, axis=1)
            risk_counts = df_master['Muc_Do_Rui_Ro'].value_counts()
            
            rc1, rc2, rc3 = st.columns(3)
            rc1.metric("🔴 Nguy cơ nợ xấu cao", f"{risk_counts.get('🔴 Nguy cơ nợ xấu cao', 0):,}")
            rc2.metric("🟡 Cảnh báo nhắc nhở", f"{risk_counts.get('🟡 Cảnh báo nhắc nhở', 0):,}")
            rc3.metric("🟢 An toàn", f"{risk_counts.get('🟢 An toàn', 0):,}")
            
            st.markdown("---")
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("#### 🔴 Top 10 Đơn Vị Nợ Cao Nhất")
                top10 = df_master.sort_values(by='tongThieu', ascending=False).head(10)
                fig, ax = plt.subplots(figsize=(8, 6))
                ax.barh(top10['tendvi'].astype(str).str[:30], top10['tongThieu'] / 1e6, color='#e74c3c')
                ax.set_xlabel('Tổng nợ (Triệu VNĐ)', fontweight='bold')
                ax.invert_yaxis()
                st.pyplot(fig)
            with col_b:
                st.markdown("#### 🔵 Tỷ Trọng Phân Bổ Mức Độ Rủi Ro")
                fig2, ax2 = plt.subplots(figsize=(7, 7))
                ax2.pie(risk_counts.values, labels=risk_counts.index, autopct='%1.1f%%', colors=['#2ecc71', '#f1c40f', '#e74c3c'], startangle=140)
                st.pyplot(fig2)

        # -------------------------------------------------------------
        # TÍNH NĂNG 2: TRỢ LÝ AI TRA CỨU NGÔN NGỮ TỰ NHIÊN
        # -------------------------------------------------------------
        elif menu == "🤖 2. Trợ Lý AI Tra Cứu Ngôn Ngữ Tự Nhiên":
            st.subheader("🤖 Trợ Lý Ảo Truy Vấn Dữ Liệu Bằng Ngôn Ngữ Tự Nhiên")
            st.info("💡 **Ví dụ câu lệnh:** Nhập `nợ > 50000000` hoặc `lao động > 100` hoặc tìm theo khối `TH`.")
            
            query = st.text_input("💬 Nhập yêu cầu tìm kiếm của bạn:")
            if query:
                try:
                    query_lower = query.lower()
                    if "nợ >" in query_lower:
                        val = float(query_lower.split("nợ >")[1].strip())
                        result_df = df_master[df_master['tongThieu'] > val]
                    elif "lao động >" in query_lower:
                        val = float(query_lower.split("lao động >")[1].strip())
                        result_df = df_master[df_master['sld'] > val]
                    elif "khối" in query_lower:
                        val = query_lower.split("khối")[1].strip().upper()
                        result_df = df_master[df_master['makhoi'].astype(str).str.contains(val, na=False)]
                    else:
                        result_df = df_master[df_master['tendvi'].astype(str).str.contains(query, case=False, na=False)]
                    
                    st.success(f"Tìm thấy **{len(result_df):,}** đơn vị phù hợp:")
                    st.dataframe(result_df[['madvi', 'tendvi', 'makhoi', 'sld', 'tongThieu', 'diachi']], use_container_width=True)
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
                st.dataframe(df_quan[['madvi', 'tendvi', 'sld', 'tongThieu']], use_container_width=True)
                
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
        # TÍNH NĂNG 4: GỬI EMAIL TỰ ĐỘNG QUA GIMAI API/SMTP
        # -------------------------------------------------------------
        elif menu == "📧 4. Gửi Email Tự Động Qua Gmail API/SMTP":
            st.subheader("📧 Hệ Thống Tự Động Gửi Email Nhắc Nợ Hàng Loạt Qua Gmail")
            st.markdown("---")
            
            st.warning("⚠️ **Lưu ý cấu hình Gmail:** Để gửi email tự động qua Python, bạn cần sử dụng **Mật khẩu ứng dụng (App Password)** của tài khoản Gmail cá nhân/cơ quan (Bật xác thực 2 bước trên Google Account -> Tạo App Password).")
            
            with st.form("gmail_form"):
                col_g1, col_g2 = st.columns(2)
                with col_g1:
                    sender_email = st.text_input("Email người gửi (Gmail của bạn):", placeholder="your_email@gmail.com")
                    sender_password = st.text_input("Mật khẩu ứng dụng Gmail (App Password 16 ký tự):", type="password")
                with col_g2:
                    email_subject = st.text_input("Tiêu đề Email:", value="THÔNG BÁO: Đôn đốc nộp tiền BHXH, BHYT, BHTN kỳ tháng 2026")
                    target_unit_email = st.selectbox("Chọn đơn vị nhận email:", df_master['madvi'].astype(str) + " - " + df_master['tendvi'].astype(str))
                
                email_body = st.text_area("Nội dung Email nhắc nợ:", value="Kính gửi Quý đơn vị,\n\nQua hệ thống theo dõi thu nộp BHXH, đơn vị hiện đang có phát sinh số tiền nợ chưa hoàn thành. Đề nghị đơn vị khẩn trương nộp tiền vào tài khoản chuyên thu để đảm bảo quyền lợi cho người lao động.\n\nTrân trọng cảm ơn!")
                
                send_btn = st.form_submit_button("🚀 Gửi Email Tự Động Ngay")
                
                if send_btn:
                    if not sender_email or not sender_password:
                        st.error("Vui lòng điền đầy đủ Email và Mật khẩu ứng dụng!")
                    else:
                        try:
                            # Lấy email người liên hệ hoặc dùng email mẫu
                            selected_madvi = target_unit_email.split(" - ")[0]
                            row_u = df_master[df_master['madvi'].astype(str) == selected_madvi].iloc[0]
                            
                            # Tìm cột email trong file nếu có, nếu không cho phép nhập
                            recipient_email = row_u.get('dsnv', 'recipient@example.com') # Hoặc lấy từ cột email nếu có
                            if "@" not in str(recipient_email):
                                recipient_email = "recipient@example.com" # Fallback
                                
                            # Thiết lập kết nối SMTP Gmail
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
                            
                            st.success(f"✅ Đã gửi email nhắc nợ thành công đến đơn vị: {row_u.get('tendvi')} (Gửi tới: {recipient_email})!")
                        except Exception as e:
                            st.error(f"❌ Gửi email thất bại. Lỗi kết nối SMTP: {e}")

        # -------------------------------------------------------------
        # TÍNH NĂNG 5: PHÂN TÍCH BIẾN ĐỘNG ĐA KỲ (MoM)
        # -------------------------------------------------------------
        elif menu == "📈 5. Phân Tích Biến Động Đa Kỳ (MoM)":
            st.subheader("📈 Phân Tích Biến Động Nợ Giữa 2 Kỳ (Tháng trước vs Tháng này)")
            
            if df_prev is not None:
                st.success("Đã nhận diện dữ liệu tháng trước! Đang tiến hành khớp nối so sánh...")
                
                # Merge dữ liệu theo madvi
                merged = pd.merge(df_master, df_prev, on='madvi', suffixes=('_nay', '_truoc'))
                if 'tongThieu_nay' in merged.columns and 'tongThieu_truoc' in merged.columns:
                    merged['Chenh_Lech_No'] = merged['tongThieu_nay'] - merged['tongThieu_truoc']
                    
                    col_m1, col_m2 = st.columns(2)
                    col_m1.metric("Tổng nợ tháng này", f"{merged['tongThieu_nay'].sum():,.0f} đ")
                    col_m2.metric("So với tháng trước", f"{merged['Chenh_Lech_No'].sum():,.0f} đ", delta_color="inverse")
                    
                    st.markdown("#### Top đơn vị nợ tăng mạnh nhất so với tháng trước:")
                    top_increase = merged.sort_values(by='Chenh_Lech_No', ascending=False).head(10)
                    st.dataframe(top_increase[['madvi', 'tendvi_nay', 'tongThieu_truoc', 'tongThieu_nay', 'Chenh_Lech_No']], use_container_width=True)
                else:
                    st.warning("Không tìm thấy cột 'tongThieu' ở cả hai file để so sánh.")
            else:
                st.info("ℹ️ Vui lòng tải thêm file Excel của **tháng trước** ở thanh bên trái (`Tải lên file Excel tháng trước`) để kích hoạt tính năng phân tích đa kỳ này.")

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
                        "tongThieu": f"{row_data.get('tongThieu', 0):,.0f}",
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
