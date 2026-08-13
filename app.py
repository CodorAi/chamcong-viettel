import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io

st.set_page_config(page_title="Viettel - Hệ thống Chấm công AI", page_icon="🏢", layout="wide")

st.markdown("""
<style>
    .main-header { font-size: 2.5rem; color: #E60000; font-weight: bold; text-align: center; margin-bottom: 0px; }
    .sub-header { font-size: 1.1rem; color: var(--text-color); text-align: center; margin-bottom: 30px; opacity: 0.8; }
    .stButton>button { background-color: #E60000; color: white; border-radius: 8px; font-weight: bold; padding: 10px 20px; border: none; transition: 0.3s; }
    .stButton>button:hover { background-color: #cc0000; box-shadow: 0px 4px 10px rgba(230, 0, 0, 0.3); color: white;}
    
    /* Logo Viettel CSS Fake */
    .viettel-logo-container {
        text-align: center;
        padding: 10px;
        margin-bottom: 10px;
    }
    .viettel-text {
        color: #E60000;
        font-weight: 900;
        font-size: 32px;
        font-family: Arial, sans-serif;
        letter-spacing: -1px;
    }
</style>
""", unsafe_allow_html=True)

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False

if not st.session_state['logged_in']:
    st.markdown('<div class="main-header">🏢 HỆ THỐNG CHẤM CÔNG THÔNG MINH</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Phần mềm độc quyền xử lý dữ liệu kiểm soát vào ra</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.info("Vui lòng nhập Mã kích hoạt bản quyền để sử dụng")
        pwd = st.text_input("🔑 Mã kích hoạt:", type="password")
        if st.button("🚀 Kích hoạt phần mềm", use_container_width=True):
            if pwd == "VIETTEL2026":
                st.session_state['logged_in'] = True
                st.rerun()
            else:
                st.error("Mã kích hoạt không đúng hoặc đã hết hạn!")
    st.stop()

st.markdown('<div class="main-header">🏢 HỆ THỐNG CHẤM CÔNG THÔNG MINH</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Bản quyền phần mềm thuộc quản lý của Viettel</div>', unsafe_allow_html=True)

with st.sidebar:
    # Thay thế ảnh bị lỗi bằng text logo đỏ đậm chuẩn màu Viettel
    st.markdown('<div class="viettel-logo-container"><div class="viettel-text">VIETTEL</div></div>', unsafe_allow_html=True)
    st.markdown("---")
    
    st.info("💡 **Mẹo:** Chọn **⋮** góc phải ➡️ **Settings** ➡️ **Theme** để đổi Giao diện Tối/Sáng.")
    
    st.header("⚙️ Cấu hình thời gian")
    
    col_m1, col_m2 = st.columns(2)
    with col_m1: start_morning_str = st.text_input("🌅 Vào sáng", value="07:30")
    with col_m2: end_morning_str = st.text_input("🕛 Ra sáng", value="12:00")
    
    col_a1, col_a2 = st.columns(2)
    with col_a1: start_afternoon_str = st.text_input("☀️ Vào chiều", value="13:30")
    with col_a2: end_afternoon_str = st.text_input("🌆 Ra chiều", value="17:00")
    
    col_c1, col_c2 = st.columns(2)
    with col_c1: grace_period = st.number_input("⏳ Ân hạn (p)", value=15, min_value=0, step=5)
    with col_c2: split_hour = st.number_input("⏱️ Giờ tách", value=13, min_value=12, max_value=14, step=1)
    
    st.markdown("---")
    if st.button("Đăng xuất"):
        st.session_state['logged_in'] = False
        st.rerun()

st.markdown("### 1️⃣ Tải dữ liệu từ máy chấm công")
uploaded_file = st.file_uploader("Kéo thả file log Excel (.xlsx)", type=["xlsx", "xls"])

def process_time(time_str):
    try: return datetime.strptime(time_str, "%H:%M").time()
    except: return None

def calc_duration(in_t, out_t):
    if pd.isnull(in_t) or pd.isnull(out_t): return 0
    t1 = datetime.combine(datetime.today(), in_t)
    t2 = datetime.combine(datetime.today(), out_t)
    return (t2 - t1).total_seconds() / 3600

def find_header(df_temp):
    for idx, row in df_temp.iterrows():
        if any(isinstance(val, str) and 'Họ tên' in val for val in row.values): return idx
    return 0

if uploaded_file is not None:
    with st.spinner("⏳ Đang xử lý dữ liệu..."):
        try:
            df_temp = pd.read_excel(uploaded_file, header=None, nrows=20)
            header_row = find_header(df_temp)
            df = pd.read_excel(uploaded_file, header=header_row)
            
            if 'Họ tên' not in df.columns or 'Thời điểm vào' not in df.columns:
                st.error("❌ File không đúng định dạng log gốc.")
                st.stop()
                
            company_col = 'Công ty' if 'Công ty' in df.columns else None
            in_punches = df[['Họ tên', 'Thời điểm vào'] + ([company_col] if company_col else [])].rename(columns={'Thời điểm vào': 'Thời gian'})
            out_punches = df[['Họ tên', 'Thời điểm ra'] + ([company_col] if company_col else [])].rename(columns={'Thời điểm ra': 'Thời gian'})
            
            all_punches = pd.concat([in_punches, out_punches]).dropna(subset=['Thời gian', 'Họ tên'])
            all_punches['Thời gian'] = pd.to_datetime(all_punches['Thời gian'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
            all_punches = all_punches.dropna(subset=['Thời gian'])
            
            if all_punches.empty:
                st.error("Không tìm thấy dữ liệu thời gian hợp lệ.")
                st.stop()
                
            all_punches['Ngày'] = all_punches['Thời gian'].dt.date
            all_punches['Giờ'] = all_punches['Thời gian'].dt.time
            
            min_date, max_date = all_punches['Ngày'].min(), all_punches['Ngày'].max()
            all_days = [min_date + timedelta(days=x) for x in range((max_date - min_date).days + 1)]
            work_days = [d for d in all_days if d.weekday() < 5]
            
            employees = all_punches['Họ tên'].unique()
            results = []
            
            std_start_m = process_time(start_morning_str)
            std_end_m = process_time(end_morning_str)
            std_start_a = process_time(start_afternoon_str)
            std_end_a = process_time(end_afternoon_str)
            
            for emp in employees:
                emp_data = all_punches[all_punches['Họ tên'] == emp]
                comp = emp_data[company_col].iloc[0] if company_col and not emp_data[company_col].isna().all() else 'Chưa phân bổ'
                
                for wd in work_days:
                    day_data = emp_data[emp_data['Ngày'] == wd].sort_values('Thời gian')
                    date_str = wd.strftime('%d/%m/%Y')
                    
                    if day_data.empty:
                        results.append({
                            'Họ tên': emp, 'Phòng ban': comp, 'Ngày': date_str,
                            'Vào sáng': '', 'Ra sáng': '', 'Vào chiều': '', 'Ra chiều': '',
                            'Tổng giờ': 0, 'Trạng thái': 'Vắng mặt', 'Ghi chú': ''
                        })
                        continue
                    
                    morning = day_data[day_data['Thời gian'].dt.hour < split_hour]
                    afternoon = day_data[day_data['Thời gian'].dt.hour >= split_hour]
                    
                    in_m = morning.iloc[0]['Giờ'] if len(morning) > 0 else None
                    out_m = morning.iloc[-1]['Giờ'] if len(morning) > 1 else None
                    in_a = afternoon.iloc[0]['Giờ'] if len(afternoon) > 0 else None
                    out_a = afternoon.iloc[-1]['Giờ'] if len(afternoon) > 1 else None
                    
                    hrs_m = calc_duration(in_m, out_m) if out_m else 0
                    hrs_a = calc_duration(in_a, out_a) if out_a else 0
                    
                    notes = []
                    if in_m and std_start_m:
                        if (datetime.combine(datetime.today(), in_m) - datetime.combine(datetime.today(), std_start_m)).total_seconds()/60 > grace_period:
                            notes.append("Đi muộn sáng")
                    if in_a and std_start_a:
                        if (datetime.combine(datetime.today(), in_a) - datetime.combine(datetime.today(), std_start_a)).total_seconds()/60 > grace_period:
                            notes.append("Đi muộn chiều")
                    if out_a and std_end_a:
                        if (datetime.combine(datetime.today(), std_end_a) - datetime.combine(datetime.today(), out_a)).total_seconds()/60 > grace_period:
                            notes.append("Về sớm chiều")
                            
                    status = 'Đi làm' if len(notes) == 0 else 'Vi phạm'
                    fmt = lambda t: t.strftime('%H:%M:%S') if t else ''
                    
                    results.append({
                        'Họ tên': emp, 'Phòng ban': comp, 'Ngày': date_str,
                        'Vào sáng': fmt(in_m), 'Ra sáng': fmt(out_m), 'Vào chiều': fmt(in_a), 'Ra chiều': fmt(out_a),
                        'Tổng giờ': round(hrs_m + hrs_a, 2),
                        'Trạng thái': status, 'Ghi chú': ', '.join(notes)
                    })
                    
            df_res = pd.DataFrame(results)
            
            st.markdown("### 2️⃣ Bảng Điều Khiển & Bộ Lọc Báo Cáo")
            
            with st.expander("🔍 Mở bộ lọc tùy chỉnh", expanded=True):
                col_f1, col_f2, col_f3 = st.columns(3)
                all_dates = ["Tất cả"] + list(df_res['Ngày'].unique())
                selected_date = col_f1.selectbox("📅 Chọn Ngày N:", all_dates)
                
                all_depts = ["Tất cả"] + list(df_res['Phòng ban'].unique())
                selected_dept = col_f2.selectbox("🏢 Chọn Phòng ban:", all_depts)
                
                selected_status = col_f3.selectbox("⚠️ Trạng thái vi phạm:", ["Tất cả", "Vi phạm (Đi muộn/Về sớm)", "Vắng mặt", "Đi làm đúng giờ"])

            df_filtered = df_res.copy()
            if selected_date != "Tất cả": df_filtered = df_filtered[df_filtered['Ngày'] == selected_date]
            if selected_dept != "Tất cả": df_filtered = df_filtered[df_filtered['Phòng ban'] == selected_dept]
                
            if selected_status == "Vi phạm (Đi muộn/Về sớm)": df_filtered = df_filtered[df_filtered['Trạng thái'] == 'Vi phạm']
            elif selected_status == "Vắng mặt": df_filtered = df_filtered[df_filtered['Trạng thái'] == 'Vắng mặt']
            elif selected_status == "Đi làm đúng giờ": df_filtered = df_filtered[df_filtered['Trạng thái'] == 'Đi làm']
                
            st.markdown("##### 📊 Thống kê nhanh theo bộ lọc:")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Tổng số lượt", len(df_filtered))
            c2.metric("Số lượt Vắng", len(df_filtered[df_filtered['Trạng thái'] == 'Vắng mặt']))
            c3.metric("Số lượt Vi phạm", len(df_filtered[df_filtered['Trạng thái'] == 'Vi phạm']))
            c4.metric("Số lượt Đi làm chuẩn", len(df_filtered[df_filtered['Trạng thái'] == 'Đi làm']))
            
            st.dataframe(df_filtered, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_filtered.to_excel(writer, index=False, sheet_name='DuLieuDaLoc')
                df_res.to_excel(writer, index=False, sheet_name='ToanBoDuLieu')
            
            st.markdown("---")
            colA, colB, colC = st.columns([1,2,1])
            with colB:
                st.download_button(
                    label=f"📥 TẢI XUỐNG BÁO CÁO ĐÃ LỌC EXCEL",
                    data=output.getvalue(),
                    file_name="Bao_Cao_Cham_Cong.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        except Exception as e:
            st.error(f"Đã xảy ra lỗi: {str(e)}")
