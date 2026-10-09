import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(
    page_title="Sales Tracking & Building Block",
    page_icon="🍺",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom styling for mobile and clean UI
st.markdown("""
<style>
    .metric-card {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 12px;
        border-left: 5px solid #1f4e79;
        margin-bottom: 10px;
    }
    .metric-title {
        font-size: 13px;
        color: #6c757d;
        font-weight: bold;
    }
    .metric-val {
        font-size: 20px;
        color: #1f4e79;
        font-weight: bold;
    }
    [data-testid="stMetricValue"] {
        font-size: 20px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("<h2 style='text-align: center; color: #1F4E79;'>📱 THEO DÕI ĐIỂM BÁN & SỐ BÁN HÀNG NGÀY</h2>", unsafe_allow_html=True)

# --- SIDEBAR: QUẢN TRỊ TẢI FILE ---
with st.sidebar:
    st.header("⚙️ Khu vực Quản lý")
    st.info("💡 Bạn chỉ cần tải 2 file gốc lên đây mỗi ngày:")
    file_sec = st.file_uploader("1. Tải lên SEC-06 (.xlsx)", type=["xlsx"])
    file_tracking = st.file_uploader("2. Tải lên Tracking Master (.xlsx)", type=["xlsx"])
    st.caption("Sau khi tải lên, hệ thống sẽ tự động ghép mã Outlet, tính toán Act Huda, Act Pre và phân chia theo từng SR.")

if not file_sec or not file_tracking:
    st.warning("👈 Vui lòng tải lên cả 2 file: **SEC-06** và **Tracking** ở thanh menu bên trái để hệ thống bắt đầu xử lý.")
    st.stop()

# --- XỬ LÝ DỮ LIỆU ---
@st.cache_data
def process_data(sec_f, track_f):
    df_sec_raw = pd.read_excel(sec_f, sheet_name="Detail")
    df_track_raw = pd.read_excel(track_f, sheet_name=" Display TOFT", header=None)
    
    # 1. Chuẩn hóa Tracking Master
    track_rows = []
    for idx in range(4, len(df_track_raw)):
        r = df_track_raw.iloc[idx].tolist()
        ocode = str(r[5]).strip() if pd.notna(r[5]) else ""
        if not ocode or ocode == "nan":
            continue
        track_rows.append({
            "Mã CT": str(r[3]).strip() if pd.notna(r[3]) else "",
            "Huyện": str(r[4]).strip() if pd.notna(r[4]) else "",
            "Mã điểm bán": ocode,
            "Tên điểm bán": str(r[6]).strip() if pd.notna(r[6]) else "",
            "Địa chỉ": str(r[7]).strip() if pd.notna(r[7]) else "",
            "SR quản lý": str(r[8]).strip() if pd.notna(r[8]) else "",
            "Level": str(r[9]).strip() if pd.notna(r[9]) else "",
            "Tar Huda": float(r[10]) if pd.notna(r[10]) and str(r[10]).replace('.','',1).isdigit() else 0.0,
            "Tar Pre": float(r[13]) if pd.notna(r[13]) and str(r[13]).replace('.','',1).isdigit() else 0.0,
        })
    df_master = pd.DataFrame(track_rows)
    df_master["Tar Tổng"] = df_master["Tar Huda"] + df_master["Tar Pre"]
    
    # 2. Xử lý SEC-06 Detail
    df_sec_raw["OutletCode"] = df_sec_raw["OutletCode"].astype(str).str.strip()
    df_sec_raw["Product"] = df_sec_raw["Product"].astype(str)
    df_sec_raw["VolumeHL"] = pd.to_numeric(df_sec_raw["VolumeHL"], errors="coerce").fillna(0.0)
    
    # Phân loại Huda vs Premium
    df_sec_raw["LoaiBia"] = "Khác"
    df_sec_raw.loc[df_sec_raw["Product"].str.contains("Huda", case=False, na=False), "LoaiBia"] = "Huda"
    df_sec_raw.loc[df_sec_raw["Product"].str.contains("Tuborg|1664", case=False, na=False), "LoaiBia"] = "Pre"
    
    # Gom tổng Actual theo từng OutletCode
    agg_sales = df_sec_raw[df_sec_raw["LoaiBia"].isin(["Huda", "Pre"])].groupby(["OutletCode", "LoaiBia"])["VolumeHL"].sum().unstack(fill_value=0.0).reset_index()
    if "Huda" not in agg_sales.columns: agg_sales["Huda"] = 0.0
    if "Pre" not in agg_sales.columns: agg_sales["Pre"] = 0.0
    agg_sales.rename(columns={"Huda": "Act Huda", "Pre": "Act Pre"}, inplace=True)
    
    # Ghép vào Tracking Master
    df_merged = pd.merge(df_master, agg_sales, left_on="Mã điểm bán", right_on="OutletCode", how="left")
    df_merged["Act Huda"] = df_merged["Act Huda"].fillna(0.0)
    df_merged["Act Pre"] = df_merged["Act Pre"].fillna(0.0)
    df_merged["Act Tổng"] = df_merged["Act Huda"] + df_merged["Act Pre"]
    
    # Đánh giá Đạt / Chưa Đạt
    df_merged["Trạng thái"] = df_merged.apply(
        lambda row: "Đạt" if (row["Act Huda"] >= row["Tar Huda"] and row["Act Pre"] >= row["Tar Pre"]) else "Chưa Đạt", 
        axis=1
    )
    
    return df_merged, df_sec_raw, df_master

df_final, df_sec_detail, df_master = process_data(file_sec, file_tracking)

# --- 3 TABS CHỨC NĂNG CHÍNH ---
tab_sr, tab_daily, tab_audit = st.tabs([
    "📊 Tiến độ từng SR", 
    "📅 Theo dõi số bán ngày", 
    "⚠️ Cảnh báo điểm mới"
])

# ================= TAB 1: TIẾN ĐỘ TỪNG SR =================
with tab_sr:
    srs = sorted(df_final["SR quản lý"].unique().tolist())
    selected_sr = st.selectbox("👉 Chọn tên Sales (SR):", srs)
    
    df_sr = df_final[df_final["SR quản lý"] == selected_sr].copy()
    
    # KPI Cards trên điện thoại
    col1, col2, col3 = st.columns(3)
    total_outlets = len(df_sr)
    achieved = len(df_sr[df_sr["Trạng thái"] == "Đạt"])
    pct_achieved = (achieved / total_outlets * 100) if total_outlets else 0
    
    act_huda = df_sr['Act Huda'].sum()
    tar_huda = df_sr['Tar Huda'].sum()
    act_pre = df_sr['Act Pre'].sum()
    tar_pre = df_sr['Tar Pre'].sum()
    
    col1.metric("Điểm bán", f"{achieved}/{total_outlets} Đạt", f"{pct_achieved:.0f}%")
    col2.metric("Huda (HL)", f"{act_huda:.2f}", f"Target: {tar_huda:.2f}")
    col3.metric("Pre (HL)", f"{act_pre:.2f}", f"Target: {tar_pre:.2f}")
    
    st.markdown("---")
    st.subheader(f"Danh sách điểm bán: {selected_sr}")
    
    # Bảng chi tiết gọn nhẹ cho di động
    display_df = df_sr[["Mã điểm bán", "Tên điểm bán", "Act Huda", "Tar Huda", "Act Pre", "Tar Pre", "Trạng thái"]].copy()
    display_df.rename(columns={
        "Act Huda": "Act Huda (HL)", "Tar Huda": "Tar Huda (HL)",
        "Act Pre": "Act Pre (HL)", "Tar Pre": "Tar Pre (HL)"
    }, inplace=True)
    
    # Định dạng hiển thị
    st.dataframe(
        display_df.style.format({
            "Act Huda (HL)": "{:.2f}",
            "Tar Huda (HL)": "{:.2f}",
            "Act Pre (HL)": "{:.2f}",
            "Tar Pre (HL)": "{:.2f}"
        }).applymap(
            lambda v: "color: green; font-weight: bold;" if v == "Đạt" else ("color: red;" if v == "Chưa Đạt" else ""),
            subset=["Trạng thái"]
        ),
        use_container_width=True,
        hide_index=True
    )

# ================= TAB 2: THEO DÕI SỐ BÁN NGÀY =================
with tab_daily:
    st.subheader("📅 Diễn biến số bán thực tế theo ngày")
    
    if "OrderedDate" in df_sec_detail.columns:
        df_sec_detail["Ngay"] = pd.to_datetime(df_sec_detail["OrderedDate"]).dt.strftime("%d/%m")
        daily_sales = df_sec_detail[df_sec_detail["LoaiBia"].isin(["Huda", "Pre"])].groupby(["Ngay", "LoaiBia"])["VolumeHL"].sum().reset_index()
        
        # Biểu đồ cột
        fig = px.bar(
            daily_sales, x="Ngay", y="VolumeHL", color="LoaiBia", barmode="group",
            title="Sản lượng Huda vs Premium theo ngày đặt hàng",
            labels={"VolumeHL": "Sản lượng (HL)", "Ngay": "Ngày", "LoaiBia": "Dòng bia"},
            color_discrete_map={"Huda": "#008080", "Pre": "#FF7F0E"}
        )
        st.plotly_chart(fig, use_container_width=True)
        
        # Bảng chi tiết ngày theo từng bạn SR
        st.markdown("#### Bảng sản lượng (HL) từng ngày theo SR:")
        pivot_sr = df_sec_detail[df_sec_detail["LoaiBia"].isin(["Huda", "Pre"])].pivot_table(
            index="Fullname", columns="Ngay", values="VolumeHL", aggfunc="sum", fill_value=0.0
        )
        st.dataframe(pivot_sr.style.format("{:.2f}"), use_container_width=True)
    else:
        st.info("Không tìm thấy trường OrderedDate trong file SEC-06.")

# ================= TAB 3: CẢNH BÁO ĐIỂM MỚI =================
with tab_audit:
    st.subheader("⚠️ Điểm phát sinh đơn trên SEC-06 nhưng chưa có trong Tracking")
    sec_outlets = set(df_sec_detail["OutletCode"].unique())
    master_outlets = set(df_master["Mã điểm bán"].unique())
    diff = sec_outlets - master_outlets
    
    if diff:
        st.warning(f"Có **{len(diff)}** điểm bán có đơn hàng trên Sec 6 nhưng chưa có trong danh sách Tracking Master!")
        df_unreg = df_sec_detail[df_sec_detail["OutletCode"].isin(diff)][
            ["OutletCode", "OutletName", "Fullname", "Product", "VolumeHL"]
        ].drop_duplicates()
        df_unreg.rename(columns={
            "OutletCode": "Mã điểm (SEC-06)",
            "OutletName": "Tên điểm bán",
            "Fullname": "SR bán trên đơn",
            "Product": "Sản phẩm",
            "VolumeHL": "Sản lượng (HL)"
        }, inplace=True)
        st.dataframe(df_unreg, use_container_width=True, hide_index=True)
    else:
        st.success("Tất cả điểm bán trên Sec 6 đều khớp chính xác với Tracking Master!")
