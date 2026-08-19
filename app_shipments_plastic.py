import io
import unicodedata
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Lọc Shipper Ngành Nhựa Có XNK Tháng Này", layout="wide"
)
st.title("📦 Danh Sách Shipper Ngành Nhựa Có Phát Sinh XNK Tháng Này")

# --- 1. UPLOAD FILE ---
st.sidebar.header("1. Tải lên dữ liệu")
file_curr = st.sidebar.file_uploader(
    "Upload file B/L Tháng Này (Excel)", type=["xlsx", "xls"]
)
file_info = st.sidebar.file_uploader(
    "Upload file Master Data Shipper (Excel/CSV)", type=["xlsx", "xls", "csv"]
)

# --- 2. CẤU HÌNH TÌM KIẾM ---
st.sidebar.header("2. Cấu hình tìm kiếm")
keyword_input = st.sidebar.text_input(
    "Từ khóa ngành nghề cần lọc:", value="nhựa"
)


def clean_str(val):
  if pd.isna(val):
    return ""
  val = str(val)
  val = unicodedata.normalize("NFC", val)
  val = (
      val.replace("\xa0", " ")
      .replace("\t", " ")
      .replace("\n", " ")
      .replace("\r", "")
  )
  return " ".join(val.split())


# Chuẩn hóa Mã Số Thuế (Xóa .0 ở cuối nếu bị pandas đọc thành float, xóa khoảng trắng)
def clean_tax_code(val):
  if pd.isna(val):
    return ""
  val_str = str(val).strip()
  if val_str.endswith(".0"):
    val_str = val_str[:-2]
  return val_str


def load_bl_file_all_sheets(uploaded_file, col_tax_target):
  all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)
  combined_list = []

  for sheet_name, df_raw in all_sheets.items():
    sheet_clean = str(sheet_name).strip()

    header_idx = None
    for idx, row in df_raw.head(20).iterrows():
      row_str = [clean_str(val).upper() for val in row.values]
      if any(col_tax_target.upper() in val for val in row_str):
        header_idx = idx
        break

    if header_idx is not None:
      df_sheet = df_raw.iloc[header_idx + 1 :].copy()
      df_sheet.columns = [clean_str(val) for val in df_raw.iloc[header_idx]]
      df_sheet = df_sheet.loc[:, ~df_sheet.columns.duplicated()]

      if col_tax_target in df_sheet.columns:
        df_sheet["Loại hình / Cảng"] = sheet_clean
        combined_list.append(df_sheet)

  if combined_list:
    return pd.concat(combined_list, ignore_index=True)
  else:
    return pd.DataFrame()


def auto_load_file_info(file_obj, col_tax_target):
  if file_obj.name.endswith(".csv"):
    df_raw = pd.read_csv(file_obj, header=None)
  else:
    excel_info = pd.ExcelFile(file_obj)
    sheet_names = excel_info.sheet_names
    df_raw = pd.read_excel(file_obj, sheet_name=sheet_names[0], header=None)

  header_idx = None
  for idx, row in df_raw.head(20).iterrows():
    row_str = [clean_str(val).upper() for val in row.values]
    if any(col_tax_target.upper() in val for val in row_str):
      header_idx = idx
      break

  if header_idx is not None:
    df_final = df_raw.iloc[header_idx + 1 :].copy()
    df_final.columns = [clean_str(val) for val in df_raw.iloc[header_idx]]
  else:
    df_raw.columns = [clean_str(val) for val in df_raw.iloc[0]]
    df_final = df_raw.iloc[1:].copy()

  df_final = df_final.loc[:, ~df_final.columns.duplicated()]
  return df_final.reset_index(drop=True)


def convert_df_to_excel(df, sheet_name_out):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name=sheet_name_out)
  return output.getvalue()


# --- XỬ LÝ CHÍNH ---
if file_curr and file_info:
  try:
    col_tax = "MST SHIPPER"
    col_shipper = "TÊN SHIPPER TRÊN B/L"
    col_agent = "AGENT HANDLE NAME & SHIPPER"
    col_industry = "Ngành nghề KD"

    # 1. Đọc File B/L
    df_curr_raw = load_bl_file_all_sheets(file_curr, col_tax)

    # 2. Đọc File Master Data
    df_info = auto_load_file_info(file_info, col_tax)

    if df_curr_raw.empty:
      st.error(
          "❌ **Không tìm thấy dữ liệu hoặc cột `MST SHIPPER` trong File B/L!**"
      )

    elif (col_tax not in df_info.columns) or (col_industry not in df_info.columns):
      st.error("❌ **Thiếu cột thông tin trong File Master Data!**")
      raw_cols = list(df_info.columns)
      st.write("📋 **Tên các cột đọc được từ Master Data:**")
      st.code(raw_cols)

      st.sidebar.markdown("---")
      st.sidebar.warning("⚠️ Chọn cột thủ công:")
      col_tax_manual = st.sidebar.selectbox("Cột Mã Số Thuế:", raw_cols)
      col_ind_manual = st.sidebar.selectbox("Cột Ngành Nghề KD:", raw_cols)
      col_tax = col_tax_manual
      col_industry = col_ind_manual

    else:
      # LÀM SẠCH CHUẨN HÓA MÃ SỐ THUẾ DÙNG CHUNG HÀM CLEAN_TAX_CODE
      df_curr_raw = df_curr_raw.dropna(subset=[col_tax])
      df_curr_raw[col_tax] = df_curr_raw[col_tax].apply(clean_tax_code)

      df_info = df_info.dropna(subset=[col_tax])
      df_info[col_tax] = df_info[col_tax].apply(clean_tax_code)

      # Chuẩn hóa cột Ngành nghề sang Unicode NFC
      df_info[col_industry] = df_info[col_industry].apply(clean_str)

      # Giữ lại các cột cần thiết từ B/L
      cols_to_keep = [col_tax]
      if col_shipper in df_curr_raw.columns:
        cols_to_keep.append(col_shipper)
      if col_agent in df_curr_raw.columns:
        cols_to_keep.append(col_agent)
      if "Loại hình / Cảng" in df_curr_raw.columns:
        cols_to_keep.append("Loại hình / Cảng")

      df_curr_clean = df_curr_raw[cols_to_keep].drop_duplicates(subset=[col_tax])
      df_info_clean = df_info[[col_tax, col_industry]].drop_duplicates(
          subset=[col_tax]
      )

      # Merge ghép bảng
      merged_df = pd.merge(
          df_curr_clean, df_info_clean, on=col_tax, how="inner"
      )

      # Chuẩn hóa từ khóa tìm kiếm sang Unicode NFC
      kw = clean_str(keyword_input).lower()

      if kw:
        condition = (
            merged_df[col_industry]
            .astype(str)
            .str.lower()
            .str.contains(kw, na=False)
        )
        final_df = merged_df[condition]
      else:
        final_df = merged_df

      # --- PHẦN BÁO CÁO THỐNG KÊ (DEBUG VIEW) ---
      st.subheader(
          f"🔍 Tìm thấy {len(final_df)} Shipper khớp điều kiện"
      )

      if len(final_df) > 0:
        st.dataframe(final_df, use_container_width=True)

        excel_data = convert_df_to_excel(final_df, "Shipper_XNK_Thang_Nay")

        st.download_button(
            label="📥 Tải danh sách Excel (.xlsx)",
            data=excel_data,
            file_name=f"shipper_{kw}_xnk_thang_nay.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
      else:
        st.warning(
            f"Báo cáo trống: Không có MST nào khớp ngành nghề '{keyword_input}'"
            " xuất hiện trong dữ liệu B/L tháng này."
        )

        # KHU VỰC SOI LỖI KHI BÁO CÁO TRỐNG
        with st.expander("🛠️ Nhấp vào đây để soi vị trí đứt gãy dữ liệu (Debug)"):
          st.write(
              f"1. Số lượng MST đọc được từ File B/L:"
              f" **{len(df_curr_clean)}**"
          )
          st.write(
              f"2. Số lượng MST đọc được từ File Master Data:"
              f" **{len(df_info_clean)}**"
          )
          st.write(
              f"3. Số MST trong File B/L KHỚP ĐƯỢC với File Master Data:"
              f" **{len(merged_df)}**"
          )
          st.write("4. Mẫu 5 dòng Ngành nghề KD trong Master Data:")
          st.dataframe(df_info_clean[[col_tax, col_industry]].head(5))

  except Exception as e:
    st.error(f"Xảy ra lỗi khi xử lý dữ liệu: {e}")
else:
  st.info(
      "👋 Vui lòng upload File B/L Tháng Này và File Master Data ở thanh bên"
      " trái."
  )
