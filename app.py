import streamlit as st
import pandas as pd
import pdfplumber
import re
import os
import io # Dùng để xử lý file trong bộ nhớ

# --- HÀM BÓC TÁCH CHO 4PS (DUY NHẤT) ---
def parse_4ps_po(full_text, page):
    """
    Hàm này được viết RIÊNG để bóc tách PO của 4PS (cả 2 loại).
    """
    st.write("  > Nhận diện: Mẫu PO của 4PS. Đang xử lý...")
    items_list = []

    order_num_match = re.search(r"Order Number\s*:\s*(\d+)", full_text)
    delivery_date_match = re.search(r"Request Del\. Time\s*:\s*(\d{2}/\d{2}/\d{4})", full_text)
    buyer_name_match = re.search(r"Buyer Name\s*:\s*([^\n]+)", full_text)
    
    order_number = order_num_match.group(1).strip() if order_num_match else None
    delivery_date = delivery_date_match.group(1).strip() if delivery_date_match else None
    buyer_name = buyer_name_match.group(1).strip() if buyer_name_match else None

    tables = page.extract_tables({"vertical_strategy": "lines", "horizontal_strategy": "lines"})
    if not tables:
        tables = page.extract_tables() 
    
    if not tables:
        st.warning(f"  [LỖI] Không tìm thấy bảng trong file 4PS.")
        return []
        
    item_table = tables[-1] 
    
    for row in item_table[1:]:
        if row and row[1] and row[1].strip() != "" and "Total" not in row:
            # Chuẩn hóa số (4PS: 180,000.00 -> 180000.00)
            quantity_str = row[4].replace(',', '') if row[4] else '0'
            price_str = row[5].replace(',', '') if row[5] else '0'

            standard_item = {
                "Order_Number": order_number,    
                "Buyer_Name": buyer_name,      
                "Delivery_Date": delivery_date,
                "Item_Code": row[1],
                "Item_Name": row[2].replace('\n', ' '),
                "Quantity": quantity_str,
                "Price": price_str
            }
            items_list.append(standard_item)
            
    return items_list

# --- HÀM TẠO EXCEL (Giữ nguyên) ---
def create_hybrid_excel(standard_df, unrecognized_files_list):
    """
    Tạo file Excel trong bộ nhớ:
    - Sheet 1: Dữ liệu đã chuẩn hóa (standard_df)
    - Các sheet khác: Dữ liệu thô từ các file không nhận diện (unrecognized_files_list)
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # --- VIẾT SHEET 1: DỮ LIỆU ĐÃ CHUẨN HÓA (4PS) ---
        if not standard_df.empty:
            standard_df.to_excel(writer, sheet_name="TongHop_4PS", index=False)
        else:
            # Tạo sheet rỗng nếu không có dữ liệu 4PS
            pd.DataFrame(["Không có dữ liệu PO 4PS nào được tìm thấy."]).to_excel(writer, sheet_name="TongHop_4PS", index=False, header=False)
        
        # --- VIẾT CÁC SHEET THÔ (RAW DUMP) ---
        if unrecognized_files_list:
            st.write("--- Đang xử lý các file PDF khác ---")
            for uploaded_file in unrecognized_files_list:
                
                # Tạo tên sheet an toàn (loại bỏ ký tự đặc biệt, giới hạn 30 ký tự)
                safe_sheet_name = re.sub(r'[\\/*?:"<>|\[\]\s]', '_', uploaded_file.name.split('.')[0])
                safe_sheet_name = safe_sheet_name[:30]
                
                try:
                    with pdfplumber.open(uploaded_file) as pdf:
                        all_rows_for_sheet = []
                        max_cols = 0
                        
                        # Lặp qua từng trang trong PDF
                        for page in pdf.pages:
                            # Dùng chiến lược "text" để cố gắng đọc các bảng không có đường kẻ
                            tables = page.extract_tables({"vertical_strategy": "text", "horizontal_strategy": "text"})
                            if not tables:
                                tables = page.extract_tables() # Thử cả cách mặc định

                            if tables:
                                for table in tables:
                                    all_rows_for_sheet.extend(table) # Thêm dữ liệu bảng
                                    # Tìm số cột tối đa để đệm
                                    if table: # Đảm bảo bảng không rỗng
                                        max_cols = max(max_cols, max(len(r) for r in table if r))
                            
                            if max_cols > 0:
                                all_rows_for_sheet.append([None] * max_cols) # Thêm 1 dòng trống
                    
                    if all_rows_for_sheet:
                        df_raw = pd.DataFrame(all_rows_for_sheet)
                        df_raw.to_excel(writer, sheet_name=safe_sheet_name, index=False, header=False)
                        st.write(f"  > Đã dump file '{uploaded_file.name}' sang sheet '{safe_sheet_name}'")
                    else:
                        pd.DataFrame([f"Không tìm thấy bảng nào trong file {uploaded_file.name}"]).to_excel(writer, sheet_name=safe_sheet_name, index=False, header=False)

                except Exception as e:
                    st.error(f"Lỗi khi dump file {uploaded_file.name}: {e}")
                    # Ghi lỗi vào sheet
                    pd.DataFrame([f"Lỗi khi xử lý file: {e}"]).to_excel(writer, sheet_name=safe_sheet_name, index=False, header=False)

    return output.getvalue()

# --- GIAO DIỆN WEB STREAMLIT (Đã cập nhật tiêu đề) ---
st.set_page_config(page_title="Công cụ tổng hợp PO", layout="wide")
st.title("🚀 Công cụ trích xuất dữ liệu PO sang Excel")
st.write("Tải lên các file PDF của 4PS và các file PDF khác.")
st.write("Các file 4PS sẽ được gộp vào sheet 'TongHop_4PS'. Các file PDF khác sẽ được trích xuất bảng thô vào các sheet riêng.")

# 1. Khu vực tải file
uploaded_files = st.file_uploader(
    "Kéo và thả file PDF của bạn vào đây:",
    type="pdf",
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"Đã nhận {len(uploaded_files)} file. Nhấn 'Xử lý' để bắt đầu.")
    
    # 2. Nút bấm xử lý
    if st.button("Xử lý tất cả file"):
        all_standardized_data = []
        unrecognized_files = [] # <-- DANH SÁCH FILE SẼ BỊ DUMP
        
        progress_bar = st.progress(0)
        
        with st.expander("Xem chi tiết quá trình xử lý:"):
            for i, uploaded_file in enumerate(uploaded_files):
                file_name = uploaded_file.name
                st.write(f"--- Đang mở file: {file_name} ---")
                
                try:
                    with pdfplumber.open(uploaded_file) as pdf:
                        if not pdf.pages:
                            st.error(f"File {file_name} bị lỗi hoặc không có trang nào.")
                            continue
                        
                        page = pdf.pages[0]
                        full_text = page.extract_text()
                        
                        items = []
                        is_recognized = False
                        
                        # --- LOGIC NHẬN DIỆN TỰ ĐỘNG (Đã rút gọn) ---
                        if "4PS CORPORATION" in full_text or "CÔNG TY TNHH MTV KITCHEN 4PS" in full_text: 
                            customer_name = "4PS"
                            items = parse_4ps_po(full_text, page)
                            is_recognized = True
                        
                        # --- XỬ LÝ KẾT QUẢ ---
                        if is_recognized:
                            for item in items:
                                item['Customer'] = customer_name 
                                item['File_Name'] = file_name 
                                all_standardized_data.append(item)
                            st.write(f"  > Hoàn tất file 4PS. Trích xuất được {len(items)} dòng sản phẩm.")
                        else:
                            # --- NẾU KHÔNG NHẬN DIỆN, THÊM VÀO DANH SÁCH CHỜ DUMP ---
                            st.info(f"  > Không phải file 4PS. Sẽ dump file này sang sheet riêng.")
                            unrecognized_files.append(uploaded_file)
                
                except Exception as e:
                    st.error(f"!!! LỖI NGHIÊM TRỌNG khi xử lý file {file_name}: {e}")
                
                progress_bar.progress((i + 1) / len(uploaded_files))

        # 3. Xử lý dữ liệu sau khi lặp
        if not all_standardized_data and not unrecognized_files:
            st.error("Hoàn tất, nhưng không có dữ liệu nào được trích xuất.")
        else:
            df_standard = pd.DataFrame(all_standardized_data)
            
            # Chỉ xử lý/hiển thị bảng chuẩn hóa nếu có
            if not df_standard.empty:
                try:
                    df_standard['Quantity'] = pd.to_numeric(df_standard['Quantity'], errors='coerce').fillna(0)
                    df_standard['Price'] = pd.to_numeric(df_standard['Price'], errors='coerce').fillna(0)
                except Exception as e:
                    st.warning(f"[CẢNH BÁO] Không thể dọn dẹp dữ liệu số: {e}")

                columns_order = [
                    'Customer', 'Order_Number', 'Buyer_Name', 'Delivery_Date', 
                    'Item_Code', 'Item_Name', 'Quantity', 'Price', 
                    'File_Name'
                ]
                final_columns = [col for col in columns_order if col in df_standard.columns]
                df_standard = df_standard[final_columns]
                
                st.success(f"🎉 XỬ LÝ DỮ LIỆU 4PS HOÀN TẤT! 🎉")
                st.write(f"Đã gộp tổng cộng {len(df_standard)} dòng dữ liệu từ {df_standard['File_Name'].nunique()} file 4PS.")
                st.dataframe(df_standard) # Hiển thị bảng kết quả chuẩn hóa
            else:
                st.info("Không tìm thấy PO 4PS nào để gộp.")

            if unrecognized_files:
                st.info(f"Sẵn sàng dump {len(unrecognized_files)} file PDF khác sang các sheet riêng.")
            
            # 4. Nút tải file
            # Gọi hàm tạo excel "hybrid" mới
            excel_data = create_hybrid_excel(df_standard, unrecognized_files)
            
            st.download_button(
                label="📥 Tải file Excel tổng hợp (Gộp 4PS + File thô)",
                data=excel_data,
                file_name="TongHop_PO_va_FileTho.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
