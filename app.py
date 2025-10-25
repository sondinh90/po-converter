import streamlit as st
import pandas as pd
import pdfplumber
import re
import os
import io # Dùng để xử lý file trong bộ nhớ

# --- HÀM BÓC TÁCH CHO MEGA ---
def parse_mega_po(full_text, page):
    """
    Hàm này được viết RIÊNG để bóc tách PO của Mega.
    """
    st.write("  > Nhận diện: Mẫu PO của Mega Market. Đang xử lý...")
    items_list = []
    
    order_num_match = re.search(r"OUR ORDER NUMBER\s*:\s*([\d/.]+)", full_text)
    delivery_date_match = re.search(r"PLANNED DELIVERY DATE\s*:\s*([\d-]+)", full_text)
    buyer_name_match = re.search(r"BUYER\s*:\s*([^\n]+)", full_text) 
    
    order_number = order_num_match.group(1).strip() if order_num_match else None
    delivery_date = delivery_date_match.group(1).strip() if delivery_date_match else None
    buyer_name = buyer_name_match.group(1).strip() if buyer_name_match else None 

    tables = page.extract_tables({"vertical_strategy": "text", "horizontal_strategy": "text"})
    if not tables:
        st.warning(f"  [LỖI] Không tìm thấy bảng trong file Mega.")
        return []

    item_table = tables[0]
    
    for row in item_table[1:]:
        if row and row[0] and row[0].strip() != "":
            # Chuẩn hóa số (Mega: 68,000.000 -> 68000.000)
            quantity_str = row[4].replace(',', '') if row[4] else '0'
            price_str = row[5].replace(',', '') if row[5] else '0'

            standard_item = {
                "Order_Number": order_number,    
                "Buyer_Name": buyer_name,      
                "Delivery_Date": delivery_date,
                "Item_Code": row[1],
                "Item_Name": row[0].replace('\n', ' '),
                "Quantity": quantity_str, # Sẽ được convert ở hàm main
                "Price": price_str       # Sẽ được convert ở hàm main
            }
            items_list.append(standard_item)
    
    return items_list

# --- HÀM BÓC TÁCH CHO 4PS ---
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

# --- HÀM BÓC TÁCH MỚI CHO AVOLTA ---
def parse_avolta_po(full_text, page):
    """
    Hàm này được viết RIÊNG để bóc tách PO của Avolta.
    """
    st.write("  > Nhận diện: Mẫu PO của Avolta. Đang xử lý...")
    items_list = []

    # 1. Trích xuất thông tin chung
    # PO No.
    order_num_match = re.search(r"PO No\.\s*([\w-]+)", full_text)
    # Delivery Date
    delivery_date_match = re.search(r"Delivery Date\s*(\d{2}/\d{2}/\d{4})", full_text)
    # Delivery Address -> Lấy dòng đầu tiên
    buyer_name_match = re.search(r"Delivery Address\s*([^\n]+)", full_text)
    
    order_number = order_num_match.group(1).strip() if order_num_match else None
    delivery_date = delivery_date_match.group(1).strip() if delivery_date_match else None
    buyer_name = buyer_name_match.group(1).strip() if buyer_name_match else None

    # 2. Trích xuất bảng
    tables = page.extract_tables() # Avolta dùng layout đơn giản
    
    if not tables or len(tables) < 3:
        st.warning(f"  [LỖI] Không tìm thấy bảng sản phẩm trong file Avolta.")
        return []
        
    item_table = tables[-1] # Bảng sản phẩm là bảng cuối cùng
    
    # 3. Đọc dữ liệu bảng
    for row in item_table[1:]: # Bỏ qua dòng tiêu đề
        # Kiểm tra cột Item No. (row[0]) có dữ liệu không
        if row and row[0] and row[0].strip() != "" and "Total" not in row[0]:
            
            # QUAN TRỌNG: Chuẩn hóa số (Avolta: 47.259,00 -> 47259.00)
            # Quantity
            quantity_str = row[2].replace('.', '').replace(',', '.') if row[2] else '0'
            # Price
            price_str = row[4].replace('.', '').replace(',', '.') if row[4] else '0'

            standard_item = {
                "Order_Number": order_number,    
                "Buyer_Name": buyer_name,      
                "Delivery_Date": delivery_date,
                "Item_Code": row[0], # Item No.
                "Item_Name": row[1].replace('\n', ' '), # Item
                "Quantity": quantity_str,
                "Price": price_str
            }
            items_list.append(standard_item)
            
    return items_list

# --- HÀM HELPER ĐỂ CHUYỂN DF SANG EXCEL TRONG BỘ NHỚ ---
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='TongHopPO')
    processed_data = output.getvalue()
    return processed_data

# --- GIAO DIỆN WEB STREAMLIT ---
st.set_page_config(page_title="Công cụ tổng hợp PO", layout="wide")
st.title("🚀 Công cụ trích xuất dữ liệu PO sang Excel")
st.write("Tải lên các file PDF của Mega, 4PS, và Avolta để tổng hợp tự động.")

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
        progress_bar = st.progress(0)
        
        with st.expander("Xem chi tiết quá trình xử lý:"):
            for i, uploaded_file in enumerate(uploaded_files):
                file_name = uploaded_file.name
                st.write(f"--- Đang mở file: {file_name} ---")
                
                try:
                    with pdfplumber.open(uploaded_file) as pdf:
                        page = pdf.pages[0]
                        full_text = page.extract_text()
                        
                        customer_name = "Unknown"
                        items = []
                        
                        # --- LOGIC NHẬN DIỆN TỰ ĐỘNG (ĐÃ CẬP NHẬT) ---
                        if "WH 79-DALAT BBXD PLATFORM" in full_text:
                            customer_name = "Mega Market"
                            items = parse_mega_po(full_text, page)
                        elif "4PS CORPORATION" in full_text or "CÔNG TY TNHH MTV KITCHEN 4PS" in full_text: 
                            customer_name = "4PS"
                            items = parse_4ps_po(full_text, page)
                        elif "Avolta" in full_text: 
                            customer_name = "Avolta"
                            items = parse_avolta_po(full_text, page)
                        else:
                            st.error(f"  [LỖI] Không nhận diện được mẫu PO cho file: {file_name}.")
                            continue
                        
                        for item in items:
                            item['Customer'] = customer_name 
                            item['File_Name'] = file_name 
                            all_standardized_data.append(item)
                        
                        st.write(f"  > Hoàn tất file. Trích xuất được {len(items)} dòng sản phẩm.")
                
                except Exception as e:
                    st.error(f"!!! LỖI NGHIÊM TRỌNG khi xử lý file {file_name}: {e}")
                
                progress_bar.progress((i + 1) / len(uploaded_files))

        # 3. Xử lý dữ liệu sau khi lặp
        if not all_standardized_data:
            st.error("Hoàn tất, nhưng không có dữ liệu nào được trích xuất.")
        else:
            df = pd.DataFrame(all_standardized_data)
            
            # Chuyển đổi tất cả các số đã được chuẩn hóa sang dạng số
            try:
                df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
                df['Price'] = pd.to_numeric(df['Price'], errors='coerce').fillna(0)
            except Exception as e:
                st.warning(f"[CẢNH BÁO] Không thể dọn dẹp dữ liệu số: {e}")

            columns_order = [
                'Customer', 'Order_Number', 'Buyer_Name', 'Delivery_Date', 
                'Item_Code', 'Item_Name', 'Quantity', 'Price', 
                'File_Name'
            ]
            final_columns = [col for col in columns_order if col in df.columns]
            df = df[final_columns]
            
            st.success(f"🎉 XỬ LÝ HOÀN TẤT! 🎉")
            st.write(f"Đã lưu tổng cộng {len(df)} dòng dữ liệu từ {df['File_Name'].nunique()} file.")
            
            # 4. Nút tải file
            excel_data = to_excel(df)
            st.download_button(
                label="📥 Tải file Excel tổng hợp",
                data=excel_data,
                file_name="TongHop_TAT_CA_DonHang.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.dataframe(df) # Hiển thị bảng kết quả
