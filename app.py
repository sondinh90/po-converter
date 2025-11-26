import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

# ==========================================
# 1. CÁC HÀM HỖ TRỢ (HELPER) - ĐÃ SỬA LOGIC SỐ
# ==========================================

def clean_avolta_number(num_str):
    """
    Xử lý số thông minh: Tự động phát hiện kiểu Âu hay kiểu Mỹ.
    """
    if not num_str: return 0.0
    s = str(num_str).strip()
    
    # Loại bỏ các ký tự lạ, chỉ giữ lại số, chấm, phẩy, trừ
    s = re.sub(r'[^\d.,-]', '', s)
    
    # TRƯỜNG HỢP 1: Có dấu phẩy (,) -> Khả năng cao là kiểu Âu (1.200,50)
    if ',' in s:
        # Nếu có cả chấm và phẩy (vd: 1.200,50) -> Bỏ chấm, thay phẩy = chấm
        if '.' in s:
            s = s.replace('.', '')
            s = s.replace(',', '.')
        # Nếu chỉ có phẩy (vd: 1200,50) -> Thay phẩy = chấm
        else:
            s = s.replace(',', '.')
            
    # TRƯỜNG HỢP 2: Không có phẩy, chỉ có chấm (vd: 10.00 hoặc 1.000)
    # Đây là ca khó. Thường Avolta dùng chấm làm ngàn (1.000).
    # Nhưng nếu PDF extract ra là 10.00 (mười) thì xóa chấm sẽ thành 1000 (sai).
    
    # Logic sửa đổi:
    # Nếu có chấm:
    # - Nếu phần sau dấu chấm có đúng 3 ký tự (vd 1.000) -> Nghi ngờ là ngàn -> Xóa chấm
    # - Nếu phần sau dấu chấm khác 3 ký tự (vd 10.00, 46.35) -> Nghi ngờ là thập phân -> Giữ nguyên
    elif '.' in s:
        parts = s.split('.')
        # Nếu phần đuôi có đúng 3 số (vd 46.350) -> Rất có thể là 46 ngàn
        if len(parts) > 1 and len(parts[-1]) == 3:
             s = s.replace('.', '')
        # Ngược lại (vd 46.35 hoặc 10.00) -> Giữ nguyên dấu chấm là thập phân
        else:
             pass 

    try:
        return float(s)
    except ValueError:
        return 0.0

# ==========================================
# 2. HÀM BÓC TÁCH 4PS (TABLE STRATEGY)
# ==========================================
def parse_4ps_po(pdf):
    st.write("  > Nhận diện: Mẫu PO của 4PS. Đang xử lý...")
    items_list = []

    # Lấy thông tin Header từ trang 1
    page1 = pdf.pages[0]
    full_text_page1 = page1.extract_text() 

    order_num_match = re.search(r"Order Number\s*:\s*(\d+)", full_text_page1)
    delivery_date_match = re.search(r"Request Del\. Time\s*:\s*(\d{2}/\d{2}/\d{4})", full_text_page1)
    buyer_name_match = re.search(r"Buyer Name\s*:\s*([^\n]+)", full_text_page1)
    
    order_number = order_num_match.group(1).strip() if order_num_match else None
    delivery_date = delivery_date_match.group(1).strip() if delivery_date_match else None
    buyer_name = buyer_name_match.group(1).strip() if buyer_name_match else None

    # Lặp qua TẤT CẢ các trang
    for i, page in enumerate(pdf.pages):
        tables = page.extract_tables({"vertical_strategy": "lines", "horizontal_strategy": "lines"})
        if not tables: tables = page.extract_tables()
        
        if not tables: continue 
            
        item_table = tables[-1] 
        for row in item_table:
            # Logic lọc rác của 4PS
            if not row or len(row) < 6: continue 
            product_code = row[1]
            if product_code == "Product Code": continue
            if (row[2] or "").strip() == "Total": continue
            if not product_code or product_code.strip() == "": continue
            
            # 4PS dùng số chuẩn (dấu phẩy ngàn, chấm thập phân) -> Chỉ cần bỏ phẩy
            quantity_str = row[4].replace(',', '') if row[4] else '0'
            price_str = row[5].replace(',', '') if row[5] else '0'

            standard_item = {
                "Order_Number": order_number,    
                "Buyer_Name": buyer_name,      
                "Delivery_Date": delivery_date,
                "Item_Code": product_code,
                "Item_Name": row[2].replace('\n', ' '),
                "Quantity": quantity_str, # Sẽ convert sau
                "Price": price_str        # Sẽ convert sau
            }
            items_list.append(standard_item)
    
    return items_list

# ==========================================
# 3. HÀM BÓC TÁCH AVOLTA (REGEX SCAN STRATEGY)
# ==========================================
def parse_avolta_po(pdf):
    st.write("  > Nhận diện: Mẫu PO Avolta (SĐT 0903613502). Đang xử lý...")
    items_list = []

    # Lấy thông tin Header từ trang 1
    page1 = pdf.pages[0]
    page1_text = page1.extract_text() or ""
    
    order_num_match = re.search(r"PO No\.[\s\S]*?(\S+)", page1_text)
    order_number = order_num_match.group(1).strip() if order_num_match else "Unknown"
    
    delivery_date_match = re.search(r"Order Date\s*(\d{2}/\d{2}/\d{4})", page1_text)
    delivery_date = delivery_date_match.group(1).strip() if delivery_date_match else None
    
    buyer_name = "Unknown"
    if "Delivery Address" in page1_text:
        parts = page1_text.split("Delivery Address")
        if len(parts) > 1:
            lines = parts[1].strip().split('\n')
            buyer_name = " ".join(lines[:2]).strip()

    # Regex quét dòng: Bắt đầu bằng SỐ (Code) + Khoảng trắng + TEXT
    line_start_pattern = re.compile(r"^(\d+)\s+(.+)")

    # Lặp qua TẤT CẢ các trang
    for page in pdf.pages:
        text = page.extract_text()
        if not text: continue
        
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            
            # Bỏ qua các dòng tiêu đề/footer
            if "PO No" in line or "Page" in line or "Total" in line or "Item No" in line:
                continue

            match = line_start_pattern.match(line)
            if match:
                # Tìm tất cả các cụm "số" trong dòng
                potential_numbers = [
                    n for n in re.findall(r'[\d.,]+', line) 
                    if any(char.isdigit() for char in n)
                ]
                
                if len(potential_numbers) >= 3:
                    item_code = potential_numbers[0]
                    
                    qty_raw = potential_numbers[1]
                    
                    if len(potential_numbers) >= 4:
                        price_raw = potential_numbers[-2]
                    else:
                        price_raw = potential_numbers[-1]
                    
                    try:
                        start_index = line.find(item_code) + len(item_code)
                        end_index = line.find(qty_raw, start_index)
                        if end_index != -1:
                            item_name = line[start_index:end_index].strip()
                        else:
                            item_name = match.group(2)
                    except:
                        item_name = match.group(2)

                    items_list.append({
                        "Order_Number": order_number,    
                        "Buyer_Name": buyer_name,      
                        "Delivery_Date": delivery_date,
                        "Item_Code": item_code,
                        "Item_Name": item_name,
                        "Quantity": clean_avolta_number(qty_raw), # Dùng hàm mới
                        "Price": clean_avolta_number(price_raw)   # Dùng hàm mới
                    })

    return items_list

# ==========================================
# 4. HÀM TẠO EXCEL (HYBRID)
# ==========================================
def create_hybrid_excel(standard_df, unrecognized_files_list):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # Sheet 1: Dữ liệu chuẩn hóa (4PS + Avolta)
        if not standard_df.empty:
            standard_df.to_excel(writer, sheet_name="TongHop_DonHang", index=False)
        else:
            pd.DataFrame(["Không có dữ liệu chuẩn hóa."]).to_excel(writer, sheet_name="TongHop_DonHang", index=False, header=False)
        
        # Các Sheet khác: Dump text thô (Layout)
        if unrecognized_files_list:
            st.write("--- Đang xử lý các file khác (Dump Text giữ Layout) ---")
            for uploaded_file in unrecognized_files_list:
                safe_sheet_name = re.sub(r'[\\/*?:"<>|\[\]\s]', '_', uploaded_file.name.split('.')[0])[:30]
                try:
                    uploaded_file.seek(0)
                    with pdfplumber.open(uploaded_file) as pdf:
                        all_lines = []
                        for page in pdf.pages:
                            # keep_blank_chars=True giúp giữ khoảng cách, nhìn giống PDF hơn
                            text = page.extract_text(layout=True, keep_blank_chars=True)
                            if text: all_lines.extend(text.split('\n'))
                            all_lines.append("--- END PAGE ---")
                    
                    if all_lines:
                        pd.DataFrame(all_lines).to_excel(writer, sheet_name=safe_sheet_name, index=False, header=False)
                except Exception as e:
                    st.error(f"Lỗi dump file {uploaded_file.name}: {e}")

    return output.getvalue()

# ==========================================
# 5. GIAO DIỆN CHÍNH (STREAMLIT APP)
# ==========================================
st.set_page_config(page_title="Công cụ tổng hợp PO", layout="wide")
st.title("🚀 Công cụ trích xuất dữ liệu PO sang Excel")
st.markdown("""
**Hỗ trợ:**
1.  **4PS Corporation:** Tự động nhận diện bảng (xử lý nhiều trang).
2.  **Avolta (SĐT 0903613502):** Tự động nhận diện dòng (xử lý nhiều trang, số kiểu Âu).
3.  **Các file khác:** Tự động chuyển toàn bộ nội dung sang sheet riêng.
""")

uploaded_files = st.file_uploader("Tải file PDF lên:", type="pdf", accept_multiple_files=True)

if uploaded_files and st.button("Xử lý tất cả file"):
    all_standardized_data = []
    unrecognized_files = []
    
    progress_bar = st.progress(0)
    
    with st.expander("Chi tiết quá trình xử lý:", expanded=True):
        for i, uploaded_file in enumerate(uploaded_files):
            file_name = uploaded_file.name
            st.write(f"--- Đang mở: **{file_name}** ---")
            
            try:
                uploaded_file.seek(0)
                with pdfplumber.open(uploaded_file) as pdf:
                    if not pdf.pages:
                        st.error("File lỗi hoặc không có trang.")
                        continue
                    
                    # Lấy text trang 1 để nhận diện
                    page1_text = pdf.pages[0].extract_text() or ""
                    
                    items = []
                    is_recognized = False
                    customer_name = ""

                    # --- LOGIC NHẬN DIỆN ---
                    # 1. Check 4PS
                    if "4PS CORPORATION" in page1_text or "CÔNG TY TNHH MTV KITCHEN 4PS" in page1_text:
                        customer_name = "4PS"
                        items = parse_4ps_po(pdf)
                        is_recognized = True
                    
                    # 2. Check Avolta (Dựa vào SĐT)
                    elif "0903613502" in page1_text:
                        customer_name = "Avolta"
                        items = parse_avolta_po(pdf)
                        is_recognized = True
                    
                    # --- KẾT QUẢ ---
                    if is_recognized:
                        for item in items:
                            item['Customer'] = customer_name
                            item['File_Name'] = file_name
                            all_standardized_data.append(item)
                        st.success(f"  > Đã xử lý xong ({customer_name}). Lấy được {len(items)} dòng.")
                    else:
                        st.info("  > Không nhận diện được mẫu. Chuyển sang chế độ dump text.")
                        unrecognized_files.append(uploaded_file)

            except Exception as e:
                st.error(f"Lỗi khi xử lý file {file_name}: {e}")
            
            progress_bar.progress((i + 1) / len(uploaded_files))

    # TỔNG HỢP VÀ TẠO FILE EXCEL
    df_standard = pd.DataFrame(all_standardized_data)
    
    if not df_standard.empty:
        # Convert số lượng/đơn giá sang số (cho 4PS, vì Avolta đã convert trong hàm parse rồi)
        try:
            # Lưu ý: Avolta đã float sẵn, 4PS đang là str -> convert lại để chắc chắn
            df_standard['Quantity'] = pd.to_numeric(df_standard['Quantity'], errors='coerce').fillna(0)
            df_standard['Price'] = pd.to_numeric(df_standard['Price'], errors='coerce').fillna(0)
        except: pass
        
        # Sắp xếp cột
        cols = ['Customer', 'Order_Number', 'Buyer_Name', 'Delivery_Date', 'Item_Code', 'Item_Name', 'Quantity', 'Price', 'File_Name']
        final_cols = [c for c in cols if c in df_standard.columns]
        df_standard = df_standard[final_cols]
        
        st.success(f"🎉 Hoàn tất! Tổng hợp được {len(df_standard)} dòng dữ liệu chuẩn hóa.")
        st.dataframe(df_standard)
    else:
        st.warning("Chưa tìm thấy dữ liệu chuẩn hóa nào (4PS/Avolta).")

    # Tạo file Excel hybrid
    excel_data = create_hybrid_excel(df_standard, unrecognized_files)
    
    st.download_button(
        label="📥 Tải file Excel kết quả",
        data=excel_data,
        file_name="TongHop_PO_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
