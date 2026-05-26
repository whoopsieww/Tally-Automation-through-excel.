import os
import fitz  # PyMuPDF
import easyocr
import xlwings as xw  # Bridge to active Excel

def inject_pdf_to_active_excel(pdf_path):
    print("🤖 Analyzing PDF text layers...")
    doc = fitz.open(pdf_path)
    all_rows = []
    is_digital = False
    
    # Test if PDF has digital text layers
    for page in doc:
        if page.get_text().strip():
            is_digital = True
            break
            
    if is_digital:
        print("⚡ Clean Digital PDF detected! Preparing live extraction...")
        for page in doc:
            blocks = page.get_text("blocks")
            blocks.sort(key=lambda b: b[1])
            for b in blocks:
                text_line = b[4].strip()
                if text_line:
                    row_data = [item.strip() for item in text_line.split('\n') if item.strip()]
                    all_rows.append(row_data)
    else:
        print("👁️ Scanned PDF image detected! Launching OCR Engine...")
        reader = easyocr.Reader(['en'])
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(dpi=300)
            img_path = f"page_{page_num}.png"
            pix.save(img_path)
            
            results = reader.readtext(img_path, detail=1)
            results.sort(key=lambda x: x[0][0][1])
            
            current_row = []
            last_y = -1
            row_threshold = 15
            
            for bbox, text, prob in results:
                y_coord = bbox[0][1]
                if last_y != -1 and (y_coord - last_y) > row_threshold:
                    if current_row:
                        all_rows.append(current_row)
                    current_row = []
                current_row.append(text.strip())
                last_y = y_coord
            if current_row:
                all_rows.append(current_row)
            os.remove(img_path)
            
    # --- CRITICAL FIX: MANAGE MULTI-LENGTH ROWS FOR XLWINGS ---
    if not all_rows:
        print("⚠️ No data extracted from the PDF.")
        return

    # Find the length of the row that has the maximum number of items
    max_len = max(len(r) for r in all_rows)
    
    # Pad any shorter rows with empty strings to make a perfect 2D grid matrix
    clean_matrix = []
    for r in all_rows:
        padded_row = r + [""] * (max_len - len(r))
        clean_matrix.append(padded_row)
            
    # --- LIVE INJECTION SECTION ---
    print("🔌 Connecting to your active Excel window...")
    try:
        wb = xw.books.active
        ws = wb.sheets.active
        
        print("🧹 Clearing previous data in the active sheet...")
        ws.clear_contents()
        
        print("✍️ Injecting fresh row data live onto your screen...")
        # Paste the safe balanced matrix starting from cell A1
        ws.range("A1").value = clean_matrix
        print("✅ Live Injection Complete! Check your open Excel sheet.")
        
    except Exception as e:
        print(f"❌ Error during active data transfer: {e}")

# --- PATH CONFIGURATION ---
project_folder = r"C:\Users\91829\Desktop\ai folder"
pdf_source = os.path.join(project_folder, "test.pdf")

inject_pdf_to_active_excel(pdf_source)