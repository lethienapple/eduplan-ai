import io
import json
import os
import re
import streamlit as st
from google import genai
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

st.set_page_config(
    page_title="EduPlan AI - Soạn Giáo Án & Đề Kiểm Tra",
    layout="wide",
    page_icon="🎓"
)

# CSS bảo mật: Ẩn vĩnh viễn biểu tượng con mắt trên mọi trình duyệt và Streamlit
st.markdown(
    """
    <style>
    button[aria-label="Show password text"],
    button[aria-label="Hide password text"] {
        display: none !important;
    }
    input::-ms-reveal,
    input::-ms-clear {
        display: none !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

CONFIG_FILE = "api_config.json"

# Quản lý lưu trữ API Key cục bộ
def load_saved_key():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("api_key", "")
        except Exception:
            return ""
    return ""

def save_key_to_file(key_to_save):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump({"api_key": key_to_save.strip()}, f)

def delete_saved_key():
    if os.path.exists(CONFIG_FILE):
        os.remove(CONFIG_FILE)

# Khởi tạo trạng thái phiên làm việc
if "api_key" not in st.session_state:
    st.session_state.api_key = load_saved_key()

if "is_generating" not in st.session_state:
    st.session_state.is_generating = False

st.title("🎓 Trợ Lý AI: Xuất Bản KHBD, Slide & Đề Kiểm Tra")
st.caption("Công cụ chuẩn hóa giáo án, kịch bản trình chiếu và ma trận đề kiểm tra theo Công văn 7991")

# Thanh điều khiển bên trái (Sidebar)
with st.sidebar:
    st.header("🔑 Cấu hình API Gemini")
    
    with st.expander("📖 Hướng dẫn lấy API Key miễn phí (3 bước)"):
        st.markdown("""
        **Bước 1:** Truy cập trang [Google AI Studio](https://aistudio.google.com/).
        
        **Bước 2:** Đăng nhập tài khoản Gmail cá nhân (`@gmail.com`).
        
        **Bước 3:** Nhấn nút **"Get API key"** (ở góc trái) -> Nhấn **"Create API key"** -> Sao chép chuỗi mã và dán vào ô bên dưới rồi bấm **"Lưu Key"**.
        
        *Lưu ý: Khóa API hoàn toàn miễn phí từ Google và được lưu an toàn trên máy tính của thầy/cô.*
        """)

    if st.session_state.api_key:
        masked_key = st.session_state.api_key[:6] + "••••••••••••••••" + st.session_state.api_key[-4:]
        st.text_input("Trạng thái Key:", value=masked_key, disabled=True)
        st.success("🔒 Key đã lưu an toàn cục bộ")
        
        if st.button("🔄 Đổi / Xóa Key", use_container_width=True, disabled=st.session_state.is_generating):
            delete_saved_key()
            st.session_state.api_key = ""
            st.rerun()
    else:
        input_key = st.text_input(
            "Nhập khóa API:",
            type="password",
            placeholder="Dán API Key tại đây..."
        )
        if st.button("💾 Lưu Key", use_container_width=True, type="primary"):
            if input_key.strip():
                save_key_to_file(input_key)
                st.session_state.api_key = input_key.strip()
                st.toast("✅ Đã lưu Key thành công!")
                st.rerun()
            else:
                st.warning("Vui lòng nhập Key trước khi bấm lưu.")

    st.markdown("---")
    st.header("⚙️ Thiết lập bài dạy")
    mon_hoc = st.text_input("Môn học:", value="Địa lí", disabled=st.session_state.is_generating)
    khoi_lop = st.selectbox("Khối lớp:", [f"Lớp {i}" for i in range(1, 13)], index=6, disabled=st.session_state.is_generating)
    ten_bai = st.text_input("Tên bài dạy / Chủ đề kiểm tra:", value="Châu Âu và Châu Á", disabled=st.session_state.is_generating)
    thoi_luong = st.text_input("Thời lượng:", value="2 tiết", disabled=st.session_state.is_generating)
    loai_sp = st.radio(
        "Chọn sản phẩm cần tạo:",
        ["Ma trận & Đề kiểm tra", "Kế hoạch bài dạy (KHBD)", "Kịch bản Slide bài giảng"],
        disabled=st.session_state.is_generating
    )

st.subheader("🎯 Yêu cầu cần đạt (YCCĐ) hoặc Nội dung cốt lõi")
yccd_input = st.text_area(
    "Dán YCCĐ trích từ chương trình môn học hoặc nội dung trọng tâm:",
    height=140,
    placeholder="Nhập các yêu cầu cần đạt, mức độ tư duy (Nhận biết, Thông hiểu, Vận dụng)...",
    disabled=st.session_state.is_generating
)

def set_cell_background(cell, fill_hex):
    """Tô màu nền nhẹ cho hàng tiêu đề của bảng"""
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

# Hàm chuyển đổi Markdown sang tệp DOCX với định dạng bảng rộng nằm ngang
def create_docx(markdown_text, title):
    doc = Document()
    
    # Thiết lập trang in nằm ngang (Landscape A4) và lề hẹp
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.6)
    section.left_margin = Inches(0.6)
    section.right_margin = Inches(0.6)

    # Tiêu đề tài liệu
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_p.add_run(title.upper())
    run_title.font.name = "Times New Roman"
    run_title.font.size = Pt(15)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0, 51, 153)

    lines = markdown_text.split('\n')
    in_table = False
    table_data = []

    def flush_table(t_data):
        if not t_data:
            return
        valid_rows = [row for row in t_data if not all(re.match(r'^:?-+:?$', c.strip()) or c.strip() == '' for c in row if c)]
        if not valid_rows:
            return
        
        max_cols = max(len(r) for r in valid_rows)
        if max_cols == 0:
            return

        table = doc.add_table(rows=len(valid_rows), cols=max_cols)
        table.style = 'Table Grid'
        table.autofit = True

        for r_idx, row in enumerate(valid_rows):
            is_header = (r_idx == 0)
            for c_idx, cell_text in enumerate(row):
                if c_idx < max_cols:
                    cell = table.cell(r_idx, c_idx)
                    clean_cell = cell_text.replace("<br>", "\n").replace("<br/>", "\n")
                    clean_cell = clean_cell.replace("; ;", "\n").replace("; -", "\n-")
                    
                    p = cell.paragraphs[0]
                    p.text = ""
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.line_spacing = 1.05
                    
                    if is_header or c_idx == 0 or any(unit in clean_cell for unit in ['đ', 'điểm', '%']):
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    if is_header:
                        set_cell_background(cell, "F2F2F2")

                    parts = re.split(r'(\*\*.*?\*\*)', clean_cell)
                    for pt in parts:
                        if pt.startswith("**") and pt.endswith("**"):
                            run = p.add_run(pt[2:-2])
                            run.bold = True
                        else:
                            run = p.add_run(pt)
                        
                        run.font.name = "Times New Roman"
                        run.font.size = Pt(9.5)
                        if is_header:
                            run.font.bold = True

    idx = 0
    while idx < len(lines):
        line = lines[idx].strip()
        
        if line.startswith("|") and line.endswith("|"):
            in_table = True
            raw_cols = [c.strip() for c in line.strip("|").split("|")]
            if not all(re.match(r'^:?-+:?$', c) for c in raw_cols if c):
                table_data.append(raw_cols)
            idx += 1
            continue
        else:
            if in_table:
                flush_table(table_data)
                table_data = []
                in_table = False

        if not line:
            idx += 1
            continue

        if line.startswith("# "):
            p = doc.add_heading(level=1)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(line[2:].strip())
            run.font.name = "Times New Roman"
            run.font.size = Pt(13)
            run.font.bold = True
        elif line.startswith("## "):
            p = doc.add_heading(level=2)
            run = p.add_run(line[3:].strip())
            run.font.name = "Times New Roman"
            run.font.size = Pt(12)
            run.font.bold = True
        elif line.startswith("### "):
            p = doc.add_heading(level=3)
            run = p.add_run(line[4:].strip())
            run.font.name = "Times New Roman"
            run.font.size = Pt(11)
            run.font.bold = True
        elif line.startswith("- ") or line.startswith("* "):
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            parts = re.split(r'(\*\*.*?\*\*)', line[2:])
            for pt in parts:
                run = p.add_run(pt[2:-2] if pt.startswith("**") and pt.endswith("**") else pt)
                run.font.name = "Times New Roman"
                run.font.size = Pt(11)
                if pt.startswith("**") and pt.endswith("**"):
                    run.bold = True
        else:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            parts = re.split(r'(\*\*.*?\*\*)', line)
            for pt in parts:
                run = p.add_run(pt[2:-2] if pt.startswith("**") and pt.endswith("**") else pt)
                run.font.name = "Times New Roman"
                run.font.size = Pt(11)
                if pt.startswith("**") and pt.endswith("**"):
                    run.bold = True
                    
        idx += 1

    if in_table:
        flush_table(table_data)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def trigger_generation():
    if not st.session_state.api_key:
        st.error("⚠️ Vui lòng nhập và lưu API Key ở cột bên trái trước.")
        return
    if not ten_bai.strip():
        st.warning("⚠️ Vui lòng nhập tên bài học trước khi tạo.")
        return
    st.session_state.is_generating = True

button_label = "⏳ Đang tiến hành tạo, vui lòng chờ..." if st.session_state.is_generating else "🚀 Bắt đầu tạo sản phẩm"
st.button(
    button_label,
    type="primary",
    disabled=st.session_state.is_generating,
    on_click=trigger_generation
)

# Xử lý sinh nội dung bằng Gemini API
if st.session_state.is_generating:
    prompt = f"""
    Vai trò: Chuyên gia khảo thí và thẩm định phương pháp dạy học theo định hướng Công văn 7991.
    Yêu cầu: Hãy biên soạn '{loai_sp}' cho bài học sau:
    - Môn học: {mon_hoc}
    - Khối lớp: {khoi_lop}
    - Tên nội dung / Chủ đề: {ten_bai}
    - Thời lượng: {thoi_luong}
    - Yêu cầu cần đạt đầu vào:
    {yccd_input if yccd_input.strip() else 'Tự động bám sát chương trình GDPT 2018 hiện hành'}
    
    QUY TẮC BẮT BUỘC VỀ BIỂU ĐIỂM VÀ ĐÁP ÁN:
    1. Trong phần HƯỚNG DẪN CHẤM VÀ BIỂU ĐIỂM PHẦN TỰ LUẬN:
       - Phải lập dạng BẢNG gồm 3 cột: | Câu | Nội dung / Hướng dẫn chấm chi tiết | Điểm |
       - MỖI Ý NHỎ HOẶC MỖI BƯỚC ĐÁP ÁN PHẢI TÁCH THÀNH MỘT DÒNG RIÊNG (Ví dụ: Câu 1a - ý 1; Câu 1a - ý 2; Câu 1b...).
       - TUYỆT ĐỐI KHÔNG gom nhiều ý nhỏ vào một ô rồi dồn dãy điểm (như 0,5 đ; 0,25 đ; ...) sang một cột.
       - Cột 'Điểm' chỉ ghi số điểm chính xác của ĐÚNG Ý ĐÓ (ví dụ: 0,5 đ hoặc 0,25 đ).
    2. Trong tất cả các bảng Markdown:
       - TUYỆT ĐỐI KHÔNG xuất thẻ HTML như <br>, <br/>, <p>.
    
    CẤU TRÚC CHI TIẾT:
    - Nếu là 'Ma trận & Đề kiểm tra':
      + Khung ma trận 4 mức độ tư duy (Nhận biết, Thông hiểu, Vận dụng, Vận dụng cao), phân chia TNKQ và Tự luận theo dạng bảng Markdown rõ ràng.
      + Bảng đặc tả chi tiết dạng bảng Markdown.
      + Đề kiểm tra chuẩn (TNKQ & Tự luận).
      + Đáp án TNKQ dạng bảng (Câu - Đáp án).
      + Hướng dẫn chấm Tự luận chi tiết chuẩn biểu điểm như quy tắc trên.
    - Nếu là 'Kế hoạch bài dạy (KHBD)': Mục tiêu -> Thiết bị dạy học -> 4 Hoạt động dạy học chuẩn 4 bước (Chuyển giao, Thực hiện, Báo cáo, Kết luận).
    - Nếu là 'Kịch bản Slide bài giảng': Tiêu đề, Nội dung trực quan, Hoạt động tương tác, Lời thoại của GV.
    """

    status_box = st.status("🚀 Đang khởi động mô hình AI...", expanded=True)
    try:
        status_box.write("📡 Kết nối tới Gemini API...")
        client = genai.Client(api_key=st.session_state.api_key)
        
        status_box.write("✍️ Đang phân tích dữ liệu và viết nội dung thời gian thực...")
        
        st.markdown("### 📋 Kết quả biên soạn trực tiếp:")
        response_container = st.empty()
        full_text = ""

        response_stream = client.models.generate_content_stream(
            model='gemini-3.8-flash',
            contents=prompt,
        )

        for chunk in response_stream:
            if chunk.text:
                full_text += chunk.text
                display_text = re.sub(r'<br\s*/?>', '  \n', full_text)
                response_container.markdown(display_text, unsafe_allow_html=True)

        status_box.update(label="🎉 Hoàn thành xuất bản nội dung!", state="complete", expanded=False)
        st.success("Đã hoàn tất toàn bộ nội dung giáo án / đề kiểm tra!")

        st.session_state["saved_text"] = full_text
        st.session_state["saved_title"] = f"{ten_bai}_{loai_sp.replace(' ', '_')}"

    except Exception as e:
        status_box.update(label="❌ Có lỗi xảy ra trong quá trình xử lý", state="error", expanded=True)
        st.error(f"Chi tiết lỗi: {str(e)}")
    finally:
        st.session_state.is_generating = False

# Nút xuất và tải file Word
if "saved_text" in st.session_state and st.session_state["saved_text"]:
    docx_file = create_docx(st.session_state["saved_text"], st.session_state["saved_title"])
    st.download_button(
        label="📥 Tải về file Word (.docx) hoàn chỉnh",
        data=docx_file,
        file_name=f"{st.session_state['saved_title']}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )