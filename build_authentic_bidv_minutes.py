#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_authentic_bidv_minutes.py
Triển khai CHÍNH XÁC 100% từng quy định trong skill bidv-meeting-minutes:
1. KHÔNG có bảng biểu ma trận nhiệm vụ trong phần thân văn bản (theo mục Gotchas của skill).
2. Đúng cấu trúc 4 phần chuẩn mực:
   - Dẫn nhập (2 đoạn văn mở đầu)
   - 1. BỐI CẢNH CUỘC HỌP (đúng 3 đoạn văn xuôi, đoạn 3 có đánh số (i)... (v)...)
   - 2. NỘI DUNG THỐNG NHẤT VÀ CHỈ ĐẠO TRIỂN KHAI:
     + 2.1. Định hướng và chỉ đạo chung (5 luận điểm nguyên tắc mở đầu bằng cụm động từ mạnh in đậm có dấu hai chấm)
     + 2.2. Nội dung triển khai cụ thể: nội dung - đơn vị thực hiện - thời hạn (8 nhóm nhiệm vụ '1. Về...:', mỗi nhóm đúng 3 gạch đầu dòng: - Nội dung..., - Đơn vị..., - Thời hạn hoàn thành in nghiêng)
   - Đoạn kết văn bản chuẩn BIDV
3. Header Table (Quốc hiệu & Tên ngân hàng) và Footer Table (Nơi nhận & Chữ ký) borderless.
4. Căn lề chuẩn: Top 0.75", Bottom 0.75", Left 1.0", Right 0.75", khổ giấy US Letter (8.5 x 11 in).
5. Phông chữ duy nhất: Times New Roman 13pt, màu đen RGB(0,0,0).
"""

import os
import subprocess
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor

OUTPUT_DOCX = "/home/hieu/Downloads/Bien_ban_hop_song_phuong_BIDV_Tencent_chuan_the_thuc_BIDV.docx"
OUTPUT_PDF = "/home/hieu/Downloads/Bien_ban_hop_song_phuong_BIDV_Tencent_chuan_the_thuc_BIDV.pdf"

def set_cell_borderless(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        r'<w:tcBorders {}>'
        r'<w:top w:val="none"/><w:left w:val="none"/>'
        r'<w:bottom w:val="none"/><w:right w:val="none"/>'
        r'</w:tcBorders>'.format(nsdecls('w'))
    )
    tcPr.append(tcBorders)

def set_run_font(run, size_pt=13.0, bold=False, italic=False):
    run.font.name = "Times New Roman"
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    rPr = run._r.get_or_add_rPr()
    rFonts = parse_xml(
        r'<w:rFonts {} w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'.format(
            nsdecls("w")
        )
    )
    rPr.append(rFonts)

doc = docx.Document()

# 1. Page margins & size (Letter 8.5 x 11.0 in, Top/Bottom 0.75", Left 1.0", Right 0.75")
for s in doc.sections:
    s.page_width = Inches(8.5)
    s.page_height = Inches(11.0)
    s.top_margin = Inches(0.75)
    s.bottom_margin = Inches(0.75)
    s.left_margin = Inches(1.0)
    s.right_margin = Inches(0.75)

# 2. Header Table (1 dòng, 2 cột không đường viền)
header_table = doc.add_table(rows=1, cols=2)
header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
header_table.autofit = False

c_left = header_table.cell(0, 0)
c_right = header_table.cell(0, 1)
c_left.width = Inches(3.25)
c_right.width = Inches(3.50)
set_cell_borderless(c_left)
set_cell_borderless(c_right)

# Left Header
p_lh = c_left.paragraphs[0]
p_lh.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_lh.paragraph_format.line_spacing = 1.15
p_lh.paragraph_format.space_after = Pt(2)
r = p_lh.add_run("NGÂN HÀNG TMCP ĐẦU TƯ\nVÀ PHÁT TRIỂN VIỆT NAM")
set_run_font(r, size_pt=12.0, bold=True)

p_lh_line = c_left.add_paragraph()
p_lh_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_lh_line.paragraph_format.space_after = Pt(2)
r = p_lh_line.add_run("————————")
set_run_font(r, size_pt=11.0)

p_lh_so = c_left.add_paragraph()
p_lh_so.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_lh_so.paragraph_format.space_after = Pt(0)
r = p_lh_so.add_run("Số:   2710   /BB-BIDV")
set_run_font(r, size_pt=12.0)

# Right Header
p_rh = c_right.paragraphs[0]
p_rh.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_rh.paragraph_format.line_spacing = 1.15
p_rh.paragraph_format.space_after = Pt(2)
r = p_rh.add_run("CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM")
set_run_font(r, size_pt=12.0, bold=True)

p_rh_tn = c_right.add_paragraph()
p_rh_tn.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_rh_tn.paragraph_format.space_after = Pt(2)
r = p_rh_tn.add_run("Độc lập - Tự do - Hạnh phúc")
set_run_font(r, size_pt=12.5, bold=True)

p_rh_line = c_right.add_paragraph()
p_rh_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_rh_line.paragraph_format.space_after = Pt(4)
r = p_rh_line.add_run("————————————")
set_run_font(r, size_pt=11.0)

p_rh_date = c_right.add_paragraph()
p_rh_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_rh_date.paragraph_format.space_after = Pt(0)
r = p_rh_date.add_run("Thâm Quyến, ngày 27 tháng 10 năm 2026")
set_run_font(r, size_pt=12.0, italic=True)

# Spacing
p_sp = doc.add_paragraph()
p_sp.paragraph_format.space_before = Pt(8)
p_sp.paragraph_format.space_after = Pt(0)

# 3. Title Block
p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_title.paragraph_format.space_before = Pt(16)
p_title.paragraph_format.space_after = Pt(6)
r = p_title.add_run("BIÊN BẢN CUỘC HỌP")
set_run_font(r, size_pt=15.0, bold=True)

p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_sub.paragraph_format.space_after = Pt(4)
r = p_sub.add_run("Phiên họp Cấp cao Song phương giữa Ban Lãnh đạo Ngân hàng TMCP Đầu tư và\nPhát triển Việt Nam (BIDV) và Ban Lãnh đạo Tập đoàn Tencent (Trung Quốc)\nvề Hợp tác Chiến lược Toàn diện trong lĩnh vực Chuyển đổi số, AI, Cloud và Thanh toán")
set_run_font(r, size_pt=13.5, bold=True)

p_tline = doc.add_paragraph()
p_tline.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_tline.paragraph_format.space_after = Pt(14)
r = p_tline.add_run("————————————")
set_run_font(r, size_pt=12.0)

# 4. Phần dẫn nhập (2 đoạn văn mở đầu chuẩn BIDV)
p_intro1 = doc.add_paragraph()
p_intro1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_intro1.paragraph_format.line_spacing = 1.15
p_intro1.paragraph_format.first_line_indent = Inches(0.5)
p_intro1.paragraph_format.space_after = Pt(6)

r = p_intro1.add_run("Ngày 27/10/2026, từ 08h30 đến 11h30, tại Phòng họp Cấp cao Trụ sở chính Tập đoàn Tencent (Tencent Binhai Mansion, TP. Thâm Quyến, Trung Quốc), đồng chí ")
set_run_font(r)
r = p_intro1.add_run("Hoàng Việt Hùng – Tổng Giám đốc BIDV")
set_run_font(r, bold=True)
r = p_intro1.add_run(" và ông ")
set_run_font(r)
r = p_intro1.add_run("Hồng Đan Nghị (Danny Hong) – Phó Chủ tịch Phụ trách Công nghệ Tài chính (Tencent FinTech), Tập đoàn Tencent")
set_run_font(r, bold=True)
r = p_intro1.add_run(" đã đồng chủ trì Phiên họp cấp cao song phương về thúc đẩy hợp tác chiến lược toàn diện giữa BIDV và Tập đoàn Tencent. Tham dự cuộc họp có đại diện Ban Lãnh đạo và cán bộ chuyên trách hai bên. Về phía BIDV có 05 Phó Tổng Giám đốc: bà Nguyễn Thị Quỳnh Giao (phụ trách Kế hoạch và Tài chính – KHTC), ông Phan Thanh Hải (phụ trách Hệ thống Công nghệ thông tin – CNTT), ông Lại Tiến Quân (phụ trách Khách hàng Doanh nghiệp – KHDN), ông Vũ Hoàng Dương (phụ trách Trí tuệ nhân tạo và Đổi mới sáng tạo – AI & ĐMST), ông Lê Huy Hoàng (phụ trách Ngân hàng Bán lẻ – NHBL và Nhân sự); ông Nguyễn Hữu Tiến – Phó Trưởng khối KHDN; ông Ngô Minh Hoàng – Phó Giám đốc Văn phòng BIDV (VP); cùng 20 Lãnh đạo các Ban, Trung tâm tại Trụ sở chính (TSC) và các Chi nhánh trọng điểm. Về phía Tập đoàn Tencent có ông Dương Văn Hối (James Yang) – Phó Tổng Giám đốc phụ trách Phát triển Kinh doanh Toàn cầu kiêm Người đại diện theo pháp luật của Tenpay Global; ông Ngô Tăng Vỹ (Sam Wu) – Giám đốc phụ trách Thị trường Việt Nam, Tencent Cloud; ông Dư Duyên Tân (Yu Yanxin) – Giám đốc Bộ phận Tư vấn Chuyển đổi số và AI; ông Bùi Trạch Lương (Pei Zeliang) – Trưởng Bộ phận Kiến trúc sư Kỹ thuật; ông Kiến Thắng (Jiansheng) – Chuyên gia cao cấp Giải pháp Ngân hàng thông minh; bà Nghê Nhã (Ni Ya) – Giám đốc Phát triển Kinh doanh Thanh toán Nội địa Tencent; ông David – Phụ trách Thị trường Việt Nam, WeChat Pay Quốc tế; cùng đội ngũ chuyên gia công nghệ, pháp chế và đối ngoại của Tập đoàn Tencent.")
set_run_font(r)

p_intro2 = doc.add_paragraph()
p_intro2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_intro2.paragraph_format.line_spacing = 1.15
p_intro2.paragraph_format.first_line_indent = Inches(0.5)
p_intro2.paragraph_format.space_after = Pt(12)

r = p_intro2.add_run("Sau khi nghe đại diện Tập đoàn Tencent trình bày báo cáo tổng quan về năng lực công nghệ, hệ sinh thái Tencent, các giải pháp AI Platform, Financial Cloud, định danh sinh trắc học eKYC và thanh toán số; ý kiến trao đổi của các đồng chí Phó Tổng Giám đốc BIDV về nhu cầu chuyển đổi AI, hiện đại hóa CNTT, mở rộng mạng lưới thanh toán hai chiều và cơ hội hợp tác đầu tư tài chính chiến lược; trên cơ sở kết quả thảo luận song phương thẳng thắn, xây dựng và Lễ ký kết Biên bản ghi nhớ hợp tác (MOU) thanh toán xuyên biên giới giữa BIDV và Tenpay Global, ")
set_run_font(r)
r = p_intro2.add_run("Tổng Giám đốc BIDV Hoàng Việt Hùng")
set_run_font(r, bold=True)
r = p_intro2.add_run(" và ")
set_run_font(r)
r = p_intro2.add_run("Phó Chủ tịch Tencent FinTech Danny Hong")
set_run_font(r, bold=True)
r = p_intro2.add_run(" thống nhất kết luận và phân công triển khai như sau:")
set_run_font(r)

# 5. Mục 1. BỐI CẢNH CUỘC HỌP (đúng 3 đoạn văn hoàn chỉnh)
p_h1 = doc.add_paragraph()
p_h1.alignment = WD_ALIGN_PARAGRAPH.LEFT
p_h1.paragraph_format.space_before = Pt(10)
p_h1.paragraph_format.space_after = Pt(6)
p_h1.paragraph_format.keep_with_next = True
r = p_h1.add_run("1. BỐI CẢNH CUỘC HỌP")
set_run_font(r, bold=True)

# Đoạn 1 Mục 1
p_bc1 = doc.add_paragraph()
p_bc1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_bc1.paragraph_format.line_spacing = 1.15
p_bc1.paragraph_format.first_line_indent = Inches(0.5)
p_bc1.paragraph_format.space_after = Pt(6)
r = p_bc1.add_run("Quan hệ đối tác hợp tác chiến lược toàn diện giữa Việt Nam và Trung Quốc đang trên đà phát triển rất tích cực với kim ngạch thương mại song phương đạt xấp xỉ 200 tỷ USD/năm, mở ra không gian giao thương, đầu tư và du lịch hết sức rộng lớn giữa hai nền kinh tế. Trong bối cảnh ngành tài chính – ngân hàng bước vào làn sóng chuyển đổi sâu sắc dưới tác động của Trí tuệ nhân tạo (AI), Điện toán đám mây (Cloud) và Dữ liệu lớn (Big Data), Đảng ủy, Hội đồng Quản trị và Ban Điều hành BIDV đã xác định chuyển đổi số và phát triển khoa học công nghệ là một trong ba trụ cột chiến lược đột phá nhằm bảo vệ và khẳng định vững chắc vị thế định chế tài chính Top 1 tại thị trường Việt Nam cả về quy mô tổng tài sản, dư nợ tín dụng, nguồn vốn huy động và tệp khách hàng.")
set_run_font(r)

# Đoạn 2 Mục 1
p_bc2 = doc.add_paragraph()
p_bc2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_bc2.paragraph_format.line_spacing = 1.15
p_bc2.paragraph_format.first_line_indent = Inches(0.5)
p_bc2.paragraph_format.space_after = Pt(6)
r = p_bc2.add_run("Tiếp nối kết quả các cuộc tiếp xúc cấp cao giữa Lãnh đạo hai bên tại Bắc Kinh năm 2025 và chuyến thăm làm việc của Lãnh đạo Tenpay Global tại Hà Nội đầu năm 2026, BIDV và Tencent đã phối hợp triển khai thành công dịch vụ chấp nhận thanh toán mã QR xuyên biên giới cho du khách Trung Quốc sử dụng WeChat Pay tại mạng lưới Đơn vị chấp nhận thanh toán (Merchant) của BIDV từ tháng 8/2026, ghi nhận sản lượng giao dịch tăng trưởng rất ấn tượng. Chuyến thăm chính thức Trụ sở chính Tập đoàn Tencent của Đoàn đại biểu cấp cao BIDV lần này nhằm mục tiêu đánh giá toàn diện tiềm năng công nghệ của đối tác, nâng tầm quan hệ từ hợp tác kỹ thuật thanh toán cục bộ sang Hợp tác Đối tác Chiến lược Toàn diện, lâu dài trên mọi phương diện công nghệ và tài chính.")
set_run_font(r)

# Đoạn 3 Mục 1 (chứa (i)... (v)...)
p_bc3 = doc.add_paragraph()
p_bc3.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_bc3.paragraph_format.line_spacing = 1.15
p_bc3.paragraph_format.first_line_indent = Inches(0.5)
p_bc3.paragraph_format.space_after = Pt(12)
r = p_bc3.add_run("Qua rà soát thực trạng tại BIDV và đối chiếu với năng lực công nghệ vượt trội của Tencent (tập đoàn sở hữu hệ sinh thái 1,4 tỷ người dùng WeChat, mô hình nền tảng Hunyuan LLM và năng lực Financial Cloud đứng đầu khu vực), hai bên đã làm rõ các điểm nghẽn và bài toán chiến lược cần tập trung giải quyết gồm: ")
set_run_font(r)
r = p_bc3.add_run("(i) Nhu cầu xây dựng Đề án Khung kiến trúc nền tảng AI Platform đồng bộ và Khung quản trị AI Governance tuân thủ chặt chẽ Luật AI, quy định an toàn thông tin và bảo vệ dữ liệu cá nhân tại Việt Nam; (ii) Hiện đại hóa hạ tầng công nghệ thông tin, ứng dụng giải pháp AIOps để vận hành tự động và nâng cao độ sẵn sàng 99.99% của hệ thống ngân hàng; (iii) Nâng cấp năng lực định danh điện tử Financial eKYC và phòng chống gian lận tài chính ứng dụng sinh trắc học tiên tiến chống giả mạo sâu (Deepfake); (iv) Thúc đẩy hoàn thiện giải pháp thanh toán hai chiều, trọng tâm là mở chiều cho khách hàng Việt Nam thanh toán thuận tiện tại Trung Quốc (Inbound China Payment) trước thềm Hội nghị Cấp cao APEC 2026; và (v) Nghiên cứu cơ chế hợp tác đầu tư tài chính chiến lược gắn với chuyển giao năng lực công nghệ dài hạn.")
set_run_font(r)

# 6. Mục 2. NỘI DUNG THỐNG NHẤT VÀ CHỈ ĐẠO TRIỂN KHAI
p_h2 = doc.add_paragraph()
p_h2.alignment = WD_ALIGN_PARAGRAPH.LEFT
p_h2.paragraph_format.space_before = Pt(10)
p_h2.paragraph_format.space_after = Pt(6)
p_h2.paragraph_format.keep_with_next = True
r = p_h2.add_run("2. NỘI DUNG THỐNG NHẤT VÀ CHỈ ĐẠO TRIỂN KHAI")
set_run_font(r, bold=True)

# 6.1. Mục 2.1. Định hướng và chỉ đạo chung (5 luận điểm nguyên tắc mở đầu bằng cụm động từ mạnh)
p_h21 = doc.add_paragraph()
p_h21.alignment = WD_ALIGN_PARAGRAPH.LEFT
p_h21.paragraph_format.space_before = Pt(6)
p_h21.paragraph_format.space_after = Pt(6)
p_h21.paragraph_format.keep_with_next = True
r = p_h21.add_run("2.1. Định hướng và chỉ đạo chung")
set_run_font(r, bold=True)

principles = [
    ("1. Chuyển đổi công nghệ gắn liền với hiệu quả thực chất: ",
     "Quá trình ứng dụng Trí tuệ nhân tạo và Điện toán đám mây tại BIDV phải bám sát định hướng kinh doanh, bắt đầu từ các bài toán cụ thể có khả năng đo lường ROI (tỷ suất hoàn vốn đầu tư) rõ ràng, giải quyết trực tiếp bài toán nâng cao năng suất lao động của cán bộ, tối ưu hóa chi phí vận hành và tăng cường trải nghiệm người dùng cuối."),
    
    ("2. Đặt an toàn hệ thống, bảo mật dữ liệu và tuân thủ pháp luật lên hàng đầu: ",
     "Mọi giải pháp hợp tác công nghệ, chuyển giao thuật toán và kết nối kỹ thuật với Tencent phải bảo đảm tuân thủ tuyệt đối quy định của Ngân hàng Nhà nước Việt Nam, Luật Giao dịch điện tử, Nghị định về bảo vệ dữ liệu cá nhân và các quy định pháp lý liên quan về an ninh mạng; tuyệt đối không để rò rỉ dữ liệu khách hàng hoặc phương hại đến an toàn hệ thống ngân hàng."),
    
    ("3. Tăng tốc hoàn thiện kiến trúc AI Platform và khung quản trị AI (AI Governance): ",
     "Xây dựng hạ tầng AI theo hướng tập trung, dùng chung và linh hoạt, tránh đầu tư dàn trải, cục bộ giữa các đơn vị; thiết lập cơ chế kiểm soát rủi ro thuật toán, tính minh bạch và độ chính xác của các mô hình Multi-Agent ngay từ giai đoạn thiết kế ban đầu."),
    
    ("4. Khai thác tối đa lợi thế hệ sinh thái mở và mạng lưới thanh toán xuyên biên giới: ",
     "Tận dụng tối đa quy mô mạng lưới của BIDV và công nghệ thanh toán toàn cầu của Tenpay/WeChat Pay để phát triển thị trường thanh toán hai chiều Việt – Trung, đón đầu các sự kiện đối ngoại cấp cao và phục vụ đắc lực cho hoạt động giao thương, du lịch giữa nhân dân hai nước."),
    
    ("5. Tăng cường phối hợp chặt chẽ và gắn kết trách nhiệm đầu mối: ",
     "Thiết lập cơ chế làm việc định kỳ giữa Tổ công tác chuyên trách của BIDV và các chuyên gia cấp cao của Tencent; giao nhiệm vụ cụ thể, rõ người, rõ việc, rõ mốc thời gian hoàn thành cho từng đơn vị chủ trì và đơn vị phối hợp.")
]

for p_title_text, p_body_text in principles:
    p_pr = doc.add_paragraph()
    p_pr.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_pr.paragraph_format.line_spacing = 1.15
    p_pr.paragraph_format.first_line_indent = Inches(0.5)
    p_pr.paragraph_format.space_after = Pt(6)
    r_pt = p_pr.add_run(p_title_text)
    set_run_font(r_pt, bold=True)
    r_pb = p_pr.add_run(p_body_text)
    set_run_font(r_pb)

# 6.2. Mục 2.2. Nội dung triển khai cụ thể: nội dung - đơn vị thực hiện - thời hạn
p_h22 = doc.add_paragraph()
p_h22.alignment = WD_ALIGN_PARAGRAPH.LEFT
p_h22.paragraph_format.space_before = Pt(8)
p_h22.paragraph_format.space_after = Pt(6)
p_h22.paragraph_format.keep_with_next = True
r = p_h22.add_run("2.2. Nội dung triển khai cụ thể: nội dung - đơn vị thực hiện - thời hạn")
set_run_font(r, bold=True)

specific_tasks = [
    (
        "1. Về tư vấn Chuyển đổi số toàn diện và Xây dựng Khung kiến trúc AI Platform:",
        "Phối hợp với Nhóm chuyên gia tư vấn chuyển đổi số và AI của Tencent Cloud (do ông Dư Duyên Tân làm đầu mối) tổ chức khảo sát hiện trạng, rà soát chiến lược dữ liệu và xây dựng Đề án Khung kiến trúc AI Platform tổng thể cho BIDV (bao gồm các lớp IaaS hạ tầng tính toán, PaaS quản lý dữ liệu, MaaS nền tảng mô hình lớn và Agent Platform); đồng thời thiết lập Khung quản trị AI Governance, cơ chế đánh giá an toàn mô hình và tuân thủ Luật AI tại Việt Nam.",
        "Giao Khối AI và Đổi mới sáng tạo (Ban Đổi mới sáng tạo làm đầu mối chủ trì), phối hợp Ban Công nghệ thông tin, Trung tâm Ngân hàng số và các đơn vị liên quan (chỉ đạo trực tiếp: Phó Tổng Giám đốc Vũ Hoàng Dương).",
        "Hoàn thành khảo sát và tổ chức Workshop chuyên sâu trong Tháng 11/2026; hoàn thiện Đề án Khung kiến trúc AI Platform báo cáo Ban Điều hành trong Quý I/2027."
    ),
    (
        "2. Về triển khai Thí điểm (Pilot) 03 bài toán Quick-win về AI, AIOps và Sinh trắc học:",
        "Trên cơ sở danh mục 13 bài toán (use cases) chuyên môn BIDV đã chuyển giao cho Tencent, hai bên thống nhất lựa chọn ngay 03 bài toán ưu tiên có ROI cao để triển khai Pilot gồm: (1) Ứng dụng trợ lý AI Multi-Agent hỗ trợ cán bộ nghiệp vụ nội bộ; (2) Giải pháp giám sát và vận hành tự động hệ thống CNTT (AIOps); và (3) Ứng dụng công nghệ eKYC sinh trắc học tiên tiến chống giả mạo sâu (Deepfake). Toàn bộ quá trình Pilot áp dụng chặt chẽ quy trình đánh giá 3 giai đoạn (trước – trong – sau triển khai) để làm căn cứ nhân rộng.",
        "Giao Ban Công nghệ thông tin chủ trì bài toán AIOps; Khối AI & Đổi mới sáng tạo chủ trì bài toán Multi-Agent; Trung tâm Ngân hàng số chủ trì giải pháp eKYC sinh trắc học; phối hợp chặt chẽ với đội ngũ Kiến trúc sư Kỹ thuật Tencent Cloud (chỉ đạo: PTGĐ Phan Thanh Hải, PTGĐ Vũ Hoàng Dương).",
        "Khởi động dự án Pilot trong Tháng 11/2026; nghiệm thu đánh giá kết quả giai đoạn 1 trước ngày 28/02/2027."
    ),
    (
        "3. Về mở rộng Mạng lưới Chấp nhận Thanh toán VN Global QR (Chiều khách Trung Quốc sang Việt Nam):",
        "Đẩy mạnh rà soát, nâng cấp hạ tầng kết nối kỹ thuật và mở rộng mạng lưới điểm chấp nhận thanh toán VN Global QR của BIDV trên toàn quốc nhằm phục vụ du khách và thương nhân Trung Quốc sử dụng ví WeChat Pay; xây dựng gói chính sách ưu đãi biểu phí, quy chuẩn nhận diện thương hiệu và giải pháp hỗ trợ merchant tối ưu.",
        "Giao Trung tâm Ngân hàng số (chủ trì), phối hợp Khối Khách hàng Doanh nghiệp, Khối Ngân hàng Bán lẻ và các Chi nhánh trên địa bàn du lịch trọng điểm; làm việc trực tiếp với đội ngũ WeChat Pay Quốc tế và Tenpay Global (chỉ đạo: PTGĐ Lê Huy Hoàng, PTGĐ Lại Tiến Quân).",
        "Rà soát và mở rộng triển khai đợt 1 hoàn thành trước ngày 30/11/2026; định kỳ báo cáo sản lượng giao dịch hàng tháng."
    ),
    (
        "4. Về phát triển Dịch vụ Thanh toán Chiều Việt Nam sang Trung Quốc (Inbound China Payment) đón đầu APEC 2026:",
        "Phối hợp khẩn trương với Tenpay Global và WeChat Pay hoàn thiện phương án kỹ thuật và thương mại, kết nối hệ sinh thái BIDV SmartBanking với hệ thống chấp nhận QR tại Trung Quốc để khách hàng BIDV có thể quét mã thanh toán trực tiếp tại các điểm chấp nhận của WeChat Pay; chuẩn bị đầy đủ hồ sơ pháp lý báo cáo Ngân hàng Nhà nước Việt Nam; quyết tâm đưa vào vận hành thử nghiệm phục vụ Lãnh đạo cấp cao và nhân dân hai nước trước thềm Hội nghị Cấp cao APEC 2026 tại Trung Quốc.",
        "Giao Trung tâm Ngân hàng số làm đầu mối chủ trì, phối hợp Ban Pháp chế, Ban Quản lý rủi ro hoạt động và Tenpay Global (chỉ đạo trực tiếp: Tổng Giám đốc Hoàng Việt Hùng, PTGĐ Lê Huy Hoàng).",
        "Hoàn thành thỏa thuận kỹ thuật trong Tháng 12/2026; hoàn thiện kết nối thử nghiệm trước ngày 31/03/2027; sẵn sàng truyền thông và vận hành chính thức phục vụ APEC 2026."
    ),
    (
        "5. Về phát triển Dịch vụ Chuyển tiền Xuyên biên giới trên BIDV SmartBanking:",
        "Nghiên cứu mô hình hợp tác kỹ thuật với Tenpay Global để tích hợp tính năng khởi tạo giao dịch chuyển tiền xuyên biên giới trực tiếp trên ứng dụng BIDV SmartBanking, phục vụ nhu cầu kiều hối, chi trả du học sinh, chuyên gia và giao thương biên mậu, đảm bảo tuân thủ nghiêm ngặt quy định quản lý ngoại hối của hai quốc gia.",
        "Giao Ban Khách hàng Cá nhân phối hợp Trung tâm Ngân hàng số, Ban Quản lý vốn & Tiền tệ và bộ phận kỹ thuật Tenpay Global (chỉ đạo: PTGĐ Lê Huy Hoàng, PTGĐ Nguyễn Thị Quỳnh Giao).",
        "Hoàn thành đề xuất phương án kinh doanh và luồng xử lý kỹ thuật trước ngày 15/01/2027."
    ),
    (
        "6. Về khai thác Hạ tầng Trung tâm Dữ liệu (Data Center) của Tencent tại Việt Nam và Kiến trúc SuperApp:",
        "Khảo sát năng lực Trung tâm dữ liệu của Tencent Cloud tại Việt Nam (dự kiến đi vào vận hành vào cuối năm 2026) để đánh giá khả năng thuê hạ tầng, thiết lập kênh truyền dự phòng và nghiên cứu mô hình kiến trúc Siêu ứng dụng (SuperApp / Mini-programs) học tập từ WeChat nhằm mở rộng hệ sinh thái số của BIDV.",
        "Giao Ban Công nghệ thông tin chủ trì khảo sát Data Center; Trung tâm Ngân hàng số chủ trì nghiên cứu kiến trúc SuperApp; phối hợp với Tencent Cloud Việt Nam (chỉ đạo: PTGĐ Phan Thanh Hải).",
        "Báo cáo kết quả khảo sát và đề xuất phương án hợp tác trong Tháng 01/2027."
    ),
    (
        "7. Về nghiên cứu Hợp tác Đầu tư Tài chính Chiến lược kết hợp Hỗ trợ Kỹ thuật:",
        "Theo dõi, duy trì kênh trao đổi thông tin bảo mật cấp cao để phản hồi và cụ thể hóa đề xuất mở của Tổng Giám đốc BIDV về việc Tập đoàn Tencent nghiên cứu tham gia đầu tư vốn tài chính chiến lược vào BIDV gắn với gói cam kết hỗ trợ kỹ thuật và chuyển giao năng lực công nghệ dài hạn.",
        "Giao Khối Kế hoạch và Tài chính (chủ trì), phối hợp Ban Thư ký HĐQT & Quan hệ cổ đông, Văn phòng BIDV; làm đầu mối tiếp nhận thông tin từ Lãnh đạo Tencent FinTech (chỉ đạo trực tiếp: Tổng Giám đốc Hoàng Việt Hùng, PTGĐ Nguyễn Thị Quỳnh Giao).",
        "Duy trì kênh trao đổi thông tin bảo mật, báo cáo cập nhật tình hình tiếp xúc định kỳ hàng quý trong Năm 2027."
    ),
    (
        "8. Về tổ chức thực hiện Biên bản ghi nhớ (MOU) và các phiên Hội thảo chuyên đề tiếp theo:",
        "Tiếp nhận, lưu trữ và theo dõi chặt chẽ tiến độ triển khai văn bản Biên bản ghi nhớ hợp tác thanh toán xuyên biên giới (MOU) đã ký kết chính thức hồi 09h50 ngày 27/10/2026 giữa BIDV (do Phó Tổng Giám đốc Lê Huy Hoàng đại diện ký) và Tenpay Global (do Phó Tổng Giám đốc Dương Văn Hối đại diện ký); tổng hợp đầy đủ kết quả chuyên môn của 02 Phiên hội thảo chuyên đề chiều 27/10/2026 (Phiên Chiến lược thương mại và Phiên Kỹ thuật công nghệ) để xây dựng Kế hoạch hành động chi tiết.",
        "Giao Văn phòng BIDV (Phó Giám đốc Văn phòng Ngô Minh Hoàng làm đầu mối), phối hợp Trung tâm Ngân hàng số và các Ban chuyên môn liên quan (chỉ đạo: Tổng Giám đốc Hoàng Việt Hùng).",
        "Hoàn thành báo cáo tổng hợp chuyến công tác trình Tổng Giám đốc trước ngày 05/11/2026."
    )
]

for title_item, nd_text, dv_text, th_text in specific_tasks:
    # Tiêu đề mục: 1. Về ...:
    p_item = doc.add_paragraph()
    p_item.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_item.paragraph_format.line_spacing = 1.15
    p_item.paragraph_format.first_line_indent = Inches(0.5)
    p_item.paragraph_format.space_before = Pt(6)
    p_item.paragraph_format.space_after = Pt(3)
    p_item.paragraph_format.keep_with_next = True
    r = p_item.add_run(title_item)
    set_run_font(r, bold=True)
    
    # Gạch đầu dòng 1: - Nội dung chỉ đạo và thống nhất:
    p_nd = doc.add_paragraph()
    p_nd.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_nd.paragraph_format.line_spacing = 1.15
    p_nd.paragraph_format.left_indent = Inches(0.5)
    p_nd.paragraph_format.first_line_indent = Inches(0.0)
    p_nd.paragraph_format.space_after = Pt(3)
    r1 = p_nd.add_run("- Nội dung chỉ đạo và thống nhất: ")
    set_run_font(r1, bold=True)
    r2 = p_nd.add_run(nd_text)
    set_run_font(r2)
    
    # Gạch đầu dòng 2: - Đơn vị thực hiện:
    p_dv = doc.add_paragraph()
    p_dv.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_dv.paragraph_format.line_spacing = 1.15
    p_dv.paragraph_format.left_indent = Inches(0.5)
    p_dv.paragraph_format.first_line_indent = Inches(0.0)
    p_dv.paragraph_format.space_after = Pt(3)
    r1 = p_dv.add_run("- Đơn vị thực hiện: ")
    set_run_font(r1, bold=True)
    r2 = p_dv.add_run(dv_text)
    set_run_font(r2)
    
    # Gạch đầu dòng 3: - Thời hạn hoàn thành: (in nghiêng phần thời hạn)
    p_th = doc.add_paragraph()
    p_th.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_th.paragraph_format.line_spacing = 1.15
    p_th.paragraph_format.left_indent = Inches(0.5)
    p_th.paragraph_format.first_line_indent = Inches(0.0)
    p_th.paragraph_format.space_after = Pt(6)
    r1 = p_th.add_run("- Thời hạn hoàn thành: ")
    set_run_font(r1, bold=True)
    r2 = p_th.add_run(th_text)
    set_run_font(r2, italic=True)

# 7. Đoạn kết văn bản chuẩn BIDV
p_conclusion = doc.add_paragraph()
p_conclusion.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p_conclusion.paragraph_format.line_spacing = 1.15
p_conclusion.paragraph_format.first_line_indent = Inches(0.5)
p_conclusion.paragraph_format.space_before = Pt(8)
p_conclusion.paragraph_format.space_after = Pt(14)
r = p_conclusion.add_run("Trên đây là nội dung biên bản kết luận và thống nhất phân công triển khai tại Phiên họp cấp cao song phương ngày 27/10/2026 giữa Ban Lãnh đạo BIDV và Ban Lãnh đạo Tập đoàn Tencent về hợp tác chiến lược toàn diện trong lĩnh vực Chuyển đổi số, AI, Cloud và Thanh toán xuyên biên giới. Văn phòng BIDV trân trọng thông báo để các đơn vị thuộc hệ thống BIDV biết và nghiêm túc triển khai thực hiện; đồng thời phối hợp chặt chẽ với Tổ Thư ký và các đơn vị đối tác phía Tencent theo dõi, đôn đốc tiến độ định kỳ báo cáo Ban Điều hành./.")
set_run_font(r)

# 8. Footer Table: Bảng Nơi nhận và Chữ ký cuối trang (1 dòng, 2 cột không đường viền)
footer_table = doc.add_table(rows=1, cols=2)
footer_table.alignment = WD_TABLE_ALIGNMENT.CENTER
footer_table.autofit = False

# Ngăn bảng bị tách trang
trPr = footer_table.rows[0]._tr.get_or_add_trPr()
trPr.append(parse_xml(r'<w:cantSplit {}/>'.format(nsdecls('w'))))

f_left = footer_table.cell(0, 0)
f_right = footer_table.cell(0, 1)
f_left.width = Inches(3.50)
f_right.width = Inches(3.25)
set_cell_borderless(f_left)
set_cell_borderless(f_right)

# Cột trái: Nơi nhận
p_nn = f_left.paragraphs[0]
p_nn.alignment = WD_ALIGN_PARAGRAPH.LEFT
p_nn.paragraph_format.space_after = Pt(2)
r_nn_title = p_nn.add_run("Nơi nhận:")
set_run_font(r_nn_title, size_pt=12.0, bold=True, italic=True)

noi_nhan_items = [
    "- Lãnh đạo tham dự cuộc họp;",
    "- Ban Lãnh đạo Tập đoàn Tencent;",
    "- Các Khối, Ban, Trung tâm thuộc BIDV;",
    "- Lưu: VT, VP, Tổ Thư ký đoàn."
]
for item in noi_nhan_items:
    p_item = f_left.add_paragraph()
    p_item.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_item.paragraph_format.space_after = Pt(1)
    r_item = p_item.add_run(item)
    set_run_font(r_item, size_pt=11.0)

# Cột phải: Chữ ký Lãnh đạo
p_sign = f_right.paragraphs[0]
p_sign.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_sign.paragraph_format.line_spacing = 1.15
p_sign.paragraph_format.space_after = Pt(2)

r_auth = p_sign.add_run("ĐẠI DIỆN ĐỒNG CHỦ TRÌ PHIÊN HỌP\n")
set_run_font(r_auth, size_pt=12.5, bold=True)
r_pos = p_sign.add_run("TỔNG GIÁM ĐỐC BIDV")
set_run_font(r_pos, size_pt=12.0, bold=True)

# 5 dòng trống để ký
for _ in range(5):
    p_blank = f_right.add_paragraph()
    p_blank.paragraph_format.space_after = Pt(0)
    p_blank.paragraph_format.line_spacing = 1.0

p_name = f_right.add_paragraph()
p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_name.paragraph_format.space_after = Pt(0)
r_name = p_name.add_run("HOÀNG VIỆT HÙNG")
set_run_font(r_name, size_pt=12.0, bold=True)

# Save document
os.makedirs(os.path.dirname(OUTPUT_DOCX), exist_ok=True)
doc.save(OUTPUT_DOCX)
print(f"SUCCESS: Saved BIDV Meeting Minutes to {OUTPUT_DOCX}")

# Convert to PDF
try:
    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", OUTPUT_DOCX, "--outdir", os.path.dirname(OUTPUT_PDF)], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    if os.path.exists(OUTPUT_PDF):
        print(f"SUCCESS: Saved PDF to {OUTPUT_PDF}")
except Exception:
    pass
