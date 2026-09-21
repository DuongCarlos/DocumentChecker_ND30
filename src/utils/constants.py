from docx.enum.text import WD_COLOR_INDEX

class Decree30Rules:
    FONT_NAME = "Times New Roman"
    FONT_SIZE_MIN = 13.0
    FONT_SIZE_MAX = 14.0
    
class ErrorMessages:
    ERR_FONT = "[LỖI: Yêu cầu phông Times New Roman theo khoản 4 Mục I Phụ lục I]"
    ERR_SIZE = "[LỖI: Cỡ chữ phần lời văn phải từ 13-14 theo điểm e khoản 6 Mục II Phụ lục I]"
    ERR_ALIGN = "[LỖI: Nội dung phải canh đều hai lề (Justified)]"
    ERR_ALIGN_CENTER = "[LỖI: Dòng này bắt buộc phải canh giữa (Center)]"
    ERR_INDENT = "[LỖI: Chữ đầu dòng phải lùi vào 1 cm hoặc 1,27 cm theo điểm e khoản 6 Mục II Phụ lục I]"
    ERR_SPACING = "[LỖI: Khoảng cách đoạn tối thiểu 6pt. Vui lòng thiết lập After 6pt HOẶC Before 6pt]"
    
    # Mã lỗi cho Quốc hiệu / Tiêu ngữ
    ERR_QUOC_HIEU_TEXT = "[LỖI: Quốc hiệu viết sai chính tả hoặc sai khoảng trắng]"
    ERR_QUOC_HIEU_FONT = "[LỖI: Quốc hiệu phải in hoa, ĐẬM, cỡ chữ 12-13 theo điểm a khoản 1 Mục II Phụ lục I]"
    ERR_TIEU_NGU_TEXT = "[LỖI: Tiêu ngữ phải dùng gạch nối '-', viết hoa chữ cái đầu cụm từ theo điểm b khoản 1 Mục II Phụ lục I]"
    ERR_TIEU_NGU_FONT = "[LỖI: Tiêu ngữ phải in thường, ĐẬM, cỡ chữ 13-14 theo điểm b khoản 1 Mục II Phụ lục I]"

    # Mã lỗi cho Ngày tháng
    ERR_DATE_FONT = "[LỖI: Địa danh, ngày tháng phải in nghiêng, cỡ 13-14 theo điểm c khoản 4 Mục II Phụ lục I]"

    # Mã lỗi cho Tên loại và Trích yếu (BÁO CÁO / Về việc...)
    ERR_TITLE_FONT = "[LỖI: Tên loại văn bản phải in ĐẬM, cỡ 13-14 theo điểm a khoản 5 Mục II Phụ lục I]"
    ERR_ABSTRACT_FONT = "[LỖI: Trích yếu phải in ĐẬM, cỡ 13-14 theo điểm b khoản 5 Mục II Phụ lục I]"
    ERR_ABSTRACT_LINE = "[LỖI: Sai thể thức! Cấm dùng gạch chân (Underline). Phải tự vẽ đường kẻ ngang dài 1/3 - 1/2 bên dưới theo điểm b khoản 5 Mục II Phụ lục I]"

    # Mã lỗi cho Nơi nhận
    ERR_NOI_NHAN_ALIGN = "[LỖI: 'Nơi nhận' và danh sách phải sát lề trái theo điểm d khoản 9 Mục II Phụ lục I]"
    ERR_NOI_NHAN_FONT = "[LỖI: 'Nơi nhận:' phải in nghiêng, ĐẬM, cỡ 12 theo điểm d khoản 9 Mục II Phụ lục I]"
    ERR_NOI_NHAN_LIST = "[LỖI: Liệt kê nơi nhận phải in thường, ĐỨNG, cỡ 11 theo điểm d khoản 9 Mục II Phụ lục I]"

class HighlightColors:
    FORMAT_ERROR = WD_COLOR_INDEX.YELLOW
    NOTE_COLOR = WD_COLOR_INDEX.TURQUOISE