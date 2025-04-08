import openpyxl
from openpyxl.utils import range_boundaries
from excel_parse import Extract_excel_data

def parse_range(range_str):
    """
    Chuyển đổi chuỗi vùng (ví dụ: "B16:J20") thành dictionary với thông tin biên.
    """
    min_col, min_row, max_col, max_row = range_boundaries(range_str)
    return {
        'range_str': range_str,
        'min_row': min_row,
        'max_row': max_row,
        'min_col': min_col,
        'max_col': max_col
    }

def match_title_with_context(title_xlsx: str, full_xlsx: str) -> dict:
    """
    Hàm nối các ô title với các ô content từ 2 file Excel:
      - title_xlsx: file Excel chứa title (ví dụ: file trống)
      - full_xlsx: file Excel chứa đầy đủ content
      
    Trả về một dictionary với key là các ô title (dạng chuỗi vùng) và value là list các ô content (dạng chuỗi vùng)
    được nối với title đó. Các ô content không có title sẽ được gán key là "không có title".
    """
    # Lấy danh sách vùng từ file title và file full
    title_ranges_str = Extract_excel_data(title_xlsx)[0]
    full_ranges_str = Extract_excel_data(full_xlsx)[0]
    
    # Các vùng context là những vùng có trong file full nhưng không có trong file title
    context_ranges_str = [item for item in full_ranges_str if item not in title_ranges_str]
    
    # Chuyển đổi chuỗi vùng thành dictionary chứa thông tin biên
    title_ranges = [parse_range(r) for r in title_ranges_str]
    context_ranges = [parse_range(r) for r in context_ranges_str]
    
    # Tập hợp để lưu các ô context đã được ghép với title (dùng range_str)
    matched_contexts_all = set()
    # Dictionary nối title với các ô content
    title_to_context = {}
    
    # Với mỗi title, tìm title gần nhất bên phải và bên dưới để xác định ranh giới của context
    for current in title_ranges:
        t_min_row = current['min_row']
        t_max_row = current['max_row']
        t_min_col = current['min_col']
        t_max_col = current['max_col']
        
        candidate_right = None
        candidate_below = None
        min_gap_right = None  # khoảng cách cột giữa current và title bên phải
        min_gap_below = None  # khoảng cách hàng giữa current và title bên dưới
        
        # Duyệt qua các title khác để tìm candidate bên phải và bên dưới
        for other in title_ranges:
            if other['range_str'] == current['range_str']:
                continue
            # Tìm title bên phải: giao nhau theo hàng và nằm sau current
            if other['min_row'] <= t_max_row and other['max_row'] >= t_min_row and other['min_col'] > t_max_col:
                gap = other['min_col'] - t_max_col
                if min_gap_right is None or gap < min_gap_right:
                    candidate_right = other
                    min_gap_right = gap
            # Tìm title bên dưới: giao nhau theo cột và nằm bên dưới current
            if other['min_col'] <= t_max_col and other['max_col'] >= t_min_col and other['min_row'] > t_max_row:
                gap = other['min_row'] - t_max_row
                if min_gap_below is None or gap < min_gap_below:
                    candidate_below = other
                    min_gap_below = gap
        
        # Xác định ranh giới cho vùng context dựa theo candidate bên phải và bên dưới
        right_boundary = candidate_right['min_col'] if candidate_right else None
        below_boundary = candidate_below['min_row'] if candidate_below else None
        
        # Tìm các context nằm trong vùng giới hạn so với title hiện tại
        matched_contexts = []
        for context in context_ranges:
            c_min_row = context['min_row']
            c_max_row = context['max_row']
            c_min_col = context['min_col']
            c_max_col = context['max_col']
            
            # Kiểm tra context bên phải của title:
            if c_min_col > t_max_col and c_min_row <= t_max_row and c_max_row >= t_min_row:
                if right_boundary is None or c_max_col < right_boundary:
                    matched_contexts.append(context)
                    matched_contexts_all.add(context['range_str'])
                    continue
            # Kiểm tra context bên dưới của title:
            if c_min_row > t_max_row and c_min_col <= t_max_col and c_max_col >= t_min_col:
                if below_boundary is None or c_max_row < below_boundary:
                    matched_contexts.append(context)
                    matched_contexts_all.add(context['range_str'])
                    continue
        
        if matched_contexts:
            title_to_context[current['range_str']] = [ctx['range_str'] for ctx in matched_contexts]
    
    # Các ô context không được nối với title nào
    unmatched_contexts = [context for context in context_ranges if context['range_str'] not in matched_contexts_all]
    if unmatched_contexts:
        title_to_context["không có title"] = [context['range_str'] for context in unmatched_contexts]
    
    return title_to_context

# Ví dụ sử dụng:
if __name__ == "__main__":
    title_file = "invoice_empty.xlsx"  # File chứa các ô title
    full_file = "invoice.xlsx"         # File chứa đầy đủ các ô (title và context)
    result_dict = match_title_with_context(title_file, full_file)
    
    print("Dictionary nối title với content:")
    for key, value in result_dict.items():
        print(f"{key}: {value}")

