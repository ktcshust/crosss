import re
import requests
import openpyxl
from openpyxl.utils import range_boundaries
import openai
import numpy as np


####################################
# Phần 1: Xử lý Excel và Extract  #
####################################
def preprocess_text(text):
    """
    Làm sạch văn bản: loại bỏ dấu nháy đơn, dấu câu và khoảng trắng thừa.
    """
    text = text.replace("'", "")
    text = re.sub(r'[^\w\s]', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def Extract_excel_data(file_path: str):
    """
    Đọc file Excel và trích xuất dữ liệu:
      - extracted_data: danh sách thông tin cell (tọa độ hoặc vùng merged)
      - extracted_value: danh sách giá trị sau khi làm sạch
      - combined_dict: dict nối extracted_data và extracted_value theo cặp key-value
    Loại bỏ các ô có giá trị chỉ chứa "n" hoặc nhiều chữ "n".
    """
    extracted_data = []
    extracted_value = []
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active  # Lấy sheet đầu tiên

    merged_ranges = list(sheet.merged_cells.ranges)

    def get_merged_range(cell, merged_ranges):
        for merged_range in merged_ranges:
            if cell.coordinate in merged_range:
                return merged_range
        return None

    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            value = repr(cell.value)
            clean_value = preprocess_text(value)
            # Loại bỏ ô chỉ chứa chữ "n" (một hoặc nhiều)
            if re.fullmatch(r"n+", clean_value):
                continue
            extracted_value.append(clean_value)

            coord = cell.coordinate
            merged_range = get_merged_range(cell, merged_ranges)
            if merged_range:
                merged_info = f"{merged_range}"
                extracted_data.append(merged_info)
            else:
                extracted_data.append(coord)

    combined_dict = dict(zip(extracted_data, extracted_value))
    return extracted_data, extracted_value, combined_dict


##############################################
# Phần 2: Ghép title với context từ 2 file xlsx #
##############################################
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
    Nối các ô title với các ô content từ 2 file Excel:
      - title_xlsx: file Excel chứa title (ví dụ: file trống)
      - full_xlsx: file Excel chứa đầy đủ content.

    Trả về dict với key là các ô title (dạng chuỗi vùng) và value là list các ô content nối với title đó.
    Các ô content không có title sẽ được gán key "không có title".

    Điều kiện đặc biệt: Nếu giá trị của ô title bắt đầu bằng dấu "*" thì vùng context của nó
    sẽ chính là tọa độ của chính ô title đó.
    """
    # Lấy thông tin từ file title: danh sách range và danh sách giá trị
    title_data = Extract_excel_data(title_xlsx)
    title_ranges_str = title_data[0]
    title_values = title_data[1]
    # Tạo map từ range_str -> value của ô title
    title_value_map = dict(zip(title_ranges_str, title_values))

    # Lấy danh sách vùng từ file full
    full_ranges_str = Extract_excel_data(full_xlsx)[0]

    # Các ô context là những ô có trong file full nhưng không có trong file title
    context_ranges_str = [item for item in full_ranges_str if item not in title_ranges_str]

    # Chuyển đổi chuỗi vùng thành dictionary chứa thông tin biên
    title_ranges = [parse_range(r) for r in title_ranges_str]
    context_ranges = [parse_range(r) for r in context_ranges_str]

    matched_contexts_all = set()
    title_to_context = {}

    # Với mỗi title, kiểm tra điều kiện đặc biệt và nếu không, thực hiện matching thông thường
    for current in title_ranges:
        curr_range_str = current['range_str']
        # Nếu giá trị của ô title bắt đầu bằng "*", gán context chính là chính nó
        if curr_range_str in title_value_map and title_value_map[curr_range_str].strip().startswith("*"):
            title_to_context[curr_range_str] = [curr_range_str]
            matched_contexts_all.add(curr_range_str)
            continue

        t_min_row = current['min_row']
        t_max_row = current['max_row']
        t_min_col = current['min_col']
        t_max_col = current['max_col']

        candidate_right = None
        candidate_below = None
        min_gap_right = None
        min_gap_below = None

        for other in title_ranges:
            if other['range_str'] == curr_range_str:
                continue
            if other['min_row'] <= t_max_row and other['max_row'] >= t_min_row and other['min_col'] > t_max_col:
                gap = other['min_col'] - t_max_col
                if min_gap_right is None or gap < min_gap_right:
                    candidate_right = other
                    min_gap_right = gap
            if other['min_col'] <= t_max_col and other['max_col'] >= t_min_col and other['min_row'] > t_max_row:
                gap = other['min_row'] - t_max_row
                if min_gap_below is None or gap < min_gap_below:
                    candidate_below = other
                    min_gap_below = gap

        right_boundary = candidate_right['min_col'] if candidate_right else None
        below_boundary = candidate_below['min_row'] if candidate_below else None

        matched_contexts = []
        for context in context_ranges:
            c_min_row = context['min_row']
            c_max_row = context['max_row']
            c_min_col = context['min_col']
            c_max_col = context['max_col']

            if c_min_col > t_max_col and c_min_row <= t_max_row and c_max_row >= t_min_row:
                if right_boundary is None or c_max_col < right_boundary:
                    matched_contexts.append(context)
                    matched_contexts_all.add(context['range_str'])
                    continue
            if c_min_row > t_max_row and c_min_col <= t_max_col and c_max_col >= t_min_col:
                if below_boundary is None or c_max_row < below_boundary:
                    matched_contexts.append(context)
                    matched_contexts_all.add(context['range_str'])
                    continue

        if matched_contexts:
            title_to_context[curr_range_str] = [ctx['range_str'] for ctx in matched_contexts]

    unmatched_contexts = [context for context in context_ranges if context['range_str'] not in matched_contexts_all]
    if unmatched_contexts:
        title_to_context["không có title"] = [context['range_str'] for context in unmatched_contexts]

    return title_to_context


##############################################
# Phần 3: Xử lý JSON API và Matching Fields   #
##############################################
def get_all_fields(data, parent_key=""):
    """
    Duyệt đệ quy qua dữ liệu JSON và trả về danh sách
    chứa tất cả các đường dẫn key, bao gồm cả các nút trung gian và nút lá.
    """
    result = []
    if isinstance(data, dict):
        for key, value in data.items():
            full_key = f"{parent_key}.{key}" if parent_key else key
            result.append(full_key)
            if isinstance(value, (dict, list)):
                result.extend(get_all_fields(value, full_key))
    elif isinstance(data, list):
        for item in data:
            result.extend(get_all_fields(item, parent_key))
    return result


def get_json_fields_from_url(url):
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        return get_all_fields(data)
    else:
        print(f"Yêu cầu thất bại với mã lỗi: {response.status_code}")
        return None


def preprocess(text):
    text = text.lower().strip().replace("", "")
    return text


def split_camel_case(text):
    text = re.sub(r'([a-z])([A-Z])', r'\1 \2', text)
    return text.lower()


def process_field_name(field):
    """
    Xử lý tên field:
      - Tách camel case (ví dụ: totalNetWeight -> total net weight)
      - Thay dấu chấm thành "'s" (amount.total -> amount's total)
      - Chuẩn hóa văn bản
    """
    field = split_camel_case(field)
    field = field.replace(".", "'s ")
    field = preprocess(field)
    return field


def extract_leaf(field):
    """
    Lấy phần cuối (leaf) của field. Ví dụ: "amount's total" sẽ trả về "total".
    """
    if "'s" in field:
        return field.split("'s")[-1].strip()
    else:
        return field


# Hàm lấy embeddings sử dụng OpenAI
def get_openai_embeddings(texts, model="text-embedding-ada-002"):
    """
    Gọi OpenAI API để lấy embeddings cho danh sách texts.
    Trả về numpy array có dạng (n_samples, embedding_dim).
    """
    response = openai.Embedding.create(input=texts, model=model)
    embeddings = [data["embedding"] for data in response["data"]]
    return np.array(embeddings)


def cosine_similarity(a, b):
    """
    Tính cosine similarity giữa hai mảng numpy a và b.
    a: (n, d)
    b: (m, d)
    Trả về ma trận similarity có dạng (n, m)
    """
    a_norm = a / np.linalg.norm(a, axis=1, keepdims=True)
    b_norm = b / np.linalg.norm(b, axis=1, keepdims=True)
    return np.dot(a_norm, b_norm.T)


def match_fields(api_url: str, xlsx_file: str) -> dict:
    """
    Lấy và xử lý các field từ API và file Excel.
    Sử dụng OpenAI để tính embeddings và matching các field.
    Tính cosine similarity trên cả full field và leaf field, sau đó kết hợp với trọng số.
    Trả về dict gồm:
      - combined_dict: cell -> value từ file Excel
      - best_match_field_dict: mapping (Excel field -> API field)
      - cell_to_best_match: mapping (cell -> API field)
    """
    # Lấy fields từ API và xử lý tên
    fields_database = get_json_fields_from_url(api_url)
    fields_database = [process_field_name(field) for field in fields_database]

    # Lấy fields từ file Excel
    _, fields_excel_raw, combined_dict = Extract_excel_data(xlsx_file)
    fields_excel = [process_field_name(field) for field in fields_excel_raw]

    # Tính embeddings cho full field sử dụng OpenAI
    embeddings_db_full = get_openai_embeddings(fields_database)
    embeddings_excel_full = get_openai_embeddings(fields_excel)

    # Lấy leaf của từng field
    leaf_db = [extract_leaf(field) for field in fields_database]
    leaf_excel = [extract_leaf(field) for field in fields_excel]

    embeddings_db_leaf = get_openai_embeddings(leaf_db)
    embeddings_excel_leaf = get_openai_embeddings(leaf_excel)

    # Tính cosine similarity
    cosine_scores_full = cosine_similarity(embeddings_db_full, embeddings_excel_full)
    cosine_scores_leaf = cosine_similarity(embeddings_db_leaf, embeddings_excel_leaf)

    threshold = 0.2
    best_match_for_excel = {}
    for j, excel_field in enumerate(fields_excel):
        best_match_score = -1
        best_match_field = None
        for i, db_field in enumerate(fields_database):
            score_full = cosine_scores_full[i][j]
            score_leaf = cosine_scores_leaf[i][j]
            # Kết hợp các score, ưu tiên leaf hơn full (0.99 vs 0.01)
            final_score = 0.01 * score_full + 0.99 * score_leaf
            if final_score > best_match_score:
                best_match_score = final_score
                best_match_field = db_field
        if best_match_score > threshold:
            best_match_for_excel[excel_field] = (best_match_field, best_match_score)

    best_match_field_dict = {excel_field: best_match_field
                             for excel_field, (best_match_field, score) in best_match_for_excel.items()}

    cell_to_best_match = {}
    for cell, value in combined_dict.items():
        processed_value = process_field_name(value)
        if processed_value in best_match_field_dict:
            cell_to_best_match[cell] = best_match_field_dict[processed_value]

    return {
        "combined_dict": combined_dict,
        "best_match_field_dict": best_match_field_dict,
        "cell_to_best_match": cell_to_best_match
    }


##############################################
# Phần 4: Tích hợp tất cả: Nối API field với Context cell
##############################################
def combine_all(api_url: str, title_xlsx: str, full_xlsx: str) -> dict:
    """
    Hàm tích hợp các bước:
      1. Lấy matching API fields với Excel từ file title (ví dụ: invoice_empty.xlsx)
      2. Lấy mapping title với context từ file title và file full (ví dụ: invoice_empty.xlsx và invoice.xlsx)
      3. Xây dựng dict nối API field với context cell, dựa vào giá trị của ô title.
         Nếu 1 title (trong combined_dict) khi xử lý trùng với key trong best_match_field_dict thì các ô context của title đó sẽ được nối với API field tương ứng.
    Trả về dict chứa:
      - combined_dict
      - best_match_field_dict
      - cell_to_best_match
      - title_to_context
      - api_field_to_context
    """
    # Bước 1: Matching API fields với Excel từ file title
    match_result = match_fields(api_url, title_xlsx)
    combined_dict = match_result["combined_dict"]
    best_match_field_dict = match_result["best_match_field_dict"]
    cell_to_best_match = match_result["cell_to_best_match"]

    # Bước 2: Lấy mapping title với context từ file title và file full
    title_to_context = match_title_with_context(title_xlsx, full_xlsx)

    # Bước 3: Xây dựng dict nối API field với context cell.
    api_field_to_context = {}
    for title_cell, context_list in title_to_context.items():
        if title_cell == "không có title":
            continue
        # Lấy giá trị của ô title từ combined_dict
        if title_cell in combined_dict:
            title_value = combined_dict[title_cell]
            processed_title = process_field_name(title_value)
            if processed_title in best_match_field_dict:
                api_field = best_match_field_dict[processed_title]
                if api_field not in api_field_to_context:
                    api_field_to_context[api_field] = []
                api_field_to_context[api_field].extend(context_list)

    return {
        "combined_dict": combined_dict,
        "best_match_field_dict": best_match_field_dict,
        "cell_to_best_match": cell_to_best_match,
        "title_to_context": title_to_context,
        "api_field_to_context": api_field_to_context
    }


##############################################################
# Ví dụ sử dụng tích hợp toàn bộ code
##############################################################
if __name__ == "__main__":
    # Thiết lập OpenAI API Key
    openai.api_key = "YOUR_OPENAI_API_KEY"

    api_url = "https://crossreach-api-dev.mystg-env.com/ai-data?order_id=1"
    title_file = "invoice_empty.xlsx"  # File chứa title
    full_file = "invoice.xlsx"  # File chứa đầy đủ (title và context)

    result = combine_all(api_url, title_file, full_file)

    print("Combined Excel Data (cell -> value):")
    print(result["combined_dict"])

    print("\nBest Match Field Dict (Excel field -> Best Match API Field):")
    print(result["best_match_field_dict"])

    print("\nCell to Best Match (Cell -> Best Match API Field):")
    print(result["cell_to_best_match"])

    print("\nTitle to Context (Title cell -> list Context cell):")
    print(result["title_to_context"])

    print("\nAPI Field to Context (API Field -> list Context cell):")
    print(result["api_field_to_context"])
