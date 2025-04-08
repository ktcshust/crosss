from sentence_transformers import SentenceTransformer, util
from json_parse import get_json_fields_from_url
from excel_parse import Extract_excel_data
import re

def preprocess(text):
    """
    Chuẩn hóa văn bản: viết thường, bỏ dấu câu.
    """
    text = text.lower().strip().replace("", "")
    return text

def split_camel_case(text):
    """
    Chuyển từ CamelCase thành dạng thường có dấu cách.
    """
    text = re.sub(r'([a-z])([A-Z])', r'\1 \2', text)
    return text.lower()

def process_field_name(field):
    """
    Xử lý tên trường từ JSON API.
    - Tách camel case: totalNetWeight -> total net weight
    - Chuyển dấu chấm thành "'s": amount.total -> amount's total
    - Chuẩn hóa văn bản
    """
    field = split_camel_case(field)
    field = field.replace(".", "'s ")  # Chuyển dấu chấm thành sở hữu "'s"
    field = preprocess(field)
    return field

def extract_leaf(field):
    """
    Trích xuất phần leaf của trường.
    Nếu có dấu sở hữu ('s) thì lấy phần sau nó,
    nếu không thì lấy phần sau dấu cách cuối cùng.
    """
    if "'s" in field:
        return field.split("'s")[-1].strip()
    else:
        return field

# Khởi tạo model
model = SentenceTransformer('sentence-transformers/all-mpnet-base-v2')

# Lấy danh sách các trường từ JSON API
fields_database = get_json_fields_from_url("https://crossreach-api-dev.mystg-env.com/ai-data?order_id=1")
fields_database = [process_field_name(field) for field in fields_database]

# Lấy danh sách các trường từ Excel
fields_excel = Extract_excel_data('invoice_empty.xlsx')[1]
fields_excel = [process_field_name(field) for field in fields_excel]

# Tạo embedding cho các trường (full field)
embeddings_db_full = model.encode(fields_database, convert_to_tensor=True)
embeddings_excel_full = model.encode(fields_excel, convert_to_tensor=True)

# Tính leaf cho mỗi trường
leaf_db = [extract_leaf(field) for field in fields_database]
leaf_excel = [extract_leaf(field) for field in fields_excel]

# Tạo embedding cho phần leaf
embeddings_db_leaf = model.encode(leaf_db, convert_to_tensor=True)
embeddings_excel_leaf = model.encode(leaf_excel, convert_to_tensor=True)

# Tính cosine similarity cho full field và leaf
cosine_scores_full = util.cos_sim(embeddings_db_full, embeddings_excel_full)
cosine_scores_leaf = util.cos_sim(embeddings_db_leaf, embeddings_excel_leaf)

# Thiết lập ngưỡng similarity cuối cùng
threshold = 0.2

# Lưu trữ trường giống nhất cho mỗi trường trong Excel (bao gồm cả score)
best_match_for_excel = {}

for j, excel_field in enumerate(fields_excel):
    best_match_score = -1
    best_match_field = None
    for i, db_field in enumerate(fields_database):
        score_full = cosine_scores_full[i][j].item()
        score_leaf = cosine_scores_leaf[i][j].item()
        final_score = 0.01 * score_full + 0.99 * score_leaf
        
        # Tìm trường giống nhất (score cao nhất)
        if final_score > best_match_score:
            best_match_score = final_score
            best_match_field = db_field
    
    # Lưu trường giống nhất từ db_field cho mỗi excel_field nếu score vượt ngưỡng
    if best_match_score > threshold:
        best_match_for_excel[excel_field] = (best_match_field, best_match_score)

# Tạo dict mới chỉ chứa key là excel_field và value là best_match_field
best_match_field_dict = {excel_field: best_match_field for excel_field, (best_match_field, score) in best_match_for_excel.items()}

print(best_match_field_dict)




