import re
import openpyxl
import torch
from sklearn.metrics.pairwise import cosine_similarity
from sentence_transformers import SentenceTransformer

# Tải mô hình contrastive learning từ Hugging Face
model = SentenceTransformer("sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2")

# Hàm làm sạch văn bản
def preprocess_text(text):
    if isinstance(text, str) and text.startswith("*"):
        prefix = "*"
        text = text[1:]
    else:
        prefix = ""
    text = text.replace("'", "")
    text = re.sub(r'[^\w\s]', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return prefix + text

# Hàm trích xuất dữ liệu từ file Excel
def extract_excel_data(file_path: str):
    extracted_value = []
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            clean_value = preprocess_text(str(cell.value))
            if re.fullmatch(r"n+", clean_value):
                continue
            extracted_value.append(clean_value)
    return extracted_value

# Đọc dữ liệu từ file Excel
values_1 = extract_excel_data("[Standard] invoice (1).xlsx")
keys_1 = extract_excel_data("[Standard] invoice (2) - Copy.xlsx")

values_2 = extract_excel_data("[Standard] packing_list.xlsx")
keys_2 = extract_excel_data("[Standard] packing_list - Copy.xlsx")

values_3 = extract_excel_data("shipping_instruction (4) (1).xlsx")
keys_3 = extract_excel_data("shipping_instruction (4) (1) - Copy.xlsx")

values_4 = extract_excel_data("食品ラベル.xlsx")
keys_4 = extract_excel_data("食品ラベル - Copy.xlsx")

# Chuyển đổi keys và values thành chuỗi để tạo embedding
text_1 = " ".join(keys_1) + " " + " ".join(map(str, values_1))
text_2 = " ".join(keys_2) + " " + " ".join(map(str, values_2))
text_3 = " ".join(keys_3) + " " + " ".join(map(str, values_3))
text_4 = " ".join(keys_4) + " " + " ".join(map(str, values_4))

# Lấy embedding bằng SentenceTransformer
def get_embedding(text):
    return model.encode(text, convert_to_tensor=True).cpu().numpy()

embedding_1 = get_embedding(text_1)
embedding_2 = get_embedding(text_2)
embedding_3 = get_embedding(text_3)
embedding_4 = get_embedding(text_4)

# Hàm tính cosine similarity
def calculate_cosine_similarity(vec1, vec2):
    return cosine_similarity([vec1], [vec2])[0][0]

# Tính toán độ tương đồng giữa các cặp file
similarity_1_2 = calculate_cosine_similarity(embedding_1, embedding_2)
similarity_1_3 = calculate_cosine_similarity(embedding_1, embedding_3)
similarity_2_3 = calculate_cosine_similarity(embedding_2, embedding_3)
similarity_1_4 = calculate_cosine_similarity(embedding_1, embedding_4)
similarity_2_4 = calculate_cosine_similarity(embedding_2, embedding_4)
similarity_3_4 = calculate_cosine_similarity(embedding_3, embedding_4)

print(f"Độ tương đồng giữa file 1 và file 2: {similarity_1_2}")
print(f"Độ tương đồng giữa file 1 và file 3: {similarity_1_3}")
print(f"Độ tương đồng giữa file 2 và file 3: {similarity_2_3}")
print(f"Độ tương đồng giữa file 1 và file 4: {similarity_1_4}")
print(f"Độ tương đồng giữa file 2 và file 4: {similarity_2_4}")
print(f"Độ tương đồng giữa file 3 và file 4: {similarity_3_4}")




