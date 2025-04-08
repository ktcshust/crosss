import re
import openpyxl
import numpy as np
import gradio as gr
from io import BytesIO
from sklearn.metrics.pairwise import cosine_similarity
from sentence_transformers import SentenceTransformer
import google.generativeai as genai

# -----------------------------
# Cấu hình Gemini API
# -----------------------------
genai.configure(api_key="")  # Thay bằng API key thật của bạn
GEMINI_MODEL_NAME = "models/embedding-001"

# -----------------------------
# Khởi tạo SentenceTransformer
# -----------------------------
st_model = SentenceTransformer("sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2")

# -----------------------------
# Các hàm tiền xử lý và đọc file Excel
# -----------------------------
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

def extract_excel_data(file_obj):
    extracted_value = []
    workbook = openpyxl.load_workbook(file_obj, data_only=True)
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

def combine_key_value_text(key_file_obj, value_file_obj):
    keys = extract_excel_data(key_file_obj)
    values = extract_excel_data(value_file_obj)
    return " ".join(keys) + " " + " ".join(map(str, values))

# -----------------------------
# Các hàm lấy embedding
# -----------------------------
def get_st_embedding(text):
    return st_model.encode(text, convert_to_tensor=True).cpu().numpy()

def get_gemini_embedding(text, model_name=GEMINI_MODEL_NAME):
    try:
        response = genai.embed_content(model=model_name, content=text)
        return np.array(response['embedding'], dtype=np.float32)
    except Exception as e:
        print(f"Error in Gemini API: {e}")
        return None

def calculate_cosine_similarity(vec1, vec2):
    return cosine_similarity([vec1], [vec2])[0][0]

# -----------------------------
# Đọc sẵn 4 file tham chiếu từ disk
# -----------------------------
REF_FILES = {
    "file1": {"keys": "[Standard] invoice (2) - Copy.xlsx", "values": "[Standard] invoice (1).xlsx"},
    "file2": {"keys": "[Standard] packing_list - Copy.xlsx", "values": "[Standard] packing_list.xlsx"},
    "file3": {"keys": "shipping_instruction (4) (1) - Copy.xlsx", "values": "shipping_instruction (4) (1).xlsx"},
    "file4": {"keys": "食品ラベル - Copy.xlsx", "values": "食品ラベル.xlsx"},
}

def load_ref_text(path_key, path_value):
    with open(path_key, "rb") as f_key, open(path_value, "rb") as f_value:
        return combine_key_value_text(f_key, f_value)

ref_texts = {ref: load_ref_text(info["keys"], info["values"]) for ref, info in REF_FILES.items()}

# Tính sẵn embedding cho các file tham chiếu theo 2 cách
ref_embeddings_st = {ref: get_st_embedding(text) for ref, text in ref_texts.items()}
ref_embeddings_gemini = {ref: get_gemini_embedding(text) for ref, text in ref_texts.items()}

# -----------------------------
# Hàm xử lý: nhận file upload và trả về bảng kết quả similarity cùng loại file được gợi ý
# -----------------------------
def process_files(key_file, value_file):
    # key_file, value_file được upload từ Gradio dưới dạng bytes
    uploaded_text = combine_key_value_text(BytesIO(key_file), BytesIO(value_file))

    embedding_st = get_st_embedding(uploaded_text)
    embedding_gemini = get_gemini_embedding(uploaded_text)

    results = {}
    for ref in ref_texts.keys():
        sim_st = calculate_cosine_similarity(embedding_st, ref_embeddings_st[ref])
        sim_gemini = calculate_cosine_similarity(embedding_gemini, ref_embeddings_gemini[ref])
        results[ref] = {"SentenceTransformer": round(sim_st, 4),
                        "Gemini_API": round(sim_gemini, 4)}

    # Tìm file tham chiếu có cosine similarity cao nhất cho mỗi phương pháp
    best_ref_st = max(results.items(), key=lambda x: x[1]["SentenceTransformer"])[0]
    best_ref_gemini = max(results.items(), key=lambda x: x[1]["Gemini_API"])[0]

    # Xây dựng thông báo kết quả trả về
    output = "Kết quả so sánh:\n"
    for ref, sims in results.items():
        output += f"{ref} --> SentenceTransformer: {sims['SentenceTransformer']}, Gemini_API: {sims['Gemini_API']}\n"

    output += "\nPhân loại file upload:\n"
    output += f"- Theo SentenceTransformer, file gần giống với: {best_ref_st}\n"
    output += f"- Theo Gemini API, file gần giống với: {best_ref_gemini}\n"
    output += "- 1: Invoice, 2: Packing list, 3: Shipping instruction, 4: Label"

    return output

# -----------------------------
# Tạo giao diện Gradio
# -----------------------------
iface = gr.Interface(
    fn=process_files,
    inputs=[
        gr.File(label="File Keys", type="binary"),
        gr.File(label="File Values", type="binary")
    ],
    outputs=gr.Textbox(label="Kết quả similarity và phân loại"),
    title="So sánh độ tương đồng nội dung",
    description="Upload 2 file Excel (Keys và Values) để tính cosine similarity với 4 file tham chiếu theo 2 cách: SentenceTransformer và Gemini API, đồng thời cho biết file upload thuộc loại nào dựa trên kết quả similarity."
)

if __name__ == "__main__":
    iface.launch(server_name="192.168.1.100", server_port=8069, share=False)






