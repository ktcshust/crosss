import pandas as pd

# Tạo DataFrame với thông tin metadata
data = {
    "id": ["inv001", "pack001", "ship001"],
    "label": [0, 1, 2],  # 0: Invoice, 1: Packing List, 2: Shipping Instruction
    "file_full": [
        "[Standard] invoice (1).xlsx",
        "[Standard] packing_list.xlsx",
        "shipping_instruction (4) (1).xlsx",
    ],
    "file_key": [
        "[Standard] invoice (2) - Copy.xlsx",
        "[Standard] packing_list - Copy.xlsx",
        "shipping_instruction (4) (1) - Copy.xlsx",
    ],
}

df = pd.DataFrame(data)

# Xuất ra file Excel
output_file = "metadata.csv"
df.to_csv(output_file, index=False)

print(f"Đã tạo file Excel: {output_file}")
