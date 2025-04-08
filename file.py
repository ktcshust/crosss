from fastapi import APIRouter, UploadFile, File, HTTPException, Depends
from sqlalchemy.orm import Session
from typing import Dict, Any, List

from app.database import get_db
from app.service.file_service import (
    save_upload_file, 
    parse_excel, 
    extract_invoice_data_from_excel
)
from app.crud import invoice as invoice_crud, company as company_crud, product as product_crud
from app.schemas.schemas import InvoiceCreate, CompanyCreate, ProductCreate, ProductInInvoice, File_process
from datetime import datetime
from app.service import excel_parse, llm_service
from app.utils.utils import escape_json_string
router = APIRouter()
ai = llm_service.AIAgentService()


@router.post("/upload/")
async def upload_file(
    file: UploadFile = File(...),
    db: Session = Depends(get_db)
):
    """Upload file và trả về dữ liệu đã phân tích"""
    if not file.filename:
        raise HTTPException(status_code=400, detail="No file provided")
    
    file_extension = file.filename.split(".")[-1].lower()
    
    if file_extension not in ["xlsx", "xls", "pdf"]:
        raise HTTPException(
            status_code=400, 
            detail="Only Excel (.xlsx, .xls) and PDF (.pdf) files are supported"
        )
    
    file_path = await save_upload_file(file)
    print(file_path)
    return {"file_path": file_path}

@router.post("/process_excel/")
async def process_upload_file(data: File_process, db: Session = Depends(get_db)):
    file_name = data.file_name
    invoice_id = data.invoice_id
    file_path = file_name
    #Trích xuất thông tin từ file excel dùng thư viện openpyxl
    excel_data = excel_parse.Extract_excel_data(file_path)
    #Dùng llm (openai) để phân loại nội dung file excel đầu vào xem thuộc loại nào
    classcify_prompt = """Based on the extracted excel file content as below, try to classify the excel file content into the following types: invoice, packing list. Then return the result in Json format. Only return JSON, no further explanation needed. Example Output: '{"type": "invoice"}'.""" + f"""The Excel data is here: {excel_data}"""
    result = ai.run(prompt = classcify_prompt)
    file_type = escape_json_string(result)["type"]
    
    if file_type == "invoice":
        data = invoice_crud.get_invoice(db = db, invoice_id = invoice_id)
        data = data.as_dict()
    else:
        data = "None"
    #print(f"data: {data}")
    #Dùng LLM để map lại data lấy được từ DATABASE sang nội dung file excel đã trích xuất được.
    tranfer_data_prompt = f"""Based on the extracted excel file content of {file_type} as below and new data extracted from database. Compare similar information fields from the database data and update to the excel file content. Note that you still need to ensure the format is the same as the excel file content with the key being the coordinates and the value being the new text value. For the titles, the field names in the excel file remain the same, only change the items related to the value. For example: 'F25': '(Country of Origin)' or 'A5': 'ご依頼主 (Sender)：' are titles so don't change them. For titles, fields that do not have a reasonable value to fill in, leave them blank. Ensure the output content follows the JSON format. Here is the excel file information: {excel_data} And here is the new data taken from the database: 
{data}"""
    
    convert_data = ai.run(prompt = tranfer_data_prompt)
    excel_data = escape_json_string(convert_data)
    return {"format_data" : excel_data}

    


