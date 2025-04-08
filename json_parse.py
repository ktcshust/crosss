import requests
import json


def get_deepest_fields(data, parent_key="", result=None):
    if result is None:
        result = []
    
    if isinstance(data, dict):
        # Kiểm tra nếu tất cả các giá trị trong dictionary là giá trị lá (không phải dict hoặc list)
        if all(not isinstance(value, (dict, list)) for value in data.values()):
            result.append(parent_key)  # Chỉ thêm nút cha nếu tất cả con là giá trị lá
        else:
            for key, value in data.items():
                full_key = f"{parent_key}.{key}" if parent_key else key
                if isinstance(value, (dict, list)):
                    get_deepest_fields(value, full_key, result)
                else:
                    result.append(full_key)  # Thêm trường lá vào kết quả

    elif isinstance(data, list) and data:
        for item in data:
            get_deepest_fields(item, parent_key, result)

    return result


def get_json_fields_from_url(url):
    """
    Gửi yêu cầu GET đến URL, lấy dữ liệu JSON và trả về danh sách các trường sâu nhất.
    """
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        return get_deepest_fields(data)
    else:
        print(f"Yêu cầu thất bại với mã lỗi: {response.status_code}")
        return None


# URL API
url = "https://crossreach-api-dev.mystg-env.com/ai-data?order_id=1"

# Lấy danh sách các trường sâu nhất từ URL
fields = get_json_fields_from_url(url)
print(fields)



