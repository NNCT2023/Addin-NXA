**📚Hướng Dẫn Sử Dụng Hàm NXA_ZConfigAPI**  

### ⚙️ NXA_ZConfigAPI: Thiết Lập Cấu Hình AI Gemini Trong Excel  

Hàm `NXA_ZConfigAPI` là một chức năng cốt lõi trong [Add-in NXA](https://github.com/XuanAn2018/Addin-NXA), cho phép bạn dễ dàng cấu hình `API Key` và mô hình `AI (model) của Gemini` để sử dụng trong các tính năng AI khác của Add-in. Hàm này giúp đảm bảo các chức năng liên quan đến AI có thể kết nối và hoạt động chính xác với dịch vụ Google Gemini.  

---

### 🎯 Mục Đích  

- Chức năng chính của `NXA_ZConfigAPI` là:  

  - **Lưu trữ cấu hình API:** 🔑 Ghi `API Key` và tên `model AI` vào một tệp tin `config.txt` đặt tại `Documents\ChatLogs`.

  - **Tạo thư mục:** 📂 Tự động tạo thư mục `ChatLogs` nếu chưa tồn tại, đảm bảo đường dẫn lưu trữ cấu hình hợp lệ.

  - **Xác thực đầu vào:** ✅ Kiểm tra tính hợp lệ của API Key và tên model được cung cấp trước khi lưu.

---

### 💻 Cách Sử Dụng

- Bạn có thể gọi hàm `NXA_ZConfigAPI` trực tiếp trong một ô Excel hoặc từ một module VBA khác.

  - **Cú pháp:**

    ```NXA_ZConfigAPI(apiKey As String, Optional model As String = "gemini-2.5-flash") As String```

---

### 📦 Tham số:

- `apiKey (Bắt buộc):` Chuỗi văn bản. 🔑 Đây là API Key của bạn từ Google AI Studio. [Bạn có thể lấy API Key tại đây](https://makersuite.google.com/app/apikey).

- `model (Tùy chọn):` Chuỗi văn bản. 🤖 Tên của mô hình Gemini bạn muốn sử dụng. Mặc định là "`gemini-2.5-flash`". Các mô hình được hỗ trợ bao gồm:
  - `"gemini-pro"`
  - `"gemini-1.5-flash"`
  - `"gemini-1.5-pro"`
  - `"gemini-2.0-flash"`
  - `"gemini-2.5-flash"`

### 📤 Giá trị trả về:

- `"Config updated successfully:` `[Tên Model]`": ✅ Nếu cấu hình được lưu thành công.

- `"Error: Invalid API Key":` ❌ Nếu API Key không hợp lệ (ví dụ: quá ngắn).

- `"Error: Invalid Model":` 🚫 Nếu tên model không được hỗ trợ.

---

### 📝 Ví dụ:


| 🏷️ Trường hợp              | 📝 Mô tả                                                         | 💡 Cú pháp hàm                                          | 📊 Kết quả                     |
|---------------------------|-----------------------------------------------------------------|--------------------------------------------------------|-------------------------------|
| 🔧 Cấu hình API           | Cấu hình API Key và mô hình Gemini-1.5-pro.                    | =NXA_ZConfigAPI("YOUR_API_KEY_HERE", "gemini-1.5-pro") | Cấu hình API thành công       |


---

### ⚠️ Lưu Ý Quan Trọng  

- **Bảo mật API Key:** 🔒 API Key của bạn là thông tin nhạy cảm. Không chia sẻ API Key công khai hoặc nhúng trực tiếp vào các tệp Excel mà người khác có thể truy cập mà không có biện pháp bảo mật.  

- **Cập nhật Models:** 🔄 Google thường xuyên cập nhật và giới thiệu các mô hình AI mới. Bạn nên kiểm tra trang tài liệu của Google AI để xem danh sách các mô hình được hỗ trợ mới nhất: `https://developers.google.com/generative-ai/models`.  
