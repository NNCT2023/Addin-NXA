**📚 Hướng Dẫn Sử Dụng Hàm NXA_SewingExpert**

### 🌟 Chức năng:

Hàm `NXA_SewingExpert` 🧵 biến Excel của bạn thành một chuyên gia về may ba lô và túi xách! 🎒👜 Hàm này sử dụng mô hình ngôn ngữ Gemini của Google AI, được tích hợp sẵn kiến thức chuyên sâu về ngành may mặc. Bạn có thể cung cấp dữ liệu về vấn đề hoặc đặt câu hỏi, và AI sẽ phân tích để đưa ra các giải pháp thực tế, chẩn đoán vấn đề và đề xuất cải tiến.

**Hàm có khả năng:**

  - Phân tích dữ liệu được cung cấp (từ một `Range`).
  - Trả lời các câu hỏi hoặc yêu cầu cụ thể của bạn.
  - Lưu trữ lịch sử tương tác để theo dõi.
  - Cung cấp phản hồi có cấu trúc theo yêu cầu.

---

### 📝 Cú pháp:

```=NXA_SewingExpert(rngData As Range, [prompt As String], [language As String], [structuredOutput As Boolean])```

### ⚙️ Các đối số:
  - `rngData (Bắt buộc)`: 🔑 Kiểu dữ liệu Range. Đây là vùng dữ liệu trong bảng tính Excel mà bạn muốn chuyên gia AI phân tích (ví dụ: dữ liệu về lỗi sản phẩm, thông số vật liệu, v.v.). Nếu không có dữ liệu, hãy chọn một ô trống.

  - `prompt (Tùy chọn)`: ✍️ Kiểu dữ liệu String. Câu hỏi hoặc yêu cầu cụ thể bạn muốn gửi đến chuyên gia AI (ví dụ: "Phân tích nguyên nhân đường may bị nhăn và đưa ra cách khắc phục.").

  - `language (Tùy chọn)`: 🌐 Kiểu dữ liệu String. Ngôn ngữ bạn muốn AI trả lời (mặc định là "vi" cho tiếng Việt). Bạn có thể dùng "en" cho tiếng Anh, v.v.

  - `structuredOutput (Tùy chọn)`: 📊 Kiểu dữ liệu Boolean. Nếu đặt là `TRUE`, AI sẽ cung cấp phản hồi có cấu trúc rõ ràng với các mục như loại vấn đề, cách khắc phục ngay lập tức, cải tiến dài hạn và ghi chú. Mặc định là `FALSE`.

---

### 🚀 Cách sử dụng:

  - **Đảm bảo Cấu hình API:** ⚙️ Trước khi sử dụng, hãy chắc chắn rằng bạn đã thiết lập đúng API Key và Model trong file `config.txt` (xem chi tiết ở phần Lưu ý Quan trọng).

  - **Nhập công thức:** ⌨️ Trong ô bạn muốn hiển thị kết quả phân tích hoặc câu trả lời từ chuyên gia, hãy nhập công thức `=NXA_SewingExpert(...)`.
Thay thế đối số:

    - `rngData:` 🖱️ Chọn vùng dữ liệu mà bạn muốn AI phân tích. Nếu bạn chỉ muốn đặt câu hỏi mà không có dữ liệu cụ thể, hãy chọn một ô trống.
    - `prompt (Tùy chọn)`: 💡 Nhập câu hỏi hoặc yêu cầu của bạn (ví dụ: "Làm thế nào để giảm thiểu đứt chỉ khi may nylon?").
    - `language (Tùy chọn)`: 🗣️ Nhập "`en`" nếu bạn muốn câu trả lời bằng tiếng Anh.
    - `structuredOutput (Tùy chọn)`: 📝 Nhập `TRUE` nếu bạn muốn câu trả lời được trình bày theo dạng gạch đầu dòng rõ ràng.
  
**Nhấn Enter:** ✅ Sau khi nhập xong, nhấn Enter. Excel sẽ gửi yêu cầu đến Gemini và hiển thị phân tích hoặc giải pháp trong ô.

---

### 💡 Ví dụ:

Giả sử bạn có một bảng dữ liệu về "`lỗi đường may`" ở vùng A1:C5 (ví dụ: Cột A: Loại lỗi, Cột B: Vật liệu, Cột C: Máy sử dụng) và bạn muốn chuyên gia AI phân tích:



+------------------------------------+-------------------------------------------------------------+-------------------------------------------------------------------------------------+-----------------------------------------------------------------------------------------------------+
| 🏷️ Ví dụ                          | 📝 Mô tả                                                     | 💡 Cú pháp hàm NXA_SewingExpert                                                      | 💬 Kết quả dự kiến                                                                                |
+------------------------------------+-------------------------------------------------------------+-------------------------------------------------------------------------------------+-----------------------------------------------------------------------------------------------------+
| 1️⃣ Phân tích tổng quát có dữ liệu | Phân tích các lỗi đường may phổ biến trong dữ liệu từ vùng | =NXA_SewingExpert(A1:C5, "Phân tích các lỗi đường may phổ biến trong dữ liệu này và | AI sẽ phân tích dữ liệu trong A1:C5 và đưa ra các nhận định, giải pháp tổng quát về lỗi đường may bằng |
|                                    | A1:C5 và đề xuất giải pháp bằng tiếng Việt, không cấu trúc. | đề xuất giải pháp.", "vi", FALSE)                                                  | tiếng Việt.                                                                                        |
+------------------------------------+-------------------------------------------------------------+-------------------------------------------------------------------------------------+-----------------------------------------------------------------------------------------------------+
| 2️⃣ Đặt câu hỏi không có dữ liệu  | Bạn không có dữ liệu cụ thể nhưng muốn hỏi chuyên gia về   | =NXA_SewingExpert("", "What are common causes of thread breakage and how to fix    | AI sẽ trả lời bằng tiếng Anh với cấu trúc như: 1. Type of issue, 2. Immediate fix, 3. Long-term    |
|                                    | vấn đề đứt chỉ bằng tiếng Anh, có cấu trúc.                 | them?", "en", TRUE)                                                                 | improvement, 4. Notes, liên quan đến vấn đề đứt chỉ.                                                |
+------------------------------------+-------------------------------------------------------------+-------------------------------------------------------------------------------------+-----------------------------------------------------------------------------------------------------+
| 3️⃣ Chỉ hiển thị phản hồi gần đây  | Không có dữ liệu hoặc prompt mới được cung cấp. Hàm sẽ hiển| =NXA_SewingExpert("", "", "vi", FALSE)                                             | Nếu không có dữ liệu hoặc prompt mới được cung cấp, hàm sẽ hiển thị 3 phản hồi gần đây nhất từ nhật |
|                                    | thị 3 phản hồi gần đây nhất từ nhật ký tương tác của bạn.   |                                                                                     | ký tương tác của bạn.                                                                               |
+------------------------------------+-------------------------------------------------------------+-------------------------------------------------------------------------------------+-----------------------------------------------------------------------------------------------------+



---

> [!IMPORTANT]
> ### 📌 Lưu ý quan trọng:

**Cấu hình API: ⚙️**
  - Hàm sử dụng thông tin cấu hình từ file `config.txt` nằm trong thư mục `C:\Users\{USERPROFILE}\Documents\ChatLogs\`.
  - File `config.txt` cần chứa `API_KEY` và `MODEL` của Gemini.
  - Nếu file `config.txt` không tồn tại, hàm sẽ tự động tạo một file mẫu với hướng dẫn chi tiết. Bạn cần thay thế `YOUR_API_KEY_HERE` bằng [API Key thực tế của bạn](https://aistudio.google.com/app/apikey).
  - Bạn có thể thay đổi `Model` muốn sử dụng (ví dụ: gemini-1.5-pro cho phân tích nâng cao, gemini-2.0-flash cho tốc độ).

---

**Kết nối internet:** 🌐 Hàm cần kết nối internet ổn định để gửi yêu cầu và nhận phản hồi từ API của Gemini.

**Định dạng dữ liệu:** 📊 Dữ liệu trong `rngData` sẽ được chuyển đổi thành JSON để gửi đến API. Hãy đảm bảo dữ liệu của bạn được sắp xếp theo dạng bảng với tiêu đề rõ ràng để Gemini hiểu tốt nhất.

**Độ chính xác:** 🎯 Chất lượng kết quả phụ thuộc vào độ rõ ràng của câu hỏi/dữ liệu và khả năng hiểu của Gemini.

**Yêu cầu [`JsonConverter`](https://github.com/VBA-tools/VBA-JSON):** 🧩 Hàm này cần có thư viện [`JsonConverter`](https://github.com/VBA-tools/VBA-JSON) để xử lý JSON. Đảm bảo bạn đã cài đặt và tham chiếu thư viện này trong VBA.

**Dịch vụ có phí:** 💰 Việc sử dụng API của Gemini có thể phát sinh chi phí theo số lượng yêu cầu và dữ liệu. Hãy tham khảo tài liệu của Google Cloud Platform.

**Lịch sử tương tác:** 💾 Mọi tương tác có dữ liệu hoặc prompt mới sẽ được ghi lại trong file `Feedback_Log.txt` tại `C:\Users\{USERPROFILE}\Documents\ChatLogs\`.

---

### 💡 Mẹo hữu ích:

- **Dữ liệu rõ ràng:** 📈 Luôn cung cấp dữ liệu rõ ràng, có cấu trúc và tiêu đề cột ý nghĩa để AI có thể phân tích chính xác nhất.

- **`Prompt cụ thể:`** 🎯 Sử dụng prompt để tùy chỉnh yêu cầu phân tích hoặc đặt câu hỏi chi tiết, giúp AI tập trung vào vấn đề bạn quan tâm.

- **Đầu ra có cấu trúc:** 📝 Khi cần thông tin để báo cáo hoặc theo dõi, hãy sử dụng `structuredOutput = TRUE` để nhận phản hồi dễ đọc và tổng hợp.

- **Kiểm tra nhật ký:** 📚 Nếu bạn không cung cấp dữ liệu hoặc `prompt` mới, hàm sẽ hiển thị các phản hồi gần đây từ nhật ký. Đây là cách hay để xem lại các tư vấn trước đó.

- **Khám phá Models:** 🧠 Thử nghiệm với các model Gemini khác nhau (ví dụ: gemini-1.5-pro cho phân tích sâu hơn) trong file config.txt để xem model nào phù hợp nhất với nhu cầu của bạn.


> [!NOTE]
> Với hàm `NXA_SewingExpert`, bạn đã có một trợ lý chuyên gia may mặc mạnh mẽ ngay trong Excel, giúp bạn nhanh chóng giải quyết các vấn đề và đưa ra quyết định sản xuất thông minh hơn! 🚀
