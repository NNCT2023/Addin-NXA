#📚 Hướng Dẫn Sử Dụng Hàm `NXA_Query`
🌟 Chức năng:
Hàm `NXA_Query` 💬 cho phép bạn đặt câu hỏi hoặc yêu cầu phân tích dữ liệu trực tiếp trong Excel, tận dụng sức mạnh của mô hình ngôn ngữ Gemini từ Google AI 🧠. Hàm sẽ lấy dữ liệu từ một vùng chọn, chuyển đổi thành định dạng phù hợp và gửi yêu cầu đến API của Gemini, sau đó trả về kết quả phân tích hoặc câu trả lời 📊.

📝 Các đối số:
`rngData (bắt buộc)`: 🔑 Kiểu dữ liệu Range. Đây là vùng dữ liệu trong bảng tính Excel mà bạn muốn truy vấn.
`prompt (bắt buộc)`: ✍️ Kiểu dữ liệu String. Đây là câu hỏi hoặc yêu cầu bạn muốn gửi đến Gemini để phân tích dữ liệu.
🚀 Cách sử dụng:
Nhập công thức: ⌨️ Nhập công thức `=NXA_Query(rngData, prompt)` vào ô bạn muốn hiển thị kết quả.
Thay thế đối số:
`rngData`: 🖱️ Chọn vùng dữ liệu chứa thông tin bạn muốn truy vấn (ví dụ: A1:C10).
`prompt`: 💡 Nhập câu hỏi hoặc yêu cầu của bạn một cách rõ ràng và cụ thể (ví dụ: "Tổng doanh số là bao nhiêu?").
Nhấn Enter: ✅ Excel sẽ gửi yêu cầu đến API của Gemini và hiển thị kết quả truy vấn trong ô đã chọn.
💡 Ví dụ:
Giả sử bạn có một bảng dữ liệu doanh số bán hàng trong vùng A1:C10 và muốn hỏi về tổng doanh số:

Excel

`=NXA_Query(A1:C10, "Tổng doanh số là bao nhiêu?")`
Kết quả: (Kết quả truy vấn, ví dụ: "Tổng doanh số là 15.000.000 VND" sẽ hiển thị trong ô).

> [!IMPORTANT]
📌 Lưu ý quan trọng:
**Cấu hình API: ⚙️**
Hàm sử dụng thông tin cấu hình từ file config.txt nằm trong thư mục `Documents\ChatLogs`.
File `config.txt` cần chứa `API_KEY` và `MODEL` của Gemini.
Nếu file `config.txt` không tồn tại, hàm sẽ tự động tạo một file mẫu với hướng dẫn chi tiết.
Bạn cần thay thế `YOUR_API_KEY_HERE` bằng [API Key thực tế của bạn](https://aistudio.google.com/app/apikey).
Bạn có thể thay đổi `Model` muốn sử dụng (ví dụ: gemini-pro).
Kết nối internet: 🌐 Hàm cần kết nối internet ổn định để gửi yêu cầu dữ liệu và nhận kết quả từ API của Gemini.
Định dạng dữ liệu: 📄 Dữ liệu trong rngData sẽ được chuyển đổi thành chuỗi văn bản để gửi đến API. Do đó, hãy đảm bảo dữ liệu trong Excel của bạn được định dạng rõ ràng, dễ hiểu.
Độ chính xác: 🎯 Chất lượng kết quả phụ thuộc vào độ phức tạp của câu hỏi và khả năng hiểu của Gemini đối với dữ liệu và yêu cầu của bạn.
Yêu cầu JsonConverter: 🧩 Hàm này cần có thư viện JsonConverter để có thể xử lý kết quả trả về dưới dạng Json từ API. Đảm bảo bạn đã cài đặt và tham chiếu thư viện này trong VBA.

> [!TIP]
💡 Mẹo hữu ích:
Phân tích phức tạp: 📊 Sử dụng hàm này để phân tích dữ liệu phức tạp hoặc đặt câu hỏi về các xu hướng, mối quan hệ trong dữ liệu mà các hàm Excel truyền thống khó thực hiện.
Kiểm tra câu hỏi: ✍️ Luôn kiểm tra kỹ câu hỏi của bạn để đảm bảo nó rõ ràng, cụ thể và dễ hiểu đối với Gemini, điều này sẽ cải thiện độ chính xác của kết quả.
Kiểm tra config.txt: 📂 Kiểm tra file config.txt định kỳ để đảm bảo API Key và Model được cung cấp đúng và còn hiệu lực.
Sử dụng JsonConverter: 💻 Đảm bảo rằng thư viện JsonConverter đã được thiết lập đúng cách trong VBA để tránh lỗi khi xử lý dữ liệu trả về từ Gemini.
Với hàm NXA_Query, bạn có thể dễ dàng truy vấn và phân tích dữ liệu trực tiếp trong Excel, khai thác tối đa sức mạnh của mô hình ngôn ngữ Gemini để có được những insight giá trị! 🚀