**📚Hướng Dẫn Sử Dụng Hàm NXA_AIMemories**

🌟 Chức năng:
Hàm `NXA_AIMemories` 🤝 giúp bạn tương tác trực tiếp với mô hình ngôn ngữ lớn (Large Language Model - LLM) của Google AI, có tên là Gemini, trong môi trường giống như trò chuyện 💬.
Bạn có thể đặt câu hỏi cho Gemini và nhận lại câu trả lời dựa trên kho kiến thức rộng lớn và khả năng xử lý ngôn ngữ của nó 🧠.
*Đặc biệt, hàm này còn có khả năng 💾 lưu trữ lịch sử trò chuyện để tạo ngữ cảnh liền mạch cho các tương tác tiếp theo của bạn.*

📝 Các đối số:
`text (bắt buộc)`: 🔑 Kiểu dữ liệu String. Đây là câu hỏi hoặc yêu cầu mà bạn muốn đặt cho Gemini.

`reset (tùy chọn)`: 🔄 Kiểu dữ liệu Boolean (mặc định là False). Đối số này xác định liệu bạn có muốn xóa lịch sử trò chuyện và bắt đầu lại một cuộc hội thoại mới tinh hay không.


🚀 Cách sử dụng:

Nhập công thức: ⌨️ Trong ô bạn muốn tương tác với Gemini, hãy nhập công thức sau: `=NXA_AIMemories(text, [reset])`.

Thay thế các đối số:

`text`: ✍️ Nhập trực tiếp câu hỏi hoặc yêu cầu của bạn vào đây.

`reset`: (Tùy chọn) 💡 Điền `True` để xóa lịch sử trò chuyện và bắt đầu lại từ đầu.

Nhấn Enter: ✅ Sau khi nhập xong, nhấn Enter để chạy hàm và nhận câu trả lời từ Gemini.


💡 Ví dụ:
Giả sử bạn muốn hỏi Gemini về thủ đô của nước Pháp:

Excel
`=NXA_AIMemories("What is the capital of France?")`

Để xóa toàn bộ ký ức của AI và bắt đầu cuộc trò chuyện mới, bạn sử dụng:

Excel
`=NXA_AIMemories(“reset”, TRUE)`

⚠️ Hãy thận trọng khi sử dụng các đoạn mã.

📌 Lưu ý quan trọng:
Kết nối internet: 🌐 Cần đảm bảo có kết nối internet ổn định để sử dụng hàm một cách suôn sẻ.

`API Key:` 🔑 Hiện tại, hướng dẫn sử dụng cung cấp `API Key mẫu` nhưng bạn cần có `API Key riêng` của mình để kích hoạt đầy đủ chức năng.

Lịch sử trò chuyện: 📂 Lịch sử trò chuyện được lưu trữ trong một tệp tin văn bản trên máy tính của bạn. Bạn có thể xóa lịch sử này một cách dễ dàng bằng cách đặt đối số reset thành `True`.

Ngôn ngữ: 🗣️ Hiện tại, Gemini chỉ hỗ trợ tiếng Anh.


🛠️ Các trường hợp sử dụng khác:
❓ Đặt câu hỏi về kiến thức tổng hợp thuộc mọi lĩnh vực.

📝 Yêu cầu Gemini thực hiện các tác vụ đơn giản bằng ngôn ngữ (ví dụ: tạo danh sách, gợi ý tên).

✍️ Sử dụng Gemini để sáng tạo nội dung văn bản (ví dụ: viết email, bài đăng blog).

🎯 Tóm tắt:
Hàm `NXA_AIMemories` là một công cụ 📈 thú vị và mạnh mẽ để tương tác với mô hình ngôn ngữ lớn Gemini. Nó cho phép bạn đặt câu hỏi, nhận câu trả lời thông minh và tận dụng tối đa khả năng xử lý ngôn ngữ vượt trội của Gemini. 
**Hãy luôn lưu ý về API Key và ngôn ngữ hỗ trợ khi sử dụng hàm này nhé!**
