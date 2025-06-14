#🌟 Hướng Dẫn Sử Dụng Hàm NXA_AIMemories 🌟
🎯 Chức năng:
💖 `Hàm NXA_AIMemories` giúp bạn tương tác với mô hình ngôn ngữ lớn (Large Language Model - LLM) của Google AI, có tên là Gemini, trong môi trường giống như trò chuyện. Bạn có thể đặt câu hỏi cho Gemini và nhận lại câu trả lời dựa trên kiến thức và khả năng xử lý ngôn ngữ của nó. Hàm này còn có khả năng lưu trữ lịch sử trò chuyện để tạo ngữ cảnh cho các tương tác tiếp theo.

📚 Các đối số:
📝 `text (bắt buộc):`` Kiểu dữ liệu String. Đây là câu hỏi hoặc yêu cầu mà bạn muốn đặt cho Gemini.
🔄 reset (tùy chọn): Kiểu dữ liệu Boolean (mặc định là False). Xác định có xóa lịch sử trò chuyện và bắt đầu lại cuộc hội thoại mới hay không.
🚀 Cách sử dụng:
✍️ Nhập công thức: Trong ô bạn muốn tương tác với Gemini, nhập công thức: =NXA_AIMemories(text, [reset]).
✏️ Thay thế các đối số:
text: Nhập câu hỏi hoặc yêu cầu của bạn.
reset: (Tùy chọn) Điền TRUE để xóa lịch sử trò chuyện và bắt đầu lại.
✅ Nhấn Enter: Nhấn Enter để chạy hàm và nhận câu trả lời từ Gemini.
💡 Ví dụ:
Giả sử bạn muốn hỏi Gemini về thủ đô của nước Pháp:

Excel

=NXA_AIMemories("What is the capital of France?")
Để xóa toàn bộ ký ức của AI:

Excel

=NXA_AIMemories(“reset”, TRUE)
⚠️ Hãy thận trọng khi sử dụng các đoạn mã.

📌 Lưu ý:
🌐 Kết nối internet: Cần đảm bảo có kết nối internet ổn định để sử dụng hàm.
🔑 API Key: Hiện tại, hướng dẫn sử dụng cung cấp API Key mẫu nhưng bạn cần có API Key riêng của mình để kích hoạt chức năng.
📜 Lịch sử trò chuyện: Lịch sử trò chuyện được lưu trữ trong tệp tin văn bản trên máy tính của bạn. Bạn có thể xóa lịch sử này bằng cách đặt đối số reset thành TRUE.
🗣️ Ngôn ngữ: Hiện tại, Gemini chỉ hỗ trợ tiếng Anh.
🌐 Các trường hợp sử dụng khác:
❓ Đặt câu hỏi về kiến thức tổng hợp.
✍️ Yêu cầu Gemini thực hiện các tác vụ đơn giản bằng ngôn ngữ.
📝 Sử dụng Gemini để sáng tạo nội dung văn bản.
📝 Tóm tắt:
Hàm NXA_AIMemories là một công cụ thú vị để tương tác với mô hình ngôn ngữ lớn Gemini. Nó cho phép bạn đặt câu hỏi, nhận câu trả lời và tận dụng khả năng xử lý ngôn ngữ của Gemini. Lưu ý về API Key và ngôn ngữ hỗ trợ khi sử dụng hàm.
