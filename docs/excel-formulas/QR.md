#📚 Hướng Dẫn Sử Dụng Hàm NNCT_QR
🌟 Chức năng:
Hàm `NNCT_QR` 📸 là một công cụ tiện lợi giúp bạn tạo và chèn mã QR (Quick Response Code) trực tiếp vào bảng tính Excel của mình. Mã QR này có thể chứa nhiều loại thông tin như URL trang web, văn bản, số điện thoại, địa chỉ email, v.v., và có thể dễ dàng được quét bằng điện thoại thông minh hoặc các thiết bị di động khác để truy cập nhanh thông tin đó 🚀.

📝 Cú pháp:
Excel

`=NNCT_QR(cellOrURL, [imgCell], [autoResize], [width], [height])`
⚙️ Các đối số:
`cellOrURL (Bắt buộc)`: 🔑 Kiểu dữ liệu String hoặc Range. Đây là nội dung bạn muốn mã hóa thành QR code. Bạn có thể nhập trực tiếp một chuỗi văn bản (ví dụ: `"https://example.com"`, "Xin chào thế giới") hoặc tham chiếu đến một ô chứa nội dung đó (ví dụ: A1).
`imgCell (Tùy chọn)`: 📍 Kiểu dữ liệu Range. Xác định vị trí ô bạn muốn chèn hình ảnh mã QR. Nếu bạn bỏ qua đối số này, hình ảnh sẽ được chèn vào vị trí mặc định (thường là ô chứa công thức hoặc gần đó).
`autoResize` (Tùy chọn): 📐 Kiểu dữ liệu Boolean (mặc định là `TRUE`).
`TRUE`: Hàm sẽ tự động điều chỉnh kích thước của hình ảnh mã QR sao cho vừa với ô chứa ảnh (`imgCell`).
`FALSE`: Hàm sẽ giữ kích thước gốc của mã QR (hoặc sử dụng width/height nếu bạn cung cấp).
`width (Tùy chọn)`: ↔️ Kiểu dữ liệu Double. Chiều rộng mong muốn của hình ảnh mã QR (tính bằng pixel). Đối số này chỉ có tác dụng khi `autoResize` được đặt là `FALSE`.
`height (Tùy chọn)`: ↕️ Kiểu dữ liệu Double. Chiều cao mong muốn của hình ảnh mã QR (tính bằng pixel). Đối số này cũng chỉ có tác dụng khi `autoResize` được đặt là `FALSE`.
🚀 Cách sử dụng:
Nhập công thức: ⌨️ Trong ô bạn muốn hình ảnh mã QR hiển thị, nhập công thức `=NNCT_QR(...)`.
Thay thế các đối số: 🖋️
`cellOrURL`: Điền nội dung bạn muốn mã hóa (ví dụ: "`https://www.google.com`") hoặc tham chiếu đến ô chứa nội dung đó (ví dụ: A1).
`imgCel`l: (Tùy chọn) Điền tham chiếu ô bạn muốn chèn hình ảnh (ví dụ: B2).
`autoResize`, `width`, `height`: (Tùy chọn) Nếu bạn muốn kiểm soát kích thước ảnh, hãy điều chỉnh các đối số này.
Nhấn Enter: ✅ Sau khi nhập xong công thức, nhấn Enter để thực hiện hàm. Mã QR sẽ được tạo và chèn vào vị trí bạn đã chỉ định.
💡 Ví dụ:
Giả sử bạn muốn tạo mã QR cho URL "`https://www.example.com`" và chèn mã QR đó vào ô B2, đồng thời tự động điều chỉnh kích thước để vừa với ô B2:

Excel

`=NNCT_QR("https://www.example.com", B2)`
Kết quả: (Mã QR cho` https://www.example.com` sẽ xuất hiện trong ô B2).

+------------+-----------------------------------------+
| ✍️ Cú pháp  | 🌟 Kết quả                             |
+------------+-----------------------------------------+
| =QR(A1)    | 🖼️ Hình ảnh mặc định, đặt bên ô A1     |
| =QR(A1,C1) | 🎯 Hình ảnh mặc định, đặt tại ô C1      |
| =QR(A1,,300x300) | 📐 Hình ảnh kích thước 300x300, đặt bên ô A1 |
| =QR(A1,C1,300x300) | 📏 Hình ảnh kích thước 300x300, đặt tại ô C1 |
+------------+-----------------------------------------+


> [!CAUTION]
⚠️ Hãy thận trọng khi sử dụng các đoạn mã.


> [!IMPORTANT]
📌 Lưu ý quan trọng:
**Kết nối Internet:** 🌐 Để tạo và tải hình ảnh mã QR, hàm cần đảm bảo có kết nối internet ổn định. Nếu không có internet, quá trình này sẽ thất bại.
Kích thước mã QR: 📏 Kích thước của mã QR ảnh hưởng trực tiếp đến khả năng quét của các thiết bị. Mã QR quá nhỏ hoặc có độ phân giải thấp có thể khó quét. Hãy đảm bảo kích thước phù hợp với mục đích sử dụng.
Nội dung mã QR: 📄 Để mã QR hoạt động hiệu quả và dễ quét, nội dung mã hóa nên ngắn gọn và súc tích. Nội dung quá dài có thể làm cho mã QR phức tạp, khó quét hoặc không hoạt động.
🎨 Các trường hợp sử dụng khác:
Tạo mã QR cho thông tin liên hệ: 📞 Dễ dàng tạo mã QR cho số điện thoại, địa chỉ email, hoặc thậm chí là thông tin liên hệ đầy đủ (vCard) để chia sẻ nhanh chóng.
Tạo mã QR dẫn đến trang web: 🔗 Dùng để tạo các mã QR dẫn đến các trang sản phẩm, bài viết blog, hoặc bất kỳ URL nào bạn muốn.
Tạo mã QR để mở file/ứng dụng: 📁 Với một số cấu hình nhất định, bạn có thể tạo mã QR để dẫn đến việc mở một file cục bộ hoặc kích hoạt một ứng dụng trên thiết bị di động (tùy thuộc vào khả năng của thiết bị quét).
🎯 Tóm tắt:
Hàm `NNCT_QR` là một công cụ cực kỳ hữu ích để tạo và chèn mã QR trực tiếp vào bảng tính Excel. Nó giúp bạn chia sẻ thông tin một cách nhanh chóng, thuận tiện và tương tác, làm cho dữ liệu của bạn trở nên sinh động và dễ tiếp cận hơn. 🌟