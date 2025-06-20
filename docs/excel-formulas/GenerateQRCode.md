**📚 Hướng Dẫn Sử Dụng Hàm NNCT_GenerateQRCode**  

### 🌟 Chức Năng:  
Hàm `NNCT_GenerateQRCode` 📸 giúp bạn tạo ra một mã QR (QR Code) dựa trên nội dung văn bản bạn cung cấp. Mã QR này là một cách nhanh chóng và dễ dàng để lưu trữ và chia sẻ thông tin, có thể quét bằng điện thoại thông minh để truy cập dữ liệu đã mã hóa.  

---

### 📝 Cú Pháp:  

`=NNCT_GenerateQRCode(content, [width], [height])`  

### ⚙️ Các Đối Số:  
  - `content (Bắt buộc)`: 🔑 Kiểu dữ liệu String. Đây là nội dung văn bản mà bạn muốn mã hóa thành mã QR (ví dụ: URL, số điện thoại, một đoạn text).  
  - `width (Tùy chọn)`: ↔️ Kiểu dữ liệu Integer. Xác định chiều rộng của mã QR (tính bằng pixel). Mặc định là 200 pixel nếu bạn không nhập.  
  - `height (Tùy chọn)`: ↕️ Kiểu dữ liệu Integer. Xác định chiều cao của mã QR (tính bằng pixel). Mặc định là 200 pixel nếu bạn không nhập.  

---

### 🚀 Cách Sử Dụng:
- **Nhập công thức:** ⌨️ Trong ô bạn muốn hiển thị liên kết đến mã QR, nhập công thức `=NNCT_GenerateQRCode(...)`.

- **Thay thế các đối số:** 🖋️
  - `content`: Nhập nội dung văn bản bạn muốn mã hóa (ví dụ: "`https://www.example.com`" hoặc tham chiếu đến một ô chứa văn bản đó).
  - `width`: (Tùy chọn) Điền giá trị chiều rộng mong muốn (ví dụ: 300).
  - `height`: (Tùy chọn) Điền giá trị chiều cao mong muốn (ví dụ: 300).
- Nhấn **Enter:** ✅ Nhấn Enter để chạy hàm. Excel sẽ hiển thị một liên kết đến mã QR được tạo trên trang web `api.qrserver.com`. Bạn có thể nhấp vào liên kết này để xem hình ảnh mã QR.

---

### 💡 Ví Dụ:  

| 🏷️ Trường hợp                     | 📝 Mô tả                                                         | 💡 Cú pháp hàm                                              | 📊 Kết quả                              |
|----------------------------------|-----------------------------------------------------------------|------------------------------------------------------------|----------------------------------------|
| 📷 Tạo mã QR mặc định            | Tạo mã QR cho URL với kích thước mặc định.                      | =NNCT_GenerateQRCode("https://www.example.com")            | Mã QR kích thước mặc định             |
| 📏 Tùy chỉnh kích thước mã QR     | Tạo mã QR với chiều rộng và cao 300 pixel.                      | =NNCT_GenerateQRCode("https://www.example.com", 300, 300)  | Mã QR kích thước 300x300 pixel        |


---

### ⚠️ Hãy thận trọng khi sử dụng các đoạn mã.

---

### 📌 Lưu Ý Quan Trọng:
- **Kết nối Internet:** 🌐 Để hàm có thể tạo và truy xuất mã QR từ dịch vụ web, bạn cần đảm bảo có kết nối internet ổn định. Hàm sẽ không hoạt động nếu không có mạng.
- **Nội dung mã hóa:** 📄 Hiện tại, hàm `NNCT_GenerateQRCode` chỉ hỗ trợ mã hóa nội dung văn bản. Đảm bảo dữ liệu bạn cung cấp là một chuỗi văn bản.
- **Kích thước mã QR:** 📏 Việc điều chỉnh width và height rất quan trọng để mã QR dễ dàng quét và hiển thị đẹp mắt. Mã QR quá nhỏ có thể khó quét, trong khi quá lớn có thể chiếm nhiều không gian không cần thiết.

---

### 🎯 Tóm Tắt:
Hàm `NNCT_GenerateQRCode` là một công cụ tiện lợi để tạo mã QR nhanh chóng ngay trong Excel. Hãy nhớ rằng bạn cần kết nối internet và nội dung mã hóa phải là văn bản. Chức năng tùy chỉnh kích thước cho phép bạn tạo mã QR phù hợp với mọi mục đích sử dụng. 🌟  
