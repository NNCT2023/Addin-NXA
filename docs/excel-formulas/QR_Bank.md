#📚 Hướng Dẫn Sử Dụng Hàm NNCT_QR_Bank
🌟 Chức năng:
Hàm `NNCT_QR_Bank` 💳 giúp bạn tạo ra một URL để tạo mã QR thanh toán ngân hàng dựa trên thông tin tài khoản và các tham số tùy chọn. Hàm này giúp bạn dễ dàng tạo mã QR thanh toán cho các giao dịch ngân hàng trực tuyến, tiện lợi và nhanh chóng! 🚀

📝 Cú pháp:
Excel

`=NNCT_QR_Bank(TenNganHang, SoTaiKhoan, TenTaiKhoan, [SoTien], [NoiDung])`
⚙️ Các đối số:
`TenNganHang (Bắt buộc)`: 🔑 Kiểu dữ liệu String. Tên ngân hàng mà bạn muốn tạo mã QR (ví dụ: "Vietcombank", "Techcombank").
`SoTaiKhoan (Bắt buộc)`: 🔑 Kiểu dữ liệu String. Số tài khoản ngân hàng của người nhận.
`TenTaiKhoan (Bắt buộc)`: 🔑 Kiểu dữ liệu String. Tên chủ tài khoản ngân hàng (người nhận).
`SoTien (Tùy chọn)`: 💰 Kiểu dữ liệu Variant. Số tiền thanh toán bạn muốn bao gồm trong mã QR.
`NoiDung (Tùy chọn)`: 📝 Kiểu dữ liệu String. Nội dung thanh toán hoặc lời nhắn cho giao dịch (ví dụ: "Thanh toan hoa don A1").
🚀 Cách sử dụng:
Nhập công thức: ⌨️ Nhập công thức `=NNCT_QR_Bank(...)` vào ô bạn muốn hiển thị URL mã QR.
Thay thế các đối số: 🖋️ Điền các thông tin tương ứng vào các đối số:
`TenNganHang`: Nhập tên ngân hàng.
`SoTaiKhoan`: Nhập số tài khoản ngân hàng.
`TenTaiKhoan`: Nhập tên chủ tài khoản.
`SoTien (Tùy chọn)`: Nhập số tiền thanh toán (ví dụ: 100000).
`NoiDung (Tùy chọn)`: Nhập nội dung thanh toán (ví dụ: "Thanh toan tien dien").
Nhấn Enter: ✅ Excel sẽ tạo URL mã QR và hiển thị nó trong ô đã chọn.
💡 Ví dụ:
Bạn muốn tạo mã QR cho giao dịch 100.000 VND đến tài khoản "Nguyen Van A" tại "Vietcombank" với nội dung "Thanh toan hoa don":

Excel

`=NNCT_QR_Bank("Vietcombank", "1234567890", "Nguyen Van A", 100000, "Thanh toan hoa don")`
Kết quả: (URL mã QR thanh toán sẽ hiển thị trong ô Excel của bạn).

> [!IMPORTANT]
📌 Lưu ý quan trọng:
Kiểm tra dữ liệu: ✅ Hàm có tính năng kiểm tra tính hợp lệ của các tham số đầu vào. Nó sẽ đảm bảo các tham số bắt buộc không bị trống, tên ngân hàng và tên tài khoản không chứa số, và số tiền chỉ chứa số hoặc dấu thập phân.
Mã hóa URL: 🔐 Hàm sẽ tự động mã hóa (encode) nội dung và tên tài khoản để đảm bảo URL được định dạng chính xác và an toàn khi gửi yêu cầu.
URL cơ bản: 🔗 URL được tạo mặc định sẽ dẫn đến một hình ảnh mã QR định dạng .png.
Tham số tùy chọn: 💡 Các tham số SoTien và NoiDung là tùy chọn. Nếu bạn không cung cấp, mã QR sẽ chỉ chứa thông tin tài khoản.
API VietQR: 🌐 Hàm này sử dụng API của `VietQR` để tạo mã QR, đảm bảo tính tương thích và đáng tin cậy.

> [!TIP]
💡 Mẹo hữu ích:
Tạo mã QR nhanh chóng: 🚀 Sử dụng hàm này để tạo mã QR thanh toán nhanh chóng cho các giao dịch hàng ngày hoặc khi cần chia sẻ thông tin thanh toán.
Kiểm tra kỹ thông tin: 👁️ Luôn kiểm tra kỹ thông tin tên ngân hàng, số tài khoản, tên tài khoản và số tiền trước khi tạo mã QR để tránh sai sót không đáng có.
Thêm nội dung giao dịch: 📝 Sử dụng tham số NoiDung để thêm thông tin chi tiết về giao dịch, giúp người chuyển tiền dễ dàng điền nội dung và bạn dễ dàng đối soát.
🌟 Tăng cường trải nghiệm với mã QR:
🚀 Mẹo 1: Mở hình ảnh QR trực tiếp trong trình duyệt Web
Để biến URL thành một liên kết có thể nhấp chuột và tự động mở hình ảnh mã QR trong trình duyệt web mặc định của bạn, bạn chỉ cần kết hợp hàm NNCT_QR_Bank với hàm HYPERLINK của Excel:

Excel

`=HYPERLINK(NNCT_QR_Bank("Vietcombank", "1234567890", "Nguyen Van A", 100000, "Thanh toan hoa don"), "Bấm vào đây để mở hình ảnh ở trình duyệt Web")`
Kết quả: (Trong ô Excel sẽ hiển thị văn bản "Bấm vào đây để mở hình ảnh ở trình duyệt Web". Khi nhấp vào, hình ảnh mã QR sẽ mở ra trong trình duyệt mặc định của bạn như Chrome, Edge, v.v.)
> [!TIP]
🖼️ Mẹo 2: Chèn ảnh QR trực tiếp vào bảng tính Excel
Nếu bạn muốn chèn hình ảnh mã QR trực tiếp vào bảng tính mà không cần mở trình duyệt, bạn có thể sử dụng hàm `NNCT_IMAGE` (một hàm tùy chỉnh) hoặc hàm IMAGE mặc định trên *Microsoft 365*.

Ví dụ 1: Tự động căn chỉnh kích thước ảnh (sử dụng `NNCT_IMAGE`)
Excel

`=NNCT_IMAGE(NNCT_QR_Bank("Vietcombank", "1234567890", "Nguyen Van A", 100000, "Thanh toan hoa don"))`
Kết quả: (Hình ảnh mã QR sẽ được tải từ URL và chèn trực tiếp vào ô trong bảng tính của bạn, với kích thước tự động căn chỉnh để vừa với ô).

Ví dụ 2: Tùy chọn kích thước ảnh (sử dụng `NNCT_IMAGE`)
Nếu bạn muốn kiểm soát kích thước của hình ảnh, bạn có thể chỉ định chiều rộng và chiều cao mong muốn. Thay `300 x 300` bằng kích thước pixel mà bạn muốn:

Excel

`=NNCT_IMAGE(NNCT_QR_Bank("Vietcombank", "1234567890", "Nguyen Van A", 100000, "Thanh toan hoa don"), FALSE, 300, 300)`
Kết quả: (Hình ảnh mã QR sẽ được tải và chèn vào bảng tính hiện tại với kích thước ảnh tùy chỉnh của bạn).

Với hàm `NNCT_QR_Bank`, bạn có thể dễ dàng tạo và tích hợp mã QR thanh toán ngân hàng trực tiếp trong Excel, giúp tiết kiệm thời gian và tăng hiệu quả trong công việc tài chính của mình! 💰✨