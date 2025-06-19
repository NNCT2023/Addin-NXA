**📚 Hướng Dẫn Sử Dụng Hàm NNCT_CallMacro**  

### 🌟 Chức năng:  
Hàm `NNCT_CallMacro` 🚀 là một công cụ mạnh mẽ cho phép bạn chèn mã VBA (Visual Basic for Applications) từ các URL an toàn trực tiếp vào sổ làm việc Excel đang mở của bạn. Điều này giúp bạn dễ dàng thêm các chức năng mới vào Excel mà không cần phải sao chép/dán mã VBA một cách thủ công.  

### 📝 Cú pháp:  

```=NNCT_CallMacro(macroInput)```  

---

### ⚙️ Các đối số:  

- `macroInput (Bắt buộc):` 🔑 Kiểu dữ liệu Variant. Đối số này xác định nguồn của mã VBA mà bạn muốn chèn:  

- `Số 1:` 🔢 Nếu bạn nhập số `1`, hàm sẽ tự động chèn một danh sách các macro cố định (hiện tại bao gồm: `AIO_MyMacro.bas`, `QR_QuickChart.bas`, `QRcode_ChatGPT.bas`) được lưu trữ trên GitHub Gist của tác giả [XuanAn2018](https://gist.githubusercontent.com/XuanAn2018/).  

- Chuỗi (URL): 🔗 Nếu bạn có URL riêng chứa mã VBA từ một nguồn tin cậy, bạn có thể nhập trực tiếp URL đó vào đối số này.  

### 🚀 Cách sử dụng:  


| 🏷️ Trường hợp                     | 📝 Mô tả                                                                 | 💡 Cú pháp hàm                                          | 📊 Kết quả                              |
|----------------------------------|-------------------------------------------------------------------------|--------------------------------------------------------|----------------------------------------|
| ⌨️ Nhập công thức                | Nhập công thức để kích hoạt chức năng chèn macro.                        | =NNCT_CallMacro(...)                                   | Kích hoạt chức năng chèn macro         |
| 🖋️ Chèn macro mặc định           | Chèn danh sách macro mặc định.                                          | =NNCT_CallMacro(1)                                     | Chèn danh sách macro mặc định          |
| 🌐 Chèn macro từ URL riêng       | Chèn macro từ URL riêng được chỉ định.                                  | =NNCT_CallMacro("URL_của_bạn")                         | Chèn macro từ URL được chỉ định        |


Nhấn **Enter:** ✅ Khi bạn nhấn Enter, hàm sẽ thực thi quá trình tải và chèn mã VBA vào sổ làm việc của bạn.  

---

### 💡 Ví dụ:  


| 🏷️ Trường hợp                     | 📝 Mô tả                                                                 | 💡 Cú pháp hàm                                                      | 📊 Kết quả                              |
|----------------------------------|-------------------------------------------------------------------------|--------------------------------------------------------------------|----------------------------------------|
| 📜 Chèn danh sách macro cố định  | Chèn ba macro mặc định từ GitHub Gist của XuanAn2018.                   | =NNCT_CallMacro(1)                                                 | Chèn AIO_MyMacro.bas, QR_QuickChart.bas, QRcode_ChatGPT.bas    |
| 🌐 Chèn macro từ URL riêng       | Chèn macro từ URL riêng, ví dụ: MyMacro.bas từ GitHub Gist.            | =NNCT_CallMacro("https://gist.githubusercontent.com/username/123abc/raw/MyMacro.bas") | Chèn macro từ URL được chỉ định        |


> [!CAUTION]
>⚠️ Hãy thận trọng khi sử dụng các đoạn mã được tải về từ internet.  


### 📌 Lưu ý quan trọng:  

**Nguồn URL an toàn:** 🔒 Hàm chỉ chấp nhận URL từ các nguồn được tin cậy (hiện tại được cấu hình để chỉ cho phép từ [XuanAn2018](https://gist.githubusercontent.com/XuanAn2018/). *Nếu bạn nhập một URL không được tin cậy, hàm sẽ cảnh báo và không chèn mã để bảo vệ sổ làm việc của bạn.*  

**Làm sạch mã VBA:** 🧹 Mã VBA được tải về sẽ được làm sạch (ví dụ: loại bỏ các ký tự không mong muốn) trước khi chèn vào sổ làm việc để đảm bảo tính tương thích. 

**Kiểm tra trùng lặp:** 🔄 Hàm sẽ kiểm tra xem module có tên tương tự đã tồn tại trong sổ làm việc chưa. Nếu có, nó sẽ bỏ qua việc chèn mã trùng lặp để tránh ghi đè hoặc xung đột.  

**Kết nối Internet:** 🌐 Để tải mã VBA từ các URL, hàm cần có kết nối internet ổn định.  

> [!NOTE]
> ### 💬 Thông báo trả về:  
>`Success`: Nếu hàm thực hiện thành công việc chèn mã VBA.  
>`Error`: Kèm theo mô tả lỗi: Nếu gặp vấn đề trong quá trình thực thi (ví dụ: không thể kết nối URL, lỗi nội bộ).  
>`Error`: Invalid input type. Enter the formula `=NXA_HELP` to view instructions for inserting macros: Nếu kiểu dữ liệu của đối số đầu vào macroInput không hợp lệ (ví dụ: nhập văn bản không phải URL hoặc số khác 1).  


### 📖 Mã VBA hỗ trợ:  
**Hiện tại, hướng dẫn này chưa liệt kê chi tiết chức năng của các macro được đề cập (AIO_MyMacro.bas, QR_QuickChart.bas, QRcode_ChatGPT.bas). Để biết thêm thông tin về chức năng của từng macro, bạn cần tham khảo tài liệu riêng của tác giả [XuanAn2018](https://gist.githubusercontent.com/XuanAn2018/) hoặc kiểm tra trực tiếp mã nguồn của các macro đó trên [GitHub Gist](https://gist.githubusercontent.com/XuanAn2018/).**  

Với hàm `NNCT_CallMacro`, bạn có thể dễ dàng mở rộng khả năng của Excel bằng cách tự động thêm các chức năng VBA tùy chỉnh một cách an toàn và hiệu quả! 🚀
