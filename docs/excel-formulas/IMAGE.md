**📚 Hướng Dẫn Sử Dụng Hàm NNCT_IMAGE**  

### 🌟 Chức năng:  
Hàm `NNCT_IMAGE` 🖼️ là một công cụ tiện lợi giúp bạn dễ dàng chèn hình ảnh từ một URL trực tiếp vào một ô tính trong Excel. Hàm này tự động xử lý toàn bộ quá trình: tải ảnh từ web, định vị ảnh trong ô, và thậm chí còn có thể tự động điều chỉnh kích thước ảnh theo ý muốn của bạn.  

### 📝 Cú pháp:  

```=NNCT_IMAGE(url, [autoFit], [height], [width])```  

---

### ⚙️ Các đối số:  
  - `url (Bắt buộc)`: 🔑 Kiểu dữ liệu String. Đây là địa chỉ URL đầy đủ của hình ảnh mà bạn muốn chèn vào Excel.  
  - `autoFit (Tùy chọn)`: 📏 Kiểu dữ liệu Boolean (mặc định là TRUE).  
  - `TRUE`: Hàm sẽ tự động điều chỉnh kích thước ảnh sao cho vừa với kích thước của ô mà bạn nhập công thức vào.  
  - `FALSE`: Hàm sẽ giữ kích thước gốc của ảnh khi chèn.  
  - `height (Tùy chọn)`: ↕️ Kiểu dữ liệu Double. Chiều cao mong muốn của ảnh (tính bằng pixel). Đối số này chỉ có tác dụng khi autoFit được đặt là `FALSE`.  
  - `width (Tùy chọn)`: ↔️ Kiểu dữ liệu Double. Chiều rộng mong muốn của ảnh (tính bằng pixel). Đối số này cũng chỉ có tác dụng khi autoFit được đặt là `FALSE`.  

---

### 🚀 Cách sử dụng:  
- **Nhập công thức:** ⌨️ Trong ô bạn muốn hình ảnh hiển thị, nhập công thức `=NNCT_IMAGE(...)`.  
  - **Thay thế các đối số:** 🖋️  
    - `url`: Nhập địa chỉ URL của hình ảnh bạn muốn chèn (ví dụ: `"https://www.example.com/logo.png"`).  
    - `autoFit`: Bạn có thể giữ mặc định là `TRUE` để ảnh tự động co giãn, hoặc nhập `FALSE` nếu bạn muốn kiểm soát kích thước thủ công.  
    - `height và width (tùy chọn)`: Chỉ điền hai đối số này khi autoFit là `FALSE`. Nhập chiều cao và chiều rộng mong muốn của ảnh tính bằng pixel (ví dụ: 200, 300).  
- **Nhấn Enter:** ✅ Sau khi nhập công thức, nhấn Enter để chạy hàm. Excel sẽ tải ảnh và chèn vào ô.  

---

### 💡 Ví dụ:  


| 🏷️ Trường hợp                     | 📝 Mô tả                                                                 | 💡 Cú pháp hàm                                                  | 📊 Kết quả                              |
|----------------------------------|-------------------------------------------------------------------------|----------------------------------------------------------------|----------------------------------------|
| 🖼️ Chèn ảnh tự động điều chỉnh    | Chèn ảnh từ URL vào ô A1, tự động điều chỉnh kích thước.                | =NNCT_IMAGE("`https://www.example.com/logo.png`")                | Ảnh tự động vừa với ô                 |
| 📏 Chèn ảnh kích thước cố định    | Chèn ảnh từ URL với chiều cao 150px, chiều rộng 200px.                 | =NNCT_IMAGE("`https://www.example.com/product.jpg`", `FALSE`, `150`, `200`) | Ảnh kích thước `150x200` pixel          |


### 📌 Lưu ý quan trọng:  
- **Kết nối Internet:** 🌐 Để hàm có thể tải ảnh từ URL, bạn cần đảm bảo có kết nối internet ổn định. Nếu không có internet, ảnh sẽ không tải được.  
- **Kích thước ảnh lớn:** ⚠️ Hình ảnh có kích thước tệp lớn có thể làm chậm quá trình tải và hiển thị trong Excel, ảnh hưởng đến hiệu suất bảng tính.  
- **Kích thước ô:** 📐 Nếu kích thước của ô quá nhỏ so với ảnh (đặc biệt khi `autoFit` là `FALSE`), ảnh có thể bị cắt hoặc méo mó. Hãy điều chỉnh kích thước ô để ảnh hiển thị đẹp nhất.  

### 🎨 Các trường hợp sử dụng khác:  
- Điều chỉnh kích thước ảnh theo ý muốn: Bạn có thể sử dụng các đối số `height` và `width` để đặt kích thước cụ thể cho ảnh, giúp bạn kiểm soát hoàn toàn việc hiển thị hình ảnh.  
- Chèn ảnh vào ô đã được gộp (`merged cells`): Hàm sẽ tự động nhận diện và điều chỉnh vị trí và kích thước ảnh sao cho phù hợp với ô đã được gộp, giúp bạn dễ dàng sắp xếp bố cục.  

---

### 🎯 Tóm tắt:  

Hàm `NNCT_IMAGE` là một công cụ hữu ích và tiện lợi để chèn hình ảnh từ URL vào Excel một cách nhanh chóng và dễ dàng. Nó giúp tiết kiệm thời gian so với việc chèn ảnh thủ công và cung cấp tùy chọn điều chỉnh kích thước linh hoạt, giúp bảng tính của bạn trở nên trực quan và sinh động hơn. 🌟  
