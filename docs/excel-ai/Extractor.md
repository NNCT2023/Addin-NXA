**📚 Hướng Dẫn Sử Dụng Hàm NXA_Extractor**  
### 🌟 Chức năng:  
- Hàm `NXA_Extractor` 🔍 sử dụng mô hình ngôn ngữ lớn Gemini (trước đây là GPT-3) của Google AI để trích xuất thông tin theo từ khóa cụ thể từ một đoạn văn bản trong Excel 📝. Hàm này giúp bạn tự động hóa việc thu thập dữ liệu quan trọng một cách hiệu quả.  

**📝 Cú pháp:**  

`=NXA_Extractor(prompt As String, keyword As String)`  

  - ⚙️ Các đối số:  
      - `prompt (Bắt buộc)`: 🔑 Kiểu dữ liệu String. Đây là đoạn văn bản chứa thông tin cần trích xuất. Bạn có thể nhập trực tiếp chuỗi văn bản hoặc tham chiếu đến một ô chứa văn bản đó.  
      - `keyword (Bắt buộc)`: 🔑 Kiểu dữ liệu String. Đây là từ khóa dùng để xác định loại thông tin cần trích xuất (ví dụ: "name" để trích xuất tên, "place" cho địa điểm, "organization" cho tổ chức).  

---

### 💡 Ví dụ:  
  - Giả sử ô A1 của bạn chứa đoạn văn bản: "John Doe is the CEO of Acme Corp. located in New York City."  

    - **Ví dụ 1**: Trích xuất tên người  
      Để trích xuất tên người từ đoạn văn bản trong ô A1, bạn sử dụng công thức:  

      - `=NXA_Extractor(A1, "name")`  
      - Kết quả trả về: "Chí Công"  

  - **Ví dụ 2:** Trích xuất địa điểm  
      Để trích xuất địa điểm từ đoạn văn bản trong ô A1, bạn sử dụng công thức:  

     - `=NXA_Extractor(A1, "place")`  
     - Kết quả trả về: "Hanoi City"  

    - **Ví dụ 3:** Trích xuất tổ chức  
      Để trích xuất tên tổ chức từ đoạn văn bản trong ô A1, bạn sử dụng công thức:  

    - `=NXA_Extractor(A1, "organization")`  
      Kết quả trả về: "Acme Corp."  

---

### ⚠️ Hãy thận trọng khi sử dụng các đoạn mã.  

---

### 📌 Lưu ý quan trọng:  
**Bảo mật API Key:** 🔑 Hàm này sử dụng [API Key của Google Cloud Platform](https://aistudio.google.com/app/apikey). API Key được minh họa trong hàm chỉ mang tính chất ví dụ và không nên sử dụng trực tiếp. Bạn cần có API Key riêng của mình để đảm bảo an toàn và kích hoạt chức năng.  

**Kiểm tra đầu vào:** ✅ Hàm sẽ tự động kiểm tra tính hợp lệ của đầu vào: đoạn văn bản và từ khóa không được để trống.  
      -Kết quả trả về: 📤  
        + Thông tin được trích xuất theo từ khóa mà bạn cung cấp.  
        + Lỗi nếu không tìm thấy thông tin theo từ khóa (ví dụ: "`No name found`" nếu không có tên nào được phát hiện).  
        + Lỗi nếu gặp vấn đề trong quá trình trích xuất hoặc kết nối API (ví dụ: `lỗi mạng`, `lỗi API Key`).  
        
**Dịch vụ có phí:** 💰 Hàm này phụ thuộc vào dịch vụ Google Cloud Platform có trả phí. Hãy kiểm tra chi phí sử dụng API của Gemini trên trang chủ Google Cloud Platform.  

---

### 📉 Các hạn chế:  

**Độ chính xác:** 🎯 Độ chính xác của kết quả trích xuất phụ thuộc vào chất lượng của đoạn văn bản đầu vào và khả năng hiểu ngữ cảnh của mô hình ngôn ngữ. Văn bản càng rõ ràng, kết quả càng chính xác.  
**Loại thông tin đơn:** 🔢 Hàm này hiện chỉ hỗ trợ trích xuất một loại thông tin theo từ khóa đơn tại một thời điểm.  

---

### 🛠️ Xử lý lỗi:  

  - **Lỗi API Key hoặc kết nối:** 🌐 Nếu gặp lỗi liên quan đến API Key hoặc kết nối, bạn cần kiểm tra lại API Key của mình (đảm bảo nó hợp lệ và đã được cấu hình đúng) và kiểm tra kết nối internet để đảm bảo ổn định.  

**Lỗi chất lượng văn bản:** ✍️ Nếu gặp lỗi do chất lượng văn bản đầu vào (ví dụ: thông tin không rõ ràng, từ khóa không khớp), hãy kiểm tra lại đoạn văn bản để đảm bảo nó rõ ràng và đầy đủ thông tin cần trích xuất.  

> Với hàm `NXA_Extractor`, bạn có thể tự động hóa việc trích xuất các loại thông tin cụ thể từ văn bản trực tiếp trong Excel, giúp tăng cường hiệu quả công việc và xử lý dữ liệu thông minh hơn! 🚀  
