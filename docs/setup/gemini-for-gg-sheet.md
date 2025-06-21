**Trước tiên bạn cần lấy [API keys in Google Cloud](https://makersuite.google.com/app/apikey)**  

### 🚀 Hướng dẫn Gán Mã JavaScript lên Google Apps Script  
- Để sử dụng đoạn mã bạn cung cấp, chúng ta cần đặt nó vào môi trường Google Apps Script. Dưới đây là các bước cụ thể:  

---

**1. 📂 Mở Google Apps Script Editor**  
  - Mở Google Sheet của bạn.  
 
  - Trên thanh menu, đi tới Extensions (Tiện ích mở rộng) ➡️ Apps Script.  
  
  - Một cửa sổ mới sẽ mở ra, đây chính là trình chỉnh sửa Google Apps Script của dự án hiện tại.  

---

**2. 📋 Dán Mã Code của bạn**  
  - Trong trình chỉnh sửa Apps Script, bạn sẽ thấy một tệp có tên Mã.gs (hoặc Code.gs) đã được tạo sẵn.  

  - Xóa bỏ mọi nội dung hiện có trong tệp này.  

  - Dán toàn bộ đoạn mã JavaScript bạn đã cung cấp vào đây:  


```
const API_KEY = 'YOUR_API_KEY'; 
const URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${API_KEY}`;

function askGPT(prompt) {
  const payload = {
    "contents":[
        {"role": "model",
         "parts":[{
           "text": "You are a helpful assisstant.  Yor name is J2TEAM GPT"}]},
        {"role": "user",
         "parts":[{
           "text": prompt }]
        },
    ]
  }

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(URL, options);
    const json = JSON.parse(response.getContentText());
    return json["candidates"][0]["content"]["parts"][0]["text"];
  } catch (e) {
    return `Error: ${e.message}`;
  }
}
```


**3. 🔑 Thay thế YOUR_API_KEY**  

- Trong đoạn mã bạn vừa dán, hãy tìm dòng const API_KEY = 'YOUR_API_KEY';.  

- Thay thế 'YOUR_API_KEY' bằng khóa API thực tế mà bạn đã lấy từ Google AI Studio (trước đây là Google Cloud Console). Đảm bảo giữ nguyên dấu nháy đơn (') xung quanh khóa API của bạn.  

---

**4. 💾 Lưu Dự án**  

- Nhấn vào biểu tượng Save project (hình đĩa mềm) hoặc tổ hợp phím Ctrl + S (Windows) / Cmd + S (Mac) để lưu dự án.  

- Bạn có thể đặt tên cho dự án nếu được yêu cầu (ví dụ: "Google Sheet AI Helper").  

---

**5. ▶️ Cấp quyền (Nếu được yêu cầu)**  

Lần đầu tiên bạn chạy một hàm sử dụng các dịch vụ của Google (như UrlFetchApp), Apps Script có thể yêu cầu bạn cấp quyền.  
Nếu có hộp thoại xuất hiện, hãy chọn Review permissions (Xem lại quyền), sau đó chọn tài khoản Google của bạn và cho phép các quyền cần thiết.  

---

**6. 📝 Sử dụng Hàm trong Google Sheet**  
- Sau khi đã thiết lập xong, bạn có thể gọi hàm askGPT trực tiếp trong bất kỳ ô nào trên Google Sheet của mình:  

---

**Để hỏi một câu hỏi:**  

| 🏷️ Trường hợp                     | 📝 Mô tả                                                         | 💡 Cú pháp hàm                     | 📊 Kết quả                              |
|----------------------------------|-----------------------------------------------------------------|------------------------------------|----------------------------------------|
| ❓ Hỏi trực tiếp                  | Hỏi thủ đô của Việt Nam trực tiếp trong ô.                      | `=askGPT("Thủ đô của Việt Nam")`     | Thủ đô của Việt Nam là Hà Nội         |
| 🔗 Hỏi từ ô khác                 | Tham chiếu câu hỏi từ ô A1, gọi API Gemini trả kết quả.         | =askGPT(A1)                        | Kết quả từ câu hỏi trong ô A1         |

---

// Author: [JUNO_OKYO - J2TEAM](https://gist.github.com/J2TEAM/91bc15ab65941d8db9cfb30de2b849a3)  
// Recreated by [orinn2k7](https://gist.github.com/orius2k7/21bb709118e264e48277f66fd222a70d)  
