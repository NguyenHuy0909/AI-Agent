# Hướng Dẫn Sử Dụng Folder VBA Project Development Framework
## Quy Trình Phát Triển Dự Án VBA Với AI Agent (Tiếng Việt)

**Ngày tạo:** 18 Tháng 4, 2026
**Phiên bản:** 1.0
**Trạng thái:** Sẵn sàng sử dụng

---

## 📁 CÓ CÁI GÌ TRONG FOLDER NÀY?

Folder này chứa **5 tài liệu hướng dẫn hoàn chỉnh** để giúp bạn xây dựng dự án VBA trong Excel một cách có hệ thống:

### 1️⃣ **README.md** - Tổng Quan Chung
- **Dùng để:** Hiểu rõ framework là gì
- **Nội dung:** Giới thiệu từng document, cách sử dụng chúng
- **Khi nào dùng:** Lần đầu tiên tìm hiểu về framework
- **Thời gian đọc:** 10 phút

### 2️⃣ **QUICK_START_GUIDE.md** ⭐ QUAN TRỌNG NHẤT
- **Dùng để:** Bắt đầu dự án của bạn trong 24 giờ đầu
- **Nội dung:** 
  - 5 bước đầu tiên để khởi động
  - Ví dụ cụ thể
  - Timeline chi tiết
  - Các mẫu câu hỏi để hỏi AI Agent
- **Khi nào dùng:** Khi bạn bắt đầu dự án mới
- **Thời gian đọc:** 15 phút + 4-5 giờ thực hiện

### 3️⃣ **VBA_AI_WORKFLOW.md** - Quy Trình Chi Tiết
- **Dùng để:** Tham khảo toàn bộ quy trình 7 giai đoạn
- **Nội dung:**
  - Giai đoạn 1: Yêu cầu & Phân tích (2 ngày)
  - Giai đoạn 2: Thiết kế & Kiến trúc (1-2 ngày)
  - Giai đoạn 3: Phát triển (4-5 ngày)
  - Giai đoạn 4: Tích hợp (1 ngày)
  - Giai đoạn 5: Kiểm thử (1-2 ngày)
  - Giai đoạn 6: Tối ưu hóa (1 ngày)
  - Giai đoạn 7: Hoàn thành (1-2 ngày)
- **Khi nào dùng:** Khi cần hướng dẫn chi tiết cho từng giai đoạn
- **Thời gian đọc:** 30 phút

### 4️⃣ **PROJECT_CHECKLIST.md** ✅ THEO DÕI TIẾN ĐỘ
- **Dùng để:** Ghi lại tiến độ dự án của bạn
- **Nội dung:**
  - Danh sách kiểm tra cho mỗi giai đoạn
  - Bảng theo dõi module
  - Danh sách lỗi (Bug Tracker)
  - Thống kê dự án
- **Khi nào dùng:** Mỗi ngày khi làm việc trên dự án
- **Cách dùng:** Điền thông tin dự án của bạn và tích vào các ô khi hoàn thành

### 5️⃣ **MODULE_TEMPLATE_AND_REFERENCE.md** 📚 TÀI LIỆU THAM KHẢO VBA
- **Dùng để:** Tra cứu cú pháp VBA, template module
- **Nội dung:**
  - Mẫu module hoàn chỉnh (copy-paste ready)
  - Hướng dẫn nhanh VBA (biến, hàm, vòng lặp, xử lý lỗi)
  - Qui ước đặt tên
  - Mẹo gỡ lỗi
  - Câu hỏi thường gặp
- **Khi nào dùng:** Khi viết code VBA hoặc cần trợ giúp cú pháp
- **Dùng để tìm:** Cách làm việc với Excel Range, xử lý lỗi, vòng lặp, v.v.

---

## 🎯 BẠN SHOULD START ĐẤY - HƯỚNG DẪN TỪNG BƯỚC

### Bước 1: Hiểu Framework (15 phút)
**Đọc các tài liệu theo thứ tự này:**
1. README.md (lần này chỉ để hiểu rõ)
2. QUICK_START_GUIDE.md (phần "GETTING STARTED: YOUR FIRST 24 HOURS")

**Mục đích:** Bạn sẽ hiểu framework là gì và cách bắt đầu

### Bước 2: Chuẩn Bị Dự Án (30 phút)
**Từ QUICK_START_GUIDE.md, hoàn thành các bước:**
1. **Step 1: Định Nghĩa Dự Án của Bạn (30 phút)**
   - Viết đó là dự án VBA của bạn sẽ làm gì
   - Liệt kê 3-5 chức năng chính
   - Xác định nguồn dữ liệu và nơi lưu kết quả
   - Xác định tiêu chí thành công

**Ví dụ:**
```
Dự Án: Xử Lý Dữ Liệu Nhân Viên
Mục Đích: Đọc dữ liệu nhân viên từ Excel, kiểm tra tính hợp lệ, 
          tính bonus, viết kết quả lại

Chức Năng Chính:
1. Đọc dữ liệu từ Sheet1
2. Kiểm tra email và định dạng điện thoại
3. Tính bonus hiệu suất (5-15% dựa trên điểm số)
4. Viết kết quả vào Sheet2

Nguồn Dữ Liệu: Sheet1 (Name, Email, Phone, Score)
Kết Quả: Sheet2 (Name, Email, Phone, Bonus Amount, Status)

Tiêu Chí Thành Công:
- Xử lý 1000 nhân viên trong 10 giây
- Kiểm tra tất cả dữ liệu chính xác
- Hiển thị thông báo lỗi nếu dữ liệu không hợp lệ
```

### Bước 3: Hỏi AI Agent Về Kiến Trúc (30 phút)
**Từ QUICK_START_GUIDE.md, hoàn thành:**
2. **Step 2: Ask AI Agent for Architecture Recommendations**
   - Copy mẫu câu hỏi từ tài liệu
   - Điền thông tin dự án của bạn
   - Gửi cho AI Agent

**AI Agent sẽ giúp bạn:**
- Xác định nên tạo bao nhiêu module
- Cách tổ chức kiến trúc
- Chiến lược xử lý lỗi
- Các best practice VBA

### Bước 4: Tạo Danh Sách Module (15 phút)
**Hoàn thành:**
3. **Step 3: Create Your Module List**
   - Dựa trên khuyến nghị của AI Agent
   - Tạo bảng danh sách module (có mẫu trong QUICK_START_GUIDE.md)
   - Ghi rõ tên module, loại, mục đích

### Bước 5: Tạo Module Đầu Tiên (2-3 giờ)
**Hoàn thành:**
4. **Step 4: Generate First Module with AI Agent (1 hour)**
   - Bắt đầu với module đơn giản nhất (thường là Config hoặc Utility)
   - Hỏi AI Agent để tạo code
   - Kiểm tra code
   - Thêm vào Excel
   - Kiểm thử

5. **Step 5: Test Your Modules (1 hour)**
   - Tạo code kiểm thử
   - Chạy kiểm thử
   - Sửa lỗi nếu có

**Kết quả:** Bạn sẽ có module đầu tiên hoạt động!

---

## 📋 CÁCH DÙNG TỪNG TÀI LIỆU TRONG CÔNG VIỆC HÀNG NGÀY

### Thứ Hai - Bắt Đầu Dự Án Mới

**Buổi Sáng:**
1. Mở **QUICK_START_GUIDE.md** (15 phút)
2. Hoàn thành Bước 1-3 (1 giờ)
3. Ghi lại thông tin dự án trong **PROJECT_CHECKLIST.md** (15 phút)

**Buổi Chiều:**
1. Hỏi AI Agent về kiến trúc (30 phút)
2. Tạo danh sách module (15 phút)
3. Bắt đầu module đầu tiên (2-3 giờ)

### Ngày Tiếp Theo - Phát Triển Tiếp

**Mỗi Module:**
1. Sử dụng **MODULE_TEMPLATE_AND_REFERENCE.md** làm mẫu
2. Hỏi AI Agent để tạo code
3. Kiểm thử code
4. Cập nhật **PROJECT_CHECKLIST.md** để ghi nhận hoàn thành

**Khi Gặp Vấn Đề:**
1. Tra cứu **MODULE_TEMPLATE_AND_REFERENCE.md** để tìm giải pháp
2. Hỏi AI Agent với mã lỗi cụ thể
3. Sửa và kiểm thử lại

### Khi Hoàn Thành Mỗi Giai Đoạn

1. Đọc phần tương ứng trong **VBA_AI_WORKFLOW.md**
2. Kiểm tra danh sách công việc trong **PROJECT_CHECKLIST.md**
3. Chuyển sang giai đoạn tiếp theo

---

## 🗂️ CÁCH TỔCHỨC CÁC MODULE

### Qui ước đặt tên folder/module:

```
Dự Án VBA của Bạn (Folder Excel)
│
├── 00_Config
│   └── ConfigConstants          (Hằng số, cấu hình)
│
├── 01_Utilities
│   ├── UtilityValidation        (Hàm kiểm tra)
│   ├── UtilityFormatting        (Hàm định dạng)
│   └── UtilityString            (Hàm xử lý chuỗi)
│
├── 02_DataAccess
│   ├── DataReader               (Đọc dữ liệu Excel)
│   └── DataWriter               (Ghi dữ liệu Excel)
│
├── 03_BusinessLogic
│   ├── Calculation              (Tính toán chính)
│   ├── Processing               (Xử lý logic chính)
│   └── Validation               (Kiểm tra business rule)
│
├── 04_Integration
│   └── UserInterface            (Tương tác worksheet)
│
└── 99_Main
    └── Orchestration            (Điểm vào chính)
```

### Tên hàm:

```
Mẫu: [Động Từ][Chủ Thể]AsKiểuDữ Liệu

Ví dụ:
- ValidateEmailAsBoolean()       (Trả về Boolean)
- CalculateTotalAsDouble()       (Trả về Double)
- GetActiveSheetAsObject()       (Trả về Object)
- FormatDateAsString()           (Trả về String)
```

---

## 💬 CÁCH HỎI AI AGENT

### Mẫu 1: Tạo Module
```
Module: [Tên Module]
Mục đích: [Module làm gì]

Hàm cần tạo:
1. FunctionName(param1 As String) As Boolean
   Mục đích: [Hàm làm gì]
   Trường hợp lỗi: [Có thể gặp lỗi gì]

2. AnotherFunction(data As Collection) As Integer
   Mục đích: [Hàm làm gì]
   Trường hợp lỗi: [Có thể gặp lỗi gì]

Yêu cầu:
- Module hoàn chỉnh với Option Explicit và comment
- Xử lý lỗi toàn diện cho mỗi hàm
- Ví dụ cách sử dụng
```

### Mẫu 2: Kiểm Tra Code
```
Vui lòng kiểm tra code này xem có vấn đề gì không:
- Bug hoặc lỗi tiềm ẩn
- Không tuân theo VBA best practice
- Vấn đề hiệu suất
- Thiếu xử lý lỗi

[Dán code của bạn ở đây]

Câu hỏi:
- Code có đúng không?
- Hiệu suất có chấp nhận được không?
- Có thể cải thiện như thế nào?
```

### Mẫu 3: Gỡ Lỗi
```
Tôi gặp lỗi này: [Thông báo lỗi]
Tại dòng: [Số dòng]
Khi tôi: [Bạn đang làm gì]

Code:
[Dán đoạn code liên quan]

Câu hỏi:
- Nguyên nhân lỗi là gì?
- Cách sửa?
- Cách phòng tránh trong tương lai?
```

### Mẫu 4: Tối Ưu Hóa
```
Hàm này quá chậm: [Tên hàm]
Hiệu suất hiện tại: [Mất X giây cho Y bản ghi]
Hiệu suất mong muốn: [Nên mất Z giây]

Code hiện tại:
[Dán code]

Câu hỏi:
- Tại sao chậm?
- Cách tối ưu?
- Cách tiếp cận tốt nhất?
```

---

## ⏱️ TIMELINE TỰA TẠO HOÀN THÀNH TRONG 8 NGÀY

| Ngày | Giai Đoạn | Công Việc | Kết Quả |
|------|-----------|----------|--------|
| 1 | Yêu Cầu & Phân Tích | Định nghĩa dự án, hỏi AI về kiến trúc | Module list hoàn thành |
| 2-3 | Thiết Kế | Tạo ConfigConstants + Utility module | 2 module hoạt động |
| 3-4 | Phát Triển | Tạo DataAccess module | Module đọc/ghi Excel hoạt động |
| 5-6 | Phát Triển | Tạo Business Logic module | Logic chính hoạt động |
| 7 | Tích Hợp & Kiểm Thử | Gộp tất cả vào Excel, kiểm thử toàn bộ | Workflow hoàn chỉnh hoạt động |
| 8 | Tối Ưu & Hoàn Thành | Cải thiện, documentation, giao hàng | Dự án hoàn thành sẵn sàng sử dụng |

---

## ✅ DANH SÁCH KIỂM TRA NHANH

### Trước Khi Bắt Đầu
- [ ] Đọc QUICK_START_GUIDE.md
- [ ] Hiểu 5 bước đầu tiên
- [ ] Chuẩn bị thông tin dự án

### Khi Bắt Đầu Dự Án
- [ ] Điền thông tin vào PROJECT_CHECKLIST.md
- [ ] Tạo danh sách module
- [ ] Hỏi AI Agent về kiến trúc
- [ ] Nhận khuyến nghị

### Khi Tạo Mỗi Module
- [ ] Sử dụng MODULE_TEMPLATE_AND_REFERENCE.md làm mẫu
- [ ] Hỏi AI Agent để tạo code
- [ ] Kiểm tra code tạo ra
- [ ] Kiểm thử module
- [ ] Cập nhật PROJECT_CHECKLIST.md

### Khi Hoàn Thành Giai Đoạn
- [ ] Kiểm tra danh sách công việc trong VBA_AI_WORKFLOW.md
- [ ] Đánh dấu hoàn thành trong PROJECT_CHECKLIST.md
- [ ] Chuyển sang giai đoạn tiếp theo

### Khi Gặp Vấn Đề
- [ ] Tra cứu MODULE_TEMPLATE_AND_REFERENCE.md
- [ ] Hỏi AI Agent với chi tiết cụ thể
- [ ] Kiểm thử giải pháp
- [ ] Ghi lại bài học

---

## 🎓 VÍ DỤ THỰC TẾ: DỰ ÁN TÍNH BONUS NHÂN VIÊN

### Ngày 1: Bắt Đầu

**Sáng (1 giờ):**
```
Dự Án: Tính Bonus Nhân Viên
Mục Đích: Đọc dữ liệu nhân viên, tính bonus 5-15% dựa trên điểm, xuất kết quả

Yêu Cầu:
- Đọc từ Sheet "Nhân Viên"
- Kiểm tra email và điện thoại hợp lệ
- Tính bonus: điểm < 50 = 5%, 50-75 = 10%, > 75 = 15%
- Ghi kết quả vào Sheet "Kết Quả"
- Xử lý 1000 nhân viên trong 10 giây
```

**Chiều (2 giờ):**
```
Hỏi AI Agent:
"Dựa trên yêu cầu này, tôi nên tạo bao nhiêu module?
Kiến trúc nên như thế nào?"

AI Agent trả lời:
"Tạo 5 module:
1. ConfigConstants - Hằng số, % bonus, tên sheet
2. UtilityValidation - Kiểm tra email, điện thoại
3. DataAccess - Đọc/ghi dữ liệu
4. Calculation - Tính bonus
5. Main - Điều phối toàn bộ"
```

### Ngày 2: Tạo Module Đầu Tiên

**Sáng (1 giờ):**
```
Hỏi AI Agent:
"Tạo ConfigConstants module với:
- Hằng số % bonus: MIN = 5%, MID = 10%, MAX = 15%
- Điểm ngưỡng: SCORE_LOW = 50, SCORE_HIGH = 75
- Tên sheet: INPUT_SHEET = 'Nhân Viên', OUTPUT_SHEET = 'Kết Quả'
- Email domain được phép: company.com, backup.company.com

Yêu cầu: Option Explicit, comments, đặt tên theo CONST_DESCRIPTION_TYPE"
```

**Chiều (2 giờ):**
```
Hỏi AI Agent:
"Tạo UtilityValidation module với 2 hàm:
1. ValidateEmailAsBoolean(email As String)
   - Kiểm tra format email hợp lệ
   - Kiểm tra domain trong danh sách được phép
   - Trả về True nếu OK

2. ValidatePhoneAsBoolean(phone As String)
   - Kiểm tra phone có chỉ số và ký tự đặc biệt
   - Phải có ít nhất 10 chữ số
   - Trả về True nếu OK

Bao gồm: Xử lý lỗi, comment chi tiết, ví dụ"
```

### Ngày 3: Tiếp Tục Phát Triển

```
Hỏi AI Agent:
"Tạo DataAccess module:
1. ReadEmployeeDataAsCollection(sheetName) As Collection
   - Đọc từ sheet (Name, Email, Phone, Score)
   - Trả về Collection với data
   
2. WriteResultsToSheet(sheetName, data As Collection) As Boolean
   - Ghi Collection vào sheet

Bao gồm: Xử lý lỗi, kiểm tra sheet tồn tại"
```

### Ngày 4: Xử Lý Chính

```
Hỏi AI Agent:
"Tạo Calculation module:
1. CalculateBonusAsDouble(score As Double) As Double
   - Nếu score < 50: return 5% của lương
   - Nếu 50 <= score <= 75: return 10%
   - Nếu score > 75: return 15%
   - Lương cơ bản = 1000 (từ constant)

Bao gồm: Xử lý lỗi, validation input"
```

### Ngày 5-6: Gộp Module

```
Tạo Main module:
1. RunBonusProcessing() As Boolean
   - Gọi DataAccess.ReadEmployeeData()
   - Với mỗi employee:
     - Validate email & phone
     - Calculate bonus
   - Ghi kết quả lại
   
Hỏi AI Agent tạo code này
```

### Ngày 7: Kiểm Thử

```
Tạo test code:
Sub TestBonusSystem()
    Dim result As Boolean
    result = Main.RunBonusProcessing()
    
    If result Then
        MsgBox "Thành công! Kiểm tra sheet 'Kết Quả'"
    Else
        MsgBox "Có lỗi. Kiểm tra Immediate Window"
    End If
End Sub

Chạy test, sửa lỗi nếu có
```

### Ngày 8: Hoàn Thành

```
- Tối ưu hóa hiệu suất
- Viết tài liệu hướng dẫn người dùng
- Kiểm thử toàn bộ
- Giao hàng cho người dùng
```

---

## 🆘 CÂU HỎI THƯỜNG GẶP

### Q: Tôi nên bắt đầu từ đâu?
**A:** Mở QUICK_START_GUIDE.md và theo 5 bước đó. Đừng vội vàng.

### Q: Module nào nên tạo trước?
**A:** Thứ tự: Config → Utility → DataAccess → BusinessLogic → Integration → Main

### Q: Nên hỏi AI Agent trong code không?
**A:** Có! Code do AI tạo ra rất tốt. Kiểm tra kỹ trước khi dùng.

### Q: Tôi nên kiểm thử như thế nào?
**A:** Test từng module riêng lẻ trước, sau đó test toàn bộ workflow.

### Q: Nếu code từ AI có lỗi?
**A:** Bình thường. Hỏi lại AI với lỗi cụ thể, AI sẽ sửa.

### Q: Tôi có thể tái sử dụng module không?
**A:** Hoàn toàn! Các module Utility thường dùng lại được.

### Q: Project này sẽ mất bao lâu?
**A:** Khoảng 8 ngày từ yêu cầu đến hoàn thành.

### Q: Tôi có thể tối ưu hóa sau không?
**A:** Có! Giai đoạn 6 là dành cho tối ưu hóa.

---

## 📌 CÁC ĐIỂM QUAN TRỌNG CẦN NHỚ

✅ **Xây dựng module từ từ** - Không làm tất cả một lúc

✅ **Kiểm thử khi đi** - Không chờ đến cuối

✅ **Kiểm tra code AI tạo** - Luôn xem qua trước khi dùng

✅ **Hỏi cụ thể** - Câu hỏi cụ thể được câu trả lời cụ thể

✅ **Lập tài liệu** - Ghi lại những quyết định quan trọng

✅ **Theo dõi tiến độ** - Sử dụng PROJECT_CHECKLIST.md hàng ngày

✅ **Hỏi khi bị stuck** - AI Agent là người giúp tốt nhất

---

## 🚀 BƯỚC TIẾP THEO

**Ngay bây giờ:**
1. [ ] Mở **QUICK_START_GUIDE.md**
2. [ ] Đọc phần "GETTING STARTED: YOUR FIRST 24 HOURS"
3. [ ] Hoàn thành Bước 1 (định nghĩa dự án của bạn)

**Trong vòng 1 giờ:**
4. [ ] Hoàn thành Bước 2-3
5. [ ] Hỏi AI Agent

**Trong vòng 4-5 giờ:**
6. [ ] Tạo module đầu tiên
7. [ ] Kiểm thử nó
8. [ ] Bạn sẽ có code hoạt động! 🎉

---

## 📞 TÓM TẮT NHANH

| Cần Làm | File Nào |
|--------|----------|
| Bắt đầu dự án mới | QUICK_START_GUIDE.md |
| Hiểu framework | README.md |
| Tham khảo quy trình 7 giai đoạn | VBA_AI_WORKFLOW.md |
| Theo dõi tiến độ | PROJECT_CHECKLIST.md |
| Tra cứu VBA, xem mẫu module | MODULE_TEMPLATE_AND_REFERENCE.md |
| Hướng dẫn này | HUONG_DAN_VIET.md (file này) |

---

**Bây giờ bạn đã sẵn sàng! Hãy bắt đầu với QUICK_START_GUIDE.md ngay.**

**Chúc bạn thành công! 🌟**
