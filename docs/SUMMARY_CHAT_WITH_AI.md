# Summary: Hướng Dẫn Chat Hiệu Quả Với AI Agent
## Gom Lại Những Điều Quan Trọng Từ Cuộc Trao Đổi

**Ngày tạo:** 18 Tháng 4, 2026  
**Mục đích:** Tài liệu tham khảo khi chat với AI để build VBA Project  
**Loại:** Quick Reference Guide

---

## 📌 PHẦN 1: CÁCH SỬ DỤNG FOLDER FRAMEWORK

### 6 File Trong Folder:

| # | File | Tiếng Việt | Mục Đích | Khi Dùng |
|---|------|-----------|---------|---------|
| 1 | README.md | ❌ | Tổng quan framework | Lần đầu hiểu framework |
| 2 | QUICK_START_GUIDE.md | ❌ | Bắt đầu 24h đầu | Khi bắt đầu dự án |
| 3 | VBA_AI_WORKFLOW.md | ❌ | 7 giai đoạn chi tiết | Tham khảo phase hiện tại |
| 4 | PROJECT_CHECKLIST.md | ✅ | Theo dõi tiến độ | Mỗi ngày ghi lại |
| 5 | MODULE_TEMPLATE_AND_REFERENCE.md | ❌ | VBA templates & ref | Tra cứu cú pháp VBA |
| 6 | HUONG_DAN_TIENG_VIET.md | ✅ | Hướng dẫn Tiếng Việt | Tất cả mọi lúc |

### Cách Sử Dụng Đúng:

```
📖 FILE .md = TÀI LIỆU THAM KHẢO (để học cách hỏi)
💬 CHAT NÀY = NƠI THỰC HIỆN (gõ yêu cầu thực tế ở đây)
```

**Quy trình:**
1. Đọc mẫu trong file .md
2. Điều chỉnh với dữ liệu cụ thể
3. **Gõ yêu cầu TRỰC TIẾP vào chat** (tại đây)
4. Tôi tạo code/trả lời

---

## 📌 PHẦN 2: CÁCH CUNG CẤP CONTEXT

### ❌ SAI - Không Đủ Context:

```
"Tạo module XYZ"
```
→ AI không biết:
- Phase nào?
- Dự án gì?
- Constraints gì?

### ✅ ĐÚNG - Context Đầy Đủ:

```
[PHASE X - TÊN PHASE]

Dự án: [Tên dự án]
Module: [Tên module hiện tại]

[Chi tiết yêu cầu cụ thể]
```

---

## 📌 PHẦN 3: PHASE SELECTION

### 7 Giai Đoạn Của Framework:

| Phase | Tên | Thời Gian | Nhiệm Vụ |
|-------|-----|----------|---------|
| **1** | Requirements & Analysis | 2 ngày | Định nghĩa yêu cầu, kiến trúc |
| **2** | Design & Architecture | 1-2 ngày | Thiết kế module, function specs |
| **3** | Implementation | 4-5 ngày | Code module, fix bug, test |
| **4** | Integration & Assembly | 1 ngày | Gộp module vào Excel |
| **5** | Testing & QA | 1-2 ngày | Kiểm thử toàn bộ |
| **6** | Optimization & Refinement | 1 ngày | Tối ưu, cải tiến |
| **7** | Final Delivery & Deployment | 1-2 ngày | Hoàn thành, giao hàng |

### Khi Nào Chọn Phase Nào?

```
✅ Yêu cầu "tạo module mới" → [PHASE 3 - IMPLEMENTATION]
✅ Yêu cầu "tạo code" → [PHASE 3 - IMPLEMENTATION]
✅ Yêu cầu "cải tiến code" → [PHASE 6 - OPTIMIZATION]
✅ Yêu cầu "kiểm tra thiết kế" → [PHASE 2 - DESIGN]
✅ Code có bug → [PHASE 3 - BUG FIX] hoặc [PHASE 5 - TESTING]
```

---

## 📌 PHẦN 4: TEMPLATE YÊUCẦU TỐI ƯU

### Template Chuẩn (Tiết Kiệm Token):

```
[PHASE X - TÊN PHASE - LOẠI YÊU CẦU]

Dự án: [Tên dự án]
Module: [Tên module - nếu có]

VẤN ĐỀ/YÊU CẦU:
[Mô tả 1-2 câu]

KỲ VỌNG vs THỰC TẾ (nếu bug):
- Kỳ vọng: [Output dự kiến]
- Thực tế: [Output hiện tại]

DỮ LIỆU MẪU (5-10 dòng):
[Dữ liệu mẫu]

CODE (chỉ phần liên quan, 15-30 dòng):
[Đoạn code]

TRIỆU CHỨNG:
[Ghi rõ vấn đề]

YÊU CẦU:
1. [Fix cụ thể]
2. [Giải thích nguyên nhân]
```

---

## 📌 PHẦN 5: MẪU YÊUUCẦU THƯỜNG GẶP

### Mẫu 1: Tạo Module Mới

```
[PHASE 3 - IMPLEMENTATION]

Dự án: [Tên dự án]
Module: [Tên module cần tạo]

Mục đích: [Module làm gì]

Hàm cần tạo:
1. Function1(param1 As Type) As ReturnType
   - Mục đích: [Ghi rõ]
   - Trường hợp lỗi: [Lỗi gì có thể gặp]

2. Function2(param1, param2) As ReturnType
   - Mục đích: [Ghi rõ]
   - Trường hợp lỗi: [Lỗi gì có thể gặp]

Yêu cầu:
- Option Explicit + comments
- Full error handling
- Ví dụ cách sử dụng
- Follow naming convention: [Verb][Subject]AsReturnType
```

### Mẫu 2: Kiểm Tra & Cải Tiến Code

```
[PHASE 6 - OPTIMIZATION]

Dự án: [Tên dự án]
Module: [Module cần optimize]

Vấn đề: [Ghi vấn đề cụ thể]
Hiệu suất hiện tại: [X giây cho Y bản ghi]
Hiệu suất mong muốn: [Z giây]

Code hiện tại:
[Dán code 15-30 dòng]

Kiểm tra:
- Bug hoặc lỗi tiềm ẩn?
- VBA best practice?
- Vấn đề hiệu suất?
- Xử lý lỗi đầy đủ?
- Có thể tối ưu như thế nào?
```

### Mẫu 3: Gỡ Lỗi (Bug Fix)

```
[PHASE 3 - IMPLEMENTATION - BUG FIX]

Dự án: [Tên dự án]
Module: [Module bị lỗi]

Vấn đề: [Mô tả vấn đề 1-2 câu]

Input/Output:
- Input: [Mẫu dữ liệu 5-10 dòng]
- Kỳ vọng: [Output dự kiến]
- Thực tế: [Output hiện tại]

Code:
[Dán hàm/phần liên quan]

Triệu chứng: [Ghi rõ triệu chứng]

Yêu cầu:
1. Nguyên nhân là gì?
2. Cách sửa?
3. Giải thích lỗi?
```

### Mẫu 4: Cập Nhật File .md

```
[PHASE X - DOCUMENTATION]

Yêu cầu cập nhật:
1. File: [Tên file nào]
   Phần: [Phần nào]
   Thay đổi: [Thay đổi gì cụ thể]

2. File: [Tên file nào]
   Phần: [Phần nào]
   Thêm: [Thêm gì]
```

---

## 📌 PHẦN 6: TIẾT KIỆM TOKEN KHI GẶP BUG

### ❌ LÃNG PHÍ TOKEN:

```
"Tôi import CSV không đúng, giúp tôi"
[Dán toàn bộ 1000 dòng code]
[Dán toàn bộ CSV file]
[Dán toàn bộ output error]

Token dùng: ~800 tokens
```

### ✅ TIẾT KIỆM TOKEN:

```
[PHASE 3 - IMPLEMENTATION - BUG FIX]

Module: DataAccess - ReadEmployeeData()
Vấn đề: Email bị cắt khi import CSV

Input: Name,Email,Phone,Score
       A,a@c.com,0123456789,85
Kỳ vọng: Email = "a@c.com" (full)
Thực tế: Email = "a@c" (cắt)

Code (liên quan):
fields = Split(line, " ")  ← Chỗ này có sai?

Sửa như thế nào?

Token dùng: ~250 tokens (-70%)
```

### Mẹo Tiết Kiệm Token:

✅ **LÀM:**
- Phát hiện vấn đề cụ thể trước hỏi
- Ghi rõ input/output/mong muốn
- Dùng mẫu nhỏ (5-10 dòng)
- Chỉ dán code liên quan (~20-30 dòng)
- Ghi triệu chứng rõ ràng
- Hỏi cụ thể

❌ **TRÁNH:**
- Dán toàn bộ module/file
- Giải thích rối rắm
- Hỏi chung chung
- Không cung cấp mẫu
- Dán output lỗi dài

---

## 📌 PHẦN 7: CÁCH CẬP NHẬT FILE .md

### Khi Nào Cập Nhật .md?

| Tình Huống | Hành Động |
|---|---|
| Tạo module mới không dự định | Cập nhật PROJECT_CHECKLIST.md |
| Phát hiện lỗi trong .md | Sửa file lỗi |
| Thay đổi timeline | Cập nhật VBA_AI_WORKFLOW.md |
| Module hoàn thành | Đánh dấu ✓ trong PROJECT_CHECKLIST.md |
| Thêm ví dụ tốt hơn | Thêm vào HUONG_DAN_TIENG_VIET.md |

### Cách Yêu Cầu:

```
❌ SAI: "Cập nhật file"
✅ ĐÚNG: "Cập nhật PROJECT_CHECKLIST.md - thêm module ErrorLogger vào Phase 3"

❌ SAI: "Sửa hướng dẫn"
✅ ĐÚNG: "Sửa HUONG_DAN_TIENG_VIET.md - timeline từ 8 ngày thành 10 ngày"
```

---

## 📌 PHẦN 8: BEST PRACTICES

### ✅ Best Practice 1: Kiểm Tra Trước Hỏi

```
Trước khi hỏi AI:
1. Xác định vấn đề cụ thể là gì
2. Tìm hiểu input/output
3. Tìm kiếm trong .md xem có giải pháp không
4. Sau đó hỏi AI (nếu cần)
```

### ✅ Best Practice 2: Cung Cấp Context

```
Luôn ghi:
- [PHASE X - TÊN PHASE]
- Tên dự án
- Module cụ thể
- Yêu cầu cụ thể
```

### ✅ Best Practice 3: Mẫu Dữ Liệu Nhỏ

```
Cần bao nhiêu mẫu? → 5-10 dòng là đủ
Không cần toàn bộ → AI hiểu từ mẫu
```

### ✅ Best Practice 4: Kiểm Thử Lập Tức

```
Nhận code từ AI:
1. Copy vào Excel
2. Test ngay
3. Nếu lỗi → hỏi AI ngay với lỗi cụ thể
4. Cập nhật PROJECT_CHECKLIST.md
```

### ✅ Best Practice 5: Ghi Lại Quyết Định

```
Khi chọn cách làm:
- Ghi tại sao chọn cách đó
- Ghi điểm mạnh/yếu
- Ghi lại trong PROJECT_CHECKLIST.md
```

---

## 📌 PHẦN 9: TIMELINE THỰC HIỆN

### 8 Ngày Hoàn Thành Dự Án:

```
Ngày 1: Requirements & Analysis (Phase 1)
- Định nghĩa dự án
- Hỏi AI về kiến trúc
- Tạo module list
- Điền PROJECT_CHECKLIST.md

Ngày 2-3: Design & Implementation (Phase 2-3)
- Tạo Config + Utility modules
- Test từng module

Ngày 3-4: Data Access (Phase 3)
- Tạo DataAccess module
- Test đọc/ghi

Ngày 5-6: Business Logic (Phase 3)
- Tạo Business Logic modules
- Test logic

Ngày 7: Integration & Testing (Phase 4-5)
- Gộp tất cả vào Excel
- Test toàn bộ workflow
- Fix bug

Ngày 8: Optimization & Delivery (Phase 6-7)
- Tối ưu hiệu suất
- Viết documentation
- Giao hàng
```

---

## 📌 PHẦN 10: QUICK REFERENCE - BẢNG TÓMLẠI

### Khi Bạn Cần...

| Cần | Làm |
|---|---|
| Bắt đầu dự án | Gõ + [PHASE 1] + định nghĩa dự án |
| Tạo module | Gõ + [PHASE 3] + template tạo module |
| Fix bug | Gõ + [PHASE 3 - BUG FIX] + template bug |
| Tối ưu code | Gõ + [PHASE 6] + template optimize |
| Cập nhật .md | Gõ + yêu cầu cụ thể + tên file |
| Tra cứu VBA | Mở MODULE_TEMPLATE_AND_REFERENCE.md |
| Theo dõi tiến độ | Cập nhật PROJECT_CHECKLIST.md |

### Cách Tiếp Kiếm Xử Dụng Folder:

```
Tối ưu nhất:
1. Đọc QUICK_START_GUIDE.md (lần đầu)
2. Đọc HUONG_DAN_TIENG_VIET.md (hàng ngày)
3. Sử dụng PROJECT_CHECKLIST.md (ghi tiến độ)
4. Tra MODULE_TEMPLATE_AND_REFERENCE.md (cần VBA syntax)
5. Chat trực tiếp với tôi (gỏ yêu cầu)
```

---

## 📌 PHẦN 11: CHECKLIST - TRƯỚC KHI HỎI AI

Trước khi gõ yêu cầu cho AI, kiểm tra:

- [ ] Bạn đã xác định vấn đề cụ thể?
- [ ] Bạn đã ghi Phase mấy?
- [ ] Bạn đã ghi tên dự án?
- [ ] Bạn đã ghi module cụ thể?
- [ ] Bạn có input/output mẫu?
- [ ] Bạn chỉ dán code liên quan?
- [ ] Bạn ghi rõ kỳ vọng vs thực tế?
- [ ] Bạn hỏi cụ thể chứ không chung chung?

**Nếu tất cả "✓" → Hỏi AI ngay!**

---

## 📌 PHẦN 12: VÍ DỤ ĐẦY ĐỦ - TÍNH BONUS NHÂN VIÊN

### Ngày 1: Bắt Đầu

```
[PHASE 1 - REQUIREMENTS & ANALYSIS]

Dự án: Tính Bonus Nhân Viên
Mục đích: Đọc dữ liệu nhân viên từ Excel, tính bonus, xuất kết quả

Yêu cầu:
1. Đọc từ Sheet "Nhân Viên" (Name, Email, Phone, Score)
2. Kiểm tra email và phone hợp lệ
3. Tính bonus: Score < 50 = 5%, 50-75 = 10%, > 75 = 15%
4. Ghi kết quả vào Sheet "Kết Quả"
5. Xử lý 1000 nhân viên trong 10 giây

Câu hỏi:
- Nên tạo bao nhiêu module?
- Cấu trúc như thế nào?
```

### Ngày 2: Tạo Module Đầu Tiên

```
[PHASE 3 - IMPLEMENTATION]

Dự án: Tính Bonus Nhân Viên
Module: ConfigConstants (Module 1)

Tạo hằng số:
- MIN_BONUS = 0.05, MID_BONUS = 0.10, MAX_BONUS = 0.15
- SCORE_LOW = 50, SCORE_HIGH = 75
- INPUT_SHEET = "Nhân Viên", OUTPUT_SHEET = "Kết Quả"
- EMAIL_DOMAINS = "company.com", "backup.company.com"
```

### Ngày 5: Gặp Bug

```
[PHASE 3 - IMPLEMENTATION - BUG FIX]

Dự án: Tính Bonus Nhân Viên
Module: DataAccess - ReadEmployeeData()

Vấn đề: Import data từ CSV bị sai

Input: Name,Email,Phone,Score
       Nguyễn A,nguyen.a@company.com,0912345678,85

Kỳ vọng: Email = "nguyen.a@company.com" (full)
Thực tế: Email = "nguyen.a@c" (bị cắt)

Code:
fields = Split(line, " ")  ← Bug ở đây?

Sửa thế nào?
```

---

## 🎯 KẾT LUẬN

### 3 Quy Tắc Vàng:

1️⃣ **Luôn cung cấp [PHASE] + Context**
   - Không phase → AI không biết bạn ở đâu

2️⃣ **Luôn cung cấp mẫu dữ liệu nhỏ**
   - 5-10 dòng đủ → Tiết kiệm token

3️⃣ **Luôn hỏi cụ thể**
   - Cụ thể → Câu trả lời tốt
   - Chung chung → Câu trả lời chung chung

### Kết Quả:

✅ Chat hiệu quả  
✅ Tiết kiệm token  
✅ Code tốt hơn  
✅ Hoàn thành nhanh hơn  

---

**Lưu ý:** File này là tài liệu tham khảo nhanh. Khi cần chi tiết, xem các file trong folder framework!

---

**Chúc bạn thành công! 🚀**
