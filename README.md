Dưới đây là **bài giới thiệu hoàn chỉnh** cho dự án **VNFastSearch – Vietnamese Fast Search DLL for Excel VBA**, **không nhắc lại tiêu đề hay chữ ký hàm C++**, chỉ tập trung vào **ý tưởng, cách dùng và lợi ích**, đúng kiểu để đưa lên GitHub README.md 👍

---

# VNFastSearch – Vietnamese Fast Search DLL for Excel VBA

**VNFastSearch** là một DLL viết bằng **C++ (C++Builder / MSVC)**, được thiết kế để **tìm kiếm tiếng Việt cực nhanh trong Excel VBA**, đặc biệt phù hợp với dữ liệu lớn (hàng nghìn đến hàng trăm nghìn dòng).

Thư viện tập trung giải quyết các hạn chế cố hữu của VBA khi:

* Tìm kiếm không dấu
* Tìm kiếm realtime khi gõ TextBox
* Tìm kiếm trên nhiều cột
* Hiệu năng kém với Like, InStr, Filter, hoặc ADODB SQL

---

## Mục tiêu chính

* Tìm kiếm **tiếng Việt không dấu**
* Tốc độ cao (cache dữ liệu trong bộ nhớ DLL)
* Dùng trực tiếp từ **Excel VBA**
* Linh hoạt: tìm toàn bộ cột hoặc chọn cột
* Dễ mở rộng cho các yêu cầu phức tạp hơn

---

## Ý tưởng hoạt động

1. Excel chuyển dữ liệu (Range) thành chuỗi Unicode
2. DLL nhận dữ liệu, tách theo dòng và cột
3. Mỗi ô được **normalize** (bỏ dấu + lowercase)
4. Dữ liệu được lưu trong bộ nhớ RAM của DLL
5. Các lần tìm kiếm sau chỉ chạy trên cache → **rất nhanh**

> Cache chỉ cần load **1 lần**, tìm kiếm có thể gọi hàng nghìn lần.

---

## Normalize tiếng Việt

DLL xử lý:

* Bỏ toàn bộ dấu tiếng Việt
* Chuyển về chữ thường
* So khớp bằng substring search

Ví dụ:

| Gốc              | Sau normalize    |
| ---------------- | ---------------- |
| Nguyễn Văn An    | nguyen van an    |
| Thành phố Hà Nội | thanh pho ha noi |
| Điện thoại       | dien thoai       |

Người dùng có thể gõ:

* `an`
* `ha noi`
* `dien`

→ vẫn tìm ra kết quả đúng.

---

## Cache dữ liệu

* Cache nằm **hoàn toàn trong DLL**
* VBA chỉ giữ dữ liệu gốc để hiển thị
* Có thể:

  * kiểm tra cache đã sẵn sàng
  * xoá cache
  * đếm số dòng trong cache

Điều này giúp:

* Tránh xử lý lại dữ liệu nhiều lần
* Giảm CPU và lag Excel
* Phù hợp tìm kiếm realtime

---

## Tìm kiếm linh hoạt

DLL hỗ trợ:

* Tìm toàn bộ cột
* Tìm theo cột chọn lọc bằng bit mask
* Tìm trong khoảng dòng (row start / end)
* Đếm nhanh số kết quả (không trả mảng)

Rất phù hợp cho:

* AutoComplete
* Search box
* Filter dữ liệu lớn
* Giao diện UserForm / ActiveX TextBox

---

## Hiệu năng

So sánh tương đối:

| Phương pháp          | Tốc độ              |
| -------------------- | ------------------- |
| VBA InStr / Like     | Chậm                |
| Filter               | Trung bình          |
| ADODB SQL Like       | Chậm – mở recordset |
| **VNFastSearch DLL** | 🚀 Rất nhanh        |

Với 10.000–50.000 dòng:

* Load cache: ~ vài chục ms
* Tìm kiếm mỗi lần gõ: gần như tức thì

---

## Cách dùng trong VBA (tóm tắt)

Quy trình chuẩn:

1. Load cache từ Sheet dữ liệu
2. Khi gõ TextBox → gọi hàm tìm kiếm
3. Nhận danh sách row index
4. Dùng `Application.Index` để ghi kết quả ra Sheet hoặc ListBox

Mã VBA được thiết kế đơn giản, dễ tái sử dụng.

---

## Khi nào nên dùng VNFastSearch

* Dữ liệu lớn
* Cần tìm không dấu
* Cần tìm nhanh khi gõ
* VBA thuần không đáp ứng được tốc độ
* Không muốn phụ thuộc ADODB / SQL

---

## Khả năng mở rộng

Thiết kế DLL cho phép mở rộng dễ dàng:

* Tìm nhiều keyword (AND / OR)
* Regex đơn giản
* Cache đa sheet
* Multi-thread load cache
* Tìm theo nhiều điều kiện

---

## Thông tin tác giả

**Kieu Manh**
Email: **[kieumanh366377@gmail.com](mailto:kieumanh366377@gmail.com)**

---


