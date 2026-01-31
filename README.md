Ok 👍 mình viết lại **bài README hoàn chỉnh để up GitHub**, có **giới thiệu ngắn gọn**, **khai báo API VBA**, và **ví dụ sử dụng cơ bản**, trình bày rõ ràng – đúng kiểu một project nghiêm túc, dễ dùng.

Bạn có thể **copy nguyên bài này** làm `README.md`.

---

# VNFastSearch – Vietnamese Fast Search DLL for Excel VBA

**VNFastSearch** là một DLL viết bằng **C++ Builder**, tối ưu cho việc **tìm kiếm tiếng Việt không dấu, realtime, tốc độ cao** trong Excel VBA.

DLL sử dụng **cache trong bộ nhớ** và thuật toán normalize tiếng Việt, cho phép tìm kiếm nhanh hơn rất nhiều so với:

* VBA thuần
* AutoFilter
* ADODB + SQL LIKE

Phù hợp cho:

* Search box realtime
* Dữ liệu lớn (vài nghìn đến hàng chục nghìn dòng)
* Tìm kiếm không phân biệt dấu, hoa thường
* Lọc theo cột hoặc theo khoảng dòng

---

## Nguyên lý hoạt động

1. VBA chuyển dữ liệu Excel → Unicode text
2. DLL:

   * Tách dòng (`\r\n`)
   * Tách cột (`\t`)
   * Normalize tiếng Việt (bỏ dấu + lowercase)
   * Lưu cache trong RAM
3. Khi tìm kiếm:

   * Chỉ so sánh chuỗi đã normalize
   * Trả về **index dòng**
4. VBA dùng index để ghi kết quả ra Sheet / ListBox

➡️ Rất nhanh và không làm Excel bị lag.

---

## Khai báo API DLL trong VBA

Tạo một **Module VBA** (ví dụ: `modVNFastSearch`) và khai báo:

```vb
Option Explicit

' Load du lieu vao bo nho DLL (cache)
Public Declare PtrSafe Sub LoadVNCache Lib "VNFastSearch.dll" ( _
    ByVal txtPtr As LongPtr, _
    ByVal colCount As Long _
)

' Tim cac dong chua keyword
' colMask = 0  -> tim tat ca cac cot
' colMask != 0 -> loc theo bit cot (bit 0 = cot A, bit 1 = cot B, ...)
Public Declare PtrSafe Function FindVNRows Lib "VNFastSearch.dll" ( _
    ByVal keyPtr As LongPtr, _
    ByVal colMask As Long, _
    ByRef rows As Long, _
    ByVal maxCount As Long _
) As Long

' Tim dong trong khoang row chi dinh (1-based theo Excel)
Public Declare PtrSafe Function FindVNRowsExRange Lib "VNFastSearch.dll" ( _
    ByVal keyPtr As LongPtr, _
    ByVal colMask As Long, _
    ByVal rowStart As Long, _
    ByVal rowEnd As Long, _
    ByRef rows As Long, _
    ByVal maxCount As Long _
) As Long

' Kiem tra cache da san sang hay chua
' 0 = chua load, 1 = da san sang
Public Declare PtrSafe Function IsVNCacheReady Lib "VNFastSearch.dll" () As Long

' Dem nhanh so dong phu hop (khong tra ve mang)
Public Declare PtrSafe Function CountVNRows Lib "VNFastSearch.dll" ( _
    ByVal keyPtr As LongPtr, _
    ByVal colMask As Long _
) As Long

' Xoa cache trong DLL
Public Declare PtrSafe Sub ClearVNCache Lib "VNFastSearch.dll" ()
```

---

## Hàm hỗ trợ: chuyển Range → Unicode text

```vb
Public Function RangeToUnicodeText(rg As Range) As String
    Dim arr As Variant
    Dim r As Long, c As Long
    Dim sb As String

    arr = rg.Value

    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            sb = sb & CStr(arr(r, c))
            If c < UBound(arr, 2) Then sb = sb & vbTab
        Next c
        sb = sb & vbCrLf
    Next r

    RangeToUnicodeText = sb
End Function
```

---

## Load cache (gọi 1 lần)

```vb
Public Sub InitVNCacheOnce()
    If IsVNCacheReady() <> 0 Then Exit Sub

    Dim rg As Range
    Set rg = Sheets("Data").Range("A1").CurrentRegion

    LoadVNCache StrPtr(RangeToUnicodeText(rg)), rg.Columns.Count
End Sub
```

📌 Nên gọi khi:

* Workbook mở
* UserForm Initialize
* Hoặc trước lần tìm kiếm đầu tiên

---

## Ví dụ tìm kiếm cơ bản (ghi kết quả ra Sheet)

```vb
Public Sub VN_Search_ToSheet( _
    ByVal keyword As String, _
    ByVal outCell As Range)

    Dim idx(1 To 1000) As Long
    Dim found As Long
    Dim i As Long
    Dim src As Range

    Call InitVNCacheOnce

    If Trim$(keyword) = "" Then
        outCell.Resize(1000, 50).ClearContents
        Exit Sub
    End If

    found = FindVNRows(StrPtr(keyword), 0, idx(1), 1000)
    If found <= 0 Then Exit Sub

    Set src = Sheets("Data").Range("A1").CurrentRegion
    outCell.Resize(1000, src.Columns.Count).ClearContents

    For i = 1 To found
        outCell.Offset(i - 1, 0).Resize(1, src.Columns.Count).Value = _
            Application.Index(src.Value, idx(i), 0)
    Next i
End Sub
```

---

## Ví dụ dùng với TextBox trên Sheet

```vb
Private Sub TextBox1_Change()
    VN_Search_ToSheet Me.TextBox1.Text, Me.Range("E3")
End Sub
```

➡️ Gõ tới đâu, tìm tới đó, không phân biệt dấu.

---

## Ví dụ lọc theo cột

```vb
' Chi tim cot A va C
' bit 0 = A, bit 2 = C -> 1 + 4 = 5
found = FindVNRows(StrPtr("lan"), 5, idx(1), 1000)
```

---

## Ví dụ đếm nhanh (không ghi kết quả)

```vb
Dim n As Long
n = CountVNRows(StrPtr("an"), 0)
Debug.Print "So dong tim duoc:", n
```

---

## Ưu điểm

* Tìm kiếm tiếng Việt **không dấu**
* Rất nhanh (cache trong RAM)
* Không làm Excel bị lag
* Hỗ trợ:

  * Lọc theo cột
  * Lọc theo khoảng dòng
  * Realtime search
* Dễ mở rộng thêm điều kiện

---

## Thông tin tác giả

**Kieu Manh**
Email: **[kieumanh366377@gmail.com](mailto:kieumanh366377@gmail.com)**

---

