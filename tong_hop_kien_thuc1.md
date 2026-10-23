# TÀI LIỆU ÔN TẬP KIẾN THỨC
Tên: Bùi Hải Đăng

**Terminal • Markdown • Microsoft Word • Excel • PowerPoint**

## 1. TERMINAL: CÁC LỆNH CƠ BẢN VÀ NÂNG CAO

Terminal là giao diện dòng lệnh giúp tương tác nhanh với hệ điều hành (PowerShell trên Windows; Terminal trên macOS/Linux).

|Loại lệnh |Lệnh tiêu biểu |Chức năng / Mô tả |
|-|-|-|
|Điều hướng & Quản lý cơ bản |`pwd`/ `Get-Location` |Hiển thị thư mục hiện tại đang làm việc.|
| |`ls`/ `Get-ChildItem` |Liệt kê các tệp và thư mục. |
| |`cd` <ThưMục> |Di chuyển vào thư mục (dùng `cd ..` để lên thư mục cha). |
| |`mkdir` <Tên> |Tạo thư mục mới. |
| |`New-Item` ... |Tạo tệp văn bản mới trên PowerShell. |
| |`Copy-Item`, `Move-Item`, `Rename-Item`, `Remove-Item` |Sao chép, di chuyển, đổi tên và xóa tệp/thư mục. |
|Nâng cao & Lọc dữ liệu |Get-ChildItem -Recurse |Liệt kê toàn bộ nội dung gồm cả thư mục con. |
| |Get-ChildItem -Force |Hiển thị cả các tệp/thư mục ẩn. |
| |`Select-String`/ `grep` |Tìm kiếm từ khóa/chuỗi trong tệp log. |
| |`Get-Content ... -Tail 20` |Xem 20 dòng cuối cùng của tệp log. |
|Kết hợp & Chuyển hướng |Dấu ` |` (Pipe) |
| |Dấu `>` / `>>` |Ghi đè (`>`) hoặc nối tiếp (`>>`) kết quả vào tệp. |


## 2. CÁC ĐỊNH DẠNG TRONG MARKDOWN

Markdown là cú pháp văn bản dạng nhẹ, phổ biến để viết tài liệu và tệp `README.md` trên GitHub.

|Mục đích định dạng |Cú pháp Markdown chuẩn |Ví dụ / Mô tả |
|-|-|-|
|Tiêu đề (Heading) |`#` đến `######` |`# Tiêu đề 1`, `## Tiêu đề 2` |
|Kiểu chữ |**nội dung**, *nội dung*, ~~nội dung~~  |In đậm, in nghiêng, gạch ngang chữ |
|Danh sách |- Mục 1, 1. Bước một |Danh sách không thứ tự và có thứ tự |
|Liên kết & Hình ảnh |[Tên](url), ![Mô tả](link-anh) |Chèn đường dẫn web và hình ảnh |
|Đoạn mã (Code) |`code`, python code |Mã trong dòng hoặc khối mã nguồn |
|Thành phần khác |> Trích dẫn, ---, - [x] Đã xong |Khối trích dẫn, đường kẻ ngang, hộp kiểm (task list)|


## 3. MICROSOFT WORD CƠ BẢN

Hướng dẫn soạn thảo văn bản chuẩn khoa học, tự động hóa tài liệu:

-   **Tạo Heading:** Dùng công cụ sẵn có tại `Home -> Styles` (Heading 1, 2, 3...) để Word nhận diện cấu trúc.
    
-   **Tạo Mục lục tự động:** Sau khi thiết lập Heading, chọn `References -> Table of Contents -> Automatic Table`. Khi có thay đổi, bấm vào mục lục rồi chọn _Update Table_.
    
-   **Đánh số trang:** Vào `Insert -> Page Number`. Dùng _Different First Page_ nếu muốn ẩn số trang ở trang bìa.
    
-   **Mẹo soạn thảo:** Dùng `Ctrl + Enter` (Page Break) để sang trang mới thay vì ấn Enter liên tục; dùng `Ctrl + J` để căn lề đều hai bên.
    

## 4. EXCEL: THỰC HÀNH HÀM VLOOKUP

Hàm `VLOOKUP` dùng để dò tìm giá trị ở cột đầu tiên của bảng tham chiếu và trả về giá trị tương ứng ở các cột tiếp theo trên cùng hàng.

### 4.1. Cú pháp chuẩn

$$\text{=VLOOKUP(lookup\_value, table\_array, col\_index\_num, [range\_lookup])}$$

-   **`lookup_value`**: Giá trị cần tìm kiếm.
    
-   **`table_array`**: Vùng bảng dữ liệu tham chiếu (cột chứa giá trị tìm nằm ở cột đầu tiên của vùng này).
    
-   **`col_index_num`**: Số thứ tự cột chứa kết quả muốn lấy (tính từ trái sang phải).
    
-   **`range_lookup`**: Chọn **`FALSE`** (hoặc `0`) để khớp chính xác tuyệt đối; chọn **`TRUE`** để khớp tương đối.
    

### 4.2. Lưu ý quan trọng

-   Dùng dấu **`$`** để cố định vùng dữ liệu (ví dụ: `$A$2:$C$5`) khi kéo công thức.
    
-   Có thể kết hợp hàm **`TRIM`** để xóa khoảng trắng thừa và **`IFERROR`** để xử lý lỗi hiển thị:
    
    $$\text{=IFERROR(VLOOKUP(TRIM(E2), \$A\$2:\$C\$5, 2, FALSE), "Không tìm thấy")}$$
    

## 5. QUY CHUẨN KHI LÀM VIỆC VỚI POWERPOINT:



-   **Quy tắc nội dung:** Mỗi slide tập trung vào **một thông điệp chính**; hạn chế chép nguyên văn đoạn văn dài; ưu tiên gạch đầu dòng ngắn, từ khóa, sơ đồ hoặc biểu đồ.
    
-   **Quy chuẩn thiết kế:** Dùng bộ màu đồng bộ, độ tương phản cao; tối đa 2 kiểu phông chữ (ưu tiên không chân như Arial, Inter, Roboto); cỡ chữ tiêu đề **28–36 pt**, nội dung **18–24 pt**.
    
-   **Phím tắt điều khiển nhanh:** `F5` (trình chiếu từ đầu), `Shift + F5` (trình chiếu từ slide hiện tại), `Ctrl + M` (thêm slide mới).
