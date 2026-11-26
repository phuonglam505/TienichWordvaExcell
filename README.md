# Tiện ích Word và Excel Việt Nam

Bộ add-in miễn phí cho Microsoft Word và Excel bản desktop trên Windows. Phiên bản ổn định hiện tại: **202611.05**.

## Tải bộ cài và mã nguồn

Các tệp phát hành riêng trong [GitHub Releases](https://github.com/phuonglam505/TienichWordvaExcell/releases/tag/addin). Bộ cài tự chứa add-in; gói mã nguồn được đóng riêng.

| Sản phẩm | Bộ cài | Mã nguồn |
|---|---|---|
| Word | [CaiDat-TienIchWord-202611.05-Setup.exe](https://github.com/phuonglam505/TienichWordvaExcell/releases/download/addin/CaiDat-TienIchWord-202611.05-Setup.exe) | [MaNguon-Word-v202611.05.zip](https://github.com/phuonglam505/TienichWordvaExcell/releases/download/addin/MaNguon-Word-v202611.05.zip) |
| Excel | [CaiDat-TienIchExcel-202611.05-Setup.exe](https://github.com/phuonglam505/TienichWordvaExcell/releases/download/addin/CaiDat-TienIchExcel-202611.05-Setup.exe) | [MaNguon-TienIchExcel-v202611.05.zip](https://github.com/phuonglam505/TienichWordvaExcell/releases/download/addin/MaNguon-TienIchExcel-v202611.05.zip) |

## Tính năng chính

**Word:** đọc số thành chữ tiếng Việt/Anh; đổi kiểu chữ; làm sạch và chuẩn hóa văn bản; chuyển bảng mã TCVN3/VNI sang Unicode; định dạng tài liệu; tạo VietQR, QR và mã vạch; dịch vùng chọn; hoàn tác thao tác được hỗ trợ; kiểm tra cập nhật.

**Excel:** đọc số thành chữ và công thức; dịch; tra cứu mã số thuế; chuyển địa chỉ hành chính cũ–mới; tách họ tên và làm sạch văn bản; tạo VietQR, QR và mã vạch; xem tỷ giá; đối chiếu hai bảng; xuất PDF hoặc in vùng chọn; hoàn tác thao tác được hỗ trợ.

Tên lệnh và bố trí có thể thay đổi theo phiên bản; rê chuột lên biểu tượng trên Ribbon để xem gợi ý.

## Cài đặt

![Các bước cài đặt](docs/images/install-flow.svg)

1. Tải bộ cài Word hoặc Excel từ bảng trên.
2. Lưu tài liệu/sổ tính và đóng tất cả cửa sổ Word hoặc Excel.
3. Chạy tệp EXE, chọn **Cài đặt / cập nhật** và chờ hoàn tất.
4. Mở lại Office và dùng tab **Tiện ích Word** hoặc **Tiện ích Excel**.

Add-in cài cho tài khoản Windows hiện tại. Thường không cần quyền quản trị, không cần bật quyền truy cập mã VBA hay thay đổi thiết lập bảo mật Office.

**Chưa thấy tab?** Thoát rồi mở lại Office. Trong Excel, vào File → Options → Add-ins → Manage: Excel Add-ins → Go và bật tiện ích. Trong Word, vào File → Options → Add-ins → Manage: COM Add-ins → Go và bật Tiện ích Word Việt Nam. Nếu Office đã vô hiệu hóa add-in, chạy lại bộ cài để cài hoặc sửa.

## Cách dùng nhanh

- Trong Word, chọn số hoặc vùng văn bản rồi bấm lệnh.
- Đổi bảng mã: chọn đúng TCVN3/VNI, xem trước rồi mới áp dụng.
- VietQR: chọn ngân hàng và nhập số tài khoản; để trống số tiền nếu muốn người chuyển tự điền.
- Dịch: chọn ngôn ngữ trong cài đặt; nên kiểm tra lại bản dịch.
- Excel có thể dùng công thức `=DocSo(A1)` và `=DiaChiMoi(A1)`. Dấu phân cách tham số phụ thuộc cài đặt vùng của Excel.
- Đối chiếu: chọn hai bảng/sheet và cột khóa. Xuất PDF/in: chọn vùng ô và xem trước.
- Dùng Ctrl+Z cho các thao tác được add-in hỗ trợ hoàn tác.

## Cập nhật và gỡ bỏ

Mở **Giới thiệu** hoặc **Kiểm tra cập nhật** nếu nút có trong bản đang dùng, hoặc tải bộ cài mới tại [Releases](https://github.com/phuonglam505/TienichWordvaExcell/releases/tag/addin). Lưu tệp, đóng Word/Excel rồi chạy bộ cài. Để gỡ, đóng Office, chạy lại bộ cài và chọn **Gỡ khỏi Word/Excel**.

## Internet và dữ liệu

Một số tính năng gọi dịch vụ bên ngoài: VietQR, tạo ảnh mã, dịch ngôn ngữ và tra cứu công khai như mã số thuế hoặc tỷ giá. Dữ liệu gửi đi tùy tính năng và nội dung được chọn. Không gửi tài liệu nhạy cảm nếu chưa được phép. Tỷ giá chỉ để tham khảo; thông tin mã số thuế/địa chỉ có thể chậm cập nhật và nên đối chiếu với cơ quan có thẩm quyền. Các thao tác văn bản cục bộ và đổi bảng mã không cần Internet.

## Mã nguồn, giấy phép và hỗ trợ

Tải riêng hai gói mã nguồn trong bảng phát hành. Word gồm giấy phép và thông tin thành phần bên thứ ba; xem các tệp `THAM-KHAO.md` và `THIRD-PARTY-NOTICE.md` trong gói tương ứng. Danh mục nguồn tham khảo không hàm ý các nhà cung cấp chứng thực sản phẩm.

Báo lỗi tại [GitHub Issues](https://github.com/phuonglam505/TienichWordvaExcell/issues), kèm phiên bản, bản Office và các bước tái hiện; không gửi dữ liệu khách hàng hoặc tài liệu mật.

## Yêu cầu

Microsoft Word/Excel bản desktop trên Windows; add-in Word yêu cầu .NET Framework 4.x. Không dành cho Office web, macOS, WPS hoặc ONLYOFFICE. Một số tính năng cần Internet.

## Phiên bản ổn định 202611.05

Cả hai bộ cài và gói mã nguồn đang phát hành đều là **202611.05**. Tải tại [GitHub Releases](https://github.com/phuonglam505/TienichWordvaExcell/releases/tag/addin).
