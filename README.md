# Tiện ích Word và Excel Việt Nam

Bộ add-in miễn phí cho Microsoft Office desktop trên Windows, đưa các công cụ văn phòng thường dùng vào Ribbon quen thuộc. Bản mã nguồn/tài liệu: **2026.11.06**.

![Thanh công cụ Word](docs/images/word-ribbon.svg)

![Thanh công cụ Excel](docs/images/excel-ribbon.svg)

## Tải bộ cài và mã nguồn

Mã nguồn và bộ cài được phát hành thành các tệp riêng trong Assets của GitHub Release; tên bắt đầu bằng `MaNguon-` là gói mã nguồn, còn `CaiDat-` là bộ cài. Cả hai add-in đều được dựng cho phiên bản 2026.11.06:

| Sản phẩm | Bộ cài (Assets của Release) | Mã nguồn |
|---|---|---|
| Word | [CaiDat-TienIchWord-2026.11.06-Setup.exe](https://github.com/phuonglam505/TienichWordvaExcell/releases/download/2026.11.06/CaiDat-TienIchWord-2026.11.06-Setup.exe) | [MaNguon-Word-v2026.11.06.zip](https://github.com/phuonglam505/TienichWordvaExcell/releases/download/2026.11.06/MaNguon-Word-v2026.11.06.zip) |
| Excel | [CaiDat-TienIchExcel-2026.11.06-Setup.exe](https://github.com/phuonglam505/TienichWordvaExcell/releases/download/2026.11.06/CaiDat-TienIchExcel-2026.11.06-Setup.exe) | [MaNguon-TienIchExcel-v2026.11.06.zip](https://github.com/phuonglam505/TienichWordvaExcell/releases/download/2026.11.06/MaNguon-TienIchExcel-v2026.11.06.zip) |

Bộ cài tự chứa add-in; không tải DLL/XLAM rời. Gói mã nguồn không trộn với EXE; có thể tải riêng từng loại ở bảng trên.

## Tính năng chính

**Word:** đọc số thành chữ Việt/Anh và chọn đơn vị; đổi hoa/thường; làm sạch, chuẩn hóa, chuyển bảng mã TCVN3/VNI sang Unicode; định dạng tài liệu; VietQR; mã QR/Code 128/Code 39/EAN-13; dịch vùng chọn; cảnh báo vùng chọn lớn; hoàn tác và kiểm tra cập nhật.

**Excel:** đọc số thành chữ và công thức; dịch; tra cứu MST; đổi địa chỉ hành chính cũ–mới; tách họ tên và làm sạch văn bản; VietQR, QR/barcode; xem tỷ giá; so sánh hai bảng theo mã; xuất PDF/in vùng chọn; hoàn tác thao tác được add-in hỗ trợ.

Tên lệnh, giới hạn và cách bố trí có thể thay đổi theo phiên bản; rê chuột lên biểu tượng Ribbon để xem hướng dẫn.

## Cài đặt

![Các bước cài đặt](docs/images/install-flow.svg)

1. Tải bộ cài Word hoặc Excel từ Assets trong GitHub Releases.
2. Lưu tài liệu/sổ tính, đóng tất cả cửa sổ Word/Excel.
3. Chạy EXE, chọn **Cài đặt / cập nhật**, chờ hoàn tất.
4. Mở lại Office, dùng tab **Tiện ích Word** hoặc **Tiện ích Excel**.

Cài cho tài khoản Windows hiện tại, thường không cần quyền quản trị. Không cần bật quyền truy cập mã VBA hay tắt bảo mật Office.

**Chưa thấy tab?** Thoát và mở lại Office. Excel: File → Options → Add-ins → Manage: Excel Add-ins → Go, bật tiện ích. Word: File → Options → Add-ins → Manage: COM Add-ins → Go, bật Tiện ích Word Việt Nam. Nếu Office từng vô hiệu hóa add-in, chạy lại bộ cài để cài/sửa.

## Cách dùng nhanh

- Chọn số hoặc vùng văn bản trước khi bấm lệnh trong Word.
- Đổi bảng mã: chọn đúng TCVN3/VNI, xem trước rồi mới áp dụng.
- VietQR: chọn ngân hàng, nhập số tài khoản; để trống số tiền nếu muốn người chuyển tự điền.
- Dịch: chọn ngôn ngữ/nhà cung cấp trong cài đặt; kiểm tra lại kết quả dịch máy.
- Excel: công thức mẫu `=DocSo(A1)`, `=DiaChiMoi(A1)`. Dấu phân cách tham số phụ thuộc thiết lập vùng của Excel; có thể cần thay `;` bằng `,`.
- Đối chiếu: chọn hai bảng/sheet và cột khóa; kiểm tra kết quả trước khi tô màu. Xuất PDF/in: chọn vùng ô, kiểm tra xem trước.
- Dùng Ctrl+Z cho các thao tác được tiện ích hỗ trợ hoàn tác.

## Cập nhật và gỡ bỏ

Dùng **Giới thiệu** hoặc **Kiểm tra cập nhật** nếu nút có trong bản đang cài, hoặc tải bộ cài mới từ Releases. Lưu tệp, đóng Word/Excel rồi chạy EXE mới. Để gỡ, đóng Office, chạy lại bộ cài và chọn **Gỡ khỏi Word/Excel**. Tùy chọn cá nhân có thể được giữ lại.

## Internet và dữ liệu

Một số tính năng chủ động gọi dịch vụ ngoài: VietQR (ảnh thanh toán), BWIP-JS/QRServer (ảnh mã), MyMemory/Google Translate web (dịch), nguồn tra cứu công khai (MST/tỷ giá). Dữ liệu gửi đi tùy tính năng và nội dung bạn chọn. Không gửi tài liệu nhạy cảm nếu chưa được phép. Tỷ giá mang tính tham khảo; thông tin MST/địa chỉ có thể chậm cập nhật và cần đối chiếu với cơ quan có thẩm quyền. Các thao tác văn bản cục bộ và đổi bảng mã không cần Internet.

## Mã nguồn, giấy phép và hỗ trợ

Mã nguồn Word/Excel tách khỏi bộ cài ở `source/Word` và `source/Excel`. Word phân phối theo GPL-2.0-or-later; xem giấy phép UniKey/third-party trong gói. Danh mục nguồn tham khảo/API nằm trong `THAM-KHAO.md` và `THIRD-PARTY-NOTICE.md` đi kèm. Đây là ghi nhận nguồn, không hàm ý các nhà cung cấp chứng thực sản phẩm.

Báo lỗi tại [GitHub Issues](https://github.com/phuonglam505/TienichWordvaExcell/issues), kèm phiên bản, bản Office và bước tái hiện; không gửi dữ liệu khách hàng/tài liệu mật.

## Yêu cầu

Microsoft Word/Excel bản desktop trên Windows; add-in Word yêu cầu .NET Framework 4.x. Không dành cho Office web, macOS, WPS hoặc ONLYOFFICE. Một số tính năng yêu cầu Internet.

## Phiên bản 2026.11.06

Cả mã nguồn và bộ cài Word/Excel đều mang phiên bản **2026.11.06**.

Tải đúng bộ cài trong [GitHub Releases](https://github.com/phuonglam505/TienichWordvaExcell/releases/tag/2026.11.06); chọn tệp có phiên bản mới nhất.


