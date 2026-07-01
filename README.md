# ExcelCSIToolBoxAddIn

ExcelCSIToolBoxAddIn là add-in VSTO cho Microsoft Excel, dùng để kết nối các workflow tính toán trong Excel với các sản phẩm CSI như **ETABS** và **SAP2000**.

Add-in cung cấp tab ribbon trong Excel, mở các cửa sổ toolbox riêng cho ETABS/SAP2000, kết nối tới model CSI đang chạy, đọc/ghi dữ liệu qua Excel range, và gom các tiện ích hậu xử lý kết quả.

## Tính Năng Chính

- Kết nối Excel với ETABS/SAP2000 thông qua CSI Open API.
- Mở ETABS Toolbox trực tiếp từ ribbon Excel.
- Đọc/xuất các bảng kết quả như base reactions, modal mass participation ratios, story forces, story drifts, mass summary by story.
- Tạo và cập nhật đối tượng/model từ dữ liệu Excel.
- Hỗ trợ workflow WPF/MVVM cho các cửa sổ công cụ.
- Có lớp AI/MCP đang phát triển cho trợ lý thao tác với model.

## Cấu Trúc Solution

```text
ExcelCSIToolBoxAddIn.sln
|
+-- ExcelCSIToolBoxAddIn              Main Excel VSTO add-in project
+-- ExcelCSIToolBox.Application       Use cases và workflow orchestration
+-- ExcelCSIToolBox.Core              Shared contracts, results, common logic
+-- ExcelCSIToolBox.Data              DTOs, mapper models, table schemas
+-- ExcelCSIToolBox.Infrastructure    ETABS/SAP2000 API adapters, Excel interop
+-- ExcelCSIToolBox.AI                AI/chatbox/MCP integration layer
+-- ExcelCSIToolBox.Tests             Unit tests
```

## Yêu Cầu Trước Khi Cài

- Windows.
- Microsoft Excel desktop app.
- Microsoft .NET Framework 4.8.
- Microsoft Visual Studio Tools for Office Runtime.
- ETABS và/hoặc SAP2000 nếu dùng các tính năng kết nối CSI.
- Các file CSI interop trong `lib/`:
  - `ETABSv1.dll`
  - `SAP2000v1.dll`

> Nếu cài bằng `publish/ExcelCSIToolbox/setup.exe`, installer có thể tự cài thêm .NET Framework 4.8 và VSTO Runtime nếu máy chưa có.

## Cài Đặt Cho User Sau Khi Clone/Download Repo

Dùng cách này nếu bạn chỉ muốn cài add-in để sử dụng trong Excel, không cần debug source code.

1. Clone repo hoặc tải file ZIP từ GitHub:

   ```powershell
   git clone <repo-url>
   ```

2. Nếu tải bằng ZIP, nên bấm chuột phải vào file ZIP, chọn **Properties**, tick **Unblock** nếu có, rồi mới extract. Bước này giúp tránh lỗi Windows chặn file `.vsto` sau khi giải nén.

3. Đóng tất cả cửa sổ Excel đang mở.

4. Mở thư mục publish:

   ```text
   publish/ExcelCSIToolbox/
   ```

5. Chạy file:

   ```text
   setup.exe
   ```

6. Nếu Windows hiện cảnh báo do add-in được ký bằng certificate tạm thời/self-signed, chỉ tiếp tục cài đặt khi bạn tin tưởng source repo.

7. Mở lại Excel. Trên ribbon sẽ có tab:

   ```text
   ExcelCSIToolBox
   ```

8. Bấm **ETABS Toolbox** hoặc các nút công cụ khác để sử dụng. Với các tính năng cần model CSI, hãy mở ETABS/SAP2000 và model trước, sau đó attach từ toolbox.

## Nếu Không Thấy Add-In Trong Excel

- Vào **Excel > File > Options > Add-ins**.
- Ở **Manage**, chọn **COM Add-ins**, bấm **Go...**.
- Kiểm tra `ExcelCSIToolBoxAddIn` đã được tick chưa.
- Nếu add-in nằm trong **Disabled Items**, chọn **Excel > File > Options > Add-ins > Manage: Disabled Items** và enable lại.
- Nếu vẫn lỗi, đóng Excel và chạy lại `publish/ExcelCSIToolbox/setup.exe`.

## Gỡ Cài Đặt

1. Đóng Excel.
2. Vào **Windows Settings > Apps > Installed apps**.
3. Tìm `ExcelCSIToolBoxAddIn`.
4. Chọn **Uninstall**.

## Build Và Debug Cho Developer

Dùng cách này nếu bạn muốn sửa code hoặc chạy debug trong Visual Studio.

1. Cài Visual Studio với workload **Office/SharePoint development**.
2. Đảm bảo máy có .NET Framework 4.8 Developer Pack.
3. Mở solution:

   ```text
   ExcelCSIToolBoxAddIn.sln
   ```

4. Restore/build solution bằng Visual Studio.
5. Chọn project `ExcelCSIToolBoxAddIn` làm startup project.
6. Bấm **Start Debugging**. Visual Studio sẽ mở Excel với add-in đã load.
7. Mở ETABS/SAP2000 trước khi test các command phụ thuộc CSI API.

## Publish Lại Installer

Khi cần tạo gói cài đặt mới:

1. Mở project `ExcelCSIToolBoxAddIn` trong Visual Studio.
2. Chọn **Build > Publish ExcelCSIToolBoxAddIn**.
3. Publish vào thư mục mặc định:

   ```text
   publish/ExcelCSIToolbox/
   ```

4. Kiểm tra trong thư mục publish có:

   ```text
   setup.exe
   ExcelCSIToolBoxAddIn.vsto
   Application Files/
   ```

5. Gửi/tổ chức repo kèm nguyên thư mục `publish/ExcelCSIToolbox/` để user cài bằng `setup.exe`.

## Troubleshooting

- **Không hiện tab ExcelCSIToolBox**: kiểm tra COM Add-ins và Disabled Items trong Excel.
- **Lỗi trust/certificate khi cài**: file publish đang dùng temporary certificate. Chỉ cài khi tin tưởng source repo, hoặc publish lại bằng certificate nội bộ của team.
- **Lỗi Windows chặn file `.vsto`**: Unblock file ZIP trước khi extract, hoặc Unblock riêng `setup.exe`/`ExcelCSIToolBoxAddIn.vsto`.
- **Lỗi thiếu VSTO Runtime**: chạy `setup.exe` thay vì mở trực tiếp file `.vsto`.
- **Lỗi kết nối ETABS/SAP2000**: mở ETABS/SAP2000 và model trước, sau đó attach lại từ toolbox. Kiểm tra version API DLL trong `lib/` có phù hợp với phiên bản CSI đang dùng không.

## Ghi Chú Cho Contributor

- Target framework: **.NET Framework 4.8**.
- Host application: Microsoft Excel thông qua VSTO.
- UI: WPF với MVVM-style ViewModels.
- CSI API access nên được cô lập trong Infrastructure adapters.
- UI chỉ nên giữ logic điều phối mỏng; workflow chính nên nằm trong Application/Core khi phù hợp.
- RefBuilder là utility phục vụ sinh scaffold/reference và không phải flow chạy chính của add-in.
