# Test/ — index

Test thuần Node.js, không cần framework hay `npm install` (chỉ dùng `node:assert/strict`,
`node:vm`, và `vendor/fflate.min.js` đã có sẵn trong repo). Chạy trực tiếp:

```bash
node Test/regression_v2_reports.js
node Test/differential_bc_compare.js
node Test/differential_bc_multimonth.js
```

Xem `../CHANGELOG.md` để biết bối cảnh từng lần sửa lỗi mà các test dưới đây đang khoá lại.

## `regression_v2_reports.js`

Suite chính, load `popup.js` + `xlsx-export.js` vào một `vm` sandbox và lắp DOM/`fetch` giả để
chạy trực tiếp các hàm xuất báo cáo. Bao phủ: So sánh 2 kỳ – Tỉnh, Lũy kế tháng – Tỉnh, Dashboard
từng tháng, native `.xlsx` khớp legacy `.xls`, và bất biến nguồn (`assertV2RecordIntegrity`/
`assertV2SheetRowIntegrity` chặn mẫu số rỗng hoặc `onTime` vượt mẫu số). Bao gồm
`testSelectedKpiComboSurvivesWholeMultiMonthPipeline` — khoá lỗi "chọn F1.1+F1.2+F1.3 nhưng file
gần như chỉ có F1.3" (xem CHANGELOG, commit `cac6f27`).

Không có phụ thuộc ngoài repo.

## `differential_bc_compare.js`

So sánh song song **Extension gốc** (`BDHN_DKCL`, chức năng So sánh 2 kỳ hoạt động đúng từ
trước) với bản hiện tại, trên cùng một fixture bưu cục. Dùng khi sửa bất kỳ thứ gì chạm vào
luồng `V2 → So sánh 2 kỳ → Bưu cục` — regression ở luồng này rất dễ tái diễn vì tính năng lũy kế
tháng dùng chung phần lớn code.

**Yêu cầu:** một checkout của Extension gốc trên máy đang chạy test. Mặc định:

```
D:/Antigravity - Project/SOSANH-Cu/SOSANH/BDHN_DKCL
```

Ghi đè bằng biến môi trường nếu checkout ở nơi khác (máy CI, máy khác):

```bash
DKCL_GOLDEN_DIR=/path/to/BDHN_DKCL node Test/differential_bc_compare.js
```

Thiếu thư mục này (hoặc thiếu `popup.js` bên trong) → test báo lỗi đọc file ngay khi khởi động,
không phải lỗi nghiệp vụ. Đây là giới hạn đã biết, không phải bug của test.

## `differential_bc_multimonth.js`

Đặc tả (characterization test) cho `V2 → So sánh lũy kế tháng → Bưu cục` — chức năng không tồn
tại ở Extension gốc nên không có oracle song song; oracle là chính fixture (nhiều bưu cục qua
nhiều tháng, có bưu cục thiếu 1 tháng, có bưu cục chỉ xuất hiện từ tháng 2, tên khác biệt
Unicode/khoảng trắng cùng mã) cộng với các quy tắc PO đã xác nhận, đặc biệt là quy tắc xếp hạng
bưu cục (trong mẫu bưu cục có dữ liệu hợp lệ của đúng tỉnh đang chọn, riêng theo KPI/tháng, đồng
tỷ lệ xếp theo mã tăng dần).

Import `createCurrent`/`makeFetch` từ `differential_bc_compare.js` để dùng chung sandbox và
fixture helper. `differential_bc_compare.js` đọc `popup.js` của Extension gốc ngay khi module
được nạp (kể cả khi golden sandbox không được gọi tới), nên **`DKCL_GOLDEN_DIR` vẫn phải trỏ
đúng chỗ để chạy được file này**, dù bản thân `differential_bc_multimonth.js` không có oracle
song song với Extension gốc.
