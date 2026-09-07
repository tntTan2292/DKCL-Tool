# Changelog — DKCL Tool

Nhật ký kỹ thuật cho các lần sửa lỗi production của extension. Mỗi mục ghi root cause, phạm vi
remediation, kết quả regression và trạng thái xác nhận của PO. Không thay thế `HDSD.md` (hướng
dẫn sử dụng cho người dùng cuối).

---

## 2026-09-07 — commit `cac6f27` — PO PASS

**Trạng thái: Đã xác nhận production.** PO đã kiểm tra và PASS cả 3 luồng dưới đây.

### Vấn đề 1 — Báo cáo Tỉnh V2, So sánh lũy kế tháng: mất KPI đã chọn

**Triệu chứng:** Tick F1.1 + F1.2 + F1.3 (bỏ F4.1), file xuất ra gần như chỉ có F1.3 — F1.1/F1.2
gần như trống hoặc bằng 0.

**Root cause (hai nguyên nhân cộng dồn, cả hai đều tồn tại ở HEAD `388dd46`):**
1. Offset cột nguồn cũ trong `V2_SOURCE_CONFIGS.values`/`bcValues` khiến mẫu số F1.1/F1.2 đọc
   trúng ô rỗng của bảng HTML API (API đã đổi bố cục cột) → `percentFrom` trả về `0`. F1.3 chỉ
   cần 2 ô nên vẫn trúng số thật, tạo cảm giác "chỉ có F1.3". *(Đã remap theo capture production
   ở các lượt sửa trước `cac6f27`; xem mục "Trạng thái parser" bên dưới.)*
2. `exportReportV2MultiMonth` truyền cố định `Object.values(MULTI_MONTH_METRICS_MAP)` cho cả
   legacy `.xls` và native `.xlsx` — bất kể KPI nào được tick, exporter luôn nhận đủ 4 KPI và
   dựng thẻ/cột/chart `0` giả cho KPI không được chọn.
3. (Phát hiện thêm trong lúc audit) Câu "Ngưỡng SSOT" trong dashboard hardcode văn bản cho cả 4
   KPI, nên ngay cả khi exporter đã lọc đúng, workbook vẫn in ra ngưỡng của KPI không hề tồn tại
   trong file.

**Remediation:**
- `exportReportV2MultiMonth` chuyển sang `getActiveMultiMonthMetrics(tuyChonGR)` làm SSOT duy
  nhất cho cả legacy `.xls` và native `.xlsx` — không còn đường nào khác để lấy danh sách KPI.
- Câu "Ngưỡng SSOT" dựng động từ `activeMetrics`, không còn liệt kê KPI không được xuất.
- `assertV2RecordIntegrity` / `assertV2SheetRowIntegrity` (đã có từ lượt trước) tiếp tục chặn
  cứng nếu mẫu số rơi vào ô rỗng — không để lặp lại kiểu lỗi #1 một cách âm thầm.

**File sửa:** `popup.js` (`exportReportV2MultiMonth`, khối `notesTextLines`).

**Regression:** `Test/regression_v2_reports.js` →
`testSelectedKpiComboSurvivesWholeMultiMonthPipeline` — dựng đúng tổ hợp F1.1+F1.2+F1.3 qua 3
tháng, kiểm tra toàn chuỗi `monthlyResults` → legacy dashboard → native model → `DuLieu_Chart` →
matrix sheet → chart parts đều có đúng 3 KPI, không KPI nào bị thiếu và không nhắc đến F4.1 ở bất
kỳ đâu trong package. Mutation kiểm chứng: quay lại `Object.values(...)` → FAIL; hardcode câu
ngưỡng → FAIL; bỏ qua lựa chọn KPI → FAIL; chỉ giữ KPI cuối (mô phỏng đúng triệu chứng gốc) →
FAIL.

---

### Vấn đề 2 — So sánh 2 kỳ – Bưu cục: gộp thành một dòng cấp tỉnh

**Triệu chứng:** Chọn cấp Bưu cục, file xuất ra không có danh sách BCVH mà chỉ có một dòng
mã/tên tỉnh.

**Root cause:** Extension gốc (`SOSANH-Cu/SOSANH/BDHN_DKCL`) dùng `config.bcCodeIndex` /
`config.bcNameIndex` để đọc mã/tên bưu cục từ một cột riêng trong bảng HTML API. Nhánh này bị
xoá ở phiên bản hiện tại của `getReportCodeIndex`/`pickReportName`, nên mọi dòng BC rơi về
`codeIndex`/`nameIndex` mặc định (cột mã/tên **tỉnh**) — toàn bộ bưu cục gộp vào một dòng.
Ngoài ra `fetchV2ReportRows`, `finalizeRows`, `getActiveV2SourceConfigs`, `buildV2ComparisonRows`
và bộ 3 hàm dựng sheet `SoSanhV2_BC` cũng bị mất nhánh BC tương ứng.

**Remediation:** Khôi phục nguyên trạng logic BC từ Extension gốc: seeding danh mục bưu cục qua
`fetchCombinedReport`, giới hạn KPI cấp BC còn F4.1 + F1.3 (`V2_BC_SOURCE_KEYS`), lọc bưu cục
đang hoạt động có dữ liệu (`isBcOperatingRowWithData`), khôi phục
`renderV2BcComparisonWorksheet`/`renderBcComparisonXmlRows`/`buildBcComparisonText` và định
tuyến BC về sheet `SoSanhV2_BC`. Gỡ 2 chỗ hardcode mã tỉnh `"10"` còn sót trong tham số request
V1, thay bằng `getV2SelectedProvinceFilter(tuyChonGR)` đọc động từ dropdown.

**File sửa:** `popup.js` (`getReportCodeIndex`, `pickReportName`, `pickV2JoinCodeText`,
`fetchCombinedReport`, `finalizeRows`, `getActiveV2SourceConfigs`, `fetchV2ReportRows`,
`finalizeV2Rows`, `buildV2ComparisonRows`, `downloadV2ComparisonExcel`,
`renderV2ReportPreview`, `exportReportV2`, + 3 hàm khôi phục nguyên trạng).

**Regression:** `Test/differential_bc_compare.js` — chạy song song Extension gốc và bản hiện
tại trên cùng một fixture (nhiều bưu cục, có bưu cục chỉ xuất hiện 1 kỳ, tên khác biệt
khoảng trắng/Unicode cùng mã), so khớp từng giá trị nghiệp vụ (mã/tên BC, sản lượng, tỷ lệ,
chênh lệch, tiêu đề cột, tổng số dòng) và request filter. 7/7 mutation bắt đúng (key theo mã
tỉnh, tên tỉnh thay tên BC, gộp mọi BC vào 1 record, join theo tên, mapping cấp tỉnh, dùng data
kỳ 1 cho cả 2 kỳ, mất BC 1 kỳ).

---

### Vấn đề 3 — Lũy kế tháng – Bưu cục: cùng lỗi gộp, cộng thêm exporter throw cứng

**Triệu chứng:** Giống Vấn đề 2, biểu hiện ở chế độ Lũy kế tháng.

**Root cause:** Tầng dữ liệu dùng chung `fetchV2ReportRows` nên tự hết lỗi sau khi Vấn đề 2 được
sửa. Tầng còn lại: cả hai exporter (`popup.js:renderV2MultiMonthDashboardWorksheet` và
`xlsx-export.js:buildModel`) bắt buộc phải tìm được một "tỉnh tiêu điểm" trong tập đơn vị. Ở BC,
đơn vị là bưu cục — tỉnh không bao giờ là một dòng — nên cả hai luôn `throw`.

**Remediation:**
- Focus cấp BC = tổng hợp (sum) các bưu cục của tỉnh đã chọn, thay vì tìm một dòng tỉnh
  (`aggregateUnitRows` trong `popup.js`, `aggregateUnits` trong `xlsx-export.js`).
- Ma trận `MaTran_BuuCuc` liệt kê từng bưu cục theo từng tháng: sản lượng, tỷ lệ, hạng, chênh
  lệch, xu hướng ▲/▼/→; N/A khi tháng đó không có dữ liệu.
- **Xếp hạng bưu cục** (bổ sung theo yêu cầu PO, xác nhận sau khi vấn đề chính đã PASS): trong
  các bưu cục **có dữ liệu hợp lệ của đúng tỉnh đang chọn**, tính riêng theo từng KPI và từng
  tháng — không bao giờ dùng thang 34 tỉnh. Đồng tỷ lệ xếp theo mã bưu cục tăng dần (dùng lại
  đúng quy tắc `assignRankByColumn` sẵn có của dự án, để cả workbook chỉ có một ngữ nghĩa hạng).
  Mẫu xếp hạng được in rõ trên sheet: *"Hạng trong các Bưu cục có dữ liệu hợp lệ thuộc tỉnh
  được chọn ... trong từng tháng, tính riêng theo từng KPI."*
- F1.1/F1.2 ở cấp BC báo lỗi rõ ràng *"không hỗ trợ theo Extension gốc"* thay vì sinh dòng `0`.

**File sửa:** `popup.js` (`aggregateUnitRows`, nhánh BC trong
`renderV2MultiMonthDashboardWorksheet`, `assertBcKpiSelection`, `V2_BC_SHORT_NAMES`), `xlsx-export.js`
(`aggregateUnits`, `monthlyByMetricFor`, `rankingSample` trong `buildModel`).

**Regression:** `Test/differential_bc_multimonth.js` — fixture nhiều bưu cục qua 3 tháng (một
bưu cục thiếu 1 tháng, một bưu cục chỉ xuất hiện từ tháng 2, tên khác biệt Unicode cùng mã).
Kiểm tra: mỗi tháng/mỗi KPI tính hạng độc lập; chỉ bưu cục của tỉnh đang chọn vào mẫu; N/A bị
loại khỏi mẫu; `0/0` hợp lệ vẫn tham gia; đổi tỷ lệ 1 BC/1 tháng chỉ đổi đúng hạng KPI/tháng đó;
legacy `.xls` và native `.xlsx` khớp từng ô hạng. Tổng cộng 7 mutation bắt đúng, gồm 2 mutation
PO chỉ định: "dùng rank 34 tỉnh cho BC" và "đưa N/A vào rank BC".

---

### Trạng thái parser (ghi nhận để tránh hiểu nhầm)

Parser V2 hiện tại đọc bảng HTML API theo **offset vị trí cố định**
(`V2_SOURCE_CONFIGS.values` / `bcValues` / `codeIndex` / `nameIndex` / `bcCodeIndex` /
`bcNameIndex`), **không phải mapping theo tên cột (header alias)**. Các offset đang dùng đã được
remap thủ công từ một capture dữ liệu production (2026-09-02) sau khi API đổi bố cục cột, và
được bọc bằng hai lớp bất biến fail-loud (`assertV2RecordIntegrity` ở tầng record,
`assertV2SheetRowIntegrity` ở tầng dòng sheet) chặn cứng khi mẫu số rỗng hoặc `onTime` vượt mẫu
số — thay vì âm thầm ghi `0`. Nếu API đổi bố cục cột lần nữa, hai lớp bất biến này sẽ dừng export
và báo lỗi kèm offset/ô đọc được/header nguồn, chứ **không tự khớp lại theo tên cột**. Việc
chuyển sang mapping theo header alias, nếu cần, là một hạng mục riêng — chưa nằm trong phạm vi
các lần sửa đã liệt kê ở trên.

### Kết quả regression tổng hợp (tại commit `cac6f27`)

```
node Test/regression_v2_reports.js       → V2 report regressions passed
node Test/differential_bc_compare.js     → BC compare differential passed
node Test/differential_bc_multimonth.js  → BC multi-month characterization passed
git diff --check                         → exit 0
```

`Test/differential_bc_compare.js` cần một checkout của Extension gốc để đối chiếu song song.
Mặc định trỏ tới `D:/Antigravity - Project/SOSANH-Cu/SOSANH/BDHN_DKCL`; ghi đè bằng biến môi
trường `DKCL_GOLDEN_DIR` nếu chạy trên máy khác (xem `Test/README.md`).

**SHA-256 các file extension tại commit `cac6f27`** (đúng bộ Chrome nạp theo `manifest.json`):

```
46a6c8112df43aa8aa92cbf16ec39068246423f73ae2b6f68d176ed96f04beba  popup.js
f648b7867bc0406b52dd9679c43766de7a54dfc7fd63bc5258bc52fbac0ba0cc  xlsx-export.js
b2b43d79b7250e11309f18cfe2695e4fada24d23118eeadc94639a0bbe60b605  popup.html
2875dde59913edb12f54c15b640362465f445ff2be47d202380a66ac33c766c5  popup.css
b79650e9ac76d13428fd112e228bae533deb61a7d56810654779cfcf8ad451b5  manifest.json
385ddec3fa37ed5af435cef51169ceab74b39d1a1a6297e8b097683c079e0e8c  background.js
5f6283411130c9c9b092f1b9ac25f48f43104318e3393b13b266d19371a29723  content-sidebar.js
418ebdf94f55508f7da5f892a740112d595928c3dfe10c01e5c58a791de92c35  vendor/fflate.min.js
```

**Xác nhận PO:** Cả 3 luồng (Lũy kế tháng – Tỉnh với tổ hợp F1.1/F1.2/F1.3, So sánh 2 kỳ – Bưu
cục, Lũy kế tháng – Bưu cục) đã được PO kiểm tra trên production và **PASS**.
