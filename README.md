# Email Reminder System - Nhắc nhở điền thiếu data

Tự động gửi email nhắc nhở nhân viên điền thiếu thông tin trong Google Sheets.

## Tính năng

- **Personalized emails**: Gửi email cá nhân với tên nhân viên
- **Smart filtering**: Chỉ check tháng <= hiện tại
- **Skip logic**: Bỏ qua hàng đã đánh dấu 'x' ở cột BX
- **Priority fields**: Ưu tiên các field quan trọng (ngày ký, doanh thu, số người khám)
- **Delay fields**: Các field có thời gian delay (chỉ check từ tháng hiện tại - 2 trở xuống)
- **Numeric handling**: Số 0 không tính là thiếu data

## Cấu hình

```javascript
const EMPLOYEE_EMAIL = "email@company.com"; // Thay email cho từng sheet
const SHEET_NAME = "file nhap chc";
```

## Các trường bắt buộc

**High Priority:**
- Ngày ký hợp đồng (D)
- Doanh thu (F)
- Số người khám (G)

**Standard:**
- Mã hợp đồng (C)
- Trạng thái ký (E)
- Ngày bắt đầu khám (J)
- Ngày kết thúc khám (K)

**Delay Fields** (chỉ check tháng <= hiện tại - 2):
- Ngày hóa đơn (H)
- Doanh thu thực hiện (I)
- Tháng GNDT (L)

## Cách sử dụng

### Functions chính

- `manualCheck()` - Test thủ công
- `dailyEmailCheck()` - Chạy tự động hàng ngày
- `testConfiguration()` - Kiểm tra cấu hình

### Setup trigger

```javascript
// Chạy thứ 3 và thứ 6
setupTuesdayFridayTriggers()

// Chạy hàng ngày
setupDailyTrigger()

// Xóa triggers
deleteTriggers()
```

### Skip rows

Đánh dấu 'x' hoặc 'X' ở cột BX để bỏ qua hàng đó trong email cảnh báo.

## Logic xử lý

1. **Filter months**: Chỉ xử lý tháng <= tháng hiện tại
2. **Check skip flag**: Bỏ qua nếu cột BX có 'x'
3. **Scan required fields**: Check từng field bắt buộc
4. **Group by month**: Nhóm missing data theo tháng
5. **Send email**: Gửi email với format HTML + plain text

## Debugging

```javascript
// Check cấu hình
testConfiguration()

// Xem log console để debug missing data logic
```

## Notes

- Delay fields chỉ áp dụng cho tháng <= (hiện tại - 2)
- Numeric fields: 0 = OK, empty/null = missing
- Text fields: empty/whitespace = missing
- Email format: HTML + fallback plain text