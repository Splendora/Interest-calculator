Công cụ tính lãi vay theo quy định của Ngân hàng Nhà nước

Tài liệu này mô tả phương pháp tính gốc, lãi và phạt trong các tập tin Excel do thư mục này sinh ra, cùng căn cứ pháp lý và hướng dẫn sử dụng.

1. Thành phần trong thư mục

Tập tin build_tinhLai_chiTiet.py sinh ra TinhLai_ChiTiet.xlsx, áp dụng cho khoản vay trả nhiều kỳ, dư nợ giảm dần, theo dõi theo từng ngày.

Tập tin build_tinhLai_donGian.py sinh ra TinhLai_DonGian.xlsx, áp dụng cho khoản vay trả một lần khi đáo hạn.

Ô nền xanh nhạt là ô nhập liệu. Ô nền xám là công thức tự tính, không chỉnh sửa trực tiếp.

2. Căn cứ pháp lý

Thông tư 39/2016/TT-NHNN, Điều 13: lãi suất quá hạn tối đa 150 phần trăm lãi suất trong hạn, bao gồm 100 phần trăm lãi theo dư nợ gốc và 50 phần trăm phạt tăng thêm. Lãi chậm trả đối với nợ lãi quá hạn tối đa 10 phần trăm một năm.

Thông tư 06/2023/TT-NHNN: thứ tự phân bổ khi có nợ quá hạn là gốc quá hạn, lãi của gốc quá hạn, gốc đến hạn, nợ lãi đến hạn, tiền phạt chậm trả, sau cùng là trả trước hạn.

Thông tư 14/2017/TT-NHNN và Nghị quyết 01/2019/NQ-HĐTP được viện dẫn cho phương pháp quy đổi lãi suất và nguyên tắc xét xử tranh chấp tín dụng.

Tiền phạt 10 phần trăm không được cộng vào dư nợ để tiếp tục tính lãi, nhằm tránh tình trạng lãi chồng lãi.

3. Trình tự nhập liệu bản chi tiết

Bước 1: nhập ngày chốt số tại sheet CẤU HÌNH.

Bước 2: kiểm tra lịch nghỉ lễ tại sheet NGÀY NGHỈ LỄ.

Bước 3: nhập các lần nhận tiền tại sheet GIẢI NGÂN.

Bước 4: nhập lịch hẹn trả gốc và lãi tại sheet LỊCH TRẢ NỢ.

Bước 5: nếu thay đổi lãi suất thì nhập tại sheet THAY ĐỔI LÃI SUẤT.

Bước 6: nhập các khoản đã trả thực tế tại sheet THANH TOÁN THỰC TẾ.

Bước 7: mở sheet BẢNG TÍNH, sao chép công thức xuống đến ngày chốt số.

Bước 8: xem kết quả tổng hợp tại sheet TỔNG HỢP.

4. Thứ tự phân bổ tiền trả trong ngày

Gọi Fr là tổng tiền khách hàng trả trong ngày. Việc phân bổ thực hiện như sau.

R là phần trả gốc quá hạn, bằng số nhỏ hơn giữa Fr và dư gốc quá hạn đầu ngày K.

S là phần trả lãi của gốc quá hạn, bằng số nhỏ hơn giữa phần tiền còn lại và tổng lãi loại này, gồm số dư đầu ngày M cộng phát sinh trong ngày P.

T là phần trả gốc đến hạn, bằng số nhỏ hơn giữa phần tiền còn lại và gốc đến hạn trong ngày D.

U là phần trả nợ lãi đến hạn, bằng số nhỏ hơn giữa phần tiền còn lại và tổng nợ lãi cơ sở, gồm số dư đầu ngày N cộng phần lãi vừa đến hạn.

V là phần trả tiền phạt 10 phần trăm, bằng số nhỏ hơn giữa phần tiền còn lại và tổng tiền phạt còn nợ cộng phát sinh trong ngày Q.

AD là phần trả trước hạn, bằng số nhỏ hơn giữa phần tiền còn lại và dư gốc trong hạn sau khi trừ gốc đến hạn.

Ô kiểm tra AC bằng F trừ đi toàn bộ các phần đã phân bổ, kết quả đúng phải bằng không.

5. Số dư cuối ngày

W là dư gốc trong hạn, bằng gốc đầu ngày J cộng giải ngân C trừ gốc đến hạn D trừ trả trước hạn AD.

X là dư gốc quá hạn, bằng gốc quá hạn đầu ngày K cộng phần gốc đến hạn chưa trả trừ phần đã trả R.

Y là lãi trong hạn chưa trả. Vào ngày đến hạn trả lãi, toàn bộ lãi được chuyển sang nợ lãi cơ sở nên Y bằng không. Ngày thường thì Y giữ nguyên.

Z là lãi của gốc quá hạn, bằng số dư M cộng phát sinh P trừ phần đã trả S.

AA là nợ lãi cơ sở, bằng số dư N cộng phần lãi mới đến hạn trừ phần đã trả U. Tuyệt đối không cộng tiền phạt Q vào đây.

AB là tiền phạt 10 phần trăm còn nợ, bằng phạt cũ cộng phát sinh Q trừ phần đã trả V. Số này không dùng để tính lãi tiếp.

Q là tiền phạt phát sinh trong ngày, bằng nợ lãi cơ sở N nhân lãi suất phạt I chia cho số ngày trong năm.

6. Các cột trong sheet BẢNG TÍNH

Cột A là ngày. Cột B cho biết có phải ngày làm việc không. Cột C là tiền nhận trong ngày. Cột D là gốc hẹn trả hôm nay. Cột E cho biết có hẹn trả lãi không. Cột F là tiền khách trả hôm nay.

Cột G, H, I là lãi suất trong hạn, quá hạn và phạt chậm trả áp dụng hôm nay.

Cột J, K là gốc đầu ngày trong hạn và quá hạn. Cột L là lãi trong hạn chưa trả. Cột M là lãi của gốc quá hạn. Cột N là nợ lãi cơ sở.

Cột O, P, Q là các khoản phát sinh trong ngày: lãi thường, lãi quá hạn, tiền phạt 10 phần trăm.

Cột R, S, T, U, V là các phần tiền trả được phân bổ theo thứ tự tại mục 4. Cột AD là trả trước hạn.

Cột W, X là gốc cuối ngày trong hạn và quá hạn. Cột Y, Z, AA, AB là các khoản lãi và phạt còn nợ. Cột AC để kiểm tra.

7. Bản đơn giản áp dụng khi nào

Khoản vay một lần, đến hạn trả cả gốc lẫn lãi thì dùng bản đơn giản cho gọn, không cần trừ dần từng ngày.

Nợ lãi tính riêng, tiền phạt tính trên nợ lãi đó, không cộng dồn nên không phát sinh lãi chồng lãi.

Khoản lãi 150 phần trăm được tách sẵn tại sheet TỔNG HỢP thành 100 phần trăm lãi gốc và 50 phần trăm phạt.

8. Kiểm soát trước khi trình số liệu

Đối chiếu số dư với sao kê của hệ thống ngân hàng.

Cập nhật lịch nghỉ Tết và nghỉ lễ hằng năm vì lịch các năm có thể thay đổi.

Xóa số liệu ví dụ trong tập tin mẫu, thay bằng số liệu thật của hồ sơ.
