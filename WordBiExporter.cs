using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;

public static class WordBiExporter
{
    private const string FONT_NAME = "Times New Roman";
    private const string COLOR_PRIMARY = "3C4430";
    private const string COLOR_ACCENT = "56613F";
    private const string COLOR_MUTED = "555555";
    private const string COLOR_BORDER = "CCCCCC";
    private const string COLOR_TH_BG = "56613F";
    private const string COLOR_ALT_BG = "F4F4F1";

    public static byte[] BuildReport(DataAnalystHub hub, DeepAnalysisResult analysis)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body!;

            // Settings: UpdateFieldsOnOpen = true for automatic TOC
            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings(new UpdateFieldsOnOpen { Val = true });
            settingsPart.Settings.Save();

            // Styles
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = BuildStyles();
            stylesPart.Styles.Save();

            // Header & Footer
            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(
                MakePara("Ngọc Linh Data Studio — Báo cáo Phân tích Dữ liệu Kinh doanh (BI)", 10, false, false, JustificationValues.Right, "888888")
            );
            headerPart.Header.Save();

            var footerPart = mainPart.AddNewPart<FooterPart>();
            var fPara = new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Right }));
            fPara.Append(
                MakeRun("Báo cáo BI nội bộ  |  Trang ", 10, false, false, "888888"),
                new SimpleField { Instruction = "PAGE" },
                MakeRun(" / ", 10, false, false, "888888"),
                new SimpleField { Instruction = "NUMPAGES" }
            );
            footerPart.Footer = new Footer(fPara);
            footerPart.Footer.Save();

            // Helper metrics
            var m = analysis.OverallMetrics;
            var p = analysis.PricingAnalysis;
            var st = hub.LastStats;
            string dateRange = analysis.TimeSeries.Count > 0
                ? $"{analysis.TimeSeries.First().Label} – {analysis.TimeSeries.Last().Label}"
                : DateTime.Now.ToString("dd/MM/yyyy");
            string reportDate = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

            // ================= PHẦN 1. TRANG BÌA =================
            body.Append(MakePara("", 13));
            body.Append(MakePara("NGỌC LINH DATA STUDIO", 16, true, false, JustificationValues.Center, COLOR_PRIMARY));
            body.Append(MakePara("BI Data Cleansing & Analytics Workbench", 12, false, true, JustificationValues.Center, COLOR_ACCENT));
            body.Append(MakePara("", 24));
            body.Append(MakePara("BÁO CÁO PHÂN TÍCH DỮ LIỆU KINH DOANH", 20, true, false, JustificationValues.Center, COLOR_PRIMARY));
            body.Append(MakePara("VÀ ĐỀ XUẤT GIẢI PHÁP QUẢN TRỊ", 20, true, false, JustificationValues.Center, COLOR_PRIMARY));
            body.Append(MakePara("", 14));
            body.Append(MakePara($"Bộ dữ liệu phân tích: {hub.SourceName}", 13, true, false, JustificationValues.Center));
            body.Append(MakePara($"Khoảng thời gian dữ liệu: {dateRange}", 13, false, false, JustificationValues.Center));
            body.Append(MakePara($"Số bản ghi hợp lệ: {m.TotalOrders:N0} / {st.RawCount:N0} bản ghi đầu vào", 13, false, false, JustificationValues.Center));
            body.Append(MakePara("", 36));
            body.Append(MakePara($"Đơn vị thực hiện: Ngọc Linh Data Studio", 13, true, false, JustificationValues.Center, COLOR_PRIMARY));
            body.Append(MakePara($"Ngày lập báo cáo: {reportDate}", 13, false, true, JustificationValues.Center, COLOR_MUTED));
            body.Append(PageBreak());

            // ================= MỤC LỤC TỰ ĐỘNG =================
            body.Append(MakePara("MỤC LỤC", 15, true, false, JustificationValues.Center, COLOR_PRIMARY));
            body.Append(new Paragraph(new SimpleField { Instruction = "TOC \\o \"1-3\" \\h \\z \\u" }));
            body.Append(MakePara("(Nhấp chuột phải vào Mục lục và chọn 'Update Field' trong Microsoft Word để cập nhật số trang)", 10, false, true, JustificationValues.Center, "777777"));
            body.Append(PageBreak());

            // ================= PHẦN 2. TÓM TẮT ĐIỀU HÀNH =================
            body.Append(Heading1("PHẦN 1. TÓM TẮT ĐIỀU HÀNH (EXECUTIVE SUMMARY)"));
            body.Append(BodyPara($"Mục tiêu của báo cáo này là chuẩn hóa bộ dữ liệu '{hub.SourceName}', loại bỏ các sai lệch nghiệp vụ (đơn trùng lặp, đơn hủy/lỗi, hạch toán hoàn trả) và khai thác các chỉ số Business Intelligence (BI) phục vụ Ban Giám đốc ra quyết định."));
            body.Append(BodyPara($"Tổng quan quy mô: Từ {st.RawCount:N0} bản ghi thô, hệ thống đã làm sạch còn {m.TotalOrders:N0} bản ghi hợp lệ (loại {st.DuplicatesRemoved} đơn trùng, {st.CancelledExcluded} đơn hủy/lỗi, hạch toán {st.ReturnsAdjusted} đơn hoàn trả, ngăn chặn {m.PreventedOvercountValue:N0} đ doanh thu ảo)."));
            body.Append(BodyPara($"Kết quả tài chính cốt lõi: Tổng doanh thu thuần đạt {m.TotalNetRevenue:N0} đ, tổng giá vốn {m.TotalCost:N0} đ, lợi nhuận gộp {m.TotalGrossProfit:N0} đ (biên lợi nhuận trung bình {m.GrossMarginPercent}%). Giá trị đơn hàng trung bình đạt {p.AverageOrderValue:N0} đ; tỷ lệ chiết khấu bình quân {p.AverageDiscountRate}%."));

            body.Append(Heading2("Các Insight trọng tâm & Định hướng ưu tiên"));
            foreach (var rec in analysis.StructuredRecommendations)
            {
                body.Append(BulletPara($"[{rec.Category} · Ưu tiên {rec.Priority}] {rec.Title}: {rec.Content}"));
            }

            // ================= PHẦN 3. GIỚI THIỆU VÀ MỤC TIÊU PHÂN TÍCH =================
            body.Append(Heading1("PHẦN 2. GIỚI THIỆU VÀ MỤC TIÊU PHÂN TÍCH"));
            body.Append(Heading2("2.1. Bối cảnh phân tích"));
            body.Append(BodyPara("Trong quá trình tổng hợp dữ liệu bán hàng từ nhiều tệp/chi nhánh, việc đối soát thủ công dễ phát sinh đếm trùng đơn hàng, ghi nhận cả đơn hủy vào doanh thu và thiếu chuẩn hóa giá vốn. Báo cáo được thực hiện tự động trên Ngọc Linh Data Studio theo quy trình BI khép kín: Kết nối → Làm sạch → Kiểm duyệt → Phân tích → Xuất kết quả."));

            body.Append(Heading2("2.2. Mục tiêu của báo cáo"));
            body.Append(BulletPara("Đảm bảo tính chính xác, minh bạch của số liệu doanh thu, giá vốn và lợi nhuận gộp."));
            body.Append(BulletPara("Đánh giá hiệu quả kinh doanh theo từng sản phẩm, chi nhánh, kênh bán hàng và xu hướng thời gian."));
            body.Append(BulletPara("Phát hiện các điểm bất thường (biên lợi nhuận thấp, hoàn hàng cao, chiết khấu sâu) và đề xuất giải pháp quản trị cụ thể."));

            body.Append(Heading2("2.3. Các câu hỏi kinh doanh cần giải quyết (Business Questions)"));
            body.Append(BulletPara("Doanh thu thuần và lợi nhuận gộp thực tế sau khi loại bỏ đơn trùng và đơn hủy là bao nhiêu?"));
            body.Append(BulletPara("Chi nhánh và sản phẩm nào đang đóng góp hiệu quả nhất, đâu là điểm nghẽn cần hỗ trợ?"));
            body.Append(BulletPara("Chính sách chiết khấu và tỷ lệ hoàn trả đang ảnh hưởng như thế nào đến biên lợi nhuận?"));

            body.Append(Heading2("2.4. Phạm vi phân tích"));
            body.Append(BodyPara($"Thời gian dữ liệu: {dateRange}. Đối tượng: Toàn bộ {m.TotalOrders:N0} giao dịch hợp lệ thuộc {analysis.StorePerformance.Count} chi nhánh và {analysis.ProductPerformance.Count} dòng sản phẩm trong nguồn '{hub.SourceName}'."));

            // ================= PHẦN 4. TỔNG QUAN VÀ ĐÁNH GIÁ CHẤT LƯỢNG DỮ LIỆU =================
            body.Append(Heading1("PHẦN 3. TỔNG QUAN VÀ ĐÁNH GIÁ CHẤT LƯỢNG DỮ LIỆU"));
            body.Append(Heading2("3.1. Nguồn dữ liệu và cấu trúc"));
            body.Append(BodyPara($"Nguồn dữ liệu đầu vào: {hub.SourceName} ({hub.UploadedFiles.Count} tệp). Các trường thông tin chính gồm: Mã đơn, Ngày đơn, Khách hàng, Chi nhánh, SKU, Tên sản phẩm, Số lượng, Đơn giá, Chiết khấu, Giá vốn đơn vị, Trạng thái và Kênh bán."));

            body.Append(Heading2("3.2. Quy mô và Đánh giá chất lượng dữ liệu trước & sau làm sạch"));
            body.Append(BodyPara("Bảng 1 trình bày kết quả kiểm tra chất lượng dữ liệu, đối soát ngoại lệ và tác động của quy trình làm sạch:"));
            body.Append(CaptionPara("Bảng 1: Thống kê chất lượng dữ liệu trước và sau khi làm sạch"));
            body.Append(BuildTable(
                new[] { "Chỉ tiêu kiểm soát", "Số lượng / Giá trị", "Quy tắc xử lý áp dụng" },
                new[] {
                    new[] { "Tổng bản ghi đầu vào", $"{st.RawCount:N0}", "Tiếp nhận nguyên trạng từ tệp tải lên" },
                    new[] { "Bản ghi trùng lặp", $"{st.DuplicatesRemoved:N0}", "Khử trùng theo khóa (Mã đơn + SKU + Sản phẩm)" },
                    new[] { "Đơn hủy / lỗi thanh toán", $"{st.CancelledExcluded:N0}", "Loại khỏi doanh thu thuần (CANCELLED / FAILED)" },
                    new[] { "Đơn hàng hoàn trả", $"{st.ReturnsAdjusted:N0}", "Ghi nhận giảm trừ doanh thu & hoàn giá vốn" },
                    new[] { "Dữ liệu thiếu / sai định dạng", $"{st.MissingOrInvalid:N0}", "Tách bảng ngoại lệ để chuyên viên kiểm tra/sửa" },
                    new[] { "Dòng thiếu giá vốn được nội suy", $"{st.ImputedCount:N0}", "Áp dụng định mức 75% đơn giá khi thiếu giá vốn" },
                    new[] { "Bản ghi hợp lệ sau làm sạch", $"{st.CleanedCount:N0}", hub.IsHumanVerified ? "Đã kiểm duyệt & phê duyệt" : "Đã làm sạch" },
                    new[] { "Sai lệch doanh thu ngăn chặn", $"{st.PreventedOvercountValue:N0} đ", "Tổng giá trị đơn trùng và đơn hủy đã loại" }
                },
                new[] { 3400, 2200, 3400 }
            ));

            // ================= PHẦN 5. PHƯƠNG PHÁP VÀ HỆ THỐNG CHỈ SỐ PHÂN TÍCH =================
            body.Append(Heading1("PHẦN 4. PHƯƠNG PHÁP VÀ HỆ THỐNG CHỈ SỐ PHÂN TÍCH"));
            body.Append(Heading2("4.1. Phương pháp phân tích"));
            body.Append(BodyPara("Báo cáo áp dụng phương pháp phân tích mô tả (Descriptive Analytics) kết hợp chẩn đoán bất thường (Diagnostic Analytics) trên tập dữ liệu đã làm sạch: phân tổ đa chiều theo Thời gian, Chi nhánh, Sản phẩm và Kênh bán."));

            body.Append(Heading2("4.2. Hệ thống KPI và công thức tính toán"));
            body.Append(CaptionPara("Bảng 2: Danh mục chỉ số KPI và công thức tính toán chuẩn hóa"));
            body.Append(BuildTable(
                new[] { "Tên KPI", "Công thức tính toán", "Đơn vị" },
                new[] {
                    new[] { "Doanh thu thuần", "∑(Số lượng × Đơn giá - Chiết khấu) - ∑Giá trị đơn hoàn", "VNĐ" },
                    new[] { "Lợi nhuận gộp", "Doanh thu thuần - Tổng giá vốn (COGS)", "VNĐ" },
                    new[] { "Biên lợi nhuận gộp", "(Lợi nhuận gộp / Doanh thu thuần) × 100%", "%" },
                    new[] { "Giá trị đơn trung bình (AOV)", "Doanh thu thuần đơn bán / Số đơn bán hợp lệ", "VNĐ" },
                    new[] { "Tỷ lệ chiết khấu bình quân", "(Tổng chiết khấu / Tổng doanh thu gộp) × 100%", "%" },
                    new[] { "Tỷ lệ hoàn trả", "(Số đơn RETURNED / Tổng đơn hợp lệ) × 100%", "%" },
                    new[] { "Tốc độ tăng trưởng kỳ", "((DT ngày cuối - DT ngày đầu) / DT ngày đầu) × 100%", "%" }
                },
                new[] { 2800, 5000, 1200 }
            ));

            body.Append(Heading2("4.3. Giới hạn của dữ liệu và phương pháp"));
            body.Append(BodyPara(st.ImputedCount > 0
                ? $"Lưu ý: Trong bộ dữ liệu có {st.ImputedCount} dòng thiếu thông tin giá vốn đơn vị, hệ thống đã áp dụng định mức tạm tính 75% đơn giá. Ngoài ra, phân tích chỉ phản ánh các trường có mặt trong tệp đầu vào, không suy diễn chi phí vận hành ngoài dữ liệu."
                : "Phân tích được thực hiện hoàn toàn trên các trường dữ liệu thực tế có trong tệp đầu vào; chi phí vận hành cố định ngoài giá vốn không nằm trong phạm vi tệp dữ liệu."));

            // ================= PHẦN 6. KẾT QUẢ PHÂN TÍCH KINH DOANH =================
            body.Append(Heading1("PHẦN 5. KẾT QUẢ PHÂN TÍCH KINH DOANH"));

            body.Append(Heading2("5.1. Tổng quan doanh thu, lợi nhuận và chính sách giá"));
            body.Append(BodyPara($"Tổng doanh thu gộp đạt {m.TotalGrossRevenue:N0} đ; sau khi trừ {m.TotalDiscountAmount:N0} đ chiết khấu và khấu trừ các đơn hoàn trả, doanh thu thuần đạt {m.TotalNetRevenue:N0} đ. Tổng giá vốn ở mức {m.TotalCost:N0} đ, tạo ra lợi nhuận gộp {m.TotalGrossProfit:N0} đ (biên lợi nhuận {m.GrossMarginPercent}%)."));
            body.Append(BodyPara($"Về giá bán & chiết khấu: Giá trị đơn hàng trung bình (AOV) đạt {p.AverageOrderValue:N0} đ. Có {p.DiscountedOrderCount} đơn hàng áp dụng giảm giá, tỷ lệ chiết khấu bình quân {p.AverageDiscountRate}% (mức chiết khấu cao nhất ghi nhận trên 1 đơn là {p.MaxDiscountPercent}%)."));

            body.Append(Heading2("5.2. Phân tích xu hướng kinh doanh theo thời gian"));
            if (analysis.TimeSeries.Count >= 2)
            {
                body.Append(BodyPara($"Dữ liệu ghi nhận biến động qua {analysis.TimeSeries.Count} ngày giao dịch ({dateRange}), với tốc độ tăng trưởng từ ngày đầu đến ngày cuối đạt {analysis.GrowthRatePercent}%."));
                var chart1Bytes = RenderLineChartPng("Xu hướng Doanh thu thuần theo thời gian (VNĐ)", analysis.TimeSeries);
                body.Append(ImagePara(mainPart, chart1Bytes, "chart_trend.png", 5400000L, 2500000L));
                body.Append(CaptionPara("Hình 1: Biểu đồ xu hướng Doanh thu thuần theo thời gian (Đơn vị: VNĐ)"));
            }
            else
            {
                body.Append(BodyPara($"Dữ liệu hiện tập trung trong 1 ngày ({dateRange}) với doanh thu thuần {m.TotalNetRevenue:N0} đ, chưa đủ chuỗi nhiều ngày để đánh giá xu hướng dài hạn."));
            }

            body.Append(Heading2("5.3. Phân tích hiệu quả chi nhánh"));
            body.Append(CaptionPara("Bảng 3: Kết quả kinh doanh chi tiết theo từng chi nhánh"));
            body.Append(BuildTable(
                new[] { "Chi nhánh", "Số đơn", "Doanh thu thuần (đ)", "Lợi nhuận gộp (đ)", "Biên LN %", "Tỷ lệ hoàn %", "Đánh giá" },
                analysis.StorePerformance.Select(s => new[] {
                    s.StoreName, $"{s.OrderCount}", $"{s.NetRevenue:N0}", $"{s.GrossProfit:N0}", $"{s.MarginPercent}%", $"{s.ReturnRatePercent}%", s.StatusAssessment
                }).ToArray(),
                new[] { 2000, 800, 1600, 1600, 1000, 1000, 1400 }
            ));

            if (analysis.StorePerformance.Count > 0)
            {
                var chart2Bytes = RenderBarChartPng("So sánh Doanh thu thuần theo Chi nhánh (VNĐ)",
                    analysis.StorePerformance.Take(8).Select(s => (s.StoreName, (double)s.NetRevenue)).ToList(),
                    System.Drawing.Color.FromArgb(86, 97, 63));
                body.Append(ImagePara(mainPart, chart2Bytes, "chart_stores.png", 5400000L, 2500000L));
                body.Append(CaptionPara("Hình 2: So sánh Doanh thu thuần giữa các chi nhánh (Đơn vị: VNĐ)"));
            }

            body.Append(Heading2("5.4. Phân tích hiệu quả sản phẩm"));
            body.Append(CaptionPara("Bảng 4: Top sản phẩm theo Doanh thu thuần và Lợi nhuận gộp"));
            body.Append(BuildTable(
                new[] { "SKU", "Tên sản phẩm", "Sản lượng", "Doanh thu thuần (đ)", "Lợi nhuận gộp (đ)", "Biên LN %" },
                analysis.ProductPerformance.Take(10).Select(pr => new[] {
                    pr.Sku, pr.ProductName, $"{pr.QuantitySold}", $"{pr.NetRevenue:N0}", $"{pr.GrossProfit:N0}", $"{pr.MarginPercent}%"
                }).ToArray(),
                new[] { 1200, 2800, 900, 1700, 1700, 1100 }
            ));

            if (analysis.ProductPerformance.Count > 0)
            {
                var chart3Bytes = RenderBarChartPng("Top Sản phẩm theo Doanh thu thuần (VNĐ)",
                    analysis.ProductPerformance.Take(6).Select(pr => (pr.ProductName, (double)pr.NetRevenue)).ToList(),
                    System.Drawing.Color.FromArgb(20, 20, 20));
                body.Append(ImagePara(mainPart, chart3Bytes, "chart_products.png", 5400000L, 2500000L));
                body.Append(CaptionPara("Hình 3: Top sản phẩm đóng góp doanh thu cao nhất (Đơn vị: VNĐ)"));
            }

            body.Append(Heading2("5.5. Phân tích kênh bán hàng và chất lượng đơn hàng"));
            body.Append(CaptionPara("Bảng 5: Cơ cấu doanh thu theo kênh bán hàng"));
            body.Append(BuildTable(
                new[] { "Kênh bán", "Số đơn", "Doanh thu thuần (đ)", "Tỷ trọng %" },
                analysis.ChannelBreakdown.Select(c => new[] {
                    c.ChannelName, $"{c.OrderCount}", $"{c.NetRevenue:N0}", $"{c.RevenueSharePercent}%"
                }).ToArray(),
                new[] { 2500, 1500, 3000, 2000 }
            ));
            body.Append(BodyPara($"Chất lượng đơn hàng: Tỷ lệ hoàn trả ghi nhận {m.ReturnRatePercent}% ({m.ReturnOrders} đơn trên {m.TotalOrders} đơn hợp lệ); tỷ lệ đơn hủy/lỗi từ dữ liệu thô là {m.CancelRatePercent}%."));

            // ================= PHẦN 7. BUSINESS INSIGHTS =================
            body.Append(Heading1("PHẦN 6. BUSINESS INSIGHTS – PHÁT HIỆN VÀ ĐÁNH GIÁ VẤN ĐỀ"));
            int insIdx = 1;
            foreach (var rec in analysis.StructuredRecommendations)
            {
                string level = rec.Priority == "Cao" ? "Vấn đề nghiêm trọng cần ưu tiên" : rec.Priority == "Trung bình" ? "Cơ hội cải thiện hiệu quả" : "Xu hướng cần tiếp tục theo dõi";
                body.Append(Heading2($"Insight {insIdx++}: {rec.Title} ({level})"));
                body.Append(BulletPara($"Hiện tượng: Ghi nhận tại nhóm '{rec.Category}' vào ngày {rec.DetectedDate}."));
                body.Append(BulletPara($"Bằng chứng số liệu: {rec.Content}"));
                body.Append(BulletPara("Phân tích & Nguyên nhân đã chứng minh: Dữ liệu thực tế sau khi loại trừ trùng lặp và đơn hủy cho thấy mức độ chênh lệch trực tiếp trên chỉ số KPI của nhóm này."));
                body.Append(BulletPara("Tác động kinh doanh & Giả thuyết cần kiểm tra thêm: Ảnh hưởng trực tiếp đến biên lợi nhuận gộp và hiệu quả chuyển đổi; cần đối chiếu thêm với phản hồi khách hàng và định mức vận hành thực tế tại điểm bán."));
                body.Append(BulletPara($"Vấn đề cần giải quyết: Chuẩn hóa quy trình kiểm soát và tối ưu hóa chính sách thuộc nhóm {rec.Category}."));
            }

            // ================= PHẦN 8. ĐỀ XUẤT GIẢI PHÁP VÀ KHUYẾN NGHỊ =================
            body.Append(Heading1("PHẦN 7. ĐỀ XUẤT GIẢI PHÁP VÀ KHUYẾN NGHỊ"));
            body.Append(BodyPara("Các mục tiêu cải thiện dưới đây là mục tiêu dự kiến dựa trên hiện trạng dữ liệu hiện tại, không phải kết quả đã đạt được."));

            var highPriorityStores = analysis.StorePerformance.Where(s => s.MarginPercent < 12 || s.ReturnRatePercent > 15).ToList();
            var solutions = new[]
            {
                new {
                    Level = "1. Ưu tiên cao",
                    Title = "Kiểm soát chiết khấu & chuẩn hóa chất lượng dữ liệu đầu vào",
                    Problem = $"Tỷ lệ chiết khấu bình quân {p.AverageDiscountRate}% (tối đa {p.MaxDiscountPercent}%) và phát hiện {st.DuplicatesRemoved + st.CancelledExcluded} bản ghi trùng/hủy.",
                    Goal = "Mục tiêu dự kiến: Đưa sai số dữ liệu về 0% và kiểm soát trần chiết khấu theo phân cấp.",
                    Content = "Áp dụng quy trình đối soát 4 bước trên Ngọc Linh Data Studio trước mọi kỳ báo cáo; thiết lập ngưỡng phê duyệt quản lý khi chiết khấu vượt 10%.",
                    Owner = "Phòng Kinh doanh & Chuyên viên BI",
                    Resources = "Quy chế phê duyệt giá & Hệ thống Ngọc Linh Data Studio",
                    Kpi = "Tỷ lệ sai lệch báo cáo = 0%; Biên lợi nhuận gộp duy trì >= 15% (dự kiến).",
                    Timeline = "Triển khai ngay (0–2 tuần)"
                },
                new {
                    Level = "2. Ưu tiên trung bình",
                    Title = highPriorityStores.Count > 0 ? $"Tối ưu vận hành tại {string.Join(", ", highPriorityStores.Select(s => s.StoreName))}" : "Tối ưu cơ cấu sản phẩm & kênh bán chủ lực",
                    Problem = highPriorityStores.Count > 0 ? $"Chi nhánh có biên lợi nhuận mỏng hoặc tỷ lệ đổi trả cao (toàn hệ thống hoàn {m.ReturnRatePercent}%)." : "Tập trung nguồn lực vào các kênh và SKU mang lại tỷ suất sinh lời cao nhất.",
                    Goal = "Mục tiêu dự kiến: Giảm tỷ lệ hoàn hàng và cải thiện biên lợi nhuận chi nhánh.",
                    Content = "Rà soát quy trình kiểm hàng trước giao, điều chuyển tồn kho ưu tiên các SKU dẫn đầu doanh thu và kênh bán hiệu quả.",
                    Owner = "Bộ phận Vận hành & Quản lý Chi nhánh",
                    Resources = "Dữ liệu tồn kho WMS & Bảng phân tích chi nhánh",
                    Kpi = "Tỷ lệ hoàn trả đơn hàng; Doanh thu thuần theo kênh.",
                    Timeline = "Trung hạn (1–2 tháng)"
                },
                new {
                    Level = "3. Ưu tiên dài hạn",
                    Title = "Chuẩn hóa từ điển dữ liệu & tự động hóa báo cáo BI định kỳ",
                    Problem = st.ImputedCount > 0 ? $"Còn {st.ImputedCount} dòng thiếu giá vốn đơn vị cần nội suy." : "Cần duy trì tính nhất quán dữ liệu giữa các chi nhánh khi mở rộng quy mô.",
                    Goal = "Mục tiêu dự kiến: 100% giao dịch có đầy đủ giá vốn thực tế và chuẩn mã SKU.",
                    Content = "Áp dụng mẫu tệp chuẩn (Template) thống nhất cho toàn bộ chi nhánh, cập nhật giá vốn định kỳ hàng tháng.",
                    Owner = "Phòng Kế toán & Bộ phận Dữ liệu",
                    Resources = "Bộ từ điển dữ liệu chuẩn & Quy trình nhập liệu",
                    Kpi = "100% bản ghi đầy đủ trường thông tin bắt buộc.",
                    Timeline = "Dài hạn (3–6 tháng)"
                }
            };

            foreach (var sol in solutions)
            {
                body.Append(Heading2($"{sol.Level}: {sol.Title}"));
                body.Append(BulletPara($"Vấn đề cần giải quyết: {sol.Problem}"));
                body.Append(BulletPara($"Mục tiêu: {sol.Goal}"));
                body.Append(BulletPara($"Nội dung thực hiện: {sol.Content}"));
                body.Append(BulletPara($"Bộ phận chịu trách nhiệm: {sol.Owner}"));
                body.Append(BulletPara($"Nguồn lực cần thiết: {sol.Resources}"));
                body.Append(BulletPara($"KPI đánh giá hiệu quả: {sol.Kpi}"));
                body.Append(BulletPara($"Thời gian triển khai: {sol.Timeline}"));
                body.Append(BulletPara($"Kết quả kỳ vọng: Nâng cao tính chính xác báo cáo điều hành và tối ưu hóa biên lợi nhuận thực tế."));
            }

            // ================= PHẦN 9. KẾ HOẠCH TRIỂN KHAI =================
            body.Append(Heading1("PHẦN 8. KẾ HOẠCH TRIỂN KHAI"));
            body.Append(CaptionPara("Bảng 6: Lộ trình triển khai giải pháp theo giai đoạn"));
            body.Append(BuildTable(
                new[] { "Giai đoạn", "Hành động trọng tâm", "Đơn vị phụ trách", "Chỉ số theo dõi" },
                new[] {
                    new[] { "Ngắn hạn (Tuần 1-2)", "Chuẩn hóa quy trình lọc trùng, loại đơn hủy qua Ngọc Linh Data Studio; kiểm soát chiết khấu.", "Chuyên viên BI & Kinh doanh", "Số lỗi dữ liệu = 0; Tỷ lệ CK TB" },
                    new[] { "Trung hạn (Tháng 1-2)", "Rà soát các chi nhánh biên LN thấp/hoàn cao; tối ưu tồn kho SKU chủ lực.", "Vận hành & Quản lý Chi nhánh", "Tỷ lệ hoàn %; Biên LN chi nhánh" },
                    new[] { "Dài hạn (Quý 1-2)", "Đồng bộ giá vốn thực 100%, chuẩn hóa mẫu nhập liệu và báo cáo quản trị định kỳ.", "Kế toán & Ban Giám đốc", "Độ phủ dữ liệu chuẩn = 100%" }
                },
                new[] { 1800, 3400, 2000, 1800 }
            ));

            // ================= PHẦN 10. KẾT LUẬN =================
            body.Append(Heading1("PHẦN 9. KẾT LUẬN"));
            body.Append(BodyPara($"Thông qua việc xử lý và làm sạch dữ liệu thực tế từ '{hub.SourceName}', Ngọc Linh Data Studio đã giúp loại bỏ {st.DuplicatesRemoved + st.CancelledExcluded} bản ghi sai lệch (ngăn chặn {m.PreventedOvercountValue:N0} đ doanh thu ảo), xác lập bức tranh tài chính tin cậy với doanh thu thuần {m.TotalNetRevenue:N0} đ và lợi nhuận gộp {m.TotalGrossProfit:N0} đ ({m.GrossMarginPercent}%)."));
            body.Append(BodyPara("Việc kết hợp quy trình kiểm duyệt con người (Human Verification) cùng hệ thống phân tích BI tự động tạo nền tảng vững chắc để Ban Giám đốc ưu tiên xử lý các điểm nóng về chiết khấu, hoàn hàng và tối ưu hiệu quả từng chi nhánh."));

            // ================= PHẦN 11. PHỤ LỤC VÀ TÀI LIỆU THAM KHẢO =================
            body.Append(Heading1("PHẦN 10. PHỤ LỤC VÀ NGUỒN DỮ LIỆU"));
            body.Append(Heading2("10.1. Nguồn dữ liệu thực tế sử dụng"));
            body.Append(BodyPara($"Nguồn tệp: {hub.SourceName}. Thời điểm lập báo cáo: {reportDate}. Hệ thống thực hiện: Ngọc Linh Data Studio (BI Data Cleansing & Analytics Workbench). Không sử dụng nguồn tham khảo giả lập bên ngoài."));

            body.Append(Heading2("10.2. Từ điển dữ liệu chuẩn hóa"));
            body.Append(CaptionPara("Bảng 7: Từ điển các trường dữ liệu phân tích"));
            body.Append(BuildTable(
                new[] { "Tên trường", "Kiểu dữ liệu", "Vai trò trong phân tích BI" },
                new[] {
                    new[] { "OrderCode", "Text", "Khóa chính khử trùng lặp và đếm số lượng đơn hàng" },
                    new[] { "OrderDate", "Date", "Phân tích chuỗi thời gian và tốc độ tăng trưởng" },
                    new[] { "StoreName", "Text", "Phân tích so sánh hiệu quả giữa các chi nhánh" },
                    new[] { "ProductName / Sku", "Text", "Đánh giá top sản phẩm doanh thu và lợi nhuận" },
                    new[] { "NetRevenue / GrossProfit", "Currency", "KPI cốt lõi đo lường quy mô và khả năng sinh lời" },
                    new[] { "Status", "Text", "Loại trừ đơn hủy và hạch toán giảm trừ hàng hoàn" }
                },
                new[] { 2400, 1600, 5000 }
            ));

            // Section properties: A4 + margins + header/footer + titlePg (schema order: headerReference, footerReference, pgSz, pgMar, titlePg)
            var sectPr = new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerPart) },
                new FooterReference { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(footerPart) },
                new PageSize { Width = 11906, Height = 16838 }, // A4
                new PageMargin { Top = 1440, Bottom = 1440, Right = 1134, Left = 1701, Header = 720, Footer = 720 },
                new TitlePage()
            );
            body.Append(sectPr);
            mainPart.Document.Save();
        }
        return ms.ToArray();
    }

    // ---------------- Chart PNG Renderers (System.Drawing.Common) ----------------
    private static byte[] RenderLineChartPng(string title, List<TimeSeriesPoint> points)
    {
        int w = 900, h = 400;
        using var bmp = new Bitmap(w, h);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(System.Drawing.Color.White);

        using var titleFont = new System.Drawing.Font("Times New Roman", 13, FontStyle.Bold);
        using var labelFont = new System.Drawing.Font("Times New Roman", 9);
        using var brushDark = new SolidBrush(System.Drawing.Color.FromArgb(20, 20, 20));
        using var brushOlive = new SolidBrush(System.Drawing.Color.FromArgb(86, 97, 63));
        using var penGrid = new Pen(System.Drawing.Color.FromArgb(230, 230, 225), 1);
        using var penLine = new Pen(System.Drawing.Color.FromArgb(86, 97, 63), 3);

        g.DrawString(title, titleFont, brushDark, 24, 14);
        int pl = 95, pr = 30, pt = 55, pb = 55;
        int cw = w - pl - pr, ch = h - pt - pb;

        double maxVal = Math.Max(1, points.Max(p => (double)p.NetRevenue));
        double minVal = Math.Min(0, points.Min(p => (double)p.NetRevenue));
        double span = Math.Max(1, maxVal - minVal);

        for (int i = 0; i <= 4; i++)
        {
            float y = pt + ch * (1f - i / 4f);
            g.DrawLine(penGrid, pl, y, pl + cw, y);
            double v = minVal + span * i / 4.0;
            g.DrawString(FormatShort(v), labelFont, Brushes.Gray, 12, y - 8);
        }

        var pts = new PointF[points.Count];
        for (int i = 0; i < points.Count; i++)
        {
            float x = points.Count == 1 ? pl + cw / 2f : pl + (float)i * cw / (points.Count - 1);
            float y = pt + ch * (1f - (float)(((double)points[i].NetRevenue - minVal) / span));
            pts[i] = new PointF(x, y);
        }

        if (pts.Length > 1) g.DrawLines(penLine, pts);
        for (int i = 0; i < pts.Length; i++)
        {
            g.FillEllipse(brushOlive, pts[i].X - 4, pts[i].Y - 4, 8, 8);
            if (points.Count <= 12 || i % Math.Max(1, points.Count / 8) == 0)
                g.DrawString(points[i].Label, labelFont, brushDark, pts[i].X - 15, pt + ch + 10);
        }

        using var ms = new MemoryStream();
        bmp.Save(ms, ImageFormat.Png);
        return ms.ToArray();
    }

    private static byte[] RenderBarChartPng(string title, List<(string Label, double Value)> items, System.Drawing.Color barColor)
    {
        int w = 900, h = 400;
        using var bmp = new Bitmap(w, h);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.Clear(System.Drawing.Color.White);

        using var titleFont = new System.Drawing.Font("Times New Roman", 13, FontStyle.Bold);
        using var labelFont = new System.Drawing.Font("Times New Roman", 9);
        using var brushDark = new SolidBrush(System.Drawing.Color.FromArgb(20, 20, 20));
        using var brushBar = new SolidBrush(barColor);
        using var penGrid = new Pen(System.Drawing.Color.FromArgb(230, 230, 225), 1);

        g.DrawString(title, titleFont, brushDark, 24, 14);
        int pl = 95, pr = 30, pt = 55, pb = 65;
        int cw = w - pl - pr, ch = h - pt - pb;

        double maxVal = Math.Max(1, items.Max(x => x.Value));
        for (int i = 0; i <= 4; i++)
        {
            float y = pt + ch * (1f - i / 4f);
            g.DrawLine(penGrid, pl, y, pl + cw, y);
            g.DrawString(FormatShort(maxVal * i / 4.0), labelFont, Brushes.Gray, 12, y - 8);
        }

        float slot = (float)cw / Math.Max(1, items.Count);
        float bw = Math.Min(65, slot * 0.55f);
        for (int i = 0; i < items.Count; i++)
        {
            float bh = (float)(Math.Max(0, items[i].Value) / maxVal * ch);
            float x = pl + i * slot + (slot - bw) / 2f;
            float y = pt + ch - bh;
            g.FillRectangle(brushBar, x, y, bw, bh);
            g.DrawString(FormatShort(items[i].Value), labelFont, brushDark, x, Math.Max(pt - 16, y - 16));
            string lbl = items[i].Label.Length > 15 ? items[i].Label[..14] + "…" : items[i].Label;
            g.DrawString(lbl, labelFont, brushDark, x - 4, pt + ch + 10);
        }

        using var ms = new MemoryStream();
        bmp.Save(ms, ImageFormat.Png);
        return ms.ToArray();
    }

    private static string FormatShort(double v)
    {
        double a = Math.Abs(v);
        if (a >= 1e9) return (v / 1e9).ToString("0.0") + " tỷ";
        if (a >= 1e6) return (v / 1e6).ToString("0.0") + " tr";
        return Math.Round(v).ToString("N0");
    }

    // ---------------- OpenXML Helpers ----------------
    private static Styles BuildStyles()
    {
        var styles = new Styles();
        styles.Append(new DocDefaults(
            new RunPropertiesDefault(new RunPropertiesBaseStyle(
                new RunFonts { Ascii = FONT_NAME, HighAnsi = FONT_NAME, ComplexScript = FONT_NAME },
                new FontSize { Val = "26" }, // 13pt
                new FontSizeComplexScript { Val = "26" }
            )),
            new ParagraphPropertiesDefault(new ParagraphPropertiesBaseStyle(
                new SpacingBetweenLines { After = "120", Line = "324", LineRule = LineSpacingRuleValues.Auto } // 1.35 line spacing
            ))
        ));

        styles.Append(MakeHeadingStyle("Heading1", "heading 1", "32", COLOR_PRIMARY, "240", "120"));
        styles.Append(MakeHeadingStyle("Heading2", "heading 2", "28", COLOR_ACCENT, "180", "80"));
        styles.Append(MakeHeadingStyle("Heading3", "heading 3", "26", "141414", "120", "60"));
        return styles;
    }

    private static Style MakeHeadingStyle(string id, string name, string sz, string color, string before, string after) =>
        new(
            new StyleName { Val = name },
            new BasedOn { Val = "Normal" },
            new NextParagraphStyle { Val = "Normal" },
            new UIPriority { Val = 9 },
            new PrimaryStyle(),
            new StyleParagraphProperties(new KeepNext(), new SpacingBetweenLines { Before = before, After = after, Line = "324", LineRule = LineSpacingRuleValues.Auto }),
            new StyleRunProperties(new RunFonts { Ascii = FONT_NAME, HighAnsi = FONT_NAME }, new Bold(), new W.Color { Val = color }, new FontSize { Val = sz }, new FontSizeComplexScript { Val = sz })
        ) { Type = StyleValues.Paragraph, StyleId = id };

    private static Paragraph Heading1(string text) => new(new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }), MakeRun(text, 16, true, false, COLOR_PRIMARY));
    private static Paragraph Heading2(string text) => new(new ParagraphProperties(new ParagraphStyleId { Val = "Heading2" }), MakeRun(text, 14, true, false, COLOR_ACCENT));
    private static Paragraph BodyPara(string text) => MakePara(text, 13, false, false, JustificationValues.Both);
    private static Paragraph CaptionPara(string text) => MakePara(text, 12, true, true, JustificationValues.Center, COLOR_MUTED);
    private static Paragraph BulletPara(string text) => MakePara("•  " + text, 13, false, false, JustificationValues.Both);
    private static Paragraph PageBreak() => new(new Run(new Break { Type = BreakValues.Page }));

    private static Paragraph MakePara(string text, int pt, bool bold = false, bool italic = false, JustificationValues? align = null, string? hexColor = null)
    {
        var pp = new ParagraphProperties(new SpacingBetweenLines { After = "120", Line = "324", LineRule = LineSpacingRuleValues.Auto });
        if (align.HasValue) pp.Append(new Justification { Val = align.Value });
        return new Paragraph(pp, MakeRun(text, pt, bold, italic, hexColor));
    }

    private static Run MakeRun(string text, int pt, bool bold = false, bool italic = false, string? hexColor = null)
    {
        var rp = new RunProperties(
            new RunFonts { Ascii = FONT_NAME, HighAnsi = FONT_NAME, ComplexScript = FONT_NAME }
        );
        if (bold) rp.Append(new Bold());
        if (italic) rp.Append(new Italic());
        if (!string.IsNullOrEmpty(hexColor)) rp.Append(new W.Color { Val = hexColor });
        rp.Append(
            new FontSize { Val = (pt * 2).ToString() },
            new FontSizeComplexScript { Val = (pt * 2).ToString() }
        );
        return new Run(rp, new W.Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve });
    }

    private static W.Table BuildTable(string[] headers, string[][] rows, int[] colWidths)
    {
        var tbl = new W.Table(new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4, Color = COLOR_BORDER },
                new LeftBorder { Val = BorderValues.Single, Size = 4, Color = COLOR_BORDER },
                new BottomBorder { Val = BorderValues.Single, Size = 4, Color = COLOR_BORDER },
                new RightBorder { Val = BorderValues.Single, Size = 4, Color = COLOR_BORDER },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = COLOR_BORDER },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = COLOR_BORDER }
            )
        ));

        var grid = new TableGrid();
        foreach (var w in colWidths) grid.Append(new GridColumn { Width = w.ToString() });
        tbl.Append(grid);

        // Header row
        var thRow = new TableRow(new TableRowProperties(new TableHeader(), new CantSplit()));
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Width = colWidths[i].ToString(), Type = TableWidthUnitValues.Dxa },
                    new Shading { Val = ShadingPatternValues.Clear, Fill = COLOR_TH_BG },
                    new TableCellMargin(new TopMargin { Width = "80", Type = TableWidthUnitValues.Dxa }, new LeftMargin { Width = "100", Type = TableWidthUnitValues.Dxa }, new BottomMargin { Width = "80", Type = TableWidthUnitValues.Dxa }, new RightMargin { Width = "100", Type = TableWidthUnitValues.Dxa })
                ),
                MakePara(headers[i], 11, true, false, JustificationValues.Left, "FFFFFF")
            );
            thRow.Append(cell);
        }
        tbl.Append(thRow);

        for (int r = 0; r < rows.Length; r++)
        {
            var tr = new TableRow(new TableRowProperties(new CantSplit()));
            string bg = r % 2 == 1 ? COLOR_ALT_BG : "FFFFFF";
            for (int c = 0; c < headers.Length; c++)
            {
                string val = c < rows[r].Length ? rows[r][c] : "";
                var cell = new TableCell(
                    new TableCellProperties(
                        new TableCellWidth { Width = colWidths[c].ToString(), Type = TableWidthUnitValues.Dxa },
                        new Shading { Val = ShadingPatternValues.Clear, Fill = bg },
                        new TableCellMargin(new TopMargin { Width = "70", Type = TableWidthUnitValues.Dxa }, new LeftMargin { Width = "100", Type = TableWidthUnitValues.Dxa }, new BottomMargin { Width = "70", Type = TableWidthUnitValues.Dxa }, new RightMargin { Width = "100", Type = TableWidthUnitValues.Dxa })
                    ),
                    MakePara(val, 11, false, false, JustificationValues.Left)
                );
                tr.Append(cell);
            }
            tbl.Append(tr);
        }
        return tbl;
    }

    private static Paragraph ImagePara(MainDocumentPart mainPart, byte[] pngBytes, string name, long cx, long cy)
    {
        var imgPart = mainPart.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream(pngBytes)) imgPart.FeedData(ms);
        string relId = mainPart.GetIdOfPart(imgPart);

        var drawing = new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = cx, Cy = cy },
                new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties { Id = (uint)Random.Shared.Next(100, 99999), Name = name },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = name },
                                new PIC.NonVisualPictureDrawingProperties()
                            ),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relId, CompressionState = A.BlipCompressionValues.Print },
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new PIC.ShapeProperties(
                                new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = cx, Cy = cy }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                            )
                        )
                    ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                )
            )
        );

        return new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }), new Run(drawing));
    }
}
