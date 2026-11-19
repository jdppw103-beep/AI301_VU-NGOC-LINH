using System.Globalization;
using System.Text;
using System.Text.Json;

if (args.Contains("--build-slides"))
{
    var outPptx = Path.Combine(Directory.GetCurrentDirectory(), "output", "Ngoc_Linh_Data_Studio_Product_Overview.pptx");
    var shotsDir = Path.Combine(Path.GetTempPath(), "slides_shots");
    PptxDeckBuilder.Build(outPptx, shotsDir);
    Console.WriteLine($"BUILT_PPTX={outPptx}");
    return;
}

var builder = WebApplication.CreateBuilder(args);
builder.WebHost.UseUrls("http://localhost:5000");
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy => policy.AllowAnyOrigin().AllowAnyHeader().AllowAnyMethod());
});

var app = builder.Build();
app.UseCors();
app.UseDefaultFiles();
app.UseStaticFiles();

// Khởi động với trạng thái trống - chờ khách hàng tải dữ liệu lên
var dataHub = new DataAnalystHub();

// ================= API ENDPOINTS CHO HỆ THỐNG BI 4 BƯỚC =================

// Overview trạng thái toàn hệ thống
app.MapGet("/api/status", () => Results.Ok(dataHub.GetPipelineStatus()));

// Xoá toàn bộ dữ liệu phiên làm việc
app.MapPost("/api/reset", () =>
{
    dataHub.Reset();
    return Results.Ok(new { message = "Đã xoá dữ liệu. Hệ thống sẵn sàng nhận tệp mới." });
});

// ---------------- BƯỚC 1: KẾT NỐI DỮ LIỆU ----------------
app.MapGet("/api/connect/template", () =>
{
    var header = "OrderCode,OrderDate,CustomerName,StoreName,Sku,ProductName,Quantity,UnitPrice,DiscountAmount,UnitCost,Status,Channel\n";
    var bytes = Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes(header)).ToArray();
    return Results.File(bytes, "text/csv", "DataLens_Template.csv");
});

app.MapGet("/api/connect/files", () => Results.Ok(new
{
    files = dataHub.GetFileSummaries(),
    summary = dataHub.GetRawSummary()
}));

app.MapPost("/api/connect/upload-json", (UploadParsedRequest req) =>
{
    if (req.Records == null || req.Records.Count == 0)
        return Results.BadRequest(new { message = "Dữ liệu tải lên trống." });

    var info = dataHub.AddOrReplaceFile(req.FileName ?? "Uploaded_Data.xlsx", req.Records, req.Headers ?? new List<string>(), req.ReplaceFileId);
    return Results.Ok(new
    {
        message = $"Đã kết nối thành công {req.Records.Count} bản ghi từ '{info.FileName}'!",
        file = info.ToSummary(),
        files = dataHub.GetFileSummaries(),
        summary = dataHub.GetRawSummary(),
        preview = dataHub.RawData.Take(50).ToList()
    });
});

app.MapPost("/api/connect/upload-file", async (HttpRequest request) =>
{
    if (!request.HasFormContentType)
        return Results.BadRequest(new { message = "Yêu cầu định dạng tệp hợp lệ (form-data)." });

    var form = await request.ReadFormAsync();
    var file = form.Files.GetFile("file");
    var replaceId = form["replaceFileId"].FirstOrDefault();
    if (file == null || file.Length == 0)
        return Results.BadRequest(new { message = "Không tìm thấy tệp đính kèm." });

    using var reader = new StreamReader(file.OpenReadStream(), Encoding.UTF8);
    var content = await reader.ReadToEndAsync();
    var (success, info, error) = dataHub.ParseAndAddFile(file.FileName, content, replaceId);

    if (!success || info == null)
        return Results.BadRequest(new { message = $"Lỗi đọc tệp: {error}" });

    return Results.Ok(new
    {
        message = $"Đã phân tích và kết nối thành công {info.RecordCount} bản ghi từ tệp '{file.FileName}'!",
        file = info.ToSummary(),
        files = dataHub.GetFileSummaries(),
        summary = dataHub.GetRawSummary(),
        preview = dataHub.RawData.Take(50).ToList()
    });
});

app.MapDelete("/api/connect/files/{id}", (string id) =>
{
    var removed = dataHub.RemoveFile(id);
    if (!removed) return Results.NotFound(new { message = "Không tìm thấy tệp cần xóa." });
    return Results.Ok(new
    {
        message = "Đã xóa tệp dữ liệu.",
        files = dataHub.GetFileSummaries(),
        summary = dataHub.GetRawSummary(),
        preview = dataHub.RawData.Take(50).ToList()
    });
});

app.MapGet("/api/connect/preview", () => Results.Ok(new
{
    files = dataHub.GetFileSummaries(),
    summary = dataHub.GetRawSummary(),
    records = dataHub.RawData.Take(50).ToList()
}));

// ---------------- BƯỚC 2: LÀM SẠCH DỮ LIỆU ----------------
app.MapPost("/api/clean/run", (CleaningConfig? config) =>
{
    if (dataHub.RawData.Count == 0)
        return Results.BadRequest(new { message = "Chưa kết nối dữ liệu ở Bước 1. Vui lòng tải tệp dữ liệu." });

    var result = dataHub.RunDataCleaning(config ?? new CleaningConfig());
    return Results.Ok(result);
});

app.MapPut("/api/clean/raw-record/{index:int}", (int index, GenericDataRecord updated) =>
{
    if (!dataHub.UpdateRawRecord(index, updated))
        return Results.NotFound(new { message = "Không tìm thấy bản ghi cần chỉnh sửa." });

    var result = dataHub.RunDataCleaning(dataHub.LastCleaningConfig);
    return Results.Ok(new
    {
        message = "Đã cập nhật bản ghi và làm sạch lại dữ liệu.",
        cleanResult = result
    });
});

app.MapDelete("/api/clean/raw-record/{index:int}", (int index) =>
{
    if (!dataHub.DeleteRawRecord(index))
        return Results.NotFound(new { message = "Không tìm thấy bản ghi cần xóa." });

    var result = dataHub.RunDataCleaning(dataHub.LastCleaningConfig);
    return Results.Ok(new
    {
        message = "Đã loại bỏ bản ghi và cập nhật dữ liệu sạch.",
        cleanResult = result
    });
});

app.MapPost("/api/clean/verify", (VerifyPayload payload) =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có dữ liệu đã làm sạch để xác nhận." });

    dataHub.IsHumanVerified = payload.Approved;
    dataHub.VerificationNotes = payload.Notes ?? (payload.Approved ? "Chuyên viên đã kiểm tra và phê duyệt số liệu sạch." : "Dữ liệu bị từ chối.");

    return Results.Ok(new
    {
        approved = dataHub.IsHumanVerified,
        notes = dataHub.VerificationNotes,
        message = dataHub.IsHumanVerified
            ? "Phê duyệt dữ liệu sạch thành công! Đã mở khóa Bước 3: Phân tích dữ liệu."
            : "Đã từ chối dữ liệu. Vui lòng rà soát hoặc chỉnh sửa bản ghi ngoại lệ."
    });
});

// ---------------- BƯỚC 3: PHÂN TÍCH DỮ LIỆU ----------------
app.MapGet("/api/analyze/dashboard", () =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Vui lòng chạy Bước 2: Làm sạch dữ liệu trước khi phân tích." });

    var report = dataHub.PerformDeepAnalysis();
    return Results.Ok(report);
});

app.MapPost("/api/analyze/evaluate", (VerifyPayload payload) =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có kết quả phân tích để đánh giá." });

    dataHub.IsResultEvaluated = payload.Approved;
    return Results.Ok(new
    {
        evaluated = dataHub.IsResultEvaluated,
        message = dataHub.IsResultEvaluated
            ? "Đã đánh giá kết quả phân tích đạt chuẩn. Sẵn sàng xuất báo cáo."
            : "Đã đặt lại trạng thái chờ đánh giá."
    });
});

app.MapPost("/api/analyze/ask-copilot", (AskQuery query) =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có dữ liệu sạch để truy vấn." });

    var answer = dataHub.AnswerAnalystQuestion(query.Question ?? "");
    return Results.Ok(answer);
});

// ---------------- BƯỚC 4: XUẤT KẾT QUẢ ----------------
app.MapGet("/api/export/excel", () =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có dữ liệu sạch để tạo Excel BI Workbook." });

    var analysis = dataHub.PerformDeepAnalysis();
    var bytes = ExcelBiExporter.BuildWorkbook(dataHub, analysis);
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Ngoc_Linh_Data_Studio_BI_Workbook.xlsx");
});

app.MapGet("/api/export/word", () =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có dữ liệu sạch để tạo Báo cáo Word BI." });

    var analysis = dataHub.PerformDeepAnalysis();
    var bytes = WordBiExporter.BuildReport(dataHub, analysis);
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Ngoc_Linh_Data_Studio_BI_Report.docx");
});

app.MapGet("/api/export/preview", () =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có dữ liệu làm sạch để xem trước." });

    var analysis = dataHub.PerformDeepAnalysis();
    return Results.Ok(new
    {
        generatedAt = DateTime.Now.ToString("dd/MM/yyyy HH:mm"),
        sourceName = dataHub.SourceName,
        isVerified = dataHub.IsHumanVerified,
        isEvaluated = dataHub.IsResultEvaluated,
        excel = new
        {
            fileName = "Ngoc_Linh_Data_Studio_BI_Workbook.xlsx",
            sheetCount = 6,
            sheets = new[]
            {
                new { name = "RAW DATA", desc = $"Dữ liệu gốc ({dataHub.RawData.Count} bản ghi, 12 trường)" },
                new { name = "CLEANED DATA", desc = $"Excel Table dữ liệu sạch ({dataHub.CleanedData.Count} bản ghi, 18 trường)" },
                new { name = "DATA DICTIONARY", desc = "Từ điển 18 trường dữ liệu chuẩn BI" },
                new { name = "CLEANING LOG", desc = $"Nhật ký làm sạch ({dataHub.Exceptions.Count} ngoại lệ, {dataHub.CleaningLogs.Count} dòng log)" },
                new { name = "ANALYSIS", desc = "Bảng tổng hợp công thức động (Doanh thu, Lợi nhuận, Chi nhánh, Sản phẩm, Thời gian, Kênh)" },
                new { name = "DASHBOARD", desc = "5 KPI Cards, 4 Biểu đồ Excel thực tế & Insight tổng quan" }
            },
            cleanRecords = dataHub.CleanedData.Count,
            rawRecords = dataHub.RawData.Count
        },
        word = new
        {
            fileName = "Ngoc_Linh_Data_Studio_BI_Report.docx",
            sectionCount = 11,
            sections = new[]
            {
                "Phần 1. Trang bìa — Ngọc Linh Data Studio",
                "Phần 2. Tóm tắt điều hành (Executive Summary)",
                "Phần 3. Giới thiệu & Mục tiêu phân tích",
                "Phần 4. Tổng quan & Đánh giá chất lượng dữ liệu",
                "Phần 5. Phương pháp & Hệ thống chỉ số KPI",
                "Phần 6. Kết quả phân tích kinh doanh (Kèm bảng & 3 biểu đồ)",
                "Phần 7. Business Insights (Hiện tượng → Bằng chứng → Phân tích → Tác động → Vấn đề)",
                "Phần 8. Đề xuất giải pháp & Khuyến nghị (3 mức ưu tiên)",
                "Phần 9. Kế hoạch triển khai (Ngắn / Trung / Dài hạn)",
                "Phần 10. Kết luận hỗ trợ ra quyết định",
                "Phần 11. Phụ lục & Từ điển dữ liệu"
            }
        },
        analysis
    });
});

app.MapGet("/api/export/csv", () =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest("Chưa có dữ liệu làm sạch để xuất.");

    var csv = dataHub.ExportToCsv();
    var bytes = Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes(csv)).ToArray();
    return Results.File(bytes, "text/csv", $"NgocLinhDataStudio_CleanedData_{DateTime.Now:yyyyMMdd_HHmm}.csv");
});

app.MapGet("/api/export/cleaned-records", () =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có dữ liệu làm sạch." });

    return Results.Ok(dataHub.CleanedData);
});

app.MapGet("/api/export/summary-report", () =>
{
    if (dataHub.CleanedData.Count == 0)
        return Results.BadRequest(new { message = "Chưa có dữ liệu làm sạch để xuất báo cáo." });

    var report = dataHub.PerformDeepAnalysis();
    return Results.Ok(new
    {
        generatedDate = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"),
        analystAgency = "Ngọc Linh Data Studio - BI Data Cleansing & Analytics Workbench",
        sourceName = dataHub.SourceName,
        executiveSummary = report.ExecutiveSummary,
        financialMetrics = report.OverallMetrics,
        pricingAnalysis = report.PricingAnalysis,
        storePerformance = report.StorePerformance,
        productPerformance = report.ProductPerformance,
        channelBreakdown = report.ChannelBreakdown,
        timeSeries = report.TimeSeries,
        structuredRecommendations = report.StructuredRecommendations,
        strategicRecommendations = report.StrategicRecommendations
    });
});

app.Run();

// ================= DATA ANALYST HUB & PROCESSING ENGINE =================

public class UploadParsedRequest
{
    public string? FileName { get; set; }
    public string? ReplaceFileId { get; set; }
    public List<string>? Headers { get; set; }
    public List<GenericDataRecord>? Records { get; set; }
}

public class CleaningConfig
{
    public bool RemoveDuplicates { get; set; } = true;
    public bool ExcludeCancelled { get; set; } = true;
    public bool DeductReturns { get; set; } = true;
    public bool ImputeMissingCogs { get; set; } = true;
    public bool NormalizeFormat { get; set; } = true;
}

public class VerifyPayload
{
    public bool Approved { get; set; } = true;
    public string? Notes { get; set; }
}

public class AskQuery
{
    public string? Question { get; set; }
}

public class GenericDataRecord
{
    public string OrderCode { get; set; } = "";
    public DateTime OrderDate { get; set; }
    public string CustomerName { get; set; } = "";
    public string StoreId { get; set; } = "";
    public string StoreName { get; set; } = "";
    public string Sku { get; set; } = "";
    public string ProductName { get; set; } = "";
    public int Quantity { get; set; } = 1;
    public decimal UnitPrice { get; set; }
    public decimal DiscountAmount { get; set; }
    public decimal UnitCost { get; set; }
    public string Status { get; set; } = "COMPLETED";
    public string Channel { get; set; } = "Store";
}

public class CleanedRecord : GenericDataRecord
{
    public decimal GrossRevenue { get; set; }
    public decimal NetRevenue { get; set; }
    public decimal TotalCost { get; set; }
    public decimal GrossProfit { get; set; }
    public decimal MarginPercent { get; set; }
    public string CleanRemark { get; set; } = "Hợp lệ";
}

public class ExceptionRecord
{
    public int RawIndex { get; set; }
    public string ExceptionType { get; set; } = ""; // DUPLICATE, CANCELLED, RETURNED, MISSING_DATA, INVALID_FORMAT
    public string Reason { get; set; } = "";
    public string ActionTaken { get; set; } = "";
    public GenericDataRecord Record { get; set; } = new();
}

public class UploadedFileInfo
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N")[..8];
    public string FileName { get; set; } = "";
    public int RecordCount { get; set; }
    public int ColumnCount { get; set; }
    public List<string> DetectedColumns { get; set; } = new();
    public List<string> MissingColumns { get; set; } = new();
    public string Status { get; set; } = "Hợp lệ";
    public DateTime UploadedAt { get; set; } = DateTime.Now;
    public List<GenericDataRecord> Records { get; set; } = new();

    public object ToSummary() => new
    {
        id = Id,
        fileName = FileName,
        recordCount = RecordCount,
        columnCount = ColumnCount,
        detectedColumns = DetectedColumns,
        missingColumns = MissingColumns,
        status = Status,
        uploadedAt = UploadedAt.ToString("dd/MM/yyyy HH:mm")
    };
}

public class DataAnalystHub
{
    public List<UploadedFileInfo> UploadedFiles { get; private set; } = new();
    public List<GenericDataRecord> RawData { get; private set; } = new();
    public List<CleanedRecord> CleanedData { get; private set; } = new();
    public List<ExceptionRecord> Exceptions { get; private set; } = new();
    public List<string> CleaningLogs { get; private set; } = new();
    public CleaningConfig LastCleaningConfig { get; private set; } = new();
    public CleaningStats LastStats { get; private set; } = new();
    public bool IsHumanVerified { get; set; } = false;
    public bool IsResultEvaluated { get; set; } = false;
    public string VerificationNotes { get; set; } = "";
    public string SourceName { get; set; } = "";

    private Dictionary<string, decimal> StandardCogs = new(StringComparer.OrdinalIgnoreCase);

    public void Reset()
    {
        UploadedFiles.Clear();
        RawData.Clear();
        CleanedData.Clear();
        Exceptions.Clear();
        CleaningLogs.Clear();
        LastStats = new();
        IsHumanVerified = false;
        IsResultEvaluated = false;
        VerificationNotes = "";
        SourceName = "";
    }

    public object GetPipelineStatus()
    {
        return new
        {
            step1_connected = RawData.Count > 0,
            step1_count = RawData.Count,
            step1_fileCount = UploadedFiles.Count,
            step1_source = SourceName,
            step2_cleaned = CleanedData.Count > 0,
            step2_count = CleanedData.Count,
            step2_exceptions = Exceptions.Count,
            step2_verified = IsHumanVerified,
            step3_analyzed = CleanedData.Count > 0 && IsHumanVerified,
            step3_evaluated = IsResultEvaluated,
            step4_ready = CleanedData.Count > 0 && IsHumanVerified
        };
    }

    public List<object> GetFileSummaries() => UploadedFiles.Select(f => f.ToSummary()).ToList();

    private void RebuildRawFromFiles()
    {
        RawData = UploadedFiles.SelectMany(f => f.Records).ToList();
        SourceName = UploadedFiles.Count switch
        {
            0 => "",
            1 => UploadedFiles[0].FileName,
            _ => $"{UploadedFiles.Count} tệp ({string.Join(", ", UploadedFiles.Select(f => f.FileName))})"
        };
        CleanedData.Clear();
        Exceptions.Clear();
        CleaningLogs.Clear();
        LastStats = new();
        IsHumanVerified = false;
        IsResultEvaluated = false;
    }

    public UploadedFileInfo AddOrReplaceFile(string fileName, List<GenericDataRecord> records, List<string> headers, string? replaceFileId = null)
    {
        // Chuẩn hóa sơ bộ các bản ghi
        foreach (var r in records)
        {
            if (r.OrderDate == default) r.OrderDate = DateTime.Today;
            r.Status = NormalizeStatus(r.Status);
            if (string.IsNullOrWhiteSpace(r.StoreName)) r.StoreName = "Chưa phân loại";
            if (string.IsNullOrWhiteSpace(r.Channel)) r.Channel = "Khác";
            if (string.IsNullOrWhiteSpace(r.Sku)) r.Sku = "N/A";
        }

        var required = new[] { "OrderCode", "ProductName", "Quantity", "UnitPrice" };
        var detected = headers.Count > 0 ? headers : new List<string> { "OrderCode", "OrderDate", "StoreName", "Sku", "ProductName", "Quantity", "UnitPrice", "DiscountAmount", "UnitCost", "Status", "Channel" };
        var missing = new List<string>();
        if (records.All(r => string.IsNullOrWhiteSpace(r.OrderCode))) missing.Add("Mã đơn");
        if (records.All(r => string.IsNullOrWhiteSpace(r.ProductName) || r.ProductName == "Không rõ tên")) missing.Add("Sản phẩm");
        if (records.All(r => r.UnitPrice == 0)) missing.Add("Đơn giá");

        var info = new UploadedFileInfo
        {
            FileName = fileName,
            RecordCount = records.Count,
            ColumnCount = detected.Count,
            DetectedColumns = detected,
            MissingColumns = missing,
            Status = missing.Count > 0 ? $"Thiếu: {string.Join(", ", missing)}" : "Cấu trúc chuẩn",
            Records = records
        };

        if (!string.IsNullOrWhiteSpace(replaceFileId))
        {
            var idx = UploadedFiles.FindIndex(f => f.Id == replaceFileId);
            if (idx >= 0) UploadedFiles[idx] = info;
            else UploadedFiles.Add(info);
        }
        else
        {
            UploadedFiles.Add(info);
        }

        RebuildRawFromFiles();
        return info;
    }

    public bool RemoveFile(string id)
    {
        var removed = UploadedFiles.RemoveAll(f => f.Id == id) > 0;
        if (removed) RebuildRawFromFiles();
        return removed;
    }

    public object GetRawSummary()
    {
        var stores = RawData.Select(r => r.StoreName).Distinct().Count();
        var skus = RawData.Select(r => r.Sku).Distinct().Count();
        var duplicates = RawData.GroupBy(r => $"{r.OrderCode}_{r.Sku}_{r.ProductName}".ToUpperInvariant()).Count(g => g.Count() > 1);
        var cancelled = RawData.Count(r => r.Status.Equals("CANCELLED", StringComparison.OrdinalIgnoreCase) || r.Status.Equals("FAILED", StringComparison.OrdinalIgnoreCase));
        var returns = RawData.Count(r => r.Status.Equals("RETURNED", StringComparison.OrdinalIgnoreCase));
        var missingOrInvalid = RawData.Count(r => r.Quantity <= 0 || r.UnitPrice <= 0 || string.IsNullOrWhiteSpace(r.OrderCode) || r.ProductName == "Không rõ tên");
        var grossValue = RawData.Sum(r => (r.Quantity * r.UnitPrice) - r.DiscountAmount);

        return new
        {
            totalRecords = RawData.Count,
            fileCount = UploadedFiles.Count,
            uniqueStores = stores,
            uniqueSkus = skus,
            duplicateGroups = duplicates,
            cancelledOrders = cancelled,
            returnedOrders = returns,
            missingOrInvalid = missingOrInvalid,
            estimatedRawGrossValue = grossValue,
            sourceName = SourceName
        };
    }

    public (bool Success, UploadedFileInfo? Info, string Error) ParseAndAddFile(string fileName, string content, string? replaceFileId = null)
    {
        try
        {
            var records = new List<GenericDataRecord>();
            var detectedHeaders = new List<string>();

            if (fileName.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
            {
                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var parsed = JsonSerializer.Deserialize<List<GenericDataRecord>>(content, options);
                if (parsed != null && parsed.Count > 0)
                {
                    records = parsed;
                    detectedHeaders = new List<string> { "OrderCode", "OrderDate", "CustomerName", "StoreName", "Sku", "ProductName", "Quantity", "UnitPrice", "DiscountAmount", "UnitCost", "Status", "Channel" };
                }
            }
            else
            {
                using var reader = new StringReader(content);
                string? headerLine = reader.ReadLine();
                if (headerLine == null) return (false, null, "Tệp trống.");

                var rawHeaders = headerLine.Split(new[] { ',', ';', '\t' }).Select(h => h.Trim().Trim('\uFEFF').Trim('"')).ToList();
                detectedHeaders = rawHeaders;
                var headers = rawHeaders.Select(h => h.ToLower()).ToList();

                int idxOrder = FindColIndex(headers, "ordercode", "order_code", "mã đơn", "mã đơn hàng", "order id", "orderid", "order", "code");
                int idxDate = FindColIndex(headers, "orderdate", "order_date", "ngày", "ngày đơn", "date");
                int idxStore = FindColIndex(headers, "storename", "store", "chi nhánh", "cửa hàng");
                int idxSku = FindColIndex(headers, "sku", "mã sp", "mã sản phẩm", "productcode");
                int idxName = FindColIndex(headers, "productname", "tên sản phẩm", "sản phẩm", "product", "item");
                int idxQty = FindColIndex(headers, "quantity", "số lượng", "qty", "sl");
                int idxPrice = FindColIndex(headers, "unitprice", "đơn giá", "price", "giá bán");
                int idxDisc = FindColIndex(headers, "discountamount", "discount", "giảm giá", "chiết khấu");
                int idxCost = FindColIndex(headers, "unitcost", "giá vốn", "cogs", "cost");
                int idxStatus = FindColIndex(headers, "status", "trạng thái", "tình trạng");
                int idxCust = FindColIndex(headers, "customername", "khách hàng", "customer");
                int idxChannel = FindColIndex(headers, "channel", "kênh", "kênh bán");

                string? line;
                int rowIdx = 1;
                while ((line = reader.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    rowIdx++;
                    var parts = SplitCsvLine(line);
                    var storeName = GetVal(parts, idxStore, "Chưa phân loại");

                    records.Add(new GenericDataRecord
                    {
                        OrderCode = GetVal(parts, idxOrder, $"ROW-{rowIdx}"),
                        StoreName = storeName,
                        StoreId = storeName,
                        Sku = GetVal(parts, idxSku, "N/A").ToUpper(),
                        ProductName = GetVal(parts, idxName, "Không rõ tên"),
                        Quantity = ParseInt(GetVal(parts, idxQty, "1")),
                        UnitPrice = ParseDecimal(GetVal(parts, idxPrice, "0")),
                        DiscountAmount = ParseDecimal(GetVal(parts, idxDisc, "0")),
                        UnitCost = ParseDecimal(GetVal(parts, idxCost, "0")),
                        Status = NormalizeStatus(GetVal(parts, idxStatus, "COMPLETED")),
                        CustomerName = GetVal(parts, idxCust, ""),
                        Channel = GetVal(parts, idxChannel, "Khác"),
                        OrderDate = ParseDate(GetVal(parts, idxDate, ""))
                    });
                }
            }

            if (records.Count == 0) return (false, null, "Không nhận diện được dòng dữ liệu hợp lệ nào.");

            var info = AddOrReplaceFile(fileName, records, detectedHeaders, replaceFileId);
            return (true, info, "");
        }
        catch (Exception ex)
        {
            return (false, null, ex.Message);
        }
    }

    public bool UpdateRawRecord(int index, GenericDataRecord updated)
    {
        if (index < 0 || index >= RawData.Count) return false;
        updated.Status = NormalizeStatus(updated.Status);
        if (updated.OrderDate == default) updated.OrderDate = RawData[index].OrderDate;
        RawData[index] = updated;
        return true;
    }

    public bool DeleteRawRecord(int index)
    {
        if (index < 0 || index >= RawData.Count) return false;
        RawData.RemoveAt(index);
        return true;
    }

    private int FindColIndex(List<string> headers, params string[] keywords)
    {
        foreach (var kw in keywords)
        {
            int exact = headers.FindIndex(h => h == kw);
            if (exact >= 0) return exact;
        }
        foreach (var kw in keywords)
        {
            int partial = headers.FindIndex(h => h.Contains(kw));
            if (partial >= 0) return partial;
        }
        return -1;
    }

    private static string NormalizeStatus(string? s)
    {
        var v = (s ?? "").Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(v)) return "COMPLETED";
        if (v.Contains("hoàn tất") || v.Contains("hoàn thành") || v.Contains("thành công") || v.Contains("complete") || v.Contains("success") || v.Contains("done")) return "COMPLETED";
        if (v.Contains("cancel") || v.Contains("hủy") || v.Contains("huỷ")) return "CANCELLED";
        if (v.Contains("fail") || v.Contains("lỗi") || v.Contains("thất bại")) return "FAILED";
        if (v.Contains("return") || v.Contains("hoàn") || v.Contains("trả")) return "RETURNED";
        return "COMPLETED";
    }

    private static DateTime ParseDate(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return DateTime.Today;
        string[] formats = { "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd", "dd-MM-yyyy", "MM/dd/yyyy", "dd/MM/yyyy HH:mm", "yyyy-MM-dd HH:mm:ss" };
        if (DateTime.TryParseExact(s.Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var d)) return d;
        return DateTime.TryParse(s, out d) ? d : DateTime.Today;
    }

    private string GetVal(List<string> parts, int index, string fallback)
    {
        if (index >= 0 && index < parts.Count && !string.IsNullOrWhiteSpace(parts[index]))
            return parts[index].Trim().Trim('"');
        return fallback;
    }

    private List<string> SplitCsvLine(string line)
    {
        var result = new List<string>();
        bool inQuotes = false;
        var sb = new StringBuilder();
        foreach (char c in line)
        {
            if (c == '\"') inQuotes = !inQuotes;
            else if ((c == ',' || c == ';' || c == '\t') && !inQuotes)
            {
                result.Add(sb.ToString());
                sb.Clear();
            }
            else sb.Append(c);
        }
        result.Add(sb.ToString());
        return result;
    }

    private int ParseInt(string s) => int.TryParse(s.Replace(".", "").Replace(",", ""), out int v) ? v : 1;
    private decimal ParseDecimal(string s) => decimal.TryParse(s.Replace(".", "").Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal v) ? v : 0;

    public object RunDataCleaning(CleaningConfig config)
    {
        LastCleaningConfig = config;
        CleaningLogs.Clear();
        CleanedData.Clear();
        Exceptions.Clear();
        IsHumanVerified = false;
        IsResultEvaluated = false;

        var seenKeys = new HashSet<string>();
        int duplicates = 0;
        int cancelled = 0;
        int returns = 0;
        int missingOrInvalid = 0;
        int imputed = 0;
        decimal preventedOvercount = 0;

        CleaningLogs.Add($"[AGENT DATA CLEANER] Bắt đầu quét {RawData.Count} bản ghi từ nguồn '{SourceName}'...");

        for (int idx = 0; idx < RawData.Count; idx++)
        {
            var r = RawData[idx];

            // Chuẩn hóa văn bản nếu bật
            if (config.NormalizeFormat)
            {
                r.OrderCode = (r.OrderCode ?? "").Trim().ToUpperInvariant();
                r.StoreName = (r.StoreName ?? "").Trim();
                r.ProductName = (r.ProductName ?? "").Trim();
                r.Sku = (r.Sku ?? "N/A").Trim().ToUpperInvariant();
                r.Channel = (r.Channel ?? "Khác").Trim();
                r.Status = NormalizeStatus(r.Status);
            }

            // Kiểm tra dữ liệu thiếu / sai định dạng
            if (r.Quantity <= 0 || r.UnitPrice <= 0 || string.IsNullOrWhiteSpace(r.OrderCode) || r.ProductName == "Không rõ tên")
            {
                missingOrInvalid++;
                Exceptions.Add(new ExceptionRecord
                {
                    RawIndex = idx,
                    ExceptionType = "MISSING_OR_INVALID",
                    Reason = r.Quantity <= 0 ? "Số lượng <= 0" : r.UnitPrice <= 0 ? "Thiếu đơn giá (<= 0)" : "Thiếu mã đơn hoặc tên sản phẩm",
                    ActionTaken = "Tạm giữ để kiểm tra",
                    Record = r
                });
                CleaningLogs.Add($"[NGOẠI LỆ] Dòng #{idx + 1} ({r.OrderCode}): thiếu dữ liệu hoặc sai định dạng (SL={r.Quantity}, Giá={r.UnitPrice:N0}).");
                continue;
            }

            var key = $"{r.OrderCode}_{r.Sku}_{r.ProductName}".ToUpperInvariant();

            // 1. Khử trùng
            if (config.RemoveDuplicates && seenKeys.Contains(key))
            {
                duplicates++;
                var dupVal = (r.Quantity * r.UnitPrice) - r.DiscountAmount;
                preventedOvercount += dupVal;
                Exceptions.Add(new ExceptionRecord
                {
                    RawIndex = idx,
                    ExceptionType = "DUPLICATE",
                    Reason = $"Trùng mã đơn & sản phẩm ({r.OrderCode})",
                    ActionTaken = "Đã loại bỏ bản ghi trùng",
                    Record = r
                });
                CleaningLogs.Add($"[DEDUPLICATE] Loại bỏ bản ghi trùng lặp: Đơn {r.OrderCode} | {r.ProductName} (Ngăn chặn {dupVal:N0} đ).");
                continue;
            }
            seenKeys.Add(key);

            // 2. Loại đơn hủy / thất bại
            if (config.ExcludeCancelled && (r.Status.Equals("CANCELLED", StringComparison.OrdinalIgnoreCase) || r.Status.Equals("FAILED", StringComparison.OrdinalIgnoreCase)))
            {
                cancelled++;
                var cancelVal = (r.Quantity * r.UnitPrice) - r.DiscountAmount;
                preventedOvercount += cancelVal;
                Exceptions.Add(new ExceptionRecord
                {
                    RawIndex = idx,
                    ExceptionType = "CANCELLED",
                    Reason = $"Trạng thái đơn: {r.Status}",
                    ActionTaken = "Loại khỏi doanh thu thuần",
                    Record = r
                });
                CleaningLogs.Add($"[EXCLUDE CANCELLED] Đơn {r.OrderCode} ({r.Status}) → loại khỏi doanh thu ({cancelVal:N0} đ).");
                continue;
            }

            // 3. Tra cứu hoặc nội suy giá vốn (COGS)
            decimal unitCost = r.UnitCost;
            if (unitCost <= 0)
            {
                if (StandardCogs.TryGetValue(r.Sku, out decimal stdCost)) unitCost = stdCost;
                else if (config.ImputeMissingCogs)
                {
                    unitCost = r.UnitPrice * 0.75m;
                    imputed++;
                }
            }

            decimal gross = r.Quantity * r.UnitPrice;
            decimal net = gross - r.DiscountAmount;
            decimal totalCost = r.Quantity * unitCost;

            // 4. Xử lý đổi trả
            string remark = unitCost != r.UnitCost && r.UnitCost <= 0 ? "Nội suy giá vốn" : "Hợp lệ";
            if (r.Status.Equals("RETURNED", StringComparison.OrdinalIgnoreCase))
            {
                returns++;
                remark = "Hàng hoàn (Giảm trừ)";
                Exceptions.Add(new ExceptionRecord
                {
                    RawIndex = idx,
                    ExceptionType = "RETURNED",
                    Reason = "Đơn hàng hoàn trả",
                    ActionTaken = config.DeductReturns ? "Đã giảm trừ DT & hoàn giá vốn" : "Giữ nguyên chờ đối soát",
                    Record = r
                });
                if (config.DeductReturns)
                {
                    gross = -gross;
                    net = -net;
                    totalCost = -totalCost;
                    CleaningLogs.Add($"[RETURN ADJUSTMENT] Đơn {r.OrderCode} là hàng trả lại → giảm trừ doanh thu & hoàn nhập giá vốn.");
                }
            }

            decimal grossProfit = net - totalCost;
            decimal marginPct = net > 0 ? Math.Round((grossProfit / net) * 100, 2) : 0;

            CleanedData.Add(new CleanedRecord
            {
                OrderCode = r.OrderCode,
                OrderDate = r.OrderDate,
                CustomerName = r.CustomerName,
                StoreId = r.StoreId,
                StoreName = r.StoreName,
                Sku = r.Sku,
                ProductName = r.ProductName,
                Quantity = r.Quantity,
                UnitPrice = r.UnitPrice,
                DiscountAmount = r.DiscountAmount,
                UnitCost = unitCost,
                GrossRevenue = gross,
                NetRevenue = net,
                TotalCost = totalCost,
                GrossProfit = grossProfit,
                MarginPercent = marginPct,
                Status = r.Status,
                Channel = r.Channel,
                CleanRemark = remark
            });
        }

        if (imputed > 0)
            CleaningLogs.Add($"[COGS IMPUTATION] {imputed} dòng thiếu giá vốn đã được áp dụng định mức tạm tính 75% đơn giá.");

        CleaningLogs.Add($"[AGENT COMPLETE] Hoàn tất làm sạch! {RawData.Count} dòng thô -> {CleanedData.Count} dòng sạch chuẩn hóa.");
        CleaningLogs.Add($"[STATISTICS] Trùng: {duplicates} · Hủy/lỗi: {cancelled} · Hoàn trả: {returns} · Ngoại lệ thiếu/sai: {missingOrInvalid} · Sai lệch ngăn ngừa: {preventedOvercount:N0} đ.");

        LastStats = new CleaningStats
        {
            RawCount = RawData.Count,
            CleanedCount = CleanedData.Count,
            DuplicatesRemoved = duplicates,
            CancelledExcluded = cancelled,
            ReturnsAdjusted = returns,
            MissingOrInvalid = missingOrInvalid,
            ImputedCount = imputed,
            PreventedOvercountValue = preventedOvercount
        };

        return new
        {
            success = true,
            rawCount = RawData.Count,
            cleanedCount = CleanedData.Count,
            duplicatesRemoved = duplicates,
            cancelledExcluded = cancelled,
            returnsAdjusted = returns,
            missingOrInvalid = missingOrInvalid,
            imputedCount = imputed,
            preventedOvercountValue = preventedOvercount,
            logs = CleaningLogs,
            exceptions = Exceptions,
            sampleCleaned = CleanedData.Take(100).ToList()
        };
    }

    public DeepAnalysisResult PerformDeepAnalysis()
    {
        var valid = CleanedData;
        decimal totalGross = valid.Sum(x => x.GrossRevenue > 0 ? x.GrossRevenue : 0);
        decimal totalDiscount = valid.Sum(x => x.Status == "RETURNED" ? 0 : x.DiscountAmount);
        decimal totalNet = valid.Sum(x => x.NetRevenue);
        decimal totalCost = valid.Sum(x => x.TotalCost);
        decimal totalProfit = totalNet - totalCost;
        decimal grossMargin = totalNet > 0 ? Math.Round((totalProfit / totalNet) * 100, 2) : 0;

        int totalOrders = valid.Count;
        int returnOrders = valid.Count(x => x.Status == "RETURNED");
        decimal returnRate = totalOrders > 0 ? Math.Round(((decimal)returnOrders / totalOrders) * 100, 2) : 0;
        decimal cancelRate = LastStats.RawCount > 0 ? Math.Round(((decimal)LastStats.CancelledExcluded / LastStats.RawCount) * 100, 2) : 0;

        // Phân tích giá bán & chiết khấu
        var nonReturnOrders = valid.Where(x => x.Status != "RETURNED").ToList();
        decimal avgOrderValue = nonReturnOrders.Count > 0 ? Math.Round(nonReturnOrders.Average(x => x.NetRevenue), 0) : 0;
        decimal avgDiscountRate = totalGross > 0 ? Math.Round((totalDiscount / totalGross) * 100, 2) : 0;
        int discountedOrderCount = nonReturnOrders.Count(x => x.DiscountAmount > 0);
        decimal maxDiscountOrderPct = nonReturnOrders.Where(x => x.GrossRevenue > 0)
            .Select(x => Math.Round((x.DiscountAmount / x.GrossRevenue) * 100, 2))
            .DefaultIfEmpty(0).Max();

        // Phân tích theo Chi nhánh
        var storeResults = valid.GroupBy(x => x.StoreName).Select(g =>
        {
            decimal net = g.Sum(x => x.NetRevenue);
            decimal cost = g.Sum(x => x.TotalCost);
            decimal profit = net - cost;
            decimal margin = net > 0 ? Math.Round((profit / net) * 100, 2) : 0;
            int rets = g.Count(x => x.Status == "RETURNED");
            decimal retRate = g.Count() > 0 ? Math.Round(((decimal)rets / g.Count()) * 100, 2) : 0;

            string storeStatus = "Vận hành tốt";
            if (margin < 12) storeStatus = "Biên lợi nhuận mỏng";
            else if (retRate > 15) storeStatus = "Cảnh báo đổi trả cao";

            return new StoreAnalysisItem
            {
                StoreName = g.Key,
                OrderCount = g.Count(),
                NetRevenue = net,
                GrossProfit = profit,
                MarginPercent = margin,
                ReturnCount = rets,
                ReturnRatePercent = retRate,
                StatusAssessment = storeStatus
            };
        }).OrderByDescending(s => s.NetRevenue).ToList();

        // Phân tích theo Sản phẩm
        var productResults = valid
            .GroupBy(x => new { x.Sku, x.ProductName })
            .Select(g =>
            {
                decimal net = g.Sum(x => x.NetRevenue);
                decimal cost = g.Sum(x => x.TotalCost);
                decimal profit = net - cost;
                decimal disc = g.Sum(x => x.Status == "RETURNED" ? 0 : x.DiscountAmount);
                return new ProductAnalysisItem
                {
                    Sku = g.Key.Sku,
                    ProductName = g.Key.ProductName,
                    QuantitySold = g.Sum(x => x.Status == "RETURNED" ? -x.Quantity : x.Quantity),
                    NetRevenue = net,
                    GrossProfit = profit,
                    DiscountAmount = disc,
                    MarginPercent = net > 0 ? Math.Round((profit / net) * 100, 2) : 0
                };
            })
            .OrderByDescending(p => p.NetRevenue)
            .ToList();

        // Phân tích Kênh bán
        var channelResults = valid
            .GroupBy(x => x.Channel)
            .Select(g => new ChannelAnalysisItem
            {
                ChannelName = g.Key,
                OrderCount = g.Count(),
                NetRevenue = g.Sum(x => x.NetRevenue),
                RevenueSharePercent = totalNet > 0 ? Math.Round((g.Sum(x => x.NetRevenue) / totalNet) * 100, 1) : 0
            })
            .OrderByDescending(c => c.NetRevenue)
            .ToList();

        // Phân tích tăng trưởng theo thời gian
        var timeSeries = valid
            .GroupBy(x => x.OrderDate.Date)
            .OrderBy(g => g.Key)
            .Select(g => new TimeSeriesPoint
            {
                Date = g.Key.ToString("yyyy-MM-dd"),
                Label = g.Key.ToString("dd/MM"),
                NetRevenue = g.Sum(x => x.NetRevenue),
                GrossProfit = g.Sum(x => x.GrossProfit),
                Orders = g.Count()
            })
            .ToList();

        decimal growthRate = 0;
        if (timeSeries.Count >= 2 && timeSeries.First().NetRevenue > 0)
        {
            growthRate = Math.Round(((timeSeries.Last().NetRevenue - timeSeries.First().NetRevenue) / timeSeries.First().NetRevenue) * 100, 2);
        }

        // Khuyến nghị có cấu trúc theo 5 nhóm nghiệp vụ (dựa 100% trên dữ liệu thực tế)
        string today = DateTime.Now.ToString("dd/MM/yyyy");
        var structuredRecs = new List<StructuredRecommendation>();

        // 1. Giá bán và chiết khấu
        structuredRecs.Add(new StructuredRecommendation
        {
            Category = "Giá bán và chiết khấu",
            Title = avgDiscountRate > 8 ? "Tỷ lệ chiết khấu bình quân vượt ngưỡng 8%" : "Cơ cấu giá bán và chiết khấu ổn định",
            Content = $"Tỷ lệ chiết khấu trung bình đạt {avgDiscountRate}% ({totalDiscount:N0} đ trên {discountedOrderCount} đơn có giảm giá, cao nhất {maxDiscountOrderPct}%). Giá trị đơn bình quân: {avgOrderValue:N0} đ.",
            DetectedDate = today,
            Priority = avgDiscountRate > 8 || maxDiscountOrderPct > 15 ? "Cao" : "Trung bình"
        });

        // 2. Tăng trưởng doanh thu
        structuredRecs.Add(new StructuredRecommendation
        {
            Category = "Tăng trưởng doanh thu",
            Title = timeSeries.Count >= 2 ? $"Xu hướng doanh thu {(growthRate >= 0 ? "tăng" : "giảm")} {Math.Abs(growthRate)}% qua {timeSeries.Count} ngày" : "Tổng hợp doanh thu thuần kỳ hiện tại",
            Content = timeSeries.Count >= 2
                ? $"Doanh thu thuần ngày đầu ({timeSeries.First().Label}) đạt {timeSeries.First().NetRevenue:N0} đ, ngày cuối ({timeSeries.Last().Label}) đạt {timeSeries.Last().NetRevenue:N0} đ."
                : $"Tổng doanh thu thuần đạt {totalNet:N0} đ trên {totalOrders} giao dịch hợp lệ.",
            DetectedDate = today,
            Priority = growthRate < 0 ? "Cao" : "Trung bình"
        });

        // 3. Hiệu quả sản phẩm
        var topProd = productResults.FirstOrDefault();
        var lowProd = productResults.Where(p => p.MarginPercent < 10).FirstOrDefault();
        if (topProd != null)
        {
            structuredRecs.Add(new StructuredRecommendation
            {
                Category = "Hiệu quả sản phẩm",
                Title = $"Sản phẩm dẫn đầu: {topProd.ProductName}",
                Content = lowProd != null && lowProd.Sku != topProd.Sku
                    ? $"'{topProd.ProductName}' đóng góp {topProd.NetRevenue:N0} đ (biên {topProd.MarginPercent}%). Lưu ý '{lowProd.ProductName}' có biên lợi nhuận thấp ({lowProd.MarginPercent}%)."
                    : $"'{topProd.ProductName}' bán {topProd.QuantitySold} đơn vị, mang lại {topProd.NetRevenue:N0} đ doanh thu và {topProd.GrossProfit:N0} đ lợi nhuận gộp.",
                DetectedDate = today,
                Priority = lowProd != null ? "Cao" : "Thấp"
            });
        }

        // 4. Hiệu quả chi nhánh
        var lowStores = storeResults.Where(s => s.MarginPercent < 12 || s.ReturnRatePercent > 15).ToList();
        var bestStore = storeResults.FirstOrDefault();
        if (bestStore != null)
        {
            structuredRecs.Add(new StructuredRecommendation
            {
                Category = "Hiệu quả chi nhánh",
                Title = lowStores.Count > 0 ? $"{lowStores.Count} chi nhánh cần hỗ trợ vận hành" : $"Chi nhánh dẫn đầu: {bestStore.StoreName}",
                Content = lowStores.Count > 0
                    ? $"Cần rà soát {string.Join(", ", lowStores.Select(s => $"{s.StoreName} (biên {s.MarginPercent}%, hoàn {s.ReturnRatePercent}%)"))}. Dẫn đầu hiện là {bestStore.StoreName} ({bestStore.NetRevenue:N0} đ)."
                    : $"Dẫn đầu là {bestStore.StoreName} đạt {bestStore.NetRevenue:N0} đ (biên {bestStore.MarginPercent}%). Tất cả {storeResults.Count} chi nhánh vận hành ổn định.",
                DetectedDate = today,
                Priority = lowStores.Count > 0 ? "Cao" : "Thấp"
            });
        }

        // 5. Quản trị lợi nhuận
        structuredRecs.Add(new StructuredRecommendation
        {
            Category = "Quản trị lợi nhuận",
            Title = grossMargin < 15 ? "Biên lợi nhuận gộp cần tối ưu" : "Biên lợi nhuận gộp đạt ngưỡng an toàn",
            Content = $"Lợi nhuận gộp đạt {totalProfit:N0} đ (biên {grossMargin}%). Đã ngăn chặn {LastStats.PreventedOvercountValue:N0} đ sai lệch từ {LastStats.DuplicatesRemoved} đơn trùng và {LastStats.CancelledExcluded} đơn hủy.",
            DetectedDate = today,
            Priority = grossMargin < 15 || returnRate > 8 ? "Cao" : "Trung bình"
        });

        var recommendations = structuredRecs.Select(r => $"[{r.Category.ToUpper()}] {r.Title}: {r.Content}").ToList();

        string execSummary = $"Phân tích {valid.Count} bản ghi hợp lệ từ '{SourceName}': doanh thu thuần đạt {totalNet:N0} đ, lợi nhuận gộp {totalProfit:N0} đ (biên {grossMargin}%). Đã loại bỏ {LastStats.DuplicatesRemoved} đơn trùng, {LastStats.CancelledExcluded} đơn hủy/lỗi và đối soát {LastStats.ReturnsAdjusted} đơn hoàn trả.";

        return new DeepAnalysisResult
        {
            ExecutiveSummary = execSummary,
            OverallMetrics = new OverallMetrics
            {
                TotalGrossRevenue = totalGross,
                TotalDiscountAmount = totalDiscount,
                TotalNetRevenue = totalNet,
                TotalCost = totalCost,
                TotalGrossProfit = totalProfit,
                GrossMarginPercent = grossMargin,
                TotalOrders = totalOrders,
                ReturnOrders = returnOrders,
                ReturnRatePercent = returnRate,
                CancelRatePercent = cancelRate,
                PreventedOvercountValue = LastStats.PreventedOvercountValue,
                IsVerified = IsHumanVerified,
                IsEvaluated = IsResultEvaluated
            },
            PricingAnalysis = new PricingAnalysis
            {
                AverageOrderValue = avgOrderValue,
                AverageDiscountRate = avgDiscountRate,
                MaxDiscountPercent = maxDiscountOrderPct,
                DiscountedOrderCount = discountedOrderCount,
                TotalDiscountAmount = totalDiscount
            },
            StorePerformance = storeResults,
            ProductPerformance = productResults,
            ChannelBreakdown = channelResults,
            TimeSeries = timeSeries,
            GrowthRatePercent = growthRate,
            StructuredRecommendations = structuredRecs,
            StrategicRecommendations = recommendations
        };
    }

    public object AnswerAnalystQuestion(string question)
    {
        var analysis = PerformDeepAnalysis();
        var q = question.ToLower();

        string answer;
        var matchedStore = analysis.StorePerformance.FirstOrDefault(s => q.Contains(s.StoreName.ToLower()));
        if (matchedStore != null)
        {
            answer = $"{matchedStore.StoreName}: doanh thu thuần {matchedStore.NetRevenue:N0} đ, lợi nhuận gộp {matchedStore.GrossProfit:N0} đ (biên {matchedStore.MarginPercent}%), {matchedStore.OrderCount} giao dịch, tỷ lệ hoàn {matchedStore.ReturnRatePercent}%. Đánh giá: {matchedStore.StatusAssessment}.";
        }
        else if (q.Contains("chiết khấu") || q.Contains("giá bán"))
        {
            answer = $"Giá trị đơn trung bình là {analysis.PricingAnalysis.AverageOrderValue:N0} đ. Tổng chiết khấu {analysis.PricingAnalysis.TotalDiscountAmount:N0} đ (tỷ lệ bình quân {analysis.PricingAnalysis.AverageDiscountRate}%, cao nhất {analysis.PricingAnalysis.MaxDiscountPercent}%).";
        }
        else if (q.Contains("tăng trưởng") || q.Contains("thời gian") || q.Contains("xu hướng"))
        {
            answer = analysis.TimeSeries.Count >= 2
                ? $"Qua {analysis.TimeSeries.Count} ngày ghi nhận, tốc độ biến động doanh thu từ ngày đầu đến ngày cuối là {analysis.GrowthRatePercent}%."
                : $"Dữ liệu hiện tập trung trong 1 ngày ({analysis.TimeSeries.FirstOrDefault()?.Label ?? "hôm nay"}) với doanh thu thuần {analysis.OverallMetrics.TotalNetRevenue:N0} đ.";
        }
        else if (q.Contains("doanh thu") || q.Contains("lợi nhuận") || q.Contains("kpi") || q.Contains("biên"))
        {
            answer = $"Tổng doanh thu thuần đạt {analysis.OverallMetrics.TotalNetRevenue:N0} đ, lợi nhuận gộp {analysis.OverallMetrics.TotalGrossProfit:N0} đ (biên lợi nhuận {analysis.OverallMetrics.GrossMarginPercent}%).";
        }
        else if (q.Contains("cửa hàng") || q.Contains("chi nhánh") || q.Contains("rủi ro"))
        {
            var bestStore = analysis.StorePerformance.FirstOrDefault();
            var riskStore = analysis.StorePerformance.OrderBy(s => s.MarginPercent).FirstOrDefault();
            answer = bestStore != null
                ? $"Chi nhánh dẫn đầu doanh thu là {bestStore.StoreName} ({bestStore.NetRevenue:N0} đ). Biên lợi nhuận thấp nhất là {riskStore?.StoreName} ({riskStore?.MarginPercent}%, hoàn {riskStore?.ReturnRatePercent}%)."
                : "Chưa có dữ liệu chi nhánh.";
        }
        else if (q.Contains("sản phẩm") || q.Contains("bán chạy") || q.Contains("sku"))
        {
            var topProd = analysis.ProductPerformance.FirstOrDefault();
            answer = topProd != null
                ? $"Sản phẩm dẫn đầu doanh thu là {topProd.ProductName} ({topProd.Sku}) với {topProd.QuantitySold} đơn vị, đạt {topProd.NetRevenue:N0} đ (biên {topProd.MarginPercent}%)."
                : "Chưa có dữ liệu sản phẩm.";
        }
        else if (q.Contains("hàng trả") || q.Contains("đổi trả") || q.Contains("hoàn"))
        {
            answer = $"Tỷ lệ hoàn trả hiện tại là {analysis.OverallMetrics.ReturnRatePercent}% ({analysis.OverallMetrics.ReturnOrders} đơn). Đã hạch toán giảm trừ doanh thu và hoàn nhập giá vốn.";
        }
        else
        {
            answer = $"{analysis.ExecutiveSummary} Bạn có thể hỏi chi tiết về doanh thu, lợi nhuận, chiết khấu, tăng trưởng, từng chi nhánh hoặc sản phẩm.";
        }

        return new
        {
            question,
            answer,
            keyMetrics = new
            {
                netRevenue = analysis.OverallMetrics.TotalNetRevenue,
                grossMargin = analysis.OverallMetrics.GrossMarginPercent,
                orders = analysis.OverallMetrics.TotalOrders
            }
        };
    }

    public string ExportToCsv()
    {
        var sb = new StringBuilder();
        sb.AppendLine("Mã Đơn,Ngày Đơn,Khách Hàng,Chi Nhánh,SKU,Tên Sản Phẩm,Số Lượng,Đơn Giá,Chiết Khấu,Giá Vốn Đơn Vị,Doanh Thu Thuần,Tổng Giá Vốn,Lợi Nhuận Gộp,Biên LN %,Trạng Thái,Kênh,Ghi Chú");

        foreach (var r in CleanedData)
        {
            sb.AppendLine($"\"{r.OrderCode}\",\"{r.OrderDate:dd/MM/yyyy}\",\"{r.CustomerName}\",\"{r.StoreName}\",\"{r.Sku}\",\"{r.ProductName}\",{r.Quantity},{r.UnitPrice},{r.DiscountAmount},{r.UnitCost},{r.NetRevenue},{r.TotalCost},{r.GrossProfit},{r.MarginPercent},\"{r.Status}\",\"{r.Channel}\",\"{r.CleanRemark}\"");
        }

        return sb.ToString();
    }
}

public class CleaningStats
{
    public int RawCount { get; set; }
    public int CleanedCount { get; set; }
    public int DuplicatesRemoved { get; set; }
    public int CancelledExcluded { get; set; }
    public int ReturnsAdjusted { get; set; }
    public int MissingOrInvalid { get; set; }
    public int ImputedCount { get; set; }
    public decimal PreventedOvercountValue { get; set; }
}

public class DeepAnalysisResult
{
    public string ExecutiveSummary { get; set; } = "";
    public OverallMetrics OverallMetrics { get; set; } = new();
    public PricingAnalysis PricingAnalysis { get; set; } = new();
    public List<StoreAnalysisItem> StorePerformance { get; set; } = new();
    public List<ProductAnalysisItem> ProductPerformance { get; set; } = new();
    public List<ChannelAnalysisItem> ChannelBreakdown { get; set; } = new();
    public List<TimeSeriesPoint> TimeSeries { get; set; } = new();
    public decimal GrowthRatePercent { get; set; }
    public List<StructuredRecommendation> StructuredRecommendations { get; set; } = new();
    public List<string> StrategicRecommendations { get; set; } = new();
}

public class OverallMetrics
{
    public decimal TotalGrossRevenue { get; set; }
    public decimal TotalDiscountAmount { get; set; }
    public decimal TotalNetRevenue { get; set; }
    public decimal TotalCost { get; set; }
    public decimal TotalGrossProfit { get; set; }
    public decimal GrossMarginPercent { get; set; }
    public int TotalOrders { get; set; }
    public int ReturnOrders { get; set; }
    public decimal ReturnRatePercent { get; set; }
    public decimal CancelRatePercent { get; set; }
    public decimal PreventedOvercountValue { get; set; }
    public bool IsVerified { get; set; }
    public bool IsEvaluated { get; set; }
}

public class PricingAnalysis
{
    public decimal AverageOrderValue { get; set; }
    public decimal AverageDiscountRate { get; set; }
    public decimal MaxDiscountPercent { get; set; }
    public int DiscountedOrderCount { get; set; }
    public decimal TotalDiscountAmount { get; set; }
}

public class StoreAnalysisItem
{
    public string StoreName { get; set; } = "";
    public int OrderCount { get; set; }
    public decimal NetRevenue { get; set; }
    public decimal GrossProfit { get; set; }
    public decimal MarginPercent { get; set; }
    public int ReturnCount { get; set; }
    public decimal ReturnRatePercent { get; set; }
    public string StatusAssessment { get; set; } = "";
}

public class ProductAnalysisItem
{
    public string Sku { get; set; } = "";
    public string ProductName { get; set; } = "";
    public int QuantitySold { get; set; }
    public decimal NetRevenue { get; set; }
    public decimal GrossProfit { get; set; }
    public decimal DiscountAmount { get; set; }
    public decimal MarginPercent { get; set; }
}

public class ChannelAnalysisItem
{
    public string ChannelName { get; set; } = "";
    public int OrderCount { get; set; }
    public decimal NetRevenue { get; set; }
    public decimal RevenueSharePercent { get; set; }
}

public class TimeSeriesPoint
{
    public string Date { get; set; } = "";
    public string Label { get; set; } = "";
    public decimal NetRevenue { get; set; }
    public decimal GrossProfit { get; set; }
    public int Orders { get; set; }
}

public class StructuredRecommendation
{
    public string Category { get; set; } = "";
    public string Title { get; set; } = "";
    public string Content { get; set; } = "";
    public string DetectedDate { get; set; } = "";
    public string Priority { get; set; } = "Trung bình";
}
