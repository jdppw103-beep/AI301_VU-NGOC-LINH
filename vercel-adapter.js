/* Static-hosting adapter for Vercel: mirrors all /api/* endpoints in-browser when no C# backend is present */
(function (global) {
  'use strict';

  const origFetch = global.fetch.bind(global);
  let staticMode = null; // null = unknown, true = Vercel static, false = localhost C# backend

  const state = {
    files: [],
    raw: [],
    cleaned: [],
    exceptions: [],
    logs: [],
    stats: { rawCount: 0, cleanedCount: 0, duplicates: 0, cancelled: 0, returns: 0, missingOrInvalid: 0, imputed: 0, prevented: 0 },
    lastCfg: { removeDuplicates: true, excludeCancelled: true, deductReturns: true, normalizeFormat: true, imputeMissingCogs: true },
    isVerified: false,
    isEvaluated: false,
    sourceName: ''
  };

  const jsonRes = (obj, status = 200) =>
    new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json; charset=utf-8' } });

  function rebuildRaw() {
    state.raw = state.files.flatMap(f => f.records);
    state.sourceName = state.files.length === 0 ? '' : state.files.length === 1 ? state.files[0].fileName : `${state.files.length} tệp (${state.files.map(f => f.fileName).join(', ')})`;
    state.cleaned = [];
    state.exceptions = [];
    state.logs = [];
    state.isVerified = false;
    state.isEvaluated = false;
  }

  function addOrReplaceFile(fileName, records, headers, replaceId) {
    const missing = [];
    if (records.every(r => !r.orderCode)) missing.push('Mã đơn');
    if (records.every(r => !r.productName || r.productName === 'Không rõ tên')) missing.push('Sản phẩm');
    if (records.every(r => !r.unitPrice)) missing.push('Đơn giá');

    const info = {
      id: Math.random().toString(36).slice(2, 10),
      fileName,
      recordCount: records.length,
      columnCount: (headers && headers.length) || 12,
      missingColumns: missing,
      status: missing.length ? `Thiếu: ${missing.join(', ')}` : 'Cấu trúc chuẩn',
      records
    };
    if (replaceId) {
      const idx = state.files.findIndex(f => f.id === replaceId);
      if (idx >= 0) state.files[idx] = info; else state.files.push(info);
    } else {
      state.files.push(info);
    }
    rebuildRaw();
    return info;
  }

  function fileSummaries() {
    return state.files.map(f => ({
      id: f.id, fileName: f.fileName, recordCount: f.recordCount, columnCount: f.columnCount,
      missingColumns: f.missingColumns, status: f.status
    }));
  }

  function rawSummary() {
    const s = DataLens.summarize(state.raw);
    return {
      totalRecords: s.totalRecords, fileCount: state.files.length,
      uniqueStores: s.uniqueStores, uniqueSkus: s.uniqueSkus,
      duplicateGroups: s.duplicateGroups, cancelledOrders: s.cancelledOrders,
      returnedOrders: s.returnedOrders, missingOrInvalid: state.raw.filter(r => r.quantity <= 0 || r.unitPrice <= 0).length,
      estimatedRawGrossValue: s.rawGrossValue, sourceName: state.sourceName
    };
  }

  function runClean(cfg) {
    state.lastCfg = Object.assign({}, state.lastCfg, cfg || {});
    const res = DataLens.clean(state.raw, state.lastCfg);
    state.cleaned = res.cleaned;
    state.logs = res.logs;
    state.stats = res.stats;
    state.isVerified = false;
    state.isEvaluated = false;

    const ex = [];
    const seen = new Set();
    state.raw.forEach((r, idx) => {
      if (r.quantity <= 0 || r.unitPrice <= 0 || !r.orderCode || r.productName === 'Không rõ tên') {
        ex.push({ rawIndex: idx, exceptionType: 'MISSING_OR_INVALID', reason: 'Thiếu dữ liệu hoặc số lượng/giá <= 0', actionTaken: 'Tạm giữ kiểm tra', record: r });
        return;
      }
      const k = `${r.orderCode}__${r.sku}__${r.productName}`.toUpperCase();
      if (state.lastCfg.removeDuplicates && seen.has(k)) {
        ex.push({ rawIndex: idx, exceptionType: 'DUPLICATE', reason: `Trùng lặp (${r.orderCode})`, actionTaken: 'Đã loại bỏ', record: r });
        return;
      }
      seen.add(k);
      if (state.lastCfg.excludeCancelled && (r.status === 'CANCELLED' || r.status === 'FAILED')) {
        ex.push({ rawIndex: idx, exceptionType: 'CANCELLED', reason: `Trạng thái ${r.status}`, actionTaken: 'Loại khỏi DT thuần', record: r });
      } else if (r.status === 'RETURNED') {
        ex.push({ rawIndex: idx, exceptionType: 'RETURNED', reason: 'Đơn hoàn trả', actionTaken: 'Giảm trừ DT & hoàn vốn', record: r });
      }
    });
    state.exceptions = ex;

    return {
      success: true,
      rawCount: res.stats.rawCount,
      cleanedCount: res.stats.cleanedCount,
      duplicatesRemoved: res.stats.duplicates,
      cancelledExcluded: res.stats.cancelled,
      returnsAdjusted: res.stats.returns,
      missingOrInvalid: ex.filter(x => x.exceptionType === 'MISSING_OR_INVALID').length,
      imputedCount: res.stats.imputed,
      preventedOvercountValue: res.stats.prevented,
      logs: res.logs,
      exceptions: ex,
      sampleCleaned: res.cleaned.slice(0, 100)
    };
  }

  function buildAnalysis() {
    const a = DataLens.analyze(state.cleaned, state.stats, {});
    const m = a.metrics;
    const nonRet = state.cleaned.filter(x => x.status !== 'RETURNED');
    const aov = nonRet.length ? Math.round(nonRet.reduce((s, x) => s + x.netRevenue, 0) / nonRet.length) : 0;
    const avgDisc = m.totalGross > 0 ? Math.round((m.totalDiscount / m.totalGross) * 10000) / 100 : 0;
    const discOrders = nonRet.filter(x => x.discountAmount > 0).length;
    const maxDisc = nonRet.reduce((mx, x) => Math.max(mx, x.grossRevenue > 0 ? Math.round((x.discountAmount / x.grossRevenue) * 10000) / 100 : 0), 0);
    const today = new Date().toLocaleDateString('vi-VN');

    const structuredRecommendations = [
      { category: 'Giá bán và chiết khấu', title: avgDisc > 8 ? 'Tỷ lệ chiết khấu bình quân vượt ngưỡng 8%' : 'Cơ cấu giá bán và chiết khấu ổn định', content: `Chiết khấu TB ${avgDisc}% (${DataLens.fmt(m.totalDiscount)} đ trên ${discOrders} đơn). Giá trị đơn TB: ${DataLens.fmt(aov)} đ.`, detectedDate: today, priority: avgDisc > 8 ? 'Cao' : 'Trung bình' },
      { category: 'Tăng trưởng doanh thu', title: 'Tổng hợp doanh thu thuần kỳ hiện tại', content: `Doanh thu thuần đạt ${DataLens.fmt(m.totalNet)} đ trên ${m.orders} giao dịch hợp lệ.`, detectedDate: today, priority: 'Trung bình' },
      { category: 'Hiệu quả sản phẩm', title: a.products[0] ? `Sản phẩm dẫn đầu: ${a.products[0].productName}` : 'Chưa có sản phẩm', content: a.products[0] ? `Đóng góp ${DataLens.fmt(a.products[0].netRevenue)} đ (biên ${a.products[0].marginPercent}%).` : '—', detectedDate: today, priority: 'Thấp' },
      { category: 'Hiệu quả chi nhánh', title: a.stores[0] ? `Chi nhánh dẫn đầu: ${a.stores[0].storeName}` : 'Chưa có chi nhánh', content: a.stores[0] ? `Đạt ${DataLens.fmt(a.stores[0].netRevenue)} đ (biên ${a.stores[0].marginPercent}%).` : '—', detectedDate: today, priority: a.stores.some(s => s.marginPercent < 12) ? 'Cao' : 'Thấp' },
      { category: 'Quản trị lợi nhuận', title: m.margin < 15 ? 'Biên lợi nhuận gộp cần tối ưu' : 'Biên lợi nhuận gộp đạt ngưỡng an toàn', content: `Lợi nhuận gộp đạt ${DataLens.fmt(m.totalProfit)} đ (biên ${m.margin}%). Ngăn chặn ${DataLens.fmt(state.stats.prevented)} đ sai lệch.`, detectedDate: today, priority: m.margin < 15 ? 'Cao' : 'Trung bình' }
    ];

    return {
      executiveSummary: a.summary,
      overallMetrics: {
        totalGrossRevenue: m.totalGross, totalDiscountAmount: m.totalDiscount, totalNetRevenue: m.totalNet,
        totalCost: m.totalCost, totalGrossProfit: m.totalProfit, grossMarginPercent: m.margin,
        totalOrders: m.orders, returnOrders: m.returnOrders, returnRatePercent: m.returnRate,
        cancelRatePercent: m.cancelRate, preventedOvercountValue: state.stats.prevented,
        isVerified: state.isVerified, isEvaluated: state.isEvaluated
      },
      pricingAnalysis: { averageOrderValue: aov, averageDiscountRate: avgDisc, maxDiscountPercent: maxDisc, discountedOrderCount: discOrders, totalDiscountAmount: m.totalDiscount },
      storePerformance: a.stores.map(s => ({ storeName: s.storeName, orderCount: s.orders, netRevenue: s.netRevenue, grossProfit: s.grossProfit, marginPercent: s.marginPercent, returnRatePercent: s.returnRate, statusAssessment: s.status })),
      productPerformance: a.products.map(p => ({ sku: p.sku, productName: p.productName, quantitySold: p.quantity, netRevenue: p.netRevenue, grossProfit: p.grossProfit, marginPercent: p.marginPercent })),
      channelBreakdown: a.channels.map(c => ({ channelName: c.channel, orderCount: c.orders, netRevenue: c.netRevenue, revenueSharePercent: c.share })),
      timeSeries: a.trend.map(t => ({ date: t.date, label: t.date.slice(5).split('-').reverse().join('/'), netRevenue: t.netRevenue })),
      growthRatePercent: a.trend.length >= 2 && a.trend[0].netRevenue > 0 ? Math.round(((a.trend[a.trend.length - 1].netRevenue - a.trend[0].netRevenue) / a.trend[0].netRevenue) * 10000) / 100 : 0,
      structuredRecommendations,
      strategicRecommendations: a.recommendations
    };
  }

  // Client-side 6-sheet Excel BI Workbook using SheetJS
  function buildExcelBlob() {
    const a = buildAnalysis();
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.raw.map(r => ({
      'Mã đơn': r.orderCode, 'Ngày': r.orderDate ? new Date(r.orderDate).toLocaleDateString('vi-VN') : '',
      'Khách hàng': r.customerName, 'Chi nhánh': r.storeName, 'SKU': r.sku, 'Sản phẩm': r.productName,
      'SL': r.quantity, 'Đơn giá': r.unitPrice, 'Chiết khấu': r.discountAmount, 'Giá vốn ĐV': r.unitCost,
      'Trạng thái': r.status, 'Kênh': r.channel
    }))), 'RAW DATA');

    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.cleaned.map(r => ({
      'Mã đơn': r.orderCode, 'Ngày': r.orderDate ? new Date(r.orderDate).toLocaleDateString('vi-VN') : '',
      'Chi nhánh': r.storeName, 'SKU': r.sku, 'Sản phẩm': r.productName, 'SL': r.quantity,
      'Đơn giá': r.unitPrice, 'Chiết khấu': r.discountAmount, 'DT Thuần': Math.round(r.netRevenue),
      'Tổng Giá Vốn': Math.round(r.totalCost), 'LN Gộp': Math.round(r.grossProfit), 'Biên LN %': r.marginPercent,
      'Trạng thái': r.status, 'Kênh': r.channel, 'Ghi chú': r.remark
    }))), 'CLEANED DATA');

    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(DataLens.COLUMN_RULES.map(c => ({
      'Tên trường': c.label, 'Mã cột': c.key, 'Bắt buộc': c.required ? 'Có' : 'Không'
    }))), 'DATA DICTIONARY');

    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.logs.map((l, i) => ({ 'STT': i + 1, 'Nhật ký làm sạch': l }))), 'CLEANING LOG');

    const m = a.overallMetrics;
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([
      { 'Chỉ số': 'Doanh thu thuần', 'Giá trị': Math.round(m.totalNetRevenue) },
      { 'Chỉ số': 'Tổng giá vốn', 'Giá trị': Math.round(m.totalCost) },
      { 'Chỉ số': 'Lợi nhuận gộp', 'Giá trị': Math.round(m.totalGrossProfit) },
      { 'Chỉ số': 'Biên lợi nhuận (%)', 'Giá trị': m.grossMarginPercent },
      { 'Chỉ số': 'Số đơn hợp lệ', 'Giá trị': m.totalOrders },
      { 'Chỉ số': 'Tỷ lệ hoàn (%)', 'Giá trị': m.returnRatePercent }
    ]), 'ANALYSIS');

    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(
      a.structuredRecommendations.map(r => ({ 'Nhóm': r.category, 'Ưu tiên': r.priority, 'Tiêu đề': r.title, 'Nội dung': r.content }))
    ), 'DASHBOARD');

    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    return new Response(out, { status: 200, headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' } });
  }

  // Client-side Word BI Report (.docx) using JSZip
  async function buildWordBlob() {
    const a = buildAnalysis();
    const m = a.overallMetrics;
    const escXml = s => String(s ?? '').replace(/[<>&'"]/g, c => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', "'": '&apos;', '"': '&quot;' }[c]));
    const p = (t, bold = false, sz = 26) => `<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>${bold ? '<w:b/>' : ''}<w:sz w:val="${sz}"/></w:rPr><w:t xml:space="preserve">${escXml(t)}</w:t></w:r></w:p>`;

    const body = [
      p('NGỌC LINH DATA STUDIO — BÁO CÁO PHÂN TÍCH BI', true, 32),
      p(`Nguồn dữ liệu: ${state.sourceName} | Ngày lập: ${new Date().toLocaleDateString('vi-VN')}`, false, 24),
      p('PHẦN 1. TÓM TẮT ĐIỀU HÀNH (EXECUTIVE SUMMARY)', true, 28),
      p(a.executiveSummary),
      p(`Doanh thu thuần: ${DataLens.fmt(m.totalNetRevenue)} đ | Lợi nhuận gộp: ${DataLens.fmt(m.totalGrossProfit)} đ | Biên LN: ${m.grossMarginPercent}%`),
      p('PHẦN 2. KẾT QUẢ THEO CHI NHÁNH', true, 28),
      ...a.storePerformance.map(s => p(`• ${s.storeName}: DT ${DataLens.fmt(s.netRevenue)} đ, LN ${DataLens.fmt(s.grossProfit)} đ (${s.marginPercent}%) — ${s.statusAssessment}`)),
      p('PHẦN 3. KẾT QUẢ THEO SẢN PHẨM', true, 28),
      ...a.productPerformance.slice(0, 10).map(pr => p(`• ${pr.productName} (${pr.sku}): SL ${pr.quantitySold}, DT ${DataLens.fmt(pr.netRevenue)} đ, Biên ${pr.marginPercent}%`)),
      p('PHẦN 4. BUSINESS INSIGHTS & KHUYẾN NGHỊ', true, 28),
      ...a.structuredRecommendations.map(r => p(`• [${r.category} · Ưu tiên ${r.priority}] ${r.title}: ${r.content}`))
    ].join('');

    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;
    const ctXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

    const zip = new JSZip();
    zip.file('[Content_Types].xml', ctXml);
    zip.folder('_rels').file('.rels', relsXml);
    zip.folder('word').file('document.xml', docXml);
    const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    return new Response(blob, { status: 200 });
  }

  async function handleStaticApi(url, init) {
    const method = (init && init.method ? init.method : 'GET').toUpperCase();
    const path = url.split('?')[0];

    if (path === '/api/status') {
      return jsonRes({
        step1_connected: state.raw.length > 0, step1_count: state.raw.length, step1_fileCount: state.files.length, step1_source: state.sourceName,
        step2_cleaned: state.cleaned.length > 0, step2_count: state.cleaned.length, step2_exceptions: state.exceptions.length,
        step2_verified: state.isVerified, step3_analyzed: state.cleaned.length > 0 && state.isVerified,
        step3_evaluated: state.isEvaluated, step4_ready: state.cleaned.length > 0 && state.isVerified
      });
    }
    if (path === '/api/reset' && method === 'POST') {
      state.files = []; rebuildRaw();
      return jsonRes({ message: 'Đã xoá dữ liệu. Hệ thống sẵn sàng nhận tệp mới.' });
    }
    if (path === '/api/connect/template') {
      return new Response(DataLens.templateCSV(), { status: 200, headers: { 'Content-Type': 'text/csv; charset=utf-8' } });
    }
    if (path === '/api/connect/preview') {
      return jsonRes({ files: fileSummaries(), summary: rawSummary(), records: state.raw.slice(0, 50) });
    }
    if (path === '/api/connect/upload-json' && method === 'POST') {
      const body = JSON.parse(init.body);
      const recs = (body.records || []).map(r => Object.assign({}, r, { orderDate: r.orderDate ? new Date(r.orderDate) : new Date() }));
      const info = addOrReplaceFile(body.fileName || 'Uploaded.xlsx', recs, body.headers, body.replaceFileId);
      return jsonRes({ message: `Đã kết nối thành công ${recs.length} bản ghi từ '${info.fileName}'!`, file: info, files: fileSummaries(), summary: rawSummary(), preview: state.raw.slice(0, 50) });
    }
    if (path === '/api/connect/upload-file' && method === 'POST') {
      const fd = init.body;
      const file = fd.get('file');
      const replaceId = fd.get('replaceFileId');
      const parsed = await DataLens.parseFile(file);
      const info = addOrReplaceFile(file.name, parsed.records, parsed.headers, replaceId);
      return jsonRes({ message: `Đã kết nối thành công ${info.recordCount} bản ghi từ tệp '${file.name}'!`, file: info, files: fileSummaries(), summary: rawSummary(), preview: state.raw.slice(0, 50) });
    }
    if (path.startsWith('/api/connect/files/') && method === 'DELETE') {
      const id = path.split('/').pop();
      state.files = state.files.filter(f => f.id !== id);
      rebuildRaw();
      return jsonRes({ message: 'Đã xóa tệp dữ liệu.', files: fileSummaries(), summary: rawSummary(), preview: state.raw.slice(0, 50) });
    }
    if (path === '/api/clean/run' && method === 'POST') {
      if (!state.raw.length) return jsonRes({ message: 'Chưa kết nối dữ liệu ở Bước 1.' }, 400);
      return jsonRes(runClean(init.body ? JSON.parse(init.body) : {}));
    }
    if (path.startsWith('/api/clean/raw-record/')) {
      const idx = parseInt(path.split('/').pop(), 10);
      if (method === 'PUT') Object.assign(state.raw[idx], JSON.parse(init.body));
      else if (method === 'DELETE') state.raw.splice(idx, 1);
      return jsonRes({ message: 'Đã cập nhật bản ghi và làm sạch lại.', cleanResult: runClean(state.lastCfg) });
    }
    if (path === '/api/clean/verify' && method === 'POST') {
      const b = JSON.parse(init.body);
      state.isVerified = !!b.approved;
      return jsonRes({ approved: state.isVerified, message: state.isVerified ? 'Phê duyệt thành công!' : 'Đã từ chối dữ liệu.' });
    }
    if (path === '/api/analyze/dashboard') {
      if (!state.cleaned.length) return jsonRes({ message: 'Vui lòng làm sạch dữ liệu trước.' }, 400);
      return jsonRes(buildAnalysis());
    }
    if (path === '/api/analyze/evaluate' && method === 'POST') {
      const b = JSON.parse(init.body);
      state.isEvaluated = !!b.approved;
      return jsonRes({ evaluated: state.isEvaluated, message: 'Đã đánh giá kết quả phân tích đạt chuẩn.' });
    }
    if (path === '/api/analyze/ask-copilot' && method === 'POST') {
      const q = JSON.parse(init.body).question || '';
      const a = DataLens.analyze(state.cleaned, state.stats, {});
      return jsonRes({ question: q, answer: DataLens.answer(q, a) });
    }
    if (path === '/api/export/preview') {
      const a = buildAnalysis();
      return jsonRes({
        generatedAt: new Date().toLocaleString('vi-VN'), sourceName: state.sourceName,
        isVerified: state.isVerified, isEvaluated: state.isEvaluated,
        excel: { fileName: 'Ngoc_Linh_Data_Studio_BI_Workbook.xlsx', sheetCount: 6, cleanRecords: state.cleaned.length, sheets: [
          { name: 'RAW DATA', desc: `${state.raw.length} bản ghi gốc` }, { name: 'CLEANED DATA', desc: `${state.cleaned.length} bản ghi sạch` },
          { name: 'DATA DICTIONARY', desc: 'Từ điển dữ liệu' }, { name: 'CLEANING LOG', desc: `${state.logs.length} dòng nhật ký` },
          { name: 'ANALYSIS', desc: 'Tổng hợp KPI' }, { name: 'DASHBOARD', desc: 'KPI & Khuyến nghị' }
        ]},
        word: { fileName: 'Ngoc_Linh_Data_Studio_BI_Report.docx', sectionCount: 4, sections: [
          'Phần 1. Tóm tắt điều hành', 'Phần 2. Kết quả theo Chi nhánh', 'Phần 3. Kết quả theo Sản phẩm', 'Phần 4. Business Insights & Khuyến nghị'
        ]},
        analysis: a
      });
    }
    if (path === '/api/export/excel') return buildExcelBlob();
    if (path === '/api/export/word') return buildWordBlob();
    if (path === '/api/export/csv') return new Response(DataLens.toCSV(state.cleaned), { status: 200, headers: { 'Content-Type': 'text/csv; charset=utf-8' } });
    if (path === '/api/export/cleaned-records') return jsonRes(state.cleaned);
    if (path === '/api/export/summary-report') {
      const a = buildAnalysis();
      return jsonRes({ generatedDate: new Date().toLocaleString('vi-VN'), analystAgency: 'Ngọc Linh Data Studio', sourceName: state.sourceName, executiveSummary: a.executiveSummary, financialMetrics: a.overallMetrics, structuredRecommendations: a.structuredRecommendations });
    }
    return jsonRes({ message: 'Endpoint not found' }, 404);
  }

  global.fetch = async function (input, init) {
    const url = typeof input === 'string' ? input : input.url;
    if (!url.startsWith('/api/')) return origFetch(input, init);
    if (staticMode === true) return handleStaticApi(url, init);
    try {
      const res = await origFetch(input, init);
      if (res.status === 404 || res.status === 405) {
        staticMode = true;
        global.__VERCEL_STATIC__ = true;
        return handleStaticApi(url, init);
      }
      staticMode = false;
      return res;
    } catch {
      staticMode = true;
      global.__VERCEL_STATIC__ = true;
      return handleStaticApi(url, init);
    }
  };
})(window);

