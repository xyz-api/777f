// Excel 转 Markdown

let parsedData = null; // { fileName, sheets: [{name, rows, cols, headers, data, colMeta}] }
let currentViewMode = 'excel'; // 'excel' or 'markdown'
let currentTab = 'overview';

document.addEventListener('DOMContentLoaded', function() {
    setupUpload();
    document.getElementById('downloadBtn').addEventListener('click', downloadZip);
    document.getElementById('previewBtn').addEventListener('click', togglePreview);
    document.getElementById('viewExcelBtn').addEventListener('click', () => switchViewMode('excel'));
    document.getElementById('viewMdBtn').addEventListener('click', () => switchViewMode('markdown'));
});

function setupUpload() {
    const area = document.getElementById('uploadArea');
    const input = document.getElementById('fileInput');
    area.onclick = () => input.click();
    area.ondragover = e => { e.preventDefault(); area.classList.add('dragover'); };
    area.ondragleave = () => area.classList.remove('dragover');
    area.ondrop = e => { e.preventDefault(); area.classList.remove('dragover'); if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); };
    input.onchange = e => { if (e.target.files[0]) handleFile(e.target.files[0]); input.value = ''; };
}

function showStatus(msg, type) {
    const el = document.getElementById('statusBar');
    el.textContent = msg;
    el.className = 'status-bar' + (type ? ' ' + type : '');
    el.classList.remove('hidden');
}

function handleFile(file) {
    showStatus('解析中...', '');
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellFormula: true, cellStyles: true, cellDates: true });
            parsedData = parseWorkbook(workbook, file.name);
            showStatus('解析完成: ' + file.name + ' (' + parsedData.sheets.length + ' 个工作表)', 'success');
            renderSummary();
            document.getElementById('resultSection').classList.remove('hidden');
            document.getElementById('previewSection').classList.add('hidden');
        } catch (err) {
            showStatus('解析失败: ' + err.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

// ==================== 解析 ====================

function parseWorkbook(workbook, fileName) {
    const sheets = [];

    for (const sheetName of workbook.SheetNames) {
        const ws = workbook.Sheets[sheetName];
        const ref = ws['!ref'];
        if (!ref) { sheets.push({ name: sheetName, rows: 0, cols: 0, headers: [], data: [], colMeta: [] }); continue; }

        const range = XLSX.utils.decode_range(ref);
        const totalRows = range.e.r - range.s.r + 1;
        const totalCols = range.e.c - range.s.c + 1;

        // 提取表头（第一行）
        const headers = [];
        for (let C = range.s.c; C <= range.e.c; C++) {
            const addr = XLSX.utils.encode_cell({ r: range.s.r, c: C });
            const cell = ws[addr];
            headers.push(cell ? formatCellValue(cell) : '列' + (C + 1));
        }

        // 提取所有数据行
        const data = [];
        for (let R = range.s.r + 1; R <= range.e.r; R++) {
            const row = [];
            let hasValue = false;
            for (let C = range.s.c; C <= range.e.c; C++) {
                const addr = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = ws[addr];
                const val = cell ? formatCellValue(cell) : '';
                if (val) hasValue = true;
                row.push(val);
            }
            if (hasValue) data.push(row);
        }

        // 分析每列的元信息
        const colMeta = [];
        for (let C = range.s.c; C <= range.e.c; C++) {
            const colIdx = C - range.s.c;
            let types = {};
            let emptyCount = 0;
            let sample = '';

            for (let R = range.s.r + 1; R <= range.e.r; R++) {
                const addr = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = ws[addr];
                if (!cell || cell.v === undefined || cell.v === null || String(cell.v).trim() === '') {
                    emptyCount++;
                    continue;
                }
                const typeName = detectCellType(cell);
                types[typeName] = (types[typeName] || 0) + 1;
                if (!sample) sample = formatCellValue(cell);
            }

            const dataRows = totalRows - 1;
            const emptyRate = dataRows > 0 ? Math.round(emptyCount / dataRows * 100) : 0;
            const mainType = Object.keys(types).sort((a, b) => types[b] - types[a])[0] || '空';

            colMeta.push({
                header: headers[colIdx] || '列' + (colIdx + 1),
                type: mainType,
                sample: sample.length > 30 ? sample.substring(0, 30) + '...' : sample,
                emptyRate: emptyRate
            });
        }

        sheets.push({ name: sheetName, rows: data.length, cols: totalCols, headers, data, colMeta });
    }

    return { fileName: fileName.replace(/\.[^/.]+$/, ''), sheets };
}

function detectCellType(cell) {
    // 超链接
    if (cell.l && cell.l.Target) return '超链接';
    if (cell.f && cell.f.includes('HYPERLINK')) return '超链接';
    // 日期
    if (cell.t === 'd') return '日期';
    if (cell.t === 'n' && cell.z && /[ymdh]/i.test(cell.z)) return '日期';
    // 数字
    if (cell.t === 'n') return '数字';
    // 布尔
    if (cell.t === 'b') return '布尔';
    // 文本
    return '文本';
}

function formatCellValue(cell) {
    if (!cell || cell.v === undefined || cell.v === null) return '';

    // 超链接：优先检查 cell.l，再检查 HYPERLINK 公式
    if (cell.l && cell.l.Target) {
        const display = String(cell.v || cell.l.Target).trim();
        return '[' + escMd(display) + '](' + cell.l.Target + ')';
    }
    if (cell.f && cell.f.includes('HYPERLINK')) {
        const match = cell.f.match(/HYPERLINK\s*\(\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\)/i);
        if (match) return '[' + escMd(match[2]) + '](' + match[1] + ')';
    }

    // 日期
    if (cell.t === 'd' && cell.v instanceof Date) {
        return formatDate(cell.v);
    }
    if (cell.t === 'n' && cell.z && /[ymdh]/i.test(cell.z)) {
        // Excel 日期序列号
        try {
            const d = XLSX.SSF.parse_date_code(cell.v);
            if (d) return d.y + '-' + String(d.m).padStart(2, '0') + '-' + String(d.d).padStart(2, '0');
        } catch (e) {}
    }

    // 数字
    if (cell.t === 'n') return String(cell.v);

    // 布尔
    if (cell.t === 'b') return cell.v ? '是' : '否';

    // 文本
    const val = String(cell.v).trim();
    return escMd(val);
}

function formatDate(d) {
    if (!(d instanceof Date) || isNaN(d.getTime())) return '';
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

function escMd(text) {
    // 转义 Markdown 表格中的管道符和换行
    return text.replace(/\|/g, '\\|').replace(/\n/g, ' ');
}

// ==================== 生成 Markdown ====================

function generateOverview() {
    if (!parsedData) return '';
    const lines = [];
    lines.push('# 文件概览: ' + parsedData.fileName);
    lines.push('');
    const totalRows = parsedData.sheets.reduce((s, sh) => s + sh.rows, 0);
    lines.push('共 ' + parsedData.sheets.length + ' 个工作表，总计 ' + totalRows + ' 行数据');
    lines.push('');

    parsedData.sheets.forEach((sheet, idx) => {
        lines.push('## Sheet ' + (idx + 1) + ': ' + sheet.name + ' (' + sheet.rows + '行 × ' + sheet.cols + '列)');
        lines.push('');
        if (sheet.colMeta.length === 0) {
            lines.push('（空工作表）');
            lines.push('');
            return;
        }
        lines.push('| 列名 | 数据类型 | 示例值 | 空值率 |');
        lines.push('|------|---------|--------|--------|');
        sheet.colMeta.forEach(col => {
            lines.push('| ' + escMd(col.header) + ' | ' + col.type + ' | ' + (col.sample || '-') + ' | ' + col.emptyRate + '% |');
        });
        lines.push('');
    });

    return lines.join('\n');
}

function generateSheetMarkdown(sheet) {
    if (sheet.rows === 0 || sheet.headers.length === 0) return '# ' + sheet.name + '\n\n（空工作表）\n';

    const lines = [];
    lines.push('# ' + sheet.name);
    lines.push('');
    lines.push(sheet.rows + '行 × ' + sheet.cols + '列');
    lines.push('');

    // 表头
    lines.push('| ' + sheet.headers.map(h => escMd(h)).join(' | ') + ' |');
    lines.push('| ' + sheet.headers.map(() => '---').join(' | ') + ' |');

    // 数据行
    for (const row of sheet.data) {
        const cells = sheet.headers.map((_, i) => row[i] !== undefined ? row[i] : '');
        lines.push('| ' + cells.join(' | ') + ' |');
    }

    lines.push('');
    return lines.join('\n');
}

// ==================== UI ====================

function renderSummary() {
    if (!parsedData) return;
    const el = document.getElementById('sheetSummary');
    let html = '<table class="sheet-summary"><thead><tr><th>#</th><th>工作表</th><th>行数</th><th>列数</th></tr></thead><tbody>';
    parsedData.sheets.forEach((sheet, i) => {
        html += '<tr><td>' + (i + 1) + '</td><td>' + sheet.name + '</td><td>' + sheet.rows + '</td><td>' + sheet.cols + '</td></tr>';
    });
    html += '</tbody></table>';
    el.innerHTML = html;
}

function togglePreview() {
    const section = document.getElementById('previewSection');
    if (section.classList.contains('hidden')) {
        renderPreviewTabs();
        section.classList.remove('hidden');
    } else {
        section.classList.add('hidden');
    }
}

function renderPreviewTabs() {
    if (!parsedData) return;
    const tabsEl = document.getElementById('previewTabs');

    // Build tab list: overview + each sheet
    let tabsHtml = '<li class="nav-item"><a class="nav-link active" href="#" data-tab="overview">数据概览</a></li>';
    parsedData.sheets.forEach((sheet, i) => {
        tabsHtml += '<li class="nav-item"><a class="nav-link" href="#" data-tab="sheet-' + i + '">' + sheet.name + '</a></li>';
    });
    tabsEl.innerHTML = tabsHtml;

    currentTab = 'overview';
    renderCurrentTab();

    // Tab click handler
    tabsEl.querySelectorAll('.nav-link').forEach(link => {
        link.addEventListener('click', function(e) {
            e.preventDefault();
            tabsEl.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
            this.classList.add('active');
            currentTab = this.dataset.tab;
            renderCurrentTab();
        });
    });
}

function switchViewMode(mode) {
    currentViewMode = mode;
    document.getElementById('viewExcelBtn').classList.toggle('active', mode === 'excel');
    document.getElementById('viewMdBtn').classList.toggle('active', mode === 'markdown');
    renderCurrentTab();
}

function renderCurrentTab() {
    const boxEl = document.getElementById('previewBox');
    const isExcel = currentViewMode === 'excel';

    boxEl.className = 'preview-box' + (isExcel ? ' excel-view' : '');

    if (currentTab === 'overview') {
        if (isExcel) {
            boxEl.innerHTML = generateOverviewHtml();
        } else {
            boxEl.textContent = generateOverview();
        }
    } else {
        const idx = parseInt(currentTab.split('-')[1]);
        const sheet = parsedData.sheets[idx];
        if (isExcel) {
            boxEl.innerHTML = generateSheetHtml(sheet);
        } else {
            boxEl.textContent = generateSheetMarkdown(sheet);
        }
    }
}

function generateOverviewHtml() {
    if (!parsedData) return '';
    const totalRows = parsedData.sheets.reduce((s, sh) => s + sh.rows, 0);
    let html = '<div style="padding:12px">';
    html += '<h6>文件概览: ' + esc(parsedData.fileName) + '</h6>';
    html += '<p style="font-size:13px;color:#666">共 ' + parsedData.sheets.length + ' 个工作表，总计 ' + totalRows + ' 行数据</p>';

    parsedData.sheets.forEach((sheet, idx) => {
        html += '<h6 style="margin-top:16px">Sheet ' + (idx + 1) + ': ' + esc(sheet.name) + ' (' + sheet.rows + '行 × ' + sheet.cols + '列)</h6>';
        if (sheet.colMeta.length === 0) {
            html += '<p style="color:#999">（空工作表）</p>';
            return;
        }
        html += '<table><thead><tr><th>列名</th><th>数据类型</th><th>示例值</th><th>空值率</th></tr></thead><tbody>';
        sheet.colMeta.forEach(col => {
            html += '<tr><td>' + esc(col.header) + '</td><td>' + esc(col.type) + '</td><td>' + esc(col.sample || '-') + '</td><td>' + col.emptyRate + '%</td></tr>';
        });
        html += '</tbody></table>';
    });

    html += '</div>';
    return html;
}

function generateSheetHtml(sheet) {
    if (sheet.rows === 0 || sheet.headers.length === 0) {
        return '<div style="padding:12px;color:#999">（空工作表）</div>';
    }
    let html = '<table><thead><tr>';
    sheet.headers.forEach(h => { html += '<th>' + esc(stripMdLink(h)) + '</th>'; });
    html += '</tr></thead><tbody>';
    for (const row of sheet.data) {
        html += '<tr>';
        sheet.headers.forEach((_, i) => {
            const val = row[i] !== undefined ? row[i] : '';
            html += '<td>' + cellToHtml(val) + '</td>';
        });
        html += '</tr>';
    }
    html += '</tbody></table>';
    return html;
}

function cellToHtml(val) {
    // Convert markdown links back to HTML links
    const linkMatch = val.match(/^\[([^\]]*)\]\(([^)]+)\)$/);
    if (linkMatch) {
        return '<a href="' + esc(linkMatch[2]) + '" target="_blank">' + esc(linkMatch[1]) + '</a>';
    }
    return esc(val.replace(/\\\|/g, '|'));
}

function stripMdLink(text) {
    const m = text.match(/^\[([^\]]*)\]\([^)]+\)$/);
    return m ? m[1] : text.replace(/\\\|/g, '|');
}

function esc(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

async function downloadZip() {
    if (!parsedData) return;

    const zip = new JSZip();
    const folderName = parsedData.fileName;
    const folder = zip.folder(folderName);

    // 概览文件
    folder.file('00_数据概览.md', generateOverview());

    // 每个 sheet 一个文件
    parsedData.sheets.forEach((sheet, i) => {
        const num = String(i + 1).padStart(2, '0');
        const safeName = sheet.name.replace(/[\\/:*?"<>|]/g, '_');
        folder.file(num + '_' + safeName + '.md', generateSheetMarkdown(sheet));
    });

    const blob = await zip.generateAsync({ type: 'blob' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = folderName + '_markdown.zip';
    a.click();
    URL.revokeObjectURL(url);
}
