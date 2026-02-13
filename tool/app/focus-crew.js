let scheduleWorkbook = null;
let focusWorkbook = null;
let scheduleData = [];
let focusSheets = [];
let scheduleColumns = [];
let resultWorkbook = null;

const CATEGORY_CONFIG = {
    '重点关注': { priority: 1, color: 'FFE5CC', label: '重点' },
    '一般关注': { priority: 2, color: 'FFF3CD', label: '一般' },
    '预防性关注': { priority: 3, color: '5b84f9', label: '预防' },
    '三新人员（不上会）': { priority: 4, color: '65c53f', label: '三新' },
    '长期关注': { priority: 5, color: 'a584ed', label: '长期' }
};

document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('scheduleFile').addEventListener('change', handleScheduleFile);
    document.getElementById('focusFile').addEventListener('change', handleFocusFile);
    document.getElementById('highlightBtn').addEventListener('click', doHighlight);
    document.getElementById('exportBtn').addEventListener('click', exportExcel);
});

function handleScheduleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            scheduleWorkbook = XLSX.read(data, {type: 'array'});
            showStatus('scheduleStatus', '已加载: ' + file.name, 'success');
            loadSchedulePreview();
        } catch (err) {
            showStatus('scheduleStatus', '文件解析失败: ' + err.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

function handleFocusFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            focusWorkbook = XLSX.read(data, {type: 'array'});
            showStatus('focusStatus', '已加载: ' + file.name, 'success');
            loadFocusPreview();
        } catch (err) {
            showStatus('focusStatus', '文件解析失败: ' + err.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

function showStatus(id, msg, type) {
    const el = document.getElementById(id);
    el.textContent = msg;
    el.className = 'status status-' + type;
}

function loadSchedulePreview() {
    if (!scheduleWorkbook) return;
    
    const sheet = scheduleWorkbook.Sheets[scheduleWorkbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, {header: 1, raw: false});
    
    if (json.length === 0) {
        showStatus('scheduleStatus', '表格为空', 'error');
        return;
    }
    
    scheduleColumns = (json[0] || []).map((c, i) => c ? String(c) : '列' + (i+1));
    scheduleData = json;
    
    renderPreview('schedule', scheduleColumns, json.slice(0, 6));
    renderSelectors('schedule', scheduleColumns);
    
    document.getElementById('scheduleConfigSection').style.display = 'block';
    checkReady();
}

function loadFocusPreview() {
    if (!focusWorkbook) return;
    
    const sheetNames = focusWorkbook.SheetNames;
    focusSheets = [];
    
    for (const sheetName of sheetNames) {
        let category = null;
        
        for (const cat of Object.keys(CATEGORY_CONFIG)) {
            if (sheetName.includes(cat) || cat.includes(sheetName)) {
                category = cat;
                break;
            }
        }
        
        if (!category) continue;
        
        const sheet = focusWorkbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, {header: 1, raw: false});
        
        if (json.length < 2) continue;
        
        const columns = (json[1] || []).map((c, i) => c ? String(c) : '列' + (i+1));
        
        focusSheets.push({
            name: sheetName,
            category: category,
            columns: columns,
            data: json
        });
    }
    
    if (focusSheets.length === 0) {
        showStatus('focusStatus', '未找到有效的关注类别工作表', 'error');
        return;
    }
    
    renderFocusSheets();
    
    showStatus('focusStatus', '已加载 ' + focusSheets.length + ' 个关注类别工作表', 'success');
    
    document.getElementById('focusConfigSection').style.display = 'block';
    checkReady();
}

function renderFocusSheets() {
    const container = document.getElementById('focusSheetsConfig');
    let html = '';
    
    focusSheets.forEach((sheetInfo, sheetIdx) => {
        const badgeClass = 'badge-' + CATEGORY_CONFIG[sheetInfo.category].priority;
        
        html += '<div class="sheet-config">';
        html += '<div class="sheet-config-title">' + sheetInfo.name;
        html += '<span class="badge ' + badgeClass + '">' + CATEGORY_CONFIG[sheetInfo.category].label + '</span>';
        html += '</div>';
        
        html += '<div class="config-row">';
        
        html += '<div class="config-item"><label>员工号列:</label>';
        html += '<select id="focusIdCol_' + sheetIdx + '">';
        html += '<option value="">请选择...</option>';
        sheetInfo.columns.forEach((col, colIdx) => {
            const selected = (col.includes('员工号') || col === '员工号') ? ' selected' : '';
            html += '<option value="' + colIdx + '"' + selected + '>' + col + '</option>';
        });
        html += '</select></div>';
        
        html += '<div class="config-item"><label>姓名列:</label>';
        html += '<select id="focusNameCol_' + sheetIdx + '">';
        html += '<option value="">请选择...</option>';
        sheetInfo.columns.forEach((col, colIdx) => {
            const selected = (col.includes('姓名') || col === '姓名') ? ' selected' : '';
            html += '<option value="' + colIdx + '"' + selected + '>' + col + '</option>';
        });
        html += '</select></div>';
        
        html += '</div>';
        
        html += '<div class="preview-wrapper"><table class="preview-table">';
        html += '<thead><tr>';
        sheetInfo.columns.forEach((col, i) => { 
            html += '<th title="列索引:' + i + '">' + col + '</th>'; 
        });
        html += '</tr></thead><tbody>';
        
        const previewRows = sheetInfo.data.slice(2, 7);
        previewRows.forEach(row => {
            html += '<tr>';
            sheetInfo.columns.forEach((_, i) => {
                const val = row[i] !== undefined && row[i] !== null ? String(row[i]) : '';
                const display = val.length > 20 ? val.substring(0, 20) + '...' : val;
                html += '<td title="' + val + '">' + display + '</td>';
            });
            html += '</tr>';
        });
        
        html += '</tbody></table></div>';
        html += '</div>';
    });
    
    container.innerHTML = html;
}

function renderPreview(type, columns, rows) {
    const table = document.getElementById(type + 'Preview');
    let html = '<thead><tr>';
    columns.forEach((col, i) => { 
        html += '<th title="列索引:' + i + '">' + col + '</th>'; 
    });
    html += '</tr></thead><tbody>';
    rows.forEach(row => {
        html += '<tr>';
        columns.forEach((_, i) => {
            const val = row[i] !== undefined && row[i] !== null ? String(row[i]) : '';
            const display = val.length > 20 ? val.substring(0, 20) + '...' : val;
            html += '<td title="' + val + '">' + display + '</td>';
        });
        html += '</tr>';
    });
    html += '</tbody>';
    table.innerHTML = html;
}

function renderSelectors(type, columns) {
    const idSelect = document.getElementById(type + 'IdCol');
    const nameSelect = document.getElementById(type + 'NameCol');
    
    idSelect.innerHTML = '<option value="">请选择...</option>';
    nameSelect.innerHTML = '<option value="">请选择...</option>';
    
    columns.forEach((col, i) => {
        const opt = '<option value="' + i + '">' + col + '</option>';
        idSelect.innerHTML += opt;
        nameSelect.innerHTML += opt;
    });
    
    columns.forEach((col, i) => {
        if (col.includes('员工号') || col === '员工号') {
            idSelect.value = i;
        }
        if (col.includes('姓名') || col === '姓名') {
            nameSelect.value = i;
        }
    });
}

function checkReady() {
    if (scheduleData.length > 0 && focusSheets.length > 0) {
        document.getElementById('actionSection').style.display = 'block';
    }
}

function doHighlight() {
    const scheduleNameCol = parseInt(document.getElementById('scheduleNameCol').value);
    
    if (isNaN(scheduleNameCol)) {
        alert('请选择审班表的姓名列');
        return;
    }
    
    const focusData = {};
    
    for (let sheetIdx = 0; sheetIdx < focusSheets.length; sheetIdx++) {
        const sheetInfo = focusSheets[sheetIdx];
        const nameColSelect = document.getElementById('focusNameCol_' + sheetIdx);
        
        if (!nameColSelect) {
            console.warn('未找到工作表 [' + sheetInfo.name + '] 的姓名列选择器');
            continue;
        }
        
        const nameCol = parseInt(nameColSelect.value);
        
        if (isNaN(nameCol)) {
            alert('请选择 [' + sheetInfo.name + '] 的姓名列');
            return;
        }
        
        for (let i = 2; i < sheetInfo.data.length; i++) {
            const row = sheetInfo.data[i];
            if (!row || row.length === 0) continue;
            
            const name = String(row[nameCol] || '').trim();
            if (!name) continue;
            
            if (!focusData[name]) {
                focusData[name] = [];
            }
            focusData[name].push(sheetInfo.category);
        }
    }
    
    const focusNames = Object.keys(focusData);
    console.log('重点人员总数:', focusNames.length);
    console.log('重点人员示例:', focusNames.slice(0, 10));
    
    if (focusNames.length === 0) {
        alert('未找到任何重点人员，请检查重点人员表的姓名列是否选择正确');
        return;
    }
    
    resultWorkbook = XLSX.utils.book_new();
    const matchedCategories = {};
    
    for (const sheetName of scheduleWorkbook.SheetNames) {
        const sourceSheet = scheduleWorkbook.Sheets[sheetName];
        const range = XLSX.utils.decode_range(sourceSheet['!ref'] || 'A1');
        
        const newSheet = {};
        let matchCount = 0;
        
        for (let R = 0; R <= range.e.r; R++) {
            for (let C = 0; C <= range.e.c; C++) {
                const addr = XLSX.utils.encode_cell({r: R, c: C});
                const sourceCell = sourceSheet[addr];
                
                if (!sourceCell) {
                    continue;
                }
                
                let cellValue = sourceCell.v;
                let cellStyle = sourceCell.s || {};
                
                if (R > 0 && C === scheduleNameCol) {
                    const name = String(cellValue || '').trim();
                    
                    if (name && focusData[name]) {
                        const categories = focusData[name];
                        const uniqueCategories = Array.from(new Set(categories));
                        
                        uniqueCategories.sort((a, b) => {
                            return CATEGORY_CONFIG[a].priority - CATEGORY_CONFIG[b].priority;
                        });
                        
                        const topCategory = uniqueCategories[0];
                        const color = CATEGORY_CONFIG[topCategory].color;
                        
                        const labels = uniqueCategories.map(cat => '[' + CATEGORY_CONFIG[cat].label + ']').join('');
                        cellValue = name + labels;
                        
                        cellStyle = {
                            fill: { patternType: "solid", fgColor: { rgb: color } },
                            alignment: { horizontal: "left", vertical: "center" },
                            border: {
                                top: {style: "thin", color: {rgb: "000000"}},
                                bottom: {style: "thin", color: {rgb: "000000"}},
                                left: {style: "thin", color: {rgb: "000000"}},
                                right: {style: "thin", color: {rgb: "000000"}}
                            }
                        };
                        
                        matchCount++;
                        
                        uniqueCategories.forEach(cat => {
                            matchedCategories[cat] = (matchedCategories[cat] || 0) + 1;
                        });
                    }
                }
                
                newSheet[addr] = {
                    v: cellValue,
                    t: sourceCell.t || 's',
                    s: cellStyle
                };
            }
        }
        
        newSheet['!ref'] = sourceSheet['!ref'];
        if (sourceSheet['!cols']) newSheet['!cols'] = sourceSheet['!cols'];
        if (sourceSheet['!rows']) newSheet['!rows'] = sourceSheet['!rows'];
        
        XLSX.utils.book_append_sheet(resultWorkbook, newSheet, sheetName);
        
        console.log('工作表 [' + sheetName + '] 匹配到 ' + matchCount + ' 人');
    }
    
    displayStats(focusNames.length, matchedCategories);
    document.getElementById('exportBtn').disabled = false;
}

function displayStats(totalFocus, matchedCategories) {
    const statsDiv = document.getElementById('resultStats');
    statsDiv.style.display = 'block';
    
    let html = '<div class="stats-item">重点人员总数: <span class="stats-num">' + totalFocus + '</span></div>';
    
    const sortedCategories = Object.keys(matchedCategories).sort((a, b) => {
        return CATEGORY_CONFIG[a].priority - CATEGORY_CONFIG[b].priority;
    });
    
    for (const cat of sortedCategories) {
        const label = CATEGORY_CONFIG[cat].label;
        const count = matchedCategories[cat];
        html += '<div class="stats-item">' + label + ': <span class="stats-num">' + count + '</span></div>';
    }
    
    statsDiv.innerHTML = html;
}

function exportExcel() {
    if (!resultWorkbook) {
        alert('请先进行标注');
        return;
    }
    XLSX.writeFile(resultWorkbook, '审班表_已标注.xlsx');
}
