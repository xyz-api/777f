const ROSTER_PATH = '../../template/机组花名册.xlsx';
let employeeData = [];
let matchedResults = [];

// 页面加载时自动加载默认花名册
document.addEventListener('DOMContentLoaded', loadDefaultRoster);

async function loadDefaultRoster() {
    try {
        showFileStatus('正在加载默认花名册...', 'loading');
        const resp = await fetch(ROSTER_PATH);
        if (!resp.ok) throw new Error('文件不存在');
        const buffer = await resp.arrayBuffer();
        parseExcelData(new Uint8Array(buffer), '机组花名册.xlsx');
    } catch (e) {
        console.error('自动加载花名册失败:', e);
        showFileStatus('请选择员工花名册文件', 'hint');
    }
}

// 解析Excel数据（复用逻辑）
function parseExcelData(data, fileName) {
    try {
        const workbook = XLSX.read(data, {type: 'array'});
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, {header: 1});
        
        employeeData = [];
        for (let i = 1; i < json.length; i++) {
            const row = json[i];
            if (row && row[0] && row[1]) {
                employeeData.push({
                    id: String(row[0]).trim(),
                    name: String(row[1]).trim()
                });
            }
        }
        
        showFileStatus('已加载: ' + fileName + '（' + employeeData.length + ' 条数据）', 'success');
        document.getElementById('searchBtn').disabled = false;
    } catch (err) {
        showFileStatus('文件解析失败: ' + err.message, 'error');
    }
}

document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        parseExcelData(new Uint8Array(e.target.result), file.name);
    };
    reader.readAsArrayBuffer(file);
});

function showFileStatus(msg, type) {
    const status = document.getElementById('fileStatus');
    status.textContent = msg;
    status.style.display = 'inline-block';
    status.style.padding = '6px 12px';
    status.style.borderRadius = '6px';
    status.style.fontSize = '13px';
    if (type === 'success') {
        status.style.background = '#dafbe1';
        status.style.color = '#1a7f37';
    } else if (type === 'error') {
        status.style.background = '#ffebe9';
        status.style.color = '#cf222e';
    } else if (type === 'loading') {
        status.style.background = '#ddf4ff';
        status.style.color = '#0969da';
    } else {
        status.style.background = '#f6f8fa';
        status.style.color = '#656d76';
    }
}

document.getElementById('searchBtn').addEventListener('click', function() {
    const text = document.getElementById('textInput').value;
    if (!text.trim()) {
        alert('请输入要查询的文本');
        return;
    }
    
    matchedResults = [];
    employeeData.forEach(emp => {
        const pos = text.indexOf(emp.name);
        if (pos !== -1) {
            matchedResults.push({ ...emp, pos });
        }
    });
    
    // 按文本中出现的位置排序
    matchedResults.sort((a, b) => a.pos - b.pos);
    
    const resultBody = document.getElementById('resultBody');
    const countInfo = document.getElementById('countInfo');
    
    if (matchedResults.length === 0) {
        resultBody.innerHTML = '<tr><td colspan="3" class="no-result">未找到匹配的员工</td></tr>';
        countInfo.textContent = '匹配到 0 个员工';
    } else {
        resultBody.innerHTML = matchedResults.map((emp, idx) => 
            '<tr><td><input type="checkbox" class="row-check" data-idx="' + idx + '"></td><td>' + emp.name + '</td><td>' + emp.id + '</td></tr>'
        ).join('');
        countInfo.textContent = '匹配到 ' + matchedResults.length + ' 个员工';
    }
    document.getElementById('selectAll').checked = false;
});

document.getElementById('clearBtn').addEventListener('click', function() {
    document.getElementById('textInput').value = '';
    document.getElementById('resultBody').innerHTML = '<tr><td colspan="3" class="no-result">匹配结果将显示在这里...</td></tr>';
    document.getElementById('countInfo').textContent = '匹配到 0 个员工';
    matchedResults = [];
    document.getElementById('selectAll').checked = false;
    document.getElementById('textInput').focus();
});

document.getElementById('copyNameBtn').addEventListener('click', function() {
    if (matchedResults.length === 0) {
        alert('没有可复制的数据，请先查询匹配。');
        return;
    }
    const selected = getSelectedResults();
    const data = selected.length > 0 ? selected : matchedResults;
    const names = data.map(emp => emp.name).join('\n');
    copyToClipboard(names, this, '复制姓名列');
});

document.getElementById('copyIdBtn').addEventListener('click', function() {
    if (matchedResults.length === 0) {
        alert('没有可复制的数据，请先查询匹配。');
        return;
    }
    const selected = getSelectedResults();
    const data = selected.length > 0 ? selected : matchedResults;
    const ids = data.map(emp => emp.id).join('\n');
    copyToClipboard(ids, this, '复制员工号列');
});

function getSelectedResults() {
    const checks = document.querySelectorAll('.row-check:checked');
    return Array.from(checks).map(cb => matchedResults[parseInt(cb.dataset.idx)]);
}

// 全选功能
document.getElementById('selectAll').addEventListener('change', function() {
    const checks = document.querySelectorAll('.row-check');
    checks.forEach(cb => {
        cb.checked = this.checked;
        cb.closest('tr').classList.toggle('selected', this.checked);
    });
});

// 单行选中高亮
document.getElementById('resultBody').addEventListener('change', function(e) {
    if (e.target.classList.contains('row-check')) {
        e.target.closest('tr').classList.toggle('selected', e.target.checked);
        // 更新全选状态
        const checks = document.querySelectorAll('.row-check');
        const allChecked = checks.length > 0 && Array.from(checks).every(cb => cb.checked);
        document.getElementById('selectAll').checked = allChecked;
    }
});

function copyToClipboard(text, btn, originalText) {
    navigator.clipboard.writeText(text).then(() => {
        btn.textContent = '已复制！';
        btn.style.backgroundColor = '#1a7f37';
        setTimeout(() => {
            btn.textContent = originalText;
            btn.style.backgroundColor = '';
        }, 2000);
    }).catch(err => {
        alert('复制失败，请手动选择文本复制。');
    });
}
