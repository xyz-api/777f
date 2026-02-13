// Markdown 对比器

let oldText = '';
let newText = '';
let oldFileName = '';
let newFileName = '';
let lastDiff = [];

document.addEventListener('DOMContentLoaded', function () {
    setupUpload('oldUpload', 'oldInput', function (name, text) {
        oldFileName = name; oldText = text;
        markUploaded('oldUpload', name, text);
        tryDiff();
    });
    setupUpload('newUpload', 'newInput', function (name, text) {
        newFileName = name; newText = text;
        markUploaded('newUpload', name, text);
        tryDiff();
    });
    document.getElementById('diffOnlyToggle').addEventListener('change', function () {
        if (lastDiff.length > 0) {
            renderSideBySide(lastDiff);
            renderUnified(lastDiff);
        }
    });
});

function setupUpload(areaId, inputId, callback) {
    const area = document.getElementById(areaId);
    const input = document.getElementById(inputId);
    area.onclick = () => input.click();
    area.ondragover = e => { e.preventDefault(); area.classList.add('dragover'); };
    area.ondragleave = () => area.classList.remove('dragover');
    area.ondrop = e => {
        e.preventDefault(); area.classList.remove('dragover');
        if (e.dataTransfer.files[0]) readFile(e.dataTransfer.files[0], callback);
    };
    input.onchange = e => { if (e.target.files[0]) readFile(e.target.files[0], callback); input.value = ''; };
}

function readFile(file, callback) {
    const reader = new FileReader();
    reader.onload = () => callback(file.name, reader.result);
    reader.readAsText(file);
}

function markUploaded(areaId, name, text) {
    const area = document.getElementById(areaId);
    area.classList.add('has-file');
    const lines = text.split('\n').length;
    area.querySelector('p').textContent = name;
    area.querySelector('small').textContent = lines + ' 行';
}

function tryDiff() {
    if (!oldText && !newText) return;
    // Allow diffing even if one side is empty
    const oldLines = oldText ? oldText.split('\n') : [];
    const newLines = newText ? newText.split('\n') : [];
    const diff = computeDiff(oldLines, newLines);
    lastDiff = diff;

    renderSideBySide(diff);
    renderUnified(diff);
    renderPreview();
    updateStats(diff);
    document.getElementById('resultSection').classList.remove('hidden');
}

// ==================== LCS Diff ====================

function computeDiff(oldLines, newLines) {
    // Myers-like LCS diff
    const N = oldLines.length;
    const M = newLines.length;

    // Build LCS table
    const dp = [];
    for (let i = 0; i <= N; i++) {
        dp[i] = new Array(M + 1).fill(0);
    }
    for (let i = 1; i <= N; i++) {
        for (let j = 1; j <= M; j++) {
            if (oldLines[i - 1] === newLines[j - 1]) {
                dp[i][j] = dp[i - 1][j - 1] + 1;
            } else {
                dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
            }
        }
    }

    // Backtrack to produce diff entries
    const result = [];
    let i = N, j = M;
    while (i > 0 || j > 0) {
        if (i > 0 && j > 0 && oldLines[i - 1] === newLines[j - 1]) {
            result.push({ type: 'equal', oldIdx: i, newIdx: j, oldLine: oldLines[i - 1], newLine: newLines[j - 1] });
            i--; j--;
        } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
            result.push({ type: 'add', newIdx: j, newLine: newLines[j - 1] });
            j--;
        } else {
            result.push({ type: 'remove', oldIdx: i, oldLine: oldLines[i - 1] });
            i--;
        }
    }
    result.reverse();

    // Merge adjacent remove+add into 'change' pairs
    const merged = [];
    let k = 0;
    while (k < result.length) {
        if (k + 1 < result.length && result[k].type === 'remove' && result[k + 1].type === 'add') {
            merged.push({
                type: 'change',
                oldIdx: result[k].oldIdx,
                newIdx: result[k + 1].newIdx,
                oldLine: result[k].oldLine,
                newLine: result[k + 1].newLine
            });
            k += 2;
        } else {
            merged.push(result[k]);
            k++;
        }
    }
    return merged;
}

// ==================== Char-level diff for changed lines ====================

function charDiff(oldStr, newStr) {
    // Simple word-level diff for better readability
    const oldChars = oldStr.split('');
    const newChars = newStr.split('');
    const N = oldChars.length;
    const M = newChars.length;

    // For very long lines, skip char diff
    if (N * M > 500000) return { oldHtml: esc(oldStr), newHtml: esc(newStr) };

    const dp = [];
    for (let i = 0; i <= N; i++) dp[i] = new Array(M + 1).fill(0);
    for (let i = 1; i <= N; i++) {
        for (let j = 1; j <= M; j++) {
            if (oldChars[i - 1] === newChars[j - 1]) dp[i][j] = dp[i - 1][j - 1] + 1;
            else dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
        }
    }

    // Backtrack
    const ops = [];
    let i = N, j = M;
    while (i > 0 || j > 0) {
        if (i > 0 && j > 0 && oldChars[i - 1] === newChars[j - 1]) {
            ops.push({ type: 'eq', oldChar: oldChars[i - 1], newChar: newChars[j - 1] });
            i--; j--;
        } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
            ops.push({ type: 'add', newChar: newChars[j - 1] });
            j--;
        } else {
            ops.push({ type: 'del', oldChar: oldChars[i - 1] });
            i--;
        }
    }
    ops.reverse();

    let oldHtml = '', newHtml = '';
    for (const op of ops) {
        if (op.type === 'eq') {
            oldHtml += esc(op.oldChar);
            newHtml += esc(op.newChar);
        } else if (op.type === 'del') {
            oldHtml += '<span class="char-del">' + esc(op.oldChar) + '</span>';
        } else {
            newHtml += '<span class="char-add">' + esc(op.newChar) + '</span>';
        }
    }
    return { oldHtml, newHtml };
}

// ==================== Filter: diff only with context ====================

function isDiffOnly() {
    return document.getElementById('diffOnlyToggle').checked;
}

function getVisibleIndices(diff, contextLines) {
    // Returns a Set of indices that should be visible (diff lines + surrounding context)
    const visible = new Set();
    for (let i = 0; i < diff.length; i++) {
        if (diff[i].type !== 'equal') {
            for (let j = Math.max(0, i - contextLines); j <= Math.min(diff.length - 1, i + contextLines); j++) {
                visible.add(j);
            }
        }
    }
    return visible;
}

// ==================== Render: Side-by-Side ====================

function renderSideBySide(diff) {
    const leftLines = [];
    const rightLines = [];
    const diffOnly = isDiffOnly();
    const visible = diffOnly ? getVisibleIndices(diff, 2) : null;
    let lastVisible = -1;

    for (let i = 0; i < diff.length; i++) {
        if (diffOnly && !visible.has(i)) continue;

        // Insert separator if there's a gap
        if (diffOnly && lastVisible >= 0 && i - lastVisible > 1) {
            leftLines.push(makeLine('', '<span class="text-muted">···</span>', 'empty-placeholder'));
            rightLines.push(makeLine('', '<span class="text-muted">···</span>', 'empty-placeholder'));
        }
        lastVisible = i;

        const d = diff[i];
        if (d.type === 'equal') {
            leftLines.push(makeLine(d.oldIdx, esc(d.oldLine), ''));
            rightLines.push(makeLine(d.newIdx, esc(d.newLine), ''));
        } else if (d.type === 'remove') {
            leftLines.push(makeLine(d.oldIdx, esc(d.oldLine), 'removed'));
            rightLines.push(makeLine('', '', 'empty-placeholder'));
        } else if (d.type === 'add') {
            leftLines.push(makeLine('', '', 'empty-placeholder'));
            rightLines.push(makeLine(d.newIdx, esc(d.newLine), 'added'));
        } else if (d.type === 'change') {
            const cd = charDiff(d.oldLine, d.newLine);
            leftLines.push(makeLine(d.oldIdx, cd.oldHtml, 'changed'));
            rightLines.push(makeLine(d.newIdx, cd.newHtml, 'changed'));
        }
    }

    const el = document.getElementById('sideBySide');
    el.innerHTML =
        '<div class="diff-pane">' + leftLines.join('') + '</div>' +
        '<div class="diff-pane">' + rightLines.join('') + '</div>';
}

function makeLine(num, contentHtml, cls) {
    return '<div class="diff-line ' + cls + '">' +
        '<span class="line-num">' + (num || '') + '</span>' +
        '<span class="line-content">' + (contentHtml || '&nbsp;') + '</span>' +
        '</div>';
}

// ==================== Render: Unified ====================

function renderUnified(diff) {
    const lines = [];
    const diffOnly = isDiffOnly();
    const visible = diffOnly ? getVisibleIndices(diff, 2) : null;
    let lastVisible = -1;

    for (let i = 0; i < diff.length; i++) {
        if (diffOnly && !visible.has(i)) continue;

        if (diffOnly && lastVisible >= 0 && i - lastVisible > 1) {
            lines.push(makeUnifiedLine('', '', '', '<span class="text-muted">···</span>', 'empty-placeholder'));
        }
        lastVisible = i;

        const d = diff[i];
        if (d.type === 'equal') {
            lines.push(makeUnifiedLine(d.oldIdx, d.newIdx, ' ', esc(d.oldLine), ''));
        } else if (d.type === 'remove') {
            lines.push(makeUnifiedLine(d.oldIdx, '', '-', esc(d.oldLine), 'removed'));
        } else if (d.type === 'add') {
            lines.push(makeUnifiedLine('', d.newIdx, '+', esc(d.newLine), 'added'));
        } else if (d.type === 'change') {
            const cd = charDiff(d.oldLine, d.newLine);
            lines.push(makeUnifiedLine(d.oldIdx, '', '-', cd.oldHtml, 'removed'));
            lines.push(makeUnifiedLine('', d.newIdx, '+', cd.newHtml, 'added'));
        }
    }
    document.getElementById('unifiedView').innerHTML = lines.join('');
}

function makeUnifiedLine(oldNum, newNum, marker, contentHtml, cls) {
    return '<div class="diff-line ' + (cls || '') + '">' +
        '<span class="line-num">' + (oldNum || '') + '</span>' +
        '<span class="line-num">' + (newNum || '') + '</span>' +
        '<span class="diff-marker">' + marker + '</span>' +
        '<span class="line-content">' + (contentHtml || '&nbsp;') + '</span>' +
        '</div>';
}

// ==================== Render: Preview ====================

function renderPreview() {
    document.getElementById('previewOld').innerHTML = marked.parse(oldText || '');
    document.getElementById('previewNew').innerHTML = marked.parse(newText || '');
}

// ==================== Stats ====================

function updateStats(diff) {
    let added = 0, removed = 0, changed = 0;
    for (const d of diff) {
        if (d.type === 'add') added++;
        else if (d.type === 'remove') removed++;
        else if (d.type === 'change') changed++;
    }
    const el = document.getElementById('diffStats');
    el.innerHTML =
        '<span class="added-count">+' + (added + changed) + '</span> ' +
        '<span class="removed-count">-' + (removed + changed) + '</span> ' +
        '<span class="text-muted">' + changed + ' 处修改</span>';
}

// ==================== Util ====================

function esc(str) {
    if (!str) return '';
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
