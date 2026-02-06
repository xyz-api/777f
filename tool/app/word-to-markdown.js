// Word 转 Markdown

let currentMarkdown = '';
let currentHtml = '';
let currentFileName = '';

document.addEventListener('DOMContentLoaded', function() {
    setupUpload();
    document.getElementById('downloadBtn').addEventListener('click', downloadMd);
    document.getElementById('copyBtn').addEventListener('click', copyMd);
    document.getElementById('includeImages').addEventListener('change', reconvert);
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

let currentArrayBuffer = null;

function showStatus(msg, type) {
    const el = document.getElementById('statusBar');
    el.textContent = msg;
    el.className = 'status-bar ' + type;
    el.classList.remove('hidden');
}

async function handleFile(file) {
    currentFileName = file.name.replace(/\.[^/.]+$/, '');
    showStatus('转换中...', 'info');
    try {
        currentArrayBuffer = await file.arrayBuffer();
        await convert();
    } catch (err) {
        showStatus('转换失败: ' + err.message, 'error');
    }
}

async function reconvert() {
    if (!currentArrayBuffer) return;
    await convert();
}

async function convert() {
    const includeImages = document.getElementById('includeImages').checked;

    const options = {
        convertImage: includeImages
            ? mammoth.images.imgElement(function(image) {
                return image.read('base64').then(function(imageBuffer) {
                    return { src: 'data:' + image.contentType + ';base64,' + imageBuffer };
                });
            })
            : { convertImage: function() { return Promise.resolve({}); } }
    };

    // mammoth 转 HTML
    const result = await mammoth.convertToHtml({ arrayBuffer: currentArrayBuffer }, includeImages ? options : {});
    currentHtml = result.value;

    // HTML 转 Markdown
    currentMarkdown = htmlToMarkdown(currentHtml);

    // 显示结果
    document.getElementById('markdownOutput').textContent = currentMarkdown;
    document.getElementById('htmlPreview').innerHTML = currentHtml;
    document.getElementById('resultSection').classList.remove('hidden');

    const warnings = result.messages.filter(m => m.type === 'warning');
    if (warnings.length > 0) {
        showStatus('转换完成 (' + warnings.length + ' 个警告)', 'success');
    } else {
        showStatus('转换完成', 'success');
    }
}

// ==================== HTML 转 Markdown ====================

function htmlToMarkdown(html) {
    const doc = new DOMParser().parseFromString('<div>' + html + '</div>', 'text/html');
    const root = doc.body.firstChild;
    return convertNode(root).trim() + '\n';
}

function convertNode(node) {
    if (node.nodeType === Node.TEXT_NODE) {
        return node.textContent;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return '';

    const tag = node.tagName.toLowerCase();
    const children = Array.from(node.childNodes).map(convertNode).join('');

    switch (tag) {
        case 'h1': return '\n# ' + children.trim() + '\n\n';
        case 'h2': return '\n## ' + children.trim() + '\n\n';
        case 'h3': return '\n### ' + children.trim() + '\n\n';
        case 'h4': return '\n#### ' + children.trim() + '\n\n';
        case 'h5': return '\n##### ' + children.trim() + '\n\n';
        case 'h6': return '\n###### ' + children.trim() + '\n\n';
        case 'p': return children.trim() + '\n\n';
        case 'br': return '\n';
        case 'strong':
        case 'b': return '**' + children + '**';
        case 'em':
        case 'i': return '*' + children + '*';
        case 'u': return children; // Markdown 无下划线
        case 'a': {
            const href = node.getAttribute('href') || '';
            return '[' + children + '](' + href + ')';
        }
        case 'img': {
            const src = node.getAttribute('src') || '';
            const alt = node.getAttribute('alt') || '图片';
            return '![' + alt + '](' + src + ')';
        }
        case 'ul': return '\n' + convertList(node, false) + '\n';
        case 'ol': return '\n' + convertList(node, true) + '\n';
        case 'li': return children.trim();
        case 'table': return '\n' + convertTable(node) + '\n';
        case 'blockquote': {
            const lines = children.trim().split('\n');
            return '\n' + lines.map(l => '> ' + l).join('\n') + '\n\n';
        }
        case 'code': return '`' + children + '`';
        case 'pre': return '\n```\n' + children.trim() + '\n```\n\n';
        case 'hr': return '\n---\n\n';
        case 'sup': return '^' + children;
        case 'sub': return '~' + children;
        default: return children;
    }
}

function convertList(node, ordered) {
    const items = Array.from(node.children).filter(c => c.tagName.toLowerCase() === 'li');
    return items.map((li, i) => {
        const prefix = ordered ? (i + 1) + '. ' : '- ';
        const content = convertNode(li);
        return prefix + content;
    }).join('\n') + '\n';
}

function convertTable(tableNode) {
    const rows = [];
    tableNode.querySelectorAll('tr').forEach(tr => {
        const cells = [];
        tr.querySelectorAll('th, td').forEach(cell => {
            cells.push(convertNode(cell).trim().replace(/\|/g, '\\|').replace(/\n/g, ' '));
        });
        rows.push(cells);
    });

    if (rows.length === 0) return '';

    const colCount = Math.max(...rows.map(r => r.length));
    // 补齐列数
    rows.forEach(r => { while (r.length < colCount) r.push(''); });

    let md = '| ' + rows[0].join(' | ') + ' |\n';
    md += '| ' + rows[0].map(() => '---').join(' | ') + ' |\n';
    for (let i = 1; i < rows.length; i++) {
        md += '| ' + rows[i].join(' | ') + ' |\n';
    }
    return md;
}

// ==================== 下载 / 复制 ====================

function downloadMd() {
    if (!currentMarkdown) return;
    const blob = new Blob([currentMarkdown], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = currentFileName + '.md';
    a.click();
    URL.revokeObjectURL(url);
}

function copyMd() {
    if (!currentMarkdown) return;
    navigator.clipboard.writeText(currentMarkdown).then(() => {
        const btn = document.getElementById('copyBtn');
        btn.textContent = '已复制';
        setTimeout(() => { btn.textContent = '复制 Markdown'; }, 1500);
    }).catch(() => {
        // fallback
        const textarea = document.createElement('textarea');
        textarea.value = currentMarkdown;
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        document.body.removeChild(textarea);
    });
}
