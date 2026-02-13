// PDF工具箱
const { PDFDocument } = PDFLib;

// 初始化
document.addEventListener('DOMContentLoaded', function() {
    pdfjsLib.GlobalWorkerOptions.workerSrc = '../../libs/pdf.worker.min.js';
    initExtract();
    initMerge();
    initPdf2Img();
    initImg2Pdf();
});

// ==================== 页面提取 ====================
const extractState = { pdfDoc: null, arrayBuffer: null, pageCount: 0, selected: new Set(), lastClicked: null };

function initExtract() {
    setupUpload('extractUpload', 'extractInput', loadExtractPdf, false);
    document.getElementById('extractRange').addEventListener('input', e => {
        extractState.selected = new Set(parseRange(e.target.value, extractState.pageCount));
        syncExtractUI();
    });
}

async function loadExtractPdf(file) {
    const info = document.getElementById('extractInfo');
    info.className = 'status-bar'; info.classList.remove('hidden');
    info.textContent = '加载中...';
    
    try {
        extractState.arrayBuffer = await file.arrayBuffer();
        extractState.pdfDoc = await pdfjsLib.getDocument({ data: extractState.arrayBuffer.slice(0) }).promise;
        extractState.pageCount = extractState.pdfDoc.numPages;
        extractState.selected.clear();
        extractState.lastClicked = null;
        
        info.textContent = `${file.name} - 共${extractState.pageCount}页`;
        document.getElementById('extractPreview').classList.remove('hidden');
        document.getElementById('extractGrid').innerHTML = '';
        document.getElementById('loadPreviewBtn').disabled = false;
        document.getElementById('loadPreviewBtn').textContent = '加载预览';
    } catch (err) {
        info.className = 'status-bar error';
        info.textContent = '加载失败: ' + err.message;
    }
}

async function startRenderPreview() {
    const btn = document.getElementById('loadPreviewBtn');
    btn.disabled = true;
    btn.textContent = '加载中...';
    
    const grid = document.getElementById('extractGrid');
    const total = extractState.pageCount;
    const info = document.getElementById('extractInfo');
    const baseInfo = info.textContent.split('|')[0].trim();
    grid.innerHTML = '';
    
    for (let i = 1; i <= total; i++) {
        info.textContent = `${baseInfo} | 渲染 ${i}/${total}...`;
        btn.textContent = `加载中 ${i}/${total}`;
        
        const page = await extractState.pdfDoc.getPage(i);
        const viewport = page.getViewport({ scale: 0.3 });
        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
        
        const div = document.createElement('div');
        div.className = 'preview-page';
        div.dataset.page = i;
        div.innerHTML = `<div class="page-label">${i}</div>`;
        div.insertBefore(canvas, div.firstChild);
        div.onclick = e => handleExtractClick(i, e);
        grid.appendChild(div);
    }
    
    info.textContent = baseInfo;
    btn.textContent = '预览已加载';
    syncExtractUI();
}

function handleExtractClick(page, e) {
    if (e.shiftKey && extractState.lastClicked) {
        const [a, b] = [extractState.lastClicked, page].sort((x, y) => x - y);
        for (let i = a; i <= b; i++) extractState.selected.add(i);
    } else if (e.ctrlKey || e.metaKey) {
        extractState.selected.has(page) ? extractState.selected.delete(page) : extractState.selected.add(page);
    } else {
        extractState.selected.clear();
        extractState.selected.add(page);
    }
    extractState.lastClicked = page;
    syncExtractUI();
}

function syncExtractUI() {
    document.querySelectorAll('#extractGrid .preview-page').forEach(el => {
        el.classList.toggle('selected', extractState.selected.has(parseInt(el.dataset.page)));
    });
    document.getElementById('extractRange').value = formatRange(extractState.selected);
    document.getElementById('extractCount').textContent = extractState.selected.size > 0 ? `已选${extractState.selected.size}页` : '';
}

function extractSelectAll() {
    for (let i = 1; i <= extractState.pageCount; i++) extractState.selected.add(i);
    syncExtractUI();
}
function extractSelectNone() {
    extractState.selected.clear();
    syncExtractUI();
}

async function doExtract() {
    const pages = extractState.selected.size > 0 
        ? Array.from(extractState.selected).sort((a, b) => a - b) 
        : Array.from({ length: extractState.pageCount }, (_, i) => i + 1);
    
    if (pages.length === 0) { alert('请选择页面'); return; }
    
    const info = document.getElementById('extractInfo');
    info.textContent = '提取中...';
    
    try {
        const srcDoc = await PDFDocument.load(extractState.arrayBuffer);
        const newDoc = await PDFDocument.create();
        const copied = await newDoc.copyPages(srcDoc, pages.map(p => p - 1));
        copied.forEach(p => newDoc.addPage(p));
        
        const bytes = await newDoc.save();
        const customName = document.getElementById('extractFileName').value.trim();
        const fileName = customName ? (customName.endsWith('.pdf') ? customName : customName + '.pdf') : `extracted_${pages.length}pages.pdf`;
        download(new Blob([bytes], { type: 'application/pdf' }), fileName);
        info.textContent = `已提取${pages.length}页`;
    } catch (err) {
        info.className = 'status-bar error';
        info.textContent = '提取失败: ' + err.message;
    }
}

// ==================== PDF合并 ====================
const mergeState = { files: [], nextId: 1 };

function initMerge() {
    setupUpload('mergeUpload', 'mergeInput', addMergeFiles, true);
}

async function addMergeFiles(files) {
    for (const file of files) {
        if (file.type !== 'application/pdf') continue;
        try {
            const ab = await file.arrayBuffer();
            const doc = await PDFDocument.load(ab);
            mergeState.files.push({
                id: mergeState.nextId++,
                name: file.name,
                size: file.size,
                pageCount: doc.getPageCount(),
                arrayBuffer: ab
            });
        } catch (err) {
            alert(`无法读取 ${file.name}`);
        }
    }
    renderMergeList();
}

function renderMergeList() {
    const list = document.getElementById('mergeList');
    list.innerHTML = mergeState.files.map((f, i) => `
        <div class="file-item" draggable="true" data-id="${f.id}">
            <span class="drag-handle">⋮⋮</span>
            <div class="file-info">
                <div class="file-name">${f.name}</div>
                <div class="file-meta">${f.pageCount}页 · ${formatSize(f.size)}</div>
            </div>
            <button class="btn btn-sm btn-outline-danger" onclick="removeMergeFile(${f.id})">×</button>
        </div>
    `).join('');
    
    document.getElementById('mergeActions').classList.toggle('hidden', mergeState.files.length === 0);
    initDragSort(list, mergeState.files);
}

function removeMergeFile(id) {
    mergeState.files = mergeState.files.filter(f => f.id !== id);
    renderMergeList();
}
function clearMergeList() {
    mergeState.files = [];
    renderMergeList();
    document.getElementById('mergeStatus').classList.add('hidden');
}

async function doMerge() {
    if (mergeState.files.length < 2) { alert('请至少添加2个PDF'); return; }
    
    const status = document.getElementById('mergeStatus');
    status.className = 'status-bar'; status.classList.remove('hidden');
    status.textContent = '合并中...';
    
    try {
        const merged = await PDFDocument.create();
        let totalPages = 0;
        
        for (const f of mergeState.files) {
            const src = await PDFDocument.load(f.arrayBuffer);
            const pages = await merged.copyPages(src, src.getPageIndices());
            pages.forEach(p => merged.addPage(p));
            totalPages += pages.length;
        }
        
        const bytes = await merged.save();
        download(new Blob([bytes], { type: 'application/pdf' }), `merged_${mergeState.files.length}files.pdf`);
        status.textContent = `已合并${mergeState.files.length}个文件，共${totalPages}页`;
    } catch (err) {
        status.className = 'status-bar error';
        status.textContent = '合并失败: ' + err.message;
    }
}

// ==================== PDF转图片 ====================
const pdf2imgState = { files: [], nextId: 1 };

function initPdf2Img() {
    setupUpload('pdf2imgUpload', 'pdf2imgInput', addPdf2ImgFiles, true);
}

async function addPdf2ImgFiles(files) {
    const fileArr = Array.isArray(files) ? files : [files];
    for (const file of fileArr) {
        if (file.type !== 'application/pdf') continue;
        const ab = await file.arrayBuffer();
        pdf2imgState.files.push({
            id: pdf2imgState.nextId++,
            name: file.name,
            baseName: file.name.replace(/\.pdf$/i, ''),
            size: file.size,
            arrayBuffer: ab
        });
    }
    renderPdf2ImgList();
}

function renderPdf2ImgList() {
    const list = document.getElementById('pdf2imgList');
    list.innerHTML = pdf2imgState.files.map(f => `
        <div class="file-item" data-id="${f.id}">
            <div class="file-info">
                <div class="file-name">${f.name}</div>
                <div class="file-meta">${formatSize(f.size)}</div>
            </div>
            <button class="btn btn-sm btn-outline-danger" onclick="removePdf2ImgFile(${f.id})">×</button>
        </div>
    `).join('');
    document.getElementById('pdf2imgOptions').classList.toggle('hidden', pdf2imgState.files.length === 0);
}

function removePdf2ImgFile(id) {
    pdf2imgState.files = pdf2imgState.files.filter(f => f.id !== id);
    renderPdf2ImgList();
}

function clearPdf2ImgList() {
    pdf2imgState.files = [];
    renderPdf2ImgList();
    document.getElementById('pdf2imgInfo').classList.add('hidden');
    document.getElementById('pdf2imgGrid').innerHTML = '';
}

async function doPdf2Img() {
    if (pdf2imgState.files.length === 0) return;
    
    const format = document.getElementById('imgFormat').value;
    const scale = parseFloat(document.getElementById('imgScale').value);
    const info = document.getElementById('pdf2imgInfo');
    const grid = document.getElementById('pdf2imgGrid');
    info.className = 'status-bar'; info.classList.remove('hidden');
    grid.innerHTML = '';
    
    const zip = new JSZip();
    const isSingle = pdf2imgState.files.length === 1;
    
    try {
        for (let fi = 0; fi < pdf2imgState.files.length; fi++) {
            const f = pdf2imgState.files[fi];
            info.textContent = `处理 ${f.name} (${fi + 1}/${pdf2imgState.files.length})...`;
            
            const pdfDoc = await pdfjsLib.getDocument({ data: f.arrayBuffer.slice(0) }).promise;
            const folder = isSingle ? zip : zip.folder(f.baseName);
            
            for (let i = 1; i <= pdfDoc.numPages; i++) {
                info.textContent = `${f.name} - 第${i}/${pdfDoc.numPages}页`;
                
                const page = await pdfDoc.getPage(i);
                const viewport = page.getViewport({ scale });
                const canvas = document.createElement('canvas');
                canvas.width = viewport.width;
                canvas.height = viewport.height;
                await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
                
                const dataUrl = canvas.toDataURL(`image/${format}`, 0.92);
                const ext = format === 'jpeg' ? 'jpg' : 'png';
                const imgName = `${f.baseName}_page${i}.${ext}`;
                folder.file(imgName, dataUrl.split(',')[1], { base64: true });
                
                // 单文件时显示预览
                if (isSingle) {
                    const item = document.createElement('div');
                    item.className = 'result-item';
                    item.innerHTML = `
                        <img src="${dataUrl}">
                        <div class="item-footer">
                            <span>第${i}页</span>
                            <a href="${dataUrl}" download="${imgName}">下载</a>
                        </div>
                    `;
                    grid.appendChild(item);
                }
            }
        }
        
        info.textContent = '打包下载中...';
        const blob = await zip.generateAsync({ type: 'blob' });
        const zipName = isSingle ? `${pdf2imgState.files[0].baseName}_images.zip` : `pdf_images_${pdf2imgState.files.length}files.zip`;
        download(blob, zipName);
        info.textContent = `转换完成，已下载zip`;
    } catch (err) {
        info.className = 'status-bar error';
        info.textContent = '转换失败: ' + err.message;
    }
}

// 删除旧的 downloadImagesZip 函数

// ==================== 图片转PDF ====================
const img2pdfState = { images: [], nextId: 1 };

function initImg2Pdf() {
    setupUpload('img2pdfUpload', 'img2pdfInput', addImg2PdfFiles, true);
}

async function addImg2PdfFiles(files) {
    for (const file of files) {
        if (!file.type.startsWith('image/')) continue;
        const dataUrl = await readFileAsDataUrl(file);
        img2pdfState.images.push({
            id: img2pdfState.nextId++,
            name: file.name,
            size: file.size,
            type: file.type,
            dataUrl
        });
    }
    renderImg2PdfList();
}

function renderImg2PdfList() {
    const list = document.getElementById('img2pdfList');
    list.innerHTML = img2pdfState.images.map(img => `
        <div class="file-item" draggable="true" data-id="${img.id}">
            <span class="drag-handle">⋮⋮</span>
            <img src="${img.dataUrl}" style="width:40px;height:40px;object-fit:cover;border-radius:4px;">
            <div class="file-info">
                <div class="file-name">${img.name}</div>
                <div class="file-meta">${formatSize(img.size)}</div>
            </div>
            <button class="btn btn-sm btn-outline-danger" onclick="removeImg2PdfFile(${img.id})">×</button>
        </div>
    `).join('');
    
    document.getElementById('img2pdfActions').classList.toggle('hidden', img2pdfState.images.length === 0);
    initDragSort(list, img2pdfState.images);
}

function removeImg2PdfFile(id) {
    img2pdfState.images = img2pdfState.images.filter(i => i.id !== id);
    renderImg2PdfList();
}
function clearImg2PdfList() {
    img2pdfState.images = [];
    renderImg2PdfList();
    document.getElementById('img2pdfStatus').classList.add('hidden');
}

async function doImg2Pdf() {
    if (img2pdfState.images.length === 0) { alert('请添加图片'); return; }
    
    const status = document.getElementById('img2pdfStatus');
    status.className = 'status-bar'; status.classList.remove('hidden');
    status.textContent = '生成中...';
    
    try {
        const pdfDoc = await PDFDocument.create();
        
        for (let i = 0; i < img2pdfState.images.length; i++) {
            status.textContent = `处理第${i + 1}/${img2pdfState.images.length}张...`;
            const img = img2pdfState.images[i];
            const bytes = await fetch(img.dataUrl).then(r => r.arrayBuffer());
            
            let embedded;
            if (img.type === 'image/png') {
                embedded = await pdfDoc.embedPng(bytes);
            } else {
                embedded = await pdfDoc.embedJpg(bytes);
            }
            
            const page = pdfDoc.addPage([embedded.width, embedded.height]);
            page.drawImage(embedded, { x: 0, y: 0, width: embedded.width, height: embedded.height });
        }
        
        const pdfBytes = await pdfDoc.save();
        download(new Blob([pdfBytes], { type: 'application/pdf' }), `images_${img2pdfState.images.length}pages.pdf`);
        status.textContent = `已生成${img2pdfState.images.length}页PDF`;
    } catch (err) {
        status.className = 'status-bar error';
        status.textContent = '生成失败: ' + err.message;
    }
}

// ==================== 工具函数 ====================
function setupUpload(areaId, inputId, handler, multiple) {
    const area = document.getElementById(areaId);
    const input = document.getElementById(inputId);
    
    area.onclick = () => input.click();
    area.ondragover = e => { e.preventDefault(); area.classList.add('dragover'); };
    area.ondragleave = () => area.classList.remove('dragover');
    area.ondrop = e => {
        e.preventDefault();
        area.classList.remove('dragover');
        handler(multiple ? Array.from(e.dataTransfer.files) : e.dataTransfer.files[0]);
    };
    input.onchange = e => {
        handler(multiple ? Array.from(e.target.files) : e.target.files[0]);
        input.value = '';
    };
}

function initDragSort(container, dataArray) {
    let dragged = null;
    container.querySelectorAll('[draggable="true"]').forEach(item => {
        item.ondragstart = () => { dragged = item; item.style.opacity = '0.5'; };
        item.ondragend = () => { item.style.opacity = '1'; dragged = null; };
        item.ondragover = e => e.preventDefault();
        item.ondrop = e => {
            e.preventDefault();
            if (dragged && item !== dragged) {
                const fromId = parseInt(dragged.dataset.id);
                const toId = parseInt(item.dataset.id);
                const fromIdx = dataArray.findIndex(x => x.id === fromId);
                const toIdx = dataArray.findIndex(x => x.id === toId);
                if (fromIdx !== -1 && toIdx !== -1) {
                    const [removed] = dataArray.splice(fromIdx, 1);
                    dataArray.splice(toIdx, 0, removed);
                    if (container.id === 'mergeList') renderMergeList();
                    else if (container.id === 'img2pdfList') renderImg2PdfList();
                }
            }
        };
    });
}

function parseRange(str, max) {
    if (!str.trim()) return [];
    const pages = new Set();
    for (const part of str.split(',')) {
        const t = part.trim();
        if (!t) continue;
        if (t.includes('-')) {
            const [a, b] = t.split('-').map(s => parseInt(s.trim()));
            if (!isNaN(a) && !isNaN(b)) {
                for (let i = Math.max(1, a); i <= Math.min(max, b); i++) pages.add(i);
            }
        } else {
            const n = parseInt(t);
            if (!isNaN(n) && n >= 1 && n <= max) pages.add(n);
        }
    }
    return Array.from(pages).sort((a, b) => a - b);
}

function formatRange(set) {
    if (set.size === 0) return '';
    const sorted = Array.from(set).sort((a, b) => a - b);
    const ranges = [];
    let start = sorted[0], end = sorted[0];
    for (let i = 1; i <= sorted.length; i++) {
        if (i < sorted.length && sorted[i] === end + 1) {
            end = sorted[i];
        } else {
            ranges.push(start === end ? String(start) : `${start}-${end}`);
            if (i < sorted.length) start = end = sorted[i];
        }
    }
    return ranges.join(',');
}

function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function readFileAsDataUrl(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

function download(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}
