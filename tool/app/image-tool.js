// 图片工具箱

// ==================== 工具函数 ====================
function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function getBaseName(filename) {
    return filename.replace(/\.[^/.]+$/, '');
}

function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}

function setupUpload(areaId, inputId, handler, multiple) {
    const area = document.getElementById(areaId);
    const input = document.getElementById(inputId);
    area.onclick = () => input.click();
    area.ondragover = e => { e.preventDefault(); area.classList.add('dragover'); };
    area.ondragleave = () => area.classList.remove('dragover');
    area.ondrop = e => {
        e.preventDefault();
        area.classList.remove('dragover');
        handler(multiple ? Array.from(e.dataTransfer.files) : [e.dataTransfer.files[0]]);
    };
    input.onchange = e => {
        handler(Array.from(e.target.files));
        input.value = '';
    };
}

function renderImageList(images, listEl, optionsEl, onRemove) {
    if (images.length === 0) {
        listEl.innerHTML = '';
        optionsEl.classList.add('hidden');
        return;
    }
    optionsEl.classList.remove('hidden');
    listEl.innerHTML = images.map((img, i) => `
        <div class="image-item">
            <img src="${img.url}">
            <div class="info">${formatSize(img.file.size)}</div>
            <div class="preview-info" id="${listEl.id}-preview-${i}"></div>
            <button class="remove-btn" data-i="${i}">&times;</button>
        </div>
    `).join('');
    listEl.querySelectorAll('.remove-btn').forEach(btn => {
        btn.onclick = () => onRemove(parseInt(btn.dataset.i));
    });
}

function renderResultItem(container, blob, text, filename) {
    const div = document.createElement('div');
    div.className = 'result-item';
    div.innerHTML = `
        <img src="${URL.createObjectURL(blob)}">
        <span class="meta">${text}</span>
        <button class="btn btn-outline-secondary btn-sm">下载</button>
    `;
    div.querySelector('button').onclick = () => downloadBlob(blob, filename);
    container.appendChild(div);
}


// ==================== 格式转换 ====================
(function() {
    const images = [];
    const listEl = document.getElementById('convertList');
    const optionsEl = document.getElementById('convertOptions');
    const resultsEl = document.getElementById('convertResults');
    const qualitySlider = document.getElementById('convertQuality');
    const qualityVal = document.getElementById('convertQualityVal');

    setupUpload('convertUpload', 'convertInput', addImages, true);
    qualitySlider.addEventListener('input', () => { qualityVal.textContent = qualitySlider.value; updatePreview(); });
    document.getElementById('convertFormat').addEventListener('change', updatePreview);
    document.getElementById('convertBtn').addEventListener('click', convert);
    document.getElementById('convertClearBtn').addEventListener('click', () => { images.length = 0; render(); resultsEl.innerHTML = ''; });

    function addImages(files) {
        files.forEach(f => { if (f.type.startsWith('image/')) images.push({ file: f, url: URL.createObjectURL(f) }); });
        render();
        updatePreview();
    }

    function render() {
        renderImageList(images, listEl, optionsEl, i => { images.splice(i, 1); render(); updatePreview(); });
    }

    function convertImage(file, format, quality) {
        return new Promise(resolve => {
            const img = new Image();
            img.onload = () => {
                const canvas = document.createElement('canvas');
                canvas.width = img.width;
                canvas.height = img.height;
                canvas.getContext('2d').drawImage(img, 0, 0);
                const mime = format === 'png' ? 'image/png' : format === 'webp' ? 'image/webp' : 'image/jpeg';
                canvas.toBlob(blob => resolve({ blob, width: img.width, height: img.height }), mime, quality);
            };
            img.src = URL.createObjectURL(file);
        });
    }

    async function updatePreview() {
        if (images.length === 0) return;
        const format = document.getElementById('convertFormat').value;
        const quality = qualitySlider.value / 100;
        for (let i = 0; i < images.length; i++) {
            const el = document.getElementById('convertList-preview-' + i);
            if (el) el.textContent = '计算中...';
        }
        for (let i = 0; i < images.length; i++) {
            const result = await convertImage(images[i].file, format, quality);
            const el = document.getElementById('convertList-preview-' + i);
            if (el) {
                const diff = ((result.blob.size / images[i].file.size - 1) * 100).toFixed(1);
                el.textContent = '→ ' + formatSize(result.blob.size) + ' (' + (diff >= 0 ? '+' : '') + diff + '%)';
            }
        }
    }

    async function convert() {
        const format = document.getElementById('convertFormat').value;
        const quality = qualitySlider.value / 100;
        resultsEl.innerHTML = '';
        for (const img of images) {
            const result = await convertImage(img.file, format, quality);
            const filename = getBaseName(img.file.name) + '.' + (format === 'jpeg' ? 'jpg' : format);
            renderResultItem(resultsEl, result.blob, result.width + 'x' + result.height + ' | ' + formatSize(result.blob.size), filename);
        }
    }
})();

// ==================== 图片压缩 ====================
(function() {
    const images = [];
    const listEl = document.getElementById('compressList');
    const optionsEl = document.getElementById('compressOptions');
    const resultsEl = document.getElementById('compressResults');
    const qualitySlider = document.getElementById('compressQuality');
    const qualityVal = document.getElementById('compressQualityVal');

    setupUpload('compressUpload', 'compressInput', addImages, true);
    qualitySlider.addEventListener('input', () => { qualityVal.textContent = qualitySlider.value; updatePreview(); });
    document.getElementById('compressMaxWidth').addEventListener('input', updatePreview);
    document.getElementById('compressBtn').addEventListener('click', compress);
    document.getElementById('compressClearBtn').addEventListener('click', () => { images.length = 0; render(); resultsEl.innerHTML = ''; });

    function addImages(files) {
        files.forEach(f => { if (f.type.startsWith('image/')) images.push({ file: f, url: URL.createObjectURL(f) }); });
        render();
        updatePreview();
    }

    function render() {
        renderImageList(images, listEl, optionsEl, i => { images.splice(i, 1); render(); updatePreview(); });
    }

    async function updatePreview() {
        if (images.length === 0) return;
        const quality = qualitySlider.value / 100;
        const maxWidth = parseInt(document.getElementById('compressMaxWidth').value) || 0;
        for (let i = 0; i < images.length; i++) {
            const el = document.getElementById('compressList-preview-' + i);
            if (el) el.textContent = '计算中...';
        }
        for (let i = 0; i < images.length; i++) {
            try {
                const opts = { maxSizeMB: 10, useWebWorker: true, initialQuality: quality };
                if (maxWidth > 0) opts.maxWidthOrHeight = maxWidth;
                const compressed = await imageCompression(images[i].file, opts);
                const el = document.getElementById('compressList-preview-' + i);
                if (el) {
                    const saved = ((1 - compressed.size / images[i].file.size) * 100).toFixed(1);
                    el.textContent = '→ ' + formatSize(compressed.size) + ' (-' + saved + '%)';
                }
            } catch (e) {
                const el = document.getElementById('compressList-preview-' + i);
                if (el) el.textContent = '';
            }
        }
    }

    async function compress() {
        const quality = qualitySlider.value / 100;
        const maxWidth = parseInt(document.getElementById('compressMaxWidth').value) || 0;
        resultsEl.innerHTML = '';
        for (const img of images) {
            try {
                const opts = { maxSizeMB: 10, useWebWorker: true, initialQuality: quality };
                if (maxWidth > 0) opts.maxWidthOrHeight = maxWidth;
                const compressed = await imageCompression(img.file, opts);
                const saved = ((1 - compressed.size / img.file.size) * 100).toFixed(1);
                const ext = compressed.type.includes('png') ? 'png' : 'jpg';
                renderResultItem(resultsEl, compressed, formatSize(img.file.size) + ' → ' + formatSize(compressed.size) + ' (节省' + saved + '%)', getBaseName(img.file.name) + '_compressed.' + ext);
            } catch (e) {
                console.error('压缩失败:', e);
            }
        }
    }
})();

// ==================== 调整尺寸 ====================
(function() {
    const images = [];
    const listEl = document.getElementById('resizeList');
    const optionsEl = document.getElementById('resizeOptions');
    const resultsEl = document.getElementById('resizeResults');
    const picaInstance = window.pica ? window.pica() : new pica();

    setupUpload('resizeUpload', 'resizeInput', addImages, true);
    document.getElementById('resizeWidth').addEventListener('input', updatePreview);
    document.getElementById('resizeHeight').addEventListener('input', updatePreview);
    document.getElementById('resizeKeepRatio').addEventListener('change', updatePreview);
    document.getElementById('resizeBtn').addEventListener('click', doResize);
    document.getElementById('resizeClearBtn').addEventListener('click', () => { images.length = 0; render(); resultsEl.innerHTML = ''; });

    function addImages(files) {
        files.forEach(f => { if (f.type.startsWith('image/')) images.push({ file: f, url: URL.createObjectURL(f) }); });
        render();
        updatePreview();
    }

    function render() {
        renderImageList(images, listEl, optionsEl, i => { images.splice(i, 1); render(); updatePreview(); });
    }

    function resizeImage(file, targetW, targetH, keepRatio) {
        return new Promise(resolve => {
            const img = new Image();
            img.onload = async () => {
                let w = targetW, h = targetH;
                if (keepRatio) {
                    const ratio = Math.min(targetW / img.width, targetH / img.height);
                    w = Math.round(img.width * ratio);
                    h = Math.round(img.height * ratio);
                }
                const srcCanvas = document.createElement('canvas');
                srcCanvas.width = img.width;
                srcCanvas.height = img.height;
                srcCanvas.getContext('2d').drawImage(img, 0, 0);
                const destCanvas = document.createElement('canvas');
                destCanvas.width = w;
                destCanvas.height = h;
                await picaInstance.resize(srcCanvas, destCanvas, { quality: 3, alpha: true });
                destCanvas.toBlob(blob => resolve({ blob, width: w, height: h }), 'image/png');
            };
            img.src = URL.createObjectURL(file);
        });
    }

    async function updatePreview() {
        if (images.length === 0) return;
        const tw = parseInt(document.getElementById('resizeWidth').value);
        const th = parseInt(document.getElementById('resizeHeight').value);
        const keep = document.getElementById('resizeKeepRatio').checked;
        for (let i = 0; i < images.length; i++) {
            const el = document.getElementById('resizeList-preview-' + i);
            if (el) el.textContent = '计算中...';
        }
        for (let i = 0; i < images.length; i++) {
            try {
                const result = await resizeImage(images[i].file, tw, th, keep);
                const el = document.getElementById('resizeList-preview-' + i);
                if (el) el.textContent = '→ ' + result.width + 'x' + result.height + ' | ' + formatSize(result.blob.size);
            } catch (e) {
                const el = document.getElementById('resizeList-preview-' + i);
                if (el) el.textContent = '';
            }
        }
    }

    async function doResize() {
        const tw = parseInt(document.getElementById('resizeWidth').value);
        const th = parseInt(document.getElementById('resizeHeight').value);
        const keep = document.getElementById('resizeKeepRatio').checked;
        resultsEl.innerHTML = '';
        for (const img of images) {
            const result = await resizeImage(img.file, tw, th, keep);
            renderResultItem(resultsEl, result.blob, result.width + 'x' + result.height + ' | ' + formatSize(result.blob.size), getBaseName(img.file.name) + '_resized.png');
        }
    }
})();


// ==================== 图片裁剪 ====================
(function() {
    let cropper = null;
    let croppedBlob = null;

    setupUpload('cropUpload', 'cropInput', files => { if (files[0]) loadImage(files[0]); }, false);
    document.getElementById('cropBtn').addEventListener('click', crop);
    document.getElementById('cropResetBtn').addEventListener('click', reset);
    document.getElementById('cropDownloadBtn').addEventListener('click', () => { if (croppedBlob) downloadBlob(croppedBlob, 'cropped.png'); });

    function loadImage(file) {
        const img = document.getElementById('cropImage');
        img.src = URL.createObjectURL(file);
        img.onload = () => {
            document.getElementById('cropEditor').classList.remove('hidden');
            document.getElementById('cropResult').classList.add('hidden');
            if (cropper) cropper.destroy();
            cropper = new Cropper(img, { aspectRatio: NaN, viewMode: 1, autoCropArea: 0.8, responsive: true });
        };
    }

    function crop() {
        if (!cropper) return;
        const canvas = cropper.getCroppedCanvas();
        canvas.toBlob(blob => {
            croppedBlob = blob;
            document.getElementById('cropOutput').src = URL.createObjectURL(blob);
            document.getElementById('cropInfo').textContent = canvas.width + 'x' + canvas.height + ' | ' + formatSize(blob.size);
            document.getElementById('cropResult').classList.remove('hidden');
        }, 'image/png');
    }

    function reset() {
        if (cropper) { cropper.destroy(); cropper = null; }
        croppedBlob = null;
        document.getElementById('cropEditor').classList.add('hidden');
        document.getElementById('cropResult').classList.add('hidden');
    }
})();

// ==================== Base64互转 ====================
(function() {
    let imageBlob = null;

    setupUpload('b64Upload', 'b64Input', files => { if (files[0]) toBase64(files[0]); }, false);
    document.getElementById('b64CopyBtn').addEventListener('click', () => {
        const output = document.getElementById('base64Output');
        output.select();
        navigator.clipboard.writeText(output.value).then(() => {
            const btn = document.getElementById('b64CopyBtn');
            btn.textContent = '已复制';
            setTimeout(() => { btn.textContent = '复制'; }, 1500);
        }).catch(() => { document.execCommand('copy'); });
    });
    document.getElementById('b64ConvertBtn').addEventListener('click', fromBase64);
    document.getElementById('b64DownloadBtn').addEventListener('click', () => { if (imageBlob) downloadBlob(imageBlob, 'image.png'); });

    function toBase64(file) {
        const reader = new FileReader();
        reader.onload = () => {
            document.getElementById('base64Output').value = reader.result;
            document.getElementById('b64ToResult').classList.remove('hidden');
        };
        reader.readAsDataURL(file);
    }

    function fromBase64() {
        let base64 = document.getElementById('base64Input').value.trim();
        if (!base64) return;
        if (!base64.startsWith('data:')) base64 = 'data:image/png;base64,' + base64;
        const img = new Image();
        img.onload = () => {
            const canvas = document.createElement('canvas');
            canvas.width = img.width;
            canvas.height = img.height;
            canvas.getContext('2d').drawImage(img, 0, 0);
            canvas.toBlob(blob => {
                imageBlob = blob;
                document.getElementById('b64Output').src = URL.createObjectURL(blob);
                document.getElementById('b64Info').textContent = img.width + 'x' + img.height + ' | ' + formatSize(blob.size);
                document.getElementById('b64FromResult').classList.remove('hidden');
            }, 'image/png');
        };
        img.onerror = () => alert('无效的Base64编码');
        img.src = base64;
    }
})();
