document.addEventListener('DOMContentLoaded', function() {
    const inputText = document.getElementById('inputText');
    const outputArea = document.getElementById('outputArea');
    const extractBtn = document.getElementById('extractBtn');
    const clearBtn = document.getElementById('clearBtn');
    const copyBtn = document.getElementById('copyBtn');
    const countInfo = document.getElementById('countInfo');
    const uniqueInfo = document.getElementById('uniqueInfo');
    
    function extractSixDigitNumbers(text) {
        const regex = /\d{6}/g;
        const matches = text.match(regex);
        return matches || [];
    }
    
    function formatNumbers(numbers) {
        const uniqueNumbers = [...new Set(numbers)];
        uniqueNumbers.sort((a, b) => a - b);
        return uniqueNumbers.join('\n');
    }
    
    function getCurrentNumbers() {
        const text = inputText.value;
        const numbers = extractSixDigitNumbers(text);
        const uniqueNumbers = [...new Set(numbers)];
        uniqueNumbers.sort((a, b) => a - b);
        return uniqueNumbers;
    }
    
    function updateStats(numbers) {
        const uniqueNumbers = [...new Set(numbers)];
        countInfo.textContent = `提取到 ${numbers.length} 个六位数字`;
        uniqueInfo.textContent = `去重后 ${uniqueNumbers.length} 个唯一数字`;
    }

    extractBtn.addEventListener('click', function() {
        const text = inputText.value;
        const numbers = extractSixDigitNumbers(text);
        
        if (numbers.length === 0) {
            outputArea.textContent = "未找到六位数字，请检查输入文本。";
            countInfo.textContent = "提取到 0 个六位数字";
            uniqueInfo.textContent = "去重后 0 个唯一数字";
            return;
        }
        
        outputArea.textContent = formatNumbers(numbers);
        updateStats(numbers);
    });
    
    clearBtn.addEventListener('click', function() {
        inputText.value = '';
        outputArea.textContent = '提取结果将显示在这里...';
        countInfo.textContent = '提取到 0 个六位数字';
        uniqueInfo.textContent = '去重后 0 个唯一数字';
        inputText.focus();
    });
    
    copyBtn.addEventListener('click', function() {
        const numbers = getCurrentNumbers();
        if (numbers.length === 0) {
            alert('没有可复制的数字，请先提取数字。');
            return;
        }
        
        navigator.clipboard.writeText(numbers.join('\n'))
            .then(() => {
                const originalText = copyBtn.innerHTML;
                copyBtn.innerHTML = '已复制！';
                copyBtn.style.backgroundColor = '#1a7f37';
                setTimeout(() => {
                    copyBtn.innerHTML = originalText;
                    copyBtn.style.backgroundColor = '';
                }, 2000);
            })
            .catch(err => {
                alert('复制失败，请手动选择文本复制。');
            });
    });
});
