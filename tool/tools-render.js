// 工具页面渲染逻辑

let currentCategory = 'all';
let searchKeyword = '';

// 渲染工具列表
function renderTools() {
    const container = document.getElementById('toolSections');
    
    // 收集所有工具
    let allItems = [];
    tools.forEach(section => {
        section.items.forEach(item => {
            allItems.push({ ...item, category: section.category });
        });
    });
    
    // 按分类筛选
    if (currentCategory !== 'all') {
        allItems = allItems.filter(item => item.category === currentCategory);
    }
    
    // 搜索筛选
    if (searchKeyword) {
        const keyword = searchKeyword.toLowerCase();
        allItems = allItems.filter(item => 
            item.name.toLowerCase().includes(keyword) || 
            item.desc.toLowerCase().includes(keyword)
        );
    }
    
    if (allItems.length === 0) {
        container.innerHTML = '<div class="no-result">没有找到匹配的工具</div>';
        return;
    }
    
    let html = '<div class="tool-list">';
    allItems.forEach(item => {
        html += `<a href="${item.url}" target="_blank" class="tool-item">
            <div class="tool-name">${highlightKeyword(item.name)}</div>
            <div class="tool-desc">${highlightKeyword(item.desc)}</div>
        </a>`;
    });
    html += '</div>';
    html += `<div class="tool-count">共 ${allItems.length} 个工具</div>`;
    
    container.innerHTML = html;
}

// 高亮搜索关键词
function highlightKeyword(text) {
    if (!searchKeyword) return text;
    const regex = new RegExp(`(${searchKeyword})`, 'gi');
    return text.replace(regex, '<mark>$1</mark>');
}

// 标签切换
document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        currentCategory = btn.dataset.category;
        renderTools();
    });
});

// 搜索功能
const searchInput = document.getElementById('searchInput');
searchInput.addEventListener('input', (e) => {
    searchKeyword = e.target.value.trim();
    renderTools();
});

// 初始化
renderTools();
