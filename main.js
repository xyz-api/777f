const modules = {
    home: { title: '首页', url: 'home/index.html' },
    tool: { title: '工具', url: 'tool/index.html' },
    docs: { title: '文档', url: 'docs/index.html' },
    links: { title: '友情链接', url: 'links/index.html' }
};

// 页面加载后初始化背景
document.addEventListener('DOMContentLoaded', initBackground);

function initBackground() {
    const app = document.getElementById('mainApp');
    const video = document.getElementById('bgVideo');
    
    if (CONFIG.bgType === 'video') {
        video.src = CONFIG.bgVideo;
        video.style.display = 'block';
        app.style.backgroundImage = 'none';
    } else {
        video.style.display = 'none';
        video.src = '';
        app.style.backgroundImage = `url('${CONFIG.bgImage}')`;
    }
}

function loadModule(name, event) {
    if (event) event.preventDefault();
    const m = modules[name];
    if (!m) return;
    document.querySelectorAll('.navbar-nav .nav-link').forEach(link => {
        link.classList.toggle('active', link.dataset.module === name);
    });
    document.getElementById('contentFrame').src = m.url;
}
