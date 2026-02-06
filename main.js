const modules = {
    home: { title: '首页', url: 'home/index.html' },
    tool: { title: '工具', url: 'tool/index.html' },
    docs: { title: '文档', url: 'docs/index.html' },
    links: { title: '友情链接', url: 'links/index.html' }
};

let clickCount = 0;
let lastClickTime = 0;

function tryLogin() {
    const input = document.getElementById('passwordInput');
    const now = Date.now();
    
    // 如果输入了字符，显示密码错误
    if (input.value.trim() !== '') {
        showError();
        clickCount = 0;
        return;
    }
    
    // 检查点击间隔，超过500ms重新计数
    if (now - lastClickTime > 500) {
        clickCount = 0;
    }
    
    lastClickTime = now;
    clickCount++;
    
    if (clickCount >= 3) {
        enterApp();
    }
}

function showError() {
    const input = document.getElementById('passwordInput');
    input.value = '';
    input.placeholder = '密码错误';
    setTimeout(() => {
        input.placeholder = '请输入密码';
    }, 1500);
}

function enterApp() {
    const welcome = document.getElementById('welcomePage');
    const main = document.getElementById('mainApp');
    
    welcome.style.display = 'none';
    main.classList.add('visible');
    initBackground();
}

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
