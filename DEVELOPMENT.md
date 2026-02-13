# 开发规范

飞行部工具箱，纯前端离线工具集。无需后端服务，双击 `index.html` 即可使用。

## 使用方法

1. 双击 `index.html` 在浏览器中打开
2. 输入密码登录
3. 导航栏选择模块：首页、工具、文档、友情链接
4. 点击工具进入使用

## 目录结构

```
├── index.html              # 主入口（Bootstrap导航栏 + iframe布局）
├── main.css                # 全局样式
├── main.js                 # 主逻辑（登录、模块加载、粒子效果）
├── assets/                 # 静态资源
│   ├── bg.webp             # 背景图
│   └── doge.gif            # 首页彩蛋
├── libs/                   # 第三方库
│   ├── bootstrap.min.css
│   ├── bootstrap.bundle.min.js
│   ├── pdf-lib.min.js      # PDF处理
│   ├── pdf.min.js          # PDF渲染
│   ├── pdf.worker.min.js
│   ├── xlsx.full.min.js    # Excel处理
│   ├── xlsx-js-style.min.js
│   ├── xlsx.bundle.js
│   ├── docxtemplater.min.js # Word模板处理
│   ├── pizzip.min.js       # ZIP处理
│   ├── jszip.min.js
│   ├── mammoth.min.js      # Word转HTML
│   ├── marked.min.js       # Markdown渲染
│   ├── tsparticles.bundle.min.js # 粒子效果
│   └── versions.json       # 库版本记录
├── home/                   # 首页模块
│   └── index.html
├── tool/                   # 工具模块
│   ├── index.html          # 工具列表页
│   ├── tools-data.js       # 工具配置
│   ├── tools-render.js     # 渲染逻辑
│   └── app/                # 具体工具页面
│       ├── word-template-filler/
│       ├── lock-entry-helper/
│       ├── aircrew-cert-form.html
│       ├── bill-gantt.html
│       ├── crew-extract-id.html
│       ├── crew-flight-stats.html
│       ├── crew-match-name-id.html
│       ├── hotel-bill-check.html
│       ├── meihua-video-form.html
│       └── sim-training-check.html
├── docs/                   # 文档模块
│   ├── index.html          # 文档列表
│   ├── docs-data.js        # 文档配置
│   ├── viewer.html         # Markdown渲染页
│   └── md/                 # Markdown文档
│       ├── word-template-filler.md
│       ├── hotel-bill-check.md
│       ├── bill-gantt.md
│       ├── lock-entry-helper.md
│       └── 更新说明.md
├── links/                  # 友情链接模块
│   └── index.html
└── template/               # 文档模板
    ├── kqdjz-sample/
    ├── meihua/
    ├── word-filler-sample/
    ├── 排班表.xlsx
    ├── 机组花名册.xlsx
    └── 通用批量锁班模板.xlsx
```

## 路径引用

- `tool/app/*.html` 引用 libs：`../../libs/xxx`
- `tool/app/*.html` 引用 template：`../../template/xxx`
- `tool/app/子目录/` 引用 libs：`../../../libs/xxx`

## 新增工具

1. 在 `tool/app/` 下创建工具文件（单文件或目录）
2. 在 `tool/tools-data.js` 的对应分类中添加配置：
   ```javascript
   { name: '工具名', desc: '描述', url: 'app/xxx.html' }
   ```

工具分类：test（测试）、training（训练）、operation（运行）、safety（安全）、tech（技术）、general（综合）、branch（分部）

## 新增文档

1. 在 `docs/md/` 下创建 `xxx.md`
2. 在 `docs/docs-data.js` 中添加配置：
   ```javascript
   { file: 'xxx', title: '文档标题' }
   ```

## 新增导航模块

1. 在根目录创建模块文件夹和 `index.html`
2. 在 `index.html` 导航栏 `<ul class="navbar-nav">` 中添加：
   ```html
   <li class="nav-item">
       <a class="nav-link" href="#" data-module="moduleName" onclick="loadModule('moduleName', event)">显示名</a>
   </li>
   ```
3. 在 `main.js` 的 `modules` 对象中添加模块路径

## 代码风格

- 统一浅色主题
- 中文界面

## 浏览器兼容

推荐 Chromium 内核浏览器（Chrome/Edge）。
