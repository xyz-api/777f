// 工具数据配置
// 添加新工具只需在对应分类的 items 数组中添加即可

const tools = [
    {
        category: 'test',
        categoryName: '测试',
        items: [
            { name: 'Word模板填充器（测试）', desc: '通用文档模板填充工具，上传配置和模板，自动生成表单并导出，具体使用方法请仔细阅读文档！', url: 'app/word-template-filler/index.html' },
            { name: '通用锁班助手（测试）', desc: 'Python脚本，自动化填写飞行门户非生产任务录入表单，支持批量录入锁班信息', url: 'app/lock-entry-helper/index.html' },
            { name: '飞行经历起落数按天统计（测试）', desc: 'Python脚本，自动化查询飞行门户飞行经历起落数并按天统计', url: 'app/flight-stats-helper/index.html' }
        ]
    },
    {
        category: 'training',
        categoryName: '训练',
        items: [
            { name: '姓名匹配员工号', desc: '从混杂文本中识别姓名，并匹配对应员工号，支持一键复制', url: 'app/crew-match-name-id.html' },
            { name: '提取员工号', desc: '从混杂文本中提取6位数字员工号，自动去重排序，支持一键复制', url: 'app/crew-extract-id.html' }
        ]
    },
    {
        category: 'operation',
        categoryName: '运行',
        items: [
            { name: '航线班次统计', desc: '根据排班表统计每人各航线班次', url: 'app/crew-flight-stats.html' },
            { name: '重点人员标注', desc: '根据重点人员表对审班表进行颜色标注，支持多类别标记', url: 'app/focus-crew.html' }
        ]
    },
    {
        category: 'safety',
        categoryName: '安全',
        items: []
    },
    {
        category: 'tech',
        categoryName: '技术',
        items: [
            { name: 'PDF加水印', desc: '手册工作使用，在PDF每页的相同位置添加图片，支持拖拽定位和精确数值调整', url: 'app/pdf-stamp.html' },
            { name: 'Markdown对比器', desc: '手册工作使用，对比两个Markdown文件的差异，支持并排/内联/渲染预览三种视图', url: 'app/md-diff.html' }
        ]
    },
    {
        category: 'general',
        categoryName: '综合',
        items: [
            { name: '酒店账单核对', desc: '对比酒店账单与入住登记表，核对用', url: 'app/hotel-bill-check.html' },
            { name: '账单甘特图', desc: '将酒店账单转换为甘特图，与飞行任务日程对比审核', url: 'app/bill-gantt.html' },
            { name: 'PDF工具箱', desc: 'PDF预览、页面提取、转图片、合并、图片转PDF，支持旋转和批量操作', url: 'app/pdf-tool.html' },
            { name: '图片工具箱', desc: '图片格式转换、压缩、调整尺寸、裁剪、Base64互转，支持批量操作', url: 'app/image-tool.html' },
            { name: 'Excel转Markdown', desc: '将Excel解析为结构化Markdown，方便提供给AI阅读，保留超链接和日期格式', url: 'app/excel-to-markdown.html' },
            { name: 'Word转Markdown', desc: '将Word文档转为Markdown，保留标题、列表、表格、超链接，支持预览和下载', url: 'app/word-to-markdown.html' }
        ]
    },
    {
        category: 'branch',
        categoryName: '分部',
        items: []
    }
];
