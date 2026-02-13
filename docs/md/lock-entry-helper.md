# 通用锁班助手（测试） - 使用教程

这是一份面向新手的详细教程，教你如何安装和使用通用锁班助手。

---

## 一、这个工具是干什么的？

通用锁班助手是一个自动化脚本，可以帮你批量填写飞行门户的"非生产任务录入"表单。

比如你有一份 Excel 表格，里面有 50 个人的锁班信息，正常情况下你需要手动填 50 次表单。用这个工具，你只需要把数据粘贴进去，它会自动帮你一条一条填写并提交。

**主要功能：**
- 批量模式：粘贴多条记录，自动逐条填表提交
- 手动模式：一条一条录入
- 白名单：只处理指定的员工号
- 冲突检测：提交后自动检查是否有冲突，有冲突会暂停让你处理

---

## 二、安装 Python

这个工具需要 Python 环境才能运行。如果你电脑上还没有 Python，按以下步骤安装：

### Windows 系统

1. 打开 Python 官网下载页面：https://www.python.org/downloads/
2. 点击黄色的 "Download Python 3.x.x" 按钮下载安装包
3. 运行安装包，**重要：勾选 "Add Python to PATH"**，然后点 "Install Now"
4. 安装完成后，打开命令提示符（按 Win+R，输入 cmd，回车），输入 `python --version`，如果显示版本号说明安装成功

### 验证安装

打开命令提示符，输入：
```
python --version
```
如果显示类似 `Python 3.11.5` 的内容，说明安装成功。

---

## 三、下载脚本文件

1. 在工具页面点击"下载 app.py"按钮
2. 把下载的 `app.py` 文件保存到一个你能找到的文件夹，比如桌面或者 `D:\tools\`

---

## 四、安装依赖库

脚本需要两个额外的库才能运行，打开命令提示符，依次执行以下命令。

**推荐使用清华镜像加速下载**（国内网络更快更稳定）：

### 1. 安装 colorama（彩色输出）
```
pip install colorama -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 2. 安装 playwright（浏览器自动化）
```
pip install playwright -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 3. 安装浏览器驱动
```
playwright install chromium
```

这一步会下载一个 Chromium 浏览器，可能需要几分钟，请耐心等待。

> 如果第3步下载很慢，可以设置环境变量使用国内CDN：
> ```
> set PLAYWRIGHT_DOWNLOAD_HOST=https://npmmirror.com/mirrors/playwright/
> playwright install chromium
> ```

---

## 五、运行脚本

1. 打开命令提示符
2. 用 cd 命令进入 app.py 所在的文件夹，比如：
   ```
   cd D:\tools
   ```
   或者如果在桌面：
   ```
   cd %USERPROFILE%\Desktop
   ```
3. 运行脚本：
   ```
   python app.py
   ```

---

## 六、使用流程

### 第一步：设置浏览器路径

程序启动后会问你浏览器路径，一般直接按回车用默认的就行。

### 第二步：设置白名单（可选）

程序会问你是否预设白名单：
- 输入 `y`：只处理白名单里的员工
- 输入 `n`：处理所有员工

如果选择设置白名单，把员工号粘贴进去，输入 `ok` 确认。

### 第三步：扫码登录

程序会自动打开浏览器并跳转到登录页面，你需要用手机扫码登录。登录成功后程序会自动跳转到录入页面。

### 第四步：选择模式

- 输入 `1`：批量模式（推荐，可以一次粘贴多条数据）
- 输入 `2`：手动模式（一条一条输入）
- 输入 `w`：设置/修改白名单
- 输入 `c`：清除白名单
- 输入 `q`：退出程序

### 第五步：粘贴数据

在批量模式下，直接从 Excel 复制数据粘贴进去，输入 `ok` 确认。

程序会解析数据并显示识别结果，确认无误后输入 `y` 开始自动填表。

---

## 七、数据格式说明

程序会自动从文本中识别以下信息：
- **员工号**：6位数字
- **姓名**：员工号后面的2-4个中文字符
- **请假类型**：需要包含完整的类型名称
- **日期**：支持 `2026-01-05` 或 `2026/1/5` 格式

### 支持的锁班类型

| 类型代码 | 名称 |
|---------|------|
| ALV | 年假（公休假） |
| ALV_FD | 飞行员公休（订座） |
| RECU_LVE | 健康疗养 |
| RECU_LVE_R | 康复疗养 |
| MAT_FA_LVE | 陪产假 |
| MAT_MO_LVE | 产假 |
| PREGNANT | 孕假 |
| PARENT_LVE | 探亲假-探父母 |
| SPOUSE_LVE | 探亲假-探配偶 |
| MARR_LVE | 婚假 |
| COMP_LVE | 丧假 |
| CHILD_LVE | 育儿假 |
| INJURY_LVE | 工伤假 |
| LWOP_LVE | 其他（事假） |
| UNPAID_LVE | 无薪 |
| HOUSE_LVE | 搬家 |
| BREED_LVE | 哺乳假 |
| PATERNITY | 独生子女护理假 |
| BIRC_LVE | 计划生育假 |
| REWARD_LVE | 奖励 |
| PENALTY | 停飞 |
| PRD_LVE | 经期假 |
| GRD | 地面班 |
| GDO | 地面休息 |
| TRNG1 | 训练 |
| BS_STUDY | 业务学习 |
| BUSINESS | 公务 |
| GRD_ONDUTY | 地面值班 |
| LG_STUDY | 语言学习/考试 |
| MEDL_CHK | 体检_临床 |
| MEDL_PHLE | 体检_抽血 |
| MEDL_EET | 体检_平板 |
| MEDL_PSYC | 体检_心理测试 |
| MTG | 会议 |
| MTG_SF | 安全讲评会 |
| DGET | 危险品培训 |
| EP | 飞行人员应急复训 |
| CRM | CRM培训 |
| T_SIM_INS | 模拟机检查 |
| T_SIM_REC | 模拟机复训 |
| T_SIM_INT | 模拟机初始 |
| T_SIM_UPG | 模拟机升级 |
| T_SIM_CON | 模拟机_转机型 |
| MAKEUP | 补考 |
| BS_CONCL | 飞行后讲评 |
| BS_CHK | 业务检查 |
| ADMN | 管理任务 |
| SOCIAL | 社会活动 |
| HANDBOOK | 手册 |
| POL_STUDY | 政治学习 |
| T/A | 部门活动 |

---

## 八、常见问题

### Q: 提示 "python 不是内部或外部命令"
A: Python 没有正确安装或没有添加到 PATH。重新安装 Python，记得勾选 "Add Python to PATH"。

### Q: pip install 下载很慢或报错
A: 使用清华镜像加速，在命令后面加 `-i https://pypi.tuna.tsinghua.edu.cn/simple`，比如：
```
pip install colorama -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### Q: playwright install chromium 下载很慢
A: 设置环境变量使用国内CDN后再安装：
```
set PLAYWRIGHT_DOWNLOAD_HOST=https://npmmirror.com/mirrors/playwright/
playwright install chromium
```

### Q: 登录后页面没有自动跳转
A: 手动进入"运行管理 - 非生产任务 - 非生产任务录入"页面，然后按回车继续。

### Q: 提交时显示有冲突
A: 程序会暂停并显示冲突信息，你可以选择：
- `s`：跳过这条，继续下一条
- `r`：重试当前这条
- `b`：返回主菜单

---

## 九、注意事项

1. 使用前请确保你有权限操作非生产任务录入
2. 建议先用少量数据测试，确认无误后再批量操作
3. 程序运行时不要关闭自动打开的浏览器窗口
4. 如果网络不稳定，可能需要多次重试
