# Cowebsite - 专业数字营销与智能建站官网

面向 AI 搜索引擎（豆包、DeepSeek 等）深度优化的企业官网，采用语义化 HTML5 + Schema.org 结构化数据 + 全面 Meta 标签体系。

## 项目结构

```
├── index.html          # 首页
├── about.html          # 关于我们
├── services.html       # 服务项目
├── contact.html        # 联系我们
├── css/
│   └── style.css       # 全局样式
├── js/
│   └── main.js         # 交互脚本
├── robots.txt          # 爬虫规则
├── sitemap.xml         # 站点地图
└── images/             # 图片资源（待添加）
```

## 运行方式

纯静态网站，无需构建工具。任选以下方式在本地预览：

```bash
# 方式一：Python
python3 -m http.server 8080

# 方式二：Node.js（需安装 http-server）
npx http-server -p 8080

# 方式三：VS Code Live Server 插件
# 右键 index.html → Open with Live Server
```

然后在浏览器访问 `http://localhost:8080`。

## AI 搜索引擎优化要点

- **Schema.org JSON-LD**：Organization、WebSite、WebPage、FAQPage、Service、BreadcrumbList、LocalBusiness 等多种结构化数据
- **语义化 HTML5**：header、nav、main、section、article、footer 等标签
- **Meta 标签体系**：title、description、keywords、Open Graph、Twitter Cards
- **robots.txt**：显式允许豆包（Bytespider）、DeepSeek、百度、Google 等爬虫
- **sitemap.xml**：完整的站点地图
- **FAQ 结构化数据**：方便 AI 搜索引擎直接提取问答内容
- **面包屑导航**：清晰的页面层级关系
- **无障碍访问**：ARIA 标签、role 属性、键盘导航支持
