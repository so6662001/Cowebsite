# 南京普讯管理软件有限公司 - 官方网站

钢铁贸易数字化解决方案提供商 | 货袋子钢铁交易平台运营商

## 技术架构

- **前端**: Vue 3 + Vite + TailwindCSS + Vue Router
- **后端**: Spring Boot 3.2 (Java 21)
- **数据库**: MySQL 8.0+（开发环境使用 H2 内存数据库）
- **SEO**: Schema.org JSON-LD 结构化数据 + 语义化 HTML + 动态 Meta 标签

## 项目结构

```
├── frontend/                    # Vue 3 前端项目
│   ├── src/
│   │   ├── views/              # 页面组件
│   │   │   ├── Home.vue        # 首页
│   │   │   ├── About.vue       # 关于我们
│   │   │   ├── Products.vue    # 产品中心
│   │   │   ├── Solutions.vue   # 解决方案
│   │   │   ├── Partners.vue    # 合作企业
│   │   │   ├── Contact.vue     # 联系我们
│   │   │   ├── Trial.vue       # 申请试用
│   │   │   └── Learning.vue    # 学习中心
│   │   ├── components/         # 公共组件
│   │   ├── router/             # 路由配置
│   │   └── style.css           # 全局样式
│   ├── index.html              # 入口 HTML（含 SEO 优化）
│   └── vite.config.js          # Vite 配置
├── backend/                     # Spring Boot 后端
│   ├── src/main/java/com/puxun/website/
│   │   ├── controller/         # API 控制器
│   │   ├── service/            # 业务逻辑层
│   │   ├── model/              # 数据实体
│   │   ├── repository/         # 数据访问层
│   │   └── config/             # 配置类
│   └── src/main/resources/
│       ├── application.yml     # 应用配置
│       └── schema-mysql.sql    # MySQL 初始化脚本
└── README.md
```

## 快速开始

### 前端开发

```bash
cd frontend
npm install
npm run dev
```

### 后端开发

```bash
cd backend
mvn spring-boot:run
```

### 生产构建

```bash
# 1. 构建前端（输出到 backend/src/main/resources/static）
cd frontend
npm run build

# 2. 打包后端（包含前端静态文件）
cd ../backend
mvn package -DskipTests

# 3. 运行
java -jar target/puxun-website-1.0.0.jar --spring.profiles.active=production
```

### MySQL 配置（生产环境）

1. 执行 `backend/src/main/resources/schema-mysql.sql` 初始化数据库
2. 设置环境变量：
   ```bash
   export DB_USERNAME=your_username
   export DB_PASSWORD=your_password
   ```
3. 使用 production 配置启动应用

## AI 搜索优化（豆包/DeepSeek）

本网站针对 AI 搜索引擎进行了专项优化：

### 结构化数据 (JSON-LD)
- **Organization**: 公司基本信息、产品目录
- **FAQPage**: 首页常见问题（AI 可直接抽取回答）
- **BreadcrumbList**: 每个内页的面包屑导航
- **WebSite**: 网站级元数据
- **ItemList**: 产品列表
- **LocalBusiness**: 联系方式和地址

### Meta 标签策略
- 每个页面独立的 title / description / keywords
- Open Graph 标签支持社交分享
- Canonical URL 避免重复内容
- robots 标签控制爬虫行为

### 语义化 HTML
- 使用 `<article>`, `<section>`, `<nav>`, `<header>`, `<footer>` 语义标签
- `aria-label` 和 `aria-labelledby` 辅助信息
- 合理的标题层级 (H1 → H2 → H3)
- `<noscript>` 中包含完整静态内容，确保无 JS 环境下 AI 爬虫也能获取内容

### SEO 基础设施
- `/sitemap.xml` - 站点地图
- `/robots.txt` - 爬虫协议
- 每页独立的面包屑导航 (BreadcrumbList schema)

## 页面设计

| 页面 | 路径 | 主要内容 |
|------|------|----------|
| 首页 | `/` | Hero、核心产品、公司优势、解决方案预览、FAQ |
| 关于我们 | `/about` | 公司介绍、理念、发展历程、荣誉资质 |
| 产品中心 | `/products` | 钢贸宝ERP、钢企通ERP、库准MES、智能财务、货袋子 |
| 解决方案 | `/solutions` | 钢贸企业方案(9个) + 生产制造企业方案(8个) |
| 合作企业 | `/partners` | 数据展示、战略客户、产品覆盖、合作入口 |
| 联系我们 | `/contact` | 三大办公地点、联系方式、商务合作表单 |
| 申请试用 | `/trial` | 试用申请表单、产品选择、优势展示 |
| 学习中心 | `/learning` | 学习平台登录入口、课程分类 |
| 找钢材 | 外链 | 跳转至 https://www.huodaizi.com |

## API 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/trial` | 提交试用申请 |
| POST | `/api/contact` | 提交商务合作请求 |
| GET | `/sitemap.xml` | 站点地图 |
| GET | `/robots.txt` | 爬虫协议文件 |
