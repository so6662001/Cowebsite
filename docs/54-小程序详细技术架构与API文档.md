# 钢贸调研小程序 详细技术架构 + API 文档

> 用途：docs/51 PRD 的技术落地  
> 受众：CTO + 开发团队  
> 开发周期：4-6 周（前端 2 + 后端 2 + 后台 PC + 测试）

---

# 第一部分：技术栈选型

## 1.1 前端
- **微信小程序原生**（更适合钢贸老板的使用习惯，性能 + 兼容性）
- 框架：原生（不用 Taro / uni-app，因为只在微信生态）
- UI 库：WeUI + 自定义组件
- 工具：微信开发者工具 + ESLint

## 1.2 后端
- **语言**：Node.js + TypeScript（团队熟悉度优先；也可以 Python + FastAPI）
- 框架：NestJS（适合中型业务）
- 部署：腾讯云 / 阿里云容器服务

## 1.3 数据库
- **主库**：MySQL 8.x（用户 / 测评 / 客户档案）
- **缓存**：Redis 6.x（会话 / 热数据 / 频次控制）
- **对象存储**：腾讯云 COS（PDF 报告 / 图片）

## 1.4 第三方
- 微信开放平台（OpenID / 手机号授权）
- 阿里云 / 腾讯云 短信
- WkHtmlToPdf 或 Puppeteer（PDF 生成）
- 飞书 / 你方 CRM API（客户分配）

## 1.5 部署架构

```
[微信用户]
    ↓ HTTPS
[微信小程序 CDN]
    ↓
[API 网关（腾讯云 / 阿里云）]
    ↓
[NestJS 后端服务 × 3 实例]
    ↓
├── [MySQL 主库 + 从库]
├── [Redis]
├── [COS 对象存储]
└── [短信 / CRM / PDF 服务]
```

---

# 第二部分：数据模型设计

## 2.1 核心表（10 张）

### 表 1：users 用户表
```sql
CREATE TABLE users (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  openid VARCHAR(64) UNIQUE NOT NULL COMMENT '微信 OpenID',
  unionid VARCHAR(64) COMMENT '微信 UnionID',
  phone VARCHAR(20) COMMENT '手机号',
  name VARCHAR(50) COMMENT '姓名',
  role ENUM('user', 'sales', 'admin') DEFAULT 'user',
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  INDEX idx_phone (phone),
  INDEX idx_openid (openid)
);
```

### 表 2：companies 公司表
```sql
CREATE TABLE companies (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  name VARCHAR(100) NOT NULL,
  business ENUM('trade', 'process_center', 'shearing', 'mixed') NOT NULL,
  scale ENUM('<3000w', '3000w-1y', '1-3y', '3-10y', '>10y') NOT NULL,
  region VARCHAR(50),
  contact_user_id BIGINT,
  segment ENUM('S', 'A', 'B', 'C') COMMENT '客户段',
  assigned_sales_id BIGINT COMMENT '分配的销售',
  status ENUM('new', 'contacted', 'demo', 'poc', 'signed', 'lost') DEFAULT 'new',
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY (contact_user_id) REFERENCES users(id),
  FOREIGN KEY (assigned_sales_id) REFERENCES users(id),
  INDEX idx_segment (segment),
  INDEX idx_status (status)
);
```

### 表 3：assessments 测评表
```sql
CREATE TABLE assessments (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  user_id BIGINT NOT NULL,
  company_id BIGINT NOT NULL,
  total_score INT,
  grade ENUM('A', 'B', 'C', 'D'),
  recommendation TEXT COMMENT '推荐方案 JSON',
  pdf_url VARCHAR(255),
  status ENUM('in_progress', 'completed') DEFAULT 'in_progress',
  started_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  completed_at TIMESTAMP,
  FOREIGN KEY (user_id) REFERENCES users(id),
  FOREIGN KEY (company_id) REFERENCES companies(id),
  INDEX idx_user (user_id),
  INDEX idx_company (company_id),
  INDEX idx_completed_at (completed_at)
);
```

### 表 4：assessment_answers 答题表
```sql
CREATE TABLE assessment_answers (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  assessment_id BIGINT NOT NULL,
  question_id INT NOT NULL COMMENT '题目 ID 1-30',
  module VARCHAR(30) NOT NULL COMMENT '7 大模块',
  selected_option INT NOT NULL,
  score INT NOT NULL,
  answered_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY (assessment_id) REFERENCES assessments(id),
  UNIQUE KEY (assessment_id, question_id)
);
```

### 表 5：assessment_module_scores 模块得分表
```sql
CREATE TABLE assessment_module_scores (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  assessment_id BIGINT NOT NULL,
  module VARCHAR(30) NOT NULL,
  score INT NOT NULL,
  max_score INT NOT NULL,
  percentage DECIMAL(5,2),
  FOREIGN KEY (assessment_id) REFERENCES assessments(id),
  INDEX idx_module (module)
);
```

### 表 6：sales 销售表
```sql
CREATE TABLE sales (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  user_id BIGINT NOT NULL,
  team ENUM('A', 'B', 'C') COMMENT '大客户/中端/电销',
  region VARCHAR(50) COMMENT '负责区域',
  max_clients INT DEFAULT 100,
  current_clients INT DEFAULT 0,
  is_active BOOLEAN DEFAULT true,
  FOREIGN KEY (user_id) REFERENCES users(id)
);
```

### 表 7：questions 题库表
```sql
CREATE TABLE questions (
  id INT PRIMARY KEY,
  module VARCHAR(30) NOT NULL,
  question_text TEXT NOT NULL,
  options JSON NOT NULL COMMENT '选项 + 分值',
  order_num INT NOT NULL,
  is_active BOOLEAN DEFAULT true
);
```

### 表 8：industry_reports 行业月报表
```sql
CREATE TABLE industry_reports (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  report_month CHAR(7) NOT NULL COMMENT 'YYYY-MM',
  region VARCHAR(50),
  total_assessments INT,
  avg_score INT,
  module_stats JSON,
  top_pains JSON,
  pdf_url VARCHAR(255),
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

### 表 9：follow_ups 销售跟进表
```sql
CREATE TABLE follow_ups (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  company_id BIGINT NOT NULL,
  sales_id BIGINT NOT NULL,
  status ENUM('assigned', 'contacted', 'demo', 'poc', 'signed', 'lost'),
  note TEXT,
  next_action_at TIMESTAMP,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY (company_id) REFERENCES companies(id),
  FOREIGN KEY (sales_id) REFERENCES sales(id)
);
```

### 表 10：appointments 预约表
```sql
CREATE TABLE appointments (
  id BIGINT PRIMARY KEY AUTO_INCREMENT,
  user_id BIGINT NOT NULL,
  company_id BIGINT NOT NULL,
  type ENUM('demo', 'consult', 'visit'),
  preferred_time DATETIME,
  status ENUM('pending', 'confirmed', 'completed', 'cancelled'),
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

---

# 第三部分：核心 API 文档（25 个）

## API 通用规范
- BaseURL：`https://api.steeltrade-survey.com/v1`
- 认证：JWT Bearer Token
- 格式：JSON
- 编码：UTF-8
- 错误码：标准 HTTP + 业务码

### 通用响应格式
```json
{
  "code": 0,           // 0 = 成功，其他 = 错误码
  "message": "success",
  "data": { ... }
}
```

---

## 3.1 用户授权与登录

### API 1：微信登录
```
POST /auth/wechat-login
```
**请求**：
```json
{
  "code": "wechat_code_from_wx.login",
  "userInfo": { "nickName": "...", "avatarUrl": "..." }
}
```
**响应**：
```json
{
  "code": 0,
  "data": {
    "token": "JWT_TOKEN",
    "user": {
      "id": 123,
      "openid": "...",
      "name": "...",
      "role": "user"
    },
    "isNewUser": true
  }
}
```

### API 2：手机号授权
```
POST /auth/phone-bind
```
**请求**：
```json
{
  "encryptedData": "...",
  "iv": "...",
  "sessionKey": "..."
}
```
**响应**：
```json
{
  "code": 0,
  "data": {
    "phone": "13800001111"
  }
}
```

---

## 3.2 测评 API

### API 3：开始测评
```
POST /assessments/start
Authorization: Bearer <token>
```
**请求**：
```json
{
  "companyName": "XX 钢贸",
  "business": "trade",
  "scale": "1-3y",
  "name": "张总",
  "phone": "13800001111"
}
```
**响应**：
```json
{
  "code": 0,
  "data": {
    "assessmentId": 456,
    "companyId": 789,
    "questions": [
      {
        "id": 1,
        "module": "利润成本",
        "question": "您今年到现在的毛利率大概是多少？",
        "options": [
          { "id": 1, "text": "≥ 5%", "score": 5 },
          { "id": 2, "text": "3-5%", "score": 3 },
          { "id": 3, "text": "1-3%", "score": 2 },
          { "id": 4, "text": "< 1%", "score": 1 }
        ]
      },
      ...
    ]
  }
}
```

### API 4：提交答案
```
POST /assessments/{id}/answer
```
**请求**：
```json
{
  "questionId": 1,
  "selectedOption": 2,
  "score": 3
}
```
**响应**：
```json
{
  "code": 0,
  "data": {
    "progress": 0.033,
    "currentScore": 3
  }
}
```

### API 5：完成测评
```
POST /assessments/{id}/complete
```
**响应**：
```json
{
  "code": 0,
  "data": {
    "totalScore": 95,
    "grade": "B",
    "moduleScores": {
      "利润成本": { "score": 18, "max": 25, "pct": 0.72 },
      "应收账款": { "score": 12, "max": 20, "pct": 0.60 },
      ...
    },
    "topPains": [
      "应收账款风险高，推 AI 风控 + 应收预警",
      "客户流失，推 CRM + AI 客户管理",
      "销售跑单，推 CRM + AI 销售助理"
    ],
    "recommendation": "经营版 ERP + AI 经营包 + 货袋子 + 风控咨询，¥15-40 万 / 年",
    "pdfUrl": "https://cos.../report-456.pdf"
  }
}
```

### API 6：获取我的报告列表
```
GET /assessments/mine
```
**响应**：
```json
{
  "code": 0,
  "data": {
    "list": [
      {
        "id": 456,
        "completedAt": "2026-06-15 10:30:00",
        "score": 95,
        "grade": "B",
        "pdfUrl": "..."
      }
    ]
  }
}
```

### API 7：获取报告详情
```
GET /assessments/{id}
```
**响应**：完整测评数据 + PDF URL

---

## 3.3 行业月报 API

### API 8：获取行业月报列表
```
GET /reports/industry?region=长三角
```
**响应**：
```json
{
  "code": 0,
  "data": {
    "list": [
      {
        "id": 1,
        "month": "2026-06",
        "region": "长三角",
        "totalAssessments": 856,
        "avgScore": 89,
        "pdfUrl": "..."
      }
    ]
  }
}
```

### API 9：获取行业月报详情
```
GET /reports/industry/{id}
```

### API 10：行业对标（个人 vs 行业）
```
GET /assessments/{id}/benchmark
```
**响应**：
```json
{
  "code": 0,
  "data": {
    "myScore": 95,
    "industryAvg": 89,
    "sameSegmentAvg": 87,
    "sameRegionAvg": 91,
    "ranking": "前 35%"
  }
}
```

---

## 3.4 预约 API

### API 11：发起预约
```
POST /appointments
```
**请求**：
```json
{
  "type": "demo",
  "preferredTime": "2026-06-20 14:00:00",
  "note": "希望专家上门"
}
```

### API 12：我的预约列表
```
GET /appointments/mine
```

---

## 3.5 销售端 API（后台 PC）

### API 13：获取我的客户列表
```
GET /sales/clients
Authorization: Bearer <sales-token>
```
**响应**：
```json
{
  "code": 0,
  "data": {
    "list": [
      {
        "id": 789,
        "name": "XX 钢贸",
        "segment": "A",
        "lastScore": 95,
        "grade": "B",
        "status": "contacted",
        "lastFollowUp": "2026-06-15"
      }
    ],
    "total": 50
  }
}
```

### API 14：客户详情
```
GET /sales/clients/{id}
```

### API 15：更新跟进状态
```
POST /sales/clients/{id}/follow-up
```
**请求**：
```json
{
  "status": "demo",
  "note": "客户对 AI 风控感兴趣",
  "nextActionAt": "2026-06-22 10:00:00"
}
```

### API 16：客户测评历史
```
GET /sales/clients/{id}/assessments
```
返回客户多次测评对比

---

## 3.6 管理后台 API（PC 端）

### API 17：管理后台 - 数据看板
```
GET /admin/dashboard
```
**响应**：
```json
{
  "code": 0,
  "data": {
    "todayAssessments": 45,
    "totalAssessments": 12450,
    "completionRate": 0.62,
    "gradeDistribution": {
      "A": 12,
      "B": 35,
      "C": 38,
      "D": 15
    },
    "salesPerformance": [
      { "salesId": 1, "name": "A1", "newClients": 18, "responseRate": 0.95 }
    ]
  }
}
```

### API 18：客户管理
```
GET /admin/clients?segment=A&region=长三角
POST /admin/clients/{id}/assign
```

### API 19：题库管理
```
GET /admin/questions
POST /admin/questions
PUT /admin/questions/{id}
DELETE /admin/questions/{id}
```

### API 20：销售分配规则
```
GET /admin/assignment-rules
PUT /admin/assignment-rules
```

### API 21：自动分配触发
```
POST /admin/auto-assign
```

---

## 3.7 报告生成 API

### API 22：生成客户专属 PDF
```
POST /reports/generate
```
**请求**：
```json
{
  "assessmentId": 456,
  "template": "client_report_v2"
}
```
**响应**：
```json
{
  "code": 0,
  "data": {
    "pdfUrl": "...",
    "expiresAt": "2026-07-15"
  }
}
```

### API 23：生成行业月报（管理后台）
```
POST /admin/reports/generate-industry
```

---

## 3.8 通知与短信

### API 24：发送通知
```
POST /notifications/send
```
**请求**：
```json
{
  "userId": 123,
  "type": "sms",  // sms / wechat
  "templateId": "assessment_complete",
  "params": { "name": "张总", "score": 95 }
}
```

### API 25：订阅消息授权
```
POST /notifications/subscribe
```

---

# 第四部分：业务流程图（关键 5 个）

## 4.1 用户测评完整流程

```
[小程序入口]
    ↓
[微信授权登录] API 1
    ↓
[填基础信息] API 3
    ↓
[逐题答题（30 题）] API 4 × 30
    ↓
[完成测评 → 评分计算] API 5
    ↓
[展示结果 + 雷达图 + 推荐]
    ↓
[后台异步任务]
   ├── 生成 PDF 报告 API 22
   ├── 短信通知用户
   ├── 客户档案入 CRM
   └── 自动分配销售 API 21
    ↓
[销售收到企业微信通知]
    ↓
[销售加客户微信 + 跟进]
```

## 4.2 自动销售分配规则（流程图）

```
[新客户完成测评]
    ↓
读取 company 基本信息
    ↓
判断 segment（S/A/B/C）
    ↓
┌──────────────┬──────────────┐
│ S 段         │ A 段         │ B 段         │ C 段
│ 大客户铁三角 │ 大客户铁三角 │ 中端铁二角 │ 电销批量
└──────────────┴──────────────┘
    ↓
匹配 region（长三角/珠三角/京津冀/...）
    ↓
匹配 sales（当前客户数 < max_clients）
    ↓
分配 → 创建 follow_ups 记录
    ↓
发送通知给 sales（企业微信）
```

## 4.3 PDF 生成流程

```
[完成测评]
    ↓
后端异步任务（队列）
    ↓
渲染 HTML 模板（含数据）
    ↓
WkHtmlToPdf 转 PDF
    ↓
上传 COS
    ↓
更新 assessment.pdf_url
    ↓
推送给用户（小程序 + 短信）
```

## 4.4 行业月报生成

```
[每月 1 日 00:00 定时任务]
    ↓
聚合上月所有 assessments
    ↓
计算行业平均 + 模块分布 + TOP 痛点
    ↓
生成 PDF 报告
    ↓
推送给所有用户 + 销售
```

## 4.5 防作弊机制

```
[用户填问卷]
    ↓
检查 1：openid 1 小时内不能重复测评
    ↓
检查 2：手机号去重
    ↓
检查 3：IP / 设备指纹
    ↓
检查 4：答题速度（< 30 秒 = 异常）
    ↓
检查 5：选项分布（全选 1 个 = 异常）
    ↓
异常 → 标记 + 不分配销售
```

---

# 第五部分：性能 / 安全 / 监控

## 5.1 性能目标

| 指标 | 目标 |
|---|---|
| API 响应时间 | P50 < 200ms / P99 < 1s |
| 小程序首屏加载 | < 2s |
| 测评 PDF 生成 | < 5s |
| 数据库 QPS | ≥ 1000 |
| 同时在线用户 | ≥ 500 |

## 5.2 安全设计

- **API 鉴权**：JWT + 过期时间 7 天
- **接口频次**：单 IP 100 次/分钟（防爬）
- **敏感数据加密**：手机号 / 公司数据 AES-256
- **HTTPS 强制**：全站 HTTPS
- **SQL 注入防护**：ORM 参数化查询
- **XSS 防护**：输入过滤 + 输出转义
- **CSRF 防护**：Token 验证

## 5.3 数据合规

- 用户协议 + 隐私政策（必须有，由法务出）
- 用户可随时**删除自己的数据**（GDPR 风格）
- 数据**不卖第三方**（明示）
- 用户授权同意才能用作行业研究（脱敏）

## 5.4 监控告警

- **APM**：阿里云 ARMS / 腾讯云 APM
- **日志**：ELK Stack
- **告警**：
  - API 错误率 > 1% → 企业微信通知
  - DB 连接池满 → 紧急告警
  - PDF 生成失败 > 5% → 通知

---

# 第六部分：开发计划（4-6 周）

## Week 1：需求 + 设计
- D1-D2：需求确认 + 数据库设计
- D3-D5：视觉设计稿（设计师）
- 输出：PRD 终版 + 表结构 + 设计稿

## Week 2：前端基础
- D6-D7：小程序框架 + 路由 + 鉴权
- D8-D10：首页 + 测评流程页面
- 输出：可跑通测评流程（无数据）

## Week 3：后端 + 集成
- D11-D13：后端 API 核心（用户 / 测评 / 报告）
- D14-D15：前后端联调
- 输出：完整测评流程跑通

## Week 4：报告 + 自动化
- D16-D17：PDF 生成 + 短信 / 通知
- D18-D19：自动分配规则 + CRM 集成
- D20：销售端基础页面（管理后台）

## Week 5：管理后台
- D21-D23：完整管理后台
- D24-D25：数据看板 + 报告生成

## Week 6：测试 + 上线
- D26-D27：全流程测试 + Bug 修复
- D28-D29：内测（10 个真实客户）
- D30：正式上线

---

# 第七部分：开发团队配置

| 角色 | 人数 | 工时 |
|---|---|---|
| 产品经理 | 1 | 全期 |
| 小程序前端 | 2 | W2-W6 |
| 后端工程师 | 2 | W2-W6 |
| PC 前端 | 1 | W4-W5 |
| 视觉设计 | 1 | W1-W3 |
| 测试 | 1 | W5-W6 |
| DevOps | 0.5 | W4-W6 |

**总人日 ≈ 100-120 人日**

---

# 第八部分：上线后的「灰度发布」

## 8.1 3 阶段灰度

### 阶段 1：内部 + 种子（W6 末 - W7）
- 9 销售内部测试
- 邀请 10 个种子客户

### 阶段 2：100 个客户（W8）
- 通过销售扫码 + 朋友圈
- 关注 Bug + 体验问题

### 阶段 3：全员开放（W9 起）
- 公众号 + 行业群推广
- 货袋子平台首页推荐

---

# 第九部分：上线后的「持续迭代」

## M1（上线第 1 月）
- Bug 修复 + UX 优化
- 收集 ≥ 50 条用户反馈

## M2-M3
- v1.1 新功能：行业对标 + 复访
- v1.2 新功能：邀请同行测评（病毒传播）

## M6+
- v2.0：与 CRM 深度集成 + AI 自动推荐套餐
- v3.0：与货袋子平台打通 → 跨平台数据

---

# 5 句话给老板

1. **6 周 / ¥8 万 + 100 人日** 可以做出完整版本
2. **核心 API 25 个**——足够完整业务闭环
3. **必须有自动销售分配** + **PDF 自动生成**
4. **安全 + 合规 + 数据隐私** 不能省
5. **3 阶段灰度发布** + **持续迭代** 保证质量

---

# 附录 A：技术栈替代方案

| 模块 | 主选 | 备选 |
|---|---|---|
| 后端语言 | Node.js + TS | Python + FastAPI / Go |
| 后端框架 | NestJS | Express / Koa / FastAPI |
| 数据库 | MySQL 8 | PostgreSQL |
| 缓存 | Redis | Memcached |
| 对象存储 | 腾讯云 COS | 阿里云 OSS |
| PDF | Puppeteer | WkHtmlToPdf |
| APM | 阿里云 ARMS | Sentry + New Relic |

---

# 附录 B：开发预算（自建 vs 外包）

## 自建（强烈推荐）
- 开发：¥0（团队成本已含）
- 服务器 + 工具：¥1-2 万 / 年
- 设计：¥2-3 万
- **合计 ¥5 万**

## 外包
- 开发：¥8-15 万
- 服务器：¥1-2 万 / 年
- 维护：¥3-5 万 / 年
- **合计 ¥15-20 万 / 年**

**建议**：自建（你方有工程团队 + 长期资产）
