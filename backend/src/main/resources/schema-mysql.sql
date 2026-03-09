-- 南京普讯管理软件有限公司官网 数据库初始化脚本
-- 数据库: MySQL 8.0+
-- 字符集: utf8mb4

CREATE DATABASE IF NOT EXISTS puxun_website
  DEFAULT CHARACTER SET utf8mb4
  DEFAULT COLLATE utf8mb4_unicode_ci;

USE puxun_website;

-- 试用申请表
CREATE TABLE IF NOT EXISTS trial_applications (
  id BIGINT AUTO_INCREMENT PRIMARY KEY,
  contact_name VARCHAR(100) NOT NULL COMMENT '联系人姓名',
  phone VARCHAR(20) NOT NULL COMMENT '手机号',
  company_name VARCHAR(200) NOT NULL COMMENT '公司名称',
  company_type VARCHAR(50) COMMENT '企业类型: steel_trade/steel_manufacture/other',
  company_scale VARCHAR(20) COMMENT '企业规模',
  products VARCHAR(500) NOT NULL COMMENT '感兴趣的产品，逗号分隔',
  remark TEXT COMMENT '补充说明',
  status VARCHAR(20) NOT NULL DEFAULT 'pending' COMMENT '状态: pending/contacted/converted/closed',
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
  updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '更新时间',
  INDEX idx_status (status),
  INDEX idx_phone (phone),
  INDEX idx_created_at (created_at)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='试用申请';

-- 商务合作/联系我们表
CREATE TABLE IF NOT EXISTS contact_requests (
  id BIGINT AUTO_INCREMENT PRIMARY KEY,
  name VARCHAR(100) NOT NULL COMMENT '姓名',
  phone VARCHAR(20) NOT NULL COMMENT '电话',
  company VARCHAR(200) NOT NULL COMMENT '公司名称',
  message TEXT COMMENT '合作意向描述',
  status VARCHAR(20) NOT NULL DEFAULT 'pending' COMMENT '状态: pending/replied/closed',
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '创建时间',
  INDEX idx_status (status),
  INDEX idx_created_at (created_at)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='商务合作请求';
