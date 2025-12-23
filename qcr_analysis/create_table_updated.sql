-- 创建QCR_data表的SQL代码（更新版）
-- 数据库：MySQL
-- 修改：允许某些字段为NULL，以适应实际数据情况

-- 创建数据库（如果不存在）
CREATE DATABASE IF NOT EXISTS local_qcr CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;

-- 设置字符集
SET NAMES utf8mb4;
SET CHARACTER SET utf8mb4;

-- 使用数据库
USE local_qcr;

-- 如果表已存在，先删除
DROP TABLE IF EXISTS QCR_data;

-- 创建表（修改版：部分字段允许NULL）
CREATE TABLE QCR_data (
    date DATE NOT NULL COMMENT '日期',
    service_order_id BIGINT PRIMARY KEY COMMENT '服务单号，主键',
    order_id BIGINT NOT NULL COMMENT '订单号',
    issue_description VARCHAR(500) NULL DEFAULT '' COMMENT '问题描述',
    sku BIGINT NULL DEFAULT 0 COMMENT 'SKU编码',
    sn_code VARCHAR(100) NULL DEFAULT '' COMMENT 'SN编码，产品身份ID',
    customer_account VARCHAR(100) NULL DEFAULT '' COMMENT '客户账号',
    product_name VARCHAR(200) NULL DEFAULT '' COMMENT '商品名称',
    mtm VARCHAR(100) NULL DEFAULT '' COMMENT 'MTM编码',
    audit_reason VARCHAR(100) NULL DEFAULT '' COMMENT '审核原因',
    issue_category VARCHAR(100) NULL DEFAULT '' COMMENT '问题分类',
    category VARCHAR(100) NULL DEFAULT '' COMMENT '分类',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP COMMENT '记录创建时间',
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT '记录更新时间'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='质量投诉记录数据表';

-- 创建索引优化查询性能
CREATE INDEX idx_date ON QCR_data(date);
CREATE INDEX idx_order_id ON QCR_data(order_id);
CREATE INDEX idx_sn_code ON QCR_data(sn_code);
CREATE INDEX idx_audit_reason ON QCR_data(audit_reason);
CREATE INDEX idx_issue_category ON QCR_data(issue_category);

-- 显示表结构
SHOW CREATE TABLE QCR_data;
DESC QCR_data;

