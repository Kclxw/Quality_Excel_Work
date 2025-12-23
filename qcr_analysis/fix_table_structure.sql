-- 一键修复表结构脚本
-- 将现有表的字段修改为允许NULL，解决数据导入问题
-- 使用方法：mysql -u root -p local_qcr < fix_table_structure.sql

USE local_qcr;

-- 显示当前表结构
SELECT '==== 修改前的表结构 ====' AS info;
DESC QCR_data;

-- 修改字段为允许NULL（保留现有数据）
SELECT '==== 开始修改表结构 ====' AS info;

ALTER TABLE QCR_data MODIFY COLUMN issue_description VARCHAR(500) NULL DEFAULT '' COMMENT '问题描述';
SELECT '✓ 修改 issue_description 完成' AS status;

ALTER TABLE QCR_data MODIFY COLUMN sku BIGINT NULL DEFAULT 0 COMMENT 'SKU编码';
SELECT '✓ 修改 sku 完成' AS status;

ALTER TABLE QCR_data MODIFY COLUMN sn_code VARCHAR(100) NULL DEFAULT '' COMMENT 'SN编码，产品身份ID';
SELECT '✓ 修改 sn_code 完成' AS status;

ALTER TABLE QCR_data MODIFY COLUMN customer_account VARCHAR(100) NULL DEFAULT '' COMMENT '客户账号';
SELECT '✓ 修改 customer_account 完成' AS status;

ALTER TABLE QCR_data MODIFY COLUMN product_name VARCHAR(200) NULL DEFAULT '' COMMENT '商品名称';
SELECT '✓ 修改 product_name 完成' AS status;

ALTER TABLE QCR_data MODIFY COLUMN mtm VARCHAR(100) NULL DEFAULT '' COMMENT 'MTM编码';
SELECT '✓ 修改 mtm 完成' AS status;

ALTER TABLE QCR_data MODIFY COLUMN audit_reason VARCHAR(100) NULL DEFAULT '' COMMENT '审核原因';
SELECT '✓ 修改 audit_reason 完成' AS status;

ALTER TABLE QCR_data MODIFY COLUMN issue_category VARCHAR(100) NULL DEFAULT '' COMMENT '问题分类';
SELECT '✓ 修改 issue_category 完成' AS status;

ALTER TABLE QCR_data MODIFY COLUMN category VARCHAR(100) NULL DEFAULT '' COMMENT '分类';
SELECT '✓ 修改 category 完成' AS status;

-- 显示修改后的表结构
SELECT '==== 修改后的表结构 ====' AS info;
DESC QCR_data;

-- 显示记录数量
SELECT '==== 当前记录统计 ====' AS info;
SELECT COUNT(*) AS total_records FROM QCR_data;

SELECT '==== 修复完成！====' AS info;

