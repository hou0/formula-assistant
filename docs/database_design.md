# 化妆品备案安全评估工具 - 数据库设计文档

## 1. 数据库概述

本系统使用 **SQLite** 作为数据库引擎，数据存储在 `userdata/safety.db` 文件中。数据库包含以下类型的表：

| 表类别 | 说明 |
|--------|------|
| **业务表** | 存储用户业务数据（原料库、配方等） |
| **法规表** | 存储《化妆品安全技术规范》相关数据 |
| **评估表** | 存储安全评估相关数据 |

---

## 2. 数据库表结构

### 2.1 业务数据表

#### 2.1.1 raw_materials（原料基本信息表）

| 字段名 | 类型 | 约束 | 说明 |
|--------|------|------|------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | 主键ID |
| name_zh | TEXT | NOT NULL | 原料中文名称 |
| name_inci | TEXT | | INCI名称 |
| composition | TEXT | NOT NULL | 成分组成（复合原料用） |
| default_purpose | TEXT | | 默认使用目的 |
| supplier_code | TEXT | | 供应商代码 |
| remarks | TEXT | | 备注 |
| trade_name | TEXT | | 商品名 |
| internal_code | TEXT | | 企业内部编码 |
| supplier_name | TEXT | | 供应商名称 |
| contact_person | TEXT | | 联系人 |
| contact_phone | TEXT | | 联系电话 |
| contact_email | TEXT | | 联系邮箱 |

#### 2.1.2 formula（配方表）

| 字段名 | 类型 | 约束 | 说明 |
|--------|------|------|------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | 主键ID |
| raw_material_id | INTEGER | NOT NULL, FOREIGN KEY | 关联原料ID |
| added_percent | REAL | NOT NULL | 添加量（百分比） |
| sort_order | INTEGER | NOT NULL | 排序顺序 |
| is_new_material | INTEGER | DEFAULT 0 | 是否新原料（0/1） |
| registration_number | TEXT | DEFAULT '' | 注册号/备案号 |
| purpose_override | TEXT | DEFAULT '' | 使用目的覆盖 |
| remarks | TEXT | DEFAULT '' | 备注 |

#### 2.1.3 custom_columns（自定义列配置表）

| 字段名 | 类型 | 约束 | 说明 |
|--------|------|------|------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | 主键ID |
| name | TEXT | NOT NULL UNIQUE | 列名 |
| width | INTEGER | DEFAULT 120 | 列宽（像素） |
| sort_order | INTEGER | DEFAULT 0 | 排序顺序 |

---

### 2.2 法规数据表

#### 2.2.1 regulation_banned_chemical（禁用化学组分表）

| 字段名 | 类型 | 约束 | 说明 |
|--------|------|------|------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | 主键ID |
| seq | TEXT | | 技术规范表序号 |
| name_zh | TEXT | NOT NULL | 中文名称 |
| name_en | TEXT | | 英文名称 |

#### 2.2.2 regulation_banned_botanical（禁用动植物组分表）

| 字段名 | 类型 | 约束 | 说明 |
|--------|------|------|------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | 主键ID |
| seq | TEXT | | 技术规范表序号 |
| name_zh | TEXT | NOT NULL | 中文名称 |
| latin_name | TEXT | | 拉丁学名 |
| notes | TEXT | | 备注 |

#### 2.2.3 regulation_restricted（限用组分表）

| 字段名 | 类型 | 约束 | 说明 |
|--------|------|------|------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | 主键ID |
| seq | TEXT | | 技术规范表序号 |
| name_zh | TEXT | NOT NULL | 中文名称 |
| name_en | TEXT | | 英文名称 |
| inci_name | TEXT | | INCI名称 |
| scope_of_use | TEXT | | 使用范围 |
| max_concentration | TEXT | | 最大浓度 |
| restrictions | TEXT | | 限制要求 |
| label_requirements | TEXT | | 标签标注要求 |

#### 2.2.4 regulation_allowed_preservative（准用防腐剂表）

| 字段名 | 类型 | 约束 | 说明 |
|--------|------|------|------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | 主键ID |
| seq | TEXT | | 技术规范表序号 |
| name_zh | TEXT | NOT NULL | 中文名称 |
| name_en | TEXT | | 英文名称 |
| inci_name | TEXT | | INCI名称 |
| max_concentration | TEXT | | 最大浓度 |
| scope_of_use | TEXT | | 使用范围与条件 |
| label_requirements | TEXT | | 标签标注要求 |

#### 2.2.5 regulation_allowed_sunscreen（准用防晒剂表）

| 字段名 | 类型 | 约束 | 说明 |
|--------|------|------|------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | 主键ID |
| seq | TEXT | | 技术规范表序号 |
| name_zh | TEXT | NOT NULL | 中文名称 |
| name_en | TEXT | | 英文名称 |
| inci_name | TEXT | | INCI名称 |
| max_concentration | TEXT | | 最大浓度 |
| restrictions | TEXT | | 限制要求 |
| label_requirements | TEXT | | 标签标注要求 |

#### 2.2.6 regulation_allowed_colorant（准用着色剂表）

| 字段名 | 类型 | 约束 | 说明 |
|--------|------|------|------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | 主键ID |
| seq | TEXT | | 技术规范表序号 |
| ci_number | TEXT | | CI号 |
| color | TEXT | | 颜色 |
| name_zh | TEXT | NOT NULL | 中文名称 |
| name_en | TEXT | | 英文名称 |
| usage_categories | TEXT | | 使用范围（4类标记） |
| restrictions | TEXT | | 限制要求 |

#### 2.2.7 regulation_allowed_hair_dye（准用染发剂表）

| 字段名 | 类型 | 约束 | 说明 |
|--------|------|------|------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | 主键ID |
| seq | TEXT | | 技术规范表序号 |
| name_zh | TEXT | NOT NULL | 中文名称 |
| name_en | TEXT | | 英文名称 |
| inci_name | TEXT | | INCI名称 |
| type | TEXT | | 氧化型/非氧化型 |
| max_concentration | TEXT | | 最大浓度 |
| restrictions | TEXT | | 限制要求 |

---

### 2.3 评估数据表

#### 2.3.1 safety_usage_market（已上市产品原料使用信息表）

| 字段名 | 类型 | 约束 | 说明 |
|--------|------|------|------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | 主键ID |
| seq | TEXT | | 目录序号 |
| name_zh | TEXT | NOT NULL | 原料中文名称 |
| name_inci | TEXT | | INCI名称 |
| application_site | TEXT | | 作用部位 |
| usage_method | TEXT | | 使用方法（驻留/淋洗） |
| concentration | TEXT | | 浓度范围 |
| remarks | TEXT | | 备注 |

#### 2.3.2 safety_usage_international（国际化妆品安全评估数据索引表）

| 字段名 | 类型 | 约束 | 说明 |
|--------|------|------|------|
| id | INTEGER | PRIMARY KEY AUTOINCREMENT | 主键ID |
| seq | TEXT | | 索引序号 |
| name_zh | TEXT | NOT NULL | 原料中文名称 |
| name_inci | TEXT | | INCI名称 |
| authority | TEXT | | 评估机构（CIR/SCCS等） |
| conclusion | TEXT | | 评估结论 |
| limit_concentration | TEXT | | 限制浓度 |
| restrictions | TEXT | | 限制条件 |

---

## 3. JSON配置文件

除SQLite数据库外，系统还使用JSON文件存储配置数据：

### 3.1 risk_substances.json（风险物质规则配置）

```json
{
  "risk_substances": [
    {
      "name": "二噁烷",
      "trigger_keywords": ["PEG", "聚乙二醇", "乙氧基化"],
      "trigger_all": false,
      "limit": "≤10mg/kg（2027年实施）",
      "reference": "《化妆品安全技术规范》",
      "description": "..."
    }
  ]
}
```

### 3.2 toxicology_seed.json（毒理学种子数据）

```json
{
  "ingredients": [
    {
      "name_zh": "甘油",
      "name_inci": "GLYCERIN",
      "noael": 1000,
      "ttc_class": "I",
      "cir_conclusion": "..."
    }
  ]
}
```

### 3.3 ingredient_catalog.json（目录I数据）

存储《已使用化妆品原料目录》I的原料数据。

### 3.4 ingredient_catalog_ii.json（目录II数据）

存储《已使用化妆品原料目录》II的原料数据。

---

## 4. 数据导入说明

### 4.1 法规数据导入

法规数据表（regulation_*）通过专门的导入脚本从官方CSV文件导入：

| 表名 | 数据来源 |
|------|----------|
| regulation_banned_chemical | 《化妆品安全技术规范》表1 |
| regulation_banned_botanical | 《化妆品安全技术规范》表2 |
| regulation_restricted | 《化妆品安全技术规范》表3 |
| regulation_allowed_preservative | 《化妆品安全技术规范》表4 |
| regulation_allowed_sunscreen | 《化妆品安全技术规范》表5 |
| regulation_allowed_colorant | 《化妆品安全技术规范》表6 |
| regulation_allowed_hair_dye | 《化妆品安全技术规范》表7 |

### 4.2 使用量数据导入

| 表名 | 数据来源 |
|------|----------|
| safety_usage_market | 《已上市产品原料使用信息》（中检院2025年2月版） |
| safety_usage_international | 《国际化妆品安全评估数据索引》 |

---

## 5. 数据库维护

### 5.1 数据备份

建议定期备份 `userdata/` 目录，包含：
- `safety.db` - 数据库文件
- `*.json` - 配置文件

### 5.2 数据更新

- **法规数据**：当《化妆品安全技术规范》更新时，需重新导入法规表
- **目录数据**：通过「同步目录」功能定期更新
- **使用量数据**：每年更新一次（根据中检院发布）

---

## 6. 数据库连接信息

```python
# 数据库路径
DB_PATH = "userdata/safety.db"

# 连接方式
conn = sqlite3.connect(DB_PATH)
conn.row_factory = sqlite3.Row  # 返回字典形式的行
```

---

**版本**: v1.0  
**创建日期**: 2026-06-08  
**最后更新**: 2026-06-08
