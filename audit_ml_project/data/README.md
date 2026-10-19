# 数据说明

## 文件

| 项目 | 内容 |
|---|---|
| 文件 | `audit_raw.csv` |
| 来源 | Kaggle 数据集 “Financial Audit Transactions Dataset”（课程指定） |
| 规模 | 7,000 行 × 59 列，约 3.6 MB |
| 缺失 | 0 |
| 标签 | `Abnormal_Label`：0 = 6,045，1 = 955（13.64%） |
| 异常类型 | 8 类（不含 Normal）；完整分布见 `notebooks/01_explore.ipynb` 输出与 `outputs/report/figures/fig1_data_profile.png`；Top-140 构成见 `outputs/top140_type_composition.csv` |
| SHA256 | `24ce85d57c4d9357caed886be14c86bc333fb8d0472359b52c40744cf5797600` |

## 使用约定

- `audit_raw.csv` 是只读原始数据；所有清洗结果写入 `outputs/audit_clean.csv`。
- 标签只用于评估，不参与训练、不参与规则定义。
- `Risk_Class`、`Abnormal_Type` 是标签/标签派生列，不能作为特征；前者删除，后者只作对照。
- 数据仅用于课程实践；真实审计场景需补充凭证、合同、审批链等外部信息。

## 字段字典

59 列的字段名、中文含义、类型、缺失率、唯一值数、示例值见 `outputs/field_dictionary.csv`；字段分组见 `outputs/data_overview.csv`。
