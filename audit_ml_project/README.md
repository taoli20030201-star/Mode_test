# 智慧审计 × 机器学习 · 期中实践

> 用 2% 审计容量，从 7,000 笔交易中筛出 140 条最值得核查的样本。
> 全流程 CPU、免费工具、约 3 天；不用 GPU、深度学习、数据库、云服务或 SHAP。

---

## 一、一句话工作流

`audit_raw.csv` → pandas 探查 → 缺失/编码/金额 log1p+z/质量检查四步清洗 → Benford 首位数基线 → Isolation Forest 打分 → Top-K 召回 + 2×2 交叉 → 人工归因 3 例 → 6 张汇报图 + 10 页主线 PPT（+3 页附录）。

---

## 二、目录结构

| 路径 | 内容 |
|---|---|
| `data/audit_raw.csv` | 原始数据，7,000 × 59，只读不改 |
| `data/README.md` | 数据来源、规模、标签分布、SHA256 |
| `notebooks/01_explore.ipynb` | 阶段一：初步探查 |
| `notebooks/02_clean.ipynb` | 阶段二：清洗与质量检查 |
| `notebooks/03_analysis.ipynb` | 阶段三：Benford / Isolation Forest / Top-K / 2×2 / 人工归因 |
| `outputs/` | 清洗数据、中间表、日志、阶段图 |
| `outputs/figures/` | 阶段一至三的 6 张分析图 |
| `outputs/report/` | 阶段四：6 张汇报图、PPT、预览、问答、演讲备注、图集说明、生成日志 |
| `requirements.txt` | Python 依赖 |

---

## 三、四阶段交付物

### 阶段一 · 获取数据

| 要求 | 交付物 |
|---|---|
| 原始 CSV | `data/audit_raw.csv` |
| 数据概览表 | `outputs/data_overview.csv`、`outputs/dtype_missing_summary.csv`、`outputs/numeric_summary.csv`、`outputs/categorical_cardinality.csv` |
| 字段字典 | `outputs/field_dictionary.csv` |
| 探查小结 | `outputs/explore_summary.txt` |
| 探查图 | `outputs/figures/01_label_distribution.png`、`outputs/figures/02_amount_skew.png` |

### 阶段二 · 清洗数据

| 要求 | 交付物 |
|---|---|
| 缺失处置表 | `outputs/missing_handling.csv` |
| 编码后特征矩阵 + 维度变化 | `outputs/audit_clean.csv`、`outputs/encoding_summary.csv`、`outputs/dimension_change.csv` |
| 金额变换前后对比 | `outputs/figures/03_amount_transform_compare.png` |
| 清洗日志 | `outputs/cleaning_log.txt`、`outputs/quality_check_before_clean.csv`、`outputs/quality_check_after_clean.csv`、`outputs/quality_check_compare.csv` |
| 清洗后数据 | `outputs/audit_clean.csv` |

### 阶段三 · 分析数据

| 要求 | 交付物 |
|---|---|
| Benford 基线 | `outputs/benford_digit_table.csv`、`outputs/benford_by_label.csv`、`outputs/benford_mad_by_label.csv`、`outputs/figures/04_benford_distribution.png` |
| Isolation Forest 分数 | `outputs/isolation_forest_scores.csv` |
| Top-K 评估 | `outputs/topk_metrics.csv`、`outputs/top140_audit_list.csv`、`outputs/figures/05_topk_recall.png` |
| 规则—模型 2×2 | `outputs/model_benford_crosstab.csv`、`outputs/model_benford_cell_stats.csv` |
| 人工归因 3 例 | `outputs/case_studies.csv` |
| 分析日志 | `outputs/analysis_log.txt` |

> `isolation_forest_scores.csv` 保留原始交易顺序，便于按 `Transaction_ID` 回连；`top140_audit_list.csv` 已按 `anomaly_score` 降序。

### 阶段四 · 整理汇报

| 要求 | 交付物 |
|---|---|
| 5—6 张汇报图 | `outputs/report/figures/fig1_data_profile.png` … `fig6_rule_model.png` |
| 10 页 PPT | `outputs/report/智慧审计_机器学习_10页汇报.pptx`（主线 10 页 + 附录 A-1–A-3） |
| 答辩问答 | `outputs/report/答辩问答.md`（3 个必答 + 23 个追问） |
| 演讲备注 | `outputs/report/演讲备注.md` |
| 图集说明 | `outputs/report/图集说明.md` |
| 生成日志 | `outputs/report/report_log.txt` |

---

## 四、怎么运行

> 以下命令默认在 `audit_ml_project/` 目录下执行；notebook 需要从 `notebooks/` 目录启动。

### 1. 环境

```bash
pip install -r requirements.txt
```

已验证：Python 3.13.9、pandas 2.3.3、numpy 2.3.5、scikit-learn 1.7.2、matplotlib 3.10.6。
课程要求 Python 3.9+；如果在 3.9 上运行，请把 matplotlib / scikit-learn 降到支持 3.9 的版本。
PPT 重新生成需要 Windows + Microsoft PowerPoint（脚本用 COM 控制）。

### 2. 按顺序运行 notebook

在 `notebooks/` 目录下依次运行：

1. `01_explore.ipynb`
2. `02_clean.ipynb`
3. `03_analysis.ipynb`

notebook 使用相对路径 `../data/audit_raw.csv` 和 `../outputs/`，从 `notebooks/` 目录启动 Jupyter 即可。

### 3. 重新生成汇报图

```bash
python outputs/report/build_report_figures.py
```

### 4. 重新生成 PPT

```bash
python outputs/report/build_report_ppt.py
```

脚本会生成 13 页（10 页主线 + 3 页备用附录）和 `preview/slide01–slide13.png`。
附录页脚标注“不计入 10 页主线”；正式汇报只放 1–10 页。

---

## 五、核心数字

| 指标 | 数值 |
|---|---|
| 原始数据 | 7,000 行 × 59 列，0 缺失 |
| 异常 / 正常 | 955 / 6,045（13.64%，6.33 : 1） |
| 清洗后 | 115 列 = 111 特征 + 2 标识 + 2 标签 |
| 删除行 | 0 |
| Benford 整体 MAD | 0.00405（高度符合） |
| Benford 正常 / 异常 MAD | 0.00325 / 0.03633 |
| Isolation Forest | 200 棵树，contamination=0.02，140 条 |
| ROC-AUC / PR-AUC | 0.8795 / 0.6466 |
| Top-140 | 129 命中，92.1% 精确率，13.5% 召回，6.75× lift |
| 2×2 双命中 | 18 条，94.4% 异常率 |
| 2×2 仅模型 | 122 条，91.8% 异常率 |
| 2×2 仅 Benford | 644 条，26.2% 异常率 |
| Jaccard | 0.023 |
| 类型覆盖 | 7 / 8 类 |
| 案例 | CN-TXN-0005220，模型排名 1 / 7,000 |

---

## 六、数据与口径

- 数据来源：Kaggle 数据集“Financial Audit Transactions Dataset”（课程指定）；本地文件 `data/audit_raw.csv`。
- SHA256：`24ce85d57c4d9357caed886be14c86bc333fb8d0472359b52c40744cf5797600`。
- 标签 `Abnormal_Label` 只用于评估，不参与训练、不参与规则定义；`Risk_Class` 已删除，`Abnormal_Type` 只作对照。
- `contamination=0.02` 是审计容量假设（7,000 × 2% = 140），不是真实异常率。
- 拓扑/关系字段的高基数 ID 未入模；城市字段暂不编码，保留省份。
- 当前为课程实践数据；真实落地需按时间切分，并考虑有监督/图特征扩展。

---

## 七、汇报入口

- [阶段四 README](outputs/report/README.md)
- [答辩问答](outputs/report/答辩问答.md)
- [演讲备注](outputs/report/演讲备注.md)
- [图集说明](outputs/report/图集说明.md)
- [PPT](outputs/report/智慧审计_机器学习_10页汇报.pptx)
- [封面预览](outputs/report/preview/slide01.png)

---

## 八、完成度自检

- [x] 阶段一：原始 CSV、数据概览、字段字典、探查图。
- [x] 阶段二：缺失处置表、编码与维度变化、金额变换图、清洗日志、`audit_clean.csv`。
- [x] 阶段三：Benford、Isolation Forest、Top-K、2×2、3 条案例、分析日志。
- [x] 阶段四：6 张汇报图、10 页主线 PPT + 3 页附录、3 个必答问题 + 23 个追问、演讲备注。
- [x] 全部数字可回到 `notebooks/` 和 `outputs/` 核对。
- [x] 全流程 CPU，无 GPU / 深度学习 / 数据库 / 云服务 / SHAP。
