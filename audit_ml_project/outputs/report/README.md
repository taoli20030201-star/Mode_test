# 阶段四 · 汇报材料

> 任务 4.1 图集、4.2 PPT、4.3 答辩准备都在这个目录。
> 分析过程见 `notebooks/03_analysis.ipynb`；中间结果见 `outputs/`。

---

## 一、目录内容

| 文件 / 目录 | 用途 |
|---|---|
| `智慧审计_机器学习_10页汇报.pptx` | 10 页主线 + 3 页备用附录（16:9，可直接放映） |
| `preview/slide01.png` … `slide13.png` | 每页预览图（主线 1–10，附录 A-1–A-3） |
| `figures/fig1_data_profile.png` | 数据概览：标签、字段分组、异常类型 |
| `figures/fig2_cleaning.png` | 清洗：59 → 115 维、0 行删除 |
| `figures/fig3_amount_transform.png` | 金额：skew 10.70 → 0.52；131 笔 500,000 元 |
| `figures/fig4_benford.png` | Benford：整体 / 正常 / 异常组 MAD |
| `figures/fig5_topk.png` | Top-K：精确率、召回率、K 选择 |
| `figures/fig6_rule_model.png` | 2×2 + 类型覆盖 |
| `答辩问答.md` | 3 个必答问题 + 23 个可能追问 + 数字速查 + 附录索引 |
| `演讲备注.md` | 逐页口播、时间分配、现场提醒、附录翻页提示 |
| `图集说明.md` | 6 张图的口径、来源、关键结论 |
| `build_report_figures.py` | 从 CSV 重新生成 6 张图 |
| `build_report_ppt.py` | 用 PowerPoint COM 重新生成 PPT 和预览图 |
| `report_log.txt` | 阶段四生成日志：输入、输出、验证、口径 |

---

## 二、10 页结构

| 页 | 内容 | 对应图 |
|---:|---|---|
| 1 | 封面：问题 → 目标 → 方法 | — |
| 2 | 学了什么：针对问题拆四步 | 原生图形 |
| 3 | 为何选审计：可验证、可解释、可落地 | 原生卡片 |
| 4 | 获取数据 | `fig1_data_profile.png` |
| 5 | 清洗（一）：59 → 115 维 | `fig2_cleaning.png` |
| 6 | 清洗（二）：金额 log1p | `fig3_amount_transform.png` |
| 7 | 分析（一）：Benford | `fig4_benford.png` |
| 8 | 分析（二）：Isolation Forest + Top-K | `fig5_topk.png` |
| 9 | 理解与修正：规则 × 模型 2×2 | `fig6_rule_model.png` |
| 10 | 收束：结论、局限、下一步 | 原生卡片 |

### 附录（备用页，不计入 10 页主线）

| 页 | 内容 | 什么时候用 |
|---:|---|---|
| A-1（slide11） | 115 维构成表 | 回答“115 维怎么来的、为什么删列” |
| A-2（slide12） | Top-K 全表 | 回答“K 怎么选、召回为什么低” |
| A-3（slide13） | 一条案例的完整字段 | 回答“模型到底在看什么、案例是不是挑的” |

---

## 三、核心数字（答辩速查）

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
| 131 笔 500,000 元 | 全部为金额突增异常 |

---

## 四、怎么重新生成

> 以下命令默认在 `audit_ml_project/` 目录下执行。

### 1. 生成图

```bash
python outputs/report/build_report_figures.py
```

前提：`notebooks/01_explore.ipynb` → `02_clean.ipynb` → `03_analysis.ipynb` 已按顺序运行，`outputs/` 下有以下中间文件：

- `audit_clean.csv`
- `isolation_forest_scores.csv`
- `benford_digit_table.csv`
- `benford_by_label.csv`
- `topk_metrics.csv`
- `model_benford_cell_stats.csv`
- `top140_type_composition.csv`
- `field_dictionary.csv`

### 2. 生成 PPT

```bash
python outputs/report/build_report_ppt.py
```

前提：

- Windows + 已安装 Microsoft PowerPoint（脚本用 COM 控制 PowerPoint）。
- Python 包：`pywin32`、`Pillow`。
- 先运行 `build_report_figures.py`，确保 6 张图存在。

脚本会：

1. 新建 10 页主线 + 3 页备用附录（共 13 页，16:9，960 × 540 pt）。
2. 插入 6 张图、原生卡片、流程图和附录表格。
3. 保存为 `智慧审计_机器学习_10页汇报.pptx`（13 页）。
4. 导出 `preview/slide01.png` … `slide13.png`。

> 脚本只操作自己新建的 PowerPoint 实例，结束后关闭该实例；不会强制结束其他 PowerPoint 进程。
> 颜色使用 COM 的 BGR 顺序；中文同时设置 `Font.NameFarEast` / `Font.NameAscii`。

---

## 五、自检清单

- [x] 封面先讲问题，再讲目标与方法；第 2 页明确承接“针对这个问题”。
- [x] 主线正好 10 页，另有 3 页备用附录；附录页脚标注“不计入 10 页主线”。
- [x] 附录 A-1/A-2/A-3 的数字与主线一致（79/28/4/2/2；K 全表；案例排名 1/7,000）。
- [x] 6 张图全部插入，文件名与图集说明一致。
- [x] 每页标题、页码、脚注完整。
- [x] 中文无乱码；无文本溢出/重叠（逐页看 `preview/`）。
- [x] 关键数字与 `notebooks/03_analysis.ipynb` 输出一致。
- [x] 标签只用于评估；`Risk_Class` 已删除，`Abnormal_Type` 不作为特征。
- [x] 500,000 元的表述准确：是“131 笔重复金额 + 金额突增”，不是“审批阈值”。

---

## 六、口径与边界

- 当前数据是课程实践数据，不是真实机构数据；结论用于方法演练。
- `contamination=0.02` 是审计容量假设，不是真实异常率。
- Top-140 召回 13.5% 是 2% 容量下的正常结果；放宽 K 可提高召回。
- 3 条案例是“模型-only”案例，用于解释模型在抓什么；不是定罪结论。
- 真实落地要按时间切分，避免未来信息泄漏；可接有监督、图特征或审计工作台。
