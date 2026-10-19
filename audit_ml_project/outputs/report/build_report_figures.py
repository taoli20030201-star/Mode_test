# -*- coding: utf-8 -*-
"""生成阶段四汇报用的 6 张统一风格图。"""
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter

BASE = Path(__file__).resolve().parents[2]
OUT = BASE / 'outputs'
FIG = OUT / 'report' / 'figures'
FIG.mkdir(parents=True, exist_ok=True)

# ---------------- 统一风格 ----------------
C = {
    'primary': '#2E86C1',
    'primary_light': '#AED6F1',
    'primary_dark': '#1F4E79',
    'accent': '#E67E22',
    'accent_light': '#F5CBA7',
    'red': '#C0392B',
    'green': '#27AE60',
    'gray': '#7F8C8D',
    'light_gray': '#ECF0F1',
    'dark': '#1B2631',
    'text': '#34495E',
}
plt.rcParams.update({
    'font.sans-serif': ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS'],
    'axes.unicode_minus': False,
    'font.size': 11,
    'axes.titlesize': 14,
    'axes.labelsize': 11,
    'xtick.labelsize': 10,
    'ytick.labelsize': 10,
    'legend.fontsize': 9.5,
    'figure.facecolor': 'white',
    'axes.facecolor': 'white',
    'axes.edgecolor': '#D5D8DC',
    'axes.linewidth': 0.9,
    'axes.grid': True,
    'grid.color': '#E5E8E8',
    'grid.linewidth': 0.8,
    'axes.spines.top': False,
    'axes.spines.right': False,
    'savefig.facecolor': 'white',
    'savefig.bbox': 'tight',
    'savefig.dpi': 200,
})


def save(fig, name):
    path = FIG / name
    fig.savefig(path, dpi=200, bbox_inches='tight')
    plt.close(fig)
    print('saved', path)


# ---------------- 数据 ----------------
raw = pd.read_csv(BASE / 'data/audit_raw.csv', encoding='utf-8')
clean = pd.read_csv(OUT / 'audit_clean.csv', encoding='utf-8-sig')
scores = pd.read_csv(OUT / 'isolation_forest_scores.csv', encoding='utf-8-sig')
digit = pd.read_csv(OUT / 'benford_digit_table.csv', encoding='utf-8-sig')
label_digit = pd.read_csv(OUT / 'benford_by_label.csv', encoding='utf-8-sig')
topk = pd.read_csv(OUT / 'topk_metrics.csv', encoding='utf-8-sig')
cell_stats = pd.read_csv(OUT / 'model_benford_cell_stats.csv', encoding='utf-8-sig')
type_comp = pd.read_csv(OUT / 'top140_type_composition.csv', encoding='utf-8-sig')
field_dict = pd.read_csv(OUT / 'field_dictionary.csv', encoding='utf-8-sig')

y = clean['Abnormal_Label'].to_numpy()
score = scores['anomaly_score'].to_numpy()
order = np.argsort(-score)
total_pos = int(y.sum())

# ================================================================
# 图 1：数据概览（标签 + 字段分组 + 异常类型）
# ================================================================
fig, axes = plt.subplots(1, 3, figsize=(12.6, 4.9),
                         gridspec_kw={'width_ratios': [0.95, 1.25, 1.15]})

# (a) 标签环形图
ax = axes[0]
vals = [6045, 955]
ax.pie(vals, colors=[C['primary_light'], C['accent']], startangle=90,
       counterclock=False, wedgeprops=dict(width=0.36, edgecolor='white', linewidth=2))
ax.text(0, 0.14, '955', ha='center', va='center', fontsize=24, fontweight='bold', color=C['accent'])
ax.text(0, -0.20, '异常\n13.64%', ha='center', va='center', fontsize=11, color=C['text'])
ax.set_title('标签分布：6.33 : 1', pad=8)
ax.legend(['正常 6,045', '异常 955'], loc='lower center', bbox_to_anchor=(0.5, -0.20),
          ncol=2, frameon=False, fontsize=9)

# (b) 字段分组
ax = axes[1]
groups = field_dict['字段分组'].value_counts()
groups = groups.sort_values(ascending=True)
colors = [C['gray'] if g == '标签' else C['primary'] for g in groups.index]
bars = ax.barh(groups.index, groups.values, color=colors, height=0.62)
for b, v in zip(bars, groups.values):
    ax.text(v + 0.15, b.get_y() + b.get_height() / 2, str(v),
            va='center', fontsize=9, color=C['text'])
ax.set_xlim(0, max(groups.values) + 2.0)
ax.set_xlabel('列数')
ax.set_title('59 列 = 11 组字段（含标签）', pad=8)
ax.grid(axis='y', visible=False)
ax.tick_params(axis='y', labelsize=9)

# (c) 异常类型
ax = axes[2]
types = raw['Abnormal_Type'].value_counts().drop('Normal')
types = types.sort_values(ascending=True)
type_zh = {
    'Unusual_Amount_Spike': '金额突增',
    'Split_Transaction': '拆分交易',
    'High_Risk_Vendor': '高风险供应商',
    'Cross_Province_Mismatch': '跨省不匹配',
    'Duplicate_Invoice': '重复发票',
    'Round_Trip_Transfer': '资金回流',
    'Temporal_Burst': '时间聚集',
    'Tax_ID_Mismatch': '税号不符',
}
type_labels = [type_zh.get(t, t) for t in types.index]
bars = ax.barh(type_labels, types.values, color=C['primary'], height=0.62)
for b, v in zip(bars, types.values):
    ax.text(v + 2, b.get_y() + b.get_height() / 2, str(v),
            va='center', fontsize=9, color=C['text'])
ax.set_xlim(0, max(types.values) + 22)
ax.set_xlabel('交易数')
ax.set_title('955 条异常 = 8 类', pad=8)
ax.grid(axis='y', visible=False)
ax.tick_params(axis='y', labelsize=9)

fig.subplots_adjust(wspace=0.55)
save(fig, 'fig1_data_profile.png')

# ================================================================
# 图 2：清洗维度变化
# ================================================================
fig, axes = plt.subplots(1, 2, figsize=(12.4, 4.8),
                         gridspec_kw={'width_ratios': [1.15, 0.95]})

# (a) 瀑布图
ax = axes[0]
labels = ['原始\n59 列', '剔除 13 列', 'One-Hot\n净增 69 列', '最终\n115 列']
x = np.arange(len(labels))
# 底部、高度、颜色
bottoms = [0, 46, 46, 0]
heights = [59, 13, 69, 115]
colors = [C['primary'], C['red'], C['green'], C['primary_dark']]
bars = ax.bar(x, heights, bottom=bottoms, color=colors, width=0.55, alpha=0.92)
# 连接线
for i in range(len(x) - 1):
    top = bottoms[i] + heights[i]
    ax.plot([x[i] + 0.275, x[i + 1] - 0.275], [top, top], color=C['gray'], lw=1, ls='--')
# 标注
for i, (b, h) in enumerate(zip(bottoms, heights)):
    if i == 0:
        txt = '59'
    elif i == 1:
        txt = '−13'
    elif i == 2:
        txt = '+69'
    else:
        txt = '115'
    ypos = b + h + 3.5
    ax.text(x[i], ypos, txt, ha='center', va='bottom', fontsize=13, fontweight='bold',
            color=colors[i])
ax.set_xticks(x)
ax.set_xticklabels(labels, fontsize=10)
ax.set_ylim(0, 135)
ax.set_ylabel('列数')
ax.set_title('59 → 115 维：不删一行，只做加减法', pad=8)
ax.grid(axis='x', visible=False)

# (b) 最终构成
ax = axes[1]
comp = pd.DataFrame({
    '类别': ['One-Hot 类别', '数值特征', '金额 log1p+z', '标识列', '标签列'],
    '列数': [79, 28, 4, 2, 2],
})
comp = comp.sort_values('列数', ascending=True)
bar_colors = [C['gray'], C['accent'], C['primary_light'], C['primary'], C['red']]
# 按排序后的类别取色
color_map = {'One-Hot 类别': C['primary'], '数值特征': C['primary_light'],
             '金额 log1p+z': C['accent'], '标识列': C['gray'], '标签列': C['red']}
bars = ax.barh(comp['类别'], comp['列数'], color=[color_map[c] for c in comp['类别']], height=0.6)
for b, v in zip(bars, comp['列数']):
    ax.text(v + 1.2, b.get_y() + b.get_height() / 2, str(v),
            va='center', fontsize=10, color=C['text'])
ax.set_xlim(0, 92)
ax.set_xlabel('列数')
ax.set_title('115 = 111 特征 + 2 标识 + 2 标签', pad=8)
ax.grid(axis='y', visible=False)
ax.text(0.98, 0.04, '0 行删除 · 0 缺失\n9 项质量检查通过',
        transform=ax.transAxes, ha='right', va='bottom', fontsize=9.5,
        color=C['green'], bbox=dict(boxstyle='round,pad=0.35', fc='white',
                                    ec=C['green'], alpha=0.9))

fig.subplots_adjust(wspace=0.42)
save(fig, 'fig2_cleaning.png')

# ================================================================
# 图 3：金额变换
# ================================================================
amount = raw['Transaction_Amount_RMB'].to_numpy()
log_amount = np.log1p(amount)

fig, axes = plt.subplots(1, 2, figsize=(12.4, 4.8))
ax = axes[0]
ax.hist(amount, bins=65, color=C['primary'], alpha=0.9)
med = np.median(amount)
ax.axvline(med, color=C['red'], ls='--', lw=1.6)
ax.text(med * 1.08, ax.get_ylim()[1] * 0.86, f'中位数\n{med:,.0f} 元',
        color=C['red'], fontsize=9.5, va='top')
ax.set_xlabel('交易金额（元）')
ax.set_ylabel('交易数')
ax.set_title('原始金额：严重右偏（skew = 10.70）', pad=8)
ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, p: f'{v/1e4:.0f}万'))
ax.text(0.97, 0.70, '最大 318 万\n≈ 中位数的 138 倍',
        transform=ax.transAxes, ha='right', va='top', fontsize=9.5, color=C['text'],
        bbox=dict(boxstyle='round,pad=0.35', fc='white', ec='#D5D8DC', alpha=0.9))

ax = axes[1]
ax.hist(log_amount, bins=65, color=C['green'], alpha=0.85)
ax.set_xlabel('log1p(交易金额)')
ax.set_ylabel('交易数')
ax.set_title('log1p 后：接近对称（skew = 0.52）', pad=8)
ax.text(0.97, 0.92, '再做 z-score\n均值 0 · 标准差 1',
        transform=ax.transAxes, ha='right', va='top', fontsize=9.5, color=C['text'],
        bbox=dict(boxstyle='round,pad=0.35', fc='white', ec='#D5D8DC', alpha=0.9))

fig.subplots_adjust(wspace=0.28)
save(fig, 'fig3_amount_transform.png')

# ================================================================
# 图 4：Benford
# ================================================================
fig, axes = plt.subplots(1, 2, figsize=(12.4, 4.9))
x = np.arange(1, 10)
w = 0.38

ax = axes[0]
ax.bar(x - w / 2, digit['实际占比'], w, color=C['primary'], label='实际观测')
ax.bar(x + w / 2, digit['Benford 期望'], w, color=C['gray'], alpha=0.55, label='Benford 期望')
# 突出数字 5
ax.axvspan(4.5, 5.5, color=C['accent'], alpha=0.12, zorder=0)
ax.annotate('实际 9.46%\n期望 7.92%', xy=(5, 0.0946), xytext=(6.2, 0.155),
            arrowprops=dict(arrowstyle='->', color=C['accent'], lw=1.4),
            fontsize=9.5, color=C['accent'], ha='left')
ax.set_xticks(x)
ax.set_xlabel('金额首位数字')
ax.set_ylabel('占比')
ax.set_ylim(0, 0.34)
ax.set_title('总体：MAD = 0.00405（高度符合）', pad=8)
ax.legend(frameon=False)

ax = axes[1]
ax.plot(x, label_digit['正常组实际占比'], marker='o', lw=1.8, color=C['primary'],
        label='正常组')
ax.plot(x, label_digit['异常组实际占比'], marker='s', lw=1.9, color=C['accent'],
        label='异常组')
ax.plot(x, label_digit['Benford 期望'], ls='--', lw=1.4, color=C['gray'],
        label='Benford 期望')
ax.axvspan(4.5, 5.5, color=C['accent'], alpha=0.10, zorder=0)
ax.annotate('异常组 19.48%', xy=(5, 0.1948), xytext=(5.6, 0.235),
            arrowprops=dict(arrowstyle='->', color=C['accent'], lw=1.4),
            fontsize=9.5, color=C['accent'])
ax.set_xticks(x)
ax.set_xlabel('金额首位数字')
ax.set_ylabel('占比')
ax.set_ylim(0, 0.34)
ax.set_title('分组：正常 0.00325 / 异常 0.03633', pad=8)
ax.legend(frameon=False, loc='upper right')

fig.subplots_adjust(wspace=0.28)
save(fig, 'fig4_benford.png')

# ================================================================
# 图 5：Top-K
# ================================================================
k_grid = np.unique(np.linspace(1, 2000, 300).astype(int))
recall_curve = np.array([y[order[:k]].sum() / total_pos for k in k_grid])
precision_curve = np.array([y[order[:k]].sum() / k for k in k_grid])

fig, axes = plt.subplots(1, 2, figsize=(12.4, 4.9))
ax = axes[0]
ax.plot(k_grid, recall_curve, color=C['primary'], lw=2.2, label='Isolation Forest')
ax.plot(k_grid, k_grid / len(y), ls='--', color=C['gray'], lw=1.3, label='随机抽查')
ax.axvline(140, color=C['accent'], ls=':', lw=1.3, label='2% 容量 = 140 条')
ks_mark = [50, 100, 140]
rs_mark = [y[order[:k]].sum() / total_pos for k in ks_mark]
ax.scatter(ks_mark, rs_mark, color=C['red'], zorder=5, s=30, label='K=50 / 100 / 140')
ax.text(0.98, 0.10,
        'K=50   5.1%\nK=100 10.1%\nK=140 13.5%',
        transform=ax.transAxes, ha='right', va='bottom', fontsize=9.2,
        bbox=dict(boxstyle='round,pad=0.35', fc='white', ec=C['red'], alpha=0.92))
ax.set_xlim(0, 2000)
ax.set_ylim(0, 1)
ax.set_xlabel('K（审计抽查条数）')
ax.set_ylabel('Recall@K（召回率）')
ax.set_title('召回率：查得越多，覆盖越高', pad=8)
ax.legend(frameon=False, loc='upper left')

ax = axes[1]
ax.plot(k_grid, precision_curve, color=C['green'], lw=2.2, label='Precision@K')
ax.axhline(y.mean(), ls='--', color=C['red'], lw=1.3, label='异常占比 13.64%')
ax.axvline(140, color=C['accent'], ls=':', lw=1.3)
for k in [50, 100, 140]:
    p = y[order[:k]].sum() / k
    ax.scatter([k], [p], color=C['red'], zorder=5, s=30)
ax.text(0.98, 0.10,
        'K=140  92.1%\nK=500  76.0%\nK=1000 57.8%',
        transform=ax.transAxes, ha='right', va='bottom', fontsize=9.2,
        bbox=dict(boxstyle='round,pad=0.35', fc='white', ec=C['green'], alpha=0.92))
ax.set_xlim(0, 2000)
ax.set_ylim(0, 1)
ax.set_xlabel('K（审计抽查条数）')
ax.set_ylabel('Precision@K（精确率）')
ax.set_title('精确率：名单质量随 K 下降', pad=8)
ax.legend(frameon=False, loc='upper right')

fig.subplots_adjust(wspace=0.28)
save(fig, 'fig5_topk.png')

# ================================================================
# 图 6：规则 × 模型 + 类型入选率
# ================================================================
fig, axes = plt.subplots(1, 2, figsize=(12.6, 5.0),
                         gridspec_kw={'width_ratios': [0.82, 1.32]})

# (a) 2×2 热力图（去掉 colorbar，避免挤压右侧标签）
ax = axes[0]
counts = np.zeros((2, 2), dtype=int)
rates = np.zeros((2, 2), dtype=float)
for _, r in cell_stats.iterrows():
    i = int(r['Benford_Flag'])
    j = int(r['model_top140'])
    counts[i, j] = int(r['记录数'])
    rates[i, j] = float(r['异常率'])
im = ax.imshow(rates, cmap='OrRd', vmin=0, vmax=1, aspect='auto')
for i in range(2):
    for j in range(2):
        ax.text(j, i, f"{counts[i, j]:,}\n异常率 {rates[i, j]:.1%}",
                ha='center', va='center', fontsize=10.5,
                color='white' if rates[i, j] > 0.55 else C['dark'],
                fontweight='bold' if rates[i, j] > 0.85 else 'normal')
ax.set_xticks([0, 1])
ax.set_xticklabels(['模型未标记', '模型 Top-140'])
ax.set_yticks([0, 1])
ax.set_yticklabels(['Benford\n未标记', 'Benford\n标记'])
ax.set_title('规则 × 模型：低重合、高互补', pad=8)
ax.grid(False)
ax.set_xlabel('Jaccard = 0.023（颜色越深 = 异常率越高）', fontsize=9.2,
              labelpad=8, color=C['text'])

# (b) 类型入选率（用中文短标签，避免长英文名挤压）
ax = axes[1]
tc = type_comp[type_comp['真实异常类型'] != 'Normal'].copy()
tc = tc.sort_values('该类入选率', ascending=True)
zh = {
    'Round_Trip_Transfer': '资金回流',
    'Duplicate_Invoice': '重复发票',
    'Tax_ID_Mismatch': '税号不符',
    'Split_Transaction': '拆分交易',
    'Unusual_Amount_Spike': '金额突增',
    'High_Risk_Vendor': '高风险供应商',
    'Temporal_Burst': '时间聚集',
    'Cross_Province_Mismatch': '跨省不匹配',
}
labels = [zh.get(t, t) for t in tc['真实异常类型']]
vals = tc['该类入选率'].to_numpy() * 100
colors = []
for t, v in zip(tc['真实异常类型'], vals):
    if t == 'Cross_Province_Mismatch':
        colors.append(C['red'])
    elif v >= 25:
        colors.append(C['accent'])
    else:
        colors.append(C['primary'])
bars = ax.barh(labels, vals, color=colors, height=0.62)
for b, v, n, total in zip(bars, vals, tc['条数'], tc['该类总数']):
    ax.text(v + 0.8, b.get_y() + b.get_height() / 2, f'{v:.1f}%  ({int(n)}/{int(total)})',
            va='center', fontsize=9.2, color=C['text'])
ax.axvline(13.5, color=C['gray'], ls='--', lw=1.2, label='整体召回 13.5%')
ax.set_xlim(0, 48)
ax.set_xlabel('该类被选入 Top-140 的比例')
ax.set_title('模型强项与短板：7 / 8 类异常被覆盖', pad=8)
ax.grid(axis='y', visible=False)
ax.tick_params(axis='y', labelsize=9.8)
ax.legend(loc='lower right', frameon=False, fontsize=9)

fig.subplots_adjust(wspace=0.34)
save(fig, 'fig6_rule_model.png')

print('\nALL FIGURES DONE')
