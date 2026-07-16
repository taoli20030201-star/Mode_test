import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import pandas as pd

# 读取训练数据
train = pd.read_csv('train.csv')

# groupby('Sex') 是把数据按性别分组
# ['Survived'].mean() 是计算每组幸存者的比例

# 1. 查看数据前5行
print("=" * 50)
print("数据前5行：")
print(train.head())

# 2. 查看数据基本信息（列名、非空数量、数据类型）
print("=" * 50)
print("数据基本信息：")
print(train.info())

# 3. 查看数据统计摘要（数值列）
print("=" * 50)
print("数值列统计摘要：")
print(train.describe())

# 4. 查看缺失值情况
print("=" * 50)
print("缺失值统计：")
print(train.isnull().sum())

# 5. 查看分类列的取值分布
print("=" * 50)
print("Survived 分布（0=遇难，1=幸存）：")
print(train['Survived'].value_counts())

print("=" * 50)
print("性别分布：")
print(train['Sex'].value_counts())

print("=" * 50)
print("舱位分布：")
print(train['Pclass'].value_counts())

# 6. 交叉统计：性别 vs 是否幸存
print("=" * 50)
print("性别与幸存情况的交叉统计：")
cross_tab = pd.crosstab(train['Sex'], train['Survived'])
print(cross_tab)

# 7. 计算男女各自的幸存率
print("=" * 50)
print("男女幸存率对比：")
survival_rate = train.groupby('Sex')['Survived'].mean() * 100
print(survival_rate)

print("=" * 50)
print("不同舱位的幸存率对比：")
print(train.groupby('Pclass')['Survived'].mean() * 100)