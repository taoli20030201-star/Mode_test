import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import pandas as pd
import numpy as np

# 1. 读取数据
train = pd.read_csv('train.csv')

# 2. 备份原始数据（以防万一）
train_original = train.copy()

# 3. 删除无用的列
train.drop(['Name', 'Ticket', 'PassengerId', 'Cabin'], axis=1, inplace=True)

# 4. 处理缺失值
# 4a. Age：用中位数填充
train['Age'].fillna(train['Age'].median(), inplace=True)

# 4b. Embarked：用出现次数最多的值（众数）填充
train['Embarked'].fillna(train['Embarked'].mode()[0], inplace=True)

# 5. 将文本标签转换为数字
# 5a. Sex: male=0, female=1
train['Sex'] = train['Sex'].map({'male': 0, 'female': 1})

# 5b. Embarked: S=0, C=1, Q=2
train['Embarked'] = train['Embarked'].map({'S': 0, 'C': 1, 'Q': 2})

# 6. 检查清洗结果
print("="*50)
print("清洗后的数据前5行：")
print(train.head())

print("="*50)
print("清洗后缺失值统计（应该全为0）：")
print(train.isnull().sum())

print("="*50)
print("清洗后的数据类型：")
print(train.dtypes)

# 保存清洗后的数据为新的 CSV 文件
train.to_csv('train_cleaned.csv', index=False)
print("✅ 清洗后的数据已保存为 train_cleaned.csv")