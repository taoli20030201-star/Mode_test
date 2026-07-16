import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import pandas as pd
import numpy as np
from sklearn.datasets import load_wine
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score

# 加载红酒数据集（13个特征，3个分类）
wine = load_wine()
X, y = wine.data, wine.target

print(f"数据集形状: {X.shape}")
print(f"特征名称: {wine.feature_names}")
print(f"标签类别: {wine.target_names}")

# 划分训练集和测试集
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# 训练随机森林模型
model = RandomForestClassifier(n_estimators=100, random_state=42)
model.fit(X_train, y_train)

# 预测并计算准确率
y_pred = model.predict(X_test)
acc = accuracy_score(y_test, y_pred)

print(f"模型训练完成！准确率: {acc:.2%}")