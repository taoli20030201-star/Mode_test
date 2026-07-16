# 用清洗后的数据训练一个模型，预测泰坦尼克号上的乘客能否幸存。
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score, confusion_matrix

# 1. 读取清洗后的数据
train = pd.read_csv('train_cleaned.csv')

# 2. 准备特征 X 和标签 y
# 特征：除了 Survived 之外的所有列
X = train.drop('Survived', axis=1)
# 标签：Survived 列
y = train['Survived']

# 3. 划分训练集和验证集（80% 训练，20% 验证）
X_train, X_val, y_train, y_val = train_test_split(X, y, test_size=0.2, random_state=42)

# 4. 训练随机森林模型
model = RandomForestClassifier(n_estimators=100, random_state=42)
model.fit(X_train, y_train)

# 5. 在验证集上预测并评估
y_pred = model.predict(X_val)
acc = accuracy_score(y_val, y_pred)

print(f"✅ 模型训练完成！验证集准确率: {acc:.2%}")

# 6. 查看哪些特征最重要
feature_importance = pd.DataFrame({
    '特征': X.columns,
    '重要性': model.feature_importances_
}).sort_values('重要性', ascending=False)

print("="*50)
print("特征重要性排序（越靠上越重要）：")
print(feature_importance)