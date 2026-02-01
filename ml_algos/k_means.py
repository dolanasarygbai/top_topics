 # Упрощенный скрипт только с базовой кластеризацией

import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans
from collections import Counter

# Загрузка данных и указываете путь до файла. В квадратных скобках название столбца для анализа 
df = pd.read_excel(r"C:\ros\messages2.xlsx") 
texts = df['Текст сообщения'].dropna().astype(str).tolist()

# Ограничиваем количество для скорости
texts = texts[:5000]

# Создаем TF-IDF вектора
vectorizer = TfidfVectorizer(max_features=1000)
X = vectorizer.fit_transform(texts)

# Кластеризация
kmeans = KMeans(n_clusters=10, random_state=42)
labels = kmeans.fit_predict(X)

# Анализ результатов
for i in range(10):
    indices = [idx for idx, label in enumerate(labels) if label == i]
    print(f"\nТема {i} ({len(indices)} сообщений):")
    
    # Примеры сообщений
    for idx in indices[:2]:
        print(f"  - {texts[idx][:100]}...")
    
    # Ключевые слова

    print(f"  Ключевые слова: ...")

