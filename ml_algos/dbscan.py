import pandas as pd
import re
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import DBSCAN
from sklearn.decomposition import PCA

# 1. Загрузка данных
file_path = r'C:\rostelecom\messages2.xlsx'
df = pd.read_excel(file_path)
# Удаляем пустые строки, если есть
df = df.dropna(subset=['Текст сообщения'])
texts = df['Текст сообщения'].astype(str).tolist()

# 2. Очистка текста
def preprocess_text(text):
    text = text.lower()
    text = re.sub(r'[^а-яё ]', ' ', text)
    return text

clean_texts = [preprocess_text(t) for t in texts]

# 3. Векторизация
# Добавляем min_df=5, чтобы игнорировать опечатки и уникальные слова
vectorizer = TfidfVectorizer(max_features=10000, min_df=5, stop_words=None)
X = vectorizer.fit_transform(clean_texts)

# 4. Кластеризация DBSCAN
# eps — радиус окрестности (ключевой параметр, подбирается экспериментально)
# min_samples — сколько сообщений должно быть в радиусе, чтобы сформировать тему
dbscan = DBSCAN(eps=0.5, min_samples=10, metric='cosine')
labels = dbscan.fit_predict(X)

# Добавляем метки кластеров в исходный фрейм
df['cluster'] = labels

# 5. Анализ результатов
print(f"Найдено тем (кластеров): {len(set(labels)) - (1 if -1 in labels else 0)}")
print(f"Сообщений вне тем (шум): {list(labels).count(-1)}\n")

# Функция для вывода ключевых слов темы
def get_top_keywords(cluster_id, n_terms=10):
    if cluster_id == -1: return "Шум/Разное"
    
    # Собираем все тексты одного кластера
    indices = [i for i, label in enumerate(labels) if label == cluster_id]
    cluster_vector = X[indices].mean(axis=0)
    
    # Сортируем слова по весу TF-IDF в этом кластере
    sorted_indices = cluster_vector.argsort().A1[::-1]
    terms = vectorizer.get_feature_names_out()
    return ", ".join([terms[i] for i in sorted_indices[:n_terms]])

# Вывод ТОП-10 самых крупных тем
top_clusters = df[df['cluster'] != -1]['cluster'].value_counts().head(10)

for cluster_id, count in top_clusters.items():
    keywords = get_top_keywords(cluster_id)
    print(f"Тема №{cluster_id} ({count} сообщений):")
    print(f"Ключевые слова: {keywords}\n")
