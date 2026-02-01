import pandas as pd
import re
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans

# 1. Загрузка данных
file_path = r'C:\rostelecom\messages2.xlsx'
df = pd.read_excel(file_path)
texts = df['Текст сообщения'].astype(str).tolist()

# 2. Предобработка текста (очистка)
def preprocess_text(text):
    text = text.lower()
    text = re.sub(r'[^а-яё ]', ' ', text)  # Оставляем только русские буквы
    return text

clean_texts = [preprocess_text(t) for t in texts]

# 3. Векторизация (TF-IDF)
# Используем стоп-слова, чтобы исключить "мусор" (и, в, на, что и т.д.)
vectorizer = TfidfVectorizer(max_features=5000, stop_words=None) 
X = vectorizer.fit_transform(clean_texts)

# 4. Кластеризация (K-Means)
num_clusters = 10
model = KMeans(n_clusters=num_clusters, random_state=42, n_init=10)
model.fit(X)

# 5. Извлечение топ-слов для каждой темы
print("Топ-10 тем обращений:\n")
order_centroids = model.cluster_centers_.argsort()[:, ::-1]
terms = vectorizer.get_feature_names_out()

for i in range(num_clusters):
    print(f"Тема №{i+1}: ", end="")
    # Выводим 7 ключевых слов, которые формируют смысл темы
    topic_words = [terms[ind] for ind in order_centroids[i, :7]]
    print(", ".join(topic_words))
