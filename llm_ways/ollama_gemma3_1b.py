"""
МИНИМАЛЬНЫЙ ТЕСТ - только самое необходимое
"""

import pandas as pd
import requests
import json

# 1. Чтение файла
print("Чтение файла messages3.xlsx...")
df = pd.read_excel("messages3.xlsx", engine='openpyxl')

# Берем первый столбец
messages = df.iloc[:, 0].dropna().astype(str).tolist()

print(f"Найдено {len(messages)} сообщений:")
for i, msg in enumerate(messages[:5], 1):
    print(f"{i}. {msg[:50]}...")

# 2. Берем только 3 сообщения для теста
sample = messages[:3]
sample_text = "\n".join([f"{i+1}. {msg}" for i, msg in enumerate(sample)])

# 3. Очень простой промт
prompt = f"""Вот обращения в банк:
{sample_text}

Назови 2 основные темы. Формат: ["тема1", "тема2"]"""

print("\nОтправка запроса к модели...")

try:
    response = requests.post(
        "http://localhost:11434/api/generate",
        json={
            "model": "gemma3:1b",
            "prompt": prompt,
            "stream": False,
            "options": {"num_predict": 100}
        },
        timeout=30
    )
    
    if response.status_code == 200:
        result = response.json().get("response", "")
        print(f"\n✅ Ответ модели:")
        print(result)
        
        # Сохраняем результат
        with open("minimal_result.txt", "w", encoding="utf-8") as f:
            f.write(result)
        print("\n✅ Результат сохранен в minimal_result.txt")
    else:
        print(f"❌ Ошибка: {response.status_code}")
        
except Exception as e:
    print(f"❌ Ошибка: {e}")
