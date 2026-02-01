# analyze.py - минимальный рабочий скрипт
import pandas as pd
import requests
import uuid
import json
import os
import time
from datetime import datetime

print("=" * 60)
print("🤖 АНАЛИЗ ОБРАЩЕНИЙ В ПОДДЕРЖКУ БАНКА")
print("=" * 60)

# 1. Загружаем конфиг
try:
    from config import AUTHORIZATION_KEY, EXCEL_PATH
    print(f"✅ Ключ загружен: {AUTHORIZATION_KEY[:20]}...")
except ImportError:
    print("❌ Создайте файл config.py с вашим ключом!")
    input("Нажмите Enter для выхода...")
    exit()

# 2. Проверяем файл
if not os.path.exists(EXCEL_PATH):
    print(f"❌ Файл не найден: {EXCEL_PATH}")
    input("Нажмите Enter для выхода...")
    exit()

# 3. Функция получения токена
def get_gigachat_token():
    """Получает токен доступа к GigaChat"""
    print("🔑 Получаю токен...")
    
    response = requests.post(
        "https://ngw.devices.sberbank.ru:9443/api/v2/oauth",
        headers={
            "Authorization": f"Basic {AUTHORIZATION_KEY}",
            "RqUID": str(uuid.uuid4()),
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json"
        },
        data={"scope": "GIGACHAT_API_PERS"},
        verify=False,
        timeout=30
    )
    
    if response.status_code == 200:
        data = response.json()
        print("✅ Токен получен")
        return data["access_token"]
    else:
        print(f"❌ Ошибка: {response.status_code}")
        print(response.text[:200])
        return None

# 4. Функция запроса к GigaChat
def ask_gigachat(prompt_text, token):
    """Отправляет запрос к GigaChat"""
    response = requests.post(
        "https://gigachat.devices.sberbank.ru/api/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        },
        json={
            "model": "GigaChat",
            "messages": [{"role": "user", "content": prompt_text}],
            "temperature": 0.3,
            "max_tokens": 3000
        },
        verify=False,
        timeout=60
    )
    
    if response.status_code == 200:
        return response.json()
    elif response.status_code == 401:  # Токен истёк
        print("🔄 Токен истёк, получаю новый...")
        return None
    else:
        print(f"⚠️ Ошибка API: {response.status_code}")
        return None

# 5. Чтение Excel файла
print(f"\n📖 Читаю файл: {EXCEL_PATH}")
try:
    df = pd.read_excel(EXCEL_PATH, engine='openpyxl')
    
    # Ищем столбец с текстом
    text_column = None
    for col in df.columns:
        col_lower = str(col).lower()
        if 'текст' in col_lower or 'сообщ' in col_lower:
            text_column = col
            break
    
    if text_column is None:
        print("⚠️ Столбец 'Текст сообщения' не найден, использую первый столбец")
        text_column = df.columns[0]
    
    # Берём до 10000 сообщений
    messages = df[text_column].dropna().astype(str).tolist()[:10000]
    print(f"✅ Прочитано {len(messages)} сообщений из столбца '{text_column}'")
    
except Exception as e:
    print(f"❌ Ошибка чтения Excel: {e}")
    input("Нажмите Enter для выхода...")
    exit()

# 6. Получаем токен
token = get_gigachat_token()
if not token:
    print("❌ Не удалось получить токен")
    input("Нажмите Enter для выхода...")
    exit()

# 7. Анализируем по частям
all_themes = {}
total_parts = min(10, (len(messages) + 999) // 1000)  # До 10 частей по 1000

print(f"\n🔨 Разбиваю на {total_parts} частей по 1000 сообщений")
print("⏳ Анализ займет 10-15 минут...")

for part_num in range(total_parts):
    start_idx = part_num * 1000
    end_idx = min((part_num + 1) * 1000, len(messages))
    part_messages = messages[start_idx:end_idx]
    
    print(f"\n📦 Часть {part_num + 1}/{total_parts} ({len(part_messages)} сообщений)...")
    
    # Подготавливаем промпт
    sample = part_messages[:30]  # Берём 30 сообщений для примера
    sample_text = "\n".join([f"{i+1}. {msg[:80]}..." for i, msg in enumerate(sample)])
    
    prompt = f"""Проанализируй обращения в поддержку банка и найди основные темы.

Инструкции:
1. Найди смысловые группы/темы обращений
2. Каждая тема должна быть развернутым описанием
3. Подсчитай примерное количество обращений по каждой теме
4. Верни результат ТОЛЬКО в формате JSON

Формат JSON:
{{
  "themes": [
    {{"name": "Название темы", "count": число, "description": "описание темы"}}
  ]
}}

Обращения для анализа (всего {len(part_messages)} сообщений):
{sample_text}
"""
    
    # Отправляем запрос
    result = ask_gigachat(prompt, token)
    
    if result is None:  # Токен истёк
        token = get_gigachat_token()
        if not token:
            break
        result = ask_gigachat(prompt, token)
    
    if result:
        try:
            answer = result['choices'][0]['message']['content']
            
            # Ищем JSON в ответе
            json_start = answer.find('{')
            json_end = answer.rfind('}') + 1
            
            if json_start != -1 and json_end > json_start:
                json_str = answer[json_start:json_end]
                data = json.loads(json_str)
                
                themes = data.get('themes', [])
                for theme in themes:
                    name = theme.get('name', '')
                    count = theme.get('count', 0)
                    
                    if name:
                        if name in all_themes:
                            all_themes[name]['count'] += count
                        else:
                            all_themes[name] = {
                                'count': count,
                                'description': theme.get('description', '')
                            }
                
                print(f"✅ Найдено {len(themes)} тем")
            else:
                print("⚠️ JSON не найден в ответе")
                
        except Exception as e:
            print(f"⚠️ Ошибка обработки: {e}")
    
    # Пауза между запросами
    if part_num < total_parts - 1:
        print("⏸️ Жду 3 секунды...")
        time.sleep(3)

# 8. Сохраняем результаты
print("\n📊 Формирую итоговый отчёт...")

if not all_themes:
    print("❌ Не удалось найти темы")
    input("Нажмите Enter для выхода...")
    exit()

# Сортируем по количеству обращений
sorted_themes = sorted(
    all_themes.items(),
    key=lambda x: x[1]['count'],
    reverse=True
)[:20]  # Топ-20

# Создаём папку для результатов
results_dir = "results"
os.makedirs(results_dir, exist_ok=True)

timestamp = datetime.now().strftime("%Y%m%d_%H%M")

# 9. Сохраняем в текстовый файл
txt_file = f"{results_dir}/report_{timestamp}.txt"
with open(txt_file, 'w', encoding='utf-8') as f:
    f.write("=" * 60 + "\n")
    f.write("ОТЧЁТ: ТОП-20 ТЕМ ОБРАЩЕНИЙ В ПОДДЕРЖКУ БАНКА\n")
    f.write("=" * 60 + "\n\n")
    
    f.write(f"Дата анализа: {datetime.now().strftime('%d.%m.%Y %H:%M')}\n")
    f.write(f"Всего проанализировано сообщений: {len(messages)}\n")
    f.write(f"Количество частей анализа: {total_parts}\n\n")
    
    f.write("ТЕМЫ ОБРАЩЕНИЙ:\n")
    f.write("-" * 60 + "\n\n")
    
    for i, (theme_name, theme_data) in enumerate(sorted_themes, 1):
        count = theme_data['count']
        description = theme_data['description']
        percentage = (count / len(messages) * 100) if messages else 0
        
        f.write(f"{i:2d}. {theme_name}\n")
        f.write(f"    📊 Количество обращений: {count} ({percentage:.1f}%)\n")
        if description:
            f.write(f"    📝 Описание: {description}\n")
        f.write("\n")

# 10. Сохраняем в JSON (для возможной дальнейшей обработки)
json_file = f"{results_dir}/data_{timestamp}.json"
with open(json_file, 'w', encoding='utf-8') as f:
    json.dump({
        "total_messages": len(messages),
        "total_parts": total_parts,
        "themes": [
            {
                "name": name,
                "count": data['count'],
                "description": data['description'],
                "percentage": (data['count'] / len(messages) * 100) if messages else 0
            }
            for name, data in sorted_themes
        ],
        "analysis_date": datetime.now().isoformat()
    }, f, ensure_ascii=False, indent=2)

print(f"\n✅ АНАЛИЗ ЗАВЕРШЁН!")
print(f"📁 Результаты сохранены в папке '{results_dir}/':")
print(f"   📄 {os.path.basename(txt_file)} - текстовый отчёт")
print(f"   📊 {os.path.basename(json_file)} - данные в JSON")
print(f"\n📧 Чтобы поделиться с коллегой:")
print(f"   1. Откройте файл '{txt_file}'")
print(f"   2. Скопируйте текст (Ctrl+A, Ctrl+C)")
print(f"   3. Отправьте коллеге в письме или чате")

print("\n" + "=" * 60)
input("Нажмите Enter для выхода...")
