import telebot
import pandas as pd
import os
from datetime import datetime  # ← ДОБАВИТЕ ЭТУ СТРОКУ


TOKEN = '8633982841:AAGtXkOnw3SAKjz_HaKR5IFYoaFEKn8e2ZA'
FILE_PATH = 'v.1.1.070326 Spreads Names 2026.xlsx'  # Файл в той же папке
SHEET_NAME = 'CME Data'     # ← ЗДЕСЬ УКАЖИТЕ НАЗВАНИЕ ВАШЕГО ЛИСТА
SHEET_NAME_2 = 'Spreads'
bot = telebot.TeleBot(TOKEN)


# Функция для форматирования даты
def format_date(date_value):
    if pd.isna(date_value):
        return ""
    try:
        if isinstance(date_value, str):
            date_obj = pd.to_datetime(date_value)
        else:
            date_obj = pd.to_datetime(date_value)
        return date_obj.strftime('%d.%m.%Y')  # 19.06.2026
    except:
        return str(date_value)  # Если не дата, оставляем как есть


# Загрузка данных
try:
    df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)
    print(f"✅ Загружен лист '{SHEET_NAME}', строк: {len(df)}")
except Exception as e:
    print(f"❌ Ошибка загрузки: {e}")
    df = pd.DataFrame()


@bot.message_handler(commands=['BuildSpread'])
def BuildSpread(message):
    bot.reply_to(message, 
        f'✅ Бот готов к работе!\n\n'
        f'📤 Формат ответа: MOEX - CME // Expiration of Spread\n\n'
        f'💬 Отправьте тикер MOEX вида NG-1.23 (в верхнем регистре)'
    )

@bot.message_handler(commands=['faq'])
def faq_command(message):
    faq_text = """
🔍 **ПОИСК ПО БД тикеров**

📋 **Как пользоваться:**
• /BuildSpread — запуск бота
•      Введите тикер MOEX (например, GOLD-6.26)
•      Пока придется учитывать регистр и вводить заглавные буквы
•      Получите все доступные вторые ноги
• /SpreadTView — запуск бота
•      Введите тикер MOEX (например, GOLD-6.26)
•      Пока придется учитывать регистр и вводить заглавные буквы
•      Получите код наиболее подходящего спреда для TView

📊 **Формат ответа:**
• Нога №1 MOEX - Нога №2 CME // Дата экспирации спреда
• Экспирацией будет считаться наиболее ранняя из 4х дат:
• • Дата экспирации на MOEX
• • Last Trade Date на CME
• • Settlement Date на CME
• • First Notice Date на CME
• Ноги на CME будут подбираться до даты экспирации на MOEX, но не позднее нее
• На CME за дату экспирации берется самая ранняя из 3х дат, указанных выше

❓ **Примеры:**
GOLD-6.26 → найдет точное совпадение по названия первой ноги и подберет доступные вторые ноги
"""
    bot.reply_to(message, faq_text)


@bot.message_handler(func=lambda msg: True)
def search(message):
    global df
    query = message.text.strip()
    
    # Номера столбцов (0=A, 8=I, 9=J, 10=K...)
    SEARCH_COL = 8   # Столбец I
    COL_I = 8        # I
    COL_J = 9        # J  
    COL_A = 0        # A
    COL_F = 5        # F
    COL_K = 10       # K
    
    # ТОЧНЫЙ поиск ВСЕХ совпадений в столбце I
    matches = df[df.iloc[:, SEARCH_COL].astype(str) == query]
    # matches = df[df.iloc[:, SEARCH_COL].astype(str).str.lower() == query.lower()]
    
    if not matches.empty:
        results = []
        for _, row in matches.iterrows():
            # Формируем строку: I - J // A - F // K
            i_val = format_date(row.iloc[8])
            j_val = format_date(row.iloc[9])
            a_val = format_date(row.iloc[0])
            f_val = format_date(row.iloc[5])
            k_val = format_date(row.iloc[10])
            
            result_line = f"{i_val} - {a_val} // {k_val}"
            results.append(result_line)
        
        # Отправляем результат
        response = f'✅ Найдено {len(results)} совпадений:\n\n'
        response += '\n'.join(results[:20])  # Максимум 20 строк
        
        if len(results) > 20:
            response += f'\n\n... и ещё {len(results)-20} строк'
            
        bot.reply_to(message, response)
    else:
        bot.reply_to(message, f'❌ Значение "{query}" не найдено')

if __name__ == '__main__':
    print("🚀 Telegram Excel Bot запущен!")
    bot.polling(none_stop=True)
