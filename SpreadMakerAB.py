import telebot
import pandas as pd
import os

TOKEN = '8633982841:AAGtXkOnw3SAKjz_HaKR5IFYoaFEKn8e2ZA'
FILE_PATH = 'v.1.1.070326 Spreads Names 2026.xlsm'  # Файл в той же папке
SHEET_NAME = 'CME Data'     # ← ЗДЕСЬ УКАЖИТЕ НАЗВАНИЕ ВАШЕГО ЛИСТА
bot = telebot.TeleBot(TOKEN)

# Загрузка данных
try:
    df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)
    print(f"✅ Загружен лист '{SHEET_NAME}', строк: {len(df)}")
except Exception as e:
    print(f"❌ Ошибка загрузки: {e}")
    df = pd.DataFrame()

@bot.message_handler(commands=['start'])
def start(message):
    bot.reply_to(message, 
        f'✅ Бот готов к работе!\n\n'
        f'📤 Формат ответа: I - J // A - F // K\n'
        f'📈 Строк в таблице: {len(df)}\n\n'
        f'💬 Отправьте тикер MOEX вида NG-1.23'
    )

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
    
    if not matches.empty:
        results = []
        for _, row in matches.iterrows():
            # Формируем строку: I - J // A - F // K
            i_val = str(row.iloc[COL_I]) if pd.notna(row.iloc[COL_I]) else ''
            j_val = str(row.iloc[COL_J]) if pd.notna(row.iloc[COL_J]) else ''
            a_val = str(row.iloc[COL_A]) if pd.notna(row.iloc[COL_A]) else ''
            f_val = str(row.iloc[COL_F]) if pd.notna(row.iloc[COL_F]) else ''
            k_val = str(row.iloc[COL_K]) if pd.notna(row.iloc[COL_K]) else ''
            
            result_line = f"{i_val} - {j_val} // {a_val} - {f_val} // {k_val}"
            results.append(result_line)
        
        # Отправляем результат
        response = f'✅ Найдено {len(results)} совпадений:\n\n'
        response += '\n'.join(results[:20])  # Максимум 20 строк
        
        if len(results) > 20:
            response += f'\n\n... и ещё {len(results)-20} строк'
            
        bot.reply_to(message, response)
    else:
        bot.reply_to(message, f'❌ Значение "{query}" не найдено в столбце I')

if __name__ == '__main__':
    print("🚀 Telegram Excel Bot запущен!")
    bot.polling(none_stop=True)
