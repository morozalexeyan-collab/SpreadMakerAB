import telebot
import pandas as pd
import os
from datetime import datetime
import threading
from flask import Flask  # ← Health-check для Render
import time
import sys

TOKEN = os.getenv('TOKEN')  # ← ENV переменная!
FILE_PATH = 'v.1.1.070326 Spreads Names 2026.xlsx'
SHEET_NAME = 'CME Data'
SHEET_NAME_2 = 'Spreads'
bot = telebot.TeleBot(TOKEN)

# 🔥 FLASK HEALTH-CHECK (РЕШАЕТ 409 CONFLICT)
app = Flask(__name__)

@app.route('/')
@app.route('/health')
def health():
    return "OK"

def run_flask():
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)

flask_thread = threading.Thread(target=run_flask, daemon=True)
flask_thread.start()

# ГЛОБАЛЬНОЕ состояние пользователя
user_state = {}  # {user_id: 'spread' или 'tview'}

def format_date(date_value):
    if pd.isna(date_value):
        return ""
    try:
        date_obj = pd.to_datetime(date_value)
        return date_obj.strftime('%d.%m.%Y')
    except:
        return str(date_value)

# Загрузка данных ГЛОБАЛЬНО
try:
    df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)
    print(f"✅ '{SHEET_NAME}': {len(df)} строк")
except Exception as e:
    print(f"❌ Ошибка: {e}")
    df = pd.DataFrame()


@bot.message_handler(commands=['spread'])
def spread_start(message):
    user_id = message.from_user.id
    user_state[user_id] = 'spread'
    bot.reply_to(message, 
        f'✅ **РЕЖИМ SPREAD** (БД: {SHEET_NAME})\n\n'
        f'💬 Введите тикер MOEX (пример: NG-9.26)'
    )

@bot.message_handler(commands=['tview'])
def tview_start(message):
    user_id = message.from_user.id
    user_state[user_id] = 'tview'
    bot.reply_to(message, 
        f'✅ **РЕЖИМ TVIEW** (БД: {SHEET_NAME_2})\n\n'
        f'💬 Введите тикер MOEX (пример: NG-9.26)'
    )

@bot.message_handler(commands=['start', 'help', 'faq'])
def start_help(message):
    bot.reply_to(message, """
🔍 **SPREADMAKER AB**

📋 **Команды:**
/spread — поиск спредов (Moex/CME Data)
/tview — коды для TView (Spreads)
/start — это меню

📋 **Как пользоваться /spread:**
• /spread — запуск бота
• Введите тикер MOEX (например, NG-6.26)
• Получите все доступные вторые ноги

📊 **Формат ответа:**
• Нога №1 MOEX - Нога №2 CME // Дата экспирации спреда
• Экспирацией будет считаться наиболее ранняя из 4х дат:
• • Дата экспирации на MOEX
• • Last Trade Date на CME
• • Settlement Date на CME
• • First Notice Date на CME
• Ноги на CME будут подбираться до даты экспирации на MOEX, но не позднее нее
• На CME за дату экспирации берется самая ранняя из 3х дат, указанных выше


📋 **Как пользоваться /tview:**
• /tview — запуск бота
• Введите тикер MOEX (например, GOLD-6.26)
• Получите код наиболее подходящего спреда для TView

📊 **Формат ответа:**
• Код спреда // Дата экспирации спреда
    """)

@bot.message_handler(commands=['tickers'])
def start_tickers(message):
    bot.reply_to(message, """
🔍 **SPREADMAKER AB**

📋 **Список доступных тикеров:**
BR
BTC
ED
GOLD
NG
PLD
PLT
SILV
UCNY
NASD
SPYF                
    """)

@bot.message_handler(commands=['restart'])  # ← ЛЮБОЙ пользователь!
def restart_command(message):
    bot.reply_to(message, "🔄 Перезапуск бота для всех...")
    os.execv(sys.executable, ['python'] + sys.argv)

# 🎯 ОСНОВНОЙ ОБРАБОТЧИК — ловит ввод ПОСЛЕ команд
@bot.message_handler(func=lambda msg: True)
def handle_message(message):
    user_id = message.from_user.id
    query = message.text.strip().upper()  # 🔥 ИЗМЕНЕНИЕ 1: ВСЕГДА БОЛЬШИЕ БУКВЫ
    
    # Проверяем состояние пользователя
    if user_id in user_state:
        mode = user_state[user_id]
        
        if mode == 'spread':
            # 🔍 СПРЕДЫ (Лист 1)
            matches = df[df.iloc[:, 8].astype(str).str.upper() == query]  # 🔥 ИЗМЕНЕНИЕ 2
            if not matches.empty:
                results = []
                for _, row in matches.iterrows():
                    i_val = format_date(row.iloc[8])
                    a_val = format_date(row.iloc[0])
                    k_val = format_date(row.iloc[10])
                    result_line = f"{i_val} - {a_val} // {k_val}"
                    results.append(result_line)
                bot.reply_to(message, f'✅ *SPREAD* Найдено: {len(results)}\n' + '\n'.join(results[:20]))
            else:
                bot.reply_to(message, f'❌ "{message.text.strip()}" → "{query}" не найдено (Spread)')
            
        elif mode == 'tview':
            # 🔍 TVIEW (Лист 2)
            try:
                df2 = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME_2)
                matches2 = df2[df2.iloc[:, 0].astype(str).str.upper() == query]  # 🔥 ИГНОР РЕГИСТРА
                if not matches2.empty:
                    results2 = []
                    for _, row in matches2.iterrows():
                        i_val = format_date(row.iloc[5])
                        j_val = format_date(row.iloc[6])
                        line = f"{i_val} // {j_val}"
                        results2.append(line)
                    bot.reply_to(message, f'✅ *TVIEW* Найдено:({len(results2)}\n' + '\n'.join(results2))
                else:
                    bot.reply_to(message, f'❌ "{message.text.strip()}" → "{query}" не найдено (TView)')
            except Exception as e:
                bot.reply_to(message, f"❌ TVIEW ошибка")
        
        # Сбрасываем состояние после ответа
        del user_state[user_id]
    else:
        # Нет режима — показываем меню
        bot.reply_to(message, "❌ Неизвестная команда\n Выберите из доступных:\n\n📋 /spread или /tview")

if __name__ == '__main__':
    print("🚀 Telegram Excel Bot запущен!")
    bot.polling(none_stop=True, timeout=30)
