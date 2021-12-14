import telebot as tb
from telebot import types
import pandas as pd
import os

TOKEN = os.environ['TGTOKEN']
Characters_excel_path = 'Characters.xlsx'

#В рабочем режиме поставить False
debug_mode = False

bot = tb.TeleBot(TOKEN)
users = {}

rules_col_names = ['Параметр пользователя', 'Атрибут персонажа', 'Баллы']
rules_col_range = 'A:C'

class User:
    def __init__(self, chat_id):
        self.chars_df = None
        self.chat_id = chat_id
        self.sex = None
        self.age = None
        self.place_pref = None
        self.archetype_pref = ''
        self.blood_type = None
        self.zodiac_sign = None
        self.problems_pref = None
        self.secret_pref = None
        self.work_pref = None
        self.favourite = None
        self.growth = None
        self.popularity_pref = None
        self.hair_pref = None
        self.eyes_pref = None

def age_convert(age):
    if age >= 18:
        return '>=18'
    else:
        return '<18'

def growth_convert(growth):
    if growth <= 200 and growth >= 180:
        return '200-180'
    elif growth <= 179 and growth >= 160:
        return '179-160'
    elif growth <= 159 and growth >= 140:
        return '159-140'
    elif growth <= 139 and growth >= 120:
        return '139-120'
    else:
        return 'Other'

def award_characters(char_df: pd.DataFrame, rule_df: pd.DataFrame, user_param: str, char_attr: str) -> bool:
    tmp = rule_df[rule_df['Параметр пользователя']==user_param].loc[:, 'Атрибут персонажа':'Баллы']
    if len(tmp.index) < 1:
        return False
    for i in tmp.index:
        char_attr_value = rule_df.loc[i, 'Атрибут персонажа']
        score = rule_df.loc[i, 'Баллы']
        char_df.loc[char_df[char_attr]==char_attr_value, 'Rating'] += score
    return True

def send_characters_table(chat_id, char_df: pd.DataFrame, phrase: str):
    bot.send_message(chat_id, '{}\n{}'.format(phrase, char_df.sort_values('Rating', ascending=False)))

@bot.message_handler(func=lambda msg: msg.text != '/start')
def undef_text(message):
    bot.send_message(message.chat.id, "Простите, но мои создатели (Великолепная Мария Панова и неповторимый Константин Чиркин) не научили меня понимать такие команды. Если вы хотите начать опрос введите /start")

@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, 'Здравствуйте, я помогу вам с выбором 2д вайфу. 2д вайфу - это кумир/романтический персонаж в аниме (любого пола).')

    #Создаем нового пользователя
    users[message.chat.id] = User(message.chat.id)
    #Импортируем для этого пользователя персонажей
    # Переменные для описания таблицы с персонажами
    characters_col_range = 'A:Q'
    col_names = ['Имя', 'Фото', 'Группа крови', 'Цвет волос', 'Наличие тайны', 'Наличие магической способности',
                 'Архетип', 'Возраст', 'Любит', 'Рост', 'Наличие проблем', 'Место проживания', 'Наличие работы',
                 'Знак зодиака', 'Пол', 'Цвет глаз', 'Популярность']
    characters_number = 22
    characters = pd.read_excel(Characters_excel_path, sheet_name='Characters',
                               usecols=characters_col_range,
                               names=col_names,
                               nrows=characters_number,
                               converters={
                                   'Возраст': age_convert,
                                   'Рост': growth_convert
                               })
    characters['Rating'] = 0
    user = users[message.chat.id]
    user.chars_df = characters

    if debug_mode:
        print('Инициализирован пользователь. Импортирована таблица персонажей:')
        print(user.chars_df)

    #Спрашиваем у пользователя его пол
    keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
    keyboard.add('Мужской', 'Женский')
    msg = bot.send_message(message.chat.id, 'Начнем с первого вопроса: Укажите Ваш пол.', reply_markup=keyboard)

    bot.register_next_step_handler(msg, get_sex)

def get_sex(message):
    if message.text not in ['Мужской', 'Женский']:
        msg = bot.send_message(message.chat.id, 'Пол может принять только одно из двух значений: "Мужской" и "Женский". Укажите ваш пол заново.')
        bot.register_next_step_handler(msg, get_sex)
    else:
        user = users[message.chat.id]

        #Импорт таблицы с правилами ранжирования по полу
        sex_rules_start = 7
        sex_rules_num = 2
        sex_rules = pd.read_excel('Characters.xlsx', sheet_name='Rules',
                                  usecols=rules_col_range,
                                  names=rules_col_names,
                                  skiprows=sex_rules_start,
                                  nrows=sex_rules_num,
                                  )
        if debug_mode:
            print('Загружена таблица с правилами ранжирования по полу')
            print(sex_rules)

        user.sex = message.text

        chars_columns_mask = [True, False, False, False, False, False, False, False, False, False, False, False,
                              False, False, True, False, False, True]
        if award_characters(user.chars_df, sex_rules, user.sex, 'Пол'):
            send_characters_table(message.chat.id, user.chars_df.loc[:, chars_columns_mask],
                                  'Ваш пол: {}. Теперь персонажи отранжированы так:'.format(user.sex))

            msg = bot.send_message(message.chat.id, 'Теперь укажите ваш возраст (0-99).')
            bot.register_next_step_handler(msg, get_age)
        else:
            msg = bot.send_message(message.chat.id, 'Произошла внутренняя ошибка. Попробуйте позже')
            bot.register_next_step_handler(msg, get_sex)


def get_age(message):
    num = -1
    try:
        num = int(message.text)
    except Exception:
        msg = bot.send_message(message.chat.id,
                               'Возраст должен быть числом. Укажите ваш возраст заново.')
        bot.register_next_step_handler(msg, get_age)
        return

    if num < 0 or num > 99:
        msg = bot.send_message(message.chat.id,
                               'Возраст должен быть в пределах от 0 до 99. Укажите ваш возраст заново.')
        bot.register_next_step_handler(msg, get_age)
    else:
        user = users[message.chat.id]

        # Импорт таблицы с правилами ранжирования по полу
        age_rules_start = 10
        age_rules_num = 2
        age_rules = pd.read_excel('Characters.xlsx', sheet_name='Rules',
                                  usecols=rules_col_range,
                                  names=rules_col_names,
                                  skiprows=age_rules_start,
                                  nrows=age_rules_num,
                                  )
        if debug_mode:
            print('Загружена таблица с правилами ранжирования по возрасту')
            print(age_rules)

        if num >= 18:
            user.age = '>=18'
        else:
            user.age = '<18'

        chars_columns_mask = [True, False, False, False, False, False, False, True, False, False, False, False,
                              False, False, False, False, False, True]
        if award_characters(user.chars_df, age_rules, user.age, 'Возраст'):
            send_characters_table(message.chat.id, user.chars_df.loc[:, chars_columns_mask],
                                  'Ваш возраст: {}. Теперь персонажи отранжированы так:'.format(user.age))

            keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
            keyboard.add('Да', 'Нет')
            msg = bot.send_message(message.chat.id, 'Верите ли вы в существование других миров?', reply_markup=keyboard)
            bot.register_next_step_handler(msg, get_place_pref_1)
        else:
            msg = bot.send_message(message.chat.id, 'Произошла внутренняя ошибка. Попробуйте позже')
            bot.register_next_step_handler(msg, get_age)

def get_place_pref_1(message):
    keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
    if message.text not in ['Да', 'Нет']:
        msg = bot.send_message(message.chat.id,
                               'Ответ может принимать только одно из двух значений: "Да" и "Нет". Укажите ваш ответ заеово.')
        bot.register_next_step_handler(msg, get_place_pref_1)
        return
    elif message.text == 'Нет':
        keyboard.add('С Земли', 'С другой планеты')
    elif message.text == 'Да':
        keyboard.add('С Земли', 'С другой планеты', 'Из параллельного мира')
    msg = bot.send_message(message.chat.id, 'Откуда должна быть ваша вайфа?', reply_markup=keyboard)
    bot.register_next_step_handler(msg, get_place_pref_2)

def get_place_pref_2(message):
    if message.text not in ['С Земли', 'С другой планеты', 'Из параллельного мира']:
        msg = bot.send_message(message.chat.id,
                               'Ответ может принимать только одно из трех значений: "С Земли", "С другой планеты" и "Из параллельного мира". Укажите ваш ответ заеово.')
        bot.register_next_step_handler(msg, get_place_pref_2)
    else:
        user = users[message.chat.id]

        # Импорт таблицы с правилами ранжирования по полу
        place_pref_rules_start = 13
        place_pref_rules_num = 3
        place_pref_rules = pd.read_excel('Characters.xlsx', sheet_name='Rules',
                                  usecols=rules_col_range,
                                  names=rules_col_names,
                                  skiprows=place_pref_rules_start,
                                  nrows=place_pref_rules_num,
                                  )
        if debug_mode:
            print('Загружена таблица с правилами ранжирования по предпочтению по месту проживания')
            print(place_pref_rules)

        if message.text == 'С Земли':
            user.place_pref = 'Земля'
        elif message.text == 'С другой планеты':
            user.place_pref = 'Другая планета'
        elif message.text == 'Из параллельного мира':
            user.place_pref = 'Параллельный мир'

        chars_columns_mask = [True, False, False, False, False, False, False, False, False, False, False, True,
                              False, False, False, False, False, True]
        if award_characters(user.chars_df, place_pref_rules, user.place_pref, 'Место проживания'):
            send_characters_table(message.chat.id, user.chars_df.loc[:, chars_columns_mask],
                                  'Ваше предпочтение по месту проживания персонажа: {}. Теперь персонажи отранжированы так:'.format(user.place_pref))
            keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
            keyboard.add('Да', 'Нет')
            msg = bot.send_message(message.chat.id, 'Активный ли вы человек?', reply_markup=keyboard)
            bot.register_next_step_handler(msg, get_archetype_pref_1)
        else:
            msg = bot.send_message(message.chat.id, 'Произошла внутренняя ошибка. Попробуйте позже')
            bot.register_next_step_handler(msg, get_place_pref_2)

def get_archetype_pref_1(message):
    if message.text not in ['Да', 'Нет']:
        msg = bot.send_message(message.chat.id,
                               'Ответ может принимать только одно из двух значений: "Да", "Нет". Укажите ваш ответ заеово.')
        bot.register_next_step_handler(msg, get_archetype_pref_1)
    else:
        user = users[message.chat.id]

        if message.text == 'Да':
            user.archetype_pref += 'Активный'
        elif message.text == 'Нет':
            user.archetype_pref += 'Не активный'

        keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
        keyboard.add('Да', 'Нет')
        msg = bot.send_message(message.chat.id, 'Храбрый ли вы человек?', reply_markup=keyboard)
        bot.register_next_step_handler(msg, get_archetype_pref_2)

def get_archetype_pref_2(message):
    if message.text not in ['Да', 'Нет']:
        msg = bot.send_message(message.chat.id,
                               'Ответ может принимать только одно из двух значений: "Да", "Нет". Укажите ваш ответ заеово.')
        bot.register_next_step_handler(msg, get_archetype_pref_2)
    else:
        user = users[message.chat.id]

        if message.text == 'Да':
            user.archetype_pref += ', Храбрый'
        elif message.text == 'Нет':
            user.archetype_pref += ', Не храбрый'

        keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
        keyboard.add('Да', 'Нет')
        msg = bot.send_message(message.chat.id, 'Нравятся ли вам персонажи со странностями?', reply_markup=keyboard)
        bot.register_next_step_handler(msg, get_archetype_pref_3)

def get_archetype_pref_3(message):
    if message.text not in ['Да', 'Нет']:
        msg = bot.send_message(message.chat.id,
                               'Ответ может принимать только одно из двух значений: "Да", "Нет". Укажите ваш ответ заеово.')
        bot.register_next_step_handler(msg, get_archetype_pref_3)
    else:
        user = users[message.chat.id]

        if message.text == 'Да':
            user.archetype_pref += ', Странный'
        elif message.text == 'Нет':
            user.archetype_pref += ', Не странный'

        if debug_mode:
            print(user.archetype_pref)

        # Импорт таблицы с правилами ранжирования по архетипу
        archetype_rules_start = 17
        archetype_rules_num = 13
        archetype_rules = pd.read_excel('Characters.xlsx', sheet_name='Rules',
                                  usecols=rules_col_range,
                                  names=rules_col_names,
                                  skiprows=archetype_rules_start,
                                  nrows=archetype_rules_num,
                                  )
        if debug_mode:
            print('Загружена таблица с правилами ранжирования по архетипу')
            print(archetype_rules)

        chars_columns_mask = [True, False, False, False, False, False, True, False, False, False, False, False,
                              False, False, False, False, False, True]
        if award_characters(user.chars_df, archetype_rules, user.archetype_pref, 'Архетип'):
            send_characters_table(message.chat.id, user.chars_df.loc[:, chars_columns_mask],
                                  'Ваше предпочтение по архетипу персонажа: {}. Теперь персонажи отранжированы так:'.format(user.archetype_pref))
            keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
            keyboard.add('I', 'II', 'III', 'IV')
            msg = bot.send_message(message.chat.id, 'Какая у вас группа крови?', reply_markup=keyboard)
            bot.register_next_step_handler(msg, get_blood_type)
        else:
            msg = bot.send_message(message.chat.id, 'Произошла внутренняя ошибка. Попробуйте позже')
            bot.register_next_step_handler(msg, get_archetype_pref_3)

def get_blood_type(message):
    if message.text not in ['I', 'II', 'III', 'IV']:
        msg = bot.send_message(message.chat.id,
                               'Группа крови может принимать только одно из четырех значений: "I", "II", "III", "IV". Укажите вашу группу крови заново.')
        bot.register_next_step_handler(msg, get_blood_type)
    else:
        user = users[message.chat.id]
        user.blood_type = message.text

        # Импорт таблицы с правилами ранжирования по архетипу
        blt_rules_start = 31
        blt_rules_num = 10
        blt_rules = pd.read_excel('Characters.xlsx', sheet_name='Rules',
                                        usecols=rules_col_range,
                                        names=rules_col_names,
                                        skiprows=blt_rules_start,
                                        nrows=blt_rules_num,
                                        )
        if debug_mode:
            print('Загружена таблица с правилами ранжирования по архетипу')
            print(blt_rules)

        chars_columns_mask = [True, False, True, False, False, False, False, False, False, False, False, False,
                              False, False, False, False, False, True]
        if award_characters(user.chars_df, blt_rules, user.blood_type, 'Группа крови'):
            send_characters_table(message.chat.id, user.chars_df.loc[:, chars_columns_mask],
                                  'Ваша группа крови: {}. Теперь персонажи отранжированы так:'.format(
                                      user.blood_type))
            keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
            keyboard.add('Овен', 'Телец', 'Близнецы', 'Рак', 'Лев', 'Дева', 'Весы', 'Скорпион', 'Стрелец',
                         'Козерог', 'Водолей', 'Рыбы')
            msg = bot.send_message(message.chat.id, 'Какой у вас знак зодиака?', reply_markup=keyboard)
            bot.register_next_step_handler(msg, get_zodiac_sign)
        else:
            msg = bot.send_message(message.chat.id, 'Произошла внутренняя ошибка. Попробуйте позже')
            bot.register_next_step_handler(msg, get_blood_type)

def get_zodiac_sign(message):
    if message.text not in ['Овен', 'Телец', 'Близнецы', 'Рак', 'Лев', 'Дева', 'Весы', 'Скорпион', 'Стрелец',
                         'Козерог', 'Водолей', 'Рыбы']:
        msg = bot.send_message(message.chat.id,
                               'Знак зодиака может принимать только одно из следующих значений: "Овен", "Телец", '
                               '"Близнецы", "Рак", "Лев", "Дева", "Весы", "Скорпион", "Стрелец", "Козерог", "Водолей", '
                               '"Рыбы". Укажите вашу группу крови заново.')
        bot.register_next_step_handler(msg, get_zodiac_sign)
    else:
        user = users[message.chat.id]
        user.zodiac_sign = message.text

        # Импорт таблицы с правилами ранжирования по архетипу
        zs_rules_start = 42
        zs_rules_num = 39
        zs_rules = pd.read_excel('Characters.xlsx', sheet_name='Rules',
                                  usecols=rules_col_range,
                                  names=rules_col_names,
                                  skiprows=zs_rules_start,
                                  nrows=zs_rules_num,
                                  )
        if debug_mode:
            print('Загружена таблица с правилами ранжирования по знаку зодиака')
            print(zs_rules)

        chars_columns_mask = [True, False, False, False, False, False, False, False, False, False, False, False,
                              False, True, False, False, False, True]
        if award_characters(user.chars_df, zs_rules, user.zodiac_sign, 'Знак зодиака'):
            send_characters_table(message.chat.id, user.chars_df.loc[:, chars_columns_mask],
                                  'Ваш знак зодиака: {}. Теперь персонажи отранжированы так:'.format(
                                      user.zodiac_sign))
            keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
            keyboard.add('Да', 'Нет')
            msg = bot.send_message(message.chat.id, 'Любите ли вы решать чужие проблемы?', reply_markup=keyboard)
            bot.register_next_step_handler(msg, get_problems_pref)
        else:
            msg = bot.send_message(message.chat.id, 'Произошла внутренняя ошибка. Попробуйте позже')
            bot.register_next_step_handler(msg, get_zodiac_sign)

def get_problems_pref(message):
    if message.text not in ['Да', 'Нет']:
        msg = bot.send_message(message.chat.id,
                               'Ответ может принимать только одно из двух значений: "Да" и "Нет". Укажите вашу ответ заново.')
        bot.register_next_step_handler(msg, get_problems_pref)
    else:
        user = users[message.chat.id]
        if message.text == "Да":
            user.problems_pref = "Есть"
        elif message.text == "Нет":
            user.problems_pref = "Нет"

        # Импорт таблицы с правилами ранжирования по архетипу
        prp_rules_start = 82
        prp_rules_num = 2
        prp_rules = pd.read_excel('Characters.xlsx', sheet_name='Rules',
                                 usecols=rules_col_range,
                                 names=rules_col_names,
                                 skiprows=prp_rules_start,
                                 nrows=prp_rules_num,
                                 )
        if debug_mode:
            print('Загружена таблица с правилами ранжирования по наличию проблем')
            print(prp_rules)

        chars_columns_mask = [True, False, False, False, False, False, False, False, False, False, True, False,
                              False, False, False, False, False, True]
        if award_characters(user.chars_df, prp_rules, user.problems_pref, 'Наличие проблем'):
            send_characters_table(message.chat.id, user.chars_df.loc[:, chars_columns_mask],
                                  'Ваше предпочтение по наличию проблем: {}. Теперь персонажи отранжированы так:'.format(
                                      user.problems_pref))
            keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
            keyboard.add('Да', 'Нет')
            msg = bot.send_message(message.chat.id, 'Умеете ли вы хранить чужие секреты?', reply_markup=keyboard)
            bot.register_next_step_handler(msg, get_secret_pref)
        else:
            msg = bot.send_message(message.chat.id, 'Произошла внутренняя ошибка. Попробуйте позже')
            bot.register_next_step_handler(msg, get_problems_pref)

def get_secret_pref(message):
    if message.text not in ['Да', 'Нет']:
        msg = bot.send_message(message.chat.id,
                               'Ответ может принимать только одно из двух значений: "Да" и "Нет". Укажите вашу ответ заново.')
        bot.register_next_step_handler(msg, get_secret_pref)
    else:
        user = users[message.chat.id]
        if message.text == "Да":
            user.secret_pref = "Есть"
        elif message.text == "Нет":
            user.secret_pref = "Нет"

        # Импорт таблицы с правилами ранжирования по архетипу
        sp_rules_start = 85
        sp_rules_num = 2
        sp_rules = pd.read_excel('Characters.xlsx', sheet_name='Rules',
                                  usecols=rules_col_range,
                                  names=rules_col_names,
                                  skiprows=sp_rules_start,
                                  nrows=sp_rules_num,
                                  )
        if debug_mode:
            print('Загружена таблица с правилами ранжирования по наличию тайны')
            print(sp_rules)

        chars_columns_mask = [True, False, False, False, True, False, False, False, False, False, False, False,
                              False, False, False, False, False, True]
        if award_characters(user.chars_df, sp_rules, user.secret_pref, 'Наличие тайны'):
            send_characters_table(message.chat.id, user.chars_df.loc[:, chars_columns_mask],
                                  'Ваше предпочтение по наличию тайны: {}. Теперь персонажи отранжированы так:'.format(
                                      user.secret_pref))
            keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
            keyboard.add('Да', 'Нет')
            msg = bot.send_message(message.chat.id, 'Считаете ли вы, что ваша вайфу должна работать?', reply_markup=keyboard)
            bot.register_next_step_handler(msg, get_work_pref)
        else:
            msg = bot.send_message(message.chat.id, 'Произошла внутренняя ошибка. Попробуйте позже')
            bot.register_next_step_handler(msg, get_secret_pref)

def get_work_pref(message):
    if message.text not in ['Да', 'Нет']:
        msg = bot.send_message(message.chat.id,
                               'Ответ может принимать только одно из двух значений: "Да" и "Нет". Укажите вашу ответ заново.')
        bot.register_next_step_handler(msg, get_work_pref)
    else:
        user = users[message.chat.id]
        if message.text == "Да":
            user.work_pref = "Есть"
        elif message.text == "Нет":
            user.work_pref = "Нет"

        # Импорт таблицы с правилами ранжирования по архетипу
        wp_rules_start = 91
        wp_rules_num = 2
        wp_rules = pd.read_excel('Characters.xlsx', sheet_name='Rules',
                                 usecols=rules_col_range,
                                 names=rules_col_names,
                                 skiprows=wp_rules_start,
                                 nrows=wp_rules_num,
                                 )
        if debug_mode:
            print('Загружена таблица с правилами ранжирования по наличию тайны')
            print(wp_rules)

        chars_columns_mask = [True, False, False, False, False, False, False, False, False, False, False, False,
                              True, False, False, False, False, True]
        if award_characters(user.chars_df, wp_rules, user.work_pref, 'Наличие работы'):
            send_characters_table(message.chat.id, user.chars_df.loc[:, chars_columns_mask],
                                  'Ваше предпочтение по наличию работы: {}. Теперь персонажи отранжированы так:'.format(
                                      user.work_pref))
            keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
            keyboard.add('Да', 'Нет')
            msg = bot.send_message(message.chat.id, 'Нравится ли вам активный отдых?',
                                   reply_markup=keyboard)
            bot.register_next_step_handler(msg, get_favourite_1)
        else:
            msg = bot.send_message(message.chat.id, 'Произошла внутренняя ошибка. Попробуйте позже')
            bot.register_next_step_handler(msg, get_work_pref)

def get_favourite_1(message):
    if message.text not in ['Да', 'Нет']:
        msg = bot.send_message(message.chat.id,
                               'Ответ может принимать только одно из двух значений: "Да" и "Нет". Укажите ваш ответ заново.')
        bot.register_next_step_handler(msg, get_favourite_1)
    else:
        if message.text == 'Да':
            user = users[message.chat.id]
            user.favourite = 'Активный отдых'

            # Импорт таблицы с правилами ранжирования по архетипу
            fv_rules_start = 94
            fv_rules_num = 21
            fv_rules = pd.read_excel('Characters.xlsx', sheet_name='Rules',
                                     usecols=rules_col_range,
                                     names=rules_col_names,
                                     skiprows=fv_rules_start,
                                     nrows=fv_rules_num,
                                     )
            if debug_mode:
                print('Загружена таблица с правилами ранжирования по наличию тайны')
                print(fv_rules)

            chars_columns_mask = [True, False, False, False, False, False, False, False, True, False, False, False,
                                  False, False, False, False, False, True]
            if award_characters(user.chars_df, fv_rules, user.favourite, 'Любит'):
                send_characters_table(message.chat.id, user.chars_df.loc[:, chars_columns_mask],
                                      'Вы любите: {}. Теперь персонажи отранжированы так:'.format(
                                          user.favourite))
                msg = bot.send_message(message.chat.id, 'Укажите ваш рост')
                bot.register_next_step_handler(msg, get_growth)
            else:
                msg = bot.send_message(message.chat.id, 'Произошла внутренняя ошибка. Попробуйте позже')
                bot.register_next_step_handler(msg, get_favourite_1)
        else:
            keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
            keyboard.add('Да', 'Нет')
            msg = bot.send_message(message.chat.id, 'Любите ли вы поесть?', reply_markup=keyboard)
            bot.register_next_step_handler(msg, get_favourite_2)

def get_favourite_2(message):
    if message.text not in ['Да', 'Нет']:
        msg = bot.send_message(message.chat.id,
                               'Ответ может принимать только одно из двух значений: "Да" и "Нет". Укажите ваш ответ заново.')
        bot.register_next_step_handler(msg, get_favourite_2)
    else:
        if message.text == 'Да':
            user = users[message.chat.id]
            user.favourite = 'Поесть'

            # Импорт таблицы с правилами ранжирования по архетипу
            fv_rules_start = 94
            fv_rules_num = 21
            fv_rules = pd.read_excel('Characters.xlsx', sheet_name='Rules',
                                     usecols=rules_col_range,
                                     names=rules_col_names,
                                     skiprows=fv_rules_start,
                                     nrows=fv_rules_num,
                                     )
            if debug_mode:
                print('Загружена таблица с правилами ранжирования по наличию тайны')
                print(fv_rules)

            chars_columns_mask = [True, False, False, False, False, False, False, False, True, False, False, False,
                                  False, False, False, False, False, True]
            if award_characters(user.chars_df, fv_rules, user.favourite, 'Любит'):
                send_characters_table(message.chat.id, user.chars_df.loc[:, chars_columns_mask],
                                      'Вы любите: {}. Теперь персонажи отранжированы так:'.format(
                                          user.favourite))
                msg = bot.send_message(message.chat.id, 'Укажите ваш рост')
                bot.register_next_step_handler(msg, get_growth)
            else:
                msg = bot.send_message(message.chat.id, 'Произошла внутренняя ошибка. Попробуйте позже')
                bot.register_next_step_handler(msg, get_favourite_1)
        else:
            keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
            keyboard.add('Да', 'Нет')
            msg = bot.send_message(message.chat.id, 'Нравится ли вам проводить время дома, завернувшись в плед?', reply_markup=keyboard)
            bot.register_next_step_handler(msg, get_favourite_3)

def get_favourite_3(message):
    if message.text not in ['Да', 'Нет']:
        msg = bot.send_message(message.chat.id,
                               'Ответ может принимать только одно из двух значений: "Да" и "Нет". Укажите ваш ответ заново.')
        bot.register_next_step_handler(msg, get_favourite_2)
    else:
        if message.text == 'Да':
            user = users[message.chat.id]
            user.favourite = 'Плед'

            # Импорт таблицы с правилами ранжирования по архетипу
            fv_rules_start = 94
            fv_rules_num = 21
            fv_rules = pd.read_excel('Characters.xlsx', sheet_name='Rules',
                                     usecols=rules_col_range,
                                     names=rules_col_names,
                                     skiprows=fv_rules_start,
                                     nrows=fv_rules_num,
                                     )
            if debug_mode:
                print('Загружена таблица с правилами ранжирования по наличию тайны')
                print(fv_rules)

            chars_columns_mask = [True, False, False, False, False, False, False, False, True, False, False, False,
                                  False, False, False, False, False, True]
            if award_characters(user.chars_df, fv_rules, user.favourite, 'Любит'):
                send_characters_table(message.chat.id, user.chars_df.loc[:, chars_columns_mask],
                                      'Вы любите: {}. Теперь персонажи отранжированы так:'.format(
                                          user.favourite))
                msg = bot.send_message(message.chat.id, 'Укажите ваш рост')
                bot.register_next_step_handler(msg, get_growth)
            else:
                msg = bot.send_message(message.chat.id, 'Произошла внутренняя ошибка. Попробуйте позже')
                bot.register_next_step_handler(msg, get_favourite_1)
        else:
            keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
            keyboard.add('Да')
            msg = bot.send_message(message.chat.id, 'Любите ли вы проводить время с друзьями?',
                                   reply_markup=keyboard)
            bot.register_next_step_handler(msg, get_favourite_4)

def get_favourite_4(message):
    if message.text not in ['Да']:
        msg = bot.send_message(message.chat.id,
                               'У вас не осталось вариантов). Укажите ваш ответ заново.')
        bot.register_next_step_handler(msg, get_favourite_2)
    else:
        if message.text == 'Да':
            user = users[message.chat.id]
            user.favourite = 'Друзья'

            # Импорт таблицы с правилами ранжирования по архетипу
            fv_rules_start = 94
            fv_rules_num = 21
            fv_rules = pd.read_excel('Characters.xlsx', sheet_name='Rules',
                                     usecols=rules_col_range,
                                     names=rules_col_names,
                                     skiprows=fv_rules_start,
                                     nrows=fv_rules_num,
                                     )
            if debug_mode:
                print('Загружена таблица с правилами ранжирования по наличию тайны')
                print(fv_rules)

            chars_columns_mask = [True, False, False, False, False, False, False, False, True, False, False, False,
                                  False, False, False, False, False, True]
            if award_characters(user.chars_df, fv_rules, user.favourite, 'Любит'):
                send_characters_table(message.chat.id, user.chars_df.loc[:, chars_columns_mask],
                                      'Вы любите: {}. Теперь персонажи отранжированы так:'.format(
                                          user.favourite))
                msg = bot.send_message(message.chat.id, 'Укажите ваш рост')
                bot.register_next_step_handler(msg, get_growth)
            else:
                msg = bot.send_message(message.chat.id, 'Произошла внутренняя ошибка. Попробуйте позже')
                bot.register_next_step_handler(msg, get_favourite_1)

def get_growth(message):
    num = -1
    try:
        num = int(message.text)
    except Exception:
        msg = bot.send_message(message.chat.id,
                               'Рост должен быть числом. Укажите ваш рост заново.')
        bot.register_next_step_handler(msg, get_growth)
        return

    if num < 120 or num > 200:
        msg = bot.send_message(message.chat.id,
                               'Рост должен быть в пределах от 120 до 200. Укажите ваш рост заново.')
        bot.register_next_step_handler(msg, get_growth)
    else:
        user = users[message.chat.id]

        # Импорт таблицы с правилами ранжирования по полу
        growth_rules_start = 116
        growth_rules_num = 4
        growth_rules = pd.read_excel('Characters.xlsx', sheet_name='Rules',
                                  usecols=rules_col_range,
                                  names=rules_col_names,
                                  skiprows=growth_rules_start,
                                  nrows=growth_rules_num,
                                  )
        if debug_mode:
            print('Загружена таблица с правилами ранжирования по росту')
            print(growth_rules)

        if num <= 200 and num >= 180:
            user.growth = '200-180'
        elif num <= 179 and num >= 160:
            user.growth = '179-160'
        elif num <= 159 and num >= 140:
            user.growth = '159-140'
        elif num <= 139 and num >= 120:
            user.growth = '139-120'

        chars_columns_mask = [True, False, False, False, False, False, False, False, False, True, False, False,
                              False, False, False, False, False, True]
        if award_characters(user.chars_df, growth_rules, user.growth, 'Рост'):
            send_characters_table(message.chat.id, user.chars_df.loc[:, chars_columns_mask],
                                  'Ваш рост: {}. Теперь персонажи отранжированы так:'.format(user.growth))

            keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
            keyboard.add('Да', 'Нет')
            msg = bot.send_message(message.chat.id, 'Ревнивый ли вы человек?', reply_markup=keyboard)
            bot.register_next_step_handler(msg, get_popularity_pref)
        else:
            msg = bot.send_message(message.chat.id, 'Произошла внутренняя ошибка. Попробуйте позже')
            bot.register_next_step_handler(msg, get_growth)

def get_popularity_pref(message):
    if message.text not in ['Да', 'Нет']:
        msg = bot.send_message(message.chat.id,
                               'Ответ может принимать только одно из двух значений: "Да" и "Нет". Укажите вашу ответ заново.')
        bot.register_next_step_handler(msg, get_popularity_pref)
    else:
        user = users[message.chat.id]
        if message.text == 'Да':
            user.popularity_pref = 'Нет'
        elif message.text == 'Нет':
            user.popularity_pref = 'Да'

        # Импорт таблицы с правилами ранжирования по архетипу
        popp_rules_start = 121
        popp_rules_num = 2
        popp_rules = pd.read_excel('Characters.xlsx', sheet_name='Rules',
                                  usecols=rules_col_range,
                                  names=rules_col_names,
                                  skiprows=popp_rules_start,
                                  nrows=popp_rules_num,
                                  )
        if debug_mode:
            print('Загружена таблица с правилами ранжирования по популярности')
            print(popp_rules)

        chars_columns_mask = [True, False, False, False, False, False, False, False, False, False, False, False,
                              False, False, False, False, True, True]
        if award_characters(user.chars_df, popp_rules, user.popularity_pref, 'Популярность'):
            send_characters_table(message.chat.id, user.chars_df.loc[:, chars_columns_mask],
                                  'Ваше предпочтение по популярности: {}. Теперь персонажи отранжированы так:'.format(
                                      user.popularity_pref))
            keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
            keyboard.add('Да', 'Нет')
            msg = bot.send_message(message.chat.id, 'Нравится ли вам яркая внешность?', reply_markup=keyboard)
            bot.register_next_step_handler(msg, get_hair_pref)
        else:
            msg = bot.send_message(message.chat.id, 'Произошла внутренняя ошибка. Попробуйте позже')
            bot.register_next_step_handler(msg, get_popularity_pref)

def get_hair_pref(message):
    if message.text not in ['Да', 'Нет']:
        msg = bot.send_message(message.chat.id,
                               'Ответ может принимать только одно из двух значений: "Да" и "Нет". Укажите вашу ответ заново.')
        bot.register_next_step_handler(msg, get_hair_pref)
    else:
        user = users[message.chat.id]
        user.hair_pref = message.text

        # Импорт таблицы с правилами ранжирования по цвету волос
        hair_rules_start = 124
        hair_rules_num = 9
        hair_rules = pd.read_excel('Characters.xlsx', sheet_name='Rules',
                                   usecols=rules_col_range,
                                   names=rules_col_names,
                                   skiprows=hair_rules_start,
                                   nrows=hair_rules_num,
                                   )
        if debug_mode:
            print('Загружена таблица с правилами ранжирования по наличию тайны')
            print(hair_rules)

        chars_columns_mask = [True, False, False, True, False, False, False, False, False, False, False, False,
                              False, False, False, False, False, True]
        if award_characters(user.chars_df, hair_rules, user.hair_pref, 'Цвет волос'):
            send_characters_table(message.chat.id, user.chars_df.loc[:, chars_columns_mask],
                                  'Вам нравятся яркие цвета волос: {}. Теперь персонажи отранжированы так:'.format(
                                      user.hair_pref))
            keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True)
            keyboard.add('Голубые', 'Карие', 'Розовые', 'Синие', 'Серые', 'Зеленые')
            msg = bot.send_message(message.chat.id, 'Какие глаза вам нравятся?',
                                   reply_markup=keyboard)
            bot.register_next_step_handler(msg, get_eyes_pref)
        else:
            msg = bot.send_message(message.chat.id, 'Произошла внутренняя ошибка. Попробуйте позже')
            bot.register_next_step_handler(msg, get_hair_pref)

def get_eyes_pref(message):
    if message.text not in ['Голубые', 'Карие', 'Розовые', 'Синие', 'Серые', 'Зеленые']:
        msg = bot.send_message(message.chat.id,
                               "Ответ может принимать только одно из значений: 'Голубые', 'Карие', 'Розовые', 'Синие', 'Серые', 'Зеленые'. Укажите вашу ответ заново.")
        bot.register_next_step_handler(msg, get_eyes_pref)
    else:
        user = users[message.chat.id]
        user.eyes_pref = message.text

        # Импорт таблицы с правилами ранжирования по цвету волос
        eyes_rules_start = 134
        eyes_rules_num = 6
        eyes_rules = pd.read_excel('Characters.xlsx', sheet_name='Rules',
                                   usecols=rules_col_range,
                                   names=rules_col_names,
                                   skiprows=eyes_rules_start,
                                   nrows=eyes_rules_num,
                                   )
        if debug_mode:
            print('Загружена таблица с правилами ранжирования по наличию тайны')
            print(eyes_rules)

        chars_columns_mask = [True, False, False, False, False, False, False, False, False, False, False, False,
                              False, False, False, True, False, True]
        if award_characters(user.chars_df, eyes_rules, user.eyes_pref, 'Цвет глаз'):
            send_characters_table(message.chat.id, user.chars_df.loc[:, chars_columns_mask],
                                  'Вам нравятся глаза: {}. Теперь персонажи отранжированы так:'.format(
                                      user.eyes_pref))
            send_best_wifu(message)
        else:
            msg = bot.send_message(message.chat.id, 'Произошла внутренняя ошибка. Попробуйте позже')
            bot.register_next_step_handler(msg, get_eyes_pref)

def send_best_wifu(message):
    user = users[message.chat.id]

    best_wifu = user.chars_df.sort_values('Rating', ascending=False)
    best_wifu_name = best_wifu.iloc[0, 0]
    best_wifu_photo_path = os.path.join(os.getcwd(), 'src', best_wifu.iloc[0, 1])

    with open(best_wifu_photo_path, 'rb') as best_wifu_photo:
        bot.send_animation(message.chat.id, animation=best_wifu_photo, caption='Ура, вы дошли до конца опроса. Ваш(а) вайфу: {}'.format(best_wifu_name))

print('bot starting...')
bot.polling(none_stop=True)