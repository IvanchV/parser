import openpyxl

# openpyxl используем только для считывания ексель файла

book = openpyxl.open("films_in_exel.xlsx", read_only=True)
sheet = book.active
dict_of_films = {}
title = []
rating = []
year = []
director = []
budget = []
general_fees = []
film_score = []
dict_of_films = {}
i = 0
for row in range(2, 52):
    title.append(sheet[row][0].value)
    rating.append(sheet[row][1].value)
    year.append(sheet[row][2].value)
    director.append(sheet[row][3].value)
    budget.append(sheet[row][4].value)
    general_fees.append(sheet[row][5].value)
    film_score.append(sheet[row][6].value)

# заполнение словаря фильмов
for item in title:
    dict_of_films[item] = {"название": title[i], "рейтинг": rating[i], "год": year[i], "режиссер": director[i],
                           "бюджет": budget[i], "общие сборы": general_fees[i], "оценка фильма": film_score[i]}
    i += 1


# функция для посика совпаденияо критериям
def intersection_list(list1, list2, list3, list4, list5, list6):
    return list(set(list1) & set(list2) & set(list3) & set(list4) & set(list5) & set(list6))


# основная функция поиска
def search(dict):
    list_of_criteria = input().split(',')
    search_list1 = list(dict.keys())
    search_list2 = list(dict.keys())
    search_list3 = list(dict.keys())
    search_list4 = list(dict.keys())
    search_list5 = list(dict.keys())
    search_list6 = list(dict.keys())

    for i in range(len(list_of_criteria)):
        list_of_criteria[i] = list_of_criteria[i].lower()

    for criteria in list_of_criteria:
        if criteria == "рейтинг":
            print("Не ниже какого рейтинга показать фильмы?")
            rating = int(input())
            search_list1.clear()
            for filmname, filmrating in dict.items():
                r = filmrating["рейтинг"]
                if r <= rating:
                    search_list1.append(f"{filmname}")

        if criteria == "год":
            print("Не позже какого года выхода  показать фильмы?")
            year = int(input())
            search_list2.clear()
            for filmname, filmyear in dict.items():
                y = filmyear["год"]
                if y >= year:
                    search_list2.append(f"{filmname}")

        if criteria == "режиссер":
            print("Вау! Ты знаешь режиссеров! Скорее напиши его имя и фамилию")
            name = input()
            search_list3.clear()
            for filmname, rname in dict.items():
                rn = rname["режиссер"]
                if name == rn:
                    search_list3.append(f"{filmname}")

        if criteria == "бюджет":
            print("Минимальный бюджет фильма:")
            money = int(input())
            search_list4.clear()
            for filmname, filmmoney in dict.items():
                fm = filmmoney["бюджет"]
                if fm >= money:
                    search_list4.append(f"{filmname}")

        if criteria == "общие сборы":
            print("Минимальные общие сборы:")
            fees = int(input())
            search_list5.clear()
            for filmname, filmfees in dict.items():
                gf = filmfees["общие сборы"]
                if gf >= fees:
                    search_list5.append(f"{filmname}")

        if criteria == "оценка фильма":
            print("Минимальная оценка фильма:")
            score = float(input())
            search_list6.clear()
            for filmname, filmscore in dict.items():
                sc = filmscore["оценка фильма"]
                if sc >= score:
                    search_list6.append(f"{filmname}")
    print("Результат поиска:")
    rez = intersection_list(search_list1, search_list2, search_list3, search_list4, search_list5, search_list6)
    if rez != []:
        for filmname in rez:
            print(f"{filmname} >>> {dict[filmname]}")
    else:
        print("Ничего не найдено")


# условный main
temp = True
while temp:
    print("Привет, давай начнем подбор, ты знаешь точное название фильма? да/нет")
    answer = input()
    try:
        if answer == "да":
            print("Название?")
            exact_name = input()
            print(f"{exact_name} >>> {dict_of_films[exact_name]}")
            print("Продолжить?да/нет")
            answer1 = input()
            if answer1 == "нет":
                temp = False
    except:
        print("Такого фильма нет")
    if temp:
        print("Не переживай, я помогу тебе подобрать фильм.")
        print('Список признаков: рейтинг, год, режиссер, бюджет, общие сборы, оценка фильма')
        print('Введите через запятую и без пробелов признаки, которые тебя интересуют(регистр неважен)')
        search(dict_of_films)
        print("Продолжить?да/нет")
        temp = input()
        if temp == "нет":
            temp = False
print("Приятного просмотра!")
