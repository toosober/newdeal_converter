# newdeal_converter

Пока работает в режиме утилиты в формате

``$ python main.py имя_эксель_файла > out``

Либо можно использовать ``Parser().parse_document``, как это показано в tests/test_main.py

Замечания Дмитрия оставляю:

Примерный план разработки

0) Что-то делать с ебучими сносками

1) Достать данные когда оно обновлено
2) Достать данные что это за таблица
3) Достать данные в каких это ценах
4) Достать единицы измерения
5) Узнать какая строчка годы
6) Достать названия счетов
7) Достать категории (ресурсы/использование)
8) Достать сами категории, так назвать файл и записать в формате
год
значение
