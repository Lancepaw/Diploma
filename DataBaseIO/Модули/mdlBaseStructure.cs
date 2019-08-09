using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataBaseIO
{
    class mdlBaseStructure
    {
        //bit - логический тип
        //binary - числовой двоичный
        //string - текстовый (255 знаков)
        //memo и text - поле MEMO
        //datetime - дата/время

        public static string[][][] masTabNames = new string[][][] 
        { 
            //01. Таблица параметров базы данных
            new string[][]
            {
                //Ссылка [i][0][0]
                new string[] { "Параметры" },
                //Ссылка [i][1][0]
                new string[] { "Код" },
                //Ссылка [i][2][0]
                new string[] { "Код" },
                //Ссылка [i][3][0]
                new string[] { "Код", "Средняя_нагрузка", "Название_вуза", "Название_кафедры", 
                               "Название_ведомства", "Префикс_вуза", "Суффикс_вуза", "Оплата_асс", 
                               "Оплата_ст_преп", "Оплата_доц", "Оплата_проф" },
                //Ссылка [i][4][0]
                new string[] { "int PRIMARY KEY", "int", "string", "string",
                               "string", "string", "string", "double", 
                               "double", "double", "double" }
            },
            
            //02. Таблица учебных годов
            new string[][]
            {
                new string[] { "Список_Учебных_годов" },

                new string[] { "Код" },

                new string[] { "Учебный год" },

                new string[] { "Код", "Учебный год" },

                new string[] { "int PRIMARY KEY", "string" }
            },
            
            //03. Таблица семестров
            new string[][]
            {
                new string[] { "Список_семестров" },

                new string[] { "Код" },
                
                new string[] { "Номер_семестра" },

                new string[] { "Код", "Номер_семестра", "Предложный_падеж" },

                new string[] { "int PRIMARY KEY", "string" , "string" }
            },

            //04. Таблица номеров недель
            new string[][]
            {
                new string[] { "Недели_1-2" },

                new string[] { "Код" },

                new string[] { "Код" },

                new string[] { "Код", "Номер_недели" },

                new string[] { "int PRIMARY KEY", "string" }
            },

            //05. Таблица дней недели
            new string[][]
            {
                //Название таблицы базы данных
                new string[] { "Список_Дни недели" }, 
                //Ключевой атрибут
                new string[] { "№" },
                //Атрибут, по которому выполняется упорядочивание
                new string[] { "№" },
                //Перечень всех атрибутов таблицы
                new string[] { "№", "Название" },
                //Перечень типов, соответствующих атрибутам
                new string[] { "int PRIMARY KEY", "string" }
            },
            
            //06. Таблица времён пар
            new string[][]
            {
                new string[] { "Список_Время пар" },

                new string[] { "№" },

                new string[] { "№" },

                new string[] { "№", "Время" },

                new string[] { "int PRIMARY KEY", "string" }
            },
            
            //07. Таблица аудиторий
            new string[][]
            {
                new string[] { "Список_Аудиторий" },

                new string[] { "№" },
               
                new string[] { "Номер аудитории" },

                new string[] { "№", "Номер аудитории" },

                new string[] { "int PRIMARY KEY", "int" }

            },

            //08. Таблица учебных дисциплин
            new string[][]
            {
                new string[] { "Список _Дисциплин" },

                new string[] { "№" },

                new string[] { "Название_дисциплины" },

                new string[] { "№", "Название_дисциплины", "Сокращённое_название_дисциплины", "Аудитории" },

                new string[] { "int PRIMARY KEY", "string", "string", "string" }
            },

            //09. Таблица учебных курсов
            new string[][]
            {
                new string[] { "Список_курсов" },

                new string[] { "Код" },

                new string[] { "Код" },

                new string[] { "Код", "Курс" },

                new string[] { "int PRIMARY KEY", "int" }
            },

            //10. Таблица видов занятий
            new string[][]
            {
                new string[] { "Список_видов_занятий" },

                new string[] { "Код" },

                new string[] { "Код" },

                new string[] { "Код", "Вид_занятия", "Кратко", "КраткоИП", "В_распределении", "Для_форм" },

                new string[] { "int PRIMARY KEY", "string", "string", "string", "string", "string" }
            },

            //11. Таблица должностей
            new string[][]
            {
                new string[] { "Список_Должностей" },

                new string[] { "№" },

                new string[] { "№" },

                new string[] { "№", "Должность", "Коротко" },

                new string[] { "int PRIMARY KEY", "string", "string" }

            },

            //12. Таблица совместительства
            new string[][]
            {
                new string[] { "Список_Совместительства" },

                new string[] { "№" },

                new string[] { "№" },

                new string[] { "№", "Совместительство" },

                new string[] { "int PRIMARY KEY", "string" }
            },

            //13. Таблица званий
            new string[][]
            {
                new string[] { "Список_Званий" },

                new string[] { "№" },

                new string[] { "№" },

                new string[] { "№", "Звание" },

                new string[] { "int PRIMARY KEY", "string" }
            },
            
            //14. Таблица степеней
            new string[][]
            {
                new string[] { "Список_Степеней" },

                new string[] { "№" },

                new string[] { "№" },

                new string[] { "№", "Степень", "Коротко" },

                new string[] { "int PRIMARY KEY", "string", "string" }
            },

            //15. Таблица кафедр
            new string[][]
            {
                new string[] { "Список_кафедр" },

                new string[] { "Код" },

                new string[] { "Код" },

                new string[] { "Код", "Кафедра", "Коротко" },

                new string[] { "int PRIMARY KEY", "string", "string" }
            },

            //16. Таблица факультетов
            new string[][]
            {
                new string[] { "Список_факультетов" },

                new string[] { "Код" },
                
                new string[] { "Код" },

                new string[] { "Код", "Факультет", "Сокр_название", "Система" },

                new string[] { "int PRIMARY KEY", "string", "string", "string" }
            },

            //17. Таблица специализаций
            new string[][]
            {
                new string[] { "Список_специализаций" },

                new string[] { "Код" },

                new string[] { "Код" },

                new string[] { "Код", "Специальность", "Сокращённое_название_специальности",
                               "Аббревиатура", "Факультет", "Сокращённое_название_специализации",
                               "Система" },

                new string[] { "int PRIMARY KEY", "string", "string", 
                               "string", "string", "string", 
                               "string" }
            },

            //18. Таблица студенческих групп
            new string[][]
            {
                new string[] { "Группы студентов" },

                new string[] { "№" },

                new string[] { "Курс" },

                new string[] { "№", "Факультет", "Специальность", "Курс" },

                new string[] { "int PRIMARY KEY", "string", "string", "int" }
            },
            
            //19. Таблица преподавателей
            new string[][]
            {
                new string[] { "Преподаватели" },

                new string[] { "№" },

                new string[] { "ФИО" },

                new string[] { "№", "ФИО", "Кафедра", "Год_рождения", "Должность", "Учёное_звание",
                               "Учёная_степень", "Совместительство", "Общий_стаж_работы",
                               "Ставка", "Максимальная_нагрузка", "Должность1", "Разгрузка",
                               "Для_диспетчерской", "Совет", "Старая_ставка", "Аудитории",
                               "Ставка1", "Ставка2", "Смена_ставок" },

                new string[] { "int PRIMARY KEY", "string", "string", "datetime", "string", "string",
                               "string", "string", "datetime",
                               "double", "int", "string", "int",
                               "string", "bit", "double", "string",
                               "double", "double", "bit"}
            },

            //20. Таблица учебной нагрузки
            new string[][]
            {
                new string[] { "Учебная_нагрузка_распр" },

                new string[] { "Код" },

                new string[] { "Семестр", "КодДок", "Курс" },

                new string[] { "Код", "Семестр", "Специальность", "Курс", "Название_дисциплины",
                               "Преподаватель", "Преподаватель2", "Преподаватель3", "Лекция", "Экзамен",
                               "Зачёт", "Рефераты, домашние задания", "Консультация", "Лаб_раб", "Практ_зан",
                               "Индив_зан", "КРАПК", "Курс_пр", "Предд_пр", "Дипл_пр", "Учебн_пр", "Произв_пр",
                               "Аспирантура", "ГАК", "Часы", "Часы_Дано", "Посещ_зан", "Для_диспетчерской", 
                               "Магистратура", "В_ведом_дисп", "Дублёр", "Связь_лб", "Часы_ЗЕТ", "Часы_Дано_ЗЕТ",
                               "Распределяемая", "Вес_единицы", "Исключить", "КодДок", "Связь_почас" },

                new string[] { "int PRIMARY KEY", "string", "string", "string", "string",
                               "string", "string", "string", "int", "int",
                               "int", "int", "int", "int", "int",
                               "int", "int", "int", "int", "int", "int", "int",
                               "int", "int", "int", "int", "int", "string",
                               "int", "bit DEFAULT true", "string", "string", "double", "double", 
                               "bit DEFAULT false", "int", "bit DEFAULT false", "int", "string" }
            },

            //21. Таблица почасовой нагрузки
            new string[][]
            {
                new string[] { "Почасовая_нагрузка_распр" },

                new string[] { "Код" },

                new string[] { "Семестр", "КодДок", "Курс" },

                new string[] { "Код", "Семестр", "Специальность", "Курс", "Название_дисциплины",
                               "Преподаватель", "Преподаватель2", "Преподаватель3", "Лекция", "Экзамен",
                               "Зачёт", "Рефераты, домашние задания", "Консультация", "Лаб_раб", "Практ_зан",
                               "Индив_зан", "КРАПК", "Курс_пр", "Предд_пр", "Дипл_пр", "Учебн_пр", "Произв_пр",
                               "Аспирантура", "ГАК", "Часы", "Часы_Дано", "Посещ_зан", "Для_диспетчерской", 
                               "Магистратура", "В_ведом_дисп", "Дублёр", "Связь_лб", "Часы_ЗЕТ", "Часы_Дано_ЗЕТ",
                               "Распределяемая", "Вес_единицы", "Исключить", "КодДок", "Связь_почас" },

                new string[] { "int PRIMARY KEY", "string", "string", "string", "string",
                               "string", "string", "string", "int", "int",
                               "int", "int", "int", "int", "int",
                               "int", "int", "int", "int", "int", "int", "int",
                               "int", "int", "int", "int", "int", "string",
                               "int", "bit DEFAULT true", "string", "string", "double", "double", 
                               "bit DEFAULT false", "int", "bit DEFAULT false", "int", "string" }
            },

            //22. Таблица дополнительной работы
            new string[][]
            {
                new string[] { "ДопРабота" },

                new string[] { "Код" },

                new string[] { "Код" },

                new string[] { "Код", "Преподаватель", "Семестр", "УчебноМР", "НаучИсР", "ОргМР", "ОбъемУМР",
                               "ОбъемНИР", "ОбъемОМР", "СрокУМР", "СрокНИР", "СрокОМР", "КомУМР", "КомНИР", "КомОМР" },

                new string[] { "int PRIMARY KEY", "string", "string", "memo", "memo", "memo", "memo", 
                               "memo", "memo", "memo", "memo", "memo", "memo", "memo", "memo" }
            },

            //23. Таблица вопросов для заседаний кафедры
            new string[][]
            {
                new string[] { "Вопросы_заседаний" },

                new string[] { "Код" },
                
                new string[] { "Дата" },

                new string[] { "Код", "Дата", "Вопрос", "Докладчик1", 
                               "Докладчик2", "Докладчик3", "Докладчик4", 
                               "Докладчик5" },

                new string[] { "int PRIMARY KEY", "datetime", "memo", "string",
                               "string", "string", "string", "string" }
            },

            //24. Таблица расписания преподавателей
            new string[][]
            {
                new string[] { "Расписание преподавателей" },

                new string[] { "ID" },

                new string[] { "Преподаватель", "Неделя", "День недели", "Время занятия" },

                new string[] { "ID", "Преподаватель", "День недели", 
                               "Время занятия", "Неделя", "Занят",
                               "Дисциплина", "Тип_Занятия", "Специальность", "Курс",
                               "Аудитория", "Семестр", "Группа", "Поток" },

                new string[] { "int PRIMARY KEY", "string", "string", 
                               "string", "string", "bit DEFAULT false",
                               "string", "string", "string", "string",
                               "string", "string", "string", "string" }
            },

            //25. Таблица студентов
            new string[][]
            {
                new string[] { "Студенты" },

                new string[] { "Код" },

                new string[] { "ФИО" },

                new string[] { "Код", "ФИО", "Специальность", "Курс", "Кафедра", "Руководитель", 
                               "Тема_работы", "В_плане", "Почасовая" },

                new string[] { "int PRIMARY KEY", "string", "string", "string", "string", "string",
                               "string", "bit DEFAULT true", "bit DEFAULT false" }
            },

            //26. Таблица итогов
            new string[][]
            {
                new string[] { "Итоги" },

                new string[] { "Код" },

                new string[] { "Код" },

                new string[] { "Код", "Факультет", "Семестр", "Лекции", "Экзамен", "Зачёт", "Реферат",
                               "Консультация", "Лабораторные", "Практические", "Индивидуальные", "КРАПК",
                               "КурсоваяР", "ПреддипломнаяП", "Диплом", "УчебнаяП", "ПроизводственнаяП",
                               "ГЭК", "Бюджет", "Всего", "Итого", "БюджетЗЕТ", "ВсегоЗЕТ", "ИтогоЗЕТ"},

                new string[] { "int PRIMARY KEY", "string", "string", "double", "double", "double", "double",
                               "double", "double", "double", "double", "double",
                               "double", "double", "double", "double", "double", 
                               "double", "double", "double", "double", "double", "double", "double"}
            },

            //27. Таблица аспирантов
            new string[][]
            {
                new string[] { "Аспиранты" },

                new string[] { "Код" },

                new string[] { "ФИО" },

                new string[] { "Код", "ФИО", "Кафедра", "Руководитель", "Курс", "В_плане",
                               "Тема_работы", "Часы", "Бюджет" },

                new string[] { "int PRIMARY KEY", "string", "string", "string", "string", "bit DEFAULT true",
                               "string", "int", "bit DEFAULT true"}
            },

            //28. Таблица учебной нагрузки (конвертация)
            new string[][]
            {
                new string[] { "Учебная_нагрузка_конв" },

                new string[] { "Код" },

                new string[] { "Код" },

                new string[] { "Код", "Семестр", "Специальность", 
                               "Курс", "Название_дисциплины",
                               "Преподаватель", "Преподаватель2", 
                               "Преподаватель3", "Вид_занятия", 
                               "Часы", "Для_диспетчерской", 
                               "В_ведом_дисп", "Дублёр", "Связь_лб",
                               "Распределяемая", "Исключить" },

                new string[] { "int PRIMARY KEY", "string", "string", 
                               "string", "string",
                               "string", "string", 
                               "string", "string", 
                               "double", "string",
                               "bit DEFAULT true", "string", "string",
                               "bit DEFAULT false", "bit DEFAULT false" }
            },

            //29. Таблица больничных листов
            new string[][]
            {
                new string[] { "Больничные_листы" },

                new string[] { "Код" },

                new string[] { "Код" },

                new string[] { "Код", "Семестр", 
                               "Преподаватель", "Открытие", 
                               "Закрытие", "Примечание" },

                new string[] { "int PRIMARY KEY", "string", 
                               "string", "datetime",
                               "datetime", "string" }
            },
        };

        public static string getTabAttributes(string[] mas)
        {
            string str = "";

            for (int i = 0; i <= mas.Length - 1; i++)
            {
                if (i != mas.Length - 1)
                {
                    str += "[" + mas[i] + "], ";
                }
                else
                {
                    str += "[" + mas[i] + "]";
                }
            }
            return str;
        }
    }
}
