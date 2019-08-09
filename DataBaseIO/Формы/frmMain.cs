using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Reflection;
using System.IO;

using Visio = Microsoft.Office.Interop.Visio;
using Word = Microsoft.Office.Interop.Word;

namespace DataBaseIO
{
    public partial class frmMain : Form
    {
        /// <summary>
        /// Метод-конструктор главной формы
        /// </summary>
        public frmMain()
        {
            InitializeComponent();
        }

        //-------------------Перечень компонентов меню

        //Создаём новый объект "Полоса меню"
        MenuStrip mnuMain = new MenuStrip();

        //-------------------------Формирование шапки меню-------------

        //Создаём новый объект "Элемент полосы меню" с именем "Файл"
        private ToolStripMenuItem mnuFile = new ToolStripMenuItem("Файл");

        //Создаём новый объект "Элемент полосы меню" с именем "Открыть"
        private ToolStripMenuItem mnuOpen = new ToolStripMenuItem("Открыть");
        //Создаём новый объект "Элемент полосы меню" с именем "Сохранить"
        private ToolStripMenuItem mnuSave = new ToolStripMenuItem("Сохранить");
        //Создаём новый объект "Элемент полосы меню" с именем "Сброс"
        private ToolStripMenuItem mnuReset = new ToolStripMenuItem("Сброс");
        //Создаём новый объект "Элемент полосы меню" с именем "Выход"
        private ToolStripMenuItem mnuExit = new ToolStripMenuItem("Выход");



        //Создаём новый объект "Элемент полосы меню" с именем "Редактирование простое"
        private ToolStripMenuItem mnuEditSmpl = new ToolStripMenuItem("Ред. простое");

        //Создаём новый объект "Элемент полосы меню" с именем "Параметры"
        private ToolStripMenuItem mnuParams = new ToolStripMenuItem("Параметры");
        //Создаём новый объект "Элемент полосы меню" с именем "Кафедры"
        private ToolStripMenuItem mnuDepart = new ToolStripMenuItem("Кафедры");
        //Создаём новый объект "Элемент полосы меню" с именем "Степени"
        private ToolStripMenuItem mnuDegree = new ToolStripMenuItem("Степени");
        //Создаём новый объект "Элемент полосы меню" с именем "Звания"
        private ToolStripMenuItem mnuStatus = new ToolStripMenuItem("Звания");
        //Создаём новый объект "Элемент полосы меню" с именем "Должности"
        private ToolStripMenuItem mnuDuty = new ToolStripMenuItem("Должности");
        //Создаём новый объект "Элемент полосы меню" с именем "Совместительства"
        private ToolStripMenuItem mnuComb = new ToolStripMenuItem("Совместительства");
        //Создаём новый объект "Элемент полосы меню" с именем "Факультеты"
        private ToolStripMenuItem mnuFaculty = new ToolStripMenuItem("Факультеты");
        //Создаём новый объект "Элемент полосы меню" с именем "Учебные годы"
        private ToolStripMenuItem mnuWorkYear = new ToolStripMenuItem("Учебные годы");
        //Создаём новый объект "Элемент полосы меню" с именем "Семестры"
        private ToolStripMenuItem mnuSemestr = new ToolStripMenuItem("Семестры");
        //Создаём новый объект "Элемент полосы меню" с именем "Дисциплины"
        private ToolStripMenuItem mnuSubject = new ToolStripMenuItem("Дисциплины");
        //Создаём новый объект "Элемент полосы меню" с именем "Курсы"
        private ToolStripMenuItem mnuKursNum = new ToolStripMenuItem("Курсы");
        //Создаём новый объект "Элемент полосы меню" с именем "Больничные листы"
        private ToolStripMenuItem mnuSickLists = new ToolStripMenuItem("Больничные листы");


        //Создаём новый объект "Элемент полосы меню" с именем "Редактирование комплексное"
        private ToolStripMenuItem mnuEditComp = new ToolStripMenuItem("Ред. комплексное");
        //Создаём новый объект "Элемент полосы меню" с именем "Документооборот"
        private ToolStripMenuItem mnuDocs = new ToolStripMenuItem("Документооборот");
        //Создаём новый объект "Элемент полосы меню" с именем "Расчёты"
        private ToolStripMenuItem mnuCount = new ToolStripMenuItem("Расчёты");
        //Создаём новый объект "Элемент полосы меню" с именем "Справка"
        private ToolStripMenuItem mnuHelp = new ToolStripMenuItem("Справка");
        //Создаём новый объект "Элемент полосы меню" с именем "В разработке"
        private ToolStripMenuItem mnuConstruction = new ToolStripMenuItem("В разработке");

        //-------------------------Формирование шапки меню-------------        

        //-------------------Перечень компонентов меню

        /// <summary>
        /// 01. Таблица параметров базы данных
        /// </summary>
        public DataTable tabParams = new DataTable();

        /// <summary>
        /// 02. Таблица Учебных годов
        /// </summary>
        public DataTable tabWorkYear = new DataTable();

        /// <summary>
        /// 03. Таблица Семестров
        /// </summary>
        public DataTable tabSemestr = new DataTable();

        /// <summary>
        /// 04. Таблица номеров недель
        /// </summary>
        public DataTable tabNumberWeek = new DataTable();

        /// <summary>
        /// 05. Таблица Дней Недели
        /// </summary>       
        public DataTable tabWeekDays = new DataTable();

        /// <summary>
        /// 06. Таблица Времени Пар
        /// </summary>
        public DataTable tabPairTime = new DataTable();

        /// <summary>
        /// 07. Таблица Аудиторий
        /// </summary>
        public DataTable tabAuditory = new DataTable();

        /// <summary>
        /// 08. Таблица Учебных дисциплин
        /// </summary>
        public DataTable tabSubject = new DataTable();

        /// <summary>
        /// 09. Таблица Номеров Курсов
        /// </summary>
        public DataTable tabKursNum = new DataTable();

        /// <summary>
        /// 10. Таблица Видов занятий
        /// </summary>
        public DataTable tabSubjectTypes = new DataTable();

        /// <summary>
        /// 11. Таблица Должностей
        /// </summary>
        public DataTable tabDuty = new DataTable();

        /// <summary>
        /// 12. Таблица Совместительства
        /// </summary>       
        public DataTable tabCombination = new DataTable();

        /// <summary>
        /// 13. Таблица Званий
        /// </summary>       
        public DataTable tabStatus = new DataTable();

        /// <summary>
        /// 14. Таблица Степеней
        /// </summary>       
        public DataTable tabDegree = new DataTable();

        /// <summary>
        /// 15. Таблица Кафедр
        /// </summary>       
        public DataTable tabDeparment = new DataTable();
        
        /// <summary>
        /// 16. Таблица Факультетов
        /// </summary>
        public DataTable tabFaculty = new DataTable();

        /// <summary>
        /// 17. Таблица cпециальностей
        /// </summary>
        public DataTable tabSpecialisation = new DataTable();

        /// <summary>
        /// 18. Таблица Студенческих групп
        /// </summary>
        public DataTable tabStudentGroups = new DataTable();

        /// <summary>
        /// 19. Таблица Преподавателей
        /// </summary>
        public DataTable tabLecturer = new DataTable();

        /// <summary>
        /// 20. Таблица штатной нагрузки
        /// </summary>
        public DataTable tabDistribution = new DataTable();

        /// <summary>
        /// 21. Таблица почасовой нагрузки
        /// </summary>
        public DataTable tabHouredDistribution = new DataTable();

        /// <summary>
        /// 22. Таблица дополнительной работы
        /// </summary>
        public DataTable tabDopWork = new DataTable();
               
        /// <summary>
        /// 23. Таблица Вопросов заседаний кафедры
        /// </summary>
        public DataTable tabQuestions = new DataTable();

        /// <summary>
        /// 24. Таблица расписания
        /// </summary>
        public DataTable tabSchedule = new DataTable();

        /// <summary>
        /// 25. Таблица студентов
        /// </summary>
        public DataTable tabStudents = new DataTable();

        /// <summary>
        /// 26. Таблица итогов
        /// </summary>
        public DataTable tabSummary = new DataTable();

        /// <summary>
        /// 27. Таблица аспирантов
        /// </summary>
        public DataTable tabPGStudents = new DataTable();

        /// <summary>
        /// 28. Таблица сконвертированной учебной нагрузки
        /// </summary>
        public DataTable tabDistribConv = new DataTable();

        /// <summary>
        /// 29. Таблица больничных листов
        /// </summary>
        public DataTable tabSickList = new DataTable();

        private DataTable getDataTableByName(string Name)
        {
            DataTable Tab = null;

            switch (Name)
            {
                case "Список_Дни недели":
                    Tab = tabWeekDays;
                    break;

                case "Список_Время пар":
                    Tab = tabPairTime;
                    break;

                case "Список_Должностей":
                    Tab = tabDuty;
                    break;

                case "Список_Аудиторий":
                    Tab = tabAuditory;
                    break;

                case "Список_семестров":
                    Tab = tabSemestr;
                    break;

                case "Список_курсов":
                    Tab = tabKursNum;
                    break;

                case "Список_Учебных_годов":
                    Tab = tabWorkYear;
                    break;

                case "Список_факультетов":
                    Tab = tabFaculty;
                    break;

                case "Группы студентов":
                    Tab = tabStudentGroups;
                    break;

                case "Преподаватели":
                    Tab = tabLecturer;
                    break;

                case "Список _Дисциплин":
                    Tab = tabSubject;
                    break;

                case "Список_видов_занятий":
                    Tab = tabSubjectTypes;
                    break;

                case "Список_специализаций":
                    Tab = tabSpecialisation;
                    break;

                case "Учебная_нагрузка_распр":
                    Tab = tabDistribution;
                    break;

                case "ДопРабота":
                    Tab = tabDopWork;
                    break;

                case "Почасовая_нагрузка_распр":
                    Tab = tabHouredDistribution;
                    break;

                case "Список_Совместительства":
                    Tab = tabCombination;
                    break;

                case "Список_Званий":
                    Tab = tabStatus;
                    break;

                case "Список_Степеней":
                    Tab = tabDegree;
                    break;

                case "Список_кафедр":
                    Tab = tabDeparment;
                    break;

                case "Параметры":
                    Tab = tabParams;
                    break;

                case "Вопросы_заседаний":
                    Tab = tabQuestions;
                    break;

                case "Недели_1-2":
                    Tab = tabNumberWeek;
                    break;

                case "Расписание преподавателей":
                    Tab = tabSchedule;
                    break;

                case "Студенты":
                    Tab = tabStudents;
                    break;

                case "Итоги":
                    Tab = tabSummary;
                    break;

                case "Аспиранты":
                    Tab = tabPGStudents;
                    break;

                case "Учебная_нагрузка_конв":
                    Tab = tabDistribConv;
                    break;

                case "Больничные_листы":
                    Tab = tabSickList;
                    break;
            }

            return Tab;
        }

        private void setDataTableByName(string Name, DataTable Tab)
        {
            switch (Name)
            {
                case "Список_Дни недели":
                    tabWeekDays = Tab;
                    break;

                case "Список_Время пар":
                    tabPairTime = Tab;
                    break;

                case "Список_Должностей":
                    tabDuty = Tab;
                    break;

                case "Список_Аудиторий":
                    tabAuditory = Tab;
                    break;

                case "Список_семестров":
                    tabSemestr = Tab;
                    break;

                case "Список_курсов":
                    tabKursNum = Tab;
                    break;

                case "Список_Учебных_годов":
                    tabWorkYear = Tab;
                    break;

                case "Список_факультетов":
                    Tab = tabFaculty;
                    break;

                case "Группы студентов":
                    Tab = tabStudentGroups;
                    break;

                case "Преподаватели":
                    tabLecturer = Tab;
                    break;

                case "Список _Дисциплин":
                    tabSubject = Tab;
                    break;

                case "Список_видов_занятий":
                    tabSubjectTypes = Tab;
                    break;

                case "Список_специализаций":
                    tabSpecialisation = Tab;
                    break;

                case "Учебная_нагрузка_распр":
                    tabDistribution = Tab;
                    break;

                case "ДопРабота":
                    tabDopWork = Tab;
                    break;

                case "Почасовая_нагрузка_распр":
                    tabHouredDistribution = Tab;
                    break;

                case "Список_Совместительства":
                    tabCombination = Tab;
                    break;

                case "Список_Званий":
                    tabStatus = Tab;
                    break;

                case "Список_Степеней":
                    tabDegree = Tab;
                    break;

                case "Список_кафедр":
                    tabDeparment = Tab;
                    break;

                case "Параметры":
                    tabParams = Tab;
                    break;

                case "Вопросы_заседаний":
                    tabQuestions = Tab;
                    break;

                case "Недели_1-2":
                    tabNumberWeek = Tab;
                    break;

                case "Расписание преподавателей":
                    tabSchedule = Tab;
                    break;

                case "Студенты":
                    tabStudents = Tab;
                    break;

                case "Итоги":
                    tabSummary = Tab;
                    break;

                case "Аспиранты":
                    tabPGStudents = Tab;
                    break;

                case "Учебная_нагрузка_конв":
                    tabDistribConv = Tab;
                    break;

                case "Больничные_листы":
                    tabSickList = Tab;
                    break;
            }
        }

        //Загрузка базы данных
        private void onLoad()
        {
            mdlData.Reg = 0;
            //Фильтрация
            //- либо по всем файлам, 
            //- либо только по файлам MS Access до 2007 выпуска
            //- либо только по файлам MS Access начная с 2007 выпуска
            dlgOpen.Filter = "Базы данных MS Access 2007+|*.accdb|Базы данных MS Access 2003|*.mdb|Все файлы|*.*";
            //Ввод заголовка диалогового окна загрузки базы данных
            dlgOpen.Title = "Открыть Базу Данных...";
            //Указание на начало работы с директорией, откуда был запущен исполняемый файл проекта
            dlgOpen.InitialDirectory = Application.StartupPath;
            //Вывод/демонстрация диалогового окна для открытия файла с базой данных
            dlgOpen.ShowDialog();

            //Если всё удачно открылось
            if (mdlData.flgReady)
            {
                //Перенастраиваем доступность элементов
                //главной формы
                mdlData.statString = "База загружена";
                mdlData.flgLoad = true;
            }

            lblStatus.Text = mdlData.statString;
        }

        /// <summary>
        /// Установка состояния кнопки Доступна/Недоступна по наполнению коллекции
        /// </summary>
        /// <param name="btn">Кнопка</param>
        /// <param name="cnt">Количество элементов коллекции</param>
        private void SetButtonState(ref Button btn, int cnt)
        {
            if (cnt > 0)
            { btn.Enabled = true; }
            else
            { btn.Enabled = false; }
        }

//-----------------------------------------------------------------------------
//-----------------------------------------------------------------------------
//-------------------Загрузка базы данных--------------------------------------
//-----------------------------------------------------------------------------
//-----------------------------------------------------------------------------

        /// <summary>
        /// Результат выбора файла в диалоговом окне открытия файла
        /// </summary>
        /// <param name="sender">Объект, содержащий сведения об источнике вызова обработчика этого события</param>
        /// <param name="e">Перечень параметров отмены действия, связанных с событием нажатия на кнопку "Отмена"</param>
        private void dlgOpen_FileOk(object sender, CancelEventArgs e)
        {
            switch (mdlData.Reg)
            {
                case 0:
                    {
                        //Если не была нажата кнопка "Отмена"
                        if (!e.Cancel)
                        {
                            //Очистка ранее созданных коллекций
                            StartClearCollections();
                            //Очистка ранее созданных таблиц
                            StartClearTables();

                            //Флаг успеха загрузки по умолчанию должен быть выставлен
                            mdlData.flgReady = true;
                            //Запись пути к Базе Данных в глобальную переменную из диалогового окна
                            mdlData.DataBasePath = dlgOpen.FileName;
                            //Переход к методу загрузки информации из базы данных
                            OpenDB(mdlData.DataBasePath, ref mdlData.flgReady);
                            //Обновляем состояние главного меню в зависимости от того, 
                            //загрузилась база данных или нет
                            InitMain(mdlData.flgLoad);
                        }

                        //Добавление к заголовку формы имени открытой базы данных
                        this.Text += ": " + mdlData.DataBasePath.Substring(mdlData.DataBasePath.LastIndexOf('\\') + 1);
                        break;
                    }
                default:
                    {
                        break;
                    }
            }
        }

        /// <summary>
        /// Метод открытия базы данных
        /// </summary>
        /// <param name="Path">Путь к базе данных</param>
        /// <param name="ready">Признак готовности/успеха считывания информации из базы данных</param>
        public void OpenDB(string Path, ref bool ready)
        {
            //Определение локальной переменной связи с базой данных
            OleDbConnection connection = new OleDbConnection();
            
            //Определение слова-сопряжения с базой данных в зависимости от версии
            //MS Access, определяемой по расширению файла
            if (Path.EndsWith(".mdb") || Path.EndsWith(".MDB"))
            {
                // MS Access до 2007 года
                connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + @"data source = " + Path;
            }
            else if (Path.EndsWith(".accdb") || Path.EndsWith(".ACCDB"))
            {
                // MS Access с 2007
                connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" + @"data source = " + Path;
                //Persist Security Info=False;
            }
            
            //Передача сформированного соединения с базой данных в глобальную переменную
            mdlData.glConn = connection;
            //Метод проверки наличия в базе данных обязательных отношений (таблиц)
            CheckTablesDB(ref ready);

            //Если предыдущий метод не дал ошибки, то реализуем проверку совместимости далее
            if (ready)
            {
                //Метод проверки наличия в базе данных обязательных атрибутов
                CheckKeyAttribsDB(ref ready);
            }

            if (ready)
            {
                //Пробуем открыть соединение с базой данных
                try
                {
                    //Открываем соединение с базой данных
                    mdlData.glConn.Open();
                }
                //Если не получилось, то сбросить флаг готовности
                //и уведомить пользователя
                catch
                {
                    MessageBox.Show("Загрузка базы данных прекращена.", "Не удалось открыть соединение с базой данных");
                    ready = false;
                }

                if (ready)
                {
                    //Заполняем таблицы (элементы DataTable)
                    CreateInnerTables(ref ready);
                }

                if (ready)
                {
                    //Проверяем, есть ли требуемые для расчёта атрибуты, а
                    //если нет, то создаём их
                    CheckFieldsDB(Path, ref ready);
                }

                //...........................................................................

                //Пытаемся закрыть соединение с базой данных
                try
                {
                    //Закрываем соединение с базой данных
                    mdlData.glConn.Close();
                }
                //Если не получается, то уведомляем об этом пользователя
                catch
                {
                    MessageBox.Show("Соединение не закрыто. Вероятно, оно и открыто не было.",
                                    "Не удалось закрыть соединение с базой данных");
                }

                if (ready)
                {
                    //Создаём коллекции с пустыми объектами
                    CreateCollections(ref ready);
                }

                //MessageBox.Show("Создана коллекция кафедр размерностью: " + mdlData.colDepart.Count.ToString(), "01. Кафедра");

                //Наполняем содержанием элементы коллекций
                LinkingCollections(ref ready);

                if (ready)
                {
                    mdlData.flgLoad = true;
                }
            }
        }

        //-----------------------------------------------------------------------------
        //-----------------------------------------------------------------------------
        //-------------------Загрузка базы данных--------------------------------------
        //-----------------------------------------------------------------------------
        //-----------------------------------------------------------------------------

        /// <summary>
        /// Проверка наличия ключевых атрибутов в базе данных
        /// </summary>
        /// <param name="ready">Возвращаемый признак готовности системы к работе с данными</param>
        private void CheckKeyAttribsDB(ref bool ready)
        {
            string Message = "";
            string tabName = "";
            string Attrib = "";
            string AttribType = "";
            bool flgContinue = true;

            //Открываем соединение с базой данных
            mdlData.glConn.Open();

            //До тех пор, пока флаг готовности не сброшен (всё проходит в штатном режиме)
            //и имеется необходимость в повторе операции проверки таблиц, повторять попытки
            while (flgContinue & ready)
            {
                //Попытка обращения к каждой обязательной таблице базы данных
                //с запросом на полную выборку
                try
                {
                    //Цикл по всем таблицам базы данных согласно модели
                    for (int i = 0; i <= mdlBaseStructure.masTabNames.Length - 1; i++)
                    {
                        //Получение имени таблицы
                        tabName = mdlBaseStructure.masTabNames[i][0][0];
                        //Цикл по атрибутам, по которым ведётся упорядочивание данных
                        for (int j = 0; j <= mdlBaseStructure.masTabNames[i][2].Length - 1; j++)
                        {
                            //Фиксируем наименование атрибута, по которому требуется упорядочивание
                            Attrib = mdlBaseStructure.masTabNames[i][2][j];
                            //Ищем в цикле такой же атрибут в перечне всех атрибутов таблицы
                            for (int k = 0; k <= mdlBaseStructure.masTabNames[i][3].Length - 1; k++)
                            {
                                //Если в перечне атрибутов нашёлся атрибут с нужным именем
                                if (mdlBaseStructure.masTabNames[i][3][k].Equals(Attrib))
                                {
                                    //То из перечня типов атрибутов забираем значение его типа
                                    //по тому же индексу и прерываем цикл
                                    AttribType = mdlBaseStructure.masTabNames[i][4][k];
                                    break;
                                }
                            }

                            //Попытка обращения с запросом на атрибут в таблице
                            TryToSelectAttrib(Attrib, tabName);
                        }
                    }

                    //Если удалось дойти до этого момента, то все запросы успено выполнились. 

                    //Сброс флага продолжения внешнего цикла
                    //При наличии всех таблиц в повторной проверке нет необходимости
                    flgContinue = false;
                }
                //Обработка исключения в случае возникновения ошибки в одном из запросов.
                //Предпринимается попытка создания отсутствующей таблицы
                catch (OleDbException e)
                {
                    Message = e.Message;
                    TryToCreateAttrib(tabName, Attrib, AttribType);
                }
            }

            //Закрыть соединение
            mdlData.glConn.Close();
        }

        /// <summary>
        /// Проверка наличия всех необходимых таблиц в базе данных
        /// </summary>
        /// <param name="ready">Возвращаемый признак готовности системы к работе с данными</param>
        private void CheckTablesDB(ref bool ready)
        {
            string tabName = "";
            string tabKey = "";
            bool flgContinue = true;

            //Открываем соединение с базой данных
            mdlData.glConn.Open();

            //Если хотя бы одна таблица из числа необходимых для 
            //нормальной работы системы в базе данных отсутсвует (не открывается),
            //то такую таблицу необходимо создать в базе данных
            
            //До тех пор, пока флаг готовности не сброшен (всё проходит в штатном режиме)
            //и имеется необходимость в повторе операции проверки таблиц, повторять попытки
            while (flgContinue & ready)
            {
                //Попытка обращения к каждой обязательной таблице базы данных
                //с запросом на полную выборку
                try
                {
                    //Цикл по всем таблицам базы данных согласно модели
                    for (int i = 0; i <= mdlBaseStructure.masTabNames.Length - 1; i++)
                    {
                        //Получение имени таблицы
                        tabName = mdlBaseStructure.masTabNames[i][0][0];
                        //Получение ключевого поля
                        tabKey = mdlBaseStructure.masTabNames[i][1][0];
                        //Попытка обращения с запросом на полную выборку
                        TryToSelectAll(tabName);
                    }

                    //Если удалось дойти до этого момента, то все запросы успено выполнились. 
                    
                    //Сброс флага продолжения внешнего цикла
                    //При наличии всех таблиц в повторной проверке нет необходимости
                    flgContinue = false;
                }
                //Обработка исключения в случае возникновения ошибки в одном из запросов.
                //Предпринимается попытка создания отсутствующей таблицы
                catch (OleDbException e)
                {
                    //Вывод ошибки об отсутствии таблицы, к которой не удалось выполнить запрос
                    MessageBox.Show("Не найдена таблица: " + tabName + "\n" + e.ToString(), "Ошибка");

                    if (MessageBox.Show("Создать таблицу " + tabName + "?",
                        "Попытка запустить систему с открытой БД",
                        MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        //Предпринимаем попытку создать таблицу базы данных
                        //Если дан положительный ответ, то
                        TryToCreateTable(tabName, tabKey);
                    }
                    //Если дан отрицательный ответ, то прекратить открытие базы данных
                    //- вылететь с ошибкой
                    else
                    {
                        ready = false;
                    }
                }
            }
            
            //Закрыть соединение
            mdlData.glConn.Close();
        }

        private void TryToSelectAll(string Name)
        {
            //Создаём переменную SQL-команды
            OleDbCommand OpenCommand = new OleDbCommand();
            //Приписываем переменной SQL-команды соединение
            OpenCommand.Connection = mdlData.glConn;

            OpenCommand.CommandText = "SELECT * FROM [" + Name + "]";
            OpenCommand.ExecuteNonQuery();
        }

        private void TryToSelectAttrib(string AttrName, string TabName)
        {
            //Создаём переменную SQL-команды
            OleDbCommand OpenCommand = new OleDbCommand();
            //Приписываем переменной SQL-команды соединение
            OpenCommand.Connection = mdlData.glConn;

            OpenCommand.CommandText = "SELECT [" + AttrName + "] FROM [" + TabName + "]";
            OpenCommand.ExecuteNonQuery();
        }

        private void TryToCreateTable(string Name, string Key)
        {
            //Создаём переменную SQL-команды
            OleDbCommand CreateCommand = new OleDbCommand();
            //Приписываем переменной SQL-команды соединение
            CreateCommand.Connection = mdlData.glConn;

            //создаём таблицу, имя которой отсутствует
            //пока без единого атрибута
            CreateCommand.CommandText = "CREATE TABLE [" + Name + "]";
            CreateCommand.ExecuteNonQuery();
            //А атрибут, являющийся ключевым полем, добавляем здесь
            CreateCommand.CommandText = "ALTER TABLE [" + Name + "] ADD ["
                                       + Key + "] int PRIMARY KEY";
            CreateCommand.ExecuteNonQuery();
        }

        private void TryToCreateAttrib(string Name, string Attrib, string AttribType)
        {
            //Создаём переменную SQL-команды
            OleDbCommand CreateCommand = new OleDbCommand();
            //Приписываем переменной SQL-команды соединение
            CreateCommand.Connection = mdlData.glConn;

            //Пытаемся дополнить таблицу атрибутом
            CreateCommand.CommandText = "ALTER TABLE [" + Name + "] ADD ["
                                       + Attrib + "] " + AttribType;
            CreateCommand.ExecuteNonQuery();
        }

        /// <summary>
        /// Создание таблиц базы данных в оперативной памяти
        /// </summary>
        private void CreateInnerTables(ref bool ready)
        {
            DataTable Tab;

            for (int i = 0; i <= mdlBaseStructure.masTabNames.Length - 1; i++)
            {
                Tab = getDataTableByName(mdlBaseStructure.masTabNames[i][0][0]);
                InnerTableFill(mdlBaseStructure.masTabNames[i][0][0], ref Tab, mdlBaseStructure.masTabNames[i][2]);
                setDataTableByName(mdlBaseStructure.masTabNames[i][0][0], Tab);
            }
        }

        /// <summary>
        /// Процедура заполнения виртуальной таблицы данными из базы данных
        /// </summary>
        /// <param name="TabName">Имя заполняемой таблицы в базе данных</param>
        /// <param name="Tab">Имя заполняемой таблицы в программе</param>
        /// <param name="Key">Строковый массив ключевых полей, в частности, 
        /// ключевое поле одно. Необходимо для упорядочивания по приоритету</param>
        private void InnerTableFill(string TabName, ref DataTable Tab, string[] Key)
        {
            string str = "";

            //Формируем SQL-запрос для поиска заданной таблицы
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            //Создаём переменную SQL-запроса
            OleDbCommand command = new OleDbCommand();
            
            //Указываем соединение для переменной SQL-запроса
            command.Connection = mdlData.glConn;
            //Указываем строку для переменную SQL-запроса

            for (int i = 0; i <= Key.Length - 1; i++)
            {
                if (i != Key.Length - 1)
                {
                    str += "[" + Key[i] + "], ";
                }
                else
                {
                    str += "[" + Key[i] + "]";
                }
            }

            //... ... ... ... ... ... ... ... ... ...

            command.CommandText = "SELECT * FROM [" + TabName + "] ORDER BY " + str;
            
            //Передаём SQL-запрос в адаптер
            adapter.SelectCommand = command;
            //Заполняем таблицу
            adapter.Fill(Tab);
        }

        private void InnerCheck(DataTable Tab, string TabName, string[] Attribs, string[] Endings, ref bool ready)
        {
            //Создаём переменную общения SQL-запросами с БД
            OleDbCommand sql = new OleDbCommand();
            //Передать соединение команде
            sql.Connection = mdlData.glConn;

            //Проверка столбцов имеет место тогда и только тогда,
            //когда таблица хоть как-то была сформирована
            if (ready)
            {
                if (Tab.Columns.Count > 0)
                {
                    for (int i = 0; i <= Attribs.Length - 1; i++)
                    {
                        if (!Tab.Columns.Contains(Attribs.GetValue(i).ToString()))
                        {
                            sql.CommandText = "ALTER TABLE [" + TabName + "] ADD [" + Attribs.GetValue(i).ToString() +
                                              "] " + Endings.GetValue(i).ToString();
                            sql.ExecuteNonQuery();
                            Tab.Columns.Add(Attribs.GetValue(i).ToString());
                        }
                    }
                }
                //Если таблица не формировалась, то необходимо предупредить пользователя
                //и запретить дальнейшее продолжение работы программы
                else
                {
                    ready = false;
                }
            }
        }
        
        /// <summary>
        /// Событие, связанное с загрузкой главной формы
        /// </summary>
        /// <param name="sender">Объект, содержащий сведения об источнике вызова обработчика этого события</param>
        /// <param name="e">Перечень параметров среды, связанных с событием загрузки главной формы</param>
        private void frmMain_Load(object sender, EventArgs e)
        {
            FormClosing += new FormClosingEventHandler(frmMain_FormClosing);

            CreateMenu();
            InitMain(mdlData.flgLoad);
        }


        void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show(this, "Действительно хотите прекратить работу с АРМом?",
                "Выход?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                e.Cancel = true;
            }
        }

        private void CreateMenu()
        {
            //-----------------Формирование пунктов "Файл"-----------------
            //По умолчанию пункт меню видимый
            mnuFile.Visible = true;

            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Выход"
            mnuExit.Click += new EventHandler(mnuExit_Click);
            //Добавляем горячую клавишу "Ctrl + Q" для пункта "Выход"
            mnuExit.ShortcutKeys = Keys.Control | Keys.Q;
            //По умолчанию пункт меню видимый
            mnuExit.Visible = true;

            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Сброс"
            mnuReset.Click += new EventHandler(mnuReset_Click);
            //Добавляем горячую клавишу "Ctrl + R" для пункта "Сброс"
            mnuReset.ShortcutKeys = Keys.Control | Keys.R;
            //По умолчанию пункт меню невидимый
            mnuReset.Visible = false;

            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Открыть"
            mnuOpen.Click += new EventHandler(mnuOpen_Click);
            //Добавляем горячую клавишу "Ctrl + L" для пункта "Открыть"
            mnuOpen.ShortcutKeys = Keys.Control | Keys.L;
            //По умолчанию пункт меню видимый
            mnuExit.Visible = true;

            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Сохранить"
            mnuSave.Click += new EventHandler(mnuSave_Click);
            //Добавляем горячую клавишу "Ctrl + S" для пункта "Сохранить"
            mnuSave.ShortcutKeys = Keys.Control | Keys.S;
            //По умолчанию пункт меню видимый
            mnuSave.Visible = true;
            //По умолчанию пункт меню недоступный
            mnuSave.Enabled = false;

            //-----------------Формирование пунктов "Файл"-----------------

            //------Формирование пунктов "Редактирование простое"----------
            //По умолчанию пункт меню невидимый
            mnuEditSmpl.Visible = false;

            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Параметры"
            mnuParams.Click += new EventHandler(mnuParams_Click);

            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Кафедры"
            mnuDepart.Click += new EventHandler(mnuDepart_Click);

            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Степени"
            mnuDegree.Click += new EventHandler(mnuDegree_Click);

            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Звания"
            mnuStatus.Click += new EventHandler(mnuStatus_Click);

            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Должности"
            mnuDuty.Click += new EventHandler(mnuDuty_Click);

            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Совместительства"
            mnuComb.Click += new EventHandler(mnuComb_Click);

            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Факультеты"
            mnuFaculty.Click += new EventHandler(mnuFaculty_Click);

            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Учебные годы"
            mnuWorkYear.Click += new EventHandler(mnuWorkYear_Click);

            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Семестры"
            mnuSemestr.Click += new EventHandler(mnuSemestr_Click);

            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Дисциплины"
            mnuSubject.Click += new EventHandler(mnuSubject_Click);

            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Курсы"
            mnuKursNum.Click += new EventHandler(mnuKursNum_Click);

            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Больничные листы"
            mnuSickLists.Click += MnuSickLists_Click;

            //Создаём новый объект "Элемент полосы меню" с именем "Дни недели"
            ToolStripMenuItem mnuWeekDay = new ToolStripMenuItem("Дни недели");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Дни недели"
            mnuWeekDay.Click += new EventHandler(mnuWeekDay_Click);

            //Создаём новый объект "Элемент полосы меню" с именем "Студенты"
            ToolStripMenuItem mnuStudents = new ToolStripMenuItem("Студенты");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Студенты"
            mnuStudents.Click += new EventHandler(mnuStudents_Click);

            //Создаём новый объект "Элемент полосы меню" с именем "Аспиранты"
            ToolStripMenuItem mnuPGStudents = new ToolStripMenuItem("Аспиранты");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Аспиранты"
            mnuPGStudents.Click += new EventHandler(mnuPGStudents_Click);
            //Создаём новый объект "Элемент полосы меню" с именем "Виды занятий"
            ToolStripMenuItem mnuSubjTypes = new ToolStripMenuItem("Виды занятий");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Виды занятий"
            mnuSubjTypes.Click += new EventHandler(mnuSubjTypes_Click);

            //------Формирование пунктов "Редактирование простое"----------

            //------Формирование пунктов "Редактирование комплексное"------
            //По умолчанию пункт меню невидимый
            mnuEditComp.Visible = false;

            //Создаём новый объект "Элемент полосы меню" с именем "Преподаватели"
            ToolStripMenuItem mnuLecturers = new ToolStripMenuItem("Преподаватели");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Преподаватели"
            mnuLecturers.Click += new EventHandler(mnuLecturers_Click);

            //Создаём новый объект "Элемент полосы меню" с именем "Специальность"
            ToolStripMenuItem mnuSpec = new ToolStripMenuItem("Специальность");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Специальность"
            mnuSpec.Click += new EventHandler(mnuSpec_Click);

            //Создаём новый объект "Элемент полосы меню" с именем "Дополнительная работа"
            ToolStripMenuItem mnuAddWork = new ToolStripMenuItem("Дополнительная работа");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Дополнительная работа"
            mnuAddWork.Click += new EventHandler(mnuAddWork_Click);

            //Создаём новый объект "Элемент полосы меню" с именем "Распределение нагрузки"
            ToolStripMenuItem mnuDistrib = new ToolStripMenuItem("Распределение нагрузки");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Распределение нагрузки"
            mnuDistrib.Click += new EventHandler(mnuDistrib_Click);
            //Добавляем горячую клавишу "Ctrl + R" для пункта "Распределение нагрузки"
            mnuDistrib.ShortcutKeys = Keys.Control | Keys.R;

            //Создаём новый объект "Элемент полосы меню" с именем "Расписание преподавателей"
            ToolStripMenuItem mnuSchedule = new ToolStripMenuItem("Расписание преподавателей");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Расписание преподавателей"
            mnuSchedule.Click += new EventHandler(mnuSchedule_Click);

            //------Формирование пунктов "Редактирование комплексное"------

            //--------------Формирование пунктов "Документооборот"---------
            //По умолчанию пункт меню невидимый
            mnuDocs.Visible = false;

            //Создаём новый объект "Элемент полосы меню" с именем "Индивидуальные планы"
            ToolStripMenuItem mnuPlans = new ToolStripMenuItem("Индивидуальные планы");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Индивидуальные планы"
            mnuPlans.Click += new EventHandler(mnuPlans_Click);
            //Добавляем горячую клавишу "Ctrl + I" для пункта "Индивидуальные планы"
            mnuDistrib.ShortcutKeys = Keys.Control | Keys.I;

            //Создаём новый объект "Элемент полосы меню" с именем "Заявка на расписание"
            ToolStripMenuItem mnuDispatch = new ToolStripMenuItem("Заявка на расписание");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Заявка на расписание"
            mnuDispatch.Click += new EventHandler(mnuDispatch_Click);

            //Создаём новый объект "Элемент полосы меню" с именем "Ведомости в учебное управление"
            ToolStripMenuItem mnuUMU = new ToolStripMenuItem("Ведомости в учебное управление");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Ведомости в учебное управление"
            mnuUMU.Click += new EventHandler(mnuUMU_Click);

            //Создаём новый объект "Элемент полосы меню" с именем "Замены по болезни"
            ToolStripMenuItem mnuSwap = new ToolStripMenuItem("Замены по болезни");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Замены по болезни"
            mnuSwap.Click += new EventHandler(mnuSwap_Click);

            //Создаём новый объект "Элемент полосы меню" с именем "Для заведующего кафедрой"
            ToolStripMenuItem mnuForChief = new ToolStripMenuItem("Для заведующего кафедрой");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Для заведующего кафедрой"
            mnuForChief.Click += new EventHandler(mnuForChief_Click);

            //Создаём новый объект "Элемент полосы меню" с именем "Протоколы кафедры"
            ToolStripMenuItem mnuProtocols = new ToolStripMenuItem("Протоколы кафедры");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Протоколы кафедры"
            mnuProtocols.Click += new EventHandler(mnuProtocols_Click);

            //--------------Формирование пунктов "Документооборот"---------

            //-----------------Формирование пунктов "Расчёты"--------------
            //По умолчанию пункт меню невидимый
            mnuCount.Visible = false;

            //Создаём новый объект "Элемент полосы меню" с именем "Расчёт ВКР"
            ToolStripMenuItem mnuVKR = new ToolStripMenuItem("Расчёт ВКР");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Расчёт ВКР"
            mnuVKR.Click += new EventHandler(mnuVKR_Click);

            //Создаём новый объект "Элемент полосы меню" с именем "Расчёт нагрузки по направлениям"
            ToolStripMenuItem mnuCountNapr = new ToolStripMenuItem("Расчёт нагрузки по направлениям");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Расчёт нагрузки по направлениям"
            mnuCountNapr.Click += new EventHandler(mnuCountNapr_Click);

            //Создаём новый объект "Элемент полосы меню" с именем "Расчёт нагрузки преп/напр"
            ToolStripMenuItem mnuCountLectNapr = new ToolStripMenuItem("Расчёт нагрузки по преп/напр");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Расчёт нагрузки преп/напр"
            mnuCountLectNapr.Click += new EventHandler(mnuCountLectNapr_Click);

            //Создаём новый объект "Элемент полосы меню" с именем "Перечень внешних"
            ToolStripMenuItem mnuFindOuterLect = new ToolStripMenuItem("Перечень внешних");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Перечень внешних"
            mnuFindOuterLect.Click += new EventHandler(mnuFindOuterLect_Click);

            //Создаём новый объект "Элемент полосы меню" с именем "Доля нагрузки внешних"
            ToolStripMenuItem mnuCountOuter = new ToolStripMenuItem("Доля нагрузки внешних");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Доля нагрузки внешних"
            mnuCountOuter.Click += new EventHandler(mnuCountOuter_Click);

            //Создаём новый объект "Элемент полосы меню" с именем "Распределяемая почасовая"
            ToolStripMenuItem mnuHouredDistrib = new ToolStripMenuItem("Распределяемая почасовая");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Распределяемая почасовая"
            mnuHouredDistrib.Click += MnuHouredDistrib_Click;

            //-----------------Формирование пунктов "Расчёты"--------------

            //-----------------Формирование пунктов "В разработке"---------
            //По умолчанию пункт меню видимый
            mnuCount.Visible = true;

            //Создаём новый объект "Элемент полосы меню" с именем "Visio Расписание"
            ToolStripMenuItem mnuVisTimeTable = new ToolStripMenuItem("Visio Расписание");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Visio Расписание"
            mnuVisTimeTable.Click += new EventHandler(mnuVisTimeTable_Click);
            //Добавляем горячую клавишу "Alt + V" для пункта "Visio Расписание"
            mnuVisTimeTable.ShortcutKeys = Keys.Alt | Keys.V;


            //Создаём новый объект "Элемент полосы меню" с именем "Visio Индивидуальное расписание"
            ToolStripMenuItem mnuVisIndTimeTable = new ToolStripMenuItem("Visio Индивидуальное Расписание");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Visio Индивидуальное расписание"
            mnuVisIndTimeTable.Click += new EventHandler(mnuVisIndTimeTable_Click);
            //Добавляем горячую клавишу "Alt + V" для пункта "Visio Индивидуальное расписание"
            mnuVisIndTimeTable.ShortcutKeys = Keys.Alt | Keys.I | Keys.V;

            //Создаём новый объект "Элемент полосы меню" с именем "Visio Тест"
            ToolStripMenuItem mnuVisTest = new ToolStripMenuItem("Visio Тест");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Visio Тест"
            mnuVisTest.Click += MnuVisTest_Click;

            //Создаём новый объект "Элемент полосы меню" с именем "Импорт Word"
            ToolStripMenuItem mnuWordImport = new ToolStripMenuItem("Импорт Word");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Импорт Word"
            mnuWordImport.Click += MnuWordImport_Click;
            //Добавляем горячую клавишу "Alt + W" для пункта "Импорт Word"
            mnuWordImport.ShortcutKeys = Keys.Alt | Keys.W;

            //-----------------Формирование пунктов "В разработке"---------

            //------------Добавление пунктов "Файл" к меню-----------------

            //Добавляем в меню "Файл" элемент "Открыть..."
            mnuFile.DropDownItems.Add(mnuOpen);
            //Добавляем в меню "Файл" элемент "Сохранить..."
            mnuFile.DropDownItems.Add(mnuSave);
            //Добавляем в меню "Файл" элемент "Сбросить"
            mnuFile.DropDownItems.Add(mnuReset);
            //Добавляем в меню "Файл" разделитель
            mnuFile.DropDownItems.Add(new ToolStripSeparator());
            //Добавляем в меню "Файл" элемент "Выход"
            mnuFile.DropDownItems.Add(mnuExit);

            //------------Добавление пунктов "Файл" к меню-----------------

            //-----Добавление пунктов "Редактирование простое" к меню------

            //Добавляем в меню "Ред. прост." элемент "Параметры"
            mnuEditSmpl.DropDownItems.Add(mnuParams);
            //Добавляем в меню "Ред. прост." разделитель
            mnuEditSmpl.DropDownItems.Add(new ToolStripSeparator());
            //Добавляем в меню "Ред. прост." элемент "Кафедры"
            mnuEditSmpl.DropDownItems.Add(mnuDepart);
            //Добавляем в меню "Ред. прост." элемент "Степени"
            mnuEditSmpl.DropDownItems.Add(mnuDegree);
            //Добавляем в меню "Ред. прост." элемент "Звания"
            mnuEditSmpl.DropDownItems.Add(mnuStatus);
            //Добавляем в меню "Ред. прост." элемент "Совместительства"
            mnuEditSmpl.DropDownItems.Add(mnuComb);
            //Добавляем в меню "Ред. прост." элемент "Факультеты"
            mnuEditSmpl.DropDownItems.Add(mnuFaculty);
            //Добавляем в меню "Ред. прост." элемент "Учебные годы"
            mnuEditSmpl.DropDownItems.Add(mnuWorkYear);
            //Добавляем в меню "Ред. прост." элемент "Семестры"
            mnuEditSmpl.DropDownItems.Add(mnuSemestr);
            //Добавляем в меню "Ред. прост." элемент "Дисциплины"
            mnuEditSmpl.DropDownItems.Add(mnuSubject);
            //Добавляем в меню "Ред. прост." элемент "Номера курсов"
            mnuEditSmpl.DropDownItems.Add(mnuKursNum);
            //Добавляем в меню "Ред. прост." элемент "Дни недели"
            mnuEditSmpl.DropDownItems.Add(mnuWeekDay);
            //Добавляем в меню "Ред. прост." элемент "Студенты"
            mnuEditSmpl.DropDownItems.Add(mnuStudents);
            //Добавляем в меню "Ред. прост." элемент "Виды занятий"
            mnuEditSmpl.DropDownItems.Add(mnuSubjTypes);
            //Добавляем в меню "Ред. прост." элемент "Аспиранты"
            mnuEditSmpl.DropDownItems.Add(mnuPGStudents);
            //Добавляем в меню "Ред. прост." элемент "Больничные листы"
            mnuEditSmpl.DropDownItems.Add(mnuSickLists);

            //-----Добавление пунктов "Редактирование простое" к меню------

            //---Добавление пунктов "Редактирование комплексное" к меню----

            //Добавляем в меню "Ред. комп." элемент "Преподаватели"
            mnuEditComp.DropDownItems.Add(mnuLecturers);
            //Добавляем в меню "Ред. комп." элемент "Специальность"
            mnuEditComp.DropDownItems.Add(mnuSpec);
            //Добавляем в меню "Ред. комп." элемент "Доп. работа"
            mnuEditComp.DropDownItems.Add(mnuAddWork);
            //Добавляем в меню "Ред. комп." элемент "Распределение"
            mnuEditComp.DropDownItems.Add(mnuDistrib);
            //Добавляем в меню "Ред. комп." элемент "Расписание"
            mnuEditComp.DropDownItems.Add(mnuSchedule);

            //---Добавление пунктов "Редактирование комплексное" к меню----

            //---------Добавление пунктов "Документооборот" к меню---------

            //Добавляем в меню "Документооборот" элемент "Индивидуальные планы"
            mnuDocs.DropDownItems.Add(mnuPlans);
            //Добавляем в меню "Документооборот" элемент "Заявка на расписание"
            mnuDocs.DropDownItems.Add(mnuDispatch);
            //Добавляем в меню "Документооборот" элемент "Ведомости в учебное управление"
            mnuDocs.DropDownItems.Add(mnuUMU);
            //Добавляем в меню "Документооборот" элемент "Замены по болезни"
            mnuDocs.DropDownItems.Add(mnuSwap);
            //Добавляем в меню "Документооборот" элемент "Для заведующего кафедрой"
            mnuDocs.DropDownItems.Add(mnuForChief);
            //Добавляем в меню "Документооборот" элемент "Протоколы кафедры"
            mnuDocs.DropDownItems.Add(mnuProtocols);

            //---------Добавление пунктов "Документооборот" к меню---------

            //------------Добавление пунктов "Расчёты" к меню--------------

            //Добавляем в меню "Расчёты" элемент "Калькулятор ВКР"
            mnuCount.DropDownItems.Add(mnuVKR);
            //Добавляем в меню "Расчёт" элемент "Расчёт нагрузки по направлениям"
            mnuCount.DropDownItems.Add(mnuCountNapr);
            //Добавляем в меню "Расчёт" элемент "Расчёт нагрузки по преп/напр"
            mnuCount.DropDownItems.Add(mnuCountLectNapr);
            //Добавляем в меню "Расчёт" элемент "Перечень внешних"
            mnuCount.DropDownItems.Add(mnuFindOuterLect);
            //Добавляем в меню "Расчёт" элемент "Доля нагрузки внешних"
            mnuCount.DropDownItems.Add(mnuCountOuter);
            //Добавляем в меню "Расчёт" элемент "Распределяемая почасовая"
            mnuCount.DropDownItems.Add(mnuHouredDistrib);

            //------------Добавление пунктов "Расчёты" к меню--------------

            //------------Добавление пунктов "В разработке" к меню---------

            //Добавляем в меню "В разработке" элемент "Visio Расписание"
            mnuConstruction.DropDownItems.Add(mnuVisTimeTable);
            //Добавляем в меню "В разработке" элемент "Visio Индивидуальное расписание"
            mnuConstruction.DropDownItems.Add(mnuVisIndTimeTable);
            //Добавляем в меню "В разработке" элемент "Visio Тест"
            mnuConstruction.DropDownItems.Add(mnuVisTest);
            //Добавляем в меню "В разработке" элемент "Импорт Word"
            mnuConstruction.DropDownItems.Add(mnuWordImport);

            //------------Добавление пунктов "В разработке" к меню---------

            //------Добавление функциональных элементов меню в полосу------

            //Добавляем меню "Файл" в полосу меню
            mnuMain.Items.Add(mnuFile);
            //Добавляем меню "Редактирование простое" в полосу меню
            mnuMain.Items.Add(mnuEditSmpl);
            //Добавляем меню "Редактирование комплексное" в полосу меню
            mnuMain.Items.Add(mnuEditComp);
            //Добавляем меню "Документооборот" в полосу меню
            mnuMain.Items.Add(mnuDocs);
            //Добавляем меню "Расчёты" в полосу меню
            mnuMain.Items.Add(mnuCount);
            //Добавляем меню "В разработке" в полосу меню
            mnuMain.Items.Add(mnuConstruction);

            //------Добавление функциональных элементов меню в полосу------

            //Размещаем полосу меню на главной форме
            mnuMain.Parent = this;
        }

        private void MnuVisTest_Click(object sender, EventArgs e)
        {
            onVisioTest();
        }

        //Нажатие на кнопку "Импорт Word"
        private void MnuWordImport_Click(object sender, EventArgs e)
        {
            //
            mdlData.Reg = 1;

            //Фильтрация
            dlgOpen.Filter = "Документ Word 2007+|*.docx|Документ Word 97-2003|*.doc|Все файлы|*.*";
            //Ввод заголовка диалогового окна загрузки документа
            dlgOpen.Title = "Открыть документ Word...";
            //Указание на начало работы с корневой директорией диска С
            dlgOpen.InitialDirectory = Application.StartupPath;
            //Вывод/демонстрация диалогового окна для открытия файла с документом

            //Если не нажата отмена в диалоговом окне выбора файла
            if (dlgOpen.ShowDialog() != DialogResult.Cancel)
            {
                //Создаём новое Word приложение
                Word._Application ObjWord = new Word.Application();
                //Добавляем новый чистый документ Word
                Word._Document ObjDoc = ObjWord.Application.Documents.Add();

                //Имя открываемого файла включая полный путь
                object filename = dlgOpen.FileName;
                //Открываем указанный документ Word
                ObjDoc = ObjWord.Documents.Open(FileName: ref filename);

                //Пробегаем все таблицы Word
                foreach (Word.Table tab in ObjDoc.Tables)
                {
                    //Пробегаем все строки таблицы Word
                    foreach (Word.Row row in tab.Rows)
                    {
                        //Интересны только те строки, в которых количество столбцов
                        //совпадает с количеством столбцов таблицы
                        if (row.Cells.Count.Equals(tab.Columns.Count))
                        {
                            //Создаём новый объект распределения, полученного из файла
                            clsDistributionFile DF = new clsDistributionFile();
                            //Начинаем инициализацию объекта
                            DF.Init(row);
                            mdlData.colDistributionFiles.Add(DF);
                        }
                    }
                }

                ObjDoc.Close();
            }

            MessageBox.Show("Данные из таблицы Microsoft Office Word загружены!", "Успех");
        }


        private void MnuHouredDistrib_Click(object sender, EventArgs e)
        {
            string str = "";

            int Sum = 0;
            int tmp;
            clsDistribution Dst;
            clsStudents Std;
            clsLecturer Lec;

            for (int i = 0; i < mdlData.colLecturer.Count; i++)
            {
                Lec = mdlData.colLecturer[i];
                for (int j = 0; j < mdlData.colDistribution.Count; j++)
                {
                    Dst = mdlData.colDistribution[j];
                    if (Dst.Lecturer != null)
                    {

                    }
                    else
                    {
                        //Если равномерно распределяемая нагрузка
                        if (Dst.flgDistrib)
                        {
                            tmp = 0;
                            for (int k = 0; k < mdlData.colStudents.Count; k++)
                            {
                                Std = mdlData.colStudents[k];
                                if (Lec.FIO.Equals(Std.Lect.FIO))
                                {
                                    if (Dst.Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        if (Dst.KursNum.Equals(mdlData.colStudents[k].KursNum) &
                                             Dst.Speciality.Equals(mdlData.colStudents[k].Speciality))
                                        {
                                            if (Std.flgHoured)
                                            {
                                                Sum += Dst.Weight;
                                                tmp += Dst.Weight;
                                            }
                                        }
                                    }

                                    if (Dst.Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        if (Dst.KursNum.Equals(mdlData.colStudents[k].KursNum) &
                                             Dst.Speciality.Equals(mdlData.colStudents[k].Speciality))
                                        {
                                            if (Std.flgHoured)
                                            {
                                                Sum += Dst.Weight;
                                                tmp += Dst.Weight;
                                            }
                                        }
                                    }
                                }
                            }

                            if (tmp > 0)
                            {
                                str += Lec.FIO + " - " + Dst.Subject.Subject + " - " + tmp + "\n\n";
                            }
                        }
                    }
                }
            }

            str += "Распределяемая почасовая нагрузка составляет: " + Sum;

            MessageBox.Show(str, "Расчёт нагрузки по направлениям");
        }

        void mnuCountOuter_Click(object sender, EventArgs e)
        {
            string str = "";
            string strLec = "";

            int SpAll = 0;
            int BachAll = 0;
            int MagAll = 0;
            int OthAll = 0;

            int SpOuter = 0;
            int BachOuter = 0;
            int MagOuter = 0;
            int OthOuter = 0;

            int SpInner = 0;
            int BachInner = 0;
            int MagInner = 0;
            int OthInner = 0;

            int SpHoured = 0;
            int BachHoured = 0;
            int MagHoured = 0;
            int OthHoured = 0;

            int tmpSum = 0;

            bool flgHaveLoad = false;

            clsLecturer Lec = null;
            clsDistribution Dst = null;

            //Расчёт суммарной нагрузки по направлениям обучения
            for (int i = 0; i < mdlData.colDistribution.Count; i++)
            {
                Dst = mdlData.colDistribution[i];
                //
                if (!Dst.flgExclude)
                {
                    if (Dst.Speciality != null)
                    {
                        switch (Dst.Speciality.Diff)
                        {
                            case "Б":
                                {
                                    BachAll += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }

                            case "УО":
                                {
                                    BachAll += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }

                            case "СУР":
                                {
                                    BachAll += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }

                            case "М":
                                {
                                    MagAll += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }

                            case "С":
                                {
                                    SpAll += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }
                            default:
                                {
                                    OthAll += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }
                        }
                    }
                    else
                    {
                        OthAll += mdlData.toSumDistributionComponents(Dst);
                    }
                }
            }

            str = "Специалитет: " + SpAll + "\n" +
                    "Бакалавриат: " + BachAll + "\n" +
                    "Магистратура: " + MagAll + "\n" +
                    "Прочее: " + OthAll + "\n\n" +
                    "Суммарно: " + (SpAll + BachAll + MagAll + OthAll);

            MessageBox.Show(str, "Расчёт нагрузки по направлениям");

            str = "";
            strLec = "Учтены: \n\n";
            
            for (int i = 0; i < mdlData.colLecturer.Count; i++)
            {
                Lec = mdlData.colLecturer[i];
                if (Lec.Combination.CombType.Equals("внешний") & Lec.Rate > 0)
                {
                    strLec += mdlData.SplitFIOString(Lec.FIO, true, false) + ", ";
                    //
                    for (int j = 0; j < mdlData.colDistribution.Count; j++)
                    {
                        Dst = mdlData.colDistribution[j];

                        //Если элемент не исключён из расчёта
                        if (!Dst.flgExclude)
                        {
                            //Если преподаватель указан
                            if (Dst.Lecturer != null)
                            {
                                //Если выбранный преподаватель указан в нагрузке
                                if (Dst.Lecturer.FIO.Equals(Lec.FIO))
                                {
                                    //Если специальность указана для элемента нагрузки
                                    if (Dst.Speciality != null)
                                    {
                                        switch (Dst.Speciality.Diff)
                                        {
                                            case "Б":
                                                {
                                                    BachOuter += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }

                                            case "УО":
                                                {
                                                    BachOuter += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }

                                            case "СУР":
                                                {
                                                    BachOuter += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }

                                            case "М":
                                                {
                                                    MagOuter += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }

                                            case "С":
                                                {
                                                    SpOuter += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }
                                            default:
                                                {
                                                    OthOuter += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }
                                        }
                                    }
                                    //
                                    else
                                    {
                                        OthOuter += mdlData.toSumDistributionComponents(Dst);
                                    }
                                }
                            }
                            //Если преподаватель не указан
                            else
                            {
                                tmpSum = 0;
                                //Если распределяемый элемент нагрузки
                                if (Dst.flgDistrib)
                                {
                                    for (int k = 0; k <= mdlData.colStudents.Count - 1; k++)
                                    {
                                        if (mdlData.colStudents[k].Lect.FIO.Equals(Lec.FIO))
                                        {
                                            if (Dst.Semestr.Equals(mdlData.colSemestr[1]))
                                            {
                                                if (Dst.KursNum.Equals(mdlData.colStudents[k].KursNum) &
                                                     Dst.Speciality.Equals(mdlData.colStudents[k].Speciality))
                                                {
                                                    tmpSum += Dst.Weight;
                                                }
                                            }

                                            if (Dst.Semestr.Equals(mdlData.colSemestr[2]))
                                            {
                                                if (Dst.KursNum.Equals(mdlData.colStudents[k].KursNum) &
                                                     Dst.Speciality.Equals(mdlData.colStudents[k].Speciality))
                                                {
                                                    tmpSum += Dst.Weight;
                                                }
                                            }
                                        }
                                    }
                                }

                                //Если специальность указана для элемента нагрузки
                                if (Dst.Speciality != null)
                                {
                                    switch (Dst.Speciality.Diff)
                                    {
                                        case "Б":
                                            {
                                                BachOuter += tmpSum;
                                                break;
                                            }

                                        case "УО":
                                            {
                                                BachOuter += tmpSum;
                                                break;
                                            }

                                        case "СУР":
                                            {
                                                BachOuter += tmpSum;
                                                break;
                                            }

                                        case "М":
                                            {
                                                MagOuter += tmpSum;
                                                break;
                                            }

                                        case "С":
                                            {
                                                SpOuter += tmpSum;
                                                break;
                                            }
                                        default:
                                            {
                                                OthOuter += tmpSum;
                                                break;
                                            }
                                    }
                                }
                                //
                                else
                                {
                                    OthOuter += tmpSum;
                                }
                            }
                        }
                    }
                }
            }

            strLec += "\n\n";

            str += "Специалитет часов: " + SpOuter + "\n" +
                   "Специалитет % общей нагрузки: " + ((Convert.ToDouble(SpOuter) / Convert.ToDouble(SpAll)) * 100).ToString("0.000") + "%\n\n" +

                   "Бакалавриат часов: " + BachOuter + "\n" +
                   "Бакалавриат % общей нагрузки: " + ((Convert.ToDouble(BachOuter) / Convert.ToDouble(BachAll)) * 100).ToString("0.000") + "%\n\n" +

                   "Магистратура часов: " + MagOuter + "\n" +
                   "Магистратура % общей нагрузки: " + ((Convert.ToDouble(MagOuter) / Convert.ToDouble(MagAll)) * 100).ToString("0.000") + "%\n\n" +

                   "Прочее часов: " + OthOuter + "\n" +
                   "Прочее % общей нагрузки: " + ((Convert.ToDouble(OthOuter) / Convert.ToDouble(OthAll)) * 100).ToString("0.000") + "%\n\n" +

                   "Суммарно часов: " + (SpOuter + BachOuter + MagOuter + OthOuter) + "\n" +
                   "% общей нагрузки: " + ((Convert.ToDouble(SpOuter + BachOuter + MagOuter + OthOuter) / Convert.ToDouble(SpAll + BachAll + MagAll + OthAll)) * 100).ToString("0.000") + "%\n\n" +

                   "Суммарно по кафедре: " + (SpAll + BachAll + MagAll + OthAll);

            MessageBox.Show(strLec + str, "Расчёт нагрузки внешних");

            str = "";
            strLec = "Учтены: \n\n";

            for (int i = 0; i < mdlData.colLecturer.Count; i++)
            {
                Lec = mdlData.colLecturer[i];
                if (!Lec.Combination.CombType.Equals("внешний") & Lec.Rate > 0)
                {
                    strLec += mdlData.SplitFIOString(Lec.FIO, true, false) + ", ";
                    //
                    for (int j = 0; j < mdlData.colDistribution.Count; j++)
                    {
                        Dst = mdlData.colDistribution[j];

                        //Если элемент не исключён из расчёта
                        if (!Dst.flgExclude)
                        {
                            //Если преподаватель указан
                            if (Dst.Lecturer != null)
                            {
                                //Если выбранный преподаватель указан в нагрузке
                                if (Dst.Lecturer.FIO.Equals(Lec.FIO))
                                {
                                    //Если специальность указана для элемента нагрузки
                                    if (Dst.Speciality != null)
                                    {
                                        switch (Dst.Speciality.Diff)
                                        {
                                            case "Б":
                                                {
                                                    BachInner += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }

                                            case "УО":
                                                {
                                                    BachInner += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }

                                            case "СУР":
                                                {
                                                    BachInner += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }

                                            case "М":
                                                {
                                                    MagInner += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }

                                            case "С":
                                                {
                                                    SpInner += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }
                                            default:
                                                {
                                                    OthInner += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }
                                        }
                                    }
                                    //
                                    else
                                    {
                                        OthInner += mdlData.toSumDistributionComponents(Dst);
                                    }
                                }
                            }
                            //Если преподаватель не указан
                            else
                            {
                                tmpSum = 0;
                                //Если распределяемый элемент нагрузки
                                if (Dst.flgDistrib)
                                {
                                    for (int k = 0; k <= mdlData.colStudents.Count - 1; k++)
                                    {
                                        if (mdlData.colStudents[k].Lect.FIO.Equals(Lec.FIO))
                                        {
                                            if (Dst.Semestr.Equals(mdlData.colSemestr[1]))
                                            {
                                                if (Dst.KursNum.Equals(mdlData.colStudents[k].KursNum) &
                                                     Dst.Speciality.Equals(mdlData.colStudents[k].Speciality))
                                                {
                                                    tmpSum += Dst.Weight;
                                                }
                                            }

                                            if (Dst.Semestr.Equals(mdlData.colSemestr[2]))
                                            {
                                                if (Dst.KursNum.Equals(mdlData.colStudents[k].KursNum) &
                                                     Dst.Speciality.Equals(mdlData.colStudents[k].Speciality))
                                                {
                                                    tmpSum += Dst.Weight;
                                                }
                                            }
                                        }
                                    }
                                }

                                //Если специальность указана для элемента нагрузки
                                if (Dst.Speciality != null)
                                {
                                    switch (Dst.Speciality.Diff)
                                    {
                                        case "Б":
                                            {
                                                BachInner += tmpSum;
                                                break;
                                            }

                                        case "УО":
                                            {
                                                BachInner += tmpSum;
                                                break;
                                            }

                                        case "СУР":
                                            {
                                                BachInner += tmpSum;
                                                break;
                                            }

                                        case "М":
                                            {
                                                MagInner += tmpSum;
                                                break;
                                            }

                                        case "С":
                                            {
                                                SpInner += tmpSum;
                                                break;
                                            }
                                        default:
                                            {
                                                OthInner += tmpSum;
                                                break;
                                            }
                                    }
                                }
                                //
                                else
                                {
                                    OthInner += tmpSum;
                                }
                            }
                        }
                    }
                }
            }

            strLec += "\n\n";

            str += "Специалитет часов: " + SpInner + "\n" +
                   "Специалитет % общей нагрузки: " + ((Convert.ToDouble(SpInner) / Convert.ToDouble(SpAll)) * 100).ToString("0.000") + "%\n\n" +

                   "Бакалавриат часов: " + BachInner + "\n" +
                   "Бакалавриат % общей нагрузки: " + ((Convert.ToDouble(BachInner) / Convert.ToDouble(BachAll)) * 100).ToString("0.000") + "%\n\n" +

                   "Магистратура часов: " + MagInner + "\n" +
                   "Магистратура % общей нагрузки: " + ((Convert.ToDouble(MagInner) / Convert.ToDouble(MagAll)) * 100).ToString("0.000") + "%\n\n" +

                   "Прочее часов: " + OthInner + "\n" +
                   "Прочее % общей нагрузки: " + ((Convert.ToDouble(OthInner) / Convert.ToDouble(OthAll)) * 100).ToString("0.000") + "%\n\n" +

                   "Суммарно часов: " + (SpInner + BachInner + MagInner + OthInner) + "\n" +
                   "% общей нагрузки: " + ((Convert.ToDouble(SpInner + BachInner + MagInner + OthInner) / Convert.ToDouble(SpAll + BachAll + MagAll + OthAll)) * 100).ToString("0.000") + "%\n\n" +

                   "Суммарно по кафедре: " + (SpAll + BachAll + MagAll + OthAll);

            MessageBox.Show(strLec + str, "Расчёт нагрузки штатных");

            str = "";
            strLec = "Учтены: \n\n";

            for (int i = 0; i < mdlData.colLecturer.Count; i++)
            {
                flgHaveLoad = false;
                Lec = mdlData.colLecturer[i];
                if (Lec.Rate == 0)
                {
                    //
                    for (int j = 0; j < mdlData.colDistribution.Count; j++)
                    {
                        Dst = mdlData.colDistribution[j];

                        //Если элемент не исключён из расчёта
                        if (!Dst.flgExclude)
                        {
                            //Если преподаватель указан
                            if (Dst.Lecturer != null)
                            {
                                //Если выбранный преподаватель указан в нагрузке
                                if (Dst.Lecturer.FIO.Equals(Lec.FIO))
                                {
                                    //Если специальность указана для элемента нагрузки
                                    if (Dst.Speciality != null)
                                    {
                                        if (mdlData.toSumDistributionComponents(Dst) > 0)
                                        {
                                            flgHaveLoad = true;
                                        }

                                        switch (Dst.Speciality.Diff)
                                        {
                                            case "Б":
                                                {
                                                    BachHoured += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }

                                            case "УО":
                                                {
                                                    BachHoured += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }

                                            case "СУР":
                                                {
                                                    BachHoured += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }

                                            case "М":
                                                {
                                                    MagHoured += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }

                                            case "С":
                                                {
                                                    SpHoured += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }
                                            default:
                                                {
                                                    OthHoured += mdlData.toSumDistributionComponents(Dst);
                                                    break;
                                                }
                                        }
                                    }
                                    //
                                    else
                                    {
                                        OthHoured += mdlData.toSumDistributionComponents(Dst);
                                    }
                                }
                            }
                            //Если преподаватель не указан
                            else
                            {
                                tmpSum = 0;
                                //Если распределяемый элемент нагрузки
                                if (Dst.flgDistrib)
                                {
                                    for (int k = 0; k <= mdlData.colStudents.Count - 1; k++)
                                    {
                                        if (mdlData.colStudents[k].Lect.FIO.Equals(Lec.FIO))
                                        {
                                            if (Dst.Semestr.Equals(mdlData.colSemestr[1]))
                                            {
                                                if (Dst.KursNum.Equals(mdlData.colStudents[k].KursNum) &
                                                     Dst.Speciality.Equals(mdlData.colStudents[k].Speciality))
                                                {
                                                    tmpSum += Dst.Weight;
                                                }
                                            }

                                            if (Dst.Semestr.Equals(mdlData.colSemestr[2]))
                                            {
                                                if (Dst.KursNum.Equals(mdlData.colStudents[k].KursNum) &
                                                     Dst.Speciality.Equals(mdlData.colStudents[k].Speciality))
                                                {
                                                    tmpSum += Dst.Weight;
                                                }
                                            }
                                        }
                                    }
                                }

                                if (tmpSum > 0)
                                {
                                    flgHaveLoad = true;
                                }

                                //Если специальность указана для элемента нагрузки
                                if (Dst.Speciality != null)
                                {
                                    switch (Dst.Speciality.Diff)
                                    {
                                        case "Б":
                                            {
                                                BachHoured += tmpSum;
                                                break;
                                            }

                                        case "УО":
                                            {
                                                BachHoured += tmpSum;
                                                break;
                                            }

                                        case "СУР":
                                            {
                                                BachHoured += tmpSum;
                                                break;
                                            }

                                        case "М":
                                            {
                                                MagHoured += tmpSum;
                                                break;
                                            }

                                        case "С":
                                            {
                                                SpHoured += tmpSum;
                                                break;
                                            }
                                        default:
                                            {
                                                OthHoured += tmpSum;
                                                break;
                                            }
                                    }
                                }
                                //
                                else
                                {
                                    OthHoured += tmpSum;
                                }
                            }
                        }
                    }

                    if (flgHaveLoad)
                    {
                        strLec += mdlData.SplitFIOString(Lec.FIO, true, false) + ", ";
                    }
                }
            }

            strLec += "\n\n";

            str += "Специалитет часов: " + SpHoured + "\n" +
                   "Специалитет % общей нагрузки: " + ((Convert.ToDouble(SpHoured) / Convert.ToDouble(SpAll)) * 100).ToString("0.000") + "%\n\n" +

                   "Бакалавриат часов: " + BachHoured + "\n" +
                   "Бакалавриат % общей нагрузки: " + ((Convert.ToDouble(BachHoured) / Convert.ToDouble(BachAll)) * 100).ToString("0.000") + "%\n\n" +

                   "Магистратура часов: " + MagHoured + "\n" +
                   "Магистратура % общей нагрузки: " + ((Convert.ToDouble(MagHoured) / Convert.ToDouble(MagAll)) * 100).ToString("0.000") + "%\n\n" +

                   "Прочее часов: " + OthHoured + "\n" +
                   "Прочее % общей нагрузки: " + ((Convert.ToDouble(OthHoured) / Convert.ToDouble(OthAll)) * 100).ToString("0.000") + "%\n\n" +

                   "Суммарно часов: " + (SpHoured + BachHoured + MagHoured + OthHoured) + "\n" +
                   "% общей нагрузки: " + ((Convert.ToDouble(SpHoured + BachHoured + MagHoured + OthHoured) / Convert.ToDouble(SpAll + BachAll + MagAll + OthAll)) * 100).ToString("0.000") + "%\n\n" +

                   "Суммарно по кафедре: " + (SpAll + BachAll + MagAll + OthAll);

            MessageBox.Show(strLec + str, "Расчёт нагрузки почасовых");
        }

        void mnuFindOuterLect_Click(object sender, EventArgs e)
        {
            string str = "";
            clsLecturer Lec = null;

            for (int i = 0; i < mdlData.colLecturer.Count; i++)
            {
                Lec = mdlData.colLecturer[i];
                if (Lec.Combination.CombType.Equals("внешний") & Lec.Rate > 0)
                {
                    str += mdlData.SplitFIOString(Lec.FIO, false, false) + ", ";
                }
            }

            MessageBox.Show(str, "Перечень внешних совместителей");
        }

        void mnuCountLectNapr_Click(object sender, EventArgs e)
        {
            string str = "";
            clsDistribution Dst = null;

            int SpOne = 0;
            int BachOne = 0;
            int MagOne = 0;
            int OthOne = 0;

            int SpAll = 0;
            int BachAll = 0;
            int MagAll = 0;
            int OthAll = 0;

            int tmpSum = 0;

            mdlData.toGenerateForm(this, new frmLectInput());

            //Расчёт суммарной нагрузки по направлениям обучения
            for (int i = 0; i < mdlData.colDistribution.Count; i++)
            {
                Dst = mdlData.colDistribution[i];
                //
                if (!Dst.flgExclude)
                {
                    if (Dst.Speciality != null)
                    {
                        switch (Dst.Speciality.Diff)
                        {
                            case "Б":
                                {
                                    BachAll += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }

                            case "УО":
                                {
                                    BachAll += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }

                            case "СУР":
                                {
                                    BachAll += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }

                            case "М":
                                {
                                    MagAll += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }

                            case "С":
                                {
                                    SpAll += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }
                            default:
                                {
                                    OthAll += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }
                        }
                    }
                    else
                    {
                        OthAll += mdlData.toSumDistributionComponents(Dst);
                    }
                }
            }

            //
            for (int i = 0; i < mdlData.colDistribution.Count; i++)
            {
                Dst = mdlData.colDistribution[i];

                //Если элемент не исключён из расчёта
                if (!Dst.flgExclude)
                {
                    //Если преподаватель указан
                    if (Dst.Lecturer != null)
                    {
                        //Если выбранный преподаватель указан в нагрузке
                        if (Dst.Lecturer.FIO.Equals(mdlData.SelectedLecturer.FIO))
                        {
                            //Если специальность указана для элемента нагрузки
                            if (Dst.Speciality != null)
                            {
                                switch (Dst.Speciality.Diff)
                                {
                                    case "Б":
                                        {
                                            BachOne += mdlData.toSumDistributionComponents(Dst);
                                            break;
                                        }

                                    case "УО":
                                        {
                                            BachOne += mdlData.toSumDistributionComponents(Dst);
                                            break;
                                        }

                                    case "СУР":
                                        {
                                            BachOne += mdlData.toSumDistributionComponents(Dst);
                                            break;
                                        }

                                    case "М":
                                        {
                                            MagOne += mdlData.toSumDistributionComponents(Dst);
                                            break;
                                        }

                                    case "С":
                                        {
                                            SpOne += mdlData.toSumDistributionComponents(Dst);
                                            break;
                                        }
                                    default:
                                        {
                                            OthOne += mdlData.toSumDistributionComponents(Dst);
                                            break;
                                        }
                                }
                            }
                            //
                            else
                            {
                                OthOne += mdlData.toSumDistributionComponents(Dst);
                            }
                        }
                    }
                    //Если преподаватель не указан
                    else
                    {
                        tmpSum = 0;
                        //Если распределяемый элемент нагрузки
                        if (Dst.flgDistrib)
                        {
                            for (int k = 0; k <= mdlData.colStudents.Count - 1; k++)
                            {
                                if (mdlData.colStudents[k].Lect.FIO.Equals(mdlData.SelectedLecturer.FIO))
                                {
                                    if (Dst.Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        if (Dst.KursNum.Equals(mdlData.colStudents[k].KursNum) &
                                             Dst.Speciality.Equals(mdlData.colStudents[k].Speciality))
                                        {
                                            tmpSum += Dst.Weight;
                                        }
                                    }

                                    if (Dst.Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        if (Dst.KursNum.Equals(mdlData.colStudents[k].KursNum) &
                                             Dst.Speciality.Equals(mdlData.colStudents[k].Speciality))
                                        {
                                            tmpSum += Dst.Weight;
                                        }
                                    }
                                }
                            }
                        }

                        //Если специальность указана для элемента нагрузки
                        if (Dst.Speciality != null)
                        {
                            switch (Dst.Speciality.Diff)
                            {
                                case "Б":
                                    {
                                        BachOne += tmpSum;
                                        break;
                                    }

                                case "УО":
                                    {
                                        BachOne += tmpSum;
                                        break;
                                    }

                                case "СУР":
                                    {
                                        BachOne += tmpSum;
                                        break;
                                    }

                                case "М":
                                    {
                                        MagOne += tmpSum;
                                        break;
                                    }

                                case "С":
                                    {
                                        SpOne += tmpSum;
                                        break;
                                    }
                                default:
                                    {
                                        OthOne += tmpSum;
                                        break;
                                    }
                            }
                        }
                        //
                        else
                        {
                            OthOne += tmpSum;
                        }                        
                    }
                }
            }

            str = mdlData.SelectedLecturer.FIO + "\n\n";

            str += "Специалитет часов: " + SpOne + "\n" +
                   "Специалитет доля ставки: " + (Convert.ToDouble(SpOne) / Convert.ToDouble(mdlData.AverageLoad)) + "\n" +
                   "Специалитет % общей нагрузки: " + (Convert.ToDouble(SpOne) / Convert.ToDouble(SpAll)) * 100 + "\n\n" +

                   "Бакалавриат часов: " + BachOne + "\n" +
                   "Бакалавриат доля ставки: " + (Convert.ToDouble(BachOne) / Convert.ToDouble(mdlData.AverageLoad)) + "\n" +
                   "Бакалавриат % общей нагрузки: " + (Convert.ToDouble(BachOne) / Convert.ToDouble(BachAll)) * 100 + "\n\n" +

                   "Магистратура часов: " + MagOne + "\n" +
                   "Магистратура доля ставки: " + (Convert.ToDouble(MagOne) / Convert.ToDouble(mdlData.AverageLoad)) + "\n" +
                   "Магистратура % общей нагрузки: " + (Convert.ToDouble(MagOne) / Convert.ToDouble(MagAll)) * 100 + "\n\n" +

                   "Прочее часов: " + OthOne + "\n\n" +
                   "Прочее доля ставки: " + (Convert.ToDouble(OthOne) / Convert.ToDouble(mdlData.AverageLoad)) + "\n" +
                   "Прочее % общей нагрузки: " + (Convert.ToDouble(OthOne) / Convert.ToDouble(OthAll)) * 100 + "\n\n" +

                   "Суммарно часов: " + (SpOne + BachOne + MagOne + OthOne) + "\n" +
                   "Доля ставки: " + ( Convert.ToDouble(SpOne + BachOne + MagOne + OthOne) / Convert.ToDouble(mdlData.AverageLoad)) + "\n" +
                   "% общей нагрузки: " + (Convert.ToDouble(SpOne + BachOne + MagOne + OthOne) / Convert.ToDouble(SpAll + BachAll + MagAll + OthAll)) * 100 + "\n\n" +

                   "Суммарно по кафедре: " + (SpAll + BachAll + MagAll + OthAll);

            MessageBox.Show(str, "Расчёт нагрузки по преп/напр");            
        }

        void mnuCountNapr_Click(object sender, EventArgs e)
        {
            string str = "";
            clsDistribution Dst = null;

            int Sp = 0;
            int Bach = 0;
            int Mag = 0;
            int Oth = 0;

            //
            for (int i = 0; i < mdlData.colDistribution.Count; i++)
            {
                Dst = mdlData.colDistribution[i];
                //
                if (!Dst.flgExclude)
                {
                    if (Dst.Speciality != null)
                    {
                        switch (Dst.Speciality.Diff)
                        {
                            case "Б":
                                {
                                    Bach += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }

                            case "УО":
                                {
                                    Bach += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }

                            case "СУР":
                                {
                                    Bach += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }

                            case "М":
                                {
                                    Mag += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }

                            case "С":
                                {
                                    Sp += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }
                            default:
                                {
                                    Oth += mdlData.toSumDistributionComponents(Dst);
                                    break;
                                }
                        }
                    }
                    else
                    {
                        Oth += mdlData.toSumDistributionComponents(Dst);
                    }
                }
            }

            str = "Специалитет: " + Sp + "\n" +
                "Бакалавриат: " + Bach + "\n" +
                "Магистратура: " + Mag + "\n" +
                "Прочее: " + Oth + "\n\n" +
                "Суммарно: " + (Sp + Bach + Mag + Oth);

            MessageBox.Show(str, "Расчёт нагрузки по направлениям");
        }

        private void MnuSickLists_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmEditSickList());
        }

        void mnuPGStudents_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmPGStudents());
        }

        void mnuProtocols_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmDepProtocol());
        }

        void mnuForChief_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmScheduleManagement());
        }

        void mnuSwap_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmPrepSwap());
        }

        void mnuUMU_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmScheduleManagement());
        }

        void mnuDispatch_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmScheduleManagement());
        }

        void mnuPlans_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmLecturerID());
        }

        void mnuVKR_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmGAKCalc());
        }

        void mnuSubjTypes_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmEditSubjectTypes());
        }

        void mnuStudents_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmEditStudents());
        }

        void mnuWeekDay_Click(object sender, EventArgs e)
        {
            
        }

        void mnuKursNum_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmEditKursNum());
        }

        void mnuSubject_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmEditSubject());
        }

        void mnuSemestr_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmEditSemestr());
        }

        void mnuWorkYear_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmEditWorkYear());
        }

        void mnuFaculty_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmEditFaculty());
        }

        void mnuComb_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmEditCombination());
        }

        void mnuDuty_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmEditDuty());
        }

        void mnuStatus_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmEditStatus());
        }

        void mnuDegree_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmEditDegree());
        }

        void mnuDepart_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmEditDepartment());
        }

        void mnuParams_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmParams());
        }

        void mnuSchedule_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmLectSchedule());
        }

        void mnuDistrib_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmEditDistribution());
        }

        void mnuAddWork_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmDopWork());
        }

        void mnuSpec_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmEditSpeciality());
        }

        void mnuLecturers_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmEditLecturer());
        }

        void mnuVisTimeTable_Click(object sender, EventArgs e)
        {
            //Составление таблицы расписания Visio
            onVisioTimeTableSimple();
        }

        void mnuVisIndTimeTable_Click(object sender, EventArgs e)
        {
            //Открытие формы вабора преподавателя для составления расписания в Visio
            toGenerateForm(new frmSelectLect());
        }

        /// <summary>
        /// Настройка начального состояния элементов управления формы
        /// </summary>
        private void InitMain(bool flg)
        {
            //Без подгруженной базы данных
            //Ничего кроме закрытия программы и открытия
            //(а также работы с отлаживаемыми функциями)
            //базы данных сделать нельзя
            if (flg)
            {
                mnuFile.Visible = true;
                mnuReset.Visible = true;
                mnuReset.Enabled = true;
                mnuSave.Visible = true;
                mnuSave.Enabled = true;

                mnuEditSmpl.Visible = true;
                mnuEditComp.Visible = true;
                mnuDocs.Visible = true;
                mnuCount.Visible = true;
                mnuConstruction.Visible = true;
            }
            else
            {
                mnuFile.Visible = true;
                mnuReset.Visible = false;
                mnuReset.Enabled = false;
                mnuSave.Visible = true;
                mnuSave.Enabled = false;
                mnuEditSmpl.Visible = false;
                mnuEditComp.Visible = false;
                mnuDocs.Visible = false;
                mnuCount.Visible = false;
                mnuConstruction.Visible = true;
            }
        }

        void mnuSave_Click(object sender, EventArgs e)
        {
            if (mdlData.flgLoad)
            {
                if (MessageBox.Show(this, "Вы уверены в необходимости сохранения?",
                    "Сохранить?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    onSave();
                }
            }
        }

        void mnuOpen_Click(object sender, EventArgs e)
        {
            if (mdlData.flgChange)
            {
                if (MessageBox.Show(this, "У Вас имеются несохранённые изменения. Продолжить?",
                    "Смена источника данных?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    onLoad();
                }
            }
            else
            {
                onLoad();
            }
        }

        void mnuReset_Click(object sender, EventArgs e)
        {
            if (mdlData.flgReady)
            {
                if (MessageBox.Show(this, "Вы уверены, что хотите сбросить результаты выполненной работы?",
                    "Сброс в исходное состояние?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    ResetState();
                    mdlData.statString = "Выполнен сброс";
                    lblStatus.Text = mdlData.statString;
                }
            }
        }

        //Прекращение работы с приложением
        void mnuExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        /// <summary>
        /// Процедура заполнения пустыми множествами коллекций по количеству строк соответствующих
        /// таблиц базы данных
        /// </summary>
        /// <param name="c">Переменная связи с Базой Данных</param>
        private void CreateCollections(ref bool ready)
        {
            //--------------------------------01. Считываем параметры
            //Если строка параметров существует
            if (tabParams.Rows.Count > 0)
            {
                //Если в базе указана средняя нагрузка
                if (tabParams.Rows[0]["Средняя_нагрузка"] != null)
                {
                    //Если средняя нагрузка - не пустая строка
                    if (!tabParams.Rows[0]["Средняя_нагрузка"].Equals(""))
                    {
                        //Если средняя нагрузка больше нуля, то
                        if (Convert.ToInt32(tabParams.Rows[0]["Средняя_нагрузка"]) > 0)
                        {
                            //пишем ту, которая заложена в базу данных
                            mdlData.AverageLoad = Convert.ToInt32(tabParams.Rows[0]["Средняя_нагрузка"]);
                        }
                        //иначе пишем стандартную среднюю нагрузку
                        else
                        {
                            mdlData.AverageLoad = 750;
                        }
                    }
                    //Если пустая строка - пишем стандартную среднюю нагрузку
                    else
                    {
                        mdlData.AverageLoad = 750;
                    }
                }
                //Если не указана - ставим стандартную среднюю нагрузку
                else
                {
                    mdlData.AverageLoad = 750;
                }
                
                //Если в базе указано название ведомства
                ControlParam(ref mdlData.MinistryName, "Название_ведомства", "ФАЖТ");

                //Если в базе указан префикс вуза
                ControlParam(ref mdlData.UniversityPrefName, "Префикс_вуза", "ФГБ ОУ ВПО");

                //Если в базе указано название вуза
                ControlParam(ref mdlData.UniversityName, "Название_вуза");

                //Если в базе указан суффикс вуза
                ControlParam(ref mdlData.UniversitySuffName, "Суффикс_вуза", "(МГУПС (МИИТ)");

                //Если в базе указано название кафедры
                ControlParam(ref mdlData.DepartmentName, "Название_кафедры", "УиЗИ");

                //Если в базе указано 
                ControlParamDbl(ref mdlData.PaymentAssist, "Оплата_асс");

                //Если в базе указано 
                ControlParamDbl(ref mdlData.PaymentStPrep, "Оплата_ст_преп");
                
                //Если в базе указано 
                ControlParamDbl(ref mdlData.PaymentDocent, "Оплата_доц");
                
                //Если в базе указано 
                ControlParamDbl(ref mdlData.PaymentProff, "Оплата_проф");
            }
            //Если строки параметров не существует
            else
            {
                //Ставим стандартную среднюю нагрузку
                mdlData.AverageLoad = 750;
                //Ставим стандартное название ведомства
                mdlData.MinistryName = "ФАЖТ";
                //Ставим стандартную аббревиатуру префикса
                mdlData.UniversityPrefName = "ФГБ ОУ ВПО";
                //Ставим кавычки вместо названия вуза
                mdlData.UniversityName = "\"\"";
                //Ставим стандартную аббревиатуру в качестве суффикса
                mdlData.UniversitySuffName = "(МГУПС (МИИТ)";
                //Ставим стандартную аббревиатуру в название кафедры
                mdlData.DepartmentName = "УиЗИ";
                //Ставим стандартную величину оплаты труда ассистента
                mdlData.PaymentAssist = 0d;
                //Ставим стандартную величину оплаты труда старшего преподавателя
                mdlData.PaymentStPrep = 0d;
                //Ставим стандартную величину оплаты труда доцента
                mdlData.PaymentDocent = 0d;
                //Ставим стандартную величину оплаты труда профессора
                mdlData.PaymentProff = 0d;
            }
            //--------------------------------01. Считываем параметры
            
            //--------------------------------02. Создаём коллекцию учебных годов
            mdlData.colWorkYear = InnerCreateCol<clsWorkYear>(tabWorkYear);
            //--------------------------------03. Создаём коллекцию семестров
            mdlData.colSemestr = InnerCreateCol<clsSemestr>(tabSemestr);
            //--------------------------------04. Создаём коллекцию номеров недели
            mdlData.colWeek = InnerCreateCol<clsWeek>(tabNumberWeek);
            //--------------------------------05. Создаём коллекцию дней недели
            mdlData.colWeekDays = InnerCreateCol<clsWeekDays>(tabWeekDays);
            //--------------------------------06. Создаём коллекцию времён пар
            mdlData.colPairTime = InnerCreateCol<clsPairTime>(tabPairTime);
            //--------------------------------07. Создаём коллекцию аудиторий
            mdlData.colAuditory = InnerCreateCol<clsAuditory>(tabAuditory);
            //--------------------------------08. Создаём коллекцию дисциплин
            mdlData.colSubject = InnerCreateCol<clsSubject>(tabSubject);
            //--------------------------------09. Создаём коллекцию номеров курсов
            mdlData.colKursNum = InnerCreateCol<clsKursNum>(tabKursNum);
            //--------------------------------10. Создаём коллекцию видов занятий
            mdlData.colSubjectType = InnerCreateCol<clsSubjectType>(tabSubjectTypes);
            //--------------------------------11. Создаём коллекцию должностей
            mdlData.colDuty = InnerCreateCol<clsDuty>(tabDuty);
            //--------------------------------12. Создаём коллекцию совместительства
            mdlData.colCombination = InnerCreateCol<clsCombination>(tabCombination);
            //--------------------------------13. Создаём коллекцию званий
            mdlData.colStatus = InnerCreateCol<clsStatus>(tabStatus);
            //--------------------------------14. Создаём коллекцию степеней
            mdlData.colDegree = InnerCreateCol<clsDegree>(tabDegree);
            //--------------------------------15. Создаём коллекцию кафедр
            mdlData.colDepart = InnerCreateCol<clsDepartment>(tabDeparment);
            //--------------------------------16. Создаём коллекцию факультетов
            mdlData.colFaculty = InnerCreateCol<clsFaculty>(tabFaculty);
            //--------------------------------17. Создаём коллекцию специальностей
            mdlData.colSpecialisation = InnerCreateCol<clsSpecialisation>(tabSpecialisation);
            //--------------------------------18. Создаём коллекцию студенческих групп
            mdlData.colStudGroup = InnerCreateCol<clsStudGroup>(tabStudentGroups);
            //--------------------------------19. Создаём коллекцию преподавателей
            mdlData.colLecturer = InnerCreateCol<clsLecturer>(tabLecturer);
            //--------------------------------20. Создаём коллекцию штатной нагрузки
            mdlData.colDistribution = InnerCreateCol<clsDistribution>(tabDistribution);
            //--------------------------------21. Создаём коллекцию почасовой нагрузки
            mdlData.colHouredDistribution = InnerCreateCol<clsDistribution>(tabHouredDistribution);
            //--------------------------------22. Создаём коллекцию дополнительной работы
            mdlData.colDopWork = InnerCreateCol<clsDopWork>(tabDopWork);
            //--------------------------------23. Создаём коллекцию вопросов заседаний кафедры
            mdlData.colQuestions = InnerCreateCol<clsQuestions>(tabQuestions);
            //--------------------------------24. Создаём коллекцию расписания преподавателей
            mdlData.colSchedule = InnerCreateCol<clsSchedule>(tabSchedule);
            //--------------------------------25. Создаём коллекцию студентов
            mdlData.colStudents = InnerCreateCol<clsStudents>(tabStudents);
            //--------------------------------26. Создаём коллекцию итогов
            mdlData.colSummary = InnerCreateCol<clsSummary>(tabSummary);
            //--------------------------------27. Создаём коллекцию аспирантов
            mdlData.colPGStudents = InnerCreateCol<clsPGStudents>(tabPGStudents);
            //--------------------------------28. Создаём коллекцию сконвертированной нагрузки
            mdlData.colDistributionDetailed = InnerCreateCol<clsDistributionDetailed>(tabDistribConv);
            //--------------------------------29. Создаём коллекцию больничных листов
            mdlData.colSickList = InnerCreateCol<clsSickList>(tabSickList);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Param"></param>
        /// <param name="tabName"></param>
        /// <param name="stdVal"></param>
        private void ControlParam(ref string Param, string tabName, string stdVal = "\"\"")
        {
            //Если в базе указано название ведомства
            if (tabParams.Rows[0][tabName] != null)
            {
                //Если название ведомства не пустая строка
                if (!tabParams.Rows[0][tabName].Equals(""))
                {
                    //то пишем название ведомства таким, какое оно заложено в базу
                    Param = tabParams.Rows[0][tabName].ToString();
                }
                //В ином случае формируем стандартную аббревиатуру
                else
                {
                    Param = stdVal;
                }
            }
            //Если не указано - формируем стандартную аббревиатуру
            else
            {
                Param = stdVal;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Param"></param>
        /// <param name="tabName"></param>
        /// <param name="stdVal"></param>
        private void ControlParamDbl(ref double Param, string tabName, double stdVal = 0d)
        {
            //Если в базе указана оплата
            if (tabParams.Rows[0][tabName] != null)
            {
                if (tabParams.Rows[0][tabName] != DBNull.Value)
                {
                    //Если оплата не пустая строка
                    if (!tabParams.Rows[0][tabName].Equals(""))
                    {
                        //то фиксируем величину оплаты, как заложено в базу
                        Param = Convert.ToDouble(tabParams.Rows[0][tabName]);
                    }
                    //В ином случае формируем стандартную аббревиатуру
                    else
                    {
                        Param = stdVal;
                    }
                }
                //
                else
                {
                    Param = stdVal;
                }
            }
            //Если не указано - формируем стандартную аббревиатуру
            else
            {
                Param = stdVal;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="someType"></typeparam>
        /// <param name="Tab"></param>
        /// <returns></returns>
        private List<someType> InnerCreateCol<someType>(DataTable Tab) where someType : new()
        {
            IList<someType> cObj = new List<someType>();

            for (int i = 1; i <= Tab.Rows.Count; i++)
            {
                someType Obj = new someType();
                cObj.Add(Obj);
            }

            return cObj.ToList<someType>();
        }

        /// <summary>
        /// Процедура связи созданных коллекций
        /// </summary>
        /// <param name="c">Переменная связи с Базой Данных</param>
        private void LinkingCollections(ref bool ready)
        {
            //--------------------------------02. Заполняем коллекцию учебных годов
            for (int i = 1; i <= mdlData.colWorkYear.Count; i++)
                mdlData.colWorkYear[i - 1].Initialize(tabWorkYear, i - 1);
            //--------------------------------03. Заполняем коллекцию семестров
            for (int i = 1; i <= mdlData.colSemestr.Count; i++)
                mdlData.colSemestr[i - 1].Initialize(tabSemestr, i - 1);     
            //--------------------------------04. Заполняем коллекцию номеров недель
            for (int i = 1; i <= mdlData.colWeek.Count; i++)
                mdlData.colWeek[i - 1].Initialize(tabNumberWeek, i - 1);
            //--------------------------------05. Заполняем коллекцию дни недель
            for (int i = 1; i <= mdlData.colWeekDays.Count; i++)
                mdlData.colWeekDays[i - 1].Initialize(tabWeekDays, i - 1);
            //--------------------------------06. Заполняем коллекцию времён пар
            for (int i = 1; i <= mdlData.colPairTime.Count; i++)
                mdlData.colPairTime[i - 1].Initialize(tabPairTime, i - 1);
            //--------------------------------07. Заполняем коллекцию аудиторий
            for (int i = 1; i <= mdlData.colAuditory.Count; i++)
                mdlData.colAuditory[i - 1].Initialize(tabAuditory, i - 1);
            //--------------------------------08. Заполняем коллекцию дисциплин
            for (int i = 1; i <= mdlData.colSubject.Count; i++)
                mdlData.colSubject[i - 1].Initialize(tabSubject, i - 1);
            //--------------------------------09. Заполняем коллекцию номеров курсов
            for (int i = 1; i <= mdlData.colKursNum.Count; i++)
                mdlData.colKursNum[i - 1].Initialize(tabKursNum, i - 1);
            //--------------------------------10. Заполняем коллекцию типов занятий
            for (int i = 1; i <= mdlData.colSubjectType.Count; i++)
                mdlData.colSubjectType[i - 1].Initialize(tabSubjectTypes, i - 1);
            //--------------------------------11. Заполняем коллекцию должностей
            for (int i = 1; i <= mdlData.colDuty.Count; i++)
                mdlData.colDuty[i - 1].Initialize(tabDuty, i - 1);
            //--------------------------------12. Заполняем коллекцию совместительства
            for (int i = 1; i <= mdlData.colCombination.Count; i++)
                mdlData.colCombination[i - 1].Initialize(tabCombination, i - 1);
            //--------------------------------13. Заполняем коллекцию званий
            for (int i = 1; i <= mdlData.colStatus.Count; i++)
                mdlData.colStatus[i - 1].Initialize(tabStatus, i - 1);
            //--------------------------------14. Заполняем коллекцию степеней
            for (int i = 1; i <= mdlData.colDegree.Count; i++)
                mdlData.colDegree[i - 1].Initialize(tabDegree, i - 1);
            //--------------------------------15. Заполняем коллекцию кафедр
            for (int i = 1; i <= mdlData.colDepart.Count; i++)
                mdlData.colDepart[i - 1].Initialize(tabDeparment, i - 1);
            //--------------------------------16. Заполняем коллекцию факультетов
            for (int i = 1; i <= mdlData.colFaculty.Count; i++)
                mdlData.colFaculty[i - 1].Initialize(tabFaculty, i - 1);
            //--------------------------------17. Заполняем коллекцию специальностей
            for (int i = 1; i <= mdlData.colSpecialisation.Count; i++)
                mdlData.colSpecialisation[i - 1].Initialize(tabSpecialisation, i - 1);
            //--------------------------------18. Заполняем коллекцию студенческих групп
            for (int i = 1; i <= mdlData.colStudGroup.Count; i++)
                mdlData.colStudGroup[i - 1].Initialize(tabStudentGroups, i - 1);
            //--------------------------------19. Заполняем коллекцию преподавателей
            for (int i = 1; i <= mdlData.colLecturer.Count; i++)
                mdlData.colLecturer[i - 1].Initialize(tabLecturer, i - 1);
            //--------------------------------20. Заполняем коллекцию штатной нагрузки
            for (int i = 1; i <= mdlData.colDistribution.Count; i++)
                mdlData.colDistribution[i - 1].Initialize(tabDistribution, i - 1);
            for (int i = 1; i <= mdlData.colDistribution.Count; i++)
                mdlData.colDistribution[i - 1].SelfLinking(tabDistribution, i - 1);
            //--------------------------------21. Заполняем коллекцию почасовой нагрузки
            for (int i = 1; i <= mdlData.colHouredDistribution.Count; i++)
                mdlData.colHouredDistribution[i - 1].Initialize(tabHouredDistribution, i - 1);
            for (int i = 1; i <= mdlData.colHouredDistribution.Count; i++)
                mdlData.colHouredDistribution[i - 1].SelfLinking(tabHouredDistribution, i - 1);
            //--------------------------------20 и 21. Перекрёстные ссылки
            for (int i = 1; i <= mdlData.colDistribution.Count; i++)
                mdlData.colDistribution[i - 1].CrossLinking(tabDistribution, i - 1, false);
            for (int i = 1; i <= mdlData.colHouredDistribution.Count; i++)
                mdlData.colHouredDistribution[i - 1].CrossLinking(tabHouredDistribution, i - 1, true);
            //--------------------------------22. Заполняем коллекцию дополнительной работы
            for (int i = 1; i <= mdlData.colDopWork.Count; i++)
                mdlData.colDopWork[i - 1].Initialize(tabDopWork, i - 1);
            //--------------------------------23. Заполняем коллекцию вопросов заседаний кафедры
            for (int i = 1; i <= mdlData.colQuestions.Count; i++)
                mdlData.colQuestions[i - 1].Initialize(tabQuestions, i - 1);
            //--------------------------------24. Заполняем коллекцию расписания преподавателей
            for (int i = 1; i <= mdlData.colSchedule.Count; i++)
                mdlData.colSchedule[i - 1].Initialize(tabSchedule, i - 1);
            //--------------------------------25. Заполняем коллекцию студентов
            for (int i = 1; i <= mdlData.colStudents.Count; i++)
                mdlData.colStudents[i - 1].Initialize(tabStudents, i - 1);
            //--------------------------------26. Заполняем коллекцию итогов
            for (int i = 1; i <= mdlData.colSummary.Count; i++)
                mdlData.colSummary[i - 1].Initialize(tabSummary, i - 1);
            //--------------------------------27. Заполняем коллекцию аспирантов
            for (int i = 1; i <= mdlData.colPGStudents.Count; i++)
                mdlData.colPGStudents[i - 1].Initialize(tabPGStudents, i - 1);
            //--------------------------------28. Заполняем коллекцию сконвертированной нагрузки
            for (int i = 1; i <= mdlData.colDistributionDetailed.Count; i++)
                mdlData.colDistributionDetailed[i - 1].Initialize(tabDistribConv, i - 1);
            //--------------------------------29. Заполняем коллекцию больничных листов
            for (int i = 1; i <= mdlData.colSickList.Count; i++)
                mdlData.colSickList[i - 1].Initialize(tabSickList, i - 1);

            //--------------------------------Заполняем коллекцию штатной нагрузки с учётои почасовой нагрузки
            mdlData.toCombineDistribution(mdlData.colDistribution, mdlData.colCombineDistribution, 
                                          mdlData.colHouredDistribution, true);
            //
            mdlData.toCompleteMassSchedule();
            //
            mdlData.toCompleteDopWork();
        }

        /// <summary>
        /// Процедура проверки наличия необходимых таблиц и 
        /// наличия в таблицах Базы Данных необходимых строк
        /// </summary>
        private void CheckFieldsDB(string p, ref bool ready)
        {
            //Создаём переменную общения SQL-запросами с БД
            OleDbCommand sql = new OleDbCommand();
            //Передать соединение команде
            sql.Connection = mdlData.glConn;

            for (int i = 0; i <= mdlBaseStructure.masTabNames.Length - 1; i++)
            {
                InnerCheck(getDataTableByName(mdlBaseStructure.masTabNames[i][0][0]),
                            mdlBaseStructure.masTabNames[i][0][0],
                            mdlBaseStructure.masTabNames[i][3],
                            mdlBaseStructure.masTabNames[i][4], 
                            ref ready);
            }
        }

        private void onSave()
        {
            // Фильтруем либо все файлы, либо только файлы MS Access
            dlgSave.Filter = "Базы данных MS Access 2007|*.accdb|Базы данных MS Access 2003|*.mdb|Все файлы|*.*";
            // Задаём заголовок для формы сохранения данных
            dlgSave.Title = "Сохранить в Базу Данных...";
            // Предупреждение о записи поверх
            dlgSave.OverwritePrompt = true;
            // В качестве директории выставляем путь к открытому файлу
            dlgSave.InitialDirectory = mdlData.DataBasePath;
            // Выводим форму открытия файла
            dlgSave.ShowDialog();

            lblStatus.Text = mdlData.statString;
        }

        /// <summary>
        /// Нажатие на кнопку "ОК" диалогового окна сохранения
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dlgSave_FileOk(object sender, CancelEventArgs e)
        {
            mdlData.flgReady = true;
            mdlData.DataBaseSavePath = dlgSave.FileName;
            SaveDB(mdlData.DataBaseSavePath, ref mdlData.flgReady);
        }

        /// <summary>
        /// Сохранение в базу данных
        /// </summary>
        /// <param name="Path"></param>
        /// <param name="ready"></param>
        public void SaveDB(string Path, ref bool ready)
        {
            //Определяем переменную связи с БД
            OleDbConnection connection = new OleDbConnection();

            if (Path != mdlData.DataBasePath)
            {
                try
                {
                    File.Copy(mdlData.DataBasePath, Path);
                }
                catch { }
            }

            //Определяем местонахождение БД и прописываем провайдер-фразу
            if (Path.EndsWith(".mdb") || Path.EndsWith(".MDB"))
            {
                // про MS Access 2003
                connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + @"data source = " + Path;
            }
            else if (Path.EndsWith(".accdb") || Path.EndsWith(".ACCDB"))
            {
                // про MS Access 2007
                connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" + @"data source = " + Path;
                //Persist Security Info=False;
            }
            //Выставляем флаг готовности
            ready = true;
            mdlData.glConn = connection;
            //Проверяем, есть ли в БД все необходимые таблицы   
            ClearTablesDB(ref ready);

            if (ready)
            {
                mdlData.glConn.Open();
                SavingCollections();
                mdlData.statString = "База сохранена";
                mdlData.glConn.Close();
            }
        }

        /// <summary>
        /// Очистка таблиц базы данных
        /// </summary>
        /// <param name="ready"></param>
        private void ClearTablesDB(ref bool ready)
        {
            //Открываем соединение
            try
            {
                mdlData.glConn.Open();
            }
            catch (OleDbException e)
            {
                MessageBox.Show(e.Message, "Ошибка", MessageBoxButtons.YesNo);
                ready = false;
            }

            if (ready)
            {
                // Пытаемся выполнить очистку содержимого
                try
                {
                    for (int i = 0; i <= mdlBaseStructure.masTabNames.Length - 1; i++)
                    {
                        InnerDelete(mdlBaseStructure.masTabNames[i][0][0], mdlBaseStructure.masTabNames[i][1][0]);
                    }
                }
                //Обрабатываем исключение в случае возникновения ошибки и
                //фиксируем ошибку снятием флага готовности и закрываем соединение
                catch (OleDbException e)
                {
                    MessageBox.Show(e.Message, "Ошибка");
                    ready = false;
                }
                mdlData.glConn.Close();
            }
        }

        private void InnerDelete(string tabName, string Key)
        {
            //Создаём переменную SQL-команды
            OleDbCommand DeleteCommand = new OleDbCommand();
            //Приписываем переменной SQL-команды соединение
            DeleteCommand.Connection = mdlData.glConn;

            DeleteCommand.CommandText = "DELETE FROM [" + tabName + "] WHERE [" + Key + "] > 0";
            // Выполнить SQL-команду
            DeleteCommand.ExecuteNonQuery();
        }

        /// <summary>
        /// Процедура переноса содержимого коллекций в Базу Данных
        /// </summary>
        /// <param name="c">Переменная связи с Базой Данных</param>
        private void SavingCollections()
        {
            int index;

            string TabName;
            string TabAttr;
            string TabVal;
            
            //Формируем SQL-запрос для поиска заданной таблицы
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            //Создаём переменную SQL-запроса
            OleDbCommand command = new OleDbCommand();
            //Конструктор команд
            OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter);
            //Указываем соединение для переменной SQL-запроса
            command.Connection = mdlData.glConn;

            //--------------------------------01. Заполняем таблицу параметров базы данных
            index = 1;
            TabName = mdlBaseStructure.masTabNames[0][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[0][3]);
            TabVal = index.ToString() + 
                     ", " + mdlData.AverageLoad.ToString() +
                     ", '" + mdlData.UniversityName.ToString() + "'" +
                     ", '" + mdlData.DepartmentName.ToString() + "'" +
                     ", '" + mdlData.MinistryName.ToString() + "'" +
                     ", '" + mdlData.UniversityPrefName.ToString() + "'" +
                     ", '" + mdlData.UniversitySuffName.ToString() + "'" +
                     ", " + mdlData.PaymentAssist.ToString().Replace(',', '.') + "" +
                     ", " + mdlData.PaymentStPrep.ToString().Replace(',', '.') + "" +
                     ", " + mdlData.PaymentDocent.ToString().Replace(',', '.') + "" +
                     ", " + mdlData.PaymentProff.ToString().Replace(',', '.') + "";
            command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
            command.ExecuteNonQuery();
            //--------------------------------01. Заполняем таблицу параметров базы данных

            //--------------------------------02. Заполняем таблицу учебных годов
            TabName = mdlBaseStructure.masTabNames[1][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[1][3]);           
            
            for (int i = 1; i <= mdlData.colWorkYear.Count; i++)
            {
                TabVal = mdlData.colWorkYear[i - 1].Save(i); 
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------02. Заполняем таблицу учебных годов

            //--------------------------------03. Заполняем таблицу семестров
            TabName = mdlBaseStructure.masTabNames[2][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[2][3]); 
            
            for (int i = 1; i <= mdlData.colSemestr.Count; i++)
            {
                TabVal = mdlData.colSemestr[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------03. Заполняем таблицу семестров

            //--------------------------------04. Заполняем таблицу номеров недель
            TabName = mdlBaseStructure.masTabNames[3][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[3][3]);

            for (int i = 1; i <= mdlData.colWeek.Count; i++)
            {
                TabVal = mdlData.colWeek[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------04. Заполняем таблицу номеров недель

            //--------------------------------05. Заполняем таблицу дней недели
            TabName = mdlBaseStructure.masTabNames[4][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[4][3]);

            for (int i = 1; i <= mdlData.colWeekDays.Count; i++)
            {
                TabVal = mdlData.colWeekDays[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------05. Заполняем таблицу дней недели

            //--------------------------------06. Заполняем таблицу времён пар
            TabName = mdlBaseStructure.masTabNames[5][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[5][3]);

            for (int i = 1; i <= mdlData.colPairTime.Count; i++)
            {
                TabVal = mdlData.colPairTime[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------06. Заполняем таблицу времён пар

            //--------------------------------07. Заполняем таблицу аудиторий
            TabName = mdlBaseStructure.masTabNames[6][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[6][3]);

            for (int i = 1; i <= mdlData.colAuditory.Count; i++)
            {
                TabVal = mdlData.colAuditory[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------07. Заполняем таблицу аудиторий

            //--------------------------------08. Заполняем таблицу дисциплин
            TabName = mdlBaseStructure.masTabNames[7][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[7][3]);

            for (int i = 1; i <= mdlData.colSubject.Count; i++)
            {
                TabVal = mdlData.colSubject[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------08. Заполняем таблицу дисциплин

            //--------------------------------09. Заполняем таблицу номеров курсов
            TabName = mdlBaseStructure.masTabNames[8][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[8][3]);
            
            for (int i = 1; i <= mdlData.colKursNum.Count; i++)
            {
                TabVal = mdlData.colKursNum[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------09. Заполняем таблицу номеров курсов

            //--------------------------------10. Заполняем таблицу видов занятий
            TabName = mdlBaseStructure.masTabNames[9][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[9][3]);

            for (int i = 1; i <= mdlData.colSubjectType.Count; i++)
            {
                TabVal = mdlData.colSubjectType[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------10. Заполняем таблицу видов занятий

            //--------------------------------11. Заполняем таблицу должностей
            TabName = mdlBaseStructure.masTabNames[10][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[10][3]);            
            
            for (int i = 1; i <= mdlData.colDuty.Count; i++)
            {
                TabVal = mdlData.colDuty[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------11. Заполняем таблицу должностей

            //--------------------------------12. Заполняем таблицу совместительства
            TabName = mdlBaseStructure.masTabNames[11][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[11][3]);  

            for (int i = 1; i <= mdlData.colCombination.Count; i++)
            {
                TabVal = mdlData.colCombination[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------12. Заполняем таблицу совместительства

            //--------------------------------13. Заполняем таблицу званий
            TabName = mdlBaseStructure.masTabNames[12][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[12][3]); 

            for (int i = 1; i <= mdlData.colStatus.Count; i++)
            {
                TabVal = mdlData.colStatus[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------13. Заполняем таблицу званий
            
            //--------------------------------14. Заполняем таблицу степеней
            TabName = mdlBaseStructure.masTabNames[13][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[13][3]); 
            
            for (int i = 1; i <= mdlData.colDegree.Count; i++)
            {
                TabVal = mdlData.colDegree[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------14. Заполняем таблицу степеней

            //--------------------------------15. Заполняем таблицу кафедр
            TabName = mdlBaseStructure.masTabNames[14][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[14][3]); 

            for (int i = 1; i <= mdlData.colDepart.Count; i++)
            {
                TabVal = mdlData.colDepart[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------15. Заполняем таблицу кафедр

            //--------------------------------16. Заполняем таблицу факультетов
            TabName = mdlBaseStructure.masTabNames[15][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[15][3]); 
            
            for (int i = 1; i <= mdlData.colFaculty.Count; i++)
            {
                TabVal = mdlData.colFaculty[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------16. Заполняем таблицу факультетов

            //--------------------------------17. Заполняем таблицу специальностей
            TabName = mdlBaseStructure.masTabNames[16][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[16][3]); 
            
            for (int i = 1; i <= mdlData.colSpecialisation.Count; i++)
            {
                TabVal = mdlData.colSpecialisation[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------17. Заполняем таблицу специальностей

            //--------------------------------18. Заполняем таблицу студенческих групп
            TabName = mdlBaseStructure.masTabNames[17][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[17][3]);

            for (int i = 1; i <= mdlData.colStudGroup.Count; i++)
            {
                TabVal = mdlData.colStudGroup[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------18. Заполняем таблицу студенческих групп

            //--------------------------------19. Заполняем таблицу преподавателей
            TabName = mdlBaseStructure.masTabNames[18][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[18][3]);

            for (int i = 1; i <= mdlData.colLecturer.Count; i++)
            {
                TabVal = mdlData.colLecturer[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------19. Заполняем таблицу преподавателей

            //--------------------------------20. Заполняем таблицу штатной нагрузки
            TabName = mdlBaseStructure.masTabNames[19][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[19][3]);

            for (int i = 1; i <= mdlData.colDistribution.Count; i++)
            {
                TabVal = mdlData.colDistribution[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();

            }
            //--------------------------------20. Заполняем таблицу штатной нагрузки

            //--------------------------------21. Заполняем таблицу почасовой нагрузки
            TabName = mdlBaseStructure.masTabNames[20][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[20][3]);

            for (int i = 1; i <= mdlData.colHouredDistribution.Count; i++)
            {
                TabVal = mdlData.colHouredDistribution[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------21. Заполняем таблицу почасовой нагрузки

            //--------------------------------22. Заполняем таблицу дополнительной работы
            TabName = mdlBaseStructure.masTabNames[21][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[21][3]);

            for (int i = 1; i <= mdlData.colDopWork.Count; i++)
            {
                TabVal = mdlData.colDopWork[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------22. Заполняем таблицу дополнительной работы

            //--------------------------------23. Заполняем таблицу вопросов на заседание кафедры
            TabName = mdlBaseStructure.masTabNames[22][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[22][3]);

            for (int i = 1; i <= mdlData.colQuestions.Count; i++)
            {
                TabVal = mdlData.colQuestions[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------23. Заполняем таблицу вопросов на заседание кафедры

            //--------------------------------24. Заполняем таблицу расписания преподавателей
            TabName = mdlBaseStructure.masTabNames[23][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[23][3]);

            for (int i = 1; i <= mdlData.colSchedule.Count; i++)
            {
                TabVal = mdlData.colSchedule[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------24. Заполняем таблицу расписания преподавателей

            //--------------------------------25. Заполняем таблицу студентов
            TabName = mdlBaseStructure.masTabNames[24][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[24][3]);

            for (int i = 1; i <= mdlData.colStudents.Count; i++)
            {
                TabVal = mdlData.colStudents[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------25. Заполняем таблицу студентов

            //--------------------------------26. Заполняем таблицу итогов
            //---(в данный момент сохраняется по кнопке, не в общем потоке)
            /*
            TabName = mdlBaseStructure.masTabNames[25][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[25][3]);

            for (int i = 1; i <= mdlData.colSummary.Count; i++)
            {
                TabVal = mdlData.colSummary[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            */ 
            //--------------------------------26. Заполняем таблицу итогов

            //--------------------------------27. Заполняем таблицу аспирантов
            TabName = mdlBaseStructure.masTabNames[26][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[26][3]);

            for (int i = 1; i <= mdlData.colPGStudents.Count; i++)
            {
                TabVal = mdlData.colPGStudents[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------27. Заполняем таблицу аспирантов

            //--------------------------------28. Заполняем таблицу сконвертированной нагрузки
            TabName = mdlBaseStructure.masTabNames[27][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[27][3]);

            for (int i = 1; i <= mdlData.colDistributionDetailed.Count; i++)
            {
                TabVal = mdlData.colDistributionDetailed[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------28. Заполняем таблицу сконвертированной нагрузки

            //--------------------------------29. Заполняем таблицу больничных листов
            TabName = mdlBaseStructure.masTabNames[28][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[28][3]);

            for (int i = 1; i <= mdlData.colSickList.Count; i++)
            {
                TabVal = mdlData.colSickList[i - 1].Save(i);
                command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
                command.ExecuteNonQuery();
            }
            //--------------------------------29. Заполняем таблицу больничных листов
        }

        /// <summary>
        /// Первичная очистка существующих коллекций
        /// </summary>
        private void StartClearCollections()
        {
            mdlData.ClearCollections();
        }

        private void StartClearTables()
        {
            tabDeparment.Clear();
            tabDegree.Clear();
            tabStatus.Clear();
            tabCombination.Clear();
            tabDuty.Clear();
            tabLecturer.Clear();
            tabSemestr.Clear();
            tabWorkYear.Clear();
            tabDopWork.Clear();
            tabFaculty.Clear();
            tabSpecialisation.Clear();
            tabKursNum.Clear();
            tabSubject.Clear();
            tabDistribution.Clear();
            tabHouredDistribution.Clear();
        }

        //Сброс формы с исходное состояние
        private void ResetState()
        {
            mdlData.ResetData();
            InitMain(mdlData.flgLoad);
        }

        //Метод вывода дочерней формы
        private void toGenerateForm(Form f)
        {
            if (mdlData.flgReady)
            {
                //Делаем наследование от главной формы
                f.Owner = this;
                //Отображаем форму на экране
                f.ShowDialog();
                //Очищаем память от формы
                f = null;
                //Заменяем строку состояния по итогам работы с формой
                lblStatus.Text = mdlData.statString;
            }
            else
            {
                MessageBox.Show(this, "Пожалуйста, загрузите сначала базу данных", "В доступе к функции отказано!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //Заполнение клеток расписания, основанное на связи расписания преподавателей
        //с конкретными элементами распределения нагрузки
        //(этот механизм ещё не готов)

        private void onVisioTimeTable()
        {
            string visDocName = Application.StartupPath + "\\myVisio.vsd";
            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;
            clsLecturer L;
            clsSchedule Sch1 = new clsSchedule();
            clsSchedule Sch2 = new clsSchedule();
            bool flgFound1Week;
            bool flgFound2Week;

            float currY = 0f;
            float currX = 0f;
            float XStart = 0f;

            Visio.Shape visTextBox;
            Visio.Page visPage;
            Visio.Application visApp;
            Visio.Document visDoc;

            visApp = new Visio.Application();
            visDoc = visApp.Documents.Add("");
            visPage = visApp.ActivePage;

            currY = 10f;
            for (int i = mdlData.colLecturer.Count - 1; i >= 0; i--)
            {
                currX = 10f;
                if (mdlData.colLecturer[i].Rate > 0)
                {
                    //Рисуем прямоугольник под Ф.И.О.
                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                       ((currX + 80f) / 25.4f), ((currY + 20f) / 25.4f));
                    //Вписываем Ф.И.О. рассматриваемого преподавателя
                    visTextBox.Text = mdlData.colLecturer[i].FIO;

                    currX += 80f;

                    //Запускаем цикл по дням недели
                    for (int j = 0; j < 5; j++)
                    {
                        //Запускаем цикл по временам пар
                        //до обеда
                        for (int k = 0; k < 3; k++)
                        {
                            visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                        ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));
                            currX += 20f;
                        }

                        //Обед
                        visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                ((currX + 10f) / 25.4f), ((currY + 20f) / 25.4f));
                        currX += 10f;

                        //Запускаем цикл по временам пар
                        //после обеда
                        for (int k = 3; k < 8; k++)
                        {
                            visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                        ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));
                            currX += 20f;
                        }
                    }

                    currY += 20f;
                }
            }

            currX = 10f;
            //Рисуем прямоугольник под надпись Ф.И.О.
            visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                               ((currX + 80f) / 25.4f), ((currY + 40f) / 25.4f));
            //Вписываем Ф.И.О. рассматриваемого преподавателя
            visTextBox.Text = "Преподаватель";

            currX += 80f;

            //Запускаем цикл по дням недели
            for (int j = 0; j < 5; j++)
            {
                XStart = currX;
                for (int k = 0; k < 3; k++)
                {
                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));
                    if (mdlData.colPairTime.Count > 0)
                    {
                        visTextBox.Text = mdlData.colPairTime[k].Time;
                    }
                    else
                    {
                        visTextBox.Text = "--:-- - --:--";
                    }
                    currX += 20f;
                }

                visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                        ((currX + 20f) / 25.4f), ((currY + 10f) / 25.4f));

                visTextBox.Text = "Обед";
                visTextBox.Rotate90();

                visApp.ActiveWindow.DeselectAll();
                visApp.ActiveWindow.Select(visTextBox, 2);
                visApp.ActiveWindow.Selection.Move(-(5f / 25.4f), (5f / 25.4f));

                currX += 10f;

                for (int k = 3; k < 8; k++)
                {
                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));
                    if (mdlData.colPairTime.Count > 0)
                    {
                        visTextBox.Text = mdlData.colPairTime[k].Time;
                    }
                    else
                    {
                        visTextBox.Text = "--:-- - --:--";
                    }
                    currX += 20f;
                }

                visTextBox = visPage.DrawRectangle((XStart / 25.4f), ((currY + 20f) / 25.4f),
                             ((currX) / 25.4f), ((currY + 40f) / 25.4f));

                if (mdlData.colWeekDays.Count > 0)
                {
                    visTextBox.Text = mdlData.colWeekDays[j].WeekDay;
                }
                else
                {
                    visTextBox.Text = "№ " + (j + 1).ToString();
                }
            }

            currY = 10f;
            for (int i = mdlData.colLecturer.Count - 1; i >= 0; i--)
            {
                currX = 10f;

                if (mdlData.colLecturer[i].Rate > 0)
                {
                    L = mdlData.colLecturer[i];

                    currX += 80f;

                    //Запускаем цикл по дням недели
                    for (int j = 0; j < 5; j++)
                    {
                        //запускаем цикл по временам занятий
                        //до обеда
                        for (int k = 0; k < 3; k++)
                        {
                            flgFound1Week = DetectTimeTableElement(ref Sch1, L, 2, 0, j, k);
                            flgFound2Week = DetectTimeTableElement(ref Sch2, L, 2, 1, j, k);

                            //Если по каждой неделе нашёлся элемент расписания
                            if (flgFound1Week & flgFound2Week)
                            {
                                visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                             ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                visTextBox.LineStyle = "None";
                                visTextBox.FillStyle = "None";
                                visTextBox.Text = "Есть";
                            }
                            else
                            {
                                //Если элемент расписания нашёлся только для одной из недель
                                if (flgFound1Week || flgFound2Week)
                                {
                                    //Элемент для первой недели
                                    if (flgFound1Week)
                                    {
                                        visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move(-(5f / 25.4f), (10f / 25.4f));

                                        if (Sch1.Link != null)
                                        {
                                            visTextBox.Text = Sch1.Link.Speciality.ShortInstitute + "-" + Sch1.Link.KursNum.Kurs
                                                                + " (" + Sch1.Link.SubjType.Short + ") " + Sch1.Auditory;
                                        }
                                        else
                                        {
                                            visTextBox.Text = "Есть";
                                        }
                                    }

                                    visTextBox = visPage.DrawLine((currX / 25.4f), (currY / 25.4f),
                                                 ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                    //Элемент для второй недели
                                    if (flgFound2Week)
                                    {
                                        visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move((5f / 25.4f), (0f / 25.4f));

                                        if (Sch2.Link != null)
                                        {
                                            visTextBox.Text = Sch2.Link.Speciality.ShortInstitute + "-" + Sch2.Link.KursNum.Kurs
                                                                + " (" + Sch2.Link.SubjType.Short + ") " + Sch2.Auditory;
                                        }
                                        else
                                        {
                                            visTextBox.Text = "Есть";
                                        }
                                    }
                                }
                            }

                            currX += 20f;
                        }

                        //Отступ обеда
                        currX += 10f;

                        //Запускаем цикл по временам пар
                        //после обеда
                        for (int k = 3; k < 8; k++)
                        {
                            flgFound1Week = DetectTimeTableElement(ref Sch1, L, 2, 0, j, k);
                            flgFound2Week = DetectTimeTableElement(ref Sch2, L, 2, 1, j, k);

                            //Если по каждой неделе нашёлся элемент расписания
                            if (flgFound1Week & flgFound2Week)
                            {
                                visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                             ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                visTextBox.LineStyle = "None";
                                visTextBox.FillStyle = "None";
                                visTextBox.Text = "Есть";
                            }
                            else
                            {
                                //Если элемент расписания нашёлся только для одной из недель
                                if (flgFound1Week || flgFound2Week)
                                {
                                    //Элемент для первой недели
                                    if (flgFound1Week)
                                    {
                                        visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move(-(5f / 25.4f), (10f / 25.4f));

                                        if (Sch1.Link != null)
                                        {
                                            visTextBox.Text = Sch1.Link.Speciality.ShortInstitute + "-" + Sch1.Link.KursNum.Kurs
                                                                + " (" + Sch1.Link.SubjType.Short + ") " + Sch1.Auditory;
                                        }
                                        else
                                        {
                                            visTextBox.Text = "Есть";
                                        }
                                    }

                                    visTextBox = visPage.DrawLine((currX / 25.4f), (currY / 25.4f),
                                                 ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                    //Элемент для второй недели
                                    if (flgFound2Week)
                                    {
                                        visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move((5f / 25.4f), (0f / 25.4f));

                                        if (Sch2.Link != null)
                                        {
                                            visTextBox.Text = Sch2.Link.Speciality.ShortInstitute + "-" + Sch2.Link.KursNum.Kurs
                                                                + " (" + Sch2.Link.SubjType.Short + ") " + Sch2.Auditory;
                                        }
                                        else
                                        {
                                            visTextBox.Text = "Есть";
                                        }
                                    }
                                }
                            }

                            currX += 20f;
                        }
                    }

                    currY += 20f;
                }
            }

            visApp.Visible = true;
        }

        private void onVisioTest()
        {
            string visDocName = Application.StartupPath + "\\myVisioTest.vsd";
            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;

            Visio.Shape visTextBox;
            Visio.Page visPage;
            Visio.Application visApp;
            Visio.Document visDoc;

            visApp = new Visio.Application();
            visDoc = visApp.Documents.Add("");
            visPage = visApp.ActivePage;

            //Рисуем прямоугольник под некоторый текст
            visTextBox = visPage.DrawRectangle((10f / 25.4f), (10f / 25.4f),
                                               (80f / 25.4f), (30f / 25.4f));
            //Вписываем некоторый текст 
            visTextBox.Text = "Тестовый текст";
            //Меняем размер шрифта текста
            visTextBox.CellsU["Char.Font"].FormulaForceU = "0";
            //Меняем размер шрифта текста
            visTextBox.CellsU["Char.Size"].FormulaForceU = "14 pt";
            //Меняем цвет шрифта текста в фигуре
            visTextBox.CellsU["Char.Color"].FormulaForceU = "RGB(0,255,0)";
            //Меняем цвет линии фигуры
            visTextBox.CellsU["LineColor"].FormulaForceU = "RGB(255,0,0)";
            //Меняем цвет подложки фигуры
            visTextBox.CellsU["FillForegnd"].FormulaForceU = "RGB(0,0,255)";

            visApp.Visible = true;
        }

        //Заполнение клеток расписания, основанное на информации,
        //указанной вручную через форму ввода
        private void onVisioTimeTableSimple()
        {
            string visDocName = Application.StartupPath + "\\myVisio.vsd";
            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;
            clsLecturer L;
            clsSchedule Sch1 = new clsSchedule();
            clsSchedule Sch2 = new clsSchedule();
            bool flgFound1Week;
            bool flgFound2Week;

            float currY = 0f;
            float currX = 0f;
            float XStart = 0f;

            float XLeft = 10f;
            float XRight = 0f;
            float YBottom = 10f;
            float YTop = 0f;

            int Semestr = 0;
            bool trigColor;

            Visio.Shape visTextBox;
            Visio.Page visPage;
            Visio.Application visApp;
            Visio.Document visDoc;

            visApp = new Visio.Application();
            visDoc = visApp.Documents.Add("");
            visPage = visApp.ActivePage;

            //Вопрос: для какого семестра создаётся расписание?
            if (MessageBox.Show(this, "Создаётся расписание для I семестра?", "Выбор семестра", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Semestr = 1;
            }
            else
            {
                Semestr = 2;
            }

            //---------------С этого момента создаётся пустая таблица----------
            //
            currY = 10f;
            trigColor = false;
            //Перебираем преподавателей
            for (int i = mdlData.colLecturer.Count - 1; i >= 0; i--)
            {
                currX = 10f;
                //if (mdlData.colLecturer[i].Rate > 0 || DetectCheckedSchedule(mdlData.colLecturer[i], Semestr))
                if (DetectCheckedSchedule(mdlData.colLecturer[i], Semestr))
                {
                    //Рисуем прямоугольник под Ф.И.О.
                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                       ((currX + 80f) / 25.4f), ((currY + 20f) / 25.4f));
                    //Вписываем Ф.И.О. рассматриваемого преподавателя
                    visTextBox.Text = mdlData.colLecturer[i].FIO;
                    visTextBox.CellsU["Char.Size"].FormulaForceU = "18 pt";

                    if (trigColor)
                    {
                        visTextBox.CellsU["FillForegnd"].FormulaForceU = "RGB(235,235,235)";
                    }
                    else
                    {
                        visTextBox.CellsU["FillForegnd"].FormulaForceU = "RGB(255,255,255)";
                    }

                    currX += 80f;

                    //Запускаем цикл по дням недели
                    for (int j = 0; j < 5; j++)
                    {
                        //Запускаем цикл по временам пар
                        //до обеда
                        for (int k = 0; k < 3; k++)
                        {
                            visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                        ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                            if (trigColor)
                            {
                                visTextBox.CellsU["FillForegnd"].FormulaForceU = "RGB(235,235,235)";
                            }
                            else
                            {
                                visTextBox.CellsU["FillForegnd"].FormulaForceU = "RGB(255,255,255)";
                            }

                            currX += 20f;
                        }

                        //Обед
                        visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                ((currX + 10f) / 25.4f), ((currY + 20f) / 25.4f));

                        if (trigColor)
                        {
                            visTextBox.CellsU["FillForegnd"].FormulaForceU = "RGB(235,235,235)";
                        }
                        else
                        {
                            visTextBox.CellsU["FillForegnd"].FormulaForceU = "RGB(255,255,255)";
                        }

                        currX += 10f;

                        //Запускаем цикл по временам пар
                        //после обеда
                        for (int k = 3; k < 8; k++)
                        {
                            visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                        ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                            if (trigColor)
                            {
                                visTextBox.CellsU["FillForegnd"].FormulaForceU = "RGB(235,235,235)";
                            }
                            else
                            {
                                visTextBox.CellsU["FillForegnd"].FormulaForceU = "RGB(255,255,255)";
                            }

                            currX += 20f;
                        }


                    }

                    trigColor = !trigColor;
                    currY += 20f;
                }
            }

            currX = 10f;
            //Рисуем прямоугольник под надпись Ф.И.О.
            visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                               ((currX + 80f) / 25.4f), ((currY + 40f) / 25.4f));
            //Вписываем Ф.И.О. рассматриваемого преподавателя
            visTextBox.Text = "Преподаватель";
            visTextBox.CellsU["Char.Size"].FormulaForceU = "24 pt";

            currX += 80f;

            //Запускаем цикл по дням недели
            for (int j = 0; j < 5; j++)
            {
                XStart = currX;
                for (int k = 0; k < 3; k++)
                {
                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));
                    if (mdlData.colPairTime.Count > 0)
                    {
                        visTextBox.Text = mdlData.colPairTime[k].Time;
                    }
                    else
                    {
                        visTextBox.Text = "--:-- - --:--";
                    }
                    currX += 20f;
                }

                visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                        ((currX + 20f) / 25.4f), ((currY + 10f) / 25.4f));
                visTextBox.Text = "Обед";
                visTextBox.Rotate90();

                visApp.ActiveWindow.DeselectAll();
                visApp.ActiveWindow.Select(visTextBox, 2);
                visApp.ActiveWindow.Selection.Move(-(5f / 25.4f), (5f / 25.4f));

                currX += 10f;

                for (int k = 3; k < 8; k++)
                {
                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));
                    if (mdlData.colPairTime.Count > 0)
                    {
                        visTextBox.Text = mdlData.colPairTime[k].Time;
                    }
                    else
                    {
                        visTextBox.Text = "--:-- - --:--";
                    }
                    currX += 20f;
                }

                visTextBox = visPage.DrawRectangle((XStart / 25.4f), ((currY + 20f) / 25.4f),
                             ((currX) / 25.4f), ((currY + 40f) / 25.4f));

                if (mdlData.colWeekDays.Count > 0)
                {
                    visTextBox.Text = mdlData.colWeekDays[j].WeekDay;
                }
                else
                {
                    visTextBox.Text = "№ " + (j + 1).ToString();
                }

                visTextBox.CellsU["Char.Size"].FormulaForceU = "24 pt";
            }

            XRight = currX;
            YTop = currY + 40f;
            YBottom = currY + 60f;

            //Изображаем прямоугольник под текст с заголовком таблицы
            visTextBox = visPage.DrawRectangle((XLeft / 25.4f), (YTop / 25.4f),
                                                     (XRight / 25.4f), ((YBottom) / 25.4f));
            visTextBox.LineStyle = "None";
            visTextBox.FillStyle = "None";

            //Снять выделение со всех элементов
            visApp.ActiveWindow.DeselectAll();
            visApp.ActiveWindow.Select(visTextBox, 2);

            visTextBox.Text = "Расписание преподавателей кафедры \"" + mdlData.DepartmentName + 
                "\" на " + (Semestr == 1 ? "I" : "II") + " семестр " + 
                mdlData.colWorkYear[mdlData.colWorkYear.Count - 2].WorkYear + " учебного года";

            //Пример покраски текста в форме (shape) в красный цвет
            //visTextBox.CellsU["Char.Color"].FormulaForceU = "RGB(255,0,0)";
            visTextBox.CellsU["Char.Size"].FormulaForceU = "48 pt";
            

            //---------------До этого момента создаётся пустая таблица---------

            //---------------С этого момента заполняется таблица---------------
            //
            currY = 10f;
            
            //Перебираем преподавателей
            for (int i = mdlData.colLecturer.Count - 1; i >= 0; i--)
            {
                currX = 10f;

                //if (mdlData.colLecturer[i].Rate > 0 || DetectCheckedSchedule(mdlData.colLecturer[i], Semestr))
                if (DetectCheckedSchedule(mdlData.colLecturer[i], Semestr))
                {
                    L = mdlData.colLecturer[i];

                    currX += 80f;

                    //Запускаем цикл по дням недели
                    for (int j = 0; j < 5; j++)
                    {
                        //запускаем цикл по временам занятий
                        //до обеда
                        for (int k = 0; k < 3; k++)
                        {
                            Sch1 = null;
                            Sch2 = null;
                            flgFound1Week = DetectTimeTableElement(ref Sch1, L, Semestr, 0, j, k);
                            flgFound2Week = DetectTimeTableElement(ref Sch2, L, Semestr, 1, j, k);

                            //Если по каждой неделе нашёлся элемент расписания
                            if (flgFound1Week & flgFound2Week)
                            {
                                //Если хотя бы что-то отсутствует у одного из элементов расписания, то
                                //выводить надпись "Есть"
                                if ((Sch1.KursNum == null || Sch1.Spec == null || Sch1.SubjectType == null) ||
                                    (Sch2.KursNum == null || Sch2.Spec == null || Sch2.SubjectType == null))
                                {
                                    //То просто пишем, что обе пары в это время есть
                                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                    visTextBox.LineStyle = "None";
                                    visTextBox.FillStyle = "None";
                                    visTextBox.Text = "Есть";
                                }
                                //Если какие-то сведения имеются
                                else
                                {
                                    //Если всё в точности совпадает, то выводим горизонтально
                                    //первый попавшийся элемент
                                    if (Sch1.SubjectType.Equals(Sch2.SubjectType) &
                                        Sch1.Spec.Equals(Sch2.Spec) &
                                        Sch1.KursNum.Equals(Sch2.KursNum) &
                                        Sch1.Auditory.Equals(Sch2.Auditory) &
                                        Sch1.Group.Equals(Sch2.Group) &
                                        Sch1.Stream.Equals(Sch2.Stream))
                                    {
                                        //Для горизонтальных надписей
                                        //не требуется расширять блок текста
                                        visTextBox = visPage.DrawRectangle(((currX) / 25.4f), (currY / 25.4f),
                                                    ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";
                                        visTextBox.Text = Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                            (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                            (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                            " (" + Sch1.SubjectType.Short + ")" + "\n" + Sch1.Auditory;
                                    }
                                    //Если хотя бы что-то не совпало, то записываем элементы через дробь
                                    else
                                    {
                                        visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));

                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move(-(3f / 25.4f), (9f / 25.4f));

                                        visTextBox.Text = Sch1.Auditory + "\n" + Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                            (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                            (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                          " (" + Sch1.SubjectType.Short + ")";

                                        visTextBox = visPage.DrawLine(((currX) / 25.4f), (currY / 25.4f),
                                                    ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                        visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move((4f / 25.4f), (1f / 25.4f));

                                        visTextBox.Text = Sch2.Spec.ShortInstitute + "-" + Sch2.KursNum.Kurs +
                                                            (Sch2.Stream.Equals("") ? (Sch2.Spec.Diff == "М" ? "5" : "1") : Sch2.Stream) +
                                                            (Sch2.Group.Equals("") ? "1" : Sch2.Group) +
                                                            " (" + Sch2.SubjectType.Short + ")" + "\n" + Sch2.Auditory;
                                    }
                                }
                            }
                            else
                            {
                                //Если элемент расписания нашёлся только для одной из недель
                                if (flgFound1Week || flgFound2Week)
                                {
                                    //Элемент для первой недели
                                    if (flgFound1Week)
                                    {
                                        visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move(-(3f / 25.4f), (9f / 25.4f));

                                        if ((Sch1.KursNum == null || Sch1.Spec == null || Sch1.SubjectType == null))
                                        {
                                            visTextBox.Text = "Есть";
                                        }
                                        else
                                        {
                                            //
                                            visTextBox.Text = Sch1.Auditory + "\n" + Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                              (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                              (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                              " (" + Sch1.SubjectType.Short + ")";
                                        }
                                    }

                                    visTextBox = visPage.DrawLine((currX / 25.4f), (currY / 25.4f),
                                                 ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                    //Элемент для второй недели
                                    if (flgFound2Week)
                                    {
                                        visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move((4f / 25.4f), (1f / 25.4f));

                                        if (Sch2.KursNum == null || Sch2.Spec == null || Sch2.SubjectType == null)
                                        {
                                            visTextBox.Text = "Есть";
                                        }
                                        else
                                        {
                                            visTextBox.Text = Sch2.Spec.ShortInstitute + "-" + Sch2.KursNum.Kurs +
                                                              (Sch2.Stream.Equals("") ? (Sch2.Spec.Diff == "М" ? "5" : "1") : Sch2.Stream) +
                                                              (Sch2.Group.Equals("") ? "1" : Sch2.Group) +
                                                              " (" + Sch2.SubjectType.Short + ")" + "\n" + Sch2.Auditory;
                                        }
                                    }
                                }
                            }

                            currX += 20f;
                        }

                        //Отступ обеда
                        currX += 10f;

                        //Запускаем цикл по временам пар
                        //после обеда
                        for (int k = 3; k < 8; k++)
                        {
                            Sch1 = null;
                            Sch2 = null;
                            
                            flgFound1Week = DetectTimeTableElement(ref Sch1, L, Semestr, 0, j, k);
                            flgFound2Week = DetectTimeTableElement(ref Sch2, L, Semestr, 1, j, k);

                            //Если по каждой неделе нашёлся элемент расписания
                            if (flgFound1Week & flgFound2Week)
                            {
                                //Если хотя бы что-то отсутствует хотя бы у одного из элементов расписания,
                                //то выводить надпись "Есть"
                                if ((Sch1.KursNum == null || Sch1.Spec == null || Sch1.SubjectType == null) ||
                                    (Sch2.KursNum == null || Sch2.Spec == null || Sch2.SubjectType == null))
                                {
                                    //То просто пишем, что обе пары в это время есть
                                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                    visTextBox.LineStyle = "None";
                                    visTextBox.FillStyle = "None";

                                    visTextBox.Text = "Есть";
                                }
                                //Если какие-то сведения имеются
                                else
                                {
                                    //Если всё в точности совпадает, то выводим горизонтально
                                    //первый попавшийся элемент
                                    if (Sch1.SubjectType.Equals(Sch2.SubjectType) &
                                        Sch1.Spec.Equals(Sch2.Spec) &
                                        Sch1.KursNum.Equals(Sch2.KursNum) &
                                        Sch1.Auditory.Equals(Sch2.Auditory) &
                                        Sch1.Group.Equals(Sch2.Group) &
                                        Sch1.Stream.Equals(Sch2.Stream))
                                    {
                                        //Если надпись горизонтальная, то
                                        //не требуется дополнительно расширять текстовый блок
                                        visTextBox = visPage.DrawRectangle(((currX) / 25.4f), (currY / 25.4f),
                                                    ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";
                                        visTextBox.Text = Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                            (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                            (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                            " (" + Sch1.SubjectType.Short + ")" + "\n" + Sch1.Auditory;
                                    }
                                    //Если хотя бы что-то не совпало, то записываем элементы через дробь
                                    else
                                    {
                                        visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move(-(3f / 25.4f), (9f / 25.4f));

                                        visTextBox.Text = Sch1.Auditory + "\n" + Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                            (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                            (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                            " (" + Sch1.SubjectType.Short + ")";

                                        visTextBox = visPage.DrawLine(((currX) / 25.4f), (currY / 25.4f),
                                                    ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                        visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move((4f / 25.4f), (1f / 25.4f));

                                        visTextBox.Text = Sch2.Spec.ShortInstitute + "-" + Sch2.KursNum.Kurs +
                                                            (Sch2.Stream.Equals("") ? (Sch2.Spec.Diff == "М" ? "5" : "1") : Sch2.Stream) +
                                                            (Sch2.Group.Equals("") ? "1" : Sch2.Group) +
                                                            " (" + Sch2.SubjectType.Short + ")" + "\n" + Sch2.Auditory;
                                    }
                                }
                            }
                            else
                            {
                                //Если элемент расписания нашёлся только для одной из недель
                                if (flgFound1Week || flgFound2Week)
                                {
                                    //Элемент для первой недели
                                    if (flgFound1Week)
                                    {
                                        visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move(-(3f / 25.4f), (9f / 25.4f));

                                        if ((Sch1.KursNum == null || Sch1.Spec == null || Sch1.SubjectType == null))
                                        {
                                            visTextBox.Text = "Есть";
                                        }
                                        else
                                        {
                                            visTextBox.Text = Sch1.Auditory + "\n" + Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                                (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                                (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                                " (" + Sch1.SubjectType.Short + ")";
                                        }
                                    }

                                    visTextBox = visPage.DrawLine((currX / 25.4f), (currY / 25.4f),
                                                 ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                    //Элемент для второй недели
                                    if (flgFound2Week)
                                    {
                                        visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move((4f / 25.4f), (1f / 25.4f));

                                        if (Sch2.KursNum == null || Sch2.Spec == null || Sch2.SubjectType == null)
                                        {
                                            visTextBox.Text = "Есть";
                                        }
                                        else
                                        {
                                            visTextBox.Text = Sch2.Spec.ShortInstitute + "-" + Sch2.KursNum.Kurs +
                                                                (Sch2.Stream.Equals("") ? (Sch2.Spec.Diff == "М" ? "5" : "1") : Sch2.Stream) +
                                                                (Sch2.Group.Equals("") ? "1" : Sch2.Group) +
                                                                " (" + Sch2.SubjectType.Short + ")" + "\n" + Sch2.Auditory;
                                        }
                                    }
                                }
                            }

                            currX += 20f;
                        }
                    }

                    currY += 20f;
                }
            }

            visApp.Visible = true;
        }

        private bool DetectCheckedSchedule(clsLecturer L, int Semestr)
        {
            bool flg = false;

            for (int i = 0; i < mdlData.colSchedule.Count; i++)
            {
                //Если совпали семестры, преподаватели и есть признак занятия
                //хотя бы у одного из элементов расписания
                if (mdlData.colSchedule[i].Lecturer.FIO.Equals(L.FIO) &
                    mdlData.colSchedule[i].Semestr.SemNum.Equals(mdlData.colSemestr[Semestr].SemNum) &
                    mdlData.colSchedule[i].Subj)
                {
                    flg = true;
                    break;
                }
            }

            return flg;
        }

        //Определение наличия интересующего элемента расписания
        private bool DetectTimeTableElement(ref clsSchedule Sch, clsLecturer L, int Semestr, int Week, int WeekDay, int PairTime)
        {
            bool flg = false;
            for (int l = mdlData.colSchedule.Count - 1; l >= 0; l--)
            {
                Sch = mdlData.colSchedule[l];
                //Если совпали преподаватели
                if (Sch.Lecturer.FIO.Equals(L.FIO))
                {
                    //Если совпали семестры
                    if (Sch.Semestr.SemNum.Equals(mdlData.colSemestr[Semestr].SemNum))
                    {
                        //Если совпали учебные недели
                        if (Sch.Week.NumberWeek.Equals(mdlData.colWeek[Week].NumberWeek))
                        {
                            //Если совпали дни недели
                            if (Sch.WeekDay.WeekDay.Equals(mdlData.colWeekDays[WeekDay].WeekDay))
                            {
                                //Если совпало время занятий
                                if (Sch.Time.Time.Equals(mdlData.colPairTime[PairTime].Time))
                                {
                                    if (Sch.Subj)
                                    {
                                        flg = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return flg;
        }

        //Определение интересующего элемента расписания
        private clsSchedule FindTimeTableElement(clsLecturer L, int Week, int WeekDay, int PairTime)
        {
            clsSchedule Sch = null;
            for (int l = mdlData.colSchedule.Count - 1; l >= 0; l--)
            {
                Sch = mdlData.colSchedule[l];
                //Если совпали преподаватели
                if (Sch.Lecturer.FIO.Equals(L.FIO))
                {
                    //Если совпали учебные недели
                    if (Sch.Week.NumberWeek.Equals(mdlData.colWeek[Week].NumberWeek))
                    {
                        //Если совпали дни недели
                        if (Sch.WeekDay.WeekDay.Equals(mdlData.colWeekDays[WeekDay].WeekDay))
                        {
                            //Если совпали дни недели
                            if (Sch.Time.Time.Equals(mdlData.colPairTime[PairTime].Time))
                            {
                                if (Sch.Subj)
                                {
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            return Sch;
        }

        /*
        private void TimeTableDraw()
        {
            string visDocName = Application.StartupPath + "\\myVisio.vsd";
            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;

            float currY = 0f;
            float currX = 0f;
            float XStart = 0f;

            Visio.Shape visTextBox;
            Visio.Page visPage;
            Visio.Application visApp;
            Visio.Document visDoc;

            visApp = new Visio.Application();
            visDoc = visApp.Documents.Add("");
            visPage = visApp.ActivePage;

            currY = 10f;
            for (int i = mdlData.colLecturer.Count - 1; i >= 0; i--)
            {
                currX = 10f;
                if (mdlData.colLecturer[i].Rate > 0)
                {
                    //Рисуем прямоугольник под Ф.И.О.
                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                       ((currX + 80f) / 25.4f), ((currY + 20f) / 25.4f));
                    //Вписываем Ф.И.О. рассматриваемого преподавателя
                    visTextBox.Text = mdlData.colLecturer[i].FIO;

                    currX += 80f;

                    //Запускаем цикл по дням недели
                    for (int j = 0; j < 5; j++)
                    {
                        for (int k = 0; k < 3; k++)
                        {
                            visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                        ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));
                            currX += 20f;
                        }

                        visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                ((currX + 10f) / 25.4f), ((currY + 20f) / 25.4f));
                        currX += 10f;

                        for (int k = 0; k < 5; k++)
                        {
                            visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                        ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));
                            currX += 20f;
                        }
                    }

                    currY += 20f;
                }
            }

            currX = 10f;
            //Рисуем прямоугольник под надпись Ф.И.О.
            visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                               ((currX + 80f) / 25.4f), ((currY + 40f) / 25.4f));
            //Вписываем Ф.И.О. рассматриваемого преподавателя
            visTextBox.Text = "Преподаватель";

            currX += 80f;

            //Запускаем цикл по дням недели
            for (int j = 0; j < 5; j++)
            {
                XStart = currX;
                for (int k = 0; k < 3; k++)
                {
                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));
                    visTextBox.Text = mdlData.colPairTime[k].Time;
                    currX += 20f;
                }

                visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                        ((currX + 20f) / 25.4f), ((currY + 10f) / 25.4f));
                visTextBox.Text = "Обед";
                visTextBox.Rotate90();

                visApp.ActiveWindow.DeselectAll();
                visApp.ActiveWindow.Select(visTextBox, 2);
                visApp.ActiveWindow.Selection.Move(-(5f / 25.4f), (5f / 25.4f));

                currX += 10f;

                for (int k = 0; k < 5; k++)
                {
                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));
                    visTextBox.Text = mdlData.colPairTime[3 + k].Time;
                    currX += 20f;
                }

                visTextBox = visPage.DrawRectangle((XStart / 25.4f), ((currY + 20f) / 25.4f),
                             ((currX) / 25.4f), ((currY + 40f) / 25.4f));
                visTextBox.Text = mdlData.colWeekDays[j].WeekDay;
            }

            visApp.Visible = true;
        }
        */
        
        //Рабочий пример вращения фигуры на 45 градусов
        /*
        private void VisioRotation45()
        {
            string visDocName = Application.StartupPath + "\\myVisio.vsd";
            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;

            float currY;
            float currX;

            Visio.Shape visTextBox;
            Visio.Page visPage;
            Visio.Application visApp;
            Visio.Document visDoc;

            visApp = new Visio.Application();
            visDoc = visApp.Documents.Add("");
            visPage = visApp.ActivePage;

            currX = 10f;
            currY = 10f;
            for (int i = 5; i >= 0; i--)
            {
                //if (mdlData.colLecturer[i].Rate > 0)
                //{
                visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                   ((currX + 70f) / 25.4f), ((currY + 10f) / 25.4f));
                visTextBox.Text = i.ToString();

                visTextBox.FillStyle = "None";
                visTextBox.LineStyle = "None";

                //Снять выделение со всех элементов
                visApp.ActiveWindow.DeselectAll();
                //visDeselect = 1;
                //visSelect = 2;
                //visSubSelect = 3;
                //visSelectAll = 4;
                //visDeselectAll = 256;
                visApp.ActiveWindow.Select(visTextBox, 2);
                //visDegrees = 81;
                visApp.ActiveWindow.Selection.Rotate(45d, 81);
                currY += 10f;
                //}
            }

            visApp.Visible = true;
        }
        */
    }
}