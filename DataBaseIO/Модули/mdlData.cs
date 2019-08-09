using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Word = Microsoft.Office.Interop.Word;

namespace DataBaseIO
{
    class mdlData
    {
        //---------------------------------------------------------------------
        //Группа сопрягающих с БД глобальных переменных
        //---------------------------------------------------------------------

        /// <summary>
        /// Переменная соединения с базой данных
        /// </summary>
        public static OleDbConnection glConn = null;

        /// <summary>
        /// Путь к файлу для загрузки информации из базы данных
        /// </summary>
        public static string DataBasePath = "";

        /// <summary>
        /// Путь сохранения в базу данных
        /// </summary>
        public static string DataBaseSavePath;

        /// <summary>
        /// Признак готовности системы к работе
        /// </summary>
        public static bool flgReady = false;

        /// <summary>
        /// Признак загрузки информации из базы данных
        /// </summary>
        public static bool flgLoad = false;

        /// <summary>
        /// Признак внесённых изменений
        /// </summary>
        public static bool flgChange = false;

        /// <summary>
        /// Признак работы с базой данных старого образца
        /// </summary>
        public static bool flgOldDB = false;

        /// <summary>
        /// Величина догрузки неразгруженных преподавателей
        /// </summary>
        public static double LoadInc = 0d;

        /// <summary>
        /// Величина средней нагрузки по кафедре
        /// </summary>
        public static int AverageLoad = 0;

        /// <summary>
        /// Наименование ведомства, которому принадлежит вуз
        /// </summary>
        public static string MinistryName = "";

        /// <summary>
        /// Величина почасовой оплаты ассистента
        /// </summary>
        public static double PaymentAssist = 0d;

        /// <summary>
        /// Величина почасовой оплаты старшего преподавателя
        /// </summary>
        public static double PaymentStPrep = 0d;

        /// <summary>
        /// Величина почасовой оплаты доцента
        /// </summary>
        public static double PaymentDocent = 0d;

        /// <summary>
        /// Величина почасовой оплаты профессора
        /// </summary>
        public static double PaymentProff = 0d;

        /// <summary>
        /// Наименование вуза
        /// </summary>
        public static string UniversityName = "";

        /// <summary>
        /// Префикс вуза
        /// </summary>
        public static string UniversityPrefName = "";

        /// <summary>
        /// Суффикс вуза
        /// </summary>
        public static string UniversitySuffName = "";

        /// <summary>
        /// Наименование кафедры
        /// </summary>
        public static string DepartmentName = "";

        //---------------------------------------------------------------------
        //Группа сопрягающих с БД глобальных переменных
        //---------------------------------------------------------------------  

        //---------------------------------------------------------------------
        //Группа глобальных переменных, хранящих настройки формы распределения
        //--------------------------------------------------------------------- 

        public static bool flgFacultyFilt = false;
        public static int inxFaculty = -1;
        public static bool flgKursFilt = false;
        public static int inxKurs = -1;
        public static bool flgSpecialityFilt = false;
        public static int inxSpeciality = -1;
        public static bool flgLecturerFilt = false;
        public static int inxLecturer = -1;
        public static bool flgSemestrFilt = false;
        public static int inxSemestr = -1;
        public static bool flgSubjectFilt = false;
        public static int inxSubject = -1;
        public static bool flgTypeFilt = false;
        public static int inxType = -1;
        public static string cmbRem = "";
        public static string statString = "Начало работы";

        //---------------------------------------------------------------------
        //Группа глобальных переменных, хранящих настройки формы распределения
        //--------------------------------------------------------------------- 

        //---------------------------------------------------------------------
        //Группа глобальных переменных, хранящих настройки формы студентов
        //--------------------------------------------------------------------- 

        public static bool flgStudLecturerFilt = false;
        public static int inxStudLecturer = -1;

        //---------------------------------------------------------------------
        //Группа глобальных переменных, хранящих настройки формы студентов
        //--------------------------------------------------------------------- 

        //---------------------------------------------------------------------
        //Упорядоченная группа коллекций
        //---------------------------------------------------------------------        

        /// <summary>
        /// 02. Коллекция учебных годов
        /// </summary>
        public static IList<clsWorkYear> colWorkYear = new List<clsWorkYear>();

        /// <summary>
        /// 03. Коллекция семестров
        /// </summary>
        public static IList<clsSemestr> colSemestr = new List<clsSemestr>();

        /// <summary>
        /// 04. Коллекция номеров недели
        /// </summary>
        public static IList<clsWeek> colWeek = new List<clsWeek>();

        /// <summary>
        /// 05. Коллекция дней недели
        /// </summary>
        public static IList<clsWeekDays> colWeekDays = new List<clsWeekDays>();

        /// <summary>
        /// 06. Коллекция времён проведения пар
        /// </summary>
        public static IList<clsPairTime> colPairTime = new List<clsPairTime>();

        /// <summary>
        /// 07. Коллекция учебных аудиторий
        /// </summary>
        public static IList<clsAuditory> colAuditory = new List<clsAuditory>();

        /// <summary>
        /// 08. Коллекция учебных дисциплин
        /// </summary>
        public static IList<clsSubject> colSubject = new List<clsSubject>();

        /// <summary>
        /// 09. Коллекция номеров учебных курсов
        /// </summary>
        public static IList<clsKursNum> colKursNum = new List<clsKursNum>();

        /// <summary>
        /// 10. Коллекция типов учебных занятий
        /// </summary>
        public static IList<clsSubjectType> colSubjectType = new List<clsSubjectType>();

        /// <summary>
        /// 11. Коллекция должностей
        /// </summary>
        public static IList<clsDuty> colDuty = new List<clsDuty>();

        /// <summary>
        /// 12. Коллекция совместительства
        /// </summary>
        public static IList<clsCombination> colCombination = new List<clsCombination>();

        /// <summary>
        /// 13. Коллекция званий
        /// </summary>
        public static IList<clsStatus> colStatus = new List<clsStatus>();

        /// <summary>
        /// 14. Коллекция степеней
        /// </summary>
        public static IList<clsDegree> colDegree = new List<clsDegree>();
        
        /// <summary>
        /// 15. Коллекция кафедр
        /// </summary>
        public static IList<clsDepartment> colDepart = new List<clsDepartment>();

        /// <summary>
        /// 16. Коллекция факультетов
        /// </summary>
        public static IList<clsFaculty> colFaculty = new List<clsFaculty>();

        /// <summary>
        /// 17. Коллекция специализаций
        /// </summary>
        public static IList<clsSpecialisation> colSpecialisation = new List<clsSpecialisation>();

        /// <summary>
        /// 18. Коллекция студенческих групп
        /// </summary>
        public static IList<clsStudGroup> colStudGroup = new List<clsStudGroup>();
        
        /// <summary>
        /// 19. Коллекция преподавателей
        /// </summary>
        public static IList<clsLecturer> colLecturer = new List<clsLecturer>();

        /// <summary>
        /// 20. Коллекция штатной нагрузки
        /// </summary>
        public static IList<clsDistribution> colDistribution = new List<clsDistribution>();

        /// <summary>
        /// 21.1 Коллекция почасовой нагрузки (фактическая)
        /// </summary>
        public static IList<clsDistribution> colHouredDistribution = new List<clsDistribution>();

        /// <summary>
        /// 21.2 Коллекция почасовой нагрузки (плановая)
        /// </summary>
        public static IList<clsDistribution> colPlanHouredDistribution = new List<clsDistribution>();

        /// <summary>
        /// 22. Коллекция дополнительной работы
        /// </summary>
        public static IList<clsDopWork> colDopWork = new List<clsDopWork>();
       
        /// <summary>
        /// 23. Коллекция вопросов заседаний кафедры
        /// </summary>
        public static IList<clsQuestions> colQuestions = new List<clsQuestions>();

        /// <summary>
        /// 24. Коллекция расписания
        /// </summary>
        public static IList<clsSchedule> colSchedule = new List<clsSchedule>();

        /// <summary>
        /// 25. Коллекция студентов
        /// </summary>
        public static IList<clsStudents> colStudents = new List<clsStudents>();

        /// <summary>
        /// 26. Коллекция итогов
        /// </summary>
        public static IList<clsSummary> colSummary = new List<clsSummary>();
        
        /// <summary>
        /// 27. Коллекция аспирантов
        /// </summary>
        public static IList<clsPGStudents> colPGStudents = new List<clsPGStudents>();

        /// <summary>
        /// 28. Коллекция детализированной (сконвертированной) учебной нагрузки
        /// </summary>
        public static IList<clsDistributionDetailed> colDistributionDetailed = new List<clsDistributionDetailed>();

        /// <summary>
        /// 29. Коллекция больничных листов
        /// </summary>
        public static IList<clsSickList> colSickList = new List<clsSickList>();

        /// <summary>
        /// Коллекция штатной нагрузки с учётом почасовой (фактическая)
        /// </summary>
        public static IList<clsDistribution> colCombineDistribution = new List<clsDistribution>();

        /// <summary>
        /// Коллекция штатной нагрузки с учётом почасовой (плановая)
        /// </summary>
        public static IList<clsDistribution> colPlanCombineDistribution = new List<clsDistribution>();

        /// <summary>
        /// Фильтрованная коллекция нагрузки
        /// </summary>
        public static IList<clsDistribution> Filtred = new List<clsDistribution>();

        /// <summary>
        /// Фильтрованная коллекция студентов
        /// </summary>
        public static IList<clsStudents> FiltredStudents = new List<clsStudents>();

        /// <summary>
        /// Фильтрованная коллекция аспирантов
        /// </summary>
        public static IList<clsPGStudents> FiltredPGStudents = new List<clsPGStudents>();

        /// <summary>
        /// Коллекция преподавателей с указанием нагрузки
        /// </summary>
        public static IList<clsLecturer_Load> colLectLoad = new List<clsLecturer_Load>();

        /// <summary>
        /// Коллекция занятий, пропущенных по болезни
        /// </summary>
        public static IList<clsScheduleSickHours> colScheduleSickHours = new List<clsScheduleSickHours>();

        /// <summary>
        /// Режим (Мария Холод)
        /// </summary>
        public static int Reg = 0;

        /// <summary>
        /// Коллекция распределения, полученная из файла Word (Мария Холод)
        /// </summary>
        public static IList<clsDistributionFile> colDistributionFiles = new List<clsDistributionFile>();

        //---------------------------------------------------------------------
        //Упорядоченная группа коллекций
        //---------------------------------------------------------------------    

        //---------------------------------------------------------------------
        //Выделенные элементы для передачи между формами
        //--------------------------------------------------------------------- 

        public static clsLecturer SelectedLecturer = null;
        public static clsSchedule SelectedScheduleElement = null;

        //---------------------------------------------------------------------
        //Выделенные элементы для передачи между формами
        //--------------------------------------------------------------------- 

        //---------------------------------------------------------------------
        //Резервный раздел
        //---------------------------------------------------------------------    

        ///// <summary>
        ///// Признак состояния вызова формы ввода объекта
        ///// </summary>
        //public static int flgInput;
        ///// <summary>
        ///// Передаваемый для редактирования код элемента
        ///// </summary>
        //public static int currentCode;

        //---------------------------------------------------------------------
        //Резервный раздел
        //---------------------------------------------------------------------  

        //---------------------------------------------------------------------
        //Глобальные функции и процедуры
        //---------------------------------------------------------------------  

        //Сброс данных к исходным значениях
        public static void ResetData()
        {
            //Удаляем глобальное соединение с базой данных
            glConn = null;
            //Удаляем путь к базе данных
            DataBasePath = "";
            //Удаляем путь сохранения в базу данных
            DataBaseSavePath = "";
            //Сбрасываем флаг готовности системы к работе
            flgReady = false;
            //Сбрасываем флаг загрузки базы данных
            flgLoad = false;
            //Сбрасываем флаг глобальных изменений
            flgChange = false;
            //Сбрасываем флаг старой базы данных
            flgOldDB = false;
            //Сбрасываем догрузку
            LoadInc = 0d;
            //Сбрасываем среднюю нагрузку
            AverageLoad = 0;
            //Сбрасываем наименование министерства
            MinistryName = "";
            //Сбрасываем величину оплаты труда ассистента
            PaymentAssist = 0d;
            //Сбрасываем величину оплаты труда старшего преподавател
            PaymentStPrep = 0d;
            //Сбрасываем величину оплаты труда доцента
            PaymentDocent = 0d;
            //Сбрасываем величину оплаты труда профессора
            PaymentProff = 0d;
            //Сбрасываем наименование университета
            UniversityName = "";
            //Сбрасываем префикс университета
            UniversityPrefName = "";
            //Сбрасываем суффикс университета
            UniversitySuffName = "";
            //Сбрасываем наименование кафедры
            DepartmentName = "";
            //Сбрасываем параметры фильтрации нагрузки по факультету
            flgFacultyFilt = false;
            inxFaculty = -1;
            //Сбрасываем параметры фильтрации нагрузки по курсу
            flgKursFilt = false;
            inxKurs = -1;
            //Сбрасываем параметры фильтрации нагрузки по специальности
            flgSpecialityFilt = false;
            inxSpeciality = -1;
            //Сбрасываем параметры фильтрации нагрузки по преподавателю
            flgLecturerFilt = false;
            inxLecturer = -1;
            //Сбрасываем параметры фильтрации нагрузки по семестру
            flgSemestrFilt = false;
            inxSemestr = -1;
            //Сбрасываем параметры фильтрации нагрузки по дисциплине
            flgSubjectFilt = false;
            inxSubject = -1;
            //Сбрасываем параметры фильтрации нагрузки по типу
            flgTypeFilt = false;
            inxType = -1;
            //????????
            cmbRem = "";
            //Переводим надпись статуса в исходное состояние
            statString = "Начало работы";

            //Вычищаем коллекции
            ClearCollections();

            //Сбрасываем глобально выбранного преподавателя
            SelectedLecturer = null;
            //Сбрасываем глобально выбранный элемент расписания
            SelectedScheduleElement = null;
        }

        public static void ClearCollections()
        {
            //Сбрасываем коллекцию учебных лет
            colWorkYear = null;
            colWorkYear = new List<clsWorkYear>();

            //Сбрасываем коллекцию семестров
            colSemestr = null;
            colSemestr = new List<clsSemestr>();

            //Сбрасываем коллекцию недель
            colWeek = null;
            colWeek = new List<clsWeek>();

            //Сбрасываем коллекцию дней недели
            colWeekDays = null;
            colWeekDays = new List<clsWeekDays>();

            //Сбрасываем коллекцию времён занятий
            colPairTime = null;
            colPairTime = new List<clsPairTime>();

            //Сбрасываем коллекцию аудиторий
            colAuditory = null;
            colAuditory = new List<clsAuditory>();

            //Сбрасываем коллекцию дисциплин
            colSubject = null;
            colSubject = new List<clsSubject>();

            //Сбрасываем коллекцию курсов
            colKursNum = null;
            colKursNum = new List<clsKursNum>();

            //Сбрасываем коллекцию видов занятий
            colSubjectType = null;
            colSubjectType = new List<clsSubjectType>();

            //Сбрасываем коллекцию должностей
            colDuty = null;
            colDuty = new List<clsDuty>();

            //Сбрасываем коллекцию совмещения
            colCombination = null;
            colCombination = new List<clsCombination>();

            //Сбрасываем коллекцию званий
            colStatus = null;
            colStatus = new List<clsStatus>();

            //Сбрасываем коллекцию степеней
            colDegree = null;
            colDegree = new List<clsDegree>();

            //Сбрасываем коллекцию кафедр
            colDepart = null;
            colDepart = new List<clsDepartment>();

            //Сбрасываем коллекцию факультетов
            colFaculty = null;
            colFaculty = new List<clsFaculty>();

            //Сбрасываем коллекцию специальностей
            colSpecialisation = null;
            colSpecialisation = new List<clsSpecialisation>();

            //Сбрасываем коллекцию студенческих групп
            colStudGroup = null;
            colStudGroup = new List<clsStudGroup>();

            //Сбрасываем коллекцию преподавателей
            colLecturer = null;
            colLecturer = new List<clsLecturer>();

            //Сбрасываем коллекцию распределения нагрузки
            colDistribution = null;
            colDistribution = new List<clsDistribution>();

            //Сбрасываем коллекцию почасовой нагрузки
            colHouredDistribution = null;
            colHouredDistribution = new List<clsDistribution>();

            //Сбрасываем коллекцию дополнительной работы
            colDopWork = null;
            colDopWork = new List<clsDopWork>();

            //Сбрасываем коллекцию вопросов, выносимых на заседание кафедры
            colQuestions = null;
            colQuestions = new List<clsQuestions>();

            //Сбрасываем коллекцию элементов расписания
            colSchedule = null;
            colSchedule = new List<clsSchedule>();

            //Сбрасываем коллекцию студентов
            colStudents = null;
            colStudents = new List<clsStudents>();

            //Сбрасываем коллекцию сводных итогов
            colSummary = null;
            colSummary = new List<clsSummary>();

            //Сбрасываем коллекцию сводных итогов
            colPGStudents = null;
            colPGStudents = new List<clsPGStudents>();

            //Сбрасываем коллекцию комбинированной нагрузки
            colCombineDistribution = null;
            colCombineDistribution = new List<clsDistribution>();

            //Сбрасываем коллекцию фильтрованной нагрузки
            Filtred = null;
            Filtred = new List<clsDistribution>();

            //Сбрасываем коллекцию фильтрованных студентов
            FiltredStudents = null;
            FiltredStudents = new List<clsStudents>();

            //Сбрасываем коллекцию связок преподаватель-нагрузка
            colLectLoad = null;
            colLectLoad = new List<clsLecturer_Load>();

            //Сбрасываем коллекцию детализированной нагрузки
            colDistributionDetailed = null;
            colDistributionDetailed = new List<clsDistributionDetailed>();
        }

        /// <summary>
        /// Процедура составления комбинированной коллекции распределения нагрузки
        /// </summary>
        /// <param name="InColl">Коллекция исходных данных</param>
        /// <param name="OutColl">Результирующая коллекция</param>
        /// <param name="DiffColl">Коллекция, по которой считается разница для результирующей</param>
        /// <param name="flgComb">Признак необходимости вычисления разницы</param>
        public static void toCombineDistribution(IList<clsDistribution> InColl, IList<clsDistribution> OutColl, 
                                                 IList<clsDistribution> DiffColl, bool flgComb = true)
        {
            clsDistribution D;
            //Очищаем коллекцию штатной нагрузки с учётом почасовой
            OutColl.Clear();
            //Заполняем коллекцию новыми объектами
            for (int i = 0; i <= InColl.Count - 1; i++)
            {
                D = new clsDistribution();
                OutColl.Add(D);
            }
            //Инициализируем коллекцию
            for (int i = 0; i <= InColl.Count - 1; i++)
            {
                OutColl[i].CopyFrom(InColl[i], true);
            }

            //Если выставлен признак необходимости комбинирования нагрузки
            if (flgComb)
            {
                //Проходим в цикле штатную нагрузку
                for (int i = 0; i <= InColl.Count - 1; i++)
                {
                    //Проходим в цикле почасовую нагрузку
                    for (int j = 0; j <= DiffColl.Count - 1; j++)
                    {
                        //Если элементы совпадают по специальности, курсу, преподавателю,
                        //дисциплине и коду документа, то комбинируем
                        if (OutColl[i].Speciality == DiffColl[j].Speciality &
                            OutColl[i].Semestr == DiffColl[j].Semestr &
                            OutColl[i].KursNum == DiffColl[j].KursNum &
                            OutColl[i].Lecturer == DiffColl[j].Lecturer &
                            OutColl[i].Subject == DiffColl[j].Subject &
                            OutColl[i].DocCode == DiffColl[j].DocCode)
                        {
                            //По почасовой должно быть строгое соответствие по
                            //перекрёстным ссылкам во избежание ошибок
                            if (OutColl[i].HouredConnect.Equals(DiffColl[j]))
                            {
                                OutColl[i].Lecture -= DiffColl[j].Lecture;
                                OutColl[i].Exam -= DiffColl[j].Exam;
                                OutColl[i].Credit -= DiffColl[j].Credit;
                                OutColl[i].RefHomeWork -= DiffColl[j].RefHomeWork;
                                OutColl[i].Tutorial -= DiffColl[j].Tutorial;
                                OutColl[i].LabWork -= DiffColl[j].LabWork;
                                OutColl[i].Practice -= DiffColl[j].Practice;
                                OutColl[i].IndividualWork -= DiffColl[j].IndividualWork;
                                OutColl[i].KRAPK -= DiffColl[j].KRAPK;
                                OutColl[i].KursProject -= DiffColl[j].KursProject;

                                //Для нагрузки, которая не распределяется равномерно
                                if (!OutColl[i].flgDistrib)
                                {
                                    OutColl[i].DiplomaPaper -= DiffColl[j].DiplomaPaper;
                                    OutColl[i].PreDiplomaPractice -= DiffColl[j].PreDiplomaPractice;
                                    OutColl[i].TutorialPractice -= DiffColl[j].TutorialPractice;
                                    OutColl[i].ProducingPractice -= DiffColl[j].ProducingPractice;
                                    OutColl[i].GAK -= DiffColl[j].GAK;
                                    OutColl[i].Magistry -= DiffColl[j].Magistry;
                                    OutColl[i].PostGrad -= DiffColl[j].PostGrad;
                                }

                                OutColl[i].Visiting -= DiffColl[j].Visiting;
                            }
                        }
                    }
                }
            }
        }

        //---------------------------------------------------------------------
        //Глобальные функции и процедуры
        //---------------------------------------------------------------------

        /// <summary>
        /// Метод конвертации Фамилии, Имени и Отчества под определённый формат
        /// </summary>
        /// <param name="Text">Входной текст с Фамилией, Именем и Отчеством</param>
        /// <param name="InitialBack">Признак инициалов слева от фамилии (false) или справа от фамилии (true)</param>
        /// <param name="Spaces">Признак наличия пробела между инициалами и фамилией (true). Нет пробела (false)</param>
        /// <returns></returns>
        public static string SplitFIOString(string Text, bool InitialBack, bool Spaces)
        {
            string Surname;
            string Name;
            string Patronymic;
            string[] FIO;
            string str = "";
            int tmpInt;

            //Разбираем строку для вывода отдельно
            //Фамилии, имени и отчества основного преподавателя
            FIO = Text.Split(new char[] { ' ' });

            //Если есть фамилия, имя и отчество
            if (FIO.GetLength(0) == 3)
            {
                Surname = FIO[0];
                Name = FIO[1].Substring(0, 1) + ".";
                Patronymic = FIO[2].Substring(0, 1) + ".";
            }
            //Если есть только фамилия и имя
            else if (FIO.GetLength(0) == 2)
            {
                Surname = FIO[0];
                if (int.TryParse(FIO[1], out tmpInt))
                {
                    Name = FIO[1];
                }
                else
                {
                    Name = FIO[1].Substring(0, 1) + ".";
                }
                Patronymic = "";
            }
            //Если есть только фамилия
            else if (FIO.GetLength(0) == 1)
            {
                Surname = FIO[0];
                Name = "";
                Patronymic = "";
            }
            else
            {
                Surname = "";
                Name = "";
                Patronymic = "";
            }

            if (InitialBack)
            {
                if (Spaces)
                {
                    str = Surname + " " + Name + " " + Patronymic;
                }
                else
                {
                    str = Surname + " " + Name + Patronymic;
                }
            }
            else
            {
                if (Spaces)
                {
                    str = Name + " " + Patronymic + " " + Surname;
                }
                else
                {
                    str = Name + Patronymic + " " + Surname;
                }
            }

            return str;
        }

        /// <summary>
        /// Метод вывода наименования месяца в родительном падеже по коду
        /// </summary>
        /// <param name="Mth">Входящий код месяца</param>
        /// <returns></returns>
        public static string getMonthStringRP(int Mth)
        {
            string str = "";

            switch (Mth)
            {
                case 1:
                    str = "января";
                    break;
                case 2:
                    str = "февраля";
                    break;
                case 3:
                    str = "марта";
                    break;
                case 4:
                    str = "апреля";
                    break;
                case 5:
                    str = "мая";
                    break;
                case 6:
                    str = "июня";
                    break;
                case 7:
                    str = "июля";
                    break;
                case 8:
                    str = "августа";
                    break;
                case 9:
                    str = "сентября";
                    break;
                case 10:
                    str = "октября";
                    break;
                case 11:
                    str = "ноября";
                    break;
                case 12:
                    str = "декабря";
                    break;
            }

            return str;
        }

        public static string getDoWString(DayOfWeek DoW)
        {
            string str = "";

            switch (DoW)
            {
                case DayOfWeek.Monday:
                    str = "в понедельник";
                    break;
                case DayOfWeek.Tuesday:
                    str = "во вторник";
                    break;
                case DayOfWeek.Wednesday:
                    str = "в среду";
                    break;
                case DayOfWeek.Thursday:
                    str = "в четверг";
                    break;
                case DayOfWeek.Friday:
                    str = "в пятницу";
                    break;
                case DayOfWeek.Saturday:
                    str = "в субботу";
                    break;
                case DayOfWeek.Sunday:
                    str = "в воскресенье";
                    break;
            }

            return str;
        }

        public static int CountStringOccurrences(string Text, string Pattern)
        {
            int count = 0;
            int i = 0;

            while ((i = Text.IndexOf(Pattern, i)) != -1)
            {
                i += Pattern.Length;
                count++;
            }

            return count;
        }

        public static string ExcelCellTranslator(int i, int j)
        {
            string cell = "";
            int x;
            int lose;

            x = j;

            if (x < 16384)
            {
                lose = (x - 1) / 676;

                if (lose > 0)
                {
                    cell += Alphabet(lose);
                    x = x - (676 * lose);
                }

                lose = (x - 1) / 26;

                if (lose > 0)
                {
                    cell += Alphabet(lose);
                    x = x - (26 * lose);
                }

                cell += Alphabet(x);
            }

            else
            {
                cell += "XFD";
            }

            cell += i.ToString();

            return cell;
        }

        /// <summary>
        /// Поиск буквы латинского алфавита по введённому номеру
        /// </summary>
        /// <param name="Num">Желаемый номер буквы латинского алфавита</param>
        /// <returns>Возвращает букву</returns>
        public static string Alphabet(int Num)
        {
            string cell = "";

            switch (Num)
            {
                case 1:
                    cell = "A";
                    break;
                case 2:
                    cell = "B";
                    break;
                case 3:
                    cell = "C";
                    break;
                case 4:
                    cell = "D";
                    break;
                case 5:
                    cell = "E";
                    break;
                case 6:
                    cell = "F";
                    break;
                case 7:
                    cell = "G";
                    break;
                case 8:
                    cell = "H";
                    break;
                case 9:
                    cell = "I";
                    break;
                case 10:
                    cell = "J";
                    break;
                case 11:
                    cell = "K";
                    break;
                case 12:
                    cell = "L";
                    break;
                case 13:
                    cell = "M";
                    break;
                case 14:
                    cell = "N";
                    break;
                case 15:
                    cell = "O";
                    break;
                case 16:
                    cell = "P";
                    break;
                case 17:
                    cell = "Q";
                    break;
                case 18:
                    cell = "R";
                    break;
                case 19:
                    cell = "S";
                    break;
                case 20:
                    cell = "T";
                    break;
                case 21:
                    cell = "U";
                    break;
                case 22:
                    cell = "V";
                    break;
                case 23:
                    cell = "W";
                    break;
                case 24:
                    cell = "X";
                    break;
                case 25:
                    cell = "Y";
                    break;
                case 26:
                    cell = "Z";
                    break;
            }

            return cell;
        }

        public static void WordPageDefault(ref Word._Application ObjWord, ref Word._Document ObjDoc, 
                                    float Left = 3f, float Right = 3f, float Top = 2f, 
                                    float Bottom = 2f)
        {
            //Делаем верхнюю границу страницы величиной 0,75 см
            ObjDoc.PageSetup.TopMargin = ObjWord.Application.CentimetersToPoints(Top);
            //Делаем нижнюю границу страницы величиной 0,75 см
            ObjDoc.PageSetup.BottomMargin = ObjWord.Application.CentimetersToPoints(Bottom);
            //Делаем левую границу страницы величиной 0,75 см
            ObjDoc.PageSetup.LeftMargin = ObjWord.Application.CentimetersToPoints(Left);
            //Делаем левую границу страницы величиной 0,75 см
            ObjDoc.PageSetup.RightMargin = ObjWord.Application.CentimetersToPoints(Right);
        }

        public static bool NonZeroDistributionOR(clsDistribution coll)
        {
            bool flg = coll.Lecture > 0 ||
                       coll.Exam > 0 ||
                       coll.Credit > 0 ||
                       coll.RefHomeWork > 0 ||
                       coll.Tutorial > 0 ||
                       coll.LabWork > 0 ||
                       coll.Practice > 0 ||
                       coll.IndividualWork > 0 ||
                       coll.KRAPK > 0 ||
                       coll.KursProject > 0 ||
                       coll.PreDiplomaPractice > 0 ||
                       coll.DiplomaPaper > 0 ||
                       coll.TutorialPractice > 0 ||
                       coll.ProducingPractice > 0 ||
                       coll.GAK > 0 ||
                       coll.PostGrad > 0 ||
                       coll.Visiting > 0 ||
                       coll.Magistry > 0;

            return flg;
        }

        public static bool NonZeroForDispatchOR(clsDistribution coll)
        {
            bool flg = coll.Lecture > 0 ||
                       coll.Practice > 0 ||
                       coll.LabWork > 0 ||
                       coll.KursProject > 0 ||
                       coll.TutorialPractice > 0 ||
                       coll.PreDiplomaPractice > 0 ||
                       coll.ProducingPractice > 0;

            return flg;
        }

        public static int toSumDistributionComponents(clsDistribution coll)
        {
            int sum;

                //1
                sum = coll.Lecture +
                    //2
                          coll.Exam +
                    //3
                          coll.Credit +
                    //4
                          coll.RefHomeWork +
                    //5
                          coll.Tutorial +
                    //6
                          coll.LabWork +
                    //7
                          coll.Practice +
                    //8
                          coll.IndividualWork +
                    //9
                          coll.KRAPK +
                    //10
                          coll.KursProject +
                    //11
                          coll.PreDiplomaPractice +
                    //12
                          coll.DiplomaPaper +
                    //13
                          coll.TutorialPractice +
                    //14
                          coll.ProducingPractice +
                    //15
                          coll.GAK +
                    //Аспирантура
                          coll.PostGrad +
                    //16
                    //coll.Hours +
                    //17
                    //coll.EnteredHours +
                    //Посещение занятий
                          coll.Visiting +
                    //Магистратура
                          coll.Magistry;
            return sum;
        }

        //Без магистратуры, дипломного проектирования, учебной практики
        public static int toSumDistributionComponentsWOCombine(clsDistribution coll)
        {
            //1
            int sum = coll.Lecture +
                //2
                      coll.Exam +
                //3
                      coll.Credit +
                //4
                      coll.RefHomeWork +
                //5
                      coll.Tutorial +
                //6
                      coll.LabWork +
                //7
                      coll.Practice +
                //8
                      coll.IndividualWork +
                //9
                      coll.KRAPK +
                //10
                      coll.KursProject +
                //11
                      coll.PreDiplomaPractice +
                //15
                      coll.GAK +
                //Аспирантура
                      coll.PostGrad +
                //16
                //coll.Hours +
                //17
                //coll.EnteredHours +
                //Посещение занятий
                      coll.Visiting;

            return sum;
        }

        public static string DateToSQL(DateTime D)
        {
            string str = "";

            if (D.Month < 10)
            {
                str = "#0" + D.Month + "/";
            }
            else
            {
                str = "#" + D.Month + "/";
            }

            if (D.Day < 10)
            {
                str += "0" + D.Day + "/";
            }
            else
            {
                str += D.Day + "/";
            }

            str += D.Year + "#";

            return str;
        }

        public static string commentLoad(clsDistribution coll)
        {
            string subjInfo = "(";

            subjInfo = "(";

            if (coll.Lecture > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "лк.-" + coll.Lecture.ToString() + "; ";
                }
                else
                {
                    subjInfo += "лк.-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            if (coll.Exam > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "экз.-" + coll.Exam.ToString() + "; ";
                }
                else
                {
                    subjInfo += "экз.-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            if (coll.Credit > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "зач.-" + coll.Credit.ToString() + "; ";
                }
                else
                {
                    subjInfo += "зач.-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            if (coll.RefHomeWork > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "реф.-" + coll.RefHomeWork.ToString() + "; ";
                }
                else
                {
                    subjInfo += "реф.-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            if (coll.Tutorial > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "конс.-" + coll.Tutorial.ToString() + "; ";
                }
                else
                {
                    subjInfo += "конс.-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            if (coll.LabWork > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "лб.-" + coll.LabWork.ToString() + "; ";
                }
                else
                {
                    subjInfo += "лб.-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            if (coll.Practice > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "пр.-" + coll.Practice.ToString() + "; ";
                }
                else
                {
                    subjInfo += "пр.-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            if (coll.IndividualWork > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "инд.-" + coll.IndividualWork.ToString() + "; ";
                }
                else
                {
                    subjInfo += "инд.-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            if (coll.KRAPK > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "ПК.-" + coll.KRAPK.ToString() + "; ";
                }
                else
                {
                    subjInfo += "ПК.-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            if (coll.KursProject > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "к/п-" + coll.KursProject.ToString() + "; ";
                }
                else
                {
                    subjInfo += "к/п-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            if (coll.PreDiplomaPractice > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "ПДП-" + coll.PreDiplomaPractice.ToString() + "; ";
                }
                else
                {
                    subjInfo += "ПДП-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            if (coll.DiplomaPaper > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "дип.-" + coll.DiplomaPaper.ToString() + "; ";
                }
                else
                {
                    subjInfo += "дип.-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            if (coll.TutorialPractice > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "уч.пр.-" + coll.TutorialPractice.ToString() + "; ";
                }
                else
                {
                    subjInfo += "уч.пр.-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            if (coll.ProducingPractice > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "пр.пр.-" + coll.ProducingPractice.ToString() + "; ";
                }
                else
                {
                    subjInfo += "пр.пр.-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            if (coll.GAK > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "ГАК.-" + coll.GAK.ToString() + "; ";
                }
                else
                {
                    subjInfo += "ГАК.-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            if (coll.PostGrad > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "асп.-" + coll.PostGrad.ToString() + "; ";
                }
                else
                {
                    subjInfo += "асп.-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            if (coll.Magistry > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "маг.-" + coll.Magistry.ToString() + "; ";
                }
                else
                {
                    subjInfo += "маг.-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            if (coll.Visiting > 0)
            {
                if (!coll.flgDistrib)
                {
                    subjInfo += "посещ.-" + coll.Visiting.ToString() + "; ";
                }
                else
                {
                    subjInfo += "посещ.-" + coll.Weight.ToString() + "н/чел.; ";
                }
            }

            subjInfo += ")";

            return subjInfo;
        }

        //Составление массива элементов расписания
        public static void toCompleteMassSchedule()
        {
            clsSchedule S;
            bool flgExist = false;
            bool flgCreateNew = false;

            //Перебираем все семестры
            for (int i0 = 0; i0 <= colSemestr.Count - 1; i0++)
            {
                //Если семестр не прочерк
                if (!colSemestr[i0].SemNum.Equals("-"))
                {
                    //Перебираем недели
                    for (int i = 0; i <= colWeek.Count - 1; i++)
                        //Перебираем дни
                        for (int j = 0; j <= colWeekDays.Count - 1; j++)
                            //Перебираем пары учебных занятий
                            for (int k = 0; k <= colPairTime.Count - 1; k++)
                                //Перебираем преподавателей
                                for (int l = 0; l <= colLecturer.Count - 1; l++)
                                {
                                    S = new clsSchedule();

                                    S.Semestr = colSemestr[i0];
                                    S.Lecturer = colLecturer[l];
                                    S.Time = colPairTime[k];
                                    S.Week = colWeek[i];
                                    S.WeekDay = colWeekDays[j];
                                    S.Subj = false;

                                    flgExist = false;
                                    for (int m = 0; m <= colSchedule.Count - 1; m++)
                                    {
                                        if (colSchedule[m].Lecturer.FIO.Equals(S.Lecturer.FIO) &
                                            colSchedule[m].Time.Time.Equals(S.Time.Time) &
                                            colSchedule[m].Week.NumberWeek.Equals(S.Week.NumberWeek) &
                                            colSchedule[m].WeekDay.WeekDay.Equals(S.WeekDay.WeekDay) &
                                            colSchedule[m].Semestr.SemNum.Equals(S.Semestr.SemNum))
                                        {
                                            flgExist = true;
                                            //Для отладки принудительно выставить второй семестр
                                            //colSchedule[m].Semestr = colSemestr[2];
                                        }
                                    }

                                    if (!flgExist)
                                    {
                                        colSchedule.Add(S);
                                        flgCreateNew = true;
                                    }

                                    S = null;
                                }
                }
            }

            if (flgCreateNew)
            {
                for (int m = 0; m <= colSchedule.Count - 1; m++)
                {
                    colSchedule[m].Code = m + 1;
                }
            }
        }

        //Вычищение коллекции от повторных элементов
        public static void toModifyDistributionToDetailed()
        {
            clsDistributionDetailed DDmain;
            clsDistributionDetailed DDdop;

            //Перебираем тех, кого сравниваем
            for (int i = 0; i < colDistributionDetailed.Count; i++)
            {
                DDmain = colDistributionDetailed[i];
                //Перебираем тех, с кем сравниваем
                for (int j = 0; j < colDistributionDetailed.Count; j++)
                {
                    DDdop = colDistributionDetailed[j];
                    //Исключение ситуации сравнения с самим собой
                    if (i != j)
                    {
                        //Совпасть должны преподаватель, дисциплина, семестр,
                        //специальность, курс, вид занятия, количество часов
                        if (DDmain.Lecturer.FIO.Equals(DDdop.Lecturer.FIO) &
                             DDmain.Subject.Subject.Equals(DDdop.Subject.Subject) &
                             DDmain.Semestr.SemNum.Equals(DDdop.Semestr.SemNum) &
                             DDmain.Speciality.ShortInstitute.Equals(DDdop.Speciality.ShortInstitute) &
                             DDmain.KursNum.Kurs.Equals(DDdop.KursNum.Kurs) &
                             DDmain.SubjType.Short.Equals(DDdop.SubjType.Short) &
                             DDmain.SubjHours.Equals(DDdop.SubjHours) )
                        {

                        }
                    }
                }
            }
        }

        public static void toConvertDistributionToDetailed(ref float sum)
        {
            int i, j, count;
            clsDistributionDetailed DD;
            clsDistribution currD;
            clsSubjectType currST;

            colDistributionDetailed.Clear();
            count = 0;

            for (i = 0; i < colDistribution.Count; i++)
            {
                if (!colDistribution[i].flgExclude)
                {
                    currD = colDistribution[i];
                    for (j = 0; j < colSubjectType.Count; j++)
                    {
                        currST = colSubjectType[j];
                        DD = new clsDistributionDetailed();
                        //формируем код по счётчику
                        DD.Code = count;
                        //переписываем принадлежность семестру
                        DD.Semestr = currD.Semestr;
                        //переписываем принадлежность специальности
                        DD.Speciality = currD.Speciality;
                        //переписываем принадлежность курсу
                        DD.KursNum = currD.KursNum;
                        //переписываем принадлежность дисциплине
                        DD.Subject = currD.Subject;
                        //переписываем связь дисциплины с преподавателем
                        DD.Lecturer = currD.Lecturer;
                        //переписываем связь дисциплины с заменяющим преподавателем
                        DD.Lecturer2 = currD.Lecturer2;
                        //переписываем связь дисциплины с резервным преподавателем
                        DD.Lecturer3 = currD.Lecturer3;
                        //переписываем связь дисциплины с преподавателем-дублёром
                        DD.Doubler = currD.Doubler;
                        //формируем вид учебной нагрузки
                        DD.SubjType = currST;
                        //
                        switch (currST.Short)
                        {
                            //Часы на лекции
                            case "лк.":
                                {
                                    DD.SubjHours = currD.Lecture;
                                    break;
                                }
                            //Часы на экзамены
                            case "экз.":
                                {
                                    DD.SubjHours = currD.Exam;
                                    break;
                                }
                            //Часы на зачёты
                            case "зач.":
                                {
                                    DD.SubjHours = currD.Credit;
                                    break;
                                }
                            //Часы на рефераты и домашние задания
                            case "реф.":
                                {
                                    DD.SubjHours = currD.RefHomeWork;
                                    break;
                                }
                            //Часы на консультации
                            case "конс.":
                                {
                                    DD.SubjHours = currD.Tutorial;
                                    break;
                                }
                            //Часы на лабораторные работы
                            case "лр.":
                                {
                                    DD.SubjHours = currD.LabWork;
                                    break;
                                }
                            //Часы на практические занятия
                            case "пр.":
                                {
                                    DD.SubjHours = currD.Practice;
                                    break;
                                }
                            //Часы на индивидуальные задания
                            case "инд.":
                                {
                                    DD.SubjHours = currD.IndividualWork;
                                    break;
                                }
                            //Часы на контрольные работы и промежуточный контроль
                            case "КРАПК":
                                {
                                    DD.SubjHours = currD.KRAPK;
                                    break;
                                }
                            //Часы на курсовой проект (курсовую работу)
                            case "к.п.(к.р.)":
                                {
                                    DD.SubjHours = currD.KursProject;
                                    break;
                                }
                            //Часы на преддипломную практику
                            case "предд.пр.":
                                {
                                    DD.SubjHours = currD.PreDiplomaPractice;
                                    break;
                                }
                            //Часы на дипломное проектирование
                            case "дипл.пр.":
                                {
                                    DD.SubjHours = currD.DiplomaPaper;
                                    break;
                                }
                            //Часы на учебную практику
                            case "уч.пр.":
                                {
                                    DD.SubjHours = currD.TutorialPractice;
                                    break;
                                }
                            //Часы на производственную практику
                            case "пр.пр.":
                                {
                                    DD.SubjHours = currD.ProducingPractice;
                                    break;
                                }
                            //Часы на аспирантуру
                            case "асп.":
                                {
                                    DD.SubjHours = currD.PostGrad;
                                    break;
                                }
                            //Часы на ГАК
                            case "ГАК":
                                {
                                    DD.SubjHours = currD.GAK;
                                    break;
                                }
                            //Часы бюджетные
                            case "бюдж.":
                                {
                                    DD.SubjHours = currD.Hours;
                                    break;
                                }
                            //Часы бюджетные в ЗЕТ
                            case "бюдж.ЗЕТ":
                                {
                                    DD.SubjHours = currD.HoursZ;
                                    break;
                                }
                            //Контрольные значения
                            case "контроль":
                                {
                                    DD.SubjHours = currD.EnteredHours;
                                    sum += currD.EnteredHours;
                                    break;
                                }
                            //Контрольные значения в ЗЕТ
                            case "контрольЗЕТ":
                                {
                                    DD.SubjHours = currD.EnteredHoursZ;
                                    break;
                                }
                            //Часы на посещение занятий
                            case "посещ.":
                                {
                                    DD.SubjHours = currD.Visiting;
                                    break;
                                }
                            //Часы на магистратуру
                            case "магистр.":
                                {
                                    DD.SubjHours = currD.Magistry;
                                    break;
                                }
                        }
                        //переписываем текст комментария к дисциплине
                        DD.Text = currD.Text;
                        //переписываем признак необходимости учёта в заявке для диспетчерской
                        DD.flgDispatch = currD.flgDispatch;
                        //переписываем связь по лабораторным работам
                        DD.LabWorkConnect = currD.LabWorkConnect;
                        //переписываем признак равномерного распределения нагрузки
                        DD.flgDistrib = currD.flgDistrib;
                        //переписываем значение веса на один элемент, связанный с нагрузкой
                        DD.Weight = currD.Weight;
                        //переписываем признак исключения из учебной нагрузки
                        DD.flgExclude = currD.flgExclude;
                        //добавляем элемент в коллекцию
                        colDistributionDetailed.Add(DD);
                        //увеличиваем значение счётчика
                        count++;
                    }
                }
            }
        }

        //Метод, который исключает из суммарного значения равномерно распределяемой нагрузку
        //именно ту её долю, которая оплачивается через почасовой фонд
        //(тернарный оператор)
        public static void toDetectUniformInHoured(ref int sum, clsDistribution dstrb, clsLecturer lect)
        {
            bool flgNeedDecrease;
            clsDistribution D;

            for (int k = 0; k < mdlData.colHouredDistribution.Count; k++)
            {
                D = mdlData.colHouredDistribution[k];
                flgNeedDecrease = true;

                //Проверяем соответствие преподавателя
                flgNeedDecrease &= (D.Lecturer != null) ?
                    (D.Lecturer.Equals(lect)) :
                    (false);
                //Проверяем соответствие номера курса
                flgNeedDecrease &= (D.KursNum != null & dstrb.KursNum != null) ?
                    D.KursNum.Equals(dstrb.KursNum) :
                    false;
                //Проверяем соответствие специальности
                flgNeedDecrease &= (D.Speciality != null & dstrb.Speciality != null) ?
                    (D.Speciality.Equals(dstrb.Speciality)) :
                    (false);
                //Проверяем соответствие дисциплины
                flgNeedDecrease &= (D.Subject != null & dstrb.Subject != null) ?
                    (D.Subject.Equals(dstrb.Subject)) :
                    (false);
                //Проверяем соответствие семестра
                flgNeedDecrease &= (D.Semestr != null & dstrb.Semestr != null) ?
                    (D.Semestr.Equals(dstrb.Semestr)) :
                    (false);

                //Если преподаватель совпал и совпала равномерно 
                //распределяемая нагрузка
                if (flgNeedDecrease)
                    sum -= mdlData.toSumDistributionComponents(D);
            }
        }

        public static void toCompleteDopWork()
        {
            clsDopWork DW;
            bool flgExist = false;
            int count = 0;

            //Последовательно перебираем всех преподавателей
            for (int i = 0; i <= colLecturer.Count - 1; i++)
                //Последовательно перебираем семестры
                for (int j = 0; j <= colSemestr.Count - 1; j++)
                {
                    if (!colSemestr[j].SemNum.Equals("-"))
                    {
                        //Создаём новую строку дополнительной работы для
                        //рассматриваемого преподавателя в рассматриваемом семестре
                        DW = new clsDopWork();

                        DW.Code = colDopWork.Count + 1;

                        DW.Lecturer = colLecturer[i];
                        DW.Semestr = colSemestr[j];
                        DW.NIR = "";
                        DW.OMR = "";
                        DW.UMR = "";

                        flgExist = false;
                        //Перебираем дополнительную работу
                        for (int k = 0; k <= colDopWork.Count - 1; k++)
                        {
                            if (colDopWork[k].Lecturer.Equals(DW.Lecturer) &
                                colDopWork[k].Semestr.Equals(DW.Semestr))
                            {
                                //Если нашлось совпадение
                                flgExist = true;
                            }
                        }

                        //Если совпадения не нашлось
                        if (!flgExist)
                        {
                            //Добавляем строчку дополнительной работы
                            colDopWork.Add(DW);
                            count++;
                        }

                        DW = null;
                    }
                }
        }

        //Метод вывода дочерней формы
        public static void toGenerateForm(Form fOw, Form f)
        {
            if (mdlData.flgReady)
            {
                //Делаем наследование от главной формы
                f.Owner = fOw;
                //Отображаем форму на экране
                f.ShowDialog();
                //Очищаем память от формы
                f = null;
            }
            else
            {
                MessageBox.Show(fOw, "Пожалуйста, загрузите сначала базу данных", "В доступе к функции отказано!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}