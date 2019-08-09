using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;

namespace DataBaseIO
{
    class clsDistribution
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Семестр +
        /// </summary>
        public clsSemestr Semestr;
        /// <summary>
        /// Специальность +
        /// </summary>
        public clsSpecialisation Speciality;
        /// <summary>
        /// Курс +
        /// </summary>
        public clsKursNum KursNum;
        /// <summary>
        /// Название учебной дисциплины +
        /// </summary>
        public clsSubject Subject;
        /// <summary>
        /// Преподаватель +
        /// </summary>
        public clsLecturer Lecturer;
        /// <summary>
        /// Замещающий преподаватель +
        /// </summary>
        public clsLecturer Lecturer2;
        /// <summary>
        /// Зам. замещающего преподавателя +
        /// </summary>
        public clsLecturer Lecturer3;
        /// <summary>
        /// Преподаватель-дублёр
        /// </summary>
        public clsLecturer Doubler;
        /// <summary>
        /// Лекции в часах +
        /// </summary>
        public int Lecture;
        /// <summary>
        /// Экзамен в часах +
        /// </summary>
        public int Exam;
        /// <summary>
        /// Зачёт в часах +
        /// </summary>
        public int Credit;
        /// <summary>
        /// Рефераты, домашние задания в часах +
        /// </summary>
        public int RefHomeWork;
        /// <summary>
        /// Консультация в часах +
        /// </summary>
        public int Tutorial;
        /// <summary>
        /// Лабораторная работа в часах +
        /// </summary>
        public int LabWork;
        /// <summary>
        /// Практическое занятие в часах +
        /// </summary>
        public int Practice;
        /// <summary>
        /// Индивидуальное занятие в часах +
        /// </summary>
        public int IndividualWork;
        /// <summary>
        /// КРАПК в часах +
        /// </summary>
        public int KRAPK;
        /// <summary>
        /// Курсовой проект в часах +
        /// </summary>
        public int KursProject;
        /// <summary>
        /// Преддипломная практика в часах +
        /// </summary>
        public int PreDiplomaPractice;
        /// <summary>
        /// Дипломный проект в часах +
        /// </summary>
        public int DiplomaPaper;
        /// <summary>
        /// Учебная практика в часах +
        /// </summary>
        public int TutorialPractice;
        /// <summary>
        /// Производственная практика в часах +
        /// </summary>
        public int ProducingPractice;
        /// <summary>
        /// Аспирантура в часах
        /// </summary>
        public int PostGrad;
        /// <summary>
        /// ГАК в часах
        /// </summary>
        public int GAK;
        /// <summary>
        /// Часы
        /// </summary>
        public int Hours;
        /// <summary>
        /// Часы ЗЕТ
        /// </summary>
        public float HoursZ;
        /// <summary>
        /// Часы Заданные
        /// </summary>
        public int EnteredHours;
        /// <summary>
        /// Часы Заданные по ЗЕТ
        /// </summary>
        public float EnteredHoursZ;
        /// <summary>
        /// Часы на посещение занятий
        /// </summary>
        public int Visiting;
        /// <summary>
        /// Примечание
        /// </summary>
        public string Text;
        /// <summary>
        /// Часы на руководство магистерской программой
        /// </summary>
        public int Magistry;
        /// <summary>
        /// Признак необходимости занесения в заявку
        /// для диспетчерской
        /// </summary>
        public bool flgDispatch;
        /// <summary>
        /// Связка по лабораторным работам
        /// </summary>
        public clsDistribution LabWorkConnect;
        /// <summary>
        /// Признак перераспределения нагрузки
        /// </summary>
        public bool flgDistrib;
        /// <summary>
        /// Вес единицы нагрузки
        /// </summary>
        public int Weight;
        /// <summary>
        /// Признак исключения из расчёта
        /// </summary>
        public bool flgExclude;
        /// <summary>
        /// Код по документу с нагрузкой (признак упорядочивания)
        /// </summary>
        public int DocCode;
        /// <summary>
        /// Связка по почасовой нагрузке
        /// </summary>
        public clsDistribution HouredConnect;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Распределение нагрузки"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            int GetCode;
            string CurrentString;
            bool Detected;
            bool Even;

            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["Код"].ToString());
            
            //Семестр
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colSemestr.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Семестр"].ToString();
                if (mdlData.colSemestr[i].SemNum == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Semestr = mdlData.colSemestr[GetCode];
            }
            else
            {
                this.Semestr = null;
                //MessageBox.Show("Не удалось определить семестр у элемента с кодом " + this.Code, "Оповещение");
            }
            
            //Специальность
            GetCode = 0;
            Detected = false;
            Even = false;

            for (int i = 0; i <= mdlData.colSpecialisation.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Специальность"].ToString();

                if (!mdlData.flgOldDB)
                {
                    if (CurrentString.StartsWith("В-"))
                    {
                        CurrentString = CurrentString.Substring(2);
                        Even = true;
                    }

                    if (mdlData.colSpecialisation[i].ShortDop == CurrentString)
                    {
                        //Если с вечернего факультета
                        if (Even)
                        {
                            if (mdlData.colSpecialisation[i].Faculty != null)
                            {
                                if (mdlData.colSpecialisation[i].Faculty.Short.Equals("ВФ"))
                                {
                                    GetCode = i;
                                    Detected = true;
                                    break;
                                }
                            }
                        }
                        //Если с дневного
                        else
                        {
                            GetCode = i;
                            Detected = true;
                            break;
                        }
                    }
                }
                else
                {
                    if (mdlData.colSpecialisation[i].ShortUpravlenie == CurrentString)
                    {
                        GetCode = i;
                        Detected = true;
                        break;
                    }
                }
            }
            if (Detected)
            {
                this.Speciality = mdlData.colSpecialisation[GetCode];
            }
            else
            {
                this.Speciality = null;
            }

            //Курс
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colKursNum.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Курс"].ToString();
                if (mdlData.colKursNum[i].Kurs.ToString() == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.KursNum = mdlData.colKursNum[GetCode];
            }
            else
            {
                this.KursNum = null;
            }

            //Название учебной дисциплины
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colSubject.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Название_дисциплины"].ToString();
                if (mdlData.colSubject[i].Subject == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Subject = mdlData.colSubject[GetCode];
            }
            else
            {
                this.Subject = null;
            }

            //Преподаватель
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Преподаватель"].ToString();
                if (mdlData.colLecturer[i].FIO == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Lecturer = mdlData.colLecturer[GetCode];
            }
            else
            {
                this.Lecturer = null;
            }
           
            //Преподаватель2
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Преподаватель2"].ToString();
                if (mdlData.colLecturer[i].FIO == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Lecturer2 = mdlData.colLecturer[GetCode];
            }
            else
            {
                this.Lecturer2 = null;
            }

            //Преподаватель3
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Преподаватель3"].ToString();
                if (mdlData.colLecturer[i].FIO == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Lecturer3 = mdlData.colLecturer[GetCode];
            }
            else
            {
                this.Lecturer3 = null;
            }

            //Дублёр
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Дублёр"].ToString();
                if (mdlData.colLecturer[i].FIO == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Doubler = mdlData.colLecturer[GetCode];
            }
            else
            {
                this.Doubler = null;
            }

            //Лекции в часах
            if (!(Tab.Rows[id]["Лекция"] is DBNull))
            {
                this.Lecture = Convert.ToInt32(Tab.Rows[id]["Лекция"].ToString());
            }
            else
            {
                this.Lecture = 0;
            }

            //Экзамен в часах
            if (!(Tab.Rows[id]["Экзамен"] is DBNull))
            {
                this.Exam = Convert.ToInt32(Tab.Rows[id]["Экзамен"].ToString());
            }
            else
            {
                this.Exam = 0;
            }

            //Зачёт в часах
            if (!(Tab.Rows[id]["Зачёт"] is DBNull))
            {
                this.Credit = Convert.ToInt32(Tab.Rows[id]["Зачёт"].ToString());
            }
            else
            {
                this.Credit = 0;
            }

            //Рефераты и домашние задания в часах
            if (!(Tab.Rows[id]["Рефераты, домашние задания"] is DBNull))
            {
                this.RefHomeWork = Convert.ToInt32(Tab.Rows[id]["Рефераты, домашние задания"].ToString());
            }
            else
            {
                this.RefHomeWork = 0;
            }

            //Консультация в часах
            if (!(Tab.Rows[id]["Консультация"] is DBNull))
            {
                this.Tutorial = Convert.ToInt32(Tab.Rows[id]["Консультация"].ToString());
            }
            else
            {
                this.Tutorial = 0;
            }

            //Лабораторные работы в часах
            if (!(Tab.Rows[id]["Лаб_раб"] is DBNull))
            {
                this.LabWork = Convert.ToInt32(Tab.Rows[id]["Лаб_раб"].ToString());
            }
            else
            {
                this.LabWork = 0;
            }

            //Практические занятия в часах
            if (!(Tab.Rows[id]["Практ_зан"] is DBNull))
            {
                this.Practice = Convert.ToInt32(Tab.Rows[id]["Практ_зан"].ToString());
            }
            else
            {
                this.Practice = 0;
            }

            //Индивидуальное занятие в часах
            if (!(Tab.Rows[id]["Индив_зан"] is DBNull))
            {
                this.IndividualWork = Convert.ToInt32(Tab.Rows[id]["Индив_зан"].ToString());
            }
            else
            {
                this.IndividualWork = 0;
            }

            //КРАПК в часах
            if (!(Tab.Rows[id]["КРАПК"] is DBNull))
            {
                this.KRAPK = Convert.ToInt32(Tab.Rows[id]["КРАПК"].ToString());
            }
            else
            {
                this.KRAPK = 0;
            }

            //Курсовой проект в часах
            if (!(Tab.Rows[id]["Курс_пр"] is DBNull))
            {
                this.KursProject = Convert.ToInt32(Tab.Rows[id]["Курс_пр"].ToString());
            }
            else
            {
                this.KursProject = 0;
            }

            //Преддипломная практика в часах
            if (!(Tab.Rows[id]["Предд_пр"] is DBNull))
            {
                this.PreDiplomaPractice = Convert.ToInt32(Tab.Rows[id]["Предд_пр"].ToString());
            }
            else
            {
                this.PreDiplomaPractice = 0;
            }

            //Дипломный проект в часах
            if (!(Tab.Rows[id]["Дипл_пр"] is DBNull))
            {
                this.DiplomaPaper = Convert.ToInt32(Tab.Rows[id]["Дипл_пр"].ToString());
            }
            else
            {
                this.DiplomaPaper = 0;
            }

            //Учебная практика в часах
            if (!(Tab.Rows[id]["Учебн_пр"] is DBNull))
            {
                this.TutorialPractice = Convert.ToInt32(Tab.Rows[id]["Учебн_пр"].ToString());
            }
            else
            {
                this.TutorialPractice = 0;
            }

            //Производственная практика в часах
            if (!(Tab.Rows[id]["Произв_пр"] is DBNull))
            {
                this.ProducingPractice = Convert.ToInt32(Tab.Rows[id]["Произв_пр"].ToString());
            }
            else
            {
                this.ProducingPractice = 0;
            }

            //Аспирантура в часах
            if (!(Tab.Rows[id]["Аспирантура"] is DBNull))
            {
                this.PostGrad = Convert.ToInt32(Tab.Rows[id]["Аспирантура"].ToString());
            }
            else
            {
                this.PostGrad = 0;
            }

            //ГАК в часах
            if (!(Tab.Rows[id]["ГАК"] is DBNull))
            {
                this.GAK = Convert.ToInt32(Tab.Rows[id]["ГАК"].ToString());
            }
            else
            {
                this.GAK = 0;
            }

            //Часы
            if (!(Tab.Rows[id]["Часы"] is DBNull))
            {
                this.Hours = Convert.ToInt32(Tab.Rows[id]["Часы"].ToString());
            }
            else
            {
                this.Hours = 0;
            }

            //Часы по ЗЕТ
            if (!(Tab.Rows[id]["Часы_ЗЕТ"] is DBNull))
            {
                this.HoursZ = Convert.ToSingle(Tab.Rows[id]["Часы_ЗЕТ"].ToString());
            }
            else
            {
                this.HoursZ = 0;
            }

            //Часы дано
            if (!(Tab.Rows[id]["Часы_Дано"] is DBNull))
            {
                this.EnteredHours = Convert.ToInt32(Tab.Rows[id]["Часы_Дано"].ToString());
            }
            else
            {
                this.EnteredHours = 0;
            }

            //Часы дано по ЗЕТ
            if (!(Tab.Rows[id]["Часы_Дано_ЗЕТ"] is DBNull))
            {
                this.EnteredHoursZ = Convert.ToSingle(Tab.Rows[id]["Часы_Дано_ЗЕТ"].ToString());
            }
            else
            {
                this.EnteredHoursZ = 0;
            }

            //Часы на посещение занятий
            if (!(Tab.Rows[id]["Посещ_зан"] is DBNull))
            {
                this.Visiting = Convert.ToInt32(Tab.Rows[id]["Посещ_зан"].ToString());
            }
            else
            {
                this.Visiting = 0;
            }

            this.Text = Tab.Rows[id]["Для_диспетчерской"].ToString();

            //Часы на руководство магистрами
            if (!(Tab.Rows[id]["Магистратура"] is DBNull))
            {
                this.Magistry = Convert.ToInt32(Tab.Rows[id]["Магистратура"].ToString());
            }
            else
            {
                this.Magistry = 0;
            }

            //Необходимость подачи нагрузки в диспетчерскую
            if (Tab.Rows[id]["В_ведом_дисп"].ToString() == "False")
            {
                this.flgDispatch = false;
            }
            else
            {
                this.flgDispatch = true;
            }

            //Равномерная распределяемость
            if (Tab.Rows[id]["Распределяемая"].ToString() == "False")
            {
                this.flgDistrib = false;
            }
            else
            {
                this.flgDistrib = true;
            }

            //Вес единицы
            if (!(Tab.Rows[id]["Вес_единицы"] is DBNull))
            {
                this.Weight = Convert.ToInt32(Tab.Rows[id]["Вес_единицы"].ToString());
            }
            else
            {
                this.Weight = 0;
            }

            //Признак исключения из расчёта нагрузки
            if (Tab.Rows[id]["Исключить"].ToString() == "False")
            {
                this.flgExclude = false;
            }
            else
            {
                this.flgExclude = true;
            }

            //Код дисциплины по документу
            if (!(Tab.Rows[id]["КодДок"] is DBNull))
            {
                this.DocCode = Convert.ToInt32(Tab.Rows[id]["КодДок"].ToString());
            }
            else
            {
                this.DocCode = 9999;
            }
        }

        public void SelfLinking(DataTable Tab, int id)
        {
            int GetCode;
            string CurrentString;
            bool Detected;

            //Связь по лабораторным работам
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Связь_лб"].ToString();
                if (mdlData.colDistribution[i].Code.ToString() == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.LabWorkConnect = mdlData.colDistribution[GetCode];
            }
            else
            {
                this.LabWorkConnect = null;
            }
        }

        public void CrossLinking(DataTable Tab, int id, bool flgHoured)
        {
            int GetCode;
            string CurrentString;
            bool Detected;

            //Применительно к учебной нагрузке - искать соответствие в почасовой
            if (!flgHoured)
            {
                //Связь по почасовой нагрузке
                GetCode = 0;
                Detected = false;
                for (int i = 0; i <= mdlData.colHouredDistribution.Count - 1; i++)
                {
                    CurrentString = Tab.Rows[id]["Связь_почас"].ToString();
                    if (mdlData.colHouredDistribution[i].Code.ToString() == CurrentString)
                    {
                        GetCode = i;
                        Detected = true;
                    }
                }
                if (Detected)
                {
                    this.HouredConnect = mdlData.colHouredDistribution[GetCode];
                }
                else
                {
                    this.HouredConnect = null;
                }
            }
            //Применительно к почасовой нагрузке - искать соответствие в учебной
            else
            {
                //Связь по почасовой нагрузке
                GetCode = 0;
                Detected = false;
                for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
                {
                    CurrentString = Tab.Rows[id]["Связь_почас"].ToString();
                    if (mdlData.colDistribution[i].Code.ToString() == CurrentString)
                    {
                        GetCode = i;
                        Detected = true;
                    }
                }
                if (Detected)
                {
                    this.HouredConnect = mdlData.colDistribution[GetCode];
                }
                else
                {
                    this.HouredConnect = null;
                }
            }
        }

        public void ClearHours()
        {
            Lecture = 0;
            Exam = 0;
            Credit= 0;
            RefHomeWork = 0;
            Tutorial = 0;
            LabWork = 0;
            Practice = 0;
            IndividualWork = 0;
            KRAPK = 0;
            KursProject = 0;
            PreDiplomaPractice = 0;
            DiplomaPaper = 0;
            TutorialPractice = 0;
            ProducingPractice = 0;
            PostGrad = 0;
            GAK = 0;
            Hours = 0;
            HoursZ = 0;
            EnteredHours = 0;
            EnteredHoursZ = 0;
            Visiting = 0;
            Magistry = 0;
        }

        /// <summary>
        /// Метод копирования в рассматриваемый элемент информации от внешнего элемента
        /// </summary>
        /// <param name="Input">Внешний элемент</param>
        /// <param name="flgHours">Флаг сохранения часов нагрузки true - сохранить, false - стереть</param>
        public void CopyFrom(clsDistribution Input, bool flgHours)
        {
            Code = Input.Code;
            Semestr = Input.Semestr;
            Speciality = Input.Speciality;
            KursNum = Input.KursNum;
            Subject = Input.Subject;
            Lecturer = Input.Lecturer;
            Lecturer2 = Input.Lecturer2;
            Lecturer3 = Input.Lecturer3;
            Doubler = Input.Doubler;

            if (flgHours)
            {
                Lecture = Input.Lecture;
                Exam = Input.Exam;
                Credit = Input.Credit;
                RefHomeWork = Input.RefHomeWork;
                Tutorial = Input.Tutorial;
                LabWork = Input.LabWork;
                Practice = Input.Practice;
                IndividualWork = Input.IndividualWork;
                KRAPK = Input.KRAPK;
                KursProject = Input.KursProject;
                PreDiplomaPractice = Input.PreDiplomaPractice;
                DiplomaPaper = Input.DiplomaPaper;
                TutorialPractice = Input.TutorialPractice;
                ProducingPractice = Input.ProducingPractice;
                PostGrad = Input.PostGrad;
                GAK = Input.GAK;
                Hours = Input.Hours;
                HoursZ = Input.HoursZ;
                EnteredHours = Input.EnteredHours;
                EnteredHoursZ = Input.EnteredHoursZ;
                Visiting = Input.Visiting;
                Magistry = Input.Magistry;
            }
            else
            {
                Lecture = 0;
                Exam = 0;
                Credit = 0;
                RefHomeWork = 0;
                Tutorial = 0;
                LabWork = 0;
                Practice = 0;
                IndividualWork = 0;
                KRAPK = 0;
                KursProject = 0;
                PreDiplomaPractice = 0;
                DiplomaPaper = 0;
                TutorialPractice = 0;
                ProducingPractice = 0;
                PostGrad = 0;
                GAK = 0;
                Hours = 0;
                HoursZ = 0;
                EnteredHours = 0;
                EnteredHoursZ = 0;
                Visiting = 0;
                Magistry = 0;
            }

            Text = Input.Text;
            flgDispatch = Input.flgDispatch;
            LabWorkConnect = Input.LabWorkConnect;
            flgDistrib = Input.flgDistrib;
            Weight = Input.Weight;
            flgExclude = Input.flgExclude;
            DocCode = Input.DocCode;
            HouredConnect = Input.HouredConnect;
        }

        public string Save(int id)
        {
            string str = "";

            //Добавляем Код
            str += this.Code.ToString() + ", ";
            
            //Добавляем Семестр
            if (this.Semestr != null)
            {
                str += "'" + this.Semestr.SemNum + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Специальность
            if (this.Speciality != null)
            {
                str += "'" + this.Speciality.ShortDop + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }
           
            //Добавляем Курс
            if (this.KursNum != null)
            {
                str += "'" + this.KursNum.Kurs + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Название дисциплины
            if (this.Subject != null)
            {
                str += "'" + this.Subject.Subject.ToString() + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Преподаватель
            if (this.Lecturer != null)
            {
                str += "'" + this.Lecturer.FIO + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Преподаватель2
            if (this.Lecturer2 != null)
            {
                str += "'" + this.Lecturer2.FIO + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Преподаватель3
            if (this.Lecturer3 != null)
            {
                str += "'" + this.Lecturer3.FIO + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Лекции
            str += "'" + this.Lecture + "', ";
            //Добавляем Экзамен
            str += "'" + this.Exam + "', ";
            //Добавляем Зачёт
            str += "'" + this.Credit + "', ";
            //Добавляем Рефераты
            str += "'" + this.RefHomeWork + "', ";
            //Добавляем Консультация
            str += "'" + this.Tutorial + "', ";
            //Добавляем Лабораторные
            str += "'" + this.LabWork + "', ";
            //Добавляем Практика
            str += "'" + this.Practice + "', ";
            //Добавляем Индивидуальные
            str += "'" + this.IndividualWork + "', ";
            //Добавляем КРАПК
            str += "'" + this.KRAPK + "', ";
            //Добавляем Курсовой проект
            str += "'" + this.KursProject + "', ";
            //Добавляем Преддипломная практика
            str += "'" + this.PreDiplomaPractice + "', ";
            //Добавляем Диплом
            str += "'" + this.DiplomaPaper + "', ";
            //Добавляем Учебная практика
            str += "'" + this.TutorialPractice + "', ";
            //Добавляем Производственная практика
            str += "'" + this.ProducingPractice + "', ";
            //Добавляем Аспирантура
            str += "'" + this.PostGrad + "', ";
            //Добавляем ГАК
            str += "'" + this.GAK + "', ";
            //Добавляем Часы
            str += "'" + this.Hours + "', ";
            //Добавляем Часы Дано
            str += "'" + this.EnteredHours + "', ";
            //Добавляем Посещение занятий
            str += "'" + this.Visiting + "', ";
            //Добавляем Текст для диспетчерской
            str += "'" + this.Text + "', ";
            //Добавляем Магистратура
            str += "'" + this.Magistry + "', ";
            //Добавляем признак необходимости внесения
            //в заявку для диспетчерской
            str += this.flgDispatch + ", ";
            //Добавляем Дублёра
            if (this.Doubler != null)
            {
                str += "'" + this.Doubler.FIO + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }
            //Связь по лабораторным работам
            if (this.LabWorkConnect != null)
            {
                str += "'" + this.LabWorkConnect.Code + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }
            //Добавляем Часы по ЗЕТ
            str += "'" + this.HoursZ + "', ";
            //Добавляем Часы Дано по ЗЕТ
            str += "'" + this.EnteredHoursZ + "', ";
            //Добавляем признак распределяемой нагрузки
            str += this.flgDistrib + ", ";
            //Добавляем Вес единицы нагрузки
            str += "'" + this.Weight + "', ";
            //Добавляем признак исключения из расчёта нагрузки
            str += this.flgExclude + ", ";
            //Добавляем код дисциплины по документу
            str += "'" + this.DocCode + "', ";
            //Связь по почасовой нагрузке
            if (this.HouredConnect != null)
            {
                str += "'" + this.HouredConnect.Code + "'";
            }
            else
            {
                str += "'" + "" + "'";
            }
            return str;
        }
    }
}