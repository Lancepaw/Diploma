using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;

namespace DataBaseIO
{
    class clsSchedule
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Преподаватель
        /// </summary>
        public clsLecturer Lecturer;
        /// <summary>
        /// Неделя
        /// </summary>
        public clsWeek Week;
        /// <summary>
        /// День недели
        /// </summary>
        public clsWeekDays WeekDay;
        /// <summary>
        /// Время занятий
        /// </summary>
        public clsPairTime Time;
        /// <summary>
        /// Занятие
        /// </summary>
        public bool Subj;
        /// <summary>
        /// Номер аудитории
        /// </summary>
        public string Auditory;
        /// <summary>
        /// Связь с элементом расписания
        /// </summary>
        public clsDistributionDetailed Link;
        /// <summary>
        /// Читаемая дисциплина
        /// </summary>
        public clsSubject Subject;
        /// <summary>
        /// Тип проводимого занятия
        /// </summary>
        public clsSubjectType SubjectType;
        /// <summary>
        /// Специальность, для которой проводится занятие
        /// </summary>
        public clsSpecialisation Spec;
        /// <summary>
        /// Курс, на котором проводится занятие
        /// </summary>
        public clsKursNum KursNum;
        /// <summary>
        /// Курс, на котором проводится занятие
        /// </summary>
        public clsSemestr Semestr;
        /// <summary>
        /// Номер группы
        /// </summary>
        public string Group;
        /// <summary>
        /// Номер потока
        /// </summary>
        public string Stream;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Расписание_преподавателей"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            int GetCode;
            string CurrentString;
            bool Detected;
            
            //Код
            //this.Code = Convert.ToInt32(Tab.Rows[id]["ID"].ToString());
            this.Code = id + 1;

            //Преподаватель
            GetCode = 0;
            Detected = false;
            CurrentString = Tab.Rows[id]["Преподаватель"].ToString();
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
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
                MessageBox.Show("Не удалось определить преподавателя у элемента с кодом " + this.Code, "Оповещение");
            }

            //Неделя
            GetCode = 0;
            Detected = false;
            CurrentString = Tab.Rows[id]["Неделя"].ToString();
            for (int i = 0; i <= mdlData.colWeek.Count - 1; i++)
            {
                if (mdlData.colWeek[i].NumberWeek == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Week = mdlData.colWeek[GetCode];
            }
            else
            {
                this.Week = null;
                MessageBox.Show("Не удалось определить неделю у элемента с кодом " + this.Code, "Оповещение");
            }

            //День недели
            GetCode = 0;
            Detected = false;
            CurrentString = Tab.Rows[id]["День недели"].ToString();
            for (int i = 0; i <= mdlData.colWeekDays.Count - 1; i++)
            {
                if (mdlData.colWeekDays[i].WeekDay == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.WeekDay = mdlData.colWeekDays[GetCode];
            }
            else
            {
                this.WeekDay = null;
                MessageBox.Show("Не удалось определить день недели у элемента с кодом " + this.Code, "Оповещение");
            }

            //Время занятий
            GetCode = 0;
            Detected = false;
            CurrentString = Tab.Rows[id]["Время занятия"].ToString();

            for (int i = 0; i <= mdlData.colPairTime.Count - 1; i++)
            {
                if (mdlData.colPairTime[i].Time == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }

            //Заглушка потом можно будет убрать
            if (!Detected)
            {
                switch (CurrentString)
                {
                    case "8.00-9.30":
                        {
                            GetCode = 0;
                            Detected = true;
                            break;
                        }

                    case "9.40-11.10":
                        {
                            GetCode = 1;
                            Detected = true;
                            break;
                        }

                    case "11.20-12.50":
                        {
                            GetCode = 2;
                            Detected = true;
                            break;
                        }

                    case "13.20-14.50":
                        {
                            GetCode = 3;
                            Detected = true;
                            break;
                        }

                    case "15.00-16.30":
                        {
                            GetCode = 4;
                            Detected = true;
                            break;
                        }

                    case "16.40-18.10":
                        {
                            GetCode = 5;
                            Detected = true;
                            break;
                        }

                    case "18.20-19.50":
                        {
                            GetCode = 6;
                            Detected = true;
                            break;
                        }

                    case "20.00-21.30":
                        {
                            GetCode = 7;
                            Detected = true;
                            break;
                        }
                }
            }

            if (Detected)
            {
                this.Time = mdlData.colPairTime[GetCode];
            }
            else
            {
                this.Time = null;
                MessageBox.Show("Не удалось определить время занятия у элемента с кодом " + this.Code, "Оповещение");
            }

            //Занят?
            this.Subj = Convert.ToBoolean(Tab.Rows[id]["Занят"].ToString());

            //Аудитория
            this.Auditory = Tab.Rows[id]["Аудитория"].ToString();

            //Дисциплина
            GetCode = 0;
            Detected = false;
            CurrentString = Tab.Rows[id]["Дисциплина"].ToString();
            for (int i = 0; i <= mdlData.colSubject.Count - 1; i++)
            {
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

            //Тип занятия
            GetCode = 0;
            Detected = false;
            CurrentString = Tab.Rows[id]["Тип_Занятия"].ToString();
            for (int i = 0; i <= mdlData.colSubjectType.Count - 1; i++)
            {
                if (mdlData.colSubjectType[i].Type == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.SubjectType = mdlData.colSubjectType[GetCode];
            }
            else
            {
                this.SubjectType = null;
            }

            //Специальность
            GetCode = 0;
            Detected = false;
            CurrentString = Tab.Rows[id]["Специальность"].ToString();
            for (int i = 0; i <= mdlData.colSpecialisation.Count - 1; i++)
            {
                if (mdlData.colSpecialisation[i].ShortDop == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Spec = mdlData.colSpecialisation[GetCode];
            }
            else
            {
                this.Spec = null;
            }

            //Номер курса
            GetCode = 0;
            Detected = false;
            CurrentString = Tab.Rows[id]["Курс"].ToString();
            for (int i = 0; i <= mdlData.colKursNum.Count - 1; i++)
            {
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
            
            //Семестр
            GetCode = 0;
            Detected = false;
            CurrentString = Tab.Rows[id]["Семестр"].ToString();
            for (int i = 0; i <= mdlData.colSemestr.Count - 1; i++)
            {
                if (mdlData.colSemestr[i].SemNum.ToString() == CurrentString)
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
            }

            //Группа
            this.Group = Tab.Rows[id]["Группа"].ToString();
            //Поток
            this.Stream = Tab.Rows[id]["Поток"].ToString();
        }

        public string Save(int id)
        {
            string str = "";

            //"ID", "Преподаватель", "День недели", "Время занятия", "Неделя", "Занят"

            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем Преподавателя
            if (this.Lecturer != null)
            {
                str += "'" + this.Lecturer.FIO + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем День Недели
            if (this.WeekDay != null)
            {
                str += "'" + this.WeekDay.WeekDay + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Время занятия
            if (this.Time!= null)
            {
                str += "'" + this.Time.Time + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Неделю
            if (this.Week != null)
            {
                str += "'" + this.Week.NumberWeek + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Занят?
            str += this.Subj.ToString() + ", ";

            //Добавляем Дисциплину
            if (this.Subject != null)
            {
                str += "'" + this.Subject.Subject + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Тип занятия
            if (this.SubjectType != null)
            {
                str += "'" + this.SubjectType.Type + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Специальность
            if (this.Spec != null)
            {
                str += "'" + this.Spec.ShortDop + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Номер курса
            if (this.KursNum != null)
            {
                str += "'" + this.KursNum.Kurs + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Аудиторию
            str += "'" + this.Auditory + "', ";

            //Добавляем Семестр
            if (this.Semestr != null)
            {
                str += "'" + this.Semestr.SemNum + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Группу
            str += "'" + this.Group + "', ";

            //Добавляем Поток
            str += "'" + this.Stream + "'";

            return str;
        }
    }
}
