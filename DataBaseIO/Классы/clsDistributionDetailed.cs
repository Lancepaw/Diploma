using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;

namespace DataBaseIO
{
    class clsDistributionDetailed
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Семестр
        /// </summary>
        public clsSemestr Semestr;
        /// <summary>
        /// Специальность
        /// </summary>
        public clsSpecialisation Speciality;
        /// <summary>
        /// Курс
        /// </summary>
        public clsKursNum KursNum;
        /// <summary>
        /// Название учебной дисциплины
        /// </summary>
        public clsSubject Subject;
        /// <summary>
        /// Преподаватель
        /// </summary>
        public clsLecturer Lecturer;
        /// <summary>
        /// Замещающий преподаватель
        /// </summary>
        public clsLecturer Lecturer2;
        /// <summary>
        /// Зам. замещающего преподавателя
        /// </summary>
        public clsLecturer Lecturer3;
        /// <summary>
        /// Преподаватель-дублёр
        /// </summary>
        public clsLecturer Doubler;
        /// <summary>
        /// Вид учебной нагрузки
        /// </summary>
        public clsSubjectType SubjType;
        /// <summary>
        /// Часы, выделяемые на вид нагрузки
        /// </summary>
        public float SubjHours;
        /// <summary>
        /// Примечание
        /// </summary>
        public string Text;
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

            this.Text = Tab.Rows[id]["Для_диспетчерской"].ToString();

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

            //Добавляем Текст для диспетчерской
            str += "'" + this.Text + "', ";
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
            //Добавляем признак распределяемой нагрузки
            str += this.flgDistrib + ", ";
            //Добавляем Вес единицы нагрузки
            str += "'" + this.Weight + "', ";
            //Добавляем признак исключения из расчёта нагрузки
            str += this.flgExclude + "";

            return str;
        }
    }
}