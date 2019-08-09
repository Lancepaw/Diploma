using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataBaseIO
{
    class clsStudents
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Фамилия, Имя, Отчество
        /// </summary>
        public string FIO;
        /// <summary>
        /// Специальность
        /// </summary>
        public clsSpecialisation Speciality;
        /// <summary>
        /// Курс
        /// </summary>
        public clsKursNum KursNum;
        /// <summary>
        /// Кафедра
        /// </summary>
        public clsDepartment Depart;
        /// <summary>
        /// Руководитель
        /// </summary>
        public clsLecturer Lect;
        /// <summary>
        /// Тема работы
        /// </summary>
        public string Theme;
        /// <summary>
        /// Признак участия в индивидуальном плане
        /// </summary>
        public bool flgPlan;
        /// <summary>
        /// Признак участия в почасовой
        /// </summary>
        public bool flgHoured;

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

            //ФИО
            this.FIO = Tab.Rows[id]["ФИО"].ToString();

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

            //Кафедра
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colDepart.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Кафедра"].ToString();
                if (mdlData.colDepart[i].Short == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Depart = mdlData.colDepart[GetCode];
            }
            else
            {
                this.Depart = null;
            }

            //Руководитель
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Руководитель"].ToString();
                if (mdlData.colLecturer[i].FIO == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Lect = mdlData.colLecturer[GetCode];
            }
            else
            {
                this.Lect = null;
            }

            //Тема_работы
            this.Theme = Tab.Rows[id]["Тема_работы"].ToString();

            //Признак участия в индивидуальном плане преподавателя
            if (Tab.Rows[id]["В_плане"].ToString() == "False")
            {
                this.flgPlan = false;
            }
            else
            {
                this.flgPlan = true;
            }

            //Признак участия в почасовой
            if (Tab.Rows[id]["Почасовая"].ToString() == "False")
            {
                this.flgHoured = false;
            }
            else
            {
                this.flgHoured = true;
            }
        }

        public string Save(int id)
        {
            string str = "";

            //Добавляем Код
            //str += this.Code.ToString() + ", ";
            str += id.ToString() + ", ";

            //Добавляем ФИО
            str += "'" + this.FIO.ToString() + "', ";

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

            //Добавляем Кафедру
            if (this.Depart != null)
            {
                str += "'" + this.Depart.Short + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Руководителя
            if (this.Lect != null)
            {
                str += "'" + this.Lect.FIO + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Тему_работы
            str += "'" + this.Theme + "', ";

            //Добавляем признак участия в индивидуальном плане преподавателя
            str += this.flgPlan + ", ";

            //Добавляем признак участия в почасовой
            str += this.flgHoured + "";

            return str;
        }
    }
}