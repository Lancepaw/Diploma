using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataBaseIO
{
    class clsPGStudents
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
        /// Кафедра
        /// </summary>
        public clsDepartment Depart;
        /// <summary>
        /// Руководитель
        /// </summary>
        public clsLecturer Lect;
        /// <summary>
        /// Курс
        /// </summary>
        public clsKursNum KursNum;
        /// <summary>
        /// Признак участия в индивидуальном плане
        /// </summary>
        public bool flgPlan;
        /// <summary>
        /// Тема работы
        /// </summary>
        public string Theme;
        /// <summary>
        /// Количество выделенных часов
        /// </summary>
        public int Hours;
        /// <summary>
        /// Признак бюджетного обучения
        /// </summary>
        public bool flgBudget;

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

            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["Код"].ToString());

            //ФИО
            this.FIO = Tab.Rows[id]["ФИО"].ToString();

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

            //Признак участия в индивидуальном плане преподавателя
            if (Tab.Rows[id]["В_плане"].ToString() == "False")
            {
                this.flgPlan = false;
            }
            else
            {
                this.flgPlan = true;
            }

            //Тема_работы
            this.Theme = Tab.Rows[id]["Тема_работы"].ToString();

            //Количество выделенных часов
            if (!(Tab.Rows[id]["Часы"] is DBNull))
            {
                this.Hours = Convert.ToInt32(Tab.Rows[id]["Часы"].ToString());
            }
            else
            {
                this.Hours = 0;
            }

            //Признак бюджетной нагрузки
            if (Tab.Rows[id]["Бюджет"].ToString() == "False")
            {
                this.flgBudget = false;
            }
            else
            {
                this.flgBudget = true;
            }
        }

        public string Save(int id)
        {
            string str = "";

            //Добавляем Код
            str += this.Code.ToString() + ", ";

            //Добавляем ФИО
            str += "'" + this.FIO.ToString() + "', ";

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

            //Добавляем Курс
            if (this.KursNum != null)
            {
                str += "'" + this.KursNum.Kurs + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем признак участия в индивидуальном плане преподавателя
            str += this.flgPlan + ", ";

            //Добавляем Тему_работы
            str += "'" + this.Theme + "', ";

            //Добавляем Часы
            str += "'" + this.Hours + "', ";

            //Добавляем признак бюджетной нагрузки
            str += this.flgBudget + "";

            return str;
        }
    }
}
