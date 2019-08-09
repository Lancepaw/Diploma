using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace DataBaseIO
{
    class clsSickList
    {
        /// <summary>
        /// Идентификатор больничного листа
        /// </summary>
        public int Code;

        /// <summary>
        /// Семестр, к которому относится больничный лист
        /// </summary>
        public clsSemestr Semestr;

        /// <summary>
        /// Преподаватель, подавший больничный лист
        /// </summary>
        public clsLecturer Lecturer;

        /// <summary>
        /// Дата открытия больничного листа
        /// </summary>
        public DateTime OpenDate;

        /// <summary>
        /// Дата закрытия больничного листа
        /// </summary>
        public DateTime CloseDate;

        /// <summary>
        /// Примечание к больничному листу
        /// </summary>
        public string Descript;

        public void Initialize(DataTable Tab, int id)
        {
            int GetCode;
            string CurrentString;
            bool Detected;

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

            //Дата открытия больничного листа
            if (!(Tab.Rows[id]["Открытие"] is DBNull))
            {
                this.OpenDate = Convert.ToDateTime(Tab.Rows[id]["Открытие"].ToString());
            }
            else
            {
                this.OpenDate = new DateTime(2014, 01, 01);
            }

            //Дата закрытия больничного листа
            if (!(Tab.Rows[id]["Закрытие"] is DBNull))
            {
                this.CloseDate = Convert.ToDateTime(Tab.Rows[id]["Закрытие"].ToString());
            }
            else
            {
                this.CloseDate = new DateTime(2014, 01, 01);
            }

            //Примечание
            this.Descript = Tab.Rows[id]["Примечание"].ToString();
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

            //Добавляем Преподавателя
            if (this.Lecturer != null)
            {
                str += "'" + this.Lecturer.FIO + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Дату открытия больничного листа
            str += mdlData.DateToSQL(this.OpenDate) + ", ";

            //Добавляем Дату закрытия больничного листа
            str += mdlData.DateToSQL(this.CloseDate) + ", ";

            //Добавляем примечание
            str += "'" + this.Descript + "'";

            return str;
        }
    }
}
