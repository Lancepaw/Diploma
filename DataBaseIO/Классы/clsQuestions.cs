using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataBaseIO
{
    class clsQuestions
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Дата
        /// </summary>
        public DateTime Date;
        /// <summary>
        /// Вопрос
        /// </summary>
        public string Question;
        /// <summary>
        /// Докладчик 1
        /// </summary>
        public clsLecturer Speaker1;
        /// <summary>
        /// Докладчик 2
        /// </summary>
        public clsLecturer Speaker2;
        /// <summary>
        /// Докладчик 3
        /// </summary>
        public clsLecturer Speaker3;
        /// <summary>
        /// Докладчик 4
        /// </summary>
        public clsLecturer Speaker4;
        /// <summary>
        /// Докладчик 5
        /// </summary>
        public clsLecturer Speaker5;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Совместительство"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            int GetCode;
            bool Detected;
            string CurrentString;

            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["Код"].ToString());
            //Дата
            this.Date = Convert.ToDateTime(Tab.Rows[id]["Дата"].ToString());
            //Вопрос
            this.Question = Tab.Rows[id]["Вопрос"].ToString();
            
            //Докладчик 1           
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Докладчик1"].ToString();
                if (mdlData.colLecturer[i].FIO == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Speaker1 = mdlData.colLecturer[GetCode];
            }
            else
            {
                this.Speaker1 = null;
            }

            //Докладчик 2           
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Докладчик2"].ToString();
                if (mdlData.colLecturer[i].FIO == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Speaker2 = mdlData.colLecturer[GetCode];
            }
            else
            {
                this.Speaker2 = null;
            }

            //Докладчик 3           
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Докладчик3"].ToString();
                if (mdlData.colLecturer[i].FIO == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Speaker3 = mdlData.colLecturer[GetCode];
            }
            else
            {
                this.Speaker3 = null;
            }

            //Докладчик 4           
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Докладчик4"].ToString();
                if (mdlData.colLecturer[i].FIO == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Speaker4 = mdlData.colLecturer[GetCode];
            }
            else
            {
                this.Speaker4 = null;
            }

            //Докладчик 5           
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Докладчик5"].ToString();
                if (mdlData.colLecturer[i].FIO == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Speaker5 = mdlData.colLecturer[GetCode];
            }
            else
            {
                this.Speaker5 = null;
            }
        }

        public string Save(int id)
        {
            string str = "";

            //"Код", "Дата", "Вопрос", "Докладчик1", "Докладчик2", "Докладчик3", "Докладчик4", "Докладчик5"

            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем Дату
            str += mdlData.DateToSQL(this.Date) + ", ";
            //Добавляем Вопрос
            str += "'" + this.Question + "', ";
            //Добавляем Докладчика1
            if (this.Speaker1 != null)
            {
                str += "'" + this.Speaker1.FIO + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }
            
            //Добавляем Докладчика2
            if (this.Speaker2 != null)
            {
                str += "'" + this.Speaker2.FIO + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Докладчика3
            if (this.Speaker3 != null)
            {
                str += "'" + this.Speaker3.FIO + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Докладчика4
            if (this.Speaker4 != null)
            {
                str += "'" + this.Speaker4.FIO + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Докладчика5
            if (this.Speaker5 != null)
            {
                str += "'" + this.Speaker5.FIO + "'";
            }
            else
            {
                str += "'" + "" + "'";
            }

            return str;
        }
    }
}
