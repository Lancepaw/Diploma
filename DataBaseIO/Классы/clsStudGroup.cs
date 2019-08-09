using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataBaseIO
{
    class clsStudGroup
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Факультет
        /// </summary>
        public clsFaculty Faculty;
        /// <summary>
        /// Специальность
        /// </summary>
        public clsSpecialisation Spec;
        /// <summary>
        /// Курс
        /// </summary>
        public clsKursNum Kurs;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Преподаватели"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            int GetCode;
            string CurrentString;
            bool Detected;

            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["№"].ToString());

            //Факультет
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colFaculty.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Факультет"].ToString();
                if (mdlData.colFaculty[i].Short == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Faculty = mdlData.colFaculty[GetCode];
            }
            else
            {
                this.Faculty = null;
            }

            //Специальность
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colSpecialisation.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Специальность"].ToString();
                if (mdlData.colSpecialisation[i].ShortUpravlenie == CurrentString)
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
                this.Kurs = mdlData.colKursNum[GetCode];
            }
            else
            {
                this.Kurs = null;
            }
        }

        public string Save(int id)
        {
            string str = "";

            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем Факультет
            if (this.Faculty != null)
            {
                str += "'" + this.Faculty.Short + "', ";
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

            //Добавляем Курс
            if (this.Kurs != null)
            {
                str += this.Kurs.Kurs.ToString();
            }
            else
            {
                str += 0.ToString();
            }

            return str;
        }
    }
}
