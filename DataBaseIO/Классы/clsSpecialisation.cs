using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;

namespace DataBaseIO
{
    class clsSpecialisation
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Специальность
        /// </summary>
        public string Specialisation;
        /// <summary>
        /// Сокращённое название специальности
        /// из учебного управления
        /// </summary>
        public string ShortUpravlenie;
        /// <summary>
        /// Сокращённое название специализации
        /// (дополнительно)
        /// </summary>
        public string ShortDop;
        /// <summary>
        /// Сокращённое название специальности
        /// учебные группы
        /// </summary>
        public string ShortInstitute;
        /// <summary>
        /// Факультет
        /// </summary>
        public clsFaculty Faculty;
        /// <summary>
        /// Образовательная система
        /// </summary>
        public string Diff;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Дополнительной работы"
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
            
            //Специальность
            this.Specialisation = Tab.Rows[id]["Специальность"].ToString();
            
            //Сокращённое название из управления
            this.ShortUpravlenie = Tab.Rows[id]["Сокращённое_название_специальности"].ToString();

            //Сокращённое название дополнительно
            this.ShortDop = Tab.Rows[id]["Сокращённое_название_специализации"].ToString();
            
            //Сокращённое название учебные группы
            this.ShortInstitute = Tab.Rows[id]["Аббревиатура"].ToString();
            
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
                MessageBox.Show("Не удалось определить факультет у элемента с кодом " + this.Code, "Оповещение");
            }

            //Образовательная система
            this.Diff = Tab.Rows[id]["Система"].ToString();
        }

        public string Save(int id)
        {
            string str = "";

            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем Специальность
            str += "'" + this.Specialisation + "', ";
            //Добавляем Сокращённое название специальность
            str += "'" + this.ShortUpravlenie + "', ";
            //Добавляем Сокращённое название учебных групп
            str += "'" + this.ShortInstitute + "', ";
           
            //Добавляем Факультет
            if (this.Faculty != null)
            {
                str += "'" + this.Faculty.Short + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }
            
            //Добавляем Сокращённое название специализации
            str += "'" + this.ShortDop + "', ";
            //Добавляем Образовательную систему
            str += "'" + this.Diff + "'";

            return str;
        }
    }
}
