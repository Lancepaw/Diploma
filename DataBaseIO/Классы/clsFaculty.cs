using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataBaseIO
{
    class clsFaculty
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Наименование факультета
        /// </summary>
        public string Faculty;
        /// <summary>
        /// Короткое наименование факультета
        /// </summary>
        public string Short;
        /// <summary>
        /// Различие по образовательной системе
        /// </summary>
        public string Diff;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Должности"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["Код"].ToString());
            //Факультет
            this.Faculty = Tab.Rows[id]["Факультет"].ToString();
            //Короткое наименование факультета
            this.Short = Tab.Rows[id]["Сокр_название"].ToString();
            //Различие по системе образования
            this.Diff = Tab.Rows[id]["Система"].ToString();
        }

        public string Save(int id)
        {
            string str = "";

            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем Название
            str += "'" + this.Faculty + "', ";
            //Добавляем Коротко
            str += "'" + this.Short + "', ";
            //Добавляем Систему
            str += "'" + this.Diff + "'";

            return str;
        }
    }
}
