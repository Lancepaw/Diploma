using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataBaseIO
{
    class clsWorkYear
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Учебный год
        /// </summary>
        public string WorkYear;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Учебные годы"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["Код"].ToString());
            //Название
            this.WorkYear = Tab.Rows[id]["Учебный год"].ToString();
        }

        public string Save(int id)
        {
            string str = "";
            
            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем Название
            str += "'" + this.WorkYear + "'";

            return str;
        }
    }
}
