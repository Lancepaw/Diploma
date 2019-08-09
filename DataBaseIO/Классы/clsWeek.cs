using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataBaseIO

{
    class clsWeek
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Номер недели
        /// </summary>
        public string NumberWeek;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Недели_1-2"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["Код"].ToString());
            //Номер недели
            this.NumberWeek = Tab.Rows[id]["Номер_недели"].ToString();
        }

        public string Save(int id)
        {
            string str = "";

            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем Номер недели
            str += "'" + this.NumberWeek + "'";

            return str;
        }
    }
}
