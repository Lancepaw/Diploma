using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataBaseIO
{
    class clsPairTime
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Время пары
        /// </summary>
        public string Time;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Список_Время пар"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["№"].ToString());
            //Время пары
            this.Time = Tab.Rows[id]["Время"].ToString();
        }

        public string Save(int id)
        {
            string str = "";

            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем Время пары
            str += "'" + this.Time + "'";

            return str;
        }
    }
}
