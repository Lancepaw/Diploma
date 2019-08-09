using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataBaseIO
{   
    class clsWeekDays
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Дни недели
        /// </summary>
        public string WeekDay;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Список_Дни недели"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["№"].ToString());
            //Название дня недели
            this.WeekDay = Tab.Rows[id]["Название"].ToString();
        }

        public string Save(int id)
        {
            string str = "";

            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем Название дня недели
            str += "'" + this.WeekDay + "'";

            return str;
        }
    }
}
