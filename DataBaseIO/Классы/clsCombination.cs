using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataBaseIO
{
    class clsCombination
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Совместительство
        /// </summary>
        public string CombType;
        
        /// <summary>
        /// Процедура инициализации текущего элемента класса "Совместительство"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["№"].ToString());
            //Тип совместительства
            this.CombType = Tab.Rows[id]["Совместительство"].ToString();
        }

        public string Save(int id)
        {
            string str = "";

            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем Тип совместительства
            str += "'" + this.CombType + "'";

            return str;
        }
    }
}
