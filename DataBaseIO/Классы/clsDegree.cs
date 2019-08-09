using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataBaseIO
{
    class clsDegree
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Название степени
        /// </summary>
        public string Degree;
        /// <summary>
        /// Название степени коротко
        /// </summary>
        public string Short;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Степень"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["№"].ToString());
            //Название степени
            this.Degree = Tab.Rows[id]["Степень"].ToString();
            //Название степени коротко
            this.Short = Tab.Rows[id]["Коротко"].ToString();
        }

        public string Save(int id)
        {
            string str = "";

            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем Название
            str += "'" + this.Degree + "', ";
            //Добавляем Коротко
            str += "'" + this.Short + "'";

            return str;
        }
    }
}
