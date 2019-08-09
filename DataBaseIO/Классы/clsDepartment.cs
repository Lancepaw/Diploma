using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataBaseIO
{
    class clsDepartment
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Название кафедры
        /// </summary>
        public string Kafedra;
        /// <summary>
        /// Короткое название кафедры
        /// </summary>
        public string Short;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Кафедры"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["Код"].ToString());
            //Название кафедры
            this.Kafedra = Tab.Rows[id]["Кафедра"].ToString();
            //Короткое название кафедры
            this.Short = Tab.Rows[id]["Коротко"].ToString();
        }

        public string Save(int id)
        {
            string str = "";

            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем Название
            str += "'" + this.Kafedra + "', ";
            //Добавляем Коротко
            str += "'" + this.Short + "'";

            return str;
        }
    }
}
