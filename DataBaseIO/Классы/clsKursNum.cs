using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataBaseIO
{
    class clsKursNum
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Номер курса
        /// </summary>
        public int Kurs;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Номеров курсов"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["Код"].ToString());
            //Название
            this.Kurs = Convert.ToInt32(Tab.Rows[id]["Курс"].ToString());
        }

        public string Save(int id)
        {
            string str = "";

            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем Название
            str += this.Kurs.ToString();

            return str;
        }
    }
}
