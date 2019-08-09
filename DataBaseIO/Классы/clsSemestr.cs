using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataBaseIO
{
    class clsSemestr
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Номер семестра
        /// </summary>
        public string SemNum;
        /// <summary>
        /// В предложном падеже
        /// </summary>
        public string About;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Семестров"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["Код"].ToString());
            //Номер семестра
            this.SemNum = Tab.Rows[id]["Номер_семестра"].ToString();
            //Семестр в предложном падеже
            this.About = Tab.Rows[id]["Предложный_падеж"].ToString();
        }

        public string Save(int id)
        {
            string str = "";

            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем Название
            str += "'" + this.SemNum + "', ";
            //Добавляем Семестр в предложном падеже
            str += "'" + this.About + "'";

            return str;
        }
    }
}
