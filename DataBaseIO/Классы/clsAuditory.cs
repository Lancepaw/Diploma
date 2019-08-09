using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataBaseIO
{
    class clsAuditory
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Номер аудитории
        /// </summary>
        public int AuditoryNum;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Аудитории"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["№"].ToString());
            //Номер аудитории
            this.AuditoryNum = Convert.ToInt32(Tab.Rows[id]["Номер аудитории"].ToString());
        }

        public string Save(int id)
        {
            string str = "";

            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем Название
            str += this.AuditoryNum.ToString();

            return str;
        }
    }
}
