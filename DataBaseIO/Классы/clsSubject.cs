using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataBaseIO
{
    class clsSubject
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Наименование дисциплины
        /// </summary>
        public string Subject;
        /// <summary>
        /// Наименование дисциплины коротко
        /// </summary>
        public string SubjectShort;
        /// <summary>
        /// Предпочтения по аудиториям
        /// </summary>
        public string Preferences;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Дисциплин"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["№"].ToString());
            //Название
            this.Subject = Tab.Rows[id]["Название_дисциплины"].ToString();
            //Короткое название дисциплины
            this.SubjectShort = Tab.Rows[id]["Сокращённое_название_дисциплины"].ToString();
            //Предпочтения по аудиториям
            this.Preferences = Tab.Rows[id]["Аудитории"].ToString();
        }

        public string Save(int id)
        {
            string str = "";

            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем Название
            str += "'" + this.Subject + "', ";
            //Добавляем Короткое название дисциплины
            str += "'" + this.SubjectShort + "', ";
            //Добавляем Пожелания по аудиториям
            str += "'" + this.Preferences + "'";

            return str;
        }
    }
}
