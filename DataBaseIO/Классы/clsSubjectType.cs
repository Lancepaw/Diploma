using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace DataBaseIO
{
    class clsSubjectType
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Вид учебного занятия
        /// </summary>
        public string Type;
        /// <summary>
        /// Вид учебного занятия (кратко)
        /// </summary>
        public string Short;
        /// <summary>
        /// Вид учебного занятия (кратко) для индивидуальных планов
        /// </summary>
        public string ShortPlan;
        /// <summary>
        /// Вид учебного занятия как в таблице распределения нагрузки
        /// </summary>
        public string LikeDistrib;
        /// <summary>
        /// Вид учебного занятия для печатных форм
        /// </summary>
        public string ForForms;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Вид занятия"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["Код"].ToString());
            //Вид
            this.Type = Tab.Rows[id]["Вид_занятия"].ToString();
            //Кратко
            this.Short = Tab.Rows[id]["Кратко"].ToString();
            //Кратко для индивидуальных планов
            this.ShortPlan = Tab.Rows[id]["КраткоИП"].ToString();
            //Как указано в таблице распределения учебной нагрузки
            this.LikeDistrib = Tab.Rows[id]["В_распределении"].ToString();
            //Вид для вставки в печатные формы
            this.ForForms = Tab.Rows[id]["Для_форм"].ToString();
        }

        public string Save(int id)
        {
            string str = "";

            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем Вид
            str += "'" + this.Type + "', ";
            //Добавляем Коротко
            str += "'" + this.Short + "', ";
            //Добавляем Коротко для индивидуального плана
            str += "'" + this.ShortPlan + "', ";
            //Добавляем надпись как в таблице распределения учебной нагрузки
            str += "'" + this.LikeDistrib + "', ";
            //Добавляем надпись для печатных форм
            str += "'" + this.ForForms + "'";

            return str;
        }
    }
}
