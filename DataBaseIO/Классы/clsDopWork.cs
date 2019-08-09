using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;

namespace DataBaseIO
{
    class clsDopWork
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Преподаватель
        /// </summary>
        public clsLecturer Lecturer;
        /// <summary>
        /// Семестр
        /// </summary>
        public clsSemestr Semestr;
        /// <summary>
        /// Учебно-Методическая Работа
        /// </summary>
        public string UMR;
        /// <summary>
        /// Научно-Исследовательская Работа
        /// </summary>
        public string NIR;
        /// <summary>
        /// Организационно-Методическая Работа
        /// </summary>
        public string OMR;
        /// <summary>
        /// Объём УМР
        /// </summary>
        public string VolumeUMR;
        /// <summary>
        /// Объём НИР
        /// </summary>
        public string VolumeNIR;
        /// <summary>
        /// Объём ОМР
        /// </summary>
        public string VolumeOMR;
        /// <summary>
        /// Срок УМР
        /// </summary>
        public string DateUMR;
        /// <summary>
        /// Срок НИР
        /// </summary>
        public string DateNIR;
        /// <summary>
        /// Срок ОМР
        /// </summary>
        public string DateOMR;
        /// <summary>
        /// Отметка УМР
        /// </summary>
        public string CommUMR;
        /// <summary>
        /// Отметка НИР
        /// </summary>
        public string CommNIR;
        /// <summary>
        /// Отметка ОМР
        /// </summary>
        public string CommOMR;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Дополнительной работы"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            int GetCode;
            string CurrentString;
            bool Detected;

            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["Код"].ToString());
            
            //Преподаватель
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Преподаватель"].ToString();
                if (mdlData.colLecturer[i].FIO == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Lecturer = mdlData.colLecturer[GetCode];
            }
            else
            {
                this.Lecturer = null;
                MessageBox.Show("Не удалось определить преподавателя у элемента с кодом " + this.Code, "Оповещение");
            }
            
            //Семестр
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colSemestr.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Семестр"].ToString();
                if (mdlData.colSemestr[i].SemNum == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Semestr = mdlData.colSemestr[GetCode];
            }
            else
            {
                this.Lecturer = null;
                MessageBox.Show("Не удалось определить семестр у элемента с кодом " + this.Code, "Оповещение");
            }

            //Учебно-Методическая Работа
            this.UMR = Tab.Rows[id]["УчебноМР"].ToString();
            //Научно-Исследовательская Работа
            this.NIR = Tab.Rows[id]["НаучИсР"].ToString();
            //Организационно-Методическая Работа
            this.OMR = Tab.Rows[id]["ОргМР"].ToString();
            
            //Обёъм УМР
            this.VolumeUMR = Tab.Rows[id]["ОбъемУМР"].ToString();
            //Объём НИР
            this.VolumeNIR = Tab.Rows[id]["ОбъемНИР"].ToString();
            //Объём ОМР
            this.VolumeOMR = Tab.Rows[id]["ОбъемОМР"].ToString();

            //Срок УМР
            this.DateUMR = Tab.Rows[id]["СрокУМР"].ToString();
            //Срок НИР
            this.DateNIR = Tab.Rows[id]["СрокНИР"].ToString();
            //Срок ОМР
            this.DateOMR = Tab.Rows[id]["СрокОМР"].ToString();

            //Отметка УМР
            this.CommUMR = Tab.Rows[id]["КомУМР"].ToString();
            //Отметка НИР
            this.CommNIR = Tab.Rows[id]["КомНИР"].ToString();
            //Отметка ОМР
            this.CommOMR = Tab.Rows[id]["КомОМР"].ToString();
        }

        public string Save(int id)
        {
            string str = "";

            //"Код", "Преподаватель", "Семестр", "УчебноМР", "НаучИсР", "ОргМР"

            //Добавляем Код
            str += id.ToString() + ", ";
            
            //Добавляем Преподавателя
            if (this.Lecturer != null)
            {
                str += "'" + this.Lecturer.FIO + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Семестр
            if (this.Semestr != null)
            {
                str += "'" + this.Semestr.SemNum + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем УМР
            str += "'" + this.UMR + "', ";
            //Добавляем НИР
            str += "'" + this.NIR + "', ";
            //Добавляем ОМР
            str += "'" + this.OMR + "', ";

            //Добавляем объём УМР
            str += "'" + this.VolumeUMR + "', ";
            //Добавляем объём НИР
            str += "'" + this.VolumeNIR + "', ";
            //Добавляем объём ОМР
            str += "'" + this.VolumeOMR + "', ";

            //Добавляем срок УМР
            str += "'" + this.DateUMR + "', ";
            //Добавляем срок НИР
            str += "'" + this.DateNIR + "', ";
            //Добавляем срок ОМР
            str += "'" + this.DateOMR + "', ";

            //Добавляем отметку УМР
            str += "'" + this.CommUMR + "', ";
            //Добавляем отметку НИР
            str += "'" + this.CommNIR + "', ";
            //Добавляем отметку ОМР
            str += "'" + this.CommOMR + "'";

            return str;
        }
    }
}
