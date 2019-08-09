using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;

namespace DataBaseIO
{
    class clsLecturer
    {
        /// <summary>
        /// Код в базе данных
        /// </summary>
        public int Code;
        /// <summary>
        /// Фамилия, Имя, Отчество
        /// </summary>
        public string FIO;
        /// <summary>
        /// Кафедра
        /// </summary>
        public clsDepartment Depart;
        /// <summary>
        /// Год рождения
        /// </summary>
        public DateTime DOB;
        /// <summary>
        /// Должность
        /// </summary>
        public clsDuty Duty;
        /// <summary>
        /// Дополнительная должность
        /// </summary>
        public clsDuty Duty1;
        /// <summary>
        /// Учёное звание
        /// </summary>
        public clsStatus Status;
        /// <summary>
        /// Учёная степень
        /// </summary>
        public clsDegree Degree;
        /// <summary>
        /// Совместительство
        /// </summary>
        public clsCombination Combination;
        /// <summary>
        /// Общий стаж работы
        /// </summary>
        public DateTime Seniority;
        /// <summary>
        /// Ставка
        /// </summary>
        public double Rate;
        /// <summary>
        /// Максимальная нагрузка
        /// </summary>
        public int MaxLoad;
        /// <summary>
        /// Разгрузка
        /// </summary>
        public int UnLoad;
        /// <summary>
        /// Примечание
        /// </summary>
        public string Text;
        /// <summary>
        /// Пожелания по аудиториям
        /// </summary>
        public string Preferences;
        /// <summary>
        /// Членство в совете
        /// </summary>
        public bool Soviet;
        /// <summary>
        /// Старая ставка
        /// </summary>
        public double OldRate;
        /// <summary>
        /// Ставка в первом семестре
        /// </summary>
        public double Rate1;
        /// <summary>
        /// Ставка во втором семестре
        /// </summary>
        public double Rate2;
        /// <summary>
        /// Переменная ставка
        /// </summary>
        public bool ChangeRate;

        /// <summary>
        /// Процедура инициализации текущего элемента класса "Преподаватели"
        /// </summary>
        /// <param name="Tab">Образ таблицы базы данных</param>
        /// <param name="id">Идентификатор текущего элемента</param>
        public void Initialize(DataTable Tab, int id)
        {
            int GetCode;
            string CurrentString;
            bool Detected;

            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["№"].ToString());
            
            //Фамилия, Имя, Отчество
            this.FIO = Tab.Rows[id]["ФИО"].ToString();
            
            //Кафедра
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colDepart.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Кафедра"].ToString();
                if (mdlData.colDepart[i].Short == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Depart = mdlData.colDepart[GetCode];
            }
            else
            {
                this.Depart = null;
            }
            
            //Год рождения
            if (!(Tab.Rows[id]["Год_рождения"] is DBNull))
            {
                this.DOB = Convert.ToDateTime(Tab.Rows[id]["Год_рождения"].ToString());
            }
            else
            {
                this.DOB = new DateTime(2014,01,01);
            }
            
            //Должность
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colDuty.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Должность"].ToString();
                if (mdlData.colDuty[i].Duty == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Duty = mdlData.colDuty[GetCode];
            }
            else
            {
                this.Duty = null;
            }
            
            //Дополнительная должность
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colDuty.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Должность1"].ToString();
                if (mdlData.colDuty[i].Duty == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Duty1 = mdlData.colDuty[GetCode];
            }
            else
            {
                this.Duty1 = null;
            }
           
            //Учёное звание
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colStatus.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Учёное_звание"].ToString();
                if (mdlData.colStatus[i].Status == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Status = mdlData.colStatus[GetCode];
            }
            else
            {
                this.Status = null;
            }

            //Учёная степень
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colDegree.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Учёная_степень"].ToString();
                if (mdlData.colDegree[i].Short == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Degree = mdlData.colDegree[GetCode];
            }
            else
            {
                this.Degree = null;
            }
            
            //Совместительство
            GetCode = 0;
            Detected = false;
            for (int i = 0; i <= mdlData.colCombination.Count - 1; i++)
            {
                CurrentString = Tab.Rows[id]["Совместительство"].ToString();
                if (mdlData.colCombination[i].CombType == CurrentString)
                {
                    GetCode = i;
                    Detected = true;
                }
            }
            if (Detected)
            {
                this.Combination = mdlData.colCombination[GetCode];
            }
            else
            {
                this.Combination = null;
                MessageBox.Show("Не удалось определить совместительство у элемента с кодом " + this.Code, "Оповещение");
            }
            
            //Общий стаж работы
            if (!(Tab.Rows[id]["Общий_стаж_работы"] is DBNull))
            {
                this.Seniority = Convert.ToDateTime(Tab.Rows[id]["Общий_стаж_работы"].ToString());
            }
            else
            {
                this.Seniority = new DateTime(2014, 01, 01); ;
            }
            //Ставка
            if (!(Tab.Rows[id]["Ставка"] is DBNull))
            {
                this.Rate = Convert.ToDouble(Tab.Rows[id]["Ставка"].ToString());
            }
            else
            {
                this.Rate = 0;
            }
            //Максимальная нагрузка
            if (!(Tab.Rows[id]["Максимальная_нагрузка"] is DBNull))
            {
                this.MaxLoad = Convert.ToInt32(Tab.Rows[id]["Максимальная_нагрузка"].ToString());
            }
            else
            {
                this.MaxLoad = 0;
            }
            //Разгрузка
            this.UnLoad = Convert.ToInt32(Tab.Rows[id]["Разгрузка"].ToString());
            //Примечание
            this.Text = Tab.Rows[id]["Для_диспетчерской"].ToString();
            //Пожелания по аудиториям
            this.Preferences = Tab.Rows[id]["Аудитории"].ToString();
            //Членство в совете
            if (Tab.Rows[id]["Совет"].ToString() == "False")
            {
                this.Soviet = false;
            }
            else
            {
                this.Soviet = true;
            }
            //Старая ставка
            if (!(Tab.Rows[id]["Старая_ставка"] is DBNull))
            {
                this.OldRate = Convert.ToDouble(Tab.Rows[id]["Старая_ставка"].ToString());
            }
            else
            {
                this.OldRate = 0;
            }

            //Ставка1
            if (!(Tab.Rows[id]["Ставка1"] is DBNull))
            {
                this.Rate1 = Convert.ToDouble(Tab.Rows[id]["Ставка1"].ToString());
            }
            else
            {
                this.Rate1 = 0;
            }

            //Ставка2
            if (!(Tab.Rows[id]["Ставка2"] is DBNull))
            {
                this.Rate2 = Convert.ToDouble(Tab.Rows[id]["Ставка2"].ToString());
            }
            else
            {
                this.Rate2 = 0;
            }

            //Переменная ставка
            if (Tab.Rows[id]["Смена_ставок"].ToString() == "False")
            {
                this.ChangeRate = false;
            }
            else
            {
                this.ChangeRate = true;
            }
        }

        public string Save(int id)
        {
            string str = "";
            
            //Добавляем Код
            str += id.ToString() + ", ";
            //Добавляем ФИО
            str += "'" + this.FIO + "', ";
            
            //Добавляем Кафедру
            if (this.Depart != null)
            {
                str += "'" + this.Depart.Short + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Дату рождения
            str += mdlData.DateToSQL(this.DOB) + ", ";

            //Добавляем Должность
            if (this.Duty != null)
            {
                str += "'" + this.Duty.Duty + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Звание
            if (this.Status != null)
            {
                str += "'" + this.Status.Status + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Степень
            if (this.Degree != null)
            {
                str += "'" + this.Degree.Short + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Совместительство
            if (this.Combination != null)
            {
                str += "'" + this.Combination.CombType + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Стаж работы
            str += mdlData.DateToSQL(this.Seniority) + ", ";
            //Добавляем Ставку
            str += "'" + this.Rate + "', ";
            //Добавляем Максимальную разгрузку
            str += "'" + this.MaxLoad + "', ";

            //Добавляем Дополнительную Должность
            if (this.Duty1 != null)
            {
                str += "'" + this.Duty1.Duty + "', ";
            }
            else
            {
                str += "'" + "" + "', ";
            }

            //Добавляем Разгрузку
            str += "'" + this.UnLoad + "', ";
            //Добавляем Текст для диспетчерской
            str += "'" + this.Text + "', ";
            //Добавляем Признак совета
            str += this.Soviet + ", ";
            //Добавляем Старую ставку
            str += "'" + this.OldRate + "', ";
            //Добавляем Пожелания по аудиториям
            str += "'" + this.Preferences + "', ";
            //Добавляем Ставку первого семестра
            str += "'" + this.Rate1 + "', ";
            //Добавляем Ставку второго семестра
            str += "'" + this.Rate2 + "', ";
            //Добавляем Признак смены ставок
            str += this.ChangeRate + "";

            return str;
        }

    }
}
