using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace DataBaseIO
{
    class clsSummary
    {
        public int Code;
        public int LectCounter;
        public int ExamCounter;
        public int CredCounter;
        public int RefCounter;
        public int TutCounter;
        public int LabCounter;
        public int PractCounter;

        public int IndCounter;
        public int KRAPKCounter;
        public int KursCounter;
        public int PreDCounter;
        public int DiplomaCounter;
        public int TutPrCounter;
        public int ProdCounter;
        public int GAKCounter;
        public int BudCounter;
        public int SumCounter;
        public int AllCounter;

        public float BudZCounter;
        public float SumZCounter;
        public float AllZCounter;

        public void Initialize(DataTable Tab, int id)
        {
            //Код
            this.Code = Convert.ToInt32(Tab.Rows[id]["Код"].ToString());
            
            //Количество лекционных часов
            this.LectCounter = Convert.ToInt32(Tab.Rows[id]["Лекции"].ToString());
            //Количество экзаменационных часов
            this.ExamCounter = Convert.ToInt32(Tab.Rows[id]["Экзамен"].ToString());
            //Количество зачётных часов
            this.CredCounter = Convert.ToInt32(Tab.Rows[id]["Зачёт"].ToString());
            //Количество часов на реферативную работу
            this.RefCounter = Convert.ToInt32(Tab.Rows[id]["Реферат"].ToString());
            //Количество часов на консультативную работу
            this.TutCounter = Convert.ToInt32(Tab.Rows[id]["Консультация"].ToString());
            //Количество часов на лабораторные работы
            this.LabCounter = Convert.ToInt32(Tab.Rows[id]["Лабораторные"].ToString());
            //Количество часов на правктические работы
            this.PractCounter = Convert.ToInt32(Tab.Rows[id]["Практические"].ToString());
        }
    }
}
