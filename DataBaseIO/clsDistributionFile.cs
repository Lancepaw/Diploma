using System.Linq;
using System.Data;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace DataBaseIO
{
    class clsDistributionFile
    {
        /// <summary>
        /// 01. № п/п
        /// </summary>
        public int NumPP;
        /// <summary>
        /// 02. № ОУП
        /// </summary>
        public int NumOUP;
        /// <summary>
        /// 03. Факультет *
        /// </summary>
        public string Faculty;
        /// <summary>
        /// 04. Курс *
        /// </summary>
        public int Course;
        /// <summary>
        /// 05. Кол-во недель
        /// </summary>
        public int NumOfWeeks;
        /// <summary>
        /// 06. Сокр. назв. сп-ти
        /// </summary>
        public string ShortSpecialisation1;
        /// <summary>
        /// 07. Сокр. назв. сп-ти
        /// </summary>
        public string ShortSpecialisation2;
        /// <summary>
        /// 08. Кол-во студентов
        /// </summary>
        public int Students;
        /// <summary>
        /// 09. Кол-во студентов(ф.б.)
        /// </summary>
        public int StudentsFB;
        /// <summary>
        /// 10. Кол-во потоков
        /// </summary>
        public int Flow;
        /// <summary>
        /// 11. Кол-во групп
        /// </summary>
        public int Groups;
        /// <summary>
        /// 12. Название дисциплины *
        /// </summary>
        public string Discipline;
        /// <summary>
        /// 13. Лекции Конс. *
        /// </summary>
        public int LecturesConsult;
        /// <summary>
        /// 14. Экзамены *
        /// </summary>
        public int Exam;
        /// <summary>
        /// 15. Зачёты *
        /// </summary>
        public int Credit;
        /// <summary>
        /// 16. Реферат, дом.зад, Эссе *
        /// </summary>
        public int HomeWork;
        /// <summary>
        /// 17. Консультации
        /// </summary>
        public int Consultations;
        /// <summary>
        /// 18. Лабораторные работы *
        /// </summary>
        public int LaboratoryWorks;
        /// <summary>
        /// 19. Практ.занятия *
        /// </summary>
        public int PracticalLess;
        /// <summary>
        /// 20. КСР
        /// </summary>
        public int KSR;
        /// <summary>
        /// 21. Контрольные работы и ПК
        /// </summary>
        public int TestPK;
        /// <summary>
        /// 22. Курс.проект/Курс.работа *
        /// </summary>
        public int CourseWork;
        /// <summary>
        /// 23. Преддипл.пратика *
        /// </summary>
        public int UndergraduatePract;
        /// <summary>
        /// 24. Подготовка ВКР
        /// </summary>
        public int PreparationVKR;
        /// <summary>
        /// 25. Учебная практика *
        /// </summary>
        public int TrainingPract;
        /// <summary>
        /// 26. Производственная практика *
        /// </summary>
        public int Internship;
        /// <summary>+
        /// 27. Государ.экзамен
        /// </summary>
        public int StateExam;
        /// <summary>
        /// 28.1 Фед.бюджет (ЗЕТ) ВНЕ СКОБОК.
        /// В БД: Hours
        /// </summary>
        public int FederalBudgetZET;
        /// <summary>
        /// 29.1 Всего (ЗЕТ) ВНЕ СКОБОК.
        /// В БД: EnteredHours
        /// </summary>
        public int TotalZET;
        /// <summary>
        /// 28.2 Фед.бюджет (ЗЕТ) В СКОБКАХ.
        /// В БД: HoursZ
        /// </summary>
        public int FederalBudgetZETskob;
        /// <summary>
        /// 29.2 Всего (ЗЕТ) В СКОБКАХ.
        /// В БД: EnteredHoursZ
        /// </summary>
        public int TotalZETskob;

        //Инициализация объекта
        public void Init(Word.Row row)
        {
            //Присылается строка таблицы Word из которой получаются ячейки
            Word.Cells cells = row.Cells;

            //У cellsTexts индексация с 0!!! (Мария Холод)
            //var cellsTexts = row.Cells.OfType<Word.Cell>().Select(cell => cell.Range.Text).ToList();

            //Обращение к тексту ячейки: cells[i].Range.Text;

            List<string> cellsTexts = row.Cells.OfType<Word.Cell>().Select(cell => cell.Range.Text).ToList();

            //У cells индексация с 1!!!
            for (int j = 1; j <= cells.Count; j++)
            {
                switch (j)
                {
                    //
                    case 1:
                        {
                            NumPP = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 2:
                        {
                            NumOUP = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 3:
                        {
                            Faculty = clsTextProcessor.WordCellToString(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 4:
                        {
                            Course = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 5:
                        {
                            NumOfWeeks = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 6:
                        {
                            ShortSpecialisation1 = clsTextProcessor.WordCellToString(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 7:
                        {
                            ShortSpecialisation2 = clsTextProcessor.WordCellToString(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 8:
                        {
                            Students = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 9:
                        {
                            StudentsFB = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 10:
                        {
                            Flow = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 11:
                        {
                            Groups = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 12:
                        {
                            Discipline = clsTextProcessor.WordCellToString(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 13:
                        {
                            LecturesConsult = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 14:
                        {
                            Exam = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 15:
                        {
                            Credit = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }

                    // !!!___???___
                    case 16:
                        {
                            HomeWork = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 17:
                        {
                            Consultations = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 18:
                        {
                            LaboratoryWorks = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 19:
                        {
                            PracticalLess = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 20:
                        {
                            KSR = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 21:
                        {
                            TestPK = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 22:
                        {
                            CourseWork = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }

                    //Тут числа?
                    case 23:
                        {
                            UndergraduatePract = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 24:
                        {
                            PreparationVKR = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 25:
                        {
                            TrainingPract = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 26:
                        {
                            Internship = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }

                    //Тут числа?
                    case 27:
                        {
                            StateExam = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 28:
                        {
                            FederalBudgetZET = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            FederalBudgetZETskob = clsTextProcessor.WordCellToSumSKOB(cells[j].Range.Text);
                            break;
                        }
                    //
                    case 29:
                        {
                            TotalZET = clsTextProcessor.WordCellToSum(cells[j].Range.Text);
                            TotalZETskob = clsTextProcessor.WordCellToSumSKOB(cells[j].Range.Text);
                            break;
                        }

                }
            }
        }
    }
}
