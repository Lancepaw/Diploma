using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace DataBaseIO
{
    public partial class frmPrepSwap : Form
    {
        public frmPrepSwap()
        {
            InitializeComponent();
        }

        //кнопка закрытия формы
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //группа действий при загрузке формы
        private void frmPrepSwap_Load(object sender, EventArgs e)
        {   
            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы преподавателей
            if (mdlData.colLecturer.Count > 0)
            {
                //Делаем доступными кнопки редактирования элементов
                //btnAdd.Enabled = true;
                //btnSave.Enabled = true;
                //btnDel.Enabled = true;
                //Заполняем комбо-боксы

                FillSemestrList();
                FillWorkYearList();
            }
            //при неудачной загрузке элементов из таблицы должностей
            else
            {
                //btnAdd.Enabled = false;
                //btnSave.Enabled = false;
                //btnDel.Enabled = false;
                //cmbDutyAddList.Enabled = false;
                //cmbDutyList.Enabled = false;
            }
        }

        private void FillSemestrList()
        {
            int NumFix = 0;
            NumFix = cmbSemestr.SelectedIndex;
            //Очищаем список
            cmbSemestr.Items.Clear();

            //Заполняем комбо-список семестрами
            for (int i = 0; i <= mdlData.colSemestr.Count - 1; i++)
            {
                cmbSemestr.Items.Add(mdlData.colSemestr[i].SemNum);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbSemestr.SelectedIndex = 0;
            }
            else
            {
                cmbSemestr.SelectedIndex = NumFix;
            }
        }

        private void FillWorkYearList()
        {
            int NumFix = 0;
            NumFix = cmbWorkYear.SelectedIndex;
            //Очищаем список
            cmbWorkYear.Items.Clear();

            //Заполняем комбо-список семестрами
            for (int i = 0; i <= mdlData.colWorkYear.Count - 1; i++)
            {
                cmbWorkYear.Items.Add(mdlData.colWorkYear[i].WorkYear);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbWorkYear.SelectedIndex = 0;
            }
            else
            {
                cmbWorkYear.SelectedIndex = NumFix;
            }
        }

        //Формирование графика замен на второй семестр
        private void btnSemestr2_Click(object sender, EventArgs e)
        {
            int curRow;

            string[] FIO;
            string Surname;
            string Name;
            string Patronymic;
            string Addition;

            //Очищаем сетку
            dgSwap.Rows.Clear();
            dgSwap.Columns.Clear();

            //Делаем невидимыми нуль-строку и нуль-столбец
            dgSwap.ColumnHeadersVisible = false;
            dgSwap.RowHeadersVisible = false;

            //Задаём количество столбцов
            //оно остаётся неизменным
            //в количестве 6 штук
            for (int i = 0; i <= 5; i++)
            {
                dgSwap.Columns.Add("", "");
            }

            //Создаём строки
            //1. под шапку таблицы

            dgSwap.Rows.Add();

            //Текущая строка первая
            curRow = 0;
            //Заполняем первую строку и одновременно формируем размерности
            dgSwap.Columns[0].Width = 30;
            dgSwap[0, curRow].Value = "№ п/п";

            dgSwap.Columns[1].Width = 350;
            dgSwap[1, curRow].Value = "Дисциплина";

            dgSwap.Columns[2].Width = 70;
            dgSwap[2, curRow].Value = "Группа";

            dgSwap.Columns[3].Width = 150;
            dgSwap[3, curRow].Value = "Основной преподаватель";

            dgSwap.Columns[4].Width = 150;
            dgSwap[4, curRow].Value = "Замещающий преподаватель";

            dgSwap.Columns[5].Width = 150;
            dgSwap[5, curRow].Value = "Резервный преподаватель";

            //Формируем строки по заменам преподавателей во втором семестре
            //прогоняем все строки фактической нагрузки
            for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
            {
                //смотрим только второй семестр
                if (mdlData.colDistribution[i].Semestr.SemNum.Equals("2 семестр"))
                {
                    //Если есть часы на ГАК, на аспирантуру, на диплом,
                    //на преддипломную практику, на производственную практику,
                    //на учебную практику, то не рассматриваем эти строки
                    if (!(mdlData.colDistribution[i].GAK > 0) & !(mdlData.colDistribution[i].PostGrad > 0) &
                        !(mdlData.colDistribution[i].DiplomaPaper > 0) & !(mdlData.colDistribution[i].PreDiplomaPractice > 0) &
                        !(mdlData.colDistribution[i].ProducingPractice > 0) & !(mdlData.colDistribution[i].TutorialPractice > 0))
                    {
                        //добавляем строку
                        dgSwap.Rows.Add();
                        //счётчик текущей строки увеличиваем на единицу
                        curRow += 1;
                        //дополнение к названию дисциплины
                        Addition = "(";
                        
                        //Если есть лекционные часы
                        if (mdlData.colDistribution[i].Lecture > 0)
                        {
                            //пишем про наличие лекции
                            Addition += "лк,";
                        }

                        //Если уже дописали вид нагрузки, то более ничего не пишем
                        if (!(Addition.EndsWith(",")))
                        {
                            //Если есть практические часы
                            if (mdlData.colDistribution[i].Practice > 0)
                            {
                                //пишем про наличие практических
                                Addition += "пр,";
                            }
                        }

                        //Если уже дописали вид нагрузки, то более ничего не пишем
                        if (!(Addition.EndsWith(",")))
                        {
                            //Если есть лабораторные часы
                            if (mdlData.colDistribution[i].LabWork > 0)
                            {
                                //пишем про наличие лабораторных
                                Addition += "лб,";
                            }
                        }

                        //Если уже дописали вид нагрузки, то более ничего не пишем
                        if (!(Addition.EndsWith(",")))
                        {
                            //Если есть курсовой проект
                            if (mdlData.colDistribution[i].KursProject > 0)
                            {
                                //пишем про наличие курсового проекта
                                Addition += "к/пр,";
                            }
                        }
                        
                        //Убираем запятую
                        if (Addition.EndsWith(","))
                        {
                            Addition = Addition.Substring(0, Addition.Length - 1);
                        }

                        //Закрываем скобку
                        Addition += ")";

                        //Если внутри скобок пустота, то
                        if (Addition == "()")
                        {
                            //убираем скобки
                            Addition = "";
                        }

                        //В номер по порядку вписываем значение счётчика
                        dgSwap[0, curRow].Value = curRow.ToString();
                        //Вписываем название дисциплины с дополнением
                        dgSwap[1, curRow].Value = mdlData.colDistribution[i].Subject.Subject.ToString() + " " + Addition;
                        //Название группы с номером курса
                        if (mdlData.colDistribution[i].KursNum != null)
                        {
                            dgSwap[2, curRow].Value = mdlData.colDistribution[i].Speciality.ShortInstitute.ToString() + "-" +
                                                      mdlData.colDistribution[i].KursNum.Kurs.ToString();
                        }

                        //Разбираем строку для вывода отдельно
                        //Фамилии, имени и отчества основного преподавателя
                        if (mdlData.colDistribution[i].Lecturer != null)
                        {
                            FIO = mdlData.colDistribution[i].Lecturer.FIO.Split(new char[] { ' ' });

                            if (FIO.GetLength(0) == 3)
                            {
                                Surname = FIO[0];
                                Name = FIO[1].Substring(0, 1) + ".";
                                Patronymic = FIO[2].Substring(0, 1) + ".";
                            }
                            else if (FIO.GetLength(0) == 2)
                            {
                                Surname = FIO[0];
                                Name = FIO[1].Substring(0, 1) + ".";
                                Patronymic = "";
                            }
                            else if (FIO.GetLength(0) == 1)
                            {
                                Surname = FIO[0];
                                Name = "";
                                Patronymic = "";
                            }
                            else
                            {
                                Surname = "";
                                Name = "";
                                Patronymic = "";
                            }

                            dgSwap[3, curRow].Value = Surname + " " + Name + Patronymic;
                        }

                        //Разбираем строку для вывода отдельно
                        //Фамилии, имени и отчества заменяющего преподавателя
                        if (mdlData.colDistribution[i].Lecturer2 != null)
                        {
                            FIO = mdlData.colDistribution[i].Lecturer2.FIO.Split(new char[] { ' ' });

                            if (FIO.GetLength(0) == 3)
                            {
                                Surname = FIO[0];
                                Name = FIO[1].Substring(0, 1) + ".";
                                Patronymic = FIO[2].Substring(0, 1) + ".";
                            }
                            else if (FIO.GetLength(0) == 2)
                            {
                                Surname = FIO[0];
                                Name = FIO[1].Substring(0, 1) + ".";
                                Patronymic = "";
                            }
                            else if (FIO.GetLength(0) == 1)
                            {
                                Surname = FIO[0];
                                Name = "";
                                Patronymic = "";
                            }
                            else
                            {
                                Surname = "";
                                Name = "";
                                Patronymic = "";
                            }

                            dgSwap[4, curRow].Value = Surname + " " + Name + Patronymic;
                        }

                        //Разбираем строку для вывода отдельно
                        //Фамилии, имени и отчества резервного преподавателя
                        if (mdlData.colDistribution[i].Lecturer3 != null)
                        {
                            FIO = mdlData.colDistribution[i].Lecturer3.FIO.Split(new char[] { ' ' });

                            if (FIO.GetLength(0) == 3)
                            {
                                Surname = FIO[0];
                                Name = FIO[1].Substring(0, 1) + ".";
                                Patronymic = FIO[2].Substring(0, 1) + ".";
                            }
                            else if (FIO.GetLength(0) == 2)
                            {
                                Surname = FIO[0];
                                Name = FIO[1].Substring(0, 1) + ".";
                                Patronymic = "";
                            }
                            else if (FIO.GetLength(0) == 1)
                            {
                                Surname = FIO[0];
                                Name = "";
                                Patronymic = "";
                            }
                            else
                            {
                                Surname = "";
                                Name = "";
                                Patronymic = "";
                            }

                            dgSwap[5, curRow].Value = Surname + " " + Name + Patronymic;
                        }
                    }
                }
            }

        }

        //Форитрование графика замен на первый семестр
        private void btnSemestr1_Click(object sender, EventArgs e)
        {
            int curRow;

            string[] FIO;
            string Surname;
            string Name;
            string Patronymic;
            string Addition;

            //Очищаем сетку
            dgSwap.Rows.Clear();
            dgSwap.Columns.Clear();

            //Делаем невидимыми нуль-строку и нуль-столбец
            dgSwap.ColumnHeadersVisible = false;
            dgSwap.RowHeadersVisible = false;

            //Задаём количество столбцов
            //оно остаётся неизменным
            //в количестве 6 штук
            for (int i = 0; i <= 5; i++)
            {
                dgSwap.Columns.Add("", "");
            }

            //Создаём строки
            //1. под шапку таблицы

            dgSwap.Rows.Add();

            //Текущая строка первая
            curRow = 0;
            //Заполняем первую строку и одновременно формируем размерности
            dgSwap.Columns[0].Width = 30;
            dgSwap[0, curRow].Value = "№ п/п";

            dgSwap.Columns[1].Width = 350;
            dgSwap[1, curRow].Value = "Дисциплина";

            dgSwap.Columns[2].Width = 70;
            dgSwap[2, curRow].Value = "Группа";

            dgSwap.Columns[3].Width = 150;
            dgSwap[3, curRow].Value = "Основной преподаватель";

            dgSwap.Columns[4].Width = 150;
            dgSwap[4, curRow].Value = "Замещающий преподаватель";

            dgSwap.Columns[5].Width = 150;
            dgSwap[5, curRow].Value = "Резервный преподаватель";

            //Формируем строки по заменам преподавателей в первом семестре
            //прогоняем все строки фактической нагрузки
            for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
            {
                //смотрим только второй семестр
                if (mdlData.colDistribution[i].Semestr.SemNum.Equals("1 семестр"))
                {
                    //Если есть часы на ГАК, на аспирантуру, на диплом,
                    //на преддипломную практику, на производственную практику,
                    //на посещение учебных занятий,
                    //на учебную практику, то не рассматриваем эти строки
                    if (!(mdlData.colDistribution[i].GAK > 0) & !(mdlData.colDistribution[i].PostGrad > 0) &
                        !(mdlData.colDistribution[i].DiplomaPaper > 0) & !(mdlData.colDistribution[i].PreDiplomaPractice > 0) &
                        !(mdlData.colDistribution[i].ProducingPractice > 0) & !(mdlData.colDistribution[i].TutorialPractice > 0) &
                        !(mdlData.colDistribution[i].Visiting > 0))
                    {

                        if ((mdlData.colDistribution[i].Subject.Subject == "Посещение занятий") ||
                            (mdlData.colDistribution[i].Subject.Subject == "Аспирантура"))
                        {
                            continue;
                        }
                        //добавляем строку
                        dgSwap.Rows.Add();
                        //счётчик текущей строки увеличиваем на единицу
                        curRow += 1;
                        //дополнение к названию дисциплины
                        Addition = "(";

                        //Если есть лекционные часы
                        if (mdlData.colDistribution[i].Lecture > 0)
                        {
                            //пишем про наличие лекции
                            Addition += "лк,";
                        }

                        //Если уже дописали вид нагрузки, то более ничего не пишем
                        if (!(Addition.EndsWith(",")))
                        {
                            //Если есть практические часы
                            if (mdlData.colDistribution[i].Practice > 0)
                            {
                                //пишем про наличие практических
                                Addition += "пр,";
                            }
                        }

                        //Если уже дописали вид нагрузки, то более ничего не пишем
                        if (!(Addition.EndsWith(",")))
                        {
                            //Если есть лабораторные часы
                            if (mdlData.colDistribution[i].LabWork > 0)
                            {
                                //пишем про наличие лабораторных
                                Addition += "лб,";
                            }
                        }

                        //Если уже дописали вид нагрузки, то более ничего не пишем
                        if (!(Addition.EndsWith(",")))
                        {
                            //Если есть курсовой проект
                            if (mdlData.colDistribution[i].KursProject > 0)
                            {
                                //пишем про наличие курсового проекта
                                Addition += "к/пр,";
                            }
                        }

                        //Убираем запятую
                        if (Addition.EndsWith(","))
                        {
                            Addition = Addition.Substring(0, Addition.Length - 1);
                        }

                        //Закрываем скобку
                        Addition += ")";

                        //Если внутри скобок пустота, то
                        if (Addition == "()")
                        {
                            //убираем скобки
                            Addition = "";
                        }

                        //В номер по порядку вписываем значение счётчика
                        dgSwap[0, curRow].Value = curRow.ToString();
                        //Вписываем название дисциплины с дополнением
                        dgSwap[1, curRow].Value = mdlData.colDistribution[i].Subject.Subject.ToString() + " " + Addition;
                        //Название группы с номером курса

                        if (!(mdlData.colDistribution[i].Speciality == null) & !(mdlData.colDistribution[i].KursNum == null))
                        {
                            dgSwap[2, curRow].Value = mdlData.colDistribution[i].Speciality.ShortInstitute.ToString() + "-" +
                                                      mdlData.colDistribution[i].KursNum.Kurs.ToString();
                        }
                        else
                        {
                            if (!(mdlData.colDistribution[i].Speciality == null))
                            {
                                dgSwap[2, curRow].Value = mdlData.colDistribution[i].Speciality.ShortInstitute.ToString();
                            }
                            else
                            {
                                if (!(mdlData.colDistribution[i].KursNum == null))
                                {
                                    dgSwap[2, curRow].Value = "???-" +
                                        mdlData.colDistribution[i].KursNum.Kurs.ToString();
                                }
                            }
                        }

                        //Разбираем строку для вывода отдельно
                        //Фамилии, имени и отчества основного преподавателя
                        FIO = mdlData.colDistribution[i].Lecturer.FIO.Split(new char[] { ' ' });

                        if (FIO.GetLength(0) == 3)
                        {
                            Surname = FIO[0];
                            Name = FIO[1].Substring(0, 1) + ".";
                            Patronymic = FIO[2].Substring(0, 1) + ".";
                        }
                        else if (FIO.GetLength(0) == 2)
                        {
                            Surname = FIO[0];
                            Name = FIO[1].Substring(0, 1) + ".";
                            Patronymic = "";
                        }
                        else if (FIO.GetLength(0) == 1)
                        {
                            Surname = FIO[0];
                            Name = "";
                            Patronymic = "";
                        }
                        else
                        {
                            Surname = "";
                            Name = "";
                            Patronymic = "";
                        }

                        dgSwap[3, curRow].Value = Surname + " " + Name + Patronymic;

                        //Разбираем строку для вывода отдельно
                        //Фамилии, имени и отчества заменяющего преподавателя
                        if (!(mdlData.colDistribution[i].Lecturer2 == null))
                        {
                            FIO = mdlData.colDistribution[i].Lecturer2.FIO.Split(new char[] { ' ' });

                            if (FIO.GetLength(0) == 3)
                            {
                                Surname = FIO[0];
                                Name = FIO[1].Substring(0, 1) + ".";
                                Patronymic = FIO[2].Substring(0, 1) + ".";
                            }
                            else if (FIO.GetLength(0) == 2)
                            {
                                Surname = FIO[0];
                                Name = FIO[1].Substring(0, 1) + ".";
                                Patronymic = "";
                            }
                            else if (FIO.GetLength(0) == 1)
                            {
                                Surname = FIO[0];
                                Name = "";
                                Patronymic = "";
                            }
                            else
                            {
                                Surname = "";
                                Name = "";
                                Patronymic = "";
                            }

                            dgSwap[4, curRow].Value = Surname + " " + Name + Patronymic;
                        }

                        //Разбираем строку для вывода отдельно
                        //Фамилии, имени и отчества резервного преподавателя
                        if (!(mdlData.colDistribution[i].Lecturer3 == null))
                        {
                            FIO = mdlData.colDistribution[i].Lecturer3.FIO.Split(new char[] { ' ' });

                            if (FIO.GetLength(0) == 3)
                            {
                                Surname = FIO[0];
                                Name = FIO[1].Substring(0, 1) + ".";
                                Patronymic = FIO[2].Substring(0, 1) + ".";
                            }
                            else if (FIO.GetLength(0) == 2)
                            {
                                Surname = FIO[0];
                                Name = FIO[1].Substring(0, 1) + ".";
                                Patronymic = "";
                            }
                            else if (FIO.GetLength(0) == 1)
                            {
                                Surname = FIO[0];
                                Name = "";
                                Patronymic = "";
                            }
                            else
                            {
                                Surname = "";
                                Name = "";
                                Patronymic = "";
                            }

                            dgSwap[5, curRow].Value = Surname + " " + Name + Patronymic;
                        }
                    }
                }
            }
        }

        private void btnWord_Click(object sender, EventArgs e)
        {  
            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;

            if (cmbSemestr.SelectedIndex > 0)
            {
                try
                {
                    //Создаём новое Word приложение
                    Word._Application ObjWord = new Word.Application();
                    //Добавляем новый чистый документ Word
                    Word._Document ObjDoc = ObjWord.Application.Documents.Add();
                    ObjDoc.Activate();

                    cfgPage(ObjDoc);

                    textHeader(ObjMissing, ObjDoc);

                    tableMain(ObjMissing, ObjDoc);

                    textSignature(ObjMissing, ObjDoc);

                    ObjWord.Visible = true;

                    ObjDoc.SaveAs(Application.StartupPath + @"\" + "График замен на " + cmbSemestr.SelectedItem.ToString() + " " +
                        DateTime.Now.Date.ToString("yyyyMMdd") + " " +
                        DateTime.Now.TimeOfDay.ToString("hhmmss") + ".docx");
                    ObjDoc.Close();
                    ObjWord.Quit();
                }
                catch
                {
                    MessageBox.Show("Возможно на этом компьютере присутствует проблема совместимости с MS Word." +
                    " Попробуйте установить версию 2007 и выше.");
                }
            }
        }

        private void textHeader(object ObjMissing, Word._Document ObjDoc)
        {
            Word.Paragraph ObjParagraph;

            //Добавляем абзац текста в начало документа
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            //Текстуальное содержимое абзаца
            ObjParagraph.Range.Text = "График замен преподавателей по болезни на " + cmbSemestr.SelectedItem.ToString() + " " +
                                      cmbWorkYear.SelectedItem.ToString() + " уч.года ";
            //Размер шрифта 14 пт
            ObjParagraph.Range.Font.Size = 14;
            //Times New Roman
            ObjParagraph.Range.Font.Name = "Times New Roman";
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Отступ в 10 пт после абзаца
            ObjParagraph.Format.SpaceAfter = 10;
            //Отступ в 0 пт до абзаца
            ObjParagraph.Format.SpaceBefore = 0;
            //Одинарный межстрочный интервал
            ObjParagraph.Format.Space1();
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textSignature(object ObjMissing, Word._Document ObjDoc)
        {
            Word.Paragraph ObjParagraph;

            //Добавляем абзац текста в начало документа
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            //Текстуальное содержимое абзаца
            ObjParagraph.Range.Text = "Зав. кафедрой УиЗИ, профессор                                                      Л.А. Баранов";
            //Размер шрифта 12 пт
            ObjParagraph.Range.Font.Size = 12;
            //Times New Roman
            ObjParagraph.Range.Font.Name = "Times New Roman";
            //Выравнивание по левому краю
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //Отступ в 0 пт до абзаца
            ObjParagraph.Format.SpaceBefore = 10;
            //Одинарный межстрочный интервал
            ObjParagraph.Format.Space1();
        }

        private void tableMain(object ObjMissing, Word._Document ObjDoc)
        {
            int curRow;

            int countCol;
            int countRow;

            string[] FIO;
            string Surname;
            string Name;
            string Patronymic;
            string Addition;

            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";

            Word.Table ObjTable;
            Word.Range ObjWordRange;

            //Задаём количество столбцов
            //оно остаётся неизменным
            //в количестве 6 штук
            countCol = 6;
            //Создаём строки
            //1. под шапку таблицы
            countRow = 1;

            //Формируем строки по заменам преподавателей в первом семестре
            //прогоняем все строки фактической нагрузки
            for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
            {
                //смотрим только второй семестр
                if (mdlData.colDistribution[i].Semestr.SemNum.Equals(cmbSemestr.SelectedItem.ToString()))
                {
                    //Если есть часы на ГАК, на аспирантуру, на диплом,
                    //на преддипломную практику, на производственную практику,
                    //на посещение учебных занятий,
                    //на учебную практику, то не рассматриваем эти строки
                    if (!(mdlData.colDistribution[i].GAK > 0) & !(mdlData.colDistribution[i].PostGrad > 0) &
                        !(mdlData.colDistribution[i].DiplomaPaper > 0) & !(mdlData.colDistribution[i].PreDiplomaPractice > 0) &
                        !(mdlData.colDistribution[i].ProducingPractice > 0) & !(mdlData.colDistribution[i].TutorialPractice > 0) &
                        !(mdlData.colDistribution[i].Visiting > 0))
                    {

                        if ((mdlData.colDistribution[i].Subject.Subject == "Посещение занятий") ||
                            (mdlData.colDistribution[i].Subject.Subject == "Аспирантура") ||
                            (mdlData.colDistribution[i].Subject.Subject == "Руководство магистрами"))
                        {
                            continue;
                        }
                        //добавляем строку
                        countRow++;
                    }
                }
            }

            //Вставляем таблицу согласно заполненной сетке и заполняем её данными о нагрузке
            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjTable = ObjDoc.Tables.Add(ObjWordRange, countRow, countCol, ref ObjMissing, ref ObjMissing);

            //Размер шрифта 10 пт
            ObjTable.Range.Font.Size = 10;
            //Выравнивание по левому краю
            ObjTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //Отступ после абзаца отсутствует
            ObjTable.Range.ParagraphFormat.SpaceAfter = 0;
            //Отступ в 0 пт до абзаца
            ObjTable.Range.ParagraphFormat.SpaceBefore = 0;
            //Одинарный межстрочный интервал
            ObjTable.Range.ParagraphFormat.Space1();
            //Границы таблицы включены
            ObjTable.Borders.Enable = 1;

            //Текущая строка первая
            curRow = 1;
            ObjTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            ObjTable.Rows[1].Range.Font.Bold = 1;
            //Заполняем первую строку и одновременно формируем размерности
            ObjTable.Cell(curRow, 1).Range.Text = "№ п/п";

            ObjTable.Cell(curRow, 2).Range.Text = "Дисциплина";

            ObjTable.Cell(curRow, 3).Range.Text = "Группа";

            ObjTable.Cell(curRow, 4).Range.Text = "Основной преподаватель";

            ObjTable.Cell(curRow, 5).Range.Text = "Замещающий преподаватель";

            ObjTable.Cell(curRow, 6).Range.Text = "Резервный преподаватель";

            //Формируем строки по заменам преподавателей в первом семестре
            //прогоняем все строки фактической нагрузки
            for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
            {
                //смотрим только второй семестр
                if (mdlData.colDistribution[i].Semestr.SemNum.Equals(cmbSemestr.SelectedItem.ToString()))
                {
                    //Если есть часы на ГАК, на аспирантуру, на диплом,
                    //на преддипломную практику, на производственную практику,
                    //на посещение учебных занятий,
                    //на учебную практику, то не рассматриваем эти строки
                    if (!(mdlData.colDistribution[i].GAK > 0) & !(mdlData.colDistribution[i].PostGrad > 0) &
                        !(mdlData.colDistribution[i].DiplomaPaper > 0) & !(mdlData.colDistribution[i].PreDiplomaPractice > 0) &
                        !(mdlData.colDistribution[i].ProducingPractice > 0) & !(mdlData.colDistribution[i].TutorialPractice > 0) &
                        !(mdlData.colDistribution[i].Visiting > 0))
                    {

                        if ((mdlData.colDistribution[i].Subject.Subject == "Посещение занятий") ||
                            (mdlData.colDistribution[i].Subject.Subject == "Аспирантура") ||
                            (mdlData.colDistribution[i].Subject.Subject == "Руководство магистрами"))
                        {
                            continue;
                        }
                        //счётчик текущей строки увеличиваем на единицу
                        curRow += 1;
                        //дополнение к названию дисциплины
                        Addition = "(";

                        //Если есть лекционные часы
                        if (mdlData.colDistribution[i].Lecture > 0)
                        {
                            //пишем про наличие лекции
                            Addition += "лк,";
                        }

                        //Если уже дописали вид нагрузки, то более ничего не пишем
                        if (!(Addition.EndsWith(",")))
                        {
                            //Если есть практические часы
                            if (mdlData.colDistribution[i].Practice > 0)
                            {
                                //пишем про наличие практических
                                Addition += "пр,";
                            }
                        }

                        //Если уже дописали вид нагрузки, то более ничего не пишем
                        if (!(Addition.EndsWith(",")))
                        {
                            //Если есть лабораторные часы
                            if (mdlData.colDistribution[i].LabWork > 0)
                            {
                                //пишем про наличие лабораторных
                                Addition += "лб,";
                            }
                        }

                        //Если уже дописали вид нагрузки, то более ничего не пишем
                        if (!(Addition.EndsWith(",")))
                        {
                            //Если есть курсовой проект
                            if (mdlData.colDistribution[i].KursProject > 0)
                            {
                                //пишем про наличие курсового проекта
                                Addition += "к/пр,";
                            }
                        }

                        //Убираем запятую
                        if (Addition.EndsWith(","))
                        {
                            Addition = Addition.Substring(0, Addition.Length - 1);
                        }

                        //Закрываем скобку
                        Addition += ")";

                        //Если внутри скобок пустота, то
                        if (Addition == "()")
                        {
                            //убираем скобки
                            Addition = "";
                        }

                        //В номер по порядку вписываем значение счётчика
                        ObjTable.Cell(curRow, 1).Range.Text = (curRow - 1).ToString();
                        //Вписываем название дисциплины с дополнением
                        ObjTable.Cell(curRow, 2).Range.Text = mdlData.colDistribution[i].Subject.Subject.ToString() + " " + Addition;
                        //Название группы с номером курса

                        if (!(mdlData.colDistribution[i].Speciality == null) & !(mdlData.colDistribution[i].KursNum == null))
                        {
                            ObjTable.Cell(curRow, 3).Range.Text = mdlData.colDistribution[i].Speciality.ShortInstitute.ToString() + "-" +
                                                    mdlData.colDistribution[i].KursNum.Kurs.ToString();
                        }
                        else
                        {
                            if (!(mdlData.colDistribution[i].Speciality == null))
                            {
                                ObjTable.Cell(curRow, 3).Range.Text = mdlData.colDistribution[i].Speciality.ShortInstitute.ToString();
                            }
                            else
                            {
                                if (!(mdlData.colDistribution[i].KursNum == null))
                                {
                                    ObjTable.Cell(curRow, 3).Range.Text = "???-" +
                                                    mdlData.colDistribution[i].KursNum.Kurs.ToString();
                                }
                            }
                        }

                        //Разбираем строку для вывода отдельно
                        //Фамилии, имени и отчества основного преподавателя
                        if (mdlData.colDistribution[i].Lecturer != null)
                        {
                            FIO = mdlData.colDistribution[i].Lecturer.FIO.Split(new char[] { ' ' });

                            if (FIO.GetLength(0) == 3)
                            {
                                Surname = FIO[0];
                                Name = FIO[1].Substring(0, 1) + ".";
                                Patronymic = FIO[2].Substring(0, 1) + ".";
                            }
                            else if (FIO.GetLength(0) == 2)
                            {
                                Surname = FIO[0];
                                Name = FIO[1].Substring(0, 1) + ".";
                                Patronymic = "";
                            }
                            else if (FIO.GetLength(0) == 1)
                            {
                                Surname = FIO[0];
                                Name = "";
                                Patronymic = "";
                            }
                            else
                            {
                                Surname = "";
                                Name = "";
                                Patronymic = "";
                            }

                            ObjTable.Cell(curRow, 4).Range.Text = Surname + " " + Name + Patronymic;
                        }

                        //Разбираем строку для вывода отдельно
                        //Фамилии, имени и отчества заменяющего преподавателя
                        if (!(mdlData.colDistribution[i].Lecturer2 == null))
                        {
                            FIO = mdlData.colDistribution[i].Lecturer2.FIO.Split(new char[] { ' ' });

                            if (FIO.GetLength(0) == 3)
                            {
                                Surname = FIO[0];
                                Name = FIO[1].Substring(0, 1) + ".";
                                Patronymic = FIO[2].Substring(0, 1) + ".";
                            }
                            else if (FIO.GetLength(0) == 2)
                            {
                                Surname = FIO[0];
                                Name = FIO[1].Substring(0, 1) + ".";
                                Patronymic = "";
                            }
                            else if (FIO.GetLength(0) == 1)
                            {
                                Surname = FIO[0];
                                Name = "";
                                Patronymic = "";
                            }
                            else
                            {
                                Surname = "";
                                Name = "";
                                Patronymic = "";
                            }

                            ObjTable.Cell(curRow, 5).Range.Text = Surname + " " + Name + Patronymic;
                        }

                        //Разбираем строку для вывода отдельно
                        //Фамилии, имени и отчества резервного преподавателя
                        if (!(mdlData.colDistribution[i].Lecturer3 == null))
                        {
                            FIO = mdlData.colDistribution[i].Lecturer3.FIO.Split(new char[] { ' ' });

                            if (FIO.GetLength(0) == 3)
                            {
                                Surname = FIO[0];
                                Name = FIO[1].Substring(0, 1) + ".";
                                Patronymic = FIO[2].Substring(0, 1) + ".";
                            }
                            else if (FIO.GetLength(0) == 2)
                            {
                                Surname = FIO[0];
                                Name = FIO[1].Substring(0, 1) + ".";
                                Patronymic = "";
                            }
                            else if (FIO.GetLength(0) == 1)
                            {
                                Surname = FIO[0];
                                Name = "";
                                Patronymic = "";
                            }
                            else
                            {
                                Surname = "";
                                Name = "";
                                Patronymic = "";
                            }

                            ObjTable.Cell(curRow, 6).Range.Text = Surname + " " + Name + Patronymic;
                        }
                    }
                }
            }

            ObjTable.Columns[1].Width = 0.94f / 0.03527f;
            ObjTable.Columns[2].Width = 7.05f / 0.03527f;
            ObjTable.Columns[3].Width = 1.70f / 0.03527f;
            ObjTable.Columns[4].Width = 3.09f / 0.03527f;
            ObjTable.Columns[5].Width = 3.62f / 0.03527f;
            ObjTable.Columns[6].Width = 3.43f / 0.03527f;
        }

        private void cfgPage(Word._Document ObjDoc)
        {
            ObjDoc.PageSetup.LeftMargin = 0.5f / 0.03527f;
            ObjDoc.PageSetup.RightMargin = 0.5f / 0.03527f;
            ObjDoc.PageSetup.TopMargin = 0.5f / 0.03527f;
            ObjDoc.PageSetup.BottomMargin = 0.5f / 0.03527f;
        }
    }
}