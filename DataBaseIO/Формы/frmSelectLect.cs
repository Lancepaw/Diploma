using System;
using System.Reflection;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;

namespace DataBaseIO
{
    public partial class frmSelectLect : Form
    {
        public frmSelectLect()
        {
            InitializeComponent();
            cmbBoxFill();
        }

        public void cmbBoxFill()
        {
            for (int i = 0; i < mdlData.colLecturer.Count - 1; i++)
            {
                cmbChoose.Items.Add(mdlData.colLecturer[i].FIO);
            }
        }

        private void onVisioTimeTableSimple()
        {
            string visDocName = Application.StartupPath + "\\myVisio.vsd";
            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;
            clsLecturer L;
            clsSchedule Sch1 = new clsSchedule();
            clsSchedule Sch2 = new clsSchedule();
            bool flgFound1Week;
            bool flgFound2Week;

            float currY = 0f;
            float currX = 0f;
            float XStart = 0f;

            float XLeft = 10f;
            float XRight = 0f;
            float YBottom = 10f;
            float YTop = 0f;

            int Semestr = 0;
            bool trigColor;

            Visio.Shape visTextBox;
            Visio.Page visPage;
            Visio.Application visApp;
            Visio.Document visDoc;

            visApp = new Visio.Application();
            visDoc = visApp.Documents.Add("");
            visPage = visApp.ActivePage;


            //Вопрос: для какого семестра создаётся расписание?
            if (MessageBox.Show(this, "Создаётся расписание для I семестра?", "Выбор семестра", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Semestr = 1;
            }
            else
            {
                Semestr = 2;
            }

            //---------------С этого момента создаётся пустая таблица----------
            //
            currY = 10f;
            trigColor = false;
            //Получаем номер преподавателя в коллекции
            int lectID = new int();
            for (int i = mdlData.colLecturer.Count - 1; i >= 0; i--)
            {
                if (mdlData.colLecturer[i].FIO == cmbChoose.SelectedItem.ToString())
                {
                    lectID = i;
                }
            }

            currX = 10f;
            //if (mdlData.colLecturer[i].Rate > 0 || DetectCheckedSchedule(mdlData.colLecturer[i], Semestr))
            if (DetectCheckedSchedule(mdlData.colLecturer[lectID], Semestr))
            {
                //Рисуем прямоугольник под Ф.И.О.
                visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                   ((currX + 80f) / 25.4f), ((currY + 20f) / 25.4f));
                //Вписываем Ф.И.О. рассматриваемого преподавателя
                visTextBox.Text = mdlData.colLecturer[lectID].FIO;
                visTextBox.CellsU["Char.Size"].FormulaForceU = "18 pt";

                if (trigColor)
                {
                    visTextBox.CellsU["FillForegnd"].FormulaForceU = "RGB(235,235,235)";
                }
                else
                {
                    visTextBox.CellsU["FillForegnd"].FormulaForceU = "RGB(255,255,255)";
                }

                currX += 80f;

                //Запускаем цикл по дням недели
                for (int j = 0; j < 5; j++)
                {
                    //Запускаем цикл по временам пар
                    //до обеда
                    for (int k = 0; k < 3; k++)
                    {
                        visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                    ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                        if (trigColor)
                        {
                            visTextBox.CellsU["FillForegnd"].FormulaForceU = "RGB(235,235,235)";
                        }
                        else
                        {
                            visTextBox.CellsU["FillForegnd"].FormulaForceU = "RGB(255,255,255)";
                        }

                        currX += 20f;
                    }

                    //Обед
                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                            ((currX + 10f) / 25.4f), ((currY + 20f) / 25.4f));

                    if (trigColor)
                    {
                        visTextBox.CellsU["FillForegnd"].FormulaForceU = "RGB(235,235,235)";
                    }
                    else
                    {
                        visTextBox.CellsU["FillForegnd"].FormulaForceU = "RGB(255,255,255)";
                    }

                    currX += 10f;

                    //Запускаем цикл по временам пар
                    //после обеда
                    for (int k = 3; k < 8; k++)
                    {
                        visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                    ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                        if (trigColor)
                        {
                            visTextBox.CellsU["FillForegnd"].FormulaForceU = "RGB(235,235,235)";
                        }
                        else
                        {
                            visTextBox.CellsU["FillForegnd"].FormulaForceU = "RGB(255,255,255)";
                        }

                        currX += 20f;
                    }


                }

                trigColor = !trigColor;
                currY += 20f;


                currX = 10f;
                //Рисуем прямоугольник под надпись Ф.И.О.
                visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                   ((currX + 80f) / 25.4f), ((currY + 40f) / 25.4f));
                //Вписываем Ф.И.О. рассматриваемого преподавателя
                visTextBox.Text = "Преподаватель";
                visTextBox.CellsU["Char.Size"].FormulaForceU = "24 pt";

                currX += 80f;

                //Запускаем цикл по дням недели
                for (int j = 0; j < 5; j++)
                {
                    XStart = currX;
                    for (int k = 0; k < 3; k++)
                    {
                        visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                    ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));
                        if (mdlData.colPairTime.Count > 0)
                        {
                            visTextBox.Text = mdlData.colPairTime[k].Time;
                        }
                        else
                        {
                            visTextBox.Text = "--:-- - --:--";
                        }
                        currX += 20f;
                    }

                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                            ((currX + 20f) / 25.4f), ((currY + 10f) / 25.4f));
                    visTextBox.Text = "Обед";
                    visTextBox.Rotate90();

                    visApp.ActiveWindow.DeselectAll();
                    visApp.ActiveWindow.Select(visTextBox, 2);
                    visApp.ActiveWindow.Selection.Move(-(5f / 25.4f), (5f / 25.4f));

                    currX += 10f;

                    for (int k = 3; k < 8; k++)
                    {
                        visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                    ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));
                        if (mdlData.colPairTime.Count > 0)
                        {
                            visTextBox.Text = mdlData.colPairTime[k].Time;
                        }
                        else
                        {
                            visTextBox.Text = "--:-- - --:--";
                        }
                        currX += 20f;
                    }

                    visTextBox = visPage.DrawRectangle((XStart / 25.4f), ((currY + 20f) / 25.4f),
                                 ((currX) / 25.4f), ((currY + 40f) / 25.4f));

                    if (mdlData.colWeekDays.Count > 0)
                    {
                        visTextBox.Text = mdlData.colWeekDays[j].WeekDay;
                    }
                    else
                    {
                        visTextBox.Text = "№ " + (j + 1).ToString();
                    }

                    visTextBox.CellsU["Char.Size"].FormulaForceU = "24 pt";
                }

                XRight = currX;
                YTop = currY + 40f;
                YBottom = currY + 60f;

                //Изображаем прямоугольник под текст с заголовком таблицы
                visTextBox = visPage.DrawRectangle((XLeft / 25.4f), (YTop / 25.4f),
                                                         (XRight / 25.4f), ((YBottom) / 25.4f));
                visTextBox.LineStyle = "None";
                visTextBox.FillStyle = "None";

                //Снять выделение со всех элементов
                visApp.ActiveWindow.DeselectAll();
                visApp.ActiveWindow.Select(visTextBox, 2);

                visTextBox.Text = "Расписание преподавателей кафедры \"" + mdlData.DepartmentName +
                    "\" на " + (Semestr == 1 ? "I" : "II") + " семестр " +
                    mdlData.colWorkYear[mdlData.colWorkYear.Count - 2].WorkYear + " учебного года";

                //Пример покраски текста в форме (shape) в красный цвет
                //visTextBox.CellsU["Char.Color"].FormulaForceU = "RGB(255,0,0)";
                visTextBox.CellsU["Char.Size"].FormulaForceU = "48 pt";


                //---------------До этого момента создаётся пустая таблица---------

                //---------------С этого момента заполняется таблица---------------
                //
                currY = 10f;

                currX = 10f;

                //if (mdlData.colLecturer[i].Rate > 0 || DetectCheckedSchedule(mdlData.colLecturer[i], Semestr))
                if (DetectCheckedSchedule(mdlData.colLecturer[lectID], Semestr))
                {
                    L = mdlData.colLecturer[lectID];

                    currX += 80f;

                    //Запускаем цикл по дням недели
                    for (int j = 0; j < 5; j++)
                    {
                        //запускаем цикл по временам занятий
                        //до обеда
                        for (int k = 0; k < 3; k++)
                        {
                            Sch1 = null;
                            Sch2 = null;
                            flgFound1Week = DetectTimeTableElement(ref Sch1, L, Semestr, 0, j, k);
                            flgFound2Week = DetectTimeTableElement(ref Sch2, L, Semestr, 1, j, k);

                            //Если по каждой неделе нашёлся элемент расписания
                            if (flgFound1Week & flgFound2Week)
                            {
                                //Если хотя бы что-то отсутствует у одного из элементов расписания, то
                                //выводить надпись "Есть"
                                if ((Sch1.KursNum == null || Sch1.Spec == null || Sch1.SubjectType == null) ||
                                    (Sch2.KursNum == null || Sch2.Spec == null || Sch2.SubjectType == null))
                                {
                                    //То просто пишем, что обе пары в это время есть
                                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                    visTextBox.LineStyle = "None";
                                    visTextBox.FillStyle = "None";
                                    visTextBox.Text = "Есть";
                                }
                                //Если какие-то сведения имеются
                                else
                                {
                                    //Если всё в точности совпадает, то выводим горизонтально
                                    //первый попавшийся элемент
                                    if (Sch1.SubjectType.Equals(Sch2.SubjectType) &
                                        Sch1.Spec.Equals(Sch2.Spec) &
                                        Sch1.KursNum.Equals(Sch2.KursNum) &
                                        Sch1.Auditory.Equals(Sch2.Auditory) &
                                        Sch1.Group.Equals(Sch2.Group) &
                                        Sch1.Stream.Equals(Sch2.Stream))
                                    {
                                        //Для горизонтальных надписей
                                        //не требуется расширять блок текста
                                        visTextBox = visPage.DrawRectangle(((currX) / 25.4f), (currY / 25.4f),
                                                    ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";
                                        visTextBox.Text = Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                            (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                            (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                            " (" + Sch1.SubjectType.Short + ")" + "\n" + Sch1.Auditory;
                                    }
                                    //Если хотя бы что-то не совпало, то записываем элементы через дробь
                                    else
                                    {
                                        visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));

                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move(-(3f / 25.4f), (9f / 25.4f));

                                        visTextBox.Text = Sch1.Auditory + "\n" + Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                            (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                            (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                          " (" + Sch1.SubjectType.Short + ")";

                                        visTextBox = visPage.DrawLine(((currX) / 25.4f), (currY / 25.4f),
                                                    ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                        visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move((4f / 25.4f), (1f / 25.4f));

                                        visTextBox.Text = Sch2.Spec.ShortInstitute + "-" + Sch2.KursNum.Kurs +
                                                            (Sch2.Stream.Equals("") ? (Sch2.Spec.Diff == "М" ? "5" : "1") : Sch2.Stream) +
                                                            (Sch2.Group.Equals("") ? "1" : Sch2.Group) +
                                                            " (" + Sch2.SubjectType.Short + ")" + "\n" + Sch2.Auditory;
                                    }
                                }
                            }
                            else
                            {
                                //Если элемент расписания нашёлся только для одной из недель
                                if (flgFound1Week || flgFound2Week)
                                {
                                    //Элемент для первой недели
                                    if (flgFound1Week)
                                    {
                                        visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move(-(3f / 25.4f), (9f / 25.4f));

                                        if ((Sch1.KursNum == null || Sch1.Spec == null || Sch1.SubjectType == null))
                                        {
                                            visTextBox.Text = "Есть";
                                        }
                                        else
                                        {
                                            //
                                            visTextBox.Text = Sch1.Auditory + "\n" + Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                              (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                              (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                              " (" + Sch1.SubjectType.Short + ")";
                                        }
                                    }

                                    visTextBox = visPage.DrawLine((currX / 25.4f), (currY / 25.4f),
                                                 ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                    //Элемент для второй недели
                                    if (flgFound2Week)
                                    {
                                        visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move((4f / 25.4f), (1f / 25.4f));

                                        if (Sch2.KursNum == null || Sch2.Spec == null || Sch2.SubjectType == null)
                                        {
                                            visTextBox.Text = "Есть";
                                        }
                                        else
                                        {
                                            visTextBox.Text = Sch2.Spec.ShortInstitute + "-" + Sch2.KursNum.Kurs +
                                                              (Sch2.Stream.Equals("") ? (Sch2.Spec.Diff == "М" ? "5" : "1") : Sch2.Stream) +
                                                              (Sch2.Group.Equals("") ? "1" : Sch2.Group) +
                                                              " (" + Sch2.SubjectType.Short + ")" + "\n" + Sch2.Auditory;
                                        }
                                    }
                                }
                            }

                            currX += 20f;
                        }

                        //Отступ обеда
                        currX += 10f;

                        //Запускаем цикл по временам пар
                        //после обеда
                        for (int k = 3; k < 8; k++)
                        {
                            Sch1 = null;
                            Sch2 = null;

                            flgFound1Week = DetectTimeTableElement(ref Sch1, L, Semestr, 0, j, k);
                            flgFound2Week = DetectTimeTableElement(ref Sch2, L, Semestr, 1, j, k);

                            //Если по каждой неделе нашёлся элемент расписания
                            if (flgFound1Week & flgFound2Week)
                            {
                                //Если хотя бы что-то отсутствует хотя бы у одного из элементов расписания,
                                //то выводить надпись "Есть"
                                if ((Sch1.KursNum == null || Sch1.Spec == null || Sch1.SubjectType == null) ||
                                    (Sch2.KursNum == null || Sch2.Spec == null || Sch2.SubjectType == null))
                                {
                                    //То просто пишем, что обе пары в это время есть
                                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                    visTextBox.LineStyle = "None";
                                    visTextBox.FillStyle = "None";

                                    visTextBox.Text = "Есть";
                                }
                                //Если какие-то сведения имеются
                                else
                                {
                                    //Если всё в точности совпадает, то выводим горизонтально
                                    //первый попавшийся элемент
                                    if (Sch1.SubjectType.Equals(Sch2.SubjectType) &
                                        Sch1.Spec.Equals(Sch2.Spec) &
                                        Sch1.KursNum.Equals(Sch2.KursNum) &
                                        Sch1.Auditory.Equals(Sch2.Auditory) &
                                        Sch1.Group.Equals(Sch2.Group) &
                                        Sch1.Stream.Equals(Sch2.Stream))
                                    {
                                        //Если надпись горизонтальная, то
                                        //не требуется дополнительно расширять текстовый блок
                                        visTextBox = visPage.DrawRectangle(((currX) / 25.4f), (currY / 25.4f),
                                                    ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";
                                        visTextBox.Text = Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                            (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                            (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                            " (" + Sch1.SubjectType.Short + ")" + "\n" + Sch1.Auditory;
                                    }
                                    //Если хотя бы что-то не совпало, то записываем элементы через дробь
                                    else
                                    {
                                        visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move(-(3f / 25.4f), (9f / 25.4f));

                                        visTextBox.Text = Sch1.Auditory + "\n" + Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                            (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                            (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                            " (" + Sch1.SubjectType.Short + ")";

                                        visTextBox = visPage.DrawLine(((currX) / 25.4f), (currY / 25.4f),
                                                    ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                        visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move((4f / 25.4f), (1f / 25.4f));

                                        visTextBox.Text = Sch2.Spec.ShortInstitute + "-" + Sch2.KursNum.Kurs +
                                                            (Sch2.Stream.Equals("") ? (Sch2.Spec.Diff == "М" ? "5" : "1") : Sch2.Stream) +
                                                            (Sch2.Group.Equals("") ? "1" : Sch2.Group) +
                                                            " (" + Sch2.SubjectType.Short + ")" + "\n" + Sch2.Auditory;
                                    }
                                }
                            }
                            else
                            {
                                //Если элемент расписания нашёлся только для одной из недель
                                if (flgFound1Week || flgFound2Week)
                                {
                                    //Элемент для первой недели
                                    if (flgFound1Week)
                                    {
                                        visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move(-(3f / 25.4f), (9f / 25.4f));

                                        if ((Sch1.KursNum == null || Sch1.Spec == null || Sch1.SubjectType == null))
                                        {
                                            visTextBox.Text = "Есть";
                                        }
                                        else
                                        {
                                            visTextBox.Text = Sch1.Auditory + "\n" + Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                                (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                                (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                                " (" + Sch1.SubjectType.Short + ")";
                                        }
                                    }

                                    visTextBox = visPage.DrawLine((currX / 25.4f), (currY / 25.4f),
                                                 ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                    //Элемент для второй недели
                                    if (flgFound2Week)
                                    {
                                        visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        //Снять выделение со всех элементов
                                        visApp.ActiveWindow.DeselectAll();
                                        visApp.ActiveWindow.Select(visTextBox, 2);
                                        visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                        visApp.ActiveWindow.Selection.Move((4f / 25.4f), (1f / 25.4f));

                                        if (Sch2.KursNum == null || Sch2.Spec == null || Sch2.SubjectType == null)
                                        {
                                            visTextBox.Text = "Есть";
                                        }
                                        else
                                        {
                                            visTextBox.Text = Sch2.Spec.ShortInstitute + "-" + Sch2.KursNum.Kurs +
                                                                (Sch2.Stream.Equals("") ? (Sch2.Spec.Diff == "М" ? "5" : "1") : Sch2.Stream) +
                                                                (Sch2.Group.Equals("") ? "1" : Sch2.Group) +
                                                                " (" + Sch2.SubjectType.Short + ")" + "\n" + Sch2.Auditory;
                                        }
                                    }
                                }
                            }

                            currX += 20f;
                        }
                    }

                    currY += 20f;
                }


                visApp.Visible = true;
            }
            else
            {
                MessageBox.Show("Расписание для данного преподавателя не найдено!", "Ошибка");
                visApp.Quit();
            }
        }

        private bool DetectCheckedSchedule(clsLecturer L, int Semestr)
        {
            bool flg = false;

            for (int i = 0; i < mdlData.colSchedule.Count; i++)
            {
                //Если совпали семестры, преподаватели и есть признак занятия
                //хотя бы у одного из элементов расписания
                if (mdlData.colSchedule[i].Lecturer.FIO.Equals(L.FIO) &
                    mdlData.colSchedule[i].Semestr.SemNum.Equals(mdlData.colSemestr[Semestr].SemNum) &
                    mdlData.colSchedule[i].Subj)
                {
                    flg = true;
                    break;
                }
            }

            return flg;
        }

        //Определение наличия интересующего элемента расписания
        private bool DetectTimeTableElement(ref clsSchedule Sch, clsLecturer L, int Semestr, int Week, int WeekDay, int PairTime)
        {
            bool flg = false;
            for (int l = mdlData.colSchedule.Count - 1; l >= 0; l--)
            {
                Sch = mdlData.colSchedule[l];
                //Если совпали преподаватели
                if (Sch.Lecturer.FIO.Equals(L.FIO))
                {
                    //Если совпали семестры
                    if (Sch.Semestr.SemNum.Equals(mdlData.colSemestr[Semestr].SemNum))
                    {
                        //Если совпали учебные недели
                        if (Sch.Week.NumberWeek.Equals(mdlData.colWeek[Week].NumberWeek))
                        {
                            //Если совпали дни недели
                            if (Sch.WeekDay.WeekDay.Equals(mdlData.colWeekDays[WeekDay].WeekDay))
                            {
                                //Если совпало время занятий
                                if (Sch.Time.Time.Equals(mdlData.colPairTime[PairTime].Time))
                                {
                                    if (Sch.Subj)
                                    {
                                        flg = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return flg;
        }

        private void btnChoose_Click(object sender, EventArgs e)
        {
            if (cmbChoose.SelectedIndex == -1)
            {
                MessageBox.Show("Выберите преподавателя из списка", "Ошибка");
            }

            else
            {
                onVisioTimeTableSimple();
            }
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            if (cmbChoose.SelectedIndex == -1)
            {
                MessageBox.Show("Выберите преподавателя из списка", "Ошибка");
            }

            else
            {
                onVisioTimeTableSimple2();
            }
        }
        private void onVisioTimeTableSimple2()
        {
            string visDocName = Application.StartupPath + "\\myVisio.vsd";
            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;
            clsLecturer L;
            clsSchedule Sch1 = new clsSchedule();
            clsSchedule Sch2 = new clsSchedule();
            bool flgFound1Week;
            bool flgFound2Week;
            bool flgDinnerPlaced = false;

            float currY = 0f;
            float currX = 0f;
            float tempX = 0f;
            float dinX = 0f;
            float XStart = 0f;

            float XLeft = 10f;
            float XRight = 0f;
            float YBottom = 10f;
            float YTop = 0f;

            int Semestr = 0;
            bool trigColor;

            Visio.Shape visTextBox;
            Visio.Page visPage;
            Visio.Application visApp;
            Visio.Document visDoc;

            visApp = new Visio.Application();
            visDoc = visApp.Documents.Add("");
            visPage = visApp.ActivePage;


            //Вопрос: для какого семестра создаётся расписание?
            if (MessageBox.Show(this, "Создаётся расписание для I семестра?", "Выбор семестра", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Semestr = 1;
            }
            else
            {
                Semestr = 2;
            }

            //---------------С этого момента создаётся пустая таблица----------
            //
            currY = 10f;
            trigColor = false;
            //Получаем номер преподавателя в коллекции
            int lectID = new int();
            for (int i = mdlData.colLecturer.Count - 1; i >= 0; i--)
            {
                if (mdlData.colLecturer[i].FIO == cmbChoose.SelectedItem.ToString())
                {
                    lectID = i;
                }
            }

            currX = 10f;

            if (DetectCheckedSchedule(mdlData.colLecturer[lectID], Semestr))
            {
                //Запускаем цикл по дням недели
                for (int j = 0; j < 6; j++)
            {
                //Рисуем поле для дня недели
                tempX = currX;
                visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                ((currX + 40f) / 25.4f), ((currY + 20f) / 25.4f));
                currX = currX + 40f;
                visTextBox.Text = mdlData.colWeekDays[5-j].WeekDay;
                //Запускаем цикл по временам пар
                //до обеда
                for (int k = 0; k < 3; k++)
                {
                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                    currX += 20f;
                }

                dinX = currX;
                currX += 10f;

                //Запускаем цикл по временам пар
                //после обеда
                for (int k = 3; k < 8; k++)
                {
                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));


                    currX += 20f;
                }


                if (j != 5)

                {
                    currX = tempX;
                }

                currY = currY + 20f;


            }

            //Обед
            currY = 10f;
            visTextBox = visPage.DrawRectangle((dinX / 25.4f), (currY / 25.4f),
                        ((dinX + 120f) / 25.4f), ((currY + 10f) / 25.4f));
            visTextBox.Text = "Обеденный перерыв";
            visTextBox.Rotate90();
            visApp.ActiveWindow.Select(visTextBox, 2);
            visApp.ActiveWindow.Selection.Move(-(55f / 25.4f), (55f / 25.4f));

            //Рисуем шапку

            currX = 10f;
            currY = 130f;

            visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                            ((currX + 40f) / 25.4f), ((currY + 25f) / 25.4f));
            visTextBox = visPage.DrawLine((currX / 25.4f), ((currY + 25f) / 25.4f),
                            ((currX + 40f) / 25.4f), (currY / 25.4f));
            //Пишем наискосок в шапку
            visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f) / 25.4f), ((currY + 10f) / 25.4f));
            visTextBox.LineStyle = "None";
            visTextBox.FillStyle = "None";

            //Снять выделение со всех элементов
            visApp.ActiveWindow.DeselectAll();
            visApp.ActiveWindow.Select(visTextBox, 2);

            visApp.ActiveWindow.Selection.Move((20f / 25.4f), (10f / 25.4f));

            visTextBox.Text = "Часы";

            visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f) / 25.4f), ((currY + 10f) / 25.4f));
            visTextBox.LineStyle = "None";
            visTextBox.FillStyle = "None";

            //Снять выделение со всех элементов
            visApp.ActiveWindow.DeselectAll();
            visApp.ActiveWindow.Select(visTextBox, 2);

            visApp.ActiveWindow.Selection.Move((0f / 25.4f), (5f / 25.4f));

            visTextBox.Text = "Дни";


            currX = currX + 40f;

            visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                            ((currX + 170f) / 25.4f), ((currY + 5f) / 25.4f));
            visTextBox.Text = "Учебная группа (поток, аббревиатура, номер, аудитория)";

            

            currY = currY + 5f;

            for (int k = 0; k < 9; k++)
            {
                if (k < 3)
                {
                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));
                        visTextBox.Text = mdlData.colPairTime[k].Time;
                    }

                else
                {
                    if (k == 3)
                    {
                        visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                            ((currX + 20f) / 25.4f), ((currY + 10f) / 25.4f));
                        visTextBox.Text = "Обед";
                        visTextBox.Rotate90();
                        visApp.ActiveWindow.Select(visTextBox, 2);
                        visApp.ActiveWindow.Selection.Move(-(5f / 25.4f), (5f / 25.4f));
                    }

                    else
                    {
                        visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));
                        visTextBox.Text = mdlData.colPairTime[k-1].Time;
                    }
                }
                if (k == 3)
                {
                    currX += 10f;
                }

                else
                {
                    currX += 20f;
                }
            }
            currX = 10f;
            currY += 20f;

            visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                ((currX + 210f) / 25.4f), ((currY + 20f) / 25.4f));
            visTextBox.Text = mdlData.colLecturer[lectID].FIO;
            visTextBox.CellsU["Char.Size"].FormulaForceU = "24 pt";

            //Конец отрисовки шапки

            //---------------С этого момента заполняется таблица---------------
            //
            currY = 10f;

            currX = 10f;

            //if (mdlData.colLecturer[i].Rate > 0 || DetectCheckedSchedule(mdlData.colLecturer[i], Semestr))
            if (DetectCheckedSchedule(mdlData.colLecturer[lectID], Semestr))
            {
                L = mdlData.colLecturer[lectID];

                currX += 40f;

                //Запускаем цикл по дням недели
                for (int j = 0; j < 6; j++)
                {
                    //запускаем цикл по временам занятий
                    //до обеда
                    for (int k = 0; k < 3; k++)
                    {
                        Sch1 = null;
                        Sch2 = null;
                        flgFound1Week = DetectTimeTableElement(ref Sch1, L, Semestr, 0, 5-j, k);
                        flgFound2Week = DetectTimeTableElement(ref Sch2, L, Semestr, 1, 5-j, k);

                        // Смена координат для перехода на новую строку
                        if (currX >= 110f)
                        {
                            currX = 50f;
                            currY += 20f;
                        }

                        else
                        {
                            if (k == 7 && currX == 120f)
                            {
                                currY += 20f;
                            }
                        }



                        //Если по каждой неделе нашёлся элемент расписания
                        if (flgFound1Week & flgFound2Week)
                        {
                            //Если хотя бы что-то отсутствует у одного из элементов расписания, то
                            //выводить надпись "Есть"
                            if ((Sch1.KursNum == null || Sch1.Spec == null || Sch1.SubjectType == null) ||
                                (Sch2.KursNum == null || Sch2.Spec == null || Sch2.SubjectType == null))
                            {
                                //То просто пишем, что обе пары в это время есть
                                
                                visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                            ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                visTextBox.LineStyle = "None";
                                visTextBox.FillStyle = "None";
                                visTextBox.Text = "Есть";
                            }
                            //Если какие-то сведения имеются
                            else
                            {
                                //Если всё в точности совпадает, то выводим горизонтально
                                //первый попавшийся элемент
                                if (Sch1.SubjectType.Equals(Sch2.SubjectType) &
                                    Sch1.Spec.Equals(Sch2.Spec) &
                                    Sch1.KursNum.Equals(Sch2.KursNum) &
                                    Sch1.Auditory.Equals(Sch2.Auditory) &
                                    Sch1.Group.Equals(Sch2.Group) &
                                    Sch1.Stream.Equals(Sch2.Stream))
                                {
                                    //Для горизонтальных надписей
                                    //не требуется расширять блок текста
                                    
                                    visTextBox = visPage.DrawRectangle(((currX) / 25.4f), (currY / 25.4f),
                                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                    visTextBox.LineStyle = "None";
                                    visTextBox.FillStyle = "None";
                                    visTextBox.Text = Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                        (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                        (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                        " (" + Sch1.SubjectType.Short + ")" + "\n" + Sch1.Auditory;
                                }
                                //Если хотя бы что-то не совпало, то записываем элементы через дробь
                                else
                                {
                                    visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                 ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));

                                    visTextBox.LineStyle = "None";
                                    visTextBox.FillStyle = "None";

                                    //Снять выделение со всех элементов
                                    visApp.ActiveWindow.DeselectAll();
                                    visApp.ActiveWindow.Select(visTextBox, 2);
                                    visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                    visApp.ActiveWindow.Selection.Move(-(3f / 25.4f), (9f / 25.4f));

                                    visTextBox.Text = Sch1.Auditory + "\n" + Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                        (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                        (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                      " (" + Sch1.SubjectType.Short + ")";

                                    visTextBox = visPage.DrawLine(((currX) / 25.4f), (currY / 25.4f),
                                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                    visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                 ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                    visTextBox.LineStyle = "None";
                                    visTextBox.FillStyle = "None";

                                    //Снять выделение со всех элементов
                                    visApp.ActiveWindow.DeselectAll();
                                    visApp.ActiveWindow.Select(visTextBox, 2);
                                    visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                    visApp.ActiveWindow.Selection.Move((4f / 25.4f), (1f / 25.4f));

                                    visTextBox.Text = Sch2.Spec.ShortInstitute + "-" + Sch2.KursNum.Kurs +
                                                        (Sch2.Stream.Equals("") ? (Sch2.Spec.Diff == "М" ? "5" : "1") : Sch2.Stream) +
                                                        (Sch2.Group.Equals("") ? "1" : Sch2.Group) +
                                                        " (" + Sch2.SubjectType.Short + ")" + "\n" + Sch2.Auditory;
                                }
                            }
                        }
                        else
                        {
                            //Если элемент расписания нашёлся только для одной из недель
                            if (flgFound1Week || flgFound2Week)
                            {
                                //Элемент для первой недели
                                if (flgFound1Week)
                                {
                                    visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                 ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                    visTextBox.LineStyle = "None";
                                    visTextBox.FillStyle = "None";

                                    //Снять выделение со всех элементов
                                    visApp.ActiveWindow.DeselectAll();
                                    visApp.ActiveWindow.Select(visTextBox, 2);
                                    visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                    visApp.ActiveWindow.Selection.Move(-(3f / 25.4f), (9f / 25.4f));

                                    if ((Sch1.KursNum == null || Sch1.Spec == null || Sch1.SubjectType == null))
                                    {
                                        visTextBox.Text = "Есть";
                                    }
                                    else
                                    {
                                        //
                                        visTextBox.Text = Sch1.Auditory + "\n" + Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                          (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                          (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                          " (" + Sch1.SubjectType.Short + ")";
                                    }
                                }

                                visTextBox = visPage.DrawLine((currX / 25.4f), (currY / 25.4f),
                                             ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                //Элемент для второй недели
                                if (flgFound2Week)
                                {
                                    // Смена координат для перехода на новую строку
                                    if (currX == 110f)
                                    {
                                        currX = 50f;
                                        currY += 20f;
                                    }
                                    visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                 ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                    visTextBox.LineStyle = "None";
                                    visTextBox.FillStyle = "None";

                                    //Снять выделение со всех элементов
                                    visApp.ActiveWindow.DeselectAll();
                                    visApp.ActiveWindow.Select(visTextBox, 2);
                                    visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                    visApp.ActiveWindow.Selection.Move((4f / 25.4f), (1f / 25.4f));

                                    if (Sch2.KursNum == null || Sch2.Spec == null || Sch2.SubjectType == null)
                                    {
                                        visTextBox.Text = "Есть";
                                    }
                                    else
                                    {
                                        visTextBox.Text = Sch2.Spec.ShortInstitute + "-" + Sch2.KursNum.Kurs +
                                                          (Sch2.Stream.Equals("") ? (Sch2.Spec.Diff == "М" ? "5" : "1") : Sch2.Stream) +
                                                          (Sch2.Group.Equals("") ? "1" : Sch2.Group) +
                                                          " (" + Sch2.SubjectType.Short + ")" + "\n" + Sch2.Auditory;
                                    }
                                }
                            }
                        }

                        currX += 20f;
                    }

                    //Отступ обеда
                    currX += 10f;

                    //Запускаем цикл по временам пар
                    //после обеда
                    for (int k = 3; k < 8; k++)
                    {
                        Sch1 = null;
                        Sch2 = null;

                        flgFound1Week = DetectTimeTableElement(ref Sch1, L, Semestr, 0, 5-j, k);
                        flgFound2Week = DetectTimeTableElement(ref Sch2, L, Semestr, 1, 5-j, k);

                        // Смена координат для перехода на новую строку
                        if (currX >= 220f)
                        {
                            currX = 120f;
                            currY += 20f;
                        }
                        
                        else
                        {
                            if (k == 7 && currX == 120f)
                            {
                                currY += 20f;
                            }
                        }



                        //Если по каждой неделе нашёлся элемент расписания
                        if (flgFound1Week & flgFound2Week)
                        {
                            //Если хотя бы что-то отсутствует хотя бы у одного из элементов расписания,
                            //то выводить надпись "Есть"
                            if ((Sch1.KursNum == null || Sch1.Spec == null || Sch1.SubjectType == null) ||
                                (Sch2.KursNum == null || Sch2.Spec == null || Sch2.SubjectType == null))
                            {
                                //То просто пишем, что обе пары в это время есть
                                visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                            ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                visTextBox.LineStyle = "None";
                                visTextBox.FillStyle = "None";

                                visTextBox.Text = "Есть";
                            }
                            //Если какие-то сведения имеются
                            else
                            {
                                //Если всё в точности совпадает, то выводим горизонтально
                                //первый попавшийся элемент
                                if (Sch1.SubjectType.Equals(Sch2.SubjectType) &
                                    Sch1.Spec.Equals(Sch2.Spec) &
                                    Sch1.KursNum.Equals(Sch2.KursNum) &
                                    Sch1.Auditory.Equals(Sch2.Auditory) &
                                    Sch1.Group.Equals(Sch2.Group) &
                                    Sch1.Stream.Equals(Sch2.Stream))
                                {
                                    //Если надпись горизонтальная, то
                                    //не требуется дополнительно расширять текстовый блок
                                    visTextBox = visPage.DrawRectangle(((currX) / 25.4f), (currY / 25.4f),
                                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                    visTextBox.LineStyle = "None";
                                    visTextBox.FillStyle = "None";
                                    visTextBox.Text = Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                        (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                        (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                        " (" + Sch1.SubjectType.Short + ")" + "\n" + Sch1.Auditory;
                                }
                                //Если хотя бы что-то не совпало, то записываем элементы через дробь
                                else
                                {
                                    visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                 ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                    visTextBox.LineStyle = "None";
                                    visTextBox.FillStyle = "None";

                                    //Снять выделение со всех элементов
                                    visApp.ActiveWindow.DeselectAll();
                                    visApp.ActiveWindow.Select(visTextBox, 2);
                                    visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                    visApp.ActiveWindow.Selection.Move(-(3f / 25.4f), (9f / 25.4f));

                                    visTextBox.Text = Sch1.Auditory + "\n" + Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                        (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                        (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                        " (" + Sch1.SubjectType.Short + ")";

                                    visTextBox = visPage.DrawLine(((currX) / 25.4f), (currY / 25.4f),
                                                ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                    visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                 ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                    visTextBox.LineStyle = "None";
                                    visTextBox.FillStyle = "None";

                                    //Снять выделение со всех элементов
                                    visApp.ActiveWindow.DeselectAll();
                                    visApp.ActiveWindow.Select(visTextBox, 2);
                                    visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                    visApp.ActiveWindow.Selection.Move((4f / 25.4f), (1f / 25.4f));

                                    visTextBox.Text = Sch2.Spec.ShortInstitute + "-" + Sch2.KursNum.Kurs +
                                                        (Sch2.Stream.Equals("") ? (Sch2.Spec.Diff == "М" ? "5" : "1") : Sch2.Stream) +
                                                        (Sch2.Group.Equals("") ? "1" : Sch2.Group) +
                                                        " (" + Sch2.SubjectType.Short + ")" + "\n" + Sch2.Auditory;
                                }
                            }
                        }
                        else
                        {
                            //Если элемент расписания нашёлся только для одной из недель
                            if (flgFound1Week || flgFound2Week)
                            {
                                //Элемент для первой недели
                                if (flgFound1Week)
                                {
                                    visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                 ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                    visTextBox.LineStyle = "None";
                                    visTextBox.FillStyle = "None";

                                    //Снять выделение со всех элементов
                                    visApp.ActiveWindow.DeselectAll();
                                    visApp.ActiveWindow.Select(visTextBox, 2);
                                    visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                    visApp.ActiveWindow.Selection.Move(-(3f / 25.4f), (9f / 25.4f));

                                    if ((Sch1.KursNum == null || Sch1.Spec == null || Sch1.SubjectType == null))
                                    {
                                        visTextBox.Text = "Есть";
                                    }
                                    else
                                    {
                                        visTextBox.Text = Sch1.Auditory + "\n" + Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                            (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                            (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                            " (" + Sch1.SubjectType.Short + ")";
                                    }
                                }

                                visTextBox = visPage.DrawLine((currX / 25.4f), (currY / 25.4f),
                                             ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                //Элемент для второй недели
                                if (flgFound2Week)
                                {
                                    visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                 ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                    visTextBox.LineStyle = "None";
                                    visTextBox.FillStyle = "None";

                                    //Снять выделение со всех элементов
                                    visApp.ActiveWindow.DeselectAll();
                                    visApp.ActiveWindow.Select(visTextBox, 2);
                                    visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                    visApp.ActiveWindow.Selection.Move((4f / 25.4f), (1f / 25.4f));

                                    if (Sch2.KursNum == null || Sch2.Spec == null || Sch2.SubjectType == null)
                                    {
                                        visTextBox.Text = "Есть";
                                    }
                                    else
                                    {
                                        visTextBox.Text = Sch2.Spec.ShortInstitute + "-" + Sch2.KursNum.Kurs +
                                                            (Sch2.Stream.Equals("") ? (Sch2.Spec.Diff == "М" ? "5" : "1") : Sch2.Stream) +
                                                            (Sch2.Group.Equals("") ? "1" : Sch2.Group) +
                                                            " (" + Sch2.SubjectType.Short + ")" + "\n" + Sch2.Auditory;
                                    }
                                }
                            }
                        }

                        currX += 20f;
                    }
                }

                currY += 20f;
            }


            visApp.Visible = true;
            }
            else
            {
                MessageBox.Show("Расписание для данного преподавателя не найдено!", "Ошибка");
                visApp.Quit();
            }
        }

        private void onVisioTimeTableForAll()
        {
                string visDocName = Application.StartupPath + "\\myVisio.vsd";
                //Задаём переменную для отсутствующего параметра
                object ObjMissing = Missing.Value;
                clsLecturer L;
                clsSchedule Sch1 = new clsSchedule();
                clsSchedule Sch2 = new clsSchedule();
                bool flgFound1Week;
                bool flgFound2Week;
                bool flgDinnerPlaced = false;

                float currY = 0f;
                float currX = 0f;
                float tempX = 0f;
                float dinX = 0f;
                float XStart = 0f;

                float XLeft = 10f;
                float XRight = 0f;
                float YBottom = 10f;
                float YTop = 0f;

                int Semestr = 0;
                bool trigColor;

                Visio.Shape visTextBox;
                Visio.Page visPage;
                Visio.Application visApp;
                Visio.Document visDoc;

                visApp = new Visio.Application();
                visApp.Visible = false;
                visDoc = visApp.Documents.Add("");
                visPage = visApp.ActivePage;


                //Вопрос: для какого семестра создаётся расписание?
                if (MessageBox.Show(this, "Создаётся расписание для I семестра?", "Выбор семестра", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Semestr = 1;
                }
                else
                {
                    Semestr = 2;
                }

            MessageBox.Show("Начинается процесс формирования карточек расписаний преподавателей кафедры. Пожалуйста, дождитесь появления сообщения о его окончании не взаимодействуя с программой.", "Внимание!");

            for (int a = 0; a < cmbChoose.Items.Count; a++)
            {
                cmbChoose.SelectedIndex = a;

                if (DetectCheckedSchedule(mdlData.colLecturer[a], Semestr))
                {

                    //---------------С этого момента создаётся пустая таблица----------
                    //
                    currY = 10f;
                    trigColor = false;
                    //Получаем номер преподавателя в коллекции
                    int lectID = new int();
                    lectID = a;

                    currX = 10f;


                    //Запускаем цикл по дням недели
                    for (int j = 0; j < 6; j++)
                    {
                        //Рисуем поле для дня недели
                        tempX = currX;
                        visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                        ((currX + 40f) / 25.4f), ((currY + 20f) / 25.4f));
                        currX = currX + 40f;
                        visTextBox.Text = mdlData.colWeekDays[5 - j].WeekDay;
                        //Запускаем цикл по временам пар
                        //до обеда
                        for (int k = 0; k < 3; k++)
                        {
                            visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                        ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                            currX += 20f;
                        }

                        dinX = currX;
                        currX += 10f;

                        //Запускаем цикл по временам пар
                        //после обеда
                        for (int k = 3; k < 8; k++)
                        {
                            visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                        ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));


                            currX += 20f;
                        }


                        if (j != 5)

                        {
                            currX = tempX;
                        }

                        currY = currY + 20f;


                    }

                    //Обед
                    currY = 10f;
                    visTextBox = visPage.DrawRectangle((dinX / 25.4f), (currY / 25.4f),
                                ((dinX + 120f) / 25.4f), ((currY + 10f) / 25.4f));
                    visTextBox.Text = "Обеденный перерыв";
                    visTextBox.Rotate90();
                    visApp.ActiveWindow.Select(visTextBox, 2);
                    visApp.ActiveWindow.Selection.Move(-(55f / 25.4f), (55f / 25.4f));

                    //Рисуем шапку

                    currX = 10f;
                    currY = 130f;

                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                    ((currX + 40f) / 25.4f), ((currY + 25f) / 25.4f));
                    visTextBox = visPage.DrawLine((currX / 25.4f), ((currY + 25f) / 25.4f),
                                    ((currX + 40f) / 25.4f), (currY / 25.4f));
                    //Пишем наискосок в шапку
                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                             ((currX + 20f) / 25.4f), ((currY + 10f) / 25.4f));
                    visTextBox.LineStyle = "None";
                    visTextBox.FillStyle = "None";

                    //Снять выделение со всех элементов
                    visApp.ActiveWindow.DeselectAll();
                    visApp.ActiveWindow.Select(visTextBox, 2);

                    visApp.ActiveWindow.Selection.Move((20f / 25.4f), (10f / 25.4f));

                    visTextBox.Text = "Часы";

                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                             ((currX + 20f) / 25.4f), ((currY + 10f) / 25.4f));
                    visTextBox.LineStyle = "None";
                    visTextBox.FillStyle = "None";

                    //Снять выделение со всех элементов
                    visApp.ActiveWindow.DeselectAll();
                    visApp.ActiveWindow.Select(visTextBox, 2);

                    visApp.ActiveWindow.Selection.Move((0f / 25.4f), (5f / 25.4f));

                    visTextBox.Text = "Дни";


                    currX = currX + 40f;

                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                    ((currX + 170f) / 25.4f), ((currY + 5f) / 25.4f));
                    visTextBox.Text = "Учебная группа (поток, аббревиатура, номер, аудитория)";



                    currY = currY + 5f;

                    for (int k = 0; k < 9; k++)
                    {
                        if (k < 3)
                        {
                            visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                        ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));
                            visTextBox.Text = mdlData.colPairTime[k].Time;
                        }

                        else
                        {
                            if (k == 3)
                            {
                                visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                    ((currX + 20f) / 25.4f), ((currY + 10f) / 25.4f));
                                visTextBox.Text = "Обед";
                                visTextBox.Rotate90();
                                visApp.ActiveWindow.Select(visTextBox, 2);
                                visApp.ActiveWindow.Selection.Move(-(5f / 25.4f), (5f / 25.4f));
                            }

                            else
                            {
                                visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                        ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));
                                visTextBox.Text = mdlData.colPairTime[k - 1].Time;
                            }
                        }
                        if (k == 3)
                        {
                            currX += 10f;
                        }

                        else
                        {
                            currX += 20f;
                        }
                    }
                    currX = 10f;
                    currY += 20f;

                    visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                        ((currX + 210f) / 25.4f), ((currY + 20f) / 25.4f));
                    visTextBox.Text = mdlData.colLecturer[lectID].FIO;
                    visTextBox.CellsU["Char.Size"].FormulaForceU = "24 pt";

                    //Конец отрисовки шапки

                    //---------------С этого момента заполняется таблица---------------
                    //
                    currY = 10f;

                    currX = 10f;

                    //if (mdlData.colLecturer[i].Rate > 0 || DetectCheckedSchedule(mdlData.colLecturer[i], Semestr))
                    if (DetectCheckedSchedule(mdlData.colLecturer[lectID], Semestr))
                    {
                        L = mdlData.colLecturer[lectID];

                        currX += 40f;

                        //Запускаем цикл по дням недели
                        for (int j = 0; j < 6; j++)
                        {
                            //запускаем цикл по временам занятий
                            //до обеда
                            for (int k = 0; k < 3; k++)
                            {
                                Sch1 = null;
                                Sch2 = null;
                                flgFound1Week = DetectTimeTableElement(ref Sch1, L, Semestr, 0, 5 - j, k);
                                flgFound2Week = DetectTimeTableElement(ref Sch2, L, Semestr, 1, 5 - j, k);

                                // Смена координат для перехода на новую строку
                                if (currX >= 110f)
                                {
                                    currX = 50f;
                                    currY += 20f;
                                }

                                else
                                {
                                    if (k == 7 && currX == 120f)
                                    {
                                        currY += 20f;
                                    }
                                }



                                //Если по каждой неделе нашёлся элемент расписания
                                if (flgFound1Week & flgFound2Week)
                                {
                                    //Если хотя бы что-то отсутствует у одного из элементов расписания, то
                                    //выводить надпись "Есть"
                                    if ((Sch1.KursNum == null || Sch1.Spec == null || Sch1.SubjectType == null) ||
                                        (Sch2.KursNum == null || Sch2.Spec == null || Sch2.SubjectType == null))
                                    {
                                        //То просто пишем, что обе пары в это время есть

                                        visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                    ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";
                                        visTextBox.Text = "Есть";
                                    }
                                    //Если какие-то сведения имеются
                                    else
                                    {
                                        //Если всё в точности совпадает, то выводим горизонтально
                                        //первый попавшийся элемент
                                        if (Sch1.SubjectType.Equals(Sch2.SubjectType) &
                                            Sch1.Spec.Equals(Sch2.Spec) &
                                            Sch1.KursNum.Equals(Sch2.KursNum) &
                                            Sch1.Auditory.Equals(Sch2.Auditory) &
                                            Sch1.Group.Equals(Sch2.Group) &
                                            Sch1.Stream.Equals(Sch2.Stream))
                                        {
                                            //Для горизонтальных надписей
                                            //не требуется расширять блок текста

                                            visTextBox = visPage.DrawRectangle(((currX) / 25.4f), (currY / 25.4f),
                                                        ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                            visTextBox.LineStyle = "None";
                                            visTextBox.FillStyle = "None";
                                            visTextBox.Text = Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                                (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                                (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                                " (" + Sch1.SubjectType.Short + ")" + "\n" + Sch1.Auditory;
                                        }
                                        //Если хотя бы что-то не совпало, то записываем элементы через дробь
                                        else
                                        {
                                            visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                         ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));

                                            visTextBox.LineStyle = "None";
                                            visTextBox.FillStyle = "None";

                                            //Снять выделение со всех элементов
                                            visApp.ActiveWindow.DeselectAll();
                                            visApp.ActiveWindow.Select(visTextBox, 2);
                                            visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                            visApp.ActiveWindow.Selection.Move(-(3f / 25.4f), (9f / 25.4f));

                                            visTextBox.Text = Sch1.Auditory + "\n" + Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                                (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                                (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                              " (" + Sch1.SubjectType.Short + ")";

                                            visTextBox = visPage.DrawLine(((currX) / 25.4f), (currY / 25.4f),
                                                        ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                            visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                         ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                            visTextBox.LineStyle = "None";
                                            visTextBox.FillStyle = "None";

                                            //Снять выделение со всех элементов
                                            visApp.ActiveWindow.DeselectAll();
                                            visApp.ActiveWindow.Select(visTextBox, 2);
                                            visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                            visApp.ActiveWindow.Selection.Move((4f / 25.4f), (1f / 25.4f));

                                            visTextBox.Text = Sch2.Spec.ShortInstitute + "-" + Sch2.KursNum.Kurs +
                                                                (Sch2.Stream.Equals("") ? (Sch2.Spec.Diff == "М" ? "5" : "1") : Sch2.Stream) +
                                                                (Sch2.Group.Equals("") ? "1" : Sch2.Group) +
                                                                " (" + Sch2.SubjectType.Short + ")" + "\n" + Sch2.Auditory;
                                        }
                                    }
                                }
                                else
                                {
                                    //Если элемент расписания нашёлся только для одной из недель
                                    if (flgFound1Week || flgFound2Week)
                                    {
                                        //Элемент для первой недели
                                        if (flgFound1Week)
                                        {
                                            visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                         ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                            visTextBox.LineStyle = "None";
                                            visTextBox.FillStyle = "None";

                                            //Снять выделение со всех элементов
                                            visApp.ActiveWindow.DeselectAll();
                                            visApp.ActiveWindow.Select(visTextBox, 2);
                                            visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                            visApp.ActiveWindow.Selection.Move(-(3f / 25.4f), (9f / 25.4f));

                                            if ((Sch1.KursNum == null || Sch1.Spec == null || Sch1.SubjectType == null))
                                            {
                                                visTextBox.Text = "Есть";
                                            }
                                            else
                                            {
                                                //
                                                visTextBox.Text = Sch1.Auditory + "\n" + Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                                  (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                                  (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                                  " (" + Sch1.SubjectType.Short + ")";
                                            }
                                        }

                                        visTextBox = visPage.DrawLine((currX / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                        //Элемент для второй недели
                                        if (flgFound2Week)
                                        {
                                            // Смена координат для перехода на новую строку
                                            if (currX == 110f)
                                            {
                                                currX = 50f;
                                                currY += 20f;
                                            }
                                            visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                         ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                            visTextBox.LineStyle = "None";
                                            visTextBox.FillStyle = "None";

                                            //Снять выделение со всех элементов
                                            visApp.ActiveWindow.DeselectAll();
                                            visApp.ActiveWindow.Select(visTextBox, 2);
                                            visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                            visApp.ActiveWindow.Selection.Move((4f / 25.4f), (1f / 25.4f));

                                            if (Sch2.KursNum == null || Sch2.Spec == null || Sch2.SubjectType == null)
                                            {
                                                visTextBox.Text = "Есть";
                                            }
                                            else
                                            {
                                                visTextBox.Text = Sch2.Spec.ShortInstitute + "-" + Sch2.KursNum.Kurs +
                                                                  (Sch2.Stream.Equals("") ? (Sch2.Spec.Diff == "М" ? "5" : "1") : Sch2.Stream) +
                                                                  (Sch2.Group.Equals("") ? "1" : Sch2.Group) +
                                                                  " (" + Sch2.SubjectType.Short + ")" + "\n" + Sch2.Auditory;
                                            }
                                        }
                                    }
                                }

                                currX += 20f;
                            }

                            //Отступ обеда
                            currX += 10f;

                            //Запускаем цикл по временам пар
                            //после обеда
                            for (int k = 3; k < 8; k++)
                            {
                                Sch1 = null;
                                Sch2 = null;

                                flgFound1Week = DetectTimeTableElement(ref Sch1, L, Semestr, 0, 5 - j, k);
                                flgFound2Week = DetectTimeTableElement(ref Sch2, L, Semestr, 1, 5 - j, k);

                                // Смена координат для перехода на новую строку
                                if (currX >= 220f)
                                {
                                    currX = 120f;
                                    currY += 20f;
                                }

                                else
                                {
                                    if (k == 7 && currX == 120f)
                                    {
                                        currY += 20f;
                                    }
                                }



                                //Если по каждой неделе нашёлся элемент расписания
                                if (flgFound1Week & flgFound2Week)
                                {
                                    //Если хотя бы что-то отсутствует хотя бы у одного из элементов расписания,
                                    //то выводить надпись "Есть"
                                    if ((Sch1.KursNum == null || Sch1.Spec == null || Sch1.SubjectType == null) ||
                                        (Sch2.KursNum == null || Sch2.Spec == null || Sch2.SubjectType == null))
                                    {
                                        //То просто пишем, что обе пары в это время есть
                                        visTextBox = visPage.DrawRectangle((currX / 25.4f), (currY / 25.4f),
                                                    ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                        visTextBox.LineStyle = "None";
                                        visTextBox.FillStyle = "None";

                                        visTextBox.Text = "Есть";
                                    }
                                    //Если какие-то сведения имеются
                                    else
                                    {
                                        //Если всё в точности совпадает, то выводим горизонтально
                                        //первый попавшийся элемент
                                        if (Sch1.SubjectType.Equals(Sch2.SubjectType) &
                                            Sch1.Spec.Equals(Sch2.Spec) &
                                            Sch1.KursNum.Equals(Sch2.KursNum) &
                                            Sch1.Auditory.Equals(Sch2.Auditory) &
                                            Sch1.Group.Equals(Sch2.Group) &
                                            Sch1.Stream.Equals(Sch2.Stream))
                                        {
                                            //Если надпись горизонтальная, то
                                            //не требуется дополнительно расширять текстовый блок
                                            visTextBox = visPage.DrawRectangle(((currX) / 25.4f), (currY / 25.4f),
                                                        ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                            visTextBox.LineStyle = "None";
                                            visTextBox.FillStyle = "None";
                                            visTextBox.Text = Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                                (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                                (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                                " (" + Sch1.SubjectType.Short + ")" + "\n" + Sch1.Auditory;
                                        }
                                        //Если хотя бы что-то не совпало, то записываем элементы через дробь
                                        else
                                        {
                                            visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                         ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                            visTextBox.LineStyle = "None";
                                            visTextBox.FillStyle = "None";

                                            //Снять выделение со всех элементов
                                            visApp.ActiveWindow.DeselectAll();
                                            visApp.ActiveWindow.Select(visTextBox, 2);
                                            visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                            visApp.ActiveWindow.Selection.Move(-(3f / 25.4f), (9f / 25.4f));

                                            visTextBox.Text = Sch1.Auditory + "\n" + Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                                (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                                (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                                " (" + Sch1.SubjectType.Short + ")";

                                            visTextBox = visPage.DrawLine(((currX) / 25.4f), (currY / 25.4f),
                                                        ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                            visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                         ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                            visTextBox.LineStyle = "None";
                                            visTextBox.FillStyle = "None";

                                            //Снять выделение со всех элементов
                                            visApp.ActiveWindow.DeselectAll();
                                            visApp.ActiveWindow.Select(visTextBox, 2);
                                            visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                            visApp.ActiveWindow.Selection.Move((4f / 25.4f), (1f / 25.4f));

                                            visTextBox.Text = Sch2.Spec.ShortInstitute + "-" + Sch2.KursNum.Kurs +
                                                                (Sch2.Stream.Equals("") ? (Sch2.Spec.Diff == "М" ? "5" : "1") : Sch2.Stream) +
                                                                (Sch2.Group.Equals("") ? "1" : Sch2.Group) +
                                                                " (" + Sch2.SubjectType.Short + ")" + "\n" + Sch2.Auditory;
                                        }
                                    }
                                }
                                else
                                {
                                    //Если элемент расписания нашёлся только для одной из недель
                                    if (flgFound1Week || flgFound2Week)
                                    {
                                        //Элемент для первой недели
                                        if (flgFound1Week)
                                        {
                                            visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                         ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                            visTextBox.LineStyle = "None";
                                            visTextBox.FillStyle = "None";

                                            //Снять выделение со всех элементов
                                            visApp.ActiveWindow.DeselectAll();
                                            visApp.ActiveWindow.Select(visTextBox, 2);
                                            visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                            visApp.ActiveWindow.Selection.Move(-(3f / 25.4f), (9f / 25.4f));

                                            if ((Sch1.KursNum == null || Sch1.Spec == null || Sch1.SubjectType == null))
                                            {
                                                visTextBox.Text = "Есть";
                                            }
                                            else
                                            {
                                                visTextBox.Text = Sch1.Auditory + "\n" + Sch1.Spec.ShortInstitute + "-" + Sch1.KursNum.Kurs +
                                                                    (Sch1.Stream.Equals("") ? (Sch1.Spec.Diff == "М" ? "5" : "1") : Sch1.Stream) +
                                                                    (Sch1.Group.Equals("") ? "1" : Sch1.Group) +
                                                                    " (" + Sch1.SubjectType.Short + ")";
                                            }
                                        }

                                        visTextBox = visPage.DrawLine((currX / 25.4f), (currY / 25.4f),
                                                     ((currX + 20f) / 25.4f), ((currY + 20f) / 25.4f));

                                        //Элемент для второй недели
                                        if (flgFound2Week)
                                        {
                                            visTextBox = visPage.DrawRectangle(((currX - 5f) / 25.4f), (currY / 25.4f),
                                                         ((currX + 20f + 5f) / 25.4f), ((currY + 10f) / 25.4f));
                                            visTextBox.LineStyle = "None";
                                            visTextBox.FillStyle = "None";

                                            //Снять выделение со всех элементов
                                            visApp.ActiveWindow.DeselectAll();
                                            visApp.ActiveWindow.Select(visTextBox, 2);
                                            visApp.ActiveWindow.Selection.Rotate(45d, 81);

                                            visApp.ActiveWindow.Selection.Move((4f / 25.4f), (1f / 25.4f));

                                            if (Sch2.KursNum == null || Sch2.Spec == null || Sch2.SubjectType == null)
                                            {
                                                visTextBox.Text = "Есть";
                                            }
                                            else
                                            {
                                                visTextBox.Text = Sch2.Spec.ShortInstitute + "-" + Sch2.KursNum.Kurs +
                                                                    (Sch2.Stream.Equals("") ? (Sch2.Spec.Diff == "М" ? "5" : "1") : Sch2.Stream) +
                                                                    (Sch2.Group.Equals("") ? "1" : Sch2.Group) +
                                                                    " (" + Sch2.SubjectType.Short + ")" + "\n" + Sch2.Auditory;
                                            }
                                        }
                                    }
                                }

                                currX += 20f;
                            }
                        }

                        currY += 20f;
                    }


                    visDoc.SaveAs(Application.StartupPath + "\\Расписания\\" + cmbChoose.Text + ".vsdx");
                }
            }

                visApp.Quit();
                MessageBox.Show("Расписания сформированы!", "Готово!");
                System.Diagnostics.Process.Start("explorer", Application.StartupPath + "\\Расписания\\");
            
        }

        private void btnAll_Click(object sender, EventArgs e)
        {
            onVisioTimeTableForAll();
        }
    }
}
