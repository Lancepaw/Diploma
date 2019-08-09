using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace DataBaseIO
{
    public partial class frmScheduleManagement : Form
    {
        public static int fWidth;
        public static int fHeigth;

        public frmScheduleManagement()
        {
            InitializeComponent();

            fWidth = this.Width;
            fHeigth = this.Height;
        }

        bool flgCombine = false;

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void DispatchGrid()
        {
            int curRow;
            bool HaveLoad;
            bool Both;

            //Очищаем сетку
            dgScheduleManagement.Rows.Clear();
            dgScheduleManagement.Columns.Clear();

            //-----------------------------------------------------------------
            //---------------Фрагмент обозначения рабочей области--------------
            //-----------------------------------------------------------------

            //Задаём количество столбцов
            //оно остаётся неизменным
            //0. Преподаватель
            //1. Курс
            //2. Специальность
            //3. Название дисциплины
            //4. Лекции
            //5. Практ.
            //6. Лаб.
            //7. Курс.пр.
            //8. Практика
            //9. (пусто)
            //10. Примечание

            for (int i = 0; i <= 10; i++)
            {
                dgScheduleManagement.Columns.Add("", "");
            }

            //Создаём строки шапки
            //0. Название университета
            //1. (пусто)
            //2. Слово "Заявка"
            //3. Наименование кафедры и учебного года
            //4. (пусто)
            //5. Семестр и его номер
            //6. (пусто)
            //7. Шапка сетки
            //8. (пусто)
            for (int i = 0; i <= 8; i++)
            {
                dgScheduleManagement.Rows.Add();
            }

            Both = false;
            if (mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum == "-")
            {
                Both = true;
            }

            //Просматриваем каждого преподавателя
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                //Считаем, по умолчанию, что преподаватель не нагружен
                HaveLoad = false;
                //Просматриваем нагрузку
                for (int j = 0; j <= mdlData.colDistribution.Count - 1; j++)
                {
                    if (!Both)
                    {
                        if (!(mdlData.colDistribution[j].Semestr == null))
                        {
                            //Если дисциплина относится ко второму семестру
                            if (mdlData.colDistribution[j].Semestr.Equals(mdlData.colSemestr[cmbSemestrList.SelectedIndex]))
                            {
                                //Если рассматриваемый преподаватель совпадает с
                                //преподавателем, указанным в нагрузке
                                if (mdlData.colLecturer[i].Equals(mdlData.colDistribution[j].Lecturer))
                                {
                                    //Если дисциплина предполагает лекционные часы,
                                    //практические занятия, лабораторные работы,
                                    //курсовой проект, учебную практику, производственную
                                    //практику или дипломную практику, то
                                    //выделяем под неё строчку текста
                                    if (mdlData.NonZeroForDispatchOR(mdlData.colDistribution[j]))
                                    {
                                        if (mdlData.colDistribution[j].flgDispatch)
                                        {
                                            //Если преподаватель что-либо из этого ведёт
                                            //добавляем строку
                                            dgScheduleManagement.Rows.Add();
                                            //Получается, что преподаватель нагружен
                                            HaveLoad = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        //Если рассматриваемый преподаватель совпадает с
                        //преподавателем, указанным в нагрузке
                        if (mdlData.colLecturer[i].Equals(mdlData.colDistribution[j].Lecturer))
                        {
                            //Если дисциплина предполагает лекционные часы,
                            //практические занятия, лабораторные работы,
                            //курсовой проект, учебную практику, производственную
                            //практику или дипломную практику, то
                            //выделяем под неё строчку текста
                            if (mdlData.NonZeroForDispatchOR(mdlData.colDistribution[j]))
                            {
                                if (mdlData.colDistribution[j].flgDispatch)
                                {
                                    //Если преподаватель что-либо из этого ведёт
                                    //добавляем строку
                                    dgScheduleManagement.Rows.Add();
                                    //Получается, что преподаватель нагружен
                                    HaveLoad = true;
                                }
                            }
                        }
                    }
                }

                //Только если у преподавателя есть нагрузка
                if (HaveLoad)
                {
                    //Разделяем преподавателей пустой строчкой
                    dgScheduleManagement.Rows.Add();
                }
            }

            dgScheduleManagement.Rows.Add();
            dgScheduleManagement.Rows.Add();
            dgScheduleManagement.Rows.Add();
            dgScheduleManagement.Rows.Add();
            dgScheduleManagement.Rows.Add();

            //-----------------------------------------------------------------
            //---------------Фрагмент обозначения рабочей области--------------
            //-----------------------------------------------------------------

            //-----------------------------------------------------------------
            //---------------Фрагмент заполнения рабочей области---------------
            //-----------------------------------------------------------------

            //Берём первую строку
            curRow = 0;
            //Вписываем в неё название университета
            dgScheduleManagement[0, curRow].Value = mdlData.UniversityPrefName + " " + 
                mdlData.UniversityName + " " + mdlData.UniversitySuffName;
            //Перешагиваем через одну строку
            curRow += 2;
            //Пишем, что мы формируем заявку
            dgScheduleManagement[0, curRow].Value = "ЗАЯВКА";
            //Идём на следующую строку
            curRow += 1;
            //Пишем наименование кафедры, номер семестра и учебный год
            //к которым относится заявка
            dgScheduleManagement[0, curRow].Value = "Кафедры " + "\"" + mdlData.DepartmentName + "\" в диспетчерскую " +
                                                    "на проведение занятий " + mdlData.colSemestr[cmbSemestrList.SelectedIndex].About +
                                                    " " + mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear + " учебного года";
            //Перешагиваем через одну строку
            curRow += 2;
            //Прописываем в строке название "Семестр"
            dgScheduleManagement[0, curRow].Value = "Семестр: ";
            //Пишем номер (код) семестра
            dgScheduleManagement[1, curRow].Value = mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum;
            //Пишем заготовку для заседания кафедры
            //(потом перенести в переменную, тягаемую из таблицы Параметры)
            dgScheduleManagement[10, curRow].Value = "Не назначать учебные часы для ВСЕХ преподавателей кафедры " +
                "по понедельникам обеих недель с 15:00 до 16:30 (заседание кафедры)";
            //Перешагиваем через одну строку
            curRow += 2;
            //Формируем шапку таблицы-заявки
            dgScheduleManagement[0, curRow].Value = "Преподаватель";
            dgScheduleManagement[1, curRow].Value = "Курс";
            dgScheduleManagement[2, curRow].Value = "Специальность";
            dgScheduleManagement[3, curRow].Value = "Название дисциплины";
            dgScheduleManagement[4, curRow].Value = "Лекция";
            dgScheduleManagement[5, curRow].Value = "Практ.зан.";
            dgScheduleManagement[6, curRow].Value = "Лаб.раб.";
            dgScheduleManagement[7, curRow].Value = "Курс.пр.";
            dgScheduleManagement[8, curRow].Value = "Практика";
            dgScheduleManagement[10, curRow].Value = "Примечания";
            //Перешагиваем на следующую строку
            curRow += 1;
            //Пишем комментарий для диспетчеров о том, как воспринимать
            //пожелания преподавателей и аудитории
            //(позже занести в переменную, хранимую в таблице Параметры)
            dgScheduleManagement[10, curRow].Value = "Пожелания по распределению и аудиториям: " +
                "А) комментарии напротив дисциплин - к дисциплинам; Б) комментарии напротив пустых " +
                "строк - к преподавателям, упомянутым над комментарием";
            //Перешагиваем на следующую строку
            curRow += 1;

            //Просматриваем каждого преподавателя
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                //По умолчанию считаем, что у преподавателя нет нагрузки
                HaveLoad = false;
                //Просматриваем нагрузку
                for (int j = 0; j <= mdlData.colDistribution.Count - 1; j++)
                {
                    if (!Both)
                    {
                        if (!(mdlData.colDistribution[j].Semestr == null))
                        {
                            //Если дисциплина относится к выбранному семестру
                            if (mdlData.colDistribution[j].Semestr.Equals(mdlData.colSemestr[cmbSemestrList.SelectedIndex]))
                            {
                                //Если рассматриваемый преподаватель совпадает с
                                //преподавателем, указанным в нагрузке
                                if (mdlData.colLecturer[i].Equals(mdlData.colDistribution[j].Lecturer))
                                {
                                    //Если дисциплина предполагает лекционные часы,
                                    //практические занятия, лабораторные работы,
                                    //курсовой проект, учебную практику, производственную
                                    //практику или дипломную практику, то
                                    //выделяем под неё строчку текста
                                    if (mdlData.NonZeroForDispatchOR(mdlData.colDistribution[j]))
                                    {
                                        if (mdlData.colDistribution[j].flgDispatch)
                                        {
                                            if (!HaveLoad)
                                            {
                                                //Печатаем фамилию, имя, отчество преподавателя
                                                dgScheduleManagement[0, curRow].Value = mdlData.colLecturer[i].FIO + " (" +
                                                    mdlData.colLecturer[i].Duty.Short + ")";
                                                //Получается, что у преподавателя есть нагрузка
                                                HaveLoad = true;
                                            }
                                            //Печатаем курс
                                            dgScheduleManagement[1, curRow].Value = mdlData.colDistribution[j].KursNum.Kurs;
                                            //Печатаем специальность
                                            dgScheduleManagement[2, curRow].Value = mdlData.colDistribution[j].Speciality.ShortUpravlenie
                                                + " (" + mdlData.colDistribution[j].Speciality.ShortInstitute + ")";
                                            //Печатаем название дисциплины
                                            dgScheduleManagement[3, curRow].Value = mdlData.colDistribution[j].Subject.Subject;
                                            //Лекционные часы
                                            dgScheduleManagement[4, curRow].Value = mdlData.colDistribution[j].Lecture;
                                            //Практические занятия в часах
                                            dgScheduleManagement[5, curRow].Value = mdlData.colDistribution[j].Practice;
                                            //Лабораторные работы в часах
                                            dgScheduleManagement[6, curRow].Value = mdlData.colDistribution[j].LabWork;
                                            //Курсовой проект в часах
                                            dgScheduleManagement[7, curRow].Value = mdlData.colDistribution[j].KursProject;
                                            //Практика в часах
                                            //Либо записываем преддипломную практику
                                            if (mdlData.colDistribution[j].PreDiplomaPractice > 0)
                                            {
                                                dgScheduleManagement[8, curRow].Value = mdlData.colDistribution[j].PreDiplomaPractice;
                                            }
                                            //Либо записываем учебную практику
                                            if (mdlData.colDistribution[j].TutorialPractice > 0)
                                            {
                                                dgScheduleManagement[8, curRow].Value = mdlData.colDistribution[j].TutorialPractice;
                                            }
                                            //Либо записываем производственную практику
                                            if (mdlData.colDistribution[j].ProducingPractice > 0)
                                            {
                                                dgScheduleManagement[8, curRow].Value = mdlData.colDistribution[j].ProducingPractice;
                                            }
                                            //Если поле практики так и осталось пустым, то
                                            //заполняем его нулём
                                            if (dgScheduleManagement[8, curRow].Value == null)
                                            {
                                                dgScheduleManagement[8, curRow].Value = 0;
                                            }
                                            //
                                            dgScheduleManagement[10, curRow].Value = mdlData.colDistribution[j].Text;
                                            //Переходим к следующей строке
                                            curRow += 1;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        //Если рассматриваемый преподаватель совпадает с
                        //преподавателем, указанным в нагрузке
                        if (mdlData.colLecturer[i].Equals(mdlData.colDistribution[j].Lecturer))
                        {
                            //Если дисциплина предполагает лекционные часы,
                            //практические занятия, лабораторные работы,
                            //курсовой проект, учебную практику, производственную
                            //практику или дипломную практику, то
                            //выделяем под неё строчку текста
                            if (mdlData.NonZeroForDispatchOR(mdlData.colDistribution[j]))
                            {
                                if (mdlData.colDistribution[j].flgDispatch)
                                {
                                    if (!HaveLoad)
                                    {
                                        //Печатаем фамилию, имя, отчество преподавателя
                                        dgScheduleManagement[0, curRow].Value = mdlData.colLecturer[i].FIO + " (" +
                                                    mdlData.colLecturer[i].Duty.Short + ")";
                                        //Получается, что у преподавателя есть нагрузка
                                        HaveLoad = true;
                                    }
                                    //Печатаем курс
                                    dgScheduleManagement[1, curRow].Value = mdlData.colDistribution[j].KursNum.Kurs;
                                    //Печатаем специальность
                                    dgScheduleManagement[2, curRow].Value = mdlData.colDistribution[j].Speciality.ShortUpravlenie
                                        + " (" + mdlData.colDistribution[j].Speciality.ShortInstitute + ")";
                                    //Печатаем название дисциплины
                                    dgScheduleManagement[3, curRow].Value = mdlData.colDistribution[j].Subject.Subject;
                                    //Лекционные часы
                                    dgScheduleManagement[4, curRow].Value = mdlData.colDistribution[j].Lecture;
                                    //Практические занятия в часах
                                    dgScheduleManagement[5, curRow].Value = mdlData.colDistribution[j].Practice;
                                    //Лабораторные работы в часах
                                    dgScheduleManagement[6, curRow].Value = mdlData.colDistribution[j].LabWork;
                                    //Курсовой проект в часах
                                    dgScheduleManagement[7, curRow].Value = mdlData.colDistribution[j].KursProject;
                                    //Практика в часах
                                    //Либо записываем преддипломную практику
                                    if (mdlData.colDistribution[j].PreDiplomaPractice > 0)
                                    {
                                        dgScheduleManagement[8, curRow].Value = mdlData.colDistribution[j].PreDiplomaPractice;
                                    }
                                    //Либо записываем учебную практику
                                    if (mdlData.colDistribution[j].TutorialPractice > 0)
                                    {
                                        dgScheduleManagement[8, curRow].Value = mdlData.colDistribution[j].TutorialPractice;
                                    }
                                    //Либо записываем производственную практику
                                    if (mdlData.colDistribution[j].ProducingPractice > 0)
                                    {
                                        dgScheduleManagement[8, curRow].Value = mdlData.colDistribution[j].ProducingPractice;
                                    }
                                    //Если поле практики так и осталось пустым, то
                                    //заполняем его нулём
                                    if (dgScheduleManagement[8, curRow].Value == null)
                                    {
                                        dgScheduleManagement[8, curRow].Value = 0;
                                    }
                                    //
                                    dgScheduleManagement[10, curRow].Value = mdlData.colDistribution[j].Text;
                                    //Переходим к следующей строке
                                    curRow += 1;
                                }
                            }
                        }
                    }
                }

                //Только, если у преподавателя есть нагрузка
                if (HaveLoad)
                {
                    //Пишем примечание конкретного преподавателя
                    dgScheduleManagement[10, curRow].Value = mdlData.colLecturer[i].Text;
                    //Разделяем преподавателей пустой строчкой
                    curRow += 1;
                }
            }

            dgScheduleManagement[2, curRow].Value = "Заведующий кафедрой УиЗИ";
            dgScheduleManagement[7, curRow].Value = "/ Л.А. Баранов /";

            curRow += 1;

            dgScheduleManagement[2, curRow].Value = "Директор ИТТСУ";
            dgScheduleManagement[7, curRow].Value = "/ П.Ф. Бестемьянов /";

            curRow += 1;

            dgScheduleManagement[2, curRow].Value = "Первый зам. директора ИТТСУ - " +
                "начальник учебного отдела";
            dgScheduleManagement[7, curRow].Value = "/ Е.В. Сердобинцев /";

            curRow += 1;

            dgScheduleManagement[2, curRow].Value = "Декан вечернего факультета";
            dgScheduleManagement[7, curRow].Value = "/ В.Ф. Ковальский /";

            curRow += 1;

            dgScheduleManagement[2, curRow].Value = "Директор ИУИТ";
            dgScheduleManagement[7, curRow].Value = "/ С.П. Вакуленко /";

            //-----------------------------------------------------------------
            //---------------Фрагмент заполнения рабочей области---------------
            //----------------------------------------------------------------- 
        }

        private void frmScheduleManagement_Load(object sender, EventArgs e)
        {
            //Очищаем сетку
            dgScheduleManagement.Rows.Clear();
            dgScheduleManagement.Columns.Clear();
            
            FillSemestrList();
            FillWorkYearList();

            optCombine.Checked = false;
            optHoured.Checked = false;
            optMain.Checked = true;

            //Делаем невидимыми нуль-строку и нуль-столбец
            dgScheduleManagement.ColumnHeadersVisible = false;
            dgScheduleManagement.RowHeadersVisible = false;

            //Выгрузка в документы (печатная форма)
            cmbForm.Items.Add("Не выбрано"); //0
            cmbForm.Items.Add("Заявка в диспетчерскую Word"); //1
            cmbForm.Items.Add("Плановая нагрузка Excel"); //2
            cmbForm.Items.Add("Выполненная нагрузка Excel"); //3
            cmbForm.Items.Add("Закреплённые дисциплины Excel"); //4
            cmbForm.Items.Add("Режим в Excel"); //5

            chkPlan.Checked = true;

            //Выгрузка в сетку (экранная форма)
            cmbGrid.Items.Add("Не выбрано"); //0
            cmbGrid.Items.Add("Заявка в диспетчерскую"); //1
            cmbGrid.Items.Add("Распределение для Зав. кафедрой"); //2
            cmbGrid.Items.Add("Плановое распределение в Уч. управление"); //3
            cmbGrid.Items.Add("Фактическое распределение в Уч. управление"); //4
            cmbGrid.Items.Add("Сведения по Курс. проектам"); //5
            cmbGrid.Items.Add("Сведения по выпускникам (ВКР)"); //6
            cmbGrid.Items.Add("Показатели нагрузки"); //7
            cmbGrid.Items.Add("Равномерно распределяемая нагрузка"); //8

            Resize += new EventHandler(frmScheduleManagement_Resize);

            //Выбрать только плановую почасовую
            for (int i = 0; i < mdlData.colHouredDistribution.Count; i++)
            {
                if (mdlData.colHouredDistribution[i].Doubler == null)
                {
                    mdlData.colPlanHouredDistribution.Add(mdlData.colHouredDistribution[i]);
                }
            }

            //Получаем плановую комбинированную нагрузку
            mdlData.toCombineDistribution(mdlData.colDistribution, mdlData.colPlanCombineDistribution,
                mdlData.colPlanHouredDistribution);
        }

        void frmScheduleManagement_Resize(object sender, EventArgs e)
        {
            if (this.Width >= fWidth & this.Height >= fHeigth)
            {
                dgScheduleManagement.Width = this.Width - 40;

                btnClose.Top = this.Height - 50 - btnClose.Height;
                btnClose.Left = this.Width - 30 - btnClose.Width;

                btnExcel.Top = btnClose.Top;
                btnExcel.Left = btnClose.Left - 10 - btnExcel.Width;

                cmbForm.Top = btnExcel.Top - 10 - cmbForm.Height;
                cmbForm.Left = btnExcel.Left - 50 - cmbForm.Width;

                btnForm.Left = cmbForm.Left;
                btnForm.Top = btnExcel.Top;

                lblForm.Left = cmbForm.Left;
                lblForm.Top = cmbForm.Top - 10 - lblForm.Height;

                cmbGrid.Top = cmbForm.Top;
                cmbGrid.Left = dgScheduleManagement.Left;

                lblGrid.Top = lblForm.Top;
                lblGrid.Left = cmbGrid.Left;

                dgScheduleManagement.Height = (lblForm.Top - 10) - dgScheduleManagement.Top;
            }
            else
            {
                this.Width = fWidth;
                this.Height = fHeigth;
            }
        }

        private void FillSemestrList()
        {
            int NumFix = 0;
            NumFix = cmbSemestrList.SelectedIndex;
            //Очищаем список
            cmbSemestrList.Items.Clear();

            //Заполняем комбо-список семестрами
            for (int i = 0; i <= mdlData.colSemestr.Count - 1; i++)
            {
                cmbSemestrList.Items.Add(mdlData.colSemestr[i].Code + ". " + mdlData.colSemestr[i].About);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbSemestrList.SelectedIndex = 0;
            }
            else
            {
                cmbSemestrList.SelectedIndex = NumFix;
            }
        }

        private void FillWorkYearList()
        {
            int NumFix = 0;
            NumFix = cmbWorkYearList.SelectedIndex;
            //Очищаем список
            cmbWorkYearList.Items.Clear();

            //Заполняем комбо-список семестрами
            for (int i = 0; i <= mdlData.colWorkYear.Count - 1; i++)
            {
                cmbWorkYearList.Items.Add(mdlData.colWorkYear[i].Code + ". " + mdlData.colWorkYear[i].WorkYear);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbWorkYearList.SelectedIndex = cmbWorkYearList.Items.Count - 2;
            }
            else
            {
                cmbWorkYearList.SelectedIndex = NumFix;
            }
        }

        private void cmbSemestrList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void optMain_CheckedChanged(object sender, EventArgs e)
        {
            flgCombine = false;
            dgScheduleManagement.Rows.Clear();
            dgScheduleManagement.Columns.Clear();
        }

        private void optHoured_CheckedChanged(object sender, EventArgs e)
        {
            flgCombine = false;
            dgScheduleManagement.Rows.Clear();
            dgScheduleManagement.Columns.Clear();
        }

        private void optCombine_CheckedChanged(object sender, EventArgs e)
        {
            flgCombine = true;
            dgScheduleManagement.Rows.Clear();
            dgScheduleManagement.Columns.Clear();
        }

        private void gridIntoExcel()
        {
            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;

            //Только если сетка создана имеет смысл выгружать что-то
            //в Excel
            if (dgScheduleManagement.RowCount > 0 &
                dgScheduleManagement.ColumnCount > 0)
            {
                try
                {
                    //Создаём новое Excel приложение
                    Excel.Application ObjExcel = new Excel.Application();
                    Excel.Workbook ObjWorkBook;
                    Excel.Worksheet ObjWorkSheet;

                    //Книга
                    ObjWorkBook = ObjExcel.Workbooks.Add(Missing.Value);
                    //Таблица
                    ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

                    for (int i = 0; i <= dgScheduleManagement.RowCount - 1; i++)
                    {
                        for (int j = 0; j <= dgScheduleManagement.ColumnCount - 1; j++)
                        {
                            if (!(dgScheduleManagement[j, i].Value == null))
                            {
                                ObjWorkSheet.Cells[i + 1, j + 1] = dgScheduleManagement[j, i].Value.ToString();
                            }
                        }
                    }

                    ObjExcel.UserControl = true;

                    ObjWorkBook.SaveCopyAs(Application.StartupPath + @"\Ведомость плановая, сетка " + 
                        DateTime.Now.Date.ToString("yyyyMMdd") + " " + 
                        DateTime.Now.TimeOfDay.ToString("hhmmss") + ".xlsx");
                    ObjWorkBook.Close(false, "", Missing.Value);

                    ObjExcel.Quit();
                }
                catch
                {
                    MessageBox.Show("Возможно на этом компьютере присутствует проблема совместимости с MS Exсel." +
                                    " Попробуйте установить версию 2007 и выше.");
                }
            }
            else
            {
                btnExcel.Enabled = false;
            }
        }

        private void UpravlenieExcelCore()
        {
            int curLect;

            int sumLecI; int sumLecII;

            int sumExamI; int sumExamII;

            int sumCredI; int sumCredII;

            int sumTutI; int sumTutII;

            int sumLabI; int sumLabII;

            int sumPracI; int sumPracII;

            int sumRefI; int sumRefII;

            int sumIndI; int sumIndII;

            int sumKRAPKI; int sumKRAPKII;

            int sumKursPrI; int sumKursPrII;

            int sumDiplI; int sumDiplII;

            int sumTutPrI; int sumTutPrII;

            int sumPreDipI; int sumPreDipII;

            int sumGAKI; int sumGAKII;

            int sumPostGrI; int sumPostGrII;

            int sumVisI; int sumVisII;

            int sumMagI; int sumMagII;

            int countStud, countWeight;

            double[] rateParams = new double[4];
            double assist;
            double hitutor;
            double proff;
            double lecturer;
            double sumRate;

            IList<clsDistribution> coll = null;

            string strSumI = "=";
            string strSumII = "=";
            string strSemI;
            string strSemII;
            string strSum = "=";

            bool accessFlg = false;

            rateParams = countRates();

            assist = rateParams[0];
            lecturer = rateParams[1];
            hitutor = rateParams[2];
            proff = rateParams[3];

            sumRate = proff + hitutor + lecturer + assist;

            coll = mdlData.colCombineDistribution;

            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;

            try
            {
                //Создаём новое Excel приложение
                Excel.Application ObjExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook;
                Excel.Worksheet ObjWorkSheet;

                ObjExcel.Visible = true;

                //Книга
                ObjWorkBook = ObjExcel.Workbooks.Add(Missing.Value);
                //Таблица
                ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
                //Альбомная ориентация страницы
                ObjWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                //Книжная ориентация страницы
                //ObjWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                //В высоту разместить на одной странице
                ObjWorkSheet.PageSetup.FitToPagesTall = 1;
                //В ширину разместить на одной странице
                ObjWorkSheet.PageSetup.FitToPagesWide = 1;

                //-----------Формируем заголовок таблицы

                //Задаём диапазон для ячеек, подлежащих форматированию
                //1-я строка, с А по AA
                var cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(1, 1),
                    mdlData.ExcelCellTranslator(1, 27));
                //Выделяем ячейки диапазона
                cells.Select();
                //Объединяем ячейки диапазона
                cells.Merge(true);
                //Выравнивание в объединённой ячейке по центру
                cells.HorizontalAlignment = Excel.Constants.xlCenter;
                //Записываем текст в объединённую ячейку
                //(считается по первой А1)
                        
                //Если плановая
                cells.Cells[1, 1] = "Сведения о распределении штатной нагрузки кафедры на учебный год";

                //Задаём диапазон для ячеек, подлежащих форматированию
                //2-я строка, с А по AA
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(2, 1),
                    mdlData.ExcelCellTranslator(2, 27));
                //Выделяем ячейки диапазона
                cells.Select();
                //Объединяем ячейки диапазона
                cells.Merge(true);
                //Выравнивание в объединённой ячейке по центру
                cells.HorizontalAlignment = Excel.Constants.xlCenter;
                //Записываем текст в объединённую ячейку
                //(считается по первой А2)
                cells.Cells[1, 1] = "Кафедра \"Управление и защита информации\" " +
                    mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear +
                    " учебный год";

                //-----------Формируем заголовок таблицы

                //-----------Формируем шапку таблицы

                for (int i = 1; i <= 27; i++)
                {
                    //Выбираем диапазон
                    cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4, i),
                        mdlData.ExcelCellTranslator(5, i));
                    //Выделяем ячейки диапазона
                    cells.Select();
                    //Объединяем ячейки диапазона
                    cells.Merge();
                    //Горизонтальное выравнивание по центру в ячейках
                    cells.HorizontalAlignment = Excel.Constants.xlCenter;
                    //Вертикальное выравнивание по центру в ячейках
                    cells.VerticalAlignment = Excel.Constants.xlCenter;
                    //Задаём границы
                    cells.Borders.Weight = 2;
                    //Перенос по словам без сокрытия под ячейками
                    cells.WrapText = true;
                    //Назначаем высоту шапки
                    cells.RowHeight = (68.25f + 66) / 2;

                    switch (i)
                    {
                        case 1:
                            cells.Cells[1, 1] = "№ п/п";
                            cells.ColumnWidth = 3.86f;
                            cells.Orientation = Excel.XlOrientation.xlHorizontal;
                            break;
                        case 2:
                            cells.Cells[1, 1] = "Фамилия, Имя, Отчество";
                            cells.ColumnWidth = 27.57f;
                            cells.Orientation = Excel.XlOrientation.xlHorizontal;
                            break;
                        case 3:
                            cells.Cells[1, 1] = "Ставка";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 4:
                            cells.Cells[1, 1] = "Ученая степень, звание";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 5:
                            cells.Cells[1, 1] = "Лекции по семестрам";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 6:
                            cells.Cells[1, 1] = "Всего лекций";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 7:
                            cells.Cells[1, 1] = "Экзамены";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 8:
                            cells.Cells[1, 1] = "Зачеты";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 9:
                            cells.Cells[1, 1] = "ПК";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 10:
                            cells.Cells[1, 1] = "Консультации";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 11:
                            cells.Cells[1, 1] = "Практические занятия";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 12:
                            cells.Cells[1, 1] = "Домашние задания и рефераты";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 13:
                            cells.Cells[1, 1] = "Текущая аттестация";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 14:
                            cells.Cells[1, 1] = "Индивидуальные занятия";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 15:
                            cells.Cells[1, 1] = "Контрольные работы";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 16:
                            cells.Cells[1, 1] = "Курсовой проект, курсовая работа";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 17:
                            cells.Cells[1, 1] = "Дипломный проект";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 18:
                            cells.Cells[1, 1] = "Учебная практика";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 19:
                            cells.Cells[1, 1] = "Преддипломная и производственная практика";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 20:
                            cells.Cells[1, 1] = "ГЭК";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 21:
                            cells.Cells[1, 1] = "Приёмная комиссия";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 22:
                            cells.Cells[1, 1] = "Лабораторные работы";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 23:
                            cells.Cells[1, 1] = "Аспирантура";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 24:
                            cells.Cells[1, 1] = "Посещение занятий";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 25:
                            cells.Cells[1, 1] = "Другие виды занятий";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 26:
                            cells.UnMerge();
                            cells.Borders.Weight = 2;
                            cells.Cells[1, 1] = "I сем.";
                            cells.Cells[2, 1] = "II сем.";
                            cells.Orientation = Excel.XlOrientation.xlHorizontal;
                            break;
                        case 27:
                            cells.Merge();
                            cells.Cells[1, 1] = "Всего за год";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                    }
                }

                //-----------Формируем шапку таблицы

                //-----------Заполняем таблицу данными

                curLect = 1;

                for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
                {
                    accessFlg = (mdlData.colLecturer[i].Rate > 0);

                    //Если прошёл отбор
                    if (accessFlg)
                    {
                        sumLecI = 0; sumLecII = 0;

                        sumExamI = 0; sumExamII = 0;

                        sumCredI = 0; sumCredII = 0;

                        sumTutI = 0; sumTutII = 0;

                        sumLabI = 0; sumLabII = 0;

                        sumPracI = 0; sumPracII = 0;

                        sumRefI = 0; sumRefII = 0;

                        sumIndI = 0; sumIndII = 0;

                        sumKRAPKI = 0; sumKRAPKII = 0;

                        sumKursPrI = 0; sumKursPrII = 0;

                        sumDiplI = 0; sumDiplII = 0;

                        sumTutPrI = 0; sumTutPrII = 0;

                        sumPreDipI = 0; sumPreDipII = 0;

                        sumGAKI = 0; sumGAKII = 0;

                        sumPostGrI = 0; sumPostGrII = 0;

                        sumVisI = 0; sumVisII = 0;

                        sumMagI = 0; sumMagII = 0;

                        //Просматриваем нагрузку
                        for (int j = 0; j <= coll.Count - 1; j++)
                        {
                            //Если строка не исключена из расчёта нагрузки
                            if (!coll[j].flgExclude)
                            {
                                //Если указан преподаватель для строки нагрузки
                                if (!(coll[j].Lecturer == null))
                                {
                                    if (coll[j].Lecturer.Equals(mdlData.colLecturer[i]))
                                    {
                                        //Если строка нагрузки первого семестра
                                        if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                        {
                                            //Данные для столбца Е (лекции по семестрам)
                                            sumLecI += coll[j].Lecture;
                                            //Данные для столбца F суммируются по I и II семестру
                                            //(всего лекций)

                                            //Данные для столбца G (экзамены)
                                            sumExamI += coll[j].Exam;
                                            //Данные для столбца H (зачёты)
                                            sumCredI += coll[j].Credit;
                                            //Данные для стобца I (ПК)

                                            //Данные для столбца J (Консультации)
                                            sumTutI += coll[j].Tutorial;
                                            
                                            sumLabI += coll[j].LabWork;
                                            //Данные для столбца K
                                            sumPracI += coll[j].Practice;
                                            //
                                            sumRefI += coll[j].RefHomeWork;
                                            sumIndI += coll[j].IndividualWork;
                                            sumKRAPKI += coll[j].KRAPK;
                                            sumKursPrI += coll[j].KursProject;
                                            sumDiplI += coll[j].DiplomaPaper;
                                            sumTutPrI += coll[j].TutorialPractice;
                                            sumPreDipI += coll[j].PreDiplomaPractice +
                                                coll[j].ProducingPractice;
                                            sumGAKI += coll[j].GAK;
                                            sumPostGrI += coll[j].PostGrad;
                                            sumVisI += coll[j].Visiting;
                                            sumMagI += coll[j].Magistry;
                                        }

                                        if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                        {
                                            sumLecII += coll[j].Lecture;
                                            sumExamII += coll[j].Exam;
                                            sumCredII += coll[j].Credit;
                                            sumTutII += coll[j].Tutorial;
                                            sumLabII += coll[j].LabWork;
                                            sumPracII += coll[j].Practice;
                                            sumRefII += coll[j].RefHomeWork;
                                            sumIndII += coll[j].IndividualWork;
                                            sumKRAPKII += coll[j].KRAPK;
                                            sumKursPrII += coll[j].KursProject;
                                            sumDiplII += coll[j].DiplomaPaper;
                                            sumTutPrII += coll[j].TutorialPractice;
                                            sumPreDipII += coll[j].PreDiplomaPractice +
                                                coll[j].ProducingPractice;
                                            sumGAKII += coll[j].GAK;
                                            sumPostGrII += coll[j].PostGrad;
                                            sumVisII += coll[j].Visiting;
                                            sumMagII += coll[j].Magistry;
                                        }

                                        accessFlg = true;
                                    }
                                }
                                //Если преподаватель не указан,
                                else
                                {
                                    //а нагрузка равномерно распределяемая
                                    if (coll[j].flgDistrib)
                                    {
                                        countWeight = 0;
                                        countStud = 0;

                                        for (int k = 0; k <= mdlData.colStudents.Count - 1; k++)
                                        {
                                            if (mdlData.colStudents[k].flgPlan)
                                            {
                                                //Если рассматриваемый преподаватель - руководитель студента
                                                //И если студент на том же курсе, где и дисциплина
                                                //И специальность студента должна соответствовать специальности нагрузки
                                                if (mdlData.colStudents[k].Lect.Equals(mdlData.colLecturer[i])
                                                    & mdlData.colStudents[k].KursNum.Equals(coll[j].KursNum)
                                                    & mdlData.colStudents[k].Speciality.Equals(coll[j].Speciality))
                                                {
                                                    countWeight += coll[j].Weight;
                                                    countStud++;
                                                }
                                            }
                                        }

                                        //
                                        if (countWeight > 0)
                                        {
                                            accessFlg = true;

                                            if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                            {
                                                if (coll[j].PreDiplomaPractice > 0 || 
                                                    coll[j].ProducingPractice > 0)
                                                {
                                                    sumPreDipI += countWeight;
                                                }

                                                if (coll[j].TutorialPractice > 0)
                                                {
                                                    sumTutPrI += countWeight;
                                                }

                                                if (coll[j].DiplomaPaper > 0)
                                                {
                                                    sumDiplI += countWeight;
                                                }                                                
                                                
                                                if (coll[j].Magistry > 0)
                                                {
                                                    sumMagI += countWeight;
                                                }
                                            }

                                            if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                            {
                                                if (coll[j].PreDiplomaPractice > 0 || 
                                                    coll[j].ProducingPractice > 0)
                                                {
                                                    sumPreDipII += countWeight;
                                                }

                                                if (coll[j].TutorialPractice > 0)
                                                {
                                                    sumTutPrII += countWeight;
                                                }

                                                if (coll[j].DiplomaPaper > 0)
                                                {
                                                    sumDiplII += countWeight;
                                                }                                                
                                                
                                                if (coll[j].Magistry > 0)
                                                {
                                                    sumMagII += countWeight;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (accessFlg)
                        {
                            strSemI = "=";
                            strSemII = "=";
                            for (int k = 5; k <= 25; k++)
                            {
                                if (k != 6)
                                {
                                    strSemI += mdlData.ExcelCellTranslator(6 + 2 * (curLect - 1), k) + "+";
                                    strSemII += mdlData.ExcelCellTranslator(7 + 2 * (curLect - 1), k) + "+";
                                }
                            }
                            strSemI = strSemI.Substring(0, strSemI.Length - 1);
                            strSemII = strSemII.Substring(0, strSemII.Length - 1);

                            for (int k = 1; k <= 30; k++)
                            {
                                //Выбираем диапазон
                                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(6 + 2 * (curLect - 1), k),
                                    mdlData.ExcelCellTranslator(7 + 2 * (curLect - 1), k));
                                //Задаём границы
                                cells.Borders.Weight = 2;

                                switch (k)
                                {
                                    //Номер по порядку
                                    case 1:
                                        cells.Merge();
                                        cells.Cells[1, 1] = curLect.ToString();
                                        //Горизонтальное выравнивание по центру в ячейках
                                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                        //Вертикальное выравнивание по центру в ячейках
                                        cells.VerticalAlignment = Excel.Constants.xlBottom;
                                        break;
                                    //Фамилия, Имя, Отчество
                                    case 2:
                                        cells.Merge();
                                        cells.Cells[1, 1] = mdlData.colLecturer[i].FIO;
                                        //Горизонтальное выравнивание по центру в ячейках
                                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                        //Вертикальное выравнивание по центру в ячейках
                                        cells.VerticalAlignment = Excel.Constants.xlCenter;
                                        cells.WrapText = true;
                                        break;
                                    //Ставка
                                    case 3:
                                        cells.Merge();
                                        cells.Cells[1, 1] = mdlData.colLecturer[i].Rate.ToString("0.00");
                                        //Горизонтальное выравнивание по центру в ячейках
                                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                        //Вертикальное выравнивание по центру в ячейках
                                        cells.VerticalAlignment = Excel.Constants.xlCenter;
                                        cells.WrapText = true;
                                        break;
                                    //Учёная степень, звание
                                    case 4:
                                        cells.Merge();
                                        cells.Cells[1, 1] = mdlData.colLecturer[i].Degree.Short + ", " +
                                            mdlData.colLecturer[i].Duty.Short;
                                        //Горизонтальное выравнивание по центру в ячейках
                                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                        //Вертикальное выравнивание по центру в ячейках
                                        cells.VerticalAlignment = Excel.Constants.xlCenter;
                                        cells.WrapText = true;
                                        break;
                                    //Лекции по семестрам
                                    case 5:
                                        cells.Cells[1, 1] = sumLecI.ToString();
                                        cells.Cells[2, 1] = sumLecII.ToString();
                                        break;
                                    //Всего лекций
                                    case 6:
                                        cells.Merge();
                                        //Горизонтальное выравнивание по центру в ячейках
                                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                        //Вертикальное выравнивание по центру в ячейках
                                        cells.VerticalAlignment = Excel.Constants.xlCenter;
                                        cells.Cells[1, 1] = (sumLecI + sumLecII).ToString();
                                        break;
                                    //Экзамены
                                    case 7:
                                        cells.Cells[1, 1] = sumExamI.ToString();
                                        cells.Cells[2, 1] = sumExamII.ToString();
                                        break;
                                    //Зачёты
                                    case 8:
                                        cells.Cells[1, 1] = sumCredI.ToString();
                                        cells.Cells[2, 1] = sumCredII.ToString();
                                        break;
                                    //ПК
                                    case 9:
                                        cells.Cells[1, 1] = 0.ToString();
                                        cells.Cells[2, 1] = 0.ToString();
                                        break;
                                    //Консультации
                                    case 10:
                                        cells.Cells[1, 1] = sumTutI.ToString();
                                        cells.Cells[2, 1] = sumTutII.ToString();
                                        break;
                                    //Практические занятия
                                    case 11:
                                        cells.Cells[1, 1] = sumPracI.ToString();
                                        cells.Cells[2, 1] = sumPracII.ToString();
                                        break;
                                    //Домашние задания и рефераты
                                    case 12:
                                        cells.Cells[1, 1] = sumRefI.ToString();
                                        cells.Cells[2, 1] = sumRefII.ToString();
                                        break;
                                    //Текущая аттестация
                                    case 13:
                                        cells.Cells[1, 1] = 0.ToString();
                                        cells.Cells[2, 1] = 0.ToString();
                                        break;
                                    //Индивидуальные занятия
                                    case 14:
                                        cells.Cells[1, 1] = sumIndI.ToString();
                                        cells.Cells[2, 1] = sumIndII.ToString();
                                        break;
                                    //Контрольные работы
                                    case 15:
                                        cells.Cells[1, 1] = sumKRAPKI.ToString();
                                        cells.Cells[2, 1] = sumKRAPKII.ToString();
                                        break;
                                    //Курсовой проект, курсовая работа
                                    case 16:
                                        cells.Cells[1, 1] = sumKursPrI.ToString();
                                        cells.Cells[2, 1] = sumKursPrII.ToString();
                                        break;
                                    //Дипломный проект
                                    case 17:
                                        cells.Cells[1, 1] = sumDiplI.ToString();
                                        cells.Cells[2, 1] = sumDiplII.ToString();
                                        break;
                                    //Учебная практика
                                    case 18:
                                        cells.Cells[1, 1] = sumTutPrI.ToString();
                                        cells.Cells[2, 1] = sumTutPrII.ToString();
                                        break;
                                    //Преддипломная и производственная практика
                                    case 19:
                                        cells.Cells[1, 1] = sumPreDipI.ToString();
                                        cells.Cells[2, 1] = sumPreDipII.ToString();
                                        break;
                                    //ГЭК
                                    case 20:
                                        cells.Cells[1, 1] = sumGAKI.ToString();
                                        cells.Cells[2, 1] = sumGAKII.ToString();
                                        break;
                                    //Приёмная комиссия
                                    case 21:
                                        cells.Cells[1, 1] = 0.ToString();
                                        cells.Cells[2, 1] = 0.ToString();
                                        break;
                                    //Лабораторные работы
                                    case 22:
                                        cells.Cells[1, 1] = sumLabI.ToString();
                                        cells.Cells[2, 1] = sumLabII.ToString();
                                        break;
                                    //Аспирантура
                                    case 23:
                                        cells.Cells[1, 1] = sumPostGrI.ToString();
                                        cells.Cells[2, 1] = sumPostGrII.ToString();
                                        break;
                                    //Посещение занятий
                                    case 24:
                                        cells.Cells[1, 1] = sumVisI.ToString();
                                        cells.Cells[2, 1] = sumVisII.ToString();
                                        break;
                                    //Другие виды занятий
                                    case 25:
                                        cells.Cells[1, 1] = sumMagI.ToString();
                                        cells.Cells[2, 1] = sumMagII.ToString();
                                        break;
                                    //
                                    case 26:
                                        cells.Cells[1, 1] = strSemI;
                                        cells.Cells[2, 1] = strSemII;
                                        break;
                                    case 27:
                                        cells.Merge();
                                        //Горизонтальное выравнивание по центру в ячейках
                                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                        //Вертикальное выравнивание по центру в ячейках
                                        cells.VerticalAlignment = Excel.Constants.xlCenter;
                                        cells.Cells[1, 1] = "=" + mdlData.ExcelCellTranslator(6 + 2 * (curLect - 1), k - 1) + "+" +
                                            mdlData.ExcelCellTranslator(7 + 2 * (curLect - 1), k - 1);
                                        break;
                                    case 28:
                                        cells.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                                        break;
                                    case 29:
                                        //Отменяем границы
                                        cells.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                                        cells.Cells[1, 1] = mdlData.colLecturer[i].Rate.ToString("0.00");
                                        break;
                                    case 30:
                                        //Отменяем границы
                                        cells.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                                        cells.Cells[1, 1] = mdlData.colLecturer[i].Degree.Short + ", " +
                                            mdlData.colLecturer[i].Duty.Short;
                                        break;
                                }
                            }

                            strSumI += mdlData.ExcelCellTranslator(6 + 2 * (curLect - 1), 26) + "+";
                            strSumII += mdlData.ExcelCellTranslator(7 + 2 * (curLect - 1), 26) + "+";
                            strSum += mdlData.ExcelCellTranslator(6 + 2 * (curLect - 1), 27) + "+";

                            curLect += 1;
                        }
                    }
                }

                strSumI = strSumI.Substring(0, strSumI.Length - 1);
                strSumII = strSumII.Substring(0, strSumII.Length - 1);
                strSum = strSum.Substring(0, strSum.Length - 1);

                //Объединённая итоговая сумма часов
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (2 * curLect), 27),
                    mdlData.ExcelCellTranslator(5 + (2 * curLect), 27));
                cells.Merge();
                //Горизонтальное выравнивание по центру в ячейках
                cells.HorizontalAlignment = Excel.Constants.xlCenter;
                //Вертикальное выравнивание по центру в ячейках
                cells.VerticalAlignment = Excel.Constants.xlCenter;
                //Задаём границы
                cells.Borders.Weight = 2;
                cells.Cells[1, 1] = strSum;

                //Итоговая сумма часов по семестрам (3 строки)
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (2 * curLect), 26),
                    mdlData.ExcelCellTranslator(5 + (2 * curLect) + 1, 26));
                cells.Cells[1, 1] = strSumI;
                cells.Cells[2, 1] = strSumII;
                //Задаём границы
                cells.Borders.Weight = 2;
                cells.Cells[3, 1] = "=" + mdlData.ExcelCellTranslator(4 + (2 * curLect), 26) + "+" +
                    mdlData.ExcelCellTranslator(5 + (2 * curLect), 26);

                //Надписи семестров по итогам
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (2 * curLect), 25),
                    mdlData.ExcelCellTranslator(5 + (2 * curLect) + 1, 25));
                cells.Cells[1, 1] = "I сем.";
                cells.Cells[2, 1] = "II сем.";

                //Надпись итого
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (2 * curLect), 24),
                    mdlData.ExcelCellTranslator(5 + (2 * curLect), 24));
                cells.Cells[1, 1] = "Итого:";

                //Суммируем ставки преподавателей
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (2 * curLect) + 1, 2),
                    mdlData.ExcelCellTranslator(5 + (2 * curLect) + 7, 3));
                cells.Cells[1, 1] = "Сумма ставок:";
                cells.Cells[1, 2] = sumRate.ToString();
                cells.Cells[2, 1] = "Сумма асс.:";
                cells.Cells[2, 2] = assist.ToString();
                cells.Cells[3, 1] = "Сумма ст.преп.:";
                cells.Cells[3, 2] = hitutor.ToString();
                cells.Cells[4, 1] = "Сумма доц.:";
                cells.Cells[4, 2] = lecturer.ToString();
                cells.Cells[5, 1] = "Сумма проф.:";
                cells.Cells[5, 2] = proff.ToString();
                cells.Cells[6, 2] = sumRate.ToString();

                //Подпись заведующего кафедрой
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (2 * curLect) + 3, 12),
                    mdlData.ExcelCellTranslator(5 + (2 * curLect) + 3, 12));
                cells.Cells[1, 1] = "Заведующий кафедрой                                                                                       / Л.А. Баранов /";

                //-----------Формируем таблицу даными

                ObjExcel.UserControl = true;

                ObjWorkBook.SaveCopyAs(Application.StartupPath + @"\Ведомость плановая - учебное управление " + 
                    DateTime.Now.Date.ToString("yyyyMMdd") + " " + 
                    DateTime.Now.TimeOfDay.ToString("hhmmss") + ".xlsx");
                ObjWorkBook.Close(false, "", Missing.Value);

                ObjExcel.Quit();
            }
            catch
            {
                MessageBox.Show("Возможно на этом компьютере присутствует проблема совместимости с MS Exсel." +
                                " Попробуйте установить версию 2007 и выше.");
            }
        }

        private void UpravlenieExcel()
        {
            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;
            
            try
            {
                //Создаём новое Excel приложение
                Excel.Application ObjExcel = new Excel.Application();

                UpravlenieExcelCore();

                ObjExcel.Quit();
            }
            catch
            {
                MessageBox.Show("Возможно на этом компьютере присутствует проблема совместимости с MS Exсel." +
                                " Попробуйте установить версию 2007 и выше.");
            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            gridIntoExcel();
        }

        private void DispatchWordTabRowsCount(ref double assist, ref double hitutor, 
                            ref double proff, ref double lecturer, ref double sumRate,
                            IList<clsDistribution> coll)
        {
            bool accessFlg = false;
            int tmpSum;

            //Задаём количество столбцов
            //оно остаётся неизменным
            //0. № п/п
            //1. Фамилия, Имя, Отчество
            //2. Учёная степень, звание
            //3. Лекции по семестрам
            //4. Всего лекций
            //5. Экзамены
            //6. Зачёты
            //7. Консультации
            //8. Лабораторные работы
            //9. Практические занятия
            //10. Домашние задания и рефераты
            //11. Текущая аттестация
            //12. Индивидуальные занятия
            //13. Контрольные работы
            //14. Курсовой проект, курсовая работа
            //15. Дипломный проект
            //16. Учебная практика
            //17. Преддипломная и производственная практика
            //18. ГЭК
            //19. Приёмная комиссия
            //20. ФПК
            //21. Аспирантура
            //22. Посещение занятий
            //23. Руководство магистерской программой
            //24. I сем./II сем.
            //25. Всего за год
            //26. (пусто)
            //27. (ставка)
            //28. (степень, звание)

            for (int i = 0; i <= 28; i++)
            {
                dgScheduleManagement.Columns.Add("", "");
            }

            //Сразу добавляем строки: под заголовок
            dgScheduleManagement.Rows.Add();
            //под пробел
            dgScheduleManagement.Rows.Add();
            //под шапку
            dgScheduleManagement.Rows.Add();

            //Просматриваем каждого преподавателя
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                //Сначала определяем, нужно ли, в принципе, 
                //рассматривать преподавателя
                if (optMain.Checked || optCombine.Checked)
                {
                    accessFlg = (mdlData.colLecturer[i].Rate > 0);
                }
                else
                {
                    if (optHoured.Checked || optMainDop.Checked || optCombineDop.Checked)
                    {
                        accessFlg = true;
                    }
                }

                //если прошёл отбор
                if (accessFlg)
                {
                    if (mdlData.colLecturer[i].Duty.Short.Equals("асс."))
                    {
                        assist += mdlData.colLecturer[i].Rate;
                    }

                    if (mdlData.colLecturer[i].Duty.Short.Equals("доц."))
                    {
                        lecturer += mdlData.colLecturer[i].Rate;
                    }

                    if (mdlData.colLecturer[i].Duty.Short.Equals("ст.преп."))
                    {
                        hitutor += mdlData.colLecturer[i].Rate;
                    }

                    if (mdlData.colLecturer[i].Duty.Short.Equals("проф."))
                    {
                        proff += mdlData.colLecturer[i].Rate;
                    }

                    if (mdlData.colLecturer[i].Duty.Short.Equals("зав.каф."))
                    {
                        proff += mdlData.colLecturer[i].Rate;
                    }

                    if (optHoured.Checked || optMainDop.Checked || optCombineDop.Checked)
                    {
                        tmpSum = 0;
                        //Просматриваем нагрузку
                        for (int j = 0; j <= coll.Count - 1; j++)
                        {
                            tmpSum += mdlData.toSumDistributionComponents(coll[j]);
                        }

                        if (tmpSum == 0)
                        {
                            accessFlg = false;
                        }
                    }

                    if (accessFlg)
                    {
                        //добавляем строку под первый семестр
                        dgScheduleManagement.Rows.Add();
                        //добавляем строку под второй семестр
                        dgScheduleManagement.Rows.Add();
                    }
                }
            }

            sumRate = proff + hitutor + lecturer + assist;

            for (int i = 0; i <= 5; i++)
            {
                dgScheduleManagement.Rows.Add();
            }
        }

        //Основное форматирование Excel планового распределения нагрузки
        private void UpravlenieExcelCore(object ObjMissing, Word._Application ObjWord)
        {
            int curRow;
            int curLect;

            int sumLecI;
            int sumLecII;

            int sumExamI;
            int sumExamII;

            int sumCredI;
            int sumCredII;

            int sumTutI;
            int sumTutII;

            int sumLabI;
            int sumLabII;

            int sumPracI;
            int sumPracII;

            int sumRefI;
            int sumRefII;

            int sumIndI;
            int sumIndII;

            int sumKRAPKI;
            int sumKRAPKII;

            int sumKursPrI;
            int sumKursPrII;

            int sumDiplI;
            int sumDiplII;

            int sumTutPrI;
            int sumTutPrII;

            int sumPreDipI;
            int sumPreDipII;

            int sumGAKI;
            int sumGAKII;

            int sumPostGrI;
            int sumPostGrII;

            int sumVisI;
            int sumVisII;

            int sumMagI;
            int sumMagII;

            int sumI;
            int sumII;

            int sumAllI;
            int sumAllII;

            int CheckSum;

            double assist = 0;
            double hitutor = 0;
            double proff = 0;
            double lecturer = 0;
            double sumRate = 0;

            IList<clsDistribution> coll = null;
            bool accessFlg = false;
            
            //Добавляем новый чистый документ Word
            Word._Document ObjDoc = ObjWord.Application.Documents.Add();
            ObjDoc.Activate();

            if (optMain.Checked || optMainDop.Checked)
            {
                coll = mdlData.colDistribution;
            }
            else
            {
                if (optHoured.Checked)
                {
                    coll = mdlData.colHouredDistribution;
                }
                else
                {
                    if (optCombine.Checked || optCombineDop.Checked)
                    {
                        coll = mdlData.colCombineDistribution;
                    }
                }
            }
            //
            DispatchWordTabRowsCount(ref assist, ref hitutor, ref proff, ref lecturer, ref sumRate, coll);

            //-----------------------------------------------------------------
            //---------------Фрагмент заполнения рабочей области---------------
            //-----------------------------------------------------------------

            //Берём первую строку (с нуля)
            curRow = 0;

            dgScheduleManagement[9, curRow].Value = "Кафедра ''Управление и защита информации'' 20   / 20   учебный год";

            //Берём третью строку (с нуля)
            curRow = 2;

            //Формируем шапку выписки
            dgScheduleManagement[0, curRow].Value = "№ п/п";
            dgScheduleManagement[1, curRow].Value = "Фамилия, Имя, Отчество";
            dgScheduleManagement[2, curRow].Value = "Учёная степень, звание";
            dgScheduleManagement[3, curRow].Value = "Лекции по семестрам";
            dgScheduleManagement[4, curRow].Value = "Всего лекций";
            dgScheduleManagement[5, curRow].Value = "Экзамены";
            dgScheduleManagement[6, curRow].Value = "Зачёты";
            dgScheduleManagement[7, curRow].Value = "Консультации";
            dgScheduleManagement[8, curRow].Value = "Лабораторные работы";
            dgScheduleManagement[9, curRow].Value = "Практические занятия";
            dgScheduleManagement[10, curRow].Value = "Домашние задания и рефераты";
            dgScheduleManagement[11, curRow].Value = "Текущая аттестация";
            dgScheduleManagement[12, curRow].Value = "Индивидуальные занятия";
            dgScheduleManagement[13, curRow].Value = "Контрольные работы";
            dgScheduleManagement[14, curRow].Value = "Курсовой проект, курсовая работа";
            dgScheduleManagement[15, curRow].Value = "Дипломный проект";
            dgScheduleManagement[16, curRow].Value = "Учебная практика";
            dgScheduleManagement[17, curRow].Value = "Преддипломная и производственная практика";
            dgScheduleManagement[18, curRow].Value = "ГЭК";
            dgScheduleManagement[19, curRow].Value = "Приёмная комиссия";
            dgScheduleManagement[20, curRow].Value = "ФПК";
            dgScheduleManagement[21, curRow].Value = "Аспирантура";
            dgScheduleManagement[22, curRow].Value = "Посещение занятий";
            dgScheduleManagement[23, curRow].Value = "Руководство магистерской программой";
            dgScheduleManagement[24, curRow].Value = "I сем./II сем.";
            dgScheduleManagement[25, curRow].Value = "Всего за год";
            dgScheduleManagement[26, curRow].Value = "";
            dgScheduleManagement[27, curRow].Value = "";
            dgScheduleManagement[28, curRow].Value = "";

            //Идём на следующую строку
            curRow += 1;

            curLect = 1;

            sumAllI = 0;
            sumAllII = 0;

            //Просматриваем каждого преподавателя
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                //Сначала определяем, нужно ли, в принципе, 
                //рассматривать преподавателя
                if (optMain.Checked || optCombine.Checked)
                {
                    accessFlg = (mdlData.colLecturer[i].Rate > 0);
                }
                else
                {
                    if (optHoured.Checked || optMainDop.Checked || optCombineDop.Checked)
                    {
                        accessFlg = true;
                    }
                }

                //Если прошёл отбор
                if (accessFlg)
                {
                    sumLecI = 0;
                    sumLecII = 0;

                    sumExamI = 0;
                    sumExamII = 0;

                    sumCredI = 0;
                    sumCredII = 0;

                    sumTutI = 0;
                    sumTutII = 0;

                    sumLabI = 0;
                    sumLabII = 0;

                    sumPracI = 0;
                    sumPracII = 0;

                    sumRefI = 0;
                    sumRefII = 0;

                    sumIndI = 0;
                    sumIndII = 0;

                    sumKRAPKI = 0;
                    sumKRAPKII = 0;

                    sumKursPrI = 0;
                    sumKursPrII = 0;

                    sumDiplI = 0;
                    sumDiplII = 0;

                    sumTutPrI = 0;
                    sumTutPrII = 0;

                    sumPreDipI = 0;
                    sumPreDipII = 0;

                    sumGAKI = 0;
                    sumGAKII = 0;

                    sumPostGrI = 0;
                    sumPostGrII = 0;

                    sumVisI = 0;
                    sumVisII = 0;

                    sumMagI = 0;
                    sumMagII = 0;

                    sumI = 0;
                    sumII = 0;

                    //Просматриваем нагрузку
                    for (int j = 0; j <= coll.Count - 1; j++)
                    {
                        if (!(coll[j].Lecturer == null))
                        {
                            if (coll[j].Lecturer.Equals(mdlData.colLecturer[i]))
                            {
                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumLecI += coll[j].Lecture;
                                    sumI += coll[j].Lecture;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumLecII += coll[j].Lecture;
                                    sumII += coll[j].Lecture;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumExamI += coll[j].Exam;
                                    sumI += coll[j].Exam;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumExamII += coll[j].Exam;
                                    sumII += coll[j].Exam;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumCredI += coll[j].Credit;
                                    sumI += coll[j].Credit;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumCredII += coll[j].Credit;
                                    sumII += coll[j].Credit;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumTutI += coll[j].Tutorial;
                                    sumI += coll[j].Tutorial;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumTutII += coll[j].Tutorial;
                                    sumII += coll[j].Tutorial;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumLabI += coll[j].LabWork;
                                    sumI += coll[j].LabWork;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumLabII += coll[j].LabWork;
                                    sumII += coll[j].LabWork;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumPracI += coll[j].Practice;
                                    sumI += coll[j].Practice;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumPracII += coll[j].Practice;
                                    sumII += coll[j].Practice;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumRefI += coll[j].RefHomeWork;
                                    sumI += coll[j].RefHomeWork;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumRefII += coll[j].RefHomeWork;
                                    sumII += coll[j].RefHomeWork;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumIndI += coll[j].IndividualWork;
                                    sumI += coll[j].IndividualWork;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumIndII += coll[j].IndividualWork;
                                    sumII += coll[j].IndividualWork;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumKRAPKI += coll[j].KRAPK;
                                    sumI += coll[j].KRAPK;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumKRAPKII += coll[j].KRAPK;
                                    sumII += coll[j].KRAPK;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumKursPrI += coll[j].KursProject;
                                    sumI += coll[j].KursProject;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumKursPrII += coll[j].KursProject;
                                    sumII += coll[j].KursProject;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumDiplI += coll[j].DiplomaPaper;
                                    sumI += coll[j].DiplomaPaper;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumDiplII += coll[j].DiplomaPaper;
                                    sumII += coll[j].DiplomaPaper;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumTutPrI += coll[j].TutorialPractice;
                                    sumI += coll[j].TutorialPractice;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumTutPrII += coll[j].TutorialPractice;
                                    sumII += coll[j].TutorialPractice;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumPreDipI += coll[j].PreDiplomaPractice +
                                        coll[j].ProducingPractice;
                                    sumI += coll[j].PreDiplomaPractice +
                                        coll[j].ProducingPractice;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumPreDipII += coll[j].PreDiplomaPractice +
                                        coll[j].ProducingPractice;
                                    sumII += coll[j].PreDiplomaPractice +
                                        coll[j].ProducingPractice;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumGAKI += coll[j].GAK;
                                    sumI += coll[j].GAK;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumGAKII += coll[j].GAK;
                                    sumII += coll[j].GAK;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumPostGrI += coll[j].PostGrad;
                                    sumI += coll[j].PostGrad;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumPostGrII += coll[j].PostGrad;
                                    sumII += coll[j].PostGrad;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumVisI += coll[j].Visiting;
                                    sumI += coll[j].Visiting;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumVisII += coll[j].Visiting;
                                    sumII += coll[j].Visiting;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumMagI += coll[j].Magistry;
                                    sumI += coll[j].Magistry;
                                }

                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumMagII += coll[j].Magistry;
                                    sumII += coll[j].Magistry;
                                }
                            }
                        }
                    }

                    sumAllI += sumI;
                    sumAllII += sumII;

                    if (optHoured.Checked || optMainDop.Checked || optCombineDop.Checked)
                    {
                        accessFlg = !(sumI + sumII == 0);
                    }

                    if (accessFlg)
                    {
                        dgScheduleManagement[0, curRow].Value = curLect + ".";
                        dgScheduleManagement[1, curRow].Value = mdlData.colLecturer[i].FIO + " (" +
                            mdlData.colLecturer[i].Rate.ToString("0.00") + " ставки)";
                        dgScheduleManagement[2, curRow].Value = mdlData.colLecturer[i].Degree.Short + ", " +
                            mdlData.colLecturer[i].Duty.Short;

                        dgScheduleManagement[27, curRow].Value = mdlData.colLecturer[i].Rate.ToString("0.00");
                        dgScheduleManagement[28, curRow].Value = mdlData.colLecturer[i].Degree.Short + ", " +
                            mdlData.colLecturer[i].Duty.Short;

                        dgScheduleManagement[3, curRow].Value = sumLecI;
                        dgScheduleManagement[5, curRow].Value = sumExamI;
                        dgScheduleManagement[6, curRow].Value = sumCredI;
                        dgScheduleManagement[7, curRow].Value = sumTutI;
                        dgScheduleManagement[8, curRow].Value = sumLabI;
                        dgScheduleManagement[9, curRow].Value = sumPracI;
                        dgScheduleManagement[10, curRow].Value = sumRefI;
                        dgScheduleManagement[11, curRow].Value = 0;
                        dgScheduleManagement[12, curRow].Value = sumIndI;
                        dgScheduleManagement[13, curRow].Value = sumKRAPKI;
                        dgScheduleManagement[14, curRow].Value = sumKursPrI;
                        dgScheduleManagement[15, curRow].Value = sumDiplI;
                        dgScheduleManagement[16, curRow].Value = sumTutPrI;
                        dgScheduleManagement[17, curRow].Value = sumPreDipI;
                        dgScheduleManagement[18, curRow].Value = sumGAKI;
                        dgScheduleManagement[19, curRow].Value = 0;
                        dgScheduleManagement[20, curRow].Value = 0;
                        dgScheduleManagement[21, curRow].Value = sumPostGrI;
                        dgScheduleManagement[22, curRow].Value = sumVisI;
                        dgScheduleManagement[23, curRow].Value = sumMagI;
                        dgScheduleManagement[24, curRow].Value = sumI;

                        curRow += 1;

                        dgScheduleManagement[3, curRow].Value = sumLecII;
                        dgScheduleManagement[4, curRow].Value = sumLecI + sumLecII;
                        dgScheduleManagement[5, curRow].Value = sumExamII;
                        dgScheduleManagement[6, curRow].Value = sumCredII;
                        dgScheduleManagement[7, curRow].Value = sumTutII;
                        dgScheduleManagement[8, curRow].Value = sumLabII;
                        dgScheduleManagement[9, curRow].Value = sumPracII;
                        dgScheduleManagement[10, curRow].Value = sumRefII;
                        dgScheduleManagement[11, curRow].Value = 0;
                        dgScheduleManagement[12, curRow].Value = sumIndII;
                        dgScheduleManagement[13, curRow].Value = sumKRAPKII;
                        dgScheduleManagement[14, curRow].Value = sumKursPrII;
                        dgScheduleManagement[15, curRow].Value = sumDiplII;
                        dgScheduleManagement[16, curRow].Value = sumTutPrII;
                        dgScheduleManagement[17, curRow].Value = sumPreDipII;
                        dgScheduleManagement[18, curRow].Value = sumGAKII;
                        dgScheduleManagement[19, curRow].Value = 0;
                        dgScheduleManagement[20, curRow].Value = 0;
                        dgScheduleManagement[21, curRow].Value = sumPostGrII;
                        dgScheduleManagement[22, curRow].Value = sumVisII;
                        dgScheduleManagement[23, curRow].Value = sumMagII;
                        dgScheduleManagement[24, curRow].Value = sumII;
                        dgScheduleManagement[25, curRow].Value = sumI + sumII;

                        curRow += 1;

                        //Увеличиваем счётчик преподавателей
                        curLect += 1;
                    }
                }
            }

            CheckSum = 0;

            for (int j = 3; j <= dgScheduleManagement.RowCount - 1; j++)
            {
                CheckSum += Convert.ToInt32(dgScheduleManagement[25, j].Value);
            }

            dgScheduleManagement[22, curRow].Value = "Итого:";
            dgScheduleManagement[23, curRow].Value = "I сем.";
            dgScheduleManagement[24, curRow].Value = sumAllI;

            curRow += 1;

            dgScheduleManagement[23, curRow].Value = "II сем.";
            dgScheduleManagement[24, curRow].Value = sumAllII;
            dgScheduleManagement[25, curRow].Value = CheckSum;

            dgScheduleManagement[1, curRow].Value = "Сумма ставок:";
            dgScheduleManagement[2, curRow].Value = sumRate.ToString("0.00");

            curRow += 1;

            dgScheduleManagement[24, curRow].Value = sumAllI + sumAllII;

            dgScheduleManagement[1, curRow].Value = "Сумма асс.:";
            dgScheduleManagement[2, curRow].Value = assist.ToString("0.00");

            curRow += 1;

            dgScheduleManagement[1, curRow].Value = "Сумма ст.преп.:";
            dgScheduleManagement[2, curRow].Value = hitutor.ToString("0.00");

            dgScheduleManagement[10, curRow].Value = "Заведующий кафедрой";
            dgScheduleManagement[18, curRow].Value = "/ Л.А. Баранов /";

            curRow += 1;

            dgScheduleManagement[1, curRow].Value = "Сумма доц.:";
            dgScheduleManagement[2, curRow].Value = lecturer.ToString("0.00");

            curRow += 1;

            dgScheduleManagement[1, curRow].Value = "Сумма проф.:";
            dgScheduleManagement[2, curRow].Value = proff.ToString("0.00");

            dgScheduleManagement[10, curRow].Value = "Директор ИТТСУ";
            dgScheduleManagement[18, curRow].Value = "/ П.Ф. Бестемьянов /";

            curRow += 1;

            dgScheduleManagement[1, curRow].Value = "";
            dgScheduleManagement[2, curRow].Value = sumRate.ToString("0.00");

            //-----------------------------------------------------------------
            //---------------Фрагмент заполнения рабочей области---------------
            //-----------------------------------------------------------------

            ObjDoc.SaveAs(Application.StartupPath + @"\"
                + " " + mdlData.colSemestr[cmbSemestrList.SelectedIndex].About
                + " " + mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear + " уч. года " +
                DateTime.Now.Date.ToString("yyyyMMdd") + " " +
                DateTime.Now.TimeOfDay.ToString("hhmmss") + ".docx");

            ObjDoc.Close();
        }

        //
        private int DispatchWordCountRows()
        {
            bool HaveLoad;
            bool Both;
            int colRows = 1;

            Both = false;
            if (mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum == "-")
            {
                Both = true;
            }

            //Просматриваем каждого преподавателя
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                //Считаем, по умолчанию, что преподаватель не нагружен
                HaveLoad = false;
                //Просматриваем нагрузку
                for (int j = 0; j <= mdlData.colDistribution.Count - 1; j++)
                {
                    //Если выбран конкретный учебный семестр
                    if (!Both)
                    {
                        if (!(mdlData.colDistribution[j].Semestr == null))
                        {
                            //Если дисциплина относится к указанному в комбинированном
                            //списке семестру
                            if (mdlData.colDistribution[j].Semestr.Equals(mdlData.colSemestr[cmbSemestrList.SelectedIndex]))
                            {
                                //Если рассматриваемый преподаватель совпадает с
                                //преподавателем, указанным в нагрузке
                                if (mdlData.colLecturer[i].Equals(mdlData.colDistribution[j].Lecturer))
                                {
                                    //Если дисциплина предполагает лекционные часы,
                                    //практические занятия, лабораторные работы,
                                    //курсовой проект, учебную практику, производственную
                                    //практику или дипломную практику, то
                                    //выделяем под неё строчку текста
                                    if (mdlData.NonZeroForDispatchOR(mdlData.colDistribution[j]))
                                    {
                                        if (mdlData.colDistribution[j].flgDispatch)
                                        {
                                            //Если преподаватель что-либо из этого ведёт
                                            //добавляем строку
                                            colRows++;
                                            //Получается, что преподаватель нагружен
                                            HaveLoad = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //Если семестр не выбран, то необходимо предоставить данные
                    //по обоим семестрам
                    else
                    {
                        //Если рассматриваемый преподаватель совпадает с
                        //преподавателем, указанным в нагрузке
                        if (mdlData.colLecturer[i].Equals(mdlData.colDistribution[j].Lecturer))
                        {
                            //Если дисциплина предполагает лекционные часы,
                            //практические занятия, лабораторные работы,
                            //курсовой проект, учебную практику, производственную
                            //практику или дипломную практику, то
                            //выделяем под неё строчку текста
                            if (mdlData.NonZeroForDispatchOR(mdlData.colDistribution[j]))
                            {
                                if (mdlData.colDistribution[j].flgDispatch)
                                {
                                    //Если преподаватель что-либо из этого ведёт
                                    //добавляем строку
                                    colRows++;
                                    //Получается, что преподаватель нагружен
                                    HaveLoad = true;
                                }
                            }
                        }
                    }
                }

                //Только если у преподавателя есть нагрузка
                if (HaveLoad)
                {
                    //Разделяем преподавателей пустой строчкой
                    colRows++;
                }
            }

            return colRows;
        }

        //Процедура формирования документа Word с заявкой в диспетчерскую
        //исходный вариант
        private void DispatchWordCore(object ObjMissing, Word._Application ObjWord)
        {
            int curRow;
            int colRows;
            int stRow = 1;
            int endRow = 1;
            bool HaveLoad;
            bool MergeCells;
            bool Both;

            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;
            Word.Table ObjTable;
            Word.Range ObjWordRange;

            //Добавляем новый чистый документ Word
            Word._Document ObjDoc = ObjWord.Application.Documents.Add();
            ObjDoc.Activate();

            //Настраиваем видимость - в процессе создания документ видно
            ObjWord.Visible = true;
            //Настройка границ
            mdlData.WordPageDefault(ref ObjWord, ref ObjDoc, 0.75f, 0.75f, 0.75f, 0.75f);
            //Настройка альбомной ориентации страницы
            ObjDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            //Настройка 80% масштаба
            ObjDoc.ActiveWindow.View.Zoom.Percentage = 80;

            //Добавляем абзац текста в документ
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            //Вписываем в него название университета
            ObjParagraph.Range.Text = mdlData.UniversityPrefName + " " +
                                      mdlData.UniversityName + " " + 
                                      mdlData.UniversitySuffName;
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //
            ObjParagraph.Range.Font.Name = "Times New Roman";
            //
            ObjParagraph.Range.Font.Size = 12;
            //
            ObjParagraph.Range.ParagraphFormat.LeftIndent = 0;
            //
            ObjParagraph.Range.ParagraphFormat.RightIndent = 0;
            //
            ObjParagraph.Range.ParagraphFormat.Space1();

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            //Пишем наименование таблицы (Заявка)
            ObjParagraph.Range.Text = "ЗАЯВКА";
            ObjParagraph.Range.InsertParagraphAfter();
            //Пишем наименование кафедры, номер семестра и учебный год
            //к которым относится заявка
            ObjParagraph.Range.Text = "Кафедры " + "\"" + mdlData.DepartmentName + "\" в диспетчерскую " +
                                                    "на проведение занятий " + mdlData.colSemestr[cmbSemestrList.SelectedIndex].About +
                                                    " " + mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear + " учебного года";

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            //Прописываем в строке название "Семестр"
            ObjParagraph.Range.Text = "Семестр: " + mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum;
            //Выравнивание по левому краю
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            //Пишем заготовку для заседания кафедры
            //(потом перенести в переменную, тягаемую из таблицы Параметры)
            ObjParagraph.Range.Text = "ПРИМЕЧАНИЕ ПО КАФЕДРЕ: просьба не назначать учебные часы для ВСЕХ преподавателей кафедры " +
                "по понедельникам обеих недель с 15:00 до 16:30 (заседание кафедры)";

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            //Пишем комментарий для диспетчеров о том, как воспринимать
            //пожелания преподавателей и аудитории
            //(позже занести в переменную, хранимую в таблице Параметры)
            ObjParagraph.Range.Text = "*КОММЕНТАРИЙ К СТРУКТУРЕ ПРИМЕЧАНИЙ: в графе собраны пожелания по распределению и аудиториям: \n" +
                "А) примечания, указанные напротив дисциплин - относятся к дисциплинам; \nБ) комментарии напротив пустых " +
                "строк - относятся к преподавателям, упомянутым НАД комментарием.";

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            //Считаем количество строк
            colRows = DispatchWordCountRows();

            //Вставляем таблицу colRows (строки) x 11 (столбцы), заполняем её данными о преподавателе
            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjTable = ObjDoc.Tables.Add(ObjWordRange, colRows, 11, ref ObjMissing, ref ObjMissing);
            //Добавляем обрамление таблицы
            ObjTable.Borders.Enable = 1;

            //Текущей строкой считаем первую
            curRow = 1;
            //Заполняем первую строку и одновременно формируем размерности
            ObjTable.Cell(curRow, 1).Range.Text = "Преподаватель";
            ObjTable.Cell(curRow, 2).Range.Text = "Курс";
            ObjTable.Cell(curRow, 3).Range.Text = "Специальность";
            ObjTable.Cell(curRow, 4).Range.Text = "Название дисциплины";
            ObjTable.Cell(curRow, 5).Range.Text = "Лекция";
            ObjTable.Cell(curRow, 6).Range.Text = "Практ.зан.";
            ObjTable.Cell(curRow, 7).Range.Text = "Лаб.раб";
            ObjTable.Cell(curRow, 8).Range.Text = "Курс.пр.";
            ObjTable.Cell(curRow, 9).Range.Text = "Практика";
            ObjTable.Cell(curRow, 10).Range.Text = "";
            ObjTable.Cell(curRow, 11).Range.Text = "Примечания*";
            
            ObjTable.Rows[curRow].Range.Font.Bold = 1;

            curRow++;

            Both = false;
            if (mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum == "-")
            {
                Both = true;
            }

            //Просматриваем каждого преподавателя
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                //По умолчанию считаем, что у преподавателя нет нагрузки
                HaveLoad = false;
                //По умолчанию считаем, что ячейки в столбце с преподавателем
                //надо объединить
                MergeCells = true;
                //Просматриваем нагрузку
                for (int j = 0; j <= mdlData.colDistribution.Count - 1; j++)
                {
                    //Если в комбинированном списке выбран
                    //конкретный семестр
                    if (!Both)
                    {
                        if (!(mdlData.colDistribution[j].Semestr == null))
                        {
                            //Если дисциплина относится к выбранному семестру
                            if (mdlData.colDistribution[j].Semestr.Equals(mdlData.colSemestr[cmbSemestrList.SelectedIndex]))
                            {
                                //Если рассматриваемый преподаватель совпадает с
                                //преподавателем, указанным в нагрузке
                                if (mdlData.colLecturer[i].Equals(mdlData.colDistribution[j].Lecturer))
                                {
                                    //Если дисциплина предполагает лекционные часы,
                                    //практические занятия, лабораторные работы,
                                    //курсовой проект, учебную практику, производственную
                                    //практику или дипломную практику, то
                                    //выделяем под неё строчку текста
                                    if (mdlData.NonZeroForDispatchOR(mdlData.colDistribution[j]))
                                    {
                                        //Если элемент распределения нагрузки необходимо внести
                                        //в заявку для диспетчерской
                                        if (mdlData.colDistribution[j].flgDispatch)
                                        {
                                            //Если выясняется, что у преподавателя есть дублёр,
                                            //то необходимо сбросить признак наличия нагрузки
                                            if (mdlData.colDistribution[j].Doubler != null)
                                            {
                                                HaveLoad = false;
                                            }

                                            //Если у преподавателя нет нагрузки, но мы попали сюда,
                                            //то надо написать его имя в отчётную форму
                                            if (!HaveLoad)
                                            {
                                                //Если у преподавателя нет дублёра,
                                                //то в заявку идёт он сам
                                                if (mdlData.colDistribution[j].Doubler == null)
                                                {
                                                    //Печатаем фамилию, имя, отчество преподавателя
                                                    ObjTable.Cell(curRow, 1).Range.Text = mdlData.colLecturer[i].FIO + " (" +
                                                        mdlData.colLecturer[i].Duty.Short + ")";
                                                    //Если мы сюда попали, то получается, что у преподавателя есть нагрузка
                                                    HaveLoad = true;
                                                }
                                                //Если у преподавателя есть дублёр, то в заявку идёт
                                                //дублёр
                                                else
                                                {
                                                    //Печатаем фамилию, имя, отчество преподавателя
                                                    ObjTable.Cell(curRow, 1).Range.Text = mdlData.colDistribution[j].Doubler.FIO + " (" +
                                                        mdlData.colDistribution[j].Doubler.Duty.Short + ")";
                                                    //Хоть и известно, что у преподавателя есть нагрузка, но признак
                                                    //её наличия не выставляем
                                                }

                                                //Выделяем имя преподавателя жирным шрифтом
                                                ObjTable.Cell(curRow, 1).Range.Font.Bold = 1;
                                                //Фиксируем первую строку, выделяемую на преподавателя
                                                stRow = curRow;
                                            }

                                            //Печатаем курс
                                            ObjTable.Cell(curRow, 2).Range.Text = mdlData.colDistribution[j].KursNum.Kurs.ToString();
                                            //Печатаем специальность
                                            ObjTable.Cell(curRow, 3).Range.Text = mdlData.colDistribution[j].Speciality.ShortUpravlenie
                                                + " (" + mdlData.colDistribution[j].Speciality.ShortInstitute + ")";
                                            //Печатаем название дисциплины
                                            ObjTable.Cell(curRow, 4).Range.Text = mdlData.colDistribution[j].Subject.Subject;
                                            //Лекционные часы
                                            ObjTable.Cell(curRow, 5).Range.Text = mdlData.colDistribution[j].Lecture.ToString();
                                            //Практические занятия в часах
                                            ObjTable.Cell(curRow, 6).Range.Text = mdlData.colDistribution[j].Practice.ToString();
                                            //Лабораторные работы в часах
                                            ObjTable.Cell(curRow, 7).Range.Text = mdlData.colDistribution[j].LabWork.ToString();
                                            //Курсовой проект в часах
                                            ObjTable.Cell(curRow, 8).Range.Text = mdlData.colDistribution[j].KursProject.ToString();
                                            //Практика в часах
                                            //Либо записываем преддипломную практику
                                            if (mdlData.colDistribution[j].PreDiplomaPractice > 0)
                                            {
                                                ObjTable.Cell(curRow, 9).Range.Text = mdlData.colDistribution[j].PreDiplomaPractice.ToString();
                                            }
                                            //Либо записываем учебную практику
                                            if (mdlData.colDistribution[j].TutorialPractice > 0)
                                            {
                                                ObjTable.Cell(curRow, 9).Range.Text = mdlData.colDistribution[j].TutorialPractice.ToString();
                                            }
                                            //Либо записываем производственную практику
                                            if (mdlData.colDistribution[j].ProducingPractice > 0)
                                            {
                                                ObjTable.Cell(curRow, 9).Range.Text = mdlData.colDistribution[j].ProducingPractice.ToString();
                                            }
                                            //Если поле практики так и осталось пустым, то
                                            //заполняем его нулём
                                            if (ObjTable.Cell(curRow, 9).Range.Text == "\r\a")
                                            {
                                                ObjTable.Cell(curRow, 9).Range.Text = (0).ToString();
                                            }
                                            
                                            //Записываем содержимое пожеланий по аудиториям
                                            ObjTable.Cell(curRow, 11).Range.Text = mdlData.colDistribution[j].Text;

                                            //Если имеется связка по лабораторным работам, то указываем в примечании
                                            //с кем в паре проводятся лабораторные работы (электроника)
                                            if (mdlData.colDistribution[j].LabWorkConnect != null)
                                            {
                                                //Если указан дублёр, то ставим дублёра
                                                if (mdlData.colDistribution[j].LabWorkConnect.Doubler != null)
                                                {
                                                    ObjTable.Cell(curRow, 11).Range.Text += " (совместно с " +
                                                        mdlData.SplitFIOString(mdlData.colDistribution[j].LabWorkConnect.Doubler.FIO, true, false) + ")";
                                                }
                                                //В ином случае указывается основной преподаватель
                                                else
                                                {
                                                    ObjTable.Cell(curRow, 11).Range.Text += " (совместно с " +
                                                        mdlData.SplitFIOString(mdlData.colDistribution[j].LabWorkConnect.Lecturer.FIO, true, false) + ")";
                                                }
                                            }

                                            //Переходим к следующей строке
                                            curRow += 1;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //Если семестр не выбран, то необходимо предоставить информацию
                    //по обоим семестрам
                    else
                    {
                        //Если рассматриваемый преподаватель совпадает с
                        //преподавателем, указанным в нагрузке
                        if (mdlData.colLecturer[i].Equals(mdlData.colDistribution[j].Lecturer))
                        {
                            //Если дисциплина предполагает лекционные часы,
                            //практические занятия, лабораторные работы,
                            //курсовой проект, учебную практику, производственную
                            //практику или дипломную практику, то
                            //выделяем под неё строчку текста
                            if (mdlData.NonZeroForDispatchOR(mdlData.colDistribution[j]))
                            {
                                if (mdlData.colDistribution[j].flgDispatch)
                                {
                                    //Если выясняется, что у преподавателя есть дублёр,
                                    //то необходимо сбросить признак наличия нагрузки
                                    if (mdlData.colDistribution[j].Doubler != null)
                                    {
                                        HaveLoad = false;
                                    }
                                    
                                    //Если у преподавателя нет нагрузки, но мы попали сюда,
                                    //то надо написать его имя в отчётную форму
                                    if (!HaveLoad)
                                    {
                                        //Если у преподавателя нет дублёра,
                                        //то в заявку идёт он сам
                                        if (mdlData.colDistribution[j].Doubler == null)
                                        {
                                            //Печатаем фамилию, имя, отчество преподавателя
                                            ObjTable.Cell(curRow, 1).Range.Text = mdlData.colLecturer[i].FIO + " (" +
                                                        mdlData.colLecturer[i].Duty.Short + ")";
                                            //Получается, что у преподавателя есть нагрузка
                                            HaveLoad = true;
                                        }
                                        //Если у преподавателя есть дублёр, то в заявку идёт
                                        //дублёр
                                        else
                                        {
                                            //Печатаем фамилию, имя, отчество преподавателя
                                            ObjTable.Cell(curRow, 1).Range.Text = mdlData.colDistribution[j].Doubler.FIO + " (" +
                                                mdlData.colDistribution[j].Doubler.Duty.Short + ")";
                                            //Хоть и известно, что у преподавателя есть нагрузка, но признак
                                            //её наличия не выставляем
                                        }
                                        //
                                        ObjTable.Cell(curRow, 1).Range.Font.Bold = 1;
                                        //
                                        stRow = curRow;
                                    }
                                    //Печатаем курс
                                    ObjTable.Cell(curRow, 2).Range.Text = mdlData.colDistribution[j].KursNum.Kurs.ToString();
                                    //Печатаем специальность
                                    ObjTable.Cell(curRow, 3).Range.Text = mdlData.colDistribution[j].Speciality.ShortUpravlenie
                                        + " (" + mdlData.colDistribution[j].Speciality.ShortInstitute + ")";
                                    //Печатаем название дисциплины
                                    ObjTable.Cell(curRow, 4).Range.Text = mdlData.colDistribution[j].Subject.Subject;
                                    //Лекционные часы
                                    ObjTable.Cell(curRow, 5).Range.Text = mdlData.colDistribution[j].Lecture.ToString();
                                    //Практические занятия в часах
                                    ObjTable.Cell(curRow, 6).Range.Text = mdlData.colDistribution[j].Practice.ToString();
                                    //Лабораторные работы в часах
                                    ObjTable.Cell(curRow, 7).Range.Text = mdlData.colDistribution[j].LabWork.ToString();
                                    //Курсовой проект в часах
                                    ObjTable.Cell(curRow, 8).Range.Text = mdlData.colDistribution[j].KursProject.ToString();
                                    //Практика в часах
                                    //Либо записываем преддипломную практику
                                    if (mdlData.colDistribution[j].PreDiplomaPractice > 0)
                                    {
                                        ObjTable.Cell(curRow, 9).Range.Text = mdlData.colDistribution[j].PreDiplomaPractice.ToString();
                                    }
                                    //Либо записываем учебную практику
                                    if (mdlData.colDistribution[j].TutorialPractice > 0)
                                    {
                                        ObjTable.Cell(curRow, 9).Range.Text = mdlData.colDistribution[j].TutorialPractice.ToString();
                                    }
                                    //Либо записываем производственную практику
                                    if (mdlData.colDistribution[j].ProducingPractice > 0)
                                    {
                                        ObjTable.Cell(curRow, 9).Range.Text = mdlData.colDistribution[j].ProducingPractice.ToString();
                                    }
                                    //Если поле практики так и осталось пустым, то
                                    //заполняем его нулём
                                    if (ObjTable.Cell(curRow, 9).Range.Text == "\r\a")
                                    {
                                        ObjTable.Cell(curRow, 9).Range.Text = (0).ToString();
                                    }
                                    //
                                    ObjTable.Cell(curRow, 11).Range.Text = mdlData.colDistribution[j].Text;
                                    //Переходим к следующей строке
                                    curRow += 1;
                                }
                            }
                        }
                    }
                }

                //Только, если у преподавателя есть нагрузка
                if (HaveLoad)
                {
                    //Пишем примечание конкретного преподавателя
                    ObjTable.Cell(curRow, 11).Range.Text = mdlData.colLecturer[i].Text;
                    //Фиксируем последнюю строку, выделяемую на преподавателя
                    endRow = curRow;
                    if (MergeCells)
                    {
                        //Объединяем ячейки в столбце преподавателей
                        ObjTable.Cell(stRow, 1).Merge(ObjTable.Cell(endRow, 1));
                    }
                    //Разделяем преподавателей пустой строчкой
                    curRow += 1;
                }
            }

            ObjTable.Columns[1].Width = ObjWord.CentimetersToPoints(3.44f);
            ObjTable.Columns[2].Width = ObjWord.CentimetersToPoints(1.5f);
            ObjTable.Columns[3].Width = ObjWord.CentimetersToPoints(2.5f);
            ObjTable.Columns[4].Width = ObjWord.CentimetersToPoints(8f);
            ObjTable.Columns[5].Width = ObjWord.CentimetersToPoints(1.25f);
            ObjTable.Columns[6].Width = ObjWord.CentimetersToPoints(1.75f);
            ObjTable.Columns[7].Width = ObjWord.CentimetersToPoints(1.25f);
            ObjTable.Columns[8].Width = ObjWord.CentimetersToPoints(1.5f);
            ObjTable.Columns[9].Width = ObjWord.CentimetersToPoints(1.5f);
            ObjTable.Columns[10].Width = ObjWord.CentimetersToPoints(0.5f);
            ObjTable.Columns[11].Width = ObjWord.CentimetersToPoints(5.5f);

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            ObjParagraph.Range.Text = "Заведующий кафедрой УиЗИ \t" + "/ Л.А. Баранов /";
            //ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjParagraph.Range.ParagraphFormat.TabStops.Add(ObjWord.CentimetersToPoints(20f), Word.WdTabAlignment.wdAlignTabRight);

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            /*
            ObjParagraph.Range.Text = "Директор ИТТСУ \t" + "/ П.Ф. Бестемьянов /";

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();
            */

            ObjParagraph.Range.Text = "Первый зам. директора ИТТСУ - " +
                "начальник учебного отдела \t" + "/ В.А. Гречишников /";

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            ObjParagraph.Range.Text = "И.о. декана вечернего факультета \t" + "/ И.В. Федякин /";

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            //ObjParagraph.Range.Text = "Директор ИУИТ \t" + "/ С.П. Вакуленко /";

            if (cmbSemestrList.SelectedIndex == 1)
            {
                ObjParagraph.Range.Text = "Первый зам. директора ИУИТ - " +
                    "начальник учебного отдела \t" + "/ Е.С. Прокофьева /";
            }

            ObjDoc.SaveAs(Application.StartupPath + @"\"
                + "Заявка в диспетчерскую " + mdlData.colSemestr[cmbSemestrList.SelectedIndex].About
                + " " + mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear.Replace("/","-") + " уч. года " +
                DateTime.Now.Date.ToString("yyyyMMdd") + " " + 
                DateTime.Now.TimeOfDay.ToString("hhmmss") + ".docx");

            ObjDoc.Close();
        }

        //Процедура формирования документа Word с заявкой в диспетчерскую
        //обновлённый вариант
        private void DispatchWordCoreNew(object ObjMissing, Word._Application ObjWord)
        {
            int curRow;
            int colRows;
            int stRow = 1;
            int endRow = 1;
            bool HaveLoad;
            bool MergeCells;
            bool Both;
            bool flgDoubler;

            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;
            Word.Table ObjTable;
            Word.Range ObjWordRange;

            //Добавляем новый чистый документ Word
            Word._Document ObjDoc = ObjWord.Application.Documents.Add();
            ObjDoc.Activate();

            //Настраиваем видимость - в процессе создания документ видно
            ObjWord.Visible = true;
            //Настройка границ
            mdlData.WordPageDefault(ref ObjWord, ref ObjDoc, 0.75f, 0.75f, 0.75f, 0.75f);
            //Настройка альбомной ориентации страницы
            ObjDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            //Настройка 80% масштаба
            ObjDoc.ActiveWindow.View.Zoom.Percentage = 80;

            //Добавляем абзац текста в документ
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            //Вписываем в него название университета
            ObjParagraph.Range.Text = mdlData.UniversityPrefName + " " +
                                      mdlData.UniversityName + " " +
                                      mdlData.UniversitySuffName;
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //
            ObjParagraph.Range.Font.Name = "Times New Roman";
            //
            ObjParagraph.Range.Font.Size = 12;
            //
            ObjParagraph.Range.ParagraphFormat.LeftIndent = 0;
            //
            ObjParagraph.Range.ParagraphFormat.RightIndent = 0;
            //
            ObjParagraph.Range.ParagraphFormat.Space1();

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            //Пишем наименование таблицы (Заявка)
            ObjParagraph.Range.Text = "ЗАЯВКА";
            ObjParagraph.Range.InsertParagraphAfter();
            //Пишем наименование кафедры, номер семестра и учебный год
            //к которым относится заявка
            ObjParagraph.Range.Text = "Кафедры " + "\"" + mdlData.DepartmentName + "\" в диспетчерскую " +
                                                    "на проведение занятий " + mdlData.colSemestr[cmbSemestrList.SelectedIndex].About +
                                                    " " + mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear + " учебного года";

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            //Прописываем в строке название "Семестр"
            ObjParagraph.Range.Text = "Семестр: " + mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum;
            //Выравнивание по левому краю
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            //Пишем заготовку для заседания кафедры
            //(потом перенести в переменную, тягаемую из таблицы Параметры)
            ObjParagraph.Range.Text = "ПРИМЕЧАНИЕ ПО КАФЕДРЕ: просьба не назначать учебные часы для ВСЕХ преподавателей кафедры " +
                "по понедельникам обеих недель с 15:20 до 16:40 (заседание кафедры)";

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            //Пишем комментарий для диспетчеров о том, как воспринимать
            //пожелания преподавателей и аудитории
            //(позже занести в переменную, хранимую в таблице Параметры)
            ObjParagraph.Range.Text = "*КОММЕНТАРИЙ К СТРУКТУРЕ ПРИМЕЧАНИЙ: в графе собраны пожелания по распределению и аудиториям: \n" +
                "А) примечания, указанные напротив дисциплин - относятся к дисциплинам; \nБ) комментарии напротив пустых " +
                "строк - относятся к преподавателям, упомянутым НАД комментарием.";

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            //Считаем количество строк
            colRows = DispatchWordCountRows();

            //Вставляем таблицу colRows (строки) x 11 (столбцы), заполняем её данными о преподавателе
            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjTable = ObjDoc.Tables.Add(ObjWordRange, colRows, 11, ref ObjMissing, ref ObjMissing);
            //Добавляем обрамление таблицы
            ObjTable.Borders.Enable = 1;

            //Текущей строкой считаем первую
            curRow = 1;
            //Заполняем первую строку и одновременно формируем размерности
            ObjTable.Cell(curRow, 1).Range.Text = "Преподаватель";
            ObjTable.Cell(curRow, 2).Range.Text = "Курс";
            ObjTable.Cell(curRow, 3).Range.Text = "Специальность";
            ObjTable.Cell(curRow, 4).Range.Text = "Название дисциплины";
            ObjTable.Cell(curRow, 5).Range.Text = "Лекция";
            ObjTable.Cell(curRow, 6).Range.Text = "Практ.зан.";
            ObjTable.Cell(curRow, 7).Range.Text = "Лаб.раб";
            ObjTable.Cell(curRow, 8).Range.Text = "Курс.пр.";
            ObjTable.Cell(curRow, 9).Range.Text = "Практика";
            ObjTable.Cell(curRow, 10).Range.Text = "";
            ObjTable.Cell(curRow, 11).Range.Text = "Примечания*";

            ObjTable.Rows[curRow].Range.Font.Bold = 1;

            curRow++;

            Both = false;
            if (mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum == "-")
            {
                Both = true;
            }

            //Просматриваем каждого преподавателя
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                //По умолчанию считаем, что у преподавателя нет нагрузки
                HaveLoad = false;
                //По умолчанию считаем, что ячейки в столбце с преподавателем
                //надо объединить
                MergeCells = true;
                //Просматриваем нагрузку
                for (int j = 0; j <= mdlData.colDistribution.Count - 1; j++)
                {
                    //Если в комбинированном списке выбран
                    //конкретный семестр (не оба семестра)
                    if (!Both)
                    {
                        if (!(mdlData.colDistribution[j].Semestr == null))
                        {
                            //Если дисциплина относится к выбранному семестру
                            if (mdlData.colDistribution[j].Semestr.Equals(mdlData.colSemestr[cmbSemestrList.SelectedIndex]))
                            {
                                //Надо проверить, записывать дублёра или основного преподавателя?
                                if (mdlData.colDistribution[j].Doubler != null)
                                {
                                    flgDoubler = mdlData.colLecturer[i].Equals(mdlData.colDistribution[j].Doubler);
                                }
                                else
                                {
                                    flgDoubler = mdlData.colLecturer[i].Equals(mdlData.colDistribution[j].Lecturer);
                                }
                                
                                //Если рассматриваемый преподаватель совпадает с
                                //преподавателем, указанным в нагрузке
                                if (flgDoubler)
                                {
                                    //Если дисциплина предполагает лекционные часы,
                                    //практические занятия, лабораторные работы,
                                    //курсовой проект, учебную практику, производственную
                                    //практику или дипломную практику, то
                                    //выделяем под неё строчку текста
                                    if (mdlData.NonZeroForDispatchOR(mdlData.colDistribution[j]))
                                    {
                                        //Если элемент распределения нагрузки необходимо внести
                                        //в заявку для диспетчерской
                                        if (mdlData.colDistribution[j].flgDispatch)
                                        {
                                            //Если у преподавателя нет нагрузки, но мы попали сюда,
                                            //то надо написать его имя в отчётную форму
                                            if (!HaveLoad)
                                            {
                                                //Если у преподавателя нет дублёра,
                                                //то в заявку идёт он сам
                                                if (mdlData.colDistribution[j].Doubler == null)
                                                {
                                                    //Печатаем фамилию, имя, отчество преподавателя
                                                    ObjTable.Cell(curRow, 1).Range.Text = mdlData.colLecturer[i].FIO + " (" +
                                                        mdlData.colLecturer[i].Duty.Short + ")";
                                                }
                                                //Если у преподавателя есть дублёр, то в заявку идёт
                                                //дублёр
                                                else
                                                {
                                                    //Печатаем фамилию, имя, отчество преподавателя
                                                    ObjTable.Cell(curRow, 1).Range.Text = mdlData.colDistribution[j].Doubler.FIO + " (" +
                                                        mdlData.colDistribution[j].Doubler.Duty.Short + ")";
                                                    //Хоть и известно, что у преподавателя есть нагрузка, но признак
                                                    //её наличия не выставляем
                                                }
                                                
                                                //Если мы сюда попали, то получается, что у преподавателя есть нагрузка
                                                HaveLoad = true;

                                                //Выделяем имя преподавателя жирным шрифтом
                                                ObjTable.Cell(curRow, 1).Range.Font.Bold = 1;
                                                //Фиксируем первую строку, выделяемую на преподавателя
                                                stRow = curRow;
                                            }

                                            //Печатаем курс
                                            ObjTable.Cell(curRow, 2).Range.Text = mdlData.colDistribution[j].KursNum.Kurs.ToString();
                                            //Печатаем специальность
                                            ObjTable.Cell(curRow, 3).Range.Text = mdlData.colDistribution[j].Speciality.ShortUpravlenie
                                                + " (" + mdlData.colDistribution[j].Speciality.ShortInstitute + ")";
                                            //Печатаем название дисциплины
                                            ObjTable.Cell(curRow, 4).Range.Text = mdlData.colDistribution[j].Subject.Subject;
                                            //Лекционные часы
                                            ObjTable.Cell(curRow, 5).Range.Text = mdlData.colDistribution[j].Lecture.ToString();
                                            //Практические занятия в часах
                                            ObjTable.Cell(curRow, 6).Range.Text = mdlData.colDistribution[j].Practice.ToString();
                                            //Лабораторные работы в часах
                                            ObjTable.Cell(curRow, 7).Range.Text = mdlData.colDistribution[j].LabWork.ToString();
                                            //Курсовой проект в часах
                                            ObjTable.Cell(curRow, 8).Range.Text = mdlData.colDistribution[j].KursProject.ToString();
                                            //Практика в часах
                                            //Либо записываем преддипломную практику
                                            if (mdlData.colDistribution[j].PreDiplomaPractice > 0)
                                            {
                                                ObjTable.Cell(curRow, 9).Range.Text = mdlData.colDistribution[j].PreDiplomaPractice.ToString();
                                            }
                                            //Либо записываем учебную практику
                                            if (mdlData.colDistribution[j].TutorialPractice > 0)
                                            {
                                                ObjTable.Cell(curRow, 9).Range.Text = mdlData.colDistribution[j].TutorialPractice.ToString();
                                            }
                                            //Либо записываем производственную практику
                                            if (mdlData.colDistribution[j].ProducingPractice > 0)
                                            {
                                                ObjTable.Cell(curRow, 9).Range.Text = mdlData.colDistribution[j].ProducingPractice.ToString();
                                            }
                                            //Если поле практики так и осталось пустым, то
                                            //заполняем его нулём
                                            if (ObjTable.Cell(curRow, 9).Range.Text == "\r\a")
                                            {
                                                ObjTable.Cell(curRow, 9).Range.Text = (0).ToString();
                                            }

                                            //Записываем содержимое пожеланий по аудиториям
                                            ObjTable.Cell(curRow, 11).Range.Text = mdlData.colDistribution[j].Text;

                                            //Если имеется связка по лабораторным работам, то указываем в примечании
                                            //с кем в паре проводятся лабораторные работы (электроника)
                                            if (mdlData.colDistribution[j].LabWorkConnect != null)
                                            {
                                                //Если указан дублёр, то ставим дублёра
                                                if (mdlData.colDistribution[j].LabWorkConnect.Doubler != null)
                                                {
                                                    ObjTable.Cell(curRow, 11).Range.Text += " (совместно с " +
                                                        mdlData.SplitFIOString(mdlData.colDistribution[j].LabWorkConnect.Doubler.FIO, true, false) + ")";
                                                }
                                                //В ином случае указывается основной преподаватель
                                                else
                                                {
                                                    ObjTable.Cell(curRow, 11).Range.Text += " (совместно с " +
                                                        mdlData.SplitFIOString(mdlData.colDistribution[j].LabWorkConnect.Lecturer.FIO, true, false) + ")";
                                                }
                                            }

                                            //Переходим к следующей строке
                                            curRow += 1;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //Если семестр не выбран, то необходимо предоставить информацию
                    //по обоим семестрам
                    else
                    {
                        //Если рассматриваемый преподаватель совпадает с
                        //преподавателем, указанным в нагрузке
                        if (mdlData.colLecturer[i].Equals(mdlData.colDistribution[j].Lecturer))
                        {
                            //Если дисциплина предполагает лекционные часы,
                            //практические занятия, лабораторные работы,
                            //курсовой проект, учебную практику, производственную
                            //практику или дипломную практику, то
                            //выделяем под неё строчку текста
                            if (mdlData.NonZeroForDispatchOR(mdlData.colDistribution[j]))
                            {
                                if (mdlData.colDistribution[j].flgDispatch)
                                {
                                    //Если выясняется, что у преподавателя есть дублёр,
                                    //то необходимо сбросить признак наличия нагрузки
                                    if (mdlData.colDistribution[j].Doubler != null)
                                    {
                                        HaveLoad = false;
                                    }

                                    //Если у преподавателя нет нагрузки, но мы попали сюда,
                                    //то надо написать его имя в отчётную форму
                                    if (!HaveLoad)
                                    {
                                        //Если у преподавателя нет дублёра,
                                        //то в заявку идёт он сам
                                        if (mdlData.colDistribution[j].Doubler == null)
                                        {
                                            //Печатаем фамилию, имя, отчество преподавателя
                                            ObjTable.Cell(curRow, 1).Range.Text = mdlData.colLecturer[i].FIO + " (" +
                                                        mdlData.colLecturer[i].Duty.Short + ")";
                                            //Получается, что у преподавателя есть нагрузка
                                            HaveLoad = true;
                                        }
                                        //Если у преподавателя есть дублёр, то в заявку идёт
                                        //дублёр
                                        else
                                        {
                                            //Печатаем фамилию, имя, отчество преподавателя
                                            ObjTable.Cell(curRow, 1).Range.Text = mdlData.colDistribution[j].Doubler.FIO + " (" +
                                                mdlData.colDistribution[j].Doubler.Duty.Short + ")";
                                            //Хоть и известно, что у преподавателя есть нагрузка, но признак
                                            //её наличия не выставляем
                                        }
                                        //
                                        ObjTable.Cell(curRow, 1).Range.Font.Bold = 1;
                                        //
                                        stRow = curRow;
                                    }
                                    //Печатаем курс
                                    ObjTable.Cell(curRow, 2).Range.Text = mdlData.colDistribution[j].KursNum.Kurs.ToString();
                                    //Печатаем специальность
                                    ObjTable.Cell(curRow, 3).Range.Text = mdlData.colDistribution[j].Speciality.ShortUpravlenie
                                        + " (" + mdlData.colDistribution[j].Speciality.ShortInstitute + ")";
                                    //Печатаем название дисциплины
                                    ObjTable.Cell(curRow, 4).Range.Text = mdlData.colDistribution[j].Subject.Subject;
                                    //Лекционные часы
                                    ObjTable.Cell(curRow, 5).Range.Text = mdlData.colDistribution[j].Lecture.ToString();
                                    //Практические занятия в часах
                                    ObjTable.Cell(curRow, 6).Range.Text = mdlData.colDistribution[j].Practice.ToString();
                                    //Лабораторные работы в часах
                                    ObjTable.Cell(curRow, 7).Range.Text = mdlData.colDistribution[j].LabWork.ToString();
                                    //Курсовой проект в часах
                                    ObjTable.Cell(curRow, 8).Range.Text = mdlData.colDistribution[j].KursProject.ToString();
                                    //Практика в часах
                                    //Либо записываем преддипломную практику
                                    if (mdlData.colDistribution[j].PreDiplomaPractice > 0)
                                    {
                                        ObjTable.Cell(curRow, 9).Range.Text = mdlData.colDistribution[j].PreDiplomaPractice.ToString();
                                    }
                                    //Либо записываем учебную практику
                                    if (mdlData.colDistribution[j].TutorialPractice > 0)
                                    {
                                        ObjTable.Cell(curRow, 9).Range.Text = mdlData.colDistribution[j].TutorialPractice.ToString();
                                    }
                                    //Либо записываем производственную практику
                                    if (mdlData.colDistribution[j].ProducingPractice > 0)
                                    {
                                        ObjTable.Cell(curRow, 9).Range.Text = mdlData.colDistribution[j].ProducingPractice.ToString();
                                    }
                                    //Если поле практики так и осталось пустым, то
                                    //заполняем его нулём
                                    if (ObjTable.Cell(curRow, 9).Range.Text == "\r\a")
                                    {
                                        ObjTable.Cell(curRow, 9).Range.Text = (0).ToString();
                                    }
                                    //
                                    ObjTable.Cell(curRow, 11).Range.Text = mdlData.colDistribution[j].Text;
                                    //Переходим к следующей строке
                                    curRow += 1;
                                }
                            }
                        }
                    }
                }

                //Только, если у преподавателя есть нагрузка
                if (HaveLoad)
                {
                    //Пишем примечание конкретного преподавателя
                    ObjTable.Cell(curRow, 11).Range.Text = mdlData.colLecturer[i].Text;
                    //Фиксируем последнюю строку, выделяемую на преподавателя
                    endRow = curRow;
                    if (MergeCells)
                    {
                        //Объединяем ячейки в столбце преподавателей
                        ObjTable.Cell(stRow, 1).Merge(ObjTable.Cell(endRow, 1));
                    }
                    //Разделяем преподавателей пустой строчкой
                    curRow += 1;
                }
            }

            ObjTable.Columns[1].Width = ObjWord.CentimetersToPoints(3.44f);
            ObjTable.Columns[2].Width = ObjWord.CentimetersToPoints(1.5f);
            ObjTable.Columns[3].Width = ObjWord.CentimetersToPoints(2.5f);
            ObjTable.Columns[4].Width = ObjWord.CentimetersToPoints(8f);
            ObjTable.Columns[5].Width = ObjWord.CentimetersToPoints(1.25f);
            ObjTable.Columns[6].Width = ObjWord.CentimetersToPoints(1.75f);
            ObjTable.Columns[7].Width = ObjWord.CentimetersToPoints(1.25f);
            ObjTable.Columns[8].Width = ObjWord.CentimetersToPoints(1.5f);
            ObjTable.Columns[9].Width = ObjWord.CentimetersToPoints(1.5f);
            ObjTable.Columns[10].Width = ObjWord.CentimetersToPoints(0.5f);
            ObjTable.Columns[11].Width = ObjWord.CentimetersToPoints(5.5f);

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            ObjParagraph.Range.Text = "Заведующий кафедрой УиЗИ \t" + "/ Л.А. Баранов /";
            //ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjParagraph.Range.ParagraphFormat.TabStops.Add(ObjWord.CentimetersToPoints(20f), Word.WdTabAlignment.wdAlignTabRight);

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            /*
            ObjParagraph.Range.Text = "Директор ИТТСУ \t" + "/ П.Ф. Бестемьянов /";

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();
            */

            ObjParagraph.Range.Text = "Первый зам. директора ИТТСУ - " +
                "начальник учебного отдела \t" + "/ В.А. Гречишников /";

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            ObjParagraph.Range.Text = "И.о. декана вечернего факультета \t" + "/ И.В. Федякин /";

            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertParagraphAfter();

            //ObjParagraph.Range.Text = "Директор ИУИТ \t" + "/ С.П. Вакуленко /";

            if (cmbSemestrList.SelectedIndex == 1)
            {
                ObjParagraph.Range.Text = "Первый зам. директора ИУИТ - " +
                    "начальник учебного отдела \t" + "/ Е.С. Прокофьева /";
            }

            ObjDoc.SaveAs(Application.StartupPath + @"\"
                + "Заявка в диспетчерскую " + mdlData.colSemestr[cmbSemestrList.SelectedIndex].About
                + " " + mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear.Replace("/", "-") + " уч. года " +
                DateTime.Now.Date.ToString("yyyyMMdd") + " " +
                DateTime.Now.TimeOfDay.ToString("hhmmss") + ".docx");

            ObjDoc.Close();
        }

        private void DispatchWord()
        {
            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;

            try
            {
                //Создаём новое Word приложение
                Word._Application ObjWord = new Word.Application();

                //DispatchWordCore(ObjMissing, ObjWord);
                DispatchWordCoreNew(ObjMissing, ObjWord);

                ObjWord.Quit();
            }
            catch
            {
                MessageBox.Show("Возможно на этом компьютере присутствует проблема совместимости с MS Word." +
                " Попробуйте установить версию 2007 и выше.");
            }
        }

        private void UpravlenieDoneExcel()
        {
            int curLect;

            //---------------------------------

            int sumLecI; int sumLecII;

            int sumExamI; int sumExamII;

            int sumCredI; int sumCredII;

            int sumTutI; int sumTutII;

            int sumLabI; int sumLabII;

            int sumPracI; int sumPracII;

            int sumRefI; int sumRefII;

            int sumIndI; int sumIndII;

            int sumKRAPKI; int sumKRAPKII;

            int sumKursPrI; int sumKursPrII;

            int sumDiplI; int sumDiplII;

            int sumTutPrI; int sumTutPrII;

            int sumPreDipI; int sumPreDipII;

            int sumGAKI; int sumGAKII;

            int sumPostGrI; int sumPostGrII;

            int sumVisI; int sumVisII;

            int sumMagI; int sumMagII;

            //---------------------------------

            int sumLecIplan; int sumLecIIplan;

            int sumExamIplan; int sumExamIIplan;

            int sumCredIplan; int sumCredIIplan;

            int sumTutIplan; int sumTutIIplan;

            int sumLabIplan; int sumLabIIplan;

            int sumPracIplan; int sumPracIIplan;

            int sumRefIplan; int sumRefIIplan;

            int sumIndIplan; int sumIndIIplan;

            int sumKRAPKIplan; int sumKRAPKIIplan;

            int sumKursPrIplan; int sumKursPrIIplan;

            int sumDiplIplan; int sumDiplIIplan;

            int sumTutPrIplan; int sumTutPrIIplan;

            int sumPreDipIplan; int sumPreDipIIplan;

            int sumGAKIplan; int sumGAKIIplan;

            int sumPostGrIplan; int sumPostGrIIplan;

            int sumVisIplan; int sumVisIIplan;

            int sumMagIplan; int sumMagIIplan;

            //---------------------------------

            int countStud, countWeight;

            string strSumI = "=";
            string strSumII = "=";
            string strSumIII = "=";
            string strSem;
            string strSemPlan;
            string fileNameAdd = "";

            double[] rateParams = new double[4];
            double assist;
            double hitutor;
            double proff;
            double lecturer;
            double sumRate;

            IList<clsDistribution> coll = null;
            IList<clsDistribution> collPlan = null;

            bool accessFlg = false;
            bool flgDoubler;
            bool flg;

            rateParams = countRates();

            assist = rateParams[0];
            lecturer = rateParams[1];
            hitutor = rateParams[2];
            proff = rateParams[3];

            sumRate = proff + hitutor + lecturer + assist;

            if (optMain.Checked || optMainDop.Checked)
            {
                coll = mdlData.colDistribution;
                fileNameAdd = " штатная ";
            }
            else
            {
                if (optHoured.Checked)
                {
                    coll = mdlData.colHouredDistribution;
                    collPlan = mdlData.colPlanHouredDistribution;
                    fileNameAdd = " почасовая ";
                }
                else
                {
                    if (optCombine.Checked || optCombineDop.Checked)
                    {
                        coll = mdlData.colCombineDistribution;
                        collPlan = mdlData.colPlanCombineDistribution;
                        fileNameAdd = " штатная с учётом почасовой ";
                    }
                }
            }

            //Если в перечне семестров указан не прочерк
            //то формируем документ в зависимости от сделанного выбора
            if (cmbSemestrList.SelectedIndex > 0)
            {
                //Задаём переменную для отсутствующего параметра
                object ObjMissing = Missing.Value;

                try
                {
                    //Создаём новое Excel приложение
                    Excel.Application ObjExcel = new Excel.Application();
                    Excel.Workbook ObjWorkBook;
                    Excel.Worksheet ObjWorkSheet;

                    //Не отслеживать заполнение таблицы
                    ObjExcel.Visible = false;

                    //Книга
                    ObjWorkBook = ObjExcel.Workbooks.Add(Missing.Value);
                    //Таблица
                    ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
                    
                    //Альбомная ориентация страницы
                    ObjWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                    //Книжная ориентация страницы
                    //ObjWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                    //В высоту разместить на одной странице
                    ObjWorkSheet.PageSetup.FitToPagesTall = 1;
                    //В ширину разместить на одной странице
                    ObjWorkSheet.PageSetup.FitToPagesWide = 1;

                    //-----------Формируем заголовок таблицы

                    //Задаём диапазон для ячеек, подлежащих форматированию
                    //1-я строка, с А по AA
                    var cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(1, 1),
                        mdlData.ExcelCellTranslator(1, 27));
                    //Выделяем ячейки диапазона
                    cells.Select();
                    //Объединяем ячейки диапазона
                    cells.Merge(true);
                    //Выравнивание в объединённой ячейке по центру
                    cells.HorizontalAlignment = Excel.Constants.xlCenter;

                    //Записываем текст в объединённую ячейку
                    //(считается по первой А1)

                    //Если исполненная
                    cells.Cells[1, 1] = "Сведения о фактически выполненной учебной нагрузке преподавателя за " + mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum;

                    //Задаём диапазон для ячеек, подлежащих форматированию
                    //2-я строка, с А по AA
                    cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(2, 1),
                        mdlData.ExcelCellTranslator(2, 27));
                    //Выделяем ячейки диапазона
                    cells.Select();
                    //Объединяем ячейки диапазона
                    cells.Merge(true);
                    //Выравнивание в объединённой ячейке по центру
                    cells.HorizontalAlignment = Excel.Constants.xlCenter;
                    //Записываем текст в объединённую ячейку
                    //(считается по первой А2)
                    cells.Cells[1, 1] = "Кафедра \"Управление и защита информации\" " +
                        mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear +
                        " учебный год";

                    //-----------Формируем заголовок таблицы

                    //-----------Формируем шапку таблицы

                    for (int i = 1; i <= 27; i++)
                    {
                        //Выбираем диапазон
                        cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4, i),
                            mdlData.ExcelCellTranslator(5, i));
                        //Выделяем ячейки диапазона
                        cells.Select();
                        //Объединяем ячейки диапазона
                        cells.Merge();
                        //Горизонтальное выравнивание по центру в ячейках
                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                        //Вертикальное выравнивание по центру в ячейках
                        cells.VerticalAlignment = Excel.Constants.xlCenter;
                        //Задаём границы
                        cells.Borders.Weight = 2;
                        //Перенос по словам без сокрытия под ячейками
                        cells.WrapText = true;
                        //Назначаем высоту шапки
                        cells.RowHeight = (68.25f + 66) / 2;

                        switch (i)
                        {
                            case 1:
                                cells.Cells[1, 1] = "№ п/п";
                                cells.ColumnWidth = 3.86f;
                                cells.Orientation = Excel.XlOrientation.xlHorizontal;
                                break;
                            case 2:
                                cells.Cells[1, 1] = "Фамилия, Имя, Отчество";
                                cells.ColumnWidth = 27.57f;
                                cells.Orientation = Excel.XlOrientation.xlHorizontal;
                                break;
                            case 3:
                                cells.Cells[1, 1] = "Ставка";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 4:
                                cells.Cells[1, 1] = "Лекции";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 5:
                                cells.Cells[1, 1] = "Экзамены";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 6:
                                cells.Cells[1, 1] = "Зачеты";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 7:
                                cells.Cells[1, 1] = "ПК";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 8:
                                cells.Cells[1, 1] = "Консультации";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 9:
                                cells.Cells[1, 1] = "Практические занятия";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 10:
                                cells.Cells[1, 1] = "Домашние задания и рефераты";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 11:
                                cells.Cells[1, 1] = "Текущая аттестация";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 12:
                                cells.Cells[1, 1] = "Индивидуальные занятия";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 13:
                                cells.Cells[1, 1] = "Контрольные работы";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 14:
                                cells.Cells[1, 1] = "Курсовой проект, курсовая работа";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 15:
                                cells.Cells[1, 1] = "Дипломный проект";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 16:
                                cells.Cells[1, 1] = "Учебная практика";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 17:
                                cells.Cells[1, 1] = "Преддипломная и производственная практика";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 18:
                                cells.Cells[1, 1] = "ГЭК";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 19:
                                cells.Cells[1, 1] = "Приёмная комиссия";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 20:
                                cells.Cells[1, 1] = "Лабораторные работы";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 21:
                                cells.Cells[1, 1] = "Аспирантура";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 22:
                                cells.Cells[1, 1] = "Посещение занятий";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 23:
                                cells.Cells[1, 1] = "Другие виды занятий";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            //Здесь указаны плановые часы
                            case 24:
                                cells.UnMerge();
                                cells.Borders.Weight = 2;
                                cells.Cells[1, 1] = mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum;
                                cells.Cells[2, 1] = "План";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            //Здесь указаны фактические часы
                            case 25:
                                cells.UnMerge();
                                cells.Borders.Weight = 2;
                                cells.Cells[2, 1] = "Фактич. выполн.";
                                cells.Orientation = Excel.XlOrientation.xlUpward;
                                break;
                            case 26:
                                cells.Cells[1, 1] = "+\n\n\n\n\n\n-";
                                cells.Orientation = Excel.XlOrientation.xlHorizontal;
                                cells.Borders[Excel.XlBordersIndex.xlDiagonalUp].Weight = 2;
                                break;
                            case 27:
                                cells.Cells[1, 1] = "Примечание";
                                cells.ColumnWidth = 12.00f;
                                cells.Orientation = Excel.XlOrientation.xlHorizontal;
                                break;
                        }
                    }

                    //Выбираем диапазон
                    cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4, 24),
                        mdlData.ExcelCellTranslator(4, 25));
                    //Выделяем ячейки диапазона
                    cells.Select();
                    //Объединяем ячейки диапазона
                    cells.Merge();
                    cells.Orientation = Excel.XlOrientation.xlHorizontal;

                    //-----------Формируем шапку таблицы

                    //-----------Заполняем таблицу данными

                    curLect = 1;

                    for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
                    {
                        flg = false;
                        //Сначала определяем, нужно ли, в принципе, 
                        //рассматривать преподавателя
                        if (optMain.Checked || optCombine.Checked)
                        {
                            accessFlg = (mdlData.colLecturer[i].Rate > 0);
                        }
                        else
                        {
                            if (optHoured.Checked || optMainDop.Checked || optCombineDop.Checked)
                            {
                                accessFlg = true;
                            }
                        }

                        //Если прошёл отбор
                        if (accessFlg)
                        {
                            sumLecI = 0; sumLecII = 0;

                            sumExamI = 0; sumExamII = 0;

                            sumCredI = 0; sumCredII = 0;

                            sumTutI = 0; sumTutII = 0;

                            sumLabI = 0; sumLabII = 0;

                            sumPracI = 0; sumPracII = 0;

                            sumRefI = 0; sumRefII = 0;

                            sumIndI = 0; sumIndII = 0;

                            sumKRAPKI = 0; sumKRAPKII = 0;

                            sumKursPrI = 0; sumKursPrII = 0;

                            sumDiplI = 0; sumDiplII = 0;

                            sumTutPrI = 0; sumTutPrII = 0;

                            sumPreDipI = 0; sumPreDipII = 0;

                            sumGAKI = 0; sumGAKII = 0;

                            sumPostGrI = 0; sumPostGrII = 0;

                            sumVisI = 0; sumVisII = 0;

                            sumMagI = 0; sumMagII = 0;

                            //Просматриваем нагрузку (фактическую)
                            for (int j = 0; j <= coll.Count - 1; j++)
                            {
                                //Если семестр соответствует выбранному из списка
                                if (coll[j].Semestr.Equals(mdlData.colSemestr[cmbSemestrList.SelectedIndex]))
                                {
                                    //Если строка нагрузки принудительно не исключена из рассмотрения
                                    if (!coll[j].flgExclude)
                                    {
                                        //Если указан какой-либо преподаватель, ведущий дисциплину
                                        if (!(coll[j].Lecturer == null))
                                        {
                                            //Если работаем с почасовой нагрузкой, то в ней может быть указан
                                            //дублёр, то есть тот, кто реально читал лекции вместо заболевшего
                                            //или уволенного преподавателя

                                            if (optHoured.Checked)
                                            {
                                                if (coll[j].Doubler != null)
                                                {
                                                    flgDoubler = coll[j].Doubler.Equals(mdlData.colLecturer[i]);
                                                }
                                                else
                                                {
                                                    flgDoubler = coll[j].Lecturer.Equals(mdlData.colLecturer[i]);
                                                }
                                            }
                                            else
                                            {
                                                flgDoubler = coll[j].Lecturer.Equals(mdlData.colLecturer[i]);
                                            }

                                            //Принимаем решение о правильности выбора преподавателя
                                            if (flgDoubler)
                                            {
                                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                                {
                                                    sumLecI += coll[j].Lecture;
                                                    sumExamI += coll[j].Exam;
                                                    sumCredI += coll[j].Credit;
                                                    sumTutI += coll[j].Tutorial;
                                                    sumLabI += coll[j].LabWork;
                                                    sumPracI += coll[j].Practice;
                                                    sumRefI += coll[j].RefHomeWork;
                                                    sumIndI += coll[j].IndividualWork;
                                                    sumKRAPKI += coll[j].KRAPK;
                                                    sumKursPrI += coll[j].KursProject;
                                                    sumDiplI += coll[j].DiplomaPaper;
                                                    sumTutPrI += coll[j].TutorialPractice;
                                                    sumPreDipI += coll[j].PreDiplomaPractice +
                                                        coll[j].ProducingPractice;
                                                    sumGAKI += coll[j].GAK;
                                                    sumPostGrI += coll[j].PostGrad;
                                                    sumVisI += coll[j].Visiting;
                                                    sumMagI += coll[j].Magistry;

                                                    if (optMainDop.Checked || optCombineDop.Checked)
                                                    {
                                                        flg = true;
                                                    }
                                                    else
                                                    {
                                                        if (optHoured.Checked || optMain.Checked || optCombine.Checked)
                                                        {
                                                            if (!flg)
                                                            {
                                                                flg = (mdlData.toSumDistributionComponents(coll[j]) != 0);
                                                            }
                                                        }
                                                    }
                                                }

                                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                                {
                                                    sumLecII += coll[j].Lecture;
                                                    sumExamII += coll[j].Exam;
                                                    sumCredII += coll[j].Credit;
                                                    sumTutII += coll[j].Tutorial;
                                                    sumLabII += coll[j].LabWork;
                                                    sumPracII += coll[j].Practice;
                                                    sumRefII += coll[j].RefHomeWork;
                                                    sumIndII += coll[j].IndividualWork;
                                                    sumKRAPKII += coll[j].KRAPK;
                                                    sumKursPrII += coll[j].KursProject;
                                                    sumDiplII += coll[j].DiplomaPaper;
                                                    sumTutPrII += coll[j].TutorialPractice;
                                                    sumPreDipII += coll[j].PreDiplomaPractice +
                                                        coll[j].ProducingPractice;
                                                    sumGAKII += coll[j].GAK;
                                                    sumPostGrII += coll[j].PostGrad;
                                                    sumVisII += coll[j].Visiting;
                                                    sumMagII += coll[j].Magistry;

                                                    if (optMainDop.Checked || optCombineDop.Checked)
                                                    {
                                                        flg = true;
                                                    }
                                                    else
                                                    {
                                                        if (optHoured.Checked || optMain.Checked || optCombine.Checked)
                                                        {
                                                            if (!flg)
                                                            {
                                                                flg = (mdlData.toSumDistributionComponents(coll[j]) != 0);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        //Если преподаватель не указан
                                        else
                                        {
                                            //а нагрузка равномерно распределяемая
                                            if (coll[j].flgDistrib)
                                            {
                                                countWeight = 0;
                                                countStud = 0;

                                                for (int k = 0; k <= mdlData.colStudents.Count - 1; k++)
                                                {
                                                    if (mdlData.colStudents[k].flgPlan)
                                                    {
                                                        //Если рассматриваемый преподаватель - руководитель студента
                                                        //И если студент на том же курсе, где и дисциплина
                                                        //И специальность студента должна соответствовать специальности нагрузки
                                                        if (mdlData.colStudents[k].Lect.Equals(mdlData.colLecturer[i])
                                                            & mdlData.colStudents[k].KursNum.Equals(coll[j].KursNum)
                                                            & mdlData.colStudents[k].Speciality.Equals(coll[j].Speciality))
                                                        {
                                                            countWeight += coll[j].Weight;
                                                            countStud++;
                                                        }
                                                    }
                                                }

                                                //
                                                if (flgCombine)
                                                {
                                                    mdlData.toDetectUniformInHoured(ref countWeight, coll[j], mdlData.colLecturer[i]);
                                                }

                                                //
                                                if (countWeight > 0)
                                                {
                                                    flg = true;

                                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                                    {
                                                        if (coll[j].PreDiplomaPractice > 0 ||
                                                            coll[j].ProducingPractice > 0)
                                                        {
                                                            sumPreDipI += countWeight;
                                                        }

                                                        if (coll[j].TutorialPractice > 0)
                                                        {
                                                            sumTutPrI += countWeight;
                                                        }

                                                        if (coll[j].DiplomaPaper > 0)
                                                        {
                                                            sumDiplI += countWeight;
                                                        }

                                                        if (coll[j].Magistry > 0)
                                                        {
                                                            sumMagI += countWeight;
                                                        }
                                                    }

                                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                                    {
                                                        if (coll[j].PreDiplomaPractice > 0 ||
                                                            coll[j].ProducingPractice > 0)
                                                        {
                                                            sumPreDipII += countWeight;
                                                        }

                                                        if (coll[j].TutorialPractice > 0)
                                                        {
                                                            sumTutPrII += countWeight;
                                                        }

                                                        if (coll[j].DiplomaPaper > 0)
                                                        {
                                                            sumDiplII += countWeight;
                                                        }

                                                        if (coll[j].Magistry > 0)
                                                        {
                                                            sumMagII += countWeight;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            sumLecIplan = 0; sumLecIIplan = 0;

                            sumExamIplan = 0; sumExamIIplan = 0;

                            sumCredIplan = 0; sumCredIIplan = 0;

                            sumTutIplan = 0; sumTutIIplan = 0;

                            sumLabIplan = 0; sumLabIIplan = 0;

                            sumPracIplan = 0; sumPracIIplan = 0;

                            sumRefIplan = 0; sumRefIIplan = 0;

                            sumIndIplan = 0; sumIndIIplan = 0;

                            sumKRAPKIplan = 0; sumKRAPKIIplan = 0;

                            sumKursPrIplan = 0; sumKursPrIIplan = 0;

                            sumDiplIplan = 0; sumDiplIIplan = 0;

                            sumTutPrIplan = 0; sumTutPrIIplan = 0;

                            sumPreDipIplan = 0; sumPreDipIIplan = 0;

                            sumGAKIplan = 0; sumGAKIIplan = 0;

                            sumPostGrIplan = 0; sumPostGrIIplan = 0;

                            sumVisIplan = 0; sumVisIIplan = 0;

                            sumMagIplan = 0; sumMagIIplan = 0;

                            //Если имеются сведения о плановой нагрузке
                            if (collPlan != null)
                            {
                                //Просматриваем нагрузку (фактическую)
                                for (int j = 0; j <= collPlan.Count - 1; j++)
                                {
                                    //Если семестр соответствует выбранному из списка
                                    if (collPlan[j].Semestr.Equals(mdlData.colSemestr[cmbSemestrList.SelectedIndex]))
                                    {
                                        //Если строка нагрузки принудительно не исключена из рассмотрения
                                        if (!collPlan[j].flgExclude)
                                        {
                                            //Если указан какой-либо преподаватель, ведущий дисциплину
                                            if (!(collPlan[j].Lecturer == null))
                                            {
                                                //Если работаем с почасовой нагрузкой, то в ней может быть указан
                                                //дублёр, то есть тот, кто реально читал лекции вместо заболевшего
                                                //или уволенного преподавателя

                                                if (optHoured.Checked)
                                                {
                                                    if (collPlan[j].Doubler != null)
                                                    {
                                                        flgDoubler = collPlan[j].Doubler.Equals(mdlData.colLecturer[i]);
                                                    }
                                                    else
                                                    {
                                                        flgDoubler = collPlan[j].Lecturer.Equals(mdlData.colLecturer[i]);
                                                    }
                                                }
                                                else
                                                {
                                                    flgDoubler = collPlan[j].Lecturer.Equals(mdlData.colLecturer[i]);
                                                }

                                                //Принимаем решение о правильности выбора преподавателя
                                                if (flgDoubler)
                                                {
                                                    if (collPlan[j].Semestr.Equals(mdlData.colSemestr[1]))
                                                    {
                                                        sumLecIplan += collPlan[j].Lecture;
                                                        sumExamIplan += collPlan[j].Exam;
                                                        sumCredIplan += collPlan[j].Credit;
                                                        sumTutIplan += collPlan[j].Tutorial;
                                                        sumLabIplan += collPlan[j].LabWork;
                                                        sumPracIplan += collPlan[j].Practice;
                                                        sumRefIplan += collPlan[j].RefHomeWork;
                                                        sumIndIplan += collPlan[j].IndividualWork;
                                                        sumKRAPKIplan += collPlan[j].KRAPK;
                                                        sumKursPrIplan += collPlan[j].KursProject;
                                                        sumDiplIplan += collPlan[j].DiplomaPaper;
                                                        sumTutPrIplan += collPlan[j].TutorialPractice;
                                                        sumPreDipIplan += collPlan[j].PreDiplomaPractice +
                                                            collPlan[j].ProducingPractice;
                                                        sumGAKIplan += collPlan[j].GAK;
                                                        sumPostGrIplan += collPlan[j].PostGrad;
                                                        sumVisIplan += collPlan[j].Visiting;
                                                        sumMagIplan += collPlan[j].Magistry;
                                                    }

                                                    if (collPlan[j].Semestr.Equals(mdlData.colSemestr[2]))
                                                    {
                                                        sumLecIIplan += collPlan[j].Lecture;
                                                        sumExamIIplan += collPlan[j].Exam;
                                                        sumCredIIplan += collPlan[j].Credit;
                                                        sumTutIIplan += collPlan[j].Tutorial;
                                                        sumLabIIplan += collPlan[j].LabWork;
                                                        sumPracIIplan += collPlan[j].Practice;
                                                        sumRefIIplan += collPlan[j].RefHomeWork;
                                                        sumIndIIplan += collPlan[j].IndividualWork;
                                                        sumKRAPKIIplan += collPlan[j].KRAPK;
                                                        sumKursPrIIplan += collPlan[j].KursProject;
                                                        sumDiplIIplan += collPlan[j].DiplomaPaper;
                                                        sumTutPrIIplan += collPlan[j].TutorialPractice;
                                                        sumPreDipIIplan += collPlan[j].PreDiplomaPractice +
                                                            collPlan[j].ProducingPractice;
                                                        sumGAKIIplan += collPlan[j].GAK;
                                                        sumPostGrIIplan += collPlan[j].PostGrad;
                                                        sumVisIIplan += collPlan[j].Visiting;
                                                        sumMagIIplan += collPlan[j].Magistry;
                                                    }
                                                }
                                            }
                                            //Если преподаватель не указан
                                            else
                                            {
                                                //а нагрузка равномерно распределяемая
                                                if (collPlan[j].flgDistrib)
                                                {
                                                    countWeight = 0;
                                                    countStud = 0;

                                                    for (int k = 0; k <= mdlData.colStudents.Count - 1; k++)
                                                    {
                                                        if (mdlData.colStudents[k].flgPlan)
                                                        {
                                                            //Если рассматриваемый преподаватель - руководитель студента
                                                            //И если студент на том же курсе, где и дисциплина
                                                            //И специальность студента должна соответствовать специальности нагрузки
                                                            if (mdlData.colStudents[k].Lect.Equals(mdlData.colLecturer[i])
                                                                & mdlData.colStudents[k].KursNum.Equals(collPlan[j].KursNum)
                                                                & mdlData.colStudents[k].Speciality.Equals(collPlan[j].Speciality))
                                                            {
                                                                countWeight += collPlan[j].Weight;
                                                                countStud++;
                                                            }
                                                        }
                                                    }

                                                    //
                                                    if (flgCombine)
                                                    {
                                                        mdlData.toDetectUniformInHoured(ref countWeight, collPlan[j], mdlData.colLecturer[i]);
                                                    }

                                                    //
                                                    if (countWeight > 0)
                                                    {
                                                        if (collPlan[j].Semestr.Equals(mdlData.colSemestr[1]))
                                                        {
                                                            if (collPlan[j].PreDiplomaPractice > 0 ||
                                                                collPlan[j].ProducingPractice > 0)
                                                            {
                                                                sumPreDipIplan += countWeight;
                                                            }

                                                            if (collPlan[j].TutorialPractice > 0)
                                                            {
                                                                sumTutPrIplan += countWeight;
                                                            }

                                                            if (collPlan[j].DiplomaPaper > 0)
                                                            {
                                                                sumDiplIplan += countWeight;
                                                            }

                                                            if (collPlan[j].Magistry > 0)
                                                            {
                                                                sumMagIplan += countWeight;
                                                            }
                                                        }

                                                        if (collPlan[j].Semestr.Equals(mdlData.colSemestr[2]))
                                                        {
                                                            if (collPlan[j].PreDiplomaPractice > 0 ||
                                                                collPlan[j].ProducingPractice > 0)
                                                            {
                                                                sumPreDipIIplan += countWeight;
                                                            }

                                                            if (collPlan[j].TutorialPractice > 0)
                                                            {
                                                                sumTutPrIIplan += countWeight;
                                                            }

                                                            if (collPlan[j].DiplomaPaper > 0)
                                                            {
                                                                sumDiplIIplan += countWeight;
                                                            }

                                                            if (collPlan[j].Magistry > 0)
                                                            {
                                                                sumMagIIplan += countWeight;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            accessFlg = flg;

                            if (accessFlg)
                            {
                                strSem = "=";
                                for (int k = 4; k <= 23; k++)
                                {
                                    strSem += mdlData.ExcelCellTranslator(6 + (curLect - 1), k) + "+";
                                }
                                strSem = strSem.Substring(0, strSem.Length - 1);

                                strSemPlan = (sumLecIplan + sumExamIplan + sumCredIplan +
                                    sumTutIplan + sumLabIplan + sumPracIplan + sumRefIplan +
                                    sumIndIplan + sumKRAPKIplan + sumKursPrIplan +
                                    sumDiplIplan + sumTutPrIplan + sumPreDipIplan + 
                                    sumGAKIplan + sumPostGrIplan + sumVisIplan + sumMagIplan).ToString();

                                if (strSemPlan.Equals("0"))
                                {
                                    strSemPlan = (sumLecIIplan + sumExamIIplan + sumCredIIplan +
                                        sumTutIIplan + sumLabIIplan + sumPracIIplan + sumRefIIplan +
                                        sumIndIIplan + sumKRAPKIIplan + sumKursPrIIplan +
                                        sumDiplIIplan + sumTutPrIIplan + sumPreDipIIplan +
                                        sumGAKIIplan + sumPostGrIIplan + sumVisIIplan + sumMagIIplan).ToString();
                                }

                                for (int k = 1; k <= 27; k++)
                                {
                                    //Выбираем диапазон
                                    cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(6 + (curLect - 1), k),
                                        mdlData.ExcelCellTranslator(6 + (curLect - 1), k));
                                    //Задаём границы
                                    cells.Borders.Weight = 2;
                                    //Горизонтальное выравнивание по центру в ячейках
                                    cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                    //Вертикальное выравнивание по центру в ячейках
                                    cells.VerticalAlignment = Excel.Constants.xlCenter;
                                    //Назначаем высоту строки
                                    cells.RowHeight = 30.00f;

                                    switch (k)
                                    {
                                        //Номер по порядку
                                        case 1:
                                            cells.Cells[1, 1] = curLect.ToString();
                                            break;
                                        //Фамилия, Имя, Отчество
                                        case 2:
                                            cells.Cells[1, 1] = mdlData.colLecturer[i].FIO;
                                            cells.WrapText = true;
                                            break;
                                        //Ставка
                                        case 3:
                                            cells.Cells[1, 1] = mdlData.colLecturer[i].Rate.ToString("0.00");
                                            //Горизонтальное выравнивание по центру в ячейках
                                            cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                            //Вертикальное выравнивание по центру в ячейках
                                            cells.VerticalAlignment = Excel.Constants.xlCenter;
                                            cells.WrapText = true;
                                            break;
                                        //Лекции
                                        case 4:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = sumLecI.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = sumLecII.ToString();
                                            }
                                            break;
                                        //Экзамены
                                        case 5:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = sumExamI.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = sumExamII.ToString();
                                            }
                                            break;
                                        //Зачёты
                                        case 6:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = sumCredI.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = sumCredII.ToString();
                                            }
                                            break;
                                        //ПК
                                        case 7:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = 0.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = 0.ToString();
                                            }
                                            break;
                                        //Консультации
                                        case 8:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = sumTutI.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = sumTutII.ToString();
                                            }
                                            break;
                                        //Практические занятия
                                        case 9:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = sumPracI.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = sumPracII.ToString();
                                            }
                                            break;
                                        //Домашние задания и рефераты
                                        case 10:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = sumRefI.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = sumRefII.ToString();
                                            }
                                            break;
                                        //Текущая аттестация
                                        case 11:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = 0.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = 0.ToString();
                                            }
                                            break;
                                        //Индивидуальные занятия
                                        case 12:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = sumIndI.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = sumIndII.ToString();
                                            }
                                            break;
                                        //Контрольные работы
                                        case 13:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = sumKRAPKI.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = sumKRAPKII.ToString();
                                            }
                                            break;
                                        //Курсовой проект, курсовая работа
                                        case 14:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = sumKursPrI.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = sumKursPrII.ToString();
                                            }
                                            break;
                                        //Дипломный проект
                                        case 15:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = sumDiplI.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = sumDiplII.ToString();
                                            }
                                            break;
                                        //Учебная практика
                                        case 16:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = sumTutPrI.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = sumTutPrII.ToString();
                                            }
                                            break;
                                        //Преддипломная и производственная практика
                                        case 17:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = sumPreDipI.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = sumPreDipII.ToString();
                                            }
                                            break;
                                        //ГЭК
                                        case 18:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = sumGAKI.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = sumGAKII.ToString();
                                            }
                                            break;
                                        //Приёмная комиссия
                                        case 19:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = 0.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = 0.ToString();
                                            }
                                            break;
                                        //Лабораторные работы
                                        case 20:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = sumLabI.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = sumLabII.ToString();
                                            }
                                            break;
                                        //Аспирантура
                                        case 21:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = sumPostGrI.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = sumPostGrII.ToString();
                                            }
                                            break;
                                        //Посещение занятий
                                        case 22:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = sumVisI.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = sumVisII.ToString();
                                            }
                                            break;
                                        //Другие виды занятий
                                        case 23:
                                            if (cmbSemestrList.SelectedIndex == 1)
                                            {
                                                cells.Cells[1, 1] = sumMagI.ToString();
                                            }
                                            else
                                            {
                                                cells.Cells[1, 1] = sumMagII.ToString();
                                            }
                                            break;
                                        //Плановая сумма по строке
                                        case 24:
                                            cells.Cells[1, 1] = strSemPlan;
                                            break;
                                        //Фактическая сумма по строке
                                        case 25:
                                            cells.Cells[1, 1] = strSem;
                                            break;
                                        //
                                        case 26:
                                            cells.Cells[1, 1] = "=" + mdlData.ExcelCellTranslator(6 + (curLect - 1), k - 1) + "-" +
                                                mdlData.ExcelCellTranslator(6 + (curLect - 1), k - 2);
                                            break;
                                        case 27:
                                            //Горизонтальное выравнивание по центру в ячейках
                                            cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                            //Вертикальное выравнивание по центру в ячейках
                                            cells.VerticalAlignment = Excel.Constants.xlCenter;
                                            break;
                                    }
                                }

                                strSumI += mdlData.ExcelCellTranslator(6 + (curLect - 1), 24) + "+";
                                strSumII += mdlData.ExcelCellTranslator(6 + (curLect - 1), 25) + "+";
                                strSumIII += mdlData.ExcelCellTranslator(6 + (curLect - 1), 26) + "+";

                                curLect += 1;
                            }
                        }
                    }

                    strSumI = strSumI.Substring(0, strSumI.Length - 1);
                    strSumII = strSumII.Substring(0, strSumII.Length - 1);
                    strSumIII = strSumIII.Substring(0, strSumIII.Length - 1);

                    //Итоговая сумма часов по плану
                    //Выбираем диапазон
                    cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (curLect) + 1, 24),
                        mdlData.ExcelCellTranslator(4 + (curLect) + 1, 24));
                    cells.Cells[1, 1] = strSumI;
                    //Задаём границы
                    cells.Borders.Weight = 2;

                    //Итоговая сумма часов по факту
                    //Выбираем диапазон
                    cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (curLect) + 1, 25),
                        mdlData.ExcelCellTranslator(4 + (curLect) + 1, 25));
                    cells.Cells[1, 1] = strSumII;
                    //Задаём границы
                    cells.Borders.Weight = 2;

                    //Итоговая сумма рассогласований
                    //Выбираем диапазон
                    cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (curLect) + 1, 26),
                        mdlData.ExcelCellTranslator(4 + (curLect) + 1, 26));
                    cells.Cells[1, 1] = strSumIII;
                    //Задаём границы
                    cells.Borders.Weight = 2;

                    //Надпись итого
                    //Выбираем диапазон
                    cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (curLect) + 1, 23),
                        mdlData.ExcelCellTranslator(4 + (curLect) + 1, 24));
                    cells.Cells[1, 1] = "итого:";

                    //Подпись заведующего кафедрой
                    //Выбираем диапазон для надписи "Заведующий кафедрой"
                    cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (curLect) + 4, 21),
                        mdlData.ExcelCellTranslator(4 + (curLect) + 4, 21));
                    cells.Cells[1, 1] = "Заведующий кафедрой";

                    //Выбираем диапазон для расшифровки подписи
                    cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (curLect) + 4, 26),
                        mdlData.ExcelCellTranslator(4 + (curLect) + 4, 26));
                    cells.Cells[1, 1] = "/ Л.А. Баранов /";

                    //Дата формирования ведомости
                    //Выбираем диапазон
                    cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (curLect) + 4, 4),
                        mdlData.ExcelCellTranslator(4 + (curLect) + 4, 4));
                    cells.Cells[1, 1] = "\"" + DateTime.Now.Day.ToString() + "\" " + mdlData.getMonthStringRP(DateTime.Now.Month) + " " + DateTime.Now.Year.ToString() + " г.";

                    //-----------Формируем таблицу даными

                    ObjExcel.UserControl = true;

                    //Сохраняем файл с уникальным именем
                    ObjWorkBook.SaveCopyAs(Application.StartupPath + @"\Ведомость фактическая" +
                        fileNameAdd +
                        mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum + " " +
                        DateTime.Now.Date.ToString("yyyyMMdd") + " " +
                        DateTime.Now.TimeOfDay.ToString("hhmmss") + ".xlsx");
                    
                    //Закрываем книгу Excel
                    ObjWorkBook.Close(false, "", Missing.Value);

                    //Закрываем приложение Excel
                    ObjExcel.Quit();
                }
                catch
                {
                    MessageBox.Show("Возможно на этом компьютере присутствует проблема совместимости с MS Exсel." +
                                    " Попробуйте установить версию 2007 и выше.", "Ошибка!");
                }
            }
            else
            {
                MessageBox.Show("Семестр не выбран", "Внимание!");
            }
        }

        private void UpravlenieFixedSubj()
        {
            int curPosition;

            clsLecturer_Subj LS;
            IList<clsLecturer_Subj> colLS = null;

            IList<clsDistribution> coll = null;

            //bool accessFlg = false;

            if (optMain.Checked || optMainDop.Checked)
            {
                coll = mdlData.colDistribution;
            }
            else
            {
                if (optHoured.Checked)
                {
                    coll = mdlData.colHouredDistribution;
                }
                else
                {
                    if (optCombine.Checked || optCombineDop.Checked)
                    {
                        coll = mdlData.colCombineDistribution;
                    }
                }
            }

            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;

            try
            {
                //Создаём новое Excel приложение
                Excel.Application ObjExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook;
                Excel.Worksheet ObjWorkSheet;

                ObjExcel.Visible = true;

                //Книга
                ObjWorkBook = ObjExcel.Workbooks.Add(Missing.Value);
                //Таблица
                ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

                //Альбомная ориентация страницы
                ObjWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                //Книжная ориентация страницы
                //ObjWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                //В высоту разместить на одной странице
                ObjWorkSheet.PageSetup.FitToPagesTall = 1;
                //В ширину разместить на одной странице
                ObjWorkSheet.PageSetup.FitToPagesWide = 1;

                //-----------Формируем заголовок таблицы

                //Задаём диапазон для ячеек, подлежащих форматированию
                //1-я строка, с А по AA
                var cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(1, 1),
                    mdlData.ExcelCellTranslator(1, 3));
                //Выделяем ячейки диапазона
                cells.Select();
                //Объединяем ячейки диапазона
                cells.Merge(true);
                //Выравнивание в объединённой ячейке по центру
                cells.HorizontalAlignment = Excel.Constants.xlCenter;

                //Записываем текст в объединённую ячейку
                //(считается по первой А1)

                //Заголовок таблицы
                cells.Cells[1, 1] = "Сведения о преподаваемых дисциплинах в институте ИТТСУ ";

                //Задаём диапазон для ячеек, подлежащих форматированию
                //2-я строка, с А по AA
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(2, 1),
                    mdlData.ExcelCellTranslator(2, 3));
                //Выделяем ячейки диапазона
                cells.Select();
                //Объединяем ячейки диапазона
                cells.Merge(true);
                //Выравнивание в объединённой ячейке по центру
                cells.HorizontalAlignment = Excel.Constants.xlCenter;
                //Записываем текст в объединённую ячейку
                //(считается по первой А2)
                cells.Cells[1, 1] = "Кафедра \"Управление и защита информации\" " +
                    mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear +
                    " учебный год";

                //-----------Формируем заголовок таблицы

                //-----------Формируем шапку таблицы

                for (int i = 1; i <= 3; i++)
                {
                    //Выбираем диапазон
                    cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(3, i),
                        mdlData.ExcelCellTranslator(3, i));
                    //Выделяем ячейки диапазона
                    cells.Select();
                    //Объединяем ячейки диапазона
                    cells.Merge();
                    //Горизонтальное выравнивание по центру в ячейках
                    cells.HorizontalAlignment = Excel.Constants.xlCenter;
                    //Вертикальное выравнивание по центру в ячейках
                    cells.VerticalAlignment = Excel.Constants.xlCenter;
                    //Задаём границы
                    cells.Borders.Weight = 2;

                    switch (i)
                    {
                        case 1:
                            cells.Cells[1, 1] = "ФИО (по алфавиту)";
                            cells.ColumnWidth = 36.00f;
                            break;
                        case 2:
                            cells.Cells[1, 1] = "Должность, степень";
                            cells.ColumnWidth = 34.00f;
                            break;
                        case 3:
                            cells.Cells[1, 1] = "Преподаваемые дисциплины";
                            cells.ColumnWidth = 120.00f;
                            break;
                    }
                }

                //-----------Формируем шапку таблицы

                //-----------Заполняем таблицу данными

                curPosition = 1;

                for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
                {
                    colLS = new List<clsLecturer_Subj>();
                    
                    //Просматриваем нагрузку за первый семестр
                    for (int j = 0; j <= coll.Count - 1; j++)
                    {
                        if (!(coll[j].Lecturer == null))
                        {
                            if (coll[j].Lecturer.Equals(mdlData.colLecturer[i]))
                            {
                                if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    if (coll[j].Lecture > 0)
                                    {
                                        LS = new clsLecturer_Subj();
                                        LS.Subject = coll[j].Subject;
                                        LS.Sem = coll[j].Semestr;
                                        LS.Spec = coll[j].Speciality;
                                        LS.Type = "лекции";
                                        colLS.Add(LS);
                                    }

                                    if (coll[j].LabWork > 0)
                                    {
                                        LS = new clsLecturer_Subj();
                                        LS.Subject = coll[j].Subject;
                                        LS.Sem = coll[j].Semestr;
                                        LS.Spec = coll[j].Speciality;
                                        LS.Type = "лабораторные работы";
                                        colLS.Add(LS);
                                    }

                                    if (coll[j].Practice > 0)
                                    {
                                        LS = new clsLecturer_Subj();
                                        LS.Subject = coll[j].Subject;
                                        LS.Sem = coll[j].Semestr;
                                        LS.Spec = coll[j].Speciality;
                                        LS.Type = "практические занятия";
                                        colLS.Add(LS);
                                    }

                                    if (coll[j].KursProject > 0)
                                    {
                                        LS = new clsLecturer_Subj();
                                        LS.Subject = coll[j].Subject;
                                        LS.Sem = coll[j].Semestr;
                                        LS.Spec = coll[j].Speciality;
                                        LS.Type = "курсовой проект (курсовая работа)";
                                        colLS.Add(LS);
                                    }

                                    if (coll[j].TutorialPractice > 0)
                                    {
                                        LS = new clsLecturer_Subj();
                                        LS.Subject = coll[j].Subject;
                                        LS.Sem = coll[j].Semestr;
                                        LS.Spec = coll[j].Speciality;
                                        LS.Type = "учебная практика";
                                        colLS.Add(LS);
                                    }

                                    if (coll[j].ProducingPractice > 0)
                                    {
                                        LS = new clsLecturer_Subj();
                                        LS.Subject = coll[j].Subject;
                                        LS.Sem = coll[j].Semestr;
                                        LS.Spec = coll[j].Speciality;
                                        LS.Type = "производственная практика";
                                        colLS.Add(LS);
                                    }
                                }
                            }
                        }
                    }

                    //Просматриваем нагрузку за второй семестр
                    for (int j = 0; j <= coll.Count - 1; j++)
                    {
                        if (!(coll[j].Lecturer == null))
                        {
                            if (coll[j].Lecturer.Equals(mdlData.colLecturer[i]))
                            {
                                if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    if (coll[j].Lecture > 0)
                                    {
                                        LS = new clsLecturer_Subj();
                                        LS.Subject = coll[j].Subject;
                                        LS.Sem = coll[j].Semestr;
                                        LS.Spec = coll[j].Speciality;
                                        LS.Type = "лекции";
                                        colLS.Add(LS);
                                    }

                                    if (coll[j].LabWork > 0)
                                    {
                                        LS = new clsLecturer_Subj();
                                        LS.Subject = coll[j].Subject;
                                        LS.Sem = coll[j].Semestr;
                                        LS.Spec = coll[j].Speciality;
                                        LS.Type = "лабораторные работы";
                                        colLS.Add(LS);
                                    }

                                    if (coll[j].Practice > 0)
                                    {
                                        LS = new clsLecturer_Subj();
                                        LS.Subject = coll[j].Subject;
                                        LS.Sem = coll[j].Semestr;
                                        LS.Spec = coll[j].Speciality;
                                        LS.Type = "практические занятия";
                                        colLS.Add(LS);
                                    }

                                    if (coll[j].KursProject > 0)
                                    {
                                        LS = new clsLecturer_Subj();
                                        LS.Subject = coll[j].Subject;
                                        LS.Sem = coll[j].Semestr;
                                        LS.Spec = coll[j].Speciality;
                                        LS.Type = "курсовой проект (курсовая работа)";
                                        colLS.Add(LS);
                                    }

                                    if (coll[j].TutorialPractice > 0)
                                    {
                                        LS = new clsLecturer_Subj();
                                        LS.Subject = coll[j].Subject;
                                        LS.Sem = coll[j].Semestr;
                                        LS.Spec = coll[j].Speciality;
                                        LS.Type = "учебная практика";
                                        colLS.Add(LS);
                                    }

                                    if (coll[j].ProducingPractice > 0)
                                    {
                                        LS = new clsLecturer_Subj();
                                        LS.Subject = coll[j].Subject;
                                        LS.Sem = coll[j].Semestr;
                                        LS.Spec = coll[j].Speciality;
                                        LS.Type = "производственная практика";
                                        colLS.Add(LS);
                                    }
                                }
                            }
                        }
                    }

                    if (colLS.Count > 0)
                    {
                        for (int l = 0; l <= colLS.Count - 1; l++)
                        {
                            for (int k = 1; k <= 3; k++)
                            {
                                //Выбираем диапазон
                                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (curPosition - 1), k),
                                    mdlData.ExcelCellTranslator(4 + (curPosition - 1), k));
                                //Задаём границы
                                cells.Borders.Weight = 2;
                                //Горизонтальное выравнивание по центру в ячейках
                                cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                //Вертикальное выравнивание по центру в ячейках
                                cells.VerticalAlignment = Excel.Constants.xlCenter;
                                //Назначаем высоту строки
                                cells.RowHeight = 30.00f;

                                switch (k)
                                {
                                    //Фамилия, Имя, Отчество
                                    case 1:
                                        cells.Cells[1, 1] = mdlData.colLecturer[i].FIO;
                                        cells.WrapText = true;
                                        break;
                                    //Должность, степень
                                    case 2:
                                        cells.Cells[1, 1] = mdlData.colLecturer[i].Duty.Duty + ", " +
                                                            mdlData.colLecturer[i].Degree.Short;
                                        cells.WrapText = true;
                                        break;
                                    //Преподаваемая дисцилина
                                    case 3:
                                        cells.Cells[1, 1] = colLS[l].Subject.Subject + " {" + colLS[l].Spec.ShortInstitute + "}, [" + colLS[l].Sem.SemNum + "], (" + colLS[l].Type + ")";
                                        cells.WrapText = true;
                                        break;
                                }
                            }

                            curPosition += 1;
                        }
                    }
                }

                //-----------Формируем таблицу даными

                ObjExcel.UserControl = true;

                ObjWorkBook.SaveCopyAs(Application.StartupPath + @"\Закреплённые дисциплины " + 
                    mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum + " " +
                    DateTime.Now.Date.ToString("yyyyMMdd") + " " +
                    DateTime.Now.TimeOfDay.ToString("hhmmss") + ".xlsx");
                ObjWorkBook.Close(false, "", Missing.Value);

                ObjExcel.Quit();
            }
            catch
            {
                MessageBox.Show("Возможно на этом компьютере присутствует проблема совместимости с MS Exсel." +
                                " Попробуйте установить версию 2007 и выше.", "Ошибка!");
            }

        }

        /// <summary>
        /// Функция сортировки нагрузки по иерархии: преподаватель, курс, дисциплина
        /// </summary>
        /// <param name="coll"></param>
        /// <returns></returns>
        private IList<clsDistribution> SortLoad(IList<clsDistribution> coll)
        {
            int i, j, k;
            //Перебираем учебные семестры
            for (i = 0; i <= mdlData.colSemestr.Count - 1; i++)
            {
                //Перебираем курсы
                for (j = -1; j <= mdlData.colKursNum.Count - 1; j++)
                {
                    //Перебираем нагрузку
                    for (k = 0; k <= mdlData.colDistribution.Count - 1; k++)
                    {
                        if (!mdlData.colDistribution[k].flgExclude)
                        {
                            //Если семестр не указан
                            if (i == 0)
                            {
                                //Если для нагрузки семестр тоже не указан
                                if (mdlData.colDistribution[k].Semestr == null)
                                {
                                    //Планируем строку в нагрузку
                                    coll.Add(mdlData.colDistribution[k]);
                                }
                            }
                            //Если семестр указан
                            else
                            {
                                //Если курс не указан
                                if (j == -1)
                                {
                                    if (mdlData.colDistribution[k].Semestr != null)
                                    {
                                        if (mdlData.colDistribution[k].KursNum == null &
                                            mdlData.colDistribution[k].Semestr.SemNum.Equals(mdlData.colSemestr[i].SemNum))
                                        {
                                            coll.Add(mdlData.colDistribution[k]);
                                        }
                                    }
                                }
                                //Если курс указан
                                else
                                {
                                    if (mdlData.colDistribution[k].Semestr != null &
                                        mdlData.colDistribution[k].KursNum != null)
                                    {
                                        if (mdlData.colDistribution[k].KursNum.Kurs.Equals(mdlData.colKursNum[j].Kurs) &
                                            mdlData.colDistribution[k].Semestr.SemNum.Equals(mdlData.colSemestr[i].SemNum))
                                        {
                                            coll.Add(mdlData.colDistribution[k]);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return coll;
        }

        /// <summary>
        /// Подготовка сетки для передачи распределения Заведующему кафедрой
        /// </summary>
        private void PrepareGridForBoss(IList<clsDistribution> coll, bool flgBoth, ref double SumRate, ref double SumRateOld,
                                        ref double sumNotUnLoadRate, ref int sumUnLoad)
        {
            int i, j, k;
            bool flgHaveLoad;
            //-----------------------------------------------------------------
            //---------------Фрагмент обозначения рабочей области--------------
            //-----------------------------------------------------------------

            //Задаём количество столбцов
            //оно остаётся неизменным
            //0. Преподаватель
            //1. Семестр
            //2. Курс
            //3. Специальность
            //4. Название дисциплины
            //5. (пусто)
            //6. Сумма часов
            //7. Баланс
            //8. Ставка
            //9. Догрузка
            //10. Разгрузка

            for (i = 0; i <= 10; i++)
            {
                dgScheduleManagement.Columns.Add("", "");
            }

            //Сразу добавляем две строки (под общую информацию)
            dgScheduleManagement.Rows.Add();
            //и под пробел
            dgScheduleManagement.Rows.Add();

            //Просматриваем каждого преподавателя
            for (i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                SumRate += mdlData.colLecturer[i].Rate;
                SumRateOld += mdlData.colLecturer[i].OldRate;

                //Считаем догрузку для преподавателей без разгрузки
                if (mdlData.colLecturer[i].UnLoad == 0)
                {
                    sumNotUnLoadRate += mdlData.colLecturer[i].Rate;
                }
                else
                {
                    sumUnLoad += mdlData.colLecturer[i].UnLoad;
                }

                //Считаем, по умолчанию, что преподаватель не нагружен
                flgHaveLoad = false;
                //Просматриваем нагрузку
                for (j = 0; j <= coll.Count - 1; j++)
                {
                    if (!coll[j].flgExclude)
                    {
                        //Случай для одного семестра
                        if (!flgBoth)
                        {
                            //Если дисциплина относится к выбранному из списка семестру
                            if (coll[j].Semestr.Equals(mdlData.colSemestr[cmbSemestrList.SelectedIndex]))
                            {
                                //Если обычная строка нагрузки
                                if (!coll[j].flgDistrib)
                                {
                                    //Если рассматриваемый преподаватель совпадает с
                                    //преподавателем, указанным в нагрузке
                                    if (mdlData.colLecturer[i].Equals(coll[j].Lecturer))
                                    {
                                        //Если дисциплина предполагает хоть какие-то часы, то
                                        //выделяем под неё строчку текста
                                        if (mdlData.NonZeroDistributionOR(coll[j]))
                                        {
                                            //Если преподаватель что-либо из этого ведёт
                                            //добавляем строку
                                            dgScheduleManagement.Rows.Add();
                                            //Получается, что преподаватель нагружен
                                            flgHaveLoad = true;
                                        }
                                    }
                                }
                                //Если равномерно распределяемая нагрузка
                                else
                                {
                                    for (k = 0; k <= mdlData.colStudents.Count - 1; k++)
                                    {
                                        if (mdlData.colStudents[k].flgPlan)
                                        {
                                            //Если рассматриваемый преподаватель - руководитель студента
                                            //И если студент на том же курсе, где и дисциплина
                                            //И специальность студента должна соответствовать специальности нагрузки
                                            if (mdlData.colStudents[k].Lect.Equals(mdlData.colLecturer[i])
                                                & mdlData.colStudents[k].KursNum.Equals(coll[j].KursNum)
                                                & mdlData.colStudents[k].Speciality.Equals(coll[j].Speciality))
                                            {
                                                //Если преподаватель что-либо из этого ведёт
                                                //добавляем строку
                                                dgScheduleManagement.Rows.Add();
                                                //Получается, что преподаватель нагружен
                                                flgHaveLoad = true;
                                                //Одной строки достаточно, прерываем цикл
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        //Случай для двух семестров
                        else
                        {
                            //Если обычная строка нагрузки
                            if (!coll[j].flgDistrib)
                            {
                                //Если рассматриваемый преподаватель совпадает с
                                //преподавателем, указанным в нагрузке
                                if (mdlData.colLecturer[i].Equals(coll[j].Lecturer))
                                {
                                    //Если дисциплина предполагает хоть какие-то часы, то
                                    //выделяем под неё строчку текста
                                    if (mdlData.NonZeroDistributionOR(coll[j]))
                                    {
                                        //Если преподаватель что-либо из этого ведёт
                                        //добавляем строку
                                        dgScheduleManagement.Rows.Add();
                                        //Получается, что преподаватель нагружен
                                        flgHaveLoad = true;
                                    }
                                }
                            }
                            //Если равномерно распределяемая нагрузка
                            else
                            {
                                for (k = 0; k <= mdlData.colStudents.Count - 1; k++)
                                {
                                    if (mdlData.colStudents[k].flgPlan)
                                    {
                                        //Если рассматриваемый преподаватель - руководитель студента
                                        //И если студент на том же курсе, где и дисциплина
                                        //И специальность студента должна соответствовать специальности нагрузки
                                        if (mdlData.colStudents[k].Lect.Equals(mdlData.colLecturer[i])
                                            & mdlData.colStudents[k].KursNum.Equals(coll[j].KursNum)
                                            & mdlData.colStudents[k].Speciality.Equals(coll[j].Speciality))
                                        {
                                            //Если преподаватель что-либо из этого ведёт
                                            //добавляем строку
                                            dgScheduleManagement.Rows.Add();
                                            //Получается, что преподаватель нагружен
                                            flgHaveLoad = true;
                                            //Одной строки достаточно, прерываем цикл
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                //Только если у преподавателя есть нагрузка
                if (flgHaveLoad)
                {
                    //Разделяем преподавателей пустой строчкой
                    dgScheduleManagement.Rows.Add();
                    dgScheduleManagement.Rows.Add();
                }
            }

            mdlData.LoadInc = Convert.ToInt32(-sumUnLoad / sumNotUnLoadRate);
        }

        /// <summary>
        /// Подготовка сетки для выдачи показателей нагрузки
        /// </summary>
        private void PrepareGridForLoad(IList<clsDistribution> coll, bool flgBoth, ref double SumRate,
                                        ref double sumNotUnLoadRate, ref int sumUnLoad)
        {
            int i, j, k;
            bool flgHaveLoad;
            //-----------------------------------------------------------------
            //---------------Фрагмент обозначения рабочей области--------------
            //-----------------------------------------------------------------

            //Задаём количество столбцов
            //оно остаётся неизменным
            //0. Преподаватель
            //1. Ставка
            //2. Норма часов
            //3. Догрузка
            //4. Разгрузка
            //5. Распределено
            //6. Отклонение от нормы
            //7. Почасовая в первом семестре
            //8. Почасовая во втором семестре
            //9. Суммарная почасовая
            //10. Компенсация отклонения
            for (i = 0; i <= 10; i++)
            {
                dgScheduleManagement.Columns.Add("", "");
            }

            //Сразу добавляем две строки (под общую информацию)
            dgScheduleManagement.Rows.Add();
            //и под пробел
            dgScheduleManagement.Rows.Add();

            //Просматриваем каждого преподавателя
            for (i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                SumRate += mdlData.colLecturer[i].Rate;

                //Считаем догрузку для преподавателей без разгрузки
                if (mdlData.colLecturer[i].UnLoad == 0)
                {
                    sumNotUnLoadRate += mdlData.colLecturer[i].Rate;
                }
                else
                {
                    sumUnLoad += mdlData.colLecturer[i].UnLoad;
                }

                //Считаем, по умолчанию, что преподаватель не нагружен
                flgHaveLoad = false;
                //Просматриваем нагрузку
                for (j = 0; j <= coll.Count - 1; j++)
                {
                    if (!coll[j].flgExclude)
                    {
                        //Если обычная строка нагрузки
                        if (!coll[j].flgDistrib)
                        {
                            //Если рассматриваемый преподаватель совпадает с
                            //преподавателем, указанным в нагрузке
                            if (mdlData.colLecturer[i].Equals(coll[j].Lecturer))
                            {
                                //Если дисциплина предполагает хоть какие-то часы, то
                                //выделяем под неё строчку текста
                                if (mdlData.NonZeroDistributionOR(coll[j]))
                                {
                                    //Получается, что преподаватель нагружен
                                    flgHaveLoad = true;
                                    break;
                                }
                            }
                        }
                        //Если равномерно распределяемая нагрузка
                        else
                        {
                            for (k = 0; k <= mdlData.colStudents.Count - 1; k++)
                            {
                                if (mdlData.colStudents[k].flgPlan)
                                {
                                    //Если рассматриваемый преподаватель - руководитель студента
                                    //И если студент на том же курсе, где и дисциплина
                                    //И специальность студента должна соответствовать специальности нагрузки
                                    if (mdlData.colStudents[k].Lect.Equals(mdlData.colLecturer[i])
                                        & mdlData.colStudents[k].KursNum.Equals(coll[j].KursNum)
                                        & mdlData.colStudents[k].Speciality.Equals(coll[j].Speciality))
                                    {
                                        //Получается, что преподаватель нагружен
                                        flgHaveLoad = true;
                                        //Одной строки достаточно, прерываем цикл
                                        break;
                                    }
                                }
                            }
                        }                       
                    }
                }

                //Только если у преподавателя есть нагрузка
                if (flgHaveLoad)
                {
                    //Переходим к следующему преподавателю
                    dgScheduleManagement.Rows.Add();
                }
            }

            //Выделяем строку под суммарную нагрузку
            dgScheduleManagement.Rows.Add();
            //Выделяем строку под суммарную нагрузку без учёта нулевых ставок
            dgScheduleManagement.Rows.Add();
            //Выделяем ещё строку под дополнительные нужды
            dgScheduleManagement.Rows.Add();

            mdlData.LoadInc = Convert.ToInt32(-sumUnLoad / sumNotUnLoadRate);
        }

        /// <summary>
        /// Подготовка сетки для выдачи равномерно распределяемой нагрузки
        /// </summary>
        private void PrepareGridForUniformLoad(IList<clsDistribution> coll)
        {
            int i, j, k;
            bool flgHaveUniformLoad;
            bool flgHaveStudent;
            //-----------------------------------------------------------------
            //---------------Фрагмент обозначения рабочей области--------------
            //-----------------------------------------------------------------

            //Задаём количество столбцов
            //оно остаётся неизменным
            //0. Преподаватель / наименование
            //1. Семестр
            //2. Группа
            //3. Курс
            //4. Часы

            for (i = 0; i <= 4; i++)
            {
                dgScheduleManagement.Columns.Add("", "");

                switch (i)
                {
                    case 0:
                    {
                        dgScheduleManagement.Columns[i].Width = 400;
                        break;
                    }
                    case 1:
                    {
                        dgScheduleManagement.Columns[i].Width = 50;
                        break;
                    }
                    case 2:
                    {  
                        dgScheduleManagement.Columns[i].Width = 50;
                        break;
                    }
                    case 3:
                    {
                        dgScheduleManagement.Columns[i].Width = 50;
                        break;
                    }
                    case 4:
                    {
                        dgScheduleManagement.Columns[i].Width = 50;
                        break;
                    }
                }
            }

            //Добавляем строку под шапку
            dgScheduleManagement.Rows.Add();

            //Просматриваем каждого преподавателя
            for (i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                //Считаем, по умолчанию, что преподаватель не нагружен
                //равномерно распределяемыми дисциплинами
                flgHaveUniformLoad = false;
                //Просматриваем нагрузку
                for (j = 0; j <= coll.Count - 1; j++)
                {
                    //Если элемент нагрузки не исключён из расчёта
                    if (!coll[j].flgExclude)
                    {
                        //Если равномерно распределяемая нагрузка
                        if (coll[j].flgDistrib)
                        {
                            flgHaveStudent = false;
                            //Просматриваем студентов
                            for (k = 0; k <= mdlData.colStudents.Count - 1; k++)
                            {
                                //Если студент учитывается в расчёте нагрузки
                                if (mdlData.colStudents[k].flgPlan)
                                {
                                    //Если рассматриваемый преподаватель - руководитель студента
                                    //И если студент на том же курсе, где и дисциплина
                                    //И специальность студента должна соответствовать специальности нагрузки
                                    if (mdlData.colStudents[k].Lect.Equals(mdlData.colLecturer[i])
                                        & mdlData.colStudents[k].KursNum.Equals(coll[j].KursNum)
                                        & mdlData.colStudents[k].Speciality.Equals(coll[j].Speciality))
                                    {
                                        //
                                        flgHaveStudent = true;
                                        //Одного совпавшего студента достаточно, прерываем цикл
                                        break;
                                    }
                                }
                            }

                            if (flgHaveStudent)
                            {
                                if (!flgHaveUniformLoad)
                                {
                                    //Получается, что у преподавателя есть нагрузка
                                    flgHaveUniformLoad = true;
                                    //Добавляем строку под Ф.И.О. преподавателя
                                    dgScheduleManagement.Rows.Add();
                                }

                                //Добавляем строку под вид нагрузки
                                dgScheduleManagement.Rows.Add();
                            }
                        }
                    }
                }
                //Завершаем просмотр нагрузки

                if (flgHaveUniformLoad)
                {
                    //Добавляем строку под разрыв между преподавателями
                    dgScheduleManagement.Rows.Add();
                }
            }
        }

        //Заполнить сетку нагрузки для заведующего кафедрой
        private void FillGridForBoss(IList<clsDistribution> coll, bool flgBoth, ref double SumRate, ref double SumRateOld,
                                        ref double sumNotUnLoadRate, ref int sumUnLoad)
        {
            //Берём первую строку (с нуля)
            int curRow = 0, sumHours, sumCurrent, sumStud, sumAll, AvgLoad = mdlData.AverageLoad;
            int i, j, k;
            bool flgHaveLoad;
            double UnLoadDopFact = 0, PlanHours = 0;

            dgScheduleManagement[0, curRow].Value = "Сред.нагр.";
            dgScheduleManagement[1, curRow].Value = mdlData.AverageLoad.ToString("0.00");

            dgScheduleManagement[2, curRow].Value = "Сумм.ставка";
            dgScheduleManagement[3, curRow].Value = SumRate.ToString("0.00") + " (" + SumRateOld.ToString("0.00") + ")";

            dgScheduleManagement[4, curRow].Value = "Сумм.разгр.";
            dgScheduleManagement[5, curRow].Value = (-sumUnLoad).ToString("0.00");

            dgScheduleManagement[6, curRow].Value = "Сумм.став.б/р";
            dgScheduleManagement[7, curRow].Value = (sumNotUnLoadRate).ToString("0.00");

            //Берём третью строку (с нуля)
            curRow = 2;
            //Формируем шапку выписки
            dgScheduleManagement[0, curRow].Value = "Преподаватель";
            dgScheduleManagement[1, curRow].Value = "Сем.";
            dgScheduleManagement[2, curRow].Value = "Курс";
            dgScheduleManagement[3, curRow].Value = "Специальность";
            dgScheduleManagement[4, curRow].Value = "Название дисциплины";
            dgScheduleManagement[6, curRow].Value = "Сумма часов";
            dgScheduleManagement[7, curRow].Value = "Баланс";
            dgScheduleManagement[8, curRow].Value = "Ставка";
            dgScheduleManagement[9, curRow].Value = "Догрузка";
            dgScheduleManagement[10, curRow].Value = "Разгрузка";
            //Идём на следующую строку
            curRow += 1;

            //Просматриваем каждого преподавателя
            sumAll = 0;
            for (i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                sumHours = 0;
                //По умолчанию считаем, что у преподавателя нет нагрузки
                flgHaveLoad = false;
                //Просматриваем нагрузку
                for (j = 0; j <= coll.Count - 1; j++)
                {
                    if (!coll[i].flgExclude)
                    {
                        sumCurrent = 0;
                        //Если работаем с одним из выбранных семестров
                        if (!flgBoth)
                        {
                            //Если дисциплина относится к выбранному семестру
                            if (coll[j].Semestr.Equals(mdlData.colSemestr[cmbSemestrList.SelectedIndex]))
                            {
                                //Если стандартная строка нагрузки
                                if (!coll[j].flgDistrib)
                                {
                                    //Если рассматриваемый преподаватель совпадает с
                                    //преподавателем, указанным в нагрузке
                                    if (mdlData.colLecturer[i].Equals(coll[j].Lecturer))
                                    {
                                        //Если дисциплина предполагает хоть какие-то часы, то
                                        //выделяем под неё строчку текста
                                        if (mdlData.NonZeroDistributionOR(coll[j]))
                                        {
                                            if (!flgHaveLoad)
                                            {
                                                //Печатаем фамилию, имя, отчество преподавателя
                                                dgScheduleManagement[0, curRow].Value = mdlData.colLecturer[i].FIO;
                                                //Получается, что у преподавателя есть нагрузка
                                                flgHaveLoad = true;
                                            }

                                            sumCurrent += mdlData.toSumDistributionComponents(coll[j]);

                                            //Печатаем код семестра, учитывая смещение
                                            dgScheduleManagement[1, curRow].Value = coll[j].Semestr.Code - 1;
                                            //Печатаем курс
                                            if (!(coll[j].KursNum == null))
                                            {
                                                dgScheduleManagement[2, curRow].Value = coll[j].KursNum.Kurs;
                                            }
                                            else
                                            {
                                                dgScheduleManagement[2, curRow].Value = "-";
                                            }
                                            //Печатаем специальность
                                            if (!(coll[j].Speciality == null))
                                            {
                                                dgScheduleManagement[3, curRow].Value = coll[j].Speciality.ShortUpravlenie
                                                        + " (" + coll[j].Speciality.ShortInstitute + ")";
                                            }
                                            else
                                            {
                                                dgScheduleManagement[3, curRow].Value = "-";
                                            }
                                            //Печатаем название дисциплины
                                            dgScheduleManagement[4, curRow].Value = coll[j].Subject.Subject + " " + mdlData.commentLoad(coll[j]);
                                            //Практические занятия в часах
                                            dgScheduleManagement[6, curRow].Value = sumCurrent;
                                            //Суммируем нагрузку преподавателя с учётом пройденных
                                            //позиций
                                            sumHours += sumCurrent;
                                            //Переходим к следующей строке
                                            curRow += 1;
                                        }
                                    }
                                }
                                //Если равномерно распределяемая нагрузка
                                else
                                {
                                    sumCurrent = 0;
                                    sumStud = 0;
                                    for (k = 0; k <= mdlData.colStudents.Count - 1; k++)
                                    {
                                        if (mdlData.colStudents[k].flgPlan)
                                        {
                                            //Если рассматриваемый преподаватель - руководитель студента
                                            //И если студент на том же курсе, где и дисциплина
                                            //И специальность студента должна соответствовать специальности нагрузки
                                            if (mdlData.colStudents[k].Lect.Equals(mdlData.colLecturer[i])
                                                & mdlData.colStudents[k].KursNum.Equals(coll[j].KursNum)
                                                & mdlData.colStudents[k].Speciality.Equals(coll[j].Speciality))
                                            {
                                                sumCurrent += coll[j].Weight;
                                                sumStud++;
                                            }
                                        }
                                    }

                                    //Если сумма изменилась
                                    if (sumCurrent > 0)
                                    {
                                        if (!flgHaveLoad)
                                        {
                                            //Печатаем фамилию, имя, отчество преподавателя
                                            dgScheduleManagement[0, curRow].Value = mdlData.colLecturer[i].FIO;
                                            //Получается, что у преподавателя есть нагрузка
                                            flgHaveLoad = true;
                                        }

                                        //Печатаем код семестра, учитывая смещение
                                        dgScheduleManagement[1, curRow].Value = coll[j].Semestr.Code - 1;
                                        //Печатаем курс
                                        if (!(coll[j].KursNum == null))
                                        {
                                            dgScheduleManagement[2, curRow].Value = coll[j].KursNum.Kurs;
                                        }
                                        else
                                        {
                                            dgScheduleManagement[2, curRow].Value = "-";
                                        }
                                        //Печатаем специальность
                                        if (!(coll[j].Speciality == null))
                                        {
                                            dgScheduleManagement[3, curRow].Value = coll[j].Speciality.ShortUpravlenie
                                                    + " (" + coll[j].Speciality.ShortInstitute + ")";
                                        }
                                        else
                                        {
                                            dgScheduleManagement[3, curRow].Value = "-";
                                        }

                                        //Печатаем название дисциплины
                                        dgScheduleManagement[4, curRow].Value = coll[j].Subject.Subject + " " + mdlData.commentLoad(coll[j]);
                                        //Практические занятия в часах
                                        dgScheduleManagement[6, curRow].Value = sumCurrent;
                                        //Суммируем нагрузку преподавателя с учётом пройденных
                                        //позиций
                                        sumHours += sumCurrent;

                                        //Переходим к следующей строке
                                        curRow += 1;
                                    }
                                }
                            }
                        }
                        //Если сведения выдаются на оба семестра
                        else
                        {
                            //Если стандартная строка нагрузки
                            if (!coll[j].flgDistrib)
                            {
                                //Если рассматриваемый преподаватель совпадает с
                                //преподавателем, указанным в нагрузке
                                if (mdlData.colLecturer[i].Equals(coll[j].Lecturer))
                                {
                                    //Если дисциплина предполагает хоть какие-то часы, то
                                    //выделяем под неё строчку текста
                                    if (mdlData.NonZeroDistributionOR(coll[j]))
                                    {
                                        if (!flgHaveLoad)
                                        {
                                            //Печатаем фамилию, имя, отчество преподавателя
                                            dgScheduleManagement[0, curRow].Value = mdlData.colLecturer[i].FIO;
                                            //Получается, что у преподавателя есть нагрузка
                                            flgHaveLoad = true;
                                        }

                                        sumCurrent += mdlData.toSumDistributionComponents(coll[j]);

                                        //Печатаем код семестра, учитывая смещение
                                        if (!(coll[j].Semestr == null))
                                        {
                                            dgScheduleManagement[1, curRow].Value = coll[j].Semestr.Code - 1;
                                        }
                                        else
                                        {
                                            dgScheduleManagement[1, curRow].Value = "-";
                                        }
                                        //Печатаем курс
                                        if (!(coll[j].KursNum == null))
                                        {
                                            dgScheduleManagement[2, curRow].Value = coll[j].KursNum.Kurs;
                                        }
                                        else
                                        {
                                            dgScheduleManagement[2, curRow].Value = "-";
                                        }
                                        //Печатаем специальность
                                        if (!(coll[j].Speciality == null))
                                        {
                                            dgScheduleManagement[3, curRow].Value = coll[j].Speciality.ShortUpravlenie
                                                + " (" + coll[j].Speciality.ShortInstitute + ")";
                                        }
                                        else
                                        {
                                            dgScheduleManagement[3, curRow].Value = "-";
                                        }
                                        //Печатаем название дисциплины
                                        dgScheduleManagement[4, curRow].Value = coll[j].Subject.Subject + " " + mdlData.commentLoad(coll[j]);
                                        //Суммарно на дисциплину
                                        dgScheduleManagement[6, curRow].Value = sumCurrent.ToString("0.00");

                                        sumHours += sumCurrent;
                                        //Переходим к следующей строке
                                        curRow += 1;
                                    }
                                }
                            }
                            //Если равномерно распределяемая нагрузка
                            else
                            {
                                sumCurrent = 0;
                                sumStud = 0;
                                for (k = 0; k <= mdlData.colStudents.Count - 1; k++)
                                {
                                    if (mdlData.colStudents[k].flgPlan)
                                    {
                                        //Если рассматриваемый преподаватель - руководитель студента
                                        //И если студент на том же курсе, где и дисциплина
                                        //И специальность студента должна соответствовать специальности нагрузки
                                        if (mdlData.colStudents[k].Lect.Equals(mdlData.colLecturer[i])
                                            & mdlData.colStudents[k].KursNum.Equals(coll[j].KursNum)
                                            & mdlData.colStudents[k].Speciality.Equals(coll[j].Speciality))
                                        {
                                            sumCurrent += coll[j].Weight;
                                            sumStud++;
                                        }
                                    }
                                }

                                //Если сумма изменилась
                                if (sumCurrent > 0)
                                {
                                    if (!flgHaveLoad)
                                    {
                                        //Печатаем фамилию, имя, отчество преподавателя
                                        dgScheduleManagement[0, curRow].Value = mdlData.colLecturer[i].FIO;
                                        //Получается, что у преподавателя есть нагрузка
                                        flgHaveLoad = true;
                                    }

                                    //Печатаем код семестра, учитывая смещение
                                    dgScheduleManagement[1, curRow].Value = coll[j].Semestr.Code - 1;
                                    //Печатаем курс
                                    if (!(coll[j].KursNum == null))
                                    {
                                        dgScheduleManagement[2, curRow].Value = coll[j].KursNum.Kurs;
                                    }
                                    else
                                    {
                                        dgScheduleManagement[2, curRow].Value = "-";
                                    }
                                    //Печатаем специальность
                                    if (!(coll[j].Speciality == null))
                                    {
                                        dgScheduleManagement[3, curRow].Value = coll[j].Speciality.ShortUpravlenie
                                                + " (" + coll[j].Speciality.ShortInstitute + ")";
                                    }
                                    else
                                    {
                                        dgScheduleManagement[3, curRow].Value = "-";
                                    }

                                    //Печатаем название дисциплины
                                    dgScheduleManagement[4, curRow].Value = coll[j].Subject.Subject + " " + mdlData.commentLoad(coll[j]);
                                    //Практические занятия в часах
                                    dgScheduleManagement[6, curRow].Value = sumCurrent;
                                    //Суммируем нагрузку преподавателя с учётом пройденных
                                    //позиций
                                    sumHours += sumCurrent;

                                    //Переходим к следующей строке
                                    curRow += 1;
                                }
                            }
                        }
                    }
                }

                //Только, если у преподавателя есть нагрузка
                if (flgHaveLoad)
                {
                    //Считаем дополнительную нагрузку преподавателя с
                    //учётом разгрузки других
                    if (mdlData.colLecturer[i].UnLoad == 0)
                    {
                        UnLoadDopFact = mdlData.colLecturer[i].Rate * mdlData.LoadInc;
                    }
                    else
                    {
                        UnLoadDopFact = 0;
                    }

                    PlanHours = (mdlData.colLecturer[i].Rate * AvgLoad) +
                                 mdlData.colLecturer[i].UnLoad + UnLoadDopFact;

                    //Пишем сумму по преподавателю
                    dgScheduleManagement[6, curRow].Value = sumHours.ToString("0.00") + " (" + PlanHours.ToString("0.00") + ")";

                    //Пишем баланс
                    dgScheduleManagement[7, curRow].Value = (sumHours - PlanHours).ToString("0.00") + " (" + ((sumHours - PlanHours) + UnLoadDopFact).ToString("0.00") + ")";

                    //Пишем ставку
                    dgScheduleManagement[8, curRow].Value = mdlData.colLecturer[i].Rate.ToString("0.00") + " (" +
                                                            mdlData.colLecturer[i].OldRate.ToString("0.00") + ")";

                    //Пишем догрузку
                    dgScheduleManagement[9, curRow].Value = UnLoadDopFact.ToString("0.00");

                    //Пишем разгрузку
                    dgScheduleManagement[10, curRow].Value = mdlData.colLecturer[i].UnLoad.ToString("0.00");

                    sumAll += sumHours;

                    //Разделяем преподавателей пустой строчкой
                    curRow += 2;
                }
            }

            //Пишем сумму всех и по всем
            curRow -= 1;
            dgScheduleManagement[6, curRow].Value = sumAll.ToString("0.00");
        }

        //Заполнить сетку нагрузки
        private void FillGridForLoad(IList<clsDistribution> coll, bool flgBoth, ref double SumRate,
                                        ref double sumNotUnLoadRate, ref int sumUnLoad)
        {
            //Берём первую строку (с нуля)
            int curRow = 0, sumHours, sumCurrent, sumStud, sumAll, AvgLoad = mdlData.AverageLoad;
            int i, j, k;
            bool flgHaveLoad;
            double UnLoadDopFact = 0, PlanHours = 0, sumDelta = 0, sumDeltaWONull = 0,
                    sumUpLoad = 0;
            int sumHoured1, sumHoured2, sumHouredAll;
            int sumHoured1Tot, sumHoured2Tot, sumHouredAllTot;
            double RealRate = 0;

            dgScheduleManagement[0, curRow].Value = "Сред.нагр.";
            dgScheduleManagement[1, curRow].Value = mdlData.AverageLoad.ToString("0.00");

            dgScheduleManagement[2, curRow].Value = "Сумм.ставка";
            dgScheduleManagement[3, curRow].Value = SumRate.ToString("0.00");

            dgScheduleManagement[4, curRow].Value = "Сумм.разгр.";
            dgScheduleManagement[5, curRow].Value = (-sumUnLoad).ToString("0.00");

            dgScheduleManagement[6, curRow].Value = "Сумм.став.б/р";
            dgScheduleManagement[7, curRow].Value = (sumNotUnLoadRate).ToString("0.00");

            //Берём третью строку (с нуля)
            curRow = 2;
            //Формируем шапку выписки
            //0. Преподаватель
            //1. Ставка
            //2. Норма часов
            //3. Догрузка
            //4. Разгрузка
            //5. Распределено
            //6. Отклонение от нормы
            //7. Почасовая в первом семестре
            //8. Почасовая во втором семестре
            //9. Суммарная почасовая
            //10. Компенсация отклонения
            dgScheduleManagement[0, curRow].Value = "Преподаватель";
            dgScheduleManagement[1, curRow].Value = "Ставка";
            dgScheduleManagement[2, curRow].Value = "Норма часов";
            dgScheduleManagement[3, curRow].Value = "Догрузка";
            dgScheduleManagement[4, curRow].Value = "Разгрузка";
            dgScheduleManagement[5, curRow].Value = "Распределено";
            dgScheduleManagement[6, curRow].Value = "Отклонение от нормы";
            dgScheduleManagement[7, curRow].Value = "Почасовая 1";
            dgScheduleManagement[8, curRow].Value = "Почасовая 2";
            dgScheduleManagement[9, curRow].Value = "Почасовая сум.";
            dgScheduleManagement[10, curRow].Value = "Компенсация отклонения";
            //Идём на следующую строку
            curRow += 1;

            sumAll = 0;
            sumHoured1Tot = 0;
            sumHoured2Tot = 0;
            sumHouredAllTot = 0;
            //Просматриваем каждого преподавателя
            for (i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                sumHours = 0;
                //По умолчанию считаем, что у преподавателя нет нагрузки
                flgHaveLoad = false;
                //Просматриваем нагрузку
                for (j = 0; j <= coll.Count - 1; j++)
                {
                    if (!coll[i].flgExclude)
                    {
                        sumCurrent = 0;

                        //Если стандартная строка нагрузки
                        if (!coll[j].flgDistrib)
                        {
                            //Если рассматриваемый преподаватель совпадает с
                            //преподавателем, указанным в нагрузке
                            if (mdlData.colLecturer[i].Equals(coll[j].Lecturer))
                            {
                                //Если дисциплина предполагает хоть какие-то часы, то
                                //выделяем под неё строчку текста
                                if (mdlData.NonZeroDistributionOR(coll[j]))
                                {
                                    if (!flgHaveLoad)
                                    {
                                        //Печатаем фамилию, имя, отчество преподавателя
                                        dgScheduleManagement[0, curRow].Value = mdlData.colLecturer[i].FIO;
                                        //Получается, что у преподавателя есть нагрузка
                                        flgHaveLoad = true;
                                    }

                                    sumCurrent += mdlData.toSumDistributionComponents(coll[j]);

                                    sumHours += sumCurrent;
                                }
                            }
                        }
                        //Если равномерно распределяемая нагрузка
                        else
                        {
                            sumCurrent = 0;
                            sumStud = 0;
                            for (k = 0; k <= mdlData.colStudents.Count - 1; k++)
                            {
                                if (mdlData.colStudents[k].flgPlan)
                                {
                                    //Если рассматриваемый преподаватель - руководитель студента
                                    //И если студент на том же курсе, где и дисциплина
                                    //И специальность студента должна соответствовать специальности нагрузки
                                    if (mdlData.colStudents[k].Lect.Equals(mdlData.colLecturer[i])
                                        & mdlData.colStudents[k].KursNum.Equals(coll[j].KursNum)
                                        & mdlData.colStudents[k].Speciality.Equals(coll[j].Speciality))
                                    {
                                        sumCurrent += coll[j].Weight;
                                        sumStud++;
                                    }
                                }
                            }

                            //Если сумма изменилась
                            if (sumCurrent > 0)
                            {
                                if (!flgHaveLoad)
                                {
                                    //Печатаем фамилию, имя, отчество преподавателя
                                    dgScheduleManagement[0, curRow].Value = mdlData.colLecturer[i].FIO;
                                    //Получается, что у преподавателя есть нагрузка
                                    flgHaveLoad = true;
                                }

                                //Суммируем нагрузку преподавателя с учётом пройденных
                                //позиций
                                sumHours += sumCurrent;
                            }
                        }
                    }

                    //Завершаем просмотр строки нагрузки
                }

                //Только, если у преподавателя есть нагрузка
                if (flgHaveLoad)
                {
                    if (mdlData.colLecturer[i].ChangeRate)
                    {
                        RealRate = (mdlData.colLecturer[i].Rate1 + mdlData.colLecturer[i].Rate2) / 2;
                    }
                    else
                    {
                        RealRate = mdlData.colLecturer[i].Rate;
                    }
                    
                    //Считаем дополнительную нагрузку преподавателя с
                    //учётом разгрузки других (ставка по отделу кадров)
                    if (mdlData.colLecturer[i].UnLoad == 0)
                    {
                        UnLoadDopFact = mdlData.colLecturer[i].Rate * mdlData.LoadInc;
                    }
                    else
                    {
                        UnLoadDopFact = 0;
                    }

                    //Здесь должна быть реальная ставка
                    PlanHours = (RealRate * AvgLoad) +
                                 mdlData.colLecturer[i].UnLoad + UnLoadDopFact;

                    //Пишем ставку (по отделу кадров)
                    dgScheduleManagement[1, curRow].Value = mdlData.colLecturer[i].Rate.ToString("0.00");
                    //Если ставка меняется от семестра к семестру, то добавляем метку "*"
                    if (mdlData.colLecturer[i].ChangeRate)
                    {
                        dgScheduleManagement[1, curRow].Value += "*";
                    }

                    //Пишем норму часов на ставку
                    dgScheduleManagement[2, curRow].Value = PlanHours.ToString("0.00");

                    //Пишем догрузку
                    dgScheduleManagement[3, curRow].Value = UnLoadDopFact.ToString("0.00");

                    //Пишем разгрузку
                    dgScheduleManagement[4, curRow].Value = mdlData.colLecturer[i].UnLoad.ToString("0.00");

                    //Пишем сумму по преподавателю
                    dgScheduleManagement[5, curRow].Value = sumHours.ToString("0.00");

                    //Пишем отклонение от нормы
                    dgScheduleManagement[6, curRow].Value = (sumHours - PlanHours).ToString("0.00");

                    //Отклонение без нулевых ставок
                    if (mdlData.colLecturer[i].Rate > 0)
                    {
                        sumDeltaWONull += (sumHours - PlanHours);
                    }

                    //Суммарное отклонение
                    sumDelta += (sumHours - PlanHours);

                    sumUpLoad += UnLoadDopFact;

                    sumAll += sumHours;

                    sumHoured1 = 0;
                    sumHoured2 = 0;
                    sumHouredAll = 0;

                    if (!coll.Equals(mdlData.colHouredDistribution))
                    {
                        for (j = 0; j <= mdlData.colHouredDistribution.Count - 1; j++)
                        {
                            //if (mdlData.colHouredDistribution[j].Lecturer.FIO.Contains("Лызлов"))
                            //{
                            //    MessageBox.Show("");
                            //}

                            if (mdlData.colHouredDistribution[j].Lecturer.Equals(mdlData.colLecturer[i]))
                            {
                                if (mdlData.colHouredDistribution[j].Semestr.Equals(mdlData.colSemestr[1]))
                                {
                                    sumHoured1 += mdlData.toSumDistributionComponents(mdlData.colHouredDistribution[j]);
                                }

                                if (mdlData.colHouredDistribution[j].Semestr.Equals(mdlData.colSemestr[2]))
                                {
                                    sumHoured2 += mdlData.toSumDistributionComponents(mdlData.colHouredDistribution[j]);
                                }
                            }
                        }
                    }

                    sumHoured1Tot += sumHoured1;
                    sumHoured2Tot += sumHoured2;
                    sumHouredAll = sumHoured1 + sumHoured2;
                    sumHouredAllTot += sumHouredAll;

                    dgScheduleManagement[7, curRow].Value = (sumHoured1).ToString("0.00");
                    dgScheduleManagement[8, curRow].Value = (sumHoured2).ToString("0.00");
                    dgScheduleManagement[9, curRow].Value = (sumHouredAll).ToString("0.00");
                    dgScheduleManagement[10, curRow].Value = ((sumHours - PlanHours) - sumHouredAll).ToString("0.00");

                    //Переходим на следующую строку
                    curRow += 1;

                }

                //Завершаем просмотр преподавателя
            }

            //Пишем суммарную догрузку
            dgScheduleManagement[3, curRow].Value = (sumUpLoad).ToString("0.00");

            //Пишем суммарную разгрузку
            dgScheduleManagement[4, curRow].Value = (-sumUnLoad).ToString("0.00");

            //Пишем сумму всех и по всем
            dgScheduleManagement[5, curRow].Value = sumAll.ToString("0.00");

            //Пишем суммарную почасовую первого семестра
            dgScheduleManagement[7, curRow].Value = sumHoured1Tot.ToString("0.00");

            //Пишем суммарную почасовую второго семестра
            dgScheduleManagement[8, curRow].Value = sumHoured2Tot.ToString("0.00");

            //Пишем суммарную почасовую
            dgScheduleManagement[9, curRow].Value = sumHouredAllTot.ToString("0.00");

            dgScheduleManagement[10, curRow].Value = (sumDelta - sumHouredAllTot).ToString("0.00");

            curRow += 1;

            //Пишем сумму отклонений по всем
            dgScheduleManagement[6, curRow].Value = sumDelta.ToString("0.00");

            curRow += 1;

            //Пишем сумму отклонений по всем
            dgScheduleManagement[5, curRow].Value = "Сумм. отклонений без нулевых ставок:";
            dgScheduleManagement[6, curRow].Value = sumDeltaWONull.ToString("0.00");
        }

        //Заполнить сетку равномерно распределяемой нагрузки
        private void FillGridForUniformLoad(IList<clsDistribution> coll)
        {
            //Берём первую строку (с нуля)
            int curRow = 0;
            int sumCurrent, sumStud;
            int i, j, k;
            bool flgHaveUniformLoad;

            //Формируем шапку выписки
            //0. Преподаватель
            //1. Семестр
            //2. Группа
            //3. Курс
            //4. Часы
            dgScheduleManagement[0, curRow].Value = "Преподаватель / наименование";
            dgScheduleManagement[1, curRow].Value = "Семестр";
            dgScheduleManagement[2, curRow].Value = "Группа";
            dgScheduleManagement[3, curRow].Value = "Курс";
            dgScheduleManagement[4, curRow].Value = "Часы";

            //Идём на следующую строку
            curRow += 1;

            //Просматриваем каждого преподавателя
            for (i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                //По умолчанию считаем, что у преподавателя нет 
                //равномерно распределяемой нагрузки
                flgHaveUniformLoad = false;
                //Просматриваем нагрузку
                for (j = 0; j <= coll.Count - 1; j++)
                {
                    //Если элемент не исключён из расчёта нагрузки
                    if (!coll[i].flgExclude)
                    {
                        //Если равномерно распределяемая нагрузка
                        if (coll[j].flgDistrib)
                        {
                            sumCurrent = 0;
                            sumStud = 0;
                            //
                            for (k = 0; k <= mdlData.colStudents.Count - 1; k++)
                            {
                                //
                                if (mdlData.colStudents[k].flgPlan)
                                {
                                    //Если рассматриваемый преподаватель - руководитель студента
                                    //И если студент на том же курсе, где и дисциплина
                                    //И специальность студента должна соответствовать специальности нагрузки
                                    if (mdlData.colStudents[k].Lect.Equals(mdlData.colLecturer[i])
                                        & mdlData.colStudents[k].KursNum.Equals(coll[j].KursNum)
                                        & mdlData.colStudents[k].Speciality.Equals(coll[j].Speciality))
                                    {
                                        sumCurrent += coll[j].Weight;
                                        sumStud++;
                                    }
                                }
                            }

                            //Если сумма изменилась
                            if (sumCurrent > 0)
                            {
                                if (!flgHaveUniformLoad)
                                {
                                    //Печатаем фамилию, имя, отчество преподавателя
                                    dgScheduleManagement[0, curRow].Value = mdlData.colLecturer[i].FIO;
                                    //Получается, что у преподавателя есть нагрузка
                                    flgHaveUniformLoad = true;
                                    //Переходим к следующей строке
                                    curRow += 1;
                                }

                                //Печатаем наименование дисциплины
                                dgScheduleManagement[0, curRow].Value = coll[j].Subject.Subject;
                                dgScheduleManagement[1, curRow].Value = coll[j].Semestr.SemNum;
                                dgScheduleManagement[2, curRow].Value = coll[j].Speciality.ShortInstitute;
                                dgScheduleManagement[3, curRow].Value = coll[j].KursNum.Kurs;
                                dgScheduleManagement[4, curRow].Value = sumCurrent;
                                //Переходим к следующей строке
                                curRow += 1;
                            }
                        }
                    }
                }
                //Завершается просмотр нагрузки

                if (flgHaveUniformLoad)
                {
                    //Переходим к следующей строке (разделяем строками преподавателей)
                    curRow += 1;
                }
            }
        }

        /// <summary>
        /// Распределение нагрузки для начальника в удобном формате
        /// </summary>
        private void BossGrid()
        {
            IList<clsDistribution> coll = new List<clsDistribution>();
            int sumUnLoad = 0;
            bool flgBoth;
            double SumRate = 0, SumRateOld = 0, sumNotUnLoadRate = 0;

            //Выполняем сортировку
            coll = SortLoad(coll);

            //Очищаем сетку
            dgScheduleManagement.Rows.Clear();
            dgScheduleManagement.Columns.Clear();

            //По умолчанию работаем с одним из выбранных семестров
            flgBoth = false;
            //Но если семестр не выбран из списка, то
            if (mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum == "-")
            {
                //считаем, что надо выдавать информацию за оба семестра
                flgBoth = true;
            }

            //Подготовка сетки для заведующего кафедрой
            PrepareGridForBoss(coll, flgBoth, ref SumRate, ref SumRateOld, ref sumNotUnLoadRate, ref sumUnLoad);

            FillGridForBoss(coll, flgBoth, ref SumRate, ref SumRateOld, ref sumNotUnLoadRate, ref sumUnLoad);

            //-----------------------------------------------------------------
            //---------------Фрагмент заполнения рабочей области---------------
            //-----------------------------------------------------------------
        }

        /// <summary>
        /// Значения нагрузки, разгрузки и перегрузки в удобном формате
        /// </summary>
        private void LoadUnloadGrid()
        {
            IList<clsDistribution> coll = new List<clsDistribution>();
            int sumUnLoad = 0;
            bool flgBoth;
            double SumRate = 0, sumNotUnLoadRate = 0;

            //Выполняем сортировку
            coll = SortLoad(coll);

            //Очищаем сетку
            dgScheduleManagement.Rows.Clear();
            dgScheduleManagement.Columns.Clear();

            //По умолчанию работаем с одним из выбранных семестров
            flgBoth = true;

            //Подготовка сетки нагрузки
            PrepareGridForLoad(coll, flgBoth, ref SumRate, ref sumNotUnLoadRate, ref sumUnLoad);

            FillGridForLoad(coll, flgBoth, ref SumRate, ref sumNotUnLoadRate, ref sumUnLoad);

            //-----------------------------------------------------------------
            //---------------Фрагмент заполнения рабочей области---------------
            //-----------------------------------------------------------------
        }

        /// <summary>
        /// Элементы равномерно распределяемой нагрузки
        /// </summary>
        private void UniformLoadGrid()
        {
            IList<clsDistribution> coll = new List<clsDistribution>();

            //Выполняем сортировку
            coll = SortLoad(coll);

            //Очищаем сетку
            dgScheduleManagement.Rows.Clear();
            dgScheduleManagement.Columns.Clear();

            //Подготовка сетки равномерно распределяемой нагрузки
            PrepareGridForUniformLoad(coll);
            //Заполнение сетки равномерно распределяемой нагрузкой
            FillGridForUniformLoad(coll);

            //-----------------------------------------------------------------
            //---------------Фрагмент заполнения рабочей области---------------
            //-----------------------------------------------------------------
        }

        private void KursGrid()
        {
            int curRow;
            int sumHours;
            int sumCurrent;
            int sumAll;
            bool HaveLoad;
            bool Both;

            double SumRate = 0;
            double SumRateOld = 0;
            int sumUnLoad = 0;
            double sumNotUnLoadRate = 0;

            //Значение средней нагрузки на кафедру
            int AvgLoad = mdlData.AverageLoad;
            //Разгрузка дополнительная фактическая
            double UnLoadDopFact = 0;
            //Плановые часы нагрузки преподавателя
            double PlanHours = 0;

            //Очищаем сетку
            dgScheduleManagement.Rows.Clear();
            dgScheduleManagement.Columns.Clear();

            //-----------------------------------------------------------------
            //---------------Фрагмент обозначения рабочей области--------------
            //-----------------------------------------------------------------

            //Задаём количество столбцов
            //оно остаётся неизменным
            //0. Преподаватель
            //1. Семестр
            //2. Курс
            //3. Специальность
            //4. Название дисциплины
            //5. (пусто)
            //6. Сумма часов
            //7. Баланс
            //8. Ставка
            //9. Догрузка
            //10. Разгрузка

            for (int i = 0; i <= 10; i++)
            {
                dgScheduleManagement.Columns.Add("", "");
            }

            Both = false;
            if (mdlData.colSemestr[cmbSemestrList.SelectedIndex].SemNum == "-")
            {
                Both = true;
            }

            //Сразу добавляем две строки (под общую информацию)
            dgScheduleManagement.Rows.Add();
            //и под пробел
            dgScheduleManagement.Rows.Add();

            //Просматриваем каждого преподавателя
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                SumRate += mdlData.colLecturer[i].Rate;
                SumRateOld += mdlData.colLecturer[i].OldRate;

                //Считаем догрузку для преподавателей без разгрузки
                if (mdlData.colLecturer[i].UnLoad == 0)
                {
                    sumNotUnLoadRate += mdlData.colLecturer[i].Rate;
                }
                else
                {
                    sumUnLoad += mdlData.colLecturer[i].UnLoad;
                }

                //Считаем, по умолчанию, что преподаватель не нагружен
                HaveLoad = false;
                //Просматриваем нагрузку
                for (int j = 0; j <= mdlData.colDistribution.Count - 1; j++)
                {
                    if (!Both)
                    {
                        //Если дисциплина относится ко второму семестру
                        if (mdlData.colDistribution[j].Semestr.Equals(mdlData.colSemestr[cmbSemestrList.SelectedIndex]))
                        {
                            //Если рассматриваемый преподаватель совпадает с
                            //преподавателем, указанным в нагрузке
                            if (mdlData.colLecturer[i].Equals(mdlData.colDistribution[j].Lecturer))
                            {
                                //Если дисциплина предполагает часы
                                //под курсовое проектирование
                                if (mdlData.colDistribution[j].KursProject > 0)
                                {
                                    //Если преподаватель что-либо из этого ведёт
                                    //добавляем строку
                                    dgScheduleManagement.Rows.Add();
                                    //Получается, что преподаватель нагружен
                                    HaveLoad = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        //Если рассматриваемый преподаватель совпадает с
                        //преподавателем, указанным в нагрузке
                        if (mdlData.colLecturer[i].Equals(mdlData.colDistribution[j].Lecturer))
                        {
                            //Если дисциплина предполагает часы
                            //под курсовое проектирование
                            if (mdlData.colDistribution[j].KursProject > 0)
                            {
                                //Если преподаватель что-либо из этого ведёт
                                //добавляем строку
                                dgScheduleManagement.Rows.Add();
                                //Получается, что преподаватель нагружен
                                HaveLoad = true;
                            }
                        }
                    }
                }

                //Только если у преподавателя есть нагрузка
                if (HaveLoad)
                {
                    //Разделяем преподавателей пустой строчкой
                    dgScheduleManagement.Rows.Add();
                    dgScheduleManagement.Rows.Add();
                }
            }

            mdlData.LoadInc = Convert.ToInt32(-sumUnLoad / sumNotUnLoadRate);

            //-----------------------------------------------------------------
            //---------------Фрагмент обозначения рабочей области--------------
            //-----------------------------------------------------------------

            //-----------------------------------------------------------------
            //---------------Фрагмент заполнения рабочей области---------------
            //-----------------------------------------------------------------

            //Берём первую строку (с нуля)
            curRow = 0;

            dgScheduleManagement[0, curRow].Value = "Сред.нагр.";
            dgScheduleManagement[1, curRow].Value = mdlData.AverageLoad.ToString("0.00");

            dgScheduleManagement[2, curRow].Value = "Сумм.ставка";
            dgScheduleManagement[3, curRow].Value = SumRate.ToString("0.00") + " (" + SumRateOld.ToString("0.00") + ")";

            dgScheduleManagement[4, curRow].Value = "Сумм.разгр.";
            dgScheduleManagement[5, curRow].Value = (-sumUnLoad).ToString("0.00");

            dgScheduleManagement[6, curRow].Value = "Сумм.став.б/р";
            dgScheduleManagement[7, curRow].Value = (sumNotUnLoadRate).ToString("0.00");

            //Берём третью строку (с нуля)
            curRow = 2;
            //Формируем шапку выписки
            dgScheduleManagement[0, curRow].Value = "Преподаватель";
            dgScheduleManagement[1, curRow].Value = "Сем.";
            dgScheduleManagement[2, curRow].Value = "Курс";
            dgScheduleManagement[3, curRow].Value = "Специальность";
            dgScheduleManagement[4, curRow].Value = "Название дисциплины";
            dgScheduleManagement[6, curRow].Value = "Сумма часов";
            dgScheduleManagement[7, curRow].Value = "Баланс";
            dgScheduleManagement[8, curRow].Value = "Ставка";
            dgScheduleManagement[9, curRow].Value = "Догрузка";
            dgScheduleManagement[10, curRow].Value = "Разгрузка";
            //Идём на следующую строку
            curRow += 1;

            //Просматриваем каждого преподавателя
            sumAll = 0;
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                sumHours = 0;
                //По умолчанию считаем, что у преподавателя нет нагрузки
                HaveLoad = false;
                //Просматриваем нагрузку
                for (int j = 0; j <= mdlData.colDistribution.Count - 1; j++)
                {
                    sumCurrent = 0;
                    if (!Both)
                    {
                        //Если дисциплина относится к выбранному семестру
                        if (mdlData.colDistribution[j].Semestr.Equals(mdlData.colSemestr[cmbSemestrList.SelectedIndex]))
                        {
                            //Если рассматриваемый преподаватель совпадает с
                            //преподавателем, указанным в нагрузке
                            if (mdlData.colLecturer[i].Equals(mdlData.colDistribution[j].Lecturer))
                            {
                                //Курсовой проект
                                if (mdlData.colDistribution[j].KursProject > 0)
                                {
                                    if (!HaveLoad)
                                    {
                                        //Печатаем фамилию, имя, отчество преподавателя
                                        dgScheduleManagement[0, curRow].Value = mdlData.colLecturer[i].FIO;
                                        //Получается, что у преподавателя есть нагрузка
                                        HaveLoad = true;
                                    }

                                    sumCurrent += mdlData.colDistribution[j].KursProject;
                                        
                                    //Печатаем код семестра, учитывая смещение
                                    dgScheduleManagement[1, curRow].Value = mdlData.colDistribution[j].Semestr.Code - 1;
                                    //Печатаем курс
                                    if (!(mdlData.colDistribution[j].KursNum == null))
                                    {
                                        dgScheduleManagement[2, curRow].Value = mdlData.colDistribution[j].KursNum.Kurs;
                                    }
                                    else
                                    {
                                        dgScheduleManagement[2, curRow].Value = "-";
                                    }
                                    //Печатаем специальность
                                    if (!(mdlData.colDistribution[j].Speciality == null))
                                    {
                                        dgScheduleManagement[3, curRow].Value = mdlData.colDistribution[j].Speciality.ShortUpravlenie
                                                + " (" + mdlData.colDistribution[j].Speciality.ShortInstitute + ")";
                                    }
                                    else
                                    {
                                        dgScheduleManagement[3, curRow].Value = "-";
                                    }
                                    //Печатаем название дисциплины
                                    if (mdlData.colDistribution[j].Subject.Subject.Length > 80)
                                    {
                                        dgScheduleManagement[4, curRow].Value = mdlData.colDistribution[j].Subject.Subject.Substring(0, 80);
                                    }
                                    else
                                    {
                                        dgScheduleManagement[4, curRow].Value = mdlData.colDistribution[j].Subject.Subject;
                                    }
                                    //Практические занятия в часах
                                    dgScheduleManagement[6, curRow].Value = sumCurrent;

                                    sumHours += sumCurrent;

                                    //Переходим к следующей строке
                                    curRow += 1;
                                }
                            }
                        }
                    }
                    else
                    {
                        //Если рассматриваемый преподаватель совпадает с
                        //преподавателем, указанным в нагрузке
                        if (mdlData.colLecturer[i].Equals(mdlData.colDistribution[j].Lecturer))
                        {
                            //Курсовой проект
                            if (mdlData.colDistribution[j].KursProject > 0)
                            {
                                if (!HaveLoad)
                                {
                                    //Печатаем фамилию, имя, отчество преподавателя
                                    dgScheduleManagement[0, curRow].Value = mdlData.colLecturer[i].FIO;
                                    //Получается, что у преподавателя есть нагрузка
                                    HaveLoad = true;
                                }

                                sumCurrent += mdlData.colDistribution[j].KursProject;

                                //Печатаем код семестра, учитывая смещение
                                if (!(mdlData.colDistribution[j].Semestr == null))
                                {
                                    dgScheduleManagement[1, curRow].Value = mdlData.colDistribution[j].Semestr.Code - 1;
                                }
                                else
                                {
                                    dgScheduleManagement[1, curRow].Value = "-";
                                }
                                //Печатаем курс
                                if (!(mdlData.colDistribution[j].KursNum == null))
                                {
                                    dgScheduleManagement[2, curRow].Value = mdlData.colDistribution[j].KursNum.Kurs;
                                }
                                else
                                {
                                    dgScheduleManagement[2, curRow].Value = "-";
                                }
                                //Печатаем специальность
                                if (!(mdlData.colDistribution[j].Speciality == null))
                                {
                                    dgScheduleManagement[3, curRow].Value = mdlData.colDistribution[j].Speciality.ShortUpravlenie
                                        + " (" + mdlData.colDistribution[j].Speciality.ShortInstitute + ")";
                                }
                                else
                                {
                                    dgScheduleManagement[3, curRow].Value = "-";
                                }
                                //Печатаем название дисциплины
                                if (mdlData.colDistribution[j].Subject.Subject.Length > 80)
                                {
                                    dgScheduleManagement[4, curRow].Value = mdlData.colDistribution[j].Subject.Subject.Substring(0, 80);
                                }
                                else
                                {
                                    dgScheduleManagement[4, curRow].Value = mdlData.colDistribution[j].Subject.Subject;
                                }
                                //Суммарно на дисциплину
                                dgScheduleManagement[6, curRow].Value = sumCurrent.ToString("0.00");

                                sumHours += sumCurrent;
                                //Переходим к следующей строке
                                curRow += 1;
                            }
                        }
                    }
                }

                //Только, если у преподавателя есть нагрузка
                if (HaveLoad)
                {

                    //Считаем дополнительную нагрузку преподавателя с
                    //учётом разгрузки других
                    if (mdlData.colLecturer[i].UnLoad == 0)
                    {
                        UnLoadDopFact = mdlData.colLecturer[i].Rate * mdlData.LoadInc;
                    }
                    else
                    {
                        UnLoadDopFact = 0;
                    }

                    PlanHours = (mdlData.colLecturer[i].Rate * AvgLoad) +
                                 mdlData.colLecturer[i].UnLoad + UnLoadDopFact;

                    //Пишем сумму по преподавателю
                    dgScheduleManagement[6, curRow].Value = sumHours.ToString("0.00") + " (" + PlanHours.ToString("0.00") + ")";

                    dgScheduleManagement[7, curRow].Value = (sumHours - PlanHours).ToString("0.00");

                    dgScheduleManagement[8, curRow].Value = mdlData.colLecturer[i].Rate.ToString("0.00") + " (" +
                                                            mdlData.colLecturer[i].OldRate.ToString("0.00") + ")";

                    dgScheduleManagement[9, curRow].Value = UnLoadDopFact.ToString("0.00");

                    dgScheduleManagement[10, curRow].Value = mdlData.colLecturer[i].UnLoad.ToString("0.00");

                    sumAll += sumHours;

                    //Разделяем преподавателей пустой строчкой
                    curRow += 2;
                }
            }

            curRow -= 1;
            dgScheduleManagement[6, curRow].Value = sumAll.ToString("0.00");
            //-----------------------------------------------------------------
            //---------------Фрагмент заполнения рабочей области---------------
            //-----------------------------------------------------------------
        }

        private void LecturerGraduateGrid()
        {
            int curRow, colStud;
            string str;
            bool HaveLoad;

            int i, j, k, l;

            //Очищаем сетку
            dgScheduleManagement.Rows.Clear();
            dgScheduleManagement.Columns.Clear();

            //-----------------------------------------------------------------
            //---------------Фрагмент обозначения рабочей области--------------
            //-----------------------------------------------------------------

            //Задаём количество столбцов
            //оно остаётся неизменным
            //0. № п/п
            //1. Руководитель
            //2. Курс
            //3. Специальность
            //4. Количество
            //5. Кто?

            for (i = 0; i <= 5; i++)
            {
                dgScheduleManagement.Columns.Add("", "");
            }

            //Сразу добавляем две строки (под общую информацию)
            dgScheduleManagement.Rows.Add();
            //и под пробел
            dgScheduleManagement.Rows.Add();

            //Просматриваем курсы
            for (i = 0; i <= mdlData.colKursNum.Count - 1; i++)
            {
                //Просматриваем каждого преподавателя
                for (j = 0; j <= mdlData.colLecturer.Count - 1; j++)
                {
                    //Просматриваем нагрузку
                    for (k = 0; k <= mdlData.colDistribution.Count - 1; k++)
                    {
                        if (mdlData.colDistribution[k].flgDistrib)
                        {
                            HaveLoad = false;
                            for (l = 0; l <= mdlData.colStudents.Count - 1; l++)
                            {
                                if (mdlData.colStudents[l].Lect.Equals(mdlData.colLecturer[j])
                                    & mdlData.colStudents[l].KursNum.Equals(mdlData.colKursNum[i])
                                    & mdlData.colStudents[l].Speciality.Equals(mdlData.colDistribution[k].Speciality))
                                {
                                    //Разделяем преподавателей пустой строчкой
                                    dgScheduleManagement.Rows.Add();
                                    HaveLoad = true;
                                    break;
                                }
                            }

                            if (HaveLoad)
                            {
                                break;
                            }
                        }
                    }
                }
            }

            //-----------------------------------------------------------------
            //---------------Фрагмент обозначения рабочей области--------------
            //-----------------------------------------------------------------

            //-----------------------------------------------------------------
            //---------------Фрагмент заполнения рабочей области---------------
            //-----------------------------------------------------------------

            //Берём первую строку (с нуля)
            curRow = 0;

            dgScheduleManagement[0, curRow].Value = "№ п/п";
            dgScheduleManagement.Columns[0].Width = 50;
            dgScheduleManagement[1, curRow].Value = "Руководитель";
            dgScheduleManagement.Columns[1].Width = 140;
            dgScheduleManagement[2, curRow].Value = "Курс";
            dgScheduleManagement.Columns[2].Width = 50;
            dgScheduleManagement[3, curRow].Value = "Специальность";
            dgScheduleManagement.Columns[3].Width = 140;
            dgScheduleManagement[4, curRow].Value = "Количество";
            dgScheduleManagement.Columns[4].Width = 70;
            dgScheduleManagement[5, curRow].Value = "Кто?";
            dgScheduleManagement.Columns[5].Width = 240;

            //Идём на следующую строку
            curRow++;

            for (i = 0; i <= mdlData.colKursNum.Count - 1; i++)
            {
                //Просматриваем каждого преподавателя
                for (j = 0; j <= mdlData.colLecturer.Count - 1; j++)
                {
                    //Просматриваем нагрузку
                    for (k = 0; k <= mdlData.colDistribution.Count - 1; k++)
                    {
                        if (mdlData.colDistribution[k].flgDistrib)
                        {
                            colStud = 0;
                            str = "";
                            for (l = 0; l <= mdlData.colStudents.Count - 1; l++)
                            {
                                if (mdlData.colStudents[l].Lect.Equals(mdlData.colLecturer[j])
                                    & mdlData.colStudents[l].KursNum.Equals(mdlData.colKursNum[i])
                                    & mdlData.colStudents[l].Speciality.Equals(mdlData.colDistribution[k].Speciality))
                                {
                                    colStud++;
                                    str += mdlData.SplitFIOString(mdlData.colStudents[l].FIO, true, false) + ", ";
                                }
                            }

                            if (colStud > 0)
                            {
                                dgScheduleManagement[0, curRow].Value = curRow;
                                dgScheduleManagement[1, curRow].Value = mdlData.colLecturer[j].FIO;
                                dgScheduleManagement[2, curRow].Value = mdlData.colKursNum[i].Kurs;
                                dgScheduleManagement[3, curRow].Value = mdlData.colDistribution[k].Speciality.ShortUpravlenie +
                                    " (" + mdlData.colDistribution[k].Speciality.ShortInstitute + ")";

                                dgScheduleManagement[4, curRow].Value = colStud;
                                str = str.Substring(0, str.Length - 2);
                                dgScheduleManagement[5, curRow].Value = str;
                                curRow++;
                                break;
                            }
                        }
                    }
                }
            }

            //-----------------------------------------------------------------
            //---------------Фрагмент заполнения рабочей области---------------
            //-----------------------------------------------------------------
        }

        private double[] countRates()
        {
            bool accessFlg = false;
            double[] param = new double[4];

            //Просматриваем каждого преподавателя
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                //Сначала определяем, нужно ли, в принципе, 
                //рассматривать преподавателя
                if (optMain.Checked || optCombine.Checked)
                {
                    accessFlg = (mdlData.colLecturer[i].Rate > 0);
                }
                else
                {
                    if (optHoured.Checked || optMainDop.Checked || optCombineDop.Checked)
                    {
                        accessFlg = true;
                    }
                }

                //если прошёл отбор
                if (accessFlg)
                {
                    if (mdlData.colLecturer[i].Duty.Short.Equals("асс."))
                    {
                        param[0] += mdlData.colLecturer[i].Rate;
                    }

                    if (mdlData.colLecturer[i].Duty.Short.Equals("доц."))
                    {
                        param[1] += mdlData.colLecturer[i].Rate;
                    }

                    if (mdlData.colLecturer[i].Duty.Short.Equals("ст.преп."))
                    {
                        param[2] += mdlData.colLecturer[i].Rate;
                    }

                    if (mdlData.colLecturer[i].Duty.Short.Equals("проф."))
                    {
                        param[3] += mdlData.colLecturer[i].Rate;
                    }

                    if (mdlData.colLecturer[i].Duty.Short.Equals("зав.каф."))
                    {
                        param[3] += mdlData.colLecturer[i].Rate;
                    }
                }
            }

            return param;
        }

        private void intoExcel2015()
        {
            int curLect;

            int sumLecI;    int sumLecII;

            int sumExamI;   int sumExamII;

            int sumCredI;   int sumCredII;

            int sumTutI;    int sumTutII;

            int sumLabI;    int sumLabII;

            int sumPracI;   int sumPracII;

            int sumRefI;    int sumRefII;

            int sumIndI;    int sumIndII;

            int sumKRAPKI;  int sumKRAPKII;

            int sumKursPrI; int sumKursPrII;

            int sumDiplI;   int sumDiplII;

            int sumTutPrI;  int sumTutPrII;

            int sumPreDipI; int sumPreDipII;

            int sumGAKI;    int sumGAKII;

            int sumPostGrI; int sumPostGrII;

            int sumVisI;    int sumVisII;

            int sumMagI;    int sumMagII;

            int countStud, countWeight;

            double[] rateParams = new double[4];
            double assist;
            double hitutor;
            double proff;
            double lecturer;
            double sumRate;

            string type = "";

            IList<clsDistribution> coll = null;
            
            bool accessFlg = false;
            bool flg;

            string strSumI = "=";
            string strSumII = "=";
            string strSemI;
            string strSemII;
            string strSum = "=";

            rateParams = countRates();

            assist = rateParams[0];
            lecturer = rateParams[1];
            hitutor = rateParams[2];
            proff = rateParams[3];

            sumRate = proff + hitutor + lecturer + assist;

            if (optMain.Checked || optMainDop.Checked)
            {
                coll = mdlData.colDistribution;
                type = " штатная общая ";
            }
            else
            {
                if (optHoured.Checked)
                {
                    if (!chkPlan.Checked)
                    {
                        coll = mdlData.colHouredDistribution;
                    }
                    else
                    {
                        coll = mdlData.colPlanHouredDistribution;
                    }

                    type = " почасовая ";
                }
                else
                {
                    if (optCombine.Checked || optCombineDop.Checked)
                    {
                        if (!chkPlan.Checked)
                        {
                            coll = mdlData.colCombineDistribution;
                        }
                        else
                        {
                            coll = mdlData.colPlanCombineDistribution;
                        }

                        type = " штатная с учётом почасовой ";
                    }
                }
            }

            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;

            try
            {
                //Создаём новое Excel приложение
                Excel.Application ObjExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook;
                Excel.Worksheet ObjWorkSheet;

                if (MessageBox.Show("Отображать ход заполнения?", "Отображение таблицы Excel", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    ObjExcel.Visible = true;
                }
                else
                {
                    ObjExcel.Visible = false;
                }

                //Книга
                ObjWorkBook = ObjExcel.Workbooks.Add(Missing.Value);
                //Таблица
                ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
                //Альбомная ориентация страницы
                ObjWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                //Книжная ориентация страницы
                //ObjWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                //В высоту разместить на одной странице
                ObjWorkSheet.PageSetup.FitToPagesTall = 1;
                //В ширину разместить на одной странице
                ObjWorkSheet.PageSetup.FitToPagesWide = 1;

                //-----------Формируем заголовок таблицы

                //Задаём диапазон для ячеек, подлежащих форматированию
                //1-я строка, с А по AA
                var cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(1, 1),
                    mdlData.ExcelCellTranslator(1, 27));
                //Выделяем ячейки диапазона
                cells.Select();
                //Объединяем ячейки диапазона
                cells.Merge(true);
                //Выравнивание в объединённой ячейке по центру
                cells.HorizontalAlignment = Excel.Constants.xlCenter;
                //Записываем текст в объединённую ячейку
                //(считается по первой А1)
                if (optHoured.Checked)
                {
                    //Если почасовая
                    if (chkPlan.Checked)
                    {
                        //Если плановая
                        cells.Cells[1, 1] = "Сведения о распределении почасовой нагрузки кафедры на учебный год";
                    }
                    else
                    {
                        //Если исполненная
                        cells.Cells[1, 1] = "Сведения о фактически выполненной (почасовой) учебной нагрузке за учебный год";
                    }
                }
                else
                {
                    //Если штатная
                    if (chkPlan.Checked)
                    {
                        //Если плановая
                        cells.Cells[1, 1] = "Сведения о распределении штатной нагрузки кафедры на учебный год";
                    }
                    else
                    {
                        //Если исполненная
                        cells.Cells[1, 1] = "Сведения о фактически выполненной (штатной) учебной нагрузке за учебный год";
                    }
                }

                //Задаём диапазон для ячеек, подлежащих форматированию
                //2-я строка, с А по AA
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(2, 1),
                    mdlData.ExcelCellTranslator(2, 27));
                //Выделяем ячейки диапазона
                cells.Select();
                //Объединяем ячейки диапазона
                cells.Merge(true);
                //Выравнивание в объединённой ячейке по центру
                cells.HorizontalAlignment = Excel.Constants.xlCenter;
                //Записываем текст в объединённую ячейку
                //(считается по первой А2)
                cells.Cells[1, 1] = "Кафедра \"Управление и защита информации\" " + 
                    mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear +
                    " учебный год";

                //-----------Формируем заголовок таблицы

                //-----------Формируем шапку таблицы

                for (int i = 1; i <= 27; i++)
                {
                    //Выбираем диапазон
                    cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4, i),
                        mdlData.ExcelCellTranslator(5, i));
                    //Выделяем ячейки диапазона
                    cells.Select();
                    //Объединяем ячейки диапазона
                    cells.Merge();
                    //Горизонтальное выравнивание по центру в ячейках
                    cells.HorizontalAlignment = Excel.Constants.xlCenter;
                    //Вертикальное выравнивание по центру в ячейках
                    cells.VerticalAlignment = Excel.Constants.xlCenter;
                    //Задаём границы
                    cells.Borders.Weight = 2;
                    //Перенос по словам без сокрытия под ячейками
                    cells.WrapText = true;
                    //Назначаем высоту шапки
                    cells.RowHeight = (68.25f + 66) / 2;

                    switch (i)
                    {
                        case 1:
                            cells.Cells[1, 1] = "№ п/п";
                            cells.ColumnWidth = 3.86f;
                            cells.Orientation = Excel.XlOrientation.xlHorizontal;
                            break;
                        case 2:
                            cells.Cells[1, 1] = "Фамилия, Имя, Отчество";
                            cells.ColumnWidth = 27.57f;
                            cells.Orientation = Excel.XlOrientation.xlHorizontal;
                            break;
                        case 3:
                            cells.Cells[1, 1] = "Ставка";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 4:
                            cells.Cells[1, 1] = "Ученая степень, звание";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 5:
                            cells.Cells[1, 1] = "Лекции по семестрам";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 6:
                            cells.Cells[1, 1] = "Всего лекций";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 7:
                            cells.Cells[1, 1] = "Экзамены";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 8:
                            cells.Cells[1, 1] = "Зачеты";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 9:
                            cells.Cells[1, 1] = "ПК";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;                        
                        case 10:
                            cells.Cells[1, 1] = "Консультации";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 11:
                            cells.Cells[1, 1] = "Практические занятия";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 12:
                            cells.Cells[1, 1] = "Домашние задания и рефераты";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 13:
                            cells.Cells[1, 1] = "Текущая аттестация";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 14:
                            cells.Cells[1, 1] = "Индивидуальные занятия";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 15:
                            cells.Cells[1, 1] = "Контрольные работы";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 16:
                            cells.Cells[1, 1] = "Курсовой проект, курсовая работа";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 17:
                            cells.Cells[1, 1] = "Дипломный проект";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 18:
                            cells.Cells[1, 1] = "Учебная практика";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 19:
                            cells.Cells[1, 1] = "Преддипломная и производственная практика";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 20:
                            cells.Cells[1, 1] = "ГЭК";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 21:
                            cells.Cells[1, 1] = "Приёмная комиссия";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 22:
                            cells.Cells[1, 1] = "Лабораторные работы";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 23:
                            cells.Cells[1, 1] = "Аспирантура";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 24:
                            cells.Cells[1, 1] = "Посещение занятий";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 25:
                            cells.Cells[1, 1] = "Другие виды занятий";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 26:
                            cells.UnMerge();
                            cells.Borders.Weight = 2;
                            cells.Cells[1, 1] = "I сем.";
                            cells.Cells[2, 1] = "II сем.";
                            cells.Orientation = Excel.XlOrientation.xlHorizontal;
                            break;
                        case 27:
                            cells.Merge();
                            cells.Cells[1, 1] = "Всего за год";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                    }
                }

                //-----------Формируем шапку таблицы

                //-----------Заполняем таблицу данными

                curLect = 1;

                for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
                {
                    if ( mdlData.colLecturer[i].FIO.Contains("Лызлов") )
                    {
                        MessageBox.Show("");
                    }

                    flg = false;
                    //Сначала определяем, нужно ли, в принципе, 
                    //рассматривать преподавателя
                    if (optMain.Checked || optCombine.Checked)
                    {
                        accessFlg = (mdlData.colLecturer[i].Rate > 0);
                    }
                    else
                    {
                        if (optHoured.Checked || optMainDop.Checked || optCombineDop.Checked)
                        {
                            accessFlg = true;
                        }
                    }

                    //Если прошёл отбор
                    if (accessFlg)
                    {
                        sumLecI = 0;    sumLecII = 0;

                        sumExamI = 0;   sumExamII = 0;

                        sumCredI = 0;   sumCredII = 0;

                        sumTutI = 0;    sumTutII = 0;

                        sumLabI = 0;    sumLabII = 0;

                        sumPracI = 0;   sumPracII = 0;

                        sumRefI = 0;    sumRefII = 0;

                        sumIndI = 0;    sumIndII = 0;

                        sumKRAPKI = 0;  sumKRAPKII = 0;

                        sumKursPrI = 0; sumKursPrII = 0;

                        sumDiplI = 0;   sumDiplII = 0;

                        sumTutPrI = 0;  sumTutPrII = 0;

                        sumPreDipI = 0; sumPreDipII = 0;

                        sumGAKI = 0;    sumGAKII = 0;

                        sumPostGrI = 0; sumPostGrII = 0;

                        sumVisI = 0;    sumVisII = 0;

                        sumMagI = 0;    sumMagII = 0;

                        //Просматриваем нагрузку
                        for (int j = 0; j <= coll.Count - 1; j++)
                        {
                            //Если строка не исключена из расчёта нагрузки
                            if (!coll[j].flgExclude)
                            {
                                if (!(coll[j].Lecturer == null))
                                {
                                    if (coll[j].Lecturer.Equals(mdlData.colLecturer[i]))
                                    {
                                        if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                        {
                                            sumLecI += coll[j].Lecture;
                                            sumExamI += coll[j].Exam;
                                            sumCredI += coll[j].Credit;
                                            sumTutI += coll[j].Tutorial;
                                            sumLabI += coll[j].LabWork;
                                            sumPracI += coll[j].Practice;
                                            sumRefI += coll[j].RefHomeWork;
                                            sumIndI += coll[j].IndividualWork;
                                            sumKRAPKI += coll[j].KRAPK;
                                            sumKursPrI += coll[j].KursProject;
                                            sumDiplI += coll[j].DiplomaPaper;
                                            sumTutPrI += coll[j].TutorialPractice;
                                            sumPreDipI += coll[j].PreDiplomaPractice +
                                                coll[j].ProducingPractice;
                                            sumGAKI += coll[j].GAK;
                                            sumPostGrI += coll[j].PostGrad;
                                            sumVisI += coll[j].Visiting;
                                            sumMagI += coll[j].Magistry;
                                        }

                                        if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                        {
                                            sumLecII += coll[j].Lecture;
                                            sumExamII += coll[j].Exam;
                                            sumCredII += coll[j].Credit;
                                            sumTutII += coll[j].Tutorial;
                                            sumLabII += coll[j].LabWork;
                                            sumPracII += coll[j].Practice;
                                            sumRefII += coll[j].RefHomeWork;
                                            sumIndII += coll[j].IndividualWork;
                                            sumKRAPKII += coll[j].KRAPK;
                                            sumKursPrII += coll[j].KursProject;
                                            sumDiplII += coll[j].DiplomaPaper;
                                            sumTutPrII += coll[j].TutorialPractice;
                                            sumPreDipII += coll[j].PreDiplomaPractice +
                                                coll[j].ProducingPractice;
                                            sumGAKII += coll[j].GAK;
                                            sumPostGrII += coll[j].PostGrad;
                                            sumVisII += coll[j].Visiting;
                                            sumMagII += coll[j].Magistry;
                                        }

                                        if (optHoured.Checked)
                                        {
                                            if (!flg)
                                            {
                                                flg = (mdlData.toSumDistributionComponents(coll[j]) != 0);
                                            }
                                        }
                                    }
                                }
                                //Если преподаватель не указан
                                else
                                {
                                    //а нагрузка равномерно распределяемая
                                    if (coll[j].flgDistrib)
                                    {
                                        countWeight = 0;
                                        countStud = 0;

                                        for (int k = 0; k <= mdlData.colStudents.Count - 1; k++)
                                        {
                                            if (mdlData.colStudents[k].flgPlan)
                                            {
                                                //Если рассматриваемый преподаватель - руководитель студента
                                                //И если студент на том же курсе, где и дисциплина
                                                //И специальность студента должна соответствовать специальности нагрузки
                                                if (mdlData.colStudents[k].Lect.Equals(mdlData.colLecturer[i])
                                                    & mdlData.colStudents[k].KursNum.Equals(coll[j].KursNum)
                                                    & mdlData.colStudents[k].Speciality.Equals(coll[j].Speciality))
                                                {
                                                    countWeight += coll[j].Weight;
                                                    countStud++;
                                                }
                                            }
                                        }

                                        //
                                        if (flgCombine)
                                        {
                                            mdlData.toDetectUniformInHoured(ref countWeight, coll[j], mdlData.colLecturer[i]);
                                        }

                                        //
                                        if (countWeight > 0)
                                        {
                                            if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                            {
                                                if (coll[j].PreDiplomaPractice > 0 ||
                                                    coll[j].ProducingPractice > 0)
                                                {
                                                    sumPreDipI += countWeight;
                                                }

                                                if (coll[j].TutorialPractice > 0)
                                                {
                                                    sumTutPrI += countWeight;
                                                }

                                                if (coll[j].DiplomaPaper > 0)
                                                {
                                                    sumDiplI += countWeight;
                                                }

                                                if (coll[j].Magistry > 0)
                                                {
                                                    sumMagI += countWeight;
                                                }
                                            }

                                            if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                            {
                                                if (coll[j].PreDiplomaPractice > 0 ||
                                                    coll[j].ProducingPractice > 0)
                                                {
                                                    sumPreDipII += countWeight;
                                                }

                                                if (coll[j].TutorialPractice > 0)
                                                {
                                                    sumTutPrII += countWeight;
                                                }

                                                if (coll[j].DiplomaPaper > 0)
                                                {
                                                    sumDiplII += countWeight;
                                                }

                                                if (coll[j].Magistry > 0)
                                                {
                                                    sumMagII += countWeight;
                                                }
                                            }

                                            if (optHoured.Checked)
                                            {
                                                flg = true;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (optHoured.Checked)
                        {
                            accessFlg = flg;
                        }

                        if (accessFlg)
                        {
                            strSemI = "=";
                            strSemII = "=";
                            for (int k = 5; k <= 25; k++)
                            {
                                if (k != 6)
                                {
                                    strSemI += mdlData.ExcelCellTranslator(6 + 2 * (curLect - 1), k) + "+";
                                    strSemII += mdlData.ExcelCellTranslator(7 + 2 * (curLect - 1), k) + "+";
                                }
                            }
                            strSemI = strSemI.Substring(0, strSemI.Length - 1);
                            strSemII = strSemII.Substring(0, strSemII.Length - 1);

                            for (int k = 1; k <= 30; k++)
                            {
                                //Выбираем диапазон
                                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(6 + 2 * (curLect - 1), k),
                                    mdlData.ExcelCellTranslator(7 + 2 * (curLect - 1), k));
                                //Задаём границы
                                cells.Borders.Weight = 2;

                                switch (k)
                                {
                                    //Номер по порядку
                                    case 1:
                                        cells.Merge();
                                        cells.Cells[1, 1] = curLect.ToString();
                                        //Горизонтальное выравнивание по центру в ячейках
                                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                        //Вертикальное выравнивание по центру в ячейках
                                        cells.VerticalAlignment = Excel.Constants.xlBottom;
                                        break;
                                    //Фамилия, Имя, Отчество
                                    case 2:
                                        cells.Merge();
                                        cells.Cells[1, 1] = mdlData.colLecturer[i].FIO;
                                        //Горизонтальное выравнивание по центру в ячейках
                                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                        //Вертикальное выравнивание по центру в ячейках
                                        cells.VerticalAlignment = Excel.Constants.xlCenter;
                                        cells.WrapText = true;
                                        break;
                                    //Ставка
                                    case 3:
                                        cells.Merge();
                                        cells.Cells[1, 1] = mdlData.colLecturer[i].Rate.ToString("0.00");
                                        //Горизонтальное выравнивание по центру в ячейках
                                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                        //Вертикальное выравнивание по центру в ячейках
                                        cells.VerticalAlignment = Excel.Constants.xlCenter;
                                        cells.WrapText = true;
                                        break;
                                    //Учёная степень, звание
                                    case 4:
                                        cells.Merge();
                                        cells.Cells[1, 1] = mdlData.colLecturer[i].Degree.Short + ", " +
                                            mdlData.colLecturer[i].Duty.Short;
                                        //Горизонтальное выравнивание по центру в ячейках
                                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                        //Вертикальное выравнивание по центру в ячейках
                                        cells.VerticalAlignment = Excel.Constants.xlCenter;
                                        cells.WrapText = true;
                                        break;
                                    //Лекции по семестрам
                                    case 5:
                                        cells.Cells[1, 1] = sumLecI.ToString();
                                        cells.Cells[2, 1] = sumLecII.ToString();
                                        break;
                                    //Всего лекций
                                    case 6:
                                        cells.Merge();
                                        //Горизонтальное выравнивание по центру в ячейках
                                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                        //Вертикальное выравнивание по центру в ячейках
                                        cells.VerticalAlignment = Excel.Constants.xlCenter;
                                        cells.Cells[1, 1] = (sumLecI + sumLecII).ToString();
                                        break;
                                    //Экзамены
                                    case 7:
                                        cells.Cells[1, 1] = sumExamI.ToString();
                                        cells.Cells[2, 1] = sumExamII.ToString();
                                        break;
                                    //Зачёты
                                    case 8:
                                        cells.Cells[1, 1] = sumCredI.ToString();
                                        cells.Cells[2, 1] = sumCredII.ToString();
                                        break;
                                    //ПК
                                    case 9:
                                        cells.Cells[1, 1] = 0.ToString();
                                        cells.Cells[2, 1] = 0.ToString();
                                        break;
                                    //Консультации
                                    case 10:
                                        cells.Cells[1, 1] = sumTutI.ToString();
                                        cells.Cells[2, 1] = sumTutII.ToString();
                                        break;
                                    //Практические занятия
                                    case 11:
                                        cells.Cells[1, 1] = sumPracI.ToString();
                                        cells.Cells[2, 1] = sumPracII.ToString();
                                        break;
                                    //Домашние задания и рефераты
                                    case 12:
                                        cells.Cells[1, 1] = sumRefI.ToString();
                                        cells.Cells[2, 1] = sumRefII.ToString();
                                        break;
                                    //Текущая аттестация
                                    case 13:
                                        cells.Cells[1, 1] = 0.ToString();
                                        cells.Cells[2, 1] = 0.ToString();
                                        break;
                                    //Индивидуальные занятия
                                    case 14:
                                        cells.Cells[1, 1] = sumIndI.ToString();
                                        cells.Cells[2, 1] = sumIndII.ToString();
                                        break;
                                    //Контрольные работы
                                    case 15:
                                        cells.Cells[1, 1] = sumKRAPKI.ToString();
                                        cells.Cells[2, 1] = sumKRAPKII.ToString();
                                        break;
                                    //Курсовой проект, курсовая работа
                                    case 16:
                                        cells.Cells[1, 1] = sumKursPrI.ToString();
                                        cells.Cells[2, 1] = sumKursPrII.ToString();
                                        break;
                                    //Дипломный проект
                                    case 17:
                                        cells.Cells[1, 1] = sumDiplI.ToString();
                                        cells.Cells[2, 1] = sumDiplII.ToString();
                                        break;
                                    //Учебная практика
                                    case 18:
                                        cells.Cells[1, 1] = sumTutPrI.ToString();
                                        cells.Cells[2, 1] = sumTutPrII.ToString();
                                        break;
                                    //Преддипломная и производственная практика
                                    case 19:
                                        cells.Cells[1, 1] = sumPreDipI.ToString();
                                        cells.Cells[2, 1] = sumPreDipII.ToString();
                                        break;
                                    //ГЭК
                                    case 20:
                                        cells.Cells[1, 1] = sumGAKI.ToString();
                                        cells.Cells[2, 1] = sumGAKII.ToString();
                                        break;
                                    //Приёмная комиссия
                                    case 21:
                                        cells.Cells[1, 1] = 0.ToString();
                                        cells.Cells[2, 1] = 0.ToString();
                                        break;
                                    //Лабораторные работы
                                    case 22:
                                        cells.Cells[1, 1] = sumLabI.ToString();
                                        cells.Cells[2, 1] = sumLabII.ToString();
                                        break;
                                    //Аспирантура
                                    case 23:
                                        cells.Cells[1, 1] = sumPostGrI.ToString();
                                        cells.Cells[2, 1] = sumPostGrII.ToString();
                                        break;
                                    //Посещение занятий
                                    case 24:
                                        cells.Cells[1, 1] = sumVisI.ToString();
                                        cells.Cells[2, 1] = sumVisII.ToString();
                                        break;
                                    //Другие виды занятий
                                    case 25:
                                        cells.Cells[1, 1] = sumMagI.ToString();
                                        cells.Cells[2, 1] = sumMagII.ToString();
                                        break;
                                    case 26:
                                        cells.Cells[1, 1] = strSemI;
                                        cells.Cells[2, 1] = strSemII;
                                        break;
                                    case 27:
                                        cells.Merge();
                                        //Горизонтальное выравнивание по центру в ячейках
                                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                        //Вертикальное выравнивание по центру в ячейках
                                        cells.VerticalAlignment = Excel.Constants.xlCenter;
                                        cells.Cells[1, 1] = "=" + mdlData.ExcelCellTranslator(6 + 2 * (curLect - 1), k - 1) + "+" +
                                            mdlData.ExcelCellTranslator(7 + 2 * (curLect - 1), k - 1);
                                        break;
                                    case 28:
                                        cells.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                                        break;
                                    case 29:
                                        //Отменяем границы
                                        cells.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                                        cells.Cells[1, 1] = mdlData.colLecturer[i].Rate.ToString("0.00");
                                        break;
                                    case 30:
                                        //Отменяем границы
                                        cells.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                                        cells.Cells[1, 1] = mdlData.colLecturer[i].Degree.Short + ", " +
                                            mdlData.colLecturer[i].Duty.Short;
                                        break;
                                }
                            }

                            strSumI += mdlData.ExcelCellTranslator(6 + 2 * (curLect - 1), 26) + "+";
                            strSumII += mdlData.ExcelCellTranslator(7 + 2 * (curLect - 1), 26) + "+";
                            strSum += mdlData.ExcelCellTranslator(6 + 2 * (curLect - 1), 27) + "+";

                            curLect += 1;
                        }
                    }
                }

                strSumI = strSumI.Substring(0, strSumI.Length - 1);
                strSumII = strSumII.Substring(0, strSumII.Length - 1);
                strSum = strSum.Substring(0, strSum.Length - 1);

                //Объединённая итоговая сумма часов
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (2 * curLect), 27),
                    mdlData.ExcelCellTranslator(5 + (2 * curLect), 27));
                cells.Merge();
                //Горизонтальное выравнивание по центру в ячейках
                cells.HorizontalAlignment = Excel.Constants.xlCenter;
                //Вертикальное выравнивание по центру в ячейках
                cells.VerticalAlignment = Excel.Constants.xlCenter;
                //Задаём границы
                cells.Borders.Weight = 2;
                cells.Cells[1, 1] = strSum;

                //Итоговая сумма часов по семестрам (3 строки)
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (2 * curLect), 26),
                    mdlData.ExcelCellTranslator(5 + (2 * curLect) + 1, 26));
                cells.Cells[1, 1] = strSumI;
                cells.Cells[2, 1] = strSumII;
                //Задаём границы
                cells.Borders.Weight = 2;
                cells.Cells[3, 1] = "=" + mdlData.ExcelCellTranslator(4 + (2 * curLect), 26) + "+" +
                    mdlData.ExcelCellTranslator(5 + (2 * curLect), 26);

                //Надписи семестров по итогам
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (2 * curLect), 25),
                    mdlData.ExcelCellTranslator(5 + (2 * curLect) + 1, 25));
                cells.Cells[1, 1] = "I сем.";
                cells.Cells[2, 1] = "II сем.";

                //Надпись итого
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (2 * curLect), 24),
                    mdlData.ExcelCellTranslator(5 + (2 * curLect), 24));
                cells.Cells[1, 1] = "Итого:";

                //Суммируем ставки преподавателей
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (2 * curLect) + 1, 2),
                    mdlData.ExcelCellTranslator(5 + (2 * curLect) + 7, 3));
                cells.Cells[1, 1] = "Сумма ставок:";
                cells.Cells[1, 2] = sumRate.ToString();
                cells.Cells[2, 1] = "Сумма асс.:";
                cells.Cells[2, 2] = assist.ToString();
                cells.Cells[3, 1] = "Сумма ст.преп.:";
                cells.Cells[3, 2] = hitutor.ToString();
                cells.Cells[4, 1] = "Сумма доц.:";
                cells.Cells[4, 2] = lecturer.ToString();
                cells.Cells[5, 1] = "Сумма проф.:";
                cells.Cells[5, 2] = proff.ToString();
                cells.Cells[6, 2] = sumRate.ToString();

                //Подпись заведующего кафедрой
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(4 + (2 * curLect) + 3, 12),
                    mdlData.ExcelCellTranslator(5 + (2 * curLect) + 3, 12));
                cells.Cells[1, 1] = "Заведующий кафедрой                                                                                       / Л.А. Баранов /";

                //-----------Формируем таблицу даными

                ObjExcel.UserControl = true;

                ObjWorkBook.SaveCopyAs(Application.StartupPath + @"\Ведомость плановая " + type + " " +
                    DateTime.Now.Year.ToString() + DateTime.Now.Year.ToString() + DateTime.Now.Day.ToString() +
                    DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + 
                    DateTime.Now.Second.ToString() + ".xlsx");

                ObjWorkBook.Close(false, "", Missing.Value);

                ObjExcel.Quit();
            }
            catch
            {
                MessageBox.Show("Возможно на этом компьютере присутствует проблема совместимости с MS Exсel." +
                                " Попробуйте установить версию 2007 и выше.");
            }
        }

        private void intoExcel()
        {
            int curLect;

            int sumLecI;
            int sumLecII;

            int sumExamI;
            int sumExamII;

            int sumCredI;
            int sumCredII;

            int sumTutI;
            int sumTutII;

            int sumLabI;
            int sumLabII;

            int sumPracI;
            int sumPracII;

            int sumRefI;
            int sumRefII;

            int sumIndI;
            int sumIndII;

            int sumKRAPKI;
            int sumKRAPKII;

            int sumKursPrI;
            int sumKursPrII;

            int sumDiplI;
            int sumDiplII;

            int sumTutPrI;
            int sumTutPrII;

            int sumPreDipI;
            int sumPreDipII;

            int sumGAKI;
            int sumGAKII;

            int sumPostGrI;
            int sumPostGrII;

            int sumVisI;
            int sumVisII;

            int sumMagI;
            int sumMagII;

            int sumI;
            int sumII;

            int sumAllI;
            int sumAllII;

            int CheckSum;

            double[] rateParams = new double[4];
            double assist;
            double hitutor;
            double proff;
            double lecturer;
            double sumRate;

            IList<clsDistribution> coll = null;
            bool accessFlg = false;

            rateParams = countRates();

            assist = rateParams[0];
            lecturer = rateParams[1];
            hitutor = rateParams[2];
            proff = rateParams[3];

            sumRate = proff + hitutor + lecturer + assist;

            if (optMain.Checked || optMainDop.Checked)
            {
                coll = mdlData.colDistribution;
            }
            else
            {
                if (optHoured.Checked)
                {
                    coll = mdlData.colHouredDistribution;
                }
                else
                {
                    if (optCombine.Checked || optCombineDop.Checked)
                    {
                        coll = mdlData.colCombineDistribution;
                    }
                }
            }

            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;

            try
            {
                //Создаём новое Excel приложение
                Excel.Application ObjExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook;
                Excel.Worksheet ObjWorkSheet;

                ObjExcel.Visible = true;

                //Книга
                ObjWorkBook = ObjExcel.Workbooks.Add(Missing.Value);
                //Таблица
                ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
                //Альбомная ориентация страницы
                ObjWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                //Книжная ориентация страницы
                //ObjWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                //В высоту разместить на одной странице
                ObjWorkSheet.PageSetup.FitToPagesTall = 1;
                //В ширину разместить на одной странице
                ObjWorkSheet.PageSetup.FitToPagesWide = 1;

                //-----------Формируем заголовок таблицы

                //Задаём диапазон для ячеек, подлежащих форматированию
                //1-я строка, с А по Z
                var cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(1, 1),
                    mdlData.ExcelCellTranslator(1, 26));
                //Выделяем ячейки диапазона
                cells.Select();
                //Объединяем ячейки диапазона
                cells.Merge(true);
                //Выравнивание в объединённой ячейке по центру
                cells.HorizontalAlignment = Excel.Constants.xlCenter;
                //Записываем текст в объединённую ячейку
                //(считается по первой А1)
                cells.Cells[1, 1] = "Кафедра \"Управление и защита информации\"";

                //-----------Формируем заголовок таблицы

                //-----------Формируем шапку таблицы

                for (int i = 1; i <= 26; i++)
                {
                    //Выбираем диапазон
                    cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(3, i),
                        mdlData.ExcelCellTranslator(4, i));
                    //Выделяем ячейки диапазона
                    cells.Select();
                    //Объединяем ячейки диапазона
                    cells.Merge();
                    //Горизонтальное выравнивание по центру в ячейках
                    cells.HorizontalAlignment = Excel.Constants.xlCenter;
                    //Вертикальное выравнивание по центру в ячейках
                    cells.VerticalAlignment = Excel.Constants.xlCenter;
                    //Задаём границы
                    cells.Borders.Weight = 2;
                    //Перенос по словам без сокрытия под ячейками
                    cells.WrapText = true;
                    //Назначаем высоту шапки
                    cells.RowHeight = (68.25f + 66) / 2;

                    switch (i)
                    {
                        case 1:
                            cells.Cells[1, 1] = "№ п/п";
                            cells.ColumnWidth = 3.86f;
                            cells.Orientation = Excel.XlOrientation.xlHorizontal;
                            break;
                        case 2:
                            cells.Cells[1, 1] = "Фамилия, Имя, Отчество";
                            cells.ColumnWidth = 27.57f;
                            cells.Orientation = Excel.XlOrientation.xlHorizontal;
                            break;
                        case 3:
                            cells.Cells[1, 1] = "Ученая степень, звание";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 4:
                            cells.Cells[1, 1] = "Лекции по семестрам";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 5:
                            cells.Cells[1, 1] = "Всего лекций";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 6:
                            cells.Cells[1, 1] = "Экзамены";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 7:
                            cells.Cells[1, 1] = "Зачеты";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 8:
                            cells.Cells[1, 1] = "Консультации";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 9:
                            cells.Cells[1, 1] = "Лабораторные работы";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 10:
                            cells.Cells[1, 1] = "Практические занятия";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 11:
                            cells.Cells[1, 1] = "Домашние задания и рефераты";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 12:
                            cells.Cells[1, 1] = "Текущая аттестация";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 13:
                            cells.Cells[1, 1] = "Индивидуальные занятия";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 14:
                            cells.Cells[1, 1] = "Контрольные работы";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 15:
                            cells.Cells[1, 1] = "Курсовой проект, курсовая работа";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 16:
                            cells.Cells[1, 1] = "Дипломный проект";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 17:
                            cells.Cells[1, 1] = "Учебная практика";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 18:
                            cells.Cells[1, 1] = "Преддипломная и производственная практика";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 19:
                            cells.Cells[1, 1] = "ГЭК";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 20:
                            cells.Cells[1, 1] = "Приёмная комиссия";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 21:
                            cells.Cells[1, 1] = "ФПК";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 22:
                            cells.Cells[1, 1] = "Аспирантура";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 23:
                            cells.Cells[1, 1] = "Посещение занятий";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 24:
                            cells.Cells[1, 1] = "Руководство магистерской программой";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                        case 25:
                            cells.UnMerge();
                            cells.Borders.Weight = 2;
                            cells.Cells[1, 1] = "I сем.";
                            cells.Cells[2, 1] = "II сем.";
                            cells.Orientation = Excel.XlOrientation.xlHorizontal;
                            break;
                        case 26:
                            cells.Merge();
                            cells.Cells[1, 1] = "Всего за год";
                            cells.Orientation = Excel.XlOrientation.xlUpward;
                            break;
                    }
                }

                //-----------Формируем шапку таблицы

                //-----------Заполняем таблицу данными

                sumAllI = 0;
                sumAllII = 0;
                CheckSum = 0;
                curLect = 1;

                for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
                {
                    //Сначала определяем, нужно ли, в принципе, 
                    //рассматривать преподавателя
                    if (optMain.Checked || optCombine.Checked)
                    {
                        accessFlg = (mdlData.colLecturer[i].Rate > 0);
                    }
                    else
                    {
                        if (optHoured.Checked || optMainDop.Checked || optCombineDop.Checked)
                        {
                            accessFlg = true;
                        }
                    }

                    //Если прошёл отбор
                    if (accessFlg)
                    {
                        sumLecI = 0;
                        sumLecII = 0;

                        sumExamI = 0;
                        sumExamII = 0;

                        sumCredI = 0;
                        sumCredII = 0;

                        sumTutI = 0;
                        sumTutII = 0;

                        sumLabI = 0;
                        sumLabII = 0;

                        sumPracI = 0;
                        sumPracII = 0;

                        sumRefI = 0;
                        sumRefII = 0;

                        sumIndI = 0;
                        sumIndII = 0;

                        sumKRAPKI = 0;
                        sumKRAPKII = 0;

                        sumKursPrI = 0;
                        sumKursPrII = 0;

                        sumDiplI = 0;
                        sumDiplII = 0;

                        sumTutPrI = 0;
                        sumTutPrII = 0;

                        sumPreDipI = 0;
                        sumPreDipII = 0;

                        sumGAKI = 0;
                        sumGAKII = 0;

                        sumPostGrI = 0;
                        sumPostGrII = 0;

                        sumVisI = 0;
                        sumVisII = 0;

                        sumMagI = 0;
                        sumMagII = 0;

                        sumI = 0;
                        sumII = 0;

                        //Просматриваем нагрузку
                        for (int j = 0; j <= coll.Count - 1; j++)
                        {
                            if (!(coll[j].Lecturer == null))
                            {
                                if (coll[j].Lecturer.Equals(mdlData.colLecturer[i]))
                                {
                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        sumLecI += coll[j].Lecture;
                                        sumI += coll[j].Lecture;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        sumLecII += coll[j].Lecture;
                                        sumII += coll[j].Lecture;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        sumExamI += coll[j].Exam;
                                        sumI += coll[j].Exam;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        sumExamII += coll[j].Exam;
                                        sumII += coll[j].Exam;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        sumCredI += coll[j].Credit;
                                        sumI += coll[j].Credit;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        sumCredII += coll[j].Credit;
                                        sumII += coll[j].Credit;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        sumTutI += coll[j].Tutorial;
                                        sumI += coll[j].Tutorial;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        sumTutII += coll[j].Tutorial;
                                        sumII += coll[j].Tutorial;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        sumLabI += coll[j].LabWork;
                                        sumI += coll[j].LabWork;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        sumLabII += coll[j].LabWork;
                                        sumII += coll[j].LabWork;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        sumPracI += coll[j].Practice;
                                        sumI += coll[j].Practice;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        sumPracII += coll[j].Practice;
                                        sumII += coll[j].Practice;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        sumRefI += coll[j].RefHomeWork;
                                        sumI += coll[j].RefHomeWork;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        sumRefII += coll[j].RefHomeWork;
                                        sumII += coll[j].RefHomeWork;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        sumIndI += coll[j].IndividualWork;
                                        sumI += coll[j].IndividualWork;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        sumIndII += coll[j].IndividualWork;
                                        sumII += coll[j].IndividualWork;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        sumKRAPKI += coll[j].KRAPK;
                                        sumI += coll[j].KRAPK;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        sumKRAPKII += coll[j].KRAPK;
                                        sumII += coll[j].KRAPK;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        sumKursPrI += coll[j].KursProject;
                                        sumI += coll[j].KursProject;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        sumKursPrII += coll[j].KursProject;
                                        sumII += coll[j].KursProject;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        sumDiplI += coll[j].DiplomaPaper;
                                        sumI += coll[j].DiplomaPaper;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        sumDiplII += coll[j].DiplomaPaper;
                                        sumII += coll[j].DiplomaPaper;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        sumTutPrI += coll[j].TutorialPractice;
                                        sumI += coll[j].TutorialPractice;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        sumTutPrII += coll[j].TutorialPractice;
                                        sumII += coll[j].TutorialPractice;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        sumPreDipI += coll[j].PreDiplomaPractice +
                                            coll[j].ProducingPractice;
                                        sumI += coll[j].PreDiplomaPractice +
                                            coll[j].ProducingPractice;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        sumPreDipII += coll[j].PreDiplomaPractice +
                                            coll[j].ProducingPractice;
                                        sumII += coll[j].PreDiplomaPractice +
                                            coll[j].ProducingPractice;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        sumGAKI += coll[j].GAK;
                                        sumI += coll[j].GAK;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        sumGAKII += coll[j].GAK;
                                        sumII += coll[j].GAK;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        sumPostGrI += coll[j].PostGrad;
                                        sumI += coll[j].PostGrad;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        sumPostGrII += coll[j].PostGrad;
                                        sumII += coll[j].PostGrad;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        sumVisI += coll[j].Visiting;
                                        sumI += coll[j].Visiting;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        sumVisII += coll[j].Visiting;
                                        sumII += coll[j].Visiting;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        sumMagI += coll[j].Magistry;
                                        sumI += coll[j].Magistry;
                                    }

                                    if (coll[j].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        sumMagII += coll[j].Magistry;
                                        sumII += coll[j].Magistry;
                                    }
                                }
                            }
                        }

                        sumAllI += sumI;
                        sumAllII += sumII;

                        if (optHoured.Checked || optMainDop.Checked || optCombineDop.Checked)
                        {
                            accessFlg = !(sumI + sumII == 0);
                        }

                        if (accessFlg)
                        {
                            for (int k = 1; k <= 29; k++)
                            {
                                //Выбираем диапазон
                                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(5 + 2 * (curLect - 1), k),
                                    mdlData.ExcelCellTranslator(6 + 2 * (curLect - 1), k));
                                //Задаём границы
                                cells.Borders.Weight = 2;

                                switch (k)
                                {
                                    case 1:
                                        cells.Merge();
                                        cells.Cells[1, 1] = curLect.ToString();
                                        //Горизонтальное выравнивание по центру в ячейках
                                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                        //Вертикальное выравнивание по центру в ячейках
                                        cells.VerticalAlignment = Excel.Constants.xlBottom;
                                        break;
                                    case 2:
                                        cells.Merge();
                                        cells.Cells[1, 1] = mdlData.colLecturer[i].FIO + " (" +
                                            mdlData.colLecturer[i].Rate.ToString("0.00") + " ставки)";
                                        //Горизонтальное выравнивание по центру в ячейках
                                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                        //Вертикальное выравнивание по центру в ячейках
                                        cells.VerticalAlignment = Excel.Constants.xlCenter;
                                        cells.WrapText = true;
                                        break;
                                    case 3:
                                        cells.Merge();
                                        cells.Cells[1, 1] = mdlData.colLecturer[i].Degree.Short + ", " +
                                            mdlData.colLecturer[i].Duty.Short;
                                        //Горизонтальное выравнивание по центру в ячейках
                                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                        //Вертикальное выравнивание по центру в ячейках
                                        cells.VerticalAlignment = Excel.Constants.xlCenter;
                                        cells.WrapText = true;
                                        break;
                                    case 4:
                                        cells.Cells[1, 1] = sumLecI.ToString();
                                        cells.Cells[2, 1] = sumLecII.ToString();
                                        break;
                                    case 5:
                                        cells.Merge();
                                        //Горизонтальное выравнивание по центру в ячейках
                                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                        //Вертикальное выравнивание по центру в ячейках
                                        cells.VerticalAlignment = Excel.Constants.xlCenter;
                                        cells.Cells[1, 1] = (sumLecI + sumLecII).ToString();
                                        break;
                                    case 6:
                                        cells.Cells[1, 1] = sumExamI.ToString();
                                        cells.Cells[2, 1] = sumExamII.ToString();
                                        break;
                                    case 7:
                                        cells.Cells[1, 1] = sumCredI.ToString();
                                        cells.Cells[2, 1] = sumCredII.ToString();
                                        break;
                                    case 8:
                                        cells.Cells[1, 1] = sumTutI.ToString();
                                        cells.Cells[2, 1] = sumTutII.ToString();
                                        break;
                                    case 9:
                                        cells.Cells[1, 1] = sumLabI.ToString();
                                        cells.Cells[2, 1] = sumLabII.ToString();
                                        break;
                                    case 10:
                                        cells.Cells[1, 1] = sumPracI.ToString();
                                        cells.Cells[2, 1] = sumPracII.ToString();
                                        break;
                                    case 11:
                                        cells.Cells[1, 1] = sumRefI.ToString();
                                        cells.Cells[2, 1] = sumRefII.ToString();
                                        break;
                                    case 12:
                                        cells.Cells[1, 1] = 0.ToString();
                                        cells.Cells[2, 1] = 0.ToString();
                                        break;
                                    case 13:
                                        cells.Cells[1, 1] = sumIndI.ToString();
                                        cells.Cells[2, 1] = sumIndII.ToString();
                                        break;
                                    case 14:
                                        cells.Cells[1, 1] = sumKRAPKI.ToString();
                                        cells.Cells[2, 1] = sumKRAPKII.ToString();
                                        break;
                                    case 15:
                                        cells.Cells[1, 1] = sumKursPrI.ToString();
                                        cells.Cells[2, 1] = sumKursPrII.ToString();
                                        break;
                                    case 16:
                                        cells.Cells[1, 1] = sumDiplI.ToString();
                                        cells.Cells[2, 1] = sumDiplII.ToString();
                                        break;
                                    case 17:
                                        cells.Cells[1, 1] = sumTutPrI.ToString();
                                        cells.Cells[2, 1] = sumTutPrII.ToString();
                                        break;
                                    case 18:
                                        cells.Cells[1, 1] = sumPreDipI.ToString();
                                        cells.Cells[2, 1] = sumPreDipII.ToString();
                                        break;
                                    case 19:
                                        cells.Cells[1, 1] = sumGAKI.ToString();
                                        cells.Cells[2, 1] = sumGAKII.ToString();
                                        break;
                                    case 20:
                                        cells.Cells[1, 1] = 0.ToString();
                                        cells.Cells[2, 1] = 0.ToString();
                                        break;
                                    case 21:
                                        cells.Cells[1, 1] = 0.ToString();
                                        cells.Cells[2, 1] = 0.ToString();
                                        break;
                                    case 22:
                                        cells.Cells[1, 1] = sumPostGrI.ToString();
                                        cells.Cells[2, 1] = sumPostGrII.ToString();
                                        break;
                                    case 23:
                                        cells.Cells[1, 1] = sumVisI.ToString();
                                        cells.Cells[2, 1] = sumVisII.ToString();
                                        break;
                                    case 24:
                                        cells.Cells[1, 1] = sumMagI.ToString();
                                        cells.Cells[2, 1] = sumMagII.ToString();
                                        break;
                                    case 25:
                                        cells.Cells[1, 1] = sumI.ToString();
                                        cells.Cells[2, 1] = sumII.ToString();
                                        break;
                                    case 26:
                                        cells.Merge();
                                        //Горизонтальное выравнивание по центру в ячейках
                                        cells.HorizontalAlignment = Excel.Constants.xlCenter;
                                        //Вертикальное выравнивание по центру в ячейках
                                        cells.VerticalAlignment = Excel.Constants.xlCenter;
                                        cells.Cells[1, 1] = (sumI + sumII).ToString();
                                        break;
                                    case 27:
                                        cells.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                                        break;
                                    case 28:
                                        //Отменяем границы
                                        cells.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                                        cells.Cells[1, 1] = mdlData.colLecturer[i].Rate.ToString("0.00");
                                        break;
                                    case 29:
                                        //Отменяем границы
                                        cells.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                                        cells.Cells[1, 1] = mdlData.colLecturer[i].Degree.Short + ", " +
                                            mdlData.colLecturer[i].Duty.Short;
                                        break;
                                }
                            }

                            CheckSum += (sumI + sumII);

                            curLect += 1;
                        }
                    }
                }

                //Объединённая итоговая сумма часов
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(3 + (2 * curLect), 26),
                    mdlData.ExcelCellTranslator(3 + (2 * curLect) + 1, 26));
                //cells.Select();
                cells.Merge();
                //Горизонтальное выравнивание по центру в ячейках
                cells.HorizontalAlignment = Excel.Constants.xlCenter;
                //Вертикальное выравнивание по центру в ячейках
                cells.VerticalAlignment = Excel.Constants.xlCenter;
                //Задаём границы
                cells.Borders.Weight = 2;
                cells.Cells[1, 1] = CheckSum.ToString();

                //Итоговая сумма часов по семестрам (3 строки)
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(3 + (2 * curLect), 25),
                    mdlData.ExcelCellTranslator(3 + (2 * curLect) + 2, 25));
                cells.Cells[1, 1] = sumAllI.ToString();
                cells.Cells[2, 1] = sumAllII.ToString();
                //Задаём границы
                cells.Borders.Weight = 2;
                cells.Cells[3, 1] = (sumAllI + sumAllII).ToString();

                //Надписи семестров по итогам
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(3 + (2 * curLect), 24),
                    mdlData.ExcelCellTranslator(3 + (2 * curLect) + 1, 24));
                cells.Cells[1, 1] = "I сем.";
                cells.Cells[2, 1] = "II сем.";

                //Надпись итого
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(3 + (2 * curLect), 23),
                    mdlData.ExcelCellTranslator(3 + (2 * curLect), 23));
                cells.Cells[1, 1] = "Итого:";

                //Суммируем ставки преподавателей
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(3 + (2 * curLect) + 1, 2),
                    mdlData.ExcelCellTranslator(3 + (2 * curLect) + 7, 3));
                cells.Cells[1, 1] = "Сумма ставок:";
                cells.Cells[1, 2] = sumRate.ToString();
                cells.Cells[2, 1] = "Сумма асс.:";
                cells.Cells[2, 2] = assist.ToString();
                cells.Cells[3, 1] = "Сумма ст.преп.:";
                cells.Cells[3, 2] = hitutor.ToString();
                cells.Cells[4, 1] = "Сумма доц.:";
                cells.Cells[4, 2] = lecturer.ToString();
                cells.Cells[5, 1] = "Сумма проф.:";
                cells.Cells[5, 2] = proff.ToString();
                cells.Cells[6, 2] = sumRate.ToString();

                //Подпись заведующего кафедрой
                //Выбираем диапазон
                cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(3 + (2 * curLect) + 3, 12),
                    mdlData.ExcelCellTranslator(3 + (2 * curLect) + 3, 12));
                cells.Cells[1, 1] = "Заведующий кафедрой                                                                                       / Л.А. Баранов /";

                //-----------Формируем таблицу даными

                ObjExcel.UserControl = true;

                ObjWorkBook.SaveCopyAs(Application.StartupPath + @"\Ведомость плановая " + 
                    DateTime.Now.Date.ToString("yyyyMMdd") + " " + 
                    DateTime.Now.TimeOfDay.ToString("hhmmss") + ".xlsx");
                ObjWorkBook.Close(false, "", Missing.Value);

                ObjExcel.Quit();
            }
            catch
            {
                MessageBox.Show("Возможно на этом компьютере присутствует проблема совместимости с MS Exсel." +
                                " Попробуйте установить версию 2007 и выше.");
            }
        }

        private void btnOnlyKurs_Click(object sender, EventArgs e)
        {
            KursGrid();
        }

        //
        private void cmbForm_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbForm.SelectedIndex > 0)
            {

            }
            else
            {

            }
        }

        //
        private void cmbGrid_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbGrid.SelectedIndex > 0)
            {
                btnExcel.Enabled = true;

                switch (cmbGrid.SelectedIndex)
                {
                    //не выбрано
                    case 0:
                    {
                        break;
                    }
                    //заявка в диспетчерскую
                    case 1:
                    {
                        DispatchGrid();
                        break;
                    }
                    //распределение для заведующего кафедрой
                    case 2:
                    {
                        BossGrid();
                        break;
                    }
                    //плановое распределение в учебное управление
                    case 3:
                    {
                        break;
                    }
                    //фактическое распределение в учебное управление
                    case 4:
                    {
                        break;
                    }
                    //сведения по курсовым проектам
                    case 5:
                    {
                        KursGrid();
                        break;
                    }
                    //сведения по выпускникам (ВКР)
                    case 6:
                    {
                        LecturerGraduateGrid();
                        break;
                    }
                    //Таблица нагрузки, разгрузки, перегрузки
                    case 7:
                    {
                        LoadUnloadGrid();
                        break;
                    }
                    //Таблица равномерно распределяемой нагрузки
                    case 8:
                    {
                        UniformLoadGrid();
                        break;
                    }
                }
            }
            else
            {
                btnExcel.Enabled = false;
                dgScheduleManagement.Rows.Clear();
                dgScheduleManagement.Columns.Clear();
            }
        }

        //
        private void btnForm_Click(object sender, EventArgs e)
        {
            //Если хотя бы какой-то индекс выбран в комбинированном списке
            if (cmbForm.SelectedIndex > 0)
            {
                //Делаем доступной кнопку "Сформировать"
                btnForm.Enabled = true;
                //В зависимости от индекса готовим определённый документ
                switch (cmbForm.SelectedIndex)
                {
                    //0. Не выбрано
                    //1. Заявка в диспетчерскую Word
                    case 1:
                        {
                            DispatchWord();
                            break;
                        }
                    //2. Плановая нагрузка в управление
                    case 2:
                        {
                            UpravlenieExcel();
                            break;
                        }
                    //3. Выполненная нагрузка в управление
                    case 3:
                        {
                            UpravlenieDoneExcel();
                            break;
                        }
                    //4. Закреплённые дисциплины Excel
                    case 4:
                        {
                            UpravlenieFixedSubj();
                            break;
                        }
                    //5. Режим в Excel
                    case 5:
                        {
                            intoExcel2015();
                            break;
                        }

                    case 6:
                        {
                            break;
                        }
                }
            }
            else
            {
                btnForm.Enabled = false;
            }
        }

        private void dgScheduleManagement_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void optMainDop_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void optCombineDop_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
