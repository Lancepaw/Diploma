using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace DataBaseIO
{
    public partial class frmLecturerID : Form
    {
        public static int fWidth;
        public static int fHeigth;
        
        public frmLecturerID()
        {
            InitializeComponent();

            fWidth = Width;
            fHeigth = Height;
        }

        bool flgCombine = false;

        static string[] NamesDiploma;
        static string[] GroupsDiploma;
        static string[] GroupsFullDiploma;

        static string[] NamesMagistry;
        static string[] GroupsMagistry;
        static string[] GroupsFullMagistry;

        static string[] NamesTutorialPr;
        static string[] GroupsTutorialPr;
        static string[] GroupsFullTutorialPr;

        static string[] NamesProdPr;
        static string[] GroupsProdPr;
        static string[] GroupsFullProdPr;

        //Событие нажатия на кнопку "Создать"
        private void btnAction_Click(object sender, EventArgs e)
        {
            //Выбрать желаемый вид документа, выдаваемый на выход
            switch (cmbForm.SelectedIndex)
            {
                //Сетку индивидуального плана
                case 0:
                    FillGrid();
                    break;
                //Документ Word с индивидуальным планом
                case 1:
                    intoWord();
                    break;
                //Документ Word с почасовым планом
                case 2:
                    intoWordHoured();
                    break;
                //Документы Word с индивидуальными планами
                case 3:
                    allIntoWord();
                    break;
                //Документы Word с почасовыми планами
                case 4:
                    allIntoWordHoured();
                    break;
                //Документ Excel
                case 5:
                    intoExcel();
                    break;
                //Документы Excel
                case 6:
                    allIntoExcel();
                    break;
                //Сетку только с дипломниками
                case 7:
                    FillGridDiplomas();
                    break;
                //Сетку только с магистрантами
                case 8:
                    FillGridMagistry();
                    break;
            }     
        }

        //Событие нажатия на кнопку "Закрыть"
        private void btnClose_Click(object sender, EventArgs e)
        {
            //Закрыть эту форму
            Close();
        }

        //Загрузка формы с планами преподавателей
        private void frmLecturerID_Load(object sender, EventArgs e)
        {
            //Настройка исходного состояния элементов управления
            //при успешной загрузке элементов из таблицы преподавателей
            if (mdlData.colLecturer.Count > 0)
            {
                //Заполняем комбо-боксы
                FillLecturerList();
                FillWorkYearList();

                //0. Нулевой элемент
                cmbForm.Items.Add("Сетку индивидуального плана");
                //1. Первый элемент
                cmbForm.Items.Add("Документ Word с индивидуальным планом");
                //2. Второй элемент
                cmbForm.Items.Add("Документ Word с почасовым планом");
                //3. Третий элемент
                cmbForm.Items.Add("Документы Word с индивидуальными планами");
                //4. Четвёртый элемент
                cmbForm.Items.Add("Документы Word с почасовыми планами");
                //5. Пятый элемент
                cmbForm.Items.Add("Документ Excel");
                //6. Шестой элемент
                cmbForm.Items.Add("Документы Excel");
                //7. Седьмой элемент
                cmbForm.Items.Add("Сетку только с дипломами");
                //8. Восьмой элемент
                cmbForm.Items.Add("Сетку только с магистрами");

                //Выставить комбинированную
                chkCombine.Checked = true;
                //Выставить признак образца после 2015 года
                chkAfter2015.Checked = true;
            }
            //при неудачной загрузке элементов из таблицы должностей
            else
            {

            }

            Resize += new EventHandler(frmLecturerID_Resize);
        }

        void frmLecturerID_Resize(object sender, EventArgs e)
        {
            if (Width >= fWidth & Height >= fHeigth)
            {
                dgNagruzka.Width = Width - 40;

                btnClose.Top = Height - 50 - btnClose.Height;
                btnClose.Left = Width - 30 - btnClose.Width;

                cmbForm.Left = dgNagruzka.Left;
                cmbForm.Top = btnClose.Top;

                btnAction.Left = cmbForm.Left + cmbForm.Width + 10;
                btnAction.Top = cmbForm.Top;

                frParams.Top = cmbForm.Top - 10 - frParams.Height;
                frParams.Left = dgNagruzka.Left;

                dgNagruzka.Height = (frParams.Top - 10) - dgNagruzka.Top;
            }
            else
            {
                Width = fWidth;
                Height = fHeigth;
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

        private void FillLecturerList()
        {
            int NumFix = 0;
            NumFix = cmbLecturerList.SelectedIndex;
            //Очищаем список
            cmbLecturerList.Items.Clear();

            //Заполняем комбо-список должностями
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                cmbLecturerList.Items.Add(mdlData.colLecturer[i].Code + ". " + mdlData.colLecturer[i].FIO);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmbLecturerList.SelectedIndex = 0;
            }
            else
            {
                cmbLecturerList.SelectedIndex = NumFix;
            }
        }

        private void cmbLecturerList_SelectedIndexChanged(object sender, EventArgs e)
        {
            int LectHours = 0;
            int ExamHours = 0;
            int CredHours = 0;
            int RefHours = 0;
            int TutHours = 0;
            int LabHours = 0;
            int PractHours = 0;
            int IndHours = 0;
            int KRAPKHours = 0;
            int KursHours = 0;
            int PreDHours = 0;
            int DiplHours = 0;
            int TutPrHours = 0;
            int ProdHours = 0;
            int GAKHours = 0;
            int PostGradHours = 0;
            int VisitingHours = 0;
            int MagistryHours = 0;
            int SumHours = 0;

            int LectHours1 = 0;
            int ExamHours1 = 0;
            int CredHours1 = 0;
            int RefHours1 = 0;
            int TutHours1 = 0;
            int LabHours1 = 0;
            int PractHours1 = 0;
            int IndHours1 = 0;
            int KRAPKHours1 = 0;
            int KursHours1 = 0;
            int PreDHours1 = 0;
            int DiplHours1 = 0;
            int TutPrHours1 = 0;
            int ProdHours1 = 0;
            int GAKHours1 = 0;
            int PostGradHours1 = 0;
            int VisitingHours1 = 0;
            int MagistryHours1 = 0;
            int SumHours1 = 0;

            int LectHours2 = 0;
            int ExamHours2 = 0;
            int CredHours2 = 0;
            int RefHours2 = 0;
            int TutHours2 = 0;
            int LabHours2 = 0;
            int PractHours2 = 0;
            int IndHours2 = 0;
            int KRAPKHours2 = 0;
            int KursHours2 = 0;
            int PreDHours2 = 0;
            int DiplHours2 = 0;
            int TutPrHours2 = 0;
            int ProdHours2 = 0;
            int GAKHours2 = 0;
            int PostGradHours2 = 0;
            int VisitingHours2 = 0;
            int MagistryHours2 = 0;
            int SumHours2 = 0;

            int sumCurrent, sumStud, i, j;

            bool flgAccess;

            IList<clsDistribution> coll;

            if (flgCombine)
            {
                coll = mdlData.colCombineDistribution;
            }
            else
            {
                coll = mdlData.colDistribution;
            }

            //Считаем количество лекционных часов преподавателя
            for (i = 0; i <= coll.Count - 1; i++)
            {
                if (!coll[i].flgExclude)
                {
                    //
                    if (!(coll[i].Lecturer == null))
                    {
                        if ((coll[i].Lecturer.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])))
                        {
                            flgAccess = true;
                            if (chkOnlyUMO.Checked)
                            {
                                if (coll[i].Speciality != null)
                                {
                                    flgAccess = (!coll[i].Speciality.Diff.Equals("УО") &
                                                 !coll[i].Speciality.Diff.Equals("СУР"));
                                }
                            }
                           
                            if (flgAccess)
                            {
                                if (coll[i].Semestr.SemNum.Equals("1 семестр"))
                                {
                                    LectHours1 += coll[i].Lecture;
                                    ExamHours1 += coll[i].Exam;
                                    CredHours1 += coll[i].Credit;
                                    RefHours1 += coll[i].RefHomeWork;
                                    TutHours1 += coll[i].Tutorial;
                                    LabHours1 += coll[i].LabWork;
                                    PractHours1 += coll[i].Practice;
                                    IndHours1 += coll[i].IndividualWork;
                                    KRAPKHours1 += coll[i].KRAPK;
                                    KursHours1 += coll[i].KursProject;

                                    if (!coll[i].flgDistrib)
                                    {
                                        MagistryHours1 += coll[i].Magistry;
                                        PostGradHours1 += coll[i].PostGrad;
                                        GAKHours1 += coll[i].GAK;
                                        ProdHours1 += coll[i].ProducingPractice;
                                        TutPrHours1 += coll[i].TutorialPractice;
                                        PreDHours1 += coll[i].PreDiplomaPractice;
                                        DiplHours1 += coll[i].DiplomaPaper;
                                    }

                                    VisitingHours1 += coll[i].Visiting;
                                }

                                if (coll[i].Semestr.SemNum.Equals("2 семестр"))
                                {
                                    LectHours2 += coll[i].Lecture;
                                    ExamHours2 += coll[i].Exam;
                                    CredHours2 += coll[i].Credit;
                                    RefHours2 += coll[i].RefHomeWork;
                                    TutHours2 += coll[i].Tutorial;
                                    LabHours2 += coll[i].LabWork;
                                    PractHours2 += coll[i].Practice;
                                    IndHours2 += coll[i].IndividualWork;
                                    KRAPKHours2 += coll[i].KRAPK;
                                    KursHours2 += coll[i].KursProject;

                                    if (!coll[i].flgDistrib)
                                    {
                                        PreDHours2 += coll[i].PreDiplomaPractice;
                                        DiplHours2 += coll[i].DiplomaPaper;
                                        TutPrHours2 += coll[i].TutorialPractice;
                                        ProdHours2 += coll[i].ProducingPractice;
                                        GAKHours2 += coll[i].GAK;
                                        PostGradHours2 += coll[i].PostGrad;
                                        MagistryHours2 += coll[i].Magistry;
                                    }

                                    VisitingHours2 += coll[i].Visiting;
                                }

                                LectHours += coll[i].Lecture;
                                ExamHours += coll[i].Exam;
                                CredHours += coll[i].Credit;
                                RefHours += coll[i].RefHomeWork;
                                TutHours += coll[i].Tutorial;
                                LabHours += coll[i].LabWork;
                                PractHours += coll[i].Practice;
                                IndHours += coll[i].IndividualWork;
                                KRAPKHours += coll[i].KRAPK;
                                KursHours += coll[i].KursProject;

                                if (!coll[i].flgDistrib)
                                {
                                    PreDHours += coll[i].PreDiplomaPractice;
                                    DiplHours += coll[i].DiplomaPaper;
                                    TutPrHours += coll[i].TutorialPractice;
                                    ProdHours += coll[i].ProducingPractice;
                                    GAKHours += coll[i].GAK;
                                    PostGradHours += coll[i].PostGrad;
                                    MagistryHours += coll[i].Magistry;
                                }

                                VisitingHours += coll[i].Visiting;
                            }
                        }
                    }
                    //
                    else
                    {
                        if (coll[i].flgDistrib)
                        {
                            flgAccess = true;
                            if (chkOnlyUMO.Checked)
                            {
                                if (coll[i].Speciality != null)
                                {
                                    flgAccess = (!coll[i].Speciality.Diff.Equals("УО") &
                                                 !coll[i].Speciality.Diff.Equals("СУР"));
                                }
                            }

                            if (flgAccess)
                            {
                                sumCurrent = 0;
                                sumStud = 0;
                                for (j = 0; j <= mdlData.colStudents.Count - 1; j++)
                                {
                                    if (mdlData.colStudents[j].flgPlan)
                                    {
                                        //Если рассматриваемый преподаватель - руководитель студента
                                        //И если студент на том же курсе, где и дисциплина
                                        //И специальность студента должна соответствовать специальности нагрузки
                                        if (mdlData.colStudents[j].Lect.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])
                                            & mdlData.colStudents[j].KursNum.Equals(coll[i].KursNum)
                                            & mdlData.colStudents[j].Speciality.Equals(coll[i].Speciality))
                                        {
                                            sumCurrent += coll[i].Weight;
                                            sumStud++;
                                        }
                                    }
                                }

                                //
                                if (flgCombine)
                                {
                                    mdlData.toDetectUniformInHoured(ref sumCurrent, coll[i], mdlData.colLecturer[cmbLecturerList.SelectedIndex]);
                                }

                                //Если сумма изменилась
                                if (sumCurrent > 0)
                                {
                                    if (coll[i].Magistry > 0)
                                    {
                                        MagistryHours += sumCurrent;
                                    }

                                    if (coll[i].DiplomaPaper > 0)
                                    {
                                        DiplHours += sumCurrent;
                                    }

                                    if (coll[i].ProducingPractice > 0)
                                    {
                                        ProdHours += sumCurrent;
                                    }

                                    if (coll[i].TutorialPractice > 0)
                                    {
                                        TutPrHours += sumCurrent;
                                    }

                                    if (coll[i].PreDiplomaPractice > 0)
                                    {
                                        PreDHours += sumCurrent;
                                    }
                                }

                                if (coll[i].Semestr.SemNum.Equals("1 семестр"))
                                {
                                    sumCurrent = 0;
                                    sumStud = 0;
                                    for (j = 0; j <= mdlData.colStudents.Count - 1; j++)
                                    {
                                        if (mdlData.colStudents[j].flgPlan)
                                        {
                                            //Если рассматриваемый преподаватель - руководитель студента
                                            //И если студент на том же курсе, где и дисциплина
                                            //И специальность студента должна соответствовать специальности нагрузки
                                            if (mdlData.colStudents[j].Lect.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])
                                                & mdlData.colStudents[j].KursNum.Equals(coll[i].KursNum)
                                                & mdlData.colStudents[j].Speciality.Equals(coll[i].Speciality))
                                            {
                                                sumCurrent += coll[i].Weight;
                                                sumStud++;
                                            }
                                        }
                                    }

                                    //
                                    if (flgCombine)
                                    {
                                        mdlData.toDetectUniformInHoured(ref sumCurrent, coll[i], mdlData.colLecturer[cmbLecturerList.SelectedIndex]);
                                    }

                                    //Если сумма изменилась
                                    if (sumCurrent > 0)
                                    {
                                        if (coll[i].Magistry > 0)
                                        {
                                            MagistryHours1 += sumCurrent;
                                        }

                                        if (coll[i].DiplomaPaper > 0)
                                        {
                                            DiplHours1 += sumCurrent;
                                        }

                                        if (coll[i].ProducingPractice > 0)
                                        {
                                            ProdHours1 += sumCurrent;
                                        }

                                        if (coll[i].TutorialPractice > 0)
                                        {
                                            TutPrHours1 += sumCurrent;
                                        }

                                        if (coll[i].PreDiplomaPractice > 0)
                                        {
                                            PreDHours1 += sumCurrent;
                                        }
                                    }
                                }

                                if (coll[i].Semestr.SemNum.Equals("2 семестр"))
                                {
                                    sumCurrent = 0;
                                    sumStud = 0;
                                    for (j = 0; j <= mdlData.colStudents.Count - 1; j++)
                                    {
                                        if (mdlData.colStudents[j].flgPlan)
                                        {
                                            //Если рассматриваемый преподаватель - руководитель студента
                                            //И если студент на том же курсе, где и дисциплина
                                            //И специальность студента должна соответствовать специальности нагрузки
                                            if (mdlData.colStudents[j].Lect.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])
                                                & mdlData.colStudents[j].KursNum.Equals(coll[i].KursNum)
                                                & mdlData.colStudents[j].Speciality.Equals(coll[i].Speciality))
                                            {
                                                sumCurrent += coll[i].Weight;
                                                sumStud++;
                                            }
                                        }
                                    }

                                    //
                                    if (flgCombine)
                                    {
                                        mdlData.toDetectUniformInHoured(ref sumCurrent, coll[i], mdlData.colLecturer[cmbLecturerList.SelectedIndex]);
                                    }

                                    //Если сумма изменилась
                                    if (sumCurrent > 0)
                                    {
                                        if (coll[i].Magistry > 0)
                                        {
                                            MagistryHours2 += sumCurrent;
                                        }

                                        if (coll[i].DiplomaPaper > 0)
                                        {
                                            DiplHours2 += sumCurrent;
                                        }

                                        if (coll[i].ProducingPractice > 0)
                                        {
                                            ProdHours2 += sumCurrent;
                                        }

                                        if (coll[i].TutorialPractice > 0)
                                        {
                                            TutPrHours2 += sumCurrent;
                                        }

                                        if (coll[i].PreDiplomaPractice > 0)
                                        {
                                            PreDHours2 += sumCurrent;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            SumHours = LectHours + ExamHours + CredHours + RefHours + TutHours + LabHours + PractHours + IndHours +
                       KRAPKHours + KursHours + PreDHours + DiplHours + TutPrHours + ProdHours + GAKHours + 
                       PostGradHours + VisitingHours + MagistryHours;
            
            SumHours1 = LectHours1 + ExamHours1 + CredHours1 + RefHours1 + TutHours1 + LabHours1 + PractHours1 + IndHours1 +
                        KRAPKHours1 + KursHours1 + PreDHours1 + DiplHours1 + TutPrHours1 + ProdHours1 + GAKHours1 +
                        PostGradHours1 + VisitingHours1 + MagistryHours1;

            SumHours2 = LectHours2 + ExamHours2 + CredHours2 + RefHours2 + TutHours2 + LabHours2 + PractHours2 + IndHours2 +
                        KRAPKHours2 + KursHours2 + PreDHours2 + DiplHours2 + TutPrHours2 + ProdHours2 + GAKHours2 +
                        PostGradHours2 + VisitingHours2 + MagistryHours2;            

            txtLecturesAll.Text = LectHours.ToString();
            txtExamAll.Text = ExamHours.ToString();
            txtCreditAll.Text = CredHours.ToString();
            txtRefAll.Text = RefHours.ToString();
            txtTutAll.Text = TutHours.ToString();
            txtLabAll.Text = LabHours.ToString();
            txtPractAll.Text = PractHours.ToString();
            txtIndAll.Text = IndHours.ToString();
            txtKRAPKAll.Text = KRAPKHours.ToString();
            txtKursAll.Text = KursHours.ToString();
            txtPreDAll.Text = PreDHours.ToString();
            txtDiplomaAll.Text = DiplHours.ToString();
            txtTutPrAll.Text = TutPrHours.ToString();
            txtProdAll.Text = ProdHours.ToString();
            txtGAKAll.Text = GAKHours.ToString();
            txtMagAll.Text = MagistryHours.ToString();
            txtVisAll.Text = VisitingHours.ToString();
            txtPGAll.Text = PostGradHours.ToString();
            txtSumAll.Text = SumHours.ToString();

            txtLectures1.Text = LectHours1.ToString();
            txtExam1.Text = ExamHours1.ToString();
            txtCred1.Text = CredHours1.ToString();
            txtRef1.Text = RefHours1.ToString();
            txtTut1.Text = TutHours1.ToString();
            txtLab1.Text = LabHours1.ToString();
            txtPract1.Text = PractHours1.ToString();
            txtInd1.Text = IndHours1.ToString();
            txtKRAPK1.Text = KRAPKHours1.ToString();
            txtKurs1.Text = KursHours1.ToString();
            txtPreD1.Text = PreDHours1.ToString();
            txtDiploma1.Text = DiplHours1.ToString();
            txtTutPr1.Text = TutPrHours1.ToString();
            txtProd1.Text = ProdHours1.ToString();
            txtGAK1.Text = GAKHours1.ToString();
            txtMag1.Text = MagistryHours1.ToString();
            txtVis1.Text = VisitingHours1.ToString();
            txtPG1.Text = PostGradHours1.ToString();
            txtSum1.Text = SumHours1.ToString();

            txtLectures2.Text = LectHours2.ToString();
            txtExam2.Text = ExamHours2.ToString();
            txtCred2.Text = CredHours2.ToString();
            txtRef2.Text = RefHours2.ToString();
            txtTut2.Text = TutHours2.ToString();
            txtLab2.Text = LabHours2.ToString();
            txtPract2.Text = PractHours2.ToString();
            txtInd2.Text = IndHours2.ToString();
            txtKRAPK2.Text = KRAPKHours2.ToString();
            txtKurs2.Text = KursHours2.ToString();
            txtPreD2.Text = PreDHours2.ToString();
            txtDiploma2.Text = DiplHours2.ToString();
            txtTutPr2.Text = TutPrHours2.ToString();
            txtProd2.Text = ProdHours2.ToString();
            txtGAK2.Text = GAKHours2.ToString();
            txtMag2.Text = MagistryHours2.ToString();
            txtVis2.Text = VisitingHours2.ToString();
            txtPG2.Text = PostGradHours2.ToString();
            txtSum2.Text = SumHours2.ToString();
        }

        private void FillGrid()
        {
            int countRow;

            IList<clsDistribution> coll;

            if (flgCombine)
            {
                coll = mdlData.colCombineDistribution;
            }
            else
            {
                coll = mdlData.colDistribution;
            }

            NamesDiploma = new string[0];
            NamesDiploma = findSameNamesDiploma(coll);

            GroupsDiploma = new string[0];
            GroupsFullDiploma = new string[0];

            findSameGroupsDiploma(coll, ref GroupsDiploma, ref GroupsFullDiploma);

            NamesMagistry = new string[0];
            NamesMagistry = findSameNamesMagistry(coll);

            GroupsMagistry = new string[0];
            GroupsFullMagistry = new string[0];

            findSameGroupsMagistry(coll, ref GroupsMagistry, ref GroupsFullMagistry);

            NamesTutorialPr = new string[0];
            NamesTutorialPr = findSameNamesTutorialPr(coll);

            GroupsTutorialPr = new string[0];
            GroupsFullTutorialPr = new string[0];

            findSameGroupsTutorialPr(coll, ref GroupsTutorialPr, ref GroupsFullTutorialPr);

            NamesProdPr = new string[0];
            NamesProdPr = findSameNamesProdPr(coll);

            GroupsProdPr = new string[0];
            GroupsFullProdPr = new string[0];

            findSameGroupsProdPr(coll, ref GroupsProdPr, ref GroupsFullProdPr);

            //Очищаем сетку
            dgNagruzka.Rows.Clear();
            dgNagruzka.Columns.Clear();

            //Делаем невидимыми нуль-строку и нуль-столбец
            dgNagruzka.ColumnHeadersVisible = false;
            dgNagruzka.RowHeadersVisible = false;

            //Задаём количество столбцов
            //оно остаётся неизменным
            //в количестве 6 штук
            for (int i = 0; i <= 5; i++)
            {
                dgNagruzka.Columns.Add("", "");
            }

            countRow = countTableRows(coll, cmbLecturerList.SelectedIndex);

            for (int i = 0; i <= countRow - 2; i++)
            {
                dgNagruzka.Rows.Add();
            }

            DataGridFiller(coll);
        }

        //
        private void FillGridDiplomas()
        {
            int countRow;

            IList<clsDistribution> coll;

            if (flgCombine)
            {
                coll = mdlData.colCombineDistribution;
            }
            else
            {
                coll = mdlData.colDistribution;
            }

            NamesDiploma = new string[0];
            NamesDiploma = findSameNamesDiploma(coll);

            GroupsDiploma = new string[0];
            GroupsFullDiploma = new string[0];

            findSameGroupsDiploma(coll, ref GroupsDiploma, ref GroupsFullDiploma);

            //Очищаем сетку
            dgNagruzka.Rows.Clear();
            dgNagruzka.Columns.Clear();

            //Делаем невидимыми нуль-строку и нуль-столбец
            dgNagruzka.ColumnHeadersVisible = false;
            dgNagruzka.RowHeadersVisible = false;

            //Задаём количество столбцов
            //оно остаётся неизменным
            //в количестве 6 штук
            for (int i = 0; i <= 5; i++)
            {
                dgNagruzka.Columns.Add("", "");
            }

            countRow = countTableRowsDiploma(coll, cmbLecturerList.SelectedIndex);

            for (int i = 0; i <= countRow - 2; i++)
            {
                dgNagruzka.Rows.Add();
            }

            DataGridFillerDiploma(coll);
        }

        //
        private void FillGridMagistry()
        {
            int countRow;

            IList<clsDistribution> coll;

            if (flgCombine)
            {
                coll = mdlData.colCombineDistribution;
            }
            else
            {
                coll = mdlData.colDistribution;
            }

            NamesMagistry = new string[0];
            NamesMagistry = findSameNamesMagistry(coll);

            GroupsMagistry = new string[0];
            GroupsFullMagistry = new string[0];

            findSameGroupsMagistry(coll, ref GroupsMagistry, ref GroupsFullMagistry);

            //Очищаем сетку
            dgNagruzka.Rows.Clear();
            dgNagruzka.Columns.Clear();

            //Делаем невидимыми нуль-строку и нуль-столбец
            dgNagruzka.ColumnHeadersVisible = false;
            dgNagruzka.RowHeadersVisible = false;

            //Задаём количество столбцов
            //оно остаётся неизменным
            //в количестве 6 штук
            for (int i = 0; i <= 5; i++)
            {
                dgNagruzka.Columns.Add("", "");
            }

            countRow = countTableRowsMagistry(coll, cmbLecturerList.SelectedIndex);

            for (int i = 0; i <= countRow - 2; i++)
            {
                dgNagruzka.Rows.Add();
            }

            DataGridFillerMagistry(coll);
        }

        //
        private void chkCombine_CheckedChanged(object sender, EventArgs e)
        {
            if (chkCombine.Checked)
            {
                flgCombine = true;
            }
            else
            {
                flgCombine = false;
            }

            FillLecturerList();
        }

        private void wordCore(object ObjMissing, Word._Application ObjWord)
        {
            bool flgBoss;

            flgBoss = (mdlData.colLecturer[cmbLecturerList.SelectedIndex].Duty.Duty.Equals("Заведующий кафедрой")
                        || mdlData.colLecturer[cmbLecturerList.SelectedIndex].Duty1.Duty.Equals("Заведующий кафедрой"));
            //Добавляем новый чистый документ Word
            Word._Document ObjDoc = ObjWord.Application.Documents.Add();
            ObjDoc.Activate();

            //Заполнение в реальном времени с отображением
            //ObjWord.Visible = true;

            ObjDoc.PageSetup.TopMargin = 0.75f / 0.03527f;
            ObjDoc.PageSetup.BottomMargin = 0.75f / 0.03527f;

            //В новых планах необходимо перенастроить границы
            if (chkAfter2015.Checked)
            {
                ObjDoc.PageSetup.LeftMargin = 2f / 0.03527f;
                ObjDoc.PageSetup.RightMargin = 2f / 0.03527f;
            }

            //-----------ПИШЕМ ЗАГЛАВИЕ УНИВЕРСИТЕТА
            spacesBefore(ObjMissing, ObjDoc, 2);
            textUniHeader(ObjMissing, ObjDoc);

            //В старых планах один пробел
            if (!chkAfter2015.Checked)
            {
                spacesAfter(ObjMissing, ObjDoc, 2);
            }
            //В новых планаз два пробела
            else
            {
                spacesAfter(ObjMissing, ObjDoc, 3);
            }

            if (!chkAfter2015.Checked)
            {
                //-----------ПИШЕМ УТВЕРЖДАЮ
                textAgreed(ObjMissing, ObjDoc);
                //-----------ПИШЕМ ДИРЕКТОРА ИНСТИТУТА
                textDirector(ObjMissing, ObjDoc);
                //-----------ПИШЕМ ФИО ДИРЕКТОРА ИНСТИТУТА
                textDirectorFIO(ObjMissing, ObjDoc);

                if (!flgBoss)
                {
                    //-----------ПИШЕМ ЗАВЕДУЮЩЕГО КАФЕДРОЙ
                    textMain(ObjMissing, ObjDoc);
                    //-----------ПИШЕМ ФИО ЗАВЕДУЮЩЕГО КАФЕДРОЙ
                    textMainFIO(ObjMissing, ObjDoc);
                }
            }
            else
            {
                //-----------ТАБЛИЧНОЕ СОГЛАСОВАНИЕ
                if (flgBoss)
                {
                    textAgreedInfoBoss(ObjMissing, ObjDoc);
                }
                else
                {
                    textAgreedInfoStd(ObjMissing, ObjDoc);
                }
            }

            //-----------ПИШЕМ ПОЛНОЕ НАЗВАНИЕ КАФЕДРЫ
            if (!chkAfter2015.Checked)
            {
                if (!flgBoss)
                {
                    spacesBefore(ObjMissing, ObjDoc, 6);
                }
                else
                {
                    spacesBefore(ObjMissing, ObjDoc, 8);
                }
            }
            else
            {
                spacesBefore(ObjMissing, ObjDoc, 5);
            }

            textDepartName(ObjMissing, ObjDoc);
            spacesAfter(ObjMissing, ObjDoc, 3);
            //-----------ПИШЕМ ИНДИВИДУАЛЬНЫЙ ПЛАН
            textPlan(ObjMissing, ObjDoc);

            if (chkAfter2015.Checked)
            {
                if (!flgBoss)
                {
                    textLecturer(ObjMissing, ObjDoc);
                }
                else
                {
                    textBoss(ObjMissing, ObjDoc);
                }
            }

            spacesAfter(ObjMissing, ObjDoc, 1);

            //-----------ПИШЕМ ГОД РАСПРОСТРАНЕНИЯ
            if (!chkAfter2015.Checked)
            {
                textPlanYear(ObjMissing, ObjDoc);
            }
            else
            {
                textPlanYearNew(ObjMissing, ObjDoc);
            }

            spacesAfter(ObjMissing, ObjDoc, 5);

            if (!chkAfter2015.Checked)
            {
                //-----------ИНФО О ПРЕПОДАВАТЕЛЕ
                textLecturerInfo(ObjMissing, ObjDoc);
                //-----------БЛОК НАПОМИНАЛОК
                spacesBefore(ObjMissing, ObjDoc, 5);
                textReminder(ObjMissing, ObjDoc);
            }
            else
            {
                //-----------ИНФО О ПРЕПОДАВАТЕЛЕ
                textLecturerInfoNew(ObjMissing, ObjDoc);
                //-----------БЛОК НАПОМИНАЛОК
                spacesBefore(ObjMissing, ObjDoc, 1);
                textReminderNew(ObjMissing, ObjDoc);
            }

            //-----------ПЕРЕХОД НА СТРАНИЦУ 2
            pageBreaker(ObjMissing, ObjDoc);
            //-----------ПИШЕМ УЧЕБНАЯ РАБОТА
            textStudWork(ObjMissing, ObjDoc);

            textStudWorkFromCol(ObjMissing, ObjDoc);

            //-----------ПЕРЕХОД НА СТРАНИЦУ 3 ИЛИ 4
            pageBreaker(ObjMissing, ObjDoc);
            //-----------ПИШЕМ УЧЕБНО-МЕТОДИЧЕСКАЯ РАБОТА
            textMethodWork(ObjMissing, ObjDoc);
            //-----------ПИШЕМ УЧЕБНО-МЕТОДИЧЕСКУЮ РАБОТУ
            textMethodWorkFromCol(ObjMissing, ObjDoc);

            if (chkAfter2015.Checked)
            {
                if (flgBoss)
                {
                    textMethodAgreed(ObjMissing, ObjDoc);
                }
            }

            //-----------ПЕРЕХОД НА СТРАНИЦУ 5 ИЛИ 6
            pageBreaker(ObjMissing, ObjDoc);
            //-----------ПИШЕМ НАУЧНО-ИССЛЕДОВАТЕЛЬСКАЯ РАБОТА
            textScienceWork(ObjMissing, ObjDoc);
            //-----------ПИШЕМ НАУЧНО-ИССЛЕДОВАТЕЛЬСКУЮ РАБОТУ
            textScienceWorkFromCol(ObjMissing, ObjDoc);

            if (chkAfter2015.Checked)
            {
                if (flgBoss)
                {
                    textScienceAgreed(ObjMissing, ObjDoc);
                }
            }

            //-----------ПЕРЕХОД НА СТРАНИЦУ 7 ИЛИ 8
            pageBreaker(ObjMissing, ObjDoc);
            //-----------ПИШЕМ ОРГАНИЗАЦИОННО-МЕТОДИЧЕСКАЯ РАБОТА
            textAdminWork(ObjMissing, ObjDoc);
            //-----------ПИШЕМ ОРГАНИЗАЦИОННО-МЕТОДИЧЕСКУЮ РАБОТУ
            textAdminWorkFromCol(ObjMissing, ObjDoc);
            //-----------ФОРМИРУЕМ СТРОКУ ДЛЯ ПОДПИСИ ПРЕПОДАВАТЕЛЯ 
            if (!flgBoss)
            {
                textSignature(ObjMissing, ObjDoc);
            }
            else
            {
                textSignatureMain(ObjMissing, ObjDoc);
            }
            //ObjWord.Visible = true;

            ObjDoc.SaveAs(Application.StartupPath + @"\"
                + mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO + " " +
                DateTime.Now.Date.ToString("yyyyMMdd") + " " + 
                DateTime.Now.TimeOfDay.ToString("hhmmss") + ".docx");

            ObjDoc.Close();
        }

        private void wordCoreHoured(object ObjMissing, Word._Application ObjWord)
        {
            //Добавляем новый чистый документ Word
            Word._Document ObjDoc = ObjWord.Application.Documents.Add();
            ObjDoc.Activate();

            ObjWord.Visible = true;
            //Настройка границ
            mdlData.WordPageDefault(ref ObjWord, ref ObjDoc, 3f, 1.5f, 0.63f, 1.27f);
            //Настройка 70% масштаба
            ObjDoc.ActiveWindow.View.Zoom.Percentage = 70;

            //-----------ПИШЕМ ЗАГЛАВИЕ УНИВЕРСИТЕТА
            textMinistryHeaderHoured(ObjMissing, ObjDoc);
            textTitleHeaderHoured(ObjMissing, ObjDoc);
            textUniversityHeaderHoured(ObjMissing, ObjDoc);
            //-----------ПИШЕМ ПАРАМЕТРЫ ИНДИВИДУАЛЬНОГО ПЛАНА
            textPlanHeaderHoured(ObjMissing, ObjDoc);
            //-----------ТАБУЛИРОВАНИЕ ПРЕПОДАВАТЕЛЯ И КАФЕДРЫ
            tabLectAndDepartHoured(ObjMissing, ObjDoc);
            //-----------ПИШЕМ ЗАГЛАВИЕ ПЕРВОЙ ТАБЛИЦЫ
            textTabFirstSemHoured(ObjMissing, ObjDoc);
            //-----------ТАБУЛИРУЕМ НАГРУЗКУ ПЕРВОГО СЕМЕСТРА
            tabFirstSemHoured(ObjMissing, ObjDoc);
            //-----------ПИШЕМ ЗАГЛАВИЕ ВТОРОЙ ТАБЛИЦЫ
            textTabSecondSemHoured(ObjMissing, ObjDoc);
            //-----------ТАБУЛИРУЕМ НАГРУЗКУ ВТОРОГО СЕМЕСТРА
            tabSecondSemHoured(ObjMissing, ObjDoc);
            //-----------ПИШЕМ ПОДВАЛ
            //textBasementHoured(ObjMissing, ObjDoc);

            ObjDoc.SaveAs(Application.StartupPath + @"\"
                + mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO + " почасовая " + 
                DateTime.Now.Date.ToString("yyyyMMdd") + " " + 
                DateTime.Now.TimeOfDay.ToString("hhmmss") + ".docx");

            ObjDoc.Close();
        }

        private void StandartTextFormat(ref Word.Paragraph ObjParagraph)
        {
            //Выравнивание по ширине листа
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            //Интервал полуторный
            ObjParagraph.Format.Space15();
            //Нет отступов перед
            ObjParagraph.Format.SpaceBefore = 0f;
            ////Нет отступов после
            ObjParagraph.Format.SpaceAfter = 0f;
            //Размер шрифта 14 пт
            ObjParagraph.Range.Font.Size = 14;
            //Семейство шрифта
            ObjParagraph.Range.Font.Name = "Times New Roman";
        }

        private void intoWord()
        {
            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;

            try
            {
                //Создаём новое Word приложение
                Word._Application ObjWord = new Word.Application();

                wordCore(ObjMissing, ObjWord);

                ObjWord.Quit();
            }
            catch
            {
                MessageBox.Show("Возможно на этом компьютере присутствует проблема совместимости с MS Word." +
                " Попробуйте установить версию 2007 и выше.");
            }
        }

        private void intoWordHoured()
        {
            int count = 0;
            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;

            //Проверяем, есть ли у преподавателя хотя бы строчка почасовой нагрузки
            for (int i = 0; i <= mdlData.colHouredDistribution.Count - 1; i++)
            {
                //Проверка по наличию фамилии, имени и отчества среди элементов
                //коллекции почасовой нагрузки преподавателей
                if (mdlData.colHouredDistribution[i].Lecturer.FIO.Equals(
                    mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO))
                {
                    //Если элемент, относящийся к преподавателю найден,
                    //то увеличиваем значение счётчика на единицу
                    count++;
                }
            }

            //Если значение счётчика больше нуля, то можно переходить к составлению
            //плана на почасовую нагрузку
            if (count > 0)
            {
                try
                {
                    //Создаём новое Word приложение
                    Word._Application ObjWord = new Word.Application();

                    wordCoreHoured(ObjMissing, ObjWord);

                    ObjWord.Quit();
                }
                catch
                {
                    MessageBox.Show("Возможно на этом компьютере присутствует проблема совместимости с MS Word.\n" +
                                    "Попробуйте установить версию 2007 и выше.", "Несовместимость версий Офисных приложений",
                                    MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("У выбранного преподавателя нет почасовой нагрузки.", "Нечего выгружать", MessageBoxButtons.OK);
            }
        }

        private void intoExcel()
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

                for (int i = 0; i <= dgNagruzka.RowCount - 1; i++)
                {
                    for (int j = 0; j <= dgNagruzka.ColumnCount - 1; j++)
                    {
                        if (!(dgNagruzka[j, i].Value == null))
                        {
                            ObjWorkSheet.Cells[i + 1, j + 1] = dgNagruzka[j, i].Value.ToString();
                        }
                    }
                }

                //Задаём диапазон для ячеек, подлежащих форматированию
                var cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(1, 1),
                    mdlData.ExcelCellTranslator(dgNagruzka.RowCount, dgNagruzka.ColumnCount));

                //внутренние вертикальные
                cells.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = 2;
                cells.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                //внутренние горизонтальные
                cells.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = 2;
                cells.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                //верхняя внешняя
                cells.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = 3;
                cells.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                //правая внешняя
                cells.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = 3;
                cells.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                //левая внешняя
                cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = 3;
                cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                //нижняя внешняя
                cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 3;
                cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                cells.EntireColumn.AutoFit();

                //Вызываем  эксель
                ObjExcel.Visible = true;
                ObjExcel.UserControl = true;

                ObjWorkBook.SaveCopyAs(Application.StartupPath + @"\" +
                    mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO + " " + 
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

        private void allIntoExcel()
        {
            try
            {
                //Создаём новое Excel приложение
                Excel.Application ObjExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook;
                Excel.Worksheet ObjWorkSheet;

                for (int l = 0; l <= cmbLecturerList.Items.Count - 1; l++)
                {
                    cmbLecturerList.SelectedIndex = l;
                    FillGrid();

                    //Книга
                    ObjWorkBook = ObjExcel.Workbooks.Add(Missing.Value);
                    //Таблица
                    ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

                    for (int i = 0; i <= dgNagruzka.RowCount - 1; i++)
                    {
                        for (int j = 0; j <= dgNagruzka.ColumnCount - 1; j++)
                        {
                            if (!(dgNagruzka[j, i].Value == null))
                            {
                                ObjWorkSheet.Cells[i + 1, j + 1] = dgNagruzka[j, i].Value.ToString();
                            }
                        }
                    }

                    //Задаём диапазон для ячеек, подлежащих форматированию
                    var cells = ObjWorkSheet.get_Range(mdlData.ExcelCellTranslator(1, 1),
                        mdlData.ExcelCellTranslator(dgNagruzka.RowCount, dgNagruzka.ColumnCount));

                    //внутренние вертикальные
                    cells.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = 2;
                    cells.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //внутренние горизонтальные
                    cells.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = 2;
                    cells.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //верхняя внешняя
                    cells.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = 3;
                    cells.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //правая внешняя
                    cells.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = 3;
                    cells.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //левая внешняя
                    cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = 3;
                    cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //нижняя внешняя
                    cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 3;
                    cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    cells.EntireColumn.AutoFit();

                    //Вызываем  эксель
                    ObjExcel.Visible = true;
                    ObjExcel.UserControl = true;

                    ObjWorkBook.SaveCopyAs(Application.StartupPath + @"\" +
                        mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO + " " + 
                        DateTime.Now.Date.ToString("yyyyMMdd") + " " +
                        DateTime.Now.TimeOfDay.ToString("hhmmss") + ".xlsx");
                    ObjWorkBook.Close(false, "", Missing.Value);

                }

                ObjExcel.Quit();
            }
            catch
            {
                MessageBox.Show("Возможно на этом компьютере присутствует проблема совместимости с MS Exсel." +
                                " Попробуйте установить версию 2007 и выше.");
            }
        }

        private void allIntoWord()
        {
            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;

            try
            {
                //Создаём новое Word приложение
                Word._Application ObjWord = new Word.Application();

                for (int k = 0; k <= cmbLecturerList.Items.Count - 1; k++)
                {
                    cmbLecturerList.SelectedIndex = k;

                    if (mdlData.colLecturer[k].Rate > 0)
                    {
                        wordCore(ObjMissing, ObjWord);
                    }
                }

                ObjWord.Quit();
            }
            catch
            {
                MessageBox.Show("Возможно на этом компьютере присутствует проблема совместимости с MS Word." +
                " Попробуйте установить версию 2007 и выше.");
            }
        }

        private void allIntoWordHoured()
        {
            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;

            try
            {
                //Создаём новое Word приложение
                Word._Application ObjWord = new Word.Application();

                for (int k = 0; k <= cmbLecturerList.Items.Count - 1; k++)
                {
                    cmbLecturerList.SelectedIndex = k;

                    wordCore(ObjMissing, ObjWord);

                }

                ObjWord.Quit();
            }
            catch
            {
                MessageBox.Show("Возможно на этом компьютере присутствует проблема совместимости с MS Word." +
                " Попробуйте установить версию 2007 и выше.");
            }
        }

        private void textUniHeader(object ObjMissing, Word._Document ObjDoc)
        {
            Word.Paragraph ObjParagraph;

            //Добавляем абзац текста в начало документа
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            
            //Текстуальное содержимое абзаца
            if (chkAfter2015.Checked)
            {
                ObjParagraph.Range.Text = mdlData.MinistryName;

                //Жирный шрифт
                ObjParagraph.Range.Font.Bold = 1;
                //Все заглавные
                ObjParagraph.Range.Font.AllCaps = 1;
                //Размер шрифта 12 пт
                ObjParagraph.Range.Font.Size = 12;
                //Times New Roman
                ObjParagraph.Range.Font.Name = "Times New Roman";
                //Выравнивание по центру
                ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                //Отступ в 0 пт после абзаца
                ObjParagraph.Format.SpaceAfter = 0;
                //Отступ в 0 пт до абзаца
                ObjParagraph.Format.SpaceBefore = 0;
                //Одинарный межстрочный интервал
                ObjParagraph.Format.Space1();
                //Добавляем ещё один абзац текста
                ObjParagraph.Range.InsertParagraphAfter();

                ObjParagraph.Range.Text = mdlData.UniversityPrefName;
                //Добавляем ещё один абзац текста
                ObjParagraph.Range.InsertParagraphAfter();

                ObjParagraph.Range.Text = mdlData.UniversityName + " " +
                                            mdlData.UniversitySuffName;
            }
            else
            {
                ObjParagraph.Range.Text =
                    "Федеральное государственное бюджетное образовательное учреждение " +
                    "высшего профессионального образования \"Московский государственный университет " +
                    "путей сообщения\" (МГУПС (МИИТ))";
            }

            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Все заглавные
            ObjParagraph.Range.Font.AllCaps = 1;
            //Размер шрифта 12 пт
            ObjParagraph.Range.Font.Size = 12;
            //Times New Roman
            ObjParagraph.Range.Font.Name = "Times New Roman";
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Отступ в 0 пт после абзаца
            ObjParagraph.Format.SpaceAfter = 0;
            //Отступ в 0 пт до абзаца
            ObjParagraph.Format.SpaceBefore = 0;
            //Одинарный межстрочный интервал
            ObjParagraph.Format.Space1();
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textBasementHoured(object ObjMissing, Word._Document ObjDoc)
        {
            Word.Paragraph ObjParagraph;
            IList<clsDistribution> coll;
            object EndOfDoc = "\\endofdoc";

            coll = mdlData.colHouredDistribution;
            //Добавляем абзац текста в начало документа
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphBefore();
            //Добавляем абзац текста в документ
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            //Текстуальное содержимое абзаца
            ObjParagraph.Range.Text = "Преподаватель \t" + mdlData.SplitFIOString(mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO, false, true);
            //Выравнивание по левому краю
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //Размер шрифта 12 пт
            ObjParagraph.Range.Font.Size = 12;
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Выставляем каретку
            ObjParagraph.TabStops.Add(6.25f / 0.03527f, Word.WdTabAlignment.wdAlignTabLeft, Word.WdTabLeader.wdTabLeaderLines);
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();

            //Добавляем абзац текста в документ
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            //Текстуальное содержимое абзаца
            ObjParagraph.Range.Text = "Директор института \t" + mdlData.SplitFIOString(mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO, false, true);
            //Выравнивание по левому краю
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //Размер шрифта 12 пт
            ObjParagraph.Range.Font.Size = 12;
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Выставляем каретку
            ObjParagraph.TabStops.Add(6.25f / 0.03527f, Word.WdTabAlignment.wdAlignTabLeft, Word.WdTabLeader.wdTabLeaderLines);
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();

            //Добавляем абзац текста в документ
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            //Текстуальное содержимое абзаца
            ObjParagraph.Range.Text = "Заведующий кафедрой \t" + mdlData.SplitFIOString(mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO, false, true);
            //Выравнивание по левому краю
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //Размер шрифта 12 пт
            ObjParagraph.Range.Font.Size = 12;
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Выставляем каретку
            ObjParagraph.TabStops.Add(6.25f / 0.03527f, Word.WdTabAlignment.wdAlignTabLeft, Word.WdTabLeader.wdTabLeaderLines);
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();

            //Передано в учебное управление
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            //Текстуальное содержимое абзаца
            ObjParagraph.Range.Text = "Передано в учебное управление \t" + "\"__\" ___________ 20____ г.";
            //Выравнивание по левому краю
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //Размер шрифта 12 пт
            ObjParagraph.Range.Font.Size = 12;
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Выставляем каретку
            ObjParagraph.TabStops.Add(6.25f / 0.03527f, Word.WdTabAlignment.wdAlignTabLeft);
        }

        private void tabSecondSemHoured(object ObjMissing, Word._Document ObjDoc)
        {
            int rows;
            int sumII;
            int countRow;

            IList<clsDistribution> coll;
            object EndOfDoc = "\\endofdoc";
            Word.Table ObjTable;
            Word.Range ObjWordRange;

            coll = mdlData.colHouredDistribution;

            //Считаем строки для второго семестра
            rows = countTableRowsHoured(coll, cmbLecturerList.SelectedIndex, 3, mdlData.colSemestr[2]);

            //Вставляем таблицу N x 6, заполняем её данными о преподавателе
            //где N - количество строк, записываемов в переменную rows
            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjTable = ObjDoc.Tables.Add(ObjWordRange, rows, 6, ref ObjMissing, ref ObjMissing);
            ObjTable.Borders.Enable = 1;
            ObjTable.Range.ParagraphFormat.SpaceAfter = 0;
            ObjTable.Range.ParagraphFormat.LeftIndent = 0;
            ObjTable.Range.ParagraphFormat.RightIndent = 0;
            ObjTable.Range.Font.Size = 12;
            ObjTable.Range.Font.Bold = 0;
            ObjTable.Range.Font.AllCaps = 0;

            ObjTable.Cell(1, 1).Range.Text = "Дисциплина";
            ObjTable.Cell(1, 2).Range.Text = "Группа";
            ObjTable.Cell(1, 3).Range.Text = "Вид занятий";
            ObjTable.Cell(1, 4).Range.Text = "Кол-во часов";
            ObjTable.Cell(1, 5).Range.Text = "Стоимость часа (руб.)";
            ObjTable.Cell(1, 6).Range.Text = "Общая стоимость оплаты";
            ObjTable.Rows[1].Range.Font.Bold = 1;

            countRow = 2;
            sumII = 0;
            for (int i = 0; i <= coll.Count - 1; i++)
            {
                if (coll[i].Semestr.SemNum.Equals(mdlData.colSemestr[2].SemNum))
                {
                    if (!(coll[i].Lecturer == null))
                    {
                        if ((coll[i].Lecturer.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])))
                        {
                            //Если есть лекционные часы - добавляем строку
                            if (!(coll[i].Lecture == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "Лекции";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].Lecture.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].Lecture;
                                countRow++;
                            }

                            //Если есть экзаменационные часы - добавляем строку
                            if (!(coll[i].Exam == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "Экзамен";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].Exam.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].Exam;
                                countRow++;
                            }

                            //Если есть зачётные часы - добавляем строку
                            if (!(coll[i].Credit == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "Зачёт";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].Credit.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].Credit;
                                countRow++;
                            }

                            //Если есть часы на реферат - добавляем строку
                            if (!(coll[i].RefHomeWork == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "Рефераты, домашнее задание";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].RefHomeWork.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].RefHomeWork;
                                countRow++;
                            }

                            //Если есть часы на консультацию - добавляем строку
                            if (!(coll[i].Tutorial == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "Консультация";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].Tutorial.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].Tutorial;
                                countRow++;
                            }

                            //Если есть часы на лабораторные работы - добавляем строку
                            if (!(coll[i].LabWork == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "Лабораторные работы";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].LabWork.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].LabWork;
                                countRow++;
                            }

                            //Если есть часы на практические занятия - добавляем строку
                            if (!(coll[i].Practice == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "Практические занятия";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].Practice.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].Practice;
                                countRow++;
                            }

                            //Если есть часы на индивидуальные занятия - добавляем строку
                            if (!(coll[i].IndividualWork == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "Индивидуальные задания";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].IndividualWork.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].IndividualWork;
                                countRow++;
                            }

                            //Если есть часы на КРАПК - добавляем строку
                            if (!(coll[i].KRAPK == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "КРАПК";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].KRAPK.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].KRAPK;
                                countRow++;
                            }

                            //Если есть часы на курсовой проект - добавляем строку
                            if (!(coll[i].KursProject == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "Курсовой проект";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].KursProject.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].KursProject;
                                countRow++;
                            }

                            //Если есть часы на преддипломную практику - добавляем строку
                            if (!(coll[i].PreDiplomaPractice == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "Преддипломная практика";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].PreDiplomaPractice.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].PreDiplomaPractice;
                                countRow++;
                            }

                            //Если есть часы на Дипломный проект - добавляем строку
                            if (!(coll[i].DiplomaPaper == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "Дипломный проект";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].DiplomaPaper.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].DiplomaPaper;
                                countRow++;
                            }

                            //Если есть часы на учебную практику - добавляем строку
                            if (!(coll[i].TutorialPractice == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "Учебная практика";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].TutorialPractice.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].TutorialPractice;
                                countRow++;
                            }

                            //Если есть часы на производственную практику - добавляем строку
                            if (!(coll[i].ProducingPractice == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "Производственная практика";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].ProducingPractice.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].ProducingPractice;
                                countRow++;
                            }

                            //Если есть часы на ГАК - добавляем строку
                            if (!(coll[i].GAK == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "ГАК";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].GAK.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].GAK;
                                countRow++;
                            }

                            //Если есть часы на Посещение занятий - добавляем строку
                            if (!(coll[i].Visiting == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "Посещение занятий";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].Visiting.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].Visiting;
                                countRow++;
                            }

                            //Если есть часы на Аспирантуру - добавляем строку
                            if (!(coll[i].PostGrad == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "Аспирантура";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].PostGrad.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].PostGrad;
                                countRow++;
                            }

                            //Если есть часы на Руководство магистрами - добавляем строку
                            if (!(coll[i].Magistry == 0))
                            {
                                ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                ObjTable.Cell(countRow, 3).Range.Text = "Руководство магистерской программой";
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].Magistry.ToString();

                                if (coll[i].Lecturer.Duty.Equals("Ассистент") ||
                                    coll[i].Lecturer.Duty.Equals("Старший преподаватель"))
                                {
                                    ObjTable.Cell(countRow, 5).Range.Text = "";
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Equals("Доцент"))
                                    {
                                        ObjTable.Cell(countRow, 5).Range.Text = "";
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Equals("Профессор"))
                                        {
                                            ObjTable.Cell(countRow, 5).Range.Text = "";
                                        }
                                    }
                                }

                                ObjTable.Cell(countRow, 6).Range.Text = "";

                                sumII += coll[i].Magistry;
                                countRow++;
                            }
                        }
                    }
                }
            }

            ObjTable.Cell(countRow, 1).Range.Text = "Всего за семестр";
            ObjTable.Cell(countRow, 4).Range.Text = sumII.ToString();
            ObjTable.Cell(countRow, 6).Range.Text = "";
            countRow++;

            ObjTable.Cell(countRow, 1).Range.Text = "Всего за год";
            ObjTable.Cell(countRow, 4).Range.Text = (sumII).ToString();
            ObjTable.Cell(countRow, 6).Range.Text = "";
            ObjTable.Rows[countRow].Range.Font.Bold = 1;
        }

        private void textTabSecondSemHoured(object ObjMissing, Word._Document ObjDoc)
        {
            Word.Paragraph ObjParagraph;
            //Добавляем абзац текста в документ
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Текстуальное содержимое абзаца
            ObjParagraph.Range.Text = "2-й семестр";
            //Размер шрифта 12 пт
            ObjParagraph.Range.Font.Size = 12;
            //Обычный шрифт
            ObjParagraph.Range.Font.Bold = 0;
        }

        /// <summary>
        /// Функция сортировки нагрузки по иерархии: преподаватель, курс, дисциплина
        /// </summary>
        /// <param name="coll"></param>
        /// <returns></returns>
        private IList<clsDistribution> SortLoad(IList<clsDistribution> coll)
        {
            int i, j, k;
            IList<clsDistribution> tmpColl = new List<clsDistribution>();

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
                                    tmpColl.Add(mdlData.colDistribution[k]);
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
                                            tmpColl.Add(mdlData.colDistribution[k]);
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
                                            tmpColl.Add(mdlData.colDistribution[k]);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return tmpColl;
        }

        /// <summary>
        /// Функция сортировки почасовой нагрузки по иерархии: преподаватель, курс, дисциплина
        /// </summary>
        /// <param name="coll"></param>
        /// <returns></returns>
        private IList<clsDistribution> SortHouredLoad(IList<clsDistribution> coll)
        {
            int i, j, k;
            IList<clsDistribution> tmpColl = new List<clsDistribution>();

            //Перебираем учебные семестры
            for (i = 0; i <= mdlData.colSemestr.Count - 1; i++)
            {
                //Перебираем курсы
                for (j = -1; j <= mdlData.colKursNum.Count - 1; j++)
                {
                    //Перебираем нагрузку
                    for (k = 0; k <= mdlData.colHouredDistribution.Count - 1; k++)
                    {
                        if (!mdlData.colHouredDistribution[k].flgExclude)
                        {
                            //Если семестр не указан
                            if (i == 0)
                            {
                                //Если для нагрузки семестр тоже не указан
                                if (mdlData.colHouredDistribution[k].Semestr == null)
                                {
                                    //Планируем строку в нагрузку
                                    tmpColl.Add(mdlData.colHouredDistribution[k]);
                                }
                            }
                            //Если семестр указан
                            else
                            {
                                //Если курс не указан
                                if (j == -1)
                                {
                                    if (mdlData.colHouredDistribution[k].Semestr != null)
                                    {
                                        if (mdlData.colHouredDistribution[k].KursNum == null &
                                            mdlData.colHouredDistribution[k].Semestr.SemNum.Equals(mdlData.colSemestr[i].SemNum))
                                        {
                                            tmpColl.Add(mdlData.colHouredDistribution[k]);
                                        }
                                    }
                                }
                                //Если курс указан
                                else
                                {
                                    if (mdlData.colHouredDistribution[k].Semestr != null &
                                        mdlData.colHouredDistribution[k].KursNum != null)
                                    {
                                        if (mdlData.colHouredDistribution[k].KursNum.Kurs.Equals(mdlData.colKursNum[j].Kurs) &
                                            mdlData.colHouredDistribution[k].Semestr.SemNum.Equals(mdlData.colSemestr[i].SemNum))
                                        {
                                            tmpColl.Add(mdlData.colHouredDistribution[k]);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return tmpColl;
        }

        private void tabFirstSemHoured(object ObjMissing, Word._Document ObjDoc)
        {
            object EndOfDoc = "\\endofdoc";
            int rows;
            int sumI;
            double sumPaymentI;
            int countRow;
            double Payment;

            bool flgPrintSubj;
            bool flgPrintGroup;

            clsDistribution prevRow;

            IList<clsDistribution> coll;            
            Word.Table ObjTable;
            Word.Range ObjWordRange;

            coll = mdlData.colHouredDistribution;

            rows = countTableRowsHoured(coll, cmbLecturerList.SelectedIndex, 2, mdlData.colSemestr[1]);

            //Вставляем таблицу N x 6, заполняем её данными о преподавателе
            //где N - количество строк, записываемов в переменную rows
            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjTable = ObjDoc.Tables.Add(ObjWordRange, rows, 6, ref ObjMissing, ref ObjMissing);

            ObjTable.Borders.Enable = 1;
            ObjTable.Range.ParagraphFormat.SpaceAfter = 0;
            ObjTable.Range.ParagraphFormat.LeftIndent = 0;
            ObjTable.Range.ParagraphFormat.RightIndent = 0;
            ObjTable.Range.Font.Size = 12;
            ObjTable.Range.Font.Bold = 0;
            ObjTable.Range.Font.AllCaps = 0;

            ObjTable.Cell(1, 1).Range.Text = "Дисциплина";
            ObjTable.Cell(1, 2).Range.Text = "Группа";
            ObjTable.Cell(1, 3).Range.Text = "Вид занятий";
            ObjTable.Cell(1, 4).Range.Text = "Кол-во часов";
            ObjTable.Cell(1, 5).Range.Text = "Стоимость часа (руб.)";
            ObjTable.Cell(1, 6).Range.Text = "Общая стоимость оплаты";
            ObjTable.Rows[1].Range.Font.Bold = 1;

            coll = SortHouredLoad(coll);

            countRow = 2;
            sumI = 0;
            sumPaymentI = 0;
            prevRow = null;
            flgPrintSubj = false;
            flgPrintGroup = false;

            for (int i = 0; i <= coll.Count - 1; i++)
            {
                if (coll[i].Semestr.SemNum.Equals(mdlData.colSemestr[1].SemNum))
                {
                    if (!(coll[i].Lecturer == null))
                    {
                        if ((coll[i].Lecturer.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])))
                        {
                            if (coll[i].Lecturer.Duty.Short.Equals("асс."))
                            {
                                Payment = mdlData.PaymentAssist;
                            }
                            else
                            {
                                if (coll[i].Lecturer.Duty.Short.Equals("ст.преп."))
                                {
                                    Payment = mdlData.PaymentStPrep;
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty.Short.Equals("доц."))
                                    {
                                        Payment = mdlData.PaymentDocent;
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty.Short.Equals("проф."))
                                        {
                                            Payment = mdlData.PaymentProff;
                                        }
                                        else
                                        {
                                            Payment = 0d;
                                        }
                                    }
                                }
                            }

                            if (Payment == 0d & coll[i].Lecturer.Duty1 != null)
                            {
                                if (coll[i].Lecturer.Duty1.Short.Equals("асс."))
                                {
                                    Payment = mdlData.PaymentAssist;
                                }
                                else
                                {
                                    if (coll[i].Lecturer.Duty1.Short.Equals("ст.преп."))
                                    {
                                        Payment = mdlData.PaymentStPrep;
                                    }
                                    else
                                    {
                                        if (coll[i].Lecturer.Duty1.Short.Equals("доц."))
                                        {
                                            Payment = mdlData.PaymentDocent;
                                        }
                                        else
                                        {
                                            if (coll[i].Lecturer.Duty1.Short.Equals("проф."))
                                            {
                                                Payment = mdlData.PaymentProff;
                                            }
                                            else
                                            {
                                                Payment = 0d;
                                            }
                                        }
                                    }
                                }
                            }

                            if (prevRow == null)
                            {
                                prevRow = coll[i];
                                flgPrintSubj = true;
                                flgPrintGroup = true;
                            }
                            else
                            {
                                if (prevRow.Subject.Equals(coll[i].Subject))
                                {
                                    flgPrintSubj = false;
                                }
                                else
                                {
                                    flgPrintSubj = true;
                                }

                                if (prevRow.Speciality.Equals(coll[i].Speciality) & prevRow.KursNum.Equals(coll[i].KursNum))
                                {
                                    flgPrintGroup = false;
                                }
                                else
                                {
                                    flgPrintGroup = true;
                                }
                            }

                            //Если есть лекционные часы - добавляем строку
                            if (!(coll[i].Lecture == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "Лекции";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                    flgPrintSubj = false;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                    flgPrintGroup = false;
                                }

                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].Lecture.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].Lecture * Payment).ToString();

                                sumI += coll[i].Lecture;
                                sumPaymentI += (coll[i].Lecture * Payment);
                                countRow++;
                            }

                            //Если есть экзаменационные часы - добавляем строку
                            if (!(coll[i].Exam == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "Экзамен";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                    flgPrintSubj = false;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                    flgPrintGroup = false;
                                }

                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].Exam.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].Exam * Payment).ToString();

                                sumI += coll[i].Exam;
                                sumPaymentI += (coll[i].Exam * Payment);
                                countRow++;
                            }

                            //Если есть зачётные часы - добавляем строку
                            if (!(coll[i].Credit == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "Зачёт";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                    flgPrintSubj = false;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                    flgPrintGroup = false;
                                }
                                
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].Credit.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].Credit * Payment).ToString();

                                sumI += coll[i].Credit;
                                sumPaymentI += (coll[i].Credit * Payment);
                                countRow++;
                            }

                            //Если есть часы на реферат - добавляем строку
                            if (!(coll[i].RefHomeWork == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "Рефераты, домашнее задание";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                    flgPrintSubj = false;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                    flgPrintGroup = false;
                                }

                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].RefHomeWork.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].RefHomeWork * Payment).ToString();

                                sumI += coll[i].RefHomeWork;
                                sumPaymentI += (coll[i].RefHomeWork * Payment);
                                countRow++;
                            }

                            //Если есть часы на консультацию - добавляем строку
                            if (!(coll[i].Tutorial == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "Консультация";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                }
                                
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].Tutorial.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].Tutorial * Payment).ToString();

                                sumI += coll[i].Tutorial;
                                sumPaymentI += (coll[i].Tutorial * Payment);
                                countRow++;
                            }

                            //Если есть часы на лабораторные работы - добавляем строку
                            if (!(coll[i].LabWork == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "Лабораторные работы";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                    flgPrintSubj = false;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                    flgPrintGroup = false;
                                }

                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].LabWork.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].LabWork * Payment).ToString();

                                sumI += coll[i].LabWork;
                                sumPaymentI += (coll[i].LabWork * Payment);
                                countRow++;
                            }

                            //Если есть часы на практические занятия - добавляем строку
                            if (!(coll[i].Practice == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "Практические занятия";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                    flgPrintSubj = false;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                    flgPrintGroup = false;
                                }
                                
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].Practice.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].Practice * Payment).ToString();

                                sumI += coll[i].Practice;
                                sumPaymentI += (coll[i].Practice * Payment);
                                countRow++;
                            }

                            //Если есть часы на индивидуальные занятия - добавляем строку
                            if (!(coll[i].IndividualWork == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "Индивидуальные задания";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                    flgPrintSubj = false;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                    flgPrintGroup = false;
                                }
                                
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].IndividualWork.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].IndividualWork * Payment).ToString();

                                sumI += coll[i].IndividualWork;
                                sumPaymentI += (coll[i].IndividualWork * Payment);
                                countRow++;
                            }

                            //Если есть часы на КРАПК - добавляем строку
                            if (!(coll[i].KRAPK == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "КРАПК";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                    flgPrintSubj = false;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                    flgPrintGroup = false;
                                }

                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].KRAPK.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].KRAPK * Payment).ToString();

                                sumI += coll[i].KRAPK;
                                sumPaymentI += (coll[i].KRAPK * Payment);
                                countRow++;
                            }

                            //Если есть часы на курсовой проект - добавляем строку
                            if (!(coll[i].KursProject == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "Курсовой проект";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                    flgPrintSubj = false;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                    flgPrintGroup = false;
                                }
                                
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].KursProject.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].KursProject * Payment).ToString();

                                sumI += coll[i].KursProject;
                                sumPaymentI += (coll[i].KursProject * Payment);
                                countRow++;
                            }

                            //Если есть часы на преддипломную практику - добавляем строку
                            if (!(coll[i].PreDiplomaPractice == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "Преддипломная практика";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                    flgPrintSubj = false;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                    flgPrintGroup = false;
                                }

                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].PreDiplomaPractice.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].PreDiplomaPractice * Payment).ToString();

                                sumI += coll[i].PreDiplomaPractice;
                                sumPaymentI += (coll[i].PreDiplomaPractice * Payment);
                                countRow++;
                            }

                            //Если есть часы на Дипломный проект - добавляем строку
                            if (!(coll[i].DiplomaPaper == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "Дипломный проект";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                    flgPrintSubj = false;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                    flgPrintGroup = false;
                                }
                                
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].DiplomaPaper.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].DiplomaPaper * Payment).ToString();

                                sumI += coll[i].DiplomaPaper;
                                sumPaymentI += (coll[i].DiplomaPaper * Payment);
                                countRow++;
                            }

                            //Если есть часы на учебную практику - добавляем строку
                            if (!(coll[i].TutorialPractice == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "Учебная практика";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                    flgPrintSubj = false;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                    flgPrintGroup = false;
                                }
                                
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].TutorialPractice.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].TutorialPractice * Payment).ToString();

                                sumI += coll[i].TutorialPractice;
                                sumPaymentI += (coll[i].TutorialPractice * Payment);
                                countRow++;
                            }

                            //Если есть часы на производственную практику - добавляем строку
                            if (!(coll[i].ProducingPractice == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "Производственная практика";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                    flgPrintSubj = false;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                    flgPrintGroup = false;
                                }
                                
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].ProducingPractice.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].ProducingPractice * Payment).ToString();

                                sumI += coll[i].ProducingPractice;
                                sumPaymentI += (coll[i].ProducingPractice * Payment);
                                countRow++;
                            }

                            //Если есть часы на ГАК - добавляем строку
                            if (!(coll[i].GAK == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "ГАК";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                    flgPrintSubj = false;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                    flgPrintGroup = false;
                                }
                                
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].GAK.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].GAK * Payment).ToString();

                                sumI += coll[i].GAK;
                                sumPaymentI += (coll[i].GAK * Payment);
                                countRow++;
                            }

                            //Если есть часы на Посещение занятий - добавляем строку
                            if (!(coll[i].Visiting == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "Посещение занятий";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                    flgPrintSubj = false;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                    flgPrintGroup = false;
                                }
                                
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].Visiting.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].Visiting * Payment).ToString();

                                sumI += coll[i].Visiting;
                                sumPaymentI += (coll[i].Visiting * Payment);
                                countRow++;
                            }

                            //Если есть часы на Аспирантуру - добавляем строку
                            if (!(coll[i].PostGrad == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "Аспирантура";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                    flgPrintSubj = false;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                    flgPrintGroup = false;
                                }
                                
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].PostGrad.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].PostGrad * Payment).ToString();

                                sumI += coll[i].PostGrad;
                                sumPaymentI += (coll[i].PostGrad * Payment);
                                countRow++;
                            }

                            //Если есть часы на Руководство магистрами - добавляем строку
                            if (!(coll[i].Magistry == 0))
                            {
                                ObjTable.Cell(countRow, 3).Range.Text = "Руководство магистерской программой";

                                if (flgPrintSubj)
                                {
                                    ObjTable.Cell(countRow, 1).Range.Text = coll[i].Subject.Subject;
                                    flgPrintSubj = false;
                                }

                                if (flgPrintGroup)
                                {
                                    ObjTable.Cell(countRow, 2).Range.Text = coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + "1";
                                    flgPrintGroup = false;
                                }
                                
                                ObjTable.Cell(countRow, 4).Range.Text = coll[i].Magistry.ToString();
                                ObjTable.Cell(countRow, 5).Range.Text = Payment.ToString();
                                ObjTable.Cell(countRow, 6).Range.Text = (coll[i].Magistry * Payment).ToString();

                                sumI += coll[i].Magistry;
                                sumPaymentI += (coll[i].Magistry * Payment);
                                countRow++;
                            }
                        }
                    }
                }
            }

            ObjTable.Cell(countRow, 1).Range.Text = "Всего за семестр";
            ObjTable.Cell(countRow, 4).Range.Text = sumI.ToString();
            ObjTable.Cell(countRow, 6).Range.Text = sumPaymentI.ToString();
        }

        private void textTabFirstSemHoured(object ObjMissing, Word._Document ObjDoc)
        {
            Word.Paragraph ObjParagraph;
            //Добавляем абзац текста в документ
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Текстуальное содержимое абзаца
            ObjParagraph.Range.Text = "1. Учебная работа.";
            //Размер шрифта 12 пт
            ObjParagraph.Range.Font.Size = 12;
            //Обычный шрифт
            ObjParagraph.Range.Font.Bold = 0;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();

            //Добавляем абзац текста в документ
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Текстуальное содержимое абзаца
            ObjParagraph.Range.Text = "1-й семестр";
            //Размер шрифта 12 пт
            ObjParagraph.Range.Font.Size = 12;
            //Обычный шрифт
            ObjParagraph.Range.Font.Bold = 0;
        }

        private void tabLectAndDepartHoured(object ObjMissing, Word._Document ObjDoc)
        {
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;
            Word.Table ObjTable;
            Word.Range ObjWordRange;

            //Вставляем таблицу 1 x 1, заполняем её данными о преподавателе
            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjTable = ObjDoc.Tables.Add(ObjWordRange, 1, 1, ref ObjMissing, ref ObjMissing);
            ObjTable.Borders.Enable = 1;
            ObjTable.Range.ParagraphFormat.SpaceAfter = 0;
            ObjTable.Range.Font.Size = 12;
            ObjTable.Range.Font.Bold = 0;
            ObjTable.Range.Font.AllCaps = 0;

            ObjTable.Cell(1, 1).Range.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO + ", " +
                                             mdlData.colLecturer[cmbLecturerList.SelectedIndex].Duty.Short + ", " +
                                             (mdlData.colLecturer[cmbLecturerList.SelectedIndex].Duty1 != null ?
                                             mdlData.colLecturer[cmbLecturerList.SelectedIndex].Duty1.Short + ", " :
                                             "") +
                                             mdlData.colLecturer[cmbLecturerList.SelectedIndex].Degree.Short;

            //Добавляем абзац текста в документ
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            //Текстуальное содержимое абзаца
            ObjParagraph.Range.Text = "(ФИО, должность, ученая степень)";
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Размер шрифта 10 пт
            ObjParagraph.Range.Font.Size = 10;
            //Обычный шрифт
            ObjParagraph.Range.Font.Bold = 0;
            //Отменить все заглавные
            ObjParagraph.Range.Font.AllCaps = 0;

            ObjTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            //Вставляем таблицу 1 x 1, заполняем её данными о преподавателе
            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            //
            ObjTable = ObjDoc.Tables.Add(ObjWordRange, 1, 1, ref ObjMissing, ref ObjMissing);
            //
            ObjTable.Borders.Enable = 1;
            //
            ObjTable.Range.ParagraphFormat.SpaceAfter = 0;
            //
            ObjTable.Range.Font.Size = 12;
            //
            ObjTable.Range.Font.Bold = 0;
            //
            ObjTable.Range.Font.AllCaps = 0;
            //Записываем наименование кафедры
            ObjTable.Cell(1, 1).Range.Text = "Кафедра: " + mdlData.DepartmentName;
            ObjTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            //Добавляем абзац текста в документ
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            //Текстуальное содержимое абзаца
            ObjParagraph.Range.Text = "(почасовой фонд: подразделение, вид деятельности)";
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Размер шрифта 10 пт
            ObjParagraph.Range.Font.Size = 10;
            //Обычный шрифт
            ObjParagraph.Range.Font.Bold = 0;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textPlanHeaderHoured(object ObjMissing, Word._Document ObjDoc)
        {
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            //Добавляем абзац текста в документ
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            //Добавляем текст
            ObjParagraph.Range.Text = "Индивидуальный план преподавателя";
            //Размер шрифта 12 пт
            ObjParagraph.Range.Font.Size = 12;
            //Все заглавные
            ObjParagraph.Range.Font.AllCaps = 1;
            //Абзацный отступ после - 0 пунктов
            ObjParagraph.Format.SpaceAfter = 0f;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();

            //Добавляем абзац текста в документ
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            //Добавляем текст
            ObjParagraph.Range.Text = "на условиях почасовой оплаты труда на " +
                mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear.Replace("/", " - ") +
                " учебный год";
            //Отменяем все заглавные
            ObjParagraph.Range.Font.AllCaps = 0;
            //Размер шрифта 10 пт
            ObjParagraph.Range.Font.Size = 12;
            //Добавить абзац
            ObjParagraph.Range.InsertParagraphAfter();
            
            //Добавляем абзац текста в документ
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjMissing);
            //Добавляем текст
            ObjParagraph.Range.Text = "по определяемой Учебным управлением учебной нагрузке";

            //Размер шрифта 12 пт
            ObjParagraph.Range.Font.Size = 12;
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Одинарный межстрочный интервал
            ObjParagraph.Format.Space1();
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Отступ слева
            ObjParagraph.Format.LeftIndent = ObjDoc.Application.CentimetersToPoints(1f);
            //Отступ справа
            ObjParagraph.Format.RightIndent = ObjDoc.Application.CentimetersToPoints(1f);
            //Абзацный отступ после - 0 пунктов
            ObjParagraph.Format.SpaceAfter = 0f;
            //Добавляем текст
            ObjParagraph.Range.InsertAfter("\n\nс «    »          " +
                mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear.Substring(0, mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear.IndexOf('/')) +
                " г. по «    »          " +
               mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear.Substring(mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear.IndexOf('/') + 1) + " г.  ");
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Размер шрифта 12 пт
            ObjParagraph.Range.Font.Size = 12;
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            //Отступ слева
            ObjParagraph.Format.LeftIndent = ObjDoc.Application.CentimetersToPoints(-1.27f);
            //Отступ справа
            ObjParagraph.Format.RightIndent = ObjDoc.Application.CentimetersToPoints(-0.64f);
        }

        private void textMinistryHeaderHoured(object ObjMissing, Word._Document ObjDoc)
        {
            object EndOfDoc = "\\endofdoc";
            object ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            Word.Paragraph ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            //Пишем наименование ведомственной принадлежности
            //(министерство)
            ObjParagraph.Range.Text = mdlData.MinistryName;
            //Одинарный межстрочный интервал
            ObjParagraph.Format.Space1();
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Все заглавные
            ObjParagraph.Range.Font.AllCaps = 1;
            //Размер шрифта 10 пт
            ObjParagraph.Range.Font.Size = 10;
            //Абзацный отступ после - 10 пунктов
            ObjParagraph.Format.SpaceAfter = 10f;
            //Отступ слева
            ObjParagraph.Format.LeftIndent = ObjDoc.Application.CentimetersToPoints(1.5f);
            //Отступ справа
            ObjParagraph.Format.RightIndent = ObjDoc.Application.CentimetersToPoints(1.5f);
            //Добавляем ещё один абзац
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textTitleHeaderHoured(object ObjMissing, Word._Document ObjDoc)
        {
            object EndOfDoc = "\\endofdoc";
            object ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            Word.Paragraph ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            //Пишем наименование префикса Университета
            //ФГБОУ ВО
            ObjParagraph.Range.Text = mdlData.UniversityPrefName;
            //Абзацный отступ после - 0 пунктов
            ObjParagraph.Format.SpaceAfter = 0f;
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textUniversityHeaderHoured(object ObjMissing, Word._Document ObjDoc)
        {
            object EndOfDoc = "\\endofdoc";
            object ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            Word.Paragraph ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            //Пишем наименование университета
            ObjParagraph.Range.Text = mdlData.UniversityName.Replace("\r\n","").Replace("«", "\"").Replace("»", "\"");
            //Размер шрифта 12 пт
            ObjParagraph.Range.Font.Size = 12;
            ObjParagraph.Range.InsertParagraphAfter();

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            //Добавляем абзац текста в документ
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            //Пишем аббревиатуру университета
            ObjParagraph.Range.Text = mdlData.UniversitySuffName;
            //Абзацный отступ после - 12 пунктов
            ObjParagraph.Format.SpaceAfter = 12f;
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textAgreed(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "Утверждаю:";
            //Выравнивание по левому краю
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //Снимаем выделение жирным
            ObjParagraph.Range.Font.Bold = 0;
            //Отступ слева в 9,5 см
            ObjParagraph.Format.LeftIndent = 9.5f / 0.03527f;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textDirector(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            //Вставляем абзац в конец документа
            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "Директор института ИТТСУ";
            //Снимаем все заглавные
            ObjParagraph.Range.Font.AllCaps = 0;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textDirectorFIO(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "П.Ф. Бестемьянов";
            //Отступ слева в 11,5 см
            ObjParagraph.Format.LeftIndent = 11.5f / 0.03527f;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textMain(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "Заведующий кафедрой \"УиЗИ\"";
            //Отступ слева в 9,5 см
            ObjParagraph.Format.LeftIndent = 9.5f / 0.03527f;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textMainFIO(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "Л.А. Баранов";
            //Отступ слева в 11,5 см
            ObjParagraph.Format.LeftIndent = 11.5f / 0.03527f;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textDepartName(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "Кафедра \"Управление и защита информации\"";
            //Отступа слева нет
            ObjParagraph.Format.LeftIndent = 0;
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Все обычные, не заглавные
            ObjParagraph.Range.Font.AllCaps = 0;
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textLecturer(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "преподавателя";
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 0;
            //Курсивный шрифт
            ObjParagraph.Range.Font.Italic = 1;
            //Все обычные, не заглавные
            ObjParagraph.Range.Font.AllCaps = 1;
            //Размер шрифта 18 пт
            ObjParagraph.Range.Font.Size = 18;
            //Разреживания шрифта нет
            ObjParagraph.Range.Font.Spacing = 0;
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textBoss(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "заведующего кафедрой";
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Курсивный шрифт
            ObjParagraph.Range.Font.Italic = 1;
            //Все обычные, не заглавные
            ObjParagraph.Range.Font.AllCaps = 1;
            //Размер шрифта 18 пт
            ObjParagraph.Range.Font.Size = 18;
            //Разреживания шрифта нет
            ObjParagraph.Range.Font.Spacing = 0;
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void spacesBefore(object ObjMissing, Word._Document ObjDoc, int i)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;
            int k;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            
            for (k = 1; k < i; k++)
            {
                //Добавляем ещё один абзац текста
                ObjParagraph.Range.InsertParagraphBefore();
            }
        }

        private void spacesAfter(object ObjMissing, Word._Document ObjDoc, int i)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;
            int k;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);

            for (k = 1; k < i; k++)
            {
                //Добавляем ещё один абзац текста
                ObjParagraph.Range.InsertParagraphAfter();
            }
        }

        private void textPlan(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "Индивидуальный  план  работы";
            //Устанавливаем все заглавные
            ObjParagraph.Range.Font.AllCaps = 1;
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Размер шрифта 20 пт
            ObjParagraph.Range.Font.Size = 20;
            //Разреживание шрифта на 2 пт
            ObjParagraph.Range.Font.Spacing = 2;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textPlanYear(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "на " + mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear + " учебный год";
            //Сбрасываем все заглавные
            ObjParagraph.Range.Font.AllCaps = 0;
            //Разреживание шрифта убираем
            ObjParagraph.Range.Font.Spacing = 0;
            //Отменяем жирный шрифт
            ObjParagraph.Range.Font.Bold = 0;
            //Размер шрифта 12 пт
            ObjParagraph.Range.Font.Size = 14;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textPlanYearNew(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "на " + mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear + " учебный год";
            //Сбрасываем все заглавные
            ObjParagraph.Range.Font.AllCaps = 0;
            //Разреживание шрифта убираем
            ObjParagraph.Range.Font.Spacing = 0;
            //Отменяем жирный шрифт
            ObjParagraph.Range.Font.Bold = 0;
            //Отменяем курсивный шрифт
            ObjParagraph.Range.Font.Italic = 0;
            //Размер шрифта 12 пт
            ObjParagraph.Range.Font.Size = 14;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textLecturerInfo(object ObjMissing, Word._Document ObjDoc)
        {
            string RateAbout;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Table ObjTable;
            Word.Range ObjWordRange;

            //Вставляем таблицу 5 x 2, заполняем её данными о преподавателе
            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjTable = ObjDoc.Tables.Add(ObjWordRange, 5, 2, ref ObjMissing, ref ObjMissing);
            ObjTable.Range.ParagraphFormat.SpaceAfter = 0;

            ObjTable.Cell(1, 1).Range.Text = "Фамилия, Имя, Отчество";
            ObjTable.Cell(1, 1).Width = 7.19f / 0.03527f;
            ObjTable.Cell(1, 1).Range.Font.Bold = 1;
            ObjTable.Cell(1, 1).Range.Font.Size = 12;
            ObjTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(1, 1).Range.ParagraphFormat.LeftIndent = 1.25f / 0.03527f;
            ObjTable.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 12;

            ObjTable.Cell(1, 2).Range.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO;
            ObjTable.Cell(1, 2).Width = 9.75f / 0.03527f;
            ObjTable.Cell(1, 2).Range.Font.Bold = 0;
            ObjTable.Cell(1, 2).Range.Font.Size = 14;
            ObjTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(1, 2).Range.ParagraphFormat.LeftIndent = 0;
            ObjTable.Cell(1, 2).Range.ParagraphFormat.SpaceAfter = 12;

            ObjTable.Cell(2, 1).Range.Text = "Должность";
            ObjTable.Cell(2, 1).Width = 7.19f / 0.03527f;
            ObjTable.Cell(2, 1).Range.Font.Bold = 1;
            ObjTable.Cell(2, 1).Range.Font.Size = 12;
            ObjTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(2, 1).Range.ParagraphFormat.LeftIndent = 3.75f / 0.03527f;
            ObjTable.Cell(2, 1).Range.ParagraphFormat.SpaceAfter = 12;

            ObjTable.Cell(2, 2).Range.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Duty.Duty;
            ObjTable.Cell(2, 2).Width = 9.75f / 0.03527f;
            ObjTable.Cell(2, 2).Range.Font.Bold = 0;
            ObjTable.Cell(2, 2).Range.Font.Size = 14;
            ObjTable.Cell(2, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(2, 2).Range.ParagraphFormat.LeftIndent = 0;
            ObjTable.Cell(2, 2).Range.ParagraphFormat.SpaceAfter = 12;

            ObjTable.Cell(3, 1).Range.Text = "Ученая степень";
            ObjTable.Cell(3, 1).Width = 7.19f / 0.03527f;
            ObjTable.Cell(3, 1).Range.Font.Bold = 1;
            ObjTable.Cell(3, 1).Range.Font.Size = 12;
            ObjTable.Cell(3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(3, 1).Range.ParagraphFormat.LeftIndent = 3.75f / 0.03527f;
            ObjTable.Cell(3, 1).Range.ParagraphFormat.SpaceAfter = 12;

            ObjTable.Cell(3, 2).Range.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Degree.Short;
            ObjTable.Cell(3, 2).Width = 9.75f / 0.03527f;
            ObjTable.Cell(3, 2).Range.Font.Bold = 0;
            ObjTable.Cell(3, 2).Range.Font.Size = 14;
            ObjTable.Cell(3, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(3, 2).Range.ParagraphFormat.LeftIndent = 0;
            ObjTable.Cell(3, 1).Range.ParagraphFormat.SpaceAfter = 12;

            ObjTable.Cell(4, 1).Range.Text = "Ученое звание";
            ObjTable.Cell(4, 1).Width = 7.19f / 0.03527f;
            ObjTable.Cell(4, 1).Range.Font.Bold = 1;
            ObjTable.Cell(4, 1).Range.Font.Size = 12;
            ObjTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(4, 1).Range.ParagraphFormat.LeftIndent = 3.75f / 0.03527f;
            ObjTable.Cell(4, 1).Range.ParagraphFormat.SpaceAfter = 12;

            ObjTable.Cell(4, 2).Range.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Status.Status;
            ObjTable.Cell(4, 2).Width = 9.75f / 0.03527f;
            ObjTable.Cell(4, 2).Range.Font.Bold = 0;
            ObjTable.Cell(4, 2).Range.Font.Size = 14;
            ObjTable.Cell(4, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(4, 2).Range.ParagraphFormat.LeftIndent = 0;
            ObjTable.Cell(4, 2).Range.ParagraphFormat.SpaceAfter = 12;

            ObjTable.Cell(5, 1).Range.Text = "Совместитель";
            ObjTable.Cell(5, 1).Width = 7.19f / 0.03527f;
            ObjTable.Cell(5, 1).Range.Font.Bold = 1;
            ObjTable.Cell(5, 1).Range.Font.Size = 12;
            ObjTable.Cell(5, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(5, 1).Range.ParagraphFormat.LeftIndent = 3.75f / 0.03527f;
            ObjTable.Cell(5, 1).Range.ParagraphFormat.SpaceAfter = 12;

            if (mdlData.colLecturer[cmbLecturerList.SelectedIndex].Rate == 1)
            {
                RateAbout = "ставка";
            }
            else
            {
                RateAbout = "ставки";
            }

            ObjTable.Cell(5, 2).Range.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Combination.CombType + "   " +
                                             mdlData.colLecturer[cmbLecturerList.SelectedIndex].Rate + " " + RateAbout;
            ObjTable.Cell(5, 2).Width = 9.75f / 0.03527f;
            ObjTable.Cell(5, 2).Range.Font.Bold = 0;
            ObjTable.Cell(5, 2).Range.Font.Size = 14;
            ObjTable.Cell(5, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(5, 2).Range.ParagraphFormat.LeftIndent = 0;
            ObjTable.Cell(5, 2).Range.ParagraphFormat.SpaceAfter = 12;
        }

        private void textLecturerInfoNew(object ObjMissing, Word._Document ObjDoc)
        {
            string RateAbout;
            string Combination1;
            string Combination2;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Table ObjTable;
            Word.Range ObjWordRange;

            //Вставляем таблицу 5 x 2, заполняем её данными о преподавателе
            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjTable = ObjDoc.Tables.Add(ObjWordRange, 6, 2, ref ObjMissing, ref ObjMissing);
            ObjTable.Range.ParagraphFormat.SpaceAfter = 0;

            ObjTable.Cell(1, 1).Range.Text = "Фамилия, Имя, Отчество";
            ObjTable.Cell(1, 1).Width = 7.19f / 0.03527f;
            ObjTable.Cell(1, 1).Range.Font.Bold = 1;
            ObjTable.Cell(1, 1).Range.Font.Size = 12;
            ObjTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(1, 1).Range.ParagraphFormat.LeftIndent = 1.25f / 0.03527f;
            ObjTable.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 12;

            ObjTable.Cell(1, 2).Range.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO;
            ObjTable.Cell(1, 2).Width = 9.75f / 0.03527f;
            ObjTable.Cell(1, 2).Range.Font.Bold = 0;
            ObjTable.Cell(1, 2).Range.Font.Size = 14;
            ObjTable.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(1, 2).Range.ParagraphFormat.LeftIndent = 0;
            ObjTable.Cell(1, 2).Range.ParagraphFormat.SpaceAfter = 12;

            ObjTable.Cell(2, 1).Range.Text = "Должность";
            ObjTable.Cell(2, 1).Width = 7.19f / 0.03527f;
            ObjTable.Cell(2, 1).Range.Font.Bold = 1;
            ObjTable.Cell(2, 1).Range.Font.Size = 12;
            ObjTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(2, 1).Range.ParagraphFormat.LeftIndent = 3.75f / 0.03527f;
            ObjTable.Cell(2, 1).Range.ParagraphFormat.SpaceAfter = 12;

            ObjTable.Cell(2, 2).Range.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Duty.Duty;
            ObjTable.Cell(2, 2).Width = 9.75f / 0.03527f;
            ObjTable.Cell(2, 2).Range.Font.Bold = 0;
            ObjTable.Cell(2, 2).Range.Font.Size = 14;
            ObjTable.Cell(2, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(2, 2).Range.ParagraphFormat.LeftIndent = 0;
            ObjTable.Cell(2, 2).Range.ParagraphFormat.SpaceAfter = 12;

            ObjTable.Cell(3, 1).Range.Text = "Ученая степень";
            ObjTable.Cell(3, 1).Width = 7.19f / 0.03527f;
            ObjTable.Cell(3, 1).Range.Font.Bold = 1;
            ObjTable.Cell(3, 1).Range.Font.Size = 12;
            ObjTable.Cell(3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(3, 1).Range.ParagraphFormat.LeftIndent = 3.75f / 0.03527f;
            ObjTable.Cell(3, 1).Range.ParagraphFormat.SpaceAfter = 12;

            ObjTable.Cell(3, 2).Range.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Degree.Short;
            ObjTable.Cell(3, 2).Width = 9.75f / 0.03527f;
            ObjTable.Cell(3, 2).Range.Font.Bold = 0;
            ObjTable.Cell(3, 2).Range.Font.Size = 14;
            ObjTable.Cell(3, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(3, 2).Range.ParagraphFormat.LeftIndent = 0;
            ObjTable.Cell(3, 1).Range.ParagraphFormat.SpaceAfter = 12;

            ObjTable.Cell(4, 1).Range.Text = "Ученое звание";
            ObjTable.Cell(4, 1).Width = 7.19f / 0.03527f;
            ObjTable.Cell(4, 1).Range.Font.Bold = 1;
            ObjTable.Cell(4, 1).Range.Font.Size = 12;
            ObjTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(4, 1).Range.ParagraphFormat.LeftIndent = 3.75f / 0.03527f;
            ObjTable.Cell(4, 1).Range.ParagraphFormat.SpaceAfter = 12;

            ObjTable.Cell(4, 2).Range.Text = mdlData.colLecturer[cmbLecturerList.SelectedIndex].Status.Status;
            ObjTable.Cell(4, 2).Width = 9.75f / 0.03527f;
            ObjTable.Cell(4, 2).Range.Font.Bold = 0;
            ObjTable.Cell(4, 2).Range.Font.Size = 14;
            ObjTable.Cell(4, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(4, 2).Range.ParagraphFormat.LeftIndent = 0;
            ObjTable.Cell(4, 2).Range.ParagraphFormat.SpaceAfter = 12;

            if (mdlData.colLecturer[cmbLecturerList.SelectedIndex].Rate == 1)
            {
                RateAbout = "ставка";
            }
            else
            {
                RateAbout = "ставки";
            }

            if (mdlData.colLecturer[cmbLecturerList.SelectedIndex].Combination.CombType == "нет" ||
                mdlData.colLecturer[cmbLecturerList.SelectedIndex].Combination.CombType == "-")
            {
                Combination1 = "";
                Combination2 = "Штатный  " + mdlData.colLecturer[cmbLecturerList.SelectedIndex].Rate + "  " + RateAbout;
            }
            else
            {
                if (mdlData.colLecturer[cmbLecturerList.SelectedIndex].Combination.CombType == "внутренний")
                {
                    Combination1 = "Совместитель внутренний  " + mdlData.colLecturer[cmbLecturerList.SelectedIndex].Rate + "  " + RateAbout;
                    Combination2 = "Совместитель внешний  ......" + "  " + RateAbout;
                }
                else
                {
                    Combination1 = "Совместитель внутренний  ......" + "  " + RateAbout;
                    Combination2 = "Совместитель внешний  " + mdlData.colLecturer[cmbLecturerList.SelectedIndex].Rate + "  " + RateAbout;
                }
            }

            ObjTable.Cell(5, 1).Range.Text = Combination1;
            ObjTable.Cell(5, 1).Width = 7.19f / 0.03527f;
            ObjTable.Cell(5, 2).Width = 9.75f / 0.03527f;

            ObjTable.Cell(5, 1).Merge(ObjTable.Cell(5, 2));

            ObjTable.Cell(5, 1).Range.Font.Bold = 1;
            ObjTable.Cell(5, 1).Range.Font.Size = 12;
            ObjTable.Cell(5, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(5, 1).Range.ParagraphFormat.LeftIndent = 1.25f / 0.03527f;
            ObjTable.Cell(5, 1).Range.ParagraphFormat.SpaceAfter = 12;

            ObjTable.Cell(6, 1).Range.Text = Combination2;
            ObjTable.Cell(6, 1).Width = 7.19f / 0.03527f;
            ObjTable.Cell(6, 2).Width = 9.75f / 0.03527f;

            ObjTable.Cell(6, 1).Merge(ObjTable.Cell(6, 2));

            ObjTable.Cell(6, 1).Range.Font.Bold = 1;
            ObjTable.Cell(6, 1).Range.Font.Size = 12;
            ObjTable.Cell(6, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(6, 1).Range.ParagraphFormat.LeftIndent = 1.25f / 0.03527f;
            ObjTable.Cell(6, 1).Range.ParagraphFormat.SpaceAfter = 12;
        }

        private void textAgreedInfoBoss(object ObjMissing, Word._Document ObjDoc)
        {
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Table ObjTable;
            Word.Range ObjWordRange;

            //Вставляем таблицу 5 x 3, заполняем её данными о преподавателе
            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjTable = ObjDoc.Tables.Add(ObjWordRange, 5, 3, ref ObjMissing, ref ObjMissing);
            ObjTable.Range.ParagraphFormat.SpaceAfter = 0;
            ObjTable.Range.Font.Size = 12;
            ObjTable.Range.Font.Bold = 0;
            ObjTable.Range.Font.AllCaps = 0;

            ObjTable.Cell(1, 1).Range.Text = "СОГЛАСОВАНО";
            ObjTable.Cell(1, 1).Width = 6.5f / 0.03527f;
            ObjTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            ObjTable.Cell(2, 1).Range.Text = "Директор института ИТТСУ";
            ObjTable.Cell(2, 1).Width = 6.5f / 0.03527f;
            ObjTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            ObjTable.Cell(3, 1).Range.Text = "П.Ф. Бестемьянов";
            ObjTable.Cell(3, 1).Width = 6.5f / 0.03527f;
            ObjTable.Cell(3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            ObjTable.Cell(4, 1).Range.Text = "/                              /";
            ObjTable.Cell(4, 1).Width = 6.5f / 0.03527f;
            ObjTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            ObjTable.Cell(5, 1).Range.Text = "";
            ObjTable.Cell(5, 1).Width = 6.5f / 0.03527f;
            ObjTable.Cell(5, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            ObjTable.Cell(1, 2).Width = 3.75f / 0.03527f;
            ObjTable.Cell(2, 2).Width = 3.75f / 0.03527f;
            ObjTable.Cell(3, 2).Width = 3.75f / 0.03527f;
            ObjTable.Cell(4, 2).Width = 3.75f / 0.03527f;
            ObjTable.Cell(5, 2).Width = 3.75f / 0.03527f;

            ObjTable.Cell(1, 3).Range.Text = "УТВЕРЖДАЮ";
            ObjTable.Cell(1, 3).Width = 6.75f / 0.03527f;
            ObjTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            ObjTable.Cell(2, 3).Range.Text = "Первый проректор -";
            ObjTable.Cell(2, 3).Width = 6.75f / 0.03527f;
            ObjTable.Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            ObjTable.Cell(3, 3).Range.Text = "проректор по учебной работе";
            ObjTable.Cell(3, 3).Width = 6.75f / 0.03527f;
            ObjTable.Cell(3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            ObjTable.Cell(4, 3).Range.Text = "В.В. Виноградов";
            ObjTable.Cell(4, 3).Width = 6.75f / 0.03527f;
            ObjTable.Cell(4, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            ObjTable.Cell(5, 3).Range.Text = "/                              /";
            ObjTable.Cell(5, 3).Width = 6.75f / 0.03527f;
            ObjTable.Cell(5, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
        }

        private void textAgreedInfoStd(object ObjMissing, Word._Document ObjDoc)
        {
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Table ObjTable;
            Word.Range ObjWordRange;

            //Вставляем таблицу 5 x 3, заполняем её данными о преподавателе
            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjTable = ObjDoc.Tables.Add(ObjWordRange, 5, 3, ref ObjMissing, ref ObjMissing);
            ObjTable.Range.ParagraphFormat.SpaceAfter = 0;
            ObjTable.Range.Font.Size = 12;
            ObjTable.Range.Font.Bold = 0;
            ObjTable.Range.Font.AllCaps = 0;

            ObjTable.Cell(1, 1).Range.Text = "СОГЛАСОВАНО";
            ObjTable.Cell(1, 1).Width = 6.5f / 0.03527f;
            ObjTable.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            ObjTable.Cell(2, 1).Range.Text = "Заведующий кафедрой";
            ObjTable.Cell(2, 1).Width = 6.5f / 0.03527f;
            ObjTable.Cell(2, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            ObjTable.Cell(3, 1).Range.Text = "Л.А. Баранов";
            ObjTable.Cell(3, 1).Width = 6.5f / 0.03527f;
            ObjTable.Cell(3, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            ObjTable.Cell(4, 1).Range.Text = "/                              /";
            ObjTable.Cell(4, 1).Width = 6.5f / 0.03527f;
            ObjTable.Cell(4, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            ObjTable.Cell(5, 1).Range.Text = "";
            ObjTable.Cell(5, 1).Width = 6.5f / 0.03527f;
            ObjTable.Cell(5, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            ObjTable.Cell(1, 2).Width = 3.75f / 0.03527f;
            ObjTable.Cell(2, 2).Width = 3.75f / 0.03527f;
            ObjTable.Cell(3, 2).Width = 3.75f / 0.03527f;
            ObjTable.Cell(4, 2).Width = 3.75f / 0.03527f;
            ObjTable.Cell(5, 2).Width = 3.75f / 0.03527f;

            ObjTable.Cell(1, 3).Range.Text = "УТВЕРЖДАЮ";
            ObjTable.Cell(1, 3).Width = 6.75f / 0.03527f;
            ObjTable.Cell(1, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            ObjTable.Cell(2, 3).Range.Text = "Директор института ИТТСУ";
            ObjTable.Cell(2, 3).Width = 6.75f / 0.03527f;
            ObjTable.Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            ObjTable.Cell(3, 3).Range.Text = "П.Ф. Бестемьянов";
            ObjTable.Cell(3, 3).Width = 6.75f / 0.03527f;
            ObjTable.Cell(3, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            ObjTable.Cell(4, 3).Range.Text = "/                              /";
            ObjTable.Cell(4, 3).Width = 6.75f / 0.03527f;
            ObjTable.Cell(4, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            ObjTable.Cell(5, 3).Range.Text = "";
            ObjTable.Cell(5, 3).Width = 6.75f / 0.03527f;
            ObjTable.Cell(5, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
        }

        private void textReminder(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "Сроки предоставления индивидуальных планов в Учебно-методическое управление:";
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //Размер шрифта 10 пт
            ObjParagraph.Range.Font.Size = 10;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertAfter("- до 1 июля, но не позднее 1 сентября – объем плановой нагрузки преподавателя;");
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertAfter("- до 1 февраля – отчет о фактической нагрузке за 1 семестр;");
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertAfter("- до 1 июля – отчет о фактической нагрузке за учебный год.");
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textReminderNew(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "Сроки предоставления индивидуальных планов в Учебно-методическое управление:";
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //Размер шрифта 10 пт
            ObjParagraph.Range.Font.Size = 10;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertAfter("- до 1 сентября – отчёт по планируемой учебной нагрузке, с возможной коррекцией по итогам " +
                                           "нового набора до 15 октября;");
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertAfter("- до 1 февраля – отчет о фактической работе за 1 семестр;");
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            ObjParagraph.Range.InsertAfter("- до 1 июля – отчет о фактической работе за учебный год.");
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textMethodAgreed(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "Согласовано:";
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //Размер шрифта 10 пт
            ObjParagraph.Range.Font.Size = 10;
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Отменяем все заглавные
            ObjParagraph.Range.Font.AllCaps = 0;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertAfter("Заместитель начальника УМУ\t_________________________ С.С. Андриянов");
            ObjParagraph.TabStops.Add(17f / 0.03527f, Word.WdTabAlignment.wdAlignTabRight);
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textScienceAgreed(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "Согласовано:";
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //Размер шрифта 10 пт
            ObjParagraph.Range.Font.Size = 10;
            //Отменяем жирный шрифт
            ObjParagraph.Range.Font.Bold = 0;
            //Отменяем все заглавные
            ObjParagraph.Range.Font.AllCaps = 0;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertAfter("Начальник Управления научно-исследовательской работы\t______________________________ А.В. Саврухин");
            ObjParagraph.TabStops.Add(17f / 0.03527f, Word.WdTabAlignment.wdAlignTabRight);
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void pageBreaker(object ObjMissing, Word._Document ObjDoc)
        {
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Range ObjWordRange;
            object ObjCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            object ObjPageBreak = Word.WdBreakType.wdPageBreak;

            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjWordRange.Collapse(ref ObjCollapseEnd);
            ObjWordRange.InsertBreak(ref ObjPageBreak);
            ObjWordRange.Collapse(ref ObjCollapseEnd);
        }

        private void textStudWork(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "Учебная работа";
            //Устанавливаем все заглавные
            ObjParagraph.Range.Font.AllCaps = 1;
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Размер шрифта 14 пт
            ObjParagraph.Range.Font.Size = 14;
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textStudWorkFromGrid(object ObjMissing, Word._Document ObjDoc)
        {
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Table ObjTable;
            Word.Range ObjWordRange;

            //Вставляем таблицу согласно заполненной сетке и заполняем её данными о нагрузке
            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjTable = ObjDoc.Tables.Add(ObjWordRange, dgNagruzka.RowCount, dgNagruzka.ColumnCount, ref ObjMissing, ref ObjMissing);
            //Сбрасываем все заглавные
            ObjTable.Range.Font.AllCaps = 0;
            //Убираем жирный шрифт
            ObjTable.Range.Font.Bold = 0;
            //Размер шрифта 10 пт
            ObjTable.Range.Font.Size = 10;
            //Выравнивание по левому краю
            ObjTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //Отступ после абзаца отсутствует
            ObjTable.Range.ParagraphFormat.SpaceAfter = 0;
            //Границы таблицы включены
            ObjTable.Borders.Enable = 1;

            ObjTable.Columns[1].Width = 5.29f / 0.03527f;
            ObjTable.Columns[2].Width = 2.49f / 0.03527f;
            ObjTable.Columns[3].Width = 2.19f / 0.03527f;
            ObjTable.Columns[4].Width = 2.65f / 0.03527f;
            ObjTable.Columns[5].Width = 1.97f / 0.03527f;
            ObjTable.Columns[6].Width = 2.25f / 0.03527f;

            for (int i = 0; i <= dgNagruzka.RowCount - 1; i++)
            {
                for (int j = 0; j <= dgNagruzka.ColumnCount - 1; j++)
                {
                    if (!(dgNagruzka[j, i].Value == null))
                    {
                        ObjTable.Cell(i + 1, j + 1).Range.Text = dgNagruzka[j, i].Value.ToString();
                    }
                }
            }
        }

        private void textStudWorkFromCol(object ObjMissing, Word._Document ObjDoc)
        {
            int countRow;
            int countCol;

            IList<clsDistribution> coll;
            
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Table ObjTable;
            Word.Range ObjWordRange;

            if (flgCombine)
            {
                coll = mdlData.colCombineDistribution;
            }
            else
            {
                coll = mdlData.colDistribution;
            }

            //Очищаем переменную, содержащую возможные названия дисциплин,
            //связанных с дипломным проектированием
            NamesDiploma = new string[0];
            NamesDiploma = findSameNamesDiploma(mdlData.colDistribution);

            GroupsDiploma = new string[0];
            GroupsFullDiploma = new string[0];

            findSameGroupsDiploma(coll, ref GroupsDiploma, ref GroupsFullDiploma);

            NamesMagistry = new string[0];
            NamesMagistry = findSameNamesMagistry(mdlData.colDistribution);

            GroupsMagistry = new string[0];
            GroupsFullMagistry = new string[0];

            findSameGroupsMagistry(coll, ref GroupsMagistry, ref GroupsFullMagistry);

            NamesTutorialPr = new string[0];
            NamesTutorialPr = findSameNamesTutorialPr(mdlData.colDistribution);

            GroupsTutorialPr = new string[0];
            GroupsFullTutorialPr = new string[0];

            findSameGroupsTutorialPr(coll, ref GroupsTutorialPr, ref GroupsFullTutorialPr);

            NamesProdPr = new string[0];
            NamesProdPr = findSameNamesProdPr(mdlData.colDistribution);

            GroupsProdPr = new string[0];
            GroupsFullProdPr = new string[0];

            findSameGroupsProdPr(coll, ref GroupsProdPr, ref GroupsFullProdPr);

            countCol = 6;
            countRow = countTableRows(coll, cmbLecturerList.SelectedIndex);

            //Вставляем таблицу согласно заполненной сетке и заполняем её данными о нагрузке
            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjTable = ObjDoc.Tables.Add(ObjWordRange, countRow, countCol, ref ObjMissing, ref ObjMissing);

            //Сбрасываем все заглавные
            ObjTable.Range.Font.AllCaps = 0;
            //Убираем жирный шрифт
            ObjTable.Range.Font.Bold = 0;
            //Размер шрифта 10 пт
            ObjTable.Range.Font.Size = 10;
            //Выравнивание по левому краю
            ObjTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Отступ после абзаца отсутствует
            ObjTable.Range.ParagraphFormat.SpaceAfter = 0;
            //Границы таблицы включены
            ObjTable.Borders.Enable = 1;

            WordTableFiller(coll, ObjTable, ObjMissing, ObjDoc);

            ObjTable.Columns[1].Width = 5.29f / 0.03527f;
            ObjTable.Columns[2].Width = 2.49f / 0.03527f;
            ObjTable.Columns[3].Width = 2.19f / 0.03527f;
            ObjTable.Columns[4].Width = 2.65f / 0.03527f;
            ObjTable.Columns[5].Width = 1.97f / 0.03527f;
            ObjTable.Columns[6].Width = 2.25f / 0.03527f;
        }

        private void textMethodWork(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "ii. Учебно-методическая работа";
            //Устанавливаем все заглавные
            ObjParagraph.Range.Font.AllCaps = 1;
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Размер шрифта 14 пт
            ObjParagraph.Range.Font.Size = 14;
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textMethodWorkFromCol(object ObjMissing, Word._Document ObjDoc)
        {
            string WorkYear;
            string FirstYear;
            string SecondYear;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Table ObjTable;
            Word.Range ObjWordRange;

            //Вставляем фиксированную таблицу 3 x 4
            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjTable = ObjDoc.Tables.Add(ObjWordRange, 3, 4, ref ObjMissing, ref ObjMissing);
            //Сбрасываем все заглавные
            ObjTable.Range.Font.AllCaps = 0;
            //Убираем жирный шрифт
            ObjTable.Range.Font.Bold = 0;
            //Размер шрифта 10 пт
            ObjTable.Range.Font.Size = 10;
            //Отступ после абзаца отсутствует
            ObjTable.Range.ParagraphFormat.SpaceAfter = 0;
            //Выравнивание в ячейках по левому краю
            ObjTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //Границы таблицы включены
            ObjTable.Borders.Enable = 1;

            ObjTable.Columns[1].Width = 7.98f / 0.03527f;
            ObjTable.Columns[2].Width = 3.07f / 0.03527f;
            ObjTable.Columns[2].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            ObjTable.Columns[3].Width = 2.84f / 0.03527f;
            ObjTable.Columns[3].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            ObjTable.Columns[4].Width = 2.68f / 0.03527f;
            ObjTable.Columns[4].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            ObjTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            ObjTable.Cell(1, 1).Range.Text = "Наименование работы";
            ObjTable.Cell(1, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            ObjTable.Cell(1, 2).Range.Text = "Объём";
            ObjTable.Cell(1, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            ObjTable.Cell(1, 3).Range.Text = "Сроки";
            ObjTable.Cell(1, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            ObjTable.Cell(1, 4).Range.Text = "Отметки о выполнении";
            ObjTable.Cell(1, 4).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            ObjTable.Rows[2].Height = 12f / 0.03527f;
            ObjTable.Rows[3].Height = 12f / 0.03527f;

            WorkYear = mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear;

            FirstYear = WorkYear.Substring(0, 4);
            SecondYear = WorkYear.Substring(5, 4);

            for (int i = 0; i <= mdlData.colDopWork.Count - 1; i++)
            {
                if (mdlData.colDopWork[i].Lecturer.FIO.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO))
                {
                    //Если первый семестр, то УМР выполняется с 1 сентября по 6 февраля
                    if (mdlData.colDopWork[i].Semestr.SemNum.Equals(mdlData.colSemestr[1].SemNum))
                    {
                        //Пишем наименование работ
                        ObjTable.Cell(2, 1).Range.Text = mdlData.colDopWork[i].UMR.ToString();

                        //Пишем объём работы, если он не задан
                        if (!mdlData.colDopWork[i].VolumeUMR.ToString().Equals(""))
                        {
                            ObjTable.Cell(2, 2).Range.Text = mdlData.colDopWork[i].VolumeUMR.ToString();
                        }
                        
                        //Пишем стандартный срок выполнения, если он не задан
                        if (mdlData.colDopWork[i].DateUMR.ToString().Equals(""))
                        {
                            ObjTable.Cell(2, 3).Range.Text = "01.09." + FirstYear + " - " + "06.02." + SecondYear;
                        }
                        else
                        {
                            ObjTable.Cell(2, 3).Range.Text = mdlData.colDopWork[i].DateUMR.ToString();
                        }

                        //Пишем отметку о выполнении работы, если она задана
                        if (!mdlData.colDopWork[i].CommUMR.ToString().Equals(""))
                        {
                            ObjTable.Cell(2, 4).Range.Text = mdlData.colDopWork[i].CommUMR.ToString();
                        }
                    }
                    //Если второй семестр, то УМР выполняется с 7 февраля по 30 июня
                    else
                    {
                        if (mdlData.colDopWork[i].Semestr.SemNum.Equals(mdlData.colSemestr[2].SemNum))
                        {
                            //Пишем наименование работ
                            ObjTable.Cell(3, 1).Range.Text = mdlData.colDopWork[i].UMR.ToString();

                            //Пишем объём работы, если он не задан
                            if (!mdlData.colDopWork[i].VolumeUMR.ToString().Equals(""))
                            {
                                ObjTable.Cell(3, 2).Range.Text = mdlData.colDopWork[i].VolumeUMR.ToString();
                            }

                            //Пишем стандартный срок выполнения, если он не задан
                            if (mdlData.colDopWork[i].DateUMR.ToString().Equals(""))
                            {
                                ObjTable.Cell(3, 3).Range.Text = "07.02." + SecondYear + " - " + "30.06." + SecondYear;
                            }
                            else
                            {
                                ObjTable.Cell(3, 3).Range.Text = mdlData.colDopWork[i].DateUMR.ToString();
                            }
                            
                            //Пишем отметку о выполнении работы, если она задана
                            if (!mdlData.colDopWork[i].CommUMR.ToString().Equals(""))
                            {
                                ObjTable.Cell(3, 4).Range.Text = mdlData.colDopWork[i].CommUMR.ToString();
                            }
                        }
                    }
                }
            }
        }

        private void textScienceWork(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "iii. Научно-исследовательская работа";
            //Устанавливаем все заглавные
            ObjParagraph.Range.Font.AllCaps = 1;
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Размер шрифта 14 пт
            ObjParagraph.Range.Font.Size = 14;
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            ObjParagraph.TabStops[ObjParagraph.TabStops.Count].Clear();
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textScienceWorkFromCol(object ObjMissing, Word._Document ObjDoc)
        {
            string WorkYear;
            string FirstYear;
            string SecondYear;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Table ObjTable;
            Word.Range ObjWordRange;

            //Вставляем фиксированную таблицу 3 x 4
            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjTable = ObjDoc.Tables.Add(ObjWordRange, 3, 4, ref ObjMissing, ref ObjMissing);
            //Сбрасываем все заглавные
            ObjTable.Range.Font.AllCaps = 0;
            //Убираем жирный шрифт
            ObjTable.Range.Font.Bold = 0;
            //Размер шрифта 10 пт
            ObjTable.Range.Font.Size = 10;
            //Отступ после абзаца отсутствует
            ObjTable.Range.ParagraphFormat.SpaceAfter = 0;
            //Выравнивание в ячейках по левому краю
            ObjTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //Границы таблицы включены
            ObjTable.Borders.Enable = 1;

            ObjTable.Columns[1].Width = 7.98f / 0.03527f;
            ObjTable.Columns[2].Width = 3.07f / 0.03527f;
            ObjTable.Columns[2].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            ObjTable.Columns[3].Width = 2.84f / 0.03527f;
            ObjTable.Columns[3].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            ObjTable.Columns[4].Width = 2.68f / 0.03527f;
            ObjTable.Columns[4].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            ObjTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            ObjTable.Cell(1, 1).Range.Text = "Наименование работы";
            ObjTable.Cell(1, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            ObjTable.Cell(1, 2).Range.Text = "Объём";
            ObjTable.Cell(1, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            ObjTable.Cell(1, 3).Range.Text = "Сроки";
            ObjTable.Cell(1, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            ObjTable.Cell(1, 4).Range.Text = "Отметки о выполнении";
            ObjTable.Cell(1, 4).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            ObjTable.Rows[2].Height = 12f / 0.03527f;
            ObjTable.Rows[3].Height = 12f / 0.03527f;

            WorkYear = mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear;

            FirstYear = WorkYear.Substring(0, 4);
            SecondYear = WorkYear.Substring(5, 4);

            for (int i = 0; i <= mdlData.colDopWork.Count - 1; i++)
            {
                if (mdlData.colDopWork[i].Lecturer.FIO.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO))
                {
                    //Если первый семестр, то НИР выполняется с 1 сентября по 31 декабря
                    if (mdlData.colDopWork[i].Semestr.SemNum.Equals(mdlData.colSemestr[1].SemNum))
                    {
                        //Пишем наименование работы
                        ObjTable.Cell(2, 1).Range.Text = mdlData.colDopWork[i].NIR.ToString();

                        //Если известен объём НИР, то пишем его
                        if (!mdlData.colDopWork[i].VolumeNIR.ToString().Equals(""))
                        {
                            ObjTable.Cell(2, 2).Range.Text = mdlData.colDopWork[i].VolumeNIR.ToString();
                        }

                        //Если сроки не указаны, то пишем стандартные
                        if (mdlData.colDopWork[i].DateNIR.ToString().Equals(""))
                        {
                            ObjTable.Cell(2, 3).Range.Text = "01.09." + FirstYear + " - " + "31.12." + FirstYear;
                        }
                        else
                        {
                            ObjTable.Cell(2, 3).Range.Text = mdlData.colDopWork[i].DateNIR.ToString();
                        }

                        //Если задан комментарий к НИР, то пишем его
                        if (!mdlData.colDopWork[i].CommNIR.ToString().Equals(""))
                        {
                            ObjTable.Cell(2, 4).Range.Text = mdlData.colDopWork[i].CommNIR.ToString();
                        }
                    }
                    //Если второй семестр, то НИР выполняется с 1 января по 30 июня
                    else
                    {
                        if (mdlData.colDopWork[i].Semestr.SemNum.Equals(mdlData.colSemestr[2].SemNum))
                        {
                            //Пишем наименование работы
                            ObjTable.Cell(3, 1).Range.Text = mdlData.colDopWork[i].NIR.ToString();

                            //Если задан объём НИР, то пишем его
                            if (!mdlData.colDopWork[i].VolumeNIR.ToString().Equals(""))
                            {
                                ObjTable.Cell(3, 2).Range.Text = mdlData.colDopWork[i].VolumeNIR.ToString();
                            }

                            //Если не задан срок НИР, то пишем стандартный
                            if (mdlData.colDopWork[i].DateNIR.ToString().Equals(""))
                            {
                                ObjTable.Cell(3, 3).Range.Text = "01.01." + SecondYear + " - " + "30.06." + SecondYear;
                            }
                            else
                            {
                                ObjTable.Cell(3, 3).Range.Text = mdlData.colDopWork[i].DateNIR.ToString();
                            }

                            //Если задан комментарий к НИР, то пишем его
                            if (!mdlData.colDopWork[i].CommNIR.ToString().Equals(""))
                            {
                                ObjTable.Cell(3, 4).Range.Text = mdlData.colDopWork[i].CommNIR.ToString();
                            }
                        }
                    }
                }
            }
        }

        private void textAdminWork(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "iv. Организационно-методическая и воспитательная работа";
            //Устанавливаем все заглавные
            ObjParagraph.Range.Font.AllCaps = 1;
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Размер шрифта 14 пт
            ObjParagraph.Range.Font.Size = 14;
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            ObjParagraph.TabStops[ObjParagraph.TabStops.Count].Clear();
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
        }

        private void textAdminWorkFromCol(object ObjMissing, Word._Document ObjDoc)
        {
            string WorkYear;
            string FirstYear;
            string SecondYear;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Table ObjTable;
            Word.Range ObjWordRange;

            //Вставляем фиксированную таблицу 3 x 4
            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjTable = ObjDoc.Tables.Add(ObjWordRange, 3, 4, ref ObjMissing, ref ObjMissing);
            //Сбрасываем все заглавные
            ObjTable.Range.Font.AllCaps = 0;
            //Убираем жирный шрифт
            ObjTable.Range.Font.Bold = 0;
            //Размер шрифта 10 пт
            ObjTable.Range.Font.Size = 10;
            //Отступ после абзаца отсутствует
            ObjTable.Range.ParagraphFormat.SpaceAfter = 0;
            //Выравнивание в ячейках по левому краю
            ObjTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //Границы таблицы включены
            ObjTable.Borders.Enable = 1;

            ObjTable.Columns[1].Width = 7.98f / 0.03527f;
            ObjTable.Columns[2].Width = 3.07f / 0.03527f;
            ObjTable.Columns[2].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            ObjTable.Columns[3].Width = 2.84f / 0.03527f;
            ObjTable.Columns[3].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            ObjTable.Columns[4].Width = 2.68f / 0.03527f;
            ObjTable.Columns[4].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            ObjTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            ObjTable.Cell(1, 1).Range.Text = "Наименование работы";
            ObjTable.Cell(1, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            ObjTable.Cell(1, 2).Range.Text = "Объём";
            ObjTable.Cell(1, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            ObjTable.Cell(1, 3).Range.Text = "Сроки";
            ObjTable.Cell(1, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            ObjTable.Cell(1, 4).Range.Text = "Отметки о выполнении";
            ObjTable.Cell(1, 4).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            ObjTable.Rows[2].Height = 11f / 0.03527f;
            ObjTable.Rows[3].Height = 11f / 0.03527f;

            WorkYear = mdlData.colWorkYear[cmbWorkYearList.SelectedIndex].WorkYear;

            FirstYear = WorkYear.Substring(0, 4);
            SecondYear = WorkYear.Substring(5, 4);

            for (int i = 0; i <= mdlData.colDopWork.Count - 1; i++)
            {
                if (mdlData.colDopWork[i].Lecturer.FIO.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO))
                {
                    if (mdlData.colDopWork[i].Semestr.SemNum.Equals(mdlData.colSemestr[1].SemNum))
                    {
                        //Пишем наименование работы
                        ObjTable.Cell(2, 1).Range.Text = mdlData.colDopWork[i].OMR.ToString();
                        
                        //Если указан объём ОМР, то пишем его
                        if (!mdlData.colDopWork[i].VolumeOMR.ToString().Equals(""))
                        {
                            ObjTable.Cell(2, 2).Range.Text = mdlData.colDopWork[i].VolumeOMR.ToString();
                        }

                        //Если срок ОМР не указан, то пишем стандартный
                        if (mdlData.colDopWork[i].DateOMR.ToString().Equals(""))
                        {
                            ObjTable.Cell(2, 3).Range.Text = "01.09." + FirstYear + " - " + "06.02." + SecondYear;
                        }
                        else
                        {
                            ObjTable.Cell(2, 3).Range.Text = mdlData.colDopWork[i].DateOMR.ToString();
                        }

                        //Если задан комментарий к ОМР, то пишем его
                        if (!mdlData.colDopWork[i].CommOMR.ToString().Equals(""))
                        {
                            ObjTable.Cell(2, 4).Range.Text = mdlData.colDopWork[i].CommOMR.ToString();
                        }
                    }
                    else
                    {
                        if (mdlData.colDopWork[i].Semestr.SemNum.Equals(mdlData.colSemestr[2].SemNum))
                        {
                            //Пишем наименование работы
                            ObjTable.Cell(3, 1).Range.Text = mdlData.colDopWork[i].OMR.ToString();

                            //Если указан объём ОМР, то пишем его
                            if (!mdlData.colDopWork[i].VolumeOMR.ToString().Equals(""))
                            {
                                ObjTable.Cell(3, 2).Range.Text = mdlData.colDopWork[i].VolumeOMR.ToString();
                            }

                            //Если срок ОМР не указан, то пишем стандартный
                            if (mdlData.colDopWork[i].DateOMR.ToString().Equals(""))
                            {
                                ObjTable.Cell(3, 3).Range.Text = "07.02." + SecondYear + " - " + "30.06." + SecondYear;
                            }
                            else
                            {
                                ObjTable.Cell(3, 3).Range.Text = mdlData.colDopWork[i].DateOMR.ToString();
                            }

                            //Если задан комментарий к ОМР, то пишем его
                            if (!mdlData.colDopWork[i].CommOMR.ToString().Equals(""))
                            {
                                ObjTable.Cell(3, 4).Range.Text = mdlData.colDopWork[i].CommOMR.ToString();
                            }
                        }
                    }
                }
            }
        }

        private void textSignature(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            //
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            //Добавляем ещё один абзац текста спереди
            ObjParagraph.Range.InsertParagraphBefore();
            //Добавляем ещё один абзац текста спереди
            ObjParagraph.Range.InsertParagraphBefore();
            //
            ObjParagraph.Range.Text = "Преподаватель ________________________________________ " + 
                mdlData.SplitFIOString(mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO, false, true);
            //Сбрасываем все заглавные
            ObjParagraph.Range.Font.AllCaps = 0;
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Размер шрифта 14 пт
            ObjParagraph.Range.Font.Size = 12;
            //Выравнивание по левому краю
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
        }

        private void textSignatureMain(object ObjMissing, Word._Document ObjDoc)
        {
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            //
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            //Добавляем ещё один абзац текста спереди
            ObjParagraph.Range.InsertParagraphBefore();
            //Добавляем ещё один абзац текста спереди
            ObjParagraph.Range.InsertParagraphBefore();
            //
            ObjParagraph.Range.Text = "Заведующий кафедрой ________________________________________ " +
                mdlData.SplitFIOString(mdlData.colLecturer[cmbLecturerList.SelectedIndex].FIO, false, true);
            //Сбрасываем все заглавные
            ObjParagraph.Range.Font.AllCaps = 0;
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Размер шрифта 14 пт
            ObjParagraph.Range.Font.Size = 12;
            //Выравнивание по левому краю
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
        }

        private static string[] findSameNamesDiploma(IList<clsDistribution> coll)
        {
            string NameDiploma;
            bool flgNamesDiploma;
            string[] NamesDiploma;

            //Очищаем переменную, содержащую возможные названия дисциплин,
            //связанных с дипломным проектированием
            NameDiploma = "";
            //Перебираем все элементы учебной нагрузки
            for (int i = 0; i <= coll.Count - 1; i++)
            {
                //Если у текущего элемента предусмотрены часы на
                //дипломное проектирование
                if (!(coll[i].DiplomaPaper == 0))
                {
                    //В массив названий предметов записываем компоненты из
                    //строки предметов, связанных с дипломным проектированием.
                    //Разделителем является ";"
                    NamesDiploma = NameDiploma.Split(new char[] { ';' });
                    //Если массив названий предметов пуст
                    if (NamesDiploma.GetLength(0) == 0)
                    {
                        //просто заносим в него название текущего предмета
                        NameDiploma += coll[i].Subject.Subject + ";";
                    }
                    //в ином случае необходима более детальная проверка
                    else
                    {
                        //выставляем признак необходимости добавления нового названия
                        //по умолчанию
                        flgNamesDiploma = true;
                        //Перебираем все элементы массива называний предметов,
                        //связанных с дипломным проектированием
                        for (int j = 0; j <= NamesDiploma.GetLength(0) - 1; j++)
                        {
                            //Если в массиве уже есть такое имя предмета, как
                            //у текущего элемента
                            if (coll[i].Subject.Subject == NamesDiploma[j])
                            {
                                //Сбрасываем признак необходимости добавления нового
                                //названия
                                flgNamesDiploma = false;
                            }
                        }
                        //Если признак необходимости добавления нового названия
                        //выставлен
                        if (flgNamesDiploma)
                        {
                            //Записываем это название в строку
                            NameDiploma += coll[i].Subject.Subject + ";";
                        }
                    }
                }
            }

            //Удаляем из строки названий замыкающий разделитель
            NameDiploma = NameDiploma.Substring(0, NameDiploma.Length - 1);
            //Укорачиваем массив названий предметов
            NamesDiploma = NameDiploma.Split(new char[] { ';' });

            return NamesDiploma;
        }

        private static string[] findSameNamesProdPr(IList<clsDistribution> coll)
        {
            string NameProdPr;
            bool flgNamesProdPr;
            string[] NamesProdPr;

            //Очищаем переменную, содержащую возможные названия дисциплин,
            //связанных с дипломным проектированием
            NameProdPr = "";
            //Перебираем все элементы учебной нагрузки
            for (int i = 0; i <= coll.Count - 1; i++)
            {
                //Если у текущего элемента предусмотрены часы на
                //дипломное проектирование
                if (!(coll[i].ProducingPractice == 0))
                {
                    //В массив названий предметов записываем компоненты из
                    //строки предметов, связанных с дипломным проектированием.
                    //Разделителем является ";"
                    NamesProdPr = NameProdPr.Split(new char[] { ';' });
                    //Если массив названий предметов пуст
                    if (NamesProdPr.GetLength(0) == 0)
                    {
                        //просто заносим в него название текущего предмета
                        NameProdPr += coll[i].Subject.Subject + ";";
                    }
                    //в ином случае необходима более детальная проверка
                    else
                    {
                        //выставляем признак необходимости добавления нового названия
                        //по умолчанию
                        flgNamesProdPr = true;
                        //Перебираем все элементы массива называний предметов,
                        //связанных с дипломным проектированием
                        for (int j = 0; j <= NamesProdPr.GetLength(0) - 1; j++)
                        {
                            //Если в массиве уже есть такое имя предмета, как
                            //у текущего элемента
                            if (coll[i].Subject.Subject == NamesProdPr[j])
                            {
                                //Сбрасываем признак необходимости добавления нового
                                //названия
                                flgNamesProdPr = false;
                            }
                        }
                        //Если признак необходимости добавления нового названия
                        //выставлен
                        if (flgNamesProdPr)
                        {
                            //Записываем это название в строку
                            NameProdPr += coll[i].Subject.Subject + ";";
                        }
                    }
                }
            }

            //Удаляем из строки названий замыкающий разделитель
            NameProdPr = NameProdPr.Substring(0, NameProdPr.Length - 1);
            //Укорачиваем массив названий предметов
            NamesProdPr = NameProdPr.Split(new char[] { ';' });

            return NamesProdPr;
        }

        private static string[] findSameNamesTutorialPr(IList<clsDistribution> coll)
        {
            string NameTutorialPr;
            bool flgNamesTutorialPr;
            string[] NamesTutorialPr;

            //Очищаем переменную, содержащую возможные названия дисциплин,
            //связанных с дипломным проектированием
            NameTutorialPr = "";
            //Перебираем все элементы учебной нагрузки
            for (int i = 0; i <= coll.Count - 1; i++)
            {
                //Если у текущего элемента предусмотрены часы на
                //дипломное проектирование
                if (!(coll[i].TutorialPractice == 0))
                {
                    //В массив названий предметов записываем компоненты из
                    //строки предметов, связанных с дипломным проектированием.
                    //Разделителем является ";"
                    NamesTutorialPr = NameTutorialPr.Split(new char[] { ';' });
                    //Если массив названий предметов пуст
                    if (NamesTutorialPr.GetLength(0) == 0)
                    {
                        //просто заносим в него название текущего предмета
                        NameTutorialPr += coll[i].Subject.Subject + ";";
                    }
                    //в ином случае необходима более детальная проверка
                    else
                    {
                        //выставляем признак необходимости добавления нового названия
                        //по умолчанию
                        flgNamesTutorialPr = true;
                        //Перебираем все элементы массива называний предметов,
                        //связанных с дипломным проектированием
                        for (int j = 0; j <= NamesTutorialPr.GetLength(0) - 1; j++)
                        {
                            //Если в массиве уже есть такое имя предмета, как
                            //у текущего элемента
                            if (coll[i].Subject.Subject == NamesTutorialPr[j])
                            {
                                //Сбрасываем признак необходимости добавления нового
                                //названия
                                flgNamesTutorialPr = false;
                            }
                        }
                        //Если признак необходимости добавления нового названия
                        //выставлен
                        if (flgNamesTutorialPr)
                        {
                            //Записываем это название в строку
                            NameTutorialPr += coll[i].Subject.Subject + ";";
                        }
                    }
                }
            }

            //Удаляем из строки названий замыкающий разделитель
            NameTutorialPr = NameTutorialPr.Substring(0, NameTutorialPr.Length - 1);
            //Укорачиваем массив названий предметов
            NamesTutorialPr = NameTutorialPr.Split(new char[] { ';' });

            return NamesTutorialPr;
        }

        private static string[] findSameNamesMagistry(IList<clsDistribution> coll)
        {
            string NameMagistry;
            bool flgNamesMagistry;
            string[] NamesMagistry;

            //Очищаем переменную, содержащую возможные названия дисциплин,
            //связанных с дипломным проектированием
            NameMagistry = "";
            //Перебираем все элементы учебной нагрузки
            for (int i = 0; i <= coll.Count - 1; i++)
            {
                //Если у текущего элемента предусмотрены часы на
                //дипломное проектирование
                if (!(coll[i].Magistry == 0))
                {
                    //В массив названий предметов записываем компоненты из
                    //строки предметов, связанных с дипломным проектированием.
                    //Разделителем является ";"
                    NamesMagistry = NameMagistry.Split(new char[] { ';' });
                    //Если массив названий предметов пуст
                    if (NamesMagistry.GetLength(0) == 0)
                    {
                        //просто заносим в него название текущего предмета
                        NameMagistry += coll[i].Subject.Subject + ";";
                    }
                    //в ином случае необходима более детальная проверка
                    else
                    {
                        //выставляем признак необходимости добавления нового названия
                        //по умолчанию
                        flgNamesMagistry = true;
                        //Перебираем все элементы массива называний предметов,
                        //связанных с дипломным проектированием
                        for (int j = 0; j <= NamesMagistry.GetLength(0) - 1; j++)
                        {
                            //Если в массиве уже есть такое имя предмета, как
                            //у текущего элемента
                            if (coll[i].Subject.Subject == NamesMagistry[j])
                            {
                                //Сбрасываем признак необходимости добавления нового
                                //названия
                                flgNamesMagistry = false;
                            }
                        }
                        //Если признак необходимости добавления нового названия
                        //выставлен
                        if (flgNamesMagistry)
                        {
                            //Записываем это название в строку
                            NameMagistry += coll[i].Subject.Subject + ";";
                        }
                    }
                }
            }

            //Удаляем из строки названий замыкающий разделитель
            NameMagistry = NameMagistry.Substring(0, NameMagistry.Length - 1);
            //Укорачиваем массив названий предметов
            NamesMagistry = NameMagistry.Split(new char[] { ';' });

            return NamesMagistry;
        }

        private void findSameGroupsDiploma(IList<clsDistribution> coll, ref string[] GroupsDiploma, ref string[] GroupsFullDiploma)
        {
            bool flgGroupsDiploma;
            string GroupDiploma;
            string GroupFullDiploma;

            //Очищаем переменную, содержащую возможные группы
            //в которых проводится дипломное проектирование
            GroupDiploma = "";
            GroupFullDiploma = "";
            //Перебираем все элементы учебной нагрузки
            for (int i = 0; i <= coll.Count - 1; i++)
            {
                //Если у текущего элемента предусмотрены часы на
                //дипломное проектирование
                if (!(coll[i].DiplomaPaper == 0))
                {
                    //В массив групп записываем компоненты из
                    //строки групп, в которых проводится дипломное 
                    //проектирование. Разделителем является ";"
                    GroupsDiploma = GroupDiploma.Split(new char[] { ';' });
                    //Если массив групп пуст
                    if (GroupsDiploma.GetLength(0) == 0)
                    {
                        //просто заносим в него название группы
                        GroupDiploma += coll[i].Speciality.ShortUpravlenie + ";";

                        GroupFullDiploma += coll[i].Speciality.ShortUpravlenie
                                            + "-" + coll[i].KursNum.Kurs + "\n ("
                                            + coll[i].Speciality.ShortInstitute + "-" +
                                            coll[i].KursNum.Kurs + "1  )" + ";";
                    }
                    //в ином случае необходима более детальная проверка
                    else
                    {
                        //выставляем признак необходимости добавления новой группы
                        //по умолчанию
                        flgGroupsDiploma = true;
                        //Перебираем все элементы массива групп,
                        //связанных с дипломным проектированием
                        for (int j = 0; j <= GroupsDiploma.GetLength(0) - 1; j++)
                        {
                            //Если в массиве уже есть такая группа, как
                            //у текущего элемента
                            if (coll[i].Speciality.ShortUpravlenie == GroupsDiploma[j])
                            {
                                //Сбрасываем признак необходимости добавления новой
                                //группы
                                flgGroupsDiploma = false;
                            }
                        }
                        //Если признак необходимости добавления новой группы
                        //выставлен
                        if (flgGroupsDiploma)
                        {
                            //Записываем эту группу в строку
                            GroupDiploma += coll[i].Speciality.ShortUpravlenie + ";";

                            GroupFullDiploma += coll[i].Speciality.ShortUpravlenie
                                                + "-" + coll[i].KursNum.Kurs + "\n ("
                                                + coll[i].Speciality.ShortInstitute + "-" +
                                                coll[i].KursNum.Kurs + "1  )" + ";";
                        }
                    }
                }
            }

            //Удаляем из строки названий замыкающий разделитель
            GroupDiploma = GroupDiploma.Substring(0, GroupDiploma.Length - 1);
            //Укорачиваем массив групп
            GroupsDiploma = GroupDiploma.Split(new char[] { ';' });

            //Удаляем из строки названий замыкающий разделитель
            GroupFullDiploma = GroupFullDiploma.Substring(0, GroupFullDiploma.Length - 1);
            //Укорачиваем массив групп
            GroupsFullDiploma = GroupFullDiploma.Split(new char[] { ';' });
        }

        private void findSameGroupsMagistry(IList<clsDistribution> coll, ref string[] GroupsMagistry, ref string[] GroupsFullMagistry)
        {
            bool flgGroupsMagistry;
            string GroupMagistry;
            string GroupFullMagistry;

            //Очищаем переменную, содержащую возможные группы
            //в которых проводится дипломное проектирование
            GroupMagistry = "";
            GroupFullMagistry = "";
            //Перебираем все элементы учебной нагрузки
            for (int i = 0; i <= coll.Count - 1; i++)
            {
                //Если у текущего элемента предусмотрены часы на
                //дипломное проектирование
                if (!(coll[i].Magistry == 0))
                {
                    //В массив групп записываем компоненты из
                    //строки групп, в которых проводится дипломное 
                    //проектирование. Разделителем является ";"
                    GroupsMagistry = GroupMagistry.Split(new char[] { ';' });
                    //Если массив групп пуст
                    if (GroupsMagistry.GetLength(0) == 0)
                    {
                        //просто заносим в него название группы
                        GroupMagistry += coll[i].Speciality.ShortUpravlenie + "-" + coll[i].KursNum.Kurs + ";";

                        GroupFullMagistry += coll[i].Speciality.ShortUpravlenie
                                                + "-" + coll[i].KursNum.Kurs + "\n ("
                                                + coll[i].Speciality.ShortInstitute + "-" +
                                                coll[i].KursNum.Kurs + "1  )" + ";";
                    }
                    //в ином случае необходима более детальная проверка
                    else
                    {
                        //выставляем признак необходимости добавления новой группы
                        //по умолчанию
                        flgGroupsMagistry = true;
                        //Перебираем все элементы массива групп,
                        //связанных с дипломным проектированием
                        for (int j = 0; j <= GroupsMagistry.GetLength(0) - 1; j++)
                        {
                            //Если в массиве уже есть такая группа, как
                            //у текущего элемента
                            if (!(coll[i].Speciality == null))
                            {
                                if ((coll[i].Speciality.ShortUpravlenie + "-" + coll[i].KursNum.Kurs) == GroupsMagistry[j])
                                {
                                    //Сбрасываем признак необходимости добавления новой
                                    //группы
                                    flgGroupsMagistry = false;
                                }
                            }
                            else
                            {
                                //Сбрасываем признак необходимости добавления новой
                                //группы
                                flgGroupsMagistry = false;
                            }
                        }
                        //Если признак необходимости добавления новой группы
                        //выставлен
                        if (flgGroupsMagistry)
                        {
                            //Записываем эту группу в строку
                            GroupMagistry += coll[i].Speciality.ShortUpravlenie + "-" + coll[i].KursNum.Kurs + ";";

                            GroupFullMagistry += coll[i].Speciality.ShortUpravlenie
                                                + "-" + coll[i].KursNum.Kurs + "\n ("
                                                + coll[i].Speciality.ShortInstitute + "-" +
                                                coll[i].KursNum.Kurs + "1  )" + ";";
                        }
                    }
                }
            }

            //Удаляем из строки названий замыкающий разделитель
            GroupMagistry = GroupMagistry.Substring(0, GroupMagistry.Length - 1);
            //Укорачиваем массив групп
            GroupsMagistry = GroupMagistry.Split(new char[] { ';' });

            //Удаляем из строки названий замыкающий разделитель
            GroupFullMagistry = GroupFullMagistry.Substring(0, GroupFullMagistry.Length - 1);
            //Укорачиваем массив групп
            GroupsFullMagistry = GroupFullMagistry.Split(new char[] { ';' });
        }

        private void findSameGroupsProdPr(IList<clsDistribution> coll, ref string[] GroupsProdPr, ref string[] GroupsFullProdPr)
        {
            bool flgGroupsProdPr;
            string GroupProdPr;
            string GroupFullProdPr;

            //Очищаем переменную, содержащую возможные группы
            //в которых проводится дипломное проектирование
            GroupProdPr = "";
            GroupFullProdPr = "";
            //Перебираем все элементы учебной нагрузки
            for (int i = 0; i <= coll.Count - 1; i++)
            {
                //Если у текущего элемента предусмотрены часы на
                //дипломное проектирование
                if (!(coll[i].ProducingPractice == 0))
                {
                    //В массив групп записываем компоненты из
                    //строки групп, в которых проводится дипломное 
                    //проектирование. Разделителем является ";"
                    GroupsProdPr = GroupProdPr.Split(new char[] { ';' });
                    //Если массив групп пуст
                    if (GroupsProdPr.GetLength(0) == 0)
                    {
                        //просто заносим в него название группы
                        GroupProdPr += coll[i].Speciality.ShortUpravlenie + "-" + coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + ";";

                        GroupFullProdPr += coll[i].Speciality.ShortUpravlenie
                                                + "-" + coll[i].KursNum.Kurs + "\n ("
                                                + coll[i].Speciality.ShortInstitute + "-" +
                                                coll[i].KursNum.Kurs + "1  )" + ";";
                    }
                    //в ином случае необходима более детальная проверка
                    else
                    {
                        //выставляем признак необходимости добавления новой группы
                        //по умолчанию
                        flgGroupsProdPr = true;
                        //Перебираем все элементы массива групп,
                        //связанных с дипломным проектированием
                        for (int j = 0; j <= GroupsProdPr.GetLength(0) - 1; j++)
                        {
                            //Если в массиве уже есть такая группа, как
                            //у текущего элемента
                            if (!(coll[i].Speciality == null))
                            {
                                if ((coll[i].Speciality.ShortUpravlenie + "-" + coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs) == GroupsProdPr[j])
                                {
                                    //Сбрасываем признак необходимости добавления новой
                                    //группы
                                    flgGroupsProdPr = false;
                                }
                            }
                            else
                            {
                                //Сбрасываем признак необходимости добавления новой
                                //группы
                                flgGroupsProdPr = false;
                            }
                        }
                        //Если признак необходимости добавления новой группы
                        //выставлен
                        if (flgGroupsProdPr)
                        {
                            //Записываем эту группу в строку
                            GroupProdPr += coll[i].Speciality.ShortUpravlenie + "-" + coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + ";";

                            GroupFullProdPr += coll[i].Speciality.ShortUpravlenie
                                                + "-" + coll[i].KursNum.Kurs + "\n ("
                                                + coll[i].Speciality.ShortInstitute + "-" +
                                                coll[i].KursNum.Kurs + "1  )" + ";";
                        }
                    }
                }
            }

            //Удаляем из строки названий замыкающий разделитель
            GroupProdPr = GroupProdPr.Substring(0, GroupProdPr.Length - 1);
            //Укорачиваем массив групп
            GroupsProdPr = GroupProdPr.Split(new char[] { ';' });

            //Удаляем из строки названий замыкающий разделитель
            GroupFullProdPr = GroupFullProdPr.Substring(0, GroupFullProdPr.Length - 1);
            //Укорачиваем массив групп
            GroupsFullProdPr = GroupFullProdPr.Split(new char[] { ';' });
        }

        private void findSameGroupsTutorialPr(IList<clsDistribution> coll, ref string[] GroupsTutorialPr, ref string[] GroupsFullTutorialPr)
        {
            bool flgGroupsTutorialPr;
            string GroupTutorialPr;
            string GroupFullTutorialPr;

            //Очищаем переменную, содержащую возможные группы
            //в которых проводится дипломное проектирование
            GroupTutorialPr = "";
            GroupFullTutorialPr = "";
            //Перебираем все элементы учебной нагрузки
            for (int i = 0; i <= coll.Count - 1; i++)
            {
                //Если у текущего элемента предусмотрены часы на
                //дипломное проектирование
                if (!(coll[i].TutorialPractice == 0))
                {
                    //В массив групп записываем компоненты из
                    //строки групп, в которых проводится дипломное 
                    //проектирование. Разделителем является ";"
                    GroupsTutorialPr = GroupTutorialPr.Split(new char[] { ';' });
                    //Если массив групп пуст
                    if (GroupsTutorialPr.GetLength(0) == 0)
                    {
                        //просто заносим в него название группы
                        GroupTutorialPr += coll[i].Speciality.ShortUpravlenie + "-" + coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + ";";

                        GroupFullTutorialPr += coll[i].Speciality.ShortUpravlenie
                                                + "-" + coll[i].KursNum.Kurs + "\n ("
                                                + coll[i].Speciality.ShortInstitute + "-" +
                                                coll[i].KursNum.Kurs + "1  )" + ";";
                    }
                    //в ином случае необходима более детальная проверка
                    else
                    {
                        //выставляем признак необходимости добавления новой группы
                        //по умолчанию
                        flgGroupsTutorialPr = true;
                        //Перебираем все элементы массива групп,
                        //связанных с дипломным проектированием
                        for (int j = 0; j <= GroupsTutorialPr.GetLength(0) - 1; j++)
                        {
                            //Если в массиве уже есть такая группа, как
                            //у текущего элемента
                            if (!(coll[i].Speciality == null))
                            {
                                if ((coll[i].Speciality.ShortUpravlenie + "-" + coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs) == GroupsTutorialPr[j])
                                {
                                    //Сбрасываем признак необходимости добавления новой
                                    //группы
                                    flgGroupsTutorialPr = false;
                                }
                            }
                            else
                            {
                                //Сбрасываем признак необходимости добавления новой
                                //группы
                                flgGroupsTutorialPr = false;
                            }
                        }
                        //Если признак необходимости добавления новой группы
                        //выставлен
                        if (flgGroupsTutorialPr)
                        {
                            //Записываем эту группу в строку
                            GroupTutorialPr += coll[i].Speciality.ShortUpravlenie + "-" + coll[i].Speciality.ShortInstitute + "-" + coll[i].KursNum.Kurs + ";";

                            GroupFullTutorialPr += coll[i].Speciality.ShortUpravlenie
                                                + "-" + coll[i].KursNum.Kurs + "\n ("
                                                + coll[i].Speciality.ShortInstitute + "-" +
                                                coll[i].KursNum.Kurs + "1  )" + ";";
                        }
                    }
                }
            }

            //Удаляем из строки названий замыкающий разделитель
            GroupTutorialPr = GroupTutorialPr.Substring(0, GroupTutorialPr.Length - 1);
            //Укорачиваем массив групп
            GroupsTutorialPr = GroupTutorialPr.Split(new char[] { ';' });

            //Удаляем из строки названий замыкающий разделитель
            GroupFullTutorialPr = GroupFullTutorialPr.Substring(0, GroupFullTutorialPr.Length - 1);
            //Укорачиваем массив групп
            GroupsFullTutorialPr = GroupFullTutorialPr.Split(new char[] { ';' });
        }

        private static int countTableRows(IList<clsDistribution> coll, int LectId)
        {
            int countRow, i, j;

            //Создаём строки
            //1. под шапку первого семестра
            //2. под надпись первого семестра
            //3. всего часов за семестр
            //4. пробел
            //5. под шапку второго семестра
            //6. под надпись второго семестра
            //7. всего часов за семестр
            //8. пробел
            //9. всего часов за учебный год

            countRow = 9;

            //Формируем строки по нагрузке преподавателя в первом семестре
            //прогоняем все нагрузочные строки
            for (i = 0; i <= coll.Count - 1; i++)
            {
                //-------------------------------------------------------------
                //1 семестр
                //-------------------------------------------------------------
                if (coll[i].Semestr.SemNum.Equals("1 семестр"))
                {
                    if (!coll[i].flgExclude)
                    {
                        //Если преподаватель определён
                        if (!(coll[i].Lecturer == null))
                        {
                            if ((coll[i].Lecturer.Equals(mdlData.colLecturer[LectId])))
                            {
                                //Если есть лекционные часы - добавляем строку
                                if (!(coll[i].Lecture == 0))
                                {
                                    countRow++;
                                }

                                //Если есть экзаменационные часы - добавляем строку
                                if (!(coll[i].Exam == 0))
                                {
                                    countRow++;
                                }

                                //Если есть зачётные часы - добавляем строку
                                if (!(coll[i].Credit == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на реферат - добавляем строку
                                if (!(coll[i].RefHomeWork == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на консультацию - добавляем строку
                                if (!(coll[i].Tutorial == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на лабораторные работы - добавляем строку
                                if (!(coll[i].LabWork == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на практические занятия - добавляем строку
                                if (!(coll[i].Practice == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на индивидуальные занятия - добавляем строку
                                if (!(coll[i].IndividualWork == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на КРАПК - добавляем строку
                                if (!(coll[i].KRAPK == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на курсовой проект - добавляем строку
                                if (!(coll[i].KursProject == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на 
                                if (!(coll[i].TutorialPractice == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на 
                                if (!(coll[i].ProducingPractice == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на 
                                if (!(coll[i].Magistry == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на преддипломную практику - добавляем строку
                                if (!(coll[i].PreDiplomaPractice == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на ГАК - добавляем строку
                                if (!(coll[i].GAK == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на Посещение занятий - добавляем строку
                                if (!(coll[i].Visiting == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на Аспирантуру - добавляем строку
                                if (!(coll[i].PostGrad == 0))
                                {
                                    countRow++;
                                }
                            }
                        }
                        //Если преподаватель не определён, то нагрузка может быть равномерно
                        //распределяемой по преподавателям в зависимости от количества студентов
                        else
                        {
                            if (coll[i].flgDistrib)
                            {
                                for (j = 0; j <= mdlData.colStudents.Count - 1; j++)
                                {
                                    if (mdlData.colStudents[j].flgPlan)
                                    {
                                        //Если рассматриваемый преподаватель - руководитель студента
                                        //И если студент на том же курсе, где и дисциплина
                                        //И специальность студента должна соответствовать специальности нагрузки
                                        if (mdlData.colStudents[j].Lect.Equals(mdlData.colLecturer[LectId])
                                            & mdlData.colStudents[j].KursNum.Equals(coll[i].KursNum)
                                            & mdlData.colStudents[j].Speciality.Equals(coll[i].Speciality))
                                        {
                                            //Если преподаватель что-либо из этого ведёт
                                            //добавляем строку
                                            countRow++;
                                            //Одной строки достаточно, прерываем цикл
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                //-------------------------------------------------------------
                //1 семестр
                //-------------------------------------------------------------

                //-------------------------------------------------------------
                //2 семестр
                //-------------------------------------------------------------
                if (coll[i].Semestr.SemNum.Equals("2 семестр"))
                {
                    if (!coll[i].flgExclude)
                    {
                        if (!(coll[i].Lecturer == null))
                        {
                            if ((coll[i].Lecturer.Equals(mdlData.colLecturer[LectId])))
                            {
                                //Если есть лекционные часы - добавляем строку
                                if (!(coll[i].Lecture == 0))
                                {
                                    countRow++;
                                }

                                //Если есть экзаменационные часы - добавляем строку
                                if (!(coll[i].Exam == 0))
                                {
                                    countRow++;
                                }

                                //Если есть зачётные часы - добавляем строку
                                if (!(coll[i].Credit == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на реферат - добавляем строку
                                if (!(coll[i].RefHomeWork == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на консультацию - добавляем строку
                                if (!(coll[i].Tutorial == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на лабораторные работы - добавляем строку
                                if (!(coll[i].LabWork == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на практические занятия - добавляем строку
                                if (!(coll[i].Practice == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на индивидуальные занятия - добавляем строку
                                if (!(coll[i].IndividualWork == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на КРАПК - добавляем строку
                                if (!(coll[i].KRAPK == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на курсовой проект - добавляем строку
                                if (!(coll[i].KursProject == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на 
                                if (!(coll[i].TutorialPractice == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на 
                                if (!(coll[i].ProducingPractice == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на 
                                if (!(coll[i].Magistry == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на преддипломную практику - добавляем строку
                                if (!(coll[i].PreDiplomaPractice == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на ГАК - добавляем строку
                                if (!(coll[i].GAK == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на Посещение занятий - добавляем строку
                                if (!(coll[i].Visiting == 0))
                                {
                                    countRow++;
                                }

                                //Если есть часы на Аспирантуру - добавляем строку
                                if (!(coll[i].PostGrad == 0))
                                {
                                    countRow++;
                                }
                            }
                        }
                        else
                        {
                            if (coll[i].flgDistrib)
                            {
                                for (j = 0; j <= mdlData.colStudents.Count - 1; j++)
                                {
                                    if (mdlData.colStudents[j].flgPlan)
                                    {
                                        //Если рассматриваемый преподаватель - руководитель студента
                                        //И если студент на том же курсе, где и дисциплина
                                        //И специальность студента должна соответствовать специальности нагрузки
                                        if (mdlData.colStudents[j].Lect.Equals(mdlData.colLecturer[LectId])
                                            & mdlData.colStudents[j].KursNum.Equals(coll[i].KursNum)
                                            & mdlData.colStudents[j].Speciality.Equals(coll[i].Speciality))
                                        {
                                            //Если преподаватель что-либо из этого ведёт
                                            //добавляем строку
                                            countRow++;
                                            //Одной строки достаточно, прерываем цикл
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                //-------------------------------------------------------------
                //2 семестр
                //-------------------------------------------------------------
            }
            
            return countRow;
        }

        //
        private static void DiplomaAddRow(IList<clsDistribution> coll, int LectId, ref int countRow)
        {
            bool flgAddRow;

            for (int i = 0; i <= NamesDiploma.GetLength(0) - 1; i++)
            {
                for (int j = 0; j <= GroupsDiploma.GetLength(0) - 1; j++)
                {
                    flgAddRow = false;
                    for (int k = 0; k <= coll.Count - 1; k++)
                    {
                        if (coll[k].Semestr.SemNum.Equals("1 семестр"))
                        {
                            if (!(coll[k].Lecturer == null))
                            {
                                if ((coll[k].Lecturer.Equals(mdlData.colLecturer[LectId])))
                                {
                                    if (!(coll[k].DiplomaPaper == 0))
                                    {
                                        if ((coll[k].Subject.Subject == NamesDiploma[i]) &
                                            (coll[k].Speciality.ShortUpravlenie == GroupsDiploma[j]))
                                        {
                                            flgAddRow = true;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (flgAddRow)
                    {
                        countRow++;
                    }

                    flgAddRow = false;
                    for (int k = 0; k <= coll.Count - 1; k++)
                    {
                        if (coll[k].Semestr.SemNum.Equals("2 семестр"))
                        {
                            if (!(coll[k].Lecturer == null))
                            {
                                if ((coll[k].Lecturer.Equals(mdlData.colLecturer[LectId])))
                                {
                                    if (!(coll[k].DiplomaPaper == 0))
                                    {
                                        if ((coll[k].Subject.Subject == NamesDiploma[i]) &
                                            (coll[k].Speciality.ShortUpravlenie == GroupsDiploma[j]))
                                        {
                                            flgAddRow = true;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (flgAddRow)
                    {
                        countRow++;
                    }
                }
            }
        }

        //
        private static int countTableRowsDiploma(IList<clsDistribution> coll, int LectId)
        {
            int countRow;

            //Создаём строки
            //1. под шапку первого семестра
            //2. под надпись первого семестра
            //3. всего часов за семестр
            //4. пробел
            //5. под шапку второго семестра
            //6. под надпись второго семестра
            //7. всего часов за семестр
            //8. пробел
            //9. всего часов за учебный год

            countRow = 9;

            DiplomaAddRow(coll, LectId, ref countRow);

            return countRow;
        }

        private static void ProdPrAddRow(IList<clsDistribution> coll, int LectId, ref int countRow)
        {
            bool flgAddRow;

            for (int i = 0; i <= NamesProdPr.GetLength(0) - 1; i++)
            {
                for (int j = 0; j <= GroupsProdPr.GetLength(0) - 1; j++)
                {
                    flgAddRow = false;
                    for (int k = 0; k <= coll.Count - 1; k++)
                    {
                        if (coll[k].Semestr.SemNum.Equals("1 семестр"))
                        {
                            if (!(coll[k].Lecturer == null))
                            {
                                if ((coll[k].Lecturer.Equals(mdlData.colLecturer[LectId])))
                                {
                                    if (!(coll[k].ProducingPractice == 0))
                                    {
                                        if ((coll[k].Subject.Subject == NamesProdPr[i]) &
                                            ((coll[k].Speciality.ShortUpravlenie + "-" + coll[k].Speciality.ShortInstitute + "-" + coll[k].KursNum.Kurs) == GroupsProdPr[j]))
                                        {
                                            flgAddRow = true;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (flgAddRow)
                    {
                        countRow++;
                    }

                    flgAddRow = false;
                    for (int k = 0; k <= coll.Count - 1; k++)
                    {
                        if (coll[k].Semestr.SemNum.Equals("2 семестр"))
                        {
                            if (!(coll[k].Lecturer == null))
                            {
                                if ((coll[k].Lecturer.Equals(mdlData.colLecturer[LectId])))
                                {
                                    if (!(coll[k].ProducingPractice == 0))
                                    {
                                        if ((coll[k].Subject.Subject == NamesProdPr[i]) &
                                            ((coll[k].Speciality.ShortUpravlenie + "-" + coll[k].Speciality.ShortInstitute + "-" + coll[k].KursNum.Kurs) == GroupsProdPr[j]))
                                        {
                                            flgAddRow = true;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (flgAddRow)
                    {
                        countRow++;
                    }
                }
            }
        }

        private static void TutorialPrAddRow(IList<clsDistribution> coll, int LectId, ref int countRow)
        {
            bool flgAddRow;

            for (int i = 0; i <= NamesTutorialPr.GetLength(0) - 1; i++)
            {
                for (int j = 0; j <= GroupsTutorialPr.GetLength(0) - 1; j++)
                {
                    flgAddRow = false;
                    for (int k = 0; k <= coll.Count - 1; k++)
                    {
                        if (coll[k].Semestr.SemNum.Equals("1 семестр"))
                        {
                            if (!(coll[k].Lecturer == null))
                            {
                                if ((coll[k].Lecturer.Equals(mdlData.colLecturer[LectId])))
                                {
                                    if (!(coll[k].TutorialPractice == 0))
                                    {
                                        if ((coll[k].Subject.Subject == NamesTutorialPr[i]) &
                                            ((coll[k].Speciality.ShortUpravlenie + "-" + coll[k].Speciality.ShortInstitute + "-" + coll[k].KursNum.Kurs) == GroupsTutorialPr[j]))
                                        {
                                            flgAddRow = true;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (flgAddRow)
                    {
                        countRow++;
                    }

                    flgAddRow = false;
                    for (int k = 0; k <= coll.Count - 1; k++)
                    {
                        if (coll[k].Semestr.SemNum.Equals("2 семестр"))
                        {
                            if (!(coll[k].Lecturer == null))
                            {
                                if ((coll[k].Lecturer.Equals(mdlData.colLecturer[LectId])))
                                {
                                    if (!(coll[k].TutorialPractice == 0))
                                    {
                                        if ((coll[k].Subject.Subject == NamesTutorialPr[i]) &
                                            ((coll[k].Speciality.ShortUpravlenie + "-" + coll[k].Speciality.ShortInstitute + "-" + coll[k].KursNum.Kurs) == GroupsTutorialPr[j]))
                                        {
                                            flgAddRow = true;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (flgAddRow)
                    {
                        countRow++;
                    }
                }
            }
        }

        private static void MagistryAddRow(IList<clsDistribution> coll, int LectId, ref int countRow)
        {
            bool flgAddRow;

            for (int i = 0; i <= NamesMagistry.GetLength(0) - 1; i++)
            {
                for (int j = 0; j <= GroupsMagistry.GetLength(0) - 1; j++)
                {
                    flgAddRow = false;
                    for (int k = 0; k <= coll.Count - 1; k++)
                    {
                        if (coll[k].Semestr.SemNum.Equals("1 семестр"))
                        {
                            if (!(coll[k].Lecturer == null))
                            {
                                if ((coll[k].Lecturer.Equals(mdlData.colLecturer[LectId])))
                                {
                                    if (!(coll[k].Magistry == 0))
                                    {
                                        if ((coll[k].Subject.Subject == NamesMagistry[i]) &
                                            ((coll[k].Speciality.ShortUpravlenie + "-" + coll[k].KursNum.Kurs) == GroupsMagistry[j]))
                                        {
                                            flgAddRow = true;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (flgAddRow)
                    {
                        countRow++;
                    }

                    flgAddRow = false;
                    for (int k = 0; k <= coll.Count - 1; k++)
                    {
                        if (coll[k].Semestr.SemNum.Equals("2 семестр"))
                        {
                            if (!(coll[k].Lecturer == null))
                            {
                                if ((coll[k].Lecturer.Equals(mdlData.colLecturer[LectId])))
                                {
                                    if (!(coll[k].Magistry == 0))
                                    {
                                        if ((coll[k].Subject.Subject == NamesMagistry[i]) &
                                            ((coll[k].Speciality.ShortUpravlenie + "-" + coll[k].KursNum.Kurs) == GroupsMagistry[j]))
                                        {
                                            flgAddRow = true;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (flgAddRow)
                    {
                        countRow++;
                    }
                }
            }
        }

        //
        private static int countTableRowsMagistry(IList<clsDistribution> coll, int LectId)
        {
            int countRow;

            //Создаём строки
            //1. под шапку первого семестра
            //2. под надпись первого семестра
            //3. всего часов за семестр
            //4. пробел
            //5. под шапку второго семестра
            //6. под надпись второго семестра
            //7. всего часов за семестр
            //8. пробел
            //9. всего часов за учебный год

            countRow = 9;

            MagistryAddRow(coll, LectId, ref countRow);

            return countRow;
        }

        private static int countTableRowsHoured(IList<clsDistribution> coll, int LectId, int AddRow,
                                                clsSemestr S)
        {
            int countRow;

            //Создаём строки
            countRow = AddRow;

            //Формируем строки по нагрузке преподавателя в первом семестре
            //прогоняем все нагрузочные строки
            for (int i = 0; i <= coll.Count - 1; i++)
            {
                if (coll[i].Semestr.SemNum.Equals(S.SemNum))
                {
                    if (!(coll[i].Lecturer == null))
                    {
                        if ((coll[i].Lecturer.Equals(mdlData.colLecturer[LectId])))
                        {
                            //Если есть лекционные часы - добавляем строку
                            if (!(coll[i].Lecture == 0))
                            {
                                countRow++;
                            }

                            //Если есть экзаменационные часы - добавляем строку
                            if (!(coll[i].Exam == 0))
                            {
                                countRow++;
                            }

                            //Если есть зачётные часы - добавляем строку
                            if (!(coll[i].Credit == 0))
                            {
                                countRow++;
                            }

                            //Если есть часы на реферат - добавляем строку
                            if (!(coll[i].RefHomeWork == 0))
                            {
                                countRow++;
                            }

                            //Если есть часы на консультацию - добавляем строку
                            if (!(coll[i].Tutorial == 0))
                            {
                                countRow++;
                            }

                            //Если есть часы на лабораторные работы - добавляем строку
                            if (!(coll[i].LabWork == 0))
                            {
                                countRow++;
                            }

                            //Если есть часы на практические занятия - добавляем строку
                            if (!(coll[i].Practice == 0))
                            {
                                countRow++;
                            }

                            //Если есть часы на индивидуальные занятия - добавляем строку
                            if (!(coll[i].IndividualWork == 0))
                            {
                                countRow++;
                            }

                            //Если есть часы на КРАПК - добавляем строку
                            if (!(coll[i].KRAPK == 0))
                            {
                                countRow++;
                            }

                            //Если есть часы на курсовой проект - добавляем строку
                            if (!(coll[i].KursProject == 0))
                            {
                                countRow++;
                            }

                            //Если есть часы на преддипломную практику - добавляем строку
                            if (!(coll[i].PreDiplomaPractice == 0))
                            {
                                countRow++;
                            }

                            //Если есть часы на дипломный проект - добавляем строку
                            if (!(coll[i].DiplomaPaper == 0))
                            {
                                countRow++;
                            }

                            //Если есть часы на учебную практику - добавляем строку
                            if (!(coll[i].TutorialPractice == 0))
                            {
                                countRow++;
                            }

                            //Если есть часы на производственную практику - добавляем строку
                            if (!(coll[i].ProducingPractice == 0))
                            {
                                countRow++;
                            }

                            //Если есть часы на ГАК - добавляем строку
                            if (!(coll[i].GAK == 0))
                            {
                                countRow++;
                            }

                            //Если есть часы на Посещение занятий - добавляем строку
                            if (!(coll[i].Visiting == 0))
                            {
                                countRow++;
                            }

                            //Если есть часы на Аспирантуру - добавляем строку
                            if (!(coll[i].PostGrad == 0))
                            {
                                countRow++;
                            }

                            //Если есть часы на Руководство магистрами - добавляем строку
                            if (!(coll[i].Magistry == 0))
                            {
                                countRow++;
                            }
                        }
                    }
                }
            }

            return countRow;
        }

        private void DataGridHeader(string sem, ref int curRow)
        {
            //Заполняем первую строку и одновременно формируем размерности
            dgNagruzka.Columns[0].Width = 300;
            dgNagruzka[0, curRow].Value = "Наименование дисциплин";

            dgNagruzka.Columns[1].Width = 150;
            dgNagruzka[1, curRow].Value = "Объединение, группа";

            dgNagruzka.Columns[2].Width = 150;
            dgNagruzka[2, curRow].Value = "Вид занятий";

            dgNagruzka.Columns[3].Width = 150;
            dgNagruzka[3, curRow].Value = "Запланировано в часах";

            dgNagruzka.Columns[4].Width = 150;
            dgNagruzka[4, curRow].Value = "Факт. в часах";

            dgNagruzka.Columns[5].Width = 150;
            dgNagruzka[5, curRow].Value = "Примечание";

            curRow += 1;

            dgNagruzka[0, curRow].Value = sem;

            curRow += 1;
        }

        private void DataGridFiller(IList<clsDistribution> coll)
        {
            int i, j;

            int curRow;
            int Sum = 0;
            int sumCurrent, sumStud;

            curRow = 0;
            //Подготавливаем шапку первого семестра
            DataGridHeader("1 семестр", ref curRow);

            //Формируем нагрузку преподавателя в первом семестре по каждой дисциплине
            for (i = 0; i <= coll.Count - 1; i++)
            {
                if (!coll[i].flgExclude)
                {
                    Sum = 0;
                    if (coll[i].Semestr.SemNum.Equals("1 семестр"))
                    {
                        if (!(coll[i].Lecturer == null))
                        {
                            if ((coll[i].Lecturer.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])))
                            {
                                //Суммируем часы в строке
                                //Sum = mdlData.toSumDistributionComponentsWOCombine(coll[i]);
                                Sum = mdlData.toSumDistributionComponents(coll[i]);

                                //Если сумма больше нуля, значит, строка,
                                //выделенная под дисциплину не фиктивная
                                if (Sum > 0)
                                {
                                    dgNagruzka[0, curRow].Value = coll[i].Subject.Subject + " (" + Sum + " час.)";

                                    if (coll[i].Speciality != null & coll[i].KursNum != null)
                                    {
                                        dgNagruzka[1, curRow].Value = coll[i].Speciality.ShortUpravlenie
                                                                        + "-" + coll[i].KursNum.Kurs + " ("
                                                                        + coll[i].Speciality.ShortInstitute + "-" +
                                                                        coll[i].KursNum.Kurs + "1  )";
                                    }

                                    //Если есть лекционные часы
                                    if (!(coll[i].Lecture == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Лк.";
                                        dgNagruzka[3, curRow].Value = coll[i].Lecture;
                                        curRow += 1;
                                    }

                                    //Если есть экзаменационные часы
                                    if (!(coll[i].Exam == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Экз.";
                                        dgNagruzka[3, curRow].Value = coll[i].Exam;
                                        curRow += 1;
                                    }

                                    //Если есть зачётные часы
                                    if (!(coll[i].Credit == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Зач.";
                                        dgNagruzka[3, curRow].Value = coll[i].Credit;
                                        curRow += 1;
                                    }

                                    //Если есть часы на реферат
                                    if (!(coll[i].RefHomeWork == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Реф.";
                                        dgNagruzka[3, curRow].Value = coll[i].RefHomeWork;
                                        curRow += 1;
                                    }

                                    //Если есть часы на консультацию
                                    if (!(coll[i].Tutorial == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Конс.";
                                        dgNagruzka[3, curRow].Value = coll[i].Tutorial;
                                        curRow += 1;
                                    }

                                    //Если есть часы на лабораторные работы
                                    if (!(coll[i].LabWork == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Лаб.р.";
                                        dgNagruzka[3, curRow].Value = coll[i].LabWork;
                                        curRow += 1;
                                    }

                                    //Если есть часы на практические занятия
                                    if (!(coll[i].Practice == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Пр.";
                                        dgNagruzka[3, curRow].Value = coll[i].Practice;
                                        curRow += 1;
                                    }

                                    //Если есть часы на индивидуальное задание
                                    if (!(coll[i].IndividualWork == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Инд.";
                                        dgNagruzka[3, curRow].Value = coll[i].IndividualWork;
                                        curRow += 1;
                                    }

                                    //Если есть часы на КРАПК
                                    if (!(coll[i].KRAPK == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "КРАПК";
                                        dgNagruzka[3, curRow].Value = coll[i].KRAPK;
                                        curRow += 1;
                                    }

                                    //Если есть часы на курсовой проект
                                    if (!(coll[i].KursProject == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Курс.пр.";
                                        dgNagruzka[3, curRow].Value = coll[i].KursProject;
                                        curRow += 1;
                                    }

                                    //Если есть часы на учебную практику
                                    if (!(coll[i].TutorialPractice == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Пр.";
                                        dgNagruzka[3, curRow].Value = coll[i].TutorialPractice;
                                        curRow += 1;
                                    }

                                    //Если есть часы на производственную практику
                                    if (!(coll[i].ProducingPractice == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Пр.";
                                        dgNagruzka[3, curRow].Value = coll[i].ProducingPractice;
                                        curRow += 1;
                                    }

                                    //Если есть часы на магистратуру
                                    if (!(coll[i].Magistry == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Мг.";
                                        dgNagruzka[3, curRow].Value = coll[i].Magistry;
                                        curRow += 1;
                                    }

                                    //Если есть часы на преддипломную практику
                                    if (!(coll[i].PreDiplomaPractice == 0))
                                    {
                                        //dgNagruzka[2, curRow].Value = "Предд.";
                                        dgNagruzka[3, curRow].Value = coll[i].PreDiplomaPractice;
                                        curRow += 1;
                                    }

                                    //Если есть часы на ГЭК
                                    if (!(coll[i].GAK == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "ГАК";
                                        dgNagruzka[3, curRow].Value = coll[i].GAK;
                                        curRow += 1;
                                    }

                                    //Если есть часы на Аспирантуру
                                    if (!(coll[i].PostGrad == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "";
                                        dgNagruzka[3, curRow].Value = coll[i].PostGrad;
                                        curRow += 1;
                                    }

                                    //Если есть часы на Посещение занятий
                                    if (!(coll[i].Visiting == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "";
                                        dgNagruzka[3, curRow].Value = coll[i].Visiting;
                                        curRow += 1;
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (coll[i].flgDistrib)
                            {
                                sumCurrent = 0;
                                sumStud = 0;

                                for (j = 0; j <= mdlData.colStudents.Count - 1; j++)
                                {
                                    if (mdlData.colStudents[j].flgPlan)
                                    {
                                        //Если рассматриваемый преподаватель - руководитель студента
                                        //И если студент на том же курсе, где и дисциплина
                                        //И специальность студента должна соответствовать специальности нагрузки
                                        if (mdlData.colStudents[j].Lect.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])
                                            & mdlData.colStudents[j].KursNum.Equals(coll[i].KursNum)
                                            & mdlData.colStudents[j].Speciality.Equals(coll[i].Speciality))
                                        {
                                            sumCurrent += coll[i].Weight;
                                            sumStud++;
                                        }
                                    }
                                }

                                //
                                if (flgCombine)
                                {
                                    mdlData.toDetectUniformInHoured(ref sumCurrent, coll[i], mdlData.colLecturer[cmbLecturerList.SelectedIndex]);
                                }

                                //Если сумма изменилась
                                if (sumCurrent > 0)
                                {
                                    if (coll[i].Speciality != null & coll[i].KursNum != null)
                                    {
                                        dgNagruzka[0, curRow].Value = coll[i].Subject.Subject + " (" + coll[i].Weight + " час. на чел.; " + sumStud + " чел.)";

                                        dgNagruzka[1, curRow].Value = coll[i].Speciality.ShortUpravlenie
                                                                        + "-" + coll[i].KursNum.Kurs + " ("
                                                                        + coll[i].Speciality.ShortInstitute + "-" +
                                                                        coll[i].KursNum.Kurs + "1  )";
                                    }

                                    if (coll[i].Magistry > 0)
                                    {
                                        dgNagruzka[2, curRow].Value = "Мг.";
                                    }

                                    if (coll[i].PreDiplomaPractice > 0 ||
                                        coll[i].TutorialPractice > 0 ||
                                        coll[i].ProducingPractice > 0)
                                    {
                                        dgNagruzka[2, curRow].Value = "Пр.";
                                    }

                                    if (coll[i].DiplomaPaper > 0)
                                    {
                                        dgNagruzka[2, curRow].Value = "ВКР";
                                    }

                                    dgNagruzka[3, curRow].Value = sumCurrent;
                                    //Переходим к следующей строке
                                    curRow += 1;
                                }
                            }
                        }
                    }
                }
            }

            //Всего часов за семестр - надпись
            dgNagruzka[0, curRow].Value = "Всего часов за семестр:";
            //Всего часов за семестр - часы
            dgNagruzka[3, curRow].Value = txtSum1.Text;
            //Пробел
            curRow += 2;
            
            //Подготавливаем шапку второго семестра
            DataGridHeader("2 семестр", ref curRow);

            //Формируем нагрузку преподавателя в первом семесте по каждой дисциплине
            for (i = 0; i <= coll.Count - 1; i++)
            {
                if (!coll[i].flgExclude)
                {
                    Sum = 0;
                    if (coll[i].Semestr.SemNum.Equals("2 семестр"))
                    {
                        if (!(coll[i].Lecturer == null))
                        {
                            if ((coll[i].Lecturer.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])))
                            {
                                //Суммируем часы в строке
                                //Sum = mdlData.toSumDistributionComponentsWOCombine(coll[i]);
                                Sum = mdlData.toSumDistributionComponents(coll[i]);

                                //Если сумма больше нуля, значит, строка,
                                //выделенная под дисциплину не фиктивная
                                if (Sum > 0)
                                {
                                    if (coll[i].Speciality != null & coll[i].KursNum != null)
                                    {
                                        dgNagruzka[0, curRow].Value = coll[i].Subject.Subject + " (" + Sum + " час.)";

                                        dgNagruzka[1, curRow].Value = coll[i].Speciality.ShortUpravlenie
                                                                        + "-" + coll[i].KursNum.Kurs + " ("
                                                                        + coll[i].Speciality.ShortInstitute + "-" +
                                                                        coll[i].KursNum.Kurs + "1  )";
                                    }

                                    //Если есть лекционные часы
                                    if (!(coll[i].Lecture == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Лк.";
                                        dgNagruzka[3, curRow].Value = coll[i].Lecture;
                                        curRow += 1;
                                    }

                                    //Если есть экзаменационные часы
                                    if (!(coll[i].Exam == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Экз.";
                                        dgNagruzka[3, curRow].Value = coll[i].Exam;
                                        curRow += 1;
                                    }

                                    //Если есть зачётные часы
                                    if (!(coll[i].Credit == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Зач.";
                                        dgNagruzka[3, curRow].Value = coll[i].Credit;
                                        curRow += 1;
                                    }

                                    //Если есть часы на реферат
                                    if (!(coll[i].RefHomeWork == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Реф.";
                                        dgNagruzka[3, curRow].Value = coll[i].RefHomeWork;
                                        curRow += 1;
                                    }

                                    //Если есть часы на консультацию
                                    if (!(coll[i].Tutorial == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Конс.";
                                        dgNagruzka[3, curRow].Value = coll[i].Tutorial;
                                        curRow += 1;
                                    }

                                    //Если есть часы на лабораторные работы
                                    if (!(coll[i].LabWork == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Лаб.р.";
                                        dgNagruzka[3, curRow].Value = coll[i].LabWork;
                                        curRow += 1;
                                    }

                                    //Если есть часы на практические занятия
                                    if (!(coll[i].Practice == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Пр.";
                                        dgNagruzka[3, curRow].Value = coll[i].Practice;
                                        curRow += 1;
                                    }

                                    //Если есть часы на индивидуальное задание
                                    if (!(coll[i].IndividualWork == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Инд.";
                                        dgNagruzka[3, curRow].Value = coll[i].IndividualWork;
                                        curRow += 1;
                                    }

                                    //Если есть часы на КРАПК
                                    if (!(coll[i].KRAPK == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "КРАПК";
                                        dgNagruzka[3, curRow].Value = coll[i].KRAPK;
                                        curRow += 1;
                                    }

                                    //Если есть часы на курсовой проект
                                    if (!(coll[i].KursProject == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Курс.пр.";
                                        dgNagruzka[3, curRow].Value = coll[i].KursProject;
                                        curRow += 1;
                                    }

                                    //Если есть часы на учебную практику
                                    if (!(coll[i].TutorialPractice == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Пр.";
                                        dgNagruzka[3, curRow].Value = coll[i].TutorialPractice;
                                        curRow += 1;
                                    }

                                    //Если есть часы на производственную практику
                                    if (!(coll[i].ProducingPractice == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Пр.";
                                        dgNagruzka[3, curRow].Value = coll[i].ProducingPractice;
                                        curRow += 1;
                                    }

                                    //Если есть часы на магистратуру
                                    if (!(coll[i].Magistry == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "Мг.";
                                        dgNagruzka[3, curRow].Value = coll[i].Magistry;
                                        curRow += 1;
                                    }

                                    //Если есть часы на преддипломную практику
                                    if (!(coll[i].PreDiplomaPractice == 0))
                                    {
                                        //dgNagruzka[2, curRow].Value = "Предд.";
                                        dgNagruzka[3, curRow].Value = coll[i].PreDiplomaPractice;
                                        curRow += 1;
                                    }

                                    //Если есть часы на ГЭК
                                    if (!(coll[i].GAK == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "ГАК";
                                        dgNagruzka[3, curRow].Value = coll[i].GAK;
                                        curRow += 1;
                                    }

                                    //Если есть часы на Аспирантуру
                                    if (!(coll[i].PostGrad == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "";
                                        dgNagruzka[3, curRow].Value = coll[i].PostGrad;
                                        curRow += 1;
                                    }

                                    //Если есть часы на Посещение занятий
                                    if (!(coll[i].Visiting == 0))
                                    {
                                        dgNagruzka[2, curRow].Value = "";
                                        dgNagruzka[3, curRow].Value = coll[i].Visiting;
                                        curRow += 1;
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (coll[i].flgDistrib)
                            {
                                sumCurrent = 0;
                                sumStud = 0;
                                for (j = 0; j <= mdlData.colStudents.Count - 1; j++)
                                {
                                    if (mdlData.colStudents[j].flgPlan)
                                    {
                                        //Если рассматриваемый преподаватель - руководитель студента
                                        //И если студент на том же курсе, где и дисциплина
                                        //И специальность студента должна соответствовать специальности нагрузки
                                        if (mdlData.colStudents[j].Lect.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])
                                            & mdlData.colStudents[j].KursNum.Equals(coll[i].KursNum)
                                            & mdlData.colStudents[j].Speciality.Equals(coll[i].Speciality))
                                        {
                                            sumCurrent += coll[i].Weight;
                                            sumStud++;
                                        }
                                    }
                                }

                                //
                                if (flgCombine)
                                {
                                    mdlData.toDetectUniformInHoured(ref sumCurrent, coll[i], mdlData.colLecturer[cmbLecturerList.SelectedIndex]);
                                }

                                //Если сумма изменилась
                                if (sumCurrent > 0)
                                {
                                    if (coll[i].Speciality != null & coll[i].KursNum != null)
                                    {
                                        dgNagruzka[0, curRow].Value = coll[i].Subject.Subject + " (" + coll[i].Weight + " час. на чел.; " + sumStud + " чел.)";

                                        dgNagruzka[1, curRow].Value = coll[i].Speciality.ShortUpravlenie
                                                                        + "-" + coll[i].KursNum.Kurs + " ("
                                                                        + coll[i].Speciality.ShortInstitute + "-" +
                                                                        coll[i].KursNum.Kurs + "1  )";
                                    }

                                    if (coll[i].Magistry > 0)
                                    {
                                        dgNagruzka[2, curRow].Value = "Мг.";
                                    }

                                    if (coll[i].PreDiplomaPractice > 0 ||
                                        coll[i].TutorialPractice > 0 ||
                                        coll[i].ProducingPractice > 0)
                                    {
                                        dgNagruzka[2, curRow].Value = "Пр.";
                                    }

                                    if (coll[i].DiplomaPaper > 0)
                                    {
                                        dgNagruzka[2, curRow].Value = "ВКР";
                                    }

                                    dgNagruzka[3, curRow].Value = sumCurrent;
                                    //Переходим к следующей строке
                                    curRow += 1;
                                }
                            }
                        }
                    }
                }
            }

            //Всего часов за семестр - надпись
            dgNagruzka[0, curRow].Value = "Всего часов за семестр:";
            //Всего часов за семестр - часы
            dgNagruzka[3, curRow].Value = txtSum2.Text;
            //Переход на следующую строку
            curRow += 2;
            //Всего часов за год - надпись
            dgNagruzka[0, curRow].Value = "Всего часов за год:";
            //Всего часов за год - часы
            dgNagruzka[3, curRow].Value = txtSumAll.Text;
        }

        private void DiplomaLoad(IList<clsDistribution> coll, string sem, ref int curRow, ref int Sum)
        {
            int OneDiploma = 0;
            int CountDiploma;
            int DiplomaSum;

            for (int i = 0; i <= NamesDiploma.GetLength(0) - 1; i++)
            {
                for (int j = 0; j <= GroupsDiploma.GetLength(0) - 1; j++)
                {
                    DiplomaSum = 0;
                    CountDiploma = 0;
                    for (int k = 0; k <= coll.Count - 1; k++)
                    {
                        if (coll[k].Semestr.SemNum.Equals(sem))
                        {
                            if (!(coll[k].Lecturer == null))
                            {
                                if ((coll[k].Lecturer.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])))
                                {
                                    if (!(coll[k].DiplomaPaper == 0))
                                    {
                                        if ((coll[k].Subject.Subject == NamesDiploma[i]) &
                                            (coll[k].Speciality.ShortUpravlenie == GroupsDiploma[j]))
                                        {
                                            if (OneDiploma == 0)
                                            {
                                                OneDiploma = coll[k].DiplomaPaper;
                                            }

                                            CountDiploma++;
                                            DiplomaSum += coll[k].DiplomaPaper;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (DiplomaSum > 0)
                    {
                        dgNagruzka[0, curRow].Value = NamesDiploma[i] + " (" + OneDiploma.ToString() + " час.)";
                        dgNagruzka[1, curRow].Value = GroupsFullDiploma[j] + " (" + CountDiploma.ToString() + " чел.)";
                        dgNagruzka[3, curRow].Value = DiplomaSum.ToString();
                        curRow += 1;
                    }

                    Sum += DiplomaSum;
                }
            }
        }

        //На выходе только сведения о дипломных проектах
        private void DataGridFillerDiploma(IList<clsDistribution> coll)
        {
            int curRow;
            int Sum = 0;
            int SumI = 0;
            int SumII = 0;

            curRow = 0;
            
            //Подготавливаем шапку первого семестра
            DataGridHeader("1 семестр", ref curRow);

            //Формируем нагрузку преподавателя в первом семестре по каждой дисциплине

            DiplomaLoad(coll, "1 семестр", ref curRow, ref SumI);

            //Всего часов за семестр - надпись
            dgNagruzka[0, curRow].Value = "Всего часов за семестр:";
            //Всего часов за семестр - часы
            dgNagruzka[3, curRow].Value = SumI.ToString();
            //Пробел
            curRow += 2;

            //Подготавливаем шапку второго семестра
            DataGridHeader("2 семестр", ref curRow);

            DiplomaLoad(coll, "2 семестр", ref curRow, ref SumII);

            Sum = SumI + SumII;

            //Всего часов за семестр - надпись
            dgNagruzka[0, curRow].Value = "Всего часов за семестр:";
            //Всего часов за семестр - часы
            dgNagruzka[3, curRow].Value = SumII.ToString();
            //Переход на следующую строку
            curRow += 2;
            //Всего часов за год - надпись
            dgNagruzka[0, curRow].Value = "Всего часов за год:";
            //Всего часов за год - часы
            dgNagruzka[3, curRow].Value = Sum.ToString();
        }

        private void MagistryLoad(IList<clsDistribution> coll, string sem, ref int curRow, ref int Sum)
        {
            int OneMagistry = 0;
            int CountMagistry;
            int MagistrySum;
            
            for (int i = 0; i <= NamesMagistry.GetLength(0) - 1; i++)
            {
                for (int j = 0; j <= GroupsMagistry.GetLength(0) - 1; j++)
                {
                    MagistrySum = 0;
                    CountMagistry = 0;
                    for (int k = 0; k <= coll.Count - 1; k++)
                    {
                        if (coll[k].Semestr.SemNum.Equals(sem))
                        {
                            if (!(coll[k].Lecturer == null))
                            {
                                if ((coll[k].Lecturer.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])))
                                {
                                    if (!(coll[k].Magistry == 0))
                                    {
                                        if ((coll[k].Subject.Subject == NamesMagistry[i]) &
                                            ((coll[k].Speciality.ShortUpravlenie + "-" + coll[k].KursNum.Kurs) == GroupsMagistry[j]))
                                        {
                                            if (OneMagistry == 0)
                                            {
                                                OneMagistry = coll[k].Magistry;
                                            }

                                            CountMagistry++;
                                            MagistrySum += coll[k].Magistry;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (MagistrySum > 0)
                    {
                        dgNagruzka[0, curRow].Value = NamesMagistry[i] + " (" + OneMagistry.ToString() + " час.)";
                        dgNagruzka[1, curRow].Value = GroupsFullMagistry[j] + " (" + CountMagistry.ToString() + " чел.)";
                        dgNagruzka[3, curRow].Value = MagistrySum.ToString();
                        curRow += 1;
                    }

                    Sum += MagistrySum;
                }
            }
        }

        //
        private void TutorialPrLoad(IList<clsDistribution> coll, string sem, ref int curRow, ref int Sum)
        {
            int OneTutorialPr = 0;
            int CountTutorialPr;
            int TutorialPrSum;

            //
            for (int i = 0; i <= NamesTutorialPr.GetLength(0) - 1; i++)
            {
                //
                for (int j = 0; j <= GroupsTutorialPr.GetLength(0) - 1; j++)
                {
                    TutorialPrSum = 0;
                    CountTutorialPr = 0;
                    
                    //
                    for (int k = 0; k <= coll.Count - 1; k++)
                    {
                        if (coll[k].Semestr.SemNum.Equals(sem))
                        {
                            if (!(coll[k].Lecturer == null))
                            {
                                if ((coll[k].Lecturer.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])))
                                {
                                    if (!(coll[k].TutorialPractice == 0))
                                    {
                                        if ((coll[k].Subject.Subject == NamesTutorialPr[i]) &
                                            ((coll[k].Speciality.ShortUpravlenie + "-" + coll[k].Speciality.ShortInstitute + "-" + coll[k].KursNum.Kurs) == GroupsTutorialPr[j]))
                                        {
                                            if (OneTutorialPr == 0)
                                            {
                                                OneTutorialPr = coll[k].TutorialPractice;
                                            }

                                            CountTutorialPr++;
                                            TutorialPrSum += coll[k].TutorialPractice;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //
                    if (TutorialPrSum > 0)
                    {
                        dgNagruzka[0, curRow].Value = NamesTutorialPr[i] + " (" + OneTutorialPr.ToString() + " час.)";
                        dgNagruzka[1, curRow].Value = GroupsFullTutorialPr[j] + " (" + CountTutorialPr.ToString() + " чел.)";
                        dgNagruzka[3, curRow].Value = TutorialPrSum.ToString();
                        curRow += 1;
                    }

                    //
                    Sum += TutorialPrSum;
                }
            }
        }

        private void ProdPrLoad(IList<clsDistribution> coll, string sem, ref int curRow, ref int Sum)
        {
            int OneProdPr = 0;
            int CountProdPr;
            int ProdPrSum;

            for (int i = 0; i <= NamesProdPr.GetLength(0) - 1; i++)
            {
                for (int j = 0; j <= GroupsProdPr.GetLength(0) - 1; j++)
                {
                    ProdPrSum = 0;
                    CountProdPr = 0;
                    for (int k = 0; k <= coll.Count - 1; k++)
                    {
                        if (coll[k].Semestr.SemNum.Equals(sem))
                        {
                            if (!(coll[k].Lecturer == null))
                            {
                                if ((coll[k].Lecturer.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])))
                                {
                                    if (!(coll[k].ProducingPractice == 0))
                                    {
                                        if ((coll[k].Subject.Subject == NamesProdPr[i]) &
                                            ((coll[k].Speciality.ShortUpravlenie + "-" + coll[k].Speciality.ShortInstitute + "-" + coll[k].KursNum.Kurs) == GroupsProdPr[j]))
                                        {
                                            if (OneProdPr == 0)
                                            {
                                                OneProdPr = coll[k].ProducingPractice;
                                            }

                                            CountProdPr++;
                                            ProdPrSum += coll[k].ProducingPractice;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (ProdPrSum > 0)
                    {
                        dgNagruzka[0, curRow].Value = NamesProdPr[i] + " (" + OneProdPr.ToString() + " час.)";
                        dgNagruzka[1, curRow].Value = GroupsFullProdPr[j] + " (" + CountProdPr.ToString() + " чел.)";
                        dgNagruzka[3, curRow].Value = ProdPrSum.ToString();
                        curRow += 1;
                    }

                    Sum += ProdPrSum;
                }
            }
        }

        //На выходе только сведения по магистратуре
        private void DataGridFillerMagistry(IList<clsDistribution> coll)
        {
            int curRow;
            int Sum = 0;
            int SumI = 0;
            int SumII = 0;

            curRow = 0;
            //Подготавливаем шапку первого семестра
            DataGridHeader("1 семестр", ref curRow);

            //Формируем нагрузку преподавателя в первом семестре по каждой дисциплине
            MagistryLoad(coll, "1 семестр", ref curRow, ref SumI);

            //Всего часов за семестр - надпись
            dgNagruzka[0, curRow].Value = "Всего часов за семестр:";
            //Всего часов за семестр - часы
            dgNagruzka[3, curRow].Value = SumI.ToString();
            //Пробел
            curRow += 2;

            //Подготавливаем шапку второго семестра
            DataGridHeader("2 семестр", ref curRow);

            MagistryLoad(coll, "2 семестр", ref curRow, ref SumII);

            Sum = SumI + SumII;

            //Всего часов за семестр - надпись
            dgNagruzka[0, curRow].Value = "Всего часов за семестр:";
            //Всего часов за семестр - часы
            dgNagruzka[3, curRow].Value = SumII.ToString();
            //Переход на следующую строку
            curRow += 2;
            //Всего часов за год - надпись
            dgNagruzka[0, curRow].Value = "Всего часов за год:";
            //Всего часов за год - часы
            dgNagruzka[3, curRow].Value = Sum.ToString();
        }

        private void ProdPrLoadWord(IList<clsDistribution> coll, Word.Table ObjTable, string sem, ref int curRow, ref int Sum)
        {
            int Summary = 0;
            int Count = 0;

            for (int i = 0; i <= NamesProdPr.GetLength(0) - 1; i++)
            {
                for (int j = 0; j <= GroupsProdPr.GetLength(0) - 1; j++)
                {
                    Sum = 0;
                    Count = 0;
                    for (int k = 0; k <= coll.Count - 1; k++)
                    {
                        if (coll[k].Semestr.SemNum.Equals(sem))
                        {
                            if (!(coll[k].Lecturer == null))
                            {
                                if ((coll[k].Lecturer.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])))
                                {
                                    if (!(coll[k].ProducingPractice == 0))
                                    {
                                        if ((coll[k].Subject.Subject == NamesProdPr[i]) &
                                            (coll[k].Speciality.ShortUpravlenie + "-" + coll[k].Speciality.ShortInstitute + "-" + coll[k].KursNum.Kurs == GroupsProdPr[j]))
                                        {
                                            Sum += coll[k].ProducingPractice;
                                            Count++;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (Sum > 0)
                    {
                        ObjTable.Cell(curRow, 1).Range.Text = NamesProdPr[i] + " (" + Count.ToString() + " чел.)";
                        ObjTable.Cell(curRow, 1).Range.Font.Italic = 1;
                        ObjTable.Cell(curRow, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                        ObjTable.Cell(curRow, 2).Range.Text = GroupsFullProdPr[j];
                        ObjTable.Cell(curRow, 4).Range.Text = Sum.ToString();
                        curRow += 1;

                        Summary += Sum;
                    }
                }
            }

            Sum = Summary;
        }

        private void TutorialPrLoadWord(IList<clsDistribution> coll, Word.Table ObjTable, string sem, ref int curRow, ref int Sum)
        {
            int Summary = 0;
            int Count = 0;

            for (int i = 0; i <= NamesTutorialPr.GetLength(0) - 1; i++)
            {
                for (int j = 0; j <= GroupsTutorialPr.GetLength(0) - 1; j++)
                {
                    Sum = 0;
                    Count = 0;
                    for (int k = 0; k <= coll.Count - 1; k++)
                    {
                        if (coll[k].Semestr.SemNum.Equals(sem))
                        {
                            if (!(coll[k].Lecturer == null))
                            {
                                if ((coll[k].Lecturer.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])))
                                {
                                    if (!(coll[k].TutorialPractice == 0))
                                    {
                                        if ((coll[k].Subject.Subject == NamesTutorialPr[i]) &
                                            (coll[k].Speciality.ShortUpravlenie + "-" + coll[k].Speciality.ShortInstitute + "-" + coll[k].KursNum.Kurs == GroupsTutorialPr[j]))
                                        {
                                            Sum += coll[k].TutorialPractice;
                                            Count++;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (Sum > 0)
                    {
                        ObjTable.Cell(curRow, 1).Range.Text = NamesTutorialPr[i] + " (" + Count.ToString() + " чел.)";
                        ObjTable.Cell(curRow, 1).Range.Font.Italic = 1;
                        ObjTable.Cell(curRow, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                        ObjTable.Cell(curRow, 2).Range.Text = GroupsFullTutorialPr[j];
                        ObjTable.Cell(curRow, 4).Range.Text = Sum.ToString();
                        curRow += 1;

                        Summary += Sum;
                    }
                }
            }

            Sum = Summary;
        }

        private void DiplomaLoadWord(IList<clsDistribution> coll, Word.Table ObjTable, string sem, ref int curRow, ref int Sum)
        {
            int Summary = 0;
            int Count = 0;

            for (int i = 0; i <= NamesDiploma.GetLength(0) - 1; i++)
            {
                for (int j = 0; j <= GroupsDiploma.GetLength(0) - 1; j++)
                {
                    Sum = 0;
                    Count = 0;
                    for (int k = 0; k <= coll.Count - 1; k++)
                    {
                        if (coll[k].Semestr.SemNum.Equals(sem))
                        {
                            if (!(coll[k].Lecturer == null))
                            {
                                if ((coll[k].Lecturer.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])))
                                {
                                    if (!(coll[k].DiplomaPaper == 0))
                                    {
                                        if ((coll[k].Subject.Subject == NamesDiploma[i]) &
                                            (coll[k].Speciality.ShortUpravlenie == GroupsDiploma[j]))
                                        {
                                            Sum += coll[k].DiplomaPaper;
                                            Count++;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (Sum > 0)
                    {
                        ObjTable.Cell(curRow, 1).Range.Text = NamesDiploma[i] + " (" + Count.ToString() + " чел.)";
                        ObjTable.Cell(curRow, 1).Range.Font.Italic = 1;
                        ObjTable.Cell(curRow, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                        ObjTable.Cell(curRow, 2).Range.Text = GroupsFullDiploma[j];
                        ObjTable.Cell(curRow, 4).Range.Text = Sum.ToString();
                        curRow += 1;

                        Summary += Sum;
                    }
                }
            }

            Sum = Summary;
        }

        private void MagistryLoadWord(IList<clsDistribution> coll, Word.Table ObjTable, string sem, ref int curRow, ref int Sum)
        {
            int Summary = 0;
            int Count = 0;

            for (int i = 0; i <= NamesMagistry.GetLength(0) - 1; i++)
            {
                for (int j = 0; j <= GroupsMagistry.GetLength(0) - 1; j++)
                {
                    Sum = 0;
                    Count = 0;
                    for (int k = 0; k <= coll.Count - 1; k++)
                    {
                        if (coll[k].Semestr.SemNum.Equals(sem))
                        {
                            if (!(coll[k].Lecturer == null))
                            {
                                if ((coll[k].Lecturer.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])))
                                {
                                    if (!(coll[k].Magistry == 0))
                                    {
                                        if ((coll[k].Subject.Subject == NamesMagistry[i]) &
                                            (coll[k].Speciality.ShortUpravlenie + "-" + coll[k].KursNum.Kurs == GroupsMagistry[j]))
                                        {
                                            Sum += coll[k].Magistry;
                                            Count++;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (Sum > 0)
                    {
                        ObjTable.Cell(curRow, 1).Range.Text = NamesMagistry[i] + " (" + Count.ToString() + " чел.)";
                        ObjTable.Cell(curRow, 1).Range.Font.Italic = 1;
                        ObjTable.Cell(curRow, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                        ObjTable.Cell(curRow, 2).Range.Text = GroupsFullMagistry[j];
                        ObjTable.Cell(curRow, 4).Range.Text = Sum.ToString();
                        curRow += 1;

                        Summary += Sum;
                    }
                }
            }

            Sum = Summary;
        }

        //
        private void WordTableHeader(IList<clsDistribution> coll, Word.Table ObjTable, string sem, ref int curRow)
        {
            //Заполняем первую строку и одновременно формируем размерности
            ObjTable.Cell(curRow, 1).Range.Text = "Наименование дисциплин";

            ObjTable.Cell(curRow, 2).Range.Text = "Объединение, группа";

            ObjTable.Cell(curRow, 3).Range.Text = "Вид занятий";

            ObjTable.Cell(curRow, 4).Range.Text = "Запланировано в часах";

            ObjTable.Cell(curRow, 5).Range.Text = "Факт. в часах";

            ObjTable.Cell(curRow, 6).Range.Text = "Примечание";

            curRow += 1;

            ObjTable.Cell(curRow, 1).Range.Text = sem;
            ObjTable.Cell(curRow, 1).Range.Font.Italic = 1;
            ObjTable.Cell(curRow, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            curRow += 1;
        }

        //
        private void WordTableFiller(IList<clsDistribution> coll, Word.Table ObjTable, object ObjMissing, Word._Document ObjDoc)
        {
            int curRow;
            int countStud;
            int Sum;
            int SumI = 0;
            int SumII = 0;
            //int DiplomaSum = 0;
            //int MagistrySum = 0;
            //int TutorialPrSum = 0;
            //int ProdPrSum = 0;
            int i, j;

            curRow = 1;

            WordTableHeader(coll, ObjTable, "1 семестр", ref curRow);

            //Формируем нагрузку преподавателя в первом семестре по каждой дисциплине
            for (i = 0; i <= coll.Count - 1; i++)
            {
                Sum = 0;
                countStud = 0;
                if (coll[i].Semestr.SemNum.Equals("1 семестр"))
                {
                    if (!coll[i].flgExclude)
                    {
                        if (!(coll[i].Lecturer == null))
                        {
                            if ((coll[i].Lecturer.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])))
                            {
                                //Суммируем часы в строке
                                //Sum = mdlData.toSumDistributionComponentsWOCombine(coll[i]);
                                Sum = mdlData.toSumDistributionComponents(coll[i]);

                                //Если сумма больше нуля, значит, строка,
                                //выделенная под дисциплину не фиктивная
                                if (Sum > 0)
                                {
                                    ObjTable.Cell(curRow, 1).Range.Text = coll[i].Subject.Subject;
                                    ObjTable.Cell(curRow, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                                    ObjTable.Cell(curRow, 1).Range.Font.Italic = 1;

                                    if (!(coll[i].KursNum == null) & !(coll[i].Speciality == null))
                                    {
                                        ObjTable.Cell(curRow, 2).Range.Text = coll[i].Speciality.ShortUpravlenie
                                                                            + "-" + coll[i].KursNum.Kurs + "\n ("
                                                                            + coll[i].Speciality.ShortInstitute + "-" +
                                                                            coll[i].KursNum.Kurs + "1  )";
                                    }

                                    //Если есть лекционные часы
                                    if (!(coll[i].Lecture == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Лк.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].Lecture.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть экзаменационные часы
                                    if (!(coll[i].Exam == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Экз.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].Exam.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть зачётные часы
                                    if (!(coll[i].Credit == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Зач.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].Credit.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на реферат
                                    if (!(coll[i].RefHomeWork == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Реф.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на консультацию
                                    if (!(coll[i].Tutorial == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Конс.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].Tutorial.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на лабораторные работы
                                    if (!(coll[i].LabWork == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Лаб.р.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].LabWork.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на практические занятия
                                    if (!(coll[i].Practice == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Пр.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].Practice.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на индивидуальное задание
                                    if (!(coll[i].IndividualWork == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Инд.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].IndividualWork.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на КРАПК
                                    if (!(coll[i].KRAPK == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "КРАПК";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].KRAPK.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на курсовой проект
                                    if (!(coll[i].KursProject == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Курс.пр.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].KursProject.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на преддипломную практику
                                    if (!(coll[i].PreDiplomaPractice == 0))
                                    {
                                        //ObjTable.Cell(curRow, 3).Range.Text = "Предд.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].PreDiplomaPractice.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на ГЭК
                                    if (!(coll[i].GAK == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "ГАК";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].GAK.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на Учебную практику
                                    if (!(coll[i].TutorialPractice == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Пр.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].TutorialPractice.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на Производственную практику
                                    if (!(coll[i].ProducingPractice == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Пр.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].ProducingPractice.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на Магистратуру
                                    if (!(coll[i].Magistry == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Мг.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].Magistry.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на Аспирантуру
                                    if (!(coll[i].PostGrad == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].PostGrad.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на Посещение занятий
                                    if (!(coll[i].Visiting == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].Visiting.ToString();
                                        curRow += 1;
                                    }
                                }

                                SumI += Sum;
                            }
                        }
                        else
                        {
                            if (coll[i].flgDistrib)
                            {
                                for (j = 0; j <= mdlData.colStudents.Count - 1; j++)
                                {
                                    if (mdlData.colStudents[j].flgPlan)
                                    {
                                        //Если рассматриваемый преподаватель - руководитель студента
                                        //И если студент на том же курсе, где и дисциплина
                                        //И специальность студента должна соответствовать специальности нагрузки
                                        if (mdlData.colStudents[j].Lect.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])
                                            & mdlData.colStudents[j].KursNum.Equals(coll[i].KursNum)
                                            & mdlData.colStudents[j].Speciality.Equals(coll[i].Speciality))
                                        {
                                            Sum += coll[i].Weight;
                                            countStud++;
                                        }
                                    }
                                }

                                //
                                if (flgCombine)
                                {
                                    mdlData.toDetectUniformInHoured(ref Sum, coll[i], mdlData.colLecturer[cmbLecturerList.SelectedIndex]);
                                }
                            }

                            if (Sum > 0)
                            {
                                ObjTable.Cell(curRow, 1).Range.Text = coll[i].Subject.Subject;
                                ObjTable.Cell(curRow, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                                ObjTable.Cell(curRow, 1).Range.Font.Italic = 1;                                
                                
                                if (coll[i].Speciality != null & coll[i].KursNum != null)
                                {
                                    ObjTable.Cell(curRow, 1).Range.Text = coll[i].Subject.Subject + " (" + coll[i].Weight + " час. на чел.; " + countStud + " чел.)";

                                    ObjTable.Cell(curRow, 2).Range.Text = coll[i].Speciality.ShortUpravlenie
                                                                    + "-" + coll[i].KursNum.Kurs + " ("
                                                                    + coll[i].Speciality.ShortInstitute + "-" +
                                                                    coll[i].KursNum.Kurs + "1  )";
                                }

                                if (coll[i].Magistry > 0)
                                {
                                    ObjTable.Cell(curRow, 3).Range.Text = "Мг.";
                                }

                                if (coll[i].PreDiplomaPractice > 0 ||
                                    coll[i].TutorialPractice > 0 ||
                                    coll[i].ProducingPractice > 0)
                                {
                                    ObjTable.Cell(curRow, 3).Range.Text = "Пр.";
                                }

                                if (coll[i].DiplomaPaper > 0)
                                {
                                    ObjTable.Cell(curRow, 3).Range.Text = "ВКР";
                                }

                                ObjTable.Cell(curRow, 4).Range.Text = Sum.ToString();
                                //Переходим к следующей строке
                                curRow += 1;

                                SumI += Sum;
                            }
                        }
                    }
                }
            }

            //TutorialPrLoadWord(coll, ObjTable, "1 семестр", ref curRow, ref TutorialPrSum);
            //SumI += TutorialPrSum;

            //ProdPrLoadWord(coll, ObjTable, "1 семестр", ref curRow, ref ProdPrSum);
            //SumI += ProdPrSum;

            //DiplomaLoadWord(coll, ObjTable, "1 семестр", ref curRow, ref DiplomaSum);
            //SumI += DiplomaSum;

            //MagistryLoadWord(coll, ObjTable, "1 семестр", ref curRow, ref MagistrySum);
            //SumI += MagistrySum;

            ObjTable.Rows[curRow].Borders[Word.WdBorderType.wdBorderLeft].Visible = false;
            ObjTable.Rows[curRow].Borders[Word.WdBorderType.wdBorderRight].Visible = false;
            ObjTable.Rows[curRow].Borders[Word.WdBorderType.wdBorderVertical].Visible = false;
            ObjTable.Rows[curRow].Borders[Word.WdBorderType.wdBorderBottom].Visible = false;

            //Всего часов за семестр - надпись
            ObjTable.Cell(curRow, 1).Range.Text = "Всего часов за семестр:";
            ObjTable.Cell(curRow, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(curRow, 1).Range.Font.Italic = 1;
            //Всего часов за семестр - часы
            ObjTable.Cell(curRow, 4).Range.Text = SumI.ToString();
            //Пробел
            curRow += 1;

            ObjTable.Rows[curRow].Borders[Word.WdBorderType.wdBorderLeft].Visible = false;
            ObjTable.Rows[curRow].Borders[Word.WdBorderType.wdBorderRight].Visible = false;
            ObjTable.Rows[curRow].Borders[Word.WdBorderType.wdBorderVertical].Visible = false;

            //Пробел
            curRow += 1;

            //pageBreaker(ObjMissing, ObjDoc);

            WordTableHeader(coll, ObjTable, "2 семестр", ref curRow);

            //Формируем нагрузку преподавателя в первом семесте по каждой дисциплине
            for (i = 0; i <= coll.Count - 1; i++)
            {
                Sum = 0;
                countStud = 0;
                if (coll[i].Semestr.SemNum.Equals("2 семестр"))
                {
                    if (!coll[i].flgExclude)
                    {
                        if (!(coll[i].Lecturer == null))
                        {
                            if ((coll[i].Lecturer.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])))
                            {
                                //Суммируем часы в строке
                                //Sum = mdlData.toSumDistributionComponentsWOCombine(coll[i]);
                                Sum = mdlData.toSumDistributionComponents(coll[i]);

                                //Если сумма больше нуля, значит, строка,
                                //выделенная под дисциплину не фиктивная
                                if (Sum > 0)
                                {
                                    ObjTable.Cell(curRow, 1).Range.Text = coll[i].Subject.Subject;
                                    ObjTable.Cell(curRow, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                                    ObjTable.Cell(curRow, 1).Range.Font.Italic = 1;

                                    if (!(coll[i].KursNum == null) & !(coll[i].Speciality == null))
                                    {
                                        ObjTable.Cell(curRow, 2).Range.Text = coll[i].Speciality.ShortUpravlenie
                                           + "-" + coll[i].KursNum.Kurs + "\n ("
                                           + coll[i].Speciality.ShortInstitute + "-" +
                                           coll[i].KursNum.Kurs + "1  )";
                                    }

                                    //Если есть лекционные часы
                                    if (!(coll[i].Lecture == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Лк.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].Lecture.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть экзаменационные часы
                                    if (!(coll[i].Exam == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Экз.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].Exam.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть зачётные часы
                                    if (!(coll[i].Credit == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Зач.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].Credit.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на реферат
                                    if (!(coll[i].RefHomeWork == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Реф.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].RefHomeWork.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на консультацию
                                    if (!(coll[i].Tutorial == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Конс.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].Tutorial.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на лабораторные работы
                                    if (!(coll[i].LabWork == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Лаб.р.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].LabWork.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на практические занятия
                                    if (!(coll[i].Practice == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Пр.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].Practice.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на индивидуальное задание
                                    if (!(coll[i].IndividualWork == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Инд.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].IndividualWork.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на КРАПК
                                    if (!(coll[i].KRAPK == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "КРАПК";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].KRAPK.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на курсовой проект
                                    if (!(coll[i].KursProject == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Курс.пр.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].KursProject.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на преддипломную практику
                                    if (!(coll[i].PreDiplomaPractice == 0))
                                    {
                                        //ObjTable.Cell(curRow, 3).Range.Text = "Предд.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].PreDiplomaPractice.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на ГЭК
                                    if (!(coll[i].GAK == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "ГАК";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].GAK.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на Учебную практику
                                    if (!(coll[i].TutorialPractice == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Пр.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].TutorialPractice.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на Производственную практику
                                    if (!(coll[i].ProducingPractice == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Пр.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].ProducingPractice.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на Магистратуру
                                    if (!(coll[i].Magistry == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Мг.";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].Magistry.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на Аспирантуру
                                    if (!(coll[i].PostGrad == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].PostGrad.ToString();
                                        curRow += 1;
                                    }

                                    //Если есть часы на Посещение занятий
                                    if (!(coll[i].Visiting == 0))
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "";
                                        ObjTable.Cell(curRow, 4).Range.Text = coll[i].Visiting.ToString();
                                        curRow += 1;
                                    }

                                }

                                SumII += Sum;
                            }
                        }
                        else
                        {
                            if (coll[i].flgDistrib)
                            {
                                for (j = 0; j <= mdlData.colStudents.Count - 1; j++)
                                {
                                    if (mdlData.colStudents[j].flgPlan)
                                    {
                                        //Если рассматриваемый преподаватель - руководитель студента
                                        //И если студент на том же курсе, где и дисциплина
                                        //И специальность студента должна соответствовать специальности нагрузки
                                        if (mdlData.colStudents[j].Lect.Equals(mdlData.colLecturer[cmbLecturerList.SelectedIndex])
                                            & mdlData.colStudents[j].KursNum.Equals(coll[i].KursNum)
                                            & mdlData.colStudents[j].Speciality.Equals(coll[i].Speciality))
                                        {
                                            Sum += coll[i].Weight;
                                            countStud++;
                                        }
                                    }
                                }

                                //
                                if (flgCombine)
                                {
                                    mdlData.toDetectUniformInHoured(ref Sum, coll[i], mdlData.colLecturer[cmbLecturerList.SelectedIndex]);
                                }

                                if (Sum > 0)
                                {
                                    ObjTable.Cell(curRow, 1).Range.Text = coll[i].Subject.Subject;
                                    ObjTable.Cell(curRow, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                                    ObjTable.Cell(curRow, 1).Range.Font.Italic = 1;                                    
                                    
                                    if (coll[i].Speciality != null & coll[i].KursNum != null)
                                    {
                                        ObjTable.Cell(curRow, 1).Range.Text = coll[i].Subject.Subject + " (" + coll[i].Weight + " час. на чел.; " + countStud + " чел.)";

                                        ObjTable.Cell(curRow, 2).Range.Text = coll[i].Speciality.ShortUpravlenie
                                                                        + "-" + coll[i].KursNum.Kurs + " ("
                                                                        + coll[i].Speciality.ShortInstitute + "-" +
                                                                        coll[i].KursNum.Kurs + "1  )";
                                    }

                                    if (coll[i].Magistry > 0)
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Мг.";
                                    }

                                    if (coll[i].PreDiplomaPractice > 0 ||
                                        coll[i].TutorialPractice > 0 ||
                                        coll[i].ProducingPractice > 0)
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "Пр.";
                                    }

                                    if (coll[i].DiplomaPaper > 0)
                                    {
                                        ObjTable.Cell(curRow, 3).Range.Text = "ВКР";
                                    }

                                    ObjTable.Cell(curRow, 4).Range.Text = Sum.ToString();
                                    //Переходим к следующей строке
                                    curRow += 1;

                                    SumII += Sum;
                                }
                            }
                        }
                    }
                }
            }

            //TutorialPrLoadWord(coll, ObjTable, "2 семестр", ref curRow, ref TutorialPrSum);
            //SumII += TutorialPrSum;

            //ProdPrLoadWord(coll, ObjTable, "2 семестр", ref curRow, ref ProdPrSum);
            //SumII += ProdPrSum;

            //DiplomaLoadWord(coll, ObjTable, "2 семестр", ref curRow, ref DiplomaSum);
            //SumII += DiplomaSum;

            //MagistryLoadWord(coll, ObjTable, "2 семестр", ref curRow, ref MagistrySum);
            //SumII += MagistrySum;

            ObjTable.Rows[curRow].Borders[Word.WdBorderType.wdBorderLeft].Visible = false;
            ObjTable.Rows[curRow].Borders[Word.WdBorderType.wdBorderRight].Visible = false;
            ObjTable.Rows[curRow].Borders[Word.WdBorderType.wdBorderVertical].Visible = false;
            ObjTable.Rows[curRow].Borders[Word.WdBorderType.wdBorderBottom].Visible = false;

            //Всего часов за семестр - надпись
            ObjTable.Cell(curRow, 1).Range.Text = "Всего часов за семестр:";
            ObjTable.Cell(curRow, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(curRow, 1).Range.Font.Italic = 1;
            //Всего часов за семестр - часы
            ObjTable.Cell(curRow, 4).Range.Text = SumII.ToString();
            //Переход на следующую строку
            curRow += 1;

            ObjTable.Rows[curRow].Borders[Word.WdBorderType.wdBorderLeft].Visible = false;
            ObjTable.Rows[curRow].Borders[Word.WdBorderType.wdBorderRight].Visible = false;
            ObjTable.Rows[curRow].Borders[Word.WdBorderType.wdBorderVertical].Visible = false;

            curRow += 1;
            //Всего часов за год - надпись
            ObjTable.Cell(curRow, 1).Range.Text = "Всего часов за год:\n";
            ObjTable.Cell(curRow, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjTable.Cell(curRow, 1).Range.Font.Italic = 1;

            Sum = SumI + SumII;
            //Всего часов за год - часы
            ObjTable.Cell(curRow, 4).Range.Text = Sum.ToString() + "\n";
            ObjTable.Cell(curRow, 4).Range.Font.Bold = 1;
        }

        private void chkAllLoad_CheckedChanged(object sender, EventArgs e)
        {
            FillLecturerList();
        }

        private void chkOnlyUMO_CheckedChanged(object sender, EventArgs e)
        {
            FillLecturerList();
        }

        private void cmbForm_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
