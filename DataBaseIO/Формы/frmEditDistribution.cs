using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DataBaseIO
{
    
    public partial class frmEditDistribution : Form
    {
        private int curNum;
        private int MaxNum;
        private static IList<clsDistribution> Selected = null;
        
        public frmEditDistribution()
        {
            InitializeComponent();
        }

        private void CreateMenu()
        {
            //Создаём новый объект "Полоса меню"
            MenuStrip mnuDistribution = new MenuStrip();

            //-------------------------Формирование шапки меню-----------------

            //Создаём новый объект "Элемент полосы меню" с именем "Основное"
            ToolStripMenuItem mnuMain = new ToolStripMenuItem("Основное");
            //Создаём новый объект "Элемент полосы меню" с именем "Редактирование"
            ToolStripMenuItem mnuEdit = new ToolStripMenuItem("Редактирование");
            //Создаём новый объект "Элемент полосы меню" с именем "Операции"
            ToolStripMenuItem mnuOper = new ToolStripMenuItem("Операции");
            //Создаём новый объект "Элемент полосы меню" с именем "Отметки"
            ToolStripMenuItem mnuChecks = new ToolStripMenuItem("Отметки");
            //Создаём новый объект "Элемент полосы меню" с именем "Переход к"
            ToolStripMenuItem mnuNavigate = new ToolStripMenuItem("Переход к");

            //-------------------------Формирование шапки меню-----------------

            //-----------------Формирование пунктов "Основное"-----------------
            
            //Создаём новый объект "Элемент полосы меню" с именем "Закрыть"
            ToolStripMenuItem mnuClose = new ToolStripMenuItem("Закрыть");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Закрыть"
            mnuClose.Click += new EventHandler(mnuClose_Click);
            //Создаём новый объект "Элемент полосы меню" с именем "Сохранить"
            ToolStripMenuItem mnuSave = new ToolStripMenuItem("Сохранить");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Сохранить"
            mnuSave.Click += new EventHandler(mnuSave_Click);
            mnuSave.ShortcutKeys = Keys.Control | Keys.S;
            //Создаём новый объект "Элемент полосы меню" с именем "Удалить"
            ToolStripMenuItem mnuDel = new ToolStripMenuItem("Удалить");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Удалить"
            mnuDel.Click += new EventHandler(mnuDel_Click);
            //Создаём новый объект "Элемент полосы меню" с именем "Создать"
            ToolStripMenuItem mnuCreate = new ToolStripMenuItem("Создать");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Создать"
            mnuCreate.Click += new EventHandler(mnuCreate_Click);
            //Создаём новый объект "Элемент полосы меню" с именем "Скопировать"
            ToolStripMenuItem mnuCopy = new ToolStripMenuItem("Скопировать");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Скопировать"
            mnuCopy.Click += new EventHandler(mnuCopy_Click);

            //-----------------Формирование пунктов "Основное"-----------------

            //------------Добавление пунктов "Основное" к меню-----------------

            //Добавляем в меню "Основное" элемент "Создать"
            mnuMain.DropDownItems.Add(mnuCreate);
            //Добавляем в меню "Основное" элемент "Копировать"
            mnuMain.DropDownItems.Add(mnuCopy);
            //Добавляем в меню "Основное" элемент "Сохранить"
            mnuMain.DropDownItems.Add(mnuSave);
            //Добавляем в меню "Основное" элемент "Удалить"
            mnuMain.DropDownItems.Add(mnuDel);
            //Добавляем в меню "Основное" элемент "Закрыть"
            mnuMain.DropDownItems.Add(mnuClose);

            //------------Добавление пунктов "Основное" к меню-----------------

            //-----------Формирование пунктов "Редактирование"-----------------

            //Создаём новый объект "Элемент полосы меню" с именем "Редактирование дисциплин"
            ToolStripMenuItem mnuEditSubj = new ToolStripMenuItem("Редактирование дисциплин");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Редактирование дисциплин"
            mnuEditSubj.Click += new EventHandler(mnuEditSubj_Click);

            //-----------Формирование пунктов "Редактирование"-----------------

            //-----------Добавление пунктов "Редактирование" к меню------------

            //Добавляем в меню "Редактирование" элемент "Редактирование дисциплин"
            mnuEdit.DropDownItems.Add(mnuEditSubj);

            //-----------Добавление пунктов "Редактирование" к меню------------

            //-----------------Формирование пунктов "Операции"-----------------

            //Создаём новый объект "Элемент полосы меню" с именем "Объединить строки дисц., закреплённых за одними людьми"
            ToolStripMenuItem mnuMergeSame = new ToolStripMenuItem("Объединить строки дисц., закреплённых за одними людьми");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Объединить строки дисц., закреплённых за одними людьми"
            mnuMergeSame.Click += new EventHandler(mnuMergeSame_Click);

            //Создаём новый объект "Элемент полосы меню" с именем "Удалить пустые строки нагрузки"
            ToolStripMenuItem mnuDelNull = new ToolStripMenuItem("Удалить пустые строки нагрузки");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Удалить пустые строки нагрузки"
            mnuDelNull.Click += new EventHandler(mnuDelNull_Click);

            //Создаём новый объект "Элемент полосы меню" с именем "Обнулить все часы в строках нагрузки"
            ToolStripMenuItem mnuClearHours = new ToolStripMenuItem("Обнулить все часы в строках нагрузки");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Обнулить все часы в строках нагрузки"
            mnuClearHours.Click += new EventHandler(mnuClearHours_Click);

            //-----------------Формирование пунктов "Операции"-----------------

            //------------Добавление пунктов "Операции" к меню-----------------

            //Добавляем в меню "Операции" элемент "Объединить строки дисц., закреплённых за одними людьми"
            mnuOper.DropDownItems.Add(mnuMergeSame);
            //Добавляем в меню "Операции" элемент "Удалить пустые строки нагрузки"
            mnuOper.DropDownItems.Add(mnuDelNull);
            //Добавляем в меню "Операции" элемент "Обнулить все часы в строках нагрузки"
            mnuOper.DropDownItems.Add(mnuClearHours);

            //------------Добавление пунктов "Операции" к меню-----------------

            //-----------------Формирование пунктов "Отметки"-----------------

            //Создаём новый объект "Элемент полосы меню" с именем "Всё в диспетчерскую"
            ToolStripMenuItem mnuDispatch = new ToolStripMenuItem("Всё в диспетчерскую");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Всё в диспетчерскую"
            mnuDispatch.Click += new EventHandler(mnuDispatch_Click);
            //Создаём новый объект "Элемент полосы меню" с именем "Ничего в диспетчерскую"
            ToolStripMenuItem mnuUnDispatch = new ToolStripMenuItem("Ничего в диспетчерскую");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Ничего в диспетчерскую"
            mnuUnDispatch.Click += new EventHandler(mnuUnDispatch_Click);
            //Создаём новый объект "Элемент полосы меню" с именем "Ничего в расчёт"
            ToolStripMenuItem mnuExclude = new ToolStripMenuItem("Ничего в расчёт");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Ничего в расчёт"
            mnuExclude.Click += new EventHandler(mnuExclude_Click);
            //Создаём новый объект "Элемент полосы меню" с именем "Всё в расчёт"
            ToolStripMenuItem mnuInclude = new ToolStripMenuItem("Всё в расчёт");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Ничего в расчёт"
            mnuInclude.Click += new EventHandler(mnuInclude_Click);

            //-----------------Формирование пунктов "Отметки"-----------------

            //------------Добавление пунктов "Отметки" к меню-----------------

            //Добавляем в меню "Отметки" элемент "Всё в диспетчерскую"
            mnuChecks.DropDownItems.Add(mnuDispatch);
            //Добавляем в меню "Отметки" элемент "Всё в расчёт"
            mnuChecks.DropDownItems.Add(mnuInclude);
            //Добавляем в меню "Отметки" элемент "Ничего в диспетчерскую"
            mnuChecks.DropDownItems.Add(mnuUnDispatch);
            //Добавляем в меню "Отметки" элемент "Ничего в расчёт"
            mnuChecks.DropDownItems.Add(mnuExclude);

            //------------Добавление пунктов "Отметки" к меню-----------------

            //-----------------Формирование пунктов "Переход к"----------------

            //Создаём новый объект "Элемент полосы меню" с именем "Суммарному распределению"
            ToolStripMenuItem mnuSumDist = new ToolStripMenuItem("Суммарному распределению");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Суммарному распределению"
            mnuSumDist.Click += new EventHandler(mnuSumDist_Click);
            //Создаём новый объект "Элемент полосы меню" с именем "Анализу"
            ToolStripMenuItem mnuAnalysis = new ToolStripMenuItem("Анализу");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Анализу"
            mnuAnalysis.Click += new EventHandler(mnuAnalysis_Click);
            //Создаём новый объект "Элемент полосы меню" с именем "Конвертации"
            ToolStripMenuItem mnuConvert = new ToolStripMenuItem("Конвертации");
            //Приписываем на событие "Нажатие кнопкой мыши" для пункта меню "Конвертации"
            mnuConvert.Click += new EventHandler(mnuConvert_Click);

            //-----------------Формирование пунктов "Переход к"----------------

            //------------Добавление пунктов "Переход к" к меню----------------

            //Добавляем в меню "Переход к" элемент "Суммарному распределению"
            mnuNavigate.DropDownItems.Add(mnuSumDist);
            //Добавляем в меню "Переход к" элемент "Анализу"
            mnuNavigate.DropDownItems.Add(mnuAnalysis);
            //Добавляем в меню "Переход к" элемент "Конвертации"
            mnuNavigate.DropDownItems.Add(mnuConvert);

            //------------Добавление пунктов "Переход к" к меню----------------

            //------Добавление функциональных элементов меню в полосу----------

            //Добавляем меню "Основное" в полосу меню
            mnuDistribution.Items.Add(mnuMain);
            //Добавляем меню "Редактирование" в полосу меню
            mnuDistribution.Items.Add(mnuEdit);
            //Добавляем меню "Операции" в полосу меню
            mnuDistribution.Items.Add(mnuOper);
            //Добавляем меню "Отметки" в полосу меню
            mnuDistribution.Items.Add(mnuChecks);
            //Добавляем меню "Переход к" в полосу меню
            mnuDistribution.Items.Add(mnuNavigate);

            //------Добавление функциональных элементов меню в полосу----------

            //Размещаем полосу меню на главной форме
            mnuDistribution.Parent = this;
        }

        void mnuEditSubj_Click(object sender, EventArgs e)
        {
            toGenerateForm(new frmEditSubject());
            FillSubjectList(ref cmbSubjectList, true);
            FillSubjectList(ref cmbSubjectFilt, false);
        }

        void mnuClearHours_Click(object sender, EventArgs e)
        {
            //Вычищаем часы из элементов коллекции основной нагрузки
            for (int i = 0; i < mdlData.colDistribution.Count; i++)
            {
                mdlData.colDistribution[i].ClearHours();
            }

            //Вычищаем часы из элементов коллекции комбинированной нагрузки
            for (int i = 0; i < mdlData.colCombineDistribution.Count; i++)
            {
                mdlData.colCombineDistribution[i].ClearHours();
            }

            //Вычищаем часы из элементов коллекции почасовой нагрузки
            for (int i = 0; i < mdlData.colHouredDistribution.Count; i++)
            {
                mdlData.colHouredDistribution[i].ClearHours();
            }
        }

        void mnuConvert_Click(object sender, EventArgs e)
        {
            float sum = 0;

            mdlData.toConvertDistributionToDetailed(ref sum);

            MessageBox.Show("Получено строк: " + mdlData.colDistributionDetailed.Count);

            mdlData.toModifyDistributionToDetailed();

            MessageBox.Show("Оставлено строк: " + mdlData.colDistributionDetailed.Count);

            //MessageBox.Show("Получено: " + mdlData.colDistributionDetailed.Count + " строк\nСуммарная нагрузка: " + 
            //    sum.ToString("0.00"), "Учебная нагрузка сконвертирована");
        }

        void mnuCopy_Click(object sender, EventArgs e)
        {
            if (chkSubjectFilt.Checked || chkKursFilt.Checked || chkSpecialityFilt.Checked ||
                chkLecturerFilt.Checked || chkSemestrFilt.Checked || chkFacultyFilt.Checked ||
                chkTypeFilt.Checked)
            {
                toCopyFilt();
            }
            else
            {
                toCopy();
            }
        }

        void mnuCreate_Click(object sender, EventArgs e)
        {
            if (chkSubjectFilt.Checked || chkKursFilt.Checked || chkSpecialityFilt.Checked ||
                chkLecturerFilt.Checked || chkSemestrFilt.Checked || chkFacultyFilt.Checked ||
                chkTypeFilt.Checked)
            {
                MessageBox.Show(this, "В настоящий момент в режиме фильтрации создание нового элемента недостуно.",
                    "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                toCreateNew();
            }
        }

        void mnuAnalysis_Click(object sender, EventArgs e)
        {
            toDoAnalysis();
        }

        void mnuDel_Click(object sender, EventArgs e)
        {
            DelObject();
        }

        void mnuSave_Click(object sender, EventArgs e)
        {
            SaveChanges();
        }

        void mnuInclude_Click(object sender, EventArgs e)
        {
            MakeAllExcluded(false);
        }

        void mnuExclude_Click(object sender, EventArgs e)
        {
            MakeAllExcluded(true);
        }

        void mnuUnDispatch_Click(object sender, EventArgs e)
        {
            MakeAllDispatched(false);
        }

        void mnuDispatch_Click(object sender, EventArgs e)
        {
            MakeAllDispatched(true);
        }

        void mnuDelNull_Click(object sender, EventArgs e)
        {
            toDelNullRows();
        }

        //Объединение строк одного преподавателя по одной дисциплине
        void mnuMergeSame_Click(object sender, EventArgs e)
        {
            toMergeSame();
        }

        void mnuClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        void mnuSumDist_Click(object sender, EventArgs e)
        {
            toOpenSumDist();
        }

        /// <summary>
        /// Загрузка формы распределения учебной нагрузки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmEditDistribution_Load(object sender, EventArgs e)
        {
            //Вызов метода создания и настройки меню
            CreateMenu();
            
            txtFind.KeyUp += new KeyEventHandler(txtFind_KeyUp);
            
            lblThisAndOtherSum.Enabled = false;
            txtThisAndOtherSum.Enabled = false;

            lblAnotherSum.Enabled = false;
            txtAnotherSum.Enabled = false;

            lblSumHours.Enabled = false;
            txtSumHours.Enabled = false;

            lblSumSubject.Enabled = false;
            txtSumSubject.Enabled = false;

            lblRows.Enabled = false;
            txtRows.Enabled = false;

            lblCode.Enabled = false;
            txtCode.Enabled = false;

            lblCount.Enabled = false;
            txtCount.Enabled = false;
            
            FillSubjectList(ref cmbSubjectList, true);
            FillSubjectList(ref cmbSubjectFilt, false);
            FillLecturerList(cmbLecturerList, true);
            FillLecturerList(cmbLecturer2List, true);
            FillLecturerList(cmbLecturer3List, true);
            FillLecturerList(cmbDoubler, true);
            FillLecturerList(cmbLecturerFilt, false);
            FillKursList(cmbKursList, true);
            FillKursList(cmbKursFilt, false);
            FillSpecialityList(cmbSpecialityList, true);
            FillSpecialityList(cmbSpecialityFilt, false);
            FillSemestrList(cmbSemestrList, true);
            FillSemestrList(cmbSemestrFilt, false);
            FillFacultyList(cmbFacultyFilt, false);

            //Заполняем элементами распределения
            FillDistributionList(cmbLabWorkConnect, mdlData.colDistribution, mdlData.cmbRem,
                                 false, 
                                 false, 
                                 false,
                                 false, 
                                 false, 
                                 false, 
                                 false,
                                 false);
            
            optHoured.Checked = false;
            optCombine.Checked = false;
            optMain.Checked = true;

            cmbSubjectFilt.Enabled = mdlData.flgSubjectFilt;
            cmbSubjectFilt.SelectedIndex = mdlData.inxSubject;

            cmbKursFilt.Enabled = mdlData.flgKursFilt;
            cmbKursFilt.SelectedIndex = mdlData.inxKurs;

            cmbSpecialityFilt.Enabled = mdlData.flgSpecialityFilt;
            cmbSpecialityFilt.SelectedIndex = mdlData.inxSpeciality;

            cmbLecturerFilt.Enabled = mdlData.flgLecturerFilt;
            cmbLecturerFilt.SelectedIndex = mdlData.inxLecturer;

            cmbSemestrFilt.Enabled = mdlData.flgSemestrFilt;
            cmbSemestrFilt.SelectedIndex = mdlData.inxSemestr;

            cmbFacultyFilt.Enabled = mdlData.flgFacultyFilt;
            cmbFacultyFilt.SelectedIndex = mdlData.inxFaculty;

            cmbTypeFilt.Enabled = mdlData.flgTypeFilt;
            cmbTypeFilt.SelectedIndex = mdlData.inxType;

            chkFacultyFilt.Checked = mdlData.flgFacultyFilt;
            chkKursFilt.Checked = mdlData.flgKursFilt;
            chkLecturerFilt.Checked = mdlData.flgLecturerFilt;
            chkSemestrFilt.Checked = mdlData.flgSemestrFilt;
            chkSpecialityFilt.Checked = mdlData.flgSpecialityFilt;
            chkSubjectFilt.Checked = mdlData.flgSubjectFilt;
            chkTypeFilt.Checked = mdlData.flgTypeFilt;

            this.KeyPress += new KeyPressEventHandler(frmEditDistribution_KeyPress);
            this.VerticalScroll.Enabled = true;
        }

        void txtFind_KeyUp(object sender, KeyEventArgs e)
        {
            //Если нажата клавиша "Enter"
            if (e.KeyCode == Keys.Enter)
            {
                for (int i = 0; i <= cmbDistributionList.Items.Count - 1; i++)
                {
                    if (cmbDistributionList.Items[i].ToString().StartsWith(txtFind.Text))
                    {
                        cmbDistributionList.SelectedIndex = i;
                        break;
                    }
                    else
                    {

                    }
                }
            }
        }

        void frmEditDistribution_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void toOpenSumDist()
        {
            //Создаём новый объект класса "Форма суммарное распределение"
            frmSumDistribution SumDistr = new frmSumDistribution();
            //Делаем наследование от формы распределения нагрузки
            SumDistr.Owner = this;
            //Отображаем форму на экране
            SumDistr.ShowDialog();
        }

        private void cmbDistributionList_SelectedIndexChanged(object sender, EventArgs e)
        {
            toolTip.SetToolTip(cmbDistributionList, cmbDistributionList.Items[cmbDistributionList.SelectedIndex].ToString());
            ShowDistribution(cmbDistributionList.SelectedIndex);
        }

        private void optHoured_CheckedChanged(object sender, EventArgs e)
        {
            if (optHoured.Checked)
            {
                if (cmbDistributionList.SelectedIndex >= 0)
                {
                    curNum = cmbDistributionList.SelectedIndex;
                }
                else
                {
                    curNum = 0;
                }

                Selected = mdlData.colHouredDistribution;

                FillDistributionList(cmbDistributionList, Selected, mdlData.cmbRem,
                                     chkSubjectFilt.Checked, 
                                     chkKursFilt.Checked, 
                                     chkSpecialityFilt.Checked, 
                                     chkLecturerFilt.Checked, 
                                     chkSemestrFilt.Checked, 
                                     chkFacultyFilt.Checked, 
                                     chkTypeFilt.Checked,
                                     false);
                
                if (curNum <= cmbDistributionList.Items.Count - 1)
                {
                    cmbDistributionList.SelectedIndex = curNum;
                }
            }
        }

        private void FillDistributionList(ComboBox cmb,
            IList<clsDistribution> collection,
            string selItem,
            bool SubjFilt, 
            bool KursFilt,
            bool SpecFilt, 
            bool LectFilt, 
            bool SemFilt, 
            bool FacFilt,
            bool TypeFilt,
            bool flgIgnoreFilt)
        {
            string Spec;
            string Kurs;
            string Subj;
            string Kod;

            bool NeedFilt = false;
            bool AddToColl = false;

            int i, j;

            //Очистка комбо-списка распределения нагрузки
            cmb.Items.Clear();

            //Если коллекция отфильтрованной нагрузки не пуста
            if (!(mdlData.Filtred == null))
            {
                //то очищаем эту коллекцию
                mdlData.Filtred.Clear();
            }

            //Если пришедшая в процедуру основная коллекция не пуста,
            //то начинаем работу по набору фильтрованных элементов
            if (!(collection == null))
            {
                if (cmb.Equals(cmbLabWorkConnect))
                {
                    cmb.Items.Add("не выбрано");
                }

                //Заполняем комбо-список распределения нагрузки
                //Сбрасываем сведения о максимальном коде элемента в коллекции
                MaxNum = 0;
                for (i = 0; i <= collection.Count - 1; i++)
                {
                    //По умолчанию считаем, что элемент нужно добавить в коллекцию
                    AddToColl = true;

                    //Если максимальный код оказался меньше кода поступившего элемента,
                    //значит найденный ранее код не максимален - заменяем его
                    if (MaxNum < collection[i].Code)
                    {
                        MaxNum = collection[i].Code;
                    }

                    if (mdlData.flgOldDB)
                    {
                        if (!(collection[i].Speciality == null))
                        {
                            Spec = collection[i].Speciality.ShortUpravlenie;
                        }
                        else
                        {
                            Spec = "";
                        }
                    }
                    else
                    {
                        if (!(collection[i].Speciality == null))
                        {
                            Spec = collection[i].Speciality.ShortDop;
                        }
                        else
                        {
                            Spec = "";
                        }
                    }

                    if (!(collection[i].KursNum == null))
                    {
                        Kurs = collection[i].KursNum.Kurs.ToString();
                    }
                    else
                    {
                        Kurs = "";
                    }

                    if (!(collection[i].Subject == null))
                    {
                        Subj = collection[i].Subject.Subject;
                    }
                    else
                    {
                        Subj = "";
                    }

                    Kod = collection[i].Code.ToString();

                    if (collection[i].Code.ToString().Length < 4)
                    {
                        for (j = 0; j < 4 - collection[i].Code.ToString().Length; j++)
                        {
                            Kod = "0" + Kod;
                        }
                    }

                    //Определяем, нужно ли фильтровать согласно флагам
                    if (SubjFilt || KursFilt ||
                        SpecFilt || LectFilt || 
                        SemFilt || FacFilt || 
                        TypeFilt)
                    {
                        NeedFilt = true;
                    }

                    //Если сложилась ситуация, что фильтровать нужно
                    if (NeedFilt)
                    {
                        //Надо проверить, не указано ли принудительно обратное
                        if (flgIgnoreFilt)
                        {
                            //Если пришли в заполнитель в режиме фильтрации, но для
                            //не основного комбинированного списка
                            NeedFilt = false;
                        }
                    }

                    //Если фильтровать не нужно, то добавляем всё
                    if (!(NeedFilt))
                    {   
                        if (collection[i].Lecturer != null)
                        {
                            cmb.Items.Add(Kod + ". " + Kurs + " - " +
                                          Spec + " - " + Subj + " - [" + 
                                          mdlData.SplitFIOString(collection[i].Lecturer.FIO, true, true) + "]");
                        }
                        else
                        {
                            cmb.Items.Add(Kod + ". " + Kurs + " - " +
                                          Spec + " - " + Subj);
                        }
                    }
                    else
                    {
                        //Если требуется фильтровать по дисциплинам
                        if (SubjFilt)
                        {
                            //Если дисциплина указана надо ещё принять решение
                            if (!(collection[i].Subject == null))
                            {
                                if (cmbSubjectFilt.SelectedIndex > -1)
                                {
                                    //Если дисциплины совпали, то нас интересует такой элемент
                                    if (collection[i].Subject.Subject.Equals(mdlData.colSubject[cmbSubjectFilt.SelectedIndex].Subject))
                                    {
                                        AddToColl = (true & AddToColl);
                                    }
                                    //Иначе не интересен
                                    else
                                    {
                                        AddToColl = (false & AddToColl);
                                    }
                                }
                                else
                                {
                                    AddToColl = (false & AddToColl);
                                }
                            }
                            //Иначе не интересен
                            else
                            {
                                AddToColl = (false & AddToColl);
                            }
                        }

                        //Если требуется фильтровать по курсам
                        if (KursFilt)
                        {
                            //Если курс указан надо ещё принять решение
                            if (!(collection[i].KursNum == null))
                            {
                                if (cmbKursFilt.SelectedIndex > -1)
                                {
                                    //Если курсы совпали, то нас интересует такой элемент
                                    if (collection[i].KursNum.Kurs.Equals(mdlData.colKursNum[cmbKursFilt.SelectedIndex].Kurs))
                                    {
                                        AddToColl = (true & AddToColl);
                                    }
                                    //Иначе не интересует
                                    else
                                    {
                                        AddToColl = (false & AddToColl);
                                    }
                                }
                                else
                                {
                                    AddToColl = (false & AddToColl);
                                }
                            }
                            //Иначе не интересен
                            else
                            {
                                AddToColl = (false & AddToColl);
                            }
                        }

                        //Если требуется фильтровать по специальности
                        if (SpecFilt)
                        {
                            if (cmbSpecialityFilt.SelectedIndex >= 0)
                            {
                                //Если специальность указана надо ещё принять решение
                                if (!(collection[i].Speciality == null))
                                {
                                    if (mdlData.flgOldDB)
                                    {
                                        if (cmbSpecialityFilt.SelectedIndex > -1)
                                        {
                                            //Если специальности совпали, то нас интересует такой элемент
                                            if (collection[i].Speciality.ShortUpravlenie.Equals(mdlData.colSpecialisation[cmbSpecialityFilt.SelectedIndex].ShortUpravlenie))
                                            {
                                                AddToColl = (true & AddToColl);
                                            }
                                            //Иначе не интересует
                                            else
                                            {
                                                AddToColl = (false & AddToColl);
                                            }
                                        }
                                        else
                                        {
                                            AddToColl = (false & AddToColl);
                                        }
                                    }
                                    else
                                    {
                                        if (cmbSpecialityFilt.SelectedIndex > -1)
                                        {
                                            //Если специальности совпали, то нас интересует такой элемент
                                            if (collection[i].Speciality.ShortDop.Equals(mdlData.colSpecialisation[cmbSpecialityFilt.SelectedIndex].ShortDop))
                                            {
                                                AddToColl = (true & AddToColl);
                                            }
                                            //Иначе не интересует
                                            else
                                            {
                                                AddToColl = (false & AddToColl);
                                            }
                                        }
                                        else
                                        {
                                            AddToColl = (false & AddToColl);
                                        }
                                    }
                                }
                                else
                                {
                                    AddToColl = (false & AddToColl);
                                }
                            }
                            //Иначе не интересен
                            else
                            {
                                AddToColl = (false & AddToColl);
                            }
                        }

                        //Если требуется фильтровать по типу специальности
                        if (TypeFilt)
                        {
                            //Если специальность указана надо ещё принять решение
                            if (!(collection[i].Speciality == null))
                            {
                                //Если специальности по типу совпали, то нас интересует такой элемент
                                if ( collection[i].Speciality.Diff.Equals(selItem) )
                                {
                                    AddToColl = (true & AddToColl);
                                }
                                //Иначе не интересует
                                else
                                {
                                    AddToColl = (false & AddToColl);
                                }
                            }
                            //Иначе не интересен
                            else
                            {
                                AddToColl = (false & AddToColl);
                            }
                        }

                        //Если требуется фильтровать по преподавателю
                        if (LectFilt)
                        {
                            //Если преподаватель указан надо ещё принять решение
                            if (!(collection[i].Lecturer == null))
                            {
                                if (cmbLecturerFilt.SelectedIndex > -1)
                                {
                                    //Если преподаватели совпали, то нас интересует такой элемент
                                    if (collection[i].Lecturer.FIO.Equals(mdlData.colLecturer[cmbLecturerFilt.SelectedIndex].FIO))
                                    {
                                        AddToColl = (true & AddToColl);
                                    }
                                    //Иначе не интересует
                                    else
                                    {
                                        AddToColl = (false & AddToColl);
                                    }
                                }
                                else
                                {
                                    AddToColl = (false & AddToColl);
                                }
                            }
                            //Иначе не интересен
                            else
                            {
                                AddToColl = (false & AddToColl);
                            }
                        }

                        //Если требуется фильтровать по семестрам
                        if (SemFilt)
                        {
                            //Если семестр указан надо ещё принять решение
                            if (!(collection[i].Semestr == null))
                            {
                                if (cmbSemestrFilt.SelectedIndex > -1)
                                {
                                    //Если семестры совпали, то нас интересует такой элемент
                                    if (collection[i].Semestr.SemNum.Equals(mdlData.colSemestr[cmbSemestrFilt.SelectedIndex].SemNum))
                                    {
                                        AddToColl = (true & AddToColl);
                                    }
                                    //в ином случае элемент нас не интересует
                                    else
                                    {
                                        AddToColl = (false & AddToColl);
                                    }
                                }
                                else
                                {
                                    AddToColl = (false & AddToColl);
                                }
                            }
                            //Если семестр не указан, то нас такой элемент не интересует
                            else
                            {
                                AddToColl = (false & AddToColl);
                            }
                        }

                        //Если требуется фильтровать по факульететам
                        if (FacFilt)
                        {
                            //Если специальность указана надо ещё принять решение
                            if (!(collection[i].Speciality == null))
                            {
                                //Если факультет указан надо ещё принять решение
                                if (!(collection[i].Speciality.Faculty == null))
                                {
                                    if (cmbFacultyFilt.SelectedIndex > -1)
                                    {
                                        //Если факультеты совпали, то нас интересует такой элемент
                                        if (collection[i].Speciality.Faculty.Faculty.Equals(mdlData.colFaculty[cmbFacultyFilt.SelectedIndex].Faculty))
                                        {
                                            AddToColl = (true & AddToColl);
                                        }
                                        //в ином случае элемент нас не интересует
                                        else
                                        {
                                            AddToColl = (false & AddToColl);
                                        }
                                    }
                                    else
                                    {
                                        AddToColl = (false & AddToColl);
                                    }
                                }
                                //в ином случае элемент нас не интересует
                                else
                                {
                                    AddToColl = (false & AddToColl);
                                }
                            }
                            //Если семестр не указан, то нас такой элемент не интересует
                            else
                            {
                                AddToColl = (false & AddToColl);
                            }
                        }

                        //Если мы прошли все ступени и остались true,
                        //значит, мы соответствуем всем требованиям
                        //и должны добавить элемент в список
                        if (AddToColl)
                        {
                            if (collection[i].Lecturer != null)
                            {
                                cmb.Items.Add(Kod + ". " + Kurs + " - " +
                                              Spec + " - " + Subj + " - [" +
                                              mdlData.SplitFIOString(collection[i].Lecturer.FIO, true, true) + "]");
                            }
                            else
                            {
                                cmb.Items.Add(Kod + ". " + Kurs + " - " +
                                              Spec + " - " + Subj);
                            }

                            mdlData.Filtred.Add(collection[i]);
                        }
                    }
                }

                if (cmb.Items.Count > 0)
                {
                    cmb.SelectedIndex = 0;
                }
                else
                {
                    cmb.SelectedIndex = -1;
                }

                if (cmb.Equals(cmbDistributionList))
                {
                    txtCount.Text = (cmb.Items.Count).ToString();
                }
            }
        }

        private void optMain_CheckedChanged(object sender, EventArgs e)
        {
            if (optMain.Checked)
            {
                if (cmbDistributionList.SelectedIndex >= 0)
                {
                    curNum = cmbDistributionList.SelectedIndex;
                }
                else
                {
                    curNum = 0;
                }

                Selected = mdlData.colDistribution;

                FillDistributionList(cmbDistributionList, Selected, mdlData.cmbRem,
                                     chkSubjectFilt.Checked, 
                                     chkKursFilt.Checked, 
                                     chkSpecialityFilt.Checked,
                                     chkLecturerFilt.Checked, 
                                     chkSemestrFilt.Checked, 
                                     chkFacultyFilt.Checked, 
                                     chkTypeFilt.Checked,
                                     false);
                
                if (curNum <= cmbDistributionList.Items.Count - 1)
                {
                    cmbDistributionList.SelectedIndex = curNum;
                }
            }
        }

        private void optCombine_CheckedChanged(object sender, EventArgs e)
        {
            if (optCombine.Checked)
            {
                if (cmbDistributionList.SelectedIndex >= 0)
                {
                    curNum = cmbDistributionList.SelectedIndex;
                }
                else
                {
                    curNum = 0;
                }

                Selected = mdlData.colCombineDistribution;

                FillDistributionList(cmbDistributionList, Selected, mdlData.cmbRem,
                                     chkSubjectFilt.Checked, 
                                     chkKursFilt.Checked, 
                                     chkSpecialityFilt.Checked,
                                     chkLecturerFilt.Checked, 
                                     chkSemestrFilt.Checked, 
                                     chkFacultyFilt.Checked,
                                     chkTypeFilt.Checked,
                                     false);

                if (curNum <= cmbDistributionList.Items.Count - 1)
                {
                    cmbDistributionList.SelectedIndex = curNum;
                }
            }
        }

        private void FillLecturerList(ComboBox cmb, bool flgReset)
        {
            int NumFix = 0;
            NumFix = cmb.SelectedIndex;
            
            //Очищаем список
            cmb.Items.Clear();

            if (!cmb.Equals(cmbLecturerFilt))
            {
                cmb.Items.Add("не выбран");
            }

            //Заполняем комбо-список преподавателями
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                cmb.Items.Add(mdlData.colLecturer[i].Code + ". " + mdlData.colLecturer[i].FIO);
            }
            
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (flgReset)
            {
                if (NumFix < 0)
                {
                    cmb.SelectedIndex = 0;
                }
                else
                {
                    cmb.SelectedIndex = NumFix;
                }
            }
        }

        private void FillSemestrList(ComboBox cmb, bool flgReset)
        {
            int NumFix = 0;
            NumFix = cmb.SelectedIndex;
            //Очищаем список
            cmb.Items.Clear();

            //Заполняем комбо-список должностями
            for (int i = 0; i <= mdlData.colSemestr.Count - 1; i++)
            {
                cmb.Items.Add(mdlData.colSemestr[i].Code + ". " + mdlData.colSemestr[i].SemNum);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (flgReset)
            {
                if (NumFix < 0)
                {
                    cmb.SelectedIndex = 0;
                }
                else
                {
                    cmb.SelectedIndex = NumFix;
                }
            }
        }

        private void FillKursList(ComboBox cmb, bool flgReset)
        {
            int NumFix = 0;
            NumFix = cmb.SelectedIndex;
            //Очищаем список
            cmb.Items.Clear();

            //Заполняем комбо-список должностями
            for (int i = 0; i <= mdlData.colKursNum.Count - 1; i++)
            {
                cmb.Items.Add(mdlData.colKursNum[i].Kurs);
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (flgReset)
            {
                if (NumFix < 0)
                {
                    cmb.SelectedIndex = 0;
                }
                else
                {
                    cmb.SelectedIndex = NumFix;
                }
            }
        }

        private void FillSpecialityList(ComboBox cmb, bool flgReset)
        {
            int NumFix = 0;
            NumFix = cmb.SelectedIndex;
            //Очищаем список
            cmb.Items.Clear();

            //Заполняем комбо-список должностями
            for (int i = 0; i <= mdlData.colSpecialisation.Count - 1; i++)
            {
                cmb.Items.Add(mdlData.colSpecialisation[i].Code + ". " + mdlData.colSpecialisation[i].ShortUpravlenie +
                    " [" + mdlData.colSpecialisation[i].ShortDop + "]" +
                    " (" + mdlData.colSpecialisation[i].ShortInstitute + ")");
            }
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (flgReset)
            {
                if (NumFix < 0)
                {
                    cmb.SelectedIndex = 0;
                }
                else
                {
                    cmb.SelectedIndex = NumFix;
                }
            }
        }

        private void FillSubjectList(ref ComboBox cmb, bool flgReset)
        {
            int NumFix = 0;
            NumFix = cmb.SelectedIndex;
            //Очищаем список
            cmb.Items.Clear();

            //Заполняем комбо-список должностями
            for (int i = 0; i <= mdlData.colSubject.Count - 1; i++)
            {
                cmb.Items.Add(mdlData.colSubject[i].Code + ". " + mdlData.colSubject[i].Subject);
            }
 
            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (flgReset)
            {
                if (NumFix < 0)
                {
                    cmb.SelectedIndex = 0;
                }
                else
                {
                    cmb.SelectedIndex = NumFix;
                }
            }
        }

        private void FillFacultyList(ComboBox cmb, bool flgReset)
        {
            int NumFix = 0;
            NumFix = cmb.SelectedIndex;
            //Очищаем список
            cmb.Items.Clear();

            //Заполняем комбо-список должностями
            for (int i = 0; i <= mdlData.colFaculty.Count - 1; i++)
            {
                cmb.Items.Add(mdlData.colFaculty[i].Code + ". " + mdlData.colFaculty[i].Short);
            }

            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (flgReset)
            {
                if (NumFix < 0)
                {
                    cmb.SelectedIndex = 0;
                }
                else
                {
                    cmb.SelectedIndex = NumFix;
                }
            }
        }

        private void chkSubjectFilt_CheckedChanged(object sender, EventArgs e)
        {
            CheckFiltParam(chkSubjectFilt, cmbSubjectFilt, mdlData.inxSubject, ref mdlData.flgSubjectFilt);
        }

        private void chkKursFilt_CheckedChanged(object sender, EventArgs e)
        {
            CheckFiltParam(chkKursFilt, cmbKursFilt, mdlData.inxKurs, ref mdlData.flgKursFilt);
        }

        private void chkSpecialityFilt_CheckedChanged(object sender, EventArgs e)
        {
            CheckFiltParam(chkSpecialityFilt, cmbSpecialityFilt, mdlData.inxSpeciality, ref mdlData.flgSpecialityFilt);
        }

        private void chkLecturerFilt_CheckedChanged(object sender, EventArgs e)
        {
            CheckFiltParam(chkLecturerFilt, cmbLecturerFilt, mdlData.inxLecturer, ref mdlData.flgLecturerFilt);
        }

        private void chkSemestrFilt_CheckedChanged(object sender, EventArgs e)
        {
            CheckFiltParam(chkSemestrFilt, cmbSemestrFilt, mdlData.inxSemestr, ref mdlData.flgSemestrFilt);
        }

        private void ShowDistribution(int ind)
        {
            //Объявление локальных переменных под счётчики
            //(счётные переменные)
            int SumRow;
            int SumSubj;
            int RowCount;
            int AnotherSum;

            IList<clsDistribution> coll = null;

            if (optMain.Checked)
            {
                coll = mdlData.colDistribution;
            }
            if (optHoured.Checked)
            {
                coll = mdlData.colHouredDistribution;
            }
            if (optCombine.Checked)
            {
                coll = mdlData.colCombineDistribution;
            }

            if (chkSubjectFilt.Checked || chkKursFilt.Checked || chkSpecialityFilt.Checked
                || chkLecturerFilt.Checked || chkSemestrFilt.Checked || chkFacultyFilt.Checked 
                || chkTypeFilt.Checked)
            {
                coll = mdlData.Filtred;
            }

            if (!(coll == null))
            {
                if (coll.Count > 0)
                {
                    //Назначаем дисциплину
                    if (!(coll[ind].Subject == null))
                    {
                        cmbSubjectList.SelectedIndex = mdlData.colSubject.IndexOf(coll[ind].Subject);
                    }
                    else
                    {
                        cmbSubjectList.SelectedIndex = -1;
                    }

                    //Назначаем курс
                    if (!(coll[ind].KursNum == null))
                    {
                        cmbKursList.SelectedIndex = mdlData.colKursNum.IndexOf(coll[ind].KursNum);
                    }
                    else
                    {
                        cmbKursList.SelectedIndex = -1;
                    }

                    //Назначаем специальность
                    if (!(coll[ind].Speciality == null))
                    {
                        cmbSpecialityList.SelectedIndex = mdlData.colSpecialisation.IndexOf(coll[ind].Speciality);
                    }
                    else
                    {
                        cmbSpecialityList.SelectedIndex = -1;
                    }

                    //Назначаем преподавателя
                    if (!(coll[ind].Lecturer == null))
                    {
                        cmbLecturerList.SelectedIndex = mdlData.colLecturer.IndexOf(coll[ind].Lecturer) + 1;
                    }
                    else
                    {
                        cmbLecturerList.SelectedIndex = 0;
                    }

                    //Назначаем замещающего преподавателя
                    if (!(coll[ind].Lecturer2 == null))
                    {
                        cmbLecturer2List.SelectedIndex = mdlData.colLecturer.IndexOf(coll[ind].Lecturer2) + 1;
                    }
                    else
                    {
                        cmbLecturer2List.SelectedIndex = 0;
                    }

                    //Назначаем дополнительного преподавателя
                    if (!(coll[ind].Lecturer3 == null))
                    {
                        cmbLecturer3List.SelectedIndex = mdlData.colLecturer.IndexOf(coll[ind].Lecturer3) + 1;
                    }
                    else
                    {
                        cmbLecturer3List.SelectedIndex = 0;
                    }

                    //Назначаем дублёра
                    if (!(coll[ind].Doubler == null))
                    {
                        cmbDoubler.SelectedIndex = mdlData.colLecturer.IndexOf(coll[ind].Doubler) + 1;
                    }
                    else
                    {
                        cmbDoubler.SelectedIndex = 0;
                    }

                    //Назначаем семестр
                    if (!(coll[ind].Semestr == null))
                    {
                        cmbSemestrList.SelectedIndex = mdlData.colSemestr.IndexOf(coll[ind].Semestr);
                    }
                    else
                    {
                        cmbSemestrList.SelectedIndex = -1;
                    }

                    //Назначаем связку по лабораторным работам
                    if (!(coll[ind].LabWorkConnect == null))
                    {
                        try
                        {
                            //Добавляется +1, поскольку в этом списке сверху добавлена строка
                            //"не выбрано" по сравнению с основным списком
                            cmbLabWorkConnect.SelectedIndex = Selected.IndexOf(coll[ind].LabWorkConnect) + 1;
                        }
                        catch
                        {
                            cmbLabWorkConnect.SelectedIndex = 0;
                        }
                    }
                    else
                    {
                        cmbLabWorkConnect.SelectedIndex = 0;
                    }

                    //Выводим штатную нагрузку

                    //Выводим код элемента
                    txtCode.Text = coll[ind].Code.ToString();
                    //Выводим лекции в часах
                    txtLecture.Text = coll[ind].Lecture.ToString();
                    //Выводим экзамены в часах
                    txtExam.Text = coll[ind].Exam.ToString();
                    //Выводим зачёты в часах
                    txtCred.Text = coll[ind].Credit.ToString();
                    //Выводим реферат в часах
                    txtRef.Text = coll[ind].RefHomeWork.ToString();
                    //Выводим консультации в часах
                    txtTut.Text = coll[ind].Tutorial.ToString();
                    //Выводим лабораторные работы в часах
                    txtLab.Text = coll[ind].LabWork.ToString();
                    //Выводим практические занятия в часах
                    txtPract.Text = coll[ind].Practice.ToString();
                    //Выводим индивидуальные занятия в часах
                    txtInd.Text = coll[ind].IndividualWork.ToString();
                    //Выводим КРАПК в часах
                    txtKRAPK.Text = coll[ind].KRAPK.ToString();
                    //Выводим курсовой проект в часах
                    txtKursPr.Text = coll[ind].KursProject.ToString();
                    //Выводим преддипломная практика в часах
                    txtPreD.Text = coll[ind].PreDiplomaPractice.ToString();
                    //Выводим диплом в часах
                    txtDiploma.Text = coll[ind].DiplomaPaper.ToString();
                    //Выводим учебную практику в часах
                    txtTutPr.Text = coll[ind].TutorialPractice.ToString();
                    //Выводим производственную практику в часах
                    txtProd.Text = coll[ind].ProducingPractice.ToString();
                    //Выводим ГЭК в часах
                    txtGAK.Text = coll[ind].GAK.ToString();
                    //Выводим госбюджетные часы
                    txtHours.Text = coll[ind].Hours.ToString();
                    //Выводим госбюджетные часы
                    txtHoursZ.Text = coll[ind].HoursZ.ToString();
                    //Выводим часы на аспирантуру
                    txtPostGrad.Text = coll[ind].PostGrad.ToString();
                    //Выводим часы на посещение занятий
                    txtVisiting.Text = coll[ind].Visiting.ToString();
                    //Выводим заданное количество часов
                    txtEnteredHours.Text = coll[ind].EnteredHours.ToString();
                    //Выводим заданное количество часов по ЗЕТ
                    txtEnteredHoursZ.Text = coll[ind].EnteredHoursZ.ToString();
                    //Выводим примечание для диспетчерской
                    txtText.Text = coll[ind].Text;
                    //Выводим часы на руководство магистерской программой
                    txtMagistry.Text = coll[ind].Magistry.ToString();
                    //Выставляем признак необходимости включения в заявку для диспетчерской
                    chkDispatch.Checked = coll[ind].flgDispatch;
                    //Выставляем признак равномерного распределения нагрузки
                    chkDistrib.Checked = coll[ind].flgDistrib;
                    //Выводим вес равномерно распределяемой нагрузки
                    txtWeight.Text = coll[ind].Weight.ToString();
                    //Выставляем признак исключения из нагрузки
                    chkExclude.Checked = coll[ind].flgExclude;
                    //Выводим код дисциплины по документу
                    txtDocCode.Text = coll[ind].DocCode.ToString();

                    //Очищение счётных переменных
                    SumSubj = 0;
                    RowCount = 0;
                    AnotherSum = 0;
                    //Перебираем элементы нагрузки
                    for (int i = 0; i <= coll.Count - 1; i++)
                    {
                        if ((coll[i].Subject == coll[ind].Subject) &
                            (coll[i].Semestr == coll[ind].Semestr) &
                            (coll[i].KursNum == coll[ind].KursNum) &
                            (coll[i].Speciality == coll[ind].Speciality))
                        {
                            //Если обычный элемент нагрузки
                            if (!coll[i].flgDistrib)
                            {
                                SumSubj += coll[i].EnteredHours;
                                RowCount += 1;

                                if (!(i == ind))
                                {
                                    AnotherSum += mdlData.toSumDistributionComponents(coll[i]);
                                }
                            }
                            //Если равномерно распределяемый элемент нагрузки
                            else
                            {
                                for (int k = 0; k <= mdlData.colStudents.Count - 1; k++)
                                {
                                    if (coll[i].Semestr.Equals(mdlData.colSemestr[1]))
                                    {
                                        if (coll[i].KursNum.Equals(mdlData.colStudents[k].KursNum) &
                                             coll[i].Speciality.Equals(mdlData.colStudents[k].Speciality))
                                        {
                                            RowCount += 1;
                                            SumSubj += coll[i].Weight;
                                        }
                                    }

                                    if (coll[i].Semestr.Equals(mdlData.colSemestr[2]))
                                    {
                                        if (coll[i].KursNum.Equals(mdlData.colStudents[k].KursNum) &
                                             coll[i].Speciality.Equals(mdlData.colStudents[k].Speciality))
                                        {
                                            RowCount += 1;
                                            SumSubj += coll[i].Weight;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    //Вычисление суммарной нагрузки по выбранной строке нагрузки
                    SumRow = mdlData.toSumDistributionComponents(coll[ind]);
                    //Вывод в текстовое поле "Итоговая сумма"
                    txtSumHours.Text = SumRow.ToString();
                    //Вывод в текстовое поле "Сумма по дисциплине"
                    txtSumSubject.Text = SumSubj.ToString();
                    //Вывод в текстовое поле "Количество строк"
                    txtRows.Text = RowCount.ToString();
                    //Вывод в текстовое поле "Сумма других"
                    txtAnotherSum.Text = AnotherSum.ToString();
                    //Вывод в текстовое поле "Сумма"
                    txtThisAndOtherSum.Text = (SumRow + AnotherSum).ToString();

                    if (coll[ind].HouredConnect != null)
                    {
                        btnPutIntoHoured.Enabled = false;
                        btnDelFromHoured.Enabled = true;
                    }
                    else
                    {
                        btnPutIntoHoured.Enabled = true;
                        btnDelFromHoured.Enabled = false;
                    }


                    //MessageBox.Show(txtLecture.BackColor.Name.ToString() + "   " + txtLecture.ForeColor.Name.ToString());
                    //MessageBox.Show(txtSumHours.BackColor.Name.ToString() + "   " + txtSumHours.ForeColor.Name.ToString());

                    /*
                    //Если количество строк более одной
                    if (Convert.ToInt32(txtRows.Text) > 1)
                    {
                        //Если "Часы_Дано" не нулевые, то применяем цветную индикацию
                        if (!txtEnteredHours.Text.Equals("0"))
                        {
                            if (txtEnteredHours.Text.Equals(txtThisAndOtherSum.Text))
                            {
                                txtEnteredHours.BackColor = Color.Green;
                                txtEnteredHours.ForeColor = Color.White;
                                txtSumHours.BackColor = Color.LightGreen;
                                txtSumHours.ForeColor = txtLecture.ForeColor;
                            }
                            else
                            {
                                txtEnteredHours.BackColor = Color.Red;
                                txtEnteredHours.ForeColor = Color.Black;
                                txtSumHours.BackColor = Color.Salmon;
                                txtSumHours.BackColor = txtLecture.ForeColor;
                            }
                        }
                        //Если "Часы_Дано" нулевые, то применяем стандартную серую индикацию
                        else
                        {
                            txtEnteredHours.BackColor = ;
                            txtEnteredHours.ForeColor = ;
                            txtSumHours.BackColor = ;
                            txtSumHours.ForeColor = ;
                        }
                    }
                    //Если строка одна
                    else
                    {
                        //Если "Часы_Дано" не нулевые, то применяем цветную индикацию
                        if (!txtEnteredHours.Text.Equals("0"))
                        {
                            if (txtEnteredHours.Text.Equals(txtThisAndOtherSum.Text))
                            {
                                txtEnteredHours.BackColor = Color.Green;
                                txtEnteredHours.ForeColor = Color.White;
                                txtSumHours.BackColor = Color.LightGreen;
                                txtSumHours.ForeColor = txtLecture.ForeColor;
                            }
                            else
                            {
                                txtEnteredHours.BackColor = Color.Red;
                                txtEnteredHours.ForeColor = Color.Black;
                                txtSumHours.BackColor = Color.Salmon;
                                txtSumHours.ForeColor = txtLecture.ForeColor;
                            }
                        }
                        //Если "Часы_Дано" нулевые, то применяем стандартную серую индикацию
                        else
                        {

                        }
                    }
                    */ 
                }
            }
        }

        private void cmbSubjectFilt_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeFiltParam(cmbSubjectFilt, ref mdlData.inxSubject);
        }

        private void ChangeFiltParam(ComboBox cmbCur, ref int inx)
        {
            if (cmbCur.SelectedIndex >= 0)
            {
                toolTip.SetToolTip(cmbCur,
                    cmbCur.Items[cmbCur.SelectedIndex].ToString());

                if (cmbCur == cmbTypeFilt)
                {
                    mdlData.cmbRem = cmbCur.SelectedItem.ToString();
                }
            }
            else
            {
                if (cmbCur == cmbTypeFilt)
                {
                    mdlData.cmbRem = "";
                }
            }

            inx = cmbCur.SelectedIndex;

            //Список увязок по лабораторным работам заполняем по основной коллекции нагрузки,
            //иначе будут проблемы с обнаружением связанных элементов, когда фильруются
            //преподаватели (увязки, как правило, по разным преподавателям)
            FillDistributionList(cmbLabWorkConnect, Selected, mdlData.cmbRem,
                                 chkSubjectFilt.Checked,
                                 chkKursFilt.Checked,
                                 chkSpecialityFilt.Checked,
                                 chkLecturerFilt.Checked,
                                 chkSemestrFilt.Checked,
                                 chkFacultyFilt.Checked,
                                 chkTypeFilt.Checked,
                                 true);

            //Основной список заполняем только по отфильтрованной коллекции "Selected"
            FillDistributionList(cmbDistributionList, Selected, mdlData.cmbRem,
                                 chkSubjectFilt.Checked,
                                 chkKursFilt.Checked,
                                 chkSpecialityFilt.Checked,
                                 chkLecturerFilt.Checked,
                                 chkSemestrFilt.Checked,
                                 chkFacultyFilt.Checked,
                                 chkTypeFilt.Checked,
                                 false);
        }

        private void cmbKursFilt_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeFiltParam(cmbKursFilt, ref mdlData.inxKurs);
        }

        private void cmbSpecialityFilt_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeFiltParam(cmbSpecialityFilt, ref mdlData.inxSpeciality);
        }

        private void cmbLecturerFilt_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeFiltParam(cmbLecturerFilt, ref mdlData.inxLecturer);
        }

        private void cmbSemestrFilt_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeFiltParam(cmbSemestrFilt, ref mdlData.inxSemestr);
        }

        /// <summary>
        /// Метод объединения строк одного преподавателя по одной дисциплине
        /// </summary>
        private void toMergeSame()
        {
            //По умолчанию признак каких-либо изменений сброшен
            bool flg = false;
            bool flgBreak = false;
            string str;

            //Запускаем первый цикл по нагрузке
            for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
            {
                //Если случилось прерывание - завершаем цикл
                if (flgBreak)
                {
                    break;
                }
                //Запускаем второй цикл по той же нагрузке
                for (int j = 0; j <= mdlData.colDistribution.Count - 1; j++)
                {
                    //Сравнивать имеет смысл только элементы с разными индексами
                    //Чтобы не случилось объединения самого себя с самим собой
                    if (!(i == j))
                    {
                        //Если строки семестра, номера курса, специальности, преподавателя и предмета
                        //не пустые, то, потенциально они могут быть одинаковыми, что требуется проверить
                        //На этом этапе отметаем объединение строк, для которых не задан преподаватель
                        //(равномерно распределяемая нагрузка)
                        //И отметаем Аспирантуру, для которой не указывается номер курса
                        if ((mdlData.colDistribution[i].Semestr != null) & (mdlData.colDistribution[j].Semestr != null) &
                            (mdlData.colDistribution[i].KursNum != null) & (mdlData.colDistribution[j].KursNum != null) &
                            (mdlData.colDistribution[i].Speciality != null) & (mdlData.colDistribution[j].Speciality != null) &
                            (mdlData.colDistribution[i].Lecturer != null) & (mdlData.colDistribution[j].Lecturer != null) &
                            (mdlData.colDistribution[i].Subject != null) & (mdlData.colDistribution[j].Subject != null))
                        {
                            // Если строки семестра, номера курса, специальности, преподавателя и предмета
                            // для различных строк сходятся, значит, их можно объединить
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals(mdlData.colDistribution[j].Semestr.SemNum) &
                                mdlData.colDistribution[i].KursNum.Kurs.Equals(mdlData.colDistribution[j].KursNum.Kurs) &
                                mdlData.colDistribution[i].Speciality.ShortUpravlenie.Equals(mdlData.colDistribution[j].Speciality.ShortUpravlenie) &
                                mdlData.colDistribution[i].Lecturer.FIO.Equals(mdlData.colDistribution[j].Lecturer.FIO) &
                                mdlData.colDistribution[i].Subject.Subject.Equals(mdlData.colDistribution[j].Subject.Subject))
                            {
                                //Перед объединением показать, что именно объединяем
                                str = "У преподавателя " + mdlData.colDistribution[i].Lecturer.FIO + "\n";
                                str += "по дисциплине " + mdlData.colDistribution[i].Subject.Subject + ":\n";
                                str += mdlData.colDistribution[i].Code + ". " + mdlData.commentLoad(mdlData.colDistribution[i]) + "\nИ\n";
                                str += mdlData.colDistribution[j].Code + ". " + mdlData.commentLoad(mdlData.colDistribution[j]);

                                //И спросить, нужно ли объединять?
                                switch (MessageBox.Show(str, "Требуется ли объединение?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question))
                                {
                                        //Если даётся положительный ответ, то строки объединяются
                                    case DialogResult.Yes:
                                        {
                                            //Физически объединять имеем право только если дисциплина
                                            //НЕ ДИПЛОМНОЕ ПРОЕКТИРОВАНИЕ, НЕ ВЫПУСКНАЯ РАБОТА и НЕ ГАК

                                            if (!(mdlData.colDistribution[i].GAK > 0) & !(mdlData.colDistribution[i].DiplomaPaper > 0))
                                            {
                                                //В первую строку добавляем значения нагрузки из второй
                                                mdlData.colDistribution[i].Lecture += mdlData.colDistribution[j].Lecture;
                                                mdlData.colDistribution[i].Exam += mdlData.colDistribution[j].Exam;
                                                mdlData.colDistribution[i].Credit += mdlData.colDistribution[j].Credit;
                                                mdlData.colDistribution[i].RefHomeWork += mdlData.colDistribution[j].RefHomeWork;
                                                mdlData.colDistribution[i].Tutorial += mdlData.colDistribution[j].Tutorial;
                                                mdlData.colDistribution[i].LabWork += mdlData.colDistribution[j].LabWork;
                                                mdlData.colDistribution[i].Practice += mdlData.colDistribution[j].Practice;
                                                mdlData.colDistribution[i].IndividualWork += mdlData.colDistribution[j].IndividualWork;
                                                mdlData.colDistribution[i].KRAPK += mdlData.colDistribution[j].KRAPK;
                                                mdlData.colDistribution[i].KursProject += mdlData.colDistribution[j].KursProject;
                                                mdlData.colDistribution[i].PreDiplomaPractice += mdlData.colDistribution[j].PreDiplomaPractice;
                                                mdlData.colDistribution[i].DiplomaPaper += mdlData.colDistribution[j].DiplomaPaper;
                                                mdlData.colDistribution[i].TutorialPractice += mdlData.colDistribution[j].TutorialPractice;
                                                mdlData.colDistribution[i].ProducingPractice += mdlData.colDistribution[j].ProducingPractice;
                                                mdlData.colDistribution[i].GAK += mdlData.colDistribution[j].GAK;
                                                mdlData.colDistribution[i].PostGrad += mdlData.colDistribution[j].PostGrad;
                                                mdlData.colDistribution[i].EnteredHours += mdlData.colDistribution[j].EnteredHours;
                                                mdlData.colDistribution[i].Hours += mdlData.colDistribution[j].Hours;
                                                mdlData.colDistribution[i].Visiting += mdlData.colDistribution[j].Visiting;
                                                mdlData.colDistribution[i].Magistry += mdlData.colDistribution[j].Magistry;

                                                //Вторую строку удаляем
                                                mdlData.colDistribution.RemoveAt(j);
                                                flg = true;
                                            }
                                            break;
                                        }
                                        //Если даётся команда отмены - инициируем прерывание
                                    case DialogResult.Cancel:
                                        {
                                            flgBreak = true;
                                            break;
                                        }
                                }

                                //Если вызвано прерывание - завершаем цикл
                                if (flgBreak)
                                {
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            //Только если изменения случились, пересчитать
            toRefreshDistList(flg);
        }

        /// <summary>
        /// Метод обновления списка строк нагрузки
        /// </summary>
        /// <param name="flg"></param>
        private void toRefreshDistList(bool flg)
        {
            if (flg)
            {
                mdlData.toCombineDistribution(mdlData.colDistribution, mdlData.colCombineDistribution, mdlData.colHouredDistribution, true);

                FillDistributionList(cmbDistributionList, Selected, mdlData.cmbRem,
                                     chkSubjectFilt.Checked,
                                     chkKursFilt.Checked,
                                     chkSpecialityFilt.Checked,
                                     chkLecturerFilt.Checked,
                                     chkSemestrFilt.Checked,
                                     chkFacultyFilt.Checked,
                                     chkTypeFilt.Checked,
                                     false);

                mdlData.statString = "Требуется сохранение";
            }
        }

        /// <summary>
        /// Метод удаления пустых строк нагрузки
        /// </summary>
        private void toDelNullRows()
        {
            int sum;
            bool flg = false;
            bool flgBreak = false;
            string str;

            //Запускаем цикл по нагрузке
            for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
            {
                sum = mdlData.toSumDistributionComponents(mdlData.colDistribution[i]);

                if (sum == 0)
                {
                    //Перед удалением показать, что именно удаляем
                    str = "У преподавателя " + mdlData.colDistribution[i].Lecturer.FIO + "\n";
                    str += "нет ничего по дисциплине " + mdlData.colDistribution[i].Subject.Subject + ":\n";
                    str += mdlData.colDistribution[i].Code + ". " + mdlData.commentLoad(mdlData.colDistribution[i]);

                    //спросить, нужно ли удалять?
                    switch (MessageBox.Show(str, "Требуется ли удаление?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question))
                    {
                        //Если даётся положительный ответ, то строки объединяются
                        case DialogResult.Yes:
                            {
                                //Cтроку удаляем
                                mdlData.colDistribution.RemoveAt(i);
                                //Фиксируем внесённые изменения
                                flg = true;
                                //Переводим параметр цикла в исходное состояние
                                i = -1;
                                //Прерываем
                                break;
                            }
                        case DialogResult.Cancel:
                            {
                                flgBreak = true;
                                break;
                            }
                    }

                    if (flgBreak)
                    {
                        break;
                    }
                }
            }

            //Только если изменения случились, пересчитать
            toRefreshDistList(flg);
        }

        private void CheckFiltParam(CheckBox chkCur, ComboBox cmbCur, int inx, ref bool flg)
        {
            if (chkCur.Checked)
            {
                cmbCur.Enabled = true;
                cmbCur.SelectedIndex = inx;

                if (cmbCur == cmbTypeFilt)
                {
                    if (cmbCur.SelectedItem != null)
                    {
                        mdlData.cmbRem = cmbCur.SelectedItem.ToString();
                    }
                    else
                    {
                        mdlData.cmbRem = "";
                    }
                }
            }
            else
            {
                cmbCur.Enabled = false;
                cmbCur.SelectedIndex = -1;

                if (cmbCur == cmbTypeFilt)
                {
                    mdlData.cmbRem = "";
                }

                FillDistributionList(cmbDistributionList, Selected, mdlData.cmbRem,
                                     chkSubjectFilt.Checked,
                                     chkKursFilt.Checked,
                                     chkSpecialityFilt.Checked,
                                     chkLecturerFilt.Checked,
                                     chkSemestrFilt.Checked,
                                     chkFacultyFilt.Checked,
                                     chkTypeFilt.Checked,
                                     false);
            }

            //Сохраняем глобальное состояние галочки
            flg = chkCur.Checked;
        }

        private void chkFacultyFilt_CheckedChanged(object sender, EventArgs e)
        {
            CheckFiltParam(chkFacultyFilt, cmbFacultyFilt, mdlData.inxFaculty, ref mdlData.flgFacultyFilt);
        }

        private void cmbFacultyFilt_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeFiltParam(cmbFacultyFilt, ref mdlData.inxFaculty);
        }

        private void SaveChanges()
        {
            int curIndex;
            IList<clsDistribution> coll = null;

            if (MessageBox.Show(this, "Сохранить выполненные изменения?", "Сохранение",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                if (optMain.Checked)
                {
                    coll = mdlData.colDistribution;
                }
                if (optHoured.Checked)
                {
                    coll = mdlData.colHouredDistribution;
                }
                if (optCombine.Checked)
                {
                    coll = mdlData.colCombineDistribution;
                }

                if (chkSubjectFilt.Checked || chkKursFilt.Checked || chkSpecialityFilt.Checked
                    || chkLecturerFilt.Checked || chkSemestrFilt.Checked || chkFacultyFilt.Checked
                    || chkTypeFilt.Checked)
                {
                    coll = mdlData.Filtred;
                }

                //Сохраняем лекционные часы
                if (!(txtLecture.Text == ""))
                {
                    coll[cmbDistributionList.SelectedIndex].Lecture = Convert.ToInt32(txtLecture.Text);
                }
                else
                {
                    coll[cmbDistributionList.SelectedIndex].Lecture = 0;
                }

                //Сохраняем экзаминационные часы
                coll[cmbDistributionList.SelectedIndex].Exam = Convert.ToInt32(txtExam.Text);
                //Сохраняем зачётные часы
                coll[cmbDistributionList.SelectedIndex].Credit = Convert.ToInt32(txtCred.Text);
                //Сохраняем часы на реферат и домашнюю работу
                coll[cmbDistributionList.SelectedIndex].RefHomeWork = Convert.ToInt32(txtRef.Text);
                //Сохраняем часы на консультацию
                coll[cmbDistributionList.SelectedIndex].Tutorial = Convert.ToInt32(txtTut.Text);
                //Сохраняем часы на лабораторные работы
                coll[cmbDistributionList.SelectedIndex].LabWork = Convert.ToInt32(txtLab.Text);
                //Сохраняем часы на практические занятия
                coll[cmbDistributionList.SelectedIndex].Practice = Convert.ToInt32(txtPract.Text);
                //Сохраняем часы на индивидуальные занятия
                coll[cmbDistributionList.SelectedIndex].IndividualWork = Convert.ToInt32(txtInd.Text);
                //Сохраняем часы на КРАПК
                coll[cmbDistributionList.SelectedIndex].KRAPK = Convert.ToInt32(txtKRAPK.Text);
                //Сохраняем часы на курсовой проект
                coll[cmbDistributionList.SelectedIndex].KursProject = Convert.ToInt32(txtKursPr.Text);
                //Сохраняем часы на преддипломную практику
                coll[cmbDistributionList.SelectedIndex].PreDiplomaPractice = Convert.ToInt32(txtPreD.Text);
                //Сохраняем часы на дипломный проект
                coll[cmbDistributionList.SelectedIndex].DiplomaPaper = Convert.ToInt32(txtDiploma.Text);
                //Сохраняем часы на учебную практику
                coll[cmbDistributionList.SelectedIndex].TutorialPractice = Convert.ToInt32(txtTutPr.Text);
                //Сохраняем часы на производственную практику
                coll[cmbDistributionList.SelectedIndex].ProducingPractice = Convert.ToInt32(txtProd.Text);
                //Сохраняем часы на ГЭК
                coll[cmbDistributionList.SelectedIndex].GAK = Convert.ToInt32(txtGAK.Text);
                //Сохраняем часы
                coll[cmbDistributionList.SelectedIndex].Hours = Convert.ToInt32(txtHours.Text);
                //Сохраняем часы в ЗЕТ
                coll[cmbDistributionList.SelectedIndex].HoursZ = Convert.ToSingle(txtHoursZ.Text);
                //Сохраняем часы_дано
                coll[cmbDistributionList.SelectedIndex].EnteredHours = Convert.ToInt32(txtEnteredHours.Text);
                //Сохраняем часы_дано в ЗЕТ
                coll[cmbDistributionList.SelectedIndex].EnteredHoursZ = Convert.ToSingle(txtEnteredHoursZ.Text);
                //Сохраняем часы на аспирантуру
                coll[cmbDistributionList.SelectedIndex].PostGrad = Convert.ToInt32(txtPostGrad.Text);
                //Сохраняем часы на посещение занятий
                coll[cmbDistributionList.SelectedIndex].Visiting = Convert.ToInt32(txtVisiting.Text);
                //Сохраняем примечание для диспетчерской
                coll[cmbDistributionList.SelectedIndex].Text = txtText.Text;
                //Сохраняем часы на магистерскую программу
                coll[cmbDistributionList.SelectedIndex].Magistry = Convert.ToInt32(txtMagistry.Text);
                //Признак необходимости записи в заявку для диспетчерской
                coll[cmbDistributionList.SelectedIndex].flgDispatch = chkDispatch.Checked;
                //Признак равномерно распределяемой нагрузки
                coll[cmbDistributionList.SelectedIndex].flgDistrib = chkDistrib.Checked;
                //Сохраняем вес единицы нагрузки
                coll[cmbDistributionList.SelectedIndex].Weight = Convert.ToInt32(txtWeight.Text);
                //Признак участия строки в расчёте нагрузки
                coll[cmbDistributionList.SelectedIndex].flgExclude = chkExclude.Checked;
                //Сохраняем введённый код дисциплины по документу
                coll[cmbDistributionList.SelectedIndex].DocCode = Convert.ToInt32(txtDocCode.Text);

                //Сохраняем дисциплину
                if (cmbSubjectList.SelectedIndex >= 0)
                {
                    coll[cmbDistributionList.SelectedIndex].Subject = mdlData.colSubject[cmbSubjectList.SelectedIndex];
                }
                else
                {
                    coll[cmbDistributionList.SelectedIndex].Subject = null;
                }

                //Сохраняем номер курса
                if (cmbKursList.SelectedIndex >= 0)
                {
                    coll[cmbDistributionList.SelectedIndex].KursNum = mdlData.colKursNum[cmbKursList.SelectedIndex];
                }
                else
                {
                    coll[cmbDistributionList.SelectedIndex].KursNum = null;
                }

                //Сохраняем специальность
                if (cmbSpecialityList.SelectedIndex >= 0)
                {
                    coll[cmbDistributionList.SelectedIndex].Speciality = mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex];
                }
                else
                {
                    coll[cmbDistributionList.SelectedIndex].Speciality = null;
                }

                if (!coll[cmbDistributionList.SelectedIndex].flgDistrib)
                {
                    //Сохраняем преподавателя
                    if (cmbLecturerList.SelectedIndex >= 0)
                    {
                        coll[cmbDistributionList.SelectedIndex].Lecturer = mdlData.colLecturer[cmbLecturerList.SelectedIndex - 1];
                    }
                    else
                    {
                        coll[cmbDistributionList.SelectedIndex].Lecturer = null;
                    }
                }
                else
                {
                    coll[cmbDistributionList.SelectedIndex].Lecturer = null;
                }

                //Сохраняем семестр
                if (cmbSemestrList.SelectedIndex >= 0)
                {
                    coll[cmbDistributionList.SelectedIndex].Semestr = mdlData.colSemestr[cmbSemestrList.SelectedIndex];
                }
                else
                {
                    coll[cmbDistributionList.SelectedIndex].Semestr = null;
                }

                //Сохраняем замещающего преподавателя
                if (cmbLecturer2List.SelectedIndex > 0)
                {
                    coll[cmbDistributionList.SelectedIndex].Lecturer2 = mdlData.colLecturer[cmbLecturer2List.SelectedIndex - 1];
                }
                else
                {
                    coll[cmbDistributionList.SelectedIndex].Lecturer2 = null;
                }

                //Сохраняем дополнительного преподавателя
                if (cmbLecturer3List.SelectedIndex > 0)
                {
                    coll[cmbDistributionList.SelectedIndex].Lecturer3 = mdlData.colLecturer[cmbLecturer3List.SelectedIndex - 1];
                }
                else
                {
                    coll[cmbDistributionList.SelectedIndex].Lecturer3 = null;
                }

                //Сохраняем преподавателя дублёра
                if (cmbDoubler.SelectedIndex > 0)
                {
                    coll[cmbDistributionList.SelectedIndex].Doubler = mdlData.colLecturer[cmbDoubler.SelectedIndex - 1];
                }
                else
                {
                    coll[cmbDistributionList.SelectedIndex].Doubler = null;
                }

                //Если существует связка по лабораторным работам, то
                //сохраняем связку по лабораторным работам
                if (cmbLabWorkConnect.SelectedIndex > 0)
                {
                    //Текущий элемент связываем с указанным
                    coll[cmbDistributionList.SelectedIndex].LabWorkConnect = Selected[cmbLabWorkConnect.SelectedIndex - 1];
                    //Указанный элемент связываем с текущим
                    Selected[cmbLabWorkConnect.SelectedIndex - 1].LabWorkConnect = coll[cmbDistributionList.SelectedIndex];
                }
                //Если не существует связки по лабораторным работам
                else
                {
                    //Если текущий элемент связан с каким-либо
                    if (coll[cmbDistributionList.SelectedIndex].LabWorkConnect != null)
                    {
                        if (coll.IndexOf(coll[cmbDistributionList.SelectedIndex].LabWorkConnect) > -1)
                        {
                            //Устраняем связь в указанном с текущим
                            coll[coll.IndexOf(coll[cmbDistributionList.SelectedIndex].LabWorkConnect)].LabWorkConnect = null;
                        }
                    }
                    //Устраняем связь в текущем с указанным
                    coll[cmbDistributionList.SelectedIndex].LabWorkConnect = null;
                }

                //Запоминаем позицию, на которой находились перед обновлением
                curIndex = cmbDistributionList.SelectedIndex;

                //Пересчитываем комбинированную нагрузку
                mdlData.toCombineDistribution(mdlData.colDistribution, mdlData.colCombineDistribution,
                                          mdlData.colHouredDistribution, true);

                //Обновляем сопряжение по лабораторным работам
                FillDistributionList(cmbLabWorkConnect, Selected, mdlData.cmbRem,
                                     chkSubjectFilt.Checked,
                                     chkKursFilt.Checked,
                                     chkSpecialityFilt.Checked,
                                     chkLecturerFilt.Checked,
                                     chkSemestrFilt.Checked,
                                     chkFacultyFilt.Checked,
                                     chkTypeFilt.Checked,
                                     true);

                //Обновляем данные по распределению
                FillDistributionList(cmbDistributionList, Selected, mdlData.cmbRem,
                                     chkSubjectFilt.Checked,
                                     chkKursFilt.Checked,
                                     chkSpecialityFilt.Checked,
                                     chkLecturerFilt.Checked,
                                     chkSemestrFilt.Checked,
                                     chkFacultyFilt.Checked,
                                     chkTypeFilt.Checked, 
                                     false);

                //Возвращаем позицию, на которой находились перед обновлением
                if ((cmbDistributionList.Items.Count > 0) &
                    (cmbDistributionList.Items.Count > curIndex))
                {
                    cmbDistributionList.SelectedIndex = curIndex;
                }
                //Иначе необходима дополнительная проверка
                //для правильного позиционирования
                else
                {
                    if (cmbDistributionList.Items.Count <= curIndex)
                    {
                        cmbDistributionList.SelectedIndex = curIndex - 1;
                    }
                    else
                    {
                        cmbDistributionList.SelectedIndex = -1;
                    }
                }

                mdlData.statString = "Последнее действие: Сохранение строки учебной нагрузки";
            }
        }

        private void DelObject()
        {
            int curIndex;
            IList<clsDistribution> coll = null;
            clsDistribution DelItem = null;

            if (MessageBox.Show(this, "Действительно удалить?", "Удаление", 
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                //Если выбрана основная нагрузка, то работаем с основной нагрузкой
                if (optMain.Checked)
                {
                    coll = mdlData.colDistribution;
                }

                //Если выбрана почасовая нагрузка, то работаем с почасовой нагрузкой
                if (optHoured.Checked)
                {
                    coll = mdlData.colHouredDistribution;
                }

                //Если выбрана комбинированная нагрузка, то работаем с комбинированной
                //нагрузкой
                if (optCombine.Checked)
                {
                    coll = mdlData.colCombineDistribution;
                }

                //Если установлен какой-либо фильтр, то работаем с фильтрованной нагрузкой
                if (chkSubjectFilt.Checked || chkKursFilt.Checked || chkSpecialityFilt.Checked
                    || chkLecturerFilt.Checked || chkSemestrFilt.Checked || chkFacultyFilt.Checked
                    || chkTypeFilt.Checked)
                {
                    coll = mdlData.Filtred;
                }

                //Запоминаем индекс до удаления элемента
                curIndex = cmbDistributionList.SelectedIndex;
                //Запоминаем удаляемый компонент
                DelItem = coll[curIndex];

                //
                for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
                {
                    if (FullyCompare(mdlData.colDistribution[i], DelItem))
                    {
                        mdlData.colDistribution.RemoveAt(i);
                    }
                }

                for (int i = 0; i <= mdlData.colHouredDistribution.Count - 1; i++)
                {
                    if (FullyCompare(mdlData.colHouredDistribution[i], DelItem))
                    {
                        mdlData.colHouredDistribution.RemoveAt(i);
                    }
                }

                for (int i = 0; i <= mdlData.colCombineDistribution.Count - 1; i++)
                {
                    if (FullyCompare(mdlData.colCombineDistribution[i], DelItem))
                    {
                        mdlData.colCombineDistribution.RemoveAt(i);
                    }
                }

                for (int i = 0; i <= mdlData.Filtred.Count - 1; i++)
                {
                    if (FullyCompare(mdlData.Filtred[i], DelItem))
                    {
                        mdlData.Filtred.RemoveAt(i);
                    }
                }

                //Обновляем данные по распределению
                FillDistributionList(cmbDistributionList, Selected, mdlData.cmbRem,
                                     chkSubjectFilt.Checked,
                                     chkKursFilt.Checked,
                                     chkSpecialityFilt.Checked,
                                     chkLecturerFilt.Checked,
                                     chkSemestrFilt.Checked,
                                     chkFacultyFilt.Checked,
                                     chkTypeFilt.Checked,
                                     false);

                //Возвращаем индекс после удаления и обновления
                if (curIndex <= cmbDistributionList.Items.Count - 1)
                {
                    cmbDistributionList.SelectedIndex = curIndex;
                }
                else
                {
                    cmbDistributionList.SelectedIndex = curIndex - 1;
                }

                mdlData.statString = "Последнее действие: Удаление строки учебной нагрузки";
            }
        }

        /// <summary>
        /// Типизированный метод поиска соответствия одного элемента нагрузки другому
        /// </summary>
        /// <param name="D1"></param>
        /// <param name="D2"></param>
        /// <returns></returns>
        private bool FullyCompare(clsDistribution D1, clsDistribution D2)
        {
            bool flgKurs = false;
            bool flgLecturer = false;
            bool flgLecturer2 = false;
            bool flgLecturer3 = false;
            bool flgSemestr = false;
            bool flgSpeciality = false;
            bool flgSubject = false;
            
            if (D1.KursNum != null & D2.KursNum != null)
            {
                flgKurs = D1.KursNum.Code.Equals(D2.KursNum.Code);
            }
            else
            {
                if (D1.KursNum == null & D2.KursNum == null)
                {
                    flgKurs = true;
                }
            }

            if (D1.Lecturer != null & D2.Lecturer != null)
            {
                flgLecturer = D1.Lecturer.Code.Equals(D2.Lecturer.Code);
            }
            else
            {
                if (D1.Lecturer == null & D2.Lecturer == null)
                {
                    flgLecturer = true;
                }
            }

            if (D1.Lecturer2 != null & D2.Lecturer2 != null)
            {
                flgLecturer2 = D1.Lecturer2.Code.Equals(D2.Lecturer2.Code);
            }
            else
            {
                if (D1.Lecturer2 == null & D2.Lecturer2 == null)
                {
                    flgLecturer2 = true;
                }
            }

            if (D1.Lecturer3 != null & D2.Lecturer3 != null)
            {
                flgLecturer3 = D1.Lecturer3.Code.Equals(D2.Lecturer3.Code);
            }
            else
            {
                if (D1.Lecturer3 == null & D2.Lecturer3 == null)
                {
                    flgLecturer3 = true;
                }
            }

            if (D1.Semestr != null & D2.Semestr != null)
            {
                flgSemestr = D1.Semestr.Code.Equals(D2.Semestr.Code);
            }
            else
            {
                if (D1.Semestr == null & D2.Semestr == null)
                {
                    flgSemestr = true;
                }
            }

            if (D1.Speciality != null & D2.Speciality != null)
            {
                flgSpeciality = D1.Speciality.Code.Equals(D2.Speciality.Code);
            }
            else
            {
                if (D1.Speciality == null & D2.Speciality == null)
                {
                    flgSpeciality = true;
                }
            }

            if (D1.Subject != null & D2.Subject != null)
            {
                flgSubject = D1.Subject.Code.Equals(D2.Subject.Code);
            }
            else
            {
                if (D1.Subject == null & D2.Subject == null)
                {
                    flgSubject = true;
                }
            }

            bool flg = (D1.Code.Equals(D2.Code) &
                        D1.Credit.Equals(D2.Credit) &
                        D1.DiplomaPaper.Equals(D2.DiplomaPaper) &
                        D1.EnteredHours.Equals(D2.EnteredHours) &
                        D1.Exam.Equals(D2.Exam) &
                        D1.GAK.Equals(D2.GAK) &
                        D1.Hours.Equals(D2.Hours) &
                        D1.IndividualWork.Equals(D2.IndividualWork) &
                        D1.KRAPK.Equals(D2.KRAPK) &
                        
                        flgKurs &
                        
                        D1.KursProject.Equals(D2.KursProject) &
                        D1.LabWork.Equals(D2.LabWork) &
                        D1.Lecture.Equals(D2.Lecture) &
                       
                        flgLecturer &
                        flgLecturer2 &
                        flgLecturer3 &
                        
                        D1.Magistry.Equals(D2.Magistry) &
                        D1.PostGrad.Equals(D2.PostGrad) &
                        D1.Practice.Equals(D2.Practice) &
                        D1.PreDiplomaPractice.Equals(D2.PreDiplomaPractice) &
                        D1.ProducingPractice.Equals(D2.ProducingPractice) &
                        D1.RefHomeWork.Equals(D2.RefHomeWork) &
                        
                        flgSemestr &
                        flgSpeciality &
                        flgSubject &
                        
                        D1.Text.Equals(D2.Text) &
                        D1.Tutorial.Equals(D2.Tutorial) &
                        D1.TutorialPractice.Equals(D2.TutorialPractice) &
                        D1.Visiting.Equals(D2.Visiting));
            return flg;
        }

        private void toCreateNew()
        {
            clsDistribution Obj;
            MaxNum += 1;
            Obj = new clsDistribution();

            Obj.Code = MaxNum;

            IList<clsDistribution> coll = null;

            if (optMain.Checked)
            {
                coll = mdlData.colDistribution;
            }
            if (optHoured.Checked)
            {
                coll = mdlData.colHouredDistribution;
            }
            if (optCombine.Checked)
            {
                coll = mdlData.colCombineDistribution;
            }

            if (chkSubjectFilt.Checked || chkKursFilt.Checked || chkSpecialityFilt.Checked
                || chkLecturerFilt.Checked || chkSemestrFilt.Checked || chkFacultyFilt.Checked
                || chkTypeFilt.Checked)
            {
                coll = mdlData.Filtred;
            }

            //Добавляем новый объект в коллекцию
            coll.Add(Obj);

            //Обновляем данные по распределению
            FillDistributionList(cmbDistributionList, Selected, mdlData.cmbRem,
                                 chkSubjectFilt.Checked,
                                 chkKursFilt.Checked,
                                 chkSpecialityFilt.Checked,
                                 chkLecturerFilt.Checked,
                                 chkSemestrFilt.Checked,
                                 chkFacultyFilt.Checked,
                                 chkTypeFilt.Checked, 
                                 false);

            //Переходим к последнему элементу
            cmbDistributionList.SelectedIndex = coll.Count - 1;
        }

        private void toCopy()
        {
            //Создаём новый объект класса "Распределение"
            clsDistribution Obj = new clsDistribution();
            //К максимальному номеру в коллекции добавляем единицу
            //для создания нового уникального элемента
            MaxNum += 1;
            //Кодом нового объекта устанавливаем увеличенный на единицу
            //максимальный номер
            Obj.Code = MaxNum;

            //Назначаем параметры выбранного объекта новому объекту

            //Копируем в новый объект лекционные часы
            Obj.Lecture = Convert.ToInt32(txtLecture.Text);
            //Копируем в новый объект экзаминационные часы
            Obj.Exam = Convert.ToInt32(txtExam.Text);
            //Копируем в новый объект зачётные часы
            Obj.Credit = Convert.ToInt32(txtCred.Text);
            //Копируем в новый объект часы на реферат и домашнюю работу
            Obj.RefHomeWork = Convert.ToInt32(txtRef.Text);
            //Копируем в новый объект часы на консультацию
            Obj.Tutorial = Convert.ToInt32(txtTut.Text);
            //Копируем в новый объект часы на лабораторные работы
            Obj.LabWork = Convert.ToInt32(txtLab.Text);
            //Копируем в новый объект часы на практические занятия
            Obj.Practice = Convert.ToInt32(txtPract.Text);
            //Копируем в новый объект часы на индивидуальные занятия
            Obj.IndividualWork = Convert.ToInt32(txtInd.Text);
            //Копируем в новый объект часы на КРАПК
            Obj.KRAPK = Convert.ToInt32(txtKRAPK.Text);
            //Копируем в новый объект часы на курсовой проект
            Obj.KursProject = Convert.ToInt32(txtKursPr.Text);
            //Копируем в новый объект часы на преддипломную практику
            Obj.PreDiplomaPractice = Convert.ToInt32(txtPreD.Text);
            //Копируем в новый объект часы на дипломный проект
            Obj.DiplomaPaper = Convert.ToInt32(txtDiploma.Text);
            //Копируем в новый объект часы на учебную практику
            Obj.TutorialPractice = Convert.ToInt32(txtTutPr.Text);
            //Копируем в новый объект часы на производственную практику
            Obj.ProducingPractice = Convert.ToInt32(txtProd.Text);
            //Копируем в новый объект часы на ГЭК
            Obj.GAK = Convert.ToInt32(txtGAK.Text);
            //Копируем в новый объект часы
            Obj.Hours = Convert.ToInt32(txtHours.Text);
            //Копируем в новый объект часы в ЗЕТ
            Obj.HoursZ = Convert.ToSingle(txtHoursZ.Text);
            //Копируем в новый объект часы_дано
            Obj.EnteredHours = Convert.ToInt32(txtEnteredHours.Text);
            //Копируем в новый объект часы_дано в ЗЕТ
            Obj.EnteredHoursZ = Convert.ToSingle(txtEnteredHoursZ.Text);
            //Копируем в новый объект часы на аспирантуру
            Obj.PostGrad = Convert.ToInt32(txtPostGrad.Text);
            //Копируем в новый объект часы на посещение занятий
            Obj.Visiting = Convert.ToInt32(txtVisiting.Text);
            //Копируем в новый объект примечание для диспетчерской
            Obj.Text = txtText.Text;
            //Копируем в новый объект часы на магистерскую программу
            Obj.Magistry = Convert.ToInt32(txtMagistry.Text);
            //Копируем вес единицы нагрузки
            Obj.Weight = Convert.ToInt32(txtWeight.Text);
            //Копируем признак учёта в заявке для диспетчерской
            Obj.flgDispatch = chkDispatch.Checked;
            //Копируем признак равномерного распределения элемента нагрузки
            Obj.flgDistrib = chkDistrib.Checked;
            //Копируем признак исключения из расчёта нагрузки
            Obj.flgExclude = chkExclude.Checked;
            //Копируем код согласно документу
            Obj.DocCode = Convert.ToInt32(txtDocCode.Text);

            //Копируем в новый объект дисциплину
            if (cmbSubjectList.SelectedIndex >= 0)
            {
                Obj.Subject = mdlData.colSubject[cmbSubjectList.SelectedIndex];
            }
            else
            {
                Obj.Subject = null;
            }
            //Копируем в новый объект номер курса
            if (cmbKursList.SelectedIndex >= 0)
            {
                Obj.KursNum = mdlData.colKursNum[cmbKursList.SelectedIndex];
            }
            else
            {
                Obj.KursNum = null;
            }
            //Копируем в новый объект специальность
            if (cmbSpecialityList.SelectedIndex >= 0)
            {
                Obj.Speciality = mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex];
            }
            else
            {
                Obj.Speciality = null;
            }
            //Копируем в новый объект преподавателя
            if (cmbLecturerList.SelectedIndex >= 0)
            {
                Obj.Lecturer = mdlData.colLecturer[cmbLecturerList.SelectedIndex - 1];
            }
            else
            {
                Obj.Lecturer = null;
            }
            //Копируем в новый объект семестр
            if (cmbSemestrList.SelectedIndex >= 0)
            {
                Obj.Semestr = mdlData.colSemestr[cmbSemestrList.SelectedIndex];
            }
            else
            {
                Obj.Semestr = null;
            }
            //Копируем в новый объект замещающего преподавателя
            //(индекс смещён, так как есть строчка "не выбран")
            if (cmbLecturer2List.SelectedIndex > 0)
            {
                Obj.Lecturer2 = mdlData.colLecturer[cmbLecturer2List.SelectedIndex - 1];
            }
            else
            {
                Obj.Lecturer2 = null;
            }
            //Сохраняем дополнительного преподавателя
            //(индекс смещён, так как есть строчка "не выбран")
            if (cmbLecturer3List.SelectedIndex > 0)
            {
                Obj.Lecturer3 = mdlData.colLecturer[cmbLecturer3List.SelectedIndex - 1];
            }
            else
            {
                Obj.Lecturer3 = null;
            }

            //Назначаем параметры выбранного объекта новому объекту

            //Сбрасываем содержимое рабочей коллекции
            IList<clsDistribution> coll = null;
            //Если работаем со штатным распределением нагрузки
            if (optMain.Checked)
            {
                //берём коллекцию штатной нагрузки
                coll = mdlData.colDistribution;
            }
            //Если работаем с почасовым распределением нагрузки
            if (optHoured.Checked)
            {
                //берём коллекцию почасовой нагрузки
                coll = mdlData.colHouredDistribution;
            }
            //Если работаем с комбинированным распределением нагрузки
            if (optCombine.Checked)
            {
                //берём коллекцию комбинированной нагрузки
                coll = mdlData.colCombineDistribution;
            }

            //Если включён хотя бы один из фильров, то
            if (chkSubjectFilt.Checked || chkKursFilt.Checked || chkSpecialityFilt.Checked
                || chkLecturerFilt.Checked || chkSemestrFilt.Checked || chkFacultyFilt.Checked
                || chkTypeFilt.Checked)
            {
                //берём коллекцию отфильтрованной нагрузки
                coll = mdlData.Filtred;
            }

            //Добавляем новый объект в коллекцию
            coll.Add(Obj);

            //Обновляем данные по распределению
            FillDistributionList(cmbDistributionList, Selected, mdlData.cmbRem,
                                 chkSubjectFilt.Checked,
                                 chkKursFilt.Checked,
                                 chkSpecialityFilt.Checked,
                                 chkLecturerFilt.Checked,
                                 chkSemestrFilt.Checked,
                                 chkFacultyFilt.Checked,
                                 chkTypeFilt.Checked,
                                 false);

            //Переходим к последнему элементу
            cmbDistributionList.SelectedIndex = coll.Count - 1;
        }

        private void toCopyFilt()
        {
            //Создаём новый объект класса "Распределение"
            clsDistribution Obj = new clsDistribution();
            //К максимальному номеру в коллекции добавляем единицу
            //для создания нового уникального элемента
            MaxNum += 1;
            //Кодом нового объекта устанавливаем увеличенный на единицу
            //максимальный номер
            Obj.Code = MaxNum;

            //Назначаем параметры выбранного объекта новому объекту

            //Копируем в новый объект лекционные часы
            Obj.Lecture = Convert.ToInt32(txtLecture.Text);
            //Копируем в новый объект экзаминационные часы
            Obj.Exam = Convert.ToInt32(txtExam.Text);
            //Копируем в новый объект зачётные часы
            Obj.Credit = Convert.ToInt32(txtCred.Text);
            //Копируем в новый объект часы на реферат и домашнюю работу
            Obj.RefHomeWork = Convert.ToInt32(txtRef.Text);
            //Копируем в новый объект часы на консультацию
            Obj.Tutorial = Convert.ToInt32(txtTut.Text);
            //Копируем в новый объект часы на лабораторные работы
            Obj.LabWork = Convert.ToInt32(txtLab.Text);
            //Копируем в новый объект часы на практические занятия
            Obj.Practice = Convert.ToInt32(txtPract.Text);
            //Копируем в новый объект часы на индивидуальные занятия
            Obj.IndividualWork = Convert.ToInt32(txtInd.Text);
            //Копируем в новый объект часы на КРАПК
            Obj.KRAPK = Convert.ToInt32(txtKRAPK.Text);
            //Копируем в новый объект часы на курсовой проект
            Obj.KursProject = Convert.ToInt32(txtKursPr.Text);
            //Копируем в новый объект часы на преддипломную практику
            Obj.PreDiplomaPractice = Convert.ToInt32(txtPreD.Text);
            //Копируем в новый объект часы на дипломный проект
            Obj.DiplomaPaper = Convert.ToInt32(txtDiploma.Text);
            //Копируем в новый объект часы на учебную практику
            Obj.TutorialPractice = Convert.ToInt32(txtTutPr.Text);
            //Копируем в новый объект часы на производственную практику
            Obj.ProducingPractice = Convert.ToInt32(txtProd.Text);
            //Копируем в новый объект часы на ГЭК
            Obj.GAK = Convert.ToInt32(txtGAK.Text);
            //Копируем в новый объект часы
            Obj.Hours = Convert.ToInt32(txtHours.Text);
            //Копируем в новый объект часы в ЗЕТ
            Obj.HoursZ = Convert.ToSingle(txtHoursZ.Text);
            //Копируем в новый объект часы_дано
            Obj.EnteredHours = Convert.ToInt32(txtEnteredHours.Text);
            //Копируем в новый объект часы_дано в ЗЕТ
            Obj.EnteredHoursZ = Convert.ToSingle(txtEnteredHoursZ.Text);
            //Копируем в новый объект часы на аспирантуру
            Obj.PostGrad = Convert.ToInt32(txtPostGrad.Text);
            //Копируем в новый объект часы на посещение занятий
            Obj.Visiting = Convert.ToInt32(txtVisiting.Text);
            //Копируем в новый объект примечание для диспетчерской
            Obj.Text = txtText.Text;
            //Копируем в новый объект часы на магистерскую программу
            Obj.Magistry = Convert.ToInt32(txtMagistry.Text);
            //Копируем вес единицы нагрузки
            Obj.Weight = Convert.ToInt32(txtWeight.Text);
            //Копируем признак учёта в заявке для диспетчерской
            Obj.flgDispatch = chkDispatch.Checked;
            //Копируем признак равномерного распределения элемента нагрузки
            Obj.flgDistrib = chkDistrib.Checked;
            //Копируем признак исключения из расчёта нагрузки
            Obj.flgExclude = chkExclude.Checked;
            //Копируем код согласно документу
            Obj.DocCode = Convert.ToInt32(txtDocCode.Text);

            //Копируем в новый объект дисциплину
            if (cmbSubjectList.SelectedIndex >= 0)
            {
                Obj.Subject = mdlData.colSubject[cmbSubjectList.SelectedIndex];
            }
            else
            {
                Obj.Subject = null;
            }

            //Копируем в новый объект номер курса
            if (cmbKursList.SelectedIndex >= 0)
            {
                Obj.KursNum = mdlData.colKursNum[cmbKursList.SelectedIndex];
            }
            else
            {
                Obj.KursNum = null;
            }

            //Копируем в новый объект специальность
            if (cmbSpecialityList.SelectedIndex >= 0)
            {
                Obj.Speciality = mdlData.colSpecialisation[cmbSpecialityList.SelectedIndex];
            }
            else
            {
                Obj.Speciality = null;
            }

            //Копируем в новый объект преподавателя
            if (cmbLecturerList.SelectedIndex >= 0)
            {
                Obj.Lecturer = mdlData.colLecturer[cmbLecturerList.SelectedIndex - 1];
            }
            else
            {
                Obj.Lecturer = null;
            }

            //Копируем в новый объект семестр
            if (cmbSemestrList.SelectedIndex >= 0)
            {
                Obj.Semestr = mdlData.colSemestr[cmbSemestrList.SelectedIndex];
            }
            else
            {
                Obj.Semestr = null;
            }

            //Копируем в новый объект замещающего преподавателя
            //(индекс смещён, так как есть строчка "не выбран")
            if (cmbLecturer2List.SelectedIndex > 0)
            {
                Obj.Lecturer2 = mdlData.colLecturer[cmbLecturer2List.SelectedIndex - 1];
            }
            else
            {
                Obj.Lecturer2 = null;
            }

            //Сохраняем дополнительного преподавателя
            //(индекс смещён, так как есть строчка "не выбран")
            if (cmbLecturer3List.SelectedIndex > 0)
            {
                Obj.Lecturer3 = mdlData.colLecturer[cmbLecturer3List.SelectedIndex - 1];
            }
            else
            {
                Obj.Lecturer3 = null;
            }

            //Назначаем параметры выбранного объекта новому объекту

            //Сбрасываем содержимое рабочей коллекции
            IList<clsDistribution> coll = null;
            //Если работаем со штатным распределением нагрузки
            if (optMain.Checked)
            {
                //берём коллекцию штатной нагрузки
                coll = mdlData.colDistribution;
            }
            //Если работаем с почасовым распределением нагрузки
            if (optHoured.Checked)
            {
                //берём коллекцию почасовой нагрузки
                coll = mdlData.colHouredDistribution;
            }
            //Если работаем с комбинированным распределением нагрузки
            if (optCombine.Checked)
            {
                //берём коллекцию комбинированной нагрузки
                coll = mdlData.colCombineDistribution;
            }

            //Если включён хотя бы один из фильров, то
            if (chkSubjectFilt.Checked || chkKursFilt.Checked || chkSpecialityFilt.Checked
                || chkLecturerFilt.Checked || chkSemestrFilt.Checked || chkFacultyFilt.Checked
                || chkTypeFilt.Checked)
            {
                //берём коллекцию отфильтрованной нагрузки
                coll = mdlData.Filtred;
            }

            //Добавляем новый объект в отображаемую коллекцию
            coll.Add(Obj);
            //Добавляем новый объект в зафиксированную коллекцию
            Selected.Add(Obj);

            //Обновляем данные по увязкам лабораторных работ
            FillDistributionList(cmbLabWorkConnect, Selected, mdlData.cmbRem,
                                 chkSubjectFilt.Checked,
                                 chkKursFilt.Checked,
                                 chkSpecialityFilt.Checked,
                                 chkLecturerFilt.Checked,
                                 chkSemestrFilt.Checked,
                                 chkFacultyFilt.Checked,
                                 chkTypeFilt.Checked,
                                 true);

            //Обновляем данные по распределению
            FillDistributionList(cmbDistributionList, Selected, mdlData.cmbRem,
                                 chkSubjectFilt.Checked,
                                 chkKursFilt.Checked,
                                 chkSpecialityFilt.Checked,
                                 chkLecturerFilt.Checked,
                                 chkSemestrFilt.Checked,
                                 chkFacultyFilt.Checked,
                                 chkTypeFilt.Checked,
                                 false);

            //Переходим к последнему элементу
            cmbDistributionList.SelectedIndex = coll.Count - 1;
        }

        private void cmbLecturer2List_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbLecturer2List.SelectedIndex >= 0)
            {
                toolTip.SetToolTip(cmbLecturer2List, cmbLecturer2List.Items[cmbLecturer2List.SelectedIndex].ToString());
            }
        }

        private void toDoAnalysis()
        {
            int AllSubjects;
            int UniqueSubjects;
            bool flgUnique;
            IList<clsSubject> Sb = new List<clsSubject>();

            AllSubjects = mdlData.colSubject.Count;
            UniqueSubjects = 0;
            for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
            {
                flgUnique = true;
                for (int j = 0; j <= Sb.Count - 1; j++)
                {
                    if (mdlData.colDistribution[i].Subject.Subject.Equals(Sb[j].Subject))
                    {
                        flgUnique = false;
                    }
                }

                if (flgUnique)
                {
                    UniqueSubjects++;
                    Sb.Add(mdlData.colDistribution[i].Subject);
                }
            }

            MessageBox.Show("Всего предметов в базе: " + AllSubjects.ToString() + "\n" +
                            "Уникальных предметов: " + UniqueSubjects.ToString() + "\n");
        }

        private void cmbSubjectList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSubjectList.SelectedIndex >= 0)
            {
                toolTip.SetToolTip(cmbSubjectList, cmbSubjectList.Items[cmbSubjectList.SelectedIndex].ToString());
            }
        }

        private void toolTip_Popup(object sender, PopupEventArgs e)
        {

        }

        private void cmbLecturerList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbLecturerList.SelectedIndex >= 0)
            {
                toolTip.SetToolTip(cmbLecturerList, cmbLecturerList.Items[cmbLecturerList.SelectedIndex].ToString());
            }
        }

        private void cmbKursList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbKursList.SelectedIndex >= 0)
            {
                toolTip.SetToolTip(cmbKursList, cmbKursList.Items[cmbKursList.SelectedIndex].ToString());
            }
        }

        private void cmbSpecialityList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSpecialityList.SelectedIndex >= 0)
            {
                toolTip.SetToolTip(cmbSpecialityList, cmbSpecialityList.Items[cmbSpecialityList.SelectedIndex].ToString());
            }
        }

        private void cmbSemestrList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSemestrList.SelectedIndex >= 0)
            {
                toolTip.SetToolTip(cmbSemestrList, cmbSemestrList.Items[cmbSemestrList.SelectedIndex].ToString());
            }
        }

        private void cmbLecturer3List_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbLecturer3List.SelectedIndex >= 0)
            {
                toolTip.SetToolTip(cmbLecturer3List, cmbLecturer3List.Items[cmbLecturer3List.SelectedIndex].ToString());
            }
        }

        //Занесение в почасовую нагрузку выбранной строки
        private void btnPutIntoHoured_Click(object sender, EventArgs e)
        {
            int curIndex;
            int maxCode = 0;
            IList<clsDistribution> coll = null;
            clsDistribution AddItem = new clsDistribution();
            clsDistribution SelItem = null;

            if (optMain.Checked)
            {
                coll = mdlData.colDistribution;
            }

            if (optHoured.Checked)
            {
                coll = mdlData.colHouredDistribution;
            }

            if (optCombine.Checked)
            {
                coll = mdlData.colCombineDistribution;
            }

            if (chkSubjectFilt.Checked || chkKursFilt.Checked || chkSpecialityFilt.Checked
                || chkLecturerFilt.Checked || chkSemestrFilt.Checked || chkFacultyFilt.Checked
                || chkTypeFilt.Checked)
            {
                coll = mdlData.Filtred;
            }

            //Запоминаем индекс переносимого элемента
            curIndex = cmbDistributionList.SelectedIndex;

            //Фиксируем выбранный элемент
            SelItem = coll[curIndex];

            //Копируем сведения из выбранного элемента в новый элемент
            AddItem.CopyFrom(SelItem, true);

            //Поиск максимального кода-идентификатора в коллекции
            //почасовой нагрузки
            for (int i = 0; i < mdlData.colHouredDistribution.Count; i++)
            {
                if (maxCode < mdlData.colHouredDistribution[i].Code)
                {
                    maxCode = mdlData.colHouredDistribution[i].Code;
                }
            }
            maxCode++;

            //Заменяем код у добавляемого элемента, 
            //чтобы не было конфликтов при сохранении
            AddItem.Code = maxCode;

            //Добавляемому элементу добавляем ссылку на выбранный
            AddItem.HouredConnect = SelItem;
            //Выбранному элементу добавляем ссылку на добавляемый
            SelItem.HouredConnect = AddItem;

            //Добавляем элемент в коллекцию почасовой нагрузки
            mdlData.colHouredDistribution.Add(AddItem);

            //Заполняем коллекцию штатной нагрузки с учётои почасовой нагрузки
            mdlData.toCombineDistribution(mdlData.colDistribution, mdlData.colCombineDistribution,
                                          mdlData.colHouredDistribution, true);

            //Обновляем данные по распределению
            FillDistributionList(cmbDistributionList, Selected, mdlData.cmbRem,
                                 chkSubjectFilt.Checked, 
                                 chkKursFilt.Checked, 
                                 chkSpecialityFilt.Checked,
                                 chkLecturerFilt.Checked, 
                                 chkSemestrFilt.Checked, 
                                 chkFacultyFilt.Checked,
                                 chkTypeFilt.Checked,
                                 false);
        }

        //
        private void btnDelFromHoured_Click(object sender, EventArgs e)
        {
            int curIndex;
            IList<clsDistribution> coll = null;
            clsDistribution AddItem = new clsDistribution();
            clsDistribution SelItem = null;

            if (optMain.Checked)
            {
                coll = mdlData.colDistribution;
            }

            if (optHoured.Checked)
            {
                coll = mdlData.colHouredDistribution;
            }

            if (optCombine.Checked)
            {
                coll = mdlData.colCombineDistribution;
            }

            if (chkSubjectFilt.Checked || chkKursFilt.Checked || chkSpecialityFilt.Checked
                || chkLecturerFilt.Checked || chkSemestrFilt.Checked || chkFacultyFilt.Checked
                || chkTypeFilt.Checked)
            {
                coll = mdlData.Filtred;
            }

            //Запоминаем индекс выбранного элемента
            curIndex = cmbDistributionList.SelectedIndex;
            //Фиксируем выбранный элемент
            SelItem = coll[curIndex];

            for (int i = 0; i < mdlData.colHouredDistribution.Count; i++)
            {
                if (SelItem.HouredConnect.Equals(mdlData.colHouredDistribution[i]))
                {
                    mdlData.colHouredDistribution.RemoveAt(i);
                    break;
                }
            }

            //Очищаем связь по почасовой нагрузке
            SelItem.HouredConnect = null;

            //Обновляем данные по распределению
            FillDistributionList(cmbDistributionList, Selected, mdlData.cmbRem,
                                 chkSubjectFilt.Checked,
                                 chkKursFilt.Checked,
                                 chkSpecialityFilt.Checked,
                                 chkLecturerFilt.Checked,
                                 chkSemestrFilt.Checked,
                                 chkFacultyFilt.Checked,
                                 chkTypeFilt.Checked,
                                 false);
        }

        //
        private void btnInxToKod_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
            {
                mdlData.colDistribution[i].Code = i + 1;
            }

            for (int i = 0; i <= mdlData.colHouredDistribution.Count - 1; i++)
            {
                mdlData.colHouredDistribution[i].Code = i + 1;
            }

            //--------------------------------Заполняем коллекцию штатной нагрузки с учётои почасовой нагрузки
            mdlData.toCombineDistribution(mdlData.colDistribution, mdlData.colCombineDistribution,
                                          mdlData.colHouredDistribution, true);

            //Обновляем данные по увязкам лабораторных работ
            FillDistributionList(cmbLabWorkConnect, Selected, mdlData.cmbRem, 
                                 chkSubjectFilt.Checked, 
                                 chkKursFilt.Checked, 
                                 chkSpecialityFilt.Checked,
                                 chkLecturerFilt.Checked, 
                                 chkSemestrFilt.Checked, 
                                 chkFacultyFilt.Checked,
                                 chkTypeFilt.Checked, 
                                 true);

            //Обновляем данные по распределению
            FillDistributionList(cmbDistributionList, Selected, mdlData.cmbRem,
                                 chkSubjectFilt.Checked, 
                                 chkKursFilt.Checked, 
                                 chkSpecialityFilt.Checked,
                                 chkLecturerFilt.Checked, 
                                 chkSemestrFilt.Checked, 
                                 chkFacultyFilt.Checked,
                                 chkTypeFilt.Checked, 
                                 false);
        }

        /// <summary>
        /// Метод установки для всех объектов нагрузки признака записи в заявку для диспетчерской
        /// </summary>
        private void MakeAllDispatched(bool flg)
        {
            int i;

            for (i = 0; i <= mdlData.colDistribution.Count - 1; i++)
            {
                mdlData.colDistribution[i].flgDispatch = flg;
            }

            for (i = 0; i <= mdlData.colHouredDistribution.Count - 1; i++)
            {
                mdlData.colHouredDistribution[i].flgDispatch = flg;
            }

            //--------------------------------Заполняем коллекцию штатной нагрузки с учётои почасовой нагрузки
            mdlData.toCombineDistribution(mdlData.colDistribution, mdlData.colCombineDistribution,
                                          mdlData.colHouredDistribution, true);
            //Обновляем данные по распределению
            FillDistributionList(cmbDistributionList, Selected, mdlData.cmbRem,
                                 chkSubjectFilt.Checked,
                                 chkKursFilt.Checked,
                                 chkSpecialityFilt.Checked,
                                 chkLecturerFilt.Checked,
                                 chkSemestrFilt.Checked,
                                 chkFacultyFilt.Checked,
                                 chkTypeFilt.Checked,
                                 false);
        }

        /// <summary>
        /// Метод установки для всех объектов нагрузки признака исключения из расчёта нагрузки
        /// </summary>
        private void MakeAllExcluded(bool flg)
        {
            int i;

            for (i = 0; i <= mdlData.colDistribution.Count - 1; i++)
            {
                mdlData.colDistribution[i].flgExclude = flg;
            }

            for (i = 0; i <= mdlData.colHouredDistribution.Count - 1; i++)
            {
                mdlData.colHouredDistribution[i].flgExclude = flg;
            }

            //--------------------------------Заполняем коллекцию штатной нагрузки с учётои почасовой нагрузки
            mdlData.toCombineDistribution(mdlData.colDistribution, mdlData.colCombineDistribution,
                                          mdlData.colHouredDistribution, true);
            //Обновляем данные по распределению
            FillDistributionList(cmbDistributionList, Selected, mdlData.cmbRem,
                                 chkSubjectFilt.Checked,
                                 chkKursFilt.Checked,
                                 chkSpecialityFilt.Checked,
                                 chkLecturerFilt.Checked,
                                 chkSemestrFilt.Checked,
                                 chkFacultyFilt.Checked,
                                 chkTypeFilt.Checked,
                                 false);
        }

        /// <summary>
        /// Нажатие на кнопку "Пожелания"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPreferences_Click(object sender, EventArgs e)
        {
            AddPreferences();
        }

        /// <summary>
        /// Метод перезаписи предпочтений по аудиториям от преподавателей в нагрузку
        /// </summary>
        private void AddPreferences()
        {
            int i;

            for (i = 0; i <= mdlData.colDistribution.Count - 1; i++)
            {
                if (mdlData.colDistribution[i].Subject.Preferences != "")
                {
                    mdlData.colDistribution[i].Text = mdlData.colDistribution[i].Subject.Preferences;
                }
                else
                {
                    if (mdlData.colDistribution[i].Lecturer.Preferences != "")
                    {
                        mdlData.colDistribution[i].Text = mdlData.colDistribution[i].Lecturer.Preferences;
                    }
                    else
                    {
                        mdlData.colDistribution[i].Text = "";
                    }
                }
            }

            //--------------------------------Заполняем коллекцию штатной нагрузки с учётои почасовой нагрузки
            mdlData.toCombineDistribution(mdlData.colDistribution, mdlData.colCombineDistribution,
                                          mdlData.colHouredDistribution, true);
            //Обновляем данные по распределению
            FillDistributionList(cmbDistributionList, Selected, mdlData.cmbRem,
                                 chkSubjectFilt.Checked,
                                 chkKursFilt.Checked,
                                 chkSpecialityFilt.Checked,
                                 chkLecturerFilt.Checked,
                                 chkSemestrFilt.Checked,
                                 chkFacultyFilt.Checked,
                                 chkTypeFilt.Checked,
                                 false);
        }

        private void chkTypeFilt_CheckedChanged(object sender, EventArgs e)
        {
            CheckFiltParam(chkTypeFilt, cmbTypeFilt, mdlData.inxType, ref mdlData.flgTypeFilt);
        }

        private void cmbTypeFilt_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeFiltParam(cmbTypeFilt, ref mdlData.inxType);
        }

        private void chkDispatch_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chkDistrib_CheckedChanged(object sender, EventArgs e)
        {
            cmbLecturerList.Enabled = !chkDistrib.Checked;
            lblLecturerList.Enabled = !chkDistrib.Checked;
        }

        private void txtFind_TextChanged(object sender, EventArgs e)
        {

        }

        private void chkExclude_CheckedChanged(object sender, EventArgs e)
        {
            //Если выставлен, то требуется запретить любое редактирование кроме сохранения
            //с изменением данного признака

            txtLecture.Enabled = !chkExclude.Checked;
            txtExam.Enabled = !chkExclude.Checked;
            txtCred.Enabled = !chkExclude.Checked;
            txtRef.Enabled = !chkExclude.Checked;
            txtTut.Enabled = !chkExclude.Checked;
            txtLab.Enabled = !chkExclude.Checked;
            txtPract.Enabled = !chkExclude.Checked;
            txtInd.Enabled = !chkExclude.Checked;
            txtKRAPK.Enabled = !chkExclude.Checked;
            txtKursPr.Enabled = !chkExclude.Checked;
            txtPreD.Enabled = !chkExclude.Checked;
            txtDiploma.Enabled = !chkExclude.Checked;
            txtTutPr.Enabled = !chkExclude.Checked;
            txtProd.Enabled = !chkExclude.Checked;
            txtGAK.Enabled = !chkExclude.Checked;
            txtHours.Enabled = !chkExclude.Checked;
            txtHoursZ.Enabled = !chkExclude.Checked;
            txtPostGrad.Enabled = !chkExclude.Checked;
            txtVisiting.Enabled = !chkExclude.Checked;
            txtMagistry.Enabled = !chkExclude.Checked;
            txtWeight.Enabled = !chkExclude.Checked;
            txtText.Enabled = !chkExclude.Checked;
            txtEnteredHours.Enabled = !chkExclude.Checked;
            txtEnteredHoursZ.Enabled = !chkExclude.Checked;

            chkDispatch.Enabled = !chkExclude.Checked;
            chkDistrib.Enabled = !chkExclude.Checked;

            cmbLabWorkConnect.Enabled = !chkExclude.Checked;
            cmbDoubler.Enabled = !chkExclude.Checked;
            cmbLecturerList.Enabled = !chkExclude.Checked;
            cmbLecturer2List.Enabled = !chkExclude.Checked;
            cmbLecturer3List.Enabled = !chkExclude.Checked;
            cmbSemestrList.Enabled = !chkExclude.Checked;
            cmbSpecialityList.Enabled = !chkExclude.Checked;
            cmbKursList.Enabled = !chkExclude.Checked;
            cmbSubjectList.Enabled = !chkExclude.Checked;
        }

        //Метод вывода дочерней формы
        private void toGenerateForm(Form f)
        {
            if (mdlData.flgReady)
            {
                //Делаем наследование от главной формы
                f.Owner = this;
                //Отображаем форму на экране
                f.ShowDialog();
                //Очищаем память от формы
                f = null;
            }
            else
            {
                MessageBox.Show(this, "Пожалуйста, загрузите сначала базу данных", "В доступе к функции отказано!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
