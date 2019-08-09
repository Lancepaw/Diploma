using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;

namespace DataBaseIO
{
    public partial class frmDepProtocol : Form
    {
        public frmDepProtocol()
        {
            InitializeComponent();
        }

        IList<clsQuestions> colQuestionsCurr = new List<clsQuestions>();

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmDepProtocol_Load(object sender, EventArgs e)
        {
            //Заполняем комбо-боксы
            FillLecturerCheckList();
            FillProtocolNumList();
            FillLecturerList(cmbSpeaker1);
            FillLecturerList(cmbSpeaker2);
            FillLecturerList(cmbSpeaker3);
            FillLecturerList(cmbSpeaker4);
            FillLecturerList(cmbSpeaker5);
            setMeetingDates();
            setNullSelection();
        }

        private void FillLecturerCheckList()
        {
            chkLstDepWorkers.Items.Clear();

            //Заполняем комбо-список
            //Сначала заведующий кафедрой
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                if (mdlData.colLecturer[i].Duty.Duty.Equals("Заведующий кафедрой"))
                {
                    chkLstDepWorkers.Items.Add(
                        mdlData.colLecturer[i].Duty.Short + ", " + 
                        mdlData.colLecturer[i].Degree.Short +
                        (mdlData.colLecturer[i].Duty1.Duty.Equals("-") ? " " : ", " + mdlData.colLecturer[i].Duty1.Short + " ") +
                        mdlData.SplitFIOString(mdlData.colLecturer[i].FIO, true, false));
                }
            }
            //Затем профессора
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                if (mdlData.colLecturer[i].Duty.Duty.Equals("Профессор"))
                {
                    chkLstDepWorkers.Items.Add(
                        mdlData.colLecturer[i].Degree.Short + ", " + 
                        mdlData.colLecturer[i].Duty.Short + " " +
                        mdlData.SplitFIOString(mdlData.colLecturer[i].FIO, true, false));
                }
            }
            //После доценты
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                if (mdlData.colLecturer[i].Duty.Duty.Equals("Доцент"))
                {
                    chkLstDepWorkers.Items.Add(
                        mdlData.colLecturer[i].Degree.Short + ", " +
                        mdlData.colLecturer[i].Duty.Short + " " +
                        mdlData.SplitFIOString(mdlData.colLecturer[i].FIO, true, false));
                }
            }
            //Вслед старшие преподавателии
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                if (mdlData.colLecturer[i].Duty.Duty.Equals("Старший преподаватель"))
                {
                    chkLstDepWorkers.Items.Add(
                        (mdlData.colLecturer[i].Degree.Degree.Equals("-") ? "" : mdlData.colLecturer[i].Degree.Short + " ") +
                        (mdlData.colLecturer[i].Duty.Duty.Equals("-") ? "" : mdlData.colLecturer[i].Duty.Short + " ") +
                        mdlData.SplitFIOString(mdlData.colLecturer[i].FIO, true, false));
                }
            }
            //Предпоследними ассистенты
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                if (mdlData.colLecturer[i].Duty.Duty.Equals("Ассистент"))
                {
                    chkLstDepWorkers.Items.Add(
                        (mdlData.colLecturer[i].Degree.Degree.Equals("-") ? "" : mdlData.colLecturer[i].Degree.Short + " ") +
                        (mdlData.colLecturer[i].Duty.Duty.Equals("-") ? "" : mdlData.colLecturer[i].Duty.Short + " ") +
                        mdlData.SplitFIOString(mdlData.colLecturer[i].FIO, true, false));
                }
            }
            //Напоследок прочие
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                if (!mdlData.colLecturer[i].Duty.Duty.Equals("Заведующий кафедрой") & 
                    !mdlData.colLecturer[i].Duty.Duty.Equals("Профессор") &
                    !mdlData.colLecturer[i].Duty.Duty.Equals("Доцент") &
                    !mdlData.colLecturer[i].Duty.Duty.Equals("Старший преподаватель") &
                    !mdlData.colLecturer[i].Duty.Duty.Equals("Ассистент"))
                {
                    chkLstDepWorkers.Items.Add(
                        (mdlData.colLecturer[i].Degree.Degree.Equals("-") ? "" : mdlData.colLecturer[i].Degree.Short + " ") +
                        (mdlData.colLecturer[i].Duty.Duty.Equals("-") ? "" : mdlData.colLecturer[i].Duty.Short + " ") +
                        mdlData.SplitFIOString(mdlData.colLecturer[i].FIO, true, false));
                }
            }
        }

        private void FillLecturerList(ComboBox cmb)
        {
            int NumFix;

            NumFix = cmb.SelectedIndex;
            //Очищаем список
            cmb.Items.Clear();

            cmb.Items.Add("(не выбран)");

            //Заполняем комбо-список преподавателями
            //и попутно считаем суммарную ставку
            for (int i = 0; i <= mdlData.colLecturer.Count - 1; i++)
            {
                cmb.Items.Add(mdlData.colLecturer[i].Code + ". " + mdlData.colLecturer[i].FIO);
            }

            //Возврат к тому же значению индекса, который был выставлен до нажатия
            //на кнопку "Сохранить"
            if (NumFix < 0)
            {
                cmb.SelectedIndex = 0;
            }
            else
            {
                cmb.SelectedIndex = NumFix;
            }
        }

        private void FillProtocolNumList()
        {
            int Num = 15;

            cmbProtocolNumList.Items.Clear();
            cmbProtocolNumList.Items.Add("-");
            for (int i = 1; i <= Num; i++) { cmbProtocolNumList.Items.Add(i); }
            cmbProtocolNumList.SelectedIndex = 0;
 
        }

        private void setMeetingDates()
        {
            //Очистка всех выделенных дат
            cldrMain.RemoveAllBoldedDates();

            //
            for (int i = 0; i <= mdlData.colQuestions.Count - 1; i++)
            {
                cldrMain.AddBoldedDate(mdlData.colQuestions[i].Date);
            }

            //Обновление в календаре выделенных дат
            cldrMain.UpdateBoldedDates();
        }

        private void btnWord_Click(object sender, EventArgs e)
        {
            intoWord();
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

        private void intoWordAnn()
        {
            //Задаём переменную для отсутствующего параметра
            object ObjMissing = Missing.Value;

            try
            {
                //Создаём новое Word приложение
                Word._Application ObjWord = new Word.Application();

                wordCoreAnn(ObjMissing, ObjWord);

                ObjWord.Quit();
            }
            catch
            {
                MessageBox.Show("Возможно на этом компьютере присутствует проблема совместимости с MS Word." +
                " Попробуйте установить версию 2007 и выше.");
            }
        }

        private void wordCore(object ObjMissing, Word._Application ObjWord)
        {
            string Sem;
            object ObjRange;
            int counter;

            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;
            Word._Document ObjDoc = ObjWord.Application.Documents.Add();
            ObjDoc.Activate();
            ObjWord.Visible = true;

            mdlData.WordPageDefault(ref ObjWord, ref ObjDoc, 3f, 1.5f, 0.75f, 2f);
           
            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            //Пишем текст с ведомственной принадлежностью вуза
            ObjParagraph.Range.Text = mdlData.MinistryName;
            //Сбрасываем предыдущие настройки формата
            ObjParagraph.Format.Reset();
            ObjParagraph.Range.Font.Reset();
            //Стандартное форматирование
            StandartTextFormat(ref ObjParagraph);
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Настраиваем отступ справа на величину 7,75 см
            ObjParagraph.Format.RightIndent = ObjWord.Application.CentimetersToPoints(7.75f);
            //Настраиваем одинарный межстрочный интервал
            ObjParagraph.Format.Space1();
            //Настраиваем отступ после абзаца текста в 1,5 ст. (чтобы не пробелом)
            ObjParagraph.Format.LineUnitAfter = 1.5f;
            //Размер шрифта 11 пт
            ObjParagraph.Range.Font.Size = 11;
            //Все заглавные
            ObjParagraph.Range.Font.AllCaps = 1;
            //Добавляем ещё один абзац текста (следующая строка)
            ObjParagraph.Range.InsertParagraphAfter();

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            //Пишем текст-префикс вуза
            ObjParagraph.Range.Text = mdlData.UniversityPrefName;
            //Сбрасываем предыдущие настройки формата
            ObjParagraph.Format.Reset();
            ObjParagraph.Range.Font.Reset();
            //Стандартное форматирование
            StandartTextFormat(ref ObjParagraph);
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Настраиваем отступ справа на величину 7,75 см
            ObjParagraph.Format.RightIndent = ObjWord.Application.CentimetersToPoints(7.75f);
            //Настраиваем одинарный межстрочный интервал
            ObjParagraph.Format.Space1();
            //Настраиваем отступ после абзаца текста в 1,5 ст. (чтобы не пробелом)
            ObjParagraph.Format.LineUnitAfter = 1.5f;
            //Размер шрифта 10 пт
            ObjParagraph.Range.Font.Size = 10;
            //Все заглавные
            ObjParagraph.Range.Font.AllCaps = 1;
            //Добавляем ещё один абзац текста (следующая строка)
            ObjParagraph.Range.InsertParagraphAfter();

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            //Пишем название вуза и суффикс (аббревиатуру)
            ObjParagraph.Range.Text = mdlData.UniversityName + " " + mdlData.UniversitySuffName;
            //Сбрасываем предыдущие настройки формата
            ObjParagraph.Format.Reset();
            ObjParagraph.Range.Font.Reset();
            //Стандартное форматирование
            StandartTextFormat(ref ObjParagraph);
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Настраиваем отступ справа на величину 7,75 см
            ObjParagraph.Format.RightIndent = ObjWord.Application.CentimetersToPoints(7.75f);
            //Настраиваем одинарный межстрочный интервал
            ObjParagraph.Format.Space1();
            //Настраиваем отступ после абзаца текста в 1,5 ст. (чтобы не пробелом)
            ObjParagraph.Format.LineUnitAfter = 1.5f;
            //Размер шрифта 12 пт
            ObjParagraph.Range.Font.Size = 12;
            //Все заглавные
            ObjParagraph.Range.Font.AllCaps = 1;
            //Шрифт не жирный, обычный
            ObjParagraph.Range.Font.Bold = 1;
            //Добавляем ещё один абзац текста (следующая строка)
            ObjParagraph.Range.InsertParagraphAfter();

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            //Пишем слово "Протокол"
            ObjParagraph.Range.Text = "Протокол";
            //Сбрасываем предыдущие настройки формата
            ObjParagraph.Format.Reset();
            ObjParagraph.Range.Font.Reset();
            //Стандартное форматирование
            StandartTextFormat(ref ObjParagraph);
            //Настраиваем одинарный межстрочный интервал
            ObjParagraph.Format.Space1();
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Настраиваем отступ справа на величину 7,75 см
            ObjParagraph.Format.RightIndent = ObjWord.Application.CentimetersToPoints(7.75f);
            //Все заглавные
            ObjParagraph.Range.Font.AllCaps = 1;
            //Добавляем ещё один абзац текста (следующая строка)
            ObjParagraph.Range.InsertParagraphAfter();

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            //Пишем дату заседания кафедры и номер протокола
            ObjParagraph.Range.Text = toFormProtocolDateAndNumber();
            //Сбрасываем предыдущие настройки формата
            ObjParagraph.Format.Reset();
            ObjParagraph.Range.Font.Reset();
            //Стандартное форматирование
            StandartTextFormat(ref ObjParagraph);
            //Настраиваем одинарный межстрочный интервал
            ObjParagraph.Format.Space1();
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Настраиваем отступ справа на величину 7,75 см
            ObjParagraph.Format.RightIndent = ObjWord.Application.CentimetersToPoints(7.75f);
            //Добавляем ещё один абзац текста (следующая строка)
            ObjParagraph.Range.InsertParagraphAfter();
            //Добавляем текст "Москва"
            ObjParagraph.Range.InsertAfter("Москва");
            //Настраиваем отступ после абзаца текста в 0 ст. (его быть не должно)
            ObjParagraph.Format.LineUnitAfter = 1.5f;
            //Добавляем ещё один абзац текста (следующая строка)
            ObjParagraph.Range.InsertParagraphAfter();

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            //Пишем наименование кафедры
            ObjParagraph.Range.Text = "Кафедра «" + mdlData.DepartmentName + "»";
            //Сбрасываем предыдущие настройки формата
            ObjParagraph.Format.Reset();
            ObjParagraph.Range.Font.Reset();
            //Стандартное форматирование
            StandartTextFormat(ref ObjParagraph);
            //Настраиваем одинарный межстрочный интервал
            ObjParagraph.Format.Space1();
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Настраиваем отступ справа на величину 7,75 см
            ObjParagraph.Format.RightIndent = ObjWord.Application.CentimetersToPoints(7.75f);
            //Настраиваем отступ после абзаца текста в 1,5 ст. (чтобы не пробелом)
            ObjParagraph.Format.LineUnitAfter = 1.5f;
            //Добавляем ещё один абзац текста (следующая строка)
            ObjParagraph.Range.InsertParagraphAfter();

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            //Пишем текст присутствовали
            ObjParagraph.Range.Text = "Присутствовали:";
            //Сбрасываем предыдущие настройки формата
            ObjParagraph.Format.Reset();
            ObjParagraph.Range.Font.Reset();
            //Стандартное форматирование
            StandartTextFormat(ref ObjParagraph);
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();

            //Перечисляем присутствовавших
            for (int i = 0; i <= chkLstDepWorkers.CheckedItems.Count - 1; i++)
            {
                ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                ObjParagraph.Range.Text = (i + 1).ToString() + ". " + chkLstDepWorkers.CheckedItems[i].ToString() + ",";
                //
                ObjParagraph.Format.Reset();
                ObjParagraph.Range.Font.Reset();
                //Стандартное форматирование
                StandartTextFormat(ref ObjParagraph);                
                //Выставляем полуторный интервал
                ObjParagraph.Format.Space15();
                //Добавляем ещё один абзац текста
                ObjParagraph.Range.InsertParagraphAfter();
            }

            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            
            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            //Пишем текст слушали
            ObjParagraph.Range.Text = "Слушали:";
            //Сбрасываем предыдущие настройки формата
            ObjParagraph.Format.Reset();
            ObjParagraph.Range.Font.Reset();
            //Стандартное форматирование
            StandartTextFormat(ref ObjParagraph);
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();

            //Перечисляем темы повестки дня
            counter = 1;
            for (int i = 0; i <= colQuestionsCurr.Count - 1; i++)
            {
                if (colQuestionsCurr[i].Speaker1 != null & colQuestionsCurr[i].Speaker2 != null &
                    colQuestionsCurr[i].Speaker3 != null & colQuestionsCurr[i].Speaker4 != null &
                    colQuestionsCurr[i].Speaker5 != null)
                {
                    ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                    ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                    ObjParagraph.Range.Text = counter.ToString() + ". Доклад " +
                            toFormSpeaker(colQuestionsCurr[i].Speaker1) + " о " +
                            colQuestionsCurr[i].Question.ToString();
                    //
                    ObjParagraph.Format.Reset();
                    ObjParagraph.Range.Font.Reset();
                    //Стандартное форматирование
                    StandartTextFormat(ref ObjParagraph);
                    //Добавляем ещё один абзац текста
                    ObjParagraph.Range.InsertParagraphAfter();
                    //
                    counter++;

                    ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                    ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                    ObjParagraph.Range.Text = counter.ToString() + ". Доклад " +
                            toFormSpeaker(colQuestionsCurr[i].Speaker2) + " о " +
                            colQuestionsCurr[i].Question.ToString();
                    //
                    ObjParagraph.Format.Reset();
                    ObjParagraph.Range.Font.Reset();
                    //Стандартное форматирование
                    StandartTextFormat(ref ObjParagraph);
                    //Добавляем ещё один абзац текста
                    ObjParagraph.Range.InsertParagraphAfter();
                    //
                    counter++;

                    ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                    ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                    ObjParagraph.Range.Text = counter.ToString() + ". Доклад " +
                            toFormSpeaker(colQuestionsCurr[i].Speaker3) + " о " +
                            colQuestionsCurr[i].Question.ToString();
                    //
                    ObjParagraph.Format.Reset();
                    ObjParagraph.Range.Font.Reset();
                    //Стандартное форматирование
                    StandartTextFormat(ref ObjParagraph);
                    //Добавляем ещё один абзац текста
                    ObjParagraph.Range.InsertParagraphAfter();
                    //
                    counter++;

                    ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                    ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                    ObjParagraph.Range.Text = counter.ToString() + ". Доклад " +
                            toFormSpeaker(colQuestionsCurr[i].Speaker4) + " о " +
                            colQuestionsCurr[i].Question.ToString();
                    //
                    ObjParagraph.Format.Reset();
                    ObjParagraph.Range.Font.Reset();
                    //Стандартное форматирование
                    StandartTextFormat(ref ObjParagraph);
                    //Добавляем ещё один абзац текста
                    ObjParagraph.Range.InsertParagraphAfter();
                    //
                    counter++;

                    ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                    ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                    ObjParagraph.Range.Text = counter.ToString() + ". Доклад " +
                            toFormSpeaker(colQuestionsCurr[i].Speaker5) + " о " +
                            colQuestionsCurr[i].Question.ToString();
                    //
                    ObjParagraph.Format.Reset();
                    ObjParagraph.Range.Font.Reset();
                    //Стандартное форматирование
                    StandartTextFormat(ref ObjParagraph);
                    //Добавляем ещё один абзац текста
                    ObjParagraph.Range.InsertParagraphAfter();
                    //
                    counter++;
                }
                else
                {
                    if (colQuestionsCurr[i].Speaker1 != null & colQuestionsCurr[i].Speaker2 != null &
                        colQuestionsCurr[i].Speaker3 != null & colQuestionsCurr[i].Speaker4 != null)
                    {
                        ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                        ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                        ObjParagraph.Range.Text = counter.ToString() + ". Доклад " +
                                toFormSpeaker(colQuestionsCurr[i].Speaker1) + " о " +
                                colQuestionsCurr[i].Question.ToString();
                        //
                        ObjParagraph.Format.Reset();
                        ObjParagraph.Range.Font.Reset();
                        //Стандартное форматирование
                        StandartTextFormat(ref ObjParagraph);
                        //Добавляем ещё один абзац текста
                        ObjParagraph.Range.InsertParagraphAfter();
                        //
                        counter++;

                        ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                        ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                        ObjParagraph.Range.Text = counter.ToString() + ". Доклад " +
                                toFormSpeaker(colQuestionsCurr[i].Speaker2) + " о " +
                                colQuestionsCurr[i].Question.ToString();
                        //
                        ObjParagraph.Format.Reset();
                        ObjParagraph.Range.Font.Reset();
                        //Стандартное форматирование
                        StandartTextFormat(ref ObjParagraph);
                        //Добавляем ещё один абзац текста
                        ObjParagraph.Range.InsertParagraphAfter();
                        //
                        counter++;

                        ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                        ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                        ObjParagraph.Range.Text = counter.ToString() + ". Доклад " +
                                toFormSpeaker(colQuestionsCurr[i].Speaker3) + " о " +
                                colQuestionsCurr[i].Question.ToString();
                        //
                        ObjParagraph.Format.Reset();
                        ObjParagraph.Range.Font.Reset();
                        //Стандартное форматирование
                        StandartTextFormat(ref ObjParagraph);
                        //Добавляем ещё один абзац текста
                        ObjParagraph.Range.InsertParagraphAfter();
                        //
                        counter++;

                        ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                        ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                        ObjParagraph.Range.Text = counter.ToString() + ". Доклад " +
                                toFormSpeaker(colQuestionsCurr[i].Speaker4) + " о " +
                                colQuestionsCurr[i].Question.ToString();
                        //
                        ObjParagraph.Format.Reset();
                        ObjParagraph.Range.Font.Reset();
                        //Стандартное форматирование
                        StandartTextFormat(ref ObjParagraph);
                        //Добавляем ещё один абзац текста
                        ObjParagraph.Range.InsertParagraphAfter();
                        //
                        counter++;
                    }
                    else
                    {
                        if (colQuestionsCurr[i].Speaker1 != null & colQuestionsCurr[i].Speaker2 != null &
                            colQuestionsCurr[i].Speaker3 != null)
                        {
                            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                            ObjParagraph.Range.Text = counter.ToString() + ". Доклад " +
                                    toFormSpeaker(colQuestionsCurr[i].Speaker1) + " о " +
                                    colQuestionsCurr[i].Question.ToString();
                            //
                            ObjParagraph.Format.Reset();
                            ObjParagraph.Range.Font.Reset();
                            //Стандартное форматирование
                            StandartTextFormat(ref ObjParagraph);
                            //Добавляем ещё один абзац текста
                            ObjParagraph.Range.InsertParagraphAfter();
                            //
                            counter++;

                            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                            ObjParagraph.Range.Text = counter.ToString() + ". Доклад " +
                                    toFormSpeaker(colQuestionsCurr[i].Speaker2) + " о " +
                                    colQuestionsCurr[i].Question.ToString();
                            //
                            ObjParagraph.Format.Reset();
                            ObjParagraph.Range.Font.Reset();
                            //Стандартное форматирование
                            StandartTextFormat(ref ObjParagraph);
                            //Добавляем ещё один абзац текста
                            ObjParagraph.Range.InsertParagraphAfter();
                            //
                            counter++;

                            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                            ObjParagraph.Range.Text = counter.ToString() + ". Доклад " +
                                    toFormSpeaker(colQuestionsCurr[i].Speaker3) + " о " +
                                    colQuestionsCurr[i].Question.ToString();
                            //
                            ObjParagraph.Format.Reset();
                            ObjParagraph.Range.Font.Reset();
                            //Стандартное форматирование
                            StandartTextFormat(ref ObjParagraph);
                            //Добавляем ещё один абзац текста
                            ObjParagraph.Range.InsertParagraphAfter();
                            //
                            counter++;
                        }
                        else
                        {
                            if (colQuestionsCurr[i].Speaker1 != null & colQuestionsCurr[i].Speaker2 != null)
                            {
                                ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                                ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                                ObjParagraph.Range.Text = counter.ToString() + ". Доклад " +
                                        toFormSpeaker(colQuestionsCurr[i].Speaker1) + " о " +
                                        colQuestionsCurr[i].Question.ToString();
                                //
                                ObjParagraph.Format.Reset();
                                ObjParagraph.Range.Font.Reset();
                                //Стандартное форматирование
                                StandartTextFormat(ref ObjParagraph);
                                //Добавляем ещё один абзац текста
                                ObjParagraph.Range.InsertParagraphAfter();
                                //
                                counter++;

                                ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                                ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                                ObjParagraph.Range.Text = counter.ToString() + ". Доклад " +
                                        toFormSpeaker(colQuestionsCurr[i].Speaker2) + " о " +
                                        colQuestionsCurr[i].Question.ToString();
                                //
                                ObjParagraph.Format.Reset();
                                ObjParagraph.Range.Font.Reset();
                                //Стандартное форматирование
                                StandartTextFormat(ref ObjParagraph);
                                //Добавляем ещё один абзац текста
                                ObjParagraph.Range.InsertParagraphAfter();
                                //
                                counter++;
                            }
                            else
                            {
                                if (colQuestionsCurr[i].Speaker1 != null)
                                {
                                    ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                                    ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                                    ObjParagraph.Range.Text = counter.ToString() + ". Доклад " +
                                            toFormSpeaker(colQuestionsCurr[i].Speaker1) + " о " +
                                            colQuestionsCurr[i].Question.ToString();
                                    //
                                    ObjParagraph.Format.Reset();
                                    ObjParagraph.Range.Font.Reset();
                                    //Стандартное форматирование
                                    StandartTextFormat(ref ObjParagraph);
                                    //Добавляем ещё один абзац текста
                                    ObjParagraph.Range.InsertParagraphAfter();
                                    //
                                    counter++;
                                }
                                else
                                {
                                    ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                                    ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                                    ObjParagraph.Range.Text = counter.ToString() + ". Доклад всех преподавателей о " +
                                            colQuestionsCurr[i].Question.ToString();
                                    //
                                    ObjParagraph.Format.Reset();
                                    ObjParagraph.Range.Font.Reset();
                                    //Стандартное форматирование
                                    StandartTextFormat(ref ObjParagraph);
                                    //Добавляем ещё один абзац текста
                                    ObjParagraph.Range.InsertParagraphAfter();
                                    //
                                    counter++;
                                }
                            }
                        }
                    }
                }
            }

            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "Постановили:";
            //
            ObjParagraph.Format.Reset();
            ObjParagraph.Range.Font.Reset();
            //Стандартное форматирование
            StandartTextFormat(ref ObjParagraph);
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();

            //Перечисляем присутствовавших
            for (int i = 1; i <= counter - 1; i++)
            {
                ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                ObjParagraph.Range.Text = i.ToString() + ". Принять к сведению и исполнению.";
                //
                ObjParagraph.Format.Reset();
                ObjParagraph.Range.Font.Reset();
                //Стандартное форматирование
                StandartTextFormat(ref ObjParagraph);
                //Добавляем ещё один абзац текста
                ObjParagraph.Range.InsertParagraphAfter();
            }

            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "Заведующий кафедрой УиЗИ, д.т.н., проф. \tЛ.А. Баранов";
            //
            ObjParagraph.Format.Reset();
            ObjParagraph.Range.Font.Reset();
            //Стандартное форматирование
            StandartTextFormat(ref ObjParagraph);
            //Вставляем каретку с выравниванием по правому краю на уровне 17f
            ObjParagraph.TabStops.Add(17f / 0.03527f, Word.WdTabAlignment.wdAlignTabRight);
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "Секретарь заседания, к.т.н., доц. \tА.И. Сафронов";
            //
            ObjParagraph.Format.Reset();
            ObjParagraph.Range.Font.Reset();
            //Стандартное форматирование
            StandartTextFormat(ref ObjParagraph);
            //Вставляем каретку с выравниванием по правому краю на уровне 17f
            ObjParagraph.TabStops.Add(17f / 0.03527f, Word.WdTabAlignment.wdAlignTabRight);
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();

            if (cldrMain.SelectionStart.Month == 1 || cldrMain.SelectionStart.Month >= 8)
            {
                Sem = "I сем";
            }
            else
            {
                Sem = "II сем";
            }  

            ObjDoc.SaveAs(Application.StartupPath + @"\"
                + "Заседание кафедры " + Sem + " №" + cmbProtocolNumList.Items[cmbProtocolNumList.SelectedIndex] + 
                " протокол " + 
                DateTime.Now.Date.ToString("yyyyMMdd") + " " + 
                DateTime.Now.TimeOfDay.ToString("hhmmss") + ".docx");

            ObjDoc.Close();
        }

        private string toFormProtocolDateAndNumber()
        {
            string Text;
            //Формируем строку текста с датой заседания кафедры
            Text = "от «" + cldrMain.SelectionStart.Day.ToString() + "» ";
            switch (cldrMain.SelectionStart.Month)
            {
                case 1:
                    Text += "января ";
                    break;
                case 2:
                    Text += "февраля ";
                    break;
                case 3:
                    Text += "марта ";
                    break;
                case 4:
                    Text += "апреля ";
                    break;
                case 5:
                    Text += "мая ";
                    break;
                case 6:
                    Text += "июня ";
                    break;
                case 7:
                    Text += "июля ";
                    break;
                case 8:
                    Text += "августа ";
                    break;
                case 9:
                    Text += "сентября ";
                    break;
                case 10:
                    Text += "октября ";
                    break;
                case 11:
                    Text += "ноября ";
                    break;
                case 12:
                    Text += "декабря ";
                    break;
            }

            Text += cldrMain.SelectionStart.Year.ToString() + " года ";

            Text += "№ " + cmbProtocolNumList.Items[cmbProtocolNumList.SelectedIndex];

            return Text;
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

        private void wordCoreAnn(object ObjMissing, Word._Application ObjWord)
        {
            string Sem = "";
            object ObjRange;
            //Задаём закладку конца документа 
            object EndOfDoc = "\\endofdoc";
            Word.Paragraph ObjParagraph;
            Word._Document ObjDoc = ObjWord.Application.Documents.Add();
            ObjDoc.Activate();
            ObjWord.Visible = true;

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);

            ObjParagraph.Range.Text = "Заседание кафедры #" + cmbProtocolNumList.Items[cmbProtocolNumList.SelectedIndex];
            //Размер шрифта 20 пт
            ObjParagraph.Range.Font.Size = 20;
            //Times New Roman
            ObjParagraph.Range.Font.Name = "Times New Roman";
            //Жирный шрифт
            ObjParagraph.Range.Font.Bold = 1;
            //Отступа слева нет
            ObjParagraph.Format.LeftIndent = 0;
            //Выравнивание по центру
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertAfter("состоится " + mdlData.getDoWString(cldrMain.SelectionStart.DayOfWeek) + ",");
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertAfter(cldrMain.SelectionStart.Day + " " +
                mdlData.getMonthStringRP(cldrMain.SelectionStart.Month) + " " + cldrMain.SelectionStart.Year + " г. в " +
                txtTime.Text);
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertAfter("в аудитории " + txtRoom.Text);
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);

            //Размер шрифта 16 пт
            ObjParagraph.Range.Font.Size = 16;
            //Обычный шрифт
            ObjParagraph.Range.Font.Bold = 0;
            //Добавляем ещё один абзац текста вперёд
            ObjParagraph.Range.InsertParagraphBefore();
            //Добавляем ещё один абзац текста вперёд
            ObjParagraph.Range.InsertParagraphBefore();
            //Добавляем ещё один абзац текста вперёд
            ObjParagraph.Range.InsertParagraphBefore();
            ObjParagraph.Range.Text = "Повестка дня:";
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();

            for (int i = 0; i <= mdlData.colQuestions.Count - 1; i++)
            {
                if (mdlData.colQuestions[i].Date.Day.Equals(cldrMain.SelectionStart.Day) &
                    mdlData.colQuestions[i].Date.Month.Equals(cldrMain.SelectionStart.Month) &
                    mdlData.colQuestions[i].Date.Year.Equals(cldrMain.SelectionStart.Year))
                {
                    ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                    ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
                    ObjParagraph.Range.Text = mdlData.colQuestions[i].Question;
                    ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    ObjParagraph.Range.ListFormat.ApplyNumberDefault(ObjMissing);

                    //Размер шрифта 16 пт
                    ObjParagraph.Range.Font.Size = 16;
                    //Обычный шрифт
                    ObjParagraph.Range.Font.Bold = 0;
                    ObjParagraph.LeftIndent = 0.63f / 0.03527f;
                    ObjParagraph.SpaceAfter = 10;
                    //Добавляем ещё один абзац текста
                    ObjParagraph.Range.InsertParagraphAfter();

                    ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
                    ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);

                    if (!(mdlData.colQuestions[i].Speaker1 == null) & !(mdlData.colQuestions[i].Speaker2 == null) &
                        !(mdlData.colQuestions[i].Speaker3 == null) & !(mdlData.colQuestions[i].Speaker4 == null) &
                        !(mdlData.colQuestions[i].Speaker5 == null))
                    {
                        ObjParagraph.Range.Text = "Докл. " + mdlData.SplitFIOString(mdlData.colQuestions[i].Speaker1.FIO, true, false) + ", " +
                            mdlData.SplitFIOString(mdlData.colQuestions[i].Speaker2.FIO, true, false) + ", " +
                            mdlData.SplitFIOString(mdlData.colQuestions[i].Speaker3.FIO, true, false) + ", " +
                            mdlData.SplitFIOString(mdlData.colQuestions[i].Speaker4.FIO, true, false) + ", " +
                            mdlData.SplitFIOString(mdlData.colQuestions[i].Speaker5.FIO, true, false);
                    }
                    else
                    {
                        if (!(mdlData.colQuestions[i].Speaker1 == null) & !(mdlData.colQuestions[i].Speaker2 == null) &
                            !(mdlData.colQuestions[i].Speaker3 == null) & !(mdlData.colQuestions[i].Speaker4 == null))
                        {
                            ObjParagraph.Range.Text = "Докл. " + mdlData.SplitFIOString(mdlData.colQuestions[i].Speaker1.FIO, true, false) + ", " +
                                mdlData.SplitFIOString(mdlData.colQuestions[i].Speaker2.FIO, true, false) + ", " +
                                mdlData.SplitFIOString(mdlData.colQuestions[i].Speaker3.FIO, true, false) + ", " +
                                mdlData.SplitFIOString(mdlData.colQuestions[i].Speaker4.FIO, true, false);
                        }
                        else
                        {
                            if (!(mdlData.colQuestions[i].Speaker1 == null) & !(mdlData.colQuestions[i].Speaker2 == null) &
                                !(mdlData.colQuestions[i].Speaker3 == null))
                            {
                                ObjParagraph.Range.Text = "Докл. " + mdlData.SplitFIOString(mdlData.colQuestions[i].Speaker1.FIO, true, false) + ", " +
                                    mdlData.SplitFIOString(mdlData.colQuestions[i].Speaker2.FIO, true, false) + ", " +
                                    mdlData.SplitFIOString(mdlData.colQuestions[i].Speaker3.FIO, true, false);
                            }
                            else
                            {
                                if (!(mdlData.colQuestions[i].Speaker1 == null) & !(mdlData.colQuestions[i].Speaker2 == null))
                                {
                                    ObjParagraph.Range.Text = "Докл. " + mdlData.SplitFIOString(mdlData.colQuestions[i].Speaker1.FIO, true, false) + ", " +
                                        mdlData.SplitFIOString(mdlData.colQuestions[i].Speaker2.FIO, true, false);
                                }
                                else
                                {
                                    if (!(mdlData.colQuestions[i].Speaker1 == null))
                                    {
                                        ObjParagraph.Range.Text = "Докл. " + mdlData.SplitFIOString(mdlData.colQuestions[i].Speaker1.FIO, true, false);
                                    }
                                    else
                                    {
                                        ObjParagraph.Range.Text = "Все преподаватели";
                                    }
                                }
                            }
                        }
                    }
                    
                    //Размер шрифта 16 пт
                    ObjParagraph.Range.Font.Size = 16;
                    //Обычный шрифт
                    ObjParagraph.Range.Font.Bold = 0;
                    ObjParagraph.Range.ListFormat.ApplyNumberDefault(Word.WdListNumberStyle.wdListNumberStyleNone);
                    ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    ObjParagraph.LeftIndent = 7f / 0.03527f;
                    ObjParagraph.SpaceAfter = 10;
                    //Добавляем ещё один абзац текста
                    ObjParagraph.Range.InsertParagraphAfter();
                }
            }

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);
            ObjParagraph.Range.Text = "Разное.";
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjParagraph.Range.ListFormat.ApplyNumberDefault(ObjMissing);

            //Размер шрифта 16 пт
            ObjParagraph.Range.Font.Size = 16;
            //Обычный шрифт
            ObjParagraph.Range.Font.Bold = 0;
            ObjParagraph.LeftIndent = 0.63f / 0.03527f;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();

            ObjRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjParagraph = ObjDoc.Content.Paragraphs.Add(ref ObjRange);

            ObjParagraph.Range.Text = "Все преподаватели";

            //Размер шрифта 16 пт
            ObjParagraph.Range.Font.Size = 16;
            //Обычный шрифт
            ObjParagraph.Range.Font.Bold = 0;
            ObjParagraph.Range.ListFormat.ApplyNumberDefault(Word.WdListNumberStyle.wdListNumberStyleNone);
            ObjParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            ObjParagraph.LeftIndent = 7f / 0.03527f;
            //Добавляем ещё один абзац текста
            ObjParagraph.Range.InsertParagraphAfter();

            if (cldrMain.SelectionStart.Month == 1 || cldrMain.SelectionStart.Month >= 8)
            { Sem = "I сем"; }
            else { Sem = "II сем"; }

            ObjDoc.SaveAs(Application.StartupPath + @"\"
                + "Заседание кафедры " + Sem + " №" + cmbProtocolNumList.Items[cmbProtocolNumList.SelectedIndex] + " " +
                DateTime.Now.Date.ToString("yyyyMMdd") + " " + 
                DateTime.Now.TimeOfDay.ToString("hhmmss") + ".docx");

            ObjDoc.Close();

        }

        //
        private void cmbProtocolNumList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        //
        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= chkLstDepWorkers.Items.Count - 1; i++)
            {
                chkLstDepWorkers.SetItemChecked(i, true);
            }
        }

        //
        private void btnClear_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= chkLstDepWorkers.Items.Count - 1; i++)
            {
                chkLstDepWorkers.SetItemChecked(i, false);
            }
        }

        //
        private void btnAnnounce_Click(object sender, EventArgs e)
        {
            intoWordAnn();
        }

        //
        private void cldrMain_DateChanged(object sender, DateRangeEventArgs e)
        {
            setCurrentQuestions();
            setCurrentProtocolNum();
            setNullSelection();
        }

        //
        private void setNullSelection()
        {
            //
            if (cldrMain.SelectionStart == null)
            {
                btnAdd.Enabled = false;
            }
            else
            {
                btnAdd.Enabled = true;
            }

            //
            if (lstQuestions.Items.Count > 0)
            {
                btnCopy.Enabled = true;
                btnCopyAll.Enabled = true;
                btnChange.Enabled = true;
                btnClear.Enabled = true;
                btnClearAll.Enabled = true;
                btnWord.Enabled = true;
                btnAnnounce.Enabled = true;
            }
            else
            {
                btnCopy.Enabled = false;
                btnCopyAll.Enabled = false;
                btnChange.Enabled = false;
                btnClear.Enabled = false;
                btnClearAll.Enabled = false;
                btnWord.Enabled = false;
                btnAnnounce.Enabled = false;
            }

            txtQuestion.Text = "";
            cmbSpeaker1.SelectedIndex = 0;
            cmbSpeaker2.SelectedIndex = 0;
            cmbSpeaker3.SelectedIndex = 0;
            cmbSpeaker4.SelectedIndex = 0;
            cmbSpeaker5.SelectedIndex = 0;

            txtQuestion.Enabled = false;
            cmbSpeaker1.Enabled = false;
            cmbSpeaker2.Enabled = false;
            cmbSpeaker3.Enabled = false;
            cmbSpeaker4.Enabled = false;
            cmbSpeaker5.Enabled = false;

            btnDel.Enabled = false;
            btnChange.Enabled = false;
        }

        private string toFormSpeaker(clsLecturer L)
        {
            string str = "";

            str = (L.Degree.Degree.Equals("-") ? "" : L.Degree.Short) +
                  (L.Duty.Duty.Equals("-") ? "" : (L.Degree.Degree.Equals("-") ? "" : ", ") + L.Duty.Short + " ") +
                  mdlData.SplitFIOString(L.FIO, true, false);

            return str;
        }

        //
        private void setCurrentQuestions()
        {
            lstQuestions.Items.Clear();
            colQuestionsCurr.Clear();

            for (int i = 0; i <= mdlData.colQuestions.Count - 1; i++)
            {
                if (mdlData.colQuestions[i].Date.Day.Equals(cldrMain.SelectionStart.Day) &
                    mdlData.colQuestions[i].Date.Month.Equals(cldrMain.SelectionStart.Month) &
                    mdlData.colQuestions[i].Date.Year.Equals(cldrMain.SelectionStart.Year))
                {
                    lstQuestions.Items.Add(mdlData.colQuestions[i].Question);
                    colQuestionsCurr.Add(mdlData.colQuestions[i]);
                }
            }
        }

        /// <summary>
        /// Процедура выставления текущего номера протокола по
        /// выбранной в календаре дате
        /// </summary>
        private void setCurrentProtocolNum()
        {
            int i;
            int count = 0;
            bool flgNeedCount = false;

            //Перебираем все выделенные даты
            for (i = 0; i <= cldrMain.BoldedDates.Length - 1; i++)
            {
                //Если выбранная дата одна из выделенных дат
                if (cldrMain.SelectionStart.Equals(cldrMain.BoldedDates.GetValue(i)))
                {
                    //выставляем признак необходимости счёта заседаний кафедры
                    flgNeedCount = true;
                    //прерываем цикл
                    break;
                }
            }

            //Если известно, что считать заседания кафедры необходимо, то
            if (flgNeedCount)
            {
                //Запускаем ещё один цикл по выделенным датам
                for (i = 0; i <= cldrMain.BoldedDates.Length - 1; i++)
                {
                    //Если год текущий - это наш случай
                    if (cldrMain.BoldedDates[i].Year == cldrMain.SelectionStart.Year)
                    {
                        //Если месяц
                        if (cldrMain.BoldedDates[i].Month == cldrMain.SelectionStart.Month)
                        {
                            if (cldrMain.BoldedDates[i].Day <= cldrMain.SelectionStart.Day)
                            {
                                count++;
                            }
                        }
                        else
                        {
                            if (cldrMain.BoldedDates[i].Month < cldrMain.SelectionStart.Month)
                            {
                                count++;
                            }
                        }
                    }
                    else
                    {
                        //Если год один из предыдущих - это тоже подходит
                        //(и пока вне зависимости от каких-либо условий)
                        if (cldrMain.BoldedDates[i].Year < cldrMain.SelectionStart.Year)
                        {
                            count++;
                        }
                    }
                }
            }
            
            //Выставляем позицию комбинированного списка
            //на значение, равное количеству заседаний кафедры
            cmbProtocolNumList.SelectedIndex = count;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            clsQuestions Q = new clsQuestions();

            Q.Code = mdlData.colQuestions.Count + 1;
            Q.Date = cldrMain.SelectionStart.Date;
            Q.Question = "Новый вопрос";
            mdlData.colQuestions.Add(Q);
            setCurrentQuestions();
            setMeetingDates();
            setNullSelection();
        }

        //
        private void lstQuestions_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstQuestions.SelectedIndex >= 0)
            {
                txtQuestion.Enabled = true;
                btnDel.Enabled = true;
                btnChange.Enabled = true;

                cmbSpeaker1.Enabled = true;
                cmbSpeaker2.Enabled = true;
                cmbSpeaker3.Enabled = true;
                cmbSpeaker4.Enabled = true;
                cmbSpeaker5.Enabled = true;

                txtQuestion.Text = colQuestionsCurr[lstQuestions.SelectedIndex].Question;

                if (colQuestionsCurr[lstQuestions.SelectedIndex].Speaker1 != null)
                {
                    cmbSpeaker1.SelectedIndex = colQuestionsCurr[lstQuestions.SelectedIndex].Speaker1.Code;
                }
                else
                {
                    cmbSpeaker1.SelectedIndex = 0;
                }

                if (colQuestionsCurr[lstQuestions.SelectedIndex].Speaker2 != null)
                {
                    cmbSpeaker2.SelectedIndex = colQuestionsCurr[lstQuestions.SelectedIndex].Speaker2.Code;
                }
                else
                {
                    cmbSpeaker2.SelectedIndex = 0;
                }

                if (colQuestionsCurr[lstQuestions.SelectedIndex].Speaker3 != null)
                {
                    cmbSpeaker3.SelectedIndex = colQuestionsCurr[lstQuestions.SelectedIndex].Speaker3.Code;
                }
                else
                {
                    cmbSpeaker3.SelectedIndex = 0;
                }

                if (colQuestionsCurr[lstQuestions.SelectedIndex].Speaker4 != null)
                {
                    cmbSpeaker4.SelectedIndex = colQuestionsCurr[lstQuestions.SelectedIndex].Speaker4.Code;
                }
                else
                {
                    cmbSpeaker4.SelectedIndex = 0;
                }

                if (colQuestionsCurr[lstQuestions.SelectedIndex].Speaker5 != null)
                {
                    cmbSpeaker5.SelectedIndex = colQuestionsCurr[lstQuestions.SelectedIndex].Speaker5.Code;
                }
                else
                {
                    cmbSpeaker5.SelectedIndex = 0;
                }
            }
            else
            {
                txtQuestion.Enabled = false;
                btnDel.Enabled = false;
                btnChange.Enabled = false;

                cmbSpeaker1.Enabled = false;
                cmbSpeaker2.Enabled = false;
                cmbSpeaker3.Enabled = false;
                cmbSpeaker4.Enabled = false;
                cmbSpeaker5.Enabled = false;
            }
        }

        private void lblSpeaker3_Click(object sender, EventArgs e)
        {
            
        }

        private void btnDel_Click(object sender, EventArgs e)
        {

        }

        private void btnChange_Click(object sender, EventArgs e)
        {
            colQuestionsCurr[lstQuestions.SelectedIndex].Question = txtQuestion.Text;

            if (cmbSpeaker1.SelectedIndex > 0)
            {
                colQuestionsCurr[lstQuestions.SelectedIndex].Speaker1 = mdlData.colLecturer[cmbSpeaker1.SelectedIndex - 1];
            }
            else
            {
                colQuestionsCurr[lstQuestions.SelectedIndex].Speaker1 = null;
            }

            if (cmbSpeaker2.SelectedIndex > 0)
            {
                colQuestionsCurr[lstQuestions.SelectedIndex].Speaker2 = mdlData.colLecturer[cmbSpeaker2.SelectedIndex - 1];
            }
            else
            {
                colQuestionsCurr[lstQuestions.SelectedIndex].Speaker2 = null;
            }

            if (cmbSpeaker3.SelectedIndex > 0)
            {
                colQuestionsCurr[lstQuestions.SelectedIndex].Speaker3 = mdlData.colLecturer[cmbSpeaker3.SelectedIndex - 1];
            }
            else
            {
                colQuestionsCurr[lstQuestions.SelectedIndex].Speaker3 = null;
            }

            if (cmbSpeaker4.SelectedIndex > 0)
            {
                colQuestionsCurr[lstQuestions.SelectedIndex].Speaker4 = mdlData.colLecturer[cmbSpeaker4.SelectedIndex - 1];
            }
            else
            {
                colQuestionsCurr[lstQuestions.SelectedIndex].Speaker4 = null;
            }

            if (cmbSpeaker5.SelectedIndex > 0)
            {
                colQuestionsCurr[lstQuestions.SelectedIndex].Speaker5 = mdlData.colLecturer[cmbSpeaker5.SelectedIndex - 1];
            }
            else
            {
                colQuestionsCurr[lstQuestions.SelectedIndex].Speaker5 = null;
            }

            setCurrentQuestions();
            setMeetingDates();
            setNullSelection();
        }
    }
}
