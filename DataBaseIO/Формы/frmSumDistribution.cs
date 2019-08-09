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
    public partial class frmSumDistribution : Form
    {
        public frmSumDistribution()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AdditionFunction(clsDistribution Dis, ref int Lk, 
            ref int Ex, ref int Cr, ref int Rf, ref int Tut, ref int Lb,
            ref int Pr, ref int Ind, ref int KRAPK, ref int Kp, ref int PD,
            ref int Dp, ref int TP, ref int PP, ref int GAK, ref int Bud,
            ref float BudZ, ref int Sum, ref float SumZ, ref int All,
            ref float AllZ)
        {
            Lk += Dis.Lecture;
            Ex += Dis.Exam;
            Cr += Dis.Credit;
            Rf += Dis.RefHomeWork;
            Tut += Dis.Tutorial;
            Lb += Dis.LabWork;
            Pr += Dis.Practice;
            Ind += Dis.IndividualWork;
            KRAPK += Dis.KRAPK;
            Kp += Dis.KursProject;
            PD += Dis.PreDiplomaPractice;
            Dp += Dis.DiplomaPaper;
            TP += Dis.TutorialPractice;
            PP += Dis.ProducingPractice;
            GAK += Dis.GAK;
            Bud += Dis.Hours;
            BudZ += Dis.HoursZ;
            Sum += Dis.EnteredHours;
            SumZ += Dis.EnteredHoursZ;
            All += mdlData.toSumDistributionComponents(Dis);
            AllZ += Dis.EnteredHoursZ;
        }

        private void frmSumDistribution_Load(object sender, EventArgs e)
        {
            clsSummary ITTSU1 = new clsSummary();
            clsSummary ITTSU2 = new clsSummary();

            mdlData.colSummary.Clear();
            //--------------------------Первый столбец значений
            //int ITTSU1LectCounter = 0;

            int ITTSU1LectBCounter = 0;
            int ITTSU1LectMCounter = 0;
            int ITTSU1LectSCounter = 0;
            int ITTSU1LectSURCounter = 0;
            int ITTSU1LectUOCounter = 0;
            int ITTSU1LectUMOCounter = 0;

            int ITTSU2LectCounter = 0;
            int ITTSU2LectBCounter = 0;
            int ITTSU2LectMCounter = 0;
            int ITTSU2LectSCounter = 0;
            int ITTSU2LectSURCounter = 0;
            int ITTSU2LectUOCounter = 0;
            int ITTSU2LectUMOCounter = 0;

            int IUIT1LectCounter = 0;
            int IUIT2LectCounter = 0;
            
            int VF1LectCounter = 0;
            int VF2LectCounter = 0;

            int Sum1LectCounter = 0;
            int Sum2LectCounter = 0;
            
            int AllLectCounter = 0;
            //--------------------------Первый столбец значений

            //--------------------------Второй столбец значений
            
            //int ITTSU1ExamCounter = 0;
            int ITTSU1ExamBCounter = 0;
            int ITTSU1ExamMCounter = 0;
            int ITTSU1ExamSCounter = 0;
            int ITTSU1ExamSURCounter = 0;
            int ITTSU1ExamUOCounter = 0;
            int ITTSU1ExamUMOCounter = 0;

            int ITTSU2ExamCounter = 0;
            int ITTSU2ExamBCounter = 0;
            int ITTSU2ExamMCounter = 0;
            int ITTSU2ExamSCounter = 0;
            int ITTSU2ExamSURCounter = 0;
            int ITTSU2ExamUOCounter = 0;
            int ITTSU2ExamUMOCounter = 0;

            int IUIT1ExamCounter = 0;
            int IUIT2ExamCounter = 0;

            int VF1ExamCounter = 0;
            int VF2ExamCounter = 0;

            int Sum1ExamCounter = 0;
            int Sum2ExamCounter = 0;

            int AllExamCounter = 0;
            //--------------------------Второй столбец значений

            //--------------------------Третий столбец значений
            
            //int ITTSU1CredCounter = 0;
            int ITTSU1CredBCounter = 0;
            int ITTSU1CredMCounter = 0;
            int ITTSU1CredSCounter = 0;
            int ITTSU1CredSURCounter = 0;
            int ITTSU1CredUOCounter = 0;
            int ITTSU1CredUMOCounter = 0;

            int ITTSU2CredCounter = 0;
            int ITTSU2CredBCounter = 0;
            int ITTSU2CredMCounter = 0;
            int ITTSU2CredSCounter = 0;
            int ITTSU2CredSURCounter = 0;
            int ITTSU2CredUOCounter = 0;
            int ITTSU2CredUMOCounter = 0;

            int IUIT1CredCounter = 0;
            int IUIT2CredCounter = 0;

            int VF1CredCounter = 0;
            int VF2CredCounter = 0;

            int Sum1CredCounter = 0;
            int Sum2CredCounter = 0;

            int AllCredCounter = 0;
            //--------------------------Третий столбец значений

            //--------------------------Четвёртый столбец значений
            
            //int ITTSU1RefCounter = 0;
            int ITTSU1RefBCounter = 0;
            int ITTSU1RefMCounter = 0;
            int ITTSU1RefSCounter = 0;
            int ITTSU1RefSURCounter = 0;
            int ITTSU1RefUOCounter = 0;
            int ITTSU1RefUMOCounter = 0;

            int ITTSU2RefCounter = 0;
            int ITTSU2RefBCounter = 0;
            int ITTSU2RefMCounter = 0;
            int ITTSU2RefSCounter = 0;
            int ITTSU2RefSURCounter = 0;
            int ITTSU2RefUOCounter = 0;
            int ITTSU2RefUMOCounter = 0;

            int IUIT1RefCounter = 0;
            int IUIT2RefCounter = 0;

            int VF1RefCounter = 0;
            int VF2RefCounter = 0;

            int Sum1RefCounter = 0;
            int Sum2RefCounter = 0;

            int AllRefCounter = 0;
            //--------------------------Четвёртый столбец значений

            //--------------------------Пятый столбец значений
            
            //int ITTSU1TutCounter = 0;
            int ITTSU1TutBCounter = 0;
            int ITTSU1TutMCounter = 0;
            int ITTSU1TutSCounter = 0;
            int ITTSU1TutSURCounter = 0;
            int ITTSU1TutUOCounter = 0;
            int ITTSU1TutUMOCounter = 0;

            int ITTSU2TutCounter = 0;
            int ITTSU2TutBCounter = 0;
            int ITTSU2TutMCounter = 0;
            int ITTSU2TutSCounter = 0;
            int ITTSU2TutSURCounter = 0;
            int ITTSU2TutUOCounter = 0;
            int ITTSU2TutUMOCounter = 0;

            int IUIT1TutCounter = 0;
            int IUIT2TutCounter = 0;

            int VF1TutCounter = 0;
            int VF2TutCounter = 0;

            int Sum1TutCounter = 0;
            int Sum2TutCounter = 0;

            int AllTutCounter = 0;
            //--------------------------Пятый столбец значений

            //--------------------------Шестой столбец значений
            
            //int ITTSU1LabCounter = 0;
            int ITTSU1LabBCounter = 0;
            int ITTSU1LabMCounter = 0;
            int ITTSU1LabSCounter = 0;
            int ITTSU1LabSURCounter = 0;
            int ITTSU1LabUOCounter = 0;
            int ITTSU1LabUMOCounter = 0;

            int ITTSU2LabCounter = 0;
            int ITTSU2LabBCounter = 0;
            int ITTSU2LabMCounter = 0;
            int ITTSU2LabSCounter = 0;
            int ITTSU2LabSURCounter = 0;
            int ITTSU2LabUOCounter = 0;
            int ITTSU2LabUMOCounter = 0;

            int IUIT1LabCounter = 0;
            int IUIT2LabCounter = 0;

            int VF1LabCounter = 0;
            int VF2LabCounter = 0;

            int Sum1LabCounter = 0;
            int Sum2LabCounter = 0;

            int AllLabCounter = 0;
            //--------------------------Шестой столбец значений

            //--------------------------Седьмой столбец значений
            
            //int ITTSU1PractCounter = 0;
            int ITTSU1PractBCounter = 0;
            int ITTSU1PractMCounter = 0;
            int ITTSU1PractSCounter = 0;
            int ITTSU1PractSURCounter = 0;
            int ITTSU1PractUOCounter = 0;
            int ITTSU1PractUMOCounter = 0;

            int ITTSU2PractCounter = 0;
            int ITTSU2PractBCounter = 0;
            int ITTSU2PractMCounter = 0;
            int ITTSU2PractSCounter = 0;
            int ITTSU2PractSURCounter = 0;
            int ITTSU2PractUOCounter = 0;
            int ITTSU2PractUMOCounter = 0;

            int IUIT1PractCounter = 0;
            int IUIT2PractCounter = 0;

            int VF1PractCounter = 0;
            int VF2PractCounter = 0;

            int Sum1PractCounter = 0;
            int Sum2PractCounter = 0;

            int AllPractCounter = 0;
            //--------------------------Седьмой столбец значений

            //--------------------------Восьмой столбец значений
            
            //int ITTSU1IndCounter = 0;
            int ITTSU1IndBCounter = 0;
            int ITTSU1IndMCounter = 0;
            int ITTSU1IndSCounter = 0;
            int ITTSU1IndSURCounter = 0;
            int ITTSU1IndUOCounter = 0;
            int ITTSU1IndUMOCounter = 0;

            int ITTSU2IndCounter = 0;
            int ITTSU2IndBCounter = 0;
            int ITTSU2IndMCounter = 0;
            int ITTSU2IndSCounter = 0;
            int ITTSU2IndSURCounter = 0;
            int ITTSU2IndUOCounter = 0;
            int ITTSU2IndUMOCounter = 0;

            int IUIT1IndCounter = 0;
            int IUIT2IndCounter = 0;

            int VF1IndCounter = 0;
            int VF2IndCounter = 0;

            int Sum1IndCounter = 0;
            int Sum2IndCounter = 0;

            int AllIndCounter = 0;
            //--------------------------Восьмой столбец значений

            //--------------------------Девятый столбец значений
            
            //int ITTSU1KRAPKCounter = 0;
            int ITTSU1KRAPKBCounter = 0;
            int ITTSU1KRAPKMCounter = 0;
            int ITTSU1KRAPKSCounter = 0;
            int ITTSU1KRAPKSURCounter = 0;
            int ITTSU1KRAPKUOCounter = 0;
            int ITTSU1KRAPKUMOCounter = 0;

            int ITTSU2KRAPKCounter = 0;
            int ITTSU2KRAPKBCounter = 0;
            int ITTSU2KRAPKMCounter = 0;
            int ITTSU2KRAPKSCounter = 0;
            int ITTSU2KRAPKSURCounter = 0;
            int ITTSU2KRAPKUOCounter = 0;
            int ITTSU2KRAPKUMOCounter = 0;

            int IUIT1KRAPKCounter = 0;
            int IUIT2KRAPKCounter = 0;

            int VF1KRAPKCounter = 0;
            int VF2KRAPKCounter = 0;

            int Sum1KRAPKCounter = 0;
            int Sum2KRAPKCounter = 0;

            int AllKRAPKCounter = 0;
            //--------------------------Девятый столбец значений

            //--------------------------Десятый столбец значений
            
            //int ITTSU1KursCounter = 0;
            int ITTSU1KursBCounter = 0;
            int ITTSU1KursMCounter = 0;
            int ITTSU1KursSCounter = 0;
            int ITTSU1KursSURCounter = 0;
            int ITTSU1KursUOCounter = 0;
            int ITTSU1KursUMOCounter = 0;

            int ITTSU2KursCounter = 0;
            int ITTSU2KursBCounter = 0;
            int ITTSU2KursMCounter = 0;
            int ITTSU2KursSCounter = 0;
            int ITTSU2KursSURCounter = 0;
            int ITTSU2KursUOCounter = 0;
            int ITTSU2KursUMOCounter = 0;

            int IUIT1KursCounter = 0;
            int IUIT2KursCounter = 0;

            int VF1KursCounter = 0;
            int VF2KursCounter = 0;

            int Sum1KursCounter = 0;
            int Sum2KursCounter = 0;

            int AllKursCounter = 0;
            //--------------------------Десятый столбец значений

            //--------------------------Одиннадцатый столбец значений
            
            //int ITTSU1PreDCounter = 0;
            int ITTSU1PreDBCounter = 0;
            int ITTSU1PreDMCounter = 0;
            int ITTSU1PreDSCounter = 0;
            int ITTSU1PreDSURCounter = 0;
            int ITTSU1PreDUOCounter = 0;
            int ITTSU1PreDUMOCounter = 0;

            int ITTSU2PreDCounter = 0;
            int ITTSU2PreDBCounter = 0;
            int ITTSU2PreDMCounter = 0;
            int ITTSU2PreDSCounter = 0;
            int ITTSU2PreDSURCounter = 0;
            int ITTSU2PreDUOCounter = 0;
            int ITTSU2PreDUMOCounter = 0;

            int IUIT1PreDCounter = 0;
            int IUIT2PreDCounter = 0;

            int VF1PreDCounter = 0;
            int VF2PreDCounter = 0;

            int Sum1PreDCounter = 0;
            int Sum2PreDCounter = 0;

            int AllPreDCounter = 0;
            //--------------------------Одиннадцатый столбец значений

            //--------------------------Двенадцатый столбец значений
            
            //int ITTSU1DiplomaCounter = 0;
            int ITTSU1DiplomaBCounter = 0;
            int ITTSU1DiplomaMCounter = 0;
            int ITTSU1DiplomaSCounter = 0;
            int ITTSU1DiplomaSURCounter = 0;
            int ITTSU1DiplomaUOCounter = 0;
            int ITTSU1DiplomaUMOCounter = 0;

            int ITTSU2DiplomaCounter = 0;
            int ITTSU2DiplomaBCounter = 0;
            int ITTSU2DiplomaMCounter = 0;
            int ITTSU2DiplomaSCounter = 0;
            int ITTSU2DiplomaSURCounter = 0;
            int ITTSU2DiplomaUOCounter = 0;
            int ITTSU2DiplomaUMOCounter = 0;

            int IUIT1DiplomaCounter = 0;
            int IUIT2DiplomaCounter = 0;

            int VF1DiplomaCounter = 0;
            int VF2DiplomaCounter = 0;

            int Sum1DiplomaCounter = 0;
            int Sum2DiplomaCounter = 0;

            int AllDiplomaCounter = 0;
            //--------------------------Двенадцатый столбец значений

            //--------------------------Тринадцатый столбец значений
            
            //int ITTSU1TutPrCounter = 0;
            int ITTSU1TutPrBCounter = 0;
            int ITTSU1TutPrMCounter = 0;
            int ITTSU1TutPrSCounter = 0;
            int ITTSU1TutPrSURCounter = 0;
            int ITTSU1TutPrUOCounter = 0;
            int ITTSU1TutPrUMOCounter = 0;

            int ITTSU2TutPrCounter = 0;
            int ITTSU2TutPrBCounter = 0;
            int ITTSU2TutPrMCounter = 0;
            int ITTSU2TutPrSCounter = 0;
            int ITTSU2TutPrSURCounter = 0;
            int ITTSU2TutPrUOCounter = 0;
            int ITTSU2TutPrUMOCounter = 0;

            int IUIT1TutPrCounter = 0;
            int IUIT2TutPrCounter = 0;

            int VF1TutPrCounter = 0;
            int VF2TutPrCounter = 0;

            int Sum1TutPrCounter = 0;
            int Sum2TutPrCounter = 0;

            int AllTutPrCounter = 0;
            //--------------------------Тринадцатый столбец значений

            //--------------------------Четырнадцатый столбец значений
            
            //int ITTSU1ProdCounter = 0;
            int ITTSU1ProdBCounter = 0;
            int ITTSU1ProdMCounter = 0;
            int ITTSU1ProdSCounter = 0;
            int ITTSU1ProdSURCounter = 0;
            int ITTSU1ProdUOCounter = 0;
            int ITTSU1ProdUMOCounter = 0;

            int ITTSU2ProdCounter = 0;
            int ITTSU2ProdBCounter = 0;
            int ITTSU2ProdMCounter = 0;
            int ITTSU2ProdSCounter = 0;
            int ITTSU2ProdSURCounter = 0;
            int ITTSU2ProdUOCounter = 0;
            int ITTSU2ProdUMOCounter = 0;

            int IUIT1ProdCounter = 0;
            int IUIT2ProdCounter = 0;

            int VF1ProdCounter = 0;
            int VF2ProdCounter = 0;

            int Sum1ProdCounter = 0;
            int Sum2ProdCounter = 0;

            int AllProdCounter = 0;
            //--------------------------Четырнадцатый столбец значений

            //--------------------------Пятнадцатый столбец значений
            
            //int ITTSU1GAKCounter = 0;
            int ITTSU1GAKBCounter = 0;
            int ITTSU1GAKMCounter = 0;
            int ITTSU1GAKSCounter = 0;
            int ITTSU1GAKSURCounter = 0;
            int ITTSU1GAKUOCounter = 0;
            int ITTSU1GAKUMOCounter = 0;

            int ITTSU2GAKCounter = 0;
            int ITTSU2GAKBCounter = 0;
            int ITTSU2GAKMCounter = 0;
            int ITTSU2GAKSCounter = 0;
            int ITTSU2GAKSURCounter = 0;
            int ITTSU2GAKUOCounter = 0;
            int ITTSU2GAKUMOCounter = 0;

            int IUIT1GAKCounter = 0;
            int IUIT2GAKCounter = 0;

            int VF1GAKCounter = 0;
            int VF2GAKCounter = 0;

            int Sum1GAKCounter = 0;
            int Sum2GAKCounter = 0;

            int AllGAKCounter = 0;
            //--------------------------Пятнадцатый столбец значений

            //--------------------------Шестнадцатый столбец значений
            
            //int ITTSU1BudCounter = 0;
            //float ITTSU1BudZCounter = 0;

            int ITTSU1BudBCounter = 0;
            int ITTSU1BudMCounter = 0;
            int ITTSU1BudSCounter = 0;
            int ITTSU1BudSURCounter = 0;
            int ITTSU1BudUOCounter = 0;
            int ITTSU1BudUMOCounter = 0;

            float ITTSU1BudBZCounter = 0;
            float ITTSU1BudMZCounter = 0;
            float ITTSU1BudSZCounter = 0;
            float ITTSU1BudSURZCounter = 0;
            float ITTSU1BudUOZCounter = 0;
            float ITTSU1BudUMOZCounter = 0;

            int ITTSU2BudCounter = 0;
            float ITTSU2BudZCounter = 0;

            int ITTSU2BudBCounter = 0;
            int ITTSU2BudMCounter = 0;
            int ITTSU2BudSCounter = 0;
            int ITTSU2BudSURCounter = 0;
            int ITTSU2BudUOCounter = 0;
            int ITTSU2BudUMOCounter = 0;

            float ITTSU2BudBZCounter = 0;
            float ITTSU2BudMZCounter = 0;
            float ITTSU2BudSZCounter = 0;
            float ITTSU2BudSURZCounter = 0;
            float ITTSU2BudUOZCounter = 0;
            float ITTSU2BudUMOZCounter = 0;

            int IUIT1BudCounter = 0;
            int IUIT2BudCounter = 0;

            float IUIT1BudZCounter = 0;
            float IUIT2BudZCounter = 0;

            int VF1BudCounter = 0;
            int VF2BudCounter = 0;

            float VF1BudZCounter = 0;
            float VF2BudZCounter = 0;

            int Sum1BudCounter = 0;
            int Sum2BudCounter = 0;

            float Sum1BudZCounter = 0;
            float Sum2BudZCounter = 0;

            int AllBudCounter = 0;
            float AllBudZCounter = 0;
            //--------------------------Шестнадцатый столбец значений

            //--------------------------Семнадцатый столбец значений
            
            //int ITTSU1SumCounter = 0;
            //float ITTSU1SumZCounter = 0;

            int ITTSU1SumBCounter = 0;
            int ITTSU1SumMCounter = 0;
            int ITTSU1SumSCounter = 0;
            int ITTSU1SumSURCounter = 0;
            int ITTSU1SumUOCounter = 0;
            int ITTSU1SumUMOCounter = 0;

            float ITTSU1SumBZCounter = 0;
            float ITTSU1SumMZCounter = 0;
            float ITTSU1SumSZCounter = 0;
            float ITTSU1SumSURZCounter = 0;
            float ITTSU1SumUOZCounter = 0;
            float ITTSU1SumUMOZCounter = 0;

            int ITTSU2SumCounter = 0;
            float ITTSU2SumZCounter = 0;

            int ITTSU2SumBCounter = 0;
            int ITTSU2SumMCounter = 0;
            int ITTSU2SumSCounter = 0;
            int ITTSU2SumSURCounter = 0;
            int ITTSU2SumUOCounter = 0;
            int ITTSU2SumUMOCounter = 0;

            float ITTSU2SumBZCounter = 0;
            float ITTSU2SumMZCounter = 0;
            float ITTSU2SumSZCounter = 0;
            float ITTSU2SumSURZCounter = 0;
            float ITTSU2SumUOZCounter = 0;
            float ITTSU2SumUMOZCounter = 0;

            int IUIT1SumCounter = 0;
            int IUIT2SumCounter = 0;

            float IUIT1SumZCounter = 0;
            float IUIT2SumZCounter = 0;

            int VF1SumCounter = 0;
            int VF2SumCounter = 0;

            float VF1SumZCounter = 0;
            float VF2SumZCounter = 0;

            int Sum1SumCounter = 0;           
            int Sum2SumCounter = 0;

            float Sum1SumZCounter = 0;
            float Sum2SumZCounter = 0;
            
            int AllSumCounter = 0;
            float AllSumZCounter = 0;
            //--------------------------Семнадцатый столбец значений

            //--------------------------Восемнадцатый столбец значений
            
            //int ITTSU1AllCounter = 0;
            //float ITTSU1AllZCounter = 0;

            int ITTSU1AllBCounter = 0;
            int ITTSU1AllMCounter = 0;
            int ITTSU1AllSCounter = 0;
            int ITTSU1AllSURCounter = 0;
            int ITTSU1AllUOCounter = 0;
            int ITTSU1AllUMOCounter = 0;

            float ITTSU1AllBZCounter = 0;
            float ITTSU1AllMZCounter = 0;
            float ITTSU1AllSZCounter = 0;
            float ITTSU1AllSURZCounter = 0;
            float ITTSU1AllUOZCounter = 0;
            float ITTSU1AllUMOZCounter = 0;

            int ITTSU2AllCounter = 0;
            float ITTSU2AllZCounter = 0;

            int ITTSU2AllBCounter = 0;
            int ITTSU2AllMCounter = 0;
            int ITTSU2AllSCounter = 0;
            int ITTSU2AllSURCounter = 0;
            int ITTSU2AllUOCounter = 0;
            int ITTSU2AllUMOCounter = 0;

            float ITTSU2AllBZCounter = 0;
            float ITTSU2AllMZCounter = 0;
            float ITTSU2AllSZCounter = 0;
            float ITTSU2AllSURZCounter = 0;
            float ITTSU2AllUOZCounter = 0;
            float ITTSU2AllUMOZCounter = 0;

            int IUIT1AllCounter = 0;
            int IUIT2AllCounter = 0;

            float IUIT1AllZCounter = 0;
            float IUIT2AllZCounter = 0;

            int VF1AllCounter = 0;
            int VF2AllCounter = 0;

            float VF1AllZCounter = 0;
            float VF2AllZCounter = 0;

            int Sum1AllCounter = 0; 
            int Sum2AllCounter = 0;

            float Sum1AllZCounter = 0;
            float Sum2AllZCounter = 0;

            int AllAllCounter = 0;
            float AllAllZCounter = 0;
            //--------------------------Восемнадцатый столбец значений

            //--------------------------Аспирантура
            int SumPostGrad1Counter = 0;
            int EntPostGrad1Counter = 0;
            int SumPostGrad2Counter = 0;
            int EntPostGrad2Counter = 0;
            //--------------------------Аспирантура

            //--------------------------Магистратура
            int SumMagistry1Counter = 0;
            int EntMagistry1Counter = 0;
            int SumMagistry2Counter = 0;
            int EntMagistry2Counter = 0;

            int SumAllMagistry1Counter = 0;
            int SumAllMagistry2Counter = 0;
            //--------------------------Магистратура

            //--------------------------Посещение занятий
            int SumVisiting1Counter = 0;
            int EntVisiting1Counter = 0;
            int SumVisiting2Counter = 0;
            int EntVisiting2Counter = 0;
            //--------------------------Посещение занятий

            //--------------------------Полная сумма часов кафедры
            int SumAllCounter = 0;
            int EntAllCounter = 0;
            //--------------------------Полная сумма часов кафедры

            for (int i = 0; i <= mdlData.colDistribution.Count - 1; i++)
            {
                //Считаем что-либо только если строка не исключена из расчёта нагрузки
                if (!mdlData.colDistribution[i].flgExclude)
                {
                    //Если не указана дисциплина, то строка не имеет значения 
                    //для полного обсчёта нагрузки
                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        //Если не указана специальность, то строка не имеет значения для
                        //полного обсчёта нагрузки
                        if (!(mdlData.colDistribution[i].Speciality == null))
                        {
                            //Проверяем суммарное распределение видов занятий по ИТТСУ за 1-й семестр                
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("1 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ИТТСУ"))
                            {

                                AdditionFunction(mdlData.colDistribution[i], ref ITTSU1.LectCounter, ref ITTSU1.ExamCounter,
                                    ref ITTSU1.CredCounter, ref ITTSU1.RefCounter, ref ITTSU1.TutCounter, ref ITTSU1.LabCounter,
                                    ref ITTSU1.PractCounter, ref ITTSU1.IndCounter, ref ITTSU1.KRAPKCounter, ref ITTSU1.KursCounter,
                                    ref ITTSU1.PreDCounter, ref ITTSU1.DiplomaCounter, ref ITTSU1.TutPrCounter, ref ITTSU1.ProdCounter,
                                    ref ITTSU1.GAKCounter, ref ITTSU1.BudCounter, ref ITTSU1.BudZCounter, ref ITTSU1.SumCounter,
                                    ref ITTSU1.SumZCounter, ref ITTSU1.AllCounter, ref ITTSU1.AllZCounter);

                            }

                            //Проверяем суммарное распределение видов занятий по ИТТСУ за 1-й семестр
                            //бакалавриат
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("1 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                mdlData.colDistribution[i].Speciality.Diff.Equals("Б") &
                                mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ИТТСУ"))
                            {

                                AdditionFunction(mdlData.colDistribution[i], ref ITTSU1LectBCounter, ref ITTSU1ExamBCounter,
                                    ref ITTSU1CredBCounter, ref ITTSU1RefBCounter, ref ITTSU1TutBCounter, ref ITTSU1LabBCounter,
                                    ref ITTSU1PractBCounter, ref ITTSU1IndBCounter, ref ITTSU1KRAPKBCounter, ref ITTSU1KursBCounter,
                                    ref ITTSU1PreDBCounter, ref ITTSU1DiplomaBCounter, ref ITTSU1TutPrBCounter, ref ITTSU1ProdBCounter,
                                    ref ITTSU1GAKBCounter, ref ITTSU1BudBCounter, ref ITTSU1BudBZCounter, ref ITTSU1SumBCounter,
                                    ref ITTSU1SumBZCounter, ref ITTSU1AllBCounter, ref ITTSU1AllBZCounter);

                            }

                            //Проверяем суммарное распределение видов занятий по ИТТСУ за 1-й семестр
                            //магистратура
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("1 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                mdlData.colDistribution[i].Speciality.Diff.Equals("М") &
                                mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ИТТСУ"))
                            {
                                AdditionFunction(mdlData.colDistribution[i], ref ITTSU1LectMCounter, ref ITTSU1ExamMCounter,
                                    ref ITTSU1CredMCounter, ref ITTSU1RefMCounter, ref ITTSU1TutMCounter, ref ITTSU1LabMCounter,
                                    ref ITTSU1PractMCounter, ref ITTSU1IndMCounter, ref ITTSU1KRAPKMCounter, ref ITTSU1KursMCounter,
                                    ref ITTSU1PreDMCounter, ref ITTSU1DiplomaMCounter, ref ITTSU1TutPrMCounter, ref ITTSU1ProdMCounter,
                                    ref ITTSU1GAKMCounter, ref ITTSU1BudMCounter, ref ITTSU1BudMZCounter, ref ITTSU1SumMCounter,
                                    ref ITTSU1SumMZCounter, ref ITTSU1AllMCounter, ref ITTSU1AllMZCounter);
                            }

                            //Проверяем суммарное распределение видов занятий по ИТТСУ за 1-й семестр
                            //специалитет
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("1 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                mdlData.colDistribution[i].Speciality.Diff.Equals("С") &
                                mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ИТТСУ"))
                            {
                                AdditionFunction(mdlData.colDistribution[i], ref ITTSU1LectSCounter, ref ITTSU1ExamSCounter,
                                    ref ITTSU1CredSCounter, ref ITTSU1RefSCounter, ref ITTSU1TutSCounter, ref ITTSU1LabSCounter,
                                    ref ITTSU1PractSCounter, ref ITTSU1IndSCounter, ref ITTSU1KRAPKSCounter, ref ITTSU1KursSCounter,
                                    ref ITTSU1PreDSCounter, ref ITTSU1DiplomaSCounter, ref ITTSU1TutPrSCounter, ref ITTSU1ProdSCounter,
                                    ref ITTSU1GAKSCounter, ref ITTSU1BudSCounter, ref ITTSU1BudSZCounter, ref ITTSU1SumSCounter,
                                    ref ITTSU1SumSZCounter, ref ITTSU1AllSCounter, ref ITTSU1AllSZCounter);
                            }

                            //Проверяем суммарное распределение видов занятий по ИТТСУ за 1-й семестр
                            //совмещение учёбы и работы
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("1 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                mdlData.colDistribution[i].Speciality.Diff.Equals("СУР") &
                                mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ИТТСУ"))
                            {
                                AdditionFunction(mdlData.colDistribution[i], ref ITTSU1LectSURCounter, ref ITTSU1ExamSURCounter,
                                    ref ITTSU1CredSURCounter, ref ITTSU1RefSURCounter, ref ITTSU1TutSURCounter, ref ITTSU1LabSURCounter,
                                    ref ITTSU1PractSURCounter, ref ITTSU1IndSURCounter, ref ITTSU1KRAPKSURCounter, ref ITTSU1KursSURCounter,
                                    ref ITTSU1PreDSURCounter, ref ITTSU1DiplomaSURCounter, ref ITTSU1TutPrSURCounter, ref ITTSU1ProdSURCounter,
                                    ref ITTSU1GAKSURCounter, ref ITTSU1BudSURCounter, ref ITTSU1BudSURZCounter, ref ITTSU1SumSURCounter,
                                    ref ITTSU1SumSURZCounter, ref ITTSU1AllSURCounter, ref ITTSU1AllSURZCounter);
                            }

                            //Проверяем суммарное распределение видов занятий по ИТТСУ за 1-й семестр
                            //сокращённое обучение
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("1 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                mdlData.colDistribution[i].Speciality.Diff.Equals("УО") &
                                mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ИТТСУ"))
                            {
                                AdditionFunction(mdlData.colDistribution[i], ref ITTSU1LectUOCounter, ref ITTSU1ExamUOCounter,
                                    ref ITTSU1CredUOCounter, ref ITTSU1RefUOCounter, ref ITTSU1TutUOCounter, ref ITTSU1LabUOCounter,
                                    ref ITTSU1PractUOCounter, ref ITTSU1IndUOCounter, ref ITTSU1KRAPKUOCounter, ref ITTSU1KursUOCounter,
                                    ref ITTSU1PreDUOCounter, ref ITTSU1DiplomaUOCounter, ref ITTSU1TutPrUOCounter, ref ITTSU1ProdUOCounter,
                                    ref ITTSU1GAKUOCounter, ref ITTSU1BudUOCounter, ref ITTSU1BudUOZCounter, ref ITTSU1SumUOCounter,
                                    ref ITTSU1SumUOZCounter, ref ITTSU1AllUOCounter, ref ITTSU1AllUOZCounter);
                            }

                            //Проверяем суммарное распределение видов занятий по ИТТСУ за 1-й семестр,
                            //находящееся в ведомстве УМО (кроме СУР и УО)
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("1 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                !mdlData.colDistribution[i].Speciality.Diff.Equals("УО") &
                                !mdlData.colDistribution[i].Speciality.Diff.Equals("СУР") &
                                mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ИТТСУ"))
                            {
                                AdditionFunction(mdlData.colDistribution[i], ref ITTSU1LectUMOCounter, ref ITTSU1ExamUMOCounter,
                                    ref ITTSU1CredUMOCounter, ref ITTSU1RefUMOCounter, ref ITTSU1TutUMOCounter, ref ITTSU1LabUMOCounter,
                                    ref ITTSU1PractUMOCounter, ref ITTSU1IndUMOCounter, ref ITTSU1KRAPKUMOCounter, ref ITTSU1KursUMOCounter,
                                    ref ITTSU1PreDUMOCounter, ref ITTSU1DiplomaUMOCounter, ref ITTSU1TutPrUMOCounter, ref ITTSU1ProdUMOCounter,
                                    ref ITTSU1GAKUMOCounter, ref ITTSU1BudUMOCounter, ref ITTSU1BudUMOZCounter, ref ITTSU1SumUMOCounter,
                                    ref ITTSU1SumUMOZCounter, ref ITTSU1AllUMOCounter, ref ITTSU1AllUMOZCounter);
                            }
                        }
                    }

                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        if (!(mdlData.colDistribution[i].Speciality == null))
                        {
                            //Проверяем суммарное распределение лекций по ИУИТу за 1-й семестр                
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("1 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ИУИТ"))
                            {
                                AdditionFunction(mdlData.colDistribution[i], ref IUIT1LectCounter, ref IUIT1ExamCounter,
                                    ref IUIT1CredCounter, ref IUIT1RefCounter, ref IUIT1TutCounter, ref IUIT1LabCounter,
                                    ref IUIT1PractCounter, ref IUIT1IndCounter, ref IUIT1KRAPKCounter, ref IUIT1KursCounter,
                                    ref IUIT1PreDCounter, ref IUIT1DiplomaCounter, ref IUIT1TutPrCounter, ref IUIT1ProdCounter,
                                    ref IUIT1GAKCounter, ref IUIT1BudCounter, ref IUIT1BudZCounter, ref IUIT1SumCounter,
                                    ref IUIT1SumZCounter, ref IUIT1AllCounter, ref IUIT1AllZCounter);
                            }
                        }
                    }

                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        if (!(mdlData.colDistribution[i].Speciality == null))
                        {
                            //Проверяем суммарное распределение лекций по вечернему факультету за 1-й семестр                
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("1 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ВФ"))
                            {
                                AdditionFunction(mdlData.colDistribution[i], ref VF1LectCounter, ref VF1ExamCounter,
                                    ref VF1CredCounter, ref VF1RefCounter, ref VF1TutCounter, ref VF1LabCounter,
                                    ref VF1PractCounter, ref VF1IndCounter, ref VF1KRAPKCounter, ref VF1KursCounter,
                                    ref VF1PreDCounter, ref VF1DiplomaCounter, ref VF1TutPrCounter, ref VF1ProdCounter,
                                    ref VF1GAKCounter, ref VF1BudCounter, ref VF1BudZCounter, ref VF1SumCounter,
                                    ref VF1SumZCounter, ref VF1AllCounter, ref VF1AllZCounter);
                            }
                        }
                    }

                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        //Проверяем суммарное распределение всех лекций за 1-й семестр                
                        if (mdlData.colDistribution[i].Semestr.SemNum.Equals("1 семестр") &
                            !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                            !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                            !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                            !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                            !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий"))
                        {
                            AdditionFunction(mdlData.colDistribution[i], ref Sum1LectCounter, ref Sum1ExamCounter,
                                ref Sum1CredCounter, ref Sum1RefCounter, ref Sum1TutCounter, ref Sum1LabCounter,
                                ref Sum1PractCounter, ref Sum1IndCounter, ref Sum1KRAPKCounter, ref Sum1KursCounter,
                                ref Sum1PreDCounter, ref Sum1DiplomaCounter, ref Sum1TutPrCounter, ref Sum1ProdCounter,
                                ref Sum1GAKCounter, ref Sum1BudCounter, ref Sum1BudZCounter, ref Sum1SumCounter,
                                ref Sum1SumZCounter, ref Sum1AllCounter, ref Sum1AllZCounter);
                        }
                    }

                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        //Проверяем суммарное распределение по аспирантуре за 1-й семестр                
                        if (mdlData.colDistribution[i].Semestr.SemNum.Equals("1 семестр") &
                            (mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") ||
                            mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)")))
                        {
                            EntPostGrad1Counter += mdlData.colDistribution[i].EnteredHours;
                            SumPostGrad1Counter += mdlData.colDistribution[i].PostGrad;
                        }
                    }

                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        //Проверяем суммарное распределение по посещению занятий за 1-й семестр                
                        if (mdlData.colDistribution[i].Semestr.SemNum.Equals("1 семестр") &
                            mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий"))
                        {
                            EntVisiting1Counter += mdlData.colDistribution[i].EnteredHours;
                            SumVisiting1Counter += mdlData.colDistribution[i].Visiting;
                        }
                    }

                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        //Проверяем суммарное распределение по магистерской программе за 1-й семестр                
                        if (mdlData.colDistribution[i].Semestr.SemNum.Equals("1 семестр") &
                            mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами"))
                        {
                            EntMagistry1Counter += mdlData.colDistribution[i].EnteredHours;
                            SumMagistry1Counter += mdlData.colDistribution[i].Magistry;
                        }
                    }

                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        //Проверяем суммарное распределение по магистерской программе за 1-й семестр                
                        if (mdlData.colDistribution[i].Semestr.SemNum.Equals("1 семестр") &
                            mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром"))
                        {
                            EntMagistry1Counter += mdlData.colDistribution[i].EnteredHours;
                            SumAllMagistry1Counter += mdlData.colDistribution[i].Magistry;
                        }
                    }

                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        if (!(mdlData.colDistribution[i].Speciality == null))
                        {
                            //Проверяем суммарное распределение видов занятий по ИТТСУ за 2-й семестр                
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("2 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ИТТСУ"))
                            {
                                AdditionFunction(mdlData.colDistribution[i], ref ITTSU2LectCounter, ref ITTSU2ExamCounter,
                                    ref ITTSU2CredCounter, ref ITTSU2RefCounter, ref ITTSU2TutCounter, ref ITTSU2LabCounter,
                                    ref ITTSU2PractCounter, ref ITTSU2IndCounter, ref ITTSU2KRAPKCounter, ref ITTSU2KursCounter,
                                    ref ITTSU2PreDCounter, ref ITTSU2DiplomaCounter, ref ITTSU2TutPrCounter, ref ITTSU2ProdCounter,
                                    ref ITTSU2GAKCounter, ref ITTSU2BudCounter, ref ITTSU2BudZCounter, ref ITTSU2SumCounter,
                                    ref ITTSU2SumZCounter, ref ITTSU2AllCounter, ref ITTSU2AllZCounter);
                            }

                            //Проверяем суммарное распределение по видам зантий по ИТТСУ за 2-й семестр
                            //бакалавриат
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("2 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                mdlData.colDistribution[i].Speciality.Diff.Equals("Б") &
                                mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ИТТСУ"))
                            {

                                AdditionFunction(mdlData.colDistribution[i], ref ITTSU2LectBCounter, ref ITTSU2ExamBCounter,
                                    ref ITTSU2CredBCounter, ref ITTSU2RefBCounter, ref ITTSU2TutBCounter, ref ITTSU2LabBCounter,
                                    ref ITTSU2PractBCounter, ref ITTSU2IndBCounter, ref ITTSU2KRAPKBCounter, ref ITTSU2KursBCounter,
                                    ref ITTSU2PreDBCounter, ref ITTSU2DiplomaBCounter, ref ITTSU2TutPrBCounter, ref ITTSU2ProdBCounter,
                                    ref ITTSU2GAKBCounter, ref ITTSU2BudBCounter, ref ITTSU2BudBZCounter, ref ITTSU2SumBCounter,
                                    ref ITTSU2SumBZCounter, ref ITTSU2AllBCounter, ref ITTSU2AllBZCounter);

                            }

                            //Проверяем суммарное распределение по видам зантий по ИТТСУ за 2-й семестр
                            //магистратура
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("2 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                mdlData.colDistribution[i].Speciality.Diff.Equals("М") &
                                mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ИТТСУ"))
                            {
                                AdditionFunction(mdlData.colDistribution[i], ref ITTSU2LectMCounter, ref ITTSU2ExamMCounter,
                                    ref ITTSU2CredMCounter, ref ITTSU2RefMCounter, ref ITTSU2TutMCounter, ref ITTSU2LabMCounter,
                                    ref ITTSU2PractMCounter, ref ITTSU2IndMCounter, ref ITTSU2KRAPKMCounter, ref ITTSU2KursMCounter,
                                    ref ITTSU2PreDMCounter, ref ITTSU2DiplomaMCounter, ref ITTSU2TutPrMCounter, ref ITTSU2ProdMCounter,
                                    ref ITTSU2GAKMCounter, ref ITTSU2BudMCounter, ref ITTSU2BudMZCounter, ref ITTSU2SumMCounter,
                                    ref ITTSU2SumMZCounter, ref ITTSU2AllMCounter, ref ITTSU2AllMZCounter);
                            }

                            //Проверяем суммарное распределение по видам зантий по ИТТСУ за 2-й семестр
                            //специалитет
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("2 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                mdlData.colDistribution[i].Speciality.Diff.Equals("С") &
                                mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ИТТСУ"))
                            {
                                AdditionFunction(mdlData.colDistribution[i], ref ITTSU2LectSCounter, ref ITTSU2ExamSCounter,
                                    ref ITTSU2CredSCounter, ref ITTSU2RefSCounter, ref ITTSU2TutSCounter, ref ITTSU2LabSCounter,
                                    ref ITTSU2PractSCounter, ref ITTSU2IndSCounter, ref ITTSU2KRAPKSCounter, ref ITTSU2KursSCounter,
                                    ref ITTSU2PreDSCounter, ref ITTSU2DiplomaSCounter, ref ITTSU2TutPrSCounter, ref ITTSU2ProdSCounter,
                                    ref ITTSU2GAKSCounter, ref ITTSU2BudSCounter, ref ITTSU2BudSZCounter, ref ITTSU2SumSCounter,
                                    ref ITTSU2SumSZCounter, ref ITTSU2AllSCounter, ref ITTSU2AllSZCounter);
                            }

                            //Проверяем суммарное распределение по видам зантий по ИТТСУ за 2-й семестр
                            //совмещение учёбы и работы
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("2 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                mdlData.colDistribution[i].Speciality.Diff.Equals("СУР") &
                                mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ИТТСУ"))
                            {
                                AdditionFunction(mdlData.colDistribution[i], ref ITTSU2LectSURCounter, ref ITTSU2ExamSURCounter,
                                    ref ITTSU2CredSURCounter, ref ITTSU2RefSURCounter, ref ITTSU2TutSURCounter, ref ITTSU2LabSURCounter,
                                    ref ITTSU2PractSURCounter, ref ITTSU2IndSURCounter, ref ITTSU2KRAPKSURCounter, ref ITTSU2KursSURCounter,
                                    ref ITTSU2PreDSURCounter, ref ITTSU2DiplomaSURCounter, ref ITTSU2TutPrSURCounter, ref ITTSU2ProdSURCounter,
                                    ref ITTSU2GAKSURCounter, ref ITTSU2BudSURCounter, ref ITTSU2BudSURZCounter, ref ITTSU2SumSURCounter,
                                    ref ITTSU2SumSURZCounter, ref ITTSU2AllSURCounter, ref ITTSU2AllSURZCounter);
                            }

                            //Проверяем суммарное распределение по видам зантий по ИТТСУ за 2-й семестр
                            //ускоренное обучение
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("2 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                mdlData.colDistribution[i].Speciality.Diff.Equals("УО") &
                                mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ИТТСУ"))
                            {
                                AdditionFunction(mdlData.colDistribution[i], ref ITTSU2LectUOCounter, ref ITTSU2ExamUOCounter,
                                    ref ITTSU2CredUOCounter, ref ITTSU2RefUOCounter, ref ITTSU2TutUOCounter, ref ITTSU2LabUOCounter,
                                    ref ITTSU2PractUOCounter, ref ITTSU2IndUOCounter, ref ITTSU2KRAPKUOCounter, ref ITTSU2KursUOCounter,
                                    ref ITTSU2PreDUOCounter, ref ITTSU2DiplomaUOCounter, ref ITTSU2TutPrUOCounter, ref ITTSU2ProdUOCounter,
                                    ref ITTSU2GAKUOCounter, ref ITTSU2BudUOCounter, ref ITTSU2BudUOZCounter, ref ITTSU2SumUOCounter,
                                    ref ITTSU2SumUOZCounter, ref ITTSU2AllUOCounter, ref ITTSU2AllUOZCounter);
                            }

                            //Проверяем суммарное распределение видов занятий по ИТТСУ за 2-й семестр,
                            //находящееся в ведомстве УМО (кроме СУР и УО)
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("2 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                !mdlData.colDistribution[i].Speciality.Diff.Equals("УО") &
                                !mdlData.colDistribution[i].Speciality.Diff.Equals("СУР") &
                                mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ИТТСУ"))
                            {
                                AdditionFunction(mdlData.colDistribution[i], ref ITTSU2LectUMOCounter, ref ITTSU2ExamUMOCounter,
                                    ref ITTSU2CredUMOCounter, ref ITTSU2RefUMOCounter, ref ITTSU2TutUMOCounter, ref ITTSU2LabUMOCounter,
                                    ref ITTSU2PractUMOCounter, ref ITTSU2IndUMOCounter, ref ITTSU2KRAPKUMOCounter, ref ITTSU2KursUMOCounter,
                                    ref ITTSU2PreDUMOCounter, ref ITTSU2DiplomaUMOCounter, ref ITTSU2TutPrUMOCounter, ref ITTSU2ProdUMOCounter,
                                    ref ITTSU2GAKUMOCounter, ref ITTSU2BudUMOCounter, ref ITTSU2BudUMOZCounter, ref ITTSU2SumUMOCounter,
                                    ref ITTSU2SumUMOZCounter, ref ITTSU2AllUMOCounter, ref ITTSU2AllUMOZCounter);
                            }
                        }
                    }

                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        if (!(mdlData.colDistribution[i].Speciality == null))
                        {
                            //Проверяем суммарное распределение лекций по ИУИТу за 2-й семестр                
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("2 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ИУИТ"))
                            {
                                AdditionFunction(mdlData.colDistribution[i], ref IUIT2LectCounter, ref IUIT2ExamCounter,
                                    ref IUIT2CredCounter, ref IUIT2RefCounter, ref IUIT2TutCounter, ref IUIT2LabCounter,
                                    ref IUIT2PractCounter, ref IUIT2IndCounter, ref IUIT2KRAPKCounter, ref IUIT2KursCounter,
                                    ref IUIT2PreDCounter, ref IUIT2DiplomaCounter, ref IUIT2TutPrCounter, ref IUIT2ProdCounter,
                                    ref IUIT2GAKCounter, ref IUIT2BudCounter, ref IUIT2BudZCounter, ref IUIT2SumCounter,
                                    ref IUIT2SumZCounter, ref IUIT2AllCounter, ref IUIT2AllZCounter);
                            }
                        }
                    }

                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        if (!(mdlData.colDistribution[i].Speciality == null))
                        {
                            //Проверяем суммарное распределение лекций по вечернему факультету за 2-й семестр
                            if (mdlData.colDistribution[i].Semestr.SemNum.Equals("2 семестр") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                                !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                                (mdlData.colDistribution[i].Speciality.Faculty.Short.Equals("ВФ")))
                            {
                                AdditionFunction(mdlData.colDistribution[i], ref VF2LectCounter, ref VF2ExamCounter,
                                    ref VF2CredCounter, ref VF2RefCounter, ref VF2TutCounter, ref VF2LabCounter,
                                    ref VF2PractCounter, ref VF2IndCounter, ref VF2KRAPKCounter, ref VF2KursCounter,
                                    ref VF2PreDCounter, ref VF2DiplomaCounter, ref VF2TutPrCounter, ref VF2ProdCounter,
                                    ref VF2GAKCounter, ref VF2BudCounter, ref VF2BudZCounter, ref VF2SumCounter,
                                    ref VF2SumZCounter, ref VF2AllCounter, ref VF2AllZCounter);
                            }
                        }
                    }

                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        //Проверяем суммарное распределение всех лекций за 2-й семестр                
                        if (mdlData.colDistribution[i].Semestr.SemNum.Equals("2 семестр") &
                            !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") &
                            !mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)") &
                            !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") &
                            !mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") &
                            !mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий"))
                        {
                            AdditionFunction(mdlData.colDistribution[i], ref Sum2LectCounter, ref Sum2ExamCounter,
                                ref Sum2CredCounter, ref Sum2RefCounter, ref Sum2TutCounter, ref Sum2LabCounter,
                                ref Sum2PractCounter, ref Sum2IndCounter, ref Sum2KRAPKCounter, ref Sum2KursCounter,
                                ref Sum2PreDCounter, ref Sum2DiplomaCounter, ref Sum2TutPrCounter, ref Sum2ProdCounter,
                                ref Sum2GAKCounter, ref Sum2BudCounter, ref Sum2BudZCounter, ref Sum2SumCounter,
                                ref Sum2SumZCounter, ref Sum2AllCounter, ref Sum2AllZCounter);
                        }
                    }

                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        //Проверяем суммарное распределение по аспирантуре за 2-й семестр                
                        if (mdlData.colDistribution[i].Semestr.SemNum.Equals("2 семестр") &
                            (mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") ||
                            mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)")))
                        {
                            EntPostGrad2Counter += mdlData.colDistribution[i].EnteredHours;
                            SumPostGrad2Counter += mdlData.colDistribution[i].PostGrad;
                        }
                    }

                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        //Проверяем суммарное распределение по посещению занятий за 2-й семестр                
                        if (mdlData.colDistribution[i].Semestr.SemNum.Equals("2 семестр") &
                            mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий"))
                        {
                            EntVisiting2Counter += mdlData.colDistribution[i].EnteredHours;
                            SumVisiting2Counter += mdlData.colDistribution[i].Visiting;
                        }
                    }

                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        //Проверяем суммарное распределение по магистерской программе за 2-й семестр                
                        if (mdlData.colDistribution[i].Semestr.SemNum.Equals("2 семестр") &
                            mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами"))
                        {
                            EntMagistry2Counter += mdlData.colDistribution[i].EnteredHours;
                            SumMagistry2Counter += mdlData.colDistribution[i].Magistry;
                        }
                    }

                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        //Проверяем суммарное распределение по магистерской программе за 2-й семестр                
                        if (mdlData.colDistribution[i].Semestr.SemNum.Equals("2 семестр") &
                            mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром"))
                        {
                            EntMagistry2Counter += mdlData.colDistribution[i].EnteredHours;
                            SumAllMagistry2Counter += mdlData.colDistribution[i].Magistry;
                        }
                    }

                    AdditionFunction(mdlData.colDistribution[i], ref AllLectCounter, ref AllExamCounter,
                        ref AllCredCounter, ref AllRefCounter, ref AllTutCounter, ref AllLabCounter,
                        ref AllPractCounter, ref AllIndCounter, ref AllKRAPKCounter, ref AllKursCounter,
                        ref AllPreDCounter, ref AllDiplomaCounter, ref AllTutPrCounter, ref AllProdCounter,
                        ref AllGAKCounter, ref AllBudCounter, ref AllBudZCounter, ref AllSumCounter,
                        ref AllSumZCounter, ref AllAllCounter, ref AllAllZCounter);

                    //В годовом отчёте не требуется считать руководства, аспирантуру и посещение занятий
                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        if (mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") |
                            mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") |
                            mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") |
                            mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") |
                            mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)"))
                        {
                            AllSumCounter -= mdlData.colDistribution[i].EnteredHours;
                        }
                    }

                    EntAllCounter += mdlData.colDistribution[i].EnteredHours;

                    //В годовом отчёте не требуется считать руководства, аспирантуру и посещение занятий
                    if (!(mdlData.colDistribution[i].Subject == null))
                    {
                        if (mdlData.colDistribution[i].Subject.Subject.Equals("Посещение занятий") |
                            mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистрами") |
                            mdlData.colDistribution[i].Subject.Subject.Equals("Руководство магистром") |
                            mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура") |
                            mdlData.colDistribution[i].Subject.Subject.Equals("Аспирантура (лекции)"))
                        {
                            AllAllCounter -= mdlData.toSumDistributionComponents(mdlData.colDistribution[i]);
                        }
                    }

                    SumAllCounter += mdlData.toSumDistributionComponents(mdlData.colDistribution[i]);
                }
            }

            //--------------------------Первый столбец значений
            txtITTSU1Lect.Text = ITTSU1.LectCounter.ToString();
            txtITTSU1LectB.Text = ITTSU1LectBCounter.ToString();
            txtITTSU1LectM.Text = ITTSU1LectMCounter.ToString();
            txtITTSU1LectS.Text = ITTSU1LectSCounter.ToString();
            txtITTSU1LectSUR.Text = ITTSU1LectSURCounter.ToString();
            txtITTSU1LectUO.Text = ITTSU1LectUOCounter.ToString();
            txtITTSU1LectUMO.Text = ITTSU1LectUMOCounter.ToString();

            txtITTSU2Lect.Text = ITTSU2LectCounter.ToString();
            txtITTSU2LectB.Text = ITTSU2LectBCounter.ToString();
            txtITTSU2LectM.Text = ITTSU2LectMCounter.ToString();
            txtITTSU2LectS.Text = ITTSU2LectSCounter.ToString();
            txtITTSU2LectSUR.Text = ITTSU2LectSURCounter.ToString();
            txtITTSU2LectUO.Text = ITTSU2LectUOCounter.ToString();
            txtITTSU2LectUMO.Text = ITTSU2LectUMOCounter.ToString();

            txtIUIT1Lect.Text = IUIT1LectCounter.ToString();
            txtIUIT2Lect.Text = IUIT2LectCounter.ToString();

            txtVF1Lect.Text = VF1LectCounter.ToString();
            txtVF2Lect.Text = VF2LectCounter.ToString();
            
            txtSum1Lect.Text = Sum1LectCounter.ToString();
            txtSum2Lect.Text = Sum2LectCounter.ToString();            

            txtAllLect.Text = AllLectCounter.ToString();
            //--------------------------Первый столбец значений

            //--------------------------Второй столбец значений
            txtITTSU1Exam.Text = ITTSU1.ExamCounter.ToString();
            txtITTSU1ExamB.Text = ITTSU1ExamBCounter.ToString();
            txtITTSU1ExamM.Text = ITTSU1ExamMCounter.ToString();
            txtITTSU1ExamS.Text = ITTSU1ExamSCounter.ToString();
            txtITTSU1ExamSUR.Text = ITTSU1ExamSURCounter.ToString();
            txtITTSU1ExamUO.Text = ITTSU1ExamUOCounter.ToString();
            txtITTSU1ExamUMO.Text = ITTSU1ExamUMOCounter.ToString();

            txtITTSU2Exam.Text = ITTSU2ExamCounter.ToString();
            txtITTSU2ExamB.Text = ITTSU2ExamBCounter.ToString();
            txtITTSU2ExamM.Text = ITTSU2ExamMCounter.ToString();
            txtITTSU2ExamS.Text = ITTSU2ExamSCounter.ToString();
            txtITTSU2ExamSUR.Text = ITTSU2ExamSURCounter.ToString();
            txtITTSU2ExamUO.Text = ITTSU2ExamUOCounter.ToString();
            txtITTSU2ExamUMO.Text = ITTSU2ExamUMOCounter.ToString();

            txtIUIT1Exam.Text = IUIT1ExamCounter.ToString();
            txtIUIT2Exam.Text = IUIT2ExamCounter.ToString();

            txtVF1Exam.Text = VF1ExamCounter.ToString();
            txtVF2Exam.Text = VF2ExamCounter.ToString();

            txtSum1Exam.Text = Sum1ExamCounter.ToString();
            txtSum2Exam.Text = Sum2ExamCounter.ToString();

            txtAllExam.Text = AllExamCounter.ToString();
            //--------------------------Второй столбец значений

            //--------------------------Третий столбец значений
            txtITTSU1Cred.Text = ITTSU1.CredCounter.ToString();
            txtITTSU1CredB.Text = ITTSU1CredBCounter.ToString();
            txtITTSU1CredM.Text = ITTSU1CredMCounter.ToString();
            txtITTSU1CredS.Text = ITTSU1CredSCounter.ToString();
            txtITTSU1CredSUR.Text = ITTSU1CredSURCounter.ToString();
            txtITTSU1CredUO.Text = ITTSU1CredUOCounter.ToString();
            txtITTSU1CredUMO.Text = ITTSU1CredUMOCounter.ToString();

            txtITTSU2Cred.Text = ITTSU2CredCounter.ToString();
            txtITTSU2CredB.Text = ITTSU2CredBCounter.ToString();
            txtITTSU2CredM.Text = ITTSU2CredMCounter.ToString();
            txtITTSU2CredS.Text = ITTSU2CredSCounter.ToString();
            txtITTSU2CredSUR.Text = ITTSU2CredSURCounter.ToString();
            txtITTSU2CredUO.Text = ITTSU2CredUOCounter.ToString();
            txtITTSU2CredUMO.Text = ITTSU2CredUMOCounter.ToString();

            txtIUIT1Cred.Text = IUIT1CredCounter.ToString();
            txtIUIT2Cred.Text = IUIT2CredCounter.ToString();

            txtVF1Cred.Text = VF1CredCounter.ToString();
            txtVF2Cred.Text = VF2CredCounter.ToString();

            txtSum1Cred.Text = Sum1CredCounter.ToString();
            txtSum2Cred.Text = Sum2CredCounter.ToString();

            txtAllCred.Text = AllCredCounter.ToString();
            //--------------------------Третий столбец значений

            //--------------------------Четвёртый столбец значений
            txtITTSU1Ref.Text = ITTSU1.RefCounter.ToString();
            txtITTSU1RefB.Text = ITTSU1RefBCounter.ToString();
            txtITTSU1RefM.Text = ITTSU1RefMCounter.ToString();
            txtITTSU1RefS.Text = ITTSU1RefSCounter.ToString();
            txtITTSU1RefSUR.Text = ITTSU1RefSURCounter.ToString();
            txtITTSU1RefUO.Text = ITTSU1RefUOCounter.ToString();
            txtITTSU1RefUMO.Text = ITTSU1RefUMOCounter.ToString();

            txtITTSU2Ref.Text = ITTSU2RefCounter.ToString();
            txtITTSU2RefB.Text = ITTSU2RefBCounter.ToString();
            txtITTSU2RefM.Text = ITTSU2RefMCounter.ToString();
            txtITTSU2RefS.Text = ITTSU2RefSCounter.ToString();
            txtITTSU2RefSUR.Text = ITTSU2RefSURCounter.ToString();
            txtITTSU2RefUO.Text = ITTSU2RefUOCounter.ToString();
            txtITTSU2RefUMO.Text = ITTSU2RefUMOCounter.ToString();

            txtIUIT1Ref.Text = IUIT1RefCounter.ToString();
            txtIUIT2Ref.Text = IUIT2RefCounter.ToString();

            txtVF1Ref.Text = VF1RefCounter.ToString();
            txtVF2Ref.Text = VF2RefCounter.ToString();

            txtSum1Ref.Text = Sum1RefCounter.ToString();
            txtSum2Ref.Text = Sum2RefCounter.ToString();

            txtAllRef.Text = AllRefCounter.ToString();
            //--------------------------Четвёртый столбец значений

            //--------------------------Пятый столбец значений
            txtITTSU1Tut.Text = ITTSU1.TutCounter.ToString();
            txtITTSU1TutB.Text = ITTSU1TutBCounter.ToString();
            txtITTSU1TutM.Text = ITTSU1TutMCounter.ToString();
            txtITTSU1TutS.Text = ITTSU1TutSCounter.ToString();
            txtITTSU1TutSUR.Text = ITTSU1TutSURCounter.ToString();
            txtITTSU1TutUO.Text = ITTSU1TutUOCounter.ToString();
            txtITTSU1TutUMO.Text = ITTSU1TutUMOCounter.ToString();

            txtITTSU2Tut.Text = ITTSU2TutCounter.ToString();
            txtITTSU2TutB.Text = ITTSU2TutBCounter.ToString();
            txtITTSU2TutM.Text = ITTSU2TutMCounter.ToString();
            txtITTSU2TutS.Text = ITTSU2TutSCounter.ToString();
            txtITTSU2TutSUR.Text = ITTSU2TutSURCounter.ToString();
            txtITTSU2TutUO.Text = ITTSU2TutUOCounter.ToString();
            txtITTSU2TutUMO.Text = ITTSU2TutUMOCounter.ToString();

            txtIUIT1Tut.Text = IUIT1TutCounter.ToString();
            txtIUIT2Tut.Text = IUIT2TutCounter.ToString();

            txtVF1Tut.Text = VF1TutCounter.ToString();
            txtVF2Tut.Text = VF2TutCounter.ToString();

            txtSum1Tut.Text = Sum1TutCounter.ToString();
            txtSum2Tut.Text = Sum2TutCounter.ToString();

            txtAllTut.Text = AllTutCounter.ToString();
            //--------------------------Пятый столбец значений

            //--------------------------Шестой столбец значений
            txtITTSU1Lab.Text = ITTSU1.LabCounter.ToString();
            txtITTSU1LabB.Text = ITTSU1LabBCounter.ToString();
            txtITTSU1LabM.Text = ITTSU1LabMCounter.ToString();
            txtITTSU1LabS.Text = ITTSU1LabSCounter.ToString();
            txtITTSU1LabSUR.Text = ITTSU1LabSURCounter.ToString();
            txtITTSU1LabUO.Text = ITTSU1LabUOCounter.ToString();
            txtITTSU1LabUMO.Text = ITTSU1LabUMOCounter.ToString();

            txtITTSU2Lab.Text = ITTSU2LabCounter.ToString();
            txtITTSU2LabB.Text = ITTSU2LabBCounter.ToString();
            txtITTSU2LabM.Text = ITTSU2LabMCounter.ToString();
            txtITTSU2LabS.Text = ITTSU2LabSCounter.ToString();
            txtITTSU2LabSUR.Text = ITTSU2LabSURCounter.ToString();
            txtITTSU2LabUO.Text = ITTSU2LabUOCounter.ToString();
            txtITTSU2LabUMO.Text = ITTSU2LabUMOCounter.ToString();

            txtIUIT1Lab.Text = IUIT1LabCounter.ToString();
            txtIUIT2Lab.Text = IUIT2LabCounter.ToString();

            txtVF1Lab.Text = VF1LabCounter.ToString();
            txtVF2Lab.Text = VF2LabCounter.ToString();

            txtSum1Lab.Text = Sum1LabCounter.ToString();
            txtSum2Lab.Text = Sum2LabCounter.ToString();

            txtAllLab.Text = AllLabCounter.ToString();
            //--------------------------Шестой столбец значений

            //--------------------------Седьмой столбец значений
            txtITTSU1Pract.Text = ITTSU1.PractCounter.ToString();
            txtITTSU1PractB.Text = ITTSU1PractBCounter.ToString();
            txtITTSU1PractM.Text = ITTSU1PractMCounter.ToString();
            txtITTSU1PractS.Text = ITTSU1PractSCounter.ToString();
            txtITTSU1PractSUR.Text = ITTSU1PractSURCounter.ToString();
            txtITTSU1PractUO.Text = ITTSU1PractUOCounter.ToString();
            txtITTSU1PractUMO.Text = ITTSU1PractUMOCounter.ToString();

            txtITTSU2Pract.Text = ITTSU2PractCounter.ToString();
            txtITTSU2PractB.Text = ITTSU2PractBCounter.ToString();
            txtITTSU2PractM.Text = ITTSU2PractMCounter.ToString();
            txtITTSU2PractS.Text = ITTSU2PractSCounter.ToString();
            txtITTSU2PractSUR.Text = ITTSU2PractSURCounter.ToString();
            txtITTSU2PractUO.Text = ITTSU2PractUOCounter.ToString();
            txtITTSU2PractUMO.Text = ITTSU2PractUMOCounter.ToString();

            txtIUIT1Pract.Text = IUIT1PractCounter.ToString();
            txtIUIT2Pract.Text = IUIT2PractCounter.ToString();

            txtVF1Pract.Text = VF1PractCounter.ToString();
            txtVF2Pract.Text = VF2PractCounter.ToString();

            txtSum1Pract.Text = Sum1PractCounter.ToString();
            txtSum2Pract.Text = Sum2PractCounter.ToString();

            txtAllPract.Text = AllPractCounter.ToString();
            //--------------------------Седьмой столбец значений

            //--------------------------Восьмой столбец значений
            txtITTSU1Ind.Text = ITTSU1.IndCounter.ToString();
            txtITTSU1IndB.Text = ITTSU1IndBCounter.ToString();
            txtITTSU1IndM.Text = ITTSU1IndMCounter.ToString();
            txtITTSU1IndS.Text = ITTSU1IndSCounter.ToString();
            txtITTSU1IndSUR.Text = ITTSU1IndSURCounter.ToString();
            txtITTSU1IndUO.Text = ITTSU1IndUOCounter.ToString();
            txtITTSU1IndUMO.Text = ITTSU1IndUMOCounter.ToString();

            txtITTSU2Ind.Text = ITTSU2IndCounter.ToString();
            txtITTSU2IndB.Text = ITTSU2IndBCounter.ToString();
            txtITTSU2IndM.Text = ITTSU2IndMCounter.ToString();
            txtITTSU2IndS.Text = ITTSU2IndSCounter.ToString();
            txtITTSU2IndSUR.Text = ITTSU2IndSURCounter.ToString();
            txtITTSU2IndUO.Text = ITTSU2IndUOCounter.ToString();
            txtITTSU2IndUMO.Text = ITTSU2IndUMOCounter.ToString();

            txtIUIT1Ind.Text = IUIT1IndCounter.ToString();
            txtIUIT2Ind.Text = IUIT2IndCounter.ToString();

            txtVF1Ind.Text = VF1IndCounter.ToString();
            txtVF2Ind.Text = VF2IndCounter.ToString();

            txtSum1Ind.Text = Sum1IndCounter.ToString();
            txtSum2Ind.Text = Sum2IndCounter.ToString();

            txtAllInd.Text = AllIndCounter.ToString();
            //--------------------------Восьмой столбец значений

            //--------------------------Девятый столбец значений
            txtITTSU1KRAPK.Text = ITTSU1.KRAPKCounter.ToString();
            txtITTSU1KRAPKB.Text = ITTSU1KRAPKBCounter.ToString();
            txtITTSU1KRAPKM.Text = ITTSU1KRAPKMCounter.ToString();
            txtITTSU1KRAPKS.Text = ITTSU1KRAPKSCounter.ToString();
            txtITTSU1KRAPKSUR.Text = ITTSU1KRAPKSURCounter.ToString();
            txtITTSU1KRAPKUO.Text = ITTSU1KRAPKUOCounter.ToString();
            txtITTSU1KRAPKUMO.Text = ITTSU1KRAPKUMOCounter.ToString();

            txtITTSU2KRAPK.Text = ITTSU2KRAPKCounter.ToString();
            txtITTSU2KRAPKB.Text = ITTSU2KRAPKBCounter.ToString();
            txtITTSU2KRAPKM.Text = ITTSU2KRAPKMCounter.ToString();
            txtITTSU2KRAPKS.Text = ITTSU2KRAPKSCounter.ToString();
            txtITTSU2KRAPKSUR.Text = ITTSU2KRAPKSURCounter.ToString();
            txtITTSU2KRAPKUO.Text = ITTSU2KRAPKUOCounter.ToString();
            txtITTSU2KRAPKUMO.Text = ITTSU2KRAPKUMOCounter.ToString();

            txtIUIT1KRAPK.Text = IUIT1KRAPKCounter.ToString();
            txtIUIT2KRAPK.Text = IUIT2KRAPKCounter.ToString();

            txtVF1KRAPK.Text = VF1KRAPKCounter.ToString();
            txtVF2KRAPK.Text = VF2KRAPKCounter.ToString();

            txtSum1KRAPK.Text = Sum1KRAPKCounter.ToString();
            txtSum2KRAPK.Text = Sum2KRAPKCounter.ToString();

            txtAllKRAPK.Text = AllKRAPKCounter.ToString();
            //--------------------------Девятый столбец значений

            //--------------------------Десятый столбец значений
            txtITTSU1Kurs.Text = ITTSU1.KursCounter.ToString();
            txtITTSU1KursB.Text = ITTSU1KursBCounter.ToString();
            txtITTSU1KursM.Text = ITTSU1KursMCounter.ToString();
            txtITTSU1KursS.Text = ITTSU1KursSCounter.ToString();
            txtITTSU1KursSUR.Text = ITTSU1KursSURCounter.ToString();
            txtITTSU1KursUO.Text = ITTSU1KursUOCounter.ToString();
            txtITTSU1KursUMO.Text = ITTSU1KursUMOCounter.ToString();

            txtITTSU2Kurs.Text = ITTSU2KursCounter.ToString();
            txtITTSU2KursB.Text = ITTSU2KursBCounter.ToString();
            txtITTSU2KursM.Text = ITTSU2KursMCounter.ToString();
            txtITTSU2KursS.Text = ITTSU2KursSCounter.ToString();
            txtITTSU2KursSUR.Text = ITTSU2KursSURCounter.ToString();
            txtITTSU2KursUO.Text = ITTSU2KursUOCounter.ToString();
            txtITTSU2KursUMO.Text = ITTSU2KursUMOCounter.ToString();

            txtIUIT1Kurs.Text = IUIT1KursCounter.ToString();
            txtIUIT2Kurs.Text = IUIT2KursCounter.ToString();

            txtVF1Kurs.Text = VF1KursCounter.ToString();
            txtVF2Kurs.Text = VF2KursCounter.ToString();

            txtSum1Kurs.Text = Sum1KursCounter.ToString();
            txtSum2Kurs.Text = Sum2KursCounter.ToString();

            txtAllKurs.Text = AllKursCounter.ToString();
            //--------------------------Десятый столбец значений

            //--------------------------Одиннадцатый столбец значений
            txtITTSU1PreD.Text = ITTSU1.PreDCounter.ToString();
            txtITTSU1PreDB.Text = ITTSU1PreDBCounter.ToString();
            txtITTSU1PreDM.Text = ITTSU1PreDMCounter.ToString();
            txtITTSU1PreDS.Text = ITTSU1PreDSCounter.ToString();
            txtITTSU1PreDSUR.Text = ITTSU1PreDSURCounter.ToString();
            txtITTSU1PreDUO.Text = ITTSU1PreDUOCounter.ToString();
            txtITTSU1PreDUMO.Text = ITTSU1PreDUMOCounter.ToString();

            txtITTSU2PreD.Text = ITTSU2PreDCounter.ToString();
            txtITTSU2PreDB.Text = ITTSU2PreDBCounter.ToString();
            txtITTSU2PreDM.Text = ITTSU2PreDMCounter.ToString();
            txtITTSU2PreDS.Text = ITTSU2PreDSCounter.ToString();
            txtITTSU2PreDSUR.Text = ITTSU2PreDSURCounter.ToString();
            txtITTSU2PreDUO.Text = ITTSU2PreDUOCounter.ToString();
            txtITTSU2PreDUMO.Text = ITTSU2PreDUMOCounter.ToString();

            txtIUIT1PreD.Text = IUIT1PreDCounter.ToString();
            txtIUIT2PreD.Text = IUIT2PreDCounter.ToString();

            txtVF1PreD.Text = VF1PreDCounter.ToString();
            txtVF2PreD.Text = VF2PreDCounter.ToString();

            txtSum1PreD.Text = Sum1PreDCounter.ToString();
            txtSum2PreD.Text = Sum2PreDCounter.ToString();

            txtAllPreD.Text = AllPreDCounter.ToString();
            //--------------------------Одиннадцатый столбец значений

            //--------------------------Двенадцатый столбец значений
            txtITTSU1Diploma.Text = ITTSU1.DiplomaCounter.ToString();
            txtITTSU1DiplomaB.Text = ITTSU1DiplomaBCounter.ToString();
            txtITTSU1DiplomaM.Text = ITTSU1DiplomaMCounter.ToString();
            txtITTSU1DiplomaS.Text = ITTSU1DiplomaSCounter.ToString();
            txtITTSU1DiplomaSUR.Text = ITTSU1DiplomaSURCounter.ToString();
            txtITTSU1DiplomaUO.Text = ITTSU1DiplomaUOCounter.ToString();
            txtITTSU1DiplomaUMO.Text = ITTSU1DiplomaUMOCounter.ToString();

            txtITTSU2Diploma.Text = ITTSU2DiplomaCounter.ToString();
            txtITTSU2DiplomaB.Text = ITTSU2DiplomaBCounter.ToString();
            txtITTSU2DiplomaM.Text = ITTSU2DiplomaMCounter.ToString();
            txtITTSU2DiplomaS.Text = ITTSU2DiplomaSCounter.ToString();
            txtITTSU2DiplomaSUR.Text = ITTSU2DiplomaSURCounter.ToString();
            txtITTSU2DiplomaUO.Text = ITTSU2DiplomaUOCounter.ToString();
            txtITTSU2DiplomaUMO.Text = ITTSU2DiplomaUMOCounter.ToString();

            txtIUIT1Diploma.Text = IUIT1DiplomaCounter.ToString();
            txtIUIT2Diploma.Text = IUIT2DiplomaCounter.ToString();

            txtVF1Diploma.Text = VF1DiplomaCounter.ToString();
            txtVF2Diploma.Text = VF2DiplomaCounter.ToString();

            txtSum1Diploma.Text = Sum1DiplomaCounter.ToString();
            txtSum2Diploma.Text = Sum2DiplomaCounter.ToString();

            txtAllDiploma.Text = AllDiplomaCounter.ToString();
            //--------------------------Двенадцатый столбец значений

            //--------------------------Тринадцатый столбец значений
            txtITTSU1TutPr.Text = ITTSU1.TutPrCounter.ToString();
            txtITTSU1TutPrB.Text = ITTSU1TutPrBCounter.ToString();
            txtITTSU1TutPrM.Text = ITTSU1TutPrMCounter.ToString();
            txtITTSU1TutPrS.Text = ITTSU1TutPrSCounter.ToString();
            txtITTSU1TutPrSUR.Text = ITTSU1TutPrSURCounter.ToString();
            txtITTSU1TutPrUO.Text = ITTSU1TutPrUOCounter.ToString();
            txtITTSU1TutPrUMO.Text = ITTSU1TutPrUMOCounter.ToString();

            txtITTSU2TutPr.Text = ITTSU2TutPrCounter.ToString();
            txtITTSU2TutPrB.Text = ITTSU2TutPrBCounter.ToString();
            txtITTSU2TutPrM.Text = ITTSU2TutPrMCounter.ToString();
            txtITTSU2TutPrS.Text = ITTSU2TutPrSCounter.ToString();
            txtITTSU2TutPrSUR.Text = ITTSU2TutPrSURCounter.ToString();
            txtITTSU2TutPrUO.Text = ITTSU2TutPrUOCounter.ToString();
            txtITTSU2TutPrUMO.Text = ITTSU2TutPrUMOCounter.ToString();

            txtIUIT1TutPr.Text = IUIT1TutPrCounter.ToString();
            txtIUIT2TutPr.Text = IUIT2TutPrCounter.ToString();

            txtVF1TutPr.Text = VF1TutPrCounter.ToString();
            txtVF2TutPr.Text = VF2TutPrCounter.ToString();

            txtSum1TutPr.Text = Sum1TutPrCounter.ToString();
            txtSum2TutPr.Text = Sum2TutPrCounter.ToString();

            txtAllTutPr.Text = AllTutPrCounter.ToString();
            //--------------------------Тринадцатый столбец значений

            //--------------------------Четырнадцатый столбец значений
            txtITTSU1Prod.Text = ITTSU1.ProdCounter.ToString();
            txtITTSU1ProdB.Text = ITTSU1ProdBCounter.ToString();
            txtITTSU1ProdM.Text = ITTSU1ProdMCounter.ToString();
            txtITTSU1ProdS.Text = ITTSU1ProdSCounter.ToString();
            txtITTSU1ProdSUR.Text = ITTSU1ProdSURCounter.ToString();
            txtITTSU1ProdUO.Text = ITTSU1ProdUOCounter.ToString();
            txtITTSU1ProdUMO.Text = ITTSU1ProdUMOCounter.ToString();

            txtITTSU2Prod.Text = ITTSU2ProdCounter.ToString();
            txtITTSU2ProdB.Text = ITTSU2ProdBCounter.ToString();
            txtITTSU2ProdM.Text = ITTSU2ProdMCounter.ToString();
            txtITTSU2ProdS.Text = ITTSU2ProdSCounter.ToString();
            txtITTSU2ProdSUR.Text = ITTSU2ProdSURCounter.ToString();
            txtITTSU2ProdUO.Text = ITTSU2ProdUOCounter.ToString();
            txtITTSU2ProdUMO.Text = ITTSU2ProdUMOCounter.ToString();

            txtIUIT1Prod.Text = IUIT1ProdCounter.ToString();
            txtIUIT2Prod.Text = IUIT2ProdCounter.ToString();

            txtVF1Prod.Text = VF1ProdCounter.ToString();
            txtVF2Prod.Text = VF2ProdCounter.ToString();

            txtSum1Prod.Text = Sum1ProdCounter.ToString();
            txtSum2Prod.Text = Sum2ProdCounter.ToString();

            txtAllProd.Text = AllProdCounter.ToString();
            //--------------------------Четырнадцатый столбец значений

            //--------------------------Пятнадцатый столбец значений
            txtITTSU1GAK.Text = ITTSU1.GAKCounter.ToString();
            txtITTSU1GAKB.Text = ITTSU1GAKBCounter.ToString();
            txtITTSU1GAKM.Text = ITTSU1GAKMCounter.ToString();
            txtITTSU1GAKS.Text = ITTSU1GAKSCounter.ToString();
            txtITTSU1GAKSUR.Text = ITTSU1GAKSURCounter.ToString();
            txtITTSU1GAKUO.Text = ITTSU1GAKUOCounter.ToString();
            txtITTSU1GAKUMO.Text = ITTSU1GAKUMOCounter.ToString();

            txtITTSU2GAK.Text = ITTSU2GAKCounter.ToString();
            txtITTSU2GAKB.Text = ITTSU2GAKBCounter.ToString();
            txtITTSU2GAKM.Text = ITTSU2GAKMCounter.ToString();
            txtITTSU2GAKS.Text = ITTSU2GAKSCounter.ToString();
            txtITTSU2GAKSUR.Text = ITTSU2GAKSURCounter.ToString();
            txtITTSU2GAKUO.Text = ITTSU2GAKUOCounter.ToString();
            txtITTSU2GAKUMO.Text = ITTSU2GAKUMOCounter.ToString();

            txtIUIT1GAK.Text = IUIT1GAKCounter.ToString();
            txtIUIT2GAK.Text = IUIT2GAKCounter.ToString();

            txtVF1GAK.Text = VF1GAKCounter.ToString();
            txtVF2GAK.Text = VF2GAKCounter.ToString();

            txtSum1GAK.Text = Sum1GAKCounter.ToString();
            txtSum2GAK.Text = Sum2GAKCounter.ToString();

            txtAllGAK.Text = AllGAKCounter.ToString();
            //--------------------------Пятнадцатый столбец значений

            //--------------------------Шестнадцатый столбец значений
            txtITTSU1Bud.Text = ITTSU1.BudCounter.ToString();
            txtITTSU1BudZ.Text = ITTSU1.BudZCounter.ToString();

            txtITTSU1BudB.Text = ITTSU1BudBCounter.ToString();
            txtITTSU1BudM.Text = ITTSU1BudMCounter.ToString();
            txtITTSU1BudS.Text = ITTSU1BudSCounter.ToString();
            txtITTSU1BudSUR.Text = ITTSU1BudSURCounter.ToString();
            txtITTSU1BudUO.Text = ITTSU1BudUOCounter.ToString();
            txtITTSU1BudUMO.Text = ITTSU1BudUMOCounter.ToString();

            txtITTSU1BudBZ.Text = ITTSU1BudBZCounter.ToString();
            txtITTSU1BudMZ.Text = ITTSU1BudMZCounter.ToString();
            txtITTSU1BudSZ.Text = ITTSU1BudSZCounter.ToString();
            txtITTSU1BudSURZ.Text = ITTSU1BudSURZCounter.ToString();
            txtITTSU1BudUOZ.Text = ITTSU1BudUOZCounter.ToString();
            txtITTSU1BudUMOZ.Text = ITTSU1BudUMOZCounter.ToString();

            txtITTSU2Bud.Text = ITTSU2BudCounter.ToString();
            txtITTSU2BudZ.Text = ITTSU2BudZCounter.ToString();

            txtITTSU2BudB.Text = ITTSU2BudBCounter.ToString();
            txtITTSU2BudM.Text = ITTSU2BudMCounter.ToString();
            txtITTSU2BudS.Text = ITTSU2BudSCounter.ToString();
            txtITTSU2BudSUR.Text = ITTSU2BudSURCounter.ToString();
            txtITTSU2BudUO.Text = ITTSU2BudUOCounter.ToString();
            txtITTSU2BudUMO.Text = ITTSU2BudUMOCounter.ToString();

            txtITTSU2BudBZ.Text = ITTSU2BudBZCounter.ToString();
            txtITTSU2BudMZ.Text = ITTSU2BudMZCounter.ToString();
            txtITTSU2BudSZ.Text = ITTSU2BudSZCounter.ToString();
            txtITTSU2BudSURZ.Text = ITTSU2BudSURZCounter.ToString();
            txtITTSU2BudUOZ.Text = ITTSU2BudUOZCounter.ToString();
            txtITTSU2BudUMOZ.Text = ITTSU2BudUMOZCounter.ToString();

            txtIUIT1Bud.Text = IUIT1BudCounter.ToString();
            txtIUIT2Bud.Text = IUIT2BudCounter.ToString();

            txtIUIT1BudZ.Text = IUIT1BudZCounter.ToString();
            txtIUIT2BudZ.Text = IUIT2BudZCounter.ToString();

            txtVF1Bud.Text = VF1BudCounter.ToString();
            txtVF2Bud.Text = VF2BudCounter.ToString();

            txtVF1BudZ.Text = VF1BudZCounter.ToString();
            txtVF2BudZ.Text = VF2BudZCounter.ToString();

            txtSum1Bud.Text = Sum1BudCounter.ToString();
            txtSum2Bud.Text = Sum2BudCounter.ToString();

            txtSum1BudZ.Text = Sum1BudZCounter.ToString();
            txtSum2BudZ.Text = Sum2BudZCounter.ToString();

            txtAllBud.Text = AllBudCounter.ToString();
            txtAllBudZ.Text = AllBudZCounter.ToString();
            //--------------------------Шестнадцатый столбец значений

            //--------------------------Семнадцатый столбец значений
            txtITTSU1Sum.Text = ITTSU1.SumCounter.ToString();
            txtITTSU1SumZ.Text = ITTSU1.SumZCounter.ToString();

            txtITTSU1SumB.Text = ITTSU1SumBCounter.ToString();
            txtITTSU1SumM.Text = ITTSU1SumMCounter.ToString();
            txtITTSU1SumS.Text = ITTSU1SumSCounter.ToString();
            txtITTSU1SumSUR.Text = ITTSU1SumSURCounter.ToString();
            txtITTSU1SumUO.Text = ITTSU1SumUOCounter.ToString();
            txtITTSU1SumUMO.Text = ITTSU1SumUMOCounter.ToString();

            txtITTSU1SumBZ.Text = ITTSU1SumBZCounter.ToString();
            txtITTSU1SumMZ.Text = ITTSU1SumMZCounter.ToString();
            txtITTSU1SumSZ.Text = ITTSU1SumSZCounter.ToString();
            txtITTSU1SumSURZ.Text = ITTSU1SumSURZCounter.ToString();
            txtITTSU1SumUOZ.Text = ITTSU1SumUOZCounter.ToString();
            txtITTSU1SumUMOZ.Text = ITTSU1SumUMOZCounter.ToString();

            txtITTSU2Sum.Text = ITTSU2SumCounter.ToString();
            txtITTSU2SumZ.Text = ITTSU2SumCounter.ToString();

            txtITTSU2SumB.Text = ITTSU2SumBCounter.ToString();
            txtITTSU2SumM.Text = ITTSU2SumMCounter.ToString();
            txtITTSU2SumS.Text = ITTSU2SumSCounter.ToString();
            txtITTSU2SumSUR.Text = ITTSU2SumSURCounter.ToString();
            txtITTSU2SumUO.Text = ITTSU2SumUOCounter.ToString();
            txtITTSU2SumUMO.Text = ITTSU2SumUMOCounter.ToString();

            txtITTSU2SumBZ.Text = ITTSU2SumBZCounter.ToString();
            txtITTSU2SumMZ.Text = ITTSU2SumMZCounter.ToString();
            txtITTSU2SumSZ.Text = ITTSU2SumSZCounter.ToString();
            txtITTSU2SumSURZ.Text = ITTSU2SumSURZCounter.ToString();
            txtITTSU2SumUOZ.Text = ITTSU2SumUOZCounter.ToString();
            txtITTSU2SumUMOZ.Text = ITTSU2SumUMOZCounter.ToString();

            txtIUIT1Sum.Text = IUIT1SumCounter.ToString();
            txtIUIT2Sum.Text = IUIT2SumCounter.ToString();

            txtIUIT1SumZ.Text = IUIT1SumZCounter.ToString();
            txtIUIT2SumZ.Text = IUIT2SumZCounter.ToString();

            txtVF1Sum.Text = VF1SumCounter.ToString();
            txtVF2Sum.Text = VF2SumCounter.ToString();

            txtVF1SumZ.Text = VF1SumZCounter.ToString();
            txtVF2SumZ.Text = VF2SumZCounter.ToString();

            txtSum1Sum.Text = Sum1SumCounter.ToString();            
            txtSum2Sum.Text = Sum2SumCounter.ToString();

            txtSum1SumZ.Text = Sum1SumZCounter.ToString();
            txtSum2SumZ.Text = Sum2SumZCounter.ToString();
            
            txtAllSum.Text = AllSumCounter.ToString();
            txtAllSumZ.Text = AllSumZCounter.ToString();
            //--------------------------Семнадцатый столбец значений

            //--------------------------Восемнадцатый столбец значений
            txtITTSU1All.Text = ITTSU1.AllCounter.ToString();
            txtITTSU1AllZ.Text = ITTSU1.AllZCounter.ToString();

            txtITTSU1AllB.Text = ITTSU1AllBCounter.ToString();
            txtITTSU1AllM.Text = ITTSU1AllMCounter.ToString();
            txtITTSU1AllS.Text = ITTSU1AllSCounter.ToString();
            txtITTSU1AllSUR.Text = ITTSU1AllSURCounter.ToString();
            txtITTSU1AllUO.Text = ITTSU1AllUOCounter.ToString();
            txtITTSU1AllUMO.Text = ITTSU1AllUMOCounter.ToString();

            txtITTSU1AllBZ.Text = ITTSU1AllBZCounter.ToString();
            txtITTSU1AllMZ.Text = ITTSU1AllMZCounter.ToString();
            txtITTSU1AllSZ.Text = ITTSU1AllSZCounter.ToString();
            txtITTSU1AllSURZ.Text = ITTSU1AllSURZCounter.ToString();
            txtITTSU1AllUOZ.Text = ITTSU1AllUOZCounter.ToString();
            txtITTSU1AllUMOZ.Text = ITTSU1AllUMOZCounter.ToString();

            txtITTSU2All.Text = ITTSU2AllCounter.ToString();
            txtITTSU2AllZ.Text = ITTSU2AllZCounter.ToString();

            txtITTSU2AllB.Text = ITTSU2AllBCounter.ToString();
            txtITTSU2AllM.Text = ITTSU2AllMCounter.ToString();
            txtITTSU2AllS.Text = ITTSU2AllSCounter.ToString();
            txtITTSU2AllSUR.Text = ITTSU2AllSURCounter.ToString();
            txtITTSU2AllUO.Text = ITTSU2AllUOCounter.ToString();
            txtITTSU2AllUMO.Text = ITTSU2AllUMOCounter.ToString();

            txtITTSU2AllBZ.Text = ITTSU2AllBZCounter.ToString();
            txtITTSU2AllMZ.Text = ITTSU2AllMZCounter.ToString();
            txtITTSU2AllSZ.Text = ITTSU2AllSZCounter.ToString();
            txtITTSU2AllSURZ.Text = ITTSU2AllSURZCounter.ToString();
            txtITTSU2AllUOZ.Text = ITTSU2AllUOZCounter.ToString();
            txtITTSU2AllUMOZ.Text = ITTSU2AllUMOZCounter.ToString();

            txtIUIT1All.Text = IUIT1AllCounter.ToString();
            txtIUIT2All.Text = IUIT2AllCounter.ToString();

            txtIUIT1AllZ.Text = IUIT1AllZCounter.ToString();
            txtIUIT2AllZ.Text = IUIT2AllZCounter.ToString();

            txtVF1All.Text = VF1AllCounter.ToString();
            txtVF2All.Text = VF2AllCounter.ToString();

            txtVF1AllZ.Text = VF1AllZCounter.ToString();
            txtVF2AllZ.Text = VF2AllZCounter.ToString();

            txtSum1All.Text = Sum1AllCounter.ToString();            
            txtSum2All.Text = Sum2AllCounter.ToString();

            txtSum1AllZ.Text = Sum1AllZCounter.ToString();
            txtSum2AllZ.Text = Sum2AllZCounter.ToString();
            
            txtAllAll.Text = AllAllCounter.ToString();
            txtAllAllZ.Text = AllAllZCounter.ToString();
            //--------------------------Восемнадцатый столбец значений

            //--------------------------Аспирантура
            txtPostGrad1Sum.Text = EntPostGrad1Counter.ToString();
            txtPostGrad1.Text = SumPostGrad1Counter.ToString();
            txtPostGrad2Sum.Text = EntPostGrad2Counter.ToString();
            txtPostGrad2.Text = SumPostGrad2Counter.ToString();
            //--------------------------Аспирантура

            //--------------------------Магистратура
            txtMag1.Text = SumMagistry1Counter.ToString();
            txtMag2.Text = SumMagistry2Counter.ToString();
            txtAllMag1.Text = SumAllMagistry1Counter.ToString();
            txtAllMag2.Text = SumAllMagistry2Counter.ToString();
            //--------------------------Магистратура

            //--------------------------Посещение занятий
            txtVisiting1Sum.Text = EntVisiting1Counter.ToString();
            txtVisiting1.Text = SumVisiting1Counter.ToString();
            txtVisiting2Sum.Text = EntVisiting2Counter.ToString();
            txtVisiting2.Text = SumVisiting2Counter.ToString();
            //--------------------------Посещение занятий

            //--------------------------Полная сумма часов кафедры
            txtLastSumSum.Text = EntAllCounter.ToString();
            txtLastSumAll.Text = SumAllCounter.ToString();
            //--------------------------Полная сумма часов кафедры

            mdlData.colSummary.Add(ITTSU1);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            mdlSaveSummary.toSaveSummary();
        }
    }
}