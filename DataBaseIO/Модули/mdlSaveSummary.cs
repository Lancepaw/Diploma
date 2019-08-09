using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;

namespace DataBaseIO
{
    class mdlSaveSummary
    {
        public static void toSaveSummary()
        {
            //int index;

            string TabName;
            string TabAttr;
            string TabVal;

            //Формируем SQL-запрос для поиска заданной таблицы
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            //Создаём переменную SQL-запроса
            OleDbCommand command = new OleDbCommand();
            //Конструктор команд
            OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter);
            //
            mdlData.glConn.Open();
            //Указываем соединение для переменной SQL-запроса
            command.Connection = mdlData.glConn;

            TabName = mdlBaseStructure.masTabNames[25][0][0];
            TabAttr = mdlBaseStructure.getTabAttributes(mdlBaseStructure.masTabNames[25][3]);

            TabVal = "";

            command.CommandText = "DELETE FROM [" + TabName + "] WHERE [Код] > 0";
            command.ExecuteNonQuery();

            for (int j = 0; j < mdlData.colSummary.Count; j++)
            {
                for (int i = 0; i <= mdlBaseStructure.masTabNames[25][3].Length - 1; i++)
                {
                    switch (i)
                    {
                        case 0:
                            {
                                TabVal += (j + 1).ToString() + ", ";
                                break;
                            }
                        case 3:
                            {
                                TabVal += mdlData.colSummary[j].LectCounter.ToString() + ", ";
                                break;
                            }
                        case 4:
                            {
                                TabVal += mdlData.colSummary[j].ExamCounter.ToString() + ", ";
                                break; 
                            }
                        case 5:
                            {
                                TabVal += mdlData.colSummary[j].CredCounter.ToString() + ", ";
                                break;
                            }
                        case 6:
                            {
                                TabVal += mdlData.colSummary[j].RefCounter.ToString() + ", ";
                                break;
                            }
                        case 7:
                            {
                                TabVal += mdlData.colSummary[j].TutCounter.ToString() + ", ";
                                break;
                            }
                        case 8:
                            {
                                TabVal += mdlData.colSummary[j].LabCounter.ToString() + ", ";
                                break;
                            }
                        case 9:
                            {
                                TabVal += mdlData.colSummary[j].PractCounter.ToString() + ", ";
                                break;
                            }
                        case 10:
                            {
                                TabVal += mdlData.colSummary[j].IndCounter.ToString() + ", ";
                                break;
                            }
                        case 11:
                            {
                                TabVal += mdlData.colSummary[j].KRAPKCounter.ToString() + ", ";
                                break;
                            }
                        case 12:
                            {
                                TabVal += mdlData.colSummary[j].KursCounter.ToString() + ", ";
                                break;
                            }
                        case 13:
                            {
                                TabVal += mdlData.colSummary[j].PreDCounter.ToString() + ", ";
                                break;
                            }
                        case 14:
                            {
                                TabVal += mdlData.colSummary[j].DiplomaCounter.ToString() + ", ";
                                break;
                            }
                        case 15:
                            {
                                TabVal += mdlData.colSummary[j].TutPrCounter.ToString() + ", ";
                                break;
                            }
                        case 16:
                            {
                                TabVal += mdlData.colSummary[j].ProdCounter.ToString() + ", ";
                                break;
                            }
                        case 17:
                            {
                                TabVal += mdlData.colSummary[j].GAKCounter.ToString() + ", ";
                                break;
                            }
                        case 18:
                            {
                                TabVal += mdlData.colSummary[j].BudCounter.ToString() + ", ";
                                break;
                            }
                        case 19:
                            {
                                TabVal += mdlData.colSummary[j].SumCounter.ToString() + ", ";
                                break;
                            }
                        case 20:
                            {
                                TabVal += mdlData.colSummary[j].AllCounter.ToString() + ", ";
                                break;
                            }
                        case 21:
                            {
                                TabVal += mdlData.colSummary[j].BudZCounter.ToString() + ", ";
                                break;
                            }
                        case 22:
                            {
                                TabVal += mdlData.colSummary[j].SumZCounter.ToString() + ", ";
                                break;
                            }
                        case 23:
                            {
                                TabVal += mdlData.colSummary[j].AllZCounter.ToString();
                                break;
                            }
                        default:
                            {
                                TabVal += "1, ";
                                break;
                            }
                    }
                }
            }

            command.CommandText = "INSERT INTO [" + TabName + "] (" + TabAttr + ") VALUES (" + TabVal + ")";
            command.ExecuteNonQuery();

            mdlData.glConn.Close();
        }
    }
}
