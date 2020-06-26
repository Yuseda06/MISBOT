using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Odbc;
using System.ComponentModel;
using System.Drawing;
using System.Data.OleDb;
using AutoIt;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace _3270_BOT
{
    public class CXChecking
    {

        GeneralEmail email = new GeneralEmail();

        public static int countTime;
        string cxDataNext;
        string currentTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        string currentTimeLess1Hour = DateTime.Now.AddHours(-1).ToString("yyyy-MM-dd HH:mm:ss");
        BTChecking bt = new BTChecking();
        public CXChecking()
        {
            if (Environment.UserName.ToString() != "Yusri")
            {
                ConnString = "\\\\172.23.16.70\\Consumer_Product\\CCCKL\\Malaysia Operations\\For Internal Use Only\\MIS Unit\\Yusri's File\\BTCX\\";
                connODBC = "\\\\172.23.16.70\\Consumer_Product\\CCCKL\\Malaysia Operations\\For Internal Use Only\\MIS Unit\\Yusri's File\\BTCX\\";
                queryODBC = "\\\\172.23.16.70\\Consumer_Product\\CCCKL\\Malaysia Operations\\For Internal Use Only\\MIS Unit\\Yusri's File\\BTCX\\";
            }
            else
            {
                ConnString = "";
                connODBC = @"C:\Users\Yusri\Desktop\MISBot Yus - DAI\MISBot\bin\Debug\;";
                queryODBC = @"C:\Users\Yusri\Desktop\MISBot Yus - DAI\MISBot\bin\Debug\";
            }
        }


        public void insertDataCXManual()
        {
            AutoItX.WinActivate("Reflection - IBM 3270 Terminal - mainfrm.rsf");
            AutoItX.Send("{F8}");
            AutoItX.Sleep(1000);
            AutoItX.Send("{DOWN 7}");
            AutoItX.Sleep(1000);
            AutoItX.Send("{RIGHT 28}");
            AutoItX.Sleep(1000);
            AutoItX.Send("+{RIGHT 5}");
            AutoItX.Sleep(1000);
            checkingCX();
        }



        public void checkingCX()
        {
            AutoItX.Send("{CTRLDOWN}c{CTRLUP}");
            AutoItX.Sleep(1000);
            string CXText = Clipboard.GetText();


            if (CXText == "RHB C")
            {

                AutoItX.Send("{DOWN 2}");
                AutoItX.Send("{LEFT 38}");

                //First copying DAI data for 17 rows (first page)
                for (int i = 0; i < 17; i++)
                {
                    cxDataNext = checkingNext(i);
                    if (cxDataNext == "-")
                    {
                        //bt.insertDataBTManual();
                        goto Finish;
                    }
                    else
                    {
                        AutoItX.Send("{CTRLDOWN}+{RIGHT 13}{CTRLUP}");
                        AutoItX.Send("{CTRLDOWN}c{CTRLUP}");
                        AutoItX.Sleep(1000);
                        string text = Clipboard.GetText().Trim();
                        cuttingData(text);
                    }

                }

                //First copying DAI data for 24 rows (other pages)
                for (int i = 0; i < 20; i++)
                {
                    AutoItX.Send("{F8}");
                    AutoItX.Sleep(1000);
                    AutoItX.Send("{DOWN 2}");
                    AutoItX.Send("{LEFT 10}");


                    for (int j = 0; j < 24; j++)
                    {
                        cxDataNext = checkingNext(j);

                        if (cxDataNext == "")
                        {
                            //AutoItX.Send("{F8}");
                            //AutoItX.Sleep(1000);
                            insertDataCXManual();
                            goto Finish;

                        }
                        else if (cxDataNext == "-")
                        {
                            //bt.insertDataBTManual();
                            goto Finish;
                        }
                        else
                        {
                            AutoItX.Send("{CTRLDOWN}+{RIGHT 13}{CTRLUP}");
                            AutoItX.Send("{CTRLDOWN}c{CTRLUP}");
                            AutoItX.Sleep(1000);
                            string text = Clipboard.GetText().Trim();
                            cuttingData(text);
                        }
                    }
                }
            }
            else
            {

                AutoItX.Send("{F8}");
                AutoItX.Sleep(1000);
                AutoItX.Send("{DOWN 7}");
                AutoItX.Sleep(1000);
                AutoItX.Send("{RIGHT 28}");
                AutoItX.Sleep(1000);
                AutoItX.Send("+{RIGHT 5}");
                AutoItX.Sleep(1000);
                checkingCX();

            }
        Finish:
            return;
        }

        string reportDate, cardNumber, lastPay, tranDate, staffID, tranAmount, monInstall, interest;
        int plan, tenure;
        DateTime extractDate;


        public void cuttingData(string data)
        {

            try
            {
                cardNumber = data.Substring(0, 16);
                tranAmount = data.Substring(18, 11).Trim();
                monInstall = data.Substring(30, 12).Trim();
                tenure = Convert.ToInt16(data.Substring(56, 2));
                interest = data.Substring(62, 11).Trim();
                reportDate = data.Substring(73, 9).Trim();

                string dd = reportDate.Substring(0, 2);
                string mm = reportDate.Substring(2, 2);
                string yy = reportDate.Substring(4, 4);

                extractDate = Convert.ToDateTime(dd + "/" + mm + "/" + yy);

                lastPay = data.Substring(82, 9).Trim();
                tranDate = data.Substring(91, 9).Trim();
                plan = Convert.ToInt32(data.Substring(116, 3));
                staffID = data.Substring(120, 6).Trim();

            }
            catch (Exception e)
            {
                staffID = "";

            }


            insertDataMDB(cardNumber, tranAmount, monInstall, tenure, "0", tranDate, reportDate, interest, plan, staffID, lastPay, 0, extractDate);

        }

        public string checkingNext(int i)
        {
            if (i != 0)
            {
                AutoItX.Send("{DOWN}");
            }
            AutoItX.Send("+{RIGHT 1}");
            AutoItX.Send("{CTRLDOWN}c{CTRLUP}");
            AutoItX.Sleep(1000);
            return Clipboard.GetText();
        }


        string ConnString;
        string connODBC;
        string queryODBC;

        private void insertDataMDB(string CardNo, string TrxAmt, string InstallAmt, int Term, string UnearnedInt, string TrxDate, string ReportDate, string InterestRate, int Plan, string StaffID, string LastDate, int Status, DateTime extractDate)
        {


            using (OleDbConnection connection = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + ConnString + "DAIData.mdb;"))

                //using (OleDbConnection connection = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=\\\\B720W999\\BTCX\\BTCXData.mdb;"))

                try
                {
                    connection.Open();

                    String my_querry = "INSERT INTO DAIDATA (CardNo, TrxAmt,            InstallAmt,        Term,   UnearnedInt,            TrxDate,           ReportDate,       InterestRate,      Plan,     StaffID, LastDate,           Status, ExtractDate, Description ) VALUES( '" + CardNo + "',          '" + TrxAmt + "',  '" + InstallAmt + "',          " + Term + ",      '" + UnearnedInt + "','" + TrxDate + "','" + ReportDate + "',  '" + InterestRate + "', " + Plan + ", '" + StaffID + "','" + LastDate + "',   " + Status + ", '" + extractDate + "','Cash Excess')";

                    OleDbCommand cmd = new OleDbCommand(my_querry, connection);
                    cmd.ExecuteNonQuery();


                }
                catch (Exception ex)
                {
                    //MessageBox.Show("Failed due to" + ex.Message);
                }
                finally
                {
                    connection.Close();
                }


        }

    }
}
