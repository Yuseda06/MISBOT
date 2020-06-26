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
    public class DAIChecking
    {
        Form1 form = new Form1();
        GeneralEmail email = new GeneralEmail();
        public static int countTime;
        string daiDataNext;
        string currentTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        string currentTimeLess1Hour = DateTime.Now.AddHours(-1).ToString("yyyy-MM-dd HH:mm:ss");

        public DAIChecking()
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


        CXChecking cx = new CXChecking();

        public void insertDataCheck()
        {
            starter("");
        }

        public void validateDAI()
        {
            using (OleDbConnection connection = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + ConnString + "DAIData.mdb;"))

            {
                try
                {
                    DateTime now = DateTime.Now;
                    //reportDate = "06022020";
                    string reportDate = now.ToString("ddMMyyyy");
                    OleDbCommand command = new OleDbCommand("SELECT * FROM DAIDATA WHERE Description = 'DAI' AND Status = 0", connection);

                    // OleDbCommand command = new OleDbCommand("SELECT * FROM DAIDATA WHERE ExtractDate = '" + reportDate + "'", connection);
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                    {
                        checkPlan(Convert.ToInt32(reader["Plan"]), Convert.ToDouble(reader["TrxAmt"]), Convert.ToInt32(reader["ID"]));

                    }

                    reader.Close();

                }
                catch (Exception ex)
                {
                    //MessageBox.Show("Error" + ex.Message);
                    Console.WriteLine(ex.Message);
                }

                connection.Close();

            }

        }

        public void checkPlan(int plan, double amount, int ID)
        {

            ODBCConnections odbc = new ODBCConnections();

            double amountPlan = Convert.ToDouble(odbc.gettingDataFromODBC(@"DSN=DAIPLAN;DBQ=" + connODBC + "DAIPLAN.xls;DefaultDir=" + connODBC + ";DriverId=790;FIL=excel 8.0;MaxBufferSize=2048;PageTimeout=5;", @"SELECT `PLAN$`.AMOUNT FROM `" + queryODBC + "DAIPLAN`.`PLAN$` `PLAN$` WHERE (`PLAN$`.PLAN=" + plan + ")", "AMOUNT"));

            if (amountPlan == 0)
            {
                //email.sendEmail("New DAI Plan " + plan + "", "", "179264", "ID " + ID + "");
            }

            if (amount < amountPlan)
            {
                updateDAIData(ID);
            }

        }



        string skillset;
        string tenureCSA;
        string teamManager;
        string unitHead;
        string tmID;
        string uhID;
        string staffName;
        string date2;

        public void gettingDAIData(int ID)
        {
            using (OleDbConnection connection = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + ConnString + "DAIData.mdb;"))

            {
                try
                {
                    OleDbCommand command = new OleDbCommand("SELECT * FROM DAIDATA WHERE ID = " + ID + "", connection);
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                    {

                        cardNumber = reader["CardNo"].ToString();
                        tranAmount = reader["TrxAmt"].ToString();
                        monInstall = reader["InstallAmt"].ToString();
                        tenure = Convert.ToInt32(reader["Term"]);
                        interest = reader["InterestRate"].ToString();
                        reportDate = reader["ReportDate"].ToString();
                        lastPay = reader["LastDate"].ToString();
                        tranDate = reader["TrxDate"].ToString();
                        plan = Convert.ToInt32(reader["Plan"]);
                        staffID = reader["StaffID"].ToString();

                    }

                    reader.Close();

                }
                catch (Exception ex)
                {
                    //MessageBox.Show("Error" + ex.Message);
                    Console.WriteLine(ex.Message);
                }

                connection.Close();



                using (OleDbConnection connection1 = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + ConnString + "AgentList.mdb;"))

                {
                    try
                    {
                        DateTime now = DateTime.Now;
                        date2 = now.Date.ToString("MMMM");

                        OleDbCommand command1 = new OleDbCommand("SELECT DISTINCT  * FROM " + date2 + " WHERE(`Staff ID`= " + staffID + ")   ", connection1);
                        OleDbDataReader reader1;
                        connection1.Open();
                        reader1 = command1.ExecuteReader();

                        while (reader1.Read())
                        {

                            staffName = reader1["WFM"].ToString();
                            skillset = reader1["Skillset"].ToString();
                            tenureCSA = reader1["Working Status"].ToString();
                            teamManager = reader1["Team Manager"].ToString();
                            unitHead = reader1["Unit Head"].ToString();
                            tmID = reader1["TM ID"].ToString();
                            uhID = reader1["UH ID"].ToString();

                        }

                        reader1.Close();

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                    connection1.Close();


                }



                string table = "<tr>" +
                    "<th >Card Number</th>" +
                    "<th >Tranx Amount</th>" +
                    "<th >Monthly Installment</th>" +
                    "<th >Tenure</th>" +
                    "<th >Interest</th>" +
                    "<th >Report Date</th>" +
                    "<th >Last Pay</th>" +
                    "<th >Tranx Date</th>" +
                    "<th >Plan</th>" +
                    "<th >Staff ID</th>" +
                    "</tr>" +
                    "<tr>" +
                        "<td>" + cardNumber + "</td>" +
                        "<td  style=\"border:0.5px dotted  Tomato;font-weight: bold; color: Red;\"  >" + tranAmount + "</td>" +
                        "<td>" + monInstall + "</td>" +
                        "<td>" + tenure + "</td>" +
                        "<td>" + interest + "</td>" +
                        "<td>" + reportDate + "</td>" +
                        "<td>" + lastPay + "</td>" +
                        "<td>" + tranDate + "</td>" +
                        "<td>" + plan + "</td>" +
                        "<td>" + staffID + "</td>" +
                    "</tr>";

                string bodies =
                        "<!DOCTYPE html>" +
                        "<html>" +
                        "<head>" +
                        "<style>" +
                        //"#customers {font-family: \"Trebuchet MS\"; font-size: 12px; border - collapse: collapse;width: 100" +
                        //"#customers td  {border:0.5px solid rgb(180, 180, 180); padding: 0.5px; text-align: center}" +
                        //"#customers tr:nth-child(even){background-color: #ffff;}" +
                        //"#customers tr:hover {background-color: #ddd;}" +
                        //"#customers th {border:0.5px solid rgb(180, 180, 180); padding: 0.5px; text-align: center ;background-color: #0099e6;color: white;}" +

                        "#customers {font-family: \"Trebuchet MS\"; font-size: 12px; border - collapse: collapse;width: 100%;}" +
                        "#customers td, #customers th {border: 1px solid #ddd;padding: 0.5px;text-align: center;}" +
                        "#customers tr:nth-child(even){background-color: #FFFFFF;}" +
                        "#customers tr:hover {background-color: #FFFFFF;}" +
                        "#customers th {padding-top: 0.5px;padding-bottom: 0.5px;text-align: center;background-color: #0099e6;color: white;}" +

                        "</style>" +
                        "</head>" +
                        "<body style=\"font-family:Trebuchet MS;font-size: 12px;\">" +
                        "<p>Dear " + staffName + ",</p>" +
                        Environment.NewLine +
                        "<p>Please be informed that you have wrongly performed the maintenance for DAI as below details.<BR>" +
                        Environment.NewLine +
                        "Please perform the necessary action for the rectification." +
                        Environment.NewLine + "</p>" +
                        "<table id=\"customers\">" + table + "</table>" +
                        Environment.NewLine +
                        "<p>Regards,<br>Resource & Capacity Management</p>" +
                        "</body>" +
                        "</html>";

                email.sendEmail("[DAI] Wrong Maintenance Performed  " + cardNumber + "", tmID + ";" + uhID + ";" + staffID, "179264;183813;409035;407074;", bodies);
                // email.sendEmail("[DAI] Wrong Maintenance Performed  " + cardNumber + "","", "179264", bodies);


            }

        }

        private void updateDAIData(int ID)
        {
            using (OleDbConnection connection = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + ConnString + "DAIData.mdb;"))

                //using (OleDbConnection connection = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=\\\\B720W999\\BTCX\\BTCXData.mdb;"))

                try
                {
                    connection.Open();

                    string my_querry = "UPDATE DAIDATA SET Status = 1 WHERE ID = " + ID + " ";
                    OleDbCommand cmd = new OleDbCommand(my_querry, connection);
                    cmd.ExecuteNonQuery();
                    connection.Close();
                    gettingDAIData(ID);


                }
                catch (Exception e)
                {

                }

        }

        public void starter(string aaa)
        {
            AutoItX.WinActivate("3270 BOT");
            System.Diagnostics.Process.Start(@"D:\Program Files\Reflection\User\mainfrm.rsf");
            AutoItX.WinWaitActive("Enter Host Name or IP Address");
            AutoItX.Send("172.24.2.2");
            AutoItX.ControlClick("Enter Host Name or IP Address", "", "Connect");
            AutoItX.Sleep(1000);
            AutoItX.Send("M");
            AutoItX.Send("{ENTER}");
            AutoItX.Sleep(1000);
            AutoItX.Send("R179264");
            AutoItX.Sleep(1000);
            AutoItX.Send("{TAB}");
            AutoItX.Sleep(1000);
            AutoItX.Send(form.retrievePassword());
            AutoItX.Sleep(1000);
            AutoItX.Send("{ENTER}");
            AutoItX.Sleep(1000);
            AutoItX.Send("CEPP");
            AutoItX.Sleep(1000);
            AutoItX.Send("{ENTER}");
            AutoItX.Sleep(1000);
            AutoItX.Send("X");
            AutoItX.Sleep(1000);
            AutoItX.Send("{ENTER}");
            AutoItX.Sleep(1000);
            AutoItX.Send("X");
            AutoItX.Sleep(1000);
            AutoItX.Send("{ENTER}");
            AutoItX.Sleep(1000);
            AutoItX.Send("X");
            AutoItX.Sleep(1000);
            AutoItX.Send("{ENTER}");
            AutoItX.Sleep(1000);
            AutoItX.Send("{DOWN 7}");
            AutoItX.Sleep(1000);
            AutoItX.Send("{RIGHT 28}");
            AutoItX.Sleep(1000);
            AutoItX.Send("+{RIGHT 3}");
            AutoItX.Sleep(1000);
            checkingDAI();

            CXChecking cx = new CXChecking();
            cx.insertDataCXManual();

            BTChecking bt = new BTChecking();
            bt.insertDataBTManual();

            validateDAI();

            AutoItX.WinActivate("Reflection - IBM 3270 Terminal - mainfrm.rsf");
            AutoItX.Send("!+{F4}", 0);


            //email.sendEmail("dai", "", "179264", aaa);
            AutoItX.WinActivate("3270 BOT");
            AutoItX.Send("!+{F4}", 0);

        }


        public void runMasterDAICXBT()
        {

            insertDataDaiManual();

            CXChecking cx = new CXChecking();
            cx.insertDataCXManual();

            BTChecking bt = new BTChecking();
            bt.insertDataBTManual();

            validateDAI();

            AutoItX.WinActivate("Reflection - IBM 3270 Terminal - mainfrm.rsf");
            AutoItX.Send("!+{F4}", 0);

        }




        public void insertDataDaiManual()
        {
            AutoItX.WinActivate("Reflection - IBM 3270 Terminal - mainfrm.rsf");
            AutoItX.Sleep(1000);
            AutoItX.Send("{DOWN 7}");
            AutoItX.Sleep(1000);
            AutoItX.Send("{RIGHT 28}");
            AutoItX.Sleep(1000);
            AutoItX.Send("+{RIGHT 3}");
            AutoItX.Sleep(1000);
            checkingDAI();


        }



        public void checkingDAI()
        {
            AutoItX.Send("{CTRLDOWN}c{CTRLUP}");
            AutoItX.Sleep(100);
            string DAIText = Clipboard.GetText();

            if (DAIText == "DIA")
            {

                AutoItX.Send("{DOWN 2}");
                AutoItX.Send("{LEFT 38}");

                //First copying DAI data for 17 rows (first page)
                for (int i = 0; i < 17; i++)
                {
                    daiDataNext = checkingNext(i);

                    if (daiDataNext == "-")
                    {
                        //cx.insertDataCXManual();
                        goto Finish;
                    }
                    else
                    {
                        AutoItX.Send("{CTRLDOWN}+{RIGHT 12}{CTRLUP}");
                        AutoItX.Send("{CTRLDOWN}c{CTRLUP}");
                        AutoItX.Sleep(1000);
                        string DAIData = Clipboard.GetText();
                        cuttingData(DAIData);
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
                        daiDataNext = checkingNext(j);

                        if (daiDataNext == "")
                        {
                            AutoItX.Send("{F8}");
                            AutoItX.Sleep(1000);
                            insertDataDaiManual();
                            goto Finish;
                        }
                        else if (daiDataNext == "-")
                        {
                            //cx.insertDataCXManual();
                            goto Finish;
                        }
                        else
                        {
                            AutoItX.Send("{CTRLDOWN}+{RIGHT 12}{CTRLUP}");
                            AutoItX.Send("{CTRLDOWN}c{CTRLUP}");
                            AutoItX.Sleep(1000);
                            string DAIData = Clipboard.GetText();
                            cuttingData(DAIData);
                        }
                    }
                }
            }
            else
            {

                AutoItX.Send("{F8}");
                AutoItX.Send("{DOWN 7}");
                AutoItX.Send("{RIGHT 28}");
                AutoItX.Send("+{RIGHT 3}");
                checkingDAI();

            }

        Finish:
            return;

        }

        string reportDate, cardNumber, lastPay, tranDate, staffID, tranAmount, monInstall, interest;

        int plan, tenure;

        DateTime extractDate;

        public void cuttingData(string data)
        {

            cardNumber = data.Substring(0, 16);
            tranAmount = data.Substring(16, 12).Trim();
            monInstall = data.Substring(29, 13).Trim();
            tenure = Convert.ToInt16(data.Substring(56, 3));
            interest = data.Substring(60, 12).Trim();
            reportDate = data.Substring(73, 9).Trim();

            string dd = reportDate.Substring(0, 2);
            string mm = reportDate.Substring(2, 2);
            string yy = reportDate.Substring(4, 4);

            extractDate = Convert.ToDateTime(dd + "/" + mm + "/" + yy);



            lastPay = data.Substring(82, 9).Trim();
            tranDate = data.Substring(91, 9).Trim();
            plan = Convert.ToInt32(data.Substring(116, 4));
            staffID = data.Substring(120, 6).Trim();

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

                    String my_querry = "INSERT INTO DAIDATA (CardNo, TrxAmt, InstallAmt, Term,  UnearnedInt, TrxDate, ReportDate, InterestRate, Plan, StaffID, LastDate, Status, ExtractDate, Description ) VALUES( '" + CardNo + "',          '" + TrxAmt + "',  '" + InstallAmt + "',          " + Term + ",      '" + UnearnedInt + "','" + TrxDate + "','" + ReportDate + "',  '" + InterestRate + "', " + Plan + ", '" + StaffID + "','" + LastDate + "',   " + Status + ", '" + extractDate + "','DAI')";

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
