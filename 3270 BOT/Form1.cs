using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using AutoIt;

namespace _3270_BOT
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Panel_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.Left += e.X - lastPoint.X;
                this.Top += e.Y - lastPoint.Y;
            }
        }
        Point lastPoint;
        private void Panel_MouseDown(object sender, MouseEventArgs e)
        {
            lastPoint = new Point(e.X, e.Y);
        }

        public string retrievePassword()
        {


            using (OleDbConnection connection = new OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=\\\\172.23.16.70\\Consumer_Product\\CCCKL\\Malaysia Operations\\For Internal Use Only\\MIS Unit\\Yusri's File\\BTCX\\CardlinkPassword.MDB;"))

            {
                try
                {

                    OleDbCommand command = new OleDbCommand("SELECT * FROM Credential", connection);

                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                    {
                        return reader["CREATES"].ToString();

                    }

                    reader.Close();

                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                connection.Close();


            }

            return null;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            TimeSpan time = DateTime.Now.TimeOfDay;
            string yyyy = DateTime.Now.ToString("yyyy").Replace('-', '/');
            string mmmm = DateTime.Now.ToString("MMMM").Replace('-', '/');
            string MM = DateTime.Now.ToString("MM").Replace('-', '/');


            if (time > new TimeSpan(06, 25, 00)        //Hours, Minutes, Seconds
             && time < new TimeSpan(06, 25, 30))
            {

                try
                {

                    System.Diagnostics.Process.Start(@"3270 BOT.exe");
                    DAIChecking dai = new DAIChecking();
                    dai.starter("");

                }
                catch (Exception a)
                {


                }



                //DAIChecking dai = new DAIChecking();
                //dai.starter();

            }
            else if (time > new TimeSpan(19, 26, 00)        //Hours, Minutes, Seconds
            && time < new TimeSpan(19, 26, 30))
            {
                try
                {

                    //System.Diagnostics.Process.Start(@"3270 BOT.exe");
                    //DAIChecking dai = new DAIChecking();
                    //dai.starter("bbb");

                }
                catch (Exception a)
                {


                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //DAIChecking dai = new DAIChecking();
            //dai.starter();

        }
    }
}
