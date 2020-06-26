using System;
using System.Windows.Forms;
using System.Linq;
using AutoIt;
using System.Data.Odbc;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing;
using System.Data.OleDb;
using Microsoft.VisualBasic;


namespace _3270_BOT
{
    class ODBCConnections
    {
        string queryString;
        string ConnectionString;
        string dataInquiry;

        public string gettingDataFromODBC(string ConnectionString, string queryString, string inputs)
        {

            try
            {
                using (OdbcConnection connection = new OdbcConnection(ConnectionString))
                {

                    OdbcCommand command = new OdbcCommand(queryString, connection);

                    connection.Open();

                    // Execute the DataReader and access the data.
                    OdbcDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                    {

                        return dataInquiry = reader[inputs].ToString();

                    }

                    reader.Close();
                    connection.Close();

                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);

            }

            return null;
        }

    }


}

