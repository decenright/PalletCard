using System;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Configuration;
using System.Data;
using System.Collections.Generic;

namespace PalletCard
{
    public partial class Home : Form
    {
        public Home()
        {
            InitializeComponent();
        }

        private void Home_Load(object sender, EventArgs e)
        {
            //OdbcConnection conn = new OdbcConnection();
            //conn.ConnectionString = "Dsn=TharData;uid=tharuser";
            //try
            //{
            //    conn.Open();
            //    using (OdbcCommand com = new OdbcCommand("SELECT JobNo FROM app_PalletOperations", conn))
            //    {
            //        using (OdbcDataReader reader = com.ExecuteReader())
            //        {
            //            while (reader.Read())
            //            {
            //                string word = reader.GetString(0);
            //                // Word is from the database. Do something with it.
            //                label1.Text = reader.GetValue(0).ToString();
            //            }
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Failed to connect to data source");
            //}
            //finally
            //{
            //    conn.Close();
            //}


            string ConnectionString = Convert.ToString("Dsn=TharData;uid=tharuser");
            string CommandText = "SELECT * FROM app_PalletOperations";

            OdbcConnection myConnection = new OdbcConnection(ConnectionString);
            OdbcCommand myCommand = new OdbcCommand(CommandText, myConnection);

            OdbcDataAdapter myAdapter = new OdbcDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet myDataSet = new DataSet();
            try
            {
                myConnection.Open();
                myAdapter.Fill(myDataSet);
            }
            catch (Exception ex)
            {
                throw (ex);
            }
            finally
            {
                myConnection.Close();
            }

            using (DataTable dt = new DataTable())
            {
                myAdapter.Fill(dt);
                dataGridView1.DataSource = dt;
            }
        }
    }
}
