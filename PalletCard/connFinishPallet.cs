using System;
using System.Data;
using System.Data.Odbc;
using System.Windows.Forms;

namespace PalletCard
{
    public class ConnFinishPallet
    {
        private OdbcConnection myConnection;
        private OdbcDataAdapter myAdapter;
        private DataTable palletCardLog;
        public OdbcCommand myCommand;

        //public ConnFinishPallet()
        //{
        //    string ConnectionString = Convert.ToString("Dsn=PalletCard;uid=PalletCardAdmin");
        //    string CommandText = "SELECT * FROM Log where AutoNum = '" + Home.tbxFinishPallet.Text + "'";
        //    OdbcConnection myConnection = new OdbcConnection(ConnectionString);
        //    OdbcCommand myCommand = new OdbcCommand(CommandText, myConnection);
        //    OdbcDataAdapter myAdapter = new OdbcDataAdapter();
        //    myAdapter.SelectCommand = myCommand;
        //    DataSet palletCardData = new DataSet();
        //    try
        //    {
        //        myConnection.Open();
        //        myAdapter.Fill(palletCardData);
        //        using (DataTable palletCardLog = new DataTable())
        //        {
        //            myAdapter.Fill(palletCardLog);
        //            dataGridView2.DataSource = palletCardLog;
        //        }
        //        pnlPalletCard3.BringToFront();

        //    }
        //    catch (Exception ex1)
        //    {
        //        MessageBox.Show(ex1.Message);
        //    }
        //    finally
        //    {
        //        myConnection.Close();
        //    }
        //}




    }
}
