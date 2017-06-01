using System;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data;

namespace PalletCard
{
    public partial class Home : Form
    {
        int    numberUp, jobGanged, paperSectionNo, heightMM, invoiceCustomerCode, qtyRequired;
        string jobNo, resourceID, name, expr1, id, workingSize, description, code, jobDesc, invoiceCustomerName, ref7;
        bool jobCompleted, jobCancelled;

        public Home()
        {
            InitializeComponent();
        }

        private void Home_Load(object sender, EventArgs e)
        {
            string ConnectionString = Convert.ToString("Dsn=TharData;uid=tharuser");
            string CommandText = "SELECT * FROM app_PalletOperations";

            OdbcConnection myConnection = new OdbcConnection(ConnectionString);
            OdbcCommand myCommand = new OdbcCommand(CommandText, myConnection);

            OdbcDataAdapter myAdapter = new OdbcDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet tharData = new DataSet();
            try
            {
                myConnection.Open();
                myAdapter.Fill(tharData);
            }
            catch (Exception ex)
            {
                throw (ex);
            }
            finally
            {
                myConnection.Close();
            }

            using (DataTable operations = new DataTable())
            {
                myAdapter.Fill(operations);
                dataGridView1.DataSource = operations;
            }
        }
//SEARCH______________________________________________________________________________________________________________________

        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("JobNo like '%{0}%'", searchBox.Text.Trim().Replace("'", "''"));
                lblJobNo.Text = dataGridView1.Rows[0].Cells[0].Value.ToString();
                lblJobNo.Visible = true;

                int resourceID = (int)dataGridView1.Rows[0].Cells[1].Value;
                if (resourceID == 5)
                {
                    lblPress.Text = "710UV";
                    lblPress.Visible = true;
                    getOperationsData();
                }
                else
                {
                    lblPress.Visible = false;
                    MessageBox.Show("The Job number you entered is not on this press");
                }
            }
            catch (Exception) { }
        }

        private void getOperationsData()
        {
            jobNo = dataGridView1.Rows[0].Cells[0].Value.ToString();
            resourceID = dataGridView1.Rows[0].Cells[1].Value.ToString();
            name = dataGridView1.Rows[0].Cells[2].Value.ToString();
            //MessageBox.Show(jobNo); MessageBox.Show(resourceID); MessageBox.Show(name);

            if  (!(string.IsNullOrEmpty(name)))
            {
                btnName.Visible = true;
                btnName.Text = name;
            }
        }

    }
}
