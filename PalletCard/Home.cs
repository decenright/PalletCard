using System;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data;
using System.Drawing;

namespace PalletCard
{
    public partial class Home : Form
    {
        public Home()
        {
            InitializeComponent();
        }
        int A = 1;
        int numberUp, jobGanged, paperSectionNo, heightMM, invoiceCustomerCode, qtyRequired;
        string jobNo, resourceID, name, expr1, id, workingSize, description, code, jobDesc, invoiceCustomerName, ref7;
        bool jobCompleted, jobCancelled;

        private void btnBack_Click(object sender, EventArgs e)
        {
            // #1. Make second form
            Home form2 = new Home();
            // #2. Set second form's size
            form2.Width = this.Width;
            form2.Height = this.Height;
            // #3. Set second form's start position as same as parent form
            form2.StartPosition = FormStartPosition.Manual;
            form2.Location = new Point(this.Location.X, this.Location.Y);
            // #4. Set parent form's visibility to true
            this.Visible = true;
            // #5. Open second dialog
            form2.ShowDialog();
            // #6. Set parent form's visibility to false
            this.Visible = false;

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
                    getName();
                }
                else
                {
                    lblPress.Visible = false;
                    MessageBox.Show("The Job number you entered is not on this press");
                }
            }
            catch (Exception) { }
        }

        private void getName()
        {
            jobNo = dataGridView1.Rows[0].Cells[0].Value.ToString();
            resourceID = dataGridView1.Rows[0].Cells[1].Value.ToString();
            //loop through datagrid rows
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                //if datagrid is not empty create a button for each row at cells[2] - "Name"
                if (!(string.IsNullOrEmpty(dataGridView1.Rows[i].Cells[2].Value.ToString())))
                {
                    for (int j = 0; j < 1; j++ )
                    {
                        Button btn = new Button();
                        this.Controls.Add(btn);
                        btn.Top = A * 80;
                        btn.Height = 50;
                        btn.Width = 233;
                        btn.BackColor = Color.SteelBlue;
                        btn.Font = new Font("Microsoft Sans Serif", 14.25f);
                        btn.ForeColor = Color.White;
                        btn.Left = 260;
                        btn.Text = dataGridView1.Rows[i].Cells[2].Value.ToString();
                        A = A + 1;
                        btn.Click += new System.EventHandler(this.btnClick);     
                    }
                }
                
            }
        }
        void btnClick(object sender, EventArgs e) {
            Button btn = sender as Button;
            //foreach (Control c in this.Controls)
            //{
            //    if (c is Button)
            //    {
            //        Button bt = c as Button;
            //        MessageBox.Show(bt.Text);
            //    }
            //}
            MessageBox.Show(btn.Text);
        }
    }
}
