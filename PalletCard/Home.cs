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

        Button palletcard = new Button();
        Button returnpaper = new Button();
        Button backupvarnish = new Button();
        Button rejectpaper = new Button();

        private void btnBack_Click(object sender, EventArgs e)
        {
            // Make second form
            Home form2 = new Home();
            // Set second form's size
            form2.Width = this.Width;
            form2.Height = this.Height;
            // Set second form's start position as same as parent form
            form2.StartPosition = FormStartPosition.Manual;
            form2.Location = new Point(this.Location.X, this.Location.Y);
            // Set parent form's visibility to true
            this.Visible = true;
            // Open second dialog
            form2.ShowDialog();
            // Set parent form's visibility to false
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
                    //getSection();

                    this.Controls.Add(palletcard);
                    palletcard.Top = 80;
                    palletcard.Left = 260;
                    palletcard.Height = 50;
                    palletcard.Width = 233;
                    palletcard.BackColor = Color.SteelBlue;
                    palletcard.Font = new Font("Microsoft Sans Serif", 14.25f);
                    palletcard.Text = "Pallet Card";
                    palletcard.ForeColor = Color.White;

                    this.Controls.Add(returnpaper);
                    returnpaper.Top = 160;
                    returnpaper.Left = 260;
                    returnpaper.Height = 50;
                    returnpaper.Width = 233;
                    returnpaper.BackColor = Color.SteelBlue;
                    returnpaper.Font = new Font("Microsoft Sans Serif", 14.25f);
                    returnpaper.Text = "Return Paper";
                    returnpaper.ForeColor = Color.White;
                    returnpaper.Click += new System.EventHandler(Returnpaper_Click);

                    this.Controls.Add(backupvarnish);
                    backupvarnish.Top = 240;
                    backupvarnish.Left = 260;
                    backupvarnish.Height = 50;
                    backupvarnish.Width = 233;
                    backupvarnish.BackColor = Color.SteelBlue;
                    backupvarnish.Font = new Font("Microsoft Sans Serif", 14.25f);
                    backupvarnish.Text = "Back Up/Varnish";
                    backupvarnish.ForeColor = Color.White;

                    this.Controls.Add(rejectpaper);
                    rejectpaper.Top = 320;
                    rejectpaper.Left = 260;
                    rejectpaper.Height = 50;
                    rejectpaper.Width = 233;
                    rejectpaper.BackColor = Color.SteelBlue;
                    rejectpaper.Font = new Font("Microsoft Sans Serif", 14.25f);
                    rejectpaper.Text = "Reject paper";
                    rejectpaper.ForeColor = Color.White;
                }
                else
                {
                    lblPress.Visible = false;
                    MessageBox.Show("The Job number you entered is not on this press");
                }
            }
            catch (Exception) { }
        }

        private void Returnpaper_Click(object sender, EventArgs e)
        {
            lblReturnPaper.Visible = true;
            palletcard.Visible = false;
            returnpaper.Visible = false;
            backupvarnish.Visible = false;
            rejectpaper.Visible = false;
            getSection();
        }

        private void getSection()
        {
            jobNo = dataGridView1.Rows[0].Cells[0].Value.ToString();
            resourceID = dataGridView1.Rows[0].Cells[1].Value.ToString();

            //loop through datagrid rows                    
            for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
            {
                //if datagrid is not empty create a button for each row at cells[2] - "Name"

                if (!(string.IsNullOrEmpty(this.dataGridView1.Rows[i].Cells[11].Value as string)))
                {
                for (int j = 0; j < 1; j++)
                {
                        Button btn = new Button();
                    this.Controls.Add(btn);
                    btn.Top = A * 80;
                    btn.Height = 50;
                    btn.Width = 500;
                    btn.BackColor = Color.SteelBlue;
                    btn.Font = new Font("Microsoft Sans Serif", 13.25f);
                    btn.ForeColor = Color.White;
                    btn.Left = 260;
                    btn.Text = this.dataGridView1.Rows[i].Cells[11].Value as string;
                        A = A + 1;
                    btn.Click += new System.EventHandler(this.getName);
                }
                }        
            }
        }

        //Dynamic button click - Return Paper work flow
        void getName(object sender, EventArgs e) {
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
