using System;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Drawing.Printing;

namespace PalletCard
{
    public partial class Home : Form
    {
        List<Panel> listPanel = new List<Panel>();
        int index;
        bool sectionbtns;
        int A = 1;
        bool control;
        int numberUp, jobGanged, paperSectionNo, heightMM, invoiceCustomerCode, qtyRequired;
        string jobNo, resourceID, name, id, workingSize, description, code, jobDesc, invoiceCustomerName, ref7;
        bool jobCompleted, jobCancelled;

        private void btnCancel_Click(object sender, EventArgs e)
        {
            string ConnectionString = Convert.ToString("Dsn=TharData;uid=tharuser");
            string CommandText = "SELECT * FROM app_PalletOperations where resourceID = 5";
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
            listPanel[0].BringToFront();
            lblJobNo.Visible = false;
            lblPress.Visible = false;
            lblReturnPaper.Visible = false;
            lbltextBoxDescription.Visible = false;
            lblWorkingSize.Visible = false;
            searchBox.Text = "";
            searchBox.Focus();
            sectionbtns = true;
            tbxPalletHeight.Text = null;
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            PrintDocument pd = new PrintDocument();
            pd.PrintPage += new PrintPageEventHandler(PrintImage);
            btnPrint.Visible = false;
            pd.Print();
            btnPrint.Visible = true;
        }

        void PrintImage(object o, PrintPageEventArgs e)
        {
            int x = SystemInformation.WorkingArea.X;
            int y = SystemInformation.WorkingArea.Y;
            int width = this.Width;
            int height = this.Height;
            Rectangle bounds = new Rectangle(x, y, width, height);
            Bitmap img = new Bitmap(width, height);
            returnpaper4.DrawToBitmap(img, bounds);
            Point p = new Point(100, 100);
            e.Graphics.DrawImage(img, p);
        }

        private void btnPalletHeight_Click(object sender, EventArgs e)
        {
            listPanel[4].BringToFront();
            lblPrint1.Text = dataGridView1.Rows[0].Cells[16].Value.ToString();
            lblPrint2.Text = dataGridView1.Rows[0].Cells[13].Value.ToString();
            lblPrint3.Text = lblPheight.Text;
            lblPrint4.Text = "Press - 710UV";
            lblPrint5.Text = "Job - " + jobNo;
            lblPrint6.Text = "Date - " + DateTime.Now.ToString("d/M/yyyy");
        }

        // Pallet Height textBox
        private void tbxPalletHeight_TextChanged(object sender, EventArgs e)
        {
            TextBox objTextBox = (TextBox)sender;
            int p1;
            int p2;
            if (!String.IsNullOrEmpty(tbxPalletHeight.Text))
            { 
            p1 = Convert.ToInt32(objTextBox.Text);
            p2 = Convert.ToInt32(this.dataGridView1.Rows[0].Cells[20].Value);
            int result = p1 * p2 /1000;
            string r1 = Convert.ToString(result);
            lblPheight.Text = (r1 + " sheets");
            }
        }

        public Home()
        {
            InitializeComponent();
        }

        private void btnReturnPaper_Click(object sender, EventArgs e)
        {
                lblReturnPaper.Visible = true;
                listPanel[2].BringToFront();
                index = 2;
                jobNo = dataGridView1.Rows[0].Cells[0].Value.ToString();

            //loop through datagridview to see if each value of field "Expr1" is the same
            string x;
            string y;
            x = dataGridView1.Rows[0].Cells[11].Value.ToString();
            control = false;
            for (int i = 1; i < this.dataGridView1.Rows.Count - 1; i++)
            {
                y = dataGridView1.Rows[i].Cells[11].Value.ToString();
                if (x == y) { control = true; }
            }
            if (control) {                               
                listPanel[3].BringToFront();
                string d = dataGridView1.Rows[0].Cells[11].Value.ToString();
                lbltextBoxDescription.Text = d;
                lbltextBoxDescription.Visible = true;
                lblWorkingSize.Visible = true;
                lblWorkingSize.Text = dataGridView1.Rows[0].Cells[13].Value.ToString();
                index = 4;
                sectionbtns = true;
            }
            else
            { //prevent section buttons from drawing again if back button is selected
                if (!sectionbtns)
                {
                    //loop through datagrid rows to create a button for each value of field "Expr1"                  
                    for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
                        {
                            //if datagrid is not empty create a button for each row at cells[2] - "Name"
                            if (!(string.IsNullOrEmpty(this.dataGridView1.Rows[i].Cells[11].Value as string)))
                            {
                                for (int j = 0; j < 1; j++)
                                {
                                //var v = this.dataGridView1.Rows[i].Cells[11].Value as string; ;
                                //if (btnBack.Text != v) { 
                                    Button btn = new Button();
                                    this.returnpaper2.Controls.Add(btn);
                                    btn.Top = A * 80;
                                    btn.Height = 48;
                                    btn.Width = 465;
                                    btn.BackColor = Color.SteelBlue;
                                    btn.Font = new Font("Microsoft Sans Serif", 13.25f);
                                    btn.ForeColor = Color.White;
                                    btn.Left = 30;
                                    btn.Text = this.dataGridView1.Rows[i].Cells[11].Value as string;
                                    A = A + 1;
                                    btn.Click += new System.EventHandler(this.expr1);
                                //}
                            }
                            }
                        }
                    }
                sectionbtns = true;
                index = 3;
            }
        }

        //private void btnNext_Click(object sender, EventArgs e)
        //{
        //    if (index < listPanel.Count - 1)
        //        listPanel[++index].BringToFront();
        //}

        private void btnBack_Click(object sender, EventArgs e)
        {
            //if (index > 0)
            //    listPanel[--index].BringToFront();
            if (index == 1)
                listPanel[0].BringToFront();
                lblReturnPaper.Visible = false;
                lbltextBoxDescription.Visible = false;
                lblWorkingSize.Visible = false;
                tbxPalletHeight.Text = "";
                lblPheight.Text = "";
            if (index == 2)
                listPanel[1].BringToFront();
                lblReturnPaper.Visible = true;
                lbltextBoxDescription.Visible = false;
                lblWorkingSize.Visible = false;               
            if (index == 3)
                listPanel[2].BringToFront();
                lblWorkingSize.Visible = false;
            if (index == 4)
                listPanel[3].BringToFront();
                lblReturnPaper.Visible = true;
                lbltextBoxDescription.Visible = false;
                lblWorkingSize.Visible = false;
                lblPheight.Text = "";
        }

        private void Home_Load(object sender, EventArgs e)
        {
            string ConnectionString = Convert.ToString("Dsn=TharData;uid=tharuser");
            string CommandText = "SELECT * FROM app_PalletOperations where resourceID = 5";
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
            listPanel.Add(returnpaper0);
            listPanel.Add(returnpaper1);
            listPanel.Add(returnpaper2);
            listPanel.Add(returnpaper3);
            listPanel.Add(returnpaper4);
            listPanel[0] = returnpaper0;
            listPanel[1] = returnpaper1;
            listPanel[2] = returnpaper2;
            listPanel[3] = returnpaper3;
            listPanel[4] = returnpaper4;
            listPanel[0].BringToFront();
        }

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
                    listPanel[1].BringToFront();
                }
                else
                {
                    lblPress.Visible = false;
                    MessageBox.Show("The Job number you entered is not on this press");
                }
            }
            catch (Exception) { }
            index = 1;
        }
        //private void getSection()
        //{
        //    jobNo = dataGridView1.Rows[0].Cells[0].Value.ToString();
        //    resourceID = dataGridView1.Rows[0].Cells[1].Value.ToString();

        //    //loop through datagrid rows                    
        //    for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
        //    {
        //        //if datagrid is not empty create a button for each row at cells[11] - "Expr1"

        //        if (!(string.IsNullOrEmpty(this.dataGridView1.Rows[i].Cells[11].Value as string)))
        //        {
        //            for (int j = 0; j < 1; j++)
        //            {
        //                Button btn = new Button();
        //                this.Controls.Add(btn);
        //                btn.Top = A * 80;
        //                btn.Height = 50;
        //                btn.Width = 500;
        //                btn.BackColor = Color.SteelBlue;
        //                btn.Font = new Font("Microsoft Sans Serif", 13.25f);
        //                btn.ForeColor = Color.White;
        //                btn.Left = 260;
        //                btn.Text = this.dataGridView1.Rows[i].Cells[11].Value as string;
        //                A = A + 1;
        //                btn.Click += new System.EventHandler(this.expr1);
        //            }
        //        }        
        //    }
        //}

        //Dynamic button click - Section buttons, Return Paper work flow
        private void expr1(object sender, EventArgs e) {
            Button btn = sender as Button;

            //foreach (Control c in this.Controls)
            //{
            //    if (c is Button)
            //    {
            //        Button bt = c as Button;
            //        MessageBox.Show(bt.Text);
            //    }
            //}
            //MessageBox.Show(btn.Text);

            listPanel[3].BringToFront();
            lblWorkingSize.Visible = true;
            lblWorkingSize.Text = dataGridView1.Rows[0].Cells[13].Value.ToString();
            index = 4;
            lbltextBoxDescription.Text = btn.Text;
            lbltextBoxDescription.Visible = true;



            //sectionbtns = true;

            //for (int ix = this.Controls.Count - 1; ix >= 0; ix--)
            //{
            //    if (this.Controls[ix] is Button ) this.Controls[ix].Dispose();
            //}
        }
    }
}
