﻿using System;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Data.SqlClient;
using System.Drawing.Imaging;

namespace PalletCard
{
    public partial class Home : Form
    {
        List<Panel> listPanel = new List<Panel>();
        int index;
        bool sectionBtns;
        bool sigBtns;
        int A = 1;
        string jobNo;
        bool searchChanged;

        public Home()
        {
            InitializeComponent();
        }
       
        private void btnBack_Click(object sender, EventArgs e)
        {
            if (index == 0)
            {
                btnBack.Visible = false;
                btnSearch.Visible = false;
                lblJobNo.Visible = false;
                lblPress.Visible = false;
            }
            else if (index == 1)
            {
                pnlHome0.BringToFront();
                lbl1.Visible = false;
                lbl2.Visible = false;
                lbl3.Visible = false;
                lbl4.Visible = false;
                tbxPalletHeight.Text = "";
                lblPheight.Text = "";
                index = 0;
            }
            else if (index == 2)
            {
                pnlHome1.BringToFront();
                lbl1.Visible = false;
                lbl2.Visible = false;
                lbl3.Visible = false;
                lbl4.Visible = false;
                btnBack.Visible = false;
                Search();
                index = 1;
                sectionBtns = false;
            }
            else if (index == 3)
            {
                pnlReturnPaper1.BringToFront();
                lblPheight.Text = "";
                tbxPalletHeight.Text = "";
                lbl2.Visible = false;
                lbl3.Visible = false;
                lbl4.Visible = false;
                index = 2;
                // if no section buttons go straight back to Choose Action screen
                if (pnlReturnPaper1.Controls.Count == 0)
                {
                    pnlHome1.BringToFront();
                    lbl1.Visible = false;
                    lbl2.Visible = false;
                    lbl3.Visible = false;
                    lbl4.Visible = false;
                    lblPheight.Text = "";
                    lbl3.Visible = false;
                    lbl4.Visible = false;
                    tbxPalletHeight.Text = "";
                    btnBack.Visible = false;
                    index = 1;
                }
            }
            else if (index == 4)
            {
                pnlReturnPaper2.BringToFront();
                lbl1.Visible = true;
                lbl2.Visible = true;
                lbl3.Visible = true;
                lbl4.Visible = true;
                lblPheight.Text = "";
                tbxPalletHeight.Text = "";
                this.ActiveControl = tbxPalletHeight;
                index = 3;
            }
            else if (index == 5)
            {
                pnlHome1.BringToFront();
                lbl1.Visible = false;
                lbl2.Visible = false;
                lbl3.Visible = false;
                lbl4.Visible = false;
                btnBack.Visible = false;
                tbxQtySheetsAffected.Text = "";
                lblBack5.Visible = false;
                lblBack6.Visible = false;
                lblBack5.Visible = true;
                lblBack6.Visible = true;
                Search();
                index = 1;
                sectionBtns = false;
            }
            else if (index == 6)
            {
                pnlRejectPaper1.BringToFront();
                lbl2.Visible = false;
                lbl3.Visible = false;
                lbl4.Visible = false;
                tbxQtySheetsAffected.Text = "";
                tbxOtherReason.Text = "";
                ckbDogEarsTIC.Checked = false;
                ckbMottle.Checked = false;
                ckbCreasing.Checked = false;
                ckbCigarRoll.Checked = false;
                ckbPalletDamage.Checked = false;
                ckbBladeLine.Checked = false;
                lblBack5.Visible = false;
                lblBack6.Visible = false;
                index = 5;
                // if no section buttons go straight back to Choose Action screen
                if (pnlRejectPaper1.Controls.Count == 0)
                {
                    pnlHome1.BringToFront();
                    lbl1.Visible = false;
                    lbl2.Visible = false;
                    lbl3.Visible = false;
                    lbl3.Visible = false;
                    lbl4.Visible = false;
                    btnBack.Visible = false;
                    tbxOtherReason.Text = "";
                    ckbDogEarsTIC.Checked = false;
                    ckbMottle.Checked = false;
                    ckbCreasing.Checked = false;
                    ckbCigarRoll.Checked = false;
                    ckbPalletDamage.Checked = false;
                    ckbBladeLine.Checked = false;
                    tbxQtySheetsAffected.Text = "";
                    index = 1;
                }
            }
            else if (index == 7)
            {
                pnlRejectPaper2.BringToFront();
                lbl1.Visible = true;
                lbl2.Visible = true;
                lbl3.Visible = false;
                lbl4.Visible = false;
                this.ActiveControl = tbxQtySheetsAffected;
                tbxQtySheetsAffected.Text = "";
                tbxOtherReason.Text = "";
                ckbDogEarsTIC.Checked = false;
                ckbMottle.Checked = false;
                ckbCreasing.Checked = false;
                ckbCigarRoll.Checked = false;
                ckbPalletDamage.Checked = false;
                ckbBladeLine.Checked = false;
                tbxQtySheetsAffected.Text = "";
                index = 6;
            }
            else if (index == 8)
            {
                pnlHome1.BringToFront();
                lbl1.Visible = false;
                lbl2.Visible = false;
                lbl3.Visible = false;
                lbl4.Visible = false;
                btnBack.Visible = false;
                lblBack5.Visible = false;
                lblBack6.Visible = false;
                lblBack5.Visible = true;
                lblBack6.Visible = true;
                Search();
                index = 1;
                sectionBtns = false;
            }
            else if (index == 9)
            {
                Search();
                pnlPalletCard1.BringToFront();
                lbl2.Visible = false;
                lbl3.Visible = false;
                lbl4.Visible = false;
                btnBack.Visible = true;
                index = 8;
                sigBtns = false;
                removeFlowLayoutBtns();
                // if no section buttons go straight back to Choose Action screen
                if (pnlPalletCard1.Controls.Count == 0)
                {
                    pnlHome1.BringToFront();
                    lbl1.Visible = false;
                    lbl2.Visible = false;
                    lbl3.Visible = false;
                    lbl3.Visible = false;
                    lbl4.Visible = false;
                    btnBack.Visible = false;
                    index = 1;
                }
            }
            else if (index == 10)
            {
                pnlPalletCard2.BringToFront();
                lbl3.Visible = false;
                lbl4.Visible = false;
                index = 9;
                // if no sig buttons go straight back to Choose Section screen
                if (pnlPalletCard2.Controls.Count == 0)
                {
                    pnlHome1.BringToFront();
                    lbl1.Visible = false;
                    lbl2.Visible = false;
                    lbl3.Visible = false;
                    lbl3.Visible = false;
                    lbl4.Visible = false;
                    btnBack.Visible = false;
                    index = 1;
                }
            }
        }



        private void Home_Load(object sender, EventArgs e)
        {
            string ConnectionString = Convert.ToString("Dsn=TharData;uid=tharuser");
            string CommandText = "SELECT * FROM app_PalletOperations where resourceID = 6";
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
            listPanel.Add(pnlHome0);
            listPanel.Add(pnlHome1);
            listPanel.Add(pnlReturnPaper1);
            listPanel.Add(pnlReturnPaper2);
            listPanel.Add(pnlReturnPaper3);
            listPanel[0] = pnlHome0;
            listPanel[1] = pnlHome1;
            listPanel[2] = pnlReturnPaper1;
            listPanel[3] = pnlReturnPaper2;
            listPanel[4] = pnlReturnPaper3;
            listPanel[0].BringToFront();
            btnBack.Visible = false;
            this.ActiveControl = tbxSearchBox;
        }


        private void Search()
        {
            {
                try
                {
                    ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("JobNo like '%{0}%'", tbxSearchBox.Text.Trim().Replace("'", "''"));
                    lblJobNo.Text = dataGridView1.Rows[0].Cells[0].Value.ToString();
                    lblJobNo.Visible = true;
                    int resourceID = (int)dataGridView1.Rows[0].Cells[1].Value;
                    if (dataGridView1.Rows[0].Cells[0].Value != null)
                    {
                        lblPress.Text = "XL106";
                        lblPress.Visible = true;
                        pnlHome1.BringToFront();
                    }
                    else
                    {
                        lblPress.Visible = false;
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Tap Cancel and search again");
                }
                index = 1;
                if (searchChanged == true)
                {
                    pnlReturnPaper1.Controls.Clear();
                }
                //reset dynamic buttons origin
                A = 1;
                btnBack.Visible = false;
            }
        }
        private void btnSearch_Click(object sender, EventArgs e)
        {
            Search();
        }

        private void searchBox_TextChanged(object sender, EventArgs e)
        {
            searchChanged = true;
        }


        private void Cancel()
        {
            string ConnectionString = Convert.ToString("Dsn=TharData;uid=tharuser");
            string CommandText = "SELECT * FROM app_PalletOperations where resourceID = 6";
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
            pnlHome0.BringToFront();
            lblJobNo.Visible = false;
            lblPress.Visible = false;
            lbl1.Visible = false;
            lbl2.Visible = false;
            lbl3.Visible = false;
            lbl4.Visible = false;
            lblPheight.Text = "";
            tbxSearchBox.Text = "";
            tbxSearchBox.Focus();
            sectionBtns = false;
            tbxPalletHeight.Text = null;
            btnSearch.Visible = true;
            btnBack.Visible = false;
            this.ActiveControl = tbxSearchBox;
            index = 0;
        }


        private void btnCancel_Click(object sender, EventArgs e)
        {
            Cancel();
        }

        private void tbxPalletHeight_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                SelectNextControl(ActiveControl, true, true, true, true);
            }
        }

        private void searchBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                btnSearch_Click(tbxSearchBox, new EventArgs());
        }

        private void removeFlowLayoutBtns()
        {
            flowLayoutPanel1.Controls.Clear();
        }

        //****************************************************************************************************
        //    RETURN PAPER WORKFLOW
        //****************************************************************************************************

        private void btnReturnPaper_Click(object sender, EventArgs e)
        {
                lbl1.Visible = true;
                lbl1.Text = "Return Paper";
                pnlReturnPaper1.BringToFront();
                index = 2;
                jobNo = dataGridView1.Rows[0].Cells[0].Value.ToString();
                btnBack.Visible = true;

            //loop through datagridview to see if each value of field "Expr1" is the same
            //(If only one datagridview row or rows are all the same then go straight to pallet height - else create dynamic buttons)
            string x;
            string y;
            x = dataGridView1.Rows[0].Cells[11].Value.ToString();
            y = dataGridView1.Rows[0].Cells[11].Value.ToString();
            for (int i = 1; i < this.dataGridView1.Rows.Count; i++)
            {
                y = dataGridView1.Rows[i].Cells[11].Value.ToString();
            }       
            if (x == y)
            {
                pnlReturnPaper2.BringToFront();
                string d = dataGridView1.Rows[0].Cells[11].Value.ToString();
                lbl2.Text = d;           
                lbl2.Visible = true;
                lbl3.Visible = true;
                lbl3.Text = this.dataGridView1.Rows[0].Cells[16].Value as string;
                lbl4.Visible = true;
                lbl4.Text = dataGridView1.Rows[0].Cells[13].Value.ToString();
                index = 3;
                sectionBtns = true;
                this.ActiveControl = tbxPalletHeight;
            }       
            else
            { //prevent section buttons from drawing again if back button is selected
                if (!sectionBtns)
                {
                    //loop through datagrid rows to create a button for each value of field "Expr1"                  
                    for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
                        {
                        //if datagrid is not empty create a button for each row at cells[11] - "Expr1"
                        if (!(string.IsNullOrEmpty(this.dataGridView1.Rows[i].Cells[11].Value as string)))

                            //offer only one button where Expr1 field has two rows with the same value
                            dataGridView1.AllowUserToAddRows = true;
                            if (! (this.dataGridView1.Rows[i].Cells[11].Value as string == this.dataGridView1.Rows[i+1].Cells[11].Value as string))
                            { 
                                {
                                    for (int j = 0; j < 1; j++)
                                    { 
                                        Button btn = new Button();
                                        this.pnlReturnPaper1.Controls.Add(btn);
                                        btn.Top = A * 100;
                                        btn.Height = 80;
                                        btn.Width = 465;
                                        btn.BackColor = Color.SteelBlue;
                                        btn.Font = new Font("Microsoft Sans Serif", 14);
                                        btn.ForeColor = Color.White;
                                        btn.Left = 30;                                     
                                        btn.Text = this.dataGridView1.Rows[i].Cells[11].Value as string;
                                        A = A + 1;
                                        btn.Click += new System.EventHandler(this.expr1);
                                    }
                                }
                            dataGridView1.AllowUserToAddRows = false;
                            }
                        }
                    }
                    sectionBtns = true;
                }          
        }

        //Dynamic button click - Section buttons, Return Paper work flow
        private void expr1(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            pnlReturnPaper2.BringToFront();

            //filter datagridview1 with the button text choice
            try
            {
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = "Expr1 like '%" + btn.Text + "%' and JobNo like '%" + lblJobNo.Text + "%'";
            }
            catch (Exception) { }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                string val = dataGridView1.Rows[i].Cells[11].Value.ToString();
                if (btn.Text == val)
                {
                    lbl4.Text = dataGridView1.Rows[i].Cells[13].Value.ToString();
                    lbl3.Text = dataGridView1.Rows[i].Cells[16].Value.ToString();
                }
            }
            lbl4.Visible = true;
            lbl3.Visible = true;
            lbl2.Visible = true;
            lbl2.Text = btn.Text;
            tbxPalletHeight.Text = "";
            this.ActiveControl = tbxPalletHeight;
            index = 3;
        }

        private void btnPalletHeight_Click(object sender, EventArgs e)
        {
            pnlReturnPaper3.BringToFront();
            lblPrint1.Text = dataGridView1.Rows[0].Cells[16].Value.ToString();
            lblPrint2.Text = dataGridView1.Rows[0].Cells[13].Value.ToString();
            lblPrint3.Text = lblPheight.Text;
            lblPrint4.Text = "Press - XL106";
            lblPrint5.Text = "Job - " + jobNo;
            lblPrint6.Text = "Date - " + DateTime.Now.ToString("d/M/yyyy");
            index = 4;
        }

        // Pallet Height textBox calculation for Return Paper
        private void tbxPalletHeight_TextChanged(object sender, EventArgs e)

        {
            try
            { 
                TextBox objTextBox = (TextBox)sender;
                double p1;
                double p2;
                if (!String.IsNullOrEmpty(tbxPalletHeight.Text))
                {
                    p1 = Convert.ToInt32(objTextBox.Text);
                    p2 = Convert.ToInt32(this.dataGridView1.Rows[0].Cells[20].Value);
                    double result = Math.Ceiling (p1 / (p2/1000)) ;
                    string r1 = Convert.ToString(result);
                    lblPheight.Text = (r1 + " sheets");
                }
            }
            catch
            {
                MessageBox.Show("Please enter a valid number");
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            PrintDocument pd = new PrintDocument();
            pd.PrintPage += new PrintPageEventHandler(PrintSig);
            btnPrint.Visible = false;
            pd.Print();
            btnPrint.Visible = true;

            DateTime CurrentDate = DateTime.Now;
            string sqlFormattedDate = CurrentDate.ToString("yyyy-MM-dd HH:mm:ss.fff");

            string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
            string Query = "insert into Log (Routine, JobNo, ResourceID, Expr1, WorkingSize, SheetQty, Description, Timestamp1) values('" + this.lbl1.Text + "','" + this.dataGridView1.Rows[0].Cells[0].Value + "','" + this.dataGridView1.Rows[0].Cells[1].Value + "','" + this.lbl2.Text + "','" + this.lbl4.Text + "','" + this.lblPrint3.Text + "','" + this.lbl3.Text + "','" + CurrentDate + "');";
            SqlConnection conDatabase = new SqlConnection(constring);
            SqlCommand cmdDatabase = new SqlCommand(Query, conDatabase);
            SqlDataReader myReader;
            try
            {
                conDatabase.Open();
                myReader = cmdDatabase.ExecuteReader();
                while (myReader.Read())
                {

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            pnlHome0.BringToFront();
            lblJobNo.Visible = false;
            lblPress.Visible = false;
            lbl1.Visible = false;
            lbl2.Visible = false;
            lbl3.Visible = false;
            lbl4.Visible = false;
            btnBack.Visible = false;
            btnCancel.Visible = false;
            Cancel();
        }

        void PrintSig(object o, PrintPageEventArgs e)
        {
            int x = SystemInformation.WorkingArea.X;
            int y = SystemInformation.WorkingArea.Y;
            int width = this.Width;
            int height = this.Height;
            Rectangle bounds = new Rectangle(x, y, width, height);
            Bitmap img = new Bitmap(width, height);
            pnlReturnPaper3.DrawToBitmap(img, bounds);
            Point p = new Point(100, 100);
            e.Graphics.DrawImage(img, p);
        }

        


//****************************************************************************************************
//REJECT PAPER WORKFLOW
//****************************************************************************************************

        private void btnRejectPaper_Click(object sender, EventArgs e)
        {
            lbl1.Visible = true;
            lbl1.Text = "Reject Paper";
            pnlRejectPaper1.BringToFront();
            index = 5;
            jobNo = dataGridView1.Rows[0].Cells[0].Value.ToString();
            btnBack.Visible = true;
            this.ActiveControl = tbxQtySheetsAffected;

            //loop through datagridview to see if each value of field "Expr1" is the same
            string x;
            string y;
            x = dataGridView1.Rows[0].Cells[11].Value.ToString();
            y = dataGridView1.Rows[0].Cells[11].Value.ToString();
            for (int i = 1; i < this.dataGridView1.Rows.Count; i++)
            {
                y = dataGridView1.Rows[i].Cells[11].Value.ToString();
            }
            if (x == y)
            {
                pnlRejectPaper2.BringToFront();
                string d = dataGridView1.Rows[0].Cells[11].Value.ToString();
                lbl2.Text = d;
                lbl2.Visible = true;
                lblBack5.Visible = false;
                lblBack6.Visible = false;
                index = 6;
                sectionBtns = true;
            }
            else
            { //prevent section buttons from drawing again if back button is selected
                if (!sectionBtns)
                {
                    //loop through datagrid rows to create a button for each value of field "Expr1"                  
                    for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
                    {
                        //if datagrid is not empty create a button for each row at cells[2] - "Name"
                        if (!(string.IsNullOrEmpty(this.dataGridView1.Rows[i].Cells[11].Value as string)))

                            //offer only one button where Expr1 field has two rows with the same value
                            dataGridView1.AllowUserToAddRows = true;
                        if (!(this.dataGridView1.Rows[i].Cells[11].Value as string == this.dataGridView1.Rows[i + 1].Cells[11].Value as string))
                            {
                                {
                                    for (int j = 0; j < 1; j++)
                                    {
                                        Button btn = new Button();
                                        this.pnlRejectPaper1.Controls.Add(btn);
                                        btn.Top = A * 100;
                                        btn.Height = 80;
                                        btn.Width = 465;
                                        btn.BackColor = Color.SteelBlue;
                                        btn.Font = new Font("Microsoft Sans Serif", 14);
                                        btn.ForeColor = Color.White;
                                        btn.Left = 30;
                                        btn.Text = this.dataGridView1.Rows[i].Cells[11].Value as string;
                                        A = A + 1;
                                        btn.Click += new System.EventHandler(this.expr2);
                                    }
                                }
                            }
                            dataGridView1.AllowUserToAddRows = false;
                    }
                }
                sectionBtns = true;
                lblBack5.Visible = false;
                lblBack6.Visible = false;
            }
        }


        //Dynamic button click - Section buttons, Reject Paper work flow
        private void expr2(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            pnlRejectPaper2.BringToFront();
            lbl2.Visible = true;
            lbl2.Text = btn.Text;
            this.ActiveControl = tbxQtySheetsAffected;
            index = 6;

            //filter datagridview1 with the button text choice
            try
            {
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = "Expr1 like '%" + btn.Text + "%' and JobNo like '%" + lblJobNo.Text + "%'";
            }
            catch (Exception) { }
        }

        private void tbxQtySheetsAffected_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                SelectNextControl(ActiveControl, true, true, true, true);
            }
        }

        private void btnOKRejectPaper_Click(object sender, EventArgs e)
        {
            //Get values of checked checkboxes
            pnlRejectPaper3.BringToFront();
            string s = "";
            foreach (Control c in this.groupBox1.Controls)
            {
                if (c is CheckBox)
                {
                    CheckBox b = (CheckBox)c;
                    if (b.Checked)
                        {
                        s = b.Text + " * " + s;
                        }
                }
            }

            lblPrint7.Text = s;
            lblPrint7.MaximumSize = new Size(450, 220);
            lblPrint7.AutoSize = true;
            lblPrint8.Text = dataGridView1.Rows[0].Cells[13].Value.ToString();
            lblPrint9.Text = tbxQtySheetsAffected.Text + " Sheets";
            int parsedValue;
            if (!int.TryParse(tbxQtySheetsAffected.Text, out parsedValue))
            {
                MessageBox.Show("Please enter a valid number in the Quantity Sheets affected box");
                pnlRejectPaper2.BringToFront();
                ActiveControl = tbxQtySheetsAffected;
                tbxQtySheetsAffected.Text = "";
                return;
            }

            lblPrint10.Text = "Press - XL106";
            lblPrint11.Text = "Job - " + jobNo;
            lblPrint12.Text = "Date - " + DateTime.Now.ToString("d/M/yyyy");
            lblPrint13.Text = tbxOtherReason.Text;
            lblPrint14.Text = this.dataGridView1.Rows[0].Cells[17].Value.ToString();
            index = 7;
        }

        private void btnRejectPaperPrint_Click(object sender, EventArgs e)
        {
            PrintDocument pd = new PrintDocument();
            pd.PrintPage += new PrintPageEventHandler(PrintImageRejectPaper);
            btnRejectPaperPrint.Visible = false;
            pd.Print();
            btnRejectPaperPrint.Visible = true;

            //string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
            //string Query = "insert into Log (Routine, JobNo, ResourceID, Description, WorkingSize, SheetQty) values('" + this.lbl1.Text + "','" + this.dataGridView1.Rows[0].Cells[0].Value + "','" + this.dataGridView1.Rows[0].Cells[1].Value + "','" + this.lbl2.Text + "','" + this.lbl4.Text + "','" + this.lblPrint3.Text + "');";
            //SqlConnection conDatabase = new SqlConnection(constring);
            //SqlCommand cmdDatabase = new SqlCommand(Query, conDatabase);
            //SqlDataReader myReader;
            //try
            //{
            //    conDatabase.Open();
            //    myReader = cmdDatabase.ExecuteReader();
            //    while (myReader.Read())
            //    {

            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
            pnlHome0.BringToFront();
            lblJobNo.Visible = false;
            lblPress.Visible = false;
            lbl1.Visible = false;
            lbl2.Visible = false;
            lbl3.Visible = false;
            lbl4.Visible = false;
            btnBack.Visible = false;
            btnCancel.Visible = false;
            Cancel();
        }

        void PrintImageRejectPaper(object o, PrintPageEventArgs e)
        {
            int x = SystemInformation.WorkingArea.X;
            int y = SystemInformation.WorkingArea.Y;
            int width = this.Width;
            int height = this.Height;
            Rectangle bounds = new Rectangle(x, y, width, height);
            Bitmap img = new Bitmap(width, height);
            pnlRejectPaper3.DrawToBitmap(img, bounds);
            Point p = new Point(100, 100);
            e.Graphics.DrawImage(img, p);
        }


//****************************************************************************************************
//    SIGNATURE
//****************************************************************************************************

        public class Line
        {
            public Line()
            {
            }

            public Line(Point startPoint, Point endPoint)
            {
                this.StartPoint = startPoint;
                this.EndPoint = endPoint;
            }

            public Point StartPoint { get; set; }
            public Point EndPoint { get; set; }
        }

        public class Glyph
        {
            public Glyph()
            {
                this.Lines = new List<Line>();
            }
            public List<Line> Lines { get; set; }
        }

        public class Signature
        {
            public Signature()
            {
                this.Glyphs = new List<Glyph>();
            }

            public List<Glyph> Glyphs { get; set; }
        }

        Boolean IsCapturing = false;
        private Point startPoint;
        private Point endPoint;
        Pen pen = new Pen(Color.Black);
        Glyph glyph = null;
        Signature signature = new Signature();
        //String fileName = @"signature.xml";

        private void SignaturePanel_MouseMove(object sender, MouseEventArgs e)
        {
            if (IsCapturing)
            {
                if (startPoint.IsEmpty && endPoint.IsEmpty)
                {
                    endPoint = e.Location;
                }
                else
                {
                    startPoint = endPoint;
                    endPoint = e.Location;
                    Line line = new Line(startPoint, endPoint);
                    glyph.Lines.Add(line);
                    DrawLine(line);
                }
            }
        }

        private void SignaturePanel_MouseUp(object sender, MouseEventArgs e)
        {
            IsCapturing = false;
            signature.Glyphs.Add(glyph);
            startPoint = new Point();
            endPoint = new Point();
        }

        private void SignaturePanel_MouseDown(object sender, MouseEventArgs e)
        {
            IsCapturing = true;
            glyph = new Glyph();
        }

        private void DrawLine(Line line)
        {
            using (Graphics graphic = this.SignaturePanel.CreateGraphics())
            {
                graphic.DrawLine(pen, line.StartPoint, line.EndPoint);
            }
        }

        private void DrawSignature()
        {
            foreach (Glyph glyph in signature.Glyphs)
            {
                foreach (Line line in glyph.Lines)
                {
                    DrawLine(line);
                }
            }
        }

        private void ClearSignaturePanel()
        {
            using (Graphics graphic = this.SignaturePanel.CreateGraphics())
            {
                SolidBrush solidBrush = new SolidBrush(Color.Gainsboro);
                graphic.FillRectangle(solidBrush, 0, 0, SignaturePanel.Width, SignaturePanel.Height);
            }
        }

        private void buttonClear_Click(object sender, EventArgs e)
        {
            ClearSignaturePanel();
        }

        private static Bitmap DrawControlToBitmap(Control control)
        {
            Bitmap bitmap = new Bitmap(control.Width, control.Height);
            Graphics graphics = Graphics.FromImage(bitmap);
            Rectangle rect = control.RectangleToScreen(control.ClientRectangle);
            graphics.CopyFromScreen(rect.Location, Point.Empty, control.Size);
            return bitmap;
        }

        private void getAutoNumber()
        {

            //string ConnectionString = Convert.ToString("Dsn=TharData;uid=tharuser");
            //string CommandText = "SELECT * FROM app_PalletOperations where resourceID = 6";
            //OdbcConnection myConnection = new OdbcConnection(ConnectionString);
            //OdbcCommand myCommand = new OdbcCommand(CommandText, myConnection);


            string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
            using (SqlConnection cs = new SqlConnection(constring))
            {
                try
                {
                    string query = "SELECT MAX(AutoNum) FROM Log";
                    //SqlCommand comSelect = new SqlCommand(query, constring);
                    //int autoNum = (int)comSelect.ExecuteScalar();


                    //SqlCommand cmd = new SqlCommand("SELECT MAX(AutoNum) FROM Log");
                    //cs.Open();
                    //int autoNum = (int)cmd.ExecuteScalar() +1;
                    //cmd.Connection = constring;
                    //cs.Close();


                    //SqlCommand cmd = new SqlCommand("SELECT TOP 1 Signature FROM Log ORDER BY Signature DESC");
                    //cs.Open();
                    //cmd.ExecuteNonQuery();
                    //cs.Close();


                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message);
                }
            }
        }
        int autoNum;

        private void btnQATravellerBlurb_Click(object sender, EventArgs e)
        {
            //PrintDocument pd = new PrintDocument();
            //pd.PrintPage += new PrintPageEventHandler(PrintSignature);
            //btnPrint.Visible = false;
            ////pd.Print();
            //btnPrint.Visible = true;

            getAutoNumber();

            DateTime CurrentDate = DateTime.Now;
            string sqlFormattedDate = CurrentDate.ToString("yyyy-MM-dd HH:mm:ss.fff");


            string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";

            Bitmap bitmap = DrawControlToBitmap(SignaturePanel);
            bitmap.Save("c://Temp//" + this.autoNum + ".jpg", ImageFormat.Jpeg);
            System.Diagnostics.Process.Start("c://Temp//"+ this.autoNum + ".jpg");
          
            using (SqlConnection cs = new SqlConnection(constring))
            {
                try
                {
                    SqlCommand cmd = new SqlCommand("insert Log(Signature, Timestamp1) values('" + this.autoNum + "','" + CurrentDate + "')", cs);                  
                    cs.Open();
                    cmd.ExecuteNonQuery();
                    cs.Close();
                    this.autoNum = this.autoNum + 1;
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message);
                }
            }

            //string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
            //string Query = "insert into Log (SignatureByte) values(@SignatureByte);";

            //SqlConnection conDatabase = new SqlConnection(constring);
            //SqlCommand cmdDatabase = new SqlCommand(Query, conDatabase);
            //SqlDataReader myReader;


            //cmdDatabase.Parameters.AddWithValue(constring, SignatureByte);


            //try
            //{
            //    conDatabase.Open();
            //    myReader = cmdDatabase.ExecuteReader();
            //    while (myReader.Read())
            //    {

            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }





//****************************************************************************************************
//   PALLET CARD
//****************************************************************************************************

        private void btnPalletCard_Click(object sender, EventArgs e)
        {
            //pnlSignature.BringToFront();

            lbl1.Visible = true;
            lbl1.Text = "Pallet Card";
            pnlPalletCard1.BringToFront();
            index = 8;
            jobNo = dataGridView1.Rows[0].Cells[0].Value.ToString();
            btnBack.Visible = true;
            //this.ActiveControl = tbxQtySheetsAffected;

            //loop through datagridview to see if each value of field "Section Name" is the same
            string x;
            string y;
            x = dataGridView1.Rows[0].Cells[15].Value.ToString();
            y = dataGridView1.Rows[0].Cells[15].Value.ToString();
            for (int i = 1; i < this.dataGridView1.Rows.Count; i++)
            {
                y = dataGridView1.Rows[i].Cells[15].Value.ToString();
            }
            if (x == y)
            {
                pnlPalletCard3.BringToFront();
                string d = dataGridView1.Rows[0].Cells[11].Value.ToString();
                lbl2.Text = d;
                lbl2.Visible = true;
                //lblBack5.Visible = false;
                //lblBack6.Visible = false;
                index = 9;
                sectionBtns = true;
            }
            else
            { //prevent section buttons from drawing again if back button is selected
                if (!sectionBtns)
                {
                    //loop through datagrid rows to create a button for each value of field "Section Name"                  
                    for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
                    {
                        //if datagrid is not empty create a button for each row at cells[2] - "Name"
                        if (!(string.IsNullOrEmpty(this.dataGridView1.Rows[i].Cells[15].Value as string)))

                            //offer only one button where SectionName field has two rows with the same value
                            dataGridView1.AllowUserToAddRows = true;
                        if (!(this.dataGridView1.Rows[i].Cells[15].Value as string == this.dataGridView1.Rows[i + 1].Cells[15].Value as string))
                        {
                            {
                                for (int j = 0; j < 1; j++)
                                {
                                    Button btn = new Button();
                                    this.pnlPalletCard1.Controls.Add(btn);
                                    btn.Top = A * 100;
                                    btn.Height = 80;
                                    btn.Width = 465;
                                    btn.BackColor = Color.SteelBlue;
                                    btn.Font = new Font("Microsoft Sans Serif", 14);
                                    btn.ForeColor = Color.White;
                                    btn.Left = 30;
                                    btn.Text = this.dataGridView1.Rows[i].Cells[15].Value as string;
                                    A = A + 1;
                                    btn.Click += new System.EventHandler(this.expr3);
                                }
                            }
                        }
                        dataGridView1.AllowUserToAddRows = false;
                    }
                }
                sectionBtns = true;
            }
        }
        //Dynamic button click - Section buttons, Pallet Card work flow
        private void expr3(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            pnlPalletCard2.BringToFront();
            lbl2.Visible = true;
            lbl2.Text = btn.Text;
            index = 9;

            //filter datagridview1 with the button text choice
            try
            {
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = "SectionName = '" + btn.Text + "' and JobNo like '%" + lblJobNo.Text + "%'";
            }
            catch (Exception) { }





            if (!sigBtns)
            {
                //loop through datagrid rows to create a button for each value of field "PaperSectionNo"                  
                for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
                {
                    //if datagrid is not empty create a button for each row at cells[19] - "PaperSectionNo"
                    //if (!(string.IsNullOrEmpty(this.dataGridView1.Rows[i].Cells[19].Value as string)))

                    //dataGridView1.AllowUserToAddRows = true;
                    //if (!(this.dataGridView1.Rows[i].Cells[19].Value as string == this.dataGridView1.Rows[i + 1].Cells[19].Value as string))
                    //{
                    //    {
                            for (int j = 0; j < 1; j++)
                    {
                        Button btnSig = new Button();
                        this.flowLayoutPanel1.Controls.Add(btnSig);
                        btnSig.Height = 70;
                        btnSig.Width = 120;
                        btnSig.BackColor = Color.SteelBlue;
                        btnSig.Font = new Font("Microsoft Sans Serif", 14);
                        btnSig.ForeColor = Color.White;
                        btnSig.Left = 30;
                        btnSig.Text = "Sig " + this.dataGridView1.Rows[i].Cells[19].Value as string;
                        btnSig.Click += new System.EventHandler(this.DynamicSigBtn);
                    }
                    //    }
                    //}
                    //dataGridView1.AllowUserToAddRows = false;

                }
            }
            sigBtns = true;
        }


        //Dynamic button click - Sig buttons, Pallet Card work flow
        private void DynamicSigBtn(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            pnlPalletCard3.BringToFront();
            lbl3.Visible = true;
            lbl3.Text = btn.Text;
            index = 10;

            //filter datagridview1 with the button text choice
            try
            {
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = "PaperSectionNo like '%" + btn.Text + "%' and JobNo like '%" + lblJobNo.Text + "%'";
            }
            catch (Exception) { }
        }

    }
}
