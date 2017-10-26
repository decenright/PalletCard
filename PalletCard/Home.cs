using System;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Data.SqlClient;
using System.Drawing.Imaging;
using System.Text.RegularExpressions;
using System.IO;
using System.ComponentModel;
using System.Net.Mail;

namespace PalletCard
{
    public partial class Home : Form
    {
        List<Panel> listPanel = new List<Panel>();
        List<string> disableSectionButtons = new List<string>();
        List<string> allSections = new List<string>();
        List<string> completedSections = new List<string>();
        List<string> numberBadList = new List<string>(0);
        List<string> sheetsAffectedList = new List<string>(0);
        List<int> wholePalletList = new List<int>(0);
        int resourceID = 6;
        int index;
        bool sectionBtns;
        bool sigBtns;
        bool badSectionLbls;
        bool backupRequired;
        bool varnishRequired;
        int A = 1;
        string jobNo;
        bool searchChanged;
        int required;
        int produced;
        int shortBy;
        int overBy;
        int oversCalc;
        int PalletNumber;
        int PaperSectionNo;
        int numberUp;
        int qtyRequired;
        int sheetsProduced;
        string badQty;
        string sheetsAffected;
        int gangRow;
        int sheetsAffectedBadSection;
        string autoNum;
        int gangWholePalletButtonPressed;
        DateTime CurrentDate= DateTime.Now;
        decimal maxPercentageShort;
        int notGangedWholePalletValue;
        string ConnectionString = Convert.ToString("Dsn=TharData;uid=tharuser");

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
                tbxPalletHeightPalletCard.Text = "";
                lblSheetCountPalletCard.Text = "";
                tbxSheetCountPalletCard.Text = "";
                lblPheightPalletCard.Text = "";
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
                    tbxPalletHeightPalletCard.Text = "";
                    lblSheetCountPalletCard.Text = "";
                    tbxSheetCountPalletCard.Text = "";
                    lblPheightPalletCard.Text = "";
                    index = 1;
                }
            }
            else if (index == 10)
            {
                Search();
                pnlPalletCard2.BringToFront();
                lbl3.Visible = false;
                lbl4.Visible = false;
                tbxPalletHeightPalletCard.Text = "";
                lblSheetCountPalletCard.Text = "";
                tbxSheetCountPalletCard.Text = "";
                lblPheightPalletCard.Text = "";
                btnBack.Visible = true;
                index = 9;
                // if no sig buttons go straight back to Choose Section screen
                if (flowLayoutPanel1.Controls.Count == 0)
                {
                    pnlPalletCard1.BringToFront();
                    lbl2.Visible = false;
                    lbl3.Visible = false;
                    lbl3.Visible = false;
                    lbl4.Visible = false;
                    tbxPalletHeightPalletCard.Text = "";
                    lblSheetCountPalletCard.Text = "";
                    tbxSheetCountPalletCard.Text = "";
                    lblPheightPalletCard.Text = "";
                    btnBack.Visible = true;
                    index = 1;
                }
            }
            else if (index == 11)
            {
                pnlPalletCard3.BringToFront();
                lbl4.Visible = false;
                lbl5.Visible = false;
                tbxPalletHeightPalletCard.Text = "";
                lblSheetCountPalletCard.Text = "";
                tbxSheetCountPalletCard.Text = "";
                lblPheightPalletCard.Text = "";
                index = 10;
            }
            else if (index == 12)
            {
                pnlPalletCard4.BringToFront();
                lbl6.Visible = false;
                this.ActiveControl = tbxExtraInfoComment;
                flowLayoutPanel2.Enabled = true;
                tbxSheetsAffectedBadSection.Enabled = true;
                tbxSheetsAffectedBadSection.Text = "";
                badSectionLbls = false;
                btnBadSectionOK.Visible = false;
                numberBadList.Clear();
                sheetsAffectedList.Clear();
                index = 11;
            }
            else if (index == 13)
            {
                pnlPalletCard5.BringToFront();
                lbl7.Visible = false;
                lbl6.Visible = false;
                lblNumberUp.Visible = false;
                lblNumberUpBadQty.Visible = false;
                tbxSheetsAffectedBadSection.Text = "";
                tbxTextBoxBadSection.Text = "";
                index = 12;
            }
            else if (index == 14)
            {
                pnlPalletCard4.BringToFront();
                tbxExtraInfoComment.Text = "";
                index = 13;
            }

            else if (index == 15)
            {
                pnlPalletCard7.BringToFront();
                index = 14;
            }

            else if (index == 16)
            {
                pnlPalletCard8.BringToFront();
                index = 15;
            }

            else if (index == 17)
            {
                pnlPalletCard8.BringToFront();
                btnIsSectionFinishedYes.Enabled = true;
                btnIsSectionFinishedNo.Enabled = true;
                btnIsSectionFinishedYes.BackColor = System.Drawing.Color.SteelBlue;
                btnIsSectionFinishedNo.BackColor = System.Drawing.Color.SteelBlue;
                index = 15;
            }
            else if (index == 18)
            {
                pnlNotification1.BringToFront();
                lbl2.Visible = false;
                lbl3.Visible = false;
                lbl4.Visible = false;
                index = 2;
                // if no section buttons go straight back to Choose Action screen
                if (pnlNotification1.Controls.Count == 0)
                {
                    pnlHome1.BringToFront();
                    lbl1.Visible = false;
                    lbl2.Visible = false;
                    lbl3.Visible = false;
                    lbl4.Visible = false;
                    lbl3.Visible = false;
                    lbl4.Visible = false;
                    btnBack.Visible = false;
                    index = 1;
                }
            }
        }

        private void Home_Load(object sender, EventArgs e)
        {
            //string ConnectionString = Convert.ToString("Dsn=TharTest;uid=tharuser");
            string CommandText = "SELECT * FROM app_PalletOperations where resourceID = '" + resourceID + "'";
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
                //dataGridView1.DataSource = operations;

                // New table to hold concatenated values - if duplicate line filter column = 0. Non duplicate lines filter column = 1. 
                // Search function includes a filter on 1 to return only the 1's
                DataTable concatenatedTable = new DataTable();
                concatenatedTable = operations.Clone();
                concatenatedTable.Columns.Add("Concat");
                concatenatedTable.Columns.Add("Filter");
                concatenatedTable.Columns["Concat"].Expression = "JobNo+ ',' + PrintMethod + ',' + PaperSectionID";

                foreach (DataRow dr in operations.Rows)
                {
                    concatenatedTable.Rows.Add(dr.ItemArray);
                }

                for (int i = 0; i < concatenatedTable.Rows.Count - 1; i++)
                    if (concatenatedTable.Rows[i][30].ToString() == concatenatedTable.Rows[i + 1][30].ToString())
                    {
                        concatenatedTable.Rows[i][31] = 0;
                    }
                    else
                        concatenatedTable.Rows[i][31] = 1;
                        concatenatedTable.Rows[concatenatedTable.Rows.Count -1][31] = 1;


                dataGridView1.DataSource = concatenatedTable;
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
                    ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("JobNo like '%{0}%' and filter = '1' ", tbxSearchBox.Text.Trim().Replace("'", "''"));
                    lblJobNo.Text = dataGridView1.Rows[0].Cells[0].Value.ToString();
                    lblJobNo.Visible = true;
                    //int resourceID = (int)dataGridView1.Rows[0].Cells[1].Value;
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
                    pnlNotification1.Controls.Clear();
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
            //string ConnectionString = Convert.ToString("Dsn=TharTest;uid=tharuser");
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
                // New table to hold concatenated values - if duplicate line filter column = 0. Non duplicate lines filter column = 1. 
                // Search function includes a filter on 1 to return only the 1's
                DataTable concatenatedTable = new DataTable();
                concatenatedTable = operations.Clone();
                concatenatedTable.Columns.Add("Concat");
                concatenatedTable.Columns.Add("Filter");
                concatenatedTable.Columns["Concat"].Expression = "JobNo+ ',' + PrintMethod + ',' + PaperSectionID";

                foreach (DataRow dr in operations.Rows)
                {
                    concatenatedTable.Rows.Add(dr.ItemArray);
                }

                for (int i = 0; i < concatenatedTable.Rows.Count - 1; i++)
                    if (concatenatedTable.Rows[i][30].ToString() == concatenatedTable.Rows[i + 1][30].ToString())
                    {
                        concatenatedTable.Rows[i][31] = 0;
                    }
                    else
                        concatenatedTable.Rows[i][31] = 1;
                        concatenatedTable.Rows[concatenatedTable.Rows.Count - 1][31] = 1;

                dataGridView1.DataSource = concatenatedTable;
            }
            pnlHome0.BringToFront();
            lblJobNo.Visible = false;
            lblPress.Visible = false;
            lbl1.Visible = false;
            lbl2.Visible = false;
            lbl3.Visible = false;
            lbl4.Visible = false;
            lbl5.Visible = false;
            lbl6.Visible = false;
            lbl7.Visible = false;
            lblPheight.Text = "";
            tbxSearchBox.Text = "";
            tbxSearchBox.Focus();
            sectionBtns = false;
            tbxPalletHeight.Text = null;
            btnSearch.Visible = true;
            btnBack.Visible = false;
            this.ActiveControl = tbxSearchBox;
            badSectionLbls = false;
            pnlPalletCard1.Controls.Clear();
            tbxPalletHeightPalletCard.Clear();
            tbxSheetCountPalletCard.Clear();
            lblSheetCountPalletCard.Text = "";
            lblPheightPalletCard.Text = "";
            backupRequired = false;
            varnishRequired = false;
            tbxSheetsAffectedBadSection.Text = "";
            tbxTextBoxBadSection.Text = "";
            tbxExtraInfoComment.Text = "";
            lblPC_IncompletePallet.Visible = false;
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
            lblPrint1.MaximumSize = new Size(450, 220);
            lblPrint1.AutoSize = true;
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
                    lblPheight.Text = (r1 + " Sheets");
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

            string sqlFormattedDate = CurrentDate.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
            string Query = "insert into Log (Routine, JobNo, ResourceID, Expr1, WorkingSize, SheetQty, Description, Timestamp1) values('" + this.lbl1.Text + "','" + this.dataGridView1.Rows[0].Cells[0].Value + "','" + resourceID + "','" + this.lbl2.Text + "','" + this.lbl4.Text + "','" + this.lblPrint3.Text + "','" + this.lbl3.Text + "','" + CurrentDate + "');";
            SqlConnection conDatabase = new SqlConnection(constring);
            SqlCommand cmdDatabase = new SqlCommand(Query, conDatabase);
            SqlDataReader myReader;
            try
            {
                conDatabase.Open();
                myReader = cmdDatabase.ExecuteReader();
                conDatabase.Close();
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
            lblPrint14.MaximumSize = new Size(450, 220);
            lblPrint14.AutoSize = true;
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
            //string Query = "insert into Log (Routine, JobNo, ResourceID, Description, WorkingSize, SheetQty) values('" + this.lbl1.Text + "','" + this.dataGridView1.Rows[0].Cells[0].Value + "','" + resourceID + "','" + this.lbl2.Text + "','" + this.lbl4.Text + "','" + this.lblPrint3.Text + "');";
            //SqlConnection conDatabase = new SqlConnection(constring);
            //SqlCommand cmdDatabase = new SqlCommand(Query, conDatabase);
            //SqlDataReader myReader;
            //try
            //{
            //   conDatabase.Open();
            //   myReader = cmdDatabase.ExecuteReader();
            //   conDatabase.Close();
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
        private object tb1;

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

        //private static Bitmap DrawControlToBitmap(Control control)
        //{
        //    Bitmap bitmap = new Bitmap(control.Width, control.Height);
        //    Graphics graphics = Graphics.FromImage(bitmap);
        //    Rectangle rect = control.RectangleToScreen(control.ClientRectangle);
        //    graphics.CopyFromScreen(rect.Location, Point.Empty, control.Size);
        //    return bitmap;
        //}

        public void SaveSignatureImageToFile()
        {
            Bitmap bmp = new Bitmap(this.SignaturePanel.Width, this.SignaturePanel.Height);
            Graphics graphics = Graphics.FromImage(bmp);
            Rectangle rect = SignaturePanel.RectangleToScreen(SignaturePanel.ClientRectangle);
            graphics.CopyFromScreen(rect.Location, Point.Empty, SignaturePanel.Size);
            bmp.Save("C:/Temp/Signatures/ " + autoNum + ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
        }

        private void btnQATravellerBlurb_Click(object sender, EventArgs e)
        {
            SaveSignatureImageToFile();
            // Requery the data to refresh dataGridView2 with the newly added PalletNumber and barCode
            string ConnectionString = Convert.ToString("Dsn=PalletCard;uid=PalletCardAdmin");
            string CommandText = "SELECT * FROM Log where JobNo = '" + lblJobNo.Text + "'";
            OdbcConnection myConnection = new OdbcConnection(ConnectionString);
            OdbcCommand myCommand = new OdbcCommand(CommandText, myConnection);
            OdbcDataAdapter myAdapter = new OdbcDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet palletCardData = new DataSet();
            try
            {
                myConnection.Open();
                myAdapter.Fill(palletCardData);
            }
            catch (Exception ex)
            {
                throw (ex);
            }
            finally
            {
                myConnection.Close();
            }
            using (DataTable palletCardLog = new DataTable())
            {
                myAdapter.Fill(palletCardLog);
                dataGridView2.DataSource = palletCardLog;
            }

            this.dataGridView2.Sort(this.dataGridView2.Columns["PalletNumber"], ListSortDirection.Descending);
            string barCode = Convert.ToString(((int)dataGridView2.Rows[0].Cells[5].Value));
            Bitmap bitMap = new Bitmap(barCode.Length * 40, 80);
            using (Graphics graphics = Graphics.FromImage(bitMap))
            {
                Font oFont = new Font("IDAutomationHC39M", 16);
                PointF point = new PointF(2f, 2f);
                SolidBrush blackBrush = new SolidBrush(Color.Black);
                SolidBrush whiteBrush = new SolidBrush(Color.White);
                graphics.FillRectangle(whiteBrush, 0, 0, bitMap.Width, bitMap.Height);
                graphics.DrawString("*" + barCode + "*", oFont, blackBrush, point);
            }
            using (MemoryStream ms = new MemoryStream())
            {
                bitMap.Save(ms, ImageFormat.Png);
                pictureBox1.Image = bitMap;
                pictureBox1.Height = bitMap.Height;
                pictureBox1.Width = bitMap.Width;
            }

            pnlPalletCardPrint.BringToFront();
            lblPC_JobNo.Text = lblJobNo.Text;
            lblPC_JobNo.Visible = true;
            lblPC_Customer.Text = dataGridView1.Rows[0].Cells[22].Value as string;
            lblPC_Customer.Visible = true;
            lblPC_Customer.MaximumSize = new Size(450, 220);
            lblPC_Customer.AutoSize = true;
            lblPC_SheetQty.Text = lbl5.Text;
            lblPC_SheetQty.Visible = true;
            lblPC_Press.Text = "Press - " + lblPress.Text;
            lblPC_Press.Visible = true;
            lblPC_Date.Text = "Date - " + DateTime.Now.ToString("d/M/yyyy");
            lblPC_Date.Visible = true;
            lblPC_Note.Text = tbxExtraInfoComment.Text + " - " + tbxTextBoxBadSection.Text;
            lblPC_Note.Visible = true;
            lblPC_PalletNumber.Text = "Pallet No " + PalletNumber.ToString();
            lblPC_PalletNumber.Visible = true;
            lblPC_Sig.Text = "Sheet " + dataGridView2.Rows[0].Cells[8].Value as string;
            lblPC_Sig.Visible = true;
            btnCancel.Visible = false;
            index = 17;
        }





//****************************************************************************************************
//   PALLET CARD
//****************************************************************************************************

        private void btnPalletCard_Click(object sender, EventArgs e)
        {
            lbl1.Visible = true;
            lbl1.Text = "Pallet Card";
            pnlPalletCard1.BringToFront();
            index = 8;
            tbxPalletHeightPalletCard.Focus();
            btnBack.Visible = true;            

            //loop through datagridview to see if each value of field "SectionName" is the same
            string x;
            string y;
            //x = SectionName value at Row 0
            x = dataGridView1.Rows[0].Cells[15].Value.ToString();
            //initialize variable y with a value - this will change to Section Name value at row 1 once it enters the loop
            y = dataGridView1.Rows[0].Cells[15].Value.ToString();
            //***************Check if SectionName field is empty
            if (!(x == ""))
            { 
                for (int i = 1; i < this.dataGridView1.Rows.Count; i++)
                {
                    //y = SectionName value at Row 1
                    y = dataGridView1.Rows[i].Cells[15].Value.ToString();                    
                }
                int rowCountA = dataGridView1.Rows.Count;
                if (x == y || rowCountA == 1)
                {
                    pnlPalletCard3.BringToFront();
                    string sectionName = dataGridView1.Rows[0].Cells[15].Value.ToString();
                    string sig = dataGridView1.Rows[0].Cells[19].Value.ToString();
                    lbl2.Text = sectionName;
                    lbl2.Visible = true;
                    lbl3.Text = "Sheet " + sig;
                    lbl3.Visible = true;
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
                            //if datagrid is not empty create a button for each row at cells[15] - "Section Name"
                            if (!(string.IsNullOrEmpty(this.dataGridView1.Rows[i].Cells[15].Value as string)))

                                //offer only one button where SectionName field has more than one row with the same value
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
                                        btn.Click += new System.EventHandler(this.sectionNameSectionBtns);
                                    }
                                }
                            }
                            dataGridView1.AllowUserToAddRows = false;
                        }
                    }
                    sectionBtns = true;                   
                }
            }
            //************ ELSE USE EXPR1 FIELD
            else
            { 
                //x = Expr1 value at Row 0
                x = dataGridView1.Rows[0].Cells[11].Value.ToString();
                {
                    for (int i = 1; i < this.dataGridView1.Rows.Count; i++)
                    {
                        //y = Expr1 value at Row 1
                        y = dataGridView1.Rows[i].Cells[11].Value.ToString();
                    }
                    int rowCountB = dataGridView1.Rows.Count;
                    if (x == y || rowCountB == 1)
                    {
                        pnlPalletCard3.BringToFront();
                        string Expr1 = dataGridView1.Rows[0].Cells[11].Value.ToString();
                        string sig = dataGridView1.Rows[0].Cells[19].Value.ToString();
                        lbl2.Text = Expr1;
                        lbl2.Visible = true;
                        lbl3.Text = "Sheet " + sig;
                        lbl3.Visible = true;
                        index = 9;
                        sectionBtns = true;
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

                                        //offer only one button where Expr1 field has more than one row with the same value
                                        dataGridView1.AllowUserToAddRows = true;
                                    if (!(this.dataGridView1.Rows[i].Cells[11].Value as string == this.dataGridView1.Rows[i + 1].Cells[11].Value as string))
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
                                                btn.Text = this.dataGridView1.Rows[i].Cells[11].Value as string;
                                                A = A + 1;
                                                btn.Click += new System.EventHandler(this.expr1SectionBtns);
                                            }
                                        }
                                    }
                                    dataGridView1.AllowUserToAddRows = false;

                                }
                            }
                            sectionBtns = true;
                        }
                 }
             }
        }

        //Dynamic button click - Section buttons, SECTION NAME, Pallet Card work flow
        private void sectionNameSectionBtns(object sender, EventArgs e)
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

            string x;
            string y;
            //x = PaperSectionNo value at Row 0
            x = dataGridView1.Rows[0].Cells[19].Value.ToString();
            //initialize variable y with a value - this will change to PaperSectionNo value at row 1 once it enters the loop
            y = dataGridView1.Rows[0].Cells[19].Value.ToString();
            {
                for (int i = 1; i < this.dataGridView1.Rows.Count; i++)
                {
                    //y = PaperSectionNo value at Row 1
                    y = dataGridView1.Rows[i].Cells[19].Value.ToString();
                }
                if (x == y)
                {
                    pnlPalletCard3.BringToFront();
                    string sig = dataGridView1.Rows[0].Cells[19].Value.ToString();
                    lbl2.Text = dataGridView1.Rows[0].Cells[15].Value.ToString();
                    lbl2.Visible = true;
                    lbl3.Text = "Sheet " + sig;
                    lbl3.Visible = true;
                    index = 10;
                    sectionBtns = true;
                }

                else {
                    if (!sigBtns)
                        {                           
                            for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
                            {
                                for (int j = 0; j < 1; j++)
                                {
                                    dataGridView1.AllowUserToAddRows = true;
                                    if (!(this.dataGridView1.Rows[i].Cells[19].Value  == this.dataGridView1.Rows[i + 1].Cells[19].Value ))
                                    {
                                        for (int k = 0; k < 1; k++)
                                        {
                                            Button btnSig = new Button();
                                            this.flowLayoutPanel1.Controls.Add(btnSig);
                                            btnSig.Text = dataGridView1.Rows[i].Cells[19].Value.ToString();
                                            if (disableSectionButtons.Contains(btnSig.Text))
                                            {
                                                btnSig.BackColor = Color.Silver;
                                                btnSig.Enabled = false;
                                            }
                                            else
                                            {
                                                btnSig.BackColor = Color.SteelBlue;
                                            }
                                            btnSig.Height = 70;
                                            btnSig.Width = 120;                                           
                                            btnSig.Font = new Font("Microsoft Sans Serif", 20);
                                            btnSig.ForeColor = Color.White;
                                            btnSig.TextAlign = ContentAlignment.MiddleCenter;
                                            btnSig.Click += new System.EventHandler(this.sectionButtonSectionName);
                                        }
                                    }
                                }
                                dataGridView1.AllowUserToAddRows = false;
                                // get all the section numbers in a list for later
                                allSections.Add(dataGridView1.Rows[i].Cells[19].Value.ToString());
                            }
                        }
                        sigBtns = true;
                    }

                string ConnectionString = Convert.ToString("Dsn=PalletCard;uid=PalletCardAdmin");
                string CommandText = "SELECT * FROM Log where JobNo = '" + lblJobNo.Text + "'";
                OdbcConnection myConnection = new OdbcConnection(ConnectionString);
                OdbcCommand myCommand = new OdbcCommand(CommandText, myConnection);
                OdbcDataAdapter myAdapter = new OdbcDataAdapter();
                myAdapter.SelectCommand = myCommand;
                DataSet palletCardData = new DataSet();
                try
                {
                    myConnection.Open();
                    myAdapter.Fill(palletCardData);
                }
                catch (Exception ex)
                {
                    throw (ex);
                }
                finally
                {
                    myConnection.Close();
                }
                using (DataTable palletCardLog = new DataTable())
                {
                    myAdapter.Fill(palletCardLog);
                    dataGridView2.DataSource = palletCardLog;
                }

                //for (int i = 0; i < this.dataGridView2.Rows.Count; i++)
                //{
                //    if (Convert.ToByte(dataGridView2.Rows[i].Cells[7].Value) == 1)
                //    {

                //            //flowLayoutPanel1.Controls.Remove();

                //    }
                //}

            }
        }

        //Dynamic button click - Section buttons, EXPR1, Pallet Card work flow
        private void expr1SectionBtns(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            pnlPalletCard2.BringToFront();
            lbl2.Visible = true;
            lbl2.Text = btn.Text;
            index = 9;

            //filter datagridview1 with the button text choice
            try
            {
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = "Expr1 = '" + btn.Text + "' and JobNo like '%" + lblJobNo.Text + "%'";
            }
            catch (Exception) { }

            string x;
            string y;
            //x = PaperSectionNo value at Row 0
            x = dataGridView1.Rows[0].Cells[19].Value.ToString();
            //initialize variable y with a value - this will change to PaperSectionNo value at row 1 once it enters the loop
            y = dataGridView1.Rows[0].Cells[19].Value.ToString();
            {
                for (int i = 1; i < this.dataGridView1.Rows.Count; i++)
                {
                    //y = PaperSectionNo value at Row 1
                    y = dataGridView1.Rows[i].Cells[19].Value.ToString();
                }
                if (x == y)
                {
                    pnlPalletCard3.BringToFront();
                    string sig = dataGridView1.Rows[0].Cells[19].Value.ToString();
                    lbl2.Text = dataGridView1.Rows[0].Cells[11].Value.ToString();
                    lbl2.Visible = true;
                    lbl3.Text = "Sheet " + sig;
                    lbl3.Visible = true;
                    index = 10;
                    sectionBtns = true;
                }

                else { 
                    if (!sigBtns)
                        {
                        //loop through datagrid rows to create a button for each value of field "PaperSectionNo"  
                            for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
                            {
                                var v = dataGridView1.Rows[i].Cells[19].Value;
                                for (int j = 0; j < 1; j++)
                                {
                                    dataGridView1.AllowUserToAddRows = true;
                                    if (!(this.dataGridView1.Rows[i].Cells[19].Value == this.dataGridView1.Rows[i + 1].Cells[19].Value))
                                    {
                                        for (int k = 0; k < 1; k++)
                                        {
                                            Button btnSig = new Button();
                                            this.flowLayoutPanel1.Controls.Add(btnSig);
                                            btnSig.Text = this.dataGridView1.Rows[i].Cells[19].Value.ToString();
                                            if (disableSectionButtons.Contains(btnSig.Text))
                                            {
                                                btnSig.BackColor = Color.Silver;
                                                btnSig.Enabled = false;
                                            }
                                            else
                                            {
                                                btnSig.BackColor = Color.SteelBlue;
                                            }
                                            btnSig.Height = 70;
                                            btnSig.Width = 120;
                                            btnSig.Font = new Font("Microsoft Sans Serif", 20);
                                            btnSig.ForeColor = Color.White;
                                            btnSig.TextAlign = ContentAlignment.MiddleCenter;
                                            btnSig.Click += new System.EventHandler(this.sectionButtonExpr1);
                                        }
                                    }
                                    dataGridView1.AllowUserToAddRows = false;
                                    // get all the section numbers in a list for later
                                    allSections.Add(dataGridView1.Rows[i].Cells[19].Value.ToString());
                                }
                            }
                        }
                        sigBtns = true;
                    }
              }
        }

//Dynamic button click - Section buttons SectionName, Pallet Card work flow
        private void sectionButtonSectionName(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            pnlPalletCard3.BringToFront();
            lbl3.Visible = true;
            lbl3.Text = "Sheet " + btn.Text;
            tbxPalletHeightPalletCard.Focus();
            index = 10;

            //filter datagridview1 with the button text choice
            try
            {
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = "SectionName like '%" + lbl2.Text + "%'  and PaperSectionNo = " + btn.Text.Trim() + " and JobNo like '%" + lblJobNo.Text + "%'";
            }
            catch (Exception) { }
        }

//Dynamic button click - Section buttons EXPR1, Pallet Card work flow
        private void sectionButtonExpr1(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            pnlPalletCard3.BringToFront();
            lbl3.Visible = true;
            lbl3.Text = "Sheet " + btn.Text;
            tbxPalletHeightPalletCard.Focus();
            index = 10;

            //filter datagridview1 with the button text choice
            try
            {
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = "Expr1 like '%" + lbl2.Text + "%'  and PaperSectionNo = " + btn.Text.Trim() + " and JobNo like '%" + lblJobNo.Text + "%'";
            }
            catch (Exception) { }
        }

// Pallet Height textBox calculation for Pallet Card
        private void tbxPalletHeightPalletCard_TextChanged(object sender, EventArgs e)
        {
            try
            {
                TextBox objTextBox = (TextBox)sender;
                double p1;
                double p2;
                if (!String.IsNullOrEmpty(tbxPalletHeightPalletCard.Text))
                {
                    p1 = Convert.ToInt32(objTextBox.Text);
                    p2 = Convert.ToInt32(this.dataGridView1.Rows[0].Cells[20].Value);
                    double result = Math.Ceiling(p1 / (p2 / 1000));
                    string r1 = Convert.ToString(result);
                    lblSheetCountPalletCard.Text = (r1 + " Sheets");
                }
            }
            catch
            {
                MessageBox.Show("Please enter a valid number");
            }
        }

        // Sheet Count textBox calculation for Pallet Card
        private void tbxSheetCountPalletCard_TextChanged(object sender, EventArgs e)
        {
            try
            {
                TextBox objTextBox = (TextBox)sender;
                double p1;
                double p2;
                if (!String.IsNullOrEmpty(tbxSheetCountPalletCard.Text))
                {
                    p1 = Convert.ToInt32(objTextBox.Text);
                    p2 = Convert.ToInt32(this.dataGridView1.Rows[0].Cells[20].Value);
                    double result = (int)(p1 * (p2 / 1000));
                    string r1 = Convert.ToString(result);
                    lblPheightPalletCard.Text = (r1 + " mm");
                }
            }
            catch
            {
                MessageBox.Show("Please enter a valid number");
            }
        }

        private void btnPalletHeightSheetCountPalletCard_Click(object sender, EventArgs e)
        {            
                if (!string.IsNullOrEmpty(tbxPalletHeightPalletCard.Text) && !string.IsNullOrEmpty(tbxSheetCountPalletCard.Text))
                {
                    lbl4.Text = tbxPalletHeightPalletCard.Text + " mm";
                    lbl4.Visible = true;
                    lbl5.Text = tbxSheetCountPalletCard.Text + " Sheets";
                    lbl5.Visible = true;
                }
                else if (!string.IsNullOrEmpty(tbxSheetCountPalletCard.Text))
                {
                    lbl5.Text = tbxSheetCountPalletCard.Text + " Sheets";
                    lbl5.Visible = true;
                    lbl4.Text = lblPheightPalletCard.Text;
                    lbl4.Visible = true;
                }
                else if (!string.IsNullOrEmpty(tbxPalletHeightPalletCard.Text))
                {
                    lbl4.Text = tbxPalletHeightPalletCard.Text + " mm";
                    lbl4.Visible = true;
                    lbl5.Text = lblSheetCountPalletCard.Text;
                    lbl5.Visible = true;
                }


            if (lbl4.Visible == false)
            {
                MessageBox.Show("please enter a value");
            }
            else
                { 
                pnlPalletCard4.BringToFront();
                this.ActiveControl = tbxExtraInfoComment;
                index = 11;
                }
        }

        // GANGPRO ROUTINE
        private void queryGangpro()
        {
            //string ConnectionString = Convert.ToString("Dsn=TharTest;uid=tharuser");
            string CommandText = "SELECT * FROM app_PalletGangPro where Prod_Job = '" + lblJobNo.Text + "'";
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

            sheetsProduced = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
            qtyRequired = Convert.ToInt32(dataGridView1.Rows[0].Cells[26].Value);
            using (DataTable gangPro = new DataTable())
            {
                myAdapter.Fill(gangPro);
                DataTable gangProTable = new DataTable();
                gangProTable = gangPro.Clone();
                gangProTable.Columns.Add("QtyRequired");
                gangProTable.Columns.Add("NumberUp1");
                gangProTable.Columns.Add("Prod_qty_required");
                gangProTable.Columns.Add("NumberUp-Bad");
                gangProTable.Columns.Add("Sheets Affected");
                gangProTable.Columns.Add("Sheets unaffected");
                gangProTable.Columns.Add("Prod_qty_Produced");
                gangProTable.Columns.Add("Qty_Short");
                gangProTable.Columns.Add("Percentage_Short");

                gangProTable.Columns["QtyRequired"].Expression = " '" + qtyRequired + "'  ";
                for (int i = 0; i < gangProTable.Rows.Count; i++)
                {
                    gangProTable.Columns["NumberUp1"].Expression = gangProTable.Columns["NumberUp"].Expression;
                }
                gangProTable.Columns["Qty_Short"].Expression = " '" + sheetsProduced + "'  ";
                gangProTable.Columns["Prod_qty_required"].Expression = "QtyRequired * NumberUp1";
                gangProTable.Columns["Sheets Affected"].Expression = tbxSheetsAffectedBadSection.Text;

                foreach (DataRow dr in gangPro.Rows)
                {
                    gangProTable.Rows.Add(dr.ItemArray);
                }
                dataGridView3.DataSource = gangProTable;
            }
        }

        // GANGCLASSIC ROUTINE
        private void queryGangClassic()
        {
            //string ConnectionString = Convert.ToString("Dsn=TharTest;uid=tharuser");
            string CommandText1 = "SELECT * FROM app_PalletGangClassic where Prod_Job = '" + lblJobNo.Text + "'";
            OdbcConnection myConnection1 = new OdbcConnection(ConnectionString);
            OdbcCommand myCommand1 = new OdbcCommand(CommandText1, myConnection1);
            OdbcDataAdapter myAdapter1 = new OdbcDataAdapter();
            myAdapter1.SelectCommand = myCommand1;
            DataSet tharData1 = new DataSet();
            try
            {
                myConnection1.Open();
                myAdapter1.Fill(tharData1);
            }
            catch (Exception ex)
            {
                throw (ex);
            }
            finally
            {
                myConnection1.Close();
            }

            sheetsProduced = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
            qtyRequired = Convert.ToInt32(dataGridView1.Rows[0].Cells[26].Value);
            using (DataTable gangClassic = new DataTable())
            {
                myAdapter1.Fill(gangClassic);
                gangClassic.Columns.Add("QtyRequired", typeof(int));
                gangClassic.Columns.Add("NumberUp");
                gangClassic.Columns.Add("NumberUp1");
                gangClassic.Columns.Add("NumberUp2", typeof(int));
                gangClassic.Columns.Add("ProdQtyRequired", typeof(int));
                gangClassic.Columns.Add("NumberUpBad", typeof(int));
                gangClassic.Columns.Add("SheetsAffected", typeof(int));
                gangClassic.Columns.Add("SheetsUnaffected", typeof(int));
                gangClassic.Columns.Add("SheetsProduced", typeof(int));
                gangClassic.Columns.Add("QtyGoodProduced", typeof(int));
                gangClassic.Columns.Add("QtyShort", typeof(int));
                gangClassic.Columns.Add("PercentageShort", typeof(decimal));
                gangClassic.Columns.Add("PercentageShort1", typeof(string));
                gangClassic.Columns.Add("PercentageShort2", typeof(decimal));
                gangClassic.Columns.Add("PercentageShort3");

                DataTable gangClassicTable = new DataTable();
                gangClassicTable = gangClassic.Clone();

                gangClassicTable.Columns["QtyRequired"].Expression = " '" + qtyRequired + "' ";
                for (int i = 0; i < gangClassic.Rows.Count; i++)
                {
                   var str = gangClassic.Rows[i][4].ToString();
                   gangClassic.Rows[i]["NumberUp"] = " '" + Regex.Match(str, @"(\d+)[^-]*$") + "'  ";
                   gangClassic.Rows[i]["NumberUp1"] = Regex.Replace(gangClassic.Rows[i]["NumberUp"].ToString(), "[^0-9.]", "");
                   gangClassic.Rows[i]["NumberUp2"] = Convert.ToInt32(gangClassic.Rows[i]["NumberUp1"]);                           
                }
                gangClassicTable.Columns["SheetsProduced"].Expression = " '" + sheetsProduced + "' ";
                gangClassicTable.Columns["ProdQtyRequired"].Expression = "QtyRequired * NumberUp2";
                gangClassicTable.Columns["SheetsAffected"].Expression = tbxSheetsAffectedBadSection.Text;
                gangClassicTable.Columns["SheetsUnaffected"].Expression = "SheetsProduced - SheetsAffected";
                gangClassicTable.Columns["QtyGoodProduced"].Expression = "(SheetsAffected * (NumberUp2 - NumberUpBad)) + (SheetsUnaffected * NumberUp2)";
                gangClassicTable.Columns["QtyShort"].Expression = "QtyGoodProduced - ProdQtyRequired";
                gangClassicTable.Columns["PercentageShort"].Expression = "QtyShort / ProdQtyRequired";
                gangClassicTable.Columns["PercentageShort1"].Expression = "-1.0";
                gangClassicTable.Columns["PercentageShort2"].Expression = "PercentageShort1";
                gangClassicTable.Columns["PercentageShort3"].Expression = "PercentageShort * PercentageShort2";

                if (numberBadList.Count != 0)
                {
                    for (int i = 0; i < gangClassic.Rows.Count; i++)
                    {
                        gangClassic.Rows[i][13] = numberBadList[i];
                        gangClassic.Rows[i][14] = sheetsAffectedList[i];
                        if(gangWholePalletButtonPressed == 1)
                        {
                            gangClassic.Rows[i][14] = wholePalletList[i];
                        }                        
                    }
                }

                foreach (DataRow dr in gangClassic.Rows)
                {
                    gangClassicTable.Rows.Add(dr.ItemArray);
                }
                dataGridView4.DataSource = gangClassicTable;

                // Find Qty bad to return to the main flow
                if (dataGridView4.Rows[0].Cells[22].Value.ToString() != "")
                    {
                    this.dataGridView4.Sort(this.dataGridView4.Columns[22], ListSortDirection.Descending);
                    maxPercentageShort = Convert.ToDecimal(dataGridView4.Rows[0].Cells[22].Value);                    
                    }

                sheetsAffectedBadSection = Convert.ToInt32((qtyRequired * (1 + maxPercentageShort)) - sheetsProduced);
                //MessageBox.Show(sheetsAffectedBadSection.ToString());

                // make sure it doesn't return a negative number
                if (sheetsAffectedBadSection < 0)
                {
                    sheetsAffectedBadSection = 0;
                }

                // Dont show OK button if Number Up(cell 13) and Sheets affected(cell 14) are empty for gang Classic
                for (int i = 0; i < gangClassic.Rows.Count; i++)
                {
                    if (dataGridView4.Rows[i].Cells[13].Value.ToString() != "" & dataGridView4.Rows[i].Cells[13].Value.ToString() != "0" & dataGridView4.Rows[i].Cells[14].Value.ToString() != "" & dataGridView4.Rows[i].Cells[14].Value.ToString() != "0")
                    {
                        btnBadSectionOK.Visible = true;
                    }
                }

            }
        }

        // NOT GANGED ROUTINE
        private void notGanged()
        {
            sheetsAffectedBadSection = Convert.ToInt32(sheetsAffected);

            // make sure it doesn't return a negative number
            if (sheetsAffectedBadSection < 0)
            {
                sheetsAffectedBadSection = 0;
            }

            if (gangWholePalletButtonPressed == 1 )
            {
                sheetsAffectedBadSection = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
            }

            MessageBox.Show(sheetsAffectedBadSection.ToString());

            foreach (Control c in flowLayoutPanel2.Controls)
            {

                if (badQty != null & badQty != "0" & sheetsAffected != null & sheetsAffected != "0")
                {
                    btnBadSectionOK.Visible = true;
                }
            }
        }

        private void tbxSheetsAffectedBadSection_TextChanged(object sender, EventArgs e)
        {
            if (tbxSheetsAffectedBadSection.Text != "")
            {
                sheetsAffectedBadSection = Convert.ToInt32(tbxSheetsAffectedBadSection.Text);
                btnBadSectionOK.Visible = true;
            }          
            lbl7.Text = dataGridView1.Rows[0].Cells[26].Value.ToString();
        }

        private void btnMarkBad_Click(object sender, EventArgs e)
        {
            this.flowLayoutPanel2.Controls.Clear();
            pnlPalletCard5.BringToFront();
            this.ActiveControl = tbxSheetsAffectedBadSection;
            tbxSheetsAffectedBadSection.Visible = true;
            btnBadSectionOK.Visible = false;
            index = 12;

            // Hide FlowLayout2 (Gang Panel) if NumberUp = 1 (will depend on wheteher it is a complete or incomplete i.e scanned line)
            if (dataGridView2.Rows.Count == 0)
                {
                    if (dataGridView1.Rows[0].Cells[12].Value.ToString() == "1")
                    {
                        pnlBadSectionGangHeader.Visible = false;
                        lblStockCode.Visible = false;
                        lblNumberUp.Visible = false;
                        lblNumberUpBadQty.Visible = false;
                        lblSheetsAffected.Visible = false;
                        flowLayoutPanel2.Visible = false;
                        btnScrollUp.Visible = false;
                        btnScrollDown.Visible = false;
                    }
                    numberUp = Convert.ToInt32(dataGridView1.Rows[0].Cells[12].Value);
                }
            else
                    if (dataGridView2.Rows[0].Cells[21].Value.ToString() == "1")
                    {
                        pnlBadSectionGangHeader.Visible = false;
                        lblStockCode.Visible = false;
                        lblNumberUp.Visible = false;
                        lblNumberUpBadQty.Visible = false;
                        lblSheetsAffected.Visible = false;
                        flowLayoutPanel2.Visible = false;
                        btnScrollUp.Visible = false;
                        btnScrollDown.Visible = false;
                        numberUp = Convert.ToInt32(dataGridView2.Rows[0].Cells[21].Value);
                    }

            qtyRequired = Convert.ToInt32(dataGridView1.Rows[0].Cells[26].Value);

            if (numberUp != 1)
            {
                tbxSheetsAffectedBadSection.Visible = false;
                lblSheets_Affected.Visible = false;

                // Filter for Ganged jobs
                if (Convert.ToInt32(dataGridView1.Rows[0].Cells[14].Value) != 0 & Convert.ToInt32(dataGridView1.Rows[0].Cells[14].Value) != 2 )
                {
                    if (Convert.ToInt32(dataGridView1.Rows[0].Cells[14].Value) == 1)
                    {
                        queryGangClassic();
                    }
                    if (Convert.ToInt32(dataGridView1.Rows[0].Cells[14].Value) == 3)
                    {
                        queryGangpro();
                    }

                }

                // if JobGanged = 0 (Main Table)
                if (Convert.ToInt32(dataGridView1.Rows[0].Cells[14].Value) == 0)
                {
                    if (numberUp > 1)
                    {
                        if (!badSectionLbls)
                        {
                            for (int tb2 = 0; tb2 < dataGridView1.Rows.Count; tb2++)
                            {
                                btnScrollUp.Visible = false;
                                btnScrollDown.Visible = false;
                                lblStockCode.Text = "Job\r\nNumber";
                                lblStockCode.Visible = true;
                                lblNumberUp.Visible = true;
                                lblNumberUp.Text = "Number\r\nUp";
                                lblNumberUpBadQty.Visible = true;
                                lblNumberUpBadQty.Text = "Quantity\r\nBad";
                                lblSheetsAffected.Visible = true;
                                lblSheetsAffected.Text = "Sheets\r\nAffected";
                                flowLayoutPanel2.HorizontalScroll.Visible = false;
                                Label lbl1 = new Label();
                                this.flowLayoutPanel2.Controls.Add(lbl1);
                                lbl1.Height = 0;
                                lbl1.Width = 430;                     
                                Label lbl2 = new Label();
                                this.flowLayoutPanel2.Controls.Add(lbl2);
                                lbl2.Height = 40;
                                lbl2.Width = 120;
                                lbl2.BackColor = Color.Silver;
                                lbl2.Font = new Font("Microsoft Sans Serif", 20);
                                lbl2.TextAlign = ContentAlignment.MiddleLeft;
                                lbl2.ForeColor = Color.Black;
                                lbl2.Margin = new Padding(0, 0, 0, 0);
                                lbl2.Left = 40;
                                lbl2.Text = this.dataGridView1.Rows[tb2].Cells[0].Value.ToString();
                                Label lbl3 = new Label();
                                this.flowLayoutPanel2.Controls.Add(lbl3);
                                lbl3.Height = 40;
                                lbl3.Width = 108;
                                lbl3.BackColor = Color.Silver;
                                lbl3.Font = new Font("Microsoft Sans Serif", 20);
                                lbl3.TextAlign = ContentAlignment.MiddleCenter;
                                lbl3.ForeColor = Color.Black;
                                lbl3.Margin = new Padding(0, 0, 0, 0);
                                lbl3.Left = 40;
                                lbl3.Text = this.dataGridView1.Rows[tb2].Cells[12].Value.ToString();
                                TextBox textBox1 = new TextBox();
                                this.flowLayoutPanel2.Controls.Add(textBox1);
                                textBox1.Height = 40;
                                textBox1.AutoSize = false;
                                textBox1.Width = 70;
                                textBox1.Multiline = false;
                                textBox1.Font = new Font(textBox1.Font.FontFamily, 20);
                                textBox1.TextAlign = HorizontalAlignment.Center;
                                textBox1.Margin = new Padding(0, 0, 0, 0);
                                //textBox1.Tag = tb1;
                                textBox1.TextChanged += new System.EventHandler(this.notGangedNumberUpBad);
                                TextBox textBox2 = new TextBox();
                                this.flowLayoutPanel2.Controls.Add(textBox2);
                                textBox2.Height = 40;
                                textBox2.AutoSize = false;
                                textBox2.Width = 85;
                                textBox2.Multiline = false;
                                textBox2.Font = new Font(textBox2.Font.FontFamily, 20);
                                textBox2.TextAlign = HorizontalAlignment.Center;
                                textBox2.Margin = new Padding(0, 0, 0, 0);
                                //textBox2.Tag = tb2;
                                textBox2.TextChanged += new System.EventHandler(notgangedSheetsAffected);
                                Button btn1 = new Button();
                                flowLayoutPanel2.Controls.Add(btn1);
                                btn1.Height = 40;
                                btn1.Width = 90;
                                btn1.BackColor = Color.SteelBlue;
                                btn1.ForeColor = Color.White;
                                btn1.Font = new Font(textBox1.Font.FontFamily, 9);
                                btn1.Text = "Whole Pallet";
                                btn1.Margin = new Padding(0, 0, 0, 0);
                                btn1.Tag = tb2;
                                btn1.Click += new System.EventHandler(notGangedWholePallet);
                                //numberBadList.Insert(tb2, "0");
                                //sheetsAffectedList.Insert(tb2, "0");
                                //wholePalletList.Insert(tb2, 0);
                            }
                        }
                        badSectionLbls = true;
                    }
                }
                //if JobGanged = 1 (PALLET_GANG_CLASSIC Table)
                else if (Convert.ToInt32(dataGridView1.Rows[0].Cells[14].Value) == 1)
                {
                    if (numberUp > 1)
                    {
                        if (!badSectionLbls)
                        {
                            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                            {
                                lblStockCode.Text = "Stock Code/\r\nJob Number";
                                lblStockCode.Visible = true;
                                lblNumberUp.Visible = true;
                                lblNumberUp.Text = "Number\r\nUp";
                                lblNumberUpBadQty.Visible = true;
                                lblNumberUpBadQty.Text = "Quantity\r\nBad";
                                lblSheetsAffected.Visible = true;
                                lblSheetsAffected.Text = "Sheets\r\nAffected";
                                flowLayoutPanel2.HorizontalScroll.Visible = false;
                                flowLayoutPanel2.VerticalScroll.Visible = false;
                                Label lbl1 = new Label();
                                this.flowLayoutPanel2.Controls.Add(lbl1);
                                lbl1.Height = 35;
                                lbl1.Width = 430;
                                lbl1.BackColor = Color.Gray;
                                lbl1.Font = new Font("Microsoft Sans Serif", 12);
                                lbl1.TextAlign = ContentAlignment.MiddleLeft;
                                lbl1.ForeColor = Color.White;
                                lbl1.Left = 40;
                                lbl1.Text = this.dataGridView4.Rows[i].Cells[3].Value.ToString();
                                Label lbl2 = new Label();
                                this.flowLayoutPanel2.Controls.Add(lbl2);
                                lbl2.Height = 40;
                                lbl2.Width = 120;
                                lbl2.BackColor = Color.Silver;
                                lbl2.Font = new Font("Microsoft Sans Serif", 20);
                                lbl2.TextAlign = ContentAlignment.MiddleLeft;
                                lbl2.ForeColor = Color.Black;
                                lbl2.Margin = new Padding(0, 0, 0, 0);
                                lbl2.Left = 40;
                                lbl2.Text = this.dataGridView4.Rows[i].Cells[1].Value.ToString();
                                Label lbl3 = new Label();
                                this.flowLayoutPanel2.Controls.Add(lbl3);
                                lbl3.Height = 40;
                                lbl3.Width = 108;
                                lbl3.BackColor = Color.Silver;
                                lbl3.Font = new Font("Microsoft Sans Serif", 20);
                                lbl3.TextAlign = ContentAlignment.MiddleCenter;
                                lbl3.ForeColor = Color.Black;
                                lbl3.Margin = new Padding(0, 0, 0, 0);
                                lbl3.Left = 40;
                                lbl3.Text = this.dataGridView4.Rows[i].Cells[11].Value.ToString();
                                TextBox textBox1 = new TextBox();
                                this.flowLayoutPanel2.Controls.Add(textBox1);
                                textBox1.Height = 40;
                                textBox1.AutoSize = false;
                                textBox1.Width = 70;
                                textBox1.Multiline = false;
                                textBox1.Font = new Font(textBox1.Font.FontFamily, 20);
                                textBox1.TextAlign = HorizontalAlignment.Center;
                                textBox1.Margin = new Padding(0, 0, 0, 0);
                                textBox1.Tag = i;
                                textBox1.TextChanged += new System.EventHandler(this.gangClassicNumberUpBad);
                                TextBox textBox2 = new TextBox();
                                this.flowLayoutPanel2.Controls.Add(textBox2);
                                textBox2.Height = 40;
                                textBox2.AutoSize = false;
                                textBox2.Width = 85;
                                textBox2.Multiline = false;
                                textBox2.Font = new Font(textBox2.Font.FontFamily, 20);
                                textBox2.TextAlign = HorizontalAlignment.Center;
                                textBox2.Margin = new Padding(0, 0, 0, 0);
                                textBox2.Tag = i;
                                textBox2.TextChanged += new System.EventHandler(gangClassicSheetsAffected);
                                Button btn1 = new Button();
                                flowLayoutPanel2.Controls.Add(btn1);
                                btn1.Height = 40;
                                btn1.Width = 73;
                                btn1.BackColor = Color.SteelBlue;
                                btn1.ForeColor = Color.White;
                                btn1.Font = new Font(textBox1.Font.FontFamily, 9);
                                btn1.Text = "Whole Pallet";
                                btn1.Margin = new Padding(0, 0, 0, 0);
                                btn1.Tag = i;
                                btn1.Click += new System.EventHandler(this.gangClassicWholePallet);
                                numberBadList.Insert(i, "0");
                                sheetsAffectedList.Insert(i, "0");
                                wholePalletList.Insert(i, 0);
                            }
                        }
                        badSectionLbls = true;
                    }
                }







                //if JobGanged = 3 (PALLET_GANG_CLASSIC Table)
                else if (Convert.ToInt32(dataGridView1.Rows[0].Cells[14].Value) == 3)
                {
                    if (numberUp > 1)
                    {
                        if (!badSectionLbls)
                        {
                            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                            {
                                lblStockCode.Text = "Stock Code/\r\nJob Number";
                                lblStockCode.Visible = true;
                                lblNumberUp.Visible = true;
                                lblNumberUp.Text = "Number\r\nUp";
                                lblNumberUpBadQty.Visible = true;
                                lblNumberUpBadQty.Text = "Quantity\r\nBad";
                                lblSheetsAffected.Visible = true;
                                lblSheetsAffected.Text = "Sheets\r\nAffected";
                                flowLayoutPanel2.HorizontalScroll.Visible = false;
                                flowLayoutPanel2.VerticalScroll.Visible = false;
                                Label lbl1 = new Label();
                                this.flowLayoutPanel2.Controls.Add(lbl1);
                                lbl1.Height = 35;
                                lbl1.Width = 430;
                                lbl1.BackColor = Color.Gray;
                                lbl1.Font = new Font("Microsoft Sans Serif", 12);
                                lbl1.TextAlign = ContentAlignment.MiddleLeft;
                                lbl1.ForeColor = Color.White;
                                lbl1.Left = 40;
                                lbl1.Text = this.dataGridView4.Rows[i].Cells[3].Value.ToString();
                                Label lbl2 = new Label();
                                this.flowLayoutPanel2.Controls.Add(lbl2);
                                lbl2.Height = 40;
                                lbl2.Width = 120;
                                lbl2.BackColor = Color.Silver;
                                lbl2.Font = new Font("Microsoft Sans Serif", 20);
                                lbl2.TextAlign = ContentAlignment.MiddleLeft;
                                lbl2.ForeColor = Color.Black;
                                lbl2.Margin = new Padding(0, 0, 0, 0);
                                lbl2.Left = 40;
                                lbl2.Text = this.dataGridView4.Rows[i].Cells[1].Value.ToString();
                                Label lbl3 = new Label();
                                this.flowLayoutPanel2.Controls.Add(lbl3);
                                lbl3.Height = 40;
                                lbl3.Width = 108;
                                lbl3.BackColor = Color.Silver;
                                lbl3.Font = new Font("Microsoft Sans Serif", 20);
                                lbl3.TextAlign = ContentAlignment.MiddleCenter;
                                lbl3.ForeColor = Color.Black;
                                lbl3.Margin = new Padding(0, 0, 0, 0);
                                lbl3.Left = 40;
                                lbl3.Text = this.dataGridView4.Rows[i].Cells[11].Value.ToString();
                                TextBox textBox1 = new TextBox();
                                this.flowLayoutPanel2.Controls.Add(textBox1);
                                textBox1.Height = 40;
                                textBox1.AutoSize = false;
                                textBox1.Width = 70;
                                textBox1.Multiline = false;
                                textBox1.Font = new Font(textBox1.Font.FontFamily, 20);
                                textBox1.TextAlign = HorizontalAlignment.Center;
                                textBox1.Margin = new Padding(0, 0, 0, 0);
                                textBox1.Tag = i;
                                textBox1.TextChanged += new System.EventHandler(this.gangClassicNumberUpBad);
                                TextBox textBox2 = new TextBox();
                                this.flowLayoutPanel2.Controls.Add(textBox2);
                                textBox2.Height = 40;
                                textBox2.AutoSize = false;
                                textBox2.Width = 85;
                                textBox2.Multiline = false;
                                textBox2.Font = new Font(textBox2.Font.FontFamily, 20);
                                textBox2.TextAlign = HorizontalAlignment.Center;
                                textBox2.Margin = new Padding(0, 0, 0, 0);
                                textBox2.Tag = i;
                                textBox2.TextChanged += new System.EventHandler(gangClassicSheetsAffected);
                                Button btn1 = new Button();
                                flowLayoutPanel2.Controls.Add(btn1);
                                btn1.Height = 40;
                                btn1.Width = 73;
                                btn1.BackColor = Color.SteelBlue;
                                btn1.ForeColor = Color.White;
                                btn1.Font = new Font(textBox1.Font.FontFamily, 9);
                                btn1.Text = "Whole Pallet";
                                btn1.Margin = new Padding(0, 0, 0, 0);
                                btn1.Tag = i;
                                btn1.Click += new System.EventHandler(this.gangClassicWholePallet);
                                numberBadList.Insert(i, "0");
                                sheetsAffectedList.Insert(i, "0");
                                wholePalletList.Insert(i, 0);
                            }
                        }
                        badSectionLbls = true;
                    }
                }









            }
            
            if (dataGridView4.RowCount != 0)
            {
                lbl7.Text = dataGridView4.Rows[0].Cells[8].Value.ToString();
            }

            if (dataGridView3.RowCount != 0)
            {
                lbl7.Text = dataGridView3.Rows[0].Cells[8].Value.ToString();
            }
        }

        private void btnScrollDown_Click(object sender, EventArgs e)
        {
            flowLayoutPanel2.AutoScrollPosition =
            new Point(0, flowLayoutPanel2.VerticalScroll.Value +
                 flowLayoutPanel2.VerticalScroll.SmallChange * 7);
        }

        private void btnScrollUp_Click(object sender, EventArgs e)
        {
            flowLayoutPanel2.AutoScrollPosition =
            new Point(0, flowLayoutPanel2.VerticalScroll.Value +
            flowLayoutPanel2.VerticalScroll.SmallChange * -7);
        }

        private void notGangedNumberUpBad(Object sender, EventArgs e)
        {
            // variable to keep the textbox value
            badQty = ((TextBox)sender).Text;
            if (badQty == "")
            {
                badQty = "0";
            }
            notGanged();
        }

        private void notgangedSheetsAffected(Object sender, EventArgs e)
        {
            gangWholePalletButtonPressed = 0;
            sheetsAffected = ((TextBox)sender).Text;

            if (sheetsAffected == "")
            {
                sheetsAffected = "0";
            }
            notGanged();
        }

        private void notGangedWholePallet(Object sender, EventArgs e)
        {
            gangWholePalletButtonPressed = 1;
            notGangedWholePalletValue = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
            notGanged();
        }

        private void gangClassicNumberUpBad(Object sender, EventArgs e)
        {
            badQty = ((TextBox)sender).Text;
            gangRow = (int)((TextBox)sender).Tag;

            if (badQty == "")
            {
                badQty = "0";
            }

            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                if (gangRow == i)
                {
                    numberBadList[i] = badQty.ToString();
                }            
            //queryGangpro();
            queryGangClassic();
        }

        private void gangClassicSheetsAffected(Object sender, EventArgs e)
        {
            gangWholePalletButtonPressed = 0;
            sheetsAffected = ((TextBox)sender).Text;
            gangRow = (int)((TextBox)sender).Tag;

            if (sheetsAffected == "")
            {
                sheetsAffected = "0";
            }

            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                if (gangRow == i)
                {
                    sheetsAffectedList[i] = sheetsAffected.ToString();
                }
            //queryGangpro();
            queryGangClassic();
        }

        private void gangClassicWholePallet(Object sender, EventArgs e)
        {
            gangRow = (int)((Button)sender).Tag;
            gangWholePalletButtonPressed = 1;
            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                if (gangRow == i)
                {
                    wholePalletList[i] = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
                }
            //queryGangpro();
            queryGangClassic();
        }






        private void gangProNumberUpBad(Object sender, EventArgs e)
        {
            badQty = ((TextBox)sender).Text;
            gangRow = (int)((TextBox)sender).Tag;

            if (badQty == "")
            {
                badQty = "0";
            }

            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                if (gangRow == i)
                {
                    numberBadList[i] = badQty.ToString();
                }
            queryGangpro();
        }

        private void gangProSheetsAffected(Object sender, EventArgs e)
        {
            gangWholePalletButtonPressed = 0;
            sheetsAffected = ((TextBox)sender).Text;
            gangRow = (int)((TextBox)sender).Tag;

            if (sheetsAffected == "")
            {
                sheetsAffected = "0";
            }

            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                if (gangRow == i)
                {
                    sheetsAffectedList[i] = sheetsAffected.ToString();
                }
            queryGangpro();
        }

        private void gangProWholePallet(Object sender, EventArgs e)
        {
            gangRow = (int)((Button)sender).Tag;
            gangWholePalletButtonPressed = 1;
            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                if (gangRow == i)
                {
                    wholePalletList[i] = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
                }
            queryGangpro();
        }






        private void btnBadSectionOK_Click(object sender, EventArgs e)
        {
            if(tbxSheetsAffectedBadSection.Visible == true)
            {
                if (tbxSheetsAffectedBadSection.Text == "")
                {
                    MessageBox.Show("Please enter a value in Sheets Affected box");
                }
            }
            pnlPalletCard8.BringToFront();
            index = 13;   
        }

        private void btnExtraInformationPalletCard_Click(object sender, EventArgs e)
        {
            pnlPalletCard7.BringToFront();
            index = 14;
        }

        private void btnFinishPalletContinue_Click(object sender, EventArgs e)
        {
            // Disable the Section button
            disableSectionButtons.Add(Convert.ToString(dataGridView1.Rows[0].Cells[19].Value));
            removeFlowLayoutBtns();
            sigBtns = false;
            pnlSignature.BringToFront();
            btnCancel.Visible = false;
        }

        private void btnPalletFinished_Click(object sender, EventArgs e)
        {
            pnlPalletCard8.BringToFront();

            if(dataGridView2.Rows.Count != 0)
            {
                btnIsSectionFinishedYes_Click(btnIsSectionFinishedYes, EventArgs.Empty);
                index = 15;
            }
            else
            { 
                if (dataGridView1.Rows[0].Cells[15].Value.ToString() == "")
                {
                    lblIsSectionFinished.Text = "Is " + dataGridView1.Rows[0].Cells[11].Value.ToString() + "\r\n" + "Section " + dataGridView1.Rows[0].Cells[19].Value.ToString() + " finished ?";
                }
                else
                {
                    lblIsSectionFinished.Text = "Is " + dataGridView1.Rows[0].Cells[15].Value.ToString() + "\r\n" + "Section " + dataGridView1.Rows[0].Cells[19].Value.ToString() + " finished ?";
                }

                // WorkingSize
                lbl6.Text = dataGridView1.Rows[0].Cells[13].Value.ToString();
                // QtyRequired
                lbl7.Text = dataGridView1.Rows[0].Cells[26].Value.ToString();

                index = 15;
            }
        }

        private void btnCancelPrintMore_Click(object sender, EventArgs e)
        {
            Cancel();
            pnlHome0.BringToFront();
            string ConnectionString = Convert.ToString("Dsn=PalletCard;uid=PalletCardAdmin");
            string CommandText = "DELETE FROM Log where JobNo = '" + lblJobNo.Text + "' and PalletNumber = PalletNumber";
            OdbcConnection myConnection = new OdbcConnection(ConnectionString);
            OdbcCommand myCommand = new OdbcCommand(CommandText, myConnection);
            OdbcDataAdapter myAdapter = new OdbcDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet palletCardData = new DataSet();
            try
            {
                myConnection.Open();
                myAdapter.Fill(palletCardData);
            }
            catch (Exception ex)
            {
                throw (ex);
            }

            finally
            {
                myConnection.Close();
            }
            using (DataTable palletCardLog = new DataTable())
            {
                myAdapter.Fill(palletCardLog);
                dataGridView2.DataSource = palletCardLog;
            }
            index = 0;
        }

        private void btnBackupRequired_Click(object sender, EventArgs e)
        {
            //SAVE TO DATABASE
            produced = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
            string sqlFormattedDate = CurrentDate.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
            string Query = "insert into Log (Routine, JobNo, PaperSectionNo, PalletNumber, Produced, QtyRequired, ResourceID, Description, WorkingSize, SheetQty, Comment, Unfinished, Timestamp1) values('" + this.lbl1.Text + "','" + this.dataGridView1.Rows[0].Cells[0].Value + "','" + this.dataGridView1.Rows[0].Cells[19].Value + "', '1', '" + produced + "', '" + this.dataGridView1.Rows[0].Cells[26].Value + "','" + resourceID + "','" + this.lbl2.Text + "','" + this.dataGridView1.Rows[0].Cells[13].Value + "','" + this.lbl5.Text + "','" + this.tbxExtraInfoComment.Text + "','1','" + CurrentDate + "');";
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

            string ConnectionString = Convert.ToString("Dsn=PalletCard;uid=PalletCardAdmin");
            string CommandText = "SELECT * FROM Log where JobNo = '" + lblJobNo.Text + "'";
            OdbcConnection myConnection = new OdbcConnection(ConnectionString);
            OdbcCommand myCommand = new OdbcCommand(CommandText, myConnection);
            OdbcDataAdapter myAdapter = new OdbcDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet palletCardData = new DataSet();
            try
            {
                myConnection.Open();
                myAdapter.Fill(palletCardData);
            }
            catch (Exception ex)
            {
                throw (ex);
            }

            finally
            {
                myConnection.Close();
            }
            using (DataTable palletCardLog = new DataTable())
            {
                myAdapter.Fill(palletCardLog);
                dataGridView2.DataSource = palletCardLog;
            }

            this.dataGridView2.Sort(this.dataGridView2.Columns["PalletNumber"], ListSortDirection.Descending);
            string barCode = Convert.ToString(((int)dataGridView2.Rows[0].Cells[5].Value));
            Bitmap bitMap = new Bitmap(barCode.Length * 40, 80);
            using (Graphics graphics = Graphics.FromImage(bitMap))
            {
                Font oFont = new Font("IDAutomationHC39M", 16);
                PointF point = new PointF(2f, 2f);
                SolidBrush blackBrush = new SolidBrush(Color.Black);
                SolidBrush whiteBrush = new SolidBrush(Color.White);
                graphics.FillRectangle(whiteBrush, 0, 0, bitMap.Width, bitMap.Height);
                graphics.DrawString("*" + barCode + "*", oFont, blackBrush, point);
            }
            using (MemoryStream ms = new MemoryStream())
            {
                bitMap.Save(ms, ImageFormat.Png);
                pictureBox1.Image = bitMap;
                pictureBox1.Height = bitMap.Height;
                pictureBox1.Width = bitMap.Width;
            }

            pnlPalletCardPrint.BringToFront();
            lblPC_IncompletePallet.Text = "INCOMPLETE";
            lblPC_IncompletePallet.Visible = true;
            lblPC_JobNo.Text = lblJobNo.Text;
            lblPC_JobNo.Visible = true;
            lblPC_Customer.Text = dataGridView1.Rows[0].Cells[22].Value as string;
            lblPC_Customer.Visible = true;
            lblPC_Customer.MaximumSize = new Size(450, 220);
            lblPC_Customer.AutoSize = true;
            lblPC_SheetQty.Text = lbl5.Text;
            lblPC_SheetQty.Visible = true;
            lblPC_Sig.Text = "Sheet " + dataGridView1.Rows[0].Cells[19].Value as string;
            lblPC_Sig.Visible = true;
            lblPC_PalletNumber.Text = "Pallet 1";
            lblPC_PalletNumber.Visible = true;
            lblPC_Press.Text = "Press - " + lblPress.Text;
            lblPC_Press.Visible = true;
            lblPC_Date.Text = "Date - " + DateTime.Now.ToString("d/M/yyyy");
            lblPC_Date.Visible = true;
            lblPC_Note.Text = tbxExtraInfoComment.Text + " - " + tbxTextBoxBadSection.Text;
            lblPC_Note.Visible = true;
            btnBack.Visible = false;
            btnCancel.Visible = false;
            index = 17;
        }

        private void btnVarnishRequired_Click(object sender, EventArgs e)
        {
            //SAVE TO DATABASE
            produced = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
            string sqlFormattedDate = CurrentDate.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
            string Query = "insert into Log (Routine, JobNo, PaperSectionNo, PalletNumber, Produced, QtyRequired, ResourceID, Description, WorkingSize, SheetQty, Comment, Unfinished, Timestamp1) values('" + this.lbl1.Text + "','" + this.dataGridView1.Rows[0].Cells[0].Value + "','" + this.dataGridView1.Rows[0].Cells[19].Value + "', '1', '" + produced + "', '" + this.dataGridView1.Rows[0].Cells[26].Value + "', '" + resourceID + "','" + this.lbl2.Text + "','" + this.dataGridView1.Rows[0].Cells[13].Value + "','" + this.lbl5.Text + "','" + this.tbxExtraInfoComment.Text + "','1','" + CurrentDate + "');";
            SqlConnection conDatabase = new SqlConnection(constring);
            SqlCommand cmdDatabase = new SqlCommand(Query, conDatabase);
            SqlDataReader myReader;
            try
            {
                conDatabase.Open();
                myReader = cmdDatabase.ExecuteReader();
                conDatabase.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            string ConnectionString = Convert.ToString("Dsn=PalletCard;uid=PalletCardAdmin");
            string CommandText = "SELECT * FROM Log where JobNo = '" + lblJobNo.Text + "'";
            OdbcConnection myConnection = new OdbcConnection(ConnectionString);
            OdbcCommand myCommand = new OdbcCommand(CommandText, myConnection);
            OdbcDataAdapter myAdapter = new OdbcDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet palletCardData = new DataSet();
            try
            {
                myConnection.Open();
                myAdapter.Fill(palletCardData);
            }
            catch (Exception ex)
            {
                throw (ex);
            }

            finally
            {
                myConnection.Close();
            }
            using (DataTable palletCardLog = new DataTable())
            {
                myAdapter.Fill(palletCardLog);
                dataGridView2.DataSource = palletCardLog;
            }

            this.dataGridView2.Sort(this.dataGridView2.Columns["PalletNumber"], ListSortDirection.Descending);
            string barCode = Convert.ToString(((int)dataGridView2.Rows[0].Cells[5].Value));
            Bitmap bitMap = new Bitmap(barCode.Length * 40, 80);
            using (Graphics graphics = Graphics.FromImage(bitMap))
            {
                Font oFont = new Font("IDAutomationHC39M", 16);
                PointF point = new PointF(2f, 2f);
                SolidBrush blackBrush = new SolidBrush(Color.Black);
                SolidBrush whiteBrush = new SolidBrush(Color.White);
                graphics.FillRectangle(whiteBrush, 0, 0, bitMap.Width, bitMap.Height);
                graphics.DrawString("*" + barCode + "*", oFont, blackBrush, point);
            }
            using (MemoryStream ms = new MemoryStream())
            {
                bitMap.Save(ms, ImageFormat.Png);
                pictureBox1.Image = bitMap;
                pictureBox1.Height = bitMap.Height;
                pictureBox1.Width = bitMap.Width;
            }

            pnlPalletCardPrint.BringToFront();
            lblPC_IncompletePallet.Text = "INCOMPLETE";
            lblPC_IncompletePallet.Visible = true;
            lblPC_JobNo.Text = lblJobNo.Text;
            lblPC_JobNo.Visible = true;
            lblPC_Customer.Text = dataGridView1.Rows[0].Cells[22].Value as string;
            lblPC_Customer.Visible = true;
            lblPC_Customer.MaximumSize = new Size(450, 220);
            lblPC_Customer.AutoSize = true;
            lblPC_SheetQty.Text = lbl5.Text;
            lblPC_SheetQty.Visible = true;
            lblPC_Sig.Text = "Sheet " + dataGridView1.Rows[0].Cells[19].Value as string;
            lblPC_Sig.Visible = true;
            lblPC_PalletNumber.Text = "Pallet 1";
            lblPC_PalletNumber.Visible = true;
            lblPC_Press.Text = "Press - " + lblPress.Text;
            lblPC_Press.Visible = true;
            lblPC_Date.Text = "Date - " + DateTime.Now.ToString("d/M/yyyy");
            lblPC_Date.Visible = true;
            lblPC_Note.Text = tbxExtraInfoComment.Text + " - " + tbxTextBoxBadSection.Text;
            lblPC_Note.Visible = true;
            btnBack.Visible = false;
            btnCancel.Visible = false;
            index = 17;
        }


        private void reQueryDataGridView2()
        {
            string ConnectionString = Convert.ToString("Dsn=PalletCard;uid=PalletCardAdmin");
            string CommandText = "SELECT * FROM Log where JobNo = '" + lblJobNo.Text + "'";
            OdbcConnection myConnection = new OdbcConnection(ConnectionString);
            OdbcCommand myCommand = new OdbcCommand(CommandText, myConnection);
            OdbcDataAdapter myAdapter = new OdbcDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet palletCardData = new DataSet();
            try
            {
                myConnection.Open();
                myAdapter.Fill(palletCardData);
            }
            catch (Exception ex)
            {
                throw (ex);
            }

            finally
            {
                //myConnection.Close();
            }
            using (DataTable palletCardLog = new DataTable())
            {
                myAdapter.Fill(palletCardLog);
                dataGridView2.DataSource = palletCardLog;
            }
        }


        private void btnIsSectionFinishedYes_Click(object sender, EventArgs e)
        {
            string ConnectionString = Convert.ToString("Dsn=PalletCard;uid=PalletCardAdmin");
            string CommandText = "SELECT * FROM Log where JobNo = '" + lblJobNo.Text + "'";
            OdbcConnection myConnection = new OdbcConnection(ConnectionString);
            OdbcCommand myCommand = new OdbcCommand(CommandText, myConnection);
            OdbcDataAdapter myAdapter = new OdbcDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet palletCardData = new DataSet();
            try
            {
                myConnection.Open();
                myAdapter.Fill(palletCardData);
            }
            catch (Exception ex)
            {
                throw (ex);
            }

            finally
            {
                //myConnection.Close();
            }
            using (DataTable palletCardLog = new DataTable())
            {
                myAdapter.Fill(palletCardLog);
                dataGridView2.DataSource = palletCardLog;
            }
            // If This job Number has not yet been recorded in the database
            if (dataGridView2.Rows.Count == 0)
            {
                PalletNumber = 1;
                dataGridView2.AllowUserToAddRows = true;
            }
            // Otherwise check if any previous Pallet Numbers("Pallet Card" Routine entries) and record as the next sequential Pallet Number
            else
            {
                try
                {
                    // (There could be entries for this job Number but for Return or reject Paper)
                    ((DataTable)dataGridView2.DataSource).DefaultView.RowFilter = "Routine like 'Pallet Card'";
                }
                catch (Exception) { }

                // if PalletNumber field is empty
                if (dataGridView2.Rows[0].Cells[4].Value as string == "")
                {
                    PalletNumber = 1;
                }
                else
                {
                    this.dataGridView2.Sort(this.dataGridView2.Columns["PalletNumber"], ListSortDirection.Descending);
                    PalletNumber = (int)dataGridView2.Rows[0].Cells[4].Value + 1;
                }
            }

            //SAVE TO DATABASE
            CurrentDate = DateTime.Now;
            dataGridView2.Refresh();
            produced = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", "")) - sheetsAffectedBadSection;
            PaperSectionNo = Convert.ToInt32(Regex.Replace(lbl3.Text, "[^0-9.]", ""));

            string sqlFormattedDate = CurrentDate.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
            string Query = "insert into Log (Routine, JobNo, PalletNumber, Unfinished, PaperSectionNo, QtyRequired, ResourceID, WorkingSize, Description, SheetQty, Comment, Timestamp1, LastPallet, Produced, Expr1, SectionName) values('" + this.lbl1.Text + "','" + lblJobNo.Text + "','" + PalletNumber + "', '0','" + PaperSectionNo + "', '" + lbl7.Text + "','" + resourceID + "','" + lbl6.Text + "','" + lbl2.Text + "','" + lbl5.Text + "','" + tbxExtraInfoComment.Text + "','" + CurrentDate + "','" + "1" + "','" + produced + "','" + lbl2.Text + "','" + lbl2.Text + "');";
            SqlConnection conDatabase = new SqlConnection(constring);
            SqlCommand cmdDatabase = new SqlCommand(Query, conDatabase);
            SqlDataReader myReader;
            try
            {
                conDatabase.Open();
                myReader = cmdDatabase.ExecuteReader();
                conDatabase.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            reQueryDataGridView2();

            // Requery Log table to refresh datagridview2 with new line
            //OdbcConnection myConnection1 = new OdbcConnection(ConnectionString);
            //OdbcCommand myCommand1 = new OdbcCommand(CommandText, myConnection1);
            //OdbcDataAdapter myAdapter1 = new OdbcDataAdapter();
            //myAdapter1.SelectCommand = myCommand;
            //DataSet palletCardData1 = new DataSet();
            //try
            //{
            //    myConnection.Open();
            //    myAdapter.Fill(palletCardData1);
            //}
            //catch (Exception ex)
            //{
            //    throw (ex);
            //}

            //finally
            //{
            //    myConnection.Close();
            //}
            //using (DataTable palletCardLog1 = new DataTable())
            //{
            //    myAdapter.Fill(palletCardLog1);
            //    dataGridView2.DataSource = palletCardLog1;
            //}

            dataGridView2.Refresh();
            //dataGridView2.AllowUserToAddRows = false;

            // Get the quantities produced from the previous pallet cards
            int sumProduced = 0;
            // ignore if the value has been retrieved by scan
            this.dataGridView2.Sort(this.dataGridView2.Columns["AutoNum"], ListSortDirection.Descending);
            if (Convert.ToInt32(dataGridView2.Rows[0].Cells[6].Value) != 0)
            {
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    dataGridView2.AllowUserToAddRows = false;
                    // Sum up only PaperSectionNo's associated with this section not all PaperSectionNo's
                    if (Convert.ToInt32(dataGridView2.Rows[i].Cells[8].Value) == PaperSectionNo)
                    {
                        sumProduced += Convert.ToInt32(dataGridView2.Rows[i].Cells[9].Value);
                    }
                }
            }
            required = Convert.ToInt32(lbl7.Text);
            produced = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", "")) - sheetsAffectedBadSection + sumProduced;
            shortBy = required - produced;
            overBy = produced - required;

            // OVER PRODUCED/UNDER PRODUCED LOGIC
            if ((required * 105 / 100) < required + 50)
            {
                oversCalc = required + 50;
            }
            else
            {
                oversCalc = (required * 105 / 100);
            }

            if (!backupRequired || !varnishRequired)
            {
                if (produced <= required)
                {
                    pnlPalletCard6.BringToFront();
                    btnBack.Visible = false;
                    lblPalletDidNotMakeQty.Text = "Job " + lblJobNo.Text + " Sheet " + dataGridView1.Rows[0].Cells[19].Value.ToString() + " has " + shortBy + " insufficient sheets";
                    //lbl7.Text = "Pallet Short";
                    lblFinishedPallets.Visible = false;

                    // Check if 1 finished pallet for each section - if not provide a warning message listing the remaing pallets to finish
                    for (int i = 0; i < this.dataGridView2.Rows.Count; i++)
                    {
                        if (Convert.ToInt32(dataGridView2.Rows[i].Cells[7].Value) == 1)
                        {
                            completedSections.Add(dataGridView2.Rows[i].Cells[8].Value.ToString());
                        }
                    }

                    List<string> sectionsNoLastFlag = new List<string>();
                    foreach (string s in allSections)
                    {
                        if (!completedSections.Contains(s))
                            sectionsNoLastFlag.Add(s);
                    }

                    lblFinishedPallets.Visible = true;
                    lblFinishedPallets.Text = "";
                    foreach (string s in sectionsNoLastFlag)
                    {
                        lblFinishedPallets.Text += "Section " + s + " is not complete" + "\r\n";
                    }
                    if (lblFinishedPallets.Text.Length > 0)
                    {
                        lblWarning.Visible = true;
                    }

                    // Send email notification
                    MailMessage mail = new MailMessage("PalletShort@colorman.ie", "declan.enright@colorman.ie", "Pallet Short", "Job Number " + lblJobNo.Text + " - Section " + dataGridView2.Rows[0].Cells[8].Value.ToString() + "- has " + shortBy + " insufficient sheets");
                    SmtpClient client = new SmtpClient("ex0101.ColorMan.local");
                    client.Port = 25;
                    client.EnableSsl = false;
                    client.Send(mail);
                }

                else if (produced > oversCalc)
                {
                    pnlPalletCard9.BringToFront();
                    btnBack.Visible = false;
                    lblPalletOverBySheets.Text = lblJobNo.Text + " is over by " + overBy;
                    //lbl7.Text = "Pallet Over";
                    // Disable the Section button
                    disableSectionButtons.Add(Convert.ToString(dataGridView1.Rows[0].Cells[19].Value));
                    removeFlowLayoutBtns();
                    sigBtns = false;

                    // Check if 1 finished pallet for each section - if not provide a warning message listing the remaing pallets to finish
                    for (int i = 0; i < this.dataGridView2.Rows.Count; i++)
                    {
                        if (Convert.ToInt32(dataGridView2.Rows[i].Cells[7].Value) == 1)
                        {
                            completedSections.Add(dataGridView2.Rows[i].Cells[8].Value.ToString());
                        }
                    }

                    List<string> sectionsNoLastFlag = new List<string>();
                    foreach (string s in allSections)
                    {
                        if (!completedSections.Contains(s))
                            sectionsNoLastFlag.Add(s);
                    }

                    lblFinishedPalletsOver.Visible = true;
                    lblFinishedPalletsOver.Text = "";
                    foreach (string s in sectionsNoLastFlag)
                    {
                        lblFinishedPalletsOver.Text += "Section " + s + " is not complete" + "\r\n";
                    }
                    if (lblFinishedPalletsOver.Text.Length > 0)
                    {
                        lblWarningOver.Visible = true;
                    }

                    // Send email notification
                    MailMessage mail = new MailMessage("PalletOver@colorman.ie", "declan.enright@colorman.ie", "Pallet Over", "Job Number " + lblJobNo.Text + " - Section " + dataGridView2.Rows[0].Cells[8].Value.ToString() + " - is over by " + overBy);
                    SmtpClient client = new SmtpClient("ex0101.ColorMan.local");
                    client.Port = 25;
                    client.EnableSsl = false;
                    client.Send(mail);
                    index = 16;
                }
                else if (produced <= oversCalc)
                {
                    pnlSignature.BringToFront();
                    btnBack.Visible = false;
                    // Disable the Section button
                    disableSectionButtons.Add(Convert.ToString(dataGridView1.Rows[0].Cells[19].Value));
                    removeFlowLayoutBtns();
                    sigBtns = false;

                    // Check if 1 finished pallet for each section - if not provide a warning message listing the remaing pallets to finish
                    for (int i = 0; i < this.dataGridView2.Rows.Count; i++)
                    {
                        if (Convert.ToInt32(dataGridView2.Rows[i].Cells[7].Value) == 1)
                        {
                            completedSections.Add(dataGridView2.Rows[i].Cells[8].Value.ToString());
                        }
                    }

                    List<string> sectionsNoLastFlag = new List<string>();
                    foreach (string s in allSections)
                    {
                        if (!completedSections.Contains(s))
                            sectionsNoLastFlag.Add(s);
                    }

                    //lblFinishedPallets.Visible = true;
                    //lblFinishedPallets.Text = "";
                    //foreach (string s in sectionsNoLastFlag)
                    //{
                    //    lblFinishedPallets.Text += "Section " + s + " is not complete" + "\r\n";
                    //}
                    //if (lblFinishedPallets.Text.Length > 0)
                    //{
                    //    lblWarning.Visible = true;
                    //}
                }
            }
            autoNum = dataGridView2.Rows[0].Cells[0].Value.ToString();
            index = 16;
        }

        private void btnIsSectionFinishedNo_Click(object sender, EventArgs e)
        {
            string ConnectionString = Convert.ToString("Dsn=PalletCard;uid=PalletCardAdmin");
            string CommandText = "SELECT * FROM Log where JobNo = '" + lblJobNo.Text + "'";
            OdbcConnection myConnection = new OdbcConnection(ConnectionString);
            OdbcCommand myCommand = new OdbcCommand(CommandText, myConnection);
            OdbcDataAdapter myAdapter = new OdbcDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet palletCardData = new DataSet();
            try
            {
                myConnection.Open();
                myAdapter.Fill(palletCardData);
            }
            catch (Exception ex)
            {
                throw (ex);
            }
            finally
            {
                myConnection.Close();
            }
            using (DataTable palletCardLog = new DataTable())
            {
                myAdapter.Fill(palletCardLog);
                dataGridView2.DataSource = palletCardLog;
            }

            // If This job Number has not yet been recorded in the database
            if (dataGridView2.Rows.Count == 0)
            {
                PalletNumber = 1;
            }
            // Otherwise check if any previous Pallet Numbers("Pallet Card" Routine entries) and record as the next sequential Pallet Number
            else
                {
                    try
                    {
                        // (There could be entries for this job Number but for Return or reject Paper)
                        ((DataTable)dataGridView2.DataSource).DefaultView.RowFilter = "Routine like 'Pallet Card'";
                    }
                    catch (Exception) { }

                    // if PalletNumber field is empty
                        if (dataGridView2.Rows[0].Cells[4].Value as string == "")
                        {
                            PalletNumber = 1;
                        }
                        else
                            {
                                this.dataGridView2.Sort(this.dataGridView2.Columns["PalletNumber"], ListSortDirection.Descending);
                                PalletNumber = (int)dataGridView2.Rows[0].Cells[4].Value + 1;
                            }
                }

            //SAVE TO DATABASE
            CurrentDate = DateTime.Now;
            produced = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", "")) - sheetsAffectedBadSection;
            dataGridView2.Refresh();
            var rowCount = dataGridView2.Rows.Count -1;
            string sqlFormattedDate = CurrentDate.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
            string Query = "insert into Log (Routine, JobNo, PalletNumber, PaperSectionNo, QtyRequired, ResourceID, WorkingSize, Description, SheetQty, Comment, Timestamp1, Produced, Unfinished) values('" + this.lbl1.Text + "','" + this.dataGridView1.Rows[0].Cells[0].Value  + "','" + PalletNumber + "','" + this.dataGridView1.Rows[0].Cells[19].Value + "', '" + this.dataGridView1.Rows[0].Cells[26].Value + "','" + resourceID + "','" + this.dataGridView1.Rows[0].Cells[13].Value + "','" + this.lbl2.Text + "','" + this.lbl5.Text + "','" + this.tbxExtraInfoComment.Text + "','" + CurrentDate + "','" + produced + "', '0');";
            SqlConnection conDatabase = new SqlConnection(constring);
            SqlCommand cmdDatabase = new SqlCommand(Query, conDatabase);
            SqlDataReader myReader;
            try
            {
                conDatabase.Open();
                myReader = cmdDatabase.ExecuteReader();
                conDatabase.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            // Requery the data to refresh dataGridView2 with the newly added PalletNumber and barCode
            OdbcConnection myConnection1 = new OdbcConnection(ConnectionString);
            OdbcCommand myCommand1 = new OdbcCommand(CommandText, myConnection1);
            OdbcDataAdapter myAdapter1 = new OdbcDataAdapter();
            myAdapter1.SelectCommand = myCommand1;
            DataSet palletCardData1 = new DataSet();
            try
            {
                myConnection.Open();
                myAdapter1.Fill(palletCardData1);
            }
            catch (Exception ex)
            {
                throw (ex);
            }
            finally
            {
                myConnection.Close();
            }
            using (DataTable palletCardLog = new DataTable())
            {
                myAdapter1.Fill(palletCardLog);
                dataGridView2.DataSource = palletCardLog;
            }

            this.dataGridView2.Sort(this.dataGridView2.Columns["PalletNumber"], ListSortDirection.Descending);
            string barCode = Convert.ToString(((int)dataGridView2.Rows[0].Cells[5].Value));
            Bitmap bitMap = new Bitmap(barCode.Length * 40, 80);
            using (Graphics graphics = Graphics.FromImage(bitMap))
            {
                Font oFont = new Font("IDAutomationHC39M", 16);
                PointF point = new PointF(2f, 2f);
                SolidBrush blackBrush = new SolidBrush(Color.Black);
                SolidBrush whiteBrush = new SolidBrush(Color.White);
                graphics.FillRectangle(whiteBrush, 0, 0, bitMap.Width, bitMap.Height);
                graphics.DrawString("*" + barCode + "*", oFont, blackBrush, point);
            }
            using (MemoryStream ms = new MemoryStream())
            {
                bitMap.Save(ms, ImageFormat.Png);
                pictureBox1.Image = bitMap;
                pictureBox1.Height = bitMap.Height;
                pictureBox1.Width = bitMap.Width;
            }

            pnlPalletCardPrint.BringToFront();
            btnBack.Visible = false;
            lblPC_JobNo.Text = lblJobNo.Text;
            lblPC_JobNo.Visible = true;
            lblPC_Customer.Text = dataGridView1.Rows[0].Cells[22].Value as string;
            lblPC_Customer.Visible = true;
            lblPC_Customer.MaximumSize = new Size(450, 220);
            lblPC_Customer.AutoSize = true;
            lblPC_SheetQty.Text = lbl5.Text;
            lblPC_SheetQty.Visible = true;
            lblPC_Press.Text = "Press - " + lblPress.Text;
            lblPC_Press.Visible = true;
            lblPC_Date.Text = "Date - " + DateTime.Now.ToString("d/M/yyyy");
            lblPC_Date.Visible = true;
            lblPC_Note.Text = tbxExtraInfoComment.Text + " - " + tbxTextBoxBadSection.Text;
            lblPC_Note.Visible = true;
            lblPC_PalletNumber.Text = "Pallet No " + PalletNumber.ToString();
            lblPC_PalletNumber.Visible = true;
            lblPC_Sig.Text = "Sheet " + dataGridView2.Rows[0].Cells[9].Value as string;
            lblPC_Sig.Visible = true;
            btnCancel.Visible = false;
            index = 17;
        }

        private void btnPalletCardPrint_Click(object sender, EventArgs e)
        {
            PrintDocument pd = new PrintDocument();
            pd.PrintPage += new PrintPageEventHandler(PrintImagePalletCard);
            btnPalletCardPrint.Visible = false;
            pd.Print();
            btnPalletCardPrint.Visible = true;
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

        void PrintImagePalletCard(object o, PrintPageEventArgs e)
        {
            int x = SystemInformation.WorkingArea.X;
            int y = SystemInformation.WorkingArea.Y;
            int width = this.Width;
            int height = this.Height;
            Rectangle bounds = new Rectangle(x, y, width, height);
            Bitmap img = new Bitmap(width, height);
            pnlPalletCardPrint.DrawToBitmap(img, bounds);
            Point p = new Point(100, 100);
            e.Graphics.DrawImage(img, p);
        }
     
        private void btnPalletOver_Click(object sender, EventArgs e)
        {
            pnlSignature.BringToFront();
        }

        private void tbxFinishPallet_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //MessageBox.Show(tbxFinishPallet.Text.ToString());
                string ConnectionString = Convert.ToString("Dsn=PalletCard;uid=PalletCardAdmin");
                string CommandText = "SELECT * FROM Log where AutoNum = '" + tbxFinishPallet.Text + "'";
                OdbcConnection myConnection = new OdbcConnection(ConnectionString);
                OdbcCommand myCommand = new OdbcCommand(CommandText, myConnection);
                OdbcDataAdapter myAdapter = new OdbcDataAdapter();
                myAdapter.SelectCommand = myCommand;
                DataSet palletCardData = new DataSet();
                try
                {
                    myConnection.Open();
                    myAdapter.Fill(palletCardData);
                }
                catch (Exception ex)
                {
                    throw (ex);
                }
                finally
                {
                    myConnection.Close();
                }
                using (DataTable palletCardLog = new DataTable())
                {
                    myAdapter.Fill(palletCardLog);
                    dataGridView2.DataSource = palletCardLog;
                }
                pnlPalletCard3.BringToFront();

                if (dataGridView2.Rows[0].Cells[25].Value != null)
                {
                    lbl2.Text = dataGridView2.Rows[0].Cells[25].Value.ToString();
                }
                else
                {
                    lbl2.Text = dataGridView2.Rows[0].Cells[22].Value.ToString();
                }

                lblJobNo.Text = dataGridView2.Rows[0].Cells[3].Value.ToString();
                lblJobNo.Visible = true;
                lblPress.Text = "XL106";
                lblPress.Visible = true;
                lbl1.Text = "Pallet Card";
                lbl1.Visible = true;
                lbl2.Visible = true;
                lbl3.Text = "Sheet " + dataGridView2.Rows[0].Cells[8].Value.ToString();
                lbl3.Visible = true;
                // WorkingSize
                lbl6.Text = dataGridView2.Rows[0].Cells[22].Value.ToString();
                // QtyRequired
                lbl7.Text = dataGridView2.Rows[0].Cells[34].Value.ToString();
            }
        }

        private void lblNumberUp_Click(object sender, EventArgs e)
        {

        }




        // NOTIFICATION PANEL
        private void btnNotificationPanel_Click(object sender, EventArgs e)
        {
            lbl1.Visible = true;
            lbl1.Text = "Notification Panel";
            pnlNotification1.BringToFront();
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
                pnlNotification2.BringToFront();
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
                        if (!(this.dataGridView1.Rows[i].Cells[11].Value as string == this.dataGridView1.Rows[i + 1].Cells[11].Value as string))
                        {
                            {
                                for (int j = 0; j < 1; j++)
                                {
                                    Button btn = new Button();
                                    pnlNotification1.Controls.Add(btn);
                                    btn.Top = A * 100;
                                    btn.Height = 80;
                                    btn.Width = 465;
                                    btn.BackColor = Color.SteelBlue;
                                    btn.Font = new Font("Microsoft Sans Serif", 14);
                                    btn.ForeColor = Color.White;
                                    btn.Left = 30;
                                    btn.Text = this.dataGridView1.Rows[i].Cells[11].Value as string;
                                    A = A + 1;
                                    btn.Click += new System.EventHandler(this.notification);
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
        private void notification(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            pnlNotification2.BringToFront();

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
            index = 18;
        }

        private void btnWaitingPlates_Click(object sender, EventArgs e)
        {
            // Send email notification
            //MailMessage mail = new MailMessage("WaitingPlates@colorman.ie", "declan.enright@colorman.ie", "Waiting for Plates", "Job Number " + lblJobNo.Text + " - Section " + dataGridView1.Rows[0].Cells[11].Value.ToString() + " - is waiting for plates");
            //SmtpClient client = new SmtpClient("ex0101.ColorMan.local");
            //client.Port = 25;
            //client.EnableSsl = false;
            //client.Send(mail);
            //MessageBox.Show("Email Notification Sent");

            MailMessage mail = new MailMessage();
            string from = "WaitingPlates@colorman.ie";
            mail.From = new MailAddress(from);
            mail.To.Add("declan.enright@colorman.ie");
            //mail.To.Add("prepress@colorman.ie");
            //mail.To.Add("production@colorman.ie");
            mail.Subject = "Waiting for Plates";
            mail.Body = "Waiting for Plates" + "Job Number " + lblJobNo.Text + " - Section " + dataGridView1.Rows[0].Cells[11].Value.ToString() + " - is waiting for plates";
            SmtpClient client = new SmtpClient("ex0101.ColorMan.local");
            client.Port = 25;
            client.EnableSsl = false;
            client.Send(mail);
            MessageBox.Show("Email Notification Sent");
        }

        private void btnWaitingPaper_Click(object sender, EventArgs e)
        {
            // Send email notification
            MailMessage mail = new MailMessage("WaitingPaper@colorman.ie", "declan.enright@colorman.ie", "Waiting for Paper", "Job Number " + lblJobNo.Text + " - Section " + dataGridView1.Rows[0].Cells[11].Value.ToString() + " - is waiting for paper");
            SmtpClient client = new SmtpClient("ex0101.ColorMan.local");
            client.Port = 25;
            client.EnableSsl = false;
            client.Send(mail);
            MessageBox.Show("Email Notification Sent");
        }

        private void btnJobLifted_Click(object sender, EventArgs e)
        {
            // Send email notification
            MailMessage mail = new MailMessage("JobLifted@colorman.ie", "declan.enright@colorman.ie", "Job Lifted", "Job Number " + lblJobNo.Text + " - Section " + dataGridView1.Rows[0].Cells[11].Value.ToString() + " - Job is lifted");
            SmtpClient client = new SmtpClient("ex0101.ColorMan.local");
            client.Port = 25;
            client.EnableSsl = false;
            client.Send(mail);
            MessageBox.Show("Email Notification Sent");
        }




    }
}
