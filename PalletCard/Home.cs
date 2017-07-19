using System;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Data.SqlClient;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;



namespace PalletCard
{
    public partial class Home : Form
    {
        List<Panel> listPanel = new List<Panel>();
        int index;
        bool sectionbtns;
        int A = 1;
        bool control;
        string jobNo;
        bool searchChanged;

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
            returnpaper0.BringToFront();
            lblJobNo.Visible = false;
            lblPress.Visible = false;
            lblReturnPaper.Visible = false;
            lbltextBoxDescription.Visible = false;
            lblPaperType.Visible = false;
            lblWorkingSize.Visible = false;
            searchBox.Text = "";
            searchBox.Focus();
            sectionbtns = false;
            tbxPalletHeight.Text = null;
            btnSearch.Visible = true;
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            PrintDocument pd = new PrintDocument();
            pd.PrintPage += new PrintPageEventHandler(PrintImage);
            btnPrint.Visible = false;
            //pd.Print();
            btnPrint.Visible = true;






            //ReportDocument myDataReport = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
            //myDataReport.Load(@"S:\\Production Admin\\Declan Enright\\Pallet Card Project\\Github\PalletCardApp\\PalletCard\\CrystalReport1.rpt");
            //myDataReport.SetParameterValue("JobNo", searchBox.Text);



            //ReportDocument myDataReport = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
            //myDataReport.Load(@"C:\Layouts\Report.rpt");
            //ParameterField myParam = new ParameterField();
            //ParameterDiscreteValue myDiscreteValue = new ParameterDiscreteValue();
            //myParam.ParameterFieldName = "MyParameter";
            //myDiscreteValue.Value = "Hello";
            //myParam.CurrentValues.Add(myDiscreteValue);
            //myDataReport.ParameterFields.Add(myParam);
            //myDataReport.SetDataSource(myDataTable);
            //myDataReport.SetParameterValue("MyParameter", "Hello");
            //Stream returnData = myDataReport.ExportToStream(PortableDocFormat);
            //myDataReport.Close();
            //return returnData;










            ReportDocument cryRpt = new ReportDocument();
            cryRpt.Load("S:\\Production Admin\\Declan Enright\\Pallet Card Project\\Github\\PalletCardApp\\PalletCard\\CrystalReport1.rpt");
            cryRpt.SetParameterValue("JobNo", searchBox.Text);

            TableLogOnInfos crtableLogoninfos = new TableLogOnInfos();
            TableLogOnInfo crtableLogoninfo = new TableLogOnInfo();
            ConnectionInfo crConnectionInfo = new ConnectionInfo();
            Tables CrTables;

            crConnectionInfo.ServerName = "CMDB2";
            crConnectionInfo.DatabaseName = "Thardata";
            crConnectionInfo.UserID = "PalletCardAdmin";
            crConnectionInfo.Password = "Pa!!etCard01";

            CrTables = cryRpt.Database.Tables;
            foreach (Table CrTable in CrTables)
            {
                crtableLogoninfo = CrTable.LogOnInfo;
                crtableLogoninfo.ConnectionInfo = crConnectionInfo;
                CrTable.ApplyLogOnInfo(crtableLogoninfo);
            }

            cryRpt.Refresh();
            cryRpt.PrintToPrinter(2, true, 1, 2);

















            //ParameterFields paramFields = crystalReportViewer1.ParameterFieldInfo;

            //string reportName; 
            //ReportDocument rpt = new ReportDocument();
            //reportName = "Works Instruction - Job Docket";
            //string reportPath = Application.StartupPath + "\\P:\\Live Reports & Documents\\Documents\\Works Instruction\\" + reportName;
            //rpt.Load(reportPath);
            //rpt.SetParameterValue("@JobNo", searchBox.Text.Trim());




            //CrystalDecisions.CrystalReports.Engine.ReportDocument cryRpt = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
            //cryRpt.Load("P:\\Live Reports & Documents\\Documents\\Works Instruction - Job Docket.rpt");

            //CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions crParameterFieldDefinitions;
            //CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition crParameterFieldDefinition;
            //CrystalDecisions.Shared.ParameterValues crParameterValues = new CrystalDecisions.Shared.ParameterValues();
            //CrystalDecisions.Shared.ParameterDiscreteValue crParameterDiscreteValue = new CrystalDecisions.Shared.ParameterDiscreteValue();

            //crParameterDiscreteValue.Value = searchBox.Text;
            //crParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields;
            //crParameterFieldDefinition = crParameterFieldDefinitions["JobNo"];
            //crParameterValues = crParameterFieldDefinition.CurrentValues;

            //crParameterValues.Clear();
            //crParameterValues.Add(crParameterDiscreteValue);
            //crParameterFieldDefinition.ApplyCurrentValues(crParameterValues);

            //crystalReportViewer1.ReportSource = cryRpt;
            //crystalReportViewer1.Refresh();





            //string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
            //string Query = "insert into Log (Routine, JobNo, ResourceID, Description, WorkingSize, SheetQty) values('" + this.lblReturnPaper.Text + "','" + this.dataGridView1.Rows[0].Cells[0].Value + "','" + this.dataGridView1.Rows[0].Cells[1].Value + "','" + this.lbltextBoxDescription.Text + "','" + this.lblWorkingSize.Text + "','" + this.lblPrint3.Text + "');";
            //SqlConnection conDatabase = new SqlConnection(constring);
            //SqlCommand cmdDatabase = new SqlCommand(Query, conDatabase);
            //SqlDataReader myReader;
            //try
            //{
            //    conDatabase.Open();
            //    myReader = cmdDatabase.ExecuteReader();
            //    MessageBox.Show("Saved");
            //    while (myReader.Read())
            //    {

            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
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

        public Home()
        {
            InitializeComponent();
        }

        

        private void btnBack_Click(object sender, EventArgs e)
        {
            if (index==0)
            {
                btnBack.Visible = false;
                btnSearch.Visible = false;
                lblJobNo.Visible = false;
                lblPress.Visible = false;
            }
            else if (index == 1)
            { 
            returnpaper0.BringToFront();
            lblReturnPaper.Visible = false;
            lbltextBoxDescription.Visible = false;
            lblPaperType.Visible = false;
            lblWorkingSize.Visible = false;
            tbxPalletHeight.Text = "";
            lblPheight.Text = "";
            index = 0;
            }
            else if (index == 2)
            { 
            returnpaper1.BringToFront();
            lblReturnPaper.Visible = false;
            lbltextBoxDescription.Visible = false;
            lblPaperType.Visible = false;    
            lblWorkingSize.Visible = false;
            index = 1;
            }
            else if (index == 3)
            { 
                returnpaper2.BringToFront();
                lblPheight.Text = "";
                tbxPalletHeight.Text = "";
                lbltextBoxDescription.Visible = false;
                lblPaperType.Visible = false;
                lblWorkingSize.Visible = false;
                index = 2;
                // if no section buttons go straight back to Choose Action screen
                if (returnpaper2.Controls.Count == 0)
                    {
                        returnpaper1.BringToFront();
                        lblReturnPaper.Visible = false;
                        lbltextBoxDescription.Visible = false;
                        lblPaperType.Visible = false;
                        lblWorkingSize.Visible = false;
                        lblPheight.Text = "";
                        lblPaperType.Visible = false;
                        lblWorkingSize.Visible = false;
                        tbxPalletHeight.Text = "";
                        index = 1;                      
                    }
                }
            else if (index == 4)
            { 
            returnpaper3.BringToFront();
            lblReturnPaper.Visible = true;
            lbltextBoxDescription.Visible = true;
            lblPaperType.Visible = true;
            lblWorkingSize.Visible = true;
            lblPheight.Text = "";
            tbxPalletHeight.Text = "";
            index = 3;
            }
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
                //if (resourceID == 5)
                //if (dataGridView1.RowCount > 0)
                if (dataGridView1.Rows[0].Cells[0].Value != null)
                {
                    lblPress.Text = "710UV";
                    lblPress.Visible = true;
                    returnpaper1.BringToFront();
                }
                else
                {
                    lblPress.Visible = false;
                    MessageBox.Show("The Job number you entered is not on this press");
                }
            }
            catch (Exception)
            {
                MessageBox.Show("The Job number you entered is not on this press");
            }
            index = 1;
            if (searchChanged == true)
                { 
                returnpaper2.Controls.Clear();
                }
            A = 1;
            btnBack.Visible = true;
        }
        
        private void btnReturnPaper_Click(object sender, EventArgs e)
        {
                lblReturnPaper.Visible = true;
                lblReturnPaper.Text = "Return Paper";
                returnpaper2.BringToFront();
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
                returnpaper3.BringToFront();
                string d = dataGridView1.Rows[0].Cells[11].Value.ToString();
                lbltextBoxDescription.Text = d;
                lbltextBoxDescription.Visible = true;
                lblPaperType.Visible = true;
                lblPaperType.Text = this.dataGridView1.Rows[0].Cells[16].Value as string;
                lblWorkingSize.Visible = true;
                lblWorkingSize.Text = dataGridView1.Rows[0].Cells[13].Value.ToString();
                index = 3;
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

                            //offer only one button where Expr1 field has two rows with the same value
                            if (! (this.dataGridView1.Rows[i].Cells[11].Value as string == this.dataGridView1.Rows[i+1].Cells[11].Value as string)) { 
                            {
                                    for (int j = 0; j < 1; j++)
                                    { 
                                        Button btn = new Button();
                                        this.returnpaper2.Controls.Add(btn);
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
                            }
                        }
                    }
                sectionbtns = true;
            }
        }

        //Dynamic button click - Section buttons, Return Paper work flow
        private void expr1(object sender, EventArgs e) {
            Button btn = sender as Button;
            returnpaper3.BringToFront();
            lblWorkingSize.Visible = true;
            lblWorkingSize.Text = dataGridView1.Rows[0].Cells[13].Value.ToString();
            lblPaperType.Visible = true;
            lblPaperType.Text = dataGridView1.Rows[0].Cells[16].Value.ToString();
            lbltextBoxDescription.Visible = true;
            lbltextBoxDescription.Text = btn.Text;
            tbxPalletHeight.Text = "";
            index = 3;    

            //filter datagridview1 with the button text choice
            try
            {
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = "Expr1 like '%" + this.dataGridView1.Rows[0].Cells[11].Value + "%'";
            }
            catch (Exception) { }


        }

        private void btnPalletCard_Click(object sender, EventArgs e)
        {

        }

        private void searchBox_TextChanged(object sender, EventArgs e)
        {
            searchChanged = true;
        }


        private void btnPalletHeight_Click(object sender, EventArgs e)
        {
            returnpaper4.BringToFront();
            lblPrint1.Text = dataGridView1.Rows[0].Cells[16].Value.ToString();
            lblPrint2.Text = dataGridView1.Rows[0].Cells[13].Value.ToString();
            lblPrint3.Text = lblPheight.Text;
            lblPrint4.Text = "Press - 710UV";
            lblPrint5.Text = "Job - " + jobNo;
            lblPrint6.Text = "Date - " + DateTime.Now.ToString("d/M/yyyy");
            index = 4;
        }

        // Pallet Height textBox calculation for Return Paper
        private void tbxPalletHeight_TextChanged(object sender, EventArgs e)
        {
            TextBox objTextBox = (TextBox)sender;
            int p1;
            int p2;
            if (!String.IsNullOrEmpty(tbxPalletHeight.Text))
            {
                p1 = Convert.ToInt32(objTextBox.Text);
                p2 = Convert.ToInt32(this.dataGridView1.Rows[0].Cells[20].Value);
                int result = p1 * p2 / 1000;
                string r1 = Convert.ToString(result);
                lblPheight.Text = (r1 + " sheets");
            }
        }

    }
}
