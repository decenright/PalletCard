﻿using System;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data;
using System.Drawing;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.IO;
using System.ComponentModel;
using System.Net.Mail;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using iTextSharp.text;
using iTextSharp.text.pdf;
using RawPrint;
using System.Linq;

namespace PalletCard
{
    public partial class Home : Form
    {

        #region Profiles

        //// Declan Testing
        //int resourceID = 6;
        //string press = "XL106";
        //string ConnectionString = Convert.ToString("Dsn=TharData;uid=tharuser");
        //string defaultEmail = "declan.enright@colorman.ie";
        //string defaultPrinter = "ProC5100S";
        ////string defaultPrinter = "ProC5100S (Pro C5100Sseries E-42B PS US1.1)";
        //////////string defaultPrinter = @"\\DC2012.ColorMan.local\Xerox 5335 PS Upstairs";

        // XL106
        int resourceID = 6;
        string press = "XL106";
        string ConnectionString = Convert.ToString("Dsn=TharData;uid=tharuser");
        //string defaultEmail = "martin@colorman.ie";
        string defaultEmail = "declan.enright@colorman.ie";
        string defaultPrinter = @"\\DC2012.ColorMan.local\Xerox 5335 PS Upstairs";

        //// SM102
        //int resourceID = 1;
        //string press = "SM102";
        //string ConnectionString = Convert.ToString("Dsn=TharData;uid=tharuser");
        ////string defaultEmail = "martin@colorman.ie";
        //string defaultEmail = "declan.enright@colorman.ie";
        //string defaultPrinter = @"\\DC2012.ColorMan.local\Xerox 5335 PS Upstairs";

        //// XL106UV
        //int resourceID = 67;
        //string press = "XL106UV";
        //string ConnectionString = Convert.ToString("Dsn=TharData;uid=tharuser");
        //string defaultEmail = "martin@colorman.ie";
        //string defaultPrinter = @"\\DC2012.ColorMan.local\Xerox 5335 PS Upstairs";
        ////string defaultEmail = "declan.enright@colorman.ie";
        ////string defaultPrinter = "ProC5100S (Pro C5100Sseries E-42B PS US1.1)";

        //// XL758
        //int resourceID = 68;
        //string press = "XL758";
        //string ConnectionString = Convert.ToString("Dsn=TharData;uid=tharuser");
        //string defaultEmail = "martin@colorman.ie";
        //string defaultPrinter = @"\\DC2012.ColorMan.local\Xerox 5335 PS Upstairs";

        #endregion

        #region Global
        List<Panel> listPanel = new List<Panel>();
        List<string> disableSectionButtons = new List<string>();
        List<string> allSections = new List<string>();
        List<string> completedSections = new List<string>();
        List<string> numberBadList = new List<string>(0);
        List<string> sheetsAffectedList = new List<string>(0);
        List<int> wholePalletList = new List<int>(0);
        int index;
        bool sectionBtns;
        bool sigBtns;
        bool badSectionLbls;
        bool backupRequired;
        bool varnishRequired;
        int DynamicButtonsOrigin = 1;
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
        int autoNum;
        int gangWholePalletButtonPressed;
        DateTime CurrentDate= DateTime.Now;
        decimal maxPercentageShort;
        int notGangedWholePalletValue;
        Boolean signed = false;
        string paperDetails = "";
        string inkDetails = "";
        string sig;
        int lastPallet = 0;
        int sumProduced;
        bool sectionFinishedClicked = false;
        bool cancelPrintMoreClicked = false;
        string badStation;
        bool isBackupVarnishFromScan = false;

        public Home()
        {
            Splash s = new Splash();
            s.Show();

            InitializeComponent();

            this.ActiveControl = textBox1;
            s.Close();
        }

        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();

        void Splash()
        {
            SplashScreen.SplashForm frm = new SplashScreen.SplashForm();
            frm.AppName = "";
            frm.Icon = Properties.Resources.Logo;
            frm.ShowIcon = true;
            frm.ShowInTaskbar = true;
            frm.BackgroundImage = Properties.Resources.Splash_Screen;

            timer.Interval = 5000;
            timer.Tick += new EventHandler(timer_Tick);


            if(timer.Interval < 5000)
            { 
            Application.Run(frm);          
            }
            frm.Close();
        }

        void timer_Tick(object sender, EventArgs e)
        {
            timer.Stop();
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
                btnCancel.Visible = true;
                tbxPalletHeight.Text = "";
                lblPheight.Text = "";
                Cancel();
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
                badSectionLbls = false;
                pnlPalletCard3.BringToFront();
                lbl4.Visible = false;
                lbl5.Visible = false;
                tbxPalletHeightPalletCard.Text = "";
                lblSheetCountPalletCard.Text = "";
                tbxSheetCountPalletCard.Text = "";
                lblPheightPalletCard.Text = "";
                tbxBadSectionComment.Text = "";
                tbxSheetsAffectedBadSection.Text = "";
                btnMarkBad.Enabled = true;
                btnMarkBad.BackColor = Color.SteelBlue;
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
                numberUp = 0;
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
                tbxBadSectionComment.Text = "";
                index = 12;
            }
            else if (index == 14)
            {
                pnlPalletCard3.BringToFront();
                tbxExtraInfoComment.Text = "";
                tbxPalletHeightPalletCard.Text = "";
                tbxSheetCountPalletCard.Text = "";
                tbxBadSectionComment.Text = "";
                tbxSheetsAffectedBadSection.Text = "";
                btnMarkBad.Enabled = true;
                btnMarkBad.BackColor = Color.SteelBlue;
                index = 10;
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
            this.WindowState = FormWindowState.Maximized;
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

                // make sure the the Operations table contains records
                if (operations.Rows.Count > 0) {

                    // New table to hold concatenated values - if is duplicate line filter then column = 0. Non duplicate lines filter column = 1. 
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
                        if (concatenatedTable.Rows[i][31].ToString() == concatenatedTable.Rows[i + 1][31].ToString())
                        {
                            concatenatedTable.Rows[i][32] = 0;
                        }
                        else
                            concatenatedTable.Rows[i][32] = 1;
                            concatenatedTable.Rows[concatenatedTable.Rows.Count - 1][32] = 1;

                            dataGridView1.DataSource = concatenatedTable;

                            // format StartOp to 24 hour time format and sort for earliest first
                            dataGridView1.Columns[30].DefaultCellStyle.Format = "MM/dd/yyyy HH:mm:ss";
                            dataGridView1.Sort(this.dataGridView1.Columns["StartOp"], ListSortDirection.Ascending);
                }              
            }
            btnBack.Visible = false;
            tbxSearchBox.Focus();
            btnMarkBad.Enabled = true;
            btnMarkBad.BackColor = Color.SteelBlue;
        }

        //Focus on Searchbox
        public void Home_SetFocus(object sender, EventArgs e)
        {
            tbxSearchBox.Focus();
        }

        private void Search()
        {   
                try
                {   
                    ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = string.Format("JobNo like '%{0}%' and Filter = 1 ", tbxSearchBox.Text.Trim().Replace("'", "''"));
                    lblJobNo.Text = dataGridView1.Rows[0].Cells[0].Value.ToString();
                    lblJobNo.Visible = true;
                    if (dataGridView1.Rows[0].Cells[0].Value != null)
                    {
                        lblPress.Text = press;
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
                    MessageBox.Show("The Job Number you entered is not on this Press");
                    Cancel();
                }
                index = 1;
                if (searchChanged == true)
                {
                    pnlReturnPaper1.Controls.Clear();
                    pnlNotification1.Controls.Clear();
                }
                //reset dynamic buttons origin
                DynamicButtonsOrigin = 1;

                btnBack.Visible = false;
                btnPalletCardPrint.Visible = true;
                sectionFinishedClicked = false;
                dataGridView1.Sort(this.dataGridView1.Columns["StartOp"], ListSortDirection.Ascending);

            // Disable Sig buttons that have already been recorded in the log table
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
                if (palletCardLog.Rows.Count != 0)
                    foreach (DataRow dr in palletCardLog.Rows)
                    {
                        if (Convert.ToInt32(dr["JobCancelled"]) != 1)
                        {
                            disableSectionButtons.Add(dr["PaperSectionNo"].ToString());
                        }                       
                    }                        
            }
        }       

        private void btnSearch_Click(object sender, EventArgs e)
        {
            Search();
        }

        private void renderCrystalReport()
        {
            ReportDocument cryRpt = new ReportDocument();
            cryRpt.Load("P:\\Live Reports & Documents\\Documents\\Works Instruction - Job Docket.rpt");

            ParameterFieldDefinitions crParameterFieldDefinitions;
            ParameterFieldDefinition crParameterFieldDefinition;
            ParameterValues crParameterValues = new ParameterValues();
            ParameterDiscreteValue crParameterDiscreteValue = new ParameterDiscreteValue();

            crParameterDiscreteValue.Value = lblJobNo.Text;
            crParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields;
            crParameterFieldDefinition = crParameterFieldDefinitions["JobNo"];
            crParameterValues = crParameterFieldDefinition.CurrentValues;

            crParameterValues.Clear();
            crParameterValues.Add(crParameterDiscreteValue);
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues);

            crystalReportViewer1.ReportSource = cryRpt;
            crystalReportViewer1.Refresh();
        }


        private void renderPalletCard()
        {
            ReportDocument cryRpt = new ReportDocument();

            cryRpt.Load("P:\\Live Reports & Documents\\Reports\\Pallet Card Front.rpt");

            // AutoNum
            ParameterFieldDefinitions crParameterFieldDefinitions;
            ParameterFieldDefinition crParameterFieldDefinition;
            ParameterValues crParameterValues = new ParameterValues();
            ParameterDiscreteValue crParameterDiscreteValue = new ParameterDiscreteValue();

            crParameterDiscreteValue.Value = autoNum;
            crParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields;
            crParameterFieldDefinition = crParameterFieldDefinitions["AutoNum"];
            crParameterValues = crParameterFieldDefinition.CurrentValues;

            crParameterValues.Clear();
            crParameterValues.Add(crParameterDiscreteValue);
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues);

            crystalReportViewer2.ReportSource = cryRpt;
            crystalReportViewer2.Refresh();
        }

        private void searchBox_TextChanged(object sender, EventArgs e)
        {
            searchChanged = true;
        }

        private void Cancel()
        {
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
                    if (concatenatedTable.Rows[i][31].ToString() == concatenatedTable.Rows[i + 1][31].ToString())
                    {
                        concatenatedTable.Rows[i][32] = 0;
                    }
                    else
                        concatenatedTable.Rows[i][32] = 1;
                        concatenatedTable.Rows[concatenatedTable.Rows.Count - 1][32] = 1;

                dataGridView1.DataSource = concatenatedTable;
                this.dataGridView1.Sort(this.dataGridView1.Columns["JobNo"], ListSortDirection.Ascending);
            }

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
            tbxBadSectionComment.Text = "";
            tbxExtraInfoComment.Text = "";
            lblPC_IncompletePallet.Visible = false;
            numberUp = 0;
            flowLayoutPanel2.Visible = true;
            pnlBadSectionGangHeader.Visible = true;
            tbxFinishPallet.Text = "";
            lastPallet = 0;
            lblPrinting.Visible = false;
            sectionFinishedClicked = false;
            btnScrollUp.Visible = false;
            btnScrollDown.Visible = false;
            btnMarkBad.Enabled = true;
            btnMarkBad.BackColor = Color.SteelBlue;

            // Reset Bad Section - flowLayoutPanel2
            pnlBadSectionGangHeader.Visible = false;
            lblStockCode.Visible = false;
            lblNumberUp.Visible = false;
            lblNumberUpBadQty.Visible = false;
            lblSheetsAffected.Visible = false;
            flowLayoutPanel2.Visible = false;
            btnScrollUp.Visible = false;
            btnScrollDown.Visible = false;
            sheetsAffectedBadSection = 0;
            dataGridView2.Columns.Clear();
            disableSectionButtons.Clear();
            sigBtns = false;
            pnlHome0.BringToFront();
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

#endregion

//****************************************************************************************************
//  RETURN PAPER WORKFLOW
//****************************************************************************************************
#region ReturnPaper

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
                                        btn.Top = DynamicButtonsOrigin * 100;
                                        btn.Height = 80;
                                        btn.Width = 465;
                                        btn.BackColor = Color.SteelBlue;
                                        btn.Font = new System.Drawing.Font("Microsoft Sans Serif", 14);
                                        btn.ForeColor = Color.White;
                                        btn.Left = 30;                                     
                                        btn.Text = this.dataGridView1.Rows[i].Cells[11].Value as string;
                                        DynamicButtonsOrigin = DynamicButtonsOrigin + 1;
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
            string Query = "insert into Log (Routine, JobNo, PalletNumber, Unfinished, PaperSectionNo, ResourceID, Expr1, WorkingSize, JobDesc, SheetQty, Description, Timestamp1) values('" + this.lbl1.Text + "', '" + this.dataGridView1.Rows[0].Cells[0].Value + "', '" + 0 + "', '" + this.dataGridView1.Rows[0].Cells[8].Value + "', '" + 0 + "', '" + resourceID + "', '" + this.lbl2.Text + "', '" + this.lbl4.Text + "', '" + this.dataGridView1.Rows[0].Cells[18].Value + "', '" + this.lblPrint3.Text + "', '" + this.lbl3.Text + "', '" + CurrentDate + "');";
            SqlConnection conDatabase = new SqlConnection(constring);
            SqlCommand cmdDatabase = new SqlCommand(Query, conDatabase);
            SqlDataReader myReader;
            try
            {
                conDatabase.Open();
                myReader = cmdDatabase.ExecuteReader();              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conDatabase.Close();
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
            System.Drawing.Rectangle bounds = new System.Drawing.Rectangle(x, y, width, height);
            Bitmap img = new Bitmap(width, height);
            pnlReturnPaper3.DrawToBitmap(img, bounds);
            Point p = new Point(100, 100);
            e.Graphics.DrawImage(img, p);
        }

#endregion

//****************************************************************************************************
//  REJECT PAPER WORKFLOW
//****************************************************************************************************
#region RejectPaper

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
                                        btn.Top = DynamicButtonsOrigin * 100;
                                        btn.Height = 80;
                                        btn.Width = 465;
                                        btn.BackColor = Color.SteelBlue;
                                        btn.Font = new System.Drawing.Font("Microsoft Sans Serif", 14);
                                        btn.ForeColor = Color.White;
                                        btn.Left = 30;
                                        btn.Text = this.dataGridView1.Rows[i].Cells[11].Value as string;
                                        DynamicButtonsOrigin = DynamicButtonsOrigin + 1;
                                        btn.Click += new System.EventHandler(this.expr2);
                                    }
                                }
                            }
                            dataGridView1.AllowUserToAddRows = false;
                    }
                }
                sectionBtns = true;
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

            string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
            string Query = "insert into Log (Routine, JobNo, PalletNumber, Unfinished, PaperSectionNo, ResourceID, Description, WorkingSize, JobDesc, SheetQty) values('" + this.lbl1.Text + "','" + this.dataGridView1.Rows[0].Cells[0].Value + "', '" + 0 + "', '" + 0 + "', '" + this.dataGridView1.Rows[0].Cells[8].Value + "', '" + resourceID + "','" + this.lbl2.Text + "', '" + this.lbl4.Text + "', '" + this.dataGridView1.Rows[0].Cells[18].Value + "', '" + this.lblPrint3.Text + "');";
            SqlConnection conDatabase = new SqlConnection(constring);
            SqlCommand cmdDatabase = new SqlCommand(Query, conDatabase);
            SqlDataReader myReader;
            try
            {
                conDatabase.Open();
                myReader = cmdDatabase.ExecuteReader();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conDatabase.Close();
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

        void PrintImageRejectPaper(object o, PrintPageEventArgs e)
        {
            int x = SystemInformation.WorkingArea.X;
            int y = SystemInformation.WorkingArea.Y;
            int width = this.Width;
            int height = this.Height;
            System.Drawing.Rectangle bounds = new System.Drawing.Rectangle(x, y, width, height);
            Bitmap img = new Bitmap(width, height);
            pnlRejectPaper3.DrawToBitmap(img, bounds);
            Point p = new Point(100, 100);
            e.Graphics.DrawImage(img, p);
        }
     
#endregion

//****************************************************************************************************
//  SIGNATURE
//****************************************************************************************************
#region Main Signature
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
                    signed = true;
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

        public void SaveSignatureImageToFile()
        {

                Bitmap bmp = new Bitmap(this.pnlSignature.Width, this.pnlSignature.Height);
                Graphics graphics = Graphics.FromImage(bmp);
                System.Drawing.Rectangle rect = pnlSignature.RectangleToScreen(pnlSignature.ClientRectangle);
                graphics.CopyFromScreen(rect.Location, Point.Empty, pnlSignature.Size);
                var time = DateTime.Now.ToString("yyyyMMdd-HH-mm-ss");
                var job = lblJobNo.Text.ToString();
                string all = job + " - " + PalletNumber + " - " + time;
                bmp.Save("P:/PalletCard/Signatures/ " + all + ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
        }

        private void btnQATravellerBlurb_Click(object sender, EventArgs e)
        {
            if (signed == true)
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

                this.dataGridView2.Sort(this.dataGridView2.Columns["AutoNum"], ListSortDirection.Descending);
                autoNum = Convert.ToInt32(dataGridView2.Rows[0].Cells[0].Value);
                pnlPalletCardPrint.BringToFront();
                btnCancel.Visible = false;
                signed = false;
                index = 17;
            }
                else
                {
                    MessageBox.Show("Please sign");
                    pnlSignature.BringToFront();
                    return;             
                }
        }       

        #endregion

#region POSA Signature

        private void SignaturePanelPosa_MouseMove(object sender, MouseEventArgs e)
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
                    DrawLinePosa(line);
                }
            }
        }

        private void SignaturePanelPosa_MouseUp(object sender, MouseEventArgs e)
        {
            IsCapturing = false;
            signature.Glyphs.Add(glyph);
            startPoint = new Point();
            endPoint = new Point();
        }

        private void SignaturePanelPosa_MouseDown(object sender, MouseEventArgs e)
        {
            IsCapturing = true;
            glyph = new Glyph();
        }

        private void DrawLinePosa(Line line)
        {
            using (Graphics graphic = this.SignaturePanelPosa.CreateGraphics())
            {
                graphic.DrawLine(pen, line.StartPoint, line.EndPoint);
            }
        }

        private void DrawSignaturePosa()
        {
            foreach (Glyph glyph in signature.Glyphs)
            {
                foreach (Line line in glyph.Lines)
                {
                    DrawLinePosa(line);
                }
            }
        }

        private void ClearSignaturePanelPosa()
        {
            using (Graphics graphic = this.SignaturePanelPosa.CreateGraphics())
            {
                SolidBrush solidBrush = new SolidBrush(Color.Gainsboro);
                graphic.FillRectangle(solidBrush, 0, 0, SignaturePanelPosa.Width, SignaturePanelPosa.Height);
            }
        }

        private void ClearSignaturePosa_Click(object sender, EventArgs e)
        {
            ClearSignaturePanelPosa();
        }

        public void SavePosaImageToFile()
        {
            Bitmap bmp = new Bitmap(this.pnlPalletCard10.Width, this.pnlPalletCard10.Height);
            Graphics graphics = Graphics.FromImage(bmp);
            System.Drawing.Rectangle rect = pnlPalletCard10.RectangleToScreen(pnlPalletCard10.ClientRectangle);
            graphics.CopyFromScreen(rect.Location, Point.Empty, pnlPalletCard10.Size);
            var time = DateTime.Now.ToString("yyyyMMdd-HH-mm-ss");
            var job = lblJobNo.Text.ToString();
            string all = job + " - " + time;
            bmp.Save("P:/PalletCard/BatchRecords/POSA/Job No " + all + ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg); 
        }




#endregion

#region MedicalPharma Signature

        private void SignaturePanelMedicalPharma_MouseMove(object sender, MouseEventArgs e)
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
                    DrawLineMedicalPharma(line);
                }
            }
        }

        private void SignaturePanelMedicalPharma_MouseUp(object sender, MouseEventArgs e)
        {
            IsCapturing = false;
            signature.Glyphs.Add(glyph);
            startPoint = new Point();
            endPoint = new Point();
        }

        private void SignaturePanelMedicalPharma_MouseDown(object sender, MouseEventArgs e)
        {
            IsCapturing = true;
            glyph = new Glyph();
        }

        private void DrawLineMedicalPharma(Line line)
        {
            using (Graphics graphic = this.SignaturePanelMedicalPharma.CreateGraphics())
            {
                graphic.DrawLine(pen, line.StartPoint, line.EndPoint);
            }
        }

        private void DrawSignatureMedicalPharma()
        {
            foreach (Glyph glyph in signature.Glyphs)
            {
                foreach (Line line in glyph.Lines)
                {
                    DrawLineMedicalPharma(line);
                }
            }
        }

        private void ClearSignaturePanelMedicalPharma()
        {
            using (Graphics graphic = this.SignaturePanelMedicalPharma.CreateGraphics())
            {
                SolidBrush solidBrush = new SolidBrush(Color.Gainsboro);
                graphic.FillRectangle(solidBrush, 0, 0, SignaturePanelMedicalPharma.Width, SignaturePanelMedicalPharma.Height);
            }
        }

        private void btnClearSignatureMedicalPharma_Click(object sender, EventArgs e)
        {
            ClearSignaturePanelMedicalPharma();
        }

        public void SaveMedicalPharmaImageToFile()
        {
            Bitmap bmp = new Bitmap(this.pnlPalletCard11.Width, this.pnlPalletCard11.Height);
            Graphics graphics = Graphics.FromImage(bmp);
            System.Drawing.Rectangle rect = pnlPalletCard11.RectangleToScreen(pnlPalletCard11.ClientRectangle);
            graphics.CopyFromScreen(rect.Location, Point.Empty, pnlPalletCard11.Size);
            var time = DateTime.Now.ToString("yyyyMMdd-HH-mm-ss");
            var job = lblJobNo.Text.ToString();
            string all = job + " - " + time;
            bmp.Save("P:/PalletCard/BatchRecords/MedicalPharma/Job No " + all + ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg);

        }

#endregion

//**************************************************************************************************** 
//  PALLET CARD
//****************************************************************************************************
#region PalletCard Search

        private void btnPalletCard_Click(object sender, EventArgs e)
        {
            lbl1.Visible = true;
            lbl1.Text = "Pallet Card";
            pnlPalletCard1.BringToFront();
            index = 8;
            tbxPalletHeightPalletCard.Focus();
            btnBack.Visible = true;
            getNumberUp();

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
                    sig = dataGridView1.Rows[0].Cells[19].Value.ToString();
                    //if (dataGridView2.Rows.Count != 0)
                    //{
                    //    sig = dataGridView2.Rows[0].Cells[8].ToString();
                    //}
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
                                        btn.Top = DynamicButtonsOrigin * 100;
                                        btn.Height = 80;
                                        btn.Width = 465;
                                        btn.BackColor = Color.SteelBlue;
                                        btn.Font = new System.Drawing.Font("Microsoft Sans Serif", 14);
                                        btn.ForeColor = Color.White;
                                        btn.Left = 30;
                                        btn.Text = this.dataGridView1.Rows[i].Cells[15].Value as string;
                                        DynamicButtonsOrigin = DynamicButtonsOrigin + 1;
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
                                                btn.Top = DynamicButtonsOrigin * 100;
                                                btn.Height = 80;
                                                btn.Width = 465;
                                                btn.BackColor = Color.SteelBlue;
                                                btn.Font = new System.Drawing.Font("Microsoft Sans Serif", 14);
                                                btn.ForeColor = Color.White;
                                                btn.Left = 30;
                                                btn.Text = this.dataGridView1.Rows[i].Cells[11].Value as string;
                                                DynamicButtonsOrigin = DynamicButtonsOrigin + 1;
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



            var distinctRowsPaperSectionNumber = (from DataGridViewRow row in dataGridView1.Rows select row.Cells[19].Value).Distinct().Count();
            var distinctRowsSectionName = (from DataGridViewRow row in dataGridView1.Rows select row.Cells[15].Value).Distinct().Count();
            var distinctRowsExpr1 = (from DataGridViewRow row in dataGridView1.Rows select row.Cells[11].Value).Distinct().Count();

            if (distinctRowsSectionName == 1 && distinctRowsExpr1 == 1 && distinctRowsPaperSectionNumber > 1)
            {
                lbl7.Text = dataGridView1.Rows[0].Cells[25].Value.ToString();

                string a;
                string b;
                //a = PaperSectionNo value at Row 0
                a = dataGridView1.Rows[0].Cells[19].Value.ToString();
                //initialize variable b with a value - this will change to PaperSectionNo value at row 1 once it enters the loop
                b = dataGridView1.Rows[0].Cells[19].Value.ToString();
                {
                    for (int i = 1; i < this.dataGridView1.Rows.Count; i++)
                    {
                        //b = PaperSectionNo value at Row 1
                        b = dataGridView1.Rows[i].Cells[19].Value.ToString();
                    }
                    if (a == b)
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

                    else
                    {
                        if (!sigBtns)
                        {
                            if (dataGridView1.Rows[0].Cells[15].Value.ToString() != "")
                            {
                                //create a button for each value of field "PaperSectionNo"  
                                DataTable data = (DataTable)(dataGridView1.DataSource);
                                DataView view = new DataView(data);
                                view.RowFilter = "SectionName = '" + dataGridView1.Rows[0].Cells[15].Value.ToString() + "' and JobNo like '%" + lblJobNo.Text + "%' and Filter = 1";
                                view.Sort = "PaperSectionNo";
                                DataTable distinctValues = view.ToTable(true, "PaperSectionNo");
                                flowLayoutPanel1.Controls.Clear();
                                foreach (DataRow row in distinctValues.Rows)
                                {
                                    Button btnSig = new Button();
                                    this.flowLayoutPanel1.Controls.Add(btnSig);
                                    btnSig.Text = row["PaperSectionNo"].ToString();
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
                                    btnSig.Font = new System.Drawing.Font("Microsoft Sans Serif", 20);
                                    btnSig.ForeColor = Color.White;
                                    btnSig.TextAlign = ContentAlignment.MiddleCenter;
                                    btnSig.Click += new System.EventHandler(this.sectionButtonSectionName);
                                    //get all the section numbers in a list for later                           
                                    allSections.Add(row["PaperSectionNo"].ToString());
                                    allSections = allSections.Distinct().ToList();
                                }
                                pnlPalletCard2.BringToFront();
                            }
                            else
                            {
                                //create a button for each value of field "PaperSectionNo"  
                                DataTable data = (DataTable)(dataGridView1.DataSource);
                                DataView view = new DataView(data);
                                view.RowFilter = "Expr1 = '" + dataGridView1.Rows[0].Cells[11].Value.ToString() + "' and JobNo like '%" + lblJobNo.Text + "%' and Filter = 1";
                                view.Sort = "PaperSectionNo";
                                DataTable distinctValues = view.ToTable(true, "PaperSectionNo");
                                flowLayoutPanel1.Controls.Clear();
                                foreach (DataRow row in distinctValues.Rows)
                                {
                                    Button btnSig = new Button();
                                    this.flowLayoutPanel1.Controls.Add(btnSig);
                                    btnSig.Text = row["PaperSectionNo"].ToString();
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
                                    btnSig.Font = new System.Drawing.Font("Microsoft Sans Serif", 20);
                                    btnSig.ForeColor = Color.White;
                                    btnSig.TextAlign = ContentAlignment.MiddleCenter;
                                    btnSig.Click += new System.EventHandler(this.sectionButtonExpr1);
                                    //get all the section numbers in a list for later                           
                                    allSections.Add(row["PaperSectionNo"].ToString());
                                    allSections = allSections.Distinct().ToList();
                                }
                                pnlPalletCard2.BringToFront();
                            }

                        }
                        sigBtns = true;
                    }
                }
            }













        }

        //Dynamic button click - Part buttons, SECTION_NAME, Pallet Card work flow
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
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = "SectionName = '" + btn.Text + "' and JobNo like '%" + lblJobNo.Text + "%' and Filter = 1";
            }
            catch (Exception) { }

            lbl7.Text = dataGridView1.Rows[0].Cells[25].Value.ToString();

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
                            //create a button for each value of field "PaperSectionNo"  
                            DataTable data = (DataTable)(dataGridView1.DataSource);
                            DataView view = new DataView(data);
                            view.RowFilter = "SectionName = '" + btn.Text + "' and JobNo like '%" + lblJobNo.Text + "%' and Filter = 1";
                            view.Sort = "PaperSectionNo";
                            DataTable distinctValues = view.ToTable(true, "PaperSectionNo");
                            flowLayoutPanel1.Controls.Clear();
                            foreach (DataRow row in distinctValues.Rows)
                            {
                                Button btnSig = new Button();
                                this.flowLayoutPanel1.Controls.Add(btnSig);
                                btnSig.Text = row["PaperSectionNo"].ToString();
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
                                btnSig.Font = new System.Drawing.Font("Microsoft Sans Serif", 20);
                                btnSig.ForeColor = Color.White;
                                btnSig.TextAlign = ContentAlignment.MiddleCenter;
                                btnSig.Click += new System.EventHandler(this.sectionButtonSectionName);
                                //get all the section numbers in a list for later                           
                                allSections.Add(row["PaperSectionNo"].ToString());
                                allSections = allSections.Distinct().ToList();
                            }
                        }
                        sigBtns = true;
                     }
            }
        }

        //Dynamic button click - Part buttons, EXPR1, Pallet Card work flow
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
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = "Expr1 = '" + btn.Text + "' and JobNo like '%" + lblJobNo.Text + "%' and Filter = 1";
            }
            catch (Exception) { }

            lbl7.Text = dataGridView1.Rows[0].Cells[25].Value.ToString();

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
                            //create a button for each value of field "PaperSectionNo"  
                            DataTable data = (DataTable)(dataGridView1.DataSource);
                            DataView view = new DataView(data);
                            view.RowFilter = "SectionName = '" + btn.Text + "' and JobNo like '%" + lblJobNo.Text + "%' and Filter = 1";
                            view.Sort = "PaperSectionNo";
                            DataTable distinctValues = view.ToTable(true, "PaperSectionNo");
                            flowLayoutPanel1.Controls.Clear();
                            foreach (DataRow row in distinctValues.Rows)
                            {
                                Button btnSig = new Button();
                                this.flowLayoutPanel1.Controls.Add(btnSig);
                                btnSig.Text = row["PaperSectionNo"].ToString();
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
                                btnSig.Font = new System.Drawing.Font("Microsoft Sans Serif", 20);
                                btnSig.ForeColor = Color.White;
                                btnSig.TextAlign = ContentAlignment.MiddleCenter;
                                btnSig.Click += new System.EventHandler(this.sectionButtonSectionName);
                                //get all the section numbers in a list for later
                                allSections.Add(row["PaperSectionNo"].ToString());
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
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = "SectionName like '%" + lbl2.Text + "%'  and PaperSectionNo = " + btn.Text.Trim() + " and JobNo like '%" + lblJobNo.Text + "%' and Filter = 1";
            }
            catch (Exception) { }

            lbl7.Text = dataGridView1.Rows[0].Cells[25].Value.ToString();
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
                ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = "Expr1 like '%" + lbl2.Text + "%'  and PaperSectionNo = " + btn.Text.Trim() + " and JobNo like '%" + lblJobNo.Text + "%' and Filter = 1";
            }
            catch (Exception) { }

            lbl7.Text = dataGridView1.Rows[0].Cells[25].Value.ToString();
        }

        private void tbxFinishPallet_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnFinishPallet.PerformClick();
            }
        }

        private void btnFinishPallet_Click(object sender, EventArgs e)
        {
            string ConnectionString = Convert.ToString("Dsn=PalletCard;uid=PalletCardAdmin");
            string CommandText = "SELECT * FROM Log where AutoNum = '" + Convert.ToInt32(tbxFinishPallet.Text) + "'";
            OdbcConnection myConnection = new OdbcConnection(ConnectionString);
            OdbcCommand myCommand = new OdbcCommand(CommandText, myConnection);
            OdbcDataAdapter myAdapter = new OdbcDataAdapter();
            myAdapter.SelectCommand = myCommand;
            DataSet palletCardData = new DataSet();
            try
            {
                myConnection.Open();
                myAdapter.Fill(palletCardData);
                using (DataTable palletCardLog = new DataTable())
                {
                    myAdapter.Fill(palletCardLog);
                    dataGridView2.DataSource = palletCardLog;
                }
                pnlPalletCard3.BringToFront();

            }
            catch (Exception ex1)
            {
                MessageBox.Show(ex1.Message);
            }
            finally
            {
                myConnection.Close();
            }


            if (dataGridView2.Rows.Count > 0)
            {
                try
                {
                    if (dataGridView2.Rows[0].Cells[24].Value != null)
                    {
                        lbl2.Text = dataGridView2.Rows[0].Cells[24].Value.ToString();
                    }
                    else
                    {
                        lbl2.Text = dataGridView2.Rows[0].Cells[20].Value.ToString();
                    }
                }
                catch (Exception ex2)
                {
                    MessageBox.Show(ex2.Message);
                }

                lblJobNo.Text = dataGridView2.Rows[0].Cells[3].Value.ToString();
                lblJobNo.Visible = true;
                lblPress.Text = press;
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
                isBackupVarnishFromScan = true;
            }
            if (lbl2.Visible == false)
            {
                MessageBox.Show("Please enter a valid Job Number");
                Cancel();
                tbxFinishPallet.Text = "";
                tbxFinishPallet.Focus();
                lblJobNo.Visible = false;
                lblPress.Visible = false;
                lbl1.Visible = false;
                lbl2.Visible = false;
                lbl3.Visible = false;
                dataGridView2.Columns.Clear();
            }
        }

#endregion

#region Pallet Height Sheet Count

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

        private void tbxSheetCountPalletCard_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                btnPalletHeightSheetCountPalletCard_Click(tbxPalletHeightPalletCard, new EventArgs());
        }

        private void tbxPalletHeightPalletCard_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                btnPalletHeightSheetCountPalletCard_Click(tbxSheetCountPalletCard, new EventArgs());
        }

        #endregion

        #region Gang

        // GANGPRO ROUTINE

        private void queryGangpro()
        {
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
                gangPro.Columns.Add("QtyRequired", typeof(int));
                gangPro.Columns.Add("ProdQtyRequired", typeof(int));
                gangPro.Columns.Add("NumberUpBad", typeof(int));
                gangPro.Columns.Add("SheetsAffected", typeof(int));
                gangPro.Columns.Add("SheetsUnaffected", typeof(int));
                gangPro.Columns.Add("SheetsProduced", typeof(int));
                gangPro.Columns.Add("QtyGoodProduced", typeof(int));
                gangPro.Columns.Add("QtyShort", typeof(int));
                gangPro.Columns.Add("PercentageShort", typeof(decimal));
                gangPro.Columns.Add("PercentageShort1", typeof(string));
                gangPro.Columns.Add("PercentageShort2", typeof(decimal));
                gangPro.Columns.Add("PercentageShort3");

                DataTable gangProTable = new DataTable();
                gangProTable = gangPro.Clone();
                gangProTable.Columns["QtyRequired"].Expression = " '" + qtyRequired + "'  ";
                gangProTable.Columns["SheetsProduced"].Expression = " '" + sheetsProduced + "' ";
                gangProTable.Columns["QtyShort"].Expression = " '" + sheetsProduced + "'  ";
                gangProTable.Columns["ProdQtyRequired"].Expression = "QtyRequired * NumberUp";
                gangProTable.Columns["SheetsAffected"].Expression = tbxSheetsAffectedBadSection.Text;
                gangProTable.Columns["SheetsUnaffected"].Expression = "SheetsProduced - SheetsAffected";
                gangProTable.Columns["QtyGoodProduced"].Expression = "(SheetsAffected * (NumberUp - NumberUpBad)) + (SheetsUnaffected * NumberUp)";
                gangProTable.Columns["QtyShort"].Expression = "QtyGoodProduced - ProdQtyRequired";
                gangProTable.Columns["PercentageShort"].Expression = "QtyShort / ProdQtyRequired";
                gangProTable.Columns["PercentageShort1"].Expression = "-1.0";
                gangProTable.Columns["PercentageShort2"].Expression = "PercentageShort1";
                gangProTable.Columns["PercentageShort3"].Expression = "PercentageShort * PercentageShort2";


                if (numberBadList.Count != 0)
                {
                    for (int i = 0; i < gangPro.Rows.Count; i++)
                    {
                        gangPro.Rows[i][10] = numberBadList[i];
                        gangPro.Rows[i][11] = sheetsAffectedList[i];
                        if (gangWholePalletButtonPressed == 1)
                        {
                            gangPro.Rows[i][11] = wholePalletList[i];
                        }
                    }
                }

                foreach (DataRow dr in gangPro.Rows)
                {
                    gangProTable.Rows.Add(dr.ItemArray);
                }
                dataGridView3.DataSource = gangProTable;

                // Find Qty bad to return to the main flow (sheetsAffectedBadSection)
                if (dataGridView3.Rows[0].Cells[19].Value.ToString() != "")
                {
                    this.dataGridView3.Sort(this.dataGridView3.Columns[19], ListSortDirection.Descending);
                    maxPercentageShort = Convert.ToDecimal(dataGridView3.Rows[0].Cells[19].Value);
                }

                sheetsAffectedBadSection = Convert.ToInt32((qtyRequired * (1 + maxPercentageShort)) - sheetsProduced);

                // make sure it doesn't return a negative number
                if (sheetsAffectedBadSection < 0)
                {
                    sheetsAffectedBadSection = 0;
                }

                // Dont show OK button if Number Up(cell 10) and Sheets affected(cell 11) are empty for gang Classic
                for (int i = 0; i < gangPro.Rows.Count; i++)
                {
                    if (dataGridView3.Rows[i].Cells[10].Value.ToString() != "" & dataGridView3.Rows[i].Cells[10].Value.ToString() != "0" & dataGridView3.Rows[i].Cells[11].Value.ToString() != "" & dataGridView3.Rows[i].Cells[11].Value.ToString() != "0")
                    {
                        btnBadSectionOK.Visible = true;
                    }
                }
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

                // Find Qty bad to return to the main flow (sheetsAffectedBadSection)
                if (dataGridView4.Rows[0].Cells[22].Value.ToString() != "")
                    {
                    this.dataGridView4.Sort(this.dataGridView4.Columns[22], ListSortDirection.Descending);
                    maxPercentageShort = Convert.ToDecimal(dataGridView4.Rows[0].Cells[22].Value);                    
                    }

                sheetsAffectedBadSection = Convert.ToInt32((qtyRequired * (1 + maxPercentageShort)) - sheetsProduced);

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
            var calc = Convert.ToDecimal(sheetsAffected) * (Convert.ToDecimal(badQty) / Convert.ToDecimal(numberUp));
            sheetsAffectedBadSection = Convert.ToInt32(calc);

            // make sure it doesn't return a negative number
            if (sheetsAffectedBadSection < 0)
            {
                sheetsAffectedBadSection = 0;
            }

            //if (gangWholePalletButtonPressed == 1 )
            //{
            //    //sheetsAffectedBadSection = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
            //    //lbl5.Text = sheetsAffectedBadSection.ToString();
            //}

            //MessageBox.Show(sheetsAffectedBadSection.ToString());

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
            lbl7.Text = dataGridView1.Rows[0].Cells[25].Value.ToString();
        }

        private void btnMarkBad_Click(object sender, EventArgs e)
        {
            this.flowLayoutPanel2.Controls.Clear();
            pnlPalletCard5.BringToFront();
            this.ActiveControl = tbxSheetsAffectedBadSection;
            tbxSheetsAffectedBadSection.Visible = true;
            lblSheets_Affected.Visible = true;
            btnBadSectionOK.Visible = false;
            index = 12;
            btnMarkBad.Enabled = false;
            btnMarkBad.BackColor = Color.Gray;


            // Hide Gang Panel (FlowLayoutPanel2) if NumberUp = 1 (Value will be from DataGridview1 or DataGridView2 depending on wheteher it is a Regular or incomplete i.e Scanned line)
            // Else Show Gang Classic, Gang Pro or Not Ganged Multi Up

            #region Regular Line
            if (dataGridView2.Rows.Count == 0)
            {
                //DATAGRIDVIEW1
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
                // if not equal to 1 then NumberUp =:
                numberUp = Convert.ToInt32(dataGridView1.Rows[0].Cells[12].Value);

                //if NumberUp field and JobGanged Field = 0
                if (numberUp == 0 & Convert.ToInt32(dataGridView1.Rows[0].Cells[14].Value) == 0)
                {
                    // Parse the NumberUp value from Section Name or Expr1 - This regex finds continuous digits before "up"
                    if (dataGridView1.Rows[0].Cells[11].Value.ToString().Contains("up"))
                    {
                        String text = dataGridView1.Rows[0].Cells[11].Value.ToString();
                        foreach (Match match in Regex.Matches(text, @"(\d+)up"))
                        {
                            //MessageBox.Show(match.Groups[1].Value);
                            numberUp = Convert.ToInt32(match.Groups[1].Value);
                        }
                    }
                    else
                    {
                        String text1 = dataGridView1.Rows[0].Cells[15].Value.ToString();
                        foreach (Match match in Regex.Matches(text1, @"(\d+)up"))
                        {
                            //MessageBox.Show(match.Groups[1].Value);
                            numberUp = Convert.ToInt32(match.Groups[1].Value);
                        }
                    }
                }

                qtyRequired = Convert.ToInt32(dataGridView1.Rows[0].Cells[25].Value);

                // IF NUMBER UP != 1 SHOW GANG PANEL AND HIDE SHEETS AFFECTED BOX
                //if (numberUp != 1)
                    if (numberUp > 1)
                    {
                    tbxSheetsAffectedBadSection.Visible = false;
                    lblSheets_Affected.Visible = false;

                    flowLayoutPanel2.Visible = true;
                    btnScrollUp.Visible = true;
                    btnScrollDown.Visible = true;
                    pnlBadSectionGangHeader.Visible = true;
                    lblStockCode.Visible = true;
                    lblNumberUp.Visible = true;
                    lblNumberUpBadQty.Visible = true;
                    lblSheetsAffected.Visible = true;
                    dataGridView4.Visible = true;
                    dataGridView3.Visible = true;

                    // Filter for Ganged jobs
                    if (Convert.ToInt32(dataGridView1.Rows[0].Cells[14].Value) != 0 & Convert.ToInt32(dataGridView1.Rows[0].Cells[14].Value) != 2)
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
                        if (!badSectionLbls)
                        {
                            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                            {
                                btnScrollUp.Visible = false;
                                btnScrollDown.Visible = false;
                                lblStockCode.Text = "Job\r\nNumber";
                                lblStockCode.Visible = true;
                                lblNumberUp.Visible = true;
                                lblNumberUp.Text = "Number\r\nUp";
                                lblNumberUpBadQty.Visible = true;
                                lblNumberUpBadQty.Text = "Bad\r\nStations";
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
                                lbl2.Width = 150;
                                lbl2.BackColor = Color.Silver;
                                lbl2.Font = new System.Drawing.Font("Microsoft Sans Serif", 20);
                                lbl2.TextAlign = ContentAlignment.MiddleLeft;
                                lbl2.ForeColor = Color.Black;
                                lbl2.Margin = new Padding(0, 0, 0, 0);
                                lbl2.Left = 40;
                                lbl2.Text = this.dataGridView1.Rows[i].Cells[0].Value.ToString();
                                Label lbl3 = new Label();
                                this.flowLayoutPanel2.Controls.Add(lbl3);
                                lbl3.Height = 40;
                                lbl3.Width = 78;
                                lbl3.BackColor = Color.Silver;
                                lbl3.Font = new System.Drawing.Font("Microsoft Sans Serif", 20);
                                lbl3.TextAlign = ContentAlignment.MiddleCenter;
                                lbl3.ForeColor = Color.Black;
                                lbl3.Margin = new Padding(0, 0, 0, 0);
                                lbl3.Left = 40;
                                lbl3.Text = this.dataGridView1.Rows[i].Cells[12].Value.ToString();
                                if (dataGridView1.Rows[i].Cells[12].Value.ToString() == "0")
                                {
                                    lbl3.Text = numberUp.ToString();
                                }
                                TextBox textBox1 = new TextBox();
                                this.flowLayoutPanel2.Controls.Add(textBox1);
                                textBox1.Height = 40;
                                textBox1.AutoSize = false;
                                textBox1.Width = 70;
                                textBox1.Multiline = false;
                                textBox1.Font = new System.Drawing.Font(textBox1.Font.FontFamily, 20);
                                textBox1.TextAlign = HorizontalAlignment.Center;
                                textBox1.Margin = new Padding(0, 0, 0, 0);
                                textBox1.TextChanged += new System.EventHandler(this.notGangedNumberUpBad);
                                TextBox textBox2 = new TextBox();
                                this.flowLayoutPanel2.Controls.Add(textBox2);
                                textBox2.Height = 40;
                                textBox2.AutoSize = false;
                                textBox2.Width = 85;
                                textBox2.Multiline = false;
                                textBox2.Font = new System.Drawing.Font(textBox2.Font.FontFamily, 20);
                                textBox2.TextAlign = HorizontalAlignment.Center;
                                textBox2.Margin = new Padding(0, 0, 0, 0);
                                textBox2.TextChanged += new System.EventHandler(notgangedSheetsAffected);
                                Button btn1 = new Button();
                                flowLayoutPanel2.Controls.Add(btn1);
                                btn1.Height = 40;
                                btn1.Width = 90;
                                btn1.BackColor = Color.SteelBlue;
                                btn1.ForeColor = Color.White;
                                btn1.Font = new System.Drawing.Font(textBox1.Font.FontFamily, 9);
                                btn1.Text = "Whole Pallet";
                                btn1.Margin = new Padding(0, 0, 0, 0);
                                btn1.Tag = i;
                                btn1.Click += new System.EventHandler(notGangedWholePallet);
                            }
                        }
                        badSectionLbls = true;
                    }

                    //if JobGanged = 1 (PALLET_GANG_CLASSIC Table)
                    else if (Convert.ToInt32(dataGridView1.Rows[0].Cells[14].Value) == 1)
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
                                lblNumberUpBadQty.Text = "Bad\r\nStations";
                                lblSheetsAffected.Visible = true;
                                lblSheetsAffected.Text = "Sheets\r\nAffected";
                                flowLayoutPanel2.HorizontalScroll.Visible = false;
                                flowLayoutPanel2.VerticalScroll.Visible = false;
                                Label lbl1 = new Label();
                                this.flowLayoutPanel2.Controls.Add(lbl1);
                                lbl1.Height = 35;
                                lbl1.Width = 430;
                                lbl1.BackColor = Color.Gray;
                                lbl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12);
                                lbl1.TextAlign = ContentAlignment.MiddleLeft;
                                lbl1.ForeColor = Color.White;
                                lbl1.Left = 40;
                                lbl1.Text = this.dataGridView4.Rows[i].Cells[3].Value.ToString();
                                Label lbl2 = new Label();
                                this.flowLayoutPanel2.Controls.Add(lbl2);
                                lbl2.Height = 40;
                                lbl2.Width = 150;
                                lbl2.BackColor = Color.Silver;
                                lbl2.Font = new System.Drawing.Font("Microsoft Sans Serif", 20);
                                lbl2.TextAlign = ContentAlignment.MiddleLeft;
                                lbl2.ForeColor = Color.Black;
                                lbl2.Margin = new Padding(0, 0, 0, 0);
                                lbl2.Left = 40;
                                lbl2.Text = this.dataGridView4.Rows[i].Cells[1].Value.ToString();
                                Label lbl3 = new Label();
                                this.flowLayoutPanel2.Controls.Add(lbl3);
                                lbl3.Height = 40;
                                lbl3.Width = 78;
                                lbl3.BackColor = Color.Silver;
                                lbl3.Font = new System.Drawing.Font("Microsoft Sans Serif", 20);
                                lbl3.TextAlign = ContentAlignment.MiddleCenter;
                                lbl3.ForeColor = Color.Black;
                                lbl3.Margin = new Padding(0, 0, 0, 0);
                                lbl3.Left = 40;
                                lbl3.Text = this.dataGridView4.Rows[i].Cells[11].Value.ToString();
                                if (dataGridView4.Rows[i].Cells[12].Value.ToString() == "0")
                                {
                                    lbl3.Text = numberUp.ToString();
                                }
                                TextBox textBox1 = new TextBox();
                                this.flowLayoutPanel2.Controls.Add(textBox1);
                                textBox1.Height = 40;
                                textBox1.AutoSize = false;
                                textBox1.Width = 70;
                                textBox1.Multiline = false;
                                textBox1.Font = new System.Drawing.Font(textBox1.Font.FontFamily, 20);
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
                                textBox2.Font = new System.Drawing.Font(textBox2.Font.FontFamily, 20);
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
                                btn1.Font = new System.Drawing.Font(textBox1.Font.FontFamily, 9);
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

                    //if JobGanged = 3 (PALLET_GANG_CLASSIC Table)
                    else if (Convert.ToInt32(dataGridView1.Rows[0].Cells[14].Value) == 3)
                    {
                        if (!badSectionLbls)
                        {
                            for (int i = 0; i < dataGridView3.Rows.Count; i++)
                            {
                                lblStockCode.Text = "Stock Code/\r\nJob Number";
                                lblStockCode.Visible = true;
                                lblNumberUp.Visible = true;
                                lblNumberUp.Text = "Number\r\nUp";
                                lblNumberUpBadQty.Visible = true;
                                lblNumberUpBadQty.Text = "Bad\r\nStations";
                                lblSheetsAffected.Visible = true;
                                lblSheetsAffected.Text = "Sheets\r\nAffected";
                                flowLayoutPanel2.HorizontalScroll.Visible = false;
                                flowLayoutPanel2.VerticalScroll.Visible = false;
                                Label lbl1 = new Label();
                                this.flowLayoutPanel2.Controls.Add(lbl1);
                                lbl1.Height = 35;
                                lbl1.Width = 430;
                                lbl1.BackColor = Color.Gray;
                                lbl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12);
                                lbl1.TextAlign = ContentAlignment.MiddleLeft;
                                lbl1.ForeColor = Color.White;
                                lbl1.Left = 40;
                                lbl1.Text = this.dataGridView3.Rows[i].Cells[4].Value.ToString();
                                Label lbl2 = new Label();
                                this.flowLayoutPanel2.Controls.Add(lbl2);
                                lbl2.Height = 40;
                                lbl2.Width = 150;
                                lbl2.BackColor = Color.Silver;
                                lbl2.Font = new System.Drawing.Font("Microsoft Sans Serif", 20);
                                lbl2.TextAlign = ContentAlignment.MiddleLeft;
                                lbl2.ForeColor = Color.Black;
                                lbl2.Margin = new Padding(0, 0, 0, 0);
                                lbl2.Left = 40;
                                lbl2.Text = this.dataGridView3.Rows[i].Cells[1].Value.ToString();
                                Label lbl3 = new Label();
                                this.flowLayoutPanel2.Controls.Add(lbl3);
                                lbl3.Height = 40;
                                lbl3.Width = 78;
                                lbl3.BackColor = Color.Silver;
                                lbl3.Font = new System.Drawing.Font("Microsoft Sans Serif", 20);
                                lbl3.TextAlign = ContentAlignment.MiddleCenter;
                                lbl3.ForeColor = Color.Black;
                                lbl3.Margin = new Padding(0, 0, 0, 0);
                                lbl3.Left = 40;
                                lbl3.Text = this.dataGridView3.Rows[i].Cells[3].Value.ToString();
                                if (dataGridView3.Rows[i].Cells[12].Value.ToString() == "0")
                                {
                                    lbl3.Text = numberUp.ToString();
                                }
                                TextBox textBox1 = new TextBox();
                                this.flowLayoutPanel2.Controls.Add(textBox1);
                                textBox1.Height = 40;
                                textBox1.AutoSize = false;
                                textBox1.Width = 70;
                                textBox1.Multiline = false;
                                textBox1.Font = new System.Drawing.Font(textBox1.Font.FontFamily, 20);
                                textBox1.TextAlign = HorizontalAlignment.Center;
                                textBox1.Margin = new Padding(0, 0, 0, 0);
                                textBox1.Tag = i;
                                textBox1.TextChanged += new System.EventHandler(this.gangProNumberUpBad);
                                TextBox textBox2 = new TextBox();
                                this.flowLayoutPanel2.Controls.Add(textBox2);
                                textBox2.Height = 40;
                                textBox2.AutoSize = false;
                                textBox2.Width = 85;
                                textBox2.Multiline = false;
                                textBox2.Font = new System.Drawing.Font(textBox2.Font.FontFamily, 20);
                                textBox2.TextAlign = HorizontalAlignment.Center;
                                textBox2.Margin = new Padding(0, 0, 0, 0);
                                textBox2.Tag = i;
                                textBox2.TextChanged += new System.EventHandler(gangProSheetsAffected);
                                Button btn1 = new Button();
                                flowLayoutPanel2.Controls.Add(btn1);
                                btn1.Height = 40;
                                btn1.Width = 73;
                                btn1.BackColor = Color.SteelBlue;
                                btn1.ForeColor = Color.White;
                                btn1.Font = new System.Drawing.Font(textBox1.Font.FontFamily, 9);
                                btn1.Text = "Whole Pallet";
                                btn1.Margin = new Padding(0, 0, 0, 0);
                                btn1.Tag = i;
                                btn1.Click += new System.EventHandler(this.gangProWholePallet);
                                numberBadList.Insert(i, "0");
                                sheetsAffectedList.Insert(i, "0");
                                wholePalletList.Insert(i, 0);
                            }
                        }
                        badSectionLbls = true;
                    }                
                }
            }
#endregion

#region Scanned Line
            else if (dataGridView2.Rows.Count != 0)
            {
                 // IF NUMBER UP != 1 SHOW GANG PANEL AND HIDE SHEETS AFFECTED BOX
                 //if (numberUp != 1)
                 if ((int)dataGridView2.Rows[0].Cells[21].Value > 1)
                 {
                        tbxSheetsAffectedBadSection.Visible = false;
                        lblSheets_Affected.Visible = false;

                        flowLayoutPanel2.Visible = true;
                        btnScrollUp.Visible = true;
                        btnScrollDown.Visible = true;
                        pnlBadSectionGangHeader.Visible = true;
                        lblStockCode.Visible = true;
                        lblNumberUp.Visible = true;
                        lblNumberUpBadQty.Visible = true;
                        lblSheetsAffected.Visible = true;




                        //DATAGRIDVIEW2
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
                        }

                    // if not equal to 1 then NumberUp =:
                    numberUp = Convert.ToInt32(dataGridView2.Rows[0].Cells[21].Value);

                    //if NumberUp field and JobGanged Field = 0
                    if (numberUp == 0 & Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) == 0)
                    {
                        // Parse the NumberUp value from Section Name or Expr1 - This regex finds continuous digits before "up"
                        if (dataGridView2.Rows[0].Cells[20].Value.ToString().Contains("up"))
                        {
                            String text = dataGridView2.Rows[0].Cells[20].Value.ToString();
                            foreach (Match match in Regex.Matches(text, @"(\d+)up"))
                            {
                                //MessageBox.Show(match.Groups[1].Value);
                                numberUp = Convert.ToInt32(match.Groups[1].Value);
                            }
                        }
                        else
                        {
                            String text1 = dataGridView2.Rows[0].Cells[24].Value.ToString();
                            foreach (Match match in Regex.Matches(text1, @"(\d+)up"))
                            {
                                //MessageBox.Show(match.Groups[1].Value);
                                numberUp = Convert.ToInt32(match.Groups[1].Value);
                            }
                        }
                    }

                    qtyRequired = Convert.ToInt32(dataGridView2.Rows[0].Cells[34].Value);

                    // IF NUMBER UP != 1 SHOW GANG PANEL AND HIDE SHEETS AFFECTED BOX
                    //if (numberUp != 1)
                    if (numberUp > 1)
                    {
                        tbxSheetsAffectedBadSection.Visible = false;
                        lblSheets_Affected.Visible = false;

                        // Filter for Ganged jobs
                        if (Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) != 0 & Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) != 2)
                        {
                            if (Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) == 1)
                            {
                                queryGangClassic();
                            }
                            if (Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) == 3)
                            {
                                queryGangpro();
                            }
                        }

                            // if JobGanged = 0 (Main Table)
                            if (Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) == 0)
                            {
                                if (!badSectionLbls)
                                {
                                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                                    {
                                        btnScrollUp.Visible = false;
                                        btnScrollDown.Visible = false;
                                        lblStockCode.Text = "Job\r\nNumber";
                                        lblStockCode.Visible = true;
                                        lblNumberUp.Visible = true;
                                        lblNumberUp.Text = "Number\r\nUp";
                                        lblNumberUpBadQty.Visible = true;
                                        lblNumberUpBadQty.Text = "Bad\r\nStations";
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
                                        lbl2.Width = 150;
                                        lbl2.BackColor = Color.Silver;
                                        lbl2.Font = new System.Drawing.Font("Microsoft Sans Serif", 20);
                                        lbl2.TextAlign = ContentAlignment.MiddleLeft;
                                        lbl2.ForeColor = Color.Black;
                                        lbl2.Margin = new Padding(0, 0, 0, 0);
                                        lbl2.Left = 40;
                                        lbl2.Text = this.dataGridView2.Rows[i].Cells[3].Value.ToString();
                                        Label lbl3 = new Label();
                                        this.flowLayoutPanel2.Controls.Add(lbl3);
                                        lbl3.Height = 40;
                                        lbl3.Width = 78;
                                        lbl3.BackColor = Color.Silver;
                                        lbl3.Font = new System.Drawing.Font("Microsoft Sans Serif", 20);
                                        lbl3.TextAlign = ContentAlignment.MiddleCenter;
                                        lbl3.ForeColor = Color.Black;
                                        lbl3.Margin = new Padding(0, 0, 0, 0);
                                        lbl3.Left = 40;
                                        lbl3.Text = this.dataGridView2.Rows[i].Cells[21].Value.ToString();
                                        if (dataGridView2.Rows[i].Cells[21].Value.ToString() == "0")
                                        {
                                            lbl3.Text = numberUp.ToString();
                                        }
                                        TextBox textBox1 = new TextBox();
                                        this.flowLayoutPanel2.Controls.Add(textBox1);
                                        textBox1.Height = 40;
                                        textBox1.AutoSize = false;
                                        textBox1.Width = 70;
                                        textBox1.Multiline = false;
                                        textBox1.Font = new System.Drawing.Font(textBox1.Font.FontFamily, 20);
                                        textBox1.TextAlign = HorizontalAlignment.Center;
                                        textBox1.Margin = new Padding(0, 0, 0, 0);
                                        textBox1.TextChanged += new System.EventHandler(this.notGangedNumberUpBad);
                                        TextBox textBox2 = new TextBox();
                                        this.flowLayoutPanel2.Controls.Add(textBox2);
                                        textBox2.Height = 40;
                                        textBox2.AutoSize = false;
                                        textBox2.Width = 85;
                                        textBox2.Multiline = false;
                                        textBox2.Font = new System.Drawing.Font(textBox2.Font.FontFamily, 20);
                                        textBox2.TextAlign = HorizontalAlignment.Center;
                                        textBox2.Margin = new Padding(0, 0, 0, 0);
                                        textBox2.TextChanged += new System.EventHandler(notgangedSheetsAffected);
                                        Button btn1 = new Button();
                                        flowLayoutPanel2.Controls.Add(btn1);
                                        btn1.Height = 40;
                                        btn1.Width = 90;
                                        btn1.BackColor = Color.SteelBlue;
                                        btn1.ForeColor = Color.White;
                                        btn1.Font = new System.Drawing.Font(textBox1.Font.FontFamily, 9);
                                        btn1.Text = "Whole Pallet";
                                        btn1.Margin = new Padding(0, 0, 0, 0);
                                        btn1.Tag = i;
                                        btn1.Click += new System.EventHandler(notGangedWholePallet);
                                    }
                                }
                                badSectionLbls = true;
                            }

                            //if JobGanged = 1 (PALLET_GANG_CLASSIC Table)
                            else if (Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) == 1)
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
                                        lblNumberUpBadQty.Text = "Bad\r\nStations";
                                        lblSheetsAffected.Visible = true;
                                        lblSheetsAffected.Text = "Sheets\r\nAffected";
                                        flowLayoutPanel2.HorizontalScroll.Visible = false;
                                        flowLayoutPanel2.VerticalScroll.Visible = false;
                                        Label lbl1 = new Label();
                                        this.flowLayoutPanel2.Controls.Add(lbl1);
                                        lbl1.Height = 35;
                                        lbl1.Width = 430;
                                        lbl1.BackColor = Color.Gray;
                                        lbl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12);
                                        lbl1.TextAlign = ContentAlignment.MiddleLeft;
                                        lbl1.ForeColor = Color.White;
                                        lbl1.Left = 40;
                                        lbl1.Text = this.dataGridView4.Rows[i].Cells[3].Value.ToString();
                                        Label lbl2 = new Label();
                                        this.flowLayoutPanel2.Controls.Add(lbl2);
                                        lbl2.Height = 40;
                                        lbl2.Width = 150;
                                        lbl2.BackColor = Color.Silver;
                                        lbl2.Font = new System.Drawing.Font("Microsoft Sans Serif", 20);
                                        lbl2.TextAlign = ContentAlignment.MiddleLeft;
                                        lbl2.ForeColor = Color.Black;
                                        lbl2.Margin = new Padding(0, 0, 0, 0);
                                        lbl2.Left = 40;
                                        lbl2.Text = this.dataGridView4.Rows[i].Cells[1].Value.ToString();
                                        Label lbl3 = new Label();
                                        this.flowLayoutPanel2.Controls.Add(lbl3);
                                        lbl3.Height = 40;
                                        lbl3.Width = 78;
                                        lbl3.BackColor = Color.Silver;
                                        lbl3.Font = new System.Drawing.Font("Microsoft Sans Serif", 20);
                                        lbl3.TextAlign = ContentAlignment.MiddleCenter;
                                        lbl3.ForeColor = Color.Black;
                                        lbl3.Margin = new Padding(0, 0, 0, 0);
                                        lbl3.Left = 40;
                                        lbl3.Text = this.dataGridView4.Rows[i].Cells[11].Value.ToString();
                                        if (dataGridView4.Rows[i].Cells[12].Value.ToString() == "0")
                                        {
                                            lbl3.Text = numberUp.ToString();
                                        }
                                        TextBox textBox1 = new TextBox();
                                        this.flowLayoutPanel2.Controls.Add(textBox1);
                                        textBox1.Height = 40;
                                        textBox1.AutoSize = false;
                                        textBox1.Width = 70;
                                        textBox1.Multiline = false;
                                        textBox1.Font = new System.Drawing.Font(textBox1.Font.FontFamily, 20);
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
                                        textBox2.Font = new System.Drawing.Font(textBox2.Font.FontFamily, 20);
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
                                        btn1.Font = new System.Drawing.Font(textBox1.Font.FontFamily, 9);
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

                            //if JobGanged = 3 (PALLET_GANG_CLASSIC Table)
                            else if (Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) == 3)
                            {
                                if (!badSectionLbls)
                                {
                                    for (int i = 0; i < dataGridView3.Rows.Count; i++)
                                    {
                                        lblStockCode.Text = "Stock Code/\r\nJob Number";
                                        lblStockCode.Visible = true;
                                        lblNumberUp.Visible = true;
                                        lblNumberUp.Text = "Number\r\nUp";
                                        lblNumberUpBadQty.Visible = true;
                                        lblNumberUpBadQty.Text = "Bad\r\nStations";
                                        lblSheetsAffected.Visible = true;
                                        lblSheetsAffected.Text = "Sheets\r\nAffected";
                                        flowLayoutPanel2.HorizontalScroll.Visible = false;
                                        flowLayoutPanel2.VerticalScroll.Visible = false;
                                        Label lbl1 = new Label();
                                        this.flowLayoutPanel2.Controls.Add(lbl1);
                                        lbl1.Height = 35;
                                        lbl1.Width = 430;
                                        lbl1.BackColor = Color.Gray;
                                        lbl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12);
                                        lbl1.TextAlign = ContentAlignment.MiddleLeft;
                                        lbl1.ForeColor = Color.White;
                                        lbl1.Left = 40;
                                        lbl1.Text = this.dataGridView3.Rows[i].Cells[4].Value.ToString();
                                        Label lbl2 = new Label();
                                        this.flowLayoutPanel2.Controls.Add(lbl2);
                                        lbl2.Height = 40;
                                        lbl2.Width = 150;
                                        lbl2.BackColor = Color.Silver;
                                        lbl2.Font = new System.Drawing.Font("Microsoft Sans Serif", 20);
                                        lbl2.TextAlign = ContentAlignment.MiddleLeft;
                                        lbl2.ForeColor = Color.Black;
                                        lbl2.Margin = new Padding(0, 0, 0, 0);
                                        lbl2.Left = 40;
                                        lbl2.Text = this.dataGridView3.Rows[i].Cells[1].Value.ToString();
                                        Label lbl3 = new Label();
                                        this.flowLayoutPanel2.Controls.Add(lbl3);
                                        lbl3.Height = 40;
                                        lbl3.Width = 78;
                                        lbl3.BackColor = Color.Silver;
                                        lbl3.Font = new System.Drawing.Font("Microsoft Sans Serif", 20);
                                        lbl3.TextAlign = ContentAlignment.MiddleCenter;
                                        lbl3.ForeColor = Color.Black;
                                        lbl3.Margin = new Padding(0, 0, 0, 0);
                                        lbl3.Left = 40;
                                        lbl3.Text = this.dataGridView3.Rows[i].Cells[3].Value.ToString();
                                        if (dataGridView3.Rows[i].Cells[12].Value.ToString() == "0")
                                        {
                                            lbl3.Text = numberUp.ToString();
                                        }
                                        TextBox textBox1 = new TextBox();
                                        this.flowLayoutPanel2.Controls.Add(textBox1);
                                        textBox1.Height = 40;
                                        textBox1.AutoSize = false;
                                        textBox1.Width = 70;
                                        textBox1.Multiline = false;
                                        textBox1.Font = new System.Drawing.Font(textBox1.Font.FontFamily, 20);
                                        textBox1.TextAlign = HorizontalAlignment.Center;
                                        textBox1.Margin = new Padding(0, 0, 0, 0);
                                        textBox1.Tag = i;
                                        textBox1.TextChanged += new System.EventHandler(this.gangProNumberUpBad);
                                        TextBox textBox2 = new TextBox();
                                        this.flowLayoutPanel2.Controls.Add(textBox2);
                                        textBox2.Height = 40;
                                        textBox2.AutoSize = false;
                                        textBox2.Width = 85;
                                        textBox2.Multiline = false;
                                        textBox2.Font = new System.Drawing.Font(textBox2.Font.FontFamily, 20);
                                        textBox2.TextAlign = HorizontalAlignment.Center;
                                        textBox2.Margin = new Padding(0, 0, 0, 0);
                                        textBox2.Tag = i;
                                        textBox2.TextChanged += new System.EventHandler(gangProSheetsAffected);
                                        Button btn1 = new Button();
                                        flowLayoutPanel2.Controls.Add(btn1);
                                        btn1.Height = 40;
                                        btn1.Width = 73;
                                        btn1.BackColor = Color.SteelBlue;
                                        btn1.ForeColor = Color.White;
                                        btn1.Font = new System.Drawing.Font(textBox1.Font.FontFamily, 9);
                                        btn1.Text = "Whole Pallet";
                                        btn1.Margin = new Padding(0, 0, 0, 0);
                                        btn1.Tag = i;
                                        btn1.Click += new System.EventHandler(this.gangProWholePallet);
                                        numberBadList.Insert(i, "0");
                                        sheetsAffectedList.Insert(i, "0");
                                        wholePalletList.Insert(i, 0);
                                    }
                                }
                            }
                       badSectionLbls = true;
                    }
                }
            }

#endregion

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

            if (badQty == null)
            {
                MessageBox.Show("Please enter a value in the Bad Stations box");
            }
            else
            {
                sheetsAffected = Convert.ToString(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
                notGanged();
                Button button = sender as Button;
                button.BackColor = Color.Gray;
            }
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
            queryGangClassic();
            Button button = sender as Button;
            button.BackColor = Color.Gray;
        }

        private void gangProNumberUpBad(Object sender, EventArgs e)
        {
            badQty = ((TextBox)sender).Text;
            gangRow = (int)((TextBox)sender).Tag;

            if (badQty == "")
            {
                badQty = "0";
            }

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
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

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
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
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
                if (gangRow == i)
                {
                    wholePalletList[i] = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
                }
            queryGangpro();
            Button button = sender as Button;
            button.BackColor = Color.Gray;
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
            pnlPalletCard4.BringToFront();
            index = 11;

            var previousValue = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
            // update label with new qty
            lbl5.Text = Convert.ToString(previousValue - sheetsAffectedBadSection);
            btnBadSectionOK.Visible = true;

            // WorkingSize
            lbl6.Text = dataGridView1.Rows[0].Cells[13].Value.ToString();
            // QtyRequired
            lbl7.Text = dataGridView1.Rows[0].Cells[25].Value.ToString();
        }

        private void btnExtraInformationPalletCard_Click(object sender, EventArgs e)
        {
            pnlPalletCard7.BringToFront();
            index = 14;
        }

#endregion

#region Finish Pallet or Cancel

        private void btnFinishPalletContinue_Click(object sender, EventArgs e)
        {
            if (dataGridView2.RowCount == 0)
            {
                int account = Convert.ToInt32(dataGridView1.Rows[0].Cells[21].Value);
                //if (account == 1784 || account == 1781 || account == 1700 || account == 1795 || account == 1720 || account == 1839 || account == 1888)
                if (account == 1720 || account == 1839)
                    {
                    pnlPalletCard10.BringToFront();
                    // Disable the Section button
                    disableSectionButtons.Add(Convert.ToString(dataGridView1.Rows[0].Cells[19].Value));
                    removeFlowLayoutBtns();
                    sigBtns = false;
                    btnCancel.Visible = false;
                }
                else if (account == 1854 || account == 1305 || account == 1337 || account == 1435 || account == 1263 || account == 1092 || account == 1385 || account == 1598 || account == 1686 || account == 1726 || account == 1921 || account == 1458)
                {
                    pnlPalletCard11.BringToFront();
                    // Disable the Section button
                    disableSectionButtons.Add(Convert.ToString(dataGridView1.Rows[0].Cells[19].Value));
                    removeFlowLayoutBtns();
                    sigBtns = false;
                    btnCancel.Visible = false;
                }
                else
                {
                    pnlSignature.BringToFront();
                    // Disable the Section button
                    disableSectionButtons.Add(Convert.ToString(dataGridView1.Rows[0].Cells[19].Value));
                    removeFlowLayoutBtns();
                    sigBtns = false;
                    btnCancel.Visible = false;
                }
            }
            else if (dataGridView2.RowCount != 0)
            {
                int account = Convert.ToInt32(dataGridView2.Rows[0].Cells[29].Value);
                //if (account == 1784 || account == 1781 || account == 1700 || account == 1795 || account == 1720 || account == 1839 || account == 1888)
                    if (account == 1720 || account == 1839)
                {
                    pnlPalletCard10.BringToFront();
                    // Disable the Section button
                    disableSectionButtons.Add(Convert.ToString(dataGridView2.Rows[0].Cells[8].Value));
                    removeFlowLayoutBtns();
                    sigBtns = false;
                    btnCancel.Visible = false;
                }
                else if (account == 1854 || account == 1305 || account == 1337 || account == 1435 || account == 1263 || account == 1092 || account == 1385 || account == 1598 || account == 1686 || account == 1726 || account == 1921 || account == 1458)
                {
                    pnlPalletCard11.BringToFront();
                    // Disable the Section button
                    disableSectionButtons.Add(Convert.ToString(dataGridView2.Rows[0].Cells[8].Value));
                    removeFlowLayoutBtns();
                    sigBtns = false;
                    btnCancel.Visible = false;
                }
                else
                {
                    pnlSignature.BringToFront();
                    // Disable the Section button
                    disableSectionButtons.Add(Convert.ToString(dataGridView1.Rows[0].Cells[19].Value));
                    removeFlowLayoutBtns();
                    sigBtns = false;
                    btnCancel.Visible = false;
                }
            }
            flowLayoutPanel2.Enabled = true;
            tbxSheetsAffectedBadSection.Enabled = true;
            tbxSheetsAffectedBadSection.Text = "";
            badSectionLbls = false;
            btnBadSectionOK.Visible = false;
            numberBadList.Clear();
            sheetsAffectedList.Clear();
            numberUp = 0;
        }

        private void btnPalletFinished_Click(object sender, EventArgs e)
        {
            pnlPalletCard8.BringToFront();

            if(dataGridView2.Rows.Count != 0)
            {
                // WorkingSize
                lbl6.Text = dataGridView2.Rows[0].Cells[22].Value.ToString();
                // QtyRequired
                lbl7.Text = dataGridView2.Rows[0].Cells[34].Value.ToString();

                btnIsSectionFinishedYes_Click(btnIsSectionFinishedYes, EventArgs.Empty);
                index = 15;
            }
            else
            { 
                if (dataGridView1.Rows[0].Cells[15].Value.ToString() == "")
                {
                    lblIsSectionFinished.Text = dataGridView1.Rows[0].Cells[11].Value.ToString() + "\r\n" + "Section " + dataGridView1.Rows[0].Cells[19].Value.ToString();
                }
                else
                {
                    lblIsSectionFinished.Text = dataGridView1.Rows[0].Cells[15].Value.ToString() + "\r\n" + "Section " + dataGridView1.Rows[0].Cells[19].Value.ToString();
                }

                // WorkingSize
                lbl6.Text = dataGridView1.Rows[0].Cells[13].Value.ToString();
                // QtyRequired
                lbl7.Text = dataGridView1.Rows[0].Cells[25].Value.ToString();

                index = 15;
            }
        }

        private void btnCancelPrintMore_Click(object sender, EventArgs e)
        {
            string ConnectionString1 = Convert.ToString("Dsn=PalletCard;uid=PalletCardAdmin");
            string CommandText1 = "SELECT * FROM Log where JobNo = '" + lblJobNo.Text + "'";
            OdbcConnection myConnection1 = new OdbcConnection(ConnectionString1);
            OdbcCommand myCommand1 = new OdbcCommand(CommandText1, myConnection1);
            OdbcDataAdapter myAdapter1 = new OdbcDataAdapter();
            myAdapter1.SelectCommand = myCommand1;
            DataSet palletCardData1 = new DataSet();
            try
            {
                myConnection1.Open();
                myAdapter1.Fill(palletCardData1);
            }
            catch (Exception ex)
            {
                throw (ex);
            }

            finally
            {
                myConnection1.Close();
            }
            using (DataTable palletCardLog1 = new DataTable())
            {

                myAdapter1.Fill(palletCardLog1);
                dataGridView5.DataSource = palletCardLog1;
            }            
            // set variables from datagridview5
            this.dataGridView5.Sort(this.dataGridView5.Columns["AutoNum"], ListSortDirection.Descending);
            int autoNum = Convert.ToInt32(dataGridView5.Rows[0].Cells[0].Value);
            int rollBackPalletNum = Convert.ToInt32(dataGridView5.Rows[0].Cells[4].Value) - 1;

            // Update datagridview2 - Active selection from FinishPallet
            string ConnectionString = Convert.ToString("Dsn=PalletCard;uid=PalletCardAdmin");
            string CommandText = "Update Log set JobCancelled = 1, Unfinished = 1, PalletNumber = '" + rollBackPalletNum + "' where AutoNum = '" + autoNum + "' ";
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
            index = 10;
            lbl4.Visible = false;
            lbl5.Visible = false;
            tbxPalletHeightPalletCard.Text = "";
            lblSheetCountPalletCard.Text = "";
            tbxSheetCountPalletCard.Text = "";
            lblPheightPalletCard.Text = "";
            btnPalletCard_Click(btnPalletCard, EventArgs.Empty);
            badSectionLbls = false;
            cancelPrintMoreClicked = true;
            btnMarkBad.Enabled = true;
            btnMarkBad.BackColor = Color.SteelBlue;
            flowLayoutPanel2.Enabled = true;
            tbxSheetsAffectedBadSection.Enabled = true;
            tbxSheetsAffectedBadSection.Text = "";
            badSectionLbls = false;
            btnBadSectionOK.Visible = false;
            numberBadList.Clear();
            sheetsAffectedList.Clear();
            tbxBadSectionComment.Text = "";
            tbxExtraInfoComment.Text = "";
            numberUp = 0;
        }

#endregion

#region Backup/Varnish
        private void btnBackupRequired_Click(object sender, EventArgs e)
        {
            // Regular line
            if (dataGridView2.RowCount == 0)
            {
                //SAVE TO DATABASE
                produced = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
                string sqlFormattedDate = CurrentDate.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
                string Query = "insert into Log (Routine, JobNo, PaperSectionNo, PalletNumber, Produced, Expr1, NumberUp, JobGanged, SectionName, JobDesc, QtyRequired, ResourceID, Description, WorkingSize, SheetQty, Comment, Unfinished, Timestamp1, InvoiceCustomerCode, BadSectionComment) values('" + lbl1.Text + "','" + dataGridView1.Rows[0].Cells[0].Value + "','" + dataGridView1.Rows[0].Cells[19].Value + "', '1', '" + produced + "','" + lbl2.Text + "', '" + numberUp + "', '" + dataGridView1.Rows[0].Cells[14].Value + "','" + lbl2.Text + "', '" + dataGridView1.Rows[0].Cells[18].Value + "', '" + dataGridView1.Rows[0].Cells[25].Value + "','" + resourceID + "','" + dataGridView1.Rows[0].Cells[16].Value + "','" + dataGridView1.Rows[0].Cells[13].Value + "','" + lbl5.Text + "','" + tbxExtraInfoComment.Text + "','2','" + CurrentDate + "', '" + dataGridView1.Rows[0].Cells[21].Value + "', '" + tbxBadSectionComment.Text + "');";
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
                finally
                {
                    conDatabase.Close();
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
            }

            // Scanned Line
            else if (dataGridView2.RowCount != 0)
            {
                //SAVE TO DATABASE
                produced = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
                string sqlFormattedDate = CurrentDate.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
                string Query = "insert into Log (Routine, JobNo, PaperSectionNo, PalletNumber, Produced, Expr1, NumberUp, JobGanged, SectionName, JobDesc, QtyRequired, ResourceID, Description, WorkingSize, SheetQty, Comment, Unfinished, Timestamp1, InvoiceCustomerCode, BadSectionComment) values('" + this.lbl1.Text + "','" + this.dataGridView2.Rows[0].Cells[3].Value + "','" + this.dataGridView2.Rows[0].Cells[8].Value + "', '" + this.dataGridView2.Rows[0].Cells[4].Value + "', '" + produced + "', '" + this.dataGridView2.Rows[0].Cells[20].Value + "', '" + numberUp + "', '" + this.dataGridView2.Rows[0].Cells[23].Value + "', '" + this.dataGridView2.Rows[0].Cells[24].Value + "', '" + this.dataGridView2.Rows[0].Cells[27].Value + "', '" + this.dataGridView2.Rows[0].Cells[34].Value + "','" + resourceID + "','" + dataGridView2.Rows[0].Cells[25].Value + "','" + this.dataGridView2.Rows[0].Cells[22].Value + "','" + this.lbl5.Text + "','" + this.tbxExtraInfoComment.Text + "','2','" + CurrentDate + "', '" + this.dataGridView2.Rows[0].Cells[29].Value + "', '" + tbxBadSectionComment.Text + "');";
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
                finally
                {
                    conDatabase.Close();
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
            }

            this.dataGridView2.Sort(this.dataGridView2.Columns["AutoNum"], ListSortDirection.Descending);

            autoNum = Convert.ToInt32(dataGridView2.Rows[0].Cells[0].Value);

            pnlPalletCardPrint.BringToFront();

            btnBack.Visible = false;
            btnCancel.Visible = false;
            btnPalletCardPrint.Visible = true;
            sectionFinishedClicked = true;
            tbxSheetsAffectedBadSection.Visible = true;
            btnMarkBad.Enabled = true;
            btnMarkBad.BackColor = Color.SteelBlue;
            index = 17;
        }

        private void btnVarnishRequired_Click(object sender, EventArgs e)
        {
            // Regular line
            if(dataGridView2.RowCount == 0)
            {
                //SAVE TO DATABASE
                produced = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
                string sqlFormattedDate = CurrentDate.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
                string Query = "insert into Log (Routine, JobNo, PaperSectionNo, PalletNumber, Produced, Expr1, NumberUp, JobGanged, SectionName, JobDesc, QtyRequired, ResourceID, Description, WorkingSize, SheetQty, Comment, Unfinished, Timestamp1, InvoiceCustomerCode, BadSectionComment) values('" + lbl1.Text + "','" + dataGridView1.Rows[0].Cells[0].Value + "','" + dataGridView1.Rows[0].Cells[19].Value + "', '1', '" + produced + "', '" + lbl2.Text + "', '" + numberUp + "', '" + dataGridView1.Rows[0].Cells[14].Value + "', '" + lbl2.Text + "', '" + dataGridView1.Rows[0].Cells[18].Value + "', '" + dataGridView1.Rows[0].Cells[25].Value + "', '" + resourceID + "','" + dataGridView1.Rows[0].Cells[16].Value + "','" + dataGridView1.Rows[0].Cells[13].Value + "','" + lbl5.Text + "','" + tbxExtraInfoComment.Text + "','3','" + CurrentDate + "', '" + dataGridView1.Rows[0].Cells[21].Value + "', '" + tbxBadSectionComment.Text + "');";
                SqlConnection conDatabase = new SqlConnection(constring);
                SqlCommand cmdDatabase = new SqlCommand(Query, conDatabase);
                SqlDataReader myReader;
                try
                {
                    conDatabase.Open();
                    myReader = cmdDatabase.ExecuteReader();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    conDatabase.Close();
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
            }

            // Scanned Line
            else if (dataGridView2.RowCount != 0)
            {
                //SAVE TO DATABASE
                produced = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
                string sqlFormattedDate = CurrentDate.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
                string Query = "insert into Log (Routine, JobNo, PaperSectionNo, PalletNumber, Produced, Expr1, NumberUp, JobGanged, SectionName, JobDesc, QtyRequired, ResourceID, Description, WorkingSize, SheetQty, Comment, Unfinished, Timestamp1, InvoiceCustomerCode, BadSectionComment) values('" + this.lbl1.Text + "','" + this.dataGridView2.Rows[0].Cells[3].Value + "','" + this.dataGridView2.Rows[0].Cells[8].Value + "', '" + this.dataGridView2.Rows[0].Cells[4].Value + "', '" + produced + "', '" + this.dataGridView2.Rows[0].Cells[20].Value + "', '" + numberUp + "', '" + this.dataGridView2.Rows[0].Cells[23].Value + "', '" + this.dataGridView2.Rows[0].Cells[24].Value + "', '" + this.dataGridView2.Rows[0].Cells[27].Value + "', '" + this.dataGridView2.Rows[0].Cells[34].Value + "', '" + resourceID + "','" + dataGridView2.Rows[0].Cells[25].Value + "','" + this.dataGridView2.Rows[0].Cells[22].Value + "','" + this.lbl5.Text + "','" + this.tbxExtraInfoComment.Text + "','3','" + CurrentDate + "', '" + this.dataGridView2.Rows[0].Cells[29].Value + "', '" + tbxBadSectionComment.Text + "');";
                SqlConnection conDatabase = new SqlConnection(constring);
                SqlCommand cmdDatabase = new SqlCommand(Query, conDatabase);
                SqlDataReader myReader;
                try
                {
                    conDatabase.Open();
                    myReader = cmdDatabase.ExecuteReader();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    conDatabase.Close();
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
            }

            this.dataGridView2.Sort(this.dataGridView2.Columns["AutoNum"], ListSortDirection.Descending);

            autoNum = Convert.ToInt32(dataGridView2.Rows[0].Cells[0].Value);

            pnlPalletCardPrint.BringToFront();
            btnBack.Visible = false;
            btnCancel.Visible = false;
            sectionFinishedClicked = true;
            tbxSheetsAffectedBadSection.Visible = true;
            btnMarkBad.Enabled = true;
            btnMarkBad.BackColor = Color.SteelBlue;
            index = 17;
        }
#endregion

#region Check Overs and Print

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
                myConnection.Close();
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
                myConnection.Close();
            }
            using (DataTable palletCardLog = new DataTable())
            {
                myAdapter.Fill(palletCardLog);
                dataGridView2.DataSource = palletCardLog;
            }
            // Filter datagridview2 for scanned line
            if (tbxFinishPallet.Text != "")
            {
                //int x = Convert.ToInt32(tbxFinishPallet.Text);
                //((DataTable)dataGridView2.DataSource).DefaultView.RowFilter = "Convert(AutoNum, 'System.String') = '%" + tbxFinishPallet.Text + "%'";
                string ConnectionString1 = Convert.ToString("Dsn=PalletCard;uid=PalletCardAdmin");
                string CommandText1 = "SELECT * FROM Log where AutoNum = '" + Convert.ToInt32(tbxFinishPallet.Text) + "'";
                OdbcConnection myConnection1 = new OdbcConnection(ConnectionString1);
                OdbcCommand myCommand1 = new OdbcCommand(CommandText1, myConnection1);
                OdbcDataAdapter myAdapter1 = new OdbcDataAdapter();
                myAdapter.SelectCommand = myCommand1;
                DataSet palletCardData1 = new DataSet();
                try
                {
                    myConnection1.Open();
                    myAdapter.Fill(palletCardData1);
                }
                catch (Exception ex)
                {
                    throw (ex);
                }

                finally
                {
                    myConnection1.Close();
                }
                using (DataTable palletCardLog1 = new DataTable())
                {
                    myAdapter.Fill(palletCardLog1);
                    dataGridView2.DataSource = palletCardLog1;
                }
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
                    catch (Exception)
                    {
                    }

                    // If no entries for Pallet Card routine
                    if (dataGridView2.Rows.Count == 0)
                    {
                        PalletNumber = 1;
                    }

                    // Do not increment PalletNumber for a scanned line
                    else if(Convert.ToInt32(dataGridView2.Rows[0].Cells[6].Value) == 2 || Convert.ToInt32(dataGridView2.Rows[0].Cells[6].Value) == 3)
                        {
                            this.dataGridView2.Sort(this.dataGridView2.Columns["PalletNumber"], ListSortDirection.Descending);
                            if (dataGridView2.RowCount == 0)
                            {
                                PalletNumber = 0;
                            }
                            PalletNumber = (int)dataGridView2.Rows[0].Cells[4].Value;
                        }

                        // Increment PalletNumber for a non-scanned line
                        else
                        {
                            dataGridView2.AllowUserToAddRows = false;
                            this.dataGridView2.Sort(this.dataGridView2.Columns["PalletNumber"], ListSortDirection.Descending);
                            if (dataGridView2.RowCount == 0)
                            {
                               PalletNumber = 1;
                            }
                            else
                            {
                                PalletNumber = (int)dataGridView2.Rows[0].Cells[4].Value + 1;
                            }
                        }
                    dataGridView2.AllowUserToAddRows = true;
                }

            // Get the quantities produced from the previous pallet cards
            sumProduced = 0;

            // ignore if the value has been retrieved by scan
            this.dataGridView2.Sort(this.dataGridView2.Columns["AutoNum"], ListSortDirection.Descending);
            PaperSectionNo = Convert.ToInt32(dataGridView2.Rows[0].Cells[8].Value);

            // If unfinished: 
            // (unfinished = 1, Backup required = 2, Varnish required = 3
            // Finished = 0. If goes through routine only once with IsSectionFinishedYes chosen then unfinished is given a value of 0.
            // IsSectionFinishedNo gives Unfinished a value of 1 to indicate that these lines need to sum up the qty Produced values.
            // Also last time it goes through (IsSectionFinishedYes) Unfinished is given value of 0
            // It wont sum up values where Unfinished = 0

            // filter to current Section Name - there may be other Section Names with the same Paper Section Number
            if (dataGridView2.Rows[0].Cells[15] == null)
            {
                ((DataTable)dataGridView2.DataSource).DefaultView.RowFilter = "Expr1 like '%" + lbl2.Text + "%'";
            }
            else
            {
                ((DataTable)dataGridView2.DataSource).DefaultView.RowFilter = "SectionName like '%" + lbl2.Text + "%'";
            }

            if (Convert.ToInt32(dataGridView2.Rows[0].Cells[6].Value) != 0)
            {
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    dataGridView2.AllowUserToAddRows = false;
                    // Sum up only PaperSectionNo's associated with this section not all PaperSectionNo's AND also not canceled
                    if (Convert.ToInt32(dataGridView2.Rows[i].Cells[8].Value) == PaperSectionNo && Convert.ToInt32(dataGridView2.Rows[i].Cells[33].Value) != 1)
                    {
                        sumProduced += Convert.ToInt32(dataGridView2.Rows[i].Cells[9].Value);
                    }
                    if (Convert.ToInt32(dataGridView2.Rows[i].Cells[6].Value) == 2 || Convert.ToInt32(dataGridView2.Rows[i].Cells[6].Value) == 3)
                    {
                        sumProduced = 0;
                    }
                }
            }

            if (cancelPrintMoreClicked == true)
            {
                sheetsAffectedBadSection = 0;
                cancelPrintMoreClicked = false;
            }

            //SAVE TO DATABASE
            CurrentDate = DateTime.Now;
            //produced = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", "")) - sheetsAffectedBadSection;
            produced = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
            PaperSectionNo = Convert.ToInt32(Regex.Replace(lbl3.Text, "[^0-9.]", ""));
            lastPallet = 1;

            string sqlFormattedDate = CurrentDate.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
            string Query;
            if (tbxFinishPallet.Text != "")
            {
                Query = "insert into Log (Routine, JobNo, PalletNumber, Unfinished, PaperSectionNo, NumberUp, JobGanged, JobDesc, QtyRequired, ResourceID, WorkingSize, Description, SheetQty, Comment, Timestamp1, LastPallet, Produced, Expr1, SectionName, InvoiceCustomerCode, InkBatch, PaperBatch, BadSectionComment) values('" + lbl1.Text + "','" + lblJobNo.Text + "','" + PalletNumber + "', '0','" + PaperSectionNo + "', '" + dataGridView2.Rows[0].Cells[21].Value + "', '" + dataGridView2.Rows[0].Cells[23].Value + "', '" + dataGridView2.Rows[0].Cells[27].Value + "', '" + lbl7.Text + "','" + resourceID + "','" + lbl6.Text + "','" + dataGridView2.Rows[0].Cells[25].Value + "','" + lbl5.Text + "','" + tbxExtraInfoComment.Text + "','" + CurrentDate + "','" + lastPallet + "','" + produced + "','" + lbl2.Text + "','" + lbl2.Text + "', '" + this.dataGridView2.Rows[0].Cells[29].Value + "', '" + inkDetails + "', '" + paperDetails + "', '" + tbxBadSectionComment.Text + "');";
            }
            else
            {
                Query = "insert into Log (Routine, JobNo, PalletNumber, Unfinished, PaperSectionNo, NumberUp, JobGanged, JobDesc, QtyRequired, ResourceID, WorkingSize, Description, SheetQty, Comment, Timestamp1, LastPallet, Produced, Expr1, SectionName, InvoiceCustomerCode, InkBatch, PaperBatch, BadSectionComment) values('" + lbl1.Text + "','" + lblJobNo.Text + "','" + PalletNumber + "', '0','" + PaperSectionNo + "', '" + numberUp + "', '" + dataGridView1.Rows[0].Cells[14].Value + "', '" + dataGridView1.Rows[0].Cells[18].Value + "', '" + lbl7.Text + "','" + resourceID + "','" + lbl6.Text + "','" + this.dataGridView1.Rows[0].Cells[16].Value + "','" + lbl5.Text + "','" + tbxExtraInfoComment.Text + "','" + CurrentDate + "','" + lastPallet + "','" + produced + "','" + lbl2.Text + "','" + lbl2.Text + "', '" + this.dataGridView1.Rows[0].Cells[21].Value + "', '" + inkDetails + "', '" + paperDetails + "', '" + tbxBadSectionComment.Text + "');";
            }
            SqlConnection conDatabase = new SqlConnection(constring);
            SqlCommand cmdDatabase = new SqlCommand(Query, conDatabase);
            SqlDataReader myReader;
            try
            {
                conDatabase.Open();
                myReader = cmdDatabase.ExecuteReader();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conDatabase.Close();
            }

            reQueryDataGridView2();

            required = Convert.ToInt32(lbl7.Text);
            //produced = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", "")) - sheetsAffectedBadSection + sumProduced;
            produced = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", "")) + sumProduced;
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
                if (produced < required)
                {
                    pnlPalletCard6.BringToFront();
                    btnBack.Visible = false;
                    lblPalletDidNotMakeQty.Text = "Job " + lblJobNo.Text + " Sheet " + dataGridView1.Rows[0].Cells[19].Value.ToString() + " has " + shortBy + " insufficient sheets";

                    lblUnfinishedSection.Visible = false;

                    // filter on Section Name
                    if (dataGridView2.Rows[0].Cells[15] == null)
                    {
                        ((DataTable)dataGridView2.DataSource).DefaultView.RowFilter = "Expr1 like '%" + lbl2.Text + "%'";
                    }
                    else
                    {
                        ((DataTable)dataGridView2.DataSource).DefaultView.RowFilter = "SectionName like '%" + lbl2.Text + "%'";
                    }

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

                    lblUnfinishedSection.Visible = true;
                    lblUnfinishedSection.Text = "";
                    foreach (string s in sectionsNoLastFlag)
                    {
                        lblUnfinishedSection.Text += "The pallet for Section " + s + " is not finished" + "\r\n";
                    }
                    if (lblUnfinishedSection.Text.Length > 0)
                    {
                        lblWarning.Visible = true;
                    }
                    else
                    {
                        lblWarning.Visible = false;
                    }

                    // Send email notification
                    try
                    {
                        MailMessage mail = new MailMessage("Press@colorman.ie", defaultEmail, "Print Job Shortage", "Job Number " + lblJobNo.Text + " - Section " + dataGridView2.Rows[0].Cells[8].Value.ToString() + "- has " + shortBy + " insufficient sheets");
                        SmtpClient client = new SmtpClient("ex0101.ColorMan.local");
                        client.Port = 25;
                        client.EnableSsl = false;
                        client.Send(mail);
                    }
                    catch
                    { }
                }

                else if (produced > oversCalc)
                {
                    pnlPalletCard9.BringToFront();
                    btnBack.Visible = false;
                    lblPalletOverBySheets.Text = lblJobNo.Text + " is over by " + overBy;
                    // Disable the Section button so it can't be pressed again
                    disableSectionButtons.Add(Convert.ToString(dataGridView1.Rows[0].Cells[19].Value));
                    removeFlowLayoutBtns();
                    sigBtns = false;

                    // filter on Section Name
                    if (dataGridView2.Rows[0].Cells[15] == null)
                    {
                        ((DataTable)dataGridView2.DataSource).DefaultView.RowFilter = "Expr1 like '%" + lbl2.Text + "%'";
                    }
                    else
                    {
                        ((DataTable)dataGridView2.DataSource).DefaultView.RowFilter = "SectionName like '%" + lbl2.Text + "%'";
                    }

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

                    lblUnfinishedSection.Visible = true;
                    lblUnfinishedSection.Text = "";

                    string[] unique = sectionsNoLastFlag.Distinct().ToArray();

                    foreach ( string s in unique)
                    {
                        lblUnfinishedSection.Text += "The pallet for Section " + s + " is not finished" + "\r\n";
                    }
                    if (lblUnfinishedSection.Text.Length > 0)
                    {
                        lblWarning.Visible = true;
                    }
                    else
                    {
                        lblWarning.Visible = false;
                    }

                    // Send email notification
                    try
                    {
                        MailMessage mail = new MailMessage("Press@colorman.ie", defaultEmail, "Print Job Over", "Job Number " + lblJobNo.Text + " - Section " + dataGridView2.Rows[0].Cells[8].Value.ToString() + " - is over by " + overBy);
                        SmtpClient client = new SmtpClient("ex0101.ColorMan.local");
                        client.Port = 25;
                        client.EnableSsl = false;
                        client.Send(mail);
                        index = 16;
                    }
                    catch
                    { }
                }
                else if (produced <= oversCalc)
                {

                    btnBack.Visible = false;
                    // Disable the Section button so it can't be pressed again
                    disableSectionButtons.Add(Convert.ToString(dataGridView1.Rows[0].Cells[19].Value));
                    removeFlowLayoutBtns();
                    sigBtns = false;

                    // filter on Section Name
                    if (dataGridView2.Rows[0].Cells[15] == null)
                    {
                        ((DataTable)dataGridView2.DataSource).DefaultView.RowFilter = "Expr1 like '%" + lbl2.Text + "%'";
                    }
                    else
                    {
                        ((DataTable)dataGridView2.DataSource).DefaultView.RowFilter = "SectionName like '%" + lbl2.Text + "%'";
                    }

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
                        //if (!completedSections.Contains(s) && !completedSections.Contains(Regex.Replace(lbl3.Text, "[^0-9.]", "")) )
                        if (!completedSections.Contains(s))
                            sectionsNoLastFlag.Add(s);
                    }

                    lblUnfinishedSection.Visible = true;
                    lblUnfinishedSection.Text = "";
                    foreach (string s in sectionsNoLastFlag)
                    {
                        lblUnfinishedSection.Text += "The pallet for Section " + s + " is not finished" + "\r\n";
                    }
                    if (lblUnfinishedSection.Text.Length > 0)
                    {
                        lblWarning.Visible = true;
                    }
                    else
                    {
                        lblWarning.Visible = false;
                    }
                        
                    if (dataGridView2.RowCount == 0)
                    {
                        int account = Convert.ToInt32(dataGridView1.Rows[0].Cells[21].Value);
                        //if (account == 1784 || account == 1781 || account == 1700 || account == 1795 || account == 1720 || account == 1839 || account == 1888)
                            if (account == 1720 || account == 1839)
                        {
                            pnlPalletCard10.BringToFront();
                        }
                        else if (account == 1854 || account == 1305 || account == 1337 || account == 1435 || account == 1263 || account == 1092 || account == 1385 || account == 1598 || account == 1686 || account == 1726 || account == 1921 || account == 1458)
                        {
                            pnlPalletCard11.BringToFront();
                        }
                        else
                        {
                            pnlSignature.BringToFront();
                        }
                    }
                    if(dataGridView2.RowCount != 0)
                    {
                        int account = Convert.ToInt32(dataGridView2.Rows[0].Cells[29].Value);
                        //if (account == 1784 || account == 1781 || account == 1700 || account == 1795 || account == 1720 || account == 1839 || account == 1888)
                            if (account == 1720 || account == 1839)
                        {
                            pnlPalletCard10.BringToFront();
                        }
                        else if (account == 1854 || account == 1305 || account == 1337 || account == 1435 || account == 1263 || account == 1092 || account == 1385 || account == 1598 || account == 1686 || account == 1726 || account == 1921 || account == 1458)
                        {
                            pnlPalletCard11.BringToFront();
                        }
                        else
                        {
                            pnlSignature.BringToFront();
                        }
                    }
                }
            }
            this.dataGridView2.Sort(this.dataGridView2.Columns["AutoNum"], ListSortDirection.Descending);
            //autoNum = Convert.ToString((int)dataGridView2.Rows[0].Cells[0].Value + 1);
            autoNum = Convert.ToInt32(dataGridView2.Rows[0].Cells[0].Value);
            btnPalletCardPrint.Visible = true;
            sectionFinishedClicked = true;
            dataGridView2.AllowUserToAddRows = false;
            btnCancel.Visible = false;
            index = 16;
        }

        private void btnSavePosaImage_Click(object sender, EventArgs e)
        {
            SavePosaImageToFile();
            pnlSignature.BringToFront();
        }

        private void btnSaveMedicalPharmaImage_Click(object sender, EventArgs e)
        {
            if (tbxInkDetails.Text != "")
            {
                if (tbxPaperDetails.Text != "")
                {
                    inkDetails = tbxInkDetails.Text;
                    paperDetails = tbxPaperDetails.Text;
                    SaveMedicalPharmaImageToFile();
                    pnlSignature.BringToFront();

                    // add ink paper detail to the database by finding the already saved line and inseting the values
                    string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
                    using (SqlConnection connection = new SqlConnection(constring))
                    using (SqlCommand command1 = connection.CreateCommand())
                    using (SqlCommand command2 = connection.CreateCommand())
                    {
                        command1.CommandText = "UPDATE log SET InkBatch = @ink Where JobNo = @jobNumber and PalletNumber = @palletNumber";
                        command2.CommandText = "UPDATE log SET PaperBatch = @paper Where JobNo = @jobNumber and PalletNumber = @palletNumber";

                        command1.Parameters.AddWithValue("@ink", inkDetails);
                        command1.Parameters.AddWithValue("@jobNumber", lblJobNo.Text);
                        command1.Parameters.AddWithValue("@palletNumber", PalletNumber);
                        command2.Parameters.AddWithValue("@paper", paperDetails);
                        command2.Parameters.AddWithValue("@jobNumber", lblJobNo.Text);
                        command2.Parameters.AddWithValue("@palletNumber", PalletNumber);

                        try
                        {
                            connection.Open();

                            command1.ExecuteNonQuery();
                            command2.ExecuteNonQuery();
                        }
                        finally
                        {
                            connection.Close();
                        }
                    }
                }
            }
            else MessageBox.Show("Please enter both Ink and Paper Details");
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
            dataGridView2.AllowUserToAddRows = false;
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
                                dataGridView2.AllowUserToAddRows = true;
                                this.dataGridView2.Sort(this.dataGridView2.Columns["PalletNumber"], ListSortDirection.Descending);
                                PalletNumber = (int)dataGridView2.Rows[0].Cells[4].Value + 1;
                            }
                                dataGridView2.AllowUserToAddRows = false;
            }

            //SAVE TO DATABASE
            CurrentDate = DateTime.Now;
            //produced = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", "")) - sheetsAffectedBadSection;
            produced = Convert.ToInt32(Regex.Replace(lbl5.Text, "[^0-9.]", ""));
            dataGridView2.Refresh();
            var rowCount = dataGridView2.Rows.Count -1;
            string sqlFormattedDate = CurrentDate.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string constring = "Data Source=APPSHARE01\\SQLEXPRESS01;Initial Catalog=PalletCard;Persist Security Info=True;User ID=PalletCardAdmin;password=Pa!!etCard01";
            string Query = "insert into Log (Routine, JobNo, PalletNumber, PaperSectionNo, NumberUp, JobGanged, JobDesc, QtyRequired, ResourceID, WorkingSize, Description, SheetQty, Comment, Timestamp1, Produced, Unfinished, InvoiceCustomerCode, InkBatch, PaperBatch, Expr1, SectionName, BadSectionComment) values('" + lbl1.Text + "','" + dataGridView1.Rows[0].Cells[0].Value  + "','" + PalletNumber + "','" + dataGridView1.Rows[0].Cells[19].Value + "', '" + numberUp + "', '" + dataGridView1.Rows[0].Cells[14].Value + "', '" + dataGridView1.Rows[0].Cells[18].Value + "', '" + dataGridView1.Rows[0].Cells[25].Value + "','" + resourceID + "','" + dataGridView1.Rows[0].Cells[13].Value + "','" + dataGridView1.Rows[0].Cells[16].Value + "','" + lbl5.Text + "','" + tbxExtraInfoComment.Text + "','" + CurrentDate + "','" + produced + "', '1', '" + dataGridView1.Rows[0].Cells[21].Value + "', '" + inkDetails + "', '" + paperDetails + "','" + lbl2.Text + "','" + lbl2.Text + "', '" + tbxBadSectionComment.Text + "');";
            SqlConnection conDatabase = new SqlConnection(constring);
            SqlCommand cmdDatabase = new SqlCommand(Query, conDatabase);
            SqlDataReader myReader;
            try
            {
                conDatabase.Open();
                myReader = cmdDatabase.ExecuteReader();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conDatabase.Close();
            }

            // Requery the data to refresh dataGridView2 with the newly added PalletNumber and barCode
            reQueryDataGridView2();

            this.dataGridView2.Sort(this.dataGridView2.Columns["AutoNum"], ListSortDirection.Descending);
            autoNum = Convert.ToInt32(dataGridView2.Rows[0].Cells[0].Value);
            pnlPalletCardPrint.BringToFront();
            btnPalletCardPrint.Visible = true;
            btnBack.Visible = false;
            btnCancel.Visible = false;
            index = 17;
        }

        private void btnPalletCardPrint_Click(object sender, EventArgs e)
        {
            lblPrinting.Visible = true;

            flowLayoutPanel2.Enabled = true;
            tbxSheetsAffectedBadSection.Enabled = true;
            tbxSheetsAffectedBadSection.Text = "";
            badSectionLbls = false;
            btnBadSectionOK.Visible = false;
            numberBadList.Clear();
            sheetsAffectedList.Clear();
            numberUp = 0;

            // Test to List Local Printers
            //foreach (string printerName in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            //{
            //    MessageBox.Show(printerName);
            //}

            btnPalletCardPrint.Visible = false;
            btnMarkBad.Enabled = true;
            btnMarkBad.BackColor = Color.SteelBlue;
            PrintImagePalletCard();

            // Absolute path of PDF to print (with filename) - location = \bin\Debug\frontBackCombined.pdf
            string Filepath = AppDomain.CurrentDomain.BaseDirectory + "frontBackCombined.pdf";           

            // The name of the PDF that will be printed (just to be shown in the print queue)
            string Filename = "frontBackCombined.pdf";
            // The name of the printer that you want to use
            string PrinterName = defaultPrinter;

            // Create an instance of the Printer
            IPrinter printer = new Printer();

            // Print the file
            //printer.PrintRawFile(PrinterName, Filepath, Filename);


            // if Is Section Finished NO - return user to Pallet Height/Sheet Count Screen
            if (sectionFinishedClicked == false)
            {
                sectionBtns = false;
                btnPalletCard.PerformClick();
                tbxSheetCountPalletCard.Text = "";
                tbxPalletHeightPalletCard.Text = "";
                dataGridView2.Columns.Clear();
                lblPrinting.Visible = false;
                tbxExtraInfoComment.Text = "";
                tbxBadSectionComment.Text = "";
                tbxSheetsAffectedBadSection.Text = "";
                tbxPaperDetails.Text = "";
                tbxInkDetails.Text = "";
                clearPosaPanel();
                tbxSearchBox.Text = "";
                tbxFinishPallet.Text = "";
                badSectionLbls = false;
                lbl5.Text = "";
            }

            // if Is Section Finished YES - return user to Choose Section or Home
            // if multiple sigs return to choose Section
            // if single sig return to Home
            // if Backup/Varnish Complete return to Home
            else if (sectionFinishedClicked == true)
            {
                sectionBtns = false;
                
                if (isBackupVarnishFromScan == true)
                {
                    tbxSearchBox.Text = "";
                    Cancel();
                    lblPrinting.Visible = false;
                    tbxExtraInfoComment.Text = "";
                    tbxBadSectionComment.Text = "";
                    tbxSheetsAffectedBadSection.Text = "";
                    tbxSheetCountPalletCard.Text = "";
                    tbxPalletHeightPalletCard.Text = "";
                    tbxPaperDetails.Text = "";
                    tbxInkDetails.Text = "";
                    clearPosaPanel();
                }
                else if (isBackupVarnishFromScan == false)
                {
                    tbxSearchBox.Text = dataGridView2.Rows[0].Cells[3].Value.ToString();
                    Search();
                    var distinctRowsPaperSectionNumber = (from DataGridViewRow row in dataGridView1.Rows select row.Cells[19].Value).Distinct().Count();
                    var distinctRowsSectionName = (from DataGridViewRow row in dataGridView1.Rows select row.Cells[15].Value).Distinct().Count();
                    var distinctRowsExpr1 = (from DataGridViewRow row in dataGridView1.Rows select row.Cells[11].Value).Distinct().Count();

                    if (distinctRowsPaperSectionNumber > 1 && distinctRowsSectionName > 1 || distinctRowsExpr1 > 1)
                    {
                        pnlHome1.BringToFront();
                        lbl2.Visible = false;
                        lbl3.Visible = false;
                        lbl4.Visible = false;
                        lbl5.Visible = false;
                        btnBack.Visible = false;
                        btnCancel.Visible = true;
                        tbxSheetCountPalletCard.Text = "";
                        tbxPalletHeightPalletCard.Text = "";
                        sectionFinishedClicked = false;
                        lblPrinting.Visible = false;
                        tbxExtraInfoComment.Text = "";
                        tbxBadSectionComment.Text = "";
                        tbxSheetsAffectedBadSection.Text = "";
                        tbxPaperDetails.Text = "";
                        tbxInkDetails.Text = "";
                        clearPosaPanel();
                    }
                    else if (distinctRowsPaperSectionNumber > 1)
                    {
                        pnlHome1.BringToFront();
                        lbl2.Visible = false;
                        lbl3.Visible = false;
                        lbl4.Visible = false;
                        lbl5.Visible = false;
                        btnBack.Visible = false;
                        btnCancel.Visible = true;
                        tbxSheetCountPalletCard.Text = "";
                        tbxPalletHeightPalletCard.Text = "";
                        sectionFinishedClicked = false;
                        lblPrinting.Visible = false;
                        tbxExtraInfoComment.Text = "";
                        tbxBadSectionComment.Text = "";
                        tbxSheetsAffectedBadSection.Text = "";
                        tbxPaperDetails.Text = "";
                        tbxInkDetails.Text = "";
                        clearPosaPanel();
                    }
                    else
                    {
                        sectionBtns = false;
                        tbxSearchBox.Text = "";
                        tbxFinishPallet.Text = "";
                        Search();
                        pnlHome0.BringToFront();
                        lblJobNo.Visible = false;
                        lblPress.Visible = false;
                        lbl1.Visible = false;
                        lbl2.Visible = false;
                        lbl3.Visible = false;
                        lbl4.Visible = false;
                        lbl5.Visible = false;
                        lblPrinting.Visible = false;
                        tbxExtraInfoComment.Text = "";
                        tbxBadSectionComment.Text = "";
                        tbxSheetsAffectedBadSection.Text = "";
                        tbxSheetCountPalletCard.Text = "";
                        tbxPalletHeightPalletCard.Text = "";
                        tbxPaperDetails.Text = "";
                        tbxInkDetails.Text = "";
                        clearPosaPanel();
                    }
                }              
            }
            dataGridView2.Columns.Clear();
            isBackupVarnishFromScan = false;
            btnCancel.Visible = true;
        }

        void clearPosaPanel()
        {
            Posa1NotChecked.BringToFront();
            Posa2NotChecked.BringToFront();
            Posa3NotChecked.BringToFront();
            Posa4NotChecked.BringToFront();
            Posa5NotChecked.BringToFront();
            Posa6NotChecked.BringToFront();
            Posa7NotChecked.BringToFront();
            Posa8NotChecked.BringToFront();
            Posa9NotChecked.BringToFront();
            PosaGripNotChecked.BringToFront();
            PosaRegistrationNotChecked.BringToFront();
            PosaColourNotChecked.BringToFront();
            PosaSTCRegisNotChecked.BringToFront();
            PosaSTCGlossNotChecked.BringToFront();
            PosaVisualDefectsNotChecked.BringToFront();
            textBox12.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";
            btnClearSignaturePosa.PerformClick();
        }

        void PrintImagePalletCard()
        {
            renderCrystalReport();
            renderPalletCard();
            Bitmap bmpDrawing1;
            System.Drawing.Rectangle rectBounds1;
            try
            {
                // Create bitmap for Pallet Card
                bmpDrawing1 = new Bitmap(840, 1052);
                // Set the bounds of the bitmap
                rectBounds1 = new System.Drawing.Rectangle(0, 0, 840, 1052);
                // Move drawing to bitmap
                pnlPalletCardPrintCR.DrawToBitmap(bmpDrawing1, rectBounds1);
                // Save the bitmap to file
                bmpDrawing1.Save("c:\\Temp\\front.jpg", System.Drawing.Imaging.ImageFormat.Bmp);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error on saving. Message: " + e.Message);
            }

            Bitmap bmpDrawing2;
            System.Drawing.Rectangle rectBounds2;
            try
            {
                // Create bitmap for Crystal Report
                bmpDrawing2 = new Bitmap(840, 1052);
                // Set the bounds of the bitmap
                rectBounds2 = new System.Drawing.Rectangle(0, 0, 840, 1052);
                // Move drawing to bitmap
                pnlPalletCardBack.DrawToBitmap(bmpDrawing2, rectBounds2);
                // Save the bitmap to file
                bmpDrawing2.Save("c:\\Temp\\back.jpg", System.Drawing.Imaging.ImageFormat.Bmp);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error on saving. Message: " + e.Message);
            }

            Document doc = new Document(iTextSharp.text.PageSize.A4, 0, 0, 0, 0);
            PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream("frontBackCombined.pdf", FileMode.Create));
            doc.Open();
            iTextSharp.text.Image pic1 = iTextSharp.text.Image.GetInstance("C:\\Temp\\front.jpg");
            iTextSharp.text.Image pic2 = iTextSharp.text.Image.GetInstance("C:\\Temp\\back.jpg");
            pic1.ScaleAbsolute(590f, 750f);
            pic2.ScaleAbsolute(590f, 750f);
            doc.Add(pic1);
            doc.Add(pic2);
            doc.Close();
        }

        private void btnPalletOver_Click(object sender, EventArgs e)
        {            
            if (dataGridView2.RowCount == 0)
            {
                int account = Convert.ToInt32(dataGridView1.Rows[0].Cells[21].Value);                    
                if (account == 1784 || account == 1781 || account == 1700 || account == 1795 || account == 1720 || account == 1839 || account == 1888)
                {
                    pnlPalletCard10.BringToFront();
                }
                else if (account == 1854 || account == 1305 || account == 1337 || account == 1435 || account == 1263 || account == 1092 || account == 1385 || account == 1598 || account == 1686 || account == 1726 || account == 1921 || account == 1458)
                {
                    pnlPalletCard11.BringToFront();
                }
                else
                {
                    pnlSignature.BringToFront();
                }
            }
            else if (dataGridView2.RowCount != 0)
            {
                int account = Convert.ToInt32(dataGridView2.Rows[0].Cells[29].Value);
                //if (account == 1784 || account == 1781 || account == 1700 || account == 1795 || account == 1720 || account == 1839 || account == 1888)
                    if (account == 1720 || account == 1839)
                {
                    pnlPalletCard10.BringToFront();
                }
                else if (account == 1854 || account == 1305 || account == 1337 || account == 1435 || account == 1263 || account == 1092 || account == 1385 || account == 1598 || account == 1686 || account == 1726 || account == 1921 || account == 1458)
                {
                    pnlPalletCard11.BringToFront();
                }
                else
                {
                    pnlSignature.BringToFront();
                }
            }
        }

        private void getNumberUp()
        {
            if (dataGridView2.Rows.Count == 0)
            {
                // if not equal to 1 then NumberUp =:
                numberUp = Convert.ToInt32(dataGridView1.Rows[0].Cells[12].Value);

                //if NumberUp field = 0
                if (numberUp == 0)
                {
                    // Parse the NumberUp value from Section Name or Expr1 - This regex finds continuous digits before "up"
                    if (dataGridView1.Rows[0].Cells[11].Value.ToString().Contains("up"))
                    {
                        String text = dataGridView1.Rows[0].Cells[11].Value.ToString();
                        foreach (Match match in Regex.Matches(text, @"(\d+)up"))
                        {
                            //MessageBox.Show(match.Groups[1].Value);
                            numberUp = Convert.ToInt32(match.Groups[1].Value);
                        }
                    }
                    else
                    {
                        String text1 = dataGridView1.Rows[0].Cells[15].Value.ToString();
                        foreach (Match match in Regex.Matches(text1, @"(\d+)up"))
                        {
                            //MessageBox.Show(match.Groups[1].Value);
                            numberUp = Convert.ToInt32(match.Groups[1].Value);
                        }
                    }
                }
            }
            else
            {
                numberUp = Convert.ToInt32(dataGridView2.Rows[0].Cells[21].Value);
            }
        }

#endregion

#region Posa Checkboxes

        private void Posa1Checked_Click(object sender, EventArgs e)
        {
            Posa1NotChecked.BringToFront();
        }

        private void Posa1NotChecked_Click(object sender, EventArgs e)
        {
            Posa1Checked.BringToFront();
        }

        private void Posa2Checked_Click(object sender, EventArgs e)
        {
            Posa2NotChecked.BringToFront();
        }

        private void Posa2NotChecked_Click(object sender, EventArgs e)
        {
            Posa2Checked.BringToFront();
        }

        private void Posa3Checked_Click(object sender, EventArgs e)
        {
            Posa3NotChecked.BringToFront();
        }
        private void Posa3NotChecked_Click(object sender, EventArgs e)
        {
            Posa3Checked.BringToFront();
        }

        private void Posa4Checked_Click(object sender, EventArgs e)
        {
            Posa4NotChecked.BringToFront();
        }

        private void Posa4NotChecked_Click(object sender, EventArgs e)
        {
            Posa4Checked.BringToFront();
        }

        private void Posa5Checked_Click(object sender, EventArgs e)
        {
            Posa5NotChecked.BringToFront();
        }

        private void Posa5NotChecked_Click(object sender, EventArgs e)
        {
            Posa5Checked.BringToFront();
        }

        private void Posa6Checked_Click(object sender, EventArgs e)
        {
            Posa6NotChecked.BringToFront();
        }

        private void Posa6NotChecked_Click(object sender, EventArgs e)
        {
            Posa6Checked.BringToFront();
        }

        private void Posa7Checked_Click(object sender, EventArgs e)
        {
            Posa7NotChecked.BringToFront();
        }

        private void Posa7NotChecked_Click(object sender, EventArgs e)
        {
            Posa7Checked.BringToFront();
        }

        private void Posa8Checked_Click(object sender, EventArgs e)
        {
            Posa8NotChecked.BringToFront();
        }

        private void Posa8NotChecked_Click(object sender, EventArgs e)
        {
            Posa8Checked.BringToFront();
        }

        private void Posa9Checked_Click(object sender, EventArgs e)
        {
            Posa9NotChecked.BringToFront();
        }

        private void Posa9NotChecked_Click(object sender, EventArgs e)
        {
            Posa9Checked.BringToFront();
        }



        private void PosaGripNotChecked_Click(object sender, EventArgs e)
        {
            PosaGripChecked.BringToFront();
        }

        private void PosaGripChecked_Click(object sender, EventArgs e)
        {
            PosaGripNotChecked.BringToFront();
        }

        private void PosaRegistrationChecked_Click(object sender, EventArgs e)
        {
            PosaRegistrationNotChecked.BringToFront();
        }

        private void PosaRegistrationNotChecked_Click(object sender, EventArgs e)
        {
            PosaRegistrationChecked.BringToFront();
        }

        private void PosaColourChecked_Click(object sender, EventArgs e)
        {
            PosaColourNotChecked.BringToFront();
        }

        private void PosaColourNotChecked_Click(object sender, EventArgs e)
        {
            PosaColourChecked.BringToFront();
        }

        private void PosaSTCRegisChecked_Click(object sender, EventArgs e)
        {
            PosaSTCRegisNotChecked.BringToFront();
        }

        private void PosaSTCRegisNotChecked_Click(object sender, EventArgs e)
        {
            PosaSTCRegisChecked.BringToFront();
        }

        private void PosaSTCGlossChecked_Click(object sender, EventArgs e)
        {
            PosaSTCGlossNotChecked.BringToFront();
        }

        private void PosaSTCGlossNotChecked_Click(object sender, EventArgs e)
        {
            PosaSTCGlossChecked.BringToFront();
        }

        private void PosaVisualDefectsChecked_Click(object sender, EventArgs e)
        {
            PosaVisualDefectsNotChecked.BringToFront();
        }

        private void PosaVisualDefectsNotChecked_Click(object sender, EventArgs e)
        {
            PosaVisualDefectsChecked.BringToFront();
        }
#endregion

//****************************************************************************************************
//  NOTIFICATION PANEL
//****************************************************************************************************
#region NotificationPanel

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
                                    btn.Top = DynamicButtonsOrigin * 100;
                                    btn.Height = 80;
                                    btn.Width = 465;
                                    btn.BackColor = Color.SteelBlue;
                                    btn.Font = new System.Drawing.Font("Microsoft Sans Serif", 14);
                                    btn.ForeColor = Color.White;
                                    btn.Left = 30;
                                    btn.Text = this.dataGridView1.Rows[i].Cells[11].Value as string;
                                    DynamicButtonsOrigin = DynamicButtonsOrigin + 1;
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
            //MailMessage mail = new MailMessage("Press@colorman.ie", defaultEmail , "Waiting for Plates", "Job Number " + lblJobNo.Text + " - Section " + dataGridView1.Rows[0].Cells[11].Value.ToString() + " - is waiting for plates");
            //SmtpClient client = new SmtpClient("ex0101.ColorMan.local");
            //client.Port = 25;
            //client.EnableSsl = false;
            //client.Send(mail);

            MailMessage mail = new MailMessage();
            string from = press + "@colorman.ie";
            mail.From = new MailAddress(from);
            mail.To.Add(defaultEmail);
            mail.Subject = "Waiting for Plates";
            mail.Body = "Waiting for Plates" + "Job Number " + lblJobNo.Text + " - Section " + dataGridView1.Rows[0].Cells[11].Value.ToString() + " - is waiting for plates";
            SmtpClient client = new SmtpClient("ex0101.ColorMan.local");
            client.Port = 25;
            client.EnableSsl = false;
            client.Send(mail);
            pnlHome1.BringToFront();
        }

        private void btnWaitingPaper_Click(object sender, EventArgs e)
        {
            // Send email notification
            MailMessage mail = new MailMessage(press + "@colorman.ie", defaultEmail , "Waiting for Paper", "Job Number " + lblJobNo.Text + " - Section " + dataGridView1.Rows[0].Cells[11].Value.ToString() + " - is waiting for paper");
            SmtpClient client = new SmtpClient("ex0101.ColorMan.local");
            client.Port = 25;
            client.EnableSsl = false;
            client.Send(mail);
            pnlHome1.BringToFront();
        }

        private void btnJobLifted_Click(object sender, EventArgs e)
        {
            // Send email notification
            MailMessage mail = new MailMessage(press + "@colorman.ie", defaultEmail , "Job Lifted", "Job Number " + lblJobNo.Text + " - Section " + dataGridView1.Rows[0].Cells[11].Value.ToString() + " - Job is lifted");
            SmtpClient client = new SmtpClient("ex0101.ColorMan.local");
            client.Port = 25;
            client.EnableSsl = false;
            client.Send(mail);
            pnlHome1.BringToFront();
        }


#endregion


    }
}
