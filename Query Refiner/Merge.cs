using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework.Forms;
using MetroFramework;
using System.Data.OleDb;
using static Query_Refiner.DictionaryInitializer;
using System.Configuration;
using System.Globalization;
using OfficeOpenXml;
using System.IO;
using System.Transactions;
using System.Diagnostics;

namespace Query_Refiner
{
    public partial class Merge : MetroForm
    {
        public Merge()
        {
            InitializeComponent();
        }
        string ImportedSheet;
        public string IO_Name;
        int MaxSchema;
        int MaxExcel;
        OleDbConnection OleDbcon;
        OleDbDataAdapter ODA;
        DataTable dt = new DataTable();
        DataTable VirtualTable;
        public DictionaryInit theDictionary;
        public string Current_Table;
        String ETA = "";
        public string Con = ConfigurationManager.ConnectionStrings["Con"].ConnectionString;

        private void Merge_Load(object sender, EventArgs e)
        {
            Resize();
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx"
            };
            openFileDialog.ShowDialog();
            if (!string.IsNullOrEmpty(openFileDialog.FileName))
            {
                lblStatuses.Text = "Loading.. Please wait.";
                IO_Name = openFileDialog.FileName;
                cmbSheets.Items.Clear();
                cmbSheets.Enabled = false;
                this.Enabled = false;
                Worker_Import.RunWorkerAsync();
            }
        }

        private void Worker_Import_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                OleDbcon =
                    new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + IO_Name +
                                        ";Extended Properties=  'Excel 12.0 Xml;HDR=Yes; IMEX = 1;TypeGuessRows=0;ImportMixedTypes=Text';");
                OleDbcon.Open();
                DataTable dt = OleDbcon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                OleDbcon.Close();
                MaxSchema = dt.Rows.Count;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    String sheetName = dt.Rows[i]["TABLE_NAME"].ToString();
                    sheetName = sheetName.Substring(0, sheetName.Length - 1);
                    cmbSheets.Invoke((MethodInvoker)delegate {
                        cmbSheets.Items.Add(sheetName);
                    });
                    Worker_Import.ReportProgress(i);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Write Excel: " + ex.Message);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Worker_Import.ReportProgress(MaxSchema);
            }
        }

        private void Worker_Import_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pBarMain.Value = (e.ProgressPercentage * 100) / MaxSchema;
        }

        private async void Worker_Import_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Red;
            }
            else
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Lime;
            }
            lblStatuses.Text = "Task completed.";
            cmbSheets.Enabled = true;
            this.Enabled = true;
            btnDownload.Enabled = false;
            await PutTaskDelay();
            pBarMain.Value = 0;
            pBarMain.Style = MetroColorStyle.Default;

        }

        public static int Truth(params bool[] booleans)
        {
            return booleans.Count(b => b);
        }

        private void btnDownload_Click(object sender, EventArgs e)
        {
            if ((chkbxViewer.Checked == true) && (DGVMerge.Rows.Count == 0))
            {
                MessageBox.Show("No rows detected!");
                return;
            }
            if (!(Truth(rdbtnYYYY.Checked == true, rdbtnMDY.Checked == true, rdbtnOA.Checked == true) >= 1))
            {
                MessageBox.Show("Select a date type!");
                return;
            }
            DialogResult dr = MessageBox.Show(
                "Caution: imported data must match database structure.", "Import caution",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (dr == DialogResult.OK)
            {
                try
                {
                    this.Enabled = false;
                    cmbSheets.Enabled = false;
                    btnDownload.Enabled = false;
                    VirtualTable = dt;
                    MaxExcel = VirtualTable.Rows.Count;
                    Worker_Transfer.RunWorkerAsync();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Validator Error: " + ex);
                }
            }
        }

        private void Worker_Transfer_DoWork(object sender, DoWorkEventArgs e)
        {
            theDictionary = new DictionaryInit();
            if (MaxExcel <= 0)
            {
                MessageBox.Show("A failed attempt has reset the inserted sheet, repeat the process correctly!");
                cmbSheets.DataSource = null;
                return;
            }
            try
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    using (OleDbCommand cmd = new OleDbCommand())
                    {
                        cmd.Connection = connection;
                        cmd.CommandText =
                                    "Insert INTO [" + Current_Table + "] (Business_Unit,Ledger,Account,Description,Span_Year,Period,Book_Code,Alt_Currency,Amount_USD,Base_Currency,PDS,Amount_PHP,Site,Stat_Site,Status_Name,Status_Regime,Department,Description_Department,Project_Number,Description_Project,Affiliate,IS_Accounts,Description_ITR) " +
                                                                  "VALUES(@Business_Unit,@Ledger,@Account,@Description,@Span_Year,@Period,@Book_Code,@Alt_Currency,@Amount_USD,@Base_Currency,@PDS,@Amount_PHP,@Site,@Stat_Site,@Status_Name,@Status_Regime,@Department,@Description_Department,@Project_Number,@Description_Project,@Affiliate,@IS_Accounts,@Description_ITR)";
                        connection.Open();
                        using (OleDbTransaction Scope = connection.BeginTransaction(System.Data.IsolationLevel.ReadCommitted))
                        {
                            try
                            {
                                cmd.Transaction = Scope;
                                int CurrentPoint;
                                for (int rowindex = 0; rowindex < MaxExcel; rowindex++)
                                {
                                    CurrentPoint = rowindex;
                                    cmd.Parameters.Clear();
                                    DictionaryFindTable(Convert.ToInt32(VirtualTable.Rows[rowindex][1]), theDictionary.BUtoQRs);
                                    cmd.Parameters.AddWithValue("@Business_Unit", VirtualTable.Rows[rowindex][1]);
                                    cmd.Parameters.AddWithValue("@Ledger", VirtualTable.Rows[rowindex][2]);
                                    cmd.Parameters.AddWithValue("@Account", VirtualTable.Rows[rowindex][3]);
                                    cmd.Parameters.AddWithValue("@Description", VirtualTable.Rows[rowindex][4]);
                                    if (rdbtnYYYY.Checked == true)
                                    {
                                        try
                                        {
                                            cmd.Parameters.AddWithValue("@Span_Year", (VirtualTable.Rows[rowindex][5] == DBNull.Value) ? VirtualTable.Rows[rowindex][9] = 1000 : new DateTime(Convert.ToInt32(VirtualTable.Rows[rowindex][5]), 1, 1));
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("Invalid date! " + ex);
                                        }
                                    }
                                    else if (rdbtnMDY.Checked == true)
                                    {
                                        string date = VirtualTable.Rows[rowindex][5].ToString();
                                        Console.WriteLine("{" + date);
                                        if (DateTime.TryParse(date, out DateTime result) == true)
                                        {
                                            try
                                            {
                                                cmd.Parameters.AddWithValue("@Span_Year", (VirtualTable.Rows[rowindex][5] == DBNull.Value) ? VirtualTable.Rows[rowindex][9] = "1/1/1000" : result.ToString("MM/dd/yyyy"));
                                                Console.WriteLine(result);
                                            }
                                            catch (FormatException)
                                            {

                                                MessageBox.Show("Invalid format!");
                                            }
                                        }
                                    }
                                    else if (rdbtnOA.Checked == true)
                                    {
                                        try
                                        {
                                            string YY = (VirtualTable.Rows[rowindex][5]).ToString();
                                            double d = double.Parse(YY);
                                            DateTime conv = DateTime.FromOADate(d);
                                            cmd.Parameters.AddWithValue("@Span_Year", conv);
                                        }
                                        catch (FormatException)
                                        {
                                            MessageBox.Show("Invalid format!");
                                        }
                                    }
                                    cmd.Parameters.AddWithValue("@Period", VirtualTable.Rows[rowindex][6]);
                                    cmd.Parameters.AddWithValue("@Book_Code", VirtualTable.Rows[rowindex][7]);
                                    cmd.Parameters.AddWithValue("@Currency", VirtualTable.Rows[rowindex][8]);
                                    cmd.Parameters.AddWithValue("@Amount_USD", (VirtualTable.Rows[rowindex][9] == null) ? VirtualTable.Rows[rowindex][9] = 0 : Convert.ToDouble(VirtualTable.Rows[rowindex][9]));
                                    cmd.Parameters.AddWithValue("@Base_Currency", VirtualTable.Rows[rowindex][10]);
                                    cmd.Parameters.AddWithValue("@PDS", VirtualTable.Rows[rowindex][11]);
                                    cmd.Parameters.AddWithValue("@Amount_PHP", (VirtualTable.Rows[rowindex][12] == null) ? VirtualTable.Rows[rowindex][12] = 0 : Convert.ToDouble(VirtualTable.Rows[rowindex][12]));
                                    cmd.Parameters.AddWithValue("@Site", VirtualTable.Rows[rowindex][13]);
                                    cmd.Parameters.AddWithValue("@Stat_Site", VirtualTable.Rows[rowindex][14]);
                                    cmd.Parameters.AddWithValue("@Status_Name", VirtualTable.Rows[rowindex][15]);
                                    cmd.Parameters.AddWithValue("@Status_Regime", VirtualTable.Rows[rowindex][16]);
                                    cmd.Parameters.AddWithValue("@Department", VirtualTable.Rows[rowindex][17]);
                                    cmd.Parameters.AddWithValue("@Description_Department", VirtualTable.Rows[rowindex][18]);
                                    cmd.Parameters.AddWithValue("@Project_Number", VirtualTable.Rows[rowindex][19]);
                                    cmd.Parameters.AddWithValue("@Description_Project", VirtualTable.Rows[rowindex][20]);
                                    cmd.Parameters.AddWithValue("@Affiliate", VirtualTable.Rows[rowindex][21]);
                                    cmd.Parameters.AddWithValue("@IS_Accounts", VirtualTable.Rows[rowindex][22]);
                                    cmd.Parameters.AddWithValue("@Description_ITR", VirtualTable.Rows[rowindex][23]);
                                    cmd.ExecuteNonQuery();
                                    Worker_Transfer.ReportProgress(rowindex);
                                    lblStatuses.Invoke((MethodInvoker)delegate
                                    {
                                        lblStatuses.Text = "Processing row " + CurrentPoint + "/" + MaxExcel;
                                    });
                                }
                                Scope.Commit();
                            }
                            catch (Exception odx)
                            {
                                MessageBox.Show(odx.Message);
                                Scope.Rollback();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Import error: " + ex);
            }
            Worker_Transfer.ReportProgress(MaxExcel);
        }

        private void Worker_Transfer_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pBarMain.Value = (e.ProgressPercentage * 100) / MaxExcel;
        }

        private async void Worker_Transfer_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Yellow;
            }
            else if (e.Error != null)
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Red;
            }
            else
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Lime;
            }
            cmbSheets.Enabled = true;
            this.Enabled = true;
            lblStatuses.Text = "Process completed.";
            btnDownload.Enabled = true;
            await PutTaskDelay();
            pBarMain.Style = MetroColorStyle.Default;
            VirtualTable.Clear();
            VirtualTable.Dispose();
            pBarMain.Value = 0;
        }

        async Task PutTaskDelay()
        {
            await Task.Delay(5000);
        }

        private DataTable GetDataTableFromDGV(DataGridView dgv)
        {
            VirtualTable = new DataTable();
            foreach (DataGridViewColumn column in dgv.Columns)
            {
                if (column.Visible)
                {
                    VirtualTable.Columns.Add();
                }
            }


            object[] cellValues = new object[dgv.Columns.Count];
            foreach (DataGridViewRow row in dgv.Rows)
            {
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    cellValues[i] = row.Cells[i].Value;
                }
                VirtualTable.Rows.Add(cellValues);
            }

            return VirtualTable;
        }

        public void Validator()
        {
            List<string> ExcludedColumnsList = new List<string> { "ID", "Business_Unit", "Ledger", "Description", "Alt_Currency", "Base_Currency", "Site", "Stat_Site", "Status_Name", "Status_Regime", "Description_Department", "Description_Project", "Affiliate", "IS_Accounts", "Description_ITR" };
            //
            foreach (DataRow row in dt.Rows)
            {
                for (int i = 0; i < row.ItemArray.Count(); i++)
                {
                    if (!ExcludedColumnsList.Contains(DGVMerge.Columns[i].Name))
                    {
                        if ((DGVMerge.Columns[i].Name == "Span_Year") && (row.ItemArray[i] == null || row.ItemArray[i] == DBNull.Value ||
                                String.IsNullOrWhiteSpace(row.ItemArray[i].ToString())))
                        {
                            row.ItemArray[i] = 1000;
                        }
                        if (row.ItemArray[i] == null || row.ItemArray[i] == DBNull.Value ||
                                String.IsNullOrWhiteSpace(row.ItemArray[i].ToString()))
                        {
                            row.ItemArray[i] = 0;
                            //DGVExcel.RefreshEdit();
                        }
                    }
                }
            }
        }

        public void DictionaryFindTable(int MapKey, Dictionary<int, DictionarySetup> AccountLexicon)
        {
            if (AccountLexicon.TryGetValue(MapKey, out DictionarySetup ClassValues))
            {
                Current_Table = ClassValues.theDescription;
                if (chkbxBS.Checked == true)
                {
                    Current_Table += "_BS";
                }
            }
        }

        private void cmbSheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            ImportedSheet = cmbSheets.Text;
            if ((Convert.ToInt32(UDMin.Value) == 0) && (Convert.ToInt32(UDMax.Value) == 0))
            {
                btnDownload.Enabled = false;
                cmbSheets.Enabled = false;
                string query = "Select * from [" + cmbSheets.Text + "$]";
                ODA = new OleDbDataAdapter(query, OleDbcon);
            }
            else if ((Math.Abs(Convert.ToInt32(UDMax.Value))) > 0)
            {
                btnDownload.Enabled = false;
                cmbSheets.Enabled = false;
                string query = "Select * from [" + cmbSheets.Text + "$" + Convert.ToInt32(UDMin.Value) + ":" + (Convert.ToInt32(UDMax.Value) + 1) + "]";
                Console.WriteLine(query);
                ODA = new OleDbDataAdapter(query, OleDbcon);
            }
            else
            {
                MessageBox.Show("Invalid Min and Max! Leave both at zero to pull the entire set of rows.");
                return;
            }
            PassiveWorker.RunWorkerAsync();
        }

        private void btnDispose_Click(object sender, EventArgs e)
        {
            DisposeThisTable("ACTB_CPSC");
            DisposeThisTable("ACTB_CSHI");
            DisposeThisTable("ACTB_CSPI");
            DisposeThisTable("ACTB_ERMI");
            DisposeThisTable("ACTB_CPI");
            DisposeThisTable("ACTB_CMPB");
            DisposeThisTable("ACTB_CGSP");
            DisposeThisTable("ACTB_CPSC_BS");
            DisposeThisTable("ACTB_CSHI_BS");
            DisposeThisTable("ACTB_CSPI_BS");
            DisposeThisTable("ACTB_ERMI_BS");
            DisposeThisTable("ACTB_CPI_BS");
            DisposeThisTable("ACTB_CMPB_BS");
            DisposeThisTable("ACTB_CGSP_BS");
            lblStatuses.Text = "All records have been purged";
        }

        public void DisposeThisTable(string Deletable)
        {
            DialogResult dr = MessageBox.Show(
"Warning: All data in "+Deletable+" will be permanently deleted. The data will no longer be recoverable, do you wish to proceed?", "Disposing records",
MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (dr == DialogResult.OK)
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    try
                    {
                        connection.Open();
                        using (OleDbTransaction Scope = connection.BeginTransaction(System.Data.IsolationLevel.ReadCommitted))
                        {
                            try
                            {

                                OleDbCommand cmd = new OleDbCommand("DELETE FROM [" + Deletable + "]", connection);
                                cmd.Transaction = Scope;
                                cmd.ExecuteNonQuery();
                                Scope.Commit();
                            }
                            catch (OleDbException odx)
                            {
                                MessageBox.Show(odx.Message);
                                Scope.Rollback();
                            }
                        }
                    }
                    catch (OleDbException es)
                    {
                        MessageBox.Show("SQL error: " + es);
                    }
                }
            }
        }

        private void PassiveWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            dt = new DataTable();
            ODA.Fill(dt);
            if (chkbxViewer.Checked == true)
            {
                DGVMerge.Invoke((MethodInvoker)delegate
                {
                    DGVMerge.DataSource = dt;
                });
            }
        }

        private void PassiveWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            lblItem.Text = DGVMerge.RowCount.ToString();
            lblStatuses.Text = "Loaded the sheet";
            cmbSheets.Enabled = true;
            btnDownload.Enabled = true;
        }

        private void chkbxViewer_CheckedChanged(object sender, EventArgs e)
        {
            if (chkbxViewer.Checked == true)
            {
                DGVMerge.Show();
                metroPanel1.Width = 725;
                this.Width = 720;//720, 424 863/865
                this.Height = 445;
            }
            else
            {
                metroPanel1.Width = 363;
                DGVMerge.Hide();
                this.Width = 210;
            }
        }

        public void Resize()
        {
            metroPanel1.Width = 363;
            DGVMerge.Hide();
            this.Width = 210;
        }

        private void rdbtnYYYY_Click(object sender, EventArgs e)
        {
            rdbtnYYYY.Checked = true;
            rdbtnMDY.Checked = false;
            rdbtnOA.Checked = false;
        }

        private void rdbtnMDY_Click(object sender, EventArgs e)
        {
            rdbtnYYYY.Checked = false;
            rdbtnMDY.Checked = true;
            rdbtnOA.Checked = false;
        }

        private void rdbtnOA_Click(object sender, EventArgs e)
        {
            rdbtnYYYY.Checked = false;
            rdbtnMDY.Checked = false;
            rdbtnOA.Checked = true;
        }

        private void btnExceltoTable_Click(object sender, EventArgs e)
        {
            lblStatuses.Text = "Starting upload process...";
            UploadExcel();
        }

        DataTable DT;

        public void UploadExcel()
        {
            if ((chkbxViewer.Checked == true) && (DGVMerge.Rows.Count == 0))
            {
                MessageBox.Show("No rows detected!");
                return;
            }
            if (!(Truth(rdbtnYYYY.Checked == true, rdbtnMDY.Checked == true, rdbtnOA.Checked == true) >= 1))
            {
                MessageBox.Show("Select a date type!");
                return;
            }
            var OFD = new OpenFileDialog();
            OFD.ShowDialog();
            string IO_Name = OFD.FileName;
            if (IO_Name == null || IO_Name == "" || IO_Name == " ")
            {
                return;
            }
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(IO_Name))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.First();
                DT = ToDataTable(ws);
            }
            //DGVMerge.DataSource = DT;
            MaxExcel = DT.Rows.Count;
            Console.WriteLine(MaxExcel);
            panel1.Enabled = false;
            EPPlusWorker.RunWorkerAsync();
        }

        public static DataTable ToDataTable(ExcelWorksheet ws, bool hasHeaderRow = true)
        {
            var tbl = new DataTable();
            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column]) tbl.Columns.Add(hasHeaderRow ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            var startRow = hasHeaderRow ? 2 : 1;
            for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                var row = tbl.NewRow();
                foreach (var cell in wsRow) row[cell.Start.Column - 1] = cell.Text;
                tbl.Rows.Add(row);
            }
            return tbl;
        }

        string rowIndexer = "";

        private void EPPlusWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            if (DT == null)
            {
                MessageBox.Show("Datatable detected no values!");
                return;
            }
            theDictionary = new DictionaryInit();
            Console.WriteLine(DateTime.Now + " < Qck");
            try
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    using (OleDbCommand cmd = new OleDbCommand())
                    {
                        cmd.Connection = connection;
                        connection.Open();
                        using (OleDbTransaction Scope = connection.BeginTransaction(System.Data.IsolationLevel.ReadCommitted))
                        {
                            try
                            {
                                cmd.Transaction = Scope;
                                int CurrentPoint;
                                for (int rowindex = 0; rowindex < MaxExcel; rowindex++)
                                {
                                    if (DT.Rows[rowindex][1] == DBNull.Value)
                                    {
                                        continue;
                                    }
                                    CurrentPoint = rowindex;
                                    cmd.Parameters.Clear();
                                    DictionaryFindTable(Convert.ToInt32(DT.Rows[rowindex][1]), theDictionary.BUtoQRs);
                                    cmd.CommandText =
                                    "Insert INTO [" + Current_Table + "] (Business_Unit,Ledger,Account,Description,Span_Year,Period,Book_Code,Alt_Currency,Amount_USD,Base_Currency,PDS,Amount_PHP,Site,Stat_Site,Status_Name,Status_Regime,Department,Description_Department,Project_Number,Description_Project,Affiliate,IS_Accounts,Description_ITR) " +
                                                                  "VALUES(@Business_Unit,@Ledger,@Account,@Description,@Span_Year,@Period,@Book_Code,@Alt_Currency,@Amount_USD,@Base_Currency,@PDS,@Amount_PHP,@Site,@Stat_Site,@Status_Name,@Status_Regime,@Department,@Description_Department,@Project_Number,@Description_Project,@Affiliate,@IS_Accounts,@Description_ITR)";
                                    cmd.Parameters.AddWithValue("@Business_Unit", DT.Rows[rowindex][1]);
                                    cmd.Parameters.AddWithValue("@Ledger", DT.Rows[rowindex][2]);
                                    cmd.Parameters.AddWithValue("@Account", DT.Rows[rowindex][3]);
                                    cmd.Parameters.AddWithValue("@Description", DT.Rows[rowindex][4]);
                                    if (rdbtnYYYY.Checked == true)
                                    {
                                        try
                                        {
                                            cmd.Parameters.AddWithValue("@Span_Year", (DT.Rows[rowindex][5] == null) ? DT.Rows[rowindex][9] = 1000 : new DateTime(Convert.ToInt32(DT.Rows[rowindex][5]), 1, 1));
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("Invalid date! " + ex);
                                        }
                                    }
                                    else if (rdbtnMDY.Checked == true)
                                    {
                                        string date = DT.Rows[rowindex][5].ToString();
                                        Console.WriteLine("{" + date);
                                        if (DateTime.TryParse(date, out DateTime result) == true)
                                        {
                                            try
                                            {
                                                cmd.Parameters.AddWithValue("@Span_Year", (DT.Rows[rowindex][5] == null) ? DT.Rows[rowindex][9] = 1000 : result.ToString("MM/dd/yyyy"));
                                                Console.WriteLine(result);
                                            }
                                            catch (FormatException)
                                            {

                                                MessageBox.Show("Invalid format!");
                                            }
                                        }
                                    }
                                    else if (rdbtnOA.Checked == true)
                                    {
                                        try
                                        {
                                            string YY = (DT.Rows[rowindex][5]).ToString();
                                            double d = double.Parse(YY);
                                            DateTime conv = DateTime.FromOADate(d);
                                            cmd.Parameters.AddWithValue("@Span_Year", conv);
                                        }
                                        catch (FormatException)
                                        {
                                            MessageBox.Show("Invalid format!");
                                        }
                                    }
                                    cmd.Parameters.AddWithValue("@Period", DT.Rows[rowindex][6]);
                                    cmd.Parameters.AddWithValue("@Book_Code", DT.Rows[rowindex][7]);
                                    cmd.Parameters.AddWithValue("@Currency", DT.Rows[rowindex][8]);
                                    if (Convert.ToDouble(DT.Rows[rowindex][9]) == 0 || Convert.ToDouble(DT.Rows[rowindex][9]) == 0.00)
                                    {
                                        cmd.Parameters.AddWithValue("@Amount_USD", 0);
                                    }
                                    else
                                    {
                                        cmd.Parameters.AddWithValue("@Amount_USD", (DT.Rows[rowindex][9] == null) ? DT.Rows[rowindex][9] = 0 : Convert.ToDouble(DT.Rows[rowindex][9]));
                                    }
                                    cmd.Parameters.AddWithValue("@Base_Currency", DT.Rows[rowindex][10]);
                                    cmd.Parameters.AddWithValue("@PDS", DT.Rows[rowindex][11]);
                                    if (Convert.ToDouble(DT.Rows[rowindex][12]) == 0 || Convert.ToDouble(DT.Rows[rowindex][12]) == 0.00)
                                    {
                                        cmd.Parameters.AddWithValue("@Amount_PHP", 0);
                                    }
                                    else
                                    {
                                        cmd.Parameters.AddWithValue("@Amount_PHP", (DT.Rows[rowindex][12] == null) ? DT.Rows[rowindex][12] = 0 : Convert.ToDouble(DT.Rows[rowindex][12]));
                                    }
                                    cmd.Parameters.AddWithValue("@Site", DT.Rows[rowindex][13]);
                                    cmd.Parameters.AddWithValue("@Stat_Site", DT.Rows[rowindex][14]);
                                    cmd.Parameters.AddWithValue("@Status_Name", DT.Rows[rowindex][15]);
                                    cmd.Parameters.AddWithValue("@Status_Regime", DT.Rows[rowindex][16]);
                                    cmd.Parameters.AddWithValue("@Department", DT.Rows[rowindex][17]);
                                    cmd.Parameters.AddWithValue("@Description_Department", DT.Rows[rowindex][18]);
                                    cmd.Parameters.AddWithValue("@Project_Number", DT.Rows[rowindex][19]);
                                    cmd.Parameters.AddWithValue("@Description_Project", DT.Rows[rowindex][20]);
                                    cmd.Parameters.AddWithValue("@Affiliate", DT.Rows[rowindex][21]);
                                    cmd.Parameters.AddWithValue("@IS_Accounts", DT.Rows[rowindex][22]);
                                    cmd.Parameters.AddWithValue("@Description_ITR", DT.Rows[rowindex][23]);
                                    rowIndexer = rowindex.ToString();
                                    cmd.ExecuteNonQuery();
                                    EPPlusWorker.ReportProgress(rowindex);
                                    lblStatuses.Invoke((MethodInvoker)delegate
                                    {
                                        lblStatuses.Text = "Processing row " + CurrentPoint + "/" + MaxExcel;
                                    });
                                }
                                Scope.Commit();
                            }
                            catch (OleDbException odx)
                            {
                                MessageBox.Show(odx.Message + rowIndexer);
                                Scope.Rollback();
                            }
                        }
                    }
                }
            }
            catch (OleDbException ex)
            {
                MessageBox.Show("Import error: " + ex);
            }
            finally
            {
                EPPlusWorker.ReportProgress(MaxExcel);
                sw.Stop();
                ETA = sw.Elapsed.ToString();
            }
        }
       
        private void EPPlusWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pBarMain.Value = (e.ProgressPercentage * 100) / MaxExcel;
        }

        private async void EPPlusWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Yellow;
                MessageBox.Show("Cancelled processing at row index: " + rowIndexer);
            }
            else if (e.Error != null)
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Red;
                MessageBox.Show("Error, stopped processing at row index: " + rowIndexer);
            }
            else
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Lime;
            }
            lblElapsedTime.Text = ETA;
            cmbSheets.Enabled = true;
            lblStatuses.Text = "Process completed";
            btnDownload.Enabled = true;
            await PutTaskDelay();
            panel1.Enabled = true;
            pBarMain.Style = MetroColorStyle.Default;
            Console.WriteLine(DateTime.Now + " < Qck");
            MessageBox.Show("Process completed "+rowIndexer+" rows!");
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }

}
