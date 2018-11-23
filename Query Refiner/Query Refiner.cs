using MetroFramework.Components;
using MetroFramework.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework;
using System.Configuration;
using System.Data.OleDb;
using System.Diagnostics;
using SD = System.Data;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using static Query_Refiner.DictionaryInitializer;
using System.Globalization;
using System.Xml;
using System.Text.RegularExpressions;
using System.IO;
using OfficeOpenXml;
using System.Web.Script.Serialization;
using Newtonsoft.Json;
using System.Transactions;
using System.Threading;

namespace Query_Refiner
{
    public partial class QueryRefiner : Form
    {
        Account AccountForm = new Account();
        DictionaryInit IncomeDictionary = new DictionaryInit();
        DictionaryInit_BS BalanceDictionary = new DictionaryInit_BS();
        DataTable VirtualTable;

        /// <summary>
        /// These handles are the parameters passed from the initial selection in the Account.cs
        /// </summary>
        /// <param name="BusinessUnit"> Parameter determines the documents BU-code, consequentially, the current table</param>
        /// <param name="DocumentType"> Decides whether to pull the table for balance sheet, or income statement</param>
        /// <param name="CurrencyType"> Determines whether to use special operations regarding the sum total amount on PHP or USD columns</param>
        public string BusinessUnit { get; set; }
        public string DocumentType { get; set; }
        public string CurrencyType { get; set; }

        /// <summary>
        /// This class handles the stopwatch and certain variables for debugging purposes
        /// </summary>
        /// <param name="FilterTable"> Returns the table being filtered </param>
        /// <param name="Occurrence"> Passes the string for time elapsed during process</param>
        public class TableStorage
        {
            public DataTable FilterTable { get; set; }
            public String Occurrence { get; set; }
        }
        /// <summary>
        /// This class handles the parameters used in the transform master function that pivots table data based on a desired set of rules or criteria.
        /// </summary>
        /// <param name="hasRevenues/Cost/Operations/Other/Income"> Determines wheether a given row has enough data to qualify adding a row for it.</param>
        /// <param name="queryBuilder"> Builds a query string by concatenating the rules of the user into a single statement</param>
        /// <param name="finalizedString"> In the event that the user opts to add bonus selections that add more columns, this string appends the request to the main query</param>
        public class TransformCriteria
        {
            public Boolean HasRevenues { get; set; }
            public Boolean HasCost { get; set; }
            public Boolean HasOperations { get; set; }
            public Boolean HasOther { get; set; }
            public Boolean HasIncome { get; set; }
            public StringBuilder QueryBuilder = new StringBuilder();
            public string FinalizedString { get; set; }
        }
        /// <summary>
        /// This class handles the parameter criterion for generating a filter
        /// </summary>
        /// <param name="hasPeriodBetween/ITR/IS/ETC"> Boolean that determines the presence of data to qualify for filtering</param>
        public class MatchesCriteria
        {
            public Boolean HasPeriodBetween { get; set; }
            public Boolean HasITR { get; set; }
            public Boolean HasIS { get; set; }
            public Boolean HasDept { get; set; }
            public Boolean HasStatus { get; set; }
            public Boolean HasSite { get; set; }
            public Boolean HasUSD { get; set; }
            public Boolean HasPHP { get; set; }
            public Boolean HasBook { get; set; }
            public Boolean HasPeriod { get; set; }
            public Boolean HasYear { get; set; }
            public Boolean HasYearBetween { get; set; }
            public Boolean HasAccount { get; set; }
            public Boolean HasPHPDecimal { get; set; }
            public Boolean HasUSDDecimal { get; set; }
            public Boolean HasOR { get; set; }
        }
        /// <summary>
        /// This class handles the actions used in filtering
        /// </summary>
        /// <param name="Starter/Ender"> Parameter that determines the start and end for year selections, same for period starter and ender</param>
        /// <param name="init/init2"> Determines the parameters used for the decimal query function </param>
        /// <param name="Acc/Year/Period/ETC"> Holds the values from the filter textboxes for use in the filter worker </param>
        /// <param name="MaxLoops/currentLoops"> MaxLoops determines the loop count while current loops determines the number of times loops will have to decrement </param>
        public class GetFilter
        {
            public string Starter { get; set; }
            public string Ender { get; set; }
            public Double init { get; set; }
            public Double init2 { get; set; }
            public string PeriodStarter { get; set; }
            public string PeriodEnder { get; set; }            
            public StringBuilder FilterBuilder { get; set; }
            public String Acc { get; set; }
            public String Year { get; set; }
            public String Period { get; set; }
            public String Book { get; set; }
            public String PHP { get; set; }
            public String USD { get; set; }
            public String Site { get; set; }
            public String Status { get; set; }
            public String Dept { get; set; }
            public String IS { get; set; }
            public String ITR { get; set; }
            public int MaxLoops { get; set; }
            public int currentLoops = 0;
        }
        /// <summary>
        /// This class handles the Json serializer and deserializers
        /// </summary>
        /// <param name="columnString"> The column string is pulled from the combo box for download purposes, or upload. </param>
        /// <param name="SelectedDictionary"> Current selected dictionary </param>
        /// <param name="Found_Dictionary"> Dictionary selected for replacement </param>
        /// <remarks> TODO: Make this work... </remarks>
        public class JsonParameters
        {
            public string ColumnString { get; set; }
            public string SelectedDictionary { get; set; }
            public Dictionary<int, DictionaryCheckup> Found_Dictionary_BS;
            public Dictionary<int, DictionarySetup> Found_Dictionary_IS;
        }
        /// <summary>
        /// This class contains various max row counts for use in various scenarios
        /// </summary>
        public class RowCountCollection
        {
            public int MaxExcel { get; set; }
            public int MaxSchema { get; set; }
            public int MaxRelocate { get; set; }
            public int MaxRows { get; set; }
        }
        /// <summary>
        /// This class contains various actions used to initialize components or variables
        /// </summary>
        public class SetupActions
        {
            public string ExcelSheet { get; set; }
            public int TimerIntervals = 0;
            public int OneTimeWarning = 0;
            public string CurrentTable { get; set; }
            public TextBox FocusedControl { get; set; }
            public string BonusSelectionBuilder { get; set; }
            public string IO_Name { get; set; }
        }
        /// <summary>
        /// This contains the connection string used for literally the whole of the program
        /// </summary>
        public string Con = ConfigurationManager.ConnectionStrings["Con"].ConnectionString;
        public OleDbConnection Conn { get; set; }

        SetupActions SetupAction = new SetupActions();
        GetFilter SetFilterParameters = new GetFilter();
        TableStorage FilterStorage = new TableStorage();
        MatchesCriteria MatchesCriterias = new MatchesCriteria();
        TransformCriteria TransformCriterias = new TransformCriteria();
        JsonParameters JsonParams = new JsonParameters();
        RowCountCollection MaxRowCount = new RowCountCollection();


        public QueryRefiner()
        {
            InitializeComponent();
            AccountForm.FormToShowOnClose = this;
            AccountForm.ShowDialog();
            this.Hide();
            return;
        }


        public void LoadData()
        {
            txtBUnow.Text = BusinessUnit;
            txtDOCUnow.Text = DocumentType;
            txtCurrency.Text = CurrencyType;
            txtCurrentBU.Text = BusinessUnit;
            if (BusinessUnit != "")
            {
                int BUtoQRConv = Convert.ToInt32(BusinessUnit);
                IncomeDictionary = new DictionaryInit();
                DictionaryFindTable(BUtoQRConv, IncomeDictionary.BUtoQRs);
                Restore_Init();
            }
            if (txtDOCUnow.Text == "Balance Sheet")
            {
                tableLayoutPanel2.Visible = false;
                groupBox2.Visible = false;
                btnMapper.Visible = false;
                flowChkBx.Visible = false;
                tableLayoutPanel3.Visible = false;
                groupBox5.Visible = true;
                btnBalances.Visible = true;
            }
            else
            {
                tableLayoutPanel2.Visible = true;
                groupBox2.Visible = true;
                btnMapper.Visible = true;
                flowChkBx.Visible = true;
                tableLayoutPanel3.Visible = true;
                groupBox5.Visible = false;
                btnBalances.Visible = false;
            }
            FilterShow();
            if (txtDOCUnow.Text == "Balance Sheet")
            {
                List<String> Lister = new List<String>() { "C", "E", "G", "H","I", "J", "K", "N", "O", "T" };
                cmbDictionList.DataSource = Lister;
            }
            else
            {
                List<String> Lister = new List<String>() { "Revenues", "Others", "Exceptions", "Projects","Income", "Expenses", "Departments", "Facilities", "Tech" };
                cmbDictionList.DataSource = Lister;
            }
            chkbxBalanceFilter.Checked = false;
            chkbxUseFilter.Checked = false;
            this.BringToFront();
            this.Focus();
        }

        //VFX: display state of records view
        public async void FilterShow()
        {
            lblFilters.Text = "Viewing first 1000 records of: " + txtCurrentBU.Text;
            lblFilters.Visible = true;
            await TaskDelay_5();
            await TaskDelay_5();
            lblFilters.Visible = false;
            lblFilters.Text = "Filter loaded to temp table";
        }

        private void Dashboard_Load (object sender, EventArgs e)
        {
            this.BringToFront();
            lblTotalRows.Text = "Total rows: " +GetRecordCount().ToString();
        }

        private int GetRecordCount()
        {
            Int32 count = 0;
            string sql = "SELECT COUNT(*) FROM "+ SetupAction.CurrentTable;
            using (OleDbConnection connection = new OleDbConnection(Con))
            {
                using (OleDbCommand cmd = new OleDbCommand(sql, connection))
                {
                    try
                    {
                        connection.Open();
                        count = (Int32)cmd.ExecuteScalar();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            return (int)count;
        }

        //QRY: Select top 1000 records, only on initialization
        public void Restore_Init()
        {
            try
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    using (OleDbTransaction Scope = connection.BeginTransaction(SD.IsolationLevel.ReadCommitted))
                    {
                        try
                        {
                            string QueryEntry = "SELECT TOP 1000 * FROM [" + SetupAction.CurrentTable + "]";
                            OleDbDataAdapter oda = new OleDbDataAdapter(QueryEntry, connection);
                            DataTable dt = new DataTable();
                            oda.SelectCommand.Transaction = Scope;
                            oda.Fill(dt);
                            Scope.Commit();
                            DGVMain.DataSource = dt;
                        }
                        catch (OleDbException odx)
                        {
                            MessageBox.Show(odx.Message);
                            Scope.Rollback();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("SQL error: " + ex);
            }
        }


        private void btnHome_Click(object sender, EventArgs e)
        {
            chkbxBalanceFilter.Checked = false;
            chkbxUseFilter.Checked = false;
            this.Hide();
            AccountForm.FormToShowOnClose = this;
            AccountForm.ShowDialog();
        }


        private void btnDB_Click(object sender, EventArgs e)
        {
            var OFD = new OpenFileDialog();
            OFD.Filter = "Excel Files|*.xls;*.xlsx";
            OFD.ShowDialog();
            Stopwatch sw = new Stopwatch();
            if (!string.IsNullOrEmpty(OFD.FileName))
            {
                SetupAction.IO_Name = OFD.FileName;
                cmbSheets.Items.Clear();
                EnablePanels("Loading... please wait!", false, true);
                Worker_Import.RunWorkerAsync();
            }
        }

        DataTable ImportTable = new DataTable();
        private void cmbSheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            ImportTable.Clear();
            EnablePanels("Loading... this may take a moment!", false, true);
            SetupAction.ExcelSheet = cmbSheets.Text;
            Thread.Sleep(3000);
            OleDbDataAdapter ODA = new OleDbDataAdapter("Select * from [" + cmbSheets.Text + "$]", Conn);
            
            try
            {
                ODA.Fill(ImportTable);
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            LoadWorker.RunWorkerAsync();
            //DGVExcel.DataSource = dt;
            tabControl.SelectedIndex = 1;
            lblItem.Text = DGVExcel.RowCount.ToString();
            CompleteProcess();
            cmbSheets.Text = "";
        }

        //VFX: for the status label
        private void Timerfx()
        {
            timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false;
        }
        
        //QRY: Refreshes the view with the full database
        private void Restore()
        {
            DGVMain.DataSource = null;
            string query = "SELECT * From [" + SetupAction.CurrentTable + "]";
            using (OleDbConnection connection = new OleDbConnection(Con))
            {
                connection.Open();
                try
                {
                    using (OleDbDataAdapter ODA = new OleDbDataAdapter(query, connection))
                    {
                        DataSet ds = new DataSet();
                        ODA.Fill(ds);
                        DGVMain.DataSource = ds.Tables[0];
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    EnablePanels("View now contains the full amount of rows", true, true);
                }
            }
        }
        
        //Initiates: [Transfer] worker
        private void btnTransfer_Click(object sender, EventArgs e)
        {
            if (DGVExcel.Rows.Count == 0)
            {
                timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false; ;
                panelColors.BackColor = Color.Red;
                MessageBox.Show("Excel data is missing!");
                lblStatus.Text = "Excel data = missing!";
                return;
            }
            if (!(Truth(rdbtnYYYY.Checked == true, rdbtnMDY.Checked == true, rdbtnOA.Checked == true) >= 1))
            {
                MessageBox.Show("Select a date type!");
                return;
            }
            DialogResult dr = MessageBox.Show(
                "Caution: imported date data types must match selection on the left pane, ensure you've selected the appropriate type!", "Import caution",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (dr == DialogResult.OK)
            {
                try
                {
                    Validator();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Validator Error: " + ex);
                }
                finally
                {
                    ClearTempTable();
                    GetDataTableFromDGV(DGVExcel);
                    MaxRowCount.MaxExcel = DGVExcel.Rows.Count;
                    EnablePanels("Loading... please wait!", false, true);
                    DGVExcel.DataSource = null;
                    Worker_Transfer.RunWorkerAsync();
                }
            }
        }
        
        //String builder function
        private void btnStrBuilder_Click(object sender, EventArgs e)
        {
            TextBox FocusedBox = FocusedTextBox();
            if (SetupAction.FocusedControl != null)
            {
                SetupAction.FocusedControl.Text = SetupAction.FocusedControl.Text + SetupAction.BonusSelectionBuilder;
            }
        }
        
        //Returns the textbox with focus
        private TextBox FocusedTextBox()
        {
            foreach (Control con in tableLayoutPanel1.Controls)
            {
                if (con.Focused == true)
                {
                    TextBox textbox = con as TextBox;
                    if (textbox != null)
                    {
                        return textbox;
                    }
                }
            }
            return null;
        }

        //Initiates: [Filter] worker
        private void btnFilter_Click(object sender, EventArgs e)
        {
            if(DGVMain.Rows.Count <= 0)
            {
                MessageBox.Show("Program did not detect any rows, returning function...");
                return;
            }
            if (SetupAction.OneTimeWarning == 0)
            {
                MessageBox.Show("Filters were reset. Note that filtering this way only generates a view.");
                SetupAction.OneTimeWarning++;
            }
            EnablePanels("Filtering... this may take a moment!", false, true);
            SetFilterParameters.currentLoops = 0;
            //Select all panels that are highlighted
            foreach (Control panel in tableLayoutPanel1.Controls)
            {
                if (panel is Panel)
                {
                    Panel Pane = panel as Panel;
                    if (panel.BackColor == Color.GreenYellow)
                    {
                        SetFilterParameters.currentLoops++;
                    }
                }
            }
            SetFilterParameters.MaxLoops = SetFilterParameters.currentLoops;
            SetFilterParameters.Acc = FilterAccount.Text;
            SetFilterParameters.Year = FilterYear.Text;
            SetFilterParameters.Period = FilterPeriod.Text;
            SetFilterParameters.Book = FilterBook.Text;
            SetFilterParameters.PHP = FilterPHP.Text;
            SetFilterParameters.USD = FilterUSD.Text;
            SetFilterParameters.Site = FilterSite.Text;
            SetFilterParameters.Status = FilterStatus.Text;
            SetFilterParameters.Dept = FilterDept.Text;
            SetFilterParameters.IS = FilterIS.Text;
            SetFilterParameters.ITR = FilterITR.Text;
            if (chkbxOR.Checked == true) { MatchesCriterias.HasOR = true; } else { MatchesCriterias.HasOR = false; }
            SetFilterParameters.FilterBuilder = new StringBuilder();
            SetFilterParameters.FilterBuilder.Append("Select * FROM ").Append(SetupAction.CurrentTable).Append(" WHERE ");
            MatchesCriterias.HasAccount = FilterParameters(panelAccount, FilterAccount, cmbAccount, 1, SetFilterParameters.Acc, "Account", "(@Account) ");
            MatchesCriterias.HasYear = FilterParameters(panelYear, FilterYear, cmbYear, 2, SetFilterParameters.Year, "Span_Year", "(@Year) ");
            MatchesCriterias.HasPeriod = FilterParameters(panelPeriod, FilterPeriod, cmbPeriod, 3, SetFilterParameters.Period, "Period", "(@Period) ");
            MatchesCriterias.HasBook = FilterParameters(panelBook, FilterBook, cmbBook, 4, SetFilterParameters.Book, "Book_Code", "(@Book) ");
            MatchesCriterias.HasPHP = FilterParameters(panelPHP, FilterPHP, cmbPHP, 5, SetFilterParameters.PHP, "Amount_PHP", "(@PHP) ");
            MatchesCriterias.HasUSD = FilterParameters(panelUSD, FilterUSD, cmbUSD, 6, SetFilterParameters.USD, "Amount_USD", "(@USD) ");
            MatchesCriterias.HasSite = FilterParameters(panelSite, FilterSite, cmbSite, 7, SetFilterParameters.Site, "Status_Name", "(@Site) ");
            MatchesCriterias.HasStatus = FilterParameters(panelStat, FilterStatus, cmbStatus, 8, SetFilterParameters.Status, "Status_Regime", "(@Status) ");
            MatchesCriterias.HasDept = FilterParameters(panelDept, FilterDept, cmbDept, 9, SetFilterParameters.Dept, "Department", "(@Dept) ");
            MatchesCriterias.HasIS = FilterParameters(panelIS, FilterIS, cmbIS, 10, SetFilterParameters.IS, "IS_Accounts", "(@IS) ");
            MatchesCriterias.HasITR = FilterParameters(panelITR, FilterITR, cmbITR, 11, SetFilterParameters.ITR, "Description_ITR", "(@ITR) ");
            ClearTempTable();
            Worker_Filter.RunWorkerAsync();
        }

        /// <summary>
        /// Assign a boolean that determines whether or not a particular filter is in place. If it is, prepare a special query when applicable. 
        /// </summary>
        /// <param name="panel">The selected panel</param>
        /// <param name="textBox">The box control for getting a value</param>
        /// <param name="dictionaryIndex">Indexer used in DictionaryFindFilter</param>
        /// <param name="replacer">TODO: Render value obsolete</param>
        /// <param name="query">Value used in overhead SQL command</param>
        /// <param name="parameters">Value used within [IN LIST] query</param>
        /// <returns>True if panel is selected to be filtered</returns>
        public Boolean FilterParameters(Panel panel, TextBox textBox, ComboBox comboBox, int dictionaryIndex, string replacer, string query, string parameters)
        {
            try
            {
                if (panel.BackColor == Color.GreenYellow)
                {
                    VerifyMissing(textBox, comboBox);
                    //[Between] Query function
                    if (comboBox.Text == "between" && panel == panelPeriod)
                    {
                        string[] BetweenCollection = FilterPeriod.Text.Split(',');
                        SetFilterParameters.PeriodStarter = BetweenCollection[0];
                        SetFilterParameters.PeriodEnder = BetweenCollection[1];
                        MatchesCriterias.HasPeriodBetween = true;
                        SetFilterParameters.FilterBuilder.Append("Period").Append(Operator(comboBox)).Append("@PeriodStarter AND @PeriodEnder ");
                    } 
                    else if (comboBox.Text == "between" && panel == panelYear)
                    {
                        string[] BetweenCollection = FilterYear.Text.Split(',');
                        SetFilterParameters.Starter = BetweenCollection[0];
                        SetFilterParameters.Ender = BetweenCollection[1];
                        MatchesCriterias.HasYearBetween = true;
                        SetFilterParameters.FilterBuilder.Append("Span_Year").Append(Operator(comboBox)).Append("@Starter AND @Ender ");
                    }
                    //If decimal type, return all decimal value matches of the given value including the given value
                    else if ((comboBox.Text == "decimal") && (panel == panelPHP))
                    {
                        Double[] decimalCollection = Array.ConvertAll<string, Double>(SetFilterParameters.PHP.Split(','), Convert.ToDouble);
                        StringBuilder queries = new StringBuilder();
                        for (int i = 0; i < decimalCollection.Count(); i++)
                        {
                            if (i > 0)
                            {
                                queries.Append(" OR ");
                            }
                            if (decimalCollection[i] >= 0)
                            {
                                queries.Append(" (Amount_PHP >= " + decimalCollection[i].ToString() + " AND Amount_PHP < " + (decimalCollection[i] + 1).ToString() + ")");
                            }
                            else
                            {
                                queries.Append(" (Amount_PHP <= " + decimalCollection[i].ToString() + " AND Amount_PHP > " + (decimalCollection[i] + 1).ToString() + ")");
                            }
                        }
                        string cap = queries.ToString();
                        SetFilterParameters.FilterBuilder.Append(cap);
                    } 
                    else if ((comboBox.Text == "decimal") && (panel == panelUSD))
                    {
                        Double[] decimalCollection = Array.ConvertAll<string, Double>(SetFilterParameters.USD.Split(','), Convert.ToDouble);
                        StringBuilder queries = new StringBuilder();
                        for (int i = 0; i < decimalCollection.Count(); i++)
                        {
                            if (i > 0)
                            {
                                queries.Append(" OR ");
                            }
                            if (decimalCollection[i] >= 0)
                            {
                                queries.Append(" (Amount_USD >= " + decimalCollection[i].ToString() + " AND Amount_USD < " + (decimalCollection[i] + 1).ToString() + ")");
                            }
                            else
                            {
                                queries.Append(" (Amount_USD <= " + decimalCollection[i].ToString() + " AND Amount_USD > " + (decimalCollection[i] + 1).ToString() + ")");
                            }
                        }
                        string cap = queries.ToString();
                        SetFilterParameters.FilterBuilder.Append(cap);
                    }
                    //[In List] Query Function
                    else
                    {
                        if (comboBox.Text == "in list")
                        {
                            parameters = "(" + FilterInList(textBox, panel) + ")";
                        }
                        SetFilterParameters.FilterBuilder.Append(query).Append(Operator(comboBox)).Append(parameters);
                    } 
                    SetFilterParameters.currentLoops--;
                    if (SetFilterParameters.currentLoops > 0 && SetFilterParameters.MaxLoops > 1) {
                        if (MatchesCriterias.HasOR != true) { SetFilterParameters.FilterBuilder.Append(" AND "); }
                            else { SetFilterParameters.FilterBuilder.Append(" OR "); } }
                    return true;
                }
                else return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            
        }

        //Algorithm to solve for [IN LIST] whitespaces
        public string FilterInList(TextBox comboco, Panel panel)
        {
            //hmmmm pointless...
            if ( panel == panelUSD || panel == panelPHP) { return String.Join(",", comboco.Text.Split(','));  }
            else { return String.Join(",", comboco.Text.Split(',').Select(x => "\'" + x + "\'"));  }
        }

        //Operator determines the appropriate query function for any given filter arrangement
        public string Operator(ComboBox Affected)
        {
            string Expression = " Null ";
            if (Affected.Text != string.Empty)
            {
                switch (Affected.Text)
                {
                    case "between":
                        Expression = " BETWEEN ";
                        break;
                    case "in list":
                        Expression = " IN ";
                        break;
                    case "equal to":
                        Expression = " = ";
                        break;
                    case "less than":
                        Expression = " < ";
                        break;
                    case "greater than":
                        Expression = " > ";
                        break;
                    case "like":
                        Expression = " LIKE ";
                        break;
                    case "decimal":
                        Expression = " ";
                        break;
                }
                return Expression;
            }
            return Expression;
        }

        //Verifies if data has correctly been entered into the filter builder
        public Boolean VerifyMissing(TextBox FoundValue, ComboBox ComboValue)
        {
            if (String.IsNullOrWhiteSpace(FoundValue.Text))
            {
                MessageBox.Show("Verified, criteria text is missing!");
                return true;
            }
            else if (String.IsNullOrWhiteSpace(ComboValue.Text))
            {
                MessageBox.Show("Verified, condition is missing!");
                return true;
            }
            else
            {
                return false;
            }
        }

        //Replaces whitespace with commas for use in filtering
        //TODO: Replace obsolete code
        public string FilterReplacer(string Input)
        {
            string[] FormattedCollection = Input.Split(' ');
            for (int i = 0; i < FormattedCollection.Length; i++)
            {
                string CurrentString = FormattedCollection.ElementAt(i);
                if (i != FormattedCollection.Length - 1)
                {
                    CurrentString += ",";
                    FormattedCollection.SetValue(CurrentString, i);
                }
            }
            string LastMove = "";
            foreach (var b in FormattedCollection)
            {
                LastMove = LastMove + string.Join("", b);
            }
            return LastMove;
        }

        //Replaces whitespace with AND for use in filtering
        //TODO: Replace obsolete code
        public string FilterSpacer(string Input)
        {
            string[] FormattedCollection = Input.Split(' ');
            for (int i = 0; i < FormattedCollection.Length; i++)
            {
                string CurrentString = FormattedCollection.ElementAt(i);
                if (i != FormattedCollection.Length - 1)
                {
                    CurrentString += " AND ";
                    FormattedCollection.SetValue(CurrentString, i);
                }
            }
            string LastMove = "";
            foreach (var b in FormattedCollection)
            {
                LastMove = LastMove + string.Join("", b);
            }
            return LastMove;
        }
        
        //Initiates: [IS] worker
        private void btnFixer_Click(object sender, EventArgs e)
        {
            if (SetupAction.OneTimeWarning == 0)
            {
                DialogResult dr = MessageBox.Show(
                "The auto-classify function uses whatever data is visible in the current view, proceed?", "Import caution",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (dr != DialogResult.OK)
                {
                    SetupAction.OneTimeWarning++;
                    return;
                }
            }
            
            if (DGVMain.Rows.Count == 0)
            {
                MessageBox.Show("No rows detected in the view");
                return;
            }
            ClearTempTable();
            GetDataTableFromDGV(DGVMain);
            EnablePanels("Classifying may take several minutes...", false, true);
            Worker_IS.RunWorkerAsync();
        }

        public async void EnablePanels(String status, Boolean b, Boolean lights)
        {
            tabControl.Enabled = b;
            panel1.Enabled = b;
            panel23.Enabled = b;
            cmbSheets.Enabled = b;

            if (b == true)
            {
                loadingSpinner.Visible = false;
                btnCancel.Enabled = false;
                if (lights == true)
                {
                    lblStatus.Text = status;
                    timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false;
                    panelColors.BackColor = Color.Green;
                }
                pbCheck.Visible = true;
                await TaskDelay_5();
                lblStatus.Text = "";
                pbCheck.Visible = false;
                panelColors.BackColor = Color.DarkGray;
            }
            else
            {
                if (lights == true)
                {
                    lblStatus.Text = status;
                    timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false;
                    panelColors.BackColor = Color.DarkGray;
                }
                loadingSpinner.Visible = true;
                btnCancel.Enabled = true;
            }
        }

        //Master function to pull data from data grid view into memory
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
            lblFilters.Visible = false;
            return VirtualTable;
        }

        /// <summary>
        /// Checks account data against the backend dictionary and updates the ITR classification
        /// </summary>
        /// <param name="mapKey">Current row's account value</param>
        /// <param name="mapCode">Current row's department value</param>
        /// <param name="rowIndex">Current index of for loop</param>
        /// <param name="accountLexicon">Specified dictionary to be searched</param>
        public void DictionaryUseAccount(int mapKey, int mapCode, int rowIndex, Dictionary<int, DictionarySetup> AccountLexicon)
        {
            if (AccountLexicon.TryGetValue(mapKey, out DictionarySetup ClassValues))
            {
                if (chkbxDesc.Checked == true)
                    VirtualTable.Rows[rowIndex][4] = ClassValues.theDescription;
                VirtualTable.Rows[rowIndex][22] = ClassValues.theClass;
            }
        }

        /// <summary>
        /// Checks the given key against the backened dictionary and returns a table name
        /// </summary>
        /// <param name="mapKey">Key is converted from literal BU code to table name</param>
        /// <param name="accountLexicon">Specified dictionary to be searched</param>
        public void DictionaryFindTable(int mapKey, Dictionary<int, DictionarySetup> accountLexicon)
        {
            if (accountLexicon.TryGetValue(mapKey, out DictionarySetup ClassValues))
            {
                SetupAction.CurrentTable = ClassValues.theDescription;
                if(txtDOCUnow.Text == "Balance Sheet")
                {
                    SetupAction.CurrentTable += "_BS";
                }
            }
        }

        /// <summary>
        /// Checks the given key against the backened dictionary and returns a nested dictionary [BS]
        /// </summary>
        /// <param name="mapKey">Key is converted from an internal value</param>
        /// <param name="accountLexicon">Specified nested dictionary to be searched</param>
        public Dictionary<int, DictionaryCheckup> Find_Dictionaries(string mapKey, Dictionary<string, Dictionary<int, DictionaryCheckup>> accountLexicon)
        {
            if (accountLexicon.TryGetValue(mapKey, out Dictionary<int, DictionaryCheckup> FoundDictionary))
            {
                return FoundDictionary;
            }
            return null;
        }

        /// <summary>
        /// Checks the given key against the backened dictionary and returns a nested dictionary [IS]
        /// </summary>
        /// <param name="mapKey">Key is converted from an internal value</param>
        /// <param name="accountLexicon">Specified nested dictionary to be searched</param>
        public Dictionary<int, DictionarySetup> Find_Dictionaries2(string mapKey, Dictionary<string, Dictionary<int, DictionarySetup>> accountLexicon)
        {
            if (accountLexicon.TryGetValue(mapKey, out Dictionary<int, DictionarySetup> FoundDictionary2))
            {
                return FoundDictionary2;
            }
            return null;
        }

        //Method overwrites the key value pairs of a dictionary with the pairs from a JSON file [BS]
        public void DictionaryReplaceBalance(Dictionary<int, DictionaryCheckup> accountLexicon, Dictionary<int, DictionaryCheckup> accountLexicon2)
        {
            if (accountLexicon != null)
            {
                accountLexicon.Clear();
                foreach (var kv in accountLexicon2)
                    if (!accountLexicon.ContainsKey(kv.Key))
                        accountLexicon.Add(kv.Key, kv.Value);
            }
        }

        //Method overwrites the key value pairs of a dictionary with the pairs from a JSON file [IS]
        public void DictionaryReplaceBalance2(Dictionary<int, DictionarySetup> accountLexicon, Dictionary<int, DictionarySetup> accountLexicon2)
        {
            if (accountLexicon != null)
            {
                accountLexicon.Clear();
                foreach (var kv in accountLexicon2)
                    if (!accountLexicon.ContainsKey(kv.Key))
                        accountLexicon.Add(kv.Key, kv.Value);
            }
        }

        //Check whether the department contains a value of direct within the dictionary
        public void DictionaryUseDepartment(int mapKey, int mapCode, int rowIndex, Dictionary<int, DictionarySetup> accountLexicon) //Compiler errors out
        {
            if (accountLexicon.TryGetValue(mapCode, out DictionarySetup ClassValues))
            {
                string foundValue = ClassValues.theClass.ToLower();
                if (foundValue == "direct")
                {
                    VirtualTable.Rows[rowIndex][22] = "Cost of Services";
                }
            }
        }
        

        public void DictionaryGetDepartmentDescription(int MapCode, int rowindex, Dictionary<int, DictionarySetup> AccountLexicon) //Compiler errors out
        {
            if (chkbxDesc.Checked != true)
            {
                return;
            }
            if (string.IsNullOrWhiteSpace(VirtualTable.Rows[rowindex][18].ToString()))
            {
                if (AccountLexicon.TryGetValue(MapCode, out DictionarySetup ClassValues))
                {
                    VirtualTable.Rows[rowindex][18] = ClassValues.theDescription;
                }
            }
        }

        //Utilizes the backend dictionary to search for a particular project match
        public void ProjectNumberSearch(int rowindex, int ProjectNumber)
        {
            if (IncomeDictionary.accountProjects.ContainsKey(ProjectNumber))
            {
                VirtualTable.Rows[rowindex][22] = "Operating Expenses";
            }
            else
            {
                VirtualTable.Rows[rowindex][22] = "Cost of Services";
            }
        }

        //Valudates cell values to ensure maximum compatibility
        //TODO: Remove obsolete code (Made obsolete by lambda/null coalescing operators)
        public void Validator()
        {
            List<string> ExcludedColumnsList = new List<string> { "ID", "Business_Unit", "Ledger", "Description", "Alt_Currency", "Base_Currency", "Site", "Stat_Site", "Status_Name", "Status_Regime", "Description_Department", "Description_Project", "Affiliate", "IS_Accounts", "Description_ITR" };
            //
            foreach (DataGridViewRow row in DGVExcel.Rows)
            {
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    if (!ExcludedColumnsList.Contains(DGVExcel.Columns[i].Name))
                    {
                        if ((DGVExcel.Columns[i].Name == "Span_Year") && (row.Cells[i].Value == null || row.Cells[i].Value == DBNull.Value ||
                                String.IsNullOrWhiteSpace(row.Cells[i].Value.ToString())))
                        {
                            row.Cells[i].Value = 1000;
                        }
                        if (row.Cells[i].Value == null || row.Cells[i].Value == DBNull.Value ||
                                String.IsNullOrWhiteSpace(row.Cells[i].Value.ToString()))
                        {
                            row.Cells[i].Value = 0;
                            //DGVExcel.RefreshEdit();
                        }
                    }
                }
            }
        }


        private void Query_Refiner_Disposed(object sender, EventArgs e)
        {
            try { Worker_IS.Dispose(); }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        //QRY: Updates records based on the changes to VirtualTable at the end of the [Transfer] worker.
        public void Recording(int rowindex)
    {
            using (OleDbCommand cmd = new OleDbCommand())
            {
                try
                {
                    using (OleDbConnection connection = new OleDbConnection(Con))
                    {
                        cmd.Connection = connection;
                        connection.Open();
                        using (OleDbTransaction Scope = connection.BeginTransaction(SD.IsolationLevel.ReadCommitted))
                        {
                            try
                            {
                                string Query = @"UPDATE [" + SetupAction.CurrentTable + "] set Description=@Description, Description_Department=@Description_Department, IS_Accounts=@IS_Accounts where ID=@ID";
                                cmd.Parameters.AddWithValue("@Description", VirtualTable.Rows[rowindex][4].ToString());
                                cmd.Parameters.AddWithValue("@Description_Department", VirtualTable.Rows[rowindex][18].ToString());
                                cmd.Parameters.AddWithValue("@IS_Accounts", VirtualTable.Rows[rowindex][22].ToString());
                                cmd.Parameters.AddWithValue("@ID", VirtualTable.Rows[rowindex][0].ToString());
                                cmd.CommandText = Query;
                                cmd.Transaction = Scope;
                                cmd.ExecuteNonQuery();
                                Scope.Commit();
                            }
                            catch (OleDbException odex)
                            {
                                MessageBox.Show(odex.Message);
                                Scope.Rollback();
                            }
                        }
                    }
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show("SQL: " + ex);
                }
            }
        }

        //QRY: Disposes all the rows of the current table.
        private void btnDispose_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show(
            "Warning: All data in the table view and Access file will be permanently deleted. The data will no longer be recoverable, do you wish to proceed?", "Disposing records",
            MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (dr == DialogResult.OK)
            {
                EnablePanels("Loading...", false, true);
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    try
                    {
                        connection.Open();
                        using (OleDbTransaction Scope = connection.BeginTransaction(SD.IsolationLevel.ReadCommitted))
                        {
                            try
                            {
                                OleDbCommand cmd = new OleDbCommand("DELETE FROM [" + SetupAction.CurrentTable + "]", connection);
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
                    finally
                    {
                        OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM [" + SetupAction.CurrentTable + "]", connection);
                        DataTable dt = new DataTable();
                        sda.Fill(dt);
                        DGVMain.DataSource = dt;
                        connection.Close();
                        DialogResult drs = MessageBox.Show(
            "Would you like to reload the view with the complete database?", "Reload?",
            MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                        if (drs == DialogResult.OK)
                        {
                            Restore();
                        }
                        CompleteProcess();
                    }
                }
            }
        }

        //VFX: Status bar effects on update.
        private void timer_Tick(object sender, EventArgs e)
        {
            SetupAction.TimerIntervals++;
            switch (SetupAction.TimerIntervals)
            {
                case 10:
                    panelColors.BackColor = Color.LightGray;
                    break;
            }
            if (SetupAction.TimerIntervals == 10)
            {
                SetupAction.TimerIntervals = 0;
                timer.Stop();
            }
            Util.Animate(panelColors, Util.Effect.Roll, 150, 360);
        }

        //QRY: Selects the full amount of rows from the database.
        private void btnReset_Click(object sender, EventArgs e)
        {
            EnablePanels("Loading... this may take a moment!", false, true);
            Restore();
            CompleteProcess();
        }

        //QRY: Clears the temporary table.
        public void ClearTempTable()
        {
            using (OleDbConnection connection = new OleDbConnection(Con))
            {
                connection.Open();
                using (OleDbTransaction Scope = connection.BeginTransaction(SD.IsolationLevel.ReadCommitted))
                {
                    try
                    {
                        OleDbCommand cmd = new OleDbCommand("DELETE FROM ACImport ", connection);
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
        }

        //Initiates: [Transforming] function.
        private void btnTransformer_Click(object sender, EventArgs e)
        {
            if(DGVMain.Rows.Count <= 0 || DGVMain.Rows == null)
            {
                MessageBox.Show("Invalid row count detected in the datagridview");
                return;
            }
            Transforming_IS();
        }

        /// <summary>
        /// Pivots the data visible in the grid view based on a set criteria
        /// </summary>
        /// <remarks>TODO: Attempt to simplify this rushed algorithm into a simple and understandable one</remarks>
        public void Transforming_IS()
        {
            StringBuilder Transformer = new StringBuilder();
            TransformCriteria CheckBoolean = new TransformCriteria();
            try
            {
                StringBuilder queryBuilder = new StringBuilder(); queryBuilder.Append(
                 " ORDER BY IIf([IS_Accounts] = 'Revenues' , 1 , IIf([IS_Accounts] = 'Cost of Services', 2 , IIf([IS_Accounts] = 'Operating Expenses', 3, IIf([IS_Accounts] = 'Other income/expense', 4, IIf([IS_Accounts] = 'Income Taxes', 5, 6))))) ASC ");
                if (chkbxBonus.Checked == true)
                {
                    TransformCriterias.QueryBuilder = new StringBuilder();
                    BonusSelectionsBuilder(comboBox1, chkbx1);
                    BonusSelectionsBuilder(comboBox2, chkbx2);
                    BonusSelectionsBuilder(comboBox3, chkbx3);
                    BonusSelectionsBuilder(comboBox4, chkbx4);
                    BonusSelectionsBuilder(comboBox5, chkbx5);
                    BonusSelectionsBuilder(comboBox6, chkbx6);
                    BonusSelectionsBuilder(comboBox7, chkbx7);
                    BonusSelectionsBuilder(comboBox8, chkbx8);
                    if (TransformCriterias.QueryBuilder != null)
                    {
                        if ((chkbxUseFilter.Checked == true) && (lblActivity.Text == "Inactive"))
                        {
                            MessageBox.Show("No filters are applied!");
                            return;
                        }
                        string FilterString = FilterReplacer(TransformCriterias.QueryBuilder.ToString());
                        char[] Char_Replacer = FilterReplacer(TransformCriterias.QueryBuilder.ToString()).ToCharArray();
                        TransformCriterias.FinalizedString = FilterString;
                        if (Char_Replacer[Char_Replacer.Length - 1] == ',')
                        {
                            TransformCriterias.FinalizedString = FilterString.Remove(FilterString.Length - 1, 1);
                        }
                        if ((chkbxUseFilter.Checked == true))
                        {
                            Transformer.Append("TRANSFORM ").Append(cmbAggregate.Text + "(" + txtData.Text + ")").Append(" SELECT DISTINCT " + TransformCriterias.FinalizedString + " FROM").Append("[ACImport]");
                        }
                        else Transformer.Append("TRANSFORM ").Append(cmbAggregate.Text + "(" + txtData.Text + ")").Append(" SELECT DISTINCT " + TransformCriterias.FinalizedString + " FROM ").Append("[" + SetupAction.CurrentTable + "]");
                        Transformer.Append(" GROUP BY " + TransformCriterias.FinalizedString + " PIVOT " + txtTopValues.Text + " ");
                    }
                    else
                    {
                        MessageBox.Show("Missing values detected in combo box!");
                        return;
                    }
                }
                else
                {
                    if (chkbxUseFilter.Checked == true)
                    {
                        Transformer.Append("TRANSFORM ").Append(cmbAggregate.Text + "(" + txtData.Text + ")").Append(" SELECT DISTINCT " + txtRowValues.Text + " FROM ").Append("[ACImport]");
                    }
                    else Transformer.Append("TRANSFORM ").Append(cmbAggregate.Text + "(" + txtData.Text + ")").Append(" SELECT DISTINCT " + txtRowValues.Text + " FROM ").Append("[" + SetupAction.CurrentTable + "]");
                    if (txtGroupBy.Enabled == false) { Transformer.Append(" GROUP BY " + txtRowValues.Text + " PIVOT " + txtTopValues.Text + " "); }
                    else if (txtGroupBy.Enabled == true) { Transformer.Append(" GROUP BY " + txtGroupBy.Text + " PIVOT " + txtTopValues.Text + " "); }
                }
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    OleDbDataAdapter da = new OleDbDataAdapter(Transformer.ToString(), Con);
                    Console.WriteLine(Transformer.ToString());
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    DGVMain.DataSource = dt;
                    connection.Close();
                    GetDataTableFromDGV(DGVMain);
                }
                List<string> strong = new List<string>();
                foreach (DataGridViewColumn col in DGVMain.Columns)
                {
                    strong.Add(col.Name);
                }
                DataColumn c = VirtualTable.Columns.Add("Sort", typeof(int));
                c.Expression = "IIf([Column1] = 'Revenues' , 1 , IIf([Column1] = 'Cost of Services', 2 , IIf([Column1] = 'Operating Expenses', 3, IIf([Column1] = 'Other income/expense', 4, IIf([Column1] = 'Income Taxes', 5, 6)))))"; //"iif(Column1='Revenues', 0, iif(Column1='Cost of Services', 1, 2))";
                DataView sorted = VirtualTable.ApplySort((r, r2) =>
                {
                    return ((int)r["Sort"]).CompareTo(((int)r2["Sort"]));
                });
                VirtualTable = sorted.ToTable();
                VirtualTable.Columns.RemoveAt(VirtualTable.Columns.Count - 1);
                if (chkbxISCheck.Checked == true)
                {
                    int MaxVirtualc = VirtualTable.Columns.Count;
                    int MaxVirtualr = VirtualTable.Rows.Count;
                    DataRow GrossProfits = VirtualTable.NewRow();
                    DataRow NetProfits = VirtualTable.NewRow();
                    DataRow GrandBottom = VirtualTable.NewRow();
                    int indexOfRevenue, indexOfCost, indexOfOperations, indexOfOthers, indexOfIncome;
                    indexOfRevenue = indexOfCost = indexOfOperations = indexOfOthers = indexOfIncome = 0;
                    if (FindRowValues("revenues", "revenue").Item1 == true)
                    {
                        CheckBoolean.HasRevenues = true;
                        indexOfRevenue = FindRowValues("revenues", "revenue").Item2;
                    }
                    if (FindRowValues("cost of services", "cost of service").Item1 == true)
                    {
                        CheckBoolean.HasCost = true;
                        indexOfCost = FindRowValues("cost of services", "cost of service").Item2;
                    }
                    if (FindRowValues("operating expenses", "operating expense").Item1 == true)
                    {
                        CheckBoolean.HasOperations = true;
                        indexOfOperations = FindRowValues("operating expenses", "operating expense").Item2;
                    }
                    if (FindRowValues("other income or expense", "other income/expense").Item1 == true)
                    {
                        CheckBoolean.HasOther = true;
                        indexOfOthers = FindRowValues("other income or expense", "other income/expense").Item2;
                    }
                    //Uncomment if income row is desired
                    //if (FindRowValues("income taxes", "income tax").Item1 == true)
                    //{
                    //    hasIncome = true;
                    //    indexOfIncome = FindRowValues("income taxes", "income tax").Item2;
                    //}
                    if (!(Truth(CheckBoolean.HasRevenues, CheckBoolean.HasCost, CheckBoolean.HasOperations, CheckBoolean.HasOther) >= 1))
                    {
                        MessageBox.Show("Data mismatch! IS class possibly missing! If the IS Account column moved to the far left, try refreshing the database!");
                        return;
                    }
                    List<Double> itemArrayRevenue = new List<Double>(); List<Double> itemArrayCost = new List<Double>(); List<Double> itemArrayOperating = new List<Double>(); List<Double> itemArrayOther = new List<Double>(); List<Double> itemArrayIncome = new List<Double>();
                    Double Valueof_Revenue, Valueof_Cost, Valueof_Operating, Valueof_Other, Valueof_Income;
                    Valueof_Revenue = Valueof_Cost = Valueof_Operating = Valueof_Other = Valueof_Income = 0;
                    for (int cols = 1; cols < MaxVirtualc; cols++)
                    {
                        List<Double> itemArrayGrandBottom = new List<Double>();
                        //Run through column, grab values
                        if (CheckBoolean.HasRevenues == true)
                        {
                            Valueof_Revenue = VirtualTable.Rows[indexOfRevenue][cols] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[indexOfRevenue][cols]);
                        }
                        if (CheckBoolean.HasCost == true)
                        {
                            Valueof_Cost = VirtualTable.Rows[indexOfCost][cols] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[indexOfCost][cols]);
                        }
                        if (CheckBoolean.HasOperations == true)
                        {
                            Valueof_Operating = VirtualTable.Rows[indexOfOperations][cols] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[indexOfOperations][cols]);
                        }
                        if (CheckBoolean.HasOther == true)
                        {
                            Valueof_Other = VirtualTable.Rows[indexOfOthers][cols] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[indexOfOthers][cols]);
                        }
                        if (CheckBoolean.HasIncome == true)
                        {
                            Valueof_Income = VirtualTable.Rows[indexOfIncome][cols] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[indexOfIncome][cols]);
                        }
                        //New row values
                        Valueof_Revenue = (Valueof_Revenue) * -1;
                        Double GrossProfit = Valueof_Revenue - Valueof_Cost;
                        Valueof_Other = (Valueof_Other) * -1;
                        Double NetProfit = Valueof_Operating - Valueof_Other;
                        //Gathering total values for last column+1
                        itemArrayRevenue.Add(Valueof_Revenue);
                        itemArrayCost.Add(Valueof_Cost);
                        itemArrayOperating.Add(Valueof_Operating);
                        itemArrayOther.Add(Valueof_Other);
                        //itemArrayIncome.Add(Valueof_Income);
                        // Gathering total values for each column bottom row
                        itemArrayGrandBottom.Add(Valueof_Revenue);
                        itemArrayGrandBottom.Add(Valueof_Cost);
                        itemArrayGrandBottom.Add(Valueof_Operating);
                        itemArrayGrandBottom.Add(Valueof_Other);
                        //itemArrayGrandBottom.Add(Valueof_Income);
                        Double GrandDouble = itemArrayGrandBottom.Sum();
                        //Data added to datarow, but not yet added to table
                        GrandBottom[0] = "[Grand Total]";
                        GrandBottom[cols] = GrandDouble;
                        GrossProfits[0] = "[Gross Profit]";
                        NetProfits[0] = "[Total Profit]";
                        GrossProfits[cols] = GrossProfit;
                        NetProfits[cols] = NetProfit;
                    } //Compute rows values
                    // Values of Grand Total
                    Double revenueValue = itemArrayRevenue.Sum();
                    Double costValue = itemArrayCost.Sum();
                    Double operationValue = itemArrayOperating.Sum();
                    Double otherValue = itemArrayOther.Sum();
                    //Double IncomeValue = itemArrayIncome.Sum();
                    //Appends Grand Total Column
                    if (chkbWithout.Checked == true)
                    {
                        DataColumn dc = new DataColumn("Grand Total", typeof(Double));
                        VirtualTable.Columns.Add(dc);
                        // Grand Total added
                        if (CheckBoolean.HasRevenues == true)
                            VirtualTable.Rows[indexOfRevenue][VirtualTable.Columns.Count - 1] = revenueValue;
                        if (CheckBoolean.HasCost == true)
                            VirtualTable.Rows[indexOfCost][VirtualTable.Columns.Count - 1] = costValue;
                        if (CheckBoolean.HasOperations == true)
                            VirtualTable.Rows[indexOfOperations][VirtualTable.Columns.Count - 1] = operationValue;
                        if (CheckBoolean.HasOther == true)
                            VirtualTable.Rows[indexOfOthers][VirtualTable.Columns.Count - 1] = otherValue;
                      //Uncomment if income row is desired.
                      //if (hasIncome == true)
                      //    VirtualTable.Rows[indexOfIncome][VirtualTable.Columns.Count - 1] = IncomeValue;
                    }
                    //Rows added
                    VirtualTable.Rows.InsertAt(GrossProfits, 2);
                    VirtualTable.Rows.InsertAt(NetProfits, MaxVirtualc);
                    //Cloning
                    DataTable VirtualTableClone = VirtualTable.Clone();
                    for (int i = 1; i < VirtualTable.Columns.Count; i++)
                    {
                        VirtualTableClone.Columns[i].DataType = typeof(Double);
                    }
                    foreach (DataRow row in VirtualTable.Rows)
                    {
                        VirtualTableClone.ImportRow(row);
                    }
                    VirtualTable.Clear();
                    VirtualTable = VirtualTableClone;
                    //Cloning End
                    DGVMain.DataSource = VirtualTable;
                    //Copy Headers
                    if (chkbWithout.Checked == true)
                    {
                        for (int i = 0; i < DGVMain.ColumnCount - 1; i++)
                        {
                            DGVMain.Columns[i].HeaderText = strong[i];
                            if (i == 0)
                            {
                                DGVMain.Columns[i + 1].DefaultCellStyle.Format = "0.00##";
                            }
                            else
                            {
                                DGVMain.Columns[i].DefaultCellStyle.Format = "0.00##";
                            }

                        }
                    } //Compute + Pivot
                    else
                    {
                        for (int i = 0; i < DGVMain.ColumnCount; i++)
                        {
                            DGVMain.Columns[i].HeaderText = strong[i];
                            if (i == 0)
                            {
                                DGVMain.Columns[i + 1].DefaultCellStyle.Format = "0.00##";
                            }
                            else
                            {
                                DGVMain.Columns[i].DefaultCellStyle.Format = "0.00##";
                            }

                        }
                    } //Pivot only
                    DGVMain.Refresh();
                }
                else
                {
                    DGVMain.DataSource = VirtualTable;
                    for (int i = 0; i < DGVMain.ColumnCount; i++)
                    {
                        DGVMain.Columns[i].HeaderText = strong[i];
                        if (i == 0)
                        {
                            DGVMain.Columns[i + 1].DefaultCellStyle.Format = "0.00##";
                        }
                        else
                        {
                            DGVMain.Columns[i].DefaultCellStyle.Format = "0.00##";
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Pivot attempt failed: " + ex);
            }
            finally
            {
                //Probably unnecessary, principle of least surprise?
                CheckBoolean.HasCost = false;
                CheckBoolean.HasIncome = false;
                CheckBoolean.HasOperations = false;
                CheckBoolean.HasOther = false;
                CheckBoolean.HasRevenues = false;
                CompleteProcess();
            }
        }

        //Simple algorithm to check if any given set of booleans returns a desired number of true answers.
        public static int Truth(params bool[] booleans)
        {
            return booleans.Count(b => b);
        }

        //Bit of Linq here that interacts with the pivoted data through a tuple, returning a Boolean and an int.
        public Tuple<Boolean, int> FindRowValues(string value1, string value2)
        {
            var row = VirtualTable
                .AsEnumerable()
                .Select((tt, index) => new
                {
                    value = tt.Field<string>("Column1"),
                    index = index
                })
                .FirstOrDefault(item =>
                 string.Equals(item.value, value1, StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(item.value, value2, StringComparison.OrdinalIgnoreCase));
            if (row != null)
            {
                var r = row.value;              
                int selectedIndex = row.index;  
                return Tuple.Create(true, selectedIndex);
            }
            else
            {
                return Tuple.Create(false, 0);
            }
        }


        public void BonusSelectionsBuilder(ComboBox combobox, CheckBox checkbox)
        {
            if (checkbox.Checked == true && combobox != null)
            {
                TransformCriterias.QueryBuilder.Append(combobox.Text + " ");
            }
        }

        private void DGVMain_SelectionChanged(object sender, EventArgs e)
        {
            if (DGVMain.RowCount > 0 && DGVMain.SelectedRows.Count > 0)
            {
                lblItem.Text = DGVMain.SelectedRows.Count + " : " + DGVMain.RowCount.ToString();
            }
        }

        //Check if background worker is doing anything and send a cancellation if it is.
        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (Worker_IS.IsBusy)
            {
                Worker_IS.CancelAsync();
            }
            else if (Worker_Import.IsBusy)
            {
                Worker_Import.CancelAsync();
            }
            else if (Worker_Transfer.IsBusy)
            {
                Worker_Transfer.CancelAsync();
            }
            else if (Worker_Filter.IsBusy)
            {
                Worker_Filter.CancelAsync();
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (txtGroupBy.Enabled == true)
            {
                txtGroupBy.Enabled = false;
            }
            else { txtGroupBy.Enabled = true; }
        }

        /// <summary>
        /// Master classify function, produces ITR descriptions based on the values of Account and Deparment
        /// </summary>
        /// <remarks>TODO: Greatly simplify background worker classes to make them self explanatory without comments and; make use of diverse classes</remarks>
        private void Worker_IS_DoWork(object sender, DoWorkEventArgs e)
        {
            IncomeDictionary = new DictionaryInit();
            MaxRowCount.MaxRows = VirtualTable.Rows.Count;
            Stopwatch sw = new Stopwatch();
            sw.Start();
            try
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    string query = @"UPDATE [" + SetupAction.CurrentTable + "] set Description=@Description, Description_Department=@Description_Department, IS_Accounts=@IS_Accounts where ID=@ID";
                    cmd.CommandText = query;
                    using (OleDbConnection connection = new OleDbConnection(Con))
                    {
                        cmd.Connection = connection;
                        connection.Open();
                        for (int rowindex = 0; rowindex < MaxRowCount.MaxRows; rowindex++)
                        {
                            cmd.Parameters.Clear();
                            int currentPoint = rowindex;
                            int accountCode = Convert.ToInt32(VirtualTable.Rows[rowindex][3]);
                            int departmentCode = Math.Abs(Convert.ToInt32(VirtualTable.Rows[rowindex][17]));
                            int projectCode = Math.Abs(Convert.ToInt32(VirtualTable.Rows[rowindex][19]));
                            int absoluteKey = Math.Abs(accountCode);
                            int absoluteCode = departmentCode;
                            while (absoluteKey >= 10) { absoluteKey /= 10; }
                            while (absoluteCode >= 10) { absoluteCode /= 10; }
                            ///<summary> Algorithm to determine the IS classification.</summary>
                            ///<param="AbsoluteKey"> Store the account key, reduce it to the first digit for checking. </param>
                            ///<param="AbsoluteCode"> Store the department key, reduce it to the first digit for checking. </param>
                            ///<remarks> In order to best accomodate the general rule, any blank cases are immediately considered to be of mixed value. </remarks>
                            if (absoluteKey == 4)
                            {
                                VirtualTable.Rows[rowindex][22] = "Revenues";
                                DictionaryUseAccount(accountCode, departmentCode, rowindex, IncomeDictionary.MasterClassificationDictionary);
                            }
                            else
                            {
                                DictionaryUseAccount(accountCode, departmentCode, rowindex, IncomeDictionary.MasterClassificationDictionary);
                            }
                            string currentState = VirtualTable.Rows[rowindex][22].ToString().ToLower();
                            if (currentState == null || currentState == "" || currentState == " ") { VirtualTable.Rows[rowindex][22] = "mixed"; currentState = "mixed"; }
                            switch (currentState)
                            {
                                case "direct":
                                    VirtualTable.Rows[rowindex][22] = "Cost of Services";
                                    break;
                                case "sga":
                                    VirtualTable.Rows[rowindex][22] = "Operating Expenses";
                                    break;
                                case "mixed":
                                    if (IncomeDictionary.accountExceptions.ContainsKey(departmentCode))
                                    {
                                        string currentState_PostException = "";
                                        //Technology
                                        if (departmentCode == 50070 || departmentCode == 50077 || departmentCode == 50079 || departmentCode == 50050)
                                        {
                                            if (!(IncomeDictionary.accountTech.ContainsKey(accountCode)))
                                            {
                                                VirtualTable.Rows[rowindex][22] = "mixed";
                                            }
                                            else
                                            {
                                                DictionaryUseAccount(accountCode, departmentCode, rowindex, IncomeDictionary.accountTech);
                                            }
                                            currentState_PostException = VirtualTable.Rows[rowindex][22].ToString().ToLower();
                                            switch (currentState_PostException)
                                            {
                                                case "direct":
                                                    VirtualTable.Rows[rowindex][22] = "Cost of Services";
                                                    break;
                                                case "sga":
                                                    VirtualTable.Rows[rowindex][22] = "Operating Expenses";
                                                    break;
                                                case "mixed":
                                                    ProjectNumberSearch(rowindex, projectCode);
                                                    break;
                                            }
                                        }
                                        //Facilities
                                        else
                                        {
                                            Console.WriteLine(accountCode + " " + departmentCode + " " + projectCode);
                                            if (!(IncomeDictionary.accountFacilities.ContainsKey(accountCode)))
                                            {
                                                VirtualTable.Rows[rowindex][22] = "mixed";
                                            }
                                            else
                                            {
                                                DictionaryUseAccount(accountCode, departmentCode, rowindex, IncomeDictionary.accountFacilities);
                                            }
                                            currentState_PostException = VirtualTable.Rows[rowindex][22].ToString().ToLower();
                                            Console.WriteLine(currentState_PostException);
                                            switch (currentState_PostException)
                                            {
                                                case "direct":
                                                    VirtualTable.Rows[rowindex][22] = "Cost of Services";
                                                    break;
                                                case "sga":
                                                    VirtualTable.Rows[rowindex][22] = "Operating Expenses";
                                                    break;
                                                case "mixed":
                                                    ProjectNumberSearch(rowindex, projectCode);
                                                    break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        switch (absoluteCode)
                                        {
                                            case 1:
                                            case 2:
                                            case 3:
                                                VirtualTable.Rows[rowindex][22] = "Cost of Services";
                                                break;
                                            default:
                                                VirtualTable.Rows[rowindex][22] = "Operating Expenses";
                                                break;
                                        }
                                        //A department class of direct overwrites any other qualification
                                        DictionaryUseDepartment(accountCode, departmentCode, rowindex, IncomeDictionary.accountDepartments);
                                        string StateOfClass_NonEx = VirtualTable.Rows[rowindex][22].ToString().ToLower();
                                        if (StateOfClass_NonEx == "direct")
                                            VirtualTable.Rows[rowindex][22] = "Cost of Services";
                                        else if (StateOfClass_NonEx == "sga")
                                            VirtualTable.Rows[rowindex][22] = "Operating Expenses";
                                    }
                                    break;
                            }
                            //After Exceptions, begin recording update
                            DictionaryGetDepartmentDescription(departmentCode, rowindex, IncomeDictionary.accountDepartments);
                            cmd.Parameters.AddWithValue("@Description", VirtualTable.Rows[rowindex][4].ToString());
                            cmd.Parameters.AddWithValue("@Description_Department", VirtualTable.Rows[rowindex][18].ToString());
                            cmd.Parameters.AddWithValue("@IS_Accounts", VirtualTable.Rows[rowindex][22].ToString());
                            cmd.Parameters.AddWithValue("@ID", VirtualTable.Rows[rowindex][0].ToString());
                            try
                            {
                                cmd.ExecuteNonQuery();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                            lblStatus.Invoke((MethodInvoker)delegate
                            {
                                lblStatus.Text = "Processing row: " + currentPoint + "/" + DGVMain.RowCount;
                            });
                            Worker_IS.ReportProgress(rowindex);
                            //Check if there is a request to cancel the process
                            if (Worker_IS.CancellationPending)
                            {
                                e.Cancel = true;
                                Worker_IS.ReportProgress(0);
                                return;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Worker_IS.ReportProgress(MaxRowCount.MaxRows);
                VirtualTable.Clear();
                VirtualTable.Dispose();
                sw.Stop();
                FilterStorage.Occurrence = sw.Elapsed.ToString();
            }
        }
        private void Worker_IS_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pBarMain.Value = (e.ProgressPercentage * 100) / MaxRowCount.MaxRows;
        }
        private async void Worker_IS_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Yellow;
                timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false;
                panelColors.BackColor = Color.Yellow;
                lblStatus.Text = "Process was cancelled";
            }
            else if (e.Error != null)
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Red;
                timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false;
                panelColors.BackColor = Color.Red;
                lblStatus.Text = "Error: The thread aborted";
            }
            else
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Lime;
                timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false;
                panelColors.BackColor = Color.Green;
                lblStatus.Text = "Process has completed";
            }
            string query = "SELECT * From [" + SetupAction.CurrentTable + "]";
            using (OleDbConnection conn = new OleDbConnection(Con))
            {
                conn.Open();
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn))
                {
                    DataSet ds = new DataSet();
                    adapter.Fill(ds);
                    DGVMain.DataSource = ds.Tables[0];
                }
            }
            lblElapsedTime.Text = FilterStorage.Occurrence;
            EnablePanels("", true, false);
            await TaskDelay_5();
            pBarMain.Value = 0;
            pBarMain.Style = MetroColorStyle.Default;
            ClearTempTable();
        }

        async Task TaskDelay_5()
        {
            await Task.Delay(5000);
        }

        private void chkbxBonus_CheckedChanged(object sender, EventArgs e)
        {
            if (flowChkBx.Enabled == true)
            {
                flowChkBx.Enabled = false;
                tableLayoutPanel3.Enabled = false;
                txtRowValues.Enabled = true;
                txtGroupBy.Enabled = false;
                checkBox2.Enabled = true;
                lblStatus.Text = "Bonus statements disabled!";
            }
            else
            {
                flowChkBx.Enabled = true;
                tableLayoutPanel3.Enabled = true;
                txtRowValues.Enabled = false;
                txtGroupBy.Enabled = false;
                checkBox2.Enabled = false;
                tp2.Focus();
                lblStatus.Text = "Bonus statements enabled!";
            }
        }
        
        private void btnRelocate_Click(object sender, EventArgs e)
        {
            UseFilter();
        }

        //Initiates: [Worker_Transfer_Temp] worker
        public void UseFilter()
        {
            ClearTempTable();
            GetDataTableFromDGV(DGVMain);
            EnablePanels("Loading filter into view!", false, true);
            MaxRowCount.MaxRelocate = VirtualTable.Rows.Count;
            Worker_Transfer_Temp.RunWorkerAsync();
        }

        /// <summary>
        /// Master manual filter, uses the current view of limited rows by transferring them to a temporary table that is only used as a view
        /// </summary>
        /// <remarks>TODO: There's probably a simpler way to do this. Through SQL, I can probably create a temp table using an internal function.</remarks>
        private void Worker_Transfer_Temp_DoWork(object sender, DoWorkEventArgs e)
        {
            if (DGVMain.Rows.Count <= 0)
            {
                MessageBox.Show("Invalid row count! Rows are less than or equal to 0");
                return;
            }
            try
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    using (OleDbTransaction Scope = connection.BeginTransaction(SD.IsolationLevel.ReadCommitted))
                    {
                        using (OleDbCommand cmd = new OleDbCommand())
                        {
                            try
                            {
                                cmd.Connection = connection;
                                cmd.Transaction = Scope;
                                cmd.CommandText =
                                "Insert INTO ACImport (Business_Unit,Ledger,Account,Description,Span_Year,Period,Book_Code,Alt_Currency,Amount_USD,Base_Currency,PDS,Amount_PHP,Site,Stat_Site,Status_Name,Status_Regime,Department,Description_Department,Project_Number,Description_Project,Affiliate,IS_Accounts,Description_ITR) " +
                                                              "VALUES(@Business_Unit,@Ledger,@Account,@Description,@Span_Year,@Period,@Book_Code,@Alt_Currency,@Amount_USD,@Base_Currency,@PDS,@Amount_PHP,@Site,@Stat_Site,@Status_Name,@Status_Regime,@Department,@Description_Department,@Project_Number,@Description_Project,@Affiliate,@IS_Accounts,@Description_ITR)";
                                for (int rowindex = 0; rowindex < MaxRowCount.MaxRelocate; rowindex++)
                                {
                                    cmd.Parameters.Clear();
                                    cmd.Parameters.AddWithValue("@Business_Unit", VirtualTable.Rows[rowindex][1]);
                                    cmd.Parameters.AddWithValue("@Ledger", VirtualTable.Rows[rowindex][2]);
                                    cmd.Parameters.AddWithValue("@Account", VirtualTable.Rows[rowindex][3]);
                                    cmd.Parameters.AddWithValue("@Description", VirtualTable.Rows[rowindex][4]);
                                    //cmd.Parameters.AddWithValue("@Span_Year", (VirtualTable.Rows[rowindex][5] == DBNull.Value) ? VirtualTable.Rows[rowindex][5] = 1000 : new DateTime(Convert.ToInt32(VirtualTable.Rows[rowindex][5]), 1, 1));
                                    cmd.Parameters.AddWithValue("@Span_Year", (VirtualTable.Rows[rowindex][5]));
                                    cmd.Parameters.AddWithValue("@Period", VirtualTable.Rows[rowindex][6]);
                                    cmd.Parameters.AddWithValue("@Book_Code", VirtualTable.Rows[rowindex][7]);
                                    cmd.Parameters.AddWithValue("@Currency", VirtualTable.Rows[rowindex][8]);
                                    cmd.Parameters.AddWithValue("@Amount_USD", (VirtualTable.Rows[rowindex][9] == DBNull.Value) ? VirtualTable.Rows[rowindex][9] = 0 : Convert.ToDouble(VirtualTable.Rows[rowindex][9]));
                                    cmd.Parameters.AddWithValue("@Base_Currency", VirtualTable.Rows[rowindex][10]);
                                    cmd.Parameters.AddWithValue("@PDS", VirtualTable.Rows[rowindex][11]);
                                    cmd.Parameters.AddWithValue("@Amount_PHP", (VirtualTable.Rows[rowindex][12] == DBNull.Value) ? VirtualTable.Rows[rowindex][12] = 0 : Convert.ToDouble(VirtualTable.Rows[rowindex][12]));
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
                                    Worker_Transfer_Temp.ReportProgress(rowindex);
                                }
                                Scope.Commit();
                            }
                            catch (OleDbException odx)
                            {
                                MessageBox.Show(odx.Message);
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
            Worker_Transfer_Temp.ReportProgress(MaxRowCount.MaxRelocate);
        }
        private void Worker_Transfer_Temp_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pBarMain.Value = (e.ProgressPercentage * 100) / MaxRowCount.MaxRelocate;
        }
        private void Worker_Transfer_Temp_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Red;
                timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false;
                panelColors.BackColor = Color.Red;
                lblStatus.Text = "Error: The thread aborted";
            }
            else
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Lime;
                timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false;
                panelColors.BackColor = Color.Green;
                lblStatus.Text = "Filter loaded successfully!";
            }
            EnablePanels("", true, false);
            pBarMain.Value = 0;
            pBarMain.Style = MetroColorStyle.Default;
        }
        
        /// <summary>
        /// Downloads the selected dictionary into JSON formatting
        /// </summary>
        /// <remarks>TODO: Refactor dictionary system completely to permit hot swapping of dictionary data (currently, an impossible feat due to referencing)</remarks>
        private void btnDownloadJson_Click(object sender, EventArgs e)
        {
            var Import_Balance = new Dictionary<int, DictionaryCheckup>();
            var Import_IS = new Dictionary<int, DictionarySetup>();
            if (txtDOCUnow.Text == "Balance Sheet")
            {
                var OverDiction = new Dictionary<string, Dictionary<int, DictionaryCheckup>>();
                var C = BalanceDictionary.C;
                var E = BalanceDictionary.E;
                var G = BalanceDictionary.G;
                var H = BalanceDictionary.H;
                var I = BalanceDictionary.I;
                var J = BalanceDictionary.J;
                var K = BalanceDictionary.K;
                var L = BalanceDictionary.L;
                var N = BalanceDictionary.N;
                var O = BalanceDictionary.O;
                var T = BalanceDictionary.T;
                OverDiction.Add("C", C);
                OverDiction.Add("E", E);
                OverDiction.Add("G", G);
                OverDiction.Add("H", H);
                OverDiction.Add("I", I);
                OverDiction.Add("J", J);
                OverDiction.Add("K", K);
                OverDiction.Add("L", L);
                OverDiction.Add("N", N);
                OverDiction.Add("O", O);
                OverDiction.Add("T", T);
                Import_Balance = Find_Dictionaries(JsonParams.SelectedDictionary, OverDiction);
            }
            else
            {
                var OverDiction = new Dictionary<string, Dictionary<int, DictionarySetup>>();
                var C = IncomeDictionary.accountRevenue;
                var E = IncomeDictionary.accountOthers;
                var G = IncomeDictionary.accountExceptions;
                var H = IncomeDictionary.accountProjects;
                var I = IncomeDictionary.accountIncome;
                var J = IncomeDictionary.accountExpenses;
                var K = IncomeDictionary.accountDepartments;
                var L = IncomeDictionary.accountFacilities;
                var N = IncomeDictionary.accountTech;
                OverDiction.Add("Revenues", C);
                OverDiction.Add("Others", E);
                OverDiction.Add("Exceptions", G);
                OverDiction.Add("Projects", H);
                OverDiction.Add("Income", I);
                OverDiction.Add("Expenses", J);
                OverDiction.Add("Departments", K);
                OverDiction.Add("Facilities", L);
                OverDiction.Add("Tech", N);
                Import_IS = Find_Dictionaries2(JsonParams.SelectedDictionary, OverDiction);
            }
            DialogResult DR = SFD.ShowDialog();
            if (!(DR == DialogResult.OK))
            {
                return;
            }
            if (!File.Exists(SFD.FileName))
            {
                using (var stream = new StreamWriter(File.Create(SFD.FileName)))
                {
                    if (txtDOCUnow.Text == "Balance Sheet")
                    {
                        var json = JsonConvert.SerializeObject(Import_Balance, Newtonsoft.Json.Formatting.Indented);
                        stream.Write(json);
                    }
                    else
                    {
                        var json = JsonConvert.SerializeObject(Import_IS, Newtonsoft.Json.Formatting.Indented);
                        stream.Write(json);
                    }
                    stream.Flush();
                }
            }
        }

        //QRY: Delete where ID = @ID
        private void btnDeleteInfo_Click(object sender, EventArgs e)
        {
            if (txtSelectID.Text != "" || txtSelectID.Text != " ")
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    try
                    {
                        connection.Open();
                        using (OleDbTransaction Scope = connection.BeginTransaction(SD.IsolationLevel.ReadCommitted))
                        {
                            try
                            {
                                OleDbCommand cmd = new OleDbCommand("DELETE FROM [" + SetupAction.CurrentTable + "] WHERE ID = @ID", connection);
                                cmd.Parameters.AddWithValue("@ID", txtSelectID.Text);
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
                        MessageBox.Show("SQL error" + es);
                    }
                    finally
                    {
                        OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM [" + SetupAction.CurrentTable + "]", connection);
                        DataTable dt = new DataTable();
                        sda.Fill(dt);
                        DGVMain.DataSource = dt;
                        connection.Close();
                    }
                }
            }

        }

        //QRY: Read Temp table data
        private void btnReadTemp_Click(object sender, EventArgs e)
        {
            using (OleDbConnection connection = new OleDbConnection(Con))
            {
                connection.Open();
                string QueryEntry = "SELECT * FROM [ACImport]";
                OleDbDataAdapter oda = new OleDbDataAdapter(
                    QueryEntry, Con);
                DataTable dt = new DataTable();
                oda.Fill(dt);
                DGVMain.DataSource = dt;
                connection.Close();
            }

        }

        private void cmbYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbYear.Text == "in list" && chkbxUseOA.Checked == false)
            {
                DialogResult dr = MessageBox.Show(
                "To enable 'in list' as search criteria for Year, you must use the OAData format. Would you like to change it to this setting?", "Year criteria",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (dr == DialogResult.OK)
                {
                    chkbxUseOA.Checked = true;
                }
                else
                {
                    MessageBox.Show("Criteria selection cancelled");
                    cmbYear.Text = "";
                }
            }
        }
        
        private void cmbColumnus_SelectedIndexChanged(object sender, EventArgs e)
        {
            JsonParams.ColumnString = cmbColumnus.Text;
        }

        //Populates combobox with the unique values of your table
        private void cmbComboValues_DropDown(object sender, EventArgs e)
        {
            if (JsonParams.ColumnString == null) { MessageBox.Show("Select a column name to search data from!"); return; }
            using (OleDbConnection connection = new OleDbConnection(Con))
            {
                cmbComboValues.DataSource = null;
                connection.Open();
                OleDbCommand cmd = new OleDbCommand("SELECT [" + JsonParams.ColumnString + "] FROM [" + SetupAction.CurrentTable + "]", connection);
                OleDbDataReader dr = cmd.ExecuteReader();
                IList<string> lister = new List<string>();
                while (dr.Read())
                {
                    lister.Add(dr[0].ToString());
                }
                lister = lister.Distinct().ToList();
                cmbComboValues.DataSource = lister;
                connection.Close();
            }
        }
        
        private void cmbComboValues_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetupAction.BonusSelectionBuilder = cmbComboValues.Text;
            lblStatus.Text = "Changed filter string!";
        }

        #region Clutter
        private void FilterAccount_Enter(object sender, EventArgs e)
        {
            SetupAction.FocusedControl = (TextBox)sender;
        }

        private void FilterYear_Enter(object sender, EventArgs e)
        {
            SetupAction.FocusedControl = (TextBox)sender;
        }

        private void FilterPeriod_Enter(object sender, EventArgs e)
        {
            SetupAction.FocusedControl = (TextBox)sender;
        }

        private void FilterBook_Enter(object sender, EventArgs e)
        {
            SetupAction.FocusedControl = (TextBox)sender;

        }

        private void FilterPHP_Enter(object sender, EventArgs e)
        {
            SetupAction.FocusedControl = (TextBox)sender;

        }

        private void FilterUSD_Enter(object sender, EventArgs e)
        {
            SetupAction.FocusedControl = (TextBox)sender;
        }

        private void FilterSite_Enter(object sender, EventArgs e)
        {
            SetupAction.FocusedControl = (TextBox)sender;

        }

        private void FilterStatus_Enter(object sender, EventArgs e)
        {
            SetupAction.FocusedControl = (TextBox)sender;

        }

        private void FilterDept_Enter(object sender, EventArgs e)
        {
            SetupAction.FocusedControl = (TextBox)sender;

        }

        private void DGVMain_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (txtCurrency.Text == "USD")
            {
                if ((this.DGVMain.Columns[e.ColumnIndex].Name == "Amount_USD") || (this.DGVMain.Columns[e.ColumnIndex].Name == "Amount_PHP"))
                {
                    if (e.Value != null)
                    {
                        DGVMain.Columns[e.ColumnIndex].DefaultCellStyle.Format = "0.00##";
                        if ((this.DGVMain.Columns[e.ColumnIndex].Name == "Amount_USD"))
                            e.CellStyle.BackColor = Color.LightSkyBlue;
                    }
                }
            }
            else if (txtCurrency.Text == "PHP")
            {
                if ((this.DGVMain.Columns[e.ColumnIndex].Name == "Amount_USD") || (this.DGVMain.Columns[e.ColumnIndex].Name == "Amount_PHP"))
                {
                    if (e.Value != null)
                    {
                        DGVMain.Columns[e.ColumnIndex].DefaultCellStyle.Format = "0.00##";
                        if ((this.DGVMain.Columns[e.ColumnIndex].Name == "Amount_PHP"))
                            e.CellStyle.BackColor = Color.LightSkyBlue;
                    }
                }
            }
            if ((this.DGVMain.Columns[e.ColumnIndex].Name == "Grand Total"))
            {
                if (e.Value != null)
                {
                    DGVMain.Columns[e.ColumnIndex].DefaultCellStyle.Format = "0.00##";
                    e.CellStyle.BackColor = Color.LightBlue;
                }
            }
            if ((this.DGVMain.Columns[e.ColumnIndex].Name == "Values"))
            {
                if (e.Value != null)
                {
                    DGVMain.Columns[e.ColumnIndex].DefaultCellStyle.Format = "0.00##";

                }
            }
        }

        private void chkbxISCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (chkbxISCheck.Checked == true)
            {
                chkbWithout.Visible = true;
            }
            else
            {
                chkbWithout.Checked = false;
                chkbWithout.Visible = false;
            }
        }

        private void FilterIS_Enter(object sender, EventArgs e)
        {
            SetupAction.FocusedControl = (TextBox)sender;
        }

        private void FilterITR_Enter(object sender, EventArgs e)
        {
            SetupAction.FocusedControl = (TextBox)sender;
        }
        #endregion

        private void rdBtnGrand_CheckedChanged(object sender, EventArgs e)
        {
            if (rdBtnGrand.Checked == true)
            {
                chkbxExportPivot.Visible = true;
            }
            else
            {
                chkbxExportPivot.Checked = false;
                chkbxExportPivot.Visible = false;
            }
        }

        public void EndRunTime()
        {
            Application.Exit();
        }
        
        private void cmbDictionList_SelectedIndexChanged(object sender, EventArgs e)
        {
            JsonParams.SelectedDictionary = cmbDictionList.Text;
        }

        /// <summary>
        /// Uploads the selected dictionary from JSON formatting
        /// </summary>
        /// <remarks>TODO: Refactor dictionary system completely to permit hot swapping of dictionary data (currently, an impossible feat due to referencing)</remarks>
        private void btnUpload_Click(object sender, EventArgs e)
        {
            if (JsonParams.SelectedDictionary == null)
            {
                MessageBox.Show("No dictionary was selected!");
                return;
            }

            if (txtDOCUnow.Text == "Balance Sheet")
            {
                var OverDiction = new Dictionary<string, Dictionary<int, DictionaryCheckup>>();
                var C = BalanceDictionary.C;
                var E = BalanceDictionary.E;
                var G = BalanceDictionary.G;
                var H = BalanceDictionary.H;
                var I = BalanceDictionary.I;
                var J = BalanceDictionary.J;
                var K = BalanceDictionary.K;
                var L = BalanceDictionary.L;
                var N = BalanceDictionary.N;
                var O = BalanceDictionary.O;
                var T = BalanceDictionary.T;
                OverDiction.Add("C", C);
                OverDiction.Add("E", E);
                OverDiction.Add("G", G);
                OverDiction.Add("H", H);
                OverDiction.Add("I", I);
                OverDiction.Add("J", J);
                OverDiction.Add("K", K);
                OverDiction.Add("L", L);
                OverDiction.Add("N", N);
                OverDiction.Add("O", O);
                OverDiction.Add("T", T);
                JsonParams.Found_Dictionary_BS = Find_Dictionaries(JsonParams.SelectedDictionary, OverDiction);
            }
            else
            {
                var OverDiction = new Dictionary<string, Dictionary<int, DictionarySetup>>();
                var C = IncomeDictionary.accountRevenue;
                var E = IncomeDictionary.accountOthers;
                var G = IncomeDictionary.accountExceptions;
                var H = IncomeDictionary.accountProjects;
                var I = IncomeDictionary.accountIncome;
                var J = IncomeDictionary.accountExpenses;
                var K = IncomeDictionary.accountDepartments;
                var L = IncomeDictionary.accountFacilities;
                var N = IncomeDictionary.accountTech;
                OverDiction.Add("Revenues", C);
                OverDiction.Add("Others", E);
                OverDiction.Add("Exxceptions", G);
                OverDiction.Add("Projects", H);
                OverDiction.Add("Income", I);
                OverDiction.Add("Expenses", J);
                OverDiction.Add("Departments", K);
                OverDiction.Add("Facilities", L);
                OverDiction.Add("Tech", N);
                JsonParams.Found_Dictionary_IS = Find_Dictionaries2(JsonParams.SelectedDictionary, OverDiction);

            }
            if ( JsonParams.Found_Dictionary_BS == null)
            {
                MessageBox.Show("Invalid dictionary selected!");
                return;
            }
            OpenFileDialog OFD = new OpenFileDialog();
            OFD.Filter = "JSON files|*.json|All files(*.*)|*.*";
            DialogResult DR = OFD.ShowDialog();
            if (OFD.FileName != null && DR == DialogResult.OK)
            {
                using (var stream = new StreamReader(File.OpenRead(OFD.FileName)))
                {
                    // Read our JSON from the file
                    var json = stream.ReadToEnd();
                    if (txtDOCUnow.Text == "Balance Sheet")
                    {
                        var json2 = JsonConvert.DeserializeObject<Dictionary<int, DictionaryCheckup>>(json);
                        DictionaryReplaceBalance(JsonParams.Found_Dictionary_BS, json2);

                    }
                    else
                    {
                        var json2 = JsonConvert.DeserializeObject<Dictionary<int, DictionarySetup>>(json);
                        DictionaryReplaceBalance2(JsonParams.Found_Dictionary_IS, json2);
                    }
                }
            }
            lblStatus.Text = "Dictionary pairs replaced!";
            timer.Enabled = true;
        }

        private void chkbxUseFilter_CheckedChanged(object sender, EventArgs e)
        {
            if (DGVMain.Rows.Count <= 0)
            {
                MessageBox.Show("Invalid row count detected in source table!");
                chkbxUseFilter.Checked = false;
                return;
            }
            if (chkbxUseFilter.Checked == true)
            {
                pbActive.Visible = true;
                pbInactive.Visible = false;
                lblActivity.Text = "Active";
                try
                {
                    UseFilter();
                }
                catch (Exception ex)
                {
                   MessageBox.Show(ex.Message);
                }
            }
            else
            {
                pbActive.Visible = false;
                pbInactive.Visible = true;
                lblActivity.Text = "Inactive";

            }
        }

        /// <summary>
        /// Background Worker for processing the filter after the Boolean checks were settled in the MatchesCriteria class
        /// </summary>
        /// <remarks>TODO: Refactor background worker to be less verbose and much faster</remarks>
        private void Worker_Filter_DoWork(object sender, DoWorkEventArgs e)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            StringBuilder FinalBuilder = SetFilterParameters.FilterBuilder;
            //Query Builder
            IncomeDictionary = new DictionaryInit();
            try
            {
                //Fire query and draw the view 
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    string query = FinalBuilder.ToString();
                    OleDbDataAdapter da = new OleDbDataAdapter(query, connection);
                    da.SelectCommand.Parameters.Clear();


                    #region Parameter Jungle
                    if (MatchesCriterias.HasAccount == true)
                        da.SelectCommand.Parameters.AddWithValue("@Account", OleDbType.Variant).Value = SetFilterParameters.Acc;
                    //Betweens
                    if (MatchesCriterias.HasYearBetween == true)
                    {
                        if (chkbxUseOA.Checked != true)
                        {
                            if (rdbtnYYYY.Checked == true)
                            {
                                try
                                {
                                    da.SelectCommand.Parameters.AddWithValue("@Starter", new DateTime(Convert.ToInt32(SetFilterParameters.Starter), 1, 1));
                                    da.SelectCommand.Parameters.AddWithValue("@Ender", new DateTime(Convert.ToInt32(SetFilterParameters.Ender), 1, 1));
                                }
                                catch (FormatException ex)
                                {
                                    MessageBox.Show("Conversion error, check if you selected an appropriate data type! " + ex);
                                }                            }
                            else if (rdbtnMDY.Checked == true)
                            {
                                try
                                {
                                    DateTime validDate = DateTime.Parse(SetFilterParameters.Starter);
                                    DateTime validDate2 = DateTime.Parse(SetFilterParameters.Ender);
                                    da.SelectCommand.Parameters.AddWithValue("@Starter", validDate);
                                    da.SelectCommand.Parameters.AddWithValue("@Ender", validDate2);
                                }
                                catch (FormatException ex)
                                {
                                    MessageBox.Show("Conversion error, check that you selected an appropriate data type! " + ex);
                                }
                            }
                            else if (rdbtnOA.Checked == true)
                            {
                                MessageBox.Show("OADate insert data type can only be used if you check Use: OADate format!");
                                return;
                            }
                        }
                        else
                        {
                            try
                            {
                                DateTime properdate = DateTime.FromOADate(Double.Parse(SetFilterParameters.Starter));
                                DateTime properdate2 = DateTime.FromOADate(Double.Parse(SetFilterParameters.Ender));
                                da.SelectCommand.Parameters.AddWithValue("@Starter", properdate);
                                da.SelectCommand.Parameters.AddWithValue("@Ender", properdate2);
                            }
                            catch (FormatException ex)
                            {
                                MessageBox.Show("Conversion error, check that you selected an appropriate data type! " + ex);
                            }
                        }
                    }
                    else if (MatchesCriterias.HasYear == true && MatchesCriterias.HasYearBetween != true)
                    {
                        if (chkbxUseOA.Checked == true)
                        {
                            try
                            {
                                if (cmbYear.Text == "in list")
                                {
                                    da.SelectCommand.Parameters.AddWithValue("@Year", DateTime.FromOADate(Double.Parse(SetFilterParameters.Year)));
                                }
                                else
                                {
                                    DateTime properdate = DateTime.FromOADate(Double.Parse(SetFilterParameters.Year));
                                    da.SelectCommand.Parameters.AddWithValue("@Year", properdate);
                                }
                            }
                            catch (FormatException ex)
                            {
                                MessageBox.Show("Conversion error, check that you selected an appropriate data type! " + ex);
                            }
                        }
                        else
                        {
                            try
                            {
                                //System.Globalization.DateTimeFormatInfo dateInfo = new System.Globalization.DateTimeFormatInfo();
                                //dateInfo.ShortDatePattern = "MM/dd/yyyy";
                                //DateTime validDate = new DateTime(Convert.ToInt32(FilterYear.Text), 1, 1);
                                DateTime validDate =  DateTime.Parse(FilterYear.Text);
                                da.SelectCommand.Parameters.AddWithValue("@Year", validDate);
                            }
                            catch (FormatException ex)
                            {
                                MessageBox.Show("Conversion error, check that you selected an appropriate data type! " + ex);
                            }
                        }
                    }
                    if (MatchesCriterias.HasPeriodBetween == true)
                    {
                        try
                        {
                            da.SelectCommand.Parameters.AddWithValue("@PeriodStarter", Convert.ToInt32(SetFilterParameters.PeriodStarter));
                            da.SelectCommand.Parameters.AddWithValue("@PeriodEnder", Convert.ToInt32(SetFilterParameters.PeriodEnder));
                        }
                        catch (FormatException ex)
                        {
                            MessageBox.Show("Conversion error, check that you selected an appropriate data type! " + ex);
                        }
                    }
                    else if (MatchesCriterias.HasPeriod == true && MatchesCriterias.HasPeriodBetween != true)
                        da.SelectCommand.Parameters.AddWithValue("@Period", OleDbType.Variant).Value = SetFilterParameters.Period;
                    if (MatchesCriterias.HasBook == true)
                        da.SelectCommand.Parameters.AddWithValue("@Book", OleDbType.Variant).Value = SetFilterParameters.Book;
                    //Currency
                    if ((MatchesCriterias.HasPHP == true) && (MatchesCriterias.HasPHPDecimal == true))
                    {
                        da.SelectCommand.Parameters.AddWithValue("@init", SetFilterParameters.init);
                        da.SelectCommand.Parameters.AddWithValue("@init2", SetFilterParameters.init2);
                    }
                    else if (MatchesCriterias.HasPHP == true)
                        da.SelectCommand.Parameters.AddWithValue("@PHP", Convert.ToDouble(SetFilterParameters.PHP));
                    if ((MatchesCriterias.HasUSD == true) && (MatchesCriterias.HasUSDDecimal == true))
                    {
                        da.SelectCommand.Parameters.AddWithValue("@init", SetFilterParameters.init);
                        da.SelectCommand.Parameters.AddWithValue("@init2", SetFilterParameters.init2);
                    }
                    else if (MatchesCriterias.HasUSD == true)
                        da.SelectCommand.Parameters.AddWithValue("@USD", Convert.ToDouble(SetFilterParameters.USD));
                    //Everything Else
                    if (MatchesCriterias.HasSite == true)
                        da.SelectCommand.Parameters.AddWithValue("@Site", OleDbType.Variant).Value = SetFilterParameters.Site;
                    if (MatchesCriterias.HasStatus == true)
                        da.SelectCommand.Parameters.AddWithValue("@Status", OleDbType.Variant).Value = SetFilterParameters.Status;
                    if (MatchesCriterias.HasDept == true)
                        da.SelectCommand.Parameters.AddWithValue("@Department", OleDbType.Variant).Value = SetFilterParameters.Dept;
                    if (MatchesCriterias.HasIS == true)
                        da.SelectCommand.Parameters.AddWithValue("@IS", OleDbType.Variant).Value = SetFilterParameters.IS;
                    if (MatchesCriterias.HasITR == true)
                        da.SelectCommand.Parameters.AddWithValue("@ITR", OleDbType.Variant).Value = SetFilterParameters.ITR;
                    #endregion
                    DataTable dt = new DataTable();
                    if (Worker_Filter.CancellationPending)
                    {
                        e.Cancel = true;
                        Worker_Filter.ReportProgress(0);
                        return;
                    }
                    da.Fill(dt);
                    FilterStorage.FilterTable = dt;
                }
            }
            catch (OleDbException ex)
            {
                MessageBox.Show("Select Failed: " + ex);
            }
            finally
            {
                sw.Stop();
                FilterStorage.Occurrence = sw.Elapsed.ToString();
                SetFilterParameters.FilterBuilder = null;
                MatchesCriterias.HasPeriodBetween = false; MatchesCriterias.HasITR = false; MatchesCriterias.HasIS = false; MatchesCriterias.HasDept = false; MatchesCriterias.HasStatus = false; MatchesCriterias.HasSite = false; MatchesCriterias.HasUSD = false; MatchesCriterias.HasPHP = false; MatchesCriterias.HasBook = false; MatchesCriterias.HasPeriod = false; MatchesCriterias.HasYear = false; MatchesCriterias.HasYearBetween = false; MatchesCriterias.HasAccount = false;
            }
        }
        private void Worker_Filter_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Yellow;
                timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false;
                panelColors.BackColor = Color.Yellow;
                lblStatus.Text = "Process was cancelled";
            }
            else if (e.Error != null)
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Red;
                timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false;
                panelColors.BackColor = Color.Red;
                lblStatus.Text = "Error: The thread aborted";
            }
            else
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Lime;
                timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false;
                panelColors.BackColor = Color.Green;
                lblStatus.Text = "Filter successful!";
            }

            if (FilterStorage.FilterTable != null)
            {
                DGVMain.DataSource = FilterStorage.FilterTable;
            }
            else
            {
                MessageBox.Show("No table was created...");
                return;
            }
            EnablePanels("", true, false);
            pBarMain.Value = 0;
            pBarMain.Style = MetroColorStyle.Default;
        }

        public async void CompleteProcess()
        {
            pBarMain.Value = 100;
            pBarMain.Style = MetroColorStyle.Lime;
            Thread.Sleep(1000);
            EnablePanels("Process complete!", true, true);
            await TaskDelay_5();
            pBarMain.Value = 0;
            pBarMain.Style = MetroColorStyle.Default;
        }
        
        private void chkbxBalanceFilter_CheckedChanged(object sender, EventArgs e)
        {
            if (DGVMain.Rows.Count <= 0)
            {
                MessageBox.Show("Invalid row count detected in source table!");
                chkbxBalanceFilter.Checked = false;
                return;
            }
            if (chkbxBalanceFilter.Checked == true)
            {
                pbActive.Visible = true;
                pbInactive.Visible = false;
                lblActivity.Text = "Active";
                UseFilter();
            }
            else
            {
                pbInactive.Visible = true;
                lblActivity.Text = "Inactive";
                pbActive.Visible = false;
            }
        }

        private void MDT_ValueChanged(object sender, EventArgs e)
        {
            if (chkbxUseOA.Checked == true)
            {
                Double trouble = Convert.ToDouble(MDT.Value.ToShortDateString());
                FilterYear.Text = Convert.ToString(trouble);
            }
            FilterYear.Text = MDT.Value.ToShortDateString();
            lblStatus.Text = "Changed date filter string!";
        }
        
        private void btnBalances_Click(object sender, EventArgs e)
        {
            if (DGVMain.Rows.Count <= 0)
            {
                MessageBox.Show("Invalid row count in the table view!");
                return;
            }
            //Ascertain whether or not a filter is in place
            DataTable CoreTable = new DataTable();
            if ((chkbxBalanceFilter.Checked == true))
            {
                using (OleDbConnection con = new OleDbConnection(Con))
                {
                    con.Open();
                    using (OleDbDataAdapter a = new OleDbDataAdapter(
                        "SELECT * FROM ACImport", con))
                    {
                        a.Fill(CoreTable);
                    }
                }
            }
            else
            {
                using (OleDbConnection con = new OleDbConnection(Con))
                {
                    con.Open();
                    using (OleDbDataAdapter a = new OleDbDataAdapter(
                        "SELECT * FROM " + SetupAction.CurrentTable, con))
                    {
                        a.Fill(CoreTable);
                    }
                }
            }
            int rowCount = CoreTable.Rows.Count;
            //New rows
            DataTable dt = new DataTable();
            dt.Columns.Add("Balance", typeof(string));
            dt.Columns.Add("Values", typeof(Double));
            #region Initialize table to match balance sheet report appearance
            DataRow Cash_Row = dt.NewRow();
            DataRow Receiveable_Row = dt.NewRow();
            DataRow Party_Row = dt.NewRow();
            DataRow Prepayment_Row = dt.NewRow();
            DataRow Property_Row = dt.NewRow();
            DataRow Deposit_Row = dt.NewRow();
            DataRow Investment_Row = dt.NewRow();
            DataRow Deferred_Row = dt.NewRow();
            DataRow Goodwill_Row = dt.NewRow();
            DataRow Intangible_Row = dt.NewRow();
            DataRow Total_Asset_Row = dt.NewRow();
            DataRow Total_Liabilities_Row = dt.NewRow();
            DataRow C_Stock_Row = dt.NewRow();
            DataRow T_Stock_Row = dt.NewRow();
            DataRow P_Stock_Row = dt.NewRow();
            DataRow CP_Stock_Row = dt.NewRow();
            DataRow APIC_Row = dt.NewRow();
            DataRow Retained_Row = dt.NewRow();
            DataRow Retained_L_Row = dt.NewRow();
            DataRow OCI_Row = dt.NewRow();
            DataRow Foreign_Row = dt.NewRow();
            DataRow Total_Equity_Row = dt.NewRow();
            DataRow Total_EandL_Row = dt.NewRow();
            //[A]
            List<Double> Cash_Value = new List<double>();
            List<Double> Receiveable_Value = new List<double>();
            List<Double> Party_Value = new List<double>();
            List<Double> Prepayment_Value = new List<double>();
            List<Double> Property_Value = new List<double>();
            List<Double> Deposit_Value = new List<double>();
            List<Double> Investment_Value = new List<double>();
            List<Double> Deferred_Value = new List<double>();
            List<Double> Goodwill_Value = new List<double>();
            List<Double> Intangible_Value = new List<double>();
            List<Double> C_Stock_Value = new List<double>();
            List<Double> T_Stock_Value = new List<double>();
            List<Double> P_Stock_Value = new List<double>();
            List<Double> CP_Stock_Value = new List<double>();
            List<Double> APIC_Value = new List<double>();
            List<Double> Retained_Value = new List<double>();
            List<Double> Retained_L_Value = new List<double>();
            List<Double> OCI_Value = new List<double>();
            List<Double> Foreign_Value = new List<double>();
            List<Double> ITR_Value = new List<double>();
            List<Double> Liabilities_Value = new List<double>();
            Double Total_Equity_Value = 0;
            Double Total_Asset_Value = 0;
            Double Total_Liabilities_Value = 0;
            Double Total_EandL_Value = 0;
            List<DataRow> ACollection = new List<DataRow>();
            List<DataRow> LCollection = new List<DataRow>();
            List<DataRow> ECollection = new List<DataRow>();
            ACollection.Add(Cash_Row);
            ACollection.Add(Receiveable_Row);
            ACollection.Add(Party_Row);
            ACollection.Add(Prepayment_Row);
            ACollection.Add(Property_Row);
            ACollection.Add(Deposit_Row);
            ACollection.Add(Investment_Row);
            ACollection.Add(Deferred_Row);
            ACollection.Add(Goodwill_Row);
            ACollection.Add(Intangible_Row);
            ACollection.Add(Total_Asset_Row);
            LCollection.Add(Total_Liabilities_Row);
            ECollection.Add(C_Stock_Row);
            ECollection.Add(T_Stock_Row);
            ECollection.Add(P_Stock_Row);
            ECollection.Add(CP_Stock_Row);
            ECollection.Add(APIC_Row);
            ECollection.Add(Retained_Row);
            ECollection.Add(Retained_L_Row);
            ECollection.Add(OCI_Row);
            ECollection.Add(Foreign_Row);
            ECollection.Add(Total_Equity_Row);
            ECollection.Add(Total_EandL_Row);
            #endregion            //Row headers
            Cash_Row[0] = "Cash and cash equivalents";
            Receiveable_Row[0] = "Receivables";
            Party_Row[0] = "Net advances to/from related parties";
            Prepayment_Row[0] = "Prepayments and other current assets";
            Property_Row[0] = "Property and equipment";
            Deposit_Row[0] = "Deposits";
            Investment_Row[0] = "Investment in subsidiaries";
            Deferred_Row[0] = "Deferred tax asset";
            Goodwill_Row[0] = "Goodwill";
            Intangible_Row[0] = "Intangible Assets";
            Total_Asset_Row[0] = "[Total Assets] ";
            Total_Liabilities_Row[0] = "[Total Liabilities] ";
            C_Stock_Row[0] = "Common Stock";
            T_Stock_Row[0] = "Treasury Stock";
            P_Stock_Row[0] = "Preferred Stock";
            CP_Stock_Row[0] = "Convertible Preferred Stock";
            APIC_Row[0] = "APIC";
            Retained_Row[0] = "Retained Earnings";
            Retained_L_Row[0] = "Retained Earnings L";
            OCI_Row[0] = "OCI-Minimum Pension Liablility";
            Foreign_Row[0] = "Forgn Curr Trnsltn";
            Total_Equity_Row[0] = "[Total Equity] ";
            Total_EandL_Row[0] = "[Total Liabilities and Equity] ";

            List<Dictionary<int, DictionaryCheckup>> Dictionaries = new List<Dictionary<int, DictionaryCheckup>> { BalanceDictionary.C, BalanceDictionary.E, BalanceDictionary.G, BalanceDictionary.H, BalanceDictionary.I, BalanceDictionary.J, BalanceDictionary.K, BalanceDictionary.L, BalanceDictionary.N, BalanceDictionary.O, BalanceDictionary.T }; // Replace ... with D,E,F, etc. until T 
            // Iterate each dictionary and if found, exit the loop.
            for (int rowindex = 0; rowindex < rowCount; rowindex++)
            {
                int Account_Key = Convert.ToInt32(CoreTable.Rows[rowindex][3].ToString());
                int AbsoluteKey = Account_Key;
                while (AbsoluteKey >= 10) { AbsoluteKey /= 10; }
                if (AbsoluteKey == 4)
                {
                    Retained_Value.Add(ReturnAmountValue(rowindex, CoreTable));
                }
                else
                {
                    foreach (var dict in Dictionaries)
                    {
                        string FoundValue = Dictionary_BalanceCheckup(Account_Key, dict);
                        if (FoundValue != null)
                        {
                            switch (FoundValue)
                            {
                                case "C":
                                    Cash_Value.Add(ReturnAmountValue(rowindex, CoreTable));
                                    break;
                                case "E":
                                    Receiveable_Value.Add(ReturnAmountValue(rowindex, CoreTable));
                                    break;
                                case "G":
                                    Prepayment_Value.Add(ReturnAmountValue(rowindex, CoreTable));
                                    break;
                                case "H":
                                    Investment_Value.Add(ReturnAmountValue(rowindex, CoreTable));
                                    break;
                                case "I":
                                    Party_Value.Add(ReturnAmountValue(rowindex, CoreTable));
                                    break;
                                case "J":
                                    Deposit_Value.Add(ReturnAmountValue(rowindex, CoreTable));
                                    break;
                                case "K":
                                    Property_Value.Add(ReturnAmountValue(rowindex, CoreTable));
                                    break;
                                case "L":
                                    if (Account_Key == 170001)
                                        Goodwill_Value.Add(ReturnAmountValue(rowindex, CoreTable));
                                    else
                                        Intangible_Value.Add(ReturnAmountValue(rowindex, CoreTable));
                                    break;
                                case "N":
                                    Liabilities_Value.Add(ReturnAmountValue(rowindex, CoreTable));
                                    break;
                                case "O":
                                    Deferred_Value.Add(ReturnAmountValue(rowindex, CoreTable));
                                    break;
                                case "T":
                                    switch (Account_Key)
                                    {
                                        case 300000: C_Stock_Value.Add(ReturnAmountValue(rowindex, CoreTable)); break;
                                        case 300001: T_Stock_Value.Add(ReturnAmountValue(rowindex, CoreTable)); break;
                                        case 300002: P_Stock_Value.Add(ReturnAmountValue(rowindex, CoreTable)); break;
                                        case 300103: CP_Stock_Value.Add(ReturnAmountValue(rowindex, CoreTable)); break;
                                        case 310000: APIC_Value.Add(ReturnAmountValue(rowindex, CoreTable)); break;
                                        case 320000: Retained_Value.Add(ReturnAmountValue(rowindex, CoreTable)); break;
                                        case 320001: Retained_Value.Add(ReturnAmountValue(rowindex, CoreTable)); break;
                                        case 320002: Retained_L_Value.Add(ReturnAmountValue(rowindex, CoreTable)); break;
                                        case 330000: Retained_Value.Add(ReturnAmountValue(rowindex, CoreTable)); break;
                                        case 350000:
                                        case 350001: OCI_Value.Add(ReturnAmountValue(rowindex, CoreTable)); break;
                                        case 360000:
                                        case 360001:
                                        case 360100: Foreign_Value.Add(ReturnAmountValue(rowindex, CoreTable)); break;
                                    }
                                    break;
                            }
                            break;
                        }
                    } 
                }
            }

            // Sum up totals regardless of availability
            Total_Asset_Value = (Cash_Value.Sum() + Receiveable_Value.Sum() + Party_Value.Sum() + Prepayment_Value.Sum() + Property_Value.Sum() + Deposit_Value.Sum() + Investment_Value.Sum() + Deferred_Value.Sum() + Goodwill_Value.Sum() + Intangible_Value.Sum());
            Total_Liabilities_Value = (Liabilities_Value.Sum());
            Total_Equity_Value = (C_Stock_Value.Sum() + T_Stock_Value.Sum() + P_Stock_Value.Sum() + CP_Stock_Value.Sum() + APIC_Value.Sum() + Retained_Value.Sum() + Retained_L_Value.Sum() + OCI_Value.Sum() + Foreign_Value.Sum());
            Total_EandL_Value = Total_Liabilities_Value + Total_Equity_Value;

            // Sum up values depending on availability
            if (Cash_Value.Count > 0) Cash_Row[1] = Cash_Value.Sum();
            if (Receiveable_Value.Count > 0) Receiveable_Row[1] = Receiveable_Value.Sum();
            if (Party_Value.Count > 0) Party_Row[1] = Party_Value.Sum();
            if (Prepayment_Value.Count > 0) Prepayment_Row[1] = Prepayment_Value.Sum();
            if (Property_Value.Count > 0) Property_Row[1] = Property_Value.Sum();
            if (Deposit_Value.Count > 0) Deposit_Row[1] = Deposit_Value.Sum();
            if (Investment_Value.Count > 0) Investment_Row[1] = Investment_Value.Sum();
            if (Deferred_Value.Count > 0) Deferred_Row[1] = Deferred_Value.Sum();
            if (Goodwill_Value.Count > 0) Goodwill_Row[1] = Goodwill_Value.Sum();
            if (Intangible_Value.Count > 0) Intangible_Row[1] = Intangible_Value.Sum();
            Total_Asset_Row[1] = Total_Asset_Value;
            if (Liabilities_Value.Count > 0) Total_Liabilities_Row[1] = Total_Liabilities_Value;
            if (C_Stock_Value.Count > 0) C_Stock_Row[1] = C_Stock_Value.Sum();
            if (T_Stock_Value.Count > 0) T_Stock_Row[1] = T_Stock_Value.Sum();
            if (P_Stock_Value.Count > 0) P_Stock_Row[1] = P_Stock_Value.Sum();
            if (CP_Stock_Value.Count > 0) CP_Stock_Row[1] = CP_Stock_Value.Sum();
            if (APIC_Value.Count > 0) APIC_Row[1] = APIC_Value.Sum();
            if (Retained_Value.Count > 0) Retained_Row[1] = Retained_Value.Sum();
            if (Retained_L_Value.Count > 0) Retained_L_Row[1] = Retained_L_Value.Sum();
            if (OCI_Value.Count > 0) OCI_Row[1] = OCI_Value.Sum();
            if (Foreign_Value.Count > 0) Foreign_Row[1] = Foreign_Value.Sum();
            Total_Equity_Row[1] = Total_Equity_Value;
            Total_EandL_Row[1] = Total_EandL_Value;

            //Append row collections and finalize report
            for (int row = 0; row < ACollection.Count(); row++)
            {
                dt.Rows.Add(ACollection.ElementAt(row));
            }
            for (int row = 0; row < LCollection.Count(); row++)
            {
                dt.Rows.Add(LCollection.ElementAt(row));
            }
            for (int row = 0; row < ECollection.Count(); row++)
            {
                dt.Rows.Add(ECollection.ElementAt(row));
            }
            DGVMain.DataSource = dt;
            CompleteProcess();
    }

        /// <summary>
        /// Determine the class grouping of a given key
        /// </summary>
        /// <parameter="mapKey">Corresponds to the letter grouping in the backend dictionary</parameter>
        public string Dictionary_BalanceCheckup(int mapKey, Dictionary<int, DictionaryCheckup> accountLexicon)
        {
            if (accountLexicon.TryGetValue(mapKey, out DictionaryCheckup ClassValues))
            {
                String Grouping = "";
                Grouping = (ClassValues.theGrouping.ToString());
                return Grouping;
            }
            else { return null;  }
        }

        /// <summary>
        /// Depending on whether the setting is at USD or PHP, get the value as Double
        /// </summary>
        /// <parameter="mapKey">Corresponds to the letter grouping in the backend dictionary</parameter>
        public Double ReturnAmountValue(int rowIndex, DataTable dataTable)
        {
           Double GetValue = 0;
           if (txtCurrency.Text == "USD")
            {
                return GetValue = (dataTable.Rows[rowIndex][9] == DBNull.Value) ? 0 : Convert.ToDouble(dataTable.Rows[rowIndex][9].ToString());
            } 
           else
            {
                return GetValue = (dataTable.Rows[rowIndex][12] == DBNull.Value) ? 0 : Convert.ToDouble(dataTable.Rows[rowIndex][12].ToString());
            }
        }

        //QRY: Loads up an excel file into memory
        private void Worker_Import_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                Conn =
                    new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + SetupAction.IO_Name +
                                        ";Extended Properties=  'Excel 12.0 Xml;HDR=Yes; IMEX = 1;TypeGuessRows=0;ImportMixedTypes=Text';");
                Conn.Open();
                DataTable dt = Conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                Conn.Close();
                MaxRowCount.MaxSchema = dt.Rows.Count;
                for (int i = 0; i < MaxRowCount.MaxSchema; i++)
                {
                    String sheetName = dt.Rows[i]["TABLE_NAME"].ToString();
                    sheetName = sheetName.Substring(0, sheetName.Length - 1);
                    cmbSheets.Invoke((MethodInvoker)delegate
                    {
                        cmbSheets.Items.Add(sheetName);
                    });
                    Worker_Import.ReportProgress(i);
                    if (Worker_Import.CancellationPending)
                    {
                        e.Cancel = true;
                        Worker_Import.ReportProgress(0);
                        return;
                    }
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
                Worker_Import.ReportProgress(MaxRowCount.MaxSchema);
            }
        }
        private void Worker_Import_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pBarMain.Value = (e.ProgressPercentage * 100) / MaxRowCount.MaxSchema;
        }
        private async void Worker_Import_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Yellow;
                timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false;
                panelColors.BackColor = Color.Yellow;
                lblStatus.Text = "Process was cancelled";
            }
            else if (e.Error != null)
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Red;
                timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false;
                panelColors.BackColor = Color.Red;
                lblStatus.Text = "Error: The thread aborted";
            }
            else
            {
                panel9.BackColor = Color.Teal;
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Lime;
                timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false;
                panelColors.BackColor = Color.Green;
                lblStatus.Text = "Schema successfully uploaded!";
            }
            EnablePanels("", true, false);
            await TaskDelay_5();
            pBarMain.Value = 0;
            pBarMain.Style = MetroColorStyle.Default;
        }

        //QRY: Tramsfers the data from memory into the Access database
        private void Worker_Transfer_DoWork(object sender, DoWorkEventArgs e)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            try
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    using (OleDbCommand cmd = new OleDbCommand())
                    {
                        cmd.Connection = connection;
                        connection.Open();
                        using (OleDbTransaction Scope = connection.BeginTransaction(SD.IsolationLevel.ReadCommitted))
                        {
                            try
                            {
                                cmd.Transaction = Scope;
                                cmd.CommandText =
                                "Insert INTO [" + SetupAction.CurrentTable + "] (Business_Unit,Ledger,Account,Description,Span_Year,Period,Book_Code,Alt_Currency,Amount_USD,Base_Currency,PDS,Amount_PHP,Site,Stat_Site,Status_Name,Status_Regime,Department,Description_Department,Project_Number,Description_Project,Affiliate,IS_Accounts,Description_ITR) " +
                                                              "VALUES(@Business_Unit,@Ledger,@Account,@Description,@Span_Year,@Period,@Book_Code,@Alt_Currency,@Amount_USD,@Base_Currency,@PDS,@Amount_PHP,@Site,@Stat_Site,@Status_Name,@Status_Regime,@Department,@Description_Department,@Project_Number,@Description_Project,@Affiliate,@IS_Accounts,@Description_ITR)";
                                for (int rowindex = 0; rowindex < MaxRowCount.MaxExcel; rowindex++)
                                {
                                    cmd.Parameters.Clear();
                                    cmd.Parameters.AddWithValue("@Business_Unit", VirtualTable.Rows[rowindex][1]);
                                    cmd.Parameters.AddWithValue("@Ledger", VirtualTable.Rows[rowindex][2]);
                                    cmd.Parameters.AddWithValue("@Account", VirtualTable.Rows[rowindex][3]);
                                    cmd.Parameters.AddWithValue("@Description", VirtualTable.Rows[rowindex][4]);
                                    if (rdbtnYYYY.Checked == true)
                                    {
                                        try
                                        {
                                            cmd.Parameters.AddWithValue("@Span_Year", (VirtualTable.Rows[rowindex][5] == null) ? VirtualTable.Rows[rowindex][9] = "1/1/1000" : new DateTime(Convert.ToInt32(VirtualTable.Rows[rowindex][5]), 1, 1));
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("Invalid date! " + ex);
                                        }
                                    }
                                    else if (rdbtnMDY.Checked == true)
                                    {
                                        string date = VirtualTable.Rows[rowindex][5].ToString();
                                        if (DateTime.TryParse(date, out DateTime result) == true)
                                        {
                                            try
                                            {
                                                cmd.Parameters.AddWithValue("@Span_Year", (VirtualTable.Rows[rowindex][5] == null) ? VirtualTable.Rows[rowindex][9] = "1/1/1000" : result.ToString("MM/dd/yyyy"));
                                                
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
                                    cmd.Parameters.AddWithValue("@Amount_USD", (VirtualTable.Rows[rowindex][9] == DBNull.Value) ? VirtualTable.Rows[rowindex][9] = 0 : Convert.ToDouble(VirtualTable.Rows[rowindex][9]));
                                    cmd.Parameters.AddWithValue("@Base_Currency", VirtualTable.Rows[rowindex][10]);
                                    cmd.Parameters.AddWithValue("@PDS", VirtualTable.Rows[rowindex][11]);
                                    cmd.Parameters.AddWithValue("@Amount_PHP", (VirtualTable.Rows[rowindex][12] == DBNull.Value) ? VirtualTable.Rows[rowindex][12] = 0 : Convert.ToDouble(VirtualTable.Rows[rowindex][12]));
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
                                    if (Worker_Transfer.CancellationPending)
                                    {
                                        e.Cancel = true;
                                        Worker_Transfer.ReportProgress(0);
                                        return;
                                    }
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
            sw.Stop();
            FilterStorage.Occurrence = sw.Elapsed.ToString();
            Worker_Transfer.ReportProgress(MaxRowCount.MaxExcel);
        }
        private async void Worker_Transfer_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Yellow;
                timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false;
                panelColors.BackColor = Color.Yellow;
                lblStatus.Text = "Process was cancelled";
            }
            else if (e.Error != null)
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Red;
                timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false;
                panelColors.BackColor = Color.Red;
                lblStatus.Text = "Error: The thread aborted";
            }
            else
            {
                pBarMain.Value = 100;
                pBarMain.Style = MetroColorStyle.Lime;
                timer.Enabled = true ? timer.Enabled = true : timer.Enabled = false;
                panelColors.BackColor = Color.Green;
                lblStatus.Text = "Import successful!";
            }
            lblElapsedTime.Text = FilterStorage.Occurrence;
            DialogResult dr = MessageBox.Show(
            "Would you like to reload the view with the database?", "Reload?",
            MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (dr == DialogResult.OK)
            {
                lblStatus.Text = "Loading... this may take a few minutes!";
                Restore();
            }
            EnablePanels("", true, false);
            await TaskDelay_5();
            pBarMain.Value = 0;
            pBarMain.Style = MetroColorStyle.Default;
        }
        private void Worker_Transfer_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pBarMain.Value = (e.ProgressPercentage * 100) / MaxRowCount.MaxExcel;
        }

        private void CopyAlltoClipboard()
        {
            //to remove the first blank column from datagridview
            DGVMain.RowHeadersVisible = false;
            DGVMain.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            DGVMain.SelectAll();
            DataObject dataObj = DGVMain.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        //Exports the visible data from the datagridview into a selected excel file, preferably, a blank one.
        private void btnToExcel_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            openFileDialog.ShowDialog();
            string pather = openFileDialog.FileName;
            FileInfo fInfo = new FileInfo(pather);
            if (pather == "" || pather == " " || pather == null)
            {
                MessageBox.Show("Path was invalid!");
                return;
            }
            if (!string.IsNullOrEmpty(pather))
            {
                if (chkbxExportPivot.Checked != true)
                {
                    EnablePanels("Loading... this may take a moment!", false, true);
                    try
                    {
                        if (DGVMain.Rows.Count <= 0)
                        {
                            MessageBox.Show("Invalid row count!");
                            return;
                        }
                        GetDataTableFromDGV(DGVMain);
                        using (ExcelPackage pck = new ExcelPackage())
                        {
                            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("From Program");
                            ws.Cells["A:X"].Style.Numberformat.Format = null;
                            ws.Cells["A1"].LoadFromDataTable(VirtualTable, true);
                            ws.Cells["A1:X1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            ws.Cells["A1:X1"].Style.Fill.BackgroundColor.SetColor(Color.Silver);
                            ws.Cells["A1"].Value = "ID";
                            ws.Cells["B1"].Value = "Business_Unit";
                            ws.Cells["C1"].Value = "Ledger";
                            ws.Cells["D1"].Value = "Account";
                            ws.Cells["E1"].Value = "Description";
                            ws.Cells["F1"].Value = "Span_Year";
                            ws.Cells["G1"].Value = "Period";
                            ws.Cells["H1"].Value = "Book_Code";
                            ws.Cells["I1"].Value = "Currency";
                            ws.Cells["J1"].Value = "Amount_USD";
                            ws.Cells["J:J"].Style.Numberformat.Format = null;
                            ws.Cells["M:M"].Style.Numberformat.Format = "0.00";
                            ws.Cells["K1"].Value = "Base_Currency";
                            ws.Cells["L1"].Value = "PDS";
                            ws.Cells["L:L"].Style.Numberformat.Format = null;
                            ws.Cells["M:M"].Style.Numberformat.Format = "0.00";
                            ws.Cells["M1"].Value = "Amount_PHP";
                            ws.Cells["M:M"].Style.Numberformat.Format = null;
                            ws.Cells["M:M"].Style.Numberformat.Format = "0.00";
                            ws.Cells["N1"].Value = "Site";
                            ws.Cells["O1"].Value = "Stat_Site";
                            ws.Cells["P1"].Value = "Status_Name";
                            ws.Cells["Q1"].Value = "Status_Regime";
                            ws.Cells["R1"].Value = "Department";
                            ws.Cells["S1"].Value = "Description_Department";
                            ws.Cells["T1"].Value = "Project_Number";
                            ws.Cells["U1"].Value = "Description_Project";
                            ws.Cells["V1"].Value = "Affiliate";
                            ws.Cells["W1"].Value = "IS_Accounts";
                            ws.Cells["X1"].Value = "Description_ITR";
                            ws.Row(1).Style.Font.Bold = true;
                            pck.SaveAs(fInfo);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Export error: " + ex.ToString());
                    }
                    finally
                    {
                        pbActive.Visible = false;
                        VirtualTable.Clear();
                        CompleteProcess();
                    } 
                }
                else
                {
                    Excel.Application xlexcel;
                    Excel.Workbook xlWorkBook;
                    Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;
                    xlexcel = new Excel.Application();
                    try
                    {
                        CopyAlltoClipboard();
                        xlexcel.Visible = true;
                        xlWorkBook = xlexcel.Workbooks.Add(pather);
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        Excel.Range cellsRange = null;
                        Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[Convert.ToInt32(txtOrigin.Text), 1];
                        CR.Select();
                        xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                        cellsRange = xlWorkSheet.UsedRange;
                        cellsRange.EntireColumn.AutoFit();
                        cellsRange.EntireRow.AutoFit();
                        var renge = cellsRange.Rows[1];
                        renge.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                        renge = null;
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        ReleaseObject(cellsRange);
                        ReleaseObject(xlWorkSheet);
                        ReleaseObject(xlWorkBook);
                        ReleaseObject(xlexcel);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                
            }
        }
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }

        }

        private void QueryRefiner_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        //QRY: Delete row by ID
        private void btnDeleteRow_Click(object sender, EventArgs e)
        {
            if (txtSelectID.Text != "" || txtSelectID.Text != " ")
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    try
                    {
                        connection.Open();
                        using (OleDbTransaction Scope = connection.BeginTransaction(SD.IsolationLevel.ReadCommitted))
                        {
                            try
                            {
                                OleDbCommand cmd = new OleDbCommand("DELETE FROM [" + SetupAction.CurrentTable + "] WHERE ID = @ID", connection);
                                cmd.Parameters.AddWithValue("@ID", txtSelectID.Text);
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
                        MessageBox.Show("SQL error" + es);
                    }
                    finally
                    {
                        OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM [" + SetupAction.CurrentTable + "]", connection);
                        DataTable dt = new DataTable();
                        sda.Fill(dt);
                        DGVMain.DataSource = dt;
                        connection.Close();
                    }
                }
            }
        }
        
        private void btnDrag_Click(object sender, EventArgs e)
        {
            if (DGVMain.SelectedCells.Count > 0)
            {
                int selectedrowindex = DGVMain.SelectedCells[0].RowIndex;
                DataGridViewRow selectedRow = DGVMain.Rows[selectedrowindex];
                string ID = Convert.ToString(selectedRow.Cells[0].Value);
                txtSelectID.Text = ID;
            }
        }

        //VFX: Filter panel gets colored on click
        private void ColorThePanel_Click(object sender, EventArgs e)
        {
            Panel clickedPanel = sender as Panel;
            if (clickedPanel != null)
            {
                if (clickedPanel.BackColor == Color.GhostWhite)
                {
                    clickedPanel.BackColor = Color.GreenYellow;
                }
                else
                {
                    clickedPanel.BackColor = Color.GhostWhite;
                }
            }
        }

        //VFX: Filter panel gets colored on double click
        private void DoubleColorPanel_DoubleClick(object sender, EventArgs e)
        {
            Panel clickedPanel = sender as Panel;
            if (clickedPanel != null)
            {
                if (clickedPanel.BackColor == Color.GhostWhite)
                {
                    clickedPanel.BackColor = Color.Plum;
                }
                else
                {
                    clickedPanel.BackColor = Color.GhostWhite;
                }
            }
        }

        //TODO: Make use of worksheet to datatable method
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

        private void LoadWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            DGVExcel.Invoke((Action)(() => DGVExcel.DataSource = ImportTable));
        }
    }

    public static class GuiExtensionMethods
    {
        public static void Enable(this Control con, bool enable)
        {
            if (con != null)
            {
                foreach (Control c in con.Controls)
                {
                    //c.Enable(enable);
                }

                try
                {
                    //con.Invoke((MethodInvoker)(() => con.Enabled = enable));
                }
                catch
                {
                }
            }
        }
    }

    public static class DataTableExtensions
    {
        public static DataView ApplySort(this DataTable table, Comparison<DataRow> comparison)
        {

            DataTable clone = table.Clone();
            List<DataRow> rows = new List<DataRow>();
            foreach (DataRow row in table.Rows)
            {
                rows.Add(row);
            }

            rows.Sort(comparison);

            foreach (DataRow row in rows)
            {
                clone.Rows.Add(row.ItemArray);
            }

            return clone.DefaultView;
        }
    }
}
