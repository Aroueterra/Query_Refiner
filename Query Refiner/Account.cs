using MetroFramework.Controls;
using MetroFramework.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace Query_Refiner
{
    public partial class Account : MetroForm
    {
        public string Con = ConfigurationManager.ConnectionStrings["Con"].ConnectionString;
        public static string passedBU;
        public static string passedDOCU;
        public static string passedCRNCY;
        public static bool SecondTime;
        Boolean SelectTable = false;
        public QueryRefiner FormToShowOnClose { get; set; }
        byte rdBoolean = 0;

        public Account()
        {
            InitializeComponent();
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
        }

        private void Account_Load(object sender, EventArgs e)
        {
            this.Opacity = 0;
            FadeIn(this, 100);
            lblHistory.Text = Con;
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
        }

        private void FadeTimer_Tick(object sender, EventArgs e)
        {
            this.Opacity += 0.5;

            if (this.Opacity == 1.00)
            {
                FadeTimer.Stop();
                Console.WriteLine("Timer Stopped!");
            }

        }

        private async void FadeIn(Form o, int interval = 80)
        {
            //Object is not fully invisible. Fade it in
            while (o.Opacity < 1.0)
            {
                await Task.Delay(interval);
                o.Opacity += 0.05;
            }
            o.Opacity = 1; //make fully visible            
        }

        private void TileCPSC_Click(object sender, EventArgs e)
        {
            StringBuilder messageBuilder = new StringBuilder(200);
            messageBuilder.Append("[30023] Convergys Philippines Services Corporation:");
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append("  \u2022 GEO: Philippines");
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append("  \u2022 Convergys One Building, 6796 Ayala Ave. cor. Salcedo St., Legaspi Village, Makati City 1229");
            textPanePEZA.Text = messageBuilder.ToString();
            lblBUitem.Text = "30023";
            rdBoolean = 1;
        }

        private void TileERMI_Click(object sender, EventArgs e)
        {
            StringBuilder messageBuilder = new StringBuilder(200);
            messageBuilder.Append("[30051] Encore Receivable Management Inc:");
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append("  \u2022 GEO: Philippines");
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append("  \u2022 5F Glorietta 5, Ayala Ave. cor. Office Drive, Brgy. San Lorenzo, Makati City 1223");
            textPanePEZA.Text = messageBuilder.ToString();
            lblBUitem.Text = "30051";
            rdBoolean = 2;
        }

        private void TileCSHI_Click(object sender, EventArgs e)
        {
            StringBuilder messageBuilder = new StringBuilder(200);
            messageBuilder.Append("30093 Convergys Singapore Holdings Inc - ROHQ:");
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append("  \u2022 GEO: Philippines");
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append("  \u2022 8F Convergys One Building, 6796 Ayala Ave. cor. Salcedo St., Legaspi Village, Makati City 1229");
            textPanePEZA.Text = messageBuilder.ToString();
            lblBUitem.Text = "30093";
            rdBoolean = 3;
        }

        private void TileCMPB_Click(object sender, EventArgs e)
        {
            StringBuilder messageBuilder = new StringBuilder(200);
            messageBuilder.Append("30114 Convergys Malaysia (Philippines) Sdn Bhd - Phil Branch:");
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append("  \u2022 GEO: Philippines");
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append("  \u2022 11F Commerce and Industry Plaza, Upper MCKinley Hill, Fort Bonifacio, Taguig City 1634");
            textPanePEZA.Text = messageBuilder.ToString();
            lblBUitem.Text = "30114";
            rdBoolean = 3;
        }

        private void TileCPI_Click(object sender, EventArgs e)
        {
            StringBuilder messageBuilder = new StringBuilder(200);
            messageBuilder.Append("30238 Convergys Philippines Inc:");
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append("  \u2022 GEO: Philippines");
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append("  \u2022 Basement, Ground, 4th to 9th Floors SLC Building 6797 Ayala Avenue Makati City 1226");
            textPanePEZA.Text = messageBuilder.ToString();
            lblBUitem.Text = "30238";
            rdBoolean = 4;
        }

        private void TileCSPI_Click(object sender, EventArgs e)
        {
            StringBuilder messageBuilder = new StringBuilder(200);
            messageBuilder.Append("30239 Convergys Services Philippines Inc:");
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append("  \u2022 GEO: Philippines");
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append("  \u2022 SM City Clark BPO Building 3 Manuel A. Roxas Highway Brgy Malabanias Angeles City ");
            textPanePEZA.Text = messageBuilder.ToString();
            lblBUitem.Text = "30239";
            rdBoolean = 5;
        }

        private void TileCGSP_Click(object sender, EventArgs e)
        {
            StringBuilder messageBuilder = new StringBuilder(200);
            messageBuilder.Append("30247 Convergys Global Services Philippines - ROHQ:");
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append("  \u2022 GEO: Philippines");
            messageBuilder.Append(Environment.NewLine);
            messageBuilder.Append("  \u2022 8F SLC Building 6797 Ayala Avenue corner VA Rufino St Makati City 1226");
            textPanePEZA.Text = messageBuilder.ToString();
            lblBUitem.Text = "30247";
            rdBoolean = 6;
        }

        private void TileReset_Click(object sender, EventArgs e)
        {
            textPanePEZA.Clear();
            textPaneBIR.Clear();
            txtCRNCY.Clear();
            txtLocation.Clear();
            rdbtnIS.Checked = false;
            rdbtnBS.Checked = false;
            lblDOCUitem.Text = "";
            lblBUitem.Text = "";
        }

        private void rdbtnIS_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbtnIS.Checked == true)
            {
                StringBuilder messageBuilderBIR = new StringBuilder(200);
                messageBuilderBIR.Append("Income Statement:");
                messageBuilderBIR.Append(Environment.NewLine);
                messageBuilderBIR.Append("  \u2022 Income statement document type pertains to the costs and expense data.");
                messageBuilderBIR.Append(Environment.NewLine);
                textPaneBIR.Text = messageBuilderBIR.ToString();
                lblDOCUitem.Text = "Income Statement";
            }
        }

        private void rdbtnBS_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbtnBS.Checked == true)
            {
                StringBuilder messageBuilderBIR = new StringBuilder(200);
                messageBuilderBIR.Append("Balance Sheet:");
                messageBuilderBIR.Append(Environment.NewLine);
                messageBuilderBIR.Append("  \u2022 Balance sheet document type pertains to the statement of Assets and Liabilities.");
                messageBuilderBIR.Append(Environment.NewLine);
                textPaneBIR.Text = messageBuilderBIR.ToString();
                lblDOCUitem.Text = "Balance Sheet";
            }
        }

        private void TileReset_Paint(object sender, PaintEventArgs e)
        {

        }
        
        private void btnPull_Click(object sender, EventArgs e)
        {
            if ((rdbtnBS.Checked == false) && (rdbtnIS.Checked == false))
            {
                MessageBox.Show("No document type selected!");
                return;
            }
            if ((rdbtnPHP.Checked == false) && (rdbtnUSD.Checked == false))
            {
                MessageBox.Show("No currency type selected!");
                return;
            }
            else if (rdBoolean == 0)
            {
                MessageBox.Show("No business unit selected!");
                return;
            }
            passedBU = lblBUitem.Text;
            passedDOCU = lblDOCUitem.Text;
            passedCRNCY = lblCRNCY.Text;
            SelectTable = true;
            this.Close();
        }

        private void Account_FormClosing(object sender, FormClosingEventArgs e)
        {
            if ((FormToShowOnClose != null) && (SelectTable == true))
            {
                SelectTable = false;
                FormToShowOnClose.BusinessUnit = passedBU;
                FormToShowOnClose.DocumentType = passedDOCU;
                FormToShowOnClose.CurrencyType = passedCRNCY;
                FormToShowOnClose.LoadData();
                FormToShowOnClose.Show();
            }
            else { Application.Exit(); }
        }

        private void rdbtnPHP_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbtnPHP.Checked == true)
            {
                StringBuilder messageBuilderBIR = new StringBuilder(200);
                messageBuilderBIR.Append("PHP:");
                messageBuilderBIR.Append(Environment.NewLine);
                messageBuilderBIR.Append("  \u2022 [PHP] preference");
                messageBuilderBIR.Append(Environment.NewLine);
                txtCRNCY.Text = messageBuilderBIR.ToString();
                lblCRNCY.Text = "PHP";
            }
        }

        private void rdbtnUSD_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbtnUSD.Checked == true)
            {
                StringBuilder messageBuilderBIR = new StringBuilder(200);
                messageBuilderBIR.Append("USD:");
                messageBuilderBIR.Append(Environment.NewLine);
                messageBuilderBIR.Append("  \u2022 [USD] preference");
                messageBuilderBIR.Append(Environment.NewLine);
                txtCRNCY.Text = messageBuilderBIR.ToString();
                lblCRNCY.Text = "USD";
            }
        }

        private void btnMerge_Click(object sender, EventArgs e)
        {
            var form = new Merge();
            form.ShowDialog();
        }

        private void txtLocation_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Here");
            OpenFileDialog OFD = new OpenFileDialog();
            OFD.Filter = "Access Database|*.accdb|All files(*.*)|*.*";
            OFD.ShowDialog();
            string FoundFile = OFD.FileName;
            txtLocation.Text = FoundFile;
            if (!string.IsNullOrEmpty(FoundFile))
            {
                try
                {
                    string Src = FoundFile;
                    string Init = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + Src;
                    ChangeConnectionString("Con", Init, "System.Data.OleDb", "Query Refiner");
                    lblStatus.Text = "Successfully replaced the data source!";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error reading file: " + ex.Message);
                    return;
                }
            }
        }

        public static bool ChangeConnectionString(string Name, string value, string providerName, string AppName)
        {
            bool retVal = false;
            try
            {
                string FILE_NAME = string.Concat(Application.StartupPath, "\\", AppName.Trim(), ".exe.Config"); //the application configuration file name
                XmlTextReader reader = new XmlTextReader(FILE_NAME);
                XmlDocument doc = new XmlDocument();
                doc.Load(reader);
                reader.Close();
                string nodeRoute = string.Concat("connectionStrings/add");

                XmlNode cnnStr = null;
                XmlElement root = doc.DocumentElement;
                XmlNodeList Settings = root.SelectNodes(nodeRoute);

                for (int i = 0; i < Settings.Count; i++)
                {
                    cnnStr = Settings[i];
                    if (cnnStr.Attributes["name"].Value.Equals(Name))
                        break;
                    cnnStr = null;
                }

                cnnStr.Attributes["connectionString"].Value = value;
                cnnStr.Attributes["providerName"].Value = providerName;
                doc.Save(FILE_NAME);
                retVal = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                retVal = false;
                //Handle the Exception as you like
            }
            finally
            {
                MessageBox.Show(
                "The program will restart in order for the changes to take effect, you may experience some error messages!", "Data source was changed",
                MessageBoxButtons.OK, MessageBoxIcon.Stop);
                try
                {
                    Application.Exit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Application runtime is ending... " + ex.Message);
                }
            }
            return retVal;
        }

        private void txtLocation_Enter(object sender, EventArgs e)
        {

        }

        private void lblExit_Click(object sender, EventArgs e)
        {
            DialogResult DR = new DialogResult();
            OpenFileDialog OFD = new OpenFileDialog();
            DR = MessageBox.Show("Are you sure you want to close the system?", "End Runtime", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (DR == DialogResult.Yes)
            {

            }
        }
    }
}
