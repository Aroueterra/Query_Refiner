namespace Query_Refiner
{
    partial class Account
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Account));
            this.txtLocation = new MetroFramework.Controls.MetroTextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.lblHistory = new MetroFramework.Controls.MetroLabel();
            this.lblStatus = new MetroFramework.Controls.MetroLabel();
            this.lblCRNCY = new MetroFramework.Controls.MetroLabel();
            this.lblDOCUitem = new MetroFramework.Controls.MetroLabel();
            this.lblBUitem = new MetroFramework.Controls.MetroLabel();
            this.metroPanel1 = new MetroFramework.Controls.MetroPanel();
            this.metroPanel2 = new MetroFramework.Controls.MetroPanel();
            this.btnPull = new System.Windows.Forms.Button();
            this.txtCRNCY = new System.Windows.Forms.TextBox();
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.rdbtnIS = new MetroFramework.Controls.MetroRadioButton();
            this.rdbtnBS = new MetroFramework.Controls.MetroRadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdbtnPHP = new MetroFramework.Controls.MetroRadioButton();
            this.rdbtnUSD = new MetroFramework.Controls.MetroRadioButton();
            this.textPaneBIR = new System.Windows.Forms.TextBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.btnMerge = new System.Windows.Forms.Button();
            this.textPanePEZA = new System.Windows.Forms.TextBox();
            this.metroPanel3 = new MetroFramework.Controls.MetroPanel();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.TileCGSP = new MetroFramework.Controls.MetroTile();
            this.TileCSPI = new MetroFramework.Controls.MetroTile();
            this.TileCPI = new MetroFramework.Controls.MetroTile();
            this.TileCMPB = new MetroFramework.Controls.MetroTile();
            this.TileCPSC = new MetroFramework.Controls.MetroTile();
            this.TileReset = new System.Windows.Forms.Panel();
            this.TileCSHI = new MetroFramework.Controls.MetroTile();
            this.TileERMI = new MetroFramework.Controls.MetroTile();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.FadeTimer = new System.Windows.Forms.Timer(this.components);
            this.lblExit = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.metroPanel2.SuspendLayout();
            this.flowLayoutPanel2.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.panel3.SuspendLayout();
            this.metroPanel3.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtLocation
            // 
            // 
            // 
            // 
            this.txtLocation.CustomButton.Image = null;
            this.txtLocation.CustomButton.Location = new System.Drawing.Point(737, 1);
            this.txtLocation.CustomButton.Name = "";
            this.txtLocation.CustomButton.Size = new System.Drawing.Size(21, 21);
            this.txtLocation.CustomButton.Style = MetroFramework.MetroColorStyle.Blue;
            this.txtLocation.CustomButton.TabIndex = 1;
            this.txtLocation.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light;
            this.txtLocation.CustomButton.UseSelectable = true;
            this.txtLocation.DisplayIcon = true;
            this.txtLocation.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.txtLocation.Icon = ((System.Drawing.Image)(resources.GetObject("txtLocation.Icon")));
            this.txtLocation.Lines = new string[] {
        "Database Location"};
            this.txtLocation.Location = new System.Drawing.Point(0, 73);
            this.txtLocation.MaxLength = 32767;
            this.txtLocation.Name = "txtLocation";
            this.txtLocation.PasswordChar = '\0';
            this.txtLocation.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.txtLocation.SelectedText = "";
            this.txtLocation.SelectionLength = 0;
            this.txtLocation.SelectionStart = 0;
            this.txtLocation.ShortcutsEnabled = true;
            this.txtLocation.ShowButton = true;
            this.txtLocation.Size = new System.Drawing.Size(759, 23);
            this.txtLocation.TabIndex = 0;
            this.txtLocation.Text = "Database Location";
            this.txtLocation.UseSelectable = true;
            this.txtLocation.WaterMarkColor = System.Drawing.Color.FromArgb(((int)(((byte)(109)))), ((int)(((byte)(109)))), ((int)(((byte)(109)))));
            this.txtLocation.WaterMarkFont = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Pixel);
            this.txtLocation.Click += new System.EventHandler(this.txtLocation_Click);
            this.txtLocation.Enter += new System.EventHandler(this.txtLocation_Enter);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Controls.Add(this.lblStatus);
            this.panel1.Controls.Add(this.txtLocation);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(20, 383);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(759, 96);
            this.panel1.TabIndex = 1;
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.lblHistory);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(0, 22);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(759, 51);
            this.panel2.TabIndex = 2;
            // 
            // lblHistory
            // 
            this.lblHistory.AutoSize = true;
            this.lblHistory.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lblHistory.FontSize = MetroFramework.MetroLabelSize.Small;
            this.lblHistory.Location = new System.Drawing.Point(0, 34);
            this.lblHistory.Name = "lblHistory";
            this.lblHistory.Size = new System.Drawing.Size(574, 15);
            this.lblHistory.TabIndex = 3;
            this.lblHistory.Text = "C:\\Users\\Aroueterra\\Documents\\Visual Studio 2015\\Bako apu\\Query Listener [1-13-17" +
    "]\\Query Listener\\ACDB.accdb";
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblStatus.Location = new System.Drawing.Point(0, 0);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(289, 19);
            this.lblStatus.TabIndex = 1;
            this.lblStatus.Text = "Your current database is set to the following file:";
            // 
            // lblCRNCY
            // 
            this.lblCRNCY.BackColor = System.Drawing.Color.White;
            this.lblCRNCY.FontSize = MetroFramework.MetroLabelSize.Small;
            this.lblCRNCY.Location = new System.Drawing.Point(6, 51);
            this.lblCRNCY.Name = "lblCRNCY";
            this.lblCRNCY.Size = new System.Drawing.Size(123, 19);
            this.lblCRNCY.TabIndex = 5;
            this.lblCRNCY.Text = "PHP";
            this.lblCRNCY.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblDOCUitem
            // 
            this.lblDOCUitem.FontSize = MetroFramework.MetroLabelSize.Small;
            this.lblDOCUitem.Location = new System.Drawing.Point(6, 34);
            this.lblDOCUitem.Name = "lblDOCUitem";
            this.lblDOCUitem.Size = new System.Drawing.Size(124, 19);
            this.lblDOCUitem.TabIndex = 4;
            this.lblDOCUitem.Text = "Income Statement";
            this.lblDOCUitem.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblBUitem
            // 
            this.lblBUitem.FontSize = MetroFramework.MetroLabelSize.Small;
            this.lblBUitem.Location = new System.Drawing.Point(6, 18);
            this.lblBUitem.Name = "lblBUitem";
            this.lblBUitem.Size = new System.Drawing.Size(122, 19);
            this.lblBUitem.TabIndex = 3;
            this.lblBUitem.Text = "[BU Code]";
            this.lblBUitem.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // metroPanel1
            // 
            this.metroPanel1.BackColor = System.Drawing.Color.MediumTurquoise;
            this.metroPanel1.HorizontalScrollbarBarColor = true;
            this.metroPanel1.HorizontalScrollbarHighlightOnWheel = false;
            this.metroPanel1.HorizontalScrollbarSize = 10;
            this.metroPanel1.Location = new System.Drawing.Point(0, 57);
            this.metroPanel1.Name = "metroPanel1";
            this.metroPanel1.Size = new System.Drawing.Size(791, 3);
            this.metroPanel1.TabIndex = 2;
            this.metroPanel1.UseCustomBackColor = true;
            this.metroPanel1.VerticalScrollbarBarColor = true;
            this.metroPanel1.VerticalScrollbarHighlightOnWheel = false;
            this.metroPanel1.VerticalScrollbarSize = 10;
            // 
            // metroPanel2
            // 
            this.metroPanel2.Controls.Add(this.btnPull);
            this.metroPanel2.Controls.Add(this.txtCRNCY);
            this.metroPanel2.Controls.Add(this.flowLayoutPanel2);
            this.metroPanel2.Controls.Add(this.textPaneBIR);
            this.metroPanel2.Controls.Add(this.panel3);
            this.metroPanel2.Controls.Add(this.textPanePEZA);
            this.metroPanel2.Controls.Add(this.metroPanel3);
            this.metroPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.metroPanel2.HorizontalScrollbarBarColor = true;
            this.metroPanel2.HorizontalScrollbarHighlightOnWheel = false;
            this.metroPanel2.HorizontalScrollbarSize = 10;
            this.metroPanel2.Location = new System.Drawing.Point(20, 60);
            this.metroPanel2.Margin = new System.Windows.Forms.Padding(3, 3, 0, 3);
            this.metroPanel2.Name = "metroPanel2";
            this.metroPanel2.Size = new System.Drawing.Size(759, 323);
            this.metroPanel2.TabIndex = 5;
            this.metroPanel2.VerticalScrollbarBarColor = true;
            this.metroPanel2.VerticalScrollbarHighlightOnWheel = false;
            this.metroPanel2.VerticalScrollbarSize = 10;
            // 
            // btnPull
            // 
            this.btnPull.BackColor = System.Drawing.Color.White;
            this.btnPull.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnPull.Image = ((System.Drawing.Image)(resources.GetObject("btnPull.Image")));
            this.btnPull.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnPull.Location = new System.Drawing.Point(478, 274);
            this.btnPull.Margin = new System.Windows.Forms.Padding(0);
            this.btnPull.Name = "btnPull";
            this.btnPull.Size = new System.Drawing.Size(130, 49);
            this.btnPull.TabIndex = 14;
            this.btnPull.Text = "Pull Table";
            this.btnPull.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnPull.UseVisualStyleBackColor = false;
            this.btnPull.Click += new System.EventHandler(this.btnPull_Click);
            // 
            // txtCRNCY
            // 
            this.txtCRNCY.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txtCRNCY.Enabled = false;
            this.txtCRNCY.Font = new System.Drawing.Font("Century Gothic", 6.75F);
            this.txtCRNCY.Location = new System.Drawing.Point(621, 234);
            this.txtCRNCY.Multiline = true;
            this.txtCRNCY.Name = "txtCRNCY";
            this.txtCRNCY.Size = new System.Drawing.Size(138, 37);
            this.txtCRNCY.TabIndex = 12;
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.flowLayoutPanel2.BackColor = System.Drawing.Color.Transparent;
            this.flowLayoutPanel2.Controls.Add(this.groupBox2);
            this.flowLayoutPanel2.Controls.Add(this.groupBox1);
            this.flowLayoutPanel2.Location = new System.Drawing.Point(467, 131);
            this.flowLayoutPanel2.Margin = new System.Windows.Forms.Padding(0);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Padding = new System.Windows.Forms.Padding(1, 0, 0, 0);
            this.flowLayoutPanel2.Size = new System.Drawing.Size(148, 140);
            this.flowLayoutPanel2.TabIndex = 11;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.rdbtnIS);
            this.groupBox2.Controls.Add(this.rdbtnBS);
            this.groupBox2.Font = new System.Drawing.Font("Century Gothic", 8.25F);
            this.groupBox2.Location = new System.Drawing.Point(4, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(144, 65);
            this.groupBox2.TabIndex = 14;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Select the document:";
            // 
            // rdbtnIS
            // 
            this.rdbtnIS.AutoSize = true;
            this.rdbtnIS.Checked = true;
            this.rdbtnIS.Location = new System.Drawing.Point(6, 20);
            this.rdbtnIS.Name = "rdbtnIS";
            this.rdbtnIS.Size = new System.Drawing.Size(120, 15);
            this.rdbtnIS.TabIndex = 0;
            this.rdbtnIS.TabStop = true;
            this.rdbtnIS.Text = "Income Statement";
            this.rdbtnIS.UseSelectable = true;
            this.rdbtnIS.CheckedChanged += new System.EventHandler(this.rdbtnIS_CheckedChanged);
            // 
            // rdbtnBS
            // 
            this.rdbtnBS.AutoSize = true;
            this.rdbtnBS.Location = new System.Drawing.Point(6, 41);
            this.rdbtnBS.Name = "rdbtnBS";
            this.rdbtnBS.Size = new System.Drawing.Size(96, 15);
            this.rdbtnBS.TabIndex = 1;
            this.rdbtnBS.Text = "Balance Sheet";
            this.rdbtnBS.UseSelectable = true;
            this.rdbtnBS.CheckedChanged += new System.EventHandler(this.rdbtnBS_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rdbtnPHP);
            this.groupBox1.Controls.Add(this.rdbtnUSD);
            this.groupBox1.Font = new System.Drawing.Font("Century Gothic", 8.25F);
            this.groupBox1.Location = new System.Drawing.Point(4, 74);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(144, 59);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Select the currency:";
            // 
            // rdbtnPHP
            // 
            this.rdbtnPHP.AutoSize = true;
            this.rdbtnPHP.Checked = true;
            this.rdbtnPHP.Location = new System.Drawing.Point(6, 19);
            this.rdbtnPHP.Margin = new System.Windows.Forms.Padding(3, 3, 100, 3);
            this.rdbtnPHP.Name = "rdbtnPHP";
            this.rdbtnPHP.Size = new System.Drawing.Size(46, 15);
            this.rdbtnPHP.TabIndex = 9;
            this.rdbtnPHP.TabStop = true;
            this.rdbtnPHP.Text = "PHP";
            this.rdbtnPHP.UseSelectable = true;
            this.rdbtnPHP.CheckedChanged += new System.EventHandler(this.rdbtnPHP_CheckedChanged);
            // 
            // rdbtnUSD
            // 
            this.rdbtnUSD.AutoSize = true;
            this.rdbtnUSD.Location = new System.Drawing.Point(6, 40);
            this.rdbtnUSD.Name = "rdbtnUSD";
            this.rdbtnUSD.Size = new System.Drawing.Size(45, 15);
            this.rdbtnUSD.TabIndex = 10;
            this.rdbtnUSD.Text = "USD";
            this.rdbtnUSD.UseSelectable = true;
            this.rdbtnUSD.CheckedChanged += new System.EventHandler(this.rdbtnUSD_CheckedChanged);
            // 
            // textPaneBIR
            // 
            this.textPaneBIR.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textPaneBIR.Enabled = false;
            this.textPaneBIR.Font = new System.Drawing.Font("Century Gothic", 7.25F);
            this.textPaneBIR.Location = new System.Drawing.Point(621, 131);
            this.textPaneBIR.Multiline = true;
            this.textPaneBIR.Name = "textPaneBIR";
            this.textPaneBIR.Size = new System.Drawing.Size(138, 100);
            this.textPaneBIR.TabIndex = 9;
            // 
            // panel3
            // 
            this.panel3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.btnMerge);
            this.panel3.Location = new System.Drawing.Point(626, 272);
            this.panel3.Margin = new System.Windows.Forms.Padding(0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(132, 51);
            this.panel3.TabIndex = 9;
            // 
            // btnMerge
            // 
            this.btnMerge.BackColor = System.Drawing.Color.White;
            this.btnMerge.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnMerge.FlatAppearance.BorderSize = 0;
            this.btnMerge.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMerge.Image = ((System.Drawing.Image)(resources.GetObject("btnMerge.Image")));
            this.btnMerge.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnMerge.Location = new System.Drawing.Point(0, 0);
            this.btnMerge.Margin = new System.Windows.Forms.Padding(0);
            this.btnMerge.Name = "btnMerge";
            this.btnMerge.Size = new System.Drawing.Size(130, 49);
            this.btnMerge.TabIndex = 15;
            this.btnMerge.Text = "Merge";
            this.btnMerge.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnMerge.UseVisualStyleBackColor = false;
            this.btnMerge.Click += new System.EventHandler(this.btnMerge_Click);
            // 
            // textPanePEZA
            // 
            this.textPanePEZA.Dock = System.Windows.Forms.DockStyle.Left;
            this.textPanePEZA.Enabled = false;
            this.textPanePEZA.Font = new System.Drawing.Font("Century Gothic", 8.25F);
            this.textPanePEZA.Location = new System.Drawing.Point(0, 131);
            this.textPanePEZA.Margin = new System.Windows.Forms.Padding(3, 3, 10, 3);
            this.textPanePEZA.Multiline = true;
            this.textPanePEZA.Name = "textPanePEZA";
            this.textPanePEZA.Size = new System.Drawing.Size(471, 192);
            this.textPanePEZA.TabIndex = 8;
            // 
            // metroPanel3
            // 
            this.metroPanel3.Controls.Add(this.tableLayoutPanel1);
            this.metroPanel3.Controls.Add(this.flowLayoutPanel1);
            this.metroPanel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.metroPanel3.HorizontalScrollbarBarColor = true;
            this.metroPanel3.HorizontalScrollbarHighlightOnWheel = false;
            this.metroPanel3.HorizontalScrollbarSize = 10;
            this.metroPanel3.Location = new System.Drawing.Point(0, 0);
            this.metroPanel3.Name = "metroPanel3";
            this.metroPanel3.Size = new System.Drawing.Size(759, 131);
            this.metroPanel3.TabIndex = 5;
            this.metroPanel3.VerticalScrollbarBarColor = true;
            this.metroPanel3.VerticalScrollbarHighlightOnWheel = false;
            this.metroPanel3.VerticalScrollbarSize = 10;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 4;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.Controls.Add(this.TileCGSP, 3, 1);
            this.tableLayoutPanel1.Controls.Add(this.TileCSPI, 2, 1);
            this.tableLayoutPanel1.Controls.Add(this.TileCPI, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.TileCMPB, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.TileCPSC, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.TileReset, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.TileCSHI, 3, 0);
            this.tableLayoutPanel1.Controls.Add(this.TileERMI, 2, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(621, 131);
            this.tableLayoutPanel1.TabIndex = 11;
            // 
            // TileCGSP
            // 
            this.TileCGSP.ActiveControl = null;
            this.TileCGSP.Location = new System.Drawing.Point(468, 68);
            this.TileCGSP.Name = "TileCGSP";
            this.TileCGSP.Size = new System.Drawing.Size(149, 59);
            this.TileCGSP.Style = MetroFramework.MetroColorStyle.Teal;
            this.TileCGSP.TabIndex = 18;
            this.TileCGSP.Text = "[30247] - CGSP";
            this.TileCGSP.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
            this.TileCGSP.TileTextFontWeight = MetroFramework.MetroTileTextWeight.Regular;
            this.TileCGSP.UseSelectable = true;
            this.TileCGSP.Click += new System.EventHandler(this.TileCGSP_Click);
            // 
            // TileCSPI
            // 
            this.TileCSPI.ActiveControl = null;
            this.TileCSPI.Location = new System.Drawing.Point(313, 68);
            this.TileCSPI.Name = "TileCSPI";
            this.TileCSPI.Size = new System.Drawing.Size(149, 59);
            this.TileCSPI.Style = MetroFramework.MetroColorStyle.Teal;
            this.TileCSPI.TabIndex = 17;
            this.TileCSPI.Text = "[30239] - CSPI";
            this.TileCSPI.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
            this.TileCSPI.TileTextFontWeight = MetroFramework.MetroTileTextWeight.Regular;
            this.TileCSPI.UseSelectable = true;
            this.TileCSPI.Click += new System.EventHandler(this.TileCSPI_Click);
            // 
            // TileCPI
            // 
            this.TileCPI.ActiveControl = null;
            this.TileCPI.Location = new System.Drawing.Point(158, 68);
            this.TileCPI.Name = "TileCPI";
            this.TileCPI.Size = new System.Drawing.Size(149, 59);
            this.TileCPI.Style = MetroFramework.MetroColorStyle.Teal;
            this.TileCPI.TabIndex = 16;
            this.TileCPI.Text = "[30238] - CPI";
            this.TileCPI.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
            this.TileCPI.TileTextFontWeight = MetroFramework.MetroTileTextWeight.Regular;
            this.TileCPI.UseSelectable = true;
            this.TileCPI.Click += new System.EventHandler(this.TileCPI_Click);
            // 
            // TileCMPB
            // 
            this.TileCMPB.ActiveControl = null;
            this.TileCMPB.Location = new System.Drawing.Point(3, 68);
            this.TileCMPB.Name = "TileCMPB";
            this.TileCMPB.Size = new System.Drawing.Size(149, 59);
            this.TileCMPB.Style = MetroFramework.MetroColorStyle.Teal;
            this.TileCMPB.TabIndex = 15;
            this.TileCMPB.Text = "[30114] - CMPB";
            this.TileCMPB.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
            this.TileCMPB.TileTextFontWeight = MetroFramework.MetroTileTextWeight.Regular;
            this.TileCMPB.UseSelectable = true;
            this.TileCMPB.Click += new System.EventHandler(this.TileCMPB_Click);
            // 
            // TileCPSC
            // 
            this.TileCPSC.ActiveControl = null;
            this.TileCPSC.Location = new System.Drawing.Point(158, 3);
            this.TileCPSC.Name = "TileCPSC";
            this.TileCPSC.Size = new System.Drawing.Size(149, 59);
            this.TileCPSC.Style = MetroFramework.MetroColorStyle.Teal;
            this.TileCPSC.TabIndex = 14;
            this.TileCPSC.Text = "[30023] - CPSC";
            this.TileCPSC.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
            this.TileCPSC.TileTextFontWeight = MetroFramework.MetroTileTextWeight.Regular;
            this.TileCPSC.UseSelectable = true;
            this.TileCPSC.Click += new System.EventHandler(this.TileCPSC_Click);
            // 
            // TileReset
            // 
            this.TileReset.BackColor = System.Drawing.Color.DarkTurquoise;
            this.TileReset.Dock = System.Windows.Forms.DockStyle.Fill;
            this.TileReset.Location = new System.Drawing.Point(3, 3);
            this.TileReset.Name = "TileReset";
            this.TileReset.Size = new System.Drawing.Size(149, 59);
            this.TileReset.TabIndex = 12;
            this.TileReset.Click += new System.EventHandler(this.TileReset_Click);
            this.TileReset.Paint += new System.Windows.Forms.PaintEventHandler(this.TileReset_Paint);
            // 
            // TileCSHI
            // 
            this.TileCSHI.ActiveControl = null;
            this.TileCSHI.Location = new System.Drawing.Point(468, 3);
            this.TileCSHI.Name = "TileCSHI";
            this.TileCSHI.Size = new System.Drawing.Size(149, 59);
            this.TileCSHI.Style = MetroFramework.MetroColorStyle.Teal;
            this.TileCSHI.TabIndex = 13;
            this.TileCSHI.Text = "[30093] - CSHI";
            this.TileCSHI.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
            this.TileCSHI.TileTextFontWeight = MetroFramework.MetroTileTextWeight.Regular;
            this.TileCSHI.UseSelectable = true;
            this.TileCSHI.Click += new System.EventHandler(this.TileCSHI_Click);
            // 
            // TileERMI
            // 
            this.TileERMI.ActiveControl = null;
            this.TileERMI.Location = new System.Drawing.Point(313, 3);
            this.TileERMI.Name = "TileERMI";
            this.TileERMI.Size = new System.Drawing.Size(149, 59);
            this.TileERMI.Style = MetroFramework.MetroColorStyle.Teal;
            this.TileERMI.TabIndex = 11;
            this.TileERMI.Text = "[30051] - ERMI";
            this.TileERMI.TileTextFontSize = MetroFramework.MetroTileTextSize.Tall;
            this.TileERMI.TileTextFontWeight = MetroFramework.MetroTileTextWeight.Regular;
            this.TileERMI.UseSelectable = true;
            this.TileERMI.Click += new System.EventHandler(this.TileERMI_Click);
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.BackColor = System.Drawing.Color.Transparent;
            this.flowLayoutPanel1.Controls.Add(this.label4);
            this.flowLayoutPanel1.Controls.Add(this.groupBox3);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Right;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(621, 0);
            this.flowLayoutPanel1.Margin = new System.Windows.Forms.Padding(0);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(138, 131);
            this.flowLayoutPanel1.TabIndex = 10;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Century Gothic", 8.25F);
            this.label4.Location = new System.Drawing.Point(3, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(127, 48);
            this.label4.TabIndex = 9;
            this.label4.Text = "Build the criteria for selecting a table from the database.";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.lblCRNCY);
            this.groupBox3.Controls.Add(this.lblDOCUitem);
            this.groupBox3.Controls.Add(this.lblBUitem);
            this.groupBox3.Font = new System.Drawing.Font("Century Gothic", 8.25F);
            this.groupBox3.Location = new System.Drawing.Point(3, 51);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(132, 76);
            this.groupBox3.TabIndex = 14;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Your Selection:";
            // 
            // FadeTimer
            // 
            this.FadeTimer.Interval = 200;
            this.FadeTimer.Tick += new System.EventHandler(this.FadeTimer_Tick);
            // 
            // lblExit
            // 
            this.lblExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblExit.AutoSize = true;
            this.lblExit.Font = new System.Drawing.Font("Century Gothic", 9.25F);
            this.lblExit.Location = new System.Drawing.Point(644, 37);
            this.lblExit.Name = "lblExit";
            this.lblExit.Size = new System.Drawing.Size(116, 17);
            this.lblExit.TabIndex = 6;
            this.lblExit.Text = "Exit Application?";
            this.lblExit.Visible = false;
            this.lblExit.Click += new System.EventHandler(this.lblExit_Click);
            // 
            // Account
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BorderStyle = MetroFramework.Forms.MetroFormBorderStyle.FixedSingle;
            this.ClientSize = new System.Drawing.Size(799, 499);
            this.Controls.Add(this.lblExit);
            this.Controls.Add(this.metroPanel2);
            this.Controls.Add(this.metroPanel1);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.MinimumSize = new System.Drawing.Size(805, 505);
            this.Name = "Account";
            this.Resizable = false;
            this.ShadowType = MetroFramework.Forms.MetroFormShadowType.AeroShadow;
            this.Style = MetroFramework.MetroColorStyle.Teal;
            this.Text = "Business unit operations:";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Account_FormClosing);
            this.Load += new System.EventHandler(this.Account_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.metroPanel2.ResumeLayout(false);
            this.metroPanel2.PerformLayout();
            this.flowLayoutPanel2.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.metroPanel3.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private MetroFramework.Controls.MetroTextBox txtLocation;
        private System.Windows.Forms.Panel panel1;
        private MetroFramework.Controls.MetroLabel lblStatus;
        private MetroFramework.Controls.MetroPanel metroPanel1;
        private MetroFramework.Controls.MetroPanel metroPanel2;
        private MetroFramework.Controls.MetroPanel metroPanel3;
        private System.Windows.Forms.TextBox textPanePEZA;
        private System.Windows.Forms.TextBox textPaneBIR;
        private System.Windows.Forms.Panel panel2;
        private MetroFramework.Controls.MetroLabel lblHistory;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private MetroFramework.Controls.MetroRadioButton rdbtnIS;
        private MetroFramework.Controls.MetroRadioButton rdbtnBS;
        private System.Windows.Forms.Timer FadeTimer;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private MetroFramework.Controls.MetroTile TileCGSP;
        private MetroFramework.Controls.MetroTile TileCSPI;
        private MetroFramework.Controls.MetroTile TileCPI;
        private MetroFramework.Controls.MetroTile TileCMPB;
        private MetroFramework.Controls.MetroTile TileCPSC;
        private System.Windows.Forms.Panel TileReset;
        private MetroFramework.Controls.MetroTile TileCSHI;
        private MetroFramework.Controls.MetroTile TileERMI;
        private MetroFramework.Controls.MetroLabel lblDOCUitem;
        private MetroFramework.Controls.MetroLabel lblBUitem;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button btnPull;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel2;
        private MetroFramework.Controls.MetroRadioButton rdbtnPHP;
        private MetroFramework.Controls.MetroRadioButton rdbtnUSD;
        private System.Windows.Forms.Label label4;
        private MetroFramework.Controls.MetroLabel lblCRNCY;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtCRNCY;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btnMerge;
        private System.Windows.Forms.Label lblExit;
    }
}