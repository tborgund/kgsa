namespace KGSA
{
    partial class FormKrav
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormKrav));
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.dataGridViewSk = new System.Windows.Forms.DataGridView();
            this.Id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Avd = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Provisjon = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Selgerkode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Navn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Avdeling = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Finans = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Strom = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TA = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Rtgsa = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.bindingNavigatorSk = new System.Windows.Forms.BindingNavigator(this.components);
            this.bindingNavigatorCountItem = new System.Windows.Forms.ToolStripLabel();
            this.bindingNavigatorMoveFirstItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMovePreviousItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorPositionItem = new System.Windows.Forms.ToolStripTextBox();
            this.bindingNavigatorSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorMoveNextItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMoveLastItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
            this.toolStripComboBoxSkFilter = new System.Windows.Forms.ToolStripComboBox();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.panel4 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.groupKrav = new System.Windows.Forms.GroupBox();
            this.checkKravMtdShowTarget = new System.Windows.Forms.CheckBox();
            this.groupBox21 = new System.Windows.Forms.GroupBox();
            this.checkKravAntallFinans = new System.Windows.Forms.CheckBox();
            this.checkKravAntallMod = new System.Windows.Forms.CheckBox();
            this.checkKravAntallStrom = new System.Windows.Forms.CheckBox();
            this.checkKravAntallRtgsa = new System.Windows.Forms.CheckBox();
            this.checkKravRtgsa = new System.Windows.Forms.CheckBox();
            this.checkKravStrom = new System.Windows.Forms.CheckBox();
            this.checkKravMod = new System.Windows.Forms.CheckBox();
            this.checkKravFinans = new System.Windows.Forms.CheckBox();
            this.label80 = new System.Windows.Forms.Label();
            this.checkKravMTD = new System.Windows.Forms.CheckBox();
            this.label84 = new System.Windows.Forms.Label();
            this.label83 = new System.Windows.Forms.Label();
            this.checkOversiktVisKrav = new System.Windows.Forms.CheckBox();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.bindingSourceSk = new System.Windows.Forms.BindingSource(this.components);
            this.panel1.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewSk)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingNavigatorSk)).BeginInit();
            this.bindingNavigatorSk.SuspendLayout();
            this.panel2.SuspendLayout();
            this.groupKrav.SuspendLayout();
            this.groupBox21.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSourceSk)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Controls.Add(this.panel4);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(776, 561);
            this.panel1.TabIndex = 0;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.dataGridViewSk);
            this.panel3.Controls.Add(this.bindingNavigatorSk);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.Padding = new System.Windows.Forms.Padding(3);
            this.panel3.Size = new System.Drawing.Size(776, 541);
            this.panel3.TabIndex = 0;
            // 
            // dataGridViewSk
            // 
            this.dataGridViewSk.AllowUserToAddRows = false;
            this.dataGridViewSk.AllowUserToDeleteRows = false;
            this.dataGridViewSk.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dataGridViewSk.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewSk.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Id,
            this.Avd,
            this.Provisjon,
            this.Selgerkode,
            this.Navn,
            this.Avdeling,
            this.Finans,
            this.Strom,
            this.TA,
            this.Rtgsa});
            this.dataGridViewSk.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewSk.Location = new System.Drawing.Point(3, 28);
            this.dataGridViewSk.Name = "dataGridViewSk";
            this.dataGridViewSk.Size = new System.Drawing.Size(770, 510);
            this.dataGridViewSk.TabIndex = 1;
            this.dataGridViewSk.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewSk_CellEndEdit);
            this.dataGridViewSk.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dataGridViewSk_CellValidating);
            this.dataGridViewSk.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(this.dataGridViewSk_RowsAdded);
            // 
            // Id
            // 
            this.Id.DataPropertyName = "Id";
            this.Id.HeaderText = "id";
            this.Id.Name = "Id";
            this.Id.ReadOnly = true;
            this.Id.Visible = false;
            // 
            // Avd
            // 
            this.Avd.DataPropertyName = "Avdeling";
            this.Avd.HeaderText = "Avd";
            this.Avd.Name = "Avd";
            this.Avd.Visible = false;
            // 
            // Provisjon
            // 
            this.Provisjon.DataPropertyName = "Provisjon";
            this.Provisjon.HeaderText = "Provisjon";
            this.Provisjon.Name = "Provisjon";
            this.Provisjon.Visible = false;
            // 
            // Selgerkode
            // 
            this.Selgerkode.DataPropertyName = "Selgerkode";
            this.Selgerkode.HeaderText = "Selgerkode";
            this.Selgerkode.MinimumWidth = 100;
            this.Selgerkode.Name = "Selgerkode";
            this.Selgerkode.ReadOnly = true;
            this.Selgerkode.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            // 
            // Navn
            // 
            this.Navn.DataPropertyName = "Navn";
            this.Navn.HeaderText = "Navn";
            this.Navn.Name = "Navn";
            this.Navn.ReadOnly = true;
            // 
            // Avdeling
            // 
            this.Avdeling.DataPropertyName = "Kategori";
            this.Avdeling.HeaderText = "Avdeling";
            this.Avdeling.MinimumWidth = 75;
            this.Avdeling.Name = "Avdeling";
            this.Avdeling.ReadOnly = true;
            this.Avdeling.Width = 75;
            // 
            // Finans
            // 
            this.Finans.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Finans.DataPropertyName = "FinansKrav";
            this.Finans.FillWeight = 67.85257F;
            this.Finans.HeaderText = "Krav: Finans";
            this.Finans.MaxInputLength = 6;
            this.Finans.Name = "Finans";
            // 
            // Strom
            // 
            this.Strom.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Strom.DataPropertyName = "StromKrav";
            this.Strom.FillWeight = 67.85257F;
            this.Strom.HeaderText = "Krav: Strøm";
            this.Strom.MaxInputLength = 6;
            this.Strom.Name = "Strom";
            // 
            // TA
            // 
            this.TA.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.TA.DataPropertyName = "ModKrav";
            this.TA.FillWeight = 67.85257F;
            this.TA.HeaderText = "Krav: TA";
            this.TA.MaxInputLength = 6;
            this.TA.Name = "TA";
            // 
            // Rtgsa
            // 
            this.Rtgsa.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Rtgsa.DataPropertyName = "RtgsaKrav";
            this.Rtgsa.FillWeight = 67.85257F;
            this.Rtgsa.HeaderText = "Krav: RTG/SA";
            this.Rtgsa.MaxInputLength = 6;
            this.Rtgsa.Name = "Rtgsa";
            // 
            // bindingNavigatorSk
            // 
            this.bindingNavigatorSk.AddNewItem = null;
            this.bindingNavigatorSk.CountItem = this.bindingNavigatorCountItem;
            this.bindingNavigatorSk.DeleteItem = null;
            this.bindingNavigatorSk.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.bindingNavigatorMoveFirstItem,
            this.bindingNavigatorMovePreviousItem,
            this.bindingNavigatorSeparator,
            this.bindingNavigatorPositionItem,
            this.bindingNavigatorCountItem,
            this.bindingNavigatorSeparator1,
            this.bindingNavigatorMoveNextItem,
            this.bindingNavigatorMoveLastItem,
            this.bindingNavigatorSeparator2,
            this.toolStripLabel1,
            this.toolStripComboBoxSkFilter,
            this.toolStripButton1});
            this.bindingNavigatorSk.Location = new System.Drawing.Point(3, 3);
            this.bindingNavigatorSk.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.bindingNavigatorSk.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.bindingNavigatorSk.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.bindingNavigatorSk.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.bindingNavigatorSk.Name = "bindingNavigatorSk";
            this.bindingNavigatorSk.PositionItem = this.bindingNavigatorPositionItem;
            this.bindingNavigatorSk.Size = new System.Drawing.Size(770, 25);
            this.bindingNavigatorSk.TabIndex = 0;
            this.bindingNavigatorSk.Text = "bindingNavigator1";
            // 
            // bindingNavigatorCountItem
            // 
            this.bindingNavigatorCountItem.Name = "bindingNavigatorCountItem";
            this.bindingNavigatorCountItem.Size = new System.Drawing.Size(35, 22);
            this.bindingNavigatorCountItem.Text = "of {0}";
            this.bindingNavigatorCountItem.ToolTipText = "Total number of items";
            // 
            // bindingNavigatorMoveFirstItem
            // 
            this.bindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMoveFirstItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMoveFirstItem.Image")));
            this.bindingNavigatorMoveFirstItem.Name = "bindingNavigatorMoveFirstItem";
            this.bindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMoveFirstItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorMoveFirstItem.Text = "Move first";
            // 
            // bindingNavigatorMovePreviousItem
            // 
            this.bindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMovePreviousItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMovePreviousItem.Image")));
            this.bindingNavigatorMovePreviousItem.Name = "bindingNavigatorMovePreviousItem";
            this.bindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMovePreviousItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorMovePreviousItem.Text = "Move previous";
            // 
            // bindingNavigatorSeparator
            // 
            this.bindingNavigatorSeparator.Name = "bindingNavigatorSeparator";
            this.bindingNavigatorSeparator.Size = new System.Drawing.Size(6, 25);
            // 
            // bindingNavigatorPositionItem
            // 
            this.bindingNavigatorPositionItem.AccessibleName = "Position";
            this.bindingNavigatorPositionItem.AutoSize = false;
            this.bindingNavigatorPositionItem.Name = "bindingNavigatorPositionItem";
            this.bindingNavigatorPositionItem.Size = new System.Drawing.Size(50, 23);
            this.bindingNavigatorPositionItem.Text = "0";
            this.bindingNavigatorPositionItem.ToolTipText = "Current position";
            // 
            // bindingNavigatorSeparator1
            // 
            this.bindingNavigatorSeparator1.Name = "bindingNavigatorSeparator1";
            this.bindingNavigatorSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // bindingNavigatorMoveNextItem
            // 
            this.bindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMoveNextItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMoveNextItem.Image")));
            this.bindingNavigatorMoveNextItem.Name = "bindingNavigatorMoveNextItem";
            this.bindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMoveNextItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorMoveNextItem.Text = "Move next";
            // 
            // bindingNavigatorMoveLastItem
            // 
            this.bindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMoveLastItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMoveLastItem.Image")));
            this.bindingNavigatorMoveLastItem.Name = "bindingNavigatorMoveLastItem";
            this.bindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMoveLastItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorMoveLastItem.Text = "Move last";
            // 
            // bindingNavigatorSeparator2
            // 
            this.bindingNavigatorSeparator2.Name = "bindingNavigatorSeparator2";
            this.bindingNavigatorSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // toolStripLabel1
            // 
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new System.Drawing.Size(36, 22);
            this.toolStripLabel1.Text = "Filter:";
            // 
            // toolStripComboBoxSkFilter
            // 
            this.toolStripComboBoxSkFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.toolStripComboBoxSkFilter.Items.AddRange(new object[] {
            "(alle)",
            "SDA",
            "AudioVideo",
            "MDA",
            "Tele",
            "Data",
            "Teknikere",
            "Kasse",
            "Aftersales",
            "Cross"});
            this.toolStripComboBoxSkFilter.Name = "toolStripComboBoxSkFilter";
            this.toolStripComboBoxSkFilter.Size = new System.Drawing.Size(100, 25);
            this.toolStripComboBoxSkFilter.DropDownClosed += new System.EventHandler(this.toolStripComboBoxSkFilter_DropDownClosed);
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(56, 22);
            this.toolStripButton1.Text = "&Lagre";
            this.toolStripButton1.Click += new System.EventHandler(this.toolStripButton1_Click);
            // 
            // panel4
            // 
            this.panel4.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel4.Location = new System.Drawing.Point(0, 541);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(776, 20);
            this.panel4.TabIndex = 1;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.groupKrav);
            this.panel2.Controls.Add(this.button3);
            this.panel2.Controls.Add(this.button2);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Right;
            this.panel2.Location = new System.Drawing.Point(776, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(330, 561);
            this.panel2.TabIndex = 1;
            // 
            // groupKrav
            // 
            this.groupKrav.Controls.Add(this.checkKravMtdShowTarget);
            this.groupKrav.Controls.Add(this.groupBox21);
            this.groupKrav.Controls.Add(this.checkKravRtgsa);
            this.groupKrav.Controls.Add(this.checkKravStrom);
            this.groupKrav.Controls.Add(this.checkKravMod);
            this.groupKrav.Controls.Add(this.checkKravFinans);
            this.groupKrav.Controls.Add(this.label80);
            this.groupKrav.Controls.Add(this.checkKravMTD);
            this.groupKrav.Controls.Add(this.label84);
            this.groupKrav.Controls.Add(this.label83);
            this.groupKrav.Controls.Add(this.checkOversiktVisKrav);
            this.groupKrav.Location = new System.Drawing.Point(6, 12);
            this.groupKrav.Name = "groupKrav";
            this.groupKrav.Size = new System.Drawing.Size(312, 369);
            this.groupKrav.TabIndex = 17;
            this.groupKrav.TabStop = false;
            this.groupKrav.Text = "Krav innstillinger";
            // 
            // checkKravMtdShowTarget
            // 
            this.checkKravMtdShowTarget.AutoSize = true;
            this.checkKravMtdShowTarget.Location = new System.Drawing.Point(50, 326);
            this.checkKravMtdShowTarget.Name = "checkKravMtdShowTarget";
            this.checkKravMtdShowTarget.Size = new System.Drawing.Size(203, 17);
            this.checkKravMtdShowTarget.TabIndex = 26;
            this.checkKravMtdShowTarget.Text = "Vis også endelig krav i egen kolonne.";
            this.checkKravMtdShowTarget.UseVisualStyleBackColor = true;
            // 
            // groupBox21
            // 
            this.groupBox21.Controls.Add(this.checkKravAntallFinans);
            this.groupBox21.Controls.Add(this.checkKravAntallMod);
            this.groupBox21.Controls.Add(this.checkKravAntallStrom);
            this.groupBox21.Controls.Add(this.checkKravAntallRtgsa);
            this.groupBox21.Location = new System.Drawing.Point(181, 108);
            this.groupBox21.Name = "groupBox21";
            this.groupBox21.Size = new System.Drawing.Size(74, 129);
            this.groupBox21.TabIndex = 14;
            this.groupBox21.TabStop = false;
            this.groupBox21.Text = "Antall solgt";
            // 
            // checkKravAntallFinans
            // 
            this.checkKravAntallFinans.AutoSize = true;
            this.checkKravAntallFinans.Location = new System.Drawing.Point(27, 25);
            this.checkKravAntallFinans.Name = "checkKravAntallFinans";
            this.checkKravAntallFinans.Size = new System.Drawing.Size(15, 14);
            this.checkKravAntallFinans.TabIndex = 18;
            this.checkKravAntallFinans.UseVisualStyleBackColor = true;
            // 
            // checkKravAntallMod
            // 
            this.checkKravAntallMod.AutoSize = true;
            this.checkKravAntallMod.Location = new System.Drawing.Point(27, 50);
            this.checkKravAntallMod.Name = "checkKravAntallMod";
            this.checkKravAntallMod.Size = new System.Drawing.Size(15, 14);
            this.checkKravAntallMod.TabIndex = 20;
            this.checkKravAntallMod.UseVisualStyleBackColor = true;
            // 
            // checkKravAntallStrom
            // 
            this.checkKravAntallStrom.AutoSize = true;
            this.checkKravAntallStrom.Location = new System.Drawing.Point(27, 76);
            this.checkKravAntallStrom.Name = "checkKravAntallStrom";
            this.checkKravAntallStrom.Size = new System.Drawing.Size(15, 14);
            this.checkKravAntallStrom.TabIndex = 22;
            this.checkKravAntallStrom.UseVisualStyleBackColor = true;
            // 
            // checkKravAntallRtgsa
            // 
            this.checkKravAntallRtgsa.AutoSize = true;
            this.checkKravAntallRtgsa.Location = new System.Drawing.Point(27, 102);
            this.checkKravAntallRtgsa.Name = "checkKravAntallRtgsa";
            this.checkKravAntallRtgsa.Size = new System.Drawing.Size(15, 14);
            this.checkKravAntallRtgsa.TabIndex = 24;
            this.checkKravAntallRtgsa.UseVisualStyleBackColor = true;
            // 
            // checkKravRtgsa
            // 
            this.checkKravRtgsa.AutoSize = true;
            this.checkKravRtgsa.Location = new System.Drawing.Point(43, 209);
            this.checkKravRtgsa.Name = "checkKravRtgsa";
            this.checkKravRtgsa.Size = new System.Drawing.Size(68, 17);
            this.checkKravRtgsa.TabIndex = 23;
            this.checkKravRtgsa.Text = "RTG/SA";
            this.checkKravRtgsa.UseVisualStyleBackColor = true;
            // 
            // checkKravStrom
            // 
            this.checkKravStrom.AutoSize = true;
            this.checkKravStrom.Location = new System.Drawing.Point(43, 183);
            this.checkKravStrom.Name = "checkKravStrom";
            this.checkKravStrom.Size = new System.Drawing.Size(53, 17);
            this.checkKravStrom.TabIndex = 21;
            this.checkKravStrom.Text = "Strøm";
            this.checkKravStrom.UseVisualStyleBackColor = true;
            // 
            // checkKravMod
            // 
            this.checkKravMod.AutoSize = true;
            this.checkKravMod.Location = new System.Drawing.Point(43, 157);
            this.checkKravMod.Name = "checkKravMod";
            this.checkKravMod.Size = new System.Drawing.Size(102, 17);
            this.checkKravMod.TabIndex = 19;
            this.checkKravMod.Text = "Trygghetsavtale";
            this.checkKravMod.UseVisualStyleBackColor = true;
            // 
            // checkKravFinans
            // 
            this.checkKravFinans.AutoSize = true;
            this.checkKravFinans.Location = new System.Drawing.Point(43, 132);
            this.checkKravFinans.Name = "checkKravFinans";
            this.checkKravFinans.Size = new System.Drawing.Size(57, 17);
            this.checkKravFinans.TabIndex = 17;
            this.checkKravFinans.Text = "Finans";
            this.checkKravFinans.UseVisualStyleBackColor = true;
            // 
            // label80
            // 
            this.label80.AutoSize = true;
            this.label80.Location = new System.Drawing.Point(23, 108);
            this.label80.Name = "label80";
            this.label80.Size = new System.Drawing.Size(66, 13);
            this.label80.TabIndex = 8;
            this.label80.Text = "Vis krav for..";
            // 
            // checkKravMTD
            // 
            this.checkKravMTD.AutoSize = true;
            this.checkKravMTD.Location = new System.Drawing.Point(26, 290);
            this.checkKravMTD.Name = "checkKravMTD";
            this.checkKravMTD.Size = new System.Drawing.Size(229, 30);
            this.checkKravMTD.TabIndex = 25;
            this.checkKravMTD.Text = "Beregn krav i henhold til hvor mange dager\r\nsom er gått av måneden (MTD)";
            this.checkKravMTD.UseVisualStyleBackColor = true;
            // 
            // label84
            // 
            this.label84.AutoSize = true;
            this.label84.Location = new System.Drawing.Point(23, 250);
            this.label84.Name = "label84";
            this.label84.Size = new System.Drawing.Size(250, 26);
            this.label84.TabIndex = 6;
            this.label84.Text = "Individuelle krav på inntjening/omsetning eller antall\r\nsettes i selgerkode liste" +
    "n, i hovedvindu.";
            // 
            // label83
            // 
            this.label83.AutoSize = true;
            this.label83.Location = new System.Drawing.Point(23, 53);
            this.label83.Name = "label83";
            this.label83.Size = new System.Drawing.Size(233, 39);
            this.label83.TabIndex = 2;
            this.label83.Text = "Sammenlign inntjening eller omsetning mot krav,\r\nmed unntak av følgende tjenester" +
    " som har antall\r\nsom krav:";
            // 
            // checkOversiktVisKrav
            // 
            this.checkOversiktVisKrav.AutoSize = true;
            this.checkOversiktVisKrav.Location = new System.Drawing.Point(26, 27);
            this.checkOversiktVisKrav.Name = "checkOversiktVisKrav";
            this.checkOversiktVisKrav.Size = new System.Drawing.Size(187, 17);
            this.checkOversiktVisKrav.TabIndex = 16;
            this.checkOversiktVisKrav.Text = "Vis krav kolonner i selger oversikt.";
            this.checkOversiktVisKrav.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button3.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button3.Location = new System.Drawing.Point(140, 526);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 12;
            this.button3.Text = "Avbryt";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button2.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button2.Location = new System.Drawing.Point(243, 526);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 13;
            this.button2.Text = "OK";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // FormKrav
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1106, 561);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormKrav";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "KGSA - Selgerkrav";
            this.panel1.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewSk)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingNavigatorSk)).EndInit();
            this.bindingNavigatorSk.ResumeLayout(false);
            this.bindingNavigatorSk.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.groupKrav.ResumeLayout(false);
            this.groupKrav.PerformLayout();
            this.groupBox21.ResumeLayout(false);
            this.groupBox21.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSourceSk)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.DataGridView dataGridViewSk;
        private System.Windows.Forms.BindingNavigator bindingNavigatorSk;
        private System.Windows.Forms.ToolStripLabel bindingNavigatorCountItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveFirstItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMovePreviousItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator;
        private System.Windows.Forms.ToolStripTextBox bindingNavigatorPositionItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator1;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveNextItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveLastItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator2;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.BindingSource bindingSourceSk;
        private System.Windows.Forms.ToolStripComboBox toolStripComboBoxSkFilter;
        private System.Windows.Forms.ToolStripLabel toolStripLabel1;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.GroupBox groupKrav;
        private System.Windows.Forms.CheckBox checkKravMtdShowTarget;
        private System.Windows.Forms.GroupBox groupBox21;
        private System.Windows.Forms.CheckBox checkKravAntallFinans;
        private System.Windows.Forms.CheckBox checkKravAntallMod;
        private System.Windows.Forms.CheckBox checkKravAntallStrom;
        private System.Windows.Forms.CheckBox checkKravAntallRtgsa;
        private System.Windows.Forms.CheckBox checkKravRtgsa;
        private System.Windows.Forms.CheckBox checkKravStrom;
        private System.Windows.Forms.CheckBox checkKravMod;
        private System.Windows.Forms.CheckBox checkKravFinans;
        private System.Windows.Forms.Label label80;
        private System.Windows.Forms.CheckBox checkKravMTD;
        private System.Windows.Forms.Label label84;
        private System.Windows.Forms.Label label83;
        private System.Windows.Forms.CheckBox checkOversiktVisKrav;
        private System.Windows.Forms.DataGridViewTextBoxColumn Id;
        private System.Windows.Forms.DataGridViewTextBoxColumn Avd;
        private System.Windows.Forms.DataGridViewTextBoxColumn Provisjon;
        private System.Windows.Forms.DataGridViewTextBoxColumn Selgerkode;
        private System.Windows.Forms.DataGridViewTextBoxColumn Navn;
        private System.Windows.Forms.DataGridViewTextBoxColumn Avdeling;
        private System.Windows.Forms.DataGridViewTextBoxColumn Finans;
        private System.Windows.Forms.DataGridViewTextBoxColumn Strom;
        private System.Windows.Forms.DataGridViewTextBoxColumn TA;
        private System.Windows.Forms.DataGridViewTextBoxColumn Rtgsa;
    }
}