namespace KGSA
{
    partial class FormBudgetDetails
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormBudgetDetails));
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Budget_id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Selgerkode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Timer = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Dager = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Multiplikator = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.bindingNavigator1 = new System.Windows.Forms.BindingNavigator(this.components);
            this.bindingNavigatorCountItem = new System.Windows.Forms.ToolStripLabel();
            this.bindingNavigatorDeleteItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMoveFirstItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMovePreviousItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorPositionItem = new System.Windows.Forms.ToolStripTextBox();
            this.bindingNavigatorSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorMoveNextItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMoveLastItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.panel2 = new System.Windows.Forms.Panel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.comboBox_Vinn = new System.Windows.Forms.ComboBox();
            this.textBox_vinn = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox_dager = new System.Windows.Forms.TextBox();
            this.dateTime_date = new System.Windows.Forms.DateTimePicker();
            this.buttonLagre = new System.Windows.Forms.Button();
            this.buttonEdit = new System.Windows.Forms.Button();
            this.comboBox_Acc = new System.Windows.Forms.ComboBox();
            this.comboBox_Finans = new System.Windows.Forms.ComboBox();
            this.comboBox_Rtgsa = new System.Windows.Forms.ComboBox();
            this.comboBox_Strom = new System.Windows.Forms.ComboBox();
            this.comboBox_TA = new System.Windows.Forms.ComboBox();
            this.textBox_acc = new System.Windows.Forms.TextBox();
            this.textBox_finans = new System.Windows.Forms.TextBox();
            this.textBox_rtgsa = new System.Windows.Forms.TextBox();
            this.textBox_strom = new System.Windows.Forms.TextBox();
            this.textBox_ta = new System.Windows.Forms.TextBox();
            this.textBox_date = new System.Windows.Forms.TextBox();
            this.textBox_margin = new System.Windows.Forms.TextBox();
            this.textBox_inntjening = new System.Windows.Forms.TextBox();
            this.textBox_omsetning = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.panel4 = new System.Windows.Forms.Panel();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.panel1.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingNavigator1)).BeginInit();
            this.bindingNavigator1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.panel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Controls.Add(this.panel4);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(615, 475);
            this.panel1.TabIndex = 0;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.dataGridView1);
            this.panel3.Controls.Add(this.bindingNavigator1);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(0, 195);
            this.panel3.Name = "panel3";
            this.panel3.Padding = new System.Windows.Forms.Padding(5);
            this.panel3.Size = new System.Drawing.Size(615, 231);
            this.panel3.TabIndex = 1;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToResizeRows = false;
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.Control;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Id,
            this.Budget_id,
            this.Selgerkode,
            this.Timer,
            this.Dager,
            this.Multiplikator});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(5, 30);
            this.dataGridView1.MultiSelect = false;
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(605, 196);
            this.dataGridView1.TabIndex = 0;
            // 
            // Id
            // 
            this.Id.DataPropertyName = "Id";
            this.Id.HeaderText = "Id";
            this.Id.Name = "Id";
            this.Id.Visible = false;
            // 
            // Budget_id
            // 
            this.Budget_id.DataPropertyName = "BudgetId";
            this.Budget_id.HeaderText = "BudgetId";
            this.Budget_id.Name = "Budget_id";
            this.Budget_id.Visible = false;
            // 
            // Selgerkode
            // 
            this.Selgerkode.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Selgerkode.DataPropertyName = "Selgerkode";
            this.Selgerkode.FillWeight = 110F;
            this.Selgerkode.HeaderText = "Selgerkode";
            this.Selgerkode.Name = "Selgerkode";
            // 
            // Timer
            // 
            this.Timer.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Timer.DataPropertyName = "Timer";
            this.Timer.HeaderText = "Timer";
            this.Timer.Name = "Timer";
            // 
            // Dager
            // 
            this.Dager.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Dager.DataPropertyName = "Dager";
            this.Dager.HeaderText = "Dager";
            this.Dager.Name = "Dager";
            // 
            // Multiplikator
            // 
            this.Multiplikator.DataPropertyName = "Multiplikator";
            this.Multiplikator.HeaderText = "Multiplikator";
            this.Multiplikator.Name = "Multiplikator";
            // 
            // bindingNavigator1
            // 
            this.bindingNavigator1.AddNewItem = null;
            this.bindingNavigator1.CountItem = this.bindingNavigatorCountItem;
            this.bindingNavigator1.DeleteItem = this.bindingNavigatorDeleteItem;
            this.bindingNavigator1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.bindingNavigatorMoveFirstItem,
            this.bindingNavigatorMovePreviousItem,
            this.bindingNavigatorSeparator,
            this.bindingNavigatorPositionItem,
            this.bindingNavigatorCountItem,
            this.bindingNavigatorSeparator1,
            this.bindingNavigatorMoveNextItem,
            this.bindingNavigatorMoveLastItem,
            this.bindingNavigatorSeparator2,
            this.bindingNavigatorDeleteItem,
            this.toolStripSeparator1,
            this.toolStripButton1});
            this.bindingNavigator1.Location = new System.Drawing.Point(5, 5);
            this.bindingNavigator1.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.bindingNavigator1.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.bindingNavigator1.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.bindingNavigator1.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.bindingNavigator1.Name = "bindingNavigator1";
            this.bindingNavigator1.PositionItem = this.bindingNavigatorPositionItem;
            this.bindingNavigator1.Size = new System.Drawing.Size(605, 25);
            this.bindingNavigator1.TabIndex = 1;
            this.bindingNavigator1.Text = "bindingNavigator1";
            // 
            // bindingNavigatorCountItem
            // 
            this.bindingNavigatorCountItem.Name = "bindingNavigatorCountItem";
            this.bindingNavigatorCountItem.Size = new System.Drawing.Size(35, 22);
            this.bindingNavigatorCountItem.Text = "of {0}";
            this.bindingNavigatorCountItem.ToolTipText = "Total number of items";
            // 
            // bindingNavigatorDeleteItem
            // 
            this.bindingNavigatorDeleteItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorDeleteItem.Image")));
            this.bindingNavigatorDeleteItem.Name = "bindingNavigatorDeleteItem";
            this.bindingNavigatorDeleteItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorDeleteItem.Size = new System.Drawing.Size(50, 22);
            this.bindingNavigatorDeleteItem.Text = "Slett";
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
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(70, 22);
            this.toolStripButton1.Text = "Time tabell";
            this.toolStripButton1.Click += new System.EventHandler(this.toolStripButton1_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.groupBox2);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(615, 195);
            this.panel2.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.comboBox_Vinn);
            this.groupBox2.Controls.Add(this.textBox_vinn);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.textBox_dager);
            this.groupBox2.Controls.Add(this.dateTime_date);
            this.groupBox2.Controls.Add(this.buttonLagre);
            this.groupBox2.Controls.Add(this.buttonEdit);
            this.groupBox2.Controls.Add(this.comboBox_Acc);
            this.groupBox2.Controls.Add(this.comboBox_Finans);
            this.groupBox2.Controls.Add(this.comboBox_Rtgsa);
            this.groupBox2.Controls.Add(this.comboBox_Strom);
            this.groupBox2.Controls.Add(this.comboBox_TA);
            this.groupBox2.Controls.Add(this.textBox_acc);
            this.groupBox2.Controls.Add(this.textBox_finans);
            this.groupBox2.Controls.Add(this.textBox_rtgsa);
            this.groupBox2.Controls.Add(this.textBox_strom);
            this.groupBox2.Controls.Add(this.textBox_ta);
            this.groupBox2.Controls.Add(this.textBox_date);
            this.groupBox2.Controls.Add(this.textBox_margin);
            this.groupBox2.Controls.Add(this.textBox_inntjening);
            this.groupBox2.Controls.Add(this.textBox_omsetning);
            this.groupBox2.Controls.Add(this.label14);
            this.groupBox2.Controls.Add(this.label13);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(0, 0);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(615, 195);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Detaljer for budsjett:";
            // 
            // comboBox_Vinn
            // 
            this.comboBox_Vinn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Vinn.Enabled = false;
            this.comboBox_Vinn.FormattingEnabled = true;
            this.comboBox_Vinn.Items.AddRange(new object[] {
            "Poeng",
            "Antall"});
            this.comboBox_Vinn.Location = new System.Drawing.Point(402, 155);
            this.comboBox_Vinn.Name = "comboBox_Vinn";
            this.comboBox_Vinn.Size = new System.Drawing.Size(99, 21);
            this.comboBox_Vinn.TabIndex = 42;
            // 
            // textBox_vinn
            // 
            this.textBox_vinn.Location = new System.Drawing.Point(296, 155);
            this.textBox_vinn.Name = "textBox_vinn";
            this.textBox_vinn.ReadOnly = true;
            this.textBox_vinn.Size = new System.Drawing.Size(100, 20);
            this.textBox_vinn.TabIndex = 41;
            this.textBox_vinn.TabStop = false;
            this.textBox_vinn.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(223, 158);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 13);
            this.label2.TabIndex = 40;
            this.label2.Text = "Vinnprodukt:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(32, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 13);
            this.label1.TabIndex = 39;
            this.label1.Text = "Åpningsdager:";
            // 
            // textBox_dager
            // 
            this.textBox_dager.Location = new System.Drawing.Point(113, 51);
            this.textBox_dager.Name = "textBox_dager";
            this.textBox_dager.ReadOnly = true;
            this.textBox_dager.Size = new System.Drawing.Size(58, 20);
            this.textBox_dager.TabIndex = 38;
            this.textBox_dager.TabStop = false;
            this.textBox_dager.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // dateTime_date
            // 
            this.dateTime_date.Location = new System.Drawing.Point(113, 25);
            this.dateTime_date.Name = "dateTime_date";
            this.dateTime_date.Size = new System.Drawing.Size(134, 20);
            this.dateTime_date.TabIndex = 37;
            this.dateTime_date.Visible = false;
            // 
            // buttonLagre
            // 
            this.buttonLagre.Location = new System.Drawing.Point(518, 96);
            this.buttonLagre.Name = "buttonLagre";
            this.buttonLagre.Size = new System.Drawing.Size(75, 23);
            this.buttonLagre.TabIndex = 36;
            this.buttonLagre.Text = "Lagre";
            this.buttonLagre.UseVisualStyleBackColor = true;
            this.buttonLagre.Click += new System.EventHandler(this.buttonLagre_Click);
            // 
            // buttonEdit
            // 
            this.buttonEdit.Location = new System.Drawing.Point(518, 54);
            this.buttonEdit.Name = "buttonEdit";
            this.buttonEdit.Size = new System.Drawing.Size(75, 23);
            this.buttonEdit.TabIndex = 35;
            this.buttonEdit.Text = "Rediger";
            this.buttonEdit.UseVisualStyleBackColor = true;
            this.buttonEdit.Click += new System.EventHandler(this.button1_Click);
            // 
            // comboBox_Acc
            // 
            this.comboBox_Acc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Acc.Enabled = false;
            this.comboBox_Acc.FormattingEnabled = true;
            this.comboBox_Acc.Items.AddRange(new object[] {
            "SoM",
            "SoB",
            "Inntjen",
            "Omset",
            "Antall"});
            this.comboBox_Acc.Location = new System.Drawing.Point(402, 129);
            this.comboBox_Acc.Name = "comboBox_Acc";
            this.comboBox_Acc.Size = new System.Drawing.Size(99, 21);
            this.comboBox_Acc.TabIndex = 34;
            // 
            // comboBox_Finans
            // 
            this.comboBox_Finans.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Finans.Enabled = false;
            this.comboBox_Finans.FormattingEnabled = true;
            this.comboBox_Finans.Items.AddRange(new object[] {
            "SoM",
            "SoB",
            "Inntjen",
            "Omset",
            "Antall"});
            this.comboBox_Finans.Location = new System.Drawing.Point(402, 103);
            this.comboBox_Finans.Name = "comboBox_Finans";
            this.comboBox_Finans.Size = new System.Drawing.Size(99, 21);
            this.comboBox_Finans.TabIndex = 33;
            // 
            // comboBox_Rtgsa
            // 
            this.comboBox_Rtgsa.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Rtgsa.Enabled = false;
            this.comboBox_Rtgsa.FormattingEnabled = true;
            this.comboBox_Rtgsa.Items.AddRange(new object[] {
            "Hitrate",
            "SoM",
            "SoB",
            "Inntjen",
            "Omset",
            "Antall"});
            this.comboBox_Rtgsa.Location = new System.Drawing.Point(402, 77);
            this.comboBox_Rtgsa.Name = "comboBox_Rtgsa";
            this.comboBox_Rtgsa.Size = new System.Drawing.Size(99, 21);
            this.comboBox_Rtgsa.TabIndex = 32;
            // 
            // comboBox_Strom
            // 
            this.comboBox_Strom.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Strom.Enabled = false;
            this.comboBox_Strom.FormattingEnabled = true;
            this.comboBox_Strom.Items.AddRange(new object[] {
            "Antall",
            "SoM",
            "SoB",
            "Inntjen"});
            this.comboBox_Strom.Location = new System.Drawing.Point(402, 51);
            this.comboBox_Strom.Name = "comboBox_Strom";
            this.comboBox_Strom.Size = new System.Drawing.Size(99, 21);
            this.comboBox_Strom.TabIndex = 31;
            // 
            // comboBox_TA
            // 
            this.comboBox_TA.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_TA.Enabled = false;
            this.comboBox_TA.FormattingEnabled = true;
            this.comboBox_TA.Items.AddRange(new object[] {
            "SoB",
            "SoM",
            "Inntjen",
            "Omset",
            "Antall"});
            this.comboBox_TA.Location = new System.Drawing.Point(402, 25);
            this.comboBox_TA.Name = "comboBox_TA";
            this.comboBox_TA.Size = new System.Drawing.Size(99, 21);
            this.comboBox_TA.TabIndex = 30;
            // 
            // textBox_acc
            // 
            this.textBox_acc.Location = new System.Drawing.Point(296, 129);
            this.textBox_acc.Name = "textBox_acc";
            this.textBox_acc.ReadOnly = true;
            this.textBox_acc.Size = new System.Drawing.Size(100, 20);
            this.textBox_acc.TabIndex = 29;
            this.textBox_acc.TabStop = false;
            this.textBox_acc.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox_finans
            // 
            this.textBox_finans.Location = new System.Drawing.Point(296, 103);
            this.textBox_finans.Name = "textBox_finans";
            this.textBox_finans.ReadOnly = true;
            this.textBox_finans.Size = new System.Drawing.Size(100, 20);
            this.textBox_finans.TabIndex = 28;
            this.textBox_finans.TabStop = false;
            this.textBox_finans.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox_rtgsa
            // 
            this.textBox_rtgsa.Location = new System.Drawing.Point(296, 77);
            this.textBox_rtgsa.Name = "textBox_rtgsa";
            this.textBox_rtgsa.ReadOnly = true;
            this.textBox_rtgsa.Size = new System.Drawing.Size(100, 20);
            this.textBox_rtgsa.TabIndex = 27;
            this.textBox_rtgsa.TabStop = false;
            this.textBox_rtgsa.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox_strom
            // 
            this.textBox_strom.Location = new System.Drawing.Point(296, 51);
            this.textBox_strom.Name = "textBox_strom";
            this.textBox_strom.ReadOnly = true;
            this.textBox_strom.Size = new System.Drawing.Size(100, 20);
            this.textBox_strom.TabIndex = 26;
            this.textBox_strom.TabStop = false;
            this.textBox_strom.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox_ta
            // 
            this.textBox_ta.Location = new System.Drawing.Point(296, 25);
            this.textBox_ta.Name = "textBox_ta";
            this.textBox_ta.ReadOnly = true;
            this.textBox_ta.Size = new System.Drawing.Size(100, 20);
            this.textBox_ta.TabIndex = 25;
            this.textBox_ta.TabStop = false;
            this.textBox_ta.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox_date
            // 
            this.textBox_date.Location = new System.Drawing.Point(113, 25);
            this.textBox_date.Name = "textBox_date";
            this.textBox_date.ReadOnly = true;
            this.textBox_date.Size = new System.Drawing.Size(100, 20);
            this.textBox_date.TabIndex = 23;
            this.textBox_date.TabStop = false;
            // 
            // textBox_margin
            // 
            this.textBox_margin.Location = new System.Drawing.Point(113, 129);
            this.textBox_margin.Name = "textBox_margin";
            this.textBox_margin.ReadOnly = true;
            this.textBox_margin.Size = new System.Drawing.Size(58, 20);
            this.textBox_margin.TabIndex = 22;
            this.textBox_margin.TabStop = false;
            this.textBox_margin.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox_inntjening
            // 
            this.textBox_inntjening.Location = new System.Drawing.Point(113, 103);
            this.textBox_inntjening.Name = "textBox_inntjening";
            this.textBox_inntjening.ReadOnly = true;
            this.textBox_inntjening.Size = new System.Drawing.Size(100, 20);
            this.textBox_inntjening.TabIndex = 21;
            this.textBox_inntjening.TabStop = false;
            this.textBox_inntjening.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox_omsetning
            // 
            this.textBox_omsetning.Location = new System.Drawing.Point(113, 77);
            this.textBox_omsetning.Name = "textBox_omsetning";
            this.textBox_omsetning.ReadOnly = true;
            this.textBox_omsetning.Size = new System.Drawing.Size(100, 20);
            this.textBox_omsetning.TabIndex = 20;
            this.textBox_omsetning.TabStop = false;
            this.textBox_omsetning.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(242, 132);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(48, 13);
            this.label14.TabIndex = 13;
            this.label14.Text = "Tilbehør:";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(249, 106);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(41, 13);
            this.label13.TabIndex = 12;
            this.label13.Text = "Finans:";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(238, 80);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(52, 13);
            this.label12.TabIndex = 11;
            this.label12.Text = "RTG/SA:";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(253, 54);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(37, 13);
            this.label11.TabIndex = 10;
            this.label11.Text = "Strøm:";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(266, 28);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(24, 13);
            this.label10.TabIndex = 9;
            this.label10.Text = "TA:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(65, 132);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(42, 13);
            this.label6.TabIndex = 5;
            this.label6.Text = "Margin:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(51, 106);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(56, 13);
            this.label5.TabIndex = 4;
            this.label5.Text = "Inntjening:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(47, 80);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(60, 13);
            this.label4.TabIndex = 3;
            this.label4.Text = "Omsetning:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(24, 28);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(83, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Budsjett måned:";
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.button3);
            this.panel4.Controls.Add(this.button2);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel4.Location = new System.Drawing.Point(0, 426);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(615, 49);
            this.panel4.TabIndex = 2;
            // 
            // button3
            // 
            this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button3.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button3.Location = new System.Drawing.Point(413, 14);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 1;
            this.button3.Text = "OK";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button2.Location = new System.Drawing.Point(528, 14);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 0;
            this.button2.Text = "Avbryt";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // FormBudgetDetails
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(615, 475);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormBudgetDetails";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Budsjett:";
            this.panel1.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingNavigator1)).EndInit();
            this.bindingNavigator1.ResumeLayout(false);
            this.bindingNavigator1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.panel4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.BindingNavigator bindingNavigator1;
        private System.Windows.Forms.ToolStripLabel bindingNavigatorCountItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorDeleteItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveFirstItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMovePreviousItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator;
        private System.Windows.Forms.ToolStripTextBox bindingNavigatorPositionItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator1;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveNextItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveLastItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.BindingSource bindingSource1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Id;
        private System.Windows.Forms.DataGridViewTextBoxColumn Budget_id;
        private System.Windows.Forms.DataGridViewTextBoxColumn Selgerkode;
        private System.Windows.Forms.DataGridViewTextBoxColumn Timer;
        private System.Windows.Forms.DataGridViewTextBoxColumn Dager;
        private System.Windows.Forms.DataGridViewTextBoxColumn Multiplikator;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private System.Windows.Forms.TextBox textBox_omsetning;
        private System.Windows.Forms.TextBox textBox_acc;
        private System.Windows.Forms.TextBox textBox_finans;
        private System.Windows.Forms.TextBox textBox_rtgsa;
        private System.Windows.Forms.TextBox textBox_strom;
        private System.Windows.Forms.TextBox textBox_ta;
        private System.Windows.Forms.TextBox textBox_date;
        private System.Windows.Forms.TextBox textBox_margin;
        private System.Windows.Forms.TextBox textBox_inntjening;
        private System.Windows.Forms.Button buttonLagre;
        private System.Windows.Forms.Button buttonEdit;
        private System.Windows.Forms.ComboBox comboBox_Acc;
        private System.Windows.Forms.ComboBox comboBox_Finans;
        private System.Windows.Forms.ComboBox comboBox_Rtgsa;
        private System.Windows.Forms.ComboBox comboBox_Strom;
        private System.Windows.Forms.ComboBox comboBox_TA;
        private System.Windows.Forms.DateTimePicker dateTime_date;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox_dager;
        private System.Windows.Forms.ComboBox comboBox_Vinn;
        private System.Windows.Forms.TextBox textBox_vinn;
        private System.Windows.Forms.Label label2;
    }
}