namespace window3
{
    partial class teleForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Title title1 = new System.Windows.Forms.DataVisualization.Charting.Title();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(teleForm));
            this.id = new System.Windows.Forms.MenuStrip();
            this.typeCombo = new System.Windows.Forms.ToolStripComboBox();
            this.id_dev = new System.Windows.Forms.ToolStripMenuItem();
            this.dev_checkAll = new System.Windows.Forms.ToolStripMenuItem();
            this.par = new System.Windows.Forms.ToolStripMenuItem();
            this.par_checkAll = new System.Windows.Forms.ToolStripMenuItem();
            this.work_time = new System.Windows.Forms.ToolStripMenuItem();
            this.power = new System.Windows.Forms.ToolStripMenuItem();
            this.loading = new System.Windows.Forms.ToolStripMenuItem();
            this.vibr = new System.Windows.Forms.ToolStripMenuItem();
            this.temp = new System.Windows.Forms.ToolStripMenuItem();
            this.dPicker1 = new System.Windows.Forms.DateTimePicker();
            this.dPicker2 = new System.Windows.Forms.DateTimePicker();
            this.id_devMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.fdToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.showButton = new System.Windows.Forms.Button();
            this.insertButton = new System.Windows.Forms.Button();
            this.infoBox = new System.Windows.Forms.Panel();
            this.chart1 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.id.SuspendLayout();
            this.id_devMenu.SuspendLayout();
            this.infoBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).BeginInit();
            this.SuspendLayout();
            // 
            // id
            // 
            this.id.BackColor = System.Drawing.Color.Black;
            this.id.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.id.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.typeCombo,
            this.id_dev,
            this.par});
            this.id.Location = new System.Drawing.Point(0, 0);
            this.id.Name = "id";
            this.id.Size = new System.Drawing.Size(891, 32);
            this.id.TabIndex = 3;
            this.id.Text = "menuStrip1";
            // 
            // typeCombo
            // 
            this.typeCombo.BackColor = System.Drawing.Color.Black;
            this.typeCombo.Font = new System.Drawing.Font("Century Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.typeCombo.ForeColor = System.Drawing.Color.White;
            this.typeCombo.Items.AddRange(new object[] {
            "График",
            "Диаграмма",
            "Таблица"});
            this.typeCombo.Name = "typeCombo";
            this.typeCombo.Overflow = System.Windows.Forms.ToolStripItemOverflow.Never;
            this.typeCombo.Size = new System.Drawing.Size(170, 28);
            this.typeCombo.Text = "Вид отображения:";
            this.typeCombo.SelectedIndexChanged += new System.EventHandler(this.всеПоля_CheckedChanged);
            // 
            // id_dev
            // 
            this.id_dev.BackColor = System.Drawing.Color.Black;
            this.id_dev.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.dev_checkAll});
            this.id_dev.Font = new System.Drawing.Font("Century Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.id_dev.ForeColor = System.Drawing.Color.White;
            this.id_dev.ImageTransparentColor = System.Drawing.Color.White;
            this.id_dev.Name = "id_dev";
            this.id_dev.Size = new System.Drawing.Size(138, 28);
            this.id_dev.Text = "Оборудование";
            // 
            // dev_checkAll
            // 
            this.dev_checkAll.CheckOnClick = true;
            this.dev_checkAll.ForeColor = System.Drawing.Color.White;
            this.dev_checkAll.Name = "dev_checkAll";
            this.dev_checkAll.Size = new System.Drawing.Size(224, 26);
            this.dev_checkAll.Text = "Отметить все";
            this.dev_checkAll.Click += new System.EventHandler(this.dev_checkAll_Click);
            // 
            // par
            // 
            this.par.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.par_checkAll,
            this.work_time,
            this.power,
            this.loading,
            this.vibr,
            this.temp});
            this.par.Font = new System.Drawing.Font("Century Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.par.ForeColor = System.Drawing.Color.White;
            this.par.Name = "par";
            this.par.Size = new System.Drawing.Size(102, 28);
            this.par.Text = "Параметр";
            // 
            // par_checkAll
            // 
            this.par_checkAll.CheckOnClick = true;
            this.par_checkAll.ForeColor = System.Drawing.Color.White;
            this.par_checkAll.Name = "par_checkAll";
            this.par_checkAll.Size = new System.Drawing.Size(354, 26);
            this.par_checkAll.Text = "Отметить все";
            this.par_checkAll.Click += new System.EventHandler(this.par_checkAll_Click);
            // 
            // work_time
            // 
            this.work_time.CheckOnClick = true;
            this.work_time.ForeColor = System.Drawing.Color.White;
            this.work_time.Name = "work_time";
            this.work_time.Size = new System.Drawing.Size(354, 26);
            this.work_time.Text = "Время работы, часы";
            this.work_time.CheckedChanged += new System.EventHandler(this.всеПоля_CheckedChanged);
            // 
            // power
            // 
            this.power.CheckOnClick = true;
            this.power.ForeColor = System.Drawing.Color.White;
            this.power.Name = "power";
            this.power.Size = new System.Drawing.Size(354, 26);
            this.power.Text = "Потребляемая мощность, кВт";
            this.power.CheckedChanged += new System.EventHandler(this.всеПоля_CheckedChanged);
            // 
            // loading
            // 
            this.loading.CheckOnClick = true;
            this.loading.ForeColor = System.Drawing.Color.White;
            this.loading.Name = "loading";
            this.loading.Size = new System.Drawing.Size(354, 26);
            this.loading.Text = "Уровень нагрузки, %";
            this.loading.CheckedChanged += new System.EventHandler(this.всеПоля_CheckedChanged);
            // 
            // vibr
            // 
            this.vibr.CheckOnClick = true;
            this.vibr.ForeColor = System.Drawing.Color.White;
            this.vibr.Name = "vibr";
            this.vibr.Size = new System.Drawing.Size(354, 26);
            this.vibr.Text = "Уровень вибрации, жесткие диски";
            this.vibr.CheckedChanged += new System.EventHandler(this.всеПоля_CheckedChanged);
            // 
            // temp
            // 
            this.temp.CheckOnClick = true;
            this.temp.ForeColor = System.Drawing.Color.White;
            this.temp.Name = "temp";
            this.temp.Size = new System.Drawing.Size(354, 26);
            this.temp.Text = "Температура, C";
            this.temp.CheckedChanged += new System.EventHandler(this.всеПоля_CheckedChanged);
            // 
            // dPicker1
            // 
            this.dPicker1.CustomFormat = "d.MM.yyyy HH:mm";
            this.dPicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dPicker1.Location = new System.Drawing.Point(470, 6);
            this.dPicker1.Name = "dPicker1";
            this.dPicker1.Size = new System.Drawing.Size(170, 25);
            this.dPicker1.TabIndex = 4;
            this.dPicker1.TextChanged += new System.EventHandler(this.DPicker_ValueChanged);
            // 
            // dPicker2
            // 
            this.dPicker2.CustomFormat = "d.MM.yyyy HH:mm";
            this.dPicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dPicker2.Location = new System.Drawing.Point(693, 6);
            this.dPicker2.Name = "dPicker2";
            this.dPicker2.Size = new System.Drawing.Size(170, 25);
            this.dPicker2.TabIndex = 5;
            this.dPicker2.TextChanged += new System.EventHandler(this.DPicker_ValueChanged);
            // 
            // id_devMenu
            // 
            this.id_devMenu.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.id_devMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fdToolStripMenuItem,
            this.toolStripMenuItem1});
            this.id_devMenu.Name = "id_devMenu";
            this.id_devMenu.Size = new System.Drawing.Size(94, 52);
            // 
            // fdToolStripMenuItem
            // 
            this.fdToolStripMenuItem.Name = "fdToolStripMenuItem";
            this.fdToolStripMenuItem.Size = new System.Drawing.Size(93, 24);
            this.fdToolStripMenuItem.Text = "fd";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(93, 24);
            // 
            // showButton
            // 
            this.showButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.showButton.Font = new System.Drawing.Font("Century Gothic", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.showButton.ForeColor = System.Drawing.Color.White;
            this.showButton.Location = new System.Drawing.Point(68, 516);
            this.showButton.Name = "showButton";
            this.showButton.Size = new System.Drawing.Size(248, 39);
            this.showButton.TabIndex = 10;
            this.showButton.Text = "Показать отчёт";
            this.showButton.UseVisualStyleBackColor = true;
            this.showButton.Click += new System.EventHandler(this.showButton_Click);
            // 
            // insertButton
            // 
            this.insertButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.insertButton.Font = new System.Drawing.Font("Century Gothic", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.insertButton.ForeColor = System.Drawing.Color.White;
            this.insertButton.Location = new System.Drawing.Point(543, 516);
            this.insertButton.Name = "insertButton";
            this.insertButton.Size = new System.Drawing.Size(244, 39);
            this.insertButton.TabIndex = 11;
            this.insertButton.Text = "Добавить запись";
            this.insertButton.UseVisualStyleBackColor = true;
            // 
            // infoBox
            // 
            this.infoBox.AutoScroll = true;
            this.infoBox.Controls.Add(this.chart1);
            this.infoBox.Location = new System.Drawing.Point(68, 68);
            this.infoBox.Name = "infoBox";
            this.infoBox.Size = new System.Drawing.Size(719, 432);
            this.infoBox.TabIndex = 12;
            // 
            // chart1
            // 
            chartArea1.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea1);
            legend1.Name = "Legend1";
            this.chart1.Legends.Add(legend1);
            this.chart1.Location = new System.Drawing.Point(136, 67);
            this.chart1.Name = "chart1";
            series1.ChartArea = "ChartArea1";
            series1.Legend = "Legend1";
            series1.Name = "Series1";
            this.chart1.Series.Add(series1);
            this.chart1.Size = new System.Drawing.Size(300, 319);
            this.chart1.TabIndex = 0;
            this.chart1.Text = "chart1";
            title1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            title1.Name = "Title1";
            this.chart1.Titles.Add(title1);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Black;
            this.label1.Font = new System.Drawing.Font("Century Gothic", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(434, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(30, 19);
            this.label1.TabIndex = 12;
            this.label1.Text = "От:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Black;
            this.label2.Font = new System.Drawing.Font("Century Gothic", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(655, 6);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(32, 19);
            this.label2.TabIndex = 13;
            this.label2.Text = "До:";
            // 
            // teleForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(27)))), ((int)(((byte)(27)))));
            this.ClientSize = new System.Drawing.Size(891, 589);
            this.Controls.Add(this.insertButton);
            this.Controls.Add(this.infoBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.showButton);
            this.Controls.Add(this.dPicker2);
            this.Controls.Add(this.dPicker1);
            this.Controls.Add(this.id);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.id;
            this.MaximizeBox = false;
            this.Name = "teleForm";
            this.Text = "Телеметрия";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.teleForm_FormClosed);
            this.Load += new System.EventHandler(this.teleForm_Load);
            this.id.ResumeLayout(false);
            this.id.PerformLayout();
            this.id_devMenu.ResumeLayout(false);
            this.infoBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.MenuStrip id;
        private System.Windows.Forms.ToolStripComboBox typeCombo;
        private System.Windows.Forms.DateTimePicker dPicker1;
        private System.Windows.Forms.DateTimePicker dPicker2;
        private System.Windows.Forms.ContextMenuStrip id_devMenu;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem fdToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem id_dev;
        private System.Windows.Forms.ToolStripMenuItem par;
        private System.Windows.Forms.ToolStripMenuItem work_time;
        private System.Windows.Forms.ToolStripMenuItem power;
        private System.Windows.Forms.ToolStripMenuItem loading;
        private System.Windows.Forms.ToolStripMenuItem temp;
        private System.Windows.Forms.ToolStripMenuItem vibr;
        private System.Windows.Forms.Button showButton;
        private System.Windows.Forms.Button insertButton;
        private System.Windows.Forms.ToolStripMenuItem dev_checkAll;
        private System.Windows.Forms.ToolStripMenuItem par_checkAll;
        private System.Windows.Forms.Panel infoBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart1;
    }
}