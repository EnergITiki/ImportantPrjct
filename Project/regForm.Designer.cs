namespace window3
{
    partial class regForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(regForm));
            this.postComboBox = new System.Windows.Forms.ComboBox();
            this.postLabel = new System.Windows.Forms.Label();
            this.nameLabel = new System.Windows.Forms.Label();
            this.nameText = new System.Windows.Forms.TextBox();
            this.sexLabel = new System.Windows.Forms.Label();
            this.sexComboBox = new System.Windows.Forms.ComboBox();
            this.ageLabel = new System.Windows.Forms.Label();
            this.ageText = new System.Windows.Forms.TextBox();
            this.mailLabel = new System.Windows.Forms.Label();
            this.mailText = new System.Windows.Forms.TextBox();
            this.expLabel = new System.Windows.Forms.Label();
            this.expText = new System.Windows.Forms.TextBox();
            this.regButton2 = new System.Windows.Forms.Button();
            this.passRegText = new System.Windows.Forms.Label();
            this.passText = new System.Windows.Forms.TextBox();
            this.logText = new System.Windows.Forms.TextBox();
            this.logLabel = new System.Windows.Forms.Label();
            this.warnLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // postComboBox
            // 
            this.postComboBox.FormattingEnabled = true;
            this.postComboBox.Items.AddRange(new object[] {
            "Сотрудник",
            "Руководитель отдела",
            "Руководитель предприятия"});
            this.postComboBox.Location = new System.Drawing.Point(211, 49);
            this.postComboBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.postComboBox.Name = "postComboBox";
            this.postComboBox.Size = new System.Drawing.Size(205, 24);
            this.postComboBox.TabIndex = 0;
            // 
            // postLabel
            // 
            this.postLabel.AutoSize = true;
            this.postLabel.Font = new System.Drawing.Font("Century Gothic", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.postLabel.ForeColor = System.Drawing.Color.White;
            this.postLabel.Location = new System.Drawing.Point(91, 48);
            this.postLabel.Name = "postLabel";
            this.postLabel.Size = new System.Drawing.Size(103, 21);
            this.postLabel.TabIndex = 1;
            this.postLabel.Text = "Должность";
            // 
            // nameLabel
            // 
            this.nameLabel.AutoSize = true;
            this.nameLabel.Font = new System.Drawing.Font("Century Gothic", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.nameLabel.ForeColor = System.Drawing.Color.White;
            this.nameLabel.Location = new System.Drawing.Point(91, 144);
            this.nameLabel.Name = "nameLabel";
            this.nameLabel.Size = new System.Drawing.Size(53, 21);
            this.nameLabel.TabIndex = 2;
            this.nameLabel.Text = "ФИО";
            // 
            // nameText
            // 
            this.nameText.Location = new System.Drawing.Point(211, 145);
            this.nameText.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.nameText.Name = "nameText";
            this.nameText.Size = new System.Drawing.Size(245, 22);
            this.nameText.TabIndex = 3;
            this.nameText.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.nameText_KeyPress);
            // 
            // sexLabel
            // 
            this.sexLabel.AutoSize = true;
            this.sexLabel.Font = new System.Drawing.Font("Century Gothic", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.sexLabel.ForeColor = System.Drawing.Color.White;
            this.sexLabel.Location = new System.Drawing.Point(91, 190);
            this.sexLabel.Name = "sexLabel";
            this.sexLabel.Size = new System.Drawing.Size(42, 21);
            this.sexLabel.TabIndex = 4;
            this.sexLabel.Text = "Пол";
            // 
            // sexComboBox
            // 
            this.sexComboBox.FormattingEnabled = true;
            this.sexComboBox.Items.AddRange(new object[] {
            "Мужской",
            "Женский"});
            this.sexComboBox.Location = new System.Drawing.Point(211, 191);
            this.sexComboBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.sexComboBox.Name = "sexComboBox";
            this.sexComboBox.Size = new System.Drawing.Size(121, 24);
            this.sexComboBox.TabIndex = 5;
            // 
            // ageLabel
            // 
            this.ageLabel.AutoSize = true;
            this.ageLabel.Font = new System.Drawing.Font("Century Gothic", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ageLabel.ForeColor = System.Drawing.Color.White;
            this.ageLabel.Location = new System.Drawing.Point(91, 236);
            this.ageLabel.Name = "ageLabel";
            this.ageLabel.Size = new System.Drawing.Size(80, 21);
            this.ageLabel.TabIndex = 6;
            this.ageLabel.Text = "Возраст";
            // 
            // ageText
            // 
            this.ageText.Location = new System.Drawing.Point(211, 238);
            this.ageText.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.ageText.Name = "ageText";
            this.ageText.Size = new System.Drawing.Size(45, 22);
            this.ageText.TabIndex = 7;
            this.ageText.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ageText_KeyPress);
            // 
            // mailLabel
            // 
            this.mailLabel.AutoSize = true;
            this.mailLabel.Font = new System.Drawing.Font("Century Gothic", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.mailLabel.ForeColor = System.Drawing.Color.White;
            this.mailLabel.Location = new System.Drawing.Point(91, 283);
            this.mailLabel.Name = "mailLabel";
            this.mailLabel.Size = new System.Drawing.Size(60, 21);
            this.mailLabel.TabIndex = 8;
            this.mailLabel.Text = "Почта";
            // 
            // mailText
            // 
            this.mailText.Location = new System.Drawing.Point(211, 284);
            this.mailText.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.mailText.Name = "mailText";
            this.mailText.Size = new System.Drawing.Size(245, 22);
            this.mailText.TabIndex = 9;
            // 
            // expLabel
            // 
            this.expLabel.AutoSize = true;
            this.expLabel.Font = new System.Drawing.Font("Century Gothic", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.expLabel.ForeColor = System.Drawing.Color.White;
            this.expLabel.Location = new System.Drawing.Point(91, 334);
            this.expLabel.Name = "expLabel";
            this.expLabel.Size = new System.Drawing.Size(56, 21);
            this.expLabel.TabIndex = 10;
            this.expLabel.Text = "Стаж";
            // 
            // expText
            // 
            this.expText.Location = new System.Drawing.Point(211, 334);
            this.expText.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.expText.Name = "expText";
            this.expText.Size = new System.Drawing.Size(45, 22);
            this.expText.TabIndex = 11;
            this.expText.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.expText_KeyPress);
            // 
            // regButton2
            // 
            this.regButton2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.regButton2.Font = new System.Drawing.Font("Century Gothic", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.regButton2.ForeColor = System.Drawing.Color.White;
            this.regButton2.Location = new System.Drawing.Point(183, 491);
            this.regButton2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.regButton2.Name = "regButton2";
            this.regButton2.Size = new System.Drawing.Size(151, 33);
            this.regButton2.TabIndex = 12;
            this.regButton2.Text = "Регистрация";
            this.regButton2.UseVisualStyleBackColor = true;
            this.regButton2.Click += new System.EventHandler(this.regButton2_Click);
            // 
            // passRegText
            // 
            this.passRegText.AutoSize = true;
            this.passRegText.Font = new System.Drawing.Font("Century Gothic", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.passRegText.ForeColor = System.Drawing.Color.White;
            this.passRegText.Location = new System.Drawing.Point(91, 382);
            this.passRegText.Name = "passRegText";
            this.passRegText.Size = new System.Drawing.Size(74, 21);
            this.passRegText.TabIndex = 13;
            this.passRegText.Text = "Пароль";
            // 
            // passText
            // 
            this.passText.Location = new System.Drawing.Point(211, 382);
            this.passText.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.passText.Name = "passText";
            this.passText.Size = new System.Drawing.Size(205, 22);
            this.passText.TabIndex = 14;
            this.passText.UseSystemPasswordChar = true;
            // 
            // logText
            // 
            this.logText.Location = new System.Drawing.Point(211, 96);
            this.logText.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.logText.Name = "logText";
            this.logText.Size = new System.Drawing.Size(245, 22);
            this.logText.TabIndex = 16;
            this.logText.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.logText_KeyPress);
            // 
            // logLabel
            // 
            this.logLabel.AutoSize = true;
            this.logLabel.Font = new System.Drawing.Font("Century Gothic", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.logLabel.ForeColor = System.Drawing.Color.White;
            this.logLabel.Location = new System.Drawing.Point(91, 95);
            this.logLabel.Name = "logLabel";
            this.logLabel.Size = new System.Drawing.Size(58, 21);
            this.logLabel.TabIndex = 15;
            this.logLabel.Text = "Логин";
            // 
            // warnLabel
            // 
            this.warnLabel.AutoSize = true;
            this.warnLabel.Font = new System.Drawing.Font("Century Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.warnLabel.ForeColor = System.Drawing.Color.White;
            this.warnLabel.Location = new System.Drawing.Point(133, 444);
            this.warnLabel.Name = "warnLabel";
            this.warnLabel.Size = new System.Drawing.Size(0, 20);
            this.warnLabel.TabIndex = 17;
            // 
            // regForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(27)))), ((int)(((byte)(27)))));
            this.ClientSize = new System.Drawing.Size(537, 583);
            this.Controls.Add(this.warnLabel);
            this.Controls.Add(this.logText);
            this.Controls.Add(this.logLabel);
            this.Controls.Add(this.passText);
            this.Controls.Add(this.passRegText);
            this.Controls.Add(this.regButton2);
            this.Controls.Add(this.expText);
            this.Controls.Add(this.expLabel);
            this.Controls.Add(this.mailText);
            this.Controls.Add(this.mailLabel);
            this.Controls.Add(this.ageText);
            this.Controls.Add(this.ageLabel);
            this.Controls.Add(this.sexComboBox);
            this.Controls.Add(this.sexLabel);
            this.Controls.Add(this.nameText);
            this.Controls.Add(this.nameLabel);
            this.Controls.Add(this.postLabel);
            this.Controls.Add(this.postComboBox);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MaximizeBox = false;
            this.Name = "regForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Регистрация";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.regForm_FormClosed);
            this.Load += new System.EventHandler(this.RegForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox postComboBox;
        private System.Windows.Forms.Label postLabel;
        private System.Windows.Forms.Label nameLabel;
        private System.Windows.Forms.TextBox nameText;
        private System.Windows.Forms.Label sexLabel;
        private System.Windows.Forms.ComboBox sexComboBox;
        private System.Windows.Forms.Label ageLabel;
        private System.Windows.Forms.TextBox ageText;
        private System.Windows.Forms.Label mailLabel;
        private System.Windows.Forms.TextBox mailText;
        private System.Windows.Forms.Label expLabel;
        private System.Windows.Forms.TextBox expText;
        private System.Windows.Forms.Button regButton2;
        private System.Windows.Forms.Label passRegText;
        private System.Windows.Forms.TextBox passText;
        private System.Windows.Forms.TextBox logText;
        private System.Windows.Forms.Label logLabel;
        private System.Windows.Forms.Label warnLabel;
    }
}