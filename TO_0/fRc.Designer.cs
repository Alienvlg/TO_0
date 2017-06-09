namespace TO_0
{
    partial class fRc
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
            this.btnExit = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.cmbBts = new System.Windows.Forms.ComboBox();
            this.tbName = new System.Windows.Forms.TextBox();
            this.tbAdress = new System.Windows.Forms.TextBox();
            this.tbInfo = new System.Windows.Forms.TextBox();
            this.tbMap = new System.Windows.Forms.TextBox();
            this.tbContact = new System.Windows.Forms.TextBox();
            this.tbAccess = new System.Windows.Forms.TextBox();
            this.tbAccesList = new System.Windows.Forms.TextBox();
            this.cmbDgu = new System.Windows.Forms.ComboBox();
            this.tbMetr = new System.Windows.Forms.TextBox();
            this.tbPwr = new System.Windows.Forms.TextBox();
            this.tbKey = new System.Windows.Forms.TextBox();
            this.tbRoof = new System.Windows.Forms.TextBox();
            this.btnCreateDir = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(807, 554);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(233, 59);
            this.btnExit.TabIndex = 51;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(114, 573);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(211, 58);
            this.btnExport.TabIndex = 50;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // cmbBts
            // 
            this.cmbBts.FormattingEnabled = true;
            this.cmbBts.Location = new System.Drawing.Point(12, 12);
            this.cmbBts.Name = "cmbBts";
            this.cmbBts.Size = new System.Drawing.Size(121, 21);
            this.cmbBts.TabIndex = 0;
            this.cmbBts.SelectedIndexChanged += new System.EventHandler(this.cmbBts_SelectedIndexChanged);
            // 
            // tbName
            // 
            this.tbName.Location = new System.Drawing.Point(139, 13);
            this.tbName.Name = "tbName";
            this.tbName.Size = new System.Drawing.Size(100, 20);
            this.tbName.TabIndex = 1;
            this.tbName.Text = "Название";
            // 
            // tbAdress
            // 
            this.tbAdress.Location = new System.Drawing.Point(245, 13);
            this.tbAdress.Name = "tbAdress";
            this.tbAdress.Size = new System.Drawing.Size(100, 20);
            this.tbAdress.TabIndex = 2;
            this.tbAdress.Text = "Адрес";
            // 
            // tbInfo
            // 
            this.tbInfo.Location = new System.Drawing.Point(351, 13);
            this.tbInfo.Name = "tbInfo";
            this.tbInfo.Size = new System.Drawing.Size(136, 20);
            this.tbInfo.TabIndex = 3;
            this.tbInfo.Text = "Экстренная информация";
            // 
            // tbMap
            // 
            this.tbMap.Location = new System.Drawing.Point(493, 13);
            this.tbMap.Name = "tbMap";
            this.tbMap.Size = new System.Drawing.Size(117, 20);
            this.tbMap.TabIndex = 4;
            this.tbMap.Text = "Месторасположение";
            // 
            // tbContact
            // 
            this.tbContact.Location = new System.Drawing.Point(626, 13);
            this.tbContact.Name = "tbContact";
            this.tbContact.Size = new System.Drawing.Size(100, 20);
            this.tbContact.TabIndex = 5;
            this.tbContact.Text = "Контактные лица";
            // 
            // tbAccess
            // 
            this.tbAccess.Location = new System.Drawing.Point(732, 13);
            this.tbAccess.Name = "tbAccess";
            this.tbAccess.Size = new System.Drawing.Size(134, 20);
            this.tbAccess.TabIndex = 6;
            this.tbAccess.Text = "Возможность прохода";
            // 
            // tbAccesList
            // 
            this.tbAccesList.Location = new System.Drawing.Point(872, 13);
            this.tbAccesList.Name = "tbAccesList";
            this.tbAccesList.Size = new System.Drawing.Size(135, 20);
            this.tbAccesList.TabIndex = 7;
            this.tbAccesList.Text = "Ограничение по списку";
            // 
            // cmbDgu
            // 
            this.cmbDgu.FormattingEnabled = true;
            this.cmbDgu.Items.AddRange(new object[] {
            "возможна",
            "не возможна"});
            this.cmbDgu.Location = new System.Drawing.Point(12, 39);
            this.cmbDgu.Name = "cmbDgu";
            this.cmbDgu.Size = new System.Drawing.Size(121, 21);
            this.cmbDgu.TabIndex = 8;
            this.cmbDgu.Text = "Установка ДГУ";
            // 
            // tbMetr
            // 
            this.tbMetr.Location = new System.Drawing.Point(139, 40);
            this.tbMetr.Name = "tbMetr";
            this.tbMetr.Size = new System.Drawing.Size(100, 20);
            this.tbMetr.TabIndex = 9;
            this.tbMetr.Text = "Длина кабеля";
            // 
            // tbPwr
            // 
            this.tbPwr.Location = new System.Drawing.Point(351, 40);
            this.tbPwr.Name = "tbPwr";
            this.tbPwr.Size = new System.Drawing.Size(136, 20);
            this.tbPwr.TabIndex = 11;
            this.tbPwr.Text = "Электроснабжение от";
            // 
            // tbKey
            // 
            this.tbKey.Location = new System.Drawing.Point(245, 40);
            this.tbKey.Name = "tbKey";
            this.tbKey.Size = new System.Drawing.Size(100, 20);
            this.tbKey.TabIndex = 10;
            this.tbKey.Text = "Ключи";
            // 
            // tbRoof
            // 
            this.tbRoof.Location = new System.Drawing.Point(493, 40);
            this.tbRoof.Name = "tbRoof";
            this.tbRoof.Size = new System.Drawing.Size(117, 20);
            this.tbRoof.TabIndex = 12;
            this.tbRoof.Text = "Выход на крышу";
            // 
            // btnCreateDir
            // 
            this.btnCreateDir.Location = new System.Drawing.Point(535, 590);
            this.btnCreateDir.Name = "btnCreateDir";
            this.btnCreateDir.Size = new System.Drawing.Size(75, 23);
            this.btnCreateDir.TabIndex = 52;
            this.btnCreateDir.Text = "create dir";
            this.btnCreateDir.UseVisualStyleBackColor = true;
            this.btnCreateDir.Click += new System.EventHandler(this.btnCreateDir_Click);
            // 
            // fRc
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1155, 643);
            this.Controls.Add(this.btnCreateDir);
            this.Controls.Add(this.tbRoof);
            this.Controls.Add(this.tbKey);
            this.Controls.Add(this.tbPwr);
            this.Controls.Add(this.tbMetr);
            this.Controls.Add(this.cmbDgu);
            this.Controls.Add(this.tbAccesList);
            this.Controls.Add(this.tbAccess);
            this.Controls.Add(this.tbContact);
            this.Controls.Add(this.tbMap);
            this.Controls.Add(this.tbInfo);
            this.Controls.Add(this.tbAdress);
            this.Controls.Add(this.tbName);
            this.Controls.Add(this.cmbBts);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnExit);
            this.Name = "fRc";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Учетная карточка";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.ComboBox cmbBts;
        private System.Windows.Forms.TextBox tbName;
        private System.Windows.Forms.TextBox tbAdress;
        private System.Windows.Forms.TextBox tbInfo;
        private System.Windows.Forms.TextBox tbMap;
        private System.Windows.Forms.TextBox tbContact;
        private System.Windows.Forms.TextBox tbAccess;
        private System.Windows.Forms.TextBox tbAccesList;
        private System.Windows.Forms.ComboBox cmbDgu;
        private System.Windows.Forms.TextBox tbMetr;
        private System.Windows.Forms.TextBox tbPwr;
        private System.Windows.Forms.TextBox tbKey;
        private System.Windows.Forms.TextBox tbRoof;
        private System.Windows.Forms.Button btnCreateDir;
    }
}