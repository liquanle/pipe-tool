namespace DataClient
{
    partial class DataImport
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DataImport));
            this.btnClose = new System.Windows.Forms.Button();
            this.SetMDB = new System.Windows.Forms.Button();
            this.openMDB = new System.Windows.Forms.OpenFileDialog();
            this.ExcelPath = new System.Windows.Forms.TextBox();
            this.oleMdbConnection = new System.Data.OleDb.OleDbConnection();
            this.SetOutMDB = new System.Windows.Forms.Button();
            this.oleMdbCommand = new System.Data.OleDb.OleDbCommand();
            this.oleDbCommandBuilder1 = new System.Data.OleDb.OleDbCommandBuilder();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtProjectName = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.lbOuput = new System.Windows.Forms.ListBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnSelectNone = new System.Windows.Forms.Button();
            this.btnSelectInvert = new System.Windows.Forms.Button();
            this.btnSelectAll = new System.Windows.Forms.Button();
            this.chkLayerList = new System.Windows.Forms.CheckedListBox();
            this.savePersonMDB = new System.Windows.Forms.SaveFileDialog();
            this.PersonMDB = new System.Windows.Forms.TextBox();
            this.AllButton = new System.Windows.Forms.Button();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnClose
            // 
            this.btnClose.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnClose.Location = new System.Drawing.Point(431, 44);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(78, 26);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "关闭";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // SetMDB
            // 
            this.SetMDB.Location = new System.Drawing.Point(8, 10);
            this.SetMDB.Name = "SetMDB";
            this.SetMDB.Size = new System.Drawing.Size(78, 23);
            this.SetMDB.TabIndex = 1;
            this.SetMDB.Text = "设置原始表";
            this.SetMDB.Click += new System.EventHandler(this.SetExcel_Click);
            // 
            // openMDB
            // 
            this.openMDB.Filter = "ACCEES文件 (*.mdb)|*.mdb";
            this.openMDB.Title = "设置原始表";
            // 
            // ExcelPath
            // 
            this.ExcelPath.Location = new System.Drawing.Point(91, 12);
            this.ExcelPath.Name = "ExcelPath";
            this.ExcelPath.Size = new System.Drawing.Size(324, 21);
            this.ExcelPath.TabIndex = 2;
            // 
            // SetOutMDB
            // 
            this.SetOutMDB.Location = new System.Drawing.Point(8, 41);
            this.SetOutMDB.Name = "SetOutMDB";
            this.SetOutMDB.Size = new System.Drawing.Size(76, 23);
            this.SetOutMDB.TabIndex = 3;
            this.SetOutMDB.Text = "输出文件";
            this.SetOutMDB.UseVisualStyleBackColor = true;
            this.SetOutMDB.Click += new System.EventHandler(this.SetOutMDB_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtProjectName);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.lbOuput);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Controls.Add(this.btnSelectNone);
            this.groupBox1.Controls.Add(this.btnSelectInvert);
            this.groupBox1.Controls.Add(this.btnSelectAll);
            this.groupBox1.Controls.Add(this.chkLayerList);
            this.groupBox1.Location = new System.Drawing.Point(8, 84);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(516, 365);
            this.groupBox1.TabIndex = 15;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "图层信息";
            // 
            // txtProjectName
            // 
            this.txtProjectName.Location = new System.Drawing.Point(79, 197);
            this.txtProjectName.Name = "txtProjectName";
            this.txtProjectName.Size = new System.Drawing.Size(423, 21);
            this.txtProjectName.TabIndex = 29;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(8, 201);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 12);
            this.label3.TabIndex = 28;
            this.label3.Text = "项目名称：";
            // 
            // lbOuput
            // 
            this.lbOuput.FormattingEnabled = true;
            this.lbOuput.ItemHeight = 12;
            this.lbOuput.Location = new System.Drawing.Point(10, 229);
            this.lbOuput.Name = "lbOuput";
            this.lbOuput.Size = new System.Drawing.Size(493, 124);
            this.lbOuput.TabIndex = 27;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Location = new System.Drawing.Point(346, 20);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(156, 164);
            this.groupBox2.TabIndex = 27;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "使用说明";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(8, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(144, 114);
            this.label1.TabIndex = 26;
            this.label1.Text = "原始表：excel格式的原始数据\r\n输出文件：mdb文件\r\n使用说明：选择要进行转换的图层，点击开始转换按钮，进行数据转换。";
            // 
            // btnSelectNone
            // 
            this.btnSelectNone.Location = new System.Drawing.Point(261, 84);
            this.btnSelectNone.Name = "btnSelectNone";
            this.btnSelectNone.Size = new System.Drawing.Size(74, 29);
            this.btnSelectNone.TabIndex = 26;
            this.btnSelectNone.Text = "全不选(&B)";
            this.btnSelectNone.UseVisualStyleBackColor = true;
            this.btnSelectNone.Click += new System.EventHandler(this.btnSelectNone_Click);
            // 
            // btnSelectInvert
            // 
            this.btnSelectInvert.Location = new System.Drawing.Point(261, 134);
            this.btnSelectInvert.Name = "btnSelectInvert";
            this.btnSelectInvert.Size = new System.Drawing.Size(74, 29);
            this.btnSelectInvert.TabIndex = 25;
            this.btnSelectInvert.Text = "反选(&F)";
            this.btnSelectInvert.UseVisualStyleBackColor = true;
            this.btnSelectInvert.Click += new System.EventHandler(this.btnSelectInvert_Click);
            // 
            // btnSelectAll
            // 
            this.btnSelectAll.Location = new System.Drawing.Point(261, 35);
            this.btnSelectAll.Name = "btnSelectAll";
            this.btnSelectAll.Size = new System.Drawing.Size(74, 29);
            this.btnSelectAll.TabIndex = 24;
            this.btnSelectAll.Text = "全选(&A)";
            this.btnSelectAll.UseVisualStyleBackColor = true;
            this.btnSelectAll.Click += new System.EventHandler(this.btnSelectAll_Click);
            // 
            // chkLayerList
            // 
            this.chkLayerList.CheckOnClick = true;
            this.chkLayerList.FormattingEnabled = true;
            this.chkLayerList.Location = new System.Drawing.Point(8, 20);
            this.chkLayerList.Name = "chkLayerList";
            this.chkLayerList.Size = new System.Drawing.Size(237, 164);
            this.chkLayerList.TabIndex = 22;
            // 
            // savePersonMDB
            // 
            this.savePersonMDB.Filter = "个人数据库(*.mdb)|*.mdb";
            this.savePersonMDB.OverwritePrompt = false;
            this.savePersonMDB.Title = "保存或选择数据库文件";
            // 
            // PersonMDB
            // 
            this.PersonMDB.Location = new System.Drawing.Point(92, 44);
            this.PersonMDB.Name = "PersonMDB";
            this.PersonMDB.Size = new System.Drawing.Size(323, 21);
            this.PersonMDB.TabIndex = 16;
            // 
            // AllButton
            // 
            this.AllButton.Location = new System.Drawing.Point(431, 10);
            this.AllButton.Name = "AllButton";
            this.AllButton.Size = new System.Drawing.Size(78, 26);
            this.AllButton.TabIndex = 18;
            this.AllButton.Text = "开始转换";
            this.AllButton.UseVisualStyleBackColor = true;
            this.AllButton.Click += new System.EventHandler(this.AllButton_Click);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "wmploc_dll_367.bmp");
            this.imageList1.Images.SetKeyName(1, "wmploc_dll_366.bmp");
            // 
            // DataImport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(536, 457);
            this.Controls.Add(this.AllButton);
            this.Controls.Add(this.PersonMDB);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.SetOutMDB);
            this.Controls.Add(this.ExcelPath);
            this.Controls.Add(this.SetMDB);
            this.Controls.Add(this.btnClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "DataImport";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "excel转mdb";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.DataImport_FormClosing);
            this.Load += new System.EventHandler(this.DataImport_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button SetMDB;
        private System.Windows.Forms.OpenFileDialog openMDB;
        private System.Windows.Forms.TextBox ExcelPath;
        private System.Data.OleDb.OleDbConnection oleMdbConnection;
        private System.Windows.Forms.Button SetOutMDB;
        private System.Data.OleDb.OleDbCommand oleMdbCommand;
        private System.Data.OleDb.OleDbCommandBuilder oleDbCommandBuilder1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.SaveFileDialog savePersonMDB;
        private System.Windows.Forms.TextBox PersonMDB;
        private System.Windows.Forms.Button AllButton;
        private System.Windows.Forms.CheckedListBox chkLayerList;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox lbOuput;
        private System.Windows.Forms.Button btnSelectNone;
        private System.Windows.Forms.Button btnSelectInvert;
        private System.Windows.Forms.Button btnSelectAll;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox txtProjectName;
        private System.Windows.Forms.Label label3;
    }
}