namespace ITAccomEmailExtractC
{
    partial class FrmEmailExport
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmEmailExport));
            this.btnEmailAtt = new System.Windows.Forms.Button();
            this.txtAttach = new System.Windows.Forms.TextBox();
            this.btnGetMailBox = new System.Windows.Forms.Button();
            this.txtbxResponse = new System.Windows.Forms.TextBox();
            this.lstBxAccount = new System.Windows.Forms.ListBox();
            this.lstBxFolder = new System.Windows.Forms.ListBox();
            this.txtbxResponseMove = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtbxAttach = new System.Windows.Forms.TextBox();
            this.txtbxMail = new System.Windows.Forms.TextBox();
            this.txtbxDesktoPath = new System.Windows.Forms.TextBox();
            this.txtbxServerPath = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnFileProcess = new System.Windows.Forms.Button();
            this.tbRecordsToInsert = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.btnGetProcess = new System.Windows.Forms.Button();
            this.btnProcessAllFiles = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.txtlstBxAccount = new System.Windows.Forms.TextBox();
            this.txtProcess = new System.Windows.Forms.TextBox();
            this.btnFileReProcess = new System.Windows.Forms.Button();
            this.btnCheckFiles = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.txtNewFile = new System.Windows.Forms.TextBox();
            this.BtnAddList = new System.Windows.Forms.Button();
            this.btnClearList = new System.Windows.Forms.Button();
            this.btnProcessOneFiles = new System.Windows.Forms.Button();
            this.BtnMarkRedo = new System.Windows.Forms.Button();
            this.btnGetProcessRedo = new System.Windows.Forms.Button();
            this.lstBxAccountOld = new System.Windows.Forms.ListBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.txtBoxFileType = new System.Windows.Forms.TextBox();
            this.btnGetoldUnprocess = new System.Windows.Forms.Button();
            this.dtePickNewDate = new System.Windows.Forms.DateTimePicker();
            this.bntoldProcess = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.txtBoxBook = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.btnFixProcess = new System.Windows.Forms.Button();
            this.btnaddDataTable = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnEmailAtt
            // 
            this.btnEmailAtt.Location = new System.Drawing.Point(532, 487);
            this.btnEmailAtt.Name = "btnEmailAtt";
            this.btnEmailAtt.Size = new System.Drawing.Size(327, 29);
            this.btnEmailAtt.TabIndex = 0;
            this.btnEmailAtt.Text = "Get ABSA SFTP Files";
            this.btnEmailAtt.UseVisualStyleBackColor = true;
            this.btnEmailAtt.Click += new System.EventHandler(this.btnEmailAtt_Click);
            // 
            // txtAttach
            // 
            this.txtAttach.Location = new System.Drawing.Point(40, 2);
            this.txtAttach.Multiline = true;
            this.txtAttach.Name = "txtAttach";
            this.txtAttach.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtAttach.Size = new System.Drawing.Size(819, 272);
            this.txtAttach.TabIndex = 1;
            // 
            // btnGetMailBox
            // 
            this.btnGetMailBox.Location = new System.Drawing.Point(45, 515);
            this.btnGetMailBox.Name = "btnGetMailBox";
            this.btnGetMailBox.Size = new System.Drawing.Size(154, 28);
            this.btnGetMailBox.TabIndex = 2;
            this.btnGetMailBox.Text = "Get Mail Box";
            this.btnGetMailBox.UseVisualStyleBackColor = true;
            this.btnGetMailBox.Click += new System.EventHandler(this.btnGetMailBox_Click);
            // 
            // txtbxResponse
            // 
            this.txtbxResponse.Location = new System.Drawing.Point(347, 492);
            this.txtbxResponse.Name = "txtbxResponse";
            this.txtbxResponse.Size = new System.Drawing.Size(70, 20);
            this.txtbxResponse.TabIndex = 3;
            this.txtbxResponse.Text = "DM9";
            this.txtbxResponse.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtbxResponse.TextChanged += new System.EventHandler(this.txtbxResponse_TextChanged);
            // 
            // lstBxAccount
            // 
            this.lstBxAccount.FormattingEnabled = true;
            this.lstBxAccount.Location = new System.Drawing.Point(45, 318);
            this.lstBxAccount.Name = "lstBxAccount";
            this.lstBxAccount.Size = new System.Drawing.Size(391, 108);
            this.lstBxAccount.TabIndex = 4;
            this.lstBxAccount.SelectedIndexChanged += new System.EventHandler(this.lstBxAccount_SelectedIndexChanged);
            // 
            // lstBxFolder
            // 
            this.lstBxFolder.FormattingEnabled = true;
            this.lstBxFolder.Location = new System.Drawing.Point(497, 318);
            this.lstBxFolder.Name = "lstBxFolder";
            this.lstBxFolder.Size = new System.Drawing.Size(362, 108);
            this.lstBxFolder.TabIndex = 5;
            // 
            // txtbxResponseMove
            // 
            this.txtbxResponseMove.Location = new System.Drawing.Point(423, 492);
            this.txtbxResponseMove.Name = "txtbxResponseMove";
            this.txtbxResponseMove.Size = new System.Drawing.Size(71, 20);
            this.txtbxResponseMove.TabIndex = 6;
            this.txtbxResponseMove.Text = "DONED9M";
            this.txtbxResponseMove.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(220, 492);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(121, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Folder Check and Move";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(220, 523);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(120, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "No Mail and attachment";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // txtbxAttach
            // 
            this.txtbxAttach.Location = new System.Drawing.Point(423, 523);
            this.txtbxAttach.Name = "txtbxAttach";
            this.txtbxAttach.Size = new System.Drawing.Size(71, 20);
            this.txtbxAttach.TabIndex = 9;
            this.txtbxAttach.Text = "0";
            this.txtbxAttach.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // txtbxMail
            // 
            this.txtbxMail.Location = new System.Drawing.Point(347, 523);
            this.txtbxMail.Name = "txtbxMail";
            this.txtbxMail.Size = new System.Drawing.Size(70, 20);
            this.txtbxMail.TabIndex = 8;
            this.txtbxMail.Text = "0";
            this.txtbxMail.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // txtbxDesktoPath
            // 
            this.txtbxDesktoPath.Location = new System.Drawing.Point(101, 458);
            this.txtbxDesktoPath.Name = "txtbxDesktoPath";
            this.txtbxDesktoPath.Size = new System.Drawing.Size(335, 20);
            this.txtbxDesktoPath.TabIndex = 11;
            // 
            // txtbxServerPath
            // 
            this.txtbxServerPath.Location = new System.Drawing.Point(524, 454);
            this.txtbxServerPath.Name = "txtbxServerPath";
            this.txtbxServerPath.Size = new System.Drawing.Size(335, 20);
            this.txtbxServerPath.TabIndex = 12;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(10, 458);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(85, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "Desktop Folder :";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(442, 461);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(76, 13);
            this.label4.TabIndex = 14;
            this.label4.Text = "Server Folder :";
            // 
            // btnFileProcess
            // 
            this.btnFileProcess.Location = new System.Drawing.Point(532, 531);
            this.btnFileProcess.Name = "btnFileProcess";
            this.btnFileProcess.Size = new System.Drawing.Size(214, 29);
            this.btnFileProcess.TabIndex = 15;
            this.btnFileProcess.Text = "Process Selected File";
            this.btnFileProcess.UseVisualStyleBackColor = true;
            this.btnFileProcess.Click += new System.EventHandler(this.btnFileProcess_Click);
            // 
            // tbRecordsToInsert
            // 
            this.tbRecordsToInsert.Location = new System.Drawing.Point(347, 558);
            this.tbRecordsToInsert.Name = "tbRecordsToInsert";
            this.tbRecordsToInsert.Size = new System.Drawing.Size(70, 20);
            this.tbRecordsToInsert.TabIndex = 17;
            this.tbRecordsToInsert.Text = "500";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(229, 561);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(90, 13);
            this.label5.TabIndex = 16;
            this.label5.Text = "Records to insert:";
            // 
            // btnGetProcess
            // 
            this.btnGetProcess.Location = new System.Drawing.Point(445, 566);
            this.btnGetProcess.Name = "btnGetProcess";
            this.btnGetProcess.Size = new System.Drawing.Size(150, 29);
            this.btnGetProcess.TabIndex = 18;
            this.btnGetProcess.Text = "Get Unprocessed Files";
            this.btnGetProcess.UseVisualStyleBackColor = true;
            this.btnGetProcess.Click += new System.EventHandler(this.btnGetProcess_Click);
            // 
            // btnProcessAllFiles
            // 
            this.btnProcessAllFiles.Location = new System.Drawing.Point(752, 566);
            this.btnProcessAllFiles.Name = "btnProcessAllFiles";
            this.btnProcessAllFiles.Size = new System.Drawing.Size(181, 29);
            this.btnProcessAllFiles.TabIndex = 19;
            this.btnProcessAllFiles.Text = "Process All Re/Unprocessed Files";
            this.btnProcessAllFiles.UseVisualStyleBackColor = true;
            this.btnProcessAllFiles.Click += new System.EventHandler(this.btnProcessAllFiles_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(10, 432);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(85, 13);
            this.label6.TabIndex = 21;
            this.label6.Text = "Desktop Folder :";
            // 
            // txtlstBxAccount
            // 
            this.txtlstBxAccount.Location = new System.Drawing.Point(101, 432);
            this.txtlstBxAccount.Name = "txtlstBxAccount";
            this.txtlstBxAccount.Size = new System.Drawing.Size(335, 20);
            this.txtlstBxAccount.TabIndex = 20;
            // 
            // txtProcess
            // 
            this.txtProcess.Location = new System.Drawing.Point(40, 280);
            this.txtProcess.Name = "txtProcess";
            this.txtProcess.Size = new System.Drawing.Size(819, 20);
            this.txtProcess.TabIndex = 22;
            // 
            // btnFileReProcess
            // 
            this.btnFileReProcess.Enabled = false;
            this.btnFileReProcess.Location = new System.Drawing.Point(601, 566);
            this.btnFileReProcess.Name = "btnFileReProcess";
            this.btnFileReProcess.Size = new System.Drawing.Size(145, 29);
            this.btnFileReProcess.TabIndex = 23;
            this.btnFileReProcess.Text = "GET Files to Re-Process";
            this.btnFileReProcess.UseVisualStyleBackColor = true;
            this.btnFileReProcess.Click += new System.EventHandler(this.btnFileReProcess_Click);
            // 
            // btnCheckFiles
            // 
            this.btnCheckFiles.Location = new System.Drawing.Point(752, 531);
            this.btnCheckFiles.Name = "btnCheckFiles";
            this.btnCheckFiles.Size = new System.Drawing.Size(177, 28);
            this.btnCheckFiles.TabIndex = 24;
            this.btnCheckFiles.Text = "Check File";
            this.btnCheckFiles.UseVisualStyleBackColor = true;
            this.btnCheckFiles.Click += new System.EventHandler(this.btnCheckFiles_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(19, 607);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(76, 13);
            this.label7.TabIndex = 26;
            this.label7.Text = "Server Folder :";
            // 
            // txtNewFile
            // 
            this.txtNewFile.Location = new System.Drawing.Point(101, 607);
            this.txtNewFile.Name = "txtNewFile";
            this.txtNewFile.Size = new System.Drawing.Size(252, 20);
            this.txtNewFile.TabIndex = 25;
            // 
            // BtnAddList
            // 
            this.BtnAddList.Location = new System.Drawing.Point(359, 607);
            this.BtnAddList.Name = "BtnAddList";
            this.BtnAddList.Size = new System.Drawing.Size(74, 21);
            this.BtnAddList.TabIndex = 27;
            this.BtnAddList.Text = "AddToList";
            this.BtnAddList.UseVisualStyleBackColor = true;
            this.BtnAddList.Click += new System.EventHandler(this.BtnAddList_Click);
            // 
            // btnClearList
            // 
            this.btnClearList.Location = new System.Drawing.Point(615, 609);
            this.btnClearList.Name = "btnClearList";
            this.btnClearList.Size = new System.Drawing.Size(91, 26);
            this.btnClearList.TabIndex = 28;
            this.btnClearList.Text = "Clear List";
            this.btnClearList.UseVisualStyleBackColor = true;
            this.btnClearList.Click += new System.EventHandler(this.btnClearList_Click);
            // 
            // btnProcessOneFiles
            // 
            this.btnProcessOneFiles.Location = new System.Drawing.Point(819, 599);
            this.btnProcessOneFiles.Name = "btnProcessOneFiles";
            this.btnProcessOneFiles.Size = new System.Drawing.Size(109, 27);
            this.btnProcessOneFiles.TabIndex = 29;
            this.btnProcessOneFiles.Text = "Process Single";
            this.btnProcessOneFiles.UseVisualStyleBackColor = true;
            this.btnProcessOneFiles.Click += new System.EventHandler(this.btnSingleProcess_Click);
            // 
            // BtnMarkRedo
            // 
            this.BtnMarkRedo.Location = new System.Drawing.Point(712, 601);
            this.BtnMarkRedo.Name = "BtnMarkRedo";
            this.BtnMarkRedo.Size = new System.Drawing.Size(101, 25);
            this.BtnMarkRedo.TabIndex = 30;
            this.BtnMarkRedo.Text = "Marked Redo";
            this.BtnMarkRedo.UseVisualStyleBackColor = true;
            this.BtnMarkRedo.Click += new System.EventHandler(this.BtnMarkRedo_Click);
            // 
            // btnGetProcessRedo
            // 
            this.btnGetProcessRedo.Location = new System.Drawing.Point(445, 607);
            this.btnGetProcessRedo.Name = "btnGetProcessRedo";
            this.btnGetProcessRedo.Size = new System.Drawing.Size(150, 29);
            this.btnGetProcessRedo.TabIndex = 31;
            this.btnGetProcessRedo.Text = "Get Redo Files";
            this.btnGetProcessRedo.UseVisualStyleBackColor = true;
            this.btnGetProcessRedo.Click += new System.EventHandler(this.btnGetProcessRedo_Click);
            // 
            // lstBxAccountOld
            // 
            this.lstBxAccountOld.FormattingEnabled = true;
            this.lstBxAccountOld.Location = new System.Drawing.Point(913, 318);
            this.lstBxAccountOld.Name = "lstBxAccountOld";
            this.lstBxAccountOld.Size = new System.Drawing.Size(391, 108);
            this.lstBxAccountOld.TabIndex = 32;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(970, 516);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(30, 13);
            this.label8.TabIndex = 36;
            this.label8.Text = "Date";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(970, 485);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(46, 13);
            this.label9.TabIndex = 34;
            this.label9.Text = "File type";
            // 
            // txtBoxFileType
            // 
            this.txtBoxFileType.Location = new System.Drawing.Point(1041, 485);
            this.txtBoxFileType.Name = "txtBoxFileType";
            this.txtBoxFileType.Size = new System.Drawing.Size(126, 20);
            this.txtBoxFileType.TabIndex = 33;
            this.txtBoxFileType.Text = "ACCOUNT";
            this.txtBoxFileType.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // btnGetoldUnprocess
            // 
            this.btnGetoldUnprocess.Location = new System.Drawing.Point(1140, 556);
            this.btnGetoldUnprocess.Name = "btnGetoldUnprocess";
            this.btnGetoldUnprocess.Size = new System.Drawing.Size(150, 29);
            this.btnGetoldUnprocess.TabIndex = 37;
            this.btnGetoldUnprocess.Text = "Get OLD Unprocessed Files";
            this.btnGetoldUnprocess.UseVisualStyleBackColor = true;
            this.btnGetoldUnprocess.Click += new System.EventHandler(this.btnGetoldUnprocess_Click);
            // 
            // dtePickNewDate
            // 
            this.dtePickNewDate.CustomFormat = "";
            this.dtePickNewDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtePickNewDate.Location = new System.Drawing.Point(1041, 516);
            this.dtePickNewDate.Name = "dtePickNewDate";
            this.dtePickNewDate.Size = new System.Drawing.Size(126, 20);
            this.dtePickNewDate.TabIndex = 38;
            // 
            // bntoldProcess
            // 
            this.bntoldProcess.Location = new System.Drawing.Point(1140, 591);
            this.bntoldProcess.Name = "bntoldProcess";
            this.bntoldProcess.Size = new System.Drawing.Size(189, 29);
            this.bntoldProcess.TabIndex = 39;
            this.bntoldProcess.Text = "Process OLD All Re/Unprocessed Files";
            this.bntoldProcess.UseVisualStyleBackColor = true;
            this.bntoldProcess.Click += new System.EventHandler(this.bntoldProcess_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(970, 454);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(32, 13);
            this.label10.TabIndex = 41;
            this.label10.Text = "Book";
            // 
            // txtBoxBook
            // 
            this.txtBoxBook.Location = new System.Drawing.Point(1041, 454);
            this.txtBoxBook.Name = "txtBoxBook";
            this.txtBoxBook.Size = new System.Drawing.Size(126, 20);
            this.txtBoxBook.TabIndex = 40;
            this.txtBoxBook.Text = "PEDCBHMD";
            this.txtBoxBook.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(1183, 516);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(33, 13);
            this.label11.TabIndex = 42;
            this.label11.Text = "Datef";
            // 
            // btnFixProcess
            // 
            this.btnFixProcess.Location = new System.Drawing.Point(953, 597);
            this.btnFixProcess.Name = "btnFixProcess";
            this.btnFixProcess.Size = new System.Drawing.Size(181, 29);
            this.btnFixProcess.TabIndex = 43;
            this.btnFixProcess.Text = "Fix OLD All Re/Unprocessed Files";
            this.btnFixProcess.UseVisualStyleBackColor = true;
            this.btnFixProcess.Click += new System.EventHandler(this.btnFixProcess_Click);
            // 
            // btnaddDataTable
            // 
            this.btnaddDataTable.Location = new System.Drawing.Point(953, 626);
            this.btnaddDataTable.Name = "btnaddDataTable";
            this.btnaddDataTable.Size = new System.Drawing.Size(164, 26);
            this.btnaddDataTable.TabIndex = 45;
            this.btnaddDataTable.Text = "Load file to Datatable";
            this.btnaddDataTable.UseVisualStyleBackColor = true;
            this.btnaddDataTable.Click += new System.EventHandler(this.btnaddDataTable_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(712, 626);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(147, 25);
            this.button1.TabIndex = 46;
            this.button1.Text = "Move redo to Unprocessed";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(953, 556);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(181, 39);
            this.button2.TabIndex = 47;
            this.button2.Text = "Process LegalR All Re/Unprocessed Files";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // FrmEmailExport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1333, 657);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnaddDataTable);
            this.Controls.Add(this.btnFixProcess);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.txtBoxBook);
            this.Controls.Add(this.bntoldProcess);
            this.Controls.Add(this.dtePickNewDate);
            this.Controls.Add(this.btnGetoldUnprocess);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.txtBoxFileType);
            this.Controls.Add(this.lstBxAccountOld);
            this.Controls.Add(this.btnGetProcessRedo);
            this.Controls.Add(this.BtnMarkRedo);
            this.Controls.Add(this.btnProcessOneFiles);
            this.Controls.Add(this.btnClearList);
            this.Controls.Add(this.BtnAddList);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtNewFile);
            this.Controls.Add(this.btnCheckFiles);
            this.Controls.Add(this.btnFileReProcess);
            this.Controls.Add(this.txtProcess);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtlstBxAccount);
            this.Controls.Add(this.btnProcessAllFiles);
            this.Controls.Add(this.btnGetProcess);
            this.Controls.Add(this.tbRecordsToInsert);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnFileProcess);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtbxServerPath);
            this.Controls.Add(this.txtbxDesktoPath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtbxAttach);
            this.Controls.Add(this.txtbxMail);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtbxResponseMove);
            this.Controls.Add(this.lstBxFolder);
            this.Controls.Add(this.lstBxAccount);
            this.Controls.Add(this.txtbxResponse);
            this.Controls.Add(this.btnGetMailBox);
            this.Controls.Add(this.txtAttach);
            this.Controls.Add(this.btnEmailAtt);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmEmailExport";
            this.Text = "IT DM9 Attachment Export  ( Ver 5.01)";
            this.Load += new System.EventHandler(this.FrmEmailExport_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnEmailAtt;
        private System.Windows.Forms.TextBox txtAttach;
        private System.Windows.Forms.Button btnGetMailBox;
        private System.Windows.Forms.TextBox txtbxResponse;
        private System.Windows.Forms.ListBox lstBxAccount;
        private System.Windows.Forms.ListBox lstBxFolder;
        private System.Windows.Forms.TextBox txtbxResponseMove;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtbxAttach;
        private System.Windows.Forms.TextBox txtbxMail;
        private System.Windows.Forms.TextBox txtbxDesktoPath;
        private System.Windows.Forms.TextBox txtbxServerPath;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnFileProcess;
        private System.Windows.Forms.TextBox tbRecordsToInsert;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnGetProcess;
        private System.Windows.Forms.Button btnProcessAllFiles;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtlstBxAccount;
        private System.Windows.Forms.TextBox txtProcess;
        private System.Windows.Forms.Button btnFileReProcess;
        private System.Windows.Forms.Button btnCheckFiles;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtNewFile;
        private System.Windows.Forms.Button BtnAddList;
        private System.Windows.Forms.Button btnClearList;
        private System.Windows.Forms.Button btnProcessOneFiles;
        private System.Windows.Forms.Button BtnMarkRedo;
        private System.Windows.Forms.Button btnGetProcessRedo;
        private System.Windows.Forms.ListBox lstBxAccountOld;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txtBoxFileType;
        private System.Windows.Forms.Button btnGetoldUnprocess;
        private System.Windows.Forms.DateTimePicker dtePickNewDate;
        private System.Windows.Forms.Button bntoldProcess;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox txtBoxBook;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button btnFixProcess;
        private System.Windows.Forms.Button btnaddDataTable;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
    }
}

