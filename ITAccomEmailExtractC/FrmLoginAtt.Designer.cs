namespace ITAccomEmailExtractC
{
    partial class FrmLoginAtt
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmLoginAtt));
            this.btnLogon = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.txtbxUser = new System.Windows.Forms.TextBox();
            this.txtbxPassword = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtbxWrong = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtOTP = new System.Windows.Forms.TextBox();
            this.btnOTP = new System.Windows.Forms.Button();
            this.webView21 = new Microsoft.Web.WebView2.WinForms.WebView2();
            ((System.ComponentModel.ISupportInitialize)(this.webView21)).BeginInit();
            this.SuspendLayout();
            // 
            // btnLogon
            // 
            this.btnLogon.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.btnLogon.Location = new System.Drawing.Point(22, 176);
            this.btnLogon.Name = "btnLogon";
            this.btnLogon.Size = new System.Drawing.Size(98, 30);
            this.btnLogon.TabIndex = 2;
            this.btnLogon.Text = "Logon";
            this.btnLogon.UseVisualStyleBackColor = false;
            this.btnLogon.Click += new System.EventHandler(this.btnLogon_Click);
            // 
            // btnExit
            // 
            this.btnExit.BackColor = System.Drawing.Color.LightCoral;
            this.btnExit.Location = new System.Drawing.Point(159, 177);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(105, 29);
            this.btnExit.TabIndex = 4;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = false;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // txtbxUser
            // 
            this.txtbxUser.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.txtbxUser.Location = new System.Drawing.Point(109, 34);
            this.txtbxUser.Name = "txtbxUser";
            this.txtbxUser.Size = new System.Drawing.Size(144, 20);
            this.txtbxUser.TabIndex = 0;
            // 
            // txtbxPassword
            // 
            this.txtbxPassword.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.txtbxPassword.Location = new System.Drawing.Point(109, 74);
            this.txtbxPassword.Name = "txtbxPassword";
            this.txtbxPassword.PasswordChar = '*';
            this.txtbxPassword.Size = new System.Drawing.Size(144, 20);
            this.txtbxPassword.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(28, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "User Name :";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(32, 77);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Password :";
            // 
            // txtbxWrong
            // 
            this.txtbxWrong.BackColor = System.Drawing.Color.Red;
            this.txtbxWrong.Location = new System.Drawing.Point(99, 141);
            this.txtbxWrong.Name = "txtbxWrong";
            this.txtbxWrong.Size = new System.Drawing.Size(174, 20);
            this.txtbxWrong.TabIndex = 6;
            this.txtbxWrong.Text = "User Name or Password Incorrect";
            this.txtbxWrong.Visible = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(32, 107);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "OTP :";
            // 
            // txtOTP
            // 
            this.txtOTP.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.txtOTP.Location = new System.Drawing.Point(109, 104);
            this.txtOTP.Name = "txtOTP";
            this.txtOTP.PasswordChar = '*';
            this.txtOTP.Size = new System.Drawing.Size(66, 20);
            this.txtOTP.TabIndex = 7;
            this.txtOTP.TextChanged += new System.EventHandler(this.txtOTP_TextChanged);
            // 
            // btnOTP
            // 
            this.btnOTP.Location = new System.Drawing.Point(181, 104);
            this.btnOTP.Name = "btnOTP";
            this.btnOTP.Size = new System.Drawing.Size(84, 22);
            this.btnOTP.TabIndex = 54;
            this.btnOTP.Text = "Get OTP";
            this.btnOTP.UseVisualStyleBackColor = true;
            this.btnOTP.Click += new System.EventHandler(this.btnOTP_Click);
            // 
            // webView21
            // 
            this.webView21.AllowExternalDrop = true;
            this.webView21.CreationProperties = null;
            this.webView21.DefaultBackgroundColor = System.Drawing.Color.White;
            this.webView21.Location = new System.Drawing.Point(3, 212);
            this.webView21.Name = "webView21";
            this.webView21.Size = new System.Drawing.Size(325, 82);
            this.webView21.Source = new System.Uri("http://192.168.3.3/justotp.aspx", System.UriKind.Absolute);
            this.webView21.TabIndex = 55;
            this.webView21.ZoomFactor = 1D;
            // 
            // FrmLoginAtt
            // 
            this.AcceptButton = this.btnLogon;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(332, 306);
            this.Controls.Add(this.webView21);
            this.Controls.Add(this.btnOTP);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtOTP);
            this.Controls.Add(this.txtbxWrong);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtbxPassword);
            this.Controls.Add(this.txtbxUser);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnLogon);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmLoginAtt";
            this.Text = "Login Dm9 Attachment (ver 6.01)";
            this.Load += new System.EventHandler(this.FrmLoginAtt_Load);
            ((System.ComponentModel.ISupportInitialize)(this.webView21)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnLogon;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.TextBox txtbxUser;
        private System.Windows.Forms.TextBox txtbxPassword;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtbxWrong;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtOTP;
        private System.Windows.Forms.Button btnOTP;
        private Microsoft.Web.WebView2.WinForms.WebView2 webView21;
    }
}