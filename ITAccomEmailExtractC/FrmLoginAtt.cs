using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ITAccomEmailExtractC
{
    public partial class FrmLoginAtt : Form
    {
        public FrmLoginAtt()
        {
            InitializeComponent();
        }

        private bool IsvalidUserPass()

        {
            string newotp = (txtOTP.Text);

           

            DataClassLegalsuiteDataContext context = new DataClassLegalsuiteDataContext();
            var q = from p in context.Employees
                    where p.LoginID.ToUpper() == txtbxUser.Text &&
                     p.Password.ToUpper() == txtbxPassword.Text &&
                    p.EmployeeTypeID != 19 &&
                   p.OTP == int.Parse(newotp)
                    select p;
            if (q.Any())
            {
                return true;
            }
            else
            {
                return false;
            }

        }



        private bool IsvalidUser()

        {
            string newotp = (txtOTP.Text);
            DataClassLegalsuiteDataContext context = new DataClassLegalsuiteDataContext();
            
            var q = from p in context.Employees
                    where p.LoginID.ToUpper() == txtbxUser.Text &&
                    p.EmployeeTypeID != 19 &&
                   p.OTP == int.Parse(newotp)

                    select p;
            if (q.Any())
            {
                return true;
            }
            else
            {
                return false;
            }

        }



        private void btnLogon_Click(object sender, EventArgs e)
        {
            if (IsvalidUserPass())
            {
                DialogResult = DialogResult.OK;
            }
            else
            {
                if (IsvalidUser())
                    {
                    txtbxWrong.Text = "Incorrect Password";
                    txtbxWrong.Visible = true;
                    txtbxPassword.Focus();
                   
                }
                else
                    {
                    txtbxWrong.Text = "Incorrect User";
                    txtbxWrong.Visible = true;
                    txtbxUser.Focus();
                }
            }

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmLoginAtt_Load(object sender, EventArgs e)
        {

        }

        private void txtOTP_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnOTP_Click(object sender, EventArgs e)
        {
            if (webView21 != null && webView21.CoreWebView2 != null)
            {
                webView21.CoreWebView2.Navigate("http://192.168.3.3/justotp.aspx");
            }

        }
    }


}
