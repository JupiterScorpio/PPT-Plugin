using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using stringrep;
using desEnDec;
using Microsoft.Win32;

namespace PPTAddin
{
    public partial class LicenseDlg : Form
    {
        public bool m_bPass;
        string str_result;
        private TextBox license_txt;
        private TextBox textBox4;
        private TextBox textBox3;
        private TextBox textBox2;
        private TextBox textBox1;
        private Button btn_cncl;
        private Button btn_ok;
        desClass des;
        public LicenseDlg()
        {
            InitializeComponent();
            m_bPass = false;
        }
        private void show_hdInfo()
        {
            string strTemp;
            int index;
            string _id = HardwareInfo.GenerateUID("pptAddin");
            strTemp = _id;
            //////////////////textbox1////////////////////////
            index = strTemp.IndexOf('-');
            string strBuf;
            strBuf = strTemp.Substring(0, index);
            textBox1.Text = strBuf;
            strTemp=strTemp.Substring(index + 1);
            index = 0;
            //////////////////textbox2////////////////////////
            index = strTemp.IndexOf('-');
            strBuf = strTemp.Substring(0, index);
            textBox2.Text = strBuf;
            strTemp = strTemp.Substring(index + 1);
            index = 0;
            //////////////////textbox3////////////////////////
            index = strTemp.IndexOf('-');
            strBuf = strTemp.Substring(0, index);
            textBox3.Text = strBuf;
            strTemp = strTemp.Substring(index + 1);
            index = 0;
            //////////////////textbox4////////////////////////
            index = strTemp.IndexOf('-');
            if (index < 0)
                index = strTemp.Length - 1;
            strBuf = strTemp.Substring(0, index);
            textBox4.Text = strBuf;
            strTemp = strTemp.Substring(index + 1);
            index = 0;
            //Class1 cls = new Class1();
            //cls.VCrypt();
        }
       


        private string LicenCalc()
        {
            des = new desClass();
            string lic1, lic2, lic3, lic4;
            lic1 = des.Encrypt(textBox1.Text);
            lic2 = des.Encrypt(textBox2.Text);
            lic3 = des.Encrypt(textBox3.Text);
            lic4 = des.Encrypt(textBox4.Text);
            str_result= string.Format("{0}-{1}-{2}-{3}", lic1, lic2, lic3, lic4);
            return str_result;
        }


        private void InitializeComponent()
        {
            this.license_txt = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btn_cncl = new System.Windows.Forms.Button();
            this.btn_ok = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // license_txt
            // 
            this.license_txt.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.license_txt.Location = new System.Drawing.Point(57, 78);
            this.license_txt.Name = "license_txt";
            this.license_txt.Size = new System.Drawing.Size(318, 26);
            this.license_txt.TabIndex = 13;
            // 
            // textBox4
            // 
            this.textBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox4.Location = new System.Drawing.Point(303, 24);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(72, 24);
            this.textBox4.TabIndex = 12;
            // 
            // textBox3
            // 
            this.textBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox3.Location = new System.Drawing.Point(207, 24);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(72, 24);
            this.textBox3.TabIndex = 11;
            // 
            // textBox2
            // 
            this.textBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox2.Location = new System.Drawing.Point(110, 24);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(72, 24);
            this.textBox2.TabIndex = 10;
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(16, 24);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(72, 24);
            this.textBox1.TabIndex = 9;
            // 
            // btn_cncl
            // 
            this.btn_cncl.Location = new System.Drawing.Point(301, 143);
            this.btn_cncl.Name = "btn_cncl";
            this.btn_cncl.Size = new System.Drawing.Size(74, 29);
            this.btn_cncl.TabIndex = 8;
            this.btn_cncl.Text = "Cancel";
            this.btn_cncl.UseVisualStyleBackColor = true;
            this.btn_cncl.Click += new System.EventHandler(this.btn_cncl_Click_1);
            // 
            // btn_ok
            // 
            this.btn_ok.Location = new System.Drawing.Point(207, 143);
            this.btn_ok.Name = "btn_ok";
            this.btn_ok.Size = new System.Drawing.Size(74, 29);
            this.btn_ok.TabIndex = 7;
            this.btn_ok.Text = "OK";
            this.btn_ok.UseVisualStyleBackColor = true;
            this.btn_ok.Click += new System.EventHandler(this.btn_ok_Click);
            // 
            // LicenseDlg
            // 
            this.ClientSize = new System.Drawing.Size(389, 187);
            this.Controls.Add(this.license_txt);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btn_cncl);
            this.Controls.Add(this.btn_ok);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "LicenseDlg";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "License Input";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.LicenseDlg_FormClosing);
            this.Load += new System.EventHandler(this.LicenseDlg_Load_1);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            if (license_txt.Text == LicenCalc())
            {
                DialogResult = DialogResult.OK;
                m_bPass = true;
                RegistryKey key = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\PPTaddin");
                key.SetValue("Set1", "addin");
                key.Close();
                MessageBox.Show("Successfully License Registered. Please restart Application");
                //ThisAddIn.ExitApp();
                this.Close();
            }
            else
            {
                m_bPass = false;
                DialogResult = DialogResult.Cancel;
                MessageBox.Show("Wrong License.Please retry.");
            }
        }

        private void LicenseDlg_Load_1(object sender, EventArgs e)
        {
            show_hdInfo();
        }

        private void btn_cncl_Click_1(object sender, EventArgs e)
        {
            m_bPass = false;
            //Application.Exit();
            this.Close();
        }

        private void LicenseDlg_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Application.Exit();
            this.Close();
        }
    }
}
