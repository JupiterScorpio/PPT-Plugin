using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PPTAddin
{
    public partial class ShortkeySel : Form
    {
        bool bsel = false;
        public ShortkeySel()
        {
            InitializeComponent();
        }

        private void ctrl_opt_CheckedChanged(object sender, EventArgs e)
        {
            ThisAddIn.nCtrlOShift = 1;
            bsel = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (bsel)
            {
                if(Globals.ThisAddIn.State==6|| Globals.ThisAddIn.State == 7)
                {
                    var frm = new CombinationalImport();
                    frm.Show();
                }
                this.Close();
            }
            else
                MessageBox.Show("Please Select a ShortKey", "Unselected Shortkey", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void ShortkeySel_Load(object sender, EventArgs e)
        {

        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            point.bfirst = false;
            this.Close();
        }

        private void shift_opt_CheckedChanged_1(object sender, EventArgs e)
        {
            ThisAddIn.nCtrlOShift = 2;
            bsel = true;
        }
    }
}
