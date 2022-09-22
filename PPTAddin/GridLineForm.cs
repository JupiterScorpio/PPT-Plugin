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
    public partial class GridLineForm : Form
    {
        bool bcmbClicked = false;
        public GridLineForm()
        {
            InitializeComponent();
        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            if (!bcmbClicked)
                MessageBox.Show("Please select GridLine Space", "Unselected Item", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
            {
                switch(comboBox1.SelectedIndex)
                {
                    case 0:
                        ThisAddIn.nGridSpace = 4;
                        break;
                    case 1:
                        ThisAddIn.nGridSpace = 8;
                        break;
                    case 2:
                        ThisAddIn.nGridSpace = 16;
                        break;
                    case 3:
                        ThisAddIn.nGridSpace = 32;
                        break;
                    case 4:
                        ThisAddIn.nGridSpace = 10;
                        break;
                    case 5:
                        ThisAddIn.nGridSpace = 20;
                        break;
                    case 6:
                        ThisAddIn.nGridSpace = 14;
                        break;
                    case 7:
                        ThisAddIn.nGridSpace = 28;
                        break;
                    case 8:
                        ThisAddIn.nGridSpace = 56;
                        break;
                }
                ThisAddIn.DrawGrid();
                this.Close();
            }            
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            bcmbClicked = true;
        }
    }
}
