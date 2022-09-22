using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Slides;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Utilities;
using Microsoft.Win32;
using System.Data;
using System.Drawing;

namespace PPTAddin
{    
    public partial class point
    {
        public bool bor;
        public GeneralHook hook = new GeneralHook();
        public static bool bfirst = false;
        string curpath;
        public void button1_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            ThisAddIn.POL = true;
            Globals.ThisAddIn.State = (uint)ThisAddIn.shapestates.Box;
            lblState.Label = "Point";
            //if(!bfirst)
            //{
            //    bfirst = true;
            //    var shrtfrm = new ShortkeySel();
            //    shrtfrm.Show();
            //}
        }
        public void button2_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            ThisAddIn.POL = false;
            Globals.ThisAddIn.State = (uint)ThisAddIn.shapestates.Line;            
            lblState.Label = "Connector";
            //ThisAddIn.ResetPthistory();
        }

        private void btn_open_Click(object sender, RibbonControlEventArgs e)
        {           
            ThisAddIn.binit = true;            
                string path;
            ThisAddIn.slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            OpenFileDialog dialog = new OpenFileDialog();
                if (DialogResult.OK == dialog.ShowDialog())
                {
                    path= dialog.FileName;
                    ThisAddIn.path = dialog.FileName;
                    ThisAddIn.newpres = Globals.ThisAddIn.Application.Presentations.Open(path,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue);
                    ThisAddIn.binit = !ThisAddIn.binit;
                }
            lblState.Label = "Open File";
        }

        public void hook_OnMouseActivity(object sender, MouseEventArgs e)
        {
            xpos.Label = ThisAddIn.curXpos.ToString();
            ypos.Label = ThisAddIn.curYpos.ToString();
        }      

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void point_Load(object sender, RibbonUIEventArgs e)
        {
            hook.InstallHook(HookHelper.HookType.MouseOperation);
            hook.OnMouseActivity += new MouseEventHandler(hook_OnMouseActivity);
            curpath = System.IO.Directory.GetCurrentDirectory() + "\\libs.vlp";
            string text = System.IO.File.ReadAllText(curpath);
            int pos = text.IndexOf("\n");
            string first = text.Substring(0, pos);
            string second = text.Substring(pos + 1, text.Length - pos - 1);
            this.btn_lib1.Label = first;
            this.btn_lib2.Label = second;
            ThisAddIn.strlib1name = this.btn_lib1.Label;
            ThisAddIn.strlib2name = this.btn_lib2.Label;
            this.editBox1.Text= this.btn_lib1.Label;
            this.editBox2.Text = this.btn_lib2.Label;
            ThisAddIn.pointRibbon = this;
                
        }

        private void btn_grid_Click(object sender, RibbonControlEventArgs e)
        {
            if(!ThisAddIn.bGridcheck)
            {
                //var frm = new GridLineForm();
                //frm.Show();
                
                if(ThisAddIn.nGridSpace!=0)
                {
                    this.btn_grid.Image = Properties.Resources.grid;
                    ThisAddIn.DrawGrid();
                    ThisAddIn.bGridcheck = true;
                    lblState.Label = "Grid";
                }                    
                else
                    MessageBox.Show("Please select GridLine Space", "Unselected Item", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                
            }else
            {
                this.btn_grid.Image = Properties.Resources.notgrid;
                ThisAddIn.DelGrid();
                ThisAddIn.bGridcheck = false;
                lblState.Label = "No Grid";
            }
            Globals.ThisAddIn.State = 5;
        }

        private void btn_Inport_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            Globals.ThisAddIn.State = 3;
            lblState.Label = "In Port";
            //if (!bfirst)
            //{
            //    bfirst = true;
            //    var shrtfrm = new ShortkeySel();
            //    shrtfrm.Show();
            //}
        }

        private void btn_Outport_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            Globals.ThisAddIn.State = 4;
            lblState.Label = "Out Port";
            //if (!bfirst)
            //{
            //    bfirst = true;
            //    var shrtfrm = new ShortkeySel();
            //    shrtfrm.Show();
            //}
        }

        private void btn_lib1_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            Globals.ThisAddIn.State = 8;
            lblState.Label = "Library1";
            MdbManger.GetInstance().SaveMDB();
            var frm = new UserLiborm();
            frm.Show();
        }

        private void btn_lib2_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            Globals.ThisAddIn.State = 9;
            lblState.Label = "Library2";
            MdbManger.GetInstance().LoadMDB();
            var frm = new UserLiborm();
            frm.Show();
        }

        private void confirm_chk_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            Globals.ThisAddIn.State = 10;
            lblState.Label = "Additional Check";
        }

        private void drop_additional_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            Globals.ThisAddIn.State = 11;
            lblState.Label = "Additional Optional";
        }

        private void btn_export_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            var frm = new ExportSetting();
            frm.Show();
        }
        
        private void gallery_comb_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            Globals.ThisAddIn.State = 6;
            lblState.Label = "Combitional";
            int selindex = gallery_comb.SelectedItemIndex;
            ThisAddIn.nCombitionalshp = Group_Param.GetGroupCombnationStr(selindex);
        }

        private void gallery_seq_Click(object sender, RibbonControlEventArgs e)
        {
            int selindex=gallery_seq.SelectedItemIndex;
            ThisAddIn.slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            Globals.ThisAddIn.State = 7;
            lblState.Label = "Sequentional";
            ThisAddIn.nSeqshp = selindex;
        }

        private void btn_setting_Click(object sender, RibbonControlEventArgs e)
        {
            var shrtfrm = new ShortkeySel();
            shrtfrm.Show();
        }

        private void rename_btn_Click(object sender, RibbonControlEventArgs e)
        {
            System.IO.File.WriteAllText(curpath,this.editBox1.Text + "\n" + this.editBox2.Text);
            this.btn_lib2.Label = this.editBox2.Text;
            this.btn_lib1.Label = this.editBox1.Text;
        }

        private void checkBox1_Click(object sender, RibbonControlEventArgs e)
        {
            if(!editBox1.Enabled)
            {
                editBox1.Enabled = true;
                editBox2.Enabled = true;
                rename_btn.Enabled = true;
            }
            else
            {
                editBox1.Enabled = false;
                editBox2.Enabled = false;
                rename_btn.Enabled = false;
            }            
        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void editBox2_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void polyln_btn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.State = 12;
            lblState.Label = "WaveForm";
            ThisAddIn.slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
        }

        private void wv_regChk_Click(object sender, RibbonControlEventArgs e)
        {
            if (wv_regChk.Checked)
                ThisAddIn.bregMode = true;
            else
                ThisAddIn.bregMode = false;
        }
        public int[] gridSpaceAry = { 4, 8, 16, 32, 10, 20, 14, 28, 56 };
        private void cmb_gridSZ_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            if (gridSpaceAry.Length <= cmb_gridSZ.SelectedItemIndex) return;
            ThisAddIn.nGridSpace = gridSpaceAry[cmb_gridSZ.SelectedItemIndex];
            
            if(ThisAddIn.bGridcheck)
            {
                ThisAddIn.DelGrid();
                ThisAddIn.DrawGrid();
            }
            
        }

        private void cmb_gridSZ_Click(object sender, RibbonControlEventArgs e)
        {
            switch (cmb_gridSZ.SelectedItemIndex)
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
        }

        private void cmb_pinSZ_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            int oldptwidth = Point_Param.pinWidth;
            Point_Param.pinWidth = Convert.ToInt32( cmb_pinSZ.SelectedItem.ToString());
            ThisAddIn.ChangePointShapeSize(ThisAddIn.ShapeTypeFlag.PIN, Point_Param.pinWidth);
            //ThisAddIn.RedrawAllPoint();
        }

        private void cmb_portSZ_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            int oldptwidth = Point_Param.portWidth;
            Point_Param.portWidth = Convert.ToInt32(cmb_portSZ.SelectedItem.ToString());
            ThisAddIn.ChangePointShapeSize(ThisAddIn.ShapeTypeFlag.INPORT | ThisAddIn.ShapeTypeFlag.OUTPORT, Point_Param.portWidth);
            //ThisAddIn.RedrawAllPort();
        }
    }
}
