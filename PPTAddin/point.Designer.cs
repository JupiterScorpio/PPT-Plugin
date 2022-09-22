
using Microsoft.Office.Tools.Ribbon;

namespace PPTAddin
{
    partial class point : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public point()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl11 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl12 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl13 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl14 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl15 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl16 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl17 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl18 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl19 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl20 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl21 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl22 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl23 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl24 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl25 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl26 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl27 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl28 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl29 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl30 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl31 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl32 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl33 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl34 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl35 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl36 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl37 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.lblX = this.Factory.CreateRibbonLabel();
            this.lblY = this.Factory.CreateRibbonLabel();
            this.lblState = this.Factory.CreateRibbonLabel();
            this.xpos = this.Factory.CreateRibbonLabel();
            this.ypos = this.Factory.CreateRibbonLabel();
            this.btn_open = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.btn_Outport = this.Factory.CreateRibbonButton();
            this.btn_Inport = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.btn_grid = this.Factory.CreateRibbonButton();
            this.cmb_pinSZ = this.Factory.CreateRibbonDropDown();
            this.cmb_portSZ = this.Factory.CreateRibbonDropDown();
            this.cmb_gridSZ = this.Factory.CreateRibbonDropDown();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.gallery_comb = this.Factory.CreateRibbonGallery();
            this.gallery_seq = this.Factory.CreateRibbonGallery();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btn_export = this.Factory.CreateRibbonButton();
            this.btn_lib1 = this.Factory.CreateRibbonButton();
            this.btn_lib2 = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.confirm_chk = this.Factory.CreateRibbonCheckBox();
            this.drop_additional = this.Factory.CreateRibbonDropDown();
            this.group_tool = this.Factory.CreateRibbonGroup();
            this.wv_regChk = this.Factory.CreateRibbonCheckBox();
            this.polyln_btn = this.Factory.CreateRibbonButton();
            this.Settings = this.Factory.CreateRibbonGroup();
            this.btn_setting = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.editBox2 = this.Factory.CreateRibbonEditBox();
            this.checkBox1 = this.Factory.CreateRibbonCheckBox();
            this.rename_btn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group_tool.SuspendLayout();
            this.Settings.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group_tool);
            this.tab1.Groups.Add(this.Settings);
            this.tab1.Label = "TabConnection";
            this.tab1.Name = "tab1";
            this.tab1.Tag = "Point_Line";
            // 
            // group1
            // 
            ribbonDialogLauncherImpl1.Image = global::PPTAddin.Properties.Resources.point;
            this.group1.DialogLauncher = ribbonDialogLauncherImpl1;
            this.group1.Items.Add(this.label1);
            this.group1.Items.Add(this.lblX);
            this.group1.Items.Add(this.lblY);
            this.group1.Items.Add(this.lblState);
            this.group1.Items.Add(this.xpos);
            this.group1.Items.Add(this.ypos);
            this.group1.Items.Add(this.btn_open);
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.btn_Outport);
            this.group1.Items.Add(this.btn_Inport);
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.btn_grid);
            this.group1.Items.Add(this.cmb_pinSZ);
            this.group1.Items.Add(this.cmb_portSZ);
            this.group1.Items.Add(this.cmb_gridSZ);
            this.group1.Label = "Point and Connector";
            this.group1.Name = "group1";
            // 
            // label1
            // 
            this.label1.Label = "State:";
            this.label1.Name = "label1";
            // 
            // lblX
            // 
            this.lblX.Label = "X:";
            this.lblX.Name = "lblX";
            // 
            // lblY
            // 
            this.lblY.Label = "Y:";
            this.lblY.Name = "lblY";
            // 
            // lblState
            // 
            this.lblState.Label = "None";
            this.lblState.Name = "lblState";
            // 
            // xpos
            // 
            this.xpos.Label = "0";
            this.xpos.Name = "xpos";
            // 
            // ypos
            // 
            this.ypos.Label = "0";
            this.ypos.Name = "ypos";
            // 
            // btn_open
            // 
            this.btn_open.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_open.Image = global::PPTAddin.Properties.Resources.open;
            this.btn_open.Label = "Open";
            this.btn_open.Name = "btn_open";
            this.btn_open.ShowImage = true;
            this.btn_open.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_open_Click);
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::PPTAddin.Properties.Resources.rectangleimages;
            this.button1.Label = "Pin";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // btn_Outport
            // 
            this.btn_Outport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_Outport.Image = global::PPTAddin.Properties.Resources.outport;
            this.btn_Outport.Label = "OutPort";
            this.btn_Outport.Name = "btn_Outport";
            this.btn_Outport.ShowImage = true;
            this.btn_Outport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Outport_Click);
            // 
            // btn_Inport
            // 
            this.btn_Inport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_Inport.Image = global::PPTAddin.Properties.Resources.inport;
            this.btn_Inport.Label = "InPort";
            this.btn_Inport.Name = "btn_Inport";
            this.btn_Inport.ShowImage = true;
            this.btn_Inport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Inport_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = global::PPTAddin.Properties.Resources.line;
            this.button2.Label = "Wire";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // btn_grid
            // 
            this.btn_grid.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_grid.Image = global::PPTAddin.Properties.Resources.notgrid;
            this.btn_grid.Label = "Grid Line";
            this.btn_grid.Name = "btn_grid";
            this.btn_grid.ShowImage = true;
            this.btn_grid.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_grid_Click);
            // 
            // cmb_pinSZ
            // 
            ribbonDropDownItemImpl1.Label = "4";
            ribbonDropDownItemImpl2.Label = "5";
            ribbonDropDownItemImpl3.Label = "6";
            ribbonDropDownItemImpl4.Label = "7";
            ribbonDropDownItemImpl5.Label = "8";
            ribbonDropDownItemImpl6.Label = "9";
            ribbonDropDownItemImpl7.Label = "10";
            this.cmb_pinSZ.Items.Add(ribbonDropDownItemImpl1);
            this.cmb_pinSZ.Items.Add(ribbonDropDownItemImpl2);
            this.cmb_pinSZ.Items.Add(ribbonDropDownItemImpl3);
            this.cmb_pinSZ.Items.Add(ribbonDropDownItemImpl4);
            this.cmb_pinSZ.Items.Add(ribbonDropDownItemImpl5);
            this.cmb_pinSZ.Items.Add(ribbonDropDownItemImpl6);
            this.cmb_pinSZ.Items.Add(ribbonDropDownItemImpl7);
            this.cmb_pinSZ.Label = "Pin Size";
            this.cmb_pinSZ.Name = "cmb_pinSZ";
            this.cmb_pinSZ.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cmb_pinSZ_SelectionChanged);
            // 
            // cmb_portSZ
            // 
            ribbonDropDownItemImpl8.Label = "4";
            ribbonDropDownItemImpl9.Label = "5";
            ribbonDropDownItemImpl10.Label = "6";
            ribbonDropDownItemImpl11.Label = "7";
            ribbonDropDownItemImpl12.Label = "8";
            ribbonDropDownItemImpl13.Label = "9";
            ribbonDropDownItemImpl14.Label = "10";
            this.cmb_portSZ.Items.Add(ribbonDropDownItemImpl8);
            this.cmb_portSZ.Items.Add(ribbonDropDownItemImpl9);
            this.cmb_portSZ.Items.Add(ribbonDropDownItemImpl10);
            this.cmb_portSZ.Items.Add(ribbonDropDownItemImpl11);
            this.cmb_portSZ.Items.Add(ribbonDropDownItemImpl12);
            this.cmb_portSZ.Items.Add(ribbonDropDownItemImpl13);
            this.cmb_portSZ.Items.Add(ribbonDropDownItemImpl14);
            this.cmb_portSZ.Label = "Port Size";
            this.cmb_portSZ.Name = "cmb_portSZ";
            this.cmb_portSZ.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cmb_portSZ_SelectionChanged);
            // 
            // cmb_gridSZ
            // 
            ribbonDropDownItemImpl15.Label = "4 pts";
            ribbonDropDownItemImpl16.Label = "8 pts(~3 mm)";
            ribbonDropDownItemImpl17.Label = "16 pts(~6 mm)";
            ribbonDropDownItemImpl18.Label = "32 pts(~12 mm)";
            ribbonDropDownItemImpl19.Label = "10 pts(~3.5 mm)";
            ribbonDropDownItemImpl20.Label = "20 pts(~7 mm)";
            ribbonDropDownItemImpl21.Label = "14 pts(~5 mm)";
            ribbonDropDownItemImpl22.Label = "28 pts(~10 mm)";
            ribbonDropDownItemImpl23.Label = "56 pts(~20 mm)";
            this.cmb_gridSZ.Items.Add(ribbonDropDownItemImpl15);
            this.cmb_gridSZ.Items.Add(ribbonDropDownItemImpl16);
            this.cmb_gridSZ.Items.Add(ribbonDropDownItemImpl17);
            this.cmb_gridSZ.Items.Add(ribbonDropDownItemImpl18);
            this.cmb_gridSZ.Items.Add(ribbonDropDownItemImpl19);
            this.cmb_gridSZ.Items.Add(ribbonDropDownItemImpl20);
            this.cmb_gridSZ.Items.Add(ribbonDropDownItemImpl21);
            this.cmb_gridSZ.Items.Add(ribbonDropDownItemImpl22);
            this.cmb_gridSZ.Items.Add(ribbonDropDownItemImpl23);
            this.cmb_gridSZ.Label = "Grid Size";
            this.cmb_gridSZ.Name = "cmb_gridSZ";
            this.cmb_gridSZ.ButtonClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cmb_gridSZ_Click);
            this.cmb_gridSZ.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cmb_gridSZ_SelectionChanged);
            // 
            // group2
            // 
            this.group2.Items.Add(this.gallery_comb);
            this.group2.Items.Add(this.gallery_seq);
            this.group2.Items.Add(this.separator1);
            this.group2.Label = "In-build";
            this.group2.Name = "group2";
            // 
            // gallery_comb
            // 
            this.gallery_comb.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.gallery_comb.Image = global::PPTAddin.Properties.Resources.com_AND;
            ribbonDropDownItemImpl24.Image = global::PPTAddin.Properties.Resources.com_AND;
            ribbonDropDownItemImpl24.Label = "And";
            ribbonDropDownItemImpl25.Image = global::PPTAddin.Properties.Resources.com_BUFFER;
            ribbonDropDownItemImpl25.Label = "Buffer";
            ribbonDropDownItemImpl26.Image = global::PPTAddin.Properties.Resources.com_NAND;
            ribbonDropDownItemImpl26.Label = "Nand";
            ribbonDropDownItemImpl27.Image = global::PPTAddin.Properties.Resources.com_NOR;
            ribbonDropDownItemImpl27.Label = "Nor";
            ribbonDropDownItemImpl28.Image = global::PPTAddin.Properties.Resources.com_NOT;
            ribbonDropDownItemImpl28.Label = "Not";
            ribbonDropDownItemImpl29.Image = global::PPTAddin.Properties.Resources.com_OR;
            ribbonDropDownItemImpl29.Label = "Or";
            ribbonDropDownItemImpl30.Image = global::PPTAddin.Properties.Resources.com_XNOR;
            ribbonDropDownItemImpl30.Label = "Xnor";
            ribbonDropDownItemImpl31.Image = global::PPTAddin.Properties.Resources.com_XOR;
            ribbonDropDownItemImpl31.Label = "Xor";
            this.gallery_comb.Items.Add(ribbonDropDownItemImpl24);
            this.gallery_comb.Items.Add(ribbonDropDownItemImpl25);
            this.gallery_comb.Items.Add(ribbonDropDownItemImpl26);
            this.gallery_comb.Items.Add(ribbonDropDownItemImpl27);
            this.gallery_comb.Items.Add(ribbonDropDownItemImpl28);
            this.gallery_comb.Items.Add(ribbonDropDownItemImpl29);
            this.gallery_comb.Items.Add(ribbonDropDownItemImpl30);
            this.gallery_comb.Items.Add(ribbonDropDownItemImpl31);
            this.gallery_comb.Label = "Combinational";
            this.gallery_comb.Name = "gallery_comb";
            this.gallery_comb.ShowImage = true;
            this.gallery_comb.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.gallery_comb_Click);
            // 
            // gallery_seq
            // 
            this.gallery_seq.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.gallery_seq.Image = global::PPTAddin.Properties.Resources.seq_D_flop;
            ribbonDropDownItemImpl32.Image = global::PPTAddin.Properties.Resources.seq_D_flop;
            ribbonDropDownItemImpl32.Label = "DFlop";
            ribbonDropDownItemImpl33.Image = global::PPTAddin.Properties.Resources.seq_latch;
            ribbonDropDownItemImpl33.Label = "Latch";
            ribbonDropDownItemImpl34.Image = global::PPTAddin.Properties.Resources.seq_synchronizer;
            ribbonDropDownItemImpl34.Label = "Synchronizer";
            this.gallery_seq.Items.Add(ribbonDropDownItemImpl32);
            this.gallery_seq.Items.Add(ribbonDropDownItemImpl33);
            this.gallery_seq.Items.Add(ribbonDropDownItemImpl34);
            this.gallery_seq.Label = "Sequential";
            this.gallery_seq.Name = "gallery_seq";
            this.gallery_seq.ShowImage = true;
            this.gallery_seq.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.gallery_seq_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // group3
            // 
            this.group3.Items.Add(this.btn_export);
            this.group3.Items.Add(this.btn_lib1);
            this.group3.Items.Add(this.btn_lib2);
            this.group3.Label = "User library";
            this.group3.Name = "group3";
            // 
            // btn_export
            // 
            this.btn_export.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_export.Image = global::PPTAddin.Properties.Resources.export;
            this.btn_export.Label = "Export";
            this.btn_export.Name = "btn_export";
            this.btn_export.ShowImage = true;
            this.btn_export.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_export_Click);
            // 
            // btn_lib1
            // 
            this.btn_lib1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_lib1.Image = global::PPTAddin.Properties.Resources.library1;
            this.btn_lib1.Label = "Library1";
            this.btn_lib1.Name = "btn_lib1";
            this.btn_lib1.ShowImage = true;
            this.btn_lib1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_lib1_Click);
            // 
            // btn_lib2
            // 
            this.btn_lib2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_lib2.Image = global::PPTAddin.Properties.Resources.library2;
            this.btn_lib2.Label = "Library2";
            this.btn_lib2.Name = "btn_lib2";
            this.btn_lib2.ShowImage = true;
            this.btn_lib2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_lib2_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.confirm_chk);
            this.group4.Items.Add(this.drop_additional);
            this.group4.Label = "Additional Buttons";
            this.group4.Name = "group4";
            // 
            // confirm_chk
            // 
            this.confirm_chk.Label = "Confirm";
            this.confirm_chk.Name = "confirm_chk";
            this.confirm_chk.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.confirm_chk_Click);
            // 
            // drop_additional
            // 
            ribbonDropDownItemImpl35.Label = "Option 1";
            ribbonDropDownItemImpl36.Label = "Option 2";
            ribbonDropDownItemImpl37.Label = "Option 3";
            this.drop_additional.Items.Add(ribbonDropDownItemImpl35);
            this.drop_additional.Items.Add(ribbonDropDownItemImpl36);
            this.drop_additional.Items.Add(ribbonDropDownItemImpl37);
            this.drop_additional.Label = "Options";
            this.drop_additional.Name = "drop_additional";
            this.drop_additional.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.drop_additional_SelectionChanged);
            // 
            // group_tool
            // 
            this.group_tool.Items.Add(this.wv_regChk);
            this.group_tool.Items.Add(this.polyln_btn);
            this.group_tool.Label = "Tools";
            this.group_tool.Name = "group_tool";
            // 
            // wv_regChk
            // 
            this.wv_regChk.Label = "Regular Mode";
            this.wv_regChk.Name = "wv_regChk";
            this.wv_regChk.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.wv_regChk_Click);
            // 
            // polyln_btn
            // 
            this.polyln_btn.Description = "WaveForm";
            this.polyln_btn.Image = global::PPTAddin.Properties.Resources.polyln;
            this.polyln_btn.Label = " WaveForm";
            this.polyln_btn.Name = "polyln_btn";
            this.polyln_btn.ShowImage = true;
            this.polyln_btn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.polyln_btn_Click);
            // 
            // Settings
            // 
            this.Settings.Items.Add(this.btn_setting);
            this.Settings.Items.Add(this.separator2);
            this.Settings.Items.Add(this.editBox1);
            this.Settings.Items.Add(this.editBox2);
            this.Settings.Items.Add(this.checkBox1);
            this.Settings.Items.Add(this.rename_btn);
            this.Settings.Label = "Setting";
            this.Settings.Name = "Settings";
            // 
            // btn_setting
            // 
            this.btn_setting.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_setting.Image = global::PPTAddin.Properties.Resources.settings;
            this.btn_setting.Label = "ShortKey Setting";
            this.btn_setting.Name = "btn_setting";
            this.btn_setting.ShowImage = true;
            this.btn_setting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_setting_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // editBox1
            // 
            this.editBox1.Enabled = false;
            this.editBox1.Label = "Library1";
            this.editBox1.Name = "editBox1";
            this.editBox1.Text = null;
            this.editBox1.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBox1_TextChanged);
            // 
            // editBox2
            // 
            this.editBox2.Enabled = false;
            this.editBox2.Label = "Library2";
            this.editBox2.Name = "editBox2";
            this.editBox2.Text = null;
            this.editBox2.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBox2_TextChanged);
            // 
            // checkBox1
            // 
            this.checkBox1.Label = "Rename Libs";
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBox1_Click);
            // 
            // rename_btn
            // 
            this.rename_btn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.rename_btn.Enabled = false;
            this.rename_btn.Image = global::PPTAddin.Properties.Resources.rename;
            this.rename_btn.Label = "Library Rename";
            this.rename_btn.Name = "rename_btn";
            this.rename_btn.ShowImage = true;
            this.rename_btn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rename_btn_Click);
            // 
            // point
            // 
            this.Name = "point";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Tag = "Point";
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.point_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group_tool.ResumeLayout(false);
            this.group_tool.PerformLayout();
            this.Settings.ResumeLayout(false);
            this.Settings.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        public Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal RibbonGroup group2;
        internal RibbonSeparator separator1;
        internal RibbonGroup group3;
        internal RibbonButton btn_lib1;
        internal RibbonSeparator separator2;
        internal RibbonButton btn_lib2;
        internal RibbonGroup group4;
        internal RibbonCheckBox confirm_chk;
        internal RibbonDropDown drop_additional;
        internal RibbonButton btn_export;
        internal RibbonGallery gallery_comb;
        internal RibbonGallery gallery_seq;
        internal RibbonGroup Settings;
        internal RibbonButton btn_setting;
        internal RibbonButton rename_btn;
        internal RibbonEditBox editBox1;
        internal RibbonEditBox editBox2;
        internal RibbonCheckBox checkBox1;
        internal RibbonGroup group_tool;
        internal RibbonButton polyln_btn;
        internal RibbonCheckBox wv_regChk;
        internal RibbonGroup group1;
        internal RibbonLabel label1;
        internal RibbonLabel lblX;
        internal RibbonLabel lblY;
        internal RibbonLabel lblState;
        internal RibbonLabel xpos;
        internal RibbonLabel ypos;
        internal RibbonButton btn_open;
        internal RibbonButton button1;
        internal RibbonButton btn_Outport;
        internal RibbonButton btn_Inport;
        internal RibbonButton button2;
        internal RibbonButton btn_grid;
        internal RibbonDropDown cmb_pinSZ;
        internal RibbonDropDown cmb_portSZ;
        internal RibbonDropDown cmb_gridSZ;
    }

    partial class ThisRibbonCollection
    {
        internal point point
        {
            get { return this.GetRibbon<point>(); }
        }
    }
}
