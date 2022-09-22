using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.ComponentModel;
using Microsoft.Office.Core;
using System.Windows.Forms;
using System.Diagnostics;
using Aspose;
using System.Drawing;
using Microsoft.Office.Tools.Ribbon;
using System.Threading;
using System.Threading.Tasks;
using stringrep;
using Utilities;
using Microsoft.Win32;
using System.Data;
using Gma.System.MouseKeyHook;
using System.Runtime.InteropServices;


namespace PPTAddin
{
    public partial class ThisAddIn
    {
        //globalKeyboardHook gkh = new globalKeyboardHook();
        #region general variable
        int isdragging;     // this stores a mouse dragging state
        bool isToolclicked;     //this stores a tool state
        public static IKeyboardMouseEvents m_Events;
        int nAllshpCnt;
        public bool isDragging = false;
        public static bool POL;             // point or line

        public PowerPoint.Shape mouseWaveLine = null;
        public PowerPoint.Shape mouseWaveCursor = null;

        public uint state;           // current button state
        public uint State
        {
            get { return state; }
            set
            {
                if (state == value)
                    return;
                //wave
                if (wavePtCount > 1)
                {
                    PowerPoint.Shape shp = AddWaveShape(null);
                    if (shp != null)
                        MdbManger.GetInstance().Modify(shp, "add");
                }
                wavePtCount = 0;
                waveLst.Clear();
                if (mouseWaveLine != null)
                    mouseWaveLine.Delete();
                mouseWaveLine = null;
                if (mouseWaveCursor != null)
                    mouseWaveCursor.Delete();
                mouseWaveCursor = null;
                

                state = value;
                if (state == 2)
                {
                    firstshp = null;
                    secondshp = null;
                }
            }
        }


        public static bool binit;
        public static string path;
        public static string userlibpath;
        public static int lineCnt = 0;
        //public static List<Point> drawLines = new List<Point>();
        public static List<Point> drawPoints = new List<Point>();
        public static List<Point> allCurDPts = new List<Point>();
        
        /// remove ShpHistory PtHistory [
        //         public static List<Point> PtHistory = new List<Point>();
        //         public static List<PowerPoint.Shape> ShpHistory = new List<PowerPoint.Shape>();
        //]

        bool bfirst = true;
        public static int linestartx, linestarty, curXpos, curYpos;
        //public int ribbonHeight = -35;
        public int ribbonHeight = 0;

        Point mousepos = Point.Empty, drag1pt, drag2pt;

        public System.Drawing.Point Origin;
        public Point wavelnBeginpt;
        public static Point tmppt = new Point(-100, -100);
        public static bool bkeypress = false;       //keyboard hook relation
        public static int nCtrlOShift = 0;          //keyboard hook relation
        public static bool bdelKeypress = false;

        public static int nGridSpace = 4;       //Grid relation
        public static bool bGridcheck;
        public static List<PowerPoint.Shape> GridlistX = new List<PowerPoint.Shape>();
        public static List<PowerPoint.Shape> GridlistY = new List<PowerPoint.Shape>();      //Grid relation

        public static string nCombitionalshp = "";
        public static int nSeqshp = -1;
        // public int connectorCnt = 0;

        public static bool bregMode;
        public static string strlib1name, strlib2name;
        bool isfirstlinePoint = false;


        public static point pointRibbon = null;


        public enum shapestates
        {
            None,
            Box,
            Line
        }

        [Flags]
        public enum ShapeTypeFlag
        {
            NONE = 0,
            PIN = 1,
            INPORT = 2,
            OUTPORT = 4
        }
        public struct POINT
        {
            public int X;
            public int Y;

            public POINT(int x, int y)
            {
                this.X = x;
                this.Y = y;
            }
        }
        POINT tmppoint;

        public GeneralHook hook = new GeneralHook();//Total hook control   
        #endregion

        #region object variable
        public static Microsoft.Office.Interop.PowerPoint.Presentation newpres;
        Microsoft.Office.Interop.PowerPoint.CustomLayout customLayout;

        Microsoft.Office.Interop.PowerPoint.Slides slides;
        public static Microsoft.Office.Interop.PowerPoint._Slide slide;
        private static Microsoft.Office.Interop.PowerPoint.Shape shp, firstshp, secondshp, bufshp = null;
        public Microsoft.Office.Interop.PowerPoint.Shape BufShp
        {
            get { return bufshp; }
            set
            {
                if (value == null)
                {

                }
                else
                {

                }
            }
        }

        Microsoft.Office.Interop.PowerPoint.Shape Sconnector;
        PowerPoint.Shape curselShp;
        PowerPoint.Shape curWaveln;
        #endregion
        public void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Class1 str = new Class1();
            //str.VCrypt();
            isDragging = false;
            nAllshpCnt = 0;
            //RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\PPTaddin");
            //if ((key == null) || (key.GetValue("Set1").ToString() != ("addin")))
            //{
            //    LicenseDlg dlg = new LicenseDlg();
            //    if (dlg.ShowDialog() == DialogResult.OK)
            //    {
            //        if (dlg.m_bPass != true)
            //        {
            //            MessageBox.Show("Wrong License.Please retry.");
            //            //Application.Exit();
            //            ThisAddIn_Shutdown(sender, e);
            //        }
            //    }
            //    else
            //    {
            //        //Application.Exit();
            //        ThisAddIn_Shutdown(sender, e);
            //    }
            //}
            //else
            //{
            //key.Close();

            Origin = new Point();
            State = (uint)shapestates.None;

            this.Application.SlideShowBegin +=
                new PowerPoint.EApplication_SlideShowBeginEventHandler(Application_SlideShow);
            this.Application.PresentationNewSlide +=
                    new PowerPoint.EApplication_PresentationNewSlideEventHandler(Application_PresentationNewSlide);
            this.Application.AfterNewPresentation +=
                new Microsoft.Office.Interop.PowerPoint.EApplication_AfterNewPresentationEventHandler(NewPresentation);
            this.Application.PresentationBeforeClose +=
                new PowerPoint.EApplication_PresentationBeforeCloseEventHandler(Application_PresentationBeforeClose);
            binit = false;
            firstshp = null;
            secondshp = null;
            bkeypress = false;
            bGridcheck = false;
            isdragging = 0;
            SetMaxScreeen_TaskBar();
            var frm = new ShortkeySel();
            frm.Show();


            this.Application.WindowSelectionChange += OnWindowSelectionChanged;

            // 
            //             PowerPoint.Presentation pptPresentation = this.Application.Presentations.Add(MsoTriState.msoTrue);
            //             PowerPoint.CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];
            //             // Create new Slide
            //             slides = pptPresentation.Slides;
            //             slide = slides.AddSlide(1, customLayout);
            //GetAllShapes();            
            //}
        }
        public void BeginHook()
        {
            KeyboardHooking.SetHook();
            SubscribeApplication();
            if (Globals.ThisAddIn.Application.ActiveWindow.WindowState != PowerPoint.PpWindowState.ppWindowMaximized)
            {
                MessageBox.Show("Please Maxmize Window", "Maxmize Window", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                return;
            }

        }
        void OnAfterDragDropOnSlide()
        {
            MessageBox.Show("Drag");
        }
        void SetMaxScreeen_TaskBar()
        {
            System.Drawing.Rectangle rect = new Rectangle(0, 0, 0, 0);

            System.Windows.Forms.Screen scr = System.Windows.Forms.Screen.PrimaryScreen;
            rect = scr.Bounds;
            int w = rect.Width;
            int h = rect.Height;
            List<Rectangle> resolist = new List<Rectangle>();
            resolist = ScreenResolution.GetScreenresol();
            ScreenResolution.SetScreenresol(resolist.Last().Width, resolist.Last().Height);

            //Taskbar.Hide();
            Taskbar.Show();
        }
        void OnWindowSelectionChanged(PowerPoint.Selection Sel)
        {
            try
            {
                if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    //                 Boolean bDlgOk = false;
                    //                 if (bdelKeypress)
                    //                 {
                    //                     DialogResult result = MessageBox.Show("Do you want to delete component circuit?", "Confirmation", MessageBoxButtons.OKCancel);
                    //                     if (result == DialogResult.OK)
                    //                     {
                    //                         bDlgOk = true;
                    //                     }
                    //                     else
                    //                     {
                    //                         //Globals.ThisAddIn.Application.CommandBars.ExecuteMso("Undo");
                    //                     }
                    //                 }
                    foreach (PowerPoint.Shape selectedShape in Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange)
                    {
                        if (selectedShape.Tags["sel"] == "unsel")
                        {
                            //selectedShape.Select(MsoTriState.msoFalse);
                            this.Application.ActiveWindow.Selection.Unselect();
                            continue;
                        }
                    }
                    //  get selected shape in group
                    if (Sel.HasChildShapeRange == true)
                    {
                        foreach (PowerPoint.Shape selectedShape in Sel.ChildShapeRange)
                        {
                            if (selectedShape.Tags["sel"] == "unsel")
                            {
                                //selectedShape.Select(MsoTriState.msoFalse);
                                this.Application.ActiveWindow.Selection.Unselect();
                                continue;
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine("-----[OnWindowSelectionChanged] exception {0}", e);
            }

        }
        public static Boolean bShapeDel = true;
        public static void OnShapeDel()
        {
            if (!bShapeDel) return;
            bShapeDel = false;
            try
            {
                PowerPoint.Selection Sel;
                Sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.ShapeRange shapeRange = Sel.ShapeRange;

                    Boolean bDlgOk = false;
                    for (int sel = shapeRange.Count; sel > 0; sel--)
                    {
                        var selectedShape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[sel];
                        if (selectedShape.AutoShapeType != MsoAutoShapeType.msoShapeNotPrimitive)
                        {
                            if (sel == shapeRange.Count)
                            {
                                DialogResult result = MessageBox.Show("Do you want to delete component circuit?", "Confirmation", MessageBoxButtons.OKCancel);
                                if (result == DialogResult.OK)
                                    bDlgOk = true;
                                else
                                {
                                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso("Undo");
                                    break;
                                }
                            }
                            //selectedShape.Delete();

                        }
                    }
                    if (bDlgOk)
                    {
                        MdbManger.GetInstance().Modify(null, "del");
                    }
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine("{0} shapeDel exception", e);
            }
            bShapeDel = true;
        }
        public static void ExitApp()
        {
            Microsoft.Office.Interop.PowerPoint.Presentation delpres = Globals.ThisAddIn.Application.ActivePresentation;
            delpres.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(delpres);
            delpres = null;
        }
        public void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            hook.UnInstallHook(HookHelper.HookType.KeyOperation);//Uninstall keyboard hook
            hook.UnInstallHook(HookHelper.HookType.MouseOperation);//Uninstall mouse hook
            KeyboardHooking.ReleaseHook();
            this.Application.WindowSelectionChange -= OnWindowSelectionChanged;
            Unsubscribe();
        }
        public void Application_SlideShow(PowerPoint.SlideShowWindow wn)
        {
            slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

        }

        private static bool bFirstSlide = true;
        public void Application_PresentationNewSlide(PowerPoint.Slide Sld)
        {
            slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            if (bFirstSlide)
            {
                while (slide.Shapes.Count > 0)
                {
                    slide.Shapes[slide.Shapes.Count].Delete();
                }
                //             foreach (PowerPoint.Shape Shpe in slide.Shapes)
                //             {
                //                 Shpe.Delete();
                //             }
                MdbManger.GetInstance().LoadMDB();
            }
            //Globals.ThisAddIn.Application.DisplayGridLines = MsoTriState.msoTrue;

        }
        public void Application_PresentationBeforeClose(PowerPoint.Presentation Pres, ref bool Cancel)
        {
            MdbManger.GetInstance().SaveMDB();
        }
        public void NewPresentation(Microsoft.Office.Interop.PowerPoint.Presentation oPres)
        {
            //             PowerPoint.CustomLayout customLayout = oPres.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];
            // 
            //             // Create new Slide
            //             slides = oPres.Slides;
            //             slide = slides.AddSlide(1, customLayout);

            BeginHook();

        }
        private Boolean CheckMouseOverPointShape(Microsoft.Office.Interop.PowerPoint.Shape shape, Point pt)
        {
            if (shape.Tags.Count == 0) return false;
            if (shape.Tags["xPoint"] != "")
            {
                float rot = shape.Rotation;
                if (rot != 0)
                    shape.Rotation = 0;
                if (shape.Left <= pt.X && pt.X <= shape.Left + shape.Width
                        && shape.Top <= pt.Y && pt.Y <= shape.Top + shape.Height)
                {
                    bufshp = shape;
                    float orgWidth = bufshp.Width;
                    bufshp.Left -= orgWidth / 2;
                    bufshp.Top -= orgWidth / 2;
                    bufshp.Width = orgWidth * 2;
                    bufshp.Height = orgWidth * 2;
                    if (rot != 0)
                        shape.Rotation = rot;

                    bufshp.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                    return true;
                }
                if (rot != 0)
                    shape.Rotation = rot;
            }
            return false;
        }
        private int checkMoveLineable(Point pt)
        {
            try
            {
                if (pt.X < 0 || pt.Y < 0)
                    return -1;

                if (bufshp != null)
                {
                    float rot = bufshp.Rotation;
                    if (rot != 0)
                        bufshp.Rotation = 0;
                    if (bufshp.Left <= pt.X && pt.X <= bufshp.Left + bufshp.Width
                        && bufshp.Top <= pt.Y && pt.Y <= bufshp.Top + bufshp.Height)
                    {
                        if (rot != 0)
                            bufshp.Rotation = rot;
                        return 0;
                    }
                    else
                    {
                        float orgWidth = bufshp.Width / 2;
                        bufshp.Left += orgWidth / 2;
                        bufshp.Top += orgWidth / 2;
                        bufshp.Width = orgWidth;
                        bufshp.Height = orgWidth;
                        bufshp.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(0, 255, 0).ToArgb();
                        if (rot != 0)
                            bufshp.Rotation = rot;
                        bufshp = null;
                    }
                }

                Microsoft.Office.Interop.PowerPoint._Slide curSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in curSlide.Shapes)
                {
                    if (shape.Type == Office.MsoShapeType.msoGroup && shape.GroupItems.Count > 0)
                    {
                        foreach (Microsoft.Office.Interop.PowerPoint.Shape subShape in shape.GroupItems)
                        {
                            if (CheckMouseOverPointShape(subShape, pt))
                                return 0;
                        }

                    }
                    if (CheckMouseOverPointShape(shape, pt))
                        return 0;

                }
                return -1;
                /// remove ShpHistory PtHistory [
                //                 for (int i = 0; i < PtHistory.Count; i++)
                //                 {
                //                     if ((PtHistory[i].X - ptwidth <= pt.X) && (pt.X <= PtHistory[i].X + ptwidth) && (PtHistory[i].Y - ptwidth <= pt.Y) && (pt.Y <= PtHistory[i].Y + ptwidth))
                //                     {
                //                         globalPointposition = PtHistory[i];
                //                         bufshp = ShpHistory[i];
                //                         bufshp.Left -= ptwidth / 2;
                //                         bufshp.Top -= ptwidth / 2;
                //                         bufshp.Width = ptwidth * 2;
                //                         bufshp.Height = ptwidth * 2;
                //                         bufshp.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                //                         return i;
                //                     }
                //                 }
                //]
            }
            catch (Exception e)
            {
                Debug.WriteLine("{0} checkMoveLineable Exception", e);
            }

            return -1;
        }
        private bool checkLineable(Point pt)        // this returns point mouse selected is in PointList or not.
        {
            bool bpoint = false;
            if (bufshp != null)
            {
                bpoint = true;
                bufshp.Fill.ForeColor.RGB = Color.Red.ToArgb();
            }
            //             for (int i = 0; i < PtHistory.Count; i++)
            //             {
            //                 if ((PtHistory[i].X- ptwidth <= pt.X) && (pt.X <= PtHistory[i].X + ptwidth) && (PtHistory[i].Y- ptwidth <= pt.Y) && (pt.Y <= PtHistory[i].Y + ptwidth))
            //                 {
            //                     bpoint = true;
            //                     ShpHistory[i].Width = ptwidth * 2;
            //                     ShpHistory[i].Height = ptwidth * 2;
            //                     ShpHistory[i].Fill.ForeColor.RGB = Color.Red.ToArgb();
            //                     bufshp = ShpHistory[i];
            //                 }
            //             }
            return bpoint;
        }
        public void DrawConnector()
        {
            if (bufshp == null)
                return;
            if (firstshp == null)
            {
                firstshp = bufshp;
                return;
            }
            else
            {
                if (firstshp == bufshp)
                    return;
                secondshp = bufshp;
                PowerPoint.Shape shp = AddWireShape(new Wire_Param(firstshp, secondshp));
                if (shp != null)
                    MdbManger.GetInstance().Modify(shp, "add");
                firstshp = null;
                secondshp = null;
            }
        }

        public PowerPoint.Shape GetShapeByName(string name)
        {
            if (name == "")
                return null;
            foreach (Microsoft.Office.Interop.PowerPoint.Slide slide in Globals.ThisAddIn.Application.ActivePresentation.Slides)
            {
                foreach (PowerPoint.Shape shp in slide.Shapes)
                {
                    if (shp.Name.CompareTo(name) == 0)
                    {
                        return shp;
                    }
                    if (shp.Type == Microsoft.Office.Core.MsoShapeType.msoGroup && shp.GroupItems.Count > 0)
                    {
                        foreach (Microsoft.Office.Interop.PowerPoint.Shape subShape in shp.GroupItems)
                        {
                            if (subShape.Name.CompareTo(name) == 0)
                            {
                                return subShape;
                            }
                        }
                        continue;
                    }

                }
            }
            return null;
        }
        

        public PowerPoint.Shape AddWireShape(Wire_Param param)
        {
            PowerPoint.Shape shape = null;

            try
            {
                if (param.beginShp == null && param.strBeginPt != "")
                    param.beginShp = GetShapeByName(param.strBeginPt);
                if (param.endShp == null && param.strEndPt != "")
                    param.endShp = GetShapeByName(param.strEndPt);

                if (param.beginShp != null && param.endShp != null && param.beginShp != param.endShp)
                {
                    var slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                    shape = slide.Shapes.AddConnector(MsoConnectorType.msoConnectorStraight, 1, 1, 1, 1);
                    shape.ConnectorFormat.BeginConnect(param.beginShp, 1);
                    shape.ConnectorFormat.EndConnect(param.endShp, 1);
                    shape.Line.Weight = 2f;
                    shape.RerouteConnections();

                    shape.Tags.Add("id", param.id.ToString());
                    shape.Tags.Add("kind", param.kind);
                    shape.Name = param.name;
                    shape.Tags.Add("pt1_name", param.strBeginPt);
                    shape.Tags.Add("pt2_name", param.strEndPt);
                    return shape;
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine("-----[AddWireShape] Exception {0}", e);
            }
            return shape;
        }

//         PowerPoint.Shape shape1 = null;
//         PowerPoint.Shape shape2 = null;
        List<PowerPoint.Shape> waveLst = new List<PowerPoint.Shape>();
        private int wavePtCount = 0;
        //        PowerPoint.GroupShape gs;
        public void DrawWaveLine(Point sldpt)
        {
            if (wavePtCount > 0)
            {
                PowerPoint.Shape waveline = slide.Shapes.AddLine(wavelnBeginpt.X, wavelnBeginpt.Y, sldpt.X, sldpt.Y);
                int lineMode = 0;
                if(wavelnBeginpt.X <= sldpt.X)
                {
                    lineMode = wavelnBeginpt.Y <= sldpt.Y ? 1: 2;
                }
                else
                {
                    lineMode = wavelnBeginpt.Y <= sldpt.Y ? 3 : 4;
                }
                waveline.Tags.Add("kind", "waveline");
                waveline.Tags.Add("sel", "unsel");
                waveline.Tags.Add("linemode", lineMode.ToString());
                waveline.GetType().InvokeMember("Locked", System.Reflection.BindingFlags.SetProperty, null, waveline, new object[] { MsoTriState.msoTrue });
                waveLst.Add(waveline);
            }
            wavelnBeginpt = sldpt;
            wavePtCount++;

        }
        public PowerPoint.Shape AddWaveShape(Wave_Param param)
        {
            try
            {
                if (param != null)
                {
                    waveLst.Clear();
                    wavePtCount = 0;
                    string[] str_pts = param.str_pts.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 0; i < str_pts.Length; i += 2)
                    {
                        Point sldPt = new Point((int)float.Parse(str_pts[i]), (int)float.Parse(str_pts[i + 1]));
                        DrawWaveLine(sldPt);
                    }
               }
                
                if (waveLst.Count > 0)
                {
                    string[] myRangeArray = new string[waveLst.Count];
                    int i = 0;
                    foreach (PowerPoint.Shape shp in waveLst)
                    {
                        shp.GetType().InvokeMember("Locked", System.Reflection.BindingFlags.SetProperty, null, shp, new object[] { MsoTriState.msoFalse });
                        myRangeArray[i] = shp.Name;
                        i++;
                        //shp.Connector.
//                         shp.ConnectorFormat.
//                         shape = slide.Shapes.AddConnector(MsoConnectorType.msoConnectorStraight, 1, 1, 1, 1);
//                         shape.ConnectorFormat.BeginConnect(param.beginShp, 1);
//                         shape.ConnectorFormat.EndConnect(param.endShp, 1);
                    }

                    PowerPoint.Shape waveShape = slide.Shapes.Range(myRangeArray).Group();
                    if (param != null)
                    {
                        waveShape.Tags.Add("id", param.id.ToString());
                        waveShape.Tags.Add("kind", param.kind);
                        waveShape.Name = param.name;
                    }
                    else
                        Wave_Param.SetShapeTags(waveShape);
                    waveLst.Clear();
                    wavePtCount = 0;
                    return waveShape;
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine("----[EndEditWave] exception{0}", e);
            }
            return null;
        }
        public void GetAllShapeCnt()
        {
            //newpres = Globals.ThisAddIn.Application.ActivePresentation;            
            //PtHistory.Clear();
            List<PowerPoint.Shape> shplist = new List<PowerPoint.Shape>();
            Point sldpt = Point.Empty;
            //shplist = slide.Shapes;
            nAllshpCnt = 0;
            foreach (PowerPoint.Shape ptshp in slide.Shapes)
            {
                if (ptshp.AutoShapeType != MsoAutoShapeType.msoShapeNotPrimitive)
                    nAllshpCnt++;
            }
        }
        public static Point ScreenPointToSlidePoint(Point point)
        {
            // Get the size of the slide (in points of the slide's coordinate system).
            int slideWidth = 0;
            int slideHeight = 0;
            try
            {
                var slide1 = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                slideWidth = (int)slide1.CustomLayout.Width;
                slideHeight = (int)slide1.CustomLayout.Height;
                // Get the screen coordinates of the upper-left hand corner of the slide.
                Point topLeft = new Point(
                    Globals.ThisAddIn.Application.ActiveWindow.PointsToScreenPixelsX(0),
                    Globals.ThisAddIn.Application.ActiveWindow.PointsToScreenPixelsY(0));
                Point bottomRight = new Point(
                    Globals.ThisAddIn.Application.ActiveWindow.PointsToScreenPixelsX(slideWidth),
                    Globals.ThisAddIn.Application.ActiveWindow.PointsToScreenPixelsY(slideHeight));
                return new Point(
                    (point.X - topLeft.X) * slideWidth / (bottomRight.X - topLeft.X),
                    (point.Y - topLeft.Y) * slideHeight / (bottomRight.Y - topLeft.Y));
            }
            catch (Exception e)
            {
                Debug.WriteLine("{0} slide get exception", e);
                return new Point(-1, -1);
            }
        }
        public static POINT ConvertScreenPointToSlideCoordinates(POINT point)
        {
            // Get the screen coordinates of the upper-left hand corner of the slide.
            POINT refPointUpperLeft = new POINT(Globals.ThisAddIn.Application.ActiveWindow.PointsToScreenPixelsX(0)
                , Globals.ThisAddIn.Application.ActiveWindow.PointsToScreenPixelsY(0));

            // Get the size of the slide (in points of the slide's coordinate system).
            var slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            var slideWidth = slide.CustomLayout.Width;
            var slideHeight = slide.CustomLayout.Height;

            Rectangle rect = new Rectangle(0, 0, 0, 0);
            Rectangle Wrect = new Rectangle(0, 0, 0, 0);
            Win32API.GetClientRect(new IntPtr(Globals.ThisAddIn.Application.ActiveWindow.HWND), ref rect);
            Win32API.GetWindowRect(new IntPtr(Globals.ThisAddIn.Application.ActiveWindow.HWND), ref Wrect);
            POINT ribheight = new POINT(0, 0);
            ribheight.X = 0;
            ribheight.Y = Wrect.Top - rect.Top;
            POINT Sheight = new POINT(0, 0);
            Sheight.Y = rect.Height;
            // Get the screen coordinates of the bottom-right hand corner of the slide.
            POINT refPointBottomRight = new POINT(
                Globals.ThisAddIn.Application.ActiveWindow.PointsToScreenPixelsX(slideWidth),
                Globals.ThisAddIn.Application.ActiveWindow.PointsToScreenPixelsY(slideHeight));
            //Win32API.POINT refPointUpperLeft1 = refPointUpperLeft;
            // Both reference points have to be converted to the PowerPoint window's coordinate system.
            Win32API.ScreenToClient(new IntPtr(Globals.ThisAddIn.Application.ActiveWindow.HWND), ref refPointUpperLeft);
            Win32API.ScreenToClient(new IntPtr(Globals.ThisAddIn.Application.ActiveWindow.HWND), ref refPointBottomRight);
            Win32API.ScreenToClient(new IntPtr(Globals.ThisAddIn.Application.ActiveWindow.HWND), ref ribheight);
            Win32API.ScreenToClient(new IntPtr(Globals.ThisAddIn.Application.ActiveWindow.HWND), ref Sheight);
            // Convert the point to the slide's coordinate system (convert the pixel coordinate inside the slide into a 0..1 interval, then scale it up by the slide's point size).
            return new POINT(
                (int)Math.Round((double)(point.X - refPointUpperLeft.X) / (refPointBottomRight.X - refPointUpperLeft.X) * slideWidth),
                /*(int)Math.Round((double)(point.Y - refPointUpperLeft.Y-Math.Abs( ribheight.Y)-1-Math.Abs(slideHeight-Sheight.Y)/2) / (refPointBottomRight.Y - refPointUpperLeft.Y) * slideHeight)*/
                (int)Math.Round((double)(point.Y - refPointUpperLeft.Y) / (refPointBottomRight.Y - refPointUpperLeft.Y) * slideHeight));
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        public void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
        public List<PowerPoint.Shape> GetshapesSlide()
        {
            List<PowerPoint.Shape> shplist = new List<PowerPoint.Shape>();
            //shplist = slide.Shapes;
            foreach (PowerPoint.Shape ptshp in slide.Shapes)
                shplist.Add(ptshp);
            return shplist;
        }
        public void makeCorrectPosition(List<PowerPoint.Shape> shplist, PowerPoint.Shape newShp)
        {
            foreach (PowerPoint.Shape inst in shplist)
            {
                if ((inst.Left > newShp.Left - Point_Param.pinWidth * 2) && (inst.Left < newShp.Left + Point_Param.pinWidth * 2) && (inst.Top < newShp.Top) && (inst.Top + inst.Height > newShp.Top + newShp.Height))
                //((inst.Left < shp.Left) && (inst.Left + inst.Width > shp.Left + shp.Left + ptwidth) && (inst.Top + inst.Height > shp.Top + ptwidth) && (inst.Top + inst.Height < shp.Top + ptwidth * 2))||  
                //((inst.Left+inst.Width>shp.Left-ptwidth)&&(inst.Left+inst.Width<shp.Left+ptwidth*2)&&(inst.Top<shp.Top)&&(inst.Top+inst.Height>shp.Top+ptwidth))||
                //((inst.Left<shp.Left)&&(inst.Left+inst.Width>shp.Left+ptwidth)&&(inst.Top>shp.Top-ptwidth)&&(inst.Top<shp.Top+ptwidth*2)))
                {
                    if (!GridlistX.Contains(inst) && !GridlistY.Contains(inst))
                        newShp.Left = inst.Left - Point_Param.pinWidth / 2;
                    ////inst.Name = "shape1";
                    ////shp.Name = "shape2";
                    ////string[] myRangeArray = new string[2];
                    ////myRangeArray[0] = "shape1";
                    ////myRangeArray[1] = "shape2";
                    ////slide.Shapes.Range(myRangeArray).Group();
                    break;
                }
                if ((inst.Left < newShp.Left) && (inst.Left + inst.Width > newShp.Left + Point_Param.pinWidth) && (inst.Top + inst.Height > newShp.Top - Point_Param.pinWidth * 2) && (inst.Top + inst.Height < newShp.Top + Point_Param.pinWidth * 2))
                {
                    if (!GridlistX.Contains(inst) && !GridlistY.Contains(inst))
                        newShp.Top = inst.Top + inst.Height - Point_Param.pinWidth / 2;
                    //inst.Name = "shape1";
                    //shp.Name = "shape2";
                    //string[] myRangeArray = new string[2];
                    //myRangeArray[0] = "shape1";
                    //myRangeArray[1] = "shape2";
                    //slide.Shapes.Range(myRangeArray).Group();
                    break;
                }
                if ((inst.Left + inst.Width > newShp.Left - Point_Param.pinWidth * 2) && (inst.Left + inst.Width < newShp.Left + Point_Param.pinWidth * 2) && (inst.Top < newShp.Top) && (inst.Top + inst.Height > newShp.Top + Point_Param.pinWidth))
                {
                    if (!GridlistX.Contains(inst) && !GridlistY.Contains(inst))
                        newShp.Left = inst.Left + inst.Width - Point_Param.pinWidth / 2;
                    //inst.Name = "shape1";
                    //shp.Name = "shape2";
                    //string[] myRangeArray = new string[2];
                    //myRangeArray[0] = "shape1";
                    //myRangeArray[1] = "shape2";
                    //slide.Shapes.Range(myRangeArray).Group();
                    break;
                }
                if ((inst.Left < newShp.Left) && (inst.Left + inst.Width > newShp.Left + Point_Param.pinWidth) && (inst.Top > newShp.Top - Point_Param.pinWidth * 2) && (inst.Top < newShp.Top + Point_Param.pinWidth * 2))
                {
                    if (!GridlistX.Contains(inst) && !GridlistY.Contains(inst))
                        newShp.Top = inst.Top - Point_Param.pinWidth / 2;
                    //inst.Name = "shape1";
                    //shp.Name = "shape2";
                    //string[] myRangeArray = new string[2];
                    //myRangeArray[0] = "shape1";
                    //myRangeArray[1] = "shape2";
                    //slide.Shapes.Range(myRangeArray).Group();
                    break;
                }
            }
        }
        public static void DrawGrid()
        {
            bGridcheck = true;
            Microsoft.Office.Interop.PowerPoint._Slide curslide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            int sld_width = (int)Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth;
            int sld_height = (int)Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight;
            for (int i = 0; i <= sld_width / nGridSpace; i++)
            {
                int gridx = i * nGridSpace;
                PowerPoint.Shape gridinst = curslide.Shapes.AddLine(gridx, 0, gridx, sld_height);
                gridinst.Tags.Add("sel", "unsel");
                gridinst.GetType().InvokeMember("Locked", System.Reflection.BindingFlags.SetProperty, null, gridinst, new object[] { MsoTriState.msoTrue });
                if (i % 4 == 0)
                {
                    gridinst.Line.Weight = 1f;
                }
                else
                    gridinst.Line.Weight = 0.1f;
                gridinst.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(128, 128, 128).ToArgb();
                GridlistX.Add(gridinst);
            }
            for (int i = 0; i <= sld_height / nGridSpace; i++)
            {
                int gridy = i * nGridSpace;
                PowerPoint.Shape gridinst = curslide.Shapes.AddLine(0, gridy, sld_width, gridy);
                gridinst.Tags.Add("sel", "unsel");
                gridinst.GetType().InvokeMember("Locked", System.Reflection.BindingFlags.SetProperty, null, gridinst, new object[] { MsoTriState.msoTrue });
                if (i % 4 == 0)
                {
                    gridinst.Line.Weight = 1f;
                }
                else
                    gridinst.Line.Weight = 0.1f;
                gridinst.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(211, 211, 211).ToArgb();
                GridlistY.Add(gridinst);
            }
        }
        public static void DelGrid()
        {
            bGridcheck = false;
            foreach (PowerPoint.Shape instshp in GridlistX)
                instshp.Delete();
            foreach (PowerPoint.Shape instshp in GridlistY)
                instshp.Delete();
            GridlistX.Clear();
            GridlistY.Clear();
        }

        public PowerPoint.Shape AddGroupShape(Group_Param param)
        {
            PowerPoint.Shape text;
            PowerPoint.Shape groupShape = null;
            //           string[] pin_names = param.pin_names.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            try
            {
                switch (param.kind)
                {
                    case "and":             //and
                        {
                            PowerPoint.Shape line1 = slide.Shapes.AddLine(0, ribbonHeight - 10, 20, ribbonHeight - 10);
                            line1.Line.Weight = 1.5f;
                            PowerPoint.Shape line2 = slide.Shapes.AddLine(0, ribbonHeight + 10, 20, ribbonHeight + 10); line2.Line.Weight = 1.5f;
                            Point delyPt = Point.Empty;
                            delyPt.X = 20;
                            delyPt.Y = ribbonHeight - 15;
                            PowerPoint.Shape Delay = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeFlowchartDelay,
                                                        delyPt.X, delyPt.Y, 40, 70);
                            Delay.Width = 30;
                            Delay.Height = 30;
                            PowerPoint.Shape line3 = slide.Shapes.AddLine(50, ribbonHeight, 70, ribbonHeight);
                            line3.Line.Weight = 1.5f;
                            line1.Name = "andline1" + param.sub_id.ToString();
                            line2.Name = "andline2" + param.sub_id.ToString();
                            line3.Name = "andline3" + param.sub_id.ToString();
                            Delay.Name = "anddelay" + param.sub_id.ToString();
                            string[] myRangeArray = new string[8];
                            myRangeArray[0] = line1.Name;
                            myRangeArray[1] = line2.Name;
                            myRangeArray[2] = line3.Name;
                            myRangeArray[3] = Delay.Name;

                            //                         if (pin_names.Length > 0)
                            //                         {
                            //                             myRangeArray[4] = pin_names[0];
                            //                             myRangeArray[5] = pin_names[1];
                            //                             myRangeArray[6] = pin_names[2];
                            //                         }
                            //                         else
                            {
                                PowerPoint.Shape andpt1 = AddPointShape(new Point_Param(new Point(0, ribbonHeight - 10)));
                                PowerPoint.Shape andpt2 = AddPointShape(new Point_Param(new Point(0, ribbonHeight + 10)));
                                PowerPoint.Shape andpt3 = AddPointShape(new Point_Param(new Point(70, ribbonHeight)));

                                andpt1.Name = "andpt1" + param.sub_id.ToString();
                                andpt2.Name = "andpt2" + param.sub_id.ToString();
                                andpt3.Name = "andpt3" + param.sub_id.ToString();
                                myRangeArray[4] = andpt1.Name;
                                myRangeArray[5] = andpt2.Name;
                                myRangeArray[6] = andpt3.Name;
                                //param.pin_names = $"{andpt1.Name},{andpt2.Name},{andpt3.Name}";
                            }

                            text = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, delyPt.X, delyPt.Y - 30, 60, 5);
                            text.TextFrame.TextRange.Text = "UC" + param.sub_id.ToString();
                            text.Name = "Ctext" + param.sub_id.ToString();
                            myRangeArray[7] = text.Name;
                            groupShape = slide.Shapes.Range(myRangeArray).Group();
                        }
                        break;
                    case "buffer":     //buffer
                        PowerPoint.Shape bufferline1 = slide.Shapes.AddLine(0, ribbonHeight, 20, ribbonHeight);
                        bufferline1.Line.Weight = 1.5f;
                        Point bufptPt = Point.Empty;
                        bufptPt.X = 20;
                        bufptPt.Y = ribbonHeight - 15;
                        PowerPoint.Shape buffer = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeIsoscelesTriangle,
                                                    bufptPt.X, bufptPt.Y, 40, 70);
                        buffer.Width = 30;
                        buffer.Height = 30;
                        buffer.Rotation = 90;
                        PowerPoint.Shape bufferline2 = slide.Shapes.AddLine(50, ribbonHeight, 70, ribbonHeight);
                        bufferline2.Line.Weight = 1.5f;
                        bufferline1.Name = "bufline1" + param.sub_id.ToString();
                        bufferline2.Name = "bufline2" + param.sub_id.ToString();
                        buffer.Name = "bufbuffer" + param.sub_id.ToString();
                        string[] strbufarry = new string[6];
                        strbufarry[0] = bufferline1.Name;
                        strbufarry[1] = bufferline2.Name;
                        strbufarry[2] = buffer.Name;

                        //                         if (pin_names.Length > 0)
                        //                         {
                        //                             strbufarry[3] = pin_names[0];
                        //                             strbufarry[4] = pin_names[1];
                        //                         }
                        //                         else
                        {
                            PowerPoint.Shape bufferpt1 = AddPointShape(new Point_Param(new Point(0, ribbonHeight)));
                            PowerPoint.Shape bufferpt2 = AddPointShape(new Point_Param(new Point(70, ribbonHeight)));
                            bufferpt1.Name = "bufferpt1" + param.sub_id.ToString();
                            bufferpt2.Name = "bufferpt2" + param.sub_id.ToString();
                            strbufarry[3] = bufferpt1.Name;
                            strbufarry[4] = bufferpt2.Name;
                            //param.pin_names = $"{bufferpt1.Name},{bufferpt2.Name}";
                        }
                        text = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, bufptPt.X, bufptPt.Y - 20, 60, 5);
                        text.Name = "Ctext" + param.sub_id.ToString();
                        strbufarry[5] = text.Name;
                        groupShape = slide.Shapes.Range(strbufarry).Group();
                        break;
                    case "nand":                                  //nand 
                        PowerPoint.Shape nandline1 = slide.Shapes.AddLine(0, ribbonHeight - 10, 20, ribbonHeight - 10);
                        nandline1.Line.Weight = 1.5f;
                        PowerPoint.Shape nandline2 = slide.Shapes.AddLine(0, ribbonHeight + 10, 20, ribbonHeight + 10); nandline2.Line.Weight = 1.5f;
                        Point nanddelyPt = Point.Empty;
                        nanddelyPt.X = 20;
                        nanddelyPt.Y = ribbonHeight - 15;
                        PowerPoint.Shape nandDelay = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeFlowchartDelay,
                                                    nanddelyPt.X, nanddelyPt.Y, 40, 70);
                        nandDelay.Width = 30;
                        nandDelay.Height = 30;
                        PowerPoint.Shape nandcircle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeFlowchartConnector,
                                                    50, ribbonHeight - Point_Param.pinWidth / 2, 50 + Point_Param.pinWidth, ribbonHeight + Point_Param.pinWidth / 2);
                        nandcircle.Width = Point_Param.pinWidth;
                        nandcircle.Height = Point_Param.pinWidth;
                        PowerPoint.Shape nandline3 = slide.Shapes.AddLine(50 + Point_Param.pinWidth, ribbonHeight, 70 + Point_Param.pinWidth, ribbonHeight);
                        nandline3.Line.Weight = 1.5f;
                        nandline1.Name = "nandline1" + param.sub_id.ToString();
                        nandline2.Name = "nandline2" + param.sub_id.ToString();
                        nandline3.Name = "nandline3" + param.sub_id.ToString();
                        nandDelay.Name = "nanddelay" + param.sub_id.ToString();
                        nandcircle.Name = "nandCircle" + param.sub_id.ToString();
                        string[] nandarray = new string[9];
                        nandarray[0] = nandline1.Name;
                        nandarray[1] = nandline2.Name;
                        nandarray[2] = nandline3.Name;
                        nandarray[3] = nandDelay.Name;
                        nandarray[4] = nandcircle.Name;

                        //                         if (pin_names.Length > 0)
                        //                         {
                        //                             nandarray[5] = pin_names[0];
                        //                             nandarray[6] = pin_names[1];
                        //                             nandarray[7] = pin_names[2];
                        //                         }
                        //                         else
                        {
                            PowerPoint.Shape nandpt1 = AddPointShape(new Point_Param(new Point(0, ribbonHeight - 10)));
                            PowerPoint.Shape nandpt2 = AddPointShape(new Point_Param(new Point(0, ribbonHeight + 10)));
                            PowerPoint.Shape nandpt3 = AddPointShape(new Point_Param(new Point(70 + Point_Param.pinWidth, ribbonHeight)));
                            nandpt1.Name = "nandpt1" + param.sub_id.ToString();
                            nandpt2.Name = "nandpt2" + param.sub_id.ToString();
                            nandpt3.Name = "nandpt3" + param.sub_id.ToString();
                            nandarray[5] = nandpt1.Name;
                            nandarray[6] = nandpt2.Name;
                            nandarray[7] = nandpt3.Name;
                            //param.pin_names = $"{nandpt1.Name},{nandpt2.Name},{nandpt3.Name}";
                        }


                        text = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, nanddelyPt.X, nanddelyPt.Y - 20, 60, 5);
                        nandarray[8] = text.Name;
                        groupShape = slide.Shapes.Range(nandarray).Group();
                        break;
                    case "nor":                     //nor
                        PowerPoint.Shape norline1 = slide.Shapes.AddLine(0, ribbonHeight - 10, 25, ribbonHeight - 10);
                        norline1.Line.Weight = 1.5f;
                        PowerPoint.Shape norline2 = slide.Shapes.AddLine(0, ribbonHeight + 10, 25, ribbonHeight + 10); norline2.Line.Weight = 1.5f;
                        Point norPt = Point.Empty;
                        norPt.X = 20;
                        norPt.Y = ribbonHeight - 15;
                        PowerPoint.Shape norDelay = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeFlowchartStoredData,
                                                    norPt.X, norPt.Y, 40, 70);
                        norDelay.Width = 30;
                        norDelay.Height = 30;
                        norDelay.Rotation = 180;
                        PowerPoint.Shape norcircle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeFlowchartConnector,
                                                    50, ribbonHeight - Point_Param.pinWidth / 2, 50 + Point_Param.pinWidth, ribbonHeight + Point_Param.pinWidth / 2);
                        norcircle.Width = Point_Param.pinWidth;
                        norcircle.Height = Point_Param.pinWidth;
                        PowerPoint.Shape norline3 = slide.Shapes.AddLine(50 + Point_Param.pinWidth, ribbonHeight, 70 + Point_Param.pinWidth, ribbonHeight);
                        norline3.Line.Weight = 1.5f;
                        norline1.Name = "norline1" + param.sub_id.ToString();
                        norline2.Name = "norline2" + param.sub_id.ToString();
                        norline3.Name = "norline3" + param.sub_id.ToString();
                        norDelay.Name = "nordelay" + param.sub_id.ToString();
                        norcircle.Name = "norCircle" + param.sub_id.ToString();
                        string[] norarray = new string[9];
                        norarray[0] = norline1.Name;
                        norarray[1] = norline2.Name;
                        norarray[2] = norline3.Name;
                        norarray[3] = norDelay.Name;
                        norarray[4] = norcircle.Name;
                        //                         if (pin_names.Length > 0)
                        //                         {
                        //                             norarray[5] = pin_names[0];
                        //                             norarray[6] = pin_names[1];
                        //                             norarray[7] = pin_names[2];
                        //                         }
                        //                         else
                        {
                            PowerPoint.Shape norpt1 = AddPointShape(new Point_Param(new Point(0, ribbonHeight - 10)));
                            PowerPoint.Shape norpt2 = AddPointShape(new Point_Param(new Point(0, ribbonHeight + 10)));
                            PowerPoint.Shape norpt3 = AddPointShape(new Point_Param(new Point(70 + Point_Param.pinWidth, ribbonHeight)));
                            norpt1.Name = "norpt1" + param.sub_id.ToString();
                            norpt2.Name = "norpt2" + param.sub_id.ToString();
                            norpt3.Name = "norpt3" + param.sub_id.ToString();
                            norarray[5] = norpt1.Name;
                            norarray[6] = norpt2.Name;
                            norarray[7] = norpt3.Name;
                            //param.pin_names = $"{norpt1.Name},{norpt2.Name},{norpt3.Name}";
                        }


                        text = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, norPt.X, norPt.Y - 20, 60, 5);
                        norarray[8] = text.Name;
                        groupShape = slide.Shapes.Range(norarray).Group();
                        break;
                    case "not":         //not
                        PowerPoint.Shape notline1 = slide.Shapes.AddLine(0, ribbonHeight, 20, ribbonHeight);
                        notline1.Line.Weight = 1.5f;
                        Point notPt = Point.Empty;
                        notPt.X = 20;
                        notPt.Y = ribbonHeight - 15;
                        PowerPoint.Shape not = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeIsoscelesTriangle,
                                                    notPt.X, notPt.Y, 40, 70);
                        not.Width = 30;
                        not.Height = 30;
                        not.Rotation = 90;
                        PowerPoint.Shape notcircle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeFlowchartConnector,
                                                    50, ribbonHeight - Point_Param.pinWidth / 2, 50 + Point_Param.pinWidth, ribbonHeight + Point_Param.pinWidth / 2);
                        notcircle.Width = Point_Param.pinWidth;
                        notcircle.Height = Point_Param.pinWidth;
                        PowerPoint.Shape notline2 = slide.Shapes.AddLine(50, ribbonHeight, 70, ribbonHeight);
                        notline2.Line.Weight = 1.5f;
                        notline1.Name = "notline1" + param.sub_id.ToString();
                        notline2.Name = "notline2" + param.sub_id.ToString();
                        not.Name = "notbuffer" + param.sub_id.ToString();
                        string[] strnotarry = new string[7];
                        strnotarry[0] = notline1.Name;
                        strnotarry[1] = notline2.Name;
                        strnotarry[2] = not.Name;
                        strnotarry[3] = notcircle.Name;
                        //                         if (pin_names.Length > 0)
                        //                         {
                        //                             strnotarry[4] = pin_names[0];
                        //                             strnotarry[5] = pin_names[1];
                        //                         }
                        //                         else
                        {
                            PowerPoint.Shape notpt1 = AddPointShape(new Point_Param(new Point(0, ribbonHeight)));
                            PowerPoint.Shape notpt2 = AddPointShape(new Point_Param(new Point(70, ribbonHeight)));
                            notpt1.Name = "notpt1" + param.sub_id.ToString();
                            notpt2.Name = "notpt2" + param.sub_id.ToString();
                            strnotarry[4] = notpt1.Name;
                            strnotarry[5] = notpt2.Name;

                            //param.pin_names = $"{notpt1.Name},{notpt2.Name}";
                        }


                        text = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, notPt.X, notPt.Y - 20, 60, 5);
                        text.TextFrame.TextRange.Text = "UC" + param.sub_id.ToString();
                        text.Name = "Ctext" + param.sub_id.ToString();
                        strnotarry[6] = text.Name;
                        groupShape = slide.Shapes.Range(strnotarry).Group();
                        break;
                    case "or":                             //or
                        PowerPoint.Shape orline1 = slide.Shapes.AddLine(0, ribbonHeight - 10, 25, ribbonHeight - 10);
                        orline1.Line.Weight = 1.5f;
                        PowerPoint.Shape orline2 = slide.Shapes.AddLine(0, ribbonHeight + 10, 25, ribbonHeight + 10); orline2.Line.Weight = 1.5f;
                        Point orPt = Point.Empty;
                        orPt.X = 20;
                        orPt.Y = ribbonHeight - 15;
                        PowerPoint.Shape orDelay = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeFlowchartStoredData,
                                                    orPt.X, orPt.Y, 40, 70);
                        orDelay.Width = 30;
                        orDelay.Height = 30;
                        orDelay.Rotation = 180;
                        PowerPoint.Shape orline3 = slide.Shapes.AddLine(50, ribbonHeight, 70, ribbonHeight);
                        orline3.Line.Weight = 1.5f;
                        orline1.Name = "orline1" + param.sub_id.ToString();
                        orline2.Name = "orline2" + param.sub_id.ToString();
                        orline3.Name = "orline3" + param.sub_id.ToString();
                        orDelay.Name = "ordelay" + param.sub_id.ToString();
                        string[] orarray = new string[8];
                        orarray[0] = orline1.Name;
                        orarray[1] = orline2.Name;
                        orarray[2] = orline3.Name;
                        orarray[3] = orDelay.Name;

                        //                         if (pin_names.Length > 0)
                        //                         {
                        //                             orarray[4] = pin_names[0];
                        //                             orarray[5] = pin_names[1];
                        //                             orarray[6] = pin_names[2];
                        //                         }
                        //                         else
                        {
                            PowerPoint.Shape orpt1 = AddPointShape(new Point_Param(new Point(0, ribbonHeight - 10)));
                            PowerPoint.Shape orpt2 = AddPointShape(new Point_Param(new Point(0, ribbonHeight + 10)));
                            PowerPoint.Shape orpt3 = AddPointShape(new Point_Param(new Point(70, ribbonHeight)));

                            orpt1.Name = "orpt1" + param.sub_id.ToString();
                            orpt2.Name = "orpt2" + param.sub_id.ToString();
                            orpt3.Name = "orpt3" + param.sub_id.ToString();
                            orarray[4] = orpt1.Name;
                            orarray[5] = orpt2.Name;
                            orarray[6] = orpt3.Name;
                            //param.pin_names = $"{orpt1.Name},{orpt2.Name},{orpt3.Name}";
                        }

                        text = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, orPt.X, orPt.Y - 20, 60, 5);
                        text.TextFrame.TextRange.Text = "UC" + param.sub_id.ToString();
                        text.Name = "Ctext" + param.sub_id.ToString();
                        orarray[7] = text.Name;
                        groupShape = slide.Shapes.Range(orarray).Group();
                        break;
                    case "xnor":         //XNOR
                        PowerPoint.Shape Xnorline1 = slide.Shapes.AddLine(0, ribbonHeight - 10, 25, ribbonHeight - 10);
                        Xnorline1.Line.Weight = 1.5f;
                        PowerPoint.Shape Xnorline2 = slide.Shapes.AddLine(0, ribbonHeight + 10, 25, ribbonHeight + 10);
                        Xnorline2.Line.Weight = 1.5f;
                        Point XnorPt = Point.Empty;
                        XnorPt.X = 20;
                        XnorPt.Y = ribbonHeight - 15;
                        PowerPoint.Shape XnorDelay = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeFlowchartStoredData,
                                                    XnorPt.X, XnorPt.Y, 40, 70);
                        XnorDelay.Width = 30;
                        XnorDelay.Height = 30;
                        XnorDelay.Rotation = 180;
                        PowerPoint.Shape xnorarc = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRightBracket,
                                                    16, XnorPt.Y, 16, 70);

                        xnorarc.Width = 2;
                        xnorarc.Height = 30;
                        xnorarc.Line.Weight = 1.5f;
                        PowerPoint.Shape Xnorcircle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeFlowchartConnector,
                                                    50, ribbonHeight - Point_Param.pinWidth / 2, 50 + Point_Param.pinWidth, ribbonHeight + Point_Param.pinWidth / 2);
                        Xnorcircle.Width = Point_Param.pinWidth;
                        Xnorcircle.Height = Point_Param.pinWidth;
                        PowerPoint.Shape Xnorline3 = slide.Shapes.AddLine(50 + Point_Param.pinWidth, ribbonHeight, 70 + Point_Param.pinWidth, ribbonHeight);
                        Xnorline3.Line.Weight = 1.5f;
                        Xnorline1.Name = "xnorline1" + param.sub_id.ToString();
                        Xnorline2.Name = "xnorline2" + param.sub_id.ToString();
                        Xnorline3.Name = "xnorline3" + param.sub_id.ToString();
                        XnorDelay.Name = "xnordelay" + param.sub_id.ToString();
                        Xnorcircle.Name = "xnorCircle" + param.sub_id.ToString();
                        string[] Xnorarray = new string[9];
                        Xnorarray[0] = Xnorline1.Name;
                        Xnorarray[1] = Xnorline2.Name;
                        Xnorarray[2] = Xnorline3.Name;
                        Xnorarray[3] = XnorDelay.Name;
                        Xnorarray[4] = Xnorcircle.Name;

                        //                         if (pin_names.Length > 0)
                        //                         {
                        //                             Xnorarray[5] = pin_names[0];
                        //                             Xnorarray[6] = pin_names[1];
                        //                             Xnorarray[7] = pin_names[2];
                        //                         }
                        //                         else
                        {
                            PowerPoint.Shape xnorpt1 = AddPointShape(new Point_Param(new Point(0, ribbonHeight - 10)));
                            PowerPoint.Shape xnorpt2 = AddPointShape(new Point_Param(new Point(0, ribbonHeight + 10)));
                            PowerPoint.Shape xnorpt3 = AddPointShape(new Point_Param(new Point(70 + Point_Param.pinWidth, ribbonHeight)));
                            xnorpt1.Name = "xnorpt1" + param.sub_id.ToString();
                            xnorpt2.Name = "xnorpt2" + param.sub_id.ToString();
                            xnorpt3.Name = "xnorpt3" + param.sub_id.ToString();
                            Xnorarray[5] = xnorpt1.Name;
                            Xnorarray[6] = xnorpt2.Name;
                            Xnorarray[7] = xnorpt3.Name;
                            //param.pin_names = $"{xnorpt1.Name},{xnorpt2.Name},{xnorpt3.Name}";
                        }

                        text = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, XnorPt.X, XnorPt.Y - 20, 60, 5);
                        text.TextFrame.TextRange.Text = "UC" + param.sub_id.ToString();
                        text.Name = "Ctext" + param.sub_id.ToString();
                        Xnorarray[8] = text.Name;
                        groupShape = slide.Shapes.Range(Xnorarray).Group();
                        break;
                    case "xor":                     //xor
                        PowerPoint.Shape xorline1 = slide.Shapes.AddLine(0, ribbonHeight - 10, 25, ribbonHeight - 10);
                        xorline1.Line.Weight = 1.5f;
                        PowerPoint.Shape xorline2 = slide.Shapes.AddLine(0, ribbonHeight + 10, 25, ribbonHeight + 10); xorline2.Line.Weight = 1.5f;
                        Point xorPt = Point.Empty;
                        xorPt.X = 20;
                        xorPt.Y = ribbonHeight - 15;
                        PowerPoint.Shape xorDelay = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeFlowchartStoredData,
                                                    xorPt.X, xorPt.Y, 40, 70);
                        xorDelay.Width = 30;
                        xorDelay.Height = 30;
                        xorDelay.Rotation = 180;
                        PowerPoint.Shape xorarc = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRightBracket,
                                                   16, xorPt.Y, 16, 70);

                        xorarc.Width = 2;
                        xorarc.Height = 30;
                        xorarc.Line.Weight = 1.5f;
                        PowerPoint.Shape xorline3 = slide.Shapes.AddLine(50, ribbonHeight, 70, ribbonHeight);
                        xorline3.Line.Weight = 1.5f;
                        xorline1.Name = "orline1" + param.sub_id.ToString();
                        xorline2.Name = "orline2" + param.sub_id.ToString();
                        xorline3.Name = "orline3" + param.sub_id.ToString();
                        xorDelay.Name = "ordelay" + param.sub_id.ToString();
                        string[] xorarray = new string[8];
                        xorarray[0] = xorline1.Name;
                        xorarray[1] = xorline2.Name;
                        xorarray[2] = xorline3.Name;
                        xorarray[3] = xorDelay.Name;

                        //                         if (pin_names.Length > 0)
                        //                         {
                        //                             xorarray[4] = pin_names[0];
                        //                             xorarray[5] = pin_names[1];
                        //                             xorarray[6] = pin_names[2];
                        //                         }
                        //                         else
                        {
                            PowerPoint.Shape xorpt1 = AddPointShape(new Point_Param(new Point(0, ribbonHeight - 10)));
                            PowerPoint.Shape xorpt2 = AddPointShape(new Point_Param(new Point(0, ribbonHeight + 10)));
                            PowerPoint.Shape xorpt3 = AddPointShape(new Point_Param(new Point(70, ribbonHeight)));
                            xorpt1.Name = "xorpt1" + param.sub_id.ToString();
                            xorpt2.Name = "xorpt2" + param.sub_id.ToString();
                            xorpt3.Name = "xorpt3" + param.sub_id.ToString();
                            xorarray[4] = xorpt1.Name;
                            xorarray[5] = xorpt2.Name;
                            xorarray[6] = xorpt3.Name;

                            //param.pin_names = $"{xorpt1.Name},{xorpt2.Name},{xorpt3.Name}";
                        }

                        text = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, xorPt.X, xorPt.Y - 20, 60, 10);
                        xorarray[7] = text.Name;
                        groupShape = slide.Shapes.Range(xorarray).Group();
                        break;
                    //sequence curcuit
                    case "dflop":             //D-flop
                        {
                            PowerPoint.Shape dflopline1 = slide.Shapes.AddLine(0, ribbonHeight - 15, 20, ribbonHeight - 15);
                            dflopline1.Line.Weight = 1.5f;
                            PowerPoint.Shape dflopline2 = slide.Shapes.AddLine(0, ribbonHeight + 15, 20, ribbonHeight + 15); dflopline2.Line.Weight = 1.5f;
                            Point delyPt = Point.Empty;
                            delyPt.X = 20;
                            delyPt.Y = ribbonHeight - 25;
                            PowerPoint.Shape dflopBox = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                                                        delyPt.X, delyPt.Y, 0, 0);
                            dflopBox.Width = 30;
                            dflopBox.Height = 50;
                            dflopBox.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                            PowerPoint.Shape dfloptriangle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeIsoscelesTriangle,
                                                        delyPt.X, ribbonHeight + 15 - Point_Param.pinWidth / 2, 0, 0);
                            dfloptriangle.Width = Point_Param.pinWidth;
                            dfloptriangle.Height = Point_Param.pinWidth;
                            dfloptriangle.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                            dfloptriangle.Rotation = 90;
                            PowerPoint.Shape dflopline3 = slide.Shapes.AddLine(50, ribbonHeight - 15, 70, ribbonHeight - 15);
                            dflopline3.Line.Weight = 1.5f;
                            dflopline1.Name = "dflopline1" + param.sub_id.ToString();
                            dflopline2.Name = "dflopline2" + param.sub_id.ToString();
                            dflopline3.Name = "dflopline3" + param.sub_id.ToString();
                            dflopBox.Name = "dflopdelay" + param.sub_id.ToString();
                            dfloptriangle.Name = "dfloptriangle" + param.sub_id.ToString();
                            string[] dflopArray = new string[9];
                            dflopArray[0] = dflopline1.Name;
                            dflopArray[1] = dflopline2.Name;
                            dflopArray[2] = dflopline3.Name;
                            dflopArray[3] = dflopBox.Name;
                            dflopArray[4] = dfloptriangle.Name;

                            PowerPoint.Shape dfloppt1 = AddPointShape(new Point_Param(new Point(0, ribbonHeight - 15)));
                            PowerPoint.Shape dfloppt2 = AddPointShape(new Point_Param(new Point(0, ribbonHeight + 15)));
                            PowerPoint.Shape dfloppt3 = AddPointShape(new Point_Param(new Point(70, ribbonHeight - 15)));

                            dfloppt1.Name = "dfloppt1" + param.sub_id.ToString();
                            dfloppt2.Name = "dfloppt2" + param.sub_id.ToString();
                            dfloppt3.Name = "dfloppt3" + param.sub_id.ToString();
                            dflopArray[5] = dfloppt1.Name;
                            dflopArray[6] = dfloppt2.Name;
                            dflopArray[7] = dfloppt3.Name;
                            text = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, delyPt.X, delyPt.Y - 30, 60, 5);
                            dflopArray[8] = text.Name;
                            groupShape = slide.Shapes.Range(dflopArray).Group();
                        }
                        break;
                    case "latch":                 //latch
                        PowerPoint.Shape latchline1 = slide.Shapes.AddLine(0, ribbonHeight - 15, 20, ribbonHeight - 15);
                        latchline1.Line.Weight = 1.5f;
                        PowerPoint.Shape latchline2 = slide.Shapes.AddLine(0, ribbonHeight + 15, 20, ribbonHeight + 15); latchline2.Line.Weight = 1.5f;
                        Point latchPt = Point.Empty;
                        latchPt.X = 20;
                        latchPt.Y = ribbonHeight - 25;
                        PowerPoint.Shape latchBox = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                                                    latchPt.X, latchPt.Y, 0, 0);
                        latchBox.Width = 30;
                        latchBox.Height = 50;
                        latchBox.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                        PowerPoint.Shape latchtriangle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                                                    latchPt.X, ribbonHeight + 15 - Point_Param.pinWidth / 2, 0, 0);
                        latchtriangle.Width = Point_Param.pinWidth;
                        latchtriangle.Height = Point_Param.pinWidth;
                        latchtriangle.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                        PowerPoint.Shape latchline3 = slide.Shapes.AddLine(50, ribbonHeight - 15, 70, ribbonHeight - 15);
                        latchline3.Line.Weight = 1.5f;
                        latchline1.Name = "dflopline1" + param.sub_id.ToString();
                        latchline2.Name = "dflopline2" + param.sub_id.ToString();
                        latchline3.Name = "dflopline3" + param.sub_id.ToString();
                        latchBox.Name = "dflopdelay" + param.sub_id.ToString();
                        latchtriangle.Name = "dfloptriangle" + param.sub_id.ToString();
                        string[] latchArray = new string[9];
                        latchArray[0] = latchline1.Name;
                        latchArray[1] = latchline2.Name;
                        latchArray[2] = latchline3.Name;
                        latchArray[3] = latchBox.Name;
                        latchArray[4] = latchtriangle.Name;

                        PowerPoint.Shape latchpt1 = AddPointShape(new Point_Param(new Point(0, ribbonHeight - 15)));
                        PowerPoint.Shape latchpt2 = AddPointShape(new Point_Param(new Point(0, ribbonHeight + 15)));
                        PowerPoint.Shape latchpt3 = AddPointShape(new Point_Param(new Point(70, ribbonHeight - 15)));
                        latchpt1.Name = "latchpt1" + param.sub_id.ToString();
                        latchpt2.Name = "latchpt2" + param.sub_id.ToString();
                        latchpt3.Name = "latchpt3" + param.sub_id.ToString();
                        latchArray[5] = latchpt1.Name;
                        latchArray[6] = latchpt2.Name;
                        latchArray[7] = latchpt3.Name;
                        text = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, latchPt.X, latchPt.Y - 30, 60, 5);
                        latchArray[8] = text.Name;
                        groupShape = slide.Shapes.Range(latchArray).Group();
                        break;
                    case "sync":                 //synchronizer
                        PowerPoint.Shape syncline1 = slide.Shapes.AddLine(0, ribbonHeight - 15, 20, ribbonHeight - 15);
                        syncline1.Line.Weight = 1.5f;
                        PowerPoint.Shape syncline2 = slide.Shapes.AddLine(0, ribbonHeight + 15, 20, ribbonHeight + 15);
                        syncline2.Line.Weight = 1.5f;
                        Point syncPt = Point.Empty;
                        syncPt.X = 20;
                        syncPt.Y = ribbonHeight - 25;
                        PowerPoint.Shape syncBox = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                                                    syncPt.X, syncPt.Y, 0, 0);
                        syncBox.Width = 30;
                        syncBox.Height = 50;
                        syncBox.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                        PowerPoint.Shape synctriangle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeIsoscelesTriangle,
                                                    syncPt.X, ribbonHeight + 15 - Point_Param.pinWidth / 2, 0, 0);
                        synctriangle.Width = Point_Param.pinWidth;
                        synctriangle.Height = Point_Param.pinWidth;
                        synctriangle.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                        synctriangle.Rotation = 90;
                        PowerPoint.Shape syncline3 = slide.Shapes.AddLine(50, ribbonHeight - 15, 90, ribbonHeight - 15);
                        syncline3.Line.Weight = 1.5f;

                        PowerPoint.Shape syncBox1 = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                                                    90, syncPt.Y, 0, 0);
                        syncBox1.Width = 30;
                        syncBox1.Height = 50;
                        syncBox1.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                        PowerPoint.Shape synctriangle1 = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeIsoscelesTriangle,
                                                    90, ribbonHeight + 15 - Point_Param.pinWidth / 2, 0, 0);
                        synctriangle1.Width = Point_Param.pinWidth;
                        synctriangle1.Height = Point_Param.pinWidth;
                        synctriangle1.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                        synctriangle1.Rotation = 90;
                        PowerPoint.Shape syncline4 = slide.Shapes.AddLine(120, ribbonHeight - 15, 140, ribbonHeight - 15);
                        syncline4.Line.Weight = 1.5f;
                        PowerPoint.Shape syncline5 = slide.Shapes.AddLine(70, ribbonHeight + 15, 90, ribbonHeight + 15);
                        syncline5.Line.Weight = 1.5f;
                        PowerPoint.Shape syncline6 = slide.Shapes.AddLine(0, ribbonHeight + 15, 0, ribbonHeight + 35);
                        syncline6.Line.Weight = 1.5f;
                        PowerPoint.Shape syncline7 = slide.Shapes.AddLine(70, ribbonHeight + 15, 70, ribbonHeight + 35);
                        syncline7.Line.Weight = 1.5f;
                        PowerPoint.Shape syncline8 = slide.Shapes.AddLine(0 - 20, ribbonHeight + 35, 70, ribbonHeight + 35);
                        syncline8.Line.Weight = 1.5f;
                        syncline1.Name = "syncline1" + param.sub_id.ToString();
                        syncline2.Name = "syncline2" + param.sub_id.ToString();
                        syncline3.Name = "syncline3" + param.sub_id.ToString();
                        syncline4.Name = "syncline4" + param.sub_id.ToString();
                        syncline5.Name = "syncline5" + param.sub_id.ToString();
                        syncline6.Name = "syncline6" + param.sub_id.ToString();
                        syncline7.Name = "syncline7" + param.sub_id.ToString();
                        syncline8.Name = "syncline8" + param.sub_id.ToString();
                        syncBox.Name = "syncBox" + param.sub_id.ToString();
                        synctriangle.Name = "synctriangle" + param.sub_id.ToString();
                        syncBox1.Name = "syncBox1" + param.sub_id.ToString();
                        synctriangle1.Name = "synctriangle1" + param.sub_id.ToString();
                        string[] syncArray = new string[16];
                        syncArray[0] = syncline1.Name;
                        syncArray[1] = syncline2.Name;
                        syncArray[2] = syncline3.Name;
                        syncArray[3] = syncline4.Name;
                        syncArray[4] = syncline5.Name;
                        syncArray[5] = syncline6.Name;
                        syncArray[6] = syncline7.Name;
                        syncArray[7] = syncline8.Name;
                        syncArray[8] = syncBox.Name;
                        syncArray[9] = synctriangle.Name;
                        syncArray[10] = syncBox1.Name;
                        syncArray[11] = synctriangle1.Name;

                        PowerPoint.Shape syncpt1 = AddPointShape(new Point_Param(new Point(0, ribbonHeight - 15)));
                        PowerPoint.Shape syncpt2 = AddPointShape(new Point_Param(new Point(140, ribbonHeight - 15)));
                        PowerPoint.Shape syncpt3 = AddPointShape(new Point_Param(new Point(0 - 20, ribbonHeight + 35)));
                        syncpt1.Name = "pt1" + param.sub_id.ToString();
                        syncpt2.Name = "pt2" + param.sub_id.ToString();
                        syncpt3.Name = "pt3" + param.sub_id.ToString();
                        syncArray[12] = syncpt1.Name;
                        syncArray[13] = syncpt2.Name;
                        syncArray[14] = syncpt3.Name;
                        text = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, syncPt.X, syncPt.Y - 30, 60, 5);
                        syncArray[15] = text.Name;
                        groupShape = slide.Shapes.Range(syncArray).Group();
                        break;
                    default:
                        return null;
                }
                if (groupShape != null)
                {
                    text.TextFrame.TextRange.Text = param.label;
                    if (param.isCombnation)
                        text.Name = "Ctext" + param.sub_id.ToString();
                    else
                        text.Name = "Stext" + param.sub_id.ToString();

                    foreach (PowerPoint.Shape subShape in groupShape.GroupItems)
                    {
                        subShape.Tags.Add("sel", "unsel");
                    }
                    text.Tags.Delete("sel");
                    groupShape.LockAspectRatio = MsoTriState.msoTrue;

                    if (param.width != 0)
                        groupShape.Width = param.width;
                    if (param.height != 0)
                        groupShape.Height = param.height;
                    groupShape.Left = param.pos.X - groupShape.Width / 2f;
                    groupShape.Top = param.pos.Y - groupShape.Height / 2f;
                    groupShape.Tags.Add("id", param.id.ToString());
                    groupShape.Tags.Add("sub_id", param.sub_id.ToString());
                    //groupShape.Tags.Add("name", param.name);
                    groupShape.Name = param.name;
                    groupShape.Tags.Add("kind", param.kind);
                    return groupShape;
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine("-------[AddGroupShape] exception {0} ", e);
            }
            return null;

        }


        public PowerPoint.Shape AddPointShape(Point_Param param)
        {
            PowerPoint.Shape shp = null;
            switch (param.kind)
            {
                case "pin":
                    shp = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                        param.pos.X - param.width / 2,
                                    param.pos.Y - param.height / 2,
                                    param.width,
                                    param.height);
                    shp.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(0, 255, 0).ToArgb();
                    break;
                case "inPort":
                    //                     if (GetShape(param.pos, ptwidth) != null)
                    //                         return;
                    shp = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapePentagon,
                                    param.pos.X - param.width / 2,
                                    param.pos.Y - param.height / 2,
                                    param.width,
                                    param.height);
                    shp.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                    break;
                case "outPort":
                    //                     if (GetShape(param.pos, ptwidth) != null)
                    //                         return;
                    shp = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapePentagon,
                                    param.pos.X - param.width / 2,
                                    param.pos.Y - param.height / 2,
                                    param.width,
                                    param.height);
                    shp.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                    shp.Rotation = 180;
                    break;
                default:
                    return null;
            }
            if (!bfirst)
                shp.Fill.ForeColor.RGB = System.Drawing.Color.Orange.ToArgb();
            bfirst = false;
            shp.Tags.Add("xPoint", param.kind);

            shp.Tags.Add("id", param.id.ToString());
            shp.Tags.Add("kind", param.kind);
            shp.Name = param.name;
            //shp.Tags.Add("name", param.name);
            return shp;
        }
        // get Shape existing on point
        private PowerPoint.Shape GetShape(Point pt, int width)
        {
            try
            {
                List<PowerPoint.Shape> shpLst = new List<PowerPoint.Shape>();
                GetShapesInSlide(shpLst, ShapeTypeFlag.PIN | ShapeTypeFlag.OUTPORT | ShapeTypeFlag.INPORT, Globals.ThisAddIn.Application.ActiveWindow.View.Slide);
                foreach (PowerPoint.Shape shape in shpLst)
                {
                    if (shape.Tags.Count == 0) continue;
                    if (shape.Tags["xPoint"] != "")
                    {
                        float rot = shape.Rotation;
                        if (rot != 0)
                            shape.Rotation = 0;
                        if (shape.Left <= pt.X && pt.X <= shape.Left + width
                                && shape.Top <= pt.Y && pt.Y <= shape.Top + width)
                        {

                            if (rot != 0)
                                shape.Rotation = rot;
                            return shape;
                        }
                        if (rot != 0)
                            shape.Rotation = rot;
                    }
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine("-----[GetShape] exception - {0}", e);
            }
            return null;
        }

        private static void GetShapes(List<PowerPoint.Shape> shpLst, ShapeTypeFlag flag, Boolean bCurSlide = true)
        {
            try
            {
                if (bCurSlide)
                {
                    GetShapesInSlide(shpLst, flag, Globals.ThisAddIn.Application.ActiveWindow.View.Slide);
                }
                else
                {
                    foreach (Microsoft.Office.Interop.PowerPoint.Slide slide in Globals.ThisAddIn.Application.ActivePresentation.Slides)
                    {
                        GetShapesInSlide(shpLst, flag, slide);
                    }
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine("-----[GetShapesInSlide] exception - {0}", e);
            }
        }

        private static void GetShapesFlag(List<PowerPoint.Shape> shpLst, ShapeTypeFlag flag, PowerPoint.Shape shp)
        {
            if ((flag & ShapeTypeFlag.INPORT) != 0 && shp.Tags["xPoint"] == "inPort")
            {
                shpLst.Add(shp);
            }
            else if ((flag & ShapeTypeFlag.OUTPORT) != 0 && shp.Tags["xPoint"] == "outPort")
            {
                shpLst.Add(shp);
            }
            else if ((flag & ShapeTypeFlag.PIN) != 0 && shp.Tags["xPoint"] == "pin")
            {
                shpLst.Add(shp);
            }
        }


        private static void GetShapesInSlide(List<PowerPoint.Shape> shpLst, ShapeTypeFlag flag, PowerPoint.Slide slide)
        {
            try
            {
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Type == Office.MsoShapeType.msoGroup && shape.GroupItems.Count > 0)
                    {
                        foreach (Microsoft.Office.Interop.PowerPoint.Shape subShape in shape.GroupItems)
                        {
                            GetShapesFlag(shpLst, flag, subShape);
                        }
                        continue;
                    }
                    GetShapesFlag(shpLst, flag, shape);
                    //shape
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine("-----[GetShapesInSlide] exception - {0}", e);
            }
        }

        public static void ChangePointShapeSize(ShapeTypeFlag flag, int shapeWidth)
        {
            try
            {
                float shapeWidth2 = shapeWidth / 2.0f;
                List<PowerPoint.Shape> shpLst = new List<PowerPoint.Shape>();
                GetShapes(shpLst, flag, false);
                foreach (PowerPoint.Shape shape in shpLst)
                {
                    float rot = shape.Rotation;
                    if (rot != 0)
                        shape.Rotation = 0;
                    float orgWidth = shape.Width;
                    shape.Left += orgWidth / 2f - shapeWidth2;
                    shape.Top += orgWidth / 2f - shapeWidth2;
                    shape.Width = shapeWidth;
                    shape.Height = shapeWidth;
                    if (rot != 0)
                        shape.Rotation = rot;
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine("-----[ChangePointShapeSize] exception - {0}", e);
            }
        }
        public static void RedrawAllPoint()
        {
            //             for (int i = 0; i < ShpHistory.Count; i++)
            //             {
            //                 if (ShpHistory[i].AutoShapeType == Office.MsoAutoShapeType.msoShapeRectangle)
            //                 {
            //                     ShpHistory[i].Width = ptwidth;
            //                     ShpHistory[i].Height = ptwidth;
            //                     ShpHistory[i].Left = ShpHistory[i].Left + ptwidthDef/2;
            //                     ShpHistory[i].Top = ShpHistory[i].Top + ptwidthDef/2;
            //                     Point tmp = Point.Empty;
            //                     tmp.X= Convert.ToInt32(ShpHistory[i].Left); 
            //                     tmp.Y= Convert.ToInt32(ShpHistory[i].Top);
            //                     PtHistory[i] = tmp;
            //                 }                    
            //             }
        }
        public static void RedrawAllPort()
        {
            //             for (int i = 0; i < ShpHistory.Count; i++)
            //             {
            //                 if (ShpHistory[i].AutoShapeType == Office.MsoAutoShapeType.msoShapePentagon)
            //                 {
            //                     ShpHistory[i].Width = ptwidth;
            //                     ShpHistory[i].Height = ptwidth;
            //                     ShpHistory[i].Left = ShpHistory[i].Left + ptwidthDef/2;
            //                     ShpHistory[i].Top= ShpHistory[i].Top + ptwidthDef/2;
            //                     Point tmp = Point.Empty;
            //                     tmp.X = Convert.ToInt32(ShpHistory[i].Left);
            //                     tmp.Y = Convert.ToInt32(ShpHistory[i].Top);
            //                     PtHistory[i] = tmp;
            //                 }
            //             }
        }
        public static void Exporrtselectedobject()
        {
            List<PowerPoint.Shape> shplist = new List<PowerPoint.Shape>();
            PowerPoint.Application pptApplication = new PowerPoint.Application();
            Microsoft.Office.Interop.PowerPoint.Slides otherslides;
            Microsoft.Office.Interop.PowerPoint._Slide newslide;
            if (Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.Count > 0)
            {
                PowerPoint.ShapeRange selShprange = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                foreach (PowerPoint.Shape shp in Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange)
                {
                    if (shp.AutoShapeType != MsoAutoShapeType.msoShapeNotPrimitive)
                        shplist.Add(shp);
                }
            }
            // Create the Presentation File
            PowerPoint.Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoFalse);
            PowerPoint.CustomLayout newlayout = pptPresentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutText];
            otherslides = pptPresentation.Slides;
            newslide = otherslides.AddSlide(1, newlayout);
            foreach (PowerPoint.Shape Shpe in newslide.Shapes)
            {
                Shpe.Delete();
            }
            foreach (PowerPoint.Shape inst in shplist)
            {
                inst.Copy();
                newslide.Shapes.Paste();
            }
            pptPresentation.SaveAs(userlibpath, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            pptPresentation.Close();
        }

        public static void importuserlib(string filename)
        {
            List<PowerPoint.Shape> shplist = new List<PowerPoint.Shape>();
            PowerPoint.Application pptApplication = new PowerPoint.Application();
            Microsoft.Office.Interop.PowerPoint.Slides otherslides;
            Microsoft.Office.Interop.PowerPoint._Slide selslide;
            PowerPoint.Presentation pptPresentation = pptApplication.Presentations.Open(filename, MsoTriState.msoFalse, MsoTriState.msoCTrue, MsoTriState.msoFalse);
            PowerPoint.CustomLayout newlayout = pptPresentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutText];
            otherslides = pptPresentation.Slides;
            selslide = otherslides[1];
            foreach (PowerPoint.Shape shp in selslide.Shapes)
            {
                if (shp.AutoShapeType != MsoAutoShapeType.msoShapeNotPrimitive)
                    shplist.Add(shp);
            }
            foreach (PowerPoint.Shape inst in shplist)
            {
                inst.Copy();
                slide.Shapes.Paste();
            }
        }


        /// <summary>
        /// //////////////////////Drag&Drop/////////////////////////////
        /// </summary>
        public void SubscribeApplication()
        {
            Unsubscribe();
            Subscribe(Hook.AppEvents());
        }

        public void SubscribeGlobal()
        {
            Unsubscribe();
            Subscribe(Hook.GlobalEvents());
        }

        public void Subscribe(IKeyboardMouseEvents events)
        {
            m_Events = events;

            m_Events.MouseDragStarted += OnMouseDragStarted;
            m_Events.MouseDragFinished += OnMouseDragFinished;
            m_Events.MouseClick += OnMouseClick;
            m_Events.MouseMove += OnMouseMove;
            m_Events.MouseDown += OnMouseDown;
            m_Events.MouseUp += OnMouseUp;
            m_Events.MouseDoubleClick += OnMouseDblClick;
        }

        public void Unsubscribe()
        {
            if (m_Events == null) return;

            m_Events.MouseDragStarted -= OnMouseDragStarted;
            m_Events.MouseDragFinished -= OnMouseDragFinished;
            m_Events.MouseClick -= OnMouseClick;
            m_Events.MouseMove -= OnMouseMove;
            m_Events.MouseDown -= OnMouseDown;
            m_Events.MouseUp -= OnMouseUp;
            m_Events.MouseDoubleClick -= OnMouseDblClick;
            m_Events.Dispose();
            m_Events = null;
        }
        public void OnMouseDblClick(object sender, MouseEventArgs e)
        {
            try
            {

                if (bkeypress && State == 12)
                {
                    slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                    int sld_width = (int)slide.CustomLayout.Width;
                    int sld_height = (int)slide.CustomLayout.Height;
                    Point slidePT = ScreenPointToSlidePoint(System.Windows.Forms.Control.MousePosition);
                    if (slidePT.X < 0 || slidePT.Y < 0 || slidePT.X >= sld_width || slidePT.Y >= sld_height)
                    {
                        State = 0;
                        return;
                    }
                    AddWaveShape(null);
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("{0} Exception ", ex);
                return;
            }

        }
        public void OnMouseDragStarted(object sender, MouseEventArgs e)
        {
            isdragging = 1;
            Debug.WriteLine("-----[Mouse DragStart] {0}, {1} - {2} ", e.X, e.Y, mousepos);
        }

        public void OnMouseDown(object sender, MouseEventArgs e)
        {
            isdragging = 0;
            drag1pt = new Point(e.X, e.Y);
            Debug.WriteLine("-----[Mouse Down] {0}, {1} ", e.X, e.Y);
        }
        public void OnMouseUp(object sender, MouseEventArgs e)
        {
            isdragging = 0;
            drag1pt = new Point(e.X, e.Y);
            Debug.WriteLine("-----[Mouse Down] {0}, {1} ", e.X, e.Y);
        }
  
        public void OnMouseClick(object sender, MouseEventArgs e)
        {
            mousepos = System.Windows.Forms.Control.MousePosition;
            try
            {
                slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                if (e.Button.ToString() == "Left")
                {
                    int viewleft = (int)Globals.ThisAddIn.Application.ActiveWindow.Left;
                    int viewtop = (int)Globals.ThisAddIn.Application.ActiveWindow.Top;
                    int viewwidth = (int)Globals.ThisAddIn.Application.ActiveWindow.Width;
                    int screenwidth = Screen.PrimaryScreen.Bounds.Width;
                    int screenheight = Screen.PrimaryScreen.Bounds.Height;
                    int viewheight = (int)Globals.ThisAddIn.Application.ActiveWindow.Height;
                    int sld_width = (int)slide.CustomLayout.Width;
                    int sld_height = (int)slide.CustomLayout.Height;
//                     sld_width = (int)Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth;
//                     sld_height = (int)Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight;

                    int dif_x = viewwidth - sld_width;
                    int dif_y = viewheight - sld_height;

                    tmppoint.X = mousepos.X;
                    tmppoint.Y = mousepos.Y - dif_y;

                    if (tmppoint.X < 0 || tmppoint.Y < 0)
                    {
                        State = 0;
                        return;
                    }
                    if (viewleft > 0 || viewtop > 0)
                    {
                        State = 0;
                        return;
                    }

                    if (State == 1)
                    {
                        if (bkeypress)
                        {
                            Point slidePT = ScreenPointToSlidePoint(mousepos);
                            if (slidePT.X < 0 || slidePT.Y < 0 || slidePT.X >= sld_width || slidePT.Y >= sld_height )
                            {
                                return;
                            }
                            PowerPoint.Shape shp = AddPointShape(new Point_Param(slidePT));
                            if(shp != null)
                                MdbManger.GetInstance().Modify(shp, "add");

                            makeCorrectPosition(GetshapesSlide(),shp);
                            //                            GetshapesSlide();
                            bkeypress = false;
                            if (!bfirst)
                                shp.Fill.ForeColor.RGB = System.Drawing.Color.Orange.ToArgb();
                        }
                    }
                    else if (State == 2)
                    {
                         Point slidePT = ScreenPointToSlidePoint(mousepos);
                        if (slidePT.X < 0 || slidePT.Y < 0 || slidePT.X >= sld_width || slidePT.Y >= sld_height)
                        {
                            return;
                        }
                        if (checkLineable(slidePT))
                        {
                            DrawConnector();
                        }
                    }
                    else if (State == 3 || State == 4)
                    {

                        Point slidePT = ScreenPointToSlidePoint(mousepos);
                        if (slidePT.X < 0 || slidePT.Y < 0 || slidePT.X >= sld_width || slidePT.Y >= sld_height)
                        {
                            State = 0;
                            return;
                        }
                        if (bkeypress)
                        {
//                             drawPoints.Add(slidePT);
//                             DrawPort();
                            PowerPoint.Shape shp = AddPointShape(new Point_Param(slidePT,State == 3?"inPort":"outPort"));
                            makeCorrectPosition(GetshapesSlide(), shp);
                            if (shp != null)
                                MdbManger.GetInstance().Modify(shp, "add");
                            bkeypress = false;
                            if (!bfirst)
                                shp.Fill.ForeColor.RGB = System.Drawing.Color.Orange.ToArgb();
                        }

                    }
                    if (State == 6)
                    {

                        Point slidePT = ScreenPointToSlidePoint(mousepos);
                        if (slidePT.X < 0 || slidePT.Y < 0 || slidePT.X >= sld_width || slidePT.Y >= sld_height)
                        {
                            return;
                        }
                        if (bkeypress)
                        {
                            PowerPoint.Shape shp = AddGroupShape(new Group_Param(nCombitionalshp, slidePT, true));
                            if (shp != null)
                                MdbManger.GetInstance().Modify(shp, "add");
                            bkeypress = false;
                            //if (!bfirst)
                            //    shp.Fill.ForeColor.RGB = System.Drawing.Color.Orange.ToArgb();
                        }
            
                    }
                    if (State == 7)
                    {

                        Point slidePT = ScreenPointToSlidePoint(mousepos);
                        if (slidePT.X < 0 || slidePT.Y < 0 || slidePT.X >= sld_width || slidePT.Y >= sld_height)
                        {
                            return;
                        }
                        if (bkeypress)
                        {
                            PowerPoint.Shape shp = AddGroupShape(new Group_Param(Group_Param.GetKindStrFromSeqid(nSeqshp), slidePT, false));
                            if (shp != null)
                                MdbManger.GetInstance().Modify(shp, "add");
                            bkeypress = false;
                        }

                    }
                    if (State == 12)
                    {
                        Point slidePT = ScreenPointToSlidePoint(mousepos);
                        if (slidePT.X < 0 || slidePT.Y < 0 || slidePT.X >= sld_width || slidePT.Y >= sld_height)
                        {
                            return;
                        }
                        if (bkeypress)
                        {
                            if (bregMode)
                                slidePT = GetCorrectPoint(slidePT);             /// ths
                            DrawWaveLine(slidePT);
                        }
                
                    }
                }
            }
            catch(System.Exception ex)
            {
                Debug.WriteLine("{0} Exception ", ex);
                return;
            }
        }
        public Point GetCorrectPoint(Point pt)
        {
            if(bGridcheck)
            {
                int x = (pt.X / nGridSpace)* nGridSpace;
                if (pt.X - x > nGridSpace / 2)
                    pt.X = x + nGridSpace;
                else
                    pt.X = x;
                int y = (pt.Y / nGridSpace) * nGridSpace;
                if (pt.Y - y > nGridSpace / 2)
                    pt.Y = y + nGridSpace;
                else
                    pt.Y = y;
                //                 foreach (PowerPoint.Shape xLind in GridlistX)
                //                 {
                //                     //var x = xLind.Nodes[1].Points[1, 1];
                //                     
                //                 }
            }
            return pt;
        }
        private void OnMouseDragFinished(object sender, MouseEventArgs e)
        {
            /// remove ShpHistory PtHistory [
            //             isdragging = 2;
            //             PowerPoint.Selection Sel;
            //             drag2pt = mousepos;
            //             Sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            //             Point prePt = ScreenPointToSlidePoint(drag1pt);
            //             Point curPt = ScreenPointToSlidePoint(drag2pt);
            //             if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            //             {
            //                 PowerPoint.ShapeRange shapeRange = Sel.ShapeRange;
            //                 for (int sel = shapeRange.Count; sel > 0; sel--)
            //                 {
            //                     var selectedShape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[sel];
            //                     if (selectedShape.AutoShapeType != MsoAutoShapeType.msoShapeNotPrimitive)
            //                     {
            //                         if (selectedShape.Type == Office.MsoShapeType.msoGroup && selectedShape.GroupItems.Count > 0)
            //                         {
            //                             for (int gi = selectedShape.GroupItems.Count; gi > 0; gi--)
            //                             {
            //                                 var groupItem = selectedShape.GroupItems[gi];
            //                                 int ptindex = ShpHistory.IndexOf(groupItem);
            //                                 if (ptindex != -1)
            //                                 {
            //                                     Point shpTopLeft = new Point((int)groupItem.Left, (int)groupItem.Top);
            //                                     shpTopLeft.X += curPt.X - prePt.X;
            //                                     shpTopLeft.Y += curPt.Y - prePt.Y;
            //                                     PtHistory[ptindex] = shpTopLeft;
            //                                 }
            //                             }
            //                         }
            //                         else
            //                         {
            //                             int ptindex = ShpHistory.IndexOf(selectedShape);
            //                             if (ptindex != -1)
            //                             {
            //                                 Point shpTopLeft = new Point((int)selectedShape.Left, (int)selectedShape.Top);
            //                                 shpTopLeft.X += curPt.X - prePt.X;
            //                                 shpTopLeft.Y += curPt.Y - prePt.Y;
            //                                 PtHistory[ptindex] = shpTopLeft;
            //                             }
            //                         }
            //     
            //                     }
            //                 }
            //             }            
            //]
        }
        public void ResetPthistory()
        {
            /// remove ShpHistory PtHistory [
            //             List<Point> PtHistory1 = new List<Point>();
            //             for (int i = 0; i < ShpHistory.Count; i++)
            //             {
            //                 PtHistory1.Add(new Point((int)ShpHistory[i].Left, (int)ShpHistory[i].Top));
            //             }
            //             PtHistory = PtHistory1;
            //]
        }
        private void OnMouseMove(object sender, MouseEventArgs e)
        {
            //return;
            mousepos = System.Windows.Forms.Control.MousePosition;
            
            //slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            if (State == 2)
            {
                Point movePT = ScreenPointToSlidePoint(mousepos);
                curXpos = movePT.X;
                curYpos = movePT.Y;
                checkMoveLineable(movePT);
            }
            else if (state == 12)
            {
                if (mouseWaveLine != null)
                    mouseWaveLine.Delete();
                mouseWaveLine = null;
                if (mouseWaveCursor != null)
                    mouseWaveCursor.Delete();
                mouseWaveCursor = null;



                int sld_width = (int)Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth;
                int sld_height = (int)Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight;

                Point slidePT = ScreenPointToSlidePoint(mousepos);
                if (slidePT.X < 0 || slidePT.Y < 0 || slidePT.X >= sld_width || slidePT.Y >= sld_height)
                {
                   return;
                }
                if (bkeypress)
                {
                    if (bregMode)
                        slidePT = GetCorrectPoint(slidePT);             /// ths

                    if (wavePtCount > 0)
                    {
                        mouseWaveLine = slide.Shapes.AddLine(wavelnBeginpt.X, wavelnBeginpt.Y, slidePT.X, slidePT.Y);
                        mouseWaveLine.Tags.Add("sel", "unsel");
                        mouseWaveLine.GetType().InvokeMember("Locked", System.Reflection.BindingFlags.SetProperty, null, mouseWaveLine, new object[] { MsoTriState.msoTrue });
                    }

                    float cursorWidth = 13f*100f/ Globals.ThisAddIn.Application.ActiveWindow.View.Zoom;
                    float lineWeight = 2f * 100f / Globals.ThisAddIn.Application.ActiveWindow.View.Zoom;
                    PowerPoint.Shape shp1 = slide.Shapes.AddLine(
                        slidePT.X - cursorWidth / 2f, slidePT.Y,
                        slidePT.X + cursorWidth / 2f, slidePT.Y);
                    shp1.Line.Weight = lineWeight;
                    shp1.Line.ForeColor.RGB = System.Drawing.Color.Green.ToArgb();
                    PowerPoint.Shape shp2 = slide.Shapes.AddLine(
                        slidePT.X, slidePT.Y - cursorWidth / 2f,
                        slidePT.X, slidePT.Y + cursorWidth / 2f);
                    shp2.Line.Weight = lineWeight;
                    string[] names = { shp1.Name, shp2.Name };
                    shp2.Line.ForeColor.RGB = System.Drawing.Color.Green.ToArgb();
                    mouseWaveCursor = slide.Shapes.Range(names).Group();
                    mouseWaveCursor.GetType().InvokeMember("Locked", System.Reflection.BindingFlags.SetProperty, null, mouseWaveCursor, new object[] { MsoTriState.msoTrue });
                }

            }
        }
    }
    public class ScreenResolution
    {
        public static int BIT_PER_SEL=0;
        [DllImport("user32.dll")]
        public static extern bool EnumDisplaySettings(
              string deviceName, int modeNum, ref DEVMODE devMode);
       
        [DllImport("user32.dll")]
        public static extern int ChangeDisplaySettings(
            ref DEVMODE devMode, int flags);

        const int ENUM_CURRENT_SETTINGS = -1;

        const int ENUM_REGISTRY_SETTINGS = -2;
        

        [StructLayout(LayoutKind.Sequential)]
        public struct DEVMODE
        {

            private const int CCHDEVICENAME = 0x20;
            private const int CCHFORMNAME = 0x20;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 0x20)]
            public string dmDeviceName;
            public short dmSpecVersion;
            public short dmDriverVersion;
            public short dmSize;
            public short dmDriverExtra;
            public int dmFields;
            public int dmPositionX;
            public int dmPositionY;
            public ScreenOrientation dmDisplayOrientation;
            public int dmDisplayFixedOutput;
            public short dmColor;
            public short dmDuplex;
            public short dmYResolution;
            public short dmTTOption;
            public short dmCollate;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 0x20)]
            public string dmFormName;
            public short dmLogPixels;
            public int dmBitsPerPel;
            public int dmPelsWidth;
            public int dmPelsHeight;
            public int dmDisplayFlags;
            public int dmDisplayFrequency;
            public int dmICMMethod;
            public int dmICMIntent;
            public int dmMediaType;
            public int dmDitherType;
            public int dmReserved1;
            public int dmReserved2;
            public int dmPanningWidth;
            public int dmPanningHeight;
        }
        
       
        public static List<Rectangle> GetScreenresol()
        {
            DEVMODE vDevMode = new DEVMODE();
            int i = 0;
            List<Rectangle> scrRectlist = new List<Rectangle>();
            Rectangle rect = new Rectangle(0, 0, 0, 0);
            while (EnumDisplaySettings(null, i, ref vDevMode))
            {
                rect.Width = vDevMode.dmPelsWidth;
                rect.Height = vDevMode.dmPelsHeight;
                scrRectlist.Add(rect);
                BIT_PER_SEL = vDevMode.dmBitsPerPel;
                i++;
            }
            return scrRectlist;
        }
        public static void SetScreenresol(int width, int height)
        {
            DEVMODE vDevMode = new DEVMODE();
            int i = 0;
            List<Rectangle> scrRectlist = new List<Rectangle>();
            Rectangle rect = new Rectangle(0, 0, 0, 0);
            while (EnumDisplaySettings(null, i, ref vDevMode))
            {
                rect.Width = vDevMode.dmPelsWidth;
                rect.Height = vDevMode.dmPelsHeight;
                scrRectlist.Add(rect);
                BIT_PER_SEL = vDevMode.dmBitsPerPel;
                i++;
            }
            vDevMode.dmPelsWidth = width;
            vDevMode.dmPelsHeight = height;
            ChangeDisplaySettings(ref vDevMode, 0);
        }
    }

    public class Taskbar
    {
        [DllImport("user32.dll")]
        private static extern int FindWindow(string className, string windowText);

        [DllImport("user32.dll")]
        private static extern int ShowWindow(int hwnd, int command);

        [DllImport("user32.dll")]
        public static extern int FindWindowEx(int parentHandle, int childAfter, string className, int windowTitle);

        [DllImport("user32.dll")]
        private static extern int GetDesktopWindow();

        private const int SW_HIDE = 0;
        private const int SW_SHOW = 1;

        protected static int Handle
        {
            get
            {
                return FindWindow("Shell_TrayWnd", "");
            }
        }

        protected static int HandleOfStartButton
        {
            get
            {
                int handleOfDesktop = GetDesktopWindow();
                int handleOfStartButton = FindWindowEx(handleOfDesktop, 0, "button", 0);
                return handleOfStartButton;
            }
        }

        private Taskbar()
        {
            // hide ctor
        }

        public static void Show()
        {
            ShowWindow(Handle, SW_SHOW);
            ShowWindow(HandleOfStartButton, SW_SHOW);
        }

        public static void Hide()
        {
            ShowWindow(Handle, SW_HIDE);
            ShowWindow(HandleOfStartButton, SW_HIDE);
        }
    }
    //public static class Taskbar
    //{
    //    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    //    private static extern IntPtr FindWindow(
    //        string lpClassName,
    //        string lpWindowName);

    //    [DllImport("user32.dll", SetLastError = true)]
    //    private static extern int SetWindowPos(
    //        IntPtr hWnd,
    //        IntPtr hWndInsertAfter,
    //        int x,
    //        int y,
    //        int cx,
    //        int cy,
    //        uint uFlags
    //    );

    //    [Flags]
    //    private enum SetWindowPosFlags : uint
    //    {
    //        HideWindow = 128,
    //        ShowWindow = 64
    //    }

    //    public static void Show()
    //    {
    //        var window = FindWindow("Shell_traywnd", "");
    //        SetWindowPos(window, IntPtr.Zero, 0, 0, 0, 0, (uint)SetWindowPosFlags.ShowWindow);
    //    }

    //    public static void Hide()
    //    {
    //        var window = FindWindow("Shell_traywnd", "");
    //        SetWindowPos(window, IntPtr.Zero, 0, 0, 0, 0, (uint)SetWindowPosFlags.HideWindow);
    //    }
    //}
}
