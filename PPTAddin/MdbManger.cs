using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Drawing;
using System.Diagnostics;
using ADOX;
using System.IO;
using System.Data.OleDb;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTAddin
{
    public class Point_Param
    {
        public static int pointId = 0;
        public static int pinWidth = 4;
        public static int portWidth = 4;
        public int id;
        public string name;
        public string kind;
        public PointF pos;
        public float width;
        public float height;
        public Point_Param()
        {
        }
        public Point_Param(Point pt, string kind = "pin")
        {
            this.id = Point_Param.pointId;
            Point_Param.pointId++;
            this.pos = pt;
            this.kind = kind;
            this.name = kind + this.id.ToString();
            if (kind == "pin")
            {
                this.width = Point_Param.pinWidth;
                this.height = Point_Param.pinWidth;
            }
            else
            {
                this.width = Point_Param.portWidth;
                this.height = Point_Param.portWidth;
            }

        }
        public Point_Param(PowerPoint.Shape shape)
        {
            id = int.Parse(shape.Tags["id"]);
            name = shape.Name;
            kind = shape.Tags["kind"];
            float rot = shape.Rotation;
            if (rot != 0)
                shape.Rotation = 0;
            pos.X = shape.Left + shape.Width / 2;
            pos.Y = shape.Top + shape.Height / 2;
            width = shape.Width;
            height = shape.Height;
            if (rot != 0)
                shape.Rotation = rot;
        }

    }

    public class Wire_Param
    {
        public static int wireId = 0;
        public int id;
        public string name;
        public string kind;
        public string strBeginPt = "";
        public string strEndPt = "";
        public PowerPoint.Shape beginShp = null;
        public PowerPoint.Shape endShp = null;

        //         public PointF pos;
        //         public float width;
        //         public float height;
        public Wire_Param()
        {

        }
        public Wire_Param(PowerPoint.Shape beginShape, PowerPoint.Shape endShape)
        {
            this.id = Wire_Param.wireId;
            Wire_Param.wireId++;
            this.kind = "wire";
            this.name = kind + this.id.ToString();
            this.beginShp = beginShape;
            this.endShp = endShape;
            if (beginShp != null)
                strBeginPt = beginShp.Name;
            if (endShp != null)
                strEndPt = endShp.Name;
        }

        public Wire_Param(PowerPoint.Shape shape)
        {
            id = int.Parse(shape.Tags["id"]);
            name = shape.Name;
            kind = shape.Tags["kind"];
            beginShp = shape.ConnectorFormat.BeginConnectedShape;
            endShp = shape.ConnectorFormat.EndConnectedShape;
            if (beginShp != null) 
                strBeginPt = beginShp.Name;
            if (endShp != null)
                strEndPt = endShp.Name;
        }

    }


    public class Group_Param
    {
        public static int groupId = 0;
        public static int CombnationalshpCnt = 0;
        public static int SeqshpCnt = 0;

        public Boolean isCombnation;

        public int id;
        public int sub_id;
        public string name;
        public string kind;
        public PointF pos;
        public float width;
        public float height;
        public string pin_names;
        public string label;
        public Group_Param()
        {

        }
        public Group_Param(PowerPoint.Shape shape)
        {
            id = int.Parse(shape.Tags["id"]);
            sub_id = int.Parse(shape.Tags["sub_id"]);
            name = shape.Name;
            kind = shape.Tags["kind"];
            float rot = shape.Rotation;
            if (rot != 0)
                shape.Rotation = 0;
            pos.X = shape.Left + shape.Width / 2;
            pos.Y = shape.Top + shape.Height / 2;
            width = shape.Width;
            height = shape.Height;
            this.label = shape.GroupItems[shape.GroupItems.Count].TextFrame.TextRange.Text;
            if (rot != 0)
                shape.Rotation = 0;
            isCombnation = IsCombination(kind);
            //pin_names = shape.Tags["pin_names"];
        }
        public Group_Param(string kind, Point pt, Boolean isCombnation)
        {
            this.id = Group_Param.groupId;
            Group_Param.groupId++;
            this.isCombnation = isCombnation;
            if (isCombnation)
            {
                this.sub_id = Group_Param.CombnationalshpCnt;
                this.label = "UC" + this.sub_id.ToString();
                Group_Param.CombnationalshpCnt++;
            }
            else
            {
                this.sub_id = Group_Param.SeqshpCnt;
                this.label = "US" + this.sub_id.ToString();
                Group_Param.SeqshpCnt++;

            }
            this.kind = kind;
            this.name = kind + this.sub_id;
            //this.pin_names = "";
            this.width = 0;
            this.height = 0;
            this.pos = pt;
            
        }
        public static Boolean IsGroup(string kind)
        {
            switch (kind)
            {
                case "and":
                    return true;
                case "buffer":
                    return true;
                case "nand":
                    return true;
                case "nor":
                    return true;
                case "not":
                    return true;
                case "or":
                    return true;
                case "xnor":
                    return true;
                case "xor":
                    return true;
                case "dflop":
                    return true;
                case "latch":
                    return true;
                case "sync":
                    return true;
            }
            return false;
        }
        public static Boolean IsCombination(string kind)
        {
            switch (kind)
            {
                case "dflop":
                    return false;
                case "latch":
                    return false;
                case "sync":
                    return false;
            }
            return true;
        }

        public static string GetKindStrFromSeqid(int seqIndex)
        {
            switch(seqIndex)
            {

                case 0://D-flop
                    return "dflop";
                case 1://latch
                    return "latch";
                case 2://synchronizer
                    return "sync";
            }
            return "";
        }
        public static string GetGroupCombnationStr(int index)
        {
            string combStr = "";
            switch (index)
            {
                case 0:
                    combStr = "and";
                    break;
                case 1:
                    combStr = "buffer";
                    break;
                case 2:
                    combStr = "nand";
                    break;
                case 3:
                    combStr = "nor";
                    break;
                case 4:
                    combStr = "not";
                    break;
                case 5:
                    combStr = "or";
                    break;
                case 6:
                    combStr = "xnor";
                    break;
                case 7:
                    combStr = "xor";
                    break;
            }
            return combStr;
        }

    }

    public class Wave_Param
    {
        public static int waveId = 0;
        public static int ptId = 0;
        public int id;
        public string name;
        public string kind;
        public int pt_count = 0;
        public string str_pts = "";
        public Wave_Param()
        {

        }
        public Wave_Param(PowerPoint.Shape shape)
        {
            this.id = int.Parse(shape.Tags["id"]);
            this.name = shape.Name;
            this.kind = shape.Tags["kind"];
            Point end_pt;
            Point begin_pt;
            foreach (PowerPoint.Shape ln in shape.GroupItems)
            {
                if(ln.Tags["kind"] == "waveline")
                {
                    if (ln.Tags["linemode"] == "1")
                    {
                        begin_pt = new Point((int)ln.Left, (int)ln.Top);
                        end_pt = new Point((int)(ln.Left + ln.Width), (int)(ln.Top + ln.Height));
                    }
                    else if (ln.Tags["linemode"] == "2")
                    {
                        begin_pt = new Point((int)ln.Left, (int)(ln.Top + ln.Height));
                        end_pt = new Point((int)(ln.Left + ln.Width), (int)ln.Top);
                    }
                    else if (ln.Tags["linemode"] == "3")
                    {
                        begin_pt = new Point((int)(ln.Left + ln.Width), (int)ln.Top);
                        end_pt = new Point((int)ln.Left, (int)(ln.Top + ln.Height));
                    }
                    else if (ln.Tags["linemode"] == "4")
                    {
                        begin_pt = new Point((int)(ln.Left + ln.Width), (int)(ln.Top + ln.Height));
                        end_pt = new Point((int)ln.Left, (int)ln.Top);
                    }
                    else
                    {
                        continue;
                    }
                    if (str_pts == "")
                    {
                        str_pts = $"{begin_pt.X},{begin_pt.Y}";
                    }
                    str_pts += $",{end_pt.X},{end_pt.Y}";
                }
                    
            }
        }
        public static void SetShapeTags(PowerPoint.Shape shape)
        {
            shape.Tags.Add("id", waveId.ToString());
            shape.Tags.Add("kind", "wave");
            shape.Name = "wave" + waveId.ToString();
            waveId++;
        }

    }

    public class MdbManger
    {
        private static MdbManger instance = null;
        public static MdbManger GetInstance()
        {
            if (instance == null)
                instance = new MdbManger();
            return instance;
        }

                        
        private static string GetTempMdbFilePath()
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\ppt_plugin\\";
            //path = "D:\\";
            string fileName = "temp.mdb";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            return path + fileName;
        }
        public Boolean CreateMDB()
        {
            string filePath = GetTempMdbFilePath();
            FileInfo fi = new FileInfo(filePath);
            if (fi.Exists)
                fi.Delete();

            ADOX.Catalog cat = new ADOX.Catalog();
            

            try
            {
                //string str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GetTempMdbFilePath();// + "; Jet OLEDB:Engine Type=5";
                string str = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + filePath;
                cat.Create(str);

                ADOX.Table nTable = new ADOX.Table();
                nTable.Name = "tblGlobal";
                nTable.Columns.Append("pin_width", DataTypeEnum.adInteger);
                nTable.Columns.Append("port_width", DataTypeEnum.adInteger);
                nTable.Columns.Append("grid_width", DataTypeEnum.adInteger);
                nTable.Columns.Append("grid_show", DataTypeEnum.adBoolean);
                cat.Tables.Append(nTable);

                nTable = new ADOX.Table();
                nTable.Name = "tblPoint";
                nTable.Columns.Append("id", DataTypeEnum.adInteger);
                nTable.Columns.Append("name", DataTypeEnum.adVarWChar, 255);
                nTable.Columns.Append("kind", DataTypeEnum.adVarWChar, 10);
                nTable.Columns.Append("x", DataTypeEnum.adSingle);
                nTable.Columns.Append("y", DataTypeEnum.adSingle);
                nTable.Columns.Append("width", DataTypeEnum.adSingle);
                nTable.Columns.Append("height", DataTypeEnum.adSingle);

                cat.Tables.Append(nTable);

                nTable = new ADOX.Table();
                nTable.Name = "tblWire";
                nTable.Columns.Append("id", DataTypeEnum.adInteger);
                nTable.Columns.Append("name", DataTypeEnum.adVarWChar, 255);
                nTable.Columns.Append("kind", DataTypeEnum.adVarWChar, 10);
                nTable.Columns.Append("pt1_name", DataTypeEnum.adVarWChar);
                nTable.Columns.Append("pt2_name", DataTypeEnum.adVarWChar);
                cat.Tables.Append(nTable);

                nTable = new ADOX.Table();
                nTable.Name = "tblGroup";
                nTable.Columns.Append("id", DataTypeEnum.adInteger);
                nTable.Columns.Append("sub_id", DataTypeEnum.adInteger);
                nTable.Columns.Append("name", DataTypeEnum.adVarWChar, 255);
                nTable.Columns.Append("kind", DataTypeEnum.adVarWChar, 10);
                nTable.Columns.Append("x", DataTypeEnum.adSingle);
                nTable.Columns.Append("y", DataTypeEnum.adSingle);
                nTable.Columns.Append("width", DataTypeEnum.adSingle);
                nTable.Columns.Append("height", DataTypeEnum.adSingle);
                nTable.Columns.Append("label", DataTypeEnum.adLongVarWChar,512);

                cat.Tables.Append(nTable);

                nTable = new ADOX.Table();
                nTable.Name = "tblWave";
                nTable.Columns.Append("id", DataTypeEnum.adInteger);
                nTable.Columns.Append("name", DataTypeEnum.adVarWChar, 255);
                nTable.Columns.Append("kind", DataTypeEnum.adVarWChar, 10);
                nTable.Columns.Append("str_pts", DataTypeEnum.adLongVarWChar);

                cat.Tables.Append(nTable);

                nTable = new ADOX.Table();
                nTable.Name = "tblTotal";
                nTable.Columns.Append("id", DataTypeEnum.adInteger);
                nTable.Columns.Append("name", DataTypeEnum.adVarWChar, 255);
                nTable.Columns.Append("kind", DataTypeEnum.adVarWChar, 10);
                nTable.Columns.Append("label", DataTypeEnum.adVarWChar, 255);

                cat.Tables.Append(nTable);

                cat = null;
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("-----[CreateMDB] excetion {0}", ex);
            }
         
            return false;
        }
        public Boolean SaveMDB()
        {
            OleDbConnection con;
            OleDbCommand cmd;
            try
            {
                string dbpath = GetTempMdbFilePath();
                FileInfo fi = new FileInfo(dbpath);
                //                 if (fi.Exists)
                //                     fi.Delete();
                //                 CreateMDB();
                if (!fi.Exists)
                    CreateMDB();

                con = new OleDbConnection($"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {dbpath} ");
                cmd = con.CreateCommand();
                con.Open();
            }
            catch(Exception e)
            {
                Debug.WriteLine("-----[SaveMDB] excetion {0}", e);
                return false;
            }
            try
            {
                cmd.Connection = con;
                cmd.CommandText = "DELETE FROM tblGlobal";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "DELETE FROM tblPoint";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "DELETE FROM tblGroup";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "DELETE FROM tblWire";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "DELETE FROM tblWave";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "DELETE FROM tblTotal";
                cmd.ExecuteNonQuery();
                int totalNum = 0;
                

                cmd.CommandText = "INSERT INTO tblGlobal(pin_width, port_width, grid_width, grid_show) " +
                    $"VALUES({Point_Param.pinWidth},{Point_Param.portWidth},{ThisAddIn.nGridSpace}, {ThisAddIn.bGridcheck.ToString()})";
                cmd.ExecuteNonQuery();


                foreach (Microsoft.Office.Interop.PowerPoint.Slide slide in Globals.ThisAddIn.Application.ActivePresentation.Slides)
                {
                   foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup && shape.GroupItems.Count > 0)
                        {
                            if (shape.Tags["kind"] == "wave")
                            {
                                Wave_Param wp = new Wave_Param(shape);
                                cmd.CommandText = "INSERT INTO tblWave(id, name, kind, str_pts) " +
                                    $"VALUES({wp.id},'{wp.name}', '{wp.kind}','{wp.str_pts}')";
                                cmd.ExecuteNonQuery();
                                cmd.CommandText = "INSERT INTO tblTotal(id, name, kind, label) " +
                                    $"VALUES({totalNum++},'{wp.name}', '{wp.kind}','')";
                                cmd.ExecuteNonQuery();
                            }
                            else if (Group_Param.IsGroup(shape.Tags["kind"]))
                            {
                                Group_Param gp = new Group_Param(shape);
                                cmd.CommandText = "INSERT INTO tblGroup(id, sub_id, name, kind, x, y, width, height, label) " +
                                    $"VALUES({gp.id},{gp.sub_id},'{gp.name}', '{gp.kind}',{gp.pos.X},{gp.pos.Y},{gp.width},{gp.height},'{gp.label}')";
                                cmd.ExecuteNonQuery();
                                cmd.CommandText = "INSERT INTO tblTotal(id, name, kind, label) " +
                                    $"VALUES({totalNum++},'{gp.name}', '{gp.kind}','{gp.label}')";
                                cmd.ExecuteNonQuery();


                                //                             if (shape.Tags["kind"] != "")
                                //                             {
                                // 
                                //                                 foreach (Microsoft.Office.Interop.PowerPoint.Shape subShape in shape.GroupItems)
                                //                                 {
                                //                                     if (subShape.Tags["kind"] == "pin" || subShape.Tags["kind"] == "inPort" || subShape.Tags["kind"] == "outPort")
                                //                                     {
                                //                                         Point_Param param = new Point_Param(subShape);
                                //                                         cmd.CommandText = "INSERT INTO tblPoint(id, name, kind, x, y, width, height) " +
                                //                                             $"VALUES({param.id},'{param.name}', '{param.kind}',{param.pos.X - gp.pos.X},{param.pos.Y - gp.pos.Y},{param.width},{param.height})";
                                //                                         cmd.ExecuteNonQuery();
                                //                                     }
                                //                                 }
                                //                             }

                            }
                        }
                        if(shape.Tags["kind"] == "pin" || shape.Tags["kind"] == "inPort" || shape.Tags["kind"] == "outPort")
                        {
                            Point_Param param = new Point_Param(shape);
                            cmd.CommandText = "INSERT INTO tblPoint(id, name, kind, x, y, width, height) " +
                                $"VALUES({param.id},'{param.name}', '{param.kind}',{param.pos.X},{param.pos.Y},{param.width},{param.height})";
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "INSERT INTO tblTotal(id, name, kind, label) " +
                                    $"VALUES({totalNum++},'{param.name}', '{param.kind}','')";
                            cmd.ExecuteNonQuery();
                        }
                        else if (shape.Tags["kind"] == "wire" )
                        {
                            Wire_Param param = new Wire_Param(shape);
                            cmd.CommandText = "INSERT INTO tblWire(id, name, kind, pt1_name, pt2_name) " +
                                $"VALUES({param.id},'{param.name}', '{param.kind}','{param.strBeginPt}','{param.strEndPt}')";
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "INSERT INTO tblTotal(id, name, kind, label) " +
                                    $"VALUES({totalNum++},'{param.name}', '{param.kind}','')";
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
                con.Close();
                return true;
            }
            catch (Exception e)
            {
                Debug.WriteLine("-----[SaveMDB] excetion {0}", e);
                con.Close();
            }
            finally
            {
            }
            return false;
        }
        public Boolean LoadMDB()
        {
            OleDbConnection con;
            try
            {
                string dbpath = GetTempMdbFilePath();
                FileInfo fi = new FileInfo(dbpath);
                if (!fi.Exists)
                    return false;
               
                con = new OleDbConnection($"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {dbpath} ");
                OleDbCommand cmd = con.CreateCommand();
                con.Open();
                cmd.Connection = con;


                cmd.CommandText = "SELECT * FROM tblGlobal";
                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        Point_Param.pinWidth = (int)reader["pin_width"];
                        Point_Param.portWidth = (int)reader["port_width"];
                        ThisAddIn.bGridcheck = bool.Parse(reader["grid_show"].ToString());
                        ThisAddIn.nGridSpace = (int)reader["grid_width"];
                        if( ThisAddIn.pointRibbon != null)
                        {
                            foreach(var item in ThisAddIn.pointRibbon.cmb_pinSZ.Items)
                            {
                                if (item.ToString() == Point_Param.pinWidth.ToString())
                                    ThisAddIn.pointRibbon.cmb_pinSZ.SelectedItem = item;
                            }
                            foreach (var item in ThisAddIn.pointRibbon.cmb_portSZ.Items)
                            {
                                if (item.ToString() == Point_Param.portWidth.ToString())
                                    ThisAddIn.pointRibbon.cmb_portSZ.SelectedItem = item;
                            }
                            
                            for (int i = 0; i< ThisAddIn.pointRibbon.gridSpaceAry.Length; i++)
                            {
                                if (ThisAddIn.pointRibbon.gridSpaceAry[i] == ThisAddIn.nGridSpace)
                                    ThisAddIn.pointRibbon.cmb_gridSZ.SelectedItem = ThisAddIn.pointRibbon.cmb_gridSZ.Items[i];
                            }
                            if(ThisAddIn.bGridcheck)
                            {
                                    ThisAddIn.DelGrid();
                                    ThisAddIn.DrawGrid();
                            }
                        }

                    }
                }

                cmd.CommandText = "SELECT * FROM tblPoint";

                

                Point_Param point_param = new Point_Param();
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        point_param.id = (int)reader["id"];
                        point_param.name = (string)reader["name"];
                        point_param.kind = (string)reader["kind"];
                        point_param.pos.X = (float)reader["x"];
                        point_param.pos.Y = (float)reader["y"];
                        point_param.width = (float)reader["width"];
                        point_param.height = (float)reader["height"];
                        if(point_param.kind == "pin")
                        {
                            point_param.width = Point_Param.pinWidth;
                            point_param.height = Point_Param.pinWidth;
                        }
                        else
                        {
                            point_param.width = Point_Param.portWidth;
                            point_param.height = Point_Param.portWidth;
                        }
                        if (Point_Param.pointId <= point_param.id)
                            Point_Param.pointId = point_param.id + 1;
                        Globals.ThisAddIn.AddPointShape(point_param);
                    }
                }

                cmd.CommandText = "SELECT * FROM tblGroup";
                cmd.Connection = con;

                Group_Param group_param = new Group_Param();
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        group_param.id = (int)reader["id"];
                        group_param.sub_id = (int)reader["sub_id"];
                        group_param.name = (string)reader["name"];
                        group_param.kind = (string)reader["kind"];
                        group_param.pos.X = (float)reader["x"];
                        group_param.pos.Y = (float)reader["y"];
                        group_param.width = (float)reader["width"];
                        group_param.height = (float)reader["height"];
                        group_param.label = (string)reader["label"];
                        group_param.isCombnation = Group_Param.IsCombination(group_param.kind);
                        if (Group_Param.groupId <= group_param.id)
                            Group_Param.groupId = group_param.id + 1;
                        if (group_param.isCombnation)
                        {
                            if (Group_Param.CombnationalshpCnt <= group_param.sub_id)
                                Group_Param.CombnationalshpCnt = group_param.sub_id + 1;
                        }
                        else
                        {
                            if (Group_Param.SeqshpCnt <= group_param.sub_id)
                                Group_Param.SeqshpCnt = group_param.sub_id + 1;
                        }

                        Globals.ThisAddIn.AddGroupShape(group_param);
                    }
                }

                cmd.CommandText = "SELECT * FROM tblWire";
                cmd.Connection = con;
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        Wire_Param wire_param = new Wire_Param();
                        wire_param.id = (int)reader["id"];
                        wire_param.name = (string)reader["name"];
                        wire_param.kind = (string)reader["kind"];
                        wire_param.strBeginPt = (string)reader["pt1_name"];
                        wire_param.strEndPt = (string)reader["pt2_name"];
                        if (Wire_Param.wireId <= group_param.id)
                            Wire_Param.wireId = group_param.id + 1;
                        Globals.ThisAddIn.AddWireShape(wire_param);
                    }
                }

                cmd.CommandText = "SELECT * FROM tblWave";
                Wave_Param wave_param = new Wave_Param();
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        wave_param.id = (int)reader["id"];
                        wave_param.name = (string)reader["name"];
                        wave_param.kind = (string)reader["kind"];
                        wave_param.str_pts = (string)reader["str_pts"];
                        if (Wave_Param.waveId <= wave_param.id)
                            Wave_Param.waveId = wave_param.id + 1;
                        PowerPoint.Shape shp = Globals.ThisAddIn.AddWaveShape(wave_param);
                    }
                }
                con.Close();
                return true;
            }
            catch (Exception e)
            {
                Debug.WriteLine("-----[SaveMDB] excetion {0}", e);
            }
            finally
            {

            }
            return false;
        }
        public void Modify(PowerPoint.Shape shape, string modifyflag)
        {
            SaveMDB();
        }
    }
}

