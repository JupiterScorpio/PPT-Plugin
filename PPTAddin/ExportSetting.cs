using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ADOX; //Requires Microsoft ADO Ext. 2.8 for DDL and Security
//using ADODB;
using DAO;
using System.Data.OleDb;
using OLEDBError;
using stdole;
using System.Data.Odbc;
using System.IO;

namespace PPTAddin
{
    public partial class ExportSetting : Form
    {
        string strtmppath;
        string currentpath;
        string dbpath;
        public ExportSetting()
        {
            InitializeComponent();            
        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            currentpath = currentpath + "\\" + strtmppath+".pptx";
            ThisAddIn.userlibpath = currentpath;
            ThisAddIn.Exporrtselectedobject();
            AddnewObj(strtmppath);
            this.Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            strtmppath = textBox1.Text;
            btn_ok.Enabled = true;
        }

        private void ExportSetting_Load(object sender, EventArgs e)
        {
            //currentpath = System.IO.Directory.GetCurrentDirectory();
            currentpath= Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        }

        private void lib1_opt_CheckedChanged(object sender, EventArgs e)
        {
            string userlibpath= currentpath + "\\" + "UserLib1";
            currentpath = userlibpath;
            if(!System.IO.Directory.Exists(currentpath))
                System.IO.Directory.CreateDirectory(currentpath);
            dbpath = currentpath + "\\" + "1.mdb";            
        }

        private void lib2_opt_CheckedChanged(object sender, EventArgs e)
        {
            string userlibpath = currentpath + "\\" + "UserLib2";
            currentpath = userlibpath;
            if (!System.IO.Directory.Exists(currentpath))
                System.IO.Directory.CreateDirectory(currentpath);
            dbpath = currentpath + "\\" + "2.mdb";
        }
        public void AddnewObj(string objName)
        {
            //Microsoft.ACE.OLEDB.12.0
            //Microsoft.Jet.OLEDB.4.0
            OleDbConnection con = new OleDbConnection($"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {dbpath} ");
            OleDbCommand cmd = con.CreateCommand();
            con.Open();
            cmd.CommandText = "insert into liblTable (`Name`) values ('" + objName + "')";
            cmd.Connection = con;
            cmd.ExecuteNonQuery();
            con.Close();
        }
    }
}
