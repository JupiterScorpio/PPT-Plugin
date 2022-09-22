using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Diagnostics;

using ADOX; 

namespace PPTAddin
{
    public class Record
    {
        public Record(object key)
        {
            this.Key = key;
            Fields = new List<object>();
        }
        public object Key;
        public List<object> Fields;
        
    }
    public partial class UserLiborm : Form
    {
        string currentpath;         //current db folder path
        string dbpath;          //database path
        string strtmppath;              //object-file name
        string strfilepath;             //file path to read
        List<string> objlist = new List<string>();
        OleDbConnection con;
        int nSelectedIndex;
        public List<string> IDlist = new List<string>();
        public UserLiborm()
        {
            InitializeComponent();
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
  
        private void UserLiborm_Load(object sender, EventArgs e)
        {
            try
            {
                currentpath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                if (Globals.ThisAddIn.State == 8)
                {
                    dbpath = currentpath + "\\" + "UserLib1" + "\\" + "1.mdb";
                    currentpath = currentpath + "\\" + "UserLib1" + "\\";
                }
                if (Globals.ThisAddIn.State == 9)
                {
                    dbpath = currentpath + "\\" + "UserLib2" + "\\" + "2.mdb";
                    currentpath = currentpath + "\\" + "UserLib2" + "\\";
                }
                con = new OleDbConnection($"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {dbpath} ");
                OleDbCommand cmd = con.CreateCommand();
                con.Open();
                cmd.CommandText = "SELECT * FROM liblTable";
                cmd.Connection = con;
                //cmd.ExecuteNonQuery();
                var dict = new Dictionary<object, Record>();
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        object key = reader[0];
                        Record rec = new Record(key);
                        for (int i = 1; i < reader.FieldCount; i++)
                        {
                            rec.Fields.Add(reader[i]);
                        }
                        dict.Add(key, rec);
                        string str = rec.Fields[0].ToString();
                        listBox1.Items.Add(str);
                        objlist.Add(str);
                        IDlist.Add(key.ToString());
                    }
                }
                con.Close();
            }
            catch(Exception ex)
            {
                return;
            }

        }

        private void btn_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_rmv_Click(object sender, EventArgs e)
        {
            nSelectedIndex = listBox1.SelectedIndex;
            if (nSelectedIndex == -1)
                return;

            listBox1.Items.RemoveAt(nSelectedIndex);
            string delpath = currentpath + objlist[nSelectedIndex]+".pptx";
            File.Delete(delpath);

            con = new OleDbConnection($"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {dbpath} ");
            OleDbCommand cmd = con.CreateCommand();
            con.Open();
            string strcmd = "DELETE FROM liblTable WHERE ID = " + IDlist[nSelectedIndex] + ";";
            cmd.CommandText = strcmd;
            cmd.Connection = con;
            cmd.ExecuteNonQuery();

            objlist.RemoveAt(nSelectedIndex);
            IDlist.RemoveAt(nSelectedIndex);
        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            strfilepath= currentpath + strtmppath + ".pptx";
            ThisAddIn.importuserlib(strfilepath);
            this.Close();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            btn_ok.Enabled = true;
            btn_rmv.Enabled = true;
            if (listBox1.SelectedIndex == -1)
                return;
            strtmppath = objlist[listBox1.SelectedIndex];
        }
    }
}
