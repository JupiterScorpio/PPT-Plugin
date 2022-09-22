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
    public partial class CombinationalImport : Form
    {
        bool bclicked;
        public CombinationalImport()
        {
            InitializeComponent();
            
        }

        private void CombinationalImport_Load(object sender, EventArgs e)
        {
            bclicked = false;
            var imageList = new ImageList();
            /*   //imageList.Images.Add("aaa",(Icon)Properties.Resources.ResourceManager.GetObject("com_AND.png"));
               // tell your ListView to use the new image list
               listView1.View = View.SmallIcon;
               // add an item
               listView1.SmallImageList = imageList;*/
            //Image img = new Bitmap(Properties.Resources.com_AND);
            if(Globals.ThisAddIn.State==6)
            {
                imageList.ImageSize = new Size(100, 100);
                imageList.Images.Add(new Bitmap(Properties.Resources.com_AND));
                imageList.Images.Add(new Bitmap(Properties.Resources.com_BUFFER));
                imageList.Images.Add(new Bitmap(Properties.Resources.com_NAND));
                imageList.Images.Add(new Bitmap(Properties.Resources.com_NOR));
                imageList.Images.Add(new Bitmap(Properties.Resources.com_NOT));
                imageList.Images.Add(new Bitmap(Properties.Resources.com_OR));
                imageList.Images.Add(new Bitmap(Properties.Resources.com_XNOR));
                imageList.Images.Add(new Bitmap(Properties.Resources.com_XOR));
                this.listView1.View = View.LargeIcon;
                for (int i = 0; i < 8; i++)
                {
                    //ListViewItem lvt = new ListViewItem(i);
                    this.listView1.Items.Add(new ListViewItem { ImageIndex = i });
                    //listView1.Items.Add(lvt);
                }
                this.listView1.LargeImageList = imageList;
            }
            if (Globals.ThisAddIn.State == 7)
            {
                imageList.ImageSize = new Size(100, 100);
                imageList.Images.Add(new Bitmap(Properties.Resources.seq_D_flop));
                imageList.Images.Add(new Bitmap(Properties.Resources.seq_latch));
                imageList.Images.Add(new Bitmap(Properties.Resources.seq_synchronizer));
                this.listView1.View = View.LargeIcon;
                for (int i = 0; i < 3; i++)
                {
                    //ListViewItem lvt = new ListViewItem(i);
                    this.listView1.Items.Add(new ListViewItem { ImageIndex = i });
                    //listView1.Items.Add(lvt);
                }
                this.listView1.LargeImageList = imageList;
            }

        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            int selindex = -1;
            if(bclicked)
            {
                if (listView1.SelectedItems.Count > 0)
                {
                    selindex = listView1.Items.IndexOf(listView1.SelectedItems[0]);
                }
                if (Globals.ThisAddIn.State == 6)
                {
                    ThisAddIn.nCombitionalshp = Group_Param.GetGroupCombnationStr(selindex);
                }
                if(Globals.ThisAddIn.State==7)
                {
                    switch (selindex)
                    {
                        case 0:
                            ThisAddIn.nSeqshp = 0;
                            break;
                        case 1:
                            ThisAddIn.nSeqshp = 1;
                            break;
                        case 2:
                            ThisAddIn.nSeqshp = 2;
                            break;
                    }
                }                  
                this.Close();
            }            
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            bclicked = true;
        }
    }
}
