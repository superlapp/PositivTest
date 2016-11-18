using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PositivTest
{
    public partial class frmMain : Form
    {
        clsWorker cw = new clsWorker();
        //---------------------------------------------------------------------
        //---------------------------------------------------------------------
        //---------------------------------------------------------------------
        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            //
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            progressBar1.Style = ProgressBarStyle.Marquee;
            //
            var task = Task<List<string>>.Factory.StartNew(() => cw.Parse(Convert.ToInt32(pagesND.Value)));
            await task;
            var list = task.Result;
            //
            if (list.Count != 0)
            {
                listBox1.Items.Clear();
                foreach (string row in list)
                {
                    listBox1.Items.Add(row);
                }
                //
                cw.WriteToXLS(list);
            }
            //
            progressBar1.Style = ProgressBarStyle.Blocks;
            button1.Enabled = true;
        }
    }
}
