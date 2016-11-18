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

        private void button1_Click(object sender, EventArgs e)
        {
            ParseAsync();
        }

        private async void ParseAsync()
        {
            button1.Enabled = false;
            progressBar1.Style = ProgressBarStyle.Marquee;
            listBox1.Items.Clear();
            //
            try
            {
                var task = Task<List<string>>.Factory.StartNew(() => cw.Parse(Convert.ToInt32(pagesND.Value)));
                await task;
                var list = task.Result;
                //
                if (list.Count != 0)
                {
                    listBox1.Items.AddRange(list.ToArray());
                    cw.WriteToXLS(list);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                progressBar1.Style = ProgressBarStyle.Blocks;
                button1.Enabled = true;
            }
        }
    }
}