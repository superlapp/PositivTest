using System;
using System.Windows.Forms;

namespace PositivTest
{
    public partial class frmMain : Form
    {
        clsWorker cw = new clsWorker();

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
            var list = cw.Parse();
            if (list.Count != 0)
            {
                foreach (string row in list)
                {
                    listBox1.Items.Add(row);
                }
                //
                cw.WriteToXLS(list);
            }
        }
    }
}
