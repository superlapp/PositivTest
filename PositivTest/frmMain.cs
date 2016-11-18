using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach (string r in cw.Parse())
            {
                listBox2.Items.Add(r);
            }
        }
    }
}
