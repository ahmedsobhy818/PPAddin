using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PPAddin.Forms
{
    public partial class ProgressForm : Form
    {
        public ProgressForm()
        {
            InitializeComponent();
            //this.Height = 1065;
            //this.Width = 4305;
            //this.Left = 90;
            //this.Top = 405;
        }

        public void  SetProgress(int x)
        {
            progressBar1.Value = x;
            //progressBar1.Refresh();
            label1.Text = x + " % completed";
            //Application.DoEvents();
            Thread.Sleep(100);
        }
    }
}
