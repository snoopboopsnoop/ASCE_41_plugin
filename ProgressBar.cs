using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ASCE_41
{
    public partial class ProgressBar : Form
    {
        public ProgressBar()
        {
            InitializeComponent();
        }

        public ProgressBar(int total) : this()
        {
            progressBar1.Maximum = total;
        }

        public int Value()
        {
            return progressBar1.Value;
        }

        public int Maximum()
        {
            return progressBar1.Maximum;
        }

        public void Increment(int i = 1)
        {
            progressBar1.Increment(i);
            UpdateBar();
        }

        public void UpdateBar()
        {
            label1.Text = $"Generating Excel File ({(100 * (double)progressBar1.Value / progressBar1.Maximum).ToString("#.##")}%)";

            Invalidate();
        }

        public void SetCurrent(int i)
        {
            progressBar1.Value = i;
        }

        public void SetMaximum(int max)
        {
            progressBar1.Maximum = max;
        } 
    }
}
