using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WordDocGenerator
{
    public partial class ProgressReport : Form
    {
        public ProgressReport()
        {
            InitializeComponent();
        }

        public void SetValue(int value)
        {
            this.progressBar1.Value = value;
        }
    }
}
