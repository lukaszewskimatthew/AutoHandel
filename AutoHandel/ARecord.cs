using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoHandel
{
    public partial class ARecord : UserControl
    {
        string date;
        string name;
        string stat;
        bool isChecked;

        public ARecord(string inDate, string inName, string inStat)
        {
            date = inDate;
            name = inName;
            stat = inStat;
            InitializeComponent();
        }

        private void ARecord_Load(object sender, EventArgs e)
        {
            textBox1.Text = date;
            textBox2.Text = name;
            textBox3.Text = stat;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            isChecked = checkBox1.Checked;
        }

        public bool IsChecked { get { return isChecked; } }
    }
}
