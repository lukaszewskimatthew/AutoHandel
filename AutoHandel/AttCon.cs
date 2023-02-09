using System;
using System.Windows.Forms;

namespace AutoHandel
{
    public partial class AttCon : UserControl
    {
        string name;

        public AttCon(string inName)
        {
            name = inName;
            InitializeComponent();
        }

        private void AttCon_Load(object sender, EventArgs e)
        {
            textBox1.Text = name;
        }

        private void SpawnClick(object sender, EventArgs e)
        {
            base.OnClick(e);
        }

        public new string Name { get { return name; } }

        public string Stauts
        {
            get { return textBox2.Text; }
            set { textBox2.Text = value; }
        }

        public string Notes
        {
            get { return textBox3.Text; }
        }
    }
}
