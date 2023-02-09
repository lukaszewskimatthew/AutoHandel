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
    public partial class DistEmail : UserControl
    {
        string dist;
        string type;
        string fun;
        string email;
        bool isChecked;

        public DistEmail(string inDist, string inType, string inFun, string inEmail)
        {
            dist = inDist;
            type = inType;
            fun = inFun;
            email = inEmail;
            InitializeComponent();
        }

        private void DistEmail_Load(object sender, EventArgs e)
        {
            textBox1.Text = dist;
            comboBox2.Text = type;
            comboBox1.Text = fun;
            textBox3.Text = email;
        }

        private void SpawnClick(object sender, EventArgs e)
        {
            base.OnClick(e);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            isChecked = checkBox1.Checked;
        }

        public bool CheckFields()
        {
            if (textBox1.Text == "Fill In" || comboBox2.Text == "Fill In" || comboBox1.Text == "Fill In" || textBox3.Text == "Fill In")
                return false;
            else { return true; }
        }

        public string Type { get { return comboBox2.Text; } }

        public string Email { get { return textBox3.Text; } }

        public string Function { get { return comboBox1.Text; } }

        public bool IsChecked { get { return isChecked; } }
    }
}
