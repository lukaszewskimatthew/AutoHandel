using System;
using System.Windows.Forms;

namespace AutoHandel
{
    public partial class GoogleAuth : Form
    {
        Form1 form;

        public GoogleAuth(Form1 inForm)
        {
            form = inForm;  
            InitializeComponent();           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            form.UserName = textBox1.Text;
            form.PassWord = textBox2.Text;
            button1.Enabled = false;
            this.Close();
        }

        private void GoogleAuth_Load(object sender, EventArgs e)
        {

        }
    }
}
