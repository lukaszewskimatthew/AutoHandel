using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoHandel
{
    public partial class GSheet : Form
    {
        OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                         "Data Source=" + Application.StartupPath + @"\Students.accdb");
        public GSheet()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string update = "Update [General] Set URL = @url";
            OleDbCommand cmd = new OleDbCommand(update, conn);
            cmd.Parameters.AddWithValue("@url", textBox1.Text);
            cmd.ExecuteNonQuery();
            cmd.Dispose();
            this.Close();
        }

        private void GSheet_Load(object sender, EventArgs e)
        {
            conn.Open();
            string gPath = "Select URL From [General] Where ID = 1";

            OleDbCommand cmd = new OleDbCommand(gPath, conn);
            OleDbDataReader dbReader = cmd.ExecuteReader();
            while (dbReader.Read())
                textBox1.Text = dbReader["URL"].ToString();
            dbReader.Close();
        }

        private void GSheet_FormClosing(object sender, FormClosingEventArgs e)
        {
            conn.Close();
        }
    }
}
