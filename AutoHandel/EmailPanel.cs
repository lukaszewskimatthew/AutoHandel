using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace AutoHandel
{
    public partial class EmailPanel : Form
    {
        OleDbDataReader reader;
        OleDbConnection conn;

        int y = 0;
        List<string> dists = new List<string>();

        public EmailPanel(OleDbConnection inConn)
        {
            conn = inConn;
            InitializeComponent();
        }

        private bool IsPresentIn(string inDist)
        {
            foreach (string dist in dists)
            {
                if (dist == inDist) { return true; }
            }
            return false;
        }

        private void EmailPanel_Load(object sender, EventArgs e)
        {
            panel1.AutoScroll = true;
            string gSchools = "Select Distract From [D&Emails]";
            OleDbCommand cmd = new OleDbCommand(gSchools, conn);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                if (!IsPresentIn(reader["Distract"].ToString()))
                {
                    dists.Add(reader["Distract"].ToString());
                    comboBox1.Items.Add(reader["Distract"].ToString());
                }
            }
            reader.Close();
            comboBox1.Items.Add("Add New");
            comboBox1.Text = comboBox1.Items[0].ToString();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Visible = false;
            if (comboBox1.Text == "Add New")
            {
                textBox1.Visible = true;
                return;
            }

            y = 0;
            panel1.Controls.Clear();
            string gEmails = "Select Distract, Type, Function, Email From [D&Emails] Where Distract = @dist";
            OleDbCommand cmd = new OleDbCommand(gEmails, conn);
            cmd.Parameters.AddWithValue("@dist", comboBox1.SelectedItem.ToString());
            reader = cmd.ExecuteReader();

            
            while (reader.Read())
            {
                DistEmail dEmail = new DistEmail(reader["Distract"].ToString(),
                                                    reader["Type"].ToString(),
                                                    reader["Function"].ToString(),
                                                    reader["Email"].ToString());
                dEmail.Location = new System.Drawing.Point(10, y * 60);
                panel1.Controls.Add(dEmail);
                y++;
            }
            reader.Close();
        }

        private void button1_Click(object sender, EventArgs e) //Add
        {
            DistEmail dEmail = new DistEmail(comboBox1.Text, "Fill In", "Fill In", "Fill In");
            dEmail.Location = new System.Drawing.Point(10, y * 60);
            panel1.Controls.Add(dEmail);
            y++;
        }

        private void button3_Click(object sender, EventArgs e) //Remove
        {         
                y = 0;

                string rEmail = "Delete From [D&Emails] Where Email = @email";
                OleDbCommand cmd;
                List<DistEmail> dEmails = new List<DistEmail>();

                foreach (DistEmail email in panel1.Controls)
                    if (email.IsChecked) { dEmails.Add(email); }

                foreach (DistEmail email in dEmails)
                {
                    cmd = new OleDbCommand(rEmail, conn);
                    cmd.Parameters.AddWithValue("@email", email.Email);
                    cmd.ExecuteNonQuery();
                    panel1.Controls.Remove(email);
                }

                foreach (DistEmail email in panel1.Controls)
                {
                    email.Location = new System.Drawing.Point(10, y * 60);
                    y++;
                }

                dEmails.Clear();           
        }

        private void button2_Click(object sender, EventArgs e) //Save
        {
            string rEmail = "Delete From [D&Emails] Where Distract = @dist";
            OleDbCommand cmd1 = new OleDbCommand(rEmail, conn);
            cmd1.Parameters.AddWithValue("@dist", comboBox1.Text);
            cmd1.ExecuteNonQuery();

            string cEmail = "Insert Into [D&Emails] (Distract, Type, Function, Email) Values " +
                            "(@dist, @type, @fun, @email)";
            OleDbCommand cmd2;

            foreach (DistEmail email in panel1.Controls)
            {
                if (!email.CheckFields())
                {
                    MessageBox.Show("All fileds must be filled in");
                    return;
                }

                cmd2 = new OleDbCommand(cEmail, conn);
                cmd2.Parameters.AddWithValue("@dist", comboBox1.Text);
                cmd2.Parameters.AddWithValue("@type", email.Type);
                cmd2.Parameters.AddWithValue("@fun", email.Function);
                cmd2.Parameters.AddWithValue("@email", email.Email);
                cmd2.ExecuteNonQuery();
                cmd2.Dispose();
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                comboBox1.Items.Add(textBox1.Text);
                comboBox1.SelectedItem = textBox1.Text;
                comboBox1.Items.Remove("Add New");
                comboBox1.Items.Add("Add New");
                textBox1.Text = "";
            }
        }
    }
}
