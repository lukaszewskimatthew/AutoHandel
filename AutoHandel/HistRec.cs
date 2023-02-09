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
    public partial class HistRec : Form
    {       
        OleDbConnection conn;
        OleDbDataReader reader;

        public HistRec(OleDbConnection inConn)
        {
            conn = inConn;
            InitializeComponent();
        }

        private void AttRec_Load(object sender, EventArgs e)
        {
            panel1.AutoScroll = true;
            string gStudents = "Select FName, LName, SDate, EDate From [Record]";
            OleDbCommand cmd = new OleDbCommand(gStudents, conn);
            reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                //if (DateTime.Parse(reader["SDate"].ToString()) <= DateTime.Now && DateTime.Now <= DateTime.Parse(reader["EDate"].ToString()).AddDays(1))
                comboBox1.Items.Add(reader["FName"].ToString() + " " + reader["LName"].ToString());
            }
            reader.Close();

            int y = 0;
            string getData = "Select Name, Status, Date From [Attendance] Where Name = @name";
            OleDbCommand cmd2;

            foreach (string name in comboBox1.Items)
            {
                cmd2 = new OleDbCommand(getData, conn);
                cmd2.Parameters.AddWithValue("@name", name);
                reader = cmd2.ExecuteReader();

                while (reader.Read())
                {
                    ARecord aRec = new ARecord(reader["Date"].ToString().Split(' ')[0], reader["Name"].ToString(), reader["Status"].ToString());
                    aRec.Location = new Point(52, 45 * y);
                    panel1.Controls.Add(aRec);
                    y += 1;
                }

                reader.Close();
                cmd2.Dispose();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int y = 0;
            panel1.Controls.Clear();
            string getAttRecs = "";

            if (radioButton1.Checked)
                getAttRecs = "Select Name, Status, Date From [Attendance] Where Name = @name";
            else if (radioButton2.Checked)
                getAttRecs = "Select SName, Date, Description From [Behavior] Where SName = @name";
            else
            {
                MessageBox.Show("Please select either attendance or behavior");
                return;
            }

            OleDbCommand cmd = new OleDbCommand(getAttRecs, conn);
            cmd.Parameters.AddWithValue("@name", comboBox1.SelectedItem.ToString());
            reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                ARecord aRec = new ARecord(null, null, null);
                if (radioButton1.Checked) 
                     aRec = new ARecord(reader["Date"].ToString().Split(' ')[0], reader["Name"].ToString(), reader["Status"].ToString());
                else
                    aRec = new ARecord(reader["Date"].ToString().Split(' ')[0], reader["SName"].ToString(), reader["Description"].ToString());
                aRec.Location = new Point(52, 45 * y);
                panel1.Controls.Add(aRec);
                y += 1;
            }
            reader.Close();
        }
    }
}
