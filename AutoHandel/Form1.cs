using System;
using System.IO;
using System.IO.Compression;
using static System.IO.Compression.ZipArchive;
using System.Net;
using System.Text;
using System.Collections.Generic;

using Spire.Xls;
using System.Drawing;
using System.Windows.Forms;
using Google.Apis.Drive.v2;
using System.Data;
using System.Data.OleDb;
using System.Drawing.Printing;
using System.Threading;
using System.Diagnostics;
using System.Net.Mail;
using System.Linq;

namespace AutoHandel  
{
    public partial class Form1 : Form
    {
        OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;" +
                         "Data Source=" + Application.StartupPath + @"\Students.accdb");
        OleDbDataReader dbReader;

        string userName = "", passWord = "";
        List<int> newStudents = new List<int>();
        List<Panel> windows = new List<Panel>();
        List<Label> names = new List<Label>();
        List<Label> sStates = new List<Label>();
        List<TextBox> notes = new List<TextBox>();
        List<Lunch> sLunches = new List<Lunch>();

        public Form1()
        {
            InitializeComponent();
        }

        private void SetAllPanels(Panel inPanel)
        {
            foreach (Panel p in windows)
            {
                if (p == inPanel) { p.Visible = true; }
                else { p.Visible = false; }
            }
        }

        private OleDbDataReader MakeReader(string cStr) //Set up Database Reader (Simple)
        {
            if (conn.State == ConnectionState.Closed) { conn.Open(); }
            OleDbCommand cmd = new OleDbCommand(cStr, conn);
            OleDbDataReader reader = cmd.ExecuteReader();
            return reader;
        }

        public static bool WebRequestTest() //Network Connection Test
        {
            try
            {
                WebRequest myRequest = WebRequest.Create("http://www.google.com");
                WebResponse myResponse = myRequest.GetResponse();
            }

            catch (WebException) { return false; }
            return true;
        }

        public void Download() //Get Student List From Google Docs
        {
            string sheet = "";
            string gLink = "Select URL From [General]";
            OleDbCommand cmd = new OleDbCommand(gLink, conn);
            dbReader = cmd.ExecuteReader();

            while (dbReader.Read())
                sheet = dbReader["URL"].ToString();
            dbReader.Close();
            cmd.Dispose();

            try
            {
                Uri link = new Uri(sheet);
                DriveService service = new DriveService();

                var stream = service.HttpClient.GetStreamAsync(link);
                var result = stream.Result;
                using (var fileStream = File.Create(Application.StartupPath + @"\GrabStudent.xlsx"))
                {
                    result.CopyTo(fileStream); //Download and Save
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            

            if (!File.Exists(Application.StartupPath + @"\GrabStudent.xlsx"))
                MessageBox.Show("File not gathered correctly");
        }

        private bool IsPresentIn(int searchNum)
        {
            foreach (int i in newStudents)
            {
                if (i == searchNum) { return true; }
            }
            return false;
        }

        private int GetLines(string path)
        {
            string line;
            int lineCount = 0;
            using (var reader = File.OpenText(@path))
            {
                while ((line = reader.ReadLine()) != null)
                {
                    if (line == ",,,,,,,,,,,,,,,,,,,,,,,,,") { break; }
                    lineCount++;
                }
            }
            return lineCount - 1; //Remove Header
        }

        private int GetRecordNumber(string table) // Next Record Number
        {
            string s = "";
            if (table == "Record") { s = "SELECT COUNT(*) FROM Record"; }
            else if (table == "Attendance") { s = "SELECT COUNT(*) FROM Attendance"; }
            else if (table == "WorkLog") { s = "SELECT COUNT(*) FROM WorkLog"; }
            else if (table == "Check") { s = "SELECT COUNT(*) FROM [Check]"; }
            else if (table == "Behavior") { s = "SELECT COUNT(*) FROM [Behavior]"; }

            if (s != "")
            {
                OleDbCommand count = new OleDbCommand(s, conn);
                return (int)count.ExecuteScalar() + 1;
            }
            else { return -1; }
        }

        private List<string> GetCurrNames() //Names in dtb
        {
            List<string> names = new List<string>();
            string getNames = "Select FName, LName, SDate, EDate From [Record]";
            OleDbCommand cmd = new OleDbCommand(getNames, conn);
            dbReader = cmd.ExecuteReader();

            while (dbReader.Read())
            {
                names.Add(dbReader["FName"].ToString() + " " + dbReader["LName"].ToString());
            }
            dbReader.Close();
            cmd.Dispose();

            return names;
        }

        private string RemoveWhiteSpace(string input)
        {
            StringBuilder output = new StringBuilder(input.Length);

            for (int index = 0; index < input.Length; index++)
            {
                if (!Char.IsWhiteSpace(input, index))
                {
                    output.Append(input[index]);
                }
            }
            return output.ToString();
        }

        private void MakeStudent(string line) //Record a new student (parse date and time)
        {
            string[] args = new string[] { "@id", "@time", "@email", "@fname", "@lname",
                                           "@grade", "@dist", "@pfname", "@plname", "@ppn",
                                           "@lelg", "@ssneeds", "@workcont", "@workpro", "@sdate",
                                           "@edate", "@incid", "@shear", "@stip", "@offssec" };

            string mkRecord = "Insert Into [Record] Values (";
            foreach (string arg in args)
            {
                if (arg == "@offssec") { mkRecord += arg; }
                else { mkRecord += (arg + ", "); }
            }
            mkRecord += ")";         

            string[] vals = line.Split('#');
            for (int x = 0; x < vals.Length; x++) { vals[x] = vals[x].Replace(@"""", ""); }

            int wSpaces = 0;
            foreach (string val in vals)
            {
                if (val == "") { wSpaces++; }
            }
            if (wSpaces == vals.Length) { return; }

            try
            {
                OleDbCommand cmd = new OleDbCommand(mkRecord, conn);
                for (int i = 0; i < args.Length; i++)
                {
                    if (i == 0)
                    {
                        int rNum = GetRecordNumber("Record");
                        cmd.Parameters.AddWithValue(args[i], rNum);

                        newStudents.Add(rNum);
                    }

                    else if (i == 1) { cmd.Parameters.AddWithValue(args[1], vals[0]); }

                    else if (i == 2)
                    {
                        string email = "";
                        string[] parts = vals[12].Split(' ');
                        foreach (string part in parts)
                        {
                            char[] chars = part.ToCharArray();
                            foreach (char c in chars)
                            {
                                if (c == '@') { email = part; }
                            }
                        }

                        for (int x = 0; x < email.Length; x++) { email = email.Replace(",", ""); }
                        cmd.Parameters.AddWithValue(args[2], email);
                    }

                    else if (i == 3 || i == 4)
                    {
                        cmd.Parameters.AddWithValue(args[i], RemoveWhiteSpace(vals[i]));
                    }

                    else if (i == 14 || i == 15)
                    {
                        DateTime time = new DateTime();
                        string[] parts = vals[i].Split(' ');
                        foreach (string part in parts)
                        {
                            if (DateTime.TryParse(part, out time)) { break; }
                        }
                        cmd.Parameters.AddWithValue(args[i], time.ToShortDateString());
                    }

                    else { cmd.Parameters.AddWithValue(args[i], vals[i]); }
                }

                cmd.ExecuteNonQuery();
                cmd.Dispose();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            foreach (Control control in this.Controls)
            {
                control.Anchor = (AnchorStyles.Top | AnchorStyles.Left);
            }

            foreach (Control control in this.Controls)
            {
                if (typeof(Panel) == control.GetType())
                {
                    if (control == panel3 || control == panel6 || control == panel8) { }

                    else { windows.Add((Panel)control); }
                }                   
            }

            conn.Open();
            SetAllPanels(panel1);
            if (!WebRequestTest()) { MessageBox.Show("Your network connection is down"); }

            listView1.View = View.Details;
            listView1.Columns.Add("Name", 150); //# gives autosize props
            listView1.Columns.Add("Start", 75);
            listView1.Columns.Add("End", 75);

            listView2.View = View.Details;
            listView2.Columns.Add("Name", 200);
            listView2.Columns.Add("Status", 100);

            listView3.View = View.Details;
            listView3.Columns.Add("Name", 150); //# gives autosize props
            listView3.Columns.Add("Start", 75);
            listView3.Columns.Add("End", 75);

            string[] subs = new string[] { "English", "Math", "Social Studies",
                                            "Science", "Gym", "Art", "Other" };
            foreach (string sub in subs) { comboBox1.Items.Add(sub); }

            dbReader = MakeReader("Select [ID], [FName], [LName], [SDate], [EDate] From Record");
            while (dbReader.Read())
            {
                if (DateTime.Parse(dbReader["SDate"].ToString()) <= DateTime.Now && DateTime.Now <= DateTime.Parse(dbReader["EDate"].ToString()).AddDays(1))
                {
                    string name = dbReader["Fname"].ToString() + " " + dbReader["LName"].ToString();
                    string[] sparts = (dbReader["SDate"].ToString()).Split(' ');
                    string[] eparts = (dbReader["EDate"].ToString()).Split(' ');

                    listView1.Items.Add(new ListViewItem(new string[] { name, sparts[0], eparts[0] }));
                }
            }
            dbReader.Close();
         }

        private void generalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetAllPanels(panel1);
        } 

        private void button1_Click(object sender, EventArgs e)
        {           
            try
            {                
                string sheet = "";
                string gLink = "Select URL From [General]";
                OleDbCommand cmd = new OleDbCommand(gLink, conn);
                dbReader = cmd.ExecuteReader();

                while (dbReader.Read())
                    sheet = dbReader["URL"].ToString();
                dbReader.Close();
                cmd.Dispose();

                Uri link = new Uri(sheet);
                DriveService service = new DriveService();

                var stream = service.HttpClient.GetStreamAsync(link);
                var result = stream.Result;
                using (var fileStream = File.Create(Application.StartupPath + @"\GrabStudent.xlsx"))
                {
                    result.CopyTo(fileStream); //Download and Save
                }

                if (!File.Exists(Application.StartupPath + @"\GrabStudent.xlsx"))
                {
                    MessageBox.Show("File not gathered correctly");
                    return;
                }

                string rAll = "Delete * From Record";
                OleDbCommand cmd2 = new OleDbCommand(rAll, conn);
                cmd2.ExecuteNonQuery();
                cmd2.Dispose();

                Workbook workbook = new Workbook();
                workbook.LoadFromFile("GrabStudent.xlsx");
                Worksheet wSheet = workbook.Worksheets[0];
                wSheet.SaveToFile(Application.StartupPath + "EasyReadStudents.txt", "#", Encoding.UTF8);


                string line;
                int lNum = 0;

                progressBar1.Minimum = 0;
                progressBar1.Maximum = GetLines(Application.StartupPath + "EasyReadStudents.txt");
                progressBar1.Step = 1;
                
                using (StreamReader readStudent = new StreamReader(Application.StartupPath + "EasyReadStudents.txt"))
                {
                    while ((line = readStudent.ReadLine()) != null)
                    {
                        if (lNum == 0) { lNum++; continue; }
                        MakeStudent(line);
                        progressBar1.PerformStep();
                        lNum++;
                    }
                }
                
            }
            catch (Exception ex) { MessageBox.Show("Error found in reading studnets\n" + ex.Message); }

            File.Delete(Application.StartupPath + "EasyReadStudents.txt");

            listView1.Items.Clear();
            int row = 0;
            dbReader = MakeReader("Select [ID], [FName], [LName], [SDate], [EDate] From Record");
            while (dbReader.Read())
            {
                if (DateTime.Parse(dbReader["SDate"].ToString()) <= DateTime.Now && DateTime.Now <= DateTime.Parse(dbReader["EDate"].ToString()).AddDays(1))
                {
                    string name = dbReader["Fname"].ToString() + " " + dbReader["LName"].ToString();
                    string[] sparts = (dbReader["SDate"].ToString()).Split(' ');
                    string[] eparts = (dbReader["EDate"].ToString()).Split(' ');

                    listView1.Items.Add(new ListViewItem(new string[] { name, sparts[0], eparts[0] }));
                    if (!IsPresentIn(int.Parse(dbReader["ID"].ToString())))
                    {
                        listView1.Items[row].ForeColor = Color.Red;
                    }
                    row++;
                }
            }
            dbReader.Close();
            button1.Enabled = false;

            try
            {
                OleDbCommand cmd1 = new OleDbCommand("Select FName, LName From [Record]", conn);
                dbReader = cmd1.ExecuteReader();

                int i = 0;
                while (dbReader.Read())
                {
                    i = 0;
                    OleDbDataReader dbReader2;
                    OleDbCommand cmd2 = new OleDbCommand("Select ID From [Check] Where SName = @name", conn);
                    cmd2.Parameters.AddWithValue("@name", dbReader["FName"] + " " + dbReader["LName"]);
                    dbReader2 = cmd2.ExecuteReader();
                    while (dbReader2.Read()) { i++; }

                    if (i == 0)
                    {
                        OleDbCommand cmd3 = new OleDbCommand("Insert Into [Check] Values (@id, @sName, @profile, @mInfo, @pif, @work)", conn);
                        cmd3.Parameters.AddWithValue("@id", GetRecordNumber("Check"));
                        cmd3.Parameters.AddWithValue("@sName", dbReader["FName"] + " " + dbReader["LName"]);
                        cmd3.Parameters.AddWithValue("@profile", false);
                        cmd3.Parameters.AddWithValue("@mInfo", false);
                        cmd3.Parameters.AddWithValue("@pif", false);
                        cmd3.Parameters.AddWithValue("@work", false);
                        cmd3.ExecuteNonQuery();
                    }
                    dbReader2.Close();

                    bool fFound = false;
                    foreach (string file in Directory.GetFiles(Application.StartupPath))
                        if (file == (Application.StartupPath + dbReader["FName"] + " " + dbReader["LName"])) { fFound = true; }

                    if (!fFound)
                        Directory.CreateDirectory(Application.StartupPath + "/" + dbReader["FName"] + " " + dbReader["LName"]);
                }
                dbReader.Close();
            }
            catch (Exception ex) { MessageBox.Show("Error found in setting up students\n" + ex.Message); }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e) //General Student Info
        {
            if (listView1.SelectedItems.Count > 0)
            {         
                listBox1.Items.Clear();
                string name = listView1.SelectedItems[0].Text;
                string[] parts = name.Split(' ');

                string fStudent = "Select * From [Record] Where " +
                      "(FName = @fname) And (LName = @lname)";

                OleDbCommand cmd = new OleDbCommand(fStudent, conn);
                cmd.Parameters.AddWithValue("@fname", parts[0]);
                cmd.Parameters.AddWithValue("@lname", parts[1]);
                dbReader = cmd.ExecuteReader();
                while (dbReader.Read())
                {
                    textBox1.Text = dbReader["Fname"].ToString() + " " + dbReader["LName"].ToString();
                    textBox2.Text = dbReader["Distract"].ToString();
                    textBox3.Text = dbReader["Grade"].ToString();
                    textBox4.Text = dbReader["WorkContact"].ToString();
                    textBox5.Text = dbReader["CEmail"].ToString();
                    textBox6.Text = dbReader["PFName"].ToString() + " " + dbReader["PLName"].ToString();
                    textBox7.Text = dbReader["PPN"].ToString();

                    textBox8.Text = dbReader["Incident"].ToString();
                    textBox9.Text = dbReader["SupHearing"].ToString();
                    textBox10.Text = dbReader["Stipulations"].ToString();
                    textBox11.Text = dbReader["OffbySSEC"].ToString();
                    break;
                }
                dbReader.Close();
                cmd.Dispose();

                string getRecord = "Select * From [Check] Where SName = @name";
                OleDbCommand cmd2 = new OleDbCommand(getRecord, conn);
                cmd2.Parameters.AddWithValue("@name", name);
                dbReader = cmd2.ExecuteReader();

                while (dbReader.Read())
                {
                    checkBox1.Checked = bool.Parse(dbReader["Profile"].ToString());
                    checkBox2.Checked = bool.Parse(dbReader["MInfo"].ToString());
                    checkBox3.Checked = bool.Parse(dbReader["PIF"].ToString());
                    checkBox4.Checked = bool.Parse(dbReader["Work"].ToString());
                }
                dbReader.Close();
                cmd2.Dispose();
            }
        }

        private void button9_Click(object sender, EventArgs e) //Checklist save
        {
            if (listView1.SelectedItems.Count > 0)
            {
                string uRecord = "Update [Check] Set [Profile] = @rProf, [MInfo] = @rMed, [PIF] = @rPIF, " +
                                                     "[Work] = @rWork Where [SName] = @name";
                OleDbCommand cmd = new OleDbCommand(uRecord, conn);
                cmd.Parameters.AddWithValue("@rProf", checkBox1.Checked);
                cmd.Parameters.AddWithValue("@rMed", checkBox2.Checked);
                cmd.Parameters.AddWithValue("@rPIF", checkBox3.Checked);
                cmd.Parameters.AddWithValue("@rWork", checkBox4.Checked);
                cmd.Parameters.AddWithValue("@name", listView1.SelectedItems[0].Text);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0) //Test and handel error
            {
                listBox1.Items.Clear();
                string getWork = "Select WName From [WorkLog] Where " +
                                 "(SName = @name) And (Subject = @sub)";
                OleDbCommand cmd = new OleDbCommand(getWork, conn);
                cmd.Parameters.AddWithValue("@name", listView1.SelectedItems[0].Text);
                cmd.Parameters.AddWithValue("@sub", comboBox1.Text);
                dbReader = cmd.ExecuteReader();

                while (dbReader.Read())
                {
                    listBox1.Items.Add(dbReader["WName"].ToString());
                }
                dbReader.Close();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && listView1.SelectedItems.Count > 0)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();

                openFileDialog1.InitialDirectory = "c:\\";
                openFileDialog1.Filter = "(*.*)|*.*";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;

                string sourcePath = @"";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    sourcePath = openFileDialog1.FileName;
                    //MessageBox.Show(sourcePath);

                    string[] parts = sourcePath.Split('\\');
                    int index = parts.Length;
                    string targetPath = Application.StartupPath + @"\" +
                        listView1.SelectedItems[0].Text + @"\" + parts[index - 1];

                    try { File.Copy(sourcePath, targetPath); }
                    catch (IOException) { MessageBox.Show("File already added"); return; }

                    string[] fileParts = parts[index - 1].Split('.');
                    string makeWork = "Insert Into [WorkLog] Values(@id, @name, @wname, @status, " +
                                      "@subject, @path, @comPath)";
                    OleDbCommand cmd = new OleDbCommand(makeWork, conn);
                    cmd.Parameters.AddWithValue("@id", GetRecordNumber("WorkLog"));
                    cmd.Parameters.AddWithValue("@name", listView1.SelectedItems[0].Text);
                    cmd.Parameters.AddWithValue("@wname", fileParts[0]);
                    cmd.Parameters.AddWithValue("@status", "Incomplete");
                    cmd.Parameters.AddWithValue("@subject", comboBox1.Text);
                    cmd.Parameters.AddWithValue("@path", targetPath);
                    cmd.Parameters.AddWithValue("@comPath", "NC");
                    cmd.ExecuteNonQuery();

                    listBox1.Items.Add(fileParts[0]);
                }
            }

            else { MessageBox.Show("Please select a student and a subject"); }
        }

        private void TextBox_Click(object sender, EventArgs e)
        {
            string action = "";
            if (radioButton1.Checked) { action = "Present"; }
            else if (radioButton2.Checked) { action = "Absent"; }
            else if (radioButton3.Checked) { action = "Tardy"; }
            else
            {
                MessageBox.Show("Must select a action");
                return;
            }

            int index = 0;
            string name = "";
            Label sendText = sender as Label;
            foreach (Label l in names)
            {
                if (l == sendText)
                {
                    name = l.Text;
                    break;
                }
                index++;
            }
            sStates[index].Text = action;
        }

        private void Atten_Click(object sender, EventArgs e)
        {
            AttCon aCon = sender as AttCon;

            foreach (System.Windows.Forms.RadioButton rb in groupBox6.Controls)
                if (rb.Checked) { aCon.Stauts = rb.Text; }
        }

        private void atToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetAllPanels(panel2);
            panel3.Visible = true;
    
            label12.Text = DateTime.Now.ToString("MM/dd/yyyy");
            panel3.AutoScroll = true;
            panel3.Controls.Clear();

            string grabStudent = "Select FName, LName, SDate, EDate From Record";
            OleDbCommand cmd = new OleDbCommand(grabStudent, conn);
            dbReader = cmd.ExecuteReader();

            int i = 0, y = 50;
            while (dbReader.Read())
            {
                if (!(DateTime.Parse(dbReader["SDate"].ToString()) <= DateTime.Now && DateTime.Now <= DateTime.Parse(dbReader["EDate"].ToString()).AddDays(1))) { continue; }

                AttCon aCon = new AttCon(dbReader["FName"].ToString() + " " + dbReader["LName"].ToString());
                aCon.Click += new EventHandler(Atten_Click);
                aCon.Location = new Point(5, y * i);
                panel3.Controls.Add(aCon);
                i++;
            }
            dbReader.Close();
            cmd.Dispose();
        }

        private void button2_Click(object sender, EventArgs e) //Reset
        {
            foreach (AttCon aCon in panel3.Controls)
                aCon.Stauts = "Unassigned";
        }

        private void button3_Click(object sender, EventArgs e) //Finish
        {
            foreach (AttCon aCon in panel3.Controls)
            {
                if (aCon.Stauts == "Unassigned")
                {
                    MessageBox.Show("Must set a state for every student");
                    listView2.Items.Clear();
                    return;
                }
                listView2.Items.Add(new ListViewItem(new string[] { aCon.Name, aCon.Stauts }));
            }
            button2.Enabled = false;
            button3.Enabled = false;
            panel3.Enabled = false;
        }

        private void button4_Click(object sender, EventArgs e) //Back
        {
            button2.Enabled = true;
            button3.Enabled = true;
            panel3.Enabled = true;
            listView2.Items.Clear();
        }

        private bool EmailIsPreset(List<string> emails, string testAddr) //Group together
        {
            foreach (string email in emails)
            {
                if (email == testAddr) { return true; }
            }
            return false;
        }

        private void button5_Click(object sender, EventArgs e) //Send
        {
            List<StateEmail> toSend = new List<StateEmail>();
            List<string> cEmails = new List<string>();

            string getEmail = "Select Distract, Type, Email From [D&Emails] Where (Function = 'Attendance') OR (Function = 'Both')";
            OleDbCommand cmd1 = new OleDbCommand(getEmail, conn);
            dbReader = cmd1.ExecuteReader();

            while (dbReader.Read()) //Get the different emails
            {
                if (!EmailIsPreset(cEmails, dbReader["Email"].ToString()))
                {
                    toSend.Add(new StateEmail(dbReader["Email"].ToString(), dbReader["Distract"].ToString(), dbReader["Type"].ToString()));
                    cEmails.Add(dbReader["Email"].ToString());
                }
            }
            dbReader.Close();
            cmd1.Dispose();

            foreach (AttCon aCon in panel3.Controls) //Go through by student
            {
                string recordAtt = "Insert Into [Attendance] Values (@id, @name, @status, @date, @notes)";
                OleDbCommand cmd2 = new OleDbCommand(recordAtt, conn);
                cmd2.Parameters.AddWithValue("@id", GetRecordNumber("Attendance"));
                cmd2.Parameters.AddWithValue("@name", aCon.Name);
                cmd2.Parameters.AddWithValue("@status", aCon.Stauts);
                cmd2.Parameters.AddWithValue("@date", label12.Text);
                cmd2.Parameters.AddWithValue("@notes", aCon.Notes);
                cmd2.ExecuteNonQuery();
                cmd2.Dispose();

                string[] parts = (aCon.Name).Split(' ');
                string findLink = "Select Distract, Grade From [Record] Where " +
                                    "(FName = @fname) And (LName = @lname)";
                OleDbCommand cmd3 = new OleDbCommand(findLink, conn);
                cmd3.Parameters.AddWithValue("@fname", parts[0]);
                cmd3.Parameters.AddWithValue("@lname", parts[1]);
                dbReader = cmd3.ExecuteReader();
                while (dbReader.Read())
                {
                    string gLevel = "";
                    int grade = int.Parse(dbReader["Grade"].ToString());
                    if (grade < 9) { gLevel = "Middle School"; }
                    else { gLevel = "High School"; }

                    foreach (StateEmail sEmail in toSend) //Assing student to email object
                    {
                        if (sEmail.Distract == dbReader["Distract"].ToString() && sEmail.GradeLevel == gLevel)
                        {
                            sEmail.AddToBody(aCon.Name + " is " + aCon.Stauts +
                                                "\nNotes: " + aCon.Notes + "\n\n");
                        }
                    }
                }
                dbReader.Close();
            }

            try
            {
                foreach (StateEmail sEmail in toSend)
                {
                    if (sEmail.Body == "") { continue; }
                    sEmail.SendEmail(userName, passWord);
                }
                    
                cEmails.Clear();


                button2.Enabled = false;
                button3.Enabled = false;
                button4.Enabled = false;
                button5.Enabled = false;
                MessageBox.Show("Day attendance recorded and schools have been notified");
            }

            catch (SmtpException ex)
            {
                MessageBox.Show("[125] " + ex.Message + "\nMail may not have sent");
                return;
            }
        }

        private void bahaviorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetAllPanels(panel4);
            listView3.Items.Clear();
            label13.Text = DateTime.Now.ToString("MM/dd/yyyy");

            dbReader = MakeReader("Select [ID], [FName], [LName], [SDate], [EDate] From Record");
            while (dbReader.Read())
            {
                if (DateTime.Parse(dbReader["SDate"].ToString()) <= DateTime.Now && DateTime.Now <= DateTime.Parse(dbReader["EDate"].ToString()).AddDays(1))
                {
                    string name = dbReader["Fname"].ToString() + " " + dbReader["LName"].ToString();
                    string[] sparts = (dbReader["SDate"].ToString()).Split(' ');
                    string[] eparts = (dbReader["EDate"].ToString()).Split(' ');

                    listView3.Items.Add(new ListViewItem(new string[] { name, sparts[0], eparts[0] }));
                }
            }
            dbReader.Close();
        }

        private string GetCurrentDate() //Parsing rbutton to date
        {
            string date = "";
            int index = 0;
            if (radioButton4.Checked) { index = 1; }
            else if (radioButton5.Checked) { index = 2; }
            else if (radioButton6.Checked) { index = 3; }
            else if (radioButton7.Checked) { index = 4; }
            else if (radioButton8.Checked) { index = 5; }

            string currDateName = DateTime.Now.DayOfWeek.ToString();
            string[] days = new string[] { "Monday", "Tuesday", "Wednesday",
                                           "Thursday", "Friday" };
            
            int readIndex = 0;
            for (int i = 0; i < days.Length; i++)
            {
                if (days[i] == currDateName) { readIndex = i + 1; }
            }

            if (index > readIndex) { date = DateTime.Now.AddDays(index - readIndex).ToShortDateString(); }
            else if (index == readIndex) { date = DateTime.Now.ToShortDateString(); }
            else if (index < readIndex) { date = DateTime.Now.AddDays(index - readIndex).ToShortDateString(); }
            return date;
        }

        private void DatRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.RadioButton rButton = sender as System.Windows.Forms.RadioButton;

            if (rButton.Checked && listView3.SelectedItems.Count > 0)
            {
                textBox12.Text = "";

                //Get behavior
                string getReport = "Select Description From [Behavior] Where " +
                                       "(SName = @name) And (Date = @date)";
                OleDbCommand cmd = new OleDbCommand(getReport, conn);
                cmd.Parameters.AddWithValue("@name", listView3.SelectedItems[0].Text);
                cmd.Parameters.AddWithValue("@date", GetCurrentDate());
                dbReader = cmd.ExecuteReader();

                while (dbReader.Read())
                {
                    textBox12.Text = dbReader["Description"].ToString();
                }

                dbReader.Close();
                cmd.Dispose();

                if (textBox12.Text == "") { textBox12.Text = "Report not yet written"; }
            }

            else if (rButton.Checked && !(listView3.SelectedItems.Count > 0))
                MessageBox.Show("Please select student name first");
        }

        private bool CheckExistingRecord(string name, string date)
        {
            string checkRec = "Select Description From[Behavior] Where " +
                              "(SName = @name) And (Date = @date)";
            OleDbCommand cmd = new OleDbCommand(checkRec, conn);
            cmd.Parameters.AddWithValue("@name", name);
            cmd.Parameters.AddWithValue("@date", date);
            dbReader = cmd.ExecuteReader();

            while (dbReader.Read()) { return true; }
            return false;
        }

        private void button7_Click(object sender, EventArgs e) //Save what has been written
        {
            if (listView3.SelectedItems.Count > 0)
            {
                string name = listView3.SelectedItems[0].Text;
                string date = GetCurrentDate();

                if (CheckExistingRecord(name, date))
                {
                    string upBehavior = "Update [Behavior] Set Description = @dist " +
                                        "Where (SName = @name) And (Date = @date)";
                    OleDbCommand cmd = new OleDbCommand(upBehavior, conn);
                    cmd.Parameters.AddWithValue("@dist", textBox12.Text);
                    cmd.Parameters.AddWithValue("@name", name);
                    cmd.Parameters.AddWithValue("@date", date);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Bahavior has been updated");
                    textBox12.Text = "";
                }

                else
                {
                    string makeBehavior = "Insert Into [Behavior] Values (@id, @name, @date, @descrip)";
                    OleDbCommand cmd = new OleDbCommand(makeBehavior, conn);
                    cmd.Parameters.AddWithValue("@id", GetRecordNumber("Behavior"));
                    cmd.Parameters.AddWithValue("@name", name);
                    cmd.Parameters.AddWithValue("@date", date);
                    cmd.Parameters.AddWithValue("@descrip", textBox12.Text);
                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                    MessageBox.Show("Behavior has been recorded");
                    textBox12.Text = "";
                }
            }

            else { MessageBox.Show("Please select student name first"); }
        }

        private void SendEmail(string inName, string inStr, string rEmail)
        {
            MailMessage mMess = new MailMessage();
            SmtpClient smtp = new SmtpClient();
            try
            {
                mMess.From = new MailAddress("iscreports@wflboces.org", "Suspension Center");
                mMess.To.Add(new MailAddress(rEmail));
                mMess.Subject = "Week Behavior For " + inName;
                mMess.Body = inStr;
                mMess.BodyEncoding = Encoding.UTF8;
                mMess.SubjectEncoding = Encoding.UTF8;

                smtp.Host = "smtp-relay.gmail.com";
                smtp.Port = 587;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new NetworkCredential("iscreports@wflboces.org", "G1nger#7");
                //smtp.EnableSsl = true; (Not sure why this makes the connection drop)
                smtp.Send(mMess);
            }

            catch (SmtpException) { MessageBox.Show("[587] Email Failure found");  }

            finally
            {
                mMess.Dispose();
                smtp.Dispose();
            }
        }

        int readIndex = 0;
        private void button8_Click(object sender, EventArgs e) //Send behavior
        {

            DialogResult result = MessageBox.Show("This will send the weekly report for everyone. " +
                                                    "Are you sure?", "Confirmation",
                                                    MessageBoxButtons.YesNoCancel);

            if (result != DialogResult.Yes) { return; }
            List<string> sNames = new List<string>();

            string gNames = "Select FName, LName, SDate, EDate From [Record]";
            OleDbCommand cmd1 = new OleDbCommand(gNames, conn);
            dbReader = cmd1.ExecuteReader();
            while (dbReader.Read())
            {
                if (DateTime.Parse(dbReader["SDate"].ToString()) <= DateTime.Now && DateTime.Now <= DateTime.Parse(dbReader["EDate"].ToString()).AddDays(1))
                    sNames.Add(dbReader["FName"].ToString() + " " +
                            dbReader["LName"].ToString());
            }

            string currDateName = DateTime.Now.DayOfWeek.ToString();
            string[] days = new string[] { "Monday", "Tuesday", "Wednesday",
                                            "Thursday", "Friday" };

            for (int i = 0; i < days.Length; i++)
            {
                if (days[i] == currDateName)
                {
                    readIndex = i + 1;
                    break;
                }
            }

            List<DateTime> dates = new List<DateTime>();
            DateTime bDate = DateTime.Now.AddDays(1 - readIndex);
            dates.Add(bDate);

            for (int i = 1; i < 5; i++)
                dates.Add(bDate.AddDays(i));

            foreach (string name in sNames)
            {
                string report = name + "\n\n";
                for (int i = 0; i < dates.Count; i++)
                {
                    report += dates[i].ToShortDateString() + "\n";
                    string sRec = "";
                    string sBehavior = "Select Description From [Behavior] Where " +
                                        "(SName = @sName) And (Date = @date)";
                    OleDbCommand cmd2 = new OleDbCommand(sBehavior, conn);
                    cmd2.Parameters.AddWithValue("@sName", name);
                    cmd2.Parameters.AddWithValue("@date", dates[i].ToShortDateString());
                    dbReader = cmd2.ExecuteReader();
                    while (dbReader.Read())
                        sRec = dbReader["Description"].ToString();

                    if (sRec != "") { report += ("  " + sRec + "\n\n"); }
                    else { report += "  None\n\n"; }
                    dbReader.Close();
                    cmd2.Dispose();
                }

                int grade = 0;
                string dist = "";

                string[] parts = name.Split(' ');
                string getEmail = "Select Distract, Grade From [Record] Where " +
                                "(FName = @fName) And (LName = @lName)";
                OleDbCommand cmd3 = new OleDbCommand(getEmail, conn);
                cmd3.Parameters.AddWithValue("@fName", parts[0]);
                cmd3.Parameters.AddWithValue("@lName", parts[1]);
                dbReader = cmd3.ExecuteReader();

                while (dbReader.Read())
                {
                    grade = int.Parse(dbReader["Grade"].ToString());
                    dist = dbReader["Distract"].ToString();
                }
                dbReader.Close();

                string school = "";
                if (grade < 9) { school = "Middle School"; }
                else { school = "High School"; }

                List<string> emails = new List<string>();
                string dEmail = "Select Email, Function From [D&Emails] Where (Distract = @dist) And (Type = @type)";
                OleDbCommand cmd4 = new OleDbCommand(dEmail, conn);
                cmd4.Parameters.AddWithValue("@dist", dist);
                cmd4.Parameters.AddWithValue("@type", school);
                dbReader = cmd4.ExecuteReader();
                while (dbReader.Read())
                {
                    if (dbReader["Function"].ToString() == "Behavior" || dbReader["Function"].ToString() == "Both")
                        emails.Add(dbReader["Email"].ToString());
                }
                dbReader.Close();

                foreach (string email in emails)
                    SendEmail(name, report, email);
                emails.Clear();
            }
            MessageBox.Show("Behavior has been sent");
        }

        private void button11_Click(object sender, EventArgs e) //View Student Report
        {
            string name = "";
            if (listView3.SelectedItems.Count > 0)
                name = listView3.SelectedItems[0].Text;
            else { MessageBox.Show("Must select a student first"); return; }

            string currDateName = DateTime.Now.DayOfWeek.ToString();
            string[] days = new string[] { "Monday", "Tuesday", "Wednesday",
                                           "Thursday", "Friday" };

            for (int i = 0; i < days.Length; i++)
            {
                if (days[i] == currDateName)
                {
                    readIndex = i + 1;
                    break;
                }
            }

            List<DateTime> dates = new List<DateTime>();
            DateTime bDate = DateTime.Now.AddDays(1 - readIndex);
            dates.Add(bDate);

            for (int i = 1; i < 5; i++)
                dates.Add(bDate.AddDays(i));

            string report = name + "\n\n";
            for (int i = 0; i < dates.Count; i++)
            {
                report += dates[i].ToShortDateString() + "\n";
                string sRec = "";
                string sBehavior = "Select Description From [Behavior] Where " +
                                   "(SName = @sName) And (Date = @date)";
                OleDbCommand cmd2 = new OleDbCommand(sBehavior, conn);
                cmd2.Parameters.AddWithValue("@sName", name);
                cmd2.Parameters.AddWithValue("@date", dates[i].ToShortDateString());
                dbReader = cmd2.ExecuteReader();
                while (dbReader.Read())
                    sRec = dbReader["Description"].ToString();

                if (sRec != "") { report += ("  " + sRec + "\n\n"); }
                else { report += "  Report Not Written Yet\n\n"; }
                dbReader.Close();
                cmd2.Dispose();
            }
            MessageBox.Show(report);
        }

        private void finishedStudentToolStripMenuItem_Click(object sender, EventArgs e) //Work report
        {
            button12.Enabled = false;
            button13.Enabled = false;

            SetAllPanels(panel5);
            listBox2.Items.Clear();

            string getStudent = "Select FName, LName, SDate, EDate From [Record]";
            OleDbCommand cmd = new OleDbCommand(getStudent, conn);
            dbReader = cmd.ExecuteReader();

            while (dbReader.Read())
            {
                if (DateTime.Parse(dbReader["SDate"].ToString()) <= DateTime.Now && DateTime.Now <= DateTime.Parse(dbReader["EDate"].ToString()).AddDays(1))
                    listBox2.Items.Add(dbReader["FName"].ToString() + " " + dbReader["LName"].ToString());
            }
            dbReader.Close();
            cmd.Dispose();
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            if (listBox2.SelectedIndex == -1) { return; }

            string name = listBox2.SelectedItem.ToString();

            string getWork = "Select WName, Subject From [WorkLog] " +
                             "Where SName = @name";
            OleDbCommand cmd = new OleDbCommand(getWork, conn);
            cmd.Parameters.AddWithValue("@name", listBox2.SelectedItem.ToString());
            dbReader = cmd.ExecuteReader();

            while (dbReader.Read())
            {
                listBox3.Items.Add(dbReader["WName"].ToString());
            }
            dbReader.Close();

            //if (listBox3.Items.Count == 0) { listBox3.Items.Add("None"); }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            panel6.Controls.Clear();
            int x = 0;
            panel6.AutoScroll = true;
            foreach (Object sItem in listBox3.SelectedItems)
            {
                Label assign = new Label();
                assign.Size = new Size(150, 28);
                assign.Text = sItem.ToString();
                assign.Font = new Font("Microsoft Sans Serif", 12F,
                    System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                assign.Location = new Point(20 + (x * 180), 22);
                assign.BorderStyle = BorderStyle.Fixed3D;
                assign.BackColor = Color.White;
                assign.Click += new EventHandler(this.Assign_Click);
                panel6.Controls.Add(assign);
                x += 1;
            }
        }

        private void Assign_Click(object sender, EventArgs e)
        {
            button15.Enabled = true;
            Label tb = sender as Label;
            string assign = tb.Text;

            string findWork = "Select * From [WorkLog] Where WName = @name";
            OleDbCommand cmd = new OleDbCommand(findWork, conn);
            cmd.Parameters.AddWithValue("@name", tb.Text);
            dbReader = cmd.ExecuteReader();

            bool finished = false;
            while (dbReader.Read())
            {
                label17.Text = "Name: " + tb.Text;
                label15.Text = "Subject: " + dbReader["Subject"].ToString();

                if (dbReader["Status"].ToString() == "Complete")
                {
                    finished = true;
                    button15.Enabled = false;
                    button12.Enabled = true;
                    //button13.Enabled = true;
                }

                label18.Text = "Assignment Path:\n" + dbReader["PathToFile"].ToString();

                if (finished)
                {
                    label19.Text = "Finished Path:\n" + dbReader["PathToFin"].ToString();
                    radioButton10.Checked = true;
                }

                else
                {
                    label19.Text = "Finished Path:    (Not Finished)";
                    radioButton9.Checked = true;
                }
                
            }
            dbReader.Close();
            cmd.Dispose();
        }

        private void button15_Click(object sender, EventArgs e) //Copy finished
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "(*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            string sourcePath = @"";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                sourcePath = openFileDialog1.FileName;
                string[] parts = sourcePath.Split('\\');
                int index = parts.Length;
                string targetPath = Application.StartupPath + @"\" +
                        listBox2.SelectedItem.ToString() + @"\f_" + parts[index - 1];

                MessageBox.Show(targetPath);

                try { File.Copy(sourcePath, targetPath); }
                catch (IOException) { MessageBox.Show("File already exists"); return; }

                string uFile = "Update [WorkLog] Set Status = @stat, PathToFin = @file " +
                               "Where (SName = @name) And (WName = @wname)";
                OleDbCommand cmd = new OleDbCommand(uFile, conn);
                cmd.Parameters.AddWithValue("@stat", "Complete");
                cmd.Parameters.AddWithValue("@file", targetPath);
                cmd.Parameters.AddWithValue("@name", listBox2.SelectedItem.ToString());
                string[] parts2 = (label17.Text).Split(' ');
                cmd.Parameters.AddWithValue("@wname", parts2[1]);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                label19.Text = "Finished Location\n" + targetPath;
                button15.Enabled = false;
                button12.Enabled = true;
            }
        }

        private void button12_Click(object sender, EventArgs e) 
        {
            button13.Enabled = true;
            string[] parts = (label17.Text).Split(' ');
            listBox4.Items.Add(parts[1]);
        }

        private void button13_Click(object sender, EventArgs e)
        {
            List<string> paths = new List<string>();
            string getFile = "Select PathToFin From [WorkLog] Where " +
                                "(SName = @name) And (WName = @wname)";

            foreach (string assign in listBox4.Items)
            {
                OleDbCommand cmd = new OleDbCommand(getFile, conn);
                cmd.Parameters.AddWithValue("@name", listBox2.SelectedItem.ToString());
                cmd.Parameters.AddWithValue("@wname", assign);
                dbReader = cmd.ExecuteReader();
                while (dbReader.Read()) { paths.Add(dbReader["PathToFin"].ToString()); }
                dbReader.Close();
                cmd.Dispose();
            }

            string zPath = Application.StartupPath + @"\tempzip.zip";
            try { File.Delete(zPath); }
            catch (DirectoryNotFoundException)
            {
                MessageBox.Show("Error removing work from student file");
                return;
            }
                
            using (ZipArchive archive = ZipFile.Open(zPath, ZipArchiveMode.Create))
            {
                foreach (string path in paths)
                {
                    archive.CreateEntryFromFile(path, Path.GetFileName(path));
                }
                archive.Dispose();
            }

            string cEmail = "";
            string[] parts = (listBox2.SelectedItem.ToString()).Split(' ');
            string fEmail = "Select CEmail From [Record] Where " +
                                "(FName = @fname) And (LName = @lname)";
            OleDbCommand cmd2 = new OleDbCommand(fEmail, conn);
            cmd2.Parameters.AddWithValue("@fname", parts[0]);
            cmd2.Parameters.AddWithValue("@lname", parts[1]);
            dbReader = cmd2.ExecuteReader();
            while (dbReader.Read()) { cEmail = dbReader["CEmail"].ToString(); }
            dbReader.Close();
            cmd2.Dispose();

            MailMessage mMess = new MailMessage();
            SmtpClient smtp = new SmtpClient();

            try
            {
                mMess.From = new MailAddress("iscreports@wflboces.org", "Suspension Center");
                mMess.To.Add(new MailAddress(cEmail));
                mMess.Subject = "Finished Work For " + listBox2.SelectedItem.ToString();
                mMess.Body = "A zip file has been added containing the finished work";
                mMess.BodyEncoding = Encoding.UTF8;
                mMess.SubjectEncoding = Encoding.UTF8;
                mMess.Attachments.Add(new Attachment(zPath));

                smtp.Host = "smtp-relay.gmail.com";
                smtp.Port = 587;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new NetworkCredential("iscreports@wflboces.org", "G1nger#7");
                //smtp.EnableSsl = true; (Not sure why this makes the connection drop)
                smtp.Send(mMess);
                MessageBox.Show("Work as been sent for " + listBox2.SelectedItem.ToString());
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }

            //"[104] Failed to send work"
            finally
            {
                mMess.Dispose();
                smtp.Dispose();           
            }
        }

        private void lunchToolStripMenuItem_Click(object sender, EventArgs e)
        {
            listBox5.Items.Clear();
            SetAllPanels(panel7);
            label20.Text = DateTime.Today.ToString("MM/dd/yyyy");

            string gName = "Select FName, LName, SDate, EDate From [Record]";
            OleDbCommand cmd = new OleDbCommand(gName, conn);
            dbReader = cmd.ExecuteReader();
            while (dbReader.Read())
                if (DateTime.Parse(dbReader["SDate"].ToString()) <= DateTime.Now && DateTime.Now <= DateTime.Parse(dbReader["EDate"].ToString()).AddDays(1))
                    listBox5.Items.Add(dbReader["FName"].ToString() + " " +
                                   dbReader["LName"].ToString());
            dbReader.Close();
            cmd.Dispose();

            panel8.AutoScroll = false;
            panel8.VerticalScroll.Enabled = false;
            panel8.VerticalScroll.Visible = false;
            panel8.VerticalScroll.Maximum = 0;
            panel8.AutoScroll = true;
        }

        private void listBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox5.SelectedIndex == -1) { return; }

            List<Label> orders = new List<Label>();
            foreach (Label label in panel8.Controls)
            {
                if (label.Text == listBox5.SelectedItem.ToString())
                {
                    MessageBox.Show("Student meal already rocorded");
                    return;
                }
            }

            label23.Text = "Student Name: " + listBox5.SelectedItem.ToString();
            string[] parts = (listBox5.SelectedItem.ToString()).Split(' ');

            string gDist = "Select Distract From [Record] Where " +
                           "(FName = @fname) And (LName = @lname)";
            OleDbCommand cmd = new OleDbCommand(gDist, conn);
            cmd.Parameters.AddWithValue("@fname", parts[0]);
            cmd.Parameters.AddWithValue("@lname", parts[1]);
            dbReader = cmd.ExecuteReader();
            while (dbReader.Read())
                label24.Text = "Distract: " + dbReader["Distract"].ToString();

            dbReader.Close();
            cmd.Dispose();
        }

        private void MealRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            groupBox17.Enabled = false;
            groupBox18.Enabled = false;

            if (radioButton15.Checked) { groupBox17.Enabled = true; }
            else if (radioButton17.Checked) { groupBox18.Enabled = true; }
        }

        private void LunchBox_Click(object sender, EventArgs e)
        {

        }

        private List<System.Windows.Forms.RadioButton> GetRadioButtons(GroupBox gb)
        {
            List<System.Windows.Forms.RadioButton> rButtons =
                new List<System.Windows.Forms.RadioButton>();
            foreach (System.Windows.Forms.RadioButton rb in gb.Controls)
                rButtons.Add(rb);

            return rButtons;
        }

        int counter = 0;
        private void button16_Click(object sender, EventArgs e)
        {
            if (listBox5.SelectedIndex == -1) { return; }

            foreach (Lunch lunch in sLunches)
            {
                if (lunch.Name == listBox5.SelectedItem.ToString())
                {
                    MessageBox.Show("Student order already recorded");
                    return;
                }
            }
            List<System.Windows.Forms.RadioButton> rButtons;

            string time = "";
            rButtons = GetRadioButtons(groupBox19);
            foreach (System.Windows.Forms.RadioButton rb in rButtons)
                if (rb.Checked) { time = rb.Text; }
            if (time == "") { MessageBox.Show("Select meal time"); return; }

            string status = "";
            rButtons = GetRadioButtons(groupBox13);
            foreach (System.Windows.Forms.RadioButton rb in rButtons)
                if (rb.Checked) { status = rb.Text; }
            if (status == "") { MessageBox.Show("Select payment code"); return; }

            string[] parts = (label24.Text).Split(' ');

            string meal = "";
            rButtons = GetRadioButtons(groupBox16);
            foreach (System.Windows.Forms.RadioButton rb in rButtons)
                if (rb.Checked) { meal = rb.Text; }
            if (meal == "") { MessageBox.Show("Select Meal"); return; }

            string add = "";
            if (radioButton15.Checked)
            {
                rButtons = GetRadioButtons(groupBox17);
                foreach (System.Windows.Forms.RadioButton rb in rButtons)
                    if (rb.Checked) { add = rb.Text; }
                if (add == "") { MessageBox.Show("Select meat choice"); return; }
            }

            else if (radioButton17.Checked)
            {
                rButtons = GetRadioButtons(groupBox18);
                foreach (System.Windows.Forms.RadioButton rb in rButtons)
                    if (rb.Checked) { add = rb.Text; }
                if (add == "") { MessageBox.Show("Select dressing choice"); return; }
            }

            else { add = "None"; }

            string milk = "";
            rButtons = GetRadioButtons(groupBox21);
            foreach (System.Windows.Forms.RadioButton rb in rButtons)
                if (rb.Checked) { milk = rb.Text; }
            if (milk == "") { MessageBox.Show("Select milk choice"); return; }

            Lunch l = new Lunch(listBox5.SelectedItem.ToString(), time, status,
                            parts[1], meal, add, milk);
            sLunches.Add(l);

            Label myLunch = new Label();
            myLunch.Size = new Size(200, 28);
            myLunch.Text = listBox5.SelectedItem.ToString();
            myLunch.Font = new Font("Microsoft Sans Serif", 12F,
                System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            myLunch.Location = new Point(20 + (counter * 220), 30);
            myLunch.BorderStyle = BorderStyle.Fixed3D;
            myLunch.BackColor = Color.White;
            myLunch.Click += new EventHandler(this.LunchBox_Click);
            panel8.Controls.Add(myLunch);
            counter++;
        }

        private void button14_Click(object sender, EventArgs e)
        {
            string pStr = "Suspension Center Food Order\n\n";
            foreach (Lunch order in sLunches)
                pStr += (order + "\n\n");

            try
            {
                PrintDocument p = new PrintDocument();
                var font = new Font("Times New Roman", 14);
                var margins = new Margins(50, 50, 50, 50);
                var layoutArea = new RectangleF(
                    margins.Left,
                    margins.Top,
                    p.DefaultPageSettings.PrintableArea.Width - (margins.Left + margins.Right),
                    p.DefaultPageSettings.PrintableArea.Height - (margins.Top + margins.Bottom));
                var layoutSize = layoutArea.Size;
                layoutSize.Height = layoutSize.Height - font.GetHeight(); // keep lastline visible
                var brush = new SolidBrush(Color.Black);

                var remainingText = pStr;

                p.PrintPage += delegate (object sender1, PrintPageEventArgs e1)
                {
                    int charsFitted, linesFilled;
                    var realsize = e1.Graphics.MeasureString(
                        remainingText,
                        font,
                        layoutSize,
                        StringFormat.GenericDefault,
                        out charsFitted,
                        out linesFilled);

                    var fitsOnPage = remainingText.Substring(0, charsFitted);
                    remainingText = remainingText.Substring(charsFitted).Trim();

                    e1.Graphics.DrawString(
                        fitsOnPage,
                        font,
                        brush,

                        layoutArea);

                    e1.HasMorePages = remainingText.Length > 0;
                };
                p.Print();
            }
            catch (Exception) { MessageBox.Show("Printing error"); }
        }

        private void changeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GSheet gSheet = new GSheet();
            gSheet.ShowDialog(this);
        }

        private string GetDataPath(string sName, string wName)
        {
            string rStr = "";
            string gPath = "Select PathToFile From [WorkLog] Where " +
                               "(SName = @sName) And (WName = wName)";
            OleDbCommand cmd = new OleDbCommand(gPath, conn);
            cmd.Parameters.AddWithValue("@sName", sName);
            cmd.Parameters.AddWithValue("@wName", wName);
            dbReader = cmd.ExecuteReader();
            while (dbReader.Read()) { rStr = dbReader["PathToFile"].ToString(); }
            dbReader.Close();
            cmd.Dispose();
            return rStr;
        }

        string item = "";
        private void Item_Click(object sender, ToolStripItemClickedEventArgs e) //Right Click Menue
        {
            if (!(listView1.SelectedItems.Count > 0)) { return; }

            if (e.ClickedItem.Text == "Print")
            {
                string path = GetDataPath(listView1.SelectedItems[0].Text, item);

                if (path != "")
                {
                    var print = new ProcessStartInfo(path);
                    print.UseShellExecute = true;
                    print.Verb = "print";

                    try { var process = Process.Start(print); }
                    catch (System.ComponentModel.Win32Exception)
                    {
                        MessageBox.Show("Print error occured, this could\n" +
                                        "be because there is not a printer\n" +
                                        "set up to work with this computer");
                        return;
                    }
                    MessageBox.Show("File sent to printer");
                }

                else { MessageBox.Show("Error in locating file to print"); }
            }

            else if (e.ClickedItem.Text == "Remove")
            {
                string path = GetDataPath(listView1.SelectedItems[0].Text, item);

                string rWork = "Delete * From [WorkLog] Where " +
                               "(SName = @sName) And (WName = @wName)";

                OleDbCommand cmd1 = new OleDbCommand(rWork, conn);
                cmd1.Parameters.AddWithValue("@sName", listView1.SelectedItems[0].Text);
                cmd1.Parameters.AddWithValue("@wName", item);
                cmd1.ExecuteNonQuery();
                cmd1.Dispose();

                listBox1.Items.Clear();
                string rSub = "Select WName From [WorkLog] Where Subject = @sub";
                OleDbCommand cmd2 = new OleDbCommand(rSub, conn);
                cmd2.Parameters.AddWithValue("@sub", comboBox1.Text);
                dbReader = cmd2.ExecuteReader();
                while (dbReader.Read())
                    listBox1.Items.Add(dbReader["WName"].ToString());
                dbReader.Close();
                cmd2.Dispose();

                try { File.Delete(path); }
                catch (DirectoryNotFoundException)
                {
                    MessageBox.Show("Error removing work from student file");
                }
            }
        }

        private void listBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && listBox1.SelectedIndex != -1)
            {
                listBox1.SelectedItem = listBox1.IndexFromPoint(Cursor.Position.X, Cursor.Position.Y);
                item = listBox1.SelectedItem.ToString();
                ContextMenuStrip menu = new ContextMenuStrip();
                menu.Items.Add("Print");
                menu.Items.Add("Remove");
                menu.ItemClicked += new ToolStripItemClickedEventHandler(Item_Click);
                menu.Show(ListBox.MousePosition);
            }
        }

        private void emailsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            EmailPanel ePanel = new EmailPanel(conn);
            ePanel.ShowDialog(this);
        }

        string student = "";
        private void Student_Click(object sender, ToolStripItemClickedEventArgs e) //Right Click Menu
        {
            //MessageBox.Show(student);
            string[] parts = student.Split(' ');
            string rStudent = "Delete From [Record] Where (FName = @fName) And (LName = @lName)";
            OleDbCommand cmd = new OleDbCommand(rStudent, conn);
            cmd.Parameters.AddWithValue("@fName", RemoveWhiteSpace(parts[0]));
            cmd.Parameters.AddWithValue("@lName", RemoveWhiteSpace(parts[1]));
            cmd.ExecuteNonQuery();
            cmd.Dispose();

            listView1.Items.Clear();
            dbReader = MakeReader("Select [ID], [FName], [LName], [SDate], [EDate] From Record");
            while (dbReader.Read())
            {
                if (DateTime.Parse(dbReader["SDate"].ToString()) <= DateTime.Now && DateTime.Now <= DateTime.Parse(dbReader["EDate"].ToString()).AddDays(1))
                {
                    string name = dbReader["FName"].ToString() + " " + dbReader["LName"].ToString();
                    if (name == student) { continue; }
                    string[] sparts = (dbReader["SDate"].ToString()).Split(' ');
                    string[] eparts = (dbReader["EDate"].ToString()).Split(' ');

                    listView1.Items.Add(new ListViewItem(new string[] { name, sparts[0], eparts[0] }));
                }
            }
            dbReader.Close();
        }

        private void listView1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && listView1.SelectedItems.Count > 0)
            {
                student = listView1.SelectedItems[0].Text;
                ContextMenuStrip menu = new ContextMenuStrip();
                menu.Items.Add("Remove");
                menu.ItemClicked += new ToolStripItemClickedEventHandler(Student_Click);
                menu.Show(ListBox.MousePosition);
            }
        }

        public string UserName { set { userName = value; } }
        public string PassWord { set { passWord = value; } }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e) { this.Close(); }

        private void attendanceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            HistRec rec = new HistRec(conn);
            rec.ShowDialog(this);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e) { conn.Close(); }
    }
}
