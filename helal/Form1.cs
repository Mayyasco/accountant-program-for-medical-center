using System;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SQLite;
using System.Text.RegularExpressions;
using System.IO;
using Microsoft.Office.Interop.Excel;

using System.Diagnostics;

namespace helal
{
    public partial class Form1 : Form
    {
        string wasel = "",w="";
        Worksheet workSheet1;
        Workbook workBook1;
        int index = 0,index1=0;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            get_clinic();
            ref_clinic();
            comboBox3.Text = "ﬂ‘›Ì…"; comboBox5.Text = "1"; comboBox6.Text = "1";
            tabControl1.TabPages.Remove(tabPage2);
            string ss = DateTime.Now.Day.ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
            textBox12.Text = ss; textBox11.Text = ss; textBox5.Text = ss; textBox16.Text = ss; textBox24.Text = ss; textBox23.Text = ss;
        }

        private void ref_clinic()
        {
            int i = 0;
            try
            {
                comboBox1.Items.Clear(); comboBox2.Items.Clear();
                SQLiteConnection m_dbConnection;
                m_dbConnection = new SQLiteConnection("Data Source=helal.db;Version=3;");
                m_dbConnection.Open();

                string sql = "select * from clinic";
                SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                SQLiteDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    comboBox1.Items.Add((string)reader["name"] + "-" + (string)reader["doctor"]); comboBox2.Items.Add((string)reader["name"] + "-" + (string)reader["doctor"]);
                    if (i == 0) { comboBox1.Text = (string)reader["name"]+ "-" + (string)reader["doctor"]; comboBox2.Text = (string)reader["name"]+ "-" + (string)reader["doctor"]; }
                    i++;
                }
                reader.Close();
                m_dbConnection.Close();
                comboBox2.Items.Add("«·ﬂ·");
            }

            catch (Exception ex)
            {
                MessageBox.Show("ÌÊÃœ ·œÌﬂ „‘ﬂ·…");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int i=0;
            try
            {
                if (dataGridView1.RowCount < 2) { MessageBox.Show("·« ÌÃÊ“ √‰ ÌﬂÊ‰ ÃœÊ· «·⁄Ì«œ«  ›«—€"); return; }

                for (i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    string s1 = (string)dataGridView1.Rows[i].Cells[0].Value;
                    string s2 = (string)dataGridView1.Rows[i].Cells[1].Value;
                    int r = i + 1;
                    if (string.IsNullOrEmpty(s1) || string.IsNullOrEmpty(s2)) { MessageBox.Show(" ÌÊÃœ „⁄·Ê„… ‰«ﬁ’… ›Ì «·”ÿ— : " + r.ToString()); return; }
                    if (s1.Trim().Length < 3 || s2.Trim().Length < 3) { MessageBox.Show("⁄œœ «·√Õ—› ·√Ì «”„ ÌÃ» √‰ ÌﬂÊ‰ √ﬂÀ— „‰ Õ—›Ì‰ ›Ì «·”ÿ— : " + r.ToString()); return; }

                }
                SQLiteConnection m_dbConnection;
                m_dbConnection = new SQLiteConnection("Data Source=helal.db;Version=3;");
                m_dbConnection.Open();

                string sql = "";
                sql = "DELETE FROM clinic";
                SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                command.ExecuteNonQuery();
                

                for ( i = 0; i < dataGridView1.Rows.Count-1; i++)
                {
                    sql = "insert into clinic values ('" + dataGridView1.Rows[i].Cells[0].Value.ToString() + "','" + dataGridView1.Rows[i].Cells[1].Value.ToString() + "')";
                    command = new SQLiteCommand(sql, m_dbConnection);
                        command.ExecuteNonQuery();
                }
                m_dbConnection.Close();
                ref_clinic();
                MessageBox.Show(" „ «·Õ›Ÿ »‰Ã«Õ");
            }

            catch (Exception ex)
            {
                MessageBox.Show("ÌÊÃœ ·œÌﬂ „‘ﬂ·…");
            } 
        
        }

        private void button8_Click(object sender, EventArgs e)
        {

            get_clinic(); 
        }

        private void get_clinic()
        {
            int i = 0;
            try
            {
                dataGridView1.Rows.Clear();
                SQLiteConnection m_dbConnection;
                m_dbConnection = new SQLiteConnection("Data Source=helal.db;Version=3;");
                m_dbConnection.Open();

                string sql = "select * from clinic";
                SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                SQLiteDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[i].Cells[0].Value = reader["name"];
                    dataGridView1.Rows[i].Cells[1].Value = reader["doctor"];
                    i++;
                }
                reader.Close();
                m_dbConnection.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show("ÌÊÃœ ·œÌﬂ „‘ﬂ·…");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int i = 0;
            try
            {
                if (!Regex.IsMatch(textBox1.Text, @"^[0-9]{9}$")) { MessageBox.Show(" √ﬂœ „‰ «œŒ«· —ﬁ„ «·ÂÊÌ… »«·‘ﬂ· «·’ÕÌÕ"); return; }
                if (textBox2.Text.Trim().Length < 3) { MessageBox.Show("ÌÃ» √‰ ÌﬂÊ‰ «·«”„ √ﬂÀ— „‰ Õ—›Ì‰"); return; }
                SQLiteConnection m_dbConnection;
                m_dbConnection = new SQLiteConnection("Data Source=helal.db;Version=3;");
                m_dbConnection.Open();

                string sql = "select * from patient where id='" + textBox1.Text + "'";
                SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                SQLiteDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    i++;
                }
                reader.Close();

                if (i != 0) { MessageBox.Show("—ﬁ„ «·ÂÊÌ… „ÊÃÊœ „”»ﬁ«");  m_dbConnection.Close(); return;}

                sql = "insert into patient values ('" + textBox1.Text + "','" + textBox2.Text + "','" +textBox3.Text + "','"+textBox4.Text+ "')";
                command = new SQLiteCommand(sql, m_dbConnection);
                command.ExecuteNonQuery();

                m_dbConnection.Close();
                textBox1.Text = ""; textBox2.Text = ""; textBox3.Text = ""; textBox4.Text = "";
                MessageBox.Show(" „ «·Õ›Ÿ »‰Ã«Õ");
            }

            catch (Exception ex)
            {
                MessageBox.Show("ÌÊÃœ ·œÌﬂ „‘ﬂ·…");
            } 
        }

        private void button7_Click(object sender, EventArgs e)
        {
           monthCalendar1.Visible = true;
           monthCalendar1.Focus();
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            string d, y, m;
            y = monthCalendar1.SelectionRange.Start.Year.ToString();
            d = monthCalendar1.SelectionRange.Start.Day.ToString();
            m = monthCalendar1.SelectionRange.Start.Month.ToString();
            textBox12.Text= d + "/" + m + "/" + y;
            monthCalendar1.Visible = false;
        }

        private void monthCalendar1_MouseLeave(object sender, EventArgs e)
        {
            monthCalendar1.Visible = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            monthCalendar2.Visible = true;
            monthCalendar2.Focus();
        }

        private void monthCalendar2_MouseLeave(object sender, EventArgs e)
        {
            monthCalendar2.Visible = false;
        }

        private void monthCalendar2_DateSelected(object sender, DateRangeEventArgs e)
        {
            string d, y, m;
            y = monthCalendar2.SelectionRange.Start.Year.ToString();
            d = monthCalendar2.SelectionRange.Start.Day.ToString();
            m = monthCalendar2.SelectionRange.Start.Month.ToString();
            textBox11.Text = d + "/" + m + "/" + y;
            monthCalendar2.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            monthCalendar3.Visible = true;
            monthCalendar3.Focus();
        }

        private void monthCalendar3_DateSelected(object sender, DateRangeEventArgs e)
        {
            string d, y, m;
            y = monthCalendar3.SelectionRange.Start.Year.ToString();
            d = monthCalendar3.SelectionRange.Start.Day.ToString();
            m = monthCalendar3.SelectionRange.Start.Month.ToString();
            textBox5.Text = d + "/" + m + "/" + y;
            monthCalendar3.Visible = false;
        }

        private void monthCalendar3_MouseLeave(object sender, EventArgs e)
        {
            monthCalendar3.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int i = 0;
            try
            {
                SQLiteConnection m_dbConnection;
                m_dbConnection = new SQLiteConnection("Data Source=helal.db;Version=3;");
                m_dbConnection.Open();

                string sql = "select * from patient where id='"+textBox8.Text+"'";
                SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                SQLiteDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    
                    textBox10.Text = (string)reader["name"];
                    textBox9.Text = (string)reader["address"];
                    textBox7.Text = (string)reader["mobile"]; textBox8.Enabled = false;
                    i++;
                }
                if (i == 0)
                { MessageBox.Show("—ﬁ„ «·ÂÊÌ… €Ì— „ÊÃÊœ");
                    textBox10.Text = ""; textBox9.Text = "";
                    textBox7.Text = ""; textBox8.Enabled = true;
                }
                reader.Close();
                m_dbConnection.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show("ÌÊÃœ ·œÌﬂ „‘ﬂ·…");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            float j1; string num = "";
            //if (textBox8.Enabled) { MessageBox.Show("√œŒ· —ﬁ„ «·ÂÊÌ… À„ «÷€ÿ “— „Ê«›ﬁ »Ã«‰»Â«"); return; }
            if (!Regex.IsMatch(textBox8.Text, @"^[0-9]{9}$")) { MessageBox.Show(" √ﬂœ „‰ «œŒ«· —ﬁ„ «·ÂÊÌ… »«·‘ﬂ· «·’ÕÌÕ"); return; }
            if (textBox10.Text.Trim().Length < 3) { MessageBox.Show("ÌÃ» √‰ ÌﬂÊ‰ «·«”„ √ﬂÀ— „‰ Õ—›Ì‰"); return; }
            if(!float.TryParse(textBox6.Text,out j1)){MessageBox.Show("√œŒ· «·—”Ê„ »«·‘ﬂ· «·’ÕÌÕ"); return;}

            int i = 0,j=0;
            try
            {
                SQLiteConnection m_dbConnection;
                m_dbConnection = new SQLiteConnection("Data Source=helal.db;Version=3;");
                m_dbConnection.Open();
                string sql = "";
                SQLiteCommand command; SQLiteDataReader reader;
                if (!checkBox1.Checked)
                {
                    sql = "select * from visit where id='" + textBox8.Text + "' and clinic='" + comboBox1.Text + "' and date='" + textBox5.Text + "'";
                    command = new SQLiteCommand(sql, m_dbConnection);
                    reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        i++;
                    }
                    reader.Close();

                    if (i != 0)
                    {
                        sql = "delete from visit where id='" + textBox8.Text + "' and clinic='" + comboBox1.Text + "' and date='" + textBox5.Text + "'";
                        command = new SQLiteCommand(sql, m_dbConnection);
                        command.ExecuteNonQuery();

                    }
                    else
                    {
                        sql = "select * from visit where id='" + textBox8.Text + "' and clinic='" + comboBox1.Text + "'";
                        command = new SQLiteCommand(sql, m_dbConnection);
                        reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            j++;
                        }
                        reader.Close();

                        if (check_date2() && j != 0)
                        {
                            DialogResult result3 = MessageBox.Show("Â–« «·„—Ì÷ ﬁœ “«— ‰›” «·⁄Ì«œ… Œ·«· «”»Ê⁄Ì‰ Â·  —Ìœ  €ÌÌ— «·—”Ê„ø ",
                              " Õ–Ì—",
                              MessageBoxButtons.YesNo,
                              MessageBoxIcon.Question,
                              MessageBoxDefaultButton.Button1);
                            if (result3 == DialogResult.Yes)
                            {
                                return;
                            }
                        }
                    }
                }
                //------------------
                sql = "select * from num";
                 command = new SQLiteCommand(sql, m_dbConnection);
                  reader = command.ExecuteReader();
                while (reader.Read())
                {
                    num = (string)reader["num"]; 
                }
                reader.Close();
                int x = int.Parse(num); x++;
                //------------------
                sql = "insert into visit values ('" + textBox8.Text + "','" + comboBox1.Text + "','" + textBox5.Text + "','" + textBox6.Text + "','" + textBox10.Text + "','" + comboBox3.Text + "','" + x.ToString() + "')";
                command = new SQLiteCommand(sql, m_dbConnection);
                command.ExecuteNonQuery();

                int y = x - 1;
                sql = "update  num set num= '"+ x.ToString() +"' where num ='"+y.ToString()+"'";
                command = new SQLiteCommand(sql, m_dbConnection);
                command.ExecuteNonQuery();
                wasel = x.ToString();
                m_dbConnection.Close();
                printDocument1.Print();
                clear_f();
                if (i != 0)
                    MessageBox.Show("Â–« «·”Ã· „ÊÃÊœ „”»ﬁ« Ê „ «” »œ«·Â »‰Ã«Õ");
                else
                    MessageBox.Show(" „ «·Õ›Ÿ »‰Ã«Õ");
                

            }

            catch (Exception ex)
            {
                MessageBox.Show("ÌÊÃœ ·œÌﬂ „‘ﬂ·…");
            } 
        
        }

        private bool check_date2()
        {
            string[] date1 = new string[3];
            date1 = textBox5.Text.Split('/'); 
            int y1 = int.Parse(date1[2]); int m1 = int.Parse(date1[1]); int day1 = int.Parse(date1[0]);
            DateTime t1 = new DateTime(y1, m1, day1);
            DateTime t2 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            DateTime t3 = new DateTime(DateTime.Today.AddDays(14).Year, DateTime.Today.AddDays(14).Month, DateTime.Today.AddDays(14).Day);
            //MessageBox.Show(t1.ToString()) ; MessageBox.Show(t2.ToString()); MessageBox.Show(t3.ToString());
            int st1 = DateTime.Compare(t1, t2);
            int st2 = DateTime.Compare(t1, t3);
            if (st1 > 0 && st2 <= 0) return true; else return false;
        }

        private void clear_f()
        {
            textBox7.Text = ""; textBox8.Text = ""; textBox9.Text = "";
            textBox10.Text = ""; textBox6.Text = ""; textBox8.Enabled = true;
            checkBox1.Checked = false;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (textBox8.Enabled)
            {
                DialogResult result3 = MessageBox.Show("Â· √‰  „ √ﬂœ „‰ «·≈·€«¡?",
                 " Õ–Ì—",
                 MessageBoxButtons.YesNo,
                 MessageBoxIcon.Question,
                 MessageBoxDefaultButton.Button2);
                if (result3 == DialogResult.No)
                {
                    return;
                }
                else if (result3 == DialogResult.Yes)
                {
                    clear_f();
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            int i = 0;
            float sum = 0;
            try
            {
                dataGridView2.Rows.Clear();
                SQLiteConnection m_dbConnection;
                m_dbConnection = new SQLiteConnection("Data Source=helal.db;Version=3;");
                m_dbConnection.Open();
                 string sql="";
                 if (comboBox2.Text == "«·ﬂ·")
                     sql = "select * from visit";
                 else sql = "select * from visit where clinic='"+comboBox2.Text+"'";
                SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                SQLiteDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (check_date1((string)reader["date"]))
                    {
                        dataGridView2.Rows.Add();
                        dataGridView2.Rows[i].Cells[0].Value = reader["name"];
                        dataGridView2.Rows[i].Cells[1].Value = reader["id"];
                        dataGridView2.Rows[i].Cells[5].Value = reader["fee"];
                        dataGridView2.Rows[i].Cells[2].Value = reader["date"];
                        dataGridView2.Rows[i].Cells[3].Value = reader["clinic"];
                        dataGridView2.Rows[i].Cells[4].Value = reader["type"];
                        dataGridView2.Rows[i].Cells[6].Value = reader["num"];
                        sum = sum + float.Parse((string)reader["fee"]);
                        i++;
                    }
                }
                reader.Close();
                m_dbConnection.Close();
                textBox13.Text = sum.ToString();
            }

            catch (Exception ex)
            {
                MessageBox.Show("ÌÊÃœ ·œÌﬂ „‘ﬂ·…");
            }
        }

        private bool check_date1(string p)
        {
            string[] date1 = new string[3];
            string[] date2 = new string[3];
            string[] date3 = new string[3];
            date1 = p.Split('/'); date2 = textBox11.Text.Split('/'); date3 = textBox12.Text.Split('/');
            int y1 = int.Parse(date1[2]); int m1 = int.Parse(date1[1]); int day1 = int.Parse(date1[0]);
            int y2 = int.Parse(date2[2]); int m2 = int.Parse(date2[1]); int day2 = int.Parse(date2[0]);
            int y3 = int.Parse(date3[2]); int m3 = int.Parse(date3[1]); int day3 = int.Parse(date3[0]);
            DateTime t1 = new DateTime(y1, m1, day1);
            DateTime t2 = new DateTime(y2, m2, day2);
            DateTime t3 = new DateTime(y3, m3, day3);
            //MessageBox.Show(t1.ToString()) ; MessageBox.Show(t2.ToString()); MessageBox.Show(t3.ToString());
            int st1 = DateTime.Compare(t1, t2);
            int st2 = DateTime.Compare(t1, t3);
            if (st1 >= 0 && st2 <= 0) return true; else return false;
        }
        private bool check_date3(string p)
        {
            string[] date1 = new string[3];
            string[] date2 = new string[3];
            string[] date3 = new string[3];
            date1 = p.Split('/'); date2 = textBox24.Text.Split('/'); date3 = textBox23.Text.Split('/');
            int y1 = int.Parse(date1[2]); int m1 = int.Parse(date1[1]); int day1 = int.Parse(date1[0]);
            int y2 = int.Parse(date2[2]); int m2 = int.Parse(date2[1]); int day2 = int.Parse(date2[0]);
            int y3 = int.Parse(date3[2]); int m3 = int.Parse(date3[1]); int day3 = int.Parse(date3[0]);
            DateTime t1 = new DateTime(y1, m1, day1);
            DateTime t2 = new DateTime(y2, m2, day2);
            DateTime t3 = new DateTime(y3, m3, day3);
            //MessageBox.Show(t1.ToString()) ; MessageBox.Show(t2.ToString()); MessageBox.Show(t3.ToString());
            int st1 = DateTime.Compare(t1, t2);
            int st2 = DateTime.Compare(t1, t3);
            if (st1 >= 0 && st2 <= 0) return true; else return false;
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            StringFormat str = new StringFormat();
            str.Alignment = StringAlignment.Center;
            str.LineAlignment = StringAlignment.Center;
            str.Trimming = StringTrimming.EllipsisCharacter;
            e.Graphics.DrawString("«·—⁄«Ì… «·’ÕÌ… «·√Ê·Ì…", new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(220, 245, 400, 60), str);
            e.Graphics.DrawString("”‰œ ﬁ»÷", new System.Drawing.Font(FontFamily.GenericSansSerif, 12, FontStyle.Regular), Brushes.Black, new RectangleF(30, 30, 60, 60), str);
            e.Graphics.DrawString("«·⁄Ì«œ… : "+comboBox1.Text, new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(220, 285, 400, 60), str);
            e.Graphics.DrawString("Ã„⁄Ì… «·Â·«· «·√Õ„— «·›·”ÿÌ‰Ì - ”·›Ì ", new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(220, 200, 400, 60), str);
            e.Graphics.DrawRectangle(Pens.Black, 15, 15, 797, 600);
            str.LineAlignment = StringAlignment.Far;
            str.Alignment = StringAlignment.Far;
            e.Graphics.DrawString("«· «—ÌŒ : "+textBox5.Text, new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(380, 360, 400, 60), str);
            e.Graphics.DrawString("«”„ «·„—Ì÷ : " + textBox10.Text, new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(280, 405, 500, 60), str);
            e.Graphics.DrawString("‰Ê⁄ «·⁄·«Ã : " + comboBox3.Text, new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(380, 450, 400, 60), str);
            e.Graphics.DrawString("«·—”Ê„ : " + textBox6.Text, new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(380, 495, 400, 60), str);
            e.Graphics.DrawString("—ﬁ„ «·ÂÊÌ… : " + textBox8.Text, new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(0, 405, 280, 60), str);
            //str.LineAlignment = StringAlignment.Center;
            //str.Alignment = StringAlignment.Center;
            e.Graphics.DrawString("—ﬁ„ «·Ê’· : "+wasel, new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(0, 360, 280, 60), str);

            string d=Directory.GetCurrentDirectory() + "\\ss.png";
            Bitmap image = new Bitmap(d);
            e.Graphics.DrawImage(image,new System.Drawing.Rectangle(310,50,image.Width, image.Height));

           
        }

        private void button11_Click(object sender, EventArgs e)
        {
           if (dataGridView2.RowCount < 2) { MessageBox.Show("«·ÃœÊ· ›«—€"); return; }
            //---
            Process[] processlist = Process.GetProcesses();

            foreach (Process theprocess in processlist)
            {
                if (theprocess.ProcessName == "EXCEL")
                {
                    MessageBox.Show("√€·ﬁ Ã„Ì⁄ „·›«  «ﬂ”·");
                    return;
                }
            }
            //----
            try
            {
                ApplicationClass app = new ApplicationClass();

                workBook1 = app.Workbooks.Open(Directory.GetCurrentDirectory() + "\\tmp.xlsx", 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                workSheet1 = (Worksheet)workBook1.ActiveSheet;
                //read from datagrid to excel
                for (int i = 0; i < dataGridView2.RowCount-1; i++)
                {

                    ((Range)workSheet1.Cells[i + 2, 1]).Value2 = dataGridView2.Rows[i].Cells[0].Value;
                    ((Range)workSheet1.Cells[i + 2, 2]).Value2=dataGridView2.Rows[i].Cells[1].Value;
                    ((Range)workSheet1.Cells[i + 2, 3]).Value2=dataGridView2.Rows[i].Cells[2].Value;
                    ((Range)workSheet1.Cells[i + 2, 4]).Value2=dataGridView2.Rows[i].Cells[3].Value;
                    ((Range)workSheet1.Cells[i + 2, 5]).Value2=dataGridView2.Rows[i].Cells[4].Value;
                    ((Range)workSheet1.Cells[i + 2, 6]).Value2 = dataGridView2.Rows[i].Cells[5].Value;
                    ((Range)workSheet1.Cells[i + 2, 7]).Value2 = dataGridView2.Rows[i].Cells[6].Value;

                    
                }
                ((Range)workSheet1.Cells[ 2, 8]).Value2 = textBox13.Text;
                //close
                //MessageBox.Show(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\ok" + ".xlsx");
                workBook1.Close(true, Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\"+comboBox2.Text + ".xlsx", false);
                //--------------------------------------------------------------------------------------------------
                //finish
                kill_excel();
                MessageBox.Show(" „  «·⁄„·Ì… »‰Ã«Õ");
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);
                workBook1.Close(false, Directory.GetCurrentDirectory() + "\\tmp.xlsx", false);
                kill_excel();
                return;
            }

        }
        private void kill_excel()
        {
            Process[] processlist = Process.GetProcesses();

            foreach (Process theprocess in processlist)
            {
                if (theprocess.ProcessName == "EXCEL")
                {
                    theprocess.Kill();
                    return;
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            int i = 0;
            float sum = 0;
            try
            {
                dataGridView3.Rows.Clear();
                SQLiteConnection m_dbConnection;
                m_dbConnection = new SQLiteConnection("Data Source=helal.db;Version=3;");
                m_dbConnection.Open();
                string sql = "";
                if (radioButton1.Checked)
                    sql = "select * from visit where id='"+textBox14.Text+"'";
                else if (radioButton2.Checked)
                    sql = "select * from visit where num='" + textBox14.Text + "'";
                else sql = "select * from visit where name like '%" + textBox14.Text + "%'";
                SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                SQLiteDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                        dataGridView3.Rows.Add();
                        dataGridView3.Rows[i].Cells[0].Value = reader["name"];
                        dataGridView3.Rows[i].Cells[1].Value = reader["id"];
                        dataGridView3.Rows[i].Cells[5].Value = reader["fee"];
                        dataGridView3.Rows[i].Cells[2].Value = reader["date"];
                        dataGridView3.Rows[i].Cells[3].Value = reader["clinic"];
                        dataGridView3.Rows[i].Cells[4].Value = reader["type"];
                        dataGridView3.Rows[i].Cells[6].Value = reader["num"];
                        i++;
                }
                if (i == 0) { MessageBox.Show("·„ Ì „ «·⁄ÀÊ— ⁄·Ï ”Ã·« "); }
                reader.Close();
                m_dbConnection.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show("ÌÊÃœ ·œÌﬂ „‘ﬂ·…");
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (dataGridView3.RowCount < 2) { MessageBox.Show("·« ”Ã·« "); return; }
            if (dataGridView3.SelectedRows.Count != 1) { MessageBox.Show("«Œ — ”Ã· Ê«Õœ ›ﬁÿ"); return; }
            index = dataGridView3.SelectedRows[0].Index;
            printDocument2.Print();
        }

        private void printDocument2_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            StringFormat str = new StringFormat();
            str.Alignment = StringAlignment.Center;
            str.LineAlignment = StringAlignment.Center;
            str.Trimming = StringTrimming.EllipsisCharacter;
            e.Graphics.DrawString("«·—⁄«Ì… «·’ÕÌ… «·√Ê·Ì…", new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(220, 245, 400, 60), str);
            e.Graphics.DrawString("”‰œ ﬁ»÷", new System.Drawing.Font(FontFamily.GenericSansSerif, 12, FontStyle.Regular), Brushes.Black, new RectangleF(30, 30, 60, 60), str);
            e.Graphics.DrawString("«·⁄Ì«œ… : " + dataGridView3.Rows[index].Cells[3].Value.ToString(), new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(220, 285, 400, 60), str);
            e.Graphics.DrawString("Ã„⁄Ì… «·Â·«· «·√Õ„— «·›·”ÿÌ‰Ì - ”·›Ì ", new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(220, 200, 400, 60), str);
            e.Graphics.DrawRectangle(Pens.Black, 15, 15, 797, 600);
            str.LineAlignment = StringAlignment.Far;
            str.Alignment = StringAlignment.Far;
            e.Graphics.DrawString("«· «—ÌŒ : " + dataGridView3.Rows[index].Cells[2].Value.ToString(), new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(380, 360, 400, 60), str);
            e.Graphics.DrawString("«”„ «·„—Ì÷ : " + dataGridView3.Rows[index].Cells[0].Value.ToString(), new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(280, 405, 500, 60), str);
            e.Graphics.DrawString("‰Ê⁄ «·⁄·«Ã : " + dataGridView3.Rows[index].Cells[4].Value.ToString(), new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(380, 450, 400, 60), str);
            e.Graphics.DrawString("«·—”Ê„ : " + dataGridView3.Rows[index].Cells[5].Value.ToString(), new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(380, 495, 400, 60), str);
            e.Graphics.DrawString("—ﬁ„ «·ÂÊÌ… : " + dataGridView3.Rows[index].Cells[1].Value.ToString(), new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(0, 405, 280, 60), str);
            //str.LineAlignment = StringAlignment.Center;
            //str.Alignment = StringAlignment.Center;
            e.Graphics.DrawString("—ﬁ„ «·Ê’· : " + dataGridView3.Rows[index].Cells[6].Value.ToString(), new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(0, 360, 280, 60), str);

            string d = Directory.GetCurrentDirectory() + "\\ss.png";
            Bitmap image = new Bitmap(d);
            e.Graphics.DrawImage(image, new System.Drawing.Rectangle(310, 50, image.Width, image.Height));
        }

        private void button14_Click(object sender, EventArgs e)
        {try{
            if (dataGridView3.RowCount < 2) { MessageBox.Show("·« ”Ã·« "); return; }
            if (dataGridView3.SelectedRows.Count != 1) { MessageBox.Show("«Œ — ”Ã· Ê«Õœ ›ﬁÿ"); return; }
            DialogResult result3 = MessageBox.Show("Â· √‰  „ √ﬂœ „‰ «·Õ–›?",
                 " Õ–Ì—",
                 MessageBoxButtons.YesNo,
                 MessageBoxIcon.Question,
                 MessageBoxDefaultButton.Button2);
            if (result3 == DialogResult.No)
            {
                return;
            }
            
            int i=0;
            SQLiteConnection m_dbConnection;
            m_dbConnection = new SQLiteConnection("Data Source=helal.db;Version=3;");
            m_dbConnection.Open();
            string sql = "delete from visit where num='" + dataGridView3.SelectedRows[0].Cells[6].Value.ToString() + "'";
            SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
            command = new SQLiteCommand(sql, m_dbConnection);
            command.ExecuteNonQuery();
            m_dbConnection.Close();
            dataGridView3.Rows.RemoveAt(dataGridView3.SelectedRows[0].Index);
            }

            catch (Exception ex)
            {
                MessageBox.Show("ÌÊÃœ ·œÌﬂ „‘ﬂ·…");
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            { textBox8.Enabled = false; textBox8.Text = "000000000"; }
            else {textBox8.Enabled = true; textBox8.Text = ""; }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            monthCalendar4.Visible = true;
        }

        private void monthCalendar4_DateSelected(object sender, DateRangeEventArgs e)
        {
            string d, y, m;
            y = monthCalendar4.SelectionRange.Start.Year.ToString();
            d = monthCalendar4.SelectionRange.Start.Day.ToString();
            m = monthCalendar4.SelectionRange.Start.Month.ToString();
            textBox16.Text = d + "/" + m + "/" + y;
            monthCalendar4.Visible = false;
        }

        private void monthCalendar4_MouseLeave(object sender, EventArgs e)
        {
            monthCalendar4.Visible = false; 
        }

        

        private void button15_Click(object sender, EventArgs e)
        {
            float j1;
            if (textBox15.Text.Trim().Length < 3) { MessageBox.Show("ÌÃ» √‰ ÌﬂÊ‰ «·«”„ √ﬂÀ— „‰ Õ—›Ì‰"); return; }
            if (!float.TryParse(textBox17.Text, out j1)) { MessageBox.Show("√œŒ· «·«ÌÃ«— »«·‘ﬂ· «·’ÕÌÕ"); return; }

            try
            {
                string num="";
                SQLiteConnection m_dbConnection;
                m_dbConnection = new SQLiteConnection("Data Source=helal.db;Version=3;");
                m_dbConnection.Open();
                string sql = "";
                SQLiteCommand command; SQLiteDataReader reader;
                 
                //------------------
                sql = "select * from num";
                command = new SQLiteCommand(sql, m_dbConnection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    num = (string)reader["num2"];
                }
                reader.Close();
                int x = int.Parse(num); x++;
                //------------------
                sql = "insert into rent values ('" +comboBox5.Text + "','" + textBox15.Text + "','" + textBox17.Text + "','" + textBox16.Text + "','" + textBox21.Text +  "','" + x.ToString() + "')";
                command = new SQLiteCommand(sql, m_dbConnection);
                command.ExecuteNonQuery();

                int y = x - 1;
                sql = "update  num set num2= '" + x.ToString() + "' where num2 ='" + y.ToString() + "'";
                command = new SQLiteCommand(sql, m_dbConnection);
                command.ExecuteNonQuery();
                w = x.ToString();
                m_dbConnection.Close();
                printDocument3.Print();
                clear_f2();
                    MessageBox.Show(" „ «·Õ›Ÿ »‰Ã«Õ");


            }

            catch (Exception ex)
            {
                MessageBox.Show("ÌÊÃœ ·œÌﬂ „‘ﬂ·…");
            } 

        }

        private void clear_f2()
        {
            textBox15.Text = "";
            textBox17.Text = "";
            textBox21.Text = "";
        }

        private void printDocument3_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            StringFormat str = new StringFormat();
            str.Alignment = StringAlignment.Center;
            str.LineAlignment = StringAlignment.Center;
            str.Trimming = StringTrimming.EllipsisCharacter;
            e.Graphics.DrawString("«” ∆Ã«—", new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(220, 245, 400, 60), str);
            e.Graphics.DrawString("”‰œ ﬁ»÷", new System.Drawing.Font(FontFamily.GenericSansSerif, 12, FontStyle.Regular), Brushes.Black, new RectangleF(30, 30, 60, 60), str);
            if (comboBox5.Text!="«·ﬁ«⁄…")
            e.Graphics.DrawString("«·„Œ“‰ : " + comboBox5.Text, new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(220, 285, 400, 60), str);
            else
            e.Graphics.DrawString("«·ﬁ«⁄…", new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(220, 285, 400, 60), str);
            e.Graphics.DrawString("Ã„⁄Ì… «·Â·«· «·√Õ„— «·›·”ÿÌ‰Ì - ”·›Ì ", new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(220, 200, 400, 60), str);
            e.Graphics.DrawRectangle(Pens.Black, 15, 15, 797, 600);
            str.LineAlignment = StringAlignment.Far;
            str.Alignment = StringAlignment.Far;
            e.Graphics.DrawString("«· «—ÌŒ : " + textBox16.Text, new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(380, 360, 400, 60), str);
            e.Graphics.DrawString("«·«”„  : " + textBox15.Text, new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(280, 405, 500, 60), str);
            e.Graphics.DrawString("«·«ÌÃ«— : " + textBox17.Text, new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(380, 450, 400, 60), str);
            if (comboBox5.Text != "«·ﬁ«⁄…")
            e.Graphics.DrawString("«·‘ÂÊ— : " + textBox21.Text, new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(380, 495, 400, 60), str);
            else
            e.Graphics.DrawString("«·«Ì«„ : " + textBox21.Text, new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(380, 495, 400, 60), str);
            e.Graphics.DrawString("—ﬁ„ «·Ê’· : " + w, new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(0, 360, 280, 60), str);
            string d = Directory.GetCurrentDirectory() + "\\ss.png";
            Bitmap image = new Bitmap(d);
            e.Graphics.DrawImage(image, new System.Drawing.Rectangle(310, 50, image.Width, image.Height));
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox5.Text != "«·ﬁ«⁄…") label28.Text = "«·‘ÂÊ—";
            else label28.Text = "«·«Ì«„";
        }

        private void button21_Click(object sender, EventArgs e)
        {
            int i = 0;
            float sum = 0;
            try
            {
                dataGridView4.Rows.Clear();
                SQLiteConnection m_dbConnection;
                m_dbConnection = new SQLiteConnection("Data Source=helal.db;Version=3;");
                m_dbConnection.Open();
                string sql = "";
                if (comboBox6.Text == "«·ﬂ·")
                    sql = "select * from rent";
                else sql = "select * from rent where m='" + comboBox6.Text + "'";
                SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                SQLiteDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (check_date3((string)reader["date"]))
                    {
                        dataGridView4.Rows.Add();
                        dataGridView4.Rows[i].Cells[0].Value = reader["m"];
                        dataGridView4.Rows[i].Cells[1].Value = reader["name"];
                        dataGridView4.Rows[i].Cells[2].Value = reader["rent"];
                        dataGridView4.Rows[i].Cells[3].Value = reader["date"];
                        dataGridView4.Rows[i].Cells[4].Value = reader["int"];
                        dataGridView4.Rows[i].Cells[5].Value = reader["num"];
                        sum = sum + float.Parse((string)reader["rent"]);
                        i++;
                    }
                }
                reader.Close();
                m_dbConnection.Close();
                textBox22.Text = sum.ToString();
            }

            catch (Exception ex)
            {
                MessageBox.Show("ÌÊÃœ ·œÌﬂ „‘ﬂ·…");
            }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            monthCalendar7.Visible = true;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            monthCalendar6.Visible = true;
        }

        private void monthCalendar7_MouseLeave(object sender, EventArgs e)
        {
            monthCalendar7.Visible = false;
        }

        private void monthCalendar6_MouseLeave(object sender, EventArgs e)
        {
            monthCalendar6.Visible = false;
        }

        private void monthCalendar7_DateSelected(object sender, DateRangeEventArgs e)
        {
            string d, y, m;
            y = monthCalendar7.SelectionRange.Start.Year.ToString();
            d = monthCalendar7.SelectionRange.Start.Day.ToString();
            m = monthCalendar7.SelectionRange.Start.Month.ToString();
            textBox24.Text = d + "/" + m + "/" + y;
            monthCalendar7.Visible = false;
        }

        private void monthCalendar6_DateSelected(object sender, DateRangeEventArgs e)
        {
            string d, y, m;
            y = monthCalendar6.SelectionRange.Start.Year.ToString();
            d = monthCalendar6.SelectionRange.Start.Day.ToString();
            m = monthCalendar6.SelectionRange.Start.Month.ToString();
            textBox23.Text = d + "/" + m + "/" + y;
            monthCalendar6.Visible = false;
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if (dataGridView4.RowCount < 2) { MessageBox.Show("«·ÃœÊ· ›«—€"); return; }
            //---
            Process[] processlist = Process.GetProcesses();

            foreach (Process theprocess in processlist)
            {
                if (theprocess.ProcessName == "EXCEL")
                {
                    MessageBox.Show("√€·ﬁ Ã„Ì⁄ „·›«  «ﬂ”·");
                    return;
                }
            }
            //----
            try
            {
                ApplicationClass app = new ApplicationClass();

                workBook1 = app.Workbooks.Open(Directory.GetCurrentDirectory() + "\\tmp1.xlsx", 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                workSheet1 = (Worksheet)workBook1.ActiveSheet;
                //read from datagrid to excel
                for (int i = 0; i < dataGridView4.RowCount - 1; i++)
                {

                    ((Range)workSheet1.Cells[i + 2, 1]).Value2 = dataGridView4.Rows[i].Cells[0].Value;
                    ((Range)workSheet1.Cells[i + 2, 2]).Value2 = dataGridView4.Rows[i].Cells[1].Value;
                    ((Range)workSheet1.Cells[i + 2, 3]).Value2 = dataGridView4.Rows[i].Cells[2].Value;
                    ((Range)workSheet1.Cells[i + 2, 4]).Value2 = dataGridView4.Rows[i].Cells[3].Value;
                    ((Range)workSheet1.Cells[i + 2, 5]).Value2 = dataGridView4.Rows[i].Cells[4].Value;
                    ((Range)workSheet1.Cells[i + 2, 6]).Value2 = dataGridView4.Rows[i].Cells[5].Value;


                }
                ((Range)workSheet1.Cells[2, 7]).Value2 = textBox22.Text;
                //close
                //MessageBox.Show(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\ok" + ".xlsx");
                workBook1.Close(true, Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\" + comboBox6.Text+"-«ÌÃ«—« " + ".xlsx", false);
                //--------------------------------------------------------------------------------------------------
                //finish
                kill_excel();
                MessageBox.Show(" „  «·⁄„·Ì… »‰Ã«Õ");
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);
                workBook1.Close(false, Directory.GetCurrentDirectory() + "\\tmp1.xlsx", false);
                kill_excel();
                return;
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            int i = 0;
            float sum = 0;
            try
            {
                dataGridView6.Rows.Clear();
                SQLiteConnection m_dbConnection;
                m_dbConnection = new SQLiteConnection("Data Source=helal.db;Version=3;");
                m_dbConnection.Open();
                string sql = "";
                if (radioButton5.Checked)
                    sql = "select * from rent where num='" + textBox25.Text + "'";
                else sql = "select * from rent where name like '%" + textBox25.Text + "%'";
                SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                SQLiteDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    dataGridView6.Rows.Add();
                    dataGridView6.Rows[i].Cells[0].Value = reader["m"];
                    dataGridView6.Rows[i].Cells[1].Value = reader["name"];
                    dataGridView6.Rows[i].Cells[2].Value = reader["rent"];
                    dataGridView6.Rows[i].Cells[3].Value = reader["date"];
                    dataGridView6.Rows[i].Cells[4].Value = reader["int"];
                    dataGridView6.Rows[i].Cells[5].Value = reader["num"];
                    i++;
                }
                if (i == 0) { MessageBox.Show("·„ Ì „ «·⁄ÀÊ— ⁄·Ï ”Ã·« "); }
                reader.Close();
                m_dbConnection.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show("ÌÊÃœ ·œÌﬂ „‘ﬂ·…");
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            if (dataGridView6.RowCount < 2) { MessageBox.Show("·« ”Ã·« "); return; }
            if (dataGridView6.SelectedRows.Count != 1) { MessageBox.Show("«Œ — ”Ã· Ê«Õœ ›ﬁÿ"); return; }
            index1 = dataGridView6.SelectedRows[0].Index;
            printDocument4.Print();
        }

        private void printDocument4_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            StringFormat str = new StringFormat();
            str.Alignment = StringAlignment.Center;
            str.LineAlignment = StringAlignment.Center;
            str.Trimming = StringTrimming.EllipsisCharacter;
            e.Graphics.DrawString("«” ∆Ã«—", new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(220, 245, 400, 60), str);
            e.Graphics.DrawString("”‰œ ﬁ»÷", new System.Drawing.Font(FontFamily.GenericSansSerif, 12, FontStyle.Regular), Brushes.Black, new RectangleF(30, 30, 60, 60), str);
            if (dataGridView6.Rows[index1].Cells[0].Value.ToString() != "«·ﬁ«⁄…")
                e.Graphics.DrawString("«·„Œ“‰ : " + comboBox5.Text, new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(220, 285, 400, 60), str);
            else
                e.Graphics.DrawString("«·ﬁ«⁄…", new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(220, 285, 400, 60), str);
            e.Graphics.DrawString("Ã„⁄Ì… «·Â·«· «·√Õ„— «·›·”ÿÌ‰Ì - ”·›Ì ", new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(220, 200, 400, 60), str);
            e.Graphics.DrawRectangle(Pens.Black, 15, 15, 797, 600);
            str.LineAlignment = StringAlignment.Far;
            str.Alignment = StringAlignment.Far;
            e.Graphics.DrawString("«· «—ÌŒ : " + dataGridView6.Rows[index1].Cells[3].Value.ToString(), new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(380, 360, 400, 60), str);
            e.Graphics.DrawString("«·«”„  : " + dataGridView6.Rows[index1].Cells[1].Value.ToString(), new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(280, 405, 500, 60), str);
            e.Graphics.DrawString("«·«ÌÃ«— : " + dataGridView6.Rows[index1].Cells[2].Value.ToString(), new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(380, 450, 400, 60), str);
            if (dataGridView6.Rows[index1].Cells[0].Value.ToString() != "«·ﬁ«⁄…")
                e.Graphics.DrawString("«·‘ÂÊ— : " + dataGridView6.Rows[index1].Cells[4].Value.ToString(), new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(380, 495, 400, 60), str);
            else
                e.Graphics.DrawString("«·«Ì«„ : " + dataGridView6.Rows[index1].Cells[4].Value.ToString(), new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(380, 495, 400, 60), str);
            e.Graphics.DrawString("—ﬁ„ «·Ê’· : " + dataGridView6.Rows[index1].Cells[5].Value.ToString(), new System.Drawing.Font(FontFamily.GenericSansSerif, 16, FontStyle.Regular), Brushes.Black, new RectangleF(0, 360, 280, 60), str);
            string d = Directory.GetCurrentDirectory() + "\\ss.png";
            Bitmap image = new Bitmap(d);
            e.Graphics.DrawImage(image, new System.Drawing.Rectangle(310, 50, image.Width, image.Height));
        }

        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView6.RowCount < 2) { MessageBox.Show("·« ”Ã·« "); return; }
                if (dataGridView6.SelectedRows.Count != 1) { MessageBox.Show("«Œ — ”Ã· Ê«Õœ ›ﬁÿ"); return; }
                DialogResult result3 = MessageBox.Show("Â· √‰  „ √ﬂœ „‰ «·Õ–›?",
                     " Õ–Ì—",
                     MessageBoxButtons.YesNo,
                     MessageBoxIcon.Question,
                     MessageBoxDefaultButton.Button2);
                if (result3 == DialogResult.No)
                {
                    return;
                }

                int i = 0;
                SQLiteConnection m_dbConnection;
                m_dbConnection = new SQLiteConnection("Data Source=helal.db;Version=3;");
                m_dbConnection.Open();
                string sql = "delete from rent where num='" + dataGridView6.SelectedRows[0].Cells[5].Value.ToString() + "'";
                SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
                command = new SQLiteCommand(sql, m_dbConnection);
                command.ExecuteNonQuery();
                m_dbConnection.Close();
                dataGridView6.Rows.RemoveAt(dataGridView6.SelectedRows[0].Index);
            }

            catch (Exception ex)
            {
                MessageBox.Show("ÌÊÃœ ·œÌﬂ „‘ﬂ·…");
            }
        }

 
     
       
       

    }
}