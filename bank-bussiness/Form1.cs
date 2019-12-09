using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace bank_bussiness
{

    public partial class Form1 : Form
    {
        private const string LECTURER = "lecturer";
        private const string SUBJECT = "subject";
        private const int MAX_MARK_RANGE = 10;
        private const int ONE= 1;
        private const int MAX_MONTH_VALUE = 12;
        private const int MAX_DAY_VALUE = 31;
        private const int MIN_YEAR_VALUE = 1970;
        private const double MAX_AVG_MARK_RANGE = 10.0;
        private const int MAX_HOURS_RANGE = 92;

        private DateTime minDateTime = new DateTime(1970, 1, 1);

        Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
        Document doc;

        public struct TableRow
        {
            public string subjectName;
            public int hours;
            public int mark;
            public double avgMark;
            public DateTime date;
            public string lecturerName;
        }

        DateTimePicker dtp = new DateTimePicker();
        System.Drawing.Rectangle _Rectangle;

        List<TableRow> rows = new List<TableRow>();

        public Form1()
        {
            InitializeComponent();
            dataGridView1.Controls.Add(dtp);
            dtp.Visible = false;
            dtp.Format = DateTimePickerFormat.Custom;
            dtp.TextChanged += new EventHandler(dtp_TextChange);
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (doc != null)
            {
                doc.Save();
                doc.Close();
                app.Quit();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = @"C:\Users\USER\Desktop";
                openFileDialog.Filter = "Office Files|*.docx;*.dotx";
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;
                    doc = app.Documents.Add(filePath);
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (doc == null)
            {
                MessageBox.Show("Please open the document.");
                return;
            }

            Table table = doc.Tables[1];

            Console.WriteLine(dataGridView1.Rows.Count);
            for (int idx = 0; idx < dataGridView1.Rows.Count - 1; idx++)
            {
                TableRow rowToAdd;
                rowToAdd.subjectName = dataGridView1.Rows[idx].Cells["subjectName"].Value.ToString();
                rowToAdd.hours = Int32.Parse(dataGridView1.Rows[idx].Cells["hours"].Value.ToString());
                rowToAdd.mark = Int32.Parse(dataGridView1.Rows[idx].Cells["mark"].Value.ToString());
                rowToAdd.avgMark = Double.Parse(dataGridView1.Rows[idx].Cells["avg_mark"].Value.ToString());
                rowToAdd.date = DateTime.Parse(dataGridView1.Rows[idx].Cells["date"].Value.ToString());
                rowToAdd.lecturerName = dataGridView1.Rows[idx].Cells["lecturerName"].Value.ToString();
                rows.Add(rowToAdd);
            }

            for (int j = 0; j < rows.Count - 1; j++)
                table.Rows.Add(table.Rows[2]);

            TableRow[] rowsArray = rows.ToArray();

            int i = -1;
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    Console.WriteLine(i);
                    if (cell.RowIndex != 1)
                    {
                        switch (cell.ColumnIndex)
                        {
                            case 1:
                                {
                                    cell.Range.Text = rowsArray[i].subjectName;
                                    break;
                                }
                            case 2:
                                {
                                    cell.Range.Text = rowsArray[i].hours.ToString();

                                    break;
                                }
                            case 3:
                                {
                                    cell.Range.Text = rowsArray[i].mark.ToString();
                                    break;
                                }
                            case 4:
                                {
                                    cell.Range.Text = rowsArray[i].avgMark.ToString();
                                    break;
                                }
                            case 5:
                                {
                                    cell.Range.Text = rowsArray[i].date.ToString().Substring(0, 10);
                                    break;
                                }
                            case 6:
                                {
                                    cell.Range.Text = rowsArray[i].lecturerName;
                                    break;
                                }
                        }
                    }
                }
                i++;
            }

            foreach (Microsoft.Office.Interop.Word.FormField field in doc.FormFields)
            {
                switch (field.Name)
                {
                    case "record_book_num":
                        {
                            field.Range.Text = numericUpDown1.Value.ToString();
                            break;
                        }
                    case "fio":
                        {
                            field.Range.Text = textBox2.Text;
                            break;
                        }
                    case "faculty":
                        {
                            field.Range.Text = textBox3.Text;
                            break;
                        }
                    case "speciality":
                        {
                            field.Range.Text = textBox4.Text;
                            break;
                        }
                    case "receipt_date":
                        {
                            field.Range.Text = dateTimePicker1.Value.ToString().Substring(0, 10);
                            break;
                        }
                    default:
                        break;
                }
            }

            if (radioButton1.Checked)
            {
                foreach (Microsoft.Office.Interop.Word.Section section in doc.Sections)
                {
                    //Get the header range and add the header details.
                    Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.InsertAfter("\n" + DateTime.Now.ToString());
                }
            }

            app.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            decimal genNumber = numericUpDown8.Value;

            int prevRowCount = dataGridView1.RowCount - 1;

            dataGridView1.RowCount += (int)genNumber;

            Random random = new Random();
            for (int idx = prevRowCount; idx < dataGridView1.RowCount - 1; idx++)
            {
                dataGridView1.Rows[idx].Cells[0].Value = SUBJECT + random.Next(DateTime.Today.Year);
                dataGridView1.Rows[idx].Cells[1].Value = random.Next(MAX_HOURS_RANGE);
                dataGridView1.Rows[idx].Cells[2].Value = random.Next(MAX_MARK_RANGE);
                dataGridView1.Rows[idx].Cells[3].Value = (random.NextDouble() * 10).ToString().Substring(0, 4);

                dataGridView1.Rows[idx].Cells[4].Value =
                    minDateTime.AddDays(random.Next((DateTime.Today - minDateTime).Days)).ToString().Split(' ')[0];
                dataGridView1.Rows[idx].Cells[5].Value = LECTURER + random.Next(DateTime.Today.Year);
                
            }
            Console.WriteLine("Generated \"" + genNumber + "\" rows");
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == -1)
                return;

            switch (dataGridView1.Columns[e.ColumnIndex].Name)
            {
                case "date":
                    _Rectangle = dataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
                    dtp.Size = new Size(_Rectangle.Width, _Rectangle.Height);
                    dtp.Location = new System.Drawing.Point(_Rectangle.X, _Rectangle.Y);
                    dtp.Visible = true;

                    break;
            }
        }

        private void dtp_TextChange(object sender, EventArgs e)
        {
            dataGridView1.CurrentCell.Value = dtp.Text.ToString();
        }

        private void dataGridView1_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            dtp.Visible = false;
        }

        private void dataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            dtp.Visible = false;
        }

    }
}
