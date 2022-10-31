using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Threading;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private static string splitString(string str, char key, int sel)
        {
            string[] arrString = str.Split(key);
            return arrString[sel];
        }
        private static string DayOfWeek(string date)
        {
            DateTime dateTime = new DateTime(int.Parse(splitString(date, '.', 2)), int.Parse(splitString(date, '.', 1)), int.Parse(splitString(date, '.', 0)));
            return dateTime.ToString("dddd");
        }
        private static string TotalWorkingTime(string timeIn, string timeOut)
        {
            // If 12h < Time-in/Time-out < 13h then working time = 0 
            if (DateTime.Parse(timeIn) >= DateTime.Parse("12:00:00") && DateTime.Parse(timeIn) <= DateTime.Parse("13:00:00") && DateTime.Parse(timeOut) > DateTime.Parse("12:00:00") && DateTime.Parse(timeOut) < DateTime.Parse("13:00:00"))
            {
                return "0.00";
            }
            else
            {
                // If 12h < Time-in < 13h then Time-in = 13h 
                if (DateTime.Parse(timeIn) > DateTime.Parse("12:00:00") && DateTime.Parse(timeIn) < DateTime.Parse("13:00:00"))
                {
                    timeIn = "13:00:00";
                }
                // If 12h < Time-out < 13h then Time-in = 12h 
                if (DateTime.Parse(timeOut) > DateTime.Parse("12:00:00") && DateTime.Parse(timeOut) < DateTime.Parse("13:00:00"))
                {
                    timeOut = "12:00:00";
                }

                TimeSpan timeSpan = DateTime.Parse(timeOut) - DateTime.Parse(timeIn);

                if (DateTime.Parse(timeIn) >= DateTime.Parse("13:00:00") || DateTime.Parse(timeOut) <= DateTime.Parse("12:00:00"))
                {
                    return (timeSpan.Hours + (float)timeSpan.Minutes / 60).ToString("0.0");
                }
                else
                {
                    return (timeSpan.Hours - 1 + (float)timeSpan.Minutes / 60).ToString("0.00");
                }
            }
        }
        private static void findMinMaxDate(string time, ref DateTime minDate, ref DateTime maxDate)
        {
            string date = splitString(time, ' ', 0);
            DateTime d = new DateTime(int.Parse(splitString(date, '.', 2)), int.Parse(splitString(date, '.', 1)), int.Parse(splitString(date, '.', 0)));
            if (d < minDate)
            {
                minDate = d;
            }
            if (d > maxDate)
            {
                maxDate = d;
            }
        }
        private void btnGenerate_Click(object sender, EventArgs e)
        {
            lState.Text = "File is opening...";

            //Kill background process Excel
            foreach (Process process in Process.GetProcesses())
            {
                if (process.ProcessName.Equals("EXCEL") && process.MainWindowHandle.ToString() == "0")
                {
                    process.Kill();
                }
            }

            ThreadStart threadStrart = new ThreadStart(Processing);
            Thread thread = new Thread(threadStrart);
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
        }

        private void Processing()
        {
            try
            {
                int j = 2;
                int c = 0;
                string tIn = "";
                string tOut = "";
                Boolean flagName = false;

                String D = "";

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(@"" + InputFile.Text);
                //Excel.Worksheet xlWorkSheet = xlWorkBook.Sheets[1];

                Global.listSheet = new string[xlWorkBook.Worksheets.Count];

                foreach (Excel.Worksheet s in xlWorkBook.Worksheets)
                {
                    Global.listSheet[c] = s.Name;
                    c++;
                }
                ConfigDialog configDialog = new ConfigDialog();
                configDialog.ShowDialog();
                if (configDialog.DialogResult == DialogResult.OK)
                {
                    Excel.Worksheet xlWorkSheet = xlWorkBook.Sheets[Global.sheetIndex];
                    Excel.Range xlRange = xlWorkSheet.UsedRange;

                    Excel.Workbook xlWorkBook1;
                    Excel.Worksheet xlWorkSheet1;
                    object missValue = System.Reflection.Missing.Value;

                    xlWorkBook1 = xlApp.Workbooks.Add(missValue);
                    xlWorkSheet1 = (Excel.Worksheet)xlWorkBook1.Worksheets.get_Item(1);

                    int rowCount = xlRange.Rows.Count;

                    D = splitString(xlRange.Cells[2, 1].Value.ToString(), ' ', 0);
                    DateTime minDate = new DateTime(int.Parse(splitString(D, '.', 2)), int.Parse(splitString(D, '.', 1)), int.Parse(splitString(D, '.', 0)));
                    DateTime maxDate = new DateTime(int.Parse(splitString(D, '.', 2)), int.Parse(splitString(D, '.', 1)), int.Parse(splitString(D, '.', 0)));

                    for (int i = 2; i <= rowCount; i++)
                    {
                        findMinMaxDate(xlRange.Cells[i, 1].Value.ToString(), ref minDate, ref maxDate);
                    }

                    xlWorkSheet1.Cells[1, 1] = "Staff's Name";
                    xlWorkSheet1.Cells[1, 2] = "Date";
                    xlWorkSheet1.Cells[1, 3] = "Day";
                    xlWorkSheet1.Cells[1, 4] = "Time in";
                    xlWorkSheet1.Cells[1, 5] = "Time out";
                    xlWorkSheet1.Cells[1, 6] = "Total Working Time";

                    DateTime d = minDate;

                    while (d.ToString("dd.MM.yyyy") != splitString(xlRange.Cells[2, 1].Value.ToString(), ' ', 0))
                    {
                        xlWorkSheet1.Cells[j, 1] = xlRange.Cells[2, 4].Value.ToString();
                        xlWorkSheet1.Cells[j, 2] = d.ToString("dd.MM.yyyy");
                        xlWorkSheet1.Cells[j, 3] = DayOfWeek(d.ToString("dd.MM.yyyy"));
                        j++;
                        if (d < maxDate)
                        {
                            d = d.AddDays(1);
                        }
                        else
                        {
                            d = minDate;
                        }
                    }
                    xlWorkSheet1.Cells[j, 1] = xlRange.Cells[2, 4].Value.ToString();
                    xlWorkSheet1.Cells[j, 2] = splitString(xlRange.Cells[2, 1].Value.ToString(), ' ', 0);
                    xlWorkSheet1.Cells[j, 3] = DayOfWeek(splitString(xlRange.Cells[2, 1].Value.ToString(), ' ', 0));

                    //Check-in
                    xlWorkSheet1.Cells[j, 4] = splitString(xlRange.Cells[2, 1].Value.ToString(), ' ', 1);
                    tIn = splitString(xlRange.Cells[2, 1].Value.ToString(), ' ', 1);

                    if (d < maxDate)
                    {
                        d = d.AddDays(1);
                    }
                    else
                    {
                        d = minDate;
                    }

                    for (int i = 3; i <= rowCount; i++)
                    {
                        if (splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 0) != splitString(xlRange.Cells[(i - 1), 1].Value.ToString(), ' ', 0) || xlRange.Cells[i, 4].Value.ToString() != xlRange.Cells[(i - 1), 4].Value.ToString())
                        {
                            if (tIn != null && tOut != null)
                            {
                                xlWorkSheet1.Cells[j, 6] = TotalWorkingTime(tIn, tOut);
                            }

                            //reset tIn, tOut
                            tIn = null;
                            tOut = null;

                            j++;

                            while (d.ToString("dd.MM.yyyy") != splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 0))
                            {
                                if (flagName == false)
                                {
                                    xlWorkSheet1.Cells[j, 1] = xlRange.Cells[i - 1, 4].Value.ToString();
                                }
                                else
                                {
                                    xlWorkSheet1.Cells[j, 1] = xlRange.Cells[i, 4].Value.ToString();
                                }

                                xlWorkSheet1.Cells[j, 2] = d.ToString("dd.MM.yyyy");
                                xlWorkSheet1.Cells[j, 3] = DayOfWeek(d.ToString("dd.MM.yyyy"));
                                j++;

                                if (d < maxDate)
                                {
                                    d = d.AddDays(1);
                                }
                                else
                                {
                                    d = minDate;
                                    flagName = true;
                                }
                            }
                            xlWorkSheet1.Cells[j, 1] = xlRange.Cells[i, 4].Value.ToString();
                            xlWorkSheet1.Cells[j, 2] = splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 0);
                            xlWorkSheet1.Cells[j, 3] = DayOfWeek(splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 0));
                            if (d < maxDate)
                            {
                                d = d.AddDays(1);
                                flagName = false;
                            }
                            else
                            {
                                d = minDate;
                                flagName = true;
                            }

                            //Check-in
                            xlWorkSheet1.Cells[j, 4] = splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 1);
                            tIn = splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 1);

                        }

                        //Check-out
                        xlWorkSheet1.Cells[j, 5] = splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 1);
                        tOut = splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 1);
                    }

                    if (tIn != null && tOut != null)
                    {
                        xlWorkSheet1.Cells[j, 6] = TotalWorkingTime(tIn, tOut);
                    }

                    if (d != minDate)
                    {
                        while (d <= maxDate)
                        {
                            j++;
                            xlWorkSheet1.Cells[j, 1] = xlRange.Cells[rowCount, 4].Value.ToString();
                            xlWorkSheet1.Cells[j, 2] = d.ToString("dd.MM.yyyy");
                            xlWorkSheet1.Cells[j, 3] = DayOfWeek(d.ToString("dd.MM.yyyy"));
                            d = d.AddDays(1);
                        }
                    }

                    //Set Column Width
                    xlWorkSheet1.Columns.AutoFit();

                    //Format Cells
                    xlWorkSheet1.get_Range("A1", "F1").Interior.Color = Excel.XlRgbColor.rgbYellow;
                    xlWorkSheet1.get_Range("A1", "F1").Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "Microsoft Excel Workbook|*.xls|Strict Open XML Spreadsheet|*.xlsx|All|*.*";
                    DialogResult result = saveFileDialog.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        xlWorkBook1.SaveAs(@"" + saveFileDialog.FileName);
                    }

                    xlWorkBook1.Close(false, missValue, missValue);

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    Marshal.ReleaseComObject(xlWorkSheet1);
                    Marshal.ReleaseComObject(xlWorkBook1);

                    Marshal.ReleaseComObject(xlRange);
                    Marshal.ReleaseComObject(xlWorkSheet);

                    xlWorkBook.Close(false, missValue, missValue);
                    Marshal.ReleaseComObject(xlWorkBook);

                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);

                    lState.Text = "Success";
                }
                else
                {
                    configDialog.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lState.Text = "Error";
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            DialogResult result = openFileDialog.ShowDialog();
            if(result == DialogResult.OK)
            {
                InputFile.Text = openFileDialog.FileName;
            }
        }
    }
}
