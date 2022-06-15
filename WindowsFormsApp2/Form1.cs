﻿using System;
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
            TimeSpan timeSpan = DateTime.Parse(timeOut) - DateTime.Parse(timeIn);

            if (DateTime.Parse(timeIn) > DateTime.Parse("13:00:00") || DateTime.Parse(timeOut) < DateTime.Parse("12:00:00"))
            {
                return (timeSpan.Hours + (float)timeSpan.Minutes / 60).ToString("0.0");
            }
            else
            {
                return (timeSpan.Hours - 1 + (float)timeSpan.Minutes / 60).ToString("0.00");
            }
        }
        private void btnGenerate_Click(object sender, EventArgs e)
        {
            label1.Text = "Processing...";

            int j = 2;
            string tIn = "";
            string tOut = "";
            Boolean flagIn = false;

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(@"D:\Data_processing\BGSV_TimeLog_Input.xlsx");
            Excel.Worksheet xlWorkSheet = xlWorkBook.Sheets[1];
            Excel.Range xlRange = xlWorkSheet.UsedRange;

            Excel.Workbook xlWorkBook1;
            Excel.Worksheet xlWorkSheet1;
            object missValue = System.Reflection.Missing.Value;

            xlWorkBook1 = xlApp.Workbooks.Add(missValue);
            xlWorkSheet1 = (Excel.Worksheet)xlWorkBook1.Worksheets.get_Item(1);

            int rowCount = xlRange.Rows.Count;

            xlWorkSheet1.Cells[1, 1] = "Staff's Name";
            xlWorkSheet1.Cells[1, 2] = "Date";
            xlWorkSheet1.Cells[1, 3] = "Day";
            xlWorkSheet1.Cells[1, 4] = "Time in";
            xlWorkSheet1.Cells[1, 5] = "Time out";
            xlWorkSheet1.Cells[1, 6] = "Total Working Time";

            /*xlWorkSheet1.Cells[2, 1] = xlRange.Cells[2, 4].Value.ToString();
            xlWorkSheet1.Cells[2, 2] = splitString(xlRange.Cells[2, 1].Value.ToString(), ' ', 0);
            xlWorkSheet1.Cells[2, 3] = DayOfWeek(splitString(xlRange.Cells[2, 1].Value.ToString(), ' ', 0));
            xlWorkSheet1.Cells[2, 4] = splitString(xlRange.Cells[2, 1].Value.ToString(), ' ', 1);
            tIn = splitString(xlRange.Cells[2, 1].Value.ToString(), ' ', 1);

            for (int i = 3; i <= rowCount; i++)
            {
                if(splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 0) != splitString(xlRange.Cells[(i - 1), 1].Value.ToString(), ' ', 0) || xlRange.Cells[i, 4].Value.ToString() != xlRange.Cells[(i - 1), 4].Value.ToString())
                {
                    xlWorkSheet1.Cells[j, 5] = splitString(xlRange.Cells[(i - 1), 1].Value.ToString(), ' ', 1);
                    tOut = splitString(xlRange.Cells[(i - 1), 1].Value.ToString(), ' ', 1);
                    xlWorkSheet1.Cells[j, 6] = TotalWorkingTime(tIn, tOut);
                    xlWorkSheet1.Cells[(j + 1), 1] = xlRange.Cells[i, 4].Value.ToString();
                    xlWorkSheet1.Cells[(j + 1), 2] = splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 0);
                    xlWorkSheet1.Cells[(j + 1), 3] = DayOfWeek(splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 0));
                    xlWorkSheet1.Cells[(j + 1), 4] = splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 1);
                    tIn = splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 1);
                    j++;
                }
            }

            xlWorkSheet1.Cells[j, 5] = splitString(xlRange.Cells[rowCount, 1].Value.ToString(), ' ', 1);
            tOut = splitString(xlRange.Cells[rowCount, 1].Value.ToString(), ' ', 1);
            xlWorkSheet1.Cells[j, 6] = TotalWorkingTime(tIn, tOut);*/
            xlWorkSheet1.Cells[2, 1] = xlRange.Cells[2, 4].Value.ToString();
            xlWorkSheet1.Cells[2, 2] = splitString(xlRange.Cells[2, 1].Value.ToString(), ' ', 0);
            xlWorkSheet1.Cells[2, 3] = DayOfWeek(splitString(xlRange.Cells[2, 1].Value.ToString(), ' ', 0));

            if (xlRange.Cells[2, 2].Value.ToString() == "entry reader 1")
            {
                xlWorkSheet1.Cells[2, 4] = splitString(xlRange.Cells[2, 1].Value.ToString(), ' ', 1);
                tIn = splitString(xlRange.Cells[2, 1].Value.ToString(), ' ', 1);
                flagIn = true;
            }

            for(int i = 3; i <= rowCount; i++)
            {
                if (splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 0) != splitString(xlRange.Cells[(i - 1), 1].Value.ToString(), ' ', 0) || xlRange.Cells[i, 4].Value.ToString() != xlRange.Cells[(i - 1), 4].Value.ToString())
                {
                    if (tIn != null && tOut != null)
                    {
                        xlWorkSheet1.Cells[j, 6] = TotalWorkingTime(tIn, tOut);
                    }
                    j++;
                    xlWorkSheet1.Cells[j, 1] = xlRange.Cells[i, 4].Value.ToString();
                    xlWorkSheet1.Cells[j, 2] = splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 0);
                    xlWorkSheet1.Cells[j, 3] = DayOfWeek(splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 0));
                    flagIn = false;
                }
                if (flagIn == false)
                {
                    if (xlRange.Cells[i, 2].Value.ToString() == "entry reader 1")
                    {
                        xlWorkSheet1.Cells[j, 4] = splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 1);
                        tIn = splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 1);
                        flagIn = true;
                    }
                }
                xlWorkSheet1.Cells[j, 5] = splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 1);
                tOut = splitString(xlRange.Cells[i, 1].Value.ToString(), ' ', 1);
            }

            if (tIn != null && tOut != null)
            {
                xlWorkSheet1.Cells[j, 6] = TotalWorkingTime(tIn, tOut);
            }

            //Set Column Width
            xlWorkSheet1.Columns.AutoFit();
            //xlWorkSheet1.Columns[1].ColumnWidth = 30;

            //Format Cells
            xlWorkSheet1.get_Range("A1", "F1").Interior.Color = Excel.XlRgbColor.rgbYellow;
            xlWorkSheet1.get_Range("A1", "F1").Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            xlWorkBook1.SaveAs("d:\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, missValue, missValue, missValue, missValue, Excel.XlSaveAsAccessMode.xlExclusive, missValue, missValue, missValue, missValue, missValue);
            xlWorkBook1.Close(true, missValue, missValue);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlWorkSheet1);
            Marshal.ReleaseComObject(xlWorkBook1);

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorkSheet);

            xlWorkBook.Close();
            Marshal.ReleaseComObject(xlWorkBook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            label1.Text = "Success";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string timeIn = "07:30:30";
            string timeOut = "09:20:30";
            TimeSpan timeSpan = DateTime.Parse(timeOut) - DateTime.Parse(timeIn);
            label1.Text = (timeSpan.Hours + (float)timeSpan.Minutes / 60).ToString("0.0");
        }
    }
}
