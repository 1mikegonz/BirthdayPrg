using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace BirthdayPrg
{
    public class Read_From_Excel
    {
        public static List<string[]> getExcelFile()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\michael.gonzalez\source\repos\BirthdayPrg\Book1.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            DateTime date = DateTime.Today; // will give the date for today
            string datePrint = date.ToShortDateString();
            string month = date.Month.ToString();
            string day = date.Day.ToString();

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            List<string[]> matches = new List<string[]>();

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {
                //for (int j = 4; j < colCount; j++)
                //{
                string comparison = xlRange.Cells[i, 4].Value2.ToString();
                string comparison2 = xlRange.Cells[i, 5].Value2.ToString();
                //new line
                //if (j == 1)
                //Console.Write("\r\n");
                int column = 4;
                if (xlRange.Cells[i, column] != null && xlRange.Cells[i, column].Value2 != null)

                    if (string.Equals(comparison, month) && column == 4)
                    {
                        if (string.Equals(comparison2, day))
                        {
                            string[] strArr = new string[colCount];
                            for (int y = 1; y <= colCount; y++)
                            {
                                strArr[y - 1] = xlRange.Cells[i, y].Value2.ToString();
                            }
                            matches.Add(strArr);
                        }
                    }
                //}
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return matches;
        }

        public static List<string[]> getWeekExcelFile()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\michael.gonzalez\source\repos\BirthdayPrg\Book1.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            DateTime date = DateTime.Today; // will give the date for today

            string month = date.Month.ToString();
            string day = date.Day.ToString();

            CultureInfo myCI = new CultureInfo("en-US");
            Calendar myCal = myCI.Calendar;

            CalendarWeekRule myCWR = myCI.DateTimeFormat.CalendarWeekRule;
            DayOfWeek myFirstDOW = myCI.DateTimeFormat.FirstDayOfWeek;

            int weekOfYear = myCal.GetWeekOfYear(DateTime.Now, myCWR, myFirstDOW);

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            List<string[]> matches = new List<string[]>();

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {
                //for (int j = 4; j < colCount; j++)
                //{
                string comparison = xlRange.Cells[i, 4].Value2.ToString();
                string comparison2 = xlRange.Cells[i, 5].Value2.ToString();
                //new line
                //if (j == 1)
                //Console.Write("\r\n");
                int column = 4;
                if (xlRange.Cells[i, column] != null && xlRange.Cells[i, column].Value2 != null)
                {
                    DateTime dateForWeekCheck = new DateTime(date.Year, Convert.ToInt32(comparison), Convert.ToInt32(comparison2));
                    int weekForDate = myCal.GetWeekOfYear(dateForWeekCheck, myCWR, myFirstDOW);
                    if (weekOfYear == weekForDate && column == 4)
                    {
                        string[] strArr = new string[colCount];
                        for (int y = 1; y <= colCount; y++)
                        {
                            strArr[y - 1] = xlRange.Cells[i, y].Value2.ToString();
                        }
                        matches.Add(strArr);
                    }
                }
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return matches;
        }
    }
}