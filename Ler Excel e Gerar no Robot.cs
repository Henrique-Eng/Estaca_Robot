using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using RobotOM;


namespace ExcelRobotConnector
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            IRobotApplication roApp = null;

            try
            {
                // --- Open Excel ---
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                workbook = excelApp.Workbooks.Open(@"C:\path\to\your\file.xlsx");
            }
        }
    }
}