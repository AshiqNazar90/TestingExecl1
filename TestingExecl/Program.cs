using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using TestingExecl.Model;
using Microsoft.Office.Interop.Excel;

namespace TestingExecl
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Account accounts = new Account();
            Student students = new Student();
            // Display the list in an Excel spreadsheet.
            DisplayInExcel(accounts.GetData(), students.GetData());

            Console.WriteLine("Success");
            Console.ReadKey();
        }

        static void DisplayInExcel(IEnumerable<Account> accounts, IEnumerable<Student> students)
        {
            Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            //Create a new excel book and sheet
            Excel.Workbook Workbooks;
            Excel.Worksheet WorkSheet;
            Excel.Worksheet WorkSheet1;
            object misValue = System.Reflection.Missing.Value;

            Workbooks = excelApp.Workbooks.Add();
            WorkSheet = (Excel.Worksheet)Workbooks.Worksheets.Add();
            WorkSheet.Name = "Accounts";
            WorkSheet1 = (Excel.Worksheet)Workbooks.Worksheets.Add();
            WorkSheet1.Name = "Student";
         
            // Header Name Loading
            HeaderRowLoading(WorkSheet, WorkSheet1);

            // incrementing the Row
            var row = 1;
            var row1 = 1;

            // Content Data Loading
            row = ContentLoading(accounts, WorkSheet, row);
            row1 = ContentStudentLoading(students, WorkSheet1, row1);

            // Footer Loading
            FooterRowLoading(accounts, WorkSheet, row);

            //SaveAndCloseExcel(excelApp, WorkBook);
            string location = @"C:\Users\SpeeHive\source\repos\TestingExecl1\ExcelReport\test1.xls";
            Workbooks.SaveAs(location);
           
            //Workbooks.Close(true);
            //excelApp.Quit();

        }

        private static void HeaderRowLoading(Excel._Worksheet workSheet, Excel._Worksheet workSheet1)
        {

            //work sheet of Account

            workSheet.Cells[1, "A"] = "Code";
            workSheet.Cells[1, "B"] = "Name";
            workSheet.Cells[1, "C"] = "Description";
            workSheet.Cells[1, "D"] = "CreatedBy";
            workSheet.Cells[1, "E"] = "ModifiedBy";
            workSheet.Cells[1, "F"] = "EventName";
            workSheet.Cells[1, "G"] = "Location";
            workSheet.Cells[1, "H"] = "Load";
            workSheet.Cells[1, "I"] = "BasicSalary";
            workSheet.Cells[1, "J"] = "HRA";
            workSheet.Cells[1, "K"] = "TotalSalary";
            workSheet.Cells[1, "L"] = "Expense";
            workSheet.Cells[1, "M"] = "BalanceAmount";
            workSheet.Cells[1, "N"] = "Comments";

            //worksheet of student
            workSheet1.Cells[1, "A"] = "Name";
            workSheet1.Cells[1, "B"] = "Age";
            workSheet1.Cells[1, "C"] = "Qualification";
            workSheet1.Cells[1, "D"] = "Height";
        }

        private static int ContentLoading(IEnumerable<Account> accounts, Excel._Worksheet workSheet, int row)
        {

            // data adding in Account
            foreach (var acct in accounts)
            {
                row++;
                workSheet.Cells[row, "A"] = acct.Code;
                workSheet.Cells[row, "B"] = acct.Name;
                workSheet.Cells[row, "C"] = acct.Description;
                workSheet.Cells[row, "D"] = acct.CreatedBy;
                workSheet.Cells[row, "E"] = acct.ModifiedBy;
                workSheet.Cells[row, "F"] = acct.EventName;
                workSheet.Cells[row, "G"] = acct.Location;
                workSheet.Cells[row, "H"] = acct.Load;
                workSheet.Cells[row, "I"] = acct.BasicSalary;
                workSheet.Cells[row, "J"] = acct.HRA;
                workSheet.Cells[row, "K"] = acct.TotalSalary = acct.BasicSalary + acct.HRA;
                workSheet.Cells[row, "L"] = acct.Expense;
                workSheet.Cells[row, "M"] = acct.BalanceAmount = acct.TotalSalary - acct.Expense;
                workSheet.Cells[row, "N"] = acct.Comments;

                //background colour setting of account

                workSheet.Range["A1", "N1"].Interior.Color = XlRgbColor.rgbBurlyWood;
                workSheet.Range["K2:K11"].Interior.Color = XlRgbColor.rgbCadetBlue;
                workSheet.Cells[row, "J"].Font.Color = XlRgbColor.rgbGreen;
                workSheet.Cells[row, "L"].Font.Color = XlRgbColor.rgbBlue;
                workSheet.Range["I:M"].NumberFormat = "0.00";//Decimal point intialiaztion

                if (acct.BalanceAmount > 0)
                {
                    workSheet.Cells[row, "M"].Interior.Color = XlRgbColor.rgbDarkOliveGreen;
                }
                else {
                    workSheet.Cells[row, "M"].Interior.Color = XlRgbColor.rgbIndianRed;
                }
                workSheet.Cells[row, "N"].HorizontalAlignment =
                 Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                workSheet.Columns["A:N"].AutoFit();
                //workSheet.Columns[2].AutoFit();
                //workSheet.Columns["N"].AutoFit();

            }
            return row;

        }

        private static int ContentStudentLoading(IEnumerable<Student> students, Excel._Worksheet workSheet1, int row1)
        {


            foreach (var stud in students)
            {
                row1++;
                workSheet1.Cells[row1, "A"] = stud.Name;
                workSheet1.Cells[row1, "B"] = stud.Age;
                workSheet1.Cells[row1, "C"] = stud.Qualification;
                workSheet1.Cells[row1, "D"] = stud.Height;
            }
            // style setting for student

            workSheet1.Range["A1", "D1"].Interior.Color = XlRgbColor.rgbDeepPink;
            workSheet1.Cells.Range["A2:A6"].Font.Color = XlRgbColor.rgbDarkGreen;
            workSheet1.Cells.Range["C2:C6"].Font.Color = XlRgbColor.rgbDarkRed;

            workSheet1.Columns["A:N"].AutoFit();
            return row1;
        }

        //Footer of Excell
        private static void FooterRowLoading(IEnumerable<Account> accounts, Excel._Worksheet workSheet, int row)
        {
            var excelApp = new Excel.Application();

            workSheet.Cells[row + 1, "J"] = "Total Amount";
            workSheet.Cells[row + 1, "J"].Interior.Color = XlRgbColor.rgbCoral;
            //workSheet.Columns[row+1,"J"].AutoFit();
            Excel.Range xlRng = workSheet.Range["K:K"];//range intialization for sum
            Excel.Range xlRng1 = workSheet.Range["L:L"];
            Excel.Range xlRng2 = workSheet.Range["M:M"];
            double sumResult = excelApp.WorksheetFunction.Sum(xlRng);// calling sum function 
            double sumResult1 = excelApp.WorksheetFunction.Sum(xlRng1);
            double sumResult2 = excelApp.WorksheetFunction.Sum(xlRng2);
            workSheet.Cells[row + 1, "K"] = sumResult;                  //assigning cell to display sum
            workSheet.Cells[row + 1, "K"].Font.Color = XlRgbColor.rgbDarkGreen;
            workSheet.Cells[row + 1, "L"] = sumResult1;
            workSheet.Cells[row + 1, "M"] = sumResult2;
            workSheet.Cells[row + 1, "M"].Font.Color = XlRgbColor.rgbDarkRed;

            //testing  calculation in console window

            Console.WriteLine("The sum of Column K is {0}.", sumResult);
            Console.WriteLine("The sum of Column L is {0}.", sumResult1);
            Console.WriteLine("The sum of Column M is {0}.", sumResult2);

        }

    }
}
