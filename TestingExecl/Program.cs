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
            DisplayInExcel(accounts.GetData(),students.GetData());

            Console.WriteLine("Success");
            Console.ReadKey();
        }

        static void DisplayInExcel(IEnumerable<Account> accounts, IEnumerable<Student> students)
        {
            var excelApp = new Excel.Application();

            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
            workSheet.Name = "Accounts";
            excelApp.Worksheets.Add();
            Excel._Worksheet workSheet1 = (Excel.Worksheet)excelApp.ActiveSheet;
            workSheet1.Name = "Student";
            // Header Name Loading
            HeaderRowLoading(workSheet,workSheet1);

            // incrementing the Row
            var row = 1;
            var row1 = 1;

            // Content Data Loading
            row = ContentLoading(accounts, workSheet, row);
            row1 = ContentStudentLoading(students, workSheet1, row1);
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

        private static int ContentLoading(IEnumerable<Account> accounts,Excel._Worksheet workSheet, int row)
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
                workSheet.Cells[row, "M"] = acct.BalanceAmount=acct.TotalSalary - acct.Expense;
                workSheet.Cells[row, "N"] = acct.Comments;

                    //background colour setting of account

                workSheet.Range["A1", "N1"].Interior.Color = XlRgbColor.rgbBurlyWood;
                workSheet.Range["K2:K11"].Interior.Color = XlRgbColor.rgbCadetBlue;
                workSheet.Cells[row, "J"].Font.Color = XlRgbColor.rgbGreen;
                workSheet.Cells[row, "L"].Font.Color=XlRgbColor.rgbBlue;
                workSheet.Range["I:M"].NumberFormat = "0.00";//Decimal point intialiaztion
              
                if (acct.BalanceAmount > 0)
                {
                    workSheet.Cells[row,"M"].Interior.Color = XlRgbColor.rgbDarkOliveGreen;
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
            return row1;
        }
    }
}
