


using Microsoft.Extensions.FileProviders;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
class Program
{
    static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        string pathSheet1 = @"D:\\Excel comparision\PreviousData.xlsx"; // Replace with the actual path of your Excel files
        string pathSheet2 = @"D:\\Excel comparision\CurrentData.xlsx";
        // Load the Excel files
        FileInfo fileInfo1 = new FileInfo(pathSheet1);
        FileInfo fileInfo2 = new FileInfo(pathSheet2);//load

        using (var package1 = new ExcelPackage(fileInfo1))
        using (var package2 = new ExcelPackage(fileInfo2))
        {
            // Get the first worksheet from each Excel file
            var sheet1 = package1.Workbook.Worksheets[0];
            var sheet2 = package2.Workbook.Worksheets[0];

            // Compare the data and write the result to Sheet2 directly
            int resultRow = 2;

            for (int row = 2; row <= sheet2.Dimension.End.Row; row++)
            {
                var app = sheet2.Cells[row, 1].Value?.ToString();
                var userId = sheet2.Cells[row, 2].Value?.ToString();

                var matchingRowSheet1 = sheet1.Cells
                    .Where(cell => cell.Start.Row > 1)
                    .FirstOrDefault(cell => cell.Text == app && cell.Offset(0, 1).Text == userId);

                if (matchingRowSheet1 != null)
                {
                    // Existing User
                    sheet2.Cells[row, 5].Value = "Existing User";
                }
                else
                {
                    // New User
                    sheet2.Cells[row, 5].Value = "New User";
                }
            }

            // Add Deleted Users
            for (int row = 2; row <= sheet1.Dimension.End.Row; row++)
            {
                var app = sheet1.Cells[row, 1].Value?.ToString();
                var userId = sheet1.Cells[row, 2].Value?.ToString();

                var matchingRowSheet2 = sheet2.Cells
                    .Where(cell => cell.Start.Row > 1)
                    .FirstOrDefault(cell => cell.Text == app && cell.Offset(0, 1).Text == userId);

                resultRow = sheet2.Dimension.End.Row + 1;

                if (matchingRowSheet2 == null)
                {
                    // Deleted User


                    sheet2.Cells[resultRow, 1].Value = sheet1.Cells[row, 1].Text;
                    sheet2.Cells[resultRow, 2].Value = sheet1.Cells[row, 2].Text;
                    sheet2.Cells[resultRow, 3].Value = sheet1.Cells[row, 3].Text;
                    sheet2.Cells[resultRow, 4].Value = sheet1.Cells[row, 4].Text;
                    sheet2.Cells[resultRow, 5].Value = "Deleted User";
                    resultRow++;

                }




            }
            int count = 0;
            for (int row = 2; row <= sheet2.Dimension.End.Row; row++)
            {
                
                var userId = sheet2.Cells[row, 2].Value?.ToString();
                if (!ContainsDigits(userId))
                {
                    // Delete the row from sheet2
                    sheet2.DeleteRow(row, 1);
                    count++;
                    row--;
                    
                }
                
            }
            Console.WriteLine($"Number of userid's deleted are {count}");


            // Save the result to the same Excel file (Sheet2)
            package2.Save();
        }
        



    }

    
    private static bool ContainsDigits(string input)
    {
        return new Regex(@"\d").IsMatch(input);
    }







}
