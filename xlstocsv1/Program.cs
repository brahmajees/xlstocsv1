using System;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        string sourceXlsxFilePath = @"d:\\bendem\NSDL\\nsdlsrc01.xlsx";
        string targetCsvFilePath = @"d:\\bendem\NSDL\\nsdlsrc01.csv";

        ConvertXlsxToCsv(sourceXlsxFilePath, targetCsvFilePath);

        Console.WriteLine("Conversion complete.");
    }

    static void ConvertXlsxToCsv(string sourceXlsxFilePath, string targetCsvFilePath)
    {
        using (var excelPackage = new ExcelPackage(new FileInfo(sourceXlsxFilePath)))
        {
            int DATA = 0;
            var worksheet = excelPackage.Workbook.Worksheets[DATA];
            int rows = worksheet.Dimension.Rows;
            int columns = worksheet.Dimension.Columns;

            using (var streamWriter = new StreamWriter(targetCsvFilePath))
            {
                // Write data rows
                for (int i = 1; i <= rows; i++)
                {
                    for (int j = 1; j <= columns; j++)
                    {
                        if (j > 1 &&  j <= 3)
                        {
                            streamWriter.Write(",");
                        }
                        var cellValue = worksheet.Cells[i, j].Value?.ToString() ?? "";
                        streamWriter.Write(cellValue);
                    }
                    streamWriter.WriteLine();
                }
            }
        }
    }
}

