using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;

Excel.Application excelApp = new Excel.Application();
Excel.Workbook workbook = excelApp.Workbooks.Open("‪C:\\Users\\ashok.c\\Downloads\\Group by's (1).xlsx");
Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[0]; // Indexing starts from 1

Dictionary<string, int> duplicateValues = new Dictionary<string, int>();

int rowCount = worksheet.UsedRange.Rows.Count;
int colCount = worksheet.UsedRange.Columns.Count;
for (int row = 1; row <= rowCount; row++)
{
    for (int col = 1; col <= colCount; col++)
    {
        Excel.Range cell = (Excel.Range)worksheet.Cells[row, col];
        string? cellValue = cell.Value?.ToString();

        if (!string.IsNullOrEmpty(cellValue))
        {
            if (duplicateValues.ContainsKey(cellValue))
            {
                duplicateValues[cellValue]++;
            }
            else
            {
                duplicateValues.Add(cellValue, 1);
            }
        }
    }
}

foreach (var kvp in duplicateValues)
{
    if (kvp.Value > 1)
    {
        Console.WriteLine($"Value: {kvp.Key}, Count: {kvp.Value}");
    }
}

// Cleanup
workbook.Close(false);
excelApp.Quit();

