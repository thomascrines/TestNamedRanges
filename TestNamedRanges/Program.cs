using System;
using OfficeOpenXml;
using System.IO;

namespace TestNamedRanges
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string testFileDirectory = "C:/Users/dsap01/source/repos/TestNamedRanges/TestNamedRanges/TestFiles/";
            string[] testFiles = { "POBE+2020+-+Blank+Return.xlsx", "POBE%2B2019%2B-%2BPublication%2Bworkbook.xlsx", "POBE+2019+-+Publication+tables.xlsx" };


            foreach (string testFile in testFiles)
            {
                string outputPath = $"{testFileDirectory}_{testFile}_summary.txt";

                File.Create(outputPath).Close();

                using (StreamWriter outputFile = new StreamWriter(outputPath, true))
                {
                    using (ExcelPackage package = new ExcelPackage(new FileInfo($"{testFileDirectory}{testFile}")))
                    {
                        outputFile.WriteLine($"File name: {testFile}\n");

                        ExcelWorksheets sheets = package.Workbook.Worksheets;

                        foreach (ExcelWorksheet sheet in sheets)
                        {
                            outputFile.WriteLine($"Sheet name: {sheet.Name}");

                            ExcelNamedRangeCollection namedRanges = sheet.Names;

                            outputFile.WriteLine("\nNamed Ranges");

                            if (namedRanges.Count == 0)
                            {
                                outputFile.WriteLine("  No named ranges in sheet");
                            }
                            else
                            {
                                foreach (ExcelNamedRange namedRange in namedRanges)
                                {
                                    outputFile.WriteLine($" Name: {namedRange.Name} | Start address: {namedRange.Start.Address} | End address: {namedRange.End.Address}");
                                }
                            }

                            outputFile.WriteLine("\nFormulas");

                            int formulaCount = 0;

                            foreach (ExcelRangeBase cell in sheet.Cells)
                            {
                                if (!string.IsNullOrEmpty(cell.Formula))
                                {
                                    var formula = cell.Formula;
                                    outputFile.WriteLine($" Cell address: {cell.Address} | Formula : {cell.Formula} | Value: {cell.Value}");

                                    formulaCount++;
                                }
                            }

                            if (formulaCount == 0)
                            {
                                outputFile.WriteLine("  No formulae in sheet");
                            }

                            outputFile.WriteLine("\n");
                        }
                    }
                }
            }
        }
    }
}
