using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using OfficeOpenXml;  // EPPlus Library

namespace AbbreviationWordAddin
{
    internal class AbbreviationManager
    {
        private static Dictionary<string, string> abbreviationDict = new Dictionary<string, string>();

        // Load abbreviations from embedded Excel file
        public static void LoadAbbreviations()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "AbbreviationWordAddin.Abbreviations.xlsx"; // Ensure the namespace matches your project

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    throw new Exception("Excel file not found in embedded resources.");
                }

                using (var package = new ExcelPackage(stream))

                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // First sheet
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)  // Skip header row
                    {
                        string phrase = worksheet.Cells[row, 1].Text.Trim();
                        string abbreviation = worksheet.Cells[row, 2].Text.Trim();

                        if (!string.IsNullOrEmpty(phrase) && !abbreviationDict.ContainsKey(phrase))
                        {
                            abbreviationDict[phrase] = abbreviation;
                        }
                    }
                }
            }
        }

        // Get abbreviation for a given phrase
        public static string GetAbbreviation(string phrase)
        {
            return abbreviationDict.ContainsKey(phrase) ? abbreviationDict[phrase] : phrase;
        }

        // Get all phrases for replacement
        public static List<string> GetAllPhrases()
        {
            return new List<string>(abbreviationDict.Keys);
        }
    }
}
