using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace POD_Extractor
{
    class ExcelOperations
    {

        public static ExcelWorksheet OpenExcelObject(String filepath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage xlPackage = new ExcelPackage(new System.IO.FileInfo(filepath));
            ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[0];
            return worksheet;
        }

        public static string SearchForPatterns(ExcelWorksheet worksheet, List<Tuple<string, int, int>> patterns)
        {

            // Define search area
            var startRow = worksheet.Dimension.Start.Row;
            var startCol = worksheet.Dimension.Start.Column;
            var endRow = worksheet.Dimension.End.Row;
            var endCol = worksheet.Dimension.End.Column;

            // Iterate all cells in range
            string matchedVal = default;
            bool match = false;

            for (int iRow = startRow; iRow <= endRow; iRow++)
            {
                for (int iCol = startCol; iCol <= endCol; iCol++)
                {
                    // Regex test cell value for CreditorName Pattern
                    var cellVal = worksheet.Cells[iRow, iCol].Value;
                    if (cellVal != null)
                    {
                        foreach (var pattern in patterns)
                        {
                            match = Regex.IsMatch(cellVal.ToString(), pattern.Item1, RegexOptions.IgnoreCase);
                            if (match == true)
                            {
                                //break from column loop
                                try
                                {
                                    matchedVal = worksheet.Cells[iRow + pattern.Item2, iCol + pattern.Item3].Value.ToString();

                                }
                                catch
                                {
                                    //matchedVal = cellVal.ToString();
                                    matchedVal = SpreadSearch(worksheet, iRow, iCol);
                                }
                                return matchedVal;
                            }
                        }
                    }
                }
            }
            return matchedVal;

        }

        public static void SegregatePODForm(ExcelWorksheet worksheet)
        {
            // Header

            // Creditor Name

            // Amount

            // Account Number

            // Details

        }

        public static String IrregularSpreadsheetSearch(ExcelWorksheet worksheet, String tablePattern, String itemPattern)
        {
            // Define search area
            var startRow = worksheet.Dimension.Start.Row;
            var startCol = worksheet.Dimension.Start.Column;
            var endRow = worksheet.Dimension.End.Row;
            var endCol = worksheet.Dimension.End.Column;

            // Iterate all cells in range
            string matchedRow = default;
            bool match = false;

            for (int iRow = startRow; iRow <= endRow; iRow++)
            {
                for (int iCol = startCol; iCol <= endCol; iCol++)
                {
                    var cellVal = worksheet.Cells[iRow, iCol].Value;
                    if(cellVal != null)
                    {
                        match = Regex.IsMatch(cellVal.ToString(), tablePattern, RegexOptions.IgnoreCase);
                        if (match == true)
                        {
                            for (int _iCol_ = iCol+1; _iCol_ <= endCol; _iCol_++)
                            {
                                var _cellVal_ = worksheet.Cells[iRow, _iCol_].Value;
                                if(_cellVal_ != null){
                                    match = Regex.IsMatch(_cellVal_.ToString(), itemPattern, RegexOptions.IgnoreCase);
                                    if (match == true)
                                    {
                                        return _cellVal_.ToString();
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return null;
        }

        private static string SpreadSearch(ExcelWorksheet worksheet, int iRow, int iCol)
        {

            // Define search area
            var endCol = worksheet.Dimension.End.Column;

            // Iterate all cells in range
            string matchedVal = default;

            for (int i = 1; i <= endCol; i++)
            {
                var cellVal = worksheet.Cells[iRow, iCol + i].Value;
                if (cellVal != null)
                {
                    matchedVal = cellVal.ToString();
                    break;
                }
                else
                {
                    continue;
                }
            }
            return matchedVal;
        }

    }
}
