using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Problem261
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application _Excel = new Excel.Application();
            Excel.Workbook wb = _Excel.Workbooks.Add();

            _Excel.DisplayAlerts = false;
            _Excel.Visible = true;

            List<string> words = GetText(args[0]);
            List<string> stopwords = GetStopwords("stop_words.txt");

            CreateExcelSpreadsheet(words, stopwords, wb);

            ExcelCleanUp(wb, _Excel);
        }

        public static void ExcelCleanUp(Excel.Workbook wb, Excel.Application _Excel)
        {
            if (wb != null)
            {
                try
                {
                    wb.Close();
                }
                catch { }

                Marshal.ReleaseComObject(wb);
                wb = null;
            }

            if (_Excel != null)
            {
                try
                {
                    _Excel.Quit();
                }
                catch { }

                Marshal.ReleaseComObject(_Excel);
                _Excel = null;
            }
        }

        public static void CreateExcelSpreadsheet(List<string> words, List<string> stopwords, Excel.Workbook wb)
        {
            Excel.Worksheet ws = wb.Worksheets.get_Item(1);

            int row = 1;
            ws.Cells[row, 1] = "words";
            ws.Cells[row, 2] = "stopwords";
            ws.Cells[row, 3] = "non_stopwords_with_duplicates";
            ws.Cells[row, 4] = "frequency_with_duplicates";
            ws.Cells[row, 5] = "non_stopwords_no_duplicates";
            ws.Cells[row, 6] = "frequency_no_duplicates";
            ws.Cells[row, 7] = "non_stopwords";
            ws.Cells[row, 8] = "stopwords";

            row = 2;

            foreach (string word in words)
            {
                if (word != "s")
                {
                    ws.Cells[row, 1] = word;
                    row++;
                }
            }

            row = 2;

            foreach (string stopword in stopwords)
            {
                ws.Cells[row, 2] = stopword;
                row++;
            }

            int total_rows = ws.UsedRange.Rows.Count;

            for (int i = 2; i < total_rows + 1; i++)
            {
                ((Excel.Range)ws.Cells[i, 3]).Value2 = "=IF(ISNUMBER(MATCH(A" + i.ToString() + ",$B$2:$B$120,0)),\"\",A" + i.ToString() + ")";
            }

            for (int i = 2; i < total_rows + 1; i++)
            {
                ((Excel.Range)ws.Cells[i, 4]).Value2 = "=IF(C" + i.ToString() + "=\"\", \"\",COUNTIF(C$2:C$999,C" + i.ToString() + "))";
            }

            Excel.Range copy_start = ws.Cells[1, 3];
            Excel.Range copy_end = ws.Cells[ws.UsedRange.Rows.Count, 4];
            Excel.Range copy = (Excel.Range)ws.get_Range(copy_start, copy_end);

            Excel.Range paste_start = ws.Cells[1, 5];
            Excel.Range paste_end = ws.Cells[ws.UsedRange.Rows.Count, 6];
            Excel.Range paste = (Excel.Range)ws.get_Range(paste_start, paste_end);

            copy.Copy();
            paste.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, true, false);
            paste.RemoveDuplicates(new object[]{1, 2}, Excel.XlYesNoGuess.xlYes);

            copy_start = ws.Cells[1, 5];
            copy_end = ws.Cells[ws.UsedRange.Rows.Count, 6];
            copy = (Excel.Range)ws.get_Range(copy_start, copy_end);

            paste_start = ws.Cells[1, 7];
            paste_end = ws.Cells[ws.UsedRange.Rows.Count, 8];
            paste = (Excel.Range)ws.get_Range(paste_start, paste_end);

            copy.Copy();
            paste.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, true, false);

            paste.AutoFilter(new object[]{1, 2}, "<>", Excel.XlAutoFilterOperator.xlAnd, Type.Missing, false);

            Excel.Range sort_range = ws.get_Range("H1");
            sort_range.Sort(sort_range, Excel.XlSortOrder.xlDescending, Type.Missing, Type.Missing, Excel.XlSortOrder.xlDescending, Type.Missing, Excel.XlSortOrder.xlDescending, Excel.XlYesNoGuess.xlYes, Type.Missing, Type.Missing, Excel.XlSortOrientation.xlSortColumns, Excel.XlSortMethod.xlPinYin, Excel.XlSortDataOption.xlSortNormal, Excel.XlSortDataOption.xlSortNormal, Excel.XlSortDataOption.xlSortNormal);

            ws.UsedRange.Columns.AutoFit();
            ws.UsedRange.Rows.AutoFit();

            wb.SaveAs(Environment.CurrentDirectory + @"\term_freq.xlsx");
            SheetCleanUp(ws);
        }

        public static void SheetCleanUp(Excel.Worksheet ws)
        {
            if (ws != null)
            {
                Marshal.ReleaseComObject(ws);
                ws = null;
            }
        }

        public static List<string> GetStopwords(string file)
        {
            return new StreamReader(file).ReadToEnd().ToLower().Replace(",\r\n\r\n", "").Split(',').ToList();
        }

        public static List<string> GetText(string file)
        {
            return Regex.Split(new StreamReader(file).ReadToEnd().ToLower().Replace("_", " "), "\\W+").ToList();
        }
    }
}
