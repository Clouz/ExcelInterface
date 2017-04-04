using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace ExcelWs
{
    public static class Interface
    {
        public static int RowCount { get; private set; }
        public static int ColCount { get; private set; }

        public static Stopwatch ExcelGetTime { get; private set; } = new Stopwatch();
        public static Stopwatch ExcelSetTime { get; private set; } = new Stopwatch();

        public static IEnumerable<List<dynamic>> GetExcelRowEnumerator(int sheetNumber, string path)
        {
            //Start timer
            ExcelGetTime.Start();

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(path, ReadOnly: true);

            //seleziono la scheda
            Excel._Worksheet worksheet = (Excel._Worksheet)workbook.Sheets[sheetNumber];
            Excel.Range range = worksheet.UsedRange;

            RowCount = range.Rows.Count;
            ColCount = range.Columns.Count;

            // Array che conterrà il foglio excel completo
            object[,] valueArray = null;

            valueArray = range.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

            // build and yield each row at a time
            for (int rowIndex = 1; rowIndex <= valueArray.GetLength(0); rowIndex++)
            {
                List<dynamic> row = new List<dynamic>(valueArray.GetLength(1));
                // build a list of column values for the row
                for (int colIndex = 1; colIndex <= valueArray.GetLength(1); colIndex++)
                {
                    row.Add(valueArray[rowIndex, colIndex]);
                }
                yield return row;
            }

            workbook.Close();

            //stop timer
            ExcelGetTime.Stop();
        }

        public static void SetExcelRow<T>(List<T> data)
        {
            try
            {
                //start timer
                ExcelSetTime.Start();

                Excel.Application excelApp = new Excel.Application();
                excelApp.Workbooks.Add();
                Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

                var prop = typeof(T).GetProperties();
                int collumn = prop.Count();
                int row = data.Count()+1;
                
                //scrivo l'intestazione
                Excel.Range head = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, collumn]];
                var prova = ListToArray(prop.Select(x => x.Name).ToList());
                head.Value2 = prova;

                //scrivo il contenuto
                Excel.Range body = workSheet.Range[workSheet.Cells[2, 1], workSheet.Cells[row, collumn]];
                body.Value2 = ListToArray(data);

                workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, collumn]].Font.Bold = true;
                workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[row, collumn]].Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[row, collumn]].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                workSheet.Range[workSheet.Columns[1], workSheet.Columns[collumn]].AutoFit();

                //rende l'oggetto visibile
                excelApp.Visible = true;

                //stop timer
                ExcelSetTime.Stop();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        
        static public string[,] ListToArray(List<string> list)
        {

            string[,] elements = new string[1, list.Count()];

            for (int i = 0; i < list.Count(); i++)
            {
                elements[0, i] = list.ElementAt(i);
            }

            return elements;
        }

        static public string[,] ListToArray<T>(List<T> list)
        {
            var props = typeof(T).GetProperties();

            string[,] elements = new string[list.Count(), props.Count()];

            for (int i = 0; i < list.Count(); i++)
            {
                int ii = 0;
                foreach (var prop in props)
                {
                    elements[i, ii] = prop.GetValue(list.ElementAt(i)).ToString();
                    ii++;
                }
            }

            return elements;
        }
    }
}

