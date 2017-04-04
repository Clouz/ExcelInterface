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
        /// <summary>
        /// Time to perform SetExcel operation
        /// </summary>
        public static Stopwatch ExcelSetTime { get; private set; } = new Stopwatch();

        /// <summary>
        /// Get Excel to List
        /// </summary>
        /// <param name="sheetNumber">Sheet number start from 1</param>
        /// <param name="path">Path of excel file</param>
        /// <param name="RowStart">Optional Row</param>
        /// <returns>List of Excel data</returns>
        public static IEnumerable<List<dynamic>> GetExcelRowEnumerator(int sheetNumber, string path, int RowStart = 1)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(path, ReadOnly: true);

            //seleziono la scheda
            Excel._Worksheet worksheet = (Excel._Worksheet)workbook.Sheets[sheetNumber];
            Excel.Range range = worksheet.UsedRange;

            // Array che conterrà il foglio excel completo
            object[,] valueArray = null;
            valueArray = range.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

            // Costruisco e yield ogni riga
            for (int rowIndex = RowStart; rowIndex <= valueArray.GetLength(0); rowIndex++)
            {
                List<dynamic> row = new List<dynamic>(valueArray.GetLength(1));

                for (int colIndex = 1; colIndex <= valueArray.GetLength(1); colIndex++)
                {
                    row.Add(valueArray[rowIndex, colIndex]);
                }
                yield return row;
            }
            workbook.Close();
        }


        /// <summary>
        /// List to Excel
        /// </summary>
        /// <typeparam name="T">Class representing data model</typeparam>
        /// <param name="data">List of data</param>
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
                head.Value2 = ListToArray(prop.Select(x => x.Name).ToList());

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
        
        /// <summary>
        /// List to multidimensional Array
        /// </summary>
        /// <param name="list">Rappresent Data</param>
        /// <returns></returns>
        static private string[,] ListToArray(List<string> list)
        {
            string[,] elements = new string[1, list.Count()];

            for (int i = 0; i < list.Count(); i++)
                elements[0, i] = list.ElementAt(i);

            return elements;
        }

        /// <summary>
        /// List<T> to multidimensional Array
        /// </summary>
        /// <typeparam name="T">Model Class</typeparam>
        /// <param name="list">Rappresent Data</param>
        /// <returns></returns>
        static private string[,] ListToArray<T>(List<T> list)
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

