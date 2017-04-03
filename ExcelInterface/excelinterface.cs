using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelWs
{
    public class Interface
    {
        public static int rowCount { get; private set; }
        public static int colCount { get; private set; }

        public static IEnumerable<List<dynamic>> GetExcelRowEnumerator(int sheetNumber, string path)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(path, ReadOnly: true);

            //seleziono la scheda
            Excel._Worksheet worksheet = (Excel._Worksheet)workbook.Sheets[sheetNumber];
            Excel.Range range = worksheet.UsedRange;

            rowCount = range.Rows.Count;
            colCount = range.Columns.Count;

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
        }

        public static void SetExcelRow(string[,] heading, string[,] value, int[] columnsInt = null)
        {
            try
            {
                Excel.Application excelApp = new Excel.Application();

                excelApp.Workbooks.Add();
                Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

                int colonne = heading.Length;
                int elementiTotali = value.Length;
                int righe = elementiTotali / colonne + 1;

                //scrivo l'intestazione
                Excel.Range testa = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, colonne]];
                testa.Value2 = heading;

                //scrivo il contenuto
                Excel.Range corpo = workSheet.Range[workSheet.Cells[2, 1], workSheet.Cells[righe, colonne]];
                corpo.Value2 = value;

                workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, colonne]].Font.Bold = true;
                workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[righe, colonne]].Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[righe, colonne]].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                workSheet.Range[workSheet.Columns[1], workSheet.Columns[colonne]].AutoFit();

                if (columnsInt != null)
                {
                    foreach (var item in columnsInt)
                    {
                        workSheet.Columns[item].NumberFormat = "0";
                    }
                }

                //rende l'oggetto visibile
                excelApp.Visible = true;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public static void SetExcelRow(List<object> data)
        {
            try
            {
                Excel.Application excelApp = new Excel.Application();
                excelApp.Workbooks.Add();
                Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

                var dataType = data.GetType();
                var info = dataType.GetProperties();
                int collumn = info.Count();
                int row = data.Count()+1;
                
                //scrivo l'intestazione
                Excel.Range testa = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, collumn]];
                testa.Value2 = ListToArray(info.Select(a => a.Name).ToList());

                //scrivo il contenuto
                Excel.Range corpo = workSheet.Range[workSheet.Cells[2, 1], workSheet.Cells[row, collumn]];
                corpo.Value2 = ListToArray(data.ToList());

                workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, collumn]].Font.Bold = true;
                workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[row, collumn]].Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[row, collumn]].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                workSheet.Range[workSheet.Columns[1], workSheet.Columns[collumn]].AutoFit();

                //rende l'oggetto visibile
                excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        static private string[,] ListToArray(List<string> list)
        {
            string[,] elements = new string[list.Count(),0];

            for (int i = 0; i < list.Count(); i++)
            {
                elements[i, 0] = list[i];
            }

            return elements;
        }

        static private string[,] ListToArray(List<object> list)
        {
            var colonne = list.GetType().GetProperties().Select(x => x.Name);

            string[,] elements = new string[list.Count(), colonne.Count()];

            for (int i = 0; i < list.Count(); i++)
            {
                int ii = 0;
                foreach (var item in colonne)
                {
                    elements[i, ii] = list.ElementAt(i).GetType().GetProperty(item).ToString();
                    ii++;
                }
            }

            return elements;
        }
    }
}

