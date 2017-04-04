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

        public static void SetExcelRow<T>(List<T> data)
        {
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
            try
            {

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        
        static public string[,] ListToArray(List<string> list)
        {

            string[,] elements = new string[list.Count(),1];

            for (int i = 0; i < list.Count(); i++)
            {
                elements[i, 0] = list.ElementAt(i);
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

