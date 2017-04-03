using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelWs;


namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            var x = Interface.GetExcelRowEnumerator(1, args[0]);

            string[,] intestazione = {{ "Azione","ID", "Modifica"}};
            string[,] contenuto = new string[Interface.rowCount, Interface.colCount];

            int i = 0;
            foreach(var xx in x)
            {
                int ii = 0;
                foreach (var item in xx)
                {
                    Console.Write($"{item}\t");
                    contenuto[i, ii] = item;
                    ii++;
                }
                i++;
                Console.WriteLine();
            }

            //Interface.SetExcelRow();

            Console.ReadLine();
        }
    }
}
