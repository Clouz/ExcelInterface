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
            //var x = Interface.GetExcelRowEnumerator(1, args[0]);

            List<Prova> prova = new List<Prova>();

            for (int i = 0; i < 10; i++)
            {
                prova.Add(new Prova {
                    col1 = i,
                    col2 = i*i
                });
            }

            var x = Interface.ListToArray(prova);

            Interface.SetExcelRow<Prova>(prova);

            Console.ReadLine();
        }


        public class Prova
        {
            public int col1 { get; set; }
            public int col2 { get; set; }
        }
    }
}
