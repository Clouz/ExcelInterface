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

            List<Prova> prova = new List<Prova>();

            for (int i = 0; i < x.Count(); i++)
            {
                prova.Add(new Prova {
                    col1 = x.ElementAt(i).ElementAt(0),
                    col2 = x.ElementAt(i).ElementAt(1)
                });
            }

            Interface.SetExcelRow(prova);


            Console.ReadLine();
        }


        public class Prova
        {
            public string col1 { get; set; }
            public string col2 { get; set; }
        }
    }
}
