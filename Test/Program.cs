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

            List<Prova> lista = new List<Prova>();

            foreach (var item in x)
            {
                lista.Add(new Prova() {
                    Nome = item[1],
                    Telefono = (int)item[0]
                });
            }

            Interface.SetExcelRow<Prova>(lista);
            Console.WriteLine($"Creating time: {Interface.ExcelSetTime.Elapsed}");

            Console.ReadLine();
        }


        public class Prova
        {
            public string Nome { get; set; }
            public int Telefono { get; set; }
        }
    }
}
