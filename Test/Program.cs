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

            for (int i = 0; i < 500; i++)
            {
                prova.Add(new Prova {
                    Nome = $"Claudio {i}",
                    Cognome = $"Mola {i+i}",
                    Telefono = i
                });
            }

            Interface.SetExcelRow<Prova>(prova);
            Console.WriteLine($"Creating time: {Interface.ExcelSetTime.Elapsed}");
            
            Console.ReadLine();
        }


        public class Prova
        {
            public string Nome { get; set; }
            public string Cognome { get; set; }
            public int Telefono { get; set; }
        }
    }
}
