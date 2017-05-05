using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelWs;
using System.Configuration;

namespace SeparaStringhe
{
    class Program
    {
        static int NumeroScheda { get; set; }
        static int NumeroColonna { get; set; }
        static int NumeroRiga { get; set; }

        static char Separatore { get; set; }
        static char CarattereRiempimento { get; set; }
        static string Pattern { get; set; }
        static string [] PatternDiviso { get; set; }

        static string PercorsoFile { get; set; }

        static List<Conversione> ColonnaExcel { get; set; } = new List<Conversione>();

        static int Main(string[] args)
        {
            try
            {
                PercorsoFile = args[0];
            }
            catch (Exception)
            {
                Console.WriteLine("Nessun File Selezionato");
                Console.ReadLine();
                return 1;
            }

            try
            {
                LeggiConfig();
                var excel = Interface.GetExcelRowEnumerator(NumeroScheda, PercorsoFile, NumeroRiga);

                foreach (var item in excel)
                    ColonnaExcel.Add(new Conversione {
                        NomeOriginale = (string)item[NumeroColonna - 1]
                    });

                foreach (var item in ColonnaExcel)
                {
                    item.NomeModificato = separa(item.NomeOriginale);
                }

                Interface.SetExcelRow<Conversione>(ColonnaExcel);

                Console.WriteLine("Operazione completata, premere invio per chiudere");
                Console.ReadLine();
                return 0;
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error: {e.ToString()}");
                return 1;
            }
        }

        public static string separa(string nome)
        {
            try
            {
                var lista = nome.Split(Separatore);
            
                string nuovaStringa = "";

                for (int i = 0; i < lista.Length; i++)
                {
                    int quantitaSpazi;
                    try
                    {
                        if (int.Parse(PatternDiviso[i]) == 0)
                        {
                            quantitaSpazi = 0;
                        }
                        else
                        {
                            quantitaSpazi = int.Parse(PatternDiviso[i]) - lista[i].Length;
                        }
                    }
                    catch (Exception)
                    {
                        quantitaSpazi = 0;
                    }

                    nuovaStringa = nuovaStringa + new string(CarattereRiempimento, quantitaSpazi) + lista[i] + Separatore;
                }
                nuovaStringa = nuovaStringa.Substring(0,nuovaStringa.Length -1);

                return nuovaStringa;
            }
            catch(Exception e)
            {
                return $"Stringa non corrispondente al pattern: {e.ToString()}";
            }

        }

        public static void LeggiConfig()
        {
            NumeroScheda = int.Parse(ConfigurationManager.AppSettings["NumeroScheda"]);
            NumeroColonna = int.Parse(ConfigurationManager.AppSettings["NumeroColonna"]);
            NumeroRiga = int.Parse(ConfigurationManager.AppSettings["NumeroRigaIniziale"]);

            Separatore = char.Parse(ConfigurationManager.AppSettings["CarattereSeparatore"]);
            CarattereRiempimento = char.Parse(ConfigurationManager.AppSettings["CarattereRiempimento"]);
            Pattern = ConfigurationManager.AppSettings["Pattern"];

            PatternDiviso = Pattern.Split(Separatore);

            Console.WriteLine("Config");
            Console.WriteLine($"\tNumeroRighe:\t\t{NumeroScheda}");
            Console.WriteLine($"\tNumeroColonna:\t\t{NumeroColonna}");
            Console.WriteLine($"\tNumeroRiga:\t\t{NumeroRiga}");
            Console.WriteLine($"\tSeparatore:\t\t{Separatore}");
            Console.WriteLine($"\tCarattereRiempimento:\t{CarattereRiempimento}");
            Console.WriteLine($"\tPattern:\t\t{Pattern}");
            Console.WriteLine($"\tPercorsoFile:\t\t{PercorsoFile}");
        }

    }

    public class Conversione
    {
        public string NomeOriginale { get; set; }
        public string NomeModificato { get; set; }
    }

    
}
