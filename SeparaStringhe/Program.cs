﻿using System;
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
        static string CarattereRiempimento { get; set; }
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


            Console.ReadLine();
            return 0;

        }

        public static string separa(string nome)
        {
            var lista = nome.Split(Separatore);

            int indice = 0;

            string nuovaStringa = "";

            foreach (var item in lista)
            {
                nuovaStringa = nuovaStringa + item;

                for (int i = item.Length; i < PatternDiviso[indice].Length; i++)
                {
                    nuovaStringa = nuovaStringa + CarattereRiempimento;
                }
                indice++;

                nuovaStringa = nuovaStringa + Separatore;
            }

            return "";
        }

        public static void LeggiConfig()
        {
            NumeroScheda = int.Parse(ConfigurationManager.AppSettings["NumeroScheda"]);
            NumeroColonna = int.Parse(ConfigurationManager.AppSettings["NumeroColonna"]);
            NumeroRiga = int.Parse(ConfigurationManager.AppSettings["NumeroRiga"]);

            Separatore = char.Parse(ConfigurationManager.AppSettings["Separatore"]);
            CarattereRiempimento = ConfigurationManager.AppSettings["CarattereRiempimento"];
            Pattern = ConfigurationManager.AppSettings["NumeroRiga"];

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