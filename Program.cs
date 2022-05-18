using IronXL;
using System;
using System.Data;
using System.IO;
using System.Linq;

namespace XpTest
{
    internal class Program
    {
        private const string _KEY = "id_fundo";
        private const string _TRIBCVM = "TRIB_LPRAZO";
        private const string _TRIBANBIMA = "tributacao_alvo";
        private const string _PATHCVM = "d:\\repos\\xptest\\docs\\fundos_cvm.xlsx";
        private const string _PATHANBIMA = "d:\\repos\\xptest\\docs\\fundos_anbima.xlsx";


        static void Main()
        {
            
            DataTable table1 = new DataTable();
            DataTable table2 = new DataTable();

            /*
            Console.Write("Path to first file: ");
            string path = Console.ReadLine();

            if (IsPathValid(path))
            {
                Console.WriteLine("Loading file 1...");
                table1 = ReadFile(path);
            }

            Console.Write("Path to second file: ");
            path = Console.ReadLine();

            if (IsPathValid(path))
            {
                Console.WriteLine("Loading file 2...");
                table2 = ReadFile(path);
            }*/

            table1 = ReadFile(_PATHCVM);
            table2 = ReadFile(_PATHANBIMA);

            if (table1 == null || table2 == null)
            {
                Console.WriteLine("You have a problem with your file");
                return;
            }

            if (table1 != null && table2 != null)
            {
                CompareTables(table1, table2);
            } else
            {
                Console.WriteLine("Please provide a valid worksheet");
            }
        }

        private static DataTable ReadFile(string path)
        {
            WorkBook workbook;
            try
            {
                // loads the file
                workbook = new WorkBook(path);

                // gets the default (first) sheet from the file
                WorkSheet sheet = workbook.DefaultWorkSheet;

                return sheet.ToDataTable(true);
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("Could not find the specified file " + path);
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }

        private static bool IsPathValid(string path)
        {
            if (path == string.Empty || path == null)
            {
                return false;
            }
            return true;
        }

        private static string RemoveSpecialChars(string input)
        {
            // removes everything that is not a number from the input string
            return new string(input.Where(x => Char.IsDigit(x)).ToArray());
        }

        private static void CompareTables(DataTable table1, DataTable table2)
        {
            DataTable table3 = new DataTable
            {
                TableName = "Result table"
            };
            table3.Columns.Clear();
            table3.Rows.Clear();


            table3.Columns.Add(new DataColumn("CNPJ", typeof(string)));
            table3.Columns.Add(new DataColumn(_TRIBCVM, typeof(string)));
            table3.Columns.Add(new DataColumn(_TRIBANBIMA, typeof(string)));
            table3.Columns.Add(new DataColumn("Resultado", typeof(string)));

            var result = from row in table1.AsEnumerable()
                         join row2 in table2.AsEnumerable()
                         on RemoveSpecialChars(row[_KEY].ToString()) equals RemoveSpecialChars(row2[_KEY].ToString())
                         where row[_KEY] != null
                         where row2[_KEY] != null
                         select new Tupla(row[_KEY].ToString(), row[_TRIBCVM].ToString(), row2[_TRIBANBIMA].ToString(), "Diferente");

            foreach(Tupla tupla in result)
            {

                table3.Rows.Add(tupla.Cnpj, tupla.TributacaoCvm, tupla.TributacaoAnbima, CompareTributes(tupla));
            }


        }

        private static string CompareTributes(Tupla t)
        {
            if (t.TributacaoCvm.Equals("S"))
            {
                if (t.TributacaoAnbima.Equals("Longo Prazo"))
                {
                    return "Igual";
                }
            }

            if (t.TributacaoCvm.Equals("N/A"))
            {
                if (t.TributacaoAnbima.Equals("Não Aplicável"))
                {
                    return "Igual";
                }
            }
            if (t.TributacaoCvm.Equals(string.Empty))
            {
                if (t.TributacaoAnbima.Equals("Indefinido"))
                {
                    return "Igual";
                }
            }
            return t.Resultado;
        }
    }
}