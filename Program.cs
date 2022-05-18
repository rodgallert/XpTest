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

        static void Main(string[] args)
        {
            Console.Write("Path to first file: ");
            string path = Console.ReadLine();

            DataTable table1 = new DataTable();
            DataTable table2 = new DataTable();


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
            }

            if (table1 == null || table2 == null)
            {
                Console.WriteLine("You have a problem with your file");
                return;
            }

            CompareTables(table1, table2);
        }

        private static DataTable ReadFile(string path)
        {
            WorkBook workbook = new WorkBook();
            try
            {
                // loads the file
                WorkBook workBook = new WorkBook(path);

                // gets the defaul (first) sheet from the file
                WorkSheet sheet = workBook.DefaultWorkSheet;

                return sheet.ToDataTable(true);
            }
            catch (FileNotFoundException ex)
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
            DataTable table3 = new DataTable();
            table3.Columns.Add("CNPJ", typeof(string));
            table3.Columns.Add("Tributacao_cvm", typeof(string));
            table3.Columns.Add("Tributacao_anbima", typeof(string));
            table3.Columns.Add("Resultado", typeof(string));

            var result = from row in table1.AsEnumerable()
                         join row2 in table2.AsEnumerable()
                         on RemoveSpecialChars(row[_KEY].ToString()) equals RemoveSpecialChars(row2[_KEY].ToString())
                         where row[_KEY] != null
                         where row2[_KEY] != null
                         select row;

            foreach (DataRow row in result)
            {
                table3.Rows.Add(row[_KEY], row[_TRIBCVM], row[_TRIBANBIMA]);
            }
        }
    }

}