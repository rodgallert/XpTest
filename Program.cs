using IronXL;
using System;
using System.Data;
using System.IO;
using System.Linq;

namespace XpTest
{
    internal class Program
    {
        const string _KEY = "id_fundo";

        static void Main(string[] args)
        {
            Console.Write("Path to first file: ");
            string path = Console.ReadLine();

            DataTable table1 = new DataTable();
            DataTable table2 = new DataTable();

            
            if (IsPathValid(path))
            {
                table1 = ReadFile(path);
            }

            Console.Write("Path to second file: ");
            path = Console.ReadLine();

            if (IsPathValid(path))
            {
                table2 = ReadFile(path);
            }

            if (table1 == null || table2 == null)
            {
                Console.WriteLine("You have a problem with your file");
                return;
            }
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
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
                return null;
            } finally { 
                workbook.Close();
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

        private static string CompareTables(DataTable table1, DataTable table2)
        {
            DataTable result = new DataTable();
            result.Columns.Add("CNPJ");
            result.Columns.Add("trib_cvm");
            result.Columns.Add("trib_anbima");
            result.Columns.Add("resultado");
        }
    }
}
