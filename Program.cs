using IronXL;
using System;
using System.Data;
using System.IO;
using System.Linq;

namespace XpTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(RemoveSpecialChars("12...581...135/./0001-.-77"));
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

                Console.WriteLine(sheet.ColumnCount);

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
    }
}
