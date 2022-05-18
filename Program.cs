using ClosedXML.Excel;
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

        static void Main()
        {
            
            DataTable table1 = new DataTable();
            DataTable table2 = new DataTable();

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
            }

            if (table1 == null || table2 == null)
            {
                Console.WriteLine("You have a problem with your file");
                return;
            }
            CompareTables(table1, table2);
        }

        /// <summary>
        /// Reads a XLS file and returns it as a DataTable object
        /// </summary>
        /// <param name="path">A string ponting to a file</param>
        /// <returns>The file converted as a DataTable</returns>
        private static DataTable ReadFile(string path)
        {
            try
            {
                // loads the file
                WorkBook workbook = new WorkBook(path);

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

        /// <summary>
        /// Checks if a provided file path is not null or empty
        /// </summary>
        /// <param name="path">A string pointing to a file</param>
        /// <returns>False if the provided string is either null or empty</returns>
        private static bool IsPathValid(string path)
        {
            if (path == string.Empty || path == null)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Checks if the input string has any special characters and remove them
        /// </summary>
        /// <param name="input">A string</param>
        /// <returns>A string with only numbers</returns>
        private static string RemoveSpecialChars(string input)
        {
            return new string(input.Where(x => Char.IsDigit(x)).ToArray());
        }

        /// <summary>
        /// Compares two provided spreadsheets and creates a third sheet with the resulting rows
        /// </summary>
        /// <param name="table1">A DataTable from a XLSX file</param>
        /// <param name="table2">A DataTable from a XLSX file</param>
        private static void CompareTables(DataTable table1, DataTable table2)
        {
            DataTable resultTable = new DataTable
            {
                TableName = "Result table",
            };

            // Creates a new, empty DataTable
            resultTable.Columns.Clear();
            resultTable.Rows.Clear();

            // Adds the columns with required results
            resultTable.Columns.Add(new DataColumn("CNPJ", typeof(string)));
            resultTable.Columns.Add(new DataColumn(_TRIBCVM, typeof(string)));
            resultTable.Columns.Add(new DataColumn(_TRIBANBIMA, typeof(string)));
            resultTable.Columns.Add(new DataColumn("Resultado", typeof(string)));
            
            // Selects a tuple from two spreadsheets, with its CNPJ, its tributes in either bases, and a default value that those tributes are not the same
            var result = from row in table1.AsEnumerable()
                         join row2 in table2.AsEnumerable()
                         on RemoveSpecialChars(row[_KEY].ToString()) equals RemoveSpecialChars(row2[_KEY].ToString())
                         where row[_KEY] != null
                         where row2[_KEY] != null
                         select new Tupla(row[_KEY].ToString(), row[_TRIBCVM].ToString(), row2[_TRIBANBIMA].ToString(), "Diferente");

            foreach(Tupla tupla in result)
            {
                resultTable.Rows.Add(RemoveSpecialChars(tupla.Cnpj), tupla.TributacaoCvm, tupla.TributacaoAnbima, CompareTributes(tupla));
            }

            resultTable.AcceptChanges();

            SaveFile(resultTable);
        }

        /// <summary>
        /// Compares two tributes from a tuple found in the two tables provided, and returns if they are equal or not
        /// </summary>
        /// <param name="t">A tuple from a joined query of two datasheets</param>
        /// <returns>Igual if the tributes are the same, keeps default value Diferente if they are not</returns>
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

        /// <summary>
        /// Saves a DataTable from the join and processing from other 2 spreadsheets
        /// </summary>
        /// <param name="table">A DataTable to be saved as a XLSX file</param>
        private static void SaveFile(DataTable table)
        {
            Console.Write("Please provide destination path, without file name: ");
            string path = Console.ReadLine();
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(table);
            wb.SaveAs(path + "\\result.xlsx");
        }
    }
}