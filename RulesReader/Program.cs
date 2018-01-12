using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace RulesReader
{
    class Program
    {
        static string fileName = Path.Combine(Directory.GetCurrentDirectory(), "rules.xlsx");
        static string connectionString = String.Format("Server=.;Database=Noition;Trusted_Connection=True;");
        static int titleRow = 1;
        static string tableName = "ChangeRules";

        static void Main(string[] args)
        {
            SetConnectionString();

            Application application = new Application();
            Workbook workbook = application.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            ReadWorkbook(workbook);
        }

        [Conditional("DEBUG")]
        public static void SetConnectionString()
        {
            connectionString = String.Format("Initial Catalog={0};Data Source={1};User ID={2};Password={3}", "Noition", "192.168.1.120", "noition", "prune2017$");
        }

        static void ReadWorkbook(Workbook workbook)
        {
            Worksheet sheet = workbook.ActiveSheet;

            long fullRow = sheet.Rows.Count;
            long lastRow = sheet.Cells[fullRow, 1].End(XlDirection.xlUp).Row;

            string query = String.Format("INSERT INTO {0} ([OwnerID], [Before], [After]) VALUES ('5f39e5a7 - f4c7 - 49af - 9b55 - 796eb8c33d33', @F0, @F1)", tableName);

            char beforeNameColumn = 'H';
            char optionNameColumn = 'I';
            char afterNameColumn = 'J';

            for (int row = titleRow + 1; row <= lastRow; row++)
            {
                //Range cells = sheet.get_Range(String.Format("{0}{1}:{2}{1}", firstColumn, row, lastColumn), Type.Missing);

                string cell = beforeNameColumn + row.ToString();
                string beforeName = workbook.ActiveSheet.Range(cell).Value;

                cell = optionNameColumn + row.ToString();
                string optionName = workbook.ActiveSheet.Range(cell).Value;

                cell = afterNameColumn + row.ToString();
                string afterName = workbook.ActiveSheet.Range(cell).Value;

                beforeName = beforeName.Trim();
                string count = String.Empty;
                string pattern = @"^(.+)(\[\d+\])$";
                Regex rgx = new Regex(pattern);
                Match match = Regex.Match(beforeName, pattern);
                if (match.Success)
                {
                    beforeName = match.Groups[1].Value;
                    count = match.Groups[2].Value;
                }
                beforeName = beforeName.Trim();

                afterName = afterName.Trim();
                rgx = new Regex(pattern);
                match = Regex.Match(afterName, pattern);
                if (match.Success)
                {
                    afterName = match.Groups[1].Value;
                }
                afterName = afterName.Trim();

                if (!String.IsNullOrWhiteSpace(optionName))
                {
                    optionName = optionName.Trim();
                    string escapedOptionName = Regex.Escape(optionName);

                    pattern = "^(.+)(" + escapedOptionName + ")$";
                    rgx = new Regex(pattern);
                    match = Regex.Match(beforeName, pattern);
                    if (match.Success)
                    {
                        beforeName = match.Groups[1].Value;
                    }
                    beforeName = beforeName.Trim();

                    match = Regex.Match(afterName, pattern);
                    if (match.Success)
                    {
                        afterName = match.Groups[1].Value;
                    }
                    afterName = afterName.Trim();
                }

                List<SqlParameter> parameters = new List<SqlParameter>();
                parameters.Add(new SqlParameter("@F0", beforeName));
                parameters.Add(new SqlParameter("@F1", afterName));

                using (SqlConnection connection = new SqlConnection(connectionString))
                using (SqlCommand command = connection.CreateCommand())
                {
                    command.CommandText = query;
                    command.Parameters.AddRange(parameters.ToArray<SqlParameter>());
                    connection.Open();
                    command.ExecuteNonQuery();

                    Console.WriteLine(row);
                }

                Console.WriteLine(row + " : " + beforeName + " --> " + afterName);
            }
        }
    }
}
