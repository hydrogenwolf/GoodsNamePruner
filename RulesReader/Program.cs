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
        static string firstColumn = "A";
        static string lastColumn = "J";
        static int titleRow = 1;
        static string tableName = "Sales";

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

            Range titles = sheet.get_Range(String.Format("{0}{1}:{2}{1}", firstColumn, titleRow, lastColumn), Type.Missing);
            string[] fields = new string[titles.Count];
            int column = 0;
            foreach (Range title in titles)
            {
                if (title.Value2 == null)
                {
                    column += 1;
                    continue;
                }

                //fields[column] = title.Value2.ToString(); // 실제 저장된 값
                fields[column] = title.Text.ToString();     // 사용자에게 보여지는 값

                column += 1;
            }

            // 칼럼 제목이 상이한 자료를 위한 보정 작업
            if (fields[6].Equals("공급가합계")) fields[6] = "택배비";

            string c = "", v = "";
            for (int i = 0; i < fields.Length; i++)
            {
                if (!String.IsNullOrEmpty(fields[i]))
                {
                    if (c.Length > 0) c += ", ";
                    if (v.Length > 0) v += ", ";

                    c += "[" + fields[i] + "]";
                    v += "@F" + i;
                }
            }
            string query = String.Format("INSERT INTO {0} ({1}) VALUES ({2})", tableName, c, v);

            char goodsNameColumn = 'H';
            char optionNameColumn = 'I';

            for (int row = titleRow + 1; row <= lastRow; row++)
            {
                //Range cells = sheet.get_Range(String.Format("{0}{1}:{2}{1}", firstColumn, row, lastColumn), Type.Missing);

                string cell = goodsNameColumn + row.ToString();
                string goodsName = workbook.ActiveSheet.Range(cell).Value;

                goodsName = goodsName.Trim();
                string count = String.Empty;
                string pattern = @"^(.+)(\[\d+\])$";
                Regex rgx = new Regex(pattern);
                Match match = Regex.Match(goodsName, pattern);
                if (match.Success)
                {
                    goodsName = match.Groups[1].Value;
                    count = match.Groups[2].Value;
                }
                goodsName = goodsName.Trim();

                Console.WriteLine(goodsName);

                /*
                Range cells = sheet.get_Range(String.Format("{0}{1}:{2}{1}", firstColumn, row, lastColumn), Type.Missing);
                bool goingOn = false;
                int i = 0;
                List<SqlParameter> parameters = new List<SqlParameter>();
                foreach (Range cell in cells)
                {
                    if (String.IsNullOrEmpty(fields[i]))
                    {
                        i += 1;
                        continue;
                    }

                    if (cell.Value2 == null)
                    {
                        parameters.Add(new SqlParameter("@F" + i, ""));
                    }
                    else
                    {
                        parameters.Add(new SqlParameter("@F" + i, cell.Value2));    // 보여지는 값이 아닌 실제 저장된 값을 저장
                        goingOn = true;
                    }

                    i += 1;
                }

                if (!goingOn) continue;    // 마지막 Row를 지정하지 않는 대신 데이터가 없는 Row가 나오면 거기에서 중지

                using (SqlConnection connection = new SqlConnection(connectionString))
                using (SqlCommand command = connection.CreateCommand())
                {
                    command.CommandText = query;
                    command.Parameters.AddRange(parameters.ToArray<SqlParameter>());
                    connection.Open();
                    command.ExecuteNonQuery();

                    Console.WriteLine(row);
                }
                 */
            }
        }
    }
}
