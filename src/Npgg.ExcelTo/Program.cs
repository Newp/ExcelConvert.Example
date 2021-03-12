using System;
using System.Linq;
using System.IO;
using ExcelDataReader;
using System.Data;
using System.Text;
using System.Collections.Generic;

namespace Npgg.ExcelTo
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var filePath = @"c:/_git/DungeonBoard1/configurations/Configuration.xlsx";
            var outputPath = "c:/_git/DungeonBoard1/server/DungeonBoard1.Common/Configurations";

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {

                    var result = reader.AsDataSet();
                    char splitCharacter = ',';
                    foreach (DataTable table in result.Tables)
                    {
                        StringBuilder sb = new StringBuilder();

                        foreach (DataRow row in table.Rows)
                        {
                            var converted = row
                                .ItemArray
                                .Select(item => item.ToString())
                                .Select(item => item.Contains('"') || item.Contains(splitCharacter) ? $"\"{item}\"" : item);
                                

                            var line = string.Join(splitCharacter, converted);
                            sb.AppendLine(line);
                        }

                        var outputFilePath= Path.Combine(outputPath, table.TableName + ".csv");

                        File.WriteAllText(outputFilePath, sb.ToString(), new UTF8Encoding(false));
                    }
                }
            }
        }

    }
}
