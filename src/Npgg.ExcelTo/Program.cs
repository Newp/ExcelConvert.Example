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

            Console.WriteLine(string.Join('\t', args));
            var filePath = args[0]; 
            var outputPath = args[1];

            var exportExtension = args[2];

            var splitCharacter = exportExtension.ToLower() switch
            {
                "csv" => ',',
                "tsv" => '\t',
                _=> throw new Exception("invalid args[2] format")
            };

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {

                    var result = reader.AsDataSet();
                    foreach (DataTable table in result.Tables)
                    {
                        StringBuilder sb = new StringBuilder();

                        foreach (DataRow row in table.Rows)
                        {
                            var converted = row
                                .ItemArray
                                .Select(item => item.ToString());
                                
                                

                            var line = string.Join(splitCharacter, converted);
                            sb.AppendLine(line);
                        }

                        var outputFilePath= Path.Combine(outputPath, $"{table.TableName}.{exportExtension}");

                        File.WriteAllText(outputFilePath, sb.ToString(), new UTF8Encoding(false));
                    }
                }
            }
        }

    }
}
