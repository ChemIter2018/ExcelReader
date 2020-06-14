using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;

namespace ExcelReader
{
    class Program
    {
        static void Main(string[] args)
        {
            var excelPath = System.Environment.CurrentDirectory + "\\Test.xlsx";
            Console.WriteLine(System.Environment.CurrentDirectory.ToString());

            using (var stream = File.Open(excelPath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    using (var result = reader.AsDataSet())
                    {
                        for (int i = 0; i < result.Tables[0].Columns.Count; i++)
                        {
                            for (int j = 0; j < result.Tables[0].Rows.Count; j++)
                            {
                                Console.WriteLine(result.Tables[0].Rows[j][i].ToString());                               
                            }
                        }

                        for (int i = 0; i < result.Tables[1].Columns.Count; i++)
                        {
                            for (int j = 0; j < result.Tables[1].Rows.Count; j++)
                            {
                                Console.WriteLine(result.Tables[1].Rows[j][i].ToString());
                            }
                        }
                    }
                }
            }
        }
    }
}
