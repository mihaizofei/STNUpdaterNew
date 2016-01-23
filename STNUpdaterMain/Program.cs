using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;

namespace STNUpdater
{
    class Program
    {
        static void Main()
        {
            var fileName = GetFileName();
            var excelConnectionString = GetExcelConnectionString(fileName);
            var dbConnectionString = GetDbConnectionString();

            var products = GetProductsFromFile(fileName, excelConnectionString);
            int warranty;

            products = products.Where(p => p.Warranty != null || Int32.TryParse(p.Warranty, out warranty)).ToList();
            
            Console.WriteLine(fileName);
            Console.WriteLine("All done!!!");
            Console.ReadLine();
        }

        private static string GetDbConnectionString()
        {
            return String.Empty;
        }

        private static List<Product> GetProductsFromFile(string fileName, string excelConnectionString)
        {
            var products = new List<Product>();
            try
            {
                using (var conn = new OleDbConnection(excelConnectionString))
                {
                    conn.Open();

                    var dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    List<string> sheetNames = new List<string>();
                    sheetNames.AddRange(dt.Rows.Cast<DataRow>().Select(row => row["TABLE_NAME"].ToString()).Where(tableName => tableName.EndsWith("$") || tableName.EndsWith("$'")));
                    foreach (var sheetName in sheetNames)
                    {
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = "SELECT * FROM [" + sheetName + "] ";

                            var adapter = new OleDbDataAdapter(cmd);
                            var ds = new DataSet();
                            adapter.Fill(ds);
                            products.AddRange(MapDataSetToProducts(ds));
                        }
                    }
                }
            }
            catch (Exception)
            {
                var message = String.Format("An error has occured: {0}", ex.Message);
                Console.WriteLine(message);
            }
            
            return products;
        }

        private static string GetFileName()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "resources", "oferte", "oferta_NOD_2015_12_08.xls");
        }

        private static string GetExcelConnectionString(string fileName)
        {
            return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\"";
        }

        private static IEnumerable<Product> MapDataSetToProducts(DataSet ds)
        {
            var products = ds.Tables[0].AsEnumerable()
                                       .Select(dataRow => new Product
                                       {
                                           Range = dataRow.Field<string>("F1"),
                                           Category = dataRow.Field<string>("F2"),
                                           Subcategory = dataRow.Field<string>("F3"),
                                           Maker = dataRow.Field<string>("F4"),
                                           Code = dataRow.Field<string>("F5"),
                                           Description = dataRow.Field<string>("F6"),
                                           Price = dataRow.Field<string>("F7"),
                                           Currency = dataRow.Field<string>("F8"),
                                           Warranty = dataRow.Field<string>("F10")
                                       }).ToList();
            return products;
        }
    }
}
