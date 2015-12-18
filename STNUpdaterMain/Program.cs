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
            var fileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "resources", "oferte", "oferta_NOD_2015_12_08.xlsx");
            var products = new List<Product>();

            var connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\""; ;
            using (var conn = new OleDbConnection(connectionString))
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


            Console.WriteLine(fileName);
            Console.ReadLine();
        }

        private static IEnumerable<Product> MapDataSetToProducts(DataSet ds)
        {
            var products = ds.Tables[0].AsEnumerable().Select(dataRow => new Product
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
