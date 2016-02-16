using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using STNUpdater.Models;

namespace STNUpdater
{
    internal class FileSource
    {
        public FileSource()
        {
            FileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "resources", "oferte", "oferta_NOD_2015_12_08.xls");
            ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\"";
        }

        public List<Product> GetProducts()
        {
            var products = new List<Product>();
            try
            {
                using (var conn = new OleDbConnection(ConnectionString))
                {
                    conn.Open();

                    var dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    var sheetNames = new List<string>();
                    if (dt != null)
                    {
                        sheetNames.AddRange(dt.Rows.Cast<DataRow>()
                                .Select(row => row["TABLE_NAME"].ToString())
                                .Where(tableName => tableName.EndsWith("$") || tableName.EndsWith("$'")));
                    }
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
            catch (Exception ex)
            {
                var message = $"An error has occured: {ex.Message}";
                Console.WriteLine(message);
            }

            int warranty;
            products = products.Where(p => p.Warranty != null && int.TryParse(p.Warranty, out warranty)).ToList();

            return products;
        }

        private static IEnumerable<Product> MapDataSetToProducts(DataSet ds)
        {
            var products = ds.Tables[0].AsEnumerable()
                                       .Select(dataRow => new Product
                                       {
                                           Name = dataRow.Field<string>("F5"),
                                           Category = dataRow.Field<string>("F2"),
                                           Maker = dataRow.Field<string>("F4"),
                                           Model = dataRow.Field<string>("F4"),
                                           Code = dataRow.Field<string>("F5"),
                                           ShortDescription = dataRow.Field<string>("F6"),
                                           Warranty = dataRow.Field<string>("F10")
                                       }).ToList();
            return products;
        }

        public string FileName { get; set; }
        public string ConnectionString { get; set; }
    }
}