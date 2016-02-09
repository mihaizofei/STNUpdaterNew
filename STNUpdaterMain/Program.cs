using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using MySql.Data.MySqlClient;
using STNUpdater.Models;

namespace STNUpdater
{
    class Program
    {
        static void Main()
        {
            var fileName = GetFileName();
            var excelConnectionString = GetExcelConnectionString(fileName);
            var dbConnectionString = GetDbConnectionString();

            var products = GetProductsFromFile(excelConnectionString);
            var categories = GetDbCategories(dbConnectionString);
            var makers = GetDbMakers(dbConnectionString);

            int warranty;
            products = products.Where(p => p.Warranty != null || int.TryParse(p.Warranty, out warranty)).ToList();
            
            Console.WriteLine(fileName);
            Console.WriteLine("All done!!!");
            Console.ReadLine();
        }

        private static List<Category> GetDbCategories(string dbConnectionString)
        {
            var results = new List<Category>();

            using (var conn = new MySqlConnection(dbConnectionString))
            using (var cmd = conn.CreateCommand())
            {    
                cmd.CommandText = "SELECT id_categorie, nume_categorie FROM cs_stonet.categorii;";
                conn.Open();
                var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var category = new Category
                    {
                        Id = Convert.ToInt32(reader["id_categorie"]),
                        Name = Convert.ToString(reader["nume_categorie"])
                    };
                    results.Add(category);
                }
                
            }

            return results;
        }

        private static List<Maker> GetDbMakers(string dbConnectionString)
        {
            var results = new List<Maker>();

            using (var conn = new MySqlConnection(dbConnectionString))
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT id_producator, nume_producator FROM cs_stonet.producatori;";
                conn.Open();
                var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var maker = new Maker
                    {
                        Id = Convert.ToInt32(reader["id_producator"]),
                        Name = Convert.ToString(reader["nume_producator"])
                    };
                    results.Add(maker);
                }

            }

            return results;
        }

        private static string GetDbConnectionString()
        {
            var connString = new MySqlConnectionStringBuilder
            {
                Server = "127.0.0.1",
                UserID = "root",
                Password = "fastweb321#",
                Database = "cs_stonet"
            };
            return "server = 127.0.0.1;uid=root;pwd=fastweb321#;database=cs_stonet;";
        }

        private static List<Product> GetProductsFromFile(string excelConnectionString)
        {
            var products = new List<Product>();
            try
            {
                using (var conn = new OleDbConnection(excelConnectionString))
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
    }
}
