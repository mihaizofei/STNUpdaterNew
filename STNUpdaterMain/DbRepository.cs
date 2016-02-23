using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using MySql.Data.MySqlClient;
using STNUpdater.Models;

namespace STNUpdater
{
    public class DbRepository
    {
        public DbRepository()
        {
            ConnectionString = ConfigurationManager.ConnectionStrings["StonetDbConnectionString"].ConnectionString;
        }

        public List<string> GetProductsNames()
        {
            Console.WriteLine("Getting existing products names from DB...");
            var results = new List<string>();

            using (var conn = new MySqlConnection(ConnectionString))
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "SELECT nume_produs FROM cs_stonet.produse;";
                conn.Open();
                var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    results.Add(Convert.ToString(reader["nume_produs"]));
                }
            }

            return results;
        }

        public List<Category> GetCategories()
        {
            Console.WriteLine("Getting categories from DB...");
            var results = new List<Category>();

            using (var conn = new MySqlConnection(ConnectionString))
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

        public List<Maker> GetMakers()
        {
            Console.WriteLine("Getting makers from DB...");
            var results = new List<Maker>();

            using (var conn = new MySqlConnection(ConnectionString))
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

        public void InsertProducts(IEnumerable<Product> products)
        {
            Console.WriteLine("Saving products in DB...");
            if (!products.Any()) return;
            using (var conn = new MySqlConnection(ConnectionString))
            using (var cmd = conn.CreateCommand())
            {
                conn.Open();
                var transaction = conn.BeginTransaction();
                cmd.Transaction = transaction;

                try
                {
                    products.ToList().ForEach(p =>
                    {
                        p.ShortDescription = p.ShortDescription.Replace("\'", "");
                        cmd.CommandText =
                            "INSERT INTO cs_stonet.produse (nume_produs, id_categorie,id_producator," +
                            "model,cod_producator,scurta_descriere,garantie)" +
                            $"VALUES('{p.Name}',{p.CategoryId},{p.MakerId},'{p.Model}','{p.Code}','{p.ShortDescription.Remove(p.ShortDescription.Length - 1)}',{p.Warranty})";
                        cmd.ExecuteNonQuery();
                    });
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    Console.WriteLine($"An error has occured in InsertProducts method: {ex.Message}");
                }
                finally
                {
                    conn.Close();
                }
            }
        }

        public string ConnectionString { get; set; }
    }
}