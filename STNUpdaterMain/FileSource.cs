using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using STNUpdater.Models;

namespace STNUpdater
{
    internal class FileSource
    {
        public FileSource()
        {
            FileName = GetFileName();
            ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\"";
        }

        private string GetFileName()
        {
            var fileName = string.Empty;

            var dialog = new OpenFileDialog
            {
                Filter = "All Files (*.xlsx)|*.xlsx|(*.xls)|*.xls",
                FilterIndex = 1,
                Multiselect = false
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                fileName = dialog.FileName;
            }

            return fileName;
        }

        public List<Product> GetProducts()
        {
            Console.WriteLine("Extracting products from file...");
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
                            if (ds.Tables[0].Columns.Count < 10) continue;
                            products.AddRange(MapDataSetToProducts(ds));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var message = $"An error has occured in GetProducts method: {ex.Message}";
                Console.WriteLine(message);
            }

            int warranty;
            products = products.Where(p => p.Warranty != null && int.TryParse(p.Warranty, out warranty)).ToList();

            return products;
        }

        private IEnumerable<Product> MapDataSetToProducts(DataSet ds)
        {
            var products = ds.Tables[0].AsEnumerable()
                                       .Select(dataRow => new Product
                                       {
                                           Name = dataRow.Field<string>("F5"),
                                           Category = MapCategoryName(dataRow.Field<string>("F2")),
                                           Maker = dataRow.Field<string>("F4"),
                                           Model = dataRow.Field<string>("F4"),
                                           Code = dataRow.Field<string>("F5"),
                                           ShortDescription = dataRow.Field<string>("F6"),
                                           Warranty = dataRow.Field<string>("F10")
                                       }).ToList();
            return products;
        }

        private string MapCategoryName(string categoryName)
        {
            if (string.IsNullOrEmpty(categoryName?.Trim())) return categoryName;

            switch (categoryName.ToLower())
            {
                case "desktop components":
                case "notebook components":
                    return "Componente";
                case "nb/dt accessories":
                    return "Accesorii";
                case "external hdd":
                    return "Hard Disk-uri";
                case "usb flash drive":
                    return "Flash USB";
                case "memory card":
                    return "Compact Flash Card";
                case "usb hub":
                case "network":
                case "switch & accessories":
                case "router & accessories":
                case "wireless":
                case "kvm & remote manag":
                case "network adapter":
                    return "Retelistica";
                case "hdd enclosure":
                    return "Hard Disk-uri";
                case "external odd":
                    return "DVD Writer";
                case "home entertainment":
                    return "Boxe";
                case "input devices":
                    return "Tastaturi";
                case "personal audio":
                    return "Casti";
                case "webcam":
                    return "Webcam";
                case "accessories":
                    return "Accesorii";
                case "other peripherals":
                    return "Periferice";
                case "tablet accessories":
                    return "Tablete";
                case "ups":
                    return "UPS";
                default:
                    return categoryName;
            }
        }

        private string FileName { get; set; }
        private string ConnectionString { get; set; }
    }
}