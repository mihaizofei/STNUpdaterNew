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

            var connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\""; ;
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();

                var dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                List<string> sheetNames = new List<string>();
                sheetNames.AddRange(dt.Rows.Cast<DataRow>().Select(row => row["TABLE_NAME"].ToString()).Where(tableName => tableName.EndsWith("$") || tableName.EndsWith("$'")));
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT * FROM [" + sheetNames[0] + "] ";

                    var adapter = new OleDbDataAdapter(cmd);
                    var ds = new DataSet();
                    adapter.Fill(ds);
                }
            }


            Console.WriteLine(fileName);
            Console.ReadLine();
        }
    }
}
