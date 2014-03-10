using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Data.OleDb;

namespace SiswaDapperNET1
{
    class Program
    {

        private static OleDbConnection GetOpenConnection()
        {
            OleDbConnection conn = null;

            try
            {

                var appDir = System.IO.Directory.GetCurrentDirectory();
                var strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + appDir + "\\SISWA.MDB;User Id=admin;Password=;";

                conn = new OleDbConnection(strConn);
                conn.Open();
            }
            catch (Exception)
            {
            }

            return conn;
        }

        private static List<Siswa> GetDataSiswa()
        {
            var daftarSiswa = new List<Siswa>();

            using (var conn = GetOpenConnection())
            {
                var strSql = "SELECT nis, nama FROM siswa";
                using (var cmd = new OleDbCommand(strSql, conn))
                {
                    using (var dtr = cmd.ExecuteReader())
                    {
                        while (dtr.Read())
                        {
                            // proses mapping dari row ke object
                            var siswa = new Siswa();
                            siswa.Nis = dtr["nis"] == null ? string.Empty : dtr.GetString(0);
                            siswa.Nama = dtr["nama"] == null ? string.Empty : dtr.GetString(1).ToUpper();

                            daftarSiswa.Add(siswa);
                        }
                    }
                }
            }

            return daftarSiswa;
        }

        static void Main(string[] args)
        {
            /*
            using (var conn = GetOpenConnection())
            {
                var strSql = "SELECT nis, nama FROM siswa";
                using (var cmd = new OleDbCommand(strSql, conn))
                {
                    using (var dtr = cmd.ExecuteReader())
                    {
                        Console.WriteLine("NIS\tNAMA");
                        Console.WriteLine("===================================");
                        while (dtr.Read())
                        {
                            Console.WriteLine(dtr["nis"] + "\t" + dtr["nama"].ToString().ToUpper());
                        }
                    }
                }
            }
            */

            Console.WriteLine("NIS\tNAMA");
            Console.WriteLine("===================================");

            var daftarSiswa = GetDataSiswa();
            foreach (var siswa in daftarSiswa)
            {
                Console.WriteLine(siswa.Nis + "\t" + siswa.Nama);
            }

            Console.ReadKey();
        }
    }
}
