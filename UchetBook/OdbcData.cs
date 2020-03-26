using System;
using System.Data;
using System.Data.Odbc;

using Microsoft.Extensions.Configuration;

namespace Vng.Uchet
{
    public class OdbcData
    {
        string? queryString;                // строка запроса

        // параметры берем из конфигурационного файла
        readonly string connectionString;

        public OdbcData() :this ("UB") 
        {
        }
        public OdbcData(string tCod) 
        {
            // файл конфигурации
            IConfigurationRoot configuration = new ConfigurationBuilder()
            .AddJsonFile("config.json", optional: true)
            .Build();

            // анонимный тип
            var config = new
            {
                MyDbConnectionString = configuration["ConnectionStrings:MyDb"],
                PathSettings = new
                {
                    Dir = configuration["AccessDb:PathDataBase"],
                    File = configuration["AccessDb:FileName"]
                },
                cnn = new
                {
                    Dsn = configuration["connectionStringsAccessDb:Dsn"],
                    Dbq = configuration["connectionStringsAccessDb:Dbq"],
                    DefaultDir = configuration["connectionStringsAccessDb:defaultdir"],
                    DriverId = configuration["connectionStringsAccessDb:driverid"],
                    Fil  =configuration["connectionStringsAccessDb:fil"],
                    MaxBufferSize = configuration["connectionStringsAccessDb:maxbuffersize"],
                    PageTimeout = configuration["connectionStringsAccessDb:pagetimeout"],
                    Uid = configuration["connectionStringsAccessDb:uid"]
                }
            };

            #region Path.Combine(cnDir, dbName)
            //readonly string cnDir;              // = "X:\\VNG\\";
            //readonly string dbName;             // = "UchDat.accdb";
            //cnDir = config.PathSettings.Dir;
            //dbName = config.PathSettings.File;
            //connectionString = @$"Dsn=MS Access Database; Dbq={Path.Combine(cnDir, dbName)};
            //                        defaultdir={cnDir};driverid=25;fil=MS Access;
            //                        maxbuffersize=2048;pagetimeout=5;uid=admin";
            #endregion

            // The connection string 
            // PM> Install-Package System.Data.Odbc -Version 4.7.0
            connectionString = config.cnn.Dsn + config.cnn.Dbq + config.cnn.DefaultDir 
                                + config.cnn.DriverId + config.cnn.Fil + config.cnn.MaxBufferSize 
                                + config.cnn.PageTimeout + config.cnn.Uid;
            switch (tCod)
            {
                case "UB":              // Книга учета
                    queryString =
                          "SELECT nn, inn, Model, GosN, GodV, Vin, KuzN, ShassiN, DvN, PtsN, dtPostup, "
                              + " Bdg, dtPrikEk, PrikEk, Pdr, dtPrikSp, PrikSp, Kuda, Appendix "
                              + " FROM tblUchetBook "
                              + " WHERE GodV > ? AND Bdg <> ? "
                              + " ORDER BY id_Book ASC, nn ASC, Model DESC;";
                    break;
                case "2":
                    Console.WriteLine("Функция пока отсутствует");
                    break;
                case "0":
                    //string queryString =
                    //    "SELECT tblCarsNG.inn2, tblCarsNG.GodV, tblCarsNG.GosN, tblCarModel.Model "
                    //        + "FROM tblCarModel INNER JOIN tblCarsNG "
                    //            + "ON tblCarModel.idModel = tblCarsNG.id_Model "
                    //        + "WHERE tblCarsNG.GodV > ? "
                    //        + "ORDER BY tblCarModel.Model DESC;";
                    break;
                default:
                    break;
            }
        }

        public DataTable LoadData()
        {
            DataTable dt = new DataTable();

            // Specify the parameter value.
            int paramValue1 = 1901;
            //string paramValue2 = "4949";
            string paramValue3 = "MB";

            // Create and open the connection in a using block. This ensures that 
            // all resources will be closed and disposed when the code exits.
            using (OdbcConnection connection = new OdbcConnection(connectionString))
            {
                //Create the Command and Parameter objects.
                OdbcCommand command = new OdbcCommand(queryString, connection);

                command.Parameters.AddWithValue("@GodV", paramValue1);
                //command.Parameters.AddWithValue("@inn", paramValue2);
                command.Parameters.AddWithValue("@Bdg", paramValue3);

                // Open the connection in a try/catch block.
                // Create and execute the DataReader.
                try
                {
                    connection.Open();
                    OdbcDataReader reader = command.ExecuteReader();

                    // если есть данные
                    if (reader.HasRows)
                    {
                        // Выгружаем DataReader в таблицу DataTable
                        dt.Load(reader);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            return dt;
        }
    }
}
