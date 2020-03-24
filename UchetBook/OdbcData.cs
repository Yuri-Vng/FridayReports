using System;
using System.Data;
using System.Data.Odbc;

namespace Vng.Uchet
{
    public class OdbcData
    {
        string? queryString;     // строка запроса
        string? connectionString;

        // параметры берем из конфигурационного файла
        readonly string cnDir = "X:\\VNG\\";
        readonly string dbName = "UchDat.accdb";

        public OdbcData() { }
        public OdbcData(string tCod) 
        {
            // The connection string 
            // PM> Install-Package System.Data.Odbc -Version 4.7.0
            connectionString = @$"Dsn=MS Access Database; Dbq={cnDir + dbName};
                                    defaultdir={cnDir};driverid=25;fil=MS Access;
                                    maxbuffersize=2048;pagetimeout=5;uid=admin";
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
