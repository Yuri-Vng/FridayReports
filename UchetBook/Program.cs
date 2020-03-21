using System;

using System.Data.Odbc;

//using System.Data;
//using System.Threading.Tasks;

// Добавим COM MS Office 15.0 Object Library и MS Ecxel 15.0 Object Library
//using Excel = Microsoft.Office.Interop.Excel;

namespace Vng.Uchet
{
    class Program
    {
        static void Main(string[] args)
        {
            //ReadDataAsync().GetAwaiter();
            LoadData();

            Console.WriteLine("Hello World!");
            //Console.ReadLine();
        }

        //http://novaevalex.blogspot.com/2013/12/fillasync-dbdataadapter-net-framework.html
        //private static async Task ReadDataAsync()
        //{
        //    var t =  new  Task(() =>  LoadData());
        //     t.Start();
        //    return;
        //}

        private static void LoadData()
        {
            // The connection string 
            // PM> Install-Package System.Data.Odbc -Version 4.7.0
            string connectionString =
                    @"Dsn=MS Access Database; Dbq=X:\VNG\UchDat.accdb;
                    defaultdir==X:\VNG;driverid=25;fil=MS Access;
                    maxbuffersize=2048;pagetimeout=5;uid=admin";

            string queryString =
                 "SELECT nn, inn, Model, GosN, GodV, Vin, KuzN, ShassiN, DvN, PtsN, dtPostup, "
                     + " Bdg, dtPrikEk, PrikEk, Pdr, dtPrikSp, PrikSp, Kuda, Appendix "
                     + " FROM tblUchetBook "
                     + " WHERE GodV > ? "
                     + " ORDER BY id_Book ASC, nn ASC, Model DESC;";

            //string queryString2 =
            //    "SELECT tblCarsNG.inn2, tblCarsNG.GodV, tblCarsNG.GosN, tblCarModel.Model "
            //        + "FROM tblCarModel INNER JOIN tblCarsNG "
            //            + "ON tblCarModel.idModel = tblCarsNG.id_Model "
            //        + "WHERE tblCarsNG.GodV > ? "
            //        + "ORDER BY tblCarModel.Model DESC;";

            // Specify the parameter value.
            int paramValue = 1900;

            // Create and open the connection in a using block. This ensures that 
            // all resources will be closed and disposed when the code exits.
            using (OdbcConnection connection = new OdbcConnection(connectionString))
            {
                //Create the Command and Parameter objects.
                //OdbcCommand command = new OdbcCommand(queryString + queryString2, connection);
                OdbcCommand command = new OdbcCommand(queryString, connection);

                //SqlCommand command = new SqlCommand(
                //    "SELECT CategoryID, CategoryName FROM dbo.Categories;" +
                //    "SELECT EmployeeID, LastName FROM dbo.Employees",
                //    connection);
                //connection.Open();

                //// создаем параметр для возраста
                //SqlParameter ageParam = new SqlParameter("@age", age);
                //// добавляем параметр к команде
                //command.Parameters.Add(ageParam);

                command.Parameters.AddWithValue("@name", paramValue);

                // Open the connection in a try/catch block.
                // Create and execute the DataReader.
                try
                {
                    connection.Open();
                    OdbcDataReader reader = command.ExecuteReader();
                    //await connection.OpenAsync();
                    //OdbcDataReader reader = await command.ExecuteReaderAsync();

                    //SqlDataReader reader = await command.ExecuteReaderAsync();
                    //while (reader.Read())
                    //{
                    //    Console.WriteLine("\t{0}\t{1}\t{2}\t{3}", reader[0], reader[1], reader[2], reader[3]);                    
                    //}

                    ReportToExcel xlTmpl = new ReportToExcel();                 
                    xlTmpl.ExelObjecCars(reader);

                    reader.Close();
                    //return reader;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
     }
 }
