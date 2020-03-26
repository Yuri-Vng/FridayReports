using System;
using System.Collections.Generic;
using System.Text;

namespace Vng.Uchet
{
    // примеры кода    
    class Help
    {
        //DateTime now = DateTime.Now;

        //string connectionString =
        //@"Dsn=MS Access Database; Dbq=X:\VNG\UchDat.accdb;
        //            defaultdir=X:\VNG;driverid=25;fil=MS Access;
        //            maxbuffersize=2048;pagetimeout=5;uid=admin";

        //connectionString = @$"Dsn=MS Access Database; Dbq={Path.Combine(cnDir, dbName)};
        //                        defaultdir={cnDir};driverid=25;fil=MS Access;
        //                        maxbuffersize=2048;pagetimeout=5;uid=admin";

        #region CurrentDirectory
        // string projectDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\.."));
        //string s = Environment.CurrentDirectory;
        //string ss = Directory.GetCurrentDirectory();
        //string projectDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\.."));
        //string directory = AppDomain.CurrentDomain.BaseDirectory;
        //string dir = this.GetType().Module.FullyQualifiedName;
        //string s = Environment.CurrentDirectory;
        //Console.WriteLine($"s={s}");
        //string ss = System.IO.Directory.GetCurrentDirectory();
        //Console.WriteLine($"ss={ss}");
        //string sss = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\.."));
        //Console.WriteLine($"sss={sss}");
        //string ssss = AppDomain.CurrentDomain.BaseDirectory;
        //Console.WriteLine($"ssss={ssss}");
        //string sssss = ss.GetType().Module.FullyQualifiedName;
        //Console.WriteLine($"sssss={sssss}");
        #endregion

        #region Асинхронный вызов
        //using System.Threading.Tasks;
        //ReadDataAsync().GetAwaiter();
        //http://novaevalex.blogspot.com/2013/12/fillasync-dbdataadapter-net-framework.html
        //private static async Task ReadDataAsync()
        //{
        //    var t =  new  Task(() =>  LoadData());
        //     t.Start();
        //    return;
        //}
        //await connection.OpenAsync();
        //OdbcDataReader reader = await command.ExecuteReaderAsync();
        //SqlDataReader reader = await command.ExecuteReaderAsync();
        #endregion

        #region Варианты форматирования
        //oExcel.ActiveCell.NumberFormat = "#,##0.0"     '"#,##0.00"
        //oExcel.ActiveCell.NumberFormat = "dd/mm/yyyyг.;@"
        //oExcel.ActiveCell.NumberFormat = "[$-FC19]dd mmmm yyyy г.;@"
        //oExcel.ActiveCell.NumberFormat = "m/d/yyyy"
        //oExcel.ActiveCell.NumberFormat = "@"

        //Excel.Range c1 = (Excel.Range)xlSheet.Cells[topRow, 2];                         //"B10"
        //Excel.Range c2 = (Excel.Range)xlSheet.Cells[topRow + dt.Rows.Count - 1, 2];
        //Excel.Range range = xlSheet.get_Range(c1, c2);
        //range.HorizontalAlignment = Excel.Constants.xlCenter;
        ////range.Font.Size = 9;
        ////range.Font.Name = "Arial";
        //range.NumberFormat = "@";

        //oExcel.Columns("A:A").ColumnWidth = 8
        //oExcel.Columns("B:B").ColumnWidth = 25
        //oExcel.Rows("4:4").Font.Bold = True
        //oExcel.Rows("4:4").RowHeight = 15
        //oExcel.Rows("4:4").Font.Size = 7
        //oExcel.Rows("4:4").VerticalAlignment = xlCenter
        //oExcel.Rows("4:4").Interior.ColorIndex = 8
        //oExcel.Rows("4:4").HorizontalAlignment = xlCenter
        #endregion

        #region Проверка на Null
        //https://www.codeproject.com/Articles/19269/Export-large-data-from-a-GridView-and-DataReader-t
        //    while (dr.Read())
        //    {
        //        sb = new StringBuilder();

        //        for (int col = 0; col < dr.FieldCount - 1; col++)
        //        {
        //            if (!dr.IsDBNull(col))
        //                sb.Append(dr.GetValue(col).ToString().Replace(",", " "));
        //            sb.Append(",");
        //        }
        //        if (!dr.IsDBNull(dr.FieldCount - 1))
        //            sb.Append(dr.GetValue(dr.FieldCount - 1).ToString().Replace(",", " "));
        //        Response.Write(sb.ToString() + "\n");
        //        Response.Flush();
        //    }
        //    dr.Dispose();
        #endregion

        #region Вaрианты загрузки данных

        //private void ExcelData(Worksheet xlS)
        //{
        //    LibToExcel xl = new LibToExcel(xlS);

        //    try
        //    {
        //        int topRow = 10;
        //        System.Data.DataTable dt = new System.Data.DataTable();

        //        // Выгружаем DataReader в таблицу dt
        //        dt.Load(Reader);

        //        int qq1 = dt.Rows.Count;

        //        // Создаём двухмерный массив и загружаем в него таблицу
        //        //https://qarchive.ru/88699_zapis__massiva_v_diapazon_excel

        //        object[,] arr = new object[dt.Rows.Count, dt.Columns.Count];
        //        for (int r = 0; r < dt.Rows.Count; r++)
        //        {
        //            DataRow dr = dt.Rows[r];
        //            for (int c = 0; c < dt.Columns.Count; c++)
        //            {
        //                arr[r, c] = dr[c];
        //            }
        //        }

        //        #region Варианты форматирования
        //        //oExcel.ActiveCell.NumberFormat = "#,##0.0"     '"#,##0.00"
        //        //oExcel.ActiveCell.NumberFormat = "dd/mm/yyyyг.;@"
        //        //oExcel.ActiveCell.NumberFormat = "[$-FC19]dd mmmm yyyy г.;@"
        //        //oExcel.ActiveCell.NumberFormat = "m/d/yyyy"
        //        //oExcel.ActiveCell.NumberFormat = "@"

        //        //Excel.Range c1 = (Excel.Range)xlSheet.Cells[topRow, 2];                         //"B10"
        //        //Excel.Range c2 = (Excel.Range)xlSheet.Cells[topRow + dt.Rows.Count - 1, 2];
        //        //Excel.Range range = xlSheet.get_Range(c1, c2);
        //        //range.HorizontalAlignment = Excel.Constants.xlCenter;
        //        ////range.Font.Size = 9;
        //        ////range.Font.Name = "Arial";
        //        //range.NumberFormat = "@";

        //        //oExcel.Columns("A:A").ColumnWidth = 8
        //        //oExcel.Columns("B:B").ColumnWidth = 25
        //        //oExcel.Rows("4:4").Font.Bold = True
        //        //oExcel.Rows("4:4").RowHeight = 15
        //        //oExcel.Rows("4:4").Font.Size = 7
        //        //oExcel.Rows("4:4").VerticalAlignment = xlCenter
        //        //oExcel.Rows("4:4").Interior.ColorIndex = 8
        //        //oExcel.Rows("4:4").HorizontalAlignment = xlCenter
        //        #endregion

        //        // форматирование столбцов для вывода данных
        //        xl.ColumnFormat(2, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "@");   //инв.№
        //        xl.ColumnFormat(3, topRow, topRow + dt.Rows.Count - 1, true, 10, 'L', "@");   //модель
        //        xl.ColumnFormat(4, topRow, topRow + dt.Rows.Count - 1, true, 0, 'C', "@");    //гос.№
        //        xl.ColumnFormat(5, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "0000");    //год
        //        xl.ColumnFormat(6, topRow, topRow + dt.Rows.Count - 1, false, 0, 'L', "@");    //VIN
        //        xl.RegionFormat(7, 9, topRow, topRow + dt.Rows.Count - 1, false, 10, 'L', "@");    //Кузов-Шасси-Двиг.

        //        //xl.ColumnFormat(7, topRow, topRow + dt.Rows.Count - 1, false, 10, 'L', "@", xlS);    //Кузов
        //        //xl.ColumnFormat(8, topRow, topRow + dt.Rows.Count - 1, false, 10, 'L', "@", xlS);    //Шасси
        //        //xl.ColumnFormat(9, topRow, topRow + dt.Rows.Count - 1, false, 10, 'L', "@", xlS);    //Двиг.

        //        xl.ColumnFormat(10, topRow, topRow + dt.Rows.Count - 1, true, 9, 'C', "@");    //ПТС
        //        xl.ColumnFormat(11, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "dd/mm/yyyy");    //дата поступления
        //        xl.ColumnFormat(12, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "@");    //бюджет
        //        xl.ColumnFormat(13, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "dd/mm/yy");    //дата приказ ввода в эксп.
        //        xl.ColumnFormat(14, topRow, topRow + dt.Rows.Count - 1, true, 0, 'L', "@");    //приказ ввода в эксп.
        //        xl.ColumnFormat(15, topRow, topRow + dt.Rows.Count - 1, true, 10, 'L', "@");    //закреплен
        //        xl.ColumnFormat(16, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "dd/mm/yy");    //дата приказ ввода в эксп.
        //        xl.ColumnFormat(17, topRow, topRow + dt.Rows.Count - 1, true, 0, 'L', "@");    //приказ списания (передачи)

        //        xl.RegionFormat(18, 19, topRow, topRow + dt.Rows.Count - 1, true, 10, 'L', "@");    //куда-примечания

        //        //xl.ColumnFormat(18, topRow, topRow + dt.Rows.Count - 1, true, 10, 'L', "@", xlS);    //куда
        //        //xl.ColumnFormat(19, topRow, topRow + dt.Rows.Count - 1, true, 10, 'L', "@", xlS);    //примечания

        //        // Определяем диапазон таблицы в который зальём массив
        //        Excel.Range c1 = (Excel.Range)xlS.Cells[topRow, 1];
        //        Excel.Range c2 = (Excel.Range)xlS.Cells[topRow + dt.Rows.Count - 1, dt.Columns.Count];
        //        Excel.Range range = xlS.get_Range(c1, c2);
        //        range.Value = arr;

        //        range.VerticalAlignment = Excel.Constants.xlCenter;

        //        #region Проверка на Null
        //        //https://www.codeproject.com/Articles/19269/Export-large-data-from-a-GridView-and-DataReader-t
        //        //    while (dr.Read())
        //        //    {
        //        //        sb = new StringBuilder();

        //        //        for (int col = 0; col < dr.FieldCount - 1; col++)
        //        //        {
        //        //            if (!dr.IsDBNull(col))
        //        //                sb.Append(dr.GetValue(col).ToString().Replace(",", " "));
        //        //            sb.Append(",");
        //        //        }
        //        //        if (!dr.IsDBNull(dr.FieldCount - 1))
        //        //            sb.Append(dr.GetValue(dr.FieldCount - 1).ToString().Replace(",", " "));
        //        //        Response.Write(sb.ToString() + "\n");
        //        //        Response.Flush();
        //        //    }
        //        //    dr.Dispose();
        //        #endregion
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine(ex.Message);
        //    }
        //}
        #endregion

        #region Вариант с двумя запросами и параметрами
        //SqlCommand command = new SqlCommand(
        //    "SELECT CategoryID, CategoryName FROM dbo.Categories;" +
        //    "SELECT EmployeeID, LastName FROM dbo.Employees",
        //    connection);
        //connection.Open();
        //OdbcCommand command = new OdbcCommand(queryString + queryString2, connection);

        //// создаем параметр для возраста
        //SqlParameter ageParam = new SqlParameter("@age", age);
        //// добавляем параметр к команде
        //command.Parameters.Add(ageParam);
        #endregion

        #region Последовательное считывание по строкам
        //while (reader.Read())
        //{
        //    Console.WriteLine("\t{0}\t{1}\t{2}\t{3}", reader[0], reader[1], reader[2], reader[3]);                    
        //}
        #endregion

        #region releaseExcel()
        ////Показываем ексель
        //xlApp.Visible = true;

        //xlApp.Interactive = true;
        //xlApp.ScreenUpdating = true;
        //xlApp.UserControl = true;

        ////Отсоединяемся от Excel
        ////releaseObject(xlSheetRange);
        ////releaseObject(xlSheet);
        //releaseObject(xlApp);
        #endregion

        #region json
        //      {
        //"myConfig": {
        //  "item1":  "config options" ,
        //  "PathDataBase":  "X:\\VNG" ,
        //  "item2":  "UchDat accdb"  
        //},
        //"exclude": [
        //  "**/bin",
        //  "**/bower_components",
        //  "**/jspm_packages",
        //  "**/node_modules",
        //  "**/obj",
        //  "**/platforms"
        //]
        //  }
        #endregion

    }
}
