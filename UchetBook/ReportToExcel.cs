using System;

using System.Data;
using System.Data.Odbc;

// Добавим COM MS Office 15.0 Object Library и MS Ecxel 15.0 Object Library
using Excel = Microsoft.Office.Interop.Excel;

#region Excel in Core.3
/*
// Если в версии .Net Core 3.1 не работает Ecxel.
// https://stackoverflow.com/questions/58130446/net-core-3-0-and-ms-office-interop
// Подписываемся в свойствах проекта на .Net Core 2.2. Выгружаем проект. 
// Открываем выгруженный проект. Копируем всё из тэга <ItemGroup>  

<ItemGroup>
    <COMReference Include = "Microsoft.Office.Core" >
      < Guid >{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}</Guid>
      <VersionMajor>2</VersionMajor>
      <VersionMinor>8</VersionMinor>
      <Lcid>0</Lcid>
      <WrapperTool>primary</WrapperTool>
      <Isolated>False</Isolated>
      <EmbedInteropTypes>True</EmbedInteropTypes>
    </COMReference>

    ... more references
</ItemGroup>

// Далее открываем наш проект .Net Core 3.1 и меняем текст на скопированный.
// Теперь Excel должен работать
*/
#endregion

namespace Vng.Uchet
{
    public class ReportToExcel
    {
        Excel.Application xlApp;                            //Екземпляр приложения Excel
        Excel.Worksheet xlSheet;                            //Лист
        //Excel.Range xlSheetRange;                           //Выделеная область

        public void ExelObjecCars(OdbcDataReader reader)
        {
            xlApp = new Excel.Application();
            int startRow = 10;
            DateTime now = DateTime.Now;

            try
            {
                //добавляем книгу
                xlApp.Workbooks.Add(Type.Missing);

                if (xlApp.Sheets.Count == 1)        //для office 13
                    xlApp.Sheets.Add();

                xlApp.EnableEvents = false;

                //I.Счета в работе
                //выбираем лист на котором будем работать (Лист 1)
                xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                //Название листа
                xlSheet.Name = "Список ТС";
                //xlSheet.Tab.ThemeColor = Excel.XlThemeColor.xlThemeColorLight2;
                //цвет вкладки
                xlSheet.Tab.Color = 255;

                //string Zagolovok = "Счета " + "qqq" + " в работе на " + now.ToString("dd.MM.yyyy") + "г.";
                string Zagolovok = "Книга учета ТС ФГКУ «УВО ВНГ России по городу Москве»";

                // названия столбцов таблицы
                Shapka(Zagolovok, xlSheet);
                // вывод данных
                ExcelData(reader, xlSheet);
                // итоги
                //Podval();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                //Показываем ексель
                xlApp.Visible = true;

                xlApp.Interactive = true;
                xlApp.ScreenUpdating = true;
                xlApp.UserControl = true;

                //        //Отсоединяемся от Excel
                //        releaseObject(xlSheetRange);
                //        releaseObject(xlSheet);
                //        releaseObject(xlApp);
            }

        }

        //создание заголовка таблицы
        private void Shapka(string Zagolovok, Excel.Worksheet xlSheet)
        {
            Excel.Range xlSheetRange;                               //Выделеная область

            //Задаем диапазон
            xlSheetRange = xlSheet.get_Range("A3", "S9");
            //Задаем выравнивание по центру для выбранного диапазона
            xlSheetRange.HorizontalAlignment = Excel.Constants.xlCenter;
            xlSheetRange.VerticalAlignment = Excel.Constants.xlCenter;
            //Задаем размер шрифта для выбранного диапазона
            xlSheetRange.Font.Size = 11;

            CellMerge(Zagolovok, "A3", "S3", 0, false, 14, 'C', 'C', xlSheet);

            //////////////////////////////////////////////////////////////////////////////////
            //рисуем шапку таблицы
            //////////////////////////////////////////////////////////////////////////////////

            //Выбираем диапазон (ячейку) для вывода 
            CellMerge("№ п/п", "A5", "A9", 5.14, true, 0, 'N', 'N', xlSheet);
            CellMerge("Инв. №", "B5", "B9", 11.57, false, 0, 'N', 'N', xlSheet);
            CellMerge("Марка, модель ТС", "C5", "C9", 16.0, true, 0, 'N', 'N', xlSheet);
            CellMerge("Гос. №", "D5", "D9", 12.14, false, 0, 'N', 'N', xlSheet);
            CellMerge("Год выпуска", "E5", "E9", 7.71, true, 0, 'N', 'N', xlSheet);
            CellMerge("VIN", "F5", "F9", 20.0, false, 0, 'N', 'N', xlSheet);
            CellMerge("№ кузова", "G5", "G9", 20.0, false, 0, 'N', 'N', xlSheet);
            CellMerge("№ шасси", "H5", "H9", 20.0, false, 0, 'N', 'N', xlSheet);
            CellMerge("№ двигателя", "I5", "I9", 20.0, false, 0, 'N', 'N', xlSheet);
            CellMerge("№ ПТС", "J5", "J9", 14.43, false, 0, 'N', 'N', xlSheet);
            CellMerge("Сведения о поступлении ТС", "K5", "L6", 0, true, 0, 'N', 'N', xlSheet);
            CellMerge("Дата", "K7", "K9", 9.86, false, 0, 'N', 'N', xlSheet);
            CellMerge("Источник приобре-тения", "L7", "L9", 10.0, true, 0, 'N', 'N', xlSheet);
            CellMerge("Приказ ввода в эксплуатацию", "M5", "N6", 13.71, true, 0, 'N', 'N', xlSheet);
            CellMerge("Дата", "M7", "M9", 9.86, false, 0, 'N', 'N', xlSheet);
            CellMerge("Номер", "N7", "N9", 10.0, false, 0, 'N', 'N', xlSheet);
            CellMerge("Орган (служба) за которым закреплено ТС", "O5", "O9", 17.57, true, 0, 'N', 'N', xlSheet);
            CellMerge("Сведения о передаче или списании ТС", "P5", "R6", 0, true, 0, 'N', 'N', xlSheet);
            CellMerge("Приказ", "P7", "Q7", 0, false, 0, 'N', 'N', xlSheet);
            CellMerge("Дата", "P8", "P9", 9.86, false, 0, 'N', 'N', xlSheet);
            CellMerge("Номер", "Q8", "Q9", 10.0, false, 0, 'N', 'N', xlSheet);
            CellMerge("Куда передено", "R7", "R9", 17.0, true, 0, 'N', 'N', xlSheet);
            CellMerge("Примечания", "S5", "S9", 20.0, false, 0, 'N', 'N', xlSheet);

            //Границы
            xlSheetRange = xlSheet.get_Range("A5", "S9");
            //Устанавливаем цвет обводки
            xlSheetRange.Borders.ColorIndex = 1;
            //Устанавливаем стиль и толщину линии
            xlSheetRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            xlSheetRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            ////Задаем выравнивание по центру
            //xlSheetRange.HorizontalAlignment = Excel.Constants.xlCenter;
            //xlSheetRange.VerticalAlignment = Excel.Constants.xlCenter;
        }

        private void CellMerge(string title, string cell1, string cell2, double tWidth, bool wrpText,
                                    double tFont, char tHor, char tVer, Excel.Worksheet xlSheet)
        {
            Excel.Range xlSheetRange;               //Выделеная область

            xlSheetRange = xlSheet.get_Range(cell1, cell2);
            if (wrpText)
            {
                xlSheetRange.WrapText = wrpText;
            }
            //Объединяем ячейки
            xlSheetRange.Merge(Type.Missing);
            xlSheetRange.Value2 = title;
            if (tWidth > 0)
            {
                xlSheetRange.ColumnWidth = tWidth;
            }
            if (tFont > 0)
            {
                xlSheetRange.Font.Size = tFont;
            }

            switch (tHor)
            {
                case 'C':                       //Задаем выравнивание по центру
                    xlSheetRange.HorizontalAlignment = Excel.Constants.xlCenter;
                    break;
                case 'L':
                    xlSheetRange.HorizontalAlignment = Excel.Constants.xlLeft;
                    break;
                case 'R':
                    xlSheetRange.HorizontalAlignment = Excel.Constants.xlRight;
                    break;
                default:
                    break;
            }

            switch (tVer)
            {
                case 'C':
                    //Задаем выравнивание по центру
                    xlSheetRange.VerticalAlignment = Excel.Constants.xlCenter;
                    break;
                case 'T':
                    xlSheetRange.VerticalAlignment = Excel.Constants.xlTop;
                    break;
                case 'B':
                    xlSheetRange.VerticalAlignment = Excel.Constants.xlBottom;
                    break;
                default:
                    break;
            }
        }

        private void ExcelData(OdbcDataReader reader, Excel.Worksheet xlSheet)
        {
            //DateTime now = DateTime.Now;
            try
            {
                ////добавляем книгу
                //xlApp.Workbooks.Add(Type.Missing);

                //if (xlApp.Sheets.Count == 1)        //для office 13
                //    xlApp.Sheets.Add();

                //xlApp.EnableEvents = false;

                ////I.Список ТС
                ////выбираем лист на котором будем работать (Лист 1)
                //xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                ////Название листа
                //xlSheet.Name = "Список ТС";
                ////xlSheet.Tab.ThemeColor = Excel.XlThemeColor.xlThemeColorLight2;
                ////цвет вкладки
                //xlSheet.Tab.Color = 255;

                ////string Zagolovok = "Книга учета ТС ФГКУ «УВО ВНГ России по городу Москве»";

                //Excel.Range xlSheetRange;               //Выделеная область

                //Задаем диапазон
           
                DataTable dt = new DataTable();
                // если есть данные
                if (reader.HasRows)
                {
                    // Выгружаем DataReader в таблицу dt
                    dt.Load(reader);
                }

                //DataTable dt2 = new DataTable();
                //dt2.Load(reader);
                
                // Если несколько наборов данных
                //while (reader.HasRows)
                //{
                //    while (reader.Read())
                //    {
                //    }
                //    reader.NextResult();
                //}

                //https://qarchive.ru/88699_zapis__massiva_v_diapazon_excel

                // Создаём двухмерный массив и загружаем в него таблицу
                object[,] arr = new object[dt.Rows.Count, dt.Columns.Count];
                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    DataRow dr = dt.Rows[r];
                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        arr[r, c] = dr[c];
                    }
                }

                //int topRow = 1;
                int topRow = 10;

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

                // форматирование столбцов для вывода данных
                ColumnFormat(2, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "@", xlSheet);   //инв.№
                ColumnFormat(3, topRow, topRow + dt.Rows.Count - 1, true, 10, 'L', "@", xlSheet);   //модель
                ColumnFormat(4, topRow, topRow + dt.Rows.Count - 1, true, 0, 'C', "@", xlSheet);    //гос.№
                ColumnFormat(5, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "0000", xlSheet);    //год
                ColumnFormat(6, topRow, topRow + dt.Rows.Count - 1, false, 0, 'L', "@", xlSheet);    //VIN
                ColumnFormat(7, topRow, topRow + dt.Rows.Count - 1, false, 0, 'L', "@", xlSheet);    //Кузов
                ColumnFormat(8, topRow, topRow + dt.Rows.Count - 1, false, 0, 'L', "@", xlSheet);    //Шасси
                ColumnFormat(9, topRow, topRow + dt.Rows.Count - 1, false, 0, 'L', "@", xlSheet);    //Двиг.
                ColumnFormat(10, topRow, topRow + dt.Rows.Count - 1, true, 0, 'C', "@", xlSheet);    //ПТС
                ColumnFormat(11, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "dd/mm/yyyy", xlSheet);    //дата поступления
                ColumnFormat(12, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "@", xlSheet);    //бюджет
                ColumnFormat(13, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "dd/mm/yy", xlSheet);    //дата приказ ввода в эксп.
                ColumnFormat(14, topRow, topRow + dt.Rows.Count - 1, true, 0, 'L', "@", xlSheet);    //приказ ввода в эксп.
                ColumnFormat(15, topRow, topRow + dt.Rows.Count - 1, true, 10, 'L', "@", xlSheet);    //закреплен
                ColumnFormat(16, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "dd/mm/yy", xlSheet);    //дата приказ ввода в эксп.
                ColumnFormat(17, topRow, topRow + dt.Rows.Count - 1, true, 0, 'L', "@", xlSheet);    //приказ списания (передачи)
                ColumnFormat(18, topRow, topRow + dt.Rows.Count - 1, true, 10, 'L', "@", xlSheet);    //куда
                ColumnFormat(19, topRow, topRow + dt.Rows.Count - 1, true, 10, 'L', "@", xlSheet);    //закреплен

                // Определяем диапазон таблицы в который зальём массив
                Excel.Range c1 = (Excel.Range)xlSheet.Cells[topRow, 1];
                Excel.Range c2 = (Excel.Range)xlSheet.Cells[topRow + dt.Rows.Count - 1, dt.Columns.Count];
                Excel.Range range = xlSheet.get_Range(c1, c2);
                range.Value = arr;
                
                range.VerticalAlignment = Excel.Constants.xlCenter;

                //oExcel.Columns("A:A").ColumnWidth = 8
                //oExcel.Columns("B:B").ColumnWidth = 25
                //oExcel.Rows("4:4").Font.Bold = True
                //oExcel.Rows("4:4").RowHeight = 15
                //oExcel.Rows("4:4").Font.Size = 7
                //oExcel.Rows("4:4").VerticalAlignment = xlCenter
                //oExcel.Rows("4:4").Interior.ColorIndex = 8
                //oExcel.Rows("4:4").HorizontalAlignment = xlCenter

                //try
                //{
                //    cn.Open();
                //    SqlDataReader dr = cmd.ExecuteReader();
                //    StringBuilder sb = new StringBuilder();
                //    //Add Header

                //    for (int count = 0; count < dr.FieldCount; count++)
                //    {
                //        if (dr.GetName(count) != null)
                //            sb.Append(dr.GetName(count));
                //        if (count < dr.FieldCount - 1)
                //        {
                //            sb.Append(",");
                //        }
                //    }
                //    Response.Write(sb.ToString() + "\n");
                //    Response.Flush();
                //    //Append Data

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
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                //Показываем ексель
                xlApp.Visible = true;

                xlApp.Interactive = true;
                xlApp.ScreenUpdating = true;
                xlApp.UserControl = true;

                //        //Отсоединяемся от Excel
                //        releaseObject(xlSheetRange);
                //        releaseObject(xlSheet);
                //        releaseObject(xlApp);
            }
        }
        private void ColumnFormat(int column, int topRow, int bottomRow, bool wrpText,
                                    double tFont, char tHor, string tFormat, Excel.Worksheet xlSheet)
        {
            Excel.Range c1 = (Excel.Range)xlSheet.Cells[topRow, column];                         //"B10"
            Excel.Range c2 = (Excel.Range)xlSheet.Cells[bottomRow, column];
            Excel.Range range = xlSheet.get_Range(c1, c2);
            if (wrpText)
            {
                range.WrapText = wrpText;
            }
            //range.HorizontalAlignment = Excel.Constants.xlCenter;
            switch (tHor)
            {
                case 'C':                       //Задаем выравнивание по центру
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    break;
                case 'L':
                    range.HorizontalAlignment = Excel.Constants.xlLeft;
                    break;
                case 'R':
                    range.HorizontalAlignment = Excel.Constants.xlRight;
                    break;
                default:
                    break;
            }
            //range.Font.Size = 9;
            if (tFont > 0)
            {
                range.Font.Size = tFont;
            }
            //range.Font.Name = "Arial";
            //range.NumberFormat = "@";
            range.NumberFormat = tFormat; 
        }
    }
}
