﻿using System;
using System.Data;
using System.Data.Odbc;

// Добавим COM MS Office 15.0 Object Library и MS Ecxel 15.0 Object Library
using Excel = Microsoft.Office.Interop.Excel;

using Vng.Common;

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
        //Excel.Application xlApp;                            //Екземпляр приложения Excel
        //Excel.Worksheet xlSheet;                            //Лист
        //Excel.Range xlSheetRange;                           //Выделеная область

        public void ExelObjecCars(OdbcDataReader reader)
        {
            Excel.Application xlApp;                            //Екземпляр приложения Excel
            Excel.Worksheet xlSheet;                            //Лист

            //Excel.Range xlSheetRange;                           //Выделеная область
            //DateTime now = DateTime.Now;

            xlApp = new Excel.Application();

            try
            {
                //добавляем книгу
                xlApp.Workbooks.Add(Type.Missing);

                if (xlApp.Sheets.Count == 1)        //для office 13
                { 
                    xlApp.Sheets.Add(); 
                }                   
                xlApp.EnableEvents = false;     //отключить события в excel

            //I.Список ТС
                //выбираем лист на котором будем работать (Лист 1)
                xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                //Название листа
                xlSheet.Name = "Список ТС";
                //цвет вкладки
                xlSheet.Tab.Color = 255;
                // Заголовог таблицы
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
        private void Shapka(string Zagolovok, Excel.Worksheet xlS)
        {
            Excel.Range xlSheetRange;               //Выделеная область
            
            LibToExcel xl = new LibToExcel();

            xl.CellMerge(title: Zagolovok, cell1: "A3", cell2: "S3", tWidth: 0, wrpText: false,
                        tFont: 14, tHor: 'C', tVer: 'C', tOrient: 0, xlSh: xlS);

            //////////////////////////////////////////////////////////////////////////////////
            //рисуем шапку таблицы
            //////////////////////////////////////////////////////////////////////////////////

            //Задаем диапазон для подписей таблицы
            xlSheetRange = xlS.get_Range("A5", "S9");
            //Задаем выравнивание по центру для выбранного диапазона
            xlSheetRange.HorizontalAlignment = Excel.Constants.xlCenter;
            xlSheetRange.VerticalAlignment = Excel.Constants.xlCenter;
            //Задаем размер шрифта для выбранного диапазона
            xlSheetRange.Font.Size = 11;
            //Границы
            //Устанавливаем цвет обводки
            xlSheetRange.Borders.ColorIndex = 1;
            //Устанавливаем стиль и толщину линии
            xlSheetRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            xlSheetRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            //Выбираем диапазон (ячейку) для вывода 
            xl.CellMerge("№ п/п", "A5", "A9", 5.14, true, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("Инв. №", "B5", "B9", 11.57, false, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("Марка, модель ТС", "C5", "C9", 16.0, true, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("Гос. №", "D5", "D9", 12.14, false, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("Год выпуска", "E5", "E9", 5.0, true, 10, 'N', 'N', 90, xlS);
            xl.CellMerge("VIN", "F5", "F9", 20.0, false, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("№ кузова", "G5", "G9", 18.57, false, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("№ шасси", "H5", "H9", 18.57, false, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("№ двигателя", "I5", "I9", 18.57, false, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("№ ПТС", "J5", "J9", 13.3, false, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("Сведения о поступлении ТС", "K5", "L6", 0, true, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("Дата", "K7", "K9", 9.86, false, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("Источник приобре-тения", "L7", "L9", 0, true, 10, 'N', 'N', 0, xlS);
            xl.CellMerge("Приказ ввода в эксплуатацию", "M5", "N6", 13.71, true, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("Дата", "M7", "M9", 9.86, false, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("Номер", "N7", "N9", 10.0, false, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("Орган (служба) за которым закреплено ТС", "O5", "O9", 17.57, true, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("Сведения о передаче или списании ТС", "P5", "R6", 0, true, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("Приказ", "P7", "Q7", 0, false, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("Дата", "P8", "P9", 9.86, false, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("Номер", "Q8", "Q9", 10.0, false, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("Куда передено", "R7", "R9", 17.0, true, 0, 'N', 'N', 0, xlS);
            xl.CellMerge("Примечания", "S5", "S9", 20.0, false, 0, 'N', 'N', 0, xlS);
        }

        private void ExcelData(OdbcDataReader reader, Excel.Worksheet xlS)
        {
            LibToExcel xl = new LibToExcel();

            try
            {
                int topRow = 10;
                DataTable dt = new DataTable();

                // Выгружаем DataReader в таблицу dt
                dt.Load(reader);

                // Создаём двухмерный массив и загружаем в него таблицу
                //https://qarchive.ru/88699_zapis__massiva_v_diapazon_excel

                object[,] arr = new object[dt.Rows.Count, dt.Columns.Count];
                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    DataRow dr = dt.Rows[r];
                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        arr[r, c] = dr[c];
                    }
                }

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

                // форматирование столбцов для вывода данных
                xl.ColumnFormat(2, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "@", xlS);   //инв.№
                xl.ColumnFormat(3, topRow, topRow + dt.Rows.Count - 1, true, 10, 'L', "@", xlS);   //модель
                xl.ColumnFormat(4, topRow, topRow + dt.Rows.Count - 1, true, 0, 'C', "@", xlS);    //гос.№
                xl.ColumnFormat(5, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "0000", xlS);    //год
                xl.ColumnFormat(6, topRow, topRow + dt.Rows.Count - 1, false, 0, 'L', "@", xlS);    //VIN
                xl.ColumnFormat(7, topRow, topRow + dt.Rows.Count - 1, false, 10, 'L', "@", xlS);    //Кузов
                xl.ColumnFormat(8, topRow, topRow + dt.Rows.Count - 1, false, 10, 'L', "@", xlS);    //Шасси
                xl.ColumnFormat(9, topRow, topRow + dt.Rows.Count - 1, false, 10, 'L', "@", xlS);    //Двиг.
                xl.ColumnFormat(10, topRow, topRow + dt.Rows.Count - 1, true, 9, 'C', "@", xlS);    //ПТС
                xl.ColumnFormat(11, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "dd/mm/yyyy", xlS);    //дата поступления
                xl.ColumnFormat(12, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "@", xlS);    //бюджет
                xl.ColumnFormat(13, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "dd/mm/yy", xlS);    //дата приказ ввода в эксп.
                xl.ColumnFormat(14, topRow, topRow + dt.Rows.Count - 1, true, 0, 'L', "@", xlS);    //приказ ввода в эксп.
                xl.ColumnFormat(15, topRow, topRow + dt.Rows.Count - 1, true, 10, 'L', "@", xlS);    //закреплен
                xl.ColumnFormat(16, topRow, topRow + dt.Rows.Count - 1, false, 0, 'C', "dd/mm/yy", xlS);    //дата приказ ввода в эксп.
                xl.ColumnFormat(17, topRow, topRow + dt.Rows.Count - 1, true, 0, 'L', "@", xlS);    //приказ списания (передачи)
                xl.ColumnFormat(18, topRow, topRow + dt.Rows.Count - 1, true, 10, 'L', "@", xlS);    //куда
                xl.ColumnFormat(19, topRow, topRow + dt.Rows.Count - 1, true, 10, 'L', "@", xlS);    //закреплен

                // Определяем диапазон таблицы в который зальём массив
                Excel.Range c1 = (Excel.Range)xlS.Cells[topRow, 1];
                Excel.Range c2 = (Excel.Range)xlS.Cells[topRow + dt.Rows.Count - 1, dt.Columns.Count];
                Excel.Range range = xlS.get_Range(c1, c2);
                range.Value = arr;
                
                range.VerticalAlignment = Excel.Constants.xlCenter;

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
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
