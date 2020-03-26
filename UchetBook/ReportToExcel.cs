using System;
using System.Data;
using System.Data.Odbc;
using System.IO;

// Добавим COM MS Office 15.0 Object Library и MS Ecxel 15.0 Object Library
using Excel = Microsoft.Office.Interop.Excel;

using Microsoft.Office.Interop.Excel;
using Microsoft.Extensions.Configuration;
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
        public Application? xlApp;                    //Екземпляр приложения Excel.Application
        Worksheet? xlSheet;                           //Лист Excel.Worksheet
       //Excel.Range xlSheetRange;                   //Выделеная область

        public OdbcDataReader? Reader { get; set; }
        public System.Data.DataTable? dt { get; set; }

        public ReportToExcel() : this (null, null) 
        {
        }

        public ReportToExcel(System.Data.DataTable? table, string? tDir)
        {
            dt = table;
            xlApp = new Application();

            // создаем отчет программно
            if  (tDir == null)                         //if (tDir == "")
            {
                //добавляем книгу
                xlApp.Workbooks.Add(Type.Missing);
                if (xlApp.Sheets.Count == 1)        //для office 13
                {
                    xlApp.Sheets.Add();
                }
            }
            // создаем отчет на основе шаблона
            else
            {
                // файл конфигурации
                IConfigurationRoot configuration = new ConfigurationBuilder()
                .AddJsonFile("config.json", optional: true)
                .Build();

                // анонимный тип
                var config = new
                {
                    Dir = configuration["Templates:Path"],
                    File = configuration["Templates:UchetBookFile"]
                };

                xlApp.Workbooks.Open(Path.GetFullPath(Path.Combine(tDir, config.Dir, config.File)));
            }
            xlApp.EnableEvents = false;     //отключить события в excel
        }

        // работаем с Excel
        public void ExelObjecCars(string zagolovok)
        {
            try
            {
                //I.Список ТС
                if (xlApp != null)
                {
                    //выбираем лист на котором будем работать (Лист 1)
                    xlSheet = (Worksheet)xlApp.Sheets[1];
                    //Название листа
                    xlSheet.Name = "Список ТС";
                    //цвет вкладки
                    xlSheet.Tab.Color = 255;

                    // названия столбцов таблицы
                    Shapka(zagolovok, xlSheet);
                    // вывод данных
                    ExcelData(xlSheet);
                    // итоги
                    //Podval();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                if (xlApp != null)
                {
                    //Показываем ексель
                    xlApp.Visible = true;

                    xlApp.Interactive = true;
                    xlApp.ScreenUpdating = true;
                    xlApp.UserControl = true;

                    //Отсоединяемся от Excel
                    //releaseObject(xlSheetRange);
                    if (xlSheet != null)
                    {
                        releaseObject(xlSheet);
                    }
                    releaseObject(xlApp);
                }
            }
        }

        // выгружаем Excel из памяти
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine(ex.ToString(), "Ошибка!");
            }
            finally
            {
                GC.Collect();
            }
        }

        //создание заголовка таблицы
        private void Shapka(string zglvk, Worksheet xlS)
        {
            Excel.Range xlSheetRange;               //Выделеная область Excel.Range

            LibToExcel xl = new LibToExcel(xlS);

            // Заголовог таблицы
            xl.CellMerge(title: zglvk, cell1: "A3", cell2: "S3", tWidth: 0, wrpText: false,
                        tFont: 14, tHor: 'C', tVer: 'C', tOrient: 0);

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
            xl.CellMerge("№ п/п", "A5", "A9", 5.14, true, 0, 'N', 'N', 0);
            xl.CellMerge("Инв. №", "B5", "B9", 11.57, false, 0, 'N', 'N', 0);
            xl.CellMerge("Марка, модель ТС", "C5", "C9", 16.0, true, 0, 'N', 'N', 0);
            xl.CellMerge("Гос. №", "D5", "D9", 12.14, false, 0, 'N', 'N', 0);
            xl.CellMerge("Год выпуска", "E5", "E9", 5.0, true, 10, 'N', 'N', 90);
            xl.CellMerge("VIN", "F5", "F9", 20.0, false, 0, 'N', 'N', 0);
            xl.CellMerge("№ кузова", "G5", "G9", 18.57, false, 0, 'N', 'N', 0);
            xl.CellMerge("№ шасси", "H5", "H9", 18.57, false, 0, 'N', 'N', 0);
            xl.CellMerge("№ двигателя", "I5", "I9", 18.57, false, 0, 'N', 'N', 0);
            xl.CellMerge("№ ПТС", "J5", "J9", 13.3, false, 0, 'N', 'N', 0);
            xl.CellMerge("Сведения о поступлении ТС", "K5", "L6", 0, true, 0, 'N', 'N', 0);
            xl.CellMerge("Дата", "K7", "K9", 9.86, false, 0, 'N', 'N', 0);
            xl.CellMerge("Источник приобре-тения", "L7", "L9", 0, true, 10, 'N', 'N', 0);
            xl.CellMerge("Приказ ввода в эксплуатацию", "M5", "N6", 13.71, true, 0, 'N', 'N', 0);
            xl.CellMerge("Дата", "M7", "M9", 9.86, false, 0, 'N', 'N', 0);
            xl.CellMerge("Номер", "N7", "N9", 10.0, false, 0, 'N', 'N', 0);
            xl.CellMerge("Орган (служба) за которым закреплено ТС", "O5", "O9", 17.57, true, 0, 'N', 'N', 0);
            xl.CellMerge("Сведения о передаче или списании ТС", "P5", "R6", 0, true, 0, 'N', 'N', 0);
            xl.CellMerge("Приказ", "P7", "Q7", 0, false, 0, 'N', 'N', 0);
            xl.CellMerge("Дата", "P8", "P9", 9.86, false, 0, 'N', 'N', 0);
            xl.CellMerge("Номер", "Q8", "Q9", 10.0, false, 0, 'N', 'N', 0);
            xl.CellMerge("Куда передено", "R7", "R9", 17.0, true, 0, 'N', 'N', 0);
            xl.CellMerge("Примечания", "S5", "S9", 20.0, false, 0, 'N', 'N', 0);
        }

        // загрузка данных
        private void ExcelData(Worksheet xlS)
        {
            LibToExcel xl = new LibToExcel(xlS);

            try
            {
                int topRow = 10;
                int bottomRow;
                int nRw;                // количество строк в таблице

                //https://qarchive.ru/88699_zapis__massiva_v_diapazon_excel
                // Создаём двухмерный массив и загружаем в него таблицу, 
                // чтобы одним махом залить его в Ecxel

                object[,] arr = new object[dt.Rows.Count, dt.Columns.Count];
                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    DataRow dr = dt.Rows[r];
                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        arr[r, c] = dr[c];
                    }
                }

                nRw = dt.Rows.Count - 1;     // количество строк
                bottomRow = topRow + nRw;

                // форматирование столбцов для вывода данных
                xl.ColumnFormat(2, topRow, bottomRow, false, 0, 'C', "@");   //инв.№
                xl.ColumnFormat(3, topRow, bottomRow, true, 10, 'L', "@");   //модель
                xl.ColumnFormat(4, topRow, bottomRow, true, 0, 'C', "@");    //гос.№
                xl.ColumnFormat(5, topRow, bottomRow, false, 0, 'C', "0000");    //год
                xl.ColumnFormat(6, topRow, bottomRow, false, 0, 'L', "@");    //VIN

                xl.RegionFormat(7, 9, topRow, bottomRow, false, 10, 'L', "@");    //Кузов-Шасси-Двиг.

                //xl.ColumnFormat(7, topRow, topRow + dt.Rows.Count - 1, false, 10, 'L', "@", xlS);    //Кузов
                //xl.ColumnFormat(8, topRow, topRow + dt.Rows.Count - 1, false, 10, 'L', "@", xlS);    //Шасси
                //xl.ColumnFormat(9, topRow, topRow + dt.Rows.Count - 1, false, 10, 'L', "@", xlS);    //Двиг.

                xl.ColumnFormat(10, topRow, bottomRow, true, 9, 'C', "@");    //ПТС
                xl.ColumnFormat(11, topRow, bottomRow, false, 0, 'C', "dd/mm/yyyy");    //дата поступления
                xl.ColumnFormat(12, topRow, bottomRow, false, 0, 'C', "@");    //бюджет
                xl.ColumnFormat(13, topRow, bottomRow, false, 0, 'C', "dd/mm/yy");    //дата приказ ввода в эксп.
                xl.ColumnFormat(14, topRow, bottomRow, true, 0, 'L', "@");    //приказ ввода в эксп.
                xl.ColumnFormat(15, topRow, bottomRow, true, 10, 'L', "@");    //закреплен
                xl.ColumnFormat(16, topRow, bottomRow, false, 0, 'C', "dd/mm/yy");    //дата приказ ввода в эксп.
                xl.ColumnFormat(17, topRow, bottomRow, true, 0, 'L', "@");    //приказ списания (передачи)

                xl.RegionFormat(18, 19, topRow, bottomRow, true, 10, 'L', "@");    //куда-примечания

                //xl.ColumnFormat(18, topRow, topRow + dt.Rows.Count - 1, true, 10, 'L', "@", xlS);    //куда
                //xl.ColumnFormat(19, topRow, topRow + dt.Rows.Count - 1, true, 10, 'L', "@", xlS);    //примечания

                // Определяем диапазон таблицы в который зальём массив
                Excel.Range c1 = (Excel.Range)xlS.Cells[topRow, 1];
                Excel.Range c2 = (Excel.Range)xlS.Cells[bottomRow, dt.Columns.Count];
                Excel.Range range = xlS.get_Range(c1, c2);
                // выгружаем таблицу
                range.Value = arr;

                range.VerticalAlignment = Excel.Constants.xlCenter;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        // итоги таблицы
        private void Podval(Worksheet xlS)
        {

        }
    }
}
