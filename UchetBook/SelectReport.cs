using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;

using static System.Console;

namespace Vng.Uchet
{
    //класс выбора требуемого отчета
    class SelectReport
    {
        public readonly static string ProjectPath;

        // статический конструктор, здесь определяем ProjectPath
        static SelectReport()
        {
            ProjectPath = string.Empty;

            // определяем каталог откуда запущено приложение, в нём будет располагаться папка
            // с файлами шаблонов Excel (настройки в config.json в разделе "Templates":)
            if (AppDomain.CurrentDomain.BaseDirectory != null)
            {
                string? path = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory);
                ProjectPath = path != null ? path : string.Empty;
            }
            if (ProjectPath == string.Empty)
            {
                WriteLine("Ошибка определения BaseDirectory.");
            }
        }

        // выбор отчета
        public static bool SelectOfReport(bool first)
        {
            byte select = byte.MaxValue;

            if (first)
            {
                WriteLine("Укажите какие отчеты следует сформировать:");
                WriteLine("\t\t\t\t\t  I.   Книга учета (777) - 1:");
                WriteLine("\t\t\t\t\t  II.  Характеристики ТС - 2:");
                WriteLine("\t\t\t\t\t  III. Отчёты  I. и II.  - 3:");
                WriteLine("\t\t\t\t\t  IV.  Пятничная справка - 4:");
                WriteLine("\t\t\t\t\t  V.   Все отчёты        - 5:");
                WriteLine("\t\t\t\t\t\t\t  Отмена - 0:");
            }
            Write("Для формирования отчета введите цифру от 0 до 9 и нажмите <Enter>: ");

            // делаем выбор
            while (select == byte.MaxValue)
            {
                select = ParsingInpt(ReadLine());
            }

            RunReport(select);                      // запускаем формирование отчета
            return select != 0 ? true : false;      // выйти из программы или выбрать новый отчет        
        }

        // определяем, что выбрал пользователь
        private static byte ParsingInpt(string Report)
        {
            // ввели число от 0 до 9
            if (byte.TryParse(Report, out byte count))
            {
                if (count >= 0 & count <= 9) 
                {
                    return count;
                }
            }
            // повторить ввод
            WriteLine();
            Write("<<< Введите цифру от 0 до 9: ");
            return byte.MaxValue;
        }

        // запускаем выбранный отчет на формирование
        private async static void RunReport(byte report)
        {
            string? dir;

            switch (report)
            {
                case 1:                 // сделан выбор - 1. это Книга учета (777)                    
                    WriteLine();
                    WriteLine("Началось формирование отчета - \"Книга учета\"");

                    #region Task
                    // (Task<bool>)qqq = (Task<bool>)true;                      
                    //// Console.WriteLine($"qqq1 = {(Task<bool>)qqq}");
                    //Task.Delay(100);
                    //Thread.Sleep(8000);
                    //Task<bool> qqq = doProg.DoExcel(projectDir);
                    #endregion

                    //определяем как строить отчет, на основе шаблона или программно                
                    dir = ReportOnTemlate("UB") == true ? ProjectPath : null;
                    Program doProg = new Program();
                    await doProg.DoExcelAsync(dir, "UB");

                    #region async
                    //// bool aaa = qqq.Result;
                    //// Console.WriteLine($"aaa = {aaa}");
                    //// Wait(qqq);
                    //do
                    //{
                    //    Console.WriteLine($"aaa1 = {aaa}");
                    //    //aaa = qqq.Result;
                    //    //Console.WriteLine($"aaa2 = {aaa}");
                    //} while (qqq.Result == false);

                    ////bool eee = false;
                    ////Console.WriteLine($"eee1 = {eee}");

                    ////eee = await doProg.DoExcel(projectDir);
                    //////Console.WriteLine($"qqq2 = {qqq ?? null}");

                    //Console.WriteLine($"eee2 = {eee}");

                    //if (qqq == 0)

                    //    if ((Task<bool>)qqq != null )
                    //{
                    //    Console.WriteLine("qqq1 = NULL");
                    //}
                    //else
                    //{
                    //Console.WriteLine("Идёт загрузка в Excel: 111111");
                    //}

                    //// создаем объект для подключения к БД и загрузки книги учета (UB)
                    //OdbcData oDbcUb = new OdbcData("UB");                         
                    //// Выгружаем reader в таблицу DataTable
                    //var xlS = new ReportToExcel(oDbcUb.LoadData(), projectDir);
                    //// Выгружаем в Excel
                    ////await Task.Run(() => xlS.ExelObjecCars("Книга учета ТС ФГКУ «УВО ВНГ России по городу Москве»"));
                    //xlS.ExelObjecCars("Книга учета ТС ФГКУ «УВО ВНГ России по городу Москве»");
                    #endregion

                    break;
                case 2:                 // сделан выбор 2. это Характеристики ТС          
                    WriteLine("Функция пока отсутствует");
                    break;
                case 3:                 // сделан выбор 3. 
                    WriteLine("Функция пока отсутствует");
                    break;
                case 4:                 // сделан выбор 4. Пятничная справка
                    WriteLine("Функция пока отсутствует");
                    break;
                case 5:                 // сделан выбор 5. 
                    WriteLine("Функция пока отсутствует");
                    break;
                case 0:                 // отмена
                    WriteLine("Выход из программы");
                    break;
                default:
                    break;
            }
        }

        // из файла конфигурации определяем как строить отчет, на основе шаблона
        // или программно для этого считываем значение из раздела "OnTemplate" 
        public static bool ReportOnTemlate(string report)
        {
            bool flag = default;

            // файл конфигурации
            IConfigurationRoot configuration = new ConfigurationBuilder()
            .AddJsonFile("config.json", optional: true)
            .Build();

            var config = new            // анонимный тип
            {
                ub = configuration["OnTemplate:BookUchet"],         //книга учета
                ts = configuration["OnTemplate:Specifications"],    //спецификация
                ps = configuration["OnTemplate:FridayReport"]       //FridayReport
            };

            switch (report)
            {
                case "UB":
                    flag = config.ub == "Yes" ? true : false;       //Yes - на основе шаблона
                    break;                                          //No - программно
                case "TS":
                    flag = config.ts == "Yes" ? true : false;
                    break;
                case "PS":
                    flag = config.ps == "Yes" ? true : false;
                    break;
                default:
                    break;
            }
            return flag;
        }
    }
}
