using System;
using System.IO;

namespace Vng.Uchet
{
    class Program
    {
        static void Main(string[] args)
        {
            string? projectDir;

            Console.WriteLine("Укажите какие отчеты следует сформировать:");
            Console.WriteLine("\t  Книга учета (777) - 1:");
            Console.WriteLine("\t\t     Отмена - 0:");
            Console.Write("Укажите цифру и нажмите Enter: ");
            string selectReport = Console.ReadLine();

            Console.WriteLine("\nСоздать отчет на основе существующего шаблона (Y/N)?");
            Console.Write("Укажите Y(да) или N(нет) и нажмите Enter: ");
            string yesNo = Console.ReadLine();

            switch (selectReport) 
            {
                case "1":
                    if (yesNo == "y" || yesNo == "н" || yesNo == "Y" || yesNo == "Н")
                    {
                        projectDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\.."));
                        // в рабочем варианте
                        //projectDir = Environment.CurrentDirectory + @"\Templates\UchetBook.xltx";
                    }
                    else
                    {
                        //projectDir = string.Empty;      //projectDir = ""
                        projectDir = default;           // или default(string); -> projectDir = null
                    }
                    // создаем объект для подключения к БД и загрузки книги учета
                    OdbcData oDbcUb = new OdbcData("UB");                         
                    // Выгружаем reader в таблицу DataTable
                    var xlS = new ReportToExcel(oDbcUb.LoadData(), projectDir);
                    // Выгружаем в Excel
                    xlS.ExelObjecCars("Книга учета ТС ФГКУ «УВО ВНГ России по городу Москве»");
                    break;
                case "2":
                    Console.WriteLine("Функция пока отсутствует");
                    break;
                case "3":
                    Console.WriteLine("Функция пока отсутствует");
                    break;
                case "0":
                    break;
                default:
                    break;
            }
            //Console.ReadLine();
        }
    }
 }
