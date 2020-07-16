using System;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;

using static System.Console;
using static Vng.Uchet.SelectReport;

namespace Vng.Uchet
{
    class Program
    {       
        //public static string ProjectPath = string.Empty;      

        static void Main(string[] args)             //static async void Main(string[] args)
        {            
            if (ProjectPath == string.Empty) {      // Vng.Uchet.SelectReport.ProjectPath
                return;                             // ошибка приложения выход    
            }        
            bool repeat = true;                     // можно формировать несколько отчетов подряд
            bool firstSelect = true;                // еще не был выбран ни один отчет

            while (repeat)                          
            {
                repeat = SelectOfReport(firstSelect);   // выбор отчета
                firstSelect = false;
            }
            ReadLine();
        }

        // Вызов DoExcelAsync с использованием Task

        //public async Task<bool> DoExcelAsync(string? tDir)
        //async Task<bool> DoExcelAsync(string? tDir, CancellationToken token)

        public async Task DoExcelAsync(string? tDir, string tTmpl)    //если tDir = null строим отчет программно
        {          
            var result = false;

            // создаем объект для подключения к БД и загрузки книги учета (UB)
            OdbcData oDbcUb = new OdbcData(tTmpl);      
            // Выгружаем reader в таблицу DataTable
            var xlS = new ReportToExcel(oDbcUb.LoadData(), tDir);

            Task generateResultTask = Task.Run(() => xlS.ExelObjecCars("Книга учета ТС ФГКУ «УВО ВНГ России по городу Москве»"));
            //Task<bool> generateResultTask = Task.Run(() => xlS.ExelObjecCars("Книга учета ТС ФГКУ «УВО ВНГ России по городу Москве»"));
            //await generateResultTask.ConfigureAwait(false);
            
            //var allTasks = new List<Task> { generateResultTask, baconTask, toastTask };
            var allTasks = new List<Task> { generateResultTask };

            while (allTasks.Any())
            {
                Task finished = await Task.WhenAny(allTasks);
                if (finished == generateResultTask)
                {
                    WriteLine("eggs are ready");
                }
                //else if (finished == baconTask)
                //{
                //    Console.WriteLine("bacon is ready");
                //}
                //else if (finished == toastTask)
                //{
                //    Console.WriteLine("toast is ready");
                //}
                allTasks.Remove(finished);
            }
            //await generateResultTask;

            //// Выгружаем в Excel
            ////await Task.Run(() => xlS.ExelObjecCars("Книга учета ТС ФГКУ «УВО ВНГ России по городу Москве»"));
            ////bFlag = result;
            ////result = await Task.Run(() => xlS.ExelObjecCars("Книга учета ТС ФГКУ «УВО ВНГ России по городу Москве»"));
            ////bFlag = result;
            ////return await Task.Run(() => xlS.ExelObjecCars("Книга учета ТС ФГКУ «УВО ВНГ России по городу Москве»"));
            ////bFlag = result;
            ////return bFlag;
            ////Console.WriteLine($"result = {result}");
            ////return result;

            WriteLine($"Факториал равен {result}");

            //return generateResultTask.Result;
            //return generateResultTask;
        }
    }
 }
