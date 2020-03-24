using System;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Data;
//using System.Drawing;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Windows.Forms;

//using Excel = Microsoft.Office.Interop.Excel;
//using Word = Microsoft.Office.Interop.Word;

//namespace AppWordExcel
//{
//    public partial class Form1 : Form
//    {
//        private Excel.Application excelapp;         //Определим глобально основной объект Excel.Application 
//        private Excel.Window excelWindow;           //объект Excel.Window

//        private Excel.Workbooks excelappworkbooks;  //ссылки на созданные книги 
//        private Excel.Workbook excelappworkbook;    //ссылка на объект - конкретную книгу 

//        private Excel.Sheets excelsheets;           //коллекции листов 
//        private Excel.Worksheet excelworksheet;     //объекты лист

//        // определения для ячеек и ячейки мы задать не можем, 
//        //так как отдельно данные объекты как самостоятельные в C# отсутствуют, 
//        //а есть понятие области выделенных ячеек, которая может включать одну или более ячеек, 
//        //с которыми можно выполнять действия
//        private Excel.Range excelcells;

//        public Form1()
//        {
//            InitializeComponent();
//        }

//        private void button1_Click(object sender, EventArgs e)
//        {
//            int i = Convert.ToInt32(((Button)(sender)).Tag);
//            switch (i)
//            {
//                case 1:
//                    //выполнять запуск Excel 
//                    excelapp = new Excel.Application(); 
//                    excelapp.Visible=true;

//                    //Свойство SheetsInNewWorkbook возвращает или устанавливает количество листов, 
//                    //автоматически помещаемых Excel в новые рабочие книги
//                    excelapp.SheetsInNewWorkbook=3;
//                    //Type.Missing - отсутствие значения. 
//                    //Некоторые методы Excel принимают необязательные параметры, которые не поддерживаются в C#. 
//                    //Для решения этой проблемы в коде на C# требуется передавать поле Type.Missing 
//                    //вместо каждого необязательного параметра, который является ссылочным типом (reference type). 
//                    excelapp.Workbooks.Add(Type.Missing);

//                    excelapp.SheetsInNewWorkbook=5;
//                    excelapp.Workbooks.Add(Type.Missing);

//                    //Получаем массив ссылок на листы выбранной книги
//                    excelsheets = excelappworkbook.Worksheets;

//                    //Выбираем лист 3
//                    excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
//                    //Делаем третий лист активным
//                    excelworksheet.Activate();
//                    //Вывод в ячейки используя номер строки и столбца Cells[строка, столбец]
//                    for (int m = 1; m < 20; m++)
//                    {
//                        for (int n = 1; n < 15; n++)
//                        {
//                            excelcells = (Excel.Range)excelworksheet.Cells[m, n];
//                            //Выводим координаты ячеек
//                            excelcells.Value2 = m.ToString() + " " + n.ToString();
//                        }
//                    }



//                    break;
//                case 2:
//                    //закрывают конкретную рабочую книгу: 
//                    excelapp.Windows[1].Close(false, Type.Missing, Type.Missing);

//                    //закрывают все книги
//                    excelapp.Workbooks.Close();
     
//                    //закрытие Excel
//                    excelapp.Quit();
//                    break;
//                default:
//                    Close();
//                    break;
//            }
//        }

//        private void button2_Click(object sender, EventArgs e)
//        {
//            button2.Click += new EventHandler(button1_Click);
//        }

//        private void button3_Click(object sender, EventArgs e)
//        {
//            button3.Click += new EventHandler(button1_Click);
//        }
//    }
//}
