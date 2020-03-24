using System;
//using System.Configuration;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Data;

//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Windows.Forms;
//using System.Data.SqlClient;

//using System.Threading;
//using System.Net;
//using System.IO;
//using System.Net.Http;
//using System.Diagnostics;

//using AviaLes;

//using Excel = Microsoft.Office.Interop.Excel;

//using System.Drawing;


////Если учитывать, что все чекбоксы называются "checkBox*", где "*" - любое кол-во цифр, 
////то их порядковый номер можно парсить из свойства Name:Код C#
////if(checkBox.Checked)//Если флажок установлен
////output += checkBox.Name.Substring(8) + " & "; //Т.к. в слове checkBox 8 букв

////2
////Вообще, как бы правильно создавать не checkBox1, checkBox2, checkBox3 и т.д., 
////а массив этих CheckBox:Код C#
////CheckBox[] array = new CheckBox[6];

////Тогда задача решается в пару строчек:Код C#
////	string myString = string.Join(" & ", array.Select((x, i) => x.Checked ? i + 1 : 0).Where(x => x != 0));
////if (string.IsNullOrEmpty(myString))
////myString = "0";

///////////////////////////////////////  ASync
//namespace AviaLesMeteo
//{
//    public partial class Form1 : Form
//    {
//        // Declare a System.Threading.CancellationTokenSource.
//        //добавим класс источников признаков отмены для асинхронных методов
//        CancellationTokenSource cts;

//        //список скачанных страниц
//        List<SiteHTML> HTMLs3 = new List<SiteHTML>();
//        List<SiteHTML> HTMLs9 = new List<SiteHTML>();

//        List<MeteoSt> Meteos = new List<MeteoSt>();

//        DateTime dtSelect;
//        ServerSelect avialesServer;

//        static Subjects[] Subj;                             //массив кодов (id) и названий субъектов

//        Excel.Application xlApp;                            //Екземпляр приложения Excel
//        Excel.Worksheet xlSheet;                            //Лист
//        Excel.Range xlSheetRange;                           //Выделеная область

//        public Form1()
//        {
//            InitializeComponent();
//        }

//        private void Form1_Load(object sender, EventArgs e)
//        {
//            //Создаю DataSet 
//            //В дальнейшем в него будут загружены данные из базы данных
//            //А за тем выгружены в таблицу Excel
//            CreateDS();
//       } 

//        private void buttonCancel_Click(object sender, EventArgs e)
//        {
//            //Решили прерватьзагрузку данных
//            if (cts != null) cts.Cancel();
//        }

//        private void buttonClose_Click(object sender, EventArgs e)
//        {
//                this.Close();
//        }

//        private async void buttonOk_Click(object sender, EventArgs e)
//        {
//            cts = new CancellationTokenSource();        //добавим возможность отмены загрузки

//            if (AnalizChkBox(out Subj))                 //если выбрали субъекты, сохраним их в массиве Subj
//                PrepareComponents();
//            else
//                return;

//            //на какую дату загружать
//            dtSelect = this.dateTimePicker.Value;       //dt = dateTimePicker.Value.ToString("yyyy-MM-dd");                
//////            int old = (dtSelect.Year < 2007 ? 1 : 0);

//            try
//            {
//                //выбираем предпочитаемый сервер загрузки
//                string serverName = this.comboBoxServer.Text;

//                textBox2.Text = string.Format("\r\n Идёт инициализация на сервере Авиалесоохраны");
//                textBox2.Refresh();

//                //создаём экземпляр класса ServerSelect сервера загрузки               
//                avialesServer = new ServerSelect(serverName);
//                //регистрация в сети
//                await avialesServer.Init();

//                ////Оператывный дежурный
//                //for (int i = 0; i < Subj.Length; i++)           //цикл по всем выбранным субъектам
//                //{
//                //    //загружаем страницу со списком метеостанций на дату dtSelect для субъекта Subj[i]
//                //    await ReadSubjects(avialesServer, dtSelect, Subj[i]);
//                //    //загружаем страницу метеостанции на дату dtSelect для субъектов из массива Subj
//                //    await ReadMeteosMS(avialesServer, dtSelect, Subj);
//                //}

//                //НИОКР
//                await ReadSubjectsNiokr(avialesServer, dtSelect, Subj);

                
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show(ex.ToString());
//            }

//            textBox1.Text = String.Format("\r\nВыгрузка в таблицу EXCEL");
//            textBox2.Clear();

//            ////Оператывный дежурный
//            //await Task.Run(() => ExelObject3(Program.tbl, Program.tblM));

//            //НИОКР
//            await Task.Run(() => ExelObjectNiokr(Program.tblP));

//            this.Close();
//        }

////Оператывный дежурный---ППО-----------------------------------------------------------------------
//        private async Task ReadSubjects(ServerSelect avialesServer, DateTime pdtSelect, Subjects pSubjVal)
//        {
//            Task sumTask = PPOPageAsyncCancel(avialesServer, pdtSelect, pSubjVal, cts.Token);

//            ////Do some other stuff...
//            ////....
//            ////Wait until sumTask is Finished...
//            await sumTask;

//            ParseSubject(PriznakYear(pdtSelect));

//            this.textBox3.Refresh();
//        }

//        private async Task PPOPageAsyncCancel(ServerSelect avialesServer, DateTime pdtSelect,
//                                                Subjects psubjVal, CancellationToken ct)
//        {
//            Uri uri = new Uri("http://www." + avialesServer.name + ".aviales.ru" + "/rus/main.sht");
//            var cc = new CookieContainer();
//            cc.SetCookies(uri, avialesServer.cookies);
//            var handler = new HttpClientHandler { CookieContainer = cc };

//            var client = new HttpClient(handler);
//            //получаем список URL-адресов субъектов
//            IEnumerable<SiteURL> urlList = LoadSubjectsURL(avialesServer, pdtSelect, psubjVal);

//            string HTML = string.Empty;

//            foreach (var url in urlList)
//            {
//                HttpResponseMessage response = await client.GetAsync(url.url, ct);
//                byte[] urlContents = await response.Content.ReadAsByteArrayAsync();

//                HTML = Encoding.GetEncoding(1251).GetString(urlContents);

//                HTMLs3.Add(new SiteHTML(HTML, url.reg, url.regNaim, url.dt,
//                                        url.meteoIndex, url.meteoNaim, url.OkrugId, url.gmt, url.ppo));
//                textBox3.Text =
//                    string.Format("\r\n{0} : {1}", url.reg, url.regNaim) + textBox3.Text;
//                label3.Text = "Метеостанции субъекта \"" + url.regNaim + "\"";

//                //textBox2.Text += string.Format("\r\n {0}", url.url);
//            }
//        }

//        public IEnumerable<SiteURL> LoadSubjectsURL(ServerSelect avialesServer,
//                                                    DateTime pdtSelect, Subjects psubjVal)
//        {
//            string meteoDate = pdtSelect.ToString("yyyy-MM-dd");

//            var urls = new List<SiteURL>();

//            urls.Add(new SiteURL(avialesServer.adrServer + "/secure/meteo/meteo_1.sht?reg_id="
//                    + psubjVal.Id + "&date=" + meteoDate + "&old=" + PriznakYear(pdtSelect),
//                    psubjVal.Id, psubjVal.Name, psubjVal.OkrugId, meteoDate, string.Empty, string.Empty, string.Empty));

//            //возвращаем ссылку на коллекцию
//            return urls;
//        }

//        private void ParseSubject(string pOld)
//        {
//            Meteos.Clear();
//            while (true)
//            {
//                if (HTMLs3.Count == 0)
//                    break;
//                else
//                    if (HTMLs3.Count > 0)
//                    {
//                        AviaLesParsePPO pp = new AviaLesParsePPO();
//                        //парсим страницу и загружаем результат в массив ppoMeteoSt
//                        string[,] ppoMeteoSt = pp.ParseSubj(HTMLs3[0].html, false);

//                        //добавляем в список метеостанций
//                        for (int i = 0; i < ppoMeteoSt.GetLength(0); i++)
//                            Meteos.Add(
//                                new MeteoSt(avialesServer.adrServer + "/secure/meteo/"
//                                                        + ppoMeteoSt[i, 0] + "&old=" + pOld,
//                                            ppoMeteoSt[i, 1], ppoMeteoSt[i, 2], HTMLs3[0].reg,
//                                            HTMLs3[0].regNaim, ppoMeteoSt[i, 17]));

//                        HTMLs3.RemoveAt(0);
//                    }
//            }

//            textBox1.Text = string.Empty;
//            textBox2.Text = string.Empty;

//            foreach (var url in Meteos)
//                textBox1.Text += string.Format("\r\n {0}: \t{1}", url.nomer, url.naim);

//            this.textBox1.Refresh();
//        }

////НИОКР--------------------------------------------------------------------------
//        private async Task ReadSubjectsNiokr(ServerSelect avialesServer, DateTime pdtSelect, Subjects[] pSubjArr)
//        {
//            Task sumTask  = PPOPageNiokrAsyncCancel(avialesServer, pdtSelect, pSubjArr, cts.Token);
//            await sumTask;

//            ParseSubjectNiokr(PriznakYear(pdtSelect));

//            this.textBox2.Refresh();
//            this.textBox3.Refresh();
//        }

//        private async Task PPOPageNiokrAsyncCancel(ServerSelect avialesServer, DateTime pdtSelect,
//                                                    Subjects[] psubjArr, CancellationToken ct)
//        {
//            Uri uri = new Uri("http://www." + avialesServer.name + ".aviales.ru" + "/rus/main.sht");
//            var cc = new CookieContainer();
//            cc.SetCookies(uri, avialesServer.cookies);
//            var handler = new HttpClientHandler { CookieContainer = cc };

//            var client = new HttpClient(handler);
//            //получаем список URL-адресов субъектов
//            IEnumerable<SiteURL> urlList = LoadSubjectsURLNiokr(avialesServer, pdtSelect, psubjArr);

//            IEnumerable<Task<SiteURL>> downloadTasksQuery =
//                         from url in urlList select ProcessURLPPO(url, client, ct);

//            List<Task<SiteURL>> downloadTasks = downloadTasksQuery.ToList();

//            while (downloadTasks.Count > 0)
//            {
//                Task<SiteURL> firstFinishedTask = await Task.WhenAny(downloadTasks);

//                downloadTasks.Remove(firstFinishedTask);

//                SiteURL ur = firstFinishedTask.Result;

//                textBox3.Text = String.Format("\r\n {0}: {1}", ur.reg, ur.regNaim) + textBox3.Text;

//                //// 1 или 0               textBox3.Text += string.Format("\r\n - {0}", HTMLs9.Count);

//////                await Task.Run(() => ParseSubjectNiokr(PriznakYear(pdtSelect)));
//            }
//            textBox2.Text += String.Format("\r\n\r\n ПАРСИНГ ЗАГРУЖЕННЫХ СТРАНИЦ");

//            //    HTMLs3.Add(new SiteHTML(HTML, url.reg, url.regNaim, url.dt,
//            //                            url.meteoIndex, url.meteoNaim, url.gmt, url.ppo));
//            //    textBox3.Text =
//            //        string.Format("\r\n{0} : {1}", url.reg, url.regNaim) + textBox3.Text;
//            //    label3.Text = "Метеостанции субъекта \"" + url.regNaim + "\"";

//            //    //textBox2.Text += string.Format("\r\n {0}", url.url);
//            //}
//        }

//        async Task<SiteURL> ProcessURLPPO(SiteURL url, HttpClient client, CancellationToken ct)
//        {
//            string HTML1 = string.Empty;

//            HttpResponseMessage response = await client.GetAsync(url.url, ct);

//            var task = response.Content.ReadAsByteArrayAsync();
//            var taskDebug = task.ContinueWith((task1)
//                    => Debug.WriteLine("URL:{0}, Thread:{1}", url, Thread.CurrentThread.ManagedThreadId));
//            byte[] urlContents = await task;

//            HTML1 = Encoding.GetEncoding(1251).GetString(urlContents);

//            HTMLs3.Add(new SiteHTML(HTML1, url.reg, url.regNaim, url.OkrugId,url.dt,
//                                        url.meteoIndex, url.meteoNaim, url.gmt, url.ppo));

//            return url;
//        }

//        public IEnumerable<SiteURL> LoadSubjectsURLNiokr(ServerSelect avialesServer,
//                                                 DateTime pdtSelect, Subjects[] psubjArr)
//        {
//            string meteoDate = pdtSelect.ToString("yyyy-MM-dd");

//            var urls = new List<SiteURL>();

//            //for (int i = 0; i < Subj.Length; i++)             //цикл по всем выбранным субъектам
//            for (int i = 0; i < psubjArr.Length; i++)           //цикл по всем выбранным субъектам
//            {
//                //получаем список URL-адресов субъектов
//                urls.Add(new SiteURL(avialesServer.adrServer + "/secure/meteo/meteo_1.sht?reg_id="
//                        + psubjArr[i].Id + "&date=" + meteoDate + "&old=" + PriznakYear(pdtSelect),
//                        psubjArr[i].Id, psubjArr[i].Name, psubjArr[i].OkrugId, 
//                        meteoDate, string.Empty, string.Empty, string.Empty));
//            }
//            //возвращаем ссылку на коллекцию
//            return urls;
//        }

//        private void ParseSubjectNiokr(string pOld)
//        {
//            Meteos.Clear();
//            while (true)
//            {
//                if (HTMLs3.Count == 0)
//                    break;
//                else
//                    if (HTMLs3.Count > 0)
//                    {
//                        AviaLesParsePPO pp = new AviaLesParsePPO();
//                        //парсим страницу и загружаем результат в массив ppoMeteoSt
//                        string[,] ppoMeteoSt = pp.ParseSubj(HTMLs3[0].html, false);

//                        ////добавляем в список метеостанций
//                        //for (int i = 0; i < ppoMeteoSt.GetLength(0); i++)
//                        //    Meteos.Add(
//                        //        new MeteoSt(avialesServer.adrServer + "/secure/meteo/"
//                        //                                + ppoMeteoSt[i, 0] + "&old=" + pOld,
//                        //                    ppoMeteoSt[i, 1], ppoMeteoSt[i, 2], HTMLs3[0].reg,
//                        //                    HTMLs3[0].regNaim, ppoMeteoSt[i, 17]));

//                        string okr = "---";
//                        if (String.IsNullOrEmpty(ppoMeteoSt[0, 0]))
//                        {
//                            //нет данных
//                        }
//                        else
//                        {
//                            //иначе грузим всё
//                            //AviaLesMeteo.LoadPage.InsMeteoSt(ppoMeteoSt, ppo, HTMLs9[0].regNaim,
//                            //                HTMLs9[0].meteoIndex, HTMLs9[0].meteoNaim, HTMLs9[0].gmt,
//                            //                dtSelect, Program.tbl, Program.tblM);

//                            switch (HTMLs3[0].OkrugId)
//                            {
//                                case "1":   okr = "ДВФО";   break;
//                                case "2":   okr = "СФО";    break;
//                                case "3":   okr = "ПФО";    break;
//                                case "4":   okr = "ЦФО";    break;
//                                case "5":   okr = "СЗФО";   break;
//                                case "6":   okr = "ЮФО";    break;
//                                case "7":   okr = "СКФО";   break;
//                                case "8":   okr = "УФО";    break;
//                                case "9":   okr = "КФО";    break; 
//                                default:    okr = "...";    break;  
//                            }

//                            AviaLesMeteo.LoadPage.InsPPO(ppoMeteoSt, okr, HTMLs3[0].regNaim, "eee", "ttt",
//                                dtSelect, Program.tblP);
//                        }

//                        HTMLs3.RemoveAt(0);
//                    }
//            }

//            //textBox1.Text = string.Empty;
//            //textBox2.Text = string.Empty;

//            //foreach (var url in Meteos)
//            //    textBox1.Text += string.Format("\r\n {0}: \t{1}", url.nomer, url.naim);

//            //this.textBox1.Refresh();
//        }

////Оператывный дежурный---Метео-----------------------------------------------------------------------
//        private async Task ReadMeteosMS(ServerSelect avialesServer, DateTime pdtSelect, Subjects[] pSubj)
//        {
//            textBox2.Clear();
//            Task sumTask = MeteoPageAsyncCancelMS(avialesServer, pdtSelect, pSubj, cts.Token);
//            await sumTask;
//        }

//        private async Task MeteoPageAsyncCancelMS(ServerSelect avialesServer, DateTime pdtSelect,
//                                                    Subjects[] pSubj, CancellationToken ct)
//        {
//            Uri uri = new Uri("http://www." + avialesServer.name + ".aviales.ru" + "/rus/main.sht");
//            var cc = new CookieContainer();
//            cc.SetCookies(uri, avialesServer.cookies);
//            var handler = new HttpClientHandler { CookieContainer = cc };

//            var client = new HttpClient(handler);

//            IEnumerable<MeteoSt> urlListM = Meteos;             //получаем список URL-адресов

//            IEnumerable<Task<MeteoSt>> downloadTasksQuery =
//                                    from url in urlListM select ProcessURLMS(url, client, ct);

//            List<Task<MeteoSt>> downloadTasks = downloadTasksQuery.ToList();
//            while (downloadTasks.Count > 0)
//            {
//                Task<MeteoSt> firstFinishedTask = await Task.WhenAny(downloadTasks);

//                downloadTasks.Remove(firstFinishedTask);

//                MeteoSt ur = firstFinishedTask.Result;

//                textBox2.Text = String.Format("\r\n {0}: {1}", ur.nomer, ur.naim) + textBox2.Text;

//                //// 1 или 0               textBox3.Text += string.Format("\r\n - {0}", HTMLs9.Count);

//                await Task.Run(() => ParseMeteo(PriznakYear(pdtSelect)));
//            }
//            textBox2.Text += String.Format("\r\n\r\n ПАРСИНГ ЗАГРУЖЕННЫХ СТРАНИЦ");
//        }

//        private void ParseMeteo(string pOld)
//        {
//            while (true)
//            {
//                if (HTMLs9.Count == 0)
//                    break;
//                else
//                    if (HTMLs9.Count > 0)
//                        if (!HTMLs9[0].ppo)           //Метео
//                        {
//                            //парсим страницу на предмет метеоданных и помещаем ее в массив MeteoData
//                            AviaLesParseMeteoData meteo = new AviaLesParseMeteoData();
//                            string[,] MeteoData;

//                            try
//                            {
//                                MeteoData = meteo.ParsePageMeteo(HTMLs9[0].html, false);
//                            }
//                            catch (Exception)
//                            {
//                                MessageBox.Show("Ошибка meteo.ParsePage!", "Meteo");
//                                break;
//                                //throw;
//                            }

//                            //парсим страницу на предмет показателей пожарной опасности (ППО) и помещаем ее в массив ppo
//                            AviaLesParseMeteoPPO ppp = new AviaLesParseMeteoPPO();
//                            string[] ppo;
//                            ppo = ppp.ParsePagePPO(HTMLs9[0].html, false);

//                            if (String.IsNullOrEmpty(MeteoData[0, 0]) || String.IsNullOrEmpty(ppo[0]))
//                            {
//                                //нет данных
//                                //сохраняем только сведения в журнал MeteoJurnal
//                                //InsMeteoJurnal(HTMLs9[0].reg, "29111", HTMLs9[0].dt);
//                            }
//                            else
//                            {
//                                //иначе грузим всё
//                                //InsMeteoSt(MeteoData, ppo, HTMLs9[0].reg, "29111", HTMLs9[0].dt);
//                                //DataTable tbl2 = InsMeteoSt(MeteoData, ppo, HTMLs9[0].reg, 
//                                //                        HTMLs9[0].meteoIndex, HTMLs9[0].meteoNaim, pdtSelect);

//                                AviaLesMeteo.LoadPage.InsMeteoSt(MeteoData, ppo, HTMLs9[0].regNaim,
//                                                HTMLs9[0].meteoIndex, HTMLs9[0].meteoNaim, HTMLs9[0].gmt,
//                                                dtSelect, Program.tbl, Program.tblM);
//                            }
//                                                    //Thread.Sleep(450); .Sleep(777); было .Sleep(500);
//                            HTMLs9.RemoveAt(0);
//                        }
//            }
//        }

//        async Task<MeteoSt> ProcessURLMS(MeteoSt url, HttpClient client, CancellationToken ct)
//        {
//            string HTML1 = string.Empty;

//            HttpResponseMessage response = await client.GetAsync(url.url, ct);

//            var task = response.Content.ReadAsByteArrayAsync();
//            var taskDebug = task.ContinueWith((task1)
//                    => Debug.WriteLine("URL:{0}, Thread:{1}", url, Thread.CurrentThread.ManagedThreadId));
//            byte[] urlContents = await task;

//            HTML1 = Encoding.GetEncoding(1251).GetString(urlContents);

//            HTMLs9.Add(new SiteHTML(HTML1, url.reg, url.regNaim, "OKR",
//               dtSelect.ToString("yyyy-MM-dd"), url.nomer, url.naim, url.gmt, false));

//            return url;
//        }

//        private string PriznakYear(DateTime dt)
//        {
//            return (dt.Year < 2007 ? 1 : 0).ToString();
//        }

//        private void CreateDS()
//        {
//            //DataSet ds = new DataSet();
//            //DataTable tbl = ds.Tables.Add("Meteo");
//            //DataTable tbl = ds.Tables.Add("Meteo");

//            Program.ds = new DataSet();
//            Program.tbl = Program.ds.Tables.Add("Ветер");
//            Program.tblM = Program.ds.Tables.Add("Метео");
//            Program.tblP = Program.ds.Tables.Add("ППО");

//            //DataColumn col = tbl.Columns.Add("idMeteo", typeof(int));
//            //col.AutoIncrement = true;
//            //col.AutoIncrementSeed = -1;
//            //col.AutoIncrementStep = -1;
//            //col.ReadOnly = true;

//            Program.tbl.Columns.Add("ФО", typeof(string));
//            Program.tbl.Columns.Add("Область", typeof(string));
//            Program.tbl.Columns.Add("Index", typeof(string));
//            Program.tbl.Columns.Add("Название", typeof(string));
//            Program.tbl.Columns.Add("Ч-й пояс", typeof(int));
//            Program.tbl.Columns.Add("Дата", typeof(string));
//            Program.tbl.Columns.Add("u0", typeof(int));
//            Program.tbl.Columns.Add("v0", typeof(int));
//            Program.tbl.Columns.Add("u3", typeof(int));
//            Program.tbl.Columns.Add("v3", typeof(int));
//            Program.tbl.Columns.Add("u6", typeof(int));
//            Program.tbl.Columns.Add("v6", typeof(int));
//            Program.tbl.Columns.Add("u9", typeof(int));
//            Program.tbl.Columns.Add("v9", typeof(int));
//            Program.tbl.Columns.Add("u12", typeof(int));
//            Program.tbl.Columns.Add("v12", typeof(int));
//            Program.tbl.Columns.Add("u15", typeof(int));
//            Program.tbl.Columns.Add("v15", typeof(int));
//            Program.tbl.Columns.Add("u18", typeof(int));
//            Program.tbl.Columns.Add("v18", typeof(int));
//            Program.tbl.Columns.Add("u21", typeof(int));
//            Program.tbl.Columns.Add("v21", typeof(int));

//            Program.tblM.Columns.Add("Регион", typeof(string));
//            Program.tblM.Columns.Add("Index", typeof(string));
//            Program.tblM.Columns.Add("Метеостанция", typeof(string));
//            Program.tblM.Columns.Add("МСК", typeof(string));
//            Program.tblM.Columns.Add("Температура", typeof(string));
//            Program.tblM.Columns.Add("Ветер", typeof(string));
//            Program.tblM.Columns.Add("Комментарий", typeof(string));

//            Program.tblP.Columns.Add("Округ", typeof(string));
//            Program.tblP.Columns.Add("Область", typeof(string));
//            Program.tblP.Columns.Add("Индекс", typeof(string));
//            Program.tblP.Columns.Add("Название", typeof(string));
//            Program.tblP.Columns.Add("Часовой пояс", typeof(string));
//            Program.tblP.Columns.Add("Широта", typeof(string));
//            Program.tblP.Columns.Add("Долгота", typeof(string));
//            Program.tblP.Columns.Add("Дата", typeof(string));
//            Program.tblP.Columns.Add("Прогноз", typeof(string));
//            Program.tblP.Columns.Add("KPPON", typeof(string));
//            Program.tblP.Columns.Add("ClaccN", typeof(string));
//            Program.tblP.Columns.Add("KPPO1", typeof(string));
//            Program.tblP.Columns.Add("Clacc1", typeof(string));
//            Program.tblP.Columns.Add("KPPO2", typeof(string));
//            Program.tblP.Columns.Add("Clacc2", typeof(string));
//            Program.tblP.Columns.Add("T", typeof(string));
//            Program.tblP.Columns.Add("TR", typeof(string));
//            Program.tblP.Columns.Add("Время", typeof(string));
//            Program.tblP.Columns.Add("Осадки", typeof(string));
//            Program.tblP.Columns.Add("Снег", typeof(string));
//        }

//        private bool AnalizChkBox(out Subjects[] Subj)
//        {
//            string output = string.Empty;                          //номер субъекта
//            string outputName = string.Empty;                      //название субъекта
//            string outputOkrugId = string.Empty;                   //код округа

//            foreach (Control control in this.Controls)
//                //Перебираем все контролы на форме
//                if (control is CheckBox)                    //Если контрол - чекбокс 
//                {
//                    CheckBox checkBox = (CheckBox)control;
//                    //Если флажок установлен и это не округ
//                    if (checkBox.Checked && checkBox.Name.Length > 11)    
//                    {
//                        output += checkBox.Tag + ",";       //Добавляем порядковый номер к строке
//                        outputName += checkBox.Text + ",";
//                        outputOkrugId += checkBox.Name.Substring(9, 1) + ",";
//                    }
//                }
//                else
//                    continue;

//            if (output.Length > 1)
//            {
//                //Обрезаем запятую и пробел в конце строки
//                output = output.Substring(0, output.Length - 1);
//                outputName = outputName.Substring(0, outputName.Length - 1);
//                outputOkrugId = outputOkrugId.Substring(0, outputOkrugId.Length - 1);
//            }

//            if (String.IsNullOrEmpty(output))
//            {
//                MessageBox.Show("Вы не выбрали ни одного субъекта!", "Внимание");
//                Subj = null;
//                return false;
//            }    
//            else
//            {
//                Subj = GetSubj(output, outputName, outputOkrugId);
//                return true;     
//            }
// //!!!           MessageBox.Show(output);
// ////           return true;        
//        }

//        private Subjects[] GetSubj(string id, string name, string okrugId)
//        {
//            string[] splitId = id.Split(',');
//            string[] splitName = name.Split(',');
//            string[] splitOkrug = okrugId.Split(',');

//            if (splitId.Length == splitName.Length)
//            {
//                Subjects[] subj = new Subjects[splitId.Length];
//                for (int i = 0; i < splitId.Length; i++)
//                {
//                    subj[i].Id = splitId[i];
//                    subj[i].Name = splitName[i];
//                    subj[i].OkrugId = splitOkrug[i];
//                }
//                return subj;
//            }
//            return null;
//        }

//        //лист1 ДАННЫЕ (МетеоДанные) 
//        //лист2 ВЕТЕР
//        private void ExelObject3(DataTable dt, DataTable dtM)
//        {
//            xlApp = new Excel.Application();

//            try
//            {
//                //добавляем книгу
//                xlApp.Workbooks.Add(Type.Missing);

//                //xlApp.SheetsInNewWorkbook = 1;
//                //MessageBox.Show(xlApp.Sheets.Count.ToString());

//                if (xlApp.Sheets.Count == 1)
//                    xlApp.Sheets.Add();
//                //else if (xlApp.Sheets.Count == 3 || xlApp.Sheets.Count == 4)
//                //    xlApp.Sheets.Delete();

//                xlApp.EnableEvents = false;
////I.
//                //выбираем лист на котором будем работать (Лист 1)
//                xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
//                //Название листа
//                xlSheet.Name = "Ветер";

//                //Выгрузка данных
//                //DataTable dt = GetData();

//                int collInd = 0;
//                int rowInd = 0;
//                string data = string.Empty;

//                DataView vue = new DataView(dt);
//                //vue.Table = dt;
//                vue.Sort = "Дата ASC, Область ASC, Название ASC";
//                DataTable dtSort = vue.ToTable();
                
//                //называем колонки
//                for (int i = 0; i < dtSort.Columns.Count; i++)
//                {
//                    data = dtSort.Columns[i].ColumnName.ToString();
//                    xlSheet.Cells[2, i + 1] = data;

//                    //выделяем первую строку
//                    xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);

//                    //делаем полужирный текст и перенос слов
//                    xlSheetRange.WrapText = true;
//                    xlSheetRange.Font.Bold = true;
//                }

//                //заполняем строки
//                for (rowInd = 0; rowInd < dtSort.Rows.Count; rowInd++)
//                    for (collInd = 0; collInd < dtSort.Columns.Count; collInd++)
//                    {
//                        data = dtSort.Rows[rowInd].ItemArray[collInd].ToString();
//                        xlSheet.Cells[rowInd + 3, collInd + 1] = data;
//                    }

//                //выбираем всю область данных
//                xlSheetRange = xlSheet.UsedRange;

//                //выравниваем строки и колонки по их содержимому
//                xlSheetRange.Columns.AutoFit();
//                xlSheetRange.Rows.AutoFit();

//                xlSheet.Cells[1, 1] =
//                    "Направление и скорость ветра для выгрузки за " + dtSelect.ToString("dd.MM.yyyy") + " г.";
////II.
//                //выбираем лист на котором будем работать (Лист 1)
//                xlSheet = (Excel.Worksheet)xlApp.Sheets[2];
//                //Название листа
//                xlSheet.Name = "Метео";

//                DataView vueM = new DataView(dtM);
//                vueM.Sort = "Регион ASC, Метеостанция ASC";
//                DataTable dtMSort = vueM.ToTable();

//                //называем колонки
//                for (int i = 0; i < dtMSort.Columns.Count; i++)
//                {
//                    data = dtMSort.Columns[i].ColumnName.ToString();
//                    xlSheet.Cells[2, i + 1] = data;

//                    //выделяем первую строку
//                    xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);

//                    //делаем полужирный текст и перенос слов
//                    xlSheetRange.WrapText = true;
//                    xlSheetRange.Font.Bold = true;
//                }

//                //заполняем строки
//                for (rowInd = 0; rowInd < dtMSort.Rows.Count; rowInd++)
//                    for (collInd = 0; collInd < dtMSort.Columns.Count; collInd++)
//                    {
//                        data = dtMSort.Rows[rowInd].ItemArray[collInd].ToString();
//                        xlSheet.Cells[rowInd + 3, collInd + 1] = data;
//                    }

//                //выбираем всю область данных
//                xlSheetRange = xlSheet.UsedRange;

//                //выравниваем строки и колонки по их содержимому
//                xlSheetRange.Columns.AutoFit();
//                xlSheetRange.Rows.AutoFit();

//                xlSheet.Cells[1, 1] = 
//                    "Фактическая метеообстановка на " + dtSelect.ToString("dd.MM.yyyy") + " г.";               
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show(ex.ToString());
//            }
//            finally
//            {
//                //Показываем ексель
//                xlApp.Visible = true;

//                xlApp.Interactive = true;
//                xlApp.ScreenUpdating = true;
//                xlApp.UserControl = true;

//                //Отсоединяемся от Excel
//                releaseObject(xlSheetRange);
//                releaseObject(xlSheet);
//                releaseObject(xlApp);
//            }
//        }

//        private void ExelObjectNiokr(DataTable dtP)
//        {
//            xlApp = new Excel.Application();

//            try
//            {
//                //добавляем книгу
//                xlApp.Workbooks.Add(Type.Missing);

//                if (xlApp.Sheets.Count == 1)
//                    xlApp.Sheets.Add();

//                xlApp.EnableEvents = false;
//                //I.
//                //выбираем лист на котором будем работать (Лист 1)
//                xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
//                //Название листа
//                xlSheet.Name = "ВСЕ";

//                //Выгрузка данных
//                //DataTable dt = GetData();

//                int collInd = 0;
//                int rowInd = 0;
//                string data = string.Empty;

//                DataView vue = new DataView(dtP);
//                //vue.Table = dt;
//                vue.Sort = "Дата ASC, Область ASC, Название ASC";
//                DataTable dtSort = vue.ToTable();

//             //   xlSheetRange.Borders.xlEdgeLeft;

//                //называем колонки
//                //for (int i = 0; i < dtSort.Columns.Count; i++)
//                //{
//                //    data = dtSort.Columns[i].ColumnName.ToString();
//                //    xlSheet.Cells[2, i + 1] = data;

//                //    //выделяем первую строку
//                //    xlSheetRange = xlSheet.get_Range("A2:Z2", Type.Missing);

//                //    //делаем полужирный текст и перенос слов
//                //    xlSheetRange.WrapText = true;
//                //    xlSheetRange.Font.Bold = true;
//                //}

//                //////////////////////////////////////////////////////////////////////////////////
//                //рисуем шапку таблицы

//                xlSheetRange = xlSheet.get_Range("A3", "T3");
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Font.Size = 14;
//                xlSheetRange.Value2 = "Метеорологические данные" + dtSelect.ToString("dd.MM.yyyy") + " г.";

//                //Задаем выравнивание по центру
//                xlSheetRange.HorizontalAlignment = Excel.Constants.xlCenter;
//                xlSheetRange.VerticalAlignment = Excel.Constants.xlCenter;

//                //Выбираем ячейку для вывода 
//                xlSheetRange = xlSheet.get_Range("A5", "A9");
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Округ";

//                xlSheetRange = xlSheet.get_Range("B5", "B9");
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Область";

//                xlSheetRange = xlSheet.get_Range("C5", "D6");
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Метеостанция";

//                xlSheetRange = xlSheet.get_Range("C7", "C9");
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Индекс";

//                xlSheetRange = xlSheet.get_Range("D7", "D9");
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Название";


//                xlSheetRange = xlSheet.get_Range("E5", "E9");
//                xlSheetRange.WrapText = true;
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Часовой пояс";


//                xlSheetRange = xlSheet.get_Range("F5", "F9");
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Широта";

//                xlSheetRange = xlSheet.get_Range("G5", "G9");
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Долгота";

//                xlSheetRange = xlSheet.get_Range("H5", "H9");
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Дата";

//                xlSheetRange = xlSheet.get_Range("I5", "I9");
//                xlSheetRange.WrapText = true;
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Прог-ноз";

//                xlSheetRange = xlSheet.get_Range("J5", "K7");
//                xlSheetRange.WrapText = true;
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Методика Нестерова";

//                xlSheetRange = xlSheet.get_Range("J8", "J9");
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "КППО";

//                xlSheetRange = xlSheet.get_Range("K8", "K9");
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Класс ПО";

//                xlSheetRange = xlSheet.get_Range("L5", "M7");
//                xlSheetRange.WrapText = true;
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "ПВ-1 (на основе влажности напочвенного покрова)";
//                xlSheetRange.Font.Size = 10;

//                xlSheetRange = xlSheet.get_Range("L8", "L9");
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "КППО";

//                xlSheetRange = xlSheet.get_Range("M8", "M9");
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Класс ПО";

//                xlSheetRange = xlSheet.get_Range("N5", "O7");
//                xlSheetRange.WrapText = true;
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "ПВ-2 (на основе лесной подстилки)";
//                xlSheetRange.Font.Size = 10;

//                xlSheetRange = xlSheet.get_Range("N8", "N9");
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "КППО";

//                xlSheetRange = xlSheet.get_Range("O8", "O9");
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Класс ПО";

//                xlSheetRange = xlSheet.get_Range("P5", "P9");
//                xlSheetRange.WrapText = true;
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Т° воздуха";

//                xlSheetRange = xlSheet.get_Range("Q5", "Q9");
//                xlSheetRange.WrapText = true;
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Точка росы";

//                xlSheetRange = xlSheet.get_Range("R5", "R9");
//                xlSheetRange.WrapText = true;
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Время измерения Т° и точки росы (местное)";
//                xlSheetRange.Font.Size = 10;

//                xlSheetRange = xlSheet.get_Range("S5", "S9");
//                xlSheetRange.WrapText = true;
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Суточные осадки на 9:00 час. (мест.вр.), мм";

//                xlSheetRange = xlSheet.get_Range("T5", "T9");
//                xlSheetRange.WrapText = true;
//                //Объединяем ячейки
//                xlSheetRange.Merge(Type.Missing);
//                xlSheetRange.Value2 = "Высота снежного покрова, см";

//                xlSheetRange = xlSheet.get_Range("A5", "T9");
//                //Устанавливаем цвет обводки
//                xlSheetRange.Borders.ColorIndex = 1;
//                //Устанавливаем стиль и толщину линии
//                xlSheetRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
//                xlSheetRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

//                //Задаем выравнивание по центру
//                xlSheetRange.HorizontalAlignment = Excel.Constants.xlCenter;
//                xlSheetRange.VerticalAlignment = Excel.Constants.xlCenter;

//                //////////////////////////////////////////////////////////////////////////////////
//                //заполняем строки
//                for (rowInd = 0; rowInd < dtSort.Rows.Count; rowInd++)
//                    for (collInd = 0; collInd < dtSort.Columns.Count; collInd++)
//                    {
//                        data = dtSort.Rows[rowInd].ItemArray[collInd].ToString();
//                        xlSheet.Cells[rowInd + 10, collInd + 1] = data;
//                    }

//                //выбираем всю область данных
//                xlSheetRange = xlSheet.UsedRange;

//                //выравниваем строки и колонки по их содержимому
//                xlSheetRange.Columns.AutoFit();
//                xlSheetRange.Rows.AutoFit();

//                //xlSheetRange = (Excel.Range)xlSheet.Columns["A", Type.Missing];
//                //xlSheetRange.ColumnWidth = 7.29;

//                //xlSheetRange = (Excel.Range)xlSheet.Columns["B", Type.Missing];
//                //xlSheetRange.ColumnWidth = 15.29;

//                //xlSheetRange = (Excel.Range)xlSheet.Columns["C", Type.Missing];
//                //xlSheetRange.ColumnWidth = 8.43;

//                //xlSheetRange = (Excel.Range)xlSheet.Columns["D", Type.Missing];
//                //xlSheetRange.ColumnWidth = 16;

//                xlSheetRange = (Excel.Range)xlSheet.Columns["E", Type.Missing];
//                xlSheetRange.ColumnWidth = 9;

//                //xlSheetRange = (Excel.Range)xlSheet.Columns["F", Type.Missing];
//                //xlSheetRange.ColumnWidth = 10;

//                //xlSheetRange = (Excel.Range)xlSheet.Columns["G", Type.Missing];
//                //xlSheetRange.ColumnWidth = 11;

//                xlSheetRange = (Excel.Range)xlSheet.Columns["H", Type.Missing];
//                xlSheetRange.ColumnWidth = 10.29;

//                xlSheetRange = (Excel.Range)xlSheet.Columns["I", Type.Missing];
//                xlSheetRange.ColumnWidth = 5.71;

//                xlSheetRange = (Excel.Range)xlSheet.Columns["J", Type.Missing];
//                xlSheetRange.ColumnWidth = 8.57;

//                //xlSheetRange = (Excel.Range)xlSheet.Columns["K", Type.Missing];
//                //xlSheetRange.ColumnWidth = 8.57;

//                xlSheetRange = (Excel.Range)xlSheet.Columns["L", Type.Missing];
//                xlSheetRange.ColumnWidth = 10.29;

//                xlSheetRange = (Excel.Range)xlSheet.Columns["M", Type.Missing];
//                xlSheetRange.ColumnWidth = 10.29;

//                xlSheetRange = (Excel.Range)xlSheet.Columns["N", Type.Missing];
//                xlSheetRange.ColumnWidth = 8.57;

//                xlSheetRange = (Excel.Range)xlSheet.Columns["O", Type.Missing];
//                xlSheetRange.ColumnWidth = 8.57;

//                xlSheetRange = (Excel.Range)xlSheet.Columns["P", Type.Missing];
//                xlSheetRange.ColumnWidth = 8.57;

//                xlSheetRange = (Excel.Range)xlSheet.Columns["Q", Type.Missing];
//                xlSheetRange.ColumnWidth = 7.57;

//                xlSheetRange = (Excel.Range)xlSheet.Columns["R", Type.Missing];
//                xlSheetRange.ColumnWidth = 10.57;

//                xlSheetRange = (Excel.Range)xlSheet.Columns["S", Type.Missing];
//                xlSheetRange.ColumnWidth = 10.57;

//                xlSheetRange = (Excel.Range)xlSheet.Columns["T", Type.Missing];
//                xlSheetRange.ColumnWidth = 10;

//                xlSheet.Cells[1, 1] =
//                    "Метеорологические данные " + dtSelect.ToString("dd.MM.yyyy") + " г.";
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show(ex.ToString());
//            }
//            finally
//            {
//                //Показываем ексель
//                xlApp.Visible = true;

//                xlApp.Interactive = true;
//                xlApp.ScreenUpdating = true;
//                xlApp.UserControl = true;

//                //Отсоединяемся от Excel
//                releaseObject(xlSheetRange);
//                releaseObject(xlSheet);
//                releaseObject(xlApp);
//            }
//        }

//        //лист НИОКР
//        private void ExelObject()
//        {
//            xlApp = new Excel.Application();

//            try
//            {
//                //добавляем книгу
//                xlApp.Workbooks.Add(Type.Missing);

//                //делаем временно неактивным документ
//                xlApp.Interactive = false;
//                xlApp.EnableEvents = false;

//                //выбираем лист на котором будем работать (Лист 1)
//                xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
//                //Название листа
//                xlSheet.Name = "Данные";

//                //Выгрузка данных
//                DataTable dt = GetData();

//                int collInd = 0;
//                int rowInd = 0;
//                string data = "";

//                //называем колонки
//                for (int i = 0; i < dt.Columns.Count; i++)
//                {
//                    data = dt.Columns[i].ColumnName.ToString();
//                    xlSheet.Cells[1, i + 1] = data;

//                    //выделяем первую строку
//                    xlSheetRange = xlSheet.get_Range("A1:Z1", Type.Missing);

//                    //делаем полужирный текст и перенос слов
//                    xlSheetRange.WrapText = true;
//                    xlSheetRange.Font.Bold = true;
//                }

//                //заполняем строки
//                for (rowInd = 0; rowInd < dt.Rows.Count; rowInd++)
//                {
//                    for (collInd = 0; collInd < dt.Columns.Count; collInd++)
//                    {
//                        data = dt.Rows[rowInd].ItemArray[collInd].ToString();
//                        xlSheet.Cells[rowInd + 2, collInd + 1] = data;
//                    }
//                }

//                //выбираем всю область данных
//                xlSheetRange = xlSheet.UsedRange;

//                //выравниваем строки и колонки по их содержимому
//                xlSheetRange.Columns.AutoFit();
//                xlSheetRange.Rows.AutoFit();
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show(ex.ToString());
//            }
//            finally
//            {
//                //Показываем ексель
//                xlApp.Visible = true;

//                xlApp.Interactive = true;
//                xlApp.ScreenUpdating = true;
//                xlApp.UserControl = true;

//                //Отсоединяемся от Excel
//                releaseObject(xlSheetRange);
//                releaseObject(xlSheet);
//                releaseObject(xlApp);
//            }
//        }
              
//        //Освобождаем ресуры (закрываем Excel)
//        void releaseObject(object obj)
//        {
//            try
//            {
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
//                obj = null;
//            }
//            catch (Exception ex)
//            {
//                obj = null;
//                MessageBox.Show(ex.ToString(), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
//            }
//            finally
//            {
//                GC.Collect();
//            }
//        }

//        private DataTable GetData()
//        {

//            string cnnSQL = ConfigurationManager.ConnectionStrings["SQLAviaLesMeteo"].ConnectionString;

//            //строка подключения к SQL-серверу
//            //cnSql.ConnectionString = @cnnSQL;               //cnSql.ConnectionString = cnnString;

//            //строка соединения
//            string connString = @cnnSQL;

//            //соединение
//            SqlConnection con = new SqlConnection(connString);

//            DataTable dt = new DataTable();
//            try
//            {
//                string query = @"SELECT TOP 10 [id] ,[dtIzm] ,[vidim] ,[veterN] ,[veterP] ,[veterV]
//                                    ,[temperature] ,[timeGMT] ,[bOsadki] ,[timeMest] ,[dtGMT] ,[id_stn]
//                                FROM [AviaLesMeteo].[dbo].[MeteoData]";

//                SqlCommand comm = new SqlCommand(query, con);

//                con.Open();
//                SqlDataAdapter da = new SqlDataAdapter(comm);
//                DataSet ds = new DataSet();
//                da.Fill(ds);
//                dt = ds.Tables[0];
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
//            }
//            finally
//            {
//                con.Close();
//                con.Dispose();
//            }
//            return dt;
//        }

//        private void PrepareComponents()
//        {
//            this.label2.Visible = true;
//            this.label3.Visible = true;
//            this.label4.Visible = true;

//            textBox3.Location = new Point(10, 55);
//            textBox3.Size = new Size(175, 473);

//            textBox1.Location = new Point(192, 55);
//            textBox1.Size = new Size(311, 473);

//            textBox2.Location = new Point(509, 55);
//            textBox2.Size = new Size(296, 473);

//            textBox3.Visible = true;
//            textBox1.Visible = true;
//            textBox2.Visible = true;

//        }

//        private void checkBox_1_CheckedChanged(object sender, EventArgs e)
//        {
//            SelectMeteo((CheckBox)this.ActiveControl);
//        }

//        private void checkBox_2_CheckedChanged(object sender, EventArgs e)
//        {
//            SelectMeteo((CheckBox)this.ActiveControl);
//        }

//        private void checkBox_3_CheckedChanged(object sender, EventArgs e)
//        {
//            SelectMeteo((CheckBox)this.ActiveControl);
//        }

//        private void checkBox_4_CheckedChanged(object sender, EventArgs e)
//        {
//            SelectMeteo((CheckBox)this.ActiveControl);
//        }

//        private void checkBox_5_CheckedChanged(object sender, EventArgs e)
//        {
//            SelectMeteo((CheckBox)this.ActiveControl);
//        }

//        private void checkBox_6_CheckedChanged(object sender, EventArgs e)
//        {
//            SelectMeteo((CheckBox)this.ActiveControl);
//        }

//        private void checkBox_7_CheckedChanged(object sender, EventArgs e)
//        {
//            SelectMeteo((CheckBox)this.ActiveControl);
//        }

//        private void checkBox_8_CheckedChanged(object sender, EventArgs e)
//        {
//            SelectMeteo((CheckBox)this.ActiveControl);
//        }

//        private void checkBox_9_CheckedChanged(object sender, EventArgs e)
//        {
//            SelectMeteo((CheckBox)this.ActiveControl);
//        }

//        private void SelectMeteo(CheckBox chkBox)
//        {
//            string qqq = string.Empty;

//            foreach (Control control in this.Controls)
//                //Перебираем все контролы на форме
//                if (control is CheckBox)                    //Если контрол - чекбокс 
//                {
//                    CheckBox checkBox = (CheckBox)control;
//                    qqq = checkBox.Name.PadRight(20, '-');
//                    //if (qqq.Substring(0, qqq.Length - 9) == "checkBox_7_")
//                    if (qqq.Substring(0, qqq.Length - 9) == chkBox.Name + "_")
//                    {
//                        //MessageBox.Show(qqq.Substring(0, qqq.Length - 9));
//                        checkBox.Checked = chkBox.Checked ? true : false;
//                    }
 
//                }
//                else
//                    continue;
//        }

//        //////////////////////////////////////////////////////////////

//        private DataTable CreateData()
//        {
//            DataSet ds = new DataSet();

//            DataTable tbl = ds.Tables.Add("Meteo");
//            //DataColumn col = tbl.Columns.Add("idMeteo");
//            DataColumn col = tbl.Columns.Add("idMeteo", typeof(int));
//            col.AutoIncrement = true;
//            col.AutoIncrementSeed = -1;
//            col.AutoIncrementStep = -1;
//            col.ReadOnly = true;

//            //col.AllowDBNull = false;
//            //col.ReadOnly = false;
//            //col.MaxLength = 5;
//            //col.Unique = true; 

//            return tbl;
//        }
 
//        /////////////////////////////////////////////////////////////////////////////////////////////////////
//        /////////////////////////////////////////////////////////////////////////////////////////////////////
//        //private IEnumerable<string> SetUpURLList2()
//        //{
//        //    //создаем коллекцию строк URL-адресов, которые хотим загрузить
//        //    var urls = new List<string> 
//        //    { 
//        //        "http://www.cbr.ru",
//        //        "http://www.yandex.ru",
//        //        "http://www.rambler.ru",
//        //        "http://www.vz.ru",
//        //        "http://www.mail.ru",
//        //        "http://msdn.microsoft.com/library/windows/apps/br211380.aspx",
//        //        "http://msdn.microsoft.com/library/windows/apps/br211380.aspx",
//        //        "http://msdn.microsoft.com",
//        //        "http://msdn.microsoft.com/en-us/library/hh290136.aspx",
//        //        "http://msdn.microsoft.com/en-us/library/ee256749.aspx",
//        //        "http://msdn.microsoft.com/en-us/library/hh290138.aspx",
//        //        "http://www.yandex.ru",
//        //        "http://msdn.microsoft.com/library/windows/apps/br211380.aspx",
//        //        "http://msdn.microsoft.com",
//        //        "http://msdn.microsoft.com/en-us/library/hh290136.aspx",
//        //        "http://msdn.microsoft.com/en-us/library/ee256749.aspx",
//        //        "http://www.rambler.ru",
//        //        "http://msdn.microsoft.com/library/windows/apps/br211380.aspx",
//        //        "http://msdn.microsoft.com",
//        //        "http://www.cbr.ru",
//        //        "http://msdn.microsoft.com/en-us/library/hh290136.aspx",
//        //        "http://msdn.microsoft.com/en-us/library/ee256749.aspx",
//        //        "http://www.regnum.ru",
//        //        "http://msdn.microsoft.com/en-us/library/hh290138.aspx",
//        //        "http://www.svpressa.ru",
//        //        "http://www.ria.ru",
//        //        "http://www.rbc.ru",
//        //        "http://www.gazeta.ru",
//        //        "http://www.lenta.ru",
//        //        "http://www.cbr.ru"
//        //    };
//        //    //возвращаем ссылку на коллекцию
//        //    return urls;
//        //}
//        /////////////////////////////////////////////////////////////////////////////////////////////////////
//        /////////////////////////////////////////////////////////////////////////////////////////////////////
//    }
//}