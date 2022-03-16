using HtmlAgilityPack;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ParserWienerborse
{
    class Program
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        private static List<string> Data = new List<string>();
        [STAThread]
        static void Main(string[] args)
        {
            logger.Info("НАЧАЛО РАБОТЫ ПАРСЕРА");
            logger.Info("Начало загрузки Prime market");
            LoadStocks("https://www.wienerborse.at/en/market-data/shares-others/prime-market/", "Prime market");
            logger.Info("Окончание загрузки рынка");

            logger.Info("Начало загрузки Standard market");
            LoadStocks("https://www.wienerborse.at/en/market-data/shares-others/standard-market/", "Standard market");
            logger.Info("Окончание загрузки рынка");

            logger.Info("Начало загрузки Direct market plus");
            LoadStocks("https://www.wienerborse.at/en/marktdaten/preisinformationen/aktien-sonstige/shares-price-list/direct-market-plus/", "Direct market plus");
            logger.Info("Окончание загрузки рынка");                  

            logger.Info("Начало загрузки Direct market");
            LoadStocks("https://www.wienerborse.at/en/marktdaten/preisinformationen/aktien-sonstige/shares-price-list/direct-market/", "Direct market");
            logger.Info("Окончание загрузки рынка");

            logger.Info("Начало загрузки Global market");
            LoadStocksPage("https://www.wienerborse.at/en/market-data/shares-others/global-market/?c52166-page=");
            logger.Info("Окончание загрузки рынка");
            
            using(StreamWriter stream = new StreamWriter(args[0] + "\\" + DateTime.Today.ToString("dd" + "MM" + "yyyy") + ".csv"))
            {
                stream.WriteLine("sep=,");
                stream.WriteLine("ISIN, Name, Last price, Perc change, Change abs, Open, High, Low, Total Volume, Total Value, MPQ, Market, Group");
                foreach (var item in Data)
                {
                    stream.WriteLine(item);
                }
            }
            logger.Info("ОКОНЧАНИЕ РАБОТЫ ПАРСЕРА");
        }

        static void LoadStocks(string url, string group)
        {
            HtmlWeb htmlWeb = new HtmlWeb();
            bool exit = true;
            int j;
            string line, name;
            while (exit)
            {
                exit = false;
                HtmlDocument webDoc = htmlWeb.Load(url);
                HtmlNodeCollection nodes = webDoc.DocumentNode.SelectNodes("//tbody/tr/td/div/a/span | //tbody/tr/td");
                if (nodes != null)
                {
                    j = 0;
                    line = "";
                    name = "";
                    string[] splitLine = new string[2];
                    foreach (var item in nodes)
                    {
                        switch (j)
                        {
                            case 0:
                                name = item.InnerText.Trim();
                                j += 1;
                                break;
                            case 1:
                                name = name.Replace(item.InnerText.Trim(), ", ");
                                line += item.InnerText.Trim().Replace(',', '.') + ", " + name;
                                j += 1;
                                break;
                            case 3:
                                if (item.InnerText == "--")
                                {
                                    line += "-, -, ";
                                }
                                else
                                {
                                    splitLine = item.InnerText.Split('%');
                                    line += splitLine[0].Replace(',', '.').Trim() + "%, " + splitLine[1].Replace(',', '.').Trim() + ", ";
                                    splitLine = null;
                                }
                                j += 1;
                                break;
                            case 4:
                                j += 1;
                                break;
                            case 11:
                                line += item.InnerText.Trim().Replace(",", ".") + ", " + group;
                                Data.Add(line);
                                line = "";
                                j = 0;
                                break;
                            default:
                                line += item.InnerText.Trim().Replace(",", ".") + ", ";
                                j += 1;
                                break;
                        }
                    }
                }                
            }
        }
        static void LoadStocksPage(string url)
        {
            int maxPage = PageStocks("https://www.wienerborse.at/en/market-data/shares-others/global-market/?c52166-page=1");
            HtmlWeb htmlWeb = new HtmlWeb();
            int j;
            string line, name;
            for (int i = 1; i <= maxPage; i++)
            {
                HtmlDocument webDoc = htmlWeb.Load(url + i);
                HtmlNodeCollection nodes = webDoc.DocumentNode.SelectNodes("//tbody/tr/td/div/a/span | //tbody/tr/td");
                if (nodes != null)
                {
                    j = 0;
                    line = "";
                    name = "";
                    string[] splitLine = new string[2];
                    foreach (var item in nodes)
                    {
                        switch (j)
                        {
                            case 0:
                                name = item.InnerText.Trim();
                                j += 1;
                                break;
                            case 1:
                                name = name.Replace(item.InnerText.Trim(), ", ");
                                line += item.InnerText.Trim().Replace(',', '.') + ", " + name;
                                j += 1;
                                break;
                            case 3:
                                if (item.InnerText == "--")
                                {
                                    line += "-, -, ";
                                }
                                else
                                {
                                    splitLine = item.InnerText.Split('%');
                                    line += splitLine[0].Replace(',', '.').Trim() + "%, " + splitLine[1].Replace(',', '.').Trim() + ", ";
                                    splitLine = null;
                                }
                                j += 1;
                                break;
                            case 4:
                                j += 1;
                                break;
                            case 5:
                                line += " , " + item.InnerText.Trim().Replace(",", ".") + ", ";                               
                                j += 1;
                                break;
                            case 10:
                                line += item.InnerText.Trim().Replace(",", ".") + ", Global market";
                                Data.Add(line);
                                line = "";
                                j = 0;
                                break;
                            default:
                                line += item.InnerText.Trim().Replace(",", ".") + ", ";
                                j += 1;
                                break;
                        }
                    }
                    logger.Info("Загрузка " + i + " страницы завершилось");
                }
            }
        }
        static int PageStocks(string url)
        {
            HtmlWeb htmlWeb = new HtmlWeb();
            HtmlDocument webDoc = htmlWeb.Load(url);
            HtmlNodeCollection nodes = webDoc.DocumentNode.SelectNodes("//div[@class=\"pull-right\"]/ul/li");
            return Convert.ToInt32(nodes[nodes.Count() - 2].InnerText);
        }
    }
}
