using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;

/* Proxy use
WebClient wc = new WebClient();
wc.Proxy = new WebProxy(host,port);
var page = wc.DownloadString(url);

HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
doc.LoadHtml(page);*/

namespace PS
{
    // A class to scrape from a certain website..
    class PhoneScraper
    {
        static Excel.Application oXL;
        static Excel._Workbook oWB;
        static Excel._Worksheet oSheet;

        // Program entry point.
        static void Main(string[] args)
        {
            oXL = new Excel.Application();
            oXL.Visible = true;

            oWB = (Excel._Workbook)(oXL.Workbooks.Add(Type.Missing));
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;

            Console.WriteLine("Excel worksheet created.");

            List<string> phonePages = GetListOfPhonePages();
            List<string> specPages = FindPhoneSpecPages(phonePages);

            int i = 1;
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            foreach (string specPage in specPages)
            {
                // Specs are extracted on multiple threads to bypass 
                // the time it takes to make web requests and load the DOM on the page.
                new Thread(() =>
                {
                    ExtractSpecs(specPage, i);
                }).Start();
                Random random = new Random();
                Thread.Sleep(random.Next(1000,3000));
                i++;
            }

            for (int j = 0; j < 1000; j++)
            {
                Thread.Sleep(1000);
                TimeSpan ts = stopWatch.Elapsed;
                // Format and display the TimeSpan value. 
                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                    ts.Hours, ts.Minutes, ts.Seconds,
                    ts.Milliseconds / 10);
                Console.WriteLine("RunTime " + elapsedTime);
            }
        }

        //A hardcoded list of phones. 
        private static List<string> GetListOfPhonePages()
        {
            return new List<string>()
            {
                "http://www.gsmarena.com/nokia-phones-1.php",
                "http://www.gsmarena.com/nokia-phones-f-1-0-p2.php",
                "http://www.gsmarena.com/nokia-phones-f-1-0-p3.php",
                "http://www.gsmarena.com/samsung-phones-9.php",
                "http://www.gsmarena.com/samsung-phones-f-9-0-p2.php",
                "http://www.gsmarena.com/samsung-phones-f-9-0-p3.php",
                "http://www.gsmarena.com/motorola-phones-4.php",
                "http://www.gsmarena.com/motorola-phones-f-4-0-p2.php",
                "http://www.gsmarena.com/motorola-phones-f-4-0-p3.php",
                "http://www.gsmarena.com/sony-phones-7.php",
                "http://www.gsmarena.com/sony-phones-f-7-0-p2.php",
                "http://www.gsmarena.com/lg-phones-20.php",
                "http://www.gsmarena.com/lg-phones-f-20-0-p2.php",
                "http://www.gsmarena.com/lg-phones-f-20-0-p3.php",
                "http://www.gsmarena.com/apple-phones-48.php",
                "http://www.gsmarena.com/htc-phones-45.php",
                "http://www.gsmarena.com/htc-phones-f-45-0-p2.php",
                "http://www.gsmarena.com/htc-phones-f-45-0-p3.php",
                "http://www.gsmarena.com/htc-phones-f-45-0-p4.php",
                "http://www.gsmarena.com/blackberry-phones-36.php",
                "http://www.gsmarena.com/huawei-phones-58.php",
                "http://www.gsmarena.com/huawei-phones-f-58-0-p2.php",
                "http://www.gsmarena.com/zte-phones-62.php",
                "http://www.gsmarena.com/zte-phones-f-62-0-p2.php"
            };
        }


        // Verify and load the nodes of the specified phone pages (Each page has like a list of 50 phones)
        private static List<string> FindPhoneSpecPages(List<string> dataPages)
        {
            var specPages = new List<string>();
            HtmlWeb web = new HtmlWeb();
            HtmlDocument doc;
            string rootURL = "http://www.gsmarena.com/";
            int i = 1;

            foreach (string dataPage in dataPages)
            {
                doc = web.Load(dataPage);
                HtmlNodeCollection specPageNodes = doc.DocumentNode.SelectNodes("//div[@class='makers']/ul/li/a");

                if (specPageNodes == null) 
                { 
                    continue; 
                }

                foreach (HtmlNode specPageNode in specPageNodes)
                {
                    string specPageLink = specPageNode.Attributes["href"].Value;
                    specPages.Add(rootURL + specPageLink);
                    Console.WriteLine(i + ":" + rootURL + specPageLink);
                    i++;
                }

                Random random = new Random();
                Thread.Sleep(random.Next(500, 1000));
            }
            return specPages;
        }

        // Extract the specs fom the DOM.
        private static void ExtractSpecs(string specPage, int index)
        {
            HtmlWeb web = new HtmlWeb();
            HtmlDocument doc;

            doc = web.Load(specPage);

            // Extract the name of the phone - from the subtitle over the image.
            HtmlNode nameNode = doc.DocumentNode.SelectSingleNode("//div[@id='ttl']/h1");
            if (nameNode != null)
            {
                string name = nameNode.InnerText;

                if (name != null)
                {
                    oSheet.Cells[index, 3] = name;
                }
            }

            // Extract the main image of the phone 
            HtmlNode imgNode = doc.DocumentNode.SelectSingleNode("//div[@id='specs-cp-pic']/a/img");
            if (imgNode != null)
            {
                string imgSrc = imgNode.Attributes["src"].Value;

                if (imgSrc != null)
                {
                    Console.WriteLine(index+":Found the image.");
                    const string IMG_PATH = @"C:\Phones1_2_3_test\";
                    using (WebClient client = new WebClient())
                    {
                        client.DownloadFile(imgSrc, IMG_PATH + "img" + index + ".png");
                    }
                }
            }

            // Extract the main data of the phone
            HtmlNodeCollection nodes = doc.DocumentNode.SelectNodes("//table/tr");
            if (nodes != null)
            {
                foreach (HtmlNode node in nodes)
                {
                    if (node == null)
                    {
                        continue;
                    }

                    HtmlNode titleNode = node.SelectSingleNode(".//td[@class='ttl']/a");
                    HtmlNode infoNode = node.SelectSingleNode(".//td[@class='nfo']");
                    HtmlNode priceNode = node.SelectSingleNode(".//td[@class='nfo']/img");

                    if (titleNode == null || infoNode == null)
                    {
                        continue;
                    }

                    string attribute = titleNode.InnerText;
                    string data = infoNode.InnerText;

                    const string PRICE_GROUP = "Price group";
                    if (attribute.Equals(PRICE_GROUP) && priceNode != null)
                    {
                        string price = priceNode.Attributes["title"].Value;
                        string fixedPrice = price.Replace("About", "").Replace("EUR", "").Trim();
                        int priceInt = 0;

                        if (Int32.TryParse(fixedPrice, out priceInt))
                        {
                            const double EUR_TO_USD = 1.25;
                            //Probably shouldn't cast like this..
                            double priceDouble = (double)priceInt;
                                    
                            int priceEUR = (int)priceDouble;
                            int priceUSD = (int)(priceDouble * EUR_TO_USD);

                            const string EUR_PRICE = "EUR Price";
                            PopulateExcelSheet(EUR_PRICE, priceEUR.ToString(), index);
                            const string USD_PRICE = "USD Price";
                            PopulateExcelSheet(USD_PRICE, priceUSD.ToString(), index);
                        }
                    }

                    else
                    {
                        PopulateExcelSheet(attribute, data, index);
                    }

                    HtmlNode headerNode = node.SelectSingleNode(".//th");
                    if (headerNode == null)
                    {
                        continue;
                    }

                    const string BATTERY = "Battery";
                    if (headerNode.InnerText.Equals(BATTERY))
                    {
                        HtmlNode battNode = node.SelectSingleNode(".//td[@class='nfo']");
                        if (battNode != null)
                        {
                            String battData = battNode.InnerText;
                            PopulateExcelSheet(BATTERY, battData, index);
                        }
                    }
                }
            }
            index++;
        }
        
        //Fills cells of the excel sheet with scraped information
        private static void PopulateExcelSheet(string attribute, string data, int index)
        {
            switch (attribute)
            {
                case "SIM":
                    oSheet.Cells[index, 4] = data;
                    break;
                case "Announced":
                    oSheet.Cells[index, 5] = data;
                    break;
                case "Status":
                    oSheet.Cells[index, 6] = data;
                    break;
                case "Dimensions":
                    oSheet.Cells[index, 7] = data;
                    break;
                case "Weight":
                    oSheet.Cells[index, 8] = data;
                    break;
                case "Type":
                    oSheet.Cells[index, 9] = data;
                    break;
                case "Size":
                    oSheet.Cells[index, 10] = data;
                    break;
                case "Card slot":
                    oSheet.Cells[index, 11] = data;
                    break;
                case "Internal":
                    oSheet.Cells[index, 12] = data;
                    break;
                case "NFC":
                    oSheet.Cells[index, 13] = data;
                    break;
                case "USB":
                    oSheet.Cells[index, 14] = data;
                    break;
                case "Primary":
                    oSheet.Cells[index, 15] = data;
                    break;
                case "Features":
                    oSheet.Cells[index, 16] = data;
                    break;
                case "Video":
                    oSheet.Cells[index, 17] = data;
                    break;
                case "Secondary":
                    oSheet.Cells[index, 18] = data;
                    break;
                case "OS":
                    oSheet.Cells[index, 19] = data;
                    break;
                case "Chipset":
                    oSheet.Cells[index, 20] = data;
                    break;
                case "CPU":
                    oSheet.Cells[index, 21] = data;
                    break;
                case "GPU":
                    oSheet.Cells[index, 22] = data;
                    break;
                case "Sensors":
                    oSheet.Cells[index, 23] = data;
                    break;
                case "Battery":
                    oSheet.Cells[index, 24] = data;
                    break;
                case "Stand-by":
                    oSheet.Cells[index, 25] = data;
                    break;
                case "Talk time":
                    oSheet.Cells[index, 26] = data;
                    break;
                case "EUR Price":
                    oSheet.Cells[index, 27] = data;
                    break;
                case "USD Price":
                    oSheet.Cells[index, 28] = data;
                    break;
            }
        }
    }
}
