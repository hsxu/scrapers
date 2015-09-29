using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;

namespace LaptopScraper
{
    class LaptopScraper
    {
        static Excel.Application oXL;
        static Excel._Workbook oWB;
        static Excel._Worksheet oSheet;
        
        static void Main(string[] args)
        {
            oXL = new Excel.Application();
            oXL.Visible = true;

            oWB = (Excel._Workbook)(oXL.Workbooks.Add(Type.Missing));
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;

            Console.WriteLine("Excel worksheet created.");

            List<string> reviewPages = GrabDataPages();
            Console.WriteLine("Data pages found.");

            List<string> specPages = NavigateToSpecs(reviewPages);
            Console.WriteLine("Spec pages found.");

            ExtractSpecs(specPages);
        }

        private static List<string> GrabDataPages()
        {
            var dataPages = new List<string>();
            HtmlWeb web = new HtmlWeb();
            HtmlDocument doc;

            for (int i = 1; i < 75; i++)
            {
                string link = "http://www.pcmag.com/products/1565/" + i;
                doc = web.Load(link);

                foreach (HtmlNode node in doc.DocumentNode.SelectNodes("//label[@for='compare-check-1']/a"))
                {
                    string specLink = node.Attributes["href"].Value;

                    if (specLink != null)
                    {
                        dataPages.Add(specLink);
                    }

                    Random r = new Random();
                    Thread.Sleep(r.Next(100, 500));
                }
                Console.WriteLine(i * 10 + " Spec links found.");

            }

            return dataPages;
        }

        private static List<string> NavigateToSpecs (List<string> dataPages)
        {
            var specPages = new List<string>();
            HtmlWeb web = new HtmlWeb();
            HtmlDocument doc;
            string rootURL = "http://www.pcmag.com";
            int i = 1;

            foreach (string dataPage in dataPages)
            {
                doc = web.Load(dataPage);
                HtmlNode specPageNode = doc.DocumentNode.SelectSingleNode("//a[@data-link-name='Tab | Specs']");

                if (specPageNode == null)
                {
                    continue;
                }

                string specPageLink = specPageNode.Attributes["href"].Value;

                if (specPageLink != null)
                {
                    specPages.Add(rootURL + specPageLink);
                }

                Random r = new Random();
                Thread.Sleep(r.Next(300, 400));
                Console.WriteLine(i + " Spec pages confirmed.");
                i++;
            }
            return specPages;
        }

        private static void ExtractSpecs(List<string> specPages)
        {
            HtmlWeb web = new HtmlWeb();
            HtmlDocument doc;
            int laptopIndex = 1;

            foreach (string specPage in specPages)
            {
                doc = web.Load(specPage);
                int attributeIndex = 1;
                
                HtmlNode MSRPnode = doc.DocumentNode.SelectSingleNode("//span[@id='msrp-value']");
                if (MSRPnode == null)
                {
                    continue;
                }

                string MSRP = MSRPnode.InnerText;
                if (MSRP != null)
                {
                    oSheet.Cells[laptopIndex, 3] = MSRP;
                }

                HtmlNode nameNode = doc.DocumentNode.SelectSingleNode("//h2[@class='heading item fn']");
                if (nameNode == null)
                {
                    continue;
                }

                string name = nameNode.InnerText;
                if (name != null)
                {
                    oSheet.Cells[laptopIndex, 2] = name;
                }
                
                HtmlNode imgNode = doc.DocumentNode.SelectSingleNode("//a[@data-link-name='View Gallery | Primary Image Area']/img");
                if (imgNode == null) 
                {
                    continue;
                }

                string imgSrc = imgNode.Attributes["src"].Value;
                if (imgSrc != null)
                {
                    Console.WriteLine("Found the image.");
                    using (WebClient client = new WebClient())
                    {
                        client.DownloadFile(imgSrc, @"c:\Laptops\img" + laptopIndex + ".png");
                    }
                }
                

                oSheet.Cells[laptopIndex, 1] = laptopIndex;     

                HtmlNodeCollection nodes = doc.DocumentNode.SelectNodes("//td[@class='spectd']");
                if (nodes == null)
                {
                    continue;
                }

                foreach (HtmlNode node in nodes)
                {
                    if (node == null)
                    {
                        continue;
                    }

                    if (node.ChildNodes != null && node.ChildNodes.Count != 0)
                    {
                        string attribute = node.ChildNodes[0].InnerText.Trim(':');
                        string data = node.ChildNodes[0].NextSibling.InnerText.Substring(1);
                        PopulateExcelSheet(attribute, data, laptopIndex);
                    }
                    attributeIndex++;
                }
                Random r = new Random();
                Thread.Sleep(r.Next(250, 650));
                laptopIndex++;
            }
        }

        private static void PopulateExcelSheet(string attribute, string data, int index)
        {
            switch (attribute)
            {
                case "Type":
                    oSheet.Cells[index, 4] = data;
                    break;
                case "Operating System":
                    oSheet.Cells[index, 5] = data;
                    break;
                case "Processor Name":
                    oSheet.Cells[index, 6] = data;
                    break;
                case "Processor Speed":
                    oSheet.Cells[index, 7] = data;
                    break;
                case "Graphics Card":
                    oSheet.Cells[index, 8] = data;
                    break;
                case "Graphics Memory":
                    oSheet.Cells[index, 9] = data;
                    break;
                case "Screen Size":
                    oSheet.Cells[index, 10] = data;
                    break;
                case "Screen Type":
                    oSheet.Cells[index, 11] = data;
                    break;
                case "Native Resolution":
                    oSheet.Cells[index, 12] = data;
                    break;
                case "RAM":
                    oSheet.Cells[index, 13] = data;
                    break;
                case "Storage Capacity (as Tested)":
                    oSheet.Cells[index, 14] = data;
                    break;
                case "Rotation Speed":
                    oSheet.Cells[index, 15] = data;
                    break;            
            }
        }


    }
}
