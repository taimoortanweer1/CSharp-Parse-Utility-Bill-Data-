using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void loadweb_Click(object sender, EventArgs e)
        {
 
            WebClient webClient = new WebClient();

            int z = 0;
            int row = 2;
            var workbook = new XLWorkbook();

            workbook.AddWorksheet("bill");
            var ws = workbook.Worksheet("bill");

            double consumerID = 6141250664000;

            for (int count = 1; count < 2000; count++, row++)
            {
                string id = z.ToString() + consumerID.ToString();
                //string page = webClient.DownloadString("http://210.56.23.106:888/iescobill/general/06141250664000");
                string page = webClient.DownloadString("http://210.56.23.106:888/iescobill/general/" + id);


                try
                {
            
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(page);

                    List<List<string>> table = doc.DocumentNode.SelectNodes("//table")
                                .Descendants("tr")
                                .Skip(1)
                                .Where(tr => tr.Elements("td").Count() > 1)
                                .Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToList())
                                .ToList();


                    ws.Cell("A1").Value = "id";
                    ws.Cell("B1").Value = "load";
                    ws.Cell("C1").Value = "feeder";
                    ws.Cell("D1").Value = "name";
                                    

                    //table[2] [2]--> load
                    //table[2] [3]--> id

                    //table[7] [1]-->feedername
                    //table[10] [0]--> name and address

                    //id ---- load---feeder name --- name and address -- 

                    int col = 1;
                    int tableCount = 0;
                    foreach (List<string> item in table)
                    {

                        if (tableCount < 16)
                        {
                            if (tableCount == 2)
                            {

                                //copy meter load data
                                ws.Cell(row, col).Value = item[3].ToString();
                                col++;

                                //copy user id 
                                ws.Cell(row, col).Value = item[2].ToString();
                                col++;
                            }

                            if (tableCount == 7)
                            {

                                ws.Cell(row, col).Value = item[1].ToString();
                                col++;
                            }

                            if (tableCount == 10)
                            {

                                ws.Cell(row, col).Value = "I-10/1";
                                col++;
                            }
                            tableCount++;
                        }

                        else if (tableCount > 27)
                        {
                            tableCount++;
                            break;
                        }

                        else
                        {
                            try
                            {
                                //month 
                                ws.Cell(1, col).Value = item[0].ToString();

                                //no of units
                                ws.Cell(row, col).Value = item[1].ToString();

                                //cost
                                //ws.Cell(row, col).Value = item[0].ToString();
                                col++;


                            }
                            catch { }
                            tableCount++;
                        }

                    }

                    

                    label1.Text = count.ToString();
                    workbook.SaveAs("data_with_units_1.xlsx");
                    

                    System.Threading.Thread.Sleep(5000);

                }
                catch
                {

                }
                Console.Write(count);
                consumerID = consumerID - 100; //for next iteration
            }//for loop


            //WebClient webClient = new WebClient();

            //int z = 0;
            //int row = 2;
            var workbook1 = new XLWorkbook();
            z = 0;
            row = 2;

            workbook1.AddWorksheet("bill");
            ws = workbook1.Worksheet("bill");

            consumerID = 6141250664000;

            for (int count = 1; count < 2000; count++, row++)
            {
                string id = z.ToString() + consumerID.ToString();
                //string page = webClient.DownloadString("http://210.56.23.106:888/iescobill/general/06141250664000");
                string page = webClient.DownloadString("http://210.56.23.106:888/iescobill/general/" + id);


                try
                {
            
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(page);

                    List<List<string>> table = doc.DocumentNode.SelectNodes("//table")
                                .Descendants("tr")
                                .Skip(1)
                                .Where(tr => tr.Elements("td").Count() > 1)
                                .Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToList())
                                .ToList();


                    ws.Cell("A1").Value = "id";
                    ws.Cell("B1").Value = "load";
                    ws.Cell("C1").Value = "feeder";
                    ws.Cell("D1").Value = "name";


                    //table[2] [2]--> load
                    //table[2] [3]--> id

                    //table[7] [1]-->feedername
                    //table[10] [0]--> name and address

                    //id ---- load---feeder name --- name and address -- 

                    int col = 1;
                    int tableCount = 0;
                    foreach (List<string> item in table)
                    {

                        if (tableCount < 16)
                        {
                            if (tableCount == 2)
                            {

                                //copy meter load data
                                ws.Cell(row, col).Value = item[3].ToString();
                                col++;

                                //copy user id 
                                ws.Cell(row, col).Value = item[2].ToString();
                                col++;
                            }

                            if (tableCount == 7)
                            {

                                ws.Cell(row, col).Value = item[1].ToString();
                                col++;
                            }

                            if (tableCount == 10)
                            {

                                ws.Cell(row, col).Value = "I-10/1";
                                col++;
                            }
                            tableCount++;
                        }

                        else if (tableCount > 27)
                        {
                            tableCount++;
                            break;
                        }

                        else
                        {
                            try
                            {
                                //month 
                                ws.Cell(1, col).Value = item[0].ToString();

                                //no of units
                                ws.Cell(row, col).Value = item[1].ToString();

                                //cost
                                //ws.Cell(row, col).Value = item[0].ToString();
                                col++;


                            }
                            catch { }
                            tableCount++;
                        }

                    }



                    label1.Text = count.ToString();
                    workbook1.SaveAs("data_with_units_2.xlsx");

                    System.Threading.Thread.Sleep(5000);

                }
                catch
                {

                }
                Console.Write(count);
                consumerID = consumerID + 100; //for next iteration
            }//for loop

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
