using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Timers;

namespace Project.NET.ChartSenderService
{
    public partial class ChartSender : ServiceBase
    {
        public string apiKey = "PC55IMQVIAWCD1X9";
        private string url;
        private string chartPath = @"C:\ChartSender\chart.png";
        private string myEmail = "mySecretEmail";
        private string myPassword = "mySecretPassword";

        private string fromSymbol = "gbp";
        private string toSymbol = "usd";
        Dictionary<DateTime, String> dict = null;

        public ChartSender()
        {
            url = String.Format("https://www.alphavantage.co/query?function=FX_DAILY&from_symbol={0}&to_symbol={1}&apikey={2}", fromSymbol, toSymbol, apiKey);
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            //System.Diagnostics.Debugger.Launch();

            Timer _timer = new Timer(1 * 60 * 1000);
            _timer.Elapsed += new System.Timers.ElapsedEventHandler(timer_Elapsed);
            _timer.Start();
        }

        private void timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            RunService();
        }

        private void RunService()
        {
            GetData();
            CreateChart();
            SendChart();
        }

        private void CreateChart()
        {
            int sizeX = 1920;
            int sizeY = 1080;

            using (Bitmap bmp = new Bitmap(sizeX, sizeY))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {

                    double maxPrice = 0, minPrice = 100;
                    foreach (var obj in dict)
                    {
                        double value = Convert.ToDouble(obj.Value);
                        if (value > maxPrice) maxPrice = value;
                        if (value < minPrice) minPrice = value;
                    }

                    double priceStep = (maxPrice - minPrice) / 5;

                    int x_indent = 100;
                    int y_indent = 100;

                    Font font = new Font(FontFamily.GenericSansSerif, 20);
                    Brush brush = new SolidBrush(Color.Black);
                    Pen pen = new Pen(brush);
                    Pen pricePen = new Pen(Color.Blue, 3);

                    g.FillRectangle(new SolidBrush(Color.Beige), 0, 0, 1920, 1080);
                    g.DrawRectangle(new Pen(Color.Black, 2), x_indent, y_indent, sizeX - 2 * x_indent, sizeY - 2 * y_indent);


                    int distance = 0;
                    int dateDistance = 180;
                    for (double i = minPrice; i < maxPrice;)
                    {
                        g.DrawString(i.ToString().Substring(0, 5), font, brush, x_indent - 95, sizeY - distance - y_indent);
                        distance += dateDistance;

                        g.DrawLine(new Pen(Color.LightGray, 1), x_indent, sizeY - distance - y_indent + 20, sizeX - x_indent, sizeY - distance - y_indent + 20);

                        i += priceStep;
                    }


                    int countOfPoints = (sizeX - 2 * x_indent) / dict.Count;
                    distance = x_indent;
                    double lastValue = y_indent;

                    double mainCenter = (sizeY - 2 * y_indent) / 2;
                    double subCenter = (maxPrice + minPrice) / 2;
                    double ratio = (sizeY - 2 * y_indent) / (maxPrice - minPrice);

                    foreach (var obj in dict)
                    {
                        double subClose = Convert.ToDouble(obj.Value) - subCenter;
                        double result = mainCenter + subClose * ratio;

                        g.FillRectangle(brush, distance + 15, y_indent + (float)result, 7, 7);
                        DrawRotatedTextAt(g, 270, obj.Key.ToString().Substring(0, 6), sizeX - distance - 10, y_indent - 10, new Font(FontFamily.GenericSansSerif, 10), brush);

                        int distance2 = distance + countOfPoints;
                        g.DrawLine(pricePen, distance, (float)lastValue + y_indent, distance2, (float)result + y_indent);

                        distance += countOfPoints;
                        lastValue = result;
                    }

                    bmp.Save(chartPath);
                }
            }
        }

        private void DrawRotatedTextAt(Graphics gr, float angle, string txt, int x, int y, Font the_font, Brush the_brush)
        {
            GraphicsState state = gr.Save();
            gr.ResetTransform();
            gr.RotateTransform(angle);
            gr.TranslateTransform(x, y, MatrixOrder.Append);

            gr.DrawString(txt, the_font, the_brush, 0, 0);

            gr.Restore(state);
        }

        private void GetData()
        {
            string json = null;
            try
            {
                json = DownloadData();
            }
            catch (DownloadErrorException ex)
            {
                Console.WriteLine(ex + ". Repeating process\n");
                RunService();
                return;
            }

            var data = (JObject)JsonConvert.DeserializeObject(json);
            List<string> closePrices = JObject.Parse(json)["Time Series FX (Daily)"].Values().Select(p => p["4. close"].Value<string>().Replace(".", ",")).ToList();
            var jsonObj = JObject.Parse(json)["Time Series FX (Daily)"].Children();

            List<DateTime> datesList = new List<DateTime>();

            foreach (var val in jsonObj)
            {
                string date = Convert.ToString(val).Substring(1, 10);
                datesList.Add(Convert.ToDateTime(date));
            }

            dict = datesList.Zip(closePrices, (k, v) => new { k, v }).ToDictionary(x => x.k, x => x.v);
        }


        private void SendChart()
        {
            MailMessage mail = new MailMessage("p.matej.skolni@gmail.com", "p.matej.skolni@gmail.com", "Subject: Project .NET", "Body: Hi, your chart subscription");
            mail.Attachments.Add(new Attachment(chartPath));

            SmtpClient client = new SmtpClient("smtp.gmail.com", 587)
            {
                Credentials = new NetworkCredential(myEmail, myPassword),
                EnableSsl = true
            };

            client.Send(mail);
            client.Dispose();
            mail.Dispose();
            
            File.Delete(chartPath);
        }

        private string DownloadData()
        {
            string json = null;
            using (WebClient wc = new WebClient())
            {
                json = wc.DownloadString(url);
            }

            if (json.Contains("Error"))
            {
                Console.WriteLine("Download failed, repeating " + fromSymbol + toSymbol);
                throw new DownloadErrorException("Error while downloading file");
            }

            if (json.Contains("Note"))
            {
                Console.WriteLine("Download failed, repeating " + fromSymbol + toSymbol);
                System.Threading.Thread.Sleep(5000);
                throw new DownloadErrorException("Limit 5 API calls per minute, please wait");
            }

            return json;
        }





        protected override void OnStop()
        {
            Trace.WriteLine("Service stopped");
        }
    }
}
