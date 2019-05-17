using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Configuration;
using System.Windows.Forms;
using System.Threading;


namespace PdfMapCreator
{
    public class HtmlToImage
    {
        public static List<Image> Img { get; set; } = new List<Image>();

        public static void ImageCapture(double latitude, double longitude)
        {
            string url = ConfigurationManager.AppSettings.Get("url");
            url = url.Replace("{latitude}", latitude.ToString());
            url = url.Replace("{longitude}", longitude.ToString());            

            int brainlessCounter = 0;
            int speedFactor = Convert.ToInt32(ConfigurationManager.AppSettings.Get("speedFactor"));

            while (brainlessCounter < speedFactor)
            {
                using (WebBrowser browser = new WebBrowser())
                {
                    browser.ScriptErrorsSuppressed = true;

                    browser.ScrollBarsEnabled = false;

                    browser.AllowNavigation = false;

                    browser.Width = 1900;

                    browser.Height = 1900;

                    browser.Navigate(url);

                    browser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(DocumentCompleted);

                    while (browser.ReadyState != WebBrowserReadyState.Complete)
                    {
                        Application.DoEvents();
                    }
                }

                if (brainlessCounter < (speedFactor-1))
                {
                    Img.RemoveAt(Img.Count - 1);
                }
                brainlessCounter++;
            }
                        
        }

        public static void DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser browser = sender as WebBrowser;

            WebBrowserReadyState loadStatus = default(WebBrowserReadyState);
            int waittime = 100;
            int counter = 0;
            while (true)
            {
                loadStatus = browser.ReadyState;
                Application.DoEvents();
                Thread.Sleep(1);

                if (counter > waittime /*|| (loadStatus == WebBrowserReadyState.Complete)*/)
                {
                    break;
                }
                counter++;
            }

            counter = 0;
            while (true)
            {
                loadStatus = browser.ReadyState;
                Application.DoEvents();

                if (loadStatus == WebBrowserReadyState.Complete)
                {
                    Bitmap screenshot = new Bitmap(browser.Width, browser.Height);
                    //ImageNativeMethods.GetImage(browser.ActiveXInstance, screenshot, Color.White);                  

                    browser.DrawToBitmap(screenshot, new Rectangle(0, 0, browser.Width, browser.Height));

                    using (MemoryStream stream = new MemoryStream())
                    {
                        screenshot.Save(stream, System.Drawing.Imaging.ImageFormat.Jpeg);

                        byte[] bytes = stream.ToArray();

                        Img.Add((Bitmap)new ImageConverter().ConvertFrom(bytes));

                        screenshot.Dispose();
                    }

                    break;
                }
                counter++;
            }
        }
       
        //private async static Task PageLoad(int TimeOut, WebBrowser browser)
        //{
        //    TaskCompletionSource<bool> PageLoaded = null;
        //    PageLoaded = new TaskCompletionSource<bool>();
        //    int TimeElapsed = 0;
        //    browser.DocumentCompleted += (s, e) =>
        //    {
        //        if (browser.ReadyState != WebBrowserReadyState.Complete) return;
        //        if (PageLoaded.Task.IsCompleted) return;
        //        string a = PageLoaded.Task.Status.ToString();
        //        PageLoaded.SetResult(true);
        //        string b = PageLoaded.Task.Status.ToString();

        //    };

        //    while (PageLoaded.Task.Status != TaskStatus.RanToCompletion)
        //    {
        //        await Task.Delay(10);
        //        TimeElapsed++;
        //        if (TimeElapsed >= TimeOut * 100) PageLoaded.TrySetResult(true);
        //    }
        //}
    }
}
