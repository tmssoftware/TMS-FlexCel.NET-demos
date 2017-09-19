using System;
using System.Drawing;
using Foundation;
using UIKit;
using System.Collections.Generic;
using FlexCel.Core;
using System.Net;
using System.IO;
using System.Runtime.Serialization.Json;
using FlexCel.Report;
using FlexCel.Render;
using FlexCel.XlsAdapter;
using System.Threading.Tasks;

namespace LangWars
{
    public partial class LangWarsViewController : UIViewController
    {
        bool HTMLLoaded = false;

        static bool UserInterfaceIdiomIsPhone
        {
            get { return UIDevice.CurrentDevice.UserInterfaceIdiom == UIUserInterfaceIdiom.Phone; }
        }

        public LangWarsViewController(IntPtr handle) : base (handle)
        {
        }

        public override void DidReceiveMemoryWarning()
        {
            // Releases the view if it doesn't have a superview.
            base.DidReceiveMemoryWarning();
			
            // Release any cached data, images, etc that aren't in use.
        }

        async public override void ViewDidLoad()
        {
            base.ViewDidLoad();

            if (File.Exists(TempHtmlPath))
            {
                try
                {
                    await DisplayLastReport();
                }
                catch (Exception ex)
                {
                    SetHTML("<html>" + System.Security.SecurityElement.Escape(ex.Message) + "</html>");
                }

            }
			
            // Perform any additional setup after loading the view, typically from a nib.
        }

        async partial void FightClick(NSObject sender)
        {
            await TryCreateReport();

        }

        async Task<bool> TryCreateReport()
        {
            try
            {
                ProgressIndicator.StartAnimating();
                try
                {
                    bool Offline = OfflineSwitch.SelectedSegment == 0;
                    await Task.Run(()=>CreateReport(Offline));
                }
                finally
                {
                    ProgressIndicator.StopAnimating();
                }
            }
            catch (Exception ex)
            {
                SetHTML("<html>" + System.Security.SecurityElement.Escape(ex.Message) + "</html>");
                return false;
            }

            LoadHTML();
            return true;
        }

        async Task DisplayLastReport()
        {
            if (!File.Exists(TempHtmlPath))
            {
                await TryCreateReport();
                return;
            }

            LoadHTML();
        }

        void CreateReport(bool Offline)
        {
            var Langs = Offline? LoadData(): FetchData();
            var xls = RunReport(Langs);
            xls.Save(TempXlsPath); //we save it to share it and to display it on startup.
            GenerateHTML(xls);
        }

        LangDataList FetchData()
        {
            var httpWebRequest = (HttpWebRequest)WebRequest.Create("https://api.stackexchange.com/2.1/tags?order=desc&sort=popular&site=stackoverflow&pagesize=5");
            httpWebRequest.AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip;
            httpWebRequest.Method = "GET";
            var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();

            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(LangDataList));
            return (LangDataList)ser.ReadObject(httpResponse.GetResponseStream());
        }

        LangDataList LoadData()
        {
            using (var offlineData = new FileStream(Path.Combine(NSBundle.MainBundle.BundlePath, "OfflineData.txt"), FileMode.Open, FileAccess.Read))
            {
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(LangDataList));
                return (LangDataList)ser.ReadObject(offlineData);
            }
        }

        ExcelFile RunReport(LangDataList langs)
        {
            ExcelFile Result = new XlsFile(Path.Combine(NSBundle.MainBundle.BundlePath, "report.template.xls"), true);
            using (FlexCelReport fr = new FlexCelReport(true))
            {
                fr.AddTable("lang", langs.items);
                fr.Run(Result);
            }
            return Result;
        }

        string TempHtmlPath
        { 
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.InternetCache), 
                    "langwars.html"); 
            } 
        }

        string TempXlsPath
        { 
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.InternetCache), 
                    "langwars.xls"); 
            } 
        }
 
        void GenerateHTML(ExcelFile xls)
        {

            using (FlexCelHtmlExport html = new FlexCelHtmlExport(xls, true))
            {
                //If we were using png, we would have to set
                //a high resolution so this looks nice in high resolution displays.
                //html.ImageResolution = 326;

                //but we will use SVG, which is vectorial:
                html.HtmlVersion = THtmlVersion.Html_5;
                html.SavedImagesFormat = THtmlImageFormat.Svg;
                html.EmbedImages = true;
                               
                html.Export(TempHtmlPath, ".");
            }

        }

        void LoadHTML()
        {
            if (!HTMLLoaded)
            {
                ResultsWindow.LoadRequest(new NSMutableUrlRequest(NSUrl.FromFilename(TempHtmlPath), 
                                                                  NSUrlRequestCachePolicy.ReloadIgnoringLocalAndRemoteCacheData, 60));
                HTMLLoaded = true;
            }
            else
            {
                //WebView doesn't refresh the images if you just load the request again (even if ignoring the cache). So we need to call Reload instead
                ResultsWindow.Reload();
            }
        }

        void SetHTML(string html)
        {
            ResultsWindow.LoadHtmlString(html, null);
            HTMLLoaded = false;
        }

        async partial void ShareClick(NSObject sender)
        {
            if (!File.Exists(TempXlsPath))
            {
                if (!await TryCreateReport()) return;
            }
            UIDocumentInteractionController docController = new UIDocumentInteractionController();
            docController.Url = NSUrl.FromFilename(TempXlsPath);
            docController.PresentOptionsMenu((UIBarButtonItem) sender, true);

        }
    }
}

