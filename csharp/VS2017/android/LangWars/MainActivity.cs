using System;
using Android.App;
using Android.Content;
using Android.Runtime;
using Android.Views;
using Android.Widget;
using Android.OS;
using System.Net;
using System.Runtime.Serialization.Json;
using System.IO;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;
using FlexCel.Render;
using Android.Webkit;
using System.Threading.Tasks;

namespace LangWars
{
    [Activity (Label = "LangWars", MainLauncher = true)]
    public class MainActivity : Activity
    {
        WebView ResultsWindow;
        CheckBox OnlineSwitch;
        bool HTMLLoaded = false;


        async protected override void OnCreate(Bundle bundle)
        {
            base.OnCreate(bundle);

            // Set our view from the "main" layout resource
            SetContentView(Resource.Layout.Main);

            Button FightButton = FindViewById<Button>(Resource.Id.FightButton);
			
            FightButton.Click += async delegate
            {
                await TryCreateReport();
            };

            Button ShareButton = FindViewById<Button>(Resource.Id.ShareButton);

            ShareButton.Click += delegate
            {
                SendFile();
            };

            ResultsWindow = FindViewById<WebView>(Resource.Id.ResultsWindow);
            OnlineSwitch = FindViewById<CheckBox>(Resource.Id.OnlineSwitch);
            await DisplayLastReport();
        }


        async Task<bool> TryCreateReport()
        {
            try
            {
                ProgressDialog progress = new ProgressDialog(this);
                progress.SetMessage("Loading...");
                progress.SetCancelable(false);
                progress.Show();
                try
                {
                    bool Offline = !OnlineSwitch.Checked;
                    await Task.Run(()=>CreateReport(Offline));
                }
                finally
                {
                    progress.Dismiss();
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
            if (TempXlsPath != null) xls.Save(TempXlsPath); //we save it to share it and to display it on startup.
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
            using (var offlineData = Assets.Open("OfflineData.txt"))
            {
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(LangDataList));
                return (LangDataList)ser.ReadObject(offlineData);
            }
        }

        ExcelFile RunReport(LangDataList langs)
        {
            ExcelFile Result = new XlsFile(true);
            using(var template = Assets.Open("report.template.xls"))
            {
                //we can't load directly from the asset stream, as we need a seekable stream.
                using (var memtemplate = new MemoryStream())
                {
                    template.CopyTo(memtemplate);
                    memtemplate.Position = 0;
                    Result.Open(memtemplate);
                }
            }
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
                    ApplicationContext.CacheDir.AbsolutePath, 
                    "langwars.html"); 
            } 
        }

        string TempXlsPath
        { 
            get
            {
                String state = Android.OS.Environment.ExternalStorageState;

                if (state == Android.OS.Environment.MediaMounted)
                {
                    return Path.Combine(
                        ApplicationContext.GetExternalFilesDir(null).AbsolutePath, //Note that we use EXTERNAL dir, because the other apps won't be able to read the internal cache.
                        "langwars.xls"); 
                }

                return null; //We can't save.
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
                ResultsWindow.LoadUrl(new Uri(TempHtmlPath).AbsoluteUri);
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
            ResultsWindow.LoadData(html, "text/html; charset=UTF-8", null);
            HTMLLoaded = false;
        }

        async void SendFile()
        {
            if (TempXlsPath == null)
            {
                new AlertDialog.Builder(this)
                    .SetTitle("Error")
                        .SetMessage("It is not possible to share the file because the external storage is not available.")
                        .Show();
                return;
            }

            if (!File.Exists(TempXlsPath))
            {
                if (!await TryCreateReport()) return;
            }

            Intent Sender = new Intent(Intent.ActionSend);
            Sender.SetType("application/vnd.ms-excel");
            Sender.PutExtra(Intent.ExtraStream, Android.Net.Uri.Parse(new Uri(TempXlsPath).AbsoluteUri));
            StartActivity(Intent.CreateChooser(Sender, "Select application"));
        }
    }
}


