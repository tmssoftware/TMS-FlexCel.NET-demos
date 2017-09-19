using System;
using System.IO;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using FlexCel.Core;
using FlexCel.Report;
using FlexCel.XlsAdapter;
using Windows.Storage;
using FlexCel.Render;

namespace LangWars
{
    class ReportGenerator
    {
        public const string TempHtmlRelFolder = "htm";
        public const string TempHtmlName = "langwars.html";
        public const string TempXlsName = "langwars.xls";

        public async Task<bool> TryCreateReport(bool Offline, Action StartProgress, Action EndProgress, Action LoadHTML, Action<string> SetHTML)
        {
            try
            {
                StartProgress();
                try
                {
                    await Task.Run(() => CreateReport(Offline));
                }
                finally
                {
                    EndProgress();
                }
            }
            catch (Exception ex)
            {
                SetHTML("<html>" + WebUtility.HtmlEncode(ex.Message) + "</html>");
                return false;
            }

            LoadHTML();
            return true;
        }

        async Task CreateReport(bool Offline)
        {
            var Langs = Offline ? await LoadData() : await FetchData();
            var xls = await RunReport(Langs);
            xls.Save(Path.Combine(TempXlsPath.Path, TempXlsName)); //we save it to share it and to display it on startup.
            await GenerateHTML(xls);
        }

        async Task<LangDataList> FetchData()
        {
            const string url = "https://api.stackexchange.com/2.1/tags?order=desc&sort=popular&site=stackoverflow&pagesize=5";

            var handler = new HttpClientHandler();
            if (handler.SupportsAutomaticDecompression)
            {
                handler.AutomaticDecompression = DecompressionMethods.GZip |
                                                 DecompressionMethods.Deflate;
            }
            var client = new HttpClient(handler);
            var response = await client.SendAsync(new HttpRequestMessage(HttpMethod.Get, new Uri(url)));

            string s = await response.Content.ReadAsStringAsync();
            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(LangDataList));

            return (LangDataList)ser.ReadObject(await response.Content.ReadAsStreamAsync());
        }

        async Task<LangDataList> LoadData()
        {
            var offlineFile = await StorageFile.GetFileFromApplicationUriAsync(new Uri("ms-appx:///Assets/Data/OfflineData.txt"));
            using (var offlineData = await offlineFile.OpenStreamForReadAsync())
            {
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(LangDataList));
                return (LangDataList)ser.ReadObject(offlineData);
            }
        }

        async Task<ExcelFile> RunReport(LangDataList langs)
        {
            ExcelFile Result = new XlsFile(true);
            var template = await StorageFile.GetFileFromApplicationUriAsync(new Uri("ms-appx:///Assets/Templates/report.template.xls"));
            using (var templateStream = await template.OpenStreamForReadAsync())
            {
                Result.Open(templateStream);
                using (FlexCelReport fr = new FlexCelReport(true))
                {
                    fr.AddTable("lang", langs.items);
                    fr.Run(Result);
                }
            }
            return Result;
        }

        public static StorageFolder TempXlsPath
        {
            get
            {
                return ApplicationData.Current.LocalFolder;
            }
        }

        public async static Task<StorageFolder> TempHtmlPath()
        {
            return await ApplicationData.Current.LocalFolder.CreateFolderAsync(TempHtmlRelFolder, CreationCollisionOption.OpenIfExists);
        }

        public async static Task<string> TempHtmlPathString()
        {
            var folder = await TempHtmlPath();
            return folder.Path;
        }

        async Task GenerateHTML(ExcelFile xls)
        {

            using (FlexCelHtmlExport html = new FlexCelHtmlExport(xls, true))
            {
                html.SavedImagesFormat = THtmlImageFormat.Svg; //vectorial so it can zoom. 
                html.EmbedImages = true;  

                //see http://blogs.windows.com/windows_phone/b/wpdev/archive/2011/03/14/managing-the-windows-phone-browser-viewport.aspx
                html.ExtraInfo.Meta = new string[] { "<meta name=\"viewport\" content=\"width=device-width\" />" };
                html.Export(Path.Combine(await TempHtmlPathString(), TempHtmlName), ".");
            }

        }
    }
}
