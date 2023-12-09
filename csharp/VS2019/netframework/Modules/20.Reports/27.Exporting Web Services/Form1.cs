/*This program is inspired on the progam by Mikhail Arkhipov
 * at http://blogs.msdn.com/mikhailarkhipov/archive/2004/08/12/213963.aspx
 * Thanks!
 */

/* UPDATE: This was patched with the info on 
 * http://weblogs.asp.net/jan/archive/2004/01/28/63771.aspx
 * to make it work.
 * 
 * Thanks again...
 * 
 * UPDATE 2!
 * The NOAA broke the service again, and it has not fixed it for more than a year. 
 * I give up. We will use http://www.webservicex.net/WeatherForecast.asmx instead.
 * The code for NOAA is still there on the SetupNOAA method, just not used so you can see it (and try it if it ever starts working again)
 *
 * UPDATE 3!
 * Now WebserviceX is not working, going back to NOAA. As you can see, it isn't very trustable that a webservice will be there in the
 * future, so this demo might not work in online mode when you try it. But you can always look at in in offline mode.
 * */


using System;
using System.Drawing;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Xml;
using System.Net;
using System.Threading;

using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;
using FlexCel.Render;

using ExportingWebServices.gov.weather.www;
using System.Globalization;

namespace ExportingWebServices
{
    /// <summary>
    /// An example that will read data from a webservice.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {
        Dictionary<string, LatLong> Cities = new Dictionary<string, LatLong>(StringComparer.CurrentCultureIgnoreCase);
        public mainForm()
        {
            InitializeComponent();
            LoadCities();
        }

        private void LoadCities()
        {
            XmlDocument xml = new XmlDocument();
            {
                string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;
                xml.Load(Path.Combine(DataPath, "cities.xml"));
                XmlNodeList latLonList = xml.GetElementsByTagName("latLonList");
                XmlNodeList cityNameList = xml.GetElementsByTagName("cityNameList");

                if (latLonList.Count != 1) throw new Exception("Invalid city list");
                if (cityNameList.Count != 1) throw new Exception("Invalid city list");

                string lats = latLonList.Item(0).InnerText;
                string cits = cityNameList.Item(0).InnerText;

                string[] latsParsed = lats.Split(' ');
                string[] citsParsed = cits.Split('|');

                if (citsParsed.Length != latsParsed.Length) throw new Exception("Invalid city list");

                edcity.BeginUpdate();
                try
                {
                    for (int i = 0; i < citsParsed.Length; i++)
                    {
                        string[] ll = latsParsed[i].Split(',');
                        if (ll.Length != 2) throw new Exception("Invalid city list");
                        Cities.Add(citsParsed[i], new LatLong(Convert.ToDecimal(ll[0], CultureInfo.InvariantCulture), Convert.ToDecimal(ll[1], CultureInfo.InvariantCulture)));
                        edcity.Items.Add(citsParsed[i]);
                    }

                    edcity.Text = "New York,NY";
                }
                finally
                {
                    edcity.EndUpdate();
                }

            }
        }

        private void Export(SaveFileDialog SaveDialog, bool ToPdf)
        {
            try
            {
                string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;

                //We will use a thread to connect, to avoid "freezing" the GUI
                WebConnectThread MyWebConnect = new WebConnectThread(reportStart, edcity.Text, DataPath, cbOffline.Checked, Cities);
                Thread WebConnect = new Thread(new ThreadStart(MyWebConnect.SetupNOAA));
                WebConnect.Start();
                using (ProgressDialog Pg = new ProgressDialog())
                {
                    Pg.ShowProgress(WebConnect);
                    if (MyWebConnect != null && MyWebConnect.MainException != null)
                    {
                        throw MyWebConnect.MainException;
                    }
                }


                if (SaveDialog.ShowDialog() == DialogResult.OK)
                {
                    if (ToPdf)
                    {
                        XlsFile xls = new XlsFile();
                        xls.Open(DataPath + "Exporting Web Services.template.xls");
                        reportStart.Run(xls);
                        using (FlexCelPdfExport PdfExport = new FlexCelPdfExport(xls, true))
                        {
                            PdfExport.Export(SaveDialog.FileName);
                        }
                    }
                    else
                    {
                        reportStart.Run(DataPath + "Exporting Web Services.template.xls", SaveDialog.FileName);
                    }

                    if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        using (Process p = new Process())
                        {               
                            p.StartInfo.FileName = SaveDialog.FileName;
                            p.StartInfo.UseShellExecute = true;
                            p.Start();
                        }              
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }


        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void btnExportPdf_Click(object sender, System.EventArgs e)
        {
            Export(saveFileDialogPdf, true);
        }

        private void btnExportXls_Click(object sender, System.EventArgs e)
        {
            Export(saveFileDialogXls, false);
        }

        private void edcity_KeyDown(object sender, KeyEventArgs e)
        {
            edcity.DroppedDown = false;
        }

    }

    struct LatLong
    {
        public decimal Latitude;
        public decimal Longitude;

        public LatLong(decimal aLatitude, decimal aLongitude)
        {
            Latitude = aLatitude;
            Longitude = aLongitude;
        }
    }

    class WebConnectThread
    {
        string CityName;
        string DataPath;
        bool UseOfflineData;
        Dictionary<string, LatLong> Cities;
        FlexCelReport ReportStart;
        private Exception FMainException;

        public WebConnectThread(FlexCelReport aReportStart, string aCityName, string aDataPath, bool aUseOfflineData, Dictionary<string, LatLong> aCities)
        {
            CityName = aCityName;
            DataPath = aDataPath;
            UseOfflineData = aUseOfflineData;
            ReportStart = aReportStart;
            Cities = aCities;
        }


        public void SetupNOAA()
        {
            try
            {
                SetupNOAA(ReportStart, CityName, DataPath, UseOfflineData, Cities);
            }
            catch (Exception ex)
            {
                FMainException = ex;
            }

        }

        public static void SetupNOAA(FlexCelReport reportStart, string CityName, string DataPath, bool UseOfflineData, Dictionary<string, LatLong> Cities)
        {
            LatLong CityCoords;
            GetCity(Cities, CityName, out CityCoords);
            reportStart.SetValue("Date", DateTime.Now);
            string forecasts;
            DateTime dtStart = DateTime.Now;

            if (UseOfflineData)
            {
                using (StreamReader fs = new StreamReader(Path.Combine(DataPath, "OfflineData.xml")))
                {
                    forecasts = fs.ReadToEnd();
                }
            }
            else
            {
                ndfdXML nd = new ndfdXML();
                forecasts = nd.NDFDgen(CityCoords.Latitude, CityCoords.Longitude, productType.glance, dtStart, dtStart.AddDays(7), unitType.m, new weatherParametersType());

#if(SAVEOFFLINEDATA)
                using (StreamWriter sw = new StreamWriter(Path.Combine(DataPath, "OfflineData.xml")))
                {
                    sw.Write(forecasts);
                }
#endif
            }

            if (String.IsNullOrEmpty(forecasts)) throw new Exception("Can't find the place " + CityName);

            DataSet ds = new DataSet();
            //Load the data into a dataset. On this web service, we cannot just call DataSet.ReadXml as the data is not on the correct format.
            XmlDocument xmlDoc = new XmlDocument();
            {
                xmlDoc.LoadXml(forecasts);
                XmlNodeList HighList = xmlDoc.SelectNodes("/dwml/data/parameters/temperature[@type='maximum']/value/text()");
                XmlNodeList LowList = xmlDoc.SelectNodes("/dwml/data/parameters/temperature[@type='minimum']/value/text()");
                XmlNodeList IconList = xmlDoc.SelectNodes("/dwml/data/parameters/conditions-icon/icon-link/text()");

                DataTable WeatherTable = ds.Tables.Add("Weather");

                WeatherTable.Columns.Add("Day", typeof(DateTime));
                WeatherTable.Columns.Add("Low", typeof(double));
                WeatherTable.Columns.Add("High", typeof(double));
                WeatherTable.Columns.Add("Icon", typeof(byte[]));

                for (int i = 0; i < Math.Min(Math.Min(HighList.Count, LowList.Count), IconList.Count); i++)
                {
                    WeatherTable.Rows.Add(new object[]{
                                                          dtStart.AddDays(i),
                                                          Convert.ToDouble(LowList[i].Value),
                                                          Convert.ToDouble(HighList[i].Value),
                                                          LoadIcon(IconList[i].Value, UseOfflineData, DataPath)});
                }
            }


            reportStart.AddTable(ds, TDisposeMode.DisposeAfterRun);
            reportStart.SetValue("Latitude", CityCoords.Latitude);
            reportStart.SetValue("Longitude", CityCoords.Longitude);
            reportStart.SetValue("Place", CityName);

        }

        private static void GetCity(Dictionary<string, LatLong> Cities, string CityName, out LatLong CityCoords)
        {
            if (!Cities.TryGetValue(CityName, out CityCoords)) throw new Exception("Can't find the city " + CityName);
        }

        internal Exception MainException
        {
            get
            {
                return FMainException;
            }
        }

        internal static byte[] LoadIcon(string url, bool useOfflineData, string dataPath)
        {
            if (url == null || url.Length == 0)
            {
                return null; //no icon for this image.
            }

            if (useOfflineData)
            {
                Uri u = new Uri(url);
                return LoadFileIcon(Path.Combine(dataPath, u.Segments[u.Segments.Length - 1]));
            }
            else
            {
#if (SAVEOFFLINEDATA)
                Uri u = new Uri(url);
                byte[] IconData = LoadWebIcon(url);
                using (FileStream fs = new FileStream(Path.Combine(dataPath, u.Segments[u.Segments.Length - 1]), FileMode.Create))
                {                    
                    fs.Write(IconData, 0, IconData.Length);
                }
#endif

                return LoadWebIcon(url);

            }
        }

        /// <summary>
        /// On a real implementation this should be cached.
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        internal static byte[] LoadWebIcon(string url)
        {
            using (WebClient wc = new WebClient())
            {
                wc.Headers.Add("user-agent", "FlexCel Webservice Example");
                return wc.DownloadData(url);
            }
        }

        private static byte[] LoadFileIcon(string filename)
        {
            using (FileStream fs = new FileStream(filename, FileMode.Open))
            {
                byte[] Result = new byte[fs.Length];
                fs.Read(Result, 0, Result.Length);
                return Result;
            }
        }


    }

}
