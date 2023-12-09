using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Xml;
using System.Net;
using System.Threading;
using System.Globalization;

using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;
using FlexCel.Render;

namespace HTML
{
    /// <summary>
    /// Shows the limited HTML support.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {

        public mainForm()
        {
            InitializeComponent();
        }

        private void Export(SaveFileDialog SaveDialog, bool ToPdf)
        {
            using (FlexCelReport reportStart = new FlexCelReport(true))
            {

                if (cbOffline.Checked && edCity.Text != "london") MessageBox.Show("Offline mode is selected, so we will show the data of london. The actual city you wrote will not be used unless you select online mode.", "Warning");
                try
                {
                    string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;
                    string OfflineDataPath = Path.Combine(DataPath, "OfflineData") + Path.DirectorySeparatorChar;

                    //We will use a thread to connect, to avoid "freezing" the GUI
                    WebConnectThread MyWebConnect = new WebConnectThread(reportStart, edCity.Text, OfflineDataPath, cbOffline.Checked);
                    Thread WebConnect = new Thread(new ThreadStart(MyWebConnect.LoadData));
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
                            xls.Open(DataPath + "HTML.template.xls");
                            reportStart.Run(xls);
                            using (FlexCelPdfExport PdfExport = new FlexCelPdfExport(xls, true))
                            {
                                PdfExport.Export(SaveDialog.FileName);
                            }
                        }
                        else
                        {
                            reportStart.Run(DataPath + "HTML.template.xls", SaveDialog.FileName);
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

        private void linkLabel1_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            using (Process p = new Process())
            {               
                p.StartInfo.FileName = (sender as LinkLabel).Text;
                p.StartInfo.UseShellExecute = true;
                p.Start();
            }              
        }
    }

    class WebConnectThread
    {
        string CityName;
        string DataPath;
        bool UseOfflineData;
        FlexCelReport ReportStart;
        private Exception FMainException;

        public WebConnectThread(FlexCelReport aReportStart, string aCityName, string aDataPath, bool aUseOfflineData)
        {
            CityName = aCityName;
            DataPath = aDataPath;
            UseOfflineData = aUseOfflineData;
            ReportStart = aReportStart;
        }


        /// <summary>
        /// This is the method we will call form a thread. It catches any internal exception.
        /// </summary>
        public void LoadData()
        {
            try
            {
                LoadData(ReportStart, CityName, DataPath, UseOfflineData);
            }
            catch (Exception ex)
            {
                FMainException = ex;
            }

        }

        public static void LoadData(FlexCelReport reportStart, string CityName, string DataPath, bool UseOfflineData)
        {
            reportStart.SetValue("Date", DateTime.Now);
            DataSet ds = new DataSet();
            ds.Locale = CultureInfo.InvariantCulture;
            ds.EnforceConstraints = false;
            ds.ReadXmlSchema(Path.Combine(DataPath, "TripSearchResponse.xsd"));
            ds.Tables["Result"].Columns.Add("ImageData", typeof(byte[])); //Add a column for the actual images.
            if (UseOfflineData)
            {
                ds.ReadXml(Path.Combine(DataPath, "OfflineData.xml"));
            }
            else
            {
                // Create the web request  
                string url = String.Format("http://travel.yahooapis.com/TripService/V1.1/tripSearch?appid=YahooDemo&query={0}&results=20", CityName);
                UriBuilder uri = new UriBuilder(url);
                HttpWebRequest request = WebRequest.Create(uri.Uri.AbsoluteUri) as HttpWebRequest;

                // Get response  
                using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
                {
                    // Load data into a dataset  
                    ds.ReadXml(response.GetResponseStream());
                }
            }

            if (ds.Tables["ResultSet"].Rows.Count <= 0) throw new Exception("Error loading the data.");
            if (Convert.ToInt32(ds.Tables["ResultSet"].Rows[0]["totalResultsReturned"]) <= 0) throw new Exception("There are no travel plans for this location");

            LoadImageData(ds, UseOfflineData, DataPath);

            /* Uncomment this code to create an offline image of the data.*/
#if (CreateOffline)
			ds.WriteXml(Path.Combine(DataPath, "OfflineData.xml"));
#endif

            reportStart.AddTable(ds);
        }

        internal Exception MainException
        {
            get
            {
                return FMainException;
            }
        }

        private static void LoadImageData(DataSet ds, bool UseOfflineData, string DataPath)
        {
            DataTable Images = ds.Tables["Image"];
            Images.PrimaryKey = new DataColumn[] { Images.Columns["Result_Id"] };
            foreach (DataRow dr in ds.Tables["Result"].Rows)
            {
                DataRow ImageRow = Images.Rows.Find(dr["Result_Id"]);

                if (ImageRow == null) continue;
                string url = Convert.ToString(ImageRow["Url"]);
                if (url != null && url.Length > 0)
                {
                    dr["ImageData"] = LoadIcon(url, UseOfflineData, DataPath);
                }
            }

        }

        internal static byte[] LoadIcon(string url, bool useOfflineData, string dataPath)
        {
            if (useOfflineData)
            {
                Uri u = new Uri(url);
                return LoadFileIcon(Path.Combine(dataPath, u.Segments[u.Segments.Length - 1]));
            }
            else
            {
                /* Uncomment this code to create an offline image of the data. */
#if (CreateOffline)
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
