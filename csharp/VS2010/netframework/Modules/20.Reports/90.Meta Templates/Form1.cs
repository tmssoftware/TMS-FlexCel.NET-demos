using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Data.OleDb;
using System.Threading;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;

using System.Xml;


namespace MetaTemplates
{
    /// <summary>
    /// Templates that self-modify themselves before running.
    /// </summary>
	public partial class mainForm: System.Windows.Forms.Form
    {

        public mainForm()
        {
            InitializeComponent();
        }

        public struct FeedData
        {
            public string Name;
            public string Url;
            public string Logo;

            public FeedData(string aName, string aUrl, string aLogo)
            {
                Name = aName;
                Url = aUrl;
                Logo = aLogo;
            }

            public override string ToString()
            {
                return Name;
            }
        }

        private FeedData[] Feeds =
                {
                new FeedData("TMS", "https://www.tmssoftware.com/rss/tms.xml", "tms.gif"),
                new FeedData("MSDN","https://sxpdata.microsoft.com/feeds/3.0/msdntn/MSDNMagazine_enus", "msdn.jpg"),
                new FeedData("SLASHDOT" , "http://rss.slashdot.org/Slashdot/slashdot", "slashdot.gif")
                };


        private void button2_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private string DataPath
        {
            get
            {
                return Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;
            }
        }

        private void Export(DataSet data)
        {
            using (FlexCelReport Report = new FlexCelReport(true))
            {
                Report.AddTable(data);
                Report.SetValue("FeedName", ((FeedData)cbFeeds.SelectedValue).Name);
                Report.SetValue("FeedUrl", ((FeedData)cbFeeds.SelectedValue).Url);
                Report.SetValue("ShowCount", cbShowFeedCount.Checked);

                using (FileStream fs = new FileStream(Path.Combine(Path.Combine(DataPath, "logos"), ((FeedData)cbFeeds.SelectedValue).Logo), FileMode.Open))
                {
                    byte[] b = new byte[fs.Length];
                    fs.Read(b, 0, b.Length);
                    Report.SetValue("Logo", b);
                }
                Report.Run(DataPath + "Meta Templates.template.xls", saveFileDialog1.FileName);
            }

        }

        private void btnExportExcel_Click(object sender, System.EventArgs e)
        {

            using (DataSet data = new DataSet())
            {
                string LocalData = Path.Combine(Path.Combine(DataPath, "data"), ((FeedData)cbFeeds.SelectedValue).Name + ".xml");

                if (cbOffline.Checked)
                {
                    data.ReadXml(LocalData);
                }
                else
                {
                    //In a real world example, this should be done on a thread, as it is done in the HTML example.
                    //To keep things simple here ,we will just "freeze" the gui while downloading the data, without
                    //providing feedback to the user.
                    XmlTextReader FeedReader = new XmlTextReader(((FeedData)cbFeeds.SelectedValue).Url);
                    data.ReadXml(FeedReader);
                }

#if (SaveForOffline)
				data.WriteXml(LocalData);
#endif


                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Export(data);

                    if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        using (Process p = new Process())
                        {               
                            p.StartInfo.FileName = saveFileDialog1.FileName;
                            p.StartInfo.UseShellExecute = true;
                            p.Start();
                        }
                    }
                }
            }
        }

        private void mainForm_Load(object sender, System.EventArgs e)
        {
            cbFeeds.DataSource = Feeds;
            cbFeeds.SelectedIndex = 0;
        }

    }
}
