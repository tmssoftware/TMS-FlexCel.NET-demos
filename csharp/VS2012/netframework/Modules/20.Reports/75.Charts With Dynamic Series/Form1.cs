using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;
using FlexCel.Demo.SharedData;


namespace ChartsWithDynamicSeries
{
    /// <summary>
    /// A report including charts which have a series per row.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {

        public mainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            AutoRun();
        }

        public void AutoRun()
        {
            using (FlexCelReport ordersReport = SharedData.CreateReport())
            {
                ordersReport.CustomizeChart += OrdersReport_CustomizeChart;
                string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    ordersReport.Run(DataPath + "Charts With Dynamic Series.template.xlsx", saveFileDialog1.FileName);

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

        private void OrdersReport_CustomizeChart(object sender, CustomizeChartEventArgs e)
        {
            if (e.ChartName == "Stock<#swap series>")
            {
                //In this event we will set the colors of the series depending on the product.
                //Let's image each product has an associated color that we want to use for its series.
                for (int subChart = 1; subChart <= e.Chart.SubchartCount; subChart++)
                {
                    for (int series = 1; series <= e.Chart.SeriesInSubchart(subChart); series++)
                    {
                        var seriesDef = e.Chart.GetSeriesInSubchart(subChart, series, true, true, true);
                        var seriesOptions = seriesDef.Options[-1];
                        seriesOptions.FillOptions = new ChartSeriesFillOptions(
                            new TShapeFill(true, new TSolidFill(ColorForProduct(series))), null, false, false);
                        e.Chart.SetSeriesInSubchart(subChart, series, seriesDef);
                    }
                }
            }
        }

        private static TDrawingColor ColorForProduct(int series)
        {
            return TDrawingColor.FromRgb((byte)((series * 24) % 255), (byte)((series * 32) % 255), (byte)((series * 16) % 255));
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }
    }

}
