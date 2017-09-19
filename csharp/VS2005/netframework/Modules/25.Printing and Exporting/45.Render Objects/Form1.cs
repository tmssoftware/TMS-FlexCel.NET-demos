using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Render;
using System.IO;
using System.Diagnostics;
using System.Reflection;

using System.Text;



namespace RenderObjects
{
    /// <summary>
    /// An Example on how to render a chart.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {
        public mainForm()
        {
            InitializeComponent();
            ResizeToolbar(mainToolbar);
        }

        private void ResizeToolbar(ToolStrip toolbar)
        {

            using (Graphics gr = CreateGraphics())
            {
                double xFactor = gr.DpiX / 96.0;
                double yFactor = gr.DpiY / 96.0;
                toolbar.ImageScalingSize = new Size((int)(24 * xFactor), (int)(24 * yFactor));
                toolbar.Width = 0; //force a recalc of the buttons.
            }
        }


        #region Global variables
        private XlsFile Xls;
        private TXlsNamedRange ValueRange;
        private double MinValue;
        private double MaxValue;
        private double StepValue;
        private double ActualValue;

        private int ChartIndex;
        private TShapeProperties ChartProps;
        #endregion


        private void InitApp()
        {
            Xls = new XlsFile();

            string TemplatePath = Path.Combine(Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), ".."), "templates") + Path.DirectorySeparatorChar;
            DirectoryInfo di = new DirectoryInfo(TemplatePath);
            FileInfo[] fi = di.GetFiles("*.xls");
            if (fi.Length == 0) throw new Exception("Sorry, no templates found in the templates folder.");

            cbTheme.Items.Clear();
            foreach (FileInfo f in fi)
            {
                cbTheme.Items.Add(new FileHolder(f.FullName));
            }

            cbTheme.SelectedIndex = 0;
        }

        private void LoadFile(string FileName)
        {
            Xls.Open(FileName);

            ActualValue = 0;

            ValueRange = Xls.GetNamedRange("Value", 0);
            if (ValueRange == null) throw new Exception("There is no range named \"value\" in the template");

            MinValue = ReadDoubleName("Minimum");
            MaxValue = ReadDoubleName("Maximum");
            StepValue = ReadDoubleName("Step");

            ChartIndex = -1;
            for (int i = 1; i <= Xls.ObjectCount; i++)
            {
                string ObjName = Xls.GetObjectName(i);
                if (String.Compare(ObjName, "DataChart", true) == 0)
                {
                    ChartIndex = i;
                    break;
                }
            }

            if (ChartIndex < 0) throw new Exception("There is no object named \"DataChart\" in the template");
            ChartProps = Xls.GetObjectProperties(ChartIndex, true);
        }

        private double ReadDoubleName(string Name)
        {
            TXlsCellRange Range = Xls.GetNamedRange(Name, 0);
            if (Range == null) throw new Exception("There is no range named " + Name + " in the template");

            object val = Xls.GetCellValue(Range.Top, Range.Left);
            if (!(val is Double)) throw new Exception("The range named " + Name + " does not contain a number");
            return (Double)val;
        }

        private void button2_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void updater_Tick(object sender, System.EventArgs e)
        {
            try
            {
                ActualValue += StepValue;
                if (ActualValue > MaxValue) ActualValue = MinValue;
                Xls.SetCellValue(ValueRange.Top, ValueRange.Left, ActualValue);
                Xls.Recalc();

                if (chartBox.Image != null) chartBox.Image.Dispose();
                chartBox.Image = GetChart();
            }
            catch (Exception ex)  //We don't want any dialog popping up every second.
            {
                labelError.Text = ex.Message;
                labelError.Dock = DockStyle.Fill;
                panelError.Dock = DockStyle.Fill;
                panelError.Visible = true;
                updater.Enabled = false;

            }
        }

        private Image GetChart()
        {
            //We could get the chart with the following command, 
            //but it would be fixed size. In this example we are going to be a little more complex.

            //Xls.RenderObject(ChartIndex);

            //A more complex way to retrieve the chart, to show how to use
            //all parameters in renderobject.

            TUIRectangle ImageDimensions;
            TPointF Origin;
            TUISize SizePixels;

            //First calculate the chart dimensions without actually rendering it. This is fast.
            Xls.RenderObject(ChartIndex, 96, ChartProps,
                SmoothingMode.AntiAlias, InterpolationMode.HighQualityBicubic, true, false,
                out Origin, out ImageDimensions, out SizePixels);

            double dpi = 96;  //default screen resolution
            if (SizePixels.Height > 0 && SizePixels.Width > 0)
            {
                double AspectX = (double)chartBox.Width / SizePixels.Width;
                double AspectY = (double)chartBox.Height / SizePixels.Height;

                double Aspect = Math.Max(AspectX, AspectY);
                //Make the dpi adjust the screen resolution and the size of the form.
                dpi = (double)(96 * Aspect);
                if (dpi < 20) dpi = 20;
                if (dpi > 500) dpi = 500;
            }

            return Xls.RenderObject(ChartIndex, dpi, ChartProps,
                SmoothingMode.AntiAlias, InterpolationMode.HighQualityBicubic, true, true,
                out Origin, out ImageDimensions, out SizePixels);


        }

        private void cbTheme_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            if (cbTheme.SelectedItem == null) return;
            LoadFile((cbTheme.SelectedItem as FileHolder).FullName);
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            if (Xls == null) InitApp();
            updater.Enabled = true;
            btnRun.Enabled = false;
            btnCancel.Enabled = true;
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            updater.Enabled = false;
            btnRun.Enabled = true;
            btnCancel.Enabled = false;
            panelError.Visible = false;
        }

    }

    internal class FileHolder
    {
        internal string FullName;
        private string Caption;

        internal FileHolder(string aFullName)
        {
            FullName = aFullName;
            Caption = Path.GetFileNameWithoutExtension(aFullName);
        }

        public override string ToString()
        {
            return Caption;
        }

    }
}
