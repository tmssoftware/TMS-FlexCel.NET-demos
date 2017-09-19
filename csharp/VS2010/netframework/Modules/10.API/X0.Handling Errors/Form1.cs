using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Text;

using FlexCel.Render;

namespace HandlingErrors
{
    /// <summary>
    /// How to handle non fatal errors with FlexCel.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {

        private FlexCelErrorEventHandler FlexCelTrace_OnErrorHandler;
        public mainForm()
        {
            InitializeComponent();

            //Create a list to hold error messages. Keeping all error messages in memory is normally not a good thing to do, 
            //but for this demo it is ok.
            ErrorList = new ArrayList();

            //Hook our error handler to FlexCel error handler.	
            FlexCelTrace_OnErrorHandler = new FlexCelErrorEventHandler(FlexCelTrace_OnError); //We will save the value of the delegate here so we can unhook the event on dispose.
            FlexCelTrace.OnError += FlexCelTrace_OnErrorHandler;
        }

        private ArrayList ErrorList;
        private static object ErrorListLock = new object(); //Used to lock ErrorList and ensure no more than one thread writes to it.

        private string PathToExe
        {
            get
            {
                return Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;
            }
        }


        private void button1_Click(object sender, System.EventArgs e)
        {
            ErrorList.Clear();
            errorBox.Text = "";

            try
            {
                DoThings();
            }
            catch (MyAbortException ex)
            {
                MessageBox.Show(ex.Message);
            }

            if (ErrorList.Count == 0) errorBox.Text = "No errors!";
            else
            {
                errorBox.Text = String.Format("There were {0} error messages" + Environment.NewLine, ErrorList.Count);
                foreach (string s in ErrorList)
                {
                    errorBox.AppendText(s + Environment.NewLine);
                }
            }
        }


        private void DoThings()
        {
            ExcelFile xls = new XlsFile(true);
            xls.NewFile(1);

            for (int r = 1; r < 2000; r++)
            {
                xls.InsertHPageBreak(r); //This won't throw an exception here, since FlexCel allows to have more than 1025 page breaks, but at the moment of saving. (since an xls file can't have more than that)
            }

            xls.SetCellValue(1, 1, "We have a page break on each row, so this will print/export as one row per page");
            xls.SetCellValue(2, 1, "??? ? ? ? ???? ????"); //Since we leave the font at arial, this won't show when exporting to pdf.

            TFlxFormat fmt = xls.GetDefaultFormat;
            fmt.Font.Name = "Arial Unicode MS";
            xls.SetCellValue(3, 1, "??? ? ? ? ???? ????", xls.AddFormat(fmt)); //this will display fine in the pdf.

            fmt.Font.Name = "ThisFontDoesntExists";
            xls.SetCellValue(4, 1, "This font doesn't exists", xls.AddFormat(fmt));

            //Tahoma doesn't have italic variant. See http://help.lockergnome.com/office/Tahoma-italic-ftopict705661.html
            //You shouldn't normally use Tahoma italics in a document. If we embedded the fonts in this pdf, the fake italics wouldn't work.
            fmt.Font.Name = "Tahoma";
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellValue(5, 1, "This is fake italics", xls.AddFormat(fmt));

            if (saveFileDialog1.ShowDialog() != DialogResult.OK) return;

            using (FlexCelPdfExport pdf = new FlexCelPdfExport(xls, true))
            {
                pdf.Export(Path.ChangeExtension(saveFileDialog1.FileName, ".pdf"));
            }

            xls.Save(saveFileDialog1.FileName + ".xls");
        }

        /// <summary>
        /// This is the generic event handler for non fatal errors. We hooked it in the mainForm constructor.
        /// </summary>
        /// <param name="e"></param>
        private void FlexCelTrace_OnError(TFlexCelErrorInfo e)
        {

            if (cbIgnoreFontErrors.Checked)
            {
                switch (e.Error)
                {
                    //Ignore this errors:
                    case FlexCelError.PdfFontNotFound:
                    case FlexCelError.PdfGlyphNotInFont:
                    case FlexCelError.PdfFauxBoldOrItalics:
                        return;
                }
            }


            //Normally tracing non fatal errors is a good idea. 
            //Depending on the listener of your trace object, you can redirect this to a log, the event viewer or wherever else.
            Trace.WriteLine(e.Message);

            //If we selected "Stop On Errors" we will abort file generation by throwing an exception that will be
            //catched in the main block.
            if (cbStopOnErrors.Checked)
            {
                throw new MyAbortException(e.Message);
            }

            //In this case this is a single thread app so locking is not really necessary,
            //but it is a good practice to always lock access to global objects in this error handler.
            //This event handler might me called from more than one thread, and you don't want to mess
            //the object collecting the messages (in this case ErrorList).
            lock (ErrorListLock)
            {
                ErrorList.Add(System.Threading.Thread.CurrentThread.Name + ": - " + e.Message);
            }
        }
    }

    /// <summary>
    /// A custom exception designed to notify us when a non fatal error must be aborted.
    /// </summary>
    public class MyAbortException: Exception
    {
        public MyAbortException(string aMessage) : base(aMessage) { }
    }
}
