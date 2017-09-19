using System;
using System.Drawing;
using Foundation;
using UIKit;
using FlexCel.Render;
using FlexCel.XlsAdapter;
using System.IO;
using System.Collections.Generic;

namespace FlexView
{
    public partial class FlexViewViewController : UIViewController
    {
        static bool UserInterfaceIdiomIsPhone
        {
            get { return UIDevice.CurrentDevice.UserInterfaceIdiom == UIUserInterfaceIdiom.Phone; }
        }

        public FlexViewViewController(IntPtr handle) : base (handle)
        {
        }

        public override void DidReceiveMemoryWarning()
        {
            // Releases the view if it doesn't have a superview.
            base.DidReceiveMemoryWarning();
			
            // Release any cached data, images, etc that aren't in use.
        }

        #region View lifecycle

        public override void ViewDidLoad()
        {
            base.ViewDidLoad();
            Viewer.LoadHtmlString("<html><h1>Please share an xls or xlsx file <br>from other app into FlexView.<br/ >" + 
                                  "For help on how to use this example, please read the " + "" +
                                  "<a href=\"http://www.tmssoftware.com/flexcel/docs/net/FlexCelViewTutorial.pdf\">tutorial</a></h1></html>", null);
			
        }

        public override void ViewWillAppear(bool animated)
        {
            base.ViewWillAppear(animated);
        }

        public override void ViewDidAppear(bool animated)
        {
            base.ViewDidAppear(animated);
        }

        public override void ViewWillDisappear(bool animated)
        {
            base.ViewWillDisappear(animated);
        }

        public override void ViewDidDisappear(bool animated)
        {
            base.ViewDidDisappear(animated);
        }

        #endregion

        NSUrl XlsUrl;
        string XlsPath;
        string PdfPath;

        public bool Open(NSUrl url)
        {
            XlsUrl = url;
            XlsPath = url.Path;
            return Refresh();
        }

        private void RemoveOldPdf()
        {
            if (PdfPath != null)
            {
                try
                {
                    File.Delete(PdfPath);
                }
                catch
                {
                    //do nothing, this was just a cache that will get deleted anyway.
                }
                PdfPath = null;
            }
        }

        private bool Refresh()
        {
            RemoveOldPdf();
            try
            {
                XlsFile xls = new XlsFile(XlsPath);

                PdfPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.InternetCache), 
                    Path.ChangeExtension(Path.GetFileName(XlsUrl.Path), ".pdf"));

                using (FlexCelPdfExport pdf = new FlexCelPdfExport(xls, true))
                {
                    using (FileStream fs = new FileStream(PdfPath, FileMode.Create))
                    {
                        pdf.Export(fs);
                    }
                }
                Viewer.LoadRequest(new NSUrlRequest(NSUrl.FromFilename(PdfPath)));

                
            }
            catch (Exception ex)
            {
                Viewer.LoadHtmlString("<html>Error opening " + System.Security.SecurityElement.Escape(Path.GetFileName(XlsUrl.Path))
                    + "<br><br>" + System.Security.SecurityElement.Escape(ex.Message) + "</html>", null);
                return false;
            }
            return true;
        }

        partial void ShareClick(UIKit.UIBarButtonItem sender)
        {
            if (PdfPath == null) 
            {
                ShowHowToUse();
                return;
            }

            UIDocumentInteractionController docController = new UIDocumentInteractionController();
            docController.Url = NSUrl.FromFilename(PdfPath);
            docController.PresentOptionsMenu(ShareButton, true);

        }

        partial void RandomizeClick(UIKit.UIBarButtonItem sender)
        {
            if (XlsUrl == null) 
            {
                ShowHowToUse();
                return;
            }

            XlsFile xls = new XlsFile(XlsPath, true);
            //We'll go through all the numeric cells and make them random numbers
            Random rnd = new Random();

            for (int row = 1; row <= xls.RowCount; row++) 
            {
                for (int colIndex = 1; colIndex < xls.ColCountInRow(row); colIndex++) 
                {
                    int XF = -1;
                    object val = xls.GetCellValueIndexed(row, colIndex, ref XF);
                    if (val is double) xls.SetCellValue(row, xls.ColFromIndex(row, colIndex), rnd.Next());
                }
            }

            //We can't save to the original file, we don't have permissions.
            XlsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.InternetCache), 
                "tmpFlexCel" + Path.GetExtension(XlsUrl.Path));

            xls.Save(XlsPath);
            Refresh();
        }

        void ShowHowToUse()
        {
            using (UIAlertView Alert = new UIAlertView("Please open FlexView from other Application.", 
                "In order to use this example you need to go to another app"
                + " like dropbox or mail, and share an xls or xlsx file with FlexView."
                + " When you click \"Share\" and select \"FlexView\" in the other app,"
                + " the file will be converted to pdf and previewed with FlexView.", null, "Ok"))
            {  
                Alert.Show();
            }
        }
    }
}

