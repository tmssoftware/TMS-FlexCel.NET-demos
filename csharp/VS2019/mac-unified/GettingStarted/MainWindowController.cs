using System;
using System.Collections.Generic;
using System.Linq;
using Foundation;
using AppKit;
using FlexCel.XlsAdapter;
using FlexCel.Core;

namespace GettingStarted
{
    public partial class MainWindowController : AppKit.NSWindowController
    {
        #region Constructors
        // Called when created from unmanaged code
        public MainWindowController(IntPtr handle) : base (handle)
        {
            Initialize();
        }
        // Called when created directly from a XIB file
        [Export ("initWithCoder:")]
        public MainWindowController(NSCoder coder) : base (coder)
        {
            Initialize();
        }
        // Call to load from the XIB/NIB file
        public MainWindowController() : base ("MainWindow")
        {
            Initialize();
        }
        // Shared initialization code
        void Initialize()
        {
        }
        #endregion
        //strongly typed window accessor
        public new MainWindow Window
        {
            get
            {
                return (MainWindow)base.Window;
            }
        }
            
        partial void CreateFile(NSObject sender)
        {
            var xls = new XlsFile(1, true);
            xls.SetCellValue(1, 1, "Hello OSX Unified!");
            xls.SetCellValue(2, 1, new TFormula("=\"Make sure to \" & \"look at the Windows examples\""));
            xls.SetCellValue(3, 1, "for information on how to use FlexCel");
            xls.SetCellValue(5, 1, "Concepts are similar, so it doesn't make sense to repeat them all here.");

            xls.AutofitCol(1, false, 1.2);

            NSSavePanel SaveDialog = new NSSavePanel();
            {
                SaveDialog.Title = "Save file as...";
                SaveDialog.AllowedFileTypes = new string[] {"xlsx", "xls"};
                SaveDialog.BeginSheet(Window, 
                (x) =>
                {
                    if (SaveDialog.Url != null)
                    {
                        xls.Save(SaveDialog.Url.Path);
                    }
                });
            }

             
        }

    }
}
