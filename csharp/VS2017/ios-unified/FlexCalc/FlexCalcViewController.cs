using System;
using System.Drawing;
using Foundation;
using UIKit;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System.IO;
using ObjCRuntime;

namespace FlexCalc
{
    public partial class FlexCalcViewController : UICollectionViewController
    {
        static NSString CalcCellId = new NSString ("CalcCell");
        string FileName
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "FlexCalc.xls"); }
        }

        ExcelFile xls;
        CalcCell ActiveEdit;
        RectangleF keybRect;

        static bool UserInterfaceIdiomIsPhone
        {
            get { return UIDevice.CurrentDevice.UserInterfaceIdiom == UIUserInterfaceIdiom.Phone; }
        }

        public FlexCalcViewController(IntPtr handle) : base (handle)
        {
        }

        public override void DidReceiveMemoryWarning()
        {
            // Releases the view if it doesn't have a superview.
            base.DidReceiveMemoryWarning();
			
            // Release any cached data, images, etc that aren't in use.
        }


        #region View lifecycle


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

        public override void ViewDidLoad ()
        {
            base.ViewDidLoad ();

            xls = new XlsFile(true);
            if (File.Exists(FileName))
                xls.Open(FileName);
            else
                CreateInitialFile(xls);

            CollectionView.RegisterClassForCell (typeof(CalcCell), CalcCellId);
            FixKeyboard();
           
            var Layout = new UICollectionViewFlowLayout();
            var w = CalcCellWidth(CollectionView.Bounds.Width);
            Layout.ItemSize = new SizeF(w, 50);
            CollectionView.SetCollectionViewLayout(Layout, false);

            //register the keyboard so we make the text fields visible
            NSNotificationCenter.DefaultCenter.AddObserver(UIKeyboard.DidShowNotification,  (x) =>
            {
                if (ActiveEdit == null) return;
                var keybRectObj = x.UserInfo.ObjectForKey(UIKeyboard.FrameBeginUserInfoKey);
                keybRect = ((NSValue)keybRectObj).RectangleFValue;
                FixKeyboard();


                var FrameRec = View.Frame;
                FrameRec.Height -=  keybRect.Height;
                if (!FrameRec.Contains(ActiveEdit.Frame.Location))
                {
                    this.CollectionView.ScrollRectToVisible(ActiveEdit.Frame, true);
                }
            });

            NSNotificationCenter.DefaultCenter.AddObserver(UIKeyboard.DidHideNotification, (x) =>
            {
                keybRect = new RectangleF(0, 0, 0, 0);
                FixKeyboard();
              
            });
          
        }

        void CreateInitialFile(ExcelFile xls)
        {
            xls.NewFile(1);
            xls.SetCellValue(1, 1, 7);
            xls.SetCellValue(2, 1, 5);
            xls.SetCellValue(3,1, new TFormula("=sum(a1:a2)^2"));
            xls.Recalc();
        }

        public void SaveConfig()
        {
            if (xls != null) xls.Save(FileName);
        }

        public override void DidRotate(UIInterfaceOrientation fromInterfaceOrientation)
        {
            FixKeyboard();
            var Layout = new UICollectionViewFlowLayout();
            var w = CalcCellWidth(CollectionView.Bounds.Width);
            Layout.ItemSize = new SizeF(w, 50);
            CollectionView.SetCollectionViewLayout(Layout, true);

        }

        void FixKeyboard()
        {
            if (InterfaceOrientation == UIInterfaceOrientation.Portrait || InterfaceOrientation == UIInterfaceOrientation.PortraitUpsideDown)
            {
                CollectionView.ContentInset = new UIEdgeInsets(20, 0, keybRect.Height + 40, 0);
            }
            else
            {
                CollectionView.ContentInset = new UIEdgeInsets(20, 0, keybRect.Width + 40, 0);
            }
        }

        public float CalcCellWidth(nfloat wd)
        {
            return 320;
        }

        public override UICollectionViewCell GetCell(UICollectionView collectionView, NSIndexPath indexPath)
        {
            var cell = (CalcCell)collectionView.DequeueReusableCell (CalcCellId, indexPath);
            cell.XlsRow = indexPath.Row + 1;
            cell.Heading = new TCellAddress(indexPath.Row + 1, 1).CellRef;
            cell.Content = GetCellOrFormula(indexPath.Row + 1);
            cell.ActiveCell = x => ActiveEdit = x;
            cell.Result = xls.GetStringFromCell(indexPath.Row + 1, 1);
            cell.OnRefresh = RefreshData;
            return cell;
        }

        string GetCellOrFormula(int row)
        {
            object cell = xls.GetCellValue(row, 1);
            if (cell == null)
                return "";
            TFormula fmla = (cell as TFormula);
            if (fmla != null)
                return fmla.Text;

            return Convert.ToString(cell);
        }  

        private void RefreshData(CalcCell changedCell)
        {
            xls.SetCellFromString(changedCell.XlsRow, 1, changedCell.Content);
            xls.Recalc();

            foreach (CalcCell cell in CollectionView.VisibleCells)
            {
                cell.Result = xls.GetStringFromCell(cell.XlsRow, 1);
            }

        }

        public override nint NumberOfSections (UICollectionView collectionView)
        {
            return 1;
        }

        public override nint GetItemsCount (UICollectionView collectionView, nint section)
        {
            return 20;
        }

        public override void ItemHighlighted(UICollectionView collectionView, NSIndexPath indexPath)
        {
            collectionView.EndEditing(false);
        }


    }
}

