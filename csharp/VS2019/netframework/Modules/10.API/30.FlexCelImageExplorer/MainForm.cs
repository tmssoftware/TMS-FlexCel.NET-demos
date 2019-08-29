using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;

using FlexCel.Core;
using FlexCel.XlsAdapter;

namespace FlexCelImageExplorer
{
    /// <summary>
    /// Image Explorer.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {

        public mainForm()
        {
            InitializeComponent();
            GetCurrencyManager.CurrentChanged += new EventHandler(CurrentRowChanged);
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

        private string CurrentFilename = null;
        private TCompressForm CompressForm;

        private CurrencyManager GetCurrencyManager
        {
            get
            {
                return (CurrencyManager)this.BindingContext[dataGrid.DataSource, dataGrid.DataMember];
            }
        }

        private int GetImagePos
        {
            get
            {
                return GetCurrencyManager.Position;
            }
        }

        private void btnExit_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void btnOpenFile_Click(object sender, System.EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            OpenFile(openFileDialog1.FileName);
            if (cbScanFolder.Checked) FillListBox();
        }

        private bool GetHasCrop(TImageProperties ImgProps)
        {
            return ImgProps.CropArea.CropFromLeft != 0 ||
                ImgProps.CropArea.CropFromRight != 0 ||
                ImgProps.CropArea.CropFromTop != 0 ||
                ImgProps.CropArea.CropFromBottom != 0;
        }

        private void FillListBox()
        {
            lblFolder.Text = "Files on folder: " + Path.GetDirectoryName(openFileDialog1.FileName);
            DirectoryInfo di = new DirectoryInfo(Path.GetDirectoryName(openFileDialog1.FileName));
            FileInfo[] Fi = di.GetFiles("*.xls");
            FilesListBox.Items.Clear();

            TImageInfo[] Files = new TImageInfo[Fi.Length];

            for (int k = 0; k < Fi.Length; k++)
            {
                FileInfo f = Fi[k];
                bool HasCrop = false;
                bool HasARGB = false;
                XlsFile x1 = new XlsFile();

                bool HasImages = false;

                try
                {
                    x1.Open(f.FullName);
                    for (int sheet = 1; sheet <= x1.SheetCount; sheet++)
                    {
                        x1.ActiveSheet = sheet;
                        for (int i = x1.ImageCount; i > 0; i--)
                        {
                            HasImages = true;
                            TImageProperties ip = x1.GetImageProperties(i);
                            if (!HasCrop) HasCrop = GetHasCrop(ip);

                            TXlsImgType imgType = TXlsImgType.Unknown;
                            using (MemoryStream ms = new MemoryStream())
                            {
                                x1.GetImage(i, ref imgType, ms);
                                FlexCel.Pdf.TPngInformation PngInfo = FlexCel.Pdf.TPdfPng.GetPngInfo(ms);
                                if (PngInfo != null)
                                {
                                    HasARGB = PngInfo.ColorType == 6;
                                }
                            }

                        }
                    }
                }
                catch (Exception)
                {
                    Files[k] = new TImageInfo(f, false, false, false, false);
                    continue;
                }

                Files[k] = new TImageInfo(f, true, HasCrop, HasImages, HasARGB);
            }

            FilesListBox.Items.AddRange(Files);
        }

        private void OpenFile(string FileName)
        {
            ImageDataTable.Rows.Clear();

            try
            {
                XlsFile Xls = new XlsFile(true);
                CurrentFilename = FileName;
                Xls.Open(FileName);

                for (int sheet = 1; sheet <= Xls.SheetCount; sheet++)
                {
                    Xls.ActiveSheet = sheet;
                    for (int i = Xls.ImageCount; i > 0; i--)
                    {
                        TXlsImgType ImageType = TXlsImgType.Unknown;
                        byte[] ImgBytes = Xls.GetImage(i, ref ImageType);
                        TImageProperties ImgProps = Xls.GetImageProperties(i);
                        object[] ImgData = new object[ImageDataTable.Columns.Count];
                        ImgData[0] = Xls.SheetName;
                        ImgData[1] = i;
                        ImgData[4] = ImageType.ToString();
                        ImgData[7] = Xls.GetImageName(i);
                        ImgData[8] = ImgBytes;
                        ImgData[9] = GetHasCrop(ImgProps);


                        using (MemoryStream ms = new MemoryStream(ImgBytes))
                        {
                            FlexCel.Pdf.TPngInformation PngInfo = FlexCel.Pdf.TPdfPng.GetPngInfo(ms);
                            if (PngInfo != null)
                            {
                                ImgData[2] = PngInfo.Width;
                                ImgData[3] = PngInfo.Height;
                                string s = String.Empty;
                                int bpp = 0;

                                if ((PngInfo.ColorType & 4) != 0)
                                {
                                    s += "ALPHA-";
                                    bpp = 1;
                                }
                                if ((PngInfo.ColorType & 2) == 0)
                                {
                                    s += "Grayscale -" + (1 << PngInfo.BitDepth).ToString() + " shades. ";
                                    bpp = 1;
                                }
                                else
                                {
                                    if ((PngInfo.ColorType & 1) == 0)
                                    {
                                        bpp += 3;
                                        s += "RGB - " + (PngInfo.BitDepth * (bpp)).ToString() + "bpp.  ";
                                    }
                                    else
                                    {
                                        s += "Indexed - " + (1 << PngInfo.BitDepth).ToString() + " colors. ";
                                        bpp = 1;
                                    }
                                }

                                ImgData[5] = s;

                                ImgData[6] = (Math.Round(PngInfo.Width * PngInfo.Height * PngInfo.BitDepth * bpp / 8f / 1024f)).ToString() + " kb.";
                            }
                            else
                            {
                                ms.Position = 0;
                                try
                                {
                                    using (Image Img = Image.FromStream(ms))
                                    {
                                        Bitmap Bmp = Img as Bitmap;
                                        if (Bmp != null)
                                        {
                                            ImgData[5] = Bmp.PixelFormat.ToString() + "bpp";
                                        }
                                        ImgData[2] = Img.Width;
                                        ImgData[3] = Img.Height;
                                    }
                                }
                                catch (Exception)
                                {
                                    ImgData[2] = -1;
                                    ImgData[3] = -1;
                                    ImgData[5] = null;
                                    ImgData[8] = null;

                                }
                            }
                        }


                        ImageDataTable.Rows.Add(ImgData);
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
                dataGrid.CaptionText = "No file selected";
                CurrentFilename = null;
                return;
            }
            dataGrid.CaptionText = "Selected file: " + FileName;
            CurrentRowChanged(GetCurrencyManager, null);
        }

        private void FilesListBox_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            TImageInfo ImageInfo = (TImageInfo)FilesListBox.SelectedItem;
            if (ImageInfo == null) return;
            OpenFile(ImageInfo.File.FullName);
        }

        public void CurrentRowChanged(object sender, System.EventArgs e)
        {
            int Pos = ((BindingManagerBase)sender).Position;
            if (Pos < 0)
            {
                PreviewBox.Image = null;
                return;
            }

            DataRowCollection r = dataSet1.Tables["ImageDataTable"].Rows;
            if (Pos < 0 || Pos >= r.Count)
            {
                PreviewBox.Image = null;
            }
            else
            {
                byte[] ImgData = r[Pos].ItemArray[8] as byte[];
                if (ImgData == null)
                {
                    PreviewBox.Image = null;
                }
                else
                {
                    using (MemoryStream ms = new MemoryStream(ImgData))
                    {
                        PreviewBox.Image = Image.FromStream(ms);
                    }
                }
            }
        }

        private void btnOpen_Click(object sender, System.EventArgs e)
        {
            if (CurrentFilename == null)
            {
                MessageBox.Show("There is no open file");
                return;
            }
            System.Diagnostics.Process.Start(CurrentFilename);
        }

        private void btnConvert_Click(object sender, System.EventArgs e)
        {
            //This is not yet implemented...
            int Pos = GetImagePos;
            if (Pos < 0)
            {
                MessageBox.Show("There is no selected image", "Error");
                return;
            }
            if (CompressForm == null) CompressForm = new TCompressForm();
            CompressForm.ImageToUse = (byte[])dataSet1.Tables["ImageDataTable"].Rows[Pos].ItemArray[8];
            CompressForm.XlsFilename = CurrentFilename;
            CompressForm.ShowDialog();
        }

        private void FilesListBox_DrawItem(object sender, System.Windows.Forms.DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            e.DrawBackground();
            Brush myBrush = Brushes.Black;

            TImageInfo ImageInfo = (TImageInfo)((ListBox)sender).Items[e.Index];
            if (!ImageInfo.HasImages)
            {
                myBrush = Brushes.Silver;
            }
            if (ImageInfo.HasCrop)
            {
                myBrush = Brushes.Red;
            }

            FontStyle NewStyle;
            if (ImageInfo.HasARGB) NewStyle = FontStyle.Bold; else NewStyle = FontStyle.Regular;
            using (Font MyFont = new Font(e.Font, NewStyle))
            {
                e.Graphics.DrawString(ImageInfo.ToString(),
                    MyFont, myBrush, e.Bounds, StringFormat.GenericDefault);
            }
            e.DrawFocusRectangle();
        }

        private void btnSaveImage_Click(object sender, System.EventArgs e)
        {
            int Pos = GetImagePos;
            if (Pos < 0)
            {
                MessageBox.Show("There is no selected image to save", "Error");
                return;
            }

            string ext = dataSet1.Tables["ImageDataTable"].Rows[Pos].ItemArray[4].ToString().ToLower();
            saveImageDialog.DefaultExt = ext;
            saveImageDialog.Filter = ext + " Images|*." + ext;
            if (saveImageDialog.ShowDialog() != DialogResult.OK) return;
            byte[] ImgData = (byte[])dataSet1.Tables["ImageDataTable"].Rows[Pos].ItemArray[8];
            using (FileStream fs = new FileStream(saveImageDialog.FileName, FileMode.Create))
            {
                fs.Write(ImgData, 0, ImgData.Length);
            }

        }

        private void btnInfo_Click(object sender, System.EventArgs e)
        {
            MessageBox.Show("FlexCelImageExplorer is a small application targeted to reduce the size on images inside Excel files.\n" +
                "On the current version you can see the image properties and extract the images to disk.");

        }

        private void btnStretchPreview_Click(object sender, EventArgs e)
        {
            if (btnStretchPreview.Checked)
                PreviewBox.SizeMode = PictureBoxSizeMode.StretchImage;
            else
                PreviewBox.SizeMode = PictureBoxSizeMode.Normal;
        }

    }

    class TImageInfo
    {
        internal FileInfo File;
        internal bool IsValidFile;
        internal bool HasCrop;
        internal bool HasImages;
        internal bool HasARGB;

        public TImageInfo(FileInfo aFile, bool aIsValidFile, bool aHasCrop, bool aHasImages, bool aHasARGB)
        {
            File = aFile;
            HasCrop = aHasCrop;
            HasImages = aHasImages;
            IsValidFile = aIsValidFile;
            HasARGB = aHasARGB;
        }

        public override string ToString()
        {
            if (!IsValidFile)
                return " (*)" + File.ToString();
            return File.ToString();
        }

    }
}
