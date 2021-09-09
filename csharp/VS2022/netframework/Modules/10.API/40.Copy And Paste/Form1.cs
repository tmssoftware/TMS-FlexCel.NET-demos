using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Text;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System.Collections.Generic;

namespace CopyAndPaste
{
    /// <summary>
    /// Copy and Paste Example.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {

        public mainForm()
        {
            InitializeComponent();
        }

        private XlsFile Xls;

        private void btnNewFile_Click(object sender, System.EventArgs e)
        {
            try
            {
                Xls = new XlsFile();
                Xls.NewFile(1, TExcelFileFormat.v2019);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            try
            {
                if (openFileDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
                Xls = new XlsFile(openFileDialog.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }


        private void DoPaste(IDataObject iData)
        {
            if (Xls == null)
            {
                MessageBox.Show("Please push the New File button before pasting");
                return;
            }

            try
            {
                if (iData.GetDataPresent(FlexCelDataFormats.Excel97))
                {
                    //DO NOT CALL -> using (MemoryStream ms = (MemoryStream)iData.GetData(FlexCelDataFormats.Excel97))
                    //You shouldn't dispose the stream, as it belongs to the Clipboard.
                    object o = iData.GetData(FlexCelDataFormats.Excel97);
                    MemoryStream ms = (MemoryStream)o;
                    {
                        Xls.PasteFromXlsClipboardFormat(1, 1, TFlxInsertMode.NoneDown, ms);
                        MessageBox.Show("NATIVE Data has been pasted at cell A1");
                    }
                }
                else
                    if (iData.GetDataPresent(DataFormats.UnicodeText))
                {
                    Xls.PasteFromTextClipboardFormat(1, 1, TFlxInsertMode.NoneDown, (string)iData.GetData(DataFormats.UnicodeText));
                    MessageBox.Show("UNICODE TEXT Data has been pasted at cell A1");
                }
                else
                        if (iData.GetDataPresent(DataFormats.Text))
                {
                    Xls.PasteFromTextClipboardFormat(1, 1, TFlxInsertMode.NoneDown, (string)iData.GetData(DataFormats.Text));
                    MessageBox.Show("TEXT Data has been pasted at cell A1");

                }
                else
                {
                    MessageBox.Show("There is no Excel or Text data on the clipboard");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Xls = new XlsFile();
                Xls.NewFile(1, TExcelFileFormat.v2019);
            }
        }

        private void btnPaste_Click(object sender, System.EventArgs e)
        {
            DoPaste(Clipboard.GetDataObject());
        }

        private void DropHere_DragOver(object sender, System.Windows.Forms.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(FlexCelDataFormats.Excel97) ||
                e.Data.GetDataPresent(DataFormats.UnicodeText) ||
                e.Data.GetDataPresent(DataFormats.Text)
                )
                e.Effect = DragDropEffects.Copy;
        }


        private void DropHere_DragDrop(object sender, System.Windows.Forms.DragEventArgs e)
        {
            DoPaste(e.Data);
        }


        private void DoCopy(bool ToClipboard)
        {
            if (Xls == null)
            {
                MessageBox.Show("Please push the New File button before copying");
                return;
            }

            //VERY IMPORTANT!!!!!
            //****************************************************************************
            //The MemoryStreams CAN NOT BE DISPOSED UNTIL WE CALL Clipboard.SetObjectData.
            //Even when we assigned the Stream with the DataObject Data, it is still in use and can't be freed.
            //****************************************************************************

            try
            {
                DataObject data = new DataObject();
                List<MemoryStream> dataStreams = new List<MemoryStream>(); //we will use this list to dispose the memorystreams after they have been used.
                try
                {
                    foreach (FlexCelClipboardFormat cf in Enum.GetValues(typeof(FlexCelClipboardFormat)))
                    {
                        MemoryStream dataStream = new MemoryStream();
                        dataStreams.Add(dataStream);
                        Xls.CopyToClipboard(cf, dataStream);
                        dataStream.Position = 0;
                        data.SetData(FlexCelDataFormats.GetString(cf), dataStream);

                    }
                    if (ToClipboard)
                        Clipboard.SetDataObject(data, true);
                    else
                        DoDragDrop(data, DragDropEffects.Copy);
                }
                finally
                {
                    foreach (MemoryStream ms in dataStreams)
                    {
                        ms.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnCopy_Click(object sender, System.EventArgs e)
        {
            DoCopy(true);
        }

        private void btnDragMe_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            DoCopy(false);
        }

    }
}
