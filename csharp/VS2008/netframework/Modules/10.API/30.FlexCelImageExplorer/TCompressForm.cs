using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;

namespace FlexCelImageExplorer
{
    /// <summary>
    /// Summary description for TCompressForm.
    /// </summary>
    public partial class TCompressForm: System.Windows.Forms.Form
    {

        public TCompressForm()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();
            cbPixelFormat.SelectedIndex = 2;

            //
            // TODO: Add any constructor code after InitializeComponent call
            //
        }

        private byte[] FImageToUse;
        private string FXlsFilename;

        internal byte[] ImageToUse
        {
            get
            {
                return FImageToUse;
            }
            set
            {
                FImageToUse = value;
                using (MemoryStream ms = new MemoryStream(value))
                {
                    pictureBox1.Image = Image.FromStream(ms);
                }
            }
        }

        internal string XlsFilename
        {
            get
            {
                return FXlsFilename;
            }
            set
            {
                FXlsFilename = value;
            }
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void TCompressForm_Load(object sender, System.EventArgs e)
        {
        }

        private void btnOk_Click(object sender, System.EventArgs e)
        {
            Close();
        }
    }
}
