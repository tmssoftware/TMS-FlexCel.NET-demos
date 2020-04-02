using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using FlexCel.Render;
using FlexCel.XlsAdapter;
using FlexCel.Pdf;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Pkcs;
using System.Drawing.Imaging;
using System.Reflection;
using System.Diagnostics;

namespace SigningPdfs
{
    public partial class mainForm: Form
    {
        public mainForm()
        {
            Application.EnableVisualStyles();
            InitializeComponent();
        }

        private void cbVisibleSignature_CheckedChanged(object sender, EventArgs e)
        {
            SignaturePicture.Visible = cbVisibleSignature.Checked;
            int delta = SignaturePicture.Height + 30;
            if (cbVisibleSignature.Checked) this.Height += delta; else this.Height -= delta;
        }

        private void SignaturePicture_Click(object sender, EventArgs e)
        {
            if (OpenImageDialog.ShowDialog() != DialogResult.OK) return;
            SignaturePicture.Load(OpenImageDialog.FileName);
        }

        private void btnCreateAndSign_Click(object sender, EventArgs e)
        {
            //Load the Excel file.
            if (OpenExcelDialog.ShowDialog() != DialogResult.OK) return;
            XlsFile xls = new XlsFile();
            xls.Open(OpenExcelDialog.FileName);

            string DataPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\..\..\";

            //Export it to pdf.
            using (FlexCelPdfExport pdf = new FlexCelPdfExport(xls, true))
            {
                pdf.FontEmbed = TFontEmbed.Embed;

                //Load the certificate and create a signer.
                //In this example we just have the password in clear. It should be kept in a SecureString.
                //Also make sure to set the flag X509KeyStorageFlags.EphemeralKeySet to avoid files created
                //on disk: https://snede.net/the-most-dangerous-constructor-in-net/
                //As X509KeyStorageFlags.EphemeralKeySet only exists in .NET 4.8 or newer, for older versions we will 
                //define it as (X509KeyStorageFlags)32. For .NET 4.8 or newer and  NET Core, you can use X509KeyStorageFlags.EphemeralKeySet
                  X509Certificate2 Cert = new X509Certificate2(DataPath + "flexcel.pfx", "password", X509KeyStorageFlags.EphemeralKeySet);  

                //Note that to use the CmsSigner class you need to add a reference to System.Security dll. 
                //It is *not* enough to add it to the using clauses, you need to add a reference to the dll.
                CmsSigner Signer = new CmsSigner(Cert);

                //By default CmsSigner uses SHA1, but SHA1 has known vulnerabilities and it is deprecated. 
                //So we will use SHA512 instead.
                //"2.16.840.1.101.3.4.2.3" is the Oid for SHA512.
                Signer.DigestAlgorithm = new System.Security.Cryptography.Oid("2.16.840.1.101.3.4.2.3");

                TPdfSignature sig;
                if (cbVisibleSignature.Checked)
                {
                    using (MemoryStream fs = new MemoryStream())
                    {
                        SignaturePicture.Image.Save(fs, ImageFormat.Png);
                        byte[] ImgData = fs.ToArray();

                        //The -1 as "page" parameter means the last page.
                        sig = new TPdfVisibleSignature(new TBuiltInSignerFactory(Signer),
                            "Signature", "I have read the document and certify it is valid.", "Springfield", "adrian@tmssoftware.com", -1, new RectangleF(50, 50, 140, 70), ImgData);
                    }
                }
                else
                {
                    sig = new TPdfSignature(new TBuiltInSignerFactory(Signer),
                                "Signature", "I have read the document and certify it is valid.", "Springfield", "adrian@tmssoftware.com");
                }

                //You must sign the document *BEFORE* starting to write it.
                pdf.Sign(sig);

                if (savePdfDialog.ShowDialog() != DialogResult.OK) return;
                using (FileStream PdfStream = new FileStream(savePdfDialog.FileName, FileMode.Create))
                {
                    pdf.BeginExport(PdfStream);
                    pdf.ExportAllVisibleSheets(false, "Signed Pdf");
                    pdf.EndExport();
                }

            }

            if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            Process.Start(savePdfDialog.FileName);

        }
    }
}
