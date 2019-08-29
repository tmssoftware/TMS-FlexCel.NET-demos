Imports System.ComponentModel
Imports System.Text
Imports FlexCel.Render
Imports FlexCel.XlsAdapter
Imports FlexCel.Pdf
Imports System.IO
Imports System.Security.Cryptography.X509Certificates
Imports System.Security.Cryptography.Pkcs
Imports System.Drawing.Imaging
Imports System.Reflection

Namespace SigningPdfs
	Partial Public Class mainForm
		Inherits Form

		Public Sub New()
			Application.EnableVisualStyles()
			InitializeComponent()
		End Sub

		Private Sub cbVisibleSignature_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles cbVisibleSignature.CheckedChanged
			SignaturePicture.Visible = cbVisibleSignature.Checked
			Dim delta As Integer = SignaturePicture.Height + 30
			If cbVisibleSignature.Checked Then
				Me.Height += delta
			Else
				Me.Height -= delta
			End If
		End Sub

		Private Sub SignaturePicture_Click(ByVal sender As Object, ByVal e As EventArgs) Handles SignaturePicture.Click
			If OpenImageDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			SignaturePicture.Load(OpenImageDialog.FileName)
		End Sub

		Private Sub btnCreateAndSign_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCreateAndSign.Click
			'Load the Excel file.
			If OpenExcelDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If
			Dim xls As New XlsFile()
			xls.Open(OpenExcelDialog.FileName)

			Dim DataPath As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & "\..\..\"

			'Export it to pdf.
			Using pdf As New FlexCelPdfExport(xls, True)
				pdf.FontEmbed = TFontEmbed.Embed

				'Load the certificate and create a signer.
				Dim Cert As New X509Certificate2(DataPath & "flexcel.pfx", "password") 'In this example we just have the password in clear. It should be kept in a SecureString.

				'Note that to use the CmsSigner class you need to add a refrence to System.Security dll. 
				'It is *not* enough to add it to the using clauses, you need to add a reference to the dll.
				Dim Signer As New CmsSigner(Cert)

				'By default CmsSigner uses SHA1, but SHA1 has known vulnerabilities and it is deprecated. 
				'So we will use SHA512 instead.
				'"2.16.840.1.101.3.4.2.3" is the Oid for SHA512.
				Signer.DigestAlgorithm = New System.Security.Cryptography.Oid("2.16.840.1.101.3.4.2.3")

				Dim sig As TPdfSignature
				If cbVisibleSignature.Checked Then
					Using fs As New MemoryStream()
						SignaturePicture.Image.Save(fs, ImageFormat.Png)
						Dim ImgData() As Byte = fs.ToArray()

						'The -1 as "page" parameter means the last page.
						sig = New TPdfVisibleSignature(New TBuiltInSignerFactory(Signer), "Signature", "I have read the document and certify it is valid.", "Springfield", "adrian@tmssoftware.com", -1, New RectangleF(50, 50, 140, 70), ImgData)
					End Using
				Else
					sig = New TPdfSignature(New TBuiltInSignerFactory(Signer), "Signature", "I have read the document and certify it is valid.", "Springfield", "adrian@tmssoftware.com")
				End If

				'You must sign the document *BEFORE* starting to write it.
				pdf.Sign(sig)

				If savePdfDialog.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
					Return
				End If
				Using PdfStream As New FileStream(savePdfDialog.FileName, FileMode.Create)
					pdf.BeginExport(PdfStream)
					pdf.ExportAllVisibleSheets(False, "Signed Pdf")
					pdf.EndExport()
				End Using

			End Using

			If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> System.Windows.Forms.DialogResult.Yes Then
				Return
			End If
			Process.Start(savePdfDialog.FileName)

		End Sub
	End Class
End Namespace
