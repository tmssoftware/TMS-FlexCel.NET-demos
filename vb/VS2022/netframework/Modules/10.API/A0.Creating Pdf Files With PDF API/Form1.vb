Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.Pdf
Imports System.IO
Imports System.Reflection
Imports System.Drawing.Drawing2D

Namespace CreatingPdfFilesWithPDFAPI
	''' <summary>
	''' Jow to create PDF files directly with FlexCel.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			If saveFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If

            Dim Underline As TUITextDecoration = New TUITextDecoration(TUIUnderline.Single)

            Dim pdf As New PdfWriter()

			Using file As New FileStream(saveFileDialog1.FileName, FileMode.Create)
				pdf.Compress = True
				pdf.BeginDoc(file)
				pdf.YAxisGrowsDown = True 'To keep it compatible with GDI+
                Using f As TUIFont = TUIFont.Create("times new roman", CSng(22.5), TUIFontStyle.Italic)
                    Using f2 As TUIFont = TUIFont.Create("Arial", CSng(12), TUIFontStyle.Italic)
                        pdf.DrawString("This is the first line on a test of many lines.", f, Underline, Brushes.Navy, 100, 100)
                        pdf.DrawString("Some unicode: " & ChrW(&HE2A).ToString() & ChrW(&HE27).ToString() & ChrW(&HE31).ToString() & ChrW(&HE2A).ToString() & ChrW(&HE14).ToString() & ChrW(&HE35).ToString(), f, Underline, Brushes.ForestGreen, 100, 200)
                        pdf.DrawString("More lines here!", f, Underline, Brushes.ForestGreen, 200, 300)
                        pdf.DrawString("And this is the last line.", f, Underline, Brushes.Black, 200, 400)
                        pdf.Properties.Author = "Adrian"
                        pdf.Properties.Title = "This is a test of FlexCel Api"
                        pdf.Properties.Keywords = "test" & vbLf & "flexcel" & vbLf & "api"
                        pdf.NewPage()
                        pdf.SaveState()
                        pdf.Rotate(200, 100, 45)
                        pdf.DrawString("Some rotated test", f, Underline, Brushes.Black, 200, 200)
                        pdf.RestoreState()
                        pdf.DrawString("Some NOT rotated text", f, Underline, Brushes.Black, 200, 200)
                        pdf.DrawString("Hello from FlexCel!", f2, Brushes.Black, 200, 50)

                        Dim points() As TPointF = { _
                            New TPointF(200, 100), _
                            New TPointF(200, 50), _
                            New TPointF(500, 50), _
                            New TPointF(700, 100) _
                        }
                        pdf.DrawLines(Pens.DarkOrchid, points)

                        Dim Coords As New RectangleF(100, 300, 100, 100)
                        Using Gradient As Brush = New LinearGradientBrush(Coords, Color.Red, Color.Blue, 0F)
                            pdf.DrawAndFillRectangle(Pens.Red, Gradient, 100, 300, 100, 100)
                        End Using
                        pdf.DrawRectangle(Pens.DarkSlateBlue, 100, 300, 50, 50)
                        pdf.DrawLine(Pens.Black, 100, 300, 200, 400)

                        Dim AssemblyPath As String = Path.GetDirectoryName(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location))
                        Using Img As Image = Image.FromFile(AssemblyPath & Path.DirectorySeparatorChar & ".." & Path.DirectorySeparatorChar & ".." & Path.DirectorySeparatorChar & "test.jpg")
                            pdf.DrawImage(Img, New RectangleF(200, 300, 200, 150), Nothing)
                        End Using
                        pdf.IntersectClipRegion(New RectangleF(100, 100, 50, 50))
                        pdf.FillRectangle(Brushes.DarkTurquoise, 100, 100, 100, 100)

                        pdf.EndDoc()
                    End Using
                End Using
            End Using
			If MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
				Process.Start(saveFileDialog1.FileName)
			End If

		End Sub
	End Class
End Namespace
