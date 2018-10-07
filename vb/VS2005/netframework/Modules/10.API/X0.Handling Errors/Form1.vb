Imports System.Collections
Imports System.ComponentModel
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports System.IO
Imports System.Reflection
Imports System.Text

Imports FlexCel.Render

Namespace HandlingErrors
	''' <summary>
	''' How to handle non fatal errors with FlexCel.
	''' </summary>
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private FlexCelTrace_OnErrorHandler As FlexCelErrorEventHandler
		Public Sub New()
			InitializeComponent()

			'Create a list to hold error messages. Keeping all error messages in memory is normally not a good thing to do, 
			'but for this demo it is ok.
			ErrorList = New ArrayList()

			'Hook our error handler to FlexCel error handler.	
			FlexCelTrace_OnErrorHandler = New FlexCelErrorEventHandler(AddressOf FlexCelTrace_OnError) 'We will save the value of the delegate here so we can unhook the event on dispose.
			AddHandler FlexCelTrace.OnError, FlexCelTrace_OnErrorHandler
		End Sub

		Private ErrorList As ArrayList
		Private Shared ErrorListLock As New Object() 'Used to lock ErrorList and ensure no more than one thread writes to it.

		Private ReadOnly Property PathToExe() As String
			Get
				Return Path.Combine(Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), ".."), "..") & Path.DirectorySeparatorChar
			End Get
		End Property


		Private Sub button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button1.Click
			ErrorList.Clear()
			errorBox.Text = ""

			Try
				DoThings()
			Catch ex As MyAbortException
				MessageBox.Show(ex.Message)
			End Try

			If ErrorList.Count = 0 Then
				errorBox.Text = "No errors!"
			Else
				errorBox.Text = String.Format("There were {0} error messages" & Environment.NewLine, ErrorList.Count)
				For Each s As String In ErrorList
					errorBox.AppendText(s & Environment.NewLine)
				Next s
			End If
		End Sub


		Private Sub DoThings()
			Dim xls As ExcelFile = New XlsFile(True)
			xls.NewFile(1, TExcelFileFormat.v2019)

			For r As Integer = 1 To 1999
				xls.InsertHPageBreak(r) 'This won't throw an exception here, since FlexCel allows to have more than 1025 page breaks, but at the moment of saving. (since an xls file can't have more than that)
			Next r

			xls.SetCellValue(1, 1, "We have a page break on each row, so this will print/export as one row per page")
			xls.SetCellValue(2, 1, "??? ? ? ? ???? ????") 'Since we leave the font at arial, this won't show when exporting to pdf.

			Dim fmt As TFlxFormat = xls.GetDefaultFormat
			fmt.Font.Name = "Arial Unicode MS"
			xls.SetCellValue(3, 1, "??? ? ? ? ???? ????", xls.AddFormat(fmt)) 'this will display fine in the pdf.

			fmt.Font.Name = "ThisFontDoesntExists"
			xls.SetCellValue(4, 1, "This font doesn't exists", xls.AddFormat(fmt))

			'Tahoma doesn't have italic variant. See http://help.lockergnome.com/office/Tahoma-italic-ftopict705661.html
			'You shouldn't normally use Tahoma italics in a document. If we embedded the fonts in this pdf, the fake italics wouldn't work.
			fmt.Font.Name = "Tahoma"
			fmt.Font.Style = TFlxFontStyles.Italic
			xls.SetCellValue(5, 1, "This is fake italics", xls.AddFormat(fmt))

			If saveFileDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
				Return
			End If

			Using pdf As New FlexCelPdfExport(xls, True)
				pdf.Export(Path.ChangeExtension(saveFileDialog1.FileName, ".pdf"))
			End Using

			xls.Save(saveFileDialog1.FileName & ".xls")
		End Sub

		''' <summary>
		''' This is the generic event handler for non fatal errors. We hooked it in the mainForm constructor.
		''' </summary>
		''' <param name="e"></param>
		Private Sub FlexCelTrace_OnError(ByVal e As TFlexCelErrorInfo)

			If cbIgnoreFontErrors.Checked Then
				Select Case e.Error
					'Ignore this errors:
					Case FlexCelError.PdfFontNotFound, FlexCelError.PdfGlyphNotInFont, FlexCelError.PdfFauxBoldOrItalics
						Return
				End Select
			End If


			'Normally tracing non fatal errors is a good idea. 
			'Depending on the listener of your trace object, you can redirect this to a log, the event viewer or wherever else.
			Trace.WriteLine(e.Message)

			'If we selected "Stop On Errors" we will abort file generation by throwing an exception that will be
			'catched in the main block.
			If cbStopOnErrors.Checked Then
				Throw New MyAbortException(e.Message)
			End If

			'In this case this is a single thread app so locking is not really necessary,
			'but it is a good practice to always lock access to global objects in this error handler.
			'This event handler might me called from more than one thread, and you don't want to mess
			'the object collecting the messages (in this case ErrorList).
			SyncLock ErrorListLock
				ErrorList.Add(System.Threading.Thread.CurrentThread.Name & ": - " & e.Message)
			End SyncLock
		End Sub
	End Class

	''' <summary>
	''' A custom exception designed to notify us when a non fatal error must be aborted.
	''' </summary>
	Public Class MyAbortException
		Inherits Exception

		Public Sub New(ByVal aMessage As String)
			MyBase.New(aMessage)
		End Sub
	End Class
End Namespace
