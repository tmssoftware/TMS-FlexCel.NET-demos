Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Data.OleDb
Imports FlexCel.Core
Imports FlexCel.XlsAdapter
Imports FlexCel.Report
Imports FlexCel.Render
Imports FlexCel.Pdf
Imports System.Globalization
Imports System.Xml
Namespace FeaturesPage
	Partial Public Class mainForm
		Inherits System.Windows.Forms.Form

		Private label1 As System.Windows.Forms.Label
		Private oleDbSelectCommand1 As System.Data.OleDb.OleDbCommand
		Private oleDbInsertCommand1 As System.Data.OleDb.OleDbCommand
		Private oleDbUpdateCommand1 As System.Data.OleDb.OleDbCommand
		Private oleDbDeleteCommand1 As System.Data.OleDb.OleDbCommand
		Private dbconnection As System.Data.OleDb.OleDbConnection
		Private categoriesAdapter As System.Data.OleDb.OleDbDataAdapter
		Private featuresAdapter As System.Data.OleDb.OleDbDataAdapter
		Private oleDbSelectCommand2 As System.Data.OleDb.OleDbCommand
		Private oleDbInsertCommand2 As System.Data.OleDb.OleDbCommand
		Private oleDbUpdateCommand2 As System.Data.OleDb.OleDbCommand
		Private oleDbDeleteCommand2 As System.Data.OleDb.OleDbCommand
		Private oleDbSelectCommand3 As System.Data.OleDb.OleDbCommand
		Private hyperlinksAdapter As System.Data.OleDb.OleDbDataAdapter
		Private components As System.ComponentModel.IContainer = Nothing

		''' <summary>
		''' Clean up any resources being used.
		''' </summary>
		Protected Overrides Sub Dispose(ByVal disposing As Boolean)
			If disposing Then
				If components IsNot Nothing Then
					components.Dispose()
				End If
			End If
			MyBase.Dispose(disposing)
		End Sub

		#Region "Windows Form Designer generated code"
		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
			Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(mainForm))
			Me.label1 = New System.Windows.Forms.Label()
			Me.oleDbSelectCommand1 = New System.Data.OleDb.OleDbCommand()
			Me.dbconnection = New System.Data.OleDb.OleDbConnection()
			Me.oleDbInsertCommand1 = New System.Data.OleDb.OleDbCommand()
			Me.oleDbUpdateCommand1 = New System.Data.OleDb.OleDbCommand()
			Me.oleDbDeleteCommand1 = New System.Data.OleDb.OleDbCommand()
			Me.categoriesAdapter = New System.Data.OleDb.OleDbDataAdapter()
			Me.featuresAdapter = New System.Data.OleDb.OleDbDataAdapter()
			Me.oleDbDeleteCommand2 = New System.Data.OleDb.OleDbCommand()
			Me.oleDbInsertCommand2 = New System.Data.OleDb.OleDbCommand()
			Me.oleDbSelectCommand2 = New System.Data.OleDb.OleDbCommand()
			Me.oleDbUpdateCommand2 = New System.Data.OleDb.OleDbCommand()
			Me.oleDbSelectCommand3 = New System.Data.OleDb.OleDbCommand()
			Me.hyperlinksAdapter = New System.Data.OleDb.OleDbDataAdapter()
			Me.mainToolbar = New System.Windows.Forms.ToolStrip()
			Me.toolStripButton4 = New System.Windows.Forms.ToolStripButton()
			Me.toolStripButton3 = New System.Windows.Forms.ToolStripButton()
			Me.toolStripButton2 = New System.Windows.Forms.ToolStripButton()
			Me.toolStripButton1 = New System.Windows.Forms.ToolStripButton()
			Me.mainToolbar.SuspendLayout()
			Me.SuspendLayout()
			' 
			' label1
			' 
			Me.label1.Location = New System.Drawing.Point(179, 69)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(200, 48)
			Me.label1.TabIndex = 4
			Me.label1.Text = "Files will be saved under a ""Features"" Folder in under where the application is r" & "unning."
			' 
			' oleDbSelectCommand1
			' 
			Me.oleDbSelectCommand1.CommandText = "SELECT Caption, CategoryId, CategoryName, Description FROM Categories"
			Me.oleDbSelectCommand1.Connection = Me.dbconnection
			' 
			' dbconnection
			' 
			Me.dbconnection.ConnectionString = resources.GetString("dbconnection.ConnectionString")
			' 
			' oleDbInsertCommand1
			' 
			Me.oleDbInsertCommand1.CommandText = "INSERT INTO Categories(Caption, CategoryName, Description) VALUES (?, ?, ?)"
			Me.oleDbInsertCommand1.Connection = Me.dbconnection
			Me.oleDbInsertCommand1.Parameters.AddRange(New System.Data.OleDb.OleDbParameter() { _
				New System.Data.OleDb.OleDbParameter("Caption", System.Data.OleDb.OleDbType.VarWChar, 255, "Caption"), _
				New System.Data.OleDb.OleDbParameter("CategoryName", System.Data.OleDb.OleDbType.VarWChar, 255, "CategoryName"), _
				New System.Data.OleDb.OleDbParameter("Description", System.Data.OleDb.OleDbType.VarWChar, 0, "Description") _
			})
			' 
			' oleDbUpdateCommand1
			' 
			Me.oleDbUpdateCommand1.CommandText = resources.GetString("oleDbUpdateCommand1.CommandText")
			Me.oleDbUpdateCommand1.Connection = Me.dbconnection
			Me.oleDbUpdateCommand1.Parameters.AddRange(New System.Data.OleDb.OleDbParameter() { _
				New System.Data.OleDb.OleDbParameter("Caption", System.Data.OleDb.OleDbType.VarWChar, 255, "Caption"), _
				New System.Data.OleDb.OleDbParameter("CategoryName", System.Data.OleDb.OleDbType.VarWChar, 255, "CategoryName"), _
				New System.Data.OleDb.OleDbParameter("Description", System.Data.OleDb.OleDbType.VarWChar, 0, "Description"), _
				New System.Data.OleDb.OleDbParameter("Original_CategoryId", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, (CByte(0)), (CByte(0)), "CategoryId", System.Data.DataRowVersion.Original, Nothing), _
				New System.Data.OleDb.OleDbParameter("Original_Caption", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, (CByte(0)), (CByte(0)), "Caption", System.Data.DataRowVersion.Original, Nothing), _
				New System.Data.OleDb.OleDbParameter("Original_Caption1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, (CByte(0)), (CByte(0)), "Caption", System.Data.DataRowVersion.Original, Nothing), _
				New System.Data.OleDb.OleDbParameter("Original_CategoryName", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, (CByte(0)), (CByte(0)), "CategoryName", System.Data.DataRowVersion.Original, Nothing), _
				New System.Data.OleDb.OleDbParameter("Original_CategoryName1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, (CByte(0)), (CByte(0)), "CategoryName", System.Data.DataRowVersion.Original, Nothing) _
			})
			' 
			' oleDbDeleteCommand1
			' 
			Me.oleDbDeleteCommand1.CommandText = "DELETE FROM Categories WHERE (CategoryId = ?) AND (Caption = ? OR ? IS NULL AND C" & "aption IS NULL) AND (CategoryName = ? OR ? IS NULL AND CategoryName IS NULL)"
			Me.oleDbDeleteCommand1.Connection = Me.dbconnection
			Me.oleDbDeleteCommand1.Parameters.AddRange(New System.Data.OleDb.OleDbParameter() { _
				New System.Data.OleDb.OleDbParameter("Original_CategoryId", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, (CByte(0)), (CByte(0)), "CategoryId", System.Data.DataRowVersion.Original, Nothing), _
				New System.Data.OleDb.OleDbParameter("Original_Caption", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, (CByte(0)), (CByte(0)), "Caption", System.Data.DataRowVersion.Original, Nothing), _
				New System.Data.OleDb.OleDbParameter("Original_Caption1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, (CByte(0)), (CByte(0)), "Caption", System.Data.DataRowVersion.Original, Nothing), _
				New System.Data.OleDb.OleDbParameter("Original_CategoryName", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, (CByte(0)), (CByte(0)), "CategoryName", System.Data.DataRowVersion.Original, Nothing), _
				New System.Data.OleDb.OleDbParameter("Original_CategoryName1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, (CByte(0)), (CByte(0)), "CategoryName", System.Data.DataRowVersion.Original, Nothing) _
			})
			' 
			' categoriesAdapter
			' 
			Me.categoriesAdapter.DeleteCommand = Me.oleDbDeleteCommand1
			Me.categoriesAdapter.InsertCommand = Me.oleDbInsertCommand1
			Me.categoriesAdapter.SelectCommand = Me.oleDbSelectCommand1
			Me.categoriesAdapter.TableMappings.AddRange(New System.Data.Common.DataTableMapping() { New System.Data.Common.DataTableMapping("Table", "Categories", New System.Data.Common.DataColumnMapping() { New System.Data.Common.DataColumnMapping("Caption", "Caption"), New System.Data.Common.DataColumnMapping("CategoryId", "CategoryId"), New System.Data.Common.DataColumnMapping("CategoryName", "CategoryName"), New System.Data.Common.DataColumnMapping("Description", "Description")})})
			Me.categoriesAdapter.UpdateCommand = Me.oleDbUpdateCommand1
			' 
			' featuresAdapter
			' 
			Me.featuresAdapter.DeleteCommand = Me.oleDbDeleteCommand2
			Me.featuresAdapter.InsertCommand = Me.oleDbInsertCommand2
			Me.featuresAdapter.SelectCommand = Me.oleDbSelectCommand2
			Me.featuresAdapter.TableMappings.AddRange(New System.Data.Common.DataTableMapping() { New System.Data.Common.DataTableMapping("Table", "Features", New System.Data.Common.DataColumnMapping() { New System.Data.Common.DataColumnMapping("Caption", "Caption"), New System.Data.Common.DataColumnMapping("CategoryId", "CategoryId"), New System.Data.Common.DataColumnMapping("Description", "Description"), New System.Data.Common.DataColumnMapping("FeaturesId", "FeaturesId")})})
			Me.featuresAdapter.UpdateCommand = Me.oleDbUpdateCommand2
			' 
			' oleDbDeleteCommand2
			' 
			Me.oleDbDeleteCommand2.CommandText = "DELETE FROM Features WHERE (FeaturesId = ?) AND (Caption = ?) AND (CategoryId = ?" & ")"
			Me.oleDbDeleteCommand2.Parameters.AddRange(New System.Data.OleDb.OleDbParameter() { _
				New System.Data.OleDb.OleDbParameter("Original_FeaturesId", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, (CByte(0)), (CByte(0)), "FeaturesId", System.Data.DataRowVersion.Original, Nothing), _
				New System.Data.OleDb.OleDbParameter("Original_Caption", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, (CByte(0)), (CByte(0)), "Caption", System.Data.DataRowVersion.Original, Nothing), _
				New System.Data.OleDb.OleDbParameter("Original_CategoryId", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, (CByte(0)), (CByte(0)), "CategoryId", System.Data.DataRowVersion.Original, Nothing) _
			})
			' 
			' oleDbInsertCommand2
			' 
			Me.oleDbInsertCommand2.CommandText = "INSERT INTO Features(Caption, CategoryId, Description) VALUES (?, ?, ?)"
			Me.oleDbInsertCommand2.Parameters.AddRange(New System.Data.OleDb.OleDbParameter() { _
				New System.Data.OleDb.OleDbParameter("Caption", System.Data.OleDb.OleDbType.VarWChar, 255, "Caption"), _
				New System.Data.OleDb.OleDbParameter("CategoryId", System.Data.OleDb.OleDbType.Integer, 0, "CategoryId"), _
				New System.Data.OleDb.OleDbParameter("Description", System.Data.OleDb.OleDbType.VarWChar, 0, "Description") _
			})
			' 
			' oleDbSelectCommand2
			' 
			Me.oleDbSelectCommand2.CommandText = "SELECT Caption, CategoryId, Description, FeaturesId FROM Features order by Positi" & "onInSheet"
			Me.oleDbSelectCommand2.Connection = Me.dbconnection
			' 
			' oleDbUpdateCommand2
			' 
			Me.oleDbUpdateCommand2.CommandText = "UPDATE Features SET Caption = ?, CategoryId = ?, Description = ? WHERE (FeaturesI" & "d = ?) AND (Caption = ?) AND (CategoryId = ?)"
			Me.oleDbUpdateCommand2.Parameters.AddRange(New System.Data.OleDb.OleDbParameter() { _
				New System.Data.OleDb.OleDbParameter("Caption", System.Data.OleDb.OleDbType.VarWChar, 255, "Caption"), _
				New System.Data.OleDb.OleDbParameter("CategoryId", System.Data.OleDb.OleDbType.Integer, 0, "CategoryId"), _
				New System.Data.OleDb.OleDbParameter("Description", System.Data.OleDb.OleDbType.VarWChar, 0, "Description"), _
				New System.Data.OleDb.OleDbParameter("Original_FeaturesId", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, (CByte(0)), (CByte(0)), "FeaturesId", System.Data.DataRowVersion.Original, Nothing), _
				New System.Data.OleDb.OleDbParameter("Original_Caption", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, (CByte(0)), (CByte(0)), "Caption", System.Data.DataRowVersion.Original, Nothing), _
				New System.Data.OleDb.OleDbParameter("Original_CategoryId", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, (CByte(0)), (CByte(0)), "CategoryId", System.Data.DataRowVersion.Original, Nothing) _
			})
			' 
			' oleDbSelectCommand3
			' 
			Me.oleDbSelectCommand3.CommandText = "SELECT FeaturesId, HiperlinksId, LinkName, Url FROM Hyperlinks"
			Me.oleDbSelectCommand3.Connection = Me.dbconnection
			' 
			' hyperlinksAdapter
			' 
			Me.hyperlinksAdapter.SelectCommand = Me.oleDbSelectCommand3
			Me.hyperlinksAdapter.TableMappings.AddRange(New System.Data.Common.DataTableMapping() { New System.Data.Common.DataTableMapping("Table", "Hyperlinks", New System.Data.Common.DataColumnMapping() { New System.Data.Common.DataColumnMapping("FeaturesId", "FeaturesId"), New System.Data.Common.DataColumnMapping("HiperlinksId", "HiperlinksId"), New System.Data.Common.DataColumnMapping("LinkName", "LinkName"), New System.Data.Common.DataColumnMapping("Url", "Url")})})
			' 
			' mainToolbar
			' 
			Me.mainToolbar.Items.AddRange(New System.Windows.Forms.ToolStripItem() { Me.toolStripButton4, Me.toolStripButton3, Me.toolStripButton2, Me.toolStripButton1})
			Me.mainToolbar.Location = New System.Drawing.Point(0, 0)
			Me.mainToolbar.Name = "mainToolbar"
			Me.mainToolbar.Size = New System.Drawing.Size(528, 38)
			Me.mainToolbar.TabIndex = 5
			Me.mainToolbar.Text = "mainToolbar"
			' 
			' toolStripButton4
			' 
			Me.toolStripButton4.Image = (CType(resources.GetObject("toolStripButton4.Image"), System.Drawing.Image))
			Me.toolStripButton4.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.toolStripButton4.Name = "toolStripButton4"
			Me.toolStripButton4.Size = New System.Drawing.Size(78, 35)
			Me.toolStripButton4.Text = "Save to Excel"
			Me.toolStripButton4.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.toolStripButton4.Click += New System.EventHandler(Me.btnExportExcel_Click)
			' 
			' toolStripButton3
			' 
			Me.toolStripButton3.Image = (CType(resources.GetObject("toolStripButton3.Image"), System.Drawing.Image))
			Me.toolStripButton3.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.toolStripButton3.Name = "toolStripButton3"
			Me.toolStripButton3.Size = New System.Drawing.Size(94, 35)
			Me.toolStripButton3.Text = "Export to HTML"
			Me.toolStripButton3.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.toolStripButton3.Click += New System.EventHandler(Me.btnExportHtml_Click)
			' 
			' toolStripButton2
			' 
			Me.toolStripButton2.Image = (CType(resources.GetObject("toolStripButton2.Image"), System.Drawing.Image))
			Me.toolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.toolStripButton2.Name = "toolStripButton2"
			Me.toolStripButton2.Size = New System.Drawing.Size(82, 35)
			Me.toolStripButton2.Text = "Export to PDF"
			Me.toolStripButton2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.toolStripButton2.Click += New System.EventHandler(Me.btnExportPDF_Click)
			' 
			' toolStripButton1
			' 
			Me.toolStripButton1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
			Me.toolStripButton1.Image = (CType(resources.GetObject("toolStripButton1.Image"), System.Drawing.Image))
			Me.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
			Me.toolStripButton1.Name = "toolStripButton1"
			Me.toolStripButton1.Size = New System.Drawing.Size(59, 35)
			Me.toolStripButton1.Text = "     E&xit     "
			Me.toolStripButton1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
'			Me.toolStripButton1.Click += New System.EventHandler(Me.button2_Click)
			' 
			' mainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(528, 126)
			Me.Controls.Add(Me.label1)
			Me.Controls.Add(Me.mainToolbar)
			Me.Name = "mainForm"
			Me.Text = "Features FlexCel"
			Me.mainToolbar.ResumeLayout(False)
			Me.mainToolbar.PerformLayout()
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub
		#End Region

		Private mainToolbar As ToolStrip
		Private WithEvents toolStripButton4 As ToolStripButton
		Private WithEvents toolStripButton3 As ToolStripButton
		Private WithEvents toolStripButton2 As ToolStripButton
		Private WithEvents toolStripButton1 As ToolStripButton
	End Class
End Namespace

