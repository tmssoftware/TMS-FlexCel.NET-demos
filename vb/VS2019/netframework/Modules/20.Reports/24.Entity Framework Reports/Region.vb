﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated from a template.
'
'     Manual changes to this file may cause unexpected behavior in your application.
'     Manual changes to this file will be overwritten if the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Namespace EntityFrameworkReports

	Partial Public Class Region
		<System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")> _
		Public Sub New()
			Me.Territories = New HashSet(Of Territory)()
		End Sub

		Public Property RegionID() As Integer
		Public Property RegionDescription() As String

		<System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")> _
		Public Overridable Property Territories() As ICollection(Of Territory)
	End Class
End Namespace
