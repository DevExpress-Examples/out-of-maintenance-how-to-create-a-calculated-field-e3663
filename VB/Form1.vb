Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Windows.Forms
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Namespace RichEditCalculatedField
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

			richEditControl1.Options.MailMerge.DataSource = ProductsTable.CreateData()
			RestoreTemplate()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			Dim mailMergeOptions As MailMergeOptions = richEditControl1.Document.CreateMailMergeOptions()
			mailMergeOptions.MergeMode = MergeMode.JoinTables

			Dim server As New RichEditDocumentServer()

			AddHandler server.CalculateDocumentVariable, AddressOf server_CalculateDocumentVariable

			richEditControl1.Document.MailMerge(mailMergeOptions, server.Document)

			richEditControl1.LoadDocument("HeaderTemplate.rtf")
			richEditControl1.Document.AppendDocumentContent(server.Document.Range)
		End Sub

		Private Sub server_CalculateDocumentVariable(ByVal sender As Object, ByVal e As CalculateDocumentVariableEventArgs)
			If e.VariableName = "Prod" Then
				Dim productId As Integer = -1

				If Int32.TryParse(e.Arguments(0).Value, productId) Then
					Dim row As DataRow = (CType(richEditControl1.Options.MailMerge.DataSource, DataTable)).Rows.Find(productId)
					Dim unitsInStock As Integer = Convert.ToInt32(row(e.Arguments(1).Value))
					Dim unitPrice As Decimal = Convert.ToDecimal(row(e.Arguments(2).Value))

					e.Value = unitsInStock * unitPrice
					e.Handled = True
				End If
			End If
		End Sub

		#Region "Helper Methods"
		Private Sub button2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button2.Click
			RestoreTemplate()
		End Sub

		Private Sub RestoreTemplate()
			richEditControl1.LoadDocument("DetailTemplate.rtf")
			ShowFieldCodes()
		End Sub

		Private Sub ShowFieldCodes()
			Dim doc As Document = richEditControl1.Document
			doc.BeginUpdate()
			For Each f As Field In doc.Fields
				f.ShowCodes = True
			Next f
			doc.EndUpdate()
		End Sub
		#End Region ' Helper Methods
	End Class
End Namespace