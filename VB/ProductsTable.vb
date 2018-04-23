Imports Microsoft.VisualBasic
Imports System
Imports System.Data

Namespace RichEditCalculatedField
	Public Class ProductsTable
		Inherits DataTable
		Public Sub New()
			MyBase.New()

			Me.TableName = "ProductsTable"

			Me.Columns.Add("ProductID", GetType(Int32))
			Me.Columns.Add("ProductName", GetType(String))
			Me.Columns.Add("UnitsInStock", GetType(Integer))
			Me.Columns.Add("UnitPrice", GetType(Decimal))
			Me.Constraints.Add("IDPK", Me.Columns("ProductID"), True)
		End Sub

		Public Shared Function CreateData() As ProductsTable
			Dim table As New ProductsTable()

			table.Rows.Add(New Object() { 1, "Chai", 39, 18 })
			table.Rows.Add(New Object() { 2, "Chang", 17, 19 })
			table.Rows.Add(New Object() { 3, "Guarana Fantastica", 20, 4.5 })
			table.Rows.Add(New Object() { 4, "Sasquatch Ale", 111, 14 })
			table.Rows.Add(New Object() { 5, "Steeleye Stout", 20, 18 })

			Return table
		End Function
	End Class
End Namespace
