Public Class MyDataGrid
    Inherits UserControl

    Public Sub SetColumns(columnNames As List(Of String))
        dgv.Columns.Clear()
        For Each name In columnNames
            dgv.Columns.Add(name, name)
        Next
    End Sub

    Public Sub LoadData(dt As DataTable)
        dgv.DataSource = dt
    End Sub

    Public Function GetData() As DataTable
        Return CType(dgv.DataSource, DataTable)
    End Function
End Class


Public Class CustomerGrid
    Inherits MyDataGrid

    Public Sub New()
        SetColumns(New List(Of String) From {"客户ID", "姓名", "电话"})
    End Sub
End Class
