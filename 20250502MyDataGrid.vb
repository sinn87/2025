Imports System.ComponentModel
Imports System.Windows.Forms

Public Class MyDataGrid
    Inherits UserControl

    Private WithEvents dgv As New DataGridView()
    Private originalValue As Object = Nothing

    Public Sub New()
        ' DataGridView を初期化してコントロールに追加
        Me.Controls.Add(dgv)
        dgv.Dock = DockStyle.Fill
        dgv.AllowUserToAddRows = False
        dgv.RowHeadersVisible = False
        dgv.SelectionMode = DataGridViewSelectionMode.CellSelect
        dgv.EditMode = DataGridViewEditMode.EditOnEnter
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    End Sub

    ''' <summary>
    ''' 列名を設定（既存の列はクリアされます）
    ''' </summary>
    Public Sub SetColumns(columnNames As List(Of String))
        dgv.Columns.Clear()
        For Each name In columnNames
            dgv.Columns.Add(name, name)
        Next
    End Sub

    ''' <summary>
    ''' DataTable を読み込んで表示
    ''' </summary>
    Public Sub LoadData(dt As DataTable)
        dgv.DataSource = Nothing
        dgv.Rows.Clear()
        dgv.Columns.Clear()

        ' 列を追加
        For Each col As DataColumn In dt.Columns
            dgv.Columns.Add(col.ColumnName, col.ColumnName)
        Next

        ' 行を追加
        For Each row As DataRow In dt.Rows
            dgv.Rows.Add(row.ItemArray)
        Next
    End Sub

    ''' <summary>
    ''' 現在の表のデータを DataTable 形式で取得
    ''' </summary>
    Public Function GetData() As DataTable
        Dim dt As New DataTable()
        For Each col As DataGridViewColumn In dgv.Columns
            dt.Columns.Add(col.HeaderText)
        Next

        For Each row As DataGridViewRow In dgv.Rows
            If Not row.IsNewRow Then
                Dim dr As DataRow = dt.NewRow()
                For i As Integer = 0 To dgv.Columns.Count - 1
                    dr(i) = row.Cells(i).Value
                Next
                dt.Rows.Add(dr)
            End If
        Next
        Return dt
    End Function

    ''' <summary>
    ''' セルの背景色（強調表示）をリセット
    ''' </summary>
    Public Sub ClearHighlights()
        For Each row As DataGridViewRow In dgv.Rows
            For Each cell As DataGridViewCell In row.Cells
                cell.Style.BackColor = Color.White
            Next
        Next
    End Sub

    ' セルの編集を開始するとき、元の値を記録
    Private Sub dgv_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgv.CellBeginEdit
        originalValue = dgv.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
    End Sub

    ' セルの値が変更された時、前の値と比較し、変更があれば背景色を黄色に
    Private Sub dgv_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgv.CellValueChanged
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim cell = dgv.Rows(e.RowIndex).Cells(e.ColumnIndex)
            Dim newValue = cell.Value

            If Not Object.Equals(originalValue, newValue) Then
                cell.Style.BackColor = Color.Yellow
            Else
                ' 値が変わっていない場合は白に戻す
                cell.Style.BackColor = Color.White
            End If
        End If
    End Sub
End Class





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
