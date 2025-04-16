获取全部控件数据（用于新登录）  

Public Function ExtractAllItems(parent As Control) As List(Of DetailItem)
    Dim list As New List(Of DetailItem)

    For Each ctrl In parent.Controls
        Dim bid As String = TryCast(ctrl.Tag, String)
        If String.IsNullOrEmpty(bid) Then Continue For

        Dim value As String = GetControlValue(ctrl)
        list.Add(New DetailItem With {.BID = bid, .Value = value})
    Next

    Return list
End Function
 获取变更控件（用于更新） 
Public Function ExtractChangedItems(parent As Control, originalDict As Dictionary(Of String, String)) As List(Of DetailItem)
    Dim list As New List(Of DetailItem)

    For Each ctrl In parent.Controls
        Dim bid As String = TryCast(ctrl.Tag, String)
        If String.IsNullOrEmpty(bid) Then Continue For

        Dim currentValue As String = GetControlValue(ctrl)
        If Not originalDict.ContainsKey(bid) Then Continue For

        If currentValue <> originalDict(bid) Then
            list.Add(New DetailItem With {.BID = bid, .Value = currentValue})
        End If
    Next

    Return list
End Function

控件取值工具函数
Public Function GetControlValue(ctrl As Control) As String
    If TypeOf ctrl Is TextBox Then
        Return DirectCast(ctrl, TextBox).Text
    ElseIf TypeOf ctrl Is ComboBox Then
        Return DirectCast(ctrl, ComboBox).Text
    ElseIf TypeOf ctrl Is CheckBox Then
        Return If(DirectCast(ctrl, CheckBox).Checked, "1", "0")
    ElseIf TypeOf ctrl Is DateTimePicker Then
        Return DirectCast(ctrl, DateTimePicker).Value.ToString("yyyy-MM-dd")
    End If
    Return ""
End Function

使用方法示例
' 新登录场景
Dim items = ExtractAllItems(Me.Panel1) ' 或当前TabPage
For Each item In items
    InsertDetail(conn, trans, aid, item, "新建")
Next

' 编辑更新场景
Dim items = ExtractChangedItems(Me.Panel1, originalDict)
For Each item In items
    InsertDetail(conn, trans, aid, item, "更新")
Next
