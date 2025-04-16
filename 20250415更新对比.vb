Public Function GetChangedItems(originalData As Dictionary(Of String, String), currentData As Dictionary(Of String, String)) As List(Of ChangeItem)
    Dim changes As New List(Of ChangeItem)

    For Each key In currentData.Keys
        Dim originalValue As String = ""
        If originalData.ContainsKey(key) Then
            originalValue = originalData(key)
        End If

        Dim newValue As String = currentData(key)

        If originalValue <> newValue Then
            changes.Add(New ChangeItem With {
                .FieldKey = key,
                .OldValue = originalValue,
                .NewValue = newValue
            })
        End If
    Next

    Return changes
End Function


' 1. 页面加载时保存原始数据
Dim _originalData As Dictionary(Of String, String)

Public Sub Initialize(mode As EntryMode, Optional key As String = "")
    If mode = EntryMode.Edit Or mode = EntryMode.ViewOnly Then
        _originalData = LoadDataAsDictionary(key)
        FillControlsFromDict(_originalData)
    Else
        ClearControls()
    End If
End Sub

' 2. 保存时，提取当前数据并做比较
Public Sub SaveData()
    Dim currentData = ExtractDictFromControls()

    ' 可选：校验 currentData

    ' 如果是编辑状态就比对
    If _mode = EntryMode.Edit Then
        Dim changes = GetChangedItems(_originalData, currentData)

        ' 打印或存入日志
        For Each change In changes
            Debug.Print($"{change.FieldKey} 发生变更: {change.OldValue} → {change.NewValue}")
        Next
    End If

    ' 最终保存 currentData（不管是否变更）
End Sub


 Bonus：只保存变更字段？
如果你数据库结构允许增量更新，那就可以这样做：
Dim updatesOnly As New Dictionary(Of String, String)
For Each change In changes
    updatesOnly(change.FieldKey) = change.NewValue
Next

' 然后用 updatesOnly 存入数据库







Public Function ExtractDictFromControls(container As Control) As Dictionary(Of String, String)
    Dim dict As New Dictionary(Of String, String)

    For Each ctrl As Control In container.Controls
        ' 递归处理子容器（比如 GroupBox, Panel, TabPage 等）
        If ctrl.HasChildren Then
            Dim childDict = ExtractDictFromControls(ctrl)
            For Each kvp In childDict
                dict(kvp.Key) = kvp.Value
            Next
        End If

        ' 确保控件设置了 Tag 才处理
        If ctrl.Tag Is Nothing OrElse String.IsNullOrWhiteSpace(ctrl.Tag.ToString()) Then Continue For

        Dim key As String = ctrl.Tag.ToString()
        Dim value As String = ""

        Select Case True
            Case TypeOf ctrl Is TextBox
                value = DirectCast(ctrl, TextBox).Text

            Case TypeOf ctrl Is ComboBox
                value = DirectCast(ctrl, ComboBox).Text

            Case TypeOf ctrl Is CheckBox
                value = If(DirectCast(ctrl, CheckBox).Checked, "1", "0")

            Case TypeOf ctrl Is DateTimePicker
                value = DirectCast(ctrl, DateTimePicker).Value.ToString("yyyy-MM-dd")

        End Select

        dict(key) = value
    Next

    Return dict
End Function

