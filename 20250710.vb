Private Sub SetSelectedMode(mode As FormMode)
    ' 当前模式变量
    currentMode = mode

    ' 遍历所有Label
    For Each ctrl In {lblView, lblEdit, lblReview, lblExport, lblClose, lblAbort}
        Dim thisMode As FormMode = CType(ctrl.Tag, FormMode)

        ' 设置字体颜色
        ctrl.ForeColor = If(thisMode = mode, Color.DodgerBlue, Color.Black)

        ' 设置下划线（假设下划线是和Label一一对应的Panel，如 pnlLineView 等）
        Dim underlinePanel As Panel = Me.Controls("pnlLine" & thisMode.ToString())
        If underlinePanel IsNot Nothing Then
            underlinePanel.BackColor = If(thisMode = mode, Color.DodgerBlue, Color.Transparent)
        End If
    Next
End Sub


Private Sub lbl_Click(sender As Object, e As EventArgs) _
    Handles lblView.Click, lblEdit.Click, lblReview.Click, lblExport.Click, lblClose.Click, lblAbort.Click

    Dim lbl = DirectCast(sender, Label)
    Dim mode As FormMode = CType(lbl.Tag, FormMode)

    SetMode(mode)
    SetSelectedMode(mode)
End Sub


| 模式        | Label 名称    | 下划线Panel名称         |
| --------- | ----------- | ------------------ |
| View      | `lblView`   | `pnlLineView`      |
| Edit      | `lblEdit`   | `pnlLineEdit`      |
| Review    | `lblReview` | `pnlLineReview`    |
| Export    | `lblExport` | `pnlLineExport`    |
| CloseCase | `lblClose`  | `pnlLineCloseCase` |
| Abort     | `lblAbort`  | `pnlLineAbort`     |
