' 监听TabControl切换
Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
    ResizeFormToFitTabPage()
End Sub

' 计算当前TabPage的大小，并调整Form1
Private Sub ResizeFormToFitTabPage()
    If TabControl1.SelectedTab IsNot Nothing Then
        ' 获取当前TabPage的内容大小
        Dim tabPageSize As Size = TabControl1.SelectedTab.PreferredSize

        ' 计算TabControl本身的额外高度（比如标题栏、边距）
        Dim extraHeight As Integer = Me.Height - TabControl1.Height
        Dim extraWidth As Integer = Me.Width - TabControl1.Width

        ' 计算新窗口大小
        Dim newWidth As Integer = tabPageSize.Width + extraWidth
        Dim newHeight As Integer = tabPageSize.Height + extraHeight

        ' 限制窗口最大大小（不超过1440x960）
        newWidth = Math.Min(newWidth, 1440)
        newHeight = Math.Min(newHeight, 960)

        ' 设置窗口大小
        Me.Size = New Size(newWidth, newHeight)

        ' 可选：居中显示
        Me.CenterToScreen()
    End If
End Sub






Private Sub AdjustFormSizeToUserControl(uc As UserControl)
    ' 获取UserControl的大小
    Dim ucSize As Size = uc.Size

    ' 计算窗口需要增加的宽度和高度
    Dim extraWidth As Integer = Me.Width - TabControl1.Width
    Dim extraHeight As Integer = Me.Height - TabControl1.Height

    ' 计算新窗口大小
    Dim newWidth As Integer = ucSize.Width + extraWidth
    Dim newHeight As Integer = ucSize.Height + extraHeight

    ' 限制窗口最大尺寸（不超过1440×960）
    newWidth = Math.Min(newWidth, 1440)
    newHeight = Math.Min(newHeight, 960)

    ' 设置窗口大小
    Me.Size = New Size(newWidth, newHeight)

    ' 可选：让窗口居中
    Me.CenterToScreen()
End Sub


