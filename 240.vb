Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.FormBorderStyle = FormBorderStyle.None  ' 去掉系统默认边框
    End Sub

    Private Sub Form1_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        Dim borderColor As Color = Color.Blue  ' 设定边框颜色
        Dim borderWidth As Integer = 6  ' 设定边框宽度
        Dim rect As Rectangle = New Rectangle(0, 0, Me.Width - 1, Me.Height - 1)
        Using pen As New Pen(borderColor, borderWidth)
            e.Graphics.DrawRectangle(pen, rect)
        End Using
    End Sub
End Class
