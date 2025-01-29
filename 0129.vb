Imports Microsoft.Office.Interop.Excel

Public Class ExcelHelper
    ' Excelを開いてシートを取得する
    ' path: Excelファイルのパス
    Public Shared Sub WriteToExcel(ByVal path As String)
        ' Excelアプリケーションを作成
        Dim excelApp As New Application()
        Dim workbook As Workbook = Nothing
        Dim worksheet As Worksheet = Nothing

        Try
            ' Excelファイルを開く
            workbook = excelApp.Workbooks.Open(path)
            worksheet = CType(workbook.Sheets(1), Worksheet)

            ' A2セルに "Hello" を書き込む
            worksheet.Cells(2, 1).Value = "Hello"

            ' 保存して閉じる
            workbook.Save()
            workbook.Close()

        Catch ex As Exception
            ' エラーハンドリング
            MsgBox("Excel処理中にエラーが発生しました: " & ex.Message)
        Finally
            ' Excelオブジェクトを解放
            If worksheet IsNot Nothing Then ReleaseObject(worksheet)
            If workbook IsNot Nothing Then ReleaseObject(workbook)
            If excelApp IsNot Nothing Then
                excelApp.Quit()
                ReleaseObject(excelApp)
            End If
        End Try
    End Sub

    ' COMオブジェクトを解放する
    Private Shared Sub ReleaseObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            End If
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Class












            
            Imports System.Data.SqlClient

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 调用方法填充 ComboBox
        LoadComboBoxData()
    End Sub

    Private Sub LoadComboBoxData()
        ' 数据库连接字符串，请替换为你的实际连接信息
        Dim connStr As String = "Server=你的服务器地址;Database=你的数据库名;User Id=你的用户名;Password=你的密码;"
        
        ' SQL 查询（假设第三列名为 column3）
        Dim query As String = "SELECT column3 FROM a"

        ' 创建数据库连接
        Using conn As New SqlConnection(connStr)
            Try
                conn.Open()
                ' 创建命令
                Using cmd As New SqlCommand(query, conn)
                    ' 执行查询
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        ' 清空 ComboBox 以防止重复加载
                        ComboBox1.Items.Clear()

                        ' 读取数据并填充 ComboBox
                        While reader.Read()
                            ComboBox1.Items.Add(reader("column3").ToString())
                        End While
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("数据库连接错误：" & ex.Message)
            End Try
        End Using
    End Sub
End Class
