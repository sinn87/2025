SELECT T1.*
FROM T100 AS T1
INNER JOIN (
    SELECT DISTINCT M3.项目ID
    FROM M320 AS M3
    INNER JOIN (
        SELECT M3.账票ID, M3.募集区分ID
        FROM M300 AS M3
        WHERE M3.账票名 = '明星片'
    ) AS M3_filtered
    ON M3.账票ID = M3_filtered.账票ID
    AND M3.募集区分ID = M3_filtered.募集区分ID
) AS M3_final
ON T1.项目ID = M3_final.项目ID
WHERE T1.案件ID = '200501001';

            
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
