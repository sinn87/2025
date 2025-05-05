##✅ 步骤一：定义每个表格的显示列配置
##你可以在模块或类中定义如下字典：

' 定义每个表名对应需要显示的列名数组vb
Private ReadOnly TableVisibleColumns As New Dictionary(Of String, String()) From {
    {"CustomerTable", {"姓名", "年龄", "电话"}},
    {"ManagerTable", {"经理名", "部门"}},
    {"StudentTable", {"学号", "姓名", "年级", "班级"}}
}

✅ 步骤二：根据表名应用显示配置
封装一个统一的设置方法：

vb.net
复制
编辑
''' <summary>
''' 根据表名自动设置 DataGridView 的可见列
''' </summary>
''' <param name="dgv">目标 DataGridView</param>
''' <param name="tableName">表名，用于查找对应列名配置</param>
Public Sub ApplyColumnVisibilityByTableName(dgv As DataGridView, tableName As String)
    If TableVisibleColumns.ContainsKey(tableName) Then
        Dim visibleCols = TableVisibleColumns(tableName)
        ApplyColumnVisibility(dgv, visibleCols)
    End If

✅ 使用示例
当你加载数据绑定完后，只需这样写：

vb.net
复制
编辑
dgv客户.DataSource = dt客户
ApplyColumnVisibilityByTableName(dgv客户, "CustomerTable")

✅ 总结优势
特性	好处
字典集中管理	各表显示逻辑清晰、集中，修改列只改一处
表名驱动	可以动态从配置、数据库、Tag 获取表名
封装简洁	不用每次写一堆列隐藏逻辑


2. 统一遍历所有 TabPage 内的 DataGridView
vb.net
复制
编辑
Private Sub btn保存_Click(sender As Object, e As EventArgs) Handles btn保存.Click
    Try
        For Each tp As TabPage In tabControl1.TabPages
            For Each ctrl As Control In tp.Controls
                If TypeOf ctrl Is DataGridView Then
                    Dim dgv As DataGridView = CType(ctrl, DataGridView)
                    If dgv.DataSource IsNot Nothing AndAlso TypeOf dgv.Tag Is GridInfo Then
                        Dim info As GridInfo = CType(dgv.Tag, GridInfo)
                        Dim dt As DataTable = CType(dgv.DataSource, DataTable)
                        SaveChanges(info.TableName, dt, info.OfferingID)
                    End If
                End If
            Next
        Next

        MessageBox.Show("所有数据保存成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    Catch ex As Exception
        MessageBox.Show("保存失败：" & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try
End Sub
