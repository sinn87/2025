' 模块名：DbDataValidator.vb
Imports System.Data.OleDb

Module DbDataValidator

    ' 检查指定的 DataTable 数据是否仍为数据库中该 ID 的最新数据
    Public Function IsDataStillLatest(dt As DataTable, conn As OleDbConnection, tableName As String) As Boolean
        Dim keys As New List(Of String)

        ' Step 1: 收集主键组合 (案件ID + ID)
        For Each row As DataRow In dt.Rows
            If row.RowState <> DataRowState.Unchanged Then
                Dim caseID = row("案件ID", DataRowVersion.Original).ToString()
                Dim id = row("ID", DataRowVersion.Original).ToString()
                keys.Add($"('{caseID}', '{id}')")
            End If
        Next

        If keys.Count = 0 Then Return True ' 没有修改过的数据

        ' Step 2: 查询这些键组合对应的最新更新日期
        Dim keyClause = String.Join(",", keys)
        Dim sql = $"SELECT 案件ID, ID, MAX([更新日期]) AS 最新更新日期 FROM {tableName} WHERE (案件ID, ID) IN ({keyClause}) GROUP BY 案件ID, ID"

        Dim latestDateDict As New Dictionary(Of String, Date)
        Using cmd As New OleDbCommand(sql, conn)
            Using reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim key = reader("案件ID").ToString() & "|" & reader("ID").ToString()
                    latestDateDict(key) = Convert.ToDateTime(reader("最新更新日期"))
                End While
            End Using
        End Using

        ' Step 3: 查询每组的状态
        Dim finalDict As New Dictionary(Of String, Integer)
        For Each kvp In latestDateDict
            Dim parts = kvp.Key.Split("|"c)
            Dim sql2 = $"SELECT 状态 FROM {tableName} WHERE 案件ID = ? AND ID = ? AND [更新日期] = ?"
            Using cmd2 As New OleDbCommand(sql2, conn)
                cmd2.Parameters.AddWithValue("?", parts(0))
                cmd2.Parameters.AddWithValue("?", parts(1))
                cmd2.Parameters.AddWithValue("?", kvp.Value)
                Dim status = cmd2.ExecuteScalar()
                If status IsNot Nothing Then
                    finalDict(kvp.Key) = Convert.ToInt32(status)
                End If
            End Using
        Next

        ' Step 4: 与原始数据比对
        For Each row As DataRow In dt.Rows
            If row.RowState <> DataRowState.Unchanged Then
                Dim key = row("案件ID", DataRowVersion.Original).ToString() & "|" & row("ID", DataRowVersion.Original).ToString()
                Dim origDate = Convert.ToDateTime(row("更新日期", DataRowVersion.Original))
                Dim origStatus = Convert.ToInt32(row("状态", DataRowVersion.Original))

                If Not latestDateDict.ContainsKey(key) Then Return False
                If latestDateDict(key) <> origDate Then Return False
                If finalDict(key) <> origStatus Then Return False
            End If
        Next

        Return True
    End Function

End Module
