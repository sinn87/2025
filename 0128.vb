Imports Microsoft.Office.Interop.Excel ' Excel操作のためのライブラリをインポート

Public Class Form1
    ' フォームが初期化されるときに実行されるコード
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 下拉リストに選択肢を追加
        ComboBox1.Items.Add("機能1")
        ComboBox1.Items.Add("機能2")
        ComboBox1.Items.Add("機能3")

        ' デフォルトで最初の選択肢を設定
        ComboBox1.SelectedIndex = 0
    End Sub

    ' ボタンをクリックしたときに実行されるコード
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' 下拉リストで選択された項目を取得
        Dim selectedItem As String = ComboBox1.SelectedItem?.ToString()

        ' 機能1が選択された場合の処理
        If selectedItem = "機能1" Then
            Call ExecuteFunction1() ' 機能1を実行する
        Else
            ' 他の機能はここに追加する
            MessageBox.Show("この機能はまだ実装されていません。")
        End If
    End Sub

    ' 機能1: Excelファイルを開き、A2セルに"hello"と記入
    Private Sub ExecuteFunction1()
        Try
            ' Excelアプリケーションを起動
            Dim excelApp As New Application

            ' 現在の実行ファイルと同じディレクトリにあるExcelファイルを開く
            Dim filePath As String = System.IO.Path.Combine(Application.StartupPath, "sample.xlsm")
            Dim workbook As Workbook = excelApp.Workbooks.Open(filePath)
            Dim worksheet As Worksheet = workbook.Sheets(1) ' 1枚目のシートを取得

            ' セルA2に"hello"を記入
            worksheet.Range("A2").Value = "hello"

            ' 変更を保存してファイルを閉じる
            workbook.Save()
            workbook.Close()

            ' Excelアプリケーションを終了
            excelApp.Quit()

            ' メモリ解放
            ReleaseObject(worksheet)
            ReleaseObject(workbook)
            ReleaseObject(excelApp)

            ' 完了メッセージを表示
            MessageBox.Show("機能1が正常に実行されました！")

        Catch ex As Exception
            ' エラーが発生した場合の処理
            MessageBox.Show($"エラーが発生しました：{ex.Message}")
        End Try
    End Sub

    ' COMオブジェクトを解放するメソッド
    Private Sub ReleaseObject(ByVal obj As Object)
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





Private Sub Button_Click(sender As Object, e As EventArgs) Handles Button.Click
    Dim allChecked As Boolean = True ' 全ての項目がチェックされていると仮定する

    ' DataGridView のすべての行をループする
    For Each row As DataGridViewRow In dataGridView.Rows
        ' チェックボックス列のチェック状態を確認
        If Convert.ToBoolean(row.Cells("CheckColumn").Value) = False Then
            allChecked = False ' チェックが外れている項目があれば、allChecked を False に設定
            Exit For ' チェックが外れている項目が見つかった時点でループを抜ける
        End If
    Next

    If allChecked Then
        ' すべての項目がチェックされている場合、確認メッセージを表示
        Dim result As DialogResult = MessageBox.Show("已确认需要审查的项目全都审查结束，确定出力资料吗？", "確認", MessageBoxButtons.YesNo)
        If result = DialogResult.Yes Then
            ' ユーザーが「Yes」を選択した場合、相応の操作を実行
            ' ここに必要な処理を追加してください
        End If
    Else
        ' チェックが外れている項目がある場合、警告メッセージを表示
        MessageBox.Show("还有没审查的项目，请确认。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End If
End Sub
