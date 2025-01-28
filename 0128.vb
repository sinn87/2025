Public Class Form1
    ' フォームが初期化されるときに実行されるコード
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 下拉リストに初期値を追加する
        ComboBox1.Items.Add("選択肢1")
        ComboBox1.Items.Add("選択肢2")
        ComboBox1.Items.Add("選択肢3")

        ' デフォルトで最初の選択肢を選択する
        ComboBox1.SelectedIndex = 0
    End Sub

    ' ボタンをクリックしたときに実行されるコード
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' 選択された項目を取得する
        Dim selectedItem As String = ComboBox1.SelectedItem?.ToString()

        ' 選択された項目があるかどうかを確認
        If selectedItem IsNot Nothing Then
            ' 選択された項目をメッセージボックスで表示
            MessageBox.Show($"あなたが選択したのは：{selectedItem}")
        Else
            ' 項目が選択されていない場合のメッセージ
            MessageBox.Show("まず、項目を選択してください！")
        End If
    End Sub
End Class
