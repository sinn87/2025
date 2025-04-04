Imports System.Drawing
Imports System.Windows.Forms

Public Class CListView
    Inherits ListView

    ' カスタムプロパティ：行の高さと選択中の色
    Private _itemHeight As Integer = 40
    Private _selectedBackColor As Color = Color.LightGreen
    Private _normalBackColor As Color = Color.White

    ' 行の高さを設定または取得
    Public Property ItemHeight As Integer
        Get
            Return _itemHeight
        End Get
        Set(value As Integer)
            _itemHeight = value
            Me.Invalidate()
        End Set
    End Property

    ' 選択中の背景色を設定または取得
    Public Property SelectedBackColor As Color
        Get
            Return _selectedBackColor
        End Get
        Set(value As Color)
            _selectedBackColor = value
            Me.Invalidate()
        End Set
    End Property

    ' 通常時の背景色を設定または取得
    Public Property NormalBackColor As Color
        Get
            Return _normalBackColor
        End Get
        Set(value As Color)
            _normalBackColor = value
            Me.Invalidate()
        End Set
    End Property

    ' コンストラクタ
    Public Sub New()
        MyBase.New()

        ' カスタム描画を有効にする
        Me.OwnerDraw = True
        Me.View = View.Details
        Me.FullRowSelect = True
        Me.GridLines = False

        ' イベントをバインドする
        AddHandler Me.DrawItem, AddressOf CListView_DrawItem
        AddHandler Me.DrawSubItem, AddressOf CListView_DrawSubItem
        AddHandler Me.DrawColumnHeader, AddressOf CListView_DrawColumnHeader
    End Sub

    ' 行を描画するイベント
    Private Sub CListView_DrawItem(sender As Object, e As DrawListViewItemEventArgs)
        Dim isSelected As Boolean = (e.State And ListViewItemStates.Selected) = ListViewItemStates.Selected

        ' 背景色を設定
        If isSelected Then
            e.Graphics.FillRectangle(New SolidBrush(_selectedBackColor), e.Bounds)
        Else
            e.Graphics.FillRectangle(New SolidBrush(_normalBackColor), e.Bounds)
        End If

        ' 文字を描画
        e.Graphics.DrawString(e.Item.Text, e.Item.Font, Brushes.Black, e.Bounds.Left, e.Bounds.Top + 10)
    End Sub

    ' サブアイテムを描画するイベント
    Private Sub CListView_DrawSubItem(sender As Object, e As DrawListViewSubItemEventArgs)
        Dim isSelected As Boolean = e.Item.Selected

        ' 背景色を設定
        If isSelected Then
            e.Graphics.FillRectangle(New SolidBrush(_selectedBackColor), e.Bounds)
        Else
            e.Graphics.FillRectangle(New SolidBrush(_normalBackColor), e.Bounds)
        End If

        ' サブアイテムの文字を描画
        e.Graphics.DrawString(e.SubItem.Text, e.Item.Font, Brushes.Black, e.Bounds.Left, e.Bounds.Top + 10)
    End Sub

    ' ヘッダーを描画するイベント
    Private Sub CListView_DrawColumnHeader(sender As Object, e As DrawListViewColumnHeaderEventArgs)
        e.DrawDefault = True ' システム標準のヘッダー描画を使用
    End Sub

    ' サイズ変更時に列幅を自動調整
    Protected Overrides Sub OnSizeChanged(e As EventArgs)
        MyBase.OnSizeChanged(e)
        For Each column As ColumnHeader In Me.Columns
            column.Width = Me.Width \ Me.Columns.Count
        Next
    End Sub
End Class
