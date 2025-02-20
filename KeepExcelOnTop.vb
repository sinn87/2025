Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
    ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal uFlags As Long) As Long

Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_SHOWWINDOW = &H40

Sub KeepExcelOnTop()
    Dim hWnd As LongPtr
    ' 获取 Excel 窗口的句柄
    hWnd = FindWindow("XLMAIN", Application.Caption)
    
    If hWnd <> 0 Then
        ' 让 Excel 窗口置顶
        SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    End If
End Sub






Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
    ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal uFlags As Long) As Long

Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_SHOWWINDOW = &H40

Sub ShowPrintPreview_B()
    ' ① b.xltm の印刷プレビューを開く
    ActiveSheet.PrintPreview
    
    ' ② 印刷プレビューウィンドウのハンドルを取得
    Dim hWndPreview As LongPtr
    Dim i As Integer
    For i = 1 To 50  ' 最大 5 秒待機
        hWndPreview = FindWindow("bosa_sdm_XL9", vbNullString)
        If hWndPreview <> 0 Then Exit For
        Application.Wait Now + TimeValue("00:00:0.1")
    Next i
    
    ' ③ b.xltm の印刷プレビューを最前面に表示
    If hWndPreview <> 0 Then
        SetWindowPos hWndPreview, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    End If

    ' ④ 5 秒待機
    Application.Wait Now + TimeValue("00:00:5")
    
    ' ⑤ a.xltm のマクロを実行し、MsgBox を表示
    Call Application.Run("a.xltm!ShowSaveMessage_A")
End Sub


Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
    ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal uFlags As Long) As Long

Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_SHOWWINDOW = &H40

Sub ShowSaveMessage_A()
    ' ① a.xltm で MsgBox を表示
    MsgBox "PDF を保存しますか？", vbInformation, "保存確認"

    ' ② MsgBox のウィンドウハンドルを取得
    Dim hWndMsgBox As LongPtr
    Dim i As Integer
    For i = 1 To 50  ' 最大 5 秒待機
        hWndMsgBox = FindWindow(vbNullString, "保存確認") ' MsgBox のタイトル
        If hWndMsgBox <> 0 Then Exit For
        Application.Wait Now + TimeValue("00:00:0.1")
    Next i
    
    ' ③ MsgBox を最前面に表示
    If hWndMsgBox <> 0 Then
        SetWindowPos hWndMsgBox, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
    End If
End Sub
