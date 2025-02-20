Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Declare PtrSafe Function FindWindowEx Lib "user32" ( _
    ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, _
    ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr

Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
    ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal uFlags As Long) As Long

Const HWND_TOPMOST = -1        ' 置顶
Const HWND_NOTOPMOST = -2      ' 取消置顶
Const SWP_NOMOVE = &H2         ' 保持当前位置
Const SWP_NOSIZE = &H1         ' 保持当前大小
Const SWP_SHOWWINDOW = &H40    ' 显示窗口

Sub SetBXLTMOnTop()
    Dim hWndExcel As LongPtr
    Dim hWndWorkbook As LongPtr
    Dim wbName As String
    wbName = "b.xltm - Excel"  ' 窗口标题，确保和实际文件名匹配
    
    ' 找到 Excel 主窗口
    hWndExcel = FindWindow("XLMAIN", vbNullString)
    
    ' 在 Excel 窗口内找到 "b.xltm" 这个文件的窗口
    If hWndExcel <> 0 Then
        hWndWorkbook = FindWindowEx(hWndExcel, 0, "XLMAIN", wbName)
        
        If hWndWorkbook <> 0 Then
            ' 置顶 b.xltm
            SetWindowPos hWndWorkbook, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
            
            ' 5 秒后取消置顶
            Application.Wait Now + TimeValue("00:00:5")
            
            SetWindowPos hWndWorkbook, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
        End If
    End If
End Sub



