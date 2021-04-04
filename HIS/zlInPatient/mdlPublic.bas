Attribute VB_Name = "mdlPublic"
Option Explicit '要求变量声明

'系统公用变量

Public gfrmMain As Object                   '导航台窗体
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gcolPrivs As Collection              '记录内部模块的权限
Public gMainPrivs As String                 '调用主界面所具有的权限,注意非内部模块权限
Public gstrPrivs As String                  '当前用户具有的当前模块的功能
Public gstrSysName As String                '系统名称
Public gstrDBUser As String                 '当前数据库用户
Public gstrUnitName As String               '用户单位名称
Public gstrProductName As String            'OEM产品名称
Public glngSys As Long
Public glngModul As Long

Public gstrSQL As String
Public gblnOK As Boolean

Public gblnLED As Boolean       '收预交款时是否使用LED语音报价
Public gblnLedWelcome As Boolean '是否在预交输完病人后提示欢迎信息

Public gobjSquare As SquareCard  '卡结算部件

'用户信息------------------------
Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    部门名称 As String
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO

'---------------------------------------
'问题27554 by lesfeng 2010-01-19
Public glngTXTProc As Long '保存默认的消息函数的地址
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const HTCAPTION = 2
Public Const GWL_WNDPROC = -4
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Public Const SRCCOPY = &HCC0020
Public Const SM_CYCAPTION = 4
Public Const CB_GETDROPPEDSTATE = &H157

'下列语句用于检测是否合法调用
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function MoveObj(lngHwnd As Long) As RECT
'功能：在对象的MouseDown事件中调用,对象必须具有Hwnd属性
'返回：相对屏幕的像素值
   
    Dim vPos As RECT
    ReleaseCapture
    SendMessage lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    GetWindowRect lngHwnd, vPos
    MoveObj = vPos
End Function

Public Function GetColNum(lvwTemp As ListView, strHead As String) As Integer
    Dim i As Integer
    For i = 1 To lvwTemp.ColumnHeaders.Count
        If lvwTemp.ColumnHeaders(i).Text = strHead Then GetColNum = i: Exit Function
    Next
End Function

Public Sub SetCenter(frm As Form)
'功能：将窗体定位在屏幕中央
    frm.Left = (Screen.width - frm.width) / 2
    frm.Top = (Screen.Height - frm.Height) / 2
End Sub

Public Function CheckLen(txt As TextBox, intLen As Integer) As Boolean
'功能：检查工本框的真实长度是否在指定限制长度内
    If LenB(StrConv(txt.Text, vbFromUnicode)) > intLen Then
        MsgBox Mid(txt.Name, 4) & "只允许输入 " & intLen & " 个字符或 " & intLen \ 2 & " 个汉字！", vbInformation, gstrSysName
        txt.SetFocus: Exit Function
    End If
    CheckLen = True
End Function

Public Function CaptionHeight() As Long
'功能:返回系统窗体标题栏高度(以象素为单位)
    CaptionHeight = GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY
End Function

Public Sub SetItemInfo(lvw As Object, pan As Object)
'功能：根据Listview当前选中行，显示在状态条上
    Dim i As Integer, strInfo As String
    
    If lvw.ListItems.Count = 0 Then Exit Sub
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    If lvw.SelectedItem.Text <> "" Then
        strInfo = "/" & lvw.ColumnHeaders(1).Text & ":" & lvw.SelectedItem.Text
    End If
    
    For i = 2 To lvw.ColumnHeaders.Count
        If lvw.SelectedItem.SubItems(i - 1) <> "" Then
            strInfo = strInfo & "/" & lvw.ColumnHeaders(i).Text & ":" & lvw.SelectedItem.SubItems(i - 1)
        End If
    Next
    If strInfo <> "" Then pan.Text = Mid(strInfo, 2)
End Sub

Public Sub CheckInputLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Public Function SetCboDefault(cbo As ComboBox) As Integer
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = 1 Then
            cbo.ListIndex = i
            SetCboDefault = i: Exit Function
        End If
    Next
End Function

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    '问题27554 by lesfeng 2010-01-19 lngTXTProc 修改为glngTXTProc
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function

Public Sub SetGridWidth(msh As Control, frmParent As Object)
'功能：自动调整表格列宽,以最小适合为准
    Dim blnRedraw As Boolean
    Dim blnDo As Boolean, i As Long, j As Long
    Dim lngStart As Long, lngEnd As Long, lngMaxWidth As Long
        
    blnRedraw = msh.Redraw
    msh.Redraw = False
    lngStart = IIf(msh.FixedRows = 0, 0, msh.FixedRows - 1)
    lngEnd = msh.Rows - 1
    
    For i = 0 To msh.Cols - 1
        lngMaxWidth = 0
        For j = lngStart To lngEnd
            blnDo = True
            If msh.MergeRow(j) Then
                If i > 0 Then If msh.TextMatrix(j, i) = msh.TextMatrix(j, i - 1) Then blnDo = False
                If i < msh.Cols - 1 Then If msh.TextMatrix(j, i) = msh.TextMatrix(j, i + 1) Then blnDo = False
            End If
            If blnDo Then
                If Len(msh.TextMatrix(j, i)) > Len(msh.TextMatrix(lngMaxWidth, i)) Then
                    lngMaxWidth = j
                End If
            End If
        Next
        msh.ColWidth(i) = IIf(frmParent.TextWidth(msh.TextMatrix(lngMaxWidth, i)) > 3000, 3000, frmParent.TextWidth(msh.TextMatrix(lngMaxWidth, i)) + 90)
    Next
    
    msh.Redraw = blnRedraw
End Sub

Public Function CheckFormInput(objForm As Object, Optional ByVal strIgnore As String, Optional ByVal strToNumText As String = "") As Boolean
    '参数:strIgnore-不检查的控件名,允许有多个,可用,号等分隔
    '参数:strToNumText--需要进行将千分位格式的金额转成正常金额格式的文本控件名称,允许有多个,可用,号等分隔
    Dim obj As Object, strText As String
    
    On Error Resume Next
    For Each obj In objForm.Controls
        If InStr("TextBox,ComboBox", TypeName(obj)) > 0 Then
            If obj.Visible And obj.Enabled And Not obj.Locked Then
                strText = ""
                Select Case TypeName(obj)
                Case "TextBox"
                    strText = obj.Text
                    If InStr(1, "," & UCase(strToNumText) & ",", "," & UCase(obj.Name) & ",") > 0 Then
                        strText = StrToNum(strText)
                    End If
                Case "ComboBox"
                    If obj.Style = 0 Then strText = obj.Text
                End Select
                If InStr(UCase(strIgnore), UCase(obj.Name)) = 0 Then
                    If InStr(strText, "'") > 0 _
                        Or InStr(strText, "|") > 0 _
                        Or InStr(strText, "~") > 0 _
                        Or InStr(strText, "^") > 0 Then
                        MsgBox "输入数据中包含非法字符！", vbInformation, gstrSysName
                        obj.SelStart = 0: obj.SelLength = Len(obj.Text)
                        obj.SetFocus: Exit Function
                    End If
                End If
            End If
        End If
    Next
    CheckFormInput = True
End Function

Public Function GetIDDate(ID As String) As String
'功能：根据身份证号返回出生日期,格式"yyyy-MM-dd"
'参数：ID=身份证号,应该为15位或18位
    Dim strTmp As String
    
    If Len(ID) = 15 Then
        strTmp = Mid(ID, 7, 6)
        If Len(strTmp) = 6 And IsNumeric(strTmp) Then
            strTmp = "19" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2)
        End If
    ElseIf Len(ID) = 18 Then
        strTmp = Mid(ID, 7, 8)
        If Len(strTmp) = 8 And IsNumeric(strTmp) Then
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2)
        End If
    End If
    If IsDate(strTmp) Then GetIDDate = strTmp
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim intDo As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(strPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Sub CboLoadData(ByRef cbo As ComboBox, ByRef rsTmp As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
    '功能:装载数据入指定的组合下拉框或网格中的下拉框中
    '参数:cbo   要装载记录集的下拉框控件
    '     rsTmp     记录集数据,要求至少有三个数据项,Id,编码，名称
    '     blnClear    装载时是否清楚原有的下拉数据,缺省为True
    
    If rsTmp.Fields.Count < 3 Then Exit Sub
    If blnClear = True Then cbo.Clear
    
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        While Not rsTmp.EOF
            cbo.AddItem rsTmp.Fields(1).Value & "-" & rsTmp.Fields(2).Value
            cbo.ItemData(cbo.NewIndex) = Val(rsTmp.Fields(0).Value)
            rsTmp.MoveNext
        Wend
        rsTmp.MoveFirst
    End If
End Sub

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    
    If gfrmMain Is Nothing Then CheckValid = True: Exit Function
    
    '获取注册表后，马上清零
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "公共全局", "公共", 0)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", 0)
    blnValid = (intAtom <> 0)
    
    '如果存在，则对串进行解析
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '如果为空，则表示非法
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '判断时间间隔是否大于1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '如果相等，则通过
                    Else
                        '不等，表示存在进位，则分应该为零
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function



