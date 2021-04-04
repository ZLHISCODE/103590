Attribute VB_Name = "mdlCISCore"
Option Explicit

Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gobjCISCore As clsCISCore

Public gstrSysName As String                '系统名称
Public gstrAviPath As String                'AVI文件的存放目录
Public gstrHelpPath As String
Public gstrDBUser As String
Public gblnOK As Boolean

Public glngSys As Long                      '用来记录系统号
Public gstrPrivs As String                  '用来记录权限

Public gstrMatch As String                  '根据本地参数“匹配模式”确定的左匹配符号
Public gstrSql As String

Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO
Public grsDuty As ADODB.Recordset '存放医生职务
Public grsSysPars As ADODB.Recordset

Public glngPen As Long '当前画笔对象
Public glngBrush As Long '当前刷子对象

Public gcurPenColor As Long '当前使用的线条色
Public gcurPenStyle As Byte '当前使用的线型
Public gcurPenWidth As Byte '当前使用的线宽
Public gcurFillColor As Long '当前使用的填充色
Public gcurFillStyle As Integer '当前使用的填充样式

Public glngTXTProc As Long '保存默认的消息函数的地址

Public Const LONG_MAX = 2147483647 'Long型最大值
'======================================================================================================================
'API定义部分
Public Const WS_CHILD = &H40000000
Public Const GWL_STYLE = (-16)
Public Const GWL_WNDPROC = -4
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Type POINTAPI
     x As Long
     y As Long
End Type
Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type
Public Const WM_UNDO = &H304
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'输入控制
Public Const EM_LINESCROLL = &HB6 'lngW=横向行数,lngL=纵向行数
Public Const EM_SCROLL = &HB5 '按滚动条几下
Public Const EM_GETFIRSTVISIBLELINE = &HCE 'lngR(>=0)
Public Const EM_GETLINECOUNT = &HBA 'lngR(>=1,包含自动折的行)
Public Const EM_LINELENGTH = &HC1 '第一行未折行前有效
Public Const EM_GETSEL = &HB0
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB

Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
'API作图
'---------------------------------------------------------------------------------------------------------------------
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type
'Style
Public Const BS_HATCHED = 2
Public Const BS_NULL = 1
Public Const BS_SOLID = 0
'Hatch
Public Const HS_BDIAGONAL = 3               '  /////
Public Const HS_CROSS = 4                   '  +++++
Public Const HS_DIAGCROSS = 5               '  xxxxx
Public Const HS_FDIAGONAL = 2               '  \\\\\
Public Const HS_HORIZONTAL = 0              '  -----
Public Const HS_VERTICAL = 1                '  |||||
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Const PS_NULL = 5
Public Const PS_SOLID = 0
Public Const PS_DOT = 2
Public Const PS_DASH = 1
Public Const PS_DASHDOT = 3
Public Const PS_DASHDOTDOT = 4
Public Const PS_INSIDEFRAME = 6
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H440328
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
'======================================================================================================================
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 '浅凹下
Public Const BDR_RAISEDINNER = &H4 '浅凸起
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '深凸起
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '深凹下
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame边线样式
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '反Frame边线样式
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000


Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = _
        " Select A.ID,C.部门ID,A.编号,A.简码,A.姓名,B.用户名" & _
        " From 人员表 A,上机人员表 B,部门人员 C" & _
        " Where A.ID = B.人员ID And A.ID = C.人员ID And C.缺省 = 1 And Upper(B.用户名) = USER"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, "mdlCISBase", strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = IIf(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.简码 = IIf(IsNull(rsTmp!简码), "", rsTmp!简码)
        UserInfo.姓名 = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
        UserInfo.用户名 = IIf(IsNull(rsTmp!用户名), "", rsTmp!用户名)
        gstrDBUser = UserInfo.用户名
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ArrayIcons(objLvw As ListView, Optional intBegin As Integer = 1, Optional blnShow As Boolean)
'功能：根据第一个图标的位置重新排列所有图标
    Dim i As Integer, t As Long
    Dim r As RECT

    Call GetClientRect(objLvw.hWnd, r)
    
    If blnShow Then
        If objLvw.ListItems(intBegin).Top < 30 Then
           objLvw.ListItems(intBegin).Top = 30
        ElseIf objLvw.ListItems(intBegin).Top + objLvw.ListItems(intBegin).Height > (r.Bottom - r.Top) * Screen.TwipsPerPixelY Then
            objLvw.ListItems(intBegin).Top = (r.Bottom - r.Top) * Screen.TwipsPerPixelY - objLvw.ListItems(intBegin).Height
        End If
    End If
    
    '下面的图标
    t = objLvw.ListItems(intBegin).Top
    For i = intBegin To objLvw.ListItems.Count
        With objLvw.ListItems(i)
            'Item的Width包含文字部分,Left仅指图标
            .Left = ((r.Right - r.Left - objLvw.Icons.ImageWidth) * Screen.TwipsPerPixelX) / 2
            .Top = t
            t = t + .Height
        End With
    Next
    
    '上面的图标
    t = objLvw.ListItems(intBegin).Top
    For i = intBegin To 1 Step -1
        With objLvw.ListItems(i)
            'Item的Width包含文字部分,Left仅指图标
            .Left = ((r.Right - r.Left - objLvw.Icons.ImageWidth) * Screen.TwipsPerPixelX) / 2
            .Top = t
            t = t - .Height
        End With
    Next
End Sub

Public Function OpenRecord(rsTmp As ADODB.Recordset, strSQL As String, ByVal strTitle As String, _
    Optional CursorType As CursorTypeEnum = adOpenKeyset, Optional LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, strTitle, strSQL)
    rsTmp.Open strSQL, gcnOracle, CursorType, LockType
    Call SQLTest
    
    Set OpenRecord = rsTmp
End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '替换为"?"参数
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        Case "Variant" '不明确类型
            strLog = Replace(strLog, "[" & i & "]", "?")
        End Select
    Next

    '清除原有参数:不然不能重复执行
    cmdData.CommandText = "" '不为空有时清除参数出错
    Do While cmdData.Parameters.Count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '创建新的参数
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "String" '字符
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, 500, varValue)
        Case "Variant" '不明确类型
        End Select
    Next

    '执行返回记录集
    If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '这句比较慢
    End If
    cmdData.CommandText = strSQL
    
    Call SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    Call SQLTest
End Function

Public Sub ResetDrawStyle()
'功能：删除当前设置的画笔和画刷
    If glngBrush <> 0 Then DeleteObject glngBrush
    If glngPen <> 0 Then DeleteObject glngPen
End Sub

Public Sub SetDrawStyleFromValue(lngHDc As Long, PenColor As Long, PenStyle As Byte, PenWidth As Byte, FillColor As Long, FillStyle As Integer)
'功能：根据指定值设置当前的画笔的画刷
    Dim vBrush As LOGBRUSH
    Dim lngPen As Long, lngBrush As Long
    
    If glngBrush <> 0 Then DeleteObject glngBrush
    If glngPen <> 0 Then DeleteObject glngPen
    
    '画笔
    lngPen = CreatePen(PenStyle, IIf(PenWidth < 1, 1, PenWidth), PenColor)
    glngPen = SelectObject(lngHDc, lngPen)
    
    '画刷
    vBrush.lbColor = FillColor
    If FillStyle = -1 Then
        vBrush.lbStyle = BS_NULL
    ElseIf FillStyle = -2 Then
        vBrush.lbStyle = BS_SOLID
    Else
        vBrush.lbStyle = BS_HATCHED
        vBrush.lbHatch = FillStyle
    End If
    lngBrush = CreateBrushIndirect(vBrush)
    glngBrush = SelectObject(lngHDc, lngBrush)
End Sub

Public Sub TextOut(objOut As Object, ByVal strOut As String, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal strFont As String, ByVal sngScale As Single)
'功能：在指定设备的指定范围内输出文字
'参数：strFont="字体,字号,字色,0000",sngScale=输出比例
'说明：自动换行,支持回车
    Dim arrFont() As String, arrLine() As String
    Dim lngWidth As Long, lngW As Long, i As Integer
    
    If Trim(Replace(strOut, vbCrLf, "")) = "" Then Exit Sub
    If strFont = "" Then Exit Sub
    
    arrFont = Split(strFont, ",")
    objOut.FontName = arrFont(0)
    objOut.FontSize = CSng(arrFont(1)) * sngScale
    objOut.ForeColor = CLng(arrFont(2))
    objOut.FontBold = Mid(arrFont(3), 1, 1) = "1"
    objOut.FontItalic = Mid(arrFont(3), 2, 1) = "1"
    objOut.FontUnderline = Mid(arrFont(3), 3, 1) = "1"
    objOut.FontStrikethru = Mid(arrFont(3), 4, 1) = "1"
    
    X1 = X1 * sngScale: Y1 = Y1 * sngScale
    X2 = X2 * sngScale: Y2 = Y2 * sngScale
        
    strOut = Replace(strOut, vbCrLf, "'")
    lngWidth = X2 - X1
    ReDim arrLine(0)
    For i = 1 To Len(strOut)
        If Mid(strOut, i, 1) = "'" Then
            lngW = 0
            Do While Mid(strOut, i, 1) = "'"
                ReDim Preserve arrLine(UBound(arrLine) + 1)
                i = i + 1
            Loop
        End If
        If i <= Len(strOut) Then
            lngW = lngW + objOut.TextWidth(Mid(strOut, i, 1))
            If lngW > lngWidth Then
                ReDim Preserve arrLine(UBound(arrLine) + 1)
                lngW = 0
            End If
            arrLine(UBound(arrLine)) = arrLine(UBound(arrLine)) & Mid(strOut, i, 1)
        End If
    Next
    objOut.CurrentY = Y1 + 2
    For i = 0 To UBound(arrLine)
        objOut.CurrentX = X1 + 2
        objOut.Print arrLine(i)
    Next
End Sub

Public Function ReadCaseMap(lngID As Long) As StdPicture
'功能：根据标记图ID返回图形对象
    Dim lngFileSize As Long, arrBin() As Byte
    Dim strFile As String, intFile As Integer
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 图形 From 病历标记图 Where 元素ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISCore", lngID)
        
    If rsTmp.EOF Then Exit Function
    If IsNull(rsTmp!图形) Then Exit Function
    
    On Error GoTo 0
    
    intFile = FreeFile
    strFile = CurDir & "\zlNewPicture" & Timer & ".pic"
    
    Open strFile For Binary As intFile
    
    lngFileSize = rsTmp.Fields("图形").ActualSize
    ReDim arrBin(lngFileSize - 1) As Byte
    arrBin() = rsTmp.Fields("图形").GetChunk(lngFileSize)
    Put intFile, , arrBin()
    Close intFile
    
    Set ReadCaseMap = VB.LoadPicture(strFile)
    Kill strFile
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ReadVoiceToFile(lngID As Long) As String
'功能：根据病历记录ID返回声音文件
    Dim lngFileSize As Long, arrBin() As Byte
    Dim strFile As String, intFile As Integer
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 声音 From 病人病历录音 Where 病历记录ID=" & lngID
    OpenRecord rsTmp, strSQL, "mdlCISCore"
        
    If rsTmp.EOF Then Exit Function
    If IsNull(rsTmp!声音) Then Exit Function
    
    intFile = FreeFile
    strFile = CurDir & "\zlNewPicture" & Timer & ".Mp3"
    
    Open strFile For Binary As intFile
    
    lngFileSize = rsTmp.Fields("声音").ActualSize
    ReDim arrBin(lngFileSize - 1) As Byte
    arrBin() = rsTmp.Fields("声音").GetChunk(lngFileSize)
    Put intFile, , arrBin()
    Close intFile
    
    ReadVoiceToFile = strFile
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    If Len(strFile) > 0 Then Close intFile: Kill strFile
    Call SaveErrLog
End Function

Public Sub ShowCaseMap(objCaseMap As StdPicture, objMapItems As MapItems, objDraw As Object, _
    Optional x As Long, Optional y As Long, Optional W As Long, Optional H As Long)
'功能：显示病历标记图内容
'参数：objDraw=显示的目标对象,它的ScaleMode必须为Pixel
'      objMapItems=病历中当前项目的标记图内容
'      X,Y,W,H=显示的目标范围,可以不指定,单位为Pixel
    Dim arrTmp() As String, arrXY() As POINTAPI
    Dim i As Integer, j As Integer
    Dim lngW As Long, lngH As Long '图片尺寸
    Dim sngScale As Single
        
    Screen.MousePointer = 11
    LockWindowUpdate objDraw.hWnd
    
    '确定图片尺寸及显示比例
    objDraw.ScaleMode = vbPixels
    lngW = objDraw.ScaleX(objCaseMap.Width, vbHimetric, vbPixels) '将HiMetric转换为Pixel
    lngH = objDraw.ScaleY(objCaseMap.Height, vbHimetric, vbPixels)
    If W = 0 Then W = objDraw.ScaleWidth
    If H = 0 Then H = objDraw.ScaleHeight
    If W / lngW < H / lngH Then
        sngScale = W / lngW
    Else
        sngScale = H / lngH
    End If
    
    objDraw.Cls
    objDraw.PaintPicture objCaseMap, x, y, lngW * sngScale, lngH * sngScale
            
    '具体标记元素
    For i = 1 To objMapItems.Count
        With objMapItems(i)
            If .类型 <> 0 Then
                Call SetDrawStyleFromValue(objDraw.hDC, .线条色, .线型, .线宽 * sngScale, .填充色, .填充方式)
            End If
            Select Case .类型
                Case 0 '文本
                    Call TextOut(objDraw, .内容, (.X1 * sngScale + x) / sngScale, (.Y1 * sngScale + y) / sngScale, (.X2 * sngScale + x) / sngScale, (.Y2 * sngScale + y) / sngScale, .字体, sngScale)
                Case 1 '线条
                    MoveToEx objDraw.hDC, .X1 * sngScale + x, .Y1 * sngScale + y, 0
                    LineTo objDraw.hDC, .X2 * sngScale + x, .Y2 * sngScale + y
                Case 2 '折线
                    arrTmp = Split(.点集, ";")
                    For j = 0 To UBound(arrTmp)
                        ReDim Preserve arrXY(j)
                        arrXY(j).x = CLng(Split(arrTmp(j), ",")(0)) * sngScale + x
                        arrXY(j).y = CLng(Split(arrTmp(j), ",")(1)) * sngScale + y
                    Next
                    Polyline objDraw.hDC, arrXY(0), UBound(arrXY) + 1
                Case 3 '矩形
                    Rectangle objDraw.hDC, .X1 * sngScale + x, .Y1 * sngScale + y, .X2 * sngScale + x, .Y2 * sngScale + y
                Case 4 '多边形
                    arrTmp = Split(.点集, ";")
                    For j = 0 To UBound(arrTmp)
                        ReDim Preserve arrXY(j)
                        arrXY(j).x = CLng(Split(arrTmp(j), ",")(0)) * sngScale + x
                        arrXY(j).y = CLng(Split(arrTmp(j), ",")(1)) * sngScale + y
                    Next
                    Polygon objDraw.hDC, arrXY(0), UBound(arrXY) + 1
                Case 5 '圆
                    Ellipse objDraw.hDC, .X1 * sngScale + x, .Y1 * sngScale + y, .X2 * sngScale + x, .Y2 * sngScale + y
            End Select
        End With
    Next
    objDraw.Refresh
    
    Call ResetDrawStyle
    
    LockWindowUpdate 0
    Screen.MousePointer = 0
End Sub

Public Function EditFlag(frmParent As Object, var元素 As Variant, Optional Flags As Variant, Optional blnViewOnly As Boolean) As MapItems
'功能：在单独的模态窗体中编辑或查看指定的病历标记图
'参数：frmParnet=调用父窗体
'      var元素=标记图元素的编码(字符型)或ID(数字型)
'      Flags=Long型：要修改的"病人病历内容"中标记图元素对应的ID；
'            MapItems：要显示的标注。
'            如果不传，则表示新增标注
'      blnViewOnly=是否只查看，不能编辑
'返回：Mapitems
'      取消编辑或查看模式返回Empty(Not isArray)。
    Dim frmNew As frmMapEdit
    Dim rsTmp As New ADODB.Recordset
    Dim arrSQL() As Variant, strSQL As String
    
    Dim objCaseMap As StdPicture, i As Long
    Dim objMapItems As New MapItems, objMapItem As MapItem
    Dim lngMapID As Long, strMapName As String
    
    Dim iMin As Long, iMax As Long, aItems() As String
    Dim strFont As String, strContent As String, strDots As String
    
    On Error GoTo errH
        
    '读取标记图元素的内容
    If TypeName(var元素) = "String" Then
        strSQL = "Select * From 病历元素目录 Where 编码=[1]"
    Else
        strSQL = "Select * From 病历元素目录 Where ID=[1]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", CStr(var元素))
    If rsTmp.EOF Then Exit Function '必须要有图形背景
    
    lngMapID = rsTmp!ID
    strMapName = rsTmp!名称 & IIf(IsNull(rsTmp!说明), "", "(" & rsTmp!说明 & ")")
    
    Set objCaseMap = ReadCaseMap(rsTmp!ID)
    If objCaseMap Is Nothing Then Exit Function '必须要有图形背景
    
    '读取标记图的标注内容
    If IsEmpty(Flags) Then Flags = 0
    
    If TypeName(Flags) = "Long" Then
        If Flags <> 0 Then
            strSQL = "Select * From 病人病历标记图 Where 病历ID=[1]"
            Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", Flags)
            Do While Not rsTmp.EOF
                With rsTmp
                    objMapItems.Add !类型, zlCommFun.NVL(!内容), _
                        IIf(IsNull(!字体), IIf(!类型 = 0, "宋体,9,0,0000", ""), !字体), _
                        zlCommFun.NVL(!点集), zlCommFun.NVL(!X1, 0), _
                        zlCommFun.NVL(!Y1, 0), zlCommFun.NVL(!X2, 0), _
                        zlCommFun.NVL(!Y2, 0), zlCommFun.NVL(!填充色, &HFFFFFF), _
                        zlCommFun.NVL(!填充方式, -1), zlCommFun.NVL(!线条色, 0), _
                        zlCommFun.NVL(!线型, 0), zlCommFun.NVL(!线宽, 1), "_" & objMapItems.Count + 1
                End With
                rsTmp.MoveNext
            Loop
        End If
    Else
        For i = 1 To Flags.Count
            Set objMapItem = Flags(i)
            '"类型,'内容','字体','点集',X1,Y1,X2,Y2,填充色,填充方式,线条色,线型,线宽"
            With objMapItem
                objMapItems.Add .类型, .内容, .字体, .点集, _
                    .X1, .Y1, .X2, .Y2, .填充色, .填充方式, _
                    .线条色, .线型, .线宽, "_" & objMapItems.Count + 1
            End With
        Next
    End If
    
    On Error GoTo 0
    
    Set frmNew = New frmMapEdit
    frmNew.mblnModi = Not blnViewOnly
    frmNew.mlngMapID = lngMapID
    frmNew.mstrMapName = strMapName
    Set frmNew.mobjCaseMap = objCaseMap
    Set frmNew.mobjMapItems = objMapItems
    frmNew.Show 1, frmParent
    
    If gblnOK Then
        Set EditFlag = frmNew.mobjMapItems
'        For i = 1 To frmNew.mobjMapItems.Count
'            Set objMapItem = frmNew.mobjMapItems(i)
'            '"类型,'内容','字体','点集',X1,Y1,X2,Y2,填充色,填充方式,线条色,线型,线宽"
'            With objMapItem
'                EditFlag.Add .类型, .内容, .字体, .点集, _
'                    .X1, .Y1, .X2, .Y2, .填充色, .填充方式, _
'                    .线条色, .线型, .线宽
'            End With
'        Next
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ShowFlagInOjbect(objDraw As Object, var元素 As Variant, Optional Flags As Variant, Optional x As Long, Optional y As Long, Optional W As Long, Optional H As Long, Optional blnMoved As Boolean = False)
'功能：在指定的对象(PictureBox或Form)上显示标记图
'参数：objDraw=PictureBox或窗体对象,它的ScaleMode必须为Pixel
'      var元素=标记图元素的编码(字符型)或ID(数字型)
'      Flags=Long型："病人病历内容"中标记图元素对应的ID；
'            MapItems：要显示的标注。
'            如果不传,仅显示标记图背景
'      X,Y,W,H=显示的目标客户端范围,可以不指定,单位为Pixel
'说明：可以利用该函数进行打印输出(因为是API作图,因此不能直接将objDraw指定为打印机,而是在PictureBox上按一定比例输出后,取PictureBox.Image输出到打印机)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim objCaseMap As StdPicture, objMapItems As New MapItems
    
    Dim i As Long, iMin As Long, iMax As Long, aItems() As String
    Dim strFont As String, strContent As String, strDots As String
    
    On Error GoTo errH
        
    '读取标记图元素的内容
    If TypeName(var元素) = "String" Then
        strSQL = "Select * From 病历元素目录 Where 编码=[1]"
    Else
        strSQL = "Select * From 病历元素目录 Where ID=[1]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", CStr(var元素))
    If rsTmp.EOF Then Exit Sub '必须要有图形背景
    
    Set objCaseMap = ReadCaseMap(rsTmp!ID)
    If objCaseMap Is Nothing Then Exit Sub '必须要有图形背景
    
    '读取标记图的标注内容
    If IsEmpty(Flags) Then Flags = 0
    
    If TypeName(Flags) = "Long" Then
        If Flags <> 0 Then
            strSQL = "Select * From 病人病历标记图 Where 病历ID=[1]"
            If blnMoved Then
                strSQL = Replace(strSQL, "病人病历标记图", "H病人病历标记图")
            End If
            Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", Flags)
            Do While Not rsTmp.EOF
                With rsTmp
                    objMapItems.Add !类型, zlCommFun.NVL(!内容), _
                        IIf(IsNull(!字体), IIf(!类型 = 0, "宋体,9,0,0000", ""), !字体), _
                        zlCommFun.NVL(!点集), zlCommFun.NVL(!X1, 0), _
                        zlCommFun.NVL(!Y1, 0), zlCommFun.NVL(!X2, 0), _
                        zlCommFun.NVL(!Y2, 0), zlCommFun.NVL(!填充色, &HFFFFFF), _
                        zlCommFun.NVL(!填充方式, -1), zlCommFun.NVL(!线条色, 0), _
                        zlCommFun.NVL(!线型, 0), zlCommFun.NVL(!线宽, 1)
                End With
                rsTmp.MoveNext
            Loop
        End If
    Else
        Set objMapItems = Flags
    End If
    
    On Error GoTo 0
    
    Call ShowCaseMap(objCaseMap, objMapItems, objDraw, x, y, W, H)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
'保存病人标记图
Public Sub SaveFlag(ByVal ContentID As Long, Flags As Variant, DataConn As Connection)
    Dim i As Long, iMin As Long, iMax As Long
    Dim strSQL As String
    Dim objMapItem As MapItem
    
    If IsEmpty(Flags) Then Exit Sub
    
    If TypeName(Flags) = "MapItems" Then
        For i = 1 To Flags.Count
            Set objMapItem = Flags(i)
            '"类型,'内容','字体','点集',X1,Y1,X2,Y2,填充色,填充方式,线条色,线型,线宽"
            With objMapItem
                strSQL = .类型 & ",'" & .内容 & "','" & .字体 & "','" & .点集 & "'," & _
                    .X1 & "," & .Y1 & "," & .X2 & "," & .Y2 & "," & .填充色 & "," & .填充方式 & "," & _
                    .线条色 & "," & .线型 & "," & .线宽
            End With
            DataConn.Execute "ZL_病人病历标记图_SAVE(" & ContentID & "," & strSQL & ")", , adCmdStoredProc
        Next
    Else
        If UBound(Flags) = -1 Then Exit Sub
    
        iMin = LBound(Flags): iMax = UBound(Flags)
        For i = iMin To iMax
            DataConn.Execute "ZL_病人病历标记图_SAVE(" & ContentID & "," & Flags(i) & ")", , adCmdStoredProc
        Next
    End If
End Sub

Public Function GetMap(ByVal lng病历ID As Long, ByVal picDraw As PictureBox, Optional blnMoved As Boolean = False) As StdPicture
    Dim rsTmp As New ADODB.Recordset
    Dim objFlags As MapItems
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select a.元素编码,b.ID From 病人病历内容 a,病历元素目录 b Where a.ID=[1] And a.元素编码=b.编码"
    If blnMoved Then
        strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "mdlCISCore", lng病历ID)
    If Not rsTmp.EOF Then
        Set objFlags = GetMapItems(lng病历ID, blnMoved)
        With picDraw
            .AutoRedraw = True: .ScaleMode = vbPixels: .Cls: .BackColor = RGB(255, 255, 255)
            Set .Picture = ReadCaseMap(rsTmp(1))
            .Width = .ScaleX(.Picture.Width, vbHimetric, vbTwips): .Height = .ScaleY(IIf(.Picture.Height = 0, 1, .Picture.Height), vbHimetric, vbTwips)
            .Width = IIf(.Width > 10000, 10000, .Width): .Height = .Height * .Width / .ScaleX(IIf(.Picture.Width = 0, 1, .Picture.Width), vbHimetric, vbTwips)
            .Cls: Set .Picture = Nothing
        End With
        ShowFlagInOjbect picDraw, CStr(rsTmp(0)), objFlags, blnMoved:=blnMoved
        Set GetMap = picDraw.Image
    Else
        Set GetMap = New StdPicture
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetMapItems(lngItemID As Long, Optional blnMoved As Boolean = False) As MapItems
'功能：获取标记对象
'参数：lngItemID："病人病历内容"中标记图元素对应的ID；
'返回：Mapitems
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    
    Set GetMapItems = New MapItems
    
    On Error GoTo DBError
    strSQL = "Select * From 病人病历标记图 Where 病历ID=[1]"
    If blnMoved Then
        strSQL = Replace(strSQL, "病人病历标记图", "H病人病历标记图")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", lngItemID)
    Do While Not rsTmp.EOF
        With rsTmp
            GetMapItems.Add !类型, zlCommFun.NVL(!内容), _
                IIf(IsNull(!字体), IIf(!类型 = 0, "宋体,9,0,0000", ""), !字体), _
                zlCommFun.NVL(!点集), zlCommFun.NVL(!X1, 0), _
                zlCommFun.NVL(!Y1, 0), zlCommFun.NVL(!X2, 0), _
                zlCommFun.NVL(!Y2, 0), zlCommFun.NVL(!填充色, &HFFFFFF), _
                zlCommFun.NVL(!填充方式, -1), zlCommFun.NVL(!线条色, 0), _
                zlCommFun.NVL(!线型, 0), zlCommFun.NVL(!线宽, 1)
        End With
        rsTmp.MoveNext
    Loop
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

'Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
''检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
'    If InStr(strInput, "'") > 0 Or InStr(strInput, ";") > 0 Or InStr(strInput, ",") > 0 Or InStr(strInput, "`") > 0 Or InStr(strInput, """") > 0 Then
'        MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
'        Exit Function
'    End If
'    If intMax > 0 Then
'        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
'            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, gstrSysName
'            Exit Function
'        End If
'    End If
'
'    StrIsValid = True
'End Function

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
    '检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If InStr(strInput, "'") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字符！", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Public Function Check是否包含(strSource As String, strTarge As String) As Boolean
    '检查strSource中的每一个字符是否在strTarge中
    Dim i As Long
    Check是否包含 = False
    
    Select Case strTarge
    Case "整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "正整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "正小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    Check是否包含 = True
End Function

Public Sub SelectRow(mshObject As Object, Optional ByVal BackColor As Long = &H8000000D, Optional ByVal ForeColor As Long = &H8000000E)
    Dim i As Integer
    Dim blnPre As Boolean
    Dim intRow As Integer
    Dim intCol As Integer
    
    With mshObject
        blnPre = .Redraw
        intRow = .Row
        intCol = .Col
        .Redraw = False
        
        For i = 0 To .Cols - 1
            .Col = i
            .CellBackColor = BackColor
            .CellForeColor = ForeColor
        Next
        
        .Row = intRow
        .Col = intCol
        .Redraw = blnPre
    End With
End Sub

Public Sub UnSelectRow(mshObject As Object, Optional lngColorSave As Long = 0)
    Dim i As Integer
    Dim blnPre As Boolean
    
    With mshObject
        blnPre = .Redraw
        .Redraw = False
        
        For i = 0 To .Cols - 1
            .Col = i
            .CellBackColor = .BackColor
            .CellForeColor = lngColorSave
        Next
        .Redraw = blnPre
    End With
End Sub
'替换所见项和特殊元素
Public Function GetSpecValue(ItemName As String, sPatientID As String, sPageID As String, iPatientType As Integer) As String
    'sPatientID：病人ID
    'sPageID：主页ID或挂号单号
    'iPatientType：0=门诊、1=住院
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTemp As String
    
    If Len(Trim(sPatientID)) = 0 Then GetSpecValue = "": Exit Function
    strSQL = ""
    Err = 0: On Error GoTo DBError
    
    Select Case ItemName
    Case "书写签名", "-1", "SXR"
        GetSpecValue = UserInfo.姓名: Exit Function
    Case "当前日期", "-2", "DQRQ"
        strSQL = "Select to_Char(SysDate,'YYYY-MM-DD') From Dual"
    Case "当前时间", "-3", "DQSJ"
        strSQL = "Select to_Char(SysDate,'YYYY-MM-DD HH24:MI:SS') From Dual"
    Case "姓名", "XM"
        strSQL = "Select 姓名 From 病人信息 Where 病人ID=" & sPatientID
    Case "性别", "XB"
        strSQL = "Select 性别 From 病人信息 Where 病人ID=" & sPatientID
    Case "年龄", "NL"
        strSQL = "Select 年龄 From 病人信息 Where 病人ID=" & sPatientID
    Case "职业", "ZY"
        strSQL = "Select 职业 From 病人信息 Where 病人ID=" & sPatientID
    Case "民族", "MZ"
        strSQL = "Select 民族 From 病人信息 Where 病人ID=" & sPatientID
    Case "国籍", "GJ"
        strSQL = "Select 国籍 From 病人信息 Where 病人ID=" & sPatientID
    Case "婚姻状况", "HYZK"
        strSQL = "Select 婚姻状况 From 病人信息 Where 病人ID=" & sPatientID
    Case "出生日期", "CSRQ"
        strSQL = "Select to_char(出生日期,'YYYY-MM-DD') From 病人信息 Where 病人ID=" & sPatientID
    Case "出生地点", "CSDD"
        strSQL = "Select 出生地点 From 病人信息 Where 病人ID=" & sPatientID
    Case "身份证号", "SFZH"
        strSQL = "Select 身份证号 From 病人信息 Where 病人ID=" & sPatientID
    Case "身份", "SF"
        strSQL = "Select 身份 From 病人信息 Where 病人ID=" & sPatientID
    Case "学历", "XL"
        strSQL = "Select 学历 From 病人信息 Where 病人ID=" & sPatientID
    Case "家庭地址", "JTDZ"
        strSQL = "Select 家庭地址 From 病人信息 Where 病人ID=" & sPatientID
    Case "家庭电话", "JTDH"
        strSQL = "Select 家庭电话 From 病人信息 Where 病人ID=" & sPatientID
    Case "工作单位", "GZDW"
        strSQL = "Select 工作单位 From 病人信息 Where 病人ID=" & sPatientID
    Case "单位电话", "DWDH"
        strSQL = "Select 单位电话 From 病人信息 Where 病人ID=" & sPatientID
    Case "门诊号", "MZH"
        strSQL = "Select 门诊号 From 病人信息 Where 病人ID=" & sPatientID
    Case "就诊卡号", "JZKH"
        strSQL = "Select 就诊卡号 From 病人信息 Where 病人ID=" & sPatientID
    Case "就诊科室", "JZKS"
        strSQL = "Select D.名称" & _
                " From 部门表 D," & _
                "      (Select Distinct 病人科室ID" & _
                "      From 病人费用记录" & _
                "      Where 病人id=" & sPatientID & " And No='" & sPageID & "'" & _
                "            And 记录性质=4 And 记录状态=1 And 收费类别='1') R" & _
                " Where D.Id=R.病人科室ID"
    Case "就诊时间", "JZSJ"
        strSQL = "Select Distinct to_char(发生时间,'YYYY-MM-DD HH24:MI:SS')" & _
                " From 病人费用记录" & _
                " Where 病人id=" & sPatientID & " And No='" & sPageID & "'" & _
                "       And 记录性质=4 And 记录状态=1 And 收费类别='1'"
    Case "是否急诊", "SFJZ"
        strSQL = "Select Distinct nvl(加班标志,0)" & _
                " From 病人费用记录" & _
                " Where 病人id=" & sPatientID & " And No='" & sPageID & "'" & _
                "       And 记录性质=4 And 记录状态=1 And 收费类别='1'"
    Case "住院号", "ZYH"
        strSQL = "Select 住院号 From 病人信息 Where 病人ID=" & sPatientID
    Case "住院次数", "ZYCS"
        strSQL = "Select 住院次数 From 病人信息 Where 病人ID=" & sPatientID
    Case "入院日期", "RYRQ"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select to_char(入院日期,'YYYY-MM-DD')" & _
                " From 病案主页" & _
                " Where 病人id=" & sPatientID & " And 主页id=" & sPageID
        End If
    Case "出院日期", "CYRQ"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select to_char(出院日期,'YYYY-MM-DD')" & _
                " From 病案主页" & _
                " Where 病人id=" & sPatientID & " And 主页id=" & sPageID
        End If
    Case "住院目的", "ZYMD"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select 住院目的" & _
                " From 病案主页" & _
                " Where 病人id=" & sPatientID & " And 主页id=" & sPageID
        End If
    Case "入院科室", "RYKS"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select D.名称" & _
                " From 部门表 D," & _
                "      (Select Distinct 入院科室ID" & _
                "       From 病案主页" & _
                "       Where 病人id=" & sPatientID & " And 主页id=" & sPageID & ") P" & _
                " Where D.Id=P.入院科室ID"
        End If
    Case "入院病区", "RYBQ"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select D.名称" & _
                " From 部门表 D," & _
                "      (Select Distinct 入院病区ID" & _
                "       From 病案主页" & _
                "       Where 病人id=" & sPatientID & " And 主页id=" & sPageID & ") P" & _
                " Where D.Id=P.入院病区ID"
        End If
    Case "当前床号", "DQCH"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select 出院病床" & _
                " From 病案主页" & _
                " Where 病人id=" & sPatientID & " And 主页id=" & sPageID
        End If
    Case "当前病区", "DQBQ"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select D.名称" & _
                " From 部门表 D," & _
                "      (Select Distinct 当前病区ID" & _
                "       From 病案主页" & _
                "       Where 病人id=" & sPatientID & " And 主页id=" & sPageID & ") P" & _
                " Where D.Id=P.当前病区ID"
        End If
    Case "当前科室", "DQKS"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select D.名称" & _
                " From 部门表 D," & _
                "      (Select Distinct 出院科室ID" & _
                "       From 病案主页" & _
                "       Where 病人id=" & sPatientID & " And 主页id=" & sPageID & ") P" & _
                " Where D.Id=P.出院科室ID"
        End If
    Case "当前病况", "DQBK"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select 当前病况" & _
                " From 病案主页" & _
                " Where 病人id=" & sPatientID & " And 主页id=" & sPageID
        End If
    Case "住院医师", "ZYYS"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select 住院医师" & _
                " From 病案主页" & _
                " Where 病人id=" & sPatientID & " And 主页id=" & sPageID
        End If
    Case "责任护士", "ZRHS"
        If iPatientType = 0 Then
            strSQL = ""
        Else
            strSQL = "Select 责任护士" & _
                " From 病案主页" & _
                " Where 病人id=" & sPatientID & " And 主页id=" & sPageID
        End If
    Case "最后诊断", "ZHZD"
        strSQL = "Select 诊断描述" & _
                " From 病人诊断记录" & _
                " Where 病人id=" & sPatientID & " And 取消时间 Is Null" & _
                "       And 记录日期 In (" & _
                "           Select Max(记录日期)" & _
                "           From 病人诊断记录" & _
                " Where 病人id=" & sPatientID & " And 取消时间 Is Null)"
        With rsTmp
            If .State = adStateOpen Then .Close
            Call SQLTest(App.ProductName, "读取替换项", strSQL): rsTmp.Open strSQL, gcnOracle, adOpenKeyset: Call SQLTest
            strTemp = ""
            Do While Not .EOF
                strTemp = strTemp & vbCrLf & .AbsolutePosition & "." & .Fields(0).Value
                .MoveNext
            Loop
            If strTemp <> "" Then strTemp = Mid(strTemp, 3)
            GetSpecValue = strTemp: Exit Function
        End With
    Case "过敏药物", "GMYW"
        strSQL = "Select 药物名" & _
                " From 病人过敏记录" & _
                " Where 病人id=" & sPatientID & " And 结果=1" & _
                "       And 记录时间 In (" & _
                "           Select Max(记录时间)" & _
                "           From 病人过敏记录" & _
                " Where 病人id=" & sPatientID & " And 结果=1)"
        With rsTmp
            If .State = adStateOpen Then .Close
            Call SQLTest(App.ProductName, "读取替换项", strSQL): rsTmp.Open strSQL, gcnOracle, adOpenKeyset: Call SQLTest
            strTemp = ""
            Do While Not .EOF
                strTemp = strTemp & vbCrLf & .AbsolutePosition & "." & .Fields(0).Value
                .MoveNext
            Loop
            If strTemp <> "" Then strTemp = Mid(strTemp, 3)
            GetSpecValue = strTemp: Exit Function
        End With
    End Select
    
    If strSQL = "" Then GetSpecValue = "": Exit Function
    With rsTmp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "读取替换项", strSQL): rsTmp.Open strSQL, gcnOracle, adOpenKeyset: Call SQLTest
        If .EOF Or .BOF Then
            GetSpecValue = ""
        Else
            GetSpecValue = IIf(IsNull(.Fields(0).Value), "", .Fields(0).Value)
        End If
    End With
    Exit Function

DBError:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function

'替换病历字符串中的变量
Public Function ReplaceString(strSource As String, sPatientID As String, sPageID As String, iPatientType As Integer, _
    Optional strVariableBegin As String = "{{", Optional strVariableEnd As String = "}}") As String
    Dim iLen1 As Integer, iLen2 As Integer
    Dim iStrPoint As Long, iStrLength As Long
    Dim iVariableBeginPos As Long, iVariableEndPos As Long
    Dim strVariable As String, strReturn As String
    iLen1 = Len(strVariableBegin): iLen2 = Len(strVariableEnd)
    
    ReplaceString = strSource
    iStrPoint = 1: iStrLength = Len(ReplaceString)
    Do While iStrPoint <= iStrLength
        If iVariableBeginPos > 0 Then
            iVariableEndPos = InStr(iVariableBeginPos + iLen1, ReplaceString, strVariableEnd)
            If iVariableEndPos = 0 Then
                Exit Do
            Else
                strVariable = Mid(ReplaceString, _
                    iVariableBeginPos + iLen1, iVariableEndPos - (iVariableBeginPos + iLen1))
                strReturn = GetSpecValue(strVariable, sPatientID, sPageID, iPatientType)
                
                ReplaceString = Replace(ReplaceString, strVariableBegin + strVariable + strVariableEnd, strReturn)
                iStrPoint = iVariableBeginPos + Len(strReturn)
                iStrLength = Len(ReplaceString)
                
                iVariableBeginPos = 0
            End If
        Else
            iVariableBeginPos = InStr(iStrPoint, ReplaceString, strVariableBegin)
            If iVariableBeginPos = 0 Then Exit Do
        End If
    Loop
End Function

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As POINTAPI
    vPoint.x = lngX / Screen.TwipsPerPixelX: vPoint.y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.x = vPoint.x * Screen.TwipsPerPixelX: vPoint.y = vPoint.y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Function GetFileId(ByVal lngWritId As Long) As Long
'功能：获取用户指定文件的病人病历记录，以便用于引入等操作
'参数：lngWritId-需要查找的病历文件
    With frmWritImp
        .lblWrit.Tag = lngWritId
        .Show 1
        GetFileId = .lngFileId
        Unload frmWritImp
    End With
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function IncStr(ByVal strVal As String) As String
'功能：对一个字符串自动加1。
'说明：每一位进位时,如果是数字,则按十进制处理,否则按26进制处理
'可提取为公共函数
    Dim i As Integer, strTmp As String, bytUp As Byte, bytAdd As Byte
    
    For i = Len(strVal) To 1 Step -1
        If i = Len(strVal) Then
            bytAdd = 1
        Else
            bytAdd = 0
        End If
        If IsNumeric(Mid(strVal, i, 1)) Then
            If CByte(Mid(strVal, i, 1)) + bytAdd + bytUp < 10 Then
                strVal = Left(strVal, i - 1) & CByte(Mid(strVal, i, 1)) + bytAdd + bytUp & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        Else
            If Asc(Mid(strVal, i, 1)) + bytAdd + bytUp <= Asc("Z") Then
                strVal = Left(strVal, i - 1) & Chr(Asc(Mid(strVal, i, 1)) + bytAdd + bytUp) & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        End If
        If bytUp = 0 Then Exit For
    Next
    IncStr = strVal
End Function

Public Function MatchIndex(ByVal lngHwnd As Long, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'功能：根据输入的字符串自动匹配ComboBox的选中项,并自动识别输入间隔
'参数：lngHwnd=ComboBox的Hwnd属性,KeyAscii=ComboBox的KeyPress事件中的KeyAscii参数,sngInterval=指定输入间隔
'返回：-2=未加处理,其它=匹配的索引(含不匹配的索引)
'说明：请将该函数在KeyPress事件中调用。

    Static lngPreTime As Single, lngPreHwnd As Long
    Static strFind As String
    Dim sngTime As Single, lngR As Long
    
    If lngPreHwnd <> lngHwnd Then lngPreTime = Empty: strFind = Empty
    lngPreHwnd = lngHwnd
    
    If KeyAscii <> 13 Then
        sngTime = Timer
        If Abs(sngTime - lngPreTime) > sngInterval Then '输入间隔(缺省为0.5秒)
            strFind = ""
        End If
        strFind = strFind & Chr(KeyAscii)
        lngPreTime = Timer
        KeyAscii = 0 '使ComboBox本身的单字匹配功能失效
        MatchIndex = SendMessage(lngHwnd, CB_FINDSTRING, -1, ByVal strFind)
        If MatchIndex = -1 Then Beep
    Else
        MatchIndex = -2 '在这里对回车不作处理
    End If
End Function

Public Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function

Public Sub CheckLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub

Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
'功能：去掉TextBox的默认右键菜单
    If msg <> WM_CONTEXTMENU Then
        WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
    End If
End Function

Public Sub RaisEffect(picBox As PictureBox, Optional intStyle As Integer, Optional strName As String = "")
'功能：将PictureBox模拟成3D平面按钮
'参数：intStyle:0=平面,-1=凹下,1=凸起
    
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            DrawEdge .hDC, PicRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            picBox.Print strName
        End If
    End With
End Sub

Public Sub PrintDiagReport(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, objParent As Object, Optional ByVal PrtMode As Integer = 2, Optional ObjPic As Object = Nothing, Optional blnMoved As Boolean = False)
'打印辅诊报告
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsImages As ADODB.Recordset
    Dim strRptName As String
    Dim aImages(1, 8) As Variant, aFlagImages(1, 8) As Variant, i As Integer
    Dim strTempPath As String, lngBuffSize As Long
    Dim intReportFormatItem As Integer
    
    Dim objFileSystem As New Scripting.FileSystemObject, strTmpFile As String
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim strDeviceNO1 As String
    Dim strDeviceNO2 As String
    Dim strNO As String, lng记录性质 As Long
    Dim iTmpFileCount As Integer, iFlagCount As Integer
    Dim objImages As New DicomImages, intRows As Integer, intCols As Integer, objAssembleImage As New DicomImage
    
    On Error GoTo DBError
    strTempPath = Space(255)
    lngBuffSize = GetTempPath(Len(strTempPath), strTempPath)
    strTempPath = Mid(strTempPath, 1, lngBuffSize)
    strTempPath = Mid(strTempPath, 1, InStrRev(strTempPath, "\"))
    
    strSQL = "Select A.NO,A.记录性质,'ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2'" & _
        " From 病人医嘱发送 A,病人病历记录 B,病历文件目录 C" & _
        " Where A.报告ID=B.ID And B.文件ID=C.ID And A.医嘱ID=[1] And A.发送号=[2]"
    If blnMoved Then
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人病历记录", "H病人病历记录")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "检查报告", lng医嘱ID, lng发送号)
    If rsTmp.EOF Then
        MsgBox "该项检查未填写报告，不能打印！", vbInformation, gstrSysName
    Else
        strRptName = rsTmp(2): strNO = rsTmp(0): lng记录性质 = rsTmp(1)
        If PrtMode = 2 Then
            If Not ReportPrintSet(gcnOracle, glngSys, strRptName, objParent) Then Exit Sub
            intReportFormatItem = GetSetting("ZLSOFT", "私有模块\ZLHIS\zl9Report\LocalSet\" & strRptName, "Format", 1)
        Else
            intReportFormatItem = GetSetting("ZLSOFT", "私有模块\ZLHIS\zl9Report\frmReport" & strRptName, "格式", 1)
        End If
        'PACS的影像图片
        strSQL = "Select A.用户名1,A.密码1,A.Host1,A.Root1,A.URL1,A.用户名2,A.密码2,A.Host2,A.Root2,A.URL2," & _
            "a.设备号1,a.设备号2,A.NO,A.记录性质 From" & _
            " (Select E.IP地址 As Host1,'/'||E.Ftp目录||'/' as Root1,e.设备号 as 设备号1," & _
            "Decode(D.接收日期,Null,'',to_Char(D.接收日期,'YYYYMMDD')||'/')" & _
            "||D.检查UID||'/'||A.图象文件 As URL1," & _
            "F.IP地址 As Host2,'/'||f.Ftp目录||'/' as Root2," & _
            "Decode(D.接收日期,Null,'',to_Char(D.接收日期,'YYYYMMDD')||'/')" & _
            "||D.检查UID||'/'||A.图象文件 As URL2,f.设备号 as 设备号2," & _
            "C.NO,C.记录性质,E.用户名 as 用户名1,E.密码 as 密码1,F.用户名 as 用户名2,F.密码 as 密码2, Rownum As Seq " & _
            " From 病人病历外部图 A,病人病历内容 B,病人医嘱发送 C,影像检查记录 D,影像设备目录 E,影像设备目录 F" & _
            " Where A.病历ID=B.ID And B.病历记录ID=C.报告ID And C.医嘱ID=D.医嘱ID" & _
            " And C.发送号=D.发送号 And D.位置一=E.设备号(+) and d.位置二=F.设备号(+)" & _
            " And C.医嘱ID=[1] And C.发送号=[2]" & _
            " Order By A.序号) A"
        If blnMoved Then
            strSQL = Replace(strSQL, "病人病历外部图", "H病人病历外部图")
            strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
            strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
            strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        End If
        Set rsTmp = OpenSQLRecord(strSQL, "检查报告", lng医嘱ID, lng发送号, intReportFormatItem)
        iTmpFileCount = rsTmp.RecordCount
        strSQL = "Select A.编号,B.名称,B.W,B.H" & _
            " From zlReports A,zlRPTItems B,病历文件目录 C,病人医嘱记录 D,诊疗单据应用 E" & _
            " Where A.ID=B.报表ID And A.编号='ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2'" & _
            " And C.ID=E.病历文件ID And D.诊疗项目ID=E.诊疗项目ID And Nvl(B.下线,0)=1 And B.类型=11" & _
            " And E.应用场合=D.病人来源 And D.ID=[1]" & _
            " And B.名称 Not Like '标记%' and b.格式号=[3]" & _
            " Order BY b.名称"
        If blnMoved Then
            strSQL = Replace(strSQL, "病人病历外部图", "H病人病历外部图")
            strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
            strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
            strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        End If
        Set rsImages = OpenSQLRecord(strSQL, "检查报告", lng医嘱ID, lng发送号, intReportFormatItem)
        If rsImages.RecordCount = 1 Then
            '图像排版
            For i = 0 To rsTmp.RecordCount - 1
                If i > 8 Then Exit For
                
                strTmpFile = App.Path & IIf(Len(App.Path) > 3, "\", "") & "TmpImage\" & objFileSystem.GetParentFolderName(rsTmp("URL1"))
                strTmpFile = Replace(strTmpFile, "/", "\")
                MkLocalDir strTmpFile
                strTmpFile = strTmpFile & "\" & objFileSystem.GetFileName(rsTmp("URL1"))
                
                If Dir(strTmpFile, vbDirectory) = "" Then
                    If strDeviceNO1 <> rsTmp("设备号1") Then
                        strDeviceNO1 = rsTmp("设备号1")
                        Inet1.FuncFtpConnect NVL(rsTmp("Host1")), NVL(rsTmp("用户名1")), NVL(rsTmp("密码1"))
                    End If
                    
                    If strDeviceNO2 <> rsTmp("设备号2") Then
                        strDeviceNO2 = rsTmp("设备号2")
                        Inet2.FuncFtpConnect NVL(rsTmp("Host2")), NVL(rsTmp("用户名2")), NVL(rsTmp("密码2"))
                    End If
                    
                    If Inet1.FuncDownloadFile(objFileSystem.GetParentFolderName(NVL(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL1"))) <> 0 Then
                        Call Inet2.FuncDownloadFile(objFileSystem.GetParentFolderName(NVL(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL2")))
                    End If
                End If
'                objAssembleImage.FileImport strTmpFile, "JPEG"
'                objImages.Add objAssembleImage
                
                objImages.AddNew
                objImages(objImages.Count).FileImport strTmpFile, "JPEG"
                
                rsTmp.MoveNext
            Next
            If objImages.Count > 0 Then
                ResizeRegion i, rsImages("W"), rsImages("H"), intRows, intCols
                Set objAssembleImage = funAssembleImage(objImages, intRows, intCols, rsImages("H"), rsImages("W"))
                strTmpFile = objFileSystem.GetParentFolderName(strTmpFile) & "\" & objFileSystem.GetTempName
                objAssembleImage.FileExport strTmpFile, "JPEG"
                    
                aImages(0, 0) = rsImages("名称")
                aImages(1, 0) = strTmpFile
            End If
            For i = 1 To 8
                aImages(0, i) = "1"
                aImages(1, i) = "1"
            Next
        Else
            For i = 0 To rsTmp.RecordCount - 1
                If i > 8 Then Exit For
                If rsImages.EOF Then Exit For
                
    '            strTmpFile = strTempPath & objFileSystem.GetFileName(rsTmp(3))
                
                strTmpFile = App.Path & IIf(Len(App.Path) > 3, "\", "") & "TmpImage\" & objFileSystem.GetParentFolderName(rsTmp("URL1"))
                strTmpFile = Replace(strTmpFile, "/", "\")
                MkLocalDir strTmpFile
                strTmpFile = strTmpFile & "\" & objFileSystem.GetFileName(rsTmp("URL1"))
                
                If Dir(strTmpFile, vbDirectory) = "" Then
                    If strDeviceNO1 <> rsTmp("设备号1") Then
                        strDeviceNO1 = rsTmp("设备号1")
                        Inet1.FuncFtpConnect NVL(rsTmp("Host1")), NVL(rsTmp("用户名1")), NVL(rsTmp("密码1"))
                    End If
                    
                    If strDeviceNO2 <> rsTmp("设备号2") Then
                        strDeviceNO2 = rsTmp("设备号2")
                        Inet2.FuncFtpConnect NVL(rsTmp("Host2")), NVL(rsTmp("用户名2")), NVL(rsTmp("密码2"))
                    End If
                    
                    'Inet.strIPAddress = Nvl(rsTmp(2)): Inet.strUser = Nvl(rsTmp(6)): Inet.strPsw = Nvl(rsTmp(7))
                    If Inet1.FuncDownloadFile(objFileSystem.GetParentFolderName(NVL(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL1"))) <> 0 Then
    '                    strTmpFile = strCachePath & Nvl(rsTmp("URL2"))
    '                        Inet.strIPAddress = Nvl(rsTmp("Host2")): Inet.strUser = Nvl(rsTmp("User2")): Inet.strPsw = Nvl(rsTmp("Pwd2"))
                        Call Inet2.FuncDownloadFile(objFileSystem.GetParentFolderName(NVL(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFileSystem.GetFileName(rsTmp("URL2")))
                    End If
                End If
                    
                aImages(0, i) = rsImages("名称")
                aImages(1, i) = strTmpFile
                rsImages.MoveNext
                rsTmp.MoveNext
            Next
            For i = rsTmp.RecordCount To 8
                aImages(0, i) = "1"
                aImages(1, i) = "1"
            Next
        End If
        
        If Not ObjPic Is Nothing Then
            '标记图的生成
            strSQL = "Select B.编号,B.名称,A.元素ID,A.内容ID,B.W,B.H From" & _
                " (Select B.ID As 元素ID,A.ID 内容ID,Rownum As Seq From 病人病历内容 A,病历元素目录 B,病人医嘱发送 C" & _
                " Where C.报告ID=A.病历记录ID AND A.元素编码=B.编码 And" & _
                " C.医嘱ID=[1] And C.发送号=[2] And A.元素类型=3) A," & _
                " (Select A.编号,B.名称,B.W,B.H,Rownum As Seq" & _
                " From zlReports A,zlRPTItems B,病历文件目录 C,病人医嘱记录 D,诊疗单据应用 E" & _
                " Where A.ID=B.报表ID And A.编号='ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2'" & _
                " And C.ID=E.病历文件ID And D.诊疗项目ID=E.诊疗项目ID And Nvl(B.下线,0)=1 And B.类型=11" & _
                " And E.应用场合=D.病人来源 And D.ID=[1]" & _
                " And B.名称 Like '标记%'" & _
                " Order BY Trunc(Y/567),Trunc(X/567)) B Where A.Seq=B.Seq"
            If rsTmp.State <> adStateClosed Then rsTmp.Close
            If blnMoved Then
                strSQL = Replace(strSQL, "病人病历内容", "H病人病历内容")
                strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            End If
            Set rsTmp = OpenSQLRecord(strSQL, "检查报告", lng医嘱ID, lng发送号)
            iFlagCount = rsTmp.RecordCount
            ObjPic.Cls
            For i = 0 To rsTmp.RecordCount - 1
                If i > 8 Then Exit For
                
                strTmpFile = strTempPath & objFileSystem.GetTempName
                
                '计算容器尺寸
                On Error Resume Next
                Set ObjPic.Picture = ReadCaseMap(rsTmp(2))
                ObjPic.Width = ObjPic.ScaleX(ObjPic.Picture.Width, vbHimetric, vbTwips): ObjPic.Height = ObjPic.ScaleY(ObjPic.Picture.Height, vbHimetric, vbTwips)
                If ObjPic.Width / ObjPic.Height > rsTmp(4) / rsTmp(5) Then
                    ObjPic.Width = ObjPic.Height * rsTmp(4) / rsTmp(5)
                Else
                    ObjPic.Height = ObjPic.Width / (rsTmp(4) / rsTmp(5))
                End If
                ObjPic.Cls: Set ObjPic.Picture = Nothing
                On Error GoTo DBError
                Call ShowMapInOjbect_1(ObjPic, rsTmp(2), rsTmp(3), blnMoved:=blnMoved)
                SavePicture ObjPic.Image, strTmpFile
                ObjPic.Cls
            
                aFlagImages(0, i) = rsTmp(1)
                aFlagImages(1, i) = strTmpFile
                rsTmp.MoveNext
            Next
            For i = rsTmp.RecordCount To 8
                aFlagImages(0, i) = "1"
                aFlagImages(1, i) = "1"
            Next
            Call ReportOpen(gcnOracle, glngSys, strRptName, objParent, _
                "NO=" & strNO, _
                "性质=" & lng记录性质, "医嘱ID=" & lng医嘱ID, _
                aImages(0, 0) & "=" & aImages(1, 0), _
                aImages(0, 1) & "=" & aImages(1, 1), _
                aImages(0, 2) & "=" & aImages(1, 2), _
                aImages(0, 3) & "=" & aImages(1, 3), _
                aImages(0, 4) & "=" & aImages(1, 4), _
                aImages(0, 5) & "=" & aImages(1, 5), _
                aImages(0, 6) & "=" & aImages(1, 6), _
                aImages(0, 7) & "=" & aImages(1, 7), _
                aImages(0, 8) & "=" & aImages(1, 8), _
                aFlagImages(0, 0) & "=" & aFlagImages(1, 0), _
                aFlagImages(0, 1) & "=" & aFlagImages(1, 1), _
                aFlagImages(0, 2) & "=" & aFlagImages(1, 2), _
                aFlagImages(0, 3) & "=" & aFlagImages(1, 3), _
                aFlagImages(0, 4) & "=" & aFlagImages(1, 4), _
                aFlagImages(0, 5) & "=" & aFlagImages(1, 5), _
                aFlagImages(0, 6) & "=" & aFlagImages(1, 6), _
                aFlagImages(0, 7) & "=" & aFlagImages(1, 7), _
                aFlagImages(0, 8) & "=" & aFlagImages(1, 8), PrtMode)
        Else
            Call ReportOpen(gcnOracle, glngSys, strRptName, objParent, _
                "NO=" & strNO, _
                "性质=" & lng记录性质, "医嘱ID=" & lng医嘱ID, _
                aImages(0, 0) & "=" & aImages(1, 0), _
                aImages(0, 1) & "=" & aImages(1, 1), _
                aImages(0, 2) & "=" & aImages(1, 2), _
                aImages(0, 3) & "=" & aImages(1, 3), _
                aImages(0, 4) & "=" & aImages(1, 4), _
                aImages(0, 5) & "=" & aImages(1, 5), _
                aImages(0, 6) & "=" & aImages(1, 6), _
                aImages(0, 7) & "=" & aImages(1, 7), _
                aImages(0, 8) & "=" & aImages(1, 8), PrtMode)
        End If
        '删除临时文件
'        For i = 0 To iTmpFileCount - 1
'            objFileSystem.DeleteFile aImages(1, i), True
'        Next
        For i = 0 To iFlagCount - 1
            If Dir(aFlagImages(1, i), vbDirectory) <> "" Then
                objFileSystem.DeleteFile aFlagImages(1, i), True
            End If
        Next
    End If
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ShowMapInOjbect_1(objDraw As Object, var元素 As Variant, Optional lng病历ID As Long, Optional x As Long, Optional y As Long, Optional W As Long, Optional H As Long, Optional blnMoved As Boolean = False)
'功能：在指定的对象(PictureBox或Form)上显示标记图
'参数：objDraw=PictureBox或窗体对象,它的ScaleMode必须为Pixel
'      var元素=标记图元素的编码(字符型)或ID(数字型)
'      lng病历ID="病人病历内容"中标记图元素对应的ID,如果不传,仅显示标记图背景
'      X,Y,W,H=显示的目标客户端范围,可以不指定,单位为Pixel
'说明：可以利用该函数进行打印输出(因为是API作图,因此不能直接将objDraw指定为打印机,而是在PictureBox上按一定比例输出后,取PictureBox.Image输出到打印机)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim objCaseMap As StdPicture, objMapItems As New MapItems
    
    On Error GoTo errH
        
    '读取标记图元素的内容
    If TypeName(var元素) = "String" Then
        strSQL = "Select * From 病历元素目录 Where 编码=[1]"
    Else
        strSQL = "Select * From 病历元素目录 Where ID=[1]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", CStr(var元素))
    If rsTmp.EOF Then Exit Sub '必须要有图形背景
    
    Set objCaseMap = ReadCaseMap(rsTmp!ID)
    If objCaseMap Is Nothing Then Exit Sub '必须要有图形背景
    
    '读取标记图的标注内容
    If lng病历ID <> 0 Then
        strSQL = "Select * From 病人病历标记图 Where 病历ID=[1]"
        If blnMoved Then
            strSQL = Replace(strSQL, "病人病历标记图", "H病人病历标记图")
        End If
        Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", lng病历ID)
        Do While Not rsTmp.EOF
            With rsTmp
                objMapItems.Add !类型, zlCommFun.NVL(!内容), _
                    IIf(IsNull(!字体), IIf(!类型 = 0, "宋体,9,0,0000", ""), !字体), _
                    zlCommFun.NVL(!点集), zlCommFun.NVL(!X1, 0), _
                    zlCommFun.NVL(!Y1, 0), zlCommFun.NVL(!X2, 0), _
                    zlCommFun.NVL(!Y2, 0), zlCommFun.NVL(!填充色, &HFFFFFF), _
                    zlCommFun.NVL(!填充方式, -1), zlCommFun.NVL(!线条色, 0), _
                    zlCommFun.NVL(!线型, 0), zlCommFun.NVL(!线宽, 1)
            End With
            rsTmp.MoveNext
        Loop
    End If
    
    On Error GoTo 0
    
    Call ShowCaseMap(objCaseMap, objMapItems, objDraw, x, y, W, H)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function CheckOneDuty(ByVal str医嘱 As String, ByVal str职务 As String, ByVal str医生 As String, ByVal bln医保 As Boolean) As String
'功能：检查当前指定药品处方职务是否符合
'参数：str医嘱=药品医嘱提示内容
'      str职务=药品处方职务
'      str医生=开嘱医生
'      bln医保=是否公费或医保病人
'      grsDuty=记录医生职务缓存
'返回：职务不满足的提示信息，如果满足则返回空。
    Const STR_职务 = "正高,副高,中级,助理/师级,员/士,,,,待聘"
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    Dim int职务A As Integer, int职务B As Integer
    
    If Len(str职务) <> 2 Or str医生 = "" Then Exit Function
    
    '取药品处方职务
    If bln医保 Then
        int职务B = Val(Right(str职务, 1))
    Else
        int职务B = Val(Left(str职务, 1))
    End If
    If int职务B = 0 Then Exit Function '不限制
    
    '取医生职务
    If grsDuty Is Nothing Then
        Set grsDuty = New ADODB.Recordset
        grsDuty.Fields.Append "医生", adVarChar, 50
        grsDuty.Fields.Append "职务", adInteger
        grsDuty.CursorLocation = adUseClient
        grsDuty.LockType = adLockOptimistic
        grsDuty.CursorType = adOpenStatic
        grsDuty.Open
    End If
    grsDuty.Filter = "医生='" & str医生 & "'"
    If grsDuty.EOF Then
        On Error GoTo errH
        strSQL = "Select 姓名,Nvl(聘任技术职务,0) as 职务 From 人员表 Where 姓名='" & str医生 & "'"
        Set rsTmp = New ADODB.Recordset
        Call OpenRecord(rsTmp, strSQL, "mdlCISWork")
        On Error GoTo 0
        If Not rsTmp.EOF Then
            grsDuty.AddNew
            grsDuty!医生 = rsTmp!姓名
            grsDuty!职务 = rsTmp!职务
            grsDuty.Update
        End If
    End If
    If Not grsDuty.EOF Then
        int职务A = grsDuty!职务
    End If
        
    '检查职务要求
    If int职务A = 0 Then
        '医生未设置职务的情况
        strMsg = """" & str医嘱 & """要求的处方职务不满足：" & vbCrLf & vbCrLf & IIf(bln医保, "对医保或公费病人,", "") & _
            "该药品要求职务至少为""" & Split(STR_职务, ",")(int职务B - 1) & """才能下达,而医生""" & str医生 & """未设置职务。"
    ElseIf int职务B < int职务A Then
        '数值越小职务越高
        strMsg = """" & str医嘱 & """要求的处方职务不满足：" & vbCrLf & vbCrLf & IIf(bln医保, "对医保或公费病人,", "") & _
            "该药品要求职务至少为""" & Split(STR_职务, ",")(int职务B - 1) & """才能下达,而医生""" & str医生 & """的职务为""" & Split(STR_职务, ",")(int职务A - 1) & """。"
    End If
    CheckOneDuty = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub MkLocalDir(ByVal strDir As String)
'功能：创建本地目录
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Function GetSysParVal(Optional ByVal int参数号 As Integer = -9999, Optional ByVal strDefault As String) As String
'功能：获取指定系统参数的值
'参数：int参数号=为-9999时，初始化参数集
'      strDefault=如果没有值或为空的缺省值
    Dim blnDo As Boolean, strSQL As String
    
    On Error GoTo errH
    
    blnDo = True
    If int参数号 <> -9999 Then
        If Not grsSysPars Is Nothing Then
            If grsSysPars.State = 1 Then blnDo = False
        End If
    End If
    If blnDo Then
        strSQL = "Select 参数号,参数名,参数值 From 系统参数表"
        Set grsSysPars = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(grsSysPars, strSQL, "GetSysParVal")
    End If
    
    If int参数号 <> -9999 Then
        grsSysPars.Filter = "参数号=" & int参数号
        If Not grsSysPars.EOF Then
            GetSysParVal = NVL(grsSysPars!参数值, strDefault)
        Else
            GetSysParVal = strDefault
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function funAssembleImage(AssembleViewer As DicomImages, ByVal intRows As Integer, ByVal intCols As Integer, _
    ByVal lngHeight As Long, ByVal lngWidth As Long) As DicomImage

'组合viewer中的显示的所有图像成一个图像

    Dim Image As New DicomImage '新图像
    Dim imgs As New DicomImages '临时存储屏幕采集的图像集
    Dim intWidth As Integer     '新图像的宽度
    Dim intHeight As Integer    '新图像的高度
    Dim Simg As New DicomImage
    Dim intLeft As Integer
    Dim intRight As Integer
    Dim intTop As Integer
    Dim intBottom As Integer
    Dim sZoom As Single
    Dim intImgRectWidth As Integer  '单张图像可占用的区域宽度
    Dim intImgRectHeight As Integer '单张图像可占用的区域高度
    Dim i As Integer
    Dim intMaxWidth As Integer      '拼接后图像的最大宽度
    Dim intMaxHeight As Integer     '拼接后图像的最大高度
    Dim intBorder As Integer        '图像之间的边距
    Dim intImgX As Integer          'X方向的图像数量
    Dim intImgY As Integer          'Y方向的图像数量
    Dim intActualSizex As Integer   '图像旋转变换后X方向的像素点数
    Dim intActualSizey As Integer   '图像旋转变换后Y方向的像素点数
    Dim intOffsetX As Integer       '拼接时X方向的位移
    Dim intOffsetY As Integer       '拼接时Y方向的位移
    Dim dlImgLabel As DicomLabel    '图像的标注
    Dim lngWhiteX As Long           '将图象底色改成白色的X宽度
    Dim lngWhiteY As Long           '将图象底色改成白色的Y高度
    
    If AssembleViewer.Count <= 0 Then
        '返回一个黑图**************
        Exit Function
    End If

    '计算新图像的宽度和高度

    '新图像的宽度和高度不能够大于intMaxWidth×intMaxHeight（宽度×高度）
    intMaxWidth = 3073
    intMaxHeight = 3073
    intBorder = 10

    intImgRectWidth = 0
    intImgRectHeight = 0

    '估算新图像的宽度和高度

    '使用原图像的宽度和高度和，并用Viewer的比例来修正。

    '估算图像的新宽高

    For i = 1 To AssembleViewer.Count
        sZoom = (lngWidth / intCols) / (AssembleViewer(i).SizeX * Screen.TwipsPerPixelX)
        If sZoom > (lngHeight / intRows) / (AssembleViewer(i).SizeY * Screen.TwipsPerPixelY) Then
            sZoom = (lngHeight / intRows) / (AssembleViewer(i).SizeY * Screen.TwipsPerPixelY)
        End If
        AssembleViewer(i).Zoom = sZoom
        '采集图像
        Set Simg = AssembleViewer(i).PrinterImage(8, 3, True, sZoom, 0, AssembleViewer(i).SizeX, 0, AssembleViewer(i).SizeY)
        imgs.Add Simg
    Next i

    '精确计算新图像的宽度和高度
    intImgRectWidth = 0
    intImgRectHeight = 0

    For i = 1 To imgs.Count
        If intImgRectWidth < imgs(i).SizeX Then intImgRectWidth = imgs(i).SizeX
        If intImgRectHeight < imgs(i).SizeY Then intImgRectHeight = imgs(i).SizeY
        imgs(i).Attributes.Add &H8, &H16, "doSOP_SecondaryCapture"
    Next i
    intImgRectWidth = intImgRectWidth + intBorder
    intImgRectHeight = intImgRectHeight + intBorder
    intWidth = intImgRectWidth * intCols
    intHeight = intImgRectHeight * intRows

    '创建新图像
    Image.Name = "print"
    Image.PatientID = "print001"
    
    Image.Attributes.Add &H8, &H16, doSOP_SecondaryCapture
    Image.Attributes.Add &H28, &H2, 3 ' samples/pixel
    Image.Attributes.Add &H28, &H4, "RGB" ' photometric interpreation  'CT都是MONOCHROME2,CR都是MONOCHROME1？
    Image.Attributes.Add &H28, &H10, intHeight  'x,Rows
    Image.Attributes.Add &H28, &H11, intWidth 'Y,Columns
    Image.Attributes.Add &H28, &H100, 8 'bits allocated
    Image.Attributes.Add &H28, &H101, 8 ' bits stored
    Image.Attributes.Add &H28, &H102, 7 ' high bit
    ReDim pix(intWidth * 3, intHeight * 3) As Byte
    For lngWhiteX = 0 To intWidth * 3
        For lngWhiteY = 0 To intHeight * 3
            pix(lngWhiteX, lngWhiteY) = 255
        Next lngWhiteY
    Next lngWhiteX
    Image.Attributes.Add &H7FE0, &H10, pix

    '拼接新图像
    For i = 1 To imgs.Count
        '计算图像内位移
        intOffsetX = (intImgRectWidth - imgs(i).SizeX - intBorder) / 2
        intOffsetY = (intImgRectHeight - imgs(i).SizeY - intBorder) / 2
        Image.Blt imgs(i), 0, 0, ((i - 1) Mod intCols) * intImgRectWidth + intOffsetX, ((i - 1) \ intCols) * intImgRectHeight + intOffsetY, imgs(i).SizeX, imgs(i).SizeY, 1, 1, 1, False
    Next i

    Set funAssembleImage = Image
End Function

Private Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer, _
    Optional ByVal MaxRows As Integer = 0, Optional ByVal MaxCols As Integer = 0)
    Dim iCols As Integer, iRows As Integer
    
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1

    Do While iRows * iCols > ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols - 1
        Else
            iRows = iRows - 1
        End If
    Loop
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    If MaxRows > 0 And iRows > MaxRows Then
        iRows = MaxRows
        iCols = CInt(ImageCount / iRows)
        If iRows * iCols < ImageCount Then iCols = iCols + 1
    End If
    If MaxCols > 0 And iCols > MaxCols Then
        iCols = MaxCols
        iRows = CInt(ImageCount / iCols)
        If iRows * iCols < ImageCount Then iRows = iRows + 1
    End If
    If MaxRows > 0 And iRows > MaxRows Then iRows = MaxRows
    
    Rows = iRows: Cols = iCols
End Sub
