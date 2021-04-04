Attribute VB_Name = "mdlPublic"
Option Explicit
'------------------------------------------------------------
'类型定义
'------------------------------------------------------------

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

'注册信息类型
Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
End Enum
'gclsPros.SecdInfoRec!改变状态
'删除行：更新行,未改变,替换行均是指表格初始数据加载为经过编辑的行
'更新行：该行的主要特征未改变，次要信息改变了
'替换行: 该行的主要信息发生改变
Public Enum Change_State
    CS_删除行 = -1
    CS_未改变 = 0
    CS_更新行 = 1
    CS_替换行 = 2
    CS_新增行 = 3
End Enum
'gclsPros.MainInfoRec!是否改变
Public Enum Main_Change_State
    MS_不判断 = -1
    MS_未改变 = 0
    MS_改变了 = 1
End Enum

'gclsPros.MainInfoRec!ExpState
'扩展：是否该信息有次级信息记录集记录
'初始扩展:在次级信息记录集初始化时扩展
'加载扩展:在数据加载时扩展
Public Enum Expan_State
    ES_不用扩展 = 0 '不扩展，加“哟”为了代码整齐
    ES_初始扩展 = 1 '初始扩展
    ES_加载扩展 = 2 '加载扩展
End Enum

Public Enum DiagMsgPos
    DMP_诊断次序 = 0
    DMP_诊断类型 = 1
    DMP_疾病编码 = 2
    DMP_诊断编码 = 3
    DMP_疾病附码 = 4
    DMP_疾病类别 = 5
    DMP_证候编码 = 6
    DMP_证候名称 = 7
    DMP_是否疑诊 = 8
End Enum
'首页操作
Public Enum MedRec_Operate
    MOP_设置 = 0
    MOP_预览 = 1
    MOP_打印 = 2
    MOP_确定 = 3
End Enum
'用户信息
Public Type TYPE_USER_INFO
    ID As Long
    编号 As String
    姓名 As String '人员姓名
    简码 As String
    DeptID As Long '部门ID
    DeptNo As String '部门编号
    DeptName As String '部门名称
    DBUser As String '数据库用户
End Type
'编码类型
Public Enum Code_Type
    ' intType参数
    'GetNextNo int序号 参数;IsHavePageNos,IsPageNosCodeRule:intType参数
    CT_病人ID = 1
    CT_住院号 = 2
    CT_住院号ex = 3
    CT_病案号 = 4
    CT_档案号 = 5
End Enum

'-----------------------------------------------------------
'常量
'------------------------------------------------------------
'API:GetSystemMetrics
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21
'GetWindowLong,SetWindowLong
Public Const GWL_WNDPROC = -4&
'CallWindowProc
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Public Const WM_MOUSEWHEEL = &H20A '鼠标滚动消息
Public Const GRD_UNEDITCELL_COLOR = &H8000000B  '未编辑的单元格颜色：灰蓝色
Public Const GRD_LOSTFOCUS_COLORSEL = &H80000010  '离开焦点时,选择的显示颜色
Public Const GRD_GOTFOCUS_COLORSEL = &H8000000D '进入控件时,选择显示颜色

Public Const CB_GETDROPPEDSTATE = &H157 '获取下拉列表状态
Public Const CB_SHOWDROPDOWN = &H14F '关闭或打开下拉列表

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2

Public Const GPAGECOLOR = vbWindowBackground
Public Const SW_RESTORE = 9
Public Const SM_CYFULLSCREEN = 17

'----------------------------------------------------------
'API声明
'----------------------------------------------------------
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private PrevWndProc     As Long

Public Function DecodeEx(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数,并发生变化
'           前一位为Boolean类型，后一位为返回值，遇到第一个为True的值就返回，不再继续判断,最后一位为True的默认值
'          如：ture,3,true,4,则返回3
    Dim i As Integer, blnObjReturn As Boolean
    i = 0
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            If IsObject(arrPar(i)) Then
                Set DecodeEx = arrPar(i): Exit Function
            Else
                DecodeEx = arrPar(i): Exit Function
            End If
        Else
            If arrPar(i) Then
                If IsObject(arrPar(i + 1)) Then
                    Set DecodeEx = arrPar(i + 1): Exit Function
                Else
                    DecodeEx = arrPar(i + 1): Exit Function
                End If
            ElseIf Not blnObjReturn Then
                blnObjReturn = IsObject(arrPar(i + 1))
            End If
            i = i + 2
        End If
    Loop
    If blnObjReturn Then Set DecodeEx = Nothing
End Function

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(gclsPros.TXTProc, hwnd, msg, wp, lp)
End Function

Public Sub SetCboFromList(ByVal arrList As Variant, ByVal arrCboInfo As Variant, Optional ByVal intDefault As Integer = -1)

'功能：将指定数据装入指定ComboBox
'参数：arrList=List String数组
'      arrCboIdx=ComboBox数组,多个ComboBox时,装入数据相同
'      intDefaut=缺省索引
    Dim i As Long, j As Long

    For i = 0 To UBound(arrCboInfo)
        arrCboInfo(i).Clear
        For j = 0 To UBound(arrList)
            arrCboInfo(i).AddItem arrList(j)
        Next
        arrCboInfo(i).ListIndex = intDefault '缺省为未选中
        arrCboInfo(i).Tag = intDefault '保存默认选项，用于清空界面后确定默认值
    Next
End Sub

Public Sub SetCboDefault(objCbo As Object, Optional ByVal intDefault As Integer = -1)
'功能：设置Cbo控件的缺省值
    objCbo.ListIndex = intDefault '缺省为未选中
    objCbo.Tag = intDefault '保存默认选项，用于清空界面后确定默认值
End Sub

Public Sub SetCboDefaultByRec(ByVal arrIndex As Variant)
'功能：通过数据字典设置Cbo控件的缺省值
    Dim i As Long
    Dim objCboTmp As ComboBox
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH
    If TypeName(arrIndex) <> "Variant()" Then
        arrIndex = Array(arrIndex)
    End If
    If TypeName(arrIndex) = "Variant()" Then
        For i = LBound(arrIndex) To UBound(arrIndex)
            Set objCboTmp = gclsPros.CurrentForm.cboBaseInfo(arrIndex(i))
            Set rsTmp = GetBaseCode(arrIndex(i))
            rsTmp.Filter = rsTmp.Filter & " And 缺省=1"
            If Not rsTmp.EOF Then
                Call Cbo.SeekIndex(objCboTmp, Val(rsTmp!ID & ""))
            End If
        Next
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Function Identity(ByRef lngCount As Long) As Long
'功能：模拟主键自增
'参数：lngCount=自增变量
    lngCount = lngCount + 1
    Identity = lngCount
End Function

Public Function GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, ByVal strKEY As String) As String
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的注册信息读取出来
    '入参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '出参数:
    '       返回的键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Dim strKeyValue As String

    On Error GoTo Errhand:

    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKEY, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKEY, "")
        Case g公共全局
            strKeyValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKEY, "")
        Case g公共模块
            strKeyValue = GetSetting("ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKEY, "")
        Case g私有全局
            strKeyValue = GetSetting("ZLSOFT", "私有全局\" & UserInfo.DBUser & "\" & strSection, strKEY, "")
        Case g私有模块
            strKeyValue = GetSetting("ZLSOFT", "私有模块\" & UserInfo.DBUser & "\" & App.ProductName & "\" & strSection, strKEY, "")
    End Select
    GetRegInFor = strKeyValue
    Exit Function
Errhand:

End Function

Public Function ShowMessage(objTmp As Object, ByVal strMsg As String, Optional ByVal blnAsk As Boolean, Optional tbsInfo As TabStrip) As VbMsgBoxResult
'功能：显示提示信息并定位在输入项目上
    Dim lngColor As Long
    On Error GoTo errH
    
    If gclsPros.FuncType <> f诊断选择 Then
        Call LocateObjectPage(objTmp)
    Else
        gclsPros.CurrentForm.tabFunc.Tabs(IIf(objTmp.Name = "vsDiagXY", "西医诊断", "中医诊断")).Selected = True
    End If
    
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        lngColor = objTmp.BackColor: objTmp.BackColor = &HC0C0FF
    Else
        lngColor = objTmp.CellBackColor: objTmp.CellBackColor = &HC0C0FF
        Call objTmp.ShowCell(objTmp.Row, objTmp.Col)
    End If
    
    If Not blnAsk Then
        MsgBox strMsg, vbInformation, gstrSysName
    Else
        ShowMessage = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
    End If
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        objTmp.BackColor = lngColor
    Else
        objTmp.CellBackColor = lngColor
    End If
    If objTmp.Enabled And objTmp.Visible Then
        If TypeName(objTmp) = "TextBox" Then zlControl.TxtSelAll objTmp
        objTmp.SetFocus
    End If
    gclsPros.CurrentForm.Refresh
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function AddErrInfo(ByVal strMsg As String, ByVal intErr As Integer, ParamArray objErr() As Variant) As Boolean
    Dim i As Long
    Dim clsErrTmp As clsErrInfo
    Dim objTmp As Object
    
    On Error GoTo errH
    
    If gclsPros.FuncType <> f医生首页 And gclsPros.FuncType <> f病案首页 Then
        Exit Function
    End If
    
    Set clsErrTmp = New clsErrInfo

    With clsErrTmp
        .IntErrType = intErr
        .StrErrInfo = strMsg
    End With

    For i = LBound(objErr) To UBound(objErr)
        Set objTmp = objErr(i)
        Call clsErrTmp.AddErrObj(objTmp)
    Next
    If intErr = 0 Then
        clsErrTmp.strErrID = "Error-" & CStr(gColErr.Count + 1)
        gColErr.Add clsErrTmp, clsErrTmp.strErrID
    ElseIf intErr = 1 Then
        clsErrTmp.strErrID = "Warn-" & CStr(gColWarn.Count + 1)
        gColWarn.Add clsErrTmp, clsErrTmp.strErrID
    End If
    
    Set clsErrTmp = Nothing
    AddErrInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '功能：提示消息框
    '参数：strMsgInfor-提示信息
    '     blnYesNo-是否提供YES或NO按钮
    '返回：blnYes-如果提供YESNO按钮,则返回YES(True)或NO(False)
    '----------------------------------------------------------------------------------------------------------------
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub

Public Function Calc段内分解时间(ByVal datBegin As Date, ByVal datEnd As Date, ByVal strPause As String, _
    ByVal str执行时间 As String, ByVal int频率次数 As Integer, ByVal int频率间隔 As Integer, ByVal str间隔单位 As String, _
    Optional ByVal dat首日日期 As Date) As String
'功能：按时间段计算各次的分解执行时间及次数
'参数：datBegin-datEnd=要计算的时间段,其中datBegin应为每个周期的开始基准时间
'      strPause=暂停的时间段
'      dat首日日期=用于首日时间计算参照
'返回："时间1,时间2,...."(yyyy-MM-dd HH:mm:ss),时间个数即为次数
'说明：1.时间段内要排除暂停的时间段,次数可能因此而减少
'      2.本函数是假定在执行时间及频率性质完全正确的情况下计算。
    Dim vCurTime As Date, vTmpTime As Date
    Dim arrTime As Variant, arrNormal As Variant, arrFirst As Variant
    Dim blnFirst As Boolean, strDetailTime As String
    Dim strTmp As String, i As Integer

    If InStr(str执行时间, ",") > 0 Then
        arrNormal = Split(Split(str执行时间, ",")(1), "-")
        arrFirst = Split(Split(str执行时间, ",")(0), "-")
    Else
        arrNormal = Split(str执行时间, "-")
        arrFirst = Array()
    End If

    vCurTime = datBegin

    If str间隔单位 = "周" Then
        vCurTime = zlCommFun.GetWeekBase(datBegin)
        If dat首日日期 <> Empty And UBound(arrFirst) <> -1 Then
            blnFirst = (vCurTime = zlCommFun.GetWeekBase(dat首日日期))
        Else
            blnFirst = False
        End If

        Do While vCurTime <= datEnd
            arrTime = IIf(blnFirst, arrFirst, arrNormal)
            blnFirst = False

            '1/8:00-3/15:00-5/9:00
            For i = 1 To int频率次数
                If i - 1 <= UBound(arrTime) Then '首周可能次数不足
                    vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                    If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                        strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                    Else
                        strTmp = Split(arrTime(i - 1), "/")(1)
                    End If
                    vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                    If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                        If Not TimeIsPause(vTmpTime, strPause) Then
                            strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                        End If
                    ElseIf vTmpTime > datEnd Then
                        Exit Do
                    End If
                End If
            Next
            vCurTime = Format(vCurTime + 7, "yyyy-MM-dd") '！！！
        Loop
    ElseIf str间隔单位 = "天" Then
        If dat首日日期 <> Empty And UBound(arrFirst) <> -1 Then
            blnFirst = (Int(vCurTime) = Int(dat首日日期))
        Else
            blnFirst = False
        End If

        Do While vCurTime <= datEnd
            arrTime = IIf(blnFirst, arrFirst, arrNormal)
            blnFirst = False

            If int频率间隔 = 1 Then
                '8:00-12:00-14:00；8-12-14
                For i = 1 To int频率次数
                    If i - 1 <= UBound(arrTime) Then '首日可能次数不足
                        If InStr(arrTime(i - 1), ":") = 0 Then
                            strTmp = arrTime(i - 1) & ":00"
                        Else
                            strTmp = arrTime(i - 1)
                        End If
                        vTmpTime = Format(vCurTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                            If Not TimeIsPause(vTmpTime, strPause) Then
                                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            End If
                        ElseIf vTmpTime > datEnd Then
                            Exit Do
                        End If
                    End If
                Next
            Else
                '1/8:00-1/15:00-2/9:00
                For i = 1 To int频率次数
                    If i - 1 <= UBound(arrTime) Then '首日可能次数不足
                        vTmpTime = vCurTime + Val(Split(arrTime(i - 1), "/")(0)) - 1
                        If InStr(Split(arrTime(i - 1), "/")(1), ":") = 0 Then
                            strTmp = Split(arrTime(i - 1), "/")(1) & ":00"
                        Else
                            strTmp = Split(arrTime(i - 1), "/")(1)
                        End If
                        vTmpTime = Format(vTmpTime, "yyyy-MM-dd") & " " & Format(strTmp, "HH:mm:ss")
                        If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                            If Not TimeIsPause(vTmpTime, strPause) Then
                                strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                            End If
                        ElseIf vTmpTime > datEnd Then
                            Exit Do
                        End If
                    End If
                Next
            End If
            vCurTime = Format(vCurTime + int频率间隔, "yyyy-MM-dd") '！！！
        Loop
    ElseIf str间隔单位 = "小时" Then
        '10:00-20:00-40:00；10-20-40；02:30
        arrTime = arrNormal
        Do While vCurTime <= datEnd
            For i = 1 To int频率次数
                If InStr(arrTime(i - 1), ":") = 0 Then
                    vTmpTime = vCurTime + (arrTime(i - 1) - 1) / 24
                Else
                    vTmpTime = vCurTime + (Split(arrTime(i - 1), ":")(0) - 1) / 24 + Split(arrTime(i - 1), ":")(1) / 60 / 24
                End If
                vTmpTime = Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                    If Not TimeIsPause(vTmpTime, strPause) Then
                        strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                    End If
                ElseIf vTmpTime > datEnd Then
                    Exit Do
                End If
            Next
            vCurTime = Format(vCurTime + int频率间隔 / 24, "yyyy-MM-dd HH:mm:ss")
        Loop
    ElseIf str间隔单位 = "分钟" Then
        '无执行时间
        Do While vCurTime <= datEnd
            vTmpTime = vCurTime

            If vTmpTime >= datBegin And vTmpTime <= datEnd Then
                If Not TimeIsPause(vTmpTime, strPause) Then
                    strDetailTime = strDetailTime & "," & Format(vTmpTime, "yyyy-MM-dd HH:mm:ss")
                End If
            ElseIf vTmpTime > datEnd Then
                Exit Do
            End If

            vCurTime = Format(vCurTime + int频率间隔 / (24 * 60), "yyyy-MM-dd HH:mm:ss")
        Loop
    End If

    Calc段内分解时间 = Mid(strDetailTime, 2)
End Function


Public Function TimeIsPause(vDate As Date, strPause As String) As Boolean
'功能：判断一个时间是否在暂停的时间段中
'参数：strPause="暂停时间,开始时间;...."
    Dim arrPause() As String, i As Long
    Dim strBegin As String, strEnd As String

    If strPause = "" Then Exit Function
    arrPause = Split(strPause, ";")
    For i = 0 To UBound(arrPause)
        strBegin = Split(arrPause(i), ",")(0)
        strEnd = Split(arrPause(i), ",")(1)
        If strEnd = "" Then strEnd = "3000-01-01 00:00:00" '可能尚未启用或暂停的时候被停止
        If Between(Format(vDate, "yyyy-MM-dd HH:mm:ss"), strBegin, strEnd) Then
            TimeIsPause = True: Exit Function
        End If
    Next
End Function

Public Function GetTextByDot(ByVal strText As String, Optional ByVal blnBefore As Boolean, Optional ByVal strSpliter As String = ".") As String
'功能: 得到圆点之后或之前的文本
    If blnBefore Then
        If InStr(strText, strSpliter) > 0 Then
            GetTextByDot = Mid(strText, 1, InStr(strText, strSpliter) - 1)
        End If
    Else
        GetTextByDot = Mid(strText, InStr(strText, strSpliter) + Len(strSpliter))
    End If
End Function

Public Function PrintVsCol(ByRef vsTmp As VSFlexGrid, Optional ByVal strName As String) As String
'功能：打印表格列,用来设置列宽
    Dim i As Long
    Dim strTmp As String
    With vsTmp
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) Then
                strTmp = strTmp & ";" & .TextMatrix(0, i) & "," & .ColWidth(i)
            End If
        Next
        strTmp = Mid(strTmp, 2)
        PrintVsCol = strName & "==" & strTmp
    End With
End Function

Public Function CreatePlugInOK(ByVal lngMod As Long) As Boolean
'功能：外挂创建与检查
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, gclsPros.SysNo, lngMod)
            Call zlPlugInErrH(Err, "Initialize")
        End If
        Err.Clear: On Error GoTo 0
    End If
    If Not gobjPlugIn Is Nothing Then CreatePlugInOK = True
End Function

Public Sub SetCtrlLocked(ByRef objInput As Object, ByVal blnLocked As Boolean, Optional ByVal blnClear As Boolean, Optional ByVal blnSetForeColor As Boolean)
'功能：锁定或解锁控件
'参数：objInput=控件对象
'         blnLocked=是否锁定
'         blnClear=是否控件输入内容
    Dim strType  As String
    Dim objCmd As CommandButton
    Dim strTmp As String

    On Error GoTo errH
    strType = TypeName(objInput)
    Select Case strType
        Case "TextBox", "ComboBox"
            If strType = "TextBox" Then
                '寻找对应按钮
                On Error Resume Next
                If objInput.Name = "txtSpecificInfo" Then
                    Set objCmd = gclsPros.CurrentForm.cmdSpecificInfo(objInput.Index)
                    strTmp = objCmd.Name
                ElseIf objInput.Name = "txtInfo" Then
                    Set objCmd = gclsPros.CurrentForm.cmdInfo(objInput.Index)
                    strTmp = objCmd.Name
                ElseIf objInput.Name = "txtAdressInfo" Then
                    Set objCmd = gclsPros.CurrentForm.cmdAdressInfo(objInput.Index)
                    strTmp = objCmd.Name
                End If
                If Err.Number = 0 Then
                    Call SetCtrlLocked(objCmd, blnLocked)
                    On Error GoTo errH
                Else
                    Err.Clear
                    On Error GoTo errH
                End If
            End If
            If blnClear And blnLocked Then
                If strType = "ComboBox" Then
                    Call zlControl.CboSetIndex(objInput.hwnd, -1)
                Else
                    objInput.Text = ""
                End If
            End If
            objInput.Locked = blnLocked
            objInput.BackColor = IIf(blnLocked, vbButtonFace, vbWindowBackground)
            objInput.TabStop = Not blnLocked
            If blnSetForeColor Then objInput.ForeColor = IIf(blnLocked, &HFF0000, &H80000008)
        Case "CheckBox"
            If blnClear And blnLocked Then
                objInput.Value = 0
            End If
            objInput.Enabled = Not blnLocked
'            objInput.BackColor = IIf(blnLocked, vbButtonFace, &H8000000F)
            objInput.TabStop = Not blnLocked
            If blnSetForeColor Then objInput.ForeColor = IIf(blnLocked, &HFF0000, &H80000008)
        Case "CommandButton"
            objInput.Enabled = Not blnLocked
        Case "MaskEdBox", "MonthView", "ListBox"
            objInput.Enabled = Not blnLocked
            objInput.BackColor = IIf(blnLocked, vbButtonFace, vbWindowBackground)
            objInput.TabStop = Not blnLocked
            If blnClear And blnLocked And strType = "MaskEdBox" Then
                objInput.Text = Replace(objInput.Mask, "#", "_")
            End If
            If objInput.Name = "mskDateInfo" Then
                Call SetCtrlLocked(gclsPros.CurrentForm.txtDateInfo(objInput.Index), blnLocked, blnClear, blnSetForeColor)
                gclsPros.CurrentForm.txtDateInfo(objInput.Index).Text = objInput.Text
                gclsPros.CurrentForm.txtDateInfo(objInput.Index).Visible = blnLocked
                objInput.Visible = Not blnLocked
            Else
                If blnSetForeColor Then objInput.ForeColor = IIf(blnLocked, &HFF0000, &H80000008)
            End If
        Case "VSFlexGrid"
            '同时注意要在键盘鼠标事件中进行一些控制
            objInput.Editable = IIf(blnLocked, flexEDNone, flexEDKbdMouse)
            objInput.BackColor = IIf(blnLocked, vbButtonFace, vbWindowBackground)
            objInput.BackColorBkg = IIf(blnLocked, vbButtonFace, vbWindowBackground)
        Case "PatiAddress"
            objInput.ControlLock = blnLocked
        Case "OptionButton"
            objInput.Enabled = Not blnLocked
'            objInput.BackColor = IIf(blnLocked, vbButtonFace, &H8000000F)
        Case "Label"
            objInput.ForeColor = IIf(Not blnLocked, gclsPros.CurrentForm.ForeColor, &H808080)
    End Select
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function ControlHaveValue(ByRef objInput As Object) As Boolean
'功能：锁定或解锁控件
'参数：objInput=控件对象
'         blnLocked=是否锁定
'         blnClear=是否控件输入内容
    Dim strType  As String

    On Error GoTo errH
    strType = TypeName(objInput)
    Select Case strType
        Case "TextBox", "ComboBox"
            ControlHaveValue = objInput.Text <> ""
        Case "CheckBox"
            ControlHaveValue = objInput.Value <> 0
        Case "MaskEdBox"
            ControlHaveValue = IsDate(objInput.Text)
        Case "PatiAddress"
            ControlHaveValue = objInput.Value <> ""
    End Select

    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ControlIsLocked(ByRef objInput As Object) As Boolean
'功能：锁定或解锁控件
'参数：objInput=控件对象
'         blnLocked=是否锁定
'         blnClear=是否控件输入内容
    Dim strType  As String

    On Error GoTo errH
    strType = TypeName(objInput)
    Select Case strType
        Case "TextBox", "ComboBox"
            ControlIsLocked = objInput.Locked
        Case "VSFlexGrid"
            ControlIsLocked = objInput.Editable = flexEDNone
        Case "PatiAddress"
            ControlIsLocked = objInput.ControlLock
        Case Else
            ControlIsLocked = Not objInput.Enabled
    End Select
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub zlVsGridLostFocus(ByVal vsGrid As VSFlexGrid, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
   '功能：离开网格控件时选择的颜色
    '入参：CustomColor-是否用自定义颜色来设置(BackColor)的方式来进行)
    '编制：刘兴洪
    '日期：2010-03-23 11:03:05
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    With vsGrid
        If CustomColor <> -1 Then
             If .Row >= .FixedRows Then
                .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = CustomColor
            End If
        Else
            .SelectionMode = flexSelectionByRow
            .FocusRect = IIf(vsGrid.Editable = flexEDNone, flexFocusHeavy, flexFocusSolid)
            .HighLight = flexHighlightAlways
            .BackColorSel = GRD_LOSTFOCUS_COLORSEL
        End If
    End With
End Sub

Public Function GetFormat(ByVal strMask As String) As String
    GetFormat = Decode(strMask, "####-##-##", "yyyy-mm-dd", "####-##-## ##:##", "yyyy-mm-dd hh:mm", "####-##-## ##:##:##", "yyyy-mm-dd hh:mm:ss", "##;##", "hh:mm", "")
End Function

Public Function SubWndProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'屏蔽掉控件的鼠标滚轮事件
    Select Case msg    '在这里进行过滤.如果知道其他的消息,也可以在这里过滤.
        Case WM_MOUSEWHEEL
            SubWndProc = 1 '屏蔽掉
            Exit Function
    End Select
    SubWndProc = CallWindowProc(PrevWndProc, hwnd, msg, wParam, lParam)             '其它消息不管
End Function

Public Sub CallHook(ByVal hwnd As Long)
    PrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SubWndProc)
End Sub

Public Sub CallUnhook(ByVal hwnd As Long)
    Dim lngReturnValue As Long
    lngReturnValue = SetWindowLong(hwnd, GWL_WNDPROC, PrevWndProc)
End Sub

