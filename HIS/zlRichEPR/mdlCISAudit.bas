Attribute VB_Name = "mdlCISAudit"
'#########################################################################
'##模 块 名：mdlCISAudit.bas
'##创 建 人：祝庆
'##日    期：2012年9月26日
'##修 改 人：
'##日    期：
'##描    述：公共函数声明等
'##版    本：
'#########################################################################

Option Explicit


'######################################################################################################################
'常量定义

Public Const strSplitCmb = "—"

Public Enum COLOR
    白色 = &H80000005
    红色 = &HFF&
    兰色 = &HFF0000
    黑色 = 0
    非焦点 = &HFFEBD7
    焦点 = &HFFCC99
    浅灰色 = &HE0E0E0
    深灰色 = &H8000000C
    灰色 = &H8000000F
    浅黄色 = &H80000018
    锁色 = &HF5F5F5
    启用色 = 0
    停用色 = 255
    拖动色 = &HFFE0D9
    公共模块色 = &HC00000
    
    报警背景色 = &H40C0&
    报警前景色 = &H8000000E
    超标背景色 = &H80C0FF
    低标背景色 = &H80FFFF
    超标前景色 = &H80000012
    默认前景色 = &H80000008
    
End Enum

'枚举
Public Enum COLOR_NativeXpPlain
    BackgroundDark = 14054755
    BackgroundLight = 15180411
    HighlightBorderBottomRight = 8388608
    HighlightBorderTopLeft = 8388608
    HighlightHot = 12775167
    HighlightPressed = 4096254
    HighlightSelected = 7323903
    NormalGroupCaptionDark = 14215660
    NormalGroupCaptionLight = 14215660
    NormalGroupCaptionTextHot = 0
    NormalGroupCaptionTextNormal = 0
    NormalGroupClient = 16244694
    NormalGroupClientBorder = 16777215
    NormalGroupClientLink = 12999969
    NormalGroupClientLinkHot = 16748098
    NormalGroupClientText = 0
    SpecialGroupCaptionDark = 14215660
    SpecialGroupCaptionLight = 14215660
    SpecialGroupCaptionTextHot = 0
    SpecialGroupCaptionTextSpecial = 0
    SpecialGroupClient = 16244694
    SpecialGroupClientBorder = 16777215
    SpecialGroupClientLink = 12999969
    SpecialGroupClientLinkHot = 16748098
    SpecialGroupClientText = 0
End Enum

Public Enum REGISTER
    注册信息
    私有模块
    私有全局
    公共模块
    公共全局
End Enum

Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
End Enum

Private mstrSQL As String
Private mstrTitle As String

'==================================================================================================
'=名称:去掉字符串中的单引号("'")(ConvertString)
'=入口参数:
'=1).sStr          类型:String
'=出口参数:空
'=功能:去掉字符串(sStr)中的单引号
'=日期:2004-12-11
'=编程:欧阳
'=说明:在SQL语句中不能带单引号
'==================================================================================================
Function ConvertString(ByVal sStr As String) As String
On Error GoTo ErrH
    ConvertString = Replace(sStr, "'", "")
    ConvertString = Replace(ConvertString, "—", "")
    ConvertString = Replace(ConvertString, "&", "")
    Exit Function
ErrH:
    Err.Clear
    ConvertString = ""
End Function

'/******************************************/
'=函数:BigNote
'=功能:放大备注字段编辑框,并返回编辑后的文本
' 参数:mStr   已经库存的字符串
'      mTitle 编辑窗口的标题
'=编程:朱红军
'=日期:2002-04-03
'/******************************************/
Function Big_Note(mStr As String, mTitle As String, Optional bReadOnly As Boolean, Optional bSqlCheck As Boolean = False) As String
On Error GoTo ErrH
    With FrmNoteBox
        .SqlCheck = bSqlCheck
        .StrText = mStr
        .StrTile = mTitle
        .ReadOnly = bReadOnly
        .Show vbModal
        DoEvents
        Big_Note = IIf(bSqlCheck, .StrText, ConvertString(.StrText))
    End With
    Set FrmNoteBox = Nothing
    Exit Function
ErrH:
    Err.Clear
End Function

Public Function CheckAuditSql_IN(strSQL As String, Optional blnMsg As Boolean = False, Optional intSource As Integer = 0) As Boolean

'=功能： 检测SQL主句功能 通过 函数 判断正误
'intSource 数据源 =0 表明是在标准版内执行，=1表明在EMR库执行
'返回：正常通过返回True
    Dim rsTemp          As ADODB.Recordset
    Dim zlCheck         As New clsCheck
    Dim strReturn       As String, strEMRSQL As String, strParm As String
    On Error GoTo ErrH
    If strSQL = "" Then CheckAuditSql_IN = True: Exit Function
    If strSQL = "" Then Exit Function
    strSQL = UCase(strSQL)
    If intSource = 1 Then
        'EMR库的查询不能被转成大写执行，有些情况可能需要区分大小写,转成大写只用于检查是否存在更新删除语句
        strEMRSQL = Replace(strSQL, "[MID]", "Hextoraw(:mid)")
        strEMRSQL = Replace(strEMRSQL, "[ALIDIN]", "Hextoraw(:alidin)")
    Else
        strSQL = Replace(strSQL, "[病人ID]", "-1")
        strSQL = Replace(strSQL, "[主页ID]", "-1")
        strSQL = "Select * From (" & strSQL & ") "
    End If
    
    If InStr(1, strSQL, "INSERT") = 1 Then
        zlCheck.Msg_OK "语法检测失败！不能写入数据！", vbCritical

        Exit Function
    ElseIf InStr(1, strSQL, "UPDATE") = 1 Then
        zlCheck.Msg_OK "语法检测失败！不能更新数据！", vbCritical

        Exit Function
    ElseIf InStr(1, strSQL, "DELETE") = 1 Then
        zlCheck.Msg_OK "语法检测失败！不能删除数据！", vbCritical

        Exit Function
    ElseIf InStr(1, strSQL, ";") > 0 Then
        zlCheck.Msg_OK "语法检测失败！不能使用“;”！", vbCritical

        Exit Function
    End If
    
    If intSource = 1 Then
        strParm = IIf(InStr(strEMRSQL, ":mid") = 0, "", "A^16^mid")
        If InStr(strEMRSQL, ":alidin") > 0 Then
            If InStr(strEMRSQL, ":mid") > 0 Then
                strParm = strParm & "|"
            End If
            strParm = strParm & "A^16^mid"
        End If
        strReturn = gobjEmr.OpenSQLRecordset(strEMRSQL, strParm, rsTemp)
        If (strReturn <> "" And InStr(strReturn, "ORA-01403") = 0) Or rsTemp Is Nothing Then
            zlCheck.Msg_OK "审查依据检测失败:" & vbCrLf & "【" & strReturn & "】" & vbCrLf & "请重新录入审查依据检测语句！", vbExclamation
            Exit Function
        End If
    Else
        strSQL = "Select * From (" & strSQL & ") "
        Set rsTemp = zlDatabase.OpenSQLRecord("select ZL_FUN_ExecSql('" & Replace(strSQL, "'", "''") & "') from dual", "mdlCISAudit")
        If rsTemp Is Nothing Then
            zlCheck.Msg_OK "语法检测失败！", vbCritical
            Exit Function
        Else
            If InStr(1, rsTemp.Fields(0), "[zlsoft]Error[zlsoft]:ORA-01403") > 0 Then
                '没找到任何数据
            ElseIf InStr(1, rsTemp.Fields(0), "[zlsoft]Error[zlsoft]") > 0 Then
                zlCheck.Msg_OK "审查依据检测失败:" & vbCrLf & "【" & Mid(rsTemp.Fields(0), 23) & "】" & vbCrLf & "请重新录入审查依据检测语句！", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    If blnMsg Then zlCheck.Msg_OK "语法检测检测成功！"
    CheckAuditSql_IN = True
    Set zlCheck = Nothing
    Exit Function
ErrH:
    zlCheck.Msg_OK "语法检测检测失败！" & vbCrLf & Err.Description, vbCritical
    '判断状态 无需进错误处理中心
    Err.Clear

    Set zlCheck = Nothing
End Function

'==============================================================================
'=功能： 检测SQL主句功能
'==============================================================================
Public Function CheckAuditSql_OUT(strSQL As String, Optional lng病人ID As Long = -1, Optional lng主页ID As Long = -1) As String
    On Error GoTo ErrH
    strSQL = UCase(strSQL)
    strSQL = Replace(strSQL, "[病人ID]", CStr(lng病人ID))
    strSQL = Replace(strSQL, "[主页ID]", CStr(lng主页ID))
    
    CheckAuditSql_OUT = strSQL
        
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetEMR_MID_ALIDIN(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByRef strMid As String, ByRef strAlidin As String) As Boolean
    On Error GoTo ErrHandle
    Dim strReturn As String, strExtend_Tag As String, rsTemp As New ADODB.Recordset
    strExtend_Tag = GetEMRIn_Tag(lngPatiID, lngPageId)
    If strExtend_Tag = "" Then Exit Function
    gstrSQL = "Select Rawtohex(ID) As ID, Rawtohex(Master_Id) As Master_Id From Bz_Act_Log Where Extend_Tag = :extendtag"
    strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, strExtend_Tag & "^16^extendtag", rsTemp)
    If strReturn <> "" Then Exit Function
    strMid = rsTemp!Master_id
    strAlidin = rsTemp!ID
    
    GetEMR_MID_ALIDIN = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'==============================================================================
'=功能： 审查项目 适用对象 与 病历种类对照
'==============================================================================
Public Function AuditFileTran(strUsed As String, intSource As Integer) As String
    Dim strType     As String
    On Error GoTo ErrH
    Select Case strUsed
        Case "2" '住院病历
            strType = IIf(intSource = 0, "2", "02")
        Case "3" '护理病历
            strType = IIf(intSource = 0, "4", "03")
        Case "4" '护理记录
            strType = "3"
        Case "6" '医嘱报告
            strType = "7"
        Case "7" '疾病证明
            strType = IIf(intSource = 0, "5", "04")
        Case "8" '知情文件
            strType = IIf(intSource = 0, "6", "05")
    End Select
    AuditFileTran = strType
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Function GetPPFS() As String
    On Error GoTo ErrH
    '输入匹配
    GetPPFS = ""
    If Val(zlDatabase.GetPara("输入匹配")) = 0 Then
        GetPPFS = "%"
    End If
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '******************************************************************************************************************
    '功能:获取特殊时间
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim intDay As Integer
    
    Select Case strMode
    Case "当  时"      '当时
        GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "今  天"       '当天
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  周"       '本周,bytFlag=1,本周开始时间,=2,本周结束时间
        intDay = Weekday(CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 7 - intDay, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  月"       '本月
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM") & "-01 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  季"      '本季度
        Select Case Format(zlDatabase.Currentdate, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-04-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-10-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "本半年"      '本半年
        If Val(Format(zlDatabase.Currentdate, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "本  年"   '全年
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
        End If
    Case "昨  天"       '昨天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "明  天"       '明天
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前三天"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -3, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一周"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -7, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前半月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -15, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -30, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前二月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -60, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前三月"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -90, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "前半年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -180, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前一年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前二年"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365 * 2, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    End Select
    
End Function


Public Function DockPannelCreate(ByRef dkpMain As DockingPane, ByVal intIndex As Integer, _
                                    ByVal lngCX As Long, ByVal lngCY As Long, _
                                    ByVal bytDirection As DockingDirection, _
                                    Optional ByVal objNeighbour As Pane = Nothing, _
                                    Optional ByVal strTitle As String, _
                                    Optional ByVal bytOptions As PaneOptions) As Pane
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Set DockPannelCreate = dkpMain.CreatePane(intIndex, lngCX, lngCY, bytDirection, objNeighbour)
    DockPannelCreate.Title = strTitle
    DockPannelCreate.Options = PaneNoCaption
    
End Function

Public Function TabControlInit(ByRef tbc As TabControl, _
                                Optional ByVal bytAppearance As XTPTabAppearanceStyle = xtpTabAppearancePropertyPage2003) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    With tbc
        
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
'            .Position = bytPosition
        End With
        
        Set .Icons = frmPubResource.imgPublic.Icons
        

        
    End With

    TabControlInit = True
    
End Function

Public Function CopyMenu(ByVal cbsMain As Object, Optional ByVal intNo As Integer = 2) As CommandBar
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrPopupItem2 As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    
    '弹出菜单处理
    
    On Error GoTo errHand
    
    If cbsMain.ActiveMenuBar.Controls(intNo).Visible = False Then Exit Function

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(intNo)
    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.Parameter = cbrControl.Parameter
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        
        If cbrControl.Type = xtpControlButtonPopup Then
            For Each cbrControl2 In cbrControl.CommandBar.Controls
                Set cbrPopupItem2 = cbrPopupItem.CommandBar.Controls.Add(xtpControlButton, cbrControl2.ID, cbrControl2.Caption)
                cbrPopupItem2.Parameter = cbrControl2.Parameter
            Next
        End If
        
    Next
    
    Set CopyMenu = cbrPopupBar
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub SendLMouseButton(ByVal lngHwnd As Long, ByVal X As Single, ByVal Y As Single)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngX As Long
    Dim lngY As Long
    Dim lngLoop As Long
    Dim lngXY As Long
            
    lngX = X / 15
    lngY = Y / 15
        
    lngXY = 2
    For lngLoop = 1 To 15
        lngXY = lngXY * 2
    Next
    
    lngXY = lngXY * lngY + lngX
    
    SendMessage lngHwnd, WM_LBUTTONDOWN, 0, ByVal lngXY
    SendMessage lngHwnd, WM_LBUTTONUP, 0, ByVal lngXY

End Sub



Public Function IncStr(ByVal strVal As String) As String
    '******************************************************************************************************************
    '功能：对一个字符串自动加1。
    '说明：每一位进位时,如果是数字,则按十进制处理,否则按26进制处理
    '******************************************************************************************************************
    Dim i As Long, strTmp As String, bytUp As Byte, bytAdd As Byte
    
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

Public Function RestoreTaskPanelPaterrn(ByVal objTpl As Object)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    With objTpl
        
        .ColorSet.BackgroundDark = COLOR_NativeXpPlain.BackgroundDark
        .ColorSet.BackgroundLight = COLOR_NativeXpPlain.BackgroundLight
        .ColorSet.HighlightBorderBottomRight = COLOR_NativeXpPlain.HighlightBorderBottomRight
        .ColorSet.HighlightBorderTopLeft = COLOR_NativeXpPlain.HighlightBorderTopLeft
        .ColorSet.HighlightHot = COLOR_NativeXpPlain.HighlightHot
        .ColorSet.HighlightPressed = COLOR_NativeXpPlain.HighlightPressed
        .ColorSet.HighlightSelected = COLOR_NativeXpPlain.HighlightSelected
        
        .ColorSet.NormalGroupCaptionDark = COLOR_NativeXpPlain.NormalGroupCaptionDark
        .ColorSet.NormalGroupCaptionLight = COLOR_NativeXpPlain.NormalGroupCaptionLight
        .ColorSet.NormalGroupCaptionTextHot = COLOR_NativeXpPlain.NormalGroupCaptionTextHot
        .ColorSet.NormalGroupCaptionTextNormal = COLOR_NativeXpPlain.NormalGroupCaptionTextNormal
        .ColorSet.NormalGroupClient = COLOR_NativeXpPlain.NormalGroupClient
        .ColorSet.NormalGroupClientBorder = COLOR_NativeXpPlain.NormalGroupClientBorder
        .ColorSet.NormalGroupClientLink = COLOR_NativeXpPlain.NormalGroupClientLink
        
        .ColorSet.NormalGroupClientLinkHot = COLOR_NativeXpPlain.NormalGroupClientLinkHot
        .ColorSet.NormalGroupClientText = COLOR_NativeXpPlain.NormalGroupClientText
        .ColorSet.SpecialGroupCaptionDark = COLOR_NativeXpPlain.SpecialGroupCaptionDark
        .ColorSet.SpecialGroupCaptionLight = COLOR_NativeXpPlain.SpecialGroupCaptionLight
        .ColorSet.SpecialGroupCaptionTextHot = COLOR_NativeXpPlain.SpecialGroupCaptionTextHot
        .ColorSet.SpecialGroupCaptionTextSpecial = COLOR_NativeXpPlain.SpecialGroupCaptionTextSpecial
        .ColorSet.SpecialGroupClient = COLOR_NativeXpPlain.SpecialGroupClient
        .ColorSet.SpecialGroupClientBorder = COLOR_NativeXpPlain.SpecialGroupClientBorder
        .ColorSet.SpecialGroupClientLink = COLOR_NativeXpPlain.SpecialGroupClientLink
        .ColorSet.SpecialGroupClientLinkHot = COLOR_NativeXpPlain.SpecialGroupClientLinkHot
        .ColorSet.SpecialGroupClientText = COLOR_NativeXpPlain.SpecialGroupClientText
    End With
End Function

Public Function RestoreDockPanelPaterrn(ByVal objDkp As Object)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    With objDkp
        
        .ColorSet.BackgroundDark = COLOR_NativeXpPlain.BackgroundDark
        .ColorSet.BackgroundLight = COLOR_NativeXpPlain.BackgroundLight
        .ColorSet.HighlightBorderBottomRight = COLOR_NativeXpPlain.HighlightBorderBottomRight
        .ColorSet.HighlightBorderTopLeft = COLOR_NativeXpPlain.HighlightBorderTopLeft
        .ColorSet.HighlightHot = COLOR_NativeXpPlain.HighlightHot
        .ColorSet.HighlightPressed = COLOR_NativeXpPlain.HighlightPressed
        .ColorSet.HighlightSelected = COLOR_NativeXpPlain.HighlightSelected
        
        .ColorSet.NormalGroupCaptionDark = COLOR_NativeXpPlain.NormalGroupCaptionDark
        .ColorSet.NormalGroupCaptionLight = COLOR_NativeXpPlain.NormalGroupCaptionLight
        .ColorSet.NormalGroupCaptionTextHot = COLOR_NativeXpPlain.NormalGroupCaptionTextHot
        .ColorSet.NormalGroupCaptionTextNormal = COLOR_NativeXpPlain.NormalGroupCaptionTextNormal
        .ColorSet.NormalGroupClient = COLOR_NativeXpPlain.NormalGroupClient
        .ColorSet.NormalGroupClientBorder = COLOR_NativeXpPlain.NormalGroupClientBorder
        .ColorSet.NormalGroupClientLink = COLOR_NativeXpPlain.NormalGroupClientLink
        
        .ColorSet.NormalGroupClientLinkHot = COLOR_NativeXpPlain.NormalGroupClientLinkHot
        .ColorSet.NormalGroupClientText = COLOR_NativeXpPlain.NormalGroupClientText
        .ColorSet.SpecialGroupCaptionDark = COLOR_NativeXpPlain.SpecialGroupCaptionDark
        .ColorSet.SpecialGroupCaptionLight = COLOR_NativeXpPlain.SpecialGroupCaptionLight
        .ColorSet.SpecialGroupCaptionTextHot = COLOR_NativeXpPlain.SpecialGroupCaptionTextHot
        .ColorSet.SpecialGroupCaptionTextSpecial = COLOR_NativeXpPlain.SpecialGroupCaptionTextSpecial
        .ColorSet.SpecialGroupClient = COLOR_NativeXpPlain.SpecialGroupClient
        .ColorSet.SpecialGroupClientBorder = COLOR_NativeXpPlain.SpecialGroupClientBorder
        .ColorSet.SpecialGroupClientLink = COLOR_NativeXpPlain.SpecialGroupClientLink
        .ColorSet.SpecialGroupClientLinkHot = COLOR_NativeXpPlain.SpecialGroupClientLinkHot
        .ColorSet.SpecialGroupClientText = COLOR_NativeXpPlain.SpecialGroupClientText
    End With
End Function

Public Function SetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strKeyValue As String) As Boolean
    '******************************************************************************************************************
    '功能： 将指定的信息保存在注册表中
    '参数： enmRegister-注册类型
    '       strSection-注册表目录
    '       strKey-键名
    '       strKeyValue-键值
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Select Case enmRegister
    Case 注册信息
        
        Call SaveSetting("ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue)
        
    Case 私有模块

        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case 私有全局

        Call SaveSetting("ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, strKeyValue)
        
    Case 公共模块

        Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case 公共全局
        
        Call SaveSetting("ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue)
        
    End Select
    
    SetRegister = True
    
errHand:
    
End Function

Public Function GetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strDefKeyValue As String) As String
    '******************************************************************************************************************
    '功能： 将指定的注册信息读取出来
    '参数： enmRegister-注册类型
    '       strSection-注册表目录
    '       strKey-键名
    '       strDefKeyValue-缺省键值
    '返回： strKeyValue-键值
    '******************************************************************************************************************

    Dim strValue As String
    
    On Error GoTo errHand
    
    Select Case enmRegister
    Case 注册信息
        
        strValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, strDefKeyValue)
        
    Case 私有模块

        strValue = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case 私有全局

        strValue = GetSetting("ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, strDefKeyValue)
        
    Case 公共模块

        strValue = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case 公共全局
        
        strValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, strDefKeyValue)
        
    End Select
    
    GetRegister = strValue
    
errHand:
End Function

Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
    '******************************************************************************************************************
    '功能：设置指定的参数值
    '参数：varPara=参数号或参数名，以数字或字符类型传入区分
    '      strValue=要设置的参数值
    '      lngModual=使用该参数的模块号，如1230
    '      blnPrivate=该参数是否用户私有参数
    '返回：设置是否成功
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    GetPara = zlDatabase.GetPara(varPara, glngSys, lngModual, strDefault, blnNotCache)

errHand:

End Function

Public Function SetPara(ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：设置指定的参数值
    '参数：varPara=参数号或参数名，以数字或字符类型传入区分
    '      strValue=要设置的参数值
    '      lngModual=使用该参数的模块号，如1230
    '      blnPrivate=该参数是否用户私有参数
    '返回：设置是否成功
    '******************************************************************************************************************

    On Error GoTo ErrH
        
    SetPara = zlDatabase.SetPara(varPara, strValue, glngSys, lngModual, blnSetup)

    Exit Function
    
ErrH:

End Function


Public Function ParamCreate(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "参数名", adVarChar, 50
        .Fields.Append "参数值", adVarChar, 50
        
        .Open
    End With
    
    ParamCreate = True
    
    Exit Function
    
errHand:
    
End Function

Public Function ParamAdd(ByRef rs As ADODB.Recordset, ByVal strParamName As String, Optional ByVal strParamValue As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.AddNew
    
    rs("参数名").Value = strParamName
    rs("参数值").Value = strParamValue
    
    ParamAdd = True
    
    Exit Function
    
errHand:
End Function

Public Function ParamRead(ByRef rs As ADODB.Recordset, ByVal strParamName As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.Filter = ""
    rs.Filter = "参数名='" & strParamName & "'"
    If rs.RecordCount > 0 Then
        ParamRead = rs("参数值").Value
    End If
    
    Exit Function
    
errHand:
End Function

Public Function ParamWrite(ByRef rs As ADODB.Recordset, ByVal strParamName As String, ByVal strParamValue As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.Filter = ""
    rs.Filter = "参数名='" & strParamName & "'"
    If rs.RecordCount > 0 Then
        rs("参数值").Value = strParamValue
    End If
    
    Exit Function
    
errHand:
End Function


Public Function ShowPubSelect(ByVal frmParent As Object, _
                                ByVal obj As Object, _
                                ByVal bytStyle As Byte, _
                                ByVal strLvw As String, _
                                ByVal strSavePath As String, _
                                ByVal strDescrible As String, _
                                ByVal rsData As ADODB.Recordset, _
                                ByRef rsResult As ADODB.Recordset, _
                                Optional ByVal lngCX As Long = 9000, _
                                Optional ByVal lngCY As Long = 4500, _
                                Optional ByVal blnMuliSel As Boolean = False, _
                                Optional ByVal strInitKey As String = "", _
                                Optional ByVal strFilterControl As String = "", _
                                Optional ByVal blnOneReturn As Boolean = False) As Byte
    '******************************************************************************************************************
    '功能：打开树型+列表结构,应用于表格控件
    '参数：
    '      bytStyle:1-TreeView;2-ListView;3-TreeView+ListView
    '返回：0:取消选择;1:选择;2:无数据返回
    '******************************************************************************************************************
    
    Dim lngX As Long
    Dim lngY As Long
    Dim lngObjHeight As Long
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI

    On Error GoTo errHand
    
    If rsData.BOF Then
        ShowPubSelect = 2
        Exit Function
    End If
    
    If blnOneReturn Then
        If rsData.RecordCount = 1 Then
            Set rsResult = rsData
            ShowPubSelect = 1
            Exit Function
        End If
    End If
    
    Call ClientToScreen(obj.hWnd, objPoint)
    
    Select Case TypeName(obj)
    Case "TextBox", "CommandButton"
    
        lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
        lngY = obj.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
        lngObjHeight = obj.Height
        
    Case Else
        lngX = objPoint.X * Screen.TwipsPerPixelX + obj.CellLeft
        lngY = objPoint.Y * Screen.TwipsPerPixelY + obj.CellTop + obj.CellHeight
        lngObjHeight = obj.CellHeight
    End Select
    
    ShowPubSelect = frmPubSelDialog.ShowDialog(frmParent, bytStyle, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, lngObjHeight, strInitKey, strSavePath, , False, blnMuliSel, strFilterControl)
                                
    If ShowPubSelect = 1 Then
        Set rsResult = rsData
        
        If rsResult.BOF Then
            ShowPubSelect = 0
        End If
        
    End If

    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function



Public Function GetQuestion(rsCondition As ADODB.Recordset, ByVal strDept As String, ByVal bytApplyMode As Byte, Optional ByVal lngKey As Long, Optional ByVal strStart As String, Optional ByVal strEnd As String, Optional ByVal lng病人ID As Long = -1, Optional ByVal lng主页ID As Long, Optional ByVal lngCur次数 As Long, Optional ByVal str开始时间 As String, Optional ByVal str结束时间 As String, Optional ByVal str反馈人 As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim strTmp As String
    Dim strSQL As String
    Dim int病人类型 As Integer
    Dim strFilter As String
    
    Dim varState As Variant
    mstrTitle = "数据包类"
    
    '1-住院医嘱;2-住院病历;3-护理病历;4-护理记录;5-首页记录;6-医嘱报告;7-疾病证明;8-知情文件
    
    '形成查病人的子ＳＱＬ语句

    If bytApplyMode > 1 Then
        '1.审查状态
        '------------------------------------------------------------------------------------------------------------------
        
        strSQL = _
            "Select B.病人id, B.主页id, B.险类,b.出院科室id" & vbNewLine & _
            "From (Select B.病人id, B.主页id, B.险类,b.出院科室id " & vbNewLine & _
            "       From 病案提交记录 C, 病案主页 B" & vbNewLine & _
            "       Where C.提交时间 Between [6] And [7] And C.病人id = B.病人id And C.主页id = B.主页id And" & vbNewLine & _
            "             B.病案状态 In ([12], [13], [14], [15])" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select B.病人id, B.主页id, B.险类,b.出院科室id " & vbNewLine & _
            "       From 病案提交记录 C, 病案主页 B" & vbNewLine & _
            "       Where C.提交时间 Between [8] And [9] And C.病人id = B.病人id And C.主页id = B.主页id And B.病案状态 = 5) B"
        
        If ParamRead(rsCondition, "出院情况") <> "" Then
            strSQL = strSQL & " Where Exists (Select 1 From 病人诊断记录 x Where x.病人id=b.病人id And x.主页id=b.主页id And x.诊断类型 In (3,13) And x.出院情况=[16])"
        End If
        
        '2.在院,出院
        '------------------------------------------------------------------------------------------------------------------
        
        If ParamRead(rsCondition, "出院情况") = "" Then
            
            strSQL = strSQL & " Union All " & vbNewLine & _
                    "Select b.病人id,b.主页id,b.险类,b.出院科室id From 病人信息 a,病案主页 b Where a.病人id=b.病人id And Nvl(b.主页ID,0)<>0 And Nvl(b.状态,0)<>1 And b.出院日期 Is Null "
                    
            If ParamRead(rsCondition, "当前病况") <> "" Then strSQL = strSQL & " And b.当前病况=[17] "
                        
            strSQL = strSQL & " Union All " & vbNewLine & _
                    "Select b.病人id,b.主页id,b.险类,b.出院科室id From 病人信息 a,病案主页 b Where a.病人id=b.病人id And Nvl(b.主页ID,0)<>0 And Nvl(b.状态,0)<>1 And b.出院日期 Between [10] And [11]"
                    
        Else
            strSQL = strSQL & " Union All " & vbNewLine & _
                    "Select b.病人id,b.主页id,b.险类,b.出院科室id From 病人信息 a,病案主页 b Where a.病人id=b.病人id And Nvl(b.主页ID,0)<>0 And Nvl(b.状态,0)<>1 And b.出院日期 Between [10] And [11] " & vbNewLine & _
                        " And Exists (Select 1 From 病人诊断记录 x Where x.病人id=b.病人id And x.主页id=b.主页id And x.诊断类型 In (3,13) And x.出院情况=[16])"
        End If
        
        strSQL = "Select b.病人id,b.主页id From (" & strSQL & ") b,Table (Cast(f_Num2List([18]) As zlTools.t_NumList)) f Where b.出院科室id=f.Column_Value "
        Select Case Val(ParamRead(rsCondition, "病人类型"))
        Case 1          '非医保病人
            strSQL = strSQL & " And b.险类 Is Null "
        Case 2          '医保病人
            strSQL = strSQL & " And b.险类 Is Not Null "
            If ParamRead(rsCondition, "医保种类") <> "" Then
                strSQL = strSQL & " And b.险类 In (" & ParamRead(rsCondition, "医保种类") & ") "
            End If
        End Select
        
        strTmp = Val(ParamRead(rsCondition, "等待接收")) & ";" & Val(ParamRead(rsCondition, "拒绝接收")) & ";" & Val(ParamRead(rsCondition, "正在审查")) & ";" & Val(ParamRead(rsCondition, "审查反馈"))
        varState = Split(strTmp, ";")
        If Val(varState(0)) = 1 Then varState(0) = 1
        If Val(varState(1)) = 1 Then varState(1) = 2
        If Val(varState(2)) = 1 Then varState(2) = 3
        If Val(varState(3)) = 1 Then varState(3) = 4
    
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    strTmp = "Decode(a.反馈对象,1,'住院医嘱',2,'住院病历',3,'护理病历',4,'护理记录',5,'首页记录',6,'医嘱报告',7,'疾病证明',8,'知情文件',9,'临床路径') As 反馈对象"
    
    '按次数查找在院审查zq
    If lngCur次数 = 0 Then
        If str反馈人 = "" Then
            strFilter = "A.反馈时间 BetWeen [20] And [21] And "
        Else
            str反馈人 = GetFeedback(str反馈人)
            strFilter = "A.反馈时间 BetWeen [20] And [21] And " & str反馈人
        End If
    ElseIf lngCur次数 = 1 Then
         If str反馈人 = "" Then
            strFilter = " (A.反馈次数 is null or A.反馈次数 =[19]) And A.反馈时间 BetWeen [20] And [21] And "
        Else
            str反馈人 = GetFeedback(str反馈人)
            strFilter = " (A.反馈次数 is null or A.反馈次数 =[19]) And A.反馈时间 BetWeen [20] And [21] And " & str反馈人
        End If
    Else
        If str反馈人 = "" Then
            strFilter = " A.反馈次数 =[19] And A.反馈时间 BetWeen [20] And [21] And "
        Else
            str反馈人 = GetFeedback(str反馈人)
            strFilter = " A.反馈次数 =[19] And A.反馈时间 BetWeen [20] And [21] And " & str反馈人
        End If
    End If
    
    Select Case bytApplyMode
    '------------------------------------------------------------------------------------------------------------------
    Case 1, 0                     '指定

        mstrSQL = _
            "Select Decode(a.反馈对象,1,'object_advice',2,'object_case',3,'object_case',4,'object_tend',5,'object_first',6,'object_report','object_file') As 图标,a.ID, a.相关id, a.提交id, a.病人id, a.主页id,a.反馈对象 As 反馈对象id," & strTmp & ", a.记录性质, a.记录状态, a.反馈意见, a.反馈项目id, a.反馈人, a.反馈时间, a.处理期限,a.评分级别," & vbNewLine & _
            "       a.处理说明, a.处理人, a.处理时间,Decode(a.反馈对象,4,b.名称,c.病历名称) As 文件名称,a.文件id,a.医嘱id,a.科室id,a.分制,a.分值,a.补充说明,a.反馈次数,a.反馈记录,e.姓名,f.名称 As 科室 " & vbNewLine & _
            "From 病案反馈记录 A,病历文件列表 b,电子病历记录 c,病案主页 d,病人信息 e,部门表 f " & vbNewLine & _
            "Where A.ID = [1] And a.文件id=b.ID(+) And a.文件id=c.ID(+) And a.病人id=d.病人id And a.主页id=d.主页id And e.病人id=d.病人id And f.ID=d.出院科室id"
    '------------------------------------------------------------------------------------------------------------------
    Case 2                      '等待修改
        
        mstrSQL = _
            "Select Decode(a.反馈对象,1,'object_advice',2,'object_case',3,'object_case',4,'object_tend',5,'object_first',6,'object_report','object_file') As 图标,A.ID, A.相关id, A.提交id, A.病人id, A.主页id,a.反馈对象 As 反馈对象id,a.文件id,a.医嘱id,a.科室id," & strTmp & ", A.反馈意见,a.分制,a.分值,a.补充说明,a.反馈记录,a.反馈次数, C.姓名,D.名称 As 科室" & vbNewLine & _
            "From 病案反馈记录 A, 病案主页 B, 病人信息 C,部门表 D" & vbNewLine & _
            "Where " & strFilter & "A.记录状态 = 1 And A.病人id = B.病人id And A.主页id = B.主页id And C.病人id = A.病人id And d.ID=b.出院科室id"
'        strFilter
        
        
        If lng病人ID > -1 Then
            mstrSQL = mstrSQL & " And a.病人id=[4] And a.主页id=[5]"
        Else
            mstrSQL = mstrSQL & " And (a.病人id,a.主页id) In (" & strSQL & ")"
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 3                      '等待复查

        mstrSQL = _
            "Select Decode(a.反馈对象,1,'object_advice',2,'object_case',3,'object_case',4,'object_tend',5,'object_first',6,'object_report','object_file') As 图标,A.ID, A.相关id, A.提交id, A.病人id, A.主页id,a.反馈对象 As 反馈对象id,a.文件id,a.医嘱id,a.科室id," & strTmp & ", A.反馈意见,a.分制,a.分值,a.补充说明,a.反馈记录,a.反馈次数, C.姓名,D.名称 As 科室" & vbNewLine & _
            "From 病案反馈记录 A, 病案主页 B, 病人信息 C,部门表 D" & vbNewLine & _
            "Where A.记录状态 = 2 And A.病人id = B.病人id And A.主页id = B.主页id And C.病人id = A.病人id And D.ID=B.出院科室id"

        If lng病人ID > -1 Then
            mstrSQL = mstrSQL & " And a.病人id=[4] And a.主页id=[5]"
        Else
            mstrSQL = mstrSQL & " And (a.病人id,a.主页id) In (" & strSQL & ")"
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 4                      '结束问题
        
        If lng病人ID > -1 Then
            mstrSQL = _
                "Select Decode(a.反馈对象,1,'object_advice',2,'object_case',3,'object_case',4,'object_tend',5,'object_first',6,'object_report','object_file') As 图标,A.ID, A.相关id, A.提交id, A.病人id, A.主页id,a.反馈对象 As 反馈对象id,a.文件id,a.医嘱id,a.科室id," & strTmp & ", A.反馈意见,a.分制,a.分值,a.补充说明,a.反馈次数,a.反馈记录, C.姓名,D.名称 As 科室" & vbNewLine & _
                "From (Select A.ID, A.相关id, A.提交id, A.病人id, A.主页id, A.反馈对象,a.文件id,a.医嘱id,a.科室id, A.反馈意见,a.分制,a.分值,a.补充说明,a.反馈次数,a.反馈记录" & vbNewLine & _
                "       From 病案反馈记录 A" & vbNewLine & _
                "       Where a.反馈时间 Between [2] And [3] And a.记录状态 = 3" & IIf(lng病人ID > -1, " And a.病人id=[4] And a.主页id=[5] ", "") & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select A.ID, A.相关id, A.提交id, A.病人id, A.主页id, A.反馈对象,a.文件id,a.医嘱id,a.科室id, A.反馈意见,a.分制,a.分值,a.补充说明,a.反馈次数,a.反馈记录" & vbNewLine & _
                "       From 病案反馈历史 A" & vbNewLine & _
                "       Where a.反馈时间 Between [2] And [3] " & IIf(lng病人ID > -1, " And a.病人id=[4] And a.主页id=[5] ", "") & ") A, 病案主页 B, 病人信息 C,部门表 D" & vbNewLine & _
                "Where A.病人id = B.病人id And A.主页id = B.主页id And C.病人id = A.病人id And D.ID=B.出院科室id"
        Else
            mstrSQL = _
                "Select Decode(a.反馈对象,1,'object_advice',2,'object_case',3,'object_case',4,'object_tend',5,'object_first',6,'object_report','object_file') As 图标,A.ID, A.相关id, A.提交id, A.病人id, A.主页id,a.反馈对象 As 反馈对象id,a.文件id,a.医嘱id,a.科室id," & strTmp & ", A.反馈意见,a.分制,a.分值,a.补充说明,a.反馈次数,a.反馈记录, C.姓名,D.名称 As 科室" & vbNewLine & _
                "From (Select A.ID, A.相关id, A.提交id, A.病人id, A.主页id, A.反馈对象,a.文件id,a.医嘱id,a.科室id, A.反馈意见,a.分制,a.分值,a.补充说明,a.反馈次数,a.反馈记录" & vbNewLine & _
                "       From 病案反馈记录 A" & vbNewLine & _
                "       Where a.反馈时间 Between [2] And [3] And a.记录状态 = 3" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select A.ID, A.相关id, A.提交id, A.病人id, A.主页id, A.反馈对象,a.文件id,a.医嘱id,a.科室id, A.反馈意见,a.分制,a.分值,a.补充说明,a.反馈次数,a.反馈记录" & vbNewLine & _
                "       From 病案反馈历史 A" & vbNewLine & _
                "       Where a.反馈时间 Between [2] And [3]) A, 病案主页 B, 病人信息 C,部门表 D" & vbNewLine & _
                "Where A.病人id = B.病人id And A.主页id = B.主页id And C.病人id = A.病人id And D.ID=B.出院科室id"
            
            mstrSQL = "Select * From (" & mstrSQL & ") A Where (a.病人id,a.主页id) In (" & strSQL & ")"
        End If
        
    End Select
    
    On Error GoTo errHand
    '------------------------------------------------------------------------------------------------------------------
    
    If bytApplyMode > 1 Then
        If strStart = "" Then
            Set GetQuestion = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey, CDate(Now), CDate(Now), lng病人ID, lng主页ID, CDate(ParamRead(rsCondition, "审查开始时间")), CDate(ParamRead(rsCondition, "审查结束时间")), CDate(ParamRead(rsCondition, "归档开始时间")), CDate(ParamRead(rsCondition, "归档结束时间")), CDate(ParamRead(rsCondition, "出院开始时间")), CDate(ParamRead(rsCondition, "出院结束时间")), Val(varState(0)), Val(varState(1)), Val(varState(2)), Val(varState(3)), ParamRead(rsCondition, "出院情况"), ParamRead(rsCondition, "当前病况"), strDept, lngCur次数, CDate(str开始时间), CDate(str结束时间)) ', UCase(str反馈人)
        Else
            Set GetQuestion = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey, CDate(strStart), CDate(strEnd), lng病人ID, lng主页ID, CDate(ParamRead(rsCondition, "审查开始时间")), CDate(ParamRead(rsCondition, "审查结束时间")), CDate(ParamRead(rsCondition, "归档开始时间")), CDate(ParamRead(rsCondition, "归档结束时间")), CDate(ParamRead(rsCondition, "出院开始时间")), CDate(ParamRead(rsCondition, "出院结束时间")), Val(varState(0)), Val(varState(1)), Val(varState(2)), Val(varState(3)), ParamRead(rsCondition, "出院情况"), ParamRead(rsCondition, "当前病况"), strDept, lngCur次数, CDate(str开始时间), CDate(str结束时间))
        End If
    Else
        If strStart = "" Then
            Set GetQuestion = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey, CDate(Now), CDate(Now), lng病人ID, lng主页ID)
        Else
            Set GetQuestion = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey, CDate(strStart), CDate(strEnd), lng病人ID, lng主页ID)
        End If
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetProjectUse(ByVal lng方案ID As Long) As ADODB.Recordset
'功能:检查该方案是否在病案反馈记录中已经被使用过
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrH
    
    strSQL = "Select Count(A.ID) as 条数 From 病案审查目录 A,病案审查分类 B,病案审查方案 C,病案反馈记录 D" & vbNewLine & _
            "Where A.分类ID= B.ID And B.方案ID = C.ID And A.id =D.反馈项目ID And C.ID=[1] And Rownum >0"
    Set GetProjectUse = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng方案ID)
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetItemUse(ByVal lng项目ID As Long) As ADODB.Recordset
'功能:检查该项目是否在病案反馈记录中已经被使用过
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrH
    
    strSQL = "Select Count(*) as 条数 From 病案反馈记录 where 反馈项目ID=[1]"
    Set GetItemUse = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng项目ID)
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRelevanceID(ByVal lng关联ID As Long) As ADODB.Recordset
'功能:检查该反馈记录中是否有相关联的ID,获取最主要反馈记录。
    Dim strSQL As String
    On Error GoTo ErrH
    
    strSQL = "Select 相关ID From 病案反馈记录 Where ID =[1]"
    Set GetRelevanceID = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng关联ID)
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetFeedback(ByVal str反馈人 As String) As String
'功能:反馈人查询串
    Dim strTemp() As String
    Dim lngRow As Long
    Dim strFeedback As String
    Dim strTempSql As String
    If InStrRev(str反馈人, ",", -1) Then
        strTemp = Split(str反馈人, ",")
        strFeedback = ""
        For lngRow = 0 To UBound(strTemp)
            strFeedback = strFeedback & "'" & strTemp(lngRow) & "'" & ","
        Next
    End If
    
    If strFeedback <> "" Then
        If Right(strFeedback, 1) = "," Then
              strTempSql = " A.反馈人 in (" & Left(strFeedback, Len(strFeedback) - 1) & ") And "
              GetFeedback = strTempSql
        End If
    Else
        If str反馈人 <> "" Then
            strTempSql = " A.反馈人 = '" & str反馈人 & "' And "
            GetFeedback = strTempSql
        End If
    End If
End Function

Public Function GetExamineStartUse() As Boolean
'功能:检查是否有启动了的审查方案
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrH
    
    strSQL = "Select ID From 病案审查方案 Where 启用时间 is not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic")
    If rsTmp.RecordCount > 0 Then
        GetExamineStartUse = True
    Else
        GetExamineStartUse = False
    End If
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetHavePath(ByVal lng部门ID As Long) As Boolean
'功能：检查指定科室或病区是否有可用的临床路径
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrH
    
    strSQL = "Select a.Id" & vbNewLine & _
            "From 临床路径目录 A, 临床路径版本 B, 临床路径科室 C," & vbNewLine & _
            "     (Select 科室id From 病区科室对应 Where 病区id = [1]" & vbNewLine & _
            "       Union" & vbNewLine & _
            "       Select ID From 部门表 Where ID = [1]) D" & vbNewLine & _
            "Where a.Id = b.路径id And a.最新版本 = b.版本号 And a.Id = c.路径id(+) And (c.科室id = d.科室id or c.科室id is null) And Rownum < 2"
    On Error GoTo ErrH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng部门ID)
    GetHavePath = Not rsTmp.EOF
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetEPRFile(ByVal lngKey As Long, Optional ByVal lng医嘱id As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    If lng医嘱id > 0 Then
        mstrSQL = "Select  '<'||c.医嘱内容||'>' || a.病历名称  As 名称 From 电子病历记录 a,病人医嘱报告 b,病人医嘱记录 c Where a.ID=[1] And a.ID=b.病历ID And b.医嘱id=[2] And b.医嘱id=c.ID"
    Else
        mstrSQL = "Select 病历名称 As 名称 From 电子病历记录 a Where a.ID=[1]"
    End If
    
    Set GetEPRFile = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey, lng医嘱id)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetEPRFileStruct(ByVal lngKey As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    mstrSQL = "Select 名称 From 病历文件列表 a Where a.ID=[1]"

    Set GetEPRFileStruct = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetSubmitID(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Long
    '******************************************************************************************************************
    '功能:返回提交ID
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    Dim rs As ADODB.Recordset
    
    mstrSQL = "Select ID From  病案提交记录 where 病人ID =[1] and 主页ID =[2] and 接收时间 is not Null"
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lng病人ID, lng主页ID)
    If rs.RecordCount = 1 Then
        GetSubmitID = NVL(rs!ID, 0)
        Exit Function
    End If
    
    GetSubmitID = 0
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetAuditInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '功能:返回提交ID
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    
    mstrSQL = "Select 数据转出,出院科室ID From 病案主页 Where 病人ID =[1] And 主页ID = [2]"
    Set GetAuditInfo = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lng病人ID, lng主页ID)
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetCISStruct(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng科室ID As Long, ByVal blnDataMove As Boolean) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：读取电子病案的结构
    '参数：
    '返回：返回记录集
    '******************************************************************************************************************
    Dim arySerial As Variant
    Dim strTmp As String
    Dim strSerial(1 To 9) As String
    Dim intCount As Integer
    Dim strSQL As String
    Dim strSQL1 As String
    Dim rs As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim blnPath As Boolean '是否有权限显示临床路径
    
    '1-住院医嘱;2-住院病历;3-护理病历;4-护理记录;5-首页记录;6-医嘱报告;7-疾病证明;8-知情文件
    On Error GoTo errHand
    
    strTmp = Trim(zlDatabase.GetPara("档案排序顺序", glngSys, 1560, "5;1;6;2;3;4;8;7;9"))
    If strTmp = "" Then strTmp = "5;1;6;2;3;4;8;7;9"
    arySerial = Split(strTmp, ";")
    For intCount = 0 To UBound(arySerial)
        strSerial(Val(arySerial(intCount))) = intCount
    Next
    
    '病人科室存在可用的临床路径时，显示临床路径记录
    '先判断是否有"临床路径应用" 序号=1256
    If GetPrivFunc(glngSys, 1256) <> "" Then
        blnPath = GetHavePath(lng科室ID) 'mlng科室ID
    End If
    
    mstrSQL = _
        "Select *" & vbNewLine & _
        "From (Select 'R5' As ID, '' As 上级id, '首页记录' As 名称, '' As 参数,1 As 末级,'object_first' As 图标,[7] As 排序 " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R2' As ID, '' As 上级id, '住院病历' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,[4] As 排序 " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R3' As ID, '' As 上级id, '护理病历' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,[5] As 排序 " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R4' As ID, '' As 上级id, '护理记录' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,[6] As 排序 " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R1' As ID, '' As 上级id, '住院医嘱' As 名称, '' As 参数,1 As 末级,'object_advice' As 图标,[3] As 排序 " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R6' As ID, '' As 上级id, '医嘱报告' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,[8] As 排序 " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R7' As ID, '' As 上级id, '疾病证明' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,[9] As 排序 " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R8' As ID, '' As 上级id, '知情文件' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,[10] As 排序 " & vbNewLine & _
        "       From Dual " & vbNewLine & _
        IIf(blnPath, " Union All Select 'R9' As ID, '' As 上级id, '临床路径' As 名称, '' As 参数,0 As 末级,'Folder' As 图标,[11] As 排序 From Dual", "")
        
    mstrSQL = mstrSQL & " Union All" & vbNewLine & _
        "Select a.上级id || 'K' || Trim(To_Char(a.ID))|| ','||Trim(To_Char(Nvl(a.医嘱id,0)))||',0' As ID, 上级id, Decode(a.医嘱id, Null, a.名称, '<'||b.医嘱内容||'>' || a.名称)||'【创建:'||To_Char(a.创建时间,'yyyy-mm-dd hh24:mi')|| Decode(a.最后版本, 1, '书写：', '修订：') || a.保存人 || '在' || To_Char(a.保存时间, 'yyyy-mm-dd hh24:mi') || Decode(Nvl(a.签名级别, 0), 0, '保存(未完成)', 1, '完成', '审签') ||Decode(a.完成时间,Null,'】',',完成:'||To_Char(a.完成时间,'yyyy-mm-dd hh24:mi')||'】') As 名称, Trim(To_Char(a.ID))||';'||Decode(a.医嘱id,Null,'0',Trim(To_Char(a.医嘱id))) As 参数,1 As 末级,Decode(病历种类,2,'object_case',3,'object_case',7,'object_report','object_file') As 图标,排序 " & vbNewLine & _
        "From (Select ID, Decode(病历种类, 2, 'R2', 4, 'R3', 7, 'R6',6,'R8',5,'R7') As 上级id,签名级别,保存时间,保存人,最后版本, a.病历名称 As 名称,c.医嘱id,a.病历种类,a.创建时间,a.完成时间,To_Char(a.创建时间,'yyyy-mm-dd hh24:mi:ss') As 排序 " & vbNewLine & _
        "       From 电子病历记录 a,病人医嘱报告 c " & vbNewLine & _
        "       Where a.病人id = [1] And a.主页id = [2] And c.病历id(+)=a.ID And 病历种类 In (2, 3, 4, 5, 6, 7)) a,病人医嘱记录 b Where a.医嘱id=b.Id(+) "

    
    strSQL1 = "Select 1 From 病人护理记录 A Where a.病人id = [1] And a.主页id = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL1, "检查是否存在老板数据", lng病人ID, lng主页ID)
    If rsTemp.RecordCount > 0 Then
        mstrSQL = mstrSQL & " Union All" & _
            "       Select 'R4K' || Trim(To_Char(A.Id))||',0,'|| Trim(To_Char(A.科室Id)) As ID, 'R4' As 上级id," & vbNewLine & _
            "              A.名称 || '(' || B.名称 || '：' || To_Char(A.开始, 'yyyy-mm-dd hh24:mi') || ' ～ ' ||" & vbNewLine & _
            "               To_Char(A.截止, 'yyyy-mm-dd hh24:mi') || ')' As 名称, Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(保留,0)))||';'||To_Char(A.开始, 'yyyy-mm-dd hh24:mi')||' ～ '||To_Char(A.截止, 'yyyy-mm-dd hh24:mi')||';'||Trim(To_Char(A.ID)) As 参数,1 As 末级,'object_tend' As 图标,To_Char(a.开始,'yyyy-mm-dd hh24:mi:ss') As 排序" & vbNewLine & _
            "       From (Select F.ID, F.编号, F.名称, R.开始, R.截止, R.科室id, 保留" & vbNewLine & _
            "              From (Select ID, 编号, 名称, 3 As 护理级别, 通用, 0 As 科室id, 保留" & vbNewLine & _
            "                     From 病历文件列表" & vbNewLine & _
            "                     Where 种类 = 3 And 保留 < 0" & vbNewLine & _
            "                     Union All" & vbNewLine & _
            "                     Select L.ID, L.编号, L.名称, F.报表 As 护理级别, L.通用, A.科室id, L.保留" & vbNewLine & _
            "                     From 病历文件列表 L, 病历页面格式 F, 病历应用科室 A" & vbNewLine & _
            "                     Where L.种类 = 3 And L.保留 = 0 And L.种类 = F.种类 And L.编号 = F.编号 And L.ID = A.文件id(+)) F," & vbNewLine & _
            "                   (Select R.科室id, Nvl(Min(R.护理级别), 3) As 护理级别, Min(R.发生时间) As 开始, Max(R.发生时间) As 截止" & vbNewLine & _
            "                     From 病人护理记录 R" & vbNewLine & _
            "                     Where R.病人来源 = 2 And R.病人id = [1] And Nvl(R.主页id, 0) = [2] And Nvl(R.婴儿, 0) = 0" & vbNewLine & _
            "                     Group By R.科室id) R" & vbNewLine & _
            "              Where (F.通用 = 1 Or F.通用 = 2 And F.科室id = R.科室id) And F.护理级别 >= R.护理级别) A, 部门表 B" & vbNewLine & _
            "       Where A.科室id = B.ID)" & vbNewLine & _
            "Order By Decode(上级id,Null,' ',上级id),排序"
    Else
        mstrSQL = mstrSQL & " Union All" & _
                   " Select 'R4K'||Trim(To_Char(A.ID))||',0,'||Trim(To_Char(A.科室Id)) As ID,'R4' As 上级id," & vbNewLine & _
                   "     A.名称||'('||B.名称||'：'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI') || '～' ||To_Char(A.截止, 'YYYY-MM-DD HH24:MI') || ')' As 名称," & vbNewLine & _
                   "      Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(保留,0)))||';'||To_Char(A.开始, 'YYYY-MM-DD HH24:MI')||'～'||To_Char(A.截止, 'YYYY-MM-DD HH24:MI')||';'||Trim(To_Char(A.ID))||';'||Trim(To_Char(A.婴儿)) As 参数," & vbNewLine & _
                   "       1 As 末级,'object_tend' As 图标,To_Char(a.开始,'YYYY-MM-DD HH24:MI:SS') As 排序" & vbNewLine & _
                   " From (" & vbNewLine & _
                   "   Select R.ID, F.编号, R.名称,R.婴儿, R.开始, NVL(R.截止,nvl(R.时间,R.开始)) 截止, R.科室id, 保留" & vbNewLine & _
                   "   From (" & vbNewLine & _
                   "       Select L.ID, L.编号, L.名称, F.报表 As 护理级别, L.通用, L.保留" & vbNewLine & _
                   "          From 病历页面格式 F, 病历文件列表 L" & vbNewLine & _
                   "          Where L.种类 = 3 And L.种类 = F.种类 And L.编号 = F.编号 And (L.通用=1 OR L.通用=2)" & vbNewLine & _
                   "" & vbNewLine & _
                   "       ) F,(" & vbNewLine & _
                   "       Select R.ID,R.科室id,R.文件名称 名称,R.格式ID,nvl(R.婴儿,0) 婴儿,Min(R.开始时间) As 开始, Max(R.结束时间) As 截止,MAX(T.发生时间) 时间" & vbNewLine & _
                   "          From 病人护理文件 R,病人护理数据 T" & vbNewLine & _
                   "          Where R.ID=T.文件ID(+) And R.病人id = [1] And Nvl(R.主页id, 0) = [2]" & vbNewLine & _
                   "          Group By R.ID,R.文件名称,R.科室id,R.格式ID,R.婴儿" & vbNewLine & _
                   "       ) R" & vbNewLine & _
                   "       Where F.ID=R.格式ID" & vbNewLine & _
                   "   ) A, 部门表 B Where A.科室id = B.ID And DECODE(A.保留,-1,0,A.婴儿)=A.婴儿)" & vbNewLine & _
                   " Order By Decode(上级id,Null,' ',上级id),排序"
    
    End If
        
        
    If lng病人ID > 0 Then
        '只处理住院病人
        If blnDataMove Then
            mstrSQL = Replace(mstrSQL, "电子病历记录", "H电子病历记录")
            mstrSQL = Replace(mstrSQL, "病人医嘱记录", "H病人医嘱记录")
            mstrSQL = Replace(mstrSQL, "病人医嘱报告", "H病人医嘱报告")
            mstrSQL = Replace(mstrSQL, "病人护理记录", "H病人护理记录")
            mstrSQL = Replace(mstrSQL, "病人护理文件", "H病人护理文件")
            mstrSQL = Replace(mstrSQL, "病人护理数据", "H病人护理数据")
        End If
    End If
    
    Set GetCISStruct = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lng病人ID, lng主页ID, strSerial(1), strSerial(2), strSerial(3), strSerial(4), strSerial(5), strSerial(6), strSerial(7), strSerial(8), strSerial(9))
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetEMRIn_Tag(ByVal lngPatiID As Long, ByVal lngPageId As Long) As String
Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    gstrSQL = "Select Nvl(a.Id, b.Id) ID" & vbNewLine & _
                "From (Select Max(ID) ID From 病人变动记录 Where 病人id = [1] And 主页id = [2] And 开始原因 = 2 And Nvl(附加床位, 0) = 0) A," & vbNewLine & _
                "     (Select Max(ID) ID From 病人变动记录 Where 病人id = [2] And 主页id = [2] And 开始原因 = 1 And Nvl(附加床位, 0) = 0) B"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病人入院ID", lngPatiID, lngPageId)
    If rsTemp Is Nothing Then Exit Function
    If NVL(rsTemp!ID) = "" Then Exit Function
    GetEMRIn_Tag = "BD_" & rsTemp!ID
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function GetEmrCISStruct(ByVal lngPatiID As Long, ByVal lngPageId As Long) As ADODB.Recordset
Dim rsTemp As New ADODB.Recordset, strExtendTag As String, strReturn As String
    If gobjEmr Is Nothing Then Set GetEmrCISStruct = Nothing: Exit Function
    strExtendTag = GetEMRIn_Tag(lngPatiID, lngPageId)
    If strExtendTag = "" Then Set GetEmrCISStruct = Nothing: Exit Function
    
    '上级ID，ID，名称，参数，图标
    gstrSQL = "Select Decode(e.Kind, '02', 'R2', '03', 'R3', '04', 'R7', '05', 'R8', 'R2') 上级id, Nvl(d.Subdoc_Id, Rawtohex(b.Id)) As ID," & vbNewLine & _
                "       d.Subdoc_Id As 子文档id," & vbNewLine & _
                "       Nvl(d.Subdoc_Title, b.Title) || Decode(d.Completor, Null, '', '【 ' || d.Completor || ' 在' || To_Char(d.Complete_Time, 'yyyy-mm-dd hh24:mi') || '签名】') As 名称" & vbNewLine & _
                "       , Rawtohex(b.Id) || Decode(d.Subdoc_Id, Null, Null, '|' || d.Subdoc_Id) As 参数, 'object_case' As 图标" & vbNewLine & _
                "From Bz_Doc_Log B," & vbNewLine & _
                "     (Select Distinct a.Id, c.Antetype_Id, c.Subdoc_Id, c.Subdoc_Title, c.Real_Doc_Id, c.Complete_Time, c.Completor" & vbNewLine & _
                "       From Bz_Act_Log A, Bz_Doc_Tasks C" & vbNewLine & _
                "       Where a.Extend_Tag = :etag And a.Id = c.Actlog_Id And c.Real_Doc_Id Is Not Null) D, Antetype_List E" & vbNewLine & _
                "Where b.Actlog_Id = d.Id And d.Real_Doc_Id = b.Id And d.Antetype_Id = e.Id And Decode(d.Subdoc_Id, Null, d.Antetype_Id, b.Antetype_Id) = b.Antetype_Id " & vbNewLine & _
                "Order By e.Code, b.Creat_Time,d.Complete_Time"
    strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, strExtendTag & "^16^etag", rsTemp)
    If strReturn <> "" Then
        MsgBox strReturn, vbCritical, gstrSysName
        Set GetEmrCISStruct = Nothing: Exit Function
    End If
    
    Set GetEmrCISStruct = rsTemp
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function GetEMRFile(ByVal strKey As String) As ADODB.Recordset
Dim strDocid As String, strSubDocid As String, strReturn As String, rsTemp As ADODB.Recordset
    On Error GoTo errHand
    If gobjEmr Is Nothing Then Exit Function
    If InStr(strKey, "|") = 0 Then
        strDocid = strKey
    Else
        strDocid = Split(strKey, "|")(0)
        strSubDocid = Split(strKey, "|")(1)
    End If
    
    gstrSQL = "Select Nvl(b.Subdoc_Title, a.Title) 名称" & vbNewLine & _
                "From Bz_Doc_Log A, Bz_Doc_Tasks B" & vbNewLine & _
                "Where a.Id = Hextoraw(:docid) And a.Id = b.Real_Doc_Id" & IIf(strSubDocid = "", "", " And b.Subdoc_Id =:subdocid")
    strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, strDocid & "^16^docid" & IIf(strSubDocid = "", "", "|" & strSubDocid & "^16^subdocid"), rsTemp)
    If strReturn <> "" Then
        MsgBox strReturn, vbCritical, gstrSysName
        Set GetEMRFile = Nothing: Exit Function
    End If
    
    Set GetEMRFile = rsTemp
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

