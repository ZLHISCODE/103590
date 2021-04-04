Attribute VB_Name = "mdlMipPollService"
'模块变量申明
'Public zlCommFun As New clsCommFun
'Public zlDataBase As New clsDatabase
'Public zlComLib As New clsComLib
'Public zlControl As New clsControl

Public gclsBusiness As New clsBusiness

Public gfrmMain As Object
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4


Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


Public Sub AddIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, ByVal strTip As String)
    '******************************************************************************************************************
    '功能：在任务栏上增加一个图标
    '参数：
    '说明：
    '******************************************************************************************************************
        
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '事件发生的载体，为了不与其它鼠标事件相冲突，所以单独放一个控件
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = strTip & Chr$(0)

    Shell_NotifyIcon NIM_ADD, t
    
End Sub

Public Sub ModifyIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, ByVal strTip As String)
    '******************************************************************************************************************
    '功能：在任务栏上增加一个图标
    '参数：
    '说明：
    '******************************************************************************************************************
        
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '事件发生的载体，为了不与其它鼠标事件相冲突，所以单独放一个控件
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = stdIcon
    t.szTip = strTip & Chr$(0)

    Shell_NotifyIcon NIM_MODIFY, t
    
End Sub


'Public Sub AddIcon(ByVal lngHwnd As Long, ByVal stdIcon As StdPicture, Optional ByVal strTip As String = "")
'
'    '功能：在任务栏上增加一个图标
'
'    Dim t As NOTIFYICONDATA
'
'    On Error Resume Next
'
'    t.cbSize = Len(t)
'    t.hwnd = lngHwnd   '事件发生的载体，为了不与其它鼠标事件相冲突，所以单独放一个控件
'    t.uId = 1&
'    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
'    t.ucallbackMessage = WM_MOUSEMOVE
'    t.hIcon = stdIcon
'    t.szTip = strTip & Chr$(0)
'
'    Shell_NotifyIcon NIM_ADD, t
'
'End Sub

Public Sub RemoveIcon(ByVal lngHwnd As Long)
    
    '功能：从任务栏上删除图标
    
    Dim t As NOTIFYICONDATA
    
    On Error Resume Next
    
    t.cbSize = Len(t)
    t.hwnd = lngHwnd   '事件发生的载体
    t.uId = 1&
    
    Shell_NotifyIcon NIM_DELETE, t
    
End Sub

Public Function GetBasePeriod(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '******************************************************************************************************************
    '功能:获取特殊时间
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim intDay As Integer
    Dim varValue As Variant
    
    If Left(strMode, 3) = "自定义" Then
        '自定义:3,4
        varValue = Split(Mid(strMode, 5), ",")
        
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", Val(varValue(0)), Now), "yyyy-MM-dd") & " 00:00:00"
        Else
            If UBound(varValue) < 1 Then
                GetBasePeriod = Format(Now, "yyyy-MM-dd") & " 23:59:59"
            Else
                GetBasePeriod = Format(DateAdd("d", Val(varValue(1)), Now), "yyyy-MM-dd") & " 23:59:59"
            End If
        End If
            
        Exit Function
    End If
    
    Select Case strMode
    Case "当  时"      '当时
        GetBasePeriod = Format(Now, "YYYY-MM-DD HH:MM:SS")
    Case "今  天"       '当天
        If bytFlag = 1 Then
            GetBasePeriod = Format(Now, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(Now, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  周"       '本周,bytFlag=1,本周开始时间,=2,本周结束时间
        intDay = Weekday(CDate(Format(Now, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 7 - intDay, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  月"       '本月
        If bytFlag = 1 Then
            GetBasePeriod = Format(Now, "YYYY-MM") & "-01 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(Now, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  季"      '本季度
        Select Case Format(Now, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetBasePeriod = Format(Now, "YYYY") & "-01-01 00:00:00"
            Else
                GetBasePeriod = Format(Now, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetBasePeriod = Format(Now, "YYYY") & "-04-01 00:00:00"
            Else
                GetBasePeriod = Format(Now, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetBasePeriod = Format(Now, "YYYY") & "-07-01 00:00:00"
            Else
                GetBasePeriod = Format(Now, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetBasePeriod = Format(Now, "YYYY") & "-10-01 00:00:00"
            Else
                GetBasePeriod = Format(Now, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "本半年"      '本半年
        If Val(Format(Now, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetBasePeriod = Format(Now, "YYYY") & "-01-01 00:00:00"
            Else
                GetBasePeriod = Format(Now, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetBasePeriod = Format(Now, "YYYY") & "-07-01 00:00:00"
            Else
                GetBasePeriod = Format(Now, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "本  年"   '全年
        If bytFlag = 1 Then
            GetBasePeriod = Format(Now, "YYYY") & "-01-01 00:00:00"
        Else
            GetBasePeriod = Format(Now, "YYYY") & "-12-31 23:59:59"
        End If
    Case "昨  天"       '昨天
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -1, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", -1, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "明  天"       '明天
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", 1, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 1, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前三天"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -3, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 0, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前一周"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -7, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 0, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前半月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -15, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 0, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前一月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -30, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 0, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前二月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -60, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 0, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前三月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -90, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 0, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    
    Case "前半年"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -180, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 0, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
        
    Case "前一年"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -365, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 0, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
        
    Case "前二年"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -365 * 2, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 0, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
        
    Case "一月前"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -365 * 50, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", -30, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
        
    Case "一周前"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -365 * 50, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", -7, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
        
    Case "一天前"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -365 * 50, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", -1, CDate(Format(Now, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    End Select
    
End Function

Public Function CommandBarExecutePublic(Control As Object, frmMain As Object, ByVal lngSys As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Select Case Control.id
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '工具栏
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            frmMain.cbsMain(lngLoop).Visible = Not frmMain.cbsMain(lngLoop).Visible
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Text      '按钮文字
    
        For lngLoop = 2 To frmMain.cbsMain.Count
            For Each objControl In frmMain.cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Size      '大图标
    
        frmMain.cbsMain.Options.LargeIcons = Not frmMain.cbsMain.Options.LargeIcons
        frmMain.cbsMain.RecalcLayout
        
    Case conMenu_View_StatusBar         '状态栏
    
        frmMain.stbThis.Visible = Not frmMain.stbThis.Visible
        frmMain.cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Help_Help              '帮助主题
    
        Call zlComLib.ShowHelp(App.ProductName, frmMain.hwnd, frmMain.Name, Int((lngSys) / 100))
        
    Case conMenu_Help_Web_Home          'Web上的中联
        
        Call zlComLib.zlHomePage(frmMain.hwnd)
        
    Case conMenu_Help_Web_Forum         'Web上的论坛
    
        Call zlComLib.zlWebForum(frmMain.hwnd)
        
    Case conMenu_Help_Web_Mail          '发送反馈
        
        Call zlComLib.zlMailTo(frmMain.hwnd)
            
    Case conMenu_Help_About             '关于
        
        Call zlComLib.ShowAbout(frmMain, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        
    Case conMenu_File_Exit             '退出
    
        Unload frmMain
        
    End Select
    
    CommandBarExecutePublic = True
End Function

Public Function CommandBarUpdatePublic(Control As Object, frmMain As Object) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Select Case Control.id
    Case conMenu_View_ToolBar_Button            '工具栏
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = frmMain.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              '图标文字
        If frmMain.cbsMain.Count >= 2 Then
            Control.Checked = Not (frmMain.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size              '大图标
        Control.Checked = frmMain.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar                 '状态栏
        Control.Checked = frmMain.stbThis.Visible
    End Select
    
    CommandBarUpdatePublic = True
End Function

Public Function SQLRecord(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    On Error GoTo errHand
    
'    Set mcolSQL = New Collection
'    mstrSplit = "'COLLECTIONFRCHENCOLLECTION'"
    
    Set rs = New ADODB.Recordset

    With rs

        .Fields.Append "SQL", adVarChar, 4000
        .Fields.Append "Trans", adTinyInt                   '1表示开始;2表示结束
        .Fields.Append "Custom", adTinyInt
        .Fields.Append "Parameter", adVarChar, 500

        .Open
    End With
    
    SQLRecord = True
    
    Exit Function
    
errHand:
    
End Function

Public Function SQLRecordAdd(ByRef rs As ADODB.Recordset, ByVal strSQL As String, Optional ByVal intTrans As Integer = 0, Optional ByVal intCustom As Integer = 0, Optional ByVal strParameter As String = "") As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim strTemp As String
    
    On Error GoTo errHand
    
'    strTemp = strSQL & mstrSplit & intTrans & mstrSplit & intCustom & mstrSplit & strParameter
'
'    mcolSQL.Add strTemp
    
    rs.AddNew
    rs("SQL").Value = strSQL
    rs("Trans").Value = intTrans
    rs("Custom").Value = intCustom
    rs("Parameter").Value = strParameter
    SQLRecordAdd = True
    
    Exit Function
    
errHand:
End Function


Public Function SQLRecordExecute(ByVal objDataOracle As clsDataOracle, ByVal rsSQL As ADODB.Recordset, Optional ByVal blnHaveTrans As Boolean = True, Optional ByRef strError As String) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim blnTran As Boolean
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim strTemp As String
    Dim aryTemp As Variant
    Dim strSQL As String
    
    On Error GoTo errHand
        
    If rsSQL.RecordCount > 0 Then

        blnTran = True

        If blnHaveTrans Then objDataOracle.BeginTrans

        rsSQL.MoveFirst

        For intLoop = 1 To rsSQL.RecordCount

            strSQL = CStr(rsSQL("SQL").Value)
            Call zlDataBase.ExecuteProcedure(strSQL, "")

            rsSQL.MoveNext
        Next

        If blnHaveTrans Then objDataOracle.CommitTrans
        blnTran = False
    End If
    
    SQLRecordExecute = True
    
    Exit Function
errHand:
    
    If blnTran And blnHaveTrans Then objDataOracle.RollbackTrans
    strError = Err.Description
'    MsgBox Err.Description
    
'    If ErrCenter = 1 Then
'        Resume
'    End If
        
End Function

