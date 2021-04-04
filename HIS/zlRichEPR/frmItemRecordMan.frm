VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmItemRecordMan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "记录频次设置"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7590
   Icon            =   "frmItemRecordMan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form15"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin ZL9BillEdit.BillEdit billTime 
      Height          =   2865
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5054
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   510
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmItemRecordMan.frx":6852
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   30
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmItemRecordMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnEdit As Boolean

Private Const conMenu_保存 = 2
Private Const conMenu_恢复 = 3
Private Const conMenu_帮助 = 4
Private Const conMenu_退出 = 5

Private Sub billTime_BeforeDeleteRow(ROW As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub billTime_EnterCell(ROW As Long, COL As Long)
    If COL < 2 Then Exit Sub
    If COL < Val(billTime.TextMatrix(ROW, 0)) + 3 Then
        billTime.ColData(COL) = 4
    Else
        billTime.ColData(COL) = 0
    End If
End Sub

Private Sub billTime_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If billTime.TxtVisible Then
        If billTime.Text = "" Then billTime.Text = " "
    End If
    
    mblnEdit = True
End Sub

Private Sub billTime_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    KeyAscii = 0
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_保存
        If Not CheckData Then Exit Sub
        If Not SaveData Then Exit Sub
        mblnEdit = False
    Case conMenu_恢复
        Call LoadData
    Case conMenu_帮助
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_退出
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    With billTime
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight
        .Height = lngBottom
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_保存
        Control.Enabled = mblnEdit
    Case conMenu_恢复
        Control.Enabled = mblnEdit
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call MainDefCommandBar
    Call LoadData
End Sub

Private Function CheckData() As Boolean
    Dim lngRow As Long, lngCount As Long
    '只要书写了名称的,就应该填写完整的时段
    
    lngCount = billTime.Rows - 1
    For lngRow = 1 To lngCount
        If billTime.TextMatrix(lngRow, 1) <> "" Then
            If Not CheckTime(lngRow, 2) Then Exit Function
            If Not CheckTime(lngRow, 3) Then Exit Function
            If Not CheckTime(lngRow, 4) Then Exit Function
            If Not CheckTime(lngRow, 5) Then Exit Function
            If Not CheckTime(lngRow, 6) Then Exit Function
            If Not CheckTime(lngRow, 7) Then Exit Function
            If Not CheckTime(lngRow, 8) Then Exit Function
        End If
    Next
    CheckData = True
End Function

Private Function CheckTime(ByVal lngRow As Long, ByVal lngCOL As Long) As Boolean
    Dim strTitle As String
    Dim strTime As String
    Dim lngHour As Long, lngMin As Long
    On Error Resume Next
    '检查时间格式合法性
    
    strTime = billTime.TextMatrix(lngRow, lngCOL)
    If strTime = "" Then
        If lngCOL <= Val(billTime.TextMatrix(lngRow, 0)) + 2 Then
            MsgBox "第" & lngRow & "行有部分数据未录入具体的时点！", vbInformation, gstrSysName
            CheckTime = False
            Exit Function
        Else
            CheckTime = True
            Exit Function
        End If
    End If
    
    strTitle = "第" & lngRow & "行第" & lngCOL & "列的"
    Err = 0
    '1、取小时
    If InStr(1, strTime, ":") = 0 Then
        lngHour = strTime
    Else
        lngHour = Split(strTime, ":")(0)
    End If
    If Err <> 0 Then
        MsgBox strTitle & "时间中含有非法字符！" & vbCrLf & _
               "时间格式为HH:mm,如05:00", vbInformation, gstrSysName
        Exit Function
    End If
    '1.1不能小于0大于23
    If lngHour < 0 Or lngHour > 23 Then
        MsgBox strTitle & "小时不能大于23或小于0！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '2、取分
    If InStr(1, strTime, ":") = 0 Then
        lngMin = "00"
    Else
        lngMin = Split(strTime, ":")(1)
    End If
    If Err <> 0 Then
        MsgBox strTitle & "时间中含有非法字符！" & vbCrLf & _
               "时间格式为HH:mm,如05:00", vbInformation, gstrSysName
        Exit Function
    End If
    '2.1不能小于0大于23
    If lngMin < 0 Or lngMin > 59 Then
        MsgBox strTitle & "分钟不能大于59或小于0！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '重新组织时间
    strTime = String(2 - Len(CStr(lngHour)), "0") & CStr(lngHour) & ":" & String(2 - Len(CStr(lngMin)), "0") & CStr(lngMin)
    billTime.TextMatrix(lngRow, lngCOL) = strTime
    
    CheckTime = True
End Function

Private Function SaveData() As Boolean
    Dim strIn As String
    Dim strSQL() As String
    Dim blnTrans As Boolean
    Dim strBegin As String, strEnd As String
    Dim lngStart As Long, lngCount As Long, lngCOL As Long
    On Error GoTo errHand
    ReDim Preserve strSQL(1 To 1)
    
    gstrSQL = "ZL_护理项目频次_DELETE"
    strSQL(ReDimArray(strSQL)) = gstrSQL
    
    lngCount = billTime.Rows - 1
    For lngStart = 1 To lngCount
        strBegin = ""
        strEnd = ""
        gstrSQL = "ZL_护理项目频次_UPDATE("
        For lngCOL = 0 To Val(billTime.TextMatrix(lngStart, 0)) - 1
            If strBegin = "" Then
                strBegin = billTime.TextMatrix(lngStart, 2 + lngCOL)
                strEnd = billTime.TextMatrix(lngStart, 3 + lngCOL)
            Else
                strBegin = Format(DateAdd("s", 60, "2010-01-01 " & strEnd & ":00"), "HH:mm")
                strEnd = billTime.TextMatrix(lngStart, 3 + lngCOL)
            End If
            
            strIn = Val(billTime.TextMatrix(lngStart, 0)) & "," & lngCOL + 1 & ",'" & strBegin & "','" & strEnd & "'," & Val(billTime.TextMatrix(lngStart, 1)) & ")"
            strIn = gstrSQL & strIn
            strSQL(ReDimArray(strSQL)) = strIn
        Next
    Next
    
    '循环执行SQL保存数据
    gcnOracle.BeginTrans
    blnTrans = True
    lngCount = UBound(strSQL)
    For lngStart = 1 To lngCount
        If strSQL(lngStart) <> "" Then
            Debug.Print strSQL(lngStart)
            Call zlDatabase.ExecuteProcedure(strSQL(lngStart), "保存护理项目频次")
        End If
    Next
    SaveData = True
    gcnOracle.CommitTrans
    Exit Function
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub LoadData()
    Dim intDo As Integer, intRow As Integer
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    mblnEdit = False
    '初始化编辑控件
    With billTime
        .ClearBill
        .Rows = 6
        .Cols = 9
        .TextMatrix(0, 0) = "频次"
        .TextMatrix(0, 1) = "取数规则"
        .TextMatrix(0, 2) = "开始时间"
        .TextMatrix(0, 3) = "分段1"
        .TextMatrix(0, 4) = "分段2"
        .TextMatrix(0, 5) = "分段3"
        .TextMatrix(0, 6) = "分段4"
        .TextMatrix(0, 7) = "分段5"
        .TextMatrix(0, 8) = "结束时间"
        .ColData(0) = 5
        .ColData(1) = 3
        .ColData(2) = 4
        .ColData(3) = 4
        .ColData(4) = 4
        .ColData(5) = 4
        .ColData(6) = 4
        .ColData(7) = 4
        .ColData(8) = 4
        .ColWidth(0) = 800
        .ColWidth(1) = 1800
        .ColWidth(2) = 900
        .ColWidth(3) = 600
        .ColWidth(4) = 600
        .ColWidth(5) = 600
        .ColWidth(6) = 600
        .ColWidth(7) = 600
        .ColWidth(8) = 900
        .PrimaryCol = 1
        .LocateCol = 1
        .ColAlignment(1) = 1
        .AllowAddRow = False
        .Active = True
        
        .AddItem "1-取第一条数据"
        .AddItem "2-取中间时点的数据"
        .AddItem "3-取最后一条数据"
        .cboStyle = DropOlnyDown
        .ListIndex = 0
        
        .TextMatrix(1, 0) = "1"
        .TextMatrix(2, 0) = "2"
        .TextMatrix(3, 0) = "3"
        .TextMatrix(4, 0) = "4"
        .TextMatrix(5, 0) = "6"
    End With
    
    '提取汇总时段数据
    strSQL = " Select 频次,序号,DECODE(类别,1,'1-取第一条数据',2,'2-取靠近中点的数据','3-取最后一条数据') AS 类别,开始,结束 " & vbNewLine & _
             " From 护理项目频次 " & vbNewLine & _
             " Order by 频次,序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取汇总时段数据")
    intRow = 0
    With rsTemp
        Do While Not .EOF
            If intRow <> IIf(!频次 = 6, 5, !频次) Then
                intDo = 1
                intRow = intRow + 1
                billTime.TextMatrix(intRow, 1) = !类别
                billTime.TextMatrix(intRow, 2) = NVL(!开始)
            End If
            
            billTime.TextMatrix(intRow, 2 + intDo) = NVL(!结束)
            
            intDo = intDo + 1
            .MoveNext
        Loop
    End With

End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup, objFile As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
    Dim lngHandel As Long

    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)
    
    'cbsMain
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.Icons = imgPublic.Icons
    
    '工具栏定义
    '-----------------------------------------------------
    cbsMain.DeleteAll
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)      '固有
    objBar.EnableDocking xtpFlagStretched
    objBar.Closeable = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_保存, "保存"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "保存数据": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_恢复, "恢复"): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "取消保存"
        Set objControl = .Add(xtpControlButton, conMenu_帮助, "帮助"): objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_退出, "退出"): objControl.Style = xtpButtonIconAndCaption
    End With
    
    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyS, conMenu_保存             '保存
    End With
End Sub
