VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClinicWorkTimeOther 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "时间段辅助设置"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd取消 
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   2460
      TabIndex        =   6
      Top             =   4590
      Width           =   1100
   End
   Begin VB.OptionButton opt时间 
      Caption         =   "分段时间间隔"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Top             =   540
      Width           =   1380
   End
   Begin VB.OptionButton opt时间 
      Caption         =   "平行时间间隔"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   180
      Value           =   -1  'True
      Width           =   1395
   End
   Begin VB.TextBox txtTimeOut 
      Height          =   300
      Left            =   1500
      TabIndex        =   1
      Text            =   "10"
      Top             =   135
      Width           =   450
   End
   Begin VB.CommandButton cmdRecalc 
      Caption         =   "辅助计算(&F)"
      Height          =   350
      Left            =   990
      TabIndex        =   0
      ToolTipText     =   "点击重新计算时段"
      Top             =   4590
      Width           =   1260
   End
   Begin MSComCtl2.UpDown udTime 
      Height          =   300
      Left            =   1980
      TabIndex        =   2
      Top             =   135
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtTimeOut"
      BuddyDispid     =   196611
      OrigLeft        =   2025
      OrigTop         =   105
      OrigRight       =   2280
      OrigBottom      =   450
      Max             =   1440
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VSFlex8Ctl.VSFlexGrid vsTime 
      Height          =   3600
      Left            =   30
      TabIndex        =   7
      Top             =   840
      Width           =   3555
      _cx             =   6271
      _cy             =   6350
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClinicWorkTimeOther.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lbl分 
      AutoSize        =   -1  'True
      Caption         =   "分"
      Height          =   180
      Left            =   2235
      TabIndex        =   5
      Top             =   195
      Width           =   180
   End
End
Attribute VB_Name = "frmClinicWorkTimeOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type TYTime
    上班时间 As String
    下班时间 As String
End Type
Private Enum m时段类型
    m平行时间间隔 = 0
    m分段时间间隔 = 1
End Enum
Private Type TyWorkTime
    上午 As TYTime
    下午 As TYTime
End Type
Private mWorkTime As TyWorkTime
Private mint缺省间隔 As Integer
Private mdtStartTime As Date
Private mdtEndTime As Date
Private mstr休息时段 As String
Private mblnOk As Boolean

Private mvarTimes As Variant
'VarTimes
'       "时间间隔":Array("时间间隔",10)
'       "分段间隔":Array("分段间隔",时间1(如:8:00～9:00),2;时间2,间隔2;....)
Public Function ShowMe(ByVal frmMain As Object, ByVal int缺省间隔 As Integer, _
    ByVal datStartTime As Date, ByVal datEndTime As Date, ByVal str休息时段 As String, ByRef varTimes As Variant) As Boolean
    '功能:程序入口
    '入参：
    '   int缺省间隔 - 缺省的按"平行时间间隔"的间隔分钟数
    '   datStartTime - 开始时间
    '   datEndTime - 终止时间
    '   str休息时段 - 格式（开始时间1～终止时间1; 开始时间2～终止时间2;….）
    '                 开始时间和终止时间格式为: HH24:MM.比如：12:00～14:00;17:30～18:00
    mint缺省间隔 = int缺省间隔
    mdtStartTime = datStartTime: mdtEndTime = datEndTime: mstr休息时段 = str休息时段
    mblnOk = False: mvarTimes = Empty
    
    Err = 0: On Error Resume Next
    Me.Show 1, frmMain
    varTimes = mvarTimes
    ShowMe = mblnOk
End Function

Private Sub cmdRecalc_Click()
    Dim strTemp As String, i As Integer

    Err = 0: On Error GoTo errHandler
    If Val(txtTimeOut.Text) > 60 Or Val(txtTimeOut.Text) < 0 Then
        MsgBox "平行时间间隔不能大于60分钟或小于0分钟！", vbInformation, gstrSysName
        txtTimeOut.Text = mint缺省间隔
        Exit Sub
    End If
    
    If opt时间(m平行时间间隔).Value Then
        mvarTimes = Array("时间间隔", Val(txtTimeOut.Text))
    Else
        With vsTime
            strTemp = ""
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("时间刻度"))) <> "" _
                    And Val(.TextMatrix(i, .ColIndex("时间间隔"))) > 0 Then
                    strTemp = strTemp & ";" & .Cell(flexcpData, i, .ColIndex("时间刻度"))
                    strTemp = strTemp & "," & Val(.TextMatrix(i, .ColIndex("时间间隔")))
                End If
            Next
        End With
        If strTemp = "" Then
            MsgBox "未设置时间间隔，请检查！", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        Else
            strTemp = Mid(strTemp, 2)
            mvarTimes = Array("分段间隔", strTemp)
        End If
    End If
    mblnOk = True: Unload Me
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo errHandler
    Call InitData
    opt时间(m分段时间间隔).Value = True
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    cmdRecalc.Top = ScaleHeight - cmdRecalc.Height - 100
    cmd取消.Top = cmdRecalc.Top
    cmd取消.Left = ScaleWidth - cmd取消.Width - 50
    
    cmdRecalc.Left = cmd取消.Left - cmdRecalc.Width - 50
    With vsTime
        .Left = Me.ScaleLeft
        .Height = cmdRecalc.Top - .Top - 50
        .Width = Me.ScaleWidth
    End With
End Sub

Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化信息
    '编制:刘兴洪
    '日期:2012-07-10 17:25:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, varData As Variant
    Dim dtCurStart As Date, dtCurEnd As Date, lngRow As Integer
    Dim varTimes As Variant, dtStart As Date, dtEnd As Date
    Dim i As Integer
    
    On Error GoTo errHandler
    With vsTime
        .Clear 1
        .Rows = 1
        .Editable = flexEDKbdMouse
        lngRow = 1
        dtCurStart = mdtStartTime
        If mstr休息时段 = "" Then
            dtCurEnd = DateAdd("h", 1, dtCurStart)
            Do While DateDiff("n", dtCurEnd, mdtEndTime) >= 0
                If DateDiff("n", dtCurEnd, mdtEndTime) < 0 Then dtCurEnd = mdtEndTime
                .Rows = .Rows + 1
                .TextMatrix(lngRow, .ColIndex("时间刻度")) = Format(dtCurStart, "HH:MM") & "～" & Format(dtCurEnd, "HH:MM")
                .Cell(flexcpData, lngRow, .ColIndex("时间刻度")) = dtCurStart & "～" & dtCurEnd
                .TextMatrix(lngRow, .ColIndex("时间间隔")) = txtTimeOut.Text
                lngRow = lngRow + 1
                dtCurStart = dtCurEnd: dtCurEnd = DateAdd("h", 1, dtCurStart)
            Loop
            If DateDiff("n", dtCurStart, mdtEndTime) > 0 Then
                .Rows = .Rows + 1
                .TextMatrix(lngRow, .ColIndex("时间刻度")) = Format(dtCurStart, "HH:MM") & "～" & Format(mdtEndTime, "HH:MM")
                .Cell(flexcpData, lngRow, .ColIndex("时间刻度")) = dtCurStart & "～" & mdtEndTime
                .TextMatrix(lngRow, .ColIndex("时间间隔")) = txtTimeOut.Text
            End If
        Else
            varTimes = Split(mstr休息时段, ";")
            For i = 0 To UBound(varTimes)
                '如果休息时段的开始时间小于当前时段的开始时间，则表示是第二天，休息时段的开始时间和终止时间都要加一天
                dtStart = CDate(Format(mdtStartTime, "yyyy-mm-dd ") & Split(varTimes(i), "-")(0))
                dtEnd = CDate(Format(mdtStartTime, "yyyy-mm-dd ") & Split(varTimes(i), "-")(1))
                If DateDiff("n", dtStart, dtCurStart) > 0 Then dtStart = DateAdd("d", 1, dtStart): dtEnd = DateAdd("d", 1, dtEnd)
                '休息时段的终止时间小于休息时段的开始时间，则休息时段的终止时间加一天
                If DateDiff("n", dtEnd, dtStart) > 0 Then dtEnd = DateAdd("d", 1, dtEnd)
                
                dtCurEnd = DateAdd("h", 1, dtCurStart)
                Do While DateDiff("n", dtCurEnd, dtStart) >= 0
                    If DateDiff("n", dtCurEnd, dtStart) < 0 Then dtCurEnd = dtStart
                    .Rows = .Rows + 1
                    .TextMatrix(lngRow, .ColIndex("时间刻度")) = Format(dtCurStart, "HH:MM") & "～" & Format(dtCurEnd, "HH:MM")
                    .Cell(flexcpData, lngRow, .ColIndex("时间刻度")) = dtCurStart & "～" & dtCurEnd
                    .TextMatrix(lngRow, .ColIndex("时间间隔")) = txtTimeOut.Text
                    lngRow = lngRow + 1
                    dtCurStart = dtCurEnd: dtCurEnd = DateAdd("h", 1, dtCurStart)
                Loop
                If DateDiff("n", dtCurStart, dtStart) > 0 Then
                    .Rows = .Rows + 1
                    .TextMatrix(lngRow, .ColIndex("时间刻度")) = Format(dtCurStart, "HH:MM") & "～" & Format(dtStart, "HH:MM")
                    .Cell(flexcpData, lngRow, .ColIndex("时间刻度")) = dtCurStart & "～" & dtStart
                    .TextMatrix(lngRow, .ColIndex("时间间隔")) = txtTimeOut.Text
                End If
                dtCurStart = dtEnd
            Next
            dtStart = mdtEndTime
            Do While DateDiff("n", dtCurEnd, mdtEndTime) >= 0
                If DateDiff("n", dtCurEnd, mdtEndTime) < 0 Then dtCurEnd = mdtEndTime
                .Rows = .Rows + 1
                .TextMatrix(lngRow, .ColIndex("时间刻度")) = Format(dtCurStart, "HH:MM") & "～" & Format(dtCurEnd, "HH:MM")
                .Cell(flexcpData, lngRow, .ColIndex("时间刻度")) = dtCurStart & "～" & dtCurEnd
                .TextMatrix(lngRow, .ColIndex("时间间隔")) = txtTimeOut.Text
                lngRow = lngRow + 1
                dtCurStart = dtCurEnd: dtCurEnd = DateAdd("h", 1, dtCurStart)
            Loop
            If DateDiff("n", dtCurStart, mdtEndTime) > 0 Then
                .Rows = .Rows + 1
                .TextMatrix(lngRow, .ColIndex("时间刻度")) = Format(dtCurStart, "HH:MM") & "～" & Format(mdtEndTime, "HH:MM")
                .Cell(flexcpData, lngRow, .ColIndex("时间刻度")) = dtCurStart & "～" & mdtEndTime
                .TextMatrix(lngRow, .ColIndex("时间间隔")) = txtTimeOut.Text
            End If
        End If
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub opt时间_Click(index As Integer)
    Err = 0: On Error GoTo errHandler
   Call SetControlEnable
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsTime_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Err = 0: On Error GoTo errHandler
     With vsTime
        If Val(.Cell(flexcpText, Row, .ColIndex("时间间隔"))) > 60 Or Val(.Cell(flexcpText, Row, .ColIndex("时间间隔"))) < 0 Then
            MsgBox "时间间隔不能大于60分钟或小于0分钟！", vbInformation, gstrSysName
           .Cell(flexcpText, Row, .ColIndex("时间间隔")) = ""
        End If
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsTime_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsTime
        Select Case Col
        Case .ColIndex("时间间隔")
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub SetControlEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件是否可用
    '编制:王吉
    '日期:2012-07-11 09:52:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If opt时间(m平行时间间隔).Value = True Then
        vsTime.Enabled = False
        txtTimeOut.Enabled = True
        udTime.Enabled = True
    ElseIf opt时间(m分段时间间隔).Value = True Then
        vsTime.Enabled = True
        txtTimeOut.Enabled = False
        udTime.Enabled = False
    End If
End Sub

Private Sub vsTime_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub
