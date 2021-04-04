VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegistPlanTimeOther 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "时间段辅助设置"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd取消 
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   2460
      TabIndex        =   7
      Top             =   4590
      Width           =   1100
   End
   Begin VB.OptionButton opt时间 
      Caption         =   "分段时间间隔"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   540
      Width           =   2850
   End
   Begin VB.OptionButton opt时间 
      Caption         =   "平行时间间隔"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   180
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox txtTimeOut 
      Height          =   300
      Left            =   1530
      TabIndex        =   2
      Text            =   "10"
      Top             =   135
      Width           =   450
   End
   Begin VB.CommandButton cmdRecalc 
      Caption         =   "辅助计算(&F)"
      Height          =   350
      Left            =   990
      TabIndex        =   1
      ToolTipText     =   "点击重新计算时段"
      Top             =   4590
      Width           =   1260
   End
   Begin VSFlex8Ctl.VSFlexGrid vsTime 
      Height          =   3600
      Left            =   60
      TabIndex        =   0
      Top             =   900
      Width           =   3705
      _cx             =   6535
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
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483634
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
      FormatString    =   $"frmRegistPlanTimeOther.frx":0000
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
   Begin MSComCtl2.UpDown udTime 
      Height          =   300
      Left            =   1980
      TabIndex        =   3
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
   Begin VB.Label lbl分 
      AutoSize        =   -1  'True
      Caption         =   "分"
      Height          =   180
      Left            =   2265
      TabIndex        =   6
      Top             =   195
      Width           =   180
   End
End
Attribute VB_Name = "frmRegistPlanTimeOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type TYTime
    上班时间 As String
    下班时间 As String
End Type
Private Type TyWorkTime
    上午 As TYTime
    下午 As TYTime
End Type
Private Enum m时段类型
    m平行时间间隔 = 0
    m分段时间间隔 = 1
End Enum
Private mWorkTime As TyWorkTime
Private mrs时段 As ADODB.Recordset
'VarTiems
'       "时间间隔"
'       "分段间隔":时间(如:8:00～9:00),2;时间2,间隔;....
Public Event zlRefreshCon(ByVal varTimes As Variant)
Public Function zlShowMe(ByVal frmMain As Object, ByVal str安排 As String, ByVal int缺省间隔 As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口(目前临时应用)
    '编制:刘兴洪
    '日期:2012-07-10 18:35:09
    '说明:
    '   21001   22001   建卡病人身份校验    IDCardCheck
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Call zlInitVar(str安排, int缺省间隔)
     Me.Show 1, frmMain
End Function
Public Sub zlInitVar(ByVal str安排 As String, ByVal int缺省间隔 As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关变量值
    '入参:str安排-时间安排,比如:上午;下午
    '        int缺省间隔-缺省的时间间隔
    '编制:刘兴洪
    '日期:2012-07-10 17:21:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str开始时间 As String, str结束时间 As String
    Dim i As Long, lngRow As Long
    Dim bln隔天 As Boolean
    Dim dtDate As Date
    Dim bln全日 As Boolean
    If mrs时段 Is Nothing Then InitData
    mrs时段.Filter = "时间段='" & str安排 & "'"
    txtTimeOut.Text = int缺省间隔
    If Not mrs时段.EOF Then
        str开始时间 = mrs时段!开始时间
        str结束时间 = mrs时段!终止时间
        If mrs时段!时间段 = "全日" Then
            bln隔天 = str开始时间 >= str结束时间
            str结束时间 = IIf(bln隔天 = False, str结束时间, mWorkTime.下午.下班时间)
        End If
    End If
    With vsTime
        .Clear 1
        .Rows = 2
        If str开始时间 = "" Then str开始时间 = mWorkTime.上午.上班时间
        If str结束时间 = "" Then str结束时间 = mWorkTime.下午.下班时间
        If str开始时间 > str结束时间 Then
            str开始时间 = "2000-01-01 " & str开始时间
            str结束时间 = "2000-01-02 " & str结束时间
        End If
        lngRow = 1
        Do While True
            dtDate = Format(CDate(str开始时间), "yyyy-mm-dd HH:00:00")
            dtDate = dtDate + 1 / 24
            If dtDate > CDate(str结束时间) Then dtDate = CDate(str结束时间)
            .TextMatrix(lngRow, .ColIndex("时间刻度")) = Format(str开始时间, "HH:MM") & "～" & Format(dtDate, "HH:MM")
            If str开始时间 < CDate(mWorkTime.上午.下班时间) Or str开始时间 >= CDate(mWorkTime.下午.上班时间) Then
                .TextMatrix(lngRow, .ColIndex("时间间隔")) = txtTimeOut.Text
            End If
            If dtDate >= str结束时间 Then Exit Do
            str开始时间 = Format(dtDate, "yyyy-mm-dd HH:MM:SS")
            .Rows = .Rows + 1
            lngRow = lngRow + 1
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub cmdRecalc_Click()
    Dim cllTime As New Collection
    Dim strTemp As String, i As Long
    cllTime.Add "", "时间间隔"
    cllTime.Add "", "分段间隔"
    If opt时间(0).Value Then
        cllTime.Remove "时间间隔"
        cllTime.Add txtTimeOut.Text, "时间间隔"
        RaiseEvent zlRefreshCon(cllTime)
        Exit Sub
    End If
    With vsTime
        strTemp = ""
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("时间刻度"))) <> "" _
                And Val(.TextMatrix(i, .ColIndex("时间间隔"))) >= 0 Then
                strTemp = strTemp & ";" & Trim(.TextMatrix(i, .ColIndex("时间刻度")))
                strTemp = strTemp & "," & Val(.TextMatrix(i, .ColIndex("时间间隔")))
            End If
        Next
    End With
    If strTemp = "" Then
        MsgBox "未设置时间间隔,请检查", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    strTemp = Mid(strTemp, 2)
    cllTime.Remove "分段间隔"
    cllTime.Add strTemp, "分段间隔"
    RaiseEvent zlRefreshCon(cllTime)
    cmd取消_Click
End Sub

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    opt时间(m分段时间间隔).Value = True
    Call InitData
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
    End With
End Sub
Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化信息
    '编制:刘兴洪
    '日期:2012-07-10 17:25:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, varData As Variant
    Dim strSQL As String
    
    On Error GoTo errHandle
    strTemp = zlDatabase.GetPara("上午上下班时间", glngSys, , "07:00:00 AND 12:00:00")
    varData = Split(UCase(strTemp & " AND "), "AND")
    If IsDate(Trim(varData(0))) Then
        mWorkTime.上午.上班时间 = Trim(varData(0))
    Else
        mWorkTime.上午.上班时间 = "07:00:00"
    End If
    If IsDate(Trim(varData(1))) Then
        mWorkTime.上午.下班时间 = Trim(varData(1))
    Else
        mWorkTime.上午.下班时间 = "12:00:00"
    End If
    strTemp = zlDatabase.GetPara("下午上下班时间", glngSys, , "14:00:00 AND 18:00:00")
    varData = Split(UCase(strTemp & " AND "), "AND")
    If IsDate(Trim(varData(0))) Then
        mWorkTime.下午.上班时间 = Trim(varData(0))
    Else
        mWorkTime.下午.上班时间 = "14:00:00"
    End If
    If IsDate(Trim(varData(1))) Then
        mWorkTime.下午.下班时间 = Trim(varData(1))
    Else
        mWorkTime.下午.下班时间 = "18:00:00"
    End If
    strSQL = "Select 时间段,to_char(开始时间,'hh24:mi:ss') as 开始时间,to_char(终止时间,'hh24:mi:ss') as 终止时间 from 时间段"
    Set mrs时段 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsTime
        .Editable = flexEDKbdMouse
    End With
    Call opt时间_Click(0)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub opt时间_Click(index As Integer)
   Call SetControlEnable
End Sub


Private Sub vsTime_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     With vsTime
        If Val(.Cell(flexcpText, Row, .ColIndex("时间间隔"))) > 60 Or Val(.Cell(flexcpText, Row, .ColIndex("时间间隔"))) < 0 Then
            MsgBox "时间间隔不能大于60分钟或小于0分钟！", vbInformation, gstrSysName
           .Cell(flexcpText, Row, .ColIndex("时间间隔")) = ""
        End If
    End With
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
    If KeyAscii > 57 Or KeyAscii < 48 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
