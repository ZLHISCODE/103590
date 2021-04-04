VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLoginLog 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   13860
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmLoginLog.frx":0000
   ScaleHeight     =   8715
   ScaleWidth      =   13860
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pctTip 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   480
      ScaleHeight     =   2025
      ScaleWidth      =   13875
      TabIndex        =   13
      Top             =   1920
      Width           =   13900
      Begin VB.Label lblTip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   105
      End
   End
   Begin VB.PictureBox pctFind 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   960
      ScaleHeight     =   495
      ScaleWidth      =   12615
      TabIndex        =   4
      Top             =   600
      Width           =   12615
      Begin VB.TextBox txtUser 
         Height          =   350
         Left            =   720
         TabIndex        =   7
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找"
         Height          =   350
         Left            =   11400
         TabIndex        =   6
         Top             =   117
         Width           =   1095
      End
      Begin VB.CommandButton cmdMore 
         Appearance      =   0  'Flat
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   6.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2520
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   115
         Width           =   375
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   345
         Left            =   4320
         TabIndex        =   8
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/MM/dd HH:mm"
         DateIsNull      =   -1  'True
         Format          =   127401987
         CurrentDate     =   43024
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   345
         Left            =   8520
         TabIndex        =   9
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/MM/dd HH:mm"
         DateIsNull      =   -1  'True
         Format          =   127401987
         CurrentDate     =   43024
      End
      Begin VB.Label lblUser 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "用户名"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   210
         Width           =   540
      End
      Begin VB.Label lblStart 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "登录开始时间"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3120
         TabIndex        =   11
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label lblEnd 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "登录结束时间"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   7320
         TabIndex        =   10
         Top             =   210
         Width           =   1080
      End
   End
   Begin VB.PictureBox pctLog 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3615
      ScaleWidth      =   13575
      TabIndex        =   1
      Top             =   1080
      Width           =   13575
      Begin VSFlex8Ctl.VSFlexGrid vsfLog 
         Height          =   2895
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   12015
         _cx             =   21193
         _cy             =   5106
         Appearance      =   1
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
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
         ExplorerBar     =   1
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
      Begin VB.Label lblLoad 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "点击查找加载数据."
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   3240
         Width           =   1530
      End
   End
   Begin VB.Label lblTblName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "登录日志"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   820
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "人员登录日志"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "frmLoginLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const conCol = "用户名,1200,1;人员,1200,1;部门,1200,1;客户端,1200,1;工作站,1200,1;状态,1200,1;登录时间,1200,1;退出时间,1200,1"
Private mrsUsers As ADODB.Recordset
Private mdtStart As Date
Private mdtEnd As Date

Private Sub cmdFind_Click()
        
    If gstrOracleBigVer = "" Then gstrOracleBigVer = GetVersion
        
    '开始结束时间必须同时为空或非空
    If (IsNull(dtpStart.value) And Not IsNull(dtpEnd.value)) Or (Not IsNull(dtpStart.value) And IsNull(dtpEnd.value)) Then
        MsgBox "开始/结束时间有误，请重新输入。"
        Exit Sub
    End If
    
    If IsNull(dtpStart.value) And txtUser.Text = "" Then
        MsgBox ("本次查询涉及大量数据,目前没有任何条件,可能会造成应用卡死,请至少填写一个条件"): Exit Sub
    End If
    
    MousePointer = vbArrowHourglass
    lblLoad.Caption = "数据量较大,正在加载,请耐心等候..."
    Call LoadLog(txtUser.Text, IIf(IsNull(dtpStart.value), CDate(0), dtpStart.value), IIf(IsNull(dtpEnd.value), CDate(0), dtpEnd.value))
    MousePointer = vbDefault
    lblLoad.Caption = "数据加载完成."
End Sub



Private Sub cmdMore_Click()
    Dim strUsers As String
    Dim p As POINTAPI
    Dim rsTmp As ADODB.Recordset
    Dim strTmp() As String, i As Integer
    
    p.x = (pctFind.Left + cmdMore.Left) / Screen.TwipsPerPixelX
    p.y = (pctFind.Top + cmdMore.Height + cmdMore.Top) / Screen.TwipsPerPixelY
    ClientToScreen Me.hwnd, p
    
    If mrsUsers Is Nothing Then
        Set mrsUsers = LoadUsers(True)
    End If
    
    strUsers = frmFindUser.ShowMe(Me, mrsUsers, Trim(txtUser.Text), p.x * Screen.TwipsPerPixelX, p.y * Screen.TwipsPerPixelY)
    txtUser.Text = strUsers

End Sub

Private Sub Form_Load()
    Call InitTable(vsfLog, conCol)
    vsfLog.Rows = 1
    mdtStart = Now: mdtEnd = Now
    Call SetTip
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    pctFind.Left = Me.ScaleWidth - pctFind.Width
    
    pctLog.Width = Me.ScaleWidth
    pctLog.Height = Me.ScaleHeight - pctLog.Top
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set mrsUsers = Nothing
End Sub

Private Sub pctLog_Resize()
    On Error Resume Next
    vsfLog.Width = pctLog.ScaleWidth - 240
    vsfLog.Height = pctLog.ScaleHeight - 240 - lblLoad.Height
    
    lblLoad.Top = vsfLog.Height + vsfLog.Top + 45
End Sub


Private Sub LoadLog(Optional strUser As String, Optional datStart As Date, Optional datEnd As Date)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    If gstrOracleBigVer = 11 Then
        strSQL = " (Select  a.Sessionid, a.Userid 用户名, a.Terminal 客户端, a.Userhost 工作站," & vbNewLine & _
                        "       Cast(From_Tz(b.Ntimestamp#, 'UTC') At Local As Date) 退出时间," & vbNewLine & _
                        "       Cast(From_Tz(a.Ntimestamp#, 'UTC') At Local As Date) 登录时间, Decode(b.Action#, 101, '退出', 102, '被终止', '未正常退出') 状态" & vbNewLine & _
                        "From Sys.Aud$ A," & vbNewLine & _
                        "     (Select b.Ntimestamp#, b.Action#, b.Sessionid" & vbNewLine & _
                        "       From Sys.Aud$ B" & vbNewLine & _
                        "       Where b.Action# In (101, 102)  " & vbNewLine & _
                        IIf(strUser = "", "", "And b.Userid in (Select Column_Value From Table(f_Str2list([1])))") & ") B" & vbNewLine & _
                        "Where  a.Action#=100 and a.Sessionid = b.Sessionid(+) " & vbNewLine & _
                        IIf(strUser = "", "", "And a.Userid in (Select Column_Value From Table(f_Str2list([1])))") & IIf(datStart = 0 Or IsNull(datStart), "", "And b.Ntimestamp# Between [2] And [3] ") & ") A"

    
    Else
        strSQL = "(Select a.Action#, a.Sessionid, a.Userid 用户名, a.Userhost 工作站, a.Terminal 客户端," & vbNewLine & _
                        "       Cast(From_Tz(a.Ntimestamp#, 'UTC') At Local As Date) 登录时间, a.Logoff$time 退出时间,decode(Action#,101,'退出',102,'被终止',100,'未正常退出') 状态" & vbNewLine & _
                        "From Sys.Aud$ A" & vbNewLine & _
                        "Where a.Action# In (100, 101, 102) " & vbNewLine & _
                        IIf(strUser = "", "", "And a.Userid In (Select Column_Value From Table(f_Str2list([1])))") & vbNewLine & _
                        IIf(datStart = 0 Or IsNull(datStart), "", "And a.Ntimestamp# Between [2] And [3]") & vbNewLine & _
                        ") A "

    End If
    strSQL = "Select a.用户名, b.人员, b.部门, a.客户端, a.工作站, a.状态, a.登录时间, a.退出时间" & vbNewLine & _
                    "From" & vbNewLine & _
                    strSQL & vbNewLine & _
                    "       ,(Select b.用户名, d.姓名 人员, e.名称 部门" & vbNewLine & _
                    "       From 上机人员表 B, 部门人员 C, 人员表 D, 部门表 E" & vbNewLine & _
                    "       Where b.人员id = c.人员id And c.人员id = d.Id And c.部门id = e.Id And c.缺省 = 1) B" & vbNewLine & _
                    "Where a.用户名 = b.用户名(+)" & vbNewLine & _
                    "Order By a.登录时间 Desc"

    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "LoadLog", UCase(strUser), datStart, datEnd)
    
    With vsfLog
        .Redraw = flexRDNone
        Set .DataSource = rsTmp
        .ColAlignment(-1) = flexAlignLeftCenter
        .AutoResize = True: .AutoSize 0, .Cols - 2
        .Redraw = flexRDDirect
        If .Rows > .FixedRows Then
            .Select .FixedRows, 0
        End If
        
    End With
End Sub

Private Sub dtpEnd_Change()
    If IsNull(dtpEnd.value) Then
        dtpStart.value = Null
    Else
        mdtEnd = dtpEnd.value
        dtpStart.value = mdtStart
    End If
End Sub

Private Sub dtpStart_Change()
    If IsNull(dtpStart.value) Then
        dtpEnd.value = Null
    Else
        mdtStart = dtpStart.value
        dtpEnd.value = mdtEnd
    End If
End Sub

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
End Function

Private Sub SetTip()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errh
    strSQL = "Select Value From v$parameter Where Name='audit_trail'"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "SetTiP")
    
    If rsTmp.RecordCount = 0 Then
        pctTip.Visible = True: lblLoad.Visible = False
        lblTip.Caption = "请检查数据库参数audit_trail" & vbNewLine & "参数值为None或False时,审计功能未开启,本功能无法使用"
    Else
        If UCase(rsTmp!value) = "NONE" Or UCase(rsTmp!value) = "FALSE" Then
            pctTip.Visible = True: lblLoad.Visible = False
            lblTip.Caption = "当前数据库参数audit_trail值为" & rsTmp!value & ",本功能无法使用，可选参数值有:os、db、xml、db,extended等" & vbNewLine & _
                                        "登录日志会占用大量空间，可以通过配置dbms_audit_mgmt.init_cleanup来自动定期清理日志" & _
                                        IIf(gstrOracleVer Like "11*", vbNewLine & "为防止日志信息过多占用System表空间，可以通过包dbms_audt_mgmt.set_audit_trail_location，将日志存储的表空间设置到非system表空间。", "")
        Else
            pctTip.Visible = False: lblLoad.Visible = True
            lblTip.Caption = ""
        End If
    End If
    
    Exit Sub
errh:
    MsgBox err.Description
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub
