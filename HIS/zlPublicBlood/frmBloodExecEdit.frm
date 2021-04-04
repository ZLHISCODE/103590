VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBloodExecEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "输血执行登记"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14520
   Icon            =   "frmBloodExecEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   14520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picPrompt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1440
      ScaleHeight     =   285
      ScaleWidth      =   9585
      TabIndex        =   28
      Top             =   8025
      Width           =   9585
      Begin VB.Label lblPrompt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   30
         TabIndex        =   29
         Top             =   45
         Width           =   10500
      End
   End
   Begin VB.PictureBox picinfo 
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   7020
      ScaleHeight     =   180
      ScaleWidth      =   3615
      TabIndex        =   26
      Top             =   90
      Width           =   3615
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "提醒：输注开始后请在4h内完成血液输注"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   3525
      End
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   7380
      Left            =   60
      ScaleHeight     =   7380
      ScaleWidth      =   14415
      TabIndex        =   0
      Top             =   480
      Width           =   14415
      Begin VB.CheckBox chkSign 
         Caption         =   "保存同时完成签名锁定"
         Height          =   180
         Left            =   11610
         TabIndex        =   35
         Top             =   1545
         Width           =   2115
      End
      Begin VB.TextBox txt执行摘要 
         Height          =   1215
         Left            =   180
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   6090
         Width           =   14130
      End
      Begin VB.Frame Frame1 
         Height          =   45
         Left            =   825
         TabIndex        =   11
         Top             =   90
         Width           =   13485
      End
      Begin VB.Frame Frame2 
         Height          =   45
         Left            =   825
         TabIndex        =   10
         Top             =   1620
         Width           =   13485
      End
      Begin VB.Frame fraExe 
         BorderStyle     =   0  'None
         Caption         =   "输注核对"
         Height          =   1245
         Index           =   2
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   14340
         Begin VB.Timer TimeFlash 
            Interval        =   250
            Left            =   10335
            Top             =   180
         End
         Begin VB.TextBox txtCheck 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   2
            Left            =   7860
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "2012-11-21 10:20"
            Top             =   135
            Width           =   1815
         End
         Begin VB.TextBox txtCheck 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   1
            Left            =   4455
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "管理员"
            Top             =   105
            Width           =   1800
         End
         Begin VB.TextBox txtCheck 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   225
            Index           =   0
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "管理员"
            Top             =   105
            Width           =   1815
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfCheck 
            Height          =   600
            Left            =   180
            TabIndex        =   3
            Top             =   525
            Width           =   14130
            _cx             =   24924
            _cy             =   1058
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
            BackColorSel    =   16761024
            ForeColorSel    =   16777215
            BackColorBkg    =   16777215
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   0
            HighLight       =   0
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   9
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmBloodExecEdit.frx":000C
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
         Begin VB.Image imgMore 
            Height          =   225
            Left            =   9705
            Picture         =   "frmBloodExecEdit.frx":0192
            Top             =   165
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Line linB 
            Index           =   2
            X1              =   7860
            X2              =   9675
            Y1              =   375
            Y2              =   375
         End
         Begin VB.Label lblExeTime 
            AutoSize        =   -1  'True
            Caption         =   "核对时间"
            Height          =   180
            Index           =   2
            Left            =   7065
            TabIndex        =   9
            Top             =   150
            Width           =   720
         End
         Begin VB.Line linB 
            Index           =   1
            X1              =   4410
            X2              =   6225
            Y1              =   345
            Y2              =   345
         End
         Begin VB.Line linB 
            Index           =   0
            X1              =   960
            X2              =   2775
            Y1              =   345
            Y2              =   345
         End
         Begin VB.Label lblCheck 
            AutoSize        =   -1  'True
            Caption         =   "复 查 人"
            Height          =   180
            Index           =   1
            Left            =   3660
            TabIndex        =   8
            Top             =   135
            Width           =   720
         End
         Begin VB.Label lblCheck 
            AutoSize        =   -1  'True
            Caption         =   "核 查 人"
            Height          =   180
            Index           =   0
            Left            =   165
            TabIndex        =   7
            Top             =   135
            Width           =   720
         End
      End
      Begin VB.Frame Frame3 
         Height          =   45
         Left            =   825
         TabIndex        =   1
         Top             =   5880
         Width           =   13485
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfExec 
         Height          =   3705
         Left            =   180
         TabIndex        =   13
         Top             =   1860
         Width           =   14130
         _cx             =   24924
         _cy             =   6535
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBloodExecEdit.frx":0593
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
         OwnerDraw       =   4
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
         Begin VB.PictureBox picDate 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1650
            ScaleHeight     =   270
            ScaleWidth      =   1725
            TabIndex        =   19
            Top             =   1515
            Visible         =   0   'False
            Width           =   1755
            Begin VB.CommandButton cmdDate 
               Height          =   270
               Left            =   1470
               Picture         =   "frmBloodExecEdit.frx":0735
               Style           =   1  'Graphical
               TabIndex        =   20
               TabStop         =   0   'False
               ToolTipText     =   "编辑(F4)"
               Top             =   0
               Width           =   270
            End
            Begin MSMask.MaskEdBox msk时间 
               Height          =   300
               Left            =   0
               TabIndex        =   21
               Top             =   30
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   529
               _Version        =   393216
               BorderStyle     =   0
               MaxLength       =   16
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "yyyy-MM-dd hh:mm"
               Mask            =   "####-##-## ##:##"
               PromptChar      =   "_"
            End
         End
         Begin VB.PictureBox picText 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   4170
            ScaleHeight     =   270
            ScaleWidth      =   750
            TabIndex        =   17
            Top             =   1530
            Visible         =   0   'False
            Width           =   780
            Begin VB.TextBox TxtEdit 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   270
               Left            =   0
               TabIndex        =   18
               Top             =   0
               Width           =   570
            End
         End
         Begin VB.ListBox lstSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   570
            ItemData        =   "frmBloodExecEdit.frx":082B
            Left            =   6330
            List            =   "frmBloodExecEdit.frx":0835
            TabIndex        =   16
            Top             =   1410
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.PictureBox picCbo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   4095
            ScaleHeight     =   270
            ScaleWidth      =   1440
            TabIndex        =   14
            Top             =   2070
            Visible         =   0   'False
            Width           =   1470
            Begin VB.ComboBox cboEdit 
               Height          =   300
               Left            =   -15
               TabIndex        =   15
               Text            =   "cboEdit"
               Top             =   -15
               Width           =   1500
            End
         End
      End
      Begin MSComCtl2.MonthView dtpDate 
         Height          =   2160
         Left            =   5580
         TabIndex        =   34
         Top             =   5175
         Visible         =   0   'False
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   3810
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   0
         StartOfWeek     =   266338305
         TitleBackColor  =   -2147483636
         TitleForeColor  =   -2147483634
         TrailingForeColor=   -2147483637
         CurrentDate     =   37904
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "输血核对"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   60
         TabIndex        =   24
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "输血巡视"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   60
         TabIndex        =   23
         Top             =   1545
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "执行摘要"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   60
         TabIndex        =   22
         Top             =   5790
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   8040
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBloodExecEdit.frx":0841
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23178
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picHide 
      Height          =   465
      Left            =   7695
      ScaleHeight     =   405
      ScaleWidth      =   1770
      TabIndex        =   30
      Top             =   4830
      Visible         =   0   'False
      Width           =   1830
      Begin VB.TextBox txt发送数次 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         ForeColor       =   &H80000011&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   165
         TabIndex        =   32
         Top             =   45
         Width           =   1005
      End
      Begin VB.TextBox txt本次数次 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   270
         TabIndex        =   31
         Top             =   45
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker dtp要求时间 
         Height          =   300
         Left            =   0
         TabIndex        =   33
         Top             =   -15
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   266338307
         CurrentDate     =   38082
      End
   End
   Begin XtremeCommandBars.CommandBars cbsExec 
      Left            =   0
      Top             =   -15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBloodExecEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnAcTive As Boolean
Private mstr缺省输血反应 As String, mstr输血反应 As String
Private mlngModul As Long
Private mlng收发ID As Long
Private mlng医嘱ID As Long, mlng相关ID As Long
Private mlng发送号 As Long
Private mlng科室ID As Long
Private mlng执行科室ID As Long
Private mstrPrivs As String
Private mblnOk As Boolean
Private mstr接收时间 As String '血液的接收时间
Private mint血袋数 As Integer, mint已执行血袋数 As Integer
Private mintTimerCount As Integer
Private mblnReturn As Boolean  '执行人快速输入匹配控制
Private mblnOnlyRead As Boolean '是否是只读模式
Private mblnShow As Boolean  '是否处于编辑状态
Private mintType As Integer    '编辑方式
Private mrsPersons As ADODB.Recordset  '人员信息
Private mrsItems As ADODB.Recordset  '生命体征项目
Private mblnChange As Boolean
Private mlngNoEditor As Long '后面的不能编辑起始列
Private mstr开始时间 As String
Private mblnFinish As Boolean  '执行登记自动完成医嘱执行

Public Function ShowEdit(ByVal frmParent As Object, ByVal lngModul As Enum_Inside_Program, ByVal lng医嘱ID As Long, _
    ByVal lng发送号 As Long, ByVal lng科室id As Long, ByVal lng收发ID As Long, ByVal lng执行科室ID As Long, Optional ByVal strPrivs As String, Optional ByVal blnOnlyRead As Boolean, Optional blnFinish As Boolean) As Boolean
    mblnOk = False
    mblnFinish = False
    mlngModul = lngModul
    mlng医嘱ID = lng医嘱ID
    mlng发送号 = lng发送号
    mlng科室ID = lng科室id
    mlng收发ID = lng收发ID
    mlng执行科室ID = lng执行科室ID
    mstrPrivs = strPrivs
    mblnOnlyRead = blnOnlyRead
    mblnShow = False
    On Error Resume Next
    Me.Show 1, frmParent
    If Err <> 0 Then Err.Clear
    blnFinish = mblnFinish
    ShowEdit = mblnOk
End Function

Public Sub ViewExecution(ByVal frmParent As Object, ByVal lng收发ID As Long)
    '功能查看输血执行
    mlng收发ID = lng收发ID
    mblnOnlyRead = True
    On Error Resume Next
    Me.Show 1, frmParent
    If Err <> 0 Then Err.Clear
End Sub

Private Sub cboEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picCbo_KeyDown(KeyCode, Shift)
End Sub

Private Sub cbsExec_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnOk As Boolean
    Dim strCheckOper As String, strCheckTime As String, strCheckResult As String
    Dim strSQL As String
    Dim intCol As Integer
    
    On Error GoTo ErrHand
    Select Case Control.id
        Case conMenu_Manage_ThingAudit '核对
            If txtCheck(2).Text <> "" Then
                MsgBox "该袋血液已经核对，不允许再次核对！", vbInformation, gstrSysName
                Exit Sub
            End If
            blnOk = frmUserCheck.ShowMe(Me, mlngModul, mlng科室ID, mlng科室ID, mstr接收时间, "", True, 执行核对)
            If blnOk = True Then
                strCheckOper = frmUserCheck.SendAndTakeOper
                strCheckTime = frmUserCheck.SendTime
                strCheckResult = frmUserCheck.CheckResult

                 strSQL = "Zl_血液执行记录_Check(" & mlng收发ID & ",'" & Split(strCheckOper, "'")(0) & "','" & Split(strCheckOper, "'")(1) & "',To_Date('" & strCheckTime & "','YYYY-MM-DD HH24:MI:SS'),'" & strCheckResult & "')"
                Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)

                txtCheck(0).Text = Split(strCheckOper, "'")(0)
                txtCheck(1).Text = Split(strCheckOper, "'")(1)
                txtCheck(2).Text = strCheckTime
                Call LoadCheckVsf
            End If
            mblnOk = blnOk
        Case conMenu_Manage_ThingDelAudit '取消
            strCheckOper = ""
            If txtCheck(0).Text <> UserInfo.姓名 And txtCheck(1).Text <> UserInfo.姓名 Then
                strCheckOper = gobjDatabase.UserIdentifyByUser(Me, "在取消核对前，请您先输入用户名和密码进行身份验证。", 100, mlngModul, "执行情况登记", , True)
                If strCheckOper = "" Then Exit Sub
                If txtCheck(0).Text <> strCheckOper And txtCheck(1).Text <> strCheckOper Then
                    MsgBox "只能取消自己核对或复查的血液!", vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                If MsgBox("你确定要取消核对吗？", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
            End If
            strSQL = "Zl_血液执行记录_Uncheck(" & mlng收发ID & ")"
            Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            txtCheck(0).Text = ""
            txtCheck(1).Text = ""
            txtCheck(2).Text = ""
            vsfCheck.Cell(flexcpText, vsfCheck.FixedRows, vsfCheck.FixedCols, vsfCheck.Rows - 1, vsfCheck.Cols - 1) = ""
            mblnOk = True
        Case conMenu_Edit_Clear
            If vsfExec.Row >= vsfExec.FixedRows Then
                If vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("签名人")) <> "" Then Exit Sub
                blnOk = vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("执行时间")) <> ""
                Call HiddenEditControl
                For intCol = vsfExec.FixedCols To vsfExec.Cols - 1
                    If vsfExec.ColHidden(intCol) = False Then
                        vsfExec.TextMatrix(vsfExec.Row, intCol) = ""
                    End If
                Next
                If blnOk Then
                    mblnChange = True
                    Call ChangeDataState
                End If
            End If
        Case conMenu_Tool_Sign, conMenu_Tool_SignEarse '签名锁定;取消签名锁定
            Call HiddenEditControl
            Call SignData(Control.id = conMenu_Tool_Sign)
        Case conMenu_Edit_Transf_Save
            mblnOk = SaveData
            If mblnOk = True Then
                mblnShow = False
                '保存时检查所有医嘱是否已经完成执行，如果是则自动完成医嘱执行
                If AutoAdviceFinish = True Then
                    mblnFinish = True
                    Unload Me
                Else
                    mblnFinish = False
                    Call RefreshDate
                    mblnChange = False
                End If
            End If
        Case conMenu_Edit_Transf_Cancle
            mblnShow = False
            Call RefreshDate
            mblnChange = False
        Case conMenu_File_Exit
            mblnChange = False
            Unload Me
    End Select
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsExec_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If mblnAcTive = True Then Exit Sub
    Select Case Control.id
        Case conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle
            Control.Visible = mblnOnlyRead = False
            Control.Enabled = mblnChange And Control.Visible
        Case conMenu_Edit_Clear
            Control.Visible = mblnOnlyRead = False
            Control.Enabled = vsfExec.Row >= vsfExec.FixedRows
            If Control.Enabled = True Then
                Control.Enabled = Control.Visible And vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("签名人")) = ""
            End If
        Case conMenu_Manage_ThingAudit
            Control.Visible = mblnOnlyRead = False
            Control.Enabled = txtCheck(2).Text = "" And Control.Visible
        Case conMenu_Manage_ThingDelAudit
            Control.Visible = mblnOnlyRead = False
            Control.Enabled = txtCheck(2).Text <> "" And mstr开始时间 = ""
        Case conMenu_Tool_Sign '签名锁定
            Control.Visible = mblnOnlyRead = False
            Control.Enabled = vsfExec.Row >= vsfExec.FixedRows
            If Control.Enabled = True Then
                Control.Enabled = mblnChange = False And vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("签名人")) = "" And Control.Visible
            End If
        Case conMenu_Tool_SignEarse '取消签名锁定
            Control.Visible = mblnOnlyRead = False
            Control.Enabled = vsfExec.Row >= vsfExec.FixedRows
            If Control.Enabled = True Then
                Control.Enabled = mblnChange = False And vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("签名人")) <> "" And Control.Visible
            End If
    End Select
End Sub

Private Sub cbsExec_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Bottom = Bottom + stbThis.Height
End Sub

Private Sub cbsExec_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call cbsExec.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    On Error Resume Next
    With picBack
        .Left = lngScaleLeft + 30
        .Top = lngScaleTop + 60
        .Width = lngScaleRight - .Left
        .Height = lngScaleBottom - .Top
    End With
    
    With picPrompt
        .Top = Me.ScaleHeight - stbThis.Height + 60
        .Height = stbThis.Height - 120
        .Left = stbThis.Panels(2).Left + 60
        .Width = stbThis.Panels(2).Width - 120
    End With
    With lblPrompt
        .FontSize = Me.FontSize
        .Width = picPrompt.Width
        .Height = TextHeight("刘")
        .Top = (picPrompt.Height - .Height) \ 2
    End With
End Sub

Private Sub cmdDate_Click()
    With dtpDate
        .Tag = "msk时间"
        .Left = picDate.Left + vsfExec.Left
        .Top = picDate.Top + picDate.Height + vsfExec.Top
        If IsDate(cmdDate.Tag) Then
            .Value = Format(cmdDate.Tag, "YYYY-MM-DD")
        Else
            .Value = Format(gobjDatabase.Currentdate, "YYYY-MM-DD")
        End If
        .Visible = True
        .ZOrder 0
    End With
End Sub

Private Sub dtpDate_DateClick(ByVal DateClicked As Date)
    Dim strDate As String
    If dtpDate.Tag = "msk时间" And msk时间.Visible = True Then
        If IsDate(msk时间.Text) Then
            strDate = Format(DateClicked, "YYYY-MM-DD") & " " & Mid(Format(msk时间.Text, "YYYY-MM-DD HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "YYYY-MM-DD") & " " & Mid(Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm"), 12, 5)
        End If
        msk时间.Text = Format(strDate, "YYYY-MM-DD HH:mm")
        dtpDate.Visible = False
        If picDate.Enabled And picDate.Visible Then picDate.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '以下字符做为数据分隔符或更新记录集的分隔符，因此不允许录入
    If KeyAscii = 39 Or KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyEscape And mblnShow Then
        mblnShow = False
        Call HiddenEditControl
    End If
End Sub

Private Sub HiddenEditControl()
    mintType = -1
    picDate.Visible = False
    picText.Visible = False
    lstSelect.Visible = False
    picCbo.Visible = False
    dtpDate.Visible = False
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    Dim blnUnLoad As Boolean
    
    On Error GoTo ErrHand
    mstr开始时间 = ""
    mlngNoEditor = 0
    mintTimerCount = 0
    mblnAcTive = True
    mintType = -1
    mblnChange = False
    picinfo.Visible = Not mblnOnlyRead
    TimeFlash.Enabled = Not mblnOnlyRead
    txt执行摘要.locked = mblnOnlyRead
    chkSign.Visible = Not mblnOnlyRead
    
    Call InitExecBar '菜单初始化
    If GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & Me.name & "\Form", "状态") <> "" Then
        DeleteSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & Me.name & "\Form", "状态"
    End If
    Call gobjComlib.RestoreWinState(Me, App.ProductName)
    strSQL = "Select 接收时间,接收状态,执行核对人,执行核对时间,执行复查人,执行复查时间 From 血液发送记录 where 收发ID=[1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng收发ID)
    If rsTmp.EOF Then
        MsgBox "血液还未发送,不能进行执行登记！", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    If gbln接收后才能执行 = True And mblnOnlyRead = False Then
        If Not (Val("" & rsTmp!接收状态) = 1 Or Val("" & rsTmp!接收状态) = 3) Then
            MsgBox "血液还未接收,不能进行执行登记！", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    End If
    If IsDate("" & rsTmp!接收时间) Then
        mstr接收时间 = Format("" & rsTmp!接收时间, "YYYY-MM-DD HH:mm")
    Else
        mstr接收时间 = ""
    End If
    
    '加载血液核对信息
    txtCheck(0).Text = rsTmp!执行核对人 & ""
    txtCheck(1).Text = rsTmp!执行复查人 & ""
    txtCheck(2).Text = Format(rsTmp!执行核对时间 & "", "YYYY-MM-DD HH:mm")
    Call LoadCheckVsf
    
     If Val(gobjDatabase.GetPara("保存执行登记同时签名", 2200, 9005, "0")) = 0 Or mblnOnlyRead = True Then
        chkSign.Value = 0
     Else
        chkSign.Value = 1
     End If
    '输血反应读取
    mstr输血反应 = ""
    strSQL = "Select 名称,缺省标志 From 输血反应"
    Call gobjDatabase.OpenRecordset(rsTmp, strSQL, "输血反应")
    Do While Not rsTmp.EOF
        mstr输血反应 = mstr输血反应 & "'" & rsTmp!名称
        If Val(rsTmp!缺省标志) = 1 Then mstr缺省输血反应 = "" & rsTmp!名称
    rsTmp.MoveNext
    Loop
    If Left(mstr输血反应, 1) = "'" Then mstr输血反应 = Mid(mstr输血反应, 2)
    '体温信息
    strSQL = "Select ID, 中文名, 长度, 单位, 数值域, 小数 From 诊治所见项目" & _
        " Where 分类id = 7 And 中文名 In ('体温', '脉搏', '收缩压', '舒张压', '呼吸')"
    Set mrsItems = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption)
    '人员信息
    Set mrsPersons = GetDataToPersons
    If RefreshDate(blnUnLoad) = False Then
        If blnUnLoad = True Then Unload Me: Exit Sub
    End If
    mblnAcTive = False
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function RefreshDate(Optional blnUnLoad As Boolean) As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHand
    Call HiddenEditControl
    '加载血液执行信息
    mstr开始时间 = ""
    Call LoadExecVsf
    If mblnOnlyRead = False Then
        '读取医嘱执行信息
        If mstr开始时间 = "" Then '未执行
            '获取上次执行信息
            strSQL = _
                " Select Sum(本次数次) as curNum" & _
                " From 病人医嘱执行" & _
                " Where 医嘱ID=[1] And 发送号=[2]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID, mlng发送号)
            If Not rsTmp.EOF Then
                txt发送数次.Tag = Nvl(rsTmp!curNum, 0) '血液医嘱的执行次数总和，每次执行一袋血，执行次数为1
            End If
            
            '计算本次执行应该的要求时间
            strSQL = "Select A.发送数次,Nvl(B.相关id, B.ID) 组ID, B.开始执行时间" & _
                " From 病人医嘱发送 A,病人医嘱记录 B" & _
                " Where A.医嘱ID=B.ID And A.医嘱ID=[1] And A.发送号=[2]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID, mlng发送号)
            
            dtp要求时间.Value = rsTmp!开始执行时间  '输血医嘱都为一次性执行的临嘱
            txt发送数次.Text = Val(rsTmp!发送数次 & "")
            mlng相关ID = rsTmp!组ID
            mint血袋数 = GetBloodNum
            mint已执行血袋数 = gobjComlib.FormatEx(Val(txt发送数次.Tag) * mint血袋数 / Val(txt发送数次.Text), 0) '上次执行的血袋数。已经有5位小数，用四舍五入就可以满足
            If mint已执行血袋数 >= mint血袋数 Then
                MsgBox "该医嘱本次发送允许执行 " & mint血袋数 & "袋，当前已经执行了 " & mint已执行血袋数 & " 袋，不能再执行。", vbInformation, gstrSysName
                blnUnLoad = True: Exit Function
            End If
    
            txt本次数次.Text = 1 '每次执行默认为一袋
        Else '已经执行
            '查询已执行的血液的总次数(不算本次)
            strSQL = "Select " & _
                " Sum(本次数次) as curNum" & _
                " From 病人医嘱执行" & _
                " Where 执行时间<>[3] And 医嘱ID=[1] And 发送号=[2]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID, mlng发送号, CDate(mstr开始时间))
            If Not rsTmp.EOF Then
                txt发送数次.Tag = Nvl(rsTmp!curNum, 0) '实际已执行的数次总量（不算本次）
            End If
            
            strSQL = "Select A.要求时间,Nvl(C.相关id, C.ID) 组ID,A.执行时间,A.本次数次,A.执行摘要,A.执行结果,A.执行人,B.发送数次" & _
                " From 病人医嘱执行 A,病人医嘱发送 B,病人医嘱记录 C" & _
                " Where A.医嘱ID=B.医嘱ID And A.发送号=B.发送号 And B.医嘱ID=C.ID" & _
                " And A.医嘱ID=[1] And A.发送号=[2] And A.执行时间=[3]"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID, mlng发送号, CDate(mstr开始时间))
            If rsTmp.EOF Then
                MsgBox "未能在病人医嘱执行中找到该血液的执行记录，请检查！", vbInformation, gstrSysName
                blnUnLoad = True: Exit Function
            End If
            
            dtp要求时间.Value = rsTmp!要求时间
            txt本次数次.Text = gobjComlib.FormatEx(Nvl(rsTmp!本次数次), 5)
            txt执行摘要.Text = "" & rsTmp!执行摘要
            txt发送数次.Text = Val(rsTmp!发送数次 & "")
            mlng相关ID = rsTmp!组ID
            If Trim(vsfExec.TextMatrix(vsfExec.FixedRows, vsfExec.ColIndex("执行人"))) = "" Then vsfExec.TextMatrix(vsfExec.FixedRows, vsfExec.ColIndex("执行人")) = rsTmp!执行人 & ""
            
            mint血袋数 = GetBloodNum
            mint已执行血袋数 = gobjComlib.FormatEx(Val(txt发送数次.Tag) * mint血袋数 / Val(txt发送数次.Text), 0) '本次的执行的血袋数
            txt本次数次.Text = gobjComlib.FormatEx(Val("" & rsTmp!本次数次) * mint血袋数, 0)
        End If
    Else
        strSQL = _
            " Select a.执行摘要" & vbNewLine & _
            " From 病人医嘱执行 a, 血液执行记录 b" & vbNewLine & _
            " Where a.医嘱id = b.医嘱id And a.发送号 = b.发送号 And a.执行时间 = b.执行时间 And b.记录性质 = 1 And Nvl(b.序号, 0) = 0 And b.收发id = [1]"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng收发ID)
        If Not rsTmp.EOF Then
            txt执行摘要.Text = "" & rsTmp!执行摘要
        End If
    End If
    RefreshDate = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetBloodNum() As Integer
'获取本次医嘱发送的数量
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrHand
    strSQL = "Select Count(收发id)  数量 From 血液发送记录 a, 血液配血记录 b Where a.配发id = b.Id And b.申请id = [1]"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng相关ID)
    GetBloodNum = rsTemp!数量
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub LoadCheckVsf()
'功能：加载血液核对表格信息
    Dim i As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim arrName, arrKey
    
    On Error GoTo ErrHand
    arrName = Array("血液效期", "血液质量", "输血装置", "姓名", "住院号", "病室", "床号", "血型", "血袋号", "血液种类", "血液剂量")
    arrKey = Array("血液效期是否有效", "血液质量是否完好", "输血装置是否完好", "姓名是否一致", "住院号是否一致", "病室是否一致", "床号是否一致", "血型是否一致", "血袋号是否正确", "血液种类是否正确", "血液剂量是否正确")
    With vsfCheck
        .Rows = 2
        .Cols = 12
        .FixedCols = 1
        .FixedRows = 1
        .Redraw = flexRDNone
        .ColWidth(0) = 1500
        .TextMatrix(0, 0) = "核查项(3查8对)"
        .TextMatrix(1, 0) = "核查结果"
        For i = 1 To .Cols - 1
            .TextMatrix(0, i) = CStr(arrName(i - 1))
            .ColKey(i) = CStr(arrKey(i - 1))
            .ColWidth(i) = 1125
        Next
        .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
        strSQL = " Select 血液效期是否有效, 血液质量是否完好, 输血装置是否完好, 姓名是否一致, 住院号是否一致, 病室是否一致, 床号是否一致, 血型是否一致, 血袋号是否正确, 血液种类是否正确, 血液剂量是否正确" & vbNewLine & _
            " From 血液核对结果" & vbNewLine & _
            " Where 收发id = [1] And 性质 = [2]"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng收发ID, 3)
        If rsTemp.RecordCount > 0 Then
            For i = 0 To rsTemp.Fields.Count - 1
                vsfCheck.TextMatrix(1, vsfCheck.ColIndex(rsTemp.Fields(i).name)) = IIf(Val("" & rsTemp.Fields(i).Value) = 1, "√", "")
            Next
        End If
        
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .RowHeight(-1) = 255
        .Redraw = flexRDDirect
    End With
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadExecVsf()
'功能：加载血液执行信息
    Dim i As Integer, intRow As Integer
    Dim arrName, arrKey, arrColWidth
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim int序号 As Integer, blnNULL As Boolean, intAddRow As Integer
    Dim strValue As String
    
    On Error GoTo ErrHand
    '初始化表格(编辑列)
    arrName = Array("执行时间", "执行人", "滴速", "输血部位有无渗漏", "管道冲洗", "使用药物", "输血反应", "反应时间", "体温", "脉搏", "呼吸", "血压", "记录性质", "序号", "登记人", "登记时间", "签名人", "签名时间", "状态")
    arrKey = Array("执行时间", "执行人", "滴速", "有无渗漏", "管道冲洗", "使用药物", "输血反应", "反应时间", "体温", "脉搏", "呼吸", "血压", "记录性质", "序号", "登记人", "登记时间", "签名人", "签名时间", "状态")
    arrColWidth = Array(1755, 900, 570, 900, 525, 525, 885, 1755, 600, 720, 720, 840, 0, 0, 900, 0, 900, 0, 0)
    With vsfExec
        .Clear
        .Cols = 21
        .Rows = 6
        .Redraw = flexRDNone
        .FixedRows = 1
        .FixedCols = 2
        .RowHeight(0) = 255
        .RowHeightMin = 255
        .MergeCells = flexMergeFixedOnly
        .MergeCol(0) = True
        .MergeCol(1) = True
        .FocusRect = IIf(mblnOnlyRead = True, flexFocusNone, flexFocusSolid)
        .BackColorSel = vbBlue
        .HighLight = flexHighlightNever
        .SelectionMode = flexSelectionFree
        
        .TextMatrix(1, 0) = "输注前15分钟"
        .TextMatrix(2, 0) = "输注过程"
        .TextMatrix(3, 0) = "输注过程"
        .TextMatrix(4, 0) = "输注结束"
        .TextMatrix(5, 0) = "输注结束4小时"
        
        .TextMatrix(1, 1) = "输注前15分钟"
        .TextMatrix(2, 1) = "15分钟后"
        .TextMatrix(3, 1) = "1小时"
        .TextMatrix(4, 1) = "输注结束"
        .TextMatrix(5, 1) = "输注结束4小时"
        .ColWidth(0) = 810
        .ColWidth(1) = 690
        
        For i = 2 To .Cols - 1
            Select Case CStr(arrName(i - 2))
                Case "体温"
                    .TextMatrix(0, i) = CStr(arrName(i - 2)) & vbLf & "(℃)"
                Case "脉搏"
                    .TextMatrix(0, i) = CStr(arrName(i - 2)) & vbLf & "(次/分)"
                Case "呼吸"
                    .TextMatrix(0, i) = CStr(arrName(i - 2)) & vbLf & "(次/分)"
                Case "血压"
                    .TextMatrix(0, i) = CStr(arrName(i - 2)) & vbLf & "(mmHg)"
                Case Else
                    .TextMatrix(0, i) = CStr(arrName(i - 2))
            End Select
            
            .ColKey(i) = CStr(arrKey(i - 2))
            .ColWidth(i) = Val(arrColWidth(i - 2))
        Next
        .ColHidden(.ColIndex("记录性质")) = True
        .ColHidden(.ColIndex("序号")) = True
        .ColHidden(.ColIndex("登记人")) = False
        .ColHidden(.ColIndex("登记时间")) = True
        .ColHidden(.ColIndex("签名人")) = False
        .ColHidden(.ColIndex("签名时间")) = True
        .ColHidden(.ColIndex("状态")) = True
        .FrozenCols = .ColIndex("执行时间")
        .Cell(flexcpAlignment, 0, .FixedCols, 0, .Cols - 1) = flexAlignCenterCenter
        mlngNoEditor = .ColIndex("记录性质")
        .TextMatrix(.FixedRows, .ColIndex("记录性质")) = 1: .TextMatrix(.FixedRows, .ColIndex("序号")) = 0: .TextMatrix(.FixedRows, .ColIndex("状态")) = 0
        .TextMatrix(.FixedRows + 1, .ColIndex("记录性质")) = 2: .TextMatrix(.FixedRows + 1, .ColIndex("序号")) = 0: .TextMatrix(.FixedRows + 1, .ColIndex("状态")) = 0
        .TextMatrix(.FixedRows + 2, .ColIndex("记录性质")) = 2: .TextMatrix(.FixedRows + 2, .ColIndex("序号")) = 1: .TextMatrix(.FixedRows + 2, .ColIndex("状态")) = 0
        .TextMatrix(.FixedRows + 3, .ColIndex("记录性质")) = 3: .TextMatrix(.FixedRows + 3, .ColIndex("序号")) = 0: .TextMatrix(.FixedRows + 3, .ColIndex("状态")) = 0
        .TextMatrix(.FixedRows + 4, .ColIndex("记录性质")) = 4: .TextMatrix(.FixedRows + 4, .ColIndex("序号")) = 0: .TextMatrix(.FixedRows + 4, .ColIndex("状态")) = 0
        '可能固定行的行高不正确需要自动调整下
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
        .Cell(flexcpFloodColor, .FixedRows, .FixedCols, .FixedRows, .Cols - 1) = vbBlack
        
        '提取数据
        strSQL = " Select 医嘱id, 发送号, 记录性质, 序号, 执行时间, 执行人, 执行科室id, 滴速, 输血反应, 反应时间, 输血部位是否渗漏 是否渗漏, 输血管道冲洗,是否使用药物, 体温, 脉搏, 呼吸, 收缩压, 舒张压, 摘要, 登记人," & vbNewLine & _
            "       登记时间, 签名人, 签名时间" & vbNewLine & _
            " From 血液执行记录" & vbNewLine & _
            " Where 收发id = [1] order by 记录性质,nvl(序号,0)"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng收发ID)
        Do While Not rsTmp.EOF
            Select Case Val("" & rsTmp!记录性质)
                Case 1
                    mstr开始时间 = Format("" & rsTmp!执行时间, "YYYY-MM-DD HH:mm:ss")
                    intRow = .FixedRows
                Case 2
                    intRow = .FixedRows + 1 + Val("" & rsTmp!序号)
                    If intRow > .Rows - 3 Then '除去后两行固定行
                        intAddRow = (intRow - .Rows + 3)
                        .Rows = .Rows + intAddRow
                        For i = .Rows - intAddRow To .Rows - 1
                            .TextMatrix(i, .ColIndex("记录性质")) = 2
                            .TextMatrix(i, .ColIndex("序号")) = Val("" & rsTmp!序号) - (.Rows - i - 1)
                            .RowPosition(i) = .Rows - 3 - intAddRow + 1
                        Next
                    End If
                Case 3
                    intRow = .Rows - 2
                Case 4
                    intRow = .Rows - 1
            End Select
            .TextMatrix(intRow, .ColIndex("执行时间")) = Format("" & rsTmp!执行时间, "YYYY-MM-DD HH:mm")
            .TextMatrix(intRow, .ColIndex("执行人")) = "" & rsTmp!执行人
            Select Case Val("" & rsTmp!滴速)
                Case -1
                    .TextMatrix(intRow, .ColIndex("滴速")) = "快速"
                Case -2
                    .TextMatrix(intRow, .ColIndex("滴速")) = "加压"
                Case Else
                    .TextMatrix(intRow, .ColIndex("滴速")) = "" & rsTmp!滴速
            End Select
            arrName = Array("是否渗漏", "是否使用药物", "输血管道冲洗")
            arrKey = Array("有无渗漏", "使用药物", "管道冲洗")
            For i = 0 To UBound(arrName)
                strValue = "" & rsTmp(CStr(arrName(i))).Value
                Select Case strValue
                Case "0"
                     .TextMatrix(intRow, .ColIndex(CStr(arrKey(i)))) = "无"
                Case "1"
                     .TextMatrix(intRow, .ColIndex(CStr(arrKey(i)))) = "有"
                Case Else
                    .TextMatrix(intRow, .ColIndex(CStr(arrKey(i)))) = ""
                End Select
            Next
            .TextMatrix(intRow, .ColIndex("输血反应")) = "" & rsTmp!输血反应
            .TextMatrix(intRow, .ColIndex("反应时间")) = Format("" & rsTmp!反应时间, "YYYY-MM-DD HH:mm")
            .TextMatrix(intRow, .ColIndex("体温")) = "" & rsTmp!体温
            .TextMatrix(intRow, .ColIndex("脉搏")) = "" & rsTmp!脉搏
            .TextMatrix(intRow, .ColIndex("呼吸")) = "" & rsTmp!呼吸
            .TextMatrix(intRow, .ColIndex("血压")) = "" & rsTmp!收缩压 & "/" & rsTmp!舒张压
            If .TextMatrix(intRow, .ColIndex("血压")) = "/" Then .TextMatrix(intRow, .ColIndex("血压")) = ""
            .TextMatrix(intRow, .ColIndex("记录性质")) = "" & rsTmp!记录性质
            .TextMatrix(intRow, .ColIndex("序号")) = "" & rsTmp!序号
            .TextMatrix(intRow, .ColIndex("登记人")) = "" & rsTmp!登记人
            .TextMatrix(intRow, .ColIndex("登记时间")) = Format("" & rsTmp!登记时间, "YYYY-MM-DD HH:mm:ss")
            .TextMatrix(intRow, .ColIndex("签名人")) = "" & rsTmp!签名人
            .TextMatrix(intRow, .ColIndex("签名时间")) = Format("" & rsTmp!签名时间, "YYYY-MM-DD HH:mm:ss")
            .TextMatrix(intRow, .ColIndex("状态")) = 1
            '.Cell(flexcpForeColor, intRow, .FixedCols, intRow, .Cols - 1) = IIf(.TextMatrix(intRow, .ColIndex("签名人")) <> "", vbRed, vbBlack)
        rsTmp.MoveNext
        Loop
        '输注中如果行数满了，则预留一行
        blnNULL = False
        For intRow = .FixedRows + 1 To .Rows - 3
            int序号 = Val(.TextMatrix(intRow, .ColIndex("序号")))
            If Val(.TextMatrix(intRow, .ColIndex("状态"))) = 0 Then
                blnNULL = True
                Exit For
            End If
        Next
        If blnNULL = False Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("记录性质")) = 2
            .TextMatrix(.Rows - 1, .ColIndex("序号")) = int序号 + 1
            .RowPosition(.Rows - 1) = .Rows - 3
        End If
        '重新赋值首列名称
         For intRow = .FixedRows + 1 To .Rows - 3
            .TextMatrix(intRow, 0) = "输注过程"
            int序号 = Val(.TextMatrix(intRow, .ColIndex("序号")))
            If int序号 = 0 Then
                .TextMatrix(intRow, 1) = "15分钟后"
            Else
                .TextMatrix(intRow, 1) = int序号 & "小时"
            End If
         Next
         '将非固定行的行高设置为最小行高
        For i = .FixedRows To .Rows - 1
            .RowHeight(i) = 300
            .MergeRow(i) = True
        Next
        .Cell(flexcpAlignment, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Redraw = flexRDDirect
        Call vsfExec_AfterRowColChange(0, 0, vsfExec.FixedRows, vsfExec.ColIndex("执行时间"))
    End With
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SaveData() As Boolean
'功能：保存执行记录
    Dim intRow As Integer, intNewRow As Integer, intCol As Integer
    Dim str开始执行时间 As String, str执行时间 As String
    Dim dbl本次数次 As Double, dbl剩余次数 As Double
    Dim blnTrans As Boolean, strSQL As String
    Dim arrSQL, i As Integer
    Dim int滴速 As Integer, str收缩压 As String, str舒张压 As String
    Dim arrMsg As Variant
    Dim blnDelete As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strCurDate As String
    Dim blnUpFirst As Boolean
    
    On Error GoTo ErrHand
    If mintType <> -1 Then Call MoveNextCell(False, True)
    
    blnDelete = True
    With vsfExec
        For intRow = .FixedRows To .Rows - 1
            If .TextMatrix(intRow, .ColIndex("执行时间")) <> "" Then
                blnDelete = False
                Exit For
            End If
        Next
    End With
    
    If blnDelete = False Then
        '未执行核对必须先核对
        If txtCheck(2).Text = "" Then
            MsgBox "请先进行输血前核对！", vbInformation, gstrSysName
            Exit Function
        End If
        
        With vsfExec
            str开始执行时间 = .TextMatrix(.FixedRows, .ColIndex("执行时间"))
            '输血开始执行时间不能空
            If str开始执行时间 = "" Then
                MsgBox "请填写输注前15分钟执行时间时间！", vbInformation, gstrSysName
                .Row = .FixedRows: .Col = .ColIndex("执行时间")
                .ShowCell .Row, .Col
                If .Enabled And .Visible Then .SetFocus
                Exit Function
            End If
            '输血开始时间不能小于核对时间
            If IsDate(str开始执行时间) And IsDate(txtCheck(2).Text) Then
                If CDate(Format(str开始执行时间, "YYYY-MM-DD HH:mm")) < CDate(Format(txtCheck(2).Text, "yyyy-MM-dd HH:mm")) Then
                    MsgBox "本次输注开始执行时间不能小于核对时间。", vbInformation, gstrSysName
                    .Row = .FixedRows: .Col = .ColIndex("执行时间")
                    .ShowCell .Row, .Col
                    If .Enabled And .Visible Then .SetFocus
                    Exit Function
                End If
            End If
            '输血时间不能小于医嘱执行要求时间
            If IsDate(str开始执行时间) And IsDate(dtp要求时间.Value) Then
                If CDate(Format(str开始执行时间, "YYYY-MM-DD HH:mm")) < CDate(Format(dtp要求时间.Value, "yyyy-MM-dd HH:mm")) Then
                    MsgBox "本次输注开始执行时间不能小于医嘱要求执行时间 " & Format(dtp要求时间.Value, "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
                    .Row = .FixedRows: .Col = .ColIndex("执行时间")
                    .ShowCell .Row, .Col
                    If .Enabled And .Visible Then .SetFocus
                    Exit Function
                End If
            End If
            '117041:开始时间相同，则时间相比最后一次时间自动加一秒,且本次开始时间不能小于上一次执行开始时间
            If IsDate(str开始执行时间) Then
                strSQL = "Select Max(执行时间) Lastdate" & vbNewLine & _
                        "From 病人医嘱执行" & vbNewLine & _
                        "Where 医嘱id = [1] And 发送号 = [2] And 执行时间 Between [3] And [4]" & IIf(mstr开始时间 <> "", " And 执行时间<>[5]", "")
                If mstr开始时间 <> "" Then
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID, mlng发送号, CDate(Format(str开始执行时间, "YYYY-MM-DD HH:mm")), CDate(Format(str开始执行时间, "YYYY-MM-DD HH:mm") & ":59"), CDate(mstr开始时间))
                Else
                    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID, mlng发送号, CDate(Format(str开始执行时间, "YYYY-MM-DD HH:mm")), CDate(Format(str开始执行时间, "YYYY-MM-DD HH:mm") & ":59"))
                End If
                If IsDate(rsTmp!Lastdate & "") Then
                        str开始执行时间 = Format(DateAdd("s", 1, CDate(Format(rsTmp!Lastdate, "yyyy-MM-dd HH:mm:ss"))), "yyyy-MM-dd HH:mm:ss")
                End If
            End If
        
            For intRow = .FixedRows To .Rows - 1
                '输血时间是否是正确的日期格式
                If IsDate(.TextMatrix(intRow, .ColIndex("执行时间"))) = False And .TextMatrix(intRow, .ColIndex("执行时间")) <> "" Then
                    MsgBox GetExecName(intRow) & "的执行时间不是有效的日期格式！", vbInformation, gstrSysName
                    .Row = intRow: .Col = .ColIndex("执行时间")
                    .ShowCell .Row, .Col
                    If .Enabled And .Visible Then .SetFocus
                    Exit Function
                End If
                If Val(.TextMatrix(intRow, .ColIndex("记录性质"))) = 1 Then
                    str执行时间 = str开始执行时间
                Else
                    str执行时间 = .TextMatrix(intRow, .ColIndex("执行时间"))
                End If
                
                '如果录入了输注过程每小时记录，则必须录入输注15分钟后内容
                If .TextMatrix(intRow, .ColIndex("执行时间")) <> "" Then
                    If Val(.TextMatrix(intRow, .ColIndex("记录性质"))) = 2 And Val(.TextMatrix(intRow, .ColIndex("序号"))) > 0 Then
                        If .TextMatrix(.FixedRows + 1, .ColIndex("执行时间")) = "" Then
                            MsgBox "录入了输注15分钟后每小时巡视记录，则必须录入输注15分钟后记录！", vbInformation, gstrSysName
                            .Row = .FixedRows + 1: .Col = .ColIndex("执行时间")
                            .ShowCell .Row, .Col
                            If .Enabled And .Visible Then .SetFocus
                            Exit Function
                        End If
                        
                        '检查上一小时是否录入了巡视记录
                        If Val(.TextMatrix(intRow - 1, .ColIndex("记录性质"))) = 2 Then
                            If .TextMatrix(intRow - 1, .ColIndex("执行时间")) = "" Then
                                MsgBox "输注15分钟后每小时巡视记录必须连续，请录入输注15分钟后" & Val(.TextMatrix(intRow - 1, .ColIndex("序号"))) & "小时巡视记录！", vbInformation, gstrSysName
                                .Row = intRow - 1: .Col = .ColIndex("执行时间")
                                .ShowCell .Row, .Col
                                If .Enabled And .Visible Then .SetFocus
                                Exit Function
                            End If
                        End If
                    End If
                End If
                
                '填写了输血开始时间，必须填写执行人
                If .TextMatrix(intRow, .ColIndex("执行时间")) <> "" And .TextMatrix(intRow, .ColIndex("执行人")) = "" Then
                    MsgBox GetExecName(intRow) & "的执行人不能为空！", vbInformation, gstrSysName
                    .Row = intRow: .Col = .ColIndex("执行人")
                    .ShowCell .Row, .Col
                    If .Enabled And .Visible Then .SetFocus
                    Exit Function
                End If
                '滴速的校验
                If InStr(1, ",快速,加压,,", "," & .TextMatrix(intRow, .ColIndex("滴速")) & ",") = 0 Then
                    If LenB(StrConv(.TextMatrix(intRow, .ColIndex("滴速")), vbFromUnicode)) > 3 Or Not IsNumeric(.TextMatrix(intRow, .ColIndex("滴速"))) Then
                        MsgBox GetExecName(intRow) & "的滴数只能是数字，且最多只允许录入3位数字！", vbInformation, gstrSysName
                        .Row = intRow: .Col = .ColIndex("滴速")
                        .ShowCell .Row, .Col
                        If .Enabled And .Visible Then .SetFocus
                        Exit Function
                    End If
                End If
                
                '本阶段的执行时间必须小于下一阶段的执行时间
                For intNewRow = intRow + 1 To .Rows - 1
                    If IsDate(.TextMatrix(intRow, .ColIndex("执行时间"))) And IsDate(.TextMatrix(intNewRow, .ColIndex("执行时间"))) Then
                        If CDate(Format(.TextMatrix(intRow, .ColIndex("执行时间")), "YYYY-MM-DD HH:mm")) >= CDate(Format(.TextMatrix(intNewRow, .ColIndex("执行时间")), "YYYY-MM-DD HH:mm")) Then
                            MsgBox GetExecName(intRow) & "的执行时间必须小于" & GetExecName(intNewRow) & "的执行时间！", vbInformation, gstrSysName
                            .Row = intRow: .Col = .ColIndex("执行时间")
                            .ShowCell .Row, .Col
                            If .Enabled And .Visible Then .SetFocus
                            Exit Function
                        End If
                    End If
                Next
                
                '未填写执行时间，但是录入了其他项目，则要求填写
                If .TextMatrix(intRow, .ColIndex("执行时间")) = "" Then
                    For intCol = .FixedCols To mlngNoEditor - 1
                        If .ColHidden(intCol) = False And Trim(.TextMatrix(intRow, intCol)) <> "" And intCol <> .ColIndex("执行时间") Then
                            MsgBox GetExecName(intRow) & "的执行时间为空，但填写了其他项目内容，请填写执行时间！", vbInformation, gstrSysName
                            .Row = intRow: .Col = .ColIndex("执行时间")
                            .ShowCell .Row, .Col
                            If .Enabled And .Visible Then .SetFocus
                            Exit Function
                        End If
                    Next
                End If
                
                '有输血反应，则必须录入输血反应情况和时间
                If .TextMatrix(intRow, .ColIndex("输血反应")) <> "" And .TextMatrix(intRow, .ColIndex("输血反应")) <> "无" Then
                    If .TextMatrix(intRow, .ColIndex("反应时间")) = "" Then
                        MsgBox GetExecName(intRow) & "存在输血反应时，则必须录入输入反应时间！", vbInformation, gstrSysName
                        .Row = intRow: .Col = .ColIndex("反应时间")
                        .ShowCell .Row, .Col
                        If .Enabled And .Visible Then .SetFocus
                        Exit Function
                    End If
                End If
                '输血反应时间校验
                If .TextMatrix(intRow, .ColIndex("反应时间")) <> "" Then
                    If IsDate(.TextMatrix(intRow, .ColIndex("反应时间"))) = False Then
                        MsgBox GetExecName(intRow) & "的输血反应时间不是有效的日期格式！", vbInformation, gstrSysName
                        .Row = intRow: .Col = .ColIndex("反应时间")
                        .ShowCell .Row, .Col
                        If .Enabled And .Visible Then .SetFocus
                        Exit Function
                    End If
                    '输血反应时间不能小于本次执行时间
                     If IsDate(str执行时间) Then
                        If CDate(Format(.TextMatrix(intRow, .ColIndex("反应时间")), "YYYY-MM-DD HH:mm")) < CDate(Format(str执行时间, "YYYY-MM-DD HH:mm")) Then
                            MsgBox GetExecName(intRow) & "的输血反应时间不能小于执行时间！", vbInformation, gstrSysName
                            .Row = intRow: .Col = .ColIndex("反应时间")
                            .ShowCell .Row, .Col
                            If .Enabled And .Visible Then .SetFocus
                            Exit Function
                        End If
                     End If
                     '本阶段的输血反应时间不能大于下一阶段的输血反应时间
                     For intNewRow = intRow + 1 To .Rows - 1
                        If IsDate(.TextMatrix(intNewRow, .ColIndex("执行时间"))) Then
                            If CDate(Format(.TextMatrix(intRow, .ColIndex("反应时间")), "YYYY-MM-DD HH:mm")) >= CDate(Format(.TextMatrix(intNewRow, .ColIndex("执行时间")), "YYYY-MM-DD HH:mm")) Then
                                MsgBox GetExecName(intRow) & "的输血反应时间必须小于" & GetExecName(intNewRow) & "的执行时间！", vbInformation, gstrSysName
                                .Row = intRow: .Col = .ColIndex("反应时间")
                                .ShowCell .Row, .Col
                                If .Enabled And .Visible Then .SetFocus
                                Exit Function
                            End If
                        End If
                     Next
                End If
                '检查生命体征数据
                If .TextMatrix(intRow, .ColIndex("体温")) <> "" And Not IsNumeric(.TextMatrix(intRow, .ColIndex("体温"))) Then
                    MsgBox GetExecName(intRow) & "的体温不是有效数字格式！", vbInformation, gstrSysName
                    .Row = intRow: .Col = .ColIndex("体温")
                    .ShowCell .Row, .Col
                    If .Enabled And .Visible Then .SetFocus
                    Exit Function
                End If
                If .TextMatrix(intRow, .ColIndex("脉搏")) <> "" And Not IsNumeric(.TextMatrix(intRow, .ColIndex("脉搏"))) Then
                    MsgBox GetExecName(intRow) & "的脉搏不是有效数字格式！", vbInformation, gstrSysName
                    .Row = intRow: .Col = .ColIndex("脉搏")
                    .ShowCell .Row, .Col
                    If .Enabled And .Visible Then .SetFocus
                    Exit Function
                End If
                If .TextMatrix(intRow, .ColIndex("呼吸")) <> "" And Not IsNumeric(.TextMatrix(intRow, .ColIndex("呼吸"))) Then
                    MsgBox GetExecName(intRow) & "的呼吸不是有效数字格式！", vbInformation, gstrSysName
                    .Row = intRow: .Col = .ColIndex("呼吸")
                    .ShowCell .Row, .Col
                    If .Enabled And .Visible Then .SetFocus
                    Exit Function
                End If
                If .TextMatrix(intRow, .ColIndex("血压")) <> "" And InStr(1, .TextMatrix(intRow, .ColIndex("血压")), "/") = 0 Then
                    MsgBox GetExecName(intRow) & "的血压不是有效血压格式！", vbInformation, gstrSysName
                    .Row = intRow: .Col = .ColIndex("血压")
                    .ShowCell .Row, .Col
                    If .Enabled And .Visible Then .SetFocus
                    Exit Function
                End If
            Next
            '录入结束后四小时，则结束时间不能为空
            If IsDate(.TextMatrix(.Rows - 1, .ColIndex("执行时间"))) And .TextMatrix(.Rows - 2, .ColIndex("执行时间")) = "" Then
                MsgBox "录入了输注结束后4小时，则必须录入输血结束！", vbInformation, gstrSysName
                .Row = .Rows - 2: .Col = .ColIndex("执行时间")
                .ShowCell .Row, .Col
                If .Enabled And .Visible Then .SetFocus
                Exit Function
            End If
        End With
        
        If gobjCommFun.ActualLen(txt执行摘要.Text) > txt执行摘要.MaxLength Then
            MsgBox "执行摘要内容过多，最多允许 " & txt执行摘要.MaxLength \ 2 & " 个汉字或 " & txt执行摘要.MaxLength & " 个字符。", vbInformation, gstrSysName
            Call gobjControl.ControlSetFocus(txt执行摘要)
            Exit Function
        End If
        
        '本次执行次数计算
        dbl本次数次 = Val(txt本次数次.Text)
        dbl剩余次数 = gobjComlib.FormatEx(Val(txt发送数次.Text) - Val(txt发送数次.Tag), 5)
        If mint血袋数 > mint已执行血袋数 Then
            dbl本次数次 = gobjComlib.FormatEx(dbl剩余次数 / (mint血袋数 - mint已执行血袋数), 5)
        Else
            dbl本次数次 = gobjComlib.FormatEx(dbl本次数次 / mint血袋数, 5)
        End If
        If Val(txt发送数次.Tag) + dbl本次数次 > Val(txt发送数次.Text) Then
            dbl本次数次 = gobjComlib.FormatEx(Val(txt发送数次.Text) - Val(txt发送数次.Tag), 5)
        End If
        
        '保存数据
        Call SetMessages(arrMsg)
        
        blnUpFirst = False
        arrSQL = Array()
        If mstr开始时间 = "" Then
            strSQL = "ZL_病人医嘱执行_Insert(" & mlng医嘱ID & "," & mlng发送号 & "," & _
                "To_Date('" & Format(dtp要求时间.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                dbl本次数次 & ",'" & txt执行摘要.Text & "','" & gobjCommFun.GetNeedName(vsfExec.TextMatrix(vsfExec.FixedRows, vsfExec.ColIndex("执行人"))) & "'," & _
                "To_Date('" & Format(str开始执行时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                1 & "," & "0," & 1 & ",'','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
        Else
            strSQL = "ZL_病人医嘱执行_Update(To_Date('" & mstr开始时间 & "','YYYY-MM-DD HH24:MI:SS')," & mlng医嘱ID & "," & mlng发送号 & "," & _
                "To_Date('" & Format(dtp要求时间.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                dbl本次数次 & ",'" & txt执行摘要.Text & "','" & gobjCommFun.GetNeedName(vsfExec.TextMatrix(vsfExec.FixedRows, vsfExec.ColIndex("执行人"))) & "'," & _
                "To_Date('" & Format(str开始执行时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & "," & 1 & ",NULL," & 1 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
            blnUpFirst = Format(mstr开始时间, "yyyy-MM-dd HH:mm:ss") <> Format(str开始执行时间, "yyyy-MM-dd HH:mm:ss")
        End If
        strCurDate = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
        With vsfExec
            For intRow = .FixedRows To .Rows - 1
                If .TextMatrix(intRow, .ColIndex("执行时间")) <> "" And (Val(.TextMatrix(intRow, .ColIndex("状态"))) <> 1 Or (blnUpFirst = True And Val(.TextMatrix(intRow, .ColIndex("记录性质"))) = 1)) Then
                    Select Case .TextMatrix(intRow, .ColIndex("滴速"))
                        Case "加压"
                            int滴速 = -2
                        Case "快速"
                            int滴速 = -1
                        Case Else
                            int滴速 = Val(.TextMatrix(intRow, .ColIndex("滴速")))
                    End Select
                    If InStr(1, .TextMatrix(intRow, .ColIndex("血压")), "/") <> 0 Then
                        str收缩压 = Mid(.TextMatrix(intRow, .ColIndex("血压")), 1, InStr(1, .TextMatrix(intRow, .ColIndex("血压")), "/") - 1)
                        str舒张压 = Mid(.TextMatrix(intRow, .ColIndex("血压")), InStr(1, .TextMatrix(intRow, .ColIndex("血压")), "/") + 1)
                    Else
                        str收缩压 = "": str舒张压 = ""
                    End If
                    If Val(.TextMatrix(intRow, .ColIndex("记录性质"))) = 1 Then
                        str执行时间 = str开始执行时间
                    Else
                        str执行时间 = Format(Format(.TextMatrix(intRow, .ColIndex("执行时间")), "YYYY-MM-DD HH:mm") & ":" & Format(strCurDate, "ss"), "YYYY-MM-DD HH:mm:ss")
                    End If
                    strSQL = "zl_血液执行记录_Update(" & mlng收发ID & "," & Val(.TextMatrix(intRow, .ColIndex("记录性质"))) & "," & Val(.TextMatrix(intRow, .ColIndex("序号"))) & ",To_Date('" & str执行时间 & "','YYYY-MM-DD HH24:MI:SS'),'" & _
                        gobjCommFun.GetNeedName(.TextMatrix(intRow, .ColIndex("执行人"))) & "'," & mlng科室ID & "," & IIf(int滴速 = 0, "NULL", int滴速) & ",'" & .TextMatrix(intRow, .ColIndex("输血反应")) & "'," & _
                        IIf(.TextMatrix(intRow, .ColIndex("反应时间")) = "", "NULL", "To_Date('" & .TextMatrix(intRow, .ColIndex("反应时间")) & "','YYYY-MM-DD HH24:MI:SS')") & "," & _
                        IIf(.TextMatrix(intRow, .ColIndex("有无渗漏")) = "无", 0, IIf(.TextMatrix(intRow, .ColIndex("有无渗漏")) = "有", 1, "NULL")) & "," & _
                        IIf(.TextMatrix(intRow, .ColIndex("管道冲洗")) = "无", 0, IIf(.TextMatrix(intRow, .ColIndex("管道冲洗")) = "有", 1, "NULL")) & "," & _
                        IIf(.TextMatrix(intRow, .ColIndex("使用药物")) = "无", 0, IIf(.TextMatrix(intRow, .ColIndex("使用药物")) = "有", 1, "NULL")) & "," & _
                        IIf(.TextMatrix(intRow, .ColIndex("体温")) = "", "NULL", .TextMatrix(intRow, .ColIndex("体温"))) & "," & IIf(.TextMatrix(intRow, .ColIndex("脉搏")) = "", "NULL", .TextMatrix(intRow, .ColIndex("脉搏"))) & "," & _
                        IIf(.TextMatrix(intRow, .ColIndex("呼吸")) = "", "NULL", .TextMatrix(intRow, .ColIndex("呼吸"))) & "," & IIf(str收缩压 = "", "NULL", str收缩压) & "," & _
                        IIf(str舒张压 = "", "NULL", str舒张压) & ",'" & UserInfo.姓名 & "',NULL,'" & txt执行摘要.Text & "')"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = strSQL
                    If chkSign.Value <> 0 Then '保存同时完成数据签名锁定
                        strSQL = "Zl_血液执行记录_Sign(" & mlng收发ID & "," & Val(.TextMatrix(intRow, .ColIndex("记录性质"))) & "," & Val(.TextMatrix(intRow, .ColIndex("序号"))) & ",'" & UserInfo.姓名 & "',1)"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = strSQL
                    End If
                ElseIf InStr(1, ",0,4,", "," & Val(.TextMatrix(intRow, .ColIndex("状态"))) & ",") = 0 And .TextMatrix(intRow, .ColIndex("执行时间")) = "" Then
                    strSQL = "Zl_血液执行记录_Delete(" & mlng收发ID & "," & Val(.TextMatrix(intRow, .ColIndex("记录性质"))) & "," & Val(.TextMatrix(intRow, .ColIndex("序号"))) & ")"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                End If
            Next
        End With
    Else
        blnDelete = False
        arrSQL = Array()
        With vsfExec
            For intRow = .FixedRows To .Rows - 1
                If InStr(1, ",0,4,", "," & Val(.TextMatrix(intRow, .ColIndex("状态"))) & ",") = 0 Then
                    blnDelete = True
                    Exit For
                End If
            Next
            If blnDelete = True Then
                Call SetMessages(arrMsg, True)
                strSQL = "ZL_病人医嘱执行_Delete(" & mlng医嘱ID & "," & mlng发送号 & ",To_Date('" & mstr开始时间 & "','YYYY-MM-DD HH24:MI:SS'))"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
                
                strSQL = "Zl_血液执行记录_Delete(" & mlng收发ID & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            Else
                GoTo GOEND
            End If
        End With
    End If
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        If CStr(arrSQL(i)) <> "" Then
            Call gobjDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        End If
    Next
    '消息数据保存
    For i = 0 To UBound(arrMsg)
        If CStr(arrMsg(i)) <> "" Then
            Call gobjDatabase.ExecuteProcedure(CStr(arrMsg(i)), Me.Caption)
        End If
    Next
    gcnOracle.CommitTrans: blnTrans = False
GOEND:
    SaveData = True
    Exit Function
ErrHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetExecName(ByVal intRow As Integer) As String
'功能：获取执行对应的阶段名称
    Dim strName As String
    With vsfExec
        Select Case Val(.TextMatrix(intRow, .ColIndex("记录性质")))
            Case 1
                strName = "输注前15分钟"
            Case 2
                If Val(.TextMatrix(intRow, .ColIndex("序号"))) <= 0 Then
                    strName = "输注15分钟后"
                Else
                    strName = "输注15分钟后" & Val(.TextMatrix(intRow, .ColIndex("序号"))) & "小时"
                End If
            Case 3
                strName = "输注结束"
            Case 4
                strName = "输注结束后4小时"
        End Select
    End With
    
    GetExecName = strName
End Function

Private Sub Form_Unload(Cancel As Integer)
    Err.Clear
    On Error Resume Next
    Call gobjDatabase.SetPara("保存执行登记同时签名", chkSign.Value, 2200, 9005)
    If Not mrsPersons Is Nothing Then
        If mrsPersons.State = adStateOpen Then mrsPersons.Close
        Set mrsPersons = Nothing
    End If
    If Not mrsItems Is Nothing Then
        If mrsItems.State = adStateOpen Then mrsItems.Close
        Set mrsItems = Nothing
    End If
    Call gobjComlib.SaveWinState(Me, App.ProductName)
    If Err <> 0 Then Err.Clear
End Sub

Private Sub lstSelect_DblClick()
    Call lstSelect_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lstSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call MoveNextCell
    End If
End Sub

Private Sub msk时间_GotFocus()
    msk时间.SelStart = 0: msk时间.SelLength = Len(msk时间.Text)
End Sub

Private Sub msk时间_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picDate_KeyDown(KeyCode, Shift)
End Sub

Private Sub picCbo_GotFocus()
    If cboEdit.Enabled And cboEdit.Visible Then cboEdit.SetFocus
End Sub

Private Sub picCbo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call MoveNextCell(, , True)
    End If
End Sub

Private Sub picDate_GotFocus()
    If msk时间.Visible And msk时间.Enabled Then msk时间.SetFocus
End Sub

Private Sub picDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call MoveNextCell
    End If
End Sub

Private Sub picText_GotFocus()
    If TxtEdit.Enabled And TxtEdit.Visible Then TxtEdit.SetFocus
End Sub

Private Sub picText_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call MoveNextCell
    End If
End Sub

Private Sub TimeFlash_Timer()
    mintTimerCount = mintTimerCount + 1
    
    If mintTimerCount Mod 2 = 0 Then
        lblTitle.ForeColor = 0
    Else
        lblTitle.ForeColor = 255
    End If
    
    If mintTimerCount = 10 Then mintTimerCount = 0
End Sub

Private Sub TxtEdit_GotFocus()
    Call gobjControl.TxtSelAll(TxtEdit)
End Sub

Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picText_KeyDown(KeyCode, Shift)
End Sub

Private Sub txt执行摘要_Change()
    If mblnAcTive = True Then Exit Sub
    mblnChange = True
End Sub

Private Sub txt执行摘要_GotFocus()
    Call gobjControl.TxtSelAll(txt执行摘要)
End Sub

Private Sub vsfExec_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim str值域 As String, int小数 As Integer, strTmp As String
    Dim arrName, i As Integer
    Dim blnMatch As Boolean
    
    Call ShowMsg("")
    On Error GoTo ErrHand
    If vsfExec.Cell(flexcpBackColor, NewRow, NewCol) <> 16772055 Then
        Call SetColBackColor(16772055)
    End If
    If mblnOnlyRead = True Then
        vsfExec.FocusRect = flexFocusNone
    Else
        If NewCol >= vsfExec.FixedCols And NewCol < mlngNoEditor Then
            vsfExec.FocusRect = flexFocusSolid
        Else
            vsfExec.FocusRect = flexFocusHeavy
        End If
    End If
    If vsfExec.TextMatrix(NewRow, vsfExec.ColIndex("签名人")) <> "" Then
        strTmp = "本条执行记录已被操作员[" & vsfExec.TextMatrix(NewRow, vsfExec.ColIndex("签名人")) & "]签名"
        Call ShowMsg(strTmp, vbRed)
        Exit Sub
    End If
    Select Case NewCol
        Case vsfExec.ColIndex("体温"), vsfExec.ColIndex("脉搏"), vsfExec.ColIndex("呼吸")
            str值域 = "": int小数 = -1
            mrsItems.Filter = "中文名='" & vsfExec.ColKey(NewCol) & "'"
            If Not mrsItems.EOF Then
                Select Case vsfExec.ColKey(NewCol)
                    Case "体温"
                        blnMatch = mrsItems!单位 & "" = "℃"
                    Case "脉搏", "呼吸"
                        blnMatch = mrsItems!单位 & "" = "次/分"
                    Case "收缩压", "舒张压"
                        blnMatch = mrsItems!单位 & "" = "mmHg"
                End Select
                If blnMatch = True Then
                    str值域 = mrsItems!数值域 & ""
                    int小数 = Val(mrsItems!小数 & "")
                End If
            End If
            If InStr(1, str值域, ";") = 0 Or int小数 = -1 Then
             '找不到则使用缺省值
                Select Case vsfExec.ColKey(NewCol)
                    Case "体温"
                        If InStr(1, str值域, ";") = 0 Then str值域 = "35;42"
                        If int小数 = -1 Then int小数 = 1
                    Case "脉搏"
                        If InStr(1, str值域, ";") = 0 Then str值域 = "20;300"
                        If int小数 = -1 Then int小数 = 0
                    Case "呼吸"
                        If InStr(1, str值域, ";") = 0 Then str值域 = "15;50"
                        If int小数 = -1 Then int小数 = 0
                    Case "收缩压", "舒张压"
                        If InStr(1, str值域, ";") = 0 Then str值域 = "50;190"
                        If int小数 = -1 Then int小数 = 0:
                End Select
            End If
            
            If Left(str值域, 1) = "." Then str值域 = "0" & str值域
            strTmp = Replace(str值域, ";", " - ")
            If int小数 = 0 Then
                strTmp = "可录入范围为 " & strTmp & " 的数"
            Else
                strTmp = "可录入范围为 " & strTmp & " 之间的数，可含" & int小数 & "位小数"
            End If
            Call ShowMsg(strTmp)
        Case vsfExec.ColIndex("血压")
            arrName = Array("收缩压", "舒张压")
            For i = 0 To UBound(arrName)
                mrsItems.Filter = "中文名='" & vsfExec.ColKey(NewCol) & "'"
                If Not mrsItems.EOF Then
                    str值域 = mrsItems!数值域 & ""
                    int小数 = Val(mrsItems!小数 & "")
                End If
                If InStr(1, str值域, ";") = 0 Then str值域 = "50;190"
                If int小数 = -1 Then int小数 = 0
                If Left(str值域, 1) = "." Then str值域 = "0" & str值域
                If i = 0 Then
                    strTmp = Replace(str值域, ";", " - ")
                Else
                    strTmp = strTmp & "/" & Replace(str值域, ";", " - ")
                End If
            Next
            strTmp = "格式为[收缩压/舒张压]，可录入范围为 " & strTmp
            Call ShowMsg(strTmp)
    End Select
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsfExec_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    If mblnAcTive = True Then Exit Sub
    Call HiddenEditControl
End Sub

Private Sub vsfExec_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mblnAcTive = True Then Exit Sub
    If mintType = -1 Then Exit Sub
    Cancel = Not MoveNextCell(True, True)
End Sub

Private Sub vsfExec_DblClick()
    Call vsfExec_KeyDown(Asc("A"), 0)
End Sub

Private Sub vsfExec_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    If DrawTimeCell(hDC, Row, Col, Left, Top, Right, Bottom) = False Then Exit Sub
    Done = True
End Sub

Private Function DrawTimeCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As Boolean
    Dim rc As RECT
    Dim rcUp As RECT
    Dim lngLoop As Long
    Dim lngSvrBkColor As Long
    Dim lngCenterTop As Long
    Dim lngPen As Long, lngOldPen As Long
    Dim lpPoint As POINTAPI
    If Not (Row = 0 And (Col = 0 Or Col = 1)) Then Exit Function
    With vsfExec
        rc.Left = Left
        rc.Top = Top
        rc.Bottom = Bottom - 1
        rc.Right = Right - Col
        lngSvrBkColor = &H8000000F
        Call SetBkColor(hDC, GetRBGFromOLEColor(lngSvrBkColor))
        Call ExtTextOut(hDC, rc.Left, rc.Top, 2, rc, " ", 1, lngLoop)
        
        '求第1列左侧顶端和第二列右侧低端连线，在两列中间的交叉点
        lngCenterTop = (.RowHeight(0) * .ColWidth(1)) / (.ColWidth(0) + .ColWidth(1)) \ 15 + 2
        '画线
        lngPen = CreatePen(PS_SOLID, 1, vbBlack)
        lngOldPen = SelectObject(hDC, lngPen)
        '绘图
        If Col = 0 Then
            Call MoveToEx(hDC, rc.Left, rc.Top, lpPoint)
            Call LineTo(hDC, rc.Right, lngCenterTop)
            '9号字体的高度和宽度就是180，像素就是12
            Call TextOut(hDC, rc.Left + 2, rc.Bottom - 12 - 2, "过程", 4)
        Else
            '绘图
            Call MoveToEx(hDC, rc.Left, lngCenterTop, lpPoint)
            Call LineTo(hDC, rc.Right, rc.Bottom)
            '9号字体的高度和宽度就是180，像素就是12
            Call TextOut(hDC, rc.Right - 24 - 2, rc.Top + 2, "内容", 4)
        End If
        '还原画笔并销毁
        Call SelectObject(hDC, lngOldPen)
        Call DeleteObject(lngPen)
    End With
End Function

Private Sub vsfExec_EnterCell()
    
    If mblnAcTive = True Or mblnOnlyRead = True Then Exit Sub
    '隐藏编辑的控件
    picDate.Visible = False
    picText.Visible = False
    lstSelect.Visible = False
    picCbo.Visible = False
    dtpDate.Visible = False
    
    With vsfExec
        If mblnShow = True Then
            If .TextMatrix(.Row, .ColIndex("签名人")) <> "" Then Exit Sub
            '未填写执行时间 不允许填写后面的
            If .TextMatrix(.Row, .ColIndex("执行时间")) = "" And .Col <> .ColIndex("执行时间") Then
                If .Col >= .FixedCols And .Col < mlngNoEditor Then
                    Call ShowMsg("填写了执行时间，才允许填写其他项", vbRed)
                End If
                Exit Sub
            End If
            '要填写输血反应时间，则必须填写输血反应
            If .Col = .ColIndex("反应时间") And InStr(1, "'无'", "'" & Trim(.TextMatrix(.Row, .ColIndex("输血反应")))) <> 0 Then
                Exit Sub
            End If
            '输血管道冲洗，只允许输血前和输血后填写
            If Val(.TextMatrix(.Row, .ColIndex("记录性质"))) = 2 And .Col = .ColIndex("管道冲洗") Then Exit Sub
            '输血结束后4小时，不填写 滴速、是否使用药物、是否渗漏、管道冲洗
            If Val(.TextMatrix(.Row, .ColIndex("记录性质"))) = 4 Then
                If .Col = .ColIndex("滴速") Or .Col = .ColIndex("有无渗漏") Or .Col = .ColIndex("使用药物") Or .Col = .ColIndex("管道冲洗") Then
                    Exit Sub
                End If
            End If
        End If
    End With
    If Not mblnShow Then Exit Sub
    '开始显示控件
    If vsfExec.Col < mlngNoEditor Then Call ShowInput
End Sub

Private Sub vsfExec_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not (vsfCheck.Col >= vsfCheck.FixedRows And vsfExec.Row >= vsfExec.FixedRows) Then Exit Sub
    If KeyCode = vbKeyReturn Then
        If mblnShow = False And vsfExec.Col = vsfExec.ColIndex("执行时间") Then
            mblnShow = True
            Call vsfExec_EnterCell
        Else
            Call MoveNextCell
        End If
    ElseIf Not (KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyDelete Or Shift <> 0) Then
            mblnShow = True
            Call vsfExec_EnterCell
    ElseIf KeyCode = vbKeyDelete Then
        If vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("签名人")) <> "" Then Exit Sub
        If vsfExec.TextMatrix(vsfExec.Row, vsfExec.Col) <> "" And vsfExec.Col <> vsfExec.ColIndex("执行时间") Then
            HiddenEditControl
            vsfExec.TextMatrix(vsfExec.Row, vsfExec.Col) = ""
            mblnChange = True
            Call ChangeDataState
        End If
    End If
End Sub

Public Function SetColBackColor(Optional ByVal lngColor As Long = 16772055) As Boolean
    '******************************************************************************************************************
    '功能:设置列的背景色
    '******************************************************************************************************************
    Dim lngLoop As Long
    
    On Error Resume Next
    
    For lngLoop = vsfExec.FixedCols To vsfExec.Cols - 1
        If vsfExec.ColHidden(lngLoop) = False Then
            vsfExec.Cell(flexcpBackColor, vsfExec.FixedRows, lngLoop, vsfExec.Rows - 1, lngLoop) = 16777215
        End If
    Next
    For lngLoop = vsfExec.FixedCols To vsfExec.Cols - 1
        If vsfExec.Cell(flexcpBackColor, vsfExec.Row, lngLoop, vsfExec.Row, lngLoop) = 16777215 Then
            vsfExec.Cell(flexcpBackColor, vsfExec.Row, lngLoop, vsfExec.Row, lngLoop) = lngColor
        End If
    Next
    If Err <> 0 Then Err.Clear
End Function

Private Sub ShowInput()
'显示对应的编辑控件
    Dim i As Integer, lngLegth As Long
    Dim strText As String
    Dim CellRect As RECT
    Dim lngFindCboIndex As Long
    Dim arrTmp
    
    With vsfExec
        If vsfExec.ColIsVisible(vsfExec.Col) = False Then
            vsfExec.LeftCol = vsfExec.Col
        End If
        If vsfExec.RowIsVisible(vsfExec.Row) = False Then
            vsfExec.TopRow = vsfExec.Row
        End If
        '确定控件的位置
        CellRect.Left = .CellLeft
        CellRect.Top = .CellTop
        CellRect.Right = .CellWidth - 10
        CellRect.Bottom = .CellHeight - 10
        strText = .TextMatrix(.Row, .Col)
        '确定要显示的控件
        Select Case .ColKey(.Col)
            Case "执行时间", "反应时间"
                mintType = 0
                picDate.Left = CellRect.Left
                picDate.Top = CellRect.Top
                picDate.Width = CellRect.Right
                picDate.Height = CellRect.Bottom
                picDate.BackColor = .BackColor
                picDate.BorderStyle = 0
                picDate.Visible = True
                picDate.ZOrder 0
                picDate.SetFocus
                picDate.Tag = .ColKey(.Col)
                '赋值
                If IsDate(strText) Then
                    msk时间.Text = Format(strText, "YYYY-MM-DD HH:mm")
                Else
                    msk时间.Text = "____-__-__ __:__"
                End If
                cmdDate.Tag = strText
            Case "有无渗漏", "使用药物", "输血反应", "管道冲洗"
                mintType = 1
                arrTmp = Split(mstr输血反应, "'")
                lstSelect.Clear
                If .ColKey(.Col) = "输血反应" And UBound(arrTmp) >= 0 Then
                    For i = 0 To UBound(arrTmp)
                        lstSelect.AddItem CStr(arrTmp(i))
                        If CStr(arrTmp(i)) = mstr缺省输血反应 Then
                            lstSelect.Selected(lstSelect.NewIndex) = True
                        End If
                    Next
                Else
                    lstSelect.AddItem "无"
                    lstSelect.AddItem "有"
                End If
                lstSelect.Left = CellRect.Left
                lstSelect.Top = CellRect.Top + CellRect.Bottom
                lstSelect.Width = CellRect.Right
                lstSelect.Height = lstSelect.ListCount * (picText.TextHeight("刘")) + picText.TextHeight("刘") \ 3
                If lstSelect.Height < CellRect.Bottom Then lstSelect.Height = CellRect.Bottom
                
                '获取最长的长度内容
                lngLegth = 0
                For i = 0 To lstSelect.ListCount - 1
                    If lngLegth < LenB(StrConv(lstSelect.List(i), vbFromUnicode)) Then
                        lngLegth = LenB(StrConv(lstSelect.List(i), vbFromUnicode))
                    End If
                    If strText = lstSelect.List(i) Then
                        lstSelect.Selected(i) = True
                    Else
                        lstSelect.Selected(i) = False
                    End If
                Next
                lstSelect.Width = lngLegth * picText.TextWidth("1") + 60    '宽度以最长的内容为准
                If lstSelect.Width < CellRect.Right Then lstSelect.Width = CellRect.Right
                If lstSelect.Height + lstSelect.Top > vsfExec.ClientHeight Then
                    If vsfExec.ClientHeight - CellRect.Top > vsfExec.ClientHeight - lstSelect.Top Then
                         lstSelect.Top = vsfExec.ClientHeight - lstSelect.Height
                         If lstSelect.Top < vsfExec.RowHeight(0) Then lstSelect.Top = vsfExec.RowHeight(0)
                         If lstSelect.Top + lstSelect.Height > vsfExec.ClientHeight Then
                             lstSelect.Height = vsfExec.ClientHeight - lstSelect.Top
                         End If
                    Else
                        lstSelect.Height = vsfExec.ClientHeight - lstSelect.Top
                    End If
                End If
                
                lstSelect.Visible = True
                lstSelect.ZOrder 0
                lstSelect.Tag = .ColKey(.Col)
                lstSelect.SetFocus
            Case "体温", "脉搏", "呼吸", "血压"
                mintType = 2
                picText.Left = CellRect.Left
                picText.Top = CellRect.Top
                picText.Width = CellRect.Right
                picText.Height = CellRect.Bottom
                picText.BackColor = .BackColor
                picText.BorderStyle = 0
                TxtEdit.Width = picText.Width
                TxtEdit.Height = picText.Height
                TxtEdit.Left = 0
                TxtEdit.Top = 0
                TxtEdit.Text = strText
                TxtEdit.Tag = strText
                picText.Visible = True
                picText.ZOrder 0
                picText.SetFocus
                picText.Tag = .ColKey(.Col)
            Case "滴速", "执行人"
                mintType = 3
                cboEdit.Clear
                cboEdit.Text = ""
                cboEdit.locked = False
                cboEdit.Tag = strText
                If .ColKey(.Col) = "滴速" Then
                    cboEdit.AddItem 15: cboEdit.ItemData(cboEdit.NewIndex) = 15
                    cboEdit.AddItem 30: cboEdit.ItemData(cboEdit.NewIndex) = 30
                    cboEdit.AddItem "快速": cboEdit.ItemData(cboEdit.NewIndex) = -1
                    cboEdit.AddItem "加压": cboEdit.ItemData(cboEdit.NewIndex) = -2
                    gobjComlib.cbo.SetText cboEdit, strText
                Else
                    mrsPersons.Filter = ""
                    lngFindCboIndex = -1
                    Do While Not mrsPersons.EOF
                        cboEdit.AddItem mrsPersons!姓名 'mrsPersons!编号 & "-" & mrsPersons!姓名
                        cboEdit.ItemData(cboEdit.NewIndex) = Val("" & mrsPersons!id)
                        If strText = "" Then
                             If mrsPersons!id = UserInfo.id Then
                                lngFindCboIndex = cboEdit.NewIndex
                            End If
                        Else
                            If strText = mrsPersons!姓名 Then
                                lngFindCboIndex = cboEdit.NewIndex
                            End If
                        End If
                        mrsPersons.MoveNext
                    Loop
                    
                    If cboEdit.ListCount > 0 And cboEdit.ListIndex = -1 Then
                        If lngFindCboIndex <> -1 Then
                            cboEdit.ListIndex = lngFindCboIndex
                        ElseIf strText <> "" Then
                            gobjComlib.cbo.SetText cboEdit, strText
                        Else
                            cboEdit.ListIndex = 0
                        End If
                    End If
                    If mlngModul = p医技工作站 Then
                        If Val(gobjDatabase.GetPara(51, 100)) = 1 Then
                            cboEdit.locked = True
                        End If
                    End If
                End If
                picCbo.Left = CellRect.Left
                picCbo.Top = CellRect.Top
                picCbo.Width = CellRect.Right
                picCbo.Height = CellRect.Bottom
                If cboEdit.locked = False Then
                    cboEdit.Width = picCbo.Width + 30
                Else
                    cboEdit.Width = picCbo.Width + 300
                End If
                gobjControl.CboSetHeight cboEdit, picCbo.Height + 30
                cboEdit.Left = -15
                cboEdit.Top = -15
                '设置展开宽度
                lngLegth = 0
                For i = 0 To cboEdit.ListCount - 1
                    If lngLegth < LenB(StrConv(cboEdit.List(i), vbFromUnicode)) Then
                        lngLegth = LenB(StrConv(cboEdit.List(i), vbFromUnicode))
                    End If
                Next i
                If lngLegth * picText.TextWidth("1") + 60 > picCbo.Width Then
                    Call gobjControl.CboSetWidth(cboEdit.hWnd, lngLegth * picText.TextWidth("1") + 60)
                Else
                    Call gobjControl.CboSetWidth(cboEdit.hWnd, picCbo.Width)
                End If
                picCbo.BackColor = .BackColor
                picCbo.BorderStyle = 0
                picCbo.Visible = True
                picCbo.ZOrder 0
                picCbo.SetFocus
                picCbo.Tag = .ColKey(.Col)
        End Select
    End With
End Sub

Private Function GetDataToPersons(Optional ByVal strIn As String = "", Optional ByVal blnRetrunSQL As Boolean, Optional strRetrunSQL As String) As ADODB.Recordset
'功能相应科室的医护人员信息
    Dim strSQL As String, strNewSQL As String, strWhere As String
    Dim blnYn As Boolean
    
    On Error GoTo ErrHand
    If strIn <> "" Then blnYn = True
    
    '医技术站，只有执行他可项目才能看到其他科室的医嘱，临床护士站可能会具有全员病人的权限，需要加上操作员本人(能在病区看到病人，要么操作就是该病区的，要么就是具有全员病区权限)
    If InStr(mstrPrivs, "执行他科项目") > 0 Or Not (mlngModul = p医技工作站) Then
        strNewSQL = " Union " & vbNewLine & _
                            " Select " & UserInfo.id & " id,'" & UserInfo.编号 & "' 编号,'" & UserInfo.姓名 & "' 姓名,'" & UserInfo.简码 & "' 简码 From Dual "
    End If
        
    If Not mlngModul = p医技工作站 Then
        strWhere = "  Exists (Select 1 From 人员性质说明 Where 人员id = a.Id And Instr(',医生,护士,', ',' || 人员性质 || ',', 1) <> 0)"
    End If
    
    '当前登录操作员优先显示在前面
    If strNewSQL = "" Then
        strSQL = "Select a.Id, a.编号, a.姓名, a.简码" & vbNewLine & _
            " From 人员表 a, 部门人员 b" & vbNewLine & _
            " Where a.Id = b.人员id And b.部门id = [1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And" & vbNewLine & _
            "      (a.站点 = ' & gstrNodeNo & ' Or a.站点 Is Null) " & vbNewLine & _
            IIf(blnYn, " And (A.编号 Like [2] Or A.简码 Like [3] Or A.姓名 Like [3])", "") & vbNewLine & _
            IIf(strWhere = "", "", " And " & strWhere) & " Order by Decode(a.id," & IIf(blnYn = True, "[4]", "[2]") & ",0,1),a.编号"
    Else
        strSQL = "Select a.Id, a.编号, a.姓名, a.简码" & vbNewLine & _
            " From 人员表 a, 部门人员 b" & vbNewLine & _
            " Where a.Id = b.人员id And b.部门id = [1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And" & vbNewLine & _
            "      (a.站点 = ' & gstrNodeNo & ' Or a.站点 Is Null)" & IIf(strWhere = "", "", " And " & strWhere) & vbNewLine & _
            strNewSQL
        If blnYn Then
            strSQL = " Select a.Id, a.编号, a.姓名, a.简码 From (" & strSQL & ") a" & vbNewLine & _
                " Where (A.编号 Like [2] Or A.简码 Like [3] Or A.姓名 Like [3])  Order by Decode(a.id,[4],0,1),a.编号"
        Else
            strSQL = " Select a.Id, a.编号, a.姓名, a.简码 From (" & strSQL & ") a Order by Decode(a.id,[2],0,1),a.编号"
        End If
    End If
    If blnRetrunSQL Then
        strRetrunSQL = strSQL
    Else
        If blnYn = True Then
            Set GetDataToPersons = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng科室ID, UCase(strIn) & "%", gstrLike & UCase(strIn) & "%", UserInfo.id)
        Else
            Set GetDataToPersons = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng科室ID, UserInfo.id)
        End If
    End If
    
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function MoveNextCell(Optional ByVal blnNext As Boolean = True, Optional ByVal blnNoMove As Boolean = False, Optional ByVal blnKeyReturn As Boolean) As Boolean
'功能：数据赋值、校验处理
    Dim strText As String
    Dim strMsg As String
    Dim intRow As Integer
    Dim blnFind As Boolean, int序号 As Integer
    Dim vPoint As RECT, strName As String
    If mblnAcTive = True Then Exit Function
    On Error GoTo ErrHand
    If mintType >= 0 Then
        Select Case mintType
            Case 0
                If msk时间.Text = "____-__-__ __:__" Then
                    strText = ""
                ElseIf Not IsDate(msk时间.Text) Then
                    strMsg = "[" & picDate.Tag & "]不是有效的时间格式！"
                    Call ShowMsg(strMsg, vbRed)
                    If picDate.Enabled And picDate.Visible Then picDate.SetFocus
'                    If IsDate(cmdDate.Tag) Then
'                        msk时间.Text = cmdDate.Tag
'                    Else
'                        msk时间.Text = "____-__-__ __:__"
'                    End If
                    Exit Function
                Else
                    strText = msk时间.Text
                End If
            Case 1
                strText = lstSelect.Text
            Case 2
                strText = TxtEdit.Text
                If CheckVitalSigns(strText, strMsg) = False Then
                    Call ShowMsg(strMsg, vbRed)
                    If picText.Enabled And picText.Visible Then picText.SetFocus
'                    strText = TxtEdit.Tag
                    Exit Function
                End If
            Case 3
                strText = cboEdit.Text
                If strText <> "" Then
                    Select Case picCbo.Tag
                        Case "滴速"
                            If InStr(1, ",快速,加压,", "," & strText & ",") = 0 And Not IsNumeric(strText) Then
                                strMsg = "录入的[滴速]不是非数字类型的【快速、加压】或整数类型！"
                                Call ShowMsg(strMsg, vbRed)
                                If picCbo.Enabled And picCbo.Visible Then picCbo.SetFocus
    '                            strText = cboEdit.Tag
                                Exit Function
                            End If
                        Case "执行人"
                            blnFind = False
                            mrsPersons.Filter = ""
                            Do While Not mrsPersons.EOF
                                If mrsPersons!姓名 & "" = strText Then
                                    blnFind = True
                                    Exit Do
                                End If
                                mrsPersons.MoveNext
                            Loop
                            If blnFind = False Then
                                If blnKeyReturn = True Then
                                    vPoint.Left = picCbo.Left + vsfExec.Left
                                    vPoint.Top = picCbo.Top + vsfExec.Top + picCbo.Height
                                    blnFind = FindPerson(strText, vPoint.Left, vPoint.Top, strName)
                                    If blnFind = True And strName <> "" Then
                                        strText = strName
                                    End If
                                End If
                            End If
                            If blnFind = False Then
                                strMsg = "录入的[执行人]不在有效的执行人范围内！"
                                Call ShowMsg(strMsg, vbRed)
                                If picCbo.Enabled And picCbo.Visible Then picCbo.SetFocus
    '                            strText = cboEdit.Tag
                                Exit Function
                            End If
                    End Select
                End If
        End Select
        If vsfExec.TextMatrix(vsfExec.Row, vsfExec.Col) <> strText Then
            If mblnChange = False Then mblnChange = True
            Call ChangeDataState
        End If
        vsfExec.TextMatrix(vsfExec.Row, vsfExec.Col) = strText
        Call HiddenEditControl
    End If
    If (vsfExec.Col = vsfExec.ColIndex("执行时间") Or vsfExec.Col = vsfExec.ColIndex("执行人")) And Val(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("记录性质"))) = 2 Then
        If vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("执行时间")) <> "" And vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("执行人")) <> "" Then
            int序号 = Val(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("序号")))
            If vsfExec.Row + 1 < vsfExec.Rows Then
                If Val(vsfExec.TextMatrix(vsfExec.Row + 1, vsfExec.ColIndex("记录性质"))) <> 2 Then
                    vsfExec.Rows = vsfExec.Rows + 1
                    vsfExec.TextMatrix(vsfExec.Rows - 1, vsfExec.ColIndex("记录性质")) = 2
                    vsfExec.TextMatrix(vsfExec.Rows - 1, vsfExec.ColIndex("序号")) = int序号 + 1
                    vsfExec.TextMatrix(vsfExec.Rows - 1, 0) = "输注过程"
                    vsfExec.TextMatrix(vsfExec.Rows - 1, 1) = int序号 + 1 & "小时"
                    vsfExec.MergeRow(vsfExec.Rows - 1) = True
                    vsfExec.RowPosition(vsfExec.Rows - 1) = vsfExec.Rows - 3
                    vsfExec.Cell(flexcpAlignment, vsfExec.FixedRows, vsfExec.FixedCols, vsfExec.Rows - 1, vsfExec.Cols - 1) = flexAlignCenterCenter
                End If
            End If
        End If
    End If
    MoveNextCell = True
    If blnNoMove = True Then Exit Function
    If blnNext Then
toMoveNextCol:
        '跳到下一列
        If vsfExec.Col < mlngNoEditor - 1 Then
            vsfExec.Col = vsfExec.Col + 1
            If vsfExec.ColWidth(vsfExec.Col) = 0 Or vsfExec.ColHidden(vsfExec.Col) Or mintType = -1 Then GoTo toMoveNextCol
        Else
toMoveNextRow:
            '跳到下一行
            mblnShow = False
            If vsfExec.Row + 1 < vsfExec.Rows Then
                vsfExec.Row = vsfExec.Row + 1
            End If
            If vsfExec.RowHidden(vsfExec.Row) Then
                If vsfExec.Row < vsfExec.Rows - 1 Then
                    If txt执行摘要.Enabled And txt执行摘要.Visible Then txt执行摘要.SetFocus
                Else
                    For intRow = vsfExec.Rows - 1 To vsfExec.FixedRows Step -1
                        If vsfExec.RowHidden(intRow) = False Then
                            vsfExec.Row = intRow
                            Exit For
                        End If
                    Next intRow
                End If
            End If
            mblnShow = True
            vsfExec.Col = vsfExec.ColIndex("执行时间")
        End If
    Else
toMovePrevCol:
        If vsfExec.Col > vsfExec.ColIndex("执行时间") Then       '护理记录单肯定有护士签名列
            vsfExec.Col = vsfExec.Col - 1
            If vsfExec.ColWidth(vsfExec.Col) = 0 Or vsfExec.ColHidden(vsfExec.Col) Or mintType = -1 Then GoTo toMovePrevCol
        Else
toMovePrevRow:
            '跳到上一行
            If vsfExec.Row > vsfExec.FixedRows Then
                vsfExec.Row = vsfExec.Row - 1
                If vsfExec.RowHidden(vsfExec.Row) Then GoTo toMovePrevRow
                vsfExec.Col = mlngNoEditor - 1
            End If
        End If
    End If
    If vsfExec.ColIsVisible(vsfExec.Col) = False Then
        vsfExec.LeftCol = vsfExec.Col
    End If
    If vsfExec.RowIsVisible(vsfExec.Row) = False Then
        vsfExec.TopRow = vsfExec.Row
    End If
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ChangeDataState()
'功能：在数据发生变化调用此过程，修改数据状态
    Select Case Val(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("状态")))
        '1,2,3表示对原有的数据进行操作，1：原始；2-修改;3-删除
        '0,4 表示对将要新增的数据操作。0-不做处理，4-新增
        Case 3
            If vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("执行时间")) <> "" Then vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("状态")) = 2
        Case 1, 2
            If vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("执行时间")) <> "" Then
                vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("状态")) = 2
            Else
                vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("状态")) = 3
            End If
        Case 0 '无数据
            If vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("执行时间")) <> "" Then vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("状态")) = 4
        Case 4
            If vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("执行时间")) = "" Then vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("状态")) = 0
    End Select
End Sub

Private Function FindPerson(ByVal strText As String, ByVal lngLeft As Long, ByVal lngTop As Long, strName As String) As Boolean
    Dim rsUser As ADODB.Recordset
    Dim strSQL As String
    Dim blnCancel As Boolean
    
    On Error GoTo ErrHand
    If strText <> "" Then
        Call GetDataToPersons(strText, True, strSQL)
        Set rsUser = gobjDatabase.ShowSQLSelect(Me, strSQL, 0, "", False, strText, "请选择人员", False, False, True, lngLeft, lngTop, 0, blnCancel, False, False, _
                    mlng科室ID, UCase(strText) & "%", gstrLike & UCase(strText) & "%", UserInfo.id)
        If Not rsUser Is Nothing Then
            If blnCancel = False Then
                If rsUser.EOF Then Exit Function
                strName = Nvl(rsUser!姓名)
            End If
        Else
            Exit Function
        End If
    End If
    FindPerson = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckVitalSigns(strText As String, strMsg As String) As Boolean
'功能：生命体征数据合法性检查
    Dim arrData, arrName
    Dim i As Integer
    Dim str值域 As String, int小数 As Integer, int长度 As Integer
    Dim dblMin As Double, dblMax As Double
    Dim blnMatch As Boolean
    
    On Error GoTo ErrHand
    arrData = Array()
    arrName = Array()
    If strText <> "" Then
        If picText.Tag = "血压" Then
            If InStr(1, strText, "/") = 0 Then
                strMsg = "录入的[" & picText.Tag & "]格式不正确，正确格式：收缩压/舒张压！"
                Exit Function
            End If
            ReDim Preserve arrData(UBound(arrData) + 1)
            arrData(UBound(arrData)) = Mid(strText, 1, InStr(1, strText, "/") - 1)
            ReDim Preserve arrName(UBound(arrName) + 1)
            arrName(UBound(arrName)) = "收缩压"
            
            ReDim Preserve arrData(UBound(arrData) + 1)
            arrData(UBound(arrData)) = Mid(strText, InStr(1, strText, "/") + 1)
            ReDim Preserve arrName(UBound(arrName) + 1)
            arrName(UBound(arrName)) = "舒张压"
        Else
            ReDim Preserve arrData(UBound(arrData) + 1)
            arrData(UBound(arrData)) = strText
            ReDim Preserve arrName(UBound(arrName) + 1)
            arrName(UBound(arrName)) = picText.Tag
        End If
        For i = 0 To UBound(arrName)
            If Not IsNumeric(arrData(i)) Then
                strMsg = "录入的[" & arrName(i) & "]不是有效的数字格式！"
                Exit Function
            End If
            '值域范围检查
            str值域 = "": int小数 = -1: int长度 = -1
            mrsItems.Filter = "中文名='" & arrName(i) & "'"
            If Not mrsItems.EOF Then
                Select Case CStr(arrName(i))
                    Case "体温"
                        blnMatch = mrsItems!单位 & "" = "℃"
                    Case "脉搏", "呼吸"
                        blnMatch = mrsItems!单位 & "" = "次/分"
                    Case "收缩压", "舒张压"
                        blnMatch = mrsItems!单位 & "" = "mmHg"
                End Select
                If blnMatch = True Then
                    str值域 = mrsItems!数值域 & ""
                    int小数 = Val(mrsItems!小数 & "")
                    int长度 = Val(mrsItems!长度 & "")
                End If
            End If
            If InStr(1, str值域, ";") = 0 Or int小数 = -1 Or int长度 - 1 Then
             '找不到则使用缺省值
                Select Case CStr(arrName(i))
                    Case "体温"
                        If InStr(1, str值域, ";") = 0 Then str值域 = "35;42"
                        If int小数 = -1 Then int小数 = 1
                        If int长度 = -1 Then int长度 = 4
                    Case "脉搏"
                        If InStr(1, str值域, ";") = 0 Then str值域 = "20;300"
                        If int小数 = -1 Then int小数 = 0
                        If int长度 = -1 Then int长度 = 3
                    Case "呼吸"
                        If InStr(1, str值域, ";") = 0 Then str值域 = "15;50"
                        If int小数 = -1 Then int小数 = 0
                        If int长度 = -1 Then int长度 = 2
                    Case "收缩压", "舒张压"
                        If InStr(1, str值域, ";") = 0 Then str值域 = "50;190"
                        If int小数 = -1 Then int小数 = 0:
                        If int长度 = -1 Then int长度 = 3
                End Select
            End If
            strText = arrData(i)
            '长度检查
            If Len(strText) > int长度 Then
                strMsg = "录入的[" & arrName(i) & "]数据超过了最大长度：" & int长度 & "！"
                Exit Function
            End If
                    
            If int小数 <> 0 Then
                If InStr(1, strText, ".") <> 0 Then
                    strText = Mid(strText, InStr(1, strText, ".") + 1)
                    If Len(strText) > int小数 Then
                        strMsg = "录入的[" & arrName(i) & "]录入小数部分超过了合法精度" & int小数 & "位！"
                        Exit Function
                    End If
                End If
            End If
            strText = arrData(i)
            If str值域 <> "" Then
                dblMin = Val(Split(str值域, ";")(0))
                dblMax = Val(Split(str值域, ";")(1))
                If Not (Val(strText) >= dblMin And Val(strText) <= dblMax) Then
                    strMsg = "录入的[" & arrName(i) & "]数据不在" & Format(dblMin, "#0.00") & "～" & Format(dblMax, "#0.00") & "的有效范围！"
                    Exit Function
                End If
            End If
            If Val(strText) < 1 And Val(strText) > 0 Then strText = "0" & Val(strText)
            arrData(i) = strText
        Next
        If picText.Tag = "血压" Then
            strText = arrData(0) & "/" & arrData(1)
        Else
            strText = arrData(0)
        End If
    End If
    CheckVitalSigns = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitExecBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
   Dim cbrCustom As CommandBarControlCustom
   
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsExec.VisualTheme = xtpThemeOfficeXP
    With Me.cbsExec.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .UseFadedIcons = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 24, 24
    End With
    Set cbsExec.Icons = gobjCommFun.GetPubIcons
    cbsExec.EnableCustomization False
    cbsExec.ActiveMenuBar.Visible = False
    
    Set objBar = cbsExec.Add("工具栏", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "核对"): objControl.ToolTipText = "输血前核对"
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingDelAudit, "取消"): objControl.ToolTipText = "取消输血前核对"
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Sign, "签名"): objControl.ToolTipText = "签名锁定": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "取消"): objControl.ToolTipText = "取消签名锁定"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Clear, "清除"): objControl.ToolTipText = "清除行内容": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "保存"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "取消")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): objControl.BeginGroup = True
        If mblnOnlyRead = False Then
            Set cbrCustom = .Add(xtpControlCustom, 0, "")
            cbrCustom.Flags = xtpFlagRightAlign
            cbrCustom.Handle = picinfo.hWnd
        End If
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsExec.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add FCONTROL, Asc("C"), conMenu_Edit_Transf_Cancle
        .Add FCONTROL, Asc("E"), conMenu_File_Exit
    End With
End Sub

Private Sub ShowMsg(ByVal strText As String, Optional lngColor As Long = vbBlack)
    lblPrompt.Caption = strText
    lblPrompt.ForeColor = lngColor
End Sub

Private Function SetMessages(ByRef arrSQL As Variant, Optional ByVal blnRead As Boolean = False) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim lng病人ID As Long, lng科室id As Long, lng病区ID As Long
    Dim lng就诊id As Long
    Dim int病人来源 As Integer
    Dim str提醒部门 As String
    Dim bln输血反应 As Boolean, intRow As Integer
    
    On Error GoTo ErrHand
    arrSQL = Array()
    strSQL = "select 主页id,挂号单,病人id,病人科室id,病人来源 from 病人医嘱记录 where id = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng相关ID)
    
    If rsTmp.State = adStateClosed Then Exit Function
    If rsTmp.RecordCount = 0 Then Exit Function
    lng病人ID = Val(rsTmp!病人id)
    lng科室id = Val(rsTmp!病人科室id)
    int病人来源 = Val(rsTmp!病人来源)
    If int病人来源 = 2 Then
        lng就诊id = Val(rsTmp!主页id)
        strSQL = "select 当前病区id from 病案主页 where 病人id = [1] and 主页id = [2]  "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng就诊id)
        lng病区ID = Val(rsTmp!当前病区id)
    Else
        lng病区ID = Val(rsTmp!病人科室id)
        strSQL = "select id 挂号id from 病人挂号记录 where no = [1] and 病人id = [2] "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, rsTmp!挂号单 & "", Val(rsTmp!病人id))
        lng就诊id = Val(rsTmp!挂号ID)
    End If
        
    If blnRead = False Then
        strSQL = "select ID,类型编码,业务标识 from 业务消息清单 where 病人ID = [1] and 就诊id = [2] and 是否已阅 = 0 "
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng就诊id)
        
        '确定要提醒的部门
        str提醒部门 = IIf(Val(lng科室id) = 0, "", lng科室id)
        If lng病区ID <> lng科室id Then
            If str提醒部门 = "" Then
                str提醒部门 = IIf(lng病区ID = 0, "", lng病区ID)
            Else
                str提醒部门 = str提醒部门 & IIf(lng病区ID = 0, "", "," & lng病区ID)
            End If
        End If
        '查询是否存在本医嘱本血袋输血反应消息
        rsTmp.Filter = "类型编码 = 'ZLHIS_BLOOD_006' And 业务标识 = '" & mlng相关ID & ":" & mlng收发ID & "'"
        '是否有输血反应
        With vsfExec
            For intRow = .FixedRows To .Rows - 1
                If IsDate(.TextMatrix(intRow, .ColIndex("执行时间"))) Then
                    If .TextMatrix(intRow, .ColIndex("输血反应")) <> "" And .TextMatrix(intRow, .ColIndex("输血反应")) <> "无" Then
                        bln输血反应 = True
                        Exit For
                    End If
                End If
            Next
        End With
        If bln输血反应 = True Then
            If rsTmp.RecordCount = 0 Then
                strSQL = "Zl_业务消息清单_Insert(" & lng病人ID & "," & lng就诊id & ","  '病人id 就诊id
                strSQL = strSQL & Val(lng科室id) & ","     '就诊科室id
                strSQL = strSQL & Val(lng病区ID) & ","      '就诊病区id
                strSQL = strSQL & int病人来源 & ","                                      '病人来源
                strSQL = strSQL & "'出现输血反应，请及时填写输血反应单。','"             '消息内容
                strSQL = strSQL & IIf(Val(int病人来源) = 1, "1000", "0100") & "','ZLHIS_BLOOD_006',"     ' 提醒场合, 类型编码
                strSQL = strSQL & "'" & mlng相关ID & ":" & mlng收发ID & "',"                      '业务标识（相关id:收发id）
                strSQL = strSQL & "1,0,NULL,'" & str提醒部门 & "',NULL)"                                                   '优先程度，是否已阅，登记时间,提醒部门
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            End If
        Else    '无输血反应，则查询是否存在有反应消息，若有，设为已读。
            If rsTmp.RecordCount > 0 Then
                strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng就诊id & ",'ZLHIS_BLOOD_006',"
                strSQL = strSQL & "3,'" & UserInfo.姓名 & "'," & lng病区ID & ",NULL,"
                strSQL = strSQL & Val(rsTmp!id) & ",NULL)"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            End If
        End If
    
        rsTmp.Filter = "类型编码 = 'ZLHIS_BLOOD_007' And 业务标识 = '" & mlng相关ID & ":" & mlng医嘱ID & ":" & mlng收发ID & "'"
        If IsDate(vsfExec.TextMatrix(vsfExec.Rows - 1, vsfExec.ColIndex("执行时间"))) Then '结束执行
            If rsTmp.RecordCount = 0 Then
                strSQL = "Zl_业务消息清单_Insert(" & lng病人ID & "," & lng就诊id & ","  '病人id 就诊id
                strSQL = strSQL & Val(lng科室id) & ","      '就诊科室id
                strSQL = strSQL & Val(lng病区ID) & ","      '就诊病区id
                strSQL = strSQL & int病人来源 & ","                                      '病人来源
                strSQL = strSQL & "'输血完成，请在24小时内收回血袋。','"                         '消息内容
                strSQL = strSQL & IIf(Val(int病人来源) = 1, "0001", "0010") & "','ZLHIS_BLOOD_007',"     ' 提醒场合, 类型编码
                strSQL = strSQL & "'" & mlng相关ID & ":" & mlng医嘱ID & ":" & mlng收发ID & "',"                      '业务标识（相关id:收发id）
                strSQL = strSQL & "1,0,NULL,'" & str提醒部门 & "',NULL)"                                                   '优先程度，是否已阅，登记时间,提醒部门                                                      '
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            End If
        Else    '
            If rsTmp.RecordCount > 0 Then
                strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & lng就诊id & ",'ZLHIS_BLOOD_007',"
                strSQL = strSQL & IIf(Val(int病人来源) = 1, 4, 3) & ",'" & UserInfo.姓名 & "'," & lng病区ID & ",NULL,"
                strSQL = strSQL & Val(rsTmp!id) & ",NULL)"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            End If
        End If
    Else
        '是否存在该血袋消息，存在则设为已读
        strSQL = "Select a.Id, a.类型编码, a.就诊id, a.业务标识" & vbNewLine & _
                    "From 业务消息清单 a" & vbNewLine & _
                    "Where a.病人id = [1] And a.就诊id = [2] And a.是否已阅 = 0 And a.类型编码 In ('ZLHIS_BLOOD_006', 'ZLHIS_BLOOD_007')"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "血袋相关消息", lng病人ID, lng就诊id)
        
        For i = 0 To 1
            '将ZLHIS_BLOOD_006的消息设为已读
            If i = 0 Then rsTmp.Filter = "业务标识 = '" & mlng相关ID & ":" & mlng收发ID & "'"
            '将ZLHIS_BLOOD_007的消息设为已读
            If i = 1 Then rsTmp.Filter = "业务标识 = '" & mlng相关ID & ":" & mlng医嘱ID & ":" & mlng收发ID & "'"
            If Not rsTmp.EOF Then
                rsTmp.MoveFirst
                Do While Not rsTmp.EOF
                    strSQL = "Zl_业务消息清单_Read(" & lng病人ID & "," & rsTmp!就诊id & ",'" & rsTmp!类型编码 & "',"
                    strSQL = strSQL & IIf(Val(int病人来源) = 1, 4, 3) & ",'" & UserInfo.姓名 & "'," & lng病区ID & ",NULL,"
                    strSQL = strSQL & Val(rsTmp!id) & ",NULL)"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                    rsTmp.MoveNext
                Loop
            End If
        Next
    End If
    SetMessages = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SignData(Optional ByVal bln签名 As Boolean = False)
    Dim strName As String, str签名人 As String
    Dim blnSign  As Boolean, strSQL As String
    Dim int记录性质 As Integer, int序号 As Integer
    
    On Error GoTo ErrHand
    int记录性质 = Val(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("记录性质")))
    int序号 = Val(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("序号")))
    If bln签名 = True Then '锁定
        If Val(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("状态"))) <> 1 Then
            MsgBox "请保存数据后在签名人！", vbInformation, gstrSysName
            Exit Sub
        End If
        If IsDate(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("执行时间"))) = False Or Trim(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("执行人"))) = "" Then
            MsgBox "执行时间和执行人不能为空！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strSQL = "Zl_血液执行记录_Sign(" & mlng收发ID & "," & int记录性质 & "," & int序号 & ",'" & UserInfo.姓名 & "',1)"
        Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
        vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("签名人")) = UserInfo.姓名
        vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("签名时间")) = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    Else
        '取消是检查是否是当前操作员，不是则需要进行身份验证
        strName = vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("签名人"))
        If strName <> UserInfo.姓名 Then
            str签名人 = gobjDatabase.UserIdentifyByUser(Me, "非本人取消，请您输入签名人的用户名和密码进行身份验证。", 100, mlngModul, "执行情况登记", , True)
            If str签名人 = "" Then Exit Sub
            If str签名人 <> strName Then
                MsgBox "只能取消自己签名的记录，当前签名人是""" & strName & """", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        '撤销签名
        strSQL = "Zl_血液执行记录_Sign(" & mlng收发ID & "," & int记录性质 & "," & int序号 & ",NULL,0)"
        Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
        vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("签名人")) = ""
        vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("签名时间")) = ""
    End If
'    vsfExec.Cell(flexcpForeColor, vsfExec.Row, vsfExec.FixedCols, vsfExec.Row, vsfExec.Cols - 1) = IIf(vsfExec.TextMatrix(vsfExec.Row, vsfExec.ColIndex("签名人")) <> "", vbRed, vbBlack)
    Call vsfExec_AfterRowColChange(0, 0, vsfExec.Row, vsfExec.Col)
    
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function AutoAdviceFinish() As Boolean
'功能：判断改医嘱下面的血液是否都已经完成执行，如果是则调用自动完成医嘱执行
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHand
    strSQL = "Select a.收发ID from 血液发送记录 A,血液发送记录 B where a.配发ID=B.配发ID and B.收发ID=[1] and nvl(a.执行状态,0) not in (2,3)"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "判断所有血液是否已经完成执行", mlng收发ID)
    If rsTmp.RecordCount = 0 Then
        '自动将医嘱标记为完成
        strSQL = "ZL_病人医嘱执行_Finish(" & mlng医嘱ID & "," & mlng发送号 & ",Null,0,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        Call gobjDatabase.ExecuteProcedure(strSQL, Me.Caption)
        AutoAdviceFinish = True
    End If
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function
