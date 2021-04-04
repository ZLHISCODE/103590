VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmFeeVerifyManger 
   Caption         =   "费用审核管理"
   ClientHeight    =   10830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15075
   Icon            =   "frmFeeVerifyManger.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10830
   ScaleWidth      =   15075
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   10470
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmFeeVerifyManger.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16907
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4584
            MinWidth        =   4584
            Picture         =   "frmFeeVerifyManger.frx":0E1E
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin VB.PictureBox picMzToZy 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   135
      ScaleHeight     =   4935
      ScaleWidth      =   14475
      TabIndex        =   3
      Top             =   780
      Width           =   14475
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   630
         TabIndex        =   27
         Top             =   150
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   635
         Appearance      =   2
         IDKindStr       =   $"frmFeeVerifyManger.frx":189F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "宋体"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12;F5"
         MustSelectItems =   "姓名"
         BackColor       =   -2147483633
      End
      Begin VB.CommandButton cmdBrush 
         Caption         =   "刷新(&N)"
         Height          =   375
         Left            =   11595
         TabIndex        =   25
         Top             =   555
         Width           =   1245
      End
      Begin VB.CheckBox chk已转出费用 
         Caption         =   "显示已转出费用"
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
         Left            =   2385
         TabIndex        =   20
         Top             =   645
         Width           =   2004
      End
      Begin VB.Frame fra单据 
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   8595
         TabIndex        =   17
         Top             =   90
         Width           =   2715
         Begin VB.ComboBox cbo收费单 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   675
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   60
            Width           =   2040
         End
         Begin VB.Label lblBill 
            AutoSize        =   -1  'True
            Caption         =   "单据"
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
            Left            =   165
            TabIndex        =   18
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.CheckBox chk审核 
         Caption         =   "显示已审核费用"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   12
         Top             =   615
         Width           =   2064
      End
      Begin VB.CommandButton cmdAllCls 
         Caption         =   "全清(&R)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1335
         TabIndex        =   15
         Top             =   4500
         Width           =   1200
      End
      Begin VB.CommandButton cmdAllSel 
         Caption         =   "全选(&A)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   90
         TabIndex        =   14
         Top             =   4500
         Width           =   1200
      End
      Begin VB.TextBox txt住院号 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6690
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   150
         Width           =   1815
      End
      Begin VB.TextBox txtOld 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   150
         Width           =   585
      End
      Begin VB.TextBox txtSex 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3915
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   150
         Width           =   600
      End
      Begin VB.TextBox txtPatient 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1305
         MaxLength       =   100
         TabIndex        =   4
         Top             =   150
         Width           =   2040
      End
      Begin VSFlex8Ctl.VSFlexGrid vsFee 
         Height          =   3120
         Left            =   108
         TabIndex        =   13
         Top             =   948
         Width           =   9816
         _cx             =   17314
         _cy             =   5503
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483628
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFeeVerifyManger.frx":1935
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
         ExplorerBar     =   2
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
      Begin VB.CommandButton cmdOk 
         Caption         =   "确定(&O)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8610
         TabIndex        =   16
         Top             =   4395
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   345
         Left            =   5700
         TabIndex        =   21
         Top             =   585
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   93323267
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   345
         Left            =   8670
         TabIndex        =   22
         Top             =   585
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   93323267
         CurrentDate     =   36588
      End
      Begin VB.Label lblSum 
         AutoSize        =   -1  'True
         Caption         =   "本次审核合计:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   150
         TabIndex        =   26
         Top             =   4170
         Width           =   1560
      End
      Begin VB.Label lbl日期 
         Caption         =   "发生日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4710
         TabIndex        =   24
         Top             =   630
         Width           =   1110
      End
      Begin VB.Label lbl至 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   8385
         TabIndex        =   23
         Top             =   675
         Width           =   120
      End
      Begin VB.Label lbl住院号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
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
         Left            =   5970
         TabIndex        =   11
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
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
         Left            =   4740
         TabIndex        =   10
         Top             =   210
         Width           =   480
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
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
         Left            =   3390
         TabIndex        =   9
         Top             =   210
         Width           =   480
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         Caption         =   "病人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Index           =   7
         Left            =   150
         TabIndex        =   5
         Top             =   195
         Width           =   480
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5010
      Left            =   495
      ScaleHeight     =   5010
      ScaleWidth      =   9510
      TabIndex        =   0
      Top             =   2280
      Width           =   9510
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   4995
         Left            =   195
         TabIndex        =   1
         Top             =   465
         Width           =   9510
         _Version        =   589884
         _ExtentX        =   16775
         _ExtentY        =   8811
         _StockProps     =   64
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   1170
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmFeeVerifyManger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mlngModule As Long, mstrPrivs As String
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar
Private Enum mPgIndex
    pg_门诊转住院 = 1
End Enum
Private mblnChange As Boolean
Private mblnFirst As Boolean
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset
Private mblnValid As Boolean
Private mblnMultiBalance As Boolean
Private Enum 医院业务
    support门诊结算作废 = 33        '医保是否支持门诊结算作废，不支持只有个人帐帐户原样退,其余的医保结算方式退为现金,支持的再判断每一种结算方式是否允许退回
    support多单据收费必须全退 = 39  '多单据收费必须全退
End Enum
Private mrsFeeList As ADODB.Recordset
Private mblnNotClick As Boolean
'-----------------------------------------------------------------------------------
'结算卡相关
Private mstrPassWord As String
'-----------------------------------------------------------------------------------

 Private Sub cbo收费单_Click()
    If mblnNotClick Then Exit Sub
    If mrsFeeList Is Nothing Then Exit Sub
    ReadListData True
End Sub

Private Sub chk审核_Click()
    If mrsFeeList Is Nothing Then Exit Sub
    ReadListData True
End Sub

Private Sub chk已转出费用_Click()
    If mrsFeeList Is Nothing Then Exit Sub
    ReadListData True
End Sub

Private Sub cmdALLCls_Click()
   Dim i As Long
    With vsFee
        '40526
        If .Rows <= 1 Or .Cols <= 0 Then Exit Sub
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> "" Then
               If Val(.TextMatrix(i, .ColIndex("转出标志"))) = 0 Then
                    .TextMatrix(i, .ColIndex("审核")) = 0
                End If
            End If
        Next
        Call SetSumMoney(True)
    End With
End Sub
Private Sub cmdAllSel_Click()
    Dim i As Long
    With vsFee
        '40526
        If .Rows <= 1 Or .Cols <= 0 Then Exit Sub
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> "" Then
                If Val(.TextMatrix(i, .ColIndex("转出标志"))) = 0 Then
                    If CheckIsInput(i) = True Then
                        .TextMatrix(i, .ColIndex("审核")) = -1
                        SetRowSelected (i)
                    End If
                End If
            End If
        Next
        Call SetSumMoney
    End With
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub cmdBrush_Click()
    If mrsInfo Is Nothing Then
        MsgBox "必须选择病人,请检查!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Sub
    ElseIf mrsInfo.State <> 1 Then
        MsgBox "必须选择病人,请检查!", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Sub
    End If
    Call ReadListData
End Sub

Private Sub cmdOK_Click()
    If SaveData = False Then
        stbThis.Panels(2).Text = "保存失败!"
        Exit Sub
    End If
    Call ReadListData
    mblnChange = False
    stbThis.Panels(2).Text = "保存成功!"
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
      If objCard.名称 Like "IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
        Exit Sub
    End If
    
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Sub
    
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

 
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC卡", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNo
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("身份证号", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub
Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关参数
    '编制:刘兴洪
    '日期:2011-02-09 11:46:35
    '---------------------------------------------------------------------------------------------------------------------------------------------


End Sub
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '------------------------------------
    Select Case Control.ID
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Parameter
            If frmFeeVrerfyParaSet.ShowMe(Me, mlngModule, mstrPrivs) = False Then Exit Sub
    Case conMenu_View_StatusBar
        stbThis.Visible = Not stbThis.Visible
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Button
        cbsThis(2).Visible = Not cbsThis(2).Visible
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each mcbrControl In cbsThis(2).Controls
            mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
        cbsThis.RecalcLayout
    Case conMenu_View_Refresh
        Call cmdBrush_Click
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                Call zlCallCustomReprot(Val(Split(Control.Parameter, ",")(0)), Trim(Split(Control.Parameter, ",")(1)))
        End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsThis_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    Err = 0: On Error Resume Next
    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With picList
        .Left = lngLeft + 50: .Top = lngTop
        .Width = lngRight - 100
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, lngID As Long, blnEnabled As Boolean
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        If Val(Me.tbPage.Selected.Tag) = mPgIndex.pg_门诊转住院 Then
            Control.Enabled = Trim(vsFee.TextMatrix(1, vsFee.ColIndex("单据号"))) <> ""
        Else
            Control.Enabled = False
        End If
    Case conMenu_View_Refresh
    End Select
End Sub
Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2011-01-25 15:22:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo ErrHand:
    Set objItem = tbPage.InsertItem(mPgIndex.pg_门诊转住院, "门诊转住院费用", picMzToZy.hWnd, 0)
    objItem.Tag = mPgIndex.pg_门诊转住院
     With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
    stbThis.Top = Me.ScaleHeight - Me.stbThis.Height
End Sub
Private Sub Form_Activate()
    Dim strKey As String
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    mblnChange = False
    
End Sub

Private Sub Form_Load()
    Dim i As Long
    RestoreWinState Me, App.ProductName
    mstrPrivs = gstrPrivs: mlngModule = glngModul
    dtpEnd.MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59")
    dtpEnd.Value = dtpEnd.MaxDate
    dtpBegin.MaxDate = dtpEnd.MaxDate
    dtpBegin.Value = dtpEnd.Value - 7
    
    Call initCardSquareData
    i = Val(zlDatabase.GetPara("门诊审核单据", glngSys, mlngModule, 2, Array(cbo收费单, lblBill), InStr(1, mstrPrivs, ";参数设置;") > 0))
    mblnFirst = True
    With cbo收费单
        mblnNotClick = True
        .AddItem "收费单"
        .ItemData(.NewIndex) = 0
        If i = 0 Then .ListIndex = .NewIndex
        .AddItem "记帐单"
        .ItemData(.NewIndex) = 1
        If i = 1 Then .ListIndex = .NewIndex
        .AddItem "收费单和记帐单"
        .ItemData(.NewIndex) = 2
        If i = 2 Then .ListIndex = .NewIndex
        If .ListIndex < 0 Then .ListIndex = .NewIndex
        mblnNotClick = False
    End With
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    Call zlDefCommandBars
    Call InitPage
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    chk审核.Value = IIf(Val(zlDatabase.GetPara("显示已审核费用", glngSys, mlngModule, 0, Array(chk审核), InStr(1, mstrPrivs, ";参数设置;") > 0)) = 0, 0, 1)
    chk已转出费用.Value = IIf(Val(zlDatabase.GetPara("显示已转出单据", glngSys, mlngModule, 0, Array(chk已转出费用), InStr(1, mstrPrivs, ";参数设置;") > 0)) = 0, 0, 1)

    Set mrsInfo = New ADODB.Recordset
    vsFee.OwnerDraw = flexODContent
    '多张单据使用多种结算方式模式
    mblnMultiBalance = zlDatabase.GetPara(79, glngSys) = "1"
    Call zlCreateObject
 End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    If mblnChange Then
        If MsgBox("注意:" & vbCrLf & "    你修改了数据,但你还未保存,是否真的要退出?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    zlDatabase.SetPara "显示已审核费用", chk审核.Value, glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "门诊审核单据", cbo收费单.ListIndex, glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0
    zlDatabase.SetPara "显示已转出单据", chk已转出费用.Value, glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0
    zl_vsGrid_Para_Save mlngModule, vsFee, Me.Caption, "明细列表", True
    SaveWinState Me, App.ProductName
    
    Call zlCloseObject
    Set mrsFeeList = Nothing
    Set mrsInfo = Nothing
End Sub
Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub

Private Sub picMzToZy_Resize()
    Err = 0: On Error Resume Next
    With picMzToZy
        cmdAllCls.Top = .ScaleHeight - cmdAllCls.Height - 50
        cmdAllSel.Top = cmdAllCls.Top
        cmdOk.Top = cmdAllCls.Top
        cmdOk.Left = .ScaleWidth - cmdOk.Width - vsFee.Left * 2
        lblSum.Top = cmdAllCls.Top - lblSum.Height - 30
        
        vsFee.Width = .ScaleWidth - vsFee.Left * 2
        vsFee.Height = lblSum.Top - vsFee.Top - 50
         'chk审核.Left = .ScaleWidth - chk审核.Width
    End With
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   If Val(tbPage.Selected.Tag) = mPgIndex.pg_门诊转住院 Then
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Else
        Exit Sub
    End If
End Sub
Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-25 15:29:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): mcbrControl.BeginGroup = True
    End With
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): mcbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
    End With
    
    '设置不常用菜单
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
     zlDefCommandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytMode=1 打印;2 预览;3 输出到EXCEL
    '编制:刘兴洪
    '日期:2011-01-25 15:14:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsGrid As VSFlexGrid, rsTemp As New ADODB.Recordset, strSQL As String
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & "门诊费用清册"
    objRow.Add "病人：" & txtPatient.Text
    objRow.Add "性别：" & txtSex.Text
    objRow.Add "年龄：" & txtOld.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Set vsGrid = vsFee
    Err = 0: On Error GoTo ErrHand:
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .Cell(flexcpData, 0, intCol) = .ColWidth(intCol)
            If .ColHidden(intCol) Or intCol = .ColIndex("选择") Then .ColWidth(intCol) = 0
        Next
    End With
    
    Set objPrint.Body = vsGrid
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
End Sub
Private Sub zlCallCustomReprot(ByVal lngSys As Long, strReprotName As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用相关的自定义报表
    '编制:刘兴洪
    '日期:2011-01-25 15:16:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As Variant, lng结帐ID As Long
    With vsFee
        If .Row > 0 Then
            strNO = Trim(.TextMatrix(.Row, .ColIndex("单据号")))
            lng结帐ID = Val(.Cell(flexcpData, .Row, .ColIndex("单据号")))
        End If
        If strNO <> "" Then
            Call ReportOpen(gcnOracle, lngSys, strReprotName, Me, _
                "NO=" & strNO, "结帐ID=" & lng结帐ID)
        Else
            Call ReportOpen(gcnOracle, lngSys, strReprotName, Me)
        End If
    End With
End Sub
Private Sub txtPatient_Change()
    txtPatient.Tag = ""
    If txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    IDKind.SetAutoReadCard (txtPatient.Text = "")
    stbThis.Panels(2).Text = ""
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    If txtPatient.Locked Then Exit Sub
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "")
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
    On Error GoTo errH
    
    If txtPatient.Locked Then Exit Sub
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
'        With frmPatiSelect
'            If (mbytUseType = 0 Or mbytUseType = 1) Then
'                .mlngUnitID = mlngUnitID
'            Else
'                .mlngUnitID = mlngDeptID
'            End If
'            .mbytUseType = mbytUseType
'            .mstrPrivs = mstrPrivs
'            Set .mfrmParent = Me
'            .Show 1, Me
'        End With
    Else
        If IDKind.GetCurCard.名称 Like "姓名*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        ElseIf IDKind.IDKind = IDKind.GetKindIndex("门诊号") Or IDKind.IDKind = IDKind.GetKindIndex("住院号") Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            End If
         Else
            txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
            '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
            txtPatient.IMEMode = 0
        End If
    End If
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        ElseIf IsNumeric(txtPatient.Tag) And mrsInfo.State = 1 Then
            KeyAscii = 0
            '刷新病人信息:"-病人ID"
            Call GetPatient(IDKind.GetCurCard, txtPatient.Tag, False)
            If mrsInfo.State = 0 Then   '
                txtPatient.Text = "": txtOld.Text = ""
                txtSex.Text = "": txt住院号.Text = ""
                Exit Sub
            End If
            Call ReadListData
            Exit Sub
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, txtPatient.Text)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnICCard As Boolean, blnMsg As Boolean, blnIDCard As Boolean
    
   '54899
    If objCard.名称 Like "IC卡*" And objCard.系统 = True And mstrPassWord <> "" Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.名称 Like "*身份证*" And objCard.系统 = True And mstrPassWord <> "" Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    
    If Not GetPatient(objCard, strInput, blnCard, blnMsg) Then
        If blnCard Then
            If Not blnMsg Then MsgBox "不能确定病人信息，请检查是否正确刷卡！", vbInformation, gstrSysName
            txtPatient.Text = "": txtOld.Text = ""
            txt住院号.Text = ""
            vsFee.Clear 1
            vsFee.Rows = 2
            Exit Sub
        End If
        If Not blnMsg Then MsgBox "不能读取病人信息！", vbInformation, gstrSysName
        zlControl.TxtSelAll txtPatient
        txtOld.Text = "": txtSex.Text = "": txt住院号.Text = ""
        vsFee.Clear 1
        vsFee.Rows = 2
        Exit Sub
    End If
    
    '读取成功
    '就诊卡密码检查
     If (objCard.名称 Like "IC卡*" Or objCard.名称 Like "*身份证*") And objCard.系统 = True And blnCard Then blnCard = False
     If Mid(gstrCardPass, 6, 1) = "1" And (blnCard Or blnICCard Or blnIDCard) Then
        If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!姓名, mrsInfo!性别, "" & mrsInfo!年龄) Then
            vsFee.Clear 1
            vsFee.Rows = 2
            Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
        End If
    End If
    Call ReadListData
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub
Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub
Private Sub txtPatient_Validate(Cancel As Boolean)
    If IsNumeric(txtPatient.Tag) And mrsInfo.State = 1 Then
        mblnValid = True
        Call txtPatient_KeyPress(13)
        mblnValid = False
    End If
End Sub
Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnOutMsg As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:blnCard=是否就诊卡刷卡
    '出参: blnOutMsg-已经提示,不用再外部再提示
    '返回:
    '编制:刘兴洪
    '日期:2011-01-25 16:57:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim strIF As String, dtDate As Date, vRect As RECT
    Dim strPati As String, blnCancel As Boolean
    
    On Error GoTo errH
    '审核标志:50459
    strSQL = _
        "Select A.病人ID,Nvl(B.主页ID,0) as 主页ID,A.门诊号 as 门诊号,A.当前床号,B.出院病床," & _
        "       Nvl(b.姓名, a.姓名) As 姓名, Nvl(b.性别, a.性别) As 性别,Nvl(B.年龄,A.年龄) as 年龄,A.IC卡号,A.就诊卡号,A.卡验证码," & _
        "       Nvl(B.费别,A.费别) as 费别,C.名称 as 当前科室,A.当前科室ID,D.名称 as 出院科室,B.出院科室ID, A.险类 as 险类,E.卡号,E.医保号,E.密码," & _
        "       A.登记时间,Nvl(B.状态,0) as 状态,Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式,Nvl(B.审核标志,0) as 审核标志,B.入院日期,B.出院日期,B.病人性质,B.病人类型" & _
        " From 病人信息 A,病案主页 B,部门表 C,部门表 D,医保病人档案 E,医保病人关联表 F" & _
        " Where A.停用时间 is NULL And A.病人ID=B.病人ID(+) And A.主页ID=B.主页ID(+) " & _
        "           And A.病人ID=F.病人ID(+) And F.标志(+)=1 And F.医保号=E.医保号(+) And F.险类=E.险类(+) And F.中心 = E.中心(+)" & _
        "           And A.当前科室ID=C.ID(+) And B.出院科室ID=D.ID(+)" & _
        "           And A.停用时间 is NULL "
    
    If blnCard = True And objCard.名称 Like "姓名*" Then  '刷卡
        lng卡类别ID = IDKind.GetDefaultCardTypeID
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.病人ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        strSQL = strSQL & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号(病人在院)
        strSQL = strSQL & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strSQL = strSQL & " And A.门诊号=[1]"
    Else '当作姓名
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                If mrsInfo.State = 1 Then
                    If mrsInfo!姓名 = Trim(txtPatient.Text) Then GetPatient = True: Exit Function
                End If
                '53816
                '通过姓名查找
                strPati = "Select A.病人ID as ID,A.病人ID,A.住院号, A.门诊号, Nvl(b.性别, a.性别) as 性别, A.年龄, A.住院次数, A.家庭地址, A.工作单位," & vbNewLine & _
                        "To_Char(A.出生日期,'YYYY-MM-DD') as 出生日期,  To_Char(B.入院日期,'YYYY-MM-DD') as 入院日期, To_Char(B.出院日期,'YYYY-MM-DD') as 出院日期" & vbNewLine & _
                        "From 病人信息 A, 病案主页 B" & vbNewLine & _
                        "Where A.病人id = B.病人id(+) And A.主页ID = B.主页id(+) And A.停用时间 Is Null And A.姓名 = [1] " & vbNewLine & strPati & vbNewLine & _
                        "Order By Decode(住院号, Null, 1, 0), 入院日期 Desc"
                        
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput)
                            
                If mrsInfo Is Nothing Or blnCancel Then
                    txtPatient.Text = "": txtOld.Text = "": txtSex.Text = ""
                    txt住院号.Text = ""
                    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
                    Set mrsInfo = New ADODB.Recordset
                    Exit Function
                End If
                strInput = "-" & Val(mrsInfo!病人ID)
                strSQL = strSQL & " And A.病人ID=[1]"
                    
            Case "医保号"
                strInput = UCase(strInput)
                strSQL = strSQL & " And A.医保号=[2]"
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.门诊号=[2]"
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [2])"
            Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strSQL = strSQL & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
    txtPatient.ForeColor = Me.ForeColor
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    
    If Not mrsInfo.EOF Then
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!卡验证码)
        txtPatient.PasswordChar = ""
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
        txtPatient.ForeColor = zlDatabase.GetPatiColor(Nvl(mrsInfo!病人类型))
        txtPatient.Text = Nvl(mrsInfo!姓名): txtOld.Text = Nvl(mrsInfo!年龄): txtSex.Text = Nvl(mrsInfo!性别)
        txt住院号.Text = Nvl(mrsInfo!门诊号)
        If IsNull(mrsInfo!入院日期) Then
            
            dtDate = zlDatabase.Currentdate '53816
            dtpEnd.MaxDate = Format(dtDate, "yyyy-mm-dd 23:59:59")
            dtpEnd.Value = dtDate
        Else
            dtDate = CDate(Format(mrsInfo!入院日期, "yyyy-mm-dd HH:MM:SS"))
            If dtDate > dtpEnd.MaxDate Then dtpEnd.MaxDate = dtDate
            
            dtpEnd.Value = Format(mrsInfo!入院日期, "yyyy-mm-dd HH:MM:SS")
            dtpEnd.MaxDate = dtpEnd.Value + 1
            dtpBegin.MaxDate = dtpEnd.Value
            '   问题:36609 比入院时间要多一天,因为可能存在病人在没有门诊结算时,先入院,再去门诊结算,从而造成门诊费用转不了的情况.
        End If
    
        If dtpBegin.Value > dtpEnd.Value Then
            dtpBegin.Value = dtpEnd.Value - 7   '减去7天
        End If
        GetPatient = True
        Exit Function
    Else
        txtPatient.Text = "": txtOld.Text = "": txtSex.Text = ""
        txt住院号.Text = ""
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Set mrsInfo = New ADODB.Recordset
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
NotFoundPati:
    txtPatient.Text = "": txtOld.Text = "": txtSex.Text = ""
    txt住院号.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Set mrsInfo = New ADODB.Recordset
End Function

Private Function ReadListData(Optional blnFilter As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取需要审核的明细数据
    '返回:读取成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2011-01-25 17:10:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, strTable As String, rsTemp As ADODB.Recordset
    Dim lngCol As Long, strSQL As String, lngRow As Long
    Dim strFilter As String, dtEndDate As Date, dtStartDate As Date
    Dim strWhere As String, strInsure As String
    
    dtEndDate = dtpEnd.Value: dtStartDate = dtpBegin.Value
    If mrsInfo Is Nothing Then
        lng病人ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng病人ID = 0
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
    End If
    
    If dtEndDate - dtStartDate > 30 Then    '大于30天,则按病人ID为过滤条件
        strWhere = " And A.病人ID=[1] And (A.发生时间+0 between [2] and [3] )"
        strInsure = " And 病人ID=[1] And (结算时间+0 between [2] and [3] )"
    Else
        strWhere = " And A.病人ID+0=[1] And (A.发生时间 between [2] and [3] )"
        strInsure = " And 病人ID+0=[1] And (结算时间 between [2] and [3] )"
    End If
    
    On Error GoTo errHandle
    If blnFilter = False Then zlCommFun.ShowFlash "正在读取收费单据,请稍候 ..."
    Screen.MousePointer = 11
    DoEvents
    Me.Refresh
    
    strTable = "" & _
    "       Select Nvl(Max(记录状态), 0) As 转出标志, Decode(Max(险类), 0, '', '√') As 医保," & vbNewLine & _
    "              Max(Decode(价格父号, Null, Decode(记录状态, 2, To_Number(Null), ID), To_Number(Null))) As ID, '收费单' As 单据, NO, 实际票号," & vbNewLine & _
    "              序号, 收费类别, 从属父号, 收费细目id, 执行部门id, Avg(Nvl(付数, 1)) As 付数, Sum(数次) 数次, 标准单价 As 单价, Sum(应收金额) As 应收金额," & vbNewLine & _
    "              Sum(实收金额) As 实收金额, 开单人, To_Char(发生时间, 'YYYY-MM-DD HH24:MI:SS') As 发生时间, Max(险类) As 险类, Max(审核人) As 审核人," & vbNewLine & _
    "              To_Char(Max(审核日期), 'YYYY-MM-DD HH24:MI:SS') As 审核日期, Min(结帐id) As 结帐id" & vbNewLine & _
    "       From (Select a.Id, m.记录状态, Nvl(b.险类, 0) As 险类, a.价格父号, a.No, a.实际票号, a.序号, a.收费类别, a.从属父号, a.收费细目id, a.执行部门id, a.付数," & vbNewLine & _
    "                     a.数次, a.标准单价, a.应收金额, a.实收金额, a.开单人, a.发生时间, m.审核人, m.审核日期, a.结帐id" & vbNewLine & _
    "              From 门诊费用记录 A," & vbNewLine & _
    "                   (Select Distinct 记录id, 险类" & vbNewLine & _
    "                     From 保险结算记录" & vbNewLine & _
    "                     Where 性质 = 1" & strInsure & ") B, 费用审核记录 M" & vbNewLine & _
    "              Where Mod(a.记录性质, 10) = 1 And a.记录状态 <> 0 " & strWhere & vbNewLine & _
    "                    And a.结帐id = b.记录id(+) And a.Id = m.费用id(+) And" & vbNewLine & _
    "                    m.性质(+) = 1 And Nvl(a.附加标志, 0) <> 9)" & vbNewLine & _
    "       Group By NO, 实际票号, 序号, 标准单价, 收费类别, 收费细目id, 从属父号, 执行部门id, 开单人, 发生时间" & vbNewLine & _
    "       Having Sum(数次) <> 0"

    
    strTable = strTable & "Union ALL " & _
    "       Select Nvl(Max(记录状态), 0) As 转出标志, Decode(Max(险类), 0, '', '√') As 医保," & vbNewLine & _
    "              Max(Decode(价格父号, Null, Decode(记录状态, 2, To_Number(Null), ID), To_Number(Null))) As ID, '收费单' As 单据, NO, 实际票号," & vbNewLine & _
    "              序号, 收费类别, 从属父号, 收费细目id, 执行部门id, Avg(Nvl(付数, 1)) As 付数, Sum(数次) 数次, 标准单价 As 单价, Max(应收金额) As 应收金额," & vbNewLine & _
    "              Max(实收金额) As 实收金额, 开单人, To_Char(发生时间, 'YYYY-MM-DD HH24:MI:SS') As 发生时间, Max(险类) As 险类, Max(审核人) As 审核人," & vbNewLine & _
    "              To_Char(Max(审核日期), 'YYYY-MM-DD HH24:MI:SS') As 审核日期, Min(结帐id) As 结帐id" & vbNewLine & _
    "       From (Select a.Id, m.记录状态, Nvl(b.险类, 0) As 险类, a.价格父号, a.No, a.实际票号, a.序号, a.收费类别, a.从属父号, a.收费细目id, a.执行部门id, a.付数," & vbNewLine & _
    "                     a.数次, a.标准单价, a.应收金额, a.实收金额, a.开单人, a.发生时间, m.审核人, m.审核日期, a.结帐id, a.结帐金额" & vbNewLine & _
    "              From 门诊费用记录 A," & vbNewLine & _
    "                   (Select Distinct 记录id, 险类" & vbNewLine & _
    "                     From 保险结算记录" & vbNewLine & _
    "                     Where 性质 = 1" & strInsure & ") B, 费用审核记录 M" & vbNewLine & _
    "              Where Mod(a.记录性质, 10) = 1 And a.记录状态 <> 0 " & strWhere & vbNewLine & _
    "                    And a.结帐id = b.记录id(+) And Exists" & vbNewLine & _
    "               (Select 1" & vbNewLine & _
    "                     From 门诊费用记录 J, 费用审核记录 K" & vbNewLine & _
    "                     Where j.Id = k.费用id And j.病人id = [1] And k.性质 = 1 And j.No = a.No And j.序号 = a.序号 And" & vbNewLine & _
    "                           Mod(j.记录性质, 10) = 1) And a.Id = m.费用id(+) And m.性质(+) = 1 And Nvl(a.附加标志, 0) <> 9)" & vbNewLine & _
    "       Group By NO, 实际票号, 序号, 标准单价, 收费类别, 收费细目id, 从属父号, 执行部门id, 开单人, 发生时间" & vbNewLine & _
    "       Having Sum(数次) = 0"

    strTable = strTable & "Union ALL " & _
    " Select  nvl(max(M.记录状态),0) as 转出标志,Decode(NULL,Null,'','√') as 医保,Max(decode(A.价格父号,NULL,ID,0))  as ID, " & _
    "       '记帐单' as 单据,A.No,A.实际票号, A.序号 as 序号,A.收费类别,A.从属父号,A.收费细目ID,A.执行部门ID, " & _
    "       Avg(Nvl(A.付数,1)) as 付数, Sum(A.数次) 数次, A.标准单价 as 单价,Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额," & _
    "       A.开单人,To_Char(A.发生时间, 'YYYY-MM-DD HH24:MI:SS') As 发生时间, NULL as 险类," & vbNewLine & _
    "       Max(M.审核人) as 审核人,To_Char(Max(M.审核日期), 'YYYY-MM-DD HH24:MI:SS') As 审核日期,Null As 结帐ID " & vbNewLine & _
    "From 门诊费用记录 A,费用审核记录 M" & vbNewLine & _
    "Where A.记录性质 = 2 And A.记录状态 <> 0 " & strWhere & _
    "           And A.ID = M.费用ID(+) And M.性质(+)=1 " & vbNewLine & _
    "Group By A.NO, A.实际票号,A.序号,A.收费类别,A.收费细目ID,A.从属父号,A.标准单价,A.执行部门id," & _
    "       A.开单人, A.发生时间 Having Sum(A.数次) <> 0" & vbNewLine
    
    strTable = strTable & "Union ALL " & _
    " Select  nvl(max(M.记录状态),0) as 转出标志,Decode(NULL,Null,'','√') as 医保,Max(decode(A.价格父号,NULL,ID,0))  as ID, " & _
    "       '记帐单' as 单据,A.No,A.实际票号, A.序号 as 序号,A.收费类别,A.从属父号,A.收费细目ID,A.执行部门ID, " & _
    "       Avg(Nvl(A.付数,1)) as 付数, Sum(A.数次) 数次, A.标准单价 as 单价,Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额," & _
    "       A.开单人,To_Char(A.发生时间, 'YYYY-MM-DD HH24:MI:SS') As 发生时间, NULL as 险类," & vbNewLine & _
    "       Max(M.审核人) as 审核人,To_Char(Max(M.审核日期), 'YYYY-MM-DD HH24:MI:SS') As 审核日期,Null As 结帐ID " & vbNewLine & _
    "From 门诊费用记录 A,费用审核记录 M" & vbNewLine & _
    "Where A.记录性质 = 2 And A.记录状态 <> 0 " & strWhere & _
    "           And A.ID = M.费用ID(+) And M.性质(+)=1 And Exists" & vbNewLine & _
    "               (Select 1" & vbNewLine & _
    "                     From 门诊费用记录 J, 费用审核记录 K" & vbNewLine & _
    "                     Where j.Id = k.费用id And j.病人id = [1] And k.性质 = 1 And j.No = a.No And j.序号 = a.序号 And" & vbNewLine & _
    "                           j.记录性质 = 2) And a.Id = m.费用id(+) And m.性质(+) = 1" & vbNewLine & _
    "Group By A.NO, A.实际票号,A.序号,A.收费类别,A.收费细目ID,A.从属父号,A.标准单价,A.执行部门id," & _
    "       A.开单人, A.发生时间 Having Sum(A.数次) = 0" & vbNewLine
    
    
    strSQL = "" & _
    " Select  A.ID,A.转出标志,decode(A.审核人,NULL,0,-1) as 审核,A.单据,A.No as 单据号,A.实际票号 As 票据号, " & _
    "       A.序号,A.从属父号,A.收费细目ID,A.执行部门ID,A.收费类别,P.类别, " & _
    "       C.编码 as 编码,C.编码||'-'||Nvl(B.名称,C.名称) as 名称,E1.名称 as 商品名,C.规格," & _
    "       A.付数, A.数次,C.计算单位," & _
    "       ltrim(to_char(A.单价,'9999990.00000')) as 单价," & _
    "       ltrim(to_char(A.应收金额,'9999990.00')) as 应收金额," & _
    "       ltrim(to_char(A.实收金额,'9999990.00')) as 实收金额," & _
    "       A.开单人,A.发生时间,A.医保, A.险类,A.审核人,A.审核日期,A.结帐ID" & vbNewLine & _
    "From (" & strTable & ") A,收费项目目录 C,收费项目别名 B,收费项目别名 E1,收费类别 P" & _
    " Where A.收费细目ID=C.ID And A.收费细目ID=B.收费细目ID(+)  And B.码类(+)=1 And B.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
    "       and A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
    "       And A.收费类别=P.编码(+)" & _
    " Order by A.单据,A.NO,A.序号"

    If mrsFeeList Is Nothing Or blnFilter = False Then
        Set mrsFeeList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, dtStartDate, dtEndDate)
    Else
        mrsFeeList.Filter = 0
    End If
    strFilter = IIf(cbo收费单.ItemData(cbo收费单.ListIndex) = 2, "", " And 单据='" & cbo收费单.Text & "'")
    strFilter = strFilter & IIf(chk审核.Value = 1, "", " And  审核=0")
    strFilter = strFilter & IIf(chk已转出费用.Value = 1, "", " And  转出标志=0")
    mrsFeeList.Filter = Mid(strFilter, 5)
    vsFee.Redraw = flexRDNone
    mblnNotClick = True
    vsFee.Clear: vsFee.Cols = 1: vsFee.Rows = 2: vsFee.FixedRows = 1
    mblnNotClick = False
    Set vsFee.DataSource = mrsFeeList
    If vsFee.Rows <= 1 Then vsFee.Rows = 2
    With vsFee
        For lngCol = 0 To .Cols - 1
             .ColAlignment(lngCol) = flexAlignLeftCenter
             .FixedAlignment(lngCol) = flexAlignCenterCenter
              .ColKey(lngCol) = Trim(.TextMatrix(0, lngCol))
              If .ColKey(lngCol) Like "*ID" Or InStr(1, ",险类,编码,序号,从属父号,转出标志,收费类别,结帐ID,", "," & .ColKey(lngCol) & ",") > 0 Then
                    .ColHidden(lngCol) = True
              ElseIf .ColKey(lngCol) Like "*数*" Or .ColKey(lngCol) Like "*价*" Or .ColKey(lngCol) Like "*额" Then
                    .ColAlignment(lngCol) = flexAlignRightCenter
              End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .ColDataType(.ColIndex("审核")) = flexDTBoolean
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsFee, Me.Caption, "明细列表", True
        If gTy_System_Para.byt药品名称显示 <> 2 Then    '0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
            .ColHidden(.ColIndex("商品名")) = True
        End If
        '画线
        Dim strNO As String, str单据 As String
        
        For lngRow = 1 To .Rows - 1
            If strNO <> Trim(.TextMatrix(lngRow, .ColIndex("单据号"))) _
                And str单据 = Trim(.TextMatrix(lngRow, .ColIndex("单据"))) And strNO <> "" Then
                '画出分隔线
                .CellBorderRange lngRow, .FixedCols, lngRow, .Cols - 1, vbBlue, 0, 1, 0, 0, 0, 0
            End If
            If str单据 <> Trim(.TextMatrix(lngRow, .ColIndex("单据"))) And str单据 <> "" Then
                .CellBorderRange lngRow, .FixedCols, lngRow, .Cols - 1, vbRed, 0, 1, 0, 0, 0, 0
            End If
            .Cell(flexcpData, lngRow, .ColIndex("单据号")) = .TextMatrix(lngRow, .ColIndex("单据号"))
            .Cell(flexcpData, lngRow, .ColIndex("审核")) = Val(.TextMatrix(lngRow, .ColIndex("审核")))
            If Val(.TextMatrix(lngRow, .ColIndex("审核"))) <> 0 Then
                Select Case Val(.TextMatrix(lngRow, .ColIndex("转出标志")))
                Case 0
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &HFF0000       '蓝色
                Case 1, 2
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H80000015
'                Case 2
'                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H80000010
                End Select
            End If
            strNO = Trim(.TextMatrix(lngRow, .ColIndex("单据号")))
            str单据 = Trim(.TextMatrix(lngRow, .ColIndex("单据")))
        Next
        .Editable = flexEDKbdMouse
    End With
    If blnFilter = False Then zlCommFun.StopFlash
    Call SetSumMoney
    Call StatusShowBillSum
    vsFee.Redraw = flexRDDirect
    Screen.MousePointer = 0
    ReadListData = True
    Exit Function
errHandle:
    vsFee.Redraw = flexRDBuffered
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
   If blnFilter = False Then zlCommFun.StopFlash
End Function
Private Sub vsFee_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsFee
        Select Case Col
        Case .ColIndex("审核")
                SetNOBill .TextMatrix(Row, .ColIndex("单据")), .TextMatrix(Row, .ColIndex("单据号")), Val(.TextMatrix(Row, .Col)) <> 0
                Call SetRowSelected(Row)
                mblnChange = True
                Call SetSumMoney
        Case Else
        End Select
    End With
End Sub
Private Sub vsFee_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsFee, Me.Caption, "明细列表", True
End Sub

Private Sub vsFee_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnNotClick = True Then Exit Sub
    If OldRow <> NewRow Then
        Call StatusShowBillSum
    End If
End Sub

Private Sub vsFee_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsFee, Me.Caption, "明细列表", True
End Sub

Private Sub vsFee_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsFee
        Select Case Col
        Case .ColIndex("审核")
            If Val(.TextMatrix(Row, .ColIndex("转出标志"))) <> 0 Then
                stbThis.Panels(2).Text = "门诊费用已经转出,不能更改审核状态"
                Cancel = True: Exit Sub
            End If
            
            If GetVsGridBoolColVal(vsFee, Row, Col) Then
                If InStr(1, mstrPrivs, ";取消他人审核;") = 0 And .TextMatrix(Row, .ColIndex("审核人")) <> UserInfo.姓名 And .TextMatrix(Row, .ColIndex("审核人")) <> "" Then
                    stbThis.Panels(2).Text = "你没有权限取消他人审核的费用"
                    Cancel = True: Exit Sub
                End If
            End If
            If CheckIsInput(Row) = False Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsFee_DblClick()
     '   vsFee.TextMatrix(vsFee.Row, vsFee.ColIndex("审核")) = "√"
End Sub

Private Sub vsFee_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按单据画线和清除线
    '编制:刘兴洪
    '日期:2011-01-26 09:57:32
    '说明:
    '       1.OwnerDraw要设置为Over(画出单元所有内容)
    '       2.Cell的GridLine从上下左右向内都是从第1根线开始
    '       3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    Dim strText As String
    strText = " "
    With vsFee
        '擦除相关行列的边线及内容
        lngLeft = .ColIndex("审核"): lngRight = .ColIndex("审核")
        
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        Call GetBillNOStartAndEndRow(Row, lngBegin, lngEnd)
        If lngBegin = lngEnd Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '底行保留下边线
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, strText, 1, 0
        Done = True
    End With
End Sub
Private Sub GetBillNOStartAndEndRow(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取单据行
    '编制:刘兴洪
    '日期:2011-01-26 10:01:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    lngBegin = lngRow: lngEnd = lngRow
    With vsFee
        For i = lngRow - 1 To .FixedRows Step -1
            If .TextMatrix(i, .ColIndex("单据号")) = .TextMatrix(lngRow, .ColIndex("单据号")) Then
                lngBegin = i
            Else
                Exit For
            End If
        Next
        For i = lngRow To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) = .TextMatrix(lngRow, .ColIndex("单据号")) Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
    End With
End Sub
Private Function SetNOBill(ByVal str单据 As String, ByVal strNO As String, ByVal blnSel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按单据全选或全清单据
    '入参:str单据-单据类型(收费单,记帐单)
    '       strNO-指定的NO
    '        blnSel:true表示全选,否则全清
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2011-01-24 10:47:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsFee
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> "" _
                And .TextMatrix(i, .ColIndex("单据号")) = strNO _
                And .TextMatrix(i, .ColIndex("单据")) = str单据 Then
                .TextMatrix(i, .ColIndex("审核")) = IIf(blnSel, -1, 0)
            End If
        Next
    End With
    SetNOBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存审核数据
    '返回:审核成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-01-26 13:31:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim cllProc As Collection, cllTemp As Collection, i As Long
    Dim blnSel As Boolean, lngRow As Long, strDate As String
    If mrsInfo Is Nothing Or mrsInfo.State = 0 Then Exit Function
    Set cllProc = New Collection: Set cllTemp = New Collection
    If mrsInfo Is Nothing Or mrsInfo.State = 0 Then Exit Function
    If Val(Nvl(mrsInfo!主页ID)) <> 0 Then
        If zlIsAllowFeeChange(Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID))) = False Then
            Exit Function
        End If
    End If
    
    '先处理取消审核部分
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    With vsFee
        If .Rows <= 1 Then Exit Function
        If .Cols <= 1 Then Exit Function
        
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("ID"))) <> 0 Then
                blnSel = GetVsGridBoolColVal(vsFee, lngRow, .ColIndex("审核"))
                If Val(.Cell(flexcpData, lngRow, .ColIndex("审核"))) <> 0 And Not blnSel Then
                    ' Zl_费用审核记录_Delete
                    strSQL = "Zl_费用审核记录_Delete("
                    '  费用id_In In 费用审核记录.费用id%Type,
                    strSQL = strSQL & "" & Val(.TextMatrix(lngRow, .ColIndex("ID"))) & ","
                    '  性质_In   In 费用审核记录.性质%Type
                    strSQL = strSQL & "1)"
                    zlAddArray cllProc, strSQL
                ElseIf Val(.Cell(flexcpData, lngRow, .ColIndex("审核"))) = 0 And blnSel Then
                    '插入
                    'Zl_费用审核记录_Insert
                    strSQL = "Zl_费用审核记录_Insert("
                    '  性质_In     In 费用审核记录.性质%Type,
                    strSQL = strSQL & "" & 1 & ","
                    '  费用id_In   In 费用审核记录.费用id%Type,
                    strSQL = strSQL & "" & Val(.TextMatrix(lngRow, .ColIndex("ID"))) & ","
                    '  病人id_In   In 费用审核记录.病人id%Type,
                    strSQL = strSQL & "" & Val(Nvl(mrsInfo!病人ID)) & ","
                    '  主页id_In   In 费用审核记录.主页id%Type,
                    strSQL = strSQL & "" & IIf(Val(Nvl(mrsInfo!主页ID)) = 0, "Null", Val(Nvl(mrsInfo!主页ID))) & ","
                    '  审核人_In   In 费用审核记录.审核人%Type,
                    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                    '  审核日期_In In 费用审核记录.审核日期%Type
                    strSQL = strSQL & "to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'))"
                    zlAddArray cllTemp, strSQL
                End If
            End If
        Next
    End With
    If cllTemp.Count = 0 And cllProc.Count = 0 Then
        MsgBox "未选择相关的单据项目,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    For i = 1 To cllTemp.Count
        zlAddArray cllProc, cllTemp(i)
    Next
    On Error GoTo errHandle
    zlExecuteProcedureArrAy cllProc, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetLocaleNO(ByVal str单据 As String, ByVal strNO As String, ByVal blnSelect As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置指定的NO
    '编制:刘兴洪
    '日期:2011-02-09 14:56:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    With vsFee
        For lngRow = 1 To .Rows - 1
            If Trim(.TextMatrix(lngRow, .ColIndex("单据号"))) = strNO _
                And Trim(.TextMatrix(lngRow, .ColIndex("单据"))) = str单据 Then
                    .TextMatrix(lngRow, .ColIndex("选择")) = IIf(blnSelect, -1, 0)
            End If
        Next
    End With
End Sub
Private Function CheckIsInput(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否允许输入更改
    '入参:lngRow-指定的行
    '出参:
    '返回:有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-09 15:04:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer, strNO As String, i As Long, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant
    Dim lng病人ID As Long, str单据 As String
    
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State <> 1 Then Exit Function
    lng病人ID = Val(Nvl(mrsInfo!病人ID))
    With vsFee
            intInsure = Val(.TextMatrix(lngRow, .ColIndex("险类")))
            strNO = .TextMatrix(lngRow, .ColIndex("单据号"))
            str单据 = .TextMatrix(lngRow, .ColIndex("单据"))
            If intInsure > 0 And str单据 = "收费单" Then
                If Not gclsInsure.GetCapability(support门诊结算作废, lng病人ID, intInsure) Then
                    stbThis.Panels(2).Text = "单据[" & strNO & "]的病人险类不支持门诊结算作废,此行不允许选择转入!"
                    Exit Function
                Else
                    '再判断该单据的每种结算方式是否支持,正常退费时,可以退为指定结算方式,此处简化规则为不允许退费
                    strTmp = GetBalanceType(strNO)
                    If strTmp <> "" Then
                        arrBalanceType = Split(strTmp, ",")
                        For i = 0 To UBound(arrBalanceType)
                            strBalanceType = arrBalanceType(i)
                            If Not gclsInsure.GetCapability(support门诊结算作废, lng病人ID, intInsure, strBalanceType) Then
                                stbThis.Panels(2).Text = "单据[" & strNO & "]的病人险类不支持" & strBalanceType & "作废,此行不允许选择转入!"
                                Exit Function
                            End If
                        Next
                    End If
                End If
            End If
    End With
    CheckIsInput = True
End Function
Private Function SetRowSelected(ByVal lngRow As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置一行的选择状态
    '       如果是多张单据中的一张,则还需同时设置多张中的其它单据
    '编制:刘兴洪
    '日期:2011-02-09 14:50:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer, strNO As String, i As Long, strTmp As String
    Dim blnSelect As Boolean, lng病人ID As Long, str单据 As String
    lng病人ID = 0
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            lng病人ID = Val(Nvl(mrsInfo!病人ID))
        End If
    End If
    With vsFee
        intInsure = Val(.TextMatrix(lngRow, .ColIndex("险类")))
        blnSelect = GetVsGridBoolColVal(vsFee, lngRow, .ColIndex("审核"))
        str单据 = Trim(.TextMatrix(lngRow, .ColIndex("单据")))
        If intInsure > 0 And str单据 = "收费单" Then '全部选择或取消
            If gclsInsure.GetCapability(support多单据收费必须全退, lng病人ID, intInsure) Or Not IsYBSingle(.TextMatrix(lngRow, .ColIndex("单据号")), intInsure) Then
                If Not SetMultiOther(lngRow, blnSelect, intInsure) Then Exit Function
            End If
        Else '现金病人需要处理多单据收费情况
            If Not SetMultiOther(lngRow, blnSelect, intInsure) Then Exit Function
        End If
    End With
    SetRowSelected = True
End Function

Private Function IsYBSingle(ByVal strNO As String, ByVal intInsure As Integer) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset, blnInsureSingle As Boolean
    
    blnInsureSingle = gclsInsure.GetCapability(83, , intInsure)
    If blnInsureSingle = False Then
        IsYBSingle = False
        Exit Function
    Else
        strSQL = "Select 1 From 医保结算明细 Where NO = [1] And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
        If rsTmp.EOF Then
            IsYBSingle = False
        Else
            If CheckAllTurn(strNO) Then
                IsYBSingle = False
            Else
                IsYBSingle = True
            End If
        End If
    End If
    
End Function

Private Function GetBalanceType(ByVal strNO As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取一张单据中的医保结算方式串
    '返回:医保结算方式串
    '编制:刘兴洪
    '日期:2011-02-09 15:01:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long
    On Error GoTo errH
    strSQL = "Select A.结算方式 From 病人预交记录 A, 结算方式 B" & vbNewLine & _
            "Where A.结算方式 = B.名称 And B.性质 In (3, 4) And A.NO = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    
    For i = 1 To rsTmp.RecordCount
        GetBalanceType = GetBalanceType & "," & rsTmp!结算方式
        rsTmp.MoveNext
    Next
    GetBalanceType = Mid(GetBalanceType, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckAllTurn(ByVal strNO As String) As Boolean
    Dim strSQL As String, rsData As ADODB.Recordset
    strSQL = "Select 1" & vbNewLine & _
            " From 病人预交记录 A," & vbNewLine & _
            "     (Select Distinct 结帐id" & vbNewLine & _
            "       From 门诊费用记录" & vbNewLine & _
            "       Where NO In (Select Distinct NO" & vbNewLine & _
            "                    From 门诊费用记录" & vbNewLine & _
            "                    Where 结帐id In" & vbNewLine & _
            "                          (Select 结帐id" & vbNewLine & _
            "                           From 病人预交记录" & vbNewLine & _
            "                           Where 结算序号 In (Select b.结算序号" & vbNewLine & _
            "                                          From 门诊费用记录 A, 病人预交记录 B" & vbNewLine & _
            "                                          Where a.No = [1] And a.记录性质 = 1 And a.记录状态 <> 0 And a.结帐id = b.结帐id))) And" & vbNewLine & _
            "             记录性质 = 1 And 记录状态 <> 0) B" & vbNewLine & _
            " Where a.结帐id = b.结帐id And a.记录性质 = 3 And (Exists (Select 1 From 医疗卡类别 Where ID = a.卡类别id And 是否全退 = 1) Or Exists" & vbNewLine & _
            "       (Select 1 From 消费卡类别目录 Where 编号 = a.结算卡序号 And 是否全退 = 1))" & vbNewLine & _
            " Group By 结算方式" & vbNewLine & _
            " Having Sum(冲预交) <> 0"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsData.EOF Then
        CheckAllTurn = False
    Else
        CheckAllTurn = True
    End If
End Function

Private Function SetMultiOther(ByVal lngRow As Long, blnSelect As Boolean, intInsure As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:多张单据整体选择或取消
    '       如果医保多张单据要求整体退费,选择其中一张时,全选多张,取消时全取消
    '入参:lngRow-当前行
    '        blnSelect-是否选中
    '        intInsure-险类
    '返回:
    '编制:刘兴洪
    '日期:2011-02-09 15:41:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, k As Long, strNO As String, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant
    Dim lng病人ID As Long, str单据 As String, blnAllTurn As Boolean
    lng病人ID = 0
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            lng病人ID = Val(Nvl(mrsInfo!病人ID))
        End If
    End If
    With vsFee
        str单据 = Trim(.TextMatrix(lngRow, .ColIndex("单据")))
        If intInsure = 0 Then
            If CheckAllTurn(.TextMatrix(lngRow, .ColIndex("单据号"))) = True Then
                blnAllTurn = True
            Else
                blnAllTurn = False
            End If
            If mblnMultiBalance Or blnAllTurn Then     '   多单据,多种结算方式
                '33635:原因是多单据且多种结算方式,不能部分退
                strNO = ""
                For k = 1 To .Rows - 1
                      If .TextMatrix(k, .ColIndex("结帐ID")) = .TextMatrix(lngRow, .ColIndex("结帐ID")) _
                        And Trim(.TextMatrix(lngRow, .ColIndex("单据号"))) <> "" _
                        And .TextMatrix(k, .ColIndex("单据")) = str单据 Then
                          If InStr(1, "," & strNO & ",", "," & .TextMatrix(k, .ColIndex("单据号")) & ",") = 0 Then
                                strNO = strNO & "," & .TextMatrix(k, .ColIndex("单据号"))
                          End If
                      End If
                Next
                If strNO <> "" Then strNO = Mid(strNO, 2)
                If InStr(1, strNO, ",") > 0 Then    '证明为多单据
                    '一院要求,只要是多单据结算的,在转时,都必须全转
                    'If CheckSingleBalance(strNo) = False Then    '是多种结算方式,则不允许退费,'全选
                        For k = 1 To .Rows - 1
                              If .TextMatrix(k, .ColIndex("结帐ID")) = .TextMatrix(lngRow, .ColIndex("结帐ID")) _
                                  And Trim(.TextMatrix(lngRow, .ColIndex("单据号"))) <> "" _
                                   And .TextMatrix(k, .ColIndex("单据")) = str单据 Then
                                    .TextMatrix(k, .ColIndex("审核")) = IIf(blnSelect, -1, 0)
                              End If
                        Next
                    'End If
                End If
            End If
            '检查是否存在消费卡的结算,如果存在,现不支持这部分数据的处理
            If strNO = "" Then strNO = .TextMatrix(lngRow, .ColIndex("单据号"))
'            If str单据 = "收费单" Then
'                If zlIsExistsSquareCard(strNO) Then
'                    stbThis.Panels(2).Text = "暂不支持对消费卡数据的门诊费用转住院费用!"
'                    For k = 1 To .Rows - 1
'                          If .TextMatrix(k, .ColIndex("单据号")) = .TextMatrix(lngRow, .ColIndex("单据号")) And Trim(.TextMatrix(lngRow, .ColIndex("单据号"))) <> "" Then
'                                .TextMatrix(k, .ColIndex("审核")) = 0
'                          End If
'                    Next
'                End If
'            End If
            '检查是否存在消费卡,如果多单据中存在消费卡,也必须全选
            SetMultiOther = True
            Exit Function
        End If
        If IsYBSingle(vsFee.TextMatrix(lngRow, .ColIndex("单据号")), intInsure) Then SetMultiOther = True: Exit Function
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("结帐ID")) = .TextMatrix(lngRow, .ColIndex("结帐ID")) _
                And i <> lngRow And .TextMatrix(i, .ColIndex("单据")) = str单据 Then
                If GetVsGridBoolColVal(vsFee, i, .ColIndex("审核")) <> GetVsGridBoolColVal(vsFee, lngRow, .ColIndex("审核")) Then
                   If intInsure <> 0 And str单据 = "收费单" And blnSelect Then
                        strNO = .TextMatrix(i, .ColIndex("单据号"))
                        '判断该单据的每种结算方式是否支持,正常退费时,可以退为指定结算方式,此处简化规则为不允许退费
                         strTmp = GetBalanceType(strNO)
                         If strTmp <> "" Then
                             arrBalanceType = Split(strTmp, ",")
                             For j = 0 To UBound(arrBalanceType)
                                 strBalanceType = arrBalanceType(j)
                                 If Not gclsInsure.GetCapability(support门诊结算作废, lng病人ID, intInsure, strBalanceType) Then
                                     stbThis.Panels(2).Text = "单据[" & strNO & "]的病人险类不支持" & strBalanceType & "作废,此行不允许选择转入!"
                                     For k = 1 To .Rows - 1
                                        If .TextMatrix(k, .ColIndex("单据号")) = .TextMatrix(i, .ColIndex("单据号")) _
                                            And .TextMatrix(k, .ColIndex("单据")) = .TextMatrix(i, .ColIndex("单据")) Then
                                            .TextMatrix(k, .ColIndex("审核")) = 0
                                        End If
                                     Next
                                     Exit Function
                                 End If
                             Next
                         End If
                    End If
                    .TextMatrix(i, .ColIndex("审核")) = IIf(blnSelect, -1, 0)
                End If
            End If
        Next
    End With
    SetMultiOther = True
End Function

Private Function CheckSingleBalance(ByVal strNO As String) As Boolean
'功能：判断指定单据中是否只有一种非医保结算方式(冲预交除外)
'       :strNO(格式为"E01,E02"):问题:34035
'参数：
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strNO = Replace(strNO, "'", "")
    CheckSingleBalance = True
    
    strSQL = "" & _
    " Select /*+ rule */ Count(Distinct A.结算方式) num" & vbNewLine & _
    " From 病人预交记录 A, 结算方式 B,Table( f_Str2list([1])) J" & vbNewLine & _
    " Where   A.记录性质 = 3 And A.记录状态 In (1, 3) " & _
    "           And A.结算方式 = B.名称 And B.性质 In (1, 2)  And A.NO = J.Column_Value"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNO)
    If rsTmp!Num > 1 Then CheckSingleBalance = False
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function zlIsExistsSquareCard(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查该单据是否为卡结算单据
    '入参:strNos-单据号(可以为多张,用逗号分离)
    '出参:
    '返回:存在,则返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-11 12:04:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strNoIns As String
    
    On Error GoTo errHandle
    
    strNoIns = Replace(strNos, "'", "")
    strSQL = "" & _
    "   Select /*+ rule */ A.ID As 卡结算id " & _
    "   From 病人卡结算记录 A, 病人预交记录 B,Table( f_Str2list([1])) J " & _
    "   Where A.结算id = B.ID and B.记录性质=3 And B.NO = J.Column_Value And Rownum = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查收费单是否存在刷卡记录", strNoIns)
    zlIsExistsSquareCard = rsTemp.EOF = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetSumMoney(Optional blnCls As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置和显示合计
    '编制:刘兴洪
    '日期:2011-03-04 14:17:20
    '问题:36285
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dblSumMoney As Double
    With vsFee
        If blnCls = False Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("审核"))) <> 0 And _
                  Val(.Cell(flexcpData, i, .ColIndex("审核"))) = 0 Then
                    dblSumMoney = dblSumMoney + Val(.TextMatrix(i, .ColIndex("实收金额")))
                End If
            Next
        Else
            dblSumMoney = 0
        End If
    End With
    lblSum.Caption = "本次审核合计:" & Format(dblSumMoney, "###0.00;-###0.00;0.00;0.00")
End Sub

Public Sub StatusShowBillSum()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据中输入的负数数量及退回科室是否正确
    '编制:刘兴洪
    '日期:2011-03-11 18:09:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, cur金额 As Currency, dbl发票金额 As Double, strNO As String, str发票号 As String
    Dim strTemp As String
    
    With vsFee
        strTemp = "": dbl发票金额 = 0: cur金额 = 0
        If Not (.Row > .Rows - 1 Or .Row < 1) Then
            strNO = .TextMatrix(.Row, .ColIndex("单据号"))
            str发票号 = .TextMatrix(.Row, .ColIndex("票据号"))
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("单据号")) = strNO Then
                        cur金额 = cur金额 + Val(.TextMatrix(i, .ColIndex("实收金额")))
                End If
                If .TextMatrix(i, .ColIndex("票据号")) = str发票号 Then
                        dbl发票金额 = dbl发票金额 + Val(.TextMatrix(i, .ColIndex("实收金额")))
                End If
            Next
            strTemp = "单据(" & strNO & ")合计:" & Format(cur金额, "###0.00;-###0.00;0.00;0.00")
            strTemp = strTemp & "  发票(" & str发票号 & ")合计:" & Format(dbl发票金额, "###0.00;-###0.00;0.00;0.00")
        End If
        stbThis.Panels(2).Text = strTemp
    End With
End Sub
Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结算卡对象的相关信息
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
End Sub

Private Sub zlCreateObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建公共事件对象
    '返回: 创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-08-28 16:16:00
    '说明:
    '问题:54896
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '创建公共对象
    Err = 0: On Error Resume Next
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
         Set mobjICCard.gcnOracle = gcnOracle
    End If
    
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    
End Sub
Private Sub zlCloseObject()
    '关闭相关对象
    Err = 0: On Error Resume Next
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
    End If
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
    End If
    Set mobjIDCard = Nothing
    Set mobjICCard = Nothing
End Sub

