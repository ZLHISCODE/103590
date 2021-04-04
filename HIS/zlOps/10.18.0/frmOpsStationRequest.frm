VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOpsStationRequest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "补录登记"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8730
   Icon            =   "frmOpsStationRequest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picButton 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8715
      TabIndex        =   58
      Top             =   5520
      Width           =   8715
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   7530
         TabIndex        =   61
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   6375
         TabIndex        =   60
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   330
         TabIndex        =   59
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lbl单量单位 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   8460
         TabIndex        =   63
         Top             =   1140
         Width           =   15
      End
      Begin VB.Label lbl总量单位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   6150
         TabIndex        =   62
         Top             =   1140
         Width           =   15
      End
   End
   Begin VB.Frame fra 
      Caption         =   "申请信息"
      Height          =   2445
      Index           =   2
      Left            =   15
      TabIndex        =   43
      Top             =   3075
      Width           =   8700
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   8
         ItemData        =   "frmOpsStationRequest.frx":000C
         Left            =   6450
         List            =   "frmOpsStationRequest.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   1350
         Width           =   2115
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H80000004&
         Caption         =   "紧急(&2)"
         Height          =   225
         Index           =   0
         Left            =   6465
         TabIndex        =   49
         Top             =   2100
         Width           =   945
      End
      Begin VB.CommandButton cmd 
         Caption         =   "…"
         Height          =   285
         Index           =   12
         Left            =   5025
         TabIndex        =   48
         ToolTipText     =   "选择项目(*)"
         Top             =   2040
         Width           =   285
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   7
         ItemData        =   "frmOpsStationRequest.frx":0010
         Left            =   6450
         List            =   "frmOpsStationRequest.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   600
         Width           =   2115
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   12
         Left            =   1110
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   46
         Top             =   2040
         Width           =   3900
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   9
         Left            =   6450
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   1725
         Width           =   2115
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   13
         Left            =   6450
         MaxLength       =   100
         TabIndex        =   44
         Top             =   240
         Width           =   2115
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   6450
         TabIndex        =   51
         Top             =   975
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   81723395
         CurrentDate     =   38022
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfAdvice 
         Height          =   1710
         Left            =   1110
         TabIndex        =   64
         Top             =   255
         Width           =   4200
         _cx             =   7408
         _cy             =   3016
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   6
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "申请手术(&R)"
         Height          =   180
         Index           =   27
         Left            =   105
         TabIndex        =   65
         Top             =   270
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "申请科室(&X)"
         Height          =   180
         Index           =   23
         Left            =   5385
         TabIndex        =   57
         Top             =   1380
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行科室(&Q)"
         Height          =   180
         Index           =   22
         Left            =   5385
         TabIndex        =   56
         Top             =   645
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "麻醉方式(&R)"
         Height          =   180
         Index           =   19
         Left            =   90
         TabIndex        =   55
         Top             =   2100
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "申请医生(&D)"
         Height          =   180
         Index           =   24
         Left            =   5385
         TabIndex        =   54
         Top             =   1755
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医生嘱托(&T)"
         Height          =   180
         Index           =   21
         Left            =   5385
         TabIndex        =   53
         Top             =   285
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行时间(&V)"
         Height          =   180
         Index           =   20
         Left            =   5385
         TabIndex        =   52
         Top             =   990
         Width           =   990
      End
   End
   Begin VB.Frame fra 
      Caption         =   "其他信息"
      Height          =   1815
      Index           =   1
      Left            =   15
      TabIndex        =   17
      Top             =   1230
      Visible         =   0   'False
      Width           =   8700
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   11
         Left            =   7320
         MaxLength       =   6
         TabIndex        =   30
         Top             =   1365
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   10
         Left            =   4695
         MaxLength       =   20
         TabIndex        =   29
         Top             =   1365
         Width           =   1545
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   9
         Left            =   1110
         MaxLength       =   50
         TabIndex        =   28
         Top             =   1365
         Width           =   2475
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   7320
         MaxLength       =   6
         TabIndex        =   27
         Top             =   975
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   4695
         MaxLength       =   20
         TabIndex        =   26
         Top             =   975
         Width           =   1545
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   6
         Left            =   1110
         MaxLength       =   100
         TabIndex        =   25
         Top             =   975
         Width           =   2190
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   6
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   585
         Width           =   2475
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   5
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   195
         Width           =   1275
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   4
         Left            =   4695
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   210
         Width           =   1545
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   3
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   210
         Width           =   2475
      End
      Begin VB.CommandButton cmd 
         Caption         =   "…"
         Height          =   270
         Index           =   6
         Left            =   3300
         TabIndex        =   20
         ToolTipText     =   "热键:F3"
         Top             =   990
         Width           =   270
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   7320
         MaxLength       =   6
         TabIndex        =   19
         Top             =   585
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   4695
         MaxLength       =   20
         TabIndex        =   18
         Top             =   585
         Width           =   1545
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住址邮编(L)"
         Height          =   180
         Index           =   18
         Left            =   6315
         TabIndex        =   42
         Top             =   1425
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭电话(K)"
         Height          =   180
         Index           =   17
         Left            =   3645
         TabIndex        =   41
         Top             =   1425
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭地址(E)"
         Height          =   180
         Index           =   16
         Left            =   105
         TabIndex        =   40
         Top             =   1425
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位邮编(&B)"
         Height          =   180
         Index           =   15
         Left            =   6315
         TabIndex        =   39
         Top             =   1035
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位电话(&T)"
         Height          =   180
         Index           =   14
         Left            =   3645
         TabIndex        =   38
         Top             =   1035
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称(&U)"
         Height          =   180
         Index           =   13
         Left            =   105
         TabIndex        =   37
         Top             =   1035
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国    籍(G)"
         Height          =   180
         Index           =   7
         Left            =   105
         TabIndex        =   36
         Top             =   270
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民    族(&P)"
         Height          =   180
         Index           =   8
         Left            =   3645
         TabIndex        =   35
         Top             =   270
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职    业(&J)"
         Height          =   180
         Index           =   10
         Left            =   105
         TabIndex        =   34
         Top             =   645
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况(&M)"
         Height          =   180
         Index           =   9
         Left            =   6315
         TabIndex        =   33
         Top             =   270
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联 系 人(&Z)"
         Height          =   180
         Index           =   12
         Left            =   6315
         TabIndex        =   32
         Top             =   645
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系电话(&W)"
         Height          =   180
         Index           =   11
         Left            =   3645
         TabIndex        =   31
         Top             =   645
         Width           =   990
      End
   End
   Begin VB.Frame fra 
      Caption         =   "基本信息"
      Height          =   1125
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   8700
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1125
         MaxLength       =   20
         TabIndex        =   9
         ToolTipText     =   "数字为就诊卡号、“－”打头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“#”收费单据号"
         Top             =   210
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   7980
         MaxLength       =   10
         TabIndex        =   8
         Top             =   210
         Width           =   585
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   6315
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   210
         Width           =   945
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   3825
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   1710
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   2
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   1905
      End
      Begin VB.CommandButton cmdMore 
         Caption         =   ">>"
         Height          =   300
         Left            =   8280
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "更多病人信息"
         Top             =   570
         Width           =   315
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   3825
         MaxLength       =   18
         TabIndex        =   3
         Top             =   210
         Width           =   1710
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   1125
         MaxLength       =   20
         TabIndex        =   2
         ToolTipText     =   "数字为就诊卡号、“－”打头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“#”收费单据号"
         Top             =   600
         Width           =   1590
      End
      Begin VB.CommandButton cmd 
         Caption         =   "…"
         Height          =   285
         Index           =   0
         Left            =   2430
         TabIndex        =   1
         ToolTipText     =   "热键:F3"
         Top             =   225
         Width           =   285
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费    别(&F)"
         Height          =   180
         Index           =   3
         Left            =   2820
         TabIndex        =   16
         Top             =   675
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓    名(&1)"
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   15
         Top             =   270
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄(&Y)"
         Height          =   180
         Index           =   2
         Left            =   7320
         TabIndex        =   14
         Top             =   270
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "付款(&A)"
         Height          =   180
         Index           =   4
         Left            =   5670
         TabIndex        =   13
         Top             =   675
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别(&S)"
         Height          =   180
         Index           =   1
         Left            =   5670
         TabIndex        =   12
         Top             =   270
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号(&I)"
         Height          =   180
         Index           =   6
         Left            =   2820
         TabIndex        =   11
         Top             =   270
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门 诊 号(&N)"
         Height          =   180
         Index           =   5
         Left            =   105
         TabIndex        =   10
         Top             =   675
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmOpsStationRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'（１）窗体级变量定义
'**********************************************************************************************************************

Private Type Items
    项目名称 As String
    麻醉方式 As String
End Type

Private usrSaveItem As Items
Private mstr收费单据号 As String
Private mblnStartUp As Boolean                          '窗体启动标志
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mint病人来源 As Integer
Private mblnDataChanged As Boolean
Private mlngDept As Long
Private mstrPrivs As String
Private WithEvents mclsVsfAdvice As clsVsf
Attribute mclsVsfAdvice.VB_VarHelpID = -1

'（２）自定义过程或函数
'**********************************************************************************************************************

Public Function ShowEdit(ByVal frmMain As Object, ByVal lngDept As Long) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '--------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
    Set mfrmMain = frmMain
    mlngDept = lngDept
    
    Call ExecuteCommand("初始控件")
    
    If ExecuteCommand("初始数据") = False Then Exit Function

    fra(1).Visible = False
    fra(2).Top = fra(2).Top - fra(1).Height
    picButton.Top = picButton.Top - fra(1).Height
    Me.Height = Me.Height - fra(1).Height
    
    mblnStartUp = False
    
    Call cbo_Click(8)
    
    cmdOK.Tag = ""
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '--------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim lng病人id As Long
    Dim str手术项目 As String
    Dim lng主页id As Long
            
    On Error GoTo errHand
    
    Call SQLRecord(rsSQL)
    
    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
        Case "初始控件"
            
            Set mclsVsfAdvice = New clsVsf
            With mclsVsfAdvice
                Call .Initialize(Me.Controls, vsfAdvice, True, True, frmPubResource.GetImageList(16))
                Call .ClearColumn
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[指示器]", False)

                Call .AppendColumn("手术名称", 3000, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("缺省", 450, flexAlignCenterCenter, flexDTBoolean, "", , True)
    
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("手术名称"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("缺省"), True, vbVsfEditCheck)
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("当前").Picture
                
                .AppendRows = True
            End With
        
        '--------------------------------------------------------------------------------------------------------------
        Case "初始数据"
            
            dtp(0).Value = Format(zlDatabase.Currentdate, dtp(0).CustomFormat)
            '性别
            gstrSQL = "Select 编码||'-'||名称  As 名称,0,缺省标志 From 性别"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            Call AddComboData(cbo(0), rs)
            If cbo(0).ListCount > 0 Then cbo(0).ListIndex = 0

            '费别
            cbo(1).Clear
            cbo(1).AddItem ""
            gstrSQL = "Select 编码||'-'||名称  As 名称,0,缺省标志 From 费别"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            Call AddComboData(cbo(1), rs, False)
            If cbo(1).ListCount > 0 Then cbo(1).ListIndex = 0

            '付款方式
            cbo(2).Clear
            cbo(2).AddItem ""
            gstrSQL = "Select 编码||'-'||名称  As 名称,0,缺省标志 From 医疗付款方式"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            Call AddComboData(cbo(2), rs, False)
            If cbo(2).ListCount > 0 Then cbo(2).ListIndex = 0
            
            '国籍
            cbo(3).Clear
            cbo(3).AddItem ""
            gstrSQL = "Select 编码||'-'||名称  As 名称,0,缺省标志 From 国籍 Order By 编码"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            Call AddComboData(cbo(3), rs, False)
            If cbo(3).ListCount > 0 Then cbo(3).ListIndex = 0
            
            '民族
            cbo(4).Clear
            cbo(4).AddItem ""
            gstrSQL = "Select 编码||'-'||名称  As 名称,0,缺省标志 From 民族"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            Call AddComboData(cbo(4), rs, False)
            If cbo(4).ListCount > 0 Then cbo(4).ListIndex = 0

            '婚姻状况
            cbo(5).Clear
            cbo(5).AddItem ""
            gstrSQL = "Select 编码||'-'||名称  As 名称,0,缺省标志 From 婚姻状况"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            Call AddComboData(cbo(5), rs, False)
            If cbo(5).ListCount > 0 Then cbo(5).ListIndex = 0

            '职业
            cbo(6).Clear
            cbo(6).AddItem ""
            gstrSQL = "Select 编码||'-'||名称  As 名称,0,缺省标志 From 职业"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            Call AddComboData(cbo(6), rs, False)
            If cbo(6).ListCount > 0 Then cbo(6).ListIndex = 0
            
            '执行科室
            gstrSQL = "Select Distinct b.编码||'-'||b.名称 As 名称,b.ID From 部门表 b Where b.ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngDept)
            Call AddComboData(cbo(7), rs)
            If cbo(7).ListCount > 0 Then cbo(7).ListIndex = 0
            
            '申请科室
            gstrSQL = GetPublicSQL(SQL.临床部门记录, "所有")
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            Call AddComboData(cbo(8), rs)
            If cbo(8).ListCount > 0 Then cbo(8).ListIndex = 0
            
            txt(1).MaxLength = 18
            txt(2).MaxLength = GetMaxLength("病人信息", "年龄")
            txt(3).MaxLength = GetMaxLength("病人信息", "门诊号")
            txt(4).MaxLength = GetMaxLength("病人信息", "联系人电话")
            txt(5).MaxLength = GetMaxLength("病人信息", "联系人姓名")
            txt(6).MaxLength = GetMaxLength("病人信息", "工作单位")
            txt(7).MaxLength = GetMaxLength("病人信息", "单位电话")
            txt(8).MaxLength = GetMaxLength("病人信息", "单位邮编")
            txt(9).MaxLength = GetMaxLength("病人信息", "家庭地址")
            txt(10).MaxLength = GetMaxLength("病人信息", "家庭电话")
            txt(11).MaxLength = GetMaxLength("病人信息", "户口邮编")
            
            txt(12).MaxLength = GetMaxLength("病人医嘱记录", "医嘱内容")
            txt(13).MaxLength = GetMaxLength("病人医嘱记录", "医生嘱托")

        '--------------------------------------------------------------------------------------------------------------
        Case "校验数据"         '检验输入的数据有效性
        
            If txt(0).Text = "" Then
                ShowSimpleMsg "手术申请必须指定做手术的病人！"
                LocationObj txt(0)
                Exit Function
            End If
            
            With vsfAdvice
                For lngLoop = 1 To .Rows - 1
                    If Val(.RowData(lngLoop)) > 0 Then
                        If Abs(Val(.TextMatrix(lngLoop, .ColIndex("缺省")))) = 1 Then
                            Exit For
                        End If
                    End If
                Next
                
                If lngLoop = .Rows Then
                    ShowSimpleMsg "必须指定一个缺省的手术！"
                    LocationGrid vsfAdvice
                    Exit Function
                End If
            End With
            
        '--------------------------------------------------------------------------------------------------------------
        Case "保存数据"         '保存更改后的数据
            
            ExecuteCommand = SaveData
            
            Exit Function
        End Select
    Next
    
    ExecuteCommand = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Function SaveData() As Boolean
    Dim lngKey As Long
    Dim intLoop As Integer
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim lng病人id As Long
    Dim str手术项目 As String
    Dim lng主页id As Long
    Dim blnTrans As Boolean
    Dim str标识号 As String
    
    On Error GoTo errHand
    
    Call SQLRecord(rsSQL)
    
    lng病人id = Val(cmd(0).Tag)
    lng主页id = IIf(mint病人来源 = 2, Val(lbl(5).Tag), 0)
    
    '------------------------------------------------------------------------------------------------------------------
    
    With vsfAdvice
        For lngLoop = 1 To .Rows - 1
            If Val(.RowData(lngLoop)) > 0 Then
                If Abs(Val(.TextMatrix(lngLoop, .ColIndex("缺省")))) = 1 Then
                    str手术项目 = Val(.RowData(lngLoop)) & ",F," & .TextMatrix(lngLoop, .ColIndex("手术名称")) & IIf(str手术项目 = "", "", ";" & str手术项目)
                Else
                    str手术项目 = IIf(str手术项目 = "", "", str手术项目 & ";") & Val(.RowData(lngLoop)) & ",F," & .TextMatrix(lngLoop, .ColIndex("手术名称"))
                End If
            End If
        Next
    End With
        
    If Val(cmd(12).Tag) > 0 Then
        str手术项目 = IIf(str手术项目 = "", "", str手术项目 & ";") & Val(cmd(12).Tag) & ",G," & txt(12).Text
    End If
    
    lngKey = zlDatabase.GetNextId("病人医嘱记录")
    
    str标识号 = "Null"
    If IsNumeric(txt(3).Text) Then str标识号 = txt(3).Text
    
    If lng病人id = 0 Then lng病人id = zlDatabase.GetNextNo(1)
    
    gstrSQL = "Zl_病人手术记录_Request("
    gstrSQL = gstrSQL & lngKey & "," & IIf(mint病人来源 = 0, 1, mint病人来源) & "," & _
                        lng病人id & "," & _
                        ZVal(lng主页id) & "," & _
                        str标识号 & ",'" & _
                        txt(0).Text & "','" & _
                        zlCommFun.GetNeedName(cbo(0).Text) & "','" & _
                        txt(2).Text & "','" & _
                        zlCommFun.GetNeedName(cbo(1).Text) & "','" & _
                        zlCommFun.GetNeedName(cbo(2).Text) & "','" & _
                        zlCommFun.GetNeedName(cbo(3).Text) & "','" & _
                        zlCommFun.GetNeedName(cbo(4).Text) & "','" & _
                        zlCommFun.GetNeedName(cbo(5).Text) & "','" & _
                        zlCommFun.GetNeedName(cbo(6).Text) & "','" & _
                        txt(1).Text & "','" & _
                        txt(6).Text & "'," & _
                        ZVal(cmd(6).Tag) & ",'" & _
                        txt(7).Text & "','" & _
                        txt(8).Text & "','" & _
                        txt(9).Text & "','" & _
                        txt(10).Text & "','" & _
                        txt(11).Text & "','" & _
                        str手术项目 & "','" & _
                        txt(13).Text & "'," & _
                        cbo(7).ItemData(cbo(7).ListIndex) & "," & _
                        chk(0).Value & ","
    gstrSQL = gstrSQL & "To_Date('" & Format(dtp(0).Value, "yyyy-MM-dd HH:mm") & ":00','yyyy-mm-dd hh24:mi:ss')," & _
                        cbo(8).ItemData(cbo(8).ListIndex) & ",'" & _
                        zlCommFun.GetNeedName(cbo(9).Text) & "'," & _
                        "Sysdate)"
                            
    Call SQLRecordAdd(rsSQL, gstrSQL, 1)
                
    
    '开始执行SQL,即提交到数据库中
    '------------------------------------------------------------------------------------------------------------------
    SaveData = SQLRecordExecute(rsSQL, Me.Caption)
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CreateOrderCharge(ByVal lngKey As Long, Optional ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '功能：从用药和材料中生成附加费用
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsPati As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    Dim rsCharge As New ADODB.Recordset
    Dim rsNo As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strNO As String
    Dim int来源 As Integer
            
    Dim lng医嘱id As Long
    Dim int父号 As Integer
    Dim lng项目ID As Long
    Dim lng执行部门ID As Long
    Dim lng病人病区ID As Long
    Dim lng病人科室ID As Long
    Dim lng类别ID As Long
    Dim strDate As String
    Dim lngLoop As Long
    Dim int保险项目否 As Integer
    Dim lng保险大类ID As Long
    Dim str保险编码 As String
    Dim cur统筹金额 As Currency
    Dim cur应收 As Currency
    Dim cur实收 As Currency
    Dim strMsg As String
    Dim dbl数量 As Double
    Dim blnTran As Boolean
    Dim cur单价 As Currency
    Dim lng报警级别 As Long
    Dim lng已报警级别 As Long
    Dim str报警方案 As String
    Dim lng级别 As Long
    Dim str已强制报警姓名 As String
    Dim bln医保 As Boolean
    Dim curMoneyTotal As Currency
    Dim str费用小数位 As String
    Dim strSQL As String
    Dim rsSQL As ADODB.Recordset
    Dim bln强制记帐 As Boolean
    Dim lng病人id As Long
    Dim lng主页id As Long
    Dim lng发送号 As Long
    Dim int记录性质 As Integer
    Dim strTmp As String
    
    On Error GoTo errHand
    
    Screen.MousePointer = 11
    
    Call SQLRecord(rsSQL)
    
    '初始设置
    '------------------------------------------------------------------------------------------------------------------
    Set rsNo = New ADODB.Recordset
    With rsNo
        .Fields.Append "No", adVarChar, 30
        .Open
    End With
    
    gstrSQL = "Select a.病人id,a.主页id,a.病人来源,b.发送号 From 病人医嘱记录 a,病人医嘱发送 b Where a.ID=[1] And a.ID=b.医嘱id"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngKey)
    If rs.BOF Then
        Screen.MousePointer = 0
        Exit Function
    End If
    
    lng病人id = rs("病人id").Value
    lng主页id = zlCommFun.NVL(rs("主页id").Value, 0)
    int来源 = rs("病人来源").Value
    int记录性质 = IIf(int来源 = 1, 1, 2)
    lng发送号 = zlCommFun.NVL(rs("发送号").Value, 0)
    
    '取费用金额保存小数
    '------------------------------------------------------------------------------------------------------------------
    str费用小数位 = ParamInfo.费用金额小数位数
    bln强制记帐 = (InStr(strPrivs, "欠费强制记帐") > 0)
    
    '获取病人的信息
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select A.姓名,A.性别,A.年龄,Nvl(B.费别,A.费别) as 费别," & _
        " A.门诊号,A.住院号,Nvl(A.当前床号,B.出院病床) as 床号," & _
        " Nvl(A.当前病区ID,B.当前病区ID) as 病人病区ID," & _
        " Nvl(A.当前科室ID,B.出院科室ID) as 病人科室ID," & _
        " Nvl(B.险类,A.险类) as 险类,C.编码 as 付款码" & _
        " From 病人信息 A,病案主页 B,医疗付款方式 C" & _
        " Where A.病人ID=[1] And A.病人ID=B.病人ID(+)" & _
        " And B.主页ID(+)=[2] And A.医疗付款方式=C.名称(+)"
    
    Set rsPati = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng病人id, lng主页id)

    If rsPati.BOF Then
        Screen.MousePointer = 0
        Exit Function
    End If
    
    bln医保 = (Val(zlCommFun.NVL(rsPati!付款码, "0")) = 1)
    
    '可能对照费用为药品费用
    '------------------------------------------------------------------------------------------------------------------
    lng类别ID = ExistIOClass(IIf(int记录性质 = 1, 8, 9)) '8:门诊划价单;9:门诊/住院记帐单
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    gstrSQL = "SELECT A.医嘱id,B.ID AS 收费细目ID," & _
                  "A.数量,A.可否分零,A.剂量系数,A.包装," & _
                  "B.计算单位," & _
                  "B.类别," & _
                  "C.现价 AS 单价," & _
                  "D.收据费目," & _
                  "C.收入项目ID," & _
                  "A.执行科室id," & _
                  "DECODE(A.主页id,NULL,F.门诊号,0,F.门诊号,F.住院号) AS 标识号," & _
                  "F.费别," & _
                  "A.病人科室id AS 当前科室ID," & _
                  "DECODE(F.当前病区ID,NULL,A.病人科室id,F.当前病区ID) AS 当前病区ID," & _
                  "F.当前床号," & _
                  "A.病人ID," & _
                  "A.主页id," & _
                  "F.姓名," & _
                  "F.性别," & _
                  "F.年龄," & _
                  "B.名称,a.No,a.记录性质 " & _
            "FROM   收费项目目录 B," & _
               "收费价目 C," & _
               "收入项目 D," & _
               "病人信息 F," & _
               "("
               
    gstrSQL = gstrSQL & _
        "SELECT AA.医嘱id,bb.No,bb.记录性质,HH.可否分零,Decode(HH.剂量系数,0,1,Null,1,HH.剂量系数) As 剂量系数,Decode(GG.病人来源,2,HH.住院包装,HH.门诊包装) As 包装,GG.病人科室id,3 AS 序号,AA.收费细目id,AA.数量,AA.执行科室id,GG.病人id,GG.主页id ,0 AS 单价 " & _
        "FROM 病人医嘱计价 AA,药品规格 HH,病人医嘱记录 GG,病人医嘱发送 BB " & _
        "Where AA.收费细目ID = HH.药品id(+) And AA.医嘱id = GG.ID And [1] In (GG.ID,GG.相关id) And BB.医嘱id=AA.医嘱id "
    
    gstrSQL = gstrSQL & _
               ") A " & _
            "Where C.收费细目id = B.ID " & _
               "AND C.收入项目ID = D.ID " & _
               "AND C.执行日期 <= SYSDATE " & _
               "AND A.数量 > 0 " & _
               "AND (C.终止日期 >= SYSDATE OR C.终止日期 IS NULL) " & _
               "AND A.收费细目id = B.ID " & _
               "AND F.病人id=A.病人id " & _
            "ORDER BY B.ID"
    
    Set rsCharge = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lngKey)
    If rsCharge.BOF Then
        Screen.MousePointer = 0
        Exit Function
    End If
    
    '
    '------------------------------------------------------------------------------------------------------------------
    With rsCharge
        
        '获取对应的医嘱信息
        gstrSQL = "Select 医嘱期效,病人科室ID,婴儿,执行频次,计价特性,诊疗项目id From 病人医嘱记录 Where ID=[1]"
        Set rsAdvice = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", Val(rsCharge("医嘱id").Value))
        If rsAdvice.BOF Then
            Screen.MousePointer = 0
            Exit Function
        End If
        
        int记录性质 = rsCharge("记录性质").Value
        strNO = rsCharge("No").Value
        rsNo.Filter = ""
        rsNo.Filter = "No='" & strNO & "'"
        If rsNo.RecordCount = 0 Then
            rsNo.AddNew
            rsNo("No").Value = strNO
        End If
        
        For lngLoop = 1 To .RecordCount
            
            dbl数量 = zlCommFun.NVL(rsCharge("数量").Value, 0)
            
            
            '病人病区科室、执行科室
            '----------------------------------------------------------------------------------------------------------
            lng病人病区ID = zlCommFun.NVL(rsPati!病人病区ID, 0)
            lng病人科室ID = zlCommFun.NVL(rsPati!病人科室ID, 0)
            If lng病人科室ID = 0 Then
                lng病人病区ID = zlCommFun.NVL(rsAdvice!病人科室ID, 0)
                lng病人科室ID = zlCommFun.NVL(rsAdvice!病人科室ID, 0)
            End If
            If lng病人科室ID = 0 Then
                lng病人病区ID = UserInfo.部门ID
                lng病人科室ID = UserInfo.部门ID
            End If
            
            lng执行部门ID = !执行科室id
            
            Select Case rsCharge("类别").Value
            Case "5", "6", "7"
                lng执行部门ID = GetDefaultDept(rsCharge("类别").Value, mint病人来源)
                
                gstrSQL = GetPublicSQL(SQL.收费执行科室, rsCharge("类别").Value)
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", lng执行部门ID, Val(rsCharge("诊疗项目id").Value), lng病人科室ID, UserInfo.部门ID)
                If rs.BOF = False Then
                    rs.Filter = ""
                    rs.Filter = "ID=" & lng执行部门ID
                    If rs.RecordCount = 0 Then
                        rs.Filter = ""
                        lng执行部门ID = rs("ID").Value
                    End If
                Else
                    lng执行部门ID = 0
                End If
            End Select
            
            If lng执行部门ID = 0 Then
                ShowSimpleMsg !名称 & "未指定执行科室，不能继续！"
                Screen.MousePointer = 0
                Exit Function
            End If
            
            cur单价 = rsCharge("单价").Value
            
            '检查普通收费项目的库存，计算实价药品/材料的单价
            '----------------------------------------------------------------------------------------------------------
            Select Case rsCharge("类别").Value
            Case "4", "5", "6", "7"
                Select Case rsCharge("类别").Value
                Case "4"
                    gstrSQL = "SELECT NVL(B.是否变价,0) AS 实价,NVL(在用分批,0) AS 分批 FROM 材料特性 A,收费项目目录 B WHERE A.材料id=B.ID AND A.材料id=[1] "
                Case "5", "6", "7"
                    '进行分零计算
                    dbl数量 = dbl数量
                    
                    If zlCommFun.NVL(rsCharge("可否分零").Value, 0) = 0 Then
                        dbl数量 = dbl数量 / zlCommFun.NVL(rsCharge("剂量系数").Value, 1)
                    Else
                        dbl数量 = IntEx(dbl数量 / zlCommFun.NVL(rsCharge("剂量系数").Value, 1) / zlCommFun.NVL(rsCharge("包装").Value, 1)) * zlCommFun.NVL(rsCharge("包装").Value, 1)
                    End If
                                            
                    gstrSQL = "SELECT NVL(I.是否变价,0) AS 实价,NVL(S.药房分批,0) AS 分批 FROM 收费项目目录 I,药品规格 S WHERE I.ID=S.药品id AND S.药品id=[1]"
                End Select
                
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlOps", Val(!收费细目id))
                If rs.BOF = False Then
                    If rs("分批").Value <> 1 And rs("实价").Value <> 1 Then
                        '是普通项目,要检查库存
                        If dbl数量 > CalcStorage(!收费细目id, lng执行部门ID, False, False) Then
                            '超过库存数量
                            Select Case GetDrugWarnOption(lng执行部门ID, rsCharge("类别").Value)
                            Case 1          '库存不足提醒
                                If MsgBox(!名称 & "库存不足，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    Screen.MousePointer = 0
                                    Exit Function
                                End If
                            Case 2          '库存不足禁止
                                MsgBox !名称 & "库存不足！", vbInformation, gstrSysName
                                Screen.MousePointer = 0
                                Exit Function
                            End Select
                        End If
                    ElseIf rs("实价") = 1 Then
                        cur单价 = CalcTimePrice(!收费细目id, lng执行部门ID, dbl数量)
                    End If
                End If
            End Select
                           
            '计算应收和实收金额
            '----------------------------------------------------------------------------------------------------------
            cur应收 = Format(dbl数量 * cur单价, str费用小数位)
            cur实收 = cur应收
            If rsPati("费别").Value <> "" Then cur实收 = Format(ActualMoney(rsPati("费别").Value, !收入项目ID, cur应收), str费用小数位)
            
            '每个收费项目的处理
            '----------------------------------------------------------------------------------------------------------
            If lng项目ID <> !收费细目id Then
            
                int父号 = lngLoop '获取价格父号
                
                '获取保险项目信息
                '------------------------------------------------------------------------------------------------------
                If int来源 = 2 And Not IsNull(rsPati!险类) And gblnInsure Then
                    strMsg = gclsInsure.GetItemInsure(lng病人id, !收费细目id, cur实收, False, rsPati!险类)
                    If strMsg <> "" Then
                        int保险项目否 = Val(Split(strMsg, ";")(0))
                        lng保险大类ID = Val(Split(strMsg, ";")(1))
                        cur统筹金额 = Format(Val(Split(strMsg, ";")(2)), "0.00")
                        str保险编码 = CStr(Split(strMsg, ";")(3))
                    End If
                End If
            End If
            lng项目ID = !收费细目id
            
            
            '如果是记帐单据，进行费用警告
            '----------------------------------------------------------------------------------------------------------
            
            If int记录性质 = 2 Then
                
                '搜索当前医嘱的最高报警级别,并与已报警级别比较
                
'                lng级别 = GetWarnGrade(lng已报警级别, !类别, bln医保, lng病人病区ID)
                
                str报警方案 = ""
                strSQL = "Select zl_PatiWarnScheme([1],[2]) As 报警方案 From Dual"
                Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlOps", lng病人id, lng主页id)
                If rs.BOF = False Then
                    str报警方案 = zlCommFun.NVL(rs("报警方案").Value)
                End If
                lng级别 = GetWarnGrade(lng已报警级别, !类别, str报警方案, lng病人病区ID)
                
                lng报警级别 = IIf(lng报警级别 > lng级别, lng报警级别, lng级别)
                lng报警级别 = IIf(lng报警级别 > lng已报警级别, lng报警级别, lng已报警级别)
                            
                '判断是否费用是否够用
                curMoneyTotal = curMoneyTotal + cur实收
                
                If lng报警级别 > lng已报警级别 Then
                    If curMoneyTotal <> 0 Then
'                        If 欠费情况(zlCommFun.NVL(rsPati!姓名), lng病人id, lng主页id, curMoneyTotal, bln医保, lng报警级别, bln强制记帐, str已强制报警姓名) = "是" Then
                        If 欠费情况(zlCommFun.NVL(rsPati!姓名), lng病人id, lng主页id, curMoneyTotal, str报警方案, lng报警级别, bln强制记帐, str已强制报警姓名) = "是" Then
                            Screen.MousePointer = 0
                            Exit Function
                        End If
                    End If
                End If
            End If
            
            '填写记录
            '----------------------------------------------------------------------------------------------------------
            If int来源 = 1 Then
                If int记录性质 = 1 Then
                    '生成门诊划价单据
                    '--------------------------------------------------------------------------------------------------
                    strSQL = _
                        "zl_门诊划价记录_Insert('" & strNO & "'," & lngLoop & "," & lng病人id & ",NULL," & _
                        ZVal(zlCommFun.NVL(rsPati!门诊号, 0)) & ",'" & zlCommFun.NVL(rsPati!付款码) & "','" & zlCommFun.NVL(rsPati!姓名) & "'," & _
                        "'" & zlCommFun.NVL(rsPati!性别) & "','" & zlCommFun.NVL(rsPati!年龄) & "','" & zlCommFun.NVL(rsPati!费别) & "',NULL," & _
                        lng病人病区ID & "," & lng病人科室ID & "," & UserInfo.部门ID & ",'" & UserInfo.姓名 & "'," & _
                        "NULL," & lng项目ID & ",'" & !类别 & "','" & !计算单位 & "',NULL,1," & dbl数量 & "," & _
                        "0," & ZVal(lng执行部门ID) & "," & IIf(int父号 = lngLoop, "NULL", int父号) & "," & _
                        !收入项目ID & ",'" & zlCommFun.NVL(!收据费目) & "'," & cur单价 & "," & cur应收 & "," & cur实收 & "," & _
                        strDate & "," & strDate & ",NULL,'" & UserInfo.姓名 & "'," & ZVal(lng类别ID) & ",NULL," & _
                        lngKey & ",'" & zlCommFun.NVL(rsAdvice!执行频次) & "',NULL,NULL," & zlCommFun.NVL(rsAdvice!医嘱期效, 0) & "," & _
                        zlCommFun.NVL(rsAdvice!计价特性, 0) & ",1)"
                    Call SQLRecordAdd(rsSQL, strSQL)
                Else
                    '生成门诊记帐单据
                    '--------------------------------------------------------------------------------------------------
                    strSQL = _
                        "zl_门诊记帐记录_Insert('" & strNO & "'," & lngLoop & "," & lng病人id & "," & _
                        ZVal(zlCommFun.NVL(rsPati!门诊号, 0)) & ",'" & zlCommFun.NVL(rsPati!姓名) & "','" & zlCommFun.NVL(rsPati!性别) & "'," & _
                        "'" & zlCommFun.NVL(rsPati!年龄) & "','" & zlCommFun.NVL(rsPati!费别) & "',NULL," & ZVal(rsAdvice!婴儿) & "," & _
                        lng病人病区ID & "," & lng病人科室ID & "," & UserInfo.部门ID & "," & _
                        "'" & UserInfo.姓名 & "',NULL," & lng项目ID & ",'" & !类别 & "'," & _
                        "'" & !计算单位 & "',1," & dbl数量 & ",0," & ZVal(lng执行部门ID) & "," & _
                        IIf(int父号 = lngLoop, "NULL", int父号) & "," & !收入项目ID & ",'" & zlCommFun.NVL(!收据费目) & "'," & cur单价 & "," & _
                        cur应收 & "," & cur实收 & "," & strDate & "," & strDate & ",NULL,NULL,'" & UserInfo.编号 & "'," & _
                        "'" & UserInfo.姓名 & "'," & ZVal(lng类别ID) & ",NULL,NULL," & lngKey & "," & _
                        "'" & zlCommFun.NVL(rsAdvice!执行频次) & "',NULL,NULL," & zlCommFun.NVL(rsAdvice!医嘱期效, 0) & "," & _
                        zlCommFun.NVL(rsAdvice!计价特性, 0) & ")"
                    Call SQLRecordAdd(rsSQL, strSQL)
                End If
            Else
                '生成住院记帐单据
                '------------------------------------------------------------------------------------------------------
                strSQL = _
                    "zl_住院记帐记录_Insert('" & strNO & "'," & lngLoop & "," & lng病人id & "," & ZVal(lng主页id) & "," & _
                    ZVal(zlCommFun.NVL(rsPati!住院号, 0)) & ",'" & zlCommFun.NVL(rsPati!姓名) & "','" & zlCommFun.NVL(rsPati!性别) & "'," & _
                    "'" & zlCommFun.NVL(rsPati!年龄) & "','" & Trim(zlCommFun.NVL(rsPati!床号)) & "','" & zlCommFun.NVL(rsPati!费别) & "'," & _
                    lng病人病区ID & "," & lng病人科室ID & ",NULL," & ZVal(rsAdvice!婴儿) & "," & _
                    UserInfo.部门ID & ",'" & UserInfo.姓名 & "',NULL," & lng项目ID & ",'" & !类别 & "'," & _
                    "'" & !计算单位 & "'," & int保险项目否 & "," & ZVal(lng保险大类ID) & ",'" & str保险编码 & "'," & _
                    "1," & dbl数量 & ",0," & ZVal(lng执行部门ID) & "," & _
                    IIf(int父号 = lngLoop, "NULL", int父号) & "," & !收入项目ID & ",'" & zlCommFun.NVL(!收据费目) & "'," & cur单价 & "," & _
                    cur应收 & "," & cur实收 & "," & cur统筹金额 & "," & strDate & "," & strDate & ",NULL,NULL," & _
                    "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "',NULL," & ZVal(lng类别ID) & ",NULL,NULL,NULL," & _
                    lngKey & ",'" & zlCommFun.NVL(rsAdvice!执行频次) & "',NULL,NULL," & zlCommFun.NVL(rsAdvice!医嘱期效, 0) & "," & _
                    zlCommFun.NVL(rsAdvice!计价特性, 0) & ",NULL)"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
            
            .MoveNext
            
        Next

    End With
    
    '
    '------------------------------------------------------------------------------------------------------------------
        
'    blnTran = True
'    gcnOracle.BeginTrans
    
    If SQLRecordExecute(rsSQL, "mdlOps", False) = False Then GoTo errHand
        
    '在提交前进行医保传输
    '------------------------------------------------------------------------------------------------------------------
    If int来源 = 2 And Not IsNull(rsPati!险类) And gblnInsure Then
        If gclsInsure.GetCapability(support记帐上传, lng病人id, rsPati!险类) And Not gclsInsure.GetCapability(support记帐完成后上传, lng病人id, rsPati!险类) Then
            If rsNo.RecordCount > 0 Then
                rsNo.MoveFirst
                Do While Not rsNo.EOF
                    strMsg = ""
                    If Not gclsInsure.TranChargeDetail(2, rsNo("No").Value, 2, 1, strMsg, rsPati!险类) Then
                        gcnOracle.RollbackTrans
                        If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                        Screen.MousePointer = 0: Exit Function
                    End If
                    rsNo.MoveNext
                Loop
            End If
        End If
    End If
    
    gcnOracle.CommitTrans
    blnTran = False
    CreateOrderCharge = True
    
    '在提交后进行医保传输
    '------------------------------------------------------------------------------------------------------------------
    If int来源 = 2 And Not IsNull(rsPati!险类) And gblnInsure Then
        If gclsInsure.GetCapability(support记帐上传, lng病人id, rsPati!险类) And gclsInsure.GetCapability(support记帐完成后上传, lng病人id, rsPati!险类) Then
            If rsNo.RecordCount > 0 Then
                rsNo.MoveFirst
                Do While Not rsNo.EOF
                    strMsg = ""
                    If Not gclsInsure.TranChargeDetail(2, rsNo("No").Value, 2, 1, strMsg, rsPati!险类) Then
                        If strMsg <> "" Then
                            MsgBox strMsg, vbInformation, gstrSysName
                        Else
                            MsgBox "单据""" & rsNo("No").Value & """的数据向医保传送失败,该单据已保存！", vbInformation, gstrSysName
                        End If
                    End If
                    rsNo.MoveNext
                Loop
            End If
        End If
    End If
        
    Screen.MousePointer = 0

    Exit Function
    
    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If blnTran Then gcnOracle.RollbackTrans
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Function

Private Sub chk_Click(Index As Integer)
    cmdOK.Tag = "Changed"
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsResult As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim lngKey As Long
    
    Select Case Index
    '------------------------------------------------------------------------------------------------------------------
    Case 0
    
        If frmPatientFind.ShowFind(Me, lngKey) Then
            If lngKey > 0 Then
                
                gstrSQL = "SELECT a.*,b.主页id FROM 病人信息 a,病案主页 b WHERE a.病人id=[1] and a.病人id=b.病人id(+) And b.出院日期 Is Null "
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
                If rs.BOF = False Then
                    cmd(Index).Tag = zlCommFun.NVL(rs("病人id").Value)
                    
                    txt(Index).Text = zlCommFun.NVL(rs("姓名").Value)
                    txt(1).Text = zlCommFun.NVL(rs("身份证号").Value)
                    txt(2).Text = zlCommFun.NVL(rs("年龄").Value)

                    If Val(zlCommFun.NVL(rs("主页id"))) > 0 Then
                        mint病人来源 = 2
                        lbl(5).Tag = Val(zlCommFun.NVL(rs("主页id")))
                        lbl(5).Caption = "住院号(&N)"
                        txt(3).Text = zlCommFun.NVL(rs("住院号"))
                    Else
                        mint病人来源 = 1
                        lbl(5).Tag = 0
                        lbl(5).Caption = "门诊号(&N)"
                        txt(3).Text = zlCommFun.NVL(rs("门诊号"))
                    End If
                    
                    txt(4).Text = zlCommFun.NVL(rs("联系人电话").Value)
                    txt(5).Text = zlCommFun.NVL(rs("联系人姓名").Value)
                    txt(6).Text = zlCommFun.NVL(rs("工作单位").Value)
                    cmd(6).Tag = zlCommFun.NVL(rs("合同单位ID").Value)
                    txt(7).Text = zlCommFun.NVL(rs("单位电话").Value)
                    txt(8).Text = zlCommFun.NVL(rs("单位邮编").Value)
                    txt(9).Text = zlCommFun.NVL(rs("家庭地址").Value)
                    txt(10).Text = zlCommFun.NVL(rs("家庭电话").Value)
                    txt(11).Text = zlCommFun.NVL(rs("户口邮编").Value)
                    
                    zlControl.CboLocate cbo(0), zlCommFun.NVL(rs("性别").Value)
                    zlControl.CboLocate cbo(1), zlCommFun.NVL(rs("费别").Value)
                    zlControl.CboLocate cbo(2), zlCommFun.NVL(rs("医疗付款方式").Value)
                    zlControl.CboLocate cbo(3), zlCommFun.NVL(rs("国籍").Value)
                    zlControl.CboLocate cbo(4), zlCommFun.NVL(rs("民族").Value)
                    zlControl.CboLocate cbo(5), zlCommFun.NVL(rs("婚姻状况").Value)
                    zlControl.CboLocate cbo(6), zlCommFun.NVL(rs("职业").Value)
                    cmdOK.Tag = "Changed"
                    txt(Index).Tag = ""
                    
                    
                End If
                
            End If
        End If
        
        LocationObj txt(Index)
    '------------------------------------------------------------------------------------------------------------------
    Case 6
    
        gstrSQL = GetPublicSQL(SQL.合约单位选择)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        If ShowPubSelect(Me, txt(Index), 3, "编码,900,0,1;名称,1500,0,1;简码,900,0,1;地址,3000,0,1", Me.Name & "\合约单位选择", "请在下表中选择一个合约单位", rsData, rs, 8790, 4500, , Val(cmd(Index).Tag)) = 1 Then
        
            txt(Index).Text = zlCommFun.NVL(rs("名称").Value)
            cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value, 0)
            cmdOK.Tag = "Changed"
            txt(Index).Tag = ""
        End If
        
        LocationObj txt(Index)
        
    '------------------------------------------------------------------------------------------------------------------
    Case 12
    
        gstrSQL = GetPublicSQL(SQL.麻醉方式选择)
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
 
        If ShowPubSelect(Me, txt(Index), 2, "编码,900,0,;名称,2400,0,;麻醉类型,900,0,", Me.Name & "\麻醉方式选择", "请从下表中选择一个麻醉方式", rsData, rs, 8790, 4500, , Val(cmd(0).Tag)) = 1 Then
            If Val(cmd(Index).Tag) <> zlCommFun.NVL(rs("ID")) Then

                txt(Index).Text = zlCommFun.NVL(rs("名称").Value)
                cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value)
                txt(Index).Tag = ""
                
                usrSaveItem.麻醉方式 = txt(Index).Text
                
                DataChanged = True

            End If
        End If
        
        LocationObj txt(Index)

    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((ParamInfo.系统号) / 100))
End Sub

Private Sub cbo_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    
    If mblnStartUp Then Exit Sub
    
    cmdOK.Tag = "Changed"
    
    If Index = 8 And cbo(Index).ListIndex > -1 Then
        
        '申请医生
        gstrSQL = GetPublicSQL(SQL.科室医生人员)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, cbo(8).ItemData(cbo(Index).ListIndex))
        Call AddComboData(cbo(9), rs)
        If cbo(9).ListCount > 0 Then cbo(9).ListIndex = 0
            
    End If
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub


Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdMore_Click()
    '
    If cmdMore.Caption = ">>" Then
        cmdMore.Caption = "<<"
        
        fra(1).Visible = True
        
        fra(2).Top = fra(2).Top + fra(1).Height
        picButton.Top = picButton.Top + fra(1).Height
        Me.Height = Me.Height + fra(1).Height
        
    Else
        cmdMore.Caption = ">>"
        
        fra(1).Visible = False
        
        fra(2).Top = fra(2).Top - fra(1).Height
        picButton.Top = picButton.Top - fra(1).Height
        Me.Height = Me.Height - fra(1).Height
    End If
    
End Sub

Private Sub cmdOK_Click()
    If cmdOK.Tag <> "" Then
        
        If ExecuteCommand("校验数据") = False Then Exit Sub
        If ExecuteCommand("保存数据") = False Then Exit Sub
        
        mblnOK = True

    End If
    
    cmdOK.Tag = ""
    Unload Me
End Sub

Private Sub dtp_Change(Index As Integer)
    cmdOK.Tag = "Changed"
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsfAdvice = Nothing
End Sub

Private Sub mclsVsfAdvice_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsfAdvice.RowData(Row)) = 0)
End Sub

Private Sub txt_Change(Index As Integer)
    cmdOK.Tag = "Changed"
    
    Select Case Index
    Case 0, 12
        txt(Index).Tag = "Changed"
    End Select
    
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 0, 5, 6, 9, 12, 13
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 12
        If KeyCode = vbKeyDelete Then
            KeyCode = 0
            txt(Index).Text = ""
            cmd(Index).Tag = ""
            txt(Index).Tag = ""
            usrSaveItem.麻醉方式 = ""
        End If
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strInput As String
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strText As String
    Dim strTmp As String
    Dim bytMode As Byte
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        '如果是在病人姓名中按了Enter,则要查找历史数据
        
        If txt(Index).Tag = "Changed" Then
            
            If InStr(txt(Index).Text, "'") Then
                ShowSimpleMsg "输入字符中有非法字符 ' ！"
                Exit Sub
            End If
                
            Select Case Index
            '----------------------------------------------------------------------------------------------------------
            Case 0

                Select Case UCase(Left(txt(Index).Text, 1))
                Case "-", "A"                 '病人id,就诊卡号
                    strInput = strInput & " AND C.病人id=" & Val(Mid(txt(Index).Text, 2))
                Case "+", "B"                 '住院号
                    strInput = strInput & " AND C.住院号=" & IIf(IsNumeric(Mid(txt(Index).Text, 2)), Mid(txt(Index).Text, 2), "0")
                Case "*", "D"                 '门诊号
                    strInput = strInput & " AND C.门诊号=" & IIf(IsNumeric(Mid(txt(Index).Text, 2)), Mid(txt(Index).Text, 2), "0")
                Case "/", "C"                 '当前床号
                    strInput = strInput & " AND C.当前床号=" & Val(Mid(txt(Index).Text, 2))
                End Select
                
                If strInput <> "" Then
                    gstrSQL = GetPublicSQL(SQL.人员过滤选择, strInput)
                    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
'                    If ShowPubSelect(Me, txt(Index), 2, "姓名,1200,0,0;性别,810,0,0;出生日期,1200,0,0;婚姻状况,900,0,0;身份证号,1500,0,0", Me.Name & "\人员过滤选择", "请从下面选择一个人员", rsData, rs, 8790, 4500) = 1 Then
                    If rs.BOF = False Then
                        txt(Index).Text = zlCommFun.NVL(rs("姓名"))
                        txt(1).Text = zlCommFun.NVL(rs("身份证号"))
                        txt(2).Text = zlCommFun.NVL(rs("年龄"))
                        
                        If Val(zlCommFun.NVL(rs("主页id"))) > 0 Then
                            lbl(5).Tag = Val(zlCommFun.NVL(rs("主页id")))
                            lbl(5).Caption = "住院号(&N)"
                            txt(3).Text = zlCommFun.NVL(rs("住院号"))
                            mint病人来源 = 2
                        Else
                            lbl(5).Tag = 0
                            lbl(5).Caption = "门诊号(&N)"
                            txt(3).Text = zlCommFun.NVL(rs("门诊号"))
                            mint病人来源 = 1
                        End If
                        
                        txt(4).Text = zlCommFun.NVL(rs("联系人电话").Value)
                        txt(5).Text = zlCommFun.NVL(rs("联系人姓名").Value)
                        txt(6).Text = zlCommFun.NVL(rs("工作单位").Value)
                        cmd(6).Tag = zlCommFun.NVL(rs("合同单位ID").Value)
                        txt(7).Text = zlCommFun.NVL(rs("单位电话").Value)
                        txt(8).Text = zlCommFun.NVL(rs("单位邮编").Value)
                        txt(9).Text = zlCommFun.NVL(rs("家庭地址").Value)
                        txt(10).Text = zlCommFun.NVL(rs("家庭电话").Value)
                        txt(11).Text = zlCommFun.NVL(rs("户口邮编").Value)
                        
                        cmd(0).Tag = zlCommFun.NVL(rs("ID"))
                        
                        zlControl.CboLocate cbo(0), zlCommFun.NVL(rs("性别").Value)
                        zlControl.CboLocate cbo(1), zlCommFun.NVL(rs("费别").Value)
                        zlControl.CboLocate cbo(2), zlCommFun.NVL(rs("医疗付款方式").Value)
                        zlControl.CboLocate cbo(3), zlCommFun.NVL(rs("国籍").Value)
                        zlControl.CboLocate cbo(4), zlCommFun.NVL(rs("民族").Value)
                        zlControl.CboLocate cbo(5), zlCommFun.NVL(rs("婚姻状况").Value)
                        zlControl.CboLocate cbo(6), zlCommFun.NVL(rs("职业").Value)
                        cmdOK.Tag = "Changed"
                    Else
                        cmd(0).Tag = ""
                        mint病人来源 = 1
                    End If
                End If
            '----------------------------------------------------------------------------------------------------------
            Case 6
            
                strInput = "%" & UCase(txt(Index).Text) & "%"
                
                gstrSQL = GetPublicSQL(SQL.合约单位过滤)
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strInput)
                If ShowPubSelect(Me, txt(Index), 2, "名称,1800,0,0;编码,900,0,0;简码,900,0,0;联系人,900,0,0;电话,1200,0,0", Me.Name & "\合约单位过滤", "请从下面选择一个合约单位", rsData, rs, 8790, 4500) = 1 Then
                                    
                    txt(Index).Text = zlCommFun.NVL(rs("名称"))
                    cmd(Index).Tag = zlCommFun.NVL(rs("ID"))
                    cmdOK.Tag = "Changed"
                Else
                    cmd(Index).Tag = ""
                End If
            
            '----------------------------------------------------------------------------------------------------------
            Case 12
                    

                txt(Index).Tag = ""
                
                strText = UCase(txt(Index).Text)
                bytMode = GetApplyMode(strText)

                strText = strText & "%"
                If ParamInfo.项目输入匹配方式 = 1 Then
                    strTmp = strText
                Else
                    strTmp = "%" & strText
                End If
                
                gstrSQL = GetPublicSQL(SQL.麻醉方式过滤, bytMode)
                
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)
                If ShowPubSelect(Me, txt(Index), 2, "编码,990,0,1;名称,1500,0,0;麻醉类型,900,0,0", Me.Name & "\麻醉方式过滤", "请从下面选择一个麻醉方式", rsData, rs, , , , Val(cmd(Index).Tag)) = 1 Then
                    If Val(cmd(Index).Tag) <> zlCommFun.NVL(rs("ID")) Then
            
                        txt(Index).Text = zlCommFun.NVL(rs("名称").Value)
                        cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value)
                        
                        DataChanged = True
                        
                        usrSaveItem.麻醉方式 = txt(Index).Text
                        
                    End If
                Else
                    txt(Index).Text = usrSaveItem.麻醉方式
                    txt(Index).Tag = ""
                    Exit Sub
                End If
            End Select
            
            txt(Index).Tag = ""
        End If
        
        zlCommFun.PressKey vbKeyTab
        
        Select Case Index
        Case 0, 6, 12
            zlCommFun.PressKey vbKeyTab
        End Select
        
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 0, 5, 6, 9, 12, 13
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    If Cancel Then Exit Sub

    Select Case Index
    Case 12
        If (txt(Index).Tag = "Changed") Then
            txt(Index).Text = usrSaveItem.麻醉方式
            txt(Index).Tag = ""
        End If
    End Select
    
End Sub


Private Sub vsfAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call mclsVsfAdvice.AfterEdit(Row, Col)
    
    With vsfAdvice
        Select Case Col
        Case .ColIndex("缺省")
            If Abs(Val(.Cell(flexcpText, Row, Col, Row, Col))) = 1 Then
                .Cell(flexcpText, 1, Col, .Rows - 1, Col) = 0
                .Cell(flexcpText, Row, Col, Row, Col) = 1
            End If
        End Select
    End With
    
    DataChanged = True
End Sub

Private Sub vsfAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsfAdvice.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsfAdvice_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsfAdvice.AppendRows = True
End Sub

Private Sub vsfAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsVsfAdvice.AppendRows = True
End Sub

Private Sub vsfAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsfAdvice.BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsfAdvice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    With vsfAdvice
        If Col = .ColIndex("手术名称") Then

            gstrSQL = GetPublicSQL(SQL.手术项目选择)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)

            If ShowPubSelect(Me, vsfAdvice, 3, "编码,1200,0,;名称,2700,0,", Me.Name & "\手术项目选择", "请从下表中选择一个手术项目", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then
                If mclsVsfAdvice.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                    ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                    Exit Sub
                End If
    
                .EditText = zlCommFun.NVL(rs("名称").Value)
                .TextMatrix(Row, mclsVsfAdvice.ColIndex("手术名称")) = zlCommFun.NVL(rs("名称").Value)
                .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                
                Call ExecuteCommand("读取执行科室")
                
                DataChanged = True
            End If
        End If
    End With
End Sub

Private Sub vsfAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mclsVsfAdvice.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsfAdvice_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    Dim bytMode As Byte
    
    With vsfAdvice
        If KeyCode = vbKeyReturn Then
            If Col = .ColIndex("手术名称") Then
                
                If InStr(.EditText, "'") > 0 Then
                    KeyCode = 0
                    .EditText = ""
                    Exit Sub
                End If

                strText = UCase(.EditText)
                bytMode = GetApplyMode(strText)

                gstrSQL = GetPublicSQL(SQL.手术项目过滤, bytMode)

                strText = strText & "%"
                If ParamInfo.项目输入匹配方式 = 1 Then
                    strTmp = strText
                Else
                    strTmp = "%" & strText
                End If
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)

                If ShowPubSelect(Me, vsfAdvice, 2, "编码,1200,0,;名称,2700,0,", Me.Name & "\手术项目过滤", "请从下表中选择一个手术项目", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then

                    If mclsVsfAdvice.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                        ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                        Exit Sub
                    End If

                    .EditText = zlCommFun.NVL(rs("名称").Value)
                    .TextMatrix(Row, .ColIndex("手术名称")) = zlCommFun.NVL(rs("名称").Value)
                    
                    .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)

                    DataChanged = True

                Else
                    KeyCode = 0

                    .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                    .EditText = .Cell(flexcpData, Row, Col)
                    .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)

                End If
            End If
        Else
            DataChanged = True
        End If
    End With
End Sub

Private Sub vsfAdvice_KeyPress(KeyAscii As Integer)
    Call mclsVsfAdvice.KeyPress(KeyAscii)
End Sub

Private Sub vsfAdvice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call mclsVsfAdvice.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsfAdvice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsfAdvice.AutoAddRow(vsfAdvice.MouseRow, vsfAdvice.MouseCol)
    End Select
End Sub

Private Sub vsfAdvice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsfAdvice.EditSelAll
End Sub

Private Sub vsfAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsfAdvice.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsfAdvice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsfAdvice.ValidateEdit(Col, Cancel)
End Sub
