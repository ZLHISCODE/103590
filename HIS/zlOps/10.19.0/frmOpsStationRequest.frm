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
      TabIndex        =   63
      Top             =   5520
      Width           =   8715
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   7530
         TabIndex        =   58
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   6375
         TabIndex        =   57
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
         TabIndex        =   65
         Top             =   1140
         Width           =   15
      End
      Begin VB.Label lbl总量单位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   6150
         TabIndex        =   64
         Top             =   1140
         Width           =   15
      End
   End
   Begin VB.Frame fra 
      Caption         =   "申请信息"
      Height          =   2445
      Index           =   2
      Left            =   15
      TabIndex        =   62
      Top             =   3075
      Width           =   8700
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   8
         ItemData        =   "frmOpsStationRequest.frx":000C
         Left            =   6450
         List            =   "frmOpsStationRequest.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   1350
         Width           =   2115
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H80000004&
         Caption         =   "紧急(&2)"
         Height          =   225
         Index           =   0
         Left            =   6465
         TabIndex        =   56
         Top             =   2100
         Width           =   945
      End
      Begin VB.CommandButton cmd 
         Caption         =   "…"
         Height          =   285
         Index           =   12
         Left            =   5025
         TabIndex        =   45
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
         TabIndex        =   49
         Top             =   600
         Width           =   2115
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   12
         Left            =   1110
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   2040
         Width           =   3900
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   9
         Left            =   6450
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   1725
         Width           =   2115
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   13
         Left            =   6450
         MaxLength       =   100
         TabIndex        =   47
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
         Format          =   80216067
         CurrentDate     =   38022
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfAdvice 
         Height          =   1710
         Left            =   1110
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   52
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
         TabIndex        =   48
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
         TabIndex        =   43
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
         TabIndex        =   46
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
         TabIndex        =   50
         Top             =   990
         Width           =   990
      End
   End
   Begin VB.Frame fra 
      Caption         =   "其他信息"
      Height          =   1815
      Index           =   1
      Left            =   15
      TabIndex        =   61
      Top             =   1230
      Visible         =   0   'False
      Width           =   8700
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   11
         Left            =   7320
         MaxLength       =   6
         TabIndex        =   40
         Top             =   1365
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   10
         Left            =   4695
         MaxLength       =   20
         TabIndex        =   38
         Top             =   1365
         Width           =   1545
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   9
         Left            =   1110
         MaxLength       =   50
         TabIndex        =   36
         Top             =   1365
         Width           =   2475
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   7320
         MaxLength       =   6
         TabIndex        =   34
         Top             =   975
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   4695
         MaxLength       =   20
         TabIndex        =   32
         Top             =   975
         Width           =   1545
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   6
         Left            =   1110
         MaxLength       =   100
         TabIndex        =   29
         Top             =   975
         Width           =   2190
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   6
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   585
         Width           =   2475
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   5
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   195
         Width           =   1275
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   4
         Left            =   4695
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   210
         Width           =   1545
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   3
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   210
         Width           =   2475
      End
      Begin VB.CommandButton cmd 
         Caption         =   "…"
         Height          =   270
         Index           =   6
         Left            =   3300
         TabIndex        =   30
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
         TabIndex        =   27
         Top             =   585
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   4695
         MaxLength       =   20
         TabIndex        =   25
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
         TabIndex        =   39
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
         TabIndex        =   37
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
         TabIndex        =   35
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
         TabIndex        =   33
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
         TabIndex        =   31
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
         TabIndex        =   28
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
         TabIndex        =   16
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
         TabIndex        =   18
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
         TabIndex        =   22
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
         TabIndex        =   20
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
         TabIndex        =   26
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
         TabIndex        =   24
         Top             =   645
         Width           =   990
      End
   End
   Begin VB.Frame fra 
      Caption         =   "基本信息"
      Height          =   1125
      Index           =   0
      Left            =   15
      TabIndex        =   60
      Top             =   60
      Width           =   8700
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1125
         MaxLength       =   20
         TabIndex        =   1
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
         TabIndex        =   6
         Top             =   210
         Width           =   945
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   3825
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   1710
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   2
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   1905
      End
      Begin VB.CommandButton cmdMore 
         Caption         =   ">>"
         Height          =   300
         Left            =   8280
         TabIndex        =   15
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
         TabIndex        =   4
         Top             =   210
         Width           =   1710
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   1125
         MaxLength       =   20
         TabIndex        =   10
         ToolTipText     =   "数字为就诊卡号、“－”打头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“#”收费单据号"
         Top             =   600
         Width           =   1590
      End
      Begin VB.CommandButton cmd 
         Caption         =   "…"
         Height          =   285
         Index           =   0
         Left            =   2430
         TabIndex        =   2
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
         TabIndex        =   11
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
         TabIndex        =   0
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
         TabIndex        =   7
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
         TabIndex        =   5
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
         TabIndex        =   3
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
         TabIndex        =   9
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
