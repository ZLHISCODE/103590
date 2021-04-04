VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmParaSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585
   Icon            =   "frmParaSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdSet 
      Caption         =   "设备(&V)"
      Height          =   350
      Left            =   5970
      TabIndex        =   65
      ToolTipText     =   "设置允许接入的设备"
      Top             =   6855
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8310
      TabIndex        =   64
      Top             =   6855
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7140
      TabIndex        =   63
      Top             =   6855
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   150
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1100
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6585
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   11615
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabMaxWidth     =   2822
      TabCaption(0)   =   "接收服务"
      TabPicture(0)   =   "frmParaSet.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frmReceiveSet"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkStorage"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "WorkList"
      TabPicture(1)   =   "frmParaSet.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmWorkList"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chkDWL"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Query/Retrieve"
      TabPicture(2)   =   "frmParaSet.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmQueryRetrieve"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "chkQuery"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "本地数据库"
      TabPicture(3)   =   "frmParaSet.frx":0060
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame6 
         Caption         =   "清空临时表"
         Height          =   2535
         Left            =   120
         TabIndex        =   55
         Top             =   700
         Width           =   9255
         Begin VB.CommandButton cmdClear 
            Caption         =   "立即清空"
            Height          =   350
            Left            =   240
            TabIndex        =   62
            Top             =   2040
            Width           =   1100
         End
         Begin VB.TextBox txtClearInterval 
            Height          =   300
            Left            =   960
            MaxLength       =   3
            TabIndex        =   61
            Top             =   1485
            Width           =   975
         End
         Begin VB.CheckBox chkAutoClear 
            Caption         =   "间隔                           天，自动清空"
            Height          =   375
            Left            =   240
            TabIndex        =   60
            Top             =   1440
            Width           =   4575
         End
         Begin VB.Frame Frame5 
            Caption         =   "数据库表"
            Height          =   855
            Left            =   240
            TabIndex        =   56
            Top             =   360
            Width           =   8775
            Begin VB.CheckBox chkClearTempTB 
               Caption         =   "影像接收序列"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   57
               Top             =   360
               Width           =   1815
            End
            Begin VB.CheckBox chkClearTempTB 
               Caption         =   "错误日志"
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   58
               Top             =   360
               Width           =   1815
            End
            Begin VB.CheckBox chkClearTempTB 
               Caption         =   "DICOM通讯日志"
               Height          =   255
               Index           =   2
               Left            =   5280
               TabIndex        =   59
               Top             =   360
               Width           =   2775
            End
         End
      End
      Begin VB.CheckBox chkStorage 
         Caption         =   "启动图像接收服务"
         Height          =   195
         Left            =   -74850
         TabIndex        =   0
         Top             =   500
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.Frame frmReceiveSet 
         Height          =   5715
         Left            =   -74850
         TabIndex        =   69
         Top             =   700
         Width           =   9255
         Begin VB.Frame Frame3 
            Caption         =   "自动匹配设置"
            Height          =   2055
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   8895
            Begin VB.CheckBox chkImageType 
               Caption         =   "根据图像类型拆分序列"
               Height          =   350
               Left            =   4440
               TabIndex        =   75
               Top             =   240
               Width           =   3015
            End
            Begin VB.CheckBox chkMatchStudyUID 
               Caption         =   "启用 ""检查UID"" 匹配"
               Height          =   350
               Left            =   120
               TabIndex        =   74
               Top             =   240
               Width           =   3015
            End
            Begin VB.Frame Frame2 
               Caption         =   "数据库项目"
               Height          =   1335
               Left            =   4440
               TabIndex        =   14
               Top             =   600
               Width           =   4365
               Begin VB.OptionButton optMatch 
                  Caption         =   "按 ""检查标识号"" 匹配"
                  Height          =   195
                  Index           =   2
                  Left            =   120
                  TabIndex        =   17
                  ToolTipText     =   "按检查标识号将病人和接收的影像进行匹配"
                  Top             =   960
                  Width           =   2775
               End
               Begin VB.OptionButton optMatch 
                  Caption         =   "按 ""病人标识号（门诊/住院号）"" 匹配"
                  Height          =   195
                  Index           =   1
                  Left            =   120
                  TabIndex        =   16
                  ToolTipText     =   "按病人标识号将病人和接收的影像进行匹配"
                  Top             =   600
                  Width           =   3975
               End
               Begin VB.OptionButton optMatch 
                  Caption         =   "按 ""检查号"" 匹配"
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  TabIndex        =   15
                  ToolTipText     =   "按检查号将病人和接收的影像进行匹配"
                  Top             =   240
                  Width           =   2265
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "图像项目"
               Height          =   1335
               Left            =   120
               TabIndex        =   10
               Top             =   600
               Width           =   4250
               Begin VB.OptionButton optImgMatch 
                  Caption         =   "Patient Name"
                  Height          =   255
                  Index           =   2
                  Left            =   240
                  TabIndex        =   13
                  Top             =   960
                  Width           =   2055
               End
               Begin VB.OptionButton optImgMatch 
                  Caption         =   "Accession Number"
                  Height          =   255
                  Index           =   1
                  Left            =   240
                  TabIndex        =   12
                  Top             =   600
                  Width           =   2055
               End
               Begin VB.OptionButton optImgMatch 
                  Caption         =   "Patient ID"
                  Height          =   255
                  Index           =   0
                  Left            =   240
                  TabIndex        =   11
                  Top             =   240
                  Width           =   2055
               End
            End
         End
         Begin VB.Frame frmAutoRoutSet 
            Caption         =   "自动路由设置"
            Height          =   2385
            Left            =   120
            TabIndex        =   18
            Top             =   3210
            Width           =   9015
            Begin VB.CommandButton cmdInsert 
               Caption         =   "添加(&A)"
               Height          =   350
               Left            =   1800
               TabIndex        =   26
               Top             =   1890
               Width           =   1100
            End
            Begin VB.CommandButton cmdModify 
               Caption         =   "修改(&M)"
               Height          =   350
               Left            =   3660
               TabIndex        =   27
               Top             =   1890
               Width           =   1100
            End
            Begin VB.CommandButton cmdDelete 
               Caption         =   "删除(&D)"
               Height          =   350
               Left            =   5520
               TabIndex        =   28
               Top             =   1890
               Width           =   1100
            End
            Begin VB.OptionButton optType 
               Caption         =   "影像类别(&S)"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   20
               Top             =   1455
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton optType 
               Caption         =   "检查设备(&R)"
               Height          =   255
               Index           =   2
               Left            =   3120
               TabIndex        =   22
               Top             =   1455
               Width           =   1335
            End
            Begin VB.ComboBox cobCondition 
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               Left            =   4530
               TabIndex        =   23
               Top             =   1425
               Width           =   1365
            End
            Begin VB.ComboBox cobCondition 
               Height          =   315
               Index           =   1
               Left            =   1530
               TabIndex        =   21
               Top             =   1440
               Width           =   1365
            End
            Begin VB.ComboBox cobDestination 
               Height          =   315
               Left            =   7290
               TabIndex        =   25
               Top             =   1425
               Width           =   1605
            End
            Begin MSFlexGridLib.MSFlexGrid MSFAutoRout 
               Height          =   1125
               Left            =   150
               TabIndex        =   19
               Top             =   240
               Width           =   8775
               _ExtentX        =   15478
               _ExtentY        =   1984
               _Version        =   393216
               FixedCols       =   0
               SelectionMode   =   1
               AllowUserResizing=   1
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "目的设备(&B)"
               Height          =   180
               Left            =   6150
               TabIndex        =   24
               Top             =   1485
               Width           =   990
            End
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "…"
            Height          =   255
            Left            =   8760
            TabIndex        =   71
            TabStop         =   0   'False
            ToolTipText     =   "选择临时目录"
            Top             =   750
            Width           =   285
         End
         Begin VB.ComboBox cboEncode 
            Height          =   300
            ItemData        =   "frmParaSet.frx":007C
            Left            =   6240
            List            =   "frmParaSet.frx":0089
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   300
            Width           =   2835
         End
         Begin VB.TextBox txtItem 
            BackColor       =   &H80000009&
            DataField       =   "315"
            Height          =   300
            Index           =   0
            Left            =   1320
            MaxLength       =   5
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   280
            Width           =   855
         End
         Begin VB.ComboBox cboDevice 
            Height          =   300
            ItemData        =   "frmParaSet.frx":00AC
            Left            =   3315
            List            =   "frmParaSet.frx":00B9
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   300
            Width           =   1575
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   465
            Left            =   150
            TabIndex        =   70
            Top             =   1050
            Width           =   8925
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   1
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   200
            TabIndex        =   8
            Top             =   720
            Width           =   7740
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "临时目录(&T)"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   780
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "压缩方式(&Y)"
            Height          =   180
            Index           =   0
            Left            =   5160
            TabIndex        =   5
            Top             =   345
            Width           =   990
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "监听端口(&P)"
            Height          =   180
            Left            =   240
            TabIndex        =   1
            Top             =   345
            Width           =   990
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "存储设备(&F)"
            Height          =   180
            Index           =   8
            Left            =   2280
            TabIndex        =   3
            Top             =   345
            Width           =   990
         End
      End
      Begin VB.CheckBox chkDWL 
         Caption         =   "启动 DICOM WorkList 服务"
         Height          =   255
         Left            =   -74880
         TabIndex        =   29
         Top             =   480
         Width           =   2775
      End
      Begin VB.Frame frmWorkList 
         Enabled         =   0   'False
         Height          =   5790
         Left            =   -74880
         TabIndex        =   68
         Top             =   705
         Width           =   9255
         Begin VB.CommandButton cmdResetWLResult 
            Caption         =   "恢复默认结果"
            Height          =   350
            Left            =   6240
            TabIndex        =   38
            Top             =   600
            Width           =   1335
         End
         Begin VB.Frame Frame8 
            Caption         =   "结果集设置"
            Height          =   4695
            Left            =   120
            TabIndex        =   39
            Top             =   960
            Width           =   8895
            Begin VB.CheckBox chkUseResult 
               Caption         =   "选择使用该结果："
               Height          =   255
               Left            =   240
               TabIndex        =   41
               Top             =   3000
               Width           =   1935
            End
            Begin VB.Frame frmSetResult 
               Height          =   1575
               Left            =   120
               TabIndex        =   72
               Top             =   3000
               Width           =   8655
               Begin VB.CommandButton cmdBuildResult 
                  Appearance      =   0  'Flat
                  Caption         =   "…"
                  Height          =   235
                  Index           =   0
                  Left            =   7950
                  MaskColor       =   &H80000000&
                  Style           =   1  'Graphical
                  TabIndex        =   73
                  Top             =   765
                  Width           =   315
               End
               Begin VB.TextBox txtResult 
                  Height          =   300
                  Index           =   1
                  Left            =   1200
                  TabIndex        =   47
                  Top             =   1080
                  Width           =   7095
               End
               Begin VB.CheckBox chkResult 
                  Caption         =   "是否递增"
                  Height          =   255
                  Left            =   7320
                  TabIndex        =   43
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.TextBox txtResult 
                  Height          =   300
                  Index           =   0
                  Left            =   1200
                  TabIndex        =   45
                  Top             =   720
                  Width           =   7095
               End
               Begin VB.Label Label12 
                  Caption         =   "强制结果值"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   46
                  Top             =   1110
                  Width           =   975
               End
               Begin VB.Label Label11 
                  Caption         =   "返回值"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   44
                  Top             =   743
                  Width           =   735
               End
               Begin VB.Label lblResult 
                  Caption         =   "结果集："
                  Height          =   255
                  Left            =   120
                  TabIndex        =   42
                  Top             =   360
                  Width           =   7215
               End
            End
            Begin MSFlexGridLib.MSFlexGrid MSFResult 
               Height          =   2535
               Left            =   120
               TabIndex        =   40
               Top             =   360
               Width           =   8655
               _ExtentX        =   15266
               _ExtentY        =   4471
               _Version        =   393216
               AllowBigSelection=   0   'False
               SelectionMode   =   1
               AllowUserResizing=   1
            End
         End
         Begin VB.CheckBox chkForceResult 
            Caption         =   "使用强制结果"
            Height          =   255
            Left            =   3600
            TabIndex        =   35
            Top             =   660
            Width           =   1515
         End
         Begin VB.CheckBox chkModel 
            Caption         =   "按检查设备过滤"
            Height          =   225
            Left            =   3600
            TabIndex        =   34
            Top             =   278
            Width           =   1755
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   6
            Left            =   7080
            MaxLength       =   4
            TabIndex        =   37
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtDWLLocalAE 
            Height          =   300
            Left            =   1080
            MaxLength       =   20
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   637
            Width           =   1695
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   4
            Left            =   1080
            MaxLength       =   5
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "检索最近                天的申请"
            Height          =   195
            Left            =   6210
            TabIndex        =   36
            Top             =   300
            Width           =   2355
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "本机AE"
            Height          =   180
            Left            =   195
            TabIndex        =   32
            Top             =   690
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "监听端口"
            Height          =   180
            Left            =   195
            TabIndex        =   30
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.CheckBox chkQuery 
         Caption         =   "启动 Query/Retrieve 查询服务"
         Height          =   255
         Left            =   -74850
         TabIndex        =   48
         Top             =   500
         Width           =   2955
      End
      Begin VB.Frame frmQueryRetrieve 
         Enabled         =   0   'False
         Height          =   975
         Left            =   -74880
         TabIndex        =   49
         Top             =   700
         Width           =   9255
         Begin VB.TextBox txtQueryAE 
            Height          =   300
            Left            =   4320
            MaxLength       =   20
            ScrollBars      =   2  'Vertical
            TabIndex        =   53
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   5
            Left            =   1245
            MaxLength       =   5
            ScrollBars      =   2  'Vertical
            TabIndex        =   51
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox chkAcceptCGET 
            Caption         =   "支持C-GET"
            Height          =   255
            Left            =   7200
            TabIndex        =   54
            Top             =   380
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "本机AE"
            Height          =   180
            Left            =   3600
            TabIndex        =   52
            Top             =   420
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "监听端口"
            Height          =   180
            Left            =   360
            TabIndex        =   50
            Top             =   420
            Width           =   720
         End
      End
   End
End
Attribute VB_Name = "frmParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ifOK As Boolean
Private mblnchkResultFocus As Boolean
Private mblnchkUseResultFocus As Boolean

Private aDevices() As Variant
Private mintMaxDevs As Integer
Private mblnModifyMWLResult As Boolean          '记录是否修改了Worklist返回值的设置

Public Function ShowMe(objParent As Object, Optional iMaxDevs As Integer = 2) As Boolean
    mintMaxDevs = iMaxDevs
    Me.Show vbModal, objParent
    ShowMe = ifOK
End Function

Private Sub cboDevice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboEncode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkAcceptCGET_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkAutoClear_Click()
    If chkAutoClear.value = 1 Then
        txtClearInterval.Enabled = True
    Else
        txtClearInterval.Enabled = False
    End If
End Sub

Private Sub chkClearTempTB_Click(Index As Integer)
    Dim i As Integer
    gstrClearTable = "" '清空表列表
    For i = 0 To 2
        If chkClearTempTB(i).value = 1 Then
            gstrClearTable = gstrClearTable & IIf(Trim(gstrClearTable) = "", "", ";") & chkClearTempTB(i).Caption
        End If
    Next i
End Sub

Private Sub chkClearTempTB_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkDWL_Click()
    If Me.chkDWL.value = 0 Then
        Me.frmWorkList.Enabled = False
    Else
        Me.frmWorkList.Enabled = True
    End If
End Sub

Private Sub chkDWL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub




Private Sub chkQuery_Click()
    If Me.chkQuery.value = 0 Then
        Me.frmQueryRetrieve.Enabled = False
    Else
        Me.frmQueryRetrieve.Enabled = True
    End If
End Sub

Private Sub chkQuery_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub


Private Sub chkResult_Click()
    If mblnchkResultFocus Then
        mblnModifyMWLResult = True
        subChangeMSFResult
    End If
End Sub

Private Sub chkResult_GotFocus()
    mblnchkResultFocus = True
End Sub

Private Sub chkResult_LostFocus()
    mblnchkResultFocus = False
End Sub


Private Sub chkStorage_Click()
    If Me.chkStorage.value = 0 Then
        Me.frmReceiveSet.Enabled = False
    Else
        Me.frmReceiveSet.Enabled = True
    End If
    
End Sub

Private Sub chkUseResult_Click()
    If mblnchkUseResultFocus Then
        mblnModifyMWLResult = True
        subChangeMSFResult
    End If
    If chkUseResult.value = 1 Then
        frmSetResult.Enabled = True
    Else
        frmSetResult.Enabled = False
    End If
End Sub

Private Sub chkUseResult_GotFocus()
    mblnchkUseResultFocus = True
End Sub

Private Sub chkUseResult_LostFocus()
    mblnchkUseResultFocus = False
End Sub

Private Sub cmdBuildResult_Click(Index As Integer)
    frmBuildResult.strReturnString = ""
    frmBuildResult.txtBuildResult.Text = Me.txtResult(Index).Text
    frmBuildResult.Show 1, Me
    If frmBuildResult.strReturnString <> "" Then
        Me.txtResult(Index).Text = frmBuildResult.strReturnString
        mblnModifyMWLResult = True
        subChangeMSFResult
    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Public Sub cmdClear_Click()
    subClearTempTable True
End Sub

Private Sub cmdDelete_Click()
    '删除自动路由数组中的值
    Dim iRow As Integer
    Dim i As Integer
    iRow = MSFAutoRout.RowSel
    '移动数组内容
    For i = iRow + 1 To UBound(aAutoRoutSetting)
        aAutoRoutSetting(i - 1).Type = aAutoRoutSetting(i).Type
        aAutoRoutSetting(i - 1).strCondition = aAutoRoutSetting(i).strCondition
        aAutoRoutSetting(i - 1).strFTPDeviceNo = aAutoRoutSetting(i).strFTPDeviceNo
    Next
    '修改数组大小
    If UBound(aAutoRoutSetting) = 0 Then Exit Sub
    ReDim Preserve aAutoRoutSetting(0 To UBound(aAutoRoutSetting) - 1)
    '刷新自动路由规则显示列表
    subFillMsfAutoRout
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdInsert_Click()
    '检查输入是否合法
    Dim iType As Integer
    iType = IIf(optType(1).value = True, 1, 2)
    If cobDestination.Text = "" Then MsgBox "请输入自动路由的目的设备。": Exit Sub
    If cobCondition(iType).Text = "" Then MsgBox IIf(iType = 1, "请输入影像类别", "请输入检查设备"): Exit Sub
    '向自动路由规则数组添加新规则
    Dim iCount As Integer
    iCount = UBound(aAutoRoutSetting) + 1
    
    ReDim a(2)
    a(1).Type = 4
    a(1).strCondition = "DFDFD"
    ReDim Preserve a(5)
    ReDim Preserve a(3)
    
    ReDim Preserve aAutoRoutSetting(0 To iCount)
    aAutoRoutSetting(iCount).Type = iType
    aAutoRoutSetting(iCount).strCondition = cobCondition(iType).Text
    aAutoRoutSetting(iCount).strFTPDeviceNo = GetDeviceNameNum(aDevices, cobDestination.Text, 1)
    '向规则列表添加新规则
    With MSFAutoRout
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = IIf(aAutoRoutSetting(iCount).Type = 1, "影像类别", "检查设备")
        .TextMatrix(.Rows - 1, 1) = aAutoRoutSetting(iCount).strCondition
        .TextMatrix(.Rows - 1, 2) = GetDeviceNameNum(aDevices, aAutoRoutSetting(iCount).strFTPDeviceNo, 0)
    End With
End Sub

Private Sub cmdModify_Click()
    '修改自动路由数组
    Dim iType  As Integer
    Dim iRow As Integer
    iRow = MSFAutoRout.RowSel
    iType = IIf(optType(1).value = True, 1, 2)
    aAutoRoutSetting(iRow).Type = iType
    aAutoRoutSetting(iRow).strCondition = cobCondition(iType).Text
    aAutoRoutSetting(iRow).strFTPDeviceNo = GetDeviceNameNum(aDevices, cobDestination.Text, 1)
    '修改自动路由规则列表
    MSFAutoRout.TextMatrix(iRow, 0) = IIf(aAutoRoutSetting(iRow).Type = 1, "影像类别", "检查设备")
    MSFAutoRout.TextMatrix(iRow, 1) = aAutoRoutSetting(iRow).strCondition
    MSFAutoRout.TextMatrix(iRow, 2) = GetDeviceNameNum(aDevices, aAutoRoutSetting(iRow).strFTPDeviceNo, 0)
End Sub

Private Sub CmdOK_Click()
    Dim strSQL As String
    Dim i As Integer
    
    On Error GoTo DBError
    '保存图像接收设置
    If Me.chkStorage.value = 1 Then
        If Len(Trim(txtItem(0))) = 0 Then
            MsgBox "请输入端口号！", vbInformation, gstrSysName
            txtItem(0).SetFocus: Exit Sub
        End If
        If Len(Trim(txtItem(1))) = 0 Then
            MsgBox "请输入临时目录！", vbInformation, gstrSysName
            txtItem(1).SetFocus: Exit Sub
        End If
        If LenB(StrConv(Trim(txtItem(1).Text), vbFromUnicode)) > txtItem(1).MaxLength Then
            MsgBox "临时目录超长（最多" & txtItem(1).MaxLength & "个字符或" & CInt(txtItem(1).MaxLength / 2) & "个汉字）！", vbInformation, gstrSysName
            txtItem(1).SetFocus: Exit Sub
        End If
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "端口", txtItem(0)
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "设备号", aDevices(0, cboDevice.ListIndex)
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "临时目录", txtItem(1)
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "无损压缩", cboEncode.ListIndex
        For i = 0 To optMatch.count - 1
            If optMatch(i).value Then Exit For
        Next
        If i > optMatch.count - 1 Then i = 0
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "数据库匹配方式", i
        
        For i = 0 To optImgMatch.count - 1
            If optImgMatch(i).value Then Exit For
        Next
        If i > optImgMatch.count - 1 Then i = 0
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "图像匹配方式", i
        
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "启用检查UID匹配", IIf(chkMatchStudyUID.value, 1, 0)
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "根据图像类型拆分序列", IIf(chkImageType.value, 1, 0)
        
        '保存自动路由设置
        Dim strAutoRoutSet As String
        If UBound(aAutoRoutSetting) >= 1 Then
            strAutoRoutSet = aAutoRoutSetting(1).Type
            strAutoRoutSet = strAutoRoutSet & "," & aAutoRoutSetting(1).strCondition
            strAutoRoutSet = strAutoRoutSet & "," & aAutoRoutSetting(1).strFTPDeviceNo
        End If
        For i = 2 To UBound(aAutoRoutSetting)
            strAutoRoutSet = strAutoRoutSet & "," & aAutoRoutSetting(i).Type
            strAutoRoutSet = strAutoRoutSet & "," & aAutoRoutSetting(i).strCondition
            strAutoRoutSet = strAutoRoutSet & "," & aAutoRoutSetting(i).strFTPDeviceNo
        Next
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "自动路由", strAutoRoutSet
    End If
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "启动接收", Me.chkStorage.value
    
    '保存WorkList的设置
    If Me.chkDWL.value = 1 Then
        '输入正确性检查
        If Len(Trim(txtItem(4))) = 0 Then
            MsgBox "请输入WorkList端口号！", vbInformation, gstrSysName
            txtItem(4).SetFocus: Exit Sub
        End If
        If Len(Trim(txtDWLLocalAE)) = 0 Then
            MsgBox "请输入WorkList的本机AE名称！", vbInformation, gstrSysName
            txtDWLLocalAE.SetFocus: Exit Sub
        End If
        '保存输入参数
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "WorkList端口", txtItem(4)
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "WorkList本机AE", txtDWLLocalAE
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "WorkList检索天数", Val(txtItem(6))
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "WorkList按设备过滤", IIf(chkModel.value, 1, 0)
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "WorkList使用强制结果", IIf(chkForceResult.value, 1, 0)
        
        '保存Worklist 返回值的修改
        If mblnModifyMWLResult = True And gcnAccess.State <> adStateClosed Then
            With Me.MSFResult
                For i = 1 To .Rows - 1
                    strSQL = "update 强制结果 set 数据值 = '" & .TextMatrix(i, 4) & "' , 强制结果值='" _
                             & .TextMatrix(i, 5) & "' , 是否递增 = " & .TextMatrix(i, 6) _
                             & ",被选择 = " & .TextMatrix(i, 3) & " where 组号 = '" _
                             & Mid(.TextMatrix(i, 0), 2, InStr(.TextMatrix(i, 0), ",") - 2) & "' and 元素号 = '" _
                             & Mid(.TextMatrix(i, 0), InStr(.TextMatrix(i, 0), ",") + 1, Len(.TextMatrix(i, 0)) - InStr(.TextMatrix(i, 0), ",") - 1) & "'"
                    gcnAccess.Execute strSQL
                Next i
            End With
        End If
    End If
    
    '保存Query/Retrieve的设置
    If Me.chkQuery.value = 1 Then
        '输入正确性检查
        If Len(Trim(txtItem(5))) = 0 Then
            MsgBox "请输入Query/Retrieve的端口号！", vbInformation, gstrSysName
            txtItem(5).SetFocus: Exit Sub
        End If
        If Len(Trim(txtQueryAE)) = 0 Then
            MsgBox "请输入Query/Retrieve的本机AE名称！", vbInformation, gstrSysName
            txtQueryAE.SetFocus: Exit Sub
        End If
        '保存输入参数
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "Query/Retrieve端口", txtItem(5)
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "Query/Retrieve本机AE", txtQueryAE
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "支持C-GET", Me.chkAcceptCGET.value
    End If
    
    '保存“本地数据库”的设置
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "自动清空临时表", gstrClearTable
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "自动清空临时表间隔天数", IIf(txtClearInterval.Enabled = True, Val(Me.txtClearInterval.Text), 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "自动清空临时表日期", Date
    
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "启动WorkList", Me.chkDWL.value
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "启动Query/Retrieve", Me.chkQuery.value
    ifOK = True
    Unload Me
    Exit Sub
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdResetWLResult_Click()
    Dim strSQL As String
    
    If gcnAccess.State = adStateClosed Then Exit Sub
    strSQL = "update 强制结果 set 数据值 = 默认值,强制结果值=默认强制结果,是否递增 = False,被选择 = 默认选择"
    gcnAccess.Execute strSQL
    subFillMsfResult
End Sub

Private Sub cmdSel_Click()
    Dim strTmp As String
    '得到路径
    strTmp = BrowPath(Me.hwnd, "请选定影像保存的临时目录：")
    '当用新的路径时才保存
    If strTmp <> "" Then
        If Mid(strTmp, Len(strTmp), 1) <> "\" Then strTmp = strTmp + "\"
        txtItem(1) = strTmp
    End If
End Sub

Private Sub cmdSet_Click()
    frmIPConfig.ShowEdit Me, mintMaxDevs
End Sub



Private Sub cobCondition_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cobDestination_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    Call CmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strExeRoom As String
    Dim strDeviceNO As String
    Dim iMatchStyle As Integer
    Dim iImgMatchStyle As Integer
    Dim strTemp As String
    Dim i As Integer
    
    ifOK = False
    
    On Error GoTo DBError
    gstrSQL = "Select 设备号,设备名 From 影像设备目录 Where 类型= [1]"
    Set rsTmp = OpenSQLRecord(gstrSQL, Me.Caption, 1)
    If rsTmp.EOF Then
        MsgBox "未定义影像存储设备，请到影像设备目录中设置！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    aDevices = rsTmp.GetRows: rsTmp.MoveFirst: strDeviceNO = rsTmp(0)
    Me.cboDevice.Clear
    Do While Not rsTmp.EOF
        cboDevice.AddItem Nvl(rsTmp(1))
        '填充自动路由设置中的目的设备下拉列表，黄捷
        cobDestination.AddItem Nvl(rsTmp(1))
        rsTmp.MoveNext
    Loop
    
    txtItem(0) = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "端口", 104)
    strDeviceNO = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "设备号", strDeviceNO)
    cboDevice.ListIndex = GetComboxIndex(aDevices, strDeviceNO)
    txtItem(1) = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "临时目录", "C:\TmpImage\")
    cboEncode.ListIndex = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "无损压缩", 0))
    iMatchStyle = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "数据库匹配方式", 0))
    optMatch(iMatchStyle).value = True
    iImgMatchStyle = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "图像匹配方式", 0))
    optImgMatch(iImgMatchStyle).value = True
    chkMatchStudyUID.value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "启用检查UID匹配", 1))
    chkImageType.value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "根据图像类型拆分序列", 0))
    
    
    '添加自动路由设置
    subFillMsfAutoRout
    '填充自动路由设置中，影像类别，和检查设备列表
    gstrSQL = "Select 编码 From 影像检查类别"
    OpenRecordset rsTmp, Me.Caption
    Do While Not rsTmp.EOF
        cobCondition(1).AddItem rsTmp(0)
        rsTmp.MoveNext
    Loop
    
    gstrSQL = "Select distinct 检查设备 From 影像检查记录"
    OpenRecordset rsTmp, Me.Caption
    Do While Not rsTmp.EOF
        cobCondition(2).AddItem Nvl(rsTmp(0))
        rsTmp.MoveNext
    Loop
    chkStorage.value = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "启动接收", 1)
    
    '填充WorkList的参数
    txtItem(4) = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "WorkList端口", 1024)
    txtDWLLocalAE = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "WorkList本机AE", "ZLPACSWL")
    chkDWL.value = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "启动WorkList", 0)
    txtItem(6) = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "WorkList检索天数", 3)
    chkModel = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "WorkList按设备过滤", 0))
    chkForceResult = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "WorkList使用强制结果", 0))
    subFillMsfResult    '填充Worklist返回值设置表
    
    '填充Query/Retrieve的参数
    txtItem(5) = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "Query/Retrieve端口", 1024)
    txtQueryAE.Text = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "Query/Retrieve本机AE", "ZLPACSQR")
    chkQuery.value = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "启动Query/Retrieve", 0)
    chkAcceptCGET.value = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "支持C-GET", 0)
    
    '填充“本地数据库”的参数
    strTemp = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "自动清空临时表", "")
    Dim strTempArray() As String
    strTempArray = Split(strTemp, ";")
    For i = 0 To 2
        chkClearTempTB(i).value = 0
    Next i
    For i = 0 To UBound(strTempArray)
        If strTempArray(i) = "影像接收序列" Then
            chkClearTempTB(0).value = 1
        ElseIf strTempArray(i) = "错误日志" Then
            chkClearTempTB(1).value = 1
        ElseIf strTempArray(i) = "DICOM通讯日志" Then
            chkClearTempTB(2).value = 1
        End If
    Next i
    txtClearInterval = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\接收服务", "自动清空临时表间隔天数", "0")
    If txtClearInterval <= 0 Then
        chkAutoClear.value = 0
        txtClearInterval.Enabled = False
    Else
        chkAutoClear.value = 1
        txtClearInterval.Enabled = True
    End If
    
    SetPrivs gstrPrivs
    
    SSTab1.Tab = 0
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub subFillMsfAutoRout()
    Dim lngRowPos As Long
    Dim i As Integer
    With MSFAutoRout
        .Clear
        .Rows = 1
        .Cols = 3
        .ColWidth(1) = 3000
        .TextMatrix(0, 0) = "条件类型"
        .TextMatrix(0, 1) = "条件内容"
        .TextMatrix(0, 2) = "目的设备"
        lngRowPos = 1
        For i = 1 To UBound(aAutoRoutSetting)
            .Rows = .Rows + 1
            .TextMatrix(lngRowPos, 0) = IIf(aAutoRoutSetting(i).Type = 1, "影像类别", "检查设备")
            .TextMatrix(lngRowPos, 1) = aAutoRoutSetting(i).strCondition
            .TextMatrix(lngRowPos, 2) = GetDeviceNameNum(aDevices, aAutoRoutSetting(i).strFTPDeviceNo, 0)
            lngRowPos = .Rows
        Next
    End With
End Sub
Private Function GetDeviceNameNum(aSource() As Variant, ByVal SeekString As String, iType As Integer) As String
    '获取设备的名称或设备号
    'iType=0---输入SeekString为设备号，返回设备名。
    'iType=1---输入SeekString为设备名，返回设备号。
    Dim i As Long
    For i = 0 To UBound(aSource, 2)
        If aSource(iType, i) = SeekString Then Exit For
    Next
    If i > UBound(aSource, 2) Then GetDeviceNameNum = "": Exit Function
    GetDeviceNameNum = IIf(iType = 1, aSource(0, i), aSource(1, i))
End Function
Private Function GetComboxIndex(aSource() As Variant, ByVal SeekString As String) As Long
    Dim i As Long
    
    For i = 0 To UBound(aSource, 2)
        If aSource(0, i) = SeekString Then Exit For
    Next
    If i > UBound(aSource, 2) Then i = 0
    GetComboxIndex = i
End Function


Private Sub MSFAutoRout_Click()
    Dim iSelected As Integer
    With MSFAutoRout
        iSelected = .RowSel
        Me.optType(IIf(.TextMatrix(iSelected, 0) = "影像类别", 1, 2)).value = True
        Me.cobCondition(IIf(.TextMatrix(iSelected, 0) = "影像类别", 1, 2)).Text = .TextMatrix(iSelected, 1)
        Me.cobDestination = .TextMatrix(iSelected, 2)
    End With
End Sub

Private Sub MSFResult_Click()
    Dim iSelected As Integer
    With MSFResult
        iSelected = .RowSel
        Me.chkUseResult.value = IIf(.TextMatrix(iSelected, 3) = "True", 1, 0)
        Me.lblResult.Caption = .TextMatrix(iSelected, 0) & " " & .TextMatrix(iSelected, 1) & " : " & .TextMatrix(iSelected, 2)
        Me.txtResult(0).Text = .TextMatrix(iSelected, 4)
        Me.txtResult(1).Text = .TextMatrix(iSelected, 5)
        Me.chkResult.value = IIf(.TextMatrix(iSelected, 6) = True, 1, 0)
    End With
End Sub

Private Sub optImgMatch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub optMatch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub optType_Click(Index As Integer)
    Me.cobCondition(Index).Enabled = True
    Me.cobCondition(IIf(Index = 1, 2, 1)).Enabled = False
End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
    On Error Resume Next
    Select Case SSTab1.Tab
        Case 0
            chkStorage.SetFocus
        Case 1
            chkDWL.SetFocus
        Case 2
            chkQuery.SetFocus
        Case 3
            chkClearTempTB(0).SetFocus
    End Select
End Sub

Private Sub txtClearInterval_GotFocus()
    With txtClearInterval
        .SelStart = 0: .SelLength = .MaxLength
    End With
End Sub

Private Sub txtClearInterval_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtClearInterval_KeyPress(KeyAscii As Integer)
    If ifEditKey(KeyAscii, False) Then Exit Sub
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then KeyAscii = 0
End Sub

Private Sub txtDWLLocalAE_GotFocus()
    txtDWLLocalAE.SelStart = 0
    txtDWLLocalAE.SelLength = Len(txtDWLLocalAE.Text)
End Sub

Private Sub txtDWLLocalAE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtItem_GotFocus(Index As Integer)
    With Me.txtItem(Index)
        .SelStart = 0: .SelLength = .MaxLength
    End With
End Sub

Private Sub txtItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtItem_KeyPress(Index As Integer, KeyAscii As Integer)
    If ifEditKey(KeyAscii, False) Then Exit Sub
    
    If LenB(StrConv(Trim(txtItem(Index).Text), vbFromUnicode)) >= txtItem(Index).MaxLength Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case Index
        Case 0, 2, 3, 6
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then KeyAscii = 0
    End Select
End Sub

Private Sub txtItem_LostFocus(Index As Integer)
    Select Case Index
        Case -1
            Call zlCommFun.OpenIme(False)
    End Select
End Sub

'判断是否为编辑键
Private Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or _
      KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function

Private Sub txtQueryAE_GotFocus()
    txtQueryAE.SelStart = 0
    txtQueryAE.SelLength = Len(txtQueryAE.Text)
End Sub

Private Sub txtQueryAE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
Private Sub SetPrivs(strPrivs As String)
    '---------------------------------------------------------------
    '功能：                                  设置人员使用权限
    '参数：
    '返回：                                  无
    '上级函数或过程：                        frmParaSet.Form_load
    '下级函数或过程：                        无
    '引用的外部参数：                        mstrPrivs
    '编制人：                                曾超 2005-8-25
    '---------------------------------------------------------------
    If InStr(strPrivs, "存储自动路由") = 0 Then
        cmdInsert.Enabled = False
        cmdModify.Enabled = False
        cmdDelete.Enabled = False
    End If
    If InStr(strPrivs, "DICOM工作列表服务") = 0 Then
        chkDWL.Enabled = False
        frmWorkList.Enabled = False
    End If
    If InStr(strPrivs, "DICOM检索服务") = 0 Then
        chkQuery.Enabled = False
        frmQueryRetrieve.Enabled = False
    End If
End Sub

Private Sub subFillMsfResult()
    Dim lngRowPos As Long
    Dim i As Integer
    Dim rsTmp As New ADODB.Recordset
    
    With MSFResult
        .Clear
        .Rows = 1
        .Cols = 7
        .ColWidth(0) = 800
        .ColWidth(1) = 1800
        .ColWidth(2) = 1800
        .ColWidth(3) = 600
        .ColWidth(4) = 1300
        .ColWidth(5) = 1300
        .ColWidth(6) = 600
        .FixedCols = 3
        .TextMatrix(0, 0) = "标记"
        .TextMatrix(0, 1) = "中文标题"
        .TextMatrix(0, 2) = "英文标题"
        .TextMatrix(0, 3) = "被选择"
        .TextMatrix(0, 4) = "数据值"
        .TextMatrix(0, 5) = "强制结果值"
        .TextMatrix(0, 6) = "递增"
        lngRowPos = 1
        If gcnAccess.State = adStateClosed Then Exit Sub
        
        Set rsTmp = gcnAccess.Execute("select * from 强制结果")
        While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(lngRowPos, 0) = "(" & rsTmp!组号 & "," & rsTmp!元素号 & ")"
            .TextMatrix(lngRowPos, 1) = Nvl(rsTmp!中文标题)
            .TextMatrix(lngRowPos, 2) = Nvl(rsTmp!英文标题)
            .TextMatrix(lngRowPos, 3) = rsTmp!被选择
            .TextMatrix(lngRowPos, 4) = Nvl(rsTmp!数据值)
            .TextMatrix(lngRowPos, 5) = Nvl(rsTmp!强制结果值)
            .TextMatrix(lngRowPos, 6) = rsTmp!是否递增
            rsTmp.MoveNext
            lngRowPos = .Rows
        Wend
    End With
End Sub

Private Sub subChangeMSFResult()
    Dim iSelect As Integer
    With Me.MSFResult
        iSelect = .RowSel
        If frmBuildResult.funVerifyResult(Me.txtResult(0).Text) <> 0 _
        Or frmBuildResult.funVerifyResult(Me.txtResult(1).Text) <> 0 Then
            Exit Sub
        End If
       .TextMatrix(iSelect, 3) = IIf(Me.chkUseResult.value = 0, "False", "True")
        .TextMatrix(iSelect, 4) = Me.txtResult(0).Text
        .TextMatrix(iSelect, 5) = Me.txtResult(1).Text
        .TextMatrix(iSelect, 6) = IIf(Me.chkResult.value = 0, "False", "True")
    End With
End Sub

Private Sub txtResult_Change(Index As Integer)
    mblnModifyMWLResult = True
    If Index = 1 Then
        Me.MSFResult.TextMatrix(Me.MSFResult.RowSel, 5) = Me.txtResult(1).Text
    End If
End Sub

Private Sub txtResult_Click(Index As Integer)
    If Index = 0 Then
        cmdBuildResult_Click (0)
    End If
End Sub

Private Sub txtResult_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        KeyAscii = 0
        cmdBuildResult_Click (0)
    End If
End Sub
