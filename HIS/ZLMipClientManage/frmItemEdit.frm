VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmItemEdit 
   BackColor       =   &H8000000D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15645
   Icon            =   "frmItemEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   15645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8415
      Index           =   2
      Left            =   5280
      ScaleHeight     =   8385
      ScaleWidth      =   9525
      TabIndex        =   41
      Top             =   360
      Width           =   9555
      Begin VB.OptionButton opt 
         Caption         =   "定时检查，每次轮询服务的服务周期在如下时间段内"
         Height          =   180
         Index           =   7
         Left            =   45
         TabIndex        =   23
         Top             =   360
         Width           =   5070
      End
      Begin VB.Frame fra 
         Height          =   7065
         Index           =   4
         Left            =   15
         TabIndex        =   42
         Top             =   525
         Width           =   9480
         Begin VB.CommandButton cmdClear 
            Caption         =   "清除(&D)"
            Enabled         =   0   'False
            Height          =   350
            Left            =   8310
            TabIndex        =   29
            Top             =   150
            Width           =   1100
         End
         Begin VB.CommandButton cmdSet 
            Caption         =   "标记(&F)"
            Enabled         =   0   'False
            Height          =   350
            Left            =   7155
            TabIndex        =   28
            Top             =   150
            Width           =   1100
         End
         Begin VB.ComboBox cbo 
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   930
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   165
            Width           =   555
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   300
            Index           =   5
            Left            =   270
            TabIndex        =   24
            Text            =   "1"
            Top             =   165
            Width           =   360
         End
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   6465
            Index           =   0
            Left            =   30
            TabIndex        =   27
            Top             =   525
            Width           =   9375
            _cx             =   16536
            _cy             =   11404
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   0   'False
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
            GridColor       =   14737632
            GridColorFixed  =   12632256
            TreeColor       =   -2147483638
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   31
            Cols            =   51
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   300
            ColWidthMin     =   180
            ColWidthMax     =   180
            ExtendLastCol   =   0   'False
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
            OutlineBar      =   4
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   1
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
         Begin MSComCtl2.UpDown upd 
            Height          =   300
            Index           =   5
            Left            =   645
            TabIndex        =   25
            Top             =   150
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txt(5)"
            BuddyDispid     =   196615
            BuddyIndex      =   5
            OrigLeft        =   3660
            OrigTop         =   1035
            OrigRight       =   3915
            OrigBottom      =   1335
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   0   'False
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "每              检查"
            Height          =   180
            Index           =   11
            Left            =   60
            TabIndex        =   43
            Top             =   225
            Width           =   1800
         End
      End
      Begin VB.OptionButton opt 
         Caption         =   "随时检查，每次轮询服务的服务周期内都会检查"
         Height          =   180
         Index           =   6
         Left            =   45
         TabIndex        =   22
         Top             =   75
         Value           =   -1  'True
         Width           =   4545
      End
   End
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8415
      Index           =   1
      Left            =   1005
      ScaleHeight     =   8385
      ScaleWidth      =   9510
      TabIndex        =   40
      Top             =   225
      Width           =   9540
      Begin VB.Frame fra 
         Height          =   1245
         Index           =   1
         Left            =   1215
         TabIndex        =   45
         Top             =   2865
         Width           =   7800
         Begin VB.OptionButton opt 
            Caption         =   "一直有效"
            Height          =   180
            Index           =   0
            Left            =   105
            TabIndex        =   16
            Top             =   300
            Value           =   -1  'True
            Width           =   2700
         End
         Begin VB.OptionButton opt 
            Caption         =   "指定时间"
            Height          =   180
            Index           =   1
            Left            =   105
            TabIndex        =   17
            Top             =   780
            Width           =   1050
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   0
            Left            =   1605
            TabIndex        =   18
            Top             =   735
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   130678787
            CurrentDate     =   41634
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   1
            Left            =   3180
            TabIndex        =   19
            Top             =   735
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   130678787
            CurrentDate     =   41634
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "从               到"
            Height          =   180
            Index           =   6
            Left            =   1410
            TabIndex        =   46
            Top             =   780
            Visible         =   0   'False
            Width           =   1710
         End
      End
      Begin VB.TextBox txt 
         Height          =   4080
         Index           =   4
         Left            =   1215
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   4200
         Width           =   7800
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   1215
         TabIndex        =   5
         Top             =   525
         Width           =   7800
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   5625
         TabIndex        =   3
         Top             =   105
         Width           =   3405
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1215
         TabIndex        =   1
         Top             =   105
         Width           =   3300
      End
      Begin VB.Frame fra 
         Height          =   2010
         Index           =   2
         Left            =   1215
         TabIndex        =   44
         Top             =   840
         Width           =   7800
         Begin MSComCtl2.UpDown upd 
            Height          =   315
            Index           =   6
            Left            =   3106
            TabIndex        =   11
            Top             =   1050
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txt(6)"
            BuddyDispid     =   196615
            BuddyIndex      =   6
            OrigLeft        =   3660
            OrigTop         =   1035
            OrigRight       =   3915
            OrigBottom      =   1335
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   0   'False
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   315
            Index           =   7
            Left            =   1515
            TabIndex        =   13
            Text            =   "24"
            Top             =   1500
            Width           =   465
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   315
            Index           =   6
            Left            =   2775
            TabIndex        =   10
            Text            =   "5"
            Top             =   1050
            Width           =   330
         End
         Begin VB.OptionButton opt 
            Caption         =   "无须重发，即首次发送失败后不再重发"
            Height          =   180
            Index           =   2
            Left            =   90
            TabIndex        =   7
            Top             =   270
            Value           =   -1  'True
            Width           =   6855
         End
         Begin VB.OptionButton opt 
            Caption         =   "限次重发，首次发送后可重发        次"
            Height          =   180
            Index           =   4
            Left            =   90
            TabIndex        =   9
            Top             =   1095
            Width           =   4605
         End
         Begin VB.OptionButton opt 
            Caption         =   "一直重发，每次轮询服务的服务周期内都会重发"
            Height          =   180
            Index           =   3
            Left            =   90
            TabIndex        =   8
            Top             =   690
            Width           =   6375
         End
         Begin MSComCtl2.UpDown upd 
            Height          =   315
            Index           =   7
            Left            =   1981
            TabIndex        =   14
            Top             =   1500
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txt(7)"
            BuddyDispid     =   196615
            BuddyIndex      =   7
            OrigLeft        =   2550
            OrigTop         =   1425
            OrigRight       =   2805
            OrigBottom      =   1725
            Max             =   96
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   0   'False
         End
         Begin VB.OptionButton opt 
            Caption         =   "限时重发，在         小时内一直重发"
            Height          =   180
            Index           =   5
            Left            =   90
            TabIndex        =   12
            Top             =   1560
            Width           =   5385
         End
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "备注说明(&N)"
         Height          =   180
         Index           =   4
         Left            =   105
         TabIndex        =   20
         Top             =   4200
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "项目名称(&T)"
         Height          =   180
         Index           =   3
         Left            =   90
         TabIndex        =   4
         Top             =   585
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "流程标识(&I)"
         Height          =   180
         Index           =   2
         Left            =   4575
         TabIndex        =   2
         Top             =   165
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "项目标识(&B)"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   0
         Top             =   165
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "重发策略(&A)"
         Height          =   180
         Index           =   5
         Left            =   90
         TabIndex        =   6
         Top             =   930
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "有效时间(&V)"
         Height          =   180
         Index           =   7
         Left            =   75
         TabIndex        =   15
         Top             =   2970
         Width           =   990
      End
   End
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8415
      Index           =   3
      Left            =   2370
      ScaleHeight     =   8385
      ScaleWidth      =   9510
      TabIndex        =   47
      Top             =   135
      Width           =   9540
      Begin VB.Frame fra 
         Caption         =   "产生消息"
         Height          =   1155
         Index           =   3
         Left            =   45
         TabIndex        =   48
         Top             =   7185
         Width           =   9405
         Begin VB.OptionButton opt 
            Caption         =   "每天产生1次"
            Height          =   180
            Index           =   9
            Left            =   2145
            TabIndex        =   33
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton opt 
            Caption         =   "每次产生"
            Height          =   180
            Index           =   8
            Left            =   825
            TabIndex        =   32
            Top             =   720
            Value           =   -1  'True
            Width           =   1050
         End
         Begin VB.OptionButton opt 
            Caption         =   "每个检查周期内产生1次"
            Height          =   180
            Index           =   10
            Left            =   3795
            TabIndex        =   34
            Top             =   720
            Width           =   2235
         End
         Begin VB.Label lbl 
            Caption         =   "在检查频率周期内并且触发条件满足后的产生消息的频率"
            Height          =   270
            Index           =   9
            Left            =   855
            TabIndex        =   51
            Top             =   345
            Width           =   8310
         End
         Begin VB.Image img 
            Height          =   480
            Index           =   0
            Left            =   90
            Picture         =   "frmItemEdit.frx":000C
            Top             =   315
            Width           =   480
         End
      End
      Begin VB.Frame fra 
         Caption         =   "触发条件"
         Height          =   7065
         Index           =   0
         Left            =   30
         TabIndex        =   49
         Top             =   60
         Width           =   9435
         Begin VB.CommandButton cmdVerfiy 
            Caption         =   "校验(&V)"
            Height          =   350
            Left            =   8220
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   315
            Width           =   1100
         End
         Begin RichTextLib.RichTextBox rtbSQL 
            Height          =   6150
            Left            =   75
            TabIndex        =   30
            Top             =   840
            Width           =   9270
            _ExtentX        =   16351
            _ExtentY        =   10848
            _Version        =   393217
            BorderStyle     =   0
            ScrollBars      =   1
            TextRTF         =   $"frmItemEdit.frx":198E
         End
         Begin VB.Label lbl 
            Caption         =   "输入有效的SQL，如果SQL执行有结果或SQL为空，则表示触发条件成立，否则不成立。"
            Height          =   270
            Index           =   8
            Left            =   630
            TabIndex        =   50
            Top             =   405
            Width           =   6765
         End
         Begin VB.Image img 
            Height          =   480
            Index           =   1
            Left            =   75
            Picture         =   "frmItemEdit.frx":1A2B
            Top             =   240
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3000
      Index           =   0
      Left            =   165
      ScaleHeight     =   3000
      ScaleWidth      =   4170
      TabIndex        =   37
      Top             =   6120
      Width           =   4170
      Begin XtremeSuiteControls.TabControl tbcPage 
         Height          =   1920
         Left            =   195
         TabIndex        =   39
         Top             =   255
         Width           =   2505
         _Version        =   589884
         _ExtentX        =   4419
         _ExtentY        =   3387
         _StockProps     =   64
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   540
         TabIndex        =   35
         Top             =   2370
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(C)"
         Height          =   350
         Left            =   1665
         TabIndex        =   36
         Top             =   2385
         Width           =   1100
      End
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1470
      TabIndex        =   38
      Top             =   75
      Width           =   1575
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private mfrmParent As Object
Private mbytMode As Byte
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mrsPara As ADODB.Recordset
Private mstrDataKey As String
Private mlngModualCode As Long
Private mblnContiune As Boolean
'Private mstrClassKey As String
Private mclsVsf As zlVSFlexGrid.clsVsf
Private mstrDataCode As String
Private mblnStartUp As Boolean

Public Event AfterNewData(ByVal DataKey As String)
Public Event AfterModifyData(ByVal DataKey As String)
Public Event AfterDeleteData(ByVal DataKey As String)
Public Event Forward(ByRef DataKey As String, ByRef Cancel As Boolean)
Public Event Backward(ByRef DataKey As String, ByRef Cancel As Boolean)

'######################################################################################################################

Public Function InitDialog(ByVal frmParent As Object, ByVal lngModualCode As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Set mfrmParent = frmParent
    mlngModualCode = lngModualCode
    InitDialog = True
    
End Function

Public Sub NewData(ByVal strDataCode As String)
    '******************************************************************************************************************
    '功能：新增消息项目
    '参数：
    '返回：
    '******************************************************************************************************************
    mbytMode = 1
    Me.Caption = "新增项目"
    mstrDataKey = ""
'    mstrClassKey = strClassKey
    mstrDataCode = strDataCode
    
    Call InitData
    Call InitTabControl
    Call InitCommandBar
    
'    Call ReadClassData(mstrClassKey)
    
    txt(1).Text = "ZLHIS_USER_"
'    DoEvents
'    txt(1).SetFocus
    
    mblnDataChanged = False
        
    Me.Show 1, mfrmParent
    
End Sub

Public Sub ModifyData(ByVal strDataCode As String, ByVal strDataKey As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mbytMode = 2
    mstrDataKey = strDataKey
    mstrDataCode = strDataCode
    Me.Caption = "修改项目"
    
    Call InitData
    Call InitTabControl
    Call InitCommandBar
    
    Call ReadData(mstrDataKey)
'    DoEvents
'    txt(1).SetFocus
    
    Me.Show 1, mfrmParent
    
End Sub

Public Sub DeleteData(ByVal strDataCode As String, ByVal strDataKey As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mbytMode = 3
    If strDataKey = "" Then Exit Sub
    mstrDataKey = strDataKey
    mstrDataCode = strDataCode
    Set mrsPara = zlCommFun.CreateParameter
    Call zlCommFun.SetParameter(mrsPara, "ID", mstrDataKey)

    If gclsBusiness.ItemEdit("Delete", mrsPara) Then
        RaiseEvent AfterDeleteData(mstrDataKey)
    End If
End Sub

'######################################################################################################################

Private Function InitTabControl() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    With tbcPage
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
'            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
        End With

        Set .Icons = zlCommFun.GetPubIcons
        
        .InsertItem 0, "基本资料", picPane(1).hWnd, 0
        .InsertItem 1, "检查频率", picPane(2).hWnd, 0
        .InsertItem 2, "触发产生", picPane(3).hWnd, 0
                                
        .Item(0).Selected = True
        
    End With
    
    InitTabControl = True
    
End Function

Private Function InitGrid(ByVal strType As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim intStartRow As Integer
                
    With vsf(0)
        mclsVsf.ClearGrid
        .MergeRow(1) = False
        .OutlineBar = flexOutlineBarNone
        .Cell(flexcpBackColor, 1, 1, 1, .Cols - 1) = .BackColor
        .Cell(flexcpData, 1, 0) = 0
        Select Case strType
        '--------------------------------------------------------------------------------------------------------------
        Case "天"
            .Rows = 2
            .TextMatrix(1, 0) = "1"
        '--------------------------------------------------------------------------------------------------------------
        Case "周"
            .Rows = 8
            For intRow = 1 To 7
                .TextMatrix(intRow, 0) = _
                        Switch(intRow = 1, "星期一", intRow = 2, "星期二", intRow = 3, "星期三", intRow = 4, "星期四", intRow = 5, "星期五", intRow = 6, "星期六", intRow = 7, "星期日")
            Next
        '--------------------------------------------------------------------------------------------------------------
        Case "月"
            .Rows = 32
            For intRow = 1 To 31
                .TextMatrix(intRow, 0) = intRow
            Next
        '--------------------------------------------------------------------------------------------------------------
        Case "年"
            .Rows = 12 + 366 + 1
            intStartRow = 1
            .OutlineBar = flexOutlineBarCompleteLeaf
            For intLoop = 1 To 12
                
                .MergeRow(intStartRow) = True
                .TextMatrix(intStartRow, 0) = Format(intLoop, "00") & "月"
                .Cell(flexcpData, intStartRow, 0) = intLoop
                .TextMatrix(intStartRow, 49) = intLoop
                .TextMatrix(intStartRow, 50) = ""
                intStartRow = intStartRow + 1
                
                Select Case intLoop
                Case 1, 3, 5, 7, 8, 10, 12
                    For intRow = intStartRow To intStartRow + 30
                        .TextMatrix(intRow, 0) = Format(intRow - intStartRow + 1, "00")
                        .TextMatrix(intRow, 49) = intLoop & "-" & (intRow - intStartRow + 1)
                        .TextMatrix(intRow, 50) = intLoop
                    Next
                    intStartRow = intStartRow + 31
                Case 2
                    For intRow = intStartRow To intStartRow + 28
                        .TextMatrix(intRow, 0) = Format(intRow - intStartRow + 1, "00")
                        .TextMatrix(intRow, 49) = intLoop & "-" & (intRow - intStartRow + 1)
                        .TextMatrix(intRow, 50) = intLoop
                    Next
                    intStartRow = intStartRow + 29
                Case 4, 6, 9, 11
                    For intRow = intStartRow To intStartRow + 29
                        .TextMatrix(intRow, 0) = Format(intRow - intStartRow + 1, "00")
                        .TextMatrix(intRow, 49) = intLoop & "-" & (intRow - intStartRow + 1)
                        .TextMatrix(intRow, 50) = intLoop
                    Next
                    intStartRow = intStartRow + 30
                End Select
            Next
            
            Call mclsVsf.ShowOutline(49, 50, .BackColorFixed)
        End Select
        
    End With
End Function

Private Function InitData() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim intLoop As Integer
    Dim intCount As Integer
    
    mblnContiune = False
        
    picPane(1).BorderStyle = 0
    picPane(2).BorderStyle = 0
    picPane(3).BorderStyle = 0
    
    Set rsTmp = gclsBusiness.ItemStruct()
    If Not (rsTmp Is Nothing) Then
        txt(1).MaxLength = rsTmp("item_code").Precision
        txt(2).MaxLength = rsTmp("item_flow").DefinedSize
        txt(3).MaxLength = rsTmp("item_title").DefinedSize
        txt(4).MaxLength = rsTmp("item_note").DefinedSize
    End If
        
    '------------------------------------------------------------------------------------------------------------------
    Set mclsVsf = New zlVSFlexGrid.clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, vsf(0), False, False, GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("", 690, flexAlignCenterCenter, flexDTString, , "", True)
                
        For intLoop = 1 To 24
            Call .AppendColumn(intLoop, 165, flexAlignCenterCenter, flexDTString, , "", True)
            Call .AppendColumn(intLoop, 165, flexAlignCenterCenter, flexDTString, , "", True)
        Next
        Call .AppendColumn("id", 0, flexAlignCenterCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("parent_id", 0, flexAlignCenterCenter, flexDTString, , "", True, , , True)
        
'        .VsfObject.OutlineCol = 0
                
    End With
    '------------------------------------------------------------------------------------------------------------------
    
    With vsf(0)
        .FixedCols = 1
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .ColWidthMax = 0
        .ColWidthMin = 0
    End With
    
    With cbo(0)
        .Clear
        .AddItem "天"
        .ItemData(.NewIndex) = 1
        .AddItem "周"
        .ItemData(.NewIndex) = 2
        .AddItem "月"
        .ItemData(.NewIndex) = 3
        .AddItem "年"
        .ItemData(.NewIndex) = 4
        .ListIndex = 0
    End With
    
    dtp(0).Value = Format(zlDataBase.Currentdate, "yyyy-MM-dd")
    dtp(1).Value = Format(zlDataBase.Currentdate, "yyyy-MM-dd")
    
    InitData = True
End Function


Private Function ReadData(ByVal strDataKey As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intRow As Integer
    Dim intDay As Integer
    Dim intMonth As Integer
    Dim intStartCol As Integer
    Dim intEndCol As Integer
    Dim rsTmp As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    On Error GoTo errHand
    
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "id", strDataKey)
    
    mblnReading = True
    Set rsTmp = gclsBusiness.ItemRead("id", rsCondition)
    If rsTmp.BOF = False Then
'        txt(0).Text = AppendCode(zlCommFun.NVL(rsTmp("folder_name").Value), zlCommFun.NVL(rsTmp("folder_code").Value))
'        cmd(0).Tag = zlCommFun.NVL(rsTmp("folder_id").Value)
        txt(1).Text = zlCommFun.NVL(rsTmp("item_code").Value)
        txt(2).Text = zlCommFun.NVL(rsTmp("item_flow").Value)
        txt(3).Text = zlCommFun.NVL(rsTmp("item_title").Value)
        txt(4).Text = zlCommFun.NVL(rsTmp("item_note").Value)
        rtbSQL.Text = zlCommFun.NVL(rsTmp("trigger_condition").Value)
        Select Case zlCommFun.NVL(rsTmp("again_policy").Value, 0)
        Case 0
            opt(2).Value = True
        Case 1
            opt(3).Value = True
        Case 2
            opt(4).Value = True
            txt(6).Text = zlCommFun.NVL(rsTmp("again_para").Value, 5)
        Case 3
            opt(5).Value = True
            txt(7).Text = zlCommFun.NVL(rsTmp("again_para").Value, 24)
        End Select
        
        Select Case zlCommFun.NVL(rsTmp("trigger_frequency").Value, 0)
        Case 0
            opt(8).Value = True
        Case 1
            opt(9).Value = True
        Case 2
            opt(10).Value = True
        End Select
        
        If Format(zlCommFun.NVL(rsTmp("start_date").Value), "yyyy-MM-dd") = "2000-01-01" And Format(zlCommFun.NVL(rsTmp("stop_date").Value), "yyyy-MM-dd") = "3000-01-01" Then
            opt(0).Value = True
        Else
            opt(1).Value = True
            dtp(0).Value = Format(zlCommFun.NVL(rsTmp("start_date").Value), dtp(0).CustomFormat)
            dtp(1).Value = Format(zlCommFun.NVL(rsTmp("stop_date").Value, "3000-01-01"), dtp(1).CustomFormat)
            dtp(0).Visible = True
            dtp(1).Visible = True
            lbl(6).Visible = True
        End If
        
        If zlCommFun.NVL(rsTmp("check_frequency").Value, 0) = 0 Then
            opt(6).Value = True
        Else
            opt(7).Value = True
            Call zlControl.CboLocate(cbo(0), zlCommFun.NVL(rsTmp("check_frequency").Value, 1), True)
            txt(5).Text = zlCommFun.NVL(rsTmp("check_freq_internal").Value, 1)
            
            Call zlCommFun.SetCondition(rsCondition, "item_id", strDataKey)
            Set rsTmp = gclsBusiness.ItemFrequencyRead("item_id", rsCondition)
            If rsTmp.BOF = False Then
                With vsf(0)
                    
                    Do While Not rsTmp.EOF
                        
                        intMonth = zlCommFun.NVL(rsTmp("freq_month").Value, 0)
                        intDay = Val(rsTmp("freq_day").Value)
                        intStartCol = GetColumn(rsTmp("freq_start").Value, 1)
                        intEndCol = GetColumn(rsTmp("freq_stop").Value, 2)
                        If intMonth = 0 Then
                            .Cell(flexcpData, intDay, intStartCol, intDay, intEndCol) = 1
                        Else
                            .Cell(flexcpData, intDay + GetDays(intMonth - 1) + intMonth, intStartCol, intDay + GetDays(intMonth - 1) + intMonth, intEndCol) = 1
                        End If
                        
                        rsTmp.MoveNext
                    Loop
                    
                End With
            End If
        End If
        
    End If
    mblnReading = False
    mblnDataChanged = False
    
    ReadData = True
    
    Exit Function
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim objFindKey As CommandBarControl
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call zlCommFun.CommandBarInit(cbsMain)
    cbsMain.VisualTheme = xtpThemeWhidbey
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份
    
    
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
        

    
    mstrFindKey = zlDataBase.GetPara("定位依据", ParamInfo.系统号, mlngModualCode, "名称")
    If mstrFindKey = "" Then mstrFindKey = "名称"

    Set mobjFindKey = zlCommFun.NewToolBar(objBar, xtpControlPopup, conMenu_View_LocationItem, mstrFindKey, True, , xtpButtonIconAndCaption)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.flags = xtpFlagRightAlign
    mobjFindKey.Style = xtpButtonIconAndCaption
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&1.名称"): objControl.Parameter = "名称"
    objControl.IconId = 1
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.编码"): objControl.Parameter = "编码"
    objControl.IconId = 1

    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, 0, "")
    cbrCustom.Handle = txtLocation.hWnd
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Filter, "搜索")
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Forward, "上一条")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Backward, "下一条")
        
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Option, IIf(mbytMode = 1, "确定之继续新增", "确定之继续修改"), False)
    objControl.IconId = conMenu_View_UnCheck
    If mbytMode <> 1 Then objControl.flags = xtpFlagRightAlign
    
    txtLocation.Visible = (mbytMode = 2)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function ValidData() As Boolean
    '******************************************************************************************************************
    '功能：校验编辑数据的有效性
    '参数：
    '返回：
    '******************************************************************************************************************
        
'
'    If Len(txt(0).Text) = 0 Then
'        ShowSimpleMsg "业务事件的名称不能为空！"
'        Call LocationObj(txt(0))
'        Exit Function
'    End If
    
    If Len(txt(1).Text) = 0 Then
        ShowSimpleMsg "消息项目标识能为空！"
        Call LocationObj(txt(1))
        Exit Function
    End If
'
'    '检查编码是否为数字字符
'    If zlCommFun.CheckStrType(Trim(txt(1).Text), 99, "0123456789") = False Then
'        ShowSimpleMsg "业务事件的编码必须为数字字符！"
'        LocationObj txt(1)
'        Exit Function
'    End If
    
    ValidData = True
    
End Function

Private Function SaveData(ByRef strDataKey As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsPara As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim strFreqContent As String
    Dim intRow As Integer
    Dim intCol As Integer
    Dim intMonth As Integer
    Dim intStartPoint As Integer
    Dim intEndPoint As Integer
    
    On Error GoTo errHand
    
    Set rsPara = zlCommFun.CreateParameter
    
    Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
'    Call zlCommFun.SetParameter(rsPara, "folder_id", Trim(cmd(0).Tag))
    Call zlCommFun.SetParameter(rsPara, "data_code", mstrDataCode)
    Call zlCommFun.SetParameter(rsPara, "item_code", Trim(txt(1).Text))
    Call zlCommFun.SetParameter(rsPara, "item_flow", Trim(txt(2).Text))
    Call zlCommFun.SetParameter(rsPara, "item_title", Trim(txt(3).Text))
    If opt(6).Value Then
        Call zlCommFun.SetParameter(rsPara, "check_frequency", 0)
        Call zlCommFun.SetParameter(rsPara, "check_freq_internal", 0)
        Call zlCommFun.SetParameter(rsPara, "freq_content", "")
    Else
        Call zlCommFun.SetParameter(rsPara, "check_frequency", cbo(0).ItemData(cbo(0).ListIndex))
        Call zlCommFun.SetParameter(rsPara, "check_freq_internal", Val(txt(5).Text))
        
        With vsf(0)
            strFreqContent = ""
                        
            For intRow = 1 To .Rows - 1
                intStartPoint = 0
                intEndPoint = 0
                
                If cbo(0).Text = "年" Then
                    If Val(.Cell(flexcpData, intRow, 0)) > 0 Then
                        intMonth = Val(.Cell(flexcpData, intRow, 0))
                    End If
                End If
                
                For intCol = 1 To 49
                    If Val(.Cell(flexcpData, intRow, intCol)) = 1 Then
                        If intStartPoint = 0 Then intStartPoint = intCol
                    ElseIf Val(.Cell(flexcpData, intRow, intCol)) = 0 Then
                        If intStartPoint > 0 Then
                            intEndPoint = intCol - 1
                            If cbo(0).Text = "年" Then
                                strFreqContent = strFreqContent & ";" & intMonth & "," & intRow - GetDays(intMonth - 1) - intMonth & "," & GetTime(intStartPoint, 1) & "," & GetTime(intEndPoint, 2)
                            Else
                                strFreqContent = strFreqContent & ";0," & intRow & "," & GetTime(intStartPoint, 1) & "," & GetTime(intEndPoint, 2)
                            End If
                            
                        End If
                        intStartPoint = 0
                        intEndPoint = 0
                    End If
                Next
            Next

        End With
        If strFreqContent <> "" Then strFreqContent = Mid(strFreqContent, 2)
        Call zlCommFun.SetParameter(rsPara, "freq_content", strFreqContent)
    End If
    
    If opt(2).Value Then
        Call zlCommFun.SetParameter(rsPara, "again_policy", 0)
        Call zlCommFun.SetParameter(rsPara, "again_para", "")
    ElseIf opt(3).Value Then
        Call zlCommFun.SetParameter(rsPara, "again_policy", 1)
        Call zlCommFun.SetParameter(rsPara, "again_para", "")
    ElseIf opt(4).Value Then
        Call zlCommFun.SetParameter(rsPara, "again_policy", 2)
        Call zlCommFun.SetParameter(rsPara, "again_para", Val(txt(6).Text))
    ElseIf opt(5).Value Then
        Call zlCommFun.SetParameter(rsPara, "again_policy", 3)
        Call zlCommFun.SetParameter(rsPara, "again_para", Val(txt(7).Text))
    End If
    
    If opt(0).Value Then
        Call zlCommFun.SetParameter(rsPara, "start_date", "2000-01-01")
        Call zlCommFun.SetParameter(rsPara, "stop_date", "3000-01-01")
    Else
        Call zlCommFun.SetParameter(rsPara, "start_date", Format(dtp(0).Value, dtp(0).CustomFormat) & " 00:00:00")
        Call zlCommFun.SetParameter(rsPara, "stop_date", Format(dtp(1).Value, dtp(1).CustomFormat) & " 23:59:59")
    End If
        
    Call zlCommFun.SetParameter(rsPara, "trigger_condition", Replace(Trim(rtbSQL.Text), "'", "''"))
    If opt(8).Value Then
        Call zlCommFun.SetParameter(rsPara, "trigger_frequency", 0)
    ElseIf opt(9).Value Then
        Call zlCommFun.SetParameter(rsPara, "trigger_frequency", 1)
    ElseIf opt(10).Value Then
        Call zlCommFun.SetParameter(rsPara, "trigger_frequency", 2)
    End If
    Call zlCommFun.SetParameter(rsPara, "item_note", Trim(txt(4).Text))
            
    '------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    Dim strLine As String
    Dim lngCount As Long
    
    If Trim(rtbSQL.Text) <> "" Then
        strTemp = ""
        strLine = ""
        lngCount = 0
        
        Set rsTmp = gclsBusiness.GetSQLField(Trim(rtbSQL.Text))
        If Not (rsTmp Is Nothing) Then
            If rsTmp.BOF = False Then
                rsTmp.MoveFirst
                Do While Not rsTmp.EOF
                    strLine = rsTmp("序号").Value
                    strLine = strLine & "," & rsTmp("名称").Value
                    strLine = strLine & "," & rsTmp("类型").Value
                    
                    If LenB(strTemp & ";" & strLine) > 3500 Then
                        If strTemp <> "" Then
                            lngCount = lngCount + 1
                            strTemp = Mid(strTemp, 2)
                            Call zlCommFun.SetParameter(rsPara, "SQL字段_" & lngCount, strTemp)
                            strTemp = ""
                        End If
                    End If
                    strTemp = strTemp & ";" & strLine
                
                    rsTmp.MoveNext
                Loop
            End If
        End If
        If strTemp <> "" Then
            lngCount = lngCount + 1
            strTemp = Mid(strTemp, 2)
            Call zlCommFun.SetParameter(rsPara, "SQL字段_" & lngCount, strTemp)
        End If
        Call zlCommFun.SetParameter(rsPara, "SQL字段个数", lngCount)
    
    End If
    
    Select Case mbytMode
    '------------------------------------------------------------------------------------------------------------------
    Case 1          '新增
        strDataKey = zlCommFun.GetGUID
        Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
        
        SaveData = gclsBusiness.ItemEdit("INSERT", rsPara)
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '修改
        SaveData = gclsBusiness.ItemEdit("UPDATE", rsPara)
    End Select
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbo_Change(Index As Integer)
    mblnDataChanged = True
End Sub

Private Sub cbo_Click(Index As Integer)
    mblnDataChanged = True
    
    Select Case Index
    Case 0
        Call InitGrid(cbo(Index).Text)
    End Select
    
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
        Select Case Index
        Case 0
            tbcPage.Item(2).Selected = True
            rtbSQL.SetFocus
        Case Else
            zlCommFun.PressKey vbKeyTab
        End Select
        
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Dim blnCancel As Boolean
    Dim strDataKey As String
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward               '上一条
        
        strDataKey = mstrDataKey
        
        RaiseEvent Forward(strDataKey, blnCancel)
        If blnCancel = False Then
        
            mstrDataKey = strDataKey
            Call ReadData(mstrDataKey)
    
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward               '下一条
        
        strDataKey = mstrDataKey
        
        RaiseEvent Backward(strDataKey, blnCancel)
        If blnCancel = False Then
            
            mstrDataKey = strDataKey
            Call ReadData(mstrDataKey)
            
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Filter
        
        Dim strText As String
        Dim rsCondition As ADODB.Recordset
        Dim rsData As ADODB.Recordset
        Dim rs As ADODB.Recordset
        
        If txtLocation.Text <> "" Then
            
            txtLocation.Tag = ""
            
            
            Set rsCondition = zlCommFun.CreateCondition
            
            Call zlCommFun.SetCondition(rsCondition, "FilterStyle", mstrFindKey)
            Call zlCommFun.SetCondition(rsCondition, "FilterText", txtLocation.Text)
            Set rsData = gclsBusiness.EventRead("FilterData", rsCondition)
            
            If zlCommFun.ShowPubSelect(Me, txtLocation, 2, "名称,1500,0,1;编码,1500,0,0;程序,1500,0,0;设备,1500,0,0", Me.Name & "\业务事件过滤", "请从下表中选择一个业务事件", rsData, rs, , , , , , True) = 1 Then
                mstrDataKey = rs("id").Value
                Call ReadData(mstrDataKey)
                txtLocation.Tag = ""
            Else
                txtLocation.Tag = ""
                Call LocationObj(txtLocation, True)
                Exit Sub
            End If
                        
            Call LocationObj(txtLocation, True)
        End If
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
    
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        mblnContiune = Not mblnContiune
    End Select
    
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    
    '窗体其它控件Resize处理
    picPane(0).Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_Filter, conMenu_View_LocationItem, conMenu_View_Backward, conMenu_View_Forward, 0
        Control.Visible = (mbytMode = 2)
        Control.Enabled = Not mblnDataChanged
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        Control.Checked = mblnContiune
        Control.IconId = IIf(mblnContiune = True, conMenu_View_Check, conMenu_View_UnCheck)
    End Select
End Sub

Private Sub cmdCancel_Click()
    '
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim lngStartRow As Long
    Dim lngStartCol As Long
    Dim lngEndRow As Long
    Dim lngEndCol As Long
    Dim intRow As Integer
    Dim intCol As Integer
    
    lngStartRow = -1
    With vsf(0)
        Call .GetSelection(lngStartRow, lngStartCol, lngEndRow, lngEndCol)
        If lngStartRow > 0 Then
        
            For intRow = lngStartRow To lngEndRow
                If Val(.Cell(flexcpData, intRow, 0)) = 0 Then
                    .Cell(flexcpData, intRow, lngStartCol, intRow, lngEndCol) = 0
                End If
            Next
            .Select lngStartRow, lngStartCol
            
            mblnDataChanged = True
        End If
    End With
End Sub

Private Sub cmdOK_Click()
        
    If mblnDataChanged = True Then
        If ValidData = True Then
                
            If SaveData(mstrDataKey) = True Then
                
                Select Case mbytMode
                Case 1
                    RaiseEvent AfterNewData(mstrDataKey)
                Case 2
                    RaiseEvent AfterModifyData(mstrDataKey)
                End Select
                
                If mblnContiune = False Then
                    mblnDataChanged = False
                    Unload Me
                Else
                    '重置环境，进入下一次新增状态
                    If mbytMode = 1 Then
                        mstrDataKey = ""
'                        txt(0).Text = ""
                        txt(1).Text = gclsBusiness.GetMaxCode("m_Event", "code")
                        txt(2).Text = ""
                        txt(3).Text = ""
                        txt(4).Text = ""
                        rtbSQL.Text = ""
                    End If
                    Call LocationObj(txt(0))
                    mblnDataChanged = False
                End If
                
            End If
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub cmdSet_Click()
    Dim lngStartRow As Long
    Dim lngStartCol As Long
    Dim lngEndRow As Long
    Dim lngEndCol As Long
    Dim intRow As Integer
    Dim intCol As Integer
    
    lngStartRow = -1
    With vsf(0)
        Call .GetSelection(lngStartRow, lngStartCol, lngEndRow, lngEndCol)
        If lngStartRow > 0 Then
            For intRow = lngStartRow To lngEndRow
                If Val(.Cell(flexcpData, intRow, 0)) = 0 Then
                    .Cell(flexcpData, intRow, lngStartCol, intRow, lngEndCol) = 1
                End If
            Next
            mblnDataChanged = True
        End If
        .Select lngStartRow, lngStartCol
    End With
End Sub

Private Sub cmdVerfiy_Click()
    '
End Sub

Private Sub dtp_Change(Index As Integer)
    mblnDataChanged = True
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
'    DoEvents
'    txt(1).SetFocus
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    
    Me.Width = 9510 - 300
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnDataChanged Then
        Cancel = (MsgBox("新增或修改的数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.系统名称) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    
    Set mobjFindKey = Nothing
    If Not (mclsVsf Is Nothing) Then
        Set mclsVsf = Nothing
    End If

    If Not (mrsPara Is Nothing) Then
        Set mrsPara = Nothing
    End If
    Set mfrmParent = Nothing
    
End Sub

Private Sub opt_Click(Index As Integer)
    Select Case Index
    Case 0, 1
        dtp(0).Visible = (opt(1).Value = True)
        dtp(1).Visible = dtp(0).Visible
        lbl(6).Visible = dtp(0).Visible
    Case 2, 3, 4, 5
        txt(6).Enabled = opt(4).Value
        upd(6).Enabled = txt(6).Enabled
        
        txt(7).Enabled = opt(5).Value
        upd(7).Enabled = txt(7).Enabled
    Case 6, 7
        cbo(0).Enabled = opt(7).Value
        txt(5).Enabled = cbo(0).Enabled
        upd(5).Enabled = cbo(0).Enabled
        vsf(0).Enabled = cbo(0).Enabled
        cmdSet.Enabled = vsf(0).Enabled
        cmdClear.Enabled = cmdSet.Enabled
    End Select
    mblnDataChanged = True
    
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
        Select Case Index
        Case 6
            If opt(6).Value Then
                tbcPage.Item(2).Selected = True
                rtbSQL.SetFocus
            End If
        Case Else
            zlCommFun.PressKey vbKeyTab
        End Select
        
    End If
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        tbcPage.Move 30, 30, picPane(Index).Width - 60, picPane(Index).Height - 600
        cmdCancel.Move picPane(Index).Width - cmdCancel.Width - 60, tbcPage.Top + tbcPage.Height + 60
        
        cmdOK.Move cmdCancel.Left - cmdOK.Width - 60, cmdCancel.Top
    Case 1
'        txt(0).Move txt(0).Left, txt(0).Top, picPane(Index).Width - txt(0).Left - cmd(0).Width - 45
'        cmd(0).Move txt(0).Left + txt(0).Width + 15
        
        txt(2).Move txt(2).Left, txt(2).Top, picPane(Index).Width - txt(2).Left
        txt(3).Width = txt(0).Width
        txt(4).Width = txt(0).Width
        txt(4).Height = picPane(Index).Height - txt(4).Top - 30
        fra(2).Width = txt(0).Width
        fra(1).Width = txt(0).Width
    Case 2
        fra(4).Move fra(4).Left, fra(4).Top, picPane(Index).Width - fra(4).Left - 15, picPane(Index).Height - fra(4).Top - 45
        vsf(0).Move 30, 525, fra(4).Width - 60, fra(4).Height - 525 - 30
        cmdClear.Left = vsf(0).Left + vsf(0).Width - cmdClear.Width
        cmdSet.Left = cmdClear.Left - cmdSet.Width - 60
    Case 3
        fra(0).Move 0, 0, picPane(Index).Width, picPane(Index).Height - fra(3).Height
        fra(3).Move fra(0).Left, fra(0).Top + fra(0).Height, fra(0).Width
        
        rtbSQL.Move rtbSQL.Left, rtbSQL.Top, fra(0).Width - 2 * rtbSQL.Left, fra(0).Height - rtbSQL.Top - 105
        cmdVerfiy.Left = rtbSQL.Left + rtbSQL.Width - cmdVerfiy.Width
    End Select
    
End Sub

Private Sub rtbSQL_Change()
    mblnDataChanged = True
End Sub

Private Sub rtbSQL_KeyPress(KeyAscii As Integer)
    '
End Sub

Private Sub txt_Change(Index As Integer)
    
    If mblnReading Then Exit Sub
    
    mblnDataChanged = True
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 4
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        
        '
        
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Select Case Index
        Case 4
            tbcPage.Item(1).Selected = True
            opt(6).SetFocus
        Case Else
            zlCommFun.PressKey vbKeyTab
        End Select
        
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 4
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
    Cancel = Not zlCommFun.StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub txtLocation_Change()
    txtLocation.Tag = "Changed"
End Sub

Private Sub txtLocation_GotFocus()
    zlControl.TxtSelAll txtLocation
End Sub

Private Sub txtLocation_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        KeyCode = 0
        txtLocation.Text = ""
        txtLocation.Tag = ""
    End If

End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim strText As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        If txtLocation.Text <> "" Then
            txtLocation.Tag = ""
            
            Dim obj As CommandBarControl
            
            Set obj = cbsMain.FindControl(, conMenu_View_Filter, True)
            If obj.Enabled = True Then
                Call cbsMain_Execute(obj)
            End If

        End If
        txtLocation.Tag = ""
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txtLocation_Validate(Cancel As Boolean)
    If (txtLocation.Tag = "Changed") Then
        txtLocation.Tag = ""
    End If
End Sub

Private Sub vsf_DrawCell(Index As Integer, ByVal hdc As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngSvrBkColor As Long
    Dim rc As RECT
    Dim rc1 As RECT
    Dim r1%, g1%, b1%
    Dim r2%, g2%, b2%
    Dim rg%, gg%, bg%
    Dim lngLoop As Long
    
'    Exit Sub
    If Row = 0 Or Col = 0 Then Exit Sub
    
    With vsf(Index)
        'flexODOver
        '--------------------------------------------------------------------------------------------------------------
'        rc.Left = Left + 50
'        rc.Top = Top + 2
'        rc.Right = Right - 5
'        rc.Bottom = Bottom - 3
        
        rc.Left = Left
        rc.Top = Top + 4
        rc.Right = Right
        rc.Bottom = Bottom - 4
        
        
        'Draw Frame
        '--------------------------------------------------------------------------------------------------------------
        lngSvrBkColor = SetBkColor(hdc, RGB(0, 0, 255))

        rc.Right = rc.Left + (rc.Right - rc.Left) * Val(.Cell(flexcpData, Row, Col))
        
        r1 = 180
        g1 = 180
        b1 = 255
        
        r2 = 180
        g2 = 180
        b2 = 255

        '画进度条
        '--------------------------------------------------------------------------------------------------------------
        rc1 = rc
        If rc.Right > rc.Left Then '
            For lngLoop = rc.Left To rc.Right Step 3

                rg = r1 + (lngLoop - rc.Left) * (r2 - r1) / (rc.Right - rc.Left)
                gg = g1 + (lngLoop - rc.Left) * (g2 - g1) / (rc.Right - rc.Left)
                bg = b1 + (lngLoop - rc.Left) * (b2 - b1) / (rc.Right - rc.Left)

'                Call SetBkColor(hdc, RGB(rg, gg, bg))
                Call SetBkColor(hdc, RGB(192, 192, 192))

                rc1.Left = lngLoop
                Call ExtTextOut(hdc, rc1.Left, rc1.Top, ETO_OPAQUE, rc1, " ", 1, lngLoop)
            Next

        End If
        
        'Done, Restore hDC And Quit
        '--------------------------------------------------------------------------------------------------------------
        Call SetBkColor(hdc, lngSvrBkColor)
        
        Done = True
    
    End With
End Sub

Private Sub vsf_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTemp As String
    Dim intMouseRow As Integer
    Dim intMouseCol As Integer
    Dim intTemp As Integer
    
    With vsf(Index)
        intMouseRow = .MouseRow
        intMouseCol = .MouseCol
        
        If intMouseRow < 1 Or intMouseCol < 1 Then
            .ToolTipText = ""
            Exit Sub
        End If
        
        If Val(.Cell(flexcpData, intMouseRow, 0)) > 0 Then
            .ToolTipText = ""
            Exit Sub
        End If
        
        intTemp = intMouseCol \ 2
        If intMouseCol Mod 2 = 1 Then
            strTemp = Format(intTemp, "00") & ":00" & "～" & Format(intTemp, "00") & ":30"
        Else
            strTemp = Format(intTemp - 1, "00") & ":30" & "～" & Format(intTemp, "00") & ":00"
        End If
        
        Select Case cbo(0).Text
        Case "天"
            
        Case "周"
            strTemp = .TextMatrix(intMouseRow, 0) & " " & strTemp
        Case "月"
            strTemp = "某月" & Val(.TextMatrix(intMouseRow, 0)) & "日 " & strTemp
        Case "年"
            strTemp = .TextMatrix(intMouseRow, 50) & "月" & Val(.TextMatrix(intMouseRow, 0)) & "日 " & strTemp
        End Select
        
        .ToolTipText = strTemp
    End With
    
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngStartRow As Long
    Dim lngStartCol As Long
    Dim lngEndRow As Long
    Dim lngEndCol As Long
    Dim intRow As Integer
    Dim intCol As Integer
    
    lngStartRow = -1
    Call vsf(0).GetSelection(lngStartRow, lngStartCol, lngEndRow, lngEndCol)
    If lngStartRow > 0 Then
        
    End If
    
End Sub

Private Function GetTime(ByVal intCol As Integer, ByVal bytMode As Byte) As String
    '******************************************************************************************************************
    '功能：根据点(列)计算出时间
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intTemp As Integer
    
    Select Case bytMode
    Case 1
        GetTime = Switch(intCol = 1, "00:00", intCol = 2, "00:30", intCol = 3, "01:00", intCol = 4, "01:30", intCol = 5, "02:00", intCol = 6, "02:30", _
                            intCol = 7, "03:00", intCol = 8, "03:30", intCol = 9, "04:00", intCol = 10, "04:30", intCol = 11, "05:00", intCol = 12, "05:30", _
                            intCol = 13, "06:00", intCol = 14, "06:30", intCol = 15, "07:00", intCol = 16, "07:30", intCol = 17, "08:00", intCol = 18, "08:30", _
                            intCol = 19, "09:00", intCol = 20, "09:30", intCol = 21, "10:00", intCol = 22, "10:30", intCol = 23, "11:00", intCol = 24, "11:30", _
                            intCol = 25, "12:00", intCol = 26, "12:30", intCol = 27, "13:00", intCol = 28, "13:30", intCol = 29, "14:00", intCol = 30, "14:30", _
                            intCol = 31, "15:00", intCol = 32, "15:30", intCol = 33, "16:00", intCol = 34, "16:30", intCol = 35, "17:00", intCol = 36, "17:30", _
                            intCol = 37, "18:00", intCol = 38, "18:30", intCol = 39, "19:00", intCol = 40, "19:30", intCol = 41, "20:00", intCol = 42, "20:30", _
                            intCol = 43, "21:00", intCol = 44, "21:30", intCol = 45, "22:00", intCol = 46, "22:30", intCol = 47, "23:00", intCol = 48, "23:30")
    Case 2
        GetTime = Switch(intCol = 1, "00:30", intCol = 2, "01:00", intCol = 3, "01:30", intCol = 4, "02:00", intCol = 5, "02:30", intCol = 6, "03:00", _
                            intCol = 7, "03:30", intCol = 8, "04:00", intCol = 9, "04:30", intCol = 10, "05:00", intCol = 11, "05:30", intCol = 12, "06:00", _
                            intCol = 13, "06:30", intCol = 14, "07:00", intCol = 15, "07:30", intCol = 16, "08:00", intCol = 17, "08:30", intCol = 18, "09:00", _
                            intCol = 19, "09:30", intCol = 20, "10:00", intCol = 21, "10:30", intCol = 22, "11:00", intCol = 23, "11:30", intCol = 24, "12:00", _
                            intCol = 25, "12:30", intCol = 26, "13:00", intCol = 27, "13:30", intCol = 28, "14:00", intCol = 29, "14:30", intCol = 30, "15:00", _
                            intCol = 31, "15:30", intCol = 32, "16:00", intCol = 33, "16:30", intCol = 34, "17:00", intCol = 35, "17:30", intCol = 36, "18:00", _
                            intCol = 37, "18:30", intCol = 38, "19:00", intCol = 39, "19:30", intCol = 40, "20:00", intCol = 41, "20:30", intCol = 42, "21:00", _
                            intCol = 43, "21:30", intCol = 44, "22:00", intCol = 45, "22:30", intCol = 46, "23:00", intCol = 47, "23:30", intCol = 48, "24:00")
                            

    End Select
    
End Function

Private Function GetColumn(ByVal strTime As String, ByVal bytMode As Byte) As Integer
    '******************************************************************************************************************
    '功能：根据时间计算出点（列）
    '参数：
    '返回：
    '******************************************************************************************************************
    Select Case bytMode
    Case 1
        GetColumn = Switch(strTime = "00:00", 1, strTime = "00:30", 2, strTime = "01:00", 3, strTime = "01:30", 4, strTime = "02:00", 5, strTime = "02:30", 6, _
                            strTime = "03:00", 7, strTime = "03:30", 8, strTime = "04:00", 9, strTime = "04:30", 10, strTime = "05:00", 11, strTime = "05:30", 12, _
                            strTime = "06:00", 13, strTime = "06:30", 14, strTime = "07:00", 15, strTime = "07:30", 16, strTime = "08:00", 17, strTime = "08:30", 18, _
                            strTime = "09:00", 19, strTime = "09:30", 20, strTime = "10:00", 21, strTime = "10:30", 22, strTime = "11:00", 23, strTime = "11:30", 24, _
                            strTime = "12:00", 25, strTime = "12:30", 26, strTime = "13:00", 27, strTime = "13:30", 28, strTime = "14:00", 29, strTime = "14:30", 30, _
                            strTime = "15:00", 31, strTime = "15:30", 32, strTime = "16:00", 33, strTime = "16:30", 34, strTime = "17:00", 35, strTime = "17:30", 36, _
                            strTime = "18:00", 37, strTime = "18:30", 38, strTime = "19:00", 39, strTime = "19:30", 40, strTime = "20:00", 41, strTime = "20:30", 42, _
                            strTime = "21:00", 43, strTime = "21:30", 44, strTime = "22:00", 45, strTime = "22:30", 46, strTime = "23:00", 47, strTime = "23:30", 48)
    Case 2
        GetColumn = Switch(strTime = "00:30", 1, strTime = "01:00", 2, strTime = "01:30", 3, strTime = "02:00", 4, strTime = "02:30", 5, strTime = "03:00", 6, _
                            strTime = "03:30", 7, strTime = "04:00", 8, strTime = "04:30", 9, strTime = "05:00", 10, strTime = "05:30", 11, strTime = "06:00", 12, _
                            strTime = "06:30", 13, strTime = "07:00", 14, strTime = "07:30", 15, strTime = "08:00", 16, strTime = "08:30", 17, strTime = "09:00", 18, _
                            strTime = "09:30", 19, strTime = "10:00", 20, strTime = "10:30", 21, strTime = "11:00", 22, strTime = "11:30", 23, strTime = "12:00", 24, _
                            strTime = "12:30", 25, strTime = "13:00", 26, strTime = "13:30", 27, strTime = "14:00", 28, strTime = "14:30", 29, strTime = "15:00", 30, _
                            strTime = "15:30", 31, strTime = "16:00", 32, strTime = "16:30", 33, strTime = "17:00", 34, strTime = "17:30", 35, strTime = "18:00", 36, _
                            strTime = "18:30", 37, strTime = "19:00", 38, strTime = "19:30", 39, strTime = "20:00", 40, strTime = "20:30", 41, strTime = "21:00", 42, _
                            strTime = "21:30", 43, strTime = "22:00", 44, strTime = "22:30", 45, strTime = "23:00", 46, strTime = "23:30", 47, strTime = "24:00", 48)
                            

    End Select
    
End Function

Private Function GetDays(ByVal intMonth As Integer) As Integer
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    
    For intLoop = 1 To intMonth
        GetDays = GetDays + Switch(intLoop = 1, 31, intLoop = 2, 29, intLoop = 3, 31, intLoop = 4, 30, intLoop = 5, 31, intLoop = 6, 30, _
                                    intLoop = 7, 31, intLoop = 8, 31, intLoop = 9, 30, intLoop = 10, 31, intLoop = 11, 30, intLoop = 12, 31)
    Next
    
End Function


