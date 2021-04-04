VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Begin VB.Form frmLabItemRef 
   BorderStyle     =   0  'None
   Caption         =   "检验项目参考值"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txt变异警示率 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   7230
      MaxLength       =   12
      TabIndex        =   38
      Top             =   465
      Width           =   645
   End
   Begin VB.TextBox txt比对失控率 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   3750
      MaxLength       =   12
      TabIndex        =   5
      Top             =   90
      Width           =   1020
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   1005
      Left            =   120
      TabIndex        =   7
      Top             =   825
      Width           =   7755
      _cx             =   13679
      _cy             =   1773
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
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
   Begin VB.TextBox txt比对警示率 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   1950
      MaxLength       =   12
      TabIndex        =   3
      Top             =   90
      Width           =   1020
   End
   Begin VB.TextBox txt变异报警率 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   5685
      MaxLength       =   12
      TabIndex        =   1
      Top             =   465
      Width           =   645
   End
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2055
      ScaleWidth      =   7740
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1860
      Width           =   7740
      Begin VB.ComboBox cbo科室 
         Height          =   300
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   975
         Width           =   2100
      End
      Begin VB.TextBox txt复查下限 
         Height          =   300
         Left            =   4185
         TabIndex        =   32
         Top             =   975
         Width           =   1020
      End
      Begin VB.TextBox txt复查上限 
         Height          =   300
         Left            =   5640
         TabIndex        =   33
         Top             =   975
         Width           =   1020
      End
      Begin VB.TextBox txt警戒下限 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   4170
         MaxLength       =   12
         TabIndex        =   28
         Top             =   660
         Width           =   1020
      End
      Begin VB.TextBox txt警戒上限 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   5625
         MaxLength       =   12
         TabIndex        =   29
         Top             =   660
         Width           =   1020
      End
      Begin VB.CheckBox chk默认 
         Alignment       =   1  'Right Justify
         Caption         =   "默认"
         Height          =   180
         Left            =   6825
         TabIndex        =   30
         Top             =   720
         Width           =   840
      End
      Begin VB.ComboBox cbo参考值 
         Height          =   300
         ItemData        =   "frmLabItemRef.frx":0000
         Left            =   3765
         List            =   "frmLabItemRef.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   345
         Width           =   1200
      End
      Begin VB.TextBox txt可偏移率 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   7005
         MaxLength       =   13
         TabIndex        =   24
         Top             =   345
         Width           =   750
      End
      Begin VB.TextBox txt备注 
         Height          =   300
         Left            =   900
         MaxLength       =   50
         TabIndex        =   34
         Top             =   1290
         Width           =   6690
      End
      Begin VB.ComboBox cbo仪器 
         Height          =   300
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   660
         Width           =   2100
      End
      Begin VB.ComboBox cbo临床特征 
         Height          =   300
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   345
         Width           =   2100
      End
      Begin VB.TextBox txt参考值 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   4905
         MaxLength       =   13
         TabIndex        =   21
         Top             =   345
         Width           =   900
      End
      Begin VB.TextBox txt参考值 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   3780
         MaxLength       =   13
         TabIndex        =   20
         Top             =   345
         Width           =   900
      End
      Begin VB.ComboBox cbo单位 
         Height          =   300
         Left            =   7020
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   30
         Width           =   750
      End
      Begin VB.TextBox txt年龄 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   6495
         MaxLength       =   3
         TabIndex        =   15
         Top             =   30
         Width           =   495
      End
      Begin VB.ComboBox cbo性别 
         Height          =   300
         Left            =   3765
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   30
         Width           =   1200
      End
      Begin VB.TextBox txt年龄 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   5775
         MaxLength       =   3
         TabIndex        =   14
         Top             =   30
         Width           =   495
      End
      Begin VB.ComboBox cbo标本 
         Height          =   300
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   30
         Width           =   2100
      End
      Begin VB.Label lbl仪器 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "对应科室"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   75
         TabIndex        =   40
         Top             =   1035
         Width           =   720
      End
      Begin VB.Label lbl警戒下限 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "警示参考              ～"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3345
         TabIndex        =   27
         Top             =   720
         Width           =   2160
      End
      Begin VB.Label lbl警戒上限 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "复查参考              ～"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3345
         TabIndex        =   39
         Top             =   1035
         Width           =   2160
      End
      Begin VB.Label lbl可偏移率 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "可偏移率"
         Height          =   180
         Left            =   6240
         TabIndex        =   23
         Top             =   405
         Width           =   720
      End
      Begin VB.Label lbl备注 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注"
         Height          =   180
         Left            =   420
         TabIndex        =   35
         Top             =   1350
         Width           =   360
      End
      Begin VB.Label lbl仪器 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "对应仪器"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   75
         TabIndex        =   25
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lbl临床特征 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "临床特征"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   75
         TabIndex        =   17
         Top             =   405
         Width           =   720
      End
      Begin VB.Label lbl参考 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "参考           ～"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3330
         TabIndex        =   19
         Top             =   405
         Width           =   1530
      End
      Begin VB.Label lbl性别 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3345
         TabIndex        =   11
         Top             =   90
         Width           =   360
      End
      Begin VB.Label lbl年龄 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "年龄      ～"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5400
         TabIndex        =   13
         Top             =   90
         Width           =   1080
      End
      Begin VB.Label lbl标本 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "标本种类"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   75
         TabIndex        =   9
         Top             =   90
         Width           =   720
      End
      Begin XtremeCommandBars.CommandBars cbsThis 
         Left            =   60
         Top             =   1665
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         VisualTheme     =   2
      End
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "变异警示"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6450
      TabIndex        =   37
      Top             =   525
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(相同样本在仪器上检测结果的比对)"
      Height          =   180
      Left            =   4890
      TabIndex        =   36
      Top             =   135
      Width           =   2880
   End
   Begin VB.Label lbl失控率 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "失控率"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3150
      TabIndex        =   4
      Top             =   150
      Width           =   540
   End
   Begin VB.Label lblList 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "参考值详细列表:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   105
      TabIndex        =   6
      Top             =   525
      Width           =   1350
   End
   Begin VB.Label lbl比对警示率 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "结果比对检查：警示率"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   105
      TabIndex        =   2
      Top             =   150
      Width           =   1800
   End
   Begin VB.Label lbl变异报警率 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "变异报警"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4920
      TabIndex        =   0
      Top             =   525
      Width           =   720
   End
End
Attribute VB_Name = "frmLabItemRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngItemID As Long          '当前显示的项目id
Private mInt类型 As Integer         '当前项目的结果类型

Private Enum mCol
    默认 = 0: 标本: 性别域: 性别: 年龄下限: 年龄上限: 年龄单位: 年龄显示: 临床特征: 参考低值: 参考高值: 参考显示: 警示低值: 警示高值: 警示显示: 复查低值: 复查高值: 复查显示: 可偏移率: 仪器Id: 仪器名: 申请科室Id: 申请科室: 备注
End Enum

'临时变量
Dim cbrControl As CommandBarControl

Dim lngCount As Long
Dim strTemp As String, aryTemp() As String

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Private Sub RowRefresh(lngRow As Long)
    '功能：设置执行数据行的年龄和参考值显示
    '参数： lngRow-数据行
    Dim lngList As Long
    
    With Me.vfgList
        
        Select Case Val(.TextMatrix(lngRow, mCol.性别域))
        Case 1: .TextMatrix(lngRow, mCol.性别) = "男"
        Case 2: .TextMatrix(lngRow, mCol.性别) = "女"
        Case Else: .TextMatrix(lngRow, mCol.性别) = ""
        End Select
        
        .TextMatrix(lngRow, mCol.年龄下限) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.年龄下限)), " .", "0."), " ", "")
        .TextMatrix(lngRow, mCol.年龄上限) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.年龄上限)), " .", "0."), " ", "")
        .TextMatrix(lngRow, mCol.参考低值) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.参考低值)), " .", "0."), " ", "")
        .TextMatrix(lngRow, mCol.参考高值) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.参考高值)), " .", "0."), " ", "")
        .TextMatrix(lngRow, mCol.警示低值) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.警示低值)), " .", "0."), " ", "")
        .TextMatrix(lngRow, mCol.警示高值) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.警示高值)), " .", "0."), " ", "")
        .TextMatrix(lngRow, mCol.复查低值) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.复查低值)), " .", "0."), " ", "")
        .TextMatrix(lngRow, mCol.复查高值) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.复查高值)), " .", "0."), " ", "")
        
        
        If Val(.TextMatrix(lngRow, mCol.年龄下限)) <> 0 And Val(.TextMatrix(lngRow, mCol.年龄上限)) <> 0 Then
            .TextMatrix(lngRow, mCol.年龄显示) = .TextMatrix(lngRow, mCol.年龄下限) & "～" & .TextMatrix(lngRow, mCol.年龄上限) & .TextMatrix(lngRow, mCol.年龄单位)
        ElseIf Val(.TextMatrix(lngRow, mCol.年龄下限)) <> 0 Then
            .TextMatrix(lngRow, mCol.年龄显示) = .TextMatrix(lngRow, mCol.年龄下限) & .TextMatrix(lngRow, mCol.年龄上限) & .TextMatrix(lngRow, mCol.年龄单位) & "～"
        ElseIf Val(.TextMatrix(lngRow, mCol.年龄上限)) <> 0 Then
            .TextMatrix(lngRow, mCol.年龄显示) = "～" & .TextMatrix(lngRow, mCol.年龄上限) & .TextMatrix(lngRow, mCol.年龄单位)
        Else
            .TextMatrix(lngRow, mCol.年龄显示) = ""
        End If
        
        .TextMatrix(lngRow, mCol.参考显示) = ""
        .TextMatrix(lngRow, mCol.警示显示) = ""
        .TextMatrix(lngRow, mCol.复查显示) = ""
        Select Case mInt类型
        Case 1, 3  '定量和定性(定性也需要有参考值)
'            .TextMatrix(lngRow, mcol.参考低值) = FormatReference(mlngItemID, .TextMatrix(lngRow, mcol.参考低值))
'            .TextMatrix(lngRow, mcol.参考高值) = FormatReference(mlngItemID, .TextMatrix(lngRow, mcol.参考高值))
            If IsNumeric(.TextMatrix(lngRow, mCol.参考低值)) = True And IsNumeric(.TextMatrix(lngRow, mCol.参考高值)) = True Then
                .TextMatrix(lngRow, mCol.参考显示) = .TextMatrix(lngRow, mCol.参考低值) & "～" & .TextMatrix(lngRow, mCol.参考高值)
            ElseIf IsNumeric(.TextMatrix(lngRow, mCol.参考低值)) = True Then
                .TextMatrix(lngRow, mCol.参考显示) = .TextMatrix(lngRow, mCol.参考低值) & "～"
            ElseIf IsNumeric(.TextMatrix(lngRow, mCol.参考高值)) = True Then
                .TextMatrix(lngRow, mCol.参考显示) = "～" & .TextMatrix(lngRow, mCol.参考高值)
            End If
        
            If IsNumeric(.TextMatrix(lngRow, mCol.警示低值)) = True And IsNumeric(.TextMatrix(lngRow, mCol.警示高值)) = True Then
                .TextMatrix(lngRow, mCol.警示显示) = .TextMatrix(lngRow, mCol.警示低值) & "～" & .TextMatrix(lngRow, mCol.警示高值)
            ElseIf IsNumeric(.TextMatrix(lngRow, mCol.警示低值)) = True Then
                .TextMatrix(lngRow, mCol.警示显示) = .TextMatrix(lngRow, mCol.警示低值) & "～"
            ElseIf IsNumeric(.TextMatrix(lngRow, mCol.警示高值)) = True Then
                .TextMatrix(lngRow, mCol.警示显示) = "～" & .TextMatrix(lngRow, mCol.警示高值)
            End If
        
            If IsNumeric(.TextMatrix(lngRow, mCol.复查低值)) = True And IsNumeric(.TextMatrix(lngRow, mCol.复查高值)) = True Then
                .TextMatrix(lngRow, mCol.复查显示) = .TextMatrix(lngRow, mCol.复查低值) & "～" & .TextMatrix(lngRow, mCol.复查高值)
            ElseIf IsNumeric(.TextMatrix(lngRow, mCol.复查低值)) = True Then
                .TextMatrix(lngRow, mCol.复查显示) = .TextMatrix(lngRow, mCol.复查低值) & "～"
            ElseIf IsNumeric(.TextMatrix(lngRow, mCol.复查高值)) = True Then
                .TextMatrix(lngRow, mCol.复查显示) = "～" & .TextMatrix(lngRow, mCol.复查高值)
            End If
        Case 2  '半定量
            For lngList = 0 To Me.cbo参考值.ListCount - 1
                If lngList = Val(.TextMatrix(lngRow, mCol.参考低值)) And IsNumeric(.TextMatrix(lngRow, mCol.参考低值)) = True Then
                    .TextMatrix(lngRow, mCol.参考显示) = Me.cbo参考值.List(lngList + 1): Exit For
                End If
            Next
        End Select
    End With
End Sub

Private Sub setListFormat(Optional blnKeepData As Boolean)
    '功能：初始化设置参考值列表
    '参数： blnKeepData-是否保留数据，即只是重新设置格式
    With Me.vfgList
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 1: .FixedRows = 1: .Cols = 24: .FixedCols = 0
        End If
        .TextMatrix(0, mCol.默认) = "默认":
        .TextMatrix(0, mCol.标本) = "标本": .TextMatrix(0, mCol.性别域) = "性别域": .TextMatrix(0, mCol.性别) = "性别"
        .TextMatrix(0, mCol.年龄下限) = "年龄下限": .TextMatrix(0, mCol.年龄上限) = "年龄上限"
        .TextMatrix(0, mCol.年龄单位) = "年龄单位": .TextMatrix(0, mCol.年龄显示) = "年龄"
        .TextMatrix(0, mCol.临床特征) = "临床特征"
        .TextMatrix(0, mCol.参考低值) = "参考低值": .TextMatrix(0, mCol.参考高值) = "参考高值": .TextMatrix(0, mCol.参考显示) = "参考值"
        .TextMatrix(0, mCol.警示低值) = "警示低值": .TextMatrix(0, mCol.警示高值) = "警示高值": .TextMatrix(0, mCol.警示显示) = "警示值"
        .TextMatrix(0, mCol.复查低值) = "复查低值": .TextMatrix(0, mCol.复查高值) = "复查高值": .TextMatrix(0, mCol.复查显示) = "复查值"
        .TextMatrix(0, mCol.可偏移率) = "可偏移率"
        .TextMatrix(0, mCol.仪器Id) = "仪器id": .TextMatrix(0, mCol.仪器名) = "仪器名"
        .TextMatrix(0, mCol.申请科室Id) = "科室id": .TextMatrix(0, mCol.申请科室) = "申请科室"
        .TextMatrix(0, mCol.备注) = "备注"
        
        .ColWidth(mCol.默认) = 500
        .ColWidth(mCol.标本) = 1000: .ColWidth(mCol.性别域) = 0: .ColWidth(mCol.性别) = 700
        .ColWidth(mCol.年龄下限) = 0: .ColWidth(mCol.年龄上限) = 0
        .ColWidth(mCol.年龄单位) = 0: .ColWidth(mCol.年龄显示) = 1000
        .ColWidth(mCol.临床特征) = 1200
        .ColWidth(mCol.参考低值) = 0: .ColWidth(mCol.参考高值) = 0: .ColWidth(mCol.参考显示) = 1300
        .ColWidth(mCol.警示低值) = 0: .ColWidth(mCol.警示高值) = 0: .ColWidth(mCol.警示显示) = 1300
        .ColWidth(mCol.复查低值) = 0: .ColWidth(mCol.复查高值) = 0: .ColWidth(mCol.复查显示) = 1300
        .ColWidth(mCol.可偏移率) = 900
        .ColWidth(mCol.仪器Id) = 0: .ColWidth(mCol.仪器名) = 1500
        .ColWidth(mCol.申请科室Id) = 0: .ColWidth(mCol.申请科室) = 1500
        .ColWidth(mCol.备注) = 1500
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .ColDataType(mCol.默认) = flexDTBoolean

        '处理年龄和参考值的显示
        For lngCount = .FixedRows To .Rows - 1
            Call RowRefresh(lngCount)
        Next
        .Redraw = flexRDDirect
    End With
    
End Sub

Public Function zlRefresh(lngItemID As Long) As Boolean
    '功能：根据项目id刷新当前显示内容
    '参数：当前项目id
    Dim rsTemp As New ADODB.Recordset
    
    mlngItemID = lngItemID
    
    '清除此前项目的显示
    Me.txt警戒下限.Text = "": Me.txt警戒上限.Text = ""
    Me.txt变异报警率.Text = "": Me.txt比对警示率.Text = "": Me.txt比对失控率.Text = "": Me.txt变异警示率.Text = ""
        
    If lngItemID = 0 Then Call setListFormat: zlRefresh = True: Exit Function
    
    '获取指定项目的信息
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select I.结果类型, I.警戒下限, I.警戒上限, I.变异报警率, I.比对警示率, I.比对失控率, I.取值序列,I.变异警示率 " & vbNewLine & _
            "From 检验项目 I, 检验报告项目 R" & vbNewLine & _
            "Where I.诊治项目id = R.报告项目id And R.诊疗项目id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    With rsTemp
        mInt类型 = 1
        If .RecordCount > 0 Then
'            Me.txt警戒下限.Text = "" & !警戒下限: Me.txt警戒上限.Text = "" & !警戒上限
            Me.txt变异报警率.Text = "" & !变异报警率: Me.txt变异警示率.Text = "" & !变异警示率
            Me.txt比对警示率.Text = "" & !比对警示率: Me.txt比对失控率.Text = "" & !比对失控率
            Me.cbo参考值.Tag = "" & !取值序列

            mInt类型 = Val("" & !结果类型)
            
            'Me.txt警戒下限.Text = Replace(Replace(" " & Trim(Me.txt警戒下限.Text), " .", "0."), " ", "")
            'Me.txt警戒上限.Text = Replace(Replace(" " & Trim(Me.txt警戒上限.Text), " .", "0."), " ", "")
            Me.txt变异报警率.Text = Replace(Replace(" " & Trim(Me.txt变异报警率.Text), " .", "0."), " ", "")
            Me.txt变异警示率.Text = Replace(Replace(" " & Trim(Me.txt变异警示率.Text), " .", "0."), " ", "")
            Me.txt比对警示率.Text = Replace(Replace(" " & Trim(Me.txt比对警示率.Text), " .", "0."), " ", "")
            Me.txt比对失控率.Text = Replace(Replace(" " & Trim(Me.txt比对失控率.Text), " .", "0."), " ", "")
        End If
    End With
    
    Me.lbl参考.Visible = False
    Me.txt参考值(0).Visible = False: Me.txt参考值(1).Visible = False: Me.txt可偏移率.Visible = False: lbl可偏移率.Visible = False
    Me.txt警戒上限.Visible = False: Me.txt警戒下限.Visible = False
    Me.txt复查上限.Visible = False: Me.txt复查下限.Visible = False
    Me.lbl警戒上限.Visible = False: Me.lbl警戒下限.Visible = False
    Me.cbo参考值.Visible = False
    Select Case mInt类型
    Case 3 '半定量
        Me.lbl参考.Visible = True
        Me.txt参考值(0).Visible = True: Me.txt参考值(1).Visible = True
        Me.lbl警戒上限.Visible = True: Me.lbl警戒下限.Visible = True
        Me.txt警戒上限.Visible = True: Me.txt警戒下限.Visible = True
        Me.txt复查上限.Visible = True: Me.txt复查下限.Visible = True
    Case 2 '文字型
        Me.lbl参考.Visible = True
        With Me.cbo参考值
            .Clear
            aryTemp() = Split(.Tag, ";")
            .AddItem ""
            For lngCount = LBound(aryTemp) To UBound(aryTemp)
                .AddItem aryTemp(lngCount)
            Next
            .Visible = True
        End With
        
    Case Else '数字型
        Me.lbl参考.Visible = True
        Me.txt参考值(0).Visible = True: Me.txt参考值(1).Visible = True: Me.txt可偏移率.Visible = True: lbl可偏移率.Visible = True
        Me.lbl警戒上限.Visible = True: Me.lbl警戒下限.Visible = True
        Me.txt警戒上限.Visible = True: Me.txt警戒下限.Visible = True
        Me.txt复查上限.Visible = True: Me.txt复查下限.Visible = True
    End Select
        
    gstrSql = "Select nvl(L.默认,0) As 默认,L.标本类型 As 标本, L.性别域, '' As 性别, L.年龄下限, L.年龄上限, L.年龄单位, '' As 年龄, L.临床特征," & vbNewLine & _
            "       L.参考低值, L.参考高值, '' As 参考,L.警示下限,L.警示上限,'' as 警示,L.复查下限,L.复查上限,'' as 复查, L.可偏移率, L.仪器id, M.名称 As 仪器名, L.申请科室ID, N.名称 as 申请科室, L.备注 " & vbNewLine & _
            "From 检验项目参考 L, 检验报告项目 R, 检验仪器 M, 部门表 N" & vbNewLine & _
            "Where L.项目id = R.报告项目id And L.仪器id = M.ID(+) And L.申请科室ID=N.ID(+) And R.诊疗项目id = [1] Order by L.标本类型"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    Set Me.vfgList.DataSource = rsTemp:
    Call setListFormat(True)
    If Me.vfgList.Rows <= Me.vfgList.FixedRows Then
        Me.vfgList.Rows = Me.vfgList.FixedRows + 1
        Me.vfgList.Row = Me.vfgList.FixedRows
    End If
    Call vfgList_RowColChange
    
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart() As Boolean
    '功能：开始项目编辑
    '参数： lngItemId-指定编辑的项目
        
    If Me.cbo标本.ListCount = 0 Then
        MsgBox "请先在字典中初始化“检验标本”！", vbInformation, gstrSysName
        zlEditStart = False: Exit Function
    End If
    
    Me.Tag = "编辑": Call Form_Resize
    If Me.Visible Then Me.txt比对警示率.SetFocus
    zlEditStart = True: Exit Function

End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mlngItemID)
End Sub

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
    Dim strValue As String
    '一般特性检查
    If Val(Me.txt警戒下限.Text) <> 0 And Val(Me.txt警戒上限.Text) <> 0 Then
        If Val(Me.txt警戒下限.Text) >= Val(Me.txt警戒上限.Text) Then
            MsgBox "警戒下限必须低于警戒上限", vbInformation, gstrSysName
            Me.txt警戒下限.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    If Val(Me.txt警戒下限.Text) <> 0 Then
        If Val(Me.txt警戒下限.Text) > 999999999 Or Val(Val(Me.txt警戒下限.Text) * 100000) - Int(Val(Val(Me.txt警戒下限.Text) * 100000)) > 0 Then
            MsgBox "警戒下限数值太大或精度太高！", vbInformation, gstrSysName
            Me.txt警戒下限.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    If Val(Me.txt警戒上限.Text) <> 0 Then
        If Val(Me.txt警戒上限.Text) > 999999999 Or Val(Val(Me.txt警戒上限.Text) * 100000) - Int(Val(Val(Me.txt警戒上限.Text) * 100000)) > 0 Then
            MsgBox "警戒上限数值太大或精度太高！", vbInformation, gstrSysName
            Me.txt警戒上限.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    If Val(Me.txt变异报警率.Text) <> 0 Then
        If Val(Me.txt变异报警率.Text) > 999999999 Or Val(Val(Me.txt变异报警率.Text) * 100000) - Int(Val(Val(Me.txt变异报警率.Text) * 100000)) > 0 Then
            MsgBox "变异报警率太大或精度太高！", vbInformation, gstrSysName
            Me.txt变异报警率.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    
    If Val(Me.txt变异警示率.Text) <> 0 Then
        If Val(Me.txt变异警示率.Text) > 999999999 Or Val(Val(Me.txt变异警示率.Text) * 100000) - Int(Val(Val(Me.txt变异警示率.Text) * 100000)) > 0 Then
            MsgBox "变异报警率太大或精度太高！", vbInformation, gstrSysName
            Me.txt变异警示率.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    
    
    If Val(Me.txt比对警示率.Text) <> 0 Then
        If Val(Me.txt比对警示率.Text) > 999999999 Or Val(Val(Me.txt比对警示率.Text) * 100000) - Int(Val(Val(Me.txt比对警示率.Text) * 100000)) > 0 Then
            MsgBox "比对警示率太大或精度太高！", vbInformation, gstrSysName
            Me.txt比对警示率.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    If Val(Me.txt比对失控率.Text) <> 0 Then
        If Val(Me.txt比对失控率.Text) > 999999999 Or Val(Val(Me.txt比对失控率.Text) * 100000) - Int(Val(Val(Me.txt比对失控率.Text) * 100000)) > 0 Then
            MsgBox "比对失控率太大或精度太高！", vbInformation, gstrSysName
            Me.txt比对失控率.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    
    '参考值列表检查
    Dim strLists As String, strItems As String
    With Me.vfgList
        strLists = ""
        For lngCount = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(lngCount, mCol.标本)) = "" Then
                MsgBox "第" & lngCount & "参考值未填写标本，请填写或删除！", vbInformation, gstrSysName
                Me.cbo标本.SetFocus: zlEditSave = 0: Exit Function
            End If
            
            strItems = Trim(.TextMatrix(lngCount, mCol.标本)) & ";"
            If Val(.TextMatrix(lngCount, mCol.性别域)) <> 0 Then strItems = strItems & Val(.TextMatrix(lngCount, mCol.性别域))
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.年龄下限)) = True Then strItems = strItems & Val(.TextMatrix(lngCount, mCol.年龄下限))
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.年龄上限)) = True Then strItems = strItems & Val(.TextMatrix(lngCount, mCol.年龄上限))
            strItems = strItems & ";" & Trim(.TextMatrix(lngCount, mCol.年龄单位))
            strItems = strItems & ";" & Trim(.TextMatrix(lngCount, mCol.临床特征))
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.参考低值)) = True Then
                strItems = strItems & .TextMatrix(lngCount, mCol.参考低值)
                strValue = .TextMatrix(lngCount, mCol.参考低值)
                If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                    MsgBox "第" & lngCount & "参考值太大或精度太高！", vbInformation, gstrSysName
                    Me.cbo标本.SetFocus: zlEditSave = 0: Exit Function
                End If
            End If
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.参考高值)) = True Then
                strItems = strItems & .TextMatrix(lngCount, mCol.参考高值)
                strValue = .TextMatrix(lngCount, mCol.参考高值)
                If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                    MsgBox "第" & lngCount & "参考值太大或精度太高！", vbInformation, gstrSysName
                    Me.cbo标本.SetFocus: zlEditSave = 0: Exit Function
                End If
            End If
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.可偏移率)) = True Then
                strItems = strItems & Val(.TextMatrix(lngCount, mCol.可偏移率))
                strValue = .TextMatrix(lngCount, mCol.可偏移率)
                If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                    MsgBox "第" & lngCount & "可偏移率太大或精度太高！", vbInformation, gstrSysName
                    Me.cbo标本.SetFocus: zlEditSave = 0: Exit Function
                End If
            End If
            strItems = strItems & ";"
            If Val(.TextMatrix(lngCount, mCol.仪器Id)) <> 0 Then strItems = strItems & Val(.TextMatrix(lngCount, mCol.仪器Id))
            strItems = strItems & ";" & Trim(.TextMatrix(lngCount, mCol.备注))
            strItems = strItems & ";" & Trim(.TextMatrix(lngCount, mCol.默认))
            
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.警示高值)) = True Then
                strItems = strItems & .TextMatrix(lngCount, mCol.警示高值)
                strValue = .TextMatrix(lngCount, mCol.警示高值)
                If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                    MsgBox "第" & lngCount & "警示值太大或精度太高！", vbInformation, gstrSysName
                    Me.cbo标本.SetFocus: zlEditSave = 0: Exit Function
                End If
            End If
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.警示低值)) = True Then
                strItems = strItems & .TextMatrix(lngCount, mCol.警示低值)
                strValue = .TextMatrix(lngCount, mCol.警示低值)
                If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                    MsgBox "第" & lngCount & "警示值太大或精度太高！", vbInformation, gstrSysName
                    Me.cbo标本.SetFocus: zlEditSave = 0: Exit Function
                End If
            End If
            
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.复查高值)) = True Then
                strItems = strItems & .TextMatrix(lngCount, mCol.复查高值)
                strValue = .TextMatrix(lngCount, mCol.复查高值)
                If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                    MsgBox "第" & lngCount & "复查值太大或精度太高！", vbInformation, gstrSysName
                    Me.cbo标本.SetFocus: zlEditSave = 0: Exit Function
                End If
            End If
            strItems = strItems & ";"
            If IsNumeric(.TextMatrix(lngCount, mCol.复查低值)) = True Then
                strItems = strItems & .TextMatrix(lngCount, mCol.复查低值)
                strValue = .TextMatrix(lngCount, mCol.复查低值)
                If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                    MsgBox "第" & lngCount & "复查值太大或精度太高！", vbInformation, gstrSysName
                    Me.cbo标本.SetFocus: zlEditSave = 0: Exit Function
                End If
            End If
            strItems = strItems & ";"
            If Val(.TextMatrix(lngCount, mCol.申请科室Id)) <> 0 Then strItems = strItems & Val(.TextMatrix(lngCount, mCol.申请科室Id))
            
            strLists = strLists & "|" & strItems
        Next
    End With
    If strLists <> "" Then strLists = Mid(strLists, 2)

    '数据保存语句组织
    gstrSql = "Zl_检验项目参考_Edit(" & mlngItemID & "," & IIf(Trim(Me.txt警戒下限.Text) = "", "''", Val(Me.txt警戒下限.Text)) & "," & _
              IIf(Val(Me.txt警戒上限.Text) = 0, "''", Val(Me.txt警戒上限.Text)) & "," & IIf(Val(Me.txt变异报警率.Text) = 0, "''", Val(Me.txt变异报警率.Text)) & _
              "," & Val(Me.txt变异警示率.Text) & _
              "," & IIf(Val(Me.txt比对警示率.Text) = 0, "''", Val(Me.txt比对警示率.Text)) & _
              "," & IIf(Val(Me.txt比对失控率.Text) = 0, "''", Val(Me.txt比对失控率.Text)) & ",'" & strLists & "')"
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    Me.Tag = "": Call Form_Resize
    zlEditSave = mlngItemID: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

Private Sub cbo标本_Click()
'   根据标本限制性别
    If Me.cbo标本.ListIndex >= 0 Then
        If Me.cbo标本.ItemData(Me.cbo标本.ListIndex) = 1 Then
            Me.cbo性别.ListIndex = 1
            Me.cbo性别.Enabled = False
        ElseIf cbo标本.ItemData(Me.cbo标本.ListIndex) = 2 Then
            Me.cbo性别.ListIndex = 2
            Me.cbo性别.Enabled = False
        Else
            Me.cbo性别.Enabled = True
        End If
    End If
End Sub

'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------

Private Sub cbo标本_GotFocus()
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub cbo标本_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo参考值_GotFocus()
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub cbo参考值_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo单位_GotFocus()
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub cbo单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo临床特征_GotFocus()
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub cbo临床特征_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo性别_GotFocus()
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub cbo性别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbo仪器_GotFocus()
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub cbo仪器_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngCurRow As Long, lngRow As Long, lngCol As Long
    Dim Str标本 As String, intRow As Integer
    With Me.vfgList
        Select Case Control.ID
        Case conMenu_Edit_NewItem
            .Rows = .Rows + 1: .Row = .Rows - 1
        Case conMenu_Edit_Delete
            If .Row = .Rows - 1 Then
                .Rows = .Rows - 1: .Row = .Rows - 1
            Else
                lngCurRow = .Row
                For lngRow = lngCurRow To .Rows - 2
                    For lngCol = 0 To .Cols - 1
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow + 1, lngCol)
                    Next
                Next
                .Rows = .Rows - 1
            End If
        Case conMenu_Edit_Adjust
            If Me.cbo标本.ListIndex = -1 Then
                .TextMatrix(.Row, mCol.标本) = ""
            Else
                .TextMatrix(.Row, mCol.标本) = Mid(Me.cbo标本.Text, 4)
            End If
            .TextMatrix(.Row, mCol.性别域) = Val(Left(Me.cbo性别.Text, 1))
            .TextMatrix(.Row, mCol.年龄下限) = Me.txt年龄(0).Text
            .TextMatrix(.Row, mCol.年龄上限) = Me.txt年龄(1).Text
            
            If Me.cbo单位.ListIndex = -1 Then
                .TextMatrix(.Row, mCol.年龄单位) = "年"
            Else
                .TextMatrix(.Row, mCol.年龄单位) = Mid(Me.cbo单位.Text, 3)
            End If
            If Me.cbo临床特征.ListIndex = -1 Then
                .TextMatrix(.Row, mCol.临床特征) = ""
            Else
                .TextMatrix(.Row, mCol.临床特征) = Mid(Me.cbo临床特征.Text, 4)
            End If
            Select Case mInt类型
            Case 1, 3
                .TextMatrix(.Row, mCol.参考低值) = FormatReference(mlngItemID, IIf(IsNumeric(txt参考值(0)), Me.txt参考值(0), ""))
                .TextMatrix(.Row, mCol.参考高值) = FormatReference(mlngItemID, IIf(IsNumeric(txt参考值(1)), Me.txt参考值(1), ""))
                .TextMatrix(.Row, mCol.警示低值) = FormatReference(mlngItemID, IIf(IsNumeric(txt警戒下限), Me.txt警戒下限, ""))
                .TextMatrix(.Row, mCol.警示高值) = FormatReference(mlngItemID, IIf(IsNumeric(txt警戒上限), Me.txt警戒上限, ""))
                .TextMatrix(.Row, mCol.复查低值) = FormatReference(mlngItemID, IIf(IsNumeric(txt复查下限), Me.txt复查下限, ""))
                .TextMatrix(.Row, mCol.复查高值) = FormatReference(mlngItemID, IIf(IsNumeric(txt复查上限), Me.txt复查上限, ""))
                .TextMatrix(.Row, mCol.可偏移率) = IIf(Val(Me.txt可偏移率.Text) = 0, "", Val(Me.txt可偏移率.Text))
                
                '--- 2012-1-19 因为加入了申请科室条件，配合贵医的修改，改为让用户自已设定每个参考的警示上限，下限。
'                For intRow = .FixedRows To .Rows - 1
'                    If .TextMatrix(intRow, mCol.标本) = .TextMatrix(.Row, mCol.标本) Then
'                        .TextMatrix(intRow, mCol.警示低值) = FormatReference(mlngItemID, IIf(IsNumeric(txt警戒下限), Me.txt警戒下限, ""))
'                        .TextMatrix(intRow, mCol.警示高值) = FormatReference(mlngItemID, IIf(IsNumeric(txt警戒上限), Me.txt警戒上限, ""))
'                        .TextMatrix(intRow, mCol.复查低值) = FormatReference(mlngItemID, IIf(IsNumeric(txt复查下限), Me.txt复查下限, ""))
'                        .TextMatrix(intRow, mCol.复查高值) = FormatReference(mlngItemID, IIf(IsNumeric(txt复查上限), Me.txt复查上限, ""))
'                        Call RowRefresh(CLng(intRow))
'                    End If
'                Next
                
            Case 2
                .TextMatrix(.Row, mCol.参考低值) = IIf(Me.cbo参考值.ListIndex = 0, "", Me.cbo参考值.ListIndex - 1)
                .TextMatrix(.Row, mCol.参考高值) = IIf(Me.cbo参考值.ListIndex = 0, "", Me.cbo参考值.ListIndex - 1)
                .TextMatrix(.Row, mCol.可偏移率) = ""
            End Select
            If Me.cbo仪器.ListIndex = -1 Then
                .TextMatrix(.Row, mCol.仪器Id) = 0: .TextMatrix(.Row, mCol.仪器名) = ""
            Else
                .TextMatrix(.Row, mCol.仪器Id) = Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex)
                .TextMatrix(.Row, mCol.仪器名) = Mid(Me.cbo仪器.Text, 5)
            End If
            .TextMatrix(.Row, mCol.备注) = Trim(Me.txt备注.Text)
            .TextMatrix(.Row, mCol.默认) = Me.chk默认.Value
            If Me.chk默认.Value = 1 Then
                For lngRow = .FixedRows To .Rows - 1
                    If lngRow <> .Row And .TextMatrix(lngRow, mCol.标本) = .TextMatrix(.Row, mCol.标本) Then
                        If .TextMatrix(lngRow, mCol.默认) = 1 Then .TextMatrix(lngRow, mCol.默认) = 0
                    End If
                Next
            End If
            If Me.cbo科室.ListIndex = -1 Then
                .TextMatrix(.Row, mCol.申请科室Id) = 0: .TextMatrix(.Row, mCol.申请科室Id) = ""
            Else
                .TextMatrix(.Row, mCol.申请科室Id) = Me.cbo科室.ItemData(Me.cbo科室.ListIndex)
                .TextMatrix(.Row, mCol.申请科室) = Mid(Me.cbo科室.Text, 5)
            End If
            Call RowRefresh(.Row)
        End Select
    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_Edit_NewItem: Control.Enabled = (Me.Tag = "编辑")
    Case conMenu_Edit_Delete, conMenu_Edit_Adjust: Control.Enabled = (Me.Tag = "编辑" And Me.vfgList.Row >= Me.vfgList.FixedRows)
    End Select
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    '内部菜单工具栏定义
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set cbsThis.Icons = zlcommfun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
    End With
    Me.cbsThis.EnableCustomization False
    
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.Position = xtpBarBottom
    Me.cbsThis.ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    With Me.cbsThis.ActiveMenuBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加新行"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除本行"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "更新到参考值列表中"): cbrControl.Flags = xtpFlagRightAlign: cbrControl.Style = xtpButtonIconAndCaption
    End With

    '基本数据装入
    aryTemp = Split("0-无区分;1-男性;2-女性", ";")
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo性别.AddItem aryTemp(lngCount)
    Next
    Me.cbo性别.ListIndex = 0
    
    aryTemp = Split("1-年;2-月;3-日;4-小时", ";")
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo单位.AddItem aryTemp(lngCount)
    Next
    Me.cbo单位.ListIndex = 0
    
    Err = 0: On Error GoTo ErrHand
 
    gstrSql = "Select 编码,名称,适用性别 From 诊疗检验标本"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Do While Not rsTemp.EOF
        Me.cbo标本.AddItem rsTemp!编码 & "-" & rsTemp!名称
        If InStr(Trim("" & rsTemp!适用性别), "男") > 0 Then
            Me.cbo标本.ItemData(Me.cbo标本.NewIndex) = 1
        ElseIf InStr(Trim("" & rsTemp!适用性别), "女") > 0 Then
            Me.cbo标本.ItemData(Me.cbo标本.NewIndex) = 2
        Else
            Me.cbo标本.ItemData(Me.cbo标本.NewIndex) = 0
        End If
        rsTemp.MoveNext
    Loop
    If Me.cbo标本.ListCount > 0 Then Me.cbo标本.ListIndex = 0
        
    gstrSql = "Select 编码,名称 From 临床特征"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.cbo临床特征.AddItem "-": Me.cbo临床特征.ListIndex = 0
    Do While Not rsTemp.EOF
        Me.cbo临床特征.AddItem rsTemp!编码 & "-" & rsTemp!名称
        rsTemp.MoveNext
    Loop

    
    gstrSql = "Select ID, 编码, 名称 From 检验仪器"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.cbo仪器.AddItem "-": Me.cbo仪器.ItemData(Me.cbo仪器.NewIndex) = 0: Me.cbo仪器.ListIndex = 0
    Do While Not rsTemp.EOF
        Me.cbo仪器.AddItem rsTemp!编码 & "-" & rsTemp!名称
        Me.cbo仪器.ItemData(Me.cbo仪器.NewIndex) = Val("" & rsTemp!ID)
        rsTemp.MoveNext
    Loop
    
    
    gstrSql = "Select Distinct a.Id, a.编码, a.名称, b.服务对象" & vbNewLine & _
            "From 部门表 A, 部门性质说明 B" & vbNewLine & _
            "Where a.Id = b.部门id And b.工作性质 = '临床' Order by a.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.cbo科室.AddItem "-": Me.cbo科室.ItemData(Me.cbo科室.NewIndex) = 0: Me.cbo科室.ListIndex = 0
    Do While Not rsTemp.EOF
        Me.cbo科室.AddItem rsTemp!编码 & "-" & rsTemp!名称
        Me.cbo科室.ItemData(Me.cbo科室.NewIndex) = Val("" & rsTemp!ID)
        rsTemp.MoveNext
    Loop

    
    '列表设置
    Call setListFormat

    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.picEdit.Top = Me.ScaleHeight - Me.picEdit.Height - 180
    If Me.Tag = "编辑" Then
        Me.vfgList.Height = Me.picEdit.Top - Me.vfgList.Top
        Me.txt警戒下限.Enabled = True: Me.txt警戒上限.Enabled = True
        Me.txt变异报警率.Enabled = True: Me.txt比对警示率.Enabled = True: Me.txt比对失控率.Enabled = True
        Me.picEdit.Enabled = True: Me.picEdit.Visible = True: Me.txt变异警示率.Enabled = True
    Else
        Me.vfgList.Height = Me.picEdit.Top + Me.picEdit.Height - Me.vfgList.Top
        Me.txt警戒下限.Enabled = False: Me.txt警戒上限.Enabled = False
        Me.txt变异报警率.Enabled = False: Me.txt比对警示率.Enabled = False: Me.txt比对失控率.Enabled = False
        Me.picEdit.Enabled = False: Me.picEdit.Visible = False: Me.txt变异警示率.Enabled = False
    End If
End Sub

Private Sub txt备注_GotFocus()
    Me.txt备注.SelStart = 0: Me.txt备注.SelLength = 1000
    Call zlcommfun.OpenIme(True)
End Sub

Private Sub txt备注_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlcommfun.PressKey(vbKeyTab)
        '自动更新到列表中
        Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Adjust)
        If cbrControl Is Nothing Then Exit Sub
        If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
        Call cbsThis_Execute(cbrControl)
        Exit Sub
    End If
    If InStr(GCST_INVALIDCHAR & ";|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt比对警示率_GotFocus()
    Me.txt比对警示率.SelStart = 0: Me.txt比对警示率.SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt比对警示率_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr(".", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt比对失控率_GotFocus()
    Me.txt比对失控率.SelStart = 0: Me.txt比对失控率.SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt比对失控率_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr(".", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt变异报警率_GotFocus()
    Me.txt变异报警率.SelStart = 0: Me.txt变异报警率.SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt变异报警率_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr(".", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt变异警示率_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr(".", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt参考值_GotFocus(Index As Integer)
    Me.txt参考值(Index).SelStart = 0: Me.txt参考值(Index).SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt参考值_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr("-.", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt警戒上限_GotFocus()
    Me.txt警戒上限.SelStart = 0: Me.txt警戒上限.SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt警戒上限_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr("-.", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt警戒下限_GotFocus()
    Me.txt警戒下限.SelStart = 0: Me.txt警戒下限.SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt警戒下限_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr("-.", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt可偏移率_GotFocus()
    Me.txt可偏移率.SelStart = 0: Me.txt可偏移率.SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt可偏移率_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr(".", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt年龄_GotFocus(Index As Integer)
    Me.txt年龄(Index).SelStart = 0: Me.txt年龄(Index).SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt年龄_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub vfgList_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    With Me.vfgList
        If .Row = .Rows - 1 Then
            Call zlcommfun.PressKey(vbKeyTab)
        Else
            .Row = .Row + 1
        End If
    End With
End Sub

Private Sub vfgList_RowColChange()
    Dim lngList As Long
    With Me.vfgList
        Me.cbo标本.ListIndex = -1: Me.cbo性别.ListIndex = 0
        Me.txt年龄(0).Text = "": Me.txt年龄(1).Text = "": Me.cbo单位.ListIndex = -1
        Me.cbo临床特征.ListIndex = 0
        Me.txt参考值(0).Text = "": Me.txt参考值(1).Text = "": Me.txt可偏移率.Text = ""
        Me.cbo参考值.ListIndex = -1
        Me.cbo仪器.ListIndex = 0: Me.txt备注.Text = ""
        Me.cbo科室.ListIndex = 0
        If .Row = 0 Then Exit Sub
        
        For lngCount = 0 To Me.cbo标本.ListCount - 1
            If Mid(Me.cbo标本.List(lngCount), 4) = .TextMatrix(.Row, mCol.标本) Then Me.cbo标本.ListIndex = lngCount: Exit For
        Next
        
        For lngCount = 0 To Me.cbo性别.ListCount - 1
            If Val(Left(Me.cbo性别.List(lngCount), 1)) = Val(.TextMatrix(.Row, mCol.性别域)) Then Me.cbo性别.ListIndex = lngCount: Exit For
        Next
        
        Me.txt年龄(0).Text = .TextMatrix(.Row, mCol.年龄下限)
        Me.txt年龄(1).Text = .TextMatrix(.Row, mCol.年龄上限)
        For lngCount = 0 To Me.cbo单位.ListCount - 1
            If Mid(Me.cbo单位.List(lngCount), 3) = .TextMatrix(.Row, mCol.年龄单位) Then Me.cbo单位.ListIndex = lngCount: Exit For
        Next
        
        For lngCount = 1 To Me.cbo临床特征.ListCount - 1
            If Mid(Me.cbo临床特征.List(lngCount), 4) = .TextMatrix(.Row, mCol.临床特征) Then Me.cbo临床特征.ListIndex = lngCount: Exit For
        Next
        
        Select Case mInt类型
        Case 1, 2, 3
            Me.txt参考值(0).Text = FormatReference(mlngItemID, .TextMatrix(.Row, mCol.参考低值))
            Me.txt参考值(1).Text = FormatReference(mlngItemID, .TextMatrix(.Row, mCol.参考高值))
            
            Me.txt警戒下限.Text = FormatReference(mlngItemID, .TextMatrix(.Row, mCol.警示低值))
            Me.txt警戒上限.Text = FormatReference(mlngItemID, .TextMatrix(.Row, mCol.警示高值))
            Me.txt复查下限.Text = FormatReference(mlngItemID, .TextMatrix(.Row, mCol.复查低值))
            Me.txt复查上限.Text = FormatReference(mlngItemID, .TextMatrix(.Row, mCol.复查高值))
            
            Me.txt可偏移率.Text = .TextMatrix(.Row, mCol.可偏移率)
        Case 3
'            For lngCount = 0 To Me.cbo参考值.ListCount - 1
'                If lngCount = Val(.TextMatrix(.Row, mcol.参考低值)) Then Me.cbo参考值.ListIndex = lngCount: Exit For
'            Next
'            Me.txt可偏移率.Text = ""
        End Select
        
        For lngList = 1 To Me.cbo仪器.ListCount - 1
            If Me.cbo仪器.ItemData(lngList) = Val(.TextMatrix(.Row, mCol.仪器Id)) Then Me.cbo仪器.ListIndex = lngList: Exit For
        Next
        Me.txt备注.Text = .TextMatrix(.Row, mCol.备注)
        Me.chk默认.Value = Val(.TextMatrix(.Row, mCol.默认))
        
        For lngList = 1 To Me.cbo科室.ListCount - 1
            If Me.cbo科室.ItemData(lngList) = Val(.TextMatrix(.Row, mCol.申请科室Id)) Then Me.cbo科室.ListIndex = lngList: Exit For
        Next
    End With
End Sub

Private Function FormatReference(lngID As Long, strReference As String) As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    strSQL = "select max( D.小数位数)  as 小数位数 from   检验报告项目 a ,  诊疗项目目录 b , 诊治所见项目 c ,检验仪器项目 d " & _
                     " Where a.诊疗项目id = b.ID And a.报告项目id = c.ID And c.ID = d.项目id And d.小数位数 Is Not Null " & _
                     " and b.id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngItemID)
    
    If IsNull(rsTmp(0)) = False And rsTmp(0) > 0 Then
        strReference = Format(strReference, "0." & Replace(Space(rsTmp(0)), " ", "0"))
    End If
    FormatReference = strReference
End Function
