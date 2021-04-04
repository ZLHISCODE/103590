VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClientsSever 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "客户端升级文件服务器配置"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6390
   Icon            =   "frmClientsSever.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vsfUse 
      Height          =   3945
      Left            =   45
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   6285
      _cx             =   11086
      _cy             =   6959
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClientsSever.frx":6852
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPart 
      Height          =   3945
      Left            =   45
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   6285
      _cx             =   11086
      _cy             =   6959
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClientsSever.frx":68C3
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
   Begin VB.Frame fraBounds 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   20
      Top             =   990
      Width           =   7365
   End
   Begin VB.Frame fraBounds 
      Height          =   45
      Index           =   0
      Left            =   -645
      TabIndex        =   19
      Top             =   6060
      Width           =   7365
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   6390
      TabIndex        =   15
      Top             =   0
      Width           =   6390
      Begin VB.Image imgCaption 
         Height          =   720
         Left            =   525
         Picture         =   "frmClientsSever.frx":6934
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblEXP 
         BackStyle       =   0  'Transparent
         Caption         =   "全选：双击最左侧表头全选，再次双击反选"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   1650
         TabIndex        =   18
         Top             =   405
         Width           =   3780
      End
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "批量设置：只对勾选项目生效"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1650
         TabIndex        =   17
         Top             =   135
         Width           =   2340
      End
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应用：只对按院区或按部门或按用途生效方案生效"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   1650
         TabIndex        =   16
         Top             =   675
         Width           =   3960
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfNode 
      Height          =   3945
      Left            =   45
      TabIndex        =   8
      Top             =   2040
      Width           =   6285
      _cx             =   11086
      _cy             =   6959
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClientsSever.frx":8476
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
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   510
      ScaleHeight     =   240
      ScaleWidth      =   2565
      TabIndex        =   12
      Top             =   1665
      Width           =   2600
      Begin VB.TextBox txtFind 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   45
         TabIndex        =   3
         Text            =   "请输入查找内容"
         Top             =   30
         Width           =   2500
      End
   End
   Begin VB.PictureBox picSeverSet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   275
      Left            =   3225
      ScaleHeight     =   240
      ScaleWidth      =   3075
      TabIndex        =   11
      Top             =   1665
      Width           =   3100
      Begin VB.CommandButton cmdOK 
         Caption         =   "批量设置"
         Height          =   300
         Left            =   2205
         TabIndex        =   5
         ToolTipText     =   "对表格中打钩的选项批量选择服务器"
         Top             =   -30
         Width           =   900
      End
      Begin VB.ComboBox cboSever 
         Height          =   300
         Left            =   -30
         TabIndex        =   4
         Top             =   -30
         Width           =   2285
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "取消(&Q)"
      Height          =   300
      Left            =   5145
      TabIndex        =   7
      Top             =   6255
      Width           =   1200
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "应用(&S)"
      Height          =   300
      Left            =   3870
      TabIndex        =   6
      ToolTipText     =   "应用只对按院区或按部门方案生效"
      Top             =   6255
      Width           =   1200
   End
   Begin VB.CheckBox chkSetClients 
      Caption         =   "对未设置升级服务器的客户端生效"
      Height          =   315
      Left            =   75
      TabIndex        =   2
      Top             =   6255
      Value           =   1  'Checked
      Width           =   3150
   End
   Begin VB.PictureBox picType 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   915
      ScaleHeight     =   435
      ScaleWidth      =   3930
      TabIndex        =   10
      Top             =   1125
      Width           =   3930
      Begin VB.OptionButton optType 
         Caption         =   "按用途"
         Height          =   240
         Index           =   2
         Left            =   2235
         TabIndex        =   21
         Top             =   120
         Width           =   1065
      End
      Begin VB.OptionButton optType 
         Caption         =   "按部门"
         Height          =   240
         Index           =   1
         Left            =   1140
         TabIndex        =   1
         Top             =   120
         Width           =   1065
      End
      Begin VB.OptionButton optType 
         Caption         =   "按院区"
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   0
         Top             =   120
         Width           =   1065
      End
   End
   Begin MSComctlLib.ImageList imgEdit 
      Left            =   5730
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsSever.frx":84DC
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsSever.frx":8A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsSever.frx":9010
            Key             =   "签名"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsSever.frx":9362
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsSever.frx":FBC4
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsSever.frx":16426
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientsSever.frx":168EE
            Key             =   "AllCheck"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFind 
      Caption         =   "查找"
      Height          =   240
      Left            =   75
      TabIndex        =   14
      Top             =   1710
      Width           =   405
   End
   Begin VB.Label lblType 
      Caption         =   "配置方式"
      Height          =   180
      Left            =   90
      TabIndex        =   13
      Top             =   1260
      Width           =   795
   End
End
Attribute VB_Name = "frmClientsSever"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnSetSuccese As Boolean

Private Enum SeverData
    Col_选择 = 0
    Col_站点 = 1 '状态值 0-正常 1-升级部件缺失(本地文件必不存在) 2-本地文件不存在 3-无需更新 4-警告但可以上传 5-已经上传
    Col_部门 = 1
    Col_用途 = 1
    Col_升级服务器 = 2
    Col_列数 = 3
End Enum

Public Sub LoadNodeData(Optional ByVal strFilter As String, Optional ByVal strLocationName As String)
    Dim i, j As Long
    Dim strSQL       As String
    Dim blnTemp As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errH

    With vsfNode
        .Redraw = flexRDNone
        .Rows = .FixedRows
'        .Cols = Col_列数

'        .Cell(flexcpText, 0, Col_站点) = "院区"
'        .Cell(flexcpAlignment, 0, Col_站点) = flexAlignCenterCenter
'
'        .Cell(flexcpText, 0, Col_升级服务器) = "升级服务器"
'        .Cell(flexcpAlignment, 0, Col_升级服务器) = flexAlignCenterCenter

        .Cell(flexcpPicture, 0, Col_选择) = imgEdit.ListImages("UnCheck").Picture
        .Cell(flexcpPictureAlignment, 0, Col_选择) = flexAlignCenterCenter
    
'        strSQL = "select 名称 as 站点,位置 as 服务器地址 from zltools.zlnodelist A,zltools.zlupgradeserver B where A.编号 = B.编号(+)"
        strSQL = "select 名称 as 站点,编号 as 编号 from zltools.zlnodelist"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)

        '数据填入
        .Rows = rsTemp.RecordCount + 2
        .Cell(flexcpText, 1, Col_部门) = "[客户端无对应站点]"
        .Cell(flexcpAlignment, 1, Col_部门) = flexAlignLeftCenter
        i = 2
        Do Until rsTemp.EOF
            .Cell(flexcpText, i, Col_站点) = Nvl(rsTemp.Fields("站点"), "")
            .Cell(flexcpData, i, Col_站点) = Nvl(rsTemp.Fields("编号"), "")
            .Cell(flexcpAlignment, i, Col_站点) = flexAlignCenterCenter

'            .Cell(flexcpText, i, Col_升级服务器) = Nvl(rsTemp.Fields("服务器地址"), "")
            .Cell(flexcpText, i, Col_升级服务器) = " "
            .Cell(flexcpAlignment, i, Col_升级服务器) = flexAlignLeftCenter
  
            rsTemp.MoveNext
            i = i + 1
        Loop
        
        LoadSever Me.vsfNode
        
        .Editable = flexEDKbdMouse
        '选中框风格
        .FocusRect = flexFocusSolid
        '最后一列自动列宽
        .ExtendLastCol = True
        '滚动画面跟随
        .ScrollTrack = True
        '自动换行
        .WordWrap = True
        '行高设置
        .RowHeightMin = 300
        .RowHeightMax = 300
        '最大宽度设置
        .ColWidthMax = 7000
        '自动适应行高、列宽
        .AutoSizeMode = flexAutoSizeRowHeight
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
'        .Redraw = flexRDBuffered
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
        .Redraw = flexRDBuffered
'        Call SetMenu

    End With
    Exit Sub
errH:
    Call MsgBox("服务器列表加载错误", vbInformation, gstrSysName)
    vsfNode.Clear
    If False Then
        Resume
    End If
End Sub

Public Sub LoadPartData(Optional ByVal strFilter As String, Optional ByVal strLocationName As String)
    Dim i, j As Long
    Dim strSQL       As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errH

    With vsfPart
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Cols = Col_列数

'        .Cell(flexcpText, 0, Col_部门) = "部门"
'        .Cell(flexcpAlignment, 0, Col_部门) = flexAlignCenterCenter
'
'        .Cell(flexcpText, 0, Col_升级服务器) = "升级服务器"
'        .Cell(flexcpAlignment, 0, Col_升级服务器) = flexAlignCenterCenter

        .Cell(flexcpPicture, 0, Col_选择) = imgEdit.ListImages("UnCheck").Picture
        .Cell(flexcpPictureAlignment, 0, Col_选择) = flexAlignCenterCenter
        
        strSQL = "select distinct 部门 from zlclients where not 部门 is null order by 部门"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        
        '数据填入
        .Rows = rsTemp.RecordCount + 2
        
        .Cell(flexcpText, 1, Col_部门) = "[客户端无对应部门]"
        .Cell(flexcpAlignment, 1, Col_部门) = flexAlignLeftCenter
        .Cell(flexcpForeColor, 1, Col_部门) = vbBlue
    
        i = 2
        Do Until rsTemp.EOF
            .Cell(flexcpText, i, Col_部门) = Nvl(rsTemp.Fields("部门"), "")
            .Cell(flexcpAlignment, i, Col_部门) = flexAlignLeftCenter

            .Cell(flexcpText, i, Col_升级服务器) = " "
            .Cell(flexcpAlignment, i, Col_升级服务器) = flexAlignLeftCenter
  
            rsTemp.MoveNext
            i = i + 1
        Loop
        
        LoadSever Me.vsfPart
        
        .Editable = flexEDKbdMouse
        '选中框风格
        .FocusRect = flexFocusSolid
        '最后一列自动列宽
        .ExtendLastCol = True
        '滚动画面跟随
        .ScrollTrack = True
        '自动换行
        .WordWrap = True
        '行高设置
        .RowHeightMin = 300
        .RowHeightMax = 300
        '最大宽度设置
        .ColWidthMax = 7000
        '自动适应行高、列宽
        .AutoSizeMode = flexAutoSizeRowHeight
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
'        .Redraw = flexRDBuffered
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
        .Redraw = flexRDBuffered
'        Call SetMenu

    End With
    Exit Sub
errH:
    Call MsgBox("服务器列表加载错误", vbInformation, gstrSysName)
    vsfNode.Clear
    If False Then
        Resume
    End If
End Sub

Public Sub LoadUseData(Optional ByVal strFilter As String, Optional ByVal strLocationName As String)
    Dim i, j As Long
    Dim strSQL       As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errH

    With vsfUse
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Cols = Col_列数

        .Cell(flexcpPicture, 0, Col_选择) = imgEdit.ListImages("UnCheck").Picture
        .Cell(flexcpPictureAlignment, 0, Col_选择) = flexAlignCenterCenter
        
        strSQL = "select distinct 用途 from zlclients where not 用途 is null order by 用途"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        
        '数据填入
        .Rows = rsTemp.RecordCount + 2
        
        .Cell(flexcpText, 1, Col_用途) = "[客户端无对应用途]"
        .Cell(flexcpAlignment, 1, Col_用途) = flexAlignLeftCenter
        .Cell(flexcpForeColor, 1, Col_用途) = vbBlue
        
        i = 2
        Do Until rsTemp.EOF
            .Cell(flexcpText, i, Col_用途) = Nvl(rsTemp.Fields("用途"), "")
            .Cell(flexcpAlignment, i, Col_用途) = flexAlignLeftCenter

            .Cell(flexcpText, i, Col_升级服务器) = " "
            .Cell(flexcpAlignment, i, Col_升级服务器) = flexAlignLeftCenter
  
            rsTemp.MoveNext
            i = i + 1
        Loop
        
        LoadSever Me.vsfUse
        
        .Editable = flexEDKbdMouse
        '选中框风格
        .FocusRect = flexFocusSolid
        '最后一列自动列宽
        .ExtendLastCol = True
        '滚动画面跟随
        .ScrollTrack = True
        '自动换行
        .WordWrap = True
        '行高设置
        .RowHeightMin = 300
        .RowHeightMax = 300
        '最大宽度设置
        .ColWidthMax = 7000
        '自动适应行高、列宽
        .AutoSizeMode = flexAutoSizeRowHeight
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
'        .Redraw = flexRDBuffered
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
        .Redraw = flexRDBuffered
'        Call SetMenu

    End With
    Exit Sub
errH:
    Call MsgBox("服务器列表加载错误", vbInformation, gstrSysName)
    vsfNode.Clear
    If False Then
        Resume
    End If
End Sub

Public Function ShowMe(frmParent As Object) As Boolean
    Me.Show 1, frmParent
    ShowMe = blnSetSuccese
    Unload Me
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim objVsf As Object
    
    If optType.Item(0).value = True Then
        Set objVsf = Me.vsfNode
    ElseIf optType.Item(1).value = True Then
        Set objVsf = Me.vsfPart
    ElseIf optType.Item(2).value = True Then
        Set objVsf = Me.vsfUse
    End If
    
    With objVsf
        If .Rows < .FixedRows Then Exit Sub
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, Col_选择) = True Then
                .Cell(flexcpText, i, Col_升级服务器) = cboSever.Text
            End If
        Next
    End With
End Sub

Private Sub cmdSet_Click()
        If optType.Item(0).value = True Then
            SaveSeverSet 0 '按站点设置
        ElseIf optType.Item(1).value = True Then
            SaveSeverSet 1 '按部门设置
        ElseIf optType.Item(2).value = True Then
            SaveSeverSet 2
        End If
        SaveConfigure
End Sub

Private Sub Form_Load()
    LoadNodeData
    LoadPartData
    LoadUseData
    LoadSetting
End Sub

Private Sub OptType_Click(Index As Integer)
    Select Case Index
        Case 0
'            LoadNodeData
            vsfNode.Visible = True
            vsfPart.Visible = False
            vsfUse.Visible = False
            txtFind.Tag = "请输入院区名称查找"
            txtFind.Text = txtFind.Tag
            txtFind.ForeColor = vbGrayText
            SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "显示表格", "0"
        Case 1
'            LoadPartData
            vsfNode.Visible = False
            vsfPart.Visible = True
            vsfUse.Visible = False
            txtFind.Tag = "请输入部门名称查找"
            txtFind.Text = txtFind.Tag
            txtFind.ForeColor = vbGrayText
            SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "显示表格", "1"
        Case 2
            vsfNode.Visible = False
            vsfPart.Visible = False
            vsfUse.Visible = True
            txtFind.Tag = "请输入用途名称查找"
            txtFind.Text = txtFind.Tag
            txtFind.ForeColor = vbGrayText
            SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "显示表格", "2"
    End Select
End Sub

Private Sub txtFind_Change()
    Dim i As Long
    Dim objVsf As Object
        
    If txtFind.Text = txtFind.Tag Then Exit Sub
    
    If optType.Item(0).value = True Then
        Set objVsf = Me.vsfNode
    ElseIf optType.Item(1).value = True Then
        Set objVsf = Me.vsfPart
    ElseIf optType.Item(2).value = True Then
        Set objVsf = Me.vsfUse
    End If

    With objVsf
        If .Rows < .FixedRows Then Exit Sub
        For i = .FixedRows To .Rows - 1
            If InStr(.TextMatrix(i, Col_部门), txtFind.Text) > 0 Then
                .RowHidden(i) = False
            Else
                .RowHidden(i) = True
            End If
        Next
    End With
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text = txtFind.Tag Then
        txtFind.Text = ""
        txtFind.ForeColor = vbBlack
    End If
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
'111
End Sub

Private Sub txtFind_LostFocus()
    If txtFind.Text = "" Then
        txtFind.Text = txtFind.Tag
        txtFind.ForeColor = vbGrayText
    End If
End Sub

Private Sub vsfNode_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     If Col = Col_部门 Or Col = Col_站点 Then Cancel = True
End Sub

Private Sub vsfNode_ChangeEdit()
    With vsfNode
'       Select Case .ComboItem(.ComboIndex)
'
'       End Select
    End With
End Sub

Private Sub LoadSetting()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngFindRow As Long
    Dim strTemp As String
    Dim intOptIndex As Integer
    
    On Error GoTo errH:
    
    '设置按站点服务器配置
    strSQL = "select 内容 from zltools.zlreginfo where 项目 = '客户端服务器-站点'"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        If strTemp <> "" Then
            strTemp = strTemp & "," & rsTemp.Fields("内容")
        Else
            strTemp = rsTemp.Fields("内容")
        End If
        rsTemp.MoveNext
    Loop
    
    strSQL = "Select a.C1 As 院区, a.C2 As 服务器编号, b.位置 As 服务器地址" & vbNewLine & _
                  "From Table(f_Str2list2('" & strTemp & "')) A, Zltools.Zlupgradeserver B" & vbNewLine & _
                  "Where a.C2 = b.编号(+)"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    
    With vsfNode
        Do Until rsTemp.EOF
            If IsNull(rsTemp.Fields("服务器编号")) = False Then
                lngFindRow = .FindRow(rsTemp.Fields("院区"), , Col_站点)
                If lngFindRow > 0 And Nvl(rsTemp.Fields("服务器地址"), "") <> "" Then
                    .TextMatrix(lngFindRow, Col_升级服务器) = Nvl(rsTemp.Fields("服务器编号"), "") & ":" & Nvl(rsTemp.Fields("服务器地址"), "")
                End If
            End If
            rsTemp.MoveNext
        Loop
    End With
    
    '设置按部门服务器配置
    strTemp = ""
    strSQL = "select 内容 from zltools.zlreginfo where 项目 = '客户端服务器-部门'"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        If strTemp <> "" Then
            strTemp = strTemp & "," & rsTemp.Fields("内容")
        Else
            strTemp = rsTemp.Fields("内容")
        End If
        rsTemp.MoveNext
    Loop
    
    strSQL = "Select a.C1 As 部门, a.C2 As 服务器编号, b.位置 As 服务器地址" & vbNewLine & _
                  "From Table(f_Str2list2('" & strTemp & "')) A, Zltools.Zlupgradeserver B" & vbNewLine & _
                  "Where a.C2 = b.编号(+)"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    
    With vsfPart
        Do Until rsTemp.EOF
            If IsNull(rsTemp.Fields("服务器编号")) = False Then
                lngFindRow = .FindRow(rsTemp.Fields("部门"), , Col_部门)
                If lngFindRow > 0 And Nvl(rsTemp.Fields("服务器地址"), "") <> "" Then
                    .TextMatrix(lngFindRow, Col_升级服务器) = Nvl(rsTemp.Fields("服务器编号"), "") & ":" & Nvl(rsTemp.Fields("服务器地址"), "")
                End If
            End If
            rsTemp.MoveNext
        Loop
    End With
    
     '设置按用途服务器配置
    strTemp = ""
    strSQL = "select 内容 from zltools.zlreginfo where 项目 = '客户端服务器-用途'"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        If strTemp <> "" Then
            strTemp = strTemp & "," & rsTemp.Fields("内容")
        Else
            strTemp = rsTemp.Fields("内容")
        End If
        rsTemp.MoveNext
    Loop
    
    strSQL = "Select a.C1 As 用途, a.C2 As 服务器编号, b.位置 As 服务器地址" & vbNewLine & _
                  "From Table(f_Str2list2('" & strTemp & "')) A, Zltools.Zlupgradeserver B" & vbNewLine & _
                  "Where a.C2 = b.编号(+)"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    
    With vsfUse
        Do Until rsTemp.EOF
            If IsNull(rsTemp.Fields("服务器编号")) = False Then
                lngFindRow = .FindRow(rsTemp.Fields("用途"), , Col_用途)
                If lngFindRow > 0 And Nvl(rsTemp.Fields("服务器地址"), "") <> "" Then
                    .TextMatrix(lngFindRow, Col_升级服务器) = Nvl(rsTemp.Fields("服务器编号"), "") & ":" & Nvl(rsTemp.Fields("服务器地址"), "")
                End If
            End If
            rsTemp.MoveNext
        Loop
    End With
    
    '设置界面参数
    strSQL = "select 内容 from zltools.zlreginfo where 项目 = '客户端服务器-配置'"
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    
    If Not rsTemp.EOF Then
        chkSetClients.value = Nvl(rsTemp.Fields("内容"), 1)
    End If
    
    '界面设置
    strSQL = "select distinct 站点 from zltools.zlclients where not 站点 is null "
    Call OpenRecordset(rsTemp, strSQL, Me.Caption)
    
    If rsTemp.EOF = True Then '没有站点数据，只显示部门、用途设置
        optType.Item(0).Enabled = False
    End If
    intOptIndex = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "显示表格", ""))
    optType.Item(intOptIndex).value = True
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

Private Sub LoadSever(ByRef objVsf As VSFlexGrid)
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    Dim blnTemp As Boolean
    
    On Error GoTo errH:
    
    cboSever.Clear
    With objVsf
        '下拉列表服务器加载
        strSQL = "select 编号,类型,位置,是否升级,是否缺省,是否收集 from zltools.zlupgradeserver order by 编号"
        Call OpenRecordset(rsTemp, strSQL, Me.Caption)
        
        .ColComboList(Col_升级服务器) = "#0*2;" & "" & vbTab & "" & vbTab & " "
        cboSever.AddItem " ", 0
        
        i = 1
        Do Until rsTemp.EOF
            If Nvl(rsTemp.Fields("是否升级"), "0") = "1" Or Nvl(rsTemp.Fields("是否缺省"), "0") = "1" Or Nvl(rsTemp.Fields("是否收集"), "0") = "1" Then
                blnTemp = True
            Else
                blnTemp = False
            End If
            If blnTemp Then
                .ColComboList(Col_升级服务器) = .ColComboList(Col_升级服务器) & _
                "|#" & i & ";" & rsTemp.Fields("编号") & "号" & vbTab & IIf(rsTemp.Fields("类型") = 0, "共享", "FTP") & vbTab & rsTemp.Fields("编号") & ":" & rsTemp.Fields("位置")
                cboSever.AddItem rsTemp.Fields("编号") & ":" & rsTemp.Fields("位置"), i
                i = i + 1
            End If
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

Private Function SaveSeverSet(intType As Integer)
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    Dim blnTemp As Boolean
    Dim strTemp As String
    Dim strSave() As String
    Dim strSettingNode As String
    Dim strSettingPart As String
    
    On Error GoTo errH:
    Select Case intType
        Case 0 '按站点设置
            With vsfNode
                If .Rows < 1 Then SaveSeverSet = False: Exit Function
                For i = 1 To .Rows - 1
                    strTemp = Trim(.Cell(flexcpTextDisplay, i, Col_升级服务器))
                    If strTemp <> "" Then
                        strSave = Split(strTemp, ":")
                        If IsNumeric(strSave(0)) = True Then
                            If .TextMatrix(i, Col_部门) = "[客户端无对应部门]" Then '缺省站点客户端设置升级服务器
                                strSQL = "update zltools.zlclients set 升级文件服务器 = " & Trim(strSave(0)) & " where 站点 is null"
                            Else
                                strSQL = "update zltools.zlclients set 升级文件服务器 = " & Trim(strSave(0)) & " where 站点 = '" & .Cell(flexcpData, i, Col_站点) & "'"
                            End If
                            If chkSetClients.value = 1 Then '对该站点未设置过升级服务器的客户端设置
                                strSQL = strSQL & " and (升级文件服务器 is null or 升级文件服务器 = 0)"
                            End If
                            gcnOracle.Execute strSQL
                        End If
                    End If
                Next
            End With
        Case 1 '按部门设置
            With vsfPart
                For i = 1 To .Rows - 1
                    strTemp = Trim(.Cell(flexcpTextDisplay, i, Col_升级服务器))
                    If strTemp <> "" Then
                        strSave = Split(strTemp, ":")
                        If IsNumeric(strSave(0)) = True Then
                            If .TextMatrix(i, Col_部门) = "[客户端无对应部门]" Then '缺省部门客户端设置升级服务器
                                strSQL = "update zltools.zlclients set 升级文件服务器 = " & Trim(strSave(0)) & " where 部门 is null"
                            Else
                                strSQL = "update zltools.zlclients set 升级文件服务器 = " & Trim(strSave(0)) & " where 部门 = '" & .TextMatrix(i, Col_部门) & "'"
                            End If
                            If chkSetClients.value = 1 Then '对该部门未设置过升级服务器的客户端设置
                                strSQL = strSQL & " and (升级文件服务器 is null or 升级文件服务器 = 0)"
                            End If
                            gcnOracle.Execute strSQL
                        End If
                    End If
                Next
            End With
        Case 2 '按用途设置
            With vsfUse
                For i = 1 To .Rows - 1
                    strTemp = Trim(.Cell(flexcpTextDisplay, i, Col_升级服务器))
                    If strTemp <> "" Then
                        strSave = Split(strTemp, ":")
                        If IsNumeric(strSave(0)) = True Then
                            If .TextMatrix(i, Col_部门) = "[客户端无对应用途]" Then '缺省部门客户端设置升级服务器
                                strSQL = "update zltools.zlclients set 升级文件服务器 = " & Trim(strSave(0)) & " where 用途 is null"
                            Else
                                strSQL = "update zltools.zlclients set 升级文件服务器 = " & Trim(strSave(0)) & " where 用途 = '" & .TextMatrix(i, Col_用途) & "'"
                            End If
                            If chkSetClients.value = 1 Then '对该部门未设置过升级服务器的客户端设置
                                strSQL = strSQL & " and (升级文件服务器 is null or 升级文件服务器 = 0)"
                            End If
                            gcnOracle.Execute strSQL
                        End If
                    End If
                Next
            End With
    End Select
    
    MsgBox "按" & Decode(intType, 0, " 站点 ", 1, " 部门 ", 2, "用途", "") & "设置完成！", vbInformation, gstrSysName
    blnSetSuccese = True
    SaveSeverSet = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    SaveSeverSet = False
    If False Then
        Resume
    End If
End Function

Private Function SaveConfigure() As Boolean
    Dim arrSQL() As Variant
    Dim i, lngCounts As Long
    Dim blnTrans As Boolean
    Dim strTemp As String
    Dim strSettingNode As String
    Dim strSettingPart As String
    
    '存储配置
    On Error GoTo errH:

    arrSQL() = Array()
    Me.Enabled = False
    
    With vsfNode
        If .Rows < 1 Then Exit Function
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - 1
            strTemp = .TextMatrix(i, Col_站点) & ":"
            If Trim(.Cell(flexcpTextDisplay, i, Col_升级服务器)) <> "" Then strTemp = strTemp & Split(Trim(.Cell(flexcpTextDisplay, i, Col_升级服务器)), ":")(0)
            If ActualLen(strSettingNode & strTemp) > 3900 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "b_Public.Zlreginfoupdate('客户端服务器-站点'," & UBound(arrSQL) + 1 & ",'" & strSettingNode & "',null)"
                    
                strSettingNode = strTemp
            Else
                If strSettingNode <> "" Then
                    strSettingNode = strSettingNode & "," & strTemp
                Else
                    strSettingNode = strTemp
                End If
            End If
        Next
        .Redraw = flexRDBuffered
        If strSettingNode <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "b_Public.Zlreginfoupdate('客户端服务器-站点'," & UBound(arrSQL) + 1 & ",'" & strSettingNode & "',null)"
        End If
    End With

    lngCounts = UBound(arrSQL)
    
    With vsfPart
        If .Rows < 1 Then Exit Function
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - 1
            strTemp = .TextMatrix(i, Col_部门) & ":"
            If Trim(.Cell(flexcpTextDisplay, i, Col_升级服务器)) <> "" Then strTemp = strTemp & Split(Trim(.Cell(flexcpTextDisplay, i, Col_升级服务器)), ":")(0)
            If ActualLen(strSettingNode & strTemp) > 3900 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "b_Public.Zlreginfoupdate('客户端服务器-部门'," & UBound(arrSQL) - lngCounts & ",'" & strSettingPart & "',null)"
                    
                strSettingPart = strTemp
            Else
                If strSettingPart <> "" Then
                    strSettingPart = strSettingPart & "," & strTemp
                Else
                    strSettingPart = strTemp
                End If
            End If
        Next
        .Redraw = flexRDBuffered
        If strSettingPart <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "b_Public.Zlreginfoupdate('客户端服务器-部门'," & UBound(arrSQL) - lngCounts & ",'" & strSettingPart & "',null)"
        End If
    End With

    lngCounts = UBound(arrSQL)
    
    With vsfUse
        If .Rows < 1 Then Exit Function
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - 1
            strTemp = .TextMatrix(i, Col_用途) & ":"
            If Trim(.Cell(flexcpTextDisplay, i, Col_升级服务器)) <> "" Then strTemp = strTemp & Split(Trim(.Cell(flexcpTextDisplay, i, Col_升级服务器)), ":")(0)
            If ActualLen(strSettingNode & strTemp) > 3900 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "b_Public.Zlreginfoupdate('客户端服务器-用途'," & UBound(arrSQL) - lngCounts & ",'" & strSettingPart & "',null)"
                    
                strSettingPart = strTemp
            Else
                If strSettingPart <> "" Then
                    strSettingPart = strSettingPart & "," & strTemp
                Else
                    strSettingPart = strTemp
                End If
            End If
        Next
        .Redraw = flexRDBuffered
        If strSettingPart <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "b_Public.Zlreginfoupdate('客户端服务器-用途'," & UBound(arrSQL) - lngCounts & ",'" & strSettingPart & "',null)"
        End If
    End With
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "b_Public.Zlreginfoupdate('客户端服务器-配置'," & 1 & ",'" & chkSetClients.value & "',null)"
    
    gcnOracle.BeginTrans: blnTrans = True
    For i = LBound(arrSQL) To UBound(arrSQL)
'        gcnOracle.Execute arrSQL(i)
'        arrSQL(i) = Replace(arrSQL(i), ")", ",'')")
        strTemp = arrSQL(i)
        Call ExecuteProcedure(strTemp, Me.Caption, gcnOracle)
    Next
    gcnOracle.CommitTrans: blnTrans = False

    Me.Enabled = True
    SaveConfigure = True
    
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    SaveConfigure = False
    MsgBox err.Description, vbExclamation, gstrSysName
    If 1 = 0 Then
        Resume
    End If
End Function

Private Sub vsfNode_DblClick()
    Dim i As Long
    Dim blnSelect As Boolean
    
    With vsfNode
        If .MouseRow = 0 And .MouseCol = Col_选择 Then
            If .Rows < .FixedRows Then Exit Sub
            
            If .Cell(flexcpPicture, 0, Col_选择) = imgEdit.ListImages("AllCheck").Picture Then
                .Cell(flexcpPicture, 0, Col_选择) = imgEdit.ListImages("UnCheck").Picture
                blnSelect = False
            Else
                .Cell(flexcpPicture, 0, Col_选择) = imgEdit.ListImages("AllCheck").Picture
                blnSelect = True
            End If
            
            For i = .FixedRows To .Rows - 1
                If Not .RowHidden(i) Then .TextMatrix(i, Col_选择) = blnSelect
            Next
        End If
     End With
End Sub

Private Sub vsfPart_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = Col_部门 Or Col = Col_站点 Then Cancel = True
End Sub


Private Sub vsfPart_DblClick()
    Dim i As Long
    Dim blnSelect As Boolean
    
    With vsfPart
        If .MouseRow = 0 And .MouseCol = Col_选择 Then
            If .Rows < .FixedRows Then Exit Sub
            
            If .Cell(flexcpPicture, 0, Col_选择) = imgEdit.ListImages("AllCheck").Picture Then
                .Cell(flexcpPicture, 0, Col_选择) = imgEdit.ListImages("UnCheck").Picture
                blnSelect = False
            Else
                .Cell(flexcpPicture, 0, Col_选择) = imgEdit.ListImages("AllCheck").Picture
                blnSelect = True
            End If
            
            For i = .FixedRows To .Rows - 1
                If Not .RowHidden(i) Then .TextMatrix(i, Col_选择) = blnSelect
            Next
        End If
     End With
End Sub

Private Sub vsfUse_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = Col_部门 Or Col = Col_站点 Then Cancel = True
End Sub

Private Sub vsfUse_DblClick()
    Dim i As Long
    Dim blnSelect As Boolean
    
    With vsfUse
        If .MouseRow = 0 And .MouseCol = Col_选择 Then
            If .Rows < .FixedRows Then Exit Sub
            
            If .Cell(flexcpPicture, 0, Col_选择) = imgEdit.ListImages("AllCheck").Picture Then
                .Cell(flexcpPicture, 0, Col_选择) = imgEdit.ListImages("UnCheck").Picture
                blnSelect = False
            Else
                .Cell(flexcpPicture, 0, Col_选择) = imgEdit.ListImages("AllCheck").Picture
                blnSelect = True
            End If
            
            For i = .FixedRows To .Rows - 1
                If Not .RowHidden(i) Then .TextMatrix(i, Col_选择) = blnSelect
            Next
        End If
     End With
End Sub
