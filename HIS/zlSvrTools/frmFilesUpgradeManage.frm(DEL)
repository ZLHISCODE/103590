VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmFilesUpgradeManage 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   ClientHeight    =   6612
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6612
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRepairFiles 
      Caption         =   "服务器文件检查(&T)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7755
      TabIndex        =   4
      ToolTipText     =   "检查服务器已上传的文件并自动修复"
      Top             =   180
      Width           =   1875
   End
   Begin VB.CommandButton cmdRepairList 
      Caption         =   "在用文件清单修正(&X)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5760
      TabIndex        =   3
      ToolTipText     =   "修正还原在用文件清单，修正为标准清单"
      Top             =   180
      Width           =   1875
   End
   Begin VB.CommandButton cmdUpLoad 
      Caption         =   "升级文件上传(&R)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3765
      TabIndex        =   2
      ToolTipText     =   "上传文件至已设置好可链接的服务器"
      Top             =   180
      Width           =   1600
   End
   Begin VB.PictureBox picBtn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1305
      ScaleHeight     =   348
      ScaleWidth      =   6780
      TabIndex        =   16
      Top             =   6060
      Width           =   6780
      Begin VB.CheckBox chkFilter 
         Caption         =   "只显示第三方部件"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4185
         TabIndex        =   14
         Top             =   30
         Width           =   2055
      End
      Begin VB.CommandButton cmdExpired 
         Caption         =   "弃用(&Q)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3015
         TabIndex        =   13
         ToolTipText     =   "弃用在用文件清单中的某个第三方文件并加入弃用文件清单"
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除(&D)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2010
         TabIndex        =   12
         ToolTipText     =   "删除在用文件清单中的某个第三方文件信息"
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "修改(&E)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1005
         TabIndex        =   11
         ToolTipText     =   "修改在用文件清单中的第三方文件信息"
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "增加(&A)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   10
         ToolTipText     =   "新增第三方文件至在用文件清单"
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   570
      ScaleHeight     =   264
      ScaleWidth      =   2676
      TabIndex        =   15
      Top             =   195
      Width           =   2700
      Begin VB.TextBox txtFind 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   45
         TabIndex        =   1
         Top             =   30
         Width           =   2580
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfFileList 
      Height          =   1005
      Left            =   600
      TabIndex        =   7
      Top             =   1905
      Width           =   3870
      _cx             =   6826
      _cy             =   1764
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483644
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483638
      BackColorAlternate=   16777215
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
      SelectionMode   =   2
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFilesUpgradeManage.frx":0000
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
      Ellipsis        =   1
      ExplorerBar     =   7
      PicturesOver    =   -1  'True
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
   Begin VSFlex8Ctl.VSFlexGrid vsfExpFileList 
      Height          =   1005
      Left            =   600
      TabIndex        =   8
      Top             =   2910
      Visible         =   0   'False
      Width           =   3870
      _cx             =   6826
      _cy             =   1764
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483644
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483638
      BackColorAlternate=   16777215
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
      SelectionMode   =   2
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFilesUpgradeManage.frx":0161
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
      Ellipsis        =   1
      ExplorerBar     =   7
      PicturesOver    =   -1  'True
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
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "第三方部件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   9
      Top             =   6120
      Width           =   900
   End
   Begin VB.Label LblBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "弃用文件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   1
      Left            =   5415
      TabIndex        =   6
      ToolTipText     =   "已经弃用的文件"
      Top             =   1100
      Width           =   855
   End
   Begin VB.Label LblBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "在用文件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   4575
      MousePointer    =   4  'Icon
      TabIndex        =   5
      ToolTipText     =   "正在使用的文件"
      Top             =   1100
      Width           =   855
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查找"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   165
      TabIndex        =   0
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "frmFilesUpgradeManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Const StopColor = vbRed '禁用时的颜色
'Const StartColor = &H80000008 '启用时的颜色
'Const SelectColor = &H8000000A '选中时背景颜色
'Const MoveColor = &H80000004 '移动时背景颜色
'Const noSelectColor = &HFFFFFF '未选中时背景颜色

'Const SelectColor = &HFFFFFF
'Const MoveColor = &H8000000A
'Const noSelectColor = &H80000004

Const COLOR_SELECT = &H80000004 '选中时背景颜色
Const COLOR_MOVE = &HFFFFFF '鼠标移动到按钮时背景颜色
Const COLOR_NOT_SELECT = &H8000000A '未选中时背景颜色

'Private mobjFindKey             As CommandBarPopup      '查询
Private mstrFindKey             As String               '查询串
Private m_strCurTypeName        As String               '当前选中的方式
Private m_strCurFileName        As String               '当前选中的名称
Private m_strCurVision          As String               '当前选中的版本
Private m_strCurEditDate        As String               '当前选中的修改日期
Private m_strCurSysNum          As String               '当前选中的系统
Private m_strCurSetupPath       As String               '当前选中的安装路径
Private m_strCurSetupPathADD       As String         '当前选中的附加安装路径
Private m_strCurSysOption       As String               '当前选中的系统参数
Private m_strCurFileExplanation As String               '当前选中的文件说明
Private m_strCurSellFile        As String               '当前选中的引用文件
Private m_blnCurReg             As Boolean              '当前选中的文件是否注册
Private m_blnCurUpData          As Boolean              '当前选中的文件是否强制覆盖
Private mintfgMainTag           As Integer              '当前表格显示 0-在用文件 1-弃用文件
Private mrsTemp      As New ADODB.Recordset
Private mstrLocationFileName As String
Public blnRefreshData As Boolean '界面切换刷新判断标志

Public Enum RegFileType
    RFT_NotReg = 0                  '不注册的对象
    RFT_NormalReg = 1               '常规注册，自动识别.NET部件，.NET部件通过Regasm注册，其他通过调用DLLRegServer注册
    RFT_NETGAC = 2                  'NET程序集注册，通过gacutil注册到全局程序集缓存
    RFT_NETServer = 3               'NET服务注册，通过installUtil进行安装卸载。
    RFT_NETComReg = 4               '.NET Com部件注册，通过调用Regasm完成
    RFT_VBComReg = 5                '通过手写注册表注册
    RFT_DelphiComReg = 6            'DelphiCom注册，通过DLLRegServer注册
    RFT_PBComReg = 7                'PBCom注册，通过DLLRegServer注册
End Enum

Public Enum FileListCol
    FC_序号 = 0
    FC_文件类型 = 1
    FC_文件名 = 2
    FC_版本号 = 3
    FC_修改日期 = 4
    FC_所属系统 = 5
    FC_业务部件 = 6
    FC_真实路径 = 7
    FC_类型ID = 8
    FC_安装路径 = 9
    FC_系统参数 = 10
    FC_自动注册 = 11
    FC_文件说明 = 12
    FC_强制覆盖 = 13
    FC_附加安装路径 = 14
    FC_列数 = 15
End Enum

Public Enum ExpFileListCol
    EFC_序号 = 0
    EFC_文件名 = 1
    EFC_系统版本 = 2
    EFC_安装路径 = 3
    EFC_系统编号 = 4
    EFC_文件说明 = 5
    EFC_列数 = 6
End Enum

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    objPrint.Title.Text = IIf(mintfgMainTag = 0, "升级文件清单", "弃用文件清单")

    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "打印时间：" & Format(date, "yyyy年MM月dd日")
    Set objPrint.Body = IIf(mintfgMainTag = 0, Me.vsfFileList, Me.vsfExpFileList)
    objPrint.BelowAppRows.Add objRow
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub chkFilter_Click()
    RefreshData
End Sub

Private Sub cmdExpired_Click()
    Call StandardAbandon
End Sub

Private Sub cmdAdd_Click()
    '增加
    Call StandardAdd
End Sub

Private Sub cmdDel_Click()
    '删除
    Call StandardDel
End Sub

Private Sub cmdEdit_Click()
    '修改
     Call StandardEdit
End Sub

Private Sub cmdRepairFiles_Click()
    Dim frmRepair As New frmFilesRepair
    
    frmRepair.ShowMe
End Sub

Private Sub cmdRepairList_Click()
    Dim frmMsgbox As New frmMessageBox
    If frmMsgbox.ShowMe(0, gstrSysName) Then
        Call ExecuteProcedure("zlFilesUpgrade_Repair", Me.Caption)
        Me.RefreshData
    End If
'    If MsgBox("修正后需要重新上传所有文件至服务器！是否需要修当前正在使用的文件清单？", vbQuestion + vbOKCancel, gstrSysName) <> vbCancel Then
'        Call ExecuteProcedure("zlFilesUpgrade_Repair", Me.Caption)
'   End If
End Sub

Private Sub cmdUpload_Click()
    Call StandardUpLoad
End Sub

Private Sub vsfExpFileList_AfterSort(ByVal Col As Long, Order As Integer)
    vsfExpFileList.Row = vsfExpFileList.FindRow(mstrLocationFileName, , 2)
    If vsfExpFileList.Row > 0 Then vsfExpFileList.ShowCell vsfExpFileList.Row, 0
End Sub

Private Sub vsfExpFileList_RowColChange()
    mstrLocationFileName = vsfExpFileList.TextMatrix(vsfExpFileList.Row, 2)
End Sub

Private Sub vsfFileList_AfterSort(ByVal Col As Long, Order As Integer)
    vsfFileList.Row = vsfFileList.FindRow(mstrLocationFileName, , 2)
    If vsfFileList.Row > 0 Then vsfFileList.ShowCell vsfFileList.Row, 0
End Sub

Private Sub vsfFileList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If LblBtn.Item(0).BackColor = COLOR_MOVE Then LblBtn.Item(0).BackColor = COLOR_NOT_SELECT
    If LblBtn.Item(1).BackColor = COLOR_MOVE Then LblBtn.Item(1).BackColor = COLOR_NOT_SELECT
End Sub

Private Sub vsfExpFileList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If LblBtn.Item(0).BackColor = COLOR_MOVE Then LblBtn.Item(0).BackColor = COLOR_NOT_SELECT
    If LblBtn.Item(1).BackColor = COLOR_MOVE Then LblBtn.Item(1).BackColor = COLOR_NOT_SELECT
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        '刷新
    End If
    If KeyCode = vbKeyDelete Then
        If cmdDel.Enabled Then
            cmdDel_Click
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If LblBtn.Item(0).BackColor = COLOR_MOVE Then LblBtn.Item(0).BackColor = COLOR_NOT_SELECT
    If LblBtn.Item(1).BackColor = COLOR_MOVE Then LblBtn.Item(1).BackColor = COLOR_NOT_SELECT
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    vsfFileList.Move 50, 950, Me.ScaleWidth - 120, Me.ScaleHeight - 950 - 600
    vsfExpFileList.Move 50, 950, Me.ScaleWidth - 120, Me.ScaleHeight - 950 - 600
    
'    LblBtn.Item(0).Move Me.ScaleWidth / 2 - LblBtn.Item(0).Width - 250, 1100
'    LblBtn.Item(1).Move Me.ScaleWidth / 2 - 265, 1100
    
    LblBtn.Item(0).Move 50, 700
    LblBtn.Item(1).Move LblBtn.Item(0).Width - 15, 700

'    LblBtn.Item(0).Move 50, 1100, vsfFileList.Width / 2
'    LblBtn.Item(1).Move vsfFileList.Width / 2 + 30, 1100, vsfFileList.Width / 2 + 15
    lblItem.Item(0).Move 100, vsfFileList.Top + vsfFileList.Height + 200
    picBtn.Move lblItem.Item(0).Left + 1100, lblItem.Item(0).Top - 50

End Sub


'==============================================================================
'=功能： 窗口初始化
'==============================================================================
Private Sub Form_Load()
  On Error GoTo errH
    KeyPreview = True
    '查找框初始化
    txtFind.Tag = "请输入文件名称查找"
    txtFind.Text = txtFind.Tag
    txtFind.ForeColor = vbGrayText
    mintfgMainTag = 0
'    LblBtn.Item(0).Move 50, 1100
'    LblBtn.Item(1).Move LblBtn.Item(0).Width + 30, 1055, LblBtn.Item(1).Width, LblBtn.Item(1).Height + 45
    '填充Combo
'    Call InitComBo
'    Call InitVsfMain

'    LoadFilesList
'    LoadExpFilesList
'    LblBtn_Click 0
'    Call SetMenu
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub


'==============================================================================
'=功能： 网格fgMain单击后刷新状态信息
'==============================================================================
Private Sub vsfFileList_Click()
    On Error GoTo errH
    vsfFileList_SelChange
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'==============================================================================
'=功能： 网格行列变化时更新基本信息
'==============================================================================
Private Sub vsfFileList_RowColChange()
    On Error GoTo errH
    Call vsfFileList_SelChange
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'==============================================================================
'=功能： 网格选择行列变化时更新基本信息
'==============================================================================
Private Sub vsfFileList_SelChange()
    Dim lngID       As Long
    On Error GoTo errH

    If vsfFileList.Row = 0 Then Exit Sub
    m_strCurTypeName = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 1)) = 0, 0, vsfFileList.Cell(flexcpText, vsfFileList.Row, 1))   '获取ID
    m_strCurFileName = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 2)) = 0, 0, vsfFileList.Cell(flexcpText, vsfFileList.Row, 2))     '文件名
    m_strCurVision = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 3)) = 0, 0, vsfFileList.Cell(flexcpText, vsfFileList.Row, 3))
    m_strCurEditDate = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 4)) = 0, 0, vsfFileList.Cell(flexcpText, vsfFileList.Row, 4))
    m_strCurSysNum = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 5)) = 0, "", vsfFileList.Cell(flexcpText, vsfFileList.Row, 5))
    m_strCurSellFile = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 6)) = 0, "", vsfFileList.Cell(flexcpText, vsfFileList.Row, 6))
    m_strCurSetupPath = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 7)) = 0, "", vsfFileList.Cell(flexcpText, vsfFileList.Row, 7))
    m_strCurSetupPathADD = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, FC_附加安装路径)) = 0, "", vsfFileList.Cell(flexcpText, vsfFileList.Row, FC_附加安装路径))
    m_strCurSysOption = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 10)) = 0, "", vsfFileList.Cell(flexcpText, vsfFileList.Row, 10))
    m_blnCurReg = IIf(vsfFileList.Cell(flexcpText, vsfFileList.Row, 11) = "是", True, False) '自动注册
    m_strCurFileExplanation = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 12)) = 0, "", vsfFileList.Cell(flexcpText, vsfFileList.Row, 12)) '文件说明
    m_blnCurUpData = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 13)) = 0, False, vsfFileList.Cell(flexcpText, vsfFileList.Row, 13)) '强制覆盖

    If m_strCurTypeName = "三方部件" Then
        cmdEdit.Enabled = True
        cmdExpired.Enabled = True
        cmdDel.Enabled = True
    Else
        cmdEdit.Enabled = False
        cmdExpired.Enabled = False
        cmdDel.Enabled = False
    End If
    mstrLocationFileName = m_strCurFileName
'    Call SetMenu
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

Private Sub vsfFileList_DblClick()
    If vsfFileList.MouseRow <> vsfFileList.Row Then Exit Sub '固定行双击无效
    If m_strCurTypeName = "三方部件" Then
        Call StandardEdit
    End If
End Sub

Public Sub LoadFilesList(Optional ByVal strFilter As String, Optional ByVal strLocationName As String)
    Dim i, j As Long
    Dim strSQL       As String
    Dim strTemp      As String
    Dim arrSys As Variant
    On Error GoTo errH
    
    If strFilter = "" Then strFilter = "0,1,2,3,4,5"
    
    With vsfFileList
        .Redraw = flexRDNone
        .Tag = ""
        .Rows = 1
        .Clear
        .Cols = FC_列数
        
        .Cell(flexcpText, 0, FC_序号) = "序号"
        .Cell(flexcpAlignment, 0, FC_序号) = flexAlignCenterCenter
        
        .Cell(flexcpText, 0, FC_文件类型) = "文件类型"
        .Cell(flexcpAlignment, 0, FC_文件类型) = flexAlignCenterCenter
        
        .Cell(flexcpText, 0, FC_文件名) = "文件名"
        .Cell(flexcpAlignment, 0, FC_文件名) = flexAlignCenterCenter
        .ColWidth(FC_文件名) = 2200
        
        .Cell(flexcpText, 0, FC_版本号) = "版本号"
        .Cell(flexcpAlignment, 0, FC_版本号) = flexAlignCenterCenter
        .ColWidth(FC_版本号) = 1200
        
        .Cell(flexcpText, 0, FC_修改日期) = "修改日期"
        .Cell(flexcpAlignment, 0, FC_修改日期) = flexAlignCenterCenter
        
        .Cell(flexcpText, 0, FC_所属系统) = "所属系统"
        .Cell(flexcpAlignment, 0, FC_所属系统) = flexAlignCenterCenter
        .ColWidth(FC_所属系统) = 1800
        
        .Cell(flexcpText, 0, FC_业务部件) = "业务部件"
        .Cell(flexcpAlignment, 0, FC_业务部件) = flexAlignCenterCenter
        .ColWidth(FC_业务部件) = 3000

        .Cell(flexcpText, 0, FC_真实路径) = "安装路径"
        .Cell(flexcpAlignment, 0, FC_真实路径) = flexAlignCenterCenter
        .ColHidden(FC_真实路径) = True
        
        .Cell(flexcpText, 0, FC_类型ID) = "类型ID"
        .Cell(flexcpAlignment, 0, FC_类型ID) = flexAlignCenterCenter
        .ColHidden(FC_类型ID) = True

        .Cell(flexcpText, 0, FC_安装路径) = "安装路径"
        .Cell(flexcpAlignment, 0, FC_安装路径) = flexAlignCenterCenter
        .ColWidth(FC_安装路径) = 2000

        .Cell(flexcpText, 0, FC_系统参数) = "系统参数"
        .Cell(flexcpAlignment, 0, FC_系统参数) = flexAlignCenterCenter
        .ColHidden(FC_系统参数) = True
        
        .Cell(flexcpText, 0, FC_自动注册) = "自动注册"
        .Cell(flexcpAlignment, 0, FC_自动注册) = flexAlignCenterCenter
        .ColWidth(FC_自动注册) = 1000

        .Cell(flexcpText, 0, FC_文件说明) = "文件说明"
        .Cell(flexcpAlignment, 0, FC_文件说明) = flexAlignCenterCenter
        .ColWidth(FC_文件说明) = 5000

        .Cell(flexcpText, 0, FC_强制覆盖) = "强制覆盖"
        .ColHidden(FC_强制覆盖) = True
        
        .Cell(flexcpText, 0, FC_附加安装路径) = "附加安装路径"
        .ColHidden(FC_附加安装路径) = True
        
        If CheckAndAdjustMustTable("zlFilesUpgrade", , True) = False Then
            Exit Sub
        End If
        
        strSQL = "Select a.序号,a.文件类型 As 类型ID,Decode(a.文件类型, 0, '公共部件', 1, '应用部件', 2, '帮助文件', 3, '其它文件', 4, '三方部件', 5, '系统文件', '未知类型') As 文件类型, a.文件名, a.文件版本号 版本号, a.修改日期," & vbNewLine & _
                        "       a.所属系统, a.业务部件,a.安装路径,a.文件说明,a.自动注册,a.附加安装路径" & vbNewLine & _
                        "From zlFilesUpgrade A" & vbNewLine & _
                        "Where a.文件类型 In (" & strFilter & ") order by lpad(a.序号,5,'0')"
        Call OpenRecordset(mrsTemp, strSQL, Me.Caption)

        '数据填入
        .Rows = mrsTemp.RecordCount + 1

        i = 1
        Do Until mrsTemp.EOF
            .Cell(flexcpText, i, FC_序号) = Nvl(mrsTemp.Fields("序号"), 0) 'mrsTemp.AbsolutePosition
            .Cell(flexcpAlignment, i, FC_序号) = flexAlignLeftCenter

            .Cell(flexcpText, i, FC_文件类型) = Nvl(mrsTemp.Fields("文件类型"))
            .Cell(flexcpAlignment, i, FC_文件类型) = flexAlignCenterCenter

            .Cell(flexcpText, i, FC_文件名) = Nvl(mrsTemp.Fields("文件名"))
            .Cell(flexcpAlignment, i, FC_文件名) = flexAlignLeftCenter

            strTemp = Nvl(mrsTemp.Fields("版本号"))
'            strTemp = GetFileVision(strTemp)

            .Cell(flexcpText, i, FC_版本号) = strTemp
            .Cell(flexcpAlignment, i, FC_版本号) = flexAlignLeftCenter

            If Nvl(mrsTemp.Fields("修改日期")) <> "" Then
                strTemp = Format(Nvl(mrsTemp.Fields("修改日期")), "yyyy-mm-dd hh:mm:ss")
            Else
                strTemp = ""
            End If

            .Cell(flexcpText, i, FC_修改日期) = strTemp
            .Cell(flexcpAlignment, i, FC_修改日期) = flexAlignCenterCenter

            strTemp = Nvl(mrsTemp.Fields("所属系统"))

            If Trim(strTemp) <> "" Then
                arrSys = Split(Trim(strTemp), ",")
                strTemp = ""
                For j = 0 To UBound(arrSys)
                    If GetSystemName(arrSys(j)) <> "" Then strTemp = strTemp & "，" & GetSystemName(arrSys(j))
                Next
                strTemp = Mid(strTemp, 2)
            End If

            .Cell(flexcpText, i, FC_所属系统) = strTemp
            .Cell(flexcpData, i, FC_所属系统) = Nvl(mrsTemp.Fields("所属系统"))
            .Cell(flexcpAlignment, i, FC_所属系统) = flexAlignLeftCenter

            .Cell(flexcpText, i, FC_业务部件) = Nvl(mrsTemp.Fields("业务部件"))
            .Cell(flexcpAlignment, i, FC_业务部件) = flexAlignLeftCenter

            .Cell(flexcpText, i, FC_真实路径) = Nvl(mrsTemp.Fields("安装路径"))
            .Cell(flexcpAlignment, i, FC_真实路径) = flexAlignLeftCenter

            .Cell(flexcpText, i, FC_类型ID) = Nvl(mrsTemp.Fields("类型ID"))
            .Cell(flexcpAlignment, i, FC_类型ID) = flexAlignLeftTop

            .Cell(flexcpText, i, FC_安装路径) = Nvl(mrsTemp.Fields("安装路径"))
            .Cell(flexcpAlignment, i, FC_安装路径) = flexAlignLeftCenter

'            .Cell(flexcpText, i, FC_所属系统) = Nvl(mrsTemp.Fields("所属系统")) 'NVL(mrsTemp.Fields("系统参数"))
'            .Cell(flexcpAlignment, i, FC_所属系统) = flexAlignCenterCenter

            .Cell(flexcpText, i, FC_自动注册) = IIf(Nvl(mrsTemp.Fields("自动注册"), "") = "0", "否", "是")
            .Cell(flexcpAlignment, i, FC_自动注册) = flexAlignCenterCenter

            .Cell(flexcpText, i, FC_文件说明) = Nvl(mrsTemp.Fields("文件说明"), "")
            .Cell(flexcpAlignment, i, FC_文件说明) = flexAlignLeftCenter

            .Cell(flexcpText, i, FC_强制覆盖) = ""
       
            .Cell(flexcpText, i, FC_附加安装路径) = Nvl(mrsTemp.Fields("附加安装路径"), "")

            mrsTemp.MoveNext
            i = i + 1
        Loop
        
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
        .AllowUserResizing = flexResizeColumns
'        .AllowSelection = False
        
        .Redraw = flexRDBuffered
        
        '刷新定位
        If strLocationName <> "" Then
            strLocationName = UCase(strLocationName)
            For j = 0 To .Rows - 1
                If UCase(.TextMatrix(j, 2)) = strLocationName Then .Row = j: Call .ShowCell(j, 2): Exit For
            Next
        Else
            If .Rows > 1 Then .Row = 1
        End If
        '刷新修改、删除按钮状态
        vsfFileList_SelChange

        If .Visible = True Then .SetFocus
'         Call SetMenu
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

Public Sub LoadExpFilesList(Optional ByVal strFilter As String, Optional ByVal strLocationName As String)
'废除表格
    Dim i, j As Long
    Dim strSQL       As String
    Dim strTemp      As String
    On Error GoTo errH

    With vsfExpFileList
        .Redraw = flexRDNone
        .Tag = ""
        .Rows = 1
        .Clear
        .Cols = EFC_列数
        
        .Cell(flexcpText, 0, EFC_序号) = "序号"
        .Cell(flexcpAlignment, 0, EFC_序号) = flexAlignCenterCenter
        
        .Cell(flexcpText, 0, EFC_文件名) = "文件名"
        .Cell(flexcpAlignment, 0, EFC_文件名) = flexAlignCenterCenter
        .ColWidth(EFC_文件名) = 1800
        
        .Cell(flexcpText, 0, EFC_系统版本) = "系统版本"
        .Cell(flexcpAlignment, 0, EFC_系统版本) = flexAlignCenterCenter
        .ColWidth(EFC_系统版本) = 1000
        
        .Cell(flexcpText, 0, EFC_安装路径) = "安装路径"
        .Cell(flexcpAlignment, 0, EFC_安装路径) = flexAlignCenterCenter
        .ColWidth(EFC_安装路径) = 3000
        
        .Cell(flexcpText, 0, EFC_系统编号) = "系统编号"
        .Cell(flexcpAlignment, 0, EFC_系统编号) = flexAlignCenterCenter
        .ColWidth(EFC_系统编号) = 1000
        
        .Cell(flexcpText, 0, EFC_文件说明) = "文件说明"
        .Cell(flexcpAlignment, 0, EFC_文件说明) = flexAlignCenterCenter
        .ColWidth(EFC_文件说明) = 5000

        If CheckAndAdjustMustTable("zlFilesUpgrade", , True) = False Then
            Exit Sub
        End If
        
        strSQL = "select * from zlfilesexpired"
        Call OpenRecordset(mrsTemp, strSQL, Me.Caption)

        '数据填入
        .Rows = mrsTemp.RecordCount + 1

        i = 1
        Do Until mrsTemp.EOF
            .Cell(flexcpText, i, EFC_序号) = i 'Nvl(mrsTemp.Fields("序号"), 0) 'mrsTemp.AbsolutePosition
            .Cell(flexcpAlignment, i, EFC_序号) = flexAlignCenterCenter

            .Cell(flexcpText, i, EFC_文件名) = Nvl(mrsTemp.Fields("文件名"))
            .Cell(flexcpAlignment, i, EFC_文件名) = flexAlignLeftCenter

            .Cell(flexcpText, i, EFC_系统版本) = Nvl(mrsTemp.Fields("系统版本"))
            .Cell(flexcpAlignment, i, EFC_系统版本) = flexAlignLeftCenter

            .Cell(flexcpText, i, EFC_安装路径) = Nvl(mrsTemp.Fields("安装路径"))
            .Cell(flexcpAlignment, i, EFC_安装路径) = flexAlignLeftCenter

            .Cell(flexcpText, i, EFC_系统编号) = Nvl(mrsTemp.Fields("系统编号"))
            .Cell(flexcpAlignment, i, EFC_系统编号) = flexAlignCenterCenter

            .Cell(flexcpText, i, EFC_文件说明) = Nvl(mrsTemp.Fields("说明"), "")
            .Cell(flexcpAlignment, i, EFC_文件说明) = flexAlignLeftCenter

            mrsTemp.MoveNext
            i = i + 1
        Loop
        
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
        .AllowUserResizing = flexResizeColumns
'        .AllowSelection = False
        
        .Redraw = flexRDBuffered
        
        '刷新定位
        If strLocationName <> "" Then
            strLocationName = UCase(strLocationName)
            For j = 0 To .Rows - 1
                If UCase(.TextMatrix(j, 2)) = strLocationName Then .Row = j: Call .ShowCell(j, 2): Exit For
            Next
        Else
            If .Rows > 1 Then .Row = 1
        End If
'        '刷新修改、删除按钮状态
'        vsfFileList_SelChange

        If .Visible = True Then .SetFocus
'         Call SetMenu
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub
'==============================================================================
'=功能： 显示记录数信息
'==============================================================================
Public Sub SetMenu()
    If mintfgMainTag = 0 Then
        frmMDIMain.stbThis.Panels(2).Text = "列表中共显示有" & vsfFileList.Rows - 1 & "行数据。"
    Else
        frmMDIMain.stbThis.Panels(2).Text = "列表中共显示有" & vsfExpFileList.Rows - 1 & "行数据。"
    End If
End Sub

'==============================================================================
'=功能： 检查表是否是新表或者表是否存在
'==============================================================================
Private Function CheckTable() As Boolean
    On Error GoTo errH
    Dim strSQL As String
    Dim i As Integer
    Dim blnUse As Boolean
    strSQL = "select * from zlFilesUpgrade where rownum =1"

    Call OpenRecordset(mrsTemp, strSQL, Me.Caption)
    If mrsTemp.RecordCount >= 0 Then
        For i = 1 To mrsTemp.Fields.Count
            If mrsTemp.Fields.Item(i - 1).name = "所属系统" Then
                blnUse = True
                Exit For
            End If
        Next

        If blnUse Then
            CheckTable = True
        Else
            MsgBox "在zlFilesUpgrade表中,没有找到相应的字段!" & vbCrLf & "请检查表结构是否为最新!", vbInformation
            CheckTable = False
        End If
    End If
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

'获取版本的直观显示值
Private Function GetFileVision(ByVal strVision As String) As String
    Dim lng版本号 As Variant
    Dim str版本号 As String
    If Len(strVision) > 0 Then
        lng版本号 = strVision
        str版本号 = Int(lng版本号 / 10 ^ 8)
        If Len(lng版本号) > 9 Then
            lng版本号 = Right(lng版本号, 9) Mod (10 ^ 8)
        Else
            lng版本号 = lng版本号 Mod (10 ^ 8)
        End If

        str版本号 = str版本号 & "." & Int(lng版本号 / 10 ^ 4)
        lng版本号 = lng版本号 Mod 10 ^ 4
        str版本号 = str版本号 & "." & lng版本号
        GetFileVision = str版本号
    End If
End Function

Private Function GetCommpentVersion(ByVal strFile As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取指定控件的版本号
    '入参:
    '出参:
    '返回:成功,返回版本号,否则返回空
    '编制:刘兴洪
    '日期:2009-01-16 16:59:34
    '-----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim strVer As String, varVersion As Variant

    err = 0: On Error Resume Next
    '获取文件版本号
    strVer = objFile.GetFileVersion(strFile)
    If err <> 0 Then
        err.Clear: err = 0
        GetCommpentVersion = ""
        Exit Function
    End If
    If Trim(strVer) <> "" Then
        varVersion = Split(strVer, ".")
        If UBound(varVersion) > 2 Then
            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(3)
        ElseIf UBound(varVersion) = 2 Then
            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(2)
        End If
    End If
    GetCommpentVersion = strVer
End Function

Private Function GetSystemName(ByVal strNum As String) As String
'传入系统编号，获得对应系统名称，若未找到
On err GoTo errH
    Select Case strNum
        Case "1", "100"
            GetSystemName = "医院系统标准版"
        Case "2", "200"
            GetSystemName = "人事工资系统"
        Case "3", "300"
            GetSystemName = "病案管理系统"
        Case "4", "400"
            GetSystemName = "物资供应系统"
        Case "5", "500"
            GetSystemName = "财务核算系统"
        Case "6", "600"
            GetSystemName = "设备管理系统"
        Case "7", "700"
            GetSystemName = "成本效益核算系统"
        Case "21", "2100"
            GetSystemName = "体检管理系统"
        Case "22", "2200"
            GetSystemName = "血库管理系统"
        Case "23", "2300"
            GetSystemName = "院感管理系统"
        Case "24", "2400"
            GetSystemName = "手麻管理系统"
        Case "25", "2500"
            GetSystemName = "临床检验管理系统"
        Case "26", "2600"
            GetSystemName = "病人自助服务系统"
    End Select
    Exit Function

errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Function

'==============================================================================
'=修改文件
'==============================================================================
Private Sub StandardEdit()
    Dim f As New frmScriptEdit
    Dim strSysNum As String
    On Error GoTo errH
    Dim strLocationName As String
    strSysNum = 100
    
    vsfFileList.Row = vsfFileList.FindRow(m_strCurFileName, , 2)
    Call f.ShowForm("修改", m_strCurTypeName, m_strCurFileName, m_strCurSysNum, strSysNum, m_strCurVision, m_strCurSetupPath, m_strCurEditDate, m_strCurSysOption, m_strCurFileExplanation, m_strCurSellFile, m_blnCurReg, m_blnCurUpData, "0", m_strCurSetupPathADD)
    Call RefreshData(m_strCurFileName)
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub


'==============================================================================
'=新增文件
'==============================================================================
Private Sub StandardAdd()
    Dim f As New frmScriptEdit
    Dim strSysNum As String
    Dim strLocationName As String

    On Error GoTo errH
    strSysNum = 1
    
    strLocationName = f.ShowForm("新增", m_strCurTypeName, m_strCurFileName, m_strCurSysNum, strSysNum, m_strCurVision, m_strCurSetupPath, m_strCurEditDate, m_strCurSysOption, m_strCurFileExplanation, m_strCurSellFile, m_blnCurReg, m_blnCurUpData, "0", m_strCurSetupPathADD)
    If strLocationName = "" Then
        Call RefreshData(m_strCurFileName)
    Else
        Call RefreshData(strLocationName)
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'==============================================================================
'=删除文件
'==============================================================================
Private Sub StandardDel()
    Dim i         As Long
    Dim strName   As String
    Dim rs        As ADODB.Recordset
    Dim strSQL    As String
    Dim strSys    As String
    Dim strSysNum As String
    Dim lngRow    As Long
    Dim lngCurRow As Long
    Dim frmMsgbox As New frmMessageBox
    On Error GoTo errH

    Select Case mintfgMainTag
        Case 0 '在用部件删除
            If vsfFileList.SelectedRows = 0 Then Exit Sub
            If m_strCurTypeName <> "三方部件" Then Exit Sub
            
            If vsfFileList.SelectedRows = 1 Then
'                If MsgBox("你确认要删除" & m_strCurFileName & "部件吗？", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then Exit Sub
                If frmMsgbox.ShowMe(2, gstrSysName, "你确认要删除" & m_strCurFileName & "部件吗？") = False Then Exit Sub
            Else
'                If MsgBox("你确认要删除选择的" & vsfFileList.SelectedRows & "个部件吗？", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then Exit Sub
                If frmMsgbox.ShowMe(2, gstrSysName, "你确认要删除选择的" & vsfFileList.SelectedRows & "个部件吗？") = False Then Exit Sub
            End If
'            gcnOracle.BeginTrans
            lngRow = vsfFileList.FindRow(CStr(m_strCurFileName), , 2)
            For i = 0 To vsfFileList.SelectedRows
                If vsfFileList.SelectedRow(i) > 0 Then
                    lngCurRow = vsfFileList.SelectedRow(i)
                    If vsfFileList.TextMatrix(lngCurRow, FC_文件类型) = "三方部件" Then
                        strName = UCase(IIf(Len(vsfFileList.Cell(flexcpText, lngCurRow, 2)) = 0, 0, vsfFileList.Cell(flexcpText, lngCurRow, 2)))
                        
                        gstrSQL = "delete zlFilesUpgrade where upper(文件名)= upper('" & strName & "')"
                        gcnOracle.Execute gstrSQL
                    End If
                End If
            Next
'            gcnOracle.CommitTrans

            If lngRow <> -1 Then
                If lngRow >= 2 And vsfFileList.Rows > 2 Then
                  vsfFileList.Select lngRow - 1, 2
                  vsfFileList.ShowCell lngRow - 1, 2
                End If
            End If
            
        Case 1 '弃用部件删除
         If vsfExpFileList.SelectedRows = 1 Then
                If MsgBox("你确认要删除" & vsfExpFileList.Cell(flexcpText, vsfExpFileList.Row, 1) & "弃用部件吗？", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then Exit Sub
            Else
                If MsgBox("你确认要删除选择的" & vsfExpFileList.SelectedRows & "弃用个部件吗？", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then Exit Sub
            End If
        '    gcnOracle.BeginTrans
'            lngRow = vsfExpFileList.FindRow(CStr(vsfExpFileList.Cell(flexcpText, lngCurRow, 1)), , 2)
            For i = 0 To vsfExpFileList.SelectedRows
                If vsfExpFileList.SelectedRow(i) Then
                    lngCurRow = vsfExpFileList.SelectedRow(i)
                    If lngCurRow <> -1 Then
                        gstrSQL = "delete zlfilesexpired where upper(文件名)= upper('" & Trim(vsfExpFileList.Cell(flexcpText, lngCurRow, 1)) & "')"
                        gcnOracle.Execute gstrSQL
                    End If
                End If
            Next
        '    gcnOracle.CommitTrans
        End Select
        Call RefreshData
    Exit Sub
errH:
'    gcnOracle.RollbackTrans
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'==============================================================================
'=弃用文件
'==============================================================================
Private Sub StandardAbandon()
    Dim strName   As String
    Dim lngRow    As Long
    Dim i As Long
    Dim lngCurRow As Long
    Dim frmMsgbox As New frmMessageBox
    
    lngRow = vsfFileList.FindRow(CStr(m_strCurFileName), , 2)
    
    If vsfFileList.SelectedRows > 1 Then
        If frmMsgbox.ShowMe(1, gstrSysName, "你确认要弃用选择的 " & vsfFileList.SelectedRows & " 个部件吗？") = False Then Exit Sub
'        If MsgBox("你确认要弃用选择的 " & vsfFileList.SelectedRows & " 个部件吗？", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then Exit Sub
    Else
        If frmMsgbox.ShowMe(1, gstrSysName, "你确认要弃用 " & vsfFileList.TextMatrix(lngRow, 2) & " 部件吗？") = False Then Exit Sub
'        If MsgBox("你确认要弃用 " & vsfFileList.TextMatrix(lngRow, 2) & " 个部件吗？", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then Exit Sub
    End If
    
    For i = 0 To vsfFileList.SelectedRows
        If vsfFileList.SelectedRow(i) > 0 Then
            lngCurRow = vsfFileList.SelectedRow(i)
            If vsfFileList.TextMatrix(lngCurRow, FC_文件类型) = "三方部件" Then
                strName = IIf(Len(vsfFileList.Cell(flexcpText, lngCurRow, 2)) = 0, 0, vsfFileList.Cell(flexcpText, lngCurRow, 2))
                strName = UCase(strName)
                gstrSQL = "  Insert Into Zlfilesexpired (文件名,安装路径,系统编号,系统版本,说明) select " & _
                                "'" & vsfFileList.Cell(flexcpText, lngCurRow, 2) & "','" & vsfFileList.Cell(flexcpText, lngCurRow, 7) & "','" & vsfFileList.Cell(flexcpData, lngCurRow, 5) & "','" & vsfFileList.Cell(flexcpText, lngCurRow, 3) & "','" & vsfFileList.Cell(flexcpText, lngCurRow, 10) & "' " & _
                                " from dual Where Not Exists " & _
                                " (Select 1 From Zlfilesexpired A Where A.文件名 = '" & vsfFileList.Cell(flexcpText, lngCurRow, 2) & "')"
                gcnOracle.Execute gstrSQL
                gstrSQL = "delete zlFilesUpgrade where upper(文件名)= upper('" & strName & "')"
                gcnOracle.Execute gstrSQL
            End If
        End If
    Next
    
    Call RefreshData
    If lngRow <> -1 Then
        If lngRow >= 2 And vsfFileList.Rows > 2 Then
          vsfFileList.Select lngRow - 1, 2
          vsfFileList.ShowCell lngRow - 1, 2
        End If
    End If
    Exit Sub
End Sub

'==============================================================================
'=上传文件
'==============================================================================
Private Sub StandardUpLoad()
    Dim frmUpload As New frmFilesUpload
    
    frmUpload.ShowMe

End Sub

Private Sub LblBtn_Click(Index As Integer)
    Select Case Index
         Case 0 '在用部件
            If mintfgMainTag = 1 Then
                LblBtn.Item(0).BackColor = COLOR_SELECT
                LblBtn.Item(1).BackColor = COLOR_NOT_SELECT
                vsfFileList.Visible = True
                vsfExpFileList.Visible = False
    
                cmdAdd.Enabled = True
                chkFilter.Visible = True
                
                mintfgMainTag = 0
                
                RefreshData

    '            LblBtn.Item(0).Move LblBtn.Item(0).Left, LblBtn.Item(0).Top - 45, LblBtn.Item(0).Width, LblBtn.Item(0).Height + 45
    '            LblBtn.Item(1).Move LblBtn.Item(1).Left, LblBtn.Item(1).Top + 45, LblBtn.Item(1).Width, LblBtn.Item(1).Height - 45
            End If
         Case 1 '弃用部件
            If mintfgMainTag = 0 Then
                LblBtn.Item(1).BackColor = COLOR_SELECT
                LblBtn.Item(0).BackColor = COLOR_NOT_SELECT
                vsfFileList.Visible = False
                vsfExpFileList.Visible = True
                
                cmdAdd.Enabled = False
                cmdEdit.Enabled = False
                cmdExpired.Enabled = False
                cmdDel.Enabled = True
                chkFilter.Visible = False
                
                mintfgMainTag = 1

                RefreshData

                If vsfExpFileList.Rows <= vsfExpFileList.FixedRows Then cmdDel.Enabled = False
    '            LblBtn.Item(0).Move LblBtn.Item(0).Left, LblBtn.Item(0).Top + 45, LblBtn.Item(0).Width, LblBtn.Item(0).Height - 45
    '            LblBtn.Item(1).Move LblBtn.Item(1).Left, LblBtn.Item(1).Top - 45, LblBtn.Item(1).Width, LblBtn.Item(1).Height + 45
            End If
    End Select
End Sub

Private Sub LblBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If LblBtn.Item(Index).BackColor = COLOR_NOT_SELECT Then
        LblBtn.Item(Index).BackColor = COLOR_MOVE
    End If
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text = txtFind.Tag Then
        txtFind.Text = ""
        txtFind.ForeColor = vbBlack
    End If
End Sub

'==============================================================================
'=快速定位
'==============================================================================
Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim lngLoop     As Long

    On Error GoTo errH

    lngRow = 0
    If txtFind.Locked Then Exit Sub
    If mstrFindKey = "名称" Then mstrFindKey = "文件名称"
    If KeyAscii = vbKeyReturn Then
        '读取大于当前行的记录数据
        For lngLoop = vsfFileList.Row + 1 To vsfFileList.Rows - 1
            If InStr(UCase(vsfFileList.TextMatrix(lngLoop, 2)), UCase(txtFind.Text)) > 0 Then
                lngRow = lngLoop
                Exit For
            End If
        Next
        '读取小于当前行的记录数据
        If lngRow = 0 Then
            For lngLoop = 0 To vsfFileList.Row
                If InStr(UCase(vsfFileList.TextMatrix(lngLoop, 2)), UCase(txtFind.Text)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
        End If
        If vsfFileList.Rows > 1 And lngRow >= 1 Then
            vsfFileList.Row = lngRow
            vsfFileList.ShowCell lngRow, 2
        End If
        'Call LocationObj(txtFind)
    End If
    If mstrFindKey = "文件名称" Then mstrFindKey = "名称"

    Exit Sub
errH:
    mstrFindKey = "名称"
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub txtFind_LostFocus()
    If txtFind.Text = "" Then
        txtFind.Text = txtFind.Tag
        txtFind.ForeColor = vbGrayText
    End If
End Sub

Private Sub InitVsfMain()
With vsfExpFileList
        .Tag = ""
'        .Redraw = flexRDNone
        .Rows = 50
        .Clear
        .Cols = 14
'        Exit Sub
        .Cell(flexcpText, 0, 0) = "序号"
        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 1) = "文件类型"
        .Cell(flexcpAlignment, 0, 1) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 2) = "文件名"
        .Cell(flexcpAlignment, 0, 2) = flexAlignCenterCenter
        .ColWidth(2) = 2200
        .Cell(flexcpText, 0, 3) = "版本号"
        .Cell(flexcpAlignment, 0, 3) = flexAlignCenterCenter
        .ColWidth(3) = 1200
        .Cell(flexcpText, 0, 4) = "修改日期"
        .Cell(flexcpAlignment, 0, 4) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 5) = "所属系统"
        .Cell(flexcpAlignment, 0, 5) = flexAlignCenterCenter
        .ColWidth(5) = 1800
        .Cell(flexcpText, 0, 6) = "业务部件"
        .Cell(flexcpAlignment, 0, 6) = flexAlignCenterCenter
        .ColWidth(6) = 3000

        .Cell(flexcpText, 0, 7) = "安装路径"
        .Cell(flexcpAlignment, 0, 7) = flexAlignCenterCenter
        .ColWidth(7) = 0

        .Cell(flexcpText, 0, 8) = "类型ID"
        .Cell(flexcpAlignment, 0, 8) = flexAlignCenterCenter
        .ColWidth(8) = 0

        .Cell(flexcpText, 0, 9) = "安装路径"
        .Cell(flexcpAlignment, 0, 9) = flexAlignCenterCenter
        .ColWidth(9) = 2000

        .Cell(flexcpText, 0, 10) = "系统参数"
        .Cell(flexcpAlignment, 0, 10) = flexAlignCenterCenter
        .ColWidth(10) = 0
        .Cell(flexcpText, 0, 11) = "自动注册"
        .Cell(flexcpAlignment, 0, 11) = flexAlignCenterCenter
        .ColWidth(11) = 1000

        .Cell(flexcpText, 0, 12) = "文件说明"
        .Cell(flexcpAlignment, 0, 12) = flexAlignCenterCenter
        .ColWidth(12) = 5000

        .Cell(flexcpText, 0, 13) = "强制覆盖"
        .Cell(flexcpAlignment, 0, 13) = flexAlignCenterCenter
        .ColWidth(13) = 0
        
        .ExtendLastCol = True
'        .ScrollTips = True
'        .FocusRect = flexFocusSolid

        .RowHeightMin = 300
        .RowHeightMax = 300
        '最大宽度设置
        .ColWidthMax = 7000
        '自动适应行高、列宽
        .AutoSizeMode = flexAutoSizeRowHeight
'        .AutoSize .ColIndex("客户端升级")
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .Redraw = flexRDBuffered
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
    End With
End Sub

Public Sub RefreshData(Optional strLocationFileName As String = "")
    Select Case mintfgMainTag
        Case 0
            If chkFilter.value = 1 Then
                Call LoadFilesList("4", strLocationFileName)
            Else
                Call LoadFilesList(, strLocationFileName)
            End If
        Case 1
            LoadExpFilesList
    End Select
    SetMenu
End Sub
