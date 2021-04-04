VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdateFile 
   BackColor       =   &H80000005&
   Caption         =   "站点运行控制"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmUpdateFile.frx":0000
   ScaleHeight     =   6180
   ScaleWidth      =   8535
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   2385
      TabIndex        =   20
      Top             =   5745
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3795
      Index           =   1
      Left            =   240
      ScaleHeight     =   3795
      ScaleWidth      =   8160
      TabIndex        =   19
      Top             =   1770
      Width           =   8160
      Begin VSFlex8Ctl.VSFlexGrid fgMain 
         Height          =   3600
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   8040
         _cx             =   14182
         _cy             =   6350
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483630
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   2
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
         FormatString    =   $"frmUpdateFile.frx":04F9
         ScrollTrack     =   0   'False
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
         WallPaperAlignment=   4
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   555
      Index           =   0
      Left            =   5355
      ScaleHeight     =   555
      ScaleWidth      =   3165
      TabIndex        =   18
      Top             =   1245
      Width           =   3165
      Begin VB.CheckBox chk部件 
         BackColor       =   &H80000005&
         Caption         =   "系统文件"
         Height          =   240
         Index           =   5
         Left            =   1065
         TabIndex        =   10
         Top             =   255
         Value           =   1  'Checked
         Width           =   1080
      End
      Begin VB.CheckBox chk部件 
         BackColor       =   &H80000005&
         Caption         =   "三方部件"
         Height          =   240
         Index           =   4
         Left            =   2130
         TabIndex        =   11
         Top             =   255
         Value           =   1  'Checked
         Width           =   1080
      End
      Begin VB.CheckBox chk部件 
         BackColor       =   &H80000005&
         Caption         =   "其他文件"
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   9
         Top             =   255
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chk部件 
         BackColor       =   &H80000005&
         Caption         =   "帮助文件"
         Height          =   240
         Index           =   2
         Left            =   2130
         TabIndex        =   8
         Top             =   0
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chk部件 
         BackColor       =   &H80000005&
         Caption         =   "应用部件"
         Height          =   240
         Index           =   1
         Left            =   1065
         TabIndex        =   7
         Top             =   0
         Value           =   1  'Checked
         Width           =   1050
      End
      Begin VB.CheckBox chk部件 
         BackColor       =   &H80000005&
         Caption         =   "公共部件"
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Value           =   1  'Checked
         Width           =   1050
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   2
      Left            =   210
      ScaleHeight     =   330
      ScaleWidth      =   6600
      TabIndex        =   17
      Top             =   1410
      Width           =   6600
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   3555
         TabIndex        =   5
         Top             =   15
         Width           =   2100
      End
      Begin VB.ComboBox cboSystem 
         Height          =   300
         Left            =   690
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   15
         Width           =   2100
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查找(&Z)"
         Height          =   180
         Left            =   2895
         TabIndex        =   4
         Top             =   75
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "系统(&X)"
         Height          =   180
         Left            =   45
         TabIndex        =   2
         Top             =   75
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "修改(&E)"
      Height          =   360
      Left            =   7455
      TabIndex        =   14
      Top             =   5730
      Width           =   945
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除(&D)"
      Height          =   360
      Left            =   6480
      TabIndex        =   13
      Top             =   5730
      Width           =   945
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "增加(&A)"
      Height          =   360
      Left            =   5505
      TabIndex        =   12
      Top             =   5730
      Width           =   945
   End
   Begin MSComctlLib.ImageList ilsIcon 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateFile.frx":065A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "刷新(&R)"
      Height          =   350
      Left            =   4395
      TabIndex        =   0
      Top             =   5745
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "升级文件管理"
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
      Left            =   195
      TabIndex        =   16
      Top             =   105
      Width           =   1440
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "升级文件分类查看,第三方部件可以进行增加、删除、修改操作。"
      Height          =   180
      Left            =   945
      TabIndex        =   15
      Top             =   705
      Width           =   5130
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   225
      Picture         =   "frmUpdateFile.frx":1124
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "frmUpdateFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const StopColor = vbRed '禁用时的颜色
Const StartColor = &H80000008 '启用时的颜色
Dim mintColumn As Integer '

Private mobjFindKey             As CommandBarPopup      '查询
Private mstrFindKey             As String               '查询串
Private m_strCurTypeName        As String               '当前选中的方式
Private m_strCurFileName        As String               '当前选中的名称
Private m_strCurVision          As String               '当前选中的版本
Private m_strCurEditDate        As String               '当前选中的修改日期
Private m_strCurSysNum          As String               '当前选中的系统
Private m_strCurSetupPath       As String               '当前选中的安装路径
Private m_strCurSysOption       As String               '当前选中的系统参数
Private m_strCurFileExplanation As String               '当前选中的文件说明
Private m_strCurSellFile        As String               '当前选中的引用文件
Private m_blnCurReg             As Boolean              '当前选中的文件是否注册
Private m_blnCurUpData          As Boolean              '当前选中的文件是否强制覆盖

Private m_lngCurRow             As Long
Dim mrsTemp      As New ADODB.Recordset

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
    
    objPrint.Title.Text = "升级文件清单"
    
  
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印时间：" & Format(date, "yyyy年MM月dd日")
    Set objPrint.Body = Me.fgMain
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

Private Sub cboSystem_Click()
    Call refData
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


Private Sub cmdFind_Click()
    txtFind_KeyPress 13
End Sub

Private Sub cmdRefresh_Click()
     Call refData '刷新
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        cmdRefresh_Click
    End If
    If KeyCode = vbKeyDelete Then
        If cmdDel.Enabled Then
            cmdDel_Click
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim lngWdt As Single

    err = 0
    On Error Resume Next
    lblNote.Width = ScaleWidth - lblNote.Left
    With cmdRefresh
        .Top = ScaleHeight - .Height - 50
    End With

    picPane(0).Move ScaleWidth - picPane(0).Width - 30, picPane(0).Top

    picPane(1).Move picPane(1).Left, picPane(1).Top, ScaleWidth - 300, cmdRefresh.Top - picPane(1).Top - 50


    With cmdAdd
        .Top = cmdRefresh.Top
        .Left = ScaleWidth - cmdAdd.Width * 3 - 30
    End With


    With cmdEdit
        .Top = cmdRefresh.Top
        .Left = cmdAdd.Left + cmdAdd.Width
    End With

    With cmdDel
        .Top = cmdRefresh.Top
        .Left = cmdEdit.Left + cmdEdit.Width
    End With

End Sub


'==============================================================================
'=功能： 窗口初始化
'==============================================================================
Private Sub Form_Load()
  On Error GoTo errH
    
    KeyPreview = True
    m_lngCurRow = -1
    '查找框初始化
    txtFind.Text = "请输入文件名称"
    txtFind.ForeColor = vbGrayText
    '填充Combo
    Call InitComBo

'    Call SetMenu
    Exit Sub
    
errH:
    MsgBox err.Description, vbInformation, "提示"
End Sub


'==============================================================================
'=功能： 网格fgMain单击后刷新状态信息
'==============================================================================
Private Sub fgMain_Click()
    On Error GoTo errH
    fgMain_SelChange
    Exit Sub
errH:
   
End Sub

'==============================================================================
'=功能： 右键菜单 fgMain
'==============================================================================
Private Sub fgMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    On Error GoTo errH
'    Select Case Button
'        Case 2          '弹出菜单处理
'            Call SendLMouseButton(fgMain.hwnd, X, Y)
'            mcbrPopupBarItem.ShowPopup
'    End Select
'    Exit Sub
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 网格行列变化时更新基本信息
'==============================================================================
Private Sub fgMain_RowColChange()
    On Error GoTo errH
    Call fgMain_SelChange
    Exit Sub
errH:

End Sub

'==============================================================================
'=功能： 网格选择行列变化时更新基本信息
'==============================================================================
Private Sub fgMain_SelChange()
    Dim lngID       As Long
    On Error GoTo errH
    
'    fgMain.WallPaper = imgBG_fg(1).Picture
    m_lngCurRow = fgMain.Row
    If m_lngCurRow = 0 Then Exit Sub
    m_strCurTypeName = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 1)) = 0, 0, fgMain.Cell(flexcpText, m_lngCurRow, 1))   '获取ID
    m_strCurFileName = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 2)) = 0, 0, fgMain.Cell(flexcpText, m_lngCurRow, 2))     '文件名
    m_strCurVision = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 3)) = 0, 0, fgMain.Cell(flexcpText, m_lngCurRow, 3))
    m_strCurEditDate = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 4)) = 0, 0, fgMain.Cell(flexcpText, m_lngCurRow, 4))
    m_strCurSysNum = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 5)) = 0, "", fgMain.Cell(flexcpText, m_lngCurRow, 5))
    m_strCurSellFile = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 6)) = 0, "", fgMain.Cell(flexcpText, m_lngCurRow, 6))
    m_strCurSetupPath = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 7)) = 0, "", fgMain.Cell(flexcpText, m_lngCurRow, 7))
    m_strCurSysOption = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 10)) = 0, "", fgMain.Cell(flexcpText, m_lngCurRow, 10))
    m_blnCurReg = IIf(fgMain.Cell(flexcpText, m_lngCurRow, 11) = "是", True, False) '自动注册
    m_strCurFileExplanation = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 12)) = 0, "", fgMain.Cell(flexcpText, m_lngCurRow, 12)) '文件说明
    m_blnCurUpData = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 13)) = 0, False, fgMain.Cell(flexcpText, m_lngCurRow, 13)) '强制覆盖
    
    If m_strCurTypeName = "三方部件" Then
        cmdEdit.Enabled = True
        cmdDel.Enabled = True
    Else
        cmdEdit.Enabled = False
        cmdDel.Enabled = False
    End If
    
    Call SetMenu
    Exit Sub
errH:
    If False Then
        Resume
    End If
End Sub

Private Sub fgMain_DblClick()
    If m_strCurTypeName = "三方部件" Then
        Call StandardEdit
    End If
End Sub

'==============================================================================
'=功能： 填充系统 ComBo
'==============================================================================
Private Sub InitComBo()
    On Error GoTo errH
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim lngDefaultNum As Long
    Dim str编号       As String
    With cboSystem
    .Clear
    strSQL = "select 编号,名称,共享号 from zlSystems order by 编号"
    Call OpenRecordset(rs, strSQL, Me.Caption)
    If rs.BOF = False Then
        rs.MoveFirst
        .AddItem "[0]所有系统"
        .ItemData(.NewIndex) = 0
        Do While Not rs.EOF
            str编号 = rs("编号").value \ 100
            .AddItem "[" & str编号 & "]" & rs("名称").value
            .ItemData(.NewIndex) = str编号
            If Nvl(rs("共享号").value, 0) = 0 Then
                lngDefaultNum = .ListCount - 1
            End If
            rs.MoveNext
        Loop
    End If
    .ListIndex = 0 'lngDefaultNum
    End With
    Exit Sub
errH:

End Sub

'==============================================================================
'=功能： 装入对应方案的评分标准
'==============================================================================
Public Sub DataLoad(Optional ByVal strFilter As String, Optional ByVal strLocationName As String)

    Dim i, j As Long
    Dim strSQL       As String
    Dim strSystemNum As String
    Dim strTypeID()  As String
    Dim strTemp      As String
    Dim arrSys         As Variant
    On Error GoTo errH
    
    With fgMain
        .Tag = ""
'        .Redraw = flexRDNone
        .Rows = 1
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
        
        If CheckTable = False Then
            Exit Sub
        End If
        
        If Len(strFilter) <> 0 Then
            If strFilter = "Clear" Then
                Exit Sub
            Else
                strSystemNum = cboSystem.ItemData(cboSystem.ListIndex)
                If strSystemNum = "" Then strSystemNum = "1"
                
                If strSystemNum = "0" Then
                     strSQL = "Select a.序号,a.文件类型 As 类型ID,Decode(a.文件类型, 0, '公共部件', 1, '应用部件', 2, '帮助文件', 3, '其它文件', 4, '三方部件', 5, '系统文件', '未知类型') As 文件类型, a.文件名, a.版本号, a.修改日期," & vbNewLine & _
                             "       a.所属系统, a.业务部件,a.安装路径,a.文件说明,a.自动注册" & vbNewLine & _
                             "From zlFilesUpgrade A" & vbNewLine & _
                             "Where a.文件类型 In (" & strFilter & ") order by lpad(a.序号,5,'0')"
                             
                              Call OpenRecordset(mrsTemp, strSQL, Me.Caption)
                              GoTo zt
                End If
                
                If InStrRev(strFilter, "0") > 0 Then
                   strTypeID = Split(strFilter, ",")
                   For i = 0 To UBound(strTypeID)
                        If strTemp = "" Then
                            strTemp = strTypeID(i)
                        Else
                            strTemp = strTemp & "," & strTypeID(i)
                        End If
                   Next
                    strSQL = "Select B.序号,B.类型ID,B.文件类型,B.文件名,B.版本号,B.修改日期,B.所属系统,B.业务部件,B.安装路径,B.文件说明,B.自动注册 From ( " & vbNewLine & _
                                "Select a.序号,a.文件类型 As 类型ID,Decode(a.文件类型, 0, '公共部件', 1, '应用部件', 2, '帮助文件', 3, '其它文件', 4, '三方部件', 5, '系统文件', '未知类型') As 文件类型, a.文件名, a.版本号, a.修改日期," & vbNewLine & _
                                "       a.所属系统, a.业务部件,a.安装路径,a.文件说明,a.自动注册" & vbNewLine & _
                                "From zlFilesUpgrade A" & vbNewLine & _
                                "Where a.文件类型 In (" & strTemp & ") And (Instr(a.所属系统, ','|| " & strSystemNum & " ||  ',' ) > 0 or a.所属系统 is null )" & vbNewLine & _
                                "Union" & vbNewLine & _
                                "Select a.序号,a.文件类型 As 类型ID,Decode(a.文件类型, 0, '公共部件', 1, '应用部件', 2, '帮助文件', 3, '其它文件', 4, '三方部件', 5, '系统文件', '未知类型') As 文件类型, a.文件名, a.版本号, a.修改日期," & vbNewLine & _
                                "       a.所属系统, a.业务部件,a.安装路径,a.文件说明,a.自动注册" & vbNewLine & _
                                "From zlFilesUpgrade A" & vbNewLine & _
                                "Where a.文件类型 =0" & vbNewLine & _
                                ") B Order by lpad(B.序号,5,'0')"

                    Call OpenRecordset(mrsTemp, strSQL, Me.Caption)
                Else
                    strSQL = "Select a.序号,a.文件类型 As 类型ID,Decode(a.文件类型, 0, '公共部件', 1, '应用部件', 2, '帮助文件', 3, '其它文件', 4, '三方部件', 5, '系统文件', '未知类型') As 文件类型, a.文件名, a.版本号, a.修改日期," & vbNewLine & _
                             "       a.所属系统, a.业务部件,a.安装路径,a.文件说明,a.自动注册" & vbNewLine & _
                             "From zlFilesUpgrade A" & vbNewLine & _
                             "Where a.文件类型 In (" & strFilter & ") And (Instr(a.所属系统, ',' || " & strSystemNum & " || ',' ) > 0 or a.所属系统 is null ) order by lpad(a.序号,5,'0')"
                    
             
                        Call OpenRecordset(mrsTemp, strSQL, Me.Caption)
                End If
            End If
        Else
            strSystemNum = cboSystem.ItemData(cboSystem.ListIndex)
            If strSystemNum = "" Then strSystemNum = "100"
    
            strSQL = "Select B.序号,B.类型ID,B.文件类型,B.文件名,B.版本号,B.修改日期,B.所属系统,B.业务部件,B.安装路径,B.文件说明,B.自动注册 From ( " & vbNewLine & _
                        "Select a.序号,a.文件类型 As 类型ID,Decode(a.文件类型, 0, '公共部件', 1, '应用部件', 2, '帮助文件', 3, '其它文件', 4, '三方部件', 5, '系统文件', '未知类型') As 文件类型, a.文件名, a.版本号, a.修改日期," & vbNewLine & _
                         "       a.所属系统, a.业务部件,a.安装路径,a.文件说明,a.自动注册" & vbNewLine & _
                         "From zlFilesUpgrade A" & vbNewLine & _
                         "Where a.文件类型 In (1, 2, 3,4) And (Instr(a.所属系统,  ',' ||  " & strSystemNum & " || ',') > 0 or a.所属系统 is null )" & vbNewLine & _
                         "Union" & vbNewLine & _
                         "Select a.序号,a.文件类型 As 类型ID,Decode(a.文件类型, 0, '公共部件', 1, '应用部件', 2, '帮助文件', 3, '其它文件', 4, '三方部件', 5, '系统文件', '未知类型') As 文件类型, a.文件名, a.版本号, a.修改日期," & vbNewLine & _
                         "       a.所属系统, a.业务部件,a.安装路径,a.文件说明,a.自动注册" & vbNewLine & _
                         "From zlFilesUpgrade A" & vbNewLine & _
                         "Where a.文件类型 =0" & vbNewLine & _
                         ") B Order by lpad(B.序号,5,'0')"
        
            Call OpenRecordset(mrsTemp, strSQL, Me.Caption)
        End If
zt:
'    .AllowSelection = False '对齐
'    .Editable = flexEDKbdMouse
'    .AllowUserResizing = flexResizeBoth
'    .AllowUserFreezing = flexFreezeBoth
'    .BackColorFrozen = 14737632
'    .GridLines = flexGridFlatVert
    .ExtendLastCol = True
'    .ScrollTips = True
    
        .FocusRect = flexFocusSolid
        '数据填入
        .Rows = mrsTemp.RecordCount + 1
    
        i = 1
        Do Until mrsTemp.EOF
            .Cell(flexcpText, i, 0) = Nvl(mrsTemp.Fields("序号"), 0) 'mrsTemp.AbsolutePosition
            .Cell(flexcpAlignment, i, 0) = flexAlignLeftCenter
            
            
            .Cell(flexcpText, i, 1) = Nvl(mrsTemp.Fields("文件类型"))
            .Cell(flexcpAlignment, i, 1) = flexAlignCenterCenter
'            If NVL(mrsTemp.Fields("文件类型")) = "应用部件" Then
'                .Cell(flexcpBackColor, i, 1) = &H80C0FF   '&H8080FF
'            End If
            .Cell(flexcpText, i, 2) = Nvl(mrsTemp.Fields("文件名"))
            .Cell(flexcpAlignment, i, 2) = flexAlignLeftCenter
            
            strTemp = Nvl(mrsTemp.Fields("版本号"))
            strTemp = GetFileVision(strTemp)
            
            .Cell(flexcpText, i, 3) = strTemp
            .Cell(flexcpAlignment, i, 3) = flexAlignCenterCenter
            
            If Nvl(mrsTemp.Fields("修改日期")) <> "" Then
                strTemp = Format(Nvl(mrsTemp.Fields("修改日期")), "yyyy-mm-dd hh:mm:ss")
            Else
                strTemp = ""
            End If
            
            .Cell(flexcpText, i, 4) = strTemp
            .Cell(flexcpAlignment, i, 4) = flexAlignCenterCenter
            
            strTemp = Nvl(mrsTemp.Fields("所属系统"))

            If Trim(strTemp) <> "" Then
                arrSys = Split(Trim(strTemp), ",")
                strTemp = ""
                For j = 0 To UBound(arrSys)
                    If GetSystemName(arrSys(j)) <> "" Then strTemp = strTemp & "，" & GetSystemName(arrSys(j))
                Next
                strTemp = Mid(strTemp, 2)
            Else
                strTemp = "所有系统"
            End If

            .Cell(flexcpText, i, 5) = strTemp
            .Cell(flexcpAlignment, i, 5) = flexAlignLeftCenter
            
            .Cell(flexcpText, i, 6) = Nvl(mrsTemp.Fields("业务部件"))
            .Cell(flexcpAlignment, i, 6) = flexAlignLeftCenter
            
            .Cell(flexcpText, i, 7) = Nvl(mrsTemp.Fields("安装路径"))
            .Cell(flexcpAlignment, i, 7) = flexAlignLeftCenter
            
            .Cell(flexcpText, i, 8) = Nvl(mrsTemp.Fields("类型ID"))
            .Cell(flexcpAlignment, i, 8) = flexAlignLeftTop
            
            .Cell(flexcpText, i, 9) = Nvl(mrsTemp.Fields("安装路径"))
            .Cell(flexcpAlignment, i, 9) = flexAlignLeftCenter
            
            .Cell(flexcpText, i, 10) = Nvl(mrsTemp.Fields("所属系统")) 'NVL(mrsTemp.Fields("系统参数"))
            .Cell(flexcpAlignment, i, 10) = flexAlignCenterCenter
            
            .Cell(flexcpText, i, 11) = IIf(Nvl(mrsTemp.Fields("自动注册"), "") = "1", "是", "否")
            .Cell(flexcpAlignment, i, 11) = flexAlignCenterCenter
            
            .Cell(flexcpText, i, 12) = Nvl(mrsTemp.Fields("文件说明"), "")
            .Cell(flexcpAlignment, i, 12) = flexAlignLeftCenter
            
            .Cell(flexcpText, i, 13) = ""
            .Cell(flexcpAlignment, i, 13) = flexAlignLeftCenter
            
            mrsTemp.MoveNext
            i = i + 1
        Loop

        '自动换行
        .WordWrap = True
        '合并单元格
        .MergeCells = 0
        .MergeCol(.ColIndex("文件类型")) = True
        .MergeCol(.ColIndex("文件名")) = True
        '隐藏单元格
        .ColWidth(.ColIndex("类型ID")) = 0
        
        '行高设置
        .RowHeightMin = 300
        .RowHeightMax = 300
        '最大宽度设置
        .ColWidthMax = 7000
        '自动适应行高、列宽
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize .ColIndex("业务部件")
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .Redraw = flexRDBuffered
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
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
        fgMain_SelChange

        .SetFocus
         Call SetMenu
    End With
    Exit Sub
errH:
    If False Then
        Resume
    End If
End Sub


'==============================================================================
'=功能： 显示记录数信息
'==============================================================================
Private Sub SetMenu()
 
    frmMDIMain.stbThis.Panels(2).Text = "列表中共显示有" & fgMain.Rows - 1 & "行数据。"

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
    If False Then
        Resume
    End If
End Function

Private Sub picPane_Resize(Index As Integer)
    Select Case Index
    Case 0
    Case 1
         fgMain.Move 0, 0, picPane(1).Width - 5, picPane(1).Height - 5
    End Select
End Sub


'==============================================================================
'=功能： 刷新数据
'==============================================================================
Private Sub refData(Optional ByVal strLocationName As String)
    Dim strTemp As String
    On Error GoTo errH
    If chk部件(0).value Then
        strTemp = "0,"
    End If
    
    If chk部件(1).value Then
        If Len(strTemp) = 0 Then
            strTemp = "1,"
        Else
            strTemp = strTemp & "1,"
        End If
    End If
    
    If chk部件(2).value Then
        If Len(strTemp) = 0 Then
            strTemp = "2,"
        Else
            strTemp = strTemp & "2,"
        End If
    End If
    
    If chk部件(3).value Then
        If Len(strTemp) = 0 Then
            strTemp = "3,"
        Else
            strTemp = strTemp & "3,"
        End If
    End If
    
    If chk部件(4).value Then
        If Len(strTemp) = 0 Then
            strTemp = "4,"
        Else
            strTemp = strTemp & "4,"
        End If
    End If
    
    If chk部件(5).value Then
        If Len(strTemp) = 0 Then
            strTemp = "5"
        Else
            strTemp = strTemp & "5"
        End If
    End If
    
    If Len(strTemp) > 0 Then
        If Right(strTemp, 1) = "," Then
            strTemp = Left(strTemp, Len(strTemp) - 1)
        End If
        Call DataLoad(strTemp, strLocationName)
    Else
        Call DataLoad("Clear")
    End If
    Exit Sub
errH:
End Sub

Private Sub chk部件_Click(Index As Integer)
    On Error GoTo errH
    Call refData
errH:

End Sub


'==============================================================================
'=修改文件
'==============================================================================
Private Sub StandardEdit()
    Dim f As New frmScriptEdit
    Dim strSysNum As String
    On Error GoTo errH
    If cboSystem.Text = "" Then
        strSysNum = 100
    Else
        strSysNum = cboSystem.ItemData(cboSystem.ListIndex)
    End If
    
    f.ShowForm "修改", m_strCurTypeName, m_strCurFileName, m_strCurSysNum, strSysNum, m_strCurVision, m_strCurSetupPath, m_strCurEditDate, m_strCurSysOption, m_strCurFileExplanation, m_strCurSellFile, m_blnCurReg, m_blnCurUpData, "0"
    If f.Moded Then
        Call refData(m_strCurFileName)
        Dim lngRow As Long
        lngRow = fgMain.FindRow(CStr(m_strCurFileName), , 2)
        If lngRow <> -1 Then
              fgMain.Select lngRow, 2
              fgMain.ShowCell lngRow, 2
        End If
    End If
    Exit Sub
errH:
 
End Sub


'==============================================================================
'=新增文件
'==============================================================================
Private Sub StandardAdd()
    Dim f As New frmScriptEdit
    Dim strSysNum As String
    Dim strLocationName As String
    
    On Error GoTo errH
    If cboSystem.Text = "" Then
        strSysNum = 1
    Else
        strSysNum = cboSystem.ItemData(cboSystem.ListIndex)
    End If
    
    strLocationName = f.ShowForm("新增", m_strCurTypeName, m_strCurFileName, m_strCurSysNum, strSysNum, m_strCurVision, m_strCurSetupPath, m_strCurEditDate, m_strCurSysOption, m_strCurFileExplanation, m_strCurSellFile, m_blnCurReg, m_blnCurUpData, "0")
    If f.Moded Then
        Call refData(strLocationName)
        fgMain.Row = fgMain.Rows
        fgMain.Select fgMain.Rows, 1, fgMain.Rows, 1
        fgMain.ro
    End If
    Exit Sub
errH:
  
End Sub

'==============================================================================
'=删除文件
'==============================================================================
Private Sub StandardDel()
    Dim i         As Long
    Dim strName   As String
    Dim lngCurRow As Long
    Dim rs        As ADODB.Recordset
    Dim strSQL    As String
    Dim strSys    As String
    Dim strSysNum As String
    Dim lngRow    As Long
    On Error GoTo errH
    
    If fgMain.SelectedRows = 0 Then Exit Sub
    
    If m_strCurTypeName <> "三方部件" Then
        Exit Sub
    End If
    
    If fgMain.SelectedRows = 1 Then
        If MsgBox("你确认要删除[" & Right(cboSystem.Text, Len(cboSystem.Text) - InStrRev(cboSystem.Text, "]", -1)) & "]" & vbCrLf & "的部件" & m_strCurFileName & "吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        If MsgBox("你确认要删除选择的" & fgMain.SelectedRows & "个部件吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
'    gcnOracle.BeginTrans
    
    
    lngRow = fgMain.FindRow(CStr(m_strCurFileName), , 2)
    
    For i = 0 To fgMain.SelectedRows
        If fgMain.SelectedRow(i) Then
            lngCurRow = fgMain.SelectedRow(i)
            If lngCurRow <> -1 Then
                strName = IIf(Len(fgMain.Cell(flexcpText, lngCurRow, 2)) = 0, 0, fgMain.Cell(flexcpText, lngCurRow, 2))
                strName = UCase(strName)
            
                gstrSQL = "delete zlFilesUpgrade where upper(文件名)= upper('" & strName & "')"
                gcnOracle.Execute gstrSQL
'                End If
            End If

        End If
    Next
    
'    gcnOracle.CommitTrans
    
    ''Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Call refData
    Call SetMenu
    
    
    If lngRow <> -1 Then
        If lngRow >= 2 And fgMain.Rows > 2 Then
          fgMain.Select lngRow - 1, 2
          fgMain.ShowCell lngRow - 1, 2
        End If
    End If
    Exit Sub
errH:
'    gcnOracle.RollbackTrans

End Sub



Private Sub txtFind_GotFocus()
    If txtFind.ForeColor = vbGrayText Then
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
        For lngLoop = fgMain.Row + 1 To fgMain.Rows - 1
            If InStr(UCase(fgMain.TextMatrix(lngLoop, 2)), UCase(txtFind.Text)) > 0 Then
                lngRow = lngLoop
                Exit For
            End If
        Next
        '读取小于当前行的记录数据
        If lngRow = 0 Then
            For lngLoop = 0 To fgMain.Row
                If InStr(UCase(fgMain.TextMatrix(lngLoop, 2)), UCase(txtFind.Text)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
        End If
        If fgMain.Rows > 1 And lngRow >= 1 Then
            fgMain.Row = lngRow
            fgMain.ShowCell lngRow, 2
        End If
        
        
        'Call LocationObj(txtFind)
    End If
    If mstrFindKey = "文件名称" Then mstrFindKey = "名称"

    Exit Sub
errH:
    mstrFindKey = "名称"
    
End Sub


Private Sub txtFind_LostFocus()
    If txtFind.Text = "" Then
        txtFind.Text = "请输入文件名称"
        txtFind.ForeColor = vbGrayText
    End If
End Sub
