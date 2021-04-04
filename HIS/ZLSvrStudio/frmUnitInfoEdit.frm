VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmUnitInfoEdit 
   BackColor       =   &H80000005&
   Caption         =   "医院信息维护"
   ClientHeight    =   6450
   ClientLeft      =   6525
   ClientTop       =   3510
   ClientWidth     =   10725
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmUnitInfoEdit.frx":0000
   ScaleHeight     =   6450
   ScaleWidth      =   10725
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboStation 
      Height          =   300
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   810
      Width           =   1125
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   9120
      ScaleHeight     =   1800
      ScaleWidth      =   1800
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "刷新(&R)"
      Height          =   350
      Left            =   8160
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1100
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   840
      ScaleHeight     =   3855
      ScaleWidth      =   8055
      TabIndex        =   2
      Top             =   1200
      Width           =   8055
      Begin VSFlex8Ctl.VSFlexGrid vsUnitInfo 
         Height          =   3615
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   7800
         _cx             =   13758
         _cy             =   6376
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
         BackColorSel    =   16761024
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   4
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   300
         RowHeightMax    =   1500
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmUnitInfoEdit.frx":04F9
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   0   'False
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
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
   End
   Begin VB.CommandButton cmdItemsDelete 
      Caption         =   "删除项目(&D)"
      Height          =   350
      Left            =   6600
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdItemsModify 
      Caption         =   "调整项目(&M)"
      Height          =   350
      Left            =   5160
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdItemsNew 
      Caption         =   "新增项目(&N)"
      Height          =   350
      Left            =   3600
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cdgPub 
      Left            =   3000
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "院区"
      Height          =   180
      Left            =   5160
      TabIndex        =   9
      Top             =   870
      Width           =   360
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "提供医院各类公共信息的定义和内容编辑的功能。"
      Height          =   180
      Left            =   960
      TabIndex        =   1
      Top             =   870
      Width           =   3960
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   240
      Picture         =   "frmUnitInfoEdit.frx":05D4
      Top             =   600
      Width           =   480
   End
   Begin VB.Image imgButtonDel 
      Height          =   240
      Left            =   2280
      Picture         =   "frmUnitInfoEdit.frx":0C6A
      Top             =   6000
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgButtonEdit 
      Height          =   240
      Left            =   2520
      Picture         =   "frmUnitInfoEdit.frx":74BC
      Top             =   6000
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "医院信息维护"
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
      TabIndex        =   0
      Top             =   105
      Width           =   1440
   End
   Begin VB.Menu mnuPop 
      Caption         =   "弹出菜单"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuPopNew 
         Caption         =   "新增项目"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuPopModfy 
         Caption         =   "调整项目"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuPopDel 
         Caption         =   "删除项目"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "frmUnitInfoEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum UnitCol
    Col_编码 = 0
    Col_项目 = 1
    Col_是否图片 = 2
    Col_内容 = 3
    Col_Edit = 4
    Col_Del = 5
    Col_存在数据 = 6
End Enum
Private mstrStation As String '站点
'===========================================================================
'==公共接口
'===========================================================================
Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口
End Sub

'===========================================================================
'==事件
'===========================================================================
Private Sub cboStation_Click()
    Dim strCurStation As String
    strCurStation = cboStation.ItemData(cboStation.ListIndex)
    If strCurStation = "-1" Then strCurStation = ""
    If strCurStation <> mstrStation And cboStation.Tag <> "" Then
        mstrStation = strCurStation
        Call RefreshData
    End If
End Sub

Private Sub cmdItemsDelete_Click()
    Dim strSQL As String
    Dim strRemarks As String
    
    With vsUnitInfo
        If .TextMatrix(.Row, Col_存在数据) <> "1" Then
            If MsgBox("确认要删除""" & .TextMatrix(.Row, Col_项目) & """吗？", vbInformation + vbYesNo, gstrSysName) = vbNo Then
                Exit Sub
            End If
        Else
            If MsgBox("项目""" & .TextMatrix(.Row, Col_项目) & """已经可能被使用，确认要删除吗？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        '验证身份并输入操作说明
        If Not CheckAuditStatus("0312", "删除项目", strRemarks) Then Exit Sub
        On Error GoTo ErrH
        strSQL = "Zltools.b_Public.Zlunitinfoitemchange(2,'" & .TextMatrix(.Row, Col_编码) & "')"
        Call ExecuteProcedure(strSQL, Me.Caption, gcnOracle)
        '插入重要操作日志
        Call SaveAuditLog(3, "删除项目", .TextMatrix(.Row, Col_项目), strRemarks)
        .RemoveItem .Row
        Call SetChange
    End With
    Exit Sub
ErrH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdItemsModify_Click()
    Dim intType As Integer, strName As String, strNo As String
    With vsUnitInfo
        strNo = .TextMatrix(.Row, Col_编码)
        strName = .TextMatrix(.Row, Col_项目)
        If frmUnitItemEdit.ShowMe(strNo, strName, intType) Then
            If Val(.TextMatrix(.Row, Col_是否图片)) <> intType Then '修改了类型，则清空数据标记为未改变
                .Redraw = flexRDNone
                .TextMatrix(.Row, Col_存在数据) = "" '标志不存在数据
                .Cell(flexcpData, .Row, Col_内容) = "" '清空图片路径
                .TextMatrix(.Row, Col_内容) = "" '清空文本内容
                Set .Cell(flexcpPicture, .Row, Col_内容) = Nothing '清空图片
            End If
            .TextMatrix(.Row, Col_项目) = strName
            .TextMatrix(.Row, Col_是否图片) = intType
            Call SetChange
        End If
    End With
End Sub

Private Sub cmdItemsNew_Click()
    Dim intType As Integer, strName As String, strNo As String
    With vsUnitInfo
        If frmUnitItemEdit.ShowMe(strNo, strName, intType) Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .TextMatrix(.Row, Col_编码) = strNo
            .TextMatrix(.Row, Col_项目) = strName
            .TextMatrix(.Row, Col_是否图片) = intType
            Call SetChange
        End If
    End With
End Sub

Private Sub cmdRefresh_Click()
    Call RefreshData
End Sub

Private Sub Form_Activate()
    picMain.Refresh
    Me.Refresh
End Sub

Private Sub Form_Load()
    Call LoadStation
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picMain.Height = Me.ScaleHeight - picMain.Top - cmdRefresh.Height - 150
    picMain.Width = Me.ScaleWidth - picMain.Left - 90
    cmdRefresh.Left = Me.ScaleWidth - cmdRefresh.Width - 60
    cmdRefresh.Top = Me.ScaleHeight - cmdRefresh.Height - 90
    Call SetCtrlPosOnLine(False, 0, cmdRefresh, (cmdRefresh.Width + cmdItemsDelete.Width + 60) * -1, cmdItemsDelete, (cmdItemsDelete.Width + cmdItemsModify.Width + 60) * -1, cmdItemsModify, (cmdItemsModify.Width + cmdItemsNew.Width + 60) * -1, cmdItemsNew)
    Call picMain_Resize
End Sub

Private Sub mnuPopDel_Click()
    Call cmdItemsDelete_Click
End Sub

Private Sub mnuPopModfy_Click()
    Call cmdItemsModify_Click
End Sub

Private Sub mnuPopNew_Click()
    Call cmdItemsNew_Click
End Sub

Private Sub picMain_Resize()
    Dim lngWith  As Long, i As Integer
    On Error Resume Next
    With vsUnitInfo
        .Redraw = flexRDNone
        .Height = picMain.ScaleHeight - 15
        .Width = picMain.ScaleWidth - 15
        '保证内容列根据拖动自动扩展
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) Then
                lngWith = lngWith + .ColWidth(i)
            End If
        Next
        lngWith = .Width - (lngWith - .ColWidth(Col_内容))
        .ColWidth(Col_内容) = lngWith - 60
        If VScrollVisible(vsUnitInfo) Then
            .ColWidth(Col_内容) = .ColWidth(Col_内容) - 300
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsUnitInfo_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Integer
    
    cmdItemsModify.Enabled = NewRow > 0
    cmdItemsDelete.Enabled = NewRow > 0
    mnuPopDel.Enabled = NewRow > 0
    mnuPopModfy.Enabled = NewRow > 0
    If NewRow = -1 Or NewCol = -1 Then Exit Sub
    With vsUnitInfo
        .Redraw = flexRDNone
        '清除图片
        For i = .FixedRows To .Rows - 1
            Set .Cell(flexcpPicture, i, Col_Edit) = Nothing
            Set .Cell(flexcpPicture, i, Col_Del) = Nothing
        Next
        .ComboList = ""
        .FocusRect = flexFocusSolid
'        .FocusRect = flexFocusHeavy
        If NewRow >= .FixedRows Then
            Set .CellButtonPicture = Nothing
            '显示图片
            If NewCol = Col_Edit Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonEdit.Picture
                If .TextMatrix(NewRow, Col_存在数据) = "1" Then
                    Set .Cell(flexcpPicture, NewRow, Col_Del) = imgButtonDel.Picture
                End If
            ElseIf NewCol = Col_Del Then
                If .TextMatrix(NewRow, Col_存在数据) = "1" Then
                    .ComboList = "..."
                    .FocusRect = flexFocusNone
                    Set .CellButtonPicture = imgButtonDel.Picture
                End If
                Set .Cell(flexcpPicture, NewRow, Col_Edit) = imgButtonEdit.Picture
            Else
                If .TextMatrix(NewRow, Col_存在数据) = "1" Then
                    Set .Cell(flexcpPicture, NewRow, Col_Del) = imgButtonDel.Picture
                End If
                Set .Cell(flexcpPicture, NewRow, Col_Edit) = imgButtonEdit.Picture
            End If
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsUnitInfo_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col <> Col_项目 And Col <> Col_内容
End Sub

Private Sub vsUnitInfo_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strReturn As String
    Dim vPoint As POINTAPI
    Dim objPic As StdPicture
    Dim mobjEdit As New frmContentEdit
    With vsUnitInfo
        If Col = Col_Del Then
            Call vsUnitInfo_KeyDown(vbKeyDelete, 0)
        ElseIf Col = Col_Edit Then
            If .TextMatrix(.Row, Col_是否图片) = "1" Then
                cdgPub.Filter = "所有图像文件|*.ico;*.bmp;*.gif;*.jpg|ICON 图标(*.ico)|*.ico|位图图像(*.bmp)|*.bmp|GIF 图像(*.gif)|*.gif|JPEG 图像(*.jpg)|*.jpg|所有文件|*.*"
                cdgPub.flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
                cdgPub.InitDir = App.Path
                cdgPub.CancelError = True
                On Error Resume Next
                cdgPub.ShowOpen
                If err.Number <> 0 Then
                    err.Clear: Exit Sub
                End If
                strReturn = cdgPub.FileName
                '检查是否支持
                On Error Resume Next
                Set objPic = LoadPicture(strReturn)
                If err.Number <> 0 Then
                    MsgBox "不支持的图像格式。", vbInformation, gstrSysName
                    err.Clear: Exit Sub
                End If
                If strReturn = "" Then Exit Sub
                .Cell(flexcpData, Row, Col_内容) = strReturn  '存储文件路径
                Call SaveData(Row)
                Call RefreshData(Row)
            Else
                strReturn = .TextMatrix(Row, Col_内容)
                '获取当前位置
                vPoint = GetCoordPos(.hwnd, .CellLeft - .ColWidth(Col_内容), .CellTop + .CellHeight)
                If mobjEdit.ShowMe(frmMDIMain, strReturn, vPoint.x, vPoint.y, , .ColWidth(Col_内容)) Then
                    If strReturn <> .TextMatrix(Row, Col_内容) Then
                        .TextMatrix(Row, Col_内容) = strReturn
                        Call SaveData(Row)
                        Call RefreshData(Row)
                    End If
                End If
                .Col = Col_内容
            End If
        End If
    End With
End Sub

Private Sub vsUnitInfo_Click()
    With vsUnitInfo
        If (.MouseCol = Col_Del Or .MouseCol = Col_Edit) And .MouseRow >= .FixedRows Then
            .Select .MouseRow, .MouseCol
            Call vsUnitInfo_CellButtonClick(.MouseRow, .MouseCol)
        End If
    End With
End Sub

Private Sub vsUnitInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnHave As Boolean
    If KeyCode = vbKeyDelete Then
        With vsUnitInfo
            If .Row >= .FixedRows Then
                '判断是否存在内容
                If .TextMatrix(.Row, Col_存在数据) = "1" Then
                    If MsgBox("确实要清除该行项目信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        If .TextMatrix(.Row, Col_是否图片) = "1" Then
                            .Cell(flexcpData, .Row, Col_内容) = ""  '清空文件路径
                        Else
                            .TextMatrix(.Row, Col_内容) = ""
                        End If
                        Call SaveData(.Row)
                        Call RefreshData(.Row)
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsUnitInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call EnterNextCell
    End If
End Sub

Private Sub vsUnitInfo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu Me.mnuPop, , picMain.Left + vsUnitInfo.Left + vsUnitInfo.CellLeft, picMain.Top + vsUnitInfo.Top + vsUnitInfo.CellTop + vsUnitInfo.CellHeight
    End If
End Sub

Private Sub vsUnitInfo_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col = Col_Del And vsUnitInfo.TextMatrix(Row, Col_存在数据) = "" Or Col = Col_内容
End Sub

'===========================================================================
'==私有方法
'===========================================================================
Private Sub LoadStation()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrH
    cboStation.Clear
    mstrStation = "-999"
    strSQL = "Select 编号, 名称 From Zlnodelist Order By 编号"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    cboStation.AddItem "公共"
    cboStation.ItemData(cboStation.NewIndex) = -1
    Do While Not rsTmp.EOF
        cboStation.AddItem rsTmp!名称 & ""
        cboStation.ItemData(cboStation.NewIndex) = Val(rsTmp!编号 & "")
        rsTmp.MoveNext
    Loop
    cboStation.Tag = "开始刷新"
    cboStation.ListIndex = 0
    Exit Sub
ErrH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub RefreshData(Optional ByVal lngCurRow As Long)
'功能：数据加载
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Integer, lngRow As Long
    Dim strTmp As String, strPreCode As String
    Dim strFile As String, strCode As String
    
    On Error GoTo ErrH
    With vsUnitInfo
        '获取信息列表，以及简单的信息内容
        If lngCurRow < .FixedRows Then
            strSQL = "Select a.编码, a.名称, a.是否图片, b.行号, b.内容" & vbNewLine & _
                    "From Zltools.Zlunitinfoitem a, (Select 行号, 内容, 项目 From Zltools.Zlreginfo Where Nvl(站点, '空空') = Nvl([2], '空空')) b" & vbNewLine & _
                    "Where a.名称 = b.项目(+)" & vbNewLine & _
                    "Order By Lpad(a.编码, 3, '0'), b.行号"
            .Rows = .FixedRows
        Else
            strCode = .TextMatrix(lngCurRow, Col_编码)
            .RowHeight(lngCurRow) = .RowHeightMin '清空数据，不能自动行高，一次回复原始行高
            strSQL = "Select a.编码, a.名称, a.是否图片, b.行号, b.内容" & vbNewLine & _
                    "From Zltools.Zlunitinfoitem a, (Select 行号, 内容, 项目 From Zltools.Zlreginfo Where Nvl(站点, '空空') = Nvl([2], '空空')) b" & vbNewLine & _
                    "Where a.名称 = b.项目(+) and a.编码=[1]" & vbNewLine & _
                    "Order By Lpad(a.编码, 3, '0'), b.行号"
        End If
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, strCode, mstrStation)
        '单行刷新，项目不存在，则刷新所有数据
        If lngCurRow >= .FixedRows And rsTmp.RecordCount = 0 Then
            Call RefreshData
            Exit Sub
        End If
        strPreCode = ""
        Do While Not rsTmp.EOF
            If strPreCode <> rsTmp!编码 Then
                If strPreCode <> "" And strTmp <> "" Then
                    .TextMatrix(lngRow, Col_内容) = strTmp
                    .TextMatrix(lngRow, Col_存在数据) = "1"  '标识存在数据
                End If
                If lngCurRow < .FixedRows Then
                    .Rows = .Rows + 1: lngRow = .Rows - 1
                Else
                    lngRow = lngCurRow
                End If
                strTmp = rsTmp!内容 & "": strPreCode = rsTmp!编码
                .TextMatrix(lngRow, Col_编码) = rsTmp!编码
                .TextMatrix(lngRow, Col_项目) = rsTmp!名称
                .TextMatrix(lngRow, Col_存在数据) = ""
                .Cell(flexcpData, .Row, Col_内容) = ""
                .TextMatrix(lngRow, Col_是否图片) = Val(rsTmp!是否图片 & "")
                '是图片时需要单独读取图片
                If Val(rsTmp!是否图片 & "") = 1 Then
                    strFile = gclsBase.ReadLob(gcnOracle, 0, rsTmp!名称 & "," & mstrStation)
                    If strFile <> "" Then
                        Set .Cell(flexcpPicture, lngRow, Col_内容) = PicDrawPicture(LoadPicture(strFile))
                        .TextMatrix(lngRow, Col_存在数据) = "1" '标识存在数据
                    Else
                        Set .Cell(flexcpPicture, lngRow, Col_内容) = Nothing
                    End If
                End If
            Else
                '文本太长时会多行存储
                strTmp = strTmp & rsTmp!内容
            End If
            rsTmp.MoveNext
        Loop
        If strPreCode <> "" And strTmp <> "" Then
            .TextMatrix(lngRow, Col_内容) = strTmp
            .TextMatrix(lngRow, Col_存在数据) = "1"  '标识存在数据
        End If
    End With
    Call SetChange
    Exit Sub
ErrH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub

Private Sub SetChange()
'功能：设置信息改变的状态
    Dim lngWith  As Long, i As Integer, lngHeight As Long
    
    On Error Resume Next
    With vsUnitInfo
        .Redraw = flexRDNone
'        .Cell(flexcpFontBold, .FixedRows, Col_项目, .Rows - 1, Col_项目) = True
'        .Cell(flexcpForeColor, .FixedRows, Col_项目, .Rows - 1, Col_项目) = &HD2BDB6
        .Cell(flexcpBackColor, .FixedRows, Col_项目, .Rows - 1, Col_项目) = &H8000000F
        '保证内容列根据拖动自动扩展
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) Then
                lngWith = lngWith + .ColWidth(i)
            End If
        Next
        lngWith = .Width - (lngWith - .ColWidth(Col_内容))
        .ColWidth(Col_内容) = lngWith
        .AutoSize (Col_内容)
        .Redraw = flexRDDirect
        If VScrollVisible(vsUnitInfo) Then
            .Redraw = flexRDNone
            .ColWidth(Col_内容) = .ColWidth(Col_内容) - 285
            .AutoSize (Col_内容)
            .Redraw = flexRDDirect
        End If
        If .Row >= .FixedRows Then
            .TopRow = .Row
            .ShowCell .Row, Col_内容
            Call vsUnitInfo_AfterRowColChange(-1, -1, .Row, .Col)
        End If
    End With
End Sub

Private Function SaveData(ByVal lngRow As Long) As Boolean
'功能：进行数据保存。
    Dim i As Integer
    Dim arrSQL() As Variant
    
    On Error GoTo ErrH
    arrSQL = Array()
    With vsUnitInfo
        If lngRow >= .FixedRows Then
            If .TextMatrix(lngRow, Col_是否图片) = "1" Then
                Call gclsBase.GetLobSql(0, .TextMatrix(lngRow, Col_项目) & "," & mstrStation, .Cell(flexcpData, lngRow, Col_内容), arrSQL)
            Else
                Call gclsBase.GetRegInfoSQL(.TextMatrix(lngRow, Col_项目), .TextMatrix(lngRow, Col_内容), mstrStation, arrSQL)
            End If
        End If
    End With
    ShowFlash ("正在保存数据，请稍候！")
    Call gclsBase.ExecuteProcedureBeach(gcnOracle, arrSQL, Me.Caption)
    ShowFlash ("")
    SaveData = True
    Exit Function
ErrH:
    If 0 = 1 Then
        Resume
    End If
    ShowFlash ("")
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Public Function PicDrawPicture(ByRef objPic As StdPicture) As IPictureDisp
'进行图片同比例缩放
    picDraw.AutoRedraw = True '必须这样才能取出画的图形
    picDraw.Cls
    picDraw.Width = picDraw.ScaleHeight * (objPic.Width / objPic.Height)
    On Error Resume Next
    picDraw.PaintPicture objPic, 0, 0, picDraw.ScaleWidth, picDraw.ScaleHeight
    Set PicDrawPicture = picDraw.Image
End Function

Public Sub EnterNextCell()
    Dim i As Long, j As Long
    
    With vsUnitInfo
        '从下一单元开始循环搜索
        If .Row < .FixedRows Then .Row = .FixedRows
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, Col_内容) To Col_Del
                If Not .ColHidden(j) Then
                    If j = Col_Del And .TextMatrix(i, Col_存在数据) = "" Then
                    
                    Else
                        Exit For
                    End If
                End If
            Next
            If j <= Col_Del Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            .ShowCell .Row, .Col
        ElseIf i = .Rows And j > Col_Del Then
            Call PressKey(vbKeyTab)
        End If
    End With
End Sub
