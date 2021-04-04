VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmIconManage 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "图标管理"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4785
   Icon            =   "frmIconManage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1380
      ScaleHeight     =   285
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   3315
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   120
      ScaleHeight     =   2865
      ScaleWidth      =   4575
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   330
      Width           =   4575
      Begin VSFlex8Ctl.VSFlexGrid vsIcon 
         Height          =   2550
         Left            =   -15
         TabIndex        =   1
         Top             =   0
         Width           =   4590
         _cx             =   8096
         _cy             =   4498
         Appearance      =   2
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   12
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   360
         RowHeightMax    =   360
         ColWidthMin     =   360
         ColWidthMax     =   360
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
      Begin VB.Label lblFunc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "替换(Ctrl+R)"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   1335
         MouseIcon       =   "frmIconManage.frx":058A
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   2655
         UseMnemonic     =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblFunc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "删除(Ctrl+D)"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   2
         Left            =   2565
         MouseIcon       =   "frmIconManage.frx":06DC
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   2655
         UseMnemonic     =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblFunc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "增加(Ctrl+N)"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   90
         MouseIcon       =   "frmIconManage.frx":082E
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   2655
         UseMnemonic     =   0   'False
         Width           =   1080
      End
   End
   Begin XtremeSuiteControls.TabControl tbcIcon 
      Height          =   3210
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4605
      _Version        =   589884
      _ExtentX        =   8123
      _ExtentY        =   5662
      _StockProps     =   64
   End
   Begin MSComDlg.CommonDialog cdgIcon 
      Left            =   495
      Top             =   3210
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmIconManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum IDX_FUNC
    FUNC_新增 = 0
    FUNC_替换 = 1
    FUNC_删除 = 2
End Enum
Private Type ICON_CONTENT
    ID As Long
    图标 As StdPicture
End Type
Private marrFixed() As ICON_CONTENT
Private marrCustom() As ICON_CONTENT
Private mblnOK As Boolean

Public Function ShowMe(frmMain As Object) As Boolean
    Me.Show 1, frmMain
    ShowMe = mblnOK
End Function

Private Sub Form_Activate()
    vsIcon.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyN And Shift = vbCtrlMask Then
        If lblFunc(FUNC_新增).Enabled Then Call lblFunc_Click(FUNC_新增)
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If lblFunc(FUNC_替换).Enabled Then Call lblFunc_Click(FUNC_替换)
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then
        If lblFunc(FUNC_删除).Enabled Then Call lblFunc_Click(FUNC_删除)
    End If
End Sub

Private Sub Form_Load()
    mblnOK = False
        
    picIcon.BackColor = Me.BackColor
    vsIcon.Cell(flexcpPictureAlignment, 0, 0, vsIcon.Rows - 1, vsIcon.Cols - 1) = 4
        
    '数组数据从下标1开始有效
    ReDim marrFixed(0): ReDim marrCustom(0)
        
    'TabControl
    '-----------------------------------------------------
    With tbcIcon
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
        End With
        'Tag存放存放当前显示页数，缺省为第1页
        .InsertItem(0, "自定义", picIcon.Hwnd, 0).Tag = "1"
        .InsertItem(1, "系统固有", picIcon.Hwnd, 0).Tag = "1"

        '因为绑定相同,最后要切换回第1个
        .Item(.ItemCount - 1).Selected = True
        .Item(0).Selected = True
    End With
    
    Call LoadIcons
    Call ShowIcons
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase marrFixed
    Erase marrCustom
End Sub

Private Sub lblFunc_Click(Index As Integer)
    Dim objIcon As StdPicture
    Dim arrSQL() As String, strSql As String
    Dim lng图标ID As Long, strFile As String
    Dim blnTran As Boolean, i As Long
    Dim lngIdx As Long
    
    Select Case Index
    Case FUNC_新增, FUNC_替换
        If Index = FUNC_新增 Then
            cdgIcon.DialogTitle = "选择要添加的图标文件"
        Else
            cdgIcon.DialogTitle = "选择要替换的图标文件"
        End If
        cdgIcon.Filter = "所有图像文件|*.ico;*.bmp;*.gif;*.jpg|ICON 图标(*.ico)|*.ico|位图图像(*.bmp)|*.bmp|GIF 图像(*.gif)|*.gif|JPEG 图像(*.jpg)|*.jpg|所有文件|*.*"
        cdgIcon.Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        cdgIcon.InitDir = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "临床路径图标目录")
        cdgIcon.CancelError = True
        On Error Resume Next
        cdgIcon.ShowOpen
        If Err.Number <> 0 Then
            Err.Clear: Exit Sub
        End If
        On Error GoTo 0
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "临床路径图标目录", gobjFile.GetFile(cdgIcon.FileName).ParentFolder.Path
        
        '检查是否支持
        On Error Resume Next
        Set objIcon = LoadPicture(cdgIcon.FileName)
        If Err.Number <> 0 Then
            MsgBox "不支持的图像格式。", vbInformation, gstrSysName
            Err.Clear: Exit Sub
        End If
        On Error GoTo 0
        Screen.MousePointer = 11
        
        '缩放成16*16的BMP图像
        picTemp.Width = Me.ScaleX(16, vbPixels, vbTwips)
        picTemp.Height = Me.ScaleY(16, vbPixels, vbTwips)
        picTemp.PaintPicture objIcon, 0, 0, picTemp.Width, picTemp.Height
        Set objIcon = picTemp.Image
        
        '保存数据库
        strFile = gobjFile.GetSpecialFolder(TemporaryFolder) & "\zlTemplate.bmp"
        Call SavePicture(objIcon, strFile)
        
        ReDim arrSQL(0)
        If Index = FUNC_新增 Then
            lng图标ID = zlDatabase.GetNextID("临床路径图标")
            arrSQL(0) = "Zl_临床路径图标_Insert(" & lng图标ID & ")"
        Else
            lng图标ID = vsIcon.Cell(flexcpData, vsIcon.Row, vsIcon.Col)
        End If
        If Not Sys.GetlobSql(glngSys, 11, lng图标ID, strFile, arrSQL()) Then
            If gobjFile.FileExists(strFile) Then Call gobjFile.DeleteFile(strFile, True)
            MsgBox "图标" & IIf(Index = FUNC_新增, "增加", "替换") & "失败！", vbExclamation, gstrSysName
            Screen.MousePointer = 0: Exit Sub
        End If
        If gobjFile.FileExists(strFile) Then Call gobjFile.DeleteFile(strFile, True)
                
        '执行SQL语句
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTran = True
        For i = LBound(arrSQL) To UBound(arrSQL)
            If arrSQL(i) <> "" Then
                Call zlDatabase.ExecuteProcedure(arrSQL(i), Me.Caption)
            End If
        Next
        gcnOracle.CommitTrans: blnTran = False
        On Error GoTo 0
        
        '刷新显示
        If Index = FUNC_新增 Then
            ReDim Preserve marrCustom(UBound(marrCustom) + 1)
            lngIdx = UBound(marrCustom)
            marrCustom(lngIdx).ID = lng图标ID
        Else
            lngIdx = vsIcon.Cols * vsIcon.Row + vsIcon.Col + 1
        End If
        Set marrCustom(lngIdx).图标 = picTemp.Image
        picTemp.Cls '如果不Cls，或者使用objIcon，将影响前面索引图标的内容
        
        Call ShowIcons(lng图标ID)
        
        mblnOK = True
        Set gcolIcons = Nothing
        Screen.MousePointer = 0
    Case FUNC_删除
        If MsgBox("确实要删除当前选中的图标吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        lng图标ID = vsIcon.Cell(flexcpData, vsIcon.Row, vsIcon.Col)
        
        On Error GoTo errH
        zlDatabase.ExecuteProcedure "Zl_临床路径图标_Delete(" & lng图标ID & ")", Me.Caption
        On Error GoTo 0
        
        lngIdx = vsIcon.Cols * vsIcon.Row + vsIcon.Col + 1
        For i = lngIdx To UBound(marrCustom) - 1
            marrCustom(i).ID = marrCustom(i + 1).ID
            Set marrCustom(i).图标 = marrCustom(i + 1).图标
        Next
        ReDim Preserve marrCustom(UBound(marrCustom) - 1)
        
        Call ShowIcons(lng图标ID)
        
        mblnOK = True
        Set gcolIcons = Nothing
    End Select
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tbcIcon_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call ShowIcons
    If Visible Then vsIcon.SetFocus
End Sub

Private Function LoadIcons() As Boolean
'功能：读取所有图标到内存
    Dim rsTmp As ADODB.Recordset
    Dim objIcon As StdPicture
    Dim strFile As String, strSql As String
    
    On Error GoTo errH
        
    Screen.MousePointer = 11
    
    strFile = gobjFile.GetSpecialFolder(TemporaryFolder) & "\zlTemplate.bmp"
    
    strSql = "Select ID,性质 From 临床路径图标 Order by 性质,ID"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
    Do While Not rsTmp.EOF
        If Sys.Readlob(glngSys, 11, rsTmp!ID, strFile) <> "" Then
            If Nvl(rsTmp!性质, 0) = 1 Then
                ReDim Preserve marrFixed(UBound(marrFixed) + 1)
                marrFixed(UBound(marrFixed)).ID = rsTmp!ID
                Set marrFixed(UBound(marrFixed)).图标 = LoadPicture(strFile)
            Else
                ReDim Preserve marrCustom(UBound(marrCustom) + 1)
                marrCustom(UBound(marrCustom)).ID = rsTmp!ID
                Set marrCustom(UBound(marrCustom)).图标 = LoadPicture(strFile)
            End If
            
            gobjFile.DeleteFile strFile
        End If
        rsTmp.MoveNext
    Loop
    
    Screen.MousePointer = 0
    LoadIcons = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowIcons(Optional ByVal lngIconID As Long)
'功能：根据当前选择的页面内容显示对应的图标
    Dim arrIcon() As ICON_CONTENT
    Dim lngRow As Long, lngCol As Long
    Dim i As Long
    
    If tbcIcon.Selected.Index = 0 Then
        arrIcon = marrCustom
    Else
        arrIcon = marrFixed
    End If
    
    With vsIcon
        .Redraw = flexRDNone
        .Rows = IntEx(UBound(arrIcon) / .Cols)

        If .Rows > 0 Then
            .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = Empty
            .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
            Set .Cell(flexcpPicture, 0, 0, .Rows - 1, .Cols - 1) = Nothing
            .Row = 0: .Col = 0
        End If
        
        lngRow = 0: lngCol = 0
        For i = 1 To UBound(arrIcon)
            .Cell(flexcpData, lngRow, lngCol) = arrIcon(i).ID
            Set .Cell(flexcpPicture, lngRow, lngCol) = arrIcon(i).图标
            
            If arrIcon(i).ID = lngIconID Then
                .Row = lngRow: .Col = lngCol
            End If
            
            If lngCol + 1 <= .Cols - 1 Then
                lngCol = lngCol + 1
            ElseIf lngRow + 1 <= .Rows - 1 Then
                lngRow = lngRow + 1
                lngCol = 0
            Else
                Exit For
            End If
        Next
        
        If .Rows > 0 Then .ShowCell .Row, .Col
        .Redraw = flexRDDirect
    End With
    
    Call SetFuncEnabled
End Sub

Private Sub SetFuncEnabled()
'功能：根据数据和界面设置功能的可用性
    Dim blnEnabled As Boolean
    
    lblFunc(FUNC_新增).Enabled = tbcIcon.Selected.Index = 0
    
    blnEnabled = tbcIcon.Selected.Index = 0
    If blnEnabled Then blnEnabled = vsIcon.Rows > 0
    If blnEnabled Then blnEnabled = vsIcon.Cell(flexcpData, vsIcon.Row, vsIcon.Col) > 0
    lblFunc(FUNC_替换).Enabled = blnEnabled
    lblFunc(FUNC_删除).Enabled = blnEnabled
End Sub

Private Sub vsIcon_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call SetFuncEnabled
End Sub

Private Sub vsIcon_DblClick()
    If Between(vsIcon.MouseCol, 0, vsIcon.Cols - 1) _
        And Between(vsIcon.MouseRow, 0, vsIcon.Rows - 1) Then
        Call vsIcon_KeyPress(13)
    End If
End Sub

Private Sub vsIcon_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsIcon.Rows > 0 And KeyCode = vbKeyDelete Then
        If lblFunc(FUNC_删除).Enabled Then
            Call lblFunc_Click(FUNC_删除)
        End If
    End If
End Sub

Private Sub vsIcon_KeyPress(KeyAscii As Integer)
    If vsIcon.Rows > 0 And KeyAscii = 13 Then
        KeyAscii = 0
        If lblFunc(FUNC_替换).Enabled Then
            Call lblFunc_Click(FUNC_替换)
        End If
    End If
End Sub
