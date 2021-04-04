VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockAnnex 
   BorderStyle     =   0  'None
   Caption         =   "附件"
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8220
   Icon            =   "frmDockAnnex.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picThis 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   1395
      ScaleHeight     =   630
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   1395
   End
   Begin MSComctlLib.ListView lvwThis 
      Height          =   330
      Left            =   1095
      TabIndex        =   0
      Top             =   30
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   582
      View            =   1
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      _Version        =   393217
      Icons           =   "imgThis"
      SmallIcons      =   "imgThis"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "文件"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   5640
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgThis 
      Left            =   645
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   5025
      Top             =   435
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin VB.Label lblThis 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "附件:"
      Height          =   180
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   450
   End
End
Attribute VB_Name = "frmDockAnnex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ICON_SIZE
    ICON_SMALL = 16
    ICON_LARGE = 32
End Enum

'-----------------------------------------------------
'菜单常量
'-----------------------------------------------------
Const conPopuPlay = 101
Const conPopuAdd = 201
Const conPopuDel = 202
Const conPopuPaste = 203

'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mlngRecordId As Long        '病历记录ID
Private mstrPrivs As String         '当前用户的附件权限
Private mblnWrite As Boolean        '是否可增删附件，已经归档的病历不允许删改附件
Private mblnMoved As Boolean        '是否已转储
Private mClipData As Variant        '剪贴板内容
Private mblnDeleted As Boolean      '删除按钮权限

'-----------------------------------------------------
'窗体公共方法
'-----------------------------------------------------
Public Sub zlRefresh(ByVal lngRecordId As Long, Optional ByVal strPrivs As String, Optional ByVal blnMoved As Boolean, Optional ByVal blnDeleted As Boolean)
    '功能：刷新病历附件列表；
    '参数：lngRecordId：电子病历记录ID；strPrivs：当前用户的附件权限
    Dim rsTemp As New ADODB.Recordset
    Dim objItem As ListItem
    Dim objIcon As StdPicture
    
    mblnMoved = blnMoved
    mblnDeleted = blnDeleted
    mlngRecordId = lngRecordId: mstrPrivs = strPrivs
    
    Set Me.lvwThis.Icons = Nothing: Set Me.lvwThis.SmallIcons = Nothing
    Me.lvwThis.ListItems.Clear: Me.imgThis.ListImages.Clear
    
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select Decode(归档人, Null, 1, 0) As 可写 From 电子病历记录 Where ID = [1]"
    If mblnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngRecordId)
    If rsTemp.RecordCount <= 0 Then
        mblnWrite = False
    ElseIf rsTemp.Fields(0).Value = 1 Then
        mblnWrite = True
    Else
        mblnWrite = False
    End If
    
    gstrSQL = "Select 序号, 文件名, 大小, 创建人, 日期 From 电子病历附件 Where 病历id = [1]"
    If mblnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngRecordId)
    With rsTemp
        Do While Not .EOF
            Set objIcon = GetFileIcon(!文件名, ICON_SMALL, True)
            Me.imgThis.ListImages.Add , , objIcon
            Set Me.lvwThis.Icons = Me.imgThis: Set Me.lvwThis.SmallIcons = Me.imgThis
            Set objItem = Me.lvwThis.ListItems.Add(, "_" & !序号, !文件名 & "(" & !大小 & "KB)")
            objItem.Tag = !文件名
            objItem.Icon = Me.imgThis.ListImages.Count: objItem.SmallIcon = objItem.Icon
            .MoveNext
        Loop
        If Me.lvwThis.ListItems.Count > 0 Then Me.lvwThis.ListItems(1).Selected = True
    End With
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetFileIcon(ByVal strFile As String, ByVal intSize As ICON_SIZE, Optional blnUntrue As Boolean) As StdPicture
    '功能：返回指定文件的大图标或小图标
    '说明：需要一个PictureBox控件，无边框，AutoRedraw = True
    '参数： strFile，包含后缀的文件名，当文件真实文件时，应该包含完整的路径名
    '       intSize，获取图标的大小
    '       blnUntrue，非真实文件，这时需要创建文件来获得相关信息
    Dim fInfo As SHFILEINFO
    Dim lngRetu As Long
    
    If blnUntrue Then
        strFile = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp")) & "\" & App.hInstance & CLng(Timer) & strFile
        gobjFSO.CreateTextFile strFile, False
    End If
    If intSize = ICON_LARGE Then
        lngRetu = SHGetFileInfo(strFile, 0, fInfo, Len(fInfo), SHGFI_SHELLICONSIZE Or SHGFI_ICON Or SHGFI_LARGEICON)
    Else
        lngRetu = SHGetFileInfo(strFile, 0, fInfo, Len(fInfo), SHGFI_SHELLICONSIZE Or SHGFI_ICON Or SHGFI_SMALLICON)
    End If
    If blnUntrue Then gobjFSO.DeleteFile strFile, True
    
    Me.picThis.Width = intSize * Screen.TwipsPerPixelX
    Me.picThis.Height = intSize * Screen.TwipsPerPixelY
    Me.picThis.Cls
    If lngRetu <> 0 Then
        DrawIconEx Me.picThis.hDC, 0, 0, fInfo.hIcon, intSize, intSize, 0, 0, DI_NORMAL
        DestroyIcon fInfo.hIcon
    End If
    Set GetFileIcon = Me.picThis.Image
End Function

Private Function GetClipboard(Optional objPic As StdPicture) As Variant
'功能：获取当前剪贴板中的数据
'返回：
'    1.文本：TypeName=String
'    2.多个文件或目录：TypeName=Variant()
'    3.图片：Not objPic is Nothing，TypeName=String(图象路径)或Long(图象句柄)
'    4.无数据：TypeName=Empty
    Dim arrFormat(8) As Long, lngFormat As Long
    Dim strFile As String, i As Long
    Dim hDrop As Long, lngFiles As Long
    Dim arrFile As Variant
    
    Set objPic = Nothing
    GetClipboard = Empty
    
    If CountClipboardFormats = 0 Then Exit Function '无数据
    
    arrFormat(0) = CF_TEXT
    arrFormat(1) = CF_OEMTEXT
    arrFormat(2) = CF_DSPTEXT
    arrFormat(3) = CF_ENHMETAFILE
    arrFormat(4) = CF_METAFILEPICT
    arrFormat(5) = CF_DSPMETAFILEPICT
    arrFormat(6) = CF_BITMAP
    arrFormat(7) = CF_DSPBITMAP
    arrFormat(8) = CF_HDROP
    lngFormat = GetPriorityClipboardFormat(arrFormat(0), 9)
    If lngFormat = 0 Then Exit Function '无数据
    If lngFormat = -1 Then Exit Function '未指定的格式

    Select Case lngFormat
    Case CF_TEXT, CF_OEMTEXT, CF_DSPTEXT '文本
        GetClipboard = Clipboard.GetText()
    Case CF_HDROP '文件/目录
        strFile = Space(MAX_PATH)
        If OpenClipboard(0&) Then
            hDrop = GetClipboardData(CF_HDROP)
            lngFiles = DragQueryFile(hDrop, -1&, "", 0)
            If lngFiles > 0 Then
                ReDim arrFile(lngFiles - 1)
                For i = 0 To lngFiles - 1
                    Call DragQueryFile(hDrop, i, strFile, Len(strFile))
                    arrFile(i) = TrimNull(strFile)
                Next i
                GetClipboard = arrFile
            End If
            Call CloseClipboard
        End If
    Case CF_BITMAP, CF_DSPBITMAP, CF_METAFILEPICT, CF_DSPMETAFILEPICT, CF_ENHMETAFILE '图片格式
        Select Case lngFormat
        Case CF_BITMAP, CF_DSPBITMAP
            Set objPic = Clipboard.GetData(vbCFBitmap)
        Case CF_METAFILEPICT, CF_DSPMETAFILEPICT
            Set objPic = Clipboard.GetData(vbCFMetafile)
        Case CF_ENHMETAFILE
            Dim hEmf As Long, Emh As ENHMETAHEADER
            
            If OpenClipboard(0&) Then
                hEmf = GetClipboardData(CF_ENHMETAFILE)
                Call GetEnhMetaFileHeader(hEmf, Len(Emh), Emh)
                Me.picThis.Cls
                Me.picThis.Width = (Emh.rclBounds.Right - Emh.rclBounds.Left + 1) * Screen.TwipsPerPixelX
                Me.picThis.Height = (Emh.rclBounds.Bottom - Emh.rclBounds.Top + 1) * Screen.TwipsPerPixelY
                Call PlayEnhMetaFile(Me.picThis.hDC, hEmf, Emh.rclBounds)
                Set objPic = Me.picThis.Image
                Call CloseClipboard
            End If
        End Select
        GetClipboard = objPic.Handle
        
        '可能有对应的图片路径
        strFile = Space(MAX_PATH)
        If OpenClipboard(0&) Then
            hDrop = GetClipboardData(CF_HDROP)
            lngFiles = DragQueryFile(hDrop, -1&, "", 0)
            If lngFiles > 0 Then
                Call DragQueryFile(hDrop, 0, strFile, Len(strFile))
                GetClipboard = TrimNull(strFile)
            End If
            Call CloseClipboard
        End If
    End Select
End Function

Private Function TrimNull(ByVal strIn As String) As String
   Dim intNul As Long

   intNul = InStr(strIn, vbNullChar)
   
   Select Case intNul
      Case Is > 1
         TrimNull = Left(strIn, intNul - 1)
      Case 1
         TrimNull = ""
      Case 0
         TrimNull = Trim(strIn)
   End Select
End Function

Private Sub AnnexPlay()
    '功能：打开播放附件
    Dim strFile As String
    Dim varRetu As Variant, strInfo As String
    
    Screen.MousePointer = vbHourglass
    strFile = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp")) & "\" & Me.lvwThis.SelectedItem.Tag
    If gobjFSO.FileExists(strFile) Then gobjFSO.DeleteFile strFile, True
    If zlBlobRead(8, mlngRecordId & "," & Mid(Me.lvwThis.SelectedItem.Key, 2), strFile, mblnMoved) = "" Then
        MsgBox "文件读取失败，请确认附件的有效性！", vbInformation, gstrSysName:
        Screen.MousePointer = vbDefault: Exit Sub
    End If
    varRetu = ShellExecute(Me.hWnd, "open", strFile, "", "", SW_SHOWNORMAL)
    If varRetu <= 32 Then
        Select Case varRetu
        Case 2: strInfo = "错误的关联"
        Case 29: strInfo = "关联失败"
        Case 30: strInfo = "关联应用程式忙碌中..."
        Case 31: strInfo = "没有关联任何应用程式"
        Case Else: strInfo = "无法识别的错误"
        End Select
        MsgBox "附件打开发生：" & strInfo, vbExclamation, gstrSysName
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Function AnnexAdd(ByVal strFile As String) As Boolean
    '功能：选择增加附件文件
    Dim objFile As File
    Dim rsTemp As New ADODB.Recordset, lngMaxNo As Long
    Dim arySql() As String, lngSql As Long, blnTran As Boolean
    
    Err = 0: On Error GoTo errHand
    Set objFile = gobjFSO.GetFile(strFile)
    Me.lblThis.Caption = "正在添加...": Me.lblThis.FontBold = True: Screen.MousePointer = vbHourglass
    
    gstrSQL = "Select Nvl(Max(序号), 0) + 1 As 序号 From 电子病历附件 Where 病历id = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngRecordId)
    lngMaxNo = rsTemp.Fields(0).Value
    
    ReDim arySql(0 To 0)
    arySql(0) = "Zl_电子病历附件_Add(" & mlngRecordId & "," & lngMaxNo & ",'" & objFile.Name & "'," & Format(objFile.Size / 1024, "0.00") & ")"
    If zlBlobSql(8, mlngRecordId & "," & lngMaxNo, strFile, arySql()) = False Then
        MsgBox strFile & "附件保存失败！", vbExclamation, gstrSysName: Exit Function
    End If

    '执行保存
    gcnOracle.BeginTrans: blnTran = True
    For lngSql = LBound(arySql) To UBound(arySql)
        Call zldatabase.ExecuteProcedure(arySql(lngSql), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTran = False
    Me.lblThis.Caption = "附件:": Me.lblThis.FontBold = False: Screen.MousePointer = vbDefault
    AnnexAdd = True
    Exit Function
errHand:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Me.lblThis.Caption = "附件:": Me.lblThis.FontBold = False: Screen.MousePointer = vbDefault
    Call SaveErrLog
End Function

Private Function AnnexDel() As Boolean
    '删除附件文件
    gstrSQL = "Zl_电子病历附件_Del(" & mlngRecordId & "," & Mid(Me.lvwThis.SelectedItem.Key, 2) & ")"
    Err = 0: On Error GoTo errHand
    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    AnnexDel = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'-----------------------------------------------------
'窗体控件事件
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngCount As Long, lngAdd As Long
    
    If mblnMoved And (Control.ID = conPopuAdd Or Control.ID = conPopuDel Or Control.ID = conPopuPaste) Then
        MsgBox "该病人的数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                        "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Select Case Control.ID
    Case conPopuAdd
        With Me.dlgThis
            .Filename = ""
            .DialogTitle = "选择添加附件"
            .ShowOpen
            If .Filename = "" Then Exit Sub
            DoEvents
            If AnnexAdd(.Filename) Then Call zlRefresh(mlngRecordId, mstrPrivs, mblnMoved, mblnDeleted)
        End With
    Case conPopuDel
        If MsgBox("真的删除附件“" & Me.lvwThis.SelectedItem.Text & "”吗？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        If AnnexDel Then Call zlRefresh(mlngRecordId, mstrPrivs, mblnMoved, mblnDeleted)
    Case conPopuPaste
        lngAdd = 0
        For lngCount = 0 To UBound(mClipData)
            If AnnexAdd(mClipData(lngCount)) Then lngAdd = lngAdd + 1
        Next
        If lngAdd > 0 Then Call zlRefresh(mlngRecordId, mstrPrivs)
    Case conPopuPlay
        Call AnnexPlay
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objPic As StdPicture
    Dim StrText As String, i As Integer

    Select Case Control.ID
    Case conPopuAdd
        Control.Enabled = (mblnWrite And InStr(1, mstrPrivs, "附件处理") > 0)
    Case conPopuDel
        Control.Enabled = (mblnWrite And InStr(1, mstrPrivs, "附件处理") > 0 And Not (Me.lvwThis.SelectedItem Is Nothing) And mblnDeleted)
    Case conPopuPaste
        Control.Enabled = (mblnWrite And InStr(1, mstrPrivs, "附件处理") > 0)
        If Control.Enabled Then
            Control.Enabled = False
            mClipData = GetClipboard(objPic)
            If TypeName(mClipData) = "Empty" Then Exit Sub
            If TypeName(mClipData) = "String" Then Exit Sub
            If TypeName(mClipData) = "Variant()" Then Control.Enabled = True
        End If
    
    Case conPopuPlay
        Control.Enabled = Not (Me.lvwThis.SelectedItem Is Nothing)
    End Select
End Sub

Private Sub Form_Load()
    Me.BackColor = Me.lvwThis.BackColor
    Me.imgThis.MaskColor = vbWhite
    Me.cbsThis.ActiveMenuBar.Visible = False
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("V"), conPopuPaste
    End With
    mlngRecordId = -1: mstrPrivs = ""
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.lblThis.Top = (Me.ScaleHeight - Me.lblThis.Height) / 2
    With Me.lvwThis
        .Width = Me.ScaleWidth - .Left
        .Top = Screen.TwipsPerPixelY * 4: .Height = Me.ScaleHeight - Screen.TwipsPerPixelY * 8
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set mClipData = Nothing
End Sub

Private Sub lvwThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    
    If Button <> vbRightButton Then Exit Sub
    Set cbrPopupBar = Me.cbsThis.Add("弹出菜单", xtpBarPopup)
    With cbrPopupBar.Controls
        Set cbrPopupItem = .Add(xtpControlButton, conPopuAdd, "添加(&A)...")
        Set cbrPopupItem = .Add(xtpControlButton, conPopuDel, "删除(&D)")
        Set cbrPopupItem = .Add(xtpControlButton, conPopuPaste, "粘贴(&V)...")
        Set cbrPopupItem = .Add(xtpControlButton, conPopuPlay, "查阅(&L)..."): cbrPopupItem.BeginGroup = True
    End With
    cbrPopupBar.ShowPopup
End Sub
