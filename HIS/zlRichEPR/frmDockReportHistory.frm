VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmDockReportHistory 
   BorderStyle     =   0  'None
   ClientHeight    =   6435
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picHistory 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2340
      Left            =   420
      ScaleHeight     =   2340
      ScaleWidth      =   3180
      TabIndex        =   4
      Top             =   225
      Width           =   3180
      Begin VB.CheckBox chkViewHistory 
         Caption         =   "查看他科历史报告"
         Height          =   300
         Left            =   105
         TabIndex        =   9
         Top             =   135
         Width           =   2190
      End
      Begin VSFlex8Ctl.VSFlexGrid vsHistory 
         Height          =   1275
         Left            =   180
         TabIndex        =   5
         Top             =   825
         Width           =   2070
         _cx             =   3651
         _cy             =   2249
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
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
         FormatString    =   ""
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
   End
   Begin VB.PictureBox picRichEdit 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   2100
      Left            =   525
      ScaleHeight     =   2100
      ScaleWidth      =   3495
      TabIndex        =   3
      Top             =   3825
      Width           =   3495
      Begin RichTextLib.RichTextBox rtbContent 
         Height          =   1635
         Left            =   225
         TabIndex        =   6
         Top             =   210
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   2884
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmDockReportHistory.frx":0000
      End
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   -45
      ScaleHeight     =   405
      ScaleWidth      =   4680
      TabIndex        =   0
      Top             =   2850
      Width           =   4680
      Begin VB.CommandButton cmdCom 
         Caption         =   "对比"
         Height          =   350
         Left            =   4020
         TabIndex        =   8
         ToolTipText     =   "当前选中影像与已打开的影像对比观片"
         Top             =   30
         Width           =   510
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "历次"
         Height          =   350
         Left            =   3510
         TabIndex        =   2
         ToolTipText     =   "当前选中历次影像独立观片"
         Top             =   30
         Width           =   510
      End
      Begin VB.CheckBox chkCopy 
         Height          =   350
         Left            =   2970
         Picture         =   "frmDockReportHistory.frx":009D
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "复制选中内容并粘贴到编辑窗口当前位置"
         Top             =   30
         Width           =   300
      End
      Begin VB.Label lblContent 
         Caption         =   "报告内容"
         Height          =   210
         Left            =   195
         TabIndex        =   1
         Top             =   105
         Width           =   795
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   30
      Top             =   45
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDockReportHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mCol
    ID = 0
    医嘱
    时间
    转出
    医嘱id
End Enum

Public Event CopyClick(ByVal strContent As String)    '复制内容
Public Event ReportCountChange(ByVal lngReportCount As Long)


Public mblnCreate As Boolean
Private mblnCopying As Boolean
Private mlngPatiId As Long
Private mlngDeptId As Long
Private mlngFileID As Long

Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngDeptId As Long, ByVal lngCurFileID As Long) As Long
Dim rsTemp As ADODB.Recordset
    On Error GoTo Errhand
    
    mblnCreate = False
    
    mlngPatiId = lngPatiID
    mlngDeptId = lngDeptId
    mlngFileID = lngCurFileID
    
    gstrSQL = "Select a.Id, c.医嘱内容, Nvl(a.完成时间, a.创建时间) 时间, 0 转出,C.ID 医嘱ID" & vbNewLine & _
                    "From 电子病历记录 A, 病人医嘱报告 B, 病人医嘱记录 C" & vbNewLine & _
                    "Where a.病人id = [1]  " & IIf(chkViewHistory.Value = 0, " And a.科室ID=[2]", "") & " And a.ID+0<>[3] And a.病历种类 = 7 And a.编辑方式 = 0 And b.病历id = a.Id And c.Id = b.医嘱id" & vbNewLine & _
                    "Union" & vbNewLine & _
                    "Select a.Id, c.医嘱内容, Nvl(a.完成时间, a.创建时间) 时间, 1 转出,C.ID 医嘱ID" & vbNewLine & _
                    "From H电子病历记录 A, H病人医嘱报告 B, H病人医嘱记录 C" & vbNewLine & _
                    "Where a.病人id = [1] " & IIf(chkViewHistory.Value = 0, " And a.科室ID=[2]", "") & " And a.ID+0<>[3] And a.病历种类 = 7 And a.编辑方式 = 0 And b.病历id = a.Id And c.Id = b.医嘱id"
                    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取历史报告", lngPatiID, lngDeptId, lngCurFileID)
    With vsHistory
        .Rows = rsTemp.RecordCount + 1
        Do Until rsTemp.EOF
            .ROWHEIGHT(rsTemp.AbsolutePosition) = 400
            .TextMatrix(rsTemp.AbsolutePosition, 0) = rsTemp!ID
            .TextMatrix(rsTemp.AbsolutePosition, 1) = rsTemp!医嘱内容
            .TextMatrix(rsTemp.AbsolutePosition, 2) = rsTemp!时间
            .TextMatrix(rsTemp.AbsolutePosition, 3) = rsTemp!转出
            .TextMatrix(rsTemp.AbsolutePosition, 4) = rsTemp!医嘱id
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount > 0 Then
            .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 2) = flexAlignLeftCenter
            .Cell(flexcpFontSize, 1, 0, .Rows - 1, .Cols - 1) = 10
        End If
    End With
    zlRefresh = rsTemp.RecordCount
    
    RaiseEvent ReportCountChange(rsTemp.RecordCount)
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume Next
    End If
End Function

Private Sub chkCopy_Click()
    
    If mblnCopying Then Exit Sub
    mblnCopying = True
    chkCopy.Value = vbUnchecked
    If rtbContent.SelText <> "" Then
        Dim strContent As String, blnReFind As Boolean, l As Long
        strContent = rtbContent.SelText
        If strContent <> "" Then
            blnReFind = True
            Do While blnReFind
                blnReFind = False '只要查到关键字就需要重新搜索,因为每次只处理一个关键字
                For l = 1 To UBound(gKeyWords)
                    If InStr(strContent, gKeyWords(l).KeyStart & "(") > 0 Then
                        strContent = Mid(strContent, 1, InStr(strContent, gKeyWords(l).KeyStart & "(") - 1) & Mid(strContent, InStr(strContent, gKeyWords(l).KeyStart & "(") + 16)
                        blnReFind = True
                    End If
                    
                    If InStr(strContent, gKeyWords(l).KeyEnd & "(") > 0 Then
                        strContent = Mid(strContent, 1, InStr(strContent, gKeyWords(l).KeyEnd & "(") - 1) & Mid(strContent, InStr(strContent, gKeyWords(l).KeyEnd & "(") + 16)
                        blnReFind = True
                    End If
                Next
                
                If InStr(strContent, "□") > 0 Then
                    strContent = Mid(strContent, 1, InStr(strContent, "□") - 1) & Mid(strContent, InStr(strContent, "□") + 1)
                    blnReFind = True
                End If
            Loop
            Debug.Print strContent
            RaiseEvent CopyClick(strContent)
        End If
    End If
    
    mblnCopying = False
End Sub

Private Sub chkViewHistory_Click()
On Error GoTo ErrHandle
    Call zlRefresh(mlngPatiId, mlngDeptId, mlngFileID)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCom_Click()
Dim lngRecordId As Long, lngMoved As Long
    lngRecordId = Val(vsHistory.TextMatrix(vsHistory.Row, mCol.医嘱id))
    lngMoved = Val(vsHistory.TextMatrix(vsHistory.Row, mCol.转出))
    If lngRecordId = 0 Then Exit Sub
    ViewImage lngRecordId, Me, lngMoved = 1, True
End Sub
Private Sub cmdView_Click()
Dim lngRecordId As Long, lngMoved As Long
    lngRecordId = Val(vsHistory.TextMatrix(vsHistory.Row, mCol.医嘱id))
    lngMoved = Val(vsHistory.TextMatrix(vsHistory.Row, mCol.转出))
    If lngRecordId = 0 Then Exit Sub
    ViewImage lngRecordId, Me, lngMoved = 1
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 101
            Item.Handle = picHistory.hwnd
        Case 102
            Item.Handle = picTitle.hwnd
        Case 103
            Item.Handle = picRichEdit.hwnd
    End Select
End Sub

Private Sub Form_Load()
Dim Pane1 As Pane, pane2 As Pane, pane3 As Pane
    With dkpMain
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = True
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = False
    End With
    
'    chkViewHistory.Visible = IIf(InStr(gstrPrivs, "PACS报告他科报告") > 0, True, False)
    
    Set Pane1 = dkpMain.CreatePane(101, 400, 100, DockTopOf, Nothing)
    Pane1.Title = "历史列表"
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane2 = dkpMain.CreatePane(102, 400, 30, DockBottomOf, Nothing)
    pane2.Title = "标头"
    pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane2.MaxTrackSize.Height = 30: pane2.MinTrackSize.Height = 30
    
    Set pane3 = dkpMain.CreatePane(103, 400, 300, DockBottomOf, Nothing)
    pane3.Title = "内容"
    pane3.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    With vsHistory
        .Clear: .Rows = 2: .Cols = 5
        .ColWidth(0) = 0: .ColWidth(1) = 2400: .ColWidth(2) = 1800: .ColWidth(3) = 0: .ColWidth(4) = 0
        .TextMatrix(0, 0) = "ID": .TextMatrix(0, 1) = "医嘱内容": .TextMatrix(0, 2) = "时间": .TextMatrix(0, 3) = "转出": .TextMatrix(0, 4) = "医嘱ID"
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If mblnCreate Then
        If Not gobjPacsCore Is Nothing Then gobjPacsCore.Closefrom
        Set gobjPacsCore = Nothing
    End If
Err.Clear
End Sub

Private Sub picHistory_Resize()
    chkViewHistory.Top = 0
    chkViewHistory.Left = 20
    chkViewHistory.Width = picHistory.ScaleWidth - 20
    
    vsHistory.Top = chkViewHistory.Height 'IIf(chkViewHistory.Visible, chkViewHistory.Height, 0)
    vsHistory.Left = 20
    vsHistory.Width = picHistory.ScaleWidth - 20
    vsHistory.Height = picHistory.ScaleHeight
End Sub

Private Sub picRichEdit_Resize()
    With rtbContent
        .Top = 0: .Left = 0
        .Width = picRichEdit.ScaleWidth: .Height = picRichEdit.ScaleHeight
    End With
End Sub

Private Sub picTitle_Resize()
    On Error Resume Next
    cmdCom.Left = picTitle.Width - cmdCom.Width - 100
    cmdView.Left = cmdCom.Left - cmdView.Width
    chkCopy.Left = cmdView.Left - chkCopy.Width - 50
End Sub
Private Sub rtbContent_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub vsHistory_RowColChange()
Dim lngRecordId As Long, lngMoved As Long
    lngRecordId = Val(vsHistory.TextMatrix(vsHistory.Row, mCol.ID))
    lngMoved = Val(vsHistory.TextMatrix(vsHistory.Row, mCol.转出))
    If lngRecordId = 0 Then Exit Sub
    Call zlRefDocment(lngRecordId, lngMoved = 1)
End Sub
Private Sub zlRefDocment(ByVal lngEPRid As Long, ByVal blnMoved As Boolean)
'功能：刷新病历显示内容；
'参数：lngEPRId-电子病历记录ID
Dim rs As New ADODB.Recordset
Dim strTemp As String, strZipFile As String

    strZipFile = zlBlobRead(5, lngEPRid, , blnMoved)
    If gobjFSO.FileExists(strZipFile) Then
        strTemp = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strTemp) Then
            rtbContent.LoadFile strTemp
            gobjFSO.DeleteFile strTemp, True
        End If
        gobjFSO.DeleteFile strZipFile, True
    End If
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub ViewImage(ByVal lng医嘱id As Long, frmParent As Object, _
                        Optional ByVal blnMoved As Boolean = False, Optional ByVal blnCompare As Boolean = False)
'功能：调用观片站
'功能：是否对比观片
    Dim strFtpHost As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strSDPath As String
    Dim strSDUser As String
    Dim strSDPwd As String
    
    On Error GoTo DBError
    
    '先判断是否存在图像，没有图像则提示并退出
    strSQL = "Select A.检查UID,Count(B.序列UID) as 序列总数 From 影像检查记录 A,影像检查序列 B Where A.检查UID=B.检查UID And A.医嘱ID=[1] Group by A.检查UID"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "观片处理", lng医嘱id)
    If rsTmp.EOF Then
        MsgBox "没有可用于观片的报告图像。", vbInformation, gstrSysName
        Exit Sub
    End If

    strFtpHost = ""
    
    '查找需要打开的所有图象信息
    strSQL = "Select /*+RULE*/ D.IP地址 As Host1,d.设备号 as 设备号1," & _
        "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'\')" & _
        "||C.检查UID||'\' As Path,E.IP地址 As Host2,e.设备号 as 设备号2, " & _
        "D.共享目录 AS 共享目录1, E.共享目录 AS 共享目录2,D.共享目录用户名 as 共享目录用户名1, " & _
        "E.共享目录用户名 AS 共享目录用户名2,D.共享目录密码 AS 共享目录密码1,E.共享目录密码 AS 共享目录密码2 " & _
        "From 影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where C.位置一=D.设备号(+) And C.位置二=E.设备号(+) And C.医嘱ID=[1] "
        
    '如果有转储标志，则读取转储的历史表
    If blnMoved Then
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "获取共享目录信息", lng医嘱id)
    
    If rsTmp.RecordCount > 0 Then
        '创建本地的缓存目录，需要在调用观片站之前先创建这个目录，观片站中只是下载，不创建本地缓存目录
        MkLocalDir App.Path & "\TmpImage\" & rsTmp("Path")
        ClearCacheFolder App.Path & "\TmpImage\"
        
        '读取FTP参数，包括用户名，密码，IP地址等
        If rsTmp("设备号1") <> "" Then
            strFtpHost = rsTmp("Host1")
            strSDPath = NVL(rsTmp("共享目录1"))
            strSDUser = NVL(rsTmp("共享目录用户名1"))
            strSDPwd = NVL(rsTmp("共享目录密码1"))
        ElseIf NVL(rsTmp("设备号2")) <> "" Then
            strFtpHost = rsTmp("Host2")
            strSDPath = NVL(rsTmp("共享目录2"))
            strSDUser = NVL(rsTmp("共享目录用户名2"))
            strSDPwd = NVL(rsTmp("共享目录密码2"))
        End If
        
        '判断共享目录是否已经连接，如果没有连接，则进行连接
        On Error Resume Next
        If strSDPath <> "" Then
            Call funcConnectShardDir("\\" & strFtpHost & "\" & strSDPath, strSDUser, strSDPwd)
        End If
        
        If gobjPacsCore Is Nothing Then
            Set gobjPacsCore = CreateObject("zl9PacsCore.clsViewer")
            mblnCreate = True
        End If
        gobjPacsCore.CallOpenViewer "", lng医嘱id, frmParent, gcnOracle, blnMoved, blnCompare
    End If

    Exit Sub
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function funcConnectShardDir(strShareRemoteDir As String, strUserName As String, strPassWord As String) As Long
    '创建网络资源
    Dim NetR As NETRESOURCE
    Dim lngResult As Long
    
    NetR.dwType = RESOURCETYPE_ANY
    NetR.lpLocalName = vbNullString
    NetR.lpRemoteName = strShareRemoteDir
    NetR.lpProvider = vbNullString
    lngResult = WNetAddConnection2(NetR, strPassWord, strUserName, 0)
    
    If lngResult <> 0 Then
        MsgBox "网络连接失败，请检查网络设置是否正确！"
    End If
    funcConnectShardDir = lngResult
End Function

Private Sub MkLocalDir(ByVal strDir As String)
'功能：创建本地目录
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Private Sub ClearCacheFolder(ByVal strCacheFolder As String)
'功能：当指定目录的大小达到一定百分比时，清空该目录
    Dim objFile As New Scripting.FileSystemObject
    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
    Dim strDriver As String
    
    On Error Resume Next
    strDriver = objFile.GetDriveName(strCacheFolder)
    Set objCurFolder = objFile.GetFolder(strCacheFolder)
    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
        objCurFolder.Delete True
    End If
End Sub

