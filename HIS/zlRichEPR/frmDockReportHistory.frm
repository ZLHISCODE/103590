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
   StartUpPosition =   3  '����ȱʡ
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
         Caption         =   "�鿴������ʷ����"
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
            Name            =   "����"
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
         Caption         =   "�Ա�"
         Height          =   350
         Left            =   4020
         TabIndex        =   8
         ToolTipText     =   "��ǰѡ��Ӱ�����Ѵ򿪵�Ӱ��Աȹ�Ƭ"
         Top             =   30
         Width           =   510
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "����"
         Height          =   350
         Left            =   3510
         TabIndex        =   2
         ToolTipText     =   "��ǰѡ������Ӱ�������Ƭ"
         Top             =   30
         Width           =   510
      End
      Begin VB.CheckBox chkCopy 
         Height          =   350
         Left            =   2970
         Picture         =   "frmDockReportHistory.frx":009D
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "����ѡ�����ݲ�ճ�����༭���ڵ�ǰλ��"
         Top             =   30
         Width           =   300
      End
      Begin VB.Label lblContent 
         Caption         =   "��������"
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
    ҽ��
    ʱ��
    ת��
    ҽ��id
End Enum

Public Event CopyClick(ByVal strContent As String)    '��������
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
    
    gstrSQL = "Select a.Id, c.ҽ������, Nvl(a.���ʱ��, a.����ʱ��) ʱ��, 0 ת��,C.ID ҽ��ID" & vbNewLine & _
                    "From ���Ӳ�����¼ A, ����ҽ������ B, ����ҽ����¼ C" & vbNewLine & _
                    "Where a.����id = [1]  " & IIf(chkViewHistory.Value = 0, " And a.����ID=[2]", "") & " And a.ID+0<>[3] And a.�������� = 7 And a.�༭��ʽ = 0 And b.����id = a.Id And c.Id = b.ҽ��id" & vbNewLine & _
                    "Union" & vbNewLine & _
                    "Select a.Id, c.ҽ������, Nvl(a.���ʱ��, a.����ʱ��) ʱ��, 1 ת��,C.ID ҽ��ID" & vbNewLine & _
                    "From H���Ӳ�����¼ A, H����ҽ������ B, H����ҽ����¼ C" & vbNewLine & _
                    "Where a.����id = [1] " & IIf(chkViewHistory.Value = 0, " And a.����ID=[2]", "") & " And a.ID+0<>[3] And a.�������� = 7 And a.�༭��ʽ = 0 And b.����id = a.Id And c.Id = b.ҽ��id"
                    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ��ʷ����", lngPatiID, lngDeptId, lngCurFileID)
    With vsHistory
        .Rows = rsTemp.RecordCount + 1
        Do Until rsTemp.EOF
            .ROWHEIGHT(rsTemp.AbsolutePosition) = 400
            .TextMatrix(rsTemp.AbsolutePosition, 0) = rsTemp!ID
            .TextMatrix(rsTemp.AbsolutePosition, 1) = rsTemp!ҽ������
            .TextMatrix(rsTemp.AbsolutePosition, 2) = rsTemp!ʱ��
            .TextMatrix(rsTemp.AbsolutePosition, 3) = rsTemp!ת��
            .TextMatrix(rsTemp.AbsolutePosition, 4) = rsTemp!ҽ��id
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
                blnReFind = False 'ֻҪ�鵽�ؼ��־���Ҫ��������,��Ϊÿ��ֻ����һ���ؼ���
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
                
                If InStr(strContent, "��") > 0 Then
                    strContent = Mid(strContent, 1, InStr(strContent, "��") - 1) & Mid(strContent, InStr(strContent, "��") + 1)
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
    lngRecordId = Val(vsHistory.TextMatrix(vsHistory.Row, mCol.ҽ��id))
    lngMoved = Val(vsHistory.TextMatrix(vsHistory.Row, mCol.ת��))
    If lngRecordId = 0 Then Exit Sub
    ViewImage lngRecordId, Me, lngMoved = 1, True
End Sub
Private Sub cmdView_Click()
Dim lngRecordId As Long, lngMoved As Long
    lngRecordId = Val(vsHistory.TextMatrix(vsHistory.Row, mCol.ҽ��id))
    lngMoved = Val(vsHistory.TextMatrix(vsHistory.Row, mCol.ת��))
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
    
'    chkViewHistory.Visible = IIf(InStr(gstrPrivs, "PACS�������Ʊ���") > 0, True, False)
    
    Set Pane1 = dkpMain.CreatePane(101, 400, 100, DockTopOf, Nothing)
    Pane1.Title = "��ʷ�б�"
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane2 = dkpMain.CreatePane(102, 400, 30, DockBottomOf, Nothing)
    pane2.Title = "��ͷ"
    pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane2.MaxTrackSize.Height = 30: pane2.MinTrackSize.Height = 30
    
    Set pane3 = dkpMain.CreatePane(103, 400, 300, DockBottomOf, Nothing)
    pane3.Title = "����"
    pane3.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    With vsHistory
        .Clear: .Rows = 2: .Cols = 5
        .ColWidth(0) = 0: .ColWidth(1) = 2400: .ColWidth(2) = 1800: .ColWidth(3) = 0: .ColWidth(4) = 0
        .TextMatrix(0, 0) = "ID": .TextMatrix(0, 1) = "ҽ������": .TextMatrix(0, 2) = "ʱ��": .TextMatrix(0, 3) = "ת��": .TextMatrix(0, 4) = "ҽ��ID"
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
    lngMoved = Val(vsHistory.TextMatrix(vsHistory.Row, mCol.ת��))
    If lngRecordId = 0 Then Exit Sub
    Call zlRefDocment(lngRecordId, lngMoved = 1)
End Sub
Private Sub zlRefDocment(ByVal lngEPRid As Long, ByVal blnMoved As Boolean)
'���ܣ�ˢ�²�����ʾ���ݣ�
'������lngEPRId-���Ӳ�����¼ID
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
Private Sub ViewImage(ByVal lngҽ��id As Long, frmParent As Object, _
                        Optional ByVal blnMoved As Boolean = False, Optional ByVal blnCompare As Boolean = False)
'���ܣ����ù�Ƭվ
'���ܣ��Ƿ�Աȹ�Ƭ
    Dim strFtpHost As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strSDPath As String
    Dim strSDUser As String
    Dim strSDPwd As String
    
    On Error GoTo DBError
    
    '���ж��Ƿ����ͼ��û��ͼ������ʾ���˳�
    strSQL = "Select A.���UID,Count(B.����UID) as �������� From Ӱ�����¼ A,Ӱ�������� B Where A.���UID=B.���UID And A.ҽ��ID=[1] Group by A.���UID"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "��Ƭ����", lngҽ��id)
    If rsTmp.EOF Then
        MsgBox "û�п����ڹ�Ƭ�ı���ͼ��", vbInformation, gstrSysName
        Exit Sub
    End If

    strFtpHost = ""
    
    '������Ҫ�򿪵�����ͼ����Ϣ
    strSQL = "Select /*+RULE*/ D.IP��ַ As Host1,d.�豸�� as �豸��1," & _
        "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'\')" & _
        "||C.���UID||'\' As Path,E.IP��ַ As Host2,e.�豸�� as �豸��2, " & _
        "D.����Ŀ¼ AS ����Ŀ¼1, E.����Ŀ¼ AS ����Ŀ¼2,D.����Ŀ¼�û��� as ����Ŀ¼�û���1, " & _
        "E.����Ŀ¼�û��� AS ����Ŀ¼�û���2,D.����Ŀ¼���� AS ����Ŀ¼����1,E.����Ŀ¼���� AS ����Ŀ¼����2 " & _
        "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) And C.ҽ��ID=[1] "
        
    '�����ת����־�����ȡת������ʷ��
    If blnMoved Then
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "��ȡ����Ŀ¼��Ϣ", lngҽ��id)
    
    If rsTmp.RecordCount > 0 Then
        '�������صĻ���Ŀ¼����Ҫ�ڵ��ù�Ƭվ֮ǰ�ȴ������Ŀ¼����Ƭվ��ֻ�����أ����������ػ���Ŀ¼
        MkLocalDir App.Path & "\TmpImage\" & rsTmp("Path")
        ClearCacheFolder App.Path & "\TmpImage\"
        
        '��ȡFTP�����������û��������룬IP��ַ��
        If rsTmp("�豸��1") <> "" Then
            strFtpHost = rsTmp("Host1")
            strSDPath = NVL(rsTmp("����Ŀ¼1"))
            strSDUser = NVL(rsTmp("����Ŀ¼�û���1"))
            strSDPwd = NVL(rsTmp("����Ŀ¼����1"))
        ElseIf NVL(rsTmp("�豸��2")) <> "" Then
            strFtpHost = rsTmp("Host2")
            strSDPath = NVL(rsTmp("����Ŀ¼2"))
            strSDUser = NVL(rsTmp("����Ŀ¼�û���2"))
            strSDPwd = NVL(rsTmp("����Ŀ¼����2"))
        End If
        
        '�жϹ���Ŀ¼�Ƿ��Ѿ����ӣ����û�����ӣ����������
        On Error Resume Next
        If strSDPath <> "" Then
            Call funcConnectShardDir("\\" & strFtpHost & "\" & strSDPath, strSDUser, strSDPwd)
        End If
        
        If gobjPacsCore Is Nothing Then
            Set gobjPacsCore = CreateObject("zl9PacsCore.clsViewer")
            mblnCreate = True
        End If
        gobjPacsCore.CallOpenViewer "", lngҽ��id, frmParent, gcnOracle, blnMoved, blnCompare
    End If

    Exit Sub
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function funcConnectShardDir(strShareRemoteDir As String, strUserName As String, strPassWord As String) As Long
    '����������Դ
    Dim NetR As NETRESOURCE
    Dim lngResult As Long
    
    NetR.dwType = RESOURCETYPE_ANY
    NetR.lpLocalName = vbNullString
    NetR.lpRemoteName = strShareRemoteDir
    NetR.lpProvider = vbNullString
    lngResult = WNetAddConnection2(NetR, strPassWord, strUserName, 0)
    
    If lngResult <> 0 Then
        MsgBox "��������ʧ�ܣ��������������Ƿ���ȷ��"
    End If
    funcConnectShardDir = lngResult
End Function

Private Sub MkLocalDir(ByVal strDir As String)
'���ܣ���������Ŀ¼
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '��ȡȫ����Ҫ������Ŀ¼��Ϣ
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '����ȫ��Ŀ¼
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Private Sub ClearCacheFolder(ByVal strCacheFolder As String)
'���ܣ���ָ��Ŀ¼�Ĵ�С�ﵽһ���ٷֱ�ʱ����ո�Ŀ¼
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

