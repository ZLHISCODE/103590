VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockAnnex 
   BorderStyle     =   0  'None
   Caption         =   "����"
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8220
   Icon            =   "frmDockAnnex.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
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
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�ļ�"
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
      Caption         =   "����:"
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
'�˵�����
'-----------------------------------------------------
Const conPopuPlay = 101
Const conPopuAdd = 201
Const conPopuDel = 202
Const conPopuPaste = 203

'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mlngRecordId As Long        '������¼ID
Private mstrPrivs As String         '��ǰ�û��ĸ���Ȩ��
Private mblnWrite As Boolean        '�Ƿ����ɾ�������Ѿ��鵵�Ĳ���������ɾ�ĸ���
Private mblnMoved As Boolean        '�Ƿ���ת��
Private mClipData As Variant        '����������
Private mblnDeleted As Boolean      'ɾ����ťȨ��

'-----------------------------------------------------
'���幫������
'-----------------------------------------------------
Public Sub zlRefresh(ByVal lngRecordId As Long, Optional ByVal strPrivs As String, Optional ByVal blnMoved As Boolean, Optional ByVal blnDeleted As Boolean)
    '���ܣ�ˢ�²��������б�
    '������lngRecordId�����Ӳ�����¼ID��strPrivs����ǰ�û��ĸ���Ȩ��
    Dim rsTemp As New ADODB.Recordset
    Dim objItem As ListItem
    Dim objIcon As StdPicture
    
    mblnMoved = blnMoved
    mblnDeleted = blnDeleted
    mlngRecordId = lngRecordId: mstrPrivs = strPrivs
    
    Set Me.lvwThis.Icons = Nothing: Set Me.lvwThis.SmallIcons = Nothing
    Me.lvwThis.ListItems.Clear: Me.imgThis.ListImages.Clear
    
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select Decode(�鵵��, Null, 1, 0) As ��д From ���Ӳ�����¼ Where ID = [1]"
    If mblnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngRecordId)
    If rsTemp.RecordCount <= 0 Then
        mblnWrite = False
    ElseIf rsTemp.Fields(0).Value = 1 Then
        mblnWrite = True
    Else
        mblnWrite = False
    End If
    
    gstrSQL = "Select ���, �ļ���, ��С, ������, ���� From ���Ӳ������� Where ����id = [1]"
    If mblnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngRecordId)
    With rsTemp
        Do While Not .EOF
            Set objIcon = GetFileIcon(!�ļ���, ICON_SMALL, True)
            Me.imgThis.ListImages.Add , , objIcon
            Set Me.lvwThis.Icons = Me.imgThis: Set Me.lvwThis.SmallIcons = Me.imgThis
            Set objItem = Me.lvwThis.ListItems.Add(, "_" & !���, !�ļ��� & "(" & !��С & "KB)")
            objItem.Tag = !�ļ���
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
    '���ܣ�����ָ���ļ��Ĵ�ͼ���Сͼ��
    '˵������Ҫһ��PictureBox�ؼ����ޱ߿�AutoRedraw = True
    '������ strFile��������׺���ļ��������ļ���ʵ�ļ�ʱ��Ӧ�ð���������·����
    '       intSize����ȡͼ��Ĵ�С
    '       blnUntrue������ʵ�ļ�����ʱ��Ҫ�����ļ�����������Ϣ
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
'���ܣ���ȡ��ǰ�������е�����
'���أ�
'    1.�ı���TypeName=String
'    2.����ļ���Ŀ¼��TypeName=Variant()
'    3.ͼƬ��Not objPic is Nothing��TypeName=String(ͼ��·��)��Long(ͼ����)
'    4.�����ݣ�TypeName=Empty
    Dim arrFormat(8) As Long, lngFormat As Long
    Dim strFile As String, i As Long
    Dim hDrop As Long, lngFiles As Long
    Dim arrFile As Variant
    
    Set objPic = Nothing
    GetClipboard = Empty
    
    If CountClipboardFormats = 0 Then Exit Function '������
    
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
    If lngFormat = 0 Then Exit Function '������
    If lngFormat = -1 Then Exit Function 'δָ���ĸ�ʽ

    Select Case lngFormat
    Case CF_TEXT, CF_OEMTEXT, CF_DSPTEXT '�ı�
        GetClipboard = Clipboard.GetText()
    Case CF_HDROP '�ļ�/Ŀ¼
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
    Case CF_BITMAP, CF_DSPBITMAP, CF_METAFILEPICT, CF_DSPMETAFILEPICT, CF_ENHMETAFILE 'ͼƬ��ʽ
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
        
        '�����ж�Ӧ��ͼƬ·��
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
    '���ܣ��򿪲��Ÿ���
    Dim strFile As String
    Dim varRetu As Variant, strInfo As String
    
    Screen.MousePointer = vbHourglass
    strFile = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp")) & "\" & Me.lvwThis.SelectedItem.Tag
    If gobjFSO.FileExists(strFile) Then gobjFSO.DeleteFile strFile, True
    If zlBlobRead(8, mlngRecordId & "," & Mid(Me.lvwThis.SelectedItem.Key, 2), strFile, mblnMoved) = "" Then
        MsgBox "�ļ���ȡʧ�ܣ���ȷ�ϸ�������Ч�ԣ�", vbInformation, gstrSysName:
        Screen.MousePointer = vbDefault: Exit Sub
    End If
    varRetu = ShellExecute(Me.hWnd, "open", strFile, "", "", SW_SHOWNORMAL)
    If varRetu <= 32 Then
        Select Case varRetu
        Case 2: strInfo = "����Ĺ���"
        Case 29: strInfo = "����ʧ��"
        Case 30: strInfo = "����Ӧ�ó�ʽæµ��..."
        Case 31: strInfo = "û�й����κ�Ӧ�ó�ʽ"
        Case Else: strInfo = "�޷�ʶ��Ĵ���"
        End Select
        MsgBox "�����򿪷�����" & strInfo, vbExclamation, gstrSysName
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Function AnnexAdd(ByVal strFile As String) As Boolean
    '���ܣ�ѡ�����Ӹ����ļ�
    Dim objFile As File
    Dim rsTemp As New ADODB.Recordset, lngMaxNo As Long
    Dim arySql() As String, lngSql As Long, blnTran As Boolean
    
    Err = 0: On Error GoTo errHand
    Set objFile = gobjFSO.GetFile(strFile)
    Me.lblThis.Caption = "�������...": Me.lblThis.FontBold = True: Screen.MousePointer = vbHourglass
    
    gstrSQL = "Select Nvl(Max(���), 0) + 1 As ��� From ���Ӳ������� Where ����id = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngRecordId)
    lngMaxNo = rsTemp.Fields(0).Value
    
    ReDim arySql(0 To 0)
    arySql(0) = "Zl_���Ӳ�������_Add(" & mlngRecordId & "," & lngMaxNo & ",'" & objFile.Name & "'," & Format(objFile.Size / 1024, "0.00") & ")"
    If zlBlobSql(8, mlngRecordId & "," & lngMaxNo, strFile, arySql()) = False Then
        MsgBox strFile & "��������ʧ�ܣ�", vbExclamation, gstrSysName: Exit Function
    End If

    'ִ�б���
    gcnOracle.BeginTrans: blnTran = True
    For lngSql = LBound(arySql) To UBound(arySql)
        Call zldatabase.ExecuteProcedure(arySql(lngSql), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTran = False
    Me.lblThis.Caption = "����:": Me.lblThis.FontBold = False: Screen.MousePointer = vbDefault
    AnnexAdd = True
    Exit Function
errHand:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Me.lblThis.Caption = "����:": Me.lblThis.FontBold = False: Screen.MousePointer = vbDefault
    Call SaveErrLog
End Function

Private Function AnnexDel() As Boolean
    'ɾ�������ļ�
    gstrSQL = "Zl_���Ӳ�������_Del(" & mlngRecordId & "," & Mid(Me.lvwThis.SelectedItem.Key, 2) & ")"
    Err = 0: On Error GoTo errHand
    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    AnnexDel = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'-----------------------------------------------------
'����ؼ��¼�
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngCount As Long, lngAdd As Long
    
    If mblnMoved And (Control.ID = conPopuAdd Or Control.ID = conPopuDel Or Control.ID = conPopuPaste) Then
        MsgBox "�ò��˵������Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                        "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Select Case Control.ID
    Case conPopuAdd
        With Me.dlgThis
            .Filename = ""
            .DialogTitle = "ѡ����Ӹ���"
            .ShowOpen
            If .Filename = "" Then Exit Sub
            DoEvents
            If AnnexAdd(.Filename) Then Call zlRefresh(mlngRecordId, mstrPrivs, mblnMoved, mblnDeleted)
        End With
    Case conPopuDel
        If MsgBox("���ɾ��������" & Me.lvwThis.SelectedItem.Text & "����", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
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
        Control.Enabled = (mblnWrite And InStr(1, mstrPrivs, "��������") > 0)
    Case conPopuDel
        Control.Enabled = (mblnWrite And InStr(1, mstrPrivs, "��������") > 0 And Not (Me.lvwThis.SelectedItem Is Nothing) And mblnDeleted)
    Case conPopuPaste
        Control.Enabled = (mblnWrite And InStr(1, mstrPrivs, "��������") > 0)
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
    Set cbrPopupBar = Me.cbsThis.Add("�����˵�", xtpBarPopup)
    With cbrPopupBar.Controls
        Set cbrPopupItem = .Add(xtpControlButton, conPopuAdd, "���(&A)...")
        Set cbrPopupItem = .Add(xtpControlButton, conPopuDel, "ɾ��(&D)")
        Set cbrPopupItem = .Add(xtpControlButton, conPopuPaste, "ճ��(&V)...")
        Set cbrPopupItem = .Add(xtpControlButton, conPopuPlay, "����(&L)..."): cbrPopupItem.BeginGroup = True
    End With
    cbrPopupBar.ShowPopup
End Sub
