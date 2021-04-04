VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmDockEPRContent 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "�����ļ����"
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3615
      ScaleWidth      =   6375
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.PictureBox picRich 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   3150
         Left            =   585
         ScaleHeight     =   3150
         ScaleWidth      =   4830
         TabIndex        =   1
         Top             =   0
         Width           =   4830
         Begin zlRichEditor.Editor edtThis 
            Height          =   2580
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   4551
            WithViewButtonas=   0   'False
            ShowRuler       =   0   'False
         End
      End
      Begin XtremeDockingPane.DockingPane dkpMan 
         Left            =   0
         Top             =   2865
         _Version        =   589884
         _ExtentX        =   450
         _ExtentY        =   423
         _StockProps     =   0
         VisualTheme     =   5
      End
   End
End
Attribute VB_Name = "frmDockEPRContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'�����¼�
'-----------------------------------------------------
Public Event DblClick()                                                 '����˫�������¼�

Private Enum FileType
    conPane_RichEpr = 1
    conPane_TablEpr = 2
    conPane_Feedback = 3
    conPane_Infection = 4
End Enum
'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mlngPatientID As Long       '����ID
Private mlngRecordId As Long        '������¼ID
Private mfrmReport  As frmDiseaseRegist    '��Ⱦ�����Խ��������
Private mObjTabEprView As cTableEPR      '�����
Private mobjInfection As Object          '�л����񹲺͹���Ⱦ�����濨

Public mIsShowAnnex As Boolean

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_RichEpr
        Item.Handle = picRich.hWnd
    Case conPane_Feedback
        Item.Handle = mfrmReport.hWnd
    Case conPane_TablEpr
        Item.Handle = mObjTabEprView.zlGetForm.hWnd
    Case conPane_Infection
        Item.Handle = mobjInfection.zlGetForm.hWnd
    End Select
End Sub

Private Sub edtThis_DblClick(ViewMode As zlRichEditor.ViewModeEnum)
    RaiseEvent DblClick
End Sub

Private Sub Form_Load()
    Dim Pane1 As Pane, pane2 As Pane, pane3 As Pane, pane4 As Pane
    On Error GoTo errHand
    Set mObjTabEprView = New cTableEPR
    mObjTabEprView.InitTableEPR gcnOracle, glngSys, gstrDBOwer
    Set mfrmReport = New frmDiseaseRegist
    Call mfrmReport.SetFrmInset(True)
    
    Set Pane1 = dkpMan.CreatePane(conPane_RichEpr, 1200, 200, DockTopOf, Nothing)
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane2 = dkpMan.CreatePane(conPane_TablEpr, 1200, 200, DockTopOf, Nothing)
    pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane2.Close
    
    Set pane3 = dkpMan.CreatePane(conPane_Feedback, 1200, 200, DockTopOf, Nothing)
    pane3.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane3.Close
    
    Set pane3 = dkpMan.CreatePane(conPane_Infection, 1200, 200, DockTopOf, Nothing)
    pane3.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane3.Close
    
    With dkpMan
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = True
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = False
    End With
        
    Set mobjInfection = DynamicCreate("zlDisReportCard.clsDisReportCard", "��Ⱦ�����濨", True)
    If Not mobjInfection Is Nothing Then
        mobjInfection.Init gcnOracle, glngSys
    End If
    
    mlngRecordId = -1
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

'-----------------------------------------------------
'���幫������
'-----------------------------------------------------

Public Sub zlRefresh(ByVal lngRecordId As Long, strAnnexRight As String, Optional ByVal blnPrivacyProtect As Boolean, _
                Optional ByVal blnMoved As Boolean, Optional ByRef blnViewFile As Boolean, Optional ByVal byteEdit As Byte, _
                Optional ByVal blnAllowDelete As Boolean)
'���ܣ�ˢ�²�����ʾ���ݣ�
'������lngRecordId�����Ӳ�����¼ID��blnPrivacyProtect���Ƿ�������˽����;strAnnexRight-��������Ȩ��,byteEdit=0 RichEdit =1 ���ʽ����;blnViewFile �Ƿ����Ԥ��
    Dim blnPrivacy As Boolean, Elements As New cEPRElements
    Dim rs As New ADODB.Recordset, lngKey As Long
    Dim strSQL As String
    
    On Error GoTo errHand
    If blnPrivacyProtect = True Then
        blnPrivacy = InStr(gstrPrivs, ";������˽����;") = 0     '������˽��Ŀ
    End If
    
    mlngRecordId = lngRecordId
    dkpMan.FindPane(conPane_RichEpr).Close
    dkpMan.FindPane(conPane_TablEpr).Close
    dkpMan.FindPane(conPane_Feedback).Close
    dkpMan.FindPane(conPane_Infection).Close
    
    If byteEdit = 1 Then
        dkpMan.ShowPane conPane_TablEpr
        Call mObjTabEprView.InitOpenEPR(Me, cprEM_�޸�, cprET_���������, mlngRecordId, False, 0)
        Call mObjTabEprView.zlRefreshDockfrm
        blnViewFile = True
    ElseIf byteEdit = 2 Then '��Ⱦ�����濨ר�ñ༭��
        dkpMan.ShowPane conPane_Infection
        strSQL = "Select ID,����ID,��ҳID From ���Ӳ�����¼ Where ID=[1]"
        If blnMoved Then strSQL = Replace(strSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
        Set rs = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngRecordId)
        mobjInfection.zlRefresh rs!����ID, rs!��ҳID, lngRecordId, blnMoved
    ElseIf byteEdit = 3 Then '��Ⱦ�����Խ��������
        dkpMan.ShowPane conPane_Feedback
        Call mfrmReport.zlRefresh(mlngRecordId)
        Call mfrmReport.SetReportTop(100)
    Else
        dkpMan.ShowPane conPane_RichEpr
        Dim strTemp As String, strZipFile As String
        Me.edtThis.Freeze
        Me.edtThis.ReadOnly = False
        Me.edtThis.NewDoc
        strZipFile = zlBlobRead(5, lngRecordId, , blnMoved)
        If gobjFSO.FileExists(strZipFile) Then
            strTemp = zlFileUnzip(strZipFile)
            If gobjFSO.FileExists(strTemp) Then
                '���ļ�
                Me.edtThis.OpenDoc strTemp
                '�����滻��Ŀ
                If blnPrivacy Then
                    '��ȡ���е�Ҫ��
                    strSQL = "Select A.ID,A.������ From ���Ӳ������� A, ��˽������Ŀ B,����������Ŀ C " & _
                        "Where A.�������� = 4 And A.�滻�� = 1 And A.�ļ�id = [1] And A.������� > 0 and B.��Ŀid = C.ID And A.Ҫ������ =C.������ And C.�滻�� = 1 "
                    If blnMoved Then strSQL = Replace(strSQL, "���Ӳ�������", "H���Ӳ�������")
                    Set rs = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngRecordId)
                    If Not rs.EOF Then
                        Do While Not rs.EOF
                            lngKey = Elements.Add(Nvl(rs("������"), 0))
                            Elements("K" & lngKey).GetElementFromDB cprET_�������༭, rs("ID"), True, IIf(blnMoved, "H���Ӳ�������", "���Ӳ�������")
                            '�滻Ҫ������
                            Elements("K" & lngKey).�����ı� = String(Len(Elements("K" & lngKey).�����ı�), "*")
                            Elements("K" & lngKey).Refresh Me.edtThis
                            rs.MoveNext
                        Loop
                    End If
                    rs.Close
                End If
                gobjFSO.DeleteFile strTemp, True
            End If
            gobjFSO.DeleteFile strZipFile, True
            Me.edtThis.SelStart = 0
            blnViewFile = True
        Else
            blnViewFile = False
        End If
        If lngRecordId > 0 Then
            '����ҳ���ʽ
            Dim mEPRFileInfo As New cEPRFileDefineInfo
            strSQL = "Select c.ID, a.��ʽ,c.����ID From   ����ҳ���ʽ a, �����ļ��б� b, ���Ӳ�����¼ c " & _
                    " Where  c.�ļ�id = b.id And a.���� = b.���� And a.��� = b.ҳ�� And c.ID = [1]"
            If blnMoved Then strSQL = Replace(strSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
            Set rs = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngRecordId)
            If Not rs.EOF Then
                mlngPatientID = rs!����ID
                mEPRFileInfo.��ʽ = gobjComlib.zlCommFun.Nvl(rs("��ʽ").Value)
                mEPRFileInfo.SetFormat Me.edtThis, mEPRFileInfo.��ʽ
                Me.edtThis.ResetWYSIWYG
            End If
            Set mEPRFileInfo = Nothing
        End If
        Call RefreshObject(lngRecordId, blnMoved)
        Me.edtThis.SelStart = 0
        Me.edtThis.UnFreeze
        edtThis.RefreshTargetDC
        Me.edtThis.ViewMode = cprNormal
        Me.edtThis.ReadOnly = True
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picMain.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload mfrmReport
    Set mfrmReport = Nothing
    Unload mObjTabEprView.zlGetForm
    Set mObjTabEprView = Nothing
    Unload mobjInfection.zlGetForm
    Set mobjInfection.zlGetForm = Nothing
    Set mobjInfection = Nothing
End Sub

Private Sub picRich_Resize()
On Error Resume Next
    edtThis.Top = 0: edtThis.Left = 0
    edtThis.Width = picRich.ScaleWidth: edtThis.Height = picRich.Height
End Sub

Private Sub RefreshObject(ByVal lngRecordId As Long, ByVal blnMoved As Boolean)
'ˢ�½����ϵ�ͼƬ,Ŀǰֻˢ��ͼƬ������Ҫʱ�ٵ���ˢ�±��
    Dim Pictures As New cEPRPictures, rsTemp As New ADODB.Recordset, lngKey As Long, Tables As New cEPRTables
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean, blnForce As Boolean
    Dim strSQL As String
    '��ȡ���е�ͼƬ
    strSQL = "Select ID, �ļ�id,��ʼ��, ��ֹ��,��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���,Ԥ�����ID " & _
        "From ���Ӳ������� " & _
        "Where �ļ�id = [1] And �������� in(3,5) And ������� Is Not Null" '����ʾ����е�ͼƬ
    If blnMoved Then strSQL = Replace(strSQL, "���Ӳ�������", "H���Ӳ�������")
    Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngRecordId)
    Do Until rsTemp.EOF
        If rsTemp!�������� = 5 Then
            lngKey = Pictures.Add(Nvl(rsTemp!������, 0))
            Call Pictures("K" & lngKey).FillPictureMember(rsTemp, IIf(blnMoved, "H���Ӳ�������", "���Ӳ�������"))
            Call Pictures("K" & lngKey).DeleteFromEditor(edtThis)
            Call Pictures("K" & lngKey).InsertIntoEditor(edtThis, -1, True)
        ElseIf rsTemp!�������� = 3 Then
            lngKey = Tables.Add(Nvl(rsTemp!������, 0))
            Call Tables("K" & lngKey).FillTableMember(rsTemp, IIf(blnMoved, "H���Ӳ�������", "���Ӳ�������"))
            
            If Tables("K" & lngKey).Cells.Count = 1 Then
                'һ����Ԫ�񣬿�����PACS�༭����д������
                If FindKey(edtThis, "T", lngKey, lKSS, lKSE, lKES, lKEE, True) Then
                    '��ɾ��
                    Call Tables("K" & lngKey).DeleteFromEditor(edtThis)
                    With edtThis
                        blnForce = .ForceEdit
                        .InProcessing = True
                        .Tag = "TableSingleCell:InsertIntoEditor"
                        .ForceEdit = True
                        .Range(lKSS, lKSS).Font.Protected = False
                        .Range(lKSS, lKSS).Font.Hidden = False
                        .Range(lKSS, lKSS) = Tables("K" & lngKey).Cells(1).�����ı�
                        .ForceEdit = blnForce
                        .UnFreeze
                        .InProcessing = False
                        .Tag = ""
                    End With
                End If
            Else
                '�����Ԫ��
                '��ɾ��
                Call Tables("K" & lngKey).DeleteFromEditor(edtThis)
                Call Tables("K" & lngKey).InsertIntoEditor(edtThis, -1)
            End If
        End If
        rsTemp.MoveNext
    Loop
End Sub





