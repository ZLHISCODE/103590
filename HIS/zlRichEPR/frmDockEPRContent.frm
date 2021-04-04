VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmDockEPRContent 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "�����ļ����"
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picRich 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   3150
      Left            =   705
      ScaleHeight     =   3150
      ScaleWidth      =   4830
      TabIndex        =   0
      Top             =   135
      Width           =   4830
      Begin zlRichEditor.Editor edtThis 
         Height          =   2580
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4551
         WithViewButtonas=   0   'False
         ShowRuler       =   0   'False
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   120
      Top             =   2985
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   45
      Top             =   45
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
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
    conPane_Infection = 3
    conPane_Annex = 4
End Enum
'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mlngPatientID As Long       '����ID
Private mlngRecordId As Long        '������¼ID
Private mfrmAnnex As frmDockAnnex    '������������
Private mObjTabEprView As cTableEPR      '�����
Private mobjInfection As Object
Public mIsShowAnnex As Boolean

Private Function CopyEnable() As Integer
On Error GoTo errHand
Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select Zl_Fun_CopyEnable([1]) CopyEnable From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngRecordId)
    If rsTemp!CopyEnable = 1 Then
        CopyEnable = 1
    Else
        CopyEnable = 0
    End If
    
    Exit Function
errHand:
    CopyEnable = 0
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_EDIT_COPY
        If CopyEnable() = 1 Then
            If Control.Enabled And Control.Visible Then '��ݼ�ִ��ʱ��Ҫ�ж�
                gstrCopyPID = CStr(mlngPatientID)
                Me.edtThis.Copy
            End If
        Else
            MsgBox "ѡ���Ĳ�����������", vbInformation, gstrSysName
        End If
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_EDIT_COPY
        Control.Enabled = edtThis.Selection.EndPos <> edtThis.Selection.StartPos
        Control.Enabled = edtThis.Selection.getType <> cprSTPicture
        Control.Visible = InStr(gstrPrivsEpr, "���ݸ���") > 0
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_RichEpr
        Item.Handle = picRich.hWnd
    Case conPane_Annex
        If Not mIsShowAnnex Then
             Item.Handle = mfrmAnnex.hWnd
        End If
    Case conPane_TablEpr
        Item.Handle = mObjTabEprView.zlGetForm.hWnd
    Case conPane_Infection
        Item.Handle = mobjInfection.zlGetForm.hWnd
    End Select
End Sub

Private Sub edtThis_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, X As Single, Y As Single)
    'û�����ݸ���Ȩ�޲�������
    If InStr(gstrPrivsEpr, "���ݸ���") = 0 Then Exit Sub
    
    Dim Popup As CommandBar
    Dim Control As CommandBarControl
    
    Set Popup = cbsThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "����(&C)")
        Popup.ShowPopup
    End With
End Sub

Private Sub edtThis_DblClick(ViewMode As zlRichEditor.ViewModeEnum)
    RaiseEvent DblClick
End Sub
Private Sub Form_Load()
Dim Pane1 As Pane, pane2 As Pane, pane3 As Pane, Pane4 As Pane
    On Error GoTo errHand
    Set mObjTabEprView = New cTableEPR
    mObjTabEprView.InitTableEPR gcnOracle, glngSys, gstrDbOwner
    Set mfrmAnnex = New frmDockAnnex
    
    Set Pane1 = dkpMan.CreatePane(conPane_RichEpr, 1200, 200, DockTopOf, Nothing)
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane2 = dkpMan.CreatePane(conPane_TablEpr, 1200, 200, DockTopOf, Nothing)
    pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane2.Close
    
    Set pane3 = dkpMan.CreatePane(conPane_Infection, 1200, 200, DockTopOf, Nothing)
    pane3.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane3.Close
    
    If Not mIsShowAnnex Then
        Set Pane4 = dkpMan.CreatePane(conPane_Annex, 1200, 15, DockBottomOf, Nothing)
        Pane4.MinTrackSize.Height = 360 / Screen.TwipsPerPixelY: Pane4.MaxTrackSize.Height = 360 / Screen.TwipsPerPixelY
        Pane4.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    End If
    
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
    
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.KeyBindings.Add FCONTROL, Asc("C"), ID_EDIT_COPY
    mlngRecordId = -1
    Exit Sub

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub Clear()
    On Error Resume Next
    dkpMan.FindPane(conPane_TablEpr).Close
    dkpMan.FindPane(conPane_Infection).Close
    dkpMan.ShowPane conPane_RichEpr
    
    edtThis.Freeze
    edtThis.ReadOnly = False
    edtThis.ForceEdit = True
    edtThis.InProcessing = True
    edtThis.Tag = "LoadFile"
    edtThis.NewDoc
    
    edtThis.UnFreeze
    edtThis.RefreshTargetDC
    edtThis.ViewMode = cprNormal
    edtThis.ReadOnly = True
    edtThis.ForceEdit = False
    edtThis.InProcessing = False
    edtThis.Tag = ""
End Sub

Public Sub zlRefresh(ByVal lngRecordId As Long, strAnnexRight As String, Optional ByVal blnPrivacyProtect As Boolean, _
                Optional ByVal blnMoved As Boolean, Optional ByRef blnViewFile As Boolean, Optional ByVal byteEdit As Byte, _
                Optional ByVal blnAllowDelete As Boolean, Optional ByVal blnClearMode As Boolean)
'���ܣ�ˢ�²�����ʾ���ݣ�
'������lngRecordId�����Ӳ�����¼ID��blnPrivacyProtect���Ƿ�������˽����;strAnnexRight-��������Ȩ��,byteEdit=0 RichEdit =1 ���ʽ����;blnViewFile �Ƿ����Ԥ��
    Dim blnPrivacy As Boolean, Elements As New cEPRElements
    Dim rs As New ADODB.Recordset, lngKey As Long
    
    On Error GoTo errHand
    If blnPrivacyProtect = True Then
        blnPrivacy = InStr(gstrPrivsEpr, ";������˽����;") = 0     '������˽��Ŀ
    End If
    
    mlngRecordId = lngRecordId
    dkpMan.FindPane(conPane_RichEpr).Close
    dkpMan.FindPane(conPane_TablEpr).Close
    dkpMan.FindPane(conPane_Infection).Close
    If byteEdit = 1 Then
        dkpMan.ShowPane conPane_TablEpr
        Call mObjTabEprView.InitOpenEPR(Me, cprEM_�޸�, cprET_���������, mlngRecordId, False, 0, , , , , , , , blnMoved)
        Call mObjTabEprView.zlRefreshDockfrm
        blnViewFile = True
    ElseIf byteEdit = 2 Then '��Ⱦ�����濨ר�ñ༭��
        dkpMan.ShowPane conPane_Infection
        gstrSQL = "Select ID,����ID,��ҳID From ���Ӳ�����¼ Where ID=[1]"
        If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
        mobjInfection.zlRefresh rs!����ID, rs!��ҳID, lngRecordId, blnMoved
    ElseIf byteEdit = 0 Then
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
                    gstrSQL = "Select A.ID,A.������ From ���Ӳ������� A, ��˽������Ŀ B,����������Ŀ C " & _
                        "Where A.�������� = 4 And A.�滻�� = 1 And A.�ļ�id = [1] And A.������� > 0 and B.��Ŀid = C.ID And A.Ҫ������ =C.������ And C.�滻�� = 1 "
                    If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�������", "H���Ӳ�������")
                    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
                    If Not rs.EOF Then
                        Do While Not rs.EOF
                            lngKey = Elements.Add(NVL(rs("������"), 0))
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
            Call BuildRTF(edtThis, lngRecordId, blnMoved)
            blnViewFile = True
        End If
        
        ' �������ĵ��Ĵ���
        If blnClearMode Then
            edtThis.AuditMode = True
            edtThis.AcceptAuditText    '���ģʽ ���������޶��ۼ�
        End If
        
        If lngRecordId > 0 Then
            '����ҳ���ʽ
            Dim mEPRFileInfo As New cEPRFileDefineInfo
            gstrSQL = "Select c.ID, a.��ʽ,c.����ID From   ����ҳ���ʽ a, �����ļ��б� b, ���Ӳ�����¼ c " & _
                    " Where  c.�ļ�id = b.id And a.���� = b.���� And a.��� = b.ҳ�� And c.ID = [1]"
            If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
            If Not rs.EOF Then
                mlngPatientID = rs!����ID
                mEPRFileInfo.��ʽ = zlCommFun.NVL(rs("��ʽ").Value)
                mEPRFileInfo.SetFormat Me.edtThis, mEPRFileInfo.��ʽ
                Me.edtThis.ResetWYSIWYG
            End If
            Set mEPRFileInfo = Nothing
        End If
        If strZipFile <> "" Then '��RTF�ļ���ˢ�¶���
            Call RefreshObject(lngRecordId, blnMoved)
        End If
        Me.edtThis.SelStart = 0
        Me.edtThis.UnFreeze
        Me.edtThis.RefreshTargetDC
        Me.edtThis.ViewMode = cprNormal
        Me.edtThis.ReadOnly = True
        
    End If
    '���ø����б�
    Call mfrmAnnex.zlRefresh(mlngRecordId, strAnnexRight, blnMoved, blnAllowDelete)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload mfrmAnnex
    Set mfrmAnnex = Nothing
    Unload mObjTabEprView.zlGetForm
    Set mObjTabEprView = Nothing
    Unload mobjInfection.zlGetForm
    Set mobjInfection.zlGetForm = Nothing
    Set mobjInfection = Nothing
End Sub

Private Sub picRich_Resize()
    edtThis.Top = 0: edtThis.Left = 0
    edtThis.Width = picRich.ScaleWidth: edtThis.Height = picRich.Height
End Sub
Private Sub RefreshObject(ByVal lngRecordId As Long, ByVal blnMoved As Boolean)
'ˢ�½����ϵ�ͼƬ,Ŀǰֻˢ��ͼƬ������Ҫʱ�ٵ���ˢ�±��
Dim Pictures As New cEPRPictures, rsTemp As New ADODB.Recordset, lngKey As Long, Tables As New cEPRTables
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean, blnForce As Boolean

    '��ȡ���е�ͼƬ
    gstrSQL = "Select ID, �ļ�id,��ʼ��, ��ֹ��,��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���,Ԥ�����ID " & _
        "From ���Ӳ������� " & _
        "Where �ļ�id = [1] And �������� in(3,5) And ������� Is Not Null" '����ʾ����е�ͼƬ
    If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�������", "H���Ӳ�������")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
    Do Until rsTemp.EOF
        If rsTemp!�������� = 5 Then
            lngKey = Pictures.Add(NVL(rsTemp!������, 0))
            Call Pictures("K" & lngKey).FillPictureMember(rsTemp, IIf(blnMoved, "H���Ӳ�������", "���Ӳ�������"))
            Call Pictures("K" & lngKey).DeleteFromEditor(edtThis)
            Call Pictures("K" & lngKey).InsertIntoEditor(edtThis, -1, True)
        ElseIf rsTemp!�������� = 3 Then
            lngKey = Tables.Add(NVL(rsTemp!������, 0))
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
