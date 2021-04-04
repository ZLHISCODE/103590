VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmDockInContent 
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
Attribute VB_Name = "frmDockInContent"
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
    conPane_Annex = 3
End Enum
'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mlngDays As Long            '-1��ʾ������ȫ����ȡ 0��ʾ������ǰѡ�в��� >0��ʾ��ȡѡ�в���ǰ��N���ڵĹ�����
Private mlngPatientID As Long       '����ID
Private mlngRecordId As Long        '������¼ID
Private mfrmAnnex As frmDockAnnex    '������������
Private mObjTabEprView As cTableEPR      '�����
Public mIsShowAnnex As Boolean
Public Sub RefreshParameter()
    mlngDays = zlDatabase.GetPara("����������Ԥ��", glngSys, 1251, -1)
End Sub
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_EDIT_COPY
        If Control.Enabled And Control.Visible Then '��ݼ�ִ��ʱ��Ҫ�ж�
            gstrCopyPID = CStr(mlngPatientID)
            Me.edtThis.Copy
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
Dim Pane1 As Pane, pane2 As Pane, pane3 As Pane
    On Error GoTo errHand
    Set mObjTabEprView = New cTableEPR
    mObjTabEprView.InitTableEPR gcnOracle, glngSys, gstrDbOwner
    Set mfrmAnnex = New frmDockAnnex
    
    Set Pane1 = dkpMan.CreatePane(conPane_RichEpr, 1200, 200, DockTopOf, Nothing)
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane2 = dkpMan.CreatePane(conPane_TablEpr, 1200, 200, DockTopOf, Nothing)
    pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane2.Close
 
    
    If Not mIsShowAnnex Then
        Set pane3 = dkpMan.CreatePane(conPane_Annex, 1200, 15, DockBottomOf, Nothing)
        pane3.MinTrackSize.Height = 360 / Screen.TwipsPerPixelY: pane3.MaxTrackSize.Height = 360 / Screen.TwipsPerPixelY
        pane3.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    End If
    
    With dkpMan
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = True
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = False
    End With
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.KeyBindings.Add FCONTROL, Asc("C"), ID_EDIT_COPY
    mlngRecordId = -1
    Call RefreshParameter
    Exit Sub

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function SetRichDocsPos(ByVal lngRecordId As Long) As Boolean
    'ͨ��ID�ȶ�λ���޷���λʱ�ټ���
    Dim lngKSS As Long, lngKSE As Long, lngKES As Long, lngKEE As Long, blnNeed As Boolean, lngKey As Long, lngLen As Long, i As Long
    lngLen = Len(edtThis.Text)
    For i = 0 To lngLen
        If FindNextKey(edtThis, i, "F", lngKey, lngKSS, lngKSE, lngKES, lngKEE, blnNeed) Then
            If edtThis.Range(lngKSE, lngKES).Text = lngRecordId Then
                Call edtThis.Range(lngKSS, lngKEE).ScrollIntoView(cprSPStart)   '  .Selected
                SetRichDocsPos = True
                Exit Function
            End If
            i = lngKEE
        Else
            Exit Function
        End If
    Next
End Function
Public Sub Clear()
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
Public Sub zlRefresh(ByVal lngRecordId As Long, strAnnexRight As String, _
                Optional ByVal blnMoved As Boolean, Optional ByVal blnShowFinal As Boolean, Optional ByVal byteEditType As Byte, _
                Optional ByVal blnAllowDelete As Boolean)
'���ܣ�ˢ�²�����ʾ���ݣ�
'������lngRecordId�����Ӳ�����¼ID��blnPrivacyProtect���Ƿ�������˽����;strAnnexRight-��������Ȩ��,byteEditType=0 RichEdit =1 ���ʽ����;blnViewFile �Ƿ����Ԥ��
    Dim blnPrivacy As Boolean, strClipboard As String
    Dim rs As New ADODB.Recordset, varPar() As String
    Dim collFile As New Collection, lngLen1 As Long, lngLen2 As Long, i As Integer, lngFileID As Long, strIDs As String, lngStart As Long, StrKey As String
    
    On Error GoTo errHand
    DoEvents '360���֣�������һ���OK
    strClipboard = Clipboard.GetText()
    dkpMan.FindPane(conPane_RichEpr).Close
    dkpMan.FindPane(conPane_TablEpr).Close
    dkpMan.FindPane(conPane_Annex).Close
    dkpMan.ShowPane conPane_RichEpr
    If lngRecordId = 0 Then Exit Sub
    
    dkpMan.ShowPane conPane_Annex
    If byteEditType = 1 Then '�����
        mlngRecordId = lngRecordId
        dkpMan.FindPane(conPane_RichEpr).Close
        dkpMan.ShowPane conPane_TablEpr
        Call mObjTabEprView.InitOpenEPR(Me, cprEM_�޸�, cprET_���������, mlngRecordId, False, 0, , , , , , , , blnMoved)
        Call mObjTabEprView.zlRefreshDockfrm
    ElseIf byteEditType = 2 Then '��Ⱦ�����濨ר�ñ༭��
'        ��Ⱦ��ҳ���Ѷ���
    ElseIf byteEditType = 0 Then 'ȫ��ʽ����
        mlngRecordId = lngRecordId
        If SetRichDocsPos(lngRecordId) Then Exit Sub
        
        '�����ĵ�����
        gstrSQL = "Select Count(C.Id) As ��Ŀ, c.����ID,c.��ҳID, c.�ļ�id, c.����ʱ��" & vbNewLine & _
                "From �����ļ��б� F, �����ļ��б� B, ���Ӳ�����¼ C" & vbNewLine & _
                "Where f.���� = b.���� And f.ҳ�� = b.ҳ�� And b.Id = c.�ļ�id And c.Id = [1]" & vbNewLine & _
                "Group By c.����ID,c.��ҳID, c.�ļ�id,c.����ʱ��"
        If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
        mlngPatientID = rs!����ID
        lngFileID = rs!�ļ�ID
        edtThis.Freeze
        edtThis.ReadOnly = False
        edtThis.ForceEdit = True
        edtThis.InProcessing = True
        edtThis.Tag = "LoadFile"
        edtThis.NewDoc
        
        If rs!��Ŀ = 1 Or mlngDays = 0 Then
            '��ȡRTF�ļ�
            Call ReadRTF(edtThis, lngRecordId, blnShowFinal, blnMoved, blnShowFinal)
        Else
            zlCommFun.ShowFlash "���Դ������ڼ��ز������ݣ�"
            strIDs = GetFileRange(rs!�ļ�ID, lngRecordId, Format(rs!����ʱ��, "yyyy-MM-dd HH:mm:ss"), 2, rs!����ID, rs!��ҳID, blnMoved)
            '��ȡ����ҳ����ļ�ID����
            gstrSQL = "Select /*+ rule*/ a.Id" & vbNewLine & _
                        "From ���Ӳ�����¼ A," & LongIDsTable(strIDs, varPar, 3) & vbNewLine & _
                        "Where a.Id = b.ID" & IIf(mlngDays = -1, "", " And A.����ʱ�� Between [1] And [2] ") & vbNewLine & _
                        "Order By a.���, a.����ʱ��"
            If blnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CDate(Format(CDate(rs!����ʱ��) - mlngDays, "yyyy-MM-dd HH:mm:ss")), _
                    CDate(Format(CDate(rs!����ʱ��) + mlngDays, "yyyy-MM-dd HH:mm:ss")), varPar(0), varPar(1), varPar(2), varPar(3), _
                    varPar(4), varPar(5), varPar(6), varPar(7), varPar(8), varPar(9))
            strIDs = ""
            Do Until rs.EOF
                strIDs = strIDs & "," & rs!ID
                rs.MoveNext
            Loop
            strIDs = Mid(strIDs, 2)
            gfrmPublic.edtBuff.Freeze
            gfrmPublic.edtBuff.ReadOnly = False
            gfrmPublic.edtBuff.ForceEdit = True
            gfrmPublic.edtBuff.InProcessing = True
            gfrmPublic.edtBuff.Tag = "LoadFile"
            For i = 0 To UBound(Split(strIDs, ","))
                zlCommFun.ShowFlash "���Դ������ڼ���" & IIf(mlngDays = -1, "", "��ѡ�ļ�ǰ��" & mlngDays & "���") & "��" & i + 1 & "�ݲ������ݣ�"
                '��ȡRTF�ļ�
                Call ReadRTF(gfrmPublic.edtBuff, Split(strIDs, ",")(i), blnShowFinal, blnMoved, blnShowFinal)
                
                '��¼�ļ�ID
                StrKey = "FS(" & Format(i, "00000000") & ",1,0)" & Split(strIDs, ",")(i) & "FE(" & Format(i, "00000000") & ",1,0)"
                'lngLen2 = Len(edtThis.Text) '���ļ���ӵ����ĵ�ĩβ
                gfrmPublic.edtBuff.Range(0, 0).Selected
                gfrmPublic.edtBuff.Range(0, 0).Text = StrKey
                gfrmPublic.edtBuff.Range(0, 0 + Len(StrKey)).Font.Protected = True
                gfrmPublic.edtBuff.Range(0, 0 + Len(StrKey)).Font.Hidden = True
                
                '׷��RTF�ļ�
                lngLen1 = Len(gfrmPublic.edtBuff.Text) '��¼��ʱ�ļ���ʼ������λ��
                lngLen2 = Len(edtThis.Text) '���ļ���ӵ����ĵ�ĩβ
                edtThis.Range(lngLen2, lngLen2).Font.Protected = False
                edtThis.Range(lngLen2, lngLen2).Selected
                gfrmPublic.edtBuff.SelectAll
                gfrmPublic.edtBuff.CopyWithFormat
                edtThis.PasteWithFormat
                lngStart = Len(edtThis.Text)
                If i < UBound(Split(strIDs, ",")) Then
                    'ֻҪ�������һ���ļ���ĩβ��֤��һ���س����Ա�׷����һ���ļ�
                    If edtThis.Range(lngStart - 2, lngStart) = vbCrLf Then
                        edtThis.Range(lngStart - 2, lngStart).Font.Hidden = False
                    Else
                        edtThis.Range(lngStart, lngStart).Text = vbCrLf
                        edtThis.Range(lngStart, lngStart + 2).Font.Hidden = False
                    End If
                End If
                edtThis.TOM.TextDocument.Range(lngStart, lngStart).Para = gfrmPublic.edtBuff.TOM.TextDocument.Range(lngLen1, lngLen1).Para '.Duplicate
            Next
        End If
        
        If lngRecordId > 0 Then
            '����ҳ���ʽ
            Dim mEPRFileInfo As New cEPRFileDefineInfo
            gstrSQL = "Select a.��ʽ From ����ҳ���ʽ a, �����ļ��б� b" & _
                    " Where b.id=[1] And a.���� = b.���� And a.��� = b.ҳ��"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
            If Not rs.EOF Then
                mEPRFileInfo.��ʽ = zlCommFun.NVL(rs("��ʽ").Value)
                mEPRFileInfo.SetFormat Me.edtThis, mEPRFileInfo.��ʽ
                Me.edtThis.ResetWYSIWYG
            End If
            Set mEPRFileInfo = Nothing
        End If
        gfrmPublic.edtBuff.UnFreeze
        gfrmPublic.edtBuff.ForceEdit = False
        edtThis.SelStart = 0
        edtThis.UnFreeze
        edtThis.RefreshTargetDC
        edtThis.ViewMode = cprNormal
        edtThis.ReadOnly = True
        edtThis.ForceEdit = False
        edtThis.InProcessing = False
        edtThis.Tag = ""
        Call SetRichDocsPos(lngRecordId)
    End If
    '���ø����б�
    Call mfrmAnnex.zlRefresh(mlngRecordId, strAnnexRight, blnMoved, blnAllowDelete)

    zlCommFun.StopFlash
    DoEvents '360���֣�������һ���OK
    Clipboard.Clear
    Clipboard.SetText strClipboard
    Exit Sub
errHand:
    zlCommFun.StopFlash
    If ErrCenter = 1 Then
        Resume
    End If
    On Error Resume Next
    gfrmPublic.edtBuff.UnFreeze
    gfrmPublic.edtBuff.ForceEdit = False
    edtThis.SelStart = 0
    edtThis.UnFreeze
    edtThis.RefreshTargetDC
    edtThis.ViewMode = cprNormal
    edtThis.ReadOnly = True
    edtThis.ForceEdit = False
    edtThis.InProcessing = False
    edtThis.Tag = ""
    Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload mfrmAnnex
    Set mfrmAnnex = Nothing
    Unload mObjTabEprView.zlGetForm
    Set mObjTabEprView = Nothing
End Sub

Private Sub picRich_Resize()
    edtThis.Top = 0: edtThis.Left = 0
    edtThis.Width = picRich.ScaleWidth: edtThis.Height = picRich.Height
End Sub
