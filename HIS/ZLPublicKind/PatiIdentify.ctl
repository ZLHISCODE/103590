VERSION 5.00
Begin VB.UserControl PatiIdentify 
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   ScaleHeight     =   990
   ScaleWidth      =   4800
   ToolboxBitmap   =   "PatiIdentify.ctx":0000
   Begin ZLPublicKind.IDKindNew IDKind 
      Height          =   330
      Left            =   150
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   135
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   9
      FontName        =   "����"
      IDKind          =   -1
      BackColor       =   -2147483633
   End
   Begin VB.PictureBox picTxtBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1110
      MousePointer    =   3  'I-Beam
      ScaleHeight     =   495
      ScaleWidth      =   3600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   3600
      Begin VB.TextBox txtPatient 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   105
         Width           =   3990
      End
      Begin VB.Shape shpLine 
         BorderColor     =   &H80000003&
         Height          =   435
         Left            =   0
         Top             =   15
         Width           =   2685
      End
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "����:"
      Height          =   180
      Left            =   105
      TabIndex        =   2
      Top             =   675
      Visible         =   0   'False
      Width           =   450
   End
End
Attribute VB_Name = "PatiIdentify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit
Public Enum Pati_ShowCardNo
    ShowNone = 0
    ShowOnlyCardNo = 1
    ShowShortNameAndCardNo = 2
    ShowFullNameAndCardNo = 3
End Enum
Public Enum Input_Appearance
    ShowFlat = 0
    Show3D = 1
    ShowDeepen3D = 2
End Enum
Public Enum Pati_InputBoxAlignment
    Input_Top_Justify = 0
    Input_Down_Justify = 1
    Input_Center = 2
End Enum
Public Enum TextAlignment
    Text_Left_Justify = 0
    Text_Right_Justify = 1
    Text_Center = 2
End Enum
'�¼�����:
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtPatient,txtPatient,-1,KeyDown
Attribute KeyDown.VB_Description = "���û���ӵ�н���Ķ����ϰ��������ʱ������"
'����֮ǰ
Event FindPatiBefore(ByVal objCard As Object, ByRef blnCard As Boolean, ByRef strShowText As String, ByRef objPatiInfor As Object, ByRef blnFindPatied As Boolean, ByRef blnCancel As Boolean)
Event FindPatiArfter(ByVal objCard As Object, ByVal blnCard As Boolean, ByRef ShowName As String, ByRef objHisPatiInfor As Object, _
        ByRef objPatiInfor As Object, ByRef strErrMsg As String, ByRef blnCancel As Boolean)

Event ItemClick(Index As Integer, objCard As Object)   'MappingInfo=IDKInd,IDKInd,-1,ItemClick
Event Click(objCard As Object)   'MappingInfo=IDKInd,IDKInd,-1,Click
Event KeyPress(KeyAscii As Integer)
Event Change()

'ȱʡ����ֵ:
Const m_def_MustBrushCard = False
Const m_def_InputBoxAlignment = Pati_InputBoxAlignment.Input_Center
Const m_def_HiddenMoseRightKey = True
Const m_def_ShowCardNO = Pati_ShowCardNo.ShowNone
Const m_def_FindPatiShowName = True
Const m_def_���˲���ID = 0
Const m_def_IDKindWidth = 0
Const m_def_OnlyReadDefaultCard = True
Const m_def_ReadThreePatiInfor = True
Const m_def_InputAppearance = 1
Const m_def_OnlyThreeCard = False
'���Ա���:
Private m_MustBrushCard As Boolean
Private m_InputBoxAlignment As Pati_InputBoxAlignment
Private m_HiddenMoseRightKey As Boolean
Private m_ShowCardNO As Pati_ShowCardNo
Private m_FindPatiShowName As Boolean
Private m_���˲���ID As Long
Private m_IDKindWidth As Long
Private m_OnlyReadDefaultCard As Boolean
Private m_ReadThreePatiInfor As Boolean
Private mblnNotAutoSel As Boolean
Private m_InputAppearance As Input_Appearance
Private m_OnlyThreeCard As Boolean
'----------------------------------------------------------------------
'�ؼ�����
Private mobjParent As Object
Private mobjCurPati As clsPatiInfor
Private mobjCurCardData As clsPatiInfor   '��ǰ�������������ݻ�IC����洢�Ĳ�����Ϣ
Private mstrInput As String
Private mblnInternal As Boolean
Private mcnOracle As ADODB.Connection
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents
Attribute mobjCommEvents.VB_VarHelpID = -1
Private Sub IDKind_Click(objCard As Object)
    RaiseEvent Click(objCard)
End Sub

Private Sub IDKind_LostFocus()
    Call IDKind.Refrash
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As Object, objPatiInfor As Object, blnCancel As Boolean)
    Dim lng����ID As Long, strErrMsg As String
    If txtPatient.Enabled = False Or txtPatient.Locked Then Exit Sub
    If FindPatiBefor(objPatiInfor.����, True, objCard, objPatiInfor) = True Then
        Exit Sub
    End If
    Call GetPatient(True, objPatiInfor.����, objCard, lng����ID, strErrMsg)
End Sub

Private Function FindPatiBefor(ByVal strInput As String, ByVal blnCard As Boolean, _
    objCard As Object, objPati As clsPatiInfor) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���֮ǰ
    '���:strInput-�����ֵ��ˢ���������ֵ
    '        blnCard-�Ƿ�ˢ�������(True,��ʾˢ�������)����Ϊ�ֹ�����
    '        objCard-��ǰ�Ŀ����
    '        objPati-��ǰ�����Ĳ�����Ϣ
    '����:
    '����:����Ӧ�ó��򷵻ز��ҵ����˻�Cancel����true,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:2012-09-24 15:13:43
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnFindPatied As Boolean, blnCancel As Boolean, strShowName As String
    Dim lng����ID As Long, strErrMsg As String
    
    If txtPatient.Enabled = False Or txtPatient.Locked Then FindPatiBefor = True: Exit Function
    
    blnFindPatied = False: blnCancel = False
    If Not blnCard Then
        '����ˢ������
        Set objPati = Nothing
    End If
    mblnInternal = True
    txtPatient.Text = strInput: strShowName = strInput
    mblnInternal = False
    RaiseEvent FindPatiBefore(objCard, blnCard, strShowName, objPati, blnFindPatied, blnCancel)
    If objCard.�ӿ���� = 0 And Not blnCard Then
        txtPatient.PasswordChar = ""
    End If
    If Trim(txtPatient.Text) <> strShowName Or strShowName = "" Then txtPatient.PasswordChar = ""
    mblnInternal = True
    txtPatient.Text = strShowName
    mblnInternal = False
    Set mobjCurCardData = objPati
    '������ҵ�,ֱ�ӾͲ���������������
    If blnFindPatied Then
        If FindPatiShowName And Not objPati Is Nothing Then
             mblnInternal = True
             txtPatient.Text = objPati.����
             mblnInternal = False
             txtPatient.PasswordChar = ""
        End If
        txtPatient.SelStart = Len(txtPatient.Text)
        Call SelPati
        Call SetShowCardNo(objCard, objPati)
        FindPatiBefor = True
        Exit Function
    End If
    txtPatient.SelStart = Len(txtPatient.Text)
    Call SelPati
    If blnCancel Then
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        TxtSelAll txtPatient
        Call SetShowCardNo(objCard, objPati)
        FindPatiBefor = True
        Exit Function
    End If
End Function


 Private Sub FindPatiedArfter(ByVal strInput As String, ByVal blnCard As Boolean, _
    objCard As Object, objPati As Object)
    Dim blnCancel As Boolean, strShowName As String, strErrMsg As String
    Dim lng����ID As Long
    
    If Not objPati Is Nothing Then
        lng����ID = objPati.����ID
    End If
    If lng����ID = 0 Then Set objPati = Nothing
    blnCancel = False: strShowName = strInput
    RaiseEvent FindPatiArfter(objCard, blnCard, strShowName, objPati, mobjCurCardData, strErrMsg, blnCancel)
    Set mobjCurPati = objPati
    mstrInput = strInput
    If Trim(txtPatient.Text) <> strShowName Or strShowName = "" Then txtPatient.PasswordChar = ""
    mblnInternal = True
    txtPatient.Text = strShowName
    mblnInternal = False
    If blnCancel Then
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        TxtSelAll txtPatient
        Call SetShowCardNo(objCard, objPati)
        Exit Sub
    End If
    If FindPatiShowName And Not objPati Is Nothing Then
         mblnInternal = True
         txtPatient.Text = objPati.����
         mblnInternal = False
         txtPatient.PasswordChar = ""
    End If
    txtPatient.SelStart = Len(txtPatient.Text)
    Call SelPati
    If objPati Is Nothing Then
        Set objPati = New clsPatiInfor
        If blnCard Then objPati.���� = strInput
    End If
    
    Call SetShowCardNo(objCard, objPati)
End Sub


Private Sub SetShowCardNo(ByVal objCard As Card, ByVal objPati As clsPatiInfor)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ʾ�Ŀ���
    '���:objCard-��ǰ�Ŀ����
    '        objPati-��ǰ�Ŀ���Ϣ
    '����:���˺�
    '����:2012-09-24 11:35:52
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strName As String, strCardNo As String
    lbl����.Caption = ""
    If objCard Is Nothing Or objPati Is Nothing Or ShowCardNo = ShowNone Then Exit Sub
    If objPati.���� = "" Then Exit Sub
    
    strCardNo = objCard.zlCardNOEncrypt(objPati.����)
    If ShowCardNo = ShowShortNameAndCardNo Then
        If objCard.�ӿ���� > 0 Then
             lbl����.Caption = "����(" & objCard.���� & "):" & strCardNo
        ElseIf Not (objCard.���� Like "����*" Or objCard.�Ƿ�ģ������) Then
            lbl����.Caption = objCard.���� & ":" & strCardNo
        End If
        Exit Sub
    End If
    If ShowCardNo = ShowFullNameAndCardNo Then
        If objCard.�ӿ���� > 0 Then
             lbl����.Caption = objCard.���� & "����:" & strCardNo
        ElseIf Not (objCard.���� Like "����*" Or objCard.�Ƿ�ģ������) Then
            lbl����.Caption = objCard.���� & ":" & strCardNo
        End If
        Exit Sub
    End If
    If ShowCardNo = ShowOnlyCardNo Then
        If objCard.�ӿ���� > 0 Then
             lbl����.Caption = strCardNo
        ElseIf Not (objCard.���� Like "����*" Or objCard.�Ƿ�ģ������) Then
            lbl����.Caption = strCardNo
        End If
        Exit Sub
    End If
    
End Sub
Private Sub picTxtBack_Click()
   If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
End Sub

Private Sub picTxtBack_GotFocus()
  If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Sub
Private Sub picTxtBack_Resize()
    Call TxtMove
End Sub
Private Sub TxtMove()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ı���ؼ��ƶ�
    '����:���˺�
    '����:2012-09-25 16:11:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngStep As Long
    Err = 0: On Error Resume Next
    If InputAppearance = ShowFlat Or InputAppearance = Show3D Then lngStep = 20
    If InputAppearance = ShowDeepen3D Then lngStep = 40
    'lngStep = IIf(InputAppearance = ShowFlat, 20, 40)
    With picTxtBack
        shpLine.Left = .ScaleLeft
        shpLine.Width = .ScaleWidth
        shpLine.Height = .ScaleHeight
        shpLine.Top = .ScaleTop
    End With
    With picTxtBack
        Select Case InputBoxAlignment
        Case Input_Top_Justify
            With picTxtBack
                txtPatient.Left = .ScaleLeft + lngStep
                txtPatient.Top = .ScaleTop + lngStep
                txtPatient.Width = .ScaleWidth - txtPatient.Left * 2
                txtPatient.Height = .ScaleHeight - txtPatient.Top * 2
            End With
        Case Input_Down_Justify
            txtPatient.Height = .TextHeight("���˺�")
            txtPatient.Top = .ScaleHeight - txtPatient.Height - lngStep
            txtPatient.Left = .ScaleLeft + lngStep
            txtPatient.Width = .ScaleWidth - txtPatient.Left * 2
        Case Else
            txtPatient.Height = .TextHeight("���˺�")
            txtPatient.Top = (.ScaleHeight - txtPatient.Height - 2 * lngStep) \ 2
            txtPatient.Left = .ScaleLeft + lngStep
            txtPatient.Width = .ScaleWidth - txtPatient.Left * 2
        End Select
    End With
End Sub

Private Sub txtPatient_Change()
    If txtPatient.Visible = False Or txtPatient.Locked = True Or txtPatient.Enabled = False Then Exit Sub
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
    If mblnInternal = True Then Exit Sub
    RaiseEvent Change
End Sub

Private Sub txtPatient_GotFocus()
    If txtPatient.Visible = False Or txtPatient.Locked = True Or txtPatient.Enabled = False Then Exit Sub
    Call IDKind.SetAutoReadCard(Trim(txtPatient.Text) = "")
    Call SetBrushCardObject(True)
    If mblnNotAutoSel Then
        mblnNotAutoSel = False
        txtPatient.SelStart = 1
        Exit Sub
    End If
    TxtSelAll txtPatient
End Sub

Private Function SetBrushCardObject(ByVal blnComm As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ���ӿ�
    '����: true-�ɹ���false-ʧ��
    '����:���ϴ�
    '����:2015/7/23 10:37:39
    '����:85565
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String, objCurCard As Card
    Dim objPubOneCard As clsPublicOneCard
    
    Err = 0: On Error Resume Next
    SetBrushCardObject = True
    
    Set objCurCard = IDKind.GetCurCard
    If objCurCard Is Nothing Then
        Set objCurCard = IDKind.GetfaultCard
    Else
        If objCurCard.�ӿ���� = 0 And objCurCard.���� Like "*����*" Then Set objCurCard = IDKind.GetfaultCard
    End If
    
    If objCurCard Is Nothing Then Exit Function
    If objCurCard.�ӿ���� = 0 Or objCurCard.�ӿڳ����� = "" Or Not (objCurCard.�Ƿ�ˢ�� Or objCurCard.�Ƿ�ɨ��) Then Exit Function
    
    If zlGetPubOneCard(mcnOracle, objPubOneCard) = False Then Exit Function
    
    If objPubOneCard.objThirdSwap.zlSetBrushCardObject(objCurCard.�ӿ����, IIf(blnComm, txtPatient, Nothing), strExpend, objCurCard.���ѿ�) Then
        If mobjCommEvents Is Nothing And AllowAutoCommCard = True Then
            Set mobjCommEvents = New clsCommEvents
            Call objPubOneCard.objThirdSwap.zlInitEvents(UserControl.hWnd, mobjCommEvents)
        End If
    End If
    Set objPubOneCard = Nothing
End Function


Private Sub txtPatient_LostFocus()
    If txtPatient.Visible = False Or txtPatient.Locked = True Or txtPatient.Enabled = False Then Exit Sub
    Call IDKind.SetAutoReadCard(False)
    Call SetBrushCardObject(False)
End Sub

 Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Or HiddenMoseRightKey = False Then Exit Sub
    glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
    Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
End Sub
Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Or HiddenMoseRightKey = False Then Exit Sub
    Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    '��Ҫ����Ƿ�ˢ��
    Dim blnCard As Boolean  '�Ƿ�ˢ��
    Dim objDefault As Card
    Dim blnCardPass As Boolean  ' ���ż�����ʾ
    Dim lngDefaultLen As Long, strInputText As String, blnCancel As Boolean, strErrMsg As String
    Dim lng����ID As Long, objPati As New clsPatiInfor, objThreePati As New clsPatiInfor
    Dim objCard As Card, blnFindPatied As Boolean
    Dim blnPass As Boolean
    
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If IDKind.GetCurCard Is Nothing Then Exit Sub
    RaiseEvent KeyPress(KeyAscii)
    If KeyAscii = 0 Then Exit Sub
    
    If IDKind.GetCurCard.���� Like "*��*��*" Or IDKind.GetCurCard.ģ�������� Then
        '103563,ֻҪ����ĵ�һ���ַ��ǡ�-+*����������ȫ���֣�����Ϊ����ˢ��
        If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
            blnPass = txtPatient.PasswordChar <> ""
            blnCard = InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
            txtPatient.IMEMode = 0
            blnPass = txtPatient.PasswordChar = "" And blnPass
            If blnPass Then
                If txtPatient.SelLength = Len(txtPatient.Text) Then
                    mblnInternal = True
                    txtPatient.Text = ""
                    mblnInternal = False
                End If
                SendKeys Chr(KeyAscii): KeyAscii = 0: Exit Sub
            End If
        End If
        
        '76196
        If FindPatiShowName = False And KeyAscii = 13 And Not mobjCurPati Is Nothing And mstrInput = txtPatient.Text Then
            Call FindPatiedArfter(mstrInput, blnCard, IDKind.GetCurCard, mobjCurPati)
            Exit Sub
        End If
        If Not (blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
            Or KeyAscii = 13 And Trim(txtPatient.Text) <> "") Then Exit Sub
        'ˢ��
        If KeyAscii <> 13 Then
             mblnInternal = True
             txtPatient.Text = txtPatient.Text & Chr(KeyAscii): txtPatient.SelStart = Len(txtPatient.Text)
             mblnInternal = False
        End If
        KeyAscii = 0
        strInputText = txtPatient.Text
        Set objCard = Nothing
        If blnCard Then
            objPati.���� = strInputText
            Set objCard = IDKind.GetfaultCard
        End If
        If objCard Is Nothing Then
            Set objCard = IDKind.GetCurCard
        End If
        
        Set objPati = Nothing
        If FindPatiBefor(strInputText, blnCard, objCard, objPati) = True Then
            Exit Sub
        End If
        If Not blnCard Then
            '����ˢ��ʱ
            If Left(strInputText, 1) = "-" And IsNumeric(Mid(strInputText, 2)) Then
                '����ID
                '78219:���ϴ�,2014/9/23,���Ʋ���ID�Ĳ�ѯ��Χ��long��������֮��
                If Val(Mid(strInputText, 2)) > 2147483647 Then
                    MsgBox "�����ID���������������ʹ��������ѯ��ʽ��", vbInformation, gstrSysName
                    Exit Sub
                End If
                lng����ID = Val(Mid(strInputText, 2))
                Call GetPatiInforFromPatiID(mcnOracle, lng����ID, objPati, strErrMsg)
            ElseIf Left(strInputText, 1) = "*" And IsNumeric(Mid(strInputText, 2)) Then '�����
                Call GetPatiInforFromPatiID(mcnOracle, 0, objPati, strErrMsg, "�����", Mid(strInputText, 2))
            ElseIf Left(strInputText, 1) = "+" And IsNumeric(Mid(strInputText, 2)) Then 'סԺ��
                Call GetPatiInforFromPatiID(mcnOracle, 0, objPati, strErrMsg, "סԺ��", Mid(strInputText, 2))
            ElseIf Left(strInputText, 1) = "/" Then
                '���������⴦��:�޲���,���ܸ��ݴ���������
                If ���˲���ID = 0 Then Exit Sub
                lng����ID = zlGetPatiIDFromBedNumber(mcnOracle, ���˲���ID, Mid(strInputText, 2))
                If lng����ID <> 0 Then Call GetPatiInforFromPatiID(mcnOracle, lng����ID, objPati, strErrMsg)
            ElseIf IsMobileNo(strInputText) Then
                '103000�����ϴ���2017/2/7�����ֻ��Ų���
                If GetPatiIDFromCardType(mcnOracle, "�ֻ���", strInputText, False, lng����ID, , strErrMsg) = False Then lng����ID = 0
                If lng����ID = 0 Then GoTo NotFound
                Call GetPatiInforFromPatiID(mcnOracle, lng����ID, objPati, strErrMsg)
            End If
            Call FindPatiedArfter(strInputText, blnCard, objCard, objPati)
            Exit Sub
        End If
        Set objCard = IDKind.GetCurCard
        Call GetPatient(blnCard, strInputText, objCard, lng����ID, strErrMsg, IDKind.GetDefaultCardTypeID, IDKind.Cards.��ȱʡ������)
        Exit Sub
    End If
    If InStr(1, ",�����,סԺ��,�ֻ���,", "," & IDKind.GetCurCard.���� & ",") > 0 Then
        'ֻ��Ϊ����
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack And InStr(",3,22,24,", "," & KeyAscii & ",") = 0 Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    End If
    
    blnCard = zlIsBrushCard(txtPatient, KeyAscii)
    If IDKind.GetCurCard.�ӿ���� > 0 And MustBrushCard Then
        '����ˢ��
        If Not blnCard And KeyAscii <> 13 And KeyAscii <> 8 Then
            If KeyAscii > 32 Then
                mblnInternal = True
                txtPatient.Text = Chr(KeyAscii): txtPatient.SelStart = 1
                mblnInternal = False
            End If
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    If Not (blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txtPatient.Text) <> "") Then Exit Sub

    If KeyAscii <> 13 Then
        mblnInternal = True
        txtPatient.Text = txtPatient.Text & Chr(KeyAscii): txtPatient.SelStart = Len(txtPatient.Text)
        mblnInternal = False
    End If
    KeyAscii = 0
    strInputText = txtPatient.Text
    
    Set objPati = Nothing
    Set objCard = IDKind.GetCurCard
    If FindPatiBefor(strInputText, blnCard, objCard, objPati) = True Then Exit Sub
    If blnCard Then 'ˢ��
        Call GetPatient(blnCard, strInputText, objCard, lng����ID, strErrMsg): Exit Sub
    End If
    '����ˢ��ʱ
    If Left(strInputText, 1) = "-" And IsNumeric(Mid(strInputText, 2)) Then
        '����ID
        lng����ID = Val(Mid(strInputText, 2))
        Call GetPatiInforFromPatiID(mcnOracle, lng����ID, objPati, strErrMsg)
    ElseIf Left(strInputText, 1) = "*" And IsNumeric(Mid(strInputText, 2)) Then '�����
        Call GetPatiInforFromPatiID(mcnOracle, 0, objPati, strErrMsg, "�����", Mid(strInputText, 2))
    ElseIf Left(strInputText, 1) = "+" And IsNumeric(Mid(strInputText, 2)) Then 'סԺ��
        Call GetPatiInforFromPatiID(mcnOracle, 0, objPati, strErrMsg, "סԺ��", Mid(strInputText, 2))
    ElseIf Left(strInputText, 1) = "/" Then
        '���������⴦��:�޲���,���ܸ��ݴ���������
        If ���˲���ID = 0 Then GoTo NotFound:
        lng����ID = zlGetPatiIDFromBedNumber(mcnOracle, ���˲���ID, Mid(strInputText, 2))
        If lng����ID <> 0 Then Call GetPatiInforFromPatiID(mcnOracle, lng����ID, objPati, strErrMsg)
    ElseIf IsMobileNo(strInputText) Then
        '103000�����ϴ���2017/2/7�����ֻ��Ų���
        If GetPatiIDFromCardType(mcnOracle, "�ֻ���", strInputText, False, lng����ID, , strErrMsg, 0) = False Then lng����ID = 0
        If lng����ID = 0 Then GoTo NotFound
        Call GetPatiInforFromPatiID(mcnOracle, lng����ID, objPati, strErrMsg)
    Else
        Call GetPatient(blnCard, strInputText, objCard, lng����ID, strErrMsg)
        Exit Sub
    End If
NotFound:
    Call FindPatiedArfter(strInputText, blnCard, objCard, objPati)
End Sub
Private Sub SelPati()
    'ѡ�в����ı�����Ϣ
    If ActiveControl Is txtPatient Then
        TxtSelAll txtPatient
    End If
End Sub
Private Sub UserControl_GotFocus()
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Sub
Private Sub UserControl_Initialize()
    glngInstanceCount = glngInstanceCount + 1
End Sub

Private Sub UserControl_Terminate()
    Err = 0: On Error Resume Next
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
    If Not mcnOracle Is Nothing Then Set mcnOracle = Nothing
    If Not mobjCurCardData Is Nothing Then Set mobjCurCardData = Nothing
    If Not mobjParent Is Nothing Then Set mobjParent = Nothing
    If Not mobjCurPati Is Nothing Then Set mobjCurPati = Nothing
End Sub
Private Sub UserControl_Resize()
    Dim lngHeight As Long
    Err = 0: On Error Resume Next
    lngHeight = IIf(lbl����.Visible = False, 0, lbl����.Height + 10)
    lbl����.Left = ScaleLeft
    lbl����.Top = ScaleHeight - lbl����.Height
    
    With IDKind
        .Left = ScaleLeft
        .Top = ScaleTop
        .Height = ScaleHeight - lngHeight
    End With
    With picTxtBack
        .Left = IDKind.Left + IDKind.Width + 20
        .Top = IDKind.Top
        .Height = IDKind.Height
        .Width = ScaleWidth - .Left
    End With
    Call SetInputAppearance
End Sub
Private Sub IDKInd_ItemClick(Index As Integer, objCard As Object)
    RaiseEvent ItemClick(Index, objCard)
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Call UserControl_Resize
End Sub
 

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKInd,IDKInd,-1,Cards
Public Property Get Cards() As Object
    Set Cards = IDKind.Cards
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=txtPatient,txtPatient,-1,Font
Public Property Get Font() As Font
    Set Font = txtPatient.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtPatient.Font = New_Font
    Set picTxtBack.Font = New_Font
    PropertyChanged "Font"
    Call UserControl_Resize
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKInd,IDKInd,-1,IDKindStr
Public Property Get IDKindStr() As String
    IDKindStr = IDKind.IDKindStr
End Property
Public Property Let IDKindStr(ByVal New_IDKindStr As String)
    IDKind.IDKindStr() = New_IDKindStr
    PropertyChanged "IDKindStr"
End Property
'
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKInd,IDKInd,-1,Font
Public Property Get IDKindFont() As Font
    Set IDKindFont = IDKind.Font
End Property
Public Property Set IDKindFont(ByVal New_IDKindFont As Font)
    Set IDKind.Font = New_IDKindFont
    PropertyChanged "IDKindFont"
End Property
'�Ӵ������м�������ֵ
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set txtPatient.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set picTxtBack.Font = picTxtBack.Font
    IDKind.IDKindStr = PropBag.ReadProperty("IDKindStr", "")
    
    Set IDKind.Font = PropBag.ReadProperty("IDKindFont", Ambient.Font)
    IDKind.AutoSize = PropBag.ReadProperty("AutoSize", False)
    IDKind.Appearance = PropBag.ReadProperty("IDKindAppearance", 1)
    m_InputAppearance = PropBag.ReadProperty("InputAppearance", m_def_InputAppearance)
    IDKind.ShowSortName = PropBag.ReadProperty("ShowSortName", False)
    IDKind.CaptionAlignment = PropBag.ReadProperty("CaptionAlignment", 2)
    m_ReadThreePatiInfor = PropBag.ReadProperty("ReadThreePatiInfor", m_def_ReadThreePatiInfor)
    IDKind.ShowPropertySet = PropBag.ReadProperty("ShowPropertySet", False)
    IDKind.DefaultCardType = PropBag.ReadProperty("DefaultCardType", "")
    m_OnlyReadDefaultCard = PropBag.ReadProperty("OnlyReadDefaultCard", m_def_OnlyReadDefaultCard)
    IDKind.BorderStyle = PropBag.ReadProperty("IDkindBorderStyle", 0)

    Call UserControl_Resize: Call SetInputAppearance
 
    m_IDKindWidth = PropBag.ReadProperty("IDKindWidth", IDKind.Width)
    IDKind.Width = m_IDKindWidth
    Dim blnEnable As Boolean
    blnEnable = PropBag.ReadProperty("Enabled", True)
    UserControl.Enabled = blnEnable
    IDKind.Enabled = blnEnable
    m_���˲���ID = PropBag.ReadProperty("���˲���ID", m_def_���˲���ID)
    m_OnlyThreeCard = PropBag.ReadProperty("OnlyThreeCard", m_def_OnlyThreeCard)
    m_FindPatiShowName = PropBag.ReadProperty("FindPatiShowName", m_def_FindPatiShowName)
    mblnInternal = True
    txtPatient.Text = PropBag.ReadProperty("Text", "")
    mblnInternal = False
    txtPatient.Locked = PropBag.ReadProperty("Locked", False)
    m_ShowCardNO = PropBag.ReadProperty("ShowCardNO", m_def_ShowCardNO)
    Call InitFace
    m_HiddenMoseRightKey = PropBag.ReadProperty("HiddenMoseRightKey", m_def_HiddenMoseRightKey)
    Set lbl����.Font = PropBag.ReadProperty("CardNoShowFont", Ambient.Font)
    lbl����.ForeColor = PropBag.ReadProperty("CardNOForColor", &H80000012)
    m_InputBoxAlignment = PropBag.ReadProperty("InputBoxAlignment", m_def_InputBoxAlignment)
    
    Call TxtMove
    txtPatient.Alignment = PropBag.ReadProperty("TextAlignment", 0)
    m_MustBrushCard = PropBag.ReadProperty("MustBrushCard", m_def_MustBrushCard)
    IDKind.AllowAutoCommCard = PropBag.ReadProperty("AllowAutoCommCard", False)
    IDKind.AllowAutoICCard = PropBag.ReadProperty("AllowAutoICCard", False)
    IDKind.AllowAutoIDCard = PropBag.ReadProperty("AllowAutoIDCard", False)
    IDKind.IDKind = PropBag.ReadProperty("IDKindIDX", 1)
    IDKind.NotContainFastKey = PropBag.ReadProperty("NotContainFastKey", "")
    IDKind.NotAutoAppendKind = PropBag.ReadProperty("NotAutoAppendKind", False)
    txtPatient.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtPatient.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    picTxtBack.BackColor = txtPatient.BackColor
    txtPatient.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtPatient.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtPatient.SelText = PropBag.ReadProperty("SelText", "")
    txtPatient.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    txtPatient.IMEMode = PropBag.ReadProperty("IMEMode", 0)
    mblnNotAutoSel = PropBag.ReadProperty("NotAutoSel", False)
    Call SetInputAppearance
End Sub


'������ֵд���洢��
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", txtPatient.Font, Ambient.Font)
    Call PropBag.WriteProperty("IDKindStr", IDKind.IDKindStr, "")
    Call PropBag.WriteProperty("IDKindFont", IDKindFont, Ambient.Font)
    Call PropBag.WriteProperty("AutoSize", IDKind.AutoSize, False)
    Call PropBag.WriteProperty("IDKindAppearance", IDKind.Appearance, 1)
    Call PropBag.WriteProperty("InputAppearance", m_InputAppearance, m_def_InputAppearance)
    Call PropBag.WriteProperty("ShowSortName", IDKind.ShowSortName, False)
    Call PropBag.WriteProperty("CaptionAlignment", IDKind.CaptionAlignment, 2)
    Call PropBag.WriteProperty("ReadThreePatiInfor", m_ReadThreePatiInfor, m_def_ReadThreePatiInfor)
    Call PropBag.WriteProperty("ShowPropertySet", IDKind.ShowPropertySet, False)
    Call PropBag.WriteProperty("DefaultCardType", IDKind.DefaultCardType, "")
    Call PropBag.WriteProperty("OnlyReadDefaultCard", m_OnlyReadDefaultCard, m_def_OnlyReadDefaultCard)
    Call PropBag.WriteProperty("IDkindBorderStyle", IDKind.BorderStyle, 0)
    Call PropBag.WriteProperty("IDKindWidth", m_IDKindWidth, m_def_IDKindWidth)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("���˲���ID", m_���˲���ID, m_def_���˲���ID)
    Call PropBag.WriteProperty("OnlyThreeCard", m_OnlyThreeCard, m_def_OnlyThreeCard)
    Call PropBag.WriteProperty("FindPatiShowName", m_FindPatiShowName, m_def_FindPatiShowName)
    Call PropBag.WriteProperty("Text", txtPatient.Text, "")
    Call PropBag.WriteProperty("Locked", txtPatient.Locked, False)
    Call PropBag.WriteProperty("ShowCardNO", m_ShowCardNO, m_def_ShowCardNO)
    Call PropBag.WriteProperty("HiddenMoseRightKey", m_HiddenMoseRightKey, m_def_HiddenMoseRightKey)
    Call PropBag.WriteProperty("CardNoShowFont", lbl����.Font, Ambient.Font)
    Call PropBag.WriteProperty("CardNOForColor", lbl����.ForeColor, &H80000012)
    Call PropBag.WriteProperty("InputBoxAlignment", m_InputBoxAlignment, m_def_InputBoxAlignment)
    Call TxtMove
    Call PropBag.WriteProperty("TextAlignment", txtPatient.Alignment, 0)
    Call PropBag.WriteProperty("MustBrushCard", m_MustBrushCard, m_def_MustBrushCard)
    Call PropBag.WriteProperty("AllowAutoCommCard", IDKind.AllowAutoCommCard, False)
    Call PropBag.WriteProperty("AllowAutoICCard", IDKind.AllowAutoICCard, False)
    Call PropBag.WriteProperty("AllowAutoIDCard", IDKind.AllowAutoIDCard, False)
    Call PropBag.WriteProperty("IDKindIDX", IDKind.IDKind, 1)
    Call PropBag.WriteProperty("NotContainFastKey", IDKind.NotContainFastKey, "")
    Call PropBag.WriteProperty("NotAutoAppendKind", IDKind.NotAutoAppendKind, False)
    Call PropBag.WriteProperty("ForeColor", txtPatient.ForeColor, &H80000008)
    Call PropBag.WriteProperty("BackColor", txtPatient.BackColor, &H80000005)
    Call PropBag.WriteProperty("SelStart", txtPatient.SelStart, 0)
    Call PropBag.WriteProperty("SelLength", txtPatient.SelLength, 0)
    Call PropBag.WriteProperty("SelText", txtPatient.SelText, "")
    Call PropBag.WriteProperty("PasswordChar", txtPatient.PasswordChar, "")
    Call PropBag.WriteProperty("IMEMode", txtPatient.IMEMode, 0)
    Call PropBag.WriteProperty("NotAutoSel", mblnNotAutoSel, False)
End Sub

Public Function IsMobileNo(ByVal strInput As String) As Boolean
    IsMobileNo = IDKind.IsMobileNo(strInput)
End Function

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKInd,IDKInd,-1,AutoSize
Public Property Get AutoSize() As Boolean
    AutoSize = IDKind.AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    IDKind.AutoSize() = New_AutoSize
    PropertyChanged "AutoSize"
    Call UserControl_Resize
End Property
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKInd,IDKInd,-1,Appearance
Public Property Get IDKindAppearance() As IDKind_Appearance
    IDKindAppearance = IDKind.Appearance
End Property

Public Property Let IDKindAppearance(ByVal New_IDKindAppearance As IDKind_Appearance)
    IDKind.Appearance() = New_IDKindAppearance
    PropertyChanged "IDKindAppearance"
End Property
Private Sub SetInputAppearance()
    Call picTxtBack_Resize
    shpLine.Visible = InputAppearance = 0
    If InputAppearance = ShowDeepen3D Then
        Call zlRaisEffect(picTxtBack, -2, " ")
    ElseIf InputAppearance = Show3D Then
        Call zlRaisEffect(picTxtBack, -1, " ")
    Else
         Call zlRaisEffect(picTxtBack, 0, " ")
    End If
End Sub

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,0
Public Property Get InputAppearance() As Input_Appearance
    InputAppearance = m_InputAppearance
End Property

Public Property Let InputAppearance(ByVal New_InputAppearance As Input_Appearance)
    m_InputAppearance = New_InputAppearance
    PropertyChanged "InputAppearance"
    Call SetInputAppearance
End Property

'Ϊ�û��ؼ���ʼ������
Private Sub UserControl_InitProperties()
    m_InputAppearance = m_def_InputAppearance
    m_ReadThreePatiInfor = m_def_ReadThreePatiInfor
    Call SetInputAppearance
    m_OnlyReadDefaultCard = m_def_OnlyReadDefaultCard
    m_IDKindWidth = m_def_IDKindWidth
    m_���˲���ID = m_def_���˲���ID
    m_OnlyThreeCard = m_def_OnlyThreeCard
    m_FindPatiShowName = m_def_FindPatiShowName
    m_ShowCardNO = m_def_ShowCardNO
    Call InitFace
    m_HiddenMoseRightKey = m_def_HiddenMoseRightKey
    m_InputBoxAlignment = m_def_InputBoxAlignment
    Call TxtMove
    m_MustBrushCard = m_def_MustBrushCard
    mblnNotAutoSel = False
End Sub

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKInd,IDKInd,-1,ShowSortName
Public Property Get ShowSortName() As Boolean
    ShowSortName = IDKind.ShowSortName
End Property

Public Property Let ShowSortName(ByVal New_ShowSortName As Boolean)
    IDKind.ShowSortName() = New_ShowSortName
    PropertyChanged "ShowSortName"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKInd,IDKInd,-1,CaptionAlignment
Public Property Get CaptionAlignment() As IDKind_CaptionAlignment
    CaptionAlignment = IDKind.CaptionAlignment
End Property
Public Property Let CaptionAlignment(ByVal New_CaptionAlignment As IDKind_CaptionAlignment)
    IDKind.CaptionAlignment() = New_CaptionAlignment
    PropertyChanged "CaptionAlignment"
End Property

Public Sub zlInit(ByVal frmMain As Object, Optional ByVal lngSys As Long, Optional ByVal lngModul As Long, _
    Optional cnOracle As ADODB.Connection, Optional strDBUser As String, _
    Optional objPublicOneCard As Object, _
    Optional strIDKindStr As String = "", _
    Optional strProductName As String = "", Optional ByVal blnIsObjRegisterAlone As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��
    '���:frmMain-���õ�������
    '     lngSys : ϵͳ���
    '     lngModul:��Ҫִ�еĹ������
    '     objPublicOneCard-�����㹫������
    '     cnOracle:����������ݿ�����
    '     strIDKindStr-���ʶ��������,�����ָ�ʽ:
    '               һ����ȱʡ��:����1|ȫ��1|������־1;��. ;����n|ȫ��n|������־n
    '               һ������չ��ʽ:����1|ȫ��1|������־1|�����ID1|���ų���1|ȱʡ��־1(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�1(1-�����ʻ�;0-�������ʻ�)|��������1(�ڼ�λ���ڼ�λ����,��Ϊ������);��
    '     strProductName-��Ʒģ������(��Ҫ�������ģ������ı���)
    '     �Ƿ�ʹ�ö�����ע�Ჿ��(True:ʹ��:zlRegisterAlone.DLL,����ʹ��zlRegister.dll)
    '����:���˺�
    '����:2012-08-16 10:45:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mobjParent = frmMain
    Set mcnOracle = cnOracle
    gblnIsObjRegisterAlone = blnIsObjRegisterAlone
    '118959:���ϴ���2018/1/3,��ˢ���ı���
    Call IDKind.zlInit(frmMain, lngSys, lngModul, cnOracle, strDBUser, objPublicOneCard, strIDKindStr, , strProductName, OnlyThreeCard, blnIsObjRegisterAlone)
    Call InitFace: lbl����.Caption = ""
    Call UserControl_Resize
End Sub
'------------------------------------------------------------------------
'ҵ�����
Private Function GetPatient(ByVal blnBrushCard As Boolean, ByVal strInput As String, ByRef objCard As Card, _
      ByRef lng����ID As Long, ByRef strErrMsg As String, Optional ByVal lngDefaultCardTypeID As Long, Optional ByVal blnȱʡ������ As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ID
    '���:strInput-��������ֵ
    '        objCard-��ǰ��ȡ�Ŀ����(����Ϊ����)
    '����:lng����ID-����:����ID
    '        objCard-���ض�ȡ�Ŀ����
    '        strErrMsg-���صĴ�����Ϣ
    '����: �ɹ�,����ָ���Ĳ���ID
    '����:���˺�
    '����:2012-08-20 17:34:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardPassWord As String, lngCardTypeID As Long
    Dim objPati As New clsPatiInfor
    
    On Error GoTo errHandle
    
    'ָ�������
    If objCard.�ӿ���� > 0 Then
       If GetPatiIDFromCardType(mcnOracle, objCard.�ӿ����, strInput, False, lng����ID, strCardPassWord, strErrMsg, lngCardTypeID) = False Then lng����ID = 0
       If lng����ID = 0 Then GoTo NotFindPati:
       
       If lngCardTypeID <> objCard.�ӿ���� And lngCardTypeID <> 0 Then
            Set objCard = IDKind.GetIDKindCard(lngCardTypeID, CardTypeID)
       End If
        Call GetPatiInforFromPatiID(mcnOracle, lng����ID, objPati, strErrMsg)
        objPati.���� = strInput: objPati.���� = strCardPassWord
       Call FindPatiedArfter(strInput, blnBrushCard, objCard, objPati)
       GetPatient = True: Exit Function
    End If
    If objCard.���� Like "����*" Or objCard.ģ�������� Then
        If blnȱʡ������ Then
            lngCardTypeID = lngDefaultCardTypeID
        Else
            lngCardTypeID = -1
        End If
        If blnBrushCard Then
            If GetPatiIDFromCardType(mcnOracle, lngCardTypeID, strInput, True, lng����ID, strCardPassWord, strErrMsg, lngCardTypeID) Then
                If lngCardTypeID <> objCard.�ӿ���� And lngCardTypeID <> 0 Then
                     Set objCard = IDKind.GetIDKindCard(lngCardTypeID, CardTypeID)
                ElseIf Not IDKind.GetfaultCard Is Nothing Then
                    Set objCard = IDKind.GetfaultCard
                End If
            ElseIf IsMobileNo(strInput) Then
                If GetPatiIDFromCardType(mcnOracle, "�ֻ���", strInput, True, lng����ID, strCardPassWord) Then
                    Set objCard = IDKind.GetIDKindCard("�ֻ���", CardTypeName)
                Else
                    lng����ID = 0
                    If strErrMsg <> "" Then MsgBox strErrMsg, vbOKOnly + vbInformation, gstrSysName
                End If
               
            Else
                lng����ID = 0
                If strErrMsg <> "" Then MsgBox strErrMsg, vbOKOnly + vbInformation, gstrSysName
            End If
            
            If lng����ID = 0 Then GoTo NotFindPati:
            Call GetPatiInforFromPatiID(mcnOracle, lng����ID, objPati, strErrMsg)
            objPati.���� = strInput: objPati.���� = strCardPassWord
            Call FindPatiedArfter(strInput, blnBrushCard, objCard, objPati)
            GetPatient = True: Exit Function
        End If
    End If
    Select Case objCard.����
    Case "����"
    Case "ҽ����", "�����", "סԺ��", "���￨", "IC����", "�ֻ���"
        If GetPatiIDFromCardType(mcnOracle, objCard.����, strInput, False, lng����ID, strCardPassWord, strErrMsg, lngCardTypeID) = False Then lng����ID = 0
        If lng����ID = 0 Then GoTo NotFindPati:
        If lngCardTypeID <> objCard.�ӿ���� And lngCardTypeID <> 0 Then
             Set objCard = IDKind.GetIDKindCard(lngCardTypeID, CardTypeID)
        End If
        Call GetPatiInforFromPatiID(mcnOracle, lng����ID, objPati, strErrMsg)
        objPati.���� = strInput: objPati.���� = strCardPassWord
        Call FindPatiedArfter(strInput, blnBrushCard, objCard, objPati)
        GetPatient = True: Exit Function
    Case Else
            GoTo NotFindPati:
    End Select
    GoTo NotFindPati:
     Exit Function
errHandle:
    strErrMsg = Err.Description
NotFindPati:
    objPati.���� = strInput
    Call FindPatiedArfter(strInput, blnBrushCard, objCard, objPati)
End Function

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    Dim objPati As clsPatiInfor, strErrMsg As String
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard(strCardType, CardTypeName)
    If objCard Is Nothing Then Exit Sub
    
    If strXmlCardInfor <> "" Then
       Call zlGetPatiInforFromXML(mcnOracle, strXmlCardInfor, objPati, strErrMsg)
    End If
    If objPati Is Nothing Then Set objPati = New clsPatiInfor
    objPati.���� = strCardNo
    Call IDKind_ReadCard(objCard, objPati, blnCancel)
End Sub


'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,True
Public Property Get ReadThreePatiInfor() As Boolean
    ReadThreePatiInfor = m_ReadThreePatiInfor
End Property

Public Property Let ReadThreePatiInfor(ByVal New_ReadThreePatiInfor As Boolean)
    m_ReadThreePatiInfor = New_ReadThreePatiInfor
    PropertyChanged "ReadThreePatiInfor"
End Property
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKInd,IDKInd,-1,ShowPropertySet
Public Property Get ShowPropertySet() As Boolean
    ShowPropertySet = IDKind.ShowPropertySet
End Property

Public Property Let ShowPropertySet(ByVal New_ShowPropertySet As Boolean)
    IDKind.ShowPropertySet() = New_ShowPropertySet
    PropertyChanged "ShowPropertySet"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKInd,IDKInd,-1,DefaultCardType
Public Property Get DefaultCardType() As String
    DefaultCardType = IDKind.DefaultCardType
End Property

Public Property Let DefaultCardType(ByVal New_DefaultCardType As String)
    IDKind.DefaultCardType() = New_DefaultCardType
    PropertyChanged "DefaultCardType"

End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,True
Public Property Get OnlyReadDefaultCard() As Boolean
    OnlyReadDefaultCard = m_OnlyReadDefaultCard
End Property

Public Property Let OnlyReadDefaultCard(ByVal New_OnlyReadDefaultCard As Boolean)
    m_OnlyReadDefaultCard = New_OnlyReadDefaultCard
    PropertyChanged "OnlyReadDefaultCard"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKind,IDKind,-1,BorderStyle
Public Property Get IDkindBorderStyle() As IDKind_BorderStyle
    IDkindBorderStyle = IDKind.BorderStyle
End Property

Public Property Let IDkindBorderStyle(ByVal New_IDkindBorderStyle As IDKind_BorderStyle)
    IDKind.BorderStyle() = New_IDkindBorderStyle
    PropertyChanged "IDkindBorderStyle"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=8,0,0,0
Public Property Get IDKindWidth() As Long
    IDKindWidth = IDKind.Width
End Property

Public Property Let IDKindWidth(ByVal New_IDKindWidth As Long)
    m_IDKindWidth = New_IDKindWidth
    IDKind.Width = m_IDKindWidth
    PropertyChanged "IDKindWidth"
    Call UserControl_Resize
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    IDKind.Enabled = New_Enabled
    txtPatient.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=8,0,0,0
Public Property Get ���˲���ID() As Long
    ���˲���ID = m_���˲���ID
End Property

Public Property Let ���˲���ID(ByVal New_���˲���ID As Long)
    m_���˲���ID = New_���˲���ID
    PropertyChanged "���˲���ID"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=8,0,0,0
Public Property Get OnlyThreeCard() As Boolean
    OnlyThreeCard = m_OnlyThreeCard
End Property

Public Property Let OnlyThreeCard(ByVal New_OnlyThreeCard As Boolean)
    m_OnlyThreeCard = New_OnlyThreeCard
    PropertyChanged "OnlyThreeCard"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,true
Public Property Get FindPatiShowName() As Boolean
    FindPatiShowName = m_FindPatiShowName
End Property

Public Property Let FindPatiShowName(ByVal New_FindPatiShowName As Boolean)
    m_FindPatiShowName = New_FindPatiShowName
    PropertyChanged "FindPatiShowName"
End Property
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=txtPatient,txtPatient,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "����/���ÿؼ��а������ı���"
    Text = txtPatient.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtPatient.Text() = New_Text
    PropertyChanged "Text"
End Property
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=txtPatient,txtPatient,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "�����ؼ��Ƿ�ɱ༭��"
    Locked = txtPatient.Locked
End Property
Public Property Let Locked(ByVal New_Locked As Boolean)
    txtPatient.Locked() = New_Locked
    PropertyChanged "Locked"
End Property
Public Property Get objTxtInput() As Object
    Set objTxtInput = txtPatient
End Property
Public Property Get objIDKind() As Object
    Set objIDKind = IDKind
End Property
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,0
Public Property Get ShowCardNo() As Pati_ShowCardNo
    ShowCardNo = m_ShowCardNO
End Property

Public Property Let ShowCardNo(ByVal New_ShowCardNO As Pati_ShowCardNo)
    m_ShowCardNO = New_ShowCardNO
    Call InitFace
    Call UserControl_Resize
    PropertyChanged "ShowCardNO"
End Property
Private Sub InitFace()
    lbl����.Visible = ShowCardNo <> ShowNone
End Sub
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,True
Public Property Get HiddenMoseRightKey() As Boolean
Attribute HiddenMoseRightKey.VB_Description = "��������Ҽ��Ŀ�ݲ˵�"
    HiddenMoseRightKey = m_HiddenMoseRightKey
End Property

Public Property Let HiddenMoseRightKey(ByVal New_HiddenMoseRightKey As Boolean)
    m_HiddenMoseRightKey = New_HiddenMoseRightKey
    PropertyChanged "HiddenMoseRightKey"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=lbl����,lbl����,-1,Font
Public Property Get CardNoShowFont() As Font
Attribute CardNoShowFont.VB_Description = "����һ�� Font ����"
    Set CardNoShowFont = lbl����.Font
End Property

Public Property Set CardNoShowFont(ByVal New_CardNoShowFont As Font)
    Set lbl����.Font = New_CardNoShowFont
    PropertyChanged "CardNoShowFont"
    Call UserControl_Resize
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=lbl����,lbl����,-1,ForeColor
Public Property Get CardNOForColor() As OLE_COLOR
Attribute CardNOForColor.VB_Description = "����/���ö������ı���ͼ�ε�ǰ��ɫ��"
    CardNOForColor = lbl����.ForeColor
End Property

Public Property Let CardNOForColor(ByVal New_CardNOForColor As OLE_COLOR)
    lbl����.ForeColor() = New_CardNOForColor
    PropertyChanged "CardNOForColor"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,0
Public Property Get InputBoxAlignment() As Pati_InputBoxAlignment
Attribute InputBoxAlignment.VB_Description = "����ʱ,�����뷽ʽ"
    InputBoxAlignment = m_InputBoxAlignment
End Property
Public Property Let InputBoxAlignment(ByVal New_InputBoxAlignment As Pati_InputBoxAlignment)
    m_InputBoxAlignment = New_InputBoxAlignment
    PropertyChanged "InputBoxAlignment"
    Call TxtMove
End Property
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=txtPatient,txtPatient,-1,Alignment
Public Property Get TextAlignment() As TextAlignment
    TextAlignment = txtPatient.Alignment
End Property

Public Property Let TextAlignment(ByVal New_TextAlignment As TextAlignment)
    txtPatient.Alignment() = New_TextAlignment
    PropertyChanged "TextAlignment"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,False
Public Property Get MustBrushCard() As Boolean
    MustBrushCard = m_MustBrushCard
End Property

Public Property Let MustBrushCard(ByVal New_MustBrushCard As Boolean)
    m_MustBrushCard = New_MustBrushCard
    PropertyChanged "MustBrushCard"
End Property

Public Function zlIsBrushCard(ByVal txtInput As Object, KeyAscii As Integer) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ�ˢ��
    '���:txtInput-�����ı���
    '       KeyAscii
    '����:
    '����:��ˢ������true,���򷵻�False
    '����:���˺�
    '����:2012-09-26 11:05:43
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static sngInputBegin As Single
    Dim sngNow As Single, blnCard As Boolean, strText As String
     'ˢ��ʱ����������ŵ��ɵ��÷�ȡ������
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then Exit Function
    blnCard = False
    '����ǰ�������ʾ������(��δ��ʾ����)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = strText & Chr(KeyAscii)
    End If
    '�ж��Ƿ���ˢ��
     If KeyAscii > 32 Then
        sngNow = Timer
        If txtInput.Text = "" Or strText = "" Then
            sngInputBegin = sngNow
        Else
            If Format((sngNow - sngInputBegin) / Len(strText), "0.000") < 0.04 Then blnCard = True '��һ̨�ʼǱ����ԣ�һ����0.014����
        End If
    End If
    zlIsBrushCard = blnCard
End Function

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKind,IDKind,-1,ActiveFastKey
Public Function ActiveFastKey() As Boolean
    If Trim(txtPatient.Text) <> "" Then Exit Function
    ActiveFastKey = IDKind.ActiveFastKey()
End Function

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKind,IDKind,-1,AllowAutoCommCard
Public Property Get AllowAutoCommCard() As Boolean
    AllowAutoCommCard = IDKind.AllowAutoCommCard
End Property

Public Property Let AllowAutoCommCard(ByVal New_AllowAutoCommCard As Boolean)
    IDKind.AllowAutoCommCard() = New_AllowAutoCommCard
    PropertyChanged "AllowAutoCommCard"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKind,IDKind,-1,AllowAutoICCard
Public Property Get AllowAutoICCard() As Boolean
    AllowAutoICCard = IDKind.AllowAutoICCard
End Property

Public Property Let AllowAutoICCard(ByVal New_AllowAutoICCard As Boolean)
    IDKind.AllowAutoICCard() = New_AllowAutoICCard
    PropertyChanged "AllowAutoICCard"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKind,IDKind,-1,AllowAutoIDCard
Public Property Get AllowAutoIDCard() As Boolean
    AllowAutoIDCard = IDKind.AllowAutoIDCard
End Property

Public Property Let AllowAutoIDCard(ByVal New_AllowAutoIDCard As Boolean)
    IDKind.AllowAutoIDCard() = New_AllowAutoIDCard
    PropertyChanged "AllowAutoIDCard"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKind,IDKind,-1,GetCardNoLen
Public Property Get GetCardNoLen() As Integer
    GetCardNoLen = IDKind.GetCardNoLen
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKind,IDKind,-1,GetCurCard
Public Property Get GetCurCard() As Object
    Set GetCurCard = IDKind.GetCurCard
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKind,IDKind,-1,GetDefaultCardNoLen
Public Property Get GetDefaultCardNoLen() As Integer
    GetDefaultCardNoLen = IDKind.GetDefaultCardNoLen
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKind,IDKind,-1,GetDefaultCardTypeID
Public Function GetDefaultCardTypeID() As Long
    GetDefaultCardTypeID = IDKind.GetDefaultCardTypeID()
End Function

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKind,IDKind,-1,GetIDKindCard
Public Function GetIDKindCard(ByVal strCardType As String, Optional MachMode As Mach_Mode) As Object
    Set GetIDKindCard = IDKind.GetIDKindCard(strCardType, MachMode)
End Function

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKind,IDKind,-1,GetKindIndex
Public Function GetKindIndex(ByVal strCardType As String) As Integer
    GetKindIndex = IDKind.GetKindIndex(strCardType)
End Function

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKind,IDKind,-1,GetfaultCard
Public Property Get GetfaultCard() As Object
    Set GetfaultCard = IDKind.GetfaultCard
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKind,IDKind,-1,IDKind
Public Property Get IDKindIDX() As Integer
    IDKindIDX = IDKind.IDKind
End Property

Public Property Let IDKindIDX(ByVal New_IDKind As Integer)
    IDKind.IDKind() = New_IDKind
    PropertyChanged "IDKindIDX"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKind,IDKind,-1,NotContainFastKey
Public Property Get NotContainFastKey() As String
    NotContainFastKey = IDKind.NotContainFastKey
End Property

Public Property Let NotContainFastKey(ByVal New_NotContainFastKey As String)
    IDKind.NotContainFastKey() = New_NotContainFastKey
    PropertyChanged "NotContainFastKey"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKind,IDKind,-1,Refrash
Public Sub Refrash()
    Call IDKind.Refrash
End Sub

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKind,IDKind,-1,NotAutoAppendKind
Public Property Get NotAutoAppendKind() As Boolean
    NotAutoAppendKind = IDKind.NotAutoAppendKind
End Property

Public Property Let NotAutoAppendKind(ByVal New_NotAutoAppendKind As Boolean)
    IDKind.NotAutoAppendKind() = New_NotAutoAppendKind
    PropertyChanged "NotAutoAppendKind"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=txtPatient,txtPatient,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "����/���ö������ı���ͼ�ε�ǰ��ɫ��"
    ForeColor = txtPatient.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtPatient.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=txtPatient,txtPatient,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "����/���ö������ı���ͼ�εı���ɫ��"
    BackColor = txtPatient.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtPatient.BackColor() = New_BackColor
    picTxtBack.BackColor() = New_BackColor
    Call SetInputAppearance
    PropertyChanged "BackColor"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=txtPatient,txtPatient,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "����/����ѡ���ı�����ʼ�㡣"
    SelStart = txtPatient.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtPatient.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=txtPatient,txtPatient,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "����/����ѡ�����ַ�����"
    SelLength = txtPatient.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtPatient.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=txtPatient,txtPatient,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "����/���ð�����ǰѡ���ı����ַ�����"
    SelText = txtPatient.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtPatient.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=txtPatient,txtPatient,-1,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "����/����һ��ֵ�������Ƿ��ڿؼ�����ʾ�û������ַ��������ַ���"
    PasswordChar = txtPatient.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    txtPatient.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=txtPatient,txtPatient,-1,IMEMode
Public Property Get IMEMode() As Integer
Attribute IMEMode.VB_Description = "����/�������뷽���༭���ĵ�ǰ����ģʽ��"
    IMEMode = txtPatient.IMEMode
End Property

Public Property Let IMEMode(ByVal New_IMEMode As Integer)
    txtPatient.IMEMode() = New_IMEMode
    PropertyChanged "IMEMode"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=IDKind,IDKind,-1,NotAutoSel
Public Property Get NotAutoSel() As Boolean
    NotAutoSel = mblnNotAutoSel
End Property

Public Property Let NotAutoSel(ByVal New_NotAutoSel As Boolean)
    mblnNotAutoSel = New_NotAutoSel
    PropertyChanged "NotAutoSel"
End Property
Public Function zlGetPatiInforFromPatiID(ByVal lng����ID As Long, ByRef objPatiInfor As Object, ByRef strErrMsg As String, Optional strOtherName As String = "", _
    Optional strOtherValue As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���ID,��ȡ����Ϣ����������Ϣ���ظ����˶���
    '���:lng����ID-����ID
    '     strOtherName-������������:�������,סԺ�ţ�ҽ���ŵ�
    '     strOtherValue-��������ֵ
    '����:objPati-���ز�����Ϣ����
    '     strErrMsg-��������ʱ�����صĴ�����Ϣ
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-10 15:17:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiInforFromPatiID = IDKind.zlGetPatiInforFromPatiID(lng����ID, objPatiInfor, strErrMsg, strOtherName, strOtherValue)
End Function

Public Function zlGetPatiIDFromCardType(ByVal strCardType As String, ByVal strCardNo As String, _
    Optional ByVal blnNotShowErrMsg As Boolean = False, Optional ByRef lng����ID As Long, _
    Optional ByRef strCardPassWord As String, Optional ByRef strErrMsg As String, _
    Optional ByRef lngCardTypeID As Long, Optional objCtl As Object = Nothing, Optional frmMain As Object, _
    Optional blnShowMergePati As Boolean = False, Optional ByRef blnOnlyContractPati As Boolean = False, _
    Optional ByRef blnCertificate As Boolean = False, Optional ByRef blnUserCancel As Boolean = False, _
    Optional ByVal lngShowCardNoTypeID As Long = 0, Optional ByVal blnNotCheckValidDate As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ����ҽ�����Ϳ���,��ȡ��Ӧ�Ĳ���ID
    '���:strCardType-�����,���Ϊ����,��Ϊ�����ID,���Ϊ�ַ�,��Ϊ�������
    '       strCardNo-����
    '       blnNotShowErrMsg-����ʾ�������ʾ��Ϣ
    '       frmMain-���õ�������
    '       objCtl-���õĿؼ�
    '       blnShowMergePati-�����ֶ�����������Ĳ���ʱ,�Ƿ���ʾ�ϲ����ܰ�ť
    '       blnOnlyContractPati-ǩԼ����
    '       blnUserCancel-ѡ�����У��û�ѡ����ȡ��
    '       lngShowCardNoTypeID-���˳���������Ϣʱ������ѡ��������ʾ�Ŀ��ŵĿ����ID,0-��ʾ����ʾ���ţ�>0��ʾ��ʾָ����������ID
    '       blnNotCheckValidDate-�Ƿ�Կ���ֹʹ��ʱ����м��,true-�������ֹʹ��ʱ��,false-���
    '����:strErrMsg-���صĴ�����Ϣ
    '       lng����ID-���صĲ���ID
    '       strCardPass-���ؿ��ŵ�����
    '       lngCardTypeID-���ؿ����ID(0��ʾ����ȷ�������ID)
    '����:��ȡ����ID�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-14 17:07:51
    '˵��:ֻ�д���ҽ�����Ĳŵ��ô˺���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiIDFromCardType = IDKind.zlGetPatiIDFromCardType(strCardType, strCardNo, blnNotShowErrMsg, lng����ID, _
        strCardPassWord, strErrMsg, lngCardTypeID, objCtl, frmMain, blnShowMergePati, blnOnlyContractPati, _
        blnCertificate, blnUserCancel, lngShowCardNoTypeID, blnNotCheckValidDate)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Property Get GetPubOneCardObject() As clsPublicOneCard
    Set GetPubOneCardObject = IDKind.GetPubOneCardObject
End Property
 
