VERSION 5.00
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#3.4#0"; "zlIDKind.ocx"
Begin VB.Form frmBlackListEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ⲡ��"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   Icon            =   "frmBlackListEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt��� 
      ForeColor       =   &H00C00000&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4245
      MaxLength       =   10
      TabIndex        =   14
      Top             =   1050
      Width           =   885
   End
   Begin VB.TextBox txtNote 
      Height          =   870
      Left            =   105
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1395
      Width           =   5220
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3855
      TabIndex        =   11
      Top             =   2400
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2760
      TabIndex        =   10
      Top             =   2400
      Width           =   1100
   End
   Begin VB.Frame fraPati 
      Height          =   960
      Left            =   105
      TabIndex        =   12
      Top             =   15
      Width           =   5220
      Begin VB.CommandButton cmdPati 
         Height          =   300
         Left            =   2625
         Picture         =   "frmBlackListEdit.frx":06EA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "ѡ����(F2)"
         Top             =   225
         Width           =   300
      End
      Begin VB.TextBox txtPatient 
         Height          =   300
         Left            =   1350
         TabIndex        =   1
         Top             =   225
         Width           =   1275
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   300
         Left            =   705
         TabIndex        =   15
         ToolTipText     =   "��ݼ�F4"
         Top             =   225
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   529
         Appearance      =   2
         IDKindStr       =   $"frmBlackListEdit.frx":0C74
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
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ţ�"
         Height          =   180
         Left            =   3960
         TabIndex        =   7
         Top             =   630
         Width           =   540
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ң�"
         Height          =   180
         Left            =   2325
         TabIndex        =   6
         Top             =   630
         Width           =   540
      End
      Begin VB.Label lbl��ʶ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʶ�ţ�"
         Height          =   180
         Left            =   330
         TabIndex        =   5
         Top             =   645
         Width           =   720
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���䣺"
         Height          =   180
         Left            =   3960
         TabIndex        =   4
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2985
         TabIndex        =   3
         Top             =   285
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   330
         TabIndex        =   0
         Top             =   285
         Width           =   360
      End
   End
   Begin VB.Label lbl��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ǼǱ��"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   3480
      TabIndex        =   13
      Top             =   1110
      Width           =   720
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ò��˼������ⲡ��������ԭ��"
      Height          =   180
      Left            =   135
      TabIndex        =   8
      Top             =   1155
      Width           =   2700
   End
End
Attribute VB_Name = "frmBlackListEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

Private mstrPrivs As String
Private mlng��� As Long
Private mblnDelete As Boolean
Private mblnOK As Boolean
Private mblnNotClick As Boolean


Public Function ShowMe(frmParent As Object, ByVal strPrivs As String, Optional ByVal lng��� As Long, Optional ByVal blnDelete As Boolean) As Boolean
    mlng��� = lng���
    mblnDelete = blnDelete
    mstrPrivs = strPrivs
    mblnNotClick = False
    Me.Show 1, frmParent
    
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim blnTrans As Boolean
    
    If Val(txtPatient.Tag) = 0 Then
        MsgBox "��ȷ��Ҫ�������ⲡ�������Ĳ��ˡ�", vbInformation, gstrSysName
        txtPatient.SetFocus: Exit Sub
    End If
    If Val(txt���.Text) = 0 Then
        MsgBox "��ȷ��Ҫ�������ⲡ�˵ĵǼǱ�š�", vbInformation, gstrSysName
        txt���.SetFocus: Exit Sub
    End If
    If txtNote.Text = "" Then
        MsgBox "������ԭ��", vbInformation, gstrSysName
        txtNote.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txtNote.Text) > txtNote.MaxLength Then
        MsgBox "����������� " & txtNote.MaxLength & " ���ַ��� " & txtNote.MaxLength \ 2 & " �����֡�", vbInformation, gstrSysName
        txtNote.SetFocus: Exit Sub
    End If
    
    If mlng��� = 0 Then
        strSQL = "ZL_���ⲡ��_Insert(" & Val(txt���.Text) & "," & Val(txtPatient.Tag) & ",'" & txtNote.Text & "')"
    Else
        If mblnDelete Then
            strSQL = "ZL_���ⲡ��_Delete(" & mlng��� & ",'" & txtNote.Text & "')"
        Else
            strSQL = "ZL_���ⲡ��_Update(" & mlng��� & "," & Val(txt���.Text) & ",'" & txtNote.Text & "')"
        End If
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
    
Private Sub cmdPati_Click()
    frmPatiSel.mstrPrivs = mstrPrivs
    frmPatiSel.Show 1, Me
    If frmPatiSel.mlng����ID <> 0 Then
        txtPatient.Text = "-" & frmPatiSel.mlng����ID
        mblnNotClick = True
        IDKind.IDKind = IDKind.GetKindIndex("����")
        mblnNotClick = False
        Call txtPatient_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intIndex As Integer
    If KeyCode = vbKeyF4 Then
        If Shift = vbCtrlMask And IDKind.Enabled Then
            intIndex = IDKind.GetKindIndex("IC����")
            If intIndex < 0 Then Exit Sub
            IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
        End If
    ElseIf KeyCode = vbKeyF2 Then
        Call cmdPati_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnOK = False
    Call CreateMobjCard
    Call CreateSquareCardObject(Me, 1101)
     '��ʼ��
    Call IDKind.zlInit(Me, 100, 1101, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    
    If Not gobjSquare.objSquareCard Is Nothing Then
        IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
    End If
    
    If mlng��� <> 0 Then
        txt���.Text = mlng���
        If mblnDelete Then
            Me.Caption = "�������ⲡ��"
            txt���.Enabled = False
        End If
        fraPati.Enabled = False
        txtPatient.Enabled = False
        cmdPati.Visible = False
        Call GetPatient(IDKind.GetCurCard, "���" & mlng���)
    Else
        txt���.Text = GetMaxNum
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand As String
    Dim strOutPatiInforXml As String
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = New clsICCard
            Call mobjICCard.SetParent(Me.hwnd)
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call txtPatient_KeyPress(vbKeyReturn)
            End If
        End If
        Exit Sub
    End If
    
    lng�����ID = objCard.�ӿ����
    If lng�����ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
    If gobjSquare.objSquareCard.zlReadCard(Me, 1101, lng�����ID, False, strExpand, strOutCardNO, strOutPatiInforXml) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    
    Set gobjSquare.objCurCard = objCard
    txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" And mblnNotClick = False Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Text <> "" Or txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNo
    If txtPatient.Text <> "" Then Call FindPati(objCard, False)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("���֤", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, False)
End Sub

Private Sub txtNote_GotFocus()
    Call zlControl.TxtSelAll(txtNote)
End Sub

Private Sub txtNote_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtPatient_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    If Not mobjIDCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then mobjIDCard.SetEnabled (True)
    If Not mobjICCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then mobjICCard.SetEnabled (True)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Function FindPati(ByVal objCard As Card, Optional blnCard As Boolean = False) As Boolean
    If Not GetPatient(objCard, txtPatient.Text, blnCard) Then
        If IsNumeric(txtPatient.Text) Then
            txtPatient.PasswordChar = "": txtPatient.IMEMode = 0: txtPatient.Text = ""
        End If
        Call zlControl.TxtSelAll(txtPatient)
        txtPatient.SetFocus: Exit Function
    Else
        txtPatient.PasswordChar = ""
        txtPatient.IMEMode = 0
        txtNote.SetFocus: Exit Function
    End If
    FindPati = True
End Function

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean = False) As Boolean
'���ܣ���ȡ������Ϣ
    Dim lng�����ID As Long, lng����ID As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnDo As Boolean
    Dim blnCode As Boolean, blnHavePassWord As Boolean
    Dim strPassWord As String, strErrMsg As String
    Dim strCard As String, strPati As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    
    If strInput Like "���*" Then
        blnCode = True
        strSQL = "Select A.����ID,A.����,A.�Ա�,A.����,A.��������,A.����," & _
            " Decode(Nvl(A.סԺ����,0),0,A.�����,A.סԺ��) as ��ʶ��," & _
            " D.���� as ����,B.��Ժ���� as ����,C.����ԭ��,C.����ԭ��,C.����ʱ��" & _
            " From ������Ϣ A,������ҳ B,���ⲡ�� C,���ű� D" & _
            " Where A.����ID=B.����ID(+) And Nvl(A.��ҳID,0)=B.��ҳID(+)" & _
            " And B.��Ժ����ID=D.ID(+) And Nvl(B.��ҳID(+),0)<>0" & _
            " And A.����ID=C.����ID And C.���=[2]"
        strInput = Mid(strInput, 3)
    Else
        blnCode = False
        strSQL = "Select A.����ID,A.����,A.�Ա�,A.����,A.��������,A.����," & _
            " Decode(Nvl(A.סԺ����,0),0,A.�����,A.סԺ��) as ��ʶ��," & _
            " D.���� as ����,B.��Ժ���� as ����" & _
            " From ������Ϣ A,������ҳ B,���ű� D" & _
            " Where A.����ID=B.����ID(+) And Nvl(A.��ҳID,0)=B.��ҳID(+)" & _
            " And B.��Ժ����ID=D.ID(+) And Nvl(B.��ҳID(+),0)<>0 And A.ͣ��ʱ�� is NULL"
            
        If blnCard = True And objCard.���� Like "����*" Then    'ˢ��
            If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
                lng�����ID = IDKind.GetfaultCard.�ӿ����
            Else
                lng�����ID = "-1"
            End If
            '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
            If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
            If lng����ID <= 0 Then GoTo NotFoundPati:
            strInput = "-" & lng����ID
            strSQL = strSQL & " And A.����ID=[1]"
            blnHavePassWord = True
        ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
            strSQL = strSQL & " And A.����ID=[1]"
        ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
            strSQL = strSQL & " And A.סԺ��=[1]"
        ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
            strSQL = strSQL & " And A.�����=[1]"
        Else
            Select Case objCard.����
                Case "����", "��������￨"
                    If gblnShowCard = True Then
                        strCard = "A.���￨�� as ���￨,A.���￨�� as ���￨��,"
                    Else
                        strCard = "LPAD('*',Length(A.���￨��),'*') as ���￨,A.���￨�� as ���￨��,"
                    End If
                    'ͨ������ģ�����Ҳ���(�������벡�˱�ʶʱ)
                    strPati = _
                        " Select A.����ID ID,A.����ID,A.�����,A.סԺ��," & strCard & "A.����,A.�Ա�,A.����,A.�ѱ� as ����ѱ�," & _
                        "   B.���� as ����,C.���� as ����,A.��ǰ���� as ����,To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��," & _
                        "   To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��,A.סԺ����,To_Char(A.��������,'YYYY-MM-DD') as ��������," & _
                        "   A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���,A.���֤��,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��," & _
                        "   Nvl(P.��������,Decode(P.����,Null,'��ͨ����','ҽ������')) ��������" & _
                        " From ������ҳ P,������Ϣ A,���ű� B,���ű� C" & _
                        " Where A.��ǰ����ID=B.ID(+) And A.��ǰ����ID=C.ID(+) And A.����ID=P.����ID(+) And A.��ҳID=P.��ҳID(+)" & _
                        "   And Nvl(P.��ҳID(+),0)<>0 And A.ͣ��ʱ�� is NULL And A.���� Like [1]" & _
                        " Order by A.����,A.�Ǽ�ʱ�� Desc"
                    
                    vRect = zlControl.GetControlRect(txtPatient.hwnd)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%")
                                
                    'ֻ��һ������ʱ,blncancel����false,��ȡ������Ҳ��һ��
                    If Not rsTmp Is Nothing Then
                        strSQL = strSQL & " And A.����ID=[1]"
                        lng����ID = Val(Nvl(rsTmp!����ID))
                        If lng����ID <= 0 Then GoTo NotFoundPati:
                        strInput = "-" & lng����ID
                    ElseIf blnCancel = True Then
                        strSQL = strSQL & " And A.����ID=[1]"
                        lng����ID = Val(txtPatient.Tag)
                        If lng����ID <= 0 Then GoTo NotFoundPati:
                        strInput = "-" & lng����ID
                    Else
                        GoTo NotFoundPati
                    End If
                Case "ҽ����"
                    strInput = UCase(strInput)
                    strSQL = strSQL & " And A.ҽ����=[2]"
                Case "�����"
                    If Not IsNumeric(strInput) Then strInput = "0"
                    strSQL = strSQL & " And A.�����=[2]"
                Case Else
                    '��������,��ȡ��صĲ���ID
                    If Val(objCard.�ӿ����) > 0 Then
                        lng�����ID = Val(objCard.�ӿ����)
                        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                        If lng����ID = 0 Then GoTo NotFoundPati:
                    Else
                        If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                            strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    End If
                    If lng����ID <= 0 Then GoTo NotFoundPati:
                    strSQL = strSQL & " And A.����ID=[1]"
                    strInput = "-" & lng����ID
                    blnHavePassWord = True
            End Select
        End If
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    
    blnDo = Not rsTmp.EOF
    If Not blnCode Then
        If blnDo Then blnDo = PatiAllow(rsTmp!����ID, rsTmp!����)
    End If
    If blnDo Then
        txtPatient.Tag = rsTmp!����ID
        txtPatient.Text = rsTmp!����
        '74426:���ϴ�,2014-7-9,����������ɫ����
        Call SetPatiColor(txtPatient, Nvl(rsTmp!��������), IIf(IsNull(rsTmp!����), Me.ForeColor, vbRed))
        lblSex.Caption = "�Ա�" & Nvl(rsTmp!�Ա�)
        lblAge.Caption = "���䣺" & Nvl(rsTmp!����)
        lbl��ʶ��.Caption = "��ʶ�ţ�" & Nvl(rsTmp!��ʶ��)
        lbl����.Caption = "���ң�" & Nvl(rsTmp!����)
        lbl����.Caption = "���ţ�" & Nvl(rsTmp!����)
        
        '�޸�ʱ�Ŷ�ȡ
        If blnCode Then
            If mblnDelete Then
                lblNote.Caption = "�����˴����������г�����ԭ��"
                txtNote.Text = ""
            ElseIf IsNull(rsTmp!����ʱ��) Then
                lblNote.Caption = "�ò��˼������ⲡ��������ԭ��"
                txtNote.Text = Nvl(rsTmp!����ԭ��)
            Else
                lblNote.Caption = "�����˴����������г�����ԭ��"
                txtNote.Text = Nvl(rsTmp!����ԭ��)
                txt���.Enabled = False
            End If
        End If
            
        GetPatient = True
    Else
NotFoundPati:
        txtPatient.Tag = ""
        txtPatient.Text = ""
        txtPatient.ForeColor = Me.ForeColor
        lblSex.Caption = "�Ա�"
        lblAge.Caption = "���䣺"
        lbl��ʶ��.Caption = "��ʶ�ţ�"
        lbl����.Caption = "���ң�"
        lbl����.Caption = "���ţ�"
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    Call IDKind.ActiveFastKey
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    Dim blnCard As Boolean
    
    If IDKind.GetCurCard.���� = "����" Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
    ElseIf IDKind.GetCurCard.���� = "�����" Or IDKind.GetCurCard.���� = "סԺ��" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        txtPatient.IMEMode = 0
    End If
    
    'ˢ����ϻ���������س�
    If blnCard And Len(Me.txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txtPatient.Text <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        
        '��ȡ������Ϣ
        Call FindPati(IDKind.GetCurCard, blnCard)
    End If
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
    txtPatient.Text = Trim(txtPatient.Text)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt���_GotFocus()
    Call zlControl.TxtSelAll(txt���)
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Function GetMaxNum() As Long
'���ܣ���ȡ�����Բ����еĵ�ǰ�����ñ��(��ȱ)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���+1 as ��� From ���ⲡ�� Minus Select ��� From ���ⲡ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If rsTmp.EOF Then
        GetMaxNum = 1
    Else
        GetMaxNum = rsTmp!���
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function PatiAllow(ByVal lng����ID As Long, ByVal str���� As String) As Boolean
'���ܣ��ж�ָ�������Ƿ���Լ������ⲡ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    'by lesfeng 2010-03-06 �����Ż����󶨱���
    strSQL = "Select ���,����ID,����ԭ��,����ʱ��,�Ǽ���,����ԭ��,����ʱ��,������ From ���ⲡ�� Where ����ʱ�� is Null And ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    If Not rsTmp.EOF Then
        MsgBox str���� & "�Ѿ��������ⲡ��,ԭ��" & vbCrLf & vbCrLf & vbTab & Nvl(rsTmp!����ԭ��, "<û��ԭ��>"), vbInformation, gstrSysName
        Exit Function
    End If
    PatiAllow = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub CreateMobjCard()
    '����������
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hwnd)
    Set mobjICCard = New clsICCard
    Call mobjICCard.SetParent(Me.hwnd)
    Set mobjICCard.gcnOracle = gcnOracle
End Sub
