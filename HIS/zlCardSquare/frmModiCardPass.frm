VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmModiCardPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����޸�"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7425
   Icon            =   "frmModiCardPass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6120
      TabIndex        =   5
      Top             =   360
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6120
      TabIndex        =   6
      Top             =   870
      Width           =   1200
   End
   Begin VB.PictureBox picPass 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   285
      ScaleHeight     =   3735
      ScaleWidth      =   5625
      TabIndex        =   0
      Top             =   285
      Width           =   5625
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1095
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   3030
         Width           =   4245
      End
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1095
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2520
         Width           =   4245
      End
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1095
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2010
         Width           =   4245
      End
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1095
         PasswordChar    =   "*"
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1470
         Width           =   4245
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   120
         X2              =   5520
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblNotes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�뽫[XX]��ˢ�����ϻ�����  Ȼ����������������ͬ�����룡"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         TabIndex        =   11
         Top             =   180
         Width           =   8550
      End
      Begin VB.Image imgFlag 
         Height          =   720
         Left            =   120
         Picture         =   "frmModiCardPass.frx":058A
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��֤"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   390
         TabIndex        =   10
         Top             =   3120
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   75
         TabIndex        =   9
         Top             =   2580
         Width           =   945
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ԭ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   8
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   390
         TabIndex        =   7
         Top             =   1500
         Width           =   630
      End
   End
   Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
      Height          =   4065
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   5955
      _Version        =   589884
      _ExtentX        =   10504
      _ExtentY        =   7170
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmModiCardPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mlngCardTypeID As Long
Private mblnCheckOldPass As Boolean

Private mobjKeyboard As Object, mblnTest As Boolean
Private mblnFirst As Boolean
Private mblnOK As Boolean

Private Enum mTextIndex
    txt_���� = 0
    txt_ԭ���� = 1
    txt_������ = 2
    txt_��֤���� = 3
End Enum

Private Type Ty_CardType '�������Ϣ
    str���� As String
    lng���ų��� As Long
    lng���볤�� As Long
    int���볤������ As Integer
    byt������� As Byte
End Type
Private mTy_CardType As Ty_CardType
Private mlngCardID  As Long
Private mstrOldPassWord As String

Public Function zlModifyPass(ByVal frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, Optional blnCheckOldPass As Boolean = True) As Boolean
    '����:����������ڲ���
    '���:frmMain-���õ�������
    '     lngModule -ģ���
    '     lngCardTypeID-���ѿ��ӿڱ��
    '����:�޸ĳɹ�,����true,���򷵻�false
    mlngModule = lngModule: mlngCardTypeID = lngCardTypeID
    mblnCheckOldPass = blnCheckOldPass
    
    mblnOK = False
    On Error Resume Next
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlModifyPass = mblnOK
End Function

Private Sub Form_Load()
    mblnFirst = True
    
    mblnTest = Val(GetSetting("ZLSOFT", "����ȫ��\zlSquareCard", "TestCardNO", 0)) = 1
    mblnTest = IsDesinMode Or mblnTest
    
    If mblnCheckOldPass = False Then
        lblNotes.Top = 180
        lblNotes.Caption = "�뽫[XX]��ˢ�����ϻ�����" & vbCrLf & "��������������ͬ�������룡"
        txtEdit(mTextIndex.txt_ԭ����).Enabled = False
        txtEdit(mTextIndex.txt_ԭ����).BackColor = &H8000000F
    Else
        lblNotes.Top = 180
        lblNotes.Caption = "�뽫[XX]��ˢ�����ϻ�����" & vbCrLf & "�����������������ͬ�������룡"
    End If
    
    If InitCardInfor() = False Then
        ShowMsgbox "��ǰ�����δ���û��ѱ�ɾ���������ܽ����޸�����������뵽����������>�豸���á����� �����ѿ����������ã�"
        Unload Me: Exit Sub
    End If
    Call ClearFace
    
    Call CreateObjectKeyboard
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    zlControl.ControlSetFocus txtEdit(mTextIndex.txt_����)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then
        'ȥ��������ţ����Ҳ�����ճ��
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Function InitCardInfor() As Boolean
    '����:��ʼ���������Ϣ
    '����:��ʼ���ɹ�,����true,���򷵻�False
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    Set rsTemp = zlGet���ѿ��ӿ�
    rsTemp.Filter = "���=" & mlngCardTypeID
    If rsTemp.EOF Then Exit Function
    
    With mTy_CardType
        .str���� = Nvl(rsTemp!����)
        .lng���ų��� = Val(Nvl(rsTemp!���ų���))
        .lng���볤�� = Val(Nvl(rsTemp!���볤��))
        .int���볤������ = Val(Nvl(rsTemp!���볤������))
        .byt������� = Nvl(rsTemp!�������)
    End With
    
    lblNotes.Caption = Replace(lblNotes.Caption, "[XX]", "[" & mTy_CardType.str���� & "]")
    InitCardInfor = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function OpenPassKeyboard(ctlText As Control, Optional blnȷ������ As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, blnȷ������) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function isValied() As Boolean
    '����:�������������Ƿ���Ч
    '����:������Ч,����true,���򷵻�False
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strCardNo As String, strPassWord As String
    
    On Error GoTo ErrHandler
    strCardNo = Trim(txtEdit(mTextIndex.txt_����).Text)
    If CheckCardNo(mlngCardTypeID, strCardNo) = False Then Exit Function
    
    strPassWord = zlCommFun.zlStringEncode(txtEdit(mTextIndex.txt_ԭ����).Text) '�������
    If strPassWord <> mstrOldPassWord And mblnCheckOldPass Then
        ShowMsgbox "��Ƭԭ������������������������룡"
        zlControl.ControlSetFocus txtEdit(mTextIndex.txt_ԭ����)
        Exit Function
    End If
    
    If txtEdit(mTextIndex.txt_������).Text = "" Then
        If MsgBox("��ǰ���õ�����Ϊ�գ�ȷ��Ҫ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            zlControl.ControlSetFocus txtEdit(mTextIndex.txt_������)
            Exit Function
        End If
    Else
        Select Case mTy_CardType.int���볤������
        Case 0
        Case 1
            If Len(txtEdit(mTextIndex.txt_������).Text) <> mTy_CardType.lng���볤�� Then
                ShowMsgbox "�����������" & mTy_CardType.lng���볤�� & "λ��"
                zlControl.ControlSetFocus txtEdit(mTextIndex.txt_������)
                zlControl.TxtSelAll txtEdit(mTextIndex.txt_������)
                Exit Function
             End If
        Case Else
            If Len(txtEdit(mTextIndex.txt_������).Text) <= Abs(mTy_CardType.int���볤������) Then
                ShowMsgbox "�����������" & Abs(mTy_CardType.int���볤������) & "λ���ϣ�"
                zlControl.ControlSetFocus txtEdit(mTextIndex.txt_������)
                zlControl.TxtSelAll txtEdit(mTextIndex.txt_������)
                Exit Function
             End If
        End Select
        If mTy_CardType.byt������� = 1 Then '����ֻ����Ϊ����
            If IsNumeric(txtEdit(mTextIndex.txt_������).Text) = False Then
                ShowMsgbox "����ֻ�ܰ������֣����������룡"
                zlControl.ControlSetFocus txtEdit(mTextIndex.txt_������)
                zlControl.TxtSelAll txtEdit(mTextIndex.txt_������)
                Exit Function
            End If
        End If
    End If
    
    If txtEdit(mTextIndex.txt_������).Text <> txtEdit(mTextIndex.txt_��֤����).Text Then
        ShowMsgbox "������������벻һ�£����������룡"
        zlControl.ControlSetFocus txtEdit(mTextIndex.txt_������)
        zlControl.TxtSelAll txtEdit(mTextIndex.txt_������)
        Exit Function
    End If
    
    isValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    If isValied = False Then Exit Sub
    
    'Zl_���ѿ�����_Update
    strSQL = "Zl_���ѿ�����_Update("
    '  ���ѿ�id_In   In ���ѿ���Ϣ.Id%Type,
    strSQL = strSQL & "" & mlngCardID & ","
    '  ����_In       In ���ѿ���Ϣ.����%Type,
    strSQL = strSQL & "'" & mstrOldPassWord & "',"
    '  �޸�����_In   In ���ѿ���Ϣ.����%Type,
    strSQL = strSQL & "'" & zlCommFun.zlStringEncode(txtEdit(mTextIndex.txt_������).Text) & "',"
    '  ����Ա����_In In ���ѿ��䶯��¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ǿ���޸�_In   In Number := 0 --�Ƿ�ǿ���޸�����,0-��,1-��
    strSQL = strSQL & "" & IIf(mblnCheckOldPass, 0, 1) & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    MsgBox "�����޸ĳɹ���", vbOKOnly + vbInformation, gstrSysName
    mblnOK = True
    Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngCardID = 0
    mstrOldPassWord = ""
End Sub

Private Sub ClearFace()
    txtEdit(mTextIndex.txt_����).PasswordChar = IIf(mTy_CardType.byt������� <> 0, "*", "")
    txtEdit(mTextIndex.txt_����).Text = ""
    txtEdit(mTextIndex.txt_������).Text = "": txtEdit(mTextIndex.txt_��֤����).Text = ""
End Sub

Private Function CheckCardNo(ByVal lngCardTypeID As Long, ByVal strCardNo As String) As Boolean
    '��鿨�ŵĺϷ���
    '���:lngCardTypeID-���ѿ��ӿڱ��
    '     strCardNO-����
    '����:�ɹ�����True,ʧ�ܷ���False
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSQL = _
        "Select a.Id, a.������, a.����, a.���, a.�ɷ��ֵ, a.�ӿڱ��, a.����, a.�������," & vbNewLine & _
        "       To_Char(a.��Ч��, 'yyyy-mm-dd hh24:mi:ss') As ��Ч��," & vbNewLine & _
        "       To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��," & vbNewLine & _
        "       To_Char(a.ͣ������, 'yyyy-mm-dd hh24:mi:ss') As ͣ������," & vbNewLine & _
        "       Decode(a.��ǰ״̬, 2, '����', 3, '�˿�', '����') As ��ǰ״̬" & vbNewLine & _
        "From ���ѿ���Ϣ A" & vbNewLine & _
        "Where a.���� = [1] And a.�ӿڱ�� = [2]" & vbNewLine & _
        "      And ��� = (Select Max(���) From ���ѿ���Ϣ B Where ���� = a.���� And �ӿڱ�� = a.�ӿڱ��)" & vbNewLine & _
        "Order By a.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCardNo, mlngCardTypeID)
    If rsTemp.EOF Then
        ShowMsgbox "δ�ҵ���ص�" & mTy_CardType.str���� & "��Ϣ�����飡"
        Exit Function
    End If
    mlngCardID = Val(Nvl(rsTemp!id))
    mstrOldPassWord = Nvl(rsTemp!����)
    
    '��鵱ǰˢ���ĺϷ���
    '�Ƿ����
    If Nvl(rsTemp!����ʱ��, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "����Ϊ" & strCardNo & "��" & mTy_CardType.str���� & "�Ѿ���" & Nvl(rsTemp!��ǰ״̬) & "��������ˢ����"
        Exit Function
    End If
    
    '�Ƿ�ͣ��
    If Nvl(rsTemp!ͣ������, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "����Ϊ" & strCardNo & "��" & mTy_CardType.str���� & "�Ѿ���ֹͣʹ�ã�������ˢ����"
        Exit Function
    End If
    
    CheckCardNo = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtEdit_Change(Index As Integer)
    Dim blnEnabled As Boolean
    
    On Error GoTo ErrHandler
    Select Case Index
    Case mTextIndex.txt_����
        blnEnabled = Trim(txtEdit(mTextIndex.txt_����).Text) <> ""
        If mblnCheckOldPass = True Then txtEdit(mTextIndex.txt_ԭ����).Enabled = blnEnabled
        txtEdit(mTextIndex.txt_������).Enabled = blnEnabled
        txtEdit(mTextIndex.txt_��֤����).Enabled = blnEnabled
            
        txtEdit(mTextIndex.txt_������).Text = ""
        txtEdit(mTextIndex.txt_��֤����).Text = ""
        txtEdit(mTextIndex.txt_ԭ����).Text = ""
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    On Error GoTo ErrHandler
    Select Case Index
    Case mTextIndex.txt_����
        txtEdit(mTextIndex.txt_����).PasswordChar = IIf(mTy_CardType.byt������� <> 0, "*", "")
    Case mTextIndex.txt_ԭ����, mTextIndex.txt_������, mTextIndex.txt_��֤����
        OpenPassKeyboard txtEdit(Index), True
    End Select
    zlControl.TxtSelAll txtEdit(Index)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean
    
    On Error GoTo ErrHandler
    Select Case Index
    Case mTextIndex.txt_����
        '�Ƿ�ˢ�����
        blnCard = KeyAscii <> 8 And Len(txtEdit(mTextIndex.txt_����).Text) = mTy_CardType.lng���ų��� - 1 _
            And txtEdit(mTextIndex.txt_����).SelLength <> Len(txtEdit(mTextIndex.txt_����).Text)
        If blnCard Or KeyAscii = 13 Then
            If KeyAscii <> 13 Then
                txtEdit(mTextIndex.txt_����).Text = txtEdit(mTextIndex.txt_����).Text & Chr(KeyAscii)
                txtEdit(mTextIndex.txt_����).SelStart = Len(txtEdit(mTextIndex.txt_����).Text)
            End If
            KeyAscii = 0
    
            If CheckCardNo(mlngCardTypeID, Trim(txtEdit(mTextIndex.txt_����).Text)) = False Then
                If txtEdit(mTextIndex.txt_����).Enabled Then txtEdit(mTextIndex.txt_����).SetFocus
                zlControl.TxtSelAll txtEdit(mTextIndex.txt_����)
                Exit Sub
            End If
            If mblnCheckOldPass Then
                zlControl.ControlSetFocus txtEdit(mTextIndex.txt_ԭ����)
            Else
                zlControl.ControlSetFocus txtEdit(mTextIndex.txt_������): Exit Sub
            End If
        Else
            If InStr(":��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then
                KeyAscii = 0 'ȥ��������ţ����Ҳ�����ճ��
            Else
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End If
            
            If mblnTest Then Exit Sub
            '��ȫˢ�����
            If KeyAscii <> 0 And KeyAscii > 32 Then
                sngNow = timer
                If txtEdit(mTextIndex.txt_����).Text = "" Then
                    sngBegin = sngNow
                ElseIf Format((sngNow - sngBegin) / (Len(txtEdit(mTextIndex.txt_����).Text) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                    txtEdit(mTextIndex.txt_����).Text = Chr(KeyAscii)
                    txtEdit(mTextIndex.txt_����).SelStart = 1
                    KeyAscii = 0
                    sngBegin = sngNow
                End If
            End If
        End If
    Case mTextIndex.txt_ԭ����
        If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
    Case mTextIndex.txt_������
        Call CheckInputPassWord(KeyAscii, mTy_CardType.byt������� = 1)
        If KeyAscii = 13 Then
            KeyAscii = 0
            If txtEdit(mTextIndex.txt_������).Text = "" And txtEdit(mTextIndex.txt_��֤����).Text = "" Then
                zlControl.ControlSetFocus cmdOK
            Else
                zlControl.ControlSetFocus txtEdit(mTextIndex.txt_��֤����)
            End If
        End If
    Case mTextIndex.txt_��֤����
        Call CheckInputPassWord(KeyAscii, mTy_CardType.byt������� = 1)
        If KeyAscii = 13 Then
            KeyAscii = 0: Call cmdOK_Click
        End If
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    On Error GoTo ErrHandler
    Select Case Index
    Case mTextIndex.txt_ԭ����, mTextIndex.txt_������, mTextIndex.txt_��֤����
        OpenPassKeyboard txtEdit(Index), False
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
