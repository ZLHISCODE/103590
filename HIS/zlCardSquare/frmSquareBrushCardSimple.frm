VERSION 5.00
Begin VB.Form frmSquareBrushCardSimple 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���㿨ˢ��"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "����"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSquareBrushCardSimple.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   465
      Left            =   5010
      TabIndex        =   12
      Top             =   3330
      Width           =   1335
   End
   Begin VB.CommandButton cmdȡ�� 
      Caption         =   "ȡ��(&C)"
      Height          =   465
      Left            =   6360
      TabIndex        =   14
      Top             =   3330
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "ˢ����Ϣ"
      Height          =   2880
      Left            =   180
      TabIndex        =   13
      Top             =   255
      Width           =   7605
      Begin VB.TextBox txtEdit 
         Height          =   405
         Index           =   3
         Left            =   1695
         MaxLength       =   100
         TabIndex        =   11
         Top             =   2115
         Width           =   5745
      End
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1695
         TabIndex        =   9
         Top             =   1575
         Width           =   2535
      End
      Begin VB.TextBox txtEdit 
         Height          =   405
         Index           =   0
         Left            =   1695
         TabIndex        =   1
         Top             =   495
         Width           =   2550
      End
      Begin VB.TextBox txtEdit 
         Height          =   405
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1695
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1020
         Width           =   2550
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��ע(&S)"
         Height          =   285
         Index           =   4
         Left            =   615
         TabIndex        =   10
         Top             =   2175
         Width           =   1020
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����ˢ��(&X)"
         Height          =   285
         Index           =   3
         Left            =   45
         TabIndex        =   8
         Top             =   1635
         Width           =   1590
      End
      Begin VB.Label lblInfor 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Index           =   1
         Left            =   5445
         TabIndex        =   7
         Top             =   1020
         Width           =   1935
      End
      Begin VB.Label lblInfor 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Index           =   0
         Left            =   5430
         TabIndex        =   3
         Top             =   495
         Width           =   1965
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   285
         Index           =   0
         Left            =   4530
         TabIndex        =   2
         Top             =   555
         Width           =   855
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   285
         Index           =   1
         Left            =   615
         TabIndex        =   0
         Top             =   555
         Width           =   1020
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&W)"
         Height          =   285
         Index           =   2
         Left            =   615
         TabIndex        =   4
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ���"
         Height          =   285
         Index           =   10
         Left            =   4260
         TabIndex        =   6
         Top             =   1080
         Width           =   1140
      End
   End
   Begin VB.Label lblʧЧ�� 
      Height          =   240
      Left            =   210
      TabIndex        =   15
      Top             =   3420
      Width           =   4455
   End
End
Attribute VB_Name = "frmSquareBrushCardSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng�ӿڱ�� As Long, mdbl����ˢ�� As Double, mstrBlanceInfor As String, mintSucces As Integer
Private Type CardInfor
    lng���ѿ�ID As Long
    str���� As String
    dbl��� As Double
    dbl������Ѷ� As Double
    dblʧЧ��� As Double '��ȡ��ԭ����,�Ƚ��ȳ��ķ���:�����ѿ����,��������ֵ��:�˽��Ϊ,���ں�δ���ѵĽ��
    str������� As String
    str�ӿ����� As String
    str���㷽ʽ As String
End Type
Private mTyCurCardInfor As CardInfor

Private Enum mtxtIdx
    idx_txt���� = 0
    idx_txt���� = 1
    idx_txt����ˢ�� = 2
    idx_txt��ע = 3
End Enum
Private Enum mlblIdx
    idx_lbl������ = 0
    idx_lbl��� = 1
    idx_lbl����ˢ�� = 3
    idx_lbl��ע = 4
End Enum
Private WithEvents mobjBrushCard As clsBrushSequareCard
Attribute mobjBrushCard.VB_VarHelpID = -1
Private mblnChange As Boolean
Private mblnCardNoSHowPW As Boolean
Private mobjKeyboard As Object

Private Function CheckDepended() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵĹ�����
    '����:���˺�
    '����:2009-12-24 12:13:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Dim rsTemp As New ADODB.Recordset
    Set rsTemp = zlGet���ѿ��ӿ�
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    rsTemp.Find "���=" & mlng�ӿڱ��, , , 1
    If rsTemp.EOF Then
        ShowMsgbox "�ӿ�δ�ҵ�(���Ϊ" & mlng�ӿڱ�� & "),����!"
        Exit Function
    End If
    With mTyCurCardInfor
        .str�ӿ����� = Nvl(rsTemp!����)
        .str���㷽ʽ = Nvl(rsTemp!���㷽ʽ)
        txtEdit(mtxtIdx.idx_txt����).MaxLength = Len(Nvl(rsTemp!ǰ׺�ı�)) + Val(Nvl(rsTemp!���ų���))
    End With
    CheckDepended = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function zlShowBrushCard(ByVal frmMain As Object, ByVal lng�ӿڱ�� As Long, dbl����ˢ�� As Double, _
    strBlanceInfor As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ˢ���ӿ�
    '��Σ�frmMain-���õ�������
    '       dbl����ˢ��-����ˢ����
    '����:strBlanceInfor-���ؽ�����Ϣ( ��||�ָ�: �ӿڱ��||���ѿ�ID(�ɴ�'')||���㷽ʽ||������||����||������ˮ��||����ʱ��(yyyy-mm-dd hh24:mi:ss)||��ע)
    '����:���óɹ�,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-06-18 15:27:01
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    mdbl����ˢ�� = dbl����ˢ��: mlng�ӿڱ�� = lng�ӿڱ��: mintSucces = 0
    If CheckDepended = False Then Exit Function
    
    txtEdit(mtxtIdx.idx_txt����ˢ��).Text = Format(dbl����ˢ��, "###0.00;-###0.00;0.00;0.00")
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlShowBrushCard = mintSucces > 0
    strBlanceInfor = mstrBlanceInfor
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmdȡ��_Click()
    mintSucces = 0: Unload Me
End Sub
Private Sub cmdȷ��_Click()
    Dim dt����ʱ�� As Date

    '������,��ʾ��Ҫ����Ƿ�Ϸ�
    If CheckInput = False Then Exit Sub
    dt����ʱ�� = zlDatabase.Currentdate
    
    ' �ӿڱ��||���ѿ�ID(�ɴ�'')||���㷽ʽ||������||����||������ˮ��||����ʱ��(yyyy-mm-dd hh24:mi:ss)||��ע
    mstrBlanceInfor = mlng�ӿڱ��
    mstrBlanceInfor = mstrBlanceInfor & "||" & mTyCurCardInfor.lng���ѿ�ID
    mstrBlanceInfor = mstrBlanceInfor & "||" & mTyCurCardInfor.str���㷽ʽ
    mstrBlanceInfor = mstrBlanceInfor & "||" & Val(txtEdit(mtxtIdx.idx_txt����ˢ��).Text)
    mstrBlanceInfor = mstrBlanceInfor & "||" & mTyCurCardInfor.str����
    mstrBlanceInfor = mstrBlanceInfor & "||" & ""
    mstrBlanceInfor = mstrBlanceInfor & "||" & Format(dt����ʱ��, "yyyy-mm-dd HH:MM:SS")
    mstrBlanceInfor = mstrBlanceInfor & "||" & Replace(Trim(txtEdit(mtxtIdx.idx_txt����ˢ��).Text), "|", "")
    mintSucces = mintSucces + 1
    mblnChange = False
    Unload Me
End Sub
Private Function CheckInput() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ�Ϸ�
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-23 17:03:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    If txtEdit(mtxtIdx.idx_txt����).Text <> Trim(txtEdit(mtxtIdx.idx_txt����).Tag) Or Trim(txtEdit(mtxtIdx.idx_txt����).Text) = "" Then
        ShowMsgbox "δˢ����ˢ������ȷ,����!"
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt����)
        Exit Function
    End If
    If CheckInputPassWord = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt����)
        Exit Function
    End If
    If CheckInputSquareMoney = False Then
        zlCtlSetFocus txtEdit(mtxtIdx.idx_txt����ˢ��)
        Exit Function
    End If
    CheckInput = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function CheckInputPassWord() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������Ƿ���ȷ
    '����:��ȷ,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-23 14:27:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Trim(txtEdit(mtxtIdx.idx_txt����).Tag) <> "" And Trim(txtEdit(mtxtIdx.idx_txt����).Text) = "" Then
        ShowMsgbox "����δ����,����!"
        Exit Function
    End If
    
    If Trim(txtEdit(mtxtIdx.idx_txt����).Tag) <> zlCommFun.zlStringEncode(Trim(txtEdit(mtxtIdx.idx_txt����).Text)) Then
        ShowMsgbox "�����������,����!"
        Exit Function
    End If
    CheckInputPassWord = True
End Function

Private Function CheckInputSquareMoney() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ı������ѽ���Ƿ���ȷ
    '����:��ȷ,����true,���򷵻�False
    '����:���˺�
    '����:2009-12-23 14:27:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If zlDblIsValid(Trim(txtEdit(mtxtIdx.idx_txt����ˢ��).Text), 16, True, True, 0, "�����") = False Then
        Exit Function
    End If
    If Val(lblInfor(mlblIdx.idx_lbl���).Caption) < Val(Trim(txtEdit(mtxtIdx.idx_txt����ˢ��).Text)) Then
        ShowMsgbox "������(" & Format(Val(lblInfor(mlblIdx.idx_lbl���).Caption), gVbFmtString.FM_���) & "Ԫ),����!"
        Exit Function
    End If
    CheckInputSquareMoney = True
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Call CreateObjectKeyboard
    '����Ƿ���������ص�ˢ������
    Set mobjBrushCard = New clsBrushSequareCard
    Call mobjBrushCard.zlInitInterFacel(mlng�ӿڱ��)
    mblnCardNoSHowPW = zlIsCardNoShowPW(mlng�ӿڱ��)
    If mblnCardNoSHowPW Then
        txtEdit(mtxtIdx.idx_txt����).PasswordChar = "*"
    Else
        txtEdit(mtxtIdx.idx_txt����).PasswordChar = ""
    End If
    
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index <> mtxtIdx.idx_txt���� Then txtEdit(Index).Tag = ""
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Select Case Index
    Case mtxtIdx.idx_txt����
        zlCommFun.OpenIme False
        gTy_TestBug.BytType = 2
        If Not mobjBrushCard Is Nothing Then Call mobjBrushCard.zlSetAutoBrush(Trim(txtEdit(Index).Text) = "")
    Case mtxtIdx.idx_txt��ע
        zlCommFun.OpenIme True
    Case Else
        zlCommFun.OpenIme False
        If Index = mtxtIdx.idx_txt���� Then
            Call OpenPassKeyboard(txtEdit(Index))
        End If
    End Select
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim str���� As String, str���� As String, lngID As Long
    Dim strCardNo As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    Select Case Index
    Case mtxtIdx.idx_txt����
        If txtEdit(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab
        '���ǿ��ܴ��ڲ���Ա��ˢ�������,����ݲ��������¹���:
        If IsDesinMode = False Then Exit Sub
        
        If txtEdit(Index).Text = "" Then
            'ֱ�ӵ�����
            If mobjBrushCard.zlReadCard(Me, strCardNo) = False Then
                Exit Sub
            End If
            txtEdit(Index).Text = strCardNo
            txtEdit(Index).Tag = strCardNo
        End If
        
        If zlBrusCard(Trim(txtEdit(Index))) = False Then
            zlCtlSetFocus txtEdit(Index)
        Else
            If txtEdit(mtxtIdx.idx_txt����).Tag = "" Then
                If txtEdit(mtxtIdx.idx_txt��ע).Enabled And txtEdit(mtxtIdx.idx_txt��ע).Visible Then txtEdit(mtxtIdx.idx_txt��ע).SetFocus
            Else
                zlCommFun.PressKey vbKeyTab
            End If
        End If
        
    Case mtxtIdx.idx_txt��ע
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    Case mtxtIdx.idx_txt����
        If CheckInputPassWord = False Then
            zlControl.ControlSetFocus txtEdit(Index): Exit Sub
        End If
        zlCommFun.PressKey vbKeyTab
    Case mtxtIdx.idx_txt����ˢ��
        If CheckInputSquareMoney = False Then
            zlControl.ControlSetFocus txtEdit(Index): Exit Sub
        End If
        zlCommFun.PressKey vbKeyTab
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
End Sub
Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim blnCard As Boolean
    
    Select Case Index
    Case mtxtIdx.idx_txt����
        If InStr(1, "'~��|`-'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        If IsDesinMode Then Exit Sub
        Call BrushCard(txtEdit(Index), KeyAscii)
    Case mtxtIdx.idx_txt��ע
        If InStr(1, "'|'", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
        blnCard = zlInputIsCard(txtEdit(Index), KeyAscii, glngSys, mblnCardNoSHowPW)
        If blnCard = True Then KeyAscii = 0
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m�ı�ʽ
    Case mtxtIdx.idx_txt����
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m�ı�ʽ
    Case mtxtIdx.idx_txt����ˢ��
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m���ʽ
    Case Else
    End Select
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = mtxtIdx.idx_txt���� Then
        Call ClosePassKeyboard(txtEdit(Index))
    End If
End Sub

Private Sub txtEdit_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> mtxtIdx.idx_txt���� Then Exit Sub
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtEdit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> mtxtIdx.idx_txt���� Then Exit Sub
    If Button = 2 Then
        Call SetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
    Case mtxtIdx.idx_txt����
    Case mtxtIdx.idx_txt��ע
    Case mtxtIdx.idx_txt����
        If CheckInputPassWord = False Then
        End If
    Case mtxtIdx.idx_txt����ˢ��
        If CheckInputSquareMoney = False Then
           'Cancel = 1
        End If
    Case Else
    End Select
End Sub

Private Function zlBrusCard(ByVal strCardNo As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ˢ������
    '���ƣ����˺�
    '���ڣ�2010-06-18 15:12:22
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim rsTemp As ADODB.Recordset, i As Long, strTemp As String, blnFind As Boolean
    
    With mTyCurCardInfor
        .dblʧЧ��� = 0
        .dbl��� = 0
        .dbl������Ѷ� = 0
        .str���� = ""
        .lng���ѿ�ID = 0
    End With
    
    gstrSQL = "" & _
    "   Select a.Id,a.������,a.����,a.���,a.�ɷ��ֵ,to_char(a.��Ч��,'yyyy-mm-dd hh24:mi:ss') as ��Ч��,  a.����," & _
    "          to_char(a.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ�� , " & _
    "          decode(a.��ǰ״̬,2,'����',3,'�˿�','����') as ��ǰ״̬, " & _
    "          to_char(a.������," & gOraFmtString.FM_��� & ") as ������ ," & _
    "          to_char(a.���۽��," & gOraFmtString.FM_��� & ") as ���۽�� ," & _
    "          to_char(a.��ֵ�ۿ���," & gOraFmtString.FM_�ۿ��� & ") as ��ֵ�ۿ��� ," & _
    "          to_char(a.���," & gOraFmtString.FM_��� & ") as ��� ," & _
    "          to_char(a.ͣ������,'yyyy-mm-dd hh24:mi:ss') as ͣ������," & _
    "          a.������� " & _
    "   From ���ѿ�Ŀ¼ A  " & _
    "   Where A.���� = [1] and A.�ӿڱ��=[2] And ��� = (Select Max(���) From ���ѿ�Ŀ¼ B Where ���� = A.���� and �ӿڱ��=A.�ӿڱ��)  " & _
    "   Order by a.���"
    Err = 0: On Error GoTo ErrHand:
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strCardNo, mlng�ӿڱ��)
    If rsTemp.EOF Then
       ShowMsgbox "δ�ҵ���ص����ѿ���¼,����!"
        Exit Function
    End If
    '���:
    '�Ƿ����
    If Nvl(rsTemp!����ʱ��, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox IIf(mblnCardNoSHowPW, "", "����Ϊ" & strCardNo & "��") & "���ѿ��Ѿ���" & Nvl(rsTemp!��ǰ״̬) & ",������ˢ��"
        Exit Function
    End If
    '�Ƿ�ͣ��
    If Nvl(rsTemp!ͣ������, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox IIf(mblnCardNoSHowPW, "", "����Ϊ" & strCardNo & "��") & "���ѿ��Ѿ���ֹͣʹ��,������ˢ��"
        Exit Function
    End If
    '�Ƿ�ͣ��
    If Nvl(rsTemp!ͣ������, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox IIf(mblnCardNoSHowPW, "", "����Ϊ" & strCardNo & "��") & "���ѿ��Ѿ���ֹͣʹ��,������ˢ��"
        Exit Function
    End If
    
    '���Ч��
    mTyCurCardInfor.dbl��� = Val(Nvl(rsTemp!���))
    lblʧЧ��.Visible = False
    If Nvl(rsTemp!��Ч��, "3000-01-01 00:00:00") < Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") Then
       '������Ч��
       If Val(Nvl(rsTemp!�ɷ��ֵ)) = 1 Then
          '������ֵ��,���ڵ�,�������ѿ�����,ֻ��������ֵ����
          mTyCurCardInfor.dblʧЧ��� = zlGetʧЧ���(Val(Nvl(rsTemp!ID)), mlng�ӿڱ��)
          mTyCurCardInfor.dbl��� = IIf(mTyCurCardInfor.dbl��� - mTyCurCardInfor.dblʧЧ��� < 0, 0, mTyCurCardInfor.dbl��� - mTyCurCardInfor.dblʧЧ���)
          If mTyCurCardInfor.dblʧЧ��� <> 0 Then
            lblʧЧ��.Caption = "��ǰ����ʧЧ���(�����)Ϊ��" & Format(mTyCurCardInfor.dblʧЧ���, gVbFmtString.FM_���) & "Ԫ"
            lblʧЧ��.Visible = True
            lblʧЧ��.ForeColor = vbRed
          End If
       Else
            '��������ֵ��,�����ٽ�������
            ShowMsgbox IIf(mblnCardNoSHowPW, "", "����Ϊ" & strCardNo & "��") & "���ѿ��Ѿ�ʧЧ,������ˢ��"
            Exit Function
       End If
    End If
    If mTyCurCardInfor.dbl��� <= 0 Then
        ShowMsgbox IIf(mblnCardNoSHowPW, "", "����Ϊ" & strCardNo & "��") & "���ѿ��Ѿ�û�����,������ˢ��"
        Exit Function
    End If
    
    With mTyCurCardInfor
        .lng���ѿ�ID = Val(Nvl(rsTemp!ID))
        .str���� = Nvl(rsTemp!����)
        .str������� = Nvl(rsTemp!�������)
    End With
    txtEdit(mtxtIdx.idx_txt����).Text = Nvl(rsTemp!����)
    txtEdit(mtxtIdx.idx_txt����).Tag = Nvl(rsTemp!����)
    lblInfor(mlblIdx.idx_lbl���).Caption = Format(Val(Nvl(rsTemp!���)), gVbFmtString.FM_���)
    lblInfor(mlblIdx.idx_lbl������).Caption = Nvl(rsTemp!������)
    txtEdit(mtxtIdx.idx_txt����).Tag = Nvl(rsTemp!����)
    'ȱʡֵ:����,ȱʡ���,����Ϊ������Ѷ�
    If mTyCurCardInfor.dbl��� < mdbl����ˢ�� Then
        txtEdit(mtxtIdx.idx_txt����ˢ��).Text = Format(mTyCurCardInfor.dbl���, gVbFmtString.FM_���)
    Else
        txtEdit(mtxtIdx.idx_txt����ˢ��).Text = Format(mdbl����ˢ��, gVbFmtString.FM_���)
    End If
    zlBrusCard = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub ClearCtlData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ؼ�����
    '����:���˺�
    '����:2009-12-24 11:11:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    txtEdit(mtxtIdx.idx_txt����ˢ��) = "0.00"
    txtEdit(mtxtIdx.idx_txt����) = ""
    txtEdit(mtxtIdx.idx_txt����) = ""
    txtEdit(mtxtIdx.idx_txt����).Tag = ""
    lblInfor(mlblIdx.idx_lbl������).Caption = ""
    lblInfor(mlblIdx.idx_lbl���).Caption = "0.00"
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub BrushCard(ByVal objEdit As Object, KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ������(Ŀǰֻ֧���п�����ˢ��)
    '����:���˺�
    '����:2010-02-09 14:07:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean
    
    '�Ƿ�ˢ�����
    blnCard = KeyAscii <> 8 And Len(objEdit.Text) = objEdit.MaxLength - 1 And objEdit.SelLength <> Len(objEdit.Text)
    
    If blnCard Then
        If KeyAscii <> 13 Then
            objEdit.Text = objEdit.Text & Chr(KeyAscii)
            objEdit.SelStart = Len(objEdit.Text)
        End If
        KeyAscii = 0
        'ˢ������:
        If zlBrusCard(Trim(objEdit)) = False Then
            zlCtlSetFocus objEdit
        Else
            If txtEdit(mtxtIdx.idx_txt����).Tag = "" Then
                If txtEdit(mtxtIdx.idx_txt��ע).Enabled And txtEdit(mtxtIdx.idx_txt��ע).Visible Then txtEdit(mtxtIdx.idx_txt��ע).SetFocus
            Else
                zlCommFun.PressKey vbKeyTab
            End If
        End If
    Else
        If InStr(":��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 'ȥ��������ţ����Ҳ�����ճ��
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        '��ȫˢ�����
        If KeyAscii <> 0 And KeyAscii > 32 Then
            sngNow = Timer
            If objEdit.Text = "" Then
                sngBegin = sngNow
            ElseIf Format((sngNow - sngBegin) / (Len(objEdit.Text) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                objEdit.Text = Chr(KeyAscii)
                objEdit.SelStart = 1
                KeyAscii = 0
                sngBegin = sngNow
            End If
        End If
    End If
End Sub
Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������봴��
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
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function OpenPassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText) = False Then Exit Function
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


