VERSION 5.00
Begin VB.Form frmEPRModelEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ʾ���༭"
   ClientHeight    =   5055
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5100
   Icon            =   "frmEPRModelEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cbo 
      Height          =   300
      Index           =   0
      Left            =   1155
      TabIndex        =   25
      Top             =   2655
      Width           =   3660
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1155
      MaxLength       =   10
      TabIndex        =   9
      Top             =   2295
      Width           =   3660
   End
   Begin VB.TextBox txt˵�� 
      Height          =   660
      Left            =   1155
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3030
      Width           =   3660
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -435
      TabIndex        =   21
      Top             =   4560
      Width           =   5760
   End
   Begin VB.Frame fraLine 
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   0
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   5910
      Begin VB.OptionButton opt���� 
         BackColor       =   &H00FDD6C6&
         Caption         =   "���ʽ����(&T): ���ַ����ñ��ʽ�����༭���༭"
         Height          =   225
         Index           =   2
         Left            =   435
         TabIndex        =   26
         Top             =   1200
         Width           =   4455
      End
      Begin VB.OptionButton opt���� 
         BackColor       =   &H00FDD6C6&
         Caption         =   "Ƭ��(&S): "
         Height          =   180
         Index           =   1
         Left            =   435
         TabIndex        =   3
         Top             =   765
         Width           =   1020
      End
      Begin VB.OptionButton opt���� 
         BackColor       =   &H00FDD6C6&
         Caption         =   "����(&M):"
         Height          =   180
         Index           =   0
         Left            =   435
         TabIndex        =   2
         Top             =   345
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ʾ������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   1
         Top             =   30
         Width           =   720
      End
      Begin VB.Image imgNote 
         Height          =   240
         Left            =   150
         Picture         =   "frmEPRModelEdit.frx":058A
         Top             =   15
         Width           =   240
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ļ�����������ݵ�ʾ��, �����༭ʱ�ɵ���ѡ����Ƭ��."
         Height          =   360
         Index           =   1
         Left            =   1485
         TabIndex        =   23
         Top             =   750
         Width           =   3420
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���������ļ���ʽ�����ݵ�ʾ���ĵ�, �����༭ʱѡ��һ�����Ĳ������Ǵ�ǰ����;"
         Height          =   360
         Index           =   0
         Left            =   1485
         TabIndex        =   22
         Top             =   330
         Width           =   3420
         WordWrap        =   -1  'True
      End
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   1155
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4110
      Width           =   2370
   End
   Begin VB.OptionButton opt��Χ 
      Caption         =   "&3)����ʹ��"
      Height          =   180
      Index           =   2
      Left            =   3675
      TabIndex        =   15
      Top             =   3780
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton opt��Χ 
      Caption         =   "&2)����ͨ��"
      Height          =   180
      Index           =   1
      Left            =   2385
      TabIndex        =   14
      Top             =   3780
      Width           =   1215
   End
   Begin VB.OptionButton opt��Χ 
      Caption         =   "&1)ȫԺͨ��"
      Height          =   180
      Index           =   0
      Left            =   1140
      TabIndex        =   13
      Top             =   3780
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2265
      TabIndex        =   19
      Top             =   4665
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3645
      TabIndex        =   20
      Top             =   4665
      Width           =   1215
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1155
      TabIndex        =   7
      Top             =   1935
      Width           =   3660
   End
   Begin VB.TextBox txt��� 
      Height          =   300
      Left            =   1155
      TabIndex        =   5
      Top             =   1575
      Width           =   3660
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   24
      Top             =   2700
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   8
      Top             =   2355
      Width           =   630
   End
   Begin VB.Label lbl˵�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "˵��(&M)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   10
      Top             =   3075
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&R)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   16
      Top             =   4170
      Width           =   630
   End
   Begin VB.Label lbl��Ա 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   3600
      TabIndex        =   18
      Top             =   4110
      Width           =   1230
   End
   Begin VB.Label lbl��Χ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ʹ��(&U)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   12
      Top             =   3780
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   6
      Top             =   1980
      Width           =   630
   End
   Begin VB.Label lbl��� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���(&B)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   4
      Top             =   1635
      Width           =   630
   End
End
Attribute VB_Name = "frmEPRModelEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'˵����
'   1���ϼ�����ͨ��������ShowMe�������������塢�༭����ID,�༭״̬����Ϣ���ݽ��뱾����
'   2���༭״̬����Me.tag��ţ��ֱ�Ϊ"����"��"�޸�"�����ϼ�����ͨ��ShowMe����
'---------------------------------------------------
Private mlngFileId As Long       '���ID
Private mlngRecID As Long       '��¼ID
Private mblnOK As Boolean        '�Ƿ���ɱ༭�˳�
 
Public Function ShowMe(ByVal frmParent As Object, ByVal blnAdd As Boolean, ByVal bytPower As Byte, ByVal lngFileId As Long _
                    , Optional ByVal lngRecId As Long, Optional ByVal EditType As Byte) As Long
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '������bytPower-����Ȩ�ޣ�=0��ȫԺ��=1�����ң�=2�����ˣ���lngFileId-���ID��lngRecID-��¼ID;EditType=0�Զ��岡�� EditType��1ϵͳ�Դ� EditType=2 ���ʽ����
    '���أ�ȷ�������������޸ĵ�ID��ȡ������0
    '---------------------------------------------------
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long
    mlngFileId = lngFileId: mlngRecID = lngRecId
    If blnAdd Then
        Me.Tag = "����": mlngRecID = 0
    Else
        Me.Tag = "�޸�"
    End If
    
    '����������Ϣ
    '------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select Distinct D.ID, D.����, D.����, R.ȱʡ, R.��Աid, P.����" & vbNewLine & _
            "From ���ű� D, ������Ա R, ��Ա�� P, �ϻ���Ա�� U, ��������˵�� C," & vbNewLine & _
            "     (Select ����, ͨ�� From �����ļ��б� Where ID = [1]) L" & vbNewLine & _
            "Where D.ID = R.����id And R.��Աid = P.ID And P.ID = U.��Աid And U.�û��� = User And D.ID = C.����id And" & vbNewLine & _
            "      C.�������� In ('�ٴ�', '���', '����', '����', '����', '����', 'Ӫ��', '���') And" & vbNewLine & _
            "      (Nvl(L.ͨ��, 0) <> 2 Or L.���� = 7 Or" & vbNewLine & _
            "      L.���� <> 7 And L.ͨ�� = 2 And D.ID In (Select ����id From ����Ӧ�ÿ��� Where �ļ�id = [1]))" & vbNewLine & _
            "Order By R.ȱʡ Desc, D.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileId)
    With rsTemp
        Do While Not .EOF
            Me.cbo����.AddItem !���� & "-" & !����
            Me.cbo����.ItemData(Me.cbo����.NewIndex) = !ID
            If !ȱʡ = 1 Then Me.cbo����.ListIndex = Me.cbo����.NewIndex
            Me.lbl��Ա.Tag = !��ԱID: Me.lbl��Ա.Caption = !����
            .MoveNext
        Loop
        If Me.cbo����.ListCount = 0 Then
            MsgBox "��Ŀǰ�����ڸò���Ӧ�ÿ��ҷ�Χ�����ܹ����ģ�", vbExclamation, gstrSysName
            ShowMe = 0: Unload Me: Exit Function
        ElseIf Me.cbo����.ListIndex = -1 Then
            Me.cbo����.ListIndex = 0
        End If
    End With
    
    cbo(0).Clear
    cbo(0).AddItem ""
    gstrSQL = "Select Distinct a.���� From ��������Ŀ¼ a Where a.�ļ�id =[1] And a.���� Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileId)
    If rsTemp.BOF = False Then
        Do While Not rsTemp.EOF
            cbo(0).AddItem rsTemp("����").Value
            rsTemp.MoveNext
        Loop
    End If
    cbo(0).ListIndex = 0
    
    If blnAdd Then
        If EditType = 2 Then
            opt����(2).Value = True: fraLine(0).Enabled = False: opt����(0).Enabled = False: opt����(1).Enabled = False
        Else
            opt����(2).Enabled = False
        End If
    End If
    '����������ȡ
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select l.���, l.����, l.����, l.����, l.����, l.˵��, l.ͨ�ü�, l.����id, d.����, d.���� As ����, l.��Աid, p.���� As ��Ա" & _
            " From ��������Ŀ¼ l, ���ű� d, ��Ա�� p" & _
            " Where l.����id = d.Id And l.��Աid = p.Id And l.id =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngRecID)
    With rsTemp
        If .RecordCount > 0 Then
            opt����(NVL(!����, 0)).Value = True
            Me.fraLine(0).Enabled = False
            Me.txt���.Text = !���
            Me.txt����.Text = "" & !����
            Me.txt����.Text = "" & !����
            Me.txt˵��.Text = "" & !˵��
            Me.opt��Χ(IIf(IsNull(!ͨ�ü�), 0, !ͨ�ü�)).Value = True
            If !��ԱID <> Me.lbl��Ա.Tag Then
                Me.lbl��Ա.Tag = !��ԱID: Me.lbl��Ա.Caption = !��Ա
                Me.cbo����.Clear
                Me.cbo����.AddItem !���� & "-" & !����
                Me.cbo����.ItemData(Me.cbo����.NewIndex) = !����ID
                Me.cbo����.ListIndex = Me.cbo����.NewIndex
                Me.cbo����.Enabled = False
            Else
                For lngCount = 0 To Me.cbo����.ListCount - 1
                    If Me.cbo����.ItemData(lngCount) = IIf(IsNull(!����ID), 0, !����ID) Then
                        Me.cbo����.ListIndex = lngCount: Exit For
                    End If
                Next
            End If
            cbo(0).Text = zlCommFun.NVL(!����)
        End If
        Me.txt���.MaxLength = .Fields("���").DefinedSize
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt˵��.MaxLength = .Fields("˵��").DefinedSize
    End With
    Select Case bytPower
    Case 2: Me.opt��Χ(0).Enabled = False: Me.opt��Χ(1).Enabled = False
    Case 1: Me.opt��Χ(0).Enabled = False
    End Select
    
    If Me.Tag = "����" Then
        Me.txt���.Text = GetMax("��������Ŀ¼", "���", 5, " Where �ļ�id=" & mlngFileId)
    End If
    
    '��ʾ����
    Me.Show vbModal, frmParent
    If mblnOK Then
        ShowMe = mlngRecID
    Else
        ShowMe = 0
    End If
    Unload Me
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = 0
End Function

Private Sub cbo_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub cbo_LostFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(cbo(Index).Text, 50)
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()

    If Trim(Me.txt���.Text) = "" Then MsgBox "�������ţ�", vbInformation, gstrSysName: Me.txt���.SetFocus: Exit Sub
    
    If Trim(Me.txt����.Text) = "" Then MsgBox "���������ƣ�", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "���Ƴ��������" & Me.txt����.MaxLength & "���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    End If
    
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "���볬�������" & Me.txt����.MaxLength & "���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt˵��.Text), vbFromUnicode)) > Me.txt˵��.MaxLength Then
        MsgBox "˵�����������" & Me.txt˵��.MaxLength & "���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName: Me.txt˵��.SetFocus: Exit Sub
    End If
    
    If Me.cbo����.ListIndex = -1 Then MsgBox "��������ң�", vbInformation, gstrSysName: Me.cbo����.SetFocus: Exit Sub
    
    '���ݱ���
    If Me.Tag = "����" Then
        mlngRecID = zlDatabase.GetNextId("��������Ŀ¼")
        gstrSQL = mlngRecID & "," & mlngFileId & ",'" & Trim(Me.txt���.Text) & "','" & Trim(Me.txt����.Text) & "','" & Trim(Me.txt����.Text) & "'"
        gstrSQL = gstrSQL & "," & IIf(Me.opt����(0).Value, 0, IIf(opt����(1).Value, 1, 2)) & ",'" & Replace(Trim(Me.txt˵��.Text), Chr(vbKeyReturn), "") & "'"
        If Me.opt��Χ(0).Value Then
            gstrSQL = gstrSQL & ",0"
        ElseIf Me.opt��Χ(1).Value Then
            gstrSQL = gstrSQL & ",1"
        Else
            gstrSQL = gstrSQL & ",2"
        End If
        gstrSQL = gstrSQL & "," & Me.cbo����.ItemData(Me.cbo����.ListIndex) & "," & Me.lbl��Ա.Tag & ",'" & cbo(0).Text & "'"
        gstrSQL = "Zl_��������Ŀ¼_Insert(" & gstrSQL & ")"
    Else
        gstrSQL = mlngRecID & ",'" & Trim(Me.txt���.Text) & "','" & Trim(Me.txt����.Text) & "','" & Trim(Me.txt����.Text) & "'"
        gstrSQL = gstrSQL & ",'" & Replace(Trim(Me.txt˵��.Text), Chr(vbKeyReturn), "") & "'"
        If Me.opt��Χ(0).Value Then
            gstrSQL = gstrSQL & ",0"
        ElseIf Me.opt��Χ(1).Value Then
            gstrSQL = gstrSQL & ",1"
        Else
            gstrSQL = gstrSQL & ",2"
        End If
        gstrSQL = gstrSQL & "," & Me.cbo����.ItemData(Me.cbo����.ListIndex) & ",'" & cbo(0).Text & "'"
        gstrSQL = "Zl_��������Ŀ¼_Update(" & gstrSQL & ")"
    End If
    
    Err = 0: On Error GoTo errHand
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    mblnOK = True: Me.Hide
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optEditType_Click(Index As Integer)
    
End Sub

Private Sub opt��Χ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub opt����_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub opt����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt���_Change()
'    txt��� = Val(txt���)
End Sub

Private Sub txt���_GotFocus()
    Me.txt���.SelStart = 0: Me.txt���.SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt���_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt���.Text, txt���.MaxLength)
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 4000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt����.Text, txt����.MaxLength)
End Sub

Private Sub txt����_Change()
    ValidControlText txt����
    Me.txt����.Text = Left(zlCommFun.SpellCode(Me.txt����.Text), 10)
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt����.Text, txt����.MaxLength)
End Sub

Private Sub txt˵��_Change()
    ValidControlText txt˵��
End Sub

Private Sub txt˵��_GotFocus()
    Me.txt˵��.SelStart = 0: Me.txt˵��.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0:  Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("'%", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt˵��_LostFocus()
    Me.txt˵��.Text = Replace(Me.txt˵��, Chr(vbKeyReturn), "")
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt˵��_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt˵��.Text, txt˵��.MaxLength)
End Sub
