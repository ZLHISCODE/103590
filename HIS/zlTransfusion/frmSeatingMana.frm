VERSION 5.00
Begin VB.Form frmSeatingMana 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��λ����"
   ClientHeight    =   3885
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   5085
   Icon            =   "frmSeatingMana.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleMode       =   0  'User
   ScaleWidth      =   5085
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkAgainAdd 
      Caption         =   "��������(&A)"
      Height          =   270
      Left            =   450
      TabIndex        =   20
      Top             =   3375
      Width           =   1665
   End
   Begin VB.Frame fraOne 
      Height          =   3165
      Left            =   150
      TabIndex        =   2
      Top             =   45
      Width           =   4755
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1005
         MaxLength       =   30
         TabIndex        =   8
         Top             =   540
         Width           =   1500
      End
      Begin VB.TextBox txt������ 
         Height          =   300
         Left            =   1005
         MaxLength       =   30
         TabIndex        =   5
         Top             =   2670
         Width           =   3510
      End
      Begin VB.OptionButton Opt���� 
         Caption         =   "��λ"
         Height          =   240
         Index           =   1
         Left            =   3810
         TabIndex        =   12
         Top             =   915
         Width           =   675
      End
      Begin VB.OptionButton Opt���� 
         Caption         =   "��λ"
         Height          =   240
         Index           =   0
         Left            =   3105
         TabIndex        =   11
         Top             =   915
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.CommandButton cmdPopu 
         Caption         =   "��"
         Height          =   240
         Left            =   4200
         TabIndex        =   7
         Top             =   1230
         Width           =   270
      End
      Begin VB.TextBox txt�շ���Ŀ 
         Height          =   300
         Left            =   1005
         MaxLength       =   103
         TabIndex        =   13
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txt��� 
         Height          =   300
         Left            =   2985
         MaxLength       =   30
         TabIndex        =   9
         Top             =   540
         Width           =   1500
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   870
         Width           =   1500
      End
      Begin VB.TextBox txt��ע 
         Height          =   705
         Left            =   1005
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1890
         Width           =   3495
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1005
         TabIndex        =   6
         Top             =   210
         Width           =   3495
      End
      Begin VB.TextBox txt�շѱ�׼ 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   1005
         TabIndex        =   3
         Top             =   1545
         Width           =   3495
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lbl��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   2700
         Width           =   720
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "����"
         Height          =   180
         Left            =   2445
         TabIndex        =   21
         Top             =   945
         Width           =   495
      End
      Begin VB.Label lbl��� 
         Alignment       =   1  'Right Justify
         Caption         =   "���"
         Height          =   180
         Index           =   0
         Left            =   2445
         TabIndex        =   19
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbl�ȼ� 
         Alignment       =   1  'Right Justify
         Caption         =   "״̬"
         Height          =   180
         Left            =   465
         TabIndex        =   18
         Top             =   915
         Width           =   495
      End
      Begin VB.Label lbl�շ���Ŀ 
         Alignment       =   1  'Right Justify
         Caption         =   "�շ���Ŀ"
         Height          =   180
         Left            =   165
         TabIndex        =   17
         Top             =   1212
         Width           =   795
      End
      Begin VB.Label lbl��ע 
         Alignment       =   1  'Right Justify
         Caption         =   "��ע"
         Height          =   180
         Left            =   165
         TabIndex        =   16
         Top             =   1890
         Width           =   795
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Caption         =   "����"
         Height          =   180
         Left            =   465
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "�շѱ�׼"
         Height          =   180
         Left            =   165
         TabIndex        =   14
         Top             =   1536
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdCance 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3525
      TabIndex        =   1
      Top             =   3345
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2265
      TabIndex        =   0
      Top             =   3345
      Width           =   1100
   End
End
Attribute VB_Name = "frmSeatingMana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintType As Integer '����ģʽ 0-���� 1-�޸�
Private mSeating As Seating
Private mstr�������� As String
Private mblnOk As Boolean '�޸ģ������Ƿ�ɹ�
Private mSelectTxt As String 'ѡ����շ���Ŀ�ı�
Private mblnShow As Boolean '�Ƿ�����ʾ����������������λʱ
Private mSeatings As Seatings
Private mfrmMain As frmDockSeat

Public Function SeatingMana(ByVal intType As Integer, ByVal curSeatings As Seatings, ByVal int��� As Integer, ByVal StrKey As String, ByVal frmParent As Form, Optional strType As String) As Boolean
    'intType: 0-���� 1-�޸�
    'curSeatings : ��λ��¼��
    'int��� : Ҫ���ӻ��޸ĵ���λ����
    'strKey : ������޸ķ�ʽ������Ҫ�޸ĵ���λ�ı��,���ӷ�ʽ�ɴ��մ�
    '
    mblnOk = False
    Set mSeatings = curSeatings
    mintType = intType
    Set mSeating = New Seating
    
    If intType = 0 Then
        mSeating.��� = mSeatings.GetNextNo(int���)
        mSeating.���� = strType
        If mSeating.���� = "" Then mSeating.���� = "��ͨ��λ"
        
    Else
        Set mSeating = mSeatings.Item(StrKey)
    End If
    mSeating.��� = int���
    mstr�������� = mSeatings.��������
    
    If (intType = 0 And Not mblnShow) Or intType = 1 Then
        Set mfrmMain = frmParent
        frmSeatingMana.Show vbModal, frmParent
    Else
        Call initForm
    End If

    SeatingMana = True
    
End Function

Private Sub cmdCance_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    
    mSeating.�շ���Ŀ = txt�շ���Ŀ
    mSeating.��ע = txt��ע

    mSeating.�շ�ϸĿID = txt�շ���Ŀ.Tag
    mSeating.״̬ = cboType.ItemData(cboType.ListIndex)
    mSeating.���� = IIf(Opt����(0).Value = True, 0, 1)
    mSeating.��������� = "" & txt������
    
    mblnOk = True
    If mintType = 0 Then
        mSeating.��� = txt���
        mSeating.���� = txt����
        If mSeating.���� = "" Then mSeating.���� = "��ͨ��λ"
        If mSeating.���� <> "��ͨ��λ" Then
            mSeating.��� = 1
        Else
            mSeating.��� = 0
        End If
        With mSeating
        Call mSeatings.Add(0, 0, 0, "", "", .���, .���, .״̬, _
                         IIf(IsNull(.�ּ�), 0, .�ּ�), IIf(IsNull(.�շ�ϸĿID), 0, .�շ�ϸĿID), "", IIf(IsNull(.��ע), "", .��ע), .����, .����, .���������, .��� & "_" & .���)
        End With
        
        Call SeatingMana(mintType, mSeatings, mSeating.���, "", Me, mSeating.����)
        If chkAgainAdd.Value = 0 Then
            Unload Me
        Else
            'Call mfrmMain.RefreshMain
        End If
    Else
        Dim strReturn As String
        
        If mSeating.���� <> "��ͨ��λ" Then
            mSeating.��� = 1
        Else
            mSeating.��� = 0
        End If
        
        With mSeating
        strReturn = .Update(mSeatings.����ID, .�շ�ϸĿID, .״̬, .�շ���Ŀ, .�ּ�, .��ע, .����, .���������)
        If strReturn <> "" Then
            MsgBox "��������ʱ���ִ���" & strReturn, vbInformation, gstrSysName
        End If
        End With
        Unload Me
    End If
End Sub

Private Sub cmdPopu_Click()
    Call ShowSelectWindow(0)
End Sub

Private Sub Form_Load()
    'WAIT:��λ�������
    Call initForm
    mblnShow = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmMain = Nothing
    mblnShow = False
End Sub


Private Sub txt�շ���Ŀ_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    If (txt�շ���Ŀ <> mSelectTxt) Or (txt�շ���Ŀ.Tag > 0) And (InStr(txt�շ���Ŀ, "]") - InStr(txt�շ���Ŀ, "[") < 1) Then
        If Trim(txt�շ���Ŀ) <> "" Then
            Call ShowSelectWindow(1)
        Else
            txt�շ���Ŀ.Tag = 0
            zlCommFun.PressKey (vbKeyTab)
        End If
    Else
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub ShowSelectWindow(ByVal intLoadType As Integer)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim blnCanel As Boolean, vRect As RECT, strTXT As String
    On Error GoTo hErr
    strTXT = UCase(Trim(txt�շ���Ŀ.Text))
    If intLoadType = 0 Then
        strSQL = "Select A.ID, A.����, A.����, A.���㵥λ, B.�ּ�, A.��������, Decode(A.�������, 1, '����', '�����סԺ') As �������," & vbNewLine & _
                "       A.ִ�п���" & vbNewLine & _
                "From (Select �ּ�, �շ�ϸĿid,�۸�ȼ� From �շѼ�Ŀ Where ��ֹ���� Is Null Or ��ֹ���� = To_Date('3000-01-01', 'YYYY-MM-DD')) B," & vbNewLine & _
                "     �շ���ĿĿ¼ A" & vbNewLine & _
                "Where A.ID = B.�շ�ϸĿid And Mod(A.�������, 2) = 1 And" & vbNewLine & _
                "      (A.վ��='" & zl9ComLib.gstrNodeNo & "' Or A.վ�� is Null) And " & vbNewLine & _
                "      (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(A.�Ƿ���, 0) = 0 And" & vbNewLine & _
                "      A.��� = 'J'" & GetPriceGradeSQL(gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "1", "2", "3")
    Else
        If InStr(txt�շ���Ŀ.Text, "]") - InStr(txt�շ���Ŀ.Text, "[") > 1 Then
            strTXT = Mid(txt�շ���Ŀ, InStr(strTXT, "[") + 1, InStr(strTXT, "]") - 2)
        End If
        
        strSQL = "Select A.ID, A.����, A.����, A.���㵥λ, B.�ּ�, A.��������, Decode(A.�������, 1, '����', '�����סԺ') As �������," & vbNewLine & _
                "       A.ִ�п���" & vbNewLine & _
                "From �շ���Ŀ���� C," & vbNewLine & _
                "     (Select �ּ�, �շ�ϸĿid,�۸�ȼ� From �շѼ�Ŀ Where ��ֹ���� Is Null Or ��ֹ���� = To_Date('3000-01-01', 'YYYY-MM-DD')) B," & vbNewLine & _
                "     �շ���ĿĿ¼ A" & vbNewLine & _
                "Where A.ID = C.�շ�ϸĿid And A.ID = B.�շ�ϸĿid And Mod(A.�������, 2) = 1 And" & vbNewLine & _
                "      (A.վ��='" & zl9ComLib.gstrNodeNo & "' Or A.վ�� is Null) And " & vbNewLine & _
                "      (C.���� Like '%" & strTXT & "%' Or A.���� Like '%" & strTXT & "%' Or A.���� Like '%" & strTXT & "%') And C.���� = 1 And" & vbNewLine & _
                "      (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(A.�Ƿ���, 0) = 0 And" & vbNewLine & _
                "      A.��� = 'J'" & GetPriceGradeSQL(gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ��Ŀ�۸�ȼ�, "A", "B", "1", "2", "3")

    End If
    
    vRect = ZLControl.GetControlRect(txt�շ���Ŀ.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ָ���շ���Ŀ", False, "", "ѡ���շ���Ŀ", False, False, True, _
                                         vRect.Left, vRect.Top, txt�շ���Ŀ.Height, blnCanel, True, True, gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ��Ŀ�۸�ȼ�)
                                         
    If Not blnCanel And Not rsTmp Is Nothing Then
        txt�շ���Ŀ = Replace("[" & zlCommFun.NVL(rsTmp.Fields("����")) & "] " & zlCommFun.NVL(rsTmp.Fields("����")), "[]", "")
        txt�շ���Ŀ.Tag = zlCommFun.NVL(rsTmp.Fields("ID"), 0)
        txt�շѱ�׼ = Format(zlCommFun.NVL(rsTmp.Fields("�ּ�"), 0), "0.00")
        mSelectTxt = txt�շ���Ŀ
        zlCommFun.PressKey (vbKeyTab)
    Else
        txt�շ���Ŀ = ""
        mSelectTxt = txt�շ���Ŀ
        txt�շ���Ŀ.Tag = 0
        txt�շѱ�׼ = "0.00"
        txt�շ���Ŀ.SetFocus
    End If
    Exit Sub
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub initForm()
    Dim str��� As String
    txt��� = mSeating.���
    txt���� = mSeating.����
    txtDept = mstr��������
    
    cboType.Clear
    
    cboType.AddItem "0-�ڱ�", 0
    cboType.ItemData(0) = 0
    cboType.AddItem "2-����", 1
    cboType.ItemData(1) = 2
    
'    Select Case mSeating.���
'    Case 0
'        str��� = "��ͨ��λ"
'    Case 1
'        str��� = "����"
'    Case 2
'        str��� = "����ҩƷ��λ"
'    Case 3
'        str��� = "VIP��λ"
'    End Select
    
    If mintType = 0 Then
        Me.Caption = "��λ���� - ����"
        txt��ע = ""
        If chkAgainAdd.Value = 0 Then
            '��������,������շ���Ŀ
            txt�շѱ�׼ = "0.00"
            txt�շ���Ŀ = ""
            txt�շ���Ŀ.Tag = 0
            
            Opt����(0).Value = True: Opt����(1).Value = False
        End If
        cboType.ListIndex = 0
        txt���.Enabled = True
        txt����.Enabled = True
        chkAgainAdd.Enabled = True
        
    Else
        txt���.Enabled = False
        Me.Caption = "��λ���� - �޸�"
        txt��ע = mSeating.��ע
        txt�շѱ�׼ = Format(mSeating.�ּ�, "0.00")
        txt�շ���Ŀ.Tag = mSeating.�շ�ϸĿID
        txt�շ���Ŀ = mSeating.�շ���Ŀ
        cboType.ListIndex = IIf(mSeating.״̬ = 0, 0, 1)
        
        If mSeating.���� = 0 Then
            Opt����(0).Value = True: Opt����(1).Value = False
        Else
            Opt����(0).Value = False: Opt����(1).Value = True
        End If
        txt������ = mSeating.���������
        txt���� = mSeating.����: txt����.Enabled = False
        If txt���� = "" Then txt���� = "��ͨ��λ"
        chkAgainAdd.Enabled = False
    End If
End Sub
