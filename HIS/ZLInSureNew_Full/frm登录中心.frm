VERSION 5.00
Begin VB.Form frm��¼���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��¼����"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   Icon            =   "frm��¼����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdModify 
      Caption         =   "�޸�����(M)"
      Height          =   375
      Left            =   90
      TabIndex        =   11
      Top             =   1650
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   26
      Left            =   1020
      TabIndex        =   13
      Top             =   270
      Visible         =   0   'False
      Width           =   2865
      Begin VB.TextBox TXT���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   26
         Left            =   1080
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   525
         Width           =   1620
      End
      Begin VB.TextBox txt�û� 
         Height          =   300
         Index           =   26
         Left            =   1080
         TabIndex        =   5
         Top             =   90
         Width           =   1620
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&P)"
         Height          =   180
         Index           =   26
         Left            =   330
         TabIndex        =   6
         Top             =   585
         Width           =   630
      End
      Begin VB.Label Lbl�û��� 
         AutoSize        =   -1  'True
         Caption         =   "����(&U)"
         Height          =   180
         Index           =   26
         Left            =   330
         TabIndex        =   4
         Top             =   150
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -240
      TabIndex        =   8
      Top             =   1440
      Width           =   4785
   End
   Begin VB.CommandButton CDMȷ�� 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1710
      TabIndex        =   9
      Top             =   1665
      Width           =   1100
   End
   Begin VB.CommandButton CMD���� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2910
      TabIndex        =   10
      Top             =   1665
      Width           =   1100
   End
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   43
      Left            =   1020
      TabIndex        =   12
      Top             =   270
      Visible         =   0   'False
      Width           =   2865
      Begin VB.TextBox txt�û� 
         Height          =   300
         Index           =   43
         Left            =   1080
         TabIndex        =   1
         Top             =   90
         Width           =   1620
      End
      Begin VB.TextBox TXT���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   43
         Left            =   1080
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   525
         Width           =   1620
      End
      Begin VB.Label Lbl�û��� 
         AutoSize        =   -1  'True
         Caption         =   "����(&U)"
         Height          =   180
         Index           =   43
         Left            =   330
         TabIndex        =   0
         Top             =   150
         Width           =   630
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&P)"
         Height          =   180
         Index           =   43
         Left            =   330
         TabIndex        =   2
         Top             =   585
         Width           =   630
      End
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   180
      Picture         =   "frm��¼����.frx":000C
      Top             =   300
      Width           =   720
   End
End
Attribute VB_Name = "frm��¼����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint���� As Integer
Private mbln�޸����� As Boolean
Private mblnLogin As Boolean
Private mstr������ As String

Public Function LoginCenter(ByVal int���� As Integer, Optional ByVal bln�޸����� As Boolean = False) As Boolean
    '��¼����
    On Error Resume Next
    mblnLogin = False
    mint���� = int����
    mbln�޸����� = bln�޸�����
    
    Me.Show 1
    LoginCenter = mblnLogin
End Function

Private Sub CDMȷ��_Click()
    If Trim(txt�û�(mint����).Text) = "" Then
        MsgBox "���������Ա���ţ�", vbInformation, gstrSysName
        txt�û�(mint����).SetFocus
        Exit Sub
    End If
    
    Call Login
End Sub

Private Sub cmdModify_Click()
    With frm�޸�����
        mstr������ = .ChangePassword(txt����(mint����).Text)
    End With
End Sub

Private Sub CMD����_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    fra(mint����).Visible = True
    cmdModify.Visible = mbln�޸�����
    
    '�Զ���¼
    Call AutoLogin
End Sub

Private Sub AutoLogin()
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    Select Case mint����
    Case TYPE_������, TYPE_��ɽ
        If mint���� = TYPE_������ Then
            gstrSQL = "Select ���,����ֵ From ���ղ��� Where ����=[1] And ��� In (4,5) Order by ���"
        Else
            gstrSQL = "Select ���,����ֵ From ���ղ��� Where ����=[2] And ��� In (2,3) Order by ���"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��¼����", mint����)
        If rsTemp.EOF Then Exit Sub
        
        Do While Not rsTemp.EOF
            If rsTemp.AbsolutePosition = 1 Then
                txt�û�(mint����) = Nvl(rsTemp!����ֵ)
            Else
                txt����(mint����) = Nvl(rsTemp!����ֵ)
            End If
            rsTemp.MoveNext
        Loop
    End Select
    
    If Trim(txt�û�(mint����)) <> "" And Trim(txt����(mint����)) <> "" Then Call Login
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Login()
    On Error GoTo errHand
    
    Select Case mint����
    Case TYPE_������
        If Not ���ýӿ�_׼��_������(Function_������.��¼����) Then Exit Sub
        '��д��ڲ���
        Call CZ_DataPut(glngInterface_������, 1, "staff_id", Trim(txt�û�(mint����).Text))
        Call CZ_DataPut(glngInterface_������, 1, "staff_pwd", txt����(mint����).Text)
        '����
        If Not ���ýӿ�_ִ��_������ Then Exit Sub
        '��¼�ɹ����������Ա����
        gCominfo_������.����Ա���� = Trim(txt�û�(mint����).Text)
    Case TYPE_��ɽ
        Dim strUserName As TStringOfChar, strPassWord As TStringOfChar
        strUserName.Data = Trim(txt�û�(mint����).Text)
        strPassWord.Data = txt����(mint����).Text
        gbytReturn_��ɽ = LS_UserLogin(strUserName, strPassWord)
        If GetErrInfo_��ɽ Then Exit Sub
    End Select
    
    If mstr������ <> "" Then
        '�޸�����
        Select Case mint����
        Case TYPE_��ɽ
            Dim strNewPwd As TStringOfChar
            strNewPwd.Data = mstr������
            gbytReturn_��ɽ = LS_ChangePwd(strPassWord, strNewPwd)
            If gbytReturn_��ɽ <> 0 Then MsgBox "�����޸�ʧ�ܣ�", vbInformation, gstrSysName
        End Select
    End If
    
    mblnLogin = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub
