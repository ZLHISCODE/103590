VERSION 5.00
Begin VB.Form frm����ѡ��_�Ͻ� 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ѡ��"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "frm����ѡ��_�Ͻ�.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -90
      TabIndex        =   3
      Top             =   1290
      Width           =   4965
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2910
      TabIndex        =   5
      Top             =   1500
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1620
      TabIndex        =   4
      Top             =   1500
      Width           =   1100
   End
   Begin VB.CommandButton cmd���� 
      Appearance      =   0  'Flat
      Caption         =   "��"
      Height          =   270
      Left            =   3750
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   810
      Width           =   285
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1230
      TabIndex        =   1
      Top             =   810
      Width           =   2835
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   195
      Index           =   1
      Left            =   540
      TabIndex        =   7
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   195
      Index           =   0
      Left            =   540
      TabIndex        =   6
      Top             =   210
      Width           =   3495
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   510
      TabIndex        =   0
      Top             =   870
      Width           =   630
   End
End
Attribute VB_Name = "frm����ѡ��_�Ͻ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mlng����ID As Long
Private mlng����ID As Long
Private mstr�������� As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mlng����ID = Val(txt����.Tag)
    If mlng����ID = 0 Then
        MsgBox "����Ҫѡ��һ�����֣�", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    
    If InStr(1, txt����.Text, ")") <> 0 Then
        mstr�������� = Mid(txt����.Text, InStr(1, txt����.Text, ")") + 1)
    End If
    
    mblnOK = True
    Unload Me
End Sub

Public Function ����ѡ��(ByVal lng����ID As String, lng����ID As Long, str�������� As String) As Boolean
    mlng����ID = lng����ID
    mlng����ID = 0
    mstr�������� = ""
    mblnOK = False
    Me.Show 1
    If mblnOK Then
        lng����ID = mlng����ID
        str�������� = mstr��������
    End If
    ����ѡ�� = mblnOK
End Function

Private Sub Form_Load()
    Dim lng����ID As Long
    Dim str�������� As String
    Dim STR���� As String
    Dim str��ᱣ�Ϻ� As String
    Dim rsTemp As New ADODB.Recordset
    '����ǰѡ��Ĳ�����ʾ����
    gstrSQL = "Select ����ID,ҽ���� From �����ʻ� Where ����=[1] And ����ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰѡ��Ĳ���ID", TYPE_�Ͻ�, mlng����ID)
    lng����ID = Nvl(rsTemp!����ID, 0)
    str��ᱣ�Ϻ� = Nvl(rsTemp!ҽ����)
    
    '��ȡ�ò��ֵ���Ϣ
    gstrSQL = "Select ���ִ���,�������� From ����Ŀ¼�� Where ID=" & lng����ID
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        If .RecordCount <> 0 Then
            Me.txt����.Text = "(" & rsTemp!���ִ��� & ")" & rsTemp!��������
            Me.txt����.Tag = lng����ID
        End If
    End With
    
    'ȡ������Ϣ
    gstrSQL = "Select ���� From ������Ϣ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", mlng����ID)
    STR���� = rsTemp!����
    
    Me.lblNote(0).Caption = "����������" & STR����
    Me.lblNote(1).Caption = "��ᱣ�Ϻţ�" & str��ᱣ�Ϻ�
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt����.Text = "" And txt����.Tag <> "" Then Exit Sub
    
    On Error GoTo errHandle
    
    strText = UCase(txt����.Text)
    If InStr(1, strText, "(") <> 0 Then
        If InStr(1, strText, ")") <> 0 Then
            strText = Mid(strText, 2, InStr(1, strText, ")") - 2)
        End If
    End If
    gstrSQL = " Select ID,���ִ��� As ����,��������,��ҽ����,�������,�����Ը�����,�����𸶽�� " & _
              " From ����Ŀ¼�� A" & _
              " Where (" & zlCommFun.GetLike("A", "���ִ���", strText) & " or " & zlCommFun.GetLike("A", "��������", strText) & " or zlspellcode(��������) Like '" & strText & "%')"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
    End With
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "�����ڸò��֣����������룡", vbInformation, gstrSysName
        txt����.Text = lbl����.Tag
        zlControl.TxtSelAll txt����
        Exit Sub
    Else
        '����ѡ����
        If rsTemp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(TYPE_�Ͻ�, rsTemp, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        txt����.Text = lbl����.Tag
        zlControl.TxtSelAll txt����
        Exit Sub
    Else
        '�϶����м�¼����
        txt����.Tag = rsTemp!ID
        txt����.Text = "(" & rsTemp!���� & ")" & rsTemp!��������
        lbl����.Tag = txt����.Text '���ڻָ���ʾ
    End If
    
    Call zlCommFun.PressKey(vbKeyTab)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmd����_Click()
    Dim blnReturn As Boolean
    Dim rsTemp As New ADODB.Recordset
        
    gstrSQL = " Select ID,���ִ��� As ����,��������,��ҽ����,�������,�����Ը�����,�����𸶽�� " & _
              " From ����Ŀ¼��"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
    End With
    
    blnReturn = frmListSel.ShowSelect(TYPE_�Ͻ�, rsTemp, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�")
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        txt����.Text = lbl����.Tag
        zlControl.TxtSelAll txt����
        Exit Sub
    Else
        '�϶����м�¼����
        txt����.Tag = rsTemp!ID
        txt����.Text = "(" & rsTemp!���� & ")" & rsTemp!��������
        lbl����.Tag = txt����.Text '���ڻָ���ʾ
    End If
End Sub
