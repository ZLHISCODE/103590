VERSION 5.00
Begin VB.Form frmIdentify���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   2140
      TabIndex        =   5
      Top             =   1785
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   400
      Left            =   1035
      TabIndex        =   4
      Top             =   1785
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Height          =   1560
      Left            =   165
      TabIndex        =   0
      Top             =   105
      Width           =   3075
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2085
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1065
         Width           =   795
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   990
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1065
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   990
         TabIndex        =   3
         Top             =   660
         Width           =   1890
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         Left            =   990
         TabIndex        =   1
         Top             =   270
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   180
         Index           =   3
         Left            =   1635
         TabIndex        =   9
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   540
         TabIndex        =   8
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IC����"
         Height          =   180
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   345
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   735
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long, mstrIdentify As String, mbytType As Byte

Public Function Identify(ByVal bytType As Byte, Optional lng����ID As Long) As String
    mbytType = bytType
    mlng����ID = lng����ID
    Me.Show 1
    Identify = mstrIdentify
End Function

Private Sub cmdCancel_Click()
    mstrIdentify = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim strIdentify As String
    Dim strAddition As String
    Dim lngSequence As String
    On Error GoTo errHandle
    
    '�ж��Ƿ�����ҽ����������
    If Trim(Text1.Text) = "" Then
        If Text1.Enabled = True Then
            MsgBox "������ҽ����������", vbInformation, gstrSysName
            Text1.SetFocus
        ElseIf Trim(txtNO.Text) = "" Then
            MsgBox "������ҽ�����˿���", vbInformation, gstrSysName
            txtNO.SetFocus
        End If
        Exit Sub
    End If
    If Not IsNumeric(Trim(Text2.Text)) And Trim(Text2.Text) <> "" Then
        MsgBox "�����������", vbInformation, gstrSysName
        Text2.SetFocus
        Exit Sub
    End If
    lngSequence = UCase(txtNO.Text)
'      strInfo='0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
'      8����;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(1,2,3);15����֤��;16�����;17�Ҷȼ�
'      18�ʻ������ۼ�;19�ʻ�֧���ۼ�;20����ͳ���ۼ�;21ͳ�ﱨ���ۼ�;22סԺ�����ۼ�;23�������
'      24��������;25�����ۼ�;26����ͳ���޶�
    
    strIdentify = lngSequence & ";"                             '0����
    strIdentify = strIdentify & lngSequence & ";"               '1ҽ���ţ����˱�ţ�
    strIdentify = strIdentify & ";"                             '2����
    strIdentify = strIdentify & Trim(Text1.Text) & ";"          '3����
    strIdentify = strIdentify & Combo1.Text & ";"               '4�Ա�
    strIdentify = strIdentify & ";"                             '5��������
    strIdentify = strIdentify & ";"                             '6���֤
    strIdentify = strIdentify & ";"                             '7.��λ����(����)
    strAddition = "0;"                                          '8.���Ĵ���
    strAddition = strAddition & ";"                             '9.˳���
    strAddition = strAddition & ";"                             '10��Ա���
    '��Ϊ����ȡ�ø��������Խ����������㹻���ֵ������֧�������ҽ��ȷ��
    strAddition = strAddition & "1000000;"                      '11�ʻ����
    strAddition = strAddition & "0;"                            '12��ǰ״̬
    strAddition = strAddition & ";"                             '13����ID
    strAddition = strAddition & "1;"                            '14��ְ(1,2,3)
    strAddition = strAddition & ";"                             '15����֤��
    strAddition = strAddition & Text2.Text & ";"                '16�����
    strAddition = strAddition & ";"                             '17�Ҷȼ�
    strAddition = strAddition & "1000000;"                      '18�ʻ������ۼ�
    strAddition = strAddition & ";"                             '19�ʻ�֧���ۼ�
    strAddition = strAddition & "0;"                            '20����ͳ���ۼ�
    strAddition = strAddition & "0;"                            '21ͳ�ﱨ���ۼ�
    strAddition = strAddition & "0;"                            '22סԺ�����ۼ�
    strAddition = strAddition & ";"                             '23��������
    
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_����)
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrIdentify = strIdentify & mlng����ID & ";" & strAddition
    End If
    Me.Hide
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error Resume Next
        cmdOK.Enabled = True
        cmdOK.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Combo1.AddItem "��"
    Combo1.AddItem "Ů"
    Combo1.ListIndex = 0
End Sub

Private Sub Text1_GotFocus()
    On Error Resume Next
    Call zlControl.TxtSelAll(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error Resume Next
        Text2.Enabled = True
        Text2.SetFocus
    End If
End Sub

Private Sub Text2_GotFocus()
    On Error Resume Next
    Call zlControl.TxtSelAll(Text2)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error Resume Next
        Combo1.Enabled = True
        Combo1.SetFocus
    End If
End Sub

Private Sub txtNO_GotFocus()
    On Error Resume Next
    Call zlControl.TxtSelAll(txtNO)
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    Dim StrInput As String, strSQL As String, rsTemp As New ADODB.Recordset
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txtNO <> "" Then
        txtNO.Text = UCase(txtNO.Text)
        gstrSQL = "Select * From �����ʻ� Where ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CStr(txtNO.Text))
        If rsTemp.EOF Then
            Text1.Enabled = True
            Text2.Enabled = True
            Combo1.Enabled = True
            Text1.SetFocus
        Else
            gstrSQL = "Select * From ������Ϣ Where ����id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rsTemp!����ID))
            If rsTemp.EOF Then
                Text1.Enabled = True
                Text2.Enabled = True
                Combo1.Enabled = True
                Text1.SetFocus
            Else
                Text1.Enabled = False
                Text2.Enabled = False
                Combo1.Enabled = False
                Text1.Text = rsTemp!����
                Text2.Text = Nvl(rsTemp!����)
                Combo1.ListIndex = IIf(Nvl(rsTemp!�Ա�, "��") = "��", 0, 1)
                cmdOK.SetFocus
            End If
        End If
    End If
End Sub
