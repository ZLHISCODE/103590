VERSION 5.00
Begin VB.Form frmѡ�����ַ��� 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   Picture         =   "frmѡ�����ַ���.frx":0000
   ScaleHeight     =   1920
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cmbSelFA 
      Height          =   300
      Left            =   225
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   810
      Width           =   3075
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "����(&O)"
      Default         =   -1  'True
      Height          =   345
      Left            =   225
      TabIndex        =   1
      Top             =   1305
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   345
      Left            =   2085
      TabIndex        =   0
      Top             =   1305
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "��ѡ����ķ������ƣ�"
      Height          =   240
      Left            =   225
      TabIndex        =   4
      Top             =   450
      Width           =   2400
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "���뷽��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   1320
   End
End
Attribute VB_Name = "frmѡ�����ַ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ID_From          As Long     'ѡ�е�Դ����ID
Private ID()            As Long     '��ѡ���ID���У���CmbBox��Ӧ

'==============================================================================
'=���ܣ�ȡ���˳�
'==============================================================================
Private Sub CmdCancel_Click()
    On Error GoTo ErrH
    Unload Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�ȷ��ѡ�з���
'==============================================================================
Private Sub CmdOK_Click()
    On Error GoTo ErrH
    If cmbSelFA.ListIndex = -1 Then MsgBox "��ѡ��һ�����������룡", vbOKOnly + vbInformation, gstrSysName: Exit Sub
    ID_From = ID(cmbSelFA.ListIndex + 1)
    Unload Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ���䷽��ѡ��������
'=���������뽫�����ID��
'==============================================================================
Public Sub FillCmbSelFA(ID_to As Long)
    Dim rsTemp      As ADODB.Recordset
    Dim i           As Long
    
    On Error GoTo ErrH
    
    cmbSelFA.Clear
    'ע����ø�ʽ���ȸ�ֵgstrSQL,Ȼ������ݼ�
    gstrSQL = "select ID,����,ѡ��,����ʱ�� from �������ַ��� where ����='סԺ' and ID <> [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, ID_to)
    rsTemp.Sort = "ѡ�� desc,���� ,����ʱ��"
    
    i = 0
    Do Until rsTemp.EOF
        i = i + 1
        ReDim Preserve ID(1 To i) As Long
        cmbSelFA.AddItem rsTemp("����"), i - 1
        ID(i) = rsTemp("ID")
        rsTemp.MoveNext
    Loop
    If i >= 1 Then cmbSelFA.ListIndex = 0
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

