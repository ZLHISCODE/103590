VERSION 5.00
Begin VB.Form FrmSet���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   Icon            =   "FrmSet����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox Txt���������� 
      Height          =   300
      Left            =   1110
      MaxLength       =   4
      TabIndex        =   3
      Top             =   600
      Width           =   3045
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3090
      TabIndex        =   5
      Top             =   1080
      Width           =   1100
   End
   Begin VB.ComboBox Cbo����ģʽ 
      Height          =   300
      Left            =   1110
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   3045
   End
   Begin VB.Label Lbl���������� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   180
      Left            =   300
      TabIndex        =   2
      Top             =   660
      Width           =   720
   End
   Begin VB.Label Lbl����ģʽ 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ģʽ"
      Height          =   180
      Left            =   300
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "FrmSet����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng���� As Long
Private blnOK As Boolean

Public Function ShowSet(ByVal lng���� As Long) As Boolean
    blnOK = False
    mlng���� = lng����
    
    Me.Show 1
    ShowSet = blnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    
    gcnOracle.BeginTrans
    gcnOracle.Execute "zl_���ղ���_Delete(" & mlng���� & ",NULL)", , adCmdStoredProc
    gcnOracle.Execute "zl_���ղ���_Insert(" & mlng���� & ",NULL,'����ģʽ'," & Cbo����ģʽ.ListIndex & ",1)", , adCmdStoredProc
    gcnOracle.Execute "zl_���ղ���_Insert(" & mlng���� & ",NULL,'����������','" & Txt����������.Text & "',2)", , adCmdStoredProc
    gcnOracle.CommitTrans
    
    blnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim intValue As Integer
    
    'װ���ʼ����
    Cbo����ģʽ.Clear
    Cbo����ģʽ.AddItem "�Ȱ����Ժ����,�ٰ����Ժ����"
    Cbo����ģʽ.AddItem "�Ȱ����Ժ����,�ٰ����Ժ����"
    
    '��ȡ����ֵ
    intValue = 0
    gstrSQL = "Select Nvl(����ֵ,0) Value From ���ղ��� Where ����=[1] And ������='����ģʽ'"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ֵ", mlng����)
    
    If Not rsTmp.EOF Then
        intValue = rsTmp!Value
    End If
    Cbo����ģʽ.ListIndex = intValue
    
    '����������
    gstrSQL = "Select Nvl(����ֵ,'') Value From ���ղ��� Where ����=[1] And ������='����������'"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ֵ", mlng����)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!Value) Then
            Txt����������.Text = rsTmp!Value
        End If
    End If
End Sub


