VERSION 5.00
Begin VB.Form frmSet�¶����� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "frmSet�¶�����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox Txt������ 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1050
      TabIndex        =   9
      Text            =   "0"
      Top             =   1080
      Width           =   915
   End
   Begin VB.TextBox txtIC�˿ں� 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3818
      TabIndex        =   7
      Text            =   "1"
      Top             =   360
      Width           =   255
   End
   Begin VB.ComboBox cbo������ 
      Height          =   300
      Left            =   1043
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   2055
   End
   Begin VB.ComboBox cbo���õ��� 
      Height          =   300
      ItemData        =   "frmSet�¶�����.frx":000C
      Left            =   1043
      List            =   "frmSet�¶�����.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1530
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   90
      TabIndex        =   2
      Top             =   2025
      Width           =   4275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2255
      TabIndex        =   4
      Top             =   2235
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   965
      TabIndex        =   3
      Top             =   2235
      Width           =   1100
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "IC����Ҫָ�������ı��:"
      Height          =   180
      Left            =   180
      TabIndex        =   10
      Top             =   780
      Width           =   2070
   End
   Begin VB.Label lblIC�˿ں� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�˿ں�"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3188
      TabIndex        =   8
      Top             =   420
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   6
      Top             =   420
      Width           =   540
   End
   Begin VB.Label lbl���õ��� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���õ���"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   0
      Top             =   1590
      Width           =   720
   End
End
Attribute VB_Name = "frmSet�¶�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnOK As Boolean

Public Function ShowSet() As Boolean
    blnOK = False
    
    Me.Show 1
    ShowSet = blnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cbo������_Click()
    Me.lblIC�˿ں�.Enabled = (cbo������.ListIndex <> 0)
    Me.txtIC�˿ں�.Enabled = (cbo������.ListIndex <> 0)
    Me.Txt������.Enabled = (cbo������.ListIndex <> 0)
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    
    gcnOracle.BeginTrans
    gcnOracle.Execute "zl_���ղ���_Delete(" & TYPE_�¶� & ",NULL)", , adCmdStoredProc
    gcnOracle.Execute "zl_���ղ���_Insert(" & TYPE_�¶� & ",NULL,'���õ���'," & Me.cbo���õ���.ListIndex & ",1)", , adCmdStoredProc
    gcnOracle.Execute "zl_���ղ���_Insert(" & gintInsure & ",NULL,'������'," & Me.Txt������.Text & ",2)", , adCmdStoredProc
    gcnOracle.CommitTrans
    
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName, "������", Me.cbo������.ListIndex)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName, "IC�豸�˿�", txtIC�˿ں�.Text)
    
    mint���õ���_�¶� = Me.cbo���õ���.ListIndex
    mintIC�������� = Me.Txt������.Text
    
    blnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    
    '���ӳ�ʼ������
    Me.cbo���õ���.Clear
    Me.cbo���õ���.AddItem "�¶�����׽�"
    '������(2005-12-28):�������õ���������
    Me.cbo���õ���.AddItem "������"
    Me.cbo���õ���.AddItem "�ѽ�"
    Me.cbo���õ���.ListIndex = 0
    
    '����ǰ�Ĳ���ȡ������ʾ�ڽ�����
    gstrSQL = "Select ������,Nvl(����ֵ,0) Value From ���ղ��� Where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ϴ���Ժ��Ϣ����ֵ", gintInsure)
    If Not rsTmp.EOF Then
       Select Case rsTmp!������
          Case "���õ���"
             Me.cbo���õ���.ListIndex = Nvl(rsTmp!Value, 0)
          Case "������"
             Me.Txt������.Text = Nvl(rsTmp!Value, 0)
       End Select
    End If

    
    Me.cbo������.Clear
    Me.cbo������.AddItem "�ſ�"
    Me.cbo������.AddItem "IC��-JKP428"
    Me.cbo������.AddItem "IC��-ICIOX"
    Me.cbo������.ListIndex = 0
    
    cbo������.ListIndex = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "������", 0)
    txtIC�˿ں�.Text = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "IC�豸�˿�", 1)

End Sub
