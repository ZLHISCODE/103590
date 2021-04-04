VERSION 5.00
Begin VB.Form frmTaskSendFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   1785
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5190
   Icon            =   "frmTaskSendParam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   60
      TabIndex        =   7
      Top             =   90
      Width           =   3720
      Begin VB.CheckBox chk 
         Caption         =   "�����ѷ���"
         Height          =   270
         Left            =   1215
         TabIndex        =   4
         Top             =   1125
         Width           =   1260
      End
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   0
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   660
         Width           =   2400
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   255
         Width           =   2400
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��첿��(&D)"
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   0
         Top             =   300
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ʱ��(&U)"
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   735
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3900
      TabIndex        =   5
      Top             =   180
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3900
      TabIndex        =   6
      Top             =   600
      Width           =   1100
   End
End
Attribute VB_Name = "frmTaskSendFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnStartUp As Boolean
Private mblnOK As Boolean

Public Function ShowFilter(ByVal frmMain As Object) As Boolean
    
    mblnOK = False
    
    If InitActivate = False Then Exit Function
    If LoadData = False Then Exit Function
        
    Me.Show 1, frmMain
    
    ShowFilter = mblnOK
    
End Function

Private Function InitActivate() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ�����ݣ������ڴ����Activate�¼�
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    gstrSQL = "SELECT A.����||'-'||A.����,ID FROM ���ű� A,��������˵�� B WHERE (A.����ʱ�� IS NULL OR A.����ʱ�� =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.����ID AND B.��������='���' ORDER BY A.����||'-'||A.����"
    
    Call OpenRecordset(rs, Me.Caption)
    If rs.BOF Then
        ShowSimpleMsg "û��������ʵĲ��ţ����ڲ��Ź��������ã�"
        Exit Function
    End If
    
    '�����ݵ��ؼ���
    Call AddComboData(cboDept, rs)
    
    '��ʼѡ�����ݴ���
    CboLocate cboDept, UserInfo.����ID, True
    If cboDept.ListCount > 0 And cboDept.ListIndex = -1 Then cboDept.ListIndex = 0
    
    cbo(0).AddItem "��  ��"
    cbo(0).AddItem "��  ��"
    cbo(0).AddItem "��  ��"
    cbo(0).AddItem "��  ��"
    cbo(0).AddItem "��  ��"
    cbo(0).AddItem "������"
    cbo(0).AddItem "��  ��"
    cbo(0).AddItem "ǰ����"
    cbo(0).AddItem "ǰһ��"
    cbo(0).AddItem "ǰ����"
    cbo(0).AddItem "ǰһ��"
    cbo(0).AddItem "ǰ����"
    cbo(0).AddItem "ǰ����"
    cbo(0).AddItem "ǰ����"
    cbo(0).AddItem "ǰһ��"
            
    
    InitActivate = True
    
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then
        'Resume
    'End If
End Function

Private Function LoadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  װ������
    '����:
    '------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHand
    
    CboLocate cboDept, Val(GetSetting("ZLSOFT", "����ȫ��\�ɱ��ӿ�", "��첿��", "0")), True
    
    chk.Value = Val(GetSetting("ZLSOFT", "����ȫ��\�ɱ��ӿ�", "�����ѷ���", "0"))
    
    On Error Resume Next
    cbo(0).Text = GetSetting("ZLSOFT", "����ȫ��\�ɱ��ӿ�", "���ʱ��", "��  ��")
    If cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
        
    LoadData = True
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then
        'Resume
    'End If
End Function

Private Function SaveData() As Boolean

    On Error GoTo errHand
    
    
    If cboDept.ListIndex = -1 Then
        Call SaveSetting("ZLSOFT", "����ȫ��\�ɱ��ӿ�", "��첿��", "0")
    Else
        Call SaveSetting("ZLSOFT", "����ȫ��\�ɱ��ӿ�", "��첿��", cboDept.ItemData(cboDept.ListIndex))
    End If
    
    Call SaveSetting("ZLSOFT", "����ȫ��\�ɱ��ӿ�", "�����ѷ���", chk.Value)
    Call SaveSetting("ZLSOFT", "����ȫ��\�ɱ��ӿ�", "���ʱ��", cbo(0).Text)
    
    SaveData = True
    
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then
        'Resume
    'End If
    
End Function


Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        PressKey vbKeyTab
    End If
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        PressKey vbKeyTab
    End If
End Sub

Private Sub chk_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveData Then
        mblnOK = True
        Unload Me
    End If
End Sub




