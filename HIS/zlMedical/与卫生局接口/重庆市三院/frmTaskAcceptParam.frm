VERSION 5.00
Begin VB.Form frmTaskAcceptParam 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ز���"
   ClientHeight    =   2595
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4830
   Icon            =   "frmTaskAcceptParam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3525
      TabIndex        =   8
      Top             =   2010
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2370
      TabIndex        =   7
      Top             =   2010
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   4560
      Begin VB.PictureBox picCmd 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4110
         ScaleHeight     =   255
         ScaleWidth      =   300
         TabIndex        =   5
         Top             =   615
         Width           =   300
         Begin VB.CommandButton cmd 
            Caption         =   "��"
            Height          =   240
            Left            =   15
            TabIndex        =   6
            Top             =   15
            Width           =   270
         End
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   1185
         TabIndex        =   2
         Top             =   585
         Width           =   3240
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   1185
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   3240
      End
      Begin VB.Label Label1 
         Caption         =   "����ĺ�Լ��λָ�ķ���������Ļ������������֣����û����Ϣ�����ں�Լ��λ�н�����"
         Height          =   570
         Left            =   1170
         TabIndex        =   9
         Top             =   1020
         Width           =   3345
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Լ��λ(&U)"
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   4
         Top             =   660
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��첿��(&D)"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   3
         Top             =   255
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmTaskAcceptParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnStartUp As Boolean

Private Type Items
    �������� As String
    ID As Long
End Type

Private usrSaveGroup As Items

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
    
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    CboLocate cboDept, Val(GetSetting("ZLSOFT", "����ȫ��\�ɱ��ӿ�", "��첿��", "0")), True

    cmd.Tag = Val(GetSetting("ZLSOFT", "����ȫ��\�ɱ��ӿ�", "��Լ��λ", "0"))
    
    gstrSQL = "Select ���� From ��Լ��λ Where ID=" & Val(cmd.Tag)
    Call OpenRecordset(rs, Me.Caption)
    If rs.BOF = False Then
        txt.Text = NVL(rs("����"))
    Else
        cmd.Tag = ""
    End If
    
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
    
    Call SaveSetting("ZLSOFT", "����ȫ��\�ɱ��ӿ�", "��Լ��λ", cmd.Tag)
    
    SaveData = True
    
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then
        'Resume
    'End If
    
End Function

Private Sub cmd_Click()
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "SELECT -1 AS ID,NULL+0 AS �ϼ�id,'0' AS ����,'����' AS ����,'' as ����,'' as ��ַ,0 AS ĩ��,'' AS ��ϵ��,'' AS �绰,'' AS �����ʼ�,'' AS ��������,'' AS �ʺ�,'' AS ��ַ,'' AS ˵�� from dual " & _
                        "Union All " & _
                        "SELECT ID,DECODE(�ϼ�id,NULL,-1,0,-1,�ϼ�id) AS �ϼ�id,����,����,����,��ַ,0 AS ĩ��,��ϵ��,�绰,�����ʼ�,��������,�ʺ�,��ַ,˵�� from ��Լ��λ  where ĩ��<>1 " & _
                        "Start With �ϼ�id is null connect by prior ID=�ϼ�id " & _
                        "Union All " & _
                        "SELECT ID,DECODE(�ϼ�id,NULL,-1,0,-1,�ϼ�id) AS �ϼ�id,����,����,����,��ַ,1 AS ĩ��,��ϵ��,�绰,�����ʼ�,��������,�ʺ�,��ַ,˵�� from ��Լ��λ  where ĩ��=1"
                        
    If ShowTxtSelectDialog(Me, txt, "����,900,0,1;����,1500,0,1;����,900,0,1;��ַ,3000,0,1", Me.Name & "\�������ѡ��", "�����±���ѡ��һ������/��λ��", strSQL, rs, 8790, 5100) Then
        
        txt.Text = NVL(rs("����").Value)
        cmd.Tag = NVL(rs("ID").Value, 0)
        
        usrSaveGroup.�������� = txt.Text

    End If
    txt.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveData Then Unload Me
End Sub

Private Sub Form_Activate()

    If mblnStartUp = False Then Exit Sub
    DoEvents
        
    If InitActivate = False Then
        mblnStartUp = False
        Unload Me
        Exit Sub
    End If
    
    mblnStartUp = False
    
    Call LoadData
    
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True

End Sub

Private Sub txt_Change()

    txt.Tag = "Changed"
    cmd.Tag = ""

End Sub

Private Sub txt_GotFocus()

    TxtSelAll txt

End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    Dim rs As New ADODB.Recordset
    Dim strFilter As String
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If txt.Tag = "Changed" Then
            
            strFilter = "'%" & UCase(txt.Text) & "%'"
            
            gstrSQL = "select ID,����,����,����,��ַ,��ϵ��,�绰,�����ʼ�,��������,�ʺ�,��ַ,˵��,ĩ�� from ��Լ��λ  where ĩ��=1 " & _
                " AND (���� Like " & strFilter & " or ���� Like " & strFilter & " OR ���� Like " & strFilter & ")"
                
            If ShowTxtFilterDialog(Me, txt, "����,1800,0,0;����,900,0,0;����,900,0,0;��ϵ��,900,0,0;�绰,1200,0,0", Me.Name & "\�������ѡ��", "�������ѡ��һ�����嵥λ", gstrSQL, rs) Then
                
                txt.Text = NVL(rs("����"))
                cmd.Tag = NVL(rs("ID"))
                txt.Tag = ""
                
                usrSaveGroup.�������� = txt.Text
                
                                
            Else
                txt.Text = usrSaveGroup.��������
                Exit Sub
            End If
        End If

        PressKey vbKeyTab
        PressKey vbKeyTab
        
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt.Locked Then
        glngTXTProc = GetWindowLong(txt.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt.Locked Then
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
    
    If txt.Tag = "Changed" Then
        txt.Text = usrSaveGroup.��������
        txt.Tag = ""
    End If

End Sub

