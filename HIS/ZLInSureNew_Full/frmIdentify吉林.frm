VERSION 5.00
Begin VB.Form frmIdentify���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������֤"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraInfo 
      Caption         =   "������Ϣ"
      Height          =   4740
      Left            =   90
      TabIndex        =   35
      Top             =   115
      Width           =   6420
      Begin VB.CommandButton Cmd����1 
         Caption         =   "��"
         Height          =   285
         Left            =   5895
         TabIndex        =   41
         Top             =   3960
         Width           =   255
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   16
         Left            =   1290
         TabIndex        =   40
         Top             =   3960
         Width           =   4890
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Left            =   1290
         TabIndex        =   37
         Top             =   4335
         Width           =   4890
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   285
         Left            =   5895
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3600
         Width           =   255
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   15
         Left            =   4515
         TabIndex        =   31
         Top             =   3255
         Width           =   1650
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   3255
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   12
         Left            =   1290
         TabIndex        =   25
         Top             =   2820
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   10
         Left            =   1290
         TabIndex        =   21
         Top             =   2390
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   8
         Left            =   1290
         TabIndex        =   17
         Top             =   1960
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   4515
         TabIndex        =   11
         Top             =   1100
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   4515
         TabIndex        =   7
         Top             =   670
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4515
         TabIndex        =   3
         Top             =   240
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   13
         Left            =   4515
         TabIndex        =   27
         Top             =   2820
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   11
         Left            =   4515
         TabIndex        =   23
         Top             =   2390
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   9
         Left            =   4515
         TabIndex        =   19
         Top             =   1960
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   7
         Left            =   4515
         TabIndex        =   15
         Top             =   1530
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   1290
         TabIndex        =   13
         Top             =   1530
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1290
         TabIndex        =   9
         Top             =   1100
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1290
         TabIndex        =   5
         Top             =   670
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1290
         TabIndex        =   1
         Top             =   240
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   14
         Left            =   1290
         TabIndex        =   42
         Top             =   3600
         Width           =   4890
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "����I"
         Height          =   180
         Index           =   15
         Left            =   810
         TabIndex        =   43
         Top             =   3660
         Width           =   450
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "����II"
         Height          =   180
         Index           =   17
         Left            =   720
         TabIndex        =   39
         Top             =   4020
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Index           =   16
         Left            =   540
         TabIndex        =   38
         Top             =   4395
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "֧�����к�"
         Height          =   180
         Index           =   16
         Left            =   3600
         TabIndex        =   30
         Top             =   3330
         Width           =   900
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   7
         Left            =   540
         TabIndex        =   28
         Top             =   3330
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "�Ƿ��ش󼲲�"
         Height          =   180
         Index           =   13
         Left            =   180
         TabIndex        =   24
         Top             =   2895
         Width           =   1080
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "סԺ����"
         Height          =   180
         Index           =   11
         Left            =   540
         TabIndex        =   20
         Top             =   2460
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "��Ա���"
         Height          =   180
         Index           =   9
         Left            =   540
         TabIndex        =   16
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   4
         Left            =   4080
         TabIndex        =   6
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "���֤����"
         Height          =   180
         Index           =   3
         Left            =   3540
         TabIndex        =   10
         Top             =   1170
         Width           =   900
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "ҽ���ʺ�"
         Height          =   180
         Index           =   1
         Left            =   3720
         TabIndex        =   2
         Top             =   315
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "סԺ��־"
         Height          =   180
         Index           =   14
         Left            =   3720
         TabIndex        =   26
         Top             =   2895
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "�Ƿ����Բ�"
         Height          =   180
         Index           =   12
         Left            =   3540
         TabIndex        =   22
         Top             =   2460
         Width           =   900
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "�ʻ����"
         Height          =   180
         Index           =   10
         Left            =   3720
         TabIndex        =   18
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   8
         Left            =   3720
         TabIndex        =   14
         Top             =   1605
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "��λ����"
         Height          =   180
         Index           =   6
         Left            =   540
         TabIndex        =   12
         Top             =   1605
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   180
         Index           =   5
         Left            =   900
         TabIndex        =   8
         Top             =   1170
         Width           =   360
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "IC����"
         Height          =   180
         Index           =   2
         Left            =   720
         TabIndex        =   4
         Top             =   750
         Width           =   540
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "���Ĵ���"
         Height          =   180
         Index           =   0
         Left            =   540
         TabIndex        =   0
         Top             =   315
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "����(&R)"
      Height          =   400
      Left            =   210
      TabIndex        =   32
      Top             =   5055
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   400
      Left            =   4245
      TabIndex        =   33
      Top             =   5055
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   5400
      TabIndex        =   34
      Top             =   5055
      Width           =   1100
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType  As Byte   'bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
Private mlng����ID As Long
Private mstrReturn As String

Public Function GetPatient(bytType As Byte, lng����ID As Long) As String
'������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    mbytType = bytType
    mlng����ID = lng����ID
    mstrReturn = ""
    Me.Show vbModal
    GetPatient = mstrReturn
    lng����ID = mlng����ID
End Function

Private Sub cboType_Click()
    If cboType.ItemData(cboType.ListIndex) <> 2 And cboType.ItemData(cboType.ListIndex) <> 3 And mbytType = 0 Then
        cmd����.Enabled = False
        txtInfo(14).Enabled = False
        
        '�º����޸���20060512
        Cmd����1.Enabled = False
        txtInfo(16).Enabled = False
    Else
        cmd����.Enabled = True
        txtInfo(14).Enabled = True
    
        '�º����޸���20060512
        Cmd����1.Enabled = True
        txtInfo(16).Enabled = True
        
    End If
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    '���˺�:20040923����
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Function IsValid() As Boolean
    '���˺�:20040923,�����鿨Ч��.
    Dim rsTemp As New ADODB.Recordset
    IsValid = False
    If txtInfo(0).Text = "" Then
        MsgBox "�������", vbInformation, "����"
        Exit Function
    End If
    
    If mbytType = 0 And (cboType.ItemData(cboType.ListIndex) = 2 Or cboType.ItemData(cboType.ListIndex) = 3) And txtInfo(14).Tag = "" And txtInfo(16).Tag = "" Then
        MsgBox "�����루��ѡ����ȷ�ı��ղ���", vbInformation, "�����֤"
        Exit Function
    End If
    If mbytType = 0 And Trim(Txt����.Tag) = "" Then
        ShowMsgbox "������������"
        Exit Function
    End If
    If mbytType <> 2 Then
        If mbytType = 4 Then
            '����鵱ǰ��̬
        Else
            '��鲡��״̬
            gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ҽ����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_����, CStr(txtInfo(1).Text))
            If rsTemp.RecordCount > 0 Then
                If rsTemp("״̬") > 0 Then
                    MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        If mbytType = 0 Or mbytType = 3 Then
            '����
            
        End If
    Else
        '�����������סԺ�ģ�ֻ��ˢ����ʾһ�����ݶ��ѣ�������
        Unload Me
        Exit Function
    End If
        
    IsValid = True
End Function
Private Sub cmdOK_Click()
    Dim strTmp As String
    Dim strAddition As String, strIdentify As String
    If IsValid = False Then Exit Sub
    
    If cboType.ListIndex >= 0 Then
        g�������_����.�������� = cboType.ItemData(cboType.ListIndex)
    End If
    
    With g�������_����
        
        If Val(txtInfo(14).Tag) <> 0 Or Val(txtInfo(16).Tag) = 0 Then
        
            .���ִ��� = cmd����.Tag
            
        End If
        
        If Val(txtInfo(14).Tag) = 0 Or Val(txtInfo(16).Tag) <> 0 Then
        
            .���ִ��� = Cmd����1.Tag
            
        End If
        
        If Val(txtInfo(14).Tag) <> 0 Or Val(txtInfo(16).Tag) <> 0 Then
        
        '�º�����20060228�޸�,��������ҽ�������������Բ�
        '��������ַָʽ,�Զ������ָ���
        
            .���ִ��� = cmd����.Tag & "," & Cmd����1.Tag
            
        End If
        
        If Txt����.Tag <> "" Then
        
        '�º�����20060228�޸�,��������ҽ�������������Բ�
        '��������ַָʽ,�Զ������ָ���
        
            .��ϱ��� = Split(Txt����.Tag, "|||")(0)
            .������� = Split(Txt����.Tag, "|||")(1)
        Else
            If mbytType = 1 Then
                If Trim(Txt����.Text) = "" Then
                    ShowMsgbox "������������"
                    If Txt����.Enabled Then Txt����.SetFocus
                    Exit Sub
                End If
                .������� = Txt����
            End If
        End If
        
        strAddition = "": strIdentify = ""
        strIdentify = .����                               '0����
        strIdentify = strIdentify & ";" & .ҽ����            '1ҽ����
        strIdentify = strIdentify & ";" & .��������            '2����  ��������
        strIdentify = strIdentify & ";" & .����               '3����
        strIdentify = strIdentify & ";" & .�Ա�                 '4�Ա�
        strIdentify = strIdentify & ";" & .��������                    '5��������
        
        strIdentify = strIdentify & ";" & .���֤��                     '6���֤
        strIdentify = strIdentify & ";" & "(" & .��λ���� & ")"                 '7.��λ����(����)
        strAddition = ";0"                                          '8.���Ĵ���
        strAddition = strAddition & ";" & .֧������                              '9 ˳���
        
        strAddition = strAddition & ";" & .�������                                           '10 ��Ա���
        strAddition = strAddition & ";" & .�ʻ����                        '11 ���
        
        strAddition = strAddition & ";0"                                        '12 ��ǰ״̬
        strAddition = strAddition & ";" & txtInfo(14).Tag                        '13 ����ID
        strAddition = strAddition & ";" & IIf(txtInfo(8).Text = "��ְ", "1", "2")   '14��ְ(1,2,3)
        
        strTmp = .����Ա��־ & "|" & .���䱣�� & "|" & .��ҽ�� & "|" & .������ϵ & "|" & .�չ˼��� & "|" & .ְ������ & "|" & .�Ƿ����Բ� & " |" & .�ش󼲲�
        strAddition = strAddition & ";" & strTmp                                '15 ����֤��
        strAddition = strAddition & ";"                                         '16 �����
        strAddition = strAddition & ";"                                         '17 �Ҷȼ�
        strAddition = strAddition & ";" & .�ʻ����                             '18 �ʻ������ۼ�
        strAddition = strAddition & ";" & "0"                                         '19 �ʻ�֧���ۼ�
        
        strAddition = strAddition & ";" & .����ͳ���ۼ�                        '20 ����ͳ���ۼ�
        strAddition = strAddition & ";" & .ͳ��֧���ۼ�                        '21 ͳ�ﱨ���ۼ�
        strAddition = strAddition & ";" & .סԺ����                            '22 סԺ�����ۼ�
        strAddition = strAddition & ";"                                      '23 �������
        strAddition = strAddition & ";"                                      '24 ��������
        strAddition = strAddition & .���߽���ۼ� & ";"                         '25 �����ۼ�
        strAddition = strAddition & ";" & .�𸶶�ҽ�Ʒ��ۼ�                  '26 ����ͳ���޶�
        
    End With
    'Me.Hide
    '���˺�:20040923����
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_����)
    
    '����������
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_���� & ",'��Ժ���','''" & g�������_����.������� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����������")
     
    If mlng����ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng����ID & strAddition
    End If
    Unload Me
End Sub

Private Sub cmdRead_Click()
    If ��ݼ���_����(mbytType) = False Then Exit Sub
    
    '��ʾ����������Ϣ
    With g�������_����
        txtInfo(0).Text = .����      '����
        txtInfo(2).Text = .����      '����
        txtInfo(5).Text = .���֤��      '���֤��
        txtInfo(3).Text = .����    '����
        txtInfo(4).Text = .�Ա�
        txtInfo(7).Text = .��������
        txtInfo(1).Text = .ҽ����      'ҽ����
        txtInfo(6).Text = .��λ����
        txtInfo(8).Text = IIf(Val(.�������) = "0", "��ְ", "����")
        
        txtInfo(9).Text = .�ʻ����
        txtInfo(10).Text = .סԺ����
        txtInfo(11).Text = IIf(.�Ƿ����Բ� = "0", "����", "��")
        txtInfo(12).Text = IIf(.�ش󼲲� = "0", "����", "��")
        txtInfo(13).Text = IIf(.סԺ��־ = "0", "��סԺ", "סԺ")
        '���˺�:20040923����
        txtInfo(15).Text = .֧������
    End With
End Sub

 Private Sub SetCtlBackColor()
    '���˺�:20040923����
    Dim i As Long
    For i = 0 To txtInfo.UBound
        If i <> 14 Or i <> 16 Then
        Else
            txtInfo(i).BackColor = &H8000000F
        End If
    Next
 End Sub

Private Sub cmd����_Click()
    Dim rsTemp As New ADODB.Recordset
    
    If mbytType = 0 Or mbytType = 3 Then
        'ֻѡ���������ֲ�
        gstrSQL = " Select A.ID,����,A.����,A.����,decode(A.���,'1','���Բ�','2','���ֲ�','��ͨ��') as ��� " & _
                "   From ���ղ��� A " & _
                "   where nvl(a.���,0)<>'0' and A.����=" & TYPE_���� & _
                "   order by ����"
    ElseIf mbytType = 1 Or mbytType = 4 Then
        gstrSQL = " Select A.ID,����,A.����,A.����,decode(A.���,'1','���Բ�','2','���ֲ�','��ͨ��') as ��� " & _
                "   From ���ղ��� A " & _
                "   where nvl(a.���,0)='0' and A.����=" & TYPE_���� & _
                "   order by ����"
    Else
        gstrSQL = " Select A.ID,����,A.����,A.����,decode(A.���,'1','���Բ�','2','���ֲ�','��ͨ��') as ��� " & _
                "   From ���ղ��� A " & _
                "   where A.����=" & TYPE_���� & _
                "   order by ����"
    End If
    
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "ҽ������", , txtInfo(14).Text)
    If rsTemp.State = 0 Then Exit Sub
    
    If Not rsTemp Is Nothing Then
        txtInfo(14).Text = "(" & rsTemp!���� & ")" & rsTemp!����
        txtInfo(14).Tag = rsTemp("ID")
        cmd����.Tag = Nvl(rsTemp!����)
        zlControl.TxtSelAll txtInfo(14)
    End If
    txtInfo(14).SetFocus
End Sub

Private Sub Cmd����1_Click()
   
'�º�����20060228�޸�

   Dim rsTemp As New ADODB.Recordset
    
    If mbytType = 0 Or mbytType = 3 Then
        'ֻѡ���������ֲ�
        gstrSQL = " Select A.ID,����,A.����,A.����,decode(A.���,'1','���Բ�','2','���ֲ�','��ͨ��') as ��� " & _
                "   From ���ղ��� A " & _
                "   where nvl(a.���,0)<>'0' and A.����=" & TYPE_���� & _
                "   order by ����"
    ElseIf mbytType = 1 Or mbytType = 4 Then
        gstrSQL = " Select A.ID,����,A.����,A.����,decode(A.���,'1','���Բ�','2','���ֲ�','��ͨ��') as ��� " & _
                "   From ���ղ��� A " & _
                "   where nvl(a.���,0)='0' and A.����=" & TYPE_���� & _
                "   order by ����"
    Else
        gstrSQL = " Select A.ID,����,A.����,A.����,decode(A.���,'1','���Բ�','2','���ֲ�','��ͨ��') as ��� " & _
                "   From ���ղ��� A " & _
                "   where A.����=" & TYPE_���� & _
                "   order by ����"
    End If
    
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "ҽ������", , txtInfo(16).Text)
    If rsTemp.State = 0 Then Exit Sub
    
    If Not rsTemp Is Nothing Then
        txtInfo(16).Text = "(" & rsTemp!���� & ")" & rsTemp!����
        txtInfo(16).Tag = rsTemp("ID")
        Cmd����1.Tag = Nvl(rsTemp!����)
        zlControl.TxtSelAll txtInfo(16)
    End If
    txtInfo(16).SetFocus
End Sub

Private Sub Form_Load()
    '���˺�:û�б�Ҫ������ص�
    cboType.Clear
    '�º�����20060512�޸�
    g�������_����.���ִ��� = ""
    
    If mbytType = 0 Or mbytType = 3 Then
        cboType.Enabled = True
        txtInfo(14).Enabled = True
        Txt����.Enabled = True
        
        '�º�����20060228�޸�
        'ԭ����Ҫ������������˾�������Բ����ܳ�������
        txtInfo(16).Enabled = True
                
        cboType.AddItem "1-��ͨ"
        cboType.ItemData(cboType.NewIndex) = 1
        cboType.AddItem "2-���Բ�"
        cboType.ItemData(cboType.NewIndex) = 2
        cboType.AddItem "3-�ش󼲲�"
        cboType.ItemData(cboType.NewIndex) = 3
        cboType.AddItem "4-�չ˶���"
        cboType.ItemData(cboType.NewIndex) = 4
        cboType.AddItem "5-������"
        cboType.ItemData(cboType.NewIndex) = 5
        cboType.AddItem "6-�ƻ�����"
        cboType.ItemData(cboType.NewIndex) = 6
        cboType.AddItem "7-����"
        cboType.ItemData(cboType.NewIndex) = 7
        cboType.ListIndex = 0
    Else
        If mbytType = 1 Then
            Txt����.Enabled = True
        Else
            Txt����.Enabled = False
        End If
        
        '�º�����20060228�޸�
        txtInfo(16).Enabled = False
        Cmd����1.Enabled = False
        
        cboType.AddItem "1-��ͨ"
        cboType.ItemData(cboType.NewIndex) = 1
        cboType.AddItem "2.�ƻ�����סԺ����"
        cboType.ItemData(cboType.NewIndex) = 6
        cboType.ListIndex = 0
    End If
    
    Call SetCtlBackColor
End Sub

Private Sub txtInfo_Change(Index As Integer)
    txtInfo(Index).Tag = ""
        
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '���˺�:20040923����
    Dim strLike As String, StrInput As String
    Dim blnCancel As Boolean
    Dim rsTmp As ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    '�º�����20060228�޸�
    If Index = 14 Or Index = 16 Then
        'ѡ����
        If txtInfo(Index).Text = "" Then
            Call zlCommFun.PressKey(vbKeyTab) '��������
        Else
            strLike = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
            StrInput = strLike & UCase(txtInfo(Index).Text)
            
            If mbytType = 0 Or mbytType = 3 Then
                'ֻѡ���������ֲ�
                gstrSQL = " Select A.ID,����,A.����,A.����,decode(A.���,'1','���Բ�','2','���ֲ�','��ͨ��') as ��� " & _
                        "   From ���ղ��� A " & _
                        "   where nvl(a.���,0)<>'0' and A.����=" & TYPE_���� & _
                        "        and ( ���� like '" & StrInput & "%' or ���� like '" & StrInput & "%' or ���� like '" & StrInput & "%' )" & _
                        "        And Rownum<=200" & _
                        "   order by ����"
            ElseIf mbytType = 1 Or mbytType = 4 Then
                gstrSQL = " Select A.ID,����,A.����,A.����,decode(A.���,'1','���Բ�','2','���ֲ�','��ͨ��') as ��� " & _
                        "   From ���ղ��� A " & _
                        "   where nvl(a.���,0)='0' and A.����=" & TYPE_���� & _
                        "        and ( ���� like '" & StrInput & "%' or ���� like '" & StrInput & "%' or ���� like '" & StrInput & "%' )" & _
                        "        And Rownum<=200" & _
                        "   order by ����"
            Else
                gstrSQL = " Select A.ID,����,A.����,A.����,decode(A.���,'1','���Բ�','2','���ֲ�','��ͨ��') as ��� " & _
                        "   From ���ղ��� A " & _
                        "   where A.����=" & TYPE_���� & _
                        "        and ( ���� like '" & StrInput & "%' or ���� like '" & StrInput & "%' or ���� like '" & StrInput & "%' )" & _
                        "        And Rownum<=200" & _
                        "   order by ����"
            End If
            Set rsTmp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "����ѡ��", , , , , , True, _
                txtInfo(Index).Left + Me.Left, _
                txtInfo(Index).Top + Me.Top, txtInfo(Index).Height, blnCancel, , True)
                
            If Not rsTmp Is Nothing Then
                txtInfo(Index).Text = "(" & rsTmp!���� & ")" & rsTmp!����
                txtInfo(Index).Tag = rsTmp("ID")
                
                '�º�����20060228�޸�
                If Index = 14 Then
                    cmd����.Tag = Nvl(rsTmp!����)
                ElseIf Index = 16 Then
                    Cmd����1.Tag = Nvl(rsTmp!����)
                End If
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��Ĳ��ֱ��롣", vbInformation, gstrSysName
                End If
                txtInfo(Index).SetFocus
            End If
        End If
    End If
    zlCommFun.PressKey vbKeyTab
End Sub


Private Sub Txt����_Change()
    Txt����.Tag = ""
End Sub

Private Sub Txt����_GotFocus()
    zlControl.TxtSelAll Txt����
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub Txt����_KeyPress(KeyAscii As Integer)
  Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strLike As String, str�Ա� As String
    Dim StrInput As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Txt����.Text = "" Then
            Call zlCommFun.PressKey(vbKeyTab) '��������
        Else
            strLike = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
            StrInput = UCase(Txt����.Text)
            str�Ա� = g�������_����.�Ա�
            If str�Ա� = "��" Then
                str�Ա� = " And (A.�Ա�����='��' Or A.�Ա����� is NULL)"
            ElseIf str�Ա� = "Ů" Then
                str�Ա� = " And (A.�Ա�����='Ů' Or A.�Ա����� is NULL)"
            End If
            
            strSQL = "Select A.ID,A.����,A.����,A.����,A.����,A.˵��,A.�Ա�����,B.���" & _
                " From ��������Ŀ¼ A,����������� B" & _
                " Where A.���=B.���� And A.��� Not IN('B','Z')" & _
                " And (A.���� Like '" & StrInput & "%'" & _
                " Or Upper(A.����) Like '" & strLike & StrInput & "%'" & _
                " Or Upper(A.����) Like '" & strLike & StrInput & "%'" & _
                " Or Upper(A.����) Like '" & strLike & StrInput & "%')" & _
                " And Rownum<=100" & str�Ա� & _
                " Order by A.���,A.����"
                
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "��������Input", , , , , , True, _
                Txt����.Left + Me.Left, _
                Txt����.Top + Me.Top, Txt����.Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                Txt����.Text = "(" & rsTmp!���� & ")" & rsTmp!����
                Txt����.Tag = rsTmp!���� & "|||" & rsTmp!����
                If cmdOK.Enabled Then
                    cmdOK.SetFocus
                Else
                    Call zlCommFun.PressKey(vbKeyTab)
                End If
            Else
                If mbytType <> 1 Then
                    If Not blnCancel Then
                        MsgBox "û���ҵ�ƥ��ļ������롣", vbInformation, gstrSysName
                    End If
                    Call Txt����_GotFocus
                    Txt����.SetFocus
                Else
                        If cmdOK.Enabled Then
                            cmdOK.SetFocus
                        Else
                            Call zlCommFun.PressKey(vbKeyTab)
                        End If
                End If
            End If
        End If
    Else
        zlControl.TxtCheckKeyPress Txt����, KeyAscii, m�ı�ʽ
    End If
End Sub


