VERSION 5.00
Begin VB.Form frmIdentify�ϳ����� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������֤"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdChangePassWord 
      Caption         =   "�޸�����(&P)"
      Height          =   350
      Left            =   45
      TabIndex        =   30
      Top             =   5415
      Width           =   1380
   End
   Begin VB.TextBox txtPassWord 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   5100
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   900
      Width           =   2535
   End
   Begin VB.TextBox txtEdit 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   300
      Left            =   5100
      MaxLength       =   25
      TabIndex        =   7
      Tag             =   "��ᱣ�Ϻ�"
      Top             =   1305
      Width           =   2535
   End
   Begin VB.CommandButton cmd�鿨 
      Caption         =   "���¶���(&R)"
      Height          =   350
      Left            =   1455
      TabIndex        =   27
      Top             =   5415
      Width           =   1305
   End
   Begin VB.ComboBox cbo�籣 
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   900
      Width           =   2805
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6600
      TabIndex        =   25
      Top             =   5415
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5445
      TabIndex        =   24
      Top             =   5415
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -60
      TabIndex        =   28
      Top             =   615
      Width           =   8340
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -555
      TabIndex        =   26
      Top             =   5220
      Width           =   8340
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   29
      Left            =   6090
      TabIndex        =   68
      ToolTipText     =   "����סԺ����"
      Top             =   4935
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   27
      Left            =   6855
      TabIndex        =   67
      Top             =   4890
      Width           =   780
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   28
      Left            =   4350
      TabIndex        =   66
      ToolTipText     =   "����סԺ����"
      Top             =   4935
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   26
      Left            =   5100
      TabIndex        =   65
      Top             =   4875
      Width           =   885
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��ȡʱ��"
      Height          =   180
      Index           =   27
      Left            =   2175
      TabIndex        =   64
      ToolTipText     =   "����סԺ����"
      Top             =   4935
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   25
      Left            =   2895
      TabIndex        =   63
      Top             =   4890
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ɷ�����"
      Height          =   180
      Index           =   26
      Left            =   240
      TabIndex        =   62
      ToolTipText     =   "�ɷ�����"
      Top             =   4935
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   24
      Left            =   960
      TabIndex        =   61
      Top             =   4890
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ѱ�����"
      Height          =   180
      Index           =   25
      Left            =   4350
      TabIndex        =   60
      ToolTipText     =   "�����ѱ������"
      Top             =   4575
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   23
      Left            =   5100
      TabIndex        =   59
      Top             =   4515
      Width           =   2535
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "סԺ����"
      Height          =   180
      Index           =   24
      Left            =   2175
      TabIndex        =   58
      ToolTipText     =   "����סԺ����"
      Top             =   4575
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   22
      Left            =   2895
      TabIndex        =   57
      Top             =   4530
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����Ա����"
      Height          =   180
      Index           =   23
      Left            =   60
      TabIndex        =   56
      Top             =   4575
      Width           =   900
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   21
      Left            =   960
      TabIndex        =   55
      Top             =   4530
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   22
      Left            =   6090
      TabIndex        =   54
      Top             =   4200
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   20
      Left            =   6855
      TabIndex        =   53
      Top             =   4155
      Width           =   780
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   21
      Left            =   4350
      TabIndex        =   52
      Top             =   4200
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   19
      Left            =   5100
      TabIndex        =   51
      Top             =   4140
      Width           =   885
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����Ա״̬"
      Height          =   180
      Index           =   20
      Left            =   1995
      TabIndex        =   50
      Top             =   4200
      Width           =   900
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   18
      Left            =   2895
      TabIndex        =   49
      Top             =   4155
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   19
      Left            =   600
      TabIndex        =   48
      Top             =   4200
      Width           =   360
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   17
      Left            =   960
      TabIndex        =   47
      Top             =   4155
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����ҽ��"
      Height          =   180
      Index           =   18
      Left            =   4350
      TabIndex        =   46
      Top             =   3810
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   16
      Left            =   6855
      TabIndex        =   45
      Top             =   3765
      Width           =   780
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����ҽ��"
      Height          =   180
      Index           =   17
      Left            =   6090
      TabIndex        =   44
      Top             =   3810
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   15
      Left            =   5100
      TabIndex        =   43
      Top             =   3750
      Width           =   885
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��ؾ�ס"
      Height          =   180
      Index           =   16
      Left            =   2175
      TabIndex        =   42
      Top             =   3810
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   14
      Left            =   2895
      TabIndex        =   41
      Top             =   3765
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����ID"
      Height          =   180
      Index           =   15
      Left            =   420
      TabIndex        =   40
      Top             =   3810
      Width           =   540
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   13
      Left            =   960
      TabIndex        =   39
      Top             =   3765
      Width           =   990
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   6855
      TabIndex        =   38
      Top             =   3368
      Width           =   780
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   11
      Left            =   5100
      TabIndex        =   37
      Top             =   3360
      Width           =   885
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "סԺ״̬"
      Height          =   180
      Index           =   14
      Left            =   6090
      TabIndex        =   36
      Top             =   3420
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��Ա����"
      Height          =   180
      Index           =   13
      Left            =   4350
      TabIndex        =   35
      Top             =   3420
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   2895
      TabIndex        =   34
      Top             =   3368
      Width           =   900
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   960
      TabIndex        =   33
      Top             =   3368
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ְ�Ƽ���"
      Height          =   180
      Index           =   12
      Left            =   2175
      TabIndex        =   32
      Top             =   3420
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ְ�񼶱�"
      Height          =   180
      Index           =   11
      Left            =   240
      TabIndex        =   31
      Top             =   3420
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   9
      Left            =   4710
      TabIndex        =   2
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "��¼��"
      Height          =   180
      Index           =   2
      Left            =   4530
      TabIndex        =   8
      Top             =   1785
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   4
      Left            =   4350
      TabIndex        =   20
      Top             =   2625
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   6
      Left            =   5100
      TabIndex        =   21
      Top             =   2565
      Width           =   2535
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   1305
      Width           =   2805
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "ҽ�����˻�����Ϣ��ʾ������ͨ��[���¶���]��ť���½��ж�ȡ���˻�����Ϣ��"
      Height          =   180
      Left            =   630
      TabIndex        =   29
      Top             =   360
      Width           =   6300
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmIdentify�ϳ�����.frx":0000
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   1365
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   1
      Left            =   600
      TabIndex        =   10
      Top             =   1785
      Width           =   360
   End
   Begin VB.Label lblInf 
      AutoSize        =   -1  'True
      Caption         =   "ҽ��֤��"
      Height          =   180
      Left            =   4350
      TabIndex        =   6
      Top             =   1365
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   180
      Index           =   3
      Left            =   600
      TabIndex        =   12
      Top             =   2190
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���֤��"
      Height          =   180
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   2625
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ʻ����"
      Height          =   180
      Index           =   6
      Left            =   4350
      TabIndex        =   16
      Top             =   2190
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�籣����"
      Height          =   180
      Index           =   7
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   8
      Left            =   2535
      TabIndex        =   14
      Top             =   2190
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��λ����"
      Height          =   180
      Index           =   10
      Left            =   240
      TabIndex        =   22
      Top             =   3015
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   1
      Left            =   960
      TabIndex        =   11
      Top             =   1725
      Width           =   2805
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   2
      Left            =   5100
      TabIndex        =   9
      Top             =   1725
      Width           =   2535
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   960
      TabIndex        =   13
      Top             =   2130
      Width           =   990
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   2895
      TabIndex        =   15
      Top             =   2145
      Width           =   870
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   5
      Left            =   960
      TabIndex        =   19
      Top             =   2565
      Width           =   2805
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   960
      TabIndex        =   23
      Top             =   2970
      Width           =   6675
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   7
      Left            =   5100
      TabIndex        =   17
      Top             =   2130
      Width           =   2535
   End
End
Attribute VB_Name = "frmIdentify�ϳ�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����,88-��סԺ��Ϣ���в�ѯ

Private mlng����ID As Long
Private mstrReturn As String
Private mintPreCol As Integer, mintsort As Integer
Private mblnFirst As Boolean        '��һ����ϵͳʱ����
Private mblnChange As Boolean
Private Sub cbo�籣_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdChangePassWord_Click()
    Dim strOldPassWord As String
    Dim strNewPassWord As String
    Dim StrInput As String
    Dim strOutput As String
    
    If cbo�籣.ListIndex < 0 Then
        ShowMsgbox "δѡ��ҽ����������,��ѡ��!"
        Exit Sub
    End If
    If Split(cbo�籣.Text, "_")(0) = "" Then
        ShowMsgbox "ҽ����������Ϊ����,������ѡ��!"
        Exit Sub
    End If
    g�������_�ϳ�����.�������� = Split(cbo�籣.Text, "-")(0)
    strNewPassWord = frm�޸�����.ChangePassword(strOldPassWord, strOldPassWord)
    If strOldPassWord = strNewPassWord Then Exit Sub
    If strNewPassWord = "" Then Exit Sub
    
    '    YBJGBH  PCHAR   ���ջ������
    '    COLDPASS    PCHAR   ������
    '    CNEWPAS PCHAR   ������

    StrInput = g�������_�ϳ�����.��������
    StrInput = StrInput & vbTab & strOldPassWord
    StrInput = StrInput & vbTab & strNewPassWord
    If ҵ������_�ϳ�����(�޸�����_����, StrInput, strOutput) = False Then Exit Sub
    MsgBox "�����޸ĳɹ�!", vbInformation + vbDefaultButton1, gstrSysName
    
End Sub

Private Sub cmd�鿨_Click()

    If mbytType = 1 Or mbytType = 4 Or mbytType = 88 Then
        If ��ȡ�α���Ա��Ϣ_סԺ = False Then
            cmdȷ��.Enabled = False
            Call ClearData
            Exit Sub
        End If
        Call LoadCtrlData
        cmdȷ��.Enabled = True
        Exit Sub
    End If
    If ��ȡ�α���Ա��Ϣ = False Then
         cmdȷ��.Enabled = False
         Call ClearData
         Exit Sub
     End If
     Call LoadCtrlData
     cmdȷ��.Enabled = True
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    Call ClearData
    
    cmdȷ��.Enabled = False
    If mbytType = 1 Or mbytType = 4 Or mbytType = 88 Then
        txtPassWord.Enabled = False
        txtPassWord.BackColor = lblEdit(0).BackColor
        txtEdit.Enabled = True
        txtEdit.BackColor = cbo�籣.BackColor
        '������:20050420 �����Ƕ�IC���������룬�ʲ��ṩ������޸Ĺ���
        cmdChangePassWord.Enabled = False
    Else
        txtPassWord.Enabled = True
        txtPassWord.BackColor = cbo�籣.BackColor
    End If
End Sub

Private Sub SetOKCtrl(ByVal blnEn As Boolean)
    cmdȷ��.Enabled = blnEn
End Sub

Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��֤���ݵĺϷ���
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim StrInput As String, strOutput As String
    Dim lng״̬ As Long
    
    IsValid = False
    If Trim(g�������_�ϳ�����.����) = "" Then
        MsgBox "��û���������֤��", vbInformation, gstrSysName
        If cmd�鿨.Enabled Then cmd�鿨.SetFocus
        Exit Function
    End If
    
     If cbo�籣.Text = "" Then
        ShowMsgbox "�籣������δѡ��"
        Exit Function
    End If
      
    If mbytType <> 2 And mbytType <> 88 Then
        If mbytType = 4 Then
            '����鵱ǰ״̬
        Else
            '��鲡��״̬
            gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ҽ����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�ϳ�����, g�������_�ϳ�����.ҽ��֤��)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("״̬") > 0 Then
                    MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Else
        '�����������סԺ�ģ�ֻ��ˢ����ʾһ�����ݶ��ѣ�������
         '��ȷ����ǰ״̬,��Ϊ��ǰ״̬�ǲ��ܸı��
        gstrSQL = "Select * from �����ʻ� where ����=[1] and  ҽ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_�ϳ�����, g�������_�ϳ�����.ҽ��֤��)
        If Not rsTemp.EOF Then
            mlng����ID = Nvl(rsTemp!����ID, 0)
        Else
            mlng����ID = 0
        End If
        mstrReturn = mlng����ID
        Unload Me
        Exit Function
    End If
    IsValid = True
End Function

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Dim lng����ID As Long
    Dim StrInput  As String, strOutput As String
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    Dim str�籣 As String
    Dim int��ǰ״̬ As Integer
    Dim lng״̬ As Long
    
    
    g�������_�ϳ�����.�������� = Split(cbo�籣.Text, "-")(0)
    g�������_�ϳ�����.�籣���� = cbo�籣.ItemData(cbo�籣.ListIndex)
    If IsValid = False Then Exit Sub
    
    int��ǰ״̬ = 0
    If mbytType = 4 Then
        '��ȷ����ǰ״̬,��Ϊ��ǰ״̬�ǲ��ܸı��
        gstrSQL = "Select * from �����ʻ� where ����=" & TYPE_�ϳ����� & " and  ҽ����='" & g�������_�ϳ�����.ҽ��֤�� & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng����ID = Nvl(rsTemp!����ID, 0)
            int��ǰ״̬ = Nvl(rsTemp!��ǰ״̬, 0)
        End If
        rsTemp.Close
    End If
    g�������_�ϳ�����.���� = txtPassWord.Text
    
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    With g�������_�ϳ�����
        
        strIdentify = .ҽ������                                '0����
        strIdentify = strIdentify & ";" & .ҽ��֤��             '1ҽ����
        strIdentify = strIdentify & ";"                    '2����
        strIdentify = strIdentify & ";" & .����               '3����
        strIdentify = strIdentify & ";" & Decode(.�Ա�, "1", "��", "2", "Ů", .�Ա�)              '4�Ա�
        strIdentify = strIdentify & ";" & .��������                '5��������
        strIdentify = strIdentify & ";" & .���֤����            '6���֤
        strIdentify = strIdentify & ";" & .��λ����     '7.��λ����(����)
        strAddition = ";0" & .�籣����                                           '8.���Ĵ���
        strAddition = strAddition & ";" & .��¼��                               '9.˳���
        strAddition = strAddition & ";" & .��Ա����                                  '10��Ա���
        strAddition = strAddition & ";" & .�ʻ����                  '11�ʻ����
        
        strAddition = strAddition & ";" & int��ǰ״̬                            '12��ǰ״̬
        strAddition = strAddition & ";"             '13����ID
        strAddition = strAddition & ";1"                        '14��ְ(1,2,3)
        strAddition = strAddition & ";" & .��������            '15����֤��
        strAddition = strAddition & ";" & .����                     '16�����
        strAddition = strAddition & ";"                         '17�Ҷȼ�
        strAddition = strAddition & ";" & .�ʻ����                           '18�ʻ������ۼ�
        strAddition = strAddition & ";0"                            '19�ʻ�֧���ۼ�
        strAddition = strAddition & ";0"                            '20���깤���ܶ�
        strAddition = strAddition & ";"                             '21סԺ�����ۼ�
    End With
    
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_�ϳ�����)
    If mlng����ID = 0 Then Exit Sub
    
    If mbytType = 1 Or mbytType = 4 Then
        '���¸�����Ϣ(���̲�������:)
        '    ����ID_IN,����ID_IN,�μӹ�������_IN,��������_IN,ְ�񼶱�_IN,ְ�Ƽ���_IN,��ؾ�ס��־_IN
        '    ��λID_IN,����_IN,סԺ����_IN,����ҽ�Ʊ�־_IN,����ҽ�Ʊ�־_IN,����Ա��־_IN,��������״̬_IN,�������״̬_IN
        '    ����Ա����״̬_IN ,����סԺ����_IN,�����ѱ������_IN,�ɷ�����_IN,��ȡʱ��_IN,סԺ��¼��_IN

        gstrSQL = "zl_�����ʻ�������Ϣ_Update("
        gstrSQL = gstrSQL & "" & mlng����ID & ","
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.����ID & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.�μӹ������� & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.�������� & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.ְ�񼶱� & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.ְ�Ƽ��� & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.��ؾ�ס��־ & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.��λID & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.���� & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.סԺ���� & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.����ҽ�Ʊ�־ & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.����ҽ�Ʊ�־ & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.����Ա��־ & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.��������״̬ & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.�������״̬ & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.����Ա����״̬ & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.����סԺ���� & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.�����ѱ������ & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.�ɷ����� & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.��ȡʱ�� & "',"
        gstrSQL = gstrSQL & "'" & g�������_�ϳ�����.סԺ��¼�� & "')"
        ExecuteProcedure_�ϳ����� "�����ʻ�������Ϣ"
    Else
    End If
    g�������_�ϳ�����.����ID = mlng����ID
    
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng����ID & strAddition
    End If
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Public Function GetPatient(Optional bytType As Byte, Optional lng����ID As Long = 0) As String
    mbytType = bytType
    mlng����ID = lng����ID
    mstrReturn = ""
    DebugTool "���������֤,����ʼ���������Ϣ"
    If Load�籣���� = False Then
        DebugTool "����ʧ��(�����֤)"
        Exit Function
    End If
    DebugTool "����ɹ�(�����֤)"
    
    Me.Show 1
    lng����ID = mlng����ID
    GetPatient = mstrReturn
End Function
Private Sub LoadCtrlData()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    With g�������_�ϳ�����
        lblEdit(0).Caption = .ҽ������
        txtEdit.Text = .ҽ��֤��
        lblEdit(1).Caption = .����
        lblEdit(2).Caption = IIf(mbytType = 1 Or mbytType = 4, .סԺ��¼��, .��¼��)
        lblEdit(3).Caption = Decode(.�Ա�, "1", "��", "2", "Ů", .�Ա�)
        lblEdit(4).Caption = .����
        lblEdit(5).Caption = .���֤����
        lblEdit(6).Caption = .��������
        lblEdit(7).Caption = Format(.�ʻ����, "####0.00;-####0.00;;")
        lblEdit(8).Caption = .��λ����
        lblEdit(9).Caption = .ְ�񼶱�
        lblEdit(10).Caption = .ְ�Ƽ���
        lblEdit(11).Caption = .��Ա����
        lblEdit(12).Caption = ""
        
        lblEdit(13).Caption = .����ID
        lblEdit(14).Caption = .��ؾ�ס��־
        lblEdit(15).Caption = .����ҽ�Ʊ�־
        lblEdit(16).Caption = .����ҽ�Ʊ�־
        lblEdit(17).Caption = .����
        lblEdit(18).Caption = .����Ա��־
        lblEdit(19).Caption = .��������״̬
        lblEdit(20).Caption = .�������״̬
        lblEdit(21).Caption = .����Ա����״̬
        lblEdit(22).Caption = .����סԺ����
        lblEdit(23).Caption = .�����ѱ������
        
        lblEdit(24).Caption = .�ɷ�����
        lblEdit(25).Caption = .��ȡʱ��
        
        lblEdit(26).Caption = .��������
        lblEdit(27).Caption = .�μӹ�������
        
    End With
End Sub

Private Sub Form_Load()
        mblnFirst = True
End Sub

Private Function Load�籣����() As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select * From ��������Ŀ¼ where ����=" & TYPE_�ϳ����� & " and ���<>0 order by ����"

    Err = 0
    On Error GoTo errHand:
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption & "�籣����Ŀ¼"
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "�����籣����Ŀ¼�����ڲ��������ػ���!"
        Exit Function
    End If
    
    With rsTemp
        cbo�籣.Clear
        Do While Not .EOF
            cbo�籣.AddItem Nvl(!����) & "--" & Nvl(!����)
            cbo�籣.ItemData(cbo�籣.NewIndex) = Nvl(!���, 0)
            .MoveNext
        Loop
    End With
    cbo�籣.ListIndex = 0
    SetDefaultSel
    Load�籣���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function SetDefaultSel() As Boolean
    Dim strReg As String
    Dim i As Integer
    
    SetDefaultSel = False
    Err = 0: On Error GoTo errHand:
    Call GetRegInFor(g����ģ��, "ҽ��", "�籣��������", strReg)
    If cbo�籣.ListCount = 0 Then Exit Function
    For i = 0 To cbo�籣.ListCount - 1
        If Split(cbo�籣.List(i) & "--", "--")(0) = strReg Then
            cbo�籣.ListIndex = i
            Exit For
        End If
    Next
    If cbo�籣.ListIndex < 0 Then
        cbo�籣.ListIndex = 0
    End If
    SetDefaultSel = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function ��ȡ�α���Ա��Ϣ_סԺ() As Boolean
    '����:��ȡ�α���Ա��Ϣ(סԺ����)
    Dim StrInput As String, strOutput As String
    Dim bln���� As Boolean
    Dim strArr As Variant
    
    ��ȡ�α���Ա��Ϣ_סԺ = False
    Err = 0: On Error GoTo errHand:
    If Trim(txtEdit.Text) <> "" Then
        StrInput = Trim(txtEdit.Text) & vbTab
        bln���� = True
    End If
    StrInput = StrInput & Split(cbo�籣.Text, "--")(0)
    '������:20050420 ����Ƕ�IC��,���ȸ���IC���ж�����ҽ��֤��,��ȡ�α���Ա�Ļ�������
    If bln���� = False Then
        If ҵ������_�ϳ�����(�����Ա����_����_����, StrInput, strOutput) = False Then
        Exit Function
        Else: strArr = Split(strOutput, "||")
              StrInput = Split(strArr(0), "--")(0) & vbTab & Split(cbo�籣.Text, "--")(0)
              txtEdit.Text = strArr(0)
              lblEdit(0).Caption = strArr(1)
        End If
    End If
    If ҵ������_�ϳ�����(�����Ա����_ҽ����_����, StrInput, strOutput) = False Then
       Exit Function
    End If
    strArr = Split(strOutput, "||")
    
    '����ID||�籣���||����||�Ա�||�������ڣ���ʽ��YYYY-MM-DD��||�μӹ�������||��������||ְ�񼶱�||ְ�Ƽ���||��Ա����||
    '��ؾ�ס��־||��λID||��λ����||����||����||ҽ��֤��||סԺ����||����ҽ�Ʊ�־||����ҽ�Ʊ�־||����Ա��־||����ҽ�ƴ���״̬||
    '����ҽ�ƴ���״̬||����Ա����״̬||������Ժ����||�����ѱ������||�ɷ�����||��ȡʱ��||סԺ��¼��||

    With g�������_�ϳ�����
        .ҽ������ = lblEdit(0).Caption
        .ҽ��֤�� = txtEdit.Text
        .��¼�� = strArr(0)
        .���� = strArr(2)
        .���֤���� = strArr(1)
        .��λ���� = strArr(12)
        .�Ա� = strArr(3)
        .�������� = zlCommFun.AddDate(strArr(4))
        .���� = Val(strArr(14))
        .�������� = Split(cbo�籣.Text, "--")(0)
        .�籣���� = cbo�籣.ItemData(cbo�籣.ListIndex)
        .����ID = strArr(0)
        .�μӹ������� = strArr(5)
        
        .�������� = strArr(6)
        .ְ�񼶱� = strArr(7)
        .ְ�Ƽ��� = strArr(8)
        .��Ա���� = strArr(9)
        .��ؾ�ס��־ = strArr(10)
        .��λID = strArr(11)
        .���� = strArr(13)   '����
        .סԺ���� = strArr(16)
        .����ҽ�Ʊ�־ = "" 'strArr(17)
        .����ҽ�Ʊ�־ = "" 'strArr(18)
        .����Ա��־ = "" 'strArr(19)
        .��������״̬ = "" 'strArr(20)
        .�������״̬ = "" 'strArr(21)
        .����Ա����״̬ = "" 'strArr(22)
        .����סԺ���� = "" 'strArr(23)
        .�����ѱ������ = "" 'strArr(24)
        .�ɷ����� = "" 'strArr(25)
        .��ȡʱ�� = "" 'strArr(26)
        .סԺ��¼�� = "" 'strArr(27)
        .strסԺ��Ϣ = strOutput
    End With
    ��ȡ�α���Ա��Ϣ_סԺ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function ��ȡ�α���Ա��Ϣ() As Boolean
    '��ȡ�α���Ա��Ϣ
    Dim StrInput As String
    Dim strOutput As String
    Dim strArr
    
    ��ȡ�α���Ա��Ϣ = False
    
    Err = 0
    On Error GoTo errHand:
    If cbo�籣.ListIndex < 0 Then
        MsgBox "δѡ���籣��������,��ѡ��!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    With g�������_�ϳ�����
        .�������� = Split(cbo�籣.Text, "--")(0)
    End With
    If txtPassWord.Text = "" Then
           MsgBox "��δ��������,������!", vbInformation + vbDefaultButton1, gstrSysName
           If txtPassWord.Enabled Then txtPassWord.SetFocus
           g�������_�ϳ�����.���� = ""
           Exit Function
    End If
    g�������_�ϳ�����.���� = txtPassWord.Text
    If ҵ������_�ϳ�����(��òα���Ա����_����, "", strOutput) = False Then
        Call ClearData
        Exit Function
    End If
    strArr = Split(strOutput, "||")
    '����:ҽ������||ҽ��֤��||���˼�¼��||����||���֤����||��λ����||�Ա�||��������
    
    With g�������_�ϳ�����
        .ҽ������ = strArr(0)
        .ҽ��֤�� = strArr(1)
        .��¼�� = strArr(2)
        .���� = strArr(3)
        .���֤���� = strArr(4)
        .��λ���� = strArr(5)
        .�Ա� = strArr(6)
        .�������� = zlCommFun.AddDate(strArr(7))
        .���� = Get����(.��������)
        .�籣���� = cbo�籣.ItemData(cbo�籣.ListIndex)
        
        .����ID = ""
        .�μӹ������� = ""
        
        .�������� = ""
        .ְ�񼶱� = ""
        .ְ�Ƽ��� = ""
        .��Ա���� = ""
        .��ؾ�ס��־ = ""
        .��λID = ""
        .���� = ""
        .סԺ���� = ""
        .����ҽ�Ʊ�־ = ""
        .����ҽ�Ʊ�־ = ""
        .����Ա��־ = ""
        .��������״̬ = ""
        .�������״̬ = ""
        .����Ա����״̬ = ""
        .����סԺ���� = ""
        .�����ѱ������ = ""
        .�ɷ����� = ""
        .��ȡʱ�� = ""
        .סԺ��¼�� = ""
        
        .strסԺ��Ϣ = ""
    End With
    
    '��ȡ�ʻ����
    '    YBJGBH  PCHAR   ���ջ������
    '    CPASSWORD   PCHAR   �ֿ��˿�����
    '�����⣬���ݻ��������ô��ȡ.
    StrInput = g�������_�ϳ�����.��������
    StrInput = StrInput & vbTab & g�������_�ϳ�����.����
    If ҵ������_�ϳ�����(��ȡ�ʻ����_����, StrInput, strOutput) = False Then Exit Function
    g�������_�ϳ�����.�ʻ���� = Val(strOutput)
    
    ��ȡ�α���Ա��Ϣ = True
    Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function
Private Function Get����(ByVal strDate As String) As Integer
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select (sysdate-to_date('" & strDate & "','yyyy-mm-dd'))/365 as ���� from dual "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If Not rsTemp.EOF Then
        Get���� = Int(Nvl(rsTemp!����, 0))
        Exit Function
    End If
    Exit Function
errHand:
End Function
Private Sub ClearData()
    Dim i As Long
    '��������Ϣ
    With g�������_�ϳ�����
        .ҽ������ = ""
        .ҽ��֤�� = ""
        .��¼�� = ""
        .���� = ""
        .���֤���� = ""
        .��λ���� = ""
        .�Ա� = ""
        .�������� = ""
        .���� = 0
   
        .�������� = ""
        .ְ�񼶱� = ""
        .ְ�Ƽ��� = ""
        .��Ա���� = ""
        .��ؾ�ס��־ = ""
        .��λID = ""
        .���� = ""
        .סԺ���� = ""
        .����ҽ�Ʊ�־ = ""
        .����ҽ�Ʊ�־ = ""
        .����Ա��־ = ""
        .��������״̬ = ""
        .�������״̬ = ""
        .����Ա����״̬ = ""
        .����סԺ���� = ""
        .�����ѱ������ = ""
        .�ɷ����� = ""
        .��ȡʱ�� = ""
        .סԺ��¼�� = ""
        
        .strסԺ��Ϣ = ""
    End With
    For i = 0 To lblEdit.UBound
        lblEdit(i).Caption = ""
    Next
End Sub
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If ��ȡ�α���Ա��Ϣ_סԺ = False Then
        cmdȷ��.Enabled = False
        Call ClearData
        Exit Sub
    End If
    Call LoadCtrlData
    cmdȷ��.Enabled = True
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit, KeyAscii, m�ı�ʽ
End Sub
Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If ��ȡ�α���Ա��Ϣ = False Then
        cmdȷ��.Enabled = False
        Call ClearData
        Exit Sub
    End If
    Call LoadCtrlData
    cmdȷ��.Enabled = True
    zlCommFun.PressKey vbKeyTab
End Sub

