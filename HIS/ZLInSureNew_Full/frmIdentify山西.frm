VERSION 5.00
Begin VB.Form frmIdentifyɽ�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   5760
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7215
   Icon            =   "frmIdentifyɽ��.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd������ 
      Caption         =   "�޸�����(&P)"
      Height          =   375
      Left            =   540
      TabIndex        =   56
      Top             =   5220
      Width           =   1215
   End
   Begin VB.CommandButton cmd������Ϣ 
      Caption         =   "��"
      Height          =   300
      Left            =   6510
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   4785
      Width           =   285
   End
   Begin VB.TextBox txtDiseaseName 
      Height          =   300
      Left            =   1260
      TabIndex        =   3
      Top             =   4785
      Width           =   5220
   End
   Begin VB.ComboBox cmbҽ����� 
      Height          =   300
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4365
      Width           =   1950
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   5220
      Width           =   1215
   End
   Begin VB.TextBox txtPin 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4095
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4335
      Width           =   1440
   End
   Begin VB.CommandButton cmdReadCar 
      Caption         =   "����(&R)"
      Height          =   345
      Left            =   5730
      TabIndex        =   2
      Top             =   4320
      Width           =   960
   End
   Begin VB.Frame Frame2 
      Caption         =   "�ʻ�������Ϣ"
      Height          =   1905
      Left            =   270
      TabIndex        =   7
      Top             =   2340
      Width           =   6645
      Begin VB.TextBox txtAcc00 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1425
         TabIndex        =   55
         Top             =   255
         Width           =   1200
      End
      Begin VB.TextBox txtAcc03 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1920
         TabIndex        =   54
         Top             =   562
         Width           =   1200
      End
      Begin VB.TextBox txtAcc05 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1890
         TabIndex        =   53
         Top             =   869
         Width           =   1200
      End
      Begin VB.TextBox txtAcc07 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1890
         TabIndex        =   52
         Top             =   1176
         Width           =   1200
      End
      Begin VB.TextBox txtAcc09 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1890
         TabIndex        =   51
         Top             =   1485
         Width           =   1200
      End
      Begin VB.TextBox txtAcc10 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5235
         TabIndex        =   47
         Top             =   1485
         Width           =   1200
      End
      Begin VB.TextBox txtAcc08 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5235
         TabIndex        =   45
         Top             =   1170
         Width           =   1200
      End
      Begin VB.TextBox txtAcc06 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5235
         TabIndex        =   44
         Top             =   870
         Width           =   1200
      End
      Begin VB.TextBox txtAcc04 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5235
         TabIndex        =   43
         Top             =   555
         Width           =   1200
      End
      Begin VB.TextBox txtAcc02 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5805
         TabIndex        =   42
         Top             =   255
         Width           =   630
      End
      Begin VB.TextBox txtAcc01 
         Enabled         =   0   'False
         Height          =   270
         Left            =   3660
         TabIndex        =   41
         Top             =   255
         Width           =   660
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "���깫��Ա����֧���ۼ�"
         Height          =   240
         Left            =   3120
         TabIndex        =   40
         Top             =   1515
         Width           =   2070
      End
      Begin VB.Label Label21 
         Caption         =   "�����ֽ�֧���ۼ�"
         Height          =   240
         Left            =   210
         TabIndex        =   39
         Top             =   1515
         Width           =   1710
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "����ͳ��֧���ۼ�"
         Height          =   240
         Left            =   3480
         TabIndex        =   38
         Top             =   1206
         Width           =   1710
      End
      Begin VB.Label Label19 
         Caption         =   "�����ʻ�֧���ۼ�"
         Height          =   240
         Left            =   210
         TabIndex        =   37
         Top             =   1206
         Width           =   1710
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "�������ͳ���ۼ�"
         Height          =   240
         Left            =   3480
         TabIndex        =   36
         Top             =   899
         Width           =   1710
      End
      Begin VB.Label Label17 
         Caption         =   "���������ۼ�"
         Height          =   240
         Left            =   210
         TabIndex        =   35
         Top             =   899
         Width           =   1710
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "�����Է��ۼ�"
         Height          =   240
         Left            =   3480
         TabIndex        =   34
         Top             =   592
         Width           =   1710
      End
      Begin VB.Label Label15 
         Caption         =   "�����ܷ���֧���ۼ�"
         Height          =   240
         Left            =   210
         TabIndex        =   33
         Top             =   592
         Width           =   1710
      End
      Begin VB.Label Label14 
         Caption         =   "����סԺ����"
         Height          =   240
         Left            =   4530
         TabIndex        =   32
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "�ʻ����"
         Height          =   240
         Left            =   2790
         TabIndex        =   31
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "�ʻ�������"
         Height          =   240
         Left            =   210
         TabIndex        =   30
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "���˻�����Ϣ"
      Height          =   1995
      Left            =   270
      TabIndex        =   6
      Top             =   255
      Width           =   6645
      Begin VB.TextBox txtEmp10 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4605
         TabIndex        =   18
         Top             =   1500
         Width           =   1860
      End
      Begin VB.TextBox txtEmp09 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1500
         TabIndex        =   17
         Top             =   1500
         Width           =   1560
      End
      Begin VB.TextBox txtEmp08 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4605
         TabIndex        =   16
         Top             =   1185
         Width           =   1875
      End
      Begin VB.TextBox txtEmp07 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1110
         TabIndex        =   15
         Top             =   1185
         Width           =   1935
      End
      Begin VB.TextBox txtEmp06 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4605
         TabIndex        =   14
         Top             =   870
         Width           =   1890
      End
      Begin VB.TextBox txtEmp05 
         Enabled         =   0   'False
         Height          =   270
         Left            =   780
         TabIndex        =   13
         Top             =   870
         Width           =   1830
      End
      Begin VB.TextBox txtEmp04 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5340
         TabIndex        =   12
         Top             =   555
         Width           =   1155
      End
      Begin VB.TextBox txtEmp03 
         Enabled         =   0   'False
         Height          =   270
         Left            =   3465
         TabIndex        =   11
         Top             =   555
         Width           =   480
      End
      Begin VB.TextBox txtEmp02 
         Enabled         =   0   'False
         Height          =   270
         Left            =   795
         TabIndex        =   10
         Top             =   555
         Width           =   1140
      End
      Begin VB.TextBox txtEmp01 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4665
         TabIndex        =   9
         Top             =   240
         Width           =   1830
      End
      Begin VB.TextBox txtEmp00 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1125
         TabIndex        =   8
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label Label11 
         Caption         =   "��Ժ״̬"
         Height          =   240
         Left            =   3795
         TabIndex        =   29
         Top             =   1530
         Width           =   810
      End
      Begin VB.Label Label10 
         Caption         =   "�չ���Ա��־"
         Height          =   240
         Left            =   240
         TabIndex        =   28
         Top             =   1530
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "ҽ����Ա���"
         Height          =   240
         Left            =   3420
         TabIndex        =   27
         Top             =   1215
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "��λ���"
         Height          =   240
         Left            =   240
         TabIndex        =   26
         Top             =   1215
         Width           =   720
      End
      Begin VB.Label Label7 
         Caption         =   "ҽ��֤��"
         Height          =   240
         Left            =   3810
         TabIndex        =   25
         Top             =   900
         Width           =   720
      End
      Begin VB.Label Label6 
         Caption         =   "����"
         Height          =   240
         Left            =   240
         TabIndex        =   24
         Top             =   900
         Width           =   720
      End
      Begin VB.Label Label5 
         Caption         =   "��������"
         Height          =   240
         Left            =   4500
         TabIndex        =   23
         Top             =   585
         Width           =   720
      End
      Begin VB.Label Label4 
         Caption         =   "�Ա�"
         Height          =   240
         Left            =   3030
         TabIndex        =   22
         Top             =   585
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "����"
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   585
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "���֤��"
         Height          =   240
         Left            =   3885
         TabIndex        =   20
         Top             =   270
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "���˱��"
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   270
         Width           =   795
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   5580
      TabIndex        =   5
      Top             =   5220
      Width           =   1215
   End
   Begin VB.Label Label25 
      Caption         =   "��Ժ����"
      Height          =   210
      Left            =   315
      TabIndex        =   49
      Top             =   4830
      Width           =   855
   End
   Begin VB.Label Label24 
      Caption         =   "ҽ�����"
      Height          =   225
      Left            =   330
      TabIndex        =   48
      Top             =   4410
      Width           =   810
   End
   Begin VB.Label Label23 
      Caption         =   "����"
      Height          =   270
      Left            =   3525
      TabIndex        =   46
      Top             =   4380
      Width           =   480
   End
End
Attribute VB_Name = "frmIdentifyɽ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mbytType As Byte            '0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�
Private mlng����ID As Long
Private mstrReturn As String

Function ��ݱ�ʶ(Optional bytType As Byte, Optional lng����ID As Long = 0) As String
    mbytType = bytType
    mlng����ID = lng����ID
    mstrReturn = "-1"
    Me.Show 1
    lng����ID = mlng����ID
    ��ݱ�ʶ = mstrReturn
End Function

Private Sub CancelButton_Click()
        mstrReturn = "-1"
    Unload Me
End Sub

Private Sub cmdReadCar_Click()
    Dim strҽ����Ա��� As String
    Dim str����Ա��־ As String
    Dim str�չ���Ա��־ As String
    
    If Len(Trim(txtPin.Text)) = 0 Then
        MsgBox "���������룡", vbInformation, gstrSysName
        Exit Sub
    Else
        If �������(Trim(txtPin.Text)) Then
            txtEmp00.Text = IIf(IsNull(g���˻�����Ϣ.���˱��00), "", g���˻�����Ϣ.���˱��00)
            txtEmp01.Text = IIf(IsNull(g���˻�����Ϣ.���֤��01), "", g���˻�����Ϣ.���֤��01)
            txtEmp02.Text = IIf(IsNull(g���˻�����Ϣ.����02), "", g���˻�����Ϣ.����02)
            txtEmp03.Text = IIf(IsNull(g���˻�����Ϣ.�Ա�03), "", g���˻�����Ϣ.�Ա�03)
            txtEmp04.Text = IIf(IsNull(g���˻�����Ϣ.��������04), "", g���˻�����Ϣ.��������04)
            txtEmp05.Text = IIf(IsNull(g���˻�����Ϣ.����05), "", g���˻�����Ϣ.����05)
            txtEmp06.Text = IIf(IsNull(g���˻�����Ϣ.ҽ��֤��06), "", g���˻�����Ϣ.ҽ��֤��06)
            txtEmp07.Text = IIf(IsNull(g���˻�����Ϣ.��λ���07), "", g���˻�����Ϣ.��λ���07)
            
            strҽ����Ա��� = IIf(IsNull(g���˻�����Ϣ.ҽ����Ա���08), "", g���˻�����Ϣ.ҽ����Ա���08)
            Select Case strҽ����Ա���
                Case 11
                    strҽ����Ա��� = "��ְ"
                Case 21
                    strҽ����Ա��� = "����"
                Case 33
                    strҽ����Ա��� = "�����Ҽ��˲о���"
                Case 91
                    strҽ����Ա��� = "������Ա"
            End Select
            txtEmp08.Text = strҽ����Ա���
            
            str����Ա��־ = IIf(IsNull(g���˻�����Ϣ.����Ա��־09), "", g���˻�����Ϣ.����Ա��־09)
            txtEmp09.Text = IIf(str����Ա��־ = 0, "��", "��")
            
            str�չ���Ա��־ = IIf(IsNull(g���˻�����Ϣ.�չ���Ա��־10), "", g���˻�����Ϣ.�չ���Ա��־10)
            txtEmp10.Text = IIf(str�չ���Ա��־ = 0, "��", "��")
            
            txtAcc00.Text = IIf(IsNull(g�ʻ�������Ϣ.�ʻ�������00), "0.00", g�ʻ�������Ϣ.�ʻ�������00)
            txtAcc01.Text = IIf(IsNull(g�ʻ�������Ϣ.�ʻ����01), "", g�ʻ�������Ϣ.�ʻ����01)
            txtAcc02.Text = IIf(IsNull(g�ʻ�������Ϣ.����סԺ����02), "0", g�ʻ�������Ϣ.����סԺ����02)
            txtAcc03.Text = IIf(IsNull(g�ʻ�������Ϣ.�����ܷ���֧���ۼ�03), "0.00", g�ʻ�������Ϣ.�����ܷ���֧���ۼ�03)
            txtAcc04.Text = IIf(IsNull(g�ʻ�������Ϣ.�����Է��ۼ�05), "0.00", g�ʻ�������Ϣ.�����Է��ۼ�05)
            txtAcc05.Text = IIf(IsNull(g�ʻ�������Ϣ.���������ۼ�06), "0.00", g�ʻ�������Ϣ.���������ۼ�06)
            txtAcc06.Text = IIf(IsNull(g�ʻ�������Ϣ.�������ͳ���ۼ�07), "0.00", g�ʻ�������Ϣ.�������ͳ���ۼ�07)
            txtAcc07.Text = IIf(IsNull(g�ʻ�������Ϣ.�����ʻ�֧���ۼ�08), "0.00", g�ʻ�������Ϣ.�����ʻ�֧���ۼ�08)
            txtAcc08.Text = IIf(IsNull(g�ʻ�������Ϣ.����ͳ��֧���ۼ�10), "0.00", g�ʻ�������Ϣ.����ͳ��֧���ۼ�10)
            txtAcc09.Text = IIf(IsNull(g�ʻ�������Ϣ.�����ֽ�֧���ۼ�11), "0.00", g�ʻ�������Ϣ.�����ֽ�֧���ۼ�11)
            txtAcc10.Text = IIf(IsNull(g�ʻ�������Ϣ.���깫��Ա����֧���ۼ�13), "0.00", g�ʻ�������Ϣ.���깫��Ա����֧���ۼ�13)
            Call �ύ_ɽ��
            '��ʱ��ȡ����Ϣ����ʾ��
            '��ʱ�ύһ�Σ������û��ڽ����ϳ�ʱ��ͣ�������������ȷ��ʱ��ҽ�������Ѿ�����״̬����ɲ�һ��
            'ȷ��ʱ����Ҫ�ٵ�һ�ζ������
             SendKeys ("{Tab}")
        Else
            mstrReturn = "-1"
        End If
    End If
    
End Sub


Private Sub cmd������_Click()
    Dim strOldPass As String, strNewPass As String
    
    strOldPass = Trim(txtPin.Text)
    strNewPass = ""
    
    strNewPass = frm�޸�����.ChangePassword("", strOldPass)
    
    If Nvl(strNewPass) = "" Then Exit Sub
    
    If �޸�����_ɽ��(strOldPass, strNewPass) Then
        txtPin.Text = strNewPass
    End If
End Sub

Private Sub Form_Load()

    '��ʼ��ҽ�����
    If mbytType = 0 Then
        '11      ��ͨ����
        '12  ��������
        '14  ����ҩ�깺ҩ
        '17  ���Ｑ��
        
       cmbҽ�����.AddItem "��ͨ����"
       cmbҽ�����.ItemData(cmbҽ�����.NewIndex) = 11
       
       cmbҽ�����.AddItem "��������"
       cmbҽ�����.ItemData(cmbҽ�����.NewIndex) = 12
       
       cmbҽ�����.AddItem "���Ｑ��"
       cmbҽ�����.ItemData(cmbҽ�����.NewIndex) = 17
       
       '��Ĭ��ֵ
       cmbҽ�����.Text = "��ͨ����"
       cmbҽ�����.Tag = 11
    End If
    
    If mbytType = 1 Then
        '21  ��ͨסԺ
        '23  ת��סԺ
        '24  תԺסԺ
        '26  ���Ｑ��ת��סԺ
       cmbҽ�����.AddItem "��ͨסԺ"
       cmbҽ�����.ItemData(cmbҽ�����.NewIndex) = 21
       
       cmbҽ�����.AddItem "תԺסԺ"
       cmbҽ�����.ItemData(cmbҽ�����.NewIndex) = 24
       
       cmbҽ�����.AddItem "���Ｑ��ת��סԺ"
       cmbҽ�����.ItemData(cmbҽ�����.NewIndex) = 26
       '��Ĭ��ֵ
       cmbҽ�����.Text = "��ͨסԺ"
       cmbҽ�����.Tag = 21
    End If
    
End Sub

Private Sub cmbҽ�����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub


Private Sub OKButton_Click()

Dim strEmpInfo As String, straccinfo As String  ''��Ÿ��˻�����Ϣ���ʻ���Ϣ
Dim strTmpSQL As String '��ʱSQL���
Dim rsTmp As New ADODB.Recordset  '��ʱ��¼��
Dim cur����ID  As Currency  '��currency�����׳��ֽ���δ֪����.
Dim str���ּ��� As String

'����ѡ��û�У�Ҫ�ж�
If txtDiseaseName.Tag = "" Then
     MsgBox "��ѡ���֣�", vbInformation, gstrSysName
     mstrReturn = "-1"
     txtDiseaseName.SetFocus
     Exit Sub
End If
  
'�����ж�,�ٴζ�����
If �������(Trim(txtPin.Text)) Then

    '���没����Ϣ�����ղ��ֱ���
      '�жϿ�����û���������,���У���ֱ��ȡ�ò���ID
    strTmpSQL = "select * from ���ղ��� where ����=" & TYPE_ɽ�� & _
                                         " and ����='" & txtDiseaseName.Tag & "'"
    Call OpenRecordset(rsTmp, "�鲡��ID", strTmpSQL)
    If rsTmp.EOF Then
        strTmpSQL = "select ���ղ���_ID.NextVal as ID from Dual "
        Call OpenRecordset(rsTmp, "ȡ����ID", strTmpSQL)
        cur����ID = 1
        If Not rsTmp.EOF Then cur����ID = rsTmp!ID
        
        strTmpSQL = "select zlspellcode('" & txtDiseaseName.Text & "') as ���� from dual"
        Call OpenRecordset(rsTmp, "ȡ���ּ���", strTmpSQL)
        str���ּ��� = rsTmp!����
        
        strTmpSQL = "zl_���ղ���_insert(" & cur����ID & "," & TYPE_ɽ�� & ",'" & _
                                         txtDiseaseName.Tag & "','" & _
                                         txtDiseaseName.Text & "','" & _
                                         str���ּ��� & "',1,NULL,NULL)"
        gcnOracle.Execute strTmpSQL, , adCmdStoredProc
        
        
        rsTmp.Close
        Set rsTmp = Nothing
    Else
       cur����ID = rsTmp!ID
    End If
   
    '�����ַ���
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    
    strEmpInfo = g���˻�����Ϣ.����05                               '0����
    strEmpInfo = strEmpInfo & ";" & g���˻�����Ϣ.���˱��00             '1ҽ����
    strEmpInfo = strEmpInfo & ";" & txtPin.Text               '2����
    strEmpInfo = strEmpInfo & ";" & g���˻�����Ϣ.����02               '3����
    strEmpInfo = strEmpInfo & ";" & g���˻�����Ϣ.�Ա�03               '4�Ա�
    strEmpInfo = strEmpInfo & ";" & Mid(g���˻�����Ϣ.��������04, 1, 4) & "-" & Mid(g���˻�����Ϣ.��������04, 5, 2) & "-" & Mid(g���˻�����Ϣ.��������04, 7, 2)        '5��������
    strEmpInfo = strEmpInfo & ";" & g���˻�����Ϣ.���֤��01           '6���֤
    strEmpInfo = strEmpInfo & ";" & g���˻�����Ϣ.��λ���07         '7.��λ����(����)
    
    straccinfo = ";0"                                          '8.���Ĵ���
    straccinfo = straccinfo & ";"                    '9.˳���
    straccinfo = straccinfo & ";" & g���˻�����Ϣ.ҽ����Ա���08           '10��Ա���
    straccinfo = straccinfo & ";" & g�ʻ�������Ϣ.�ʻ�������00      '11�ʻ����
    straccinfo = straccinfo & ";0" ' & g���˻�����Ϣ.��Ժ״̬16                             '12��ǰ״̬
    straccinfo = straccinfo & ";" & cur����ID                  '13����ID
    straccinfo = straccinfo & ";1"                            '14��ְ(1,2,3)
    straccinfo = straccinfo & ";"                             '15����֤��
    straccinfo = straccinfo & ";"                             '16�����
    straccinfo = straccinfo & ";1"                            '17�Ҷȼ�
    straccinfo = straccinfo & ";" & g�ʻ�������Ϣ.�ʻ�������00      '18�ʻ������ۼ�
    straccinfo = straccinfo & ";0"                              '19�ʻ�֧���ۼ�
    straccinfo = straccinfo & ";0"                            '20���깤���ܶ�
    straccinfo = straccinfo & ";"      '21
    straccinfo = straccinfo & ";" & g�ʻ�������Ϣ.����סԺ����02      '22סԺ�����ۼ�
    
    mlng����ID = BuildPatiInfo(0, strEmpInfo & straccinfo, mlng����ID, TYPE_ɽ��)
    
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_ɽ�� & ",'�������','''" & cmbҽ�����.ItemData(cmbҽ�����.ListIndex) & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "Ӧ�����")
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strEmpInfo & ";" & mlng����ID & straccinfo
    End If
    Unload Me
Else
    mstrReturn = "-1"
End If


End Sub

Private Sub txtDiseaseName_KeyPress(KeyAscii As Integer)
  ''������ѡ����
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtDiseaseName.Text) = "" Then Exit Sub
    Call ����ѡ��
End Sub

Private Sub ����ѡ��(Optional strLoad As String = 1)
    Dim rsTmp As ADODB.Recordset
    Dim strTmpSQL As String
    If strLoad = 1 Then
        strTmpSQL = "select rownum as ID,aka120  ���ֱ���,aka121 ��������,aka066 ������,aae035 ������� from ka06" & _
                    " where aka120 like '%" & Trim(txtDiseaseName.Text) & "%' or aka121 like '%" & Trim(txtDiseaseName.Text) & "%' or Upper(aka066) like '%" & UCase(Trim(txtDiseaseName.Text)) & "%'"
    Else
        strTmpSQL = "select rownum as ID,aka120  ���ֱ���,aka121 ��������,aka066 ������,aae035 ������� from ka06"
    End If
    Set rsTmp = frmPubSel.ShowSelect(Me, strTmpSQL, 0, "����", True, , , , , gcnSxDr)
    If rsTmp Is Nothing Then Exit Sub
    txtDiseaseName.Text = rsTmp!��������
    txtDiseaseName.Tag = rsTmp!���ֱ���
    OKButton.SetFocus
End Sub

Private Sub cmd������Ϣ_Click()
    Call ����ѡ��(0)
End Sub

Private Sub txtPin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{Tab}")
End Sub


