VERSION 5.00
Begin VB.Form frmIdentify���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   3615
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6000
   Icon            =   "frmIdentify����.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt��λ���� 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1215
      TabIndex        =   26
      Top             =   1470
      Width           =   4350
   End
   Begin VB.CommandButton cmdˢ�� 
      Caption         =   "ˢ��(&R)"
      Height          =   375
      Left            =   465
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3465
      TabIndex        =   24
      Top             =   2280
      Width           =   1185
   End
   Begin VB.TextBox txt��Ա״̬ 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3465
      TabIndex        =   22
      Top             =   1065
      Width           =   1185
   End
   Begin VB.TextBox txt��Ա��� 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1215
      TabIndex        =   20
      Top             =   1065
      Width           =   1185
   End
   Begin VB.TextBox txt�ʻ���� 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1215
      TabIndex        =   18
      Top             =   2280
      Width           =   1185
   End
   Begin VB.TextBox txt����״̬ 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4395
      TabIndex        =   16
      Top             =   255
      Width           =   1185
   End
   Begin VB.TextBox txt�������� 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4395
      TabIndex        =   14
      Top             =   660
      Width           =   1185
   End
   Begin VB.TextBox txt�������� 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3465
      TabIndex        =   12
      Top             =   1875
      Width           =   1185
   End
   Begin VB.TextBox txt�ʻ�״̬ 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1215
      TabIndex        =   10
      Top             =   1875
      Width           =   1185
   End
   Begin VB.TextBox txt�Ա� 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2940
      TabIndex        =   6
      Top             =   660
      Width           =   360
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1035
      TabIndex        =   5
      Top             =   660
      Width           =   1185
   End
   Begin VB.TextBox txt�籣�� 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1035
      TabIndex        =   3
      Top             =   255
      Width           =   2250
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   4425
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   2985
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "���ߣ�"
      Height          =   225
      Left            =   2580
      TabIndex        =   25
      Top             =   2340
      Width           =   930
   End
   Begin VB.Label Label11 
      Caption         =   "��Ա״̬��"
      Height          =   225
      Left            =   2520
      TabIndex        =   23
      Top             =   1140
      Width           =   930
   End
   Begin VB.Label Label10 
      Caption         =   "��Ա���"
      Height          =   225
      Left            =   300
      TabIndex        =   21
      Top             =   1140
      Width           =   930
   End
   Begin VB.Label Label9 
      Caption         =   "�ʻ���"
      Height          =   225
      Left            =   300
      TabIndex        =   19
      Top             =   2340
      Width           =   930
   End
   Begin VB.Label Label8 
      Caption         =   "����״̬��"
      Height          =   255
      Left            =   3465
      TabIndex        =   17
      Top             =   270
      Width           =   930
   End
   Begin VB.Label Label7 
      Caption         =   "�������ڣ�"
      Height          =   225
      Left            =   3480
      TabIndex        =   15
      Top             =   720
      Width           =   930
   End
   Begin VB.Label Label6 
      Caption         =   "�������ͣ�"
      Height          =   225
      Left            =   2565
      TabIndex        =   13
      Top             =   1950
      Width           =   930
   End
   Begin VB.Label Label5 
      Caption         =   "�ʻ�״̬��"
      Height          =   225
      Left            =   300
      TabIndex        =   11
      Top             =   1950
      Width           =   930
   End
   Begin VB.Label Label4 
      Caption         =   "��λ���ƣ�"
      Height          =   255
      Left            =   300
      TabIndex        =   9
      Top             =   1530
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "�Ա�"
      Height          =   255
      Left            =   2355
      TabIndex        =   8
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label2 
      Caption         =   "������"
      Height          =   255
      Left            =   300
      TabIndex        =   7
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "�籣�ţ�"
      Height          =   255
      Left            =   300
      TabIndex        =   4
      Top             =   300
      Width           =   780
   End
End
Attribute VB_Name = "frmIdentify����"
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
    mstrReturn = "99"
    
    Me.Show 1
    lng����ID = mlng����ID
    ��ݱ�ʶ = mstrReturn
End Function

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub cmdˢ��_Click()
  Dim strSBH As String * 20 '�籣��
 ' Dim strSbh1 As String '����ȡ����Ա��Ϣʱ���ݵ��籣��,����25λ
  
  Dim net As String * 1 '����״̬
  Dim rylx As String * 1 '��Ա����
  Dim zhzt As String * 3 '�ʻ�״̬
  Dim tmp As Double '����ֵ
  Dim Zhye As Double '�ʻ����
  
  Dim fhz As String * 1000 '����ֵ

 Dim zffs As String * 1 '֧����ʽ
 Dim tcqfx As Double 'ͳ������
 
net = Space(1)
strSBH = Space(20)
rylx = Space(1)
zhzt = Space(3)
zffs = Space(1)

If mbytType = 0 Then 'begin ����

Zhye = 0

'beging 1ˢ��
  tmp = mzcsh(strSBH, net, rylx, zhzt, Zhye)
  Call WriteBusinessLOG("mzcsh", "sbh,net,rylx,zhzt,zhye", tmp & "," & strSBH & "," & net & "," & rylx & "," & zhzt & "," & Zhye)
  
  If tmp = 0 Then
    txt�籣��.Text = Mid(strSBH, 1, 18)
    
    txt�ʻ�״̬.Tag = Mid(zhzt, 1, 3)
    If txt�ʻ�״̬.Tag = "002" Then
        txt�ʻ�״̬.Text = "�ʻ�����"
    Else
        txt�ʻ�״̬.Text = "�ʻ�����"
    End If
    
    txt��������.Tag = Mid(rylx, 1, 1)
    Select Case txt��������.Tag
    Case 1
        txt��������.Text = "��ͨ����"
    Case 2
        txt��������.Text = "���°���"
    Case 3
        txt��������.Text = "��������"
    End Select
    
    txt����״̬.Tag = Mid(net, 1, 1)
    If txt����״̬.Tag = 1 Then
        txt����״̬.Text = "ͨ"
    Else
        txt����״̬.Text = "��ͨ"
    End If
    
    If Mid(rylx, 1, 1) <> 2 And Mid(zhzt, 1, 3) = "002" Then
        txt�ʻ����.Text = Val(Zhye)
    Else
        txt�ʻ����.Text = 0
    End If
    '2Ȼ��ȡ�û�����Ϣ
    fhz = Space(1000)
    'strSbh1 = Trim(txt�籣��.Text) & Space(25 - Len(Trim(txt�籣��.Text)))
    tmp = getyhxx_vb(txt�籣��.Text, fhz)
    Call WriteBusinessLOG("getyhxx", txt�籣��.Text & "," & Trim(fhz), tmp)

    If Trim(fhz) <> "" Then
        txt����.Text = Split(fhz, ",")(0)
        txt�Ա�.Text = Split(fhz, ",")(1)
        txt��λ����.Text = Split(fhz, ",")(2)
        txt��Ա���.Text = Split(fhz, ",")(3)
        txt��Ա״̬.Text = Split(fhz, ",")(4)
        txt��������.Text = Split(fhz, ",")(5)
    End If
    OKButton.Enabled = True
    SendKeys ("{Tab}")
  End If
  'end ˢ��
  If tmp = 1 Then MsgBox "������Ϣ�޸��籣��", vbInformation, gstrSysName
  If tmp = 2 Then MsgBox "���ػ�����Ϣ����Ҫ����", vbInformation, gstrSysName
  If tmp = 99 Then MsgBox "����", vbInformation, gstrSysName

End If 'end ����

If mbytType = 1 Then 'beging סԺ
    tmp = rycsh(strSBH, zhzt, zffs, net, Zhye, tcqfx)
    Call WriteBusinessLOG("zycsh", "sbh, zhzt, zffs, net, Zhye, tcqfx", tmp & "," & strSBH & "," & zhzt & "," & zffs & "," & net & "," & Zhye & "," & tcqfx)
    If tmp = 0 Then
        txt�籣��.Text = Mid(strSBH, 1, 18)
        
        txt�ʻ�״̬.Tag = Mid(zhzt, 1, 3)
        If txt�ʻ�״̬.Tag = "002" Then
            txt�ʻ�״̬.Text = "�ʻ�����"
        Else
            txt�ʻ�״̬.Text = "�ʻ�����"
        End If
            
        txt��������.Tag = Mid(zffs, 1, 1)
        If txt��������.Tag = 1 Then
            txt��������.Text = "�ʻ�������"
        Else
            txt��������.Text = "�����ʻ�"
        End If
        
        txt����״̬.Tag = Mid(net, 1, 1)
        If txt����״̬.Tag = 1 Then
            txt����״̬.Text = "ͨ"
        Else
            txt����״̬.Text = "��ͨ"
        End If
        
        txt����.Text = Val(tcqfx)
        
        If txt��������.Tag <> 1 And txt�ʻ�״̬.Tag = "002" Then
            txt�ʻ����.Text = Val(Zhye)
        Else
            txt�ʻ����.Text = 0
        End If
        
        '2Ȼ��ȡ�û�����Ϣ
        fhz = Space(1000)
        'strSbh1 = Trim(txt�籣��.Text) & Space(25 - Len(Trim(txt�籣��.Text)))
        tmp = getyhxx_vb(Trim(txt�籣��.Text), fhz)
        Call WriteBusinessLOG("getyhxx", Trim(txt�籣��.Text) & "," & Trim(fhz), tmp)
        
        If Trim(fhz) <> "" Then
            txt����.Text = Split(fhz, ",")(0)
            txt�Ա�.Text = Split(fhz, ",")(1)
            txt��λ����.Text = Split(fhz, ",")(2)
            txt��Ա���.Text = Split(fhz, ",")(3)
            txt��Ա״̬.Text = Split(fhz, ",")(4)
            txt��������.Text = Split(fhz, ",")(5)
        End If
        OKButton.Enabled = True
        SendKeys ("{Tab}")
    End If
  'end ˢ��
    If tmp = 1 Then MsgBox "������Ϣ�޸��籣��", vbInformation, gstrSysName
    If tmp = 2 Then MsgBox "���ػ�����Ϣ����Ҫ����", vbInformation, gstrSysName
    If tmp = 99 Then MsgBox "����", vbInformation, gstrSysName
End If 'end סԺ
End Sub

Private Sub Form_Load()
    '��ʼ���ؼ�
    If mbytType = 0 Then
        txt����.Visible = False
        Label12.Visible = False
        Label6.Caption = "�������ͣ�"
    End If
    
    If mbytType = 1 Then
        txt����.Visible = True
        Label12.Visible = True
        Label6.Caption = "֧����ʽ��"
    End If
    OKButton.Enabled = False
End Sub

Private Sub OKButton_Click()
    Dim strEmpInfo As String
    Dim straccinfo As String

    '�����ַ���
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    
    strEmpInfo = txt�籣��.Text                               '0����
    strEmpInfo = strEmpInfo & ";" & txt�籣��.Text              '1ҽ����
    strEmpInfo = strEmpInfo & ";" & txt����״̬.Tag             '2����  ��ҽ���б����������״̬
    strEmpInfo = strEmpInfo & ";" & txt����.Text                '3����
    strEmpInfo = strEmpInfo & ";" & txt�Ա�.Text                '4�Ա�
    strEmpInfo = strEmpInfo & ";" & txt��������.Text         '5��������
    strEmpInfo = strEmpInfo & ";"             '6���֤
    strEmpInfo = strEmpInfo & ";" & txt��λ����          '7.��λ����(����)
    
    straccinfo = ";0"                                          '8.���Ĵ���
    straccinfo = straccinfo & ";"                    '9.˳���
    straccinfo = straccinfo & ";" & txt��������.Tag             '10��Ա���
    straccinfo = straccinfo & ";" & Val(txt�ʻ����.Text)        '11�ʻ����
    straccinfo = straccinfo & ";" & txt��Ա״̬.Tag   ' & g���˻�����Ϣ.��Ժ״̬16                             '12��ǰ״̬
    straccinfo = straccinfo & ";"                   '13����ID
    straccinfo = straccinfo & ";1"                            '14��ְ(1,2,3)
    straccinfo = straccinfo & ";"                             '15����֤��
    straccinfo = straccinfo & ";"                             '16�����
    straccinfo = straccinfo & ";1"                            '17�Ҷȼ�
    straccinfo = straccinfo & ";0"       '18�ʻ������ۼ�
    straccinfo = straccinfo & ";0"                              '19�ʻ�֧���ۼ�
    straccinfo = straccinfo & ";0"                            '20���깤���ܶ�
    straccinfo = straccinfo & ";"      '21
    straccinfo = straccinfo & ";"       '22סԺ�����ۼ�
    
    mlng����ID = BuildPatiInfo(0, strEmpInfo & straccinfo, mlng����ID, TYPE_����)
    
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_���� & ",'�������','''" & txt��������.Tag & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "Ӧ�����")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_���� & ",'��Ա���','''" & txt�ʻ�״̬.Tag & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "�ʻ�״̬")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_���� & ",'�ʻ����','''" & txt�ʻ����.Text & "''')"
    Call zldatabase.ExecuteProcedure(gstrSQL, "�ʻ����")
    
    
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strEmpInfo & ";" & mlng����ID & straccinfo
    End If
    Unload Me

End Sub

