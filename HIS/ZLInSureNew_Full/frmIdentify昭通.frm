VERSION 5.00
Begin VB.Form frmIdentify��ͨ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdRead 
      Caption         =   "ˢ��(&R)"
      Height          =   400
      Left            =   3030
      TabIndex        =   17
      Top             =   2130
      Width           =   1100
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   6
      Left            =   1050
      TabIndex        =   15
      Top             =   1440
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   1050
      TabIndex        =   0
      Top             =   120
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   1050
      TabIndex        =   8
      Top             =   560
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   4260
      TabIndex        =   7
      Top             =   560
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   4
      Left            =   1050
      TabIndex        =   6
      Top             =   1000
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   5
      Left            =   4260
      TabIndex        =   5
      Top             =   1000
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4260
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   2085
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -98
      TabIndex        =   4
      Top             =   1890
      Width           =   6810
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   400
      Left            =   4140
      TabIndex        =   2
      Top             =   2130
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   5245
      TabIndex        =   3
      Top             =   2130
      Width           =   1100
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "�ʻ����"
      Height          =   180
      Index           =   2
      Left            =   210
      TabIndex        =   16
      Top             =   1530
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Index           =   1
      Left            =   210
      TabIndex        =   14
      Top             =   210
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "ҽ �� ��"
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   13
      Top             =   650
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Index           =   3
      Left            =   3420
      TabIndex        =   12
      Top             =   650
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Index           =   4
      Left            =   210
      TabIndex        =   11
      Top             =   1090
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "���֤��"
      Height          =   180
      Index           =   5
      Left            =   3420
      TabIndex        =   10
      Top             =   1090
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Index           =   10
      Left            =   3420
      TabIndex        =   9
      Top             =   210
      Width           =   720
   End
End
Attribute VB_Name = "frmIdentify��ͨ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte, blnEOF As Boolean
Public mstrPatient As String, mstrOther As String, mstr���� As String, mstr���� As String

Public Function GetPatient(bytType As Byte) As String
    '������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    mbytType = bytType
    Me.Show vbModal
    GetPatient = mstrPatient & mstrOther
End Function

Private Sub cmdCancel_Click()
    mstrPatient = "": mstrOther = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If txtInfo(2).Text = "" Or txtInfo(3).Text = "" Then Exit Sub
    mstrOther = "": mstrPatient = ""
    mstrPatient = txtInfo(0).Text & ";"                                 '0 ����
    mstrPatient = mstrPatient & txtInfo(2).Text & ";"                   '1 ҽ���ʺ�
    mstrPatient = mstrPatient & txtInfo(1).Text & ";"                   '2 ����
    mstrPatient = mstrPatient & txtInfo(3).Text & ";"                   '3 ����
    mstrPatient = mstrPatient & txtInfo(4).Text & ";"                   '4 �Ա�
    mstrPatient = mstrPatient & ";"                                     '5 ��������
    mstrPatient = mstrPatient & txtInfo(5).Text & ";"                   '6 ���֤
    mstrPatient = mstrPatient & ";"                                     '7 ��λ����/����
    
    mstrOther = mstrOther & ";"                                         '8 ҽ����������(����)
    mstrOther = mstrOther & ";"                                         '9 ˳���
    mstrOther = mstrOther & ";"                                         '10 ���
    mstrOther = mstrOther & txtInfo(6).Text & ";"                       '11 ���
    mstrOther = mstrOther & ";"                                         '12 ��ǰ״̬
    mstrOther = mstrOther & ";"                                         '13 ����ID
    mstrOther = mstrOther & ";"                                         '14 ��ְ״̬
    mstrOther = mstrOther & ";"                                         '15 ����֤��
    mstrOther = mstrOther & ";"                                         '16 �����
    mstrOther = mstrOther & ";"                                         '17 �Ҷȼ�
    mstrOther = mstrOther & txtInfo(6).Text & ";"                       '18 �ʻ������ۼ�
    mstrOther = mstrOther & ";"                                         '19 �ʻ�֧���ۼ�
    mstrOther = mstrOther & ";"                                         '20 ����ͳ���ۼ�
    mstrOther = mstrOther & ";"                                         '21 ͳ�ﱨ���ۼ�
    mstrOther = mstrOther & ";"                                         '22 סԺ�����ۼ�
    mstrOther = mstrOther & ";"                                         '23 �������
    mstrOther = mstrOther & ";"                                         '24 ��������
    mstrOther = mstrOther & ";"                                         '25 �����ۼ�
    mstrOther = mstrOther & ";"                                         '26 ����ͳ���޶�
    
    Unload Me
End Sub

Private Sub cmdRead_Click()
    Dim intPort As Integer
    intPort = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��ǰʹ�õĴ���", 1)
    Me.txtInfo(0).Text = frmConn��ͨ.readCard(intPort)
    If Me.txtInfo(0).Text <> "" Then
        Me.txtInfo(1).Text = frmConn��ͨ.readPassword(intPort)
        Call readInfo
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    txtInfo(0).SetFocus
    cmdOK.Enabled = False
End Sub

Private Sub Timer1_Timer()
    blnEOF = True
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    txtInfo(Index).SelStart = 0
    txtInfo(Index).SelLength = Len(txtInfo(Index).Text)
End Sub

Private Function readInfo() As Boolean
    Dim strPara As String, strReturn() As String
    strPara = txtInfo(0).Text & vbTab & IIf(txtInfo(1).Text = "", " ", txtInfo(1).Text)
    If frmConn��ͨ.Execute("I200", 0, strPara, "���ڶ�ȡ����ҽ����Ϣ......") = False Then
'        txtInfo(0).SetFocus
        Exit Function
    End If
    If frmConn��ͨ.Query(0, 1) = False Then Exit Function
    mstr���� = txtInfo(0).Text
    mstr���� = IIf(txtInfo(1).Text = "", " ", txtInfo(1).Text)
    cmdOK.Enabled = True
    strReturn = Split(Replace(frmConn��ͨ.strReturnInfo, " ", ""), vbTab)
    txtInfo(2).Text = strReturn(0)
    txtInfo(3).Text = strReturn(1)
    txtInfo(4).Text = IIf(strReturn(2) = 0, "Ů", "��")
    txtInfo(5).Text = strReturn(3)
    txtInfo(6).Text = strReturn(4)
    readInfo = True
    cmdOK.SetFocus
End Function
