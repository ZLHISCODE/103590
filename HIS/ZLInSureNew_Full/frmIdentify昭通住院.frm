VERSION 5.00
Begin VB.Form frmIdentify��ͨסԺ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   15
      Left            =   4230
      TabIndex        =   36
      Top             =   3480
      Width           =   2085
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "ˢ��(&R)"
      Height          =   400
      Left            =   3030
      TabIndex        =   35
      Top             =   4260
      Width           =   1100
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   120
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   14
      Left            =   4237
      TabIndex        =   31
      Top             =   3060
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   13
      Left            =   1027
      TabIndex        =   29
      Top             =   3135
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   12
      Left            =   4237
      TabIndex        =   27
      Top             =   2640
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   11
      Left            =   1027
      TabIndex        =   25
      Top             =   2700
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   10
      Left            =   4237
      TabIndex        =   23
      Top             =   2220
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   9
      Left            =   1027
      TabIndex        =   21
      Top             =   2265
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   8
      Left            =   4237
      TabIndex        =   19
      Top             =   1800
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   7
      Left            =   1027
      TabIndex        =   17
      Top             =   1830
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Index           =   6
      Left            =   1027
      TabIndex        =   15
      Top             =   3570
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   1027
      TabIndex        =   0
      Top             =   540
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   1027
      TabIndex        =   8
      Top             =   975
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   4237
      TabIndex        =   7
      Top             =   960
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   4
      Left            =   1027
      TabIndex        =   6
      Top             =   1410
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   5
      Left            =   4237
      TabIndex        =   5
      Top             =   1380
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4237
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   540
      Width           =   2085
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -98
      TabIndex        =   4
      Top             =   4020
      Width           =   6810
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   400
      Left            =   4147
      TabIndex        =   2
      Top             =   4260
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   5257
      TabIndex        =   3
      Top             =   4260
      Width           =   1100
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "�ɷ�ʱ��"
      Height          =   180
      Index           =   16
      Left            =   3390
      TabIndex        =   37
      Top             =   3570
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��Ժ״̬"
      Height          =   180
      Index           =   15
      Left            =   180
      TabIndex        =   33
      Top             =   210
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   14
      Left            =   3397
      TabIndex        =   32
      Top             =   3150
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   13
      Left            =   187
      TabIndex        =   30
      Top             =   3225
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��λ����"
      Height          =   180
      Index           =   12
      Left            =   3397
      TabIndex        =   28
      Top             =   2730
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��λ����"
      Height          =   180
      Index           =   11
      Left            =   187
      TabIndex        =   26
      Top             =   2790
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "�� �� ��"
      Height          =   180
      Index           =   9
      Left            =   3397
      TabIndex        =   24
      Top             =   2310
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "סԺ����"
      Height          =   180
      Index           =   8
      Left            =   187
      TabIndex        =   22
      Top             =   2355
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "����״̬"
      Height          =   180
      Index           =   7
      Left            =   3397
      TabIndex        =   20
      Top             =   1890
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Index           =   6
      Left            =   187
      TabIndex        =   18
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "�ʻ����"
      Height          =   180
      Index           =   2
      Left            =   187
      TabIndex        =   16
      Top             =   3660
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Index           =   1
      Left            =   187
      TabIndex        =   14
      Top             =   630
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "ҽ �� ��"
      Height          =   180
      Index           =   0
      Left            =   187
      TabIndex        =   13
      Top             =   1065
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Index           =   3
      Left            =   3397
      TabIndex        =   12
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Index           =   4
      Left            =   187
      TabIndex        =   11
      Top             =   1500
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "���֤��"
      Height          =   180
      Index           =   5
      Left            =   3397
      TabIndex        =   10
      Top             =   1470
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Index           =   10
      Left            =   3397
      TabIndex        =   9
      Top             =   630
      Width           =   720
   End
End
Attribute VB_Name = "frmIdentify��ͨסԺ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte, blnEOF As Boolean
Public mstrPatient As String, mstrOther As String, mstrState As String

Public Function GetPatient(bytType As Byte, strinState As String) As String
    '������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    mbytType = bytType
    Combo1.AddItem "��ͨסԺ"
    Combo1.AddItem "תԺ"
    Combo1.AddItem "��������"
    Combo1.ListIndex = 0
    Me.Show vbModal
    strinState = mstrState
    GetPatient = mstrPatient & mstrOther
End Function

Private Sub cmdCancel_Click()
    mstrPatient = "": mstrOther = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    mstrOther = "": mstrPatient = ""
    If txtInfo(2).Text = "" Or txtInfo(3).Text = "" Then Exit Sub
    mstrState = Combo1.ListIndex + 1
    mstrPatient = txtInfo(0).Text & ";"                                 '0 ����
    mstrPatient = mstrPatient & txtInfo(2).Text & ";"                   '1 ҽ���ʺ�
    mstrPatient = mstrPatient & txtInfo(1).Text & ";"                   '2 ����
    mstrPatient = mstrPatient & txtInfo(3).Text & ";"                   '3 ����
    mstrPatient = mstrPatient & txtInfo(4).Text & ";"                   '4 �Ա�
    mstrPatient = mstrPatient & txtInfo(0).Tag & ";"                    '5 ��������
    mstrPatient = mstrPatient & txtInfo(5).Text & ";"                   '6 ���֤
    mstrPatient = mstrPatient & txtInfo(11).Text & ";"                  '7 ��λ����/����
    
    mstrOther = mstrOther & txtInfo(13).Text & ";"                      '8 ҽ����������(����)
    mstrOther = mstrOther & ";"                                         '9 ˳���
    mstrOther = mstrOther & ";"                                         '10 ���
    mstrOther = mstrOther & txtInfo(6).Text & ";"                       '11 ���
    mstrOther = mstrOther & ";"                                         '12 ��ǰ״̬
    mstrOther = mstrOther & ";"                                         '13 ����ID
    mstrOther = mstrOther & txtInfo(8).Text & ";"                       '14 ��ְ״̬
    mstrOther = mstrOther & ";"                                         '15 ����֤��
    mstrOther = mstrOther & ";"                                         '16 �����
    mstrOther = mstrOther & ";"                                         '17 �Ҷȼ�
    mstrOther = mstrOther & txtInfo(6).Text & ";"                       '18 �ʻ������ۼ�
    mstrOther = mstrOther & ";"                                         '19 �ʻ�֧���ۼ�
    mstrOther = mstrOther & ";"                                         '20 ����ͳ���ۼ�
    mstrOther = mstrOther & ";"                                         '21 ͳ�ﱨ���ۼ�
    mstrOther = mstrOther & ";"                                         '22 סԺ�����ۼ�
    mstrOther = mstrOther & ";"                                         '23 �������
    mstrOther = mstrOther & txtInfo(10).Text & ";"                      '24 ��������
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
    strPara = txtInfo(0).Text & vbTab & IIf(txtInfo(1).Text = "", " ", txtInfo(1).Text) & vbTab & Combo1.ListIndex + 1
    If frmConn��ͨ.Execute("I300", 0, strPara, "���ڶ�ȡ����ҽ����Ϣ......") = False Then
'        txtInfo(0).SetFocus
        Exit Function
    End If
    If frmConn��ͨ.Query(0, 1) = False Then Exit Function
    cmdOK.Enabled = True
    strReturn = Split(Replace(frmConn��ͨ.strReturnInfo, " ", ""), vbTab)
    txtInfo(2).Text = strReturn(0)
    txtInfo(3).Text = strReturn(1)
    txtInfo(4).Text = IIf(strReturn(3) = 0, "Ů", "��")
    txtInfo(5).Text = strReturn(2)
    txtInfo(6).Text = strReturn(12)
    txtInfo(7).Text = strReturn(4)
    txtInfo(8).Text = IIf(strReturn(5) = 1, "��ְ", "����")
    txtInfo(9).Text = strReturn(6)
    txtInfo(10).Text = strReturn(7)
    txtInfo(11).Text = strReturn(8)
    txtInfo(12).Text = strReturn(9)
    txtInfo(13).Text = strReturn(10)
    txtInfo(14).Text = strReturn(11)
    txtInfo(0).Tag = Left(strReturn(13), 4) & "-" & Mid(strReturn(13), 5, 2) & "-" & Right(strReturn(13), 2)
    txtInfo(15).Text = strReturn(14)
    cmdOK.SetFocus
    readInfo = True
End Function
