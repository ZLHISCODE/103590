VERSION 5.00
Begin VB.Form frmIdentify������ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   4298
      TabIndex        =   13
      Top             =   225
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   1208
      TabIndex        =   12
      Top             =   660
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   4298
      TabIndex        =   11
      Top             =   660
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   1208
      TabIndex        =   10
      Top             =   1080
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   4
      Left            =   4298
      TabIndex        =   9
      Top             =   1080
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   5
      Left            =   1208
      TabIndex        =   8
      Top             =   1485
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   6
      Left            =   4298
      TabIndex        =   7
      Top             =   1485
      Width           =   2085
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   7
      Left            =   1208
      TabIndex        =   6
      Top             =   1905
      Width           =   5175
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   8
      Left            =   1208
      TabIndex        =   5
      Top             =   2310
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   4
      Top             =   2790
      Width           =   6810
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "����(&R)"
      Height          =   400
      Left            =   428
      TabIndex        =   3
      Top             =   3000
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   400
      Left            =   4178
      TabIndex        =   2
      Top             =   3000
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   5283
      TabIndex        =   1
      Top             =   3000
      Width           =   1100
   End
   Begin VB.TextBox txtPass 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1208
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   225
      Width           =   2085
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "ҽ������"
      Height          =   180
      Index           =   0
      Left            =   3540
      TabIndex        =   23
      Top             =   315
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "����״̬"
      Height          =   180
      Index           =   1
      Left            =   450
      TabIndex        =   22
      Top             =   750
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Index           =   3
      Left            =   3540
      TabIndex        =   21
      Top             =   750
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Index           =   4
      Left            =   450
      TabIndex        =   20
      Top             =   1170
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   5
      Left            =   3540
      TabIndex        =   19
      Top             =   1170
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��ϵ�绰"
      Height          =   180
      Index           =   6
      Left            =   450
      TabIndex        =   18
      Top             =   1575
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   7
      Left            =   3540
      TabIndex        =   17
      Top             =   1575
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��ͥסַ"
      Height          =   180
      Index           =   8
      Left            =   450
      TabIndex        =   16
      Top             =   1995
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "������λ"
      Height          =   180
      Index           =   9
      Left            =   450
      TabIndex        =   15
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "IC������"
      Height          =   180
      Index           =   10
      Left            =   450
      TabIndex        =   14
      Top             =   315
      Width           =   720
   End
End
Attribute VB_Name = "frmIdentify������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mbytType As Byte, mstrPatient As String, mstrOther As String, mintסԺ���� As Integer
Private strTransNO As String, cur֧���ۼ� As Currency, cur�����ۼ� As Currency, strPara As String, _
    strReturn As String, blnReadCard As Boolean
 
Public Function GetPatient(bytType As Byte) As String
'������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    mbytType = bytType
    Me.Show vbModal
    GetPatient = mstrPatient & mstrOther
End Function

Private Sub cmdCancel_Click()
    mstrPatient = ""
    mstrOther = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    '17-��������֧����18-סԺ����֧����19-����סԺ������20-������ã�21-סԺ���ã�22-�ʻ����
    '23-�μ�ͳ��֧�����ã�24-ͳ��֧�����ã�25-�μӴ�֧�����ã�26-��֧�����ã�27-�Ƿ�����α�����
    '28-�α����ޣ�29-ҽ��״̬(0����)
    Dim datCurr As Date
    If blnReadCard = False Then
        MsgBox "���ȶ���", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error Resume Next
    If UCase(txtInfo(4).Text) = "YYYY-MM-DD" Then
        txtInfo(4).Enabled = True
        MsgBox "��������ȷ�ĳ�������", vbInformation, gstrSysName
        txtInfo(4).SetFocus
        txtInfo(4).SelStart = 0
        txtInfo(4).SelLength = Len(txtInfo(4).Text)
        On Error GoTo 0
        Exit Sub
    Else
        datCurr = CDate(txtInfo(4).Text)
        If Err.Number <> 0 Then
            MsgBox "�밴��ʽ:yyyy-mm-dd������ȷ�ĳ�������", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    mstrOther = "": mstrPatient = ""
    strReturn = Me.Tag
    mstrPatient = txtInfo(0).Text & ";"                                 '0 ����
    mstrPatient = mstrPatient & txtInfo(0).Text & ";"                   '1 ҽ���ʺ�
    mstrPatient = mstrPatient & txtPass.Text & ";"                      '2 ����
    mstrPatient = mstrPatient & txtInfo(2).Text & ";"                   '3 ����
    mstrPatient = mstrPatient & txtInfo(3).Text & ";"                   '4 �Ա�
    mstrPatient = mstrPatient & txtInfo(4).Text & ";"                   '5 ��������
    mstrPatient = mstrPatient & ";"                                     '6 ���֤
    mstrPatient = mstrPatient & txtInfo(8).Text & ";"                   '7 ��λ����/����
        
    mstrOther = mstrOther & ";"                                         '8 ҽ����������(����)
    mstrOther = mstrOther & txtInfo(0).Tag & ";"                        '9 ˳���
    mstrOther = mstrOther & ";"                                         '10 ���
    mstrOther = mstrOther & Split(strReturn, ",")(22) & ";"             '11 ���
    mstrOther = mstrOther & ";"                                         '12 ��ǰ״̬
    mstrOther = mstrOther & ";"                                         '13 ����ID
    mstrOther = mstrOther & IIf(txtInfo(1).Text = "��ְ", "1", "3") & ";"
    mstrOther = mstrOther & ";"                                         '15 ����֤��
    mstrOther = mstrOther & ";"                                         '16 �����
    mstrOther = mstrOther & ";"                                         '17 �Ҷȼ�
    mstrOther = mstrOther & Split(strReturn, ",")(22) & ";"             '18 �ʻ������ۼ�
    mstrOther = mstrOther & ";"                                         '19 �ʻ�֧���ۼ�
    mstrOther = mstrOther & Split(strReturn, ",")(23) & ";"             '20 ����ͳ���ۼ�
    mstrOther = mstrOther & Split(strReturn, ",")(24) & ";"             '21 ͳ�ﱨ���ۼ�
    mstrOther = mstrOther & Split(strReturn, ",")(19) & ";"             '22 סԺ�����ۼ�
    mstrOther = mstrOther & ";"                                         '23 �������
    mstrOther = mstrOther & Split(strReturn, ",")(18) & ";"             '24 ��������
    mstrOther = mstrOther & ";"                                         '25 �����ۼ�
    mstrOther = mstrOther & ";"                                         '26 ����ͳ���޶�
    
    mintסԺ���� = CInt(Split(strReturn, ",")(19))
    
    Me.Hide
End Sub

Private Sub cmdRead_Click()
    Dim lngReturn As Long, strReturn As String, strErrInfo As String, strInfo() As String
    If Trim(txtPass.Text) = "" Then
        MsgBox "������IC������", vbInformation, "����"
        Exit Sub
    End If
    lngReturn = init_com(intCOM����)
    If lngReturn <> 0 Then
        MsgBox "��ʼ���˿ڴ���", vbInformation, "����"
        Exit Sub
    End If
    
    lngReturn = sele_card(43)
    If lngReturn <> 0 Then
        MsgBox "���忨���ʹ���", vbInformation, "����"
        GoTo powerOFF
    End If
    
    If power_on() <> 0 Then
        MsgBox "���ϵ����", vbInformation, "����"
        GoTo powerOFF
    End If
    
    strReturn = Space(129)
    lngReturn = rd_str(1, 0, 128, strReturn)
    If lngReturn <> 0 Then
        MsgBox "��ȡ����Ϣ����", vbInformation, "����"
        GoTo powerOFF
    End If
    
    On Error GoTo powerOFF
    strInfo = Split(Trim(strReturn), "@")
    txtInfo(0).Text = strInfo(2)
    For lngReturn = 1 To 8
        If InStr(strInfo(lngReturn + 3), Chr(0)) > 0 Then
            strInfo(lngReturn + 3) = Left(strInfo(lngReturn + 3), InStr(strInfo(lngReturn + 3), Chr(0)) - 1)
        End If
        txtInfo(lngReturn).Text = IIf(lngReturn <> 3, IIf(lngReturn <> 1, strInfo(lngReturn + 3), IIf(strInfo(lngReturn + 3) = "0", "����", "��ְ")), IIf(strInfo(lngReturn + 3) = "0", "��", "Ů"))
    Next
    txtInfo(0).Tag = strInfo(0)
    
    blnReadCard = True
    cmdOK.SetFocus

powerOFF:
    Call power_off
    Call close_com
End Sub

Private Sub txtPass_GotFocus()
    txtPass.SelStart = 0
    txtPass.SelLength = Len(txtPass.Text)
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(txtPass.Text) = "" Then
            txtPass_GotFocus
            Exit Sub
        End If
        cmdRead_Click
        If blnReadCard = True Then cmdOK.SetFocus
    End If
End Sub

Private Sub clsText()
    Dim iLoop As Long
    For iLoop = 0 To 8
        txtInfo(iLoop).Text = ""
    Next
End Sub


