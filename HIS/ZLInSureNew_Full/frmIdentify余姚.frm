VERSION 5.00
Begin VB.Form frmIdentify��Ҧ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   4
      Left            =   1095
      TabIndex        =   0
      Top             =   135
      Width           =   2310
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   2060
      TabIndex        =   11
      Top             =   2595
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   400
      Left            =   770
      TabIndex        =   10
      Top             =   2595
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -30
      TabIndex        =   9
      Top             =   2400
      Width           =   3990
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   3
      Left            =   1095
      TabIndex        =   8
      Top             =   1920
      Width           =   2310
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   2
      Left            =   1095
      TabIndex        =   6
      Top             =   1470
      Width           =   2310
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   1
      Left            =   1095
      TabIndex        =   4
      Top             =   1035
      Width           =   2310
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   0
      Left            =   1095
      TabIndex        =   2
      Top             =   585
      Width           =   2310
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���￨��"
      Height          =   180
      Index           =   4
      Left            =   300
      TabIndex        =   12
      Top             =   225
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��λ���"
      Height          =   180
      Index           =   3
      Left            =   300
      TabIndex        =   7
      Top             =   2010
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   180
      Index           =   2
      Left            =   660
      TabIndex        =   5
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   1
      Left            =   660
      TabIndex        =   3
      Top             =   1125
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���˱��"
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   675
      Width           =   720
   End
End
Attribute VB_Name = "frmIdentify��Ҧ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte
Private mlng����ID As Long
Public mstrPatient As String, mstrOther As String

Public Function GetPatient(bytType As Byte, lng����ID As Long) As String
    '������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    mbytType = bytType
    mlng����ID = lng����ID
    Me.Show vbModal
'    gstrIC���� = Right(Space(18) & txtInfo(0).Text, 18) & _
'                 String(18, "0") & _
'                 Right(Space(20) & txtInfo(1).Text, 20) & _
'                 Right(Space(2) & txtInfo(2).Text, 2) & _
'                 String(56, "0") & _
'                 Right(Space(10) & txtInfo(3).Text, 10) & _
'                 String(2, "0") & String(126 + 85 + 146 + 116 * 6, "0")
                 
    gstrIC���� = Space(18 - LenB(StrConv(txtInfo(0).Text, vbFromUnicode))) & txtInfo(0).Text & _
                 String(18, "0") & _
                 Space(20 - LenB(StrConv(txtInfo(1).Text, vbFromUnicode))) & txtInfo(1).Text & _
                 Space(2 - LenB(StrConv(txtInfo(2).Text, vbFromUnicode))) & txtInfo(2).Text & _
                 String(56, "0") & _
                 Space(10 - LenB(StrConv(txtInfo(3).Text, vbFromUnicode))) & txtInfo(3).Text & _
                 String(2, "0") & String(126 + 85 + 146 + 116 * 6, "0")
    GetPatient = mstrPatient & mstrOther
    If GetPatient <> "" Then lng����ID = mlng����ID
End Function

Private Sub cmdCancel_Click()
    mstrPatient = "": mstrOther = ""
    Me.Hide
    'ȡ��
End Sub

Private Sub cmdOK_Click()
    'ȷ��
    If txtInfo(2).Text <> "��" And txtInfo(2).Text <> "Ů" Then
        MsgBox "�Ա������롰�С���Ů��", vbInformation, gstrSysName
        Exit Sub
    End If
    If Trim(txtInfo(0).Text) = "" Or Trim(txtInfo(1).Text) = "" Or Trim(txtInfo(2).Text) = "" Or Trim(txtInfo(3).Text) = "" Then
        MsgBox "�����������ĸ��������Ϣ", vbInformation, gstrSysName
        Exit Sub
    End If
    mstrOther = "": mstrPatient = ""
    
    mstrPatient = txtInfo(0).Text & ";"                                 '0 ����
    mstrPatient = mstrPatient & txtInfo(0).Text & ";"                   '1 ҽ���ʺ�
    mstrPatient = mstrPatient & ";"                                     '2 ����
    mstrPatient = mstrPatient & txtInfo(1).Text & ";"                   '3 ����
    mstrPatient = mstrPatient & txtInfo(2).Text & ";"                   '4 �Ա�
    mstrPatient = mstrPatient & ";"                                     '5 ��������
    mstrPatient = mstrPatient & ";"                                     '6 ���֤
    mstrPatient = mstrPatient & "(" & txtInfo(3).Text & ")" & ";"       '7 ��λ����/����
    
    mstrOther = mstrOther & gstrҽ���������� & ";"                      '8 ҽ����������(����)
    mstrOther = mstrOther & ";"                                         '9 ˳���
    mstrOther = mstrOther & ";"                                         '10 ���
    mstrOther = mstrOther & ";"                                         '11 ���
    mstrOther = mstrOther & ";"                                         '12 ��ǰ״̬
    mstrOther = mstrOther & ";"                                         '13 ����ID
    mstrOther = mstrOther & ";"                                         '14 ��ְ״̬
    mstrOther = mstrOther & ";"                                         '15 ����֤��
    mstrOther = mstrOther & ";"                                         '16 �����
    mstrOther = mstrOther & ";"                                         '17 �Ҷȼ�
    mstrOther = mstrOther & ";"                                         '18 �ʻ������ۼ�
    mstrOther = mstrOther & ";"                                         '19 �ʻ�֧���ۼ�
    mstrOther = mstrOther & ";"                                         '20 ����ͳ���ۼ�
    mstrOther = mstrOther & ";"                                         '21 ͳ�ﱨ���ۼ�
    mstrOther = mstrOther & ";"                                         '22 סԺ�����ۼ�
    mstrOther = mstrOther & ";"                                         '23 �������
    mstrOther = mstrOther & ";"                                         '24 ��������
    mstrOther = mstrOther & ";"                                         '25 �����ۼ�
    mstrOther = mstrOther & ";"                                         '26 ����ͳ���޶�
    
    Me.Hide
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    txtInfo(Index).SelStart = 0
    txtInfo(Index).SelLength = Len(txtInfo(Index).Text)
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, rs��Ϣ As New ADODB.Recordset
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Index = 4 Then
        If Trim(txtInfo(4).Text) <> "" Then
            gstrSQL = "Select * From ������Ϣ Where ���￨��=[1]"
            Set rs��Ϣ = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CStr(Trim(UCase(txtInfo(4).Text))))
            If Not rs��Ϣ.EOF Then
                mlng����ID = rs��Ϣ!����ID
                gstrSQL = "Select ���� From �����ʻ� Where ����ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs��Ϣ!����ID))
                If rsTemp.EOF Then
                    txtInfo(1).Text = rs��Ϣ!����
                    txtInfo(2).Text = rs��Ϣ!�Ա�
                Else
                    txtInfo(0).Text = rsTemp(0)
                    gstrSQL = "Select A.����,A.�Ա�,nvl(B.��λ����,'0') From ������Ϣ A,�����ʻ� B Where A.����ID=B.����ID And B.����=[1] And B.����=[2]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��Ҧ, CStr(Trim(txtInfo(0).Text)))
                    If Not rsTemp.EOF Then
                        txtInfo(1).Text = rsTemp(0)
                        txtInfo(2).Text = rsTemp(1)
                        txtInfo(3).Text = rsTemp(2)
                    End If
                End If
            End If
        End If
        txtInfo(0).SetFocus
    ElseIf Index = 0 Then
        If Trim(txtInfo(4).Text) = "" Then
            gstrSQL = "Select A.����,A.�Ա�,nvl(B.��λ����,'0') From ������Ϣ A,�����ʻ� B Where A.����ID=B.����ID And B.����=[1] And B.����=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_��Ҧ, CStr(Trim(txtInfo(0).Text)))
            If Not rsTemp.EOF Then
                txtInfo(1).Text = rsTemp(0)
                txtInfo(2).Text = rsTemp(1)
                txtInfo(3).Text = rsTemp(2)
            End If
        End If
        txtInfo(1).SetFocus
    ElseIf Index = 3 Then
        cmdOK.SetFocus
    Else
        txtInfo(Index + 1).SetFocus
    End If
End Sub

