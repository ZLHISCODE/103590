VERSION 5.00
Begin VB.Form frmIdentify�㽭 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -473
      TabIndex        =   21
      Top             =   1965
      Width           =   6450
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   8
      Left            =   3671
      TabIndex        =   20
      Top             =   1485
      Width           =   1755
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   7
      Left            =   956
      TabIndex        =   18
      Top             =   1485
      Width           =   1755
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   6
      Left            =   3671
      TabIndex        =   16
      Top             =   1035
      Width           =   1755
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   5
      Left            =   956
      TabIndex        =   14
      Top             =   1035
      Width           =   1755
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   4
      Left            =   3671
      TabIndex        =   12
      Top             =   585
      Width           =   1755
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   956
      TabIndex        =   10
      Top             =   585
      Width           =   1755
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   3671
      TabIndex        =   8
      Top             =   150
      Width           =   1755
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   956
      TabIndex        =   0
      Top             =   150
      Width           =   1755
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   0
      Left            =   956
      MaxLength       =   10
      TabIndex        =   4
      Top             =   150
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   4326
      TabIndex        =   3
      Top             =   2145
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3191
      TabIndex        =   2
      Top             =   2145
      Width           =   1100
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "����(&R)"
      Height          =   400
      Left            =   176
      TabIndex        =   1
      Top             =   2145
      Width           =   1100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   8
      Left            =   2880
      TabIndex        =   19
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   7
      Left            =   525
      TabIndex        =   17
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   180
      Index           =   6
      Left            =   3240
      TabIndex        =   15
      Top             =   1110
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��Ա״̬"
      Height          =   180
      Index           =   5
      Left            =   165
      TabIndex        =   13
      Top             =   1110
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���֤��"
      Height          =   180
      Index           =   4
      Left            =   2880
      TabIndex        =   11
      Top             =   660
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   3
      Left            =   525
      TabIndex        =   9
      Top             =   660
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ҽ����"
      Height          =   180
      Index           =   2
      Left            =   3060
      TabIndex        =   7
      Top             =   225
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   1
      Left            =   525
      TabIndex        =   6
      Top             =   225
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   0
      Left            =   529
      TabIndex        =   5
      Top             =   225
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "frmIdentify�㽭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte
Public mstrPatient As String, mstrOther As String

Public Function GetPatient(bytType As Byte, str���� As String) As String
    '������bytType-ʶ�����ͣ�0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
    mbytType = bytType
    Me.Show vbModal
    str���� = txtInfo(1).Text
    GetPatient = mstrPatient & mstrOther
End Function

Private Sub cmdCancel_Click()
    mstrPatient = "": mstrOther = ""
    Me.Hide
    'ȡ��
End Sub

Private Sub cmdOK_Click()
    'ȷ��
    mstrOther = "": mstrPatient = ""
reQuery1:
    gstrInfo = Space(1024)
    glngReturn = QUERY_HANDLE("13|" & txtInfo(2).Text & "|DF0431|", gstrInfo)  '��ȡ������Ϣ
    WriteInfo Trim(gstrInfo)
    If CheckReturn�㽭() = False Then
        If gstrInfo = "" Then
            GoTo reQuery1
        Else
            Exit Sub
        End If
    End If
    txtInfo(0).Tag = "(" & Split(gstrInfo, "|")(8) & ")"
    
reQuery2:
    gstrInfo = Space(1024)
    glngReturn = QUERY_HANDLE("13|" & txtInfo(2).Text & "|DF0432|", gstrInfo)  '��ȡ������Ϣ
    WriteInfo Trim(gstrInfo)
    If CheckReturn�㽭() = False Then
        If gstrInfo = "" Then
            GoTo reQuery2
        Else
            Exit Sub
        End If
    End If
    txtInfo(2).Tag = Split(gstrInfo, "|")(0)
    
    mstrPatient = txtInfo(1).Text & ";"                                 '0 ����
    mstrPatient = mstrPatient & txtInfo(2).Text & ";"                   '1 ҽ���ʺ�
    mstrPatient = mstrPatient & txtInfo(0).Text & ";"                   '2 ����
    mstrPatient = mstrPatient & txtInfo(3).Text & ";"                   '3 ����
    mstrPatient = mstrPatient & txtInfo(6).Text & ";"                   '4 �Ա�
    mstrPatient = mstrPatient & txtInfo(8).Text & ";"                   '5 ��������
    mstrPatient = mstrPatient & txtInfo(4).Text & ";"                   '6 ���֤
    mstrPatient = mstrPatient & txtInfo(0).Tag & ";"                    '7 ��λ����/����
    
    mstrOther = mstrOther & txtInfo(1).Tag & ";"                        '8 ҽ����������(����)
    mstrOther = mstrOther & ";"                                         '9 ˳���
    mstrOther = mstrOther & ";"                                         '10 ���
    mstrOther = mstrOther & txtInfo(2).Tag & ";"                        '11 ���
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

Private Sub cmdRead_Click()
    '����
    Dim strTemp As String
    Dim strfksj As String
    Dim datCurr As Date
    
    datCurr = zlDatabase.Currentdate
    
    WriteInfo "��ʼ����"
    cmdOK.Enabled = False
    If txtInfo(1).Text = "" Then
        WriteInfo "��IC����ȡ����"
        glngReturn = readCardID(strTemp)
        WriteInfo "�������أ�" & strTemp
        If glngReturn < 0 Then Exit Sub
        txtInfo(1).Text = Trim(strTemp)
    Else
        WriteInfo "�ֹ����뿨��"
    End If
reQuery01:
      '*******06��06��13�ղ��䣬����ֿ����ʸ���鹦��********
      'ȡ����ʱ��
      gstrInfo = Space(1024)
      glngReturn = QUERY_HANDLE("13|" & txtInfo(1).Text & "|MF11|", gstrInfo)
      If CheckReturn�㽭() = False Then
        If gstrInfo = "" Then
            GoTo reQuery01
        Else
            Exit Sub
        End If
     End If
      WriteInfo Trim(gstrInfo)
     strfksj = Trim(Split(gstrInfo, "|")(4))
     'ȡ�ʸ���Ϣ
      gstrInfo = Space(1024)
      glngReturn = QUERY_HANDLE("04|" & txtInfo(1).Text & "|" & Format(datCurr, "yyyymmdd") & "|" & strfksj & "|", gstrInfo)
      If CheckReturn�㽭() = False Then
        If gstrInfo = "" Then
            GoTo reQuery01
        Else
            Exit Sub
        End If
     End If
      WriteInfo Trim(gstrInfo)
           
      If Trim(Split(gstrInfo, "|")(0)) = "0" Then '�޷���
      Else
         If Trim(Split(gstrInfo, "|")(0)) = "1" Then '���˷���
             MsgBox "�ÿ������˷��������ڷ�Χ��" & Trim(Split(gstrInfo, "|")(1)) = "1" & "---" & Trim(Split(gstrInfo, "|")(2)) & " ����ԭ��" & Trim(Split(gstrInfo, "|")(3))
         Else '��λ����
             MsgBox "�ÿ�����λ���������ڷ�Χ��" & Trim(Split(gstrInfo, "|")(1)) = "1" & "---" & Trim(Split(gstrInfo, "|")(2)) & " ����ԭ��" & Trim(Split(gstrInfo, "|")(3))
         End If
         Exit Sub
      End If
                 
      '*********����*********
    
    gstrInfo = Space(1024)
     
    glngReturn = QUERY_HANDLE("13|" & txtInfo(1).Text & "|MF12|", gstrInfo)
    WriteInfo Trim(gstrInfo)
    If CheckReturn�㽭() = False Then
        If gstrInfo = "" Then
            GoTo reQuery01
        Else
            Exit Sub
        End If
    End If
    cmdOK.Enabled = True
    txtInfo(1).Text = Trim(Split(gstrInfo, "|")(0))
    txtInfo(2).Text = Trim(Split(gstrInfo, "|")(1))
    txtInfo(3).Text = Trim(Split(gstrInfo, "|")(3))
    txtInfo(4).Text = Trim(Split(gstrInfo, "|")(2))
'    txtInfo(5).Text = IIf(Trim(Split(gstrInfo, "|")(4)) = "1", "����Ա", "�ǹ���Ա")
    txtInfo(6).Text = IIf(Trim(Split(gstrInfo, "|")(5)) = "1", "��", "Ů")
    txtInfo(7).Text = Trim(Split(gstrInfo, "|")(6))
    strTemp = Trim(Split(gstrInfo, "|")(7))
    txtInfo(8).Text = Left(strTemp, 4) & "-" & Mid(strTemp, 5, 2) & "-" & Mid(strTemp, 7)
    On Error Resume Next
    cmdOK.SetFocus
    cmdRead.Enabled = False
reQuery1:
    gstrInfo = Space(1024)
    glngReturn = QUERY_HANDLE("13|" & txtInfo(2).Text & "|DF0431|", gstrInfo)  '��ȡ������Ϣ
    WriteInfo Trim(gstrInfo)
    If CheckReturn�㽭() = False Then
        If gstrInfo = "" Then
            GoTo reQuery1
        Else
            Exit Sub
        End If
    End If
    txtInfo(5).Text = IIf(Trim(Split(gstrInfo, "|")(9)) = "1", "����Ա", "�ǹ���Ա")
End Sub

Private Sub Form_Load()
    cmdOK.Enabled = False
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 And KeyAscii = vbKeyReturn Then
        cmdRead_Click
    End If
End Sub
