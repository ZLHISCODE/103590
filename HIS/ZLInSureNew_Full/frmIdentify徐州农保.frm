VERSION 5.00
Begin VB.Form frmIdentify����ũ�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������֤"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   400
      Left            =   2463
      TabIndex        =   6
      Top             =   2535
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   400
      Left            =   1365
      TabIndex        =   5
      Top             =   2535
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -120
      TabIndex        =   12
      Top             =   2295
      Width           =   3990
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   4
      Left            =   1193
      TabIndex        =   4
      Top             =   1860
      Width           =   2370
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   1193
      TabIndex        =   3
      Top             =   1425
      Width           =   2370
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   1193
      TabIndex        =   2
      Top             =   990
      Width           =   2370
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   1193
      TabIndex        =   1
      Top             =   585
      Width           =   2370
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   1193
      TabIndex        =   0
      Top             =   195
      Width           =   2370
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   4
      Left            =   788
      TabIndex        =   11
      Top             =   1935
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   180
      Index           =   3
      Left            =   788
      TabIndex        =   10
      Top             =   1500
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   2
      Left            =   428
      TabIndex        =   9
      Top             =   1065
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����ҽ����"
      Height          =   180
      Index           =   1
      Left            =   255
      TabIndex        =   8
      Top             =   660
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ҽ��סԺ��"
      Height          =   180
      Index           =   0
      Left            =   248
      TabIndex        =   7
      Top             =   270
      Width           =   900
   End
End
Attribute VB_Name = "frmIdentify����ũ��"
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
    If Trim(txtEdit(1).Text) = "" Then
        MsgBox "�������벡��ҽ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    txtEdit(4).Tag = CStr(Year(Date) - CInt(txtEdit(4).Text)) & "-01-01"
    mstrOther = "": mstrPatient = ""
    
    mstrPatient = txtEdit(1).Text & ";"                                 '0 ����
    mstrPatient = mstrPatient & txtEdit(1).Text & ";"                   '1 ҽ���ʺ�
    mstrPatient = mstrPatient & ";"                                     '2 ����
    mstrPatient = mstrPatient & txtEdit(2).Text & ";"                   '3 ����
    mstrPatient = mstrPatient & txtEdit(3).Text & ";"                   '4 �Ա�
    mstrPatient = mstrPatient & txtEdit(4).Tag & ";"                    '5 ��������
    mstrPatient = mstrPatient & ";"                                     '6 ���֤
    mstrPatient = mstrPatient & ";"                                     '7 ��λ����/����
        
    mstrOther = mstrOther & ";"                                         '8 ҽ����������(����)
    mstrOther = mstrOther & txtEdit(0).Tag & ";"                        '9 ˳���
    mstrOther = mstrOther & ";"                                         '10 ���
    mstrOther = mstrOther & ";"                                         '11 ���
    mstrOther = mstrOther & ";"                                         '12 ��ǰ״̬
    mstrOther = mstrOther & ";"                                         '13 ����ID
    mstrOther = mstrOther & ";"
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

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Index = 0 Then
        Set rsTemp = gcn����ũ��.Execute("Select * From inPatient Where No='" & Trim(txtEdit(Index)) & "'")
        If rsTemp.EOF Then
            MsgBox "ҽ��ǰ�û���û�иò��˵�סԺ��Ϣ�����Ƚ���ҽ����Ժ�Ǽ�", vbInformation, gstrSysName
            txtEdit(Index).SelStart = 0
            txtEdit(Index).SelLength = Len(txtEdit(Index).Text)
        Else
            txtEdit(0).Tag = rsTemp!ID
            txtEdit(2).Text = Trim(rsTemp!Name)
            txtEdit(3).Text = rsTemp!Sex
            txtEdit(4).Text = rsTemp!age
            If Nvl(rsTemp!id_card, "") = "" Then
                txtEdit(1).Enabled = True
                On Error Resume Next
                DoEvents
                txtEdit(1).SetFocus
                On Error GoTo 0
            Else
                txtEdit(1).Text = rsTemp!id_card
            End If
        End If
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    txtEdit(Index).SelStart = 0
    txtEdit(Index).SelLength = Len(txtEdit(Index).Text)
End Sub
