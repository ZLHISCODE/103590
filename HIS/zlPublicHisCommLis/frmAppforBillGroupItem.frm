VERSION 5.00
Begin VB.Form frmAppforBillGroupItem 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "��������"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   0
      TabIndex        =   5
      Top             =   1710
      Width           =   3705
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1230
      TabIndex        =   3
      Top             =   990
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   780
      TabIndex        =   2
      Top             =   2010
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2280
      TabIndex        =   1
      Top             =   2010
      Width           =   1335
   End
   Begin VB.TextBox txtNO 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1230
      TabIndex        =   0
      Top             =   360
      Width           =   2025
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��������:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   6
      Top             =   1050
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "����:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   4
      Top             =   420
      Width           =   600
   End
End
Attribute VB_Name = "frmAppforBillGroupItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnfrmShow As Boolean                     '�����Ƿ���ʾ
Private mlngkeyID As Long                          '����ID
Private mstrNO As String                           '����
Private mstrName As String                         '����
Private mlngMainID As Long                         '���뵥id
Private mstrNametext As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveDate = True Then
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    If mblnfrmShow = False Then
        If mlngkeyID = 0 Then
            Call GetMaxNO
            Me.TxtName.SetFocus
        Else
            Me.txtNO = mstrNO
            Me.TxtName = mstrName
        End If
        mblnfrmShow = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnfrmShow = False
End Sub
Private Sub txtName_Change()
    If LenB(StrConv(TxtName.Text, vbFromUnicode)) > 20 Then MsgBox "���Ʋ��ܳ���20���ֽ�!", vbExclamation + vbOKOnly, "���ƹ���": TxtName.Text = mstrNametext
End Sub
Private Sub txtName_GotFocus()
    TxtName.SelStart = 0
    TxtName.SelLength = Len(TxtName)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdOK_Click
    End If
    mstrNametext = TxtName.Text
End Sub

Private Function SaveDate() As Boolean
          Dim strSQL As String
              
1         On Error GoTo SaveDate_Error

2         If Trim(Me.txtNO.Text) = "" Then
3             MsgBox "������������ܱ���!", vbInformation, "�������뵥����"
4             Me.txtNO.SetFocus
5             Exit Function
6         End If
          
7         If Trim(Me.TxtName.Text) = "" Then
8             MsgBox "���������ƺ���ܱ���!", vbInformation, "�������뵥����"
9             Me.TxtName.SetFocus
10            Exit Function
11        End If
          
          '����
12        strSQL = "Zl_�������뵥����_EDIT(1," & mlngkeyID & ",'" & Me.txtNO & "','" & Me.TxtName & "'," & mlngMainID & ")"
13        ComExecuteProc Sel_Lis_DB, strSQL, "�����������"
14        SaveDBLog 18, 6, 0, "����", "������Ŀ����:" & TxtName.Text, 1012, "���뵥����"
15        SaveDate = True


16        Exit Function
SaveDate_Error:
17        Call writeErrLog("zl9LisInsideComm", "frmAppforBillGroupItem", "ִ��(SaveDate)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
18        Err.Clear
          
End Function

Private Sub txtNO_GotFocus()
    txtNO.SelStart = 0
    txtNO.SelLength = Len(txtNO)
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtName.SetFocus
    End If
End Sub

Private Sub GetMaxNO()
          '���ܣ�         ��ʼ������
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
              
1         On Error GoTo GetMaxNO_Error

2         strSQL = "select nvl(max(����),0) ���� from �������뵥���� "
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�������뵥����")
4         If rsTmp("����") = 0 Then
5             Me.txtNO = "001"
6         Else
7             Me.txtNO = Format(Val(rsTmp("����")) + 1, "000")
8         End If


9         Exit Sub
GetMaxNO_Error:
10        Call writeErrLog("zl9LisInsideComm", "frmAppforBillGroupItem", "ִ��(GetMaxNO)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
11        Err.Clear
          
End Sub

Public Sub showMe(objfrm As Object, lngMainID As Long, lngID As Long, strNO As String, strName As String)
    '����           ��������
    
    mlngMainID = lngMainID
    mlngkeyID = lngID
    mstrNO = strNO
    mstrName = strName
    Me.Show vbModal, objfrm
End Sub


