VERSION 5.00
Begin VB.Form frmLabToCompare 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ϊ�ȶԱ걾"
   ClientHeight    =   2310
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5730
   Icon            =   "frmLabToCompare.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame frmLine 
      Height          =   30
      Left            =   -30
      TabIndex        =   7
      Top             =   1650
      Width           =   5745
   End
   Begin VB.ComboBox cbo�ȶԺ� 
      Height          =   300
      Left            =   2265
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3195
      TabIndex        =   1
      Top             =   1770
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4380
      TabIndex        =   0
      Top             =   1770
      Width           =   1100
   End
   Begin VB.Label lbl�ȶԺ� 
      AutoSize        =   -1  'True
      Caption         =   "����������Ϊ"
      Height          =   180
      Left            =   1140
      TabIndex        =   6
      Top             =   1140
      Width           =   1080
   End
   Begin VB.Label lbl�������� 
      AutoSize        =   -1  'True
      Caption         =   "����������####"
      Height          =   180
      Left            =   1140
      TabIndex        =   4
      Top             =   795
      Width           =   1260
   End
   Begin VB.Label lbl�������� 
      AutoSize        =   -1  'True
      Caption         =   "�������ڣ�####"
      Height          =   180
      Left            =   1140
      TabIndex        =   3
      Top             =   480
      Width           =   1260
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����Ϊ�������߱��ȶ���������ѡ��Ҫָ���ȶԺź�ȷ����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   780
      TabIndex        =   2
      Top             =   135
      Width           =   4860
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   210
      Picture         =   "frmLabToCompare.frx":058A
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frmLabToCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngID As Long   '��ǰid
Private mblnOK As Boolean

'��ʱ����
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long

Public Function ShowMe(ByVal frmParent As Form, lngID As Long) As Boolean
    Dim strDate As String
    mlngID = lngID
    
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select L.ҽ��id, To_Char(L.����ʱ��, 'yyyy-mm-dd') As ����ʱ��, M.���� As ����" & vbNewLine & _
            "From ����걾��¼ L, �������� M" & vbNewLine & _
            "Where L.����id = M.ID And L.ID = [1] And L.����ʱ�� Is Not Null"

    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngID)
    With rsTemp
        If .RecordCount <= 0 Then
            MsgBox "����ĿΪ�ֹ���Ŀ���������δ��д�������������Ϊ�ȶ�������", vbInformation, gstrSysName
            Unload Me: ShowMe = False: Exit Function
        End If
        Me.lbl��������.Caption = "�������ڣ�" & Format(!����ʱ��, "yyyy-MM-dd")
        Me.lbl��������.Caption = "����������" & !����
    End With
    
    With Me.cbo�ȶԺ�
        .Clear: .AddItem "�ȶ�1": .AddItem "�ȶ�2": .AddItem "�ȶ�3":: .AddItem "�ȶ�4": .AddItem "�ȶ�5": .ListIndex = 0
    End With
    
    Me.Show vbModal, frmParent
    ShowMe = mblnOK: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = False: Exit Function
End Function

Private Sub cmdCancel_Click()
    mblnOK = False: Unload Me
End Sub

Private Sub cmdOK_Click()

    gstrSql = "Zl_����걾��¼_��Ϊ�ȶ�(" & mlngID & "," & Me.cbo�ȶԺ�.ListIndex + 1 & ")"
    Err = 0: On Error GoTo ErrHand
    zldatabase.ExecuteProcedure gstrSql, Me.Caption
    mblnOK = True: Unload Me
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    Call zlCommFun.OpenIme(False)
End Sub


