VERSION 5.00
Begin VB.Form frmSet�ش�У԰�� 
   AutoRedraw      =   -1  'True
   Caption         =   "����"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   ControlBox      =   0   'False
   Icon            =   "frmSet�ش�У԰��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2520
   ScaleWidth      =   3735
   StartUpPosition =   1  '����������
   Begin VB.TextBox Txt�޶� 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   1500
      TabIndex        =   1
      Text            =   "2000"
      Top             =   1380
      Width           =   930
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   90
      ScaleHeight     =   375
      ScaleWidth      =   7575
      TabIndex        =   7
      Top             =   855
      Width           =   7575
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1395
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "1"
         Top             =   75
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ����(&D)"
         Height          =   180
         Index           =   3
         Left            =   300
         TabIndex        =   9
         Top             =   135
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�Ŵ���"
         Height          =   180
         Index           =   4
         Left            =   1800
         TabIndex        =   8
         Top             =   135
         Width           =   540
      End
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   -165
      TabIndex        =   6
      Top             =   1905
      Width           =   7755
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2595
      TabIndex        =   3
      Top             =   2070
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1335
      TabIndex        =   2
      Top             =   2070
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   30
      TabIndex        =   5
      Top             =   690
      Width           =   7665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ԫ"
      Height          =   180
      Left            =   2505
      TabIndex        =   11
      Top             =   1425
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÿ�ν����޶�"
      Height          =   180
      Left            =   390
      TabIndex        =   10
      Top             =   1470
      Width           =   1080
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmSet�ش�У԰��.frx":000C
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "�����豸�Ĵ��ںźͽ����޶�."
      Height          =   315
      Left            =   540
      TabIndex        =   4
      Top             =   390
      Width           =   7125
   End
End
Attribute VB_Name = "frmSet�ش�У԰��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean
Private mlngҽ������ As Long
Private mlng���� As Long

Private Sub cmdCancel_Click()
    mblnReturn = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngRow As Long
    
    If Trim(txtEdit) = "" Then Exit Sub
    SaveRegInFor g����ģ��, "����", "�˿ں�", Me.txtEdit
    gintComport_�ش�У԰�� = Val(txtEdit)
    
    'ɾ���Ѿ�����
    On Error GoTo errHand
    gstrSQL = "zl_���ղ���_Update(" & mlng���� & ",NULL,'�����޶�' ,'" & Val(Txt�޶�.Text) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mblnReturn = True
    Unload Me
    Exit Sub
errHand:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim strReg As String
    mblnReturn = False
    
    Call GetRegInFor(g����ģ��, "����", "�˿ں�", strReg)
    
    If Val(strReg) = 0 Then
        txtEdit.Text = 0
    Else
        txtEdit.Text = Val(strReg)
    End If
    
     gstrSQL = "Select * From ���ղ��� where ������ ='�����޶�' and ����=" & mlng����
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If Not rsTemp.EOF Then
            Txt�޶�.Text = Format(Val(NVL(rsTemp!����ֵ)), "####0.00;-####0.00; ;")
    End If
End Sub

Public Function ShowME(ByVal lng���� As Long, ByVal lngҽ������ As Long) As Boolean
    mlngҽ������ = lngҽ������
    mlng���� = lng����
    Me.Show 1
    ShowME = mblnReturn
End Function
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
        zlControl.TxtCheckKeyPress txtEdit, KeyAscii, m����ʽ
End Sub



Private Sub Txt�޶�_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress Txt�޶�, KeyAscii, m���ʽ
End Sub
