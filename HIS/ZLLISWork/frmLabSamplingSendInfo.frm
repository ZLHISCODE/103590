VERSION 5.00
Begin VB.Form frmLabSamplingSendInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�걾�ͼ�"
   ClientHeight    =   1656
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   3948
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1656
   ScaleWidth      =   3948
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2616
      TabIndex        =   5
      Top             =   1200
      Width           =   1104
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   1320
      TabIndex        =   4
      Top             =   1200
      Width           =   1104
   End
   Begin VB.Frame Frame1 
      Height          =   36
      Left            =   -312
      TabIndex        =   3
      Top             =   1032
      Width           =   7308
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "��ӡ�ͼ쵥"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Value           =   1  'Checked
      Width           =   1956
   End
   Begin VB.TextBox txtSendName 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1176
      TabIndex        =   0
      Top             =   240
      Width           =   2208
   End
   Begin VB.Label lblSendinfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ͼ���:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   456
      TabIndex        =   1
      Top             =   300
      Width           =   636
   End
End
Attribute VB_Name = "frmLabSamplingSendInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mblnOK As Boolean                                   '����ȷ��Ϊ�棬����ȡ����
Dim mstrName As String                                  '�ͼ�����
Dim mblnPrint As Boolean                                '�Ƿ��ӡ�嵥


Private Sub CmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mstrName = Me.txtSendName
    mblnPrint = chkPrint.Value
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    Me.txtSendName = UserInfo.����
End Sub

Public Function ShowMe(Objfrm As Object, strName As String, blnPrint As Boolean) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����                               ���ͼ촰��
    '����
    '       objfrm                      ���������
    '       strName                     �ͼ���
    '       blnPrint                    �Ƿ��ӡ
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    mstrName = strName
    mblnPrint = blnPrint
    Me.Show vbModal, Objfrm
    ShowMe = mblnOK
    strName = mstrName
    blnPrint = mblnPrint
End Function

Private Sub Form_Unload(Cancel As Integer)
    
End Sub

Private Sub txtSendName_GotFocus()
    Me.txtSendName.SelStart = 0
    Me.txtSendName.SelLength = Len(Me.txtSendName)
End Sub
