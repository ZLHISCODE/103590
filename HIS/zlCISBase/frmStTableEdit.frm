VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmStTableEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "������"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6060
   Icon            =   "frmStTableEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin RichTextLib.RichTextBox rtfTableTitle 
      Height          =   4815
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   8493
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      MaxLength       =   250
      Appearance      =   0
      TextRTF         =   $"frmStTableEdit.frx":17D2A
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6060
      TabIndex        =   6
      Top             =   5490
      Width           =   6060
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ����&C��"
         Height          =   350
         Left            =   4800
         TabIndex        =   5
         Top             =   160
         Width           =   1100
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   3360
         TabIndex        =   4
         Top             =   160
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   1
      Top             =   180
      Width           =   4695
   End
   Begin VB.Label lblTableTile 
      Caption         =   "����ͷ(&T)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblStTableName 
      Caption         =   "������(&N)"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmStTableEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng��� As Long
Private mlngStPathID As Long
Private mstrName As String
Private mstrTableTitle As String
Private mblnOK As Boolean

Public Function ShowMe(ByRef FrmParent As Object, ByRef lngStPathID As Long, Optional ByVal lng��� As Long, Optional ByVal strName As String, Optional ByVal strTableTitle As String) As Boolean
'���ܣ���ʾ�����»�������
'������ lngStPathID ��׼·��ID
'       lng��� ��׼·������� 0-������>0 -�޸�
'       strName ��׼·��������
'       strTableTitle ��׼·������ͷ

    mlngStPathID = lngStPathID
    mlng��� = lng���
    mstrName = strName
    mblnOK = False
    mstrTableTitle = strTableTitle
    Me.Show 1, FrmParent
    ShowMe = mblnOK
    lngStPathID = mlngStPathID
    
End Function

Private Sub cmdCancel_Click()

    mblnOK = False
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    Dim strSql As String
    
    On Error GoTo errH
    If Trim(txtName.Text) = "" Then
        MsgBox "�����Ʋ���Ϊ��", vbInformation, gstrSysName
        txtName.SetFocus
        Exit Sub
    End If
    
    If mstrName = Trim(txtName.Text) And mstrTableTitle = Trim(rtfTableTitle.Text) Then
        mblnOK = False
        Unload Me
    Else
        mblnOK = True
    End If
    
    mstrName = Trim(txtName.Text)
    mstrTableTitle = Trim(rtfTableTitle.Text)
    
    strSql = "Zl_��׼·����_Update(" & mlngStPathID & "," & mlng��� & ",'" & mstrName & "','" & mstrTableTitle & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()

    txtName.SetFocus
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
    
End Sub

Private Sub Form_Load()
'���ݴ���Ĳ�����ʼ������
    If mlng��� = 0 Then
        Me.Caption = "������"
    Else
        Me.Caption = "�޸ı�"
        txtName.Text = mstrName
        rtfTableTitle.Text = mstrTableTitle
    End If
     
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'���ܣ��س���λ��һ���ؼ�
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub


Private Sub Form_Resize()
    
    If Me.WindowState = vbMaximized Or Me.WindowState = vbMinimized Then Exit Sub
    '������ı䴰���С
    If Me.Width < 6150 Then Me.Width = 6150
    If Me.Height < 6500 Then Me.Height = 6500
    
    rtfTableTitle.Height = picBottom.Top - 20 - rtfTableTitle.Top
    
End Sub

Private Sub txtName_GotFocus()
'���ܣ���ȡ����ȫѡ
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub
