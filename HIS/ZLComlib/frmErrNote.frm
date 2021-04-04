VERSION 5.00
Begin VB.Form frmErrNote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ע��"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00E1E1FF&
      Caption         =   "����(&E)"
      Height          =   350
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1155
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton cmdCopyScreen 
      Caption         =   "��ͼ(&S)"
      Height          =   350
      Left            =   3045
      TabIndex        =   5
      Top             =   1155
      Width           =   1080
   End
   Begin VB.TextBox txtHelp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmErrNote.frx":0000
      Top             =   1605
      Width           =   4290
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1905
      TabIndex        =   3
      Top             =   1155
      Width           =   1080
   End
   Begin VB.PictureBox picS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3165
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   156
      TabIndex        =   6
      Top             =   1065
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Label lblNote 
      Caption         =   "    �����������û��Ķ�ռ�����°�װ�˲���ϵͳ�����Ĵ����ų���ռʹ�������Բ������У����貿����װ��ϵͳ��"
      Height          =   585
      Left            =   900
      TabIndex        =   2
      Top             =   465
      Width           =   3255
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblScrip 
      AutoSize        =   -1  'True
      Caption         =   "˵����"
      Height          =   180
      Left            =   900
      TabIndex        =   1
      Top             =   180
      Width           =   540
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      Caption         =   "�����ţ�"
      Height          =   180
      Left            =   1770
      TabIndex        =   0
      Top             =   180
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmErrNote.frx":0067
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmErrNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mtmrConnect  As clsTimer '�Զ�����Timer
Attribute mtmrConnect.VB_VarHelpID = -1

Private Sub cmdCopyScreen_Click()
     Call SaveScreen(txtHelp.Text, picS)
End Sub

Private Sub cmdEnd_Click()
    If cmdEnd.Caption = "����(&E)" Then
        TerminateProcess GetCurrentProcess, 0
    Else
       If CheckAdoConnction(False) = False Then
         cmdOK_Click
       End If
    End If
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Public Sub ShowEdit(lngErrNum As Long, strNote As String, strErrInfo As String, ByVal blnConnect As Boolean)
'���ܣ���ʾ������ʾ����
'������lngErrNum   ������
'      strNote     ��������
'      strErrInfo  ��ϸ�Ĵ�����Ϣ
    
    lblNumber.Caption = "��ţ�" & lngErrNum
    lblNote.Caption = Space(4) & strNote
    txtHelp.Text = strErrInfo

    
    If blnConnect Then
        cmdEnd.Caption = "����(&S)"
        cmdEnd.Visible = True

        txtHelp.Text = "��ORACLE�����������Ѿ��Ͽ�," & vbNewLine & "����������ָ����ֶ�������������,���(����)!"
        
        '�����Զ���������,Ĭ��10����������
        Set mtmrConnect = New clsTimer
        mtmrConnect.Interval = 10000
    Else
        If lngErrNum = -2147467259 _
            And (InStr(strErrInfo, "E_FAIL") > 0 _
                Or InStr(UCase(strErrInfo), "UNKNOW") > 0 _
                Or strErrInfo = "δָ���Ĵ���") Then
            '�����ṩ������������񷵻�һ�� E_FAIL ״̬
            cmdEnd.Visible = True
        ElseIf lngErrNum = -2147217900 _
            And (InStr(strErrInfo, "ORA-00028") > 0 _
                Or InStr(strErrInfo, "ORA-01012") > 0 _
                Or InStr(strErrInfo, "ORA-03113") > 0) Then
            'ORA-00028: ���ĻỰ������ֹ
            'ORA-01012: û�е�¼
            'ORA-03113: ͨ��ͨ�����ļ�����
            cmdEnd.Visible = True
        End If
    End If
    
    frmErrNote.Show vbModal
    If blnConnect Then
        cmdEnd.Caption = "����(&E)"
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub mtmrConnect_ThatTime()
'    mtmrConnect.Interval = 0
    If CheckAdoConnction(False) = False Then
      If ObjPtr(mtmrConnect) > 0 Then
        mtmrConnect.Interval = 0
      End If
      cmdOK_Click
    End If
End Sub

