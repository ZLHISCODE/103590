VERSION 5.00
Begin VB.Form frmErrAsk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ʾ"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdRetry 
      Caption         =   "����(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1365
      TabIndex        =   4
      Top             =   1335
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2280
      TabIndex        =   3
      Top             =   1335
      Width           =   900
   End
   Begin VB.TextBox txtHelp 
      Height          =   1635
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmErrAsk.frx":0000
      Top             =   1800
      Width           =   4275
   End
   Begin VB.CommandButton cmdCopyScreen 
      Caption         =   "��ͼ(&S)"
      Height          =   350
      Left            =   3195
      TabIndex        =   1
      Top             =   1335
      Width           =   900
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00E1E1FF&
      Caption         =   "����(&E)"
      Height          =   350
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1335
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox picS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3390
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   156
      TabIndex        =   5
      Top             =   1275
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   285
      Picture         =   "frmErrAsk.frx":0065
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      Caption         =   "�����ţ�"
      Height          =   180
      Left            =   1965
      TabIndex        =   9
      Top             =   105
      Width           =   900
   End
   Begin VB.Label lblScrip 
      AutoSize        =   -1  'True
      Caption         =   "˵����"
      Height          =   180
      Left            =   975
      TabIndex        =   8
      Top             =   105
      Width           =   540
   End
   Begin VB.Label lblNote 
      Caption         =   "    �����������û��Ķ�ռ�����°�װ�˲���ϵͳ�����Ĵ����ų���ռʹ�������Բ������У����貿����װ��ϵͳ��"
      Height          =   585
      Left            =   975
      TabIndex        =   7
      Top             =   315
      Width           =   3390
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAsk 
      AutoSize        =   -1  'True
      Caption         =   "����һ����"
      Height          =   180
      Left            =   975
      TabIndex        =   6
      Top             =   1005
      Width           =   1080
   End
End
Attribute VB_Name = "frmErrAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mtmrConnect  As clsTimer '�Զ�����Timer
Attribute mtmrConnect.VB_VarHelpID = -1
Private mbytReturn As Byte

Private Sub cmdCancel_Click()
    mbytReturn = 0
    Unload Me
End Sub

Private Sub cmdCopyScreen_Click()
    Call SaveScreen(txtHelp.Text, picS)
End Sub

Private Sub cmdEnd_Click()
    If cmdRetry.Caption = "����(&S)" Then
        TerminateProcess GetCurrentProcess, 0
    Else
       If CheckAdoConnction(False) = False Then
         cmdCancel_Click
       End If
    End If
End Sub

Private Sub cmdRetry_Click()
    mbytReturn = 1
    Unload Me
End Sub

Public Function ShowEdit(lngErrNum As Long, strNote As String, strErrInfo As String, ByVal blnConnect As Boolean) As Byte
'���ܣ���ʾ������ʾ���ڣ�����ѡ������
'������lngErrNum   ������
'      strNote     ��������
'      strErrInfo  ��ϸ�Ĵ�����Ϣ
'���أ���һ����������ʾ��1-���ԣ�0-ȡ��
    mbytReturn = 0
    
    If gblnSilentMode Then
        gstrErrorContent = strErrInfo
        ShowEdit = 0
        Exit Function
    End If
    
    lblNumber.Caption = "��ţ�" & lngErrNum
    lblNote.Caption = Space(4) & strNote
    txtHelp.Text = strErrInfo
    
    If blnConnect Then
        cmdRetry.Caption = "����(&S)"
        cmdRetry.Visible = True
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

    frmErrAsk.Show vbModal
    If blnConnect Then
        cmdRetry.Caption = "����(&O)"
    End If
    ShowEdit = mbytReturn
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        mbytReturn = 0
        Unload Me
    End If
End Sub

Private Sub mtmrConnect_ThatTime()
    If CheckAdoConnction(False) = False Then
      If ObjPtr(mtmrConnect) > 0 Then
        mtmrConnect.Interval = 0
      End If
      cmdCancel_Click
    End If
End Sub


