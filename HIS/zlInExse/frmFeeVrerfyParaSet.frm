VERSION 5.00
Begin VB.Form frmFeeVrerfyParaSet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��������"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   420
      Left            =   3090
      TabIndex        =   1
      Top             =   2730
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   420
      Left            =   4695
      TabIndex        =   5
      Top             =   2730
      Width           =   1500
   End
   Begin VB.CheckBox chk��� 
      Caption         =   "����תסԺ���������(&V)"
      Height          =   270
      Left            =   1155
      TabIndex        =   0
      Top             =   1710
      Width           =   3405
   End
   Begin VB.Frame fraSplit 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   1
      Left            =   -45
      TabIndex        =   3
      Top             =   2400
      Width           =   8925
   End
   Begin VB.Frame fraSplit 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   1275
      Width           =   8925
   End
   Begin VB.Label lblTittle 
      Caption         =   $"frmFeeVrerfyParaSet.frx":0000
      Height          =   945
      Left            =   990
      TabIndex        =   4
      Top             =   315
      Width           =   5205
   End
   Begin VB.Image imgPit 
      Height          =   720
      Left            =   105
      Picture         =   "frmFeeVrerfyParaSet.frx":0093
      Top             =   435
      Width           =   720
   End
End
Attribute VB_Name = "frmFeeVrerfyParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************
'**�ò����ѷ���,�Ѿ������˷��ò��ֵĹ������ֽ�������

Private mlngModule As String, mstrPrivs As String, mblnOk As Boolean
Public Function ShowMe(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ز��������
    '����:�������óɹ�������true,���򷵻�False
    '����:���˺�
    '����:2011-02-09 11:35:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnOk = False: mlngModule = lngModule: mstrPrivs = strPrivs
    Me.Show 1, frmMain
    ShowMe = mblnOk
End Function
Private Sub LoadPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز���
    '����:���˺�
    '����:2011-02-09 11:36:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    chk���.Value = IIf(Val(zlDatabase.GetPara("����תסԺ�����", glngSys, mlngModule, 0, Array(chk���), InStr(1, mstrPrivs, ";��������;") > 0)) = 1, 1, 0)
End Sub
Private Sub SavePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:���˺�
    '����:2011-02-09 11:36:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call zlDatabase.SetPara("����תסԺ�����", IIf(chk���.Value = 1, 1, 0), glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call SavePara
    Unload Me
    mblnOk = True
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    Call LoadPara
End Sub
