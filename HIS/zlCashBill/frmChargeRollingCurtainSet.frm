VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChargeRollingCurtainSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7290
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChargeRollingCurtainSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   345
      Left            =   2205
      TabIndex        =   5
      Top             =   2835
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "hh:mm:ss"
      Format          =   115933186
      CurrentDate     =   42402
   End
   Begin VB.CheckBox chkTime 
      Caption         =   "ȱʡ����ʱ��"
      Height          =   345
      Left            =   645
      TabIndex        =   4
      Top             =   2835
      Width           =   1635
   End
   Begin VB.Frame fraPringMode 
      Caption         =   "�ɿ����ӡ��ʽ"
      Height          =   1590
      Left            =   645
      TabIndex        =   11
      Top             =   1125
      Width           =   6030
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "��ӡ����(&S)"
         Height          =   375
         Index           =   0
         Left            =   4125
         TabIndex        =   3
         Top             =   638
         Width           =   1530
      End
      Begin VB.OptionButton optPrintMode 
         Caption         =   "���ʺ�ѡ���Ƿ��ӡ(&3)"
         Height          =   300
         Index           =   2
         Left            =   1365
         TabIndex        =   2
         Top             =   1050
         Width           =   2760
      End
      Begin VB.OptionButton optPrintMode 
         Caption         =   "���ʺ��Զ���ӡ(&2)"
         Height          =   300
         Index           =   1
         Left            =   1365
         TabIndex        =   1
         Top             =   675
         Value           =   -1  'True
         Width           =   2190
      End
      Begin VB.OptionButton optPrintMode 
         Caption         =   "���ʺ󲻴�ӡ(&1)"
         Height          =   300
         Index           =   0
         Left            =   1395
         TabIndex        =   0
         Top             =   315
         Width           =   1935
      End
      Begin VB.Image imgPrint 
         Height          =   720
         Left            =   375
         Picture         =   "frmChargeRollingCurtainSet.frx":06EA
         Top             =   465
         Width           =   720
      End
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
      Left            =   -855
      TabIndex        =   9
      Top             =   855
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
      Index           =   1
      Left            =   -90
      TabIndex        =   7
      Top             =   3375
      Width           =   8925
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6000
      TabIndex        =   8
      Top             =   3750
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4665
      TabIndex        =   6
      Top             =   3735
      Width           =   1100
   End
   Begin VB.Image imgNotes 
      Height          =   720
      Left            =   270
      Picture         =   "frmChargeRollingCurtainSet.frx":15B4
      Top             =   165
      Width           =   720
   End
   Begin VB.Label lblTittle 
      Caption         =   "   ����ݾ���������ò������ɿ����ӡ��ʽ��Ҫ�������շ�Ա�����ʺ���ƽɿ���Ĵ�ӡ��ʽ��"
      Height          =   600
      Left            =   1080
      TabIndex        =   10
      Top             =   405
      Width           =   5985
   End
End
Attribute VB_Name = "frmChargeRollingCurtainSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As String, mstrPrivs As String, mblnOK As Boolean
Public Function ShowMe(ByVal frmMain As Form, _
    ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ز��������
    '����:�������óɹ�������true,���򷵻�False
    '����:���˺�
    '����:2013-09-12 14:33:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnOK = False: mlngModule = lngModule: mstrPrivs = strPrivs
    Me.Show 1, frmMain
    ShowMe = mblnOK
End Function

Private Sub LoadPara()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز���
    '����:���˺�
    '����:2013-09-12 15:26:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strValue As String
    i = Val(zlDatabase.GetPara("�ɿ����ӡ��ʽ", glngSys, mlngModule, 0, Array(optPrintMode(0), optPrintMode(1), optPrintMode(2)), InStr(1, mstrPrivs, ";��������;") > 0))
    If i > 2 Or i < 0 Then
        optPrintMode(0).Value = True
    Else
        optPrintMode(i).Value = True
    End If
    strValue = zlDatabase.GetPara("ȱʡ����ʱ��", glngSys, mlngModule, "", dtpTime, InStr(1, mstrPrivs, ";��������;") > 0)
    If strValue = "" Then
        dtpTime.Enabled = False
        chkTime.Value = 0
    Else
        dtpTime.Enabled = True
        chkTime.Value = 1
        dtpTime.Value = Format(strValue, "hh:mm:ss")
    End If
End Sub

Private Sub SavePara()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:���˺�
    '����:2013-09-12 15:28:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call zlDatabase.SetPara("�ɿ����ӡ��ʽ", IIf(optPrintMode(0).Value, 0, IIf(optPrintMode(1).Value, 1, 2)), glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0)
    If chkTime.Value = 1 Then
        Call zlDatabase.SetPara("ȱʡ����ʱ��", Format(dtpTime.Value, "hh:mm:ss"), glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0)
    Else
        Call zlDatabase.SetPara("ȱʡ����ʱ��", "", glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0)
    End If
End Sub

Private Sub chkTime_Click()
    If chkTime.Value = 1 Then
        dtpTime.Enabled = True
    Else
        dtpTime.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Call SavePara
    Unload Me
    mblnOK = True
End Sub
Private Sub cmdPrintSetup_Click(Index As Integer)
        Call ReportPrintSet(gcnOracle, glngSys, "zl" & Int(glngSys / 100) & "_INSIDE_1506", Me)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    Call LoadPara
End Sub

