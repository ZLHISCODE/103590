VERSION 5.00
Begin VB.Form frmSquareAffirmParaSet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��������"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   3615
      TabIndex        =   11
      Top             =   -135
      Width           =   30
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3930
      TabIndex        =   8
      Top             =   165
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3930
      TabIndex        =   9
      Top             =   660
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "��ӡ����(&S)"
      Height          =   345
      Left            =   3930
      TabIndex        =   10
      Top             =   2835
      Width           =   1425
   End
   Begin VB.Frame fraRecored 
      Caption         =   "������˺����Ʊ�ݴ�ӡ��ʽ"
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   195
      TabIndex        =   4
      Top             =   1830
      Width           =   3210
      Begin VB.OptionButton optRecordPrint 
         Caption         =   "�Զ���ӡ"
         Height          =   180
         Index           =   1
         Left            =   525
         TabIndex        =   6
         Top             =   720
         Width           =   1020
      End
      Begin VB.OptionButton optRecordPrint 
         Caption         =   "ѡ���Ƿ��ӡ"
         Height          =   180
         Index           =   2
         Left            =   525
         TabIndex        =   7
         Top             =   1020
         Width           =   1380
      End
      Begin VB.OptionButton optRecordPrint 
         Caption         =   "����ӡ"
         Height          =   180
         Index           =   0
         Left            =   525
         TabIndex        =   5
         Top             =   450
         Value           =   -1  'True
         Width           =   900
      End
   End
   Begin VB.Frame fraCharge 
      Caption         =   "����ȷ�����շ�Ʊ�ݴ�ӡ��ʽ"
      ForeColor       =   &H00000000&
      Height          =   1380
      Left            =   195
      TabIndex        =   0
      Top             =   270
      Width           =   3180
      Begin VB.OptionButton optChargePrint 
         Caption         =   "����ӡ"
         Height          =   180
         Index           =   0
         Left            =   465
         TabIndex        =   1
         Top             =   375
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton optChargePrint 
         Caption         =   "ѡ���Ƿ��ӡ"
         Height          =   180
         Index           =   2
         Left            =   465
         TabIndex        =   3
         Top             =   945
         Width           =   1380
      End
      Begin VB.OptionButton optChargePrint 
         Caption         =   "�Զ���ӡ"
         Height          =   180
         Index           =   1
         Left            =   465
         TabIndex        =   2
         Top             =   645
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmSquareAffirmParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mstrPrivs As String, mblnOk As Boolean
Private Const mlngModul = 1151
Public Function SetPara(ByVal frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ѵ���ز����������
    '����:���ȷ��,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-11 00:16:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnOk = False
    Me.Show 1, frmMain
    SetPara = mblnOk
End Function
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Function SavePara() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:���˺�
    '����:2011-08-10 23:37:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "�շѴ�ӡ��ʽ", IIf(optChargePrint(0).value, 0, IIf(optChargePrint(1).value, 1, 2)), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "��˴�ӡ��ʽ", IIf(optRecordPrint(0).value, 0, IIf(optRecordPrint(1).value, 1, 2)), glngSys, mlngModul, blnHavePrivs
    'zlDatabase.setPara "ҩƷ��λ", IIf(opt��λ(0).Value, 0, 1), glngSys, mlngModul, blnHavePrivs
    
    SavePara = True
End Function
 Private Sub cmdOK_Click()
    If SavePara = False Then Exit Sub
    mblnOk = True
    Unload Me
End Sub
Private Sub cmdPrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1151", Me)
End Sub

Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2011-08-10 23:48:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long
    Dim blnHavePrivs As Boolean, strValue As String
    Dim j As Long
    
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    i = Val(zlDatabase.GetPara("�շѴ�ӡ��ʽ", glngSys, mlngModul, , Array(optChargePrint(0), optChargePrint(1), optChargePrint(2)), blnHavePrivs))
    i = IIf(i < 0, 0, i): i = IIf(i > 2, 2, i)
    optChargePrint(i).value = True
    
    i = Val(zlDatabase.GetPara("��˴�ӡ��ʽ", glngSys, mlngModul, , Array(optRecordPrint(0), optRecordPrint(1), optRecordPrint(2)), blnHavePrivs))
    i = IIf(i < 0, 0, i): i = IIf(i > 2, 2, i)
    optRecordPrint(i).value = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    mstrPrivs = ";" & GetPrivFunc(glngSys, mlngModul) & ";"
    Call InitPara
End Sub

