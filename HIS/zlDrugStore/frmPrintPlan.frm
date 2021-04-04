VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPrintPlan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ӡ����"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7635
   Icon            =   "frmPrintPlan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   7635
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ProgressBar prgPlan 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer timerAuto 
      Interval        =   200
      Left            =   6000
      Top             =   0
   End
   Begin VB.Label lblPlan 
      Caption         =   "����ɣ�20%"
      Height          =   180
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblCur 
      Caption         =   "�Ѵ�ӡ����20"
      Height          =   180
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblSum 
      Caption         =   "�ܴ�ӡ����100"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmPrintPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngCur As Long
Private mlngSum As Long
Private mlngTemp As Long
Private mstrPrint As String
Private mintNum As Integer
Private mIntCount As Integer
Private mlngRow As Long


Private Sub Form_Load()
    mlngCur = 0
    mlngTemp = 0
    Me.timerAuto.Enabled = True
End Sub

Private Sub timerAuto_Timer()
    Dim strPrintStatus As String
    Dim strJobStatus As String
    Dim blnReturn As Boolean
    Dim dateNow As Date
    Dim arrParams As Variant
    Dim lngRow As Long
    Dim strTemp As String
    Dim intCount As Integer
    Dim blnTemp As Boolean
    Dim i As Integer
    Dim intNum As Integer
    Dim j As Integer
    
    '����ӡ��״̬,�������ӡ��������ʾ�쳣
    
    Do While Not blnTemp
        blnReturn = CheckPrinter(strPrintStatus, strJobStatus)
        If blnReturn Then
            blnTemp = True
        Else
            If MsgBox("��ӡ�������쳣���Ƿ����ԣ�", vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                blnTemp = False
            Else
                blnTemp = True
                Unload Me
                Exit Sub
            End If
        End If
    Loop
    
    
    dateNow = zldatabase.Currentdate
    intNum = 20
    arrParams = Split(mstrPrint, ",")
    
    '�������ݣ�Ĭ��20��һ���ύ
    arrParams = Split(mstrPrint, ",")
    For lngRow = mlngCur To UBound(arrParams)
        strTemp = strTemp & Str(arrParams(lngRow)) & ","
        If arrParams(lngRow) <> "" And (lngRow + 1 = intNum * mIntCount Or lngRow + 1 = mlngSum) Then
            '���´�ӡ��־
            gstrSQL = "Zl_��Һ��ҩ��¼_��ӡ("
            '��ҩID
            gstrSQL = gstrSQL & "'" & strTemp & "'"
            '��ӡʱ��
            gstrSQL = gstrSQL & ",To_Date('" & dateNow & "','yyyy-MM-dd hh24:mi:ss')"
            gstrSQL = gstrSQL & ")"
            Call zldatabase.ExecuteProcedure(gstrSQL, "���´�ӡ��־")
            Exit For
        End If
    Next
    
    '���ô�ӡ���
    arrParams = GetArrayByStr(mstrPrint, 3950, ",")
    For lngRow = 0 To UBound(arrParams)
        '���´�ӡ���
        gstrSQL = "Zl_��Һ��ҩ��¼_�������("
        '��ҩID
        gstrSQL = gstrSQL & "'" & arrParams(lngRow) & "'"
        gstrSQL = gstrSQL & ",To_Date('" & dateNow & "','yyyy-MM-dd hh24:mi:ss')"
        gstrSQL = gstrSQL & ")"
        Call zldatabase.ExecuteProcedure(gstrSQL, "���´�ӡ���")
    Next
    
    For j = 0 To UBound(Split(strTemp, ",")) - 1
        If Split(strTemp, ",")(j) <> "" Then
            mlngCur = mlngCur + 1
            For i = 1 To mintNum
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1345_1", Me, _
                    "��ҩID=" & Val(Split(strTemp, ",")(j)), _
                    "PrintEmpty=0", 2)
                 Me.prgPlan.Value = Int((mlngCur * mintNum - mintNum + i) / (mlngSum * mintNum) * 100)
                 lblCur.Caption = "�Ѵ�ӡ����" & (mlngCur * mintNum - mintNum + i)
                 lblPlan.Caption = "����ɣ�" & Int((mlngCur * mintNum - mintNum + i) / (mlngSum * mintNum) * 100) & "%"
    
            Next
            
        End If
    Next
    
'    mlngCur = lngRow
    '���������
    
    DoEvents
    strTemp = ""
    mIntCount = mIntCount + 1
    
    If mlngSum = mlngCur Then
        Unload Me
        Exit Sub
    End If
    
    Sleep (5000)
End Sub

Public Sub ShowMe(ByVal frmParent As Form, ByVal strPrint As String, ByVal intNum As Integer)
    'frmParent:������
    'strPrint:��ӡ����ҩID��
    'intNum:��ҩ����ӡ����
    mintNum = intNum
    mstrPrint = strPrint
    mIntCount = 1
    mlngSum = UBound(Split(strPrint, ",")) + 1
    lblSum.Caption = "�ܴ�ӡ����" & mlngSum * mintNum
    lblCur.Caption = "�Ѵ�ӡ����0"
    lblPlan.Caption = "����ɣ�0%"
    
    Me.Show 1, frmParent
End Sub





