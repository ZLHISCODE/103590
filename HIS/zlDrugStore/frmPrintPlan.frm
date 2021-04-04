VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPrintPlan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "打印进度"
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
   StartUpPosition =   3  '窗口缺省
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
      Caption         =   "已完成：20%"
      Height          =   180
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblCur 
      Caption         =   "已打印数：20"
      Height          =   180
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblSum 
      Caption         =   "总打印数：100"
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
    
    '检查打印机状态,正常则打印，否则提示异常
    
    Do While Not blnTemp
        blnReturn = CheckPrinter(strPrintStatus, strJobStatus)
        If blnReturn Then
            blnTemp = True
        Else
            If MsgBox("打印机出现异常，是否重试？", vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
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
    
    '处理数据，默认20个一次提交
    arrParams = Split(mstrPrint, ",")
    For lngRow = mlngCur To UBound(arrParams)
        strTemp = strTemp & Str(arrParams(lngRow)) & ","
        If arrParams(lngRow) <> "" And (lngRow + 1 = intNum * mIntCount Or lngRow + 1 = mlngSum) Then
            '更新打印标志
            gstrSQL = "Zl_输液配药记录_打印("
            '配药ID
            gstrSQL = gstrSQL & "'" & strTemp & "'"
            '打印时间
            gstrSQL = gstrSQL & ",To_Date('" & dateNow & "','yyyy-MM-dd hh24:mi:ss')"
            gstrSQL = gstrSQL & ")"
            Call zldatabase.ExecuteProcedure(gstrSQL, "更新打印标志")
            Exit For
        End If
    Next
    
    '设置打印序号
    arrParams = GetArrayByStr(mstrPrint, 3950, ",")
    For lngRow = 0 To UBound(arrParams)
        '更新打印序号
        gstrSQL = "Zl_输液配药记录_设置序号("
        '配药ID
        gstrSQL = gstrSQL & "'" & arrParams(lngRow) & "'"
        gstrSQL = gstrSQL & ",To_Date('" & dateNow & "','yyyy-MM-dd hh24:mi:ss')"
        gstrSQL = gstrSQL & ")"
        Call zldatabase.ExecuteProcedure(gstrSQL, "更新打印序号")
    Next
    
    For j = 0 To UBound(Split(strTemp, ",")) - 1
        If Split(strTemp, ",")(j) <> "" Then
            mlngCur = mlngCur + 1
            For i = 1 To mintNum
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1345_1", Me, _
                    "配药ID=" & Val(Split(strTemp, ",")(j)), _
                    "PrintEmpty=0", 2)
                 Me.prgPlan.Value = Int((mlngCur * mintNum - mintNum + i) / (mlngSum * mintNum) * 100)
                 lblCur.Caption = "已打印数：" & (mlngCur * mintNum - mintNum + i)
                 lblPlan.Caption = "已完成：" & Int((mlngCur * mintNum - mintNum + i) / (mlngSum * mintNum) * 100) & "%"
    
            Next
            
        End If
    Next
    
'    mlngCur = lngRow
    '处理进度条
    
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
    'frmParent:父窗体
    'strPrint:打印的配药ID串
    'intNum:配药单打印次数
    mintNum = intNum
    mstrPrint = strPrint
    mIntCount = 1
    mlngSum = UBound(Split(strPrint, ",")) + 1
    lblSum.Caption = "总打印数：" & mlngSum * mintNum
    lblCur.Caption = "已打印数：0"
    lblPlan.Caption = "已完成：0%"
    
    Me.Show 1, frmParent
End Sub





