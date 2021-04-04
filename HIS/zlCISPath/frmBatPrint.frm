VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmBatPrint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "路径表批量打印"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5745
   Icon            =   "frmBatPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdPrintSetup 
      Caption         =   "打印设置"
      Height          =   300
      Left            =   240
      TabIndex        =   7
      Top             =   1720
      Width           =   1100
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "开始打印"
      Default         =   -1  'True
      Height          =   300
      Left            =   2040
      TabIndex        =   6
      Top             =   1720
      Width           =   1100
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "预览(&V)"
      Height          =   300
      Left            =   3240
      TabIndex        =   5
      Top             =   1720
      Width           =   1100
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5160
      Top             =   360
   End
   Begin VB.CommandButton cmdCancle 
      Cancel          =   -1  'True
      Caption         =   "取消(&E)"
      Height          =   300
      Left            =   4440
      TabIndex        =   2
      Top             =   1720
      Width           =   1100
   End
   Begin VB.PictureBox picTime 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   240
      Picture         =   "frmBatPrint.frx":6852
      ScaleHeight     =   255
      ScaleWidth      =   5325
      TabIndex        =   1
      Top             =   975
      Width           =   5320
   End
   Begin VB.PictureBox picTime 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      Picture         =   "frmBatPrint.frx":719E
      ScaleHeight     =   255
      ScaleWidth      =   5325
      TabIndex        =   0
      Top             =   960
      Width           =   5320
   End
   Begin XtremeSuiteControls.TabControl tbcPath 
      Height          =   3090
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   5475
      _Version        =   589884
      _ExtentX        =   9657
      _ExtentY        =   5450
      _StockProps     =   64
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmBatPrint.frx":7A54
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblMsg 
      Caption         =   "共打印    个病人，正在打印第    个病人。"
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmBatPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmPath As frmPathTable
Private mfrmPathOut As frmPathTableOut
Private mrsTmp As Recordset
Private mlngI As Long
Private mlngCount As Long
Private mblnOk As Boolean
Private mblnPrinted As Boolean
Private mbytFunc As Byte

Public Sub ShowMe(ByVal frmParent As Object, ByVal rsTmp As ADODB.Recordset, Optional ByVal bytFunc As Byte)
'功能:bytFunc=0 0-临床路径跟踪;1-门诊临床路径跟踪
    On Error Resume Next
    Set mrsTmp = rsTmp
    mbytFunc = bytFunc
    Me.Show , frmParent
End Sub

Public Sub BatPrint()
'功能：批量打印
    mblnPrinted = False
    If mlngCount = 0 Then Unload Me: Exit Sub
    
    If Not mrsTmp.EOF Then
        lblMsg.Caption = "共打印 " & mlngCount & " 个病人," & "当前正在打印第 " & mlngI & " 个病人【" & mrsTmp!姓名 & "】。"
        If mbytFunc = 0 Then
            Call mfrmPath.zlRefresh(Val(mrsTmp!病人ID & ""), Val(mrsTmp!主页ID & ""), Val(mrsTmp!病区ID & ""), Val(mrsTmp!科室ID & ""), Val(mrsTmp!病人状态 & ""), Val(mrsTmp!数据转出 & "") = 1)
            Call mfrmPath.zlPrintOutPut(1, True)
        Else
            Call mfrmPathOut.zlRefresh(Val(mrsTmp!病人ID & ""), Val(mrsTmp!挂号ID & ""), mrsTmp!NO & "", Val(mrsTmp!科室ID & ""), Val(mrsTmp!病人状态 & ""), Val(mrsTmp!数据转出 & "") = 1)
            Call mfrmPathOut.zlPrintOutPut(1, True)
        End If
        picTime(1).Width = picTime(1).Width + (picTime(0).Width / mlngCount)
        Me.Refresh
        mlngI = mlngI + 1
    Else
        Unload Me
    End If
    mblnPrinted = True
End Sub

Private Sub cmdCancle_Click()
    '这里易用性考虑，正在打印时，按ESC不退出打印，而是先暂停打印
    If cmdStart.Tag = "Stop" Then
        Call cmdStart_Click
    Else
        If mlngI > 1 And mblnOk Then
            If MsgBox("已经开始打印，取消将停止打印后面的病人，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                Unload Me
            End If
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub cmdPreview_Click()
'预览当前记录集的病人路径表
    If Not mrsTmp.EOF Then
        If mbytFunc = 0 Then
            Call mfrmPath.zlRefresh(Val(mrsTmp!病人ID & ""), Val(mrsTmp!主页ID & ""), Val(mrsTmp!病区ID & ""), Val(mrsTmp!科室ID & ""), Val(mrsTmp!病人状态 & ""), Val(mrsTmp!数据转出 & "") = 1)
            Call mfrmPath.zlPrintOutPut(2, True)
        Else
            Call mfrmPathOut.zlRefresh(Val(mrsTmp!病人ID & ""), Val(mrsTmp!挂号ID & ""), mrsTmp!NO & "", Val(mrsTmp!科室ID & ""), Val(mrsTmp!病人状态 & ""), Val(mrsTmp!数据转出 & "") = 1)
            Call mfrmPathOut.zlPrintOutPut(2, True)
        End If
    Else
        MsgBox "没有可浏览的病人路径表。", vbInformation, Me.Caption
    End If
End Sub

Private Sub cmdPrintSetup_Click()
    '打印设置
    Call zlPrintSet
End Sub

Private Sub cmdStart_Click()
     '开始打印后按钮设置为暂停打印
    If cmdStart.Tag = "Start" Then
        Timer1.Enabled = True
        cmdStart.Tag = "Stop"
        cmdStart.Caption = "暂停打印"
        mblnOk = True
    Else
        Timer1.Enabled = False
        cmdStart.Tag = "Start"
        cmdStart.Caption = "开始打印"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And cmdStart.Tag = "Stop" Then
        Call cmdStart_Click
    End If
End Sub

Private Sub Form_Load()
    Dim tabItem As TabControlItem
    If mbytFunc = 0 Then
        Set mfrmPath = New frmPathTable
    Else
        Set mfrmPathOut = New frmPathTableOut
    End If
    picTime(1).Width = 0
    mlngI = 1
    mlngCount = mrsTmp.RecordCount
    mrsTmp.MoveFirst
    cmdStart.Tag = "Start"
    mblnOk = False
    mblnPrinted = True
    lblMsg.Caption = "共打印 " & mlngCount & " 个病人," & "是否要开始打印这些病人的路径表？"
    
    'TabControl
    '-----------------------------------------------------
    With Me.tbcPath
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .Color = xtpTabColorOffice2003
        End With
        If mbytFunc = 0 Then
            Set tabItem = .InsertItem(0, "病人临床路径", mfrmPath.Hwnd, 0)
        Else
            Set tabItem = .InsertItem(0, "病人门诊路径", mfrmPathOut.Hwnd, 0)
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mbytFunc = 0 Then
        Unload mfrmPath
    Else
        Unload mfrmPathOut
    End If
    Set mfrmPath = Nothing
    Set mrsTmp = Nothing
End Sub

Private Sub Timer1_Timer()
    '必须等上次打印完成后再开始下一个病人的打印
    If mblnPrinted Then
        Call BatPrint
        If Not mrsTmp Is Nothing Then mrsTmp.MoveNext
    End If
End Sub
