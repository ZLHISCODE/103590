VERSION 5.00
Begin VB.Form frmDiffPriceRecalCard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "卫材差价计算"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   Icon            =   "frmDiffPriceRecalCard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra 
      Height          =   30
      Left            =   -30
      TabIndex        =   7
      Top             =   900
      Width           =   5805
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   150
      TabIndex        =   6
      Top             =   2385
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   50
      Left            =   -765
      TabIndex        =   5
      Top             =   2145
      Width           =   7815
   End
   Begin VB.ComboBox cboPeriod 
      Height          =   300
      Left            =   1545
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1350
      Width           =   2925
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4140
      TabIndex        =   1
      Top             =   2385
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2835
      TabIndex        =   0
      Top             =   2385
      Width           =   1100
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   120
      Picture         =   "frmDiffPriceRecalCard.frx":000C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblPeriod 
      AutoSize        =   -1  'True
      Caption         =   "期间"
      Height          =   180
      Left            =   1140
      TabIndex        =   4
      Top             =   1425
      Width           =   360
   End
   Begin VB.Label lblMemo 
      Caption         =   "材料差价重算，将会导致原有材料差价数据的变化，且不可恢复，因此建议不要对一般人员授予该权限，以免导致对数据的影响。"
      ForeColor       =   &H00C00000&
      Height          =   600
      Left            =   735
      TabIndex        =   2
      Top             =   345
      Width           =   4890
   End
End
Attribute VB_Name = "frmDiffPriceRecalCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub CmdSave_Click()
    Dim strPeriod As String
    
    If cboPeriod.ListIndex = -1 Then Exit Sub
    On Error GoTo ErrHandle
    strPeriod = Mid(cboPeriod.Text, 1, 4) & Mid(cboPeriod.Text, 6, 2)
          
    DoEvents
    FS.ShowFlash "正在计算和更新，请等待。。。"
    gstrSQL = "zl_材料差价重整_UPDATE('" & strPeriod & "')"
    Me.MousePointer = vbHourglass
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Me.MousePointer = vbDefault
   
    MsgBox "差价重算成功！", vbOKOnly + vbInformation, gstrSysName
    FS.StopFlash
    Unload Me
    Exit Sub

ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Command1_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub

Private Sub Form_Load()
    Dim rsPeriod As New Recordset
    
    RestoreWinState Me, App.Title
    
    On Error GoTo ErrHandle
    gstrSQL = "select 期间 from 期间表 WHERE 期间>TO_CHAR(ADD_MONTHS(SYSDATE,-2),'yyyymm') AND 开始日期<SYSDATE "
    Call zlDatabase.OpenRecordset(rsPeriod, gstrSQL, Me.Caption)
    
    If rsPeriod.EOF Then
        MsgBox "没有设置相应的期间!", vbOKOnly, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    With cboPeriod
        .Clear
        Do While Not rsPeriod.EOF
            .AddItem Mid(rsPeriod.Fields(0), 1, 4) & "年" & Mid(rsPeriod.Fields(0), 5) & "月"
            rsPeriod.MoveNext
        Loop
        .ListIndex = .ListCount - 1
        rsPeriod.Close
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.Title
End Sub
