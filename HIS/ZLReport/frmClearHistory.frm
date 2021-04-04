VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClearHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "清除历史数据源"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "frmClearHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1430
      MaxLength       =   10
      TabIndex        =   6
      Text            =   "5"
      Top             =   1470
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   300
      Left            =   1080
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   44761091
      CurrentDate     =   41634
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   0
      TabIndex        =   4
      Top             =   2400
      Width           =   5910
   End
   Begin VB.OptionButton optType 
      Caption         =   "清除                之前修改的历史记录"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   1890
      Width           =   3975
   End
   Begin VB.OptionButton optType 
      Caption         =   "清除最近            次修改之前的历史记录"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Value           =   -1  'True
      Width           =   4095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2040
      TabIndex        =   1
      Top             =   2520
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3270
      TabIndex        =   0
      Top             =   2520
      Width           =   1100
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   600
      Picture         =   "frmClearHistory.frx":6852
      Stretch         =   -1  'True
      Top             =   760
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "   注意：删除历史数据源记录后，将不可恢复，请慎重选择！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "    请选择一种模式来清除历史数据源，请注意：清除历史数据源记录是针对所有报表的！"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmClearHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    On Error GoTo errH
    If optType(0).Value Then
        '按最新N次删除
        If MsgBox("是否删除最近" & Val(txtCount.Text) & "次之前的历史数据源记录？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
            gcnOracle.Execute "Delete From Zlrptsqlshistory A" & vbNewLine & _
                                "Where (a.报表id, a.修改时间) Not In (Select 报表id, 修改时间" & vbNewLine & _
                                "                           From (Select 报表id, 修改时间, Row_Number() Over(Partition By 报表id Order By 报表id, 修改时间 Desc) As Top" & vbNewLine & _
                                "                                  From (Select b.报表id, b.修改时间" & vbNewLine & _
                                "                                         From Zlrptsqlshistory B" & vbNewLine & _
                                "                                         Group By b.修改时间, b.报表id" & vbNewLine & _
                                "                                         Order By b.报表id, b.修改时间 Desc)) B" & vbNewLine & _
                                "                           Where Top <= " & Val(txtCount.Text) & ")"
            MsgBox "删除成功。", vbInformation, Me.Caption
        End If
    Else
        '按时间删除
        If MsgBox("是否删除" & Format(dtpTime.Value, "yyyy-mm-dd") & "之前的历史数据源记录？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
            gcnOracle.Execute "Delete From Zlrptsqlshistory A" & vbNewLine & _
                                "Where A.修改时间< To_date('" & Format(dtpTime.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')"
            MsgBox "删除成功。", vbInformation, Me.Caption
        End If
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    dtpTime.Value = DateAdd("M", -1, Currentdate)
End Sub

Private Sub optType_Click(Index As Integer)
     dtpTime.Enabled = optType(1).Value
     txtCount.Enabled = optType(0).Value
End Sub
