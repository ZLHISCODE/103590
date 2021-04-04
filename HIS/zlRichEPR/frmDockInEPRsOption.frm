VERSION 5.00
Begin VB.Form frmDockInEPRsOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病历选项"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   Icon            =   "frmDockInEPRsOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4365
      TabIndex        =   2
      Top             =   3375
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3000
      TabIndex        =   1
      Top             =   3375
      Width           =   1125
   End
   Begin VB.Frame fraDockInEprs 
      Height          =   3165
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   5445
      Begin VB.TextBox txtDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         Left            =   3240
         TabIndex        =   13
         Text            =   "7"
         Top             =   2805
         Width           =   360
      End
      Begin VB.OptionButton optRead 
         Caption         =   "连续预览，读取选中文件前后    天的共享病历。"
         Height          =   195
         Index           =   2
         Left            =   630
         TabIndex        =   12
         Top             =   2790
         Width           =   4305
      End
      Begin VB.OptionButton optRead 
         Caption         =   "不连续预览，选中一份文件读一次。"
         Height          =   195
         Index           =   1
         Left            =   630
         TabIndex        =   11
         Top             =   2470
         Width           =   4305
      End
      Begin VB.OptionButton optRead 
         Caption         =   "连续预览，首次读取全部共享病历，后续只定位。"
         Height          =   195
         Index           =   0
         Left            =   630
         TabIndex        =   9
         Top             =   2150
         Value           =   -1  'True
         Width           =   4305
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   30
         TabIndex        =   6
         Top             =   1725
         Width           =   5400
      End
      Begin VB.CheckBox chkShareWrited 
         Caption         =   "共享病历必须先书写被共享病历"
         Height          =   180
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   3720
      End
      Begin VB.CheckBox chkAutoShowNewPane 
         Caption         =   "自动显示新增面板"
         Height          =   180
         Left            =   360
         TabIndex        =   4
         Top             =   705
         Width           =   3720
      End
      Begin VB.CheckBox chkPageprogression 
         Caption         =   "(转科后要求书写)的共享病历另起一页打印"
         Height          =   330
         Left            =   360
         TabIndex        =   3
         Top             =   300
         Width           =   3720
      End
      Begin VB.TextBox txtfolding 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   2880
         TabIndex        =   7
         Text            =   "6"
         Top             =   1440
         Width           =   360
      End
      Begin VB.Line Line2 
         X1              =   3225
         X2              =   3615
         Y1              =   2985
         Y2              =   2985
      End
      Begin VB.Label Label2 
         Caption         =   "共享病历连读预览"
         Height          =   180
         Left            =   360
         TabIndex        =   10
         Top             =   1845
         Width           =   1650
      End
      Begin VB.Line Line1 
         X1              =   2865
         X2              =   3255
         Y1              =   1620
         Y2              =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "列表产生滚动条后，共享病历超    行自动折叠"
         Height          =   240
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   3780
      End
   End
End
Attribute VB_Name = "frmDockInEPRsOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModul As Long, mstrPrivs As String
Public Sub ShowMe(ByVal lngModul As Long, ByVal strPrivs As String)
    mlngModul = lngModul: mstrPrivs = strPrivs
    Me.Show 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Call zlDatabase.SetPara("转科后要求书写的共享病历另起一页打印", chkPageprogression.Value, glngSys, mlngModul)
    Call zlDatabase.SetPara("自动显示新增面板", chkAutoShowNewPane.Value, glngSys, mlngModul)
    Call zlDatabase.SetPara("共享病历必须先书写被共享病历", chkShareWrited.Value, glngSys, mlngModul)
    Call zlDatabase.SetPara("共享病历折叠起始行数", Abs(txtfolding.Text), glngSys, mlngModul)
    Select Case True
        Case optRead(0).Value
            Call zlDatabase.SetPara("共享病历连读预览", "-1", glngSys, mlngModul)
        Case optRead(1).Value
            Call zlDatabase.SetPara("共享病历连读预览", "0", glngSys, mlngModul)
        Case optRead(2).Value
            Call zlDatabase.SetPara("共享病历连读预览", Abs(txtDays.Text), glngSys, mlngModul)
    End Select
    Unload Me
End Sub

Private Sub Form_Load()
Dim lngDays As Long '-1表示共享病历全部读取 0表示仅读当前选中病历 >0表示读取选中病历前后N天内的共享病历
    chkAutoShowNewPane.Value = zlDatabase.GetPara("自动显示新增面板", glngSys, mlngModul, "1", Array(chkAutoShowNewPane), InStr(mstrPrivs, "参数设置") > 0)
    chkPageprogression.Value = zlDatabase.GetPara("转科后要求书写的共享病历另起一页打印", glngSys, mlngModul, "1", Array(chkPageprogression), InStr(mstrPrivs, "参数设置") > 0)
    chkShareWrited.Value = zlDatabase.GetPara("共享病历必须先书写被共享病历", glngSys, mlngModul, "1", Array(chkShareWrited), InStr(mstrPrivs, "参数设置") > 0)
    txtfolding.Text = zlDatabase.GetPara("共享病历折叠起始行数", glngSys, mlngModul, "6", Array(txtfolding, Label1), InStr(mstrPrivs, "参数设置") > 0)
    lngDays = zlDatabase.GetPara("共享病历连读预览", glngSys, 1251, -1, Array(optRead(0), optRead(1), optRead(2), txtDays), InStr(mstrPrivs, "参数设置") > 0)
    Select Case lngDays
        Case -1
            optRead(0).Value = True
        Case 0
            optRead(1).Value = True
        Case Else
            optRead(2).Value = True
            txtDays.Text = lngDays
    End Select
End Sub
Private Sub optRead_Click(Index As Integer)
    If Index <> 2 Then
        txtDays.Enabled = False
    Else
        txtDays.Enabled = True
    End If
End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtfolding_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
