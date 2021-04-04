VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmVersionOut 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "临床路径"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4860
   Icon            =   "frmVersionOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra标准费用 
      Caption         =   "标准费用"
      Height          =   1140
      Left            =   495
      TabIndex        =   24
      Top             =   2190
      Width           =   3885
      Begin VB.TextBox txt费用 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   825
         MaxLength       =   10
         TabIndex        =   9
         Top             =   300
         Width           =   1080
      End
      Begin VB.OptionButton opt费用 
         Height          =   210
         Index           =   1
         Left            =   270
         TabIndex        =   10
         Top             =   735
         Width           =   210
      End
      Begin VB.OptionButton opt费用 
         Height          =   210
         Index           =   0
         Left            =   270
         TabIndex        =   8
         Top             =   345
         Value           =   -1  'True
         Width           =   210
      End
      Begin VB.TextBox txt费用 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   840
         MaxLength       =   10
         TabIndex        =   11
         Top             =   690
         Width           =   1080
      End
      Begin VB.TextBox txt费用 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2355
         MaxLength       =   10
         TabIndex        =   12
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label lblCost 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "估算(Ctrl+E)"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2385
         MouseIcon       =   "frmVersionOut.frx":058A
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≤             元"
         Height          =   180
         Index           =   3
         Left            =   615
         TabIndex        =   26
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "元 -             元"
         Height          =   180
         Index           =   0
         Left            =   1965
         TabIndex        =   25
         Top             =   750
         Width           =   1710
      End
   End
   Begin VB.Frame fra就诊日 
      Caption         =   "标准治疗时间"
      Height          =   1140
      Left            =   495
      TabIndex        =   21
      Top             =   990
      Width           =   3885
      Begin VB.OptionButton opt天数 
         Height          =   210
         Index           =   1
         Left            =   270
         TabIndex        =   3
         Top             =   705
         Width           =   210
      End
      Begin VB.OptionButton opt天数 
         Height          =   210
         Index           =   0
         Left            =   270
         TabIndex        =   0
         Top             =   345
         Value           =   -1  'True
         Width           =   210
      End
      Begin MSComCtl2.UpDown ud天数 
         Height          =   300
         Index           =   2
         Left            =   3210
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt天数(2)"
         BuddyDispid     =   196616
         BuddyIndex      =   2
         OrigLeft        =   2265
         OrigTop         =   1815
         OrigRight       =   2520
         OrigBottom      =   2010
         Max             =   999
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown ud天数 
         Height          =   300
         Index           =   1
         Left            =   1680
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   660
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt天数(1)"
         BuddyDispid     =   196616
         BuddyIndex      =   1
         OrigLeft        =   2265
         OrigTop         =   1815
         OrigRight       =   2520
         OrigBottom      =   2010
         Max             =   999
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown ud天数 
         Height          =   300
         Index           =   0
         Left            =   1680
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt天数(0)"
         BuddyDispid     =   196616
         BuddyIndex      =   0
         OrigLeft        =   2265
         OrigTop         =   1815
         OrigRight       =   2520
         OrigBottom      =   2010
         Max             =   999
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt天数 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   825
         MaxLength       =   3
         TabIndex        =   4
         Top             =   660
         Width           =   840
      End
      Begin VB.TextBox txt天数 
         Alignment       =   2  'Center
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   825
         MaxLength       =   3
         TabIndex        =   1
         Top             =   300
         Width           =   840
      End
      Begin VB.TextBox txt天数 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2355
         MaxLength       =   3
         TabIndex        =   6
         Top             =   675
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≤             天"
         Height          =   180
         Index           =   2
         Left            =   615
         TabIndex        =   23
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "天 -             天"
         Height          =   180
         Index           =   1
         Left            =   1965
         TabIndex        =   22
         Top             =   720
         Width           =   1710
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2145
      TabIndex        =   15
      Top             =   4365
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3240
      TabIndex        =   16
      Top             =   4365
      Width           =   1100
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   4860
      TabIndex        =   18
      Top             =   0
      Width           =   4860
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路径版本信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1065
         TabIndex        =   20
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "  设置当前临床路径版本的标准治疗时间、标准费用，以及说明信息。"
         Height          =   360
         Left            =   1065
         TabIndex        =   19
         Top             =   360
         Width           =   3480
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   105
         Picture         =   "frmVersionOut.frx":06DC
         Top             =   45
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   10000
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.TextBox txt说明 
      Height          =   660
      Left            =   1065
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   3450
      Width           =   3315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   10000
      Y1              =   4230
      Y2              =   4230
   End
   Begin VB.Label lbl说明 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "说明"
      Height          =   180
      Left            =   585
      TabIndex        =   17
      Top             =   3510
      Width           =   360
   End
End
Attribute VB_Name = "frmVersionOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event CheckDataValid(Version As TYPE_PATH_VERSION, Cancel As Boolean)
Public Event CalcPathCost(CostMin As Currency, CostMax As Currency)

Private mvVersion       As TYPE_PATH_VERSION
Private mblnOK          As Boolean
Private mlng路径ID      As Long
Private mlngBegin       As Long
Private mlngPreStepID   As Long
Private mcolBegin       As Collection

Public Function ShowMe(frmParent As Object, vVersion As TYPE_PATH_VERSION, ByVal lng路径ID As Long, Optional ByVal lngNew阶段ID As Long) As Boolean
    mvVersion = vVersion
    mlng路径ID = lng路径ID
    mlngPreStepID = lngNew阶段ID
    Me.Show 1, frmParent
    If mblnOK Then vVersion = mvVersion
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim blnCancel As Boolean

    '检查数据
    If opt天数(0).Value Then
        If txt天数(0).Text = "" Or Val(txt天数(0).Text) <= 0 Then
            MsgBox "请输入一个有效的天数值。", vbInformation, gstrSysName
            txt天数(0).SetFocus: Exit Sub
        End If
    ElseIf opt天数(1).Value Then
        If txt天数(1).Text = "" Or Val(txt天数(1).Text) <= 0 Then
            MsgBox "请输入一个有效的天数值。", vbInformation, gstrSysName
            txt天数(1).SetFocus: Exit Sub
        End If
        If txt天数(2).Text = "" Or Val(txt天数(2).Text) <= 0 Then
            MsgBox "请输入一个有效的天数值。", vbInformation, gstrSysName
            txt天数(2).SetFocus: Exit Sub
        End If
        If Val(txt天数(2).Text) <= Val(txt天数(1).Text) Then
            MsgBox "最高天数应该大于(>)最低天数。", vbInformation, gstrSysName
            txt天数(2).SetFocus: Exit Sub
        End If
    End If
    '标准费用可以不设置
    If opt费用(0).Value Then
        If txt费用(0).Text <> "" And Val(txt费用(0).Text) <= 0 Then
            MsgBox "请输入一个有效的费用值。", vbInformation, gstrSysName
            txt费用(0).SetFocus: Exit Sub
        End If
    ElseIf opt费用(1).Value Then
        If txt费用(1).Text <> "" And Val(txt费用(1).Text) <= 0 Then
            MsgBox "请输入一个有效的费用值。", vbInformation, gstrSysName
            txt费用(1).SetFocus: Exit Sub
        End If
        If txt费用(2).Text <> "" And Val(txt费用(2).Text) <= 0 Then
            MsgBox "请输入一个有效的费用值。", vbInformation, gstrSysName
            txt费用(2).SetFocus: Exit Sub
        End If
        If txt费用(1).Text <> "" And txt费用(2).Text = "" _
            Or txt费用(1).Text = "" And txt费用(2).Text <> "" Then
            MsgBox "请输入一个有效的费用值。", vbInformation, gstrSysName
            If txt费用(2).Text = "" Then txt费用(2).SetFocus
            If txt费用(1).Text = "" Then txt费用(1).SetFocus
            Exit Sub
        End If
        If Val(txt费用(2).Text) <= Val(txt费用(1).Text) Then
            MsgBox "最高费用应该高于(>)最低费用。", vbInformation, gstrSysName
            txt费用(2).SetFocus: Exit Sub
        End If
    End If
    If zlCommFun.ActualLen(txt说明.Text) > txt说明.MaxLength Then
        MsgBox "说明内容太长，最多只允许 " & txt说明.MaxLength \ 2 & " 个汉字或者 " & txt说明.MaxLength & " 个字符。", vbInformation, gstrSysName
        txt说明.SetFocus: Exit Sub
    End If
    
    '收集数据
    If opt天数(0).Value Then
        mvVersion.标准治疗时间 = txt天数(0).Text
    ElseIf opt天数(1).Value Then
        mvVersion.标准治疗时间 = txt天数(1).Text & "-" & txt天数(2).Text
    End If

    If opt费用(0).Value Then
        mvVersion.标准费用 = txt费用(0).Text

    ElseIf opt费用(1).Value Then
        mvVersion.标准费用 = txt费用(1).Text & "-" & txt费用(2).Text
    End If
    mvVersion.版本说明 = txt说明.Text
    
    RaiseEvent CheckDataValid(mvVersion, blnCancel)
    If blnCancel Then Exit Sub
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyE And Shift = vbCtrlMask Then
        If lblCost.Visible And lblCost.Enabled Then Call lblCost_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim strSql As String, rsTmp As Recordset
    
    mblnOK = False

    If mvVersion.标准治疗时间 <> "" Then
        If InStr(mvVersion.标准治疗时间, "-") = 0 Then
            opt天数(0).Value = True
            txt天数(0).Text = mvVersion.标准治疗时间
        Else
            opt天数(1).Value = True
            txt天数(1).Text = Split(mvVersion.标准治疗时间, "-")(0)
            txt天数(2).Text = Split(mvVersion.标准治疗时间, "-")(1)
        End If
    End If
    
    If mvVersion.标准费用 <> "" Then
        If InStr(mvVersion.标准费用, "-") = 0 Then
            opt费用(0).Value = True
            txt费用(0).Text = mvVersion.标准费用
        Else
            opt费用(1).Value = True
            txt费用(1).Text = Split(mvVersion.标准费用, "-")(0)
            txt费用(2).Text = Split(mvVersion.标准费用, "-")(1)
        End If
    End If
    txt说明.Text = mvVersion.版本说明
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngBegin = 0
End Sub

Private Sub lblCost_Click()
    Dim curCostMin As Currency
    Dim curCostMax As Currency
    
    RaiseEvent CalcPathCost(curCostMin, curCostMax)
    If curCostMin <> 0 And curCostMax <> 0 Then
        If curCostMin = curCostMax Then
            opt费用(0).Value = True
            txt费用(0).Text = IntEx(curCostMin)
            txt费用(0).SetFocus
            Call txt费用_GotFocus(0)
        Else
            opt费用(1).Value = True
            txt费用(1).Text = IntEx(curCostMin)
            txt费用(2).Text = IntEx(curCostMax)
            txt费用(1).SetFocus
            Call txt费用_GotFocus(1)
        End If
    End If
End Sub

Private Sub opt费用_Click(Index As Integer)
    If opt费用(0).Value Then
        txt费用(0).Enabled = True
        txt费用(1).Enabled = False: txt费用(2).Enabled = False
        
        txt费用(0).BackColor = txt说明.BackColor
        txt费用(1).BackColor = Me.BackColor
        txt费用(2).BackColor = Me.BackColor
    ElseIf opt费用(1).Value Then
        txt费用(0).Enabled = False
        txt费用(1).Enabled = True: txt费用(2).Enabled = True
        
        txt费用(0).BackColor = Me.BackColor
        txt费用(1).BackColor = txt说明.BackColor
        txt费用(2).BackColor = txt说明.BackColor
    End If
End Sub

Private Sub opt天数_Click(Index As Integer)
    If opt天数(0).Value Then
        txt天数(0).Enabled = True: ud天数(0).Enabled = True
        txt天数(1).Enabled = False: txt天数(2).Enabled = False
        ud天数(1).Enabled = False: ud天数(2).Enabled = False
        
        txt天数(0).BackColor = txt说明.BackColor
        txt天数(1).BackColor = Me.BackColor
        txt天数(2).BackColor = Me.BackColor
    ElseIf opt天数(1).Value Then
        txt天数(0).Enabled = False: ud天数(0).Enabled = False
        txt天数(1).Enabled = True: txt天数(2).Enabled = True
        ud天数(1).Enabled = True: ud天数(2).Enabled = True
        
        txt天数(0).BackColor = Me.BackColor
        txt天数(1).BackColor = txt说明.BackColor
        txt天数(2).BackColor = txt说明.BackColor
    End If
End Sub

Private Sub txt费用_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt费用(Index))
End Sub

Private Sub txt费用_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt说明_GotFocus()
    Call zlControl.TxtSelAll(txt说明)
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txt天数_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt天数(Index))
End Sub

Private Sub txt天数_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
