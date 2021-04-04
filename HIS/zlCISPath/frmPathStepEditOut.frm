VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPathStepEditOut 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "阶段设置"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmPathStepEditOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt天数 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2835
      MaxLength       =   3
      TabIndex        =   10
      Top             =   1020
      Width           =   420
   End
   Begin VB.TextBox txt天数 
      Alignment       =   2  'Center
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1545
      MaxLength       =   3
      TabIndex        =   9
      Top             =   1005
      Width           =   420
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2040
      TabIndex        =   7
      Top             =   3690
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3135
      TabIndex        =   6
      Top             =   3690
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
      ScaleWidth      =   4755
      TabIndex        =   3
      Top             =   0
      Width           =   4755
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "时间阶段"
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
         TabIndex        =   5
         Top             =   120
         Width           =   780
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "  设置路径表中的一个时间阶段，可以是具体某个天数，也可以是一个天数范围。"
         Height          =   360
         Left            =   1065
         TabIndex        =   4
         Top             =   360
         Width           =   3240
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   105
         Picture         =   "frmPathStepEditOut.frx":058A
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
   Begin VB.TextBox txt名称 
      Alignment       =   2  'Center
      Height          =   660
      Left            =   1200
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "换行：Ctrl+Enter"
      Top             =   1560
      Width           =   2880
   End
   Begin VB.TextBox txt说明 
      Height          =   660
      Left            =   1200
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2760
      Width           =   2880
   End
   Begin VB.ComboBox cbo分类 
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   2340
      Width           =   2880
   End
   Begin MSComCtl2.UpDown ud天数 
      Height          =   300
      Index           =   1
      Left            =   3255
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1020
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txt天数(1)"
      BuddyDispid     =   196627
      BuddyIndex      =   1
      OrigLeft        =   2265
      OrigTop         =   1815
      OrigRight       =   2520
      OrigBottom      =   2010
      Max             =   999
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   0   'False
   End
   Begin MSComCtl2.UpDown ud天数 
      Height          =   300
      Index           =   0
      Left            =   1965
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1005
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txt天数(0)"
      BuddyDispid     =   196627
      BuddyIndex      =   0
      OrigLeft        =   2265
      OrigTop         =   1815
      OrigRight       =   2520
      OrigBottom      =   2010
      Max             =   999
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label lbl天数 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "天数："
      Height          =   180
      Left            =   720
      TabIndex        =   17
      Top             =   1065
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "第"
      Height          =   180
      Index           =   0
      Left            =   1290
      TabIndex        =   16
      Top             =   1065
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "天  -         天"
      Height          =   180
      Index           =   1
      Left            =   2310
      TabIndex        =   15
      Top             =   1065
      Width           =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   10000
      Y1              =   3555
      Y2              =   3555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   10000
      Y1              =   3570
      Y2              =   3570
   End
   Begin VB.Label lbl名称 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "名称"
      Height          =   180
      Left            =   720
      TabIndex        =   14
      Top             =   1620
      Width           =   360
   End
   Begin VB.Label lbl说明 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "说明"
      Height          =   180
      Left            =   720
      TabIndex        =   13
      Top             =   2820
      Width           =   360
   End
   Begin VB.Label lbl分类 
      Caption         =   "分类"
      Height          =   180
      Left            =   720
      TabIndex        =   12
      Top             =   2400
      Width           =   540
   End
End
Attribute VB_Name = "frmPathStepEditOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event CheckDataValid(TimeStep As TYPE_PATH_STEP, Cancel As Boolean)

Private mvStep      As TYPE_PATH_STEP
Private mvPreStep   As TYPE_PATH_STEP
Private mvNextStep  As TYPE_PATH_STEP
Private mstr分类s   As String
Private mblnOK      As Boolean

Public Function ShowEdit(frmParent As Object, vStep As TYPE_PATH_STEP, _
    vPreStep As TYPE_PATH_STEP, vNextStep As TYPE_PATH_STEP, ByVal str分类s As String) As Boolean
'功能：设置当前选择时间阶段的详细内容
'参数：vStep=主要是修改时当前时间阶段的内容，类型中的"父ID<>0"表示设置分支
'      mvPreStep,mvNextStep=前后相邻的一个时间阶段的内容，用于新增时参考
'      str分类s=当前路径表中，前后阶段备用分支的分类名串，用"|"间隔
    mvStep = vStep
    mvPreStep = vPreStep
    mvNextStep = vNextStep
    mstr分类s = str分类s
    
    Me.Show 1, frmParent
    
    If mblnOK Then vStep = mvStep
    ShowEdit = mblnOK
End Function

Private Sub cbo分类_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "|" Then KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim blnCancel As Boolean
    Dim strTmp As String, i As Integer
    
    '检查数据
    If txt天数(0).Text <> "" And Val(txt天数(0).Text) <= 0 Then
        MsgBox "请输入一个有效的开始天数值。", vbInformation, gstrSysName
        txt天数(0).SetFocus: Exit Sub
    End If
    If txt天数(1).Text <> "" And Val(txt天数(1).Text) <= 0 Then
        MsgBox "请输入一个有效的结束天数值。", vbInformation, gstrSysName
        txt天数(0).SetFocus: Exit Sub
    End If
    If txt天数(0).Text <> "" And txt天数(1).Text <> "" Then
        If Val(txt天数(1).Text) < Val(txt天数(0).Text) Then
            MsgBox "结束天数应该大于开始天数。", vbInformation, gstrSysName
            txt天数(1).SetFocus: Exit Sub
        ElseIf Val(txt天数(0).Text) = Val(txt天数(1).Text) Then
            MsgBox "指定为某一个天数时，不需要输入结束天数。", vbInformation, gstrSysName
            txt天数(1).SetFocus: Exit Sub
        End If
    End If
    If txt天数(1).Text <> "" And txt天数(0).Text = "" Then
        MsgBox "请输入开始天数。", vbInformation, gstrSysName
        txt天数(0).SetFocus: Exit Sub
    End If
    
    If Trim(txt名称.Text) = "" Then
        MsgBox "请输入时间阶段的名称。", vbInformation, gstrSysName
        txt名称.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txt名称.Text) > txt名称.MaxLength Then
        MsgBox "名称内容太长，最多允许 " & txt名称.MaxLength \ 2 & " 个汉字或者 " & txt名称.MaxLength & " 个字符。", vbInformation, gstrSysName
        txt名称.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(cbo分类.Text) > 50 Then
        MsgBox "分类内容太长，最多允许 25 个汉字或者 50 个字符。", vbInformation, gstrSysName
        cbo分类.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txt说明.Text) > txt说明.MaxLength Then
        MsgBox "说明内容太长，最多允许 " & txt说明.MaxLength \ 2 & " 个汉字或者 " & txt说明.MaxLength & " 个字符。", vbInformation, gstrSysName
        txt说明.SetFocus: Exit Sub
    End If
    
    '收集数据
    mvStep.名称 = txt名称.Text
    mvStep.说明 = txt说明.Text
    mvStep.开始天数 = Val(txt天数(0).Text)
    mvStep.结束天数 = Val(txt天数(1).Text)
    mvStep.分类 = cbo分类.Text
        
    '主程序检查
    If mvStep.父ID = 0 Then
        RaiseEvent CheckDataValid(mvStep, blnCancel)
        If blnCancel Then Exit Sub
        '允许不指定天数范围
        If txt天数(0).Text = "" And txt天数(1).Text = "" And txt天数(0).Enabled Then
            If MsgBox("没有确定该时间阶段所对应的天数范围，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    mblnOK = False
    
    txt名称.Text = mvStep.名称
    cbo分类.Text = mvStep.分类
    txt说明.Text = mvStep.说明
    txt天数(0).Text = IIf(mvStep.开始天数 = 0, "", mvStep.开始天数)
    txt天数(1).Text = IIf(mvStep.结束天数 = 0, "", mvStep.结束天数)
    
    '新设置时，根据前一个阶段的天数范围进行缺省
    If mvStep.名称 = "" Then
        If mvPreStep.名称 <> "" Then
            If mvPreStep.结束天数 <> 0 Then
                txt天数(0).Text = mvPreStep.结束天数 + 1
            ElseIf mvPreStep.开始天数 <> 0 Then
                txt天数(0).Text = mvPreStep.开始天数 + 1
            End If
        Else
            txt天数(0).Text = "1"
        End If
        If mvNextStep.名称 <> "" And txt天数(0).Text <> "" Then
            If mvNextStep.开始天数 <> 0 And mvNextStep.开始天数 - 1 > Val(txt天数(0).Text) Then
                txt天数(1).Text = mvNextStep.开始天数 - 1
            End If
        End If
        If txt天数(0).Text <> "" Then
            Call MakeStepName
        End If
    End If
    
    '备用分支只允许修改说明
    If mvStep.父ID <> 0 Then
        Me.Caption = "分支设置"
        txt名称.Enabled = False
        txt名称.BackColor = Me.BackColor
        For i = 0 To txt天数.UBound
            txt天数(i).Enabled = False
            txt天数(i).BackColor = Me.BackColor
        Next
        For i = 0 To ud天数.UBound
            ud天数(i).Enabled = False
        Next
    End If
    
    '备用分支才设置分类
    If mvStep.父ID = 0 Then
        lbl说明.Top = lbl说明.Top - cbo分类.Height - (cbo分类.Top - txt名称.Top - txt名称.Height)
        txt说明.Top = txt说明.Top - cbo分类.Height - (cbo分类.Top - txt名称.Top - txt名称.Height)
        cmdOK.Top = cmdOK.Top - cbo分类.Height - (cbo分类.Top - txt名称.Top - txt名称.Height)
        cmdCancel.Top = cmdCancel.Top - cbo分类.Height - (cbo分类.Top - txt名称.Top - txt名称.Height)
        
        Line1(0).Y1 = Line1(0).Y1 - cbo分类.Height - (cbo分类.Top - txt名称.Top - txt名称.Height)
        Line1(0).Y2 = Line1(0).Y1
        Line1(1).Y1 = Line1(1).Y1 - cbo分类.Height - (cbo分类.Top - txt名称.Top - txt名称.Height)
        Line1(1).Y2 = Line1(1).Y1
        
        Me.Height = Me.Height - cbo分类.Height - (cbo分类.Top - txt名称.Top - txt名称.Height)
    
        lbl分类.Visible = False
        cbo分类.Visible = False
    Else
        For i = 0 To UBound(Split(mstr分类s, "|"))
            cbo分类.AddItem Split(mstr分类s, "|")(i)
        Next
    End If
End Sub

Private Sub txt名称_GotFocus()
    Call zlControl.TxtSelAll(txt名称)
End Sub

Private Sub txt说明_GotFocus()
    Call zlControl.TxtSelAll(txt说明)
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txt天数_Change(Index As Integer)
    txt天数(1).Enabled = txt天数(0).Text <> ""
    ud天数(1).Enabled = txt天数(1).Enabled
    If Not txt天数(1).Enabled Then
        txt天数(1).Text = ""
        txt天数(1).BackColor = Me.BackColor
    Else
        txt天数(1).BackColor = txt天数(0).BackColor
    End If
    
    If Visible Then Call MakeStepName
End Sub

Private Sub MakeStepName()
    Dim str天数 As String
    Dim i As Long
    
    If txt天数(1).Text <> "" Then
        str天数 = "就诊第" & txt天数(0).Text & "-" & txt天数(1).Text & "天"
    Else
        str天数 = "就诊第" & txt天数(0).Text & "天"
    End If
    
    txt名称.Text = str天数
End Sub

Private Sub txt天数_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt天数(Index))
End Sub

Private Sub txt天数_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
