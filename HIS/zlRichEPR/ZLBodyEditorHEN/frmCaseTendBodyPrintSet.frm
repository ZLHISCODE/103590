VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCaseTendBodyPrintSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印选项"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frmCaseTendBodyPrintSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra 
      Caption         =   "其他"
      Height          =   1020
      Left            =   120
      TabIndex        =   9
      Top             =   2625
      Width           =   4380
      Begin VB.CheckBox chk 
         Caption         =   "不打印心率和脉搏间的连线和阴影(&8)"
         Height          =   195
         Index           =   0
         Left            =   915
         TabIndex        =   12
         Top             =   720
         Value           =   1  'Checked
         Width           =   3360
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   960
         TabIndex        =   11
         Top             =   255
         Width           =   3210
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "质控号(&5)"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   810
      End
   End
   Begin VB.Frame fra打印 
      Caption         =   "打印页脚"
      Height          =   1080
      Left            =   120
      TabIndex        =   4
      Top             =   1500
      Width           =   4380
      Begin VB.CheckBox chk周数 
         Caption         =   "打印住院周数(&7)"
         Height          =   195
         Left            =   525
         TabIndex        =   8
         Top             =   765
         Value           =   1  'Checked
         Width           =   1650
      End
      Begin VB.CheckBox chk页号 
         Caption         =   "打印页号，第一页页号表示为(&3)"
         Height          =   195
         Left            =   525
         TabIndex        =   5
         Top             =   405
         Value           =   1  'Checked
         Width           =   2910
      End
      Begin VB.TextBox txt起始 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1050
         MaxLength       =   3
         TabIndex        =   20
         Text            =   "25"
         Top             =   1680
         Visible         =   0   'False
         Width           =   600
      End
      Begin MSComCtl2.UpDown UD页号 
         Height          =   300
         Left            =   3795
         TabIndex        =   7
         Top             =   345
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt页号"
         BuddyDispid     =   196617
         OrigLeft        =   1590
         OrigTop         =   1365
         OrigRight       =   1830
         OrigBottom      =   1665
         Max             =   999
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UD起始 
         Height          =   300
         Left            =   1665
         TabIndex        =   21
         Top             =   1680
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   25
         BuddyControl    =   "txt起始"
         BuddyDispid     =   196616
         OrigLeft        =   1590
         OrigTop         =   705
         OrigRight       =   1830
         OrigBottom      =   1005
         Max             =   460
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt页号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3435
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "1"
         Top             =   360
         Width           =   360
      End
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2130
         Left            =   2850
         ScaleHeight     =   491.128
         ScaleMode       =   0  'User
         ScaleWidth      =   491.128
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2175
         Visible         =   0   'False
         Width           =   2130
         Begin VB.PictureBox picPaper 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   1485
            Left            =   405
            ScaleHeight     =   1455
            ScaleMode       =   0  'User
            ScaleWidth      =   1140
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "拖动蓝色线条改变起始位置"
            Top             =   270
            Width           =   1170
            Begin VB.PictureBox pic起始 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Left            =   0
               MousePointer    =   7  'Size N S
               ScaleHeight     =   15
               ScaleMode       =   0  'User
               ScaleWidth      =   1140
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   135
               Width           =   1140
            End
         End
         Begin VB.PictureBox picShadow 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1485
            Left            =   450
            ScaleHeight     =   1485
            ScaleWidth      =   1170
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   315
            Width           =   1170
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起始位置"
         Height          =   180
         Left            =   255
         TabIndex        =   23
         Top             =   1740
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   180
         Left            =   1965
         TabIndex        =   22
         Top             =   1710
         Visible         =   0   'False
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   4620
      TabIndex        =   13
      Top             =   165
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4620
      TabIndex        =   14
      Top             =   570
      Width           =   1100
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "预览(&V)"
      Height          =   350
      Left            =   4620
      TabIndex        =   15
      Top             =   165
      Width           =   1100
   End
   Begin VB.Frame fra病历 
      Caption         =   "打印范围"
      Height          =   1380
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   4380
      Begin VB.OptionButton opt全部 
         Caption         =   "打印全部体温单(&6)"
         Height          =   180
         Left            =   480
         TabIndex        =   3
         Top             =   1005
         Width           =   2775
      End
      Begin VB.OptionButton opt连续 
         Caption         =   "从当前体温表开始连续打印(&2)"
         Height          =   180
         Left            =   480
         TabIndex        =   2
         Top             =   675
         Width           =   2775
      End
      Begin VB.OptionButton opt当前 
         Caption         =   "只打印当前选择的体温表(&1)"
         Height          =   180
         Left            =   480
         TabIndex        =   1
         Top             =   345
         Value           =   -1  'True
         Width           =   2745
      End
   End
End
Attribute VB_Name = "frmCaseTendBodyPrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytOpt As Byte

Private mblnFirst As Boolean
Private mintPrintRange As Integer
Private mlngBeginY As Long
Private mintBeginPage As Integer
Private mlngWidth As Long '自定义纸张宽度,Twip
Private mlngHeight As Long '自定义纸张高度'Twip
Private mlngLeft As Long '左边距'mm
Private mlngRight As Long '右边距'mm
Private mlngTop As Long '上边距'mm
Private mlngBottom As Long '下边距'mm

Private mstrPrivs As String

Private Sub chk页号_Click()
    txt页号.Enabled = chk页号.Value = 1
    UD页号.Enabled = chk页号.Value = 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    If Not GetValue Then Exit Sub
    mbytOpt = 1
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Call zlDatabase.SetPara("质控号", txt.Text, glngSys, 1255)
    If Not GetValue Then Exit Sub
    mbytOpt = 2
    Unload Me
End Sub

Private Sub Form_Load()
    mbytOpt = 0
    
    '显示纸张打印位置调整图
        
    mlngWidth = Val(zlDatabase.GetPara("体温单宽度", glngSys, 1255, Printer.Width))
    mlngHeight = Val(zlDatabase.GetPara("体温单高度", glngSys, 1255, Printer.Height))
    mlngLeft = Val(zlDatabase.GetPara("体温单左边距", glngSys, 1255, OFFSET_LEFT))
    mlngRight = Val(zlDatabase.GetPara("体温单右边距", glngSys, 1255, OFFSET_RIGHT))
    mlngTop = Val(zlDatabase.GetPara("体温单上边距", glngSys, 1255, OFFSET_TOP))
    mlngBottom = Val(zlDatabase.GetPara("体温单下边距", glngSys, 1255, OFFSET_BOTTOM))
    
    txt.Text = zlDatabase.GetPara("质控号", glngSys, 1255, "", Array(txt), InStr(mstrPrivs, "护理选项设置") > 0)
    
    If mlngWidth > mlngHeight Then
        picBack.ScaleWidth = mlngWidth / 56.7 * 1.1
        picBack.ScaleHeight = mlngWidth / 56.7 * 1.1
    Else
        picBack.ScaleWidth = mlngHeight / 56.7 * 1.1
        picBack.ScaleHeight = mlngHeight / 56.7 * 1.1
    End If
    picPaper.Width = mlngWidth / 56.7
    picPaper.Height = mlngHeight / 56.7
    picPaper.Left = (picBack.ScaleWidth - picPaper.Width) / 2
    picPaper.Top = (picBack.ScaleHeight - picPaper.Height) / 2
    picShadow.Width = picPaper.Width
    picShadow.Height = picPaper.Height
    picShadow.Left = picPaper.Left + 5
    picShadow.Top = picPaper.Top + 5
    
    picPaper.ScaleWidth = mlngWidth / 56.7
    picPaper.ScaleHeight = mlngHeight / 56.7
    
    '显初始位置
    If Not (mlngBeginY >= mlngTop And mlngBeginY <= picPaper.ScaleHeight - mlngBottom * 2) Then
        mlngBeginY = mlngTop
    End If
    pic起始.Left = 0
    pic起始.Width = picPaper.ScaleWidth
    pic起始.Top = mlngBeginY
    
    UD起始.Min = mlngTop
    UD起始.Max = picPaper.ScaleHeight - 2 * mlngBottom
    UD起始.Value = mlngBeginY
    
    pic起始.ScaleHeight = 1 '不然不能拖动
    
    Call DrawPage
    
    mintPrintRange = Val(zlDatabase.GetPara("连续打印", glngSys, 1255, "1", Array(opt当前, opt连续, opt全部), InStr(mstrPrivs, "护理选项设置") > 0))
    Select Case mintPrintRange
    Case 0
        opt当前.Value = True
    Case 1
        opt连续.Value = True
    Case 2
        opt全部.Value = True
    End Select
    
    chk页号.Value = Val(zlDatabase.GetPara("打印页号", glngSys, 1255, "1", Array(chk页号), InStr(mstrPrivs, "护理选项设置") > 0))
    txt页号.Text = Val(zlDatabase.GetPara("起始页号", glngSys, 1255, "1", Array(txt页号, UD页号), InStr(mstrPrivs, "护理选项设置") > 0))
    chk周数.Value = Val(zlDatabase.GetPara("打印周数", glngSys, 1255, "0", Array(chk周数), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(0).Value = Val(zlDatabase.GetPara("不打印脉搏短绌图形", glngSys, 1255, "0", Array(chk(0)), InStr(mstrPrivs, "护理选项设置") > 0))
    
    mintBeginPage = Val(txt页号.Text)
    
    UD页号.Value = IIf(mintBeginPage = 0, 1, mintBeginPage)

End Sub

Public Function PrintSet(objParent As Object, ByVal blnFirst As Boolean, ByRef intPrintRange As Integer, ByRef lngBeginY As Long, ByRef intBeginPage As Integer, ByVal strPrivs As String) As Byte
'功能：调用打印选项
'参数：blnFirst=是否第一次调用,否则只有"确定","取消",且不允许修改病历打印份数
'      blnCurCase=T=只打印当前病历,F=从当前病历开始连续打印病历
'      lngBeginY=本次病历开始打印位置'mm
'      intBeginPage=起始页号,为0表示不打印页号
'返回：0-取消,1-预览,2-打印
    
    mstrPrivs = strPrivs
    mblnFirst = blnFirst
    mintPrintRange = intPrintRange
    mlngBeginY = lngBeginY
    mintBeginPage = intBeginPage
        
    If Not mblnFirst Then
        opt当前.Enabled = False
        opt连续.Enabled = False
        
        cmdPrint.Visible = False
        cmdCancel.Top = cmdPrint.Top
        cmdPreview.Caption = "确定(&O)"
        cmdPreview.Default = True
    End If
    Me.Show 1, objParent
    
    intPrintRange = mintPrintRange
    lngBeginY = mlngBeginY
    intBeginPage = mintBeginPage
    
    PrintSet = mbytOpt
End Function

Private Sub Form_Unload(Cancel As Integer)
    
    If opt当前.Value Then
        Call zlDatabase.SetPara("连续打印", "0", glngSys, 1255)
    ElseIf opt连续.Value Then
        Call zlDatabase.SetPara("连续打印", "1", glngSys, 1255)
    Else
        Call zlDatabase.SetPara("连续打印", "2", glngSys, 1255)
    End If
    
    Call zlDatabase.SetPara("打印页号", chk页号.Value, glngSys, 1255)
    Call zlDatabase.SetPara("起始页号", Val(txt页号.Text), glngSys, 1255)
    Call zlDatabase.SetPara("打印周数", chk周数.Value, glngSys, 1255)
    Call zlDatabase.SetPara("不打印脉搏短绌图形", chk(0).Value, glngSys, 1255)
    Call zlDatabase.SetPara("质控号", txt.Text, glngSys, 1255)
    
End Sub

Private Sub pic起始_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If pic起始.Top + y > UD起始.Max Or pic起始.Top + y < UD起始.Min Then Exit Sub
        pic起始.Top = pic起始.Top + y
        UD起始.Value = pic起始.Top
        Call DrawPage
        Me.Refresh
    End If
End Sub

Private Sub txt起始_Change()
    If Val(txt起始.Text) >= UD起始.Min And Val(txt起始.Text) <= UD起始.Max Then
        UD起始.Value = Val(txt起始.Text)
    End If
End Sub

Private Sub txt起始_GotFocus()
    zlControl.TxtSelAll txt起始
End Sub

Private Sub txt起始_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt页号_GotFocus()
    zlControl.TxtSelAll txt页号
End Sub

Private Sub txt页号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Function GetValue() As Boolean
    If Not (Val(txt起始.Text) >= UD起始.Min And Val(txt起始.Text) <= UD起始.Max) Then
        MsgBox "起始位置应该在 " & UD起始.Min & " 至 " & UD起始.Max & " 之间！", vbInformation, gstrSysName
        txt起始.SetFocus: Exit Function
    End If
    
    If opt当前.Value Then
        mintPrintRange = 0
    ElseIf opt连续.Value Then
        mintPrintRange = 1
    Else
        mintPrintRange = 2
    End If

    mlngBeginY = Val(txt起始.Text)
    If chk页号.Value = 1 Then
        mintBeginPage = Val(txt页号.Text)
    Else
        mintBeginPage = 0
    End If
    
    GetValue = True
End Function

Private Sub UD起始_Change()
    pic起始.Top = UD起始.Value
    Call DrawPage
End Sub

Private Sub DrawPage()
    picPaper.Cls
    picPaper.Line (0, mlngTop)-(picPaper.ScaleWidth, mlngTop), &H808080
    picPaper.Line (0, picPaper.ScaleHeight - mlngBottom)-(picPaper.ScaleWidth, picPaper.ScaleHeight - mlngBottom), &H808080
    picPaper.Line (mlngLeft, 0)-(mlngLeft, picPaper.ScaleHeight), &H808080
    picPaper.Line (picPaper.ScaleWidth - mlngRight, 0)-(picPaper.ScaleWidth - mlngRight, picPaper.ScaleHeight), &H808080
    
    picPaper.Line (mlngLeft, UD起始.Value)-(picPaper.ScaleWidth - mlngRight, picPaper.ScaleHeight - mlngBottom), &H808080, B
End Sub




