VERSION 5.00
Begin VB.Form frmCheckQueryPass 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7785
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txt 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   20
      PasswordChar    =   "*"
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   2025
      Width           =   1950
   End
   Begin VB.PictureBox picKey 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4410
      Left            =   3495
      ScaleHeight     =   4410
      ScaleWidth      =   4095
      TabIndex        =   0
      Top             =   30
      Width           =   4095
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   2
         Left            =   1410
         TabIndex        =   1
         Top             =   30
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "2"
         backcolor       =   16777215
         forecolor       =   255
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   3
         Left            =   2085
         TabIndex        =   2
         Top             =   30
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "3"
         backcolor       =   16777215
         forecolor       =   255
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   4
         Left            =   2760
         TabIndex        =   3
         Top             =   30
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "4"
         backcolor       =   16777215
         forecolor       =   255
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   5
         Left            =   3435
         TabIndex        =   4
         Top             =   30
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "5"
         backcolor       =   16777215
         forecolor       =   255
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   6
         Left            =   60
         TabIndex        =   5
         Top             =   660
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "6"
         backcolor       =   16777215
         forecolor       =   255
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   7
         Left            =   735
         TabIndex        =   6
         Top             =   660
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "7"
         backcolor       =   16777215
         forecolor       =   255
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   8
         Left            =   1410
         TabIndex        =   7
         Top             =   660
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "8"
         backcolor       =   16777215
         forecolor       =   255
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   36
         Left            =   60
         TabIndex        =   8
         Top             =   3795
         Width           =   1275
         _extentx        =   2249
         _extenty        =   1005
         caption         =   "  确定 "
         backcolor       =   16777215
         fontsize        =   10.5
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   20
         Left            =   1410
         TabIndex        =   9
         Top             =   1905
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "K"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   22
         Left            =   2760
         TabIndex        =   10
         Top             =   1905
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "M"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   23
         Left            =   3435
         TabIndex        =   11
         Top             =   1905
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "N"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   37
         Left            =   1410
         TabIndex        =   12
         Top             =   3795
         Width           =   1275
         _extentx        =   2249
         _extenty        =   1005
         caption         =   "  清除 "
         backcolor       =   16777215
         fontsize        =   10.5
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   24
         Left            =   60
         TabIndex        =   13
         Top             =   2550
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "O"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   25
         Left            =   735
         TabIndex        =   14
         Top             =   2550
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "P"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   26
         Left            =   1410
         TabIndex        =   15
         Top             =   2550
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "Q"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   27
         Left            =   2085
         TabIndex        =   16
         Top             =   2550
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "R"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   28
         Left            =   2760
         TabIndex        =   17
         Top             =   2550
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "S"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   29
         Left            =   3435
         TabIndex        =   18
         Top             =   2550
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "T"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   30
         Left            =   60
         TabIndex        =   19
         Top             =   3180
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "U"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   32
         Left            =   1410
         TabIndex        =   20
         Top             =   3180
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "W"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   33
         Left            =   2085
         TabIndex        =   21
         Top             =   3180
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "X"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   34
         Left            =   2760
         TabIndex        =   22
         Top             =   3180
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "Y"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   35
         Left            =   3435
         TabIndex        =   23
         Top             =   3180
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "Z"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   9
         Left            =   2085
         TabIndex        =   27
         Top             =   660
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "9"
         backcolor       =   16777215
         forecolor       =   255
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   12
         Left            =   60
         TabIndex        =   28
         Top             =   1305
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "C"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   13
         Left            =   735
         TabIndex        =   29
         Top             =   1305
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "D"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   14
         Left            =   1410
         TabIndex        =   30
         Top             =   1305
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "E"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   15
         Left            =   2085
         TabIndex        =   31
         Top             =   1305
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "F"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   16
         Left            =   2760
         TabIndex        =   32
         Top             =   1305
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "G"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   17
         Left            =   3435
         TabIndex        =   33
         Top             =   1305
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "H"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   18
         Left            =   60
         TabIndex        =   34
         Top             =   1905
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "I"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   19
         Left            =   735
         TabIndex        =   35
         Top             =   1905
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "J"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   0
         Left            =   60
         TabIndex        =   36
         Top             =   30
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "0"
         backcolor       =   16777215
         forecolor       =   255
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   10
         Left            =   2760
         TabIndex        =   37
         Top             =   660
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "A"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   1
         Left            =   735
         TabIndex        =   38
         Top             =   30
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "1"
         backcolor       =   16777215
         forecolor       =   255
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   21
         Left            =   2085
         TabIndex        =   39
         Top             =   1905
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "L"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   31
         Left            =   735
         TabIndex        =   40
         Top             =   3180
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "V"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   11
         Left            =   3435
         TabIndex        =   41
         Top             =   660
         Width           =   600
         _extentx        =   1058
         _extenty        =   1005
         caption         =   "B"
         backcolor       =   16777215
         fontsize        =   10.5
         fontbold        =   -1  'True
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
      Begin zl9NewQuery.ctlButton UsrCmd 
         Height          =   570
         Index           =   38
         Left            =   2760
         TabIndex        =   42
         Top             =   3795
         Width           =   1275
         _extentx        =   2249
         _extenty        =   1005
         caption         =   "  退出"
         backcolor       =   16777215
         fontsize        =   10.5
         autosize        =   0   'False
         buttonheight    =   450
         textaligment    =   2
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入您的查询密码。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   870
      TabIndex        =   26
      Top             =   1095
      Width           =   2100
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frmCheckQueryPass.frx":0000
      Top             =   960
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   540
      TabIndex        =   25
      Top             =   2100
      Width           =   420
   End
End
Attribute VB_Name = "frmCheckQueryPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mblnOK As Boolean
Public mstrPass As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyNumpad0 To vbKeyNumpad9
            Call UsrCmd_CommandClick(KeyCode - 96)
        Case vbKey0 To vbKey9
            Call UsrCmd_CommandClick(KeyCode - 48)
        Case vbKeyA To vbKeyZ
            Call UsrCmd_CommandClick(KeyCode - 55)
        Case vbKeyReturn, vbKeySeparator
            Call UsrCmd_CommandClick(36)
        Case vbKeyEscape
            Call UsrCmd_CommandClick(38)
        Case vbKeyDelete, vbKeyDecimal
            Call UsrCmd_CommandClick(37)
        Case vbKeyBack
            Call UsrCmd_CommandClick(40)
        Case Else
            KeyCode = 0
    End Select
End Sub

Private Sub Form_Load()
Dim i As Integer
    For i = 0 To UsrCmd.UBound
        UsrCmd(i).ShowPicture = False
        UsrCmd(i).TextAligment = 1
    Next
End Sub

Private Sub UsrCmd_CommandClick(Index As Integer)
    Select Case Index
        Case 0 To 9
            txt.Text = txt.Text & Index
            txt.SetFocus: txt.SelStart = Len(txt.Text)
        Case 10 To 35
            txt.Text = txt.Text & Chr(Index + 87)
            txt.SetFocus: txt.SelStart = Len(txt.Text)
        Case 36 '确定
            mstrPass = txt.Text
            mblnOK = True
            Unload Me
        Case 37 '取消
            txt.Text = ""
            txt.SetFocus: txt.SelStart = Len(txt.Text)
        Case 38 '退出
            mblnOK = False
            Unload Me
        Case 40 '退格
            txt.Text = Mid(txt.Text, 1, IIf((Len(txt.Text) - 1) < 0, 0, Len(txt.Text) - 1)):: txt.SelStart = Len(txt.Text)
    End Select
    Debug.Print txt.Text
End Sub

Public Function GetPwd(frmParent As Form) As Boolean
   mblnOK = False
   Me.Show 1, frmParent
   GetPwd = mblnOK
End Function

