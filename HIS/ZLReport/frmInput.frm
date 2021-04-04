VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmInput 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmInput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdAtt 
      Caption         =   "…"
      Height          =   285
      Left            =   2145
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1590
      Width           =   300
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   4200
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "字体加粗"
      Height          =   255
      Left            =   3000
      TabIndex        =   20
      Top             =   1610
      Width           =   1095
   End
   Begin VB.PictureBox Pic对齐 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   2970
      ScaleHeight     =   810
      ScaleWidth      =   1350
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   690
      Visible         =   0   'False
      Width           =   1350
      Begin VB.CheckBox Chk对齐 
         Height          =   285
         Index           =   8
         Left            =   900
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   540
         Value           =   2  'Grayed
         Width           =   465
      End
      Begin VB.CheckBox Chk对齐 
         Height          =   285
         Index           =   7
         Left            =   900
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   270
         Value           =   2  'Grayed
         Width           =   465
      End
      Begin VB.CheckBox Chk对齐 
         Height          =   285
         Index           =   6
         Left            =   900
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         Value           =   2  'Grayed
         Width           =   465
      End
      Begin VB.CheckBox Chk对齐 
         Height          =   285
         Index           =   5
         Left            =   450
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   540
         Value           =   2  'Grayed
         Width           =   465
      End
      Begin VB.CheckBox Chk对齐 
         Height          =   285
         Index           =   4
         Left            =   450
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   270
         Value           =   2  'Grayed
         Width           =   465
      End
      Begin VB.CheckBox Chk对齐 
         Height          =   285
         Index           =   3
         Left            =   450
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Value           =   2  'Grayed
         Width           =   465
      End
      Begin VB.CheckBox Chk对齐 
         Height          =   285
         Index           =   2
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   540
         Value           =   2  'Grayed
         Width           =   465
      End
      Begin VB.CheckBox Chk对齐 
         Height          =   285
         Index           =   1
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   270
         Value           =   2  'Grayed
         Width           =   465
      End
      Begin VB.CheckBox Chk对齐 
         Height          =   285
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Value           =   2  'Grayed
         Width           =   465
      End
   End
   Begin VB.CommandButton CmdSelect 
      Height          =   285
      Left            =   4335
      Picture         =   "frmInput.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1230
      Width           =   270
   End
   Begin VB.TextBox Txt对齐 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1230
      Width           =   3030
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3495
      TabIndex        =   4
      Top             =   2235
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1965
      TabIndex        =   3
      Top             =   2235
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   -135
      TabIndex        =   7
      Top             =   2085
      Width           =   5595
   End
   Begin VB.TextBox txtValue 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   855
      Width           =   3300
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1320
      ScaleHeight     =   240
      ScaleWidth      =   1035
      TabIndex        =   22
      Top             =   1590
      Width           =   1095
      Begin VB.PictureBox picFontColor 
         BackColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   23
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "字体颜色"
      Height          =   180
      Left            =   525
      TabIndex        =   19
      Top             =   1650
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   255
      Picture         =   "frmInput.frx":0F50
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "对齐方式"
      Height          =   180
      Left            =   525
      TabIndex        =   8
      Top             =   1290
      Width           =   720
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "内容"
      Height          =   180
      Left            =   885
      TabIndex        =   6
      Top             =   915
      Width           =   360
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "请在下面的输入框中输入合法的内容！"
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   945
      TabIndex        =   5
      Top             =   195
      Width           =   3585
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'入口参数
Public I_strInfo As String '提示信息
Public I_strTitle As String '窗体标题
Public I_intMaxLen As Integer '最大字符数
Public I_blnAllowNULL As Boolean '是允许输入空串
Public I_bytType As Byte '输入类型,0-无,1-数字,2-日期
Public I_strMask As String '输入字符的限制码
'入/出口参数
Public IO_strValue As String '输入内容
Public IO_IntAlig As Integer '对齐方式
Public IO_FontBold As Integer '字体加粗
Public IO_FontColor As Long  '字体颜色
Private BlnIn As Boolean


Private Sub Chk对齐_Click(Index As Integer)
    Dim i As Integer
    If BlnIn = False Then Exit Sub
    BlnIn = False

    For i = 0 To Chk对齐.count - 1
        Chk对齐(i).Value = 0
    Next
    Chk对齐(Index).Value = 1
    IO_IntAlig = Index
    
    Select Case Index
    Case 0
        Txt对齐 = "左上"
    Case 1
        Txt对齐 = "左中"
    Case 2
        Txt对齐 = "左下"
    Case 3
        Txt对齐 = "中上"
    Case 4
        Txt对齐 = "正中"
    Case 5
        Txt对齐 = "中下"
    Case 6
        Txt对齐 = "右上"
    Case 7
        Txt对齐 = "右中"
    Case 8
        Txt对齐 = "右下"
    End Select
    
    BlnIn = True
    
    On Error Resume Next
    Txt对齐.SetFocus
End Sub

Private Sub Chk对齐_LostFocus(Index As Integer)
    Pic对齐.Visible = False
End Sub

Private Sub cmdAtt_Click()
    On Error Resume Next
    
    cdg.CancelError = True
    cdg.Flags = &H1 Or &H2
    cdg.Color = picFontColor.BackColor
    cdg.ShowColor
    If Err.Number = 0 Then
        picFontColor.BackColor = cdg.Color
    Else
        Err.Clear
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not I_blnAllowNULL And txtValue.Text = "" Then
        MsgBox "不允许输入空值！", vbInformation, App.Title: txtValue.SetFocus: Exit Sub
    End If
    If InString(txtValue.Text, "'|~^") Then
        MsgBox "输入了非法字符！", vbInformation, App.Title
        txtValue.SetFocus: Exit Sub
    End If
    If TLen(txtValue.Text) > I_intMaxLen And I_intMaxLen <> 0 Then
        MsgBox "输入内容长度不能超过 " & I_intMaxLen & " 个字符！", vbInformation, App.Title: txtValue.SetFocus: Exit Sub
    End If
    If I_bytType = 1 And Not IsNumeric(txtValue.Text) Then
        MsgBox "请输入数字型数据！", vbInformation, App.Title: txtValue.SetFocus: Exit Sub
    End If
    If I_bytType = 2 And Not IsDate(txtValue.Text) Then
        MsgBox "请输入日期型数据！", vbInformation, App.Title: txtValue.SetFocus: Exit Sub
    End If
    IO_strValue = txtValue.Text
    IO_FontColor = picFontColor.BackColor
    IO_FontBold = chkBold.Value
    gblnOK = True
    Hide
End Sub

Private Sub CmdSelect_Click()
    On Error Resume Next
    Chk对齐(IO_IntAlig).Value = 1
    Chk对齐(IO_IntAlig).SetFocus
    Pic对齐.Visible = Pic对齐.Visible Xor True
End Sub

Private Sub Form_Load()
    gblnOK = False
    Caption = I_strTitle
    lblInfo.Caption = I_strInfo
    txtValue.Text = IO_strValue
    Chk对齐(IO_IntAlig).Value = 1
    picFontColor.BackColor = IO_FontColor
    chkBold.Value = IO_FontBold
    
    Select Case IO_IntAlig
    Case 0
        Txt对齐 = "左上"
    Case 1
        Txt对齐 = "左中"
    Case 2
        Txt对齐 = "左下"
    Case 3
        Txt对齐 = "中上"
    Case 4
        Txt对齐 = "正中"
    Case 5
        Txt对齐 = "中下"
    Case 6
        Txt对齐 = "右上"
    Case 7
        Txt对齐 = "右中"
    Case 8
        Txt对齐 = "右下"
    End Select
    BlnIn = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    I_intMaxLen = 0
    I_blnAllowNULL = False
    I_strMask = ""
    I_bytType = 0
End Sub

Private Sub txtValue_GotFocus()
    SelAll txtValue
    Pic对齐.Visible = False
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    If I_strMask <> "" Then
        If InStr(I_strMask & Chr(8) & Chr(3) & Chr(22) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If InStr("#@~`'""|^↑←", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Txt对齐_GotFocus()
    Pic对齐.Visible = False
End Sub
