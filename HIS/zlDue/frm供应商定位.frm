VERSION 5.00
Begin VB.Form frm供应商定位 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "供应商定位"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra 
      Height          =   60
      Left            =   -30
      TabIndex        =   20
      Top             =   600
      Width           =   5580
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   105
      TabIndex        =   15
      Top             =   3045
      Width           =   1100
   End
   Begin VB.Frame fraTemp3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   2700
      Visible         =   0   'False
      Width           =   4830
      Begin VB.OptionButton optSelect 
         Caption         =   "模糊查找(&3)"
         Height          =   180
         Index           =   3
         Left            =   2865
         TabIndex        =   11
         Top             =   0
         Value           =   -1  'True
         Width           =   1800
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "精确查找(&2)"
         Height          =   180
         Index           =   2
         Left            =   975
         TabIndex        =   10
         Top             =   0
         Width           =   1395
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "查找方式："
         Height          =   180
         Index           =   4
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.Frame fraTemp2 
      Height          =   30
      Left            =   -15
      TabIndex        =   17
      Top             =   2910
      Visible         =   0   'False
      Width           =   5580
   End
   Begin VB.CommandButton cmdAdva 
      Caption         =   "高级(&A)"
      Height          =   350
      Left            =   3270
      TabIndex        =   13
      Top             =   3045
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4365
      TabIndex        =   14
      Top             =   3045
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "定位(&F)"
      Height          =   350
      Left            =   2175
      TabIndex        =   12
      Top             =   3045
      Width           =   1100
   End
   Begin VB.Frame fraTemp1 
      Height          =   30
      Left            =   -30
      TabIndex        =   16
      Top             =   2115
      Width           =   5580
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Index           =   2
      Left            =   1335
      MaxLength       =   40
      TabIndex        =   5
      Tag             =   "名称"
      Top             =   1695
      Width           =   3735
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Index           =   1
      Left            =   1335
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "简码"
      Top             =   1260
      Width           =   1230
   End
   Begin VB.Frame fraTemp4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   180
      Left            =   120
      TabIndex        =   19
      Top             =   2340
      Visible         =   0   'False
      Width           =   4830
      Begin VB.OptionButton optSelect 
         Caption         =   "满足其中任意条件(&1)"
         Height          =   180
         Index           =   1
         Left            =   2865
         TabIndex        =   8
         Top             =   0
         Value           =   -1  'True
         Width           =   2145
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "满足所有条件(&0)"
         Height          =   180
         Index           =   0
         Left            =   975
         TabIndex        =   7
         Top             =   0
         Width           =   1755
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "条件设置："
         Height          =   180
         Index           =   3
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Index           =   0
      Left            =   1335
      MaxLength       =   8
      TabIndex        =   1
      Tag             =   "编码"
      Top             =   840
      Width           =   1230
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "通过输入以下条件,将定位到你所需要查找的供应商."
      Height          =   180
      Left            =   750
      TabIndex        =   21
      Top             =   375
      Width           =   4140
   End
   Begin VB.Image img晋升 
      Height          =   480
      Left            =   120
      Picture         =   "frm供应商定位.frx":0000
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "编    码(&D)"
      Height          =   180
      Index           =   0
      Left            =   315
      TabIndex        =   0
      Top             =   900
      Width           =   990
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "单位名称(&N)"
      Height          =   180
      Index           =   2
      Left            =   315
      TabIndex        =   4
      Top             =   1755
      Width           =   990
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "简    码(&J)"
      Height          =   180
      Index           =   1
      Left            =   315
      TabIndex        =   2
      Top             =   1320
      Width           =   990
   End
End
Attribute VB_Name = "frm供应商定位"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrSql As String
Private mstrOthers(0 To 2) As String '0-编码,1-简码,2-名称

Public Function getSql(ByRef strOthers() As String) As String
    cmdAdva_Click
    Me.Show vbModal
    getSql = mstrSql
    strOthers = mstrOthers
End Function

Private Sub cmdAdva_Click()
    If Left(cmdAdva.Caption, 2) = "高级" Then
        Me.Height = 3950
        cmdHelp.Top = Me.fraTemp1.Top + 1000
        cmdAdva.Top = cmdHelp.Top
        cmdOk.Top = cmdHelp.Top
        cmdCancel.Top = cmdHelp.Top
        fraTemp2.Top = cmdHelp.Top - 100
        
        cmdAdva.Caption = "隐藏(&A)"
        fraTemp2.Visible = True
        fraTemp3.Visible = True
        fraTemp4.Visible = True
    Else
        fraTemp2.Visible = False
        fraTemp3.Visible = False
        fraTemp4.Visible = False
        Me.Height = 3050
        cmdHelp.Top = Me.fraTemp1.Top + 100
        cmdAdva.Top = cmdHelp.Top
        cmdOk.Top = cmdHelp.Top
        cmdCancel.Top = cmdHelp.Top
        fraTemp2.Top = cmdHelp.Top - 100
        cmdAdva.Caption = "高级(&A)"
    End If
End Sub

Private Sub cmdCancel_Click()
    mstrSql = ""
    Me.Hide
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strLinkStr As String, strLeftStr As String, strRightStr As String, intTemp As Integer, strTemp As String
    If optSelect(0).Value Then
        strLinkStr = " And "
    Else
        strLinkStr = " Or "
    End If
    If optSelect(2).Value Then
'        strLeftStr = " = '"
'        strRightStr = "'"
        strLeftStr = " = "
        strRightStr = ""
    Else
'        strLeftStr = " Like '" & IIf(gstrMatchMethod = "0", "%", "")
'        strRightStr = "%'"
        strLeftStr = " Like "
        strRightStr = "%"
    End If
    mstrSql = "Select ID,上级ID,类型,名称 From 供应商 Where ("
    strTemp = ""
    For intTemp = 0 To 2
'        If Trim(txtFind(intTemp)) <> "" Then strTemp = strTemp & IIf(strTemp = "", "", strLinkStr) & "Upper(" & txtFind(intTemp).Tag & ") " & strLeftStr & UCase(txtFind(intTemp)) & strRightStr
        If Trim(txtFind(intTemp)) <> "" Then
            strTemp = strTemp & IIf(strTemp = "", "", strLinkStr) & txtFind(intTemp).Tag & strLeftStr & "Upper([" & (intTemp + 8) & "])"
            mstrOthers(intTemp) = IIf(gstrMatchMethod = "0", "%", "") & UCase(txtFind(intTemp)) & strRightStr
        End If
    Next
    If strTemp = "" Then
        MsgBox "请指定最少一个定位条件，需要退出请点击“取消”按钮。", vbInformation, Me.Caption
        txtFind(0).SetFocus
        Exit Sub
    End If
    mstrSql = mstrSql & strTemp & ") and 末级=1"
    Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub optSelect_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtFind_GotFocus(Index As Integer)
    Dim blnOpen As Boolean
    
   Select Case Index
    Case 2
            blnOpen = True
    Case Else
            blnOpen = False
    End Select
    SetTxtGotFocus txtFind(Index), blnOpen
End Sub

Private Sub txtFind_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtFind_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m数字式
    Else
        zlControl.TxtCheckKeyPress txtFind, KeyAscii, m文本式
    End If
End Sub

Private Sub txtFind_LostFocus(Index As Integer)
        ImeLanguage False
End Sub

