VERSION 5.00
Begin VB.Form frmPrintDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "中联检验报告打印动态库DEMO"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9660
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrintDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdGetPatiList 
      Caption         =   "取住院清单"
      Height          =   360
      Left            =   795
      TabIndex        =   17
      Top             =   4365
      Width           =   1200
   End
   Begin VB.TextBox txtPatiDate 
      Height          =   300
      Left            =   840
      TabIndex        =   15
      Top             =   3915
      Width           =   1200
   End
   Begin VB.TextBox txtPatiNo 
      Height          =   300
      Left            =   840
      TabIndex        =   13
      Top             =   3570
      Width           =   1200
   End
   Begin VB.TextBox txtNo 
      Height          =   300
      Left            =   780
      TabIndex        =   11
      Top             =   2745
      Width           =   1200
   End
   Begin VB.TextBox txtSVr 
      Height          =   300
      Left            =   780
      TabIndex        =   9
      Top             =   1050
      Width           =   1200
   End
   Begin VB.TextBox txtPwd 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   780
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   637
      Width           =   1200
   End
   Begin VB.TextBox txtUser 
      Height          =   300
      Left            =   780
      TabIndex        =   5
      Top             =   225
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      Height          =   345
      Left            =   8355
      TabIndex        =   4
      Top             =   4365
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Height          =   345
      Left            =   6540
      TabIndex        =   3
      Top             =   4365
      Width           =   1095
   End
   Begin VB.CommandButton cmdGetList 
      Caption         =   "取清单(&G)"
      Height          =   345
      Left            =   750
      TabIndex        =   2
      Top             =   3105
      Width           =   1095
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "登录(&L)"
      Height          =   350
      Left            =   465
      TabIndex        =   1
      Top             =   1650
      Width           =   1100
   End
   Begin VB.TextBox txtList 
      Height          =   4125
      Left            =   2190
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   135
      Width           =   7350
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "出院日期"
      Height          =   195
      Index           =   3
      Left            =   75
      TabIndex        =   16
      Top             =   3945
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "住院号"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   14
      Top             =   3570
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "门诊号"
      Height          =   195
      Index           =   2
      Left            =   165
      TabIndex        =   12
      Top             =   2745
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "服务器"
      Height          =   195
      Left            =   165
      TabIndex        =   10
      Top             =   1065
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "口　令"
      Height          =   195
      Left            =   165
      TabIndex        =   8
      Top             =   652
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户名"
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   6
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "frmPrintDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mzlLisPrint As Object

Private Sub cmdExit_Click()
    If mzlLisPrint Is Nothing Then
        Call mzlLisPrint.zlLoginOut
        Set mzlLisPrint = Nothing
    End If
    Unload Me
End Sub

Private Sub cmdGetPatiList_Click()
    '取住院清单
    Dim strReturn As String
    Dim varTmp As Variant, i As Integer
    
    On Error GoTo hErr
    If txtPatiNo = "" Or txtPatiDate = "" Then
        MsgBox "请填写住院号和出院日期!"
        Exit Sub
    End If
    If Not mzlLisPrint Is Nothing Then
        strReturn = mzlLisPrint.zlGetZyPrintList(txtPatiNo, txtPatiDate)
        If Left(strReturn, 2) <> "OK" Then
            MsgBox strReturn
        Else
            varTmp = Split(strReturn, "^")
            txtList.Text = ""
            For i = LBound(varTmp) To UBound(varTmp)
                If Trim(varTmp(i)) <> "OK" And Trim$(varTmp(i)) <> "" Then
                    If txtList <> "" Then
                        txtList = txtList & vbNewLine & Trim(varTmp(i))
                    Else
                        txtList = Trim(varTmp(i))
                    End If
                End If
            Next
        End If
    End If
    Exit Sub
hErr:
    MsgBox Err.Description
End Sub

Private Sub cmdLogin_Click()
    '登录
    Dim strReturn As String
    On Error GoTo hErr
    
    If mzlLisPrint Is Nothing Then Set mzlLisPrint = CreateObject("zlLisComPrint.clsComPrint")
    If Not mzlLisPrint Is Nothing Then
        strReturn = mzlLisPrint.zllogin(txtUser, txtPwd, txtSVr)
        If strReturn <> "OK" Then MsgBox strReturn
    End If
    Exit Sub
hErr:
    MsgBox Err.Description
End Sub

Private Sub cmdGetList_Click()
    '取可打印清单
    Dim strReturn As String
    Dim varTmp As Variant, i As Integer
    
    On Error GoTo hErr
    If txtNo = "" Then
        MsgBox "请填写门诊号!"
        Exit Sub
    End If
    If Not mzlLisPrint Is Nothing Then
        strReturn = mzlLisPrint.zlGetPrintList(txtNo, 700)
        If Left(strReturn, 2) <> "OK" Then
            MsgBox strReturn
        Else
            varTmp = Split(strReturn, "^")
            txtList.Text = ""
            For i = LBound(varTmp) To UBound(varTmp)
                If Trim(varTmp(i)) <> "OK" And Trim$(varTmp(i)) <> "" Then
                    If txtList <> "" Then
                        txtList = txtList & vbNewLine & Trim(varTmp(i))
                    Else
                        txtList = Trim(varTmp(i))
                    End If
                End If
            Next
        End If
    End If
    Exit Sub
hErr:
    MsgBox Err.Description
End Sub

Private Sub cmdPrint_Click()
    '打印报告
    Dim strLine As String
    Dim varTmp As Variant
    Dim strReturn As String, i As Integer, lngID As Long
    On Error GoTo hErr
    If txtList = "" Then
        MsgBox "没有可打印的报告！"
        Exit Sub
    End If
    If Not mzlLisPrint Is Nothing Then
        varTmp = Split(txtList, vbNewLine)
        For i = LBound(varTmp) To UBound(varTmp)
            strLine = varTmp(i)
            lngID = Val(Split(strLine, "|")(0))
            If lngID > 0 Then
                strReturn = mzlLisPrint.zlPrintReport(lngID)
                If strReturn <> "OK" Then
                    
                    MsgBox "打印" & strLine & "失败！" & vbNewLine & strReturn
                End If
            End If
        Next
    End If
    Exit Sub
hErr:
    MsgBox Err.Description
End Sub
