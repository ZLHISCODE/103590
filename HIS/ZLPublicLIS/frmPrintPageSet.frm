VERSION 5.00
Begin VB.Form frmPrintPageSet 
   BorderStyle     =   0  'None
   Caption         =   "打印页码设置"
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmd取消 
      Caption         =   "取 消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txt说明 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmPrintPageSet.frx":0000
      Top             =   120
      Width           =   8055
   End
   Begin VB.TextBox txt选择页码 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   3000
      Width           =   4095
   End
   Begin VB.CommandButton cmd确认 
      Caption         =   "确  认"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txt结束页 
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txt起始页 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "选择打印页码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "结束页"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "起始页"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "frmPrintPageSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mstrPrintPage As String
Private mintPageCount As Integer

Public Function ShowMe(intPageCount As Integer) As String
    mintPageCount = intPageCount
    Me.Show 1
    ShowMe = mstrPrintPage
End Function

Private Sub cmd取消_Click()
    mstrPrintPage = ""
    Unload Me
    
End Sub

Private Sub cmd确认_Click()
    Dim strTemp As String
    Dim Temp As String
    Dim i As Integer
    If Trim(txt选择页码.Text) <> "" Then
        strTemp = Trim(txt选择页码.Text)
        For i = 1 To Len(strTemp)
            Temp = Mid(strTemp, i, i + 1)
            If (Asc(Temp) >= 48 And Asc(Temp) <= 57) Or Asc(Temp) = 44 Then
                
            Else
                MsgBox "输入内容包含非法内容,只能输入数字和英文逗号(1,2,3,4,5,6,7,8,9,0,)", vbOKOnly, gstrSysName
                mstrPrintPage = ""
                Exit Sub
            End If
        Next
        mstrPrintPage = Trim(txt选择页码.Text)
    Else
        If Trim(txt结束页.Text) = "" Then
            MsgBox "结束页都必须输入,请重新输入", vbOKOnly, gstrSysName
            Exit Sub
        End If
        If Trim(txt起始页.Text) = "" Then
            MsgBox "起始页都必须输入,请重新输入", vbOKOnly, gstrSysName
            Exit Sub
        End If
        If CInt(Trim(txt结束页.Text)) > mintPageCount Then
            MsgBox "结束页号不能大于总页数,请重新输入", vbOKOnly, gstrSysName
            Exit Sub
        End If
        If CInt(Trim(txt起始页.Text)) <= 0 Or CInt(Trim(txt起始页.Text)) > mintPageCount Then
            MsgBox "起始页号必须大于0小于等于总页数,请重新输入", vbOKOnly, gstrSysName
            Exit Sub
        End If
        
        If CInt(Trim(txt结束页.Text)) <= 0 Or CInt(Trim(txt起始页.Text)) <= 0 Then
            MsgBox "起始页和结束页都必须大于0,请重新输入", vbOKOnly, gstrSysName
            mstrPrintPage = ""
            txt结束页.Text = ""
            txt起始页.Text = ""
            Exit Sub
        End If
        If CInt(Trim(txt结束页.Text)) = CInt(Trim(txt起始页.Text)) Then
            mstrPrintPage = CInt(Trim(txt起始页.Text))
        ElseIf CInt(Trim(txt结束页.Text)) < CInt(Trim(txt起始页.Text)) Then
            MsgBox "结束页必须大于等于起始页,请重新输入", vbOKOnly, gstrSysName
            mstrPrintPage = ""
            txt结束页.Text = ""
            txt起始页.Text = ""
        Else
            mstrPrintPage = ""
            For i = CInt(Trim(txt起始页.Text)) To CInt(Trim(txt结束页.Text))
                mstrPrintPage = mstrPrintPage & "," & i
            Next
        End If
    End If
    If Mid(mstrPrintPage, 1, 1) = "," Then
        mstrPrintPage = Mid(mstrPrintPage, 2)
    End If
    If Mid(mstrPrintPage, Len(mstrPrintPage), 1) = "," Then
        mstrPrintPage = Mid(mstrPrintPage, Len(mstrPrintPage), 1)
    End If
    Unload Me
End Sub

'Private Sub txt结束页_Change()
'    If CInt(Trim(txt结束页.Text)) > 0 Then
'
'    Else
'        txt结束页.Text = ""
'    End If
'End Sub

Private Sub txt结束页_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        Exit Sub
    Else
        MsgBox "输入非数字内容,请重新输入", vbOKOnly, gstrSysName
        Exit Sub
    End If
End Sub

Private Sub txt起始页_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
        Exit Sub
    Else
        MsgBox "输入非数字内容,请重新输入", vbOKOnly, gstrSysName
        Exit Sub
    End If
    
End Sub



Private Sub txt选择页码_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 44 Or KeyAscii = 8 Then
        Exit Sub
    Else
        MsgBox "输入非数字内容,请重新输入", vbOKOnly, gstrSysName
        Exit Sub
    End If
End Sub
