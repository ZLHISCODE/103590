VERSION 5.00
Begin VB.Form Frm乐山_提示 
   Caption         =   "提示"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   3405
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Cmd确定 
      Caption         =   "确定"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txt说明 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Lbl说明 
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Frm乐山_提示"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mstr返回 As String
Public Function 出生日期更改_乐山(ByVal lng性质 As Long, ByVal str说明 As String) As String
    '功能:
    '参数:lng性质 1 表示月份错误；2 表示日期错误
    '返回: 返回正确的数据

    If lng性质 = 1 Then
       Me.Lbl说明.Caption = "该病人的出生月份（返回为：" & str说明 & "）错误，请输入正确的值："
    Else
       Me.Lbl说明.Caption = "该病人的出生日期（返回为：" & str说明 & "）错误，请输入正确的值："
    End If
    Me.Show 1
    出生日期更改_乐山 = mstr返回
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub Cmd确定_Click()
    If txt说明.Text <> "" Then
       mstr返回 = Me.txt说明.Text
       Unload Me
    Else
       MsgBox "请输入正确的数值。"
       Exit Sub
    End If
End Sub

Private Sub Form_Activate()

    Me.txt说明.SetFocus
    
End Sub

