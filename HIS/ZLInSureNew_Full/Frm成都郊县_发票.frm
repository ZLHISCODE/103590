VERSION 5.00
Begin VB.Form Frm成都郊县_发票 
   Caption         =   "发票号处理"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4395
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定"
      Height          =   405
      Left            =   1680
      TabIndex        =   2
      Top             =   1620
      Width           =   975
   End
   Begin VB.TextBox TXT发票号 
      Height          =   405
      Left            =   690
      TabIndex        =   1
      Top             =   930
      Width           =   3105
   End
   Begin VB.Label lbl说明 
      Caption         =   "请输入发票号码："
      BeginProperty Font 
         Name            =   "华文中宋"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   420
      Width           =   2325
   End
End
Attribute VB_Name = "Frm成都郊县_发票"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mstr返回 As String
Public Function 发票号() As String
    '功能:
    '参数:lng性质 1 表示月份错误；2 表示日期错误
    '返回: 返回正确的数据

    Me.Show 1
    发票号 = mstr返回
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmd确定_Click()
    If TXT发票号.Text <> "" Then
       mstr返回 = Me.TXT发票号.Text
       Unload Me
    Else
       MsgBox "请输入发票号。"
       Exit Sub
    End If
End Sub

Private Sub Form_Activate()

    Me.TXT发票号.SetFocus
    
End Sub


