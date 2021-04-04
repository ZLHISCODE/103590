VERSION 5.00
Begin VB.Form frmSQLTrace 
   BackColor       =   &H80000005&
   Caption         =   "跟踪工具"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmSQLTrace.frx":0000
   ScaleHeight     =   5760
   ScaleWidth      =   6315
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdEnter 
      Caption         =   "现在进入跟踪工具(&E)… "
      Height          =   350
      Left            =   840
      TabIndex        =   0
      Top             =   3600
      Width           =   2190
   End
   Begin VB.Image imgMain 
      Height          =   720
      Left            =   360
      Picture         =   "frmSQLTrace.frx":803A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Height          =   3330
      Left            =   870
      TabIndex        =   2
      Top             =   615
      Width           =   4140
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SQL跟踪"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmSQLTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFilePath As String

Private Sub cmdEnter_Click()
    On Error GoTo errh
    
    Call ShowFlash("正在加载SQL跟踪工具...")
    'Shell "E:\vb project\zlSvrTools\SQLTrace.exe zlUserName=ZLHISzlPassword=HISzlServer=QZYY"
    Shell mstrFilePath & "\ZLSQLTrace.exe zlUserName=" & gstrUserName & "zlPassword=" & gstrPassword & "zlServer=" & gstrServer
    Call ShowFlash("")
    Exit Sub
errh:
    Call ShowFlash("")
    MsgBox "请检查" & mstrFilePath & "\ZLSQLTrace.exe  是否存在。"
End Sub

Private Sub Form_Load()
    
    lblMain.Caption = "本工具通过Oracle的SQLTrace功能来跟踪和分析SQL性能问题。" & _
    vbCrLf & vbCrLf & "支持对指定的用户会话进行SQL跟踪，从服务器获取SQLTrace文件到客户端，以及解析SQLTrace文件。" & _
    vbCrLf & vbCrLf & "支持对多个SQLTrace文件进行对比，以及快速过滤出含有性能问题的SQL语句，方便地查看分析执行计划。"
    
    '设置获取ZLSQLTrace.EXE的路径
    mstrFilePath = GetSetting("ZLSOFT", "公共全局", "程序路径", App.Path)
    If mstrFilePath = App.Path Then
        mstrFilePath = App.Path
    Else
        'C:\APPSOFT\ZLHIS+.exe
        mstrFilePath = Mid(mstrFilePath, 1, InStrRev(mstrFilePath, "\") - 1)
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
    
    With lblMain
        .Top = imgMain.Top
        .Height = Me.ScaleHeight - .Top * 2
        .Left = imgMain.Left * 2 + imgMain.Width
        .Width = Me.ScaleWidth - .Left - imgMain.Left
    End With

    Dim intCount As Integer, intRows As Integer, aryRow() As String
    intRows = 1
    aryRow() = Split(lblMain.Caption, vbCrLf)
    For intCount = 0 To UBound(aryRow)
        intRows = intRows + TextWidth(aryRow(intCount)) \ (lblMain.Width - 90) + 1
    Next
    If intRows * TextHeight("A") < lblMain.Height + TextHeight("A") Then
        cmdEnter.Top = lblMain.Top + intRows * TextHeight("A")
    Else
        cmdEnter.Top = lblMain.Top + lblMain.Height + TextHeight("A")
    End If
    cmdEnter.Left = lblMain.Left
    
End Sub

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

End Sub

