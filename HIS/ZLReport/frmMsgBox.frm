VERSION 5.00
Begin VB.Form frmMsgBox 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "frmMsgBox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdDo 
      Caption         =   "###"
      Height          =   350
      Index           =   0
      Left            =   1695
      TabIndex        =   0
      Top             =   900
      Width           =   1100
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   270
      Picture         =   "frmMsgBox.frx":000C
      Top             =   210
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   3
      Left            =   270
      Picture         =   "frmMsgBox.frx":08D6
      Top             =   210
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   2
      Left            =   270
      Picture         =   "frmMsgBox.frx":11A0
      Top             =   210
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMsgBox.frx":1A6A
      Height          =   360
      Left            =   960
      TabIndex        =   1
      Top             =   210
      Width           =   3150
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   270
      Picture         =   "frmMsgBox.frx":1AB6
      Top             =   210
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrInfo As String
Private mstrCaption As String
Private mstrCmds As String
Private mvStyle As VbMsgBoxStyle

Public Function ShowMsgBox(ByVal strCaption As String, ByVal strInfo As String, ByVal strCmds As String, _
    frmParent As Object, Optional vStyle As VbMsgBoxStyle = vbQuestion) As String
'参数：strCaption=消息窗体标题
'      strInfo=具体提示内容,可用"^"表示换行,">"表示缩进。
'      strCmds=按钮描述,如"重试(&R),!忽略(&A),?取消(&C)"
'              至少要有两个按钮,"!"表示缺省定位按钮,"?"表示取消按钮
'              每个按钮文字最多支持4个汉字
'      vStyle=vbInformation,vbQuestion,vbExclamation,vbCritical
'返回：按钮文字,如"按钮2"(不包含()和&),如果按关闭或取消则返回""
    Dim intMouse As Integer
    
    mstrCaption = strCaption
    mstrInfo = strInfo
    mstrCmds = strCmds
    mvStyle = vStyle
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = 0
    Me.Show 1, frmParent
    Screen.MousePointer = intMouse
    
    ShowMsgBox = mstrCmds
End Function

Private Sub cmdDo_Click(Index As Integer)
    mstrCmds = Replace(Split(cmdDo(Index).Caption, "(")(0), "&", "")
    If cmdDo(Index).Cancel Then mstrCmds = ""
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    '直接按单键热键
    If (KeyCode >= vbKey0 And KeyCode <= vbKey9 _
        Or KeyCode >= vbKeyA And KeyCode <= vbKeyZ) And Shift = 0 Then
        For i = 0 To cmdDo.UBound
            If InStr(cmdDo(i).Caption, "&") > 0 Then
                If Mid(cmdDo(i).Caption, InStr(cmdDo(i).Caption, "&") + 1, 1) = Chr(KeyCode) Then
                    Call cmdDo_Click(i): Exit Sub
                End If
            End If
        Next
    ElseIf KeyCode = vbKeyEscape Then
        mstrCmds = "": Unload Me
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '点击窗体关闭按钮
    If UnloadMode = vbFormControlMenu Then mstrCmds = ""
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    
    '缺省定位到缺省按钮上
    For i = 0 To cmdDo.UBound
        If cmdDo(i).Tag = "Default" Then
            cmdDo(i).SetFocus: Exit Sub
        End If
    Next
    VBA.Beep
End Sub

Private Sub Form_Load()
    Dim arrCmds As Variant, i As Integer
    Dim lngCmdW As Long, lngCmdL As Long
    
    Me.Caption = mstrCaption
    lblInfo.Caption = Replace(Replace(mstrInfo, "^", vbCrLf), ">", "　　")
    arrCmds = Split(mstrCmds, ","): mstrCmds = ""
    Select Case mvStyle
        Case vbInformation
            imgIcon(0).Visible = True
        Case vbQuestion
            imgIcon(1).Visible = True
        Case vbExclamation
            imgIcon(2).Visible = True
        Case vbCritical
            imgIcon(3).Visible = True
    End Select
    
    Me.Height = lblInfo.Top + lblInfo.Height + 1150
    If Me.Height < 1800 Then Me.Height = 1800
    
    '根据按钮确定按钮宽度
    For i = 0 To UBound(arrCmds)
        If LenB(StrConv(Replace(Split(cmdDo(i).Caption, "(")(0), "&", ""), vbFromUnicode)) > 2 Then
            Me.cmdDo(0).Width = 1300: Exit For
        End If
    Next
    lngCmdW = (UBound(arrCmds) + 1) * (cmdDo(0).Width + 100)
    
    '确定窗体宽度和按钮整体位置
    Me.Width = lblInfo.Left + lblInfo.Width + 500
    If Me.Width < lblInfo.Left + lngCmdW + 500 Then
        Me.Width = lblInfo.Left + lngCmdW + 500
    End If
    If Me.Width < 4500 Then Me.Width = 4500
    lngCmdL = (Me.ScaleWidth - lngCmdW) / 2 + 200
        
    '加载按钮
    For i = 0 To UBound(arrCmds)
        If i > 0 Then Load cmdDo(i)
        
        cmdDo(i).Caption = arrCmds(i)
        If Left(cmdDo(i).Caption, 1) = "!" Or Mid(cmdDo(i).Caption, 2, 1) = "!" Then
            cmdDo(i).Caption = Replace(cmdDo(i).Caption, "!", "")
            cmdDo(i).Tag = "Default"
        End If
        If Left(cmdDo(i).Caption, 1) = "?" Or Mid(cmdDo(i).Caption, 2, 1) = "?" Then
            cmdDo(i).Caption = Replace(cmdDo(i).Caption, "?", "")
            cmdDo(i).Cancel = True
        End If
        cmdDo(i).Left = lngCmdL + (cmdDo(0).Width + 100) * i
        cmdDo(i).Top = Me.ScaleHeight - cmdDo(i).Height - 180
        cmdDo(i).Visible = True
    Next
End Sub

