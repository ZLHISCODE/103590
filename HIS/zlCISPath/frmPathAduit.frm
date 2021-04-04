VERSION 5.00
Begin VB.Form frmPathAduit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "审核"
   ClientHeight    =   3810
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6135
   Icon            =   "frmPathAduit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3720
      TabIndex        =   6
      Top             =   3285
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4815
      TabIndex        =   5
      Top             =   3285
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   6030
   End
   Begin VB.TextBox txtContent 
      Height          =   1140
      Left            =   345
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1815
      Width           =   5550
   End
   Begin VB.OptionButton optAduit 
      Caption         =   "审核不通过"
      Height          =   225
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Top             =   1005
      Width           =   1305
   End
   Begin VB.OptionButton optAduit 
      Caption         =   "审核通过"
      Height          =   225
      Index           =   0
      Left            =   330
      TabIndex        =   1
      Top             =   1005
      Value           =   -1  'True
      Width           =   1305
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   765
      Width           =   6030
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "通过或不通过的理由(&M):"
      Height          =   180
      Left            =   330
      TabIndex        =   8
      Top             =   1490
      Width           =   1980
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   180
      Picture         =   "frmPathAduit.frx":6852
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "仔细检查临床路径表单内容是否符合要求，决定通过或不通过。"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   810
      TabIndex        =   7
      Top             =   180
      Width           =   5115
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPathAduit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng路径ID As Long
Private mlng版本号 As Long
                   
Private mintFunc  As Integer     '1=审核, 2=药剂科审核
Private mblnOK As Boolean

Public Function ShowAudit(ByVal frmParent As Object, ByVal lng路径ID As Long, ByVal lng版本号 As Long, ByVal intFunc As Integer) As Boolean
    On Error GoTo errHand
 
    mlng路径ID = lng路径ID
    mlng版本号 = lng版本号
    mintFunc = intFunc
    mblnOK = False
    Me.Show 1, frmParent
    
    ShowAudit = mblnOK
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim intType As Integer

    If zlCommFun.ActualLen(txtContent.Text) > 200 Then
        MsgBox "审核理由最多只允许 100 个汉字或 200 个字符。", vbInformation, gstrSysName
        txtContent.SetFocus: Exit Sub
    End If
    If optAduit(1).Value And Trim(txtContent.Text) = "" Then
        MsgBox "审核不通过时必须录入原因。", vbInformation, gstrSysName
        txtContent.SetFocus: Exit Sub
    End If
    
    If mintFunc = 1 Then '医务科审核
        intType = IIf(optAduit(0).Value = True, 1, 2)
    ElseIf mintFunc = 2 Then
        intType = IIf(optAduit(0).Value = True, 3, 4)
    End If
    
    On Error GoTo errH
    
    strSql = "Select Nvl(审核状态, 0) As 审核状态 From 临床路径目录 Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "查询路径审核状态", mlng路径ID)
    If rsTmp.RecordCount > 0 Then
       If (InStr(",1,2,", intType) > 0 And Not (Val(rsTmp!审核状态 & "") = 1 Or Val(rsTmp!审核状态 & "") = 2)) Or _
            (InStr(",3,4,", intType) > 0 And Val(rsTmp!审核状态 & "") <> 1) Then
           MsgBox "当前路径状态已改变不能进行审核,请刷新后再试!", vbInformation, gstrSysName
           Unload Me
           Exit Sub
       End If
    End If

    strSql = "Zl_临床路径审核_Insert(" & intType & "," & mlng路径ID & "," & mlng版本号 & ",'" & Trim(txtContent.Text) & "','" & UserInfo.姓名 & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "临床路径审核")
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub txtContent_GotFocus()
    zlControl.TxtSelAll txtContent
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtContent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtContent_LostFocus()
    Me.txtContent.Text = Replace(Me.txtContent, Chr(vbKeyReturn), "")
    Call zlCommFun.OpenIme(False)
End Sub

