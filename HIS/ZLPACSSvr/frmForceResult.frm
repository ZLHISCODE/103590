VERSION 5.00
Begin VB.Form frmBuildResult 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置自定义结果"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
   Icon            =   "frmForceResult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdVerify 
      Caption         =   "验证(&V)"
      Height          =   350
      Left            =   1080
      TabIndex        =   6
      Top             =   4080
      Width           =   1100
   End
   Begin VB.TextBox txtBuildResult 
      Height          =   3135
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   4695
   End
   Begin VB.ListBox lstDBItem 
      Height          =   3120
      ItemData        =   "frmForceResult.frx":000C
      Left            =   120
      List            =   "frmForceResult.frx":000E
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4920
      TabIndex        =   1
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3000
      TabIndex        =   0
      Top             =   4080
      Width           =   1100
   End
   Begin VB.Label Label2 
      Caption         =   "自定义结果："
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "可选的数据库项目："
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmBuildResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strReturnString As String

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    If funVerifyResult(Me.txtBuildResult.Text) = 0 Then
        strReturnString = Me.txtBuildResult.Text
        Unload Me
    End If
End Sub

Private Sub cmdVerify_Click()
    funVerifyResult Me.txtBuildResult.Text
End Sub

Public Function funVerifyResult(strString As String) As Integer
    '返回值：1-格式错误，[]不匹配；2－[]中不是数据库字段。
    Dim strTemp As String
    Dim strField As String
    Dim i As Integer, iPoint As Integer
    
    funVerifyResult = 0
    strTemp = strString
    
    '在字符串中，不允许出现'号
    If InStr(strTemp, "'") > 0 Then
        funVerifyResult = 1
        MsgBox "自定义结果格式出现非法字符，请修改后重新验证。"
        Exit Function
    End If
    
    iPoint = 1
    Do While iPoint <= Len(strTemp)
        iPoint = InStr(iPoint, strTemp, "[")
        If iPoint = 0 Then Exit Do
        
        If InStr(iPoint, strTemp, "]") = 0 Then
            funVerifyResult = 1
            Exit Do
        End If
        
        strField = Mid(strTemp, iPoint + 1, InStr(iPoint, strTemp, "]") - iPoint - 1)
        For i = 0 To Me.lstDBItem.ListCount - 1
            If strField = Me.lstDBItem.list(i) Then Exit For
        Next i
        If i >= Me.lstDBItem.ListCount Then
            funVerifyResult = 2
            Exit Do
        End If
'        strTemp = Right(strTemp, Len(strTemp) - InStr(strTemp, "]"))
        iPoint = iPoint + 1
    Loop
    
    '错误处理
    If funVerifyResult = 1 Then
        MsgBox "自定义结果格式有错误，“[”和“]”数量不匹配，请修改后重新验证。"
    ElseIf funVerifyResult = 2 Then
        MsgBox "自定义数据有错误，“[”和“]”中包含的文字不是系统提供的数据库字段，请修改后重新验证。"
    End If
End Function

Private Sub Form_Load()
    Me.lstDBItem.Clear
    Me.lstDBItem.AddItem "CallingAET"
    Me.lstDBItem.AddItem "首次日期"
    Me.lstDBItem.AddItem "首次时间"
    Me.lstDBItem.AddItem "影像类别"
    Me.lstDBItem.AddItem "执行间"
    Me.lstDBItem.AddItem "执行过程"
    Me.lstDBItem.AddItem "医嘱ID"
    Me.lstDBItem.AddItem "发送号"
    Me.lstDBItem.AddItem "检查号"
    Me.lstDBItem.AddItem "标识号"
    Me.lstDBItem.AddItem "英文名"
    Me.lstDBItem.AddItem "性别"
    Me.lstDBItem.AddItem "年龄"
    Me.lstDBItem.AddItem "出生日期"
    Me.lstDBItem.AddItem "中文名"
    Me.lstDBItem.AddItem "检查设备"
    strReturnString = ""
End Sub

Private Sub lstDBItem_DblClick()
Dim intStart As Integer
    Dim strTemp As String
    intStart = Me.txtBuildResult.SelStart
    strTemp = Me.txtBuildResult.Text
    Me.txtBuildResult.Text = Left(strTemp, intStart) & "[" & Me.lstDBItem.list(Me.lstDBItem.ListIndex) _
                            & "]" & Right(strTemp, Len(strTemp) - intStart)
    Me.txtBuildResult.SelStart = intStart + Len(Me.lstDBItem.list(Me.lstDBItem.ListIndex)) + 2
    Me.txtBuildResult.SetFocus
End Sub
