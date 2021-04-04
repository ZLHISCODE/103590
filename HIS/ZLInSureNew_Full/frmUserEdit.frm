VERSION 5.00
Begin VB.Form frmUserEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "用户编辑"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   1670
      TabIndex        =   10
      Top             =   1650
      Width           =   1100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   560
      TabIndex        =   9
      Top             =   1650
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -30
      TabIndex        =   8
      Top             =   1470
      Width           =   3405
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   1028
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1470
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1028
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1035
      Width           =   2115
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1028
      MaxLength       =   6
      TabIndex        =   3
      Top             =   585
      Width           =   2115
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1028
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   150
      Width           =   2115
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "权限"
      Height          =   180
      Index           =   1
      Left            =   548
      TabIndex        =   6
      Top             =   1530
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "确认密码"
      Height          =   180
      Index           =   1
      Left            =   188
      TabIndex        =   4
      Top             =   1095
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Index           =   0
      Left            =   548
      TabIndex        =   2
      Top             =   645
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "操作员"
      Height          =   180
      Index           =   0
      Left            =   368
      TabIndex        =   0
      Top             =   210
      Width           =   540
   End
End
Attribute VB_Name = "frmUserEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private isCancel As Boolean, strName As Long

Private Sub Command1_Click()
    Dim rsTemp As New ADODB.Recordset, strTemp As String, strUser As String, strPara As String
    If Text1.Text = "" Then
        MsgBox "信息设置不完整", vbInformation, "错误"
        Exit Sub
    End If
    
    strTemp = Mid(Combo1.Text, 2)
    strTemp = Left(strTemp, InStr(strTemp, "]") - 1)
    If strName = 0 Then
        Set rsTemp = gcn昭通.Execute("Select * From tab_czry Where hisid=" & strTemp)
    Else
        If strTemp <> strName Then
            Set rsTemp = gcn昭通.Execute("Select * From tab_czry Where hisid=" & strTemp)
        Else
            Set rsTemp = gcn昭通.Execute("Select * From tab_czry Where 1=2")
        End If
    End If
    If Not rsTemp.EOF Then
        MsgBox "选择的操作员已经具有其他权限", vbInformation, "错误"
        Exit Sub
    End If
    
    If Text1.Text <> Text2.Text Then
        MsgBox "密码输入错误", vbInformation, "错误"
        Exit Sub
    End If
    If Len(Text1.Text) < 6 Then
        MsgBox "密码必须输入6位，不能含空格", vbInformation, "错误"
        Exit Sub
    End If
    
    strUser = Mid(Combo1.Text, InStr(Combo1.Text, "]") + 1)
    If Combo1.Tag <> "" Then
        gcn昭通.Execute "Update tab_czry Set name='" & strUser & "',password='" & Text1.Text & "',POPEDOM=" & _
            "12290,hisid=" & strTemp & " Where oper=" & Combo1.Tag
    Else
        Set rsTemp = gcn昭通.Execute("Select nvl(max(oper),-1) from tab_czry")
        gcn昭通.Execute "Insert into tab_czry values (" & rsTemp(0) + 1 & ",'" & strUser & "','" & _
            Text1.Text & "',12290," & strTemp & ")"
        Combo1.Tag = rsTemp(0) + 1
    End If
    strPara = Combo1.Tag & vbTab & strUser & vbTab & Text1.Text & vbTab & 12290
    If frmConn昭通.Execute("I050", 0, strPara, "正在更新操作员数据......") = False Then Exit Sub
    
    Unload Me
End Sub

Private Sub Command2_Click()
    isCancel = True
    Unload Me
End Sub

Public Function userEdit(intEditMode As Integer, Optional intoper As Integer) As Boolean
    Dim rsTemp As New ADODB.Recordset
    strName = 0
'    Combo2.Enabled = True
    If intEditMode <> 0 Then
        Set rsTemp = gcn昭通.Execute("Select * from tab_czry where oper=" & intoper)
        strName = rsTemp!hisid
'        If intoper = 0 Then
'            Combo2.Enabled = False
'        Else
'            Combo2.Enabled = True
'        End If
    End If
    
    Set rsTemp = gcnOracle.Execute("Select a.ID,a.编号,a.姓名,a.简码 from 人员表 A,上机人员表 B where b.人员ID=a.id")
    While Not rsTemp.EOF
        Combo1.AddItem "[" & rsTemp!ID & "]" & rsTemp!姓名
        If rsTemp!ID = strName Then Combo1.ListIndex = Combo1.ListCount - 1
            
        rsTemp.MoveNext
    Wend
    
    If intEditMode <> 0 Then
        Set rsTemp = gcn昭通.Execute("Select * from tab_czry where oper=" & intoper)
        Text1.Text = rsTemp!password
        Text2.Text = Text1.Text
        Combo1.Tag = intoper
    End If
'    Combo2.AddItem "系统管理"
'    Combo2.AddItem "门诊业务"
'    Combo2.AddItem "住院业务"
'    If intEditMode <> 0 Then
'        If rsTemp!popedom = 2 Then
'            Combo2.ListIndex = 0
'        ElseIf rsTemp!popedom = 4096 Then
'            Combo2.ListIndex = 1
'        ElseIf rsTemp!popedom = 8192 Then
'            Combo2.ListIndex = 2
'        End If
'    End If
    
    On Error Resume Next
    Combo1.SetFocus
    isCancel = False
    Me.Show vbModal
    userEdit = Not isCancel
End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2.SetFocus
        Exit Sub
    End If
    If KeyAscii = vbKeyBack Then Exit Sub
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", UCase(Chr(KeyAscii))) > 0 Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1.SetFocus
        Exit Sub
    End If
    If KeyAscii = vbKeyBack Then Exit Sub
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", UCase(Chr(KeyAscii))) > 0 Then
        Exit Sub
    Else
        KeyAscii = 0
    End If

End Sub
