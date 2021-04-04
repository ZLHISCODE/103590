VERSION 5.00
Begin VB.Form frm病种选择_贵阳 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ICD疾病编码选择"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   Icon            =   "frm病种选择_贵阳.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt性别 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1080
      TabIndex        =   3
      Top             =   570
      Width           =   525
   End
   Begin VB.TextBox txt姓名 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1080
      TabIndex        =   1
      Top             =   180
      Width           =   1095
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1800
      TabIndex        =   6
      Top             =   1410
      Width           =   1100
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3060
      TabIndex        =   7
      Top             =   1410
      Width           =   1100
   End
   Begin VB.TextBox txt疾病信息 
      Height          =   300
      Left            =   1080
      TabIndex        =   5
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label lbl性别 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "性别"
      Height          =   180
      Left            =   660
      TabIndex        =   2
      Top             =   630
      Width           =   360
   End
   Begin VB.Label lbl姓名 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      Height          =   180
      Left            =   660
      TabIndex        =   0
      Top             =   240
      Width           =   360
   End
   Begin VB.Label lbl疾病信息 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ICD编码"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   390
      TabIndex        =   4
      Top             =   1020
      Width           =   630
   End
End
Attribute VB_Name = "frm病种选择_贵阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrICD编码 As String
Private mlng病人ID As Long

Public Function ChooseDisease(ByVal lng病人ID As Long) As String
    mlng病人ID = lng病人ID
    Me.Show 1
    ChooseDisease = mstrICD编码
End Function

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    mstrICD编码 = txt疾病信息.Tag
    
    If Trim(mstrICD编码) = "" Then
        MsgBox "必须要选择一种疾病！", vbInformation, gstrSysName
        txt疾病信息.SetFocus
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select 姓名,性别" & _
        " From 病人信息 " & _
        " Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人信息", mlng病人ID)
    
    With rsTemp
        Me.txt姓名 = !姓名
        Me.txt性别 = !性别
    End With
End Sub

Private Sub txt疾病信息_GotFocus()
    Call zlControl.TxtSelAll(txt疾病信息)
End Sub

Private Sub txt疾病信息_KeyPress(KeyAscii As Integer)
    Dim strLike As String
    Dim StrInput As String
    Dim str性别 As String
    Dim blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    If txt疾病信息.Text = lbl疾病信息.Tag And txt疾病信息.Text <> "" Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf txt疾病信息.Text = "" Then
        txt疾病信息.Tag = "": lbl疾病信息.Tag = ""
        Call zlCommFun.PressKey(vbKeyTab) '允许不输入
    Else
        strLike = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
        StrInput = UCase(txt疾病信息.Text)
        str性别 = txt性别.Text
        If str性别 = "男" Then
            str性别 = " And (A.性别限制='男' Or A.性别限制 is NULL)"
        ElseIf str性别 = "女" Then
            str性别 = " And (A.性别限制='女' Or A.性别限制 is NULL)"
        Else
            str性别 = ""
        End If
        gstrSQL = "Select A.ID,A.编码,A.附码,A.名称,A.简码,A.说明,A.性别限制,B.类别" & _
            " From 疾病编码目录 A,疾病编码类别 B" & _
            " Where A.类别=B.编码 And A.类别 Not IN('B','Z')" & _
            " And (A.编码 Like '" & StrInput & "%'" & _
            " Or Upper(A.名称) Like '" & strLike & StrInput & "%'" & _
            " Or Upper(A.简码) Like '" & strLike & StrInput & "%'" & _
            " Or Upper(A.附码) Like '" & strLike & StrInput & "%')" & _
            " And Rownum<=100" & str性别 & _
            " Order by A.类别,A.编码"
        Set rsTemp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "疾病编码Input", , , , , , True, _
            txt疾病信息.Left + Me.Left, _
            txt疾病信息.Top + Me.Top, txt疾病信息.Height, blnCancel, , True)
        If Not rsTemp Is Nothing Then
            txt疾病信息.Tag = rsTemp!编码
            txt疾病信息.Text = "(" & rsTemp!编码 & ")" & rsTemp!名称
            lbl疾病信息.Tag = txt疾病信息.Text '用于恢复显示
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            txt疾病信息.Tag = StrInput
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
