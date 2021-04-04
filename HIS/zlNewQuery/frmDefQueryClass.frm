VERSION 5.00
Begin VB.Form frmDefQueryClass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "页面分类"
   ClientHeight    =   1965
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4740
   Icon            =   "frmDefQueryClass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3510
      TabIndex        =   10
      Top             =   630
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3510
      TabIndex        =   9
      Top             =   150
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   75
      TabIndex        =   11
      Top             =   60
      Width           =   3195
      Begin VB.TextBox txt 
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   1065
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "11111"
         Top             =   270
         Width           =   900
      End
      Begin VB.TextBox txtTemp 
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   825
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "编码"
         Text            =   "1111111111"
         Top             =   225
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "…"
         Height          =   285
         Left            =   2775
         TabIndex        =   8
         Top             =   1350
         Width           =   285
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   825
         MaxLength       =   30
         TabIndex        =   3
         Top             =   615
         Width           =   1935
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   825
         MaxLength       =   15
         TabIndex        =   5
         Top             =   975
         Width           =   1935
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   3
         Left            =   825
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1350
         Width           =   1935
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "编码(&B)"
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   0
         Top             =   285
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   675
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "简码(&S)"
         Height          =   180
         Index           =   2
         Left            =   135
         TabIndex        =   4
         Top             =   1065
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "上级(&U)"
         Height          =   180
         Index           =   3
         Left            =   135
         TabIndex        =   6
         Top             =   1410
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmDefQueryClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const mlng编码长度 As Long = 10

Private mlngKey As Long
Private mlngUpKey As Long
Private mstr上级ID As String
Private mstr上级编码 As String
Private mstr编码 As String
Private mblnOK As Boolean

Private Function GetTreeCode(ByVal lngUpKey As Long) As Boolean
    '获取树型结构的编码规则,包括上级编码,本级编码
    
    If lngUpKey = 0 Then
        '如果没有指定上级
        mstr上级编码 = ""
        txtTemp.Text = ""
        
        txt(3).Text = "无"
        
        '取得上级编码，本级编码长度等值
        txtTemp.MaxLength = GetLocalCodeLength("", "咨询页面目录")
        
    Else
        '指定了上级
        gstrSQL = "select 编码 as 上级编码,页面名称 as 上级名称,页面序号 as 上级ID from 咨询页面目录 where 页面序号=[1]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngUpKey)
                        
        mstr上级ID = IIf(IsNull(gRs("上级ID")), "", gRs("上级ID"))
        mstr上级编码 = IIf(IsNull(gRs("上级编码")), "", gRs("上级编码"))
        txt(3).Text = IIf(IsNull(gRs("上级名称")), "无", gRs("上级名称"))
        txt(3).Tag = lngUpKey
        txtTemp.MaxLength = 0
        txtTemp.Text = mstr上级编码
        
        '判断编码是否满了
        If Len(mstr上级编码) >= mlng编码长度 Then
            MsgBox "不能再增加子分类了，编码长度已经用尽。", vbExclamation, gstrSysName
            Exit Function
        End If
        
        '取得上级编码，本级编码长度等值
        txtTemp.MaxLength = GetLocalCodeLength(mstr上级ID, "咨询页面目录")
    End If
        
    txt(0).MaxLength = IIf(txtTemp.MaxLength = 0, mlng编码长度, txtTemp.MaxLength) - Len(mstr上级编码)
    txt(0).Text = Mid(txt(0).Text, Len(txtTemp.Text) + 1)
    
    If mlngKey = 0 Then txt(0).Text = GetMaxLocalCode(mstr上级ID, "咨询页面目录")
    
    GetTreeCode = True
End Function

Public Function ShowEdit(ByVal frmParent As Form, ByVal lngKey As Long, ByVal lngUpKey As Long) As Boolean
    
    mblnOK = False
    
    mlngUpKey = lngUpKey
    mlngKey = lngKey
    
    If lngKey > 0 Then
        '修改分类
        gstrSQL = "Select 编码,页面名称,简码 from 咨询页面目录 where 页面序号=[1]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        If gRs.BOF = False Then
            txt(0).Text = IIf(IsNull(gRs("编码")), "", gRs("编码"))
            txt(1).Text = IIf(IsNull(gRs("页面名称")), "", gRs("页面名称"))
            txt(2).Text = IIf(IsNull(gRs("简码")), "", gRs("简码"))
            mstr编码 = txt(0).Text
        End If
    End If
    
    If GetTreeCode(mlngUpKey) = False Then Exit Function
                    
    cmdOK.Tag = ""
    
    Me.Show 1, frmParent
    
    ShowEdit = mblnOK
End Function

Private Function CheckValid() As Boolean
    txt(0).Text = Trim(txt(0).Text)

    If txtTemp.MaxLength = 0 Then
        If Len(txt(0).Text) = 0 Then
            MsgBox "编码不能为空。", vbExclamation, gstrSysName
            txt(0).SetFocus
            Exit Function
        End If
    Else
        If Len(txt(0).Text) < txt(0).MaxLength Then
            MsgBox "编码的长度不够。", vbExclamation, gstrSysName
            txt(0).SetFocus
            Exit Function
        End If
    End If
    If Not IsNumeric(txt(0).Text) Or InStr(txt(0).Text, ",") > 0 Or InStr(txt(0).Text, ".") > 0 Or InStr(txt(0).Text, "-") > 0 Then
        MsgBox "编码应由数字组成。", vbExclamation, gstrSysName
        txt(0).SetFocus
        Exit Function
    End If
    If Len(Trim(txt(1).Text)) = 0 Then
        MsgBox "名称不能为空。", vbExclamation, gstrSysName
        txt(1).Text = ""
        txt(1).SetFocus
        Exit Function
    End If
    
    CheckValid = True
End Function
Private Function SaveData() As Boolean
    Dim lng序号 As Long
    
    If cmdOK.Tag = "1" Then
                                
        If CheckValid = False Then Exit Function
                                        
        If mlngKey = 0 Then
            lng序号 = NextValue("咨询页面目录", "页面序号")
            gstrSQL = "zl_咨询页面目录_insert(" & lng序号 & ",'" & txt(1).Text & "',NULL,NULL,NULL,NULL,NULL," & IIf(Val(txt(3).Tag) = 0, "NULL", Val(txt(3).Tag)) & ",0,'" & txtTemp.Text & txt(0).Text & "','" & txt(2).Text & "')"
        Else
            lng序号 = mlngKey
            gstrSQL = "zl_咨询页面目录_update(" & lng序号 & ",'" & txt(1).Text & "',NULL,NULL,NULL,NULL," & IIf(Val(txt(3).Tag) = 0, "NULL", Val(txt(3).Tag)) & ",'" & txtTemp.Text & txt(0).Text & "','" & txt(2).Text & "'," & Len(mstr编码) + 1 & ")"
        End If
        
        On Error GoTo errHand
        
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        
        Call frmDefQuery.RefreshClass(CStr(lng序号))
    End If
    SaveData = True
    
    Exit Function
    
errHand:
    If ErrCenter() = -1 Then Resume
    
End Function

Private Sub cmdOpen_Click()
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim str名称 As String
    Dim strID As String
    Dim strRerurnID As String
    Dim str编码 As String
    Dim int编码  As Integer
    
    strSQL = "Select 页面序号 AS id,上级序号 AS 上级id,页面名称 AS 名称,编码,0 as 末级 From 咨询页面目录 Where (末级 IS NULL OR 末级=0)  Start with 上级序号 is null connect by prior 页面序号 =上级序号 "
    
    strID = txt(3).Tag
    str名称 = txt(3).Text
    str编码 = txtTemp.Text & txt(0).Text
        
    blnRe = frm树型选择.ShowTree(strSQL, strID, str名称, mstr上级编码, mlngKey, Me.Caption, "所有页面分类", , mstr编码)

    If blnRe Then       '新的本级的宽度
        
        mlngUpKey = Val(strID)
        txt(3).Tag = strID
        txt(3).Text = str名称
        Call GetTreeCode(mlngUpKey)
        txt(0).Text = GetMaxLocalCode(strID, "咨询页面目录")
        cmdOK.Tag = "1"
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Tag = "1" Then
        If MsgBox("更改后的查询目录必须保存才能生效" & vbCrLf & "确认不保存就退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveData Then
        mblnOK = True
        
        If mlngKey = 0 Then
            txt(0).Text = ""
            txt(1).Text = ""
            txt(2).Text = ""
            
            txt(0).Text = GetMaxLocalCode(txt(3).Tag, "咨询页面目录")
            
            txt(0).SetFocus
            cmdOK.Tag = ""
        Else
            cmdOK.Tag = ""
            Unload Me
        End If
    End If
    
End Sub

Private Sub txt_Change(Index As Integer)
    cmdOK.Tag = "1"
    If Index = 1 Then
        txt(2).Text = zlCommFun.SpellCode(txt(1).Text)
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Call SelAll(txt(Index))
    If Index = 0 Then zlCommFun.OpenIme True
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
        If Index = 3 Then SendKeys "{TAB}"
    Else
        If Index = 2 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If Index = 3 And Chr(KeyAscii) = "*" Then Call cmdOpen_Click
    End If
    
End Sub

Private Sub txt_LostFocus(Index As Integer)
    If Index = 0 Then zlCommFun.OpenIme
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub txtTemp_Change()
    txt(0).Width = txtTemp.Width - TextWidth(txtTemp.Text) - 120
    txt(0).Left = txtTemp.Left + TextWidth(txtTemp.Text) + 60
End Sub
