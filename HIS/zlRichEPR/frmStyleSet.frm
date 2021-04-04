VERSION 5.00
Object = "*\A..\zlRichEditor\zlRichEdit.vbp"
Begin VB.Form frmStyleSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "常用样式设置"
   ClientHeight    =   5625
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame frmLine 
      Height          =   30
      Index           =   1
      Left            =   2730
      TabIndex        =   12
      Top             =   4950
      Width           =   2655
   End
   Begin VB.Frame frmLine 
      Height          =   30
      Index           =   0
      Left            =   165
      TabIndex        =   11
      Top             =   555
      Width           =   5220
   End
   Begin VB.TextBox txt编号 
      Height          =   300
      Left            =   855
      TabIndex        =   1
      Top             =   135
      Width           =   795
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Left            =   2655
      TabIndex        =   3
      Top             =   135
      Width           =   2685
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4260
      TabIndex        =   10
      Top             =   5115
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3165
      TabIndex        =   9
      Top             =   5115
      Width           =   1100
   End
   Begin zlRichEditor.Editor edt范例 
      Height          =   1890
      Left            =   150
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   930
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   3334
      Border          =   -1  'True
      PaperWidth      =   1907
      WithViewButtonas=   0   'False
      PaperKind       =   256
      ShowRuler       =   0   'False
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "字体(&F)…"
      Height          =   350
      Index           =   1
      Left            =   1260
      TabIndex        =   8
      Top             =   5115
      Width           =   1100
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "段落(&P)…"
      Height          =   350
      Index           =   0
      Left            =   150
      TabIndex        =   7
      Top             =   5115
      Width           =   1100
   End
   Begin VB.Label lbl范例 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "样式范例:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   150
      TabIndex        =   4
      Top             =   705
      Width           =   810
   End
   Begin VB.Label lbl编号 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "编号(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   150
      TabIndex        =   0
      Top             =   195
      Width           =   630
   End
   Begin VB.Label lbl名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "名称(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1935
      TabIndex        =   2
      Top             =   195
      Width           =   630
   End
   Begin VB.Label lbl描述 
      Caption         =   "###"
      Height          =   1935
      Left            =   150
      TabIndex        =   6
      Top             =   2895
      Width           =   5205
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmStyleSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngOldCode As Long, mlngNewCode As Long
Public Function ShowMe(ByVal frmParent As Object, ByVal blnAdd As Boolean, Optional ByVal lngCode As Long) As Long
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '参数： frmParent-父窗体
    '       blnAdd-是否增加
    '       lngCode-修改时需要指定的样式的编号
    '返回：确定返回新增或修改的编号；取消返回0
    '---------------------------------------------------
Dim rsTemp As New ADODB.Recordset
Dim aryFormat() As String

    If blnAdd Then
        Me.Tag = "新增": mlngOldCode = 0
    Else
        Me.Tag = "修改": mlngOldCode = lngCode
    End If
    Me.edt范例.Text = "样式设置范例"
    Me.edt范例.PaperWidth = 3000
    Me.edt范例.ResetWYSIWYG
    
    '原样式信息提取
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select 编号, 名称, 段落样式, 字体样式, 系统 From 病历常用样式 Where 编号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngOldCode)
    With rsTemp
        Me.txt编号.MaxLength = 3: Me.txt名称.MaxLength = .Fields("名称").DefinedSize
        If .RecordCount > 0 Then
            Me.txt编号.Text = Format(!编号, String(Me.txt编号.MaxLength, "0")): Me.txt名称 = "" & !名称
            If Val("" & !系统) = 1 Then Me.txt编号.Enabled = False: Me.txt名称.Enabled = False
            
            If "" & !字体样式 <> "" Then
                aryFormat = Split("" & !字体样式, ";")
                Me.edt范例.ForceEdit = True
                With Me.edt范例.Range(0, Len(Me.edt范例.Text)).Font
                    If Trim(aryFormat(0)) <> "" Then .Name = aryFormat(0)
                    If Val(aryFormat(1)) > 0 Then .Size = Val(aryFormat(1))
                    .Bold = IIf(Mid(aryFormat(2), 1, 1) = 1, True, False)
                    .Italic = IIf(Mid(aryFormat(2), 2, 1) = 1, True, False)
                    .Superscript = IIf(Mid(aryFormat(2), 7, 1) = 1, True, False)
                    .Subscript = IIf(Mid(aryFormat(2), 8, 1) = 1, True, False)
                    .ForeColor = Val(aryFormat(5))
                End With
                Me.edt范例.ForceEdit = False
            End If

            If "" & !段落样式 <> "" Then
                aryFormat = Split("" & !段落样式, ";")
                Me.edt范例.ForceEdit = True
                With Me.edt范例.Range(0, Len(Me.edt范例.Text)).Para
                    If Mid(aryFormat(0), 2, 1) < 9 Then .ListAlignment = Mid(aryFormat(0), 2, 1)                       '正常取值为：0、1、2
                    If Val(aryFormat(1)) <> -9999999 Then .Style = Val(aryFormat(1))        '正常取值： -1 ~ -10，需要先于行距设置，否则行距失效
                    
                    If Val(aryFormat(2)) <> -9999999 Then
                        .ListType = Val(aryFormat(2))     '正常取值：0 ～ 6、65536、131072、196608
                        .ListStart = Val(aryFormat(3))
                    End If
                    If Val(aryFormat(4)) <> tomUndefined Then .FirstLineIndent = Val(aryFormat(4)) '首行缩进一般是正数
                    If Val(aryFormat(5)) <> tomUndefined Then .LeftIndent = Val(aryFormat(5))
                    If Val(aryFormat(6)) <> tomUndefined Then .RightIndent = Val(aryFormat(6))
                    If Val(aryFormat(8)) <> tomUndefined Then .ListTab = Val(aryFormat(8))
                    If Val(aryFormat(9)) <> tomUndefined Then .SpaceBefore = Val(aryFormat(9))
                    If Val(aryFormat(10)) <> tomUndefined Then .SpaceAfter = Val(aryFormat(10))
                    
                    If Mid(aryFormat(0), 3, 1) < 9 And Val(aryFormat(7)) <> tomUndefined Then .SetLineSpacing Mid(aryFormat(0), 3, 1), Val(aryFormat(7))
                    If Mid(aryFormat(0), 1, 1) < 9 Then .Alignment = Mid(aryFormat(0), 1, 1)                           '正常取值为：0、1、2
                End With
                Me.edt范例.ForceEdit = False
            End If
            
        End If
    End With
    
    If Me.Tag = "新增" Then
        gstrSQL = "Select nvl(max(编号),'" & String(Me.txt编号.MaxLength, "0") & "') as 编号 From 病历常用样式"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        Me.txt编号.Text = Format(Val(rsTemp!编号) + 1, String(Me.txt编号.MaxLength, "0"))
    End If
    
    Me.lbl描述.Caption = zlStyleDesc
    '显示窗体
    Me.Show vbModal, frmParent
    ShowMe = mlngNewCode
    Unload Me
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = 0
End Function

Private Sub cmdCancel_Click()
    mlngNewCode = 0: Me.Hide
End Sub

Private Sub cmdOK_Click()
    '基本检查
    If Trim(Me.txt编号.Text) = "" Then MsgBox "请输入编号！", vbInformation, gstrSysName: Me.txt编号.SetFocus: Exit Sub
    If Len(Me.txt编号.Text) < Me.txt编号.MaxLength Then MsgBox "编号长度不足！", vbInformation, gstrSysName: Me.txt编号.SetFocus: Exit Sub
    If Trim(Me.txt名称.Text) = "" Then MsgBox "请输入名称！", vbInformation, gstrSysName: Me.txt名称.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > Me.txt名称.MaxLength Then
        MsgBox "名称超长（最多" & Me.txt名称.MaxLength & "个字符或等长的汉字）！", vbInformation, gstrSysName: Me.txt名称.SetFocus: Exit Sub
    End If
    
    '样式组织
    Dim strFont As String, strPara As String
    With Me.edt范例.Range(0, Len(Me.edt范例.Text)).Font
        strFont = .Name
        strFont = strFont & ";" & .Size
        strFont = strFont & ";" & Abs(CInt(.Bold)) & Abs(CInt(.Italic)) & Abs(CInt(.Hidden)) & Abs(CInt(.Protected)) _
                & Abs(CInt(.Link)) & Abs(CInt(.Strikethrough)) & Abs(CInt(.Superscript)) & Abs(CInt(.Subscript))
        strFont = strFont & ";" & .Underline
        strFont = strFont & ";" & .BackColor
        strFont = strFont & ";" & .ForeColor
    End With
    With Me.edt范例.Range(0, Len(Me.edt范例.Text)).Para
        strPara = .Alignment & .ListAlignment & .LineSpacingRule
        strPara = strPara & ";" & .Style
        strPara = strPara & ";" & .ListType
        strPara = strPara & ";" & .ListStart
        strPara = strPara & ";" & .FirstLineIndent
        strPara = strPara & ";" & .LeftIndent
        strPara = strPara & ";" & .RightIndent
        strPara = strPara & ";" & .LineSpacing
        strPara = strPara & ";" & .ListTab
        strPara = strPara & ";" & .SpaceBefore
        strPara = strPara & ";" & .SpaceAfter
    End With
    
    '保存操作
    If Me.Tag = "新增" Then
        gstrSQL = "Zl_病历常用样式_Insert(" & Trim(Me.txt编号.Text) & ",'" & Trim(Me.txt名称.Text) & "','" & strPara & "','" & strFont & "')"
    Else
        gstrSQL = "Zl_病历常用样式_Update(" & mlngOldCode & "," & Trim(Me.txt编号.Text) & ",'" & Trim(Me.txt名称.Text) & "','" & strPara & "','" & strFont & "')"
    End If
    Err = 0: On Error GoTo errHand
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    mlngNewCode = Val(Trim(Me.txt编号.Text)): Me.Hide
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSet_Click(Index As Integer)
    Dim blnSet As Boolean
    With Me.edt范例
        .SelStart = 0: .SelLength = Len(.Text)
        .ForceEdit = True
        If Index = 0 Then
            blnSet = .ShowParaDlg(False)
        Else
            blnSet = .ShowFontDlg(2 ^ 5 + 2 ^ 4 + 2 ^ 3 + 2 ^ 2 + 2 ^ 1 + 2 ^ 0)
        End If
        .ForceEdit = False
        .SelStart = 0
    End With
    Me.lbl描述.Caption = zlStyleDesc
End Sub

Private Sub Form_Activate()
    If Me.txt编号.Visible And Me.txt编号.Enabled Then Me.txt编号.SetFocus
End Sub

Private Sub txt编号_Change()
    txt编号 = Val(txt编号)
End Sub

Private Sub txt编号_GotFocus()
    Me.txt编号.SelStart = 0: Me.txt编号.SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt编号_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt名称_Change()
    ValidControlText txt名称
End Sub

Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Function zlStyleDesc() As String
    '根据当前样式填写样式说明文本
    Dim strStyle As String
    
    With Me.edt范例.Range(0, Len(Me.edt范例.Text)).Font
        strStyle = "字体样式:" & .Name
        strStyle = strStyle & ", 尺寸:" & .Size
        strStyle = strStyle & IIf(.Bold, ", 粗体", "")
        strStyle = strStyle & IIf(.Italic, ", 斜体", "")
        strStyle = strStyle & IIf(.Superscript, ", 上标", "")
        strStyle = strStyle & IIf(.Subscript, ", 下标", "")
        If .ForeColor = tomAutoColor Then
            strStyle = strStyle & ", 前景色:自动"
        Else
            strStyle = strStyle & ", 前景色:" & .ForeColor
        End If
    End With

    strStyle = strStyle & vbCrLf & vbCrLf & "段落样式:"
    With Me.edt范例.Range(0, Len(Me.edt范例.Text)).Para
        Select Case .Alignment   '正常取值为：0、1、2
        Case 0: strStyle = strStyle & "左对齐"
        Case 1: strStyle = strStyle & "居中"
        Case 2: strStyle = strStyle & "右对齐"
        End Select
        If .Style = cprPSNormal Then
            strStyle = strStyle & ", 大纲层次: 正文"
        Else
            strStyle = strStyle & ", 大纲层次: 标题" & Abs(.Style) - 1
        End If
        strStyle = strStyle & ", 首行缩进:" & .FirstLineIndent
        strStyle = strStyle & ", 左端缩进:" & .LeftIndent
        strStyle = strStyle & ", 右端缩进:" & .RightIndent
        
        Select Case .LineSpacingRule
        Case cprLSSignle:   strStyle = strStyle & ", 单倍行距"
        Case cprLS1pt5:     strStyle = strStyle & ", 1.5倍行距"
        Case cprLSDouble:   strStyle = strStyle & ", 两倍行距"
        Case cprLSAtLeast:   strStyle = strStyle & ", 最小行距:" & .LineSpacing
        Case cprLSExactly:   strStyle = strStyle & ", 精确行距:" & .LineSpacing
        Case cprLSMultiple:   strStyle = strStyle & ", 多倍行距:" & .LineSpacing
        End Select
        strStyle = strStyle & ", 段前间距:" & .SpaceBefore
        strStyle = strStyle & ", 段后间距:" & .SpaceAfter
    End With
    zlStyleDesc = strStyle
End Function
