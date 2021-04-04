VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDefTreePage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择页面"
   ClientHeight    =   3420
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5250
   Icon            =   "frmDefTreePage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList ils16 
      Left            =   4290
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefTreePage.frx":000C
            Key             =   "default"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3345
      Left            =   60
      TabIndex        =   21
      Top             =   30
      Width           =   3600
      Begin VB.CommandButton cmdPage 
         Caption         =   "…"
         Height          =   240
         Left            =   3165
         TabIndex        =   4
         Top             =   570
         Width           =   255
      End
      Begin VB.TextBox txtPage 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   795
         MaxLength       =   20
         TabIndex        =   3
         Top             =   540
         Width           =   2655
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   2655
         TabIndex        =   10
         Text            =   "cbo"
         Top             =   1260
         Width           =   795
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   2
         Left            =   795
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1275
         Width           =   1125
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   795
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   900
         Width           =   2655
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "…"
         Height          =   240
         Left            =   1635
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1665
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Height          =   930
         Left            =   795
         ScaleHeight     =   870
         ScaleWidth      =   2595
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2340
         Width           =   2655
         Begin VB.Label lblSample 
            AutoSize        =   -1  'True
            Caption         =   "示例说明"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   15
            TabIndex        =   22
            Top             =   30
            Width           =   840
         End
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   795
         MaxLength       =   20
         TabIndex        =   1
         Top             =   180
         Width           =   2655
      End
      Begin VB.ComboBox cboIcon 
         Height          =   300
         Left            =   795
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1980
         Width           =   1710
      End
      Begin zl9NewQuery.ctlPicture UsrPicture 
         Height          =   435
         Index           =   1
         Left            =   2865
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1830
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   767
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "字体(&3)"
         Height          =   180
         Left            =   135
         TabIndex        =   5
         Top             =   975
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "字形(&4)"
         Height          =   180
         Left            =   135
         TabIndex        =   7
         Top             =   1350
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "大小(&5)"
         Height          =   180
         Left            =   1995
         TabIndex        =   9
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "示例(&8)"
         Height          =   180
         Left            =   135
         TabIndex        =   17
         Top             =   2355
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "颜色(&6)"
         Height          =   180
         Left            =   135
         TabIndex        =   11
         Top             =   1710
         Width           =   630
      End
      Begin VB.Label lblColor 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "▇▇▇▇"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   780
         TabIndex        =   12
         Top             =   1635
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "页面(&2)"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   585
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "名称(&1)"
         Height          =   180
         Left            =   120
         TabIndex        =   0
         Top             =   225
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "图标(&7)"
         Height          =   180
         Left            =   135
         TabIndex        =   14
         Top             =   2025
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3885
      TabIndex        =   20
      Top             =   585
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3885
      TabIndex        =   19
      Top             =   135
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4065
      Top             =   1545
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDefTreePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnFirst As Boolean
Private blnOK As Boolean
Private mKey As Long
Private mUpKey As Long
Private mlngPageKey As Long

Private Sub ApplySample()
    '---------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '---------------------------------------------------------------------------------------
    
    On Error GoTo errHand
    
    With lblSample
        .FontName = cbo(0).Text
        .FontSize = IIf(Val(cbo(1).Text) <= 0, 12, Val(cbo(1).Text))
        .FontBold = (cbo(2).Text = "粗体" Or cbo(2).Text = "粗斜体")
        .FontItalic = (cbo(2).Text = "斜体" Or cbo(2).Text = "粗斜体")
        .ForeColor = lblColor.ForeColor
        .Caption = IIf(txt.Text = "", "示例说明", txt.Text)
    End With
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function CheckValid() As Boolean
    '---------------------------------------------------------------------------------------
    '功能：检查输入数据的有效性
    '返回：若合法返回True;否则返回False
    '---------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim blnWarn As Boolean
                
    If Trim(txt.Text) = "" Then
        strTmp = "必须定义一个显示名称！"
        GoTo Finally
    End If
    
    If mlngPageKey <= 0 Then
        strTmp = "必须要选对一个对应页面！"
        GoTo Finally
    End If
        
    '超过8个汉字时，警告在显示时可能不正常
    If LenB(StrConv(txt.Text, vbFromUnicode)) > 16 Then
        strTmp = "标题超过了8个汉字或16个字母，显示时可能不能完全显示！"
        blnWarn = True
        GoTo Finally
    End If
    
    If Val(cbo(1).Text) <= 0 Or Val(cbo(1).Text) > 99 Then
        strTmp = "必须输入大于0且小于99的数字！"
        GoTo Finally
    End If
    
Finally:
    If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
            
    '如果仅仅是警告,则仍然合法
    CheckValid = (strTmp = "" Or blnWarn)
    
End Function

Private Sub cbo_Change(Index As Integer)
    cmdOK.Tag = "1"
    Call ApplySample
End Sub

Private Sub cbo_Click(Index As Integer)
    cmdOK.Tag = "1"
    Call ApplySample
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cbocboPage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    End If
End Sub

Private Sub cboIcon_Click()
    If mblnFirst Then Exit Sub
    
    '更新图片显示
    Dim rs As New ADODB.Recordset
    
    UsrPicture(1).Cls
    
    If cboIcon.ItemData(cboIcon.ListIndex) > 0 Then
        gstrSQL = "select 序号,宽度,高度,类型 from 咨询图片元素 where 序号=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(cboIcon.ItemData(cboIcon.ListIndex)))
        If rs.BOF = False Then
            Call UsrPicture(1).ShowPictureByFieldNew(rs!序号, rs!宽度 * Screen.TwipsPerPixelX, rs!高度 * Screen.TwipsPerPixelY, IIf(IsNull(rs!类型), 0, rs!类型))
        End If
        CloseRecord rs
    Else
        Call UsrPicture(1).ShowByIPictureDisp(ils16.ListImages(1).Picture, 16 * Screen.TwipsPerPixelX, 16 * Screen.TwipsPerPixelY)
    End If
    cmdOK.Tag = "1"
End Sub

Private Sub cboIcon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveData Then
        blnOK = True
        If mKey = 0 Then
            txt.Text = ""
            txt.SetFocus
            cmdOK.Tag = ""
        Else
            cmdOK.Tag = ""
            Unload Me
        End If
    End If
End Sub

Private Sub cmdOpen_Click()
    '设置指定单元格的文字颜色,可以一次指定多个单元格
    Dim i As Long
    Dim j As Long
    
    On Error Resume Next
    dlg.CancelError = True
    dlg.flags = &H1
    dlg.Color = lblColor.ForeColor
    dlg.ShowColor
    If Err.Number = 0 Then
        lblColor.ForeColor = dlg.Color
        Call ApplySample
        cmdOK.Tag = "1"
    Else
        Err.Clear
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdPage_Click()
    Dim strID As String
    Dim blnRe As Boolean
    Dim str名称 As String
    Dim str编码 As String
        
    gstrSQL = "Select 页面序号 AS id,上级序号 AS 上级id,页面名称 AS 名称,编码,末级 From 咨询页面目录 where 页面序号>0 Start with 上级序号 is null connect by prior 页面序号 =上级序号"
    str名称 = txtPage.Text
                
    strID = CStr(mlngPageKey)
    strID = IIf(Val(strID) = 0, "", strID)
    
    blnRe = frm树型选择.ShowTree(gstrSQL, strID, str名称, str编码, "", Me.Caption, "所有页面分类", , "", True)
    
    If blnRe Then       '新的本级的宽度
        txtPage.Text = str名称
        mlngPageKey = Val(strID)
        txtPage.ForeColor = &HFF0000
        txtPage.BackColor = &HE0E0E0
        txtPage.Tag = ""
        
        cmdOK.Tag = "1"
    End If
    
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    
    Dim lngLoop As Long
    Dim rsFont As New ADODB.Recordset
    
    With rsFont
        .Fields.Append "字体名称", adVarChar, 100
        .Open
        .ActiveConnection = Nothing
    End With
    
    For lngLoop = 0 To Screen.FontCount - 1
        rsFont.AddNew
        rsFont.Fields(0).Value = Screen.Fonts(lngLoop)
    Next
        
    rsFont.Sort = "字体名称"
    
    If rsFont.RecordCount > 0 Then
        rsFont.MoveFirst
        Do While Not rsFont.EOF
            cbo(0).AddItem rsFont.Fields(0).Value
            rsFont.MoveNext
        Loop
    End If
    
    cbo(0).Text = "黑体"
    
    cbo(1).AddItem "8"
    cbo(1).AddItem "9"
    cbo(1).AddItem "10"
    cbo(1).AddItem "11"
    cbo(1).AddItem "12"
    cbo(1).AddItem "14"
    cbo(1).AddItem "16"
    cbo(1).AddItem "18"
    cbo(1).AddItem "20"
    cbo(1).AddItem "22"
    cbo(1).AddItem "24"
    cbo(1).AddItem "26"
    cbo(1).AddItem "28"
    cbo(1).AddItem "36"
    cbo(1).AddItem "48"
    cbo(1).AddItem "72"
    cbo(1).ListIndex = 4
    
    cbo(2).AddItem "常规"
    cbo(2).AddItem "斜体"
    cbo(2).AddItem "粗体"
    cbo(2).AddItem "粗斜体"
    cbo(2).ListIndex = 0
    
    '加载图标集合
    cboIcon.AddItem "缺省图标"
    gstrSQL = "select 名称,序号 from 咨询图片元素 where 性质=3"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF
            cboIcon.AddItem IIf(IsNull(gRs!名称), "", gRs!名称)
            cboIcon.ItemData(cboIcon.NewIndex) = gRs!序号
            gRs.MoveNext
        Wend
    End If
    cboIcon.ListIndex = 0
    
    If mKey <> 0 Then
        gstrSQL = "select A.名称,A.页面,A.页面图标,A.字体,A.大小,A.字形,A.颜色,A.序号,A.父序号,B.页面名称,B.页面序号 from 咨询页面排列 A,咨询页面目录 B where A.页面=B.页面序号 and A.序号=[1]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mKey)
        If gRs.BOF = False Then
            txt.Text = IIf(IsNull(gRs!名称), "", gRs!名称)
            'cboPage.Text = IIf(IsNull(gRs!页面名称), "", gRs!页面名称)
            
            txtPage.Text = IIf(IsNull(gRs!页面名称), "", gRs!页面名称)
            mlngPageKey = IIf(IsNull(gRs!页面序号), 0, gRs!页面序号)
            
            cboIcon.ListIndex = FindCboIndex(cboIcon, IIf(IsNull(gRs!页面图标), 0, gRs!页面图标))
            If cboIcon.ListIndex = -1 Then cboIcon.ListIndex = 0
            
            On Error Resume Next
            lblColor.ForeColor = IIf(IsNull(gRs!颜色), &HFF0000, gRs!颜色)
            cbo(0).Text = IIf(IsNull(gRs!字体), "黑体", gRs!字体)
            cbo(1).Text = IIf(IsNull(gRs!大小), "12", gRs!大小)
            
            Select Case IIf(IsNull(gRs!字形), 1, gRs!字形)
            Case 1
                cbo(2).ListIndex = 0
            Case 2
                cbo(2).ListIndex = 1
            Case 3
                cbo(2).ListIndex = 2
            Case 4
                cbo(2).ListIndex = 3
            End Select
            
            If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).Text = "黑体"
            If cbo(2).ListCount > 0 And cbo(2).ListIndex = -1 Then cbo(2).Text = "常规"
        End If
    End If
    mblnFirst = False
    
    Call cboIcon_Click
    
    txtPage.ForeColor = &HFF0000
    txtPage.BackColor = &HE0E0E0
    txtPage.Tag = ""
    
    cmdOK.Tag = ""
    Call SelAll(txt)
End Sub

Private Sub Form_Load()
    blnOK = False
    mblnFirst = True
End Sub

Public Function ShowTreePageBox(frmMain As Form, ByVal Key As Long, ByVal UpKey As Long) As Boolean
    mKey = Key
    mUpKey = UpKey
    frmDefTreePage.Show 1, frmMain
    ShowTreePageBox = blnOK
    
End Function

Private Function SaveData() As Boolean
    Dim lng序号 As Long
    Dim bytFontForm As Byte
    
    If cmdOK.Tag = "1" Then
        
        If CheckValid = False Then Exit Function
        
        Select Case cbo(2).Text
        Case "常规"
            bytFontForm = 1
        Case "斜体"
            bytFontForm = 2
        Case "粗体"
            bytFontForm = 3
        Case "粗斜体"
            bytFontForm = 4
        End Select
        
        If mKey = 0 Then

            lng序号 = NextValue("咨询页面排列", "序号")
            gstrSQL = "zl_咨询页面排列_insert(" & lng序号 & "," & IIf(mUpKey = 0, "NULL", mUpKey) & ",'" & txt.Text & "'," & mlngPageKey & "," & IIf(cboIcon.ItemData(cboIcon.ListIndex) = 0, "NULL", cboIcon.ItemData(cboIcon.ListIndex)) & ",'" & _
                                            IIf(cbo(0).Text = "", "黑体", cbo(0).Text) & "'," & _
                                            IIf(Val(cbo(1).Text) <= 0, 12, Val(cbo(1).Text)) & "," & _
                                            bytFontForm & "," & _
                                            lblColor.ForeColor & ")"
        Else
            lng序号 = mKey
            gstrSQL = "zl_咨询页面排列_update(" & lng序号 & "," & IIf(mUpKey = 0, "NULL", mUpKey) & ",'" & txt.Text & "'," & mlngPageKey & "," & IIf(cboIcon.ItemData(cboIcon.ListIndex) = 0, "NULL", cboIcon.ItemData(cboIcon.ListIndex)) & ",'" & _
                                            IIf(cbo(0).Text = "", "黑体", cbo(0).Text) & "'," & _
                                            IIf(Val(cbo(1).Text) <= 0, 12, Val(cbo(1).Text)) & "," & _
                                            bytFontForm & "," & _
                                            lblColor.ForeColor & ")"
        End If
        
        On Error GoTo errHand
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        Call frmDefTree.AddLvwItem(lng序号)
    End If
    SaveData = True
    Exit Function
errHand:
    If ErrCenter() = -1 Then Resume
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Tag = "1" Then
        If MsgBox("更改后的显示页面信息必须保存才能生效" & vbCrLf & "确认不保存就退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = True
    End If
End Sub

Private Sub txt_Change()
    cmdOK.Tag = "1"
End Sub

Private Sub txt_GotFocus()
    SelAll txt
    zlCommFun.OpenIme True
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    End If
    If CheckIsInclude(UCase(Chr(KeyAscii)), "'") = True Then KeyAscii = 0
End Sub

Private Sub txt_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
End Sub

Private Sub txtPage_Change()
    txtPage.Tag = "Changed"
    mlngPageKey = 0
    txtPage.ForeColor = &H0&
    txtPage.BackColor = &H80000005
    cmdOK.Tag = "1"
End Sub

Private Sub txtPage_GotFocus()
    Call SelAll(txtPage)
End Sub

Private Sub txtPage_KeyPress(KeyAscii As Integer)
    Dim strInput As String
    Dim strColWidth As String
    Dim strColAlign As String
    Dim lngPostion  As Long
    Dim sglX As Single
    Dim sglY As Single
    
    If KeyAscii = vbKeyReturn Then
        If txtPage.Tag = "Changed" Then
            If InStr(txtPage.Text, "'") > 0 Then
                MsgBox "内容中有非法字符！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            strInput = "'%" & txtPage.Text & "%'"
            
            gstrSQL = "Select 编码,页面名称 AS 名称,简码,页面序号 From 咨询页面目录  where 页面序号>0 and 末级=1 AND (编码 Like [1] or 简码 Like [1] or 页面名称 Like [1])"
            Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strInput)
            If gRs.BOF = False Then
                If gRs.RecordCount = 1 Then
                    lngPostion = 1
                Else
                    strColWidth = "900;2300;900;0"
                    strColAlign = "1;1;1;1"
                    Call CalcXY(Me, txtPage.Left + 30, txtPage.Top + txtPage.Height, sglX, sglY)
                    lngPostion = frmSelectList.ShowSelectList(Me, sglX, sglY, 4800, 2400, gRs, strColWidth, strColAlign)
                End If
                If lngPostion > 0 Then
                    gRs.MoveFirst
                    gRs.Move lngPostion - 1
                                    
                    txtPage.Text = IIf(IsNull(gRs("名称")), "", gRs("名称"))
                    mlngPageKey = IIf(IsNull(gRs("页面序号")), 0, gRs("页面序号"))
                Else
                    mlngPageKey = 0
                    txtPage.Text = ""
                End If
            Else
                mlngPageKey = 0
                txtPage.Text = ""
            End If
            txtPage.ForeColor = &HFF0000
            txtPage.BackColor = &HE0E0E0
            txtPage.Tag = ""
        Else
            SendKeys "{TAB}"
            SendKeys "{TAB}"
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtPage_Validate(Cancel As Boolean)
    If txtPage.Tag = "Changed" Then Cancel = True
End Sub

Private Sub CalcXY(objFrm As Form, objX As Single, objY As Single, sglX As Single, sglY As Single)
    sglX = objFrm.Left + objX + Screen.TwipsPerPixelX
    sglY = objFrm.Top + objFrm.Height - objFrm.ScaleHeight + objY
    If sglX + 6030 > Screen.Width Then
        sglX = Screen.Width - 6030
    End If
    If sglY + 3195 > Screen.Height Then
        sglY = sglY - 3195
    End If
End Sub
