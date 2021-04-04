VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmParagraph 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "段落"
   ClientHeight    =   5610
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5520
   Icon            =   "frmParagraph.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdDefault 
      Caption         =   "默认间距(&D)..."
      Height          =   350
      Left            =   180
      TabIndex        =   38
      Top             =   5130
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4200
      TabIndex        =   37
      Top             =   5130
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2985
      TabIndex        =   36
      Top             =   5130
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   4
      Left            =   120
      TabIndex        =   35
      Top             =   4980
      Width           =   5250
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   3
      Left            =   585
      TabIndex        =   33
      Top             =   3150
      Width           =   4785
   End
   Begin VB.TextBox txtLineSpacing 
      Enabled         =   0   'False
      Height          =   300
      Left            =   4170
      MaxLength       =   6
      TabIndex        =   28
      Top             =   2595
      Width           =   825
   End
   Begin VB.ComboBox cboSpaceRule 
      Height          =   300
      ItemData        =   "frmParagraph.frx":000C
      Left            =   2475
      List            =   "frmParagraph.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   2595
      Width           =   1470
   End
   Begin VB.TextBox txtSpaceAfter 
      Height          =   300
      Left            =   990
      MaxLength       =   6
      TabIndex        =   23
      Text            =   "0"
      Top             =   2595
      Width           =   825
   End
   Begin VB.TextBox txtSpaceBefore 
      Height          =   300
      Left            =   990
      MaxLength       =   6
      TabIndex        =   20
      Text            =   "0"
      Top             =   2190
      Width           =   825
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   2
      Left            =   585
      TabIndex        =   32
      Top             =   1995
      Width           =   4785
   End
   Begin VB.TextBox txtFirstIndent 
      Height          =   300
      Left            =   4170
      MaxLength       =   6
      TabIndex        =   16
      Top             =   1455
      Width           =   825
   End
   Begin VB.ComboBox cboSpecialIdent 
      Height          =   300
      ItemData        =   "frmParagraph.frx":0063
      Left            =   2475
      List            =   "frmParagraph.frx":0070
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1455
      Width           =   1470
   End
   Begin VB.TextBox txtRightIndent 
      Height          =   300
      Left            =   990
      MaxLength       =   6
      TabIndex        =   11
      Text            =   "0"
      Top             =   1455
      Width           =   825
   End
   Begin VB.TextBox txtLeftIndent 
      Height          =   300
      Left            =   990
      MaxLength       =   6
      TabIndex        =   8
      Text            =   "0"
      Top             =   1050
      Width           =   825
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   585
      TabIndex        =   31
      Top             =   840
      Width           =   4785
   End
   Begin VB.ComboBox cboStyle 
      Height          =   300
      ItemData        =   "frmParagraph.frx":008E
      Left            =   4140
      List            =   "frmParagraph.frx":00B0
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   330
      Width           =   1110
   End
   Begin VB.ComboBox cboAlignment 
      Height          =   300
      ItemData        =   "frmParagraph.frx":00FD
      Left            =   1365
      List            =   "frmParagraph.frx":010A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   330
      Width           =   1425
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   585
      TabIndex        =   0
      Top             =   180
      Width           =   4785
   End
   Begin zlRichEditor.Document docSample 
      Height          =   1545
      Left            =   315
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3315
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   2725
      BackColor       =   0
      WYSIWYG         =   0   'False
   End
   Begin MSComCtl2.UpDown udRightIndent 
      Height          =   300
      Left            =   1815
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1455
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txtRightIndent"
      BuddyDispid     =   196619
      OrigLeft        =   1755
      OrigTop         =   1575
      OrigRight       =   1995
      OrigBottom      =   1875
      Max             =   200
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udLeftIndent 
      Height          =   300
      Left            =   1815
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1050
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txtLeftIndent"
      BuddyDispid     =   196620
      OrigLeft        =   1830
      OrigTop         =   893
      OrigRight       =   2070
      OrigBottom      =   1178
      Max             =   200
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udFirstIndent 
      Height          =   300
      Left            =   5010
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1455
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txtFirstIndent"
      BuddyDispid     =   196617
      OrigLeft        =   1755
      OrigTop         =   1575
      OrigRight       =   1995
      OrigBottom      =   1875
      Max             =   200
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udSpaceAfter 
      Height          =   300
      Left            =   1815
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2595
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txtSpaceAfter"
      BuddyDispid     =   196615
      OrigLeft        =   1755
      OrigTop         =   1575
      OrigRight       =   1995
      OrigBottom      =   1875
      Max             =   100
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udSpaceBefore 
      Height          =   300
      Left            =   1815
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2190
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txtSpaceBefore"
      BuddyDispid     =   196616
      OrigLeft        =   1830
      OrigTop         =   893
      OrigRight       =   2070
      OrigBottom      =   1178
      Max             =   100
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udLineSpacing 
      Height          =   300
      Left            =   5010
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2595
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtLineSpacing"
      BuddyDispid     =   196613
      OrigLeft        =   1755
      OrigTop         =   1575
      OrigRight       =   1995
      OrigBottom      =   1875
      Max             =   50
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   0   'False
   End
   Begin VB.Label lblSample 
      AutoSize        =   -1  'True
      Caption         =   "预览"
      Height          =   180
      Left            =   165
      TabIndex        =   34
      Top             =   3075
      Width           =   360
   End
   Begin VB.Label lblLineSpacing 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "设置值(&A)"
      Height          =   180
      Left            =   4170
      TabIndex        =   27
      Top             =   2355
      Width           =   810
   End
   Begin VB.Label lblSpaceRule 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "行距模式(&N)"
      Height          =   180
      Left            =   2505
      TabIndex        =   25
      Top             =   2355
      Width           =   990
   End
   Begin VB.Label lblSpaceAfter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "段后(&E)"
      Height          =   180
      Left            =   330
      TabIndex        =   22
      Top             =   2655
      Width           =   630
   End
   Begin VB.Label lblSpaceBefore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "段前(&B)"
      Height          =   180
      Left            =   330
      TabIndex        =   19
      Top             =   2250
      Width           =   630
   End
   Begin VB.Label lblSpace 
      AutoSize        =   -1  'True
      Caption         =   "间距"
      Height          =   180
      Left            =   165
      TabIndex        =   18
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label lblFirstIndent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "缩进值(&Y)"
      Height          =   180
      Left            =   4170
      TabIndex        =   15
      Top             =   1215
      Width           =   810
   End
   Begin VB.Label lblSpecialIdent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "特殊格式(&S)"
      Height          =   180
      Left            =   2505
      TabIndex        =   13
      Top             =   1215
      Width           =   990
   End
   Begin VB.Label lblRightIndent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "右(&R)"
      Height          =   180
      Left            =   330
      TabIndex        =   10
      Top             =   1515
      Width           =   450
   End
   Begin VB.Label lblLeftIndent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "左(&L)"
      Height          =   180
      Left            =   330
      TabIndex        =   7
      Top             =   1110
      Width           =   450
   End
   Begin VB.Label lblIndent 
      AutoSize        =   -1  'True
      Caption         =   "缩进"
      Height          =   180
      Left            =   165
      TabIndex        =   6
      Top             =   765
      Width           =   360
   End
   Begin VB.Label lblStyle 
      AutoSize        =   -1  'True
      Caption         =   "大纲样式(&O)"
      Height          =   180
      Left            =   3105
      TabIndex        =   4
      Top             =   390
      Width           =   990
   End
   Begin VB.Label lblAlignment 
      AutoSize        =   -1  'True
      Caption         =   "对齐方式(&G)"
      Height          =   180
      Left            =   330
      TabIndex        =   2
      Top             =   390
      Width           =   990
   End
   Begin VB.Label lblGeneral 
      AutoSize        =   -1  'True
      Caption         =   "常规"
      Height          =   180
      Left            =   165
      TabIndex        =   1
      Top             =   105
      Width           =   360
   End
End
Attribute VB_Name = "frmParagraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const conSpaceValue As Integer = 12     '最小方式和精确方式下的默认的行间距数值
Const conSpaceMultiple As Single = 1.5  '多倍行距方式下的默认的行间距数值
Const conSampleStart As Long = 102      '示范段落的开始位置，修改时必须注意处理前一段和后一段文字的长度

Dim blnOK As Boolean
Dim blnInSel As Boolean
Dim intCount As Integer

Public Function ShowMe(curParagraph As cPara, Optional blnHideStyle As Boolean, Optional strSample As String) As Boolean
    '功能：显示本段落对话框
    '参数：
    '   curParagraph,需要设置的段落对象
    '   blnHideStyle,是否禁止大纲样式设置
    '   strSample,显示样本文字
    
    '示范段落文字处理
    If Trim(strSample) = "" Then strSample = "当前段落…"
    With Me.docSample
        .Text = ""
        For intCount = 1 To 20
            .Text = .Text & "前一段落…"
        Next
        .Text = .Text & vbCrLf & strSample & vbCrLf
        For intCount = 1 To 20
            .Text = .Text & "后一段落…"
        Next
        .SelStart = 0: .SelLength = Len(.Text): .ZoomFactor = 0.5
        .SelStart = 0: .SelLength = 100
        .Selection.Font.ForeColor = RGB(192, 192, 192)
        .Selection.Font.Protected = True
        .SelStart = Len(.Text) - 100: .SelLength = 100
        .Selection.Font.ForeColor = RGB(192, 192, 192)
        .Selection.Font.Protected = True
        .SelStart = 0 'conSampleStart
    End With
    
    '段落对象属性值获取
    With curParagraph
        If blnHideStyle Then
            Me.lblStyle.Visible = False: Me.cboStyle.Visible = False
        Else
            Me.cboStyle.ListIndex = Abs(.Style) - 1
        End If
        
        If .FirstLineIndent >= 0 Then
            Me.cboSpecialIdent.ListIndex = Sgn(.FirstLineIndent)
            If .FirstLineIndent = tomUndefined Then
                Me.txtFirstIndent = ""
            Else
                Me.txtFirstIndent.Text = Abs(.FirstLineIndent)
            End If
            If .LeftIndent = tomUndefined Then
                Me.txtLeftIndent = ""
            Else
                Me.txtLeftIndent.Text = .LeftIndent
            End If
        Else
            Me.cboSpecialIdent.ListIndex = 2
            If .FirstLineIndent = tomUndefined Then
                Me.txtFirstIndent = ""
            Else
                Me.txtFirstIndent.Text = Abs(.FirstLineIndent)
            End If
            If .LeftIndent = tomUndefined Then
                Me.txtLeftIndent = ""
            Else
                Me.txtLeftIndent.Text = .LeftIndent
            End If
        End If
        Me.txtRightIndent.Text = IIf(.RightIndent = tomUndefined, 0, .RightIndent)
        
        Me.txtSpaceBefore.Text = IIf(.SpaceBefore = tomUndefined, 0, .SpaceBefore)
        Me.txtSpaceAfter.Text = IIf(.SpaceAfter = tomUndefined, 0, .SpaceAfter)
        If .LineSpacing <> tomUndefined Then
            Me.cboSpaceRule.ListIndex = .LineSpacingRule
            Select Case Me.cboSpaceRule.ListIndex
            Case cprLSSignle, cprLS1pt5, cprLSDouble
                Me.txtLineSpacing.Text = ""
            Case cprLSAtLeast, cprLSExactly, cprLSMultiple
                Me.txtLineSpacing.Text = .LineSpacing
            End Select
        End If
        If .Alignment <> tomUndefined Then
            
            Me.cboAlignment.ListIndex = IIf(.Alignment > 2, 2, .Alignment)
        End If
    End With
    
    Me.docSample.ReadOnly = True
    blnOK = False
    Me.Show 1
    If blnOK = False Then Unload Me: ShowMe = False: Exit Function
    
    With Me.docSample
        .SelStart = conSampleStart
        .ReadOnly = False
        If blnHideStyle = False Then curParagraph.Style = .Selection.Para.Style
        
        Call curParagraph.SetIndents(.Selection.Para.FirstLineIndent, .Selection.Para.LeftIndent, .Selection.Para.RightIndent)
        
        curParagraph.SpaceBefore = .Selection.Para.SpaceBefore
        curParagraph.SpaceAfter = .Selection.Para.SpaceAfter
        
        If Me.cboSpaceRule.ListIndex <> -1 Then Call curParagraph.SetLineSpacing(.Selection.Para.LineSpacingRule, .Selection.Para.LineSpacing)

        If Me.cboAlignment.ListIndex <> -1 Then
            curParagraph.Alignment = IIf(.Selection.Para.Alignment > 2, 2, .Selection.Para.Alignment)
        End If
    End With
    
    ShowMe = True: Unload Me
End Function

Private Sub IndentModify()
    '功能：根据当前设置，更改缩进
    Dim sglFirst As Single, sglLeft As Single, sglRight As Single
    Select Case Me.cboSpecialIdent.ListIndex
    Case 0
        sglFirst = 0
        sglLeft = Val(Me.txtLeftIndent.Text)
    Case 1
        sglFirst = Val(Me.txtFirstIndent.Text)
        sglLeft = Val(Me.txtLeftIndent.Text)
    Case 2
        sglFirst = -1 * Val(Me.txtFirstIndent.Text)
        sglLeft = Val(Me.txtLeftIndent.Text) + Val(Me.txtFirstIndent.Text)
    End Select
    sglRight = Val(Me.txtRightIndent.Text)
    
    With Me.docSample
        .ReadOnly = False: .SelStart = conSampleStart
        Call .Selection.Para.SetIndents(sglFirst, sglLeft, sglRight)
        .ReadOnly = True
    End With
End Sub

Private Sub cboAlignment_Click()
    With Me.docSample
        .ReadOnly = False: .SelStart = conSampleStart
        .Selection.Para.Alignment = Me.cboAlignment.ListIndex
        .ReadOnly = True
    End With
End Sub

Private Sub cboAlignment_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub cboSpaceRule_Click()
    blnInSel = True
    Select Case Me.cboSpaceRule.ListIndex
    Case cprLSSignle, cprLS1pt5, cprLSDouble      '0-单倍行距,1-1.5倍行距,2-两倍行距,忽略Spacing的值。
        Me.txtLineSpacing.Enabled = False
        Me.udLineSpacing.Enabled = False
        Me.txtLineSpacing.Text = ""
    Case cprLSAtLeast, cprLSExactly      '3-最小行距为1行，否则显示精确值;4-精确行距
        Me.txtLineSpacing.Enabled = True
        Me.udLineSpacing.Enabled = True
        If Val(Me.txtLineSpacing.Text) < conSpaceValue Then Me.txtLineSpacing.Text = conSpaceValue
    Case cprLSMultiple      '5-多倍行距。按行数计算。如1.2表示行距为1.2倍标准行距。
        Me.txtLineSpacing.Enabled = True
        Me.udLineSpacing.Enabled = True
        Me.txtLineSpacing.Text = conSpaceMultiple
    End Select
    
    With Me.docSample
        .ReadOnly = False: .SelStart = conSampleStart
        Select Case Me.cboSpaceRule.ListIndex
        Case cprLSSignle
            'Call .Selection.Para.SetLineSpacing(Me.cboSpaceRule.ListIndex, 0)
            .Selection.Para.SetLineSpacing cprLSMultiple, 1#
        Case cprLS1pt5
            .Selection.Para.SetLineSpacing cprLSMultiple, 1.5
        Case cprLSDouble
            .Selection.Para.SetLineSpacing cprLSMultiple, 2#
        Case cprLSAtLeast, cprLSExactly, cprLSMultiple
            Call .Selection.Para.SetLineSpacing(Me.cboSpaceRule.ListIndex, Val(Me.txtLineSpacing.Text))
        End Select
        .ReadOnly = True
    End With
    blnInSel = False

End Sub

Private Sub cboSpaceRule_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub cboSpecialIdent_Click()
    blnInSel = True
    If Me.cboSpecialIdent.ListIndex = 0 Then
        Me.txtFirstIndent.Text = "": Me.txtFirstIndent.Enabled = False: Me.udFirstIndent.Enabled = False
    Else
        Me.txtFirstIndent.Enabled = True: Me.udFirstIndent.Enabled = True
        If Val(Me.txtFirstIndent.Text) = 0 Then
            Me.txtFirstIndent.Text = Me.docSample.DefaultTabStop
        End If
    End If
    Call IndentModify
    blnInSel = False
End Sub

Private Sub cboSpecialIdent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub cboStyle_Click()
    With Me.docSample
        .ReadOnly = False: .SelStart = conSampleStart
        .Selection.Para.Style = -1 * (Me.cboStyle.ListIndex + 1)
        .ReadOnly = True
    End With
End Sub

Private Sub cboStyle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    blnOK = False: Me.Hide
End Sub

Private Sub cmdDefault_Click()
    Dim strMsgInfo As String
    strMsgInfo = "是否将当前设置的段前、段后、行距作为默认间距保存？" & _
        vbCrLf & "此更改将影响新的文档。"
    If MsgBox(strMsgInfo, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
    
    With Me.docSample
        .SelStart = conSampleStart
        .ReadOnly = False
        SaveSetting UCase(App.ProductName), "PARA", UCase("SpaceBefore"), .Selection.Para.SpaceBefore
        SaveSetting UCase(App.ProductName), "PARA", UCase("SpaceAfter"), .Selection.Para.SpaceAfter
        SaveSetting UCase(App.ProductName), "PARA", UCase("LineSpacingRule"), .Selection.Para.LineSpacingRule
        SaveSetting UCase(App.ProductName), "PARA", UCase("LineSpacing"), .Selection.Para.LineSpacing
        .ReadOnly = True
    End With
End Sub

Private Sub cmdOK_Click()
    blnOK = True: Me.Hide
End Sub

Private Sub Form_Activate()
    Me.cboAlignment.SetFocus
End Sub

Private Sub txtFirstIndent_Change()
    If blnInSel Then Exit Sub
    Call IndentModify
End Sub

Private Sub txtFirstIndent_GotFocus()
    With Me.txtFirstIndent
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call OpenIme(False)
End Sub

Private Sub txtFirstIndent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890." & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtLeftIndent_Change()
    Call IndentModify
End Sub

Private Sub txtLeftIndent_GotFocus()
    With Me.txtLeftIndent
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call OpenIme(False)
End Sub

Private Sub txtLeftIndent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890." & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtLineSpacing_Change()
    If blnInSel Then Exit Sub
    With Me.docSample
        .ReadOnly = False: .SelStart = conSampleStart
        Select Case Me.cboSpaceRule.ListIndex
        Case cprLSSignle, cprLS1pt5, cprLSDouble
            Call .Selection.Para.SetLineSpacing(Me.cboSpaceRule.ListIndex, 0)
        Case cprLSAtLeast, cprLSExactly, cprLSMultiple
            Call .Selection.Para.SetLineSpacing(Me.cboSpaceRule.ListIndex, Val(Me.txtLineSpacing.Text))
        End Select
        .ReadOnly = True
    End With
End Sub

Private Sub txtLineSpacing_GotFocus()
    With Me.txtLineSpacing
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call OpenIme(False)
End Sub

Private Sub txtLineSpacing_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890." & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtRightIndent_Change()
    Call IndentModify
End Sub

Private Sub txtRightIndent_GotFocus()
    With Me.txtRightIndent
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call OpenIme(False)
End Sub

Private Sub txtRightIndent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890." & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtSpaceAfter_Change()
    With Me.docSample
        .ReadOnly = False: .SelStart = conSampleStart
        .Selection.Para.SpaceAfter = Val(Me.txtSpaceAfter.Text)
        .ReadOnly = True
    End With
End Sub

Private Sub txtSpaceAfter_GotFocus()
    With Me.txtSpaceAfter
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call OpenIme(False)
End Sub

Private Sub txtSpaceAfter_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890." & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtSpaceBefore_Change()
    With Me.docSample
        .ReadOnly = False: .SelStart = conSampleStart
        .Selection.Para.SpaceBefore = Val(Me.txtSpaceBefore.Text)
        .ReadOnly = True
    End With
End Sub

Private Sub txtSpaceBefore_GotFocus()
    With Me.txtSpaceBefore
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call OpenIme(False)
End Sub

Private Sub txtSpaceBefore_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890." & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
