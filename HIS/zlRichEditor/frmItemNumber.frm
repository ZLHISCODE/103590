VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmItemNumber 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "项目符号或编号"
   ClientHeight    =   5985
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5610
   Icon            =   "frmItemNumber.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   Begin zlRichEditor.Document docSample 
      Height          =   1680
      Left            =   270
      TabIndex        =   27
      Top             =   3555
      Width           =   5145
      _extentx        =   9075
      _extenty        =   2963
      backcolor       =   0
      wysiwyg         =   0
   End
   Begin VB.CheckBox chkPlain 
      Caption         =   "无括号与句点(&N)"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "H:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   8
      EndProperty
      Height          =   195
      Left            =   3615
      TabIndex        =   11
      Top             =   1860
      Width           =   1665
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   3
      Left            =   630
      TabIndex        =   25
      Top             =   3330
      Width           =   4785
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   4
      Left            =   165
      TabIndex        =   24
      Top             =   5340
      Width           =   5250
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3030
      TabIndex        =   22
      Top             =   5490
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4245
      TabIndex        =   23
      Top             =   5490
      Width           =   1100
   End
   Begin VB.TextBox txtListTab 
      Height          =   300
      Left            =   2310
      MaxLength       =   6
      TabIndex        =   17
      Text            =   "0"
      Top             =   2775
      Width           =   1080
   End
   Begin VB.TextBox txtListStart 
      Height          =   300
      Left            =   3990
      MaxLength       =   6
      TabIndex        =   20
      Text            =   "1"
      Top             =   2775
      Width           =   1080
   End
   Begin VB.ComboBox cboAlignment 
      Height          =   300
      ItemData        =   "frmItemNumber.frx":000C
      Left            =   420
      List            =   "frmItemNumber.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2775
      Width           =   1575
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   615
      TabIndex        =   12
      Top             =   2310
      Width           =   4800
   End
   Begin VB.CheckBox chkPeriod 
      Caption         =   "添加句点(&P)"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "H:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   8
      EndProperty
      Height          =   195
      Left            =   2055
      TabIndex        =   10
      Top             =   1860
      Width           =   1530
   End
   Begin VB.CheckBox chkParenthese 
      Caption         =   "添加圆括号(&B)"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "H:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   8
      EndProperty
      Height          =   195
      Left            =   420
      TabIndex        =   9
      Top             =   1860
      Width           =   1530
   End
   Begin VB.OptionButton optListType 
      Caption         =   "大写罗马数字编号(&6)"
      Height          =   180
      Index           =   6
      Left            =   3105
      TabIndex        =   8
      Top             =   1455
      Width           =   2130
   End
   Begin VB.OptionButton optListType 
      Caption         =   "小写罗马数字编号(&5)"
      Height          =   180
      Index           =   5
      Left            =   3105
      TabIndex        =   7
      Top             =   1110
      Width           =   2130
   End
   Begin VB.OptionButton optListType 
      Caption         =   "大写字母编号(&4)"
      Height          =   180
      Index           =   4
      Left            =   3105
      TabIndex        =   6
      Top             =   765
      Width           =   2130
   End
   Begin VB.OptionButton optListType 
      Caption         =   "小写字母编号(&3)"
      Height          =   180
      Index           =   3
      Left            =   3105
      TabIndex        =   5
      Top             =   435
      Width           =   2130
   End
   Begin VB.OptionButton optListType 
      Caption         =   "阿拉伯数字编号(&2)"
      Height          =   180
      Index           =   2
      Left            =   420
      TabIndex        =   4
      Top             =   1110
      Width           =   2130
   End
   Begin VB.OptionButton optListType 
      Caption         =   "项目符号(&1)"
      Height          =   180
      Index           =   1
      Left            =   420
      TabIndex        =   3
      Top             =   765
      Width           =   2130
   End
   Begin VB.OptionButton optListType 
      Caption         =   "非项目符号与编号(&0)"
      Height          =   180
      Index           =   0
      Left            =   420
      TabIndex        =   2
      Top             =   435
      Value           =   -1  'True
      Width           =   2130
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   975
      TabIndex        =   0
      Top             =   195
      Width           =   4440
   End
   Begin MSComCtl2.UpDown udListStart 
      Height          =   300
      Left            =   5070
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2775
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtListStart"
      BuddyDispid     =   196616
      OrigLeft        =   1260
      OrigTop         =   4320
      OrigRight       =   1500
      OrigBottom      =   4620
      Max             =   1000
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udListTab 
      Height          =   300
      Left            =   3375
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2775
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txtListTab"
      BuddyDispid     =   196615
      OrigLeft        =   1830
      OrigTop         =   893
      OrigRight       =   2070
      OrigBottom      =   1178
      Max             =   200
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label lblSample 
      AutoSize        =   -1  'True
      Caption         =   "示范"
      Height          =   180
      Left            =   225
      TabIndex        =   26
      Top             =   3255
      Width           =   360
   End
   Begin VB.Label lblListTab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "符号文字间距(&T)"
      Height          =   180
      Left            =   2310
      TabIndex        =   16
      Top             =   2520
      Width           =   1350
   End
   Begin VB.Label lblListStart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始编号(&S)"
      Height          =   180
      Left            =   3990
      TabIndex        =   19
      Top             =   2520
      Width           =   990
   End
   Begin VB.Label lblAlignment 
      AutoSize        =   -1  'True
      Caption         =   "符号对齐方式(&A)"
      Height          =   180
      Left            =   420
      TabIndex        =   14
      Top             =   2520
      Width           =   1350
   End
   Begin VB.Label lblOther 
      AutoSize        =   -1  'True
      Caption         =   "其他"
      Height          =   180
      Left            =   225
      TabIndex        =   13
      Top             =   2235
      Width           =   360
   End
   Begin VB.Label lblListType 
      AutoSize        =   -1  'True
      Caption         =   "符号类型"
      Height          =   180
      Left            =   225
      TabIndex        =   1
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmItemNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const conSampleLen As Integer = 30     '每行示范文字的长度

Dim blnOK As Boolean
Dim lngListType As Long

Public Function ShowMe(curParagraph As cPara) As Boolean
    '功能：显示本段落对话框
    '参数：
    '   curParagraph,需要设置的段落对象
    
    lngListType = 0
    '示范段落文字处理
    With Me.docSample
        .Text = String(conSampleLen, "▂")
        .Text = .Text & vbCrLf & .Text & vbCrLf & .Text
        .Range(0, Len(.Text)).Font.ForeColor = RGB(128, 128, 128)
    End With
    
    '段落对象属性值获取
    With curParagraph
        Me.cboAlignment.ListIndex = .ListAlignment
        Me.txtListStart.Text = IIf(.ListStart = 0, 1, .ListStart)
        Me.txtListTab.Text = IIf(.ListTab = 0, 25, .ListTab)
        
        lngListType = .ListType
        If lngListType >= cprLTPlain Then
            lngListType = lngListType - cprLTPlain
            Me.chkPlain.Value = 1
        ElseIf lngListType >= cprLTPeriod Then
            lngListType = lngListType - cprLTPeriod
            Me.chkPeriod.Value = 1
        ElseIf lngListType >= cprLTParenthese Then
            lngListType = lngListType - cprLTParenthese
            Me.chkParenthese.Value = 1
        End If
        If lngListType >= cprLTNone And lngListType <= cprLTNumberAsUCRoman Then
            Me.optListType(lngListType).Value = True
        End If
    End With
    DisplayEffects
    
    Me.docSample.ReadOnly = True
    blnOK = False
    Me.Show 1
    If blnOK = False Then Unload Me: ShowMe = False: Exit Function
    
    With Me.docSample
        .SelStart = 0
        .ReadOnly = False
        If Me.cboAlignment.ListIndex <> -1 Then curParagraph.ListAlignment = .Selection.Para.ListAlignment
        curParagraph.ListStart = .Selection.Para.ListStart
        curParagraph.ListTab = .Selection.Para.ListTab
        curParagraph.ListType = .Selection.Para.ListType
    End With
    
    ShowMe = True: Unload Me
End Function

Private Sub DisplayEffects()
    '功能：当设置改变时显示效果
    Dim lngLength As Long
    lngLength = Len(Me.docSample.Text)
    
'    If Me.Visible = False Then Exit Sub
    With Me.docSample
        .ReadOnly = False
        
        Select Case lngListType
        Case 0
            .Range(0, lngLength).Para.ListType = cprLTNone
        Case 1
            .Range(0, lngLength).Para.ListTab = Val(Me.txtListTab.Text)
            .Range(0, lngLength).Para.ListType = cprLTBullet
        Case Else
            .Range(0, lngLength).Para.ListTab = Val(Me.txtListTab.Text)
            If Me.cboAlignment.ListIndex <> -1 Then .Range(0, lngLength).Para.ListAlignment = Me.cboAlignment.ListIndex
            .Range(0, lngLength).Para.ListStart = Val(Me.txtListStart.Text)
            If Me.chkPlain.Value = 1 Then
                .Range(0, lngLength).Para.ListType = lngListType + cprLTPlain
            ElseIf Me.chkParenthese.Value = 1 Then
                .Range(0, lngLength).Para.ListType = lngListType + cprLTParenthese
            ElseIf Me.chkPeriod.Value = 1 Then
                .Range(0, lngLength).Para.ListType = lngListType + cprLTPeriod
            Else
                .Range(0, lngLength).Para.ListType = lngListType
            End If
        End Select
        .ReadOnly = True
    End With
End Sub

Private Sub cboAlignment_Click()
    Call DisplayEffects
End Sub

Private Sub cboAlignment_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub chkParenthese_Click()
    If Me.chkParenthese.Value = 1 Then Me.chkPeriod.Value = 0: Me.chkPlain.Value = 0
    Call DisplayEffects
End Sub

Private Sub chkParenthese_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub chkPeriod_Click()
    If Me.chkPeriod.Value = 1 Then Me.chkParenthese.Value = 0: Me.chkPlain.Value = 0
    Call DisplayEffects
End Sub

Private Sub chkPeriod_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub chkPlain_Click()
    If Me.chkPlain.Value = 1 Then Me.chkParenthese.Value = 0: Me.chkPeriod.Value = 0
    Call DisplayEffects
End Sub

Private Sub chkPlain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    blnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
    blnOK = True: Me.Hide
End Sub

Private Sub optListType_Click(Index As Integer)
    lngListType = Index
    Select Case Index
    Case 0
        Me.chkParenthese.Value = 0: Me.chkParenthese.Enabled = False
        Me.chkPeriod.Value = 0: Me.chkPeriod.Enabled = False
        Me.chkPlain.Value = 0: Me.chkPlain.Enabled = False
        Me.cboAlignment.Enabled = False
        Me.txtListTab.Enabled = False: Me.udListTab.Enabled = False
        Me.txtListStart.Enabled = False: Me.udListStart.Enabled = False
    Case 1
        Me.chkParenthese.Value = 0: Me.chkParenthese.Enabled = False
        Me.chkPeriod.Value = 0: Me.chkPeriod.Enabled = False
        Me.chkPlain.Value = 0: Me.chkPlain.Enabled = False
        Me.cboAlignment.Enabled = False
        Me.txtListTab.Enabled = True: Me.udListTab.Enabled = True
        Me.txtListStart.Enabled = False: Me.udListStart.Enabled = False
    Case Else
        Me.chkParenthese.Enabled = True
        Me.chkPeriod.Enabled = True
        Me.chkPlain.Enabled = True
        Me.cboAlignment.Enabled = True
        Me.txtListTab.Enabled = True: Me.udListTab.Enabled = True
        Me.txtListStart.Enabled = True: Me.udListStart.Enabled = True
    End Select
    Call DisplayEffects
End Sub

Private Sub optListType_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub txtListStart_Change()
    If Val(Me.txtListStart.Text) > Me.udListStart.Max Then Me.txtListStart.Text = Me.udListStart.Max
    If Val(Me.txtListStart.Text) < Me.udListStart.Min Then Me.txtListStart.Text = Me.udListStart.Min
    Call DisplayEffects
End Sub

Private Sub txtListStart_GotFocus()
    With Me.txtListTab
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call OpenIme(False)
End Sub

Private Sub txtListStart_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtListTab_Change()
    If Val(Me.txtListTab.Text) > Me.udListTab.Max Then Me.txtListTab.Text = Me.udListTab.Max
    Call DisplayEffects
End Sub

Private Sub txtListTab_GotFocus()
    With Me.txtListTab
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call OpenIme(False)
End Sub

Private Sub txtListTab_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
