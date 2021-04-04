VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRSearchElement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "按要素的检索条件"
   ClientHeight    =   4935
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7515
   Icon            =   "frmEPRSearchElement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.OptionButton optAsk 
      Caption         =   "满足任一条件(&2)"
      Height          =   180
      Index           =   1
      Left            =   4305
      TabIndex        =   18
      Top             =   780
      Width           =   1665
   End
   Begin VB.OptionButton optAsk 
      Caption         =   "满足全部条件(&1)"
      Height          =   180
      Index           =   0
      Left            =   2565
      TabIndex        =   17
      Top             =   780
      Value           =   -1  'True
      Width           =   1665
   End
   Begin VB.CommandButton cmdAppend 
      Caption         =   "添加(&A)"
      Height          =   350
      Left            =   6195
      TabIndex        =   10
      Top             =   3480
      Width           =   1200
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   6195
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   915
      Width           =   1200
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   2925
      Left            =   -5610
      TabIndex        =   15
      Top             =   915
      Visible         =   0   'False
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   5159
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Frame fraDefine 
      Caption         =   "检索条件定义:"
      Height          =   1050
      Left            =   90
      TabIndex        =   3
      Top             =   3855
      Width           =   7335
      Begin VB.ComboBox cboValue 
         Height          =   300
         Left            =   3255
         TabIndex        =   9
         Top             =   525
         Width           =   3960
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Left            =   90
         TabIndex        =   5
         Top             =   525
         Width           =   1950
      End
      Begin VB.ComboBox cboFormula 
         Height          =   300
         Left            =   2085
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   525
         Width           =   1140
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "项目(&I):"
         Height          =   180
         Left            =   90
         TabIndex        =   4
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblFormula 
         AutoSize        =   -1  'True
         Caption         =   "条件(&F):"
         Height          =   180
         Left            =   2085
         TabIndex        =   6
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         Caption         =   "值(&V):"
         Height          =   180
         Left            =   3255
         TabIndex        =   8
         Top             =   300
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "移除(&R)"
      Height          =   350
      Left            =   6195
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3090
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6195
      TabIndex        =   14
      Top             =   465
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6195
      TabIndex        =   13
      Top             =   90
      Width           =   1200
   End
   Begin VB.Frame fraCodex 
      Height          =   30
      Left            =   75
      TabIndex        =   12
      Top             =   630
      Width           =   5910
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   -135
      Top             =   4755
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
            Picture         =   "frmEPRSearchElement.frx":058A
            Key             =   "ITEM"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   2745
      Left            =   120
      TabIndex        =   2
      Top             =   1050
      Width           =   5835
      _cx             =   10292
      _cy             =   4842
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   105
      Picture         =   "frmEPRSearchElement.frx":09DC
      Top             =   75
      Width           =   480
   End
   Begin VB.Label lblConditions 
      AutoSize        =   -1  'True
      Caption         =   "检索条件列表(&L):"
      Height          =   180
      Left            =   90
      TabIndex        =   1
      Top             =   795
      Width           =   1440
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "设置病历中包含的“固定诊治要素”检索条件，以便精确地检索到希望的病历记录。"
      Height          =   360
      Left            =   885
      TabIndex        =   0
      Top             =   120
      Width           =   5040
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmEPRSearchElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gstrMatch As String

Const conCol项目ID As Integer = 0
Const conCol项目名 As Integer = 1
Const conCol类型 As Integer = 2
Const conCol关系式 As Integer = 3
Const conCol条件值 As Integer = 4

'窗体变量
Private mblnOK As Boolean

Public Function ShowMe(ByVal frmParent As Object, ByRef strTerms As String) As Boolean
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '参数： frmParent-父窗体
    '       strTerms-要求条款
    '返回：确定更改返回True；取消返回False
    '---------------------------------------------------
Dim aryTerm() As String, aryField() As String
Dim lngCount As Long
    
    If strTerms <> "" Then
        If Val(Left(strTerms, 1)) = 0 Then Me.optAsk(1).Value = True
        aryTerm = Split(Mid(strTerms, 3), "|")
        With Me.vfgThis
            .Redraw = flexRDNone
            For lngCount = 0 To UBound(aryTerm)
                .Rows = .Rows + 1
                aryField = Split(aryTerm(lngCount), ";")
                .TextMatrix(.Rows - 1, conCol项目ID) = Val(aryField(conCol项目ID))
                .TextMatrix(.Rows - 1, conCol项目名) = aryField(conCol项目名)
                .TextMatrix(.Rows - 1, conCol类型) = Val(aryField(conCol类型))
                .TextMatrix(.Rows - 1, conCol关系式) = aryField(conCol关系式)
                .TextMatrix(.Rows - 1, conCol条件值) = aryField(conCol条件值)
            Next
            .Redraw = flexRDDirect
            If .Rows > .FixedRows Then .Row = .Rows - 1
        End With
    End If
    
    Me.Show vbModal, frmParent
    
    If mblnOK Then
        With Me.vfgThis
            strTerms = ""
            For lngCount = .FixedRows To .Rows - 1
                strTerms = strTerms & _
                    "|" & .TextMatrix(lngCount, conCol项目ID) & _
                    ";" & .TextMatrix(lngCount, conCol项目名) & _
                    ";" & .TextMatrix(lngCount, conCol类型) & _
                    ";" & .TextMatrix(lngCount, conCol关系式) & _
                    ";" & .TextMatrix(lngCount, conCol条件值)
            Next
        End With
        If strTerms <> "" Then strTerms = IIf(Me.optAsk(0).Value, 1, 0) & strTerms
    End If
    ShowMe = mblnOK: Unload Me
End Function

Private Sub cboFormula_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cboValue_Change()
    ValidControlText cboValue
End Sub

Private Sub cboValue_GotFocus()
    Me.cboValue.SelStart = 0: Me.cboValue.SelLength = 100
End Sub

Private Sub cboValue_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmdAppend_Click()
Dim strTemp As String
    '细则定义正确性检查
    If Trim(Me.txtItem.Tag) <> Trim(Me.txtItem.Text) Or Trim(Me.txtItem.Text) = "" Then
        MsgBox "未指定明确的评估细则项目！", vbExclamation, gstrSysName
        Me.txtItem.SetFocus: Exit Sub
    End If
    If Trim(Me.cboFormula.Text) = "" Then
        MsgBox "未指定明确的评估细则关系式！", vbExclamation, gstrSysName
        Me.cboFormula.SetFocus: Exit Sub
    End If
    If Me.cboValue.Enabled Then
        If Trim(Me.cboValue.Text) = "" Then
            MsgBox "未指定的评估细则条件值！", vbExclamation, gstrSysName
            Me.cboValue.SetFocus: Exit Sub
        End If
        strTemp = zlVerifyForm
        If strTemp <> "" Then
            MsgBox strTemp, vbExclamation, gstrSysName
            Me.cboValue.SetFocus: Exit Sub
        End If
    End If
    
    '将细则添加到表格中
    With Me.vfgThis
        .Redraw = flexRDNone
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, conCol项目ID) = Val(Me.lblItem.Tag)
        .TextMatrix(.Rows - 1, conCol项目名) = Trim(Me.txtItem.Text)
        .TextMatrix(.Rows - 1, conCol类型) = Val(Me.lblFormula.Tag)
        .TextMatrix(.Rows - 1, conCol关系式) = Trim(Me.cboFormula.Text)
        .TextMatrix(.Rows - 1, conCol条件值) = Trim(Me.cboValue.Text)
        .Row = .Rows - 1
        .Col = conCol项目名
        .Redraw = flexRDDirect
    End With

    
    '清除细则定义控件内容，以便定义新的细则：
    Me.lblItem.Tag = ""
    Me.txtItem.Text = ""
    Me.txtItem.Tag = ""
    Me.lblValue.Tag = ""
    Me.cboValue.Text = ""
    Me.lblFormula.Tag = ""
    Me.cboFormula.Clear
    Me.vfgThis.SetFocus
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    mblnOK = True: Me.Hide
End Sub

Private Sub cmdRemove_Click()
Dim rsTemp As New ADODB.Recordset
    If Val(Me.vfgThis.TextMatrix(Me.vfgThis.Row, conCol项目ID)) = 0 Then Exit Sub
    
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select i.Id, i.编码, i.中文名, i.英文名, i.类型, i.长度, i.小数, i.单位, Decode(i.替换域, 2, '', i.数值域) As 数值域" & _
            " From 诊治所见项目 i" & _
            " Where i.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(Me.vfgThis.TextMatrix(Me.vfgThis.Row, conCol项目ID)))
    With rsTemp
        Me.lblItem.Tag = !ID
        Me.txtItem.Text = !中文名
        Me.txtItem.Tag = !中文名
        Me.lblFormula.Tag = IIf(IsNull(!类型), 0, !类型)
        Me.lblValue.Tag = "" & !数值域
        Me.cboValue.Tag = "" & !单位
        Call zlAdjustForm
    End With
    
    Err = 0: On Error GoTo 0
    With Me.vfgThis
        Me.cboFormula.Text = .TextMatrix(.Row, conCol关系式)
        Me.cboValue.Text = .TextMatrix(.Row, conCol条件值)
        .Rows = .Rows - 1: .Row = .Rows - 1
    End With
    Me.txtItem.SetFocus
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.lvwList.Visible Then
        Me.lvwList.Visible = False
        Me.txtItem.SetFocus
    Else
        Call cmdCancel_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim lngCount As Long
    With Me.vfgThis
        .Redraw = flexRDNone
        .Rows = 1: .Cols = 5
        For lngCount = 0 To .Cols - 1
            .ColAlignment(lngCount) = 1
        Next
        .TextMatrix(0, conCol项目ID) = "项目ID"
        .TextMatrix(0, conCol项目名) = "项目名"
        .TextMatrix(0, conCol类型) = "类型"
        .TextMatrix(0, conCol关系式) = "关系式"
        .TextMatrix(0, conCol条件值) = "条件值"
        
        .ColWidth(conCol项目ID) = 0
        .ColWidth(conCol项目名) = 1600
        .ColWidth(conCol类型) = 0
        .ColWidth(conCol关系式) = 900
        .ColWidth(conCol条件值) = .Width - .ColWidth(conCol项目名) - .ColWidth(conCol关系式) - 250
        .Redraw = flexRDDirect
    End With
    With Me.lvwList.ColumnHeaders
        .Clear
        .Add , "中文名", "中文名", 1800
        .Add , "编码", "编码", 1000
        .Add , "类型", "类型", 600
        .Add , "数值域", "数值域", 4000
    End With
    Me.lvwList.ColumnHeaders("编码").Position = 1
End Sub

Private Sub lvwList_DblClick()
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwList
        Me.lblItem.Tag = Mid(.SelectedItem.Key, 2)
        Me.txtItem.Text = Split(.SelectedItem.Tag, ",")(0)
        Me.txtItem.Tag = Split(.SelectedItem.Tag, ",")(0)
        Me.lblFormula.Tag = Split(.SelectedItem.Tag, ",")(1)
        Me.lblValue.Tag = .SelectedItem.SubItems(Me.lvwList.ColumnHeaders("数值域").Index - 1)
        Me.cboValue.Tag = Split(.SelectedItem.Tag, ",")(2)
        Call zlAdjustForm
        Me.txtItem.SetFocus
        Call zlCommFun.PressKey(vbKeyTab)
    End With
End Sub

Private Sub lvwList_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
        Call lvwList_DblClick
    End Select
End Sub

Private Sub lvwList_LostFocus()
    Me.lvwList.Visible = False
End Sub

Private Sub txtItem_Change()
    ValidControlText txtItem
End Sub

Private Sub txtItem_GotFocus()
    Me.txtItem.SelStart = 0: Me.txtItem.SelLength = 100
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
    If InStr(" ~!@#$^&*()+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select i.Id, i.编码, i.中文名, i.英文名, i.类型, i.长度, i.小数, i.单位, Decode(i.替换域, 2, '', i.数值域) As 数值域" & _
            " From 诊治所见项目 i" & _
            " Where i.类型 In (0, 1) And (i.编码 Like [1] || '%' Or i.中文名 Like '" & gstrMatch & "'|| [1] ||'%' Or Upper(i.英文名) Like '" & gstrMatch & "'|| [1] ||'%')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(Me.txtItem.Text))
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "未找到指定诊治要素！", vbExclamation, gstrSysName
            Me.txtItem.SelStart = 0: Me.txtItem.SelLength = 100
            Me.txtItem.SetFocus
            Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.lblItem.Tag = !ID
            Me.txtItem.Text = !中文名
            Me.txtItem.Tag = !中文名
            Me.lblFormula.Tag = IIf(IsNull(!类型), 0, !类型)
            Me.lblValue.Tag = IIf(IsNull(!数值域), "", !数值域)
            Me.cboValue.Tag = IIf(IsNull(!单位), "", !单位)
            Call zlAdjustForm
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            Exit Sub
        End If
        
        Me.lvwList.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwList.ListItems.Add(, "_" & !ID, !中文名 & IIf(IsNull(!英文名), "", "(" & !英文名 & ")"), "ITEM", "ITEM")
            objItem.SubItems(Me.lvwList.ColumnHeaders("编码").Index - 1) = !编码
            Select Case IIf(IsNull(!类型), 0, !类型)
            Case 0
                objItem.SubItems(Me.lvwList.ColumnHeaders("类型").Index - 1) = "数值"
            Case 1
                objItem.SubItems(Me.lvwList.ColumnHeaders("类型").Index - 1) = "文字"
            End Select
            objItem.SubItems(Me.lvwList.ColumnHeaders("数值域").Index - 1) = IIf(IsNull(!数值域), "", !数值域)
            objItem.Tag = !中文名 & "," & IIf(IsNull(!类型), 0, !类型) & "," & IIf(IsNull(!单位), "", !单位)
            .MoveNext
        Loop
        With Me.lvwList
            .ListItems(1).Selected = True
            .Left = Me.fraDefine.Left + Me.txtItem.Left
            .Top = Me.fraDefine.Top + Me.txtItem.Top - .Height
            .Visible = True
            .SetFocus
        End With
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfgThis_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub zlAdjustForm()
Dim lngCount As Long
    '-------------------------------------------------
    '调整条件表达式的可选范围
    '入参： 保存在Me.lblFormula.Tag中的数值类型，Me.lblValue.Tag中的数值域
    '-------------------------------------------------
    Dim aryValue() As String
    Me.cboValue.Clear
    Me.cboValue.Enabled = False
    Me.cboFormula.Clear
    Select Case Val(Me.lblFormula.Tag)
    Case 0  '数值
        If Me.cboValue.Tag = "" Then
            Me.lblValue.Caption = "值(&V):(数值型)"
        Else
            Me.lblValue.Caption = "值(&V):(数值型 单位:" & Me.cboValue.Tag & ")"
        End If
        Me.cboFormula.AddItem "等于"
        Me.cboFormula.AddItem "不等于"
        Me.cboFormula.AddItem "大于"
        Me.cboFormula.AddItem "小于"
        Me.cboFormula.AddItem "至多"
        Me.cboFormula.AddItem "至少"
        Me.cboFormula.AddItem "介于"
        Me.cboFormula.AddItem "存在"
        Me.cboFormula.AddItem "不存在"
        Me.cboFormula.ListIndex = 0
        Me.cboValue.Enabled = True
    Case 1  '文字
        Me.lblValue.Caption = "值(&V):(文字型)"
        Me.cboFormula.AddItem "等于"
        Me.cboFormula.AddItem "不等于"
        Me.cboFormula.AddItem "包含"
        Me.cboFormula.AddItem "不包含"
        Me.cboFormula.AddItem "存在"
        Me.cboFormula.AddItem "不存在"
        Me.cboFormula.ListIndex = 0
        Me.cboValue.Enabled = True
    Case Else
    End Select
    
    aryValue = Split(Me.lblValue.Tag, ";")
    For lngCount = LBound(aryValue) To UBound(aryValue)
        Me.cboValue.AddItem aryValue(lngCount)
    Next
End Sub

Private Function zlVerifyForm() As String
    '-------------------------------------------------
    '判断条件表达式数值输入的正确性
    '入参：保存在Me.lblFormula.Tag中的数值类型
    '       Me.lblValue.Tag中的数值域，
    '       Me.lblFormula.text中的关系式
    '       Me.lblValue.text中的输入
    '出参：正确返回""，否则返回错误信息
    '-------------------------------------------------
Dim aryValue() As String
Dim lngCount As Long
    zlVerifyForm = ""
    Select Case Val(Me.lblFormula.Tag)
    Case 0  '数值
        Select Case Me.cboFormula.Text
        Case "等于", "不等于", "大于", "小于", "至多", "至少"
            Me.cboValue.Text = Val(Me.cboValue.Text)
        Case "介于"
            aryValue = Split(Trim(Me.cboValue.Text), ",")
            If UBound(aryValue) <> 1 Then
                zlVerifyForm = "条件值未按“介于”要求规则“值1,值2”形式组织填写！": Exit Function
            End If
            Me.cboValue.Text = Val(aryValue(0)) & "," & Val(aryValue(1))
        Case "存在", "不存在"
            aryValue = Split(Trim(Me.cboValue.Text), ",")
            If UBound(aryValue) < 1 Then
                zlVerifyForm = "如果仅为单条件值，没必要采用“存在”或“不存在”的关系式！": Exit Function
            End If
            Me.cboValue.Text = ""
            For lngCount = LBound(aryValue) To UBound(aryValue)
                Me.cboValue.Text = Me.cboValue.Text & "," & Val(aryValue(lngCount))
            Next
            Me.cboValue.Text = Mid(Me.cboValue.Text, 2)
        End Select
    Case 1  '文字
        Select Case Me.cboFormula.Text
        Case "等于", "不等于", "包含", "不包含"
        Case "存在", "不存在"
            aryValue = Split(Trim(Me.cboValue.Text), ",")
            If UBound(aryValue) < 1 Then
                zlVerifyForm = "如果仅为单条件值，没必要采用“存在”或“不存在”的关系式！": Exit Function
            End If
        End Select
    End Select
End Function

