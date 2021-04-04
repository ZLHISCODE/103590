VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmData 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   Icon            =   "frmData.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdSetup 
      Caption         =   "列特性设置"
      Height          =   350
      Left            =   2640
      TabIndex        =   26
      Top             =   3075
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   2160
      Left            =   2655
      TabIndex        =   12
      Top             =   840
      Width           =   4500
      Begin VB.CommandButton cmdRelationSetup 
         Caption         =   "参数"
         Height          =   300
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "选择排序字段(F3)"
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox chkBold 
         Caption         =   "字体加粗"
         Height          =   255
         Left            =   2305
         TabIndex        =   22
         Top             =   1815
         Width           =   1095
      End
      Begin VB.CommandButton cmdAtt 
         Caption         =   "…"
         Height          =   285
         Left            =   1785
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1800
         Width           =   300
      End
      Begin VB.CommandButton cmdRelation 
         Height          =   270
         Left            =   3480
         Picture         =   "frmData.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1455
         Width           =   270
      End
      Begin VB.TextBox txtRelation 
         Height          =   300
         Left            =   945
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2820
      End
      Begin VB.OptionButton optDesc 
         Caption         =   "降序"
         Enabled         =   0   'False
         Height          =   180
         Left            =   2295
         TabIndex        =   6
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "升序"
         Enabled         =   0   'False
         Height          =   180
         Left            =   1305
         TabIndex        =   5
         Top             =   1200
         Width           =   810
      End
      Begin VB.CommandButton cmdOrder 
         Height          =   300
         Left            =   945
         Picture         =   "frmData.frx":0258
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "选择排序字段(F3)"
         Top             =   810
         Width           =   315
      End
      Begin VB.CommandButton cmdData 
         Height          =   300
         Left            =   945
         Picture         =   "frmData.frx":0432
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "选择表项数据(F2)"
         Top             =   375
         Width           =   315
      End
      Begin VB.TextBox txtOrder 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   810
         Width           =   2475
      End
      Begin VB.TextBox txtData 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   2475
      End
      Begin VB.TextBox txtFormat 
         BackColor       =   &H00EBFFFF&
         Height          =   300
         Left            =   945
         MaxLength       =   50
         TabIndex        =   7
         Top             =   810
         Width           =   2775
      End
      Begin VB.PictureBox picFont 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   960
         ScaleHeight     =   240
         ScaleWidth      =   1035
         TabIndex        =   23
         Top             =   1800
         Width           =   1095
         Begin VB.PictureBox picFontColor 
            BackColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   24
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Label lblFont 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "字体颜色"
         Height          =   180
         Left            =   165
         TabIndex        =   25
         Top             =   1860
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "关联报表"
         Height          =   180
         Left            =   165
         TabIndex        =   19
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "排序字段"
         Height          =   180
         Left            =   165
         TabIndex        =   15
         Top             =   885
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "表项数据"
         Height          =   180
         Left            =   165
         TabIndex        =   14
         Top             =   435
         Width           =   720
      End
      Begin VB.Label lblFormat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "格式串"
         Height          =   180
         Left            =   345
         TabIndex        =   13
         Top             =   870
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3840
      Left            =   2565
      TabIndex        =   11
      Top             =   -195
      Width           =   30
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   3135
      Left            =   90
      TabIndex        =   0
      Top             =   330
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   5530
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      PathSeparator   =   "."
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5955
      TabIndex        =   9
      Top             =   3075
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4365
      TabIndex        =   8
      Top             =   3075
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "排序字段用于对表项进行强行排序，其内容应该与该表项对应，如果选择不适当，将不能得到期望的结果。通常应该使用排序字段。"
      ForeColor       =   &H000000C0&
      Height          =   720
      Left            =   3300
      TabIndex        =   16
      Top             =   135
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   2900
      Picture         =   "frmData.frx":060C
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "数据源："
      Height          =   180
      Left            =   105
      TabIndex        =   10
      Top             =   90
      Width           =   720
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public frmParent As Object
Public I_strTitle As String '入：窗体标题
Public I_strClass As String '入：分类表格内容
Public I_strFormat As String '入：格式串内容
Public I_strOrder As String '入：排序字段内容
Public IO_strNode As String '入/出：选择的项目
Public I_bytType As Byte '入：0=纵向分类,1=横向分类,2=统计项,以决定可选项目类型
Public IO_FontBold As Integer '入/出：字体加粗
Public IO_FontColor As Long  '入/出：字体颜色
Public I_strSummaryFile As String  '汇总表所有数据项字段
Public objReport As Report
Public mintEleID As Integer             '选择元素的ID

Public mobjRelations As RPTRelations  '出/关联报表项目
Public mobjColProtertys As RPTColProtertys '出/列特性设置

Private Sub cmdAtt_Click()
    On Error Resume Next
    
    cdg.CancelError = True
    cdg.Flags = &H1 Or &H2
    cdg.Color = picFontColor.BackColor
    cdg.ShowColor
    If Err.Number = 0 Then
        picFontColor.BackColor = cdg.Color
    Else
        Err.Clear
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdData_Click()
    If tvw.SelectedItem.Parent Is Nothing Then VBA.Beep: Exit Sub
    If frmParent.mnuClass_State.Visible Then
        Select Case Val(tvw.SelectedItem.Tag)
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal _
                , adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt _
                , adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            MsgBox "“表项数据”不支持数字类型的字段！", vbInformation, App.Title
            Exit Sub
        End Select
    End If
    txtData.Text = tvw.SelectedItem.Text
    tvw.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim intType As Integer, objNode As Object
    Dim i As Long
    
    If txtData.Text = "" Then
        MsgBox "没有选择表格数据项目！", vbInformation, App.Title: tvw.SetFocus: Exit Sub
    End If
    If txtOrder.Text <> "" And txtOrder.Visible And Not optOrder.Value And Not optDesc.Value Then
        MsgBox "没有选择排序方式！", vbInformation, App.Title: optOrder.SetFocus: Exit Sub
    End If
    If LenB(StrConv(txtFormat.Text, vbFromUnicode)) > 50 Then
        MsgBox "该统计项的格式串过长,不能超过50个字符！", vbInformation, App.Title
        txtFormat.SetFocus: Exit Sub
    End If
    
    For Each objNode In tvw.Nodes
        If objNode.Text = txtData.Text Then intType = Val(objNode.Tag): Exit For
    Next
    
    If I_bytType = 2 And Not IsType(intType, adNumeric) Then
        MsgBox "统计项目请选择数字型项目！", vbInformation, App.Title: tvw.SetFocus: Exit Sub
    End If

    IO_FontColor = picFontColor.BackColor
    IO_FontBold = chkBold.Value
    gblnOK = True
    Hide
End Sub

Private Sub cmdOrder_Click()
    If tvw.SelectedItem.Parent Is Nothing Then VBA.Beep: Exit Sub
    If txtOrder.Text <> tvw.SelectedItem.Text Then optOrder.Value = True
    txtOrder.Text = tvw.SelectedItem.Text
    Call cmdData_Click
    tvw.SetFocus
End Sub

Private Sub cmdRelationSetup_Click()
    Call frmRelationSetup.ShowMe(Me, Val(txtRelation.Tag), I_strClass, objReport, txtRelation.Text, mobjRelations, 0)
End Sub

Private Sub cmdSetup_Click()
    Dim objSet As New frmColSetup
    
    Call objSet.ShowMe(Me, mobjColProtertys, 0, I_strClass, I_strSummaryFile, I_strSummaryFile)
End Sub

Private Sub Form_Activate()
    If tvw.Visible And tvw.Enabled Then tvw.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        cmdData_Click
    ElseIf KeyCode = vbKeyF3 And cmdOrder.Visible Then
        cmdOrder_Click
    End If
End Sub

Private Sub Form_Load()
    gblnOK = False
    Caption = I_strTitle
    Call CopySubTree
    If Not tvw.SelectedItem.Parent Is Nothing Then txtData.Text = IO_strNode
    If I_bytType <> 2 Then
        '分类项
        lblFormat.Visible = False
        txtFormat.Visible = False
        cmdSetup.Visible = False
        If I_strOrder <> "" Then
            If Left(I_strOrder, 1) = "," Then
                txtOrder.Text = Mid(I_strOrder, 2)
                optDesc.Value = True
            Else
                txtOrder.Text = I_strOrder
                optOrder.Value = True
            End If
        End If
    Else
        '统计项
        lblOrder.Visible = False
        txtOrder.Visible = False
        optOrder.Visible = False
        optDesc.Visible = False
        cmdOrder.Visible = False
        lblFont.Visible = False
        picFont.Visible = False
        cmdAtt.Visible = False
        chkBold.Visible = False
        txtFormat.Text = I_strFormat
        lblInfo.Caption = "统计项目是报表数据的最终统计结果。格式串则指明该统计项目输出时的格式形式，格式字符与VB的格式字符兼容。"
    End If
    txtRelation.Text = mobjRelations.Item(1).关联报表名称
    txtRelation.Tag = mobjRelations.Item(1).关联报表ID
    cmdRelationSetup.Enabled = Val(txtRelation.Tag) <> 0
    picFontColor.BackColor = IO_FontColor
    chkBold.Value = IO_FontBold
End Sub

Private Sub Form_Unload(Cancel As Integer)
    I_strTitle = ""
    I_strClass = ""
    I_strOrder = ""
    I_bytType = 0
End Sub


Private Sub tvw_DblClick()
    Call cmdData_Click
End Sub

Private Sub txtformat_GotFocus()
    SelAll txtFormat
End Sub

Private Sub tvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: cmdOK_Click
End Sub

Private Sub CopySubTree()
    Dim objNode As Object, tmpNode As Object
    
    For Each objNode In frmParent.tvwSQL.Nodes
        If mdlPublic.GetStdNodeText(objNode.Text) = I_strClass And objNode.Children <> 0 And objNode.Key <> "Root" Then Exit For
    Next
    
    tvw.Nodes.Clear
    Set tvw.ImageList = frmParent.tvwSQL.ImageList
    
    Set tmpNode = tvw.Nodes.Add(, , objNode.Key, objNode.Text, objNode.Image, objNode.SelectedImage)
    tmpNode.Expanded = True
    tmpNode.Selected = True
    
    Set objNode = objNode.Child
    Do While Not objNode Is Nothing
        If Not IsType(Val(objNode.Tag), adLongVarBinary) Then
            Set tmpNode = tvw.Nodes.Add(objNode.Parent.Key, 4, objNode.Key, objNode.Text, objNode.Image, objNode.SelectedImage)
            tmpNode.Tag = objNode.Tag
            If mdlPublic.GetStdNodeText(tmpNode.Text) = IO_strNode Then tmpNode.Selected = True
        End If
        Set objNode = objNode.Next
    Loop
End Sub

Private Sub txtFormat_KeyPress(KeyAscii As Integer)
    If InStr("!^|@'`~""", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txtOrder_Change()
    optDesc.Enabled = txtOrder.Text <> ""
    optOrder.Enabled = txtOrder.Text <> ""
End Sub

Private Sub txtOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        txtOrder.Text = ""
        optOrder.Value = False
        optDesc.Value = False
    End If
End Sub

Private Sub txtRelation_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strInfo As String
    Dim lngReportID As Long
    Dim intEleID As Integer
    
    intEleID = mintEleID
    If KeyCode = vbKeyReturn Then
        lngReportID = Val(FindReport(txtRelation.Text, txtRelation.hwnd, strInfo, Val(txtRelation.Tag), _
                                    objReport, mobjRelations, 1, Me, intEleID))
        If lngReportID <> 0 Then
            txtRelation.Text = strInfo
            txtRelation.Tag = lngReportID
        Else
            txtRelation.SetFocus
        End If
    End If
    
    cmdRelationSetup.Enabled = Val(txtRelation.Tag) <> 0
End Sub

Private Sub txtRelation_LostFocus()
    If txtRelation.Text = "" Then txtRelation.Tag = ""
    cmdRelationSetup.Enabled = Val(txtRelation.Tag) <> 0
End Sub

Private Sub cmdRelation_Click()
    Dim strInfo As String
    Dim lngReportID As Long
    Dim intEleID As Integer
    
    intEleID = mintEleID
    lngReportID = Val(FindReport("", txtRelation.hwnd, strInfo, Val(txtRelation.Tag), objReport, _
                                    mobjRelations, 1, Me, intEleID))
    
    If lngReportID <> 0 Then
        txtRelation.Text = strInfo
        txtRelation.Tag = lngReportID
    Else
        txtRelation.SetFocus
    End If
    
    cmdRelationSetup.Enabled = Val(txtRelation.Tag) <> 0
End Sub

