VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInsElement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "插入要素"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   Icon            =   "frmInsElement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3690
      Index           =   0
      Left            =   585
      TabIndex        =   35
      Tag             =   "1000"
      Top             =   1065
      Visible         =   0   'False
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   6509
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2505
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsElement.frx":058A
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsElement.frx":0B24
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsElement.frx":10BE
            Key             =   "item"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picVBar 
      BackColor       =   &H8000000C&
      Height          =   5850
      Left            =   3420
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5850
      ScaleWidth      =   30
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   6105
      Left            =   3555
      ScaleHeight     =   6105
      ScaleWidth      =   4620
      TabIndex        =   29
      Top             =   45
      Width           =   4620
      Begin VB.TextBox txtTip 
         Height          =   945
         Left            =   885
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         ToolTipText     =   "最长250个汉字"
         Top             =   3960
         Width           =   3285
      End
      Begin VB.CheckBox chkDyn 
         Caption         =   "自定义(&K)"
         Height          =   225
         Left            =   3120
         TabIndex        =   16
         Top             =   2498
         Width           =   1110
      End
      Begin VB.CheckBox chkItemMust 
         Caption         =   "必填要素"
         Height          =   210
         Left            =   2580
         TabIndex        =   25
         ToolTipText     =   "是否必填要素，在诊治所见项目中定义"
         Top             =   5340
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   3285
         TabIndex        =   28
         Top             =   5745
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "保存(&S)"
         Height          =   350
         Left            =   2190
         TabIndex        =   27
         Top             =   5745
         Width           =   1100
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "插入(&I)"
         Height          =   350
         Left            =   1080
         TabIndex        =   26
         Top             =   5745
         Width           =   1100
      End
      Begin VB.CheckBox chkProtect 
         Caption         =   "保留对象(&P)"
         Height          =   225
         Left            =   915
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   5355
         Width           =   1485
      End
      Begin VB.CheckBox chkToString 
         Caption         =   "自动转为文本(&X)"
         Height          =   225
         Left            =   2595
         TabIndex        =   23
         Top             =   5025
         Width           =   1710
      End
      Begin VB.ComboBox cbo替换域 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "frmInsElement.frx":1658
         Left            =   915
         List            =   "frmInsElement.frx":165A
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   4980
         Width           =   1560
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   2
         Left            =   105
         TabIndex        =   33
         Top             =   5655
         Width           =   4305
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   1
         Left            =   105
         TabIndex        =   32
         Top             =   2355
         Width           =   4305
      End
      Begin VB.CheckBox chk形态 
         Caption         =   "展开(&E)"
         Height          =   225
         Left            =   2130
         TabIndex        =   15
         Top             =   2498
         Width           =   945
      End
      Begin VB.OptionButton opt固定 
         Caption         =   "插入临时诊治要素(&A)"
         Height          =   180
         Index           =   0
         Left            =   1125
         TabIndex        =   1
         Top             =   585
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton opt固定 
         Caption         =   "插入固定诊治要素(&B)"
         Height          =   180
         Index           =   1
         Left            =   1125
         TabIndex        =   2
         Top             =   885
         Width           =   2775
      End
      Begin VB.TextBox txt单位 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1635
         Width           =   1080
      End
      Begin VB.TextBox txt值域 
         Height          =   660
         Left            =   915
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   2820
         Width           =   3285
      End
      Begin VB.ComboBox cbo类型 
         Height          =   300
         ItemData        =   "frmInsElement.frx":165C
         Left            =   915
         List            =   "frmInsElement.frx":165E
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1635
         Width           =   1080
      End
      Begin VB.TextBox txt名称 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   915
         MaxLength       =   40
         TabIndex        =   4
         Top             =   1275
         Width           =   3285
      End
      Begin VB.TextBox txt长度 
         Height          =   300
         Left            =   915
         MaxLength       =   3
         TabIndex        =   10
         Top             =   1995
         Width           =   1080
      End
      Begin VB.TextBox txt小数 
         Height          =   300
         Left            =   3120
         MaxLength       =   1
         TabIndex        =   12
         Top             =   1995
         Width           =   1080
      End
      Begin VB.ComboBox cbo表示 
         Height          =   300
         ItemData        =   "frmInsElement.frx":1660
         Left            =   915
         List            =   "frmInsElement.frx":1662
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2460
         Width           =   1125
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   0
         Left            =   105
         TabIndex        =   30
         Top             =   1155
         Width           =   4305
      End
      Begin VB.Label lblTip 
         AutoSize        =   -1  'True
         Caption         =   "提示(&M)"
         Height          =   180
         Left            =   195
         TabIndex        =   20
         Top             =   3960
         Width           =   630
      End
      Begin VB.Image imgNote 
         Height          =   480
         Left            =   150
         Picture         =   "frmInsElement.frx":1664
         Top             =   90
         Width           =   480
      End
      Begin VB.Label lbl要素性质 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "可自行设置新的临时要素，或从列表中选择要素作为临时或固定要素插入："
         Height          =   360
         Left            =   705
         TabIndex        =   0
         Top             =   120
         Width           =   3420
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl值域 
         AutoSize        =   -1  'True
         Caption         =   "值域(&V)"
         Height          =   180
         Left            =   195
         TabIndex        =   17
         Top             =   2880
         Width           =   630
      End
      Begin VB.Label lbl名称 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "名称(&N)"
         Height          =   180
         Left            =   195
         TabIndex        =   3
         Top             =   1335
         Width           =   630
      End
      Begin VB.Label lbl类型 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "类型(&T)"
         Height          =   180
         Left            =   195
         TabIndex        =   5
         Top             =   1695
         Width           =   630
      End
      Begin VB.Label lbl长度 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "长度(&L)"
         Height          =   180
         Left            =   195
         TabIndex        =   9
         Top             =   2055
         Width           =   630
      End
      Begin VB.Label lbl小数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "小数(&D)"
         Height          =   180
         Left            =   2415
         TabIndex        =   11
         Top             =   2055
         Width           =   630
      End
      Begin VB.Label lbl单位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位(&U)"
         Height          =   180
         Left            =   2415
         TabIndex        =   7
         Top             =   1695
         Width           =   630
      End
      Begin VB.Label lbl表示 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "表示(&F)"
         Height          =   180
         Left            =   195
         TabIndex        =   13
         Top             =   2520
         Width           =   630
      End
      Begin VB.Label lbl填写说明 
         AutoSize        =   -1  'True
         Caption         =   "以分号分隔填写可选的数值，例如：A;B;C;D"
         Height          =   390
         Left            =   915
         TabIndex        =   19
         Top             =   3555
         Width           =   3105
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.TabControl tbcKind 
      Height          =   5445
      Left            =   45
      TabIndex        =   34
      Top             =   300
      Width           =   2850
      _Version        =   589884
      _ExtentX        =   5027
      _ExtentY        =   9604
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmInsElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'################################################################################################################
'窗体变量
Private mblnOk As Boolean
Private mblnOnlyAutoElement As Boolean
'################################################################################################################
'临时变量
Dim EditMode As EditModeEnum
Dim Element As cEPRElement

'################################################################################################################
'自定义事件
Public Event pOK(Ele As cEPRElement)        '保存数据
Public Event pCancel()                      '取消修改
'################################################################################################################
'## 功能：  上级程序调用本窗体的接口函数，传递参数，并显示窗体
'##
'## 参数：  frmParent       :父窗体
'##         oElement        :传入的诊治要素对象
'##         blnExtend       :是否包括“展开”输入形态的设置
'##         blnCanProtect   :是否允许设置要素为“保留”对象
'##         blnOnlyAutoElement:只可使用自动替换要素
'################################################################################################################
Public Function ShowMe(ByRef frmParent As Object, _
    Optional oElement As cEPRElement, _
    Optional blnExtend As Boolean = True, _
    Optional blnCanProtect As Boolean = False, _
    Optional blnOnlyAutoElement As Boolean = False) As Boolean
    
Dim aryTemp() As String
Dim lngCount As Long
'    Me.picVBar.BackColor = Me.BackColor
'    Call Form_Resize
    If blnCanProtect Then
        '允许设置保留
        chkProtect.Enabled = True
    Else
        chkProtect.Enabled = False
    End If
    If blnExtend = False Then Me.chk形态.Visible = False
    mblnOnlyAutoElement = blnOnlyAutoElement
    
    '填写需要选择的数据
    aryTemp = Split("0-数值;1-文字", ";")
    Me.cbo类型.Clear
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo类型.AddItem aryTemp(lngCount)
    Next
    Me.cbo类型.ListIndex = 1
    
    aryTemp = Split("0-不处理;1-自动替换;2-字典项目", ";")
    Me.cbo替换域.Clear
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo替换域.AddItem aryTemp(lngCount)
    Next
    Me.cbo替换域.ListIndex = 0
    chkToString.Visible = False
    
    If oElement Is Nothing Then
        EditMode = cprEM_新增
        Set Element = New cEPRElement
        Call zlRefElementByObject(Element, True)
    Else
        EditMode = cprEM_修改
        Set Element = oElement.Clone(True)
        Call zlRefElementByObject(Element, True)
    End If
    
    '显示窗体
    Me.Show 1
    If mblnOk = False Then ShowMe = False: Exit Function
    
    '返回结果对象
    ShowMe = True
    Unload Me
End Function

Private Sub cbo表示_Click()
    Me.txt值域.Enabled = True
    Select Case Left(Me.cbo类型.Text, 1)
    Case 0
        Select Case Left(Me.cbo表示.Text, 1)
        Case 0: Me.lbl填写说明.Caption = "可以按“最小值;最大值”形式指定数值限制，例如：0;100"
        Case 1: Me.lbl填写说明.Caption = "可以按“最小值;最大值”形式指定数值限制，例如：0;100"
        Case 2: Me.lbl填写说明.Caption = "需要按分号(;)分隔指定排斥可选的不同数值，例如：1;3;5"
        End Select
    Case 1
        Select Case Left(Me.cbo表示.Text, 1)
        Case 0: Me.lbl填写说明.Caption = "自由文本输入，不需要设置值域限制": Me.txt值域.Enabled = False: Me.txt值域.Text = ""
        Case 2: Me.lbl填写说明.Caption = "需要按分号(;)分隔指定互相排斥可选的文字，例如：正常;异常"
        Case 3: Me.lbl填写说明.Caption = "需要按分号(;)分隔指定可选的数值，例如：畏光;眼痛;耳鸣"
        End Select
    End Select
    Select Case Left(Me.cbo表示.Text, 1)
        Case 0, 1
            chk形态.Enabled = False: Me.chk形态.Value = 0
            chkDyn.Enabled = False: chkDyn.Value = vbUnchecked
        Case 2, 3
            Me.chk形态.Enabled = True
            chkDyn.Enabled = True
    End Select
End Sub

Private Sub cbo表示_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo类型_Change()
    Me.picBack.Tag = ""
End Sub

Private Sub cbo类型_Click()
Dim aryTemp() As String
Dim lngCount As Long
    '0-数值；1-文字；2-日期；3-逻辑
    Me.txt小数.Enabled = Me.opt固定(0).Value
    Select Case Left(Me.cbo类型.Text, 1)
    Case 0
        aryTemp = Split("0-文本;1-上下", ";")
    Case 1
        Me.txt小数.Text = 0: Me.txt小数.Enabled = False
        aryTemp = Split("0-文本;2-单选;3-复选", ";")
    End Select
    Me.cbo表示.Clear
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo表示.AddItem aryTemp(lngCount)
    Next
    Me.cbo表示.ListIndex = 0
End Sub

Private Sub cbo类型_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo替换域_Click()
    If Me.cbo替换域.ListIndex = 2 Then  '字典项目
        Me.cbo表示.ListIndex = 0: Me.cbo表示.Enabled = False
        chkToString.Visible = False
    Else
        chkToString.Visible = (cbo替换域.ListIndex = 1) '替换项目
        Me.cbo表示.Enabled = True
    End If
End Sub

Private Sub chk形态_Click()
    If chk形态.Value = 1 Then
        txtTip.Text = "": txtTip.Enabled = False: lblTip.Enabled = False
    Else
        txtTip.Enabled = True: lblTip.Enabled = True
        If Not Element Is Nothing Then
            txtTip.Text = Element.提示
        End If
    End If
End Sub

Private Sub chk形态_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    RaiseEvent pCancel
    mblnOk = False
End Sub

Private Sub cmdInsert_Click()
Dim lngCount As Long
    If Me.opt固定(0).Value Then
        If Trim(Me.txt名称.Text) = "" Then MsgBox "请输入要素名称！", vbInformation, gstrSysName: Me.txt名称.SetFocus: Exit Sub
        If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > 40 Then MsgBox "名称超长（最多40个字符或20个汉字）！", vbInformation, gstrSysName: Me.txt名称.SetFocus: Exit Sub
        If LenB(StrConv(Trim(Me.txt单位.Text), vbFromUnicode)) > 10 Then MsgBox "单位超长（最多10个字符或5个汉字）！", vbInformation, gstrSysName: Me.txt单位.SetFocus: Exit Sub
        If Val(Me.txt长度.Text) = 0 Then MsgBox "未正确设置长度！", vbExclamation, gstrSysName: Me.txt长度.SetFocus: Exit Sub
        If Val(Me.txt小数.Text) <> 0 And Val(Me.txt长度.Text) - Val(Me.txt小数.Text) < 2 Then MsgBox "未正确设置长度！", vbExclamation, gstrSysName: Me.txt长度.SetFocus: Exit Sub
    Else
        If Val(Me.picBack.Tag) = 0 Then MsgBox "不是按规定选择的固定诊治要素！", vbExclamation, gstrSysName: Exit Sub
    End If
    Select Case Left(Me.cbo表示.Text, 1)
    Case 2, 3
        If Trim(Me.txt值域.Text) = "" Then MsgBox "单选复选类型，必须设置可选项目！", vbExclamation, gstrSysName: Me.txt值域.SetFocus: Exit Sub
    End Select
    
    '保存诊治要素相关信息，激发 pOK() 事件，提交修改！
    Dim aryTemp
    With Element
        .要素名称 = Trim(Me.txt名称.Text)
        .诊治要素ID = IIf(Me.opt固定(0).Value, 0, Val(Me.picBack.Tag))
        .要素类型 = Left(Me.cbo类型.Text, 1)
        .要素长度 = Val(Me.txt长度.Text)
        .要素小数 = IIf(.要素类型 = 0, Val(Me.txt小数.Text), 0)
        .要素单位 = Trim(Me.txt单位.Text)
        .要素表示 = Left(Me.cbo表示.Text, 1)
        .替换域 = IIf(Me.opt固定(0).Value, 0, Me.cbo替换域.ListIndex)
        .自动转文本 = IIf(Me.chk形态.Visible, IIf(Me.chkToString.Value = vbChecked, True, False), False)
        .必填 = Me.chkItemMust.Value
        .动态域 = chkDyn.Value
        .提示 = zl9ComLib.zlStr.ToVarchar(txtTip.Text, 500)
        If chkProtect.Enabled Then
            .保留对象 = IIf(chkProtect.Value = vbChecked, True, False)
        End If
        
        If .要素类型 = 0 Then
            Select Case .要素表示
            Case 0, 1
                If Trim(Me.txt值域.Text) = "" Then
                    .要素值域 = ""
                Else
                    aryTemp = Split(Trim(Me.txt值域.Text), ";")
                    .要素值域 = Val(aryTemp(0)) & ";" & Val(aryTemp(1))
                End If
            Case 2
                aryTemp = Split(Trim(Me.txt值域.Text), ";")
                For lngCount = 0 To UBound(aryTemp)
                    aryTemp(lngCount) = Val(aryTemp(lngCount))
                Next
                .要素值域 = Join(aryTemp(0), ";")
            Case Else
                .要素值域 = ""
            End Select
        Else
            Select Case .要素表示
            Case 2, 3
                .要素值域 = Trim(Me.txt值域.Text)
                If chkDyn.Value = 1 And InStr(.要素值域, "自定义") = 0 Then .要素值域 = .要素值域 & ";自定义"
            Case Else
                .要素值域 = ""
            End Select
        End If
        .输入形态 = IIf(Me.chk形态.Visible, Me.chk形态.Value, 0)
        
        If EditMode = cprEM_修改 Then
            If .输入形态 = 1 Then
                '展开形式默认文本内容：
                Dim T As Variant, i As Long, strContent As String
                T = Split(.要素值域, ";")
                For i = 0 To UBound(T)
                    strContent = strContent & IIf(.要素表示 = 3, "□", "○") & T(i) & IIf(i = UBound(T), "", "  ")   '○●□■
                Next
                .内容文本 = strContent
            Else
                .内容文本 = ""
            End If
        Else
            .内容文本 = ""
        End If
    End With
    RaiseEvent pOK(Element)
    
    mblnOk = True
    
    '重新初始化
    Set Element = New cEPRElement
    Call zlRefElementByObject(Element, True)
End Sub

Private Sub cmdOK_Click()
    Call cmdInsert_Click
    Me.Hide
End Sub

Private Sub Form_Load()
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node

    With Me.tbcKind
        .SetImageList Me.imgList
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .COLOR = xtpTabColorOffice2003
            .ShowIcons = True
            .Position = xtpTabPositionTop
        End With
    End With
    
    '调入已经设置的诊治所见性质
    Err = 0: On Error GoTo errHand
    gstrSQL = "select 编码,名称 from 诊治所见性质 order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmInsElement")
    With rsTemp
        Do While Not .EOF
            If .AbsolutePosition > Me.tvwClass.Count Then Load Me.tvwClass(.AbsolutePosition - 1)
            Me.tbcKind.InsertItem(.AbsolutePosition - 1, !编码 & "." & !名称, Me.tvwClass(.AbsolutePosition - 1).hwnd, 0).Tag = "" & !编码
            .MoveNext
        Loop
    End With
    
    Dim intKind As Long
    gstrSQL = "select ID,上级ID,编码,名称,简码" & _
            " From 诊治所见分类" & _
            " Where 性质 = [1]" & _
            " start with 上级ID is null" & _
            " connect by prior ID=上级ID"
    For intKind = 0 To Me.tvwClass.Count - 1
        Me.tvwClass(intKind).Nodes.Clear
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmInsElement", Val(Me.tbcKind.Item(intKind).Tag))
        With rsTemp
            Do While Not .EOF
                If IsNull(!上级ID) Then
                    Set objNode = Me.tvwClass(intKind).Nodes.Add(, , "_" & !ID, "[" & !编码 & "]" & !名称, "close")
                Else
                    Set objNode = Me.tvwClass(intKind).Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, "[" & !编码 & "]" & !名称, "close")
                End If
                objNode.Sorted = True
                objNode.Tag = IIf(IsNull(!简码), "", !简码)
                objNode.ExpandedImage = "expend"
                .MoveNext
            Loop
        End With
        If Me.tvwClass(intKind).Nodes.Count > 0 Then Me.tvwClass(intKind).Nodes(1).Selected = True
    Next
    If Me.tbcKind.ItemCount > 0 Then Me.tbcKind.Item(0).Selected = True
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
'    Dim lngHWarp As Long, lngWWarp As Long
'    lngHWarp = Me.Height - Me.ScaleHeight
'    lngWWarp = Me.Width - Me.ScaleWidth
'    With Me.picVBar
'        .Top = Me.ScaleTop: .Height = Me.ScaleHeight
'        If .Left < 0 Then .Left = 0
'        If .Left > 6000 Then .Left = 6000
'    End With
    With Me.tbcKind
        .Top = Me.ScaleTop: .Height = Me.ScaleHeight
        .Left = Me.ScaleLeft: .Width = Me.picBack.Left - .Left - 30
    End With
'    With Me.picBack
'        .Left = Me.picVBar.Left + Me.picVBar.Width
'        .Top = Me.ScaleTop
'    End With
'    Me.Width = Me.picBack.Left + Me.picBack.Width + lngWWarp
'    Me.Height = Me.picBack.Height + lngHWarp
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    imgList.ListImages.Clear
    ImageList_Destroy imgList.hImageList
End Sub

Private Sub opt固定_Click(Index As Integer)
    Me.txt名称.Enabled = Me.opt固定(0).Value
    Me.cbo类型.Enabled = Me.opt固定(0).Value
    Me.txt长度.Enabled = Me.opt固定(0).Value
    Me.txt小数.Enabled = Me.opt固定(0).Value
    If Me.opt固定(0).Value = True Then
        Me.cbo替换域.Tag = Me.cbo替换域.ListIndex: Me.cbo替换域.ListIndex = 0
    Else
        Me.cbo替换域.ListIndex = Val(Me.cbo替换域.Tag)
    End If
End Sub

Private Sub opt固定_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub



Private Sub picVBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 1 Then Me.picVBar.Left = Me.picVBar.Left + X: Me.picVBar.BackColor = RGB(192, 192, 192)
End Sub

Private Sub picVBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Me.picVBar.BackColor = Me.BackColor
'    If Button = 1 Then Call Form_Resize
End Sub

Private Sub tvwClass_DblClick(Index As Integer)
    If Me.tvwClass(Index).SelectedItem Is Nothing Then Exit Sub
    If Left(Me.tvwClass(Index).SelectedItem.Key, 1) <> "I" Then Exit Sub
    Call zlRefElementByString(Me.tvwClass(Index).SelectedItem.Tag)
    Me.opt固定(1).Value = True
End Sub

Private Sub tvwClass_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node

    If Node.Children > 0 Then Exit Sub
    If Left(Node.Key, 1) <> "_" Then Exit Sub
    
    Err = 0: On Error GoTo errHand
    gstrSQL = "select  ID,编码,中文名,类型,长度,小数,小数,单位,表示法,数值域,替换域,必填,动态域,临床意义 提示" & _
            " from 诊治所见项目 I" & _
            " where 分类ID=[1] And 类型 in (0,1)"
    If mblnOnlyAutoElement Then
        gstrSQL = gstrSQL & " and 替换域=1"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmInsElement", CLng(Mid(Node.Key, 2)))
    With rsTemp
        Do While Not .EOF
            Set objNode = Me.tvwClass(Index).Nodes.Add(Node.Key, tvwChild, "I" & !ID, "[" & !编码 & "]" & !中文名, "item")
            objNode.Tag = !中文名 & "|" & !ID & "|" & !类型 & "|" & !长度 & "|" & !小数 & "|" & !单位
            Select Case Val("" & !表示法)
            Case 5: objNode.Tag = objNode.Tag & "|1||0" & "|" & !替换域 & "|0|0|" & !必填 & "|" & NVL(!动态域, 0) & "|" & NVL(!提示, "")
            Case 4: objNode.Tag = objNode.Tag & "|2|" & !数值域 & "|0" & "|" & !替换域 & "|0|0|" & !必填 & "|" & NVL(!动态域, 0) & "|" & NVL(!提示, "")
            Case Else: objNode.Tag = objNode.Tag & "|" & !表示法 & "|" & !数值域 & "|0" & "|" & !替换域 & "|0|0|" & !必填 & "|" & NVL(!动态域, 0) & "|" & NVL(!提示, "")
            End Select
            .MoveNext
        Loop
    End With
    Node.Expanded = True
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtTip_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("%&_|'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt长度_Change()
    Me.picBack.Tag = ""
End Sub

Private Sub txt长度_GotFocus()
    Me.txt长度.SelStart = 0: Me.txt长度.SelLength = 100
End Sub

Private Sub txt长度_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt单位_Change()
    ValidControlText txt单位
End Sub

Private Sub txt单位_GotFocus()
    Me.txt单位.SelStart = 0: Me.txt单位.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt单位_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt名称_Change()
    ValidControlText txt名称
    Me.picBack.Tag = ""
End Sub

Private Sub txt小数_Change()
    Me.picBack.Tag = ""
End Sub

Private Sub txt小数_GotFocus()
    Me.txt小数.SelStart = 0: Me.txt小数.SelLength = 100
End Sub

Private Sub txt小数_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt名称_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt值域_Change()
    ValidControlText txt值域
    '去掉特殊字符：○●□■
    txt值域 = Replace(txt值域, "○", "")
    txt值域 = Replace(txt值域, "●", "")
    txt值域 = Replace(txt值域, "□", "")
    txt值域 = Replace(txt值域, "■", "")
    If cbo类型.ListIndex = 1 And Left(Me.cbo表示.Text, 1) <> 0 Then
        '文本，单选/复选
        On Error Resume Next
        Dim lngNum As Long, T As Variant
        T = Split(txt值域.Text, ";")
        txt长度.Text = Len(txt值域.Text) + (UBound(T) + 1) * 2 + 4
    End If
End Sub

Private Sub txt值域_GotFocus()
    If Me.cbo表示.ListIndex = 0 Then
        Call zlCommFun.OpenIme(False)
    End If
End Sub

Private Sub txt值域_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
    Case vbKeyReturn: KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    Case Else
        If Me.cbo类型.ListIndex = 0 Then
            If InStr("0123456789.;-", Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
    End Select
End Sub

'################################################################################################################
'## 功能：  将元素按照对象属性填写到编辑控件
'##
'## 参数：
'##         Element     :传入的诊治要素对象
'##         blnFromOut  :是否外部提供修改的元素
'################################################################################################################
Public Sub zlRefElementByObject(ByRef Ele As cEPRElement, Optional blnFromOut As Boolean)
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
    Dim intKind As Integer, lngItemID As Long
    lngItemID = Val(Ele.诊治要素ID)
    If lngItemID <> 0 Then
        gstrSQL = "Select c.性质, i.分类id, i.Id From 诊治所见项目 i, 诊治所见分类 c Where i.分类id = c.Id And i.Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmInsElement", lngItemID)
        For intKind = 0 To Me.tbcKind.ItemCount - 1
            If Val(Me.tbcKind.Item(intKind).Tag) = rsTemp!性质 Then
                Me.tbcKind.Item(intKind).Selected = True
                Err = 0: On Error Resume Next
                Set objNode = Nothing
                Set objNode = Me.tvwClass(intKind).Nodes("_" & rsTemp!分类id)
                If Not (objNode Is Nothing) Then Call tvwClass_NodeClick(intKind, objNode)
                
                Set objNode = Nothing
                Set objNode = Me.tvwClass(intKind).Nodes("I" & lngItemID)
                If Not (objNode Is Nothing) Then
                    objNode.Selected = True
                    objNode.EnsureVisible
                End If
                
                Exit For
            End If
        Next
    End If
    
    Dim strElement As String
    strElement = Ele.要素名称 & "|" & Ele.诊治要素ID & "|" & Ele.要素类型 & "|" & Ele.要素长度 & "|" & Ele.要素小数 & "|" & Ele.要素单位 _
                & "|" & Ele.要素表示 & "|" & Ele.要素值域 & "|" & Ele.输入形态 & "|" & Ele.替换域 & "|" & IIf(Ele.自动转文本, 1, 0) _
                & "|" & IIf(Ele.保留对象, 1, 0) & "|" & Ele.必填 & "|" & Ele.动态域 & "|" & Ele.提示
    Call zlRefElementByString(strElement, blnFromOut)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'################################################################################################################
'## 功能：  将元素按照属性分解填写到编辑控件
'##
'## 参数：
'##         strElement  :按|分隔的元素属性串
'##         blnFromOut  :是否外部提供修改的元素
'################################################################################################################
Public Sub zlRefElementByString(ByVal strElement As String, Optional blnFromOut As Boolean)
Dim aryTemp() As String
Dim lngCount As Long
    aryTemp = Split(strElement, "|")
    Me.txt名称.Text = aryTemp(0)
    Me.cbo类型.ListIndex = IIf(aryTemp(0) = "", 1, Val(aryTemp(2)))
    Me.txt长度.Text = Val(aryTemp(3))
    Me.txt小数.Text = Val(aryTemp(4))
    Me.txt单位.Text = aryTemp(5)
    For lngCount = 0 To Me.cbo表示.ListCount - 1
        If Val(Left(Me.cbo表示.List(lngCount), 1)) = Val(aryTemp(6)) Then
            Me.cbo表示.ListIndex = lngCount: Exit For
        End If
    Next
    Me.txt值域.Text = aryTemp(7)
    If UBound(aryTemp) >= 8 And Me.chk形态.Enabled Then Me.chk形态.Value = aryTemp(8)
    Me.cbo替换域.ListIndex = Val(aryTemp(9)): Me.cbo替换域.Tag = Val(aryTemp(9))
    Me.chkToString.Value = IIf(Val(aryTemp(10)) = 0, vbUnchecked, vbChecked)
    Me.chkToString.Visible = (Me.cbo替换域.ListIndex = 1)
    Me.chkProtect.Value = IIf(Val(aryTemp(11)) = 1, vbChecked, vbUnchecked)
    Me.chkItemMust.Value = aryTemp(12)
    Me.chkDyn.Value = Val(aryTemp(13))
    Me.txtTip.Text = IIf(chk形态.Value = 1, "", NVL(aryTemp(14), "")) '展开型要素不给于提示
    lblTip.Enabled = (chk形态.Value = 0)

    'ID，最后设置；避免在其他设置中被更改事件清除
    If blnFromOut Then
        If Val(aryTemp(1)) = 0 Then
            Me.opt固定(0).Value = True
        Else
            Me.opt固定(1).Value = True
        End If
    End If
    Me.picBack.Tag = Val(aryTemp(1))
End Sub
