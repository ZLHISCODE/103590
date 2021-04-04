VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLabItemBase 
   BorderStyle     =   0  'None
   Caption         =   "项目基础信息"
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   Enabled         =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox chk多参考 
      Caption         =   "多参考项目"
      Height          =   210
      Left            =   6735
      TabIndex        =   56
      Top             =   525
      Width           =   1200
   End
   Begin VB.TextBox txt检验方法 
      Height          =   300
      Left            =   1215
      MaxLength       =   40
      ScrollBars      =   2  'Vertical
      TabIndex        =   50
      Top             =   3825
      Width           =   6630
   End
   Begin VB.TextBox txt名称拼音 
      Height          =   300
      Left            =   1215
      MaxLength       =   12
      TabIndex        =   54
      Top             =   1221
      Width           =   1335
   End
   Begin VB.TextBox txt项目名称 
      Height          =   300
      Left            =   1215
      MaxLength       =   60
      TabIndex        =   53
      Top             =   849
      Width           =   3975
   End
   Begin VB.TextBox txt项目编码 
      Height          =   300
      Left            =   1215
      MaxLength       =   13
      TabIndex        =   52
      Top             =   477
      Width           =   1335
   End
   Begin VB.TextBox txt英文缩写 
      Height          =   300
      Left            =   1215
      MaxLength       =   10
      TabIndex        =   51
      Top             =   1965
      Width           =   1335
   End
   Begin VB.TextBox txt取值序列 
      Height          =   300
      Left            =   1215
      MaxLength       =   200
      ScrollBars      =   2  'Vertical
      TabIndex        =   49
      ToolTipText     =   "(半定量和定性项目可设置取值序列，多个可选取值是采用“;”分隔)"
      Top             =   3453
      Width           =   6630
   End
   Begin VB.TextBox txt计算公式 
      Height          =   300
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   48
      Top             =   3081
      Width           =   6330
   End
   Begin VB.TextBox txt诊疗分类 
      Height          =   300
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   47
      Top             =   105
      Width           =   3660
   End
   Begin VB.ComboBox cbo标本类型 
      Height          =   300
      ItemData        =   "frmLabItemBase.frx":0000
      Left            =   1215
      List            =   "frmLabItemBase.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   2337
      Width           =   1365
   End
   Begin VB.TextBox txtAlias 
      Height          =   300
      Left            =   1215
      MaxLength       =   60
      TabIndex        =   45
      Top             =   1593
      Width           =   3975
   End
   Begin VB.ComboBox cbo试管 
      Height          =   300
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   2709
      Width           =   1365
   End
   Begin VB.TextBox txt排列序号 
      Height          =   300
      Left            =   6315
      MaxLength       =   14
      TabIndex        =   41
      Top             =   2337
      Width           =   1410
   End
   Begin VB.Frame fra酶标 
      Caption         =   "酶标项目公式"
      Height          =   1365
      Left            =   90
      TabIndex        =   33
      Top             =   4215
      Width           =   7770
      Begin VB.TextBox txtCutOff公式 
         Height          =   300
         Left            =   1110
         MaxLength       =   200
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   975
         Width           =   6555
      End
      Begin VB.TextBox txt弱阳性公式 
         Height          =   300
         Left            =   1110
         MaxLength       =   200
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   600
         Width           =   6555
      End
      Begin VB.TextBox txt阳性公式 
         Height          =   300
         Left            =   1110
         MaxLength       =   200
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   240
         Width           =   6555
      End
      Begin VB.Label lblCutOff公式 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CutOff公式"
         Height          =   180
         Left            =   165
         TabIndex        =   39
         Top             =   1020
         Width           =   900
      End
      Begin VB.Label lbl弱阳性公式 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "弱阳性公式"
         Height          =   180
         Left            =   165
         TabIndex        =   38
         Top             =   690
         Width           =   900
      End
      Begin VB.Label lbl阳性公式 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "阳性公式"
         Height          =   180
         Left            =   165
         TabIndex        =   37
         Top             =   315
         Width           =   900
      End
   End
   Begin VB.CheckBox chkPrivacy 
      Caption         =   "隐私项目"
      Height          =   210
      Left            =   6735
      TabIndex        =   32
      Top             =   150
      Width           =   1065
   End
   Begin VB.CommandButton cmdFormula 
      Caption         =   "∑"
      Height          =   285
      Left            =   7530
      TabIndex        =   25
      Top             =   3075
      Width           =   300
   End
   Begin VB.OptionButton OptApplyType 
      Caption         =   "应用于本类试管编码"
      Height          =   285
      Left            =   5835
      TabIndex        =   31
      Top             =   2762
      Width           =   1995
   End
   Begin VB.OptionButton OptApplyOnly 
      Caption         =   "应用于本项试管编码"
      Height          =   285
      Left            =   3780
      TabIndex        =   30
      Top             =   2762
      Value           =   -1  'True
      Width           =   1995
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   2835
      Left            =   -3540
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   240
      Visible         =   0   'False
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   5001
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.TextBox txt默认结果 
      Height          =   300
      Left            =   6330
      MaxLength       =   40
      TabIndex        =   23
      Top             =   1965
      Width           =   1380
   End
   Begin VB.ComboBox cbo结果范围 
      Height          =   300
      ItemData        =   "frmLabItemBase.frx":0004
      Left            =   6330
      List            =   "frmLabItemBase.frx":0006
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1593
      Width           =   1380
   End
   Begin VB.CommandButton cmd诊疗分类 
      Caption         =   "&P"
      Height          =   285
      Left            =   4875
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   113
      Width           =   285
   End
   Begin VB.ComboBox cbo项目类别 
      Height          =   300
      ItemData        =   "frmLabItemBase.frx":0008
      Left            =   6330
      List            =   "frmLabItemBase.frx":000A
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   849
      Width           =   1380
   End
   Begin VB.ComboBox cbo结果类型 
      Height          =   300
      ItemData        =   "frmLabItemBase.frx":000C
      Left            =   6330
      List            =   "frmLabItemBase.frx":000E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   1221
      Width           =   1380
   End
   Begin VB.ComboBox cbo操作类型 
      Height          =   300
      Left            =   3975
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   1200
   End
   Begin VB.CheckBox chk组合项目 
      Caption         =   "组合检验项目"
      Height          =   210
      Left            =   5310
      TabIndex        =   16
      Top             =   525
      Width           =   1440
   End
   Begin VB.ComboBox cbo适用性别 
      Height          =   300
      Left            =   3975
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2337
      Width           =   1200
   End
   Begin VB.TextBox txt名称五笔 
      Height          =   300
      Left            =   3255
      MaxLength       =   12
      TabIndex        =   7
      Top             =   1221
      Width           =   1320
   End
   Begin VB.CheckBox chk单独应用 
      Caption         =   "允许单独应用"
      Height          =   210
      Left            =   5295
      TabIndex        =   15
      Top             =   150
      Value           =   1  'Checked
      Width           =   1530
   End
   Begin VB.TextBox txt计算单位 
      Height          =   300
      Left            =   3975
      MaxLength       =   10
      TabIndex        =   11
      Top             =   1965
      Width           =   1200
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   7680
      Top             =   1650
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
            Picture         =   "frmLabItemBase.frx":0010
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabItemBase.frx":006E
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabItemBase.frx":00CC
            Key             =   "Dept"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "实验方法(&Y)"
      Height          =   180
      Left            =   165
      TabIndex        =   55
      Top             =   3870
      Width           =   990
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3375
      TabIndex        =   43
      Top             =   2709
      Width           =   300
   End
   Begin VB.Label Label2 
      Caption         =   "试管颜色"
      Height          =   225
      Left            =   2595
      TabIndex        =   42
      Top             =   2762
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "排列序号(&L)"
      Height          =   180
      Left            =   5295
      TabIndex        =   40
      Top             =   2385
      Width           =   990
   End
   Begin VB.Label lbl别名 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "别    名(&B)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   8
      Top             =   1649
      Width           =   990
   End
   Begin VB.Label lbl试管编码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "试管编码(&C)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   29
      Top             =   2762
      Width           =   990
   End
   Begin VB.Label lbl标本类型 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "标本类型(&M)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   12
      Top             =   2391
      Width           =   990
   End
   Begin VB.Label lbl默认结果 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "默认结果(&R)"
      Height          =   180
      Left            =   5310
      TabIndex        =   22
      Top             =   2025
      Width           =   990
   End
   Begin VB.Label lbl结果范围 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结果范围(W)"
      Height          =   180
      Left            =   5310
      TabIndex        =   27
      Top             =   1650
      Width           =   990
   End
   Begin VB.Label lbl诊疗分类 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "诊疗分类(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   0
      Top             =   165
      Width           =   990
   End
   Begin VB.Label lbl项目类别 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "项目性质(&K)"
      Height          =   180
      Left            =   5310
      TabIndex        =   17
      Top             =   900
      Width           =   990
   End
   Begin VB.Label lbl结果类型 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结果类型(T)"
      Height          =   180
      Left            =   5310
      TabIndex        =   19
      Top             =   1275
      Width           =   990
   End
   Begin VB.Label lbl计算公式 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "计算公式(F)"
      Height          =   180
      Left            =   165
      TabIndex        =   24
      Top             =   3133
      Width           =   990
   End
   Begin VB.Label lbl取值序列 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "取值序列(&P)"
      Height          =   180
      Left            =   165
      TabIndex        =   26
      Top             =   3510
      Width           =   990
   End
   Begin VB.Label lbl英文缩写 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "英文缩写(&E)"
      Height          =   180
      Left            =   165
      TabIndex        =   9
      Top             =   2020
      Width           =   990
   End
   Begin VB.Label lbl操作类型 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "检验类型(&T)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2985
      TabIndex        =   3
      Top             =   540
      Width           =   990
   End
   Begin VB.Label lbl项目编码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "项目编码(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   2
      Top             =   536
      Width           =   990
   End
   Begin VB.Label lbl项目名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "中文名称(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   5
      Top             =   907
      Width           =   990
   End
   Begin VB.Label lbl适用性别 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "适用性别(&X)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2940
      TabIndex        =   13
      Top             =   2391
      Width           =   990
   End
   Begin VB.Label lbl名称简码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "名称简码(&S)               (拼音)                 (五笔)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   6
      Top             =   1278
      Width           =   4950
   End
   Begin VB.Label lbl计算单位 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "计算单位(&U)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2940
      TabIndex        =   10
      Top             =   2020
      Width           =   990
   End
End
Attribute VB_Name = "frmLabItemBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngItemID As Long          '当前显示的项目id

Dim objNode As Node
Dim strTemp As String, aryTemp() As String
Dim lngCount As Long

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Public Function zlRefresh(lngItemId As Long) As Boolean
    '功能：根据项目id刷新当前显示内容
    Dim rsTemp As New ADODB.Recordset, rsGS As New ADODB.Recordset
    Dim strTmp As String, strItem As String, lngLength As Long
    mlngItemID = lngItemId
    
    '清除此前项目的显示
    Me.cbo操作类型.ListIndex = -1
    Me.txt项目编码.Text = "": Me.txt项目名称.Text = ""
    Me.txt名称拼音.Text = "": Me.txt名称五笔.Text = ""
    Me.txt英文缩写.Text = "": Me.txt计算单位.Text = ""
    Me.cbo标本类型.ListIndex = -1: Me.cbo适用性别.ListIndex = -1
    Me.cbo项目类别.ListIndex = -1: Me.cbo结果类型.ListIndex = -1: Me.txt默认结果.Text = ""
    Me.txt计算公式.Text = "": Me.txt取值序列.Text = "": Me.txtAlias.Text = ""
    Me.OptApplyOnly.Value = True
    Me.chkPrivacy.Value = 0
    Me.txt阳性公式.Text = ""
    Me.txt弱阳性公式.Text = ""
    Me.txtCutOff公式.Text = ""
    Me.txt检验方法.Text = ""
    Me.chk多参考.Value = 0
    
    If lngItemId = 0 Then zlRefresh = True: Exit Function
    
    '获取指定项目的信息
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select 分类id, 操作类型, 编码, 名称, 计算单位, 标本部位, 适用性别, 单独应用, 组合项目,试管编码 From 诊疗项目目录 Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    With rsTemp
        If .RecordCount > 0 Then
            If Val("" & !分类id) > 0 Then
                Me.tvwClass.Nodes("_" & !分类id).Selected = True
                Me.txt诊疗分类.Text = Me.tvwClass.SelectedItem.Text
                Me.txt诊疗分类.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
            End If
            For lngCount = 0 To Me.cbo操作类型.ListCount - 1
                If Mid(Me.cbo操作类型.List(lngCount), 4) = "" & !操作类型 Then Me.cbo操作类型.ListIndex = lngCount: Exit For
            Next
            Me.txt项目编码.Text = "" & !编码
            Me.txt项目名称.Text = "" & !名称
            Me.txt计算单位.Text = "" & !计算单位

            If "" & !试管编码 = "" Or "" & !试管编码 = "NULL" Then
                Me.cbo试管.ListIndex = 0
            Else
                For lngCount = 0 To Me.cbo试管.ListCount - 1
                    If Split(Me.cbo试管.List(lngCount), "-")(0) = "" & !试管编码 Then Me.cbo试管.ListIndex = lngCount: Exit For
                Next
            End If
            
            For lngCount = 0 To Me.cbo标本类型.ListCount - 1
                If Mid(Me.cbo标本类型.List(lngCount), 4) = "" & !标本部位 Then Me.cbo标本类型.ListIndex = lngCount: Exit For
            Next
            For lngCount = 0 To Me.cbo适用性别.ListCount - 1
                If Left(Me.cbo适用性别.List(lngCount), 1) = "" & !适用性别 Then Me.cbo适用性别.ListIndex = lngCount: Exit For
            Next
            Me.chk单独应用.Value = IIf(Val("" & !单独应用) = 1, 1, 0)
            Me.chk组合项目.Value = IIf(Val("" & !组合项目) = 1, 1, 0)
            If Me.chk组合项目.Value = 1 Then
                Me.chk单独应用.Enabled = False
            Else
                Me.chk单独应用.Enabled = True
            End If
        End If
    End With
        
    gstrSql = "Select 名称,性质,简码,码类 From 诊疗项目别名 Where 诊疗项目ID=[1] And 性质=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    With rsTemp
        Do While Not .EOF
            If !性质 = 1 And !码类 = 1 Then Me.txt名称拼音.Text = !简码
            If !性质 = 1 And !码类 = 2 Then Me.txt名称五笔.Text = !简码
            .MoveNext
        Loop
    End With
            
    gstrSql = "select 名称,性质,简码,码类 from 诊疗项目别名 where 诊疗项目ID=" & lngItemId
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    Do While Not rsTemp.EOF
        If rsTemp!性质 = 9 And rsTemp!码类 = 1 Then Me.txtAlias.Text = rsTemp!名称
        If rsTemp!性质 = 9 And rsTemp!码类 = 2 Then Me.txtAlias.Text = rsTemp!名称
        rsTemp.MoveNext
    Loop
            
    '查询基础检验项目对应的检验指标
    If Me.chk组合项目.Value = 0 Then
        gstrSql = "Select A.缩写, A.项目类别, A.结果类型, A.结果范围, A.默认值, A.计算公式, A.取值序列,A.隐私项目, " & vbNewLine & _
                  " A.阳性公式, A.弱阳性公式, A.CutOff公式,A.排列序号, A.检验方法, A.多参考 " & vbNewLine & _
                "From 检验项目 A, 检验报告项目 C" & vbNewLine & _
                "Where A.诊治项目id = C.报告项目id And C.诊疗项目id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
        With rsTemp
            If .RecordCount > 0 Then
                Me.txt英文缩写.Text = "" & !缩写
                For lngCount = 0 To Me.cbo项目类别.ListCount - 1
                    If Left(Me.cbo项目类别.List(lngCount), 1) = "" & !项目类别 Then Me.cbo项目类别.ListIndex = lngCount: Exit For
                Next
                For lngCount = 0 To Me.cbo结果类型.ListCount - 1
                    If Left(Me.cbo结果类型.List(lngCount), 1) = "" & !结果类型 Then Me.cbo结果类型.ListIndex = lngCount: Exit For
                Next
                Me.txt默认结果.Text = "" & !默认值
                Me.txt计算公式.Text = "" & !计算公式
                Me.chkPrivacy.Value = Nvl(!隐私项目, 0)
                Me.txt排列序号.Text = "" & !排列序号
                
                If Me.txt计算公式.Text <> "" Then
                    Do While Me.txt计算公式.Text Like "*[[]*[]]*"
                        strTmp = strTmp & Mid(Me.txt计算公式.Text, 1, InStr(Me.txt计算公式.Text, "[") - 1)
                        lngLength = InStr(Me.txt计算公式.Text, "]") - InStr(Me.txt计算公式.Text, "[") - 1
                        strItem = Mid(Me.txt计算公式.Text, InStr(Me.txt计算公式.Text, "[") + 1, lngLength)
                        gstrSql = "Select 诊治项目ID,缩写 From 检验项目 Where 诊治项目ID=[1] "
                        Set rsGS = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(strItem))
                        Do Until rsGS.EOF
                            If Trim("" & rsGS.Fields("缩写")) <> "" Then
                                strTmp = strTmp & "[" & Trim("" & rsGS.Fields("缩写")) & "]"
                            Else
                                strTmp = strTmp & "[" & Val(strItem) & "]"
                            End If
                            rsGS.MoveNext
                        Loop
                        Me.txt计算公式.Text = Mid(Me.txt计算公式.Text, InStr(Me.txt计算公式.Text, "]") + 1)
                    Loop
                    strTmp = strTmp & Mid(Me.txt计算公式.Text, InStr(Me.txt计算公式.Text, "]") + 1)
                    Me.txt计算公式.Text = strTmp
                End If
                
                Me.txt取值序列.Text = "" & !取值序列
                
                Me.txt阳性公式.Text = "" & !阳性公式
                Me.txt弱阳性公式.Text = "" & !弱阳性公式
                Me.txtCutOff公式.Text = "" & !CutOff公式
                Me.txt检验方法.Text = "" & !检验方法
                
                Me.chk多参考.Value = Val("" & !多参考)
            End If
        End With
    End If
    
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngItemId As Long) As Boolean
    '功能：开始项目编辑
    '参数： blnAdd-是否增加，否则为修改
    '       lngItemId-增加的参照项目，或者指定编辑的项目
    Dim rsTemp As New ADODB.Recordset
    If Me.tvwClass.Nodes.Count = 0 Then
        MsgBox "请先在字典中初始化“诊疗分类目录”！", vbInformation, gstrSysName
        zlEditStart = False: Exit Function
    End If
    If Me.cbo操作类型.ListCount = 0 Then
        MsgBox "请先在字典中初始化“检验类型”！", vbInformation, gstrSysName
        zlEditStart = False: Exit Function
    End If
    
    If blnAdd Then
        Err = 0: On Error GoTo ErrHand
        If Val(zlDatabase.GetPara(61, glngSys)) = 0 Then '诊疗项目编码递增模式
            gstrSql = "Select Nvl(Max(编码),'0000000') As 编码 From 诊疗项目目录 Where 类别 >= 'A'"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "zlEditStart")
            
            Me.txt项目编码.Text = IncStr(rsTemp!编码)
        Else
            strTemp = Mid(Me.txt诊疗分类.Text, 2, InStr(1, Me.txt诊疗分类.Text, "]") - 2)
            gstrSql = "Select Nvl(Max(编码),'0000000') As 编码" & _
                    " From 诊疗项目目录" & _
                    " Where 类别 >= 'A' And 编码 Like '" & strTemp & "%'"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "zlEditStart")
            
            Err = 0: On Error Resume Next
            If rsTemp!编码 = "0000000" Then
                Me.txt项目编码.Text = strTemp & IncStr(rsTemp!编码)
            Else
                Me.txt项目编码.Text = strTemp & IncStr(Right(rsTemp!编码, Len(rsTemp!编码) - Len(strTemp)))
            End If
        End If
        
        '清除并设置默认值
        Me.txt项目名称.Text = "": Me.cbo试管.ListIndex = 0
        Me.txt名称拼音.Text = "": Me.txt名称五笔.Text = ""
        Me.txt英文缩写.Text = "": Me.txt默认结果.Text = ""
        Me.txt计算公式.Text = "": Me.txt取值序列.Text = ""
        Me.txtAlias.Text = ""
        If lngItemId = 0 Then
            '说明以前没有可继承信息，需要设置部分默认值
            Me.cbo操作类型.ListIndex = 0
            Me.cbo标本类型.ListIndex = 0: Me.cbo适用性别.ListIndex = 0
            Me.cbo项目类别.ListIndex = 0: Me.cbo结果类型.ListIndex = 0
        End If
    End If

    mlngItemID = lngItemId
    Me.Enabled = True: Me.Tag = IIf(blnAdd, "增加", "修改")
    Me.BackColor = RGB(250, 250, 250): Me.chk单独应用.BackColor = Me.BackColor: Me.chk组合项目.BackColor = Me.BackColor
    Me.chkPrivacy.BackColor = Me.BackColor: Me.chk多参考.BackColor = Me.BackColor
    Me.OptApplyOnly.BackColor = Me.BackColor: Me.OptApplyType.BackColor = Me.BackColor
    Me.txt诊疗分类.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Function IncStr(ByVal strVal As String, Optional intUpDown As Integer, Optional ByRef strErr As String) As String
'功能：对一个字符串自动加1。
'说明：每一位进位时,如果是数字,则按十进制处理,否则按26进制处理
'参数：strVal=要加1的字符串
'      intUpDown = 0 加1 =1 减1
    Dim strValuse As String
    Dim intAdd As Integer
    Dim intUp As Integer
    Dim strValue As String
    Dim strValueOne As String
    Dim strHead As String
    
    Dim i  As Integer
    
    On Error GoTo errH
    
    strVal = UCase(strVal)

    For i = Len(strVal) To 1 Step -1
        strValueOne = Mid(strVal, i, 1)
        If Asc(strValueOne) >= Asc("0") And Asc(strValueOne) <= Asc("9") Then
        Else
            '不是数字
            strHead = Mid$(strVal, 1, i)
            strVal = Mid$(strVal, i + 1)
            Exit For
        End If
    Next
    
    strVal = UCase(strVal)
    
    If intUpDown = 0 Then
        '加1
        For i = Len(strVal) To 1 Step -1
            If i = Len(strVal) Then
                intAdd = 1
            Else
                intAdd = 0
            End If
            strValueOne = Mid(strVal, i, 1)
    
            If IsNumeric(strValueOne) Then
                If Val(strValueOne) + intAdd + intUp < 10 Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
                    strValue = "0" & strValue
                    intUp = 1
                End If
            Else
                If Asc(strValueOne) + intAdd + intUp <= Asc("Z") Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
                    strValue = "A" & strValue
                    intUp = 1
                End If
            End If
        Next
        
        If intUp = 1 Then
            If IsNumeric(strValueOne) Then
                strValue = "1" & strValue
            Else
                strValue = "A" & strValue
            End If
        End If
        IncStr = IIf(strHead <> "", strHead & strValue, strValue)
    Else
        For i = Len(strVal) To 1 Step -1
            If i = Len(strVal) Then
                intAdd = -1
            Else
                intAdd = 0
            End If
            strValueOne = Mid(strVal, i, 1)
    
            If IsNumeric(strValueOne) Then
                If Val(strValueOne) + intAdd + intUp >= 0 Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
                    strValue = "9" & strValue
                    intUp = -1
                End If
            Else
                If Asc(strValueOne) + intAdd + intUp >= Asc("A") Then
                    strValue = Chr(Asc(strValueOne) + intAdd + intUp) & strValue
                    intUp = 0
                Else
                    If intAdd = 0 Then
                        strValue = "Z" & strValue
                    End If
                    intUp = -1
                End If
            End If
        Next
        
        If intUp = 1 Then
            strValue = -1
        End If
        
        If Mid(strValue, 1, 1) = "0" Or Mid(strValue, 1, 1) = "A" Then
            strValue = Mid(strValue, 2)
            If strValue = "" Then strValue = 1
        End If
        IncStr = IIf(strHead <> "", strHead & strValue, strValue)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Enabled = False: Me.Tag = ""
    Me.BackColor = Me.cmd诊疗分类.BackColor: Me.chk单独应用.BackColor = Me.BackColor: Me.chk组合项目.BackColor = Me.BackColor:
    Me.chkPrivacy.BackColor = Me.BackColor: Me.chk多参考.BackColor = Me.BackColor
    Me.OptApplyOnly.BackColor = Me.cmd诊疗分类.BackColor: Me.OptApplyType.BackColor = Me.cmd诊疗分类.BackColor
    Call Me.zlRefresh(mlngItemID)
End Sub

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
    Dim lngNewId As Long
    Dim rsGS As New ADODB.Recordset, strItem As String, strTmp As String, lngLength As Long
    Dim str试管 As String
    
    '一般特性检查
    Err = 0: On Error GoTo ErrHand
    If Trim(Me.txt项目编码.Text) = "" Then
        MsgBox "请输入项目编码！", vbInformation, gstrSysName
        Me.txt项目编码.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt项目编码.Text), vbFromUnicode)) > Me.txt项目编码.MaxLength Then
        MsgBox "项目编码的长度超长（最多" & Me.txt项目编码.MaxLength & " 个字符）！", vbInformation, gstrSysName
        Me.txt项目编码.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt项目名称.Text) = "" Then
        MsgBox "请输入项目名称！", vbInformation, gstrSysName
        Me.txt项目名称.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt项目名称.Text), vbFromUnicode)) > Me.txt项目名称.MaxLength Then
        MsgBox "项目名称超长（最多" & Me.txt项目名称.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt项目名称.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt计算单位.Text), vbFromUnicode)) > Me.txt计算单位.MaxLength Then
        MsgBox "计算单位超长（最多" & Me.txt计算单位.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt计算单位.SetFocus: zlEditSave = 0: Exit Function
    End If
    '10804 并发操作时，检查检验类型是否被删除
    If Not zlExistItem("诊疗检验类型", "名称", Mid(Me.cbo操作类型.Text, InStr(1, Me.cbo操作类型, "-") + 1), "检验类型：" & Mid(Me.cbo操作类型.Text, InStr(1, Me.cbo操作类型, "-") + 1)) Then
        Me.cbo操作类型.SetFocus: zlEditSave = 0: Exit Function
    End If
'    If Trim(Me.txt英文缩写.Text) = "" And Me.cbo项目类别.ListIndex = 0 Then
'        MsgBox "基础项目请输入英文缩写！", vbInformation, gstrSysName
'        Me.txt英文缩写.SetFocus: zlEditSave = 0: Exit Function
'    End If
    If Trim(Me.txt名称拼音) = "" Then
        Me.txt名称拼音.Text = zlStr.GetCodeByORCL(Me.txt项目名称.Text, False, Me.txt名称拼音.MaxLength)
    End If
    
    If Trim(Me.txt名称五笔) = "" Then
        Me.txt名称五笔.Text = zlStr.GetCodeByORCL(Me.txt项目名称.Text, True, Me.txt名称五笔.MaxLength)
    End If
    If LenB(StrConv(Trim(Me.txt英文缩写.Text), vbFromUnicode)) > Me.txt英文缩写.MaxLength Then
        MsgBox "英文缩写超长（最多" & Me.txt英文缩写.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt英文缩写.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    '结果类型 为 非定量的，检查是否被其他计算项目引用
    If Mid(Me.cbo结果类型.Text, 1, 1) <> "1" And Me.Tag <> "增加" Then
        gstrSql = "Select 诊治项目id, 缩写, B.中文名, B.编码" & vbNewLine & _
                "From 诊治所见项目 B, 检验项目 A" & vbNewLine & _
                "Where A.诊治项目id = B.ID And" & vbNewLine & _
                "      计算公式 Like (Select '%' || Chr(91) || A.报告项目id || Chr(93) || '%' From 检验报告项目 A ,诊疗项目目录 B Where A.诊疗项目id=B.ID and B.组合项目=0 and A.诊疗项目id = [1])"
        Set rsGS = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID)
        strTmp = "此项目被以下项目引用，不能更改结果类型！"
        Do Until rsGS.EOF
            strItem = strItem & "(" & rsGS.Fields("编码") & ")" & rsGS.Fields("中文名") & vbNewLine
            rsGS.MoveNext
        Loop
        If strItem <> "" Then
            MsgBox strTmp & vbNewLine & strItem, vbInformation, Me.Caption
            Exit Function
        End If
    End If
    
    strItem = "": strTmp = ""
    If Me.txt计算公式.Text <> "" And Me.cbo项目类别.ListIndex = 2 Then
         
        Do While Me.txt计算公式.Text Like "*[[]*[]]*"
            strTmp = strTmp & Mid(Me.txt计算公式.Text, 1, InStr(Me.txt计算公式.Text, "[") - 1)
            lngLength = InStr(Me.txt计算公式.Text, "]") - InStr(Me.txt计算公式.Text, "[") - 1
            strItem = Mid(Me.txt计算公式.Text, InStr(Me.txt计算公式.Text, "[") + 1, lngLength)
            gstrSql = "Select 诊治项目ID,缩写 From 检验项目 Where (诊治项目ID=[1] or 缩写=[2]) "
            Set rsGS = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(strItem), strItem)
            Do Until rsGS.EOF
                strTmp = strTmp & "[" & Val("" & rsGS.Fields("诊治项目ID")) & "]"
                rsGS.MoveNext
            Loop
            Me.txt计算公式.Text = Mid(Me.txt计算公式.Text, InStr(Me.txt计算公式.Text, "]") + 1)
        Loop
        strTmp = strTmp & Mid(Me.txt计算公式.Text, InStr(Me.txt计算公式.Text, "]") + 1)
        Me.txt计算公式.Text = strTmp
    Else
        Me.txt计算公式.Text = ""
    End If
    
    If LenB(StrConv(Trim(Me.txt计算公式.Text), vbFromUnicode)) > Me.txt计算公式.MaxLength Then
        MsgBox "计算公式超长（最多" & Me.txt计算公式.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt计算公式.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    If LenB(StrConv(Trim(Me.txt默认结果.Text), vbFromUnicode)) > Me.txt默认结果.MaxLength Then
        MsgBox "默认结果超长（最多" & Me.txt默认结果.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt默认结果.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt取值序列.Text), vbFromUnicode)) > Me.txt取值序列.MaxLength Then
        MsgBox "取值序列超长（最多" & Me.txt取值序列.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt取值序列.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    If LenB(StrConv(Trim(Me.txt阳性公式.Text), vbFromUnicode)) > Me.txt阳性公式.MaxLength Then
        MsgBox "阳性公式超长（最多" & Me.txt阳性公式.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt阳性公式.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt弱阳性公式.Text), vbFromUnicode)) > Me.txt弱阳性公式.MaxLength Then
        MsgBox "弱阳性公式超长（最多" & Me.txt弱阳性公式.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt弱阳性公式.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txtCutOff公式.Text), vbFromUnicode)) > Me.txtCutOff公式.MaxLength Then
        MsgBox "CutOff公式超长（最多" & Me.txtCutOff公式.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txtCutOff公式.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Me.txt阳性公式.Text <> "" And txt阳性公式.Enabled Then
        strTmp = UCase(Me.txt阳性公式.Text)
        strTmp = Replace(strTmp, "OD", "100"): strTmp = Replace(strTmp, "NC", "1"): strTmp = Replace(strTmp, "BC", "1")
        strTmp = Replace(strTmp, "QC", "1"): strTmp = Replace(strTmp, "PC", "1")
        If Check_Expression(0, strTmp) = False Then
            MsgBox "阳性公式错误，请检查！", vbInformation, Me.Caption
            Me.txt阳性公式.SetFocus: zlEditSave = 0: Exit Function
        End If
        Me.txt阳性公式.Text = UCase(Me.txt阳性公式.Text)
    End If
    If Me.txt弱阳性公式.Text <> "" And txt弱阳性公式.Enabled Then
        strTmp = UCase(Me.txt弱阳性公式.Text)
        strTmp = Replace(strTmp, "OD", "100"): strTmp = Replace(strTmp, "NC", "1"): strTmp = Replace(strTmp, "BC", "1")
        strTmp = Replace(strTmp, "QC", "1"): strTmp = Replace(strTmp, "PC", "1")
        If Check_Expression(0, strTmp) = False Then
            MsgBox "弱阳性公式错误，请检查！", vbInformation, Me.Caption
            Me.txt弱阳性公式.SetFocus: zlEditSave = 0: Exit Function
        End If
        Me.txt弱阳性公式.Text = UCase(Me.txt弱阳性公式.Text)
    End If
    If Me.txtCutOff公式.Text <> "" And txtCutOff公式.Enabled Then
        strTmp = UCase(Me.txtCutOff公式.Text)
        strTmp = Replace(strTmp, "OD", "100"): strTmp = Replace(strTmp, "NC", "1"): strTmp = Replace(strTmp, "BC", "1")
        strTmp = Replace(strTmp, "QC", "1"): strTmp = Replace(strTmp, "PC", "1")
        If Check_Expression(1, strTmp) = False Then
            MsgBox "CutOff公式错误，请检查！", vbInformation, Me.Caption
            Me.txtCutOff公式.SetFocus: zlEditSave = 0: Exit Function
        End If
        Me.txtCutOff公式.Text = UCase(Me.txtCutOff公式.Text)
    End If
    '数据保存语句组织
    If Me.Tag = "增加" Then
        lngNewId = zlDatabase.GetNextId("诊疗项目目录")
        If zlClinicCodeRepeat(Trim(Me.txt项目编码.Text)) = True Then zlEditSave = 0: Exit Function
    Else
        If zlClinicCodeRepeat(Trim(Me.txt项目编码.Text), mlngItemID) = True Then zlEditSave = 0: Exit Function
        '检查项目是否还在
        If zlExistItem("诊疗项目目录", "ID", mlngItemID, Trim(Me.txt项目名称.Text)) = False Then zlEditSave = 0: Exit Function
       
    End If

    gstrSql = Me.txt诊疗分类.Tag & ",'" & Mid(Me.cbo操作类型.Text, InStr(1, Me.cbo操作类型.Text, "-") + 1) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt项目编码.Text) & "','" & Trim(Me.txt项目名称.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt名称拼音.Text) & "','" & Trim(Me.txt名称五笔.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txtAlias.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt英文缩写.Text) & "','" & Trim(Me.txt计算单位.Text) & "'"
    If InStr(Me.cbo标本类型.Text, "-") > 0 Then
        gstrSql = gstrSql & ",'" & Mid(Me.cbo标本类型.Text, InStr(Me.cbo标本类型.Text, "-") + 1) & "'," & Me.cbo适用性别.ListIndex
    Else
        gstrSql = gstrSql & ",'" & Mid(Me.cbo标本类型.Text, 4) & "'," & Me.cbo适用性别.ListIndex
    End If
    gstrSql = gstrSql & "," & Me.chk单独应用.Value & "," & Me.chk组合项目.Value
    
    '增加排列序号
    gstrSql = gstrSql & "," & IIf(Trim(Me.txt排列序号.Text) = "", "Null", Val(Me.txt排列序号))
    '-- 2008-12-24 加入 检验方法
    gstrSql = gstrSql & ",'" & Trim(Me.txt检验方法) & "'"
    
    If Me.chk组合项目.Value = 0 Then
        If Me.cbo项目类别.ListIndex = -1 Then
            MsgBox "非组合项目必须说明项目性质！", vbInformation, gstrSysName
            Me.cbo项目类别.SetFocus: zlEditSave = 0: Exit Function
        End If
        gstrSql = gstrSql & "," & Left(Me.cbo项目类别.Text, 1)
        
        If Me.cbo结果类型.ListIndex = -1 Then
            MsgBox "非组合项目必须说明结果类型！", vbInformation, gstrSysName
            Me.cbo结果类型.SetFocus: zlEditSave = 0: Exit Function
        End If
        gstrSql = gstrSql & "," & Left(Me.cbo结果类型.Text, 1)
        
        gstrSql = gstrSql & ",'" & Me.cbo结果范围.Text & "','" & Trim(Me.txt默认结果.Text) & "'"
        gstrSql = gstrSql & ",'" & Trim(Me.txt计算公式.Text) & "','" & Trim(Me.txt取值序列.Text) & "'"
        gstrSql = gstrSql & "," & Me.chkPrivacy.Value
        gstrSql = gstrSql & "," & Me.chk多参考.Value
        
        '-begin 20080318 加入酶标项目的内容
        If txt阳性公式.Enabled Then gstrSql = gstrSql & ",'" & Trim(Me.txt阳性公式) & "'"
        If txt弱阳性公式.Enabled Then gstrSql = gstrSql & ",'" & Trim(Me.txt弱阳性公式) & "'"
        If txtCutOff公式.Enabled Then gstrSql = gstrSql & ",'" & Trim(Me.txtCutOff公式) & "'"
        '-- End 20080318 加入酶标项目的内容
        
        
    End If
    
    If Me.Tag = "增加" Then
        gstrSql = "Zl_检验项目_Edit(1," & lngNewId & "," & gstrSql & ")"
    Else
        gstrSql = "Zl_检验项目_Edit(2," & mlngItemID & "," & gstrSql & ")"
    End If
    
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
'    If Trim(Me.Txt试管编码) <> "" Then
        str试管 = ""
        If Me.cbo试管.ListIndex > 0 Then
            str试管 = Me.cbo试管.List(Me.cbo试管.ListIndex)
            str试管 = Split(str试管, "-")(0)
        End If
        gstrSql = "Zl_诊疗项目目录_Batch_Update('" & IIf(Trim(str试管) = "", "NULL", str试管) & "'," & IIf(Me.Tag = "增加", lngNewId, mlngItemID) & _
        ",'" & IIf(Me.OptApplyOnly.Value = True, "", Mid(Me.cbo操作类型.Text, InStr(1, Me.cbo操作类型, "-") + 1)) & "')"
        zlDatabase.ExecuteProcedure gstrSql, gstrSysName
'    End If
    
    If Me.Tag = "增加" Then mlngItemID = lngNewId
    Me.Enabled = False: Me.Tag = ""
    Me.BackColor = Me.cmd诊疗分类.BackColor: Me.chk单独应用.BackColor = Me.BackColor: Me.chk组合项目.BackColor = Me.BackColor
    Me.OptApplyOnly.BackColor = Me.cmd诊疗分类.BackColor: Me.OptApplyType.BackColor = Me.cmd诊疗分类.BackColor
    zlEditSave = mlngItemID: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

Private Function Check_Expression(ByVal intType As Integer, strExpression As String) As Boolean
    '验证酶标表达示是否正确
    'inttype =0 逻辑表达式，1＝合法的计算表达式
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    If intType = 0 Then
        strSql = "Select 1 From Dual Where " & strExpression
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        Check_Expression = True
    ElseIf intType = 1 Then
        strSql = "Select " & strExpression & " From Dual "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        Check_Expression = True
    End If
    Exit Function
errHandle:
    Check_Expression = False
End Function
'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------

Private Sub cbo标本类型_Click()
'   根据标本限制性别
    If Me.cbo标本类型.ListIndex >= 0 Then
        If Me.cbo标本类型.ItemData(Me.cbo标本类型.ListIndex) = 1 Then
            Me.cbo适用性别.ListIndex = 1
            Me.cbo适用性别.Enabled = False
        ElseIf cbo标本类型.ItemData(Me.cbo标本类型.ListIndex) = 2 Then
            Me.cbo适用性别.ListIndex = 2
            Me.cbo适用性别.Enabled = False
        Else
            Me.cbo适用性别.Enabled = True
        End If
    End If
End Sub

Private Sub cbo标本类型_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo标本类型_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo操作类型_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo操作类型_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo结果范围_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo结果范围_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo结果类型_Click()
    Select Case Left(Me.cbo结果类型, 1)
    Case 1: Me.txt取值序列.Text = "": Me.txt取值序列.Enabled = False
    Case 2
        Me.txt取值序列.Enabled = True
        If Me.txt取值序列.Text = "-;±;+;++;+++;++++" Then Me.txt取值序列 = ""
    Case 3
        Me.txt取值序列.Enabled = True
        If Trim(Me.txt取值序列.Text) = "" Then Me.txt取值序列.Text = "-;±;+;++;+++;++++"
    End Select
End Sub

Private Sub cbo结果类型_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo结果类型_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo试管_Click()
    Dim lngColor As Long
    
    If cbo试管.ListIndex > 0 Then
        
        lngColor = cbo试管.ItemData(cbo试管.ListIndex)
        If lngColor < 0 Then lngColor = 0
        On Error Resume Next
        lblColor.BackColor = lngColor
    Else
        lblColor.BackColor = Label2.BackColor
    End If
End Sub

Private Sub cbo试管_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo试管_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo适用性别_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo适用性别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo项目类别_Click()
    '微生物项目不能设置组合项目
    If Me.cbo项目类别.ListIndex = 1 Then
        Me.chk组合项目.Value = 0: Me.chk组合项目.Enabled = False
    Else
        Me.chk组合项目.Enabled = True
    End If
    '非计算项目不能设置计算公式
    If Me.cbo项目类别.ListIndex = 2 Then
        Me.txt计算公式.Text = Me.txt计算公式.Tag: Me.txt计算公式.Enabled = True
        Me.cmdFormula.Enabled = True
    Else
        Me.txt计算公式.Tag = Me.txt计算公式.Text: Me.txt计算公式.Enabled = False
        Me.cmdFormula.Enabled = False
    End If
    
    '非酶标项目不能设置酶标公式
    If Me.cbo项目类别.ListIndex = 3 Then
        Me.txt阳性公式.Text = Me.txt阳性公式.Tag: Me.txt阳性公式.Enabled = True
        Me.txt弱阳性公式.Text = Me.txt弱阳性公式.Tag: Me.txt弱阳性公式.Enabled = True
        Me.txtCutOff公式.Text = Me.txtCutOff公式.Tag: Me.txtCutOff公式.Enabled = True
    Else
        Me.txt阳性公式.Tag = Me.txt阳性公式.Text: Me.txt阳性公式.Enabled = False
        Me.txt弱阳性公式.Tag = Me.txt弱阳性公式.Text: Me.txt弱阳性公式.Enabled = False
        Me.txtCutOff公式.Tag = Me.txtCutOff公式.Text: Me.txtCutOff公式.Enabled = False
    End If
End Sub

Private Sub cbo项目类别_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo项目类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk单独应用_Click()
    If Me.chk组合项目.Value = 0 Then
        Me.chkPrivacy.Visible = True
        Me.chk多参考.Visible = True
    Else
        Me.chkPrivacy.Value = 0
        Me.chkPrivacy.Visible = False
        Me.chk多参考.Value = 0
        Me.chk多参考.Visible = False
    End If
End Sub

Private Sub chk单独应用_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chk单独应用_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk组合项目_Click()
    If Me.chk组合项目.Value = 0 Then
        Me.txt英文缩写.Text = Me.txt英文缩写.Tag: Me.txt英文缩写.Enabled = True
        Me.cbo项目类别.ListIndex = Val(Me.cbo项目类别.Tag): Me.cbo项目类别.Enabled = True
        Me.cbo结果类型.ListIndex = Val(Me.cbo结果类型.Tag): Me.cbo结果类型.Enabled = True
        Me.cbo结果范围.ListIndex = Val(Me.cbo结果范围.Tag): Me.cbo结果范围.Enabled = True
        Me.txt默认结果.Text = Me.txt默认结果.Tag: Me.txt默认结果.Enabled = True
        
        Me.txt计算公式.Enabled = (Me.cbo项目类别.ListIndex = 2)
        If Me.txt计算公式.Enabled Then Me.txt计算公式.Text = Me.txt计算公式.Tag
        
        Me.txt取值序列.Tag = Me.txt取值序列.Text: Me.txt取值序列.Enabled = True
        chk单独应用.Enabled = True
        Me.chkPrivacy.Visible = True
        
        Me.txt阳性公式.Enabled = (Me.cbo项目类别.ListIndex = 3)
        Me.txt弱阳性公式.Enabled = (Me.cbo项目类别.ListIndex = 3)
        Me.txtCutOff公式.Enabled = (Me.cbo项目类别.ListIndex = 3)
        If Me.txt阳性公式.Enabled Then Me.txt阳性公式.Text = Me.txt阳性公式.Tag
        If Me.txt弱阳性公式.Enabled Then Me.txt弱阳性公式.Text = Me.txt弱阳性公式.Tag
        If Me.txtCutOff公式.Enabled Then Me.txtCutOff公式.Text = Me.txtCutOff公式.Tag
        
        Me.chk多参考.Visible = True
    Else
        Me.txt英文缩写.Tag = Me.txt英文缩写.Text: Me.txt英文缩写.Text = "": Me.txt英文缩写.Enabled = False
        Me.cbo项目类别.Tag = Me.cbo项目类别.ListIndex: Me.cbo项目类别.ListIndex = -1: Me.cbo项目类别.Enabled = False
        Me.cbo结果类型.Tag = Me.cbo结果类型.ListIndex: Me.cbo结果类型.ListIndex = -1: Me.cbo结果类型.Enabled = False
        Me.cbo结果范围.Tag = Me.cbo结果范围.ListIndex: Me.cbo结果范围.ListIndex = -1: Me.cbo结果范围.Enabled = False
        Me.txt默认结果.Tag = Me.txt默认结果.Text: Me.txt默认结果.Text = "": Me.txt默认结果.Enabled = False
        Me.txt计算公式.Tag = Me.txt计算公式.Text: Me.txt计算公式.Text = "": Me.txt计算公式.Enabled = False
        Me.txt取值序列.Tag = Me.txt取值序列.Text: Me.txt取值序列.Text = "": Me.txt取值序列.Enabled = False
        
        Me.txt阳性公式.Tag = Me.txt阳性公式.Text: Me.txt阳性公式.Text = "": Me.txt阳性公式.Enabled = False
        Me.txt弱阳性公式.Tag = Me.txt弱阳性公式.Text: Me.txt弱阳性公式.Text = "": Me.txt弱阳性公式.Enabled = False
        Me.txtCutOff公式.Tag = Me.txtCutOff公式.Text: Me.txtCutOff公式.Text = "": Me.txtCutOff公式.Enabled = False
        
        chk单独应用.Enabled = False: chk单独应用.Value = 1: Me.chkPrivacy.Visible = False: Me.chkPrivacy.Value = 0: Me.chk多参考.Visible = False:  Me.chk多参考.Value = 0
    End If
End Sub

Private Sub chk组合项目_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chk组合项目_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdFormula_Click()
    txt计算公式 = FrmLabItemFormula.DefFormula(mlngItemID, txt计算公式, Me)
End Sub

Private Sub cmd诊疗分类_Click()
    With Me.tvwClass
        .Left = Me.txt诊疗分类.Left
        .Top = Me.txt诊疗分类.Top + Me.txt诊疗分类.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.tvwClass.Visible Then
        On Error Resume Next
        Me.tvwClass.Visible = False: Me.txt诊疗分类.SetFocus: Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
    
Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo ErrHand
    
    '字段长度限制
    gstrSql = "Select A.编码, A.名称, A.计算单位, B.简码 From 诊疗项目目录 A, 诊疗项目别名 B " & _
            " Where A.ID = B.诊疗项目id And A.ID = 0 And B.码类 = 1"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
    With rsTemp
        Me.txt项目编码.MaxLength = .Fields("编码").DefinedSize
        Me.txt项目名称.MaxLength = .Fields("名称").DefinedSize
        Me.txt计算单位.MaxLength = .Fields("计算单位").DefinedSize
        Me.txt名称拼音.MaxLength = .Fields("简码").DefinedSize
        Me.txt名称五笔.MaxLength = .Fields("简码").DefinedSize
    End With
    
    gstrSql = "Select A.缩写, A.默认值, A.计算公式, A.取值序列, A.阳性公式, A.弱阳性公式, A.CutOff公式 From 检验项目 A Where A.诊治项目ID = 0"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
    With rsTemp
        Me.txt英文缩写.MaxLength = .Fields("缩写").DefinedSize
        Me.txt英文缩写.MaxLength = .Fields("默认值").DefinedSize
        Me.txt计算公式.MaxLength = .Fields("计算公式").DefinedSize
        Me.txt取值序列.MaxLength = .Fields("取值序列").DefinedSize
        Me.txt阳性公式.MaxLength = .Fields("阳性公式").DefinedSize
        Me.txt弱阳性公式.MaxLength = .Fields("弱阳性公式").DefinedSize
        Me.txtCutOff公式.MaxLength = .Fields("CutOff公式").DefinedSize
    End With
    
    '诊疗分类装入
    gstrSql = "select ID,上级ID,编码,名称,简码" & _
            " From 诊疗分类目录" & _
            " Where 类型 = 5" & _
            " start with 上级ID is null" & _
            " connect by prior ID=上级ID"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
'        If .BOF Or .EOF Then MsgBox "请首先建立诊疗分类项目!", vbExclamation, gstrSysName: Exit Sub
    Me.tvwClass.Nodes.Clear
    With rsTemp
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            End If
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!简码), "", !简码)
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        If Me.tvwClass.Nodes.Count > 0 Then
            Me.tvwClass.Nodes(1).Selected = True
            Me.txt诊疗分类.Text = Me.tvwClass.SelectedItem.Text
            Me.txt诊疗分类.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        End If
    End With
    
    '检验操作类型
    gstrSql = "Select 编码,名称 From 诊疗检验类型"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
    Me.cbo操作类型.Clear
    
    With rsTemp
        Do While Not .EOF
            Me.cbo操作类型.AddItem !编码 & "-" & !名称
            .MoveNext
        Loop
        If Me.cbo操作类型.ListCount > 0 Then Me.cbo操作类型.ListIndex = 0
    
        '性别数据要先于检验标本类型装入，因检验标本类型要用到这个数据
        aryTemp = Split("0-无性别区分;1-男性;2-女性", ";")
        For lngCount = LBound(aryTemp) To UBound(aryTemp)
            Me.cbo适用性别.AddItem aryTemp(lngCount)
        Next
        Me.cbo适用性别.ListIndex = 0
    End With
        '检验操作类型
        
        gstrSql = "Select 编码,名称,适用性别 From 诊疗检验标本"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
        Me.cbo标本类型.Clear
    With rsTemp
        Do While Not .EOF
            Me.cbo标本类型.AddItem !编码 & "-" & !名称
            If InStr(Trim("" & !适用性别), "男") > 0 Then
                Me.cbo标本类型.ItemData(Me.cbo标本类型.NewIndex) = 1
            ElseIf InStr(Trim("" & !适用性别), "女") > 0 Then
                Me.cbo标本类型.ItemData(Me.cbo标本类型.NewIndex) = 2
            Else
                Me.cbo标本类型.ItemData(Me.cbo标本类型.NewIndex) = 0
            End If
            
            .MoveNext
        Loop
        If Me.cbo标本类型.ListCount > 0 Then Me.cbo标本类型.ListIndex = 0
    End With
    
    '检验结果范围
    gstrSql = "Select Distinct 分类  From 检验结果描述"
'    If .State = adStateOpen Then .Close
'    Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'    Call SQLTest
    Me.cbo结果范围.Clear
    Me.cbo结果范围.AddItem ""
    With rsTemp
        Do While Not .EOF
            Me.cbo结果范围.AddItem "" & !分类
            .MoveNext
        Loop
        If Me.cbo结果范围.ListCount > 0 Then Me.cbo结果范围.ListIndex = 0
    
    End With
    '其他固定内容装入
    aryTemp = Split("1-普通检验;2-微生物;3-计算项目;4-酶标项目", ";")
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo项目类别.AddItem aryTemp(lngCount)
    Next
    Me.cbo项目类别.ListIndex = 0
    
    aryTemp = Split("1-定量;2-定性;3-半定量", ";")
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo结果类型.AddItem aryTemp(lngCount)
    Next
    Me.cbo结果类型.ListIndex = 0
    
    '试管
    gstrSql = "Select 编码,名称,颜色 From 采血管类型"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.cbo试管.Clear
    Me.cbo试管.AddItem "<未设置>"
    Me.cbo试管.ItemData(cbo试管.NewIndex) = Me.BackColor
    Do Until rsTemp.EOF
        Me.cbo试管.AddItem rsTemp!编码 & "-" & rsTemp!名称
        Me.cbo试管.ItemData(cbo试管.NewIndex) = rsTemp!颜色
        rsTemp.MoveNext
    Loop
    If Me.cbo试管.ListCount > 0 Then Me.cbo试管.ListIndex = 0
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txt诊疗分类.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    Me.txt诊疗分类.Text = Me.tvwClass.SelectedItem.Text
    Me.txt诊疗分类.SetFocus
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeySpace
        Call tvwClass_DblClick
    Case vbKeyEscape
        Call tvwClass_LostFocus
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    If Me.cmd诊疗分类 Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub

Private Sub txtAlias_GotFocus()
    Me.txtAlias.SelStart = 0
    Me.txtAlias.SelLength = Len(Me.txtAlias.Text)
End Sub

Private Sub txtAlias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtCutOff公式_GotFocus()
    Me.txtCutOff公式.SelStart = 0: Me.txtCutOff公式.SelLength = 1000
End Sub

Private Sub txtCutOff公式_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt计算单位_GotFocus()
    Me.txt计算单位.SelStart = 0: Me.txt计算单位.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt计算单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
'    If InStr(" ~!@#$%^&*_|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt计算公式_GotFocus()
    Me.txt计算公式.SelStart = 0: Me.txt计算公式.SelLength = 1000
End Sub

Private Sub txt计算公式_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt检验方法_GotFocus()
    Me.txt检验方法.SelStart = 0: Me.txt检验方法.SelLength = Me.txt检验方法.MaxLength
End Sub

Private Sub txt检验方法_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt名称拼音_GotFocus()
    Me.txt名称拼音.SelStart = 0: Me.txt名称拼音.SelLength = 1000
End Sub

Private Sub txt名称拼音_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt名称五笔_GotFocus()
    Me.txt名称五笔.SelStart = 0: Me.txt名称五笔.SelLength = 1000
End Sub

Private Sub txt名称五笔_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt默认结果_GotFocus()
    Me.txt默认结果.SelStart = 0: Me.txt默认结果.SelLength = 1000
End Sub

Private Sub txt默认结果_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt排列序号_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt取值序列_GotFocus()
    Me.txt取值序列.SelStart = 0: Me.txt取值序列.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt取值序列_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt弱阳性公式_GotFocus()
    Me.txt弱阳性公式.SelStart = 0: Me.txt弱阳性公式.SelLength = 1000
End Sub

Private Sub txt弱阳性公式_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt项目编码_GotFocus()
    Me.txt项目编码.SelStart = 0: Me.txt项目编码.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt项目编码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt项目名称_GotFocus()
    Me.txt项目名称.SelStart = 0: Me.txt项目名称.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt项目名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt项目名称.Text = MoveSpecialChar(txt项目名称.Text)
        Me.txt名称拼音.Text = zlStr.GetCodeByORCL(Me.txt项目名称.Text, False, Me.txt名称拼音.MaxLength)
        Me.txt名称五笔.Text = zlStr.GetCodeByORCL(Me.txt项目名称.Text, True, Me.txt名称五笔.MaxLength)
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt项目名称_LostFocus()
    Me.txt名称拼音.Text = zlStr.GetCodeByORCL(Me.txt项目名称.Text, False, Me.txt名称拼音.MaxLength)
    Me.txt名称五笔.Text = zlStr.GetCodeByORCL(Me.txt项目名称.Text, True, Me.txt名称五笔.MaxLength)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt阳性公式_GotFocus()
    Me.txt阳性公式.SelStart = 0: Me.txt阳性公式.SelLength = 1000
End Sub

Private Sub txt阳性公式_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt英文缩写_GotFocus()
    Me.txt英文缩写.SelStart = 0: Me.txt英文缩写.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt英文缩写_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt诊疗分类_Change()
    Me.txt诊疗分类.SelStart = 0: Me.txt诊疗分类.SelLength = 1000
End Sub

Private Sub txt诊疗分类_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub


