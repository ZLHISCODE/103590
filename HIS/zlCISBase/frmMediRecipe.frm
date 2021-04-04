VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmMediRecipe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "中药配方编辑"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   Icon            =   "frmMediRecipe.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.OptionButton optDosageType 
      Caption         =   "忽略形态"
      Height          =   255
      Index           =   0
      Left            =   825
      TabIndex        =   43
      Top             =   2640
      Width           =   1095
   End
   Begin VB.OptionButton optDosageType 
      Caption         =   "免煎剂"
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   42
      Top             =   2640
      Width           =   855
   End
   Begin VB.OptionButton optDosageType 
      Caption         =   "饮片"
      Height          =   255
      Index           =   2
      Left            =   3755
      TabIndex        =   41
      Top             =   2640
      Width           =   735
   End
   Begin VB.OptionButton optDosageType 
      Caption         =   "散装"
      Height          =   255
      Index           =   1
      Left            =   2470
      TabIndex        =   40
      Top             =   2640
      Width           =   735
   End
   Begin VB.ComboBox cmbStationNo 
      Height          =   300
      Left            =   825
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   2160
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.CommandButton cmd参考 
      Caption         =   "…"
      Height          =   285
      Left            =   4140
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1800
      Width           =   285
   End
   Begin VB.TextBox txt参考 
      Height          =   300
      Left            =   825
      TabIndex        =   34
      Top             =   1785
      Width           =   3240
   End
   Begin VB.TextBox txt说明 
      Height          =   300
      Left            =   5820
      MaxLength       =   30
      TabIndex        =   17
      Top             =   1815
      Width           =   3600
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Index           =   1
      Left            =   825
      MaxLength       =   40
      TabIndex        =   12
      Top             =   1425
      Width           =   3600
   End
   Begin VB.TextBox txt拼音 
      Height          =   300
      Index           =   1
      Left            =   5820
      MaxLength       =   12
      TabIndex        =   14
      Top             =   1425
      Width           =   1350
   End
   Begin VB.TextBox txt五笔 
      Height          =   300
      Index           =   1
      Left            =   7800
      MaxLength       =   12
      TabIndex        =   15
      Top             =   1425
      Width           =   960
   End
   Begin VB.TextBox txt疗程 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   825
      MaxLength       =   50
      TabIndex        =   26
      Top             =   5910
      Width           =   1020
   End
   Begin VB.ComboBox cbo频率 
      Height          =   300
      Left            =   780
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   5520
      Width           =   2115
   End
   Begin VB.ComboBox cbo煎法 
      Height          =   300
      Left            =   4155
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   5520
      Width           =   2115
   End
   Begin VB.ComboBox cbo用法 
      Height          =   300
      Left            =   7530
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   5520
      Width           =   2115
   End
   Begin VB.Frame fraLine 
      Height          =   15
      Index           =   1
      Left            =   -45
      TabIndex        =   33
      Top             =   6360
      Width           =   10410
   End
   Begin VB.TextBox txt五笔 
      Height          =   300
      Index           =   0
      Left            =   7800
      MaxLength       =   12
      TabIndex        =   10
      Top             =   1050
      Width           =   960
   End
   Begin VB.TextBox txt拼音 
      Height          =   300
      Index           =   0
      Left            =   5820
      MaxLength       =   12
      TabIndex        =   9
      Top             =   1050
      Width           =   1350
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Index           =   0
      Left            =   825
      MaxLength       =   40
      TabIndex        =   7
      Top             =   1050
      Width           =   3600
   End
   Begin VB.TextBox txt编码 
      Height          =   300
      Left            =   825
      MaxLength       =   13
      TabIndex        =   2
      Top             =   675
      Width           =   3600
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2280
      Left            =   4080
      TabIndex        =   31
      Top             =   8400
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4022
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3570
      Left            =   120
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   8400
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   6297
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8880
      TabIndex        =   28
      Top             =   6525
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   135
      Picture         =   "frmMediRecipe.frx":058A
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6525
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7800
      TabIndex        =   27
      Top             =   6525
      Width           =   1100
   End
   Begin VB.TextBox txt分类 
      Height          =   300
      Left            =   5820
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   4
      Top             =   675
      Width           =   3255
   End
   Begin VB.CommandButton cmd分类 
      Caption         =   "&P"
      Height          =   285
      Left            =   9105
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   690
      Width           =   285
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Index           =   0
      Left            =   0
      TabIndex        =   32
      Top             =   480
      Width           =   10410
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   6960
      Top             =   8280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediRecipe.frx":06D4
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediRecipe.frx":0C6E
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediRecipe.frx":1208
            Key             =   "药品"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediRecipe.frx":17A2
            Key             =   "脚注"
         EndProperty
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit msfRecipe 
      Height          =   2415
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4260
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Label lblDosageType 
      AutoSize        =   -1  'True
      Caption         =   "形态(&B)"
      Height          =   180
      Left            =   150
      TabIndex        =   39
      Top             =   2640
      Width           =   630
   End
   Begin VB.Label lblStationNo 
      AutoSize        =   -1  'True
      Caption         =   "院区(&Z)"
      Height          =   180
      Left            =   150
      TabIndex        =   38
      Top             =   2220
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "参考(&F)"
      Height          =   180
      Left            =   150
      TabIndex        =   36
      Top             =   1845
      Width           =   630
   End
   Begin VB.Label lbl说明 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "说明(&Z)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5175
      TabIndex        =   16
      Top             =   1875
      Width           =   630
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   135
      Picture         =   "frmMediRecipe.frx":207C
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lbl名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "别名(&A)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   11
      Top             =   1485
      Width           =   630
   End
   Begin VB.Label lbl简码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "简码(&M)                (拼音)            (五笔)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   5160
      TabIndex        =   13
      Top             =   1485
      Width           =   4230
   End
   Begin VB.Label lbl疗程 
      AutoSize        =   -1  'True
      Caption         =   "疗程(&T)            (根据治疗需要，一般需要连续服用该配方的持续天数)。"
      Height          =   180
      Left            =   135
      TabIndex        =   25
      Top             =   5955
      Width           =   6210
   End
   Begin VB.Label lbl频率 
      AutoSize        =   -1  'True
      Caption         =   "频率(&L)"
      Height          =   180
      Left            =   135
      TabIndex        =   19
      Top             =   5580
      Width           =   630
   End
   Begin VB.Label lbl煎法 
      AutoSize        =   -1  'True
      Caption         =   "煎法(&J)"
      Height          =   180
      Left            =   3510
      TabIndex        =   21
      Top             =   5580
      Width           =   630
   End
   Begin VB.Label lbl用法 
      AutoSize        =   -1  'True
      Caption         =   "用法(&U)"
      Height          =   180
      Left            =   6885
      TabIndex        =   23
      Top             =   5580
      Width           =   630
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    按组方原则，参考权威方剂资料，将中草药组成常用的配方，以方便医生下达中药医嘱时，能迅速准确地完成处方。"
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   780
      TabIndex        =   0
      Top             =   105
      Width           =   9645
   End
   Begin VB.Label lbl简码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "简码(&S)                (拼音)            (五笔)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   5160
      TabIndex        =   8
      Top             =   1110
      Width           =   4230
   End
   Begin VB.Label lbl名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "名称(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   6
      Top             =   1110
      Width           =   630
   End
   Begin VB.Label lbl编码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "编码(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   1
      Top             =   735
      Width           =   630
   End
   Begin VB.Label lbl分类 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "分类(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5160
      TabIndex        =   3
      Top             =   735
      Width           =   630
   End
End
Attribute VB_Name = "frmMediRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、上级程序通过本窗体ShowMe函数，将父窗体、权限、编辑项目的分类ID、ID,编辑状态等信息传递进入本程序
'   2、编辑状态：由Me.tag存放，分别为"增加"、"修改"、"查阅"，由上级程序通过ShowMe传入
'---------------------------------------------------
Private lngClassId As Long       '被编辑的分类ID，上级程序通过ShowMe传递进入
Private lngItemId As Long        '被编辑的项目ID，修改、查阅时由上级程序通过ShowMe传递进入,增加时为0，
Private mbyt中药味数 As Byte    '中药配方每行中药味数

Private strInputed As String
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim strTemp As String, aryTemp() As String
Dim intCount As Integer, intFence As Integer
Private Const mstrChar As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789."
Private mblnDosage As Boolean '是否点击了配方
Private mstrMedi As String  '编辑之前

Dim mstrMatch As String, strRefer As String '参考名称
Private mblnOK As Boolean
Private mintOldShape As Integer '记录选中的形态 '0-忽略形态 1-散装 2-饮片 3-免煎剂
Private mblnLoad As Boolean  '窗体是否加载完 true-加载完 false-未加载完
Private mblnClickNo As Boolean  '点击了提示框中的“否”按钮 true-点击了
Private Enum 配方列表
    空白 = 0
    药名ID = 1
    规格ID = 2
    名称 = 3
    数量 = 4
    单位 = 5
    脚注 = 6
    列数 = 7
End Enum
Private Sub GetDefineSize()
    '功能：得到数据库的表字段的长度
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
    gstrSql = "Select A.编码,A.标本部位,B.名称,B.简码 From 诊疗项目目录 A, 诊疗项目别名 B Where A.ID=B.诊疗项目ID and A.ID=0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)
    
    txt编码.MaxLength = rsTmp.Fields("编码").DefinedSize
    txt名称(0).MaxLength = rsTmp.Fields("名称").DefinedSize
    txt名称(1).MaxLength = rsTmp.Fields("名称").DefinedSize
    txt拼音(0).MaxLength = rsTmp.Fields("简码").DefinedSize
    txt拼音(1).MaxLength = rsTmp.Fields("简码").DefinedSize
    txt五笔(0).MaxLength = rsTmp.Fields("简码").DefinedSize
    txt五笔(1).MaxLength = rsTmp.Fields("简码").DefinedSize
    txt说明.MaxLength = rsTmp.Fields("标本部位").DefinedSize

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Function ShowMe(ByVal frmParent As Object, ByVal byt状态 As Byte, ByVal lng分类id As Long, Optional ByVal lng项目id As Long) As Boolean
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '---------------------------------------------------
    Dim intDosageType As Integer
    
    Me.Tag = Switch(byt状态 = 0, "增加", byt状态 = 1, "修改", byt状态 = 2, "查阅")
    lngClassId = lng分类id: lngItemId = lng项目id
    
    '填写需要选择的数据
    Err = 0: On Error GoTo ErrHand
    
    If Me.Tag = "新增" Then
        intDosageType = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "配方形态", 0)
    Else
        gstrSql = "select distinct nvl(配方类型,0) 配方类型 from 诊疗项目组合 where 诊疗组合ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "showme", lngItemId)
        If rsTemp.RecordCount > 0 Then
            intDosageType = rsTemp!配方类型
        End If
    End If
    
    If intDosageType < 0 Or intDosageType > 3 Then
        intDosageType = 0
    End If
    optDosageType(intDosageType).Value = True
    
    gstrSql = "select ID,上级ID,编码,名称,简码" & _
            " From 诊疗分类目录" & _
            " Where 类型 = 4" & _
            " start with 上级ID is null" & _
            " connect by prior ID=上级ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")

     With rsTemp
        If .BOF Or .EOF Then MsgBox "请首先建立配方诊疗分类项目之后增加配方", vbExclamation, gstrSysName: Unload Me: Exit Function
        Me.tvwClass.Nodes.Clear
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
        Me.tvwClass.Nodes("_" & lng分类id).Selected = True
        Me.txt分类.Text = Me.tvwClass.SelectedItem.Text
        Me.txt分类.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        
        gstrSql = "select 编码||'-'||名称 as 名称 from 诊疗频率项目 where 适用范围=2 order by 编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")
        
        Me.cbo频率.Clear
        Do While Not rsTemp.EOF
            Me.cbo频率.AddItem rsTemp!名称
            rsTemp.MoveNext
        Loop
        If Me.cbo频率.ListCount = 0 Then
            Me.cbo频率.Enabled = False
        Else
            Me.cbo频率.ListIndex = 0
        End If
        
        gstrSql = "select ID,rownum||'-'||名称 as 名称 from 诊疗项目目录 where 类别='E' and 操作类型='3' order by 编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")

        Me.cbo煎法.Clear
        Me.cbo煎法.AddItem "": Me.cbo煎法.ItemData(Me.cbo煎法.NewIndex) = 0
        Do While Not rsTemp.EOF
            Me.cbo煎法.AddItem rsTemp!名称: Me.cbo煎法.ItemData(Me.cbo煎法.NewIndex) = rsTemp!ID
            rsTemp.MoveNext
        Loop
        If Me.cbo煎法.ListCount = 0 Then
            Me.cbo煎法.Enabled = False
        Else
            Me.cbo煎法.ListIndex = 0
        End If
        
        gstrSql = "select ID,rownum||'-'||名称 as 名称 from 诊疗项目目录 where 类别='E' and 操作类型='4' order by 编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")

        Me.cbo用法.Clear
        Do While Not rsTemp.EOF
            Me.cbo用法.AddItem rsTemp!名称: Me.cbo用法.ItemData(Me.cbo用法.NewIndex) = rsTemp!ID
            rsTemp.MoveNext
        Loop
        If Me.cbo用法.ListCount = 0 Then
            Me.cbo用法.Enabled = False
        Else
            Me.cbo用法.ListIndex = 0
        End If
    End With
    If Me.cbo用法.Enabled = False And Me.cbo煎法.Enabled = False Then
        Me.cbo频率.Enabled = False: Me.txt疗程.Enabled = False
    End If

    '显示窗体
    Me.Show 1, frmParent
    ShowMe = mblnOK
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbo煎法_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo频率_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo用法_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String
    Dim strSql As String
    Dim str站点 As String
    Dim str编码 As String
    Dim i As Integer
    Dim intDosage As Integer '配方类型
    Dim intGroup As Integer '第几组    一行三味药有三组 一行四位药有4组.....
    Dim strCheckID As String    '药名id+规格id
    
    If optDosageType(3).Value = True And cbo煎法.ItemData(cbo煎法.ListIndex) <> 0 Then
        MsgBox "免煎剂不用设置煎法！", vbInformation, gstrSysName
        cbo煎法.SetFocus
        Exit Sub
    End If
    
    For i = 0 To optDosageType.UBound
        If optDosageType(i).Value = True Then
            intDosage = i
            Exit For
        End If
    Next
    
    '重新检查名称，并去掉特殊字符
    strTmp = MoveSpecialChar(txt名称(0).Text)
    If txt名称(0).Text <> strTmp Then
        txt名称(0).Text = strTmp
        Me.txt拼音(0).Text = zlStr.GetCodeByORCL(Me.txt名称(0).Text, False)
        Me.txt五笔(0).Text = zlStr.GetCodeByORCL(Me.txt名称(0).Text, True)
    End If
    strTmp = MoveSpecialChar(txt名称(1).Text)
    If txt名称(1).Text <> strTmp Then
        txt名称(1).Text = strTmp
        Me.txt拼音(1).Text = zlStr.GetCodeByORCL(Me.txt名称(1).Text, False)
        Me.txt五笔(1).Text = zlStr.GetCodeByORCL(Me.txt名称(1).Text, True)
    End If
    
    '一般特性检查
    If Trim(Me.txt编码.Text) = "" Then MsgBox "请输入编码！", vbInformation, gstrSysName: Me.txt编码.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt编码.Text), vbFromUnicode)) > Me.txt编码.MaxLength Then MsgBox "编码的超长（最多" & Me.txt编码.MaxLength & "个字符）！", vbInformation, gstrSysName: Me.txt编码.SetFocus: Exit Sub
    If Trim(Me.txt名称(0).Text) = "" Then MsgBox "请输入名称！", vbInformation, gstrSysName: Me.txt名称(0).SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt名称(0).Text), vbFromUnicode)) > Me.txt名称(0).MaxLength Then
        MsgBox "名称超长（" & Me.txt名称(0).MaxLength & "个字符或" & Me.txt名称(0).MaxLength / 2 & "个汉字）！", vbInformation, gstrSysName: Me.txt名称(0).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt名称(1).Text), vbFromUnicode)) > Me.txt名称(1).MaxLength Then
        MsgBox "别名超长（" & Me.txt名称(1).MaxLength & "个字符或" & Me.txt名称(1).MaxLength / 2 & "个汉字）！", vbInformation, gstrSysName: Me.txt名称(1).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt拼音(0).Text), vbFromUnicode)) > Me.txt拼音(0).MaxLength Then
        MsgBox "名称拼音简码超长（" & Me.txt拼音(0).MaxLength & "个字符）！", vbInformation, gstrSysName: Me.txt拼音(0).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt拼音(1).Text), vbFromUnicode)) > Me.txt拼音(1).MaxLength Then
        MsgBox "别名拼音简码超长（" & Me.txt拼音(1).MaxLength & "个字符）！", vbInformation, gstrSysName: Me.txt拼音(1).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt五笔(0).Text), vbFromUnicode)) > Me.txt五笔(0).MaxLength Then
        MsgBox "名称五笔简码超长（" & Me.txt五笔(0).MaxLength & "个字符）！", vbInformation, gstrSysName: Me.txt五笔(0).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt五笔(1).Text), vbFromUnicode)) > Me.txt五笔(1).MaxLength Then
        MsgBox "别名五笔简码超长（" & Me.txt五笔(1).MaxLength & "个字符）！", vbInformation, gstrSysName: Me.txt五笔(1).SetFocus: Exit Sub
    End If
    If Val(Me.txt疗程.Text) > 100 Then MsgBox "系统不允许设置太长的疗程（为0表示不设置疗程）！", vbExclamation, gstrSysName: Me.txt疗程.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt说明.Text), vbFromUnicode)) > Me.txt说明.MaxLength Then
        MsgBox "说明超长（" & Me.txt说明.MaxLength & "个字符或" & Me.txt说明.MaxLength / 2 & "个汉字）！", vbInformation, gstrSysName: Me.txt说明.SetFocus: Exit Sub
    End If
    
    '新增项目时，保证不出现重复编码，如果有重复自动在原编码基础上加1，直到不重复
    str编码 = Trim(txt编码.Text)
    If Me.Tag = "增加" Then
        Do While True
            gstrSql = "select a.编码 from 诊疗项目目录 a,诊疗项目类别 b where a.编码=[1] and a.类别=b.编码"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "编码是否重复", str编码)
            If rsTemp.RecordCount <> 0 Then
                str编码 = zlCommFun.IncStr(str编码)
            Else
                Exit Do
            End If
        Loop
    End If
    
    Dim strMembers As String
    strTemp = "": strMembers = ""
        
    If mbyt中药味数 = 3 Then
        intGroup = 2
    ElseIf mbyt中药味数 = 4 Then
        intGroup = 3
    End If
    With Me.msfRecipe
        For intCount = 1 To .Rows - 1
            For intFence = 0 To intGroup
                If Val(.TextMatrix(intCount, 配方列表.列数 * intFence + 配方列表.药名ID)) <> 0 Then   'id不能为空
'                    If (Val(.TextMatrix(intCount, 配方列表.列数 * intFence + 配方列表.数量)) = 0) Then   '数量不能为0
'                        MsgBox "数量不能为0，请输入数量！", vbInformation, gstrSysName
'                        .SetFocus
'                        .Row = intCount
'                        .Col = 配方列表.列数 * intFence + 配方列表.数量
'                        Exit Sub
'                    End If
                    
                    If strCheckID <> "" Then
                        If InStr(1, ";" & strCheckID & ";", ";" & Val(.TextMatrix(intCount, 配方列表.列数 * intFence + 配方列表.药名ID)) & "+" & Val(.TextMatrix(intCount, 配方列表.列数 * intFence + 配方列表.规格ID))) > 0 Then
                            MsgBox "配方中“" & .TextMatrix(intCount, 配方列表.列数 * intFence + 配方列表.名称) & "”重复应用！", vbInformation, gstrSysName
                            .SetFocus
                            Exit Sub
                        End If
                    End If
                    
                    strCheckID = IIf(strCheckID = "", "", strCheckID & ";") & Val(.TextMatrix(intCount, 配方列表.列数 * intFence + 配方列表.药名ID)) & "+" & Val(.TextMatrix(intCount, 配方列表.列数 * intFence + 配方列表.规格ID))

                    strMembers = strMembers & "|" & Val(.TextMatrix(intCount, 配方列表.列数 * intFence + 配方列表.药名ID)) & _
                            "^" & IIf(Val(.TextMatrix(intCount, 配方列表.列数 * intFence + 配方列表.规格ID)) = 0, Null, Val(.TextMatrix(intCount, 配方列表.列数 * intFence + 配方列表.规格ID))) & _
                            "^" & Val(.TextMatrix(intCount, 配方列表.列数 * intFence + 配方列表.数量)) & _
                            "^" & Trim(.TextMatrix(intCount, 配方列表.列数 * intFence + 配方列表.脚注))
                End If
            Next
        Next
        If strCheckID = "" Then MsgBox "未定义配方组成！", vbInformation, gstrSysName: .SetFocus: Exit Sub
    End With
    strMembers = Mid(strMembers, 2)
    
    If cmbStationNo.Text = "" Then
        str站点 = "Null"
    Else
        str站点 = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
    End If
    
    '数据保存
    If Me.Tag = "增加" Then
        lngItemId = zlDatabase.GetNextId("诊疗项目目录")
'        If zlClinicCodeRepeat(Trim(Me.txt编码.Text)) = True Then Exit Sub
    Else
        If zlClinicCodeRepeat(str编码, lngItemId) = True Then Exit Sub
    End If
    gstrSql = lngItemId & "," & Me.txt分类.Tag & ",'" & str编码 & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt名称(0).Text) & "','" & Trim(Me.txt拼音(0).Text) & "','" & Trim(Me.txt五笔(0).Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt名称(1).Text) & "','" & Trim(Me.txt拼音(1).Text) & "','" & Trim(Me.txt五笔(1).Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt说明.Text) & "'"
    If Me.cbo频率.Enabled Then
        gstrSql = gstrSql & ",'" & Left(Me.cbo频率.Text, InStr(1, Me.cbo频率.Text, "-") - 1) & "'"
    Else
        gstrSql = gstrSql & ",null"
    End If
    If Me.cbo煎法.Enabled Then
        gstrSql = gstrSql & "," & Me.cbo煎法.ItemData(Me.cbo煎法.ListIndex)
    Else
        gstrSql = gstrSql & ",0"
    End If
    If Me.cbo用法.Enabled Then
        gstrSql = gstrSql & "," & Me.cbo用法.ItemData(Me.cbo用法.ListIndex)
    Else
        gstrSql = gstrSql & ",0"
    End If
    gstrSql = gstrSql & "," & Val(Me.txt疗程.Text)
    
    gstrSql = gstrSql & "," & IIf(Me.txt参考.Tag = "", "Null", Me.txt参考.Tag)
    
    gstrSql = "zl_中药配方_UPDATE(" & gstrSql & ",'" & strMembers & "'," & IIf(cmbStationNo.Visible = False Or cmbStationNo.Text = "", "Null", str站点) & "," & intDosage & ")"
    
    Err = 0: On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        
    '保存配方类型到注册表
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "配方形态", intDosage
    
    If Me.Tag = "增加" Then
        If GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\诊疗项目增加\", "连续", 0) = 1 Then
            lngItemId = 0
            Call Form_Activate
            Me.txt编码.SetFocus
            Exit Sub
        End If
    End If
    mblnOK = True
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd参考_Click()
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = SelectRefer
    If Not rsTmp Is Nothing Then
        Me.txt参考 = rsTmp("名称"): Me.txt参考.Tag = rsTmp("ID"): strRefer = Me.txt参考
    Else
        MsgBox "没有找到可参考的项目。", vbInformation, Me.Caption
    End If
End Sub

Private Function SelectRefer(Optional ByVal strName As String = "") As ADODB.Recordset
    Dim strSql As String, strSQLItem As String
    Dim rsTmp As New ADODB.Recordset, iAttr As Integer
    
    On Error GoTo ErrHand
    strSql = "Select 类型 From 诊疗分类目录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngClassId)
    
    If rsTmp.EOF Then
        iAttr = -1
    Else
        iAttr = rsTmp(0)
    End If
    If Len(strName) = 0 Then
        strSql = "Select level as 级数,0 As 末级,ID,上级ID,编码,名称,'' As 说明 From 诊疗参考分类 a" & _
            " Where 类型=" & iAttr & _
            " Start With a.上级id Is Null Connect By Prior a.id=a.上级id " & _
            " Union All" & _
            " Select 999 as 级数,1,ID,分类ID,编码,名称,说明 From 诊疗参考目录 a Where 类型=" & iAttr & " Order By 级数,编码"
    Else
        strSQLItem = " From 诊疗参考目录 A,诊疗参考别名 B" & _
            " Where A.ID=B.参考目录ID And A.类型=" & iAttr & _
            " And (Upper(A.编码) Like '" & UCase(strName) & "%'" & _
            " Or Upper(A.名称) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.名称) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.简码) Like '" & mstrMatch & UCase(strName) & "%')"

        strSql = "Select Distinct 0 As 末级,ID,上级ID,编码,名称,'' As 说明 From 诊疗参考分类 a" & _
            " Where 类型=" & iAttr & _
            " Start With ID In (Select 分类ID " & strSQLItem & ") Connect By Prior a.上级id=a.id " & _
            " Union All" & _
            " Select Distinct 1,A.ID,A.分类ID,A.编码,A.名称,A.说明 " & strSQLItem & " Order By 编码"
    End If
    Set SelectRefer = zlDatabase.ShowSelect(Me, strSql, 2, "参考", , , , , True)
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub cmd分类_Click()
    With Me.tvwClass
        .Left = Me.txt分类.Left
        .Top = Me.txt分类.Top + Me.txt分类.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
'提取执行项目的信息
    Dim lngCol As Long
    Dim intBe As Integer
    
    Err = 0: On Error GoTo ErrHand
    
    If mblnDosage = True Then Exit Sub

    gstrSql = "select A.编码,A.名称,A.标本部位 as 说明,A.建档时间,nvl(A.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间," & _
              " A.参考目录id,B.名称 As 参考名称,A.站点 " & _
              " from 诊疗项目目录 A,诊疗参考目录 B" & _
              " where A.参考目录Id = B.id(+) And A.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)

    With rsTemp
        Me.txt编码.MaxLength = .Fields("编码").DefinedSize
        If .RecordCount > 0 Then
            Me.txt编码.Text = !编码: Me.txt名称(0).Text = !名称: Me.txt说明.Text = IIf(IsNull(!说明), "", !说明)
            Me.txt参考.Text = NVL(!参考名称)
            Me.txt参考.Tag = NVL(!参考目录ID)
            SetStationNo IIf(IsNull(!站点), "", !站点)
            strRefer = Me.txt参考.Text
        End If
    End With

'    gstrSql = "Select 0 As 类型, b.序号, a.Id, a.编码, a.名称, a.计算单位, b.单次用量, b.医生嘱托" & vbNewLine & _
'        "From 诊疗项目目录 A, 诊疗项目组合 B" & vbNewLine & _
'        "Where a.Id = b.诊疗项目id And b.收费细目id Is Null And b.诊疗组合id = [1]" & vbNewLine & _
'        "Union All" & vbNewLine & _
'        "Select 1 As 类型, b.序号, a.Id, a.编码, a.名称 || '(' || a.规格 || ')' 名称, a.计算单位, b.单次用量, b.医生嘱托" & vbNewLine & _
'        "From 收费项目目录 A, 诊疗项目组合 B" & vbNewLine & _
'        "Where a.Id = b.收费细目id And b.诊疗项目id Is Null And b.诊疗组合id = [1]" & vbNewLine & _
'        "Order By 序号"

    gstrSql = "Select b.序号, b.诊疗项目Id As 药名id, b.收费细目Id As 规格id, a.名称, c.规格, a.计算单位, b.单次用量, b.医生嘱托 " & vbNewLine & _
        "From 诊疗项目目录 A, 诊疗项目组合 B, 收费项目目录 C " & vbNewLine & _
        "Where a.Id = b.诊疗项目id And b.收费细目id = c.Id(+) And b.诊疗组合id = [1] " & vbNewLine & _
        "Order By b.序号 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    
    With rsTemp
        Me.msfRecipe.ClearBill
        Do While Not .EOF
            If Me.msfRecipe.Rows - 1 < ((.AbsolutePosition - 1) \ mbyt中药味数) + 1 Then Me.msfRecipe.Rows = Me.msfRecipe.Rows + 1
            intFence = (.AbsolutePosition - 1) Mod mbyt中药味数
            Me.msfRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt中药味数 + 1, intFence * 配方列表.列数 + 配方列表.药名ID) = !药名ID
            Me.msfRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt中药味数 + 1, intFence * 配方列表.列数 + 配方列表.规格ID) = IIf(IsNull(!规格ID), 0, !规格ID)
            Me.msfRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt中药味数 + 1, intFence * 配方列表.列数 + 配方列表.名称) = !名称 & IIf(IsNull(!规格), "", "(" & !规格 & ")")
            Me.msfRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt中药味数 + 1, intFence * 配方列表.列数 + 配方列表.数量) = FormatEx(IIf(IsNull(!单次用量), 0, !单次用量), 2)
            Me.msfRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt中药味数 + 1, intFence * 配方列表.列数 + 配方列表.单位) = IIf(IsNull(!计算单位), "", !计算单位)
            Me.msfRecipe.TextMatrix((.AbsolutePosition - 1) \ mbyt中药味数 + 1, intFence * 配方列表.列数 + 配方列表.脚注) = IIf(IsNull(!医生嘱托), "", !医生嘱托)
            .MoveNext
        Loop
    End With

    gstrSql = "select R.用法ID,R.性质,R.频次,R.疗程" & _
              " from 诊疗用法用量 R" & _
              " where R.项目ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)

    With rsTemp
        Do While Not .EOF
            If !性质 = 0 Then
                For intCount = 0 To Me.cbo用法.ListCount - 1
                    If Me.cbo用法.ItemData(intCount) = !用法ID Then Me.cbo用法.ListIndex = intCount: Exit For
                Next
            End If
            If !性质 = 1 Then
                For intCount = 0 To Me.cbo煎法.ListCount - 1
                    If Me.cbo煎法.ItemData(intCount) = !用法ID Then Me.cbo煎法.ListIndex = intCount: Exit For
                Next
            End If
            For intCount = 0 To Me.cbo频率.ListCount - 1
                If Left(Me.cbo频率.List(intCount), InStr(1, Me.cbo频率.List(intCount), "-") - 1) = IIf(IsNull(!频次), "", !频次) Then
                    Me.cbo频率.ListIndex = intCount: Exit For
                End If
            Next
            Me.txt疗程.Text = IIf(IsNull(!疗程), 0, !疗程)
            .MoveNext
        Loop
    End With

    If Me.Tag = "增加" Then
        lngItemId = 0

        If Val(zlDatabase.GetPara(61, glngSys)) = 0 Then    '诊疗项目编码递增模式
            gstrSql = "select nvl(max(编码),'0000000') as 编码" & _
                      " From 诊疗项目目录"
            '            If rsTemp.State = adStateOpen Then rsTemp.Close
            '            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Activate")
            '            Call SQLTest
            Me.txt编码.Text = Right(String(10, "0") & Val(rsTemp!编码) + 1, Len(rsTemp!编码))
        Else
            strTemp = Mid(Me.txt分类.Text, 2, InStr(1, Me.txt分类.Text, "]") - 2)
            gstrSql = "select nvl(max(编码),'0000000') as 编码" & _
                      " From 诊疗项目目录" & _
                      " Where  编码 like [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, "8" & strTemp & "%")
            Err = 0: On Error Resume Next
            Me.txt编码.Text = "8" & strTemp & Right(String(10, "0") & Val(rsTemp!编码) + 1, Len(rsTemp!编码) - 1 - Len(strTemp))
        End If

        Me.txt名称(0).Text = ""
        Me.txt参考 = "": Me.txt参考.Tag = "": strRefer = ""
    Else
        gstrSql = "select 名称,性质,简码,码类 from 诊疗项目别名 where 诊疗项目ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
        With rsTemp
            Do While Not .EOF
                If !性质 = 1 And !码类 = 1 Then Me.txt拼音(0).Text = !简码
                If !性质 = 1 And !码类 = 2 Then Me.txt五笔(0).Text = !简码
                If !性质 = 9 Then Me.txt名称(1).Text = !名称
                If !性质 = 9 And !码类 = 1 Then Me.txt拼音(1).Text = !简码
                If !性质 = 9 And !码类 = 2 Then Me.txt五笔(1).Text = !简码
                .MoveNext
            Loop
        End With
    End If

    If Me.Tag = "查阅" Then
        Me.cmdOK.Visible = False
        Me.cmdCancel.Caption = "关闭(&C)"
        Me.txt分类.Enabled = False: Me.cmd分类.Enabled = False
        Me.txt编码.Enabled = False
        Me.txt名称(0).Enabled = False: Me.txt拼音(0).Enabled = False: Me.txt五笔(0).Enabled = False
        Me.txt名称(1).Enabled = False: Me.txt拼音(1).Enabled = False: Me.txt五笔(1).Enabled = False
        Me.txt说明.Enabled = False
        Me.cbo频率.Enabled = False: Me.cbo煎法.Enabled = False: Me.cbo用法.Enabled = False
        Me.txt疗程.Enabled = False: Me.msfRecipe.Active = False
        Me.txt参考.Enabled = False
        Me.cmd参考.Enabled = False
    End If
    
    '设置颜色
    For lngCol = 0 To msfRecipe.Cols - 1
        If lngCol Mod 配方列表.列数 = 0 Then
            For intBe = 0 To 配方列表.列数 - 1
                msfRecipe.SetColColor lngCol + intBe, &H8000000F
            Next
            lngCol = lngCol + 配方列表.列数
        End If
    Next
    mblnLoad = True
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub IniStationNo()
    Dim dblHeight As Double
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    
'    lblStationNo.Visible = False
'    cmbStationNo.Visible = False
'
'    If gstrNodeNo <> "-" Then
    On Error GoTo ErrHand
    lblStationNo.Visible = True
    cmbStationNo.Visible = True
    
    strSql = "select 编号,名称 from zlnodelist"
    Set rsRecord = zlDatabase.OpenSQLRecord(strSql, "站点查询")
    With cmbStationNo
        .AddItem ""
        Do While Not rsRecord.EOF
            .AddItem rsRecord!编号 & "-" & rsRecord!名称
            rsRecord.MoveNext
        Loop
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetStationNo(ByVal strNo As String)
    Dim n As Integer
    
'    If gstrNodeNo = "-" Then Exit Sub
    
    If strNo = "" Then
        cmbStationNo.ListIndex = 0
    Else
        For n = 1 To cmbStationNo.ListCount - 1
            If Mid(cmbStationNo.List(n), 1, InStr(1, cmbStationNo.List(n), "-") - 1) = strNo Then
                cmbStationNo.ListIndex = n
            End If
        Next
    End If
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.tvwClass.Visible Then
        Me.tvwClass.Visible = False: Me.txt分类.SetFocus: Exit Sub
    ElseIf Me.lvwItems.Visible Then
        Me.lvwItems.Visible = False: Me.msfRecipe.SetFocus: Exit Sub
    End If
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call GetDefineSize
    Call IniStationNo
    
    ''中药配方每行中药味数
    mbyt中药味数 = zlDatabase.GetPara(213, glngSys)
    With Me.msfRecipe
        strTemp = "   配方组成：  (*按顺序选择药品、输入单味用量、必要时填写煎服脚注)"
        intCount = (.Width - Me.TextWidth(strTemp)) \ Me.TextWidth(Space(1))
        strTemp = strTemp & Space(intCount - 2)
        .Active = True
        .Rows = 2: .Cols = mbyt中药味数 * 配方列表.列数
        .MsfObj.AllowUserResizing = flexResizeNone
        .MsfObj.ScrollBars = flexScrollBarBoth 'flexScrollBarVertical
        .MsfObj.GridColor = &H80000005: .MsfObj.BackColorBkg = &H80000005
        .MsfObj.MergeCells = flexMergeFree
        .MsfObj.MergeRow(0) = True
    
        .TxtCheck = True
        '        .TextMask = mstrChar
        For intCount = 0 To .Cols - 1
            .TextMatrix(0, intCount) = strTemp
            Select Case (intCount Mod 配方列表.列数)
            Case 配方列表.空白
                .ColData(intCount) = 5: .ColWidth(intCount) = IIf(mbyt中药味数 = 4, 200, 370)
            Case 配方列表.药名ID
                .ColData(intCount) = 5: .ColWidth(intCount) = 0
            Case 配方列表.规格ID
                .ColData(intCount) = 5: .ColWidth(intCount) = 0
            Case 配方列表.名称
                .ColData(intCount) = 1: .ColWidth(intCount) = IIf(mbyt中药味数 = 4, 1500, 1700)
            Case 配方列表.数量
                .ColData(intCount) = 4: .ColWidth(intCount) = IIf(mbyt中药味数 = 4, 500, 700)
            Case 配方列表.单位
                .ColData(intCount) = 5: .ColWidth(intCount) = IIf(mbyt中药味数 = 4, 300, 400)
            Case 配方列表.脚注
                .ColData(intCount) = 4: .ColWidth(intCount) = IIf(mbyt中药味数 = 4, 850, 1000)
            End Select
        Next
'        .PrimaryCol = 3: .LocateCol = 3
        .Row = 1: .Col = 2
    End With

    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 1500
        .Add , "编码", "编码", 900
    End With
    With Me.lvwItems
        .Width = 2600
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1
        .SortOrder = lvwAscending
    End With
    mstrMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
    strRefer = ""
    
    mblnOK = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnDosage = False
    mblnLoad = False
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    Dim lngCol As Long
    Dim intBe As Integer
    
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItems
        If .SelectedItem.Icon = "药品" Then
            Me.msfRecipe.Text = .SelectedItem.Text
            Me.msfRecipe.TextMatrix(Me.msfRecipe.Row, Me.msfRecipe.Col - 1) = Mid(.SelectedItem.Key, 2)
            Me.msfRecipe.TextMatrix(Me.msfRecipe.Row, Me.msfRecipe.Col) = Me.msfRecipe.Text
            Me.msfRecipe.TextMatrix(Me.msfRecipe.Row, Me.msfRecipe.Col + 2) = .SelectedItem.Tag
        Else
            Me.msfRecipe.Text = .SelectedItem.Text
            Me.msfRecipe.TextMatrix(Me.msfRecipe.Row, Me.msfRecipe.Col) = Me.msfRecipe.Text
            
            With msfRecipe
                If Val(.TextMatrix(.Row, .Col - 4)) <> 0 Then 'id为空不能新增行
                    If .Row = .Rows - 1 And .Col = .Cols - 1 Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                        .Col = 3
                    End If
                    
                    '设置颜色
                    For lngCol = 0 To .Cols - 1
                        If lngCol Mod 配方列表.列数 = 0 Then
                            For intBe = 0 To 配方列表.列数 - 1
                                .SetColColor lngCol + intBe, &H8000000F
                            Next
                            lngCol = lngCol + 配方列表.列数
                        End If
                    Next
                    
                End If
            End With
        End If
        Me.msfRecipe.SetFocus
        Call zlCommFun.PressKey(vbKeyReturn)
    End With
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Private Sub msfRecipe_CommandClick()
    Dim i As Integer
    Dim intDosageType As Integer
    Dim intCount As Integer
    Dim strName As String
    Dim intCurRow As Integer
    Dim intCurCol As Integer
    
'    If (Me.msfRecipe.Cols Mod 7) <> 2 Then Exit Sub
    For i = 0 To optDosageType.UBound
        If optDosageType(i).Value = True Then
            intDosageType = i
            Exit For
        End If
    Next
    
    intCurRow = msfRecipe.Row
    intCurCol = msfRecipe.Col
    mblnDosage = True
    frmMediDosage.ShowMe intDosageType, Me, "", strName
    
    If strName <> "" Then
        With msfRecipe
            If CheckDoubDosage(strName) = False Then
                MsgBox "该药品已经存在，重复药品不能录入多次！", vbInformation, gstrSysName
                msfRecipe.Row = intCurRow
                msfRecipe.Col = intCurCol
                msfRecipe.SetFocus
                Exit Sub
            Else
                .Row = intCurRow
                .Col = intCurCol
                .SetFocus
                msfRecipe.TextMatrix(intCurRow, intCurCol - 2) = Split(strName, ",")(0) '药名id
                Me.msfRecipe.TextMatrix(intCurRow, intCurCol - 1) = Split(strName, ",")(1) '规格id
                Me.msfRecipe.Text = Split(strName, ",")(2)  '名称
                Me.msfRecipe.TextMatrix(intCurRow, intCurCol) = Me.msfRecipe.Text '名称
                Me.msfRecipe.TextMatrix(intCurRow, intCurCol + 2) = Split(strName, ",")(3) '单位
            End If
        End With
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckDoubDosage(ByVal strName As String) As Boolean
    '检查配方中药品是否重复
    Dim intRow As Integer
    Dim intFence As Integer
    Dim intGroup As Integer
    Dim strCheckID As String '药名ID+规格ID
    
    If mbyt中药味数 = 3 Then
        intGroup = 2
    ElseIf mbyt中药味数 = 4 Then
        intGroup = 3
    End If
    
    strCheckID = Split(strName, ",")(0) & "+" & Split(strName, ",")(1)
    
    With msfRecipe
        For intRow = 1 To .Rows - 1
            For intFence = 0 To intGroup
                If Val(.TextMatrix(intRow, 配方列表.列数 * intFence + 配方列表.药名ID)) <> 0 Then  'id不能为空
                    If strCheckID = .TextMatrix(intRow, 配方列表.列数 * intFence + 配方列表.药名ID) & "+" & .TextMatrix(intRow, 配方列表.列数 * intFence + 配方列表.规格ID) Then
                        Exit Function
                    End If
                    
'                    If InStr(1, Mid(strName, 1, InStr(3, strName, ",") - 1), .TextMatrix(intRow, 配方列表.列数 * intFence + 1) & "," & Trim(.TextMatrix(intRow, 配方列表.列数 * intFence + 2))) > 0 Then
'                        Exit Function
'                    End If
                End If
            Next
        Next
        CheckDoubDosage = True
    End With
End Function

Private Sub msfRecipe_EditChange(curText As String)
    mstrMedi = msfRecipe.TextMatrix(msfRecipe.Row, msfRecipe.Col)
End Sub

Private Sub msfRecipe_EditKeyDown(KeyCode As Integer, Shift As Integer)
'    If InStr("'`?|/\;,%", Chr(KeyCode)) > 0 Then KeyCode = 0
End Sub

Private Sub msfRecipe_EditKeyPress(KeyAscii As Integer)
    If InStr("'`?|/\;,%", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub msfRecipe_EnterCell(Row As Long, Col As Long)
    strInputed = Me.msfRecipe.TextMatrix(Row, Col)
End Sub

Private Sub msfRecipe_GotFocus()
    If Me.lvwItems.Visible Then Me.lvwItems.SetFocus
End Sub

Private Sub msfRecipe_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim i As Integer
    Dim intDosageType As Integer
    Dim intCount As Integer
    Dim strName As String
    Dim intCurRow As Integer
    Dim intCurCol As Integer
    Dim lngCol As Long
    Dim intBe As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.msfRecipe.Active = False Then Exit Sub
    With Me.msfRecipe
'        .TextMask = mstrChar
        Select Case (.Col Mod 配方列表.列数)
        Case 配方列表.空白, 配方列表.药名ID, 配方列表.规格ID, 配方列表.单位
            Exit Sub
        Case 配方列表.数量
            If .TxtVisible = False Then
                If Trim(.TextMatrix(.Row, .Col)) = "" Then .TextMatrix(.Row, .Col) = "0"
                .TextMatrix(.Row, .Col) = FormatEx(.TextMatrix(.Row, .Col), 2)
            Else
                If Trim(.Text) = "" Then .Text = 0: .TextMatrix(.Row, .Col) = "0"
                .Text = FormatEx(.Text, 2)
            End If
'            If Val(.Text) = 0 Then
'                Cancel = True
'            End If
            Exit Sub
        Case 配方列表.脚注
            .TextMask = ""
            If .TxtVisible = False Then
                If Trim(.TextMatrix(.Row, .Col)) = "" Then .TextMatrix(.Row, .Col) = Space(1)
                strTemp = UCase(Trim(.TextMatrix(.Row, .Col)))
            Else
                If Trim(.Text) = "" Then .Text = Space(1): .TextMatrix(.Row, .Col) = Space(1)
                strTemp = UCase(Trim(.Text))
            End If
            
            If strTemp = "" Or Not IsNumeric(strTemp) Then
                With msfRecipe
                    If Val(.TextMatrix(.Row, .Col - 4)) = 0 And Val(.TextMatrix(.Row, .Col - 5)) = 0 Then 'id为空不能新增行
                    Else
                        If .Row = .Rows - 1 And .Col = .Cols - 1 Then
                            .Rows = .Rows + 1
                            .Row = .Rows - 1
                            .Col = 3
                            
                            '设置颜色
                            For lngCol = 0 To .Cols - 1
                                If lngCol Mod 配方列表.列数 = 0 Then
                                    For intBe = 0 To 配方列表.列数 - 1
                                        .SetColColor lngCol + intBe, &H8000000F
                                    Next
                                    lngCol = lngCol + 配方列表.列数
                                End If
                            Next
                        End If
                    End If
                End With
                Exit Sub
            End If
            
            gstrSql = "select 编码,名称" & _
                    " from 中药煎服脚注" & _
                    " where (编码 like [1] or 名称 like [2] or 简码 like [2])" & _
                    " order by 编码"
            Err = 0: On Error GoTo ErrHand
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
            
            With rsTemp
                If .BOF Or .EOF Then Exit Sub
                If .RecordCount = 1 Then Me.msfRecipe.Text = !名称: Me.msfRecipe.TextMatrix(Me.msfRecipe.Row, Me.msfRecipe.Col) = Me.msfRecipe.Text: Exit Sub
                Me.lvwItems.ListItems.Clear
                Do While Not .EOF
                    Set objItem = Me.lvwItems.ListItems.Add(, "_" & !编码, !名称)
                    objItem.Icon = "脚注": objItem.SmallIcon = "脚注"
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
                    .MoveNext
                Loop
                Me.lvwItems.ListItems(1).Selected = True
            End With
        Case 配方列表.名称
            If .TxtVisible = False Then
                If .TextMatrix(.Row, .Col) = "" Then
                    If .Col <> 3 Then Call zlCommFun.PressKey(vbKeyTab)
                    Exit Sub
                End If
                strTemp = UCase(Trim(.TextMatrix(.Row, .Col)))
            Else
                If Trim(.Text) = "" Then
                    If .Col <> 3 Then .SetFocus: Call zlCommFun.PressKey(vbKeyTab)
                    Exit Sub
                End If
                strTemp = UCase(Trim(.Text))
            End If
            If strInputed = strTemp Then Exit Sub
            
            For i = 0 To optDosageType.UBound
                If optDosageType(i).Value = True Then
                    intDosageType = i
                    Exit For
                End If
            Next
            
            intCurRow = msfRecipe.Row
            intCurCol = msfRecipe.Col
            mblnDosage = True
            frmMediDosage.ShowMe intDosageType, Me, strTemp, strName
            
            If strName <> "" Then
                With msfRecipe
                    If CheckDoubDosage(strName) = False Then
                        MsgBox "该药品已经存在，重复药品不能录入多次！", vbInformation, gstrSysName
                        .Row = intCurRow
                        .Col = intCurCol
                        .SetFocus
                        Exit Sub
                    Else
                        .Row = intCurRow
                        .Col = intCurCol
                        .SetFocus
                        msfRecipe.TextMatrix(intCurRow, intCurCol - 2) = Split(strName, ",")(0) '药名id
                        Me.msfRecipe.TextMatrix(intCurRow, intCurCol - 1) = Split(strName, ",")(1) '规格id
                        Me.msfRecipe.Text = Split(strName, ",")(2) '名称
                        Me.msfRecipe.TextMatrix(intCurRow, intCurCol) = Me.msfRecipe.Text '名称
                        Me.msfRecipe.TextMatrix(intCurRow, intCurCol + 2) = Split(strName, ",")(3) '单位
                    End If
                End With
            Else
                .Text = mstrMedi
                Me.msfRecipe.TextMatrix(intCurRow, intCurCol) = .Text '名称
            End If
            Exit Sub
        End Select
    End With
    
    With Me.lvwItems
        .Left = Me.msfRecipe.Left
        For intCount = 0 To Me.msfRecipe.Col - 1
            .Left = .Left + Me.msfRecipe.ColWidth(intCount)
        Next
        .Top = Me.msfRecipe.Top
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Cancel = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msfRecipe_KeyPress(KeyAscii As Integer)
    If InStr("'`?|/\;,%", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub optDosageType_Click(Index As Integer)
    Dim intCol As Integer
    Dim intGroup As Integer
    Dim intFence As Integer
    Dim blnHaveData As Boolean
    
    If mblnLoad = True Then '界面加载完了后才允许触发
        If (mintOldShape = 1 And Index = 2) Or (mintOldShape = 2 And Index = 1) Then
        Else
            If mbyt中药味数 = 3 Then
                intGroup = 2
            ElseIf mbyt中药味数 = 4 Then
                intGroup = 3
            End If
            For intCount = 1 To msfRecipe.Rows - 1
                For intFence = 0 To intGroup
                    If Val(msfRecipe.TextMatrix(intCount, 配方列表.列数 * intFence + 配方列表.规格ID)) <> 0 Or Val(msfRecipe.TextMatrix(intCount, 配方列表.列数 * intFence + 配方列表.药名ID)) <> 0 Then   'id不能为空
                        blnHaveData = True
                        Exit For
                    End If
                Next
                If blnHaveData = True Then
                    Exit For
                End If
            Next
            If blnHaveData = True And mblnClickNo = False Then
                If MsgBox("形态改变，将清空配方表格中数据，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    With msfRecipe
                        .Rows = 2
                        For intCol = 0 To .Cols - 1
                            .TextMatrix(1, intCol) = ""
                        Next
                    End With
                Else
                    mblnClickNo = True
                    optDosageType(mintOldShape).Value = True
                End If
            End If
        End If
    End If
End Sub

Private Sub optDosageType_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    
    mintOldShape = 0
    mblnClickNo = False
    For i = optDosageType.LBound To optDosageType.UBound
        If optDosageType(i).Value = True Then
            mintOldShape = i
            Exit For
        End If
    Next
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txt分类.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    Me.txt分类.Text = Me.tvwClass.SelectedItem.Text
    Me.txt分类.SetFocus
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
    If Me.cmd分类 Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub

Private Sub txt编码_GotFocus()
    Me.txt编码.SelStart = 0: Me.txt编码.SelLength = 100
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
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

Private Sub txt参考_GotFocus()
    Me.txt参考.SelStart = 0: Me.txt参考.SelLength = 100
End Sub


Private Sub txt参考_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        If Me.txt参考 <> strRefer Then
            Set rsTmp = SelectRefer(Trim(Me.txt参考))
            If rsTmp Is Nothing Then
                Me.txt参考 = strRefer
                Me.SetFocus
                MsgBox "没有找到可参考的项目。", vbInformation, Me.Caption
                Exit Sub
            Else
                Me.txt参考 = rsTmp("名称"): Me.txt参考.Tag = rsTmp("ID"): strRefer = Me.txt参考
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If InStr(" ~!@#$%^&|=`;'""?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub txt参考_LostFocus()
    If Me.txt参考 <> strRefer Then
        Me.txt参考 = strRefer
    End If
End Sub


Private Sub txt分类_GotFocus()
    Me.txt分类.SelStart = 0: Me.txt分类.SelLength = 100
End Sub

Private Sub txt分类_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt疗程_GotFocus()
    Me.txt疗程.SelStart = 0: Me.txt疗程.SelLength = 100
End Sub

Private Sub txt疗程_KeyPress(KeyAscii As Integer)
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

Private Sub txt名称_GotFocus(Index As Integer)
    Me.txt名称(Index).SelStart = 0: Me.txt名称(Index).SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt名称(Index).Text = MoveSpecialChar(txt名称(Index).Text)
        Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt名称_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Me.txt拼音(Index).Text = zlStr.GetCodeByORCL(Me.txt名称(Index).Text, False, Me.txt拼音(Index).MaxLength)
    Me.txt五笔(Index).Text = zlStr.GetCodeByORCL(Me.txt名称(Index).Text, True, Me.txt五笔(Index).MaxLength)
End Sub

Private Sub txt名称_LostFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt拼音_GotFocus(Index As Integer)
    Me.txt拼音(Index).SelStart = 0: Me.txt拼音(Index).SelLength = 100
End Sub

Private Sub txt拼音_KeyPress(Index As Integer, KeyAscii As Integer)
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

Private Sub txt说明_GotFocus()
    Me.txt说明.SelStart = 0: Me.txt说明.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt说明_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt五笔_GotFocus(Index As Integer)
    Me.txt五笔(Index).SelStart = 0: Me.txt五笔(Index).SelLength = 100
End Sub

Private Sub txt五笔_KeyPress(Index As Integer, KeyAscii As Integer)
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



