VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemsEdit 
   Caption         =   "体检项目设置"
   ClientHeight    =   6375
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   10650
   Icon            =   "frmItemsEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraTitle 
      Height          =   600
      Left            =   75
      TabIndex        =   3
      Top             =   15
      Width           =   6870
      Begin VB.ComboBox cbo 
         Height          =   300
         Left            =   4020
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   195
         Width           =   2745
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&1.组别"
         Height          =   180
         Index           =   4
         Left            =   3435
         TabIndex        =   19
         Top             =   255
         Width           =   540
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检项目设置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   90
         TabIndex        =   4
         Top             =   195
         Width           =   1800
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6015
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmItemsEdit.frx":076A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13705
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fra2 
      Height          =   3000
      Left            =   45
      TabIndex        =   1
      Top             =   525
      Width           =   8475
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   2
         Left            =   4620
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   14
         Top             =   180
         Width           =   1020
      End
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   1
         Left            =   3075
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   13
         Top             =   180
         Width           =   870
      End
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   0
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   180
         Width           =   930
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   6
         Left            =   7995
         Picture         =   "frmItemsEdit.frx":0FFE
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "快捷键：F4"
         Top             =   165
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   5
         Left            =   7620
         Picture         =   "frmItemsEdit.frx":1E40
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "快捷键：F3"
         Top             =   165
         Width           =   345
      End
      Begin zl9Medical.VsfGrid vsf 
         Height          =   1650
         Left            =   90
         TabIndex        =   0
         Top             =   555
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   2910
      End
      Begin zl9Medical.VsfGrid vsfPrice 
         Height          =   1635
         Left            =   5070
         TabIndex        =   11
         Top             =   1005
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   2884
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "折扣(Z)"
         Height          =   180
         Index           =   17
         Left            =   3975
         TabIndex        =   17
         Top             =   240
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检价格(E)"
         Height          =   180
         Index           =   18
         Left            =   2070
         TabIndex        =   16
         Top             =   240
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "基本价格(&B)"
         Height          =   180
         Index           =   19
         Left            =   90
         TabIndex        =   15
         Top             =   240
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   9825
      Top             =   3645
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
            Picture         =   "frmItemsEdit.frx":23CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemsEdit.frx":7434
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemsEdit.frx":772E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemsEdit.frx":7CC8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraButton 
      Height          =   705
      Left            =   90
      TabIndex        =   5
      Top             =   3450
      Width           =   8460
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   5775
         TabIndex        =   8
         Top             =   210
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   6975
         TabIndex        =   7
         Top             =   210
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   75
         TabIndex        =   6
         Top             =   210
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmItemsEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngLoop As Long
Private mblnDataChange As Boolean
Private mrsItems As New ADODB.Recordset
Private mblnChanged As Boolean
Private mblnGroup As Boolean
Private mlngDept As Long
Private mstrGroup As String
Private mbytMode As Byte
Private mlngKey As Long
Private mlngArchiveKey As Long
Private mblnNo As Boolean
Private mstrSQL As String
Private mstr性别 As String
Private Enum mCol
    项目 = 1
    执行科室
    检查部位
    采集方式
    采集科室
    检验标本
    基本价格
    体检价格
    折扣
    体检类型
    类别
    结算方式
    执行科室id
    采集方式id
    采集科室id
    检查部位id
    计费明细
    新加
    前景色
    删除
    公共
    清单id
    
    p计价项目 = 1
    p名称
    p计算单位
    p数次
    p标准单价
    p体检单价
    p折扣
    p标准金额
    p体检金额
    p执行科室
    p执行科室id
    p收费项目id
    p计价性质
    p类别
    p可用库存
End Enum

'（２）自定义过程或函数************************************************************************************************
Private Property Let DataChange(ByVal vData As Boolean)
        mblnDataChange = vData
End Property

Private Property Get DataChange() As Boolean
        DataChange = mblnDataChange
End Property

Private Function CreatePriceList(ByVal intRow As Integer) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim strKeys As String
    
    strKeys = CStr(Val(vsf.RowData(intRow))) & "'" & CStr(Val(vsf.TextMatrix(intRow, mCol.采集方式id))) & "'" & vsf.TextMatrix(intRow, mCol.检查部位id)
    
    Dim str计价项目 As String
    Dim str计价性质 As String
    
    vsfPrice.Rows = 2
    str计价项目 = vsfPrice.TextMatrix(1, mCol.p计价项目)
    str计价性质 = vsfPrice.TextMatrix(1, mCol.p计价性质)
    
    vsfPrice.Body.Cell(flexcpText, 1, mCol.p计价项目 + 1, 1, vsfPrice.Cols - 1) = ""
    vsfPrice.RowData(1) = 0

    vsfPrice.TextMatrix(1, mCol.p计价项目) = str计价项目
    vsfPrice.TextMatrix(1, mCol.p计价性质) = str计价性质
    
    If vsfPrice.ComboList(mCol.p计价项目) <> "" Then
        vsfPrice.TextMatrix(1, mCol.p计价项目) = Split(vsfPrice.ComboList(mCol.p计价项目), "|")(0)
    End If
    
    mstrSQL = GetPublicSQL(SQL.体检项目价表, strKeys)
    If vsf.TextMatrix(intRow, mCol.检查部位id) = "" Then
        '检验或单部位检查
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, Val(vsf.RowData(intRow)), Val(vsf.TextMatrix(intRow, mCol.采集方式id)))
    Else
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    End If
    
    If rs.BOF = False Then
        With vsfPrice
            Do While Not rs.EOF
                
                If Val(.TextMatrix(.Rows - 1, mCol.p收费项目id)) > 0 Then
                    .Rows = .Rows + 1
                End If
                
                If zlCommFun.NVL(rs("计价性质")) = 2 Then
                    .TextMatrix(.Rows - 1, mCol.p计价项目) = "采集方式-" & vsf.TextMatrix(vsf.Row, mCol.采集方式)
                ElseIf vsf.TextMatrix(vsf.Row, mCol.类别) = "检验" Then
                    .TextMatrix(.Rows - 1, mCol.p计价项目) = "检验项目-" & vsf.TextMatrix(vsf.Row, mCol.项目)
                Else
                    .TextMatrix(.Rows - 1, mCol.p计价项目) = "检查项目-" & vsf.TextMatrix(vsf.Row, mCol.项目)
                End If
                
                .TextMatrix(.Rows - 1, mCol.p名称) = zlCommFun.NVL(rs("名称"))
                .TextMatrix(.Rows - 1, mCol.p计算单位) = zlCommFun.NVL(rs("计算单位"))
                .TextMatrix(.Rows - 1, mCol.p数次) = zlCommFun.NVL(rs("收费数量"))
                .TextMatrix(.Rows - 1, mCol.p标准单价) = zlCommFun.NVL(rs("现价"))
                .TextMatrix(.Rows - 1, mCol.p体检单价) = zlCommFun.NVL(rs("现价"))
                .TextMatrix(.Rows - 1, mCol.p折扣) = 10
                .TextMatrix(.Rows - 1, mCol.p标准金额) = zlCommFun.NVL(rs("收费数量"), 0) * zlCommFun.NVL(rs("现价"), 0)
                .TextMatrix(.Rows - 1, mCol.p体检金额) = zlCommFun.NVL(rs("收费数量"), 0) * zlCommFun.NVL(rs("现价"), 0)
                .TextMatrix(.Rows - 1, mCol.p收费项目id) = zlCommFun.NVL(rs("ID"))
                
                .TextMatrix(vsfPrice.Rows - 1, mCol.p计价性质) = zlCommFun.NVL(rs("计价性质"))
                
                .TextMatrix(vsfPrice.Rows - 1, mCol.p类别) = zlCommFun.NVL(rs("类别"))
                
                Call SetRowDefault(zlCommFun.NVL(rs("ID"), 0), vsfPrice.Rows - 1, "收费执行科室")
                
                If InStr("567", .TextMatrix(.Rows - 1, mCol.p类别)) > 0 Then
                    .TextMatrix(.Rows - 1, mCol.p可用库存) = GetStorage(Val(.RowData(.Rows - 1)), Val(.TextMatrix(.Rows - 1, mCol.p执行科室id)))
                    Call PromptStorageWarn(Val(.TextMatrix(.Rows - 1, mCol.p数次)), Val(.TextMatrix(.Rows - 1, mCol.p可用库存)), .TextMatrix(.Rows - 1, mCol.p名称), .TextMatrix(.Rows - 1, mCol.p执行科室), .TextMatrix(.Rows - 1, mCol.p计算单位), 1)
                End If
                                
                rs.MoveNext
            Loop
        End With
    End If
    
    vsf.TextMatrix(intRow, mCol.基本价格) = SumPrice(1)
    vsf.TextMatrix(intRow, mCol.体检价格) = SumPrice(2)
    
End Function


Public Function ShowEdit(ByVal frmMain As Object, _
                        ByVal lngKey As Long, _
                        ByRef rsItems As ADODB.Recordset, _
                        ByVal lngDept As Long, _
                        Optional blnGroup As Boolean = False, _
                        Optional ByVal bytMode As Byte = 1, _
                        Optional ByVal lngArchiveKey As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:  frmMain         调用窗体对象
    '       lngKey          预约登记id
    '返回:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    Dim varGroup As Variant
    
    mblnStartUp = True
    mblnOK = False

    'bytMode    1表示接受前,2表示接受后
    mblnNo = True
    
    Set mfrmMain = frmMain
        
    mlngKey = lngKey
    Call CopyRecord(rsItems, mrsItems)
    mblnGroup = blnGroup
    mlngDept = lngDept
    mbytMode = bytMode
    mlngArchiveKey = lngArchiveKey
    
    Call ClearData
    If InitData = False Then Exit Function
    If ReadData() = False Then Exit Function
            
    DataChange = False
    
    mblnNo = False
    
    Call cbo_Click
    
    vsf.Col = 2
    vsf.Col = 1
    
    Me.Show 1, frmMain
    
    rsItems.Filter = ""
    If mblnOK Then Call CopyRecord(mrsItems, rsItems)
    
    ShowEdit = mblnOK
    
End Function

Private Function ChangeTotal(ByVal dbMoney As Double, ByVal dbTmp As Double, Optional ByVal bytMode As Byte = 1) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim db折扣 As Double
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngRow As Long
    Dim dbTotal As Double

    If bytMode = 1 Then
        '变化金额
        
        If dbMoney = 0 Then Exit Function
        '1.计算折扣
        db折扣 = Format(10 * dbTmp / dbMoney, "0.0000")

    Else
        '变化折扣
        db折扣 = dbTmp

    End If
    
    txtSum(1).Text = Format(dbMoney * db折扣 / 10, "0.00")
    txtSum(2).Text = Format(db折扣, "0.0000")
    dbTotal = 0
    
    For lngLoop = 1 To vsf.Rows - 1
    
        vsf.TextMatrix(lngLoop, mCol.折扣) = db折扣
        vsf.TextMatrix(lngLoop, mCol.体检价格) = Format(vsf.TextMatrix(lngLoop, mCol.基本价格) * (db折扣 / 10), "0.00")
        
        dbTotal = dbTotal + Val(vsf.TextMatrix(lngLoop, mCol.体检价格))
                    
        varRow = Split(vsf.TextMatrix(lngLoop, mCol.计费明细), ";")
        For lngRow = 0 To UBound(varRow)
            If varRow(lngRow) <> "" Then
                varCol = Split(varRow(lngRow), ":")
                varCol(4) = Format(Val(varCol(3)) * (db折扣 / 10), "0.00000")
                varCol(10) = db折扣
            End If
            varRow(lngRow) = Join(varCol, ":")
        Next
        vsf.TextMatrix(lngLoop, mCol.计费明细) = Join(varRow, ";")
    Next

    '误差处理
    '------------------------------------------------------------------------------------------------------------------
    If dbTotal <> Val(txtSum(1).Text) Then

        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.TextMatrix(lngLoop, mCol.体检价格)) <> 0 Then
            
                vsf.TextMatrix(lngLoop, mCol.体检价格) = Val(vsf.TextMatrix(lngLoop, mCol.体检价格)) + (Val(txtSum(1).Text) - dbTotal)
                
                If Val(vsf.TextMatrix(lngLoop, mCol.基本价格)) <> 0 Then
                    vsf.TextMatrix(lngLoop, mCol.折扣) = Format(10 * Val(vsf.TextMatrix(lngLoop, mCol.体检价格)) / Val(vsf.TextMatrix(lngLoop, mCol.基本价格)), "0.0000")
                Else
                    vsf.TextMatrix(lngLoop, mCol.折扣) = 0
                End If
                
                varRow = Split(vsf.TextMatrix(lngLoop, mCol.计费明细), ";")
                For lngRow = 0 To UBound(varRow)
                    If varRow(lngRow) <> "" Then
                        varCol = Split(varRow(lngRow), ":")
                        If Val(varCol(4)) <> 0 Then
                            varCol(4) = Val(varCol(4)) + (Val(txtSum(1).Text) - dbTotal)
                            If Val(varCol(3)) <> 0 Then
                                varCol(10) = Format(10 * Val(varCol(4)) / Val(varCol(3)), "0.0000")
                            Else
                                varCol(10) = 0
                            End If
                        End If
                    End If
                    varRow(lngRow) = Join(varCol, ":")
                Next
                vsf.TextMatrix(lngLoop, mCol.计费明细) = Join(varRow, ";")
                Exit For
            End If
        Next
    End If

    ChangeTotal = True
    
End Function

Private Function ChangeItem(ByVal dbMoney As Double, ByVal dbTmp As Double, Optional ByVal bytMode As Byte = 1, Optional ByVal blnUpdate As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim dbSum As Double
    Dim db折扣 As Double
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngRow As Long
    
    
    If blnUpdate Then
        If dbMoney = 0 Then Exit Function
        
        Call WritePrice(vsf.Row)
        
        If bytMode = 1 Then
            '变化金额
            
            '1.计算折扣
            db折扣 = Format(10 * dbTmp / dbMoney, "0.0000")
        Else
            '变化折扣
            db折扣 = dbTmp
            
        End If
        
        vsf.TextMatrix(vsf.Row, mCol.体检价格) = Format(dbMoney * db折扣 / 10, "0.00")
        vsf.TextMatrix(vsf.Row, mCol.折扣) = Format(db折扣, "0.0000")
    End If
    
    '更新总体
    '------------------------------------------------------------------------------------------------------------------
    dbSum = 0
    For lngLoop = 1 To vsf.Rows - 1
       dbSum = dbSum + Val(vsf.TextMatrix(lngLoop, mCol.基本价格))
    Next
    txtSum(0).Text = Format(dbSum, "0.00")
    
    dbSum = 0
    For lngLoop = 1 To vsf.Rows - 1
       dbSum = dbSum + Val(vsf.TextMatrix(lngLoop, mCol.体检价格))
    Next
    txtSum(1).Text = Format(dbSum, "0.00")
    
    If Val(txtSum(0).Text) <> 0 Then
        txtSum(2).Text = Format(10 * Val(txtSum(1).Text) / Val(txtSum(0).Text), "0.0000")
    Else
        txtSum(2).Text = "0.0000"
    End If
    
    '更新价格
    '------------------------------------------------------------------------------------------------------------------
    If blnUpdate Then
        varRow = Split(vsf.TextMatrix(vsf.Row, mCol.计费明细), ";")
        For lngRow = 0 To UBound(varRow)
            If varRow(lngRow) <> "" Then
                varCol = Split(varRow(lngRow), ":")
                varCol(4) = Format(Val(varCol(3)) * (db折扣 / 10), "0.00000")
                varCol(10) = db折扣
            End If
            varRow(lngRow) = Join(varCol, ":")
        Next
        vsf.TextMatrix(vsf.Row, mCol.计费明细) = Join(varRow, ";")
    End If
        
    ChangeItem = True
    
End Function

Private Function ChangePrice(ByVal dbMoney As Double, ByVal dbTmp As Double, Optional ByVal bytMode As Byte = 1) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim dbSum As Double
    Dim db折扣 As Double
    
    
    If bytMode = 1 Then
        '变化金额
        If dbMoney = 0 Then Exit Function
        '1.计算折扣
        db折扣 = Format(10 * dbTmp / dbMoney, "0.0000")
    Else
        '变化折扣
        db折扣 = dbTmp
        
    End If
    
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.p体检单价) = Format(dbMoney * db折扣 / 10, "0.00000")
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.p折扣) = Format(db折扣, "0.0000")
    
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.p体检金额) = Val(vsfPrice.TextMatrix(vsfPrice.Row, mCol.p数次)) * Val(vsfPrice.TextMatrix(vsfPrice.Row, mCol.p体检单价))
    
    '更新项目
    '------------------------------------------------------------------------------------------------------------------
    dbSum = 0
    For lngLoop = 1 To vsfPrice.Rows - 1
       dbSum = dbSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.p标准金额))
    Next
    vsf.TextMatrix(vsf.Row, mCol.基本价格) = dbSum
    
    dbSum = 0
    For lngLoop = 1 To vsfPrice.Rows - 1
       dbSum = dbSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.p体检金额))
    Next
    vsf.TextMatrix(vsf.Row, mCol.体检价格) = dbSum
    
    If Val(vsf.TextMatrix(vsf.Row, mCol.基本价格)) <> 0 Then
        vsf.TextMatrix(vsf.Row, mCol.折扣) = Format(10 * Val(vsf.TextMatrix(vsf.Row, mCol.体检价格)) / Val(vsf.TextMatrix(vsf.Row, mCol.基本价格)), "0.0000")
    Else
        vsf.TextMatrix(vsf.Row, mCol.折扣) = "0.0000"
    End If
    
    '更新总体
    '------------------------------------------------------------------------------------------------------------------
    dbSum = 0
    For lngLoop = 1 To vsf.Rows - 1
       dbSum = dbSum + Val(vsf.TextMatrix(lngLoop, mCol.基本价格))
    Next
    txtSum(0).Text = Format(dbSum, "0.00")
    
    dbSum = 0
    For lngLoop = 1 To vsf.Rows - 1
       dbSum = dbSum + Val(vsf.TextMatrix(lngLoop, mCol.体检价格))
    Next
    txtSum(1).Text = Format(dbSum, "0.00")
    
    If Val(txtSum(0).Text) <> 0 Then
        txtSum(2).Text = Format(10 * Val(txtSum(1).Text) / Val(txtSum(0).Text), "0.0000")
    Else
        txtSum(2).Text = "0.0000"
    End If
        
    ChangePrice = True
    
End Function


Private Function SumPrice(ByVal bytMode As Byte) As Single
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim sglSum As Single
    
    For lngLoop = 1 To vsfPrice.Rows - 1
        If bytMode = 2 Then
            sglSum = sglSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.p体检金额))
        Else
            sglSum = sglSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.p标准单价)) * Val(vsfPrice.TextMatrix(lngLoop, mCol.p数次))
        End If
    Next
    SumPrice = sglSum
    
End Function

Private Function ClearData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------

    cbo.Clear
    Call ResetVsf(vsf)
    Call ResetVsf(vsfPrice)
    DataChange = False
    
        
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  初始化设置
    '返回:  True        初始化成功
    '       False       初始化失败
    '------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHand

    With vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "名称", 2100, 1, "...", 1
        .NewColumn "执行科室", 1080, 1, " ", 1
        
        .NewColumn "检查部位", 1800, 1, "...", 1
        .NewColumn "采集方式", 1200, 1, " ", 1
        .NewColumn "采集科室", 1080, 1, " ", 1
        
        .NewColumn "检验标本", 900, 1, " ", 1
        .NewColumn "基本价格", 900, 7
        .NewColumn "体检价格", 900, 7
        .NewColumn "折扣", 900, 7, , 1
        .NewColumn "体检类型", 0, 1
        .NewColumn "类别", 0, 1
        .NewColumn "结算方式", 900, 1, "记帐|收费", 1
        .NewColumn "执行科室id", 0, 1
        .NewColumn "采集方式id", 0, 1
        .NewColumn "采集科室id", 0, 1
        .NewColumn "检查部位id", 0, 1
        .NewColumn "计费明细", 0, 1
        .NewColumn "新加", 0, 1
        .NewColumn "前景色", 0, 1
        .NewColumn "删除", 0, 1
        .NewColumn "公共", 0, 1
        .NewColumn "清单id", 0, 1
        .FixedCols = 1
        
        .SelectMode = True
        
        .Body.ColFormat(mCol.基本价格) = "0.00"
        .Body.ColFormat(mCol.体检价格) = "0.00"
        .Body.ColFormat(mCol.折扣) = "0.0000"
    End With
    
    With vsfPrice
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "计价项目", 2100, 1, " ", 1
        .NewColumn "收费项目", 2700, 1, "...", 1
        .NewColumn "单位", 600, 1
        .NewColumn "数次", 540, 7, , 1
        .NewColumn "标准单价", 900, 7
        .NewColumn "体检单价", 900, 7, , 1
        .NewColumn "折扣", 900, 7, , 1
        .NewColumn "标准价格", 900, 7
        .NewColumn "体检价格", 900, 7
        .NewColumn "执行科室", 1080, 1, " ", 1
        .NewColumn "执行科室id", 0
        .NewColumn "收费项目id", 0
        .NewColumn "计价性质", 0
        .NewColumn "类别", 0
        .NewColumn "", 0
        .FixedCols = 1
        .Body.ColFormat(mCol.p标准单价) = "0.00000"
        .Body.ColFormat(mCol.p体检单价) = "0.00000"
        .Body.ColFormat(mCol.p标准金额) = "0.00"
        .Body.ColFormat(mCol.p体检金额) = "0.00"
        .Body.ColFormat(mCol.p折扣) = "0.0000"
        .SelectMode = True
    End With
    
    mstrGroup = ""
    cbo.AddItem "缺省"
               
    If mblnGroup = False Then
        cbo.Visible = False
        lbl(4).Visible = False
'        cmd(5).Visible = False
'        cmd(6).Visible = False
        lblTitle.Caption = "人员项目"
    Else
'        cmd(5).Visible = True
'        cmd(6).Visible = True
        lblTitle.Caption = "组别项目"
    End If
        
    InitData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ReadPrice(ByVal intRow As Integer) As Boolean
    '读取对应的计费明细
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngRow As Long
    Dim lngCol As Long
    
    Call ResetVsf(vsfPrice)
    
    If intRow = 0 Then Exit Function
    
    If vsf.TextMatrix(intRow, mCol.计费明细) <> "" Then
        
        varRow = Split(vsf.TextMatrix(intRow, mCol.计费明细), ";")
        
        vsfPrice.Rows = UBound(varRow) + 2
        
        For lngRow = 0 To UBound(varRow)
            If varRow(lngRow) <> "" Then
                varCol = Split(varRow(lngRow), ":")
                For lngCol = 0 To UBound(varCol)
                    
                    If Val(varCol(6)) = 2 Then
                        vsfPrice.TextMatrix(lngRow + 1, mCol.p计价项目) = "采集方式-" & Trim(vsf.TextMatrix(vsf.Row, mCol.采集方式))
                    ElseIf Trim(vsf.TextMatrix(vsf.Row, mCol.类别)) = "检验" Then
                        vsfPrice.TextMatrix(lngRow + 1, mCol.p计价项目) = "检验项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目))
                    Else
                        vsfPrice.TextMatrix(lngRow + 1, mCol.p计价项目) = "检查项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目))
                    End If
                    
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p名称) = varCol(0)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p计算单位) = varCol(1)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p数次) = varCol(2)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p标准单价) = varCol(3)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p体检单价) = varCol(4)
                    
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p标准金额) = Val(varCol(2)) * Val(varCol(3))
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p体检金额) = Val(varCol(2)) * Val(varCol(4))
                                        
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p收费项目id) = varCol(5)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p计价性质) = varCol(6)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p执行科室) = varCol(7)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p执行科室id) = varCol(8)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p类别) = varCol(9)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p折扣) = varCol(10)
                    
                    vsfPrice.RowData(lngRow + 1) = Val(varCol(5))
                    
                Next
            End If
        Next
        
    End If
    
    ReadPrice = True
End Function

Private Function WritePrice(ByVal intRow As Integer) As Boolean
    Dim strTmp As String
    Dim lngRow As Long
    Dim varCol As Variant
    
    On Error GoTo errHand
    
    If intRow <= 0 Then Exit Function
    
    For lngRow = 1 To vsfPrice.Rows - 1
        If Val(vsfPrice.TextMatrix(lngRow, mCol.p收费项目id)) > 0 Then
            
            varCol = Split(String(11, ":"), ":")
            
            varCol(0) = vsfPrice.TextMatrix(lngRow, mCol.p名称)
            varCol(1) = vsfPrice.TextMatrix(lngRow, mCol.p计算单位)
            varCol(2) = vsfPrice.TextMatrix(lngRow, mCol.p数次)
            varCol(3) = vsfPrice.TextMatrix(lngRow, mCol.p标准单价)
            varCol(4) = vsfPrice.TextMatrix(lngRow, mCol.p体检单价)
            varCol(5) = vsfPrice.TextMatrix(lngRow, mCol.p收费项目id)
            varCol(6) = vsfPrice.TextMatrix(lngRow, mCol.p计价性质)
            varCol(7) = vsfPrice.TextMatrix(lngRow, mCol.p执行科室)
            varCol(8) = vsfPrice.TextMatrix(lngRow, mCol.p执行科室id)
            varCol(9) = vsfPrice.TextMatrix(lngRow, mCol.p类别)
            varCol(10) = vsfPrice.TextMatrix(lngRow, mCol.p折扣)
            
            If strTmp = "" Then
                strTmp = Join(varCol, ":")
            Else
                strTmp = strTmp & ";" & Join(varCol, ":")
            End If
        End If
    Next
    
    vsf.TextMatrix(intRow, mCol.计费明细) = strTmp
    
    WritePrice = True
    
errHand:
    
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  读取数据
    '参数:  lngKey      体检类型序号
    '返回:  True        读取成功
    '       False       读取失败
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
            
    On Error GoTo errHand

    
    '读取体检组别及体检项目
    
    mblnNo = True
    
    cbo.Clear
    
    If mblnGroup = False Then
        
        gstrSQL = "Select 组别名称,性别 From 体检人员档案 Where 病人id=[2] and 登记id=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, mlngArchiveKey)
        If rs.BOF = False Then
            cbo.AddItem zlCommFun.NVL(rs("组别名称"))
            mstr性别 = zlCommFun.NVL(rs("性别"), "未知")
        End If
        
        If cbo.ListCount = 0 Then cbo.AddItem "缺省"
    Else
        gstrSQL = "SELECT A.组别名称, rownum AS ID FROM 体检组别 A WHERE A.登记id=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then
            Do While Not rs.EOF
                cbo.AddItem zlCommFun.NVL(rs("组别名称"))
                rs.MoveNext
            Loop
        Else
            cbo.AddItem "缺省"
        End If
    End If
    
    If cbo.ListIndex = -1 And cbo.ListCount > 0 Then cbo.ListIndex = 0
    '读取体检项目
    mblnNo = False
    
    Call cbo_Click
    
    

    ReadData = True
    
    Exit Function
    
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function CheckHave(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  检查是否有重复的项目
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) = lngKey And vsf.Row <> lngLoop Then
            CheckHave = True
            Exit Function
        End If
    Next
End Function

Private Function ShowOpenList(Optional strText As String, Optional ByVal lngCol As Long = 0) As Byte
    '------------------------------------------------------------------------------------------------------------------
    '功能:  以列表方式显示数据
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strClass As String
    Dim strPath As String
    Dim lngLoop As Long
    Dim strTmp As String
    On Error GoTo errHand
    
    ShowOpenList = 2

    Select Case lngCol
        Case mCol.项目
        
            strText = UCase(strText)
            
            strLvw = "编码,1200,0,1;名称,2700,0,0;单位,900,0,0;标本部位,900,0,0;类别,900,0,0"
            strPath = Me.Name & "\体检项目选择"
            
            gstrSQL = GetPublicSQL(SQL.体检项目过滤选择, strText)
            
            If ParamInfo.项目输入匹配方式 = 1 Then
                strTmp = strText & "%"
            Else
                strTmp = "%" & strText & "%"
            End If
            
            Dim bytParam1 As Byte
            Dim bytParam2 As Byte
            
            Select Case mstr性别
            Case "男"
                bytParam1 = 1
            Case "女"
                bytParam2 = 2
            End Select
            
            If Trim(vsf.TextMatrix(vsf.Row, mCol.类别)) = "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "D", strText & "%", strTmp, bytParam1, bytParam2)
            ElseIf Trim(vsf.TextMatrix(vsf.Row, mCol.类别)) = "检验" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "", strText & "%", strTmp, bytParam1, bytParam2)
            Else
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "D", "", strText & "%", strTmp, bytParam1, bytParam2)
            End If
            
'            If Trim(vsf.TextMatrix(vsf.Row, mCol.类别)) = "" Then
'                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "D", strText & "%", strTmp)
'            ElseIf Trim(vsf.TextMatrix(vsf.Row, mCol.类别)) = "检验" Then
'                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "", strText & "%", strTmp)
'            Else
'                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "D", "", strText & "%", strTmp)
'            End If

        Case mCol.执行科室
            
            strLvw = "编码,1200,0,1;名称,3300,0,0;简码,1200,0,0"
            strPath = Me.Name & "\执行科室选择"
            
            gstrSQL = GetPublicSQL(SQL.诊疗执行科室)
            If gstrSQL = "" Then Exit Function
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)), mlngDept, UserInfo.部门ID, "%" & UCase(strText) & "%")
            
        Case mCol.检查部位
            
            strText = "'%" & UCase(strText) & "%'"
            
            strLvw = "名称,3300,0,0"
            strPath = Me.Name & "\检查部位选择"
            
            gstrSQL = "select B.标本部位 AS 名称,B.ID,0 AS 选择 from 诊疗项目组合 A,诊疗项目目录 B WHERE (B.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or B.撤档时间 is NULL) AND A.诊疗项目ID=B.ID AND A.诊疗组合ID=" & Val(vsf.RowData(vsf.Row)) & ""
            
            rs.CursorLocation = adUseClient
            If rs.State = adStateOpen Then rs.Close
            rs.Open gstrSQL, gcnOracle, adOpenStatic, adLockOptimistic
            
        Case mCol.采集方式
            
            strText = "%" & UCase(strText) & "%"
            
            strLvw = "编码,1200,0,1;名称,3300,0,0"
            strPath = Me.Name & "\采集方式选择"
            
            gstrSQL = "SELECT A.ID,A.编码,A.名称 " & _
                "FROM 诊疗项目目录 A,诊疗用法用量 B " & _
                "WHERE (A.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or A.撤档时间 is NULL) AND A.类别='E' AND A.操作类型='6' AND A.ID=B.用法id AND B.项目id=[1] "
            gstrSQL = gstrSQL & " AND (UPPER(A.编码) Like [2] OR A.名称 Like [2] OR A.ID IN (SELECT 诊疗项目id FROM 诊疗项目别名 WHERE (名称 Like [2] OR UPPER(简码) Like [2])))"
            
'            Call OpenRecord(rs, gstrSQL, Me.Caption)
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)), strText)
            
            If rs.BOF Then
                gstrSQL = "SELECT A.ID,A.编码,A.名称 " & _
                    "FROM 诊疗项目目录 A " & _
                    "WHERE (A.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or A.撤档时间 is NULL) AND A.类别='E' AND A.操作类型='6' "
                gstrSQL = gstrSQL & " AND (UPPER(A.编码) Like [1] OR A.名称 Like [1] OR A.ID IN (SELECT 诊疗项目id FROM 诊疗项目别名 WHERE (名称 Like [1] OR UPPER(简码) Like [1])))"
                    
            End If
'            Call OpenRecord(rs, gstrSQL, Me.Caption)
            
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText)
            
    End Select

    If rs.BOF Then
        ShowOpenList = 0
        Exit Function
    End If
    If rs.RecordCount = 1 And strText <> "'%%'" Then GoTo PointOver
    Call CalcPosition(sglX, sglY, vsf)
    
    If lngCol = mCol.检查部位 Then
        If vsf.TextMatrix(vsf.Row, mCol.检查部位id) <> "" Then
            Do While Not rs.EOF
                If InStr("," & vsf.TextMatrix(vsf.Row, mCol.检查部位id) & ",", "," & rs("ID").Value & ",") > 0 Then rs("选择").Value = 1
                rs.MoveNext
            Loop
        End If
        rs.MoveFirst
        
        If frmSelectDialog.ShowSelect(Me, 2, rs, strLvw, "请从下面选择多个项目,然后回车或双击退出", sglX + 60, sglY + 30, 9000, 4500, 300, , strPath, , False, True) Then GoTo PointOver
        
    Else
        If frmSelectDialog.ShowSelect(Me, 2, rs, strLvw, "请从下面选择一个项目", sglX + 60, sglY + 30, 9000, 4500, 300, , strPath, , False) Then GoTo PointOver
    End If
        
    Exit Function
    
PointOver:
    Select Case lngCol
        Case mCol.项目
            If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                MsgBox "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已经被选择！", vbInformation, gstrSysName
                Exit Function
            End If
            
            vsf.Cell(flexcpText, vsf.Row, mCol.项目 + 1, vsf.Row, vsf.Cols - 1) = ""
            
            vsf.EditText = zlCommFun.NVL(rs("名称").Value)
            vsf.TextMatrix(vsf.Row, mCol.类别) = zlCommFun.NVL(rs("类别").Value)
            vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
            vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
            vsf.RowData(vsf.Row) = zlCommFun.NVL(rs("ID").Value)
            
        Case mCol.执行科室
        
            vsf.EditText = zlCommFun.NVL(rs("名称").Value)
            vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
            vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
            vsf.TextMatrix(vsf.Row, mCol.执行科室id) = zlCommFun.NVL(rs("ID").Value)
        
        Case mCol.检查部位
            
            vsf.TextMatrix(vsf.Row, vsf.Col) = ""
            vsf.TextMatrix(vsf.Row, mCol.检查部位id) = ""
            
            rs.Filter = ""
            rs.Filter = "选择=1"
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                Do While Not rs.EOF
                    vsf.TextMatrix(vsf.Row, vsf.Col) = vsf.TextMatrix(vsf.Row, vsf.Col) & zlCommFun.NVL(rs("名称").Value) & ","
                    vsf.TextMatrix(vsf.Row, mCol.检查部位id) = vsf.TextMatrix(vsf.Row, mCol.检查部位id) & zlCommFun.NVL(rs("ID").Value) & ","
                    rs.MoveNext
                Loop
                
                If vsf.TextMatrix(vsf.Row, mCol.检查部位) <> "" Then vsf.TextMatrix(vsf.Row, mCol.检查部位) = Mid(vsf.TextMatrix(vsf.Row, mCol.检查部位), 1, Len(vsf.TextMatrix(vsf.Row, mCol.检查部位)) - 1)
                If vsf.TextMatrix(vsf.Row, mCol.检查部位id) <> "" Then vsf.TextMatrix(vsf.Row, mCol.检查部位id) = Mid(vsf.TextMatrix(vsf.Row, mCol.检查部位id), 1, Len(vsf.TextMatrix(vsf.Row, mCol.检查部位id)) - 1)
                
            End If
        Case mCol.采集方式
        
            vsf.EditText = zlCommFun.NVL(rs("名称").Value)
            vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
            vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
            vsf.TextMatrix(vsf.Row, mCol.采集方式id) = zlCommFun.NVL(rs("ID").Value)
    End Select
    
    ShowOpenList = 1
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ReadSample(ByVal lng诊疗项目id As Long) As String
    '------------------------------------------------------------------------------------------------------------------
    '功能:获取可选的采集方式下拉数据
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    gstrSQL = "SELECT 1 FROM 诊疗项目目录 WHERE 组合项目=1 AND ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng诊疗项目id)
    If rs.BOF = False Then
        '是组合项目
        
        gstrSQL = "SELECT DISTINCT A.标本类型 AS 名称 FROM 检验项目参考 A,检验报告项目 B,诊疗项目目录 C " & _
                "WHERE C.ID<>[1] AND nvl(C.组合项目,0)=0 " & _
                    "AND B.报告项目id=A.项目id "
                    
        gstrSQL = gstrSQL & "AND B.诊疗项目id IN (SELECT C.ID " & _
                     "FROM 检验报告项目 A," & _
                          "(SELECT 报告项目id FROM 检验报告项目 WHERE 诊疗项目id = [1]) B," & _
                          "诊疗项目目录 C,诊治所见项目 D,检验项目 E,检验报告项目 F " & _
                    "WHERE A.报告项目id = B.报告项目id AND A.诊疗项目id <> [1] AND " & _
                          "nvl(C.组合项目,0) = 0 AND A.诊疗项目id = C.ID AND C.ID=F.诊疗项目id AND F.报告项目id=D.ID AND D.ID=E.诊治项目id)"
                                  
    Else
        gstrSQL = "SELECT A.标本类型 AS 名称 FROM 检验项目参考 A,检验报告项目 B,诊疗项目目录 C " & _
                "WHERE C.ID=[1] AND nvl(C.组合项目,0)=0 AND B.诊疗项目id=[1] and B.报告项目id=A.项目id"
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng诊疗项目id)

    If rs.BOF = False Then
        Do While Not rs.EOF
            ReadSample = ReadSample & rs("名称").Value & "|"
            rs.MoveNext
        Loop
    Else
        
        '没有对应时，读取所有标本类型
        gstrSQL = "SELECT 名称 FROM 诊疗检验标本 A "
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If rs.BOF = False Then
            Do While Not rs.EOF
                ReadSample = ReadSample & rs("名称").Value & "|"
                rs.MoveNext
            Loop
        End If
        
    End If
    
    Exit Function
        
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function SetRowData(ByVal lngKey As Long, ByVal intRow As Integer, ParamArray arryMode() As Variant) As Boolean
'------------------------------------------------------------------------------------------------------------------
    '功能:设置行数据（随行不同而不同）
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim strCombList As String
    
    On Error Resume Next
    
    For lngLoop = 0 To UBound(arryMode)
        Select Case arryMode(lngLoop)
        Case "收费执行科室"
        
            If InStr("4,5,6,7", vsfPrice.TextMatrix(intRow, mCol.p类别)) > 0 Then
                gstrSQL = GetPublicSQL(SQL.药品执行科室)
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, vsfPrice.TextMatrix(intRow, mCol.p类别))
            Else
                gstrSQL = GetPublicSQL(SQL.收费执行科室, "1")
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, mlngDept, UserInfo.部门ID, "%%")
            End If
            If rs.RecordCount > 1 Then
                vsfPrice.EditMode(mCol.p执行科室) = 1
                vsfPrice.Body.ColComboList(mCol.p执行科室) = vsfPrice.Body.BuildComboList(rs, "名称", "ID")
            Else
                vsfPrice.EditMode(mCol.p执行科室) = 0
                vsfPrice.Body.ColComboList(mCol.p执行科室) = ""
            End If
        
        Case "计价项目"
            
            If Trim(vsf.TextMatrix(intRow, mCol.类别)) = "检查" Then
                strCombList = "检查项目-" & Trim(vsf.TextMatrix(intRow, mCol.项目))
                vsfPrice.EditMode(mCol.p计价项目) = 0
                vsfPrice.Body.ColComboList(mCol.p计价项目) = ""
                
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p计价项目) = strCombList
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p计价性质) = "1"
            Else
                strCombList = "检验项目-" & Trim(vsf.TextMatrix(intRow, mCol.项目))
                If Val(vsf.TextMatrix(intRow, mCol.采集方式id)) > 0 Then
                    strCombList = strCombList & "|采集方式-" & Trim(vsf.TextMatrix(intRow, mCol.采集方式))
                    vsfPrice.EditMode(mCol.p计价项目) = 1
                    vsfPrice.Body.ColComboList(mCol.p计价项目) = strCombList
                Else
                    vsfPrice.EditMode(mCol.p计价项目) = 0
                    vsfPrice.Body.ColComboList(mCol.p计价项目) = ""
                End If
            End If
            
        Case "诊疗执行科室"
        
            gstrSQL = GetPublicSQL(SQL.诊疗执行科室, "1")
            If gstrSQL <> "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, mlngDept, UserInfo.部门ID, "%%")
                If rs.RecordCount > 1 Then
                    vsf.EditMode(mCol.执行科室) = 1
                    vsf.Body.ColComboList(mCol.执行科室) = vsf.Body.BuildComboList(rs, "名称", "ID")
                Else
                    vsf.EditMode(mCol.执行科室) = 0
                    vsf.Body.ColComboList(mCol.执行科室) = ""
                End If
            End If
        
        Case "采集方式"
        
            gstrSQL = "SELECT A.名称 AS 名称,A.ID FROM 诊疗项目目录 A,诊疗用法用量 B WHERE A.ID=B.用法id AND A.类别='E' AND A.操作类型='6' AND B.项目ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.RecordCount > 1 Then
                vsf.EditMode(mCol.采集方式) = 1
                vsf.Body.ColComboList(mCol.采集方式) = vsf.Body.BuildComboList(rs, "名称", "ID")
            Else
                gstrSQL = "SELECT A.名称 AS 名称,A.ID FROM 诊疗项目目录 A WHERE A.类别='E' AND A.操作类型='6'"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                If rs.RecordCount > 1 Then
                    vsf.EditMode(mCol.采集方式) = 1
                    vsf.Body.ColComboList(mCol.采集方式) = vsf.Body.BuildComboList(rs, "名称", "ID")
                Else
                    vsf.EditMode(mCol.采集方式) = 0
                    vsf.Body.ColComboList(mCol.采集方式) = ""
                End If
            End If
            
        Case "采集科室"
        
            gstrSQL = GetPublicSQL(SQL.诊疗执行科室)
            If gstrSQL <> "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.TextMatrix(intRow, mCol.采集方式id)), mlngDept, UserInfo.部门ID, "%%")
                If rs.RecordCount > 1 Then
                    vsf.EditMode(mCol.采集科室) = 1
                    vsf.Body.ColComboList(mCol.采集科室) = vsf.Body.BuildComboList(rs, "*名称", "ID")
                Else
                    vsf.EditMode(mCol.采集科室) = 0
                    vsf.Body.ColComboList(mCol.采集科室) = ""
                End If
            End If
        
        Case "检验标本"
        
            gstrSQL = "SELECT 1 FROM 诊疗项目目录 WHERE 组合项目=1 AND ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                '是组合项目
                
                gstrSQL = "SELECT DISTINCT A.标本类型 AS 名称 FROM 检验项目参考 A,检验报告项目 B,诊疗项目目录 C " & _
                        "WHERE C.ID<>[1] AND nvl(C.组合项目,0)=0 " & _
                            "AND B.报告项目id=A.项目id and rownum<2"
                            
                gstrSQL = gstrSQL & "AND B.诊疗项目id IN (SELECT C.ID " & _
                             "FROM 检验报告项目 A," & _
                                  "(SELECT 报告项目id FROM 检验报告项目 WHERE 诊疗项目id = [1]) B," & _
                                  "诊疗项目目录 C,诊治所见项目 D,检验项目 E,检验报告项目 F " & _
                            "WHERE A.报告项目id = B.报告项目id AND A.诊疗项目id <> [1] AND " & _
                                  "nvl(C.组合项目,0) = 0 AND A.诊疗项目id = C.ID AND C.ID=F.诊疗项目id AND F.报告项目id=D.ID AND D.ID=E.诊治项目id)  and rownum<2 "
                                          
            Else
                gstrSQL = "SELECT A.标本类型 AS 名称 FROM 检验项目参考 A,检验报告项目 B,诊疗项目目录 C " & _
                        "WHERE C.ID=[1] AND nvl(C.组合项目,0)=0 AND B.诊疗项目id=[1] and B.报告项目id=A.项目id  and rownum<2"
            End If
        
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.RecordCount > 1 Then
                
                vsf.EditMode(mCol.检验标本) = 1
                vsf.Body.ColComboList(mCol.检验标本) = vsf.Body.BuildComboList(rs, "名称", "名称")
                
            Else
                
                '没有对应时，读取所有标本类型
                gstrSQL = "SELECT 名称 FROM 诊疗检验标本 A"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                If rs.RecordCount > 1 Then
                
                    vsf.EditMode(mCol.检验标本) = 1
                    vsf.Body.ColComboList(mCol.检验标本) = vsf.Body.BuildComboList(rs, "名称", "名称")
                Else
                    vsf.EditMode(mCol.检验标本) = 0
                    vsf.Body.ColComboList(mCol.检验标本) = ""
                End If
                
            End If
        
        End Select
    Next
    
    SetRowData = True
    
End Function

Private Function SetRowDefault(ByVal lngKey As Long, ByVal intRow As Integer, ParamArray arryMode() As Variant) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:获取缺省
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim strCombList As String
    
    On Error GoTo errHand
    
    For lngLoop = 0 To UBound(arryMode)
        
        Select Case arryMode(lngLoop)
        Case "结算方式"
            
            If mblnGroup Then
                vsf.TextMatrix(vsf.Row, mCol.结算方式) = "记帐"
            Else
                vsf.TextMatrix(vsf.Row, mCol.结算方式) = "收费"
            End If
            
        Case "执行科室"
'            lng开单科室id = mlngDept
            
            gstrSQL = GetPublicSQL(SQL.诊疗执行科室)
            If gstrSQL <> "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, mlngDept, UserInfo.部门ID, "%%")
                If rs.BOF = False Then
                    vsf.TextMatrix(vsf.Row, mCol.执行科室) = zlCommFun.NVL(rs("名称").Value)
                    vsf.TextMatrix(vsf.Row, mCol.执行科室id) = zlCommFun.NVL(rs("ID").Value)
                Else
                    vsf.TextMatrix(vsf.Row, mCol.执行科室) = gstrDeptName
                    vsf.TextMatrix(vsf.Row, mCol.执行科室id) = UserInfo.部门ID
                End If
            End If
        
        Case "采集方式"
            
            gstrSQL = "SELECT A.名称 AS 名称,A.ID FROM 诊疗项目目录 A,诊疗用法用量 B WHERE A.ID=B.用法id AND A.类别='E' AND A.操作类型='6' AND B.项目ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                vsf.TextMatrix(vsf.Row, mCol.采集方式) = zlCommFun.NVL(rs("名称").Value)
                vsf.TextMatrix(vsf.Row, mCol.采集方式id) = zlCommFun.NVL(rs("ID").Value)
            Else
                gstrSQL = "SELECT A.名称 AS 名称,A.ID FROM 诊疗项目目录 A WHERE A.类别='E' AND A.操作类型='6'"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                If rs.BOF = False Then
                    vsf.TextMatrix(vsf.Row, mCol.采集方式) = zlCommFun.NVL(rs("名称").Value)
                    vsf.TextMatrix(vsf.Row, mCol.采集方式id) = zlCommFun.NVL(rs("ID").Value)
                End If
            End If
            
        Case "采集科室"
                    
            gstrSQL = GetPublicSQL(SQL.诊疗执行科室)
            If gstrSQL <> "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.TextMatrix(vsf.Row, mCol.采集方式id)), mlngDept, UserInfo.部门ID, "%%")
                If rs.BOF = False Then
                    vsf.TextMatrix(vsf.Row, mCol.采集科室) = zlCommFun.NVL(rs("名称").Value)
                    vsf.TextMatrix(vsf.Row, mCol.采集科室id) = zlCommFun.NVL(rs("ID").Value)
                End If
            End If
        
        Case "检验标本"
        
            gstrSQL = "SELECT 1 FROM 诊疗项目目录 WHERE 组合项目=1 AND ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                '是组合项目
                
                gstrSQL = "SELECT DISTINCT A.标本类型 AS 名称 FROM 检验项目参考 A,检验报告项目 B,诊疗项目目录 C " & _
                        "WHERE C.ID<>[1] AND nvl(C.组合项目,0)=0 " & _
                            "AND B.报告项目id=A.项目id and rownum<2"
                            
                gstrSQL = gstrSQL & "AND B.诊疗项目id IN (SELECT C.ID " & _
                             "FROM 检验报告项目 A," & _
                                  "(SELECT 报告项目id FROM 检验报告项目 WHERE 诊疗项目id = [1]) B," & _
                                  "诊疗项目目录 C,诊治所见项目 D,检验项目 E,检验报告项目 F " & _
                            "WHERE A.报告项目id = B.报告项目id AND A.诊疗项目id <> [1] AND " & _
                                  "nvl(C.组合项目,0) = 0 AND A.诊疗项目id = C.ID AND C.ID=F.诊疗项目id AND F.报告项目id=D.ID AND D.ID=E.诊治项目id)  and rownum<2 "
                                          
            Else
                gstrSQL = "SELECT A.标本类型 AS 名称 FROM 检验项目参考 A,检验报告项目 B,诊疗项目目录 C " & _
                        "WHERE C.ID=[1] AND nvl(C.组合项目,0)=0 AND B.诊疗项目id=[1] and B.报告项目id=A.项目id  and rownum<2"
            End If
        
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                vsf.TextMatrix(vsf.Row, mCol.检验标本) = rs("名称").Value
            Else
                
                '没有对应时，读取所有标本类型
                gstrSQL = "SELECT 名称 FROM 诊疗检验标本 A where rownum<2"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                If rs.BOF = False Then
                    vsf.TextMatrix(vsf.Row, mCol.检验标本) = rs("名称").Value
                End If
                
            End If
        
        Case "收费执行科室"
            If InStr("4,5,6,7", vsfPrice.TextMatrix(intRow, mCol.p类别)) > 0 Then
                gstrSQL = GetPublicSQL(SQL.药品执行科室)
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, vsfPrice.TextMatrix(intRow, mCol.p类别))
            Else
                gstrSQL = GetPublicSQL(SQL.收费执行科室)
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, mlngDept, UserInfo.部门ID, "%%")
            End If
            If rs.BOF = False Then
                vsfPrice.TextMatrix(intRow, mCol.p执行科室) = zlCommFun.NVL(rs("名称").Value)
                vsfPrice.TextMatrix(intRow, mCol.p执行科室id) = zlCommFun.NVL(rs("ID").Value)
            Else
                vsfPrice.TextMatrix(intRow, mCol.p执行科室) = vsf.TextMatrix(vsf.Row, mCol.执行科室)
                vsfPrice.TextMatrix(intRow, mCol.p执行科室id) = vsf.TextMatrix(vsf.Row, mCol.执行科室id)
            End If
        Case "计价项目"
        
            If Trim(vsf.TextMatrix(vsf.Row, mCol.类别)) = "检查" Then
                strCombList = "检查项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目))
                vsfPrice.EditMode(mCol.p计价项目) = 0
                vsfPrice.Body.ColComboList(mCol.p计价项目) = ""
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p计价项目) = strCombList
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p计价性质) = "1"
            Else
                strCombList = "检验项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目))
                If Val(vsf.TextMatrix(vsf.Row, mCol.采集方式id)) > 0 Then
                    strCombList = strCombList & "|采集方式-" & Trim(vsf.TextMatrix(vsf.Row, mCol.采集方式))
                    vsfPrice.EditMode(mCol.p计价项目) = 1
                    vsfPrice.Body.ColComboList(mCol.p计价项目) = strCombList
                Else
                    vsfPrice.EditMode(mCol.p计价项目) = 0
                    vsfPrice.Body.ColComboList(mCol.p计价项目) = ""
                End If
            End If

        End Select
    Next
    
    SetRowDefault = True
    
    Exit Function
    
errHand:
    
End Function

Private Function SaveItems(ByVal strGroup As String) As Boolean
    
    Dim lngLoop As Long
    
    On Error GoTo errHand

    '保存所选择的检验项目
    mrsItems.Filter = ""
    mrsItems.Filter = "组别='" & strGroup & "' AND 删除<>'1'"
    
    Call DeleteRecord(mrsItems)
    
    For lngLoop = 1 To vsf.Rows - 1
        
        If Val(vsf.RowData(lngLoop)) > 0 Then
            mrsItems.AddNew
            
            mrsItems("组别").Value = strGroup
            mrsItems("ID").Value = vsf.RowData(lngLoop)
            mrsItems("类别").Value = vsf.TextMatrix(lngLoop, mCol.类别)
            mrsItems("名称").Value = vsf.TextMatrix(lngLoop, mCol.项目)
            mrsItems("执行科室").Value = vsf.TextMatrix(lngLoop, mCol.执行科室)
            mrsItems("检查部位").Value = vsf.TextMatrix(lngLoop, mCol.检查部位)
            mrsItems("采集方式").Value = vsf.TextMatrix(lngLoop, mCol.采集方式)
            mrsItems("采集科室").Value = vsf.TextMatrix(lngLoop, mCol.采集科室)
            mrsItems("检验标本").Value = vsf.TextMatrix(lngLoop, mCol.检验标本)
            mrsItems("体检类型").Value = vsf.TextMatrix(lngLoop, mCol.体检类型)
            mrsItems("基本价格").Value = vsf.TextMatrix(lngLoop, mCol.基本价格)
            mrsItems("体检价格").Value = vsf.TextMatrix(lngLoop, mCol.体检价格)
            mrsItems("折扣").Value = vsf.TextMatrix(lngLoop, mCol.折扣)
            mrsItems("结算方式").Value = vsf.TextMatrix(lngLoop, mCol.结算方式)
            mrsItems("执行科室id").Value = vsf.TextMatrix(lngLoop, mCol.执行科室id)
            mrsItems("采集方式id").Value = vsf.TextMatrix(lngLoop, mCol.采集方式id)
            mrsItems("采集科室id").Value = vsf.TextMatrix(lngLoop, mCol.采集科室id)
            mrsItems("检查部位id").Value = vsf.TextMatrix(lngLoop, mCol.检查部位id)
            mrsItems("计费明细").Value = vsf.TextMatrix(lngLoop, mCol.计费明细)
            mrsItems("新加").Value = vsf.TextMatrix(lngLoop, mCol.新加)
            mrsItems("前景色").Value = vsf.TextMatrix(lngLoop, mCol.前景色)
            mrsItems("删除").Value = ""
            mrsItems("公共").Value = vsf.TextMatrix(lngLoop, mCol.公共)
            mrsItems("清单id").Value = vsf.TextMatrix(lngLoop, mCol.清单id)
            
        End If
    Next
    
    SaveItems = True
    
errHand:

End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  校验数据的有效性
    '返回:  True        数据有效
    '       False       数据无效
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
                
   
    If Val(vsf.TextMatrix(lngLoop, mCol.体检价格)) < 0 Then
            
            ShowSimpleMsg "体检价格不能为负数"
            vsf.Row = lngLoop
            vsf.Col = mCol.体检价格
            vsf.ShowCell vsf.Row, vsf.Col
            vsf.SetFocus
            
            Exit Function
        End If
        
'        If Val(vsf.TextMatrix(lngLoop, mCol.体检价格)) > Val(vsf.TextMatrix(lngLoop, mCol.基本价格)) Then
'
'            ShowSimpleMsg "体检价格不能大于基本价格"
'            vsf.Row = lngLoop
'            vsf.Col = mCol.体检价格
'            vsf.ShowCell vsf.Row, vsf.Col
'            vsf.SetFocus
'
'            Exit Function
'        End If
    
    ValidEdit = True
    
End Function

Private Function ReadItems(ByVal strGroup As String) As Boolean
    
    mrsItems.Filter = ""
    mrsItems.Filter = "组别='" & strGroup & "' AND 删除<>'1'"
    If mrsItems.RecordCount > 0 Then
        mrsItems.MoveFirst
        Call FillGrid(vsf, mrsItems)
    End If
    
    ReadItems = True
    
End Function

Private Function ReadTemplate(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim rsPrice As New ADODB.Recordset
    
    Dim strKeys As String
    Dim bytParam1 As Byte
    Dim bytParam2 As Byte
    
    bytParam1 = 1
    bytParam2 = 2
            
    Select Case mstr性别
    Case "男"
        bytParam1 = 1
    Case "女"
        bytParam2 = 2
    End Select
        
    On Error GoTo errHand
    
    gstrSQL = "SELECT DISTINCT A.ID,DECODE(A.类别,'C','检验','D','检查') AS 类别,A.编码,A.名称,C.名称 AS 体检类型,D.名称 As 采集方式,B.采集方式id,B.检验标本,B.检查部位,B.检查部位id " & _
                "FROM 诊疗项目目录 A,体检类型目录 B,体检类型 C,诊疗项目目录 D " & _
                "WHERE A.ID=B.诊疗项目ID AND C.序号=B.序号 AND D.ID(+)=B.采集方式id AND B.序号=[1] And Nvl(a.适用性别,0) In (0,[2],[3])"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, bytParam1, bytParam2)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            vsf.Row = vsf.Rows - 1
            If Val(vsf.RowData(vsf.Row)) > 0 Then
                vsf.Rows = vsf.Rows + 1
                vsf.Row = vsf.Rows - 1
            End If
            
            If CheckHave(rs("ID").Value) = False Then
            
                vsf.TextMatrix(vsf.Row, mCol.类别) = zlCommFun.NVL(rs("类别").Value)
                vsf.TextMatrix(vsf.Row, mCol.项目) = zlCommFun.NVL(rs("名称").Value)
                vsf.TextMatrix(vsf.Row, mCol.体检类型) = zlCommFun.NVL(rs("体检类型").Value)
                
                vsf.TextMatrix(vsf.Row, mCol.检验标本) = zlCommFun.NVL(rs("检验标本").Value)
                vsf.TextMatrix(vsf.Row, mCol.检查部位) = zlCommFun.NVL(rs("检查部位").Value)
                vsf.TextMatrix(vsf.Row, mCol.采集方式) = zlCommFun.NVL(rs("采集方式").Value)
                vsf.TextMatrix(vsf.Row, mCol.采集方式id) = zlCommFun.NVL(rs("采集方式id").Value)
                vsf.TextMatrix(vsf.Row, mCol.检查部位id) = zlCommFun.NVL(rs("检查部位id").Value)
                
                vsf.RowData(vsf.Row) = zlCommFun.NVL(rs("ID").Value)
            End If
                        
            If vsf.TextMatrix(vsf.Row, mCol.类别) = "检验" Then
                Call SetRowDefault(Val(vsf.RowData(vsf.Row)), vsf.Row, "执行科室", "采集方式", "采集科室", "检验标本", "结算方式", "计价项目")
                
            ElseIf vsf.TextMatrix(vsf.Row, mCol.类别) = "检查" Then
                Call SetRowDefault(Val(vsf.RowData(vsf.Row)), vsf.Row, "执行科室", "结算方式", "计价项目")
            End If
            
            gstrSQL = "Select y.类别,z.数次,y.名称,y.计算单位,x.现价,y.id,Nvl(z.计价性质,1) As 计价性质 " & _
                        "From " & _
                            "( Select a.序号,a.诊疗项目id,a.收费细目id,Sum(c.现价) As 现价 " & _
                              "From 收费价目 c, " & _
                                   "体检类型计价 a " & _
                              "Where a.收费细目id = c.收费细目id " & _
                                    "and c.执行日期<=SYSDATE and (c.终止日期 IS NULL OR c.终止日期>SYSDATE) " & _
                                    "and A.序号=[2] " & _
                                    "and A.诊疗项目id=[1] " & _
                              "Group by a.序号,a.诊疗项目id,a.收费细目id " & _
                            ") x, " & _
                            "收费项目目录 y, " & _
                            "体检类型计价 z " & _
                        "Where x.收费细目id = y.ID " & _
                              "and z.序号=x.序号 " & _
                              "and z.诊疗项目id=x.诊疗项目id " & _
                              "and z.收费细目id=x.收费细目id "
                        
            Set rsPrice = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)), lngKey)
            If rsPrice.BOF = False Then
                With vsfPrice
                    Do While Not rsPrice.EOF
                        
                        If Val(.TextMatrix(.Rows - 1, mCol.p收费项目id)) > 0 Then
                            .Rows = .Rows + 1
                        End If
                        
                        .TextMatrix(.Rows - 1, mCol.p名称) = zlCommFun.NVL(rsPrice("名称"))
                        .TextMatrix(.Rows - 1, mCol.p计算单位) = zlCommFun.NVL(rsPrice("计算单位"))
                        .TextMatrix(.Rows - 1, mCol.p数次) = zlCommFun.NVL(rsPrice("数次"))
                        .TextMatrix(.Rows - 1, mCol.p标准单价) = zlCommFun.NVL(rsPrice("现价"))
                        .TextMatrix(.Rows - 1, mCol.p体检单价) = zlCommFun.NVL(rsPrice("现价"))
                        .TextMatrix(.Rows - 1, mCol.p标准金额) = zlCommFun.NVL(rsPrice("数次"), 0) * zlCommFun.NVL(rsPrice("现价"), 0)
                        .TextMatrix(.Rows - 1, mCol.p体检金额) = zlCommFun.NVL(rsPrice("数次"), 0) * zlCommFun.NVL(rsPrice("现价"), 0)
                        .TextMatrix(.Rows - 1, mCol.p收费项目id) = zlCommFun.NVL(rsPrice("ID"))
                        .TextMatrix(.Rows - 1, mCol.p计价性质) = zlCommFun.NVL(rsPrice("计价性质"))
                        .TextMatrix(.Rows - 1, mCol.p类别) = zlCommFun.NVL(rsPrice("类别"))
                        .RowData(.Rows - 1) = zlCommFun.NVL(rsPrice("ID"), 0)
                        
                        If zlCommFun.NVL(rsPrice("计价性质"), 1) = 2 Then
                            .TextMatrix(.Rows - 1, mCol.p计价项目) = "采集方式-" & Trim(vsf.TextMatrix(vsf.Row, mCol.采集方式))
                        ElseIf Trim(vsf.TextMatrix(vsf.Row, mCol.类别)) = "检验" Then
                            .TextMatrix(.Rows - 1, mCol.p计价项目) = "检验项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目))
                        Else
                            .TextMatrix(.Rows - 1, mCol.p计价项目) = "检查项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目))
                        End If
                        
                        Call SetRowDefault(Val(.RowData(.Rows - 1)), .Rows - 1, "收费执行科室")
                        
                        If InStr("567", .TextMatrix(.Rows - 1, mCol.p类别)) > 0 Then
                            .TextMatrix(.Rows - 1, mCol.p可用库存) = GetStorage(Val(.RowData(.Rows - 1)), Val(.TextMatrix(.Rows - 1, mCol.p执行科室id)))
                            Call PromptStorageWarn(Val(.TextMatrix(.Rows - 1, mCol.p数次)), Val(.TextMatrix(.Rows - 1, mCol.p可用库存)), .TextMatrix(.Rows - 1, mCol.p名称), .TextMatrix(.Rows - 1, mCol.p执行科室), .TextMatrix(.Rows - 1, mCol.p计算单位), 1)
                        End If
                        rsPrice.MoveNext
                    Loop
                End With
                
                vsf.TextMatrix(vsf.Row, mCol.基本价格) = SumPrice(1)
                vsf.TextMatrix(vsf.Row, mCol.体检价格) = SumPrice(2)
                
            End If
            
            Call vsf_BeforeRowColChange(0, 0, vsf.Row, vsf.Col, False)
            Call vsfPrice_BeforeRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col, False)
            Call WritePrice(vsf.Row)
                                    
            rs.MoveNext
        Loop
    End If
    
    ReadTemplate = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub cbo_Click()
    If mblnNo Then Exit Sub
    
    If mstrGroup <> cbo.Text Then
        
        Call WritePrice(vsf.Row)
        Call SaveItems(mstrGroup)
        
        mstrGroup = cbo.Text
        
        Call ResetVsf(vsf)
        Call ResetVsf(vsfPrice)
        
        Call ReadItems(mstrGroup)
        Call ReadPrice(vsf.Row)
        
        Call ChangeItem(Val(vsf.TextMatrix(vsf.Row, mCol.基本价格)), Val(vsf.TextMatrix(vsf.Row, mCol.体检价格)), 1, False)

    End If
       
End Sub

Private Sub cbo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then

        zlCommFun.PressKey vbKeyTab

    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim lngKey As Long
    Dim strDate As String
    Dim objPoint As POINTAPI
    Dim strTmp As String
    Dim rsData As New ADODB.Recordset
    Dim rsPrice As New ADODB.Recordset

    Dim lngLoop As Long
    Dim objItem As ListItem
    Dim intRow As Long
    Dim strKeys As String

    On Error GoTo errHand
    
    Call ClientToScreen(cmd(Index).hWnd, objPoint)
    
    Select Case Index

    Case 5
            
        gstrSQL = GetPublicSQL(SQL.体检项目选择)

        Select Case mstr性别
        Case "男"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 1, 1)
        Case "女"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 2, 2)
        Case Else
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 1, 2)
        End Select
        
        If ShowTxtSelect(Me, cmd(Index), "编码,1200,0,1;名称,2700,0,0;单位,900,0,0;类别,900,0,0", Me.Name & "\体检项目选择", "请从列表中选择一个体检项目。", rsData, rs, 8790, 5100, True) Then

            rs.Filter = 0
            rs.Filter = "选择=1"
            If rs.RecordCount > 0 Then

                rs.MoveFirst
                Do While Not rs.EOF
                    '选取了一个项目
                    vsf.Row = 0

                    If CheckHave(zlCommFun.NVL(rs("ID").Value)) = False Then

                        If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                            vsf.Rows = vsf.Rows + 1
                        End If
                        intRow = vsf.Rows - 1
                        vsf.Row = vsf.Rows - 1

                        vsf.Cell(flexcpText, intRow, mCol.项目 + 1, intRow, vsf.Cols - 1) = ""

                        vsf.TextMatrix(intRow, mCol.类别) = zlCommFun.NVL(rs("类别").Value)
                        vsf.TextMatrix(intRow, mCol.项目) = zlCommFun.NVL(rs("名称").Value)
                        vsf.RowData(intRow) = zlCommFun.NVL(rs("ID").Value)

'                        Call DefaultValue(Val(vsf.RowData(intRow)), 1)
'                        If vsf.TextMatrix(intRow, mCol.类别) = "检验" Then
'                            Call DefaultValue(Val(vsf.RowData(intRow)), 2)
'                            Call DefaultValue(Val(vsf.RowData(intRow)), 3)
'                        End If
                        
                        Call CreatePriceList(intRow)
                        Call WritePrice(intRow)

                        DataChange = True
                    End If

                    rs.MoveNext
                Loop
            End If

        End If

        EnterFocus vsf

    Case 6
    
        gstrSQL = GetPublicSQL(SQL.体检类型分类选择)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, IIf(mblnGroup, 2, 1))
        If ShowTxtSelect(Me, cmd(Index), "编码,1080,0,1;名称,2400,0,0;简码,900,0,0;说明,1500,0,0", Me.Name & "\体检类型选择", "请从列表中选择一个体检类型。", rsData, rs, 8790, 5100, True) Then

            rs.Filter = 0
            rs.Filter = "选择=1"
            If rs.RecordCount > 0 Then

                If Val(vsf.RowData(1)) > 0 Then
                    
                    If Not (mblnGroup = False And mbytMode = 2) Then
                        If MsgBox("是否要清除已选择的体检项目？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        
                            '记录
                            If Trim(cbo.Text) <> "" Then
                                For lngLoop = 1 To vsf.Rows - 1
                                    mrsItems.Filter = ""
                                    mrsItems.Filter = "组别='" & Trim(cbo.Text) & "' AND 清单id=" & Val(vsf.TextMatrix(lngLoop, mCol.清单id))
                                    If mrsItems.RecordCount > 0 Then
                                        mrsItems.MoveFirst
                                        mrsItems("删除").Value = "1"
                                    End If
                                Next
                            End If
                            
                            Call ResetVsf(vsf)
                            Call ResetVsf(vsfPrice)
                        End If
                    End If
                    
                End If

                rs.MoveFirst

                Do While Not rs.EOF

                    Call ReadTemplate(rs("ID").Value)
                    rs.MoveNext

                Loop

                DataChange = True
            End If

        End If

        EnterFocus vsf
ErrHandler:

    End Select
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub cmdOK_Click()
        
    Dim lngKey As Long
    
    If Trim(cbo.Text) <> "" Then
        Call WritePrice(vsf.Row)
        Call SaveItems(Trim(cbo.Text))
    End If
        
    If ValidEdit = False Then Exit Sub
    
    mrsItems.Filter = ""
    
    mblnOK = True
    DataChange = False
    
    Unload Me
    
End Sub


Private Sub Form_Load()
    glngFormW = 10770
    glngFormH = 6780
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    
    Call RestoreWinState(Me, App.ProductName)
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    With fraTitle
        .Left = 0
        .Top = -90
        .Width = Me.ScaleWidth - .Left
    End With
    cbo.Move fraTitle.Width - cbo.Width - 45, cbo.Top
    lbl(4).Move cbo.Left - lbl(4).Width - 30
    
    With fra2
        .Left = 0
        .Top = fraTitle.Top + fraTitle.Height - 90
        .Width = fraTitle.Width
        .Height = Me.ScaleHeight - .Top - stbThis.Height - fraButton.Height + 90
    End With

    vsf.Move 45, vsf.Top, fra2.Width - vsf.Left - 45, fra2.Height - vsf.Top - 45 - vsfPrice.Height - 45
    
    With vsfPrice
        .Left = vsf.Left
        .Top = vsf.Top + vsf.Height + 45
        .Width = fra2.Width - .Left - 45
    End With
    
    cmd(6).Left = fra2.Width - cmd(6).Width - 60
    cmd(5).Left = cmd(6).Left - cmd(5).Width - 45
    
    With fraButton
        .Left = fra2.Left
        .Top = fra2.Top + fra2.Height - 90
        .Width = fra2.Width
    End With
    
    cmdCancel.Left = fraButton.Width - cmdCancel.Width - 60
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 45
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If DataChange Then
        Cancel = (MsgBox("数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    Call SaveWinState(Me, App.ProductName)
    
    If Not InDesign Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    
End Sub

Private Sub txtSum_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txtSum(Index)
        
End Sub

Private Sub txtSum_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        zlCommFun.PressKey vbKeyTab

    End If
End Sub

Private Sub txtSum_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtSum(Index).Locked Then
        glngTXTProc = GetWindowLong(txtSum(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtSum(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtSum_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtSum(Index).Locked Then
        Call SetWindowLong(txtSum(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub


Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    If vsf.Rows = 2 And Val(vsf.RowData(1)) = 0 Then
        Call ResetVsf(vsfPrice)
    Else
        Call ReadPrice(vsf.Row)
    End If
    
    Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.基本价格)), Val(vsf.TextMatrix(Row, mCol.体检价格)), 1)
    
    DataChange = True
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case mCol.执行科室
        
        vsf.TextMatrix(Row, mCol.执行科室id) = vsf.Body.ComboData
        vsf.TextMatrix(Row, mCol.执行科室) = vsf.Cell(flexcpTextDisplay, Row, mCol.执行科室)
        
    Case mCol.采集方式
        
        vsf.TextMatrix(Row, mCol.采集方式id) = vsf.Body.ComboData
        vsf.TextMatrix(Row, mCol.采集方式) = vsf.Cell(flexcpTextDisplay, Row, mCol.采集方式)
        
    Case mCol.采集科室
    
        vsf.TextMatrix(Row, mCol.采集科室id) = vsf.Body.ComboData
        vsf.TextMatrix(Row, mCol.采集科室) = vsf.Cell(flexcpTextDisplay, Row, mCol.采集科室)
        
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.体检价格
        
        Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.基本价格)), Val(vsf.TextMatrix(Row, mCol.体检价格)), 1)
        Call ReadPrice(Row)

    '------------------------------------------------------------------------------------------------------------------
    Case mCol.折扣
        
        Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.基本价格)), Val(vsf.TextMatrix(Row, mCol.折扣)), 2)
        Call ReadPrice(Row)
        
    End Select
    DataChange = True
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If NewRow = OldRow Then Exit Sub
    
    Call ReadPrice(NewRow)
    
    Call vsfPrice_BeforeRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col, False)

End Sub


Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case mbytMode
    Case 1
        If Val(vsf.TextMatrix(Row, mCol.公共)) = 1 Then
            
            If mblnGroup = False Then
                If MsgBox("此项目为团体公共项目，是否真的要删除？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If

    Case 2
'        If Val(vsf.TextMatrix(Row, mCol.公共)) = 1 Then
            If MsgBox("此项目已经开始，是否真的要取消？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
                Exit Sub
            End If
            
'        End If
    End Select
    
    '记录
    If Trim(cbo.Text) <> "" Then
        mrsItems.Filter = ""
'        mrsItems.Filter = "组别='" & Trim(cbo.Text) & "' AND ID=" & Val(vsf.RowData(Row))
        mrsItems.Filter = "组别='" & Trim(cbo.Text) & "' AND 清单id=" & Val(vsf.TextMatrix(Row, mCol.清单id))
        If mrsItems.RecordCount > 0 Then
            mrsItems.MoveFirst
            mrsItems("删除").Value = "1"
        End If
    End If

End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsf.RowData(Row)) <= 0)
    Cancel = (Val(vsf.TextMatrix(Row, mCol.执行科室id)) <= 0)
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    On Error GoTo errHand
    
    If NewRow = OldRow Then Exit Sub
    
    vsf.EditMode(mCol.项目) = 0
    vsf.EditMode(mCol.执行科室) = 0
    vsf.EditMode(mCol.检查部位) = 0
    vsf.EditMode(mCol.采集方式) = 0
    vsf.EditMode(mCol.采集科室) = 0
    vsf.EditMode(mCol.检验标本) = 0
    vsf.EditMode(mCol.结算方式) = 0
    vsf.EditMode(mCol.体检价格) = 0
    vsf.EditMode(mCol.折扣) = 0
    
    vsf.ComboList(mCol.项目) = ""
    vsf.ComboList(mCol.执行科室) = ""
    vsf.ComboList(mCol.检查部位) = ""
    vsf.ComboList(mCol.采集方式) = ""
    vsf.ComboList(mCol.采集科室) = ""
    vsf.ComboList(mCol.检验标本) = ""
    vsf.ComboList(mCol.结算方式) = ""

'    '设置编辑状态

    If Val(vsf.TextMatrix(NewRow, mCol.新加)) = 0 Then
        vsf.EditMode(mCol.项目) = 1
        vsf.EditMode(mCol.执行科室) = 1
        vsf.EditMode(mCol.结算方式) = 1
        vsf.EditMode(mCol.体检价格) = 1
        vsf.EditMode(mCol.折扣) = 1
        vsf.ComboList(mCol.项目) = "..."
        vsf.ComboList(mCol.执行科室) = " "
        vsf.ComboList(mCol.结算方式) = "记帐|收费"

        Select Case vsf.TextMatrix(NewRow, mCol.类别)
        Case "检查"
            vsf.EditMode(mCol.检查部位) = 1
             vsf.ComboList(mCol.检查部位) = "..."
        Case "检验"
            vsf.EditMode(mCol.采集方式) = 1
            vsf.EditMode(mCol.采集科室) = 1
            vsf.EditMode(mCol.检验标本) = 1

            vsf.ComboList(mCol.采集方式) = " "
            vsf.ComboList(mCol.采集科室) = " "
            vsf.ComboList(mCol.检验标本) = " "

        End Select

        vsfPrice.ComboList(mCol.p名称) = "..."
        vsfPrice.ComboList(mCol.p执行科室) = " "
        vsfPrice.ComboList(mCol.p计价项目) = " "
        vsfPrice.EditMode(mCol.p名称) = 1
        vsfPrice.EditMode(mCol.p数次) = 1
        vsfPrice.EditMode(mCol.p体检单价) = 1
        vsfPrice.EditMode(mCol.p执行科室) = 1
        vsfPrice.EditMode(mCol.p计价项目) = 1
        vsfPrice.EditMode(mCol.p折扣) = 1
    Else
        vsfPrice.ComboList(mCol.p名称) = ""
        vsfPrice.EditMode(mCol.p名称) = 0
        vsfPrice.EditMode(mCol.p数次) = 0
        vsfPrice.EditMode(mCol.p体检单价) = 0
        vsfPrice.EditMode(mCol.p执行科室) = 0
        vsfPrice.EditMode(mCol.p计价项目) = 0
        vsfPrice.EditMode(mCol.p折扣) = 0
    End If
    
    
    If Val(vsf.TextMatrix(OldRow, mCol.新加)) = 0 And OldRow > 0 Then
        Call WritePrice(OldRow)
    End If
    
    If Val(vsf.TextMatrix(NewRow, mCol.新加)) = 0 Then
        If vsf.TextMatrix(NewRow, mCol.类别) = "检验" Then
            Call SetRowData(Val(vsf.RowData(NewRow)), NewRow, "计价项目", "诊疗执行科室", "采集方式", "检验标本")
            Call SetRowData(Val(vsf.RowData(NewRow)), NewRow, "采集科室")
        ElseIf vsf.TextMatrix(NewRow, mCol.类别) = "检查" Then
            Call SetRowData(Val(vsf.RowData(NewRow)), NewRow, "计价项目", "诊疗执行科室")
        End If
    End If
    
    Exit Sub
    
errHand:

    
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim bytResult As Byte
    Dim rsPrice As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strKeys As String
    Dim rsData As New ADODB.Recordset
    
    Select Case Col
        Case mCol.项目
            
            gstrSQL = GetPublicSQL(SQL.体检项目选择)
            
            Select Case mstr性别
            Case "男"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 1, 1)
            Case "女"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 2, 2)
            Case Else
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 1, 2)
            End Select
        
            If ShowGrdSelect(Me, vsf, "编码,1200,0,1;名称,2700,0,0;单位,900,0,0;标本部位,900,0,0;类别,900,0,0", Me.Name & "\体检项目选择", "请从列表中选择一个体检项目。", rsData, rs, 8790, 4500) Then
                '选取了一个项目
                If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                    ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                    Exit Sub
                End If
                
                vsf.Cell(flexcpText, Row, mCol.项目 + 1, Row, vsf.Cols - 1) = ""
                
                vsf.EditText = zlCommFun.NVL(rs("名称").Value)
                vsf.TextMatrix(Row, mCol.类别) = zlCommFun.NVL(rs("类别").Value)
                vsf.TextMatrix(Row, mCol.项目) = zlCommFun.NVL(rs("名称").Value)
                vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                
                If vsf.TextMatrix(Row, mCol.类别) = "检验" Then
                    Call SetRowDefault(Val(vsf.RowData(Row)), Row, "执行科室", "采集方式", "采集科室", "检验标本", "结算方式", "计价项目")
                    
                ElseIf vsf.TextMatrix(Row, mCol.类别) = "检查" Then
                    Call SetRowDefault(Val(vsf.RowData(Row)), Row, "执行科室", "结算方式", "计价项目")
                End If
                
                Call CreatePriceList(Row)
                Call vsf_BeforeRowColChange(0, 0, vsf.Row, vsf.Col, False)
                Call vsfPrice_BeforeRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col, False)
                
                Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.基本价格)), Val(vsf.TextMatrix(Row, mCol.体检价格)), 1)
                
                DataChange = True
                
            End If
    End Select
    
    Select Case Col
            
        Case mCol.执行科室
            
            bytResult = ShowOpenList("", mCol.执行科室)
            If bytResult = 0 Then ShowSimpleMsg "没有找到相匹配的项目！"
            If bytResult = 1 Then DataChange = True
                
        Case mCol.检查部位
            
            bytResult = ShowOpenList("", mCol.检查部位)
            If bytResult = 0 Then ShowSimpleMsg "没有找到相匹配的项目！"
            If bytResult = 1 Then
                     
                Call CreatePriceList(Row)
                
                DataChange = True
            End If
            
        Case mCol.采集方式
            bytResult = ShowOpenList("", mCol.采集方式)
            If bytResult = 0 Then ShowSimpleMsg "没有找到相匹配的项目！"
            If bytResult = 1 Then
                Call CreatePriceList(Row)
                DataChange = True
            End If
            
    End Select
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim bytResult As Byte
    Dim rs As New ADODB.Recordset
    
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." And Col = mCol.项目 Then
            
            If InStr(vsf.EditText, "'") > 0 Then
                KeyCode = 0
                vsf.EditText = ""
                Cancel = True
                Exit Sub
            End If
                        
            bytResult = ShowOpenList(UCase(vsf.EditText), Col)
            
            If bytResult = 0 Then
                '没有匹配的项目
                KeyCode = 0
                Cancel = True
                
                vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                vsf.EditText = vsf.Cell(flexcpData, Row, Col)
                vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                
                MsgBox "没有找到相匹配的体检项目！", vbInformation, gstrSysName
            End If
            
            If bytResult = 1 Then
                '选取了一个项目
                DataChange = True
                
                If Col = mCol.项目 Then
                
                    If vsf.TextMatrix(Row, mCol.类别) = "检验" Then
                        Call SetRowDefault(Val(vsf.RowData(Row)), Row, "执行科室", "采集方式", "采集科室", "检验标本", "结算方式", "计价项目")
                        
                    ElseIf vsf.TextMatrix(Row, mCol.类别) = "检查" Then
                        Call SetRowDefault(Val(vsf.RowData(Row)), Row, "执行科室", "结算方式", "计价项目")
                    End If
                    
                    Call CreatePriceList(Row)
                    
                    Call vsf_BeforeRowColChange(0, 0, vsf.Row, vsf.Col, False)
                    Call vsfPrice_BeforeRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col, False)
                    
                    Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.基本价格)), Val(vsf.TextMatrix(Row, mCol.体检价格)), 1)
                    
'                    Call DefaultValue(Val(vsf.RowData(Row)), 1)
'
'                    If vsf.TextMatrix(Row, mCol.类别) = "检验" Then
'                        Call DefaultValue(Val(vsf.RowData(Row)), 2)
'                        Call DefaultValue(Val(vsf.RowData(Row)), 3)
'                    End If
                    
'                    Call CreatePriceList(Row)
                    
                End If
                
            End If
            
            If bytResult = 2 Then
                '取消了本次选择
                KeyCode = 0
                Cancel = True
                
                vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                
            End If
            
        End If
    Else
        DataChange = True
    End If
End Sub

Private Sub vsf_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    
    If KeyAscii = vbKeyReturn Then
                
        If Col = 1 Then
            If Trim(vsf.TextMatrix(Row, Col)) = "" Then
                
                KeyAscii = 0
                
                cmdOK.SetFocus
                
                Cancel = True
                
            End If
        End If
    End If
    
End Sub
Private Sub vsfPrice_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    
    Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p标准单价)), Val(vsfPrice.TextMatrix(Row, mCol.p体检单价)), 1)
    
    DataChange = True
End Sub

Private Sub vsfPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case mCol.p计价项目
    
        If Left(vsfPrice.TextMatrix(Row, mCol.p计价项目), 4) = "采集方式" Then
            vsfPrice.TextMatrix(Row, mCol.p计价性质) = "2"
        Else
            vsfPrice.TextMatrix(Row, mCol.p计价性质) = "1"
        End If
        vsfPrice.TextMatrix(Row, mCol.p计价项目) = vsfPrice.Cell(flexcpTextDisplay, Row, mCol.p计价项目)
        
'    Case mCol.p数次, mCol.p体检单价
'
'        vsfPrice.TextMatrix(Row, mCol.p标准金额) = Val(vsfPrice.TextMatrix(Row, mCol.p标准单价)) * Val(vsfPrice.TextMatrix(Row, mCol.p数次))
'        vsfPrice.TextMatrix(Row, mCol.p体检金额) = Val(vsfPrice.TextMatrix(Row, mCol.p体检单价)) * Val(vsfPrice.TextMatrix(Row, mCol.p数次))
'
'        vsf.TextMatrix(vsf.Row, mCol.基本价格) = SumPrice(1)
'        vsf.TextMatrix(vsf.Row, mCol.体检价格) = SumPrice(2)
        
    Case mCol.p数次
        vsfPrice.TextMatrix(Row, mCol.p标准金额) = Val(vsfPrice.TextMatrix(Row, mCol.p标准单价)) * Val(vsfPrice.TextMatrix(Row, mCol.p数次))
        vsfPrice.TextMatrix(Row, mCol.p体检金额) = Val(vsfPrice.TextMatrix(Row, mCol.p体检单价)) * Val(vsfPrice.TextMatrix(Row, mCol.p数次))
        
        Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p标准单价)), Val(vsfPrice.TextMatrix(Row, mCol.p体检单价)), 1)
                
        If InStr("567", vsfPrice.TextMatrix(Row, mCol.p类别)) > 0 Then
            Call PromptStorageWarn(Val(vsfPrice.TextMatrix(Row, mCol.p数次)), Val(vsfPrice.TextMatrix(Row, mCol.p可用库存)), vsfPrice.TextMatrix(Row, mCol.p名称), vsfPrice.TextMatrix(Row, mCol.p执行科室), vsfPrice.TextMatrix(Row, mCol.p计算单位), 1)
        End If
            
    Case mCol.p体检单价
        
        Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p标准单价)), Val(vsfPrice.TextMatrix(Row, mCol.p体检单价)), 1)
    
    Case mCol.p折扣
        
        Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p标准单价)), Val(vsfPrice.TextMatrix(Row, mCol.p折扣)), 2)
        
    Case mCol.p执行科室
        vsfPrice.TextMatrix(Row, mCol.p执行科室id) = vsfPrice.Body.ComboData
        vsfPrice.TextMatrix(Row, mCol.p执行科室) = vsfPrice.Cell(flexcpTextDisplay, Row, mCol.p执行科室)
    End Select
    
    DataChange = True
    
End Sub

Private Sub vsfPrice_AfterNewRow(ByVal Row As Long, Col As Long)
    
    If Row > 1 Then
        vsfPrice.TextMatrix(Row, mCol.p计价项目) = vsfPrice.TextMatrix(Row - 1, mCol.p计价项目)
        If Left(vsfPrice.TextMatrix(Row, mCol.p计价项目), 4) = "采集方式" Then
            vsfPrice.TextMatrix(Row, mCol.p计价性质) = "2"
        Else
            vsfPrice.TextMatrix(Row, mCol.p计价性质) = "1"
        End If
    End If
    
End Sub

Private Sub vsfPrice_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str计价项目 As String
    Dim str计价性质 As String
    
    If vsfPrice.Rows = 2 Then
        
        str计价项目 = vsfPrice.TextMatrix(1, mCol.p计价项目)
        str计价性质 = vsfPrice.TextMatrix(1, mCol.p计价性质)
        
        vsfPrice.Body.Cell(flexcpText, 1, mCol.p计价项目 + 1, 1, vsfPrice.Cols - 1) = ""
        vsfPrice.RowData(1) = 0

        vsfPrice.TextMatrix(1, mCol.p计价项目) = str计价项目
        vsfPrice.TextMatrix(1, mCol.p计价性质) = str计价性质
        Call vsfPrice_AfterDeleteRow(1, Col)
        
        Cancel = True
    End If
End Sub

Private Sub vsfPrice_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    If OldRow = NewRow Then Exit Sub
    
    Call SetRowData(Val(vsfPrice.RowData(NewRow)), NewRow, "收费执行科室")
    
End Sub

Private Sub vsfPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    If Col = mCol.p名称 Then
            
        gstrSQL = GetPublicSQL(SQL.收费项目选择)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        If ShowGrdSelect(Me, vsfPrice, "编码,1200,0,1;名称,2700,0,0;单位,600,0,0;规格,1200,0,0;单价,900,0,0;类别,900,0,0", Me.Name & "\收费项目选择", "请从列表中选择一个收费项目。", rsData, rs, 8790, 5100) Then

            If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                Exit Sub
            End If
            
            With vsfPrice
                .EditText = zlCommFun.NVL(rs("名称").Value)
                .TextMatrix(Row, mCol.p名称) = zlCommFun.NVL(rs("名称").Value)
                .TextMatrix(Row, mCol.p计算单位) = zlCommFun.NVL(rs("单位").Value)
                
                .TextMatrix(Row, mCol.p标准单价) = zlCommFun.NVL(rs("单价").Value, 0)
                .TextMatrix(Row, mCol.p体检单价) = .TextMatrix(Row, mCol.p标准单价)
                
                .TextMatrix(Row, mCol.p收费项目id) = zlCommFun.NVL(rs("ID").Value, 0)
                If Val(.TextMatrix(Row, mCol.p数次)) < 1 Then .TextMatrix(Row, mCol.p数次) = 1
                
                .TextMatrix(Row, mCol.p标准金额) = Val(.TextMatrix(Row, mCol.p标准单价)) * Val(.TextMatrix(Row, mCol.p数次))
                .TextMatrix(Row, mCol.p体检金额) = .TextMatrix(Row, mCol.p标准金额)
                .TextMatrix(Row, mCol.p类别) = zlCommFun.NVL(rs("类别").Value)
                
                .RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                
                Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p标准单价)), Val(vsfPrice.TextMatrix(Row, mCol.p体检单价)), 1)
                
                Call SetRowDefault(Val(.RowData(Row)), Row, "收费执行科室")
                Call SetRowData(Val(.RowData(Row)), Row, "收费执行科室")
            
                If InStr("567", .TextMatrix(Row, mCol.p类别)) > 0 Then
                    .TextMatrix(Row, mCol.p可用库存) = GetStorage(Val(.RowData(Row)), Val(.TextMatrix(Row, mCol.p执行科室id)))
                    Call PromptStorageWarn(Val(.TextMatrix(Row, mCol.p数次)), Val(.TextMatrix(Row, mCol.p可用库存)), .TextMatrix(Row, mCol.p名称), .TextMatrix(Row, mCol.p执行科室), .TextMatrix(Row, mCol.p计算单位), 1)
                End If
                
                
            End With
            
            DataChange = True

        End If
        
    End If
End Sub

Private Sub vsfPrice_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." Then
            
            If InStr(vsfPrice.EditText, "'") > 0 Then
                KeyCode = 0
                vsfPrice.EditText = ""
                Cancel = True
                Exit Sub
            End If
    
            Select Case Col
                Case mCol.p名称
                    
                    strText = UCase(vsfPrice.EditText)
                    gstrSQL = GetPublicSQL(SQL.收费项目过滤, strText)
                    
                    If ParamInfo.项目输入匹配方式 = 1 Then
                        strTmp = strText & "%"
                    Else
                        strTmp = "%" & strText & "%"
                    End If
                    
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText & "%", strTmp)
                    
                    If ShowGrdFilter(Me, vsfPrice, "编码,1200,0,1;名称,2700,0,0;单位,600,0,0;规格,1200,0,0;单价,900,0,0;类别,900,0,0", Me.Name & "\收费项目过滤", "请从列表中选择一个收费项目。", rsData, rs, 8790, 5100) Then
                        
                        If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                            Exit Sub
                        End If
                        With vsfPrice
                            .EditText = zlCommFun.NVL(rs("名称").Value)
                            .TextMatrix(Row, mCol.p名称) = zlCommFun.NVL(rs("名称").Value)
                            .TextMatrix(Row, mCol.p计算单位) = zlCommFun.NVL(rs("单位").Value)
                            
                            .TextMatrix(Row, mCol.p标准单价) = zlCommFun.NVL(rs("单价").Value, 0)
                            .TextMatrix(Row, mCol.p体检单价) = .TextMatrix(Row, mCol.p标准单价)
                            
                            .TextMatrix(Row, mCol.p收费项目id) = zlCommFun.NVL(rs("ID").Value, 0)
                            If Val(.TextMatrix(Row, mCol.p数次)) < 1 Then .TextMatrix(Row, mCol.p数次) = 1
                            
                            .TextMatrix(Row, mCol.p标准金额) = Val(.TextMatrix(Row, mCol.p标准单价)) * Val(.TextMatrix(Row, mCol.p数次))
                            .TextMatrix(Row, mCol.p体检金额) = .TextMatrix(Row, mCol.p标准金额)
                            .TextMatrix(Row, mCol.p类别) = zlCommFun.NVL(rs("类别").Value)
                            
                            .RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                            
                            Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p标准单价)), Val(vsfPrice.TextMatrix(Row, mCol.p体检单价)), 1)
                            
                            Call SetRowDefault(Val(.RowData(Row)), Row, "收费执行科室")
                            Call SetRowData(Val(.RowData(Row)), Row, "收费执行科室")
                            
                            If InStr("567", .TextMatrix(Row, mCol.p类别)) > 0 Then
                                .TextMatrix(Row, mCol.p可用库存) = GetStorage(Val(.RowData(Row)), Val(.TextMatrix(Row, mCol.p执行科室id)))
                                Call PromptStorageWarn(Val(.TextMatrix(Row, mCol.p数次)), Val(.TextMatrix(Row, mCol.p可用库存)), .TextMatrix(Row, mCol.p名称), .TextMatrix(Row, mCol.p执行科室), .TextMatrix(Row, mCol.p计算单位), 1)
                            End If
                        End With
                        
                        DataChange = True
                    Else
                        KeyCode = 0
                        Cancel = True
                        
                        vsfPrice.Cell(flexcpData, Row, Col) = vsfPrice.Cell(flexcpData, Row, Col)
                        vsfPrice.EditText = vsfPrice.Cell(flexcpData, Row, Col)
                        vsfPrice.TextMatrix(Row, Col) = vsfPrice.Cell(flexcpData, Row, Col)
                        
                    End If
            End Select
        End If
    Else
        DataChange = True
    End If
End Sub









