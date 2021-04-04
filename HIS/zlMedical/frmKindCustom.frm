VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKindCustom 
   Caption         =   "体检项目设置"
   ClientHeight    =   7035
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   10815
   Icon            =   "frmKindCustom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   10815
   Begin VB.PictureBox picBack 
      Height          =   3870
      Left            =   0
      ScaleHeight     =   3810
      ScaleWidth      =   10440
      TabIndex        =   4
      Top             =   645
      Width           =   10500
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   2
         Left            =   8760
         MaxLength       =   16
         TabIndex        =   16
         Top             =   60
         Width           =   1020
      End
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   1
         Left            =   7215
         MaxLength       =   16
         TabIndex        =   15
         Top             =   60
         Width           =   870
      End
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   0
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   60
         Width           =   930
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&P"
         Height          =   285
         Left            =   3675
         TabIndex        =   6
         Top             =   75
         Width           =   300
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   1110
         TabIndex        =   5
         Top             =   60
         Width           =   2550
      End
      Begin zl9Medical.VsfGrid vsf 
         Height          =   1500
         Left            =   60
         TabIndex        =   8
         Top             =   405
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   2646
      End
      Begin zl9Medical.VsfGrid vsfPrice 
         Height          =   1635
         Left            =   60
         TabIndex        =   9
         Top             =   2130
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   2884
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "折扣(Z)"
         Height          =   180
         Index           =   3
         Left            =   8115
         TabIndex        =   17
         Top             =   120
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检价格(E)"
         Height          =   180
         Index           =   2
         Left            =   6210
         TabIndex        =   13
         Top             =   120
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "基本价格(&B)"
         Height          =   180
         Index           =   0
         Left            =   4230
         TabIndex        =   12
         Top             =   120
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检类型(&T)"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   7
         Top             =   120
         Width           =   990
      End
   End
   Begin VB.Frame fraTitle 
      Height          =   600
      Left            =   0
      TabIndex        =   10
      Top             =   -90
      Width           =   6870
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
         TabIndex        =   11
         Top             =   195
         Width           =   1800
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6675
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmKindCustom.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13996
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
   Begin VB.Frame fraButton 
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   4455
      Width           =   6870
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   4230
         TabIndex        =   3
         Top             =   225
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5400
         TabIndex        =   2
         Top             =   225
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmKindCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean
Private mstrName As String
Private Enum mCol
    项目类别 = 1
    项目名称
    采集方式
    
    检验标本
    
    检查部位
    
    基本价格
    体检价格
    折扣
    计费明细
    采集方式id
    检查部位id
    
    计价项目 = 1
    收费项目
    计算单位
    收费数次
    收费单价
    体检单价
    p折扣
    收费金额
    体检金额
    收费项目id
    计价性质
End Enum

Private mstrSQL As String

'（２）自定义过程或函数************************************************************************************************

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
    str计价项目 = vsfPrice.TextMatrix(1, mCol.计价项目)
    str计价性质 = vsfPrice.TextMatrix(1, mCol.计价性质)
    
    vsfPrice.Body.Cell(flexcpText, 1, mCol.计价项目 + 1, 1, vsfPrice.Cols - 1) = ""
    vsfPrice.RowData(1) = 0

    vsfPrice.TextMatrix(1, mCol.计价项目) = str计价项目
    vsfPrice.TextMatrix(1, mCol.计价性质) = str计价性质

    
    mstrSQL = GetPublicSQL(SQL.体检项目价表, strKeys)
    
    If vsf.TextMatrix(intRow, mCol.检查部位id) = "" Then
        '检验或单部位检查
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, Val(vsf.RowData(intRow)), Val(vsf.TextMatrix(intRow, mCol.采集方式id)))
    Else
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    End If
    
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If Val(vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.收费项目id)) > 0 Then
                vsfPrice.Rows = vsfPrice.Rows + 1
            End If
            
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.收费项目) = zlCommFun.NVL(rs("名称"))
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.计算单位) = zlCommFun.NVL(rs("计算单位"))
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.收费数次) = zlCommFun.NVL(rs("收费数量"))
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.收费单价) = zlCommFun.NVL(rs("现价"))
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.体检单价) = zlCommFun.NVL(rs("现价"))
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p折扣) = 10
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.收费金额) = zlCommFun.NVL(rs("收费数量"), 0) * zlCommFun.NVL(rs("现价"), 0)
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.体检金额) = zlCommFun.NVL(rs("收费数量"), 0) * zlCommFun.NVL(rs("现价"), 0)
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.收费项目id) = zlCommFun.NVL(rs("ID"))
            
            rs.MoveNext
        Loop
    End If
    
    vsf.TextMatrix(intRow, mCol.基本价格) = SumPrice(mCol.收费金额)
    vsf.TextMatrix(intRow, mCol.体检价格) = SumPrice(mCol.体检金额)

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
    
    If dbMoney = 0 Then Exit Function
    
    If bytMode = 1 Then
        '变化金额
        
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
                varCol(6) = Format(Val(varCol(3)) * (db折扣 / 10), "0.00000")
                varCol(7) = db折扣
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
                        If Val(varCol(6)) <> 0 Then
                            varCol(6) = Val(varCol(6)) + (Val(txtSum(1).Text) - dbTotal)
                            If Val(varCol(3)) <> 0 Then
                                varCol(7) = Format(10 * Val(varCol(6)) / Val(varCol(3)), "0.0000")
                            Else
                                varCol(7) = 0
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
        If bytMode = 1 Then
            '变化金额
            If dbMoney = 0 Then Exit Function
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
                varCol(6) = Format(Val(varCol(3)) * (db折扣 / 10), "0.00000")
                varCol(7) = db折扣
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
    
    If dbMoney = 0 Then Exit Function
    
    If bytMode = 1 Then
        '变化金额
        
        '1.计算折扣
        db折扣 = Format(10 * dbTmp / dbMoney, "0.0000")
    Else
        '变化折扣
        db折扣 = dbTmp
        
    End If
    
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.体检单价) = Format(dbMoney * db折扣 / 10, "0.00000")
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.p折扣) = Format(db折扣, "0.0000")
    
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.体检金额) = Val(vsfPrice.TextMatrix(vsfPrice.Row, mCol.收费数次)) * Val(vsfPrice.TextMatrix(vsfPrice.Row, mCol.体检单价))
    
    '更新项目
    '------------------------------------------------------------------------------------------------------------------
    dbSum = 0
    For lngLoop = 1 To vsfPrice.Rows - 1
       dbSum = dbSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.收费金额))
    Next
    vsf.TextMatrix(vsf.Row, mCol.基本价格) = dbSum
    
    dbSum = 0
    For lngLoop = 1 To vsfPrice.Rows - 1
       dbSum = dbSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.体检金额))
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

Private Function SumPrice(ByVal intCol As Integer, Optional ByVal bytMode As Byte = 1) As Single
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim sglSum As Single
    
    If bytMode = 2 Then
        '体检类型的价格
        For lngLoop = 1 To vsf.Rows - 1
           sglSum = sglSum + Val(vsf.TextMatrix(lngLoop, intCol))
        Next
    Else
        For lngLoop = 1 To vsfPrice.Rows - 1
           sglSum = sglSum + Val(vsfPrice.TextMatrix(lngLoop, intCol))
        Next
    End If
    SumPrice = sglSum
    
End Function

Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '值域:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
            
    txt.Locked = False
    cmd.Enabled = True
    
    If vData = False Then
        cmdOK.Tag = ""
    Else
        cmdOK.Tag = "Changed"
        txt.Locked = True
        cmd.Enabled = False
    End If
End Property

Private Property Get EditChanged() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '值域:
    '------------------------------------------------------------------------------------------------------------------
            
    EditChanged = (cmdOK.Tag = "Changed")
    
End Property

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
        
        Case "计价项目"
            
            If Trim(vsf.TextMatrix(intRow, mCol.项目类别)) = "检查" Then
                strCombList = "检查项目-" & Trim(vsf.TextMatrix(intRow, mCol.项目名称))
                vsfPrice.EditMode(mCol.计价项目) = 0
                vsfPrice.Body.ColComboList(mCol.计价项目) = ""
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.计价项目) = strCombList
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.计价性质) = "1"
            Else
                strCombList = "检验项目-" & Trim(vsf.TextMatrix(intRow, mCol.项目名称))
                If Val(vsf.TextMatrix(intRow, mCol.采集方式id)) > 0 Then
                    strCombList = strCombList & "|采集方式-" & Trim(vsf.TextMatrix(intRow, mCol.采集方式))
                    vsfPrice.EditMode(mCol.计价项目) = 1
                    vsfPrice.Body.ColComboList(mCol.计价项目) = strCombList
                Else
                    vsfPrice.EditMode(mCol.计价项目) = 0
                    vsfPrice.Body.ColComboList(mCol.计价项目) = ""
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
        Case "计价项目"
        
            If Trim(vsf.TextMatrix(vsf.Row, mCol.项目类别)) = "检查" Then
                strCombList = "检查项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目名称))
                vsfPrice.EditMode(mCol.计价项目) = 0
                vsfPrice.Body.ColComboList(mCol.计价项目) = ""
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.计价项目) = strCombList
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.计价性质) = "1"
            Else
                strCombList = "检验项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目名称))
                If Val(vsf.TextMatrix(vsf.Row, mCol.采集方式id)) > 0 Then
                    strCombList = strCombList & "|采集方式-" & Trim(vsf.TextMatrix(vsf.Row, mCol.采集方式))
                    vsfPrice.EditMode(mCol.计价项目) = 1
                    vsfPrice.Body.ColComboList(mCol.计价项目) = strCombList
                Else
                    vsfPrice.EditMode(mCol.计价项目) = 0
                    vsfPrice.Body.ColComboList(mCol.计价项目) = ""
                End If
            End If

        End Select
    Next
    
    SetRowDefault = True
    
    Exit Function
    
errHand:
    
End Function

Private Function SetDefault(ByVal lng诊疗项目id As Long, ParamArray arryMode() As Variant) As Boolean
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
            
        Case "采集方式"
        
            gstrSQL = "SELECT A.名称 AS 名称,A.ID FROM 诊疗项目目录 A,诊疗用法用量 B WHERE A.ID=B.用法id AND A.类别='E' AND A.操作类型='6' AND B.项目ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng诊疗项目id)
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
        
        Case "计价项目"
        
            If Trim(vsf.TextMatrix(vsf.Row, mCol.项目类别)) = "检查" Then
            
                strCombList = "检查项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目名称))
                vsfPrice.EditMode(mCol.计价项目) = 0
                vsfPrice.Body.ColComboList(mCol.计价项目) = ""
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.计价项目) = strCombList
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.计价性质) = "1"
                
            Else
            
                strCombList = "检验项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目名称))
                If Val(vsf.TextMatrix(vsf.Row, mCol.采集方式id)) > 0 Then
                    strCombList = strCombList & "|采集方式-" & Trim(vsf.TextMatrix(vsf.Row, mCol.采集方式))
                    vsfPrice.EditMode(mCol.计价项目) = 1
                    vsfPrice.Body.ColComboList(mCol.计价项目) = strCombList
                Else
                    vsfPrice.EditMode(mCol.计价项目) = 0
                    vsfPrice.Body.ColComboList(mCol.计价项目) = ""
                End If
                
            End If
            
        Case "检验标本"
        
            gstrSQL = "SELECT 1 FROM 诊疗项目目录 WHERE 组合项目=1 AND ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng诊疗项目id)
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
        
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng诊疗项目id)
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
            
        End Select
    Next
    
        
    SetDefault = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ReadSample(ByVal lng诊疗项目id As Long) As String
    '------------------------------------------------------------------------------------------------------------------
    '功能:获取可选的采集方式下拉数据
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand

    gstrSQL = "SELECT distinct A.标本类型 AS 名称 FROM 检验项目参考 A,检验报告项目 B,诊疗项目目录 C " & _
            "WHERE C.ID=[1] AND B.诊疗项目id=[1] and B.报告项目id=A.项目id"

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


Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    
    vsf.Rows = 2
    vsf.RowData(1) = 0
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    
    Call ResetVsf(vsfPrice)
        
End Function

Public Function ShowEdit(ByVal frmMain As Object, ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
            
    mlngKey = lngKey
    
    
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    Call InitSysPara
    
    If mlngKey > 0 Then
        Call ReadData(mlngKey)
        Call ReadPrice(vsf.Row)
    End If
            
    EditChanged = False
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  读取数据
    '参数:  lngKey      体检类型序号
    '返回:  True        读取成功
    '       False       读取失败
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
        
    On Error GoTo errHand
    
    gstrSQL = "SELECT * FROM 体检类型 WHERE 序号=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then txt.Text = rs("名称").Value
    mstrName = txt.Text
    
    stbThis.Panels(2).Text = "体检类型:" & txt.Text & "  编码:" & rs("编码").Value
    
    gstrSQL = GetPublicSQL(SQL.体检类型项目, CStr(lngKey))
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                vsf.Rows = vsf.Rows + 1
            End If
            
            vsf.RowData(vsf.Rows - 1) = zlCommFun.NVL(rs("ID"), 0)
            vsf.TextMatrix(vsf.Rows - 1, mCol.项目名称) = zlCommFun.NVL(rs("项目名称"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.项目类别) = zlCommFun.NVL(rs("项目类别"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.检查部位) = zlCommFun.NVL(rs("检查部位"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.采集方式) = zlCommFun.NVL(rs("采集方式"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.检验标本) = zlCommFun.NVL(rs("检验标本"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.基本价格) = zlCommFun.NVL(rs("基本价格"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.体检价格) = zlCommFun.NVL(rs("体检价格"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.折扣) = zlCommFun.NVL(rs("折扣"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.采集方式id) = zlCommFun.NVL(rs("采集方式id"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.检查部位id) = zlCommFun.NVL(rs("检查部位id"))
            
            vsf.TextMatrix(vsf.Rows - 1, mCol.计费明细) = GetTypePriceList(lngKey, zlCommFun.NVL(rs("ID"), 0))
                        
            rs.MoveNext
        Loop
    End If
    
    Call ChangeItem(Val(vsf.TextMatrix(vsf.Row, mCol.基本价格)), Val(vsf.TextMatrix(vsf.Row, mCol.体检价格)), 1, False)

                
    ReadData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    
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
        .NewColumn "项目类别", 900, 1
        .NewColumn "项目名称", 2700, 1, "...", 1
        .NewColumn "采集方式", 1200, 1
        .NewColumn "检验标本", 900, 1
        .NewColumn "检查部位", 1500, 1
        .NewColumn "基本价格", 900, 7
        .NewColumn "体检价格", 900, 7, , 1
        .NewColumn "折扣", 900, 7, , 1
        .NewColumn "计费明细", 0, 1
        .NewColumn "采集方式id", 0, 1
        .NewColumn "检查部位id", 0, 1
        .FixedCols = 1
        
        .Body.ColFormat(mCol.基本价格) = "0.00"
        .Body.ColFormat(mCol.体检价格) = "0.00"
        .Body.ColFormat(mCol.折扣) = "0.0000"
        .SelectMode = True
    End With
    
    With vsfPrice
        .Cols = 0
        
        .NewColumn "", 255, 4
        
        .NewColumn "计价项目", 3000, 1, " |", 1
        .NewColumn "收费项目", 2100, 1, "...", 1
        .NewColumn "单位", 900, 1
        .NewColumn "数次", 600, 7, , 1
        .NewColumn "收费单价", 900, 7
        .NewColumn "体检单价", 900, 7, , 1
        .NewColumn "折扣", 900, 7, , 1
        .NewColumn "收费金额", 900, 7
        .NewColumn "体检金额", 900, 7
        .NewColumn "收费项目id", 0, 1
        .NewColumn "计价性质", 0, 1
        
        .Body.ColFormat(mCol.收费单价) = "0.00000"
        .Body.ColFormat(mCol.收费金额) = "0.00"
        .Body.ColFormat(mCol.p折扣) = "0.0000"
        .Body.ColFormat(mCol.体检单价) = "0.00000"
        .Body.ColFormat(mCol.体检金额) = "0.00"
        
        .FixedCols = 1
        .SelectMode = True
    End With
    
    InitData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
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
    
                    If Val(varCol(5)) = 2 Then
                        vsfPrice.TextMatrix(lngRow + 1, mCol.计价项目) = "采集方式-" & Trim(vsf.TextMatrix(vsf.Row, mCol.采集方式))
                    ElseIf Trim(vsf.TextMatrix(vsf.Row, mCol.项目类别)) = "检验" Then
                        vsfPrice.TextMatrix(lngRow + 1, mCol.计价项目) = "检验项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目名称))
                    Else
                        vsfPrice.TextMatrix(lngRow + 1, mCol.计价项目) = "检查项目-" & Trim(vsf.TextMatrix(vsf.Row, mCol.项目名称))
                    End If
                    
                    vsfPrice.TextMatrix(lngRow + 1, mCol.收费项目) = varCol(0)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.计算单位) = varCol(1)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.收费数次) = varCol(2)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.收费单价) = varCol(3)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.体检单价) = varCol(6)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.收费金额) = Val(varCol(2)) * Val(varCol(3))
                    vsfPrice.TextMatrix(lngRow + 1, mCol.体检金额) = Val(varCol(2)) * Val(varCol(6))
                    vsfPrice.TextMatrix(lngRow + 1, mCol.收费项目id) = varCol(4)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.计价性质) = varCol(5)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p折扣) = varCol(7)
                    
                Next
            End If
        Next
        
    End If
    
    If vsf.TextMatrix(intRow, mCol.项目类别) = "检验" Then
        Call SetRowData(Val(vsf.RowData(intRow)), intRow, "计价项目")
    ElseIf vsf.TextMatrix(intRow, mCol.项目类别) = "检查" Then
        Call SetRowData(Val(vsf.RowData(intRow)), intRow, "计价项目")
    End If
    
    ReadPrice = True
    
End Function

Private Function WritePrice(ByVal intRow As Integer) As Boolean
    Dim strTmp As String
    Dim lngRow As Long
    Dim varCol As Variant
    
    On Error GoTo errHand
    
    For lngRow = 1 To vsfPrice.Rows - 1
        If Val(vsfPrice.TextMatrix(lngRow, mCol.收费项目id)) > 0 Then
            
            varCol = Split(String(8, ":"), ":")
                                
            varCol(0) = vsfPrice.TextMatrix(lngRow, mCol.收费项目)
            varCol(1) = vsfPrice.TextMatrix(lngRow, mCol.计算单位)
            varCol(2) = vsfPrice.TextMatrix(lngRow, mCol.收费数次)
            varCol(3) = vsfPrice.TextMatrix(lngRow, mCol.收费单价)
            varCol(4) = vsfPrice.TextMatrix(lngRow, mCol.收费项目id)
            varCol(5) = vsfPrice.TextMatrix(lngRow, mCol.计价性质)
            varCol(6) = vsfPrice.TextMatrix(lngRow, mCol.体检单价)
            varCol(7) = vsfPrice.TextMatrix(lngRow, mCol.p折扣)
                        
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

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  校验数据的有效性
    '返回:  True        数据有效
    '       False       数据无效
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If lngLoop <> vsf.Rows - 1 Then
            If Val(vsf.RowData(lngLoop)) = 0 Then
                ShowSimpleMsg "第 " & lngLoop & " 行数据输入不完整，必须输入有效的体检项目！"
                LocationGrid vsf, lngLoop, mCol.项目名称
                Exit Function
            End If
        End If
    Next
    
    ValidEdit = True
    
End Function

Private Function SaveEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  保存数据
    '返回:  True        保存成功
    '       False       保存失败
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim strTmp As String
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngRow As Long
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    strSQL(ReDimArray(strSQL)) = "ZL_体检类型目录_DELETE(" & mlngKey & ")"
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            
            
            strTmp = ""
            If vsf.TextMatrix(lngLoop, mCol.计费明细) <> "" Then
                varRow = Split(vsf.TextMatrix(lngLoop, mCol.计费明细), ";")
                For lngRow = 0 To UBound(varRow)
                    
                    varCol = Split(varRow(lngRow), ":")
                    
                    If strTmp <> "" Then strTmp = strTmp & ";"
                    strTmp = strTmp & varCol(4) & ":" & varCol(2) & ":" & varCol(5) & ":" & Format(varCol(7) / 10, "0.00000")
                            
                Next
            End If
            
            strSQL(ReDimArray(strSQL)) = "ZL_体检类型目录_INSERT(" & mlngKey & "," & _
                                                                Val(vsf.RowData(lngLoop)) & ",'" & _
                                                                vsf.TextMatrix(lngLoop, mCol.检查部位) & "'," & _
                                                                Val(vsf.TextMatrix(lngLoop, mCol.采集方式id)) & ",'" & _
                                                                vsf.TextMatrix(lngLoop, mCol.检查部位id) & "','" & _
                                                                vsf.TextMatrix(lngLoop, mCol.检验标本) & "','" & _
                                                                strTmp & "')"
        End If
    Next
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveEdit = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran Then gcnOracle.RollbackTrans
    
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


'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub cmd_Click()
    
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    gstrSQL = GetPublicSQL(SQL.体检类型选择)
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    If ShowTxtSelect(Me, txt, "编码,1080,0,1;名称,2400,0,0;简码,900,0,0;说明,1500,0,0", Me.Name & "\体检类型选择", "请从列表中选择一个体检类型。", rsData, rs, 8790, 5100) Then
        
        Call ClearData
        
        txt.Text = zlCommFun.NVL(rs("名称"))
        mlngKey = zlCommFun.NVL(rs("ID"))
        
        Call ReadData(mlngKey)
        Call ReadPrice(vsf.Row)
        
        txt.Tag = ""
        mstrName = txt.Text
        
        EditChanged = False
        
    End If

    Call LocationObj(txt)

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    Call WritePrice(vsf.Row)
    
    If ValidEdit = False Then Exit Sub
    
    If SaveEdit Then
       
        Call mfrmMain.EditRefresh("体检类型", mlngKey)
        
        If mlngKey = 0 Then
            
            EditChanged = False
        Else
            EditChanged = False
        End If
        
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With fraTitle
        .Left = 0
        .Top = -90
        .Width = Me.ScaleWidth - .Left
    End With
    
    With picBack
        .Left = 0
        .Top = fraTitle.Top + fraTitle.Height
        .Width = fraTitle.Width
        .Height = Me.ScaleHeight - .Top - stbThis.Height - fraButton.Height
    End With
    
    With fraButton
        .Left = picBack.Left
        .Top = picBack.Top + picBack.Height - 90
        .Width = picBack.Width
    End With
    
    cmdCancel.Left = fraButton.Width - cmdCancel.Width - 60
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 45

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If EditChanged Then
        Cancel = (MsgBox("新增或修改的数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picBack_Resize()
    On Error Resume Next
    
    With vsf
        
        .Left = 60
        .Top = txt.Top + txt.Height + 45
        .Width = picBack.Width - .Left - 60
        .Height = picBack.Height - .Top - vsfPrice.Height - 60 - 45
        
    End With
    
    With vsfPrice
        .Left = vsf.Left
        .Top = vsf.Top + vsf.Height + 45
        .Width = picBack.Width - .Left - 60
    End With
    
End Sub

Private Sub txt_Change()
    txt.Tag = "Changed"
End Sub

Private Sub txt_GotFocus()
    zlControl.TxtSelAll txt
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    Dim strText As String
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        If txt.Tag = "Changed" Then
            
            txt.Tag = ""
            strText = UCase(txt.Text)
            
            gstrSQL = GetPublicSQL(SQL.体检类型过滤选择)
            
            If ParamInfo.项目输入匹配方式 = 1 Then
                strTmp = strText & "%"
            Else
                strTmp = "%" & strText & "%"
            End If
                    
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText & "%", strTmp)
            
            If ShowTxtFilter(Me, txt, "编码,1080,0,1;名称,2400,0,0;简码,900,0,0;说明,1500,0,0", Me.Name & "\体检类型过滤选择", "请从列表中选择一个体检类型。", rsData, rs) Then
                
                Call ClearData
                
                txt.Text = zlCommFun.NVL(rs("名称"))
                mlngKey = zlCommFun.NVL(rs("ID"))
                
                Call ReadData(mlngKey)
                Call ReadPrice(vsf.Row)
                
                txt.Tag = ""
                mstrName = txt.Text
                
                zlCommFun.PressKey vbKeyTab
                zlCommFun.PressKey vbKeyTab
            Else
                txt.Text = mstrName
            End If
            txt.Tag = ""
            Call LocationObj(txt)
            
        Else
            zlCommFun.PressKey vbKeyTab
            zlCommFun.PressKey vbKeyTab
        End If
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
    
End Sub

Private Sub txt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 And txt.Locked Then
        glngTXTProc = GetWindowLong(txt.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt.Locked Then
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Cancel As Boolean)

    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)

    If txt.Tag = "Changed" Then txt.Text = mstrName
    
End Sub

Private Sub txtSum_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txtSum(Index)
        
End Sub

Private Sub txtSum_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        Call WritePrice(vsf.Row)
        
        If Index = 1 Then Call ChangeTotal(Val(txtSum(0).Text), Val(txtSum(1).Text), 1)
        If Index = 2 Then Call ChangeTotal(Val(txtSum(0).Text), Val(txtSum(2).Text), 2)
        
        Call ReadPrice(vsf.Row)
   
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
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

Private Sub txtSum_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txtSum(Index).Text, txtSum(Index).MaxLength)
    
    If Index = 1 Then
        If InStr(txtSum(1).Text, ".") > 0 Then
            If Len(Mid(txtSum(1).Text, InStr(txtSum(1).Text, ".") + 1)) > 2 Then
                MsgBox "只允许输入两位小数位数。", vbExclamation, gstrSysName
                Cancel = True
            End If
        End If
    End If
    
End Sub

Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    If vsf.Rows = 2 And Val(vsf.RowData(1)) = 0 Then
        Call ResetVsf(vsfPrice)
    Else
        Call ReadPrice(vsf.Row)
    End If
    
    Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.基本价格)), Val(vsf.TextMatrix(Row, mCol.体检价格)), 1)
    
    EditChanged = True
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim db折扣 As Double
    Dim lngLoop As Long
    
    Select Case Col
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.项目类别
        If vsf.EditText <> vsf.Cell(flexcpData, Row, Col) Then
            vsf.RowData(Row) = 0
            vsf.Cell(flexcpText, Row, mCol.项目名称, Row, vsf.Cols - 1) = ""
            
            EditChanged = True
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.体检价格
        
        Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.基本价格)), Val(vsf.TextMatrix(Row, mCol.体检价格)), 1)
        Call ReadPrice(Row)

    '------------------------------------------------------------------------------------------------------------------
    Case mCol.折扣
        
        Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.基本价格)), Val(vsf.TextMatrix(Row, mCol.折扣)), 2)
        Call ReadPrice(Row)
        
    End Select
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    '设置编辑状态
    
    If NewRow = OldRow Then Exit Sub
    
    Select Case vsf.TextMatrix(NewRow, mCol.项目类别)
        Case "检查"
            vsf.EditMode(mCol.采集方式) = 0
            vsf.EditMode(mCol.检验标本) = 0
            vsf.EditMode(mCol.检查部位) = 1
            
            vsf.ComboList(mCol.采集方式) = ""
            vsf.ComboList(mCol.检验标本) = ""
            vsf.ComboList(mCol.检查部位) = "..."
        Case "检验"
            vsf.EditMode(mCol.采集方式) = 1
            vsf.EditMode(mCol.检验标本) = 1
            vsf.EditMode(mCol.检查部位) = 0
            
            vsf.ComboList(mCol.采集方式) = "..."
            vsf.ComboList(mCol.检验标本) = " "
            vsf.ComboList(mCol.检查部位) = ""
    End Select
    
    '计价项目列表
    
    mstrSQL = "Select * From 体检类型目录 where 序号=[1] And 诊疗项目id=1"
    
    Call ReadPrice(NewRow)

End Sub

Private Sub vsf_BeforeComboList(ByVal OldCol As Long, ByVal NewCol As Long, ComboList As String, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    
    '产生下接数据列表
    
    Select Case NewCol
    Case mCol.检验标本
        
        ComboList = ReadSample(Val(vsf.RowData(vsf.Row)))
        
    End Select
    
    If ComboList = "" Then ComboList = " |"
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsf.RowData(Row)) <= 0)
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    On Error GoTo errHand
    
    Call WritePrice(OldRow)
    
    Exit Sub
    
errHand:
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    Select Case Col
        Case mCol.项目名称
            
            gstrSQL = GetPublicSQL(SQL.体检项目选择)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 1, 2)
            If ShowGrdSelect(Me, vsf, "编码,1200,0,1;名称,2700,0,0;单位,900,0,0;类别,900,0,0", Me.Name & "\体检项目选择", "请从列表中选择一个体检项目。", rsData, rs, 8790, 5100) Then
                
                If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                    ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                    Exit Sub
                End If
                
                vsf.EditText = zlCommFun.NVL(rs("名称").Value)
                vsf.TextMatrix(Row, mCol.项目类别) = zlCommFun.NVL(rs("类别").Value)
                vsf.TextMatrix(Row, mCol.项目名称) = zlCommFun.NVL(rs("名称").Value)
                vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                                
                If vsf.TextMatrix(Row, mCol.项目类别) = "检验" Then
                    Call SetDefault(Val(vsf.RowData(Row)), "执行科室", "检验标本", "采集方式", "采集执行", "计价项目")
                Else
                    Call SetDefault(Val(vsf.RowData(Row)), "执行科室", "计价项目")
                End If
                
                Call CreatePriceList(Row)
                Call WritePrice(Row)
                
                Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.基本价格)), Val(vsf.TextMatrix(Row, mCol.体检价格)), 1)
                
                EditChanged = True
                
            End If
        
        Case mCol.检查部位
                        
            gstrSQL = "select B.标本部位 AS 名称,B.ID,0 AS 选择 from 诊疗项目组合 A,诊疗项目目录 B WHERE (B.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or B.撤档时间 is NULL) AND A.诊疗项目ID=B.ID AND A.诊疗组合ID=[1]"
            
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(Row)))
            
            If ShowGrdSelect(Me, vsf, "名称,3300,0,0", Me.Name & "\检查部位选择", "请从列表中选择一个检查部位。", rsData, rs, 8790, 5100) Then
                
                vsf.TextMatrix(Row, vsf.Col) = ""
                vsf.TextMatrix(Row, mCol.检查部位id) = ""
                
                rs.Filter = ""
                rs.Filter = "选择=1"
                If rs.RecordCount > 0 Then
                    rs.MoveFirst
                    Do While Not rs.EOF
                        vsf.TextMatrix(Row, vsf.Col) = vsf.TextMatrix(Row, vsf.Col) & zlCommFun.NVL(rs("名称").Value) & ","
                        vsf.TextMatrix(Row, mCol.检查部位id) = vsf.TextMatrix(Row, mCol.检查部位id) & zlCommFun.NVL(rs("ID").Value) & ","
                        rs.MoveNext
                    Loop
                    
                    If vsf.TextMatrix(Row, mCol.检查部位) <> "" Then vsf.TextMatrix(Row, mCol.检查部位) = Mid(vsf.TextMatrix(Row, mCol.检查部位), 1, Len(vsf.TextMatrix(Row, mCol.检查部位)) - 1)
                    If vsf.TextMatrix(Row, mCol.检查部位id) <> "" Then vsf.TextMatrix(Row, mCol.检查部位id) = Mid(vsf.TextMatrix(Row, mCol.检查部位id), 1, Len(vsf.TextMatrix(Row, mCol.检查部位id)) - 1)
                    
                End If
                                
                Call CreatePriceList(Row)
                
                EditChanged = True
                
            End If
            
        Case mCol.采集方式
        
            
            gstrSQL = "SELECT 1 As 末级,A.ID,A.编码,A.名称 " & _
                "FROM 诊疗项目目录 A,诊疗用法用量 B " & _
                "WHERE (A.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or A.撤档时间 is NULL) AND A.类别='E' AND A.操作类型='6' AND A.ID=B.用法id AND B.项目id=[1]"
            
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(Row)))
            If rs.BOF Then
            
                gstrSQL = "SELECT 1 As 末级,A.ID,A.编码,A.名称 " & _
                    "FROM 诊疗项目目录 A WHERE (A.撤档时间 = To_Date('30000101', 'YYYYMMDD') Or A.撤档时间 is NULL) AND A.类别='E' AND A.操作类型='6' "
            End If
               
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            If ShowGrdSelect(Me, vsf, "编码,1200,0,1;名称,3300,0,0", Me.Name & "\采集方式选择", "请从列表中选择一个采集方式。", rsData, rs, 8790, 5100) Then
                
                vsf.Cell(flexcpData, Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
                vsf.TextMatrix(Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
                vsf.TextMatrix(Row, mCol.采集方式id) = zlCommFun.NVL(rs("ID").Value)
                
                Call CreatePriceList(Row)
                
                EditChanged = True
            End If
    End Select

End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim strTmp As String
    Dim strText As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." Then
            
            If InStr(vsf.EditText, "'") > 0 Then
                KeyCode = 0
                vsf.EditText = ""
                Cancel = True
                Exit Sub
            End If
    
            Select Case Col
                Case mCol.项目名称
                    
                    strText = UCase(vsf.EditText)
                    gstrSQL = GetPublicSQL(SQL.体检项目过滤选择, strText)
                    
                    If ParamInfo.项目输入匹配方式 = 1 Then
                        strTmp = strText & "%"
                    Else
                        strTmp = "%" & strText & "%"
                    End If

                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "D", strText & "%", strTmp, 1, 2)
                    
                    If ShowGrdFilter(Me, vsf, "编码,1200,0,1;名称,2700,0,0;单位,900,0,0;类别,900,0,0", Me.Name & "\体检项目过滤选择", "请从列表中选择一个体检项目。", rsData, rs, 8790, 5100) Then

                        If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                            Exit Sub
                        End If

                        vsf.EditText = zlCommFun.NVL(rs("名称").Value)
                        vsf.TextMatrix(Row, mCol.项目类别) = zlCommFun.NVL(rs("类别").Value)
                        vsf.TextMatrix(Row, mCol.项目名称) = zlCommFun.NVL(rs("名称").Value)
                        vsf.Cell(flexcpData, Row, Col) = vsf.TextMatrix(Row, Col)
                        vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                        
                        If vsf.TextMatrix(Row, mCol.项目类别) = "检验" Then
                            Call SetDefault(Val(vsf.RowData(Row)), "执行科室", "检验标本", "采集方式", "采集执行", "计价项目")
                        Else
                            Call SetDefault(Val(vsf.RowData(Row)), "执行科室", "计价项目")
                        End If
                        
                        Call CreatePriceList(Row)
                        
                        Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.基本价格)), Val(vsf.TextMatrix(Row, mCol.体检价格)), 1)
                
                        EditChanged = True
                    Else
                        KeyCode = 0
                        Cancel = True
                        
                        vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                        vsf.EditText = vsf.Cell(flexcpData, Row, Col)
                        vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                        
                    End If
            End Select
        End If
    Else
        EditChanged = True
    End If
End Sub


Private Sub vsfPrice_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)

    Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.收费单价)), Val(vsfPrice.TextMatrix(Row, mCol.体检单价)), 1)
    
    EditChanged = True
    
End Sub

Private Sub vsfPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    Select Case Col
    Case mCol.计价项目
    
        If Left(vsfPrice.TextMatrix(Row, mCol.计价项目), 4) = "采集方式" Then
            vsfPrice.TextMatrix(Row, mCol.计价性质) = "2"
        Else
            vsfPrice.TextMatrix(Row, mCol.计价性质) = "1"
        End If
        
    Case mCol.收费数次
        vsfPrice.TextMatrix(Row, mCol.收费金额) = Val(vsfPrice.TextMatrix(Row, mCol.收费单价)) * Val(vsfPrice.TextMatrix(Row, mCol.收费数次))
        vsfPrice.TextMatrix(Row, mCol.体检金额) = Val(vsfPrice.TextMatrix(Row, mCol.体检单价)) * Val(vsfPrice.TextMatrix(Row, mCol.收费数次))
        
        Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.收费单价)), Val(vsfPrice.TextMatrix(Row, mCol.体检单价)), 1)
                
    Case mCol.体检单价
        
        Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.收费单价)), Val(vsfPrice.TextMatrix(Row, mCol.体检单价)), 1)
    
    Case mCol.p折扣
        
        Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.收费单价)), Val(vsfPrice.TextMatrix(Row, mCol.p折扣)), 2)
        
    End Select
    
    EditChanged = True
    
End Sub

Private Sub vsfPrice_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str计价项目 As String
    Dim str计价性质 As String
    
    If vsfPrice.Rows = 2 Then
        
        str计价项目 = vsfPrice.TextMatrix(1, mCol.计价项目)
        str计价性质 = vsfPrice.TextMatrix(1, mCol.计价性质)
        
        vsfPrice.Body.Cell(flexcpText, 1, mCol.计价项目 + 1, 1, vsfPrice.Cols - 1) = ""
        vsfPrice.RowData(1) = 0

        vsfPrice.TextMatrix(1, mCol.计价项目) = str计价项目
        vsfPrice.TextMatrix(1, mCol.计价性质) = str计价性质
        Call vsfPrice_AfterDeleteRow(1, Col)
        
        Cancel = True
    End If
End Sub

Private Sub vsfPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    If Col = mCol.收费项目 Then
            
        gstrSQL = GetPublicSQL(SQL.收费项目选择)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If ShowGrdSelect(Me, vsfPrice, "编码,1200,0,1;名称,2700,0,0;单位,900,0,0;单价,900,0,0;类别,900,0,0", Me.Name & "\收费项目选择", "请从列表中选择一个收费项目。", rsData, rs, 8790, 5100) Then

            If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                Exit Sub
            End If

            vsfPrice.EditText = zlCommFun.NVL(rs("名称").Value)
            vsfPrice.TextMatrix(Row, mCol.项目类别) = zlCommFun.NVL(rs("类别").Value)
            vsfPrice.TextMatrix(Row, mCol.项目名称) = zlCommFun.NVL(rs("名称").Value)
            vsfPrice.TextMatrix(Row, mCol.计算单位) = zlCommFun.NVL(rs("单位").Value)
            vsfPrice.TextMatrix(Row, mCol.收费单价) = zlCommFun.NVL(rs("单价").Value, 0)
            vsfPrice.TextMatrix(Row, mCol.体检单价) = zlCommFun.NVL(rs("单价").Value, 0)
            vsfPrice.TextMatrix(Row, mCol.收费项目id) = zlCommFun.NVL(rs("ID").Value, 0)
            
            If Val(vsfPrice.TextMatrix(Row, mCol.收费数次)) < 1 Then vsfPrice.TextMatrix(Row, mCol.收费数次) = 1
            
            vsfPrice.TextMatrix(Row, mCol.收费金额) = zlCommFun.NVL(rs("单价").Value, 0) * Val(vsfPrice.TextMatrix(Row, mCol.收费数次))
            vsfPrice.TextMatrix(Row, mCol.体检金额) = zlCommFun.NVL(rs("单价").Value, 0) * Val(vsfPrice.TextMatrix(Row, mCol.收费数次))
            
            vsfPrice.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
            
            vsf.TextMatrix(vsf.Row, mCol.基本价格) = SumPrice(mCol.收费金额)
            vsf.TextMatrix(vsf.Row, mCol.体检价格) = SumPrice(mCol.体检金额)
            
            Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.收费单价)), Val(vsfPrice.TextMatrix(Row, mCol.体检单价)), 1)
                
            EditChanged = True

        End If
        
    End If
            
End Sub

Private Sub vsfPrice_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strText As String
    Dim strTmp As String
    
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." Then
            
            If InStr(vsfPrice.EditText, "'") > 0 Then
                KeyCode = 0
                vsfPrice.EditText = ""
                Cancel = True
                Exit Sub
            End If
    
            Select Case Col
                Case mCol.收费项目
                    
                    strText = UCase(vsfPrice.EditText)
                    
                    gstrSQL = GetPublicSQL(SQL.收费项目过滤, strText)
                    
                    If ParamInfo.项目输入匹配方式 = 1 Then
                        strTmp = strText & "%"
                    Else
                        strTmp = "%" & strText & "%"
                    End If
                    
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText & "%", strTmp)
                    
                    If ShowGrdFilter(Me, vsfPrice, "编码,1200,0,1;名称,2700,0,0;单位,900,0,0;单价,900,0,0;类别,900,0,0", Me.Name & "\收费项目过滤", "请从列表中选择一个收费项目。", rsData, rs, 8790, 5100) Then
                        
                        
                        If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                            Exit Sub
                        End If
            
                        vsfPrice.EditText = zlCommFun.NVL(rs("名称").Value)
                        vsfPrice.TextMatrix(Row, mCol.项目类别) = zlCommFun.NVL(rs("类别").Value)
                        vsfPrice.TextMatrix(Row, mCol.项目名称) = zlCommFun.NVL(rs("名称").Value)
                        vsfPrice.TextMatrix(Row, mCol.计算单位) = zlCommFun.NVL(rs("单位").Value)
                        vsfPrice.TextMatrix(Row, mCol.收费单价) = zlCommFun.NVL(rs("单价").Value, 0)
                        vsfPrice.TextMatrix(Row, mCol.体检单价) = zlCommFun.NVL(rs("单价").Value, 0)
                        vsfPrice.TextMatrix(Row, mCol.收费项目id) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        If Val(vsfPrice.TextMatrix(Row, mCol.收费数次)) < 1 Then vsfPrice.TextMatrix(Row, mCol.收费数次) = 1
                        
                        vsfPrice.TextMatrix(Row, mCol.收费金额) = zlCommFun.NVL(rs("单价").Value, 0) * Val(vsfPrice.TextMatrix(Row, mCol.收费数次))
                        vsfPrice.TextMatrix(Row, mCol.体检金额) = zlCommFun.NVL(rs("单价").Value, 0) * Val(vsfPrice.TextMatrix(Row, mCol.收费数次))
                        
                        vsfPrice.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                        
                        Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.收费单价)), Val(vsfPrice.TextMatrix(Row, mCol.体检单价)), 1)
                        
                        EditChanged = True
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
        EditChanged = True
    End If
End Sub




