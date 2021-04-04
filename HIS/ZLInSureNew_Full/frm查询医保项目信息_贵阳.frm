VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm查询医保项目信息_贵阳 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查询医保项目信息"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "frm查询医保项目信息_贵阳.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd定位 
      Caption         =   "定位(&L)"
      Height          =   350
      Left            =   2730
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "在查询的结果集上定位某个项目"
      Top             =   120
      Width           =   1100
   End
   Begin VB.CommandButton cmd查询 
      Caption         =   "查询(&R)"
      Height          =   350
      Left            =   7140
      TabIndex        =   5
      ToolTipText     =   "从中心提取指定支付类别下所有项目或某个项目的报销信息"
      Top             =   120
      Width           =   1100
   End
   Begin VB.ComboBox cbo支付类别 
      Height          =   300
      Left            =   5070
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   150
      Width           =   1905
   End
   Begin VB.TextBox txt编码 
      Height          =   300
      Left            =   1020
      TabIndex        =   1
      Top             =   150
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   4695
      Left            =   90
      TabIndex        =   6
      Top             =   600
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorSel    =   12285290
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin VB.Label lbl支付类别 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "支付类别(&P)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3990
      TabIndex        =   3
      Top             =   210
      Width           =   990
   End
   Begin VB.Label lbl编码 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "编码(&B)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   630
   End
End
Attribute VB_Name = "frm查询医保项目信息_贵阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Enum Columns
    编码
    名称
    最高限价
    自付比例
    生育标志
    工伤标志
    特殊报销标志    '01-普通项目；02-政策内全自付项目(纳入医疗补助范围)；03-医疗照顾人员特殊项目；04-基金直接支付项目
    包干结算类别    '01-普通项目；02-包干结算加收范围内项目；03-医疗照顾人员特殊项目； 04-基金直接支付项目；05-包干结算加收自费项目
    列数
End Enum

Private Sub cmd查询_Click()
    Dim arrData
    Dim str类别 As String
    Dim rsData As New ADODB.Recordset
    '查询指定项目或所有项目其它支付类别的信息
    On Error GoTo errHand
    
    With rsData
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .Fields.Append "CLASSCODE", adVarChar, 6    '大类编码
        .Fields.Append "CODE", adVarChar, 20        '编码
        .Fields.Append "NAME", adVarChar, 300       '名称
        .Fields.Append "PY", adVarChar, 150         '拼音简码
        .Fields.Append "MEMO", adVarChar, 500       '附注
        .Open
    End With
    Call InitBill(True)
    
    str类别 = cbo支付类别.ItemData(cbo支付类别.ListIndex)
    If Not 医保项目_贵阳(rsData, str类别) Then Exit Sub
    
    '装入数据
    mshDetail.Redraw = False
    With rsData
        Do While Not .EOF
            arrData = Split(!Memo, "|")
            mshDetail.TextMatrix(.AbsolutePosition, 编码) = !CODE
            mshDetail.TextMatrix(.AbsolutePosition, 名称) = !Name
            mshDetail.TextMatrix(.AbsolutePosition, 最高限价) = arrData(0)
            mshDetail.TextMatrix(.AbsolutePosition, 自付比例) = arrData(1)
            mshDetail.TextMatrix(.AbsolutePosition, 生育标志) = arrData(2)
            mshDetail.TextMatrix(.AbsolutePosition, 工伤标志) = arrData(3)
            mshDetail.TextMatrix(.AbsolutePosition, 特殊报销标志) = Switch(arrData(4) = "01", "普通项目", arrData(4) = "02", "政策内全自付项目(纳入医疗补助范围)", arrData(4) = "03", "医疗照顾人员特殊项目", arrData(4) = "04", "基金直接支付项目")
            mshDetail.TextMatrix(.AbsolutePosition, 包干结算类别) = Switch(arrData(5) = "01", "普通项目", arrData(5) = "02", "包干结算加收范围内项目", arrData(5) = "03", "医疗照顾人员特殊项目", arrData(5) = "04", "基金直接支付项目", arrData(5) = "05", "包干结算加收自费项目")
            .MoveNext
        Loop
    End With
    mshDetail.Redraw = True
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mshDetail.Redraw = True
End Sub

Private Sub cmd定位_Click()
    Dim intDO As Integer, intMAX As Integer
    
    If txt编码.Text = "" Then Exit Sub
    intMAX = mshDetail.Rows - 1
    For intDO = 1 To intMAX
        If txt编码.Text = (mshDetail.TextMatrix(intDO, 编码)) Then
            With mshDetail
                .TopRow = intDO
                .Row = intDO: .RowSel = intDO
                .COL = 0: .ColSel = .Cols - 1
            End With
            Exit Sub
        End If
    Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With cbo支付类别
        .Clear
        .AddItem "普通住院"
        .ItemData(.NewIndex) = 12
        .AddItem "离休住院"
        .ItemData(.NewIndex) = 22
        .AddItem "工伤住院"
        .ItemData(.NewIndex) = 42
        .AddItem "生育住院"
        .ItemData(.NewIndex) = 32
        .AddItem "普通门诊"
        .ItemData(.NewIndex) = 11
        .AddItem "离休门诊"
        .ItemData(.NewIndex) = 21
        .AddItem "工伤门诊"
        .ItemData(.NewIndex) = 41
        .AddItem "生育门诊"
        .ItemData(.NewIndex) = 31
        .ListIndex = 0
    End With
End Sub

Private Sub InitBill(Optional ByVal blnInit As Boolean)
    With mshDetail
        If blnInit Then
            .Clear
            .Rows = 2: .Cols = 列数
            .TextMatrix(0, 编码) = "编码"
            .TextMatrix(0, 名称) = "名称"
            .TextMatrix(0, 最高限价) = "最高限价"
            .TextMatrix(0, 自付比例) = "自付比例"
            .TextMatrix(0, 生育标志) = "生育标志"
            .TextMatrix(0, 工伤标志) = "工伤标志"
            .TextMatrix(0, 特殊报销标志) = "特殊报销标志"
            .TextMatrix(0, 包干结算类别) = "包干结算类别"
        End If
        .ColWidth(编码) = 1000
        .ColWidth(名称) = 1500
        .ColWidth(最高限价) = 1000
        .ColWidth(自付比例) = 1000
        .ColWidth(生育标志) = 500
        .ColWidth(工伤标志) = 500
        .ColWidth(特殊报销标志) = 1500
        .ColWidth(包干结算类别) = 1500
    End With
End Sub
