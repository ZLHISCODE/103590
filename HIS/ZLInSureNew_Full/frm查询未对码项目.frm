VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm查询未对码项目 
   Caption         =   "查询未对码项目"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8280
   Icon            =   "frm查询未对码项目.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   8280
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdExcel 
      Caption         =   "输出&EXCEL"
      Height          =   350
      Left            =   150
      TabIndex        =   2
      Top             =   4920
      Width           =   1100
   End
   Begin VB.CommandButton cmd退出 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Default         =   -1  'True
      Height          =   350
      Left            =   6900
      TabIndex        =   1
      Top             =   4920
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   4785
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   8440
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frm查询未对码项目.frx":0E42
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm查询未对码项目"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure As Integer

Public Sub ShowME(ByVal objParent As Object, ByVal intinsure As Integer)
    mintInsure = intinsure
    Me.Show 1, objParent
End Sub

Private Sub cmdExcel_Click()
    '输出到EXCEL
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    Dim bytStyle As Byte
    
    intRow = mshList.Row
    bytStyle = 3
    
    '表头
    objOut.Title.Text = "未对码项目清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    mshList.Redraw = False
    Set objOut.Body = mshList
    
    '输出
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mshList.Row = intRow
    mshList.COL = 0: mshList.ColSel = mshList.Cols - 1
    mshList.Redraw = True
End Sub

Private Sub cmd退出_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    '显示指定医保所有未对码的项目
    '格式：收费细目ID|收费类别|项目编码|项目名称|单价|规格|建档时间
    Dim i As Integer, j As Integer
    Dim rsTemp As New ADODB.Recordset
    
    With mshList
        .Cols = 6
        .TextMatrix(0, 0) = "收费细目ID"
        .TextMatrix(0, 1) = "收费类别"
        .TextMatrix(0, 2) = "项目编码"
        .TextMatrix(0, 3) = "项目名称"
        .TextMatrix(0, 4) = "规格"
        .TextMatrix(0, 5) = "建档时间"
        
        j = .Cols - 1
        For i = 0 To j
            .ColAlignmentFixed(i) = 4
        Next
        .ColAlignment(2) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 1
        .ColWidth(0) = 0
        .ColWidth(1) = 810
        .ColWidth(2) = 1050
        .ColWidth(3) = 2130
        .ColWidth(4) = 2820
    End With
    
    '读取所有未对码项目
    gstrSQL = "Select C.ID As 收费细目ID,B.类别 As 收费类别,C.编码 As 项目编码,C.名称 As 项目名称,DECODE(C.规格,'┆','',C.规格) AS 规格,C.建档时间 " & _
             " From  " & _
             " (Select ID As 收费细目ID " & _
             " From 收费细目 " & _
             " Minus  " & _
             " Select 收费细目ID " & _
             " From 保险支付项目 " & _
             " Where 险类=[1]) A,收费类别 B,收费细目 C " & _
             " Where A.收费细目ID=C.Id And B.编码=C.类别 " & _
             " And (C.撤档时间 Is NULL Or to_char(C.撤档时间,'yyyy-MM-dd')='3000-01-01')" & _
             " Order By C.类别,C.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取所有未对码项目", mintInsure)
    If rsTemp.RecordCount = 0 Then Exit Sub
    Set mshList.DataSource = rsTemp
    mshList.ColWidth(0) = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    With cmd退出
        .Left = Me.ScaleWidth - .Width - 150
        .Top = Me.ScaleHeight - .Height - 150
    End With
    cmdExcel.Top = cmd退出.Top
    
    With mshList
        .Height = cmd退出.Top - 150
        .Width = Me.ScaleWidth
    End With
End Sub
