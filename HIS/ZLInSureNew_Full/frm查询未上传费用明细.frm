VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm查询未上传费用明细 
   Caption         =   "未上传的处方明细清单，请检查"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   Icon            =   "frm查询未上传费用明细.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   8910
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
      Left            =   7620
      TabIndex        =   0
      Top             =   4920
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   4785
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
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
      MouseIcon       =   "frm查询未上传费用明细.frx":0E42
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm查询未上传费用明细"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mintInsure As Integer

Public Function ShowME(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intinsure As Integer) As Boolean
    '如果还存在未上传的处方明细则返回假，同时显示给操作员供检查
    On Error Resume Next
    mblnOK = False
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mintInsure = intinsure
    Me.Show 1
    ShowME = mblnOK
End Function

Private Sub cmdExcel_Click()
    '输出到EXCEL
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    Dim bytStyle As Byte
    
    intRow = mshList.Row
    bytStyle = 3
    
    '表头
    objOut.Title.Text = "未上传的处方明细清单-" & mlng病人ID
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
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    '提取指定病人指定住院的所有未上传明细
    gstrSQL = "Select DECODE(A.记录性质,3,'自动记帐','记帐') AS 单据,DECODE(A.记录状态,2,'冲销','正常') AS 类型,A.No,A.序号,E.名称," & _
             " trim(to_char(Nvl(A.数次,0)*Nvl(A.付数,1),'90009990.00')) As 数量,trim(to_char(A.标准单价,'90009990.00')) AS 标准单价," & _
             " trim(to_char(A.实收金额,'90009990.00')) AS 实收金额,F.项目编码 AS 医保编码" & _
             " From 住院费用记录 A,病人信息 B,病案主页 C,保险帐户 D,收费细目 E,保险支付项目 F " & _
             " Where A.病人ID=B.病人ID And B.病人ID=C.病人ID And A.主页ID=C.主页ID And A.病人ID=D.病人ID And D.险类=" & mintInsure & _
             " And Nvl(记录状态,0)<>0 And Nvl(附加标志,0)<>9 And Nvl(实收金额,0)<>0 And Nvl(A.是否上传,0)=0 And Nvl(记帐费用,0)=1 " & _
             " And (Nvl(A.门诊标志,0)<>1 And Nvl(A.门诊标志,0)<>4)" & _
             " And A.收费细目ID=E.Id And E.ID=F.收费细目ID(+) And F.险类(+)=" & mintInsure & _
             " And A.病人ID=[1] And A.主页ID=[2]" & _
             " Order By A.登记时间,No,序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取指定病人指定住院的所有未上传明细", mlng病人ID, mlng主页ID)
    If rsTemp.RecordCount = 0 Then
        mblnOK = True
        Unload Me
        Exit Sub
    End If
    
    Set mshList.DataSource = rsTemp
    With mshList
        .ColWidth(0) = 660
        .ColWidth(1) = 495
        .ColWidth(2) = 810
        .ColWidth(3) = 495
        .ColWidth(4) = 2070
        .ColWidth(5) = 1035
        .ColWidth(6) = 1035
        .ColWidth(7) = 990
        .ColWidth(8) = 1200
        .ColAlignment(4) = 1
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
        .ColAlignment(7) = 7
        .ColAlignment(8) = 1
    End With
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
