VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm手册内容 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "手册内容"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   Icon            =   "frm手册内容.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   7980
      TabIndex        =   2
      Top             =   150
      Width           =   1100
   End
   Begin VB.CommandButton cmd打印 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   6720
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   150
      Width           =   1100
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5475
      Index           =   0
      Left            =   30
      ScaleHeight     =   5415
      ScaleWidth      =   9315
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   660
      Width           =   9375
      Begin VB.VScrollBar vsScroll 
         Height          =   5385
         Index           =   0
         Left            =   9090
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   225
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshHead 
         Height          =   1155
         Index           =   0
         Left            =   90
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   870
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2037
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   -2147483628
         BackColorBkg    =   -2147483628
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         ScrollBars      =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDeal 
         Height          =   1155
         Index           =   0
         Left            =   90
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2010
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2037
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorFixed  =   -2147483628
         BackColorBkg    =   -2147483628
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         ScrollBars      =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin VB.Label lbl单位 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单位：元、角、分"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   7530
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "门诊特殊病费用"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   30
         TabIndex        =   4
         Top             =   180
         Width           =   9045
      End
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5475
      Index           =   1
      Left            =   30
      ScaleHeight     =   5415
      ScaleWidth      =   9315
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   660
      Width           =   9375
      Begin VB.VScrollBar vsScroll 
         Height          =   5385
         Index           =   1
         Left            =   9090
         TabIndex        =   14
         Top             =   0
         Width           =   225
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshHead 
         Height          =   1155
         Index           =   1
         Left            =   90
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   870
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2037
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   -2147483628
         BackColorBkg    =   -2147483628
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         ScrollBars      =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDeal 
         Height          =   1155
         Index           =   1
         Left            =   90
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2010
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2037
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorFixed  =   -2147483628
         BackColorBkg    =   -2147483628
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         ScrollBars      =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin VB.Label lbl单位 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单位：元、角、分"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   7560
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "住院费用"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   30
         TabIndex        =   12
         Top             =   180
         Width           =   9045
      End
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "这是您需要在病人医保手册上填写的内容："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   900
      TabIndex        =   0
      Top             =   240
      Width           =   3420
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frm手册内容.frx":1272
      Stretch         =   -1  'True
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "frm手册内容"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private mblnStartup As Boolean
Private mdbl差值 As Double
Private mblnOutPatient As Boolean       '门诊
Private mrsHead As New ADODB.Recordset
Private mrsDeal As New ADODB.Recordset

Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

Private Enum 页面
    门诊 = 0
    住院
End Enum

Private Sub cmd打印_Click()
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim intIndex As Integer
    Dim bytMode As Byte
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    
    bytMode = 1
    intIndex = IIf(mblnOutPatient, 门诊, 住院)
    
    Set objPrint = New zlPrintGrds
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = Trim(lblTitle(intIndex).Caption)
        
    objRow.Add lbl单位(intIndex).Caption
    objPrint.UnderAppRows.Add objRow
    
    Set objPrint.Grds = New Collection
    objPrint.Grds.Add mshHead(intIndex)
    objPrint.Grds.Add mshDeal(intIndex)
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewGrds objPrint, 1
          Case 2
              zlPrintOrViewGrds objPrint, 2
          Case 3
              zlPrintOrViewGrds objPrint, 3
      End Select
    Else
        zlPrintOrViewGrds objPrint, bytMode
    End If
End Sub

Private Sub cmd确定_Click()
    Unload Me
End Sub

Public Sub ShowBalance(ByVal rsHead As ADODB.Recordset, ByVal rsDeal As ADODB.Recordset, Optional ByVal bln门诊 As Boolean = True)
    mblnOutPatient = bln门诊
    Set mrsHead = rsHead
    Set mrsDeal = rsDeal
    Me.Show 1
End Sub

Private Sub Form_Activate()
    If Not mblnStartup Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim dbl高度 As Double
    Dim objHead As MSHFlexGrid
    Dim objDeal As MSHFlexGrid
    
    '基础设置
    mdbl差值 = 0
    picBack(IIf(mblnOutPatient, 门诊, 住院)).Visible = True
    picBack(IIf(mblnOutPatient, 门诊, 住院)).ZOrder
    Set objHead = IIf(mblnOutPatient, mshHead(门诊), mshHead(住院))
    Set objDeal = IIf(mblnOutPatient, mshDeal(门诊), mshDeal(住院))
    Call InitStruct
    Call LoadData
    Call SetRowHeight(objHead)
    Call SetRowHeight(objDeal)
    
    '调整位置
    objHead.Height = objHead.Rows * 700
    objDeal.Height = objDeal.Rows * 700
    objDeal.Top = objHead.Top + objHead.Height
    mblnStartup = True
    If mblnOutPatient Then Exit Sub
    
    vsScroll(住院).Visible = (objDeal.Top + objDeal.Height > picBack(住院).Height)
    With vsScroll(住院)
        .Value = 0
        .Min = 0
        .Max = (mshDeal(住院).Top + mshDeal(住院).Height) / picBack(住院).Height
        .LargeChange = 1
    End With
    
    dbl高度 = mshDeal(住院).Top + mshDeal(住院).Height
    '计算差值
    mdbl差值 = dbl高度 / (vsScroll(住院).Max + 1)
End Sub

Private Sub InitStruct()
    Dim arrHead, arrDeal
    Dim intCol As Long, intCols As Integer
    Dim strHead As String, strDeal As String
    Dim objHead As MSHFlexGrid
    Dim objDeal As MSHFlexGrid
    Const strHead_门诊 As String = "医院名称,2500|就诊日期,2000|医院级别,1500|初步诊断,2900"
    Const strHead_住院 As String = "医院名称,2400|入院-出院日期（年、月、日）,2500|医院级别,1200|初步" & vbCrLf & "诊断,500|入院类型,800|中途转院" & vbCrLf & "转出日期,1500"
    Const strDeal_门诊 As String = "费用总额,1200|统筹支付,1200|大额/" & vbCrLf & "公务员支付,1200|个人自付,1600|个人自费,1200|统筹封顶后" & vbCrLf & "医保内金额,1200|日期、" & vbCrLf & "经办签章,1300"
    Const strDeal_住院 As String = "费用总额,1200|统筹支付,1200|大额/" & vbCrLf & "公务员支付,1200|个人自付,1600|个人自费,1200|统筹封顶后" & vbCrLf & "医保内金额,1200|日期、" & vbCrLf & "经办签章,1300"
    
    If mblnOutPatient Then
        strHead = strHead_门诊
        strDeal = strDeal_门诊
        Set objHead = mshHead(门诊)
        Set objDeal = mshDeal(门诊)
    Else
        strHead = strHead_住院
        strDeal = strDeal_住院
        Set objHead = mshHead(住院)
        Set objDeal = mshDeal(住院)
    End If
    
    '设置表头
    arrHead = Split(strHead, "|")
    intCols = UBound(arrHead)
    objHead.Cols = intCols + 1
    For intCol = 0 To intCols
        objHead.TextMatrix(0, intCol) = Split(arrHead(intCol), ",")(0)
        objHead.ColWidth(intCol) = Split(arrHead(intCol), ",")(1)
        objHead.ColAlignmentFixed(intCol) = 4
        objHead.ColAlignment(intCol) = IIf(intCol = intCols, 7, 1)
    Next
    '设置待遇表格
    arrDeal = Split(strDeal, "|")
    intCols = UBound(arrDeal)
    objDeal.Cols = intCols + 1
    For intCol = 0 To intCols
        objDeal.TextMatrix(0, intCol) = Split(arrDeal(intCol), ",")(0)
        objDeal.ColWidth(intCol) = Split(arrDeal(intCol), ",")(1)
        objDeal.ColAlignmentFixed(intCol) = 4
        objDeal.ColAlignment(intCol) = IIf(intCol = 3, 1, 7)
    Next
End Sub

Private Sub LoadData()
    Dim objMsh As MSHFlexGrid
    '根据记录集的数据显示
    If mblnOutPatient Then
        '门诊只可能有一条记录
        Set objMsh = mshDeal(门诊)
        With mshHead(门诊)
            .TextMatrix(1, 0) = Nvl(mrsHead!医院名称)
            .TextMatrix(1, 1) = Nvl(mrsHead!就诊日期)
            .TextMatrix(1, 2) = Nvl(mrsHead!医院级别)
            .TextMatrix(1, 3) = Nvl(mrsHead!初步诊断)
        End With
    Else
        Set objMsh = mshDeal(住院)
        With mshHead(住院)
            If mrsHead.RecordCount <> 0 Then mrsHead.MoveFirst
            Do While Not mrsHead.EOF
                If mrsHead.AbsolutePosition > 1 Then .Rows = .Rows + 1
                .TextMatrix(mrsHead.AbsolutePosition, 0) = Nvl(mrsHead!医院名称)
                .TextMatrix(mrsHead.AbsolutePosition, 1) = Nvl(mrsHead!入院日期) & "-" & Nvl(mrsHead!转出日期)
                .TextMatrix(mrsHead.AbsolutePosition, 2) = Nvl(mrsHead!医院级别)
                .TextMatrix(mrsHead.AbsolutePosition, 3) = Nvl(mrsHead!初步诊断)
                .TextMatrix(mrsHead.AbsolutePosition, 4) = Nvl(mrsHead!入院类型)
                .TextMatrix(mrsHead.AbsolutePosition, 5) = Nvl(mrsHead!转出日期)
                mrsHead.MoveNext
            Loop
        End With
    End If
    
    '统一将待遇信息写入（格式一样）
    With objMsh
        If mrsDeal.RecordCount <> 0 Then mrsDeal.MoveFirst
        Do While Not mrsDeal.EOF
            If mrsDeal.AbsolutePosition > 1 Then .Rows = .Rows + 1
            .TextMatrix(mrsDeal.AbsolutePosition, 0) = Format(Nvl(mrsDeal!费用总额, 0), "#0.00;-#0.00;0.00;")
            .TextMatrix(mrsDeal.AbsolutePosition, 1) = Format(Nvl(mrsDeal!统筹支付, 0), "#0.00;-#0.00;0.00;")
            .TextMatrix(mrsDeal.AbsolutePosition, 2) = Format(Nvl(mrsDeal!大额支付, 0), "#0.00;-#0.00;0.00;")
            .TextMatrix(mrsDeal.AbsolutePosition, 3) = "自付1：" & Format(Nvl(mrsDeal!个人自付, 0), "#0.00;-#0.00;0.00;") & _
                vbCrLf & "自付2：" & Format(Nvl(mrsDeal!首先自付, 0), "#0.00;-#0.00;0.00;")
            .TextMatrix(mrsDeal.AbsolutePosition, 4) = Format(Nvl(mrsDeal!个人自费, 0), "#0.00;-#0.00;0.00;")
            .TextMatrix(mrsDeal.AbsolutePosition, 5) = Format(Nvl(mrsDeal!统筹封顶后医保内金额, 0), "#0.00;-#0.00;0.00;")
            .TextMatrix(mrsDeal.AbsolutePosition, 6) = Nvl(mrsDeal!经办日期)
            mrsDeal.MoveNext
        Loop
    End With
End Sub

Private Sub SetRowHeight(ByVal objMsh As MSHFlexGrid)
    Dim intRow As Integer, intRows As Integer
    intRows = objMsh.Rows - 1
    For intRow = 0 To intRows
        objMsh.RowHeight(intRow) = 700
    Next
End Sub

Private Sub vsScroll_Change(Index As Integer)
    Static intValue As Integer          '上次的值
    Dim intCur As Integer               '当前的值
    Dim dbl差值 As Double
    If Index = 门诊 Then Exit Sub
    
    intCur = vsScroll(Index).Value
    dbl差值 = mdbl差值 * (intValue - intCur)
    intValue = intCur
    
    '移动所有控件
    picBack(Index).AutoRedraw = False
    lblTitle(Index).Top = lblTitle(Index).Top + dbl差值
    lbl单位(Index).Top = lbl单位(Index).Top + dbl差值
    mshHead(Index).Top = mshHead(Index).Top + dbl差值
    mshDeal(Index).Top = mshDeal(Index).Top + dbl差值
    picBack(Index).AutoRedraw = True
End Sub
