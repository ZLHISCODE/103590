VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPreview 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "preView"
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form24"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtLength 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1005
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   90
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1995
      LargeChange     =   10
      Left            =   3540
      Max             =   100
      SmallChange     =   2
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   285
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   90
      ScaleHeight     =   2655
      ScaleWidth      =   3225
      TabIndex        =   0
      Top             =   630
      Width           =   3255
      Begin VSFlex8Ctl.VSFlexGrid VsfData 
         Height          =   1455
         Left            =   570
         TabIndex        =   5
         Top             =   930
         Width           =   2265
         _cx             =   3995
         _cy             =   2566
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   3
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPreview.frx":0000
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblDownTable 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "表下项可换行"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   4
         Top             =   1020
         Width           =   1125
      End
      Begin VB.Label lblUpTable 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "表上项可换行"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   270
         TabIndex        =   3
         Top             =   600
         Width           =   1125
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "标题栏"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1380
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Line lineRight 
         X1              =   1380
         X2              =   1380
         Y1              =   360
         Y2              =   2220
      End
      Begin VB.Line lineLeft 
         X1              =   720
         X2              =   720
         Y1              =   360
         Y2              =   2220
      End
      Begin VB.Line lineBottom 
         X1              =   630
         X2              =   2790
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line lineTop 
         X1              =   630
         X2              =   2790
         Y1              =   600
         Y2              =   600
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objStream As TextStream
Dim lngFormat As Long               '格式ID
Dim lngFile As Long                 '病人护理文件.ID
Dim mlngRows As Long
Dim mstrSQL As String
Dim mstrSQL中 As String
Dim mstrSQL内 As String
Dim mstrSQL列 As String
Dim mstrSQL条件 As String

'病历文件格式定义相关
Private mintTabTiers As Integer     '表头层次
Private mintTagFormHour As Integer  '开始时间条件
Private mintTagToHour As Integer    '截止时间条件
Private mobjTagFont As New StdFont  '条件样式字体
Private mlngTagColor As Long        '条件样式颜色
Private mstrPaperSet As String      '格式
Private mstrPageHead As String      '页眉
Private mstrPageFoot As String      '页脚
Private mblnChildForm As Boolean
Private mlngActiveRows As Long      '有效数据行
Private mstrSubhead As String       '表上标签
Private mstrTabHead As String       '表头单元
Private mstrPreHead As String       '需处理的列,文本型项目所属列或绑定多个项目的列
Private mstrColWidth As String      '列宽序列串
Private mstrColumns As String       '当前护理文件各列对应的项目
Private lngCurColor As Long, strCurFont As String, objFont As StdFont
Private mrsItems As New ADODB.Recordset

Dim dblTitle As Double      '标题栏的高度
Dim dblUpTable As Double    '表上项的高度
Dim dblDownTable As Double  '表下项的高度

Private Const EM_GETLINECOUNT = &HBA&        '获取行数。
Private Const EM_GETLINE = &HC4&             '发送一行文本到指定缓冲区。
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub Form_Load()
    Dim lngRows As Long                             '计算所得有效行数
    Dim lngFixRows As Long                          '固定行数
    Dim dblRowHeight As Double                      '行高
    Dim lngParent As Long
    Dim strUpText As String
    Dim lngHeight As Long, lngWidth As Long         '有效高度，宽度
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    Dim rsTemp As New ADODB.Recordset
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    
    '设置页面格式
    Call zlGetPrinterSet
    
    '获取打印机当前状态
    picDraw.Height = Printer.Height
    picDraw.Width = Printer.Width
    picDraw.ScaleHeight = Printer.ScaleHeight
    picDraw.ScaleWidth = Printer.ScaleWidth
    '页边距
    lngTop = marrFormat(6)
    lngBottom = marrFormat(7)
    lngLeft = marrFormat(4)
    lngRight = marrFormat(5)
    '实际有效高度，宽度
    lngHeight = picDraw.ScaleHeight - lngTop - lngBottom
    lngWidth = picDraw.ScaleWidth - lngLeft - lngRight
    
    '上,下边距(lngTop , lngBottom)
    '左,右边距(lngLeft , lngRight)
    lineTop.X1 = 0
    lineTop.X2 = picDraw.ScaleWidth
    lineTop.Y1 = lngTop
    lineTop.Y2 = lngTop
    lineBottom.X1 = 0
    lineBottom.X2 = picDraw.ScaleWidth
    lineBottom.Y1 = picDraw.ScaleHeight - lngBottom
    lineBottom.Y2 = lineBottom.Y1
    
    lineLeft.X1 = lngLeft
    lineLeft.X2 = lngLeft
    lineLeft.Y1 = 0
    lineLeft.Y2 = picDraw.ScaleHeight
    lineRight.X1 = picDraw.ScaleWidth - lngRight
    lineRight.X2 = lineRight.X1
    lineRight.Y1 = 0
    lineRight.Y2 = picDraw.ScaleHeight
    
    '准备处理表格内容,包括标题栏,表下标签及表体
    gstrSQL = "" & _
            "SELECT id, 文件id, nvl(父id,0) 父ID, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, " & vbNewLine & _
            "       预制提纲id, 复用提纲, 使用时机, 诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
            "FROM 病历文件结构 A" & vbNewLine & _
            "WHERE A.文件ID=[1]" & vbNewLine & _
            "ORDER BY A.父ID,A.对象序号"
    Set rsTemp = OpenSQLRecord(gstrSQL, "提取护理文件定义", lngFormat)
    
    With rsTemp
        '1、标题栏从上边距开始
        .Filter = "对象序号=1 And 父ID=0"
        lngParent = !ID
        '行高
        .Filter = "父ID=" & lngParent & " And 对象序号=3"
        dblRowHeight = !内容文本
        '固定行数
        .Filter = "父ID=" & lngParent & " And 对象序号=1"
        lngFixRows = !内容文本
        '标题栏的字体设置
        .Filter = "父ID=" & lngParent & " And 对象序号=8"
        lblTitle.FontName = Split(!内容文本, ",")(0)
        lblTitle.FontSize = Split(!内容文本, ",")(1)
        .Filter = "父ID=" & lngParent & " And 对象序号=7"
        lblTitle.Caption = !内容文本
        '表上项的字体设置
        .Filter = "父ID=" & lngParent & " And 对象序号=4"
        lblUpTable.FontName = Split(!内容文本, ",")(0)
        lblUpTable.FontSize = Split(!内容文本, ",")(1)
        .Filter = "父ID=" & lngParent & " And 对象序号=5"
        lblUpTable.BackColor = !内容文本
        
        '设置标题栏坐标
        picDraw.FontName = lblTitle.FontName
        picDraw.FontSize = lblTitle.FontSize
        lblTitle.Left = lngLeft
        lblTitle.Top = lngTop + 30
        lblTitle.Width = lngWidth
        lblTitle.Height = picDraw.TextHeight("a")
        
        '2、表上标签从标题栏下开始
        .Filter = "对象序号=2 And 父ID=0"
        lngParent = !ID
        .Filter = "父ID=" & lngParent
        Do While Not .EOF
            strUpText = strUpText & IIf(strUpText = "", "", "  ") & IIf(!是否换行 = 0, "", vbCrLf) & NVL(!内容文本) & !要素名称
            .MoveNext
        Loop
        If strUpText <> "" Then
            lblUpTable.Caption = strUpText
            lblUpTable.AutoSize = True
        End If
        '设置表上项坐标
        picDraw.FontName = lblUpTable.FontName
        picDraw.FontSize = lblUpTable.FontSize
        lblUpTable.Left = lngLeft
        lblUpTable.Top = lblTitle.Top + lblTitle.Height + 30
        lblUpTable.Width = picDraw.ScaleWidth
        
        '3、设置表格
    lngHeight = lngHeight - lblUpTable.Height - lblTitle.Height
        VsfData.Top = lblUpTable.Top + lblUpTable.Height + 30
        VsfData.Left = lngLeft
        VsfData.Width = lngWidth
        lngHeight = lngHeight + lngTop - VsfData.Top
        VsfData.Height = lngHeight
    lngRows = CLng(lngHeight \ dblRowHeight) - lngFixRows
        VsfData.Rows = lngFixRows + lngRows
        VsfData.FixedRows = lngFixRows
        VsfData.RowHeightMin = dblRowHeight
        
        mlngRows = lngRows
    End With
    
    Call VScroll1_Change
    
    If mrsItems.State = 0 Then
        '打开现存在的所有护理记录项目
        gstrSQL = " Select 项目序号,项目名称,项目类型,项目性质,项目长度,项目小数,项目表示,项目单位,项目值域,护理等级,应用方式" & _
                  " From 护理记录项目 B" & _
                  " Where B.应用方式<>0 " & _
                  " Order by 项目序号"
        Set mrsItems = OpenSQLRecord(gstrSQL, "打开现存在的所有护理记录项目")
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
End Sub

Private Sub VScroll1_Change()
    picDraw.Top = -1 * VScroll1.Value * (picDraw.Height - Me.Height) / 100
End Sub

Public Function ShowMe(ByVal objParent As Object, ByVal lngFileID As Long, arrData, Optional ByVal blnHide As Boolean = True) As Long
    '读取护理记录单的格式
    mlngRows = 0
    lngFormat = lngFileID
    marrFormat = arrData
    If blnHide Then
        Unload frmPreview
        Load frmPreview
    Else
        Me.Show 1, objParent
    End If
    ShowMe = mlngRows
End Function

Public Function AnaliseData(ByVal objParent As Object, ByVal lngFileID As Long, arrData, objStream_ As TextStream) As Boolean
    lngFormat = lngFileID
    marrFormat = arrData
    Set objStream = objStream_
    
    Unload frmPreview
    Load frmPreview
    
    If Not ReadStruDef Then
        '没有需要解析的列,因此直接返回解析成功,应该不存在这种情况
        AnaliseData = (mstrPreHead = "")
        Exit Function
    End If
    If Not ReadData Then Exit Function
    
    AnaliseData = True
End Function

Private Function ReadData() As Boolean
    Dim rsPati As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '读取所有使用该记录文件的病人列表
    gstrSQL = "Select ID,科室ID,病人ID,主页ID,婴儿 from 病人护理文件 where nvl(解析,0)=0 And 格式ID=[1]"
    Set rsPati = OpenSQLRecord(gstrSQL, "提取使用该护理文件的病人列表", lngFormat)
    Do While Not rsPati.EOF
        '装入数据
        lngFile = rsPati!ID
        gstrSQL = mstrSQL
        Set rsTemp = OpenSQLRecord(gstrSQL, "提取护理数据", CLng(rsPati!ID), CLng(rsPati!病人ID), CLng(rsPati!主页ID), CLng(rsPati!婴儿))
        '绑定数据并设置护理记录单的格式,同时实现一行数据分行显示的功能
        Call PreTendFormat(rsTemp)
        '解析每行数据
        If Not ParseData Then Exit Function
        
        gcnOracle.Execute "ZL_病人护理文件_解析(" & rsPati!ID & ")", , adCmdStoredProc
        'objStream.WriteLine "文件ID:" & rsPati!ID & "，病人ID=" & rsPati!病人ID & ";主页ID=" & rsPati!主页ID & ";婴儿=" & rsPati!婴儿 & "于" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "完成..."
        
        If gintAutoRUN = 1 Then
            If Format(Now, "HH:mm") >= gstrEndTime Then
                Exit Do
            End If
        End If
        rsPati.MoveNext
    Loop
    
    ReadData = True
    Exit Function
errHand:
    MsgBox Err.Description
End Function

Private Function ParseData() As Boolean
    Dim arrCol, arrData
    Dim blnNewPage As Boolean
    Dim lngMutilRow As Long, lngRecord As Long
    Dim lngRow As Long, lngCount As Long
    Dim lngCol As Long, lngMAX As Long
    Dim lngStartPage As Long, lngEndPage As Long, lngStartRow As Long, lngEndRow As Long
    On Error GoTo errHand
    '循环解析所有行数据(一列绑定多个项目,或者项目为文本型)
    
    arrCol = Split(mstrPreHead, ",")
    lngMAX = UBound(arrCol)
    lngCount = VsfData.Rows - 1
    lngStartPage = 1: lngEndPage = 1: lngStartRow = 1: lngEndRow = 1
    
    For lngRow = 1 To lngCount
        lngMutilRow = 0
        lngRecord = Val(VsfData.TextMatrix(lngRow, VsfData.Cols - 1))
        If lngRecord <> 0 Then
            For lngCol = 0 To lngMAX
                If VsfData.TextMatrix(lngRow, arrCol(lngCol)) <> "" Then
                '准备赋值
                With txtLength
                    .Width = VsfData.ColWidth(arrCol(lngCol))
                    .Text = VsfData.TextMatrix(lngRow, arrCol(lngCol))
                    .FontName = VsfData.FontName
                    .FontSize = VsfData.FontSize
                End With
                arrData = GetData(txtLength.Text)
                If UBound(arrData) > lngMutilRow Then lngMutilRow = UBound(arrData)
                End If
            Next
            
            lngEndRow = (lngStartRow + lngMutilRow)
reSub:
            If lngEndRow > mlngActiveRows Then
                blnNewPage = True
                lngEndRow = lngEndRow - mlngActiveRows
                lngEndPage = lngEndPage + 1
                GoTo reSub
            End If
            
            '一行结束时，产生打印解析数据
            gstrSQL = "ZL_病人护理打印升迁_UPDATE(" & lngRecord & "," & lngFile & "," & lngMutilRow + 1 & "," & lngStartPage & "," & lngStartRow & "," & lngEndPage & "," & lngEndRow & ")"
            gcnOracle.Execute gstrSQL, , adCmdStoredProc
            
            If blnNewPage Then
                lngStartPage = lngEndPage
                blnNewPage = False
            End If
            lngStartRow = lngEndRow + 1
            If lngStartRow > mlngActiveRows Then
                lngStartRow = lngStartRow - mlngActiveRows
                lngStartPage = lngStartPage + 1
                If lngEndPage < lngStartPage Then lngEndPage = lngStartPage
            End If
        End If
    Next
    
    ParseData = True
    Exit Function
errHand:
    MsgBox Err.Description
End Function

Private Sub PreTendFormat(ByVal rsTemp As ADODB.Recordset)
    Dim blnTag As Boolean
    Dim aryItem() As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String
    On Error GoTo errHand
    
    '设置护理记录单的格式
    With VsfData
        .FixedRows = 3
        .Clear
        Set .DataSource = rsTemp
        
        '表头填写
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        .ColHidden(.Cols - 1) = True
        .ColHidden(.Cols - 2) = True
        .ColHidden(.Cols - 3) = True
        
        '设置列头
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + 1) = strCell
        Next
        
        '列宽设置
        Dim blnAlign As Boolean
        aryItem = Split(mstrColWidth, ",")
        For lngCount = 2 To .Cols - 1
            If Not .ColHidden(lngCount) Then
                .ColWidth(lngCount) = Val(Split(aryItem(lngCount - 2), "`")(0))
                If InStr(1, aryItem(lngCount - 2), "`") <> 0 Then
                    blnAlign = True
                    .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - 2), "`")(1))
                End If
            End If
        Next
        
        '固定行格式为居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        '再按列合并
        For lngCount = 0 To VsfData.Cols - 1
            VsfData.MergeCol(lngCount) = True
        Next
        .AutoSize 0, .Cols - 1
        
        If blnAlign = False Then
            '改为根据用户的设置显示列对齐方式
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .RowHeight(lngCount) < .RowHeightMin Then .RowHeight(lngCount) = .RowHeightMin
        Next
        Select Case mintTabTiers
        Case 1
            .RowHidden(0) = False
            .RowHidden(1) = True
            .RowHidden(2) = True
        Case 2
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = True
        Case 3
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = False
        End Select
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
    End With
    Exit Sub
errHand:
    MsgBox Err.Description
End Sub

Private Function ReadStruDef() As Boolean
    Dim arrCol
    Dim intCol As Integer, intCount As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '读取病历文件格式定义
    gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表格样式'" & _
        " Order By d.对象序号"
    Set rsTemp = OpenSQLRecord(gstrSQL, "读取病历文件格式定义", lngFormat)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !要素名称
            Case "表头层数": mintTabTiers = Val("" & !内容文本)
            Case "总列数":  VsfData.Cols = Val("" & !内容文本)
            Case "最小行高": VsfData.RowHeightMin = Val("" & !内容文本)
            Case "文本字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set VsfData.Font = objFont
                Set lblUpTable.Font = VsfData.Font
                Set Font = lblUpTable.Font
                
            Case "文本颜色": VsfData.ForeColor = Val("" & !内容文本)
            Case "表格颜色": VsfData.GridColor = Val("" & !内容文本): VsfData.GridColorFixed = VsfData.GridColor
            
            Case "标题文本": lblTitle.Caption = "" & !内容文本
            Case "标题字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set lblTitle.Font = objFont
                lblTitle.AutoSize = False
            
            Case "开始时间": mintTagFormHour = Val("" & !内容文本)
            Case "终止时间": mintTagToHour = Val("" & !内容文本)
            Case "条件字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set mobjTagFont = objFont
            Case "条件颜色": mlngTagColor = Val("" & !内容文本)
            Case "有效数据行": mlngActiveRows = Val(!内容文本)
            End Select
            .MoveNext
        Loop
    End With
    
    gstrSQL = "Select 格式, 页眉, 页脚,报表 From 病历页面格式 Where 种类 = 3 And 编号 In (Select 页面 From 病历文件列表 Where Id = [1])"
    Set rsTemp = OpenSQLRecord(gstrSQL, "读取病历页面格式", lngFormat)
    If Not rsTemp.EOF Then
        mstrPaperSet = "" & rsTemp!格式: mstrPageHead = "" & rsTemp!页眉: mstrPageFoot = "" & rsTemp!页脚
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称, Nvl(d.是否换行, 0) As 是否换行" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表上标签'" & _
        " Order By d.对象序号"
    Set rsTemp = OpenSQLRecord(gstrSQL, "读取表上标签定义", lngFormat)
    With rsTemp
        mstrSubhead = ""
        Do While Not .EOF
            mstrSubhead = mstrSubhead & "|" & IIf(!是否换行 = 0, "", vbCrLf) & !内容文本 & "{" & !要素名称 & "}"
            .MoveNext
        Loop
        If mstrSubhead <> "" Then mstrSubhead = Mid(mstrSubhead, 2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.对象序号, d.内容行次, d.内容文本" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表头单元'" & _
        " Order By d.对象序号"
    Set rsTemp = OpenSQLRecord(gstrSQL, "读取表头单元定义", lngFormat)
    With rsTemp
        mstrTabHead = ""
        Do While Not .EOF
            mstrTabHead = mstrTabHead & "|" & !内容行次 - 1 & "," & !对象序号 & "," & !内容文本
            .MoveNext
        Loop
        If mstrTabHead <> "" Then mstrTabHead = Mid(mstrTabHead, 2)
    End With
    
    '查询语句组织
    '------------------------------------------------------------------------------------------------------------------
    Dim strSql外 As String, str格式 As String
    Dim bln日期 As Boolean, bln时间 As Boolean, bln护士 As Boolean
    Dim bln签名人 As Boolean, bln签名时间 As Boolean, bln签名日期 As Boolean
    Dim lngColumn As Long
    
    gstrSQL = "Select d.对象序号, d.对象属性, d.内容行次, d.内容文本, d.要素名称, d.要素单位,d.要素表示 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表列集合'" & _
        " Order By d.对象序号, d.内容行次"
    Set rsTemp = OpenSQLRecord(gstrSQL, "读取表列集合定义", lngFormat)
    With rsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = ""
        mstrSQL内 = "": mstrSQL中 = "": strSql外 = "": mstrSQL列 = "": mstrSQL条件 = ""
        bln日期 = False: bln时间 = False: bln护士 = False
        bln签名人 = False: bln签名时间 = False: bln签名日期 = False
        Do While Not .EOF
            
            If lngColumn <> !对象序号 Then
                mstrColumns = mstrColumns & IIf(mstrColumns = "", "", ";1;" & str格式) & "|" & !对象序号 & ";" & !要素名称
                mstrColWidth = mstrColWidth & "," & !对象属性
                str格式 = ""
                If !要素名称 <> "" Then
                    str格式 = "{" & NVL(!内容文本) & "[" & !要素名称 & "]" & NVL(!要素单位) & "}"
                    mstrSQL列 = mstrSQL列 & "," & Mid(strSql外, 3) & " As C" & Format(lngColumn, "00")
                Else
                    If strSql外 <> "" Then
                        mstrSQL列 = mstrSQL列 & "," & Mid(strSql外, 3) & " As C" & Format(lngColumn, "00")
                    Else
                        mstrSQL列 = mstrSQL列 & ",'' As C" & Format(lngColumn, "00")
                    End If
                End If
                strSql外 = ""
                lngColumn = !对象序号
            Else
                mstrColumns = mstrColumns & "," & !要素名称
                str格式 = str格式 & "{" & NVL(!内容文本) & "[" & !要素名称 & "]" & NVL(!要素单位) & "}"
            End If
            
            Select Case !要素名称
            Case "日期"
                bln日期 = True
                mstrSQL中 = mstrSQL中 & ",日期"
                mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'yyyy-mm-dd') As 日期"
                strSql外 = strSql外 & "||" & !要素名称
            Case "时间"
                bln时间 = True
                mstrSQL中 = mstrSQL中 & ",时间"
                mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'hh24:mi') As 时间"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "签名人"
                bln签名人 = True
                mstrSQL中 = mstrSQL中 & ",签名人"
                mstrSQL内 = mstrSQL内 & ",l.签名人 As 签名人"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "签名时间"
                bln签名时间 = True
                mstrSQL中 = mstrSQL中 & ",签名时间"
                mstrSQL内 = mstrSQL内 & ",Decode(a.项目名称,Null,Null,Substr(a.项目名称,12,5)) As 签名时间"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "签名日期"
                bln签名日期 = True
                mstrSQL中 = mstrSQL中 & ",签名日期"
                mstrSQL内 = mstrSQL内 & ",Decode(a.项目名称,Null,Null,Substr(a.项目名称, 1,11)) As 签名日期"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "护士"
                bln护士 = True
                mstrSQL中 = mstrSQL中 & ",护士"
                mstrSQL内 = mstrSQL内 & ",l.保存人 As 护士"
                strSql外 = strSql外 & "||" & !要素名称
            Case Else
                If !要素名称 <> "" Then
                    mstrSQL中 = mstrSQL中 & ",Max(""" & !要素名称 & """) As """ & !要素名称 & """"
                    mstrSQL条件 = mstrSQL条件 & " Or """ & !要素名称 & """ Is Not Null"
                    strSql外 = strSql外 & "||""" & !要素名称 & """"
                    
                    If Trim("" & !内容文本) = "" And Trim("" & !要素单位) = "" Then
                        mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,c.记录内容), '') As """ & !要素名称 & """"
                    Else
                        mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,Decode(c.记录内容,Null,Null,'" & !内容文本 & "'||c.记录内容||'" & !要素单位 & "')), '') As """ & !要素名称 & """"
                    End If
                End If
            End Select
            .MoveNext
        Loop
        
        mstrColWidth = Mid(mstrColWidth, 2)
        '加入最后一列的格式
        mstrColumns = mstrColumns & IIf(mstrColumns = "", "", ";1;" & str格式) '& "|" & !对象序号 & ";" & !要素名称
        mstrColumns = Mid(mstrColumns, 2)     '格式如:列号;项目名称1,项目名称2|列号...,实例;1;体温|2;脉搏|3...
        If Mid(strSql外, 3) <> "" Then
            mstrSQL列 = mstrSQL列 & "," & Mid(strSql外, 3) & " As C" & Format(lngColumn, "00")
        Else
            mstrSQL列 = mstrSQL列 & ",'' As C" & Format(lngColumn, "00")
        End If
        
        If mstrSQL条件 <> "" Then mstrSQL条件 = "(" & Mid(mstrSQL条件, 5) & ")"
        
        '如果没有出现日期，时间，护士，则内层需要补充，以保证中层分组的正常：
        If bln日期 = False Then mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'yyyy-mm-dd') As 日期"
        If bln时间 = False Then mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'hh24:mi') As 时间"
        If bln护士 = False Then mstrSQL内 = mstrSQL内 & ",l.保存人 As 护士"
        
        If bln签名人 = False Then mstrSQL内 = mstrSQL内 & ",l.签名人 As 签名人"
        If bln签名日期 = False Then mstrSQL内 = mstrSQL内 & ",Decode(a.项目名称,Null,Null,Substr(a.项目名称,1,11)) As 签名日期"
        If bln签名时间 = False Then mstrSQL内 = mstrSQL内 & ",Decode(a.项目名称,Null,Null,Substr(a.项目名称,12,5)) As 签名时间"
        
        If Mid(mstrSQL中, 2) = "" Then
            MsgBox "对不起，您没有定义当前护理单的显示列信息，请在病历文件管理中定义！"
            Exit Function
        End If
        
        '程序内部控制增加固定列
        mstrSQL中 = mstrSQL中 & ",MAX(证书ID) AS 证书ID,MAX(签名级别) AS 签名级别,MAX(记录ID) AS 记录ID"
        mstrSQL内 = mstrSQL内 & ",A.项目ID AS 证书ID,NVL(A.记录内容,'护士') AS 签名级别,C.记录ID"
        mstrSQL列 = mstrSQL列 & ",证书ID,签名级别,记录ID"
        
        '分析哪些列的数据需要进行打印解析处理
        Dim arrData
        Dim strtodo As String
        Dim intto As Integer, intdo As Integer
        mstrPreHead = ""
        arrCol = Split(mstrColumns, "|")
        intCount = UBound(arrCol)
        For intCol = 0 To intCount
            If UBound(Split(Split(arrCol(intCol), ";")(3), "}{")) > 0 Then
                '只要有一个不是数字型则作为文本型处理
                
                strtodo = Split(arrCol(intCol), ";")(3)
                strtodo = Replace(strtodo, "]}{[", "||")
                strtodo = Replace(Replace(strtodo, "{[", ""), "]}", "")
                arrData = Split(strtodo, "||")
                intdo = UBound(arrData)
                For intto = 0 To intdo
                    mrsItems.Filter = "项目名称='" & arrData(intto) & "'"
                    If mrsItems.RecordCount <> 0 Then
                        '如果用户设置项目时都是设置成文本型,那么长度在20及以上的项目才检查,用户设置将数字型的设置成数字型才正确
                        If mrsItems!项目类型 = 1 And mrsItems!项目表示 = 0 And mrsItems!项目长度 >= 20 Then
                            mstrPreHead = mstrPreHead & "," & Val(Split(arrCol(intCol), ";")(0)) + 1    '有两列固定的列，而列序号从0开始，因此+1
                            Exit For
                        End If
                    End If
                Next
            Else
                '检查是否为文本项
                mrsItems.Filter = "项目名称='" & Replace(Replace(Split(arrCol(intCol), ";")(3), "{[", ""), "]}", "") & "'"
                If mrsItems.RecordCount <> 0 Then
                    '如果用户设置项目时都是设置成文本型,那么长度在20及以上的项目才检查,用户设置将数字型的设置成数字型才正确
                    If mrsItems!项目类型 = 1 And mrsItems!项目表示 = 0 And mrsItems!项目长度 >= 20 Then
                        mstrPreHead = mstrPreHead & "," & Val(Split(arrCol(intCol), ";")(0)) + 1    '有两列固定的列，而列序号从0开始，因此+1
                    End If
                End If
            End If
        Next
        
        mrsItems.Filter = 0
        If mstrPreHead = "" Then Exit Function
        mstrPreHead = Mid(mstrPreHead, 2)
        Call SQLCombination
    End With
    
    ReadStruDef = True
    Exit Function
errHand:
    MsgBox Err.Description
End Function

Private Sub SQLCombination()
    mstrSQL = "Select 备用,发生时间," & Mid(mstrSQL列, 10) & vbCrLf & _
                " From (Select 记录组号,时间 as 备用,发生时间," & Mid(mstrSQL中, 2) & vbCrLf & _
                "        From (Select c.记录组号,l.发生时间," & Mid(mstrSQL内, 2) & vbCrLf & _
                "               From 病人护理数据 l, 病人护理明细 c,病人护理明细 a,病人护理文件 f " & vbCrLf & _
                "               Where l.Id = c.记录id And l.文件ID=f.ID " & _
                "               And a.记录id(+)=l.ID And a.记录类型(+)=5 And Nvl(a.终止版本,0)=0 And c.终止版本 Is Null And c.记录类型<>5  " & _
                "               And f.id=[1] And f.病人id = [2] And f.主页id = [3] And Nvl(f.婴儿,0)=[4] )" & vbCrLf & _
                IIf(mstrSQL条件 <> "", "Where " & mstrSQL条件, "") & _
                "       Group By 日期, 时间, 发生时间,记录组号,护士,签名人,签名日期,签名时间" & _
                                "       Order By 日期, 时间, 发生时间,记录组号,护士,签名人,签名日期,签名时间)"
End Sub

'######################################################################################################################
'**********************************************************************************************************************
'以#分隔的区域内的代码都与分行相关,没事别动
Private Function GetData(ByVal strInput As String) As Variant
    Dim arrData
    Dim strData As String
    Dim strLine(256) As Byte
    Dim lngRow As Long, lngRows As Long
    
    GetData = ""
    lngRows = SendMessage(txtLength.hWnd, EM_GETLINECOUNT, 0&, 0&)
    For lngRow = 1 To lngRows
        Call ClearArray(strLine)
        Call SendMessage(txtLength.hWnd, EM_GETLINE, lngRow - 1, strLine(0))
        strData = StrConv(strLine, vbUnicode)
        strData = TruncZero(strData)
        GetData = GetData & IIf(GetData = "", "", "|ZYB.ZLSOFT|") & strData
    Next
    GetData = Split(GetData, "|ZYB.ZLSOFT|")
End Function

Private Sub ClearArray(strLine() As Byte)
    Dim intdo As Integer, intMax As Integer
    intMax = UBound(strLine)
    For intdo = 0 To intMax
        strLine(intdo) = 0
    Next
    strLine(1) = 1
End Sub

Private Function TrimStr(ByVal str As String) As String
'功能：去掉字符串中\0以后的字符，并且去掉两端的空格

    If InStr(str, Chr(0)) > 0 Then
        TrimStr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        TrimStr = Trim(str)
    End If
End Function

Private Function TruncZero(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function
