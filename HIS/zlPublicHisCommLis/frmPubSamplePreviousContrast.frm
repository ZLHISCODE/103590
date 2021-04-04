VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form frmPubSamplePreviousContrast 
   BorderStyle     =   0  'None
   Caption         =   "历次比对"
   ClientHeight    =   6165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PICContrast 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5565
      Left            =   0
      ScaleHeight     =   5565
      ScaleWidth      =   7005
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7005
      Begin VB.PictureBox PicContrast_Bottom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCDBD8&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1635
         Left            =   0
         ScaleHeight     =   1635
         ScaleWidth      =   5280
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1110
         Width           =   5280
         Begin VB.OptionButton optContrast 
            Appearance      =   0  'Flat
            BackColor       =   &H00FCDBD8&
            Caption         =   "结果值(&2)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   2745
            TabIndex        =   8
            Top             =   60
            Width           =   1500
         End
         Begin VB.OptionButton optContrast 
            Appearance      =   0  'Flat
            BackColor       =   &H00FCDBD8&
            Caption         =   "变异率(&1)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   1170
            TabIndex        =   7
            Top             =   38
            Value           =   -1  'True
            Width           =   1395
         End
         Begin C1Chart2D8.Chart2D chtContrast 
            Height          =   975
            Left            =   60
            TabIndex        =   9
            Top             =   480
            Width           =   1005
            _Version        =   524288
            _Revision       =   7
            _ExtentX        =   1773
            _ExtentY        =   1720
            _StockProps     =   0
            ControlProperties=   "frmPubSamplePreviousContrast.frx":0000
         End
         Begin VB.Label lblCht 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "图形内容"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   30
            TabIndex        =   10
            Top             =   60
            Width           =   960
         End
      End
      Begin VB.PictureBox PicContrast_Top 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   30
         ScaleHeight     =   975
         ScaleWidth      =   5970
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   5970
         Begin VB.TextBox txtMaxDay 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1650
            TabIndex        =   2
            Text            =   "30"
            Top             =   60
            Width           =   705
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFContrast 
            Height          =   1335
            Left            =   60
            TabIndex        =   3
            Top             =   480
            Width           =   2265
            _cx             =   3995
            _cy             =   2355
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
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
            BackColorSel    =   16777215
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483635
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   2
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   350
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
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
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   0
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
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
         Begin VB.Label lblContrast 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "刷新"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00F56C58&
            Height          =   240
            Left            =   2790
            MouseIcon       =   "frmPubSamplePreviousContrast.frx":0595
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   90
            Width           =   480
         End
         Begin VB.Label lblMaxDay 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "最大跟踪天数:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   90
            TabIndex        =   4
            Top             =   90
            Width           =   1560
         End
      End
   End
End
Attribute VB_Name = "frmPubSamplePreviousContrast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'创    建:蔡青松
'创建时间:2019-07-22
'模块功能:标本结果历次比对
'---------------------------------------------------------------------------------------

Option Explicit

Private mlngSampleID As Long        '标本ID
Private mdteS As Date               '报告时间
Private mintVersion As Integer      '版本


'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-07-22
'功    能:  加载数据
'入    参:
'           cnOracle        连接对象
'           标本ID
'           dteS            标本核收日期
'           intVersion      版本25=新版，10=老版
'出    参:
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Public Sub InitData(ByVal lngSampleID As Long, ByVal dteS As Date, Optional ByVal intVersion As Integer = 25)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngPaitID As Long
    
    mlngSampleID = lngSampleID
    mdteS = dteS
    mintVersion = intVersion
    
    If Val(txtMaxDay) > 365 Then
        If MsgBox("录入的最大跟踪天数超过一年，是否继续查看历次数据，如果继续可能会导致加载慢的情况！", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Exit Sub
        End If
    End If
    
    Me.VSFContrast.Rows = 1
    Me.VSFContrast.Rows = 2
    
    If intVersion = 25 Then
        strSQL = "select 病人ID from 检验报告记录 where ID=[1]"
        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "病人ID", lngSampleID)
    Else
        strSQL = "select 病人ID from 检验标本记录 where ID=[1]"
        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "病人ID", lngSampleID)
    End If
    
    If Not rsTmp.EOF Then lngPaitID = Val(rsTmp("病人ID") & "")
    If lngPaitID = 0 Then Exit Sub
    Call LoadContrastDBWriteVSF(Me.VSFContrast, lngSampleID, lngPaitID, dteS, Val(txtMaxDay.Text), intVersion)
End Sub

Public Function LoadContrastDBWriteVSF(vsfList As VSFlexGrid, lngSampleID As Long, lngPatientID As Long, SampleReportDate As Date, _
                                       intMaxDay As Integer, ByVal intVersion As Integer) As Boolean
      '功能                   从数据库中读出比对数据写入VSF中
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim intCol As Integer
          Dim dblTmp As Double
          Dim strData As String
          Dim blnTre As Boolean       '是否是耐受试验标本


1         On Error GoTo LoadContrastDBWriteVSF_Error

2         blnTre = gobjLiscomlib.IsTre(lngSampleID)

3         If intVersion = 25 Then
4             If blnTre Then
                  '耐受试验
5                 strSQL = "Select b.id, b.中文名, b.英文名, b.单位, a.id 次数, c.报告时间, a.检验结果, e.耐受时间, b.变异报警率, b.结果类型, a.结果标志" & vbCrLf & _
                         "   From 检验报告明细 A, 检验指标 B, 检验报告记录 C, 耐受试验标本 D, 检验耐受时间方案 E" & vbCrLf & _
                         "   Where a.项目ID = b.id And a.标本id = c.id And a.id = d.报告明细id And d.耐受方案id = e.id And a.标本ID =[1] order by a.id desc"
6                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读入比对数据", lngSampleID)
7             Else

8                 strSQL = "Select " & vbNewLine & _
                         " B.Id, B.中文名, B.英文名, B.单位, A.次数, A.报告时间, A.检验结果, A.核收时间, B.变异报警率, B.结果类型, A.结果标志" & vbNewLine & _
                           "From (Select B.项目id 检验项目id, B.次数, B.报告时间, B.检验结果, B.核收时间, B.结果标志" & vbNewLine & _
                         "       From (Select A.Id 次数, A.病人id, A.姓名, A.性别, A.标本类型, A.报告时间, B.项目id, A.核收时间, B.结果标志, a.主页id" & vbNewLine & _
                         "              From 检验报告记录 A, 检验报告明细 B" & vbNewLine & _
                         "              Where A.Id = B.标本id And A.Id = [1] ) A," & vbNewLine & _
                         "            (Select A.Id 次数, A.病人id, A.姓名, A.性别, A.标本类型, A.报告时间, B.项目id, B.检验结果, A.核收时间, B.结果标志, a.主页id" & vbNewLine & _
                         "              From 检验报告记录 A, 检验报告明细 B" & vbNewLine & _
                         "              Where A.Id = B.标本id And A.病人id = [2] And" & vbNewLine & _
                         "                    报告时间 Between [3] And" & vbNewLine & _
                         "                    [4] And A.Id <= [1] ) B," & vbNewLine & _
                         "            (Select A.Id 次数" & vbNewLine & _
                         "              From 检验报告记录 A, 检验报告明细 B" & vbNewLine & _
                         "              Where A.Id = B.标本id And A.病人id = [2] And" & vbNewLine & _
                         "                    报告时间+0 Between [3] And" & vbNewLine & _
                         "                    [4] And A.Id <= [1] " & vbNewLine & _
                         "              Group By A.Id" & vbNewLine & _
                         "              Having Count(A.Id) > 0) C" & vbNewLine & _
                         "       Where A.病人id = B.病人id And A.项目id + 0 = B.项目id And Nvl(A.标本类型, 0) = Nvl(B.标本类型, 0) And A.姓名 = B.姓名 And A.性别 = B.性别 And a.主页ID = b.主页ID And" & vbNewLine & _
                         "             B.次数 = C.次数) A, 检验指标 B" & vbNewLine & _
                           "Where A.检验项目id = B.Id" & vbNewLine & _
                           "Order By  A.次数 Desc ,LPad(B.排列序号, 10, '0'), B.Id"

9                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读入比对数据", lngSampleID, lngPatientID, CDate(Format(SampleReportDate - intMaxDay, "yyyy-MM-dd") & " 00:00:00"), _
                                            CDate(Format(SampleReportDate, "yyyy-MM-dd") & " 23:59:59"))
10            End If
11        Else
12            strSQL = "    Select " & vbNewLine & _
                     "       i.Id, i.名称 As 中文名, v.缩写 As 英文名, i.计算单位 As 单位, a.次数, a.报告时间, a.检验结果, v.变异报警率, v.结果类型" & vbNewLine & _
                     "       From (Select b.检验项目id, b.次数, b.报告时间, b.检验结果" & vbNewLine & _
                     "              From (Select a.Id 次数, a.病人id, a.标本类型, a.审核时间 报告时间, b.检验项目id, b.检验结果" & vbNewLine & _
                     "                     From 检验标本记录 A, 检验普通结果 B" & vbNewLine & _
                     "                     Where a.Id = b.检验标本id And a.Id = [1] And 病人id = [2] And b.检验结果 Is Not Null) A," & vbNewLine & _
                     "                   (Select a.Id 次数, a.病人id, a.标本类型, a.审核时间 报告时间, b.检验项目id, b.检验结果" & vbNewLine & _
                     "                     From 检验标本记录 A, 检验普通结果 B" & vbNewLine & _
                     "                     Where a.Id = b.检验标本id And a.Id < [1] And 病人id = [2]  And  审核时间 Between [3] And [4]  And b.检验结果 Is Not Null) B" & vbNewLine & _
                     "              Where a.病人id = b.病人id And a.检验项目id + 0 = b.检验项目id) A, 检验项目 V, 检验报告项目 R, 诊疗项目目录 I" & vbNewLine & _
                     "       Where A.检验项目id = v.诊治项目id And A.检验项目id = r.报告项目id And r.诊疗项目id = i.ID And i.组合项目 <> 1"
13            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "读入比对数据", lngSampleID, lngPatientID, CDate(Format(SampleReportDate - intMaxDay, "yyyy-MM-dd") & " 00:00:00"), _
                                     CDate(Format(SampleReportDate, "yyyy-MM-dd") & " 23:59:59"))
14        End If

15        gobjLiscomlib.vfgSetting 0, vsfList
16        With vsfList

17            .Rows = 1
18            .Cols = 2
19            .FixedRows = 1
              '        .FixedCols = 1
20            .TextMatrix(0, 0) = "检验项目": .ColWidth(0) = 2500
21            .TextMatrix(0, 1) = "项目id": .ColWidth(1) = 0: .ColHidden(1) = True
              Dim strTimes As String
              Dim i As Long
              Dim strBT As String
              Dim strBH As String
              Dim J As Long
22            Do Until rsTmp.EOF
                  '按照次数进行区分历次结果，开始就写入本次结果
                  '最开始写入的时候 ，没有次数，通过本次结果写入上次结果
23                If strTimes = "" Or strTimes = rsTmp("次数") & "" Then
24                    If strBT <> "**" Then
25                        strBH = "**"
26                        .Rows = .Rows + 1
27                        intCol = 1
28                        If .Cols - 1 < intCol Then
29                            .Cols = .Cols + 1
30                            .ColWidth(intCol) = 1500
31                        End If

32                        If intCol = 1 Then
                              '写入项目
33                            .TextMatrix(.Rows - 1, 0) = rsTmp("中文名") & "(" & rsTmp("英文名") & ")"
34                            .TextMatrix(.Rows - 1, 1) = rsTmp("id")
35                        End If
36                        intCol = intCol + 1
37                        If .Cols - 1 < intCol Then
38                            .Cols = .Cols + 1
39                            .ColWidth(intCol) = 1200: .ColAlignment(intCol) = flexAlignLeftCenter
40                            If blnTre Then
41                                .TextMatrix(0, intCol) = rsTmp("耐受时间") & ""
42                            Else
43                                .TextMatrix(0, intCol) = "本次"
44                            End If
45                        End If
                          '写入内容
46                        .TextMatrix(.Rows - 1, intCol) = rsTmp("检验结果") & "(已做)"
47                        .Cell(flexcpBackColor, .Rows - 1, intCol) = GetValColour(Val(rsTmp("结果标志") & ""))
48                    End If
49                End If
                  '上一次的结果写入下一列
50                If strTimes <> "" And strBH <> "**" Then
51                    If strTimes <> rsTmp("次数") & "" Then
52                        If strBT = "" Then strBT = "**"
53                        For i = 1 To .Rows - 1
54                            If .TextMatrix(i, 1) = rsTmp("id") Then
55                                .Cols = .Cols + 1
56                                If blnTre Then
57                                    strData = rsTmp("耐受时间") & ""
58                                    .ColWidth(.Cols - 1) = 2000: .ColAlignment(.Cols - 1) = flexAlignLeftCenter: .TextMatrix(0, .Cols - 1) = strData
59                                Else
60                                    strData = Format(rsTmp("核收时间") & "", "yyyy-MM-dd HH:mm")

61                                    .ColWidth(.Cols - 1) = 2000: .ColAlignment(.Cols - 1) = flexAlignLeftCenter: .TextMatrix(0, .Cols - 1) = "上" & .Cols - 3 & "次" & "(" & strData & ")"
62                                End If
63                                dblTmp = Val(CalcVolatility(.TextMatrix(i, 1), .TextMatrix(i, .Cols - 1)))
64                                If dblTmp <> 0 And Val(rsTmp("变异报警率") & "") <> 0 Then
65                                    If dblTmp > Val(rsTmp("变异报警率") & "") Then
66                                        .Cell(flexcpBackColor, i, .Cols - 1) = RGB(248, 194, 169)
67                                    End If
68                                End If
                                  '写入内容
69                                .TextMatrix(i, .Cols - 1) = rsTmp("检验结果") & "(已做)"
70                                .Cell(flexcpBackColor, i, .Cols - 1) = GetValColour(Val(rsTmp("结果标志") & ""))
71                            End If
72                        Next
73                    Else
74                        For i = 1 To .Rows - 1
75                            If .TextMatrix(i, 1) = rsTmp("id") Then
                                  '写入内容
76                                .TextMatrix(i, .Cols - 1) = rsTmp("检验结果") & "(已做)"
77                                .Cell(flexcpBackColor, i, .Cols - 1) = GetValColour(Val(rsTmp("结果标志") & ""))
78                            End If
79                        Next
80                    End If
81                End If
82                strTimes = rsTmp("次数") & ""
83                strBH = ""
84                rsTmp.MoveNext
85            Loop
86            For i = 2 To .Cols - 1
87                For J = 1 To .Rows - 1
88                    If .TextMatrix(J, i) <> "" Then
89                        .TextMatrix(J, i) = Replace(.TextMatrix(J, i), "(已做)", "")
90                        If .TextMatrix(J, i) = "" Then
91                            .TextMatrix(J, i) = "无结果"
92                        End If
93                    Else
94                        .TextMatrix(J, i) = "未做"
95                    End If
96                Next
97            Next

98            If .Rows > 1 Then
99                .Row = 1
100               Call VSFContrast_SelChange
101           End If

102       End With



103       Exit Function
LoadContrastDBWriteVSF_Error:
104       Call WriteErrLog("zl9LisInsideComm", "frmPubSamplePreviousContrast", "执行(LoadContrastDBWriteVSF)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
105       Err.Clear

End Function

Private Sub Form_Resize()
    On Error Resume Next
    With PICContrast
        .Left = 0
        .Top = 0
        .Width = Me.Width
        .Height = Me.Height
    End With
End Sub

Private Sub lblContrast_Click()
    Call InitData(mlngSampleID, mdteS, mintVersion)
End Sub

Private Sub PicContrast_Bottom_Resize()
    On Error Resume Next
    With Me.chtContrast
        .Top = lblCht.Top + lblCht.Height + 75
        .Left = 0
        .Width = Me.PicContrast_Bottom.ScaleWidth
        .Height = Me.PicContrast_Bottom.ScaleHeight
    End With
End Sub

Private Sub PICContrast_Resize()
    On Error Resume Next
    With Me.PicContrast_Top
        .Top = 0
        .Left = 0
        .Width = Me.PICContrast.ScaleWidth
        .Height = Me.PICContrast.ScaleHeight / 2
    End With
    With Me.PicContrast_Bottom
        .Top = PicContrast_Top.Top + PicContrast_Top.Height + 25
        .Left = 0
        .Width = Me.PicContrast_Top.Width
        .Height = Me.PICContrast.ScaleHeight - .Top
    End With
End Sub

Private Sub PicContrast_Top_Resize()
    On Error Resume Next
    With Me.VSFContrast
        .Top = Me.lblMaxDay.Top + lblMaxDay.Height + 50
        .Left = 0
        .Width = PicContrast_Top.ScaleWidth
        .Height = PicContrast_Top.ScaleHeight - .Top
    End With
End Sub

Private Sub txtMaxDay_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Then Exit Sub
End Sub

Private Sub VSFContrast_SelChange()
    Dim strErr As String
    Dim intType As Integer
    With Me.VSFContrast
        If .Cols < 1 Then Exit Sub
        If .Row > 0 Then
            If Me.optContrast(0).value = True Then
                intType = 1
            Else
                intType = 2
            End If
            If LoadVSFContrastToCht(Me.VSFContrast, Me.chtContrast, .Row, intType, strErr) = False Then
                MsgBox strErr, vbInformation, gSysInfo.AppName
            End If
        End If
    End With
End Sub

Public Function LoadVSFContrastToCht(vsfList As VSFlexGrid, chtObj As Chart2D, intRow As Integer, intType As Integer, strErr As String) As Boolean
          '功能           从VSF读出数据写入Cht控件
          Dim intCol As Integer
          Dim dblMax As Double
          Dim i As Integer

1         On Error GoTo LoadVSFContrastToCht_Error

2         chtObj.ChartGroups(1).Data.NumSeries = 0
3         With chtObj.ChartGroups(1)
4             .ChartType = oc2dTypePlot  '折线
5             .Styles(oc2dTypePlot).Symbol.Shape = oc2dShapeBox
6             With .Data
7                 .Layout = oc2dDataArray
8                 .NumSeries = 1
9                 .NumPoints(1) = vsfList.Cols - 1
10            End With
11        End With

12        With chtObj.ChartArea
13            .Axes("X").MajorGrid.Spacing.IsDefault = True
14            .Axes("Y").MajorGrid.Spacing.IsDefault = True
15            .Axes("X").AnnotationMethod = oc2dAnnotateValueLabels   '横坐标显示值提示

16        End With
17        With chtObj.ChartGroups(1).Data
18            For intCol = 2 To vsfList.Cols - 1
19                If IsNumeric(vsfList.TextMatrix(vsfList.Row, intCol)) = True Then
20                    i = i + 1
21                    Select Case intType
                          Case 1
22                            If intCol = 2 Then
23                                If vsfList.TextMatrix(vsfList.Row, 2) <> "" Then
24                                    .Y(1, i) = 0
25                                End If
26                            Else
27                                If IsNumeric(vsfList.TextMatrix(vsfList.Row, 2)) = True And IsNumeric(vsfList.TextMatrix(vsfList.Row, intCol)) = True Then
28                                    If CalcVolatility(vsfList.TextMatrix(vsfList.Row, 2), vsfList.TextMatrix(vsfList.Row, intCol)) <> "" Then
29                                        .Y(1, i) = Val(CalcVolatility(vsfList.TextMatrix(vsfList.Row, 2), vsfList.TextMatrix(vsfList.Row, intCol)))
30                                    Else
31                                        .Y(1, i) = 1E+308
32                                    End If
33                                End If
34                            End If
35                        Case 2
36                            If IsNumeric(vsfList.TextMatrix(vsfList.Row, intCol)) = True Then
37                                .Y(1, i) = IIf(vsfList.TextMatrix(vsfList.Row, intCol) = "", 1E+308, vsfList.TextMatrix(vsfList.Row, intCol))
38                            End If
39                    End Select
40                    If Abs(.Y(1, i)) > Abs(dblMax) And .Y(1, i) <> 1E+308 Then
41                        dblMax = .Y(1, i)
42                    End If
43                End If
44            Next
45        End With

46        With chtObj.ChartArea
47            Select Case intType
                  Case 1              '变异率
48                    .Axes("Y").DataMax = Abs(dblMax)
49                    .Axes("Y").DataMin = Abs(dblMax) * -1
50                    .Axes("Y").Origin = 0
51                Case 2              '结果值
52                    .Axes("Y").DataMax = Abs(dblMax) + Abs(dblMax) / 100 * 10
53                    .Axes("Y").DataMin = 0
54                    .Axes("Y").Origin = 0
55            End Select
56        End With
57        LoadVSFContrastToCht = True



58        Exit Function
LoadVSFContrastToCht_Error:
59        Call WriteErrLog("zl9LisInsideComm", "frmPubSamplePreviousContrast", "执行(LoadVSFContrastToCht)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
60        Err.Clear

End Function



Public Function CalcVolatility(strCalcA As String, strCalcB As String) As String
    '计算变异率

    On Error Resume Next

    If strCalcA = "" Or strCalcB = "" Then
        CalcVolatility = ""
        Exit Function
    End If
    If Val(strCalcA) = 0 Or Val(strCalcB) = 0 Then
        CalcVolatility = ""
    End If

    '计算
    CalcVolatility = (Val(strCalcB) - Val(strCalcA)) / Val(strCalcA) * 100
End Function
