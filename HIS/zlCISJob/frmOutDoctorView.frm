VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOutDoctorView 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "门诊诊疗一栏"
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   15855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PicCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   -360
      ScaleHeight     =   7695
      ScaleWidth      =   15630
      TabIndex        =   1
      Top             =   0
      Width           =   15630
      Begin VB.PictureBox picBottom 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   6570
         Left            =   480
         ScaleHeight     =   6570
         ScaleWidth      =   15630
         TabIndex        =   4
         Top             =   1080
         Width           =   15630
         Begin VSFlex8Ctl.VSFlexGrid vsView 
            Height          =   6570
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   15630
            _cx             =   27570
            _cy             =   11589
            Appearance      =   0
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
            MouseIcon       =   "frmOutDoctorView.frx":0000
            BackColor       =   -2147483643
            ForeColor       =   0
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483641
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   14737632
            GridColorFixed  =   10526880
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   4
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmOutDoctorView.frx":0162
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   -1  'True
            MergeCells      =   1
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   -1  'True
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
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   2
            FrozenCols      =   1
            AllowUserFreezing=   0
            BackColorFrozen =   16777215
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin MSComctlLib.ImageList imgFlag 
            Left            =   13920
            Top             =   -120
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   8
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   12
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":0260
                  Key             =   "报告已阅"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":077A
                  Key             =   "报告"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":0C94
                  Key             =   "上次可用"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":219C
                  Key             =   "上次高亮"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":36A4
                  Key             =   "上次不可用"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":4200
                  Key             =   "下次高亮"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":5708
                  Key             =   "下次可用"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":6C10
                  Key             =   "下次不可用"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":776C
                  Key             =   "只显示本科"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":A3E4
                  Key             =   "只显示本科高亮"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":D05C
                  Key             =   "显示所有就诊"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":FCD4
                  Key             =   "显示所有就诊高亮"
               EndProperty
            EndProperty
         End
      End
      Begin VB.PictureBox picTop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1005
         Left            =   360
         ScaleHeight     =   1634.146
         ScaleMode       =   0  'User
         ScaleWidth      =   15630
         TabIndex        =   2
         Top             =   0
         Width           =   15630
         Begin VB.PictureBox pitBtn 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   2
            Left            =   6360
            Picture         =   "frmOutDoctorView.frx":1294C
            ScaleHeight     =   480
            ScaleWidth      =   2400
            TabIndex        =   8
            Top             =   1560
            Width           =   2400
         End
         Begin VB.PictureBox pitBtn 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   1
            Left            =   3480
            Picture         =   "frmOutDoctorView.frx":1A690
            ScaleHeight     =   480
            ScaleWidth      =   2400
            TabIndex        =   7
            Top             =   1440
            Width           =   2400
         End
         Begin VB.PictureBox pitBtn 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   0
            Left            =   360
            Picture         =   "frmOutDoctorView.frx":223D4
            ScaleHeight     =   480
            ScaleWidth      =   2400
            TabIndex        =   6
            Top             =   1440
            Width           =   2400
         End
         Begin VB.Image imgBtn 
            Height          =   375
            Index           =   1
            Left            =   13800
            Picture         =   "frmOutDoctorView.frx":2A118
            Stretch         =   -1  'True
            Top             =   480
            Width           =   1050
         End
         Begin VB.Image imgBtn 
            Height          =   375
            Index           =   2
            Left            =   120
            Picture         =   "frmOutDoctorView.frx":2B610
            Stretch         =   -1  'True
            Top             =   480
            Width           =   2250
         End
         Begin VB.Image imgBtn 
            Height          =   375
            Index           =   0
            Left            =   12480
            Picture         =   "frmOutDoctorView.frx":2E278
            Stretch         =   -1  'True
            Top             =   480
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "诊疗一览"
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
            Left            =   6360
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.Label lblW 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "frmOutDoctorView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'入口参数
Private mlng病人ID As Long
Private mlng科室ID As Long

'程序变量
Private mrs就诊时间 As ADODB.Recordset        '保存病人历次就诊时间:（字段:序号、就诊ID、挂号单、执行时间）
Private mrs诊断 As ADODB.Recordset            '诊断信息
Private mrsDrug As ADODB.Recordset
Private mrs检查检验 As ADODB.Recordset
Private mrs病历文书 As ADODB.Recordset
Private mrsEMR As ADODB.Recordset

Private mrs其他医嘱 As ADODB.Recordset

Private mlngPrev As Long   '前一个   缺省 4
Private mlngNext As Long   '后一个   缺省 1
Private mbytFontSize As Long   '字体大小

Private mcolCate As Collection

Private mstr就诊ID As String
Private mstr挂号单 As String
Private mstrFontUnderLine As String   '标记下划线行  行号|列1|列2
Private mintIndex  As Integer        '标记当前高亮图标
Private mbytShow As Byte                '0-显示所有就诊记录;1-仅显示本科室就诊记录 (缺省显示所有就诊记录)

Private Const mstr分类 As String = "就诊时间|诊断|用药跟踪|病历文书|检查|检验|其他医嘱"
Private Const mlngSubCol  As Long = 5 '标注(口\滴\针)|图标(┏┃┗)|名称(药品名称)|疗程|间隔列

'颜色深度:浅红,淡红,鲜红,深红,暗红
Private Enum CONST_COLOR
    '红色
    COLOR_浅红 = &HC0C0FF          '浅红
    COLOR_淡红 = &H8080FF
    COLOR_鲜红 = &HFF&
    COLOR_深红 = &HC0&
    COLOR_暗红 = &H80&
    '橙色
    COLOR_浅橙 = &HC0E0FF
    COLOR_淡橙 = &H80C0FF
    COLOR_鲜橙 = &H80FF&
    COLOR_深橙 = &H40C0&
    COLOR_暗橙 = &H4080&
    '黄色
    COLOR_浅黄 = &HC0FFFF
    COLOR_淡黄 = &H80FFFF
    COLOR_鲜黄 = &HFFFF&
    COLOR_深黄 = &HC0C0&
    COLOR_暗黄 = &H8080&
    '绿色
    COLOR_浅绿 = &HC0FFC0
    COLOR_淡绿 = &H80FF80
    COLOR_鲜绿 = &HFF00&
    COLOR_深绿 = &HC000&
    COLOR_暗绿 = &H8000&
    '青色
    COLOR_浅青 = &HFFFFC0
    COLOR_淡青 = &HFFFF80
    COLOR_鲜青 = &HFFFF00
    COLOR_深青 = &HC0C000
    COLOR_暗青 = &H808000
    '蓝色
    COLOR_浅蓝 = &HFFC0C0
    COLOR_淡色 = &HFF8080
    COLOR_鲜蓝 = &HFF0000
    COLOR_深蓝 = &HC00000
    COLOR_暗蓝 = &H800000
    '紫色
    COLOR_浅紫 = &HFFC0FF
    COLOR_淡紫 = &HFF80FF
    COLOR_鲜紫 = &HFF00FF
    COLOR_深紫 = &HC000C0
    COLOR_暗紫 = &H800080
    '白色
    COLOR_白色 = &H80000005
    COLOR_FORMBK = &H8000000B
    COLOR_CENTERBK = &H808000
End Enum


Private Enum CONST_CATEGORY
    CATE_就诊时间 = 0
    CATE_诊断 = 1
    CATE_用药跟踪 = 2
    CATE_病历文书 = 3
    CATE_检查 = 4
    CATE_检验 = 5
    CATE_其他医嘱 = 6
End Enum

Private Enum CONST_IX_CMD
    CMD_PREV = 0
    CMD_NEXT = 1
    CMD_ORTHER = 2
End Enum

Private Enum CONST_SUBCOL
    SUBCOL_标注 = 0
    SUBCOL_图标 = 1
    SUBCOL_名称 = 2
    SUBCOL_疗程 = 3
    SUBCOL_间隔 = 4
End Enum

Public Function zlRefresh(frmParent As Object, ByVal lng病人ID As Long, ByVal lng科室ID As Long) As Boolean
'功能：
    If lng病人ID = 0 Then
        Exit Function
    End If
    mlng病人ID = lng病人ID
    mlng科室ID = lng科室ID
    mlngPrev = 3     '默认值
    mlngNext = 1     '默认值
    Call SubRefresh
End Function

Private Sub LoadView()
'功能:加载视图
    Dim strTmp As String
    Dim i As Long, k As Long, j As Long
    Dim lng相关ID As Long, lngCount As Long
    Dim lngColor As Long, lngContW As Long
    Dim lng挂号id As Long
    Dim lngDay As Long
    Dim strFilter As String, strContent As String
    Dim strDrug As String
    Dim strType As String
    Dim strDay As String
    Dim strMerge As String     '记录一并给药需要合并的行号
    Dim lngRow As Long, lngMergeRow As Long
    Dim lngDrugCol As Long
    Dim rsDrug As ADODB.Recordset
    Dim blnAddRow As Boolean
    Dim blnMerge As Boolean
    With vsView
        If mrs就诊时间 Is Nothing Then Exit Sub
        
        .Redraw = flexRDNone
        mrs就诊时间.Filter = "序号 >=" & mlngNext & " And 序号 <= " & mlngPrev
        mrs就诊时间.MoveLast
        For i = 1 To mrs就诊时间.RecordCount
            '就诊时间
            If NVL(mrs就诊时间!序号, 0) = 1 Then
                strTmp = "本次就诊"
            ElseIf NVL(mrs就诊时间!序号, 0) = 2 Then
                strTmp = "上次就诊"
            Else
                strTmp = ""
            End If
            strTmp = strTmp & " " & Format(mrs就诊时间!执行时间 & "", "YYYY-MM-DD hh:mm") & IIf(Val(mrs就诊时间!执行部门ID & "") <> mlng科室ID, "[" & mrs就诊时间!科室名称 & "]", "")
            .Cell(flexcpText, 0, (i - 1) * mlngSubCol + 1, 0, i * mlngSubCol) = strTmp
            .ColData((i - 1) * mlngSubCol + SUBCOL_名称 + 1) = CLng(mrs就诊时间!就诊id) ' 记录每列的就诊ID  '医嘱名称列
            .Cell(flexcpData, mcolCate("_" & CATE_就诊时间).lngBeginRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = CStr(mrs就诊时间!挂号单) '记录下挂号单
            
            '界面格式调整
            .Cell(flexcpAlignment, 0, 1, 0, .Cols - 1) = flexAlignCenterCenter '就诊时间
            .Cell(flexcpAlignment, 1, 1, 1, .Cols - 1) = flexAlignLeftCenter   '诊断
   
            '就诊诊断
            mrs诊断.Filter = "就诊ID =" & mrs就诊时间!就诊id
            strTmp = ""
            For j = 1 To mrs诊断.RecordCount
                strTmp = strTmp & "," & mrs诊断!诊断描述 & ""
                mrs诊断.MoveNext
            Next
            lngContW = .ColWidth(SUBCOL_标注) + .ColWidth(SUBCOL_名称) + .ColWidth(SUBCOL_疗程) - 300
            If strTmp <> "" Then
                strTmp = Mid(strTmp, 2)
                .Cell(flexcpData, mcolCate("_" & CATE_诊断).lngBeginRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = strTmp
                strTmp = GetSubString(strTmp, lngContW)
                .Cell(flexcpText, mcolCate("_" & CATE_诊断).lngBeginRow, (i - 1) * mlngSubCol + 1, mcolCate("_" & CATE_诊断).lngBeginRow, i * mlngSubCol) = strTmp
            Else
                .Cell(flexcpText, mcolCate("_" & CATE_诊断).lngBeginRow, (i - 1) * mlngSubCol + 1, mcolCate("_" & CATE_诊断).lngBeginRow, i * mlngSubCol) = IIf(i Mod 2 = 0, " ", "  ") '用于诊断行的合并处理
            End If
            
            '用药跟踪
            lngDrugCol = (i - 1) * mlngSubCol + SUBCOL_名称 + 1
            strFilter = "挂号单 ='" & mrs就诊时间!挂号单 & "' And 诊疗类别 <> 'E'"
            mrsDrug.Filter = strFilter
            Set rsDrug = zlDatabase.CopyNewRec(mrsDrug)
            lngRow = mcolCate("_" & CATE_用药跟踪).lngBeginRow
            lngCount = 0
            For k = 1 To rsDrug.RecordCount
                blnAddRow = True: blnMerge = False
                '诊疗类别=5,6
                '医嘱内容:药品类型(滴,针,口)vbTab药品名称 vbTAB 疗程(单位为天,大于7天默认为7天)
                If rsDrug!诊疗类别 & "" = "5" Or rsDrug!诊疗类别 & "" = "6" Then
                    .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1) = CLng(rsDrug!相关ID & "")      '用于标识同一组药品
                    '标记一并给药┏┃┗
                    mrsDrug.Filter = "相关ID=" & CLng(rsDrug!相关ID & "")
                    
                    If mrsDrug.RecordCount > 1 Then
                        If lngMergeRow = 0 Then lngMergeRow = 1
                        blnMerge = True:
                        If lngMergeRow = 1 Then
                            .TextMatrix(lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1) = "┏"
                            lngMergeRow = lngMergeRow + 1
                            lngCount = lngCount + 1
                        ElseIf lngMergeRow = mrsDrug.RecordCount Then
                            .TextMatrix(lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1) = "┗"
                            lngMergeRow = 0
                        ElseIf lngMergeRow <> 1 And lngMergeRow <> mrsDrug.RecordCount Then
                            .TextMatrix(lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1) = "┃"
                            lngMergeRow = lngMergeRow + 1
                        End If
                    Else
                        lngMergeRow = 0
                        lngCount = lngCount + 1
                    End If
                        
                    If lng相关ID <> CLng(rsDrug!相关ID & "") Then '避免一并给药重复取
                        lng相关ID = CLng(rsDrug!相关ID & "")

                        mrsDrug.Filter = "ID=" & lng相关ID     '给药途径
                        '0-其他治疗类别,1-输液类,2-注射类,3-皮试,4-口服
                        If NVL(mrsDrug!执行分类) = "1" Then
                            strType = "滴"
                            lngColor = COLOR_浅蓝
                        ElseIf NVL(mrsDrug!执行分类) = "2" Then
                            strType = "针"
                            lngColor = COLOR_浅青
                        ElseIf NVL(mrsDrug!执行分类) = "3" Then
                            strType = "皮"
                            lngColor = COLOR_浅红
                        ElseIf NVL(mrsDrug!执行分类) = "4" Then
                            strType = "口"
                            lngColor = COLOR_浅绿
                        Else
                            strType = " "
                            lngColor = COLOR_浅橙
                        End If
       
                        '疗程
                        If IsNull(rsDrug!天数) Then
                            '通过单量,总量计算天数
                            If NVL(rsDrug!单次用量, 0) <> 0 Then
                                lngDay = Calc缺省药品天数(NVL(rsDrug!总给予量, 0), NVL(rsDrug!单次用量, 0), _
                                        NVL(rsDrug!频率次数, 0), NVL(rsDrug!频率间隔, 0), NVL(rsDrug!间隔单位, 0), _
                                        NVL(rsDrug!剂量系数, 0), NVL(rsDrug!门诊包装, 0), _
                                        NVL(rsDrug!可否分零, 0))
                            Else
                                lngDay = 7  '未设置单量时缺省设为7天
                            End If
                            
                        Else
                            '直接取数据
                            lngDay = NVL(rsDrug!天数, 0)
                        End If
                        strDay = "(" & lngDay & "天)"
                    End If
                    '合并同组药品下的标注
                    .Cell(flexcpText, lngRow, (i - 1) * mlngSubCol + SUBCOL_标注 + 1) = strType & IIf(lngCount Mod 2 = 0, "", vbTab)
                    .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_标注 + 1) = lngColor
                    '一并给药标注需要合并
                    .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = rsDrug!医嘱内容 & ""
                    If Not blnMerge Then
                        .Cell(flexcpText, lngRow, (i - 1) * mlngSubCol + SUBCOL_疗程 + 1) = strDay  '合并列Draw_Cell处理
                    End If
                    .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_疗程 + 1) = lngDay    '用于背景色控制显示长度
              
                    
                ElseIf rsDrug!诊疗类别 & "" = "7" Then
                    '中药配方
                    If NVL(rsDrug!相关ID, 0) <> lng相关ID Then
                        lng相关ID = CLng(rsDrug!相关ID & "")
                        lngCount = lngCount + 1
                        mrsDrug.Filter = "ID=" & lng相关ID
                        .Cell(flexcpData, lngRow, lngDrugCol) = lng相关ID    '用于标识同一组药品
                        .Cell(flexcpText, lngRow, (i - 1) * mlngSubCol + SUBCOL_标注 + 1) = "草" & IIf(lngCount Mod 2 = 0, "", vbTab)
                        .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_标注 + 1) = COLOR_浅紫
                        .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = mrsDrug!医嘱内容 & ""
                        
                        .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_疗程 + 1) = 7  '缺省添满
                    Else
                        blnAddRow = False
                    End If
                    
                End If
                If blnAddRow Then
                    lngRow = lngRow + 1
                End If
                rsDrug.MoveNext
            Next
            '检查
            strFilter = "挂号单 ='" & mrs就诊时间!挂号单 & "' and 诊疗类别 ='D' "
            mrs检查检验.Filter = strFilter
            lngRow = mcolCate("_" & CATE_检查).lngBeginRow
            
            lngContW = .ColWidth(SUBCOL_名称 + 1)
            
            For k = 1 To mrs检查检验.RecordCount
                .MergeCol((i - 1) * mlngSubCol + SUBCOL_疗程 + 1) = False
                '文本内容记录下来便于气泡提示
                strContent = mrs检查检验!医嘱内容 & ""
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1) = strContent
                strContent = GetSubString(strContent, lngContW)
                .Cell(flexcpText, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = strContent
                .Cell(flexcpAlignment, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = flexAlignLeftCenter
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = mrs检查检验!ID & ""  '记录下检查医嘱组ID
                
                '已出报告 淡蓝色字体显示并加别针图标
                If Not (Val(mrs检查检验!报告ID & "") = 0 And Val(mrs检查检验!检查报告ID & "") = 0) Then
                    .Cell(flexcpForeColor, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = COLOR_鲜蓝
                    Set .Cell(flexcpPicture, lngRow, (i - 1) * mlngSubCol + SUBCOL_标注 + 1) = imgFlag.ListImages("报告").Picture
                    .Cell(flexcpPictureAlignment, lngRow, (i - 1) * mlngSubCol + SUBCOL_标注 + 1) = flexAlignCenterCenter
                    .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_标注 + 1) = mrs检查检验!报告ID & "|" & mrs检查检验!检查报告ID
                End If
                .TextMatrix(lngRow, (i - 1) * mlngSubCol + SUBCOL_疗程 + 1) = "观片" & IIf(lngRow Mod 2 = 0, "", vbTab)
                .Cell(flexcpForeColor, lngRow, (i - 1) * mlngSubCol + SUBCOL_疗程 + 1) = COLOR_鲜蓝
            
                .MergeRow(lngRow) = True
                lngRow = lngRow + 1
                mrs检查检验.MoveNext
            Next
            '检验
            strFilter = "挂号单 ='" & mrs就诊时间!挂号单 & "' and 诊疗类别 ='E' And 操作类型='6' "
            mrs检查检验.Filter = strFilter
            lngRow = mcolCate("_" & CATE_检验).lngBeginRow
            lngContW = .ColWidth(SUBCOL_名称 + 1)
                
            For k = 1 To mrs检查检验.RecordCount
                strContent = mrs检查检验!医嘱内容 & ""
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1) = strContent
                strContent = GetSubString(strContent, lngContW)
                .Cell(flexcpText, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = strContent
                .TextMatrix(lngRow, (i - 1) * mlngSubCol + SUBCOL_疗程 + 1) = IIf(lngRow Mod 2 = 0, "", vbTab)
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = mrs检查检验!ID & ""  '记录ID
                .Cell(flexcpAlignment, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = flexAlignLeftCenter
                
                 '已出报告 淡蓝色字体显示并加别针图标
                If Val(mrs检查检验!报告ID & "") <> 0 Then
                    .Cell(flexcpForeColor, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = COLOR_鲜蓝
                    Set .Cell(flexcpPicture, lngRow, (i - 1) * mlngSubCol + SUBCOL_标注 + 1) = imgFlag.ListImages("报告").Picture
                    .Cell(flexcpPictureAlignment, lngRow, (i - 1) * mlngSubCol + SUBCOL_标注 + 1) = flexAlignCenterCenter
                    .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_标注 + 1) = mrs检查检验!报告ID & ""
                End If
                
                .MergeRow(lngRow) = True
                lngRow = lngRow + 1
                mrs检查检验.MoveNext
            Next
            '其他
            
            strFilter = "挂号单 ='" & mrs就诊时间!挂号单 & "'"
            mrs其他医嘱.Filter = strFilter
            lngRow = mcolCate("_" & CATE_其他医嘱).lngBeginRow
            lngContW = .ColWidth(SUBCOL_名称 + 1)
            
            For k = 1 To mrs其他医嘱.RecordCount
                strContent = mrs其他医嘱!医嘱内容 & ""
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1) = strContent
                strContent = GetSubString(strContent, lngContW)
                .Cell(flexcpText, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = strContent
                .Cell(flexcpAlignment, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_疗程 + 1) = flexAlignLeftCenter
                .MergeRow(lngRow) = True
                lngRow = lngRow + 1
                mrs其他医嘱.MoveNext
            Next
            
        
            '病历
            strFilter = "挂号ID=" & mrs就诊时间!就诊id
            mrs病历文书.Filter = strFilter
            lngRow = mcolCate("_" & CATE_病历文书).lngBeginRow
            lngContW = .ColWidth(SUBCOL_名称 + 1)
            '旧版病历
            For k = 1 To mrs病历文书.RecordCount
                strContent = mrs病历文书!病历名称 & ""
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1) = strContent
                strContent = GetSubString(strContent, lngContW)
                .Cell(flexcpText, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = strContent
                .Cell(flexcpAlignment, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = flexAlignLeftCenter
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = mrs病历文书!ID & ""    '记录下病历记录ID
                
                .MergeRow(lngRow) = True
                lngRow = lngRow + 1
                mrs病历文书.MoveNext
            Next
            '新版病历
            mrsEMR.Filter = "挂号ID=" & mrs就诊时间!就诊id
            
            For k = 1 To mrsEMR.RecordCount
                strContent = mrsEMR!病历名称 & ""
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1) = strContent
                strContent = GetSubString(strContent, lngContW)
                .Cell(flexcpText, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = strContent
                .Cell(flexcpAlignment, lngRow, (i - 1) * mlngSubCol + SUBCOL_图标 + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = flexAlignLeftCenter
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = mrsEMR!ID & ""    '记录下病历记录ID
                .MergeRow(lngRow) = True
                lngRow = lngRow + 1
                mrsEMR.MoveNext
            Next
            mrs就诊时间.MovePrevious
        Next
        .Redraw = flexRDDirect
        .Row = 0 '获取焦点
    End With

    
End Sub

Private Sub ReadRegister()
'功能:读取登记信息
    Dim strSql As String
    Dim i As Long, j As Long
    Dim rsEmr As ADODB.Recordset
    Dim strMsg As String
    
    On Error GoTo errH
    '1-就诊时间
    strSql = "Select Rownum As 序号, b.就诊ID,b.挂号单,b.执行时间,b.执行状态,b.执行部门ID,b.科室名称  " & vbNewLine & _
            "From (Select a.Id As 就诊id,a.NO as 挂号单,a.执行时间,a.执行状态,a.执行部门ID,d.名称 as 科室名称 " & vbNewLine & _
            "       From 病人挂号记录 A,部门表 D " & vbNewLine & _
            "       Where a.执行部门ID =d.Id(+) and a.病人id = [1] And a.记录性质 = 1 And a.记录状态 = 1 " & IIf(mbytShow = 0, "", " And a.执行部门ID =[2]") & vbNewLine & _
            "       Order By a.执行时间 Desc) B"
    
    Set mrs就诊时间 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng科室ID)
    
    '2-诊断记录
    mrs就诊时间.Filter = "序号 >=" & mlngNext & " And 序号 <= " & mlngPrev
    
    mstr挂号单 = "": mstr就诊ID = ""
    For i = 1 To mrs就诊时间.RecordCount
        mstr就诊ID = mstr就诊ID & "," & mrs就诊时间!就诊id
        mstr挂号单 = mstr挂号单 & "," & mrs就诊时间!挂号单
        mrs就诊时间.MoveNext
    Next
    mstr就诊ID = mstr就诊ID & ","
    mstr挂号单 = mstr挂号单 & ","
    
    strSql = "Select a.主页id As 就诊id, a.诊断描述" & vbNewLine & _
            "From 病人诊断记录 A" & vbNewLine & _
            "Where 病人id = [1] And Instr([2], ',' || 主页id || ',') > 0 " & vbNewLine & _
            " order by a.主页id,a.诊断次序"
            
    Set mrs诊断 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mstr就诊ID)
    
    '3-医嘱记录
    strSql = "Select a.挂号单, a.Id, a.相关id, a.序号, a.医嘱期效, a.医嘱内容, a.标本部位, a.诊疗类别, a.诊疗项目id, a.天数, a.单次用量, a.总给予量, a.执行频次, b.操作类型, b.执行分类," & vbNewLine & _
            "       a.频率次数, a.频率间隔, a.间隔单位, b.计算单位 As 单量单位, c.剂量系数, c.门诊包装, c.门诊可否分零 As 可否分零 " & vbNewLine & _
            "From 病人医嘱记录 A, 诊疗项目目录 B, 药品规格 C" & vbNewLine & _
            "Where a.诊疗项目id = b.Id And a.收费细目id = c.药品id(+) And a.病人id = [1] And a.医嘱状态 = 8 And" & vbNewLine & _
            "      (a.诊疗类别 In ('5', '6', '7') Or (a.诊疗类别 = 'E' And b.操作类型 In ('1', '2', '3', '4'))) And" & vbNewLine & _
            "      Instr([2], ',' || a.挂号单 || ',') > 0" & vbNewLine & _
            "Order By a.挂号单, a.序号"

    Set mrsDrug = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mstr挂号单)
    '检查检验
    strSql = "Select a.挂号单,a.ID,a.医嘱内容,a.诊疗类别 ,b.操作类型, Max(c.病历id) As 报告id, Max(c.检查报告id) As 检查报告id," & vbNewLine & _
        "       Decode(Max(Nvl(c.查阅状态, 0)), Min(Nvl(c.查阅状态, 0)), Max(Nvl(c.查阅状态, 0)), 2) As 查阅状态" & vbNewLine & _
        "From 病人医嘱记录 A, 诊疗项目目录 B, 病人医嘱报告 C" & vbNewLine & _
        "Where a.诊疗项目id = b.Id And a.Id = c.医嘱id(+) And a.病人id = [1] And a.医嘱状态 = 8 And" & vbNewLine & _
        "      (a.诊疗类别 = 'D' And a.相关id Is Null Or a.诊疗类别 = 'E' And b.操作类型 = '6') And" & vbNewLine & _
        "      Instr([2], ',' || a.挂号单 || ',') > 0" & vbNewLine & _
        "Group By a.挂号单,a.ID,a.医嘱内容,a.序号,a.诊疗类别, b.操作类型" & vbNewLine & _
        "Order By a.挂号单, a.序号"

    Set mrs检查检验 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mstr挂号单)
    '其他医嘱
    '1-排除 治疗类E：2-给药方法(西药);3-中药煎法;4-中药用(服)法;6-采集方法(检验);8-输血途径 ;只加载;0-普通,1-过敏试验,5-特殊治疗
    '2-排除 麻醉,麻醉同手术一并使用
    '3-排除 检验,检查
    '4-排除 5-西药,6-中成药,7-草药
    '5-手术\输血 只加载主医嘱行 相关ID IS NULL
    '6-其他类 Z
    strSql = "Select a.挂号单, a.医嘱内容, a.诊疗类别" & vbNewLine & _
        "From 病人医嘱记录 A, 诊疗项目目录 B" & vbNewLine & _
        "Where a.诊疗项目id = b.Id And a.病人id = [1] And a.医嘱状态 = 8 And Not a.诊疗类别 In ('G', 'D', 'C', '5', '6', '7') And" & vbNewLine & _
        "      Not (NVL(b.操作类型,0) In ('2', '3', '4', '6', '8') And a.诊疗类别 = 'E') And NVL(相关id,0)=0 And" & vbNewLine & _
        "      Instr([2], ',' || a.挂号单 || ',') > 0" & vbNewLine & _
        "Order By a.挂号单, a.序号"
    Set mrs其他医嘱 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mstr挂号单)
    
    '门诊病历
    strSql = "Select ID , Nvl(主页id, 0) as 挂号ID,病历种类, 病历名称 " & vbNewLine & _
            "From 电子病历记录" & vbNewLine & _
            "Where 病人来源 = 1 And (病历种类 In (1, 6) Or (病历种类 = 5 And 编辑方式 <> 2)) And 病人id = [1] And Instr([2], ',' || Nvl(主页id, 0) || ',') > 0 " & vbNewLine & _
            "Order By  Nvl(主页id, 0),病历种类, 序号, 创建时间"
    Set mrs病历文书 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mstr就诊ID)
    '新版病历
    Set mrsEMR = InitRS
    If Not gobjEmr Is Nothing Then
        mrs就诊时间.MoveFirst
        For i = 1 To mrs就诊时间.RecordCount
            '新版病历提供接口：GetOutEPRRecord(挂号ID)返回每次就诊的病历情况（ID,Title）。
            On Error Resume Next
            strMsg = gobjEmr.GetOutEPRRecord(mrs就诊时间!就诊id & "", rsEmr)
            err.Clear: On Error GoTo 0
            If Not rsEmr Is Nothing Then
                For j = 1 To rsEmr.RecordCount
                    mrsEMR.AddNew
                    mrsEMR!挂号ID = CLng(mrs就诊时间!就诊id & "")
                    mrsEMR!ID = rsEmr!ID
                    mrsEMR!病历名称 = rsEmr!Title
                    mrsEMR.Update
                    rsEmr.MoveNext
                Next
            End If
            mrs就诊时间.MoveNext
        Next
    End If
    If mrsEMR.RecordCount > 0 Then mrsEMR.MoveFirst
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SetViewRows() As Integer
'功能:设置视图显示总行数
    Dim udtCate As TYPE_CATE
    Dim arrTmp As Variant
    Dim lngTmp As Long
    Dim lngRows As Long
    Dim i As Long
    
    Set mcolCate = New Collection
    arrTmp = Split(mstr分类, "|")
    lngRows = 0
    For i = CATE_就诊时间 To CATE_其他医嘱
        '分类之间再增加一行作为间隔行
        If i >= CATE_诊断 And i <= CATE_其他医嘱 Then
            lngRows = lngRows + 1
        End If
        
        lngTmp = CalculateMaxRows(i)
        
        udtCate.strName = arrTmp(i)
        udtCate.lngBeginRow = lngRows
        udtCate.lngEndRow = lngRows + lngTmp - 1
        
        mcolCate.Add udtCate, "_" & i
        lngRows = lngRows + lngTmp
    Next

    SetViewRows = lngRows
End Function

Private Function CalculateMaxRows(ByVal bytFun As Byte) As Integer
'---------------------------------------------------------------------------------------------
'功能:计算最大行数
'参数:返回各分类应该显示行数
'---------------------------------------------------------------------------------------------
    Dim intNum As Integer
    Dim intMaxNum As Integer
    Dim i As Long, j As Long
    
    Dim blnDo As Boolean
    
    Dim str挂号单 As String
    Dim str相关ID As String
    Dim strErr As String
    
    
    Dim rsEmr As ADODB.Recordset
    Dim arrTmp As Variant
    
    blnDo = Not mrs就诊时间 Is Nothing
    
    Select Case bytFun
    Case CATE_就诊时间, CATE_诊断
        intMaxNum = 1
    Case CATE_用药跟踪
        intMaxNum = 10
        If blnDo Then
            mrsDrug.Filter = "诊疗类别 ='5' or 诊疗类别 ='6' or 诊疗类别 ='7'"
            For j = 1 To mrsDrug.RecordCount
                If str挂号单 <> mrsDrug!挂号单 & "" Then
                    str挂号单 = mrsDrug!挂号单
                    If intMaxNum < intNum Then intMaxNum = intNum '记录最大行
                    intNum = 0
                End If
                '-中药配方,多个药品只显示一行
                If NVL(mrsDrug!诊疗类别, "") = "7" And str相关ID <> mrsDrug!相关ID & "" Then  '中药配方
                    str相关ID = mrsDrug!相关ID & ""       '标记用于判断
                    intNum = intNum + 1
                ElseIf InStr(",5,6,", mrsDrug!诊疗类别 & "") > 1 Then  '西药及成药
                '-一并给药,有几个药品就占用几行
                    intNum = intNum + 1
                End If
                '最后一行时,再一次记录最大行
                If j = mrsDrug.RecordCount Then
                    If intMaxNum <= intNum Then intMaxNum = intNum + 1 '记录最大行
                End If
                mrsDrug.MoveNext
            Next
        End If
     Case CATE_检查
        intMaxNum = 5
        If blnDo Then
            mrs就诊时间.MoveFirst
            For j = 1 To mrs就诊时间.RecordCount
                mrs检查检验.Filter = "挂号单 ='" & mrs就诊时间!挂号单 & "' And 诊疗类别 ='D'"
                If intMaxNum <= mrs检查检验.RecordCount Then intMaxNum = mrs检查检验.RecordCount + 1
                
                mrs就诊时间.MoveNext
            Next
        End If
     Case CATE_病历文书
        intMaxNum = 5
        '病历
        If blnDo Then
            mrs就诊时间.MoveFirst
            For j = 1 To mrs就诊时间.RecordCount
                mrs病历文书.Filter = "挂号ID =" & mrs就诊时间!就诊id
                mrsEMR.Filter = "挂号ID =" & mrs就诊时间!就诊id
                If intMaxNum <= (mrs病历文书.RecordCount + mrsEMR.RecordCount) Then intMaxNum = (mrs病历文书.RecordCount + mrsEMR.RecordCount) + 1
                mrs就诊时间.MoveNext
            Next
        End If
    Case CATE_检验
        intMaxNum = 5
        If blnDo Then
            mrs就诊时间.MoveFirst
            For j = 1 To mrs就诊时间.RecordCount
                mrs检查检验.Filter = "挂号单 ='" & mrs就诊时间!挂号单 & "' And 诊疗类别 ='E' And 操作类型 = '6'"   '检验显示采集方法行
                If intMaxNum <= mrs检查检验.RecordCount Then intMaxNum = mrs检查检验.RecordCount + 1
                
                mrs就诊时间.MoveNext
            Next
        End If
    Case CATE_其他医嘱
        intMaxNum = 2
        If blnDo Then
            mrs就诊时间.MoveFirst
            For j = 1 To mrs就诊时间.RecordCount
                mrs其他医嘱.Filter = "挂号单 ='" & mrs就诊时间!挂号单 & "'"
                If intMaxNum <= mrs其他医嘱.RecordCount Then intMaxNum = mrs其他医嘱.RecordCount + 1
                mrs就诊时间.MoveNext
            Next
        End If
    End Select
    CalculateMaxRows = intMaxNum
End Function

Private Sub SubRefresh(Optional ByVal Index As Integer = -1)
'功能:刷新

    If Index <> -1 Then
        If Index = CMD_PREV Then
            If imgBtn(CMD_PREV).Enabled = False Then Exit Sub
            mlngPrev = mlngPrev + 1
            mlngNext = mlngNext + 1
        ElseIf Index = CMD_NEXT Then
            If imgBtn(CMD_NEXT).Enabled = False Then Exit Sub
            mlngPrev = mlngPrev - 1
            mlngNext = mlngNext - 1
        ElseIf Index = CMD_ORTHER Then
            mlngPrev = 3
            mlngNext = 1
            If mbytShow = 1 Then
                Set imgBtn(CMD_ORTHER).Picture = imgFlag.ListImages("显示所有就诊").Picture
                 mbytShow = 0
            Else
                Set imgBtn(CMD_ORTHER).Picture = imgFlag.ListImages("只显示本科").Picture
                mbytShow = 1
            End If
        End If
    Else
        If CheckRegister Then
            imgBtn(CMD_ORTHER).Visible = True
        Else
            imgBtn(CMD_ORTHER).Visible = False
        End If
    End If
    
    mstrFontUnderLine = ""
    
    Call ReadRegister
    Call InitVsView
    Call ResizeVsView
    Call LoadView
    mrs就诊时间.Filter = ""
    imgBtn(CMD_PREV).Enabled = mlngPrev < mrs就诊时间.RecordCount
    imgBtn(CMD_NEXT).Enabled = mlngNext > 1
 
    Set imgBtn(CMD_PREV).Picture = IIf(imgBtn(CMD_PREV).Enabled, imgFlag.ListImages("上次可用").Picture, imgFlag.ListImages("上次不可用").Picture)
    Set imgBtn(CMD_NEXT).Picture = IIf(imgBtn(CMD_NEXT).Enabled, imgFlag.ListImages("下次可用").Picture, imgFlag.ListImages("下次不可用").Picture)
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyRight Then
        Call SubRefresh(CMD_NEXT)
    ElseIf KeyCode = vbKeyLeft Then
        Call SubRefresh(CMD_PREV)
    End If
End Sub

Private Sub Form_Load()
    '加载界面
    '表格初始化
    mbytFontSize = 9
    mstrFontUnderLine = ""
    imgBtn(CMD_PREV).Enabled = False
    imgBtn(CMD_NEXT).Enabled = False
    imgBtn(CMD_ORTHER).Visible = False
    Set imgBtn(CMD_PREV).Picture = IIf(imgBtn(CMD_PREV).Enabled, imgFlag.ListImages("上次可用").Picture, imgFlag.ListImages("上次不可用").Picture)
    Set imgBtn(CMD_NEXT).Picture = IIf(imgBtn(CMD_NEXT).Enabled, imgFlag.ListImages("下次可用").Picture, imgFlag.ListImages("下次不可用").Picture)
    Set imgBtn(CMD_ORTHER).Picture = IIf(mbytShow = 0, imgFlag.ListImages("只显示本科").Picture, imgFlag.ListImages("显示所有就诊").Picture)
    '表格缺省设置
    
    Call InitVsView
End Sub

Private Sub InitVsView()
'--------------------------------------------------------------------------------------------------------------------------------------------
'功能:初始化表格对象
'--------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim arrTmp As Variant
    
    With vsView
        .Cols = 0: .Rows = 0
        .Cols = 3 * mlngSubCol + 1
        .Rows = SetViewRows
        .FixedRows = mcolCate("_" & CATE_诊断).lngEndRow + 1  '就诊时间,诊断
        .FixedCols = 1
        .ExtendLastCol = False
        .RowHeightMin = 300
        .BackColorFixed = COLOR_白色
        .FixedAlignment(0) = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignCenterCenter
        .MergeCol(0) = True
        .MergeRow(mcolCate("_" & CATE_就诊时间).lngBeginRow) = True
        .MergeRow(mcolCate("_" & CATE_诊断).lngBeginRow) = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .MergeCells = flexMergeFree
        .MergeCellsFixed = flexMergeRestrictRows
        .MergeCompare = flexMCExact
        .SelectionMode = flexSelectionFree
        .GridLines = flexGridFlat
 
        '加载分类列
        For i = 0 To mcolCate.Count - 1
            .Cell(flexcpText, mcolCate("_" & i).lngBeginRow, 0, mcolCate("_" & i).lngEndRow, 0) = mcolCate("_" & i).strName
            .Cell(flexcpFontBold, mcolCate("_" & i).lngBeginRow, 0, mcolCate("_" & i).lngEndRow, 0) = flexcpFontBold
            .Cell(flexcpFontSize, mcolCate("_" & i).lngBeginRow, 0, mcolCate("_" & i).lngEndRow, 0) = mbytFontSize
            .Cell(flexcpBackColor, mcolCate("_" & i).lngBeginRow, 0, mcolCate("_" & i).lngEndRow, 0) = &H8000000F
        Next
    End With
End Sub

Private Sub InitFrom()
'功能:表格宽度固定
    Dim lngWFrm As Long
    Dim lngHTop As Long
    
    On Error Resume Next
    '界面尺寸、大小
    PicCenter.Move 0, 0, Me.Width, Me.Height
    lngWFrm = PicCenter.Width
    If lngWFrm < 7035 Then
        lngWFrm = 7035
    End If
    
    lngHTop = 1000
    picTop.Move 0, 0, lngWFrm, lngHTop
    picBottom.Move 0, lngHTop + 45, lngWFrm, Me.Height - lngHTop - 45
    vsView.Move 0, 0, picBottom.Width, picBottom.Height
    imgBtn(CMD_NEXT).Width = 1050: imgBtn(CMD_NEXT).Height = 610
    imgBtn(CMD_PREV).Width = 1050: imgBtn(CMD_PREV).Height = 610
    imgBtn(CMD_ORTHER).Width = 2250: imgBtn(CMD_ORTHER).Height = 610
    
    lbl(0).Move lngWFrm / 2 - lbl(0).Width / 2, lngHTop / 2
    imgBtn(CMD_PREV).Move lngWFrm - (imgBtn(CMD_PREV).Width + imgBtn(CMD_NEXT).Width + 200), 850
    imgBtn(CMD_NEXT).Move lngWFrm - (imgBtn(CMD_NEXT).Width + 100), 850
    imgBtn(CMD_ORTHER).Move 45, 850
    
    '背景颜色设置
    PicCenter.BackColor = COLOR_CENTERBK
    Me.BackColor = COLOR_FORMBK      '窗口背景白色
    'VS重新布局 刷新界面重新截取字符
    If Me.Visible Then
        Call LoadView
        Call ResizeVsView
    End If
End Sub

Private Sub Form_Resize()
    Dim lngSpace As Long
    
    On Error Resume Next
    Call InitFrom
End Sub

Private Sub imgBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strKey As String
    Dim strRePic As String
    
    If imgBtn(Index).Enabled = False Then Exit Sub
    
    Select Case Index
    Case CMD_PREV
        strKey = "上次高亮"
    Case CMD_NEXT
        strKey = "下次高亮"
    Case CMD_ORTHER
        If mbytShow = 0 Then
            strKey = "只显示本科高亮"
        Else
            strKey = "显示所有就诊高亮"
        End If
    End Select
    Set imgBtn(Index).Picture = imgFlag.ListImages(strKey).Picture
    mintIndex = Index
End Sub

Private Sub ShowTipInfo(ByVal objHwnd As Long, ByVal strInfo As String)
    If strInfo <> "" Then
        Call zlCommFun.ShowTipInfo(objHwnd, strInfo, True, , 4500)
    Else
        Call zlCommFun.ShowTipInfo(0, strInfo)
    End If
End Sub


Private Sub imgBtn_Click(Index As Integer)
    SubRefresh Index
End Sub

Private Sub ResizeVsView()
    Dim i As Long, j As Long
    Dim lngMainW As Long
    Dim lngSubW As Long
    Dim lngCount As Long
    Dim lngRow As Long
    Dim strSpace As String
    Dim objImage As Object
    
    Dim udtCate As TYPE_CATE
    
    On Error Resume Next
    
    With vsView
        '设置列宽
        strSpace = "  "    '一个空格
        .FontSize = mbytFontSize
        .RowHeightMin = IIf(mbytFontSize = 9, 300, 400)
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, .Cols - 1) = mbytFontSize
        .ColWidth(0) = IIf(mbytFontSize = 9, 1000, 1200)
        lngMainW = (.Width - .ColWidth(0) - 480) / 3
        If lngMainW < 2800 Then lngMainW = 2800
        For i = 1 To .Cols - 1
            Select Case i Mod mlngSubCol
            Case SUBCOL_标注 + 1 '标注列（滴|口|针）
                .ColWidth(i) = 300
                .ColAlignment(i) = flexAlignCenterCenter
                .MergeCol(i) = True      '允许标注列合并
            Case SUBCOL_图标 + 1  '一并给药标识列(┏┃┗)
                .ColWidth(i) = IIf(mbytFontSize = 9, 200, 250)
                .ColAlignment(i) = flexAlignRightCenter
            Case SUBCOL_名称 + 1   '药品名
                .ColWidth(i) = lngMainW - IIf(mbytFontSize = 9, 1160, 1215)
                .ColAlignment(i) = flexAlignLeftCenter
            Case SUBCOL_疗程 + 1  '疗程信息
                 .ColWidth(i) = 700
                .ColAlignment(i) = flexAlignCenterCenter
            Case Else
                '间隔列
                .ColWidth(i) = 15
                Call .Select(mcolCate("_" & CATE_病历文书).lngBeginRow, i, mcolCate("_" & CATE_其他医嘱).lngEndRow, i)
                Call .CellBorder(.GridColorFixed, 1, 1, 1, 1, -1, -1)
            End Select
            
        Next
        '间隔行处理
        For i = CATE_诊断 To CATE_其他医嘱
            lngRow = mcolCate("_" & i).lngBeginRow - 1
            .RowHidden(lngRow) = True
        Next
        
        '诊断与就诊时间行合并处理
        For i = 1 To 3
            udtCate = mcolCate("_" & CATE_就诊时间)
            If .Cell(flexcpData, udtCate.lngBeginRow, (i - 1) * mlngSubCol + SUBCOL_名称 + 1) = "" Then   '挂号单为空时代表该列没有数据
                udtCate = mcolCate("_" & CATE_就诊时间)
                .Cell(flexcpText, udtCate.lngBeginRow, (i - 1) * 5 + 1, udtCate.lngEndRow, i * 5) = IIf(i Mod 2, strSpace, strSpace & strSpace)
                udtCate = mcolCate("_" & CATE_诊断)
                .Cell(flexcpText, udtCate.lngBeginRow, (i - 1) * 5 + 1, udtCate.lngEndRow, i * 5) = IIf(i Mod 2, strSpace, strSpace & strSpace)
            End If
        Next
        
        udtCate = mcolCate("_" & CATE_就诊时间)
        .Cell(flexcpAlignment, udtCate.lngBeginRow, 1, udtCate.lngEndRow, .Cols - 1) = flexAlignCenterCenter   '就诊时间
        udtCate = mcolCate("_" & CATE_诊断)
        .Cell(flexcpAlignment, udtCate.lngBeginRow, 1, udtCate.lngEndRow, .Cols - 1) = flexAlignLeftCenter     '诊断
        
        
        '用药跟踪
        udtCate = mcolCate("_" & CATE_用药跟踪)
        Call .Select(udtCate.lngEndRow, 1, udtCate.lngEndRow, .Cols - 1)
        Call .CellBorder(.GridColorFixed, -1, -1, -1, 1, -1, -1)
        '病历文书
        udtCate = mcolCate("_" & CATE_病历文书)
        Call .Select(udtCate.lngEndRow, 1, udtCate.lngEndRow, .Cols - 1)
        Call .CellBorder(.GridColorFixed, -1, -1, -1, 1, -1, -1)
        '检查
        udtCate = mcolCate("_" & CATE_检查)
        Call .Select(udtCate.lngEndRow, 1, udtCate.lngEndRow, .Cols - 1)
        Call .CellBorder(.GridColorFixed, -1, -1, -1, 1, -1, -1)
        '检验
        udtCate = mcolCate("_" & CATE_检验)
        Call .Select(udtCate.lngEndRow, 1, udtCate.lngEndRow, .Cols - 1)
        Call .CellBorder(.GridColorFixed, -1, -1, -1, 1, -1, -1)


        '最后一列
        Call .Select(0, 0, .Rows - 1, .Cols - 1)
        Call .CellBorder(.GridColorFixed, -1, -1, 1, 1, -1, -1)

        .AutoSize 0, .Cols - 1, , 45
        .Row = 0   '获得焦点
        
    End With
End Sub

Private Sub picTop_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If mintIndex >= 0 Then
        SetImageDefault
    End If
End Sub

Private Sub vsView_Click()
    Dim lng医嘱ID As Long
    Dim lng就诊ID As Long
    Dim lng报告ID As Long
    Dim str检查报告ID As String
    Dim lngCol As Long
    Dim strMsg As String
    Dim strTmp As String
    
    Dim blnMoved As Boolean
    With vsView
         If .Col <= .FixedCols - 1 Then Exit Sub
         If .Row <= .FixedRows - 1 Then Exit Sub
         .Redraw = flexRDNone
         lngCol = IIf(.Col Mod mlngSubCol = 0, .Col - mlngSubCol, (.Col \ mlngSubCol) * mlngSubCol)
         If .Row >= mcolCate("_" & CATE_检查).lngBeginRow And .Row <= mcolCate("_" & CATE_检查).lngEndRow Then
            lng就诊ID = CLng(.ColData(lngCol + SUBCOL_名称 + 1))
            lng医嘱ID = CLng(.Cell(flexcpData, .Row, lngCol + SUBCOL_名称 + 1))
            If .Col Mod mlngSubCol = SUBCOL_疗程 + 1 And .Cell(flexcpFontUnderline, .Row, .Col) = True Then
                '观片
                mrs就诊时间.Filter = "就诊ID=" & lng就诊ID
                If NVL(mrs就诊时间!执行状态, 0) = 1 Then '完成就诊的情况,检查数据是否转出
                    blnMoved = zlDatabase.NOMoved("病人挂号记录", mrs就诊时间!挂号单 & "")
                End If
               
                If CreateObjectPacs(gobjPublicPacs) Then
                    Call gobjPublicPacs.ShowImage(lng医嘱ID, Me, blnMoved)
                End If
            ElseIf .Col Mod mlngSubCol = SUBCOL_名称 And Not .Cell(flexcpPicture, .Row, lngCol + SUBCOL_标注 + 1) Is Nothing Then
                '查阅报告
                strTmp = .Cell(flexcpData, .Row, lngCol + SUBCOL_标注 + 1)
                lng报告ID = CLng(Split(strTmp, "|")(0))
                str检查报告ID = Split(strTmp, "|")(1)
                Call FuncEPRReport(Me, lng医嘱ID, "D", lng报告ID, str检查报告ID, 1)
            End If
         ElseIf .Row >= mcolCate("_" & CATE_检验).lngBeginRow And .Row <= mcolCate("_" & CATE_检验).lngEndRow Then
            If .Col Mod mlngSubCol = SUBCOL_名称 And Not .Cell(flexcpPicture, .Row, lngCol + SUBCOL_标注 + 1) Is Nothing Then
                lng医嘱ID = CLng(.Cell(flexcpData, .Row, lngCol + SUBCOL_名称 + 1))
                lng报告ID = CLng(.Cell(flexcpData, .Row, lngCol + SUBCOL_标注 + 1))
                Call FuncEPRReport(Me, lng医嘱ID, "", lng报告ID, , 1)
            End If
         ElseIf .Row >= mcolCate("_" & CATE_病历文书).lngBeginRow And .Row <= mcolCate("_" & CATE_病历文书).lngEndRow And .MousePointer = flexCustom Then
             
            strTmp = CStr(.Cell(flexcpData, .Row, lngCol + SUBCOL_名称 + 1))
            If strTmp = "" Then Exit Sub
            If Len(strTmp) < 32 Then
                lng报告ID = CLng(strTmp) '老版病历查看
                Call gobjRichEPR.ViewDocument(Me, lng报告ID, False)
            ElseIf Len(strTmp) = 32 And Not gobjEmr Is Nothing Then
                '新版病历
                On Error Resume Next
                strMsg = gobjEmr.OpenOutEPR(strTmp)
                err.Clear: On Error GoTo 0
            End If
         End If
         .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsView_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngColor As Long, lngclrg, k As Long, n As Long
    Dim r1 As Integer, g1 As Integer, b1 As Integer
    Dim r2 As Integer, g2 As Integer, b2 As Integer
    Dim rg As Integer, gg As Integer, bg As Integer
    
    Dim lngFontW As Long
    Dim lng组ID As Long
    Dim strContent As String
    
    Dim vRect As RECT, vRect1 As RECT, vRect2 As RECT

    If mcolCate Is Nothing Then Exit Sub
    
    With vsView
        If .RowHidden(Row) = True Then Exit Sub
        '分类行背景色处理
        If mcolCate("_" & CATE_就诊时间).lngBeginRow <= Row And Row <= mcolCate("_" & CATE_其他医嘱).lngEndRow And Col = 0 Then
            '获取矩形框
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right - 1
            vRect.Bottom = Bottom - 1
            'draw frame
            lngColor = SetBkColor(hDC, 0)
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, k

            ' get colors
            r1 = 250: g1 = 250: b1 = 250   '渐变起始
            r2 = 229: g2 = 229: b2 = 229   '渐变终止
            ' show color
            vRect2 = vRect
            vRect2.Bottom = vRect.Bottom - (vRect.Bottom - vRect.Top) / 2
            vRect1 = vRect2

            For k = vRect2.Top To vRect2.Bottom
                rg = r1 + (k - vRect2.Top) * (r2 - r1) / (vRect2.Bottom - vRect2.Top)
                gg = g1 + (k - vRect2.Top) * (g2 - g1) / (vRect2.Bottom - vRect2.Top)
                bg = b1 + (k - vRect2.Top) * (b2 - b1) / (vRect2.Bottom - vRect2.Top)
                lngclrg = RGB(rg, gg, bg)
                SetBkColor hDC, lngclrg
                vRect1.Top = k
                ExtTextOut hDC, vRect1.Left, vRect1.Top, ETO_OPAQUE, vRect1, " ", 1, k
            Next
            ' get colors
            r1 = 229: g1 = 229: b1 = 229   '渐变起始
            r2 = 250: g2 = 250: b2 = 250   '渐变终止
            ' show color
            vRect2 = vRect
            vRect2.Top = vRect.Top + (vRect.Bottom - vRect.Top) / 2
            vRect1 = vRect2
            For k = vRect2.Top To vRect2.Bottom
                rg = r1 + (k - vRect2.Top) * (r2 - r1) / (vRect2.Bottom - vRect2.Top)
                gg = g1 + (k - vRect2.Top) * (g2 - g1) / (vRect2.Bottom - vRect2.Top)
                bg = b1 + (k - vRect2.Top) * (b2 - b1) / (vRect2.Bottom - vRect2.Top)
                lngclrg = RGB(rg, gg, bg)
                SetBkColor hDC, lngclrg
                vRect1.Top = k
                ExtTextOut hDC, vRect1.Left, vRect1.Top, ETO_OPAQUE, vRect1, " ", 1, k
            Next

            SetBkColor hDC, lngColor
            '将单元格字体绘到矩形区域
            strContent = .Cell(flexcpText, Row, Col)
            lblW.Caption = strContent: lblW.AutoSize = True
            vRect1.Top = vRect.Top + (vRect.Bottom - vRect.Top) / 2 - (lblW.Height / 2) / Screen.TwipsPerPixelY

            vRect1.Left = vRect.Left + (vRect.Right - vRect.Left) / 2 - (lblW.Width / 2) / Screen.TwipsPerPixelX

            TextOut hDC, vRect1.Left, vRect1.Top, strContent, LenB(StrConv(strContent, vbFromUnicode))
        End If

        If Not (Col >= 1 And Col < vsView.Cols - 1) Then Exit Sub
        If mcolCate("_" & CATE_用药跟踪).lngBeginRow <= Row And Row <= mcolCate("_" & CATE_用药跟踪).lngEndRow Then
            If Col Mod mlngSubCol = SUBCOL_标注 Then Exit Sub
           '清楚右边线
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
            If Col Mod mlngSubCol = SUBCOL_名称 + 1 Then
                If .Cell(flexcpData, Row, (Col \ mlngSubCol) * mlngSubCol + SUBCOL_名称 + 1) <> "" Then
                    vRect.Left = Left
                    vRect.Top = Top + 1
                    vRect.Right = Left + (Right - Left) / 7 * Val(.Cell(flexcpData, Row, (Col \ mlngSubCol) * mlngSubCol + SUBCOL_疗程 + 1))
                    vRect.Bottom = Bottom - 2
                    SetBkColor hDC, OS.SysColor2RGB(.Cell(flexcpData, Row, (Col \ mlngSubCol) * mlngSubCol + SUBCOL_标注 + 1))
                    '设置矩形区域背景色
                    ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, "", 0, 0
                     '恢复背景色为窗体背景
                    SetBkColor hDC, OS.SysColor2RGB(.BackColor)
                    vRect.Left = Left + 1
                    vRect.Top = Top + 4
                    vRect.Right = Right - 1
                    vRect.Bottom = Bottom - 1

                    '字体长度超过行宽时 截取部分+“...”的方式显示
                    strContent = .Cell(flexcpData, Row, Col)
                    lngFontW = .Cell(flexcpWidth, Row, Col)
                    strContent = GetSubString(strContent, lngFontW)
                    '将单元格字体绘到矩形区域
                    TextOut hDC, vRect.Left, vRect.Top, strContent, LenB(StrConv(.Cell(flexcpData, Row, Col), vbFromUnicode))
                End If
            ElseIf Col Mod mlngSubCol = SUBCOL_疗程 + 1 Then
                 '一并给药疗程合并显示,要求从下向上合并处理,否则合并内容被下面表格掩盖
                If .Cell(flexcpText, Row, (Col \ mlngSubCol) * mlngSubCol + SUBCOL_图标 + 1) = "┗" Then
                    lngEnd = 0
                    lng组ID = CLng(.Cell(flexcpData, Row, (Col \ mlngSubCol) * mlngSubCol + SUBCOL_图标 + 1))
                    For k = 1 To .Rows - 1
                        If lng组ID <> CLng(.Cell(flexcpData, Row - k, (Col \ mlngSubCol) * mlngSubCol + SUBCOL_图标 + 1)) Then
                             k = k - 1
                             Exit For
                        End If
                    Next
                    
                    vRect.Top = Top - (Bottom - Top) * k
                    vRect.Left = Left
                    vRect.Right = Right
                    vRect.Bottom = Bottom - 1
                    If vRect.Top < (.RowPos(mcolCate("_" & CATE_诊断).lngEndRow) / Screen.TwipsPerPixelY + (Bottom - Top)) Then
                    '合并矩形框超过固定行时,取固定行边缘值
                        vRect.Top = (.RowPos(mcolCate("_" & CATE_诊断).lngEndRow) / Screen.TwipsPerPixelY + (Bottom - Top))
                    End If
                    '设置矩形区域背景色
                    ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, "", 0, 0
                    
                    strContent = "(" & .Cell(flexcpData, Row, Col) & "天)"
                    lblW.Caption = strContent: lblW.AutoSize = True
                    vRect.Top = vRect.Top + (vRect.Bottom - vRect.Top) / 2 - (lblW.Height / 2) / Screen.TwipsPerPixelY
                    vRect.Left = vRect.Left + (vRect.Right - vRect.Left) / 2 - (lblW.Width / 2) / Screen.TwipsPerPixelX
                    
                    TextOut hDC, vRect.Left, vRect.Top, strContent, LenB(StrConv(strContent, vbFromUnicode))
                    
                End If
            End If
        End If

        If mcolCate("_" & CATE_病历文书).lngBeginRow <= Row And Row <= mcolCate("_" & CATE_其他医嘱).lngEndRow Then
            If Col Mod mlngSubCol = SUBCOL_标注 Then Exit Sub
           '清楚右边线
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom - 1

            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

        End If

    End With
End Sub

Private Sub vsView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngColor As Long
    
    Dim strInfo As String
    
    Dim arrTmp As Variant
    
    With vsView
        If mcolCate Is Nothing Then Exit Sub
        lngRow = .MouseRow: lngCol = .MouseCol
        If lngRow = -1 Or lngCol = -1 Then Exit Sub
        .MousePointer = flexDefault
        If mstrFontUnderLine <> "" Then
            arrTmp = Split(mstrFontUnderLine, "|")
            If UBound(arrTmp) >= 3 Then
                lngColor = Val(arrTmp(3))
            Else
                lngColor = vbBlack
            End If
            .Cell(flexcpForeColor, arrTmp(0), arrTmp(1), arrTmp(0), arrTmp(2)) = lngColor
            .Cell(flexcpFontUnderline, arrTmp(0), arrTmp(1), arrTmp(0), arrTmp(2)) = False
        End If
        
        If lngRow >= mcolCate("_" & CATE_诊断).lngBeginRow And lngRow <= mcolCate("_" & CATE_诊断).lngEndRow And lngCol > 0 Then
            If lngCol Mod mlngSubCol >= 1 And lngCol Mod mlngSubCol <= mlngSubCol - 1 Then
                If .Cell(flexcpText, lngRow, lngCol) <> "" And Right(.Cell(flexcpText, lngRow, lngCol), 3) = "..." Then
                    strInfo = .Cell(flexcpData, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_名称 + 1)
                    ShowTipInfo .hwnd, strInfo
                    Exit Sub
                End If
            End If
            
        ElseIf lngRow >= mcolCate("_" & CATE_用药跟踪).lngBeginRow And lngRow <= mcolCate("_" & CATE_用药跟踪).lngEndRow Then
            If lngCol Mod mlngSubCol = (SUBCOL_名称 + 1) Then
                If .Cell(flexcpData, lngRow, lngCol) <> "" Then
                    strInfo = .Cell(flexcpData, lngRow, lngCol)
                    ShowTipInfo .hwnd, strInfo
                    Exit Sub
                End If
            End If
        ElseIf lngRow >= mcolCate("_" & CATE_检查).lngBeginRow And lngRow <= mcolCate("_" & CATE_检查).lngEndRow Then
            If lngCol Mod mlngSubCol = SUBCOL_名称 + 1 Then
                If .Cell(flexcpText, lngRow, lngCol) <> "" Then
                    strInfo = .Cell(flexcpData, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_图标 + 1)
                    ShowTipInfo .hwnd, strInfo
                End If
                
                If Not .Cell(flexcpPicture, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_标注 + 1) Is Nothing Then
                    .Cell(flexcpFontUnderline, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_图标 + 1, lngRow, lngCol) = True
                    mstrFontUnderLine = lngRow & "|" & (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_图标 + 1 & "|" & lngCol & "|" & COLOR_鲜蓝
                    .MousePointer = flexCustom
                End If
                Exit Sub
            ElseIf lngCol Mod mlngSubCol = (SUBCOL_疗程 + 1) And Replace(.TextMatrix(lngRow, lngCol), vbTab, "") = "观片" Then
                '每次诊疗 疗程一列
                .Cell(flexcpFontUnderline, lngRow, lngCol) = True
                mstrFontUnderLine = lngRow & "|" & lngCol & "|" & lngCol & "|" & COLOR_鲜蓝
                .MousePointer = flexCustom
            Else
                .MousePointer = flexDefault
            End If
        ElseIf lngRow >= mcolCate("_" & CATE_检验).lngBeginRow And lngRow <= mcolCate("_" & CATE_检验).lngEndRow Then
            If lngCol Mod mlngSubCol = (SUBCOL_名称 + 1) Then
                If .Cell(flexcpText, lngRow, lngCol) <> "" Then
                    strInfo = .Cell(flexcpData, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_图标 + 1)
                    ShowTipInfo .hwnd, strInfo
                End If
                    
                If Not .Cell(flexcpPicture, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_标注 + 1) Is Nothing Then
                    .Cell(flexcpFontUnderline, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_图标 + 1, lngRow, lngCol) = True
                    mstrFontUnderLine = lngRow & "|" & (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_图标 + 1 & "|" & lngCol & "|" & COLOR_鲜蓝
                    .MousePointer = flexCustom
                End If
            End If
            Exit Sub
        ElseIf lngRow >= mcolCate("_" & CATE_病历文书).lngBeginRow And lngRow <= mcolCate("_" & CATE_病历文书).lngEndRow Then
            If lngCol Mod mlngSubCol = SUBCOL_名称 + 1 And .TextMatrix(lngRow, lngCol) <> "" Then
                strInfo = .Cell(flexcpData, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_图标 + 1)
                ShowTipInfo .hwnd, strInfo

                .Cell(flexcpFontUnderline, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_图标 + 1, lngRow, lngCol) = True
                mstrFontUnderLine = lngRow & "|" & (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_图标 + 1 & "|" & lngCol
                .Cell(flexcpForeColor, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_图标 + 1, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_名称 + 1) = COLOR_深橙
                .MousePointer = flexCustom
            End If
            Exit Sub
        ElseIf lngRow >= mcolCate("_" & CATE_其他医嘱).lngBeginRow And lngRow <= mcolCate("_" & CATE_其他医嘱).lngEndRow Then
            If .Cell(flexcpText, lngRow, lngCol) <> "" Then
                strInfo = .Cell(flexcpData, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_图标 + 1)
                Call ShowTipInfo(.hwnd, strInfo)
            End If
            Exit Sub
        End If
        
        If strInfo = "" Then
            ShowTipInfo 0, strInfo
        End If
    
        If mintIndex >= 0 Then
            SetImageDefault
        End If
    End With
End Sub

Private Function InitRS() As ADODB.Recordset
'功能:构造记录集
    Dim rs As ADODB.Recordset
    Dim strFields As String
    Dim strFieldName As String
    Dim lngLen As Long
    Dim FieldType As DataTypeEnum
    Dim i As Long, j As Long
    
    Dim arrField As Variant
    Dim arrSubFeld As Variant '字段名称|字段类型|字段长度 缺省字段类型 为adVarChar
    

    strFields = "ID|adVarChar|32,挂号ID|adBigInt|18,病历名称||100"

    
    Set rs = New ADODB.Recordset
    '-----------------------------------------
    With rs.Fields
        arrField = Split(strFields, ",")
        For i = LBound(arrField) To UBound(arrField)
            arrSubFeld = Split(arrField(i), "|")
            strFieldName = arrSubFeld(0)
            Select Case UCase(arrSubFeld(1) & "")
            Case UCase("adVarChar")
                FieldType = adVarChar
            Case UCase("adBigInt")
                FieldType = adBigInt
            Case Else
                FieldType = adVarChar
            End Select
            lngLen = Val(arrSubFeld(2))
            .Append strFieldName, FieldType, lngLen
        Next
    End With
    '---------------------------------------
    rs.CursorLocation = adUseClient
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenStatic
    rs.Open
    '----------------------------------
    Set InitRS = rs
End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
    If mbytFontSize = IIf(bytSize = 0, 9, 12) Then Exit Sub
    mbytFontSize = IIf(bytSize = 0, 9, 12)
    lblW.FontSize = mbytFontSize
    Call ResizeVsView
    Call LoadView     '重新加载数据
End Sub

Private Function GetSubString(ByVal strSource As String, ByVal lngShowW As Long) As String
'----------------------------------------------------------------------
'功能：字符串超过显示宽度时截取部分显示
'参数:strSource-需要截取的长度
'    lngShowW-显示宽度
'返回:截取的字符串 格式：截取字符串 + “...”
'----------------------------------------------------------------------
    Dim strRet As String
    Dim lngSingleWord As Long, lngSumWord As Long
    Dim lngSumLen As Long, i As Long
    
    Dim lngPosBegin As Long, lngPosEnd As Long, lngPosMid As Long
    Dim blnTag As Boolean
    
    If strSource = "" Then Exit Function
    
    lblW.AutoSize = True
    lblW.FontSize = mbytFontSize
    lblW.Caption = strSource
    lngSumWord = lblW.Width - 15     '实际字符宽度
    '字体宽度大于显示时截取
    If lngSumWord > lngShowW Then
        lblW.Caption = "\"
        lngSingleWord = lblW.Width - 15         '单个字符宽度
        lngShowW = lngShowW - lngSingleWord * 3      '预留省略号"..."长度
        lngSumLen = Len(strSource)              '总字符字符个数

        lngPosBegin = 1: lngPosEnd = lngSumLen
        
        For i = 1 To lngSumLen
            lngPosMid = (lngPosBegin + lngPosEnd) \ 2
            lblW.Caption = Mid(strSource, 1, lngPosMid)
            
            If lblW.Width < lngShowW Then
                lngPosBegin = lngPosMid
                blnTag = True
            ElseIf lblW.Width > lngShowW Then
                lngPosEnd = lngPosMid
                blnTag = False
            End If
            
            If (lngPosBegin + lngPosEnd) \ 2 = lngPosMid Then
                lngPosMid = IIf(blnTag, lngPosMid, lngPosMid - 1)
                strRet = Mid(strSource, 1, lngPosMid) & "..."
                Exit For
            End If
        Next
        
    Else
        strRet = strSource
    End If
    GetSubString = strRet
End Function

Private Function CheckRegister() As Boolean
'功能:检查病人就诊记录中是否存在其他科室的就诊记录
'返回:T-存在;F-不存在
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    strSql = " Select 1  " & vbNewLine & _
        "       From 病人挂号记录 A " & vbNewLine & _
        "       Where a.病人id = [1] And a.记录性质 = 1 And a.记录状态 = 1 And a.执行部门ID<>[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, mlng科室ID)
    
    CheckRegister = rsTmp.RecordCount > 0
End Function

Private Sub SetImageDefault()
'功能:设置缺省图片
    If imgBtn(mintIndex).Enabled = False Then Exit Sub
    Select Case mintIndex
    Case CMD_PREV
        Set imgBtn(mintIndex).Picture = imgFlag.ListImages("上次可用").Picture
    Case CMD_NEXT
        Set imgBtn(mintIndex).Picture = imgFlag.ListImages("下次可用").Picture
    Case CMD_ORTHER
        Set imgBtn(mintIndex).Picture = imgFlag.ListImages(IIf(mbytShow = 0, "只显示本科", "显示所有就诊")).Picture
    End Select
    mintIndex = -1
End Sub
