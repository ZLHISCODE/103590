VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDiseaseReportPrint 
   Caption         =   "阳性报告打印"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13545
   Icon            =   "frmDiseaseReportPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   13545
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   12600
      ScaleHeight     =   1170
      ScaleWidth      =   780
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   810
      Begin VB.Label lblSelect 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓  名↓"
         Height          =   180
         Index           =   3
         Left            =   30
         TabIndex        =   13
         Top             =   900
         Width           =   720
      End
      Begin VB.Label lblSelect 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "条  码↓"
         Height          =   180
         Index           =   0
         Left            =   30
         TabIndex        =   11
         Top             =   30
         Width           =   720
      End
      Begin VB.Label lblSelect 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号↓"
         Height          =   180
         Index           =   1
         Left            =   30
         TabIndex        =   10
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblSelect 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号↓"
         Height          =   180
         Index           =   2
         Left            =   30
         TabIndex        =   9
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5745
      Left            =   390
      ScaleHeight     =   5745
      ScaleWidth      =   11205
      TabIndex        =   1
      Top             =   1500
      Width           =   11205
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   4635
         Left            =   480
         TabIndex        =   2
         Top             =   270
         Width           =   9375
         _cx             =   16536
         _cy             =   8176
         Appearance      =   1
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
         Editable        =   0
         ShowComboButton =   1
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
   End
   Begin VB.PictureBox picTop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   360
      ScaleHeight     =   765
      ScaleWidth      =   12945
      TabIndex        =   0
      Top             =   570
      Width           =   12975
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   930
         TabIndex        =   17
         Top             =   60
         Width           =   2325
      End
      Begin VB.CheckBox chekPrint 
         BackColor       =   &H80000005&
         Caption         =   "显示已打印"
         Height          =   225
         Left            =   10830
         TabIndex        =   14
         Top             =   105
         Width           =   1215
      End
      Begin VB.TextBox txtFind 
         Height          =   270
         Left            =   8220
         TabIndex        =   7
         Top             =   75
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPStart 
         Height          =   285
         Left            =   4170
         TabIndex        =   4
         Top             =   75
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Format          =   222953473
         CurrentDate     =   43161
      End
      Begin MSComCtl2.DTPicker DTPEnd 
         Height          =   285
         Left            =   5880
         TabIndex        =   5
         Top             =   75
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Format          =   222953473
         CurrentDate     =   43161
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前科室"
         Height          =   180
         Index           =   3
         Left            =   60
         TabIndex        =   16
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "你知道吗,点击带“↓”的标签可以切换过滤条件。当条码输入框中有数据时时间过滤条件是无效的。黄色警示图标表示超出打印次数"
         ForeColor       =   &H0000C000&
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   10530
      End
      Begin VB.Label lblSel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "条  码↓"
         Height          =   180
         Left            =   7440
         TabIndex        =   12
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         Height          =   180
         Index           =   1
         Left            =   5580
         TabIndex        =   6
         Top             =   120
         Width           =   180
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "报告时间"
         Height          =   180
         Index           =   0
         Left            =   3360
         TabIndex        =   3
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imgVsf 
      Left            =   11970
      Top             =   4770
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportPrint.frx":6852
            Key             =   "医生打印"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportPrint.frx":D0B4
            Key             =   "自助打印"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportPrint.frx":13916
            Key             =   "禁止打印"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   60
      Top             =   90
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDiseaseReportPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrDBUser As String        '用户ID
Private mlngDeptID As Long          '当前选择科室ID
Private mintDeptType As Integer     '科室服务对象
Private mrsDept As ADODB.Recordset  '人员所在科室
'Private mstrDeptIDS As String      '使用“所有科室”选项无法获取报告打印科室的服务对象，导致不能限制打印次数

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/3/2
'功    能:显示窗体
'入    参:
'           objFrm          调用窗体
'           lngDeptID       科室ID
'           intDeptType     科室类型 0=其他,1=门诊,2=住院
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Sub ShowMe(ByVal strDBUser As String)
    mstrDBUser = strDBUser
    Me.Show
End Sub

Private Sub lblDrowBorder(objLbl As Label, objPic As PictureBox)
    '画Labled的边框线,当鼠标移动到Lable上时,呈现3D效果
    
    objPic.Line (objLbl.Left - 2, objLbl.Top - 2)-(objLbl.Left + objLbl.Width - 2, objLbl.Top - 2), &H8000000F '上边线
    objPic.Line (objLbl.Left + objLbl.Width, objLbl.Top)-(objLbl.Left + objLbl.Width, objLbl.Top + objLbl.Height), vbBlack '右边线
    objPic.Line (objLbl.Left + objLbl.Width, objLbl.Top + objLbl.Height)-(objLbl.Left, objLbl.Top + objLbl.Height), vbBlack '下边线
    objPic.Line (objLbl.Left - 2, objLbl.Top + objLbl.Height)-(objLbl.Left - 2, objLbl.Top - 2), &H8000000F '左边线

End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/3/9
'功    能:改变选项时获取部门ID和部门服务对象
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub cboDept_Click()
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim blnMZ As Boolean
          Dim blnZY As Boolean
          
1         On Error GoTo cboDept_Click_Error

2         If Me.Visible = False Then Exit Sub
3         With Me.cboDept
4             mlngDeptID = Val(.ItemData(.ListIndex))
5             If mlngDeptID = 0 Then Exit Sub
              
6             strSQL = "select distinct 服务对象 from 部门性质说明 where 部门ID=[1] and (工作性质 = '临床' Or 工作性质 = '治疗' Or 工作性质 = '检验')"
7             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "部门服务对象", mlngDeptID)
8             Do While Not rsTmp.EOF
9                 If Val(rsTmp("服务对象") & "") = 1 Then
10                    blnMZ = True
11                ElseIf Val(rsTmp("服务对象") & "") = 2 Then
12                    blnZY = True
13                ElseIf Val(rsTmp("服务对象") & "") = 3 Then
14                    blnMZ = True
15                    blnZY = True
16                End If
17                rsTmp.MoveNext
18            Loop
19        End With
20        If blnMZ = True And blnZY = False Then
21            mintDeptType = 1
22        ElseIf blnMZ = False And blnZY = True Then
23            mintDeptType = 2
24        ElseIf blnMZ = True And blnZY = True Then
25            mintDeptType = 3
26        End If


27        Exit Sub
cboDept_Click_Error:
28        Call WriteErrLog("zlPublicHisCommLis", "frmDiseaseReportPrint", "执行(cboDept_Click)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
29        Err.Clear
          
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    Dim strFind As String
    
    If KeyAscii <> 13 Then Exit Sub
    With Me.cboDept
        strFind = Trim(.Text)
        If strFind = "" Then
            mrsDept.Filter = ""
        Else
            mrsDept.Filter = " 编码 like '%" & strFind & "%' or 名称 like '%" & strFind & "%' or 简码 like '%" & strFind & "%'"
        End If
    End With
    Call setDataToCbo(mrsDept)
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Browse_Find            '查找
            Call FindData
        Case ConMenu_Browse_SelAll          '全选
            Call SelOrDelAll(1)
        Case ConMenu_Browse_ClsAll          '全清
            Call SelOrDelAll(0)
        Case ConMenu_Browse_Print           '打印
            Call BatchPrintReport(2)
        Case ConMenu_Browse_PrintSet        '打印设置
           Call BatchPrintReport(3)
        Case ConMenu_Browse_PrintView       '预览
            Call BatchPrintReport(1)
        Case ConMenu_Browse_Exit            '退出
            Unload Me
    End Select
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/3/5
'功    能:打印报告
'入    参:
'           1=预览,2=打印,3=打印设置
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub BatchPrintReport(ByVal byRunMode As Byte)
          Dim lngRow As Long
          Dim lngSampleID As Long
          Dim strPrintCount As String
          Dim lngPrintCount As Long
          
1         On Error GoTo BatchPrintReport_Error

2         strPrintCount = ComGetPara(Sel_Lis_DB, "医生站传染病报告打印次数", gSysInfo.SysNo, gSysInfo.ModlNo)
3         Select Case mintDeptType
              Case 1  '门诊
4                 lngPrintCount = Val(Split(strPrintCount, "|")(0))
5             Case 2  '住院
6                 lngPrintCount = Val(Split(strPrintCount, "|")(1))
7             Case 3  '门诊和住院，以小的为准
8                 If Val(Split(strPrintCount, "|")(0)) > Val(Split(strPrintCount, "|")(1)) Then
9                     lngPrintCount = Val(Split(strPrintCount, "|")(1))
10                Else
11                    lngPrintCount = Val(Split(strPrintCount, "|")(0))
12                End If
13            Case 0  '其他
14                lngPrintCount = Val(Split(strPrintCount, "|")(2))
15        End Select
              
16        With Me.VSFList
17            For lngRow = 1 To .Rows - 1
18                If .Cell(flexcpChecked, lngRow, .ColIndex("选择")) = 1 Then
19                    lngSampleID = Val(.TextMatrix(lngRow, .ColIndex("ID")))
20                    Call PrintReport(lngSampleID, byRunMode, lngRow, lngPrintCount)
21                    If byRunMode = 3 Then Exit Sub
22                End If
23            Next
24        End With


25        Exit Sub
BatchPrintReport_Error:
26        Call WriteErrLog("zlPublicHisCommLis", "frmDiseaseReportPrint", "执行(BatchPrintReport)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
27        Err.Clear
End Sub

Private Function PrintReport(lngSampleID As Long, Optional byRunMode As Byte = 2, Optional lngRow As Long, Optional lngPrintCount As Long) As Boolean
          '功能       打印报告
          Dim intCount As Integer
          Dim strNO As String
          Dim intSel As Integer
          Dim strChart(0 To 8) As String
          Dim strSQL As String
          Dim strTmp As String
          Dim rsTmp As ADODB.Recordset
          Dim rsReportFormat As ADODB.Recordset


1         On Error GoTo PrintReport_Error

2         strSQL = "select b.id 仪器id ,b.名称 仪器名称,b.仪器类别,Nvl(a.病人来源,1) 病人来源,a.报告时间,a.阳性报告,a.标本序号,a.医生站打印 from 检验报告记录 a,检验仪器记录 b where a.仪器id = b.id and a.id = [1]"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "报告打印", lngSampleID)

4         If rsTmp.RecordCount = 0 Then Exit Function

          '对比打印次数和参数
5         If lngPrintCount > 0 Then
6             If Val(rsTmp("医生站打印") & "") > lngPrintCount Then
7                 With Me.VSFList
8                     .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
9                     .Cell(flexcpPicture, lngRow, .ColIndex("打印方式")) = imgVsf.ListImages("禁止打印").ExtractIcon
10                End With
11                PrintReport = False
12                Exit Function
13            End If
14        End If

15        strSQL = "select id,编码,名称,门诊单据,住院单据,体检单据,院外单据,门诊格式,住院格式,体检格式,院外格式,格式数量," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(门诊单据, '00000')) || '-2' 门诊单据号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(住院单据, '00000')) || '-2' 住院单据号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(体检单据, '00000')) || '-2' 体检单据号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(院外单据, '00000')) || '-2' 院外单据号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(门诊格式, '00000')) || '-2' 门诊格式号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(住院格式, '00000')) || '-2' 住院格式号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(体检格式, '00000')) || '-2' 体检格式号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(院外格式, '00000')) || '-2' 院外格式号" & vbNewLine & _
                      "from 检验仪器记录 where id = [1] "

16        Set rsReportFormat = ComOpenSQL(Sel_Lis_DB, strSQL, "检验技师站", Val(rsTmp("仪器ID") & ""))


17        rsReportFormat.Filter = "id=" & Val(rsTmp("仪器ID") & "")
18        If Val(rsTmp("仪器类别")) = 1 Then
19            If Val(rsTmp("阳性报告") & "") = 1 Then
                  '阳性
20                intSel = 0
21            Else
                  '阴性
22                intSel = 1
23            End If
24        Else
25            intCount = GetSampleValCount(lngSampleID)
              '没有结果时提示
26            If intCount = 0 Then
27                Exit Function
28            End If
29            If rsReportFormat.RecordCount > 0 Then
30                If Val(rsReportFormat("格式数量") & "") > 0 Then
31                    If intCount > Val(rsReportFormat("格式数量") & "") Then
32                        intSel = 0
33                    Else
34                        intSel = 1
35                    End If
36                End If
37            Else
38                intSel = 0
39            End If

40        End If
41        Select Case Val(rsTmp("病人来源"))
              Case 1
42                If intSel = 0 Then
43                    strNO = rsReportFormat("门诊单据号")
44                Else
45                    strNO = rsReportFormat("门诊格式号")
46                End If
47            Case 2
48                If intSel = 0 Then
49                    strNO = rsReportFormat("住院单据号")
50                Else
51                    strNO = rsReportFormat("住院格式号")
52                End If
53            Case 3
54                If intSel = 0 Then
55                    strNO = rsReportFormat("住院单据号")
56                Else
57                    strNO = rsReportFormat("住院格式号")
58                End If
59            Case 4
60                If intSel = 0 Then
61                    strNO = rsReportFormat("院外单据号")
62                Else
63                    strNO = rsReportFormat("院外格式号")
64                End If
65            Case Else
66                If intSel = 0 Then
67                    strNO = rsReportFormat("门诊单据号")
68                Else
69                    strNO = rsReportFormat("门诊格式号")
70                End If
71        End Select
72        If byRunMode = 3 Then
73            If strNO <> "" Then
74                FunReportPrintSet gcnLisOracle, gSysInfo.SysNo, strNO, Me
75            End If
76        Else
             '读图像
77            strTmp = "开始读入图像:" & Now & vbCrLf
78            If ReadSampleImage(lngSampleID, strChart, "", 25) = False Then
79                Exit Function
80            End If
81            strTmp = strTmp & "读入图像完成:" & Now & vbCrLf

82            FunReportOpen gcnLisOracle, gSysInfo.SysNo, strNO, Me, "标本ID=" & lngSampleID, "图形1=" & strChart(0), "图形2=" & strChart(1), "图形3=" & strChart(2), _
                      "图形4=" & strChart(3), "图形5=" & strChart(4), "图形6=" & strChart(5), "图形7=" & strChart(6), "图形8=" & strChart(7), _
                      "图形9=" & strChart(8), byRunMode
83            strTmp = strTmp & "打印完成:" & Now & vbCrLf

              '对于审核过的标本标识
84            strSQL = "Zl_检验报告打印_Edit(1," & lngSampleID & ",1)"
85            Call ComExecuteProc(Sel_Lis_DB, strSQL, "打印标本")
86            strTmp = strTmp & "完成打印:" & Now

87            SaveDBLog 18, 6, lngSampleID, "打印", "报告打印", 2500, "临床实验室管理"
88        End If

89        PrintReport = True

          '发送刷新科内概况已打印标签申请
90        Call SendMessage("RefreshDeptSurvey7")


91        Exit Function
PrintReport_Error:
92        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(PrintReport)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
93        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/3/5
'功    能:全选/全清
'入    参:
'           intType     0=全清,1=权限
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub SelOrDelAll(ByVal intType As Integer)
          Dim lngRow As Long
1         On Error GoTo SelOrDelAll_Error

2         With Me.VSFList
3             For lngRow = 1 To .Rows - 1
4                 .Cell(flexcpChecked, lngRow, .ColIndex("选择")) = intType
5             Next
6         End With


7         Exit Sub
SelOrDelAll_Error:
8         Call WriteErrLog("zlPublicHisCommLis", "frmDiseaseReportPrint", "执行(SelOrDelAll)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
9         Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/3/2
'功    能:查找数据
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub FindData()
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim lngRow As Long
          Dim lngDeptID As Long
          Dim strErr As String
          
1         On Error GoTo FindData_Error
          
          '获取科室ID
2         With Me.cboDept
3             If .Text <> "所有科室" Then
4                 lngDeptID = Val(.ItemData(.ListIndex))
5                 If lngDeptID = 0 Then Exit Sub
6             End If
7         End With
      '    If mstrDeptIDS = "" Then Exit Sub

              
          '部门ID
      '    With Me.cboDept
      '        If .Text <> "所有科室" Then
      '            strSQL = "Select Distinct '0' 选择,decode(Sign(Nvl(a.医生站打印, 0)),1,'医生打印',decode(Sign(Nvl(a.自助打印次数, 0)),1,'自助打印','')) 打印方式,a.医生站打印, a.Id, a.标本序号 标本号, a.姓名, decode(a.性别,1,'男',2,'女','未知') 性别, a.年龄," & vbCrLf & _
      '                " a.标本类型, a.审核时间 报告时间, a.审核人 报告人,a.样本条码, a.申请项目" & vbCrLf & _
      '                " From 检验报告记录 A, 检验申请组合 B, 检验报卡打印科室 C" & vbCrLf & _
      '                " Where a.Id = b.标本id And b.组合id = c.组合id And a.是否传染病 = 1 And 审核人 Is Not Null And c.部门id = [1] "
      '        Else
8                 strSQL = "Select /*+cardinality(d,10)*/ Distinct '0' 选择,decode(Sign(Nvl(a.医生站打印, 0)),1,'医生打印',decode(Sign(Nvl(a.自助打印次数, 0)),1,'自助打印','')) 打印方式,a.医生站打印, a.Id, a.标本序号 标本号, a.姓名, decode(a.性别,1,'男',2,'女','未知') 性别, a.年龄," & vbCrLf & _
                      " a.标本类型, a.审核时间 报告时间, a.审核人 报告人,a.样本条码, a.申请项目" & vbCrLf & _
                      " From 检验报告记录 A, 检验申请组合 B, 检验报卡打印科室 C" & vbCrLf & _
                      " Where a.Id = b.标本id And b.组合id = c.组合id And a.是否传染病 = 1 And 审核人 Is Not Null And c.部门id in (Select Column_Value From Table(Cast(f_Str2list([1]) As zltools.t_strlist)) d)"
      '        End If
      '    End With
          
          '过滤条件
9         If Trim(Me.txtFind.Text) <> "" Then
10            Select Case Me.lblSel.Caption
                  Case "条  码↓"
11                    strSQL = strSQL & " and a.样本条码=[2]"
12                Case "门诊号↓"
13                    strSQL = strSQL & " and a.门诊号=[2]"
14                Case "住院号↓"
15                    strSQL = strSQL & " and a.住院号=[2]"
16                Case "姓  名↓"
17                    strSQL = strSQL & " and a.姓名 like [2]"
18            End Select
19        End If
          
          '报告时间
20        If Trim(Me.txtFind.Text) = "" Then
21            strSQL = strSQL & " and a.审核时间 between [3] and [4]"
              '高峰时段限制查询
22            If Not funCheckRushHours(2500, 2001, "浏览检验结果", CDate(Format(Me.DTPStart.value, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(Me.DTPEnd.value, "yyyy-mm-dd") & " 23:59:59")) Then Exit Sub
23        End If
          
          '是否显示已打印
24        If Me.chekPrint.value <> 1 Then
25            strSQL = strSQL & " and nvl(a.医生站打印,0)=0 and nvl(a.自助打印次数,0)=0 "
26        End If
          
27        strSQL = strSQL & " order by a.审核时间"
      '    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验报告记录", IIf(cboDept.Text <> "所有科室", lngDeptID, mstrDeptIDS), Trim(Me.txtFind.Text), CDate(Format(Me.DTPStart.value, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(Me.DTPEnd.value, "yyyy-mm-dd") & " 23:59:59"))
28        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验报告记录", lngDeptID, IIf(Me.lblSel.Caption = "姓  名↓", "%" & Trim(Me.txtFind.Text) & "%", Trim(Me.txtFind.Text)), CDate(Format(Me.DTPStart.value, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(Me.DTPEnd.value, "yyyy-mm-dd") & " 23:59:59"))
29        If vfgLoadFromRecord(Me.VSFList, rsTmp, strErr) = False Then
30            MsgBox strErr
31            Exit Sub
32        End If
33        With Me.VSFList
34            .ColDataType(.ColIndex("选择")) = flexDTBoolean
35            .ColWidth(.ColIndex("选择")) = 500
36            .ColWidth(.ColIndex("打印方式")) = 250
37            .ColWidth(.ColIndex("标本号")) = 800
38            .ColWidth(.ColIndex("姓名")) = 1000
39            .ColWidth(.ColIndex("性别")) = 500
40            .ColWidth(.ColIndex("年龄")) = 800
41            .ColWidth(.ColIndex("标本类型")) = 1000
42            .ColWidth(.ColIndex("报告时间")) = 2000
43            .ColWidth(.ColIndex("报告人")) = 1000
44            .ColWidth(.ColIndex("样本条码")) = 1500
45            .ColWidth(.ColIndex("申请项目")) = 1000
46            .ExtendLastCol = True
              
47            .Cell(flexcpPicture, 0, .ColIndex("打印方式")) = imgVsf.ListImages("医生打印").ExtractIcon
              
48            For lngRow = 1 To .Rows - 1
49                If Trim(.TextMatrix(lngRow, .ColIndex("打印方式"))) = "医生打印" Then
50                    .Cell(flexcpPicture, lngRow, .ColIndex("打印方式")) = imgVsf.ListImages("医生打印").ExtractIcon
51                End If
                  
52                If Trim(.TextMatrix(lngRow, .ColIndex("打印方式"))) = "自助打印" Then
53                    .Cell(flexcpPicture, lngRow, .ColIndex("打印方式")) = imgVsf.ListImages("自助打印").ExtractIcon
54                End If
55            Next
56        End With


57        Exit Sub
FindData_Error:
58        Call WriteErrLog("zlPublicHisCommLis", "frmDiseaseReportPrint", "执行(FindData)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
59        Err.Clear
End Sub

Private Sub cbrMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    On Error Resume Next
    With Me.picTop
        .Left = Left
        .Top = Top
        .Width = Right - Left
    End With
    With Me.picMain
        .Left = Left
        .Top = Me.picTop.Top + Me.picTop.Height
        .Width = Me.picTop.Width
        .Height = Bottom - .Top
    End With
    With picSel
        .Left = Me.picTop.Left + lblSel.Left - 10
        .Top = Me.picTop.Top + lblSel.Top + lblSel.Height
    End With
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/3/5
'功    能:快捷键，如果使用快捷键无效请检查快捷键是否被其他程序占用
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 1      '全选
            Call SelOrDelAll(1)
        Case 4      '全清
            Call SelOrDelAll(0)
        Case 16     '打印
            Call BatchPrintReport(2)
        Case 21     '预览
            Call BatchPrintReport(1)
        Case 17     '退出
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Me.cbrMain.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrMain.EnableCustomization False

    '-----------------------------------------------------
    '菜单定义
    Me.cbrMain.ActiveMenuBar.Title = "菜单"
    Me.cbrMain.ActiveMenuBar.Visible = False
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Find, "查找(F5)")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_SelAll, "全选(Crl+A)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_ClsAll, "全清(Crl+D)")
        Set cbrControl = .Add(xtpControlSplitButtonPopup, ConMenu_Browse_Print, "打印(Crl+P)"): cbrControl.BeginGroup = True
        cbrControl.Style = xtpButtonIconAndCaption
        With cbrControl.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintSet, "打印设置  ")
        End With
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintView, "预览(Crl+U)")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Exit, "退出(Crl+Q)"): cbrControl.BeginGroup = True
    End With

    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type = xtpControlButton Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    '快键绑定
    With Me.cbrMain.KeyBindings
        .Add 0, VK_F5, ConMenu_Browse_Find
    End With
    
    Call intData
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/3/5
'功    能:初始化数据
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub intData()
          Dim strTitle As String
          
          '初始化VSF
1         On Error GoTo intData_Error

2         strTitle = "选择,500,1;打印方式,250,1;标本号,800,1;姓名,1000,1;性别,500,1;年龄,800,1;标本类型,1000,1;" & _
                      "报告时间,2000,1;报告人,1000,1;样本条码,1500,1;申请项目,1000,1"
3         Call vfgSetting(0, Me.VSFList, strTitle)
4         With Me.VSFList
5             .ExtendLastCol = True
6              .Cell(flexcpPicture, 0, .ColIndex("打印方式")) = imgVsf.ListImages("医生打印").ExtractIcon
7         End With
          
          '获取服务器时间
8         Me.DTPStart.value = Currentdate
9         Me.DTPEnd.value = Me.DTPStart.value
          
          '获取用户当前科室
10        Call getUserDept


11        Exit Sub
intData_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "frmDiseaseReportPrint", "执行(intData)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
13        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/3/9
'功    能:读取人员科室
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub getUserDept()
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
1         On Error GoTo getUserDept_Error

2         strSQL = "Select a.Id, a.编码, a.名称, a.简码" & vbCrLf & _
                  " From 部门表 A, 上机人员表 B, 部门人员 C" & vbCrLf & _
                  " Where a.Id = c.部门id And b.人员id = c.人员id And b.用户名 = [1] And a.撤档时间 > Sysdate "
3         If gUserInfo.NodeNo <> "-" Then
4             strSQL = strSQL & " And (a.站点 = 1 Or a.站点 Is Null)"
5         End If
6         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "人员部门", mstrDBUser)
7         Set mrsDept = rsTmp
8         Call setDataToCbo(rsTmp)

9         Exit Sub
getUserDept_Error:
10        Call WriteErrLog("zlPublicHisCommLis", "frmDiseaseReportPrint", "执行(getUserDept)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
11        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/3/9
'功    能:将数据绑定到下拉列表中
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub setDataToCbo(ByVal rsTmp As ADODB.Recordset)
1         On Error GoTo setDataToCbo_Error

2         With Me.cboDept
3             .Clear
      '        .AddItem "所有科室"
4             Do While Not rsTmp.EOF
5                 .AddItem "[" & rsTmp("编码") & "]" & rsTmp("名称")
6                 .ItemData(.ListCount - 1) = Val(rsTmp("ID") & "")
      '            If Val(rsTmp("ID") & "") <> 0 Then mstrDeptIDS = mstrDeptIDS & "," & rsTmp("ID")
7                 rsTmp.MoveNext
8             Loop
      '        If mstrDeptIDS <> "" Then mstrDeptIDS = Mid(mstrDeptIDS, 2)
9             If .ListCount > 0 Then .ListIndex = 0
10        End With


11        Exit Sub
setDataToCbo_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "frmDiseaseReportPrint", "执行(setDataToCbo)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
13        Err.Clear

End Sub

Private Sub Form_Unload(Cancel As Integer)
'    mstrDeptIDS = ""
    mstrDBUser = ""
    mlngDeptID = 0
    Set mrsDept = Nothing
End Sub

Private Sub lblSel_Click()
    If Me.picSel.Visible = False Then
        Me.picSel.Visible = True
    Else
        Me.picSel.Visible = False
    End If
End Sub

Private Sub lblSel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call lblDrowBorder(lblSel, picTop)
End Sub

Private Sub lblSelect_Click(Index As Integer)
    Me.lblSel.Caption = Me.lblSelect(Index).Caption
    Me.picSel.Visible = False
End Sub

Private Sub lblSelect_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call lblDrowBorder(lblSelect(Index), picSel)
End Sub

Private Sub PicMain_Resize()
    On Error Resume Next
    With Me.VSFList
        .Left = 0
        .Top = 0
        .Width = Me.picMain.Width
        .Height = Me.picMain.Height
    End With
End Sub

Private Sub picSel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSel.Cls
End Sub

Private Sub picTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picTop.Cls
End Sub

Private Sub txtFind_GotFocus()
    With Me.txtFind
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call FindData
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/3/5
'功    能:选择/取消选择
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub VSFList_Click()
    Dim lngRow As Long
    Dim lngCol As Long
    
    On Error Resume Next
    
    With Me.VSFList
        lngRow = .MouseRow
        lngCol = .MouseCol
        
        If lngRow <= 0 Or lngCol <> .ColIndex("选择") Then Exit Sub
        If .Cell(flexcpChecked, lngRow, lngCol) = 1 Then
            .Cell(flexcpChecked, lngRow, lngCol) = 0
        Else
            .Cell(flexcpChecked, lngRow, lngCol) = 1
        End If
    End With
End Sub
