VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.UserControl uclReport 
   ClientHeight    =   7050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9195
   ScaleHeight     =   7050
   ScaleWidth      =   9195
   Begin VB.PictureBox picComment 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6990
      Left            =   4590
      ScaleHeight     =   6990
      ScaleWidth      =   3495
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2730
      Width           =   3495
      Begin VB.TextBox txtResultComment 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   330
         Width           =   3255
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSupplement 
         Height          =   1095
         Left            =   270
         TabIndex        =   3
         Top             =   1890
         Width           =   3225
         _cx             =   5689
         _cy             =   1931
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
      Begin VB.Label lblResultComment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结果说明:"
         Height          =   180
         Left            =   60
         TabIndex        =   6
         Top             =   90
         Width           =   810
      End
      Begin VB.Label lblSupplement 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "补充报告:"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   1590
         Width           =   810
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3210
      ScaleHeight     =   945
      ScaleWidth      =   4125
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   4155
      Begin XtremeSuiteControls.TabControl tabPage 
         Height          =   735
         Left            =   180
         TabIndex        =   1
         Top             =   120
         Width           =   585
         _Version        =   589884
         _ExtentX        =   1032
         _ExtentY        =   1296
         _StockProps     =   64
      End
   End
   Begin XtremeDockingPane.DockingPane dkpPage 
      Left            =   810
      Top             =   810
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "uclReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobjFrm As Object
Private mobjSamplePreviousContrast As Object            '历次比对窗体
Private mobjSampleImgs As Object                        '标本图像查看窗体
Private mlngSampleID As Long                            '标本ID
Private mdteSampleTime As Date
Private mintVersion As Integer
Private mintSampleType As Integer                       '1=微生物报告

Private Sub dkpPage_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error Resume Next
    Select Case Item.ID
        Case 1
            If mobjFrm Is Nothing Then Set mobjFrm = New frmAdviceReprotBrowse
            Item.Handle = mobjFrm.hWnd
        Case 2
            Item.Handle = picPage.hWnd
    End Select
End Sub

Private Sub picComment_Resize()
    On Error Resume Next
    If mintSampleType = 1 Or mintVersion = 10 Then
        With lblResultComment
            .Top = 100
            .Left = 0
        End With
        With txtResultComment
            .Top = lblResultComment.Top + lblResultComment.Height
            .Left = 100
            .Width = picComment.Width - 200
            .Height = picComment.Height - .Top
        End With
        lblSupplement.Visible = False
        vsfSupplement.Visible = False
    Else
        With lblResultComment
            .Top = 100
            .Left = 0
        End With
        With txtResultComment
            .Top = lblResultComment.Top + lblResultComment.Height
            .Left = 100
            .Width = picComment.Width - 200
            .Height = (picComment.Height - 600) / 2
        End With

        With lblSupplement
            .Left = 0
            .Top = txtResultComment.Top + txtResultComment.Height + 300
            .Visible = True
        End With
        With vsfSupplement
            .Left = 100
            .Top = lblSupplement.Top + lblSupplement.Height
            .Width = picComment.Width - 200
            .Height = txtResultComment.Height
            .Visible = True
        End With
    End If

End Sub

Private Sub picPage_Resize()
    On Error Resume Next
    With tabPage
        .Left = 0
        .Top = 0
        .Width = picPage.Width
        .Height = picPage.Height
    End With
End Sub

Private Sub tabPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Select Case Item.Index
    Case 1                          '历次
        Call mobjSamplePreviousContrast.InitData(mlngSampleID, mdteSampleTime, mintVersion)
    Case 2                        '图形
        Call mobjSampleImgs.ReadImages(mlngSampleID, mintVersion)
    End Select
End Sub

Private Sub UserControl_Initialize()
    Dim objPanle As Pane
    Dim strDiagnoseID As String
    Dim strDiagnoseStr As String

    Set objPanle = dkpPage.CreatePane(1, 2, 1, DockLeftOf, Nothing)  '参数一次是：分页索引,宽占比，高占比，对应位置，参照对象
    objPanle.Options = PaneNoCaption '是否可以浮动
    Set objPanle = dkpPage.CreatePane(2, 1, 1, DockRightOf, dkpPage.Panes(1))
    objPanle.Options = PaneNoCaption '是否可以浮动
    
    dkpPage.Options.ThemedFloatingFrames = True
    dkpPage.Options.UseSplitterTracker = False '实时拖动
    dkpPage.Options.AlphaDockingContext = True
    dkpPage.Options.CloseGroupOnButtonClick = True
    dkpPage.Options.HideClient = True

End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-07-27
'功    能:  通过标本ID显示报告（报告查询模块调用）
'入    参:
'           objFrm          调用窗体
'           mblnDoctorShow  是否是医生站调用
'           lngPaintID      病人ID
'           lngSampleID     标本ID
'           intVersion      报告版本，25=新版LIS，10=老版LIS
'           intSampleType   是否是微生物报告，0=普通报告，1=微生物报告
'           intPositive     报告类型，1=药敏报告，3=PDF报告，其他=阴性报告
'           strDiagnosis    诊断
'           strResult       备注
'           intCount        老版LIS结果次数
'           dteSampleTime   标本核收时间
'           strPrivs        人员权限
'出    参:
'           strThirdReport  三方报告
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Public Function GetSampleReport(objFrm As Object, ByVal blnDoctorShow As Boolean, ByVal lngPaintID As Long, ByVal lngSampleID As Long, ByVal intVersion As Long, _
                                ByVal intSampleType As Integer, Optional ByVal intPositive As Integer, _
                                Optional ByVal strDiagnosis As String, Optional ByVal strResult As String, _
                                Optional ByVal intCount As Integer, Optional ByVal dteSampleTime As Date, _
                                Optional ByVal strPrivs As String, Optional ByRef strThirdReport As String) As Long
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim lngRow As Long
          Dim strSupplementID As String       '补充报告指标ID，多个使用逗号分割

1         On Error GoTo GetSampleReport_Error

2         mlngSampleID = lngSampleID
3         mdteSampleTime = dteSampleTime
4         mintVersion = intVersion
5         mintSampleType = intSampleType

6         tabPage.RemoveAll
7         Set mobjSamplePreviousContrast = Nothing
8         Set mobjSampleImgs = Nothing

9         With tabPage
10            .Icons = frmPubIcons.imgPublic.Icons
11            .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
12            .PaintManager.Color = xtpTabColorOffice2003
13            .PaintManager.OneNoteColors = False
14            .PaintManager.BoldSelected = True

              '备注
15            .InsertItem 0, "备注", picComment.hWnd, ConTab_Sample_Comment
16            .PaintManager.Position = xtpTabPositionTop
17            .PaintManager.Layout = xtpTabLayoutAutoSize
18            .PaintManager.ShowIcons = True

              '历次
19            Set mobjSamplePreviousContrast = New frmPubSamplePreviousContrast
20            .InsertItem 1, "历次", mobjSamplePreviousContrast.hWnd, ConTab_Sample_History
21            .PaintManager.Position = xtpTabPositionTop
22            .PaintManager.Layout = xtpTabLayoutAutoSize
23            .PaintManager.ShowIcons = True

              '图像
24            Set mobjSampleImgs = New frmPubSampleImgs
25            .InsertItem 2, "图像", mobjSampleImgs.hWnd, ConTab_Sample_Image
26            .PaintManager.Position = xtpTabPositionTop
27            .PaintManager.Layout = xtpTabLayoutAutoSize
28            .PaintManager.ShowIcons = True

29            .Item(0).Selected = True
30        End With

          '查新标本结果备注、补充报告等信息
31        If intVersion = 25 Then
              '结果说明
32            strSQL = "select 结果说明 from  检验报告记录 where ID=[1]"
33            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验报告记录", lngSampleID)
34            If Not rsTmp.EOF Then
35                txtResultComment.Text = rsTmp("结果说明") & ""
36            End If
              '补充报告
37            Set rsTmp = gobjSample.GetSupplementReportFromDB(lngSampleID)
38            Call gobjLiscomlib.SetDataToVSF(vsfSupplement, rsTmp)
39            Call gobjLiscomlib.GetVSFTitle(vsfSupplement, App.EXEName, UserControl.Name, True, rsTmp)      '读取列顺序
40            With vsfSupplement
41                .SelectionMode = flexSelectionFree
42                .ColHidden(.ColIndex("ID")) = True
43                .ColHidden(.ColIndex("补充报告ID")) = True
44                .ColHidden(.ColIndex("项目ID")) = True
45                .ColHidden(.ColIndex("仪器ID")) = True
46                .ColHidden(.ColIndex("结果标志")) = True
47                .ColHidden(.ColIndex("参考高值")) = True
48                .ColHidden(.ColIndex("参考低值")) = True
                  '设置颜色
49                For lngRow = 1 To .Rows - 1
50                    strSupplementID = strSupplementID & "," & .TextMatrix(lngRow, .ColIndex("ID"))

51                    If Val(.TextMatrix(lngRow, .ColIndex("结果标志"))) = 2 Then .TextMatrix(lngRow, .ColIndex("结果标志")) = "↓"
52                    If Val(.TextMatrix(lngRow, .ColIndex("结果标志"))) = 3 Then .TextMatrix(lngRow, .ColIndex("结果标志")) = "↑"
53                    If Val(.TextMatrix(lngRow, .ColIndex("结果标志"))) = 4 Then .TextMatrix(lngRow, .ColIndex("结果标志")) = "异常"
54                    If Val(.TextMatrix(lngRow, .ColIndex("结果标志"))) = 5 Then .TextMatrix(lngRow, .ColIndex("结果标志")) = "↓↓"
55                    If Val(.TextMatrix(lngRow, .ColIndex("结果标志"))) = 6 Then .TextMatrix(lngRow, .ColIndex("结果标志")) = "↑↑"
56                Next

57                If strSupplementID <> "" Then strSupplementID = Mid(strSupplementID, 2)
58            End With

59        End If

60        GetSampleReport = mobjFrm.ShowReportByID(objFrm, blnDoctorShow, lngPaintID, lngSampleID, intVersion, intSampleType, intPositive, strDiagnosis, strResult, intCount, strSupplementID, strPrivs, strThirdReport)


61        Exit Function
GetSampleReport_Error:
62        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "uclReport", "执行(GetSampleReport)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
63        Err.Clear
End Function

Public Sub UnloadCrl()
    If Not mobjFrm Is Nothing Then
        Unload mobjFrm
        Set mobjFrm = Nothing
    End If
    If Not mobjSamplePreviousContrast Is Nothing Then
        Unload mobjSamplePreviousContrast
        Set mobjSamplePreviousContrast = Nothing
    End If
    If Not mobjSampleImgs Is Nothing Then
        Unload mobjSampleImgs
        Set mobjSampleImgs = Nothing
    End If
End Sub

