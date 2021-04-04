VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCISAduitPDF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "输出档案到PDF"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12525
   Icon            =   "frmCISAduitPDF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   12525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdUnSelectAll 
      Cancel          =   -1  'True
      Caption         =   "全清(&U)"
      Height          =   350
      Left            =   7845
      TabIndex        =   14
      Top             =   5685
      Width           =   1200
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "全选(&S)"
      Height          =   350
      Left            =   6660
      TabIndex        =   15
      Top             =   5685
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Caption         =   "待输出病人清单(由主界面条件过滤)"
      Height          =   5415
      Left            =   90
      TabIndex        =   7
      Top             =   150
      Width           =   8955
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   5010
         Left            =   45
         TabIndex        =   8
         ToolTipText     =   "双击选中"
         Top             =   315
         Width           =   8820
         _cx             =   15557
         _cy             =   8837
         Appearance      =   2
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         Begin VB.PictureBox picInfo 
            BackColor       =   &H00FFEBD7&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   7470
            Picture         =   "frmCISAduitPDF.frx":000C
            ScaleHeight     =   225
            ScaleMode       =   0  'User
            ScaleWidth      =   283.333
            TabIndex        =   13
            Top             =   285
            Width           =   250
         End
         Begin MSComctlLib.ImageList img16 
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCISAduitPDF.frx":685E
                  Key             =   "Selected"
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Frame fraPageScope 
      Caption         =   "输出选项(&R)"
      Height          =   5415
      Left            =   9120
      TabIndex        =   2
      Top             =   150
      Width           =   3330
      Begin VB.CommandButton cmdPath 
         Caption         =   "…"
         Height          =   315
         Left            =   2910
         TabIndex        =   3
         Top             =   4898
         Width           =   210
      End
      Begin VB.ListBox lst 
         Height          =   4470
         Left            =   180
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   240
         Width           =   2985
      End
      Begin VB.Frame Frame1 
         Height          =   120
         Left            =   30
         TabIndex        =   4
         Top             =   4710
         Width           =   3270
      End
      Begin VB.TextBox txtPath 
         Height          =   300
         Left            =   960
         TabIndex        =   11
         Top             =   4905
         Width           =   1965
      End
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   4930
         Width           =   2205
      End
      Begin VB.Label Label1 
         Caption         =   "输出位置"
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   4995
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "输出设备"
         Height          =   255
         Left            =   150
         TabIndex        =   10
         Top             =   4995
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   11250
      TabIndex        =   1
      Top             =   5685
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   9120
      TabIndex        =   0
      Top             =   5685
      Width           =   1200
   End
   Begin VB.Label Label2 
      Height          =   270
      Left            =   90
      TabIndex        =   9
      Top             =   5730
      Width           =   6495
   End
End
Attribute VB_Name = "frmCISAduitPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnDoctorAdvice As Boolean
Private mblnPDF As Boolean
Private mstrPrintDocIDs As String '共享病历的子文档只打印一次
Private WithEvents mclsDockAduits As zlRichEPR.clsDockAduits
Attribute mclsDockAduits.VB_VarHelpID = -1
Private mfrmTipInfo As New frmTipInfo
Private Enum mCol
    选择
    病人ID
    主页ID
    姓名
    床号
    住院号
    性别
    年龄
    出院科室
    出院科室ID
    入院日期
    出院日期
    打印记录
End Enum

Public Sub ShowMe(ByVal frmObj As Object, ByVal parVsf As VSFlexGrid, ByVal intType As Integer, ByVal blnDoctorAdvice As Boolean, ByVal blnPDF As Boolean)
Dim strPath As String, strSelect As String, strTmp As String, intCount As Integer, arySerial As Variant, strPrinterName As String
'blnDoctorAdvice False=医嘱本 True=医嘱单
'----------------------------------------------------------------------------------------
'1-住院医嘱;2-住院病历;3-护理病历;4-护理记录;5-首页记录;6-医嘱报告;7-疾病证明;8-知情文件
    On Error GoTo errHand
    mblnPDF = blnPDF
    If blnPDF = False Then '设备输出，非 输出PDF
        strPrinterName = GetRegister(私有模块, "打印档案", "打印机", Printer.DeviceName)
        strSelect = "," & GetRegister(私有模块, "打印档案", "打印内容", "1,2,3,4,5,6,7,8,9") & ","
        With cboPrinterName
            .Clear
            For intCount = 0 To Printers.count - 1
                .AddItem Printers(intCount).DeviceName
                If Printers(intCount).DeviceName = strPrinterName Then .ListIndex = intCount
            Next
        End With
        Call zlControl.CboSetWidth(cboPrinterName.hWnd, 3000)
        Label1.Visible = False
        txtPath.Visible = False
        cmdPath.Visible = False
        picInfo.Visible = False
        Me.Caption = "输出档案"
    Else '输出PDF
        Me.Caption = "输出档案到PDF"
        strSelect = "," & GetRegister(私有模块, "打印档案", "输出PDF", "1,2,3,4,5,6,7,8,9") & ","
        Label3.Visible = False
        cboPrinterName.Visible = False
        picInfo.Visible = True
        strPath = GetRegister(私有模块, "打印档案", "PDF位置", App.Path)
        txtPath.Text = strPath: txtPath.ToolTipText = strPath
    End If
    
    mstrPrintDocIDs = ""
    mblnDoctorAdvice = blnDoctorAdvice
    Call FillVfg(parVsf, intType)
    
    strTmp = Trim(zlDatabase.GetPara("档案排序顺序", ParamInfo.系统号, 1560, "5;1;6;2;3;4;8;7;9"))
    If strTmp = "" Then strTmp = "5;1;6;2;3;4;8;7;9"
    arySerial = Split(strTmp, ";")
    
    With lst
        For intCount = 0 To UBound(arySerial)
            Select Case Val(arySerial(intCount))
            Case 1
                .AddItem "住院医嘱": .ItemData(.NewIndex) = 1
                If InStr(strSelect, ",1,") > 0 Then .Selected(.NewIndex) = True
            Case 2
                .AddItem "住院病历": .ItemData(.NewIndex) = 2
                If InStr(strSelect, ",2,") > 0 Then .Selected(.NewIndex) = True
            Case 3
                .AddItem "护理病历": .ItemData(.NewIndex) = 3
                If InStr(strSelect, ",3,") > 0 Then .Selected(.NewIndex) = True
            Case 4
                .AddItem "护理记录": .ItemData(.NewIndex) = 4
                If InStr(strSelect, ",4,") > 0 Then .Selected(.NewIndex) = True
            Case 5
                .AddItem "首页正面": .ItemData(.NewIndex) = 5
                If InStr(strSelect, ",5,") > 0 Then .Selected(.NewIndex) = True
                .AddItem "首页反面": .ItemData(.NewIndex) = 52
                If InStr(strSelect, ",52,") > 0 Then .Selected(.NewIndex) = True
                .AddItem "首页附页一": .ItemData(.NewIndex) = 53
                If InStr(strSelect, ",53,") > 0 Then .Selected(.NewIndex) = True
                .AddItem "首页附页二": .ItemData(.NewIndex) = 54
                If InStr(strSelect, ",54,") > 0 Then .Selected(.NewIndex) = True
            Case 6
                .AddItem "医嘱报告": .ItemData(.NewIndex) = 6
                If InStr(strSelect, ",6,") > 0 Then .Selected(.NewIndex) = True
            Case 7
                .AddItem "疾病证明": .ItemData(.NewIndex) = 7
                If InStr(strSelect, ",7,") > 0 Then .Selected(.NewIndex) = True
            Case 8
                .AddItem "知情文件": .ItemData(.NewIndex) = 8
                If InStr(strSelect, ",8,") > 0 Then .Selected(.NewIndex) = True
            Case 9
                .AddItem "临床路径": .ItemData(.NewIndex) = 9
                If InStr(strSelect, ",9,") > 0 Then .Selected(.NewIndex) = True
            End Select
        Next

        .ListIndex = 0
    End With
    
    Me.Show 1, frmObj
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub FillVfg(ByVal parVsf As VSFlexGrid, ByVal intType As Integer)
Dim i As Integer
    On Error GoTo errHand
    With vsf
        .Clear
        .Rows = parVsf.Rows
        .Cols = 13
        .RowHeight(0) = 350
        .TextMatrix(0, mCol.选择) = "选择"
        .TextMatrix(0, mCol.病人ID) = "病人ID"
        .TextMatrix(0, mCol.主页ID) = "主页ID"
        .TextMatrix(0, mCol.姓名) = "姓名"
        .TextMatrix(0, mCol.床号) = "床号"
        .TextMatrix(0, mCol.住院号) = "住院号"
        .TextMatrix(0, mCol.性别) = "性别"
        .TextMatrix(0, mCol.年龄) = "年龄"
        .TextMatrix(0, mCol.出院科室) = "出院科室"
        .TextMatrix(0, mCol.出院科室ID) = "出院科室ID"
        .TextMatrix(0, mCol.入院日期) = "入院日期"
        .TextMatrix(0, mCol.出院日期) = "出院日期"
        .TextMatrix(0, mCol.打印记录) = "打印记录"
        
        .ColWidth(mCol.选择) = 400
        .ColWidth(mCol.病人ID) = 0
        .ColWidth(mCol.主页ID) = 0
        .ColWidth(mCol.姓名) = 1200
        .ColWidth(mCol.床号) = 600
        .ColWidth(mCol.住院号) = 1200
        .ColWidth(mCol.入院日期) = 1800
        .ColWidth(mCol.性别) = 500
        .ColWidth(mCol.出院科室ID) = 0
        If intType = 0 Then
            .ColWidth(mCol.出院日期) = 1800
            .ColWidth(mCol.出院科室) = 800
        Else
            .ColWidth(mCol.出院日期) = 0
            .ColWidth(mCol.出院科室) = 0
        End If
        .ColWidth(mCol.年龄) = 500
        .ColWidth(mCol.打印记录) = 0
                
        For i = 1 To parVsf.Rows - 1
            .RowHeight(i) = 350
            .Cell(flexcpData, i, mCol.选择) = 0
            .TextMatrix(i, mCol.病人ID) = parVsf.TextMatrix(i, parVsf.ColIndex("病人ID"))
            .TextMatrix(i, mCol.主页ID) = parVsf.TextMatrix(i, parVsf.ColIndex("主页ID"))
            If intType = 1 Then '在院
                .TextMatrix(i, mCol.姓名) = parVsf.TextMatrix(i, parVsf.ColIndex("姓名"))
                .TextMatrix(i, mCol.床号) = parVsf.TextMatrix(i, parVsf.ColIndex("床号"))
                .TextMatrix(i, mCol.住院号) = parVsf.TextMatrix(i, parVsf.ColIndex("住院号"))
                .TextMatrix(i, mCol.性别) = parVsf.TextMatrix(i, parVsf.ColIndex("性别"))
                .TextMatrix(i, mCol.年龄) = parVsf.TextMatrix(i, parVsf.ColIndex("年龄"))
                .TextMatrix(i, mCol.入院日期) = parVsf.TextMatrix(i, parVsf.ColIndex("入院日期"))
                .TextMatrix(i, mCol.出院科室ID) = parVsf.TextMatrix(i, parVsf.ColIndex("出院科室ID"))
            Else
                .TextMatrix(i, mCol.姓名) = parVsf.TextMatrix(i, parVsf.ColIndex("姓名"))
                If parVsf.ColIndex("床号") <> -1 Then
                    .TextMatrix(i, mCol.床号) = parVsf.TextMatrix(i, parVsf.ColIndex("床号"))
                End If
                If parVsf.ColIndex("住院号") <> -1 Then
                    .TextMatrix(i, mCol.住院号) = parVsf.TextMatrix(i, parVsf.ColIndex("住院号"))
                End If
                If parVsf.ColIndex("年龄") <> -1 Then
                    .TextMatrix(i, mCol.年龄) = parVsf.TextMatrix(i, parVsf.ColIndex("年龄"))
                End If
                If parVsf.ColIndex("出院科室") <> -1 Then
                    .TextMatrix(i, mCol.出院科室) = parVsf.TextMatrix(i, parVsf.ColIndex("出院科室"))
                End If
                
                If parVsf.ColIndex("入院日期") <> -1 Then
                    .TextMatrix(i, mCol.入院日期) = parVsf.TextMatrix(i, parVsf.ColIndex("入院日期"))
                End If
                
                If parVsf.ColIndex("出院日期") <> -1 Then
                    .TextMatrix(i, mCol.出院日期) = parVsf.TextMatrix(i, parVsf.ColIndex("出院日期"))
                End If
                .TextMatrix(i, mCol.出院科室ID) = parVsf.TextMatrix(i, parVsf.ColIndex("出院科室ID"))
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
        .Row = parVsf.Row
        .TopRow = .Row
        If .Rows = 2 Then vsf_DblClick
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub PrintWithActiveEXE(ByVal strRegRange As String, ByVal strRange As String, ByVal strParPath As String)
Dim i As Long, strPrintContent As String, rs As New ADODB.Recordset, l As Long, objPrint As Object
Dim objRichEMR As Object, arrPar As Variant, arrParOne As Variant, j As Long, strEmrId As String
Dim varParam As Variant, strReportNO As String, lng科室ID As Long, blnNewTends As Boolean, intSel As Integer, strEprName As String
Dim lngPatient As Long, lngPageID As Long, lngDept As Long, lngInNo As Long, strPath As String, strFileName As String, blnDataMove As Boolean, strName As String
    On Error GoTo errHand
    For i = 1 To vsf.Rows - 1
        With vsf
            If .Cell(flexcpData, i, mCol.选择 = 1) Then '选中才输出
                l = l + 1
                lngPatient = Val(.TextMatrix(i, mCol.病人ID))
                lngInNo = Val(.TextMatrix(i, mCol.住院号))
                lngPageID = Val(.TextMatrix(i, mCol.主页ID))
                lngDept = Val(.TextMatrix(i, mCol.出院科室ID))
                strName = NVL(.TextMatrix(i, mCol.姓名))
                strPath = strParPath & "\" & strName & "_" & lngInNo
                If Not gobjFSO.FolderExists(strPath) Then
                    Call gobjFSO.CreateFolder(strPath)
                End If

                '读取记录
                Set rs = gclsPackage.GetCISStruct(lngPatient, lngPageID, lngDept, blnDataMove)
                Do Until rs.EOF
                    If NVL(rs("上级id").Value) = "" Then
                        If InStr(strRange, "," & rs("ID").Value & ",") > 0 Then
                            Select Case rs("ID").Value
                            Case "R5" '首页
                                '系统号,报表编号,病人id,主页id,1正/2反/3附一/4附二,PDFFileName
                                lng科室ID = GetlngID(lngPatient, lngPageID)
                                Select Case Val(zlDatabase.GetPara("病案首页标准", glngSys, 1261, "0"))
                                Case 0 '卫生部标准
                                    If Have部门性质(lng科室ID, "中医科") Then
                                        strReportNO = "ZL1_INSIDE_1261_4"
                                    Else
                                        strReportNO = "ZL1_INSIDE_1261_1"
                                    End If
                                Case 1    '四川省标准
                                    If Have部门性质(lng科室ID, "中医科") Then
                                        strReportNO = "ZL1_INSIDE_1261_6"
                                    Else
                                        strReportNO = "ZL1_INSIDE_1261_5"
                                    End If
                                Case 2    '云南省标准
                                    If Have部门性质(lng科室ID, "中医科") Then
                                        strReportNO = "ZL1_INSIDE_1261_8"
                                    Else
                                        strReportNO = "ZL1_INSIDE_1261_7"
                                    End If
                                Case 3     '湖南省标准
                                    If Have部门性质(lng科室ID, "中医科") Then
                                        strReportNO = "ZL1_INSIDE_1261_10"
                                    Else
                                        strReportNO = "ZL1_INSIDE_1261_9"
                                    End If
                                Case Else '当期修改时未定义
                                    If Have部门性质(lng科室ID, "中医科") Then
                                        strReportNO = "ZL1_INSIDE_1261_4"
                                    Else
                                        strReportNO = "ZL1_INSIDE_1261_1"
                                    End If
                                End Select
                                If InStr("," & strRegRange & ",", ",5,") > 0 Then '正面
                                    strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_首页正面.PDF"
                                    strPrintContent = strPrintContent & "|" & strName & ",R5," & glngSys & "," & strReportNO & "," & lngPatient & "," & lngPageID & ",1," & strFileName
                                End If
                                
                                If InStr("," & strRegRange & ",", ",52,") > 0 Then '反面
                                    strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_首页反面.PDF"
                                    strPrintContent = strPrintContent & "|" & strName & ",R5," & glngSys & "," & strReportNO & "," & lngPatient & "," & lngPageID & ",2," & strFileName
                                End If
                                
                                If InStr("," & strRegRange & ",", ",53,") > 0 Then '附一
                                    strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_首页附页一.PDF"
                                    strPrintContent = strPrintContent & "|" & strName & ",R5," & glngSys & "," & strReportNO & "," & lngPatient & "," & lngPageID & ",3," & strFileName
                                End If
                                
                                If InStr("," & strRegRange & ",", ",54,") > 0 Then '附二
                                    strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_首页附页二.PDF"
                                    strPrintContent = strPrintContent & "|" & strName & ",R5," & glngSys & "," & strReportNO & "," & lngPatient & "," & lngPageID & ",4," & strFileName
                                End If
                            Case "R1"               '医嘱
                                '系统号,报表编号,病人id,主页id,医嘱单A0,A1/医嘱本B,PDFFileName
                                If mblnDoctorAdvice Then
                                    strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_医嘱.PDF"
                                    strPrintContent = strPrintContent & "|" & strName & ",R1," & glngSys & ",zl1_INSIDE_1254_1," & lngPatient & "," & lngPageID & ",A0," & strFileName
                                    strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_临嘱.PDF"
                                    strPrintContent = strPrintContent & "|" & strName & ",R1," & glngSys & ",zl1_INSIDE_1254_2," & lngPatient & "," & lngPageID & ",A1," & strFileName
                                Else
                                    strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_医嘱.PDF"
                                    strPrintContent = strPrintContent & "|" & strName & ",R1," & glngSys & ",ZL1_INSIDE_1560," & lngPatient & "," & lngPageID & ",B," & strFileName
                                End If
                            Case "R9"               '临床路径
                                'FileName,病人ID,主页ID
                                strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_临床路径.PDF"
                                strPrintContent = strPrintContent & "|" & strName & ",R9," & glngSys & "," & strFileName & "," & lngPatient & "," & lngPageID
                            End Select
                        End If
                    Else
                        If InStr(strRange, "," & rs("上级id").Value & ",") > 0 Then
                            varParam = Split(rs("参数").Value, ";")
                            Select Case rs("上级id").Value
                            Case "R2"               '住院病历
                                '系统号,FileName,ID
                                strEprName = Split(rs("名称").Value, "【")(0)
                                strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & Val(varParam(0)) & ".PDF"
                                strPrintContent = strPrintContent & "|" & strName & ",R2," & glngSys & "," & strFileName & "," & Val(varParam(0))
                            Case "R3"               '护理病历
                                '系统号,FileName,ID
                                strEprName = Split(rs("名称").Value, "【")(0)
                                strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & Val(varParam(0)) & ".PDF"
                                strPrintContent = strPrintContent & "|" & strName & ",R3," & glngSys & "," & strFileName & "," & Val(varParam(0))
                            Case "R4"               '护理记录
                                '系统号,新版N/旧版O,体温单1/护理记录单2/产程图3,FileName,病人ID,主页ID,科室ID,婴儿序号,lngKey/lngFileID,Period
                                blnNewTends = Get新版护理(lngPatient, lngPageID)
                                If blnNewTends = False Then
                                    If UBound(varParam) >= 1 Then
                                        If Val(varParam(1)) = -1 Then '体温单
                                            strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_体温单_" & Val(varParam(0)) & ".PDF"
                                            strPrintContent = strPrintContent & "|" & strName & ",R4," & glngSys & ",O,1," & strFileName & "," & lngPatient & "," & lngPageID & "," & Val(Split(varParam(0), "_")(0)) & "," & Val(varParam(4))
                                        Else '护理记录
                                            strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_护理记录_" & Val(varParam(3)) & ".PDF"
                                            strPrintContent = strPrintContent & "|" & strName & ",R4," & glngSys & ",O,2," & strFileName & "," & lngPatient & "," & lngPageID & "," & Val(Split(varParam(0), "_")(0)) & "," & Val(varParam(4)) & "," & Val(varParam(3)) & "," & CStr(varParam(2))
                                        End If
                                    End If
                                Else
                                    '此参数保存 保留
                                    varParam = Split(rs("参数").Value, ";")
                                    If UBound(varParam) >= 1 Then
                                        Select Case Val(varParam(1))
                                            Case -1 '体温单
                                                intSel = 1
                                            Case 1  '产程图
                                                intSel = 3
                                            Case Else '记录单
                                                intSel = 2
                                        End Select
                                        strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & Decode(intSel, 1, "体温单", 2, "护理记录", "产程图") & "_" & Val(varParam(3)) & ".PDF"
                                        strPrintContent = strPrintContent & "|" & strName & ",R4," & glngSys & ",N," & intSel & "," & strFileName & "," & lngPatient & "," & lngPageID & "," & Val(varParam(0)) & "," & Val(varParam(4)) & "," & Val(varParam(3))
                                    End If
                                End If
                            Case "R6"               '医嘱报告
                                '系统号,FileName,ID
                                strEprName = Split(rs("名称").Value, "【")(0)
                                If UBound(Split(strEprName, ">")) > 0 Then
                                    strEprName = Split(strEprName, ">")(1)
                                End If
                                strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & Val(varParam(0)) & ".PDF"
                                strPrintContent = strPrintContent & "|" & strName & ",R6," & glngSys & "," & strFileName & "," & Val(varParam(0))
                            Case "R7"               '疾病证明
                                '系统号,FileName,ID
                                strEprName = Split(rs("名称").Value, "【")(0)
                                strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & Val(varParam(0)) & ".PDF"
                                strPrintContent = strPrintContent & "|" & strName & ",R7," & glngSys & "," & strFileName & "," & Val(varParam(0))
                            Case "R8"               '知情文件
                                '系统号,FileName,ID
                                strEprName = Split(rs("名称").Value, "【")(0)
                                strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & Val(varParam(0)) & ".PDF"
                                strPrintContent = strPrintContent & "|" & strName & ",R8," & glngSys & "," & strFileName & "," & Val(varParam(0))
                            End Select
                        End If
                    End If
                    rs.MoveNext
                Loop
                
                If Not gobjEmr Is Nothing Then
                    If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
                        Set gobjEmr = Nothing
                    End If
                    If Not gobjEmr Is Nothing Then
                        Set rs = gclsPackage.GetEmrCISStruct(lngPatient, lngPageID)
                        Do Until rs.EOF
                            strEmrId = Split(rs!参数, "|")(0)
                            If InStr(strPrintContent, strEmrId) = 0 Then
                                If UBound(Split(rs!参数, "|")) = 0 Then
                                    strEprName = Split(rs("名称").Value, "【")(0)
                                Else
                                    strEprName = rs!Title
                                End If
                                strFileName = strPath & "\" & strName & "_" & lngInNo & "_" & lngPageID & "_" & strEprName & "_" & strEmrId & ".PDF"
                                strPrintContent = strPrintContent & "|" & strName & ",EMR," & glngSys & "," & strFileName & "," & strEmrId
                            End If
                            rs.MoveNext
                        Loop
                    End If
                End If
                
                ''病人循环,每10个病人输出一次，以减少zlCisAuditPrint初始化对象时间
                If l Mod 10 = 0 Then
                    strPrintContent = Mid(strPrintContent, 2)
                    Set objPrint = Nothing
                    Set objPrint = CreateObject("zlCisAuditPrint.clsPrint")
                    Call objPrint.PrintDocument(Me, gstrInputSeverName, gstrInputUser, gstrInputPwd, strPrintContent, "TinyPDF")
                    
                    '新病历输出
                    arrPar = Split(strPrintContent, "|")
                    For j = 0 To UBound(arrPar)
                        arrParOne = Split(arrPar(j), ",")
                        If arrParOne(1) = "EMR" Then             '新病历
                            Label2.Caption = "开始输出" & arrParOne(0) & "病历"
                                                            
                            If objRichEMR Is Nothing Then
                                Set objRichEMR = DynamicCreate("zlRichEMR.clsDockContent", "新版病历", False)
                                If Not objRichEMR Is Nothing Then Call objRichEMR.Init(gobjEmr, gcnOracle, glngSys, 0)
                            End If
                            Call objRichEMR.zlShowDoc(arrParOne(4), "")
                            Call zlCommFun.PDFFile(arrParOne(3))
                            Call objRichEMR.zlPrintDoc(False, "TinyPDF")
                        End If
                    Next
                    
                    l = 0: strPrintContent = ""
                End If
            End If
        End With
    Next
    
    If l <> 0 Then
        strPrintContent = Mid(strPrintContent, 2)
        Set objPrint = Nothing
        Set objPrint = CreateObject("zlCisAuditPrint.clsPrint")
        Call objPrint.PrintDocument(Me, gstrInputSeverName, gstrInputUser, gstrInputPwd, strPrintContent, "TinyPDF")
        
        '新病历输出
        arrPar = Split(strPrintContent, "|")
        For j = 0 To UBound(arrPar)
            arrParOne = Split(arrPar(j), ",")
            If arrParOne(1) = "EMR" Then             '新病历
                Label2.Caption = "开始输出" & arrParOne(0) & "病历"
                                                
                If objRichEMR Is Nothing Then
                    Set objRichEMR = DynamicCreate("zlRichEMR.clsDockContent", "新版病历", False)
                    If Not objRichEMR Is Nothing Then Call objRichEMR.Init(gobjEmr, gcnOracle, glngSys, 0)
                End If
                Call objRichEMR.zlShowDoc(arrParOne(4), "")
                Call zlCommFun.PDFFile(arrParOne(3))
                Call objRichEMR.zlPrintDoc(False, "TinyPDF")
            End If
        Next
        
        l = 0: strPrintContent = ""
    End If
    Exit Sub
errHand:
    zlCommFun.StopFlash
    If ErrCenter = 1 Then
        Resume
    End If
    Label2.Caption = ""
    mstrPrintDocIDs = ""
End Sub

Private Sub cmdOk_Click()
Dim strRange As String, strRegRange As String, i As Integer, strErr As String
Dim strParPath As String, strPrinterName As String
    
    '统计并记录打印类别
    strRange = ""
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) = True Then
            strRegRange = strRegRange & "," & lst.ItemData(i)
            If InStr(",5,52,53,54,", "," & lst.ItemData(i) & ",") > 0 Then '首页的正反面，类型都是5
                If InStr(strRange, "R5") = 0 Then '没加
                    strRange = strRange & ",R5"
                End If
            Else
                strRange = strRange & ",R" & lst.ItemData(i)
            End If
        End If
    Next
    If strRange <> "" Then
        strRange = strRange & ","
        strRegRange = Mid(strRegRange, 2)
    Else
        MsgBox "请选择需要输出的档案！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mblnPDF Then
        Call SetRegister(私有模块, "打印档案", "输出PDF", strRegRange)
        '输出位置
        If txtPath.Text = "" Then
            MsgBox "请选择输出档案的位置！", vbInformation, gstrSysName
            Exit Sub
        Else
            strParPath = txtPath.Text
            If gobjFSO.FolderExists(strParPath) = False Then
                MsgBox "指定目录不存在，请检查！", vbInformation, gstrSysName
                Exit Sub
            End If
            Call SetRegister(私有模块, "打印档案", "PDF位置", txtPath.Text)
        End If
        On Error Resume Next
        Err.Clear
        Call zlCommFun.PDFInitialize(strErr)
        If Err.Number <> 0 Then
            Err.Raise vbObjectError, , "PDF设备初始化失败"
        End If
    Else
        strPrinterName = cboPrinterName.Text
        Call SetRegister(私有模块, "打印档案", "打印内容", strRegRange)
        Call SetRegister(私有模块, "打印档案", "打印机", strPrinterName)
    End If
    On Error GoTo 0

    cmdCancel.Enabled = False: cmdOK.Enabled = False: fraPageScope.Enabled = False
    cmdSelectAll.Enabled = False: cmdUnSelectAll.Enabled = False: Frame2.Enabled = False
    
    If mblnPDF Then
        Call PrintWithActiveEXE(strRegRange, strRange, strParPath)
    Else
        Call PrintDocument(strRegRange, strRange, strParPath, strPrinterName)
    End If
    
    Label2.Caption = "已完成输出"
    mstrPrintDocIDs = ""
    cmdCancel.Enabled = True: cmdOK.Enabled = True: fraPageScope.Enabled = True: cmdSelectAll.Enabled = True: cmdUnSelectAll.Enabled = True: Frame2.Enabled = True
End Sub


Private Sub PrintDocument(ByVal strRegRange As String, ByVal strRange As String, ByVal strParPath As String, ByVal strPrinterName As String)
Dim i As Integer, rs As New ADODB.Recordset, blnTrans As Boolean, lngNo As Long
Dim clsPath As zlCISPath.clsDockPath, clsTendsNew As zl9TendFile.clsTendFile, objPacsDoc As Object
Dim varParam As Variant, strReportNO As String, lng科室ID As Long, blnNewTends As Boolean, intSel As Integer, strEprName As String
Dim lngPatient As Long, lngPageID As Long, lngDept As Long, lngInNo As Long, strPath As String, strFileName As String, blnDataMove As Boolean, strName As String
    
    On Error GoTo errHand

    '输出对象
    If mclsDockAduits Is Nothing Then
        Set mclsDockAduits = New zlRichEPR.clsDockAduits
    End If
    Set clsPath = New zlCISPath.clsDockPath
    Set clsTendsNew = New zl9TendFile.clsTendFile: Call clsTendsNew.InitTendFile(gcnOracle, glngSys)
    
    '调用打印
    For i = 1 To vsf.Rows - 1
        With vsf
            If .Cell(flexcpData, i, mCol.选择 = 1) Then '选中才输出
                .Row = i
                lngPatient = Val(.TextMatrix(i, mCol.病人ID))
                lngInNo = Val(.TextMatrix(i, mCol.住院号))
                lngPageID = Val(.TextMatrix(i, mCol.主页ID))
                lngDept = Val(.TextMatrix(i, mCol.出院科室ID))
                strName = NVL(.TextMatrix(i, mCol.姓名))
                '批量打印
                Call gclsPackage.FuncPrintBatch(lngPatient, lngPageID, lngDept, strRange, strRegRange, mclsDockAduits, clsPath, clsTendsNew, _
                    mblnPDF, strParPath, strName, lngInNo, Me, Label2.Caption, blnDataMove, strPrinterName, mblnDoctorAdvice, mstrPrintDocIDs)
                .TopRow = i
                .Cell(flexcpData, i, mCol.选择) = 0
                Set .Cell(flexcpPicture, i, mCol.选择) = Nothing
            End If
        End With
    Next
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Label2.Caption = ""
    mstrPrintDocIDs = ""
End Sub

Private Sub cmdPath_Click()
Dim strPath As String
    strPath = zl9Comlib.OS.OpenDir(Me.hWnd, "请选择导出文件位置")
    If strPath = "" Then Exit Sub
    txtPath.Text = strPath: txtPath.ToolTipText = strPath
End Sub

Private Sub cmdSelectAll_Click()
    vsf.Cell(flexcpData, 1, mCol.选择, vsf.Rows - 1, mCol.选择) = 1
    Set vsf.Cell(flexcpPicture, 1, mCol.选择, vsf.Rows - 1, mCol.选择) = img16.ListImages("Selected").Picture
End Sub

Private Sub cmdUnSelectAll_Click()
    vsf.Cell(flexcpData, 1, mCol.选择, vsf.Rows - 1, mCol.选择) = 0
    Set vsf.Cell(flexcpPicture, 1, mCol.选择, vsf.Rows - 1, mCol.选择) = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
If cmdCancel.Enabled = False Then
    Cancel = 1
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If cmdCancel.Enabled = False Then
    Cancel = 1
End If

Set mclsDockAduits = Nothing
Unload mfrmTipInfo
Set mfrmTipInfo = Nothing
End Sub

Private Sub mclsDockAduits_AfterEprPrint(ByVal lngRecordId As Long)
    mstrPrintDocIDs = mstrPrintDocIDs & lngRecordId & ","
End Sub
Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'显示指定病历列表行的历史签名记录
Dim strTipInfo As String, lngRow As Long
    If picInfo.Visible = False Then Exit Sub
    
    lngRow = vsf.MouseRow
    If lngRow <= 0 Then Exit Sub
    
    strTipInfo = vsf.Cell(flexcpData, lngRow, mCol.打印记录)

    If strTipInfo = "" Then '如果没有获取过，则立即获取并记录在列表中
        strTipInfo = GetPrintLog(vsf.TextMatrix(lngRow, mCol.病人ID), vsf.TextMatrix(lngRow, mCol.主页ID)) '提取打印记录
        vsf.Cell(flexcpData, lngRow, mCol.打印记录) = strTipInfo
    End If
    
    mfrmTipInfo.ShowTipInfo picInfo.hWnd, strTipInfo, True
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    If picInfo.Visible Then
        picInfo.Move vsf.Cell(flexcpLeft, NewTopRow, mCol.姓名) + vsf.Cell(flexcpWidth, NewTopRow, mCol.姓名) - picInfo.Width - 30
    End If
End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
        Call vsf_DblClick
    End If
End Sub

Private Sub vsf_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngCol As Long, lngRow As Long
    lngCol = vsf.MouseCol: lngRow = vsf.MouseRow
    If lngRow <= 0 Then picInfo.Visible = False: Exit Sub
    If Val(vsf.TextMatrix(lngRow, mCol.病人ID)) <> 0 Then
        If Val(picInfo.Tag) = lngRow And picInfo.Visible Then Exit Sub
        picInfo.Tag = lngRow
        picInfo.Move vsf.Cell(flexcpLeft, lngRow, mCol.姓名) + vsf.Cell(flexcpWidth, lngRow, mCol.姓名) - picInfo.Width - 30, vsf.Cell(flexcpTop, lngRow, mCol.姓名) + 15
        If vsf.RowSel = lngRow Then
            picInfo.BackColor = vsf.BackColorSel
        Else
            picInfo.BackColor = &H80000005
        End If
        picInfo.Visible = True
    Else
        picInfo.Visible = False
    End If
End Sub
Private Sub vsf_SelChange()
    If picInfo.Visible Then
        picInfo.BackColor = vsf.BackColorSel
    End If
End Sub
Private Sub vsf_DblClick()
Dim lngRow As Long, l As Long, lCheck As Long
    With vsf
        lngRow = .Row
        If lngRow < 1 Then Exit Sub
        If .Cell(flexcpData, lngRow, mCol.选择) = 0 Then
            .Cell(flexcpData, lngRow, mCol.选择) = 1
            Set .Cell(flexcpPicture, lngRow, mCol.选择) = img16.ListImages("Selected").Picture
        Else
            .Cell(flexcpData, lngRow, mCol.选择) = 0
            Set .Cell(flexcpPicture, lngRow, mCol.选择) = Nothing
        End If
        
        For l = 1 To .Rows - 1
            If .Cell(flexcpData, l, mCol.选择) = 1 Then
                lCheck = lCheck + 1
            End If
        Next
        Frame2.Caption = "待输出病人清单(由主界面条件过滤)" & " 共" & .Rows - 1 & "行，已选中" & lCheck & "行"
    End With
End Sub
Private Function GetPrintLog(ByVal lngPatient As Long, ByVal lngPageID As Long) As String
Dim rs As New ADODB.Recordset
    gstrSQL = "Select 打印次数 As 打印次, 打印内容, 打印人, 打印时间 From 病案打印记录 Where 病人id = [1] And 主页id = [2] Order By 打印时间, 打印序号"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatient, lngPageID)
    Do Until rs.EOF
        GetPrintLog = GetPrintLog & vbCrLf & Rpad(rs!打印人, 10) & Rpad(Format(rs!打印时间, "yyyy-mm-dd hh:MM"), 20) & Rpad(rs!打印内容, 40)
        rs.MoveNext
    Loop
    GetPrintLog = Rpad("打印人", 10) & Rpad("打印时间", 20) & Rpad("打印内容", 40) & GetPrintLog
End Function
