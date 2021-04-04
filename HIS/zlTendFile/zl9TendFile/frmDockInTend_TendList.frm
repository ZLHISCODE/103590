VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockInTend_TendList 
   BorderStyle     =   0  'None
   Caption         =   "护理文件列表"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer TimFresh 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   0
      Top             =   3630
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   3300
      Left            =   30
      ScaleHeight     =   3300
      ScaleWidth      =   6690
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   6690
      Begin MSComctlLib.ImageList imgData 
         Left            =   1005
         Top             =   1695
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
               Picture         =   "frmDockInTend_TendList.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockInTend_TendList.frx":6862
               Key             =   "体温"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockInTend_TendList.frx":6DFC
               Key             =   "普通"
            EndProperty
         EndProperty
      End
      Begin VB.Frame fra 
         Height          =   525
         Left            =   0
         TabIndex        =   1
         Top             =   -90
         Width           =   6015
         Begin VB.ComboBox cboBaby 
            Height          =   300
            Left            =   630
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   150
            Width           =   1350
         End
         Begin VB.Label lbl病人 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "查看"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   4
            Top             =   210
            Width           =   360
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgFile 
         Height          =   1095
         Left            =   -15
         TabIndex        =   3
         Top             =   435
         Width           =   6060
         _cx             =   10689
         _cy             =   1931
         Appearance      =   2
         BorderStyle     =   0
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Bindings        =   "frmDockInTend_TendList.frx":7396
      Left            =   135
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmDockInTend_TendList.frx":73AA
   End
End
Attribute VB_Name = "frmDockInTend_TendList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'绑定快捷键时,ID值如大于无符号整型的取值范围则无法绑定,也就是0-65535
Private Const conMenu_Add As Long = 32761 '新增
Private Const conMenu_Modify As Long = 32762 '修改
Private Const conMenu_Delete As Long = 32763 '删除

Private Enum mCol
    f标志 = 0: fID: f格式ID: f文件: f开始日期: f科室ID: f科室: f保留: f创建日期
End Enum

Private mblnInit As Boolean
Private mblnNoRefresh As Boolean
Private mstrPrivs As String                             '当前使用者对本程序(1255)的权限串
Private mlngPatiID As Long                              '病人id
Private mlngPageId As Long                              '主页id
Private mlngDeptId As Long                              '当前操作科室id，如病人科室和当前科室不一致，则不能操作归档外的功能
Private mlngFileID As Long                              '需要定位到的文件ID
Private mlngFormatID As Long                            '文件格式ID
Private mlng序号 As Integer                             '选择病人本人或婴儿
Private mblnEdit As Boolean                             '是否允许操作，通常由上级程序根据当前操作科室是否当前病人病区决定。
Private mblnDoctorStation As Boolean
Private mintCurveReSize As Integer                      '体温单查看是否是缩小模式 0缩小 1原始大小
Private rsTemp As New ADODB.Recordset
Private mintBaby As Integer
Private mfrmMain As Object
Private mbytFontSize As Byte
Private mblnChange As Boolean                           '修改标志
Private mblnSign As Boolean                             '签名标志
Private mblnArchive As Boolean                          '归档标志
Private mblnRefreshFontSize As Boolean                  '是否在刷新数据后自动调用设置字体功能(对于内部调用的自动刷新)
Private mblnTemparatureChat As Boolean                  '是否是标准体温单

'在已可方便查看体温单与护理记录单的情况下,弹出式查看已失去意义,先写到这里
Public Event Activate()         '更新按钮与菜单
Public Event ViewFile(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, _
    ByVal intBaby As Integer, ByVal blnChildForm As Boolean, ByVal strPrivs As String, ByVal blnEdit As Boolean, ByVal bytSize As Byte)
Public Event ViewAnimalHeat(strPara As String, bytMode As Byte, strPrivs As String, ByVal bytSize As Byte)
Public Event ShowData(intBaby As Integer, lngFile As Long, lngDept As Long, bytSel As Byte, ByVal intCurveReSize As Integer)                 '通知数据页面刷新
Public Event PrintTendFile(ByVal bytKind As Byte, ByVal bytMode As Byte)
Public Event SaveDocument(blnSave As Boolean)                                                               '假则恢复数据
Public Event SignDocument(blnOK As Boolean, blnVerify As Boolean, blnExchange As Boolean)                                           '假则取消签名
Public Event ArchiveDocument(blnOK As Boolean)                                                              '假则取消归档
Public Event SignMarker()
Public Event ViewCaveData(ByVal intDataEditor As Integer)
Public Event Viewpartogram(strPara As String, bytMode As Byte, strPrivs As String, ByVal bytSize As Byte)
Public Event ViewpartogramEditor(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, _
    ByVal intBaby As Integer, ByVal strPrivs As String, ByVal bytSize As Byte)
Public Event ViewReSetFontSize(ByVal intSEL As Integer, ByVal bytSize As Byte)
Public Event BulkPrintDocument(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, ByVal intBaby As Integer)

Public Sub SetChange(ByVal blnChange As Boolean)
    mblnChange = blnChange
End Sub

Public Sub SetState(ByVal blnSign As Boolean, ByVal blnArchive As Boolean)
    mblnArchive = blnArchive
    mblnSign = blnSign
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strInfo As String
    Dim byt护理等级 As Byte
    Dim rs As New ADODB.Recordset
    
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_FileMan
        Call frmNurseFileMan.ShowEditor(mlngPatiID, mlngPageId, mintBaby, mstrPrivs, False, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
    Case conMenu_File_Open
        Call vfgFile_DblClick
'        With vfgFile
'            strInfo = Val(.TextMatrix(.ROW, mCol.f科室ID))
'            If Val(.TextMatrix(.ROW, mCol.f保留)) = -1 Then
'                '体温单查看：病人ID;主页ID;病区ID;出院;编辑;婴儿
'                If Not CreateBodyEditor Then Exit Sub
'                RaiseEvent ViewAnimalHeat(mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";" & Val(.TextMatrix(.ROW, mCol.fID)) & ";0;0;" & mintBaby & ";1", 0, mstrPrivs)
'            ElseIf Val(.TextMatrix(.ROW, mCol.f保留)) = 1 Then
'                '产程图查看:文件ID;病人ID;主页ID;病区ID
'                If Not CreatePartogram Then Exit Sub
'                RaiseEvent Viewpartogram(Val(.TextMatrix(.ROW, mCol.fID)) & ";" & mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId, 1, mstrPrivs)
'            Else
'                RaiseEvent ViewFile(Val(.TextMatrix(.ROW, mCol.fID)), mlngPatiId, mlngPageId, mlngDeptId, mintBaby, False, mstrPrivs, True)
'            End If
'        End With
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintSet
        With vfgFile
            If Val(.TextMatrix(.ROW, mCol.f保留)) = -1 Then
                If Not CreateBodyEditor Then Exit Sub
                Call gobjBodyEditor.zlPrintSet(Me)
            ElseIf Val(.TextMatrix(.ROW, mCol.f保留)) = 1 Then
                If Not CreatePartogram Then Exit Sub
                Call gobjPartogram.zlPrintSet(Me, 1)
            Else
                frmPrintSet.Show 1
            End If
        End With
    Case conMenu_File_Preview
        ''1-预览,2-打印
        With vfgFile
            If Val(.TextMatrix(.ROW, mCol.f保留)) = -1 Then
                RaiseEvent PrintTendFile(1, 1)
            ElseIf Val(.TextMatrix(.ROW, mCol.f保留)) = 1 Then
                RaiseEvent PrintTendFile(3, 1)
            Else
                RaiseEvent PrintTendFile(2, 1)
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print
        With vfgFile
            If Val(.TextMatrix(.ROW, mCol.f保留)) = -1 Then
                RaiseEvent PrintTendFile(1, 2)
            ElseIf Val(.TextMatrix(.ROW, mCol.f保留)) = 1 Then
                RaiseEvent PrintTendFile(3, 2)
            Else
                RaiseEvent PrintTendFile(2, 2)
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel
        With vfgFile
            If Val(.TextMatrix(.ROW, mCol.f保留)) = -1 Then
                MsgBox "对不起，体温单不支持输出到Excel！", vbInformation, gstrSysName
            ElseIf Val(.TextMatrix(.ROW, mCol.f保留)) = 1 Then
                MsgBox "对不起，产程图不支持输出到Excel！", vbInformation, gstrSysName
            Else
                RaiseEvent PrintTendFile(2, 3)
            End If
        End With
    '51588:刘鹏飞,2012-12-12,护理文件添加批量打印
    Case conMenu_File_Print * 100# + 1
        gstrSQL = "SELECT A.ID FROM 病人护理文件 A,病历文件列表 B WHERE A.格式ID=B.ID  And A.病人ID=[1] And A.主页ID=[2]"
        Set rs = zldatabase.OpenSQLRecord(gstrSQL, "检查是否存在要打印的文件", mlngPatiID, mlngPageId)
        If rs.RecordCount = 0 Then
            MsgBox "该病人没有任何护理文件,请先添加护理文件！", vbInformation, gstrSysName
            Exit Sub
        End If
        RaiseEvent BulkPrintDocument(mlngPatiID, mlngPageId, mlngDeptId, mintBaby)
    Case conMenu_Tool_Sign
        RaiseEvent SignDocument(True, False, False)
    '51589:刘鹏飞,2013-03-01,添加交班签名
    Case conMenu_Tool_SignShiftExchange  '交班签名
        RaiseEvent SignDocument(True, False, True)
    Case conMenu_Tool_SignEarse
        RaiseEvent SignDocument(False, False, False)
    Case conMenu_Tool_SignAuditAffirm
        RaiseEvent SignDocument(True, True, False)
    Case conMenu_Tool_SignAuditCancel
        RaiseEvent SignDocument(False, True, False)
    Case conMenu_Edit_Archive * 10
        RaiseEvent ArchiveDocument(True)
    Case conMenu_Edit_UnArchive
        If Val(Split(EprIsCommit(mlngPatiID, mlngPageId), "|")(2)) = 0 Then
            MsgBox "该病人的病案已提交审查[状态：" & gstrMecState & "]，不能撤销归档，请取消审查后再试！", vbInformation, gstrSysName
            Exit Sub
        End If
        RaiseEvent ArchiveDocument(False)
    Case conMenu_Edit_Save
        RaiseEvent SaveDocument(True)
    Case conMenu_Tool_SignVerify
        RaiseEvent SignMarker
    Case conMenu_Edit_Transf_Cancle
        RaiseEvent SaveDocument(False)
    Case conMenu_File_PrintDayDetail, conMenu_Edit_Curve, conMenu_Edit_CurveTable, conMenu_Edit_Curve_Show, conMenu_Edit_Surgery_Edit '批量录入,设置记录,显示,手术/分娩设置
        If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) = -1 Then
            On Error Resume Next
            Dim strDLL As String
            Dim strSQL As String
            Dim objChart As Object
            Dim rsTemp As New ADODB.Recordset
            
            strSQL = " Select 新部件 From 体温部件 Where Nvl(启用,0)=1"
            Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "提取体温部件")
            If Err <> 0 Then
                strDLL = "zl9TemperatureChart"
            Else
                If rsTemp.RecordCount = 0 Then
                    strDLL = "zl9TemperatureChart"
                Else
                    strDLL = NVL(rsTemp!新部件, "zl9TemperatureChart")
                End If
            End If
            
            Err = 0
            strDLL = strDLL & ".clsBodyEditor"
            Set objChart = CreateObject(strDLL)
            If Err <> 0 Then
                MsgBox "    创建体温部件失败！" & vbCrLf & "    程序将创建标准的体温部件进行数据展现，请检查指定的体温部件是否存在或已损坏！" & vbCrLf & "    详细错误：" & Err.Description, vbInformation, gstrSysName
                
                '如果创建指定的体温部件出错则创建标准的体温部件，因为这里不处理的话，后面可能存在直接使用体温部件中的对象，从而导致程序崩溃
                strDLL = "zl9TemperatureChart.clsBodyEditor"
                Set objChart = CreateObject(strDLL)
            End If
            
            On Error GoTo ErrHand
            Call objChart.InitBodyEditor(glngSys, gcnOracle)
            Select Case Control.ID
                Case conMenu_File_PrintDayDetail
                    Call objChart.BodyMutilEditor(Me, mlngDeptId, mstrPrivs, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
                Case conMenu_Edit_Curve
                    RaiseEvent ViewCaveData(0)
                Case conMenu_Edit_CurveTable
                    RaiseEvent ViewCaveData(-1)
                Case Else
                    RaiseEvent ViewCaveData(1)
            End Select
        Else
            If Control.ID <> conMenu_File_PrintDayDetail Then Exit Sub
            Dim frmTendFileMutil As New frmTendFileMutilEditor
            Call frmTendFileMutil.ShowMe(Me, mlngDeptId, mstrPrivs, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
        End If
    Case conMenu_Edit_Billing '产程数据编辑
        If Not CreatePartogram Then Exit Sub
        RaiseEvent ViewpartogramEditor(Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)), mlngPatiID, mlngPageId, mlngDeptId, 0, mstrPrivs, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
    Case conMenu_Tool_Option '护理选项
        '参数重整后，因产程图和记录单参数都是公共的，取消原有参数设置界面
        If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) = 1 Then '产程图
'            If Not CreatePartogram Then Exit Sub
'            If gobjPartogram.zlPartogramPara(Me, mstrPrivs) Then
'                Call RefreshData(mlngPatiID, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit, Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)))
'            End If
        ElseIf Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) = -1 Then '体温单
            If Not CreateBodyEditor Then Exit Sub
            If gobjBodyEditor.GetCaseTendBodyPara.ShowPara(Me, mstrPrivs) Then
                Call RefreshData(mlngPatiID, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit, Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)))
            End If
        Else '记录单
'            If frmTendPara.ShowPara(Me, mstrPrivs) Then
'                Call RefreshData(mlngPatiID, mlngPageId, mlngDeptId, mblnDoctorStation, mblnEdit, Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)))
'            End If
        End If
    End Select
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If

    Call SaveErrLog
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    If Not mblnInit Then Exit Sub
    
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_EditPopup
         Control.Visible = (mblnDoctorStation = False)
         Control.Enabled = Control.Visible
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_FileMan
        Control.Visible = (InStr(1, mstrPrivs, "护理文件管理") > 0 And mblnDoctorStation = False And Not gblnMoved)
        Control.Enabled = (mlngPatiID > 0) And Not mblnArchive And Control.Visible And mblnEdit
    Case conMenu_File_Open
        Control.Visible = True
        Control.Enabled = (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintSet
        'Control.Enabled = Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) = -1
    Case conMenu_File_Preview, conMenu_File_Print
        Control.Enabled = (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel
        Control.Enabled = (vfgFile.Rows > 1 And (InStr(1, ",-1,1,", "," & Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) & ",") = 0))
    Case conMenu_File_ExportToXML, conMenu_File_RowPrint, conMenu_Edit_Audit, conMenu_Edit_Sort, _
        conMenu_Tool_Monitor, conMenu_Edit_Archive * 10 + 1
        Control.Visible = False
        Control.Enabled = False
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_PrintDayDetail
        Control.Enabled = (mblnEdit And mlngPatiID > 0 And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f保留)) <> 1))

        Control.Visible = (InStr(1, mstrPrivs, "护理记录登记") > 0 And mblnDoctorStation = False And Not gblnMoved And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f保留)) <> 1))
        If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "护理记录登记") > 0)
    Case conMenu_Edit_Curve, conMenu_Edit_CurveTable '设置记录
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "体温单作图") > 0) And Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) = -1
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Control.Visible And mblnEdit
    Case conMenu_Edit_Curve_Show '显示
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "体温单作图") > 0) And Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) = -1 And mblnTemparatureChat = False
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Control.Visible And mblnEdit
    Case conMenu_Edit_Surgery_Edit '手术/分娩设置
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "体温单作图") > 0) And Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) = -1 And mblnTemparatureChat = True
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Control.Visible And mblnEdit
    Case conMenu_Edit_Billing  '产程数据编辑
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "产程图作图") > 0) And Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) = 1
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Not mblnArchive And Control.Visible And mblnEdit
    Case conMenu_Tool_Sign  '签名
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "护理记录签名") > 0) And (InStr(1, ",-1,1,", "," & Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) & ",") = 0)
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Not mblnArchive And Not mblnChange And Control.Visible And mblnEdit
    '51589:刘鹏飞,2013-03-01,添加交班签名
    Case conMenu_Tool_SignShiftExchange  '交班签名
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "护理记录签名") > 0) And (InStr(1, ",-1,1,", "," & Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) & ",") = 0)
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Not mblnArchive And Not mblnChange And Control.Visible And mblnEdit
    Case conMenu_Tool_SignEarse  '取消签名
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "取消记录签名") > 0) And (InStr(1, ",-1,1,", "," & Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) & ",") = 0)
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Not mblnArchive And Not mblnChange And Control.Visible And mblnEdit
    Case conMenu_Tool_SignAuditAffirm, conMenu_Tool_SignAuditCancel  '审签,取消审签
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "护理记录审签") > 0) And (InStr(1, ",-1,1,", "," & Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) & ",") = 0)
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And Not mblnArchive And Not mblnChange And Control.Visible And mblnEdit
        If Control.ID = conMenu_Tool_SignAuditCancel And Control.Enabled Then
            Control.Enabled = (InStr(1, mstrPrivs, "取消记录签名") > 0)
        End If
    Case conMenu_Edit_Archive * 10 '归档
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "护理记录归档") > 0)
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) > 0) And Not mblnArchive And Not mblnChange And Control.Visible And mblnEdit
    Case conMenu_Edit_UnArchive  '取消归档
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "取消记录归档") > 0)
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) > 0) And mblnArchive And Not mblnChange And Control.Visible And mblnEdit
    Case conMenu_Edit_Save  '保存
        Control.Visible = Not mblnDoctorStation And Not gblnMoved And (InStr(1, mstrPrivs, "护理记录登记") > 0)
        Control.Enabled = mblnChange And Not mblnArchive And Control.Visible And (InStr(1, ",-1,1,", "," & Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) & ",") = 0)
    Case conMenu_Edit_Transf_Cancle  '取消
        Control.Visible = Not mblnDoctorStation And Not gblnMoved
        Control.Enabled = mblnChange And Not mblnArchive And Control.Visible And (InStr(1, ",-1,1,", "," & Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) & ",") = 0)
    Case conMenu_Tool_SignVerify
        Control.Visible = (Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) = 0) And Not mblnDoctorStation And Not gblnMoved
        Control.Enabled = Not mblnChange And Not mblnArchive And Control.Visible And Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) = 0 And mblnEdit
    Case conMenu_Tool_Option '护理选项
        If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) = 1 Then
            Control.Caption = "产程选项"
        Else
            Control.Caption = "护理选项"
        End If
        '参数重整后，因产程图和记录单参数都是公共的，取消原有参数设置界面
        Control.Visible = Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) = -1
        Control.Enabled = (mlngPatiID > 0) And (Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) > 0) And Control.Visible
    End Select
    
End Sub

Private Function zlRefData() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim intRow As Integer
    Dim lngID As Long
    
    Err = 0
    On Error GoTo ErrHand
    '------------------------------------------------------------------------------------------------------------------
    '护理文件刷新
    
    With vfgFile
        .Clear
        .Rows = 2
        .Cols = 9
        .FixedCols = 1
        
        .TextMatrix(0, mCol.f标志) = ""
        .TextMatrix(0, mCol.fID) = "ID"
        .TextMatrix(0, mCol.f格式ID) = "格式ID"
        .TextMatrix(0, mCol.f文件) = "文件"
        .TextMatrix(0, mCol.f开始日期) = "开始日期"
        .TextMatrix(0, mCol.f科室ID) = "科室id"
        .TextMatrix(0, mCol.f科室) = "科室"
        .TextMatrix(0, mCol.f保留) = "保留"
        .TextMatrix(0, mCol.f创建日期) = "创建日期"
        
        Set .Cell(flexcpPicture, 1, mCol.f标志) = Nothing
        .TextMatrix(1, mCol.fID) = ""
        .TextMatrix(1, mCol.f格式ID) = ""
        .TextMatrix(1, mCol.f文件) = ""
        .TextMatrix(1, mCol.f开始日期) = ""
        .TextMatrix(1, mCol.f科室ID) = ""
        .TextMatrix(1, mCol.f科室) = ""
        .TextMatrix(1, mCol.f保留) = ""
        .TextMatrix(1, mCol.f创建日期) = ""
        
        .ColWidth(mCol.f标志) = 270
        .ColWidth(mCol.fID) = 0: .ColWidth(mCol.f格式ID) = 0: .ColWidth(mCol.f文件) = 2000: .ColWidth(mCol.f开始日期) = 1200
        .ColWidth(mCol.f科室ID) = 0: .ColWidth(mCol.f科室) = 1200: .ColWidth(mCol.f保留) = 0: .ColWidth(mCol.f创建日期) = 0
    End With
    
    intRow = vfgFile.FixedRows
    '--------------------------------------------------------------------------------------------------------------
    gstrSQL = "" & _
        " SELECT A.ID,A.格式ID,A.科室ID,C.名称 AS 科室,A.文件名称,A.开始时间,A.创建时间,B.保留,b.编号" & vbNewLine & _
        " FROM 病人护理文件 A,病历文件列表 B,部门表 C" & vbNewLine & _
        " WHERE A.格式ID=B.ID AND A.科室ID=C.ID And A.病人ID=[1] And A.主页ID=[2] And A.婴儿=[3]" & _
        " ORDER BY B.保留,A.开始时间 "
    Call SQLDIY(gstrSQL)
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiID, mlngPageId, mintBaby)
    
    With Me.vfgFile
        Do While Not rsTemp.EOF
            If Val(.TextMatrix(.Rows - 1, mCol.fID)) > 0 Then .AddItem ""
            If rsTemp!保留 = -1 Then
                Set .Cell(flexcpPicture, .Rows - 1, mCol.f标志) = imgData.ListImages("体温").Picture
            Else
                Set .Cell(flexcpPicture, .Rows - 1, mCol.f标志) = imgData.ListImages("普通").Picture
            End If
            
            lngID = Val(NVL(rsTemp!ID, 0))
            If mlngFormatID > 0 And mlngFormatID = Val(NVL(rsTemp!格式ID)) Then mlngFileID = lngID
            
            If mlngFileID <> 0 And lngID = mlngFileID Then
                intRow = .Rows - 1
            End If
            .TextMatrix(.Rows - 1, mCol.fID) = lngID
            .TextMatrix(.Rows - 1, mCol.f格式ID) = NVL(rsTemp!格式ID, 0)
            .TextMatrix(.Rows - 1, mCol.f文件) = NVL(rsTemp!文件名称)
            .TextMatrix(.Rows - 1, mCol.f开始日期) = Format(NVL(rsTemp!开始时间), "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(.Rows - 1, mCol.f科室ID) = NVL(rsTemp!科室ID)
            .TextMatrix(.Rows - 1, mCol.f科室) = NVL(rsTemp!科室)
            .TextMatrix(.Rows - 1, mCol.f保留) = NVL(rsTemp!保留)
            .TextMatrix(.Rows - 1, mCol.f创建日期) = Format(NVL(rsTemp!创建时间), "yyyy-MM-dd HH:mm:ss")
            
            rsTemp.MoveNext
        Loop
    End With
    
    '选择行
    Call vfgFile.Select(intRow, mCol.fID)
    
    If mblnEdit = True Then
        '41778,刘鹏飞,2012-09-06
        '如果病人老板和新版数据都已经存在，不做任何限制。如果只有老板数据，没有新版。则不能添加文件。
        '婴儿应该和母亲使用同一套系统。
        gstrSQL = " Select 1 序号 From 病人护理记录 Where 病人id = [1] And 主页id = [2] And Rownum < 2" & vbNewLine & _
                "   Union All" & vbNewLine & _
                "   Select 2 序号 From 病人护理文件 Where 病人id = [1] And 主页id = [2] And Rownum < 2"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiID, mlngPageId)
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveLast
            If Val(rsTemp!序号) = 1 Then mblnEdit = False
        End If
    End If
    
    zlRefData = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function InitData(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mstrPrivs = strPrivs
    Set mfrmMain = frmMain
    
    If ExecuteCommand("初始控件") = False Or ExecuteCommand("初始数据") = False Then Exit Function
    Call ExecuteCommand("读注册表")
    Call ExecuteCommand("控件状态")
    
End Function

Public Function RefreshData(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lngDeptID As Long, ByVal blnDoctorStation As Boolean, _
    ByVal blnEdit As Boolean, Optional ByVal lngFileID As Long = 0, Optional ByVal lng序号 As Integer = 0, Optional ByVal intCurveReSize As Integer = 0, _
    Optional blnRefreshFontSize As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：刷新数据
    '参数： lngFileID 为0默认选择第一个文件，不为0则选择此文件。int序号 0为病人本人 其他为婴儿序号
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    
    mblnInit = False
    mlngPatiID = lng病人ID
    mlngPageId = lng主页ID
    mlngDeptId = lngDeptID
    mblnEdit = blnEdit And Not gblnMoved
    mintCurveReSize = intCurveReSize
    mblnDoctorStation = blnDoctorStation
    mblnRefreshFontSize = blnRefreshFontSize
    '文件ID<>0说明是修改或添加了文件 =0表明是切换了病人。
    '修改添加文件将自动定位到变更的文件，却换病人将定位到项目格式的文件上(没有相同格式的文件定位到第一个文件)
    If lngFileID <> 0 Then
        mlngFileID = lngFileID
        mlngFormatID = 0
    End If
    mlng序号 = IIf(lng序号 < 0, 0, lng序号)
    
    Call ExecuteCommand("刷新数据")
End Function

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim strNow As String
    Dim strNote As String
    Dim byt护理等级 As Byte
    Static strPatient As String     '病人ID|主页ID|婴儿
    
    On Error GoTo ErrHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        Call InitCommandBar
'        Set mclsDockAduits = New zlRichEPR.clsDockAduits
'        Call FormSetCaption(mclsDockAduits.zlGetFormTendBody, False, False)
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"

        
    '------------------------------------------------------------------------------------------------------------------
    Case "控件状态"

        
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新状态"
        

        
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新数据"
    
        '判断病人是否已转出
        '因为该函数内外都在调用,参数不好变,直接读取
        '------------------------------------------------------------------------------------------------------------------
        gblnMoved = False
        
        If mlngPatiID <> 0 Then
            '检查病人所有文件是否已经提交到病案室、数据是否转出
            gstrSQL = "Select 数据转出 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
            Set rs = zldatabase.OpenSQLRecord(gstrSQL, "判断数据是否转出", mlngPatiID, mlngPageId)
            gblnMoved = NVL(rs!数据转出, 0) <> 0
        End If
        
        mblnNoRefresh = True
        cboBaby.Clear
        cboBaby.AddItem "病人本人"
        gstrSQL = "Select a.序号,Decode(a.婴儿姓名,Null,NVL(C.姓名,b.姓名) ||'之子'||Trim(To_Char(a.序号,'9')),a.婴儿姓名) As 婴儿姓名" & _
            " From 病人信息 b,病案主页 C,病人新生儿记录 a Where b.病人id=C.病人id And A.病人ID=C.病人ID And A.主页ID=C.主页ID And C.病人id=[1] And C.主页id=[2]  Order By a.序号"
        Set rs = zldatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngPatiID, mlngPageId)
        If rs.BOF = False Then
            Do While Not rs.EOF
                cboBaby.AddItem rs("婴儿姓名").Value: cboBaby.ItemData(cboBaby.NewIndex) = Val(NVL(rs("序号").Value, 0))
                If cboBaby.ListIndex = -1 And Val(NVL(rs("序号").Value, 0)) = mlng序号 Then cboBaby.ListIndex = cboBaby.NewIndex
                rs.MoveNext
            Loop
        End If
        If cboBaby.ListIndex = -1 And cboBaby.ListCount > 0 Then cboBaby.ListIndex = 0
        cboBaby.Enabled = (cboBaby.ListCount > 1)
        
        Call zlRefData
        mblnNoRefresh = False
        Call ExecuteCommand("显示文件内容", vfgFile.ROW)
        
        mblnInit = True
    '------------------------------------------------------------------------------------------------------------------
    Case "显示文件内容"
        'todo:应该传文件ID,但老程序只接受格式ID,需要修改程序
        
        If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) <> 0 Then mlngFileID = Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID))
        If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f格式ID)) <> 0 Then mlngFormatID = Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f格式ID))
        RaiseEvent ShowData(mintBaby, Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)), mlngDeptId, Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) + 1, mintCurveReSize)
        If mblnRefreshFontSize = True And Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) <> 0 Then
            RaiseEvent ViewReSetFontSize(Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) + 1, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
        End If
        mblnRefreshFontSize = True
        If Not mblnDoctorStation And mblnEdit = True Then
            If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) = 0 And InStr(1, mstrPrivs, "护理文件管理") > 0 And strPatient <> mlngPatiID & "|" & mlngPageId & "|" & mintBaby Then
                '病人病案已提交审查,不能添加文件
                If Val(Split(EprIsCommit(mlngPatiID, mlngPageId), "|")(0)) = 0 Then Exit Function
                If frmNurseFileMan.ShowEditor(mlngPatiID, mlngPageId, mintBaby, mstrPrivs, True, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))) Then
                    Call ExecuteCommand("刷新数据")
                End If
            End If
            strPatient = mlngPatiID & "|" & mlngPageId & "|" & mintBaby
        End If
    
    End Select

    ExecuteCommand = True

    GoTo EndHand
    
    '------------------------------------------------------------------------------------------------------------------
ErrHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
EndHand:

End Function

Private Sub InitCommandBar()
    Dim strDLL As String
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    gstrSQL = " Select 新部件 From 体温部件 Where Nvl(启用,0)=1"
    Call zldatabase.OpenRecordset(rsTemp, gstrSQL, "提取体温部件")
    If Err <> 0 Then
        mblnTemparatureChat = True
    Else
        If rsTemp.RecordCount = 0 Then
            mblnTemparatureChat = True
        Else
            If rsTemp!新部件 = "zl9TemperatureChart" Then
                mblnTemparatureChat = True
            Else
                mblnTemparatureChat = False
            End If
        End If
    End If
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.VisualTheme = xtpThemeOffice2003
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = imgPublic.Icons
    cbsMain.ActiveMenuBar.Visible = False
    
End Sub

Private Sub cboBaby_Click()
    mlng序号 = cboBaby.ItemData(cboBaby.ListIndex)
    If mintBaby = mlng序号 Then Exit Sub
    mintBaby = mlng序号
'    mblnRefresh = True
    If mblnNoRefresh Then Exit Sub
    mblnNoRefresh = True
    Call zlRefData
    Call ExecuteCommand("显示文件内容", vfgFile.ROW)
    mblnNoRefresh = False
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngLoop As Long, intNORule As Integer
    Dim DBeginTime As Date
    Dim lngFileID As Long
    Dim blnTrans As Boolean
    Dim ArrSQL()
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    Select Case Control.ID
        Case conMenu_Add
            If Val(Split(EprIsCommit(mlngPatiID, mlngPageId), "|")(0)) = 0 Then
                MsgBox "该病人的病案已提交审查[状态：" & gstrMecState & "]，不能添加文件，请取消审查后再试！", vbInformation, gstrSysName
                Exit Sub
            End If
            If frmNurseFileEdit.ShowEditor(mlngPatiID, mlngPageId, mintBaby, mlngDeptId, "", 0, lngFileID) Then
                mintBaby = -1: mblnNoRefresh = False
                mlngFileID = lngFileID: mlngFormatID = 0
                cboBaby_Click
            End If
        Case conMenu_Modify
            lngFileID = Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID))
            If lngFileID = 0 Then Exit Sub
            
            If Val(Split(EprIsCommit(mlngPatiID, mlngPageId), "|")(2)) = 0 Then
                MsgBox "该病人的病案已提交审查[状态：" & gstrMecState & "]，不能修改文件，请取消审查后再试！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If frmNurseFileEdit.ShowEditor(mlngPatiID, mlngPageId, mintBaby, mlngDeptId, "", lngFileID) Then
                mintBaby = -1: mblnNoRefresh = False
                mlngFileID = lngFileID: mlngFormatID = 0
                cboBaby_Click
            End If
        Case conMenu_Delete
            lngFileID = Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID))
            If lngFileID = 0 Then Exit Sub
            
            If Val(Split(EprIsCommit(mlngPatiID, mlngPageId), "|")(1)) = 0 Then
                MsgBox "该病人的病案已提交审查[状态：" & gstrMecState & "]，不能删除文件，请取消审查后再试！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            '91844,连接护理明细表
            If Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f保留)) = -1 Then
                gstrSQL = "SELECT A.ID,B.开始时间" & _
                    " FROM 病人护理数据 A, 病人护理文件 B,病人护理明细 C" & _
                    " Where  a.文件id = b.Id And b.Id = [1]  And c.记录id = a.Id And Rownum < 2"
            Else
                gstrSQL = "SELECT A.ID,B.开始时间" & _
                    " FROM 病人护理数据 A,病人护理打印 C,病人护理文件 B,病人护理明细 D" & _
                    " WHERE B.ID=[1] And A.文件ID=B.ID and A.文件ID=C.文件ID And A.ID=C.记录ID And d.记录Id=a.Id And RowNum<2"
            End If
            Call SQLDIY(gstrSQL)
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "检查是否存在数据", lngFileID)
            If rsTemp.RecordCount > 0 Then
                MsgBox "该文件已经产生护理数据不允许删除,请检查！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f保留)) = -1 Then
                DBeginTime = CDate(Format(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f开始日期), "YYYY-MM-DD HH:mm:ss"))
                gstrSQL = " Select A.ID,A.开始时间" & _
                    " From 病人护理文件 A,病历文件列表 B" & _
                    " Where A.格式ID=B.ID And A.病人ID=[1] And A.主页ID=[2] And A.婴儿=[3] And B.保留=-1 order by A.开始时间 DESC"
                Call SQLDIY(gstrSQL)
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "判断是否已定义体温单", mlngPatiID, mlngPageId, mintBaby)
                rsTemp.Filter = "开始时间> '" & CStr(DBeginTime) & "'"
                If rsTemp.RecordCount > 0 Then
                    MsgBox "该文件之后还存在其他的体温单文件,文件只能从后往前删除,请检查！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            If MsgBox("你确定要删除" & vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f文件) & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            'If MsgBox("该文件所有的护理数据也将一并删除，请再次确认是否删除！", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            ArrSQL = Array()
            gstrSQL = "ZL_病人护理文件_DELETE(" & lngFileID & ")"
            ReDim Preserve ArrSQL(UBound(ArrSQL) + 1)
            ArrSQL(UBound(ArrSQL)) = gstrSQL
            
            If Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f保留)) = -1 Then
                rsTemp.Filter = "开始时间< '" & CStr(DBeginTime) & "'"
                rsTemp.Sort = "开始时间 DESC"
                If rsTemp.RecordCount > 0 Then
                    '取消上一体温单文件的结束时间
                    gstrSQL = "ZL_病人护理文件_STATE(" & Val(rsTemp!ID) & ",1,NULL)"
                    ReDim Preserve ArrSQL(UBound(ArrSQL) + 1)
                    ArrSQL(UBound(ArrSQL)) = gstrSQL
                End If
            End If
            
            '删除护理记录单时，如果是文件页码顺序编号需要重算该文件之后的文件页码
            '此处不关心文件存在合并的情况，(因为删除已经控制，对于文件如果存在合并信息则不能删除)
            intNORule = zldatabase.GetPara("护理文件页码规则", glngSys, 1255, 0)
            If InStr(1, ",-1,1,", "," & Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f保留)) & ",") = 0 And intNORule <> 0 Then
                
                gstrSQL = " Select id " & vbNewLine & _
                    " From (" & vbNewLine & _
                    "   With 病人护理文件_F1 As" & vbNewLine & _
                    "   (Select a.Id, a.续打id, 开始时间, 创建时间" & vbNewLine & _
                    "   From 病人护理文件 a, 病历文件列表 b" & vbNewLine & _
                    "   Where a.格式id = b.Id And b.种类 = 3 And b.保留 <> 1 And b.保留 <> -1 And a.病人id = [1] And a.主页id = [2] And Nvl(a.婴儿, 0) = [3])" & vbNewLine & _
                    "   Select Id" & vbNewLine & _
                    "   From (Select Id, 开始时间, 创建时间" & vbNewLine & _
                    "       From 病人护理文件_F1 a" & vbNewLine & _
                    "       Where Not Exists (Select 1 From 病人护理文件_F1 Where a.Id = 续打id))" & vbNewLine & _
                    "   Where id<>[4] And (开始时间>[5] OR (开始时间=[5] And 创建时间>[6])) " & vbNewLine & _
                    "   Order by 开始时间)"
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取该文件之后的护理文件", mlngPatiID, mlngPageId, mintBaby, lngFileID, _
                    CDate(Format(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f开始日期), "YYYY-MM-DD HH:mm:ss")), CDate(Format(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.f创建日期), "YYYY-MM-DD HH:mm:ss")))
                If rsTemp.RecordCount > 0 Then
                    gstrSQL = "Zl_病人护理打印_Batchretrypage(" & rsTemp!ID & ",'1;0')"
                    ReDim Preserve ArrSQL(UBound(ArrSQL) + 1)
                    ArrSQL(UBound(ArrSQL)) = gstrSQL
                End If
            End If
            
            If UBound(ArrSQL) > 0 Then gcnOracle.BeginTrans: blnTrans = True
            For lngLoop = 0 To UBound(ArrSQL)
                If CStr(ArrSQL(lngLoop)) <> "" Then Call zldatabase.ExecuteProcedure(CStr(ArrSQL(lngLoop)), "文件删除")
            Next
            If UBound(ArrSQL) > 0 Then gcnOracle.CommitTrans: blnTrans = False
            
            mintBaby = -1: mblnNoRefresh = False
            mlngFileID = lngFileID: mlngFormatID = 0
            cboBaby_Click
    End Select
    
    Exit Sub
ErrHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Add, conMenu_Modify, conMenu_Delete
            Control.Visible = (InStr(1, mstrPrivs, "护理文件管理") > 0 And mblnDoctorStation = False And Not gblnMoved)
            Control.Enabled = (mlngPatiID > 0) And Not mblnArchive And Control.Visible And mblnEdit
            If Control.ID = conMenu_Modify And Control.Enabled = True Then
                Control.Enabled = (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0)
            ElseIf Control.ID = conMenu_Delete And Control.Enabled = True Then
                Control.Enabled = (Val(vfgFile.TextMatrix(Me.vfgFile.ROW, mCol.fID)) <> 0) And (InStr(1, mstrPrivs, "护理文件删除") <> 0)
            End If
    End Select
End Sub

Private Sub Form_Activate()
    RaiseEvent Activate
End Sub

Private Sub Form_Load()
    mblnInit = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picPane.Move 0, 0, Me.Width, Me.Height
    fra.Move 10, 10, Me.Width - 30, fra.Height
    vfgFile.Move 10, fra.Height + 10, Me.Width - 20, Me.Height - vfgFile.Top - 20
End Sub

Private Sub TimFresh_Timer()
    Dim blnFileChange As Boolean
    Dim lngFileID As Long
    Dim lngBaby As Long
    Dim i As Long
    
    If Not mblnInit Then Exit Sub
    If gobjBodyEditor Is Nothing Then Exit Sub
    On Error Resume Next
    Call gobjBodyEditor.zlFileChange(blnFileChange, lngFileID, lngBaby)
    If Err <> 0 Then Err.Clear
    If blnFileChange = False Then Exit Sub
    '根据体温单选择的文件重新定位文件，保持文件列表和体温单选择一致
    If cboBaby.ItemData(cboBaby.ListIndex) = lngBaby Then
        For i = vfgFile.FixedRows To vfgFile.Rows - 1
            If Val(vfgFile.TextMatrix(i, mCol.fID)) = lngFileID And Val(vfgFile.TextMatrix(i, mCol.f保留)) = -1 Then
                Call vfgFile.Select(i, mCol.fID)
                Exit For
            End If
        Next i
    Else
        For i = 0 To cboBaby.ListCount - 1
           If lngBaby = cboBaby.ItemData(i) Then
               mlngFileID = lngFileID: mlngFormatID = 0
               cboBaby.ListIndex = i
               Exit For
           End If
        Next
    End If
End Sub

Private Sub vfgFile_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnNoRefresh Then Exit Sub
    If OldRow <> NewRow Then
        
        Call ExecuteCommand("显示文件内容", NewRow)
'        DoEvents
'        On Error Resume Next
'        vfgFile.SetFocus
    End If
End Sub

Private Sub vfgFile_DblClick()
    Dim lng科室ID As Long
    Dim intEdit As Integer
    
    If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) = 0 Then Exit Sub
    
    lng科室ID = Val(Me.vfgFile.TextMatrix(vfgFile.ROW, mCol.f科室ID))

    If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) = -1 Then
        '体温单查看：病人ID;主页ID;病区ID;出院;编辑;婴儿
        intEdit = 0
        If (InStr(1, ";" & mstrPrivs & ";", ";体温单作图;") > 0 And mblnDoctorStation = False) Then
            If (mblnEdit And mlngPatiID > 0 And mblnArchive = False) Then
                intEdit = 1
            End If
        End If
        If Not CreateBodyEditor Then Exit Sub
        RaiseEvent ViewAnimalHeat(mlngPatiID & ";" & mlngPageId & ";" & mlngDeptId & ";" & Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) & ";0;" & intEdit & ";" & mintBaby & ";1", 0, mstrPrivs, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
    ElseIf Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) = 1 Then
        '产程图查看
        intEdit = 0
        If (InStr(1, ";" & mstrPrivs & ";", ";产程图作图;") > 0 And mblnDoctorStation = False) Then
            If (mblnEdit And mlngPatiID > 0 And mblnArchive = False) Then
                intEdit = 1
            End If
        End If
        If Not CreatePartogram Then Exit Sub
        RaiseEvent Viewpartogram(Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) & ";" & mlngPatiID & ";" & mlngPageId & ";" & mlngDeptId & ";" & intEdit, 1, mstrPrivs, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
        
    Else
        With vfgFile
            RaiseEvent ViewFile(Val(.TextMatrix(.ROW, mCol.fID)), mlngPatiID, mlngPageId, mlngDeptId, mintBaby, False, mstrPrivs, mblnEdit, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
        End With
    End If

End Sub

Private Sub vfgFile_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call vfgFile_DblClick
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-19 15:16
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
    '进行具体模块字体放大缩小功能
    If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) = 0 Then Exit Sub
    RaiseEvent ViewReSetFontSize(Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) + 1, bytSize)
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-19 15:16
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CtlFont  As StdFont
    Dim objCtrl As Control
    Dim bytSize As Byte
    Dim intCol As Integer
    bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
    
    Me.FontSize = mbytFontSize
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Label")
            objCtrl.FontSize = mbytFontSize
            objCtrl.Height = TextHeight("刘") + 20
        Case UCase("ComboBox")
           objCtrl.FontSize = mbytFontSize
        Case UCase("Frame")
           objCtrl.FontSize = mbytFontSize
        Case UCase("VSFlexGrid")
            objCtrl.FontSize = mbytFontSize
            For intCol = 0 To objCtrl.Cols
                Select Case intCol
                    Case mCol.f文件, mCol.f开始日期, mCol.f科室
                        objCtrl.ColWidth(intCol) = BlowUp(CDbl(objCtrl.ColWidth(intCol)))
                End Select
            Next intCol
        End Select
    Next
    fra.Height = cboBaby.Height + 200
    Call Form_Resize
End Sub

Private Function BlowUp(ByRef dblChange As Double) As Double
    '放大：字体，单元格宽度
    BlowUp = dblChange
    If mbytFontSize = 9 Or mbytFontSize = 0 Then Exit Function
    BlowUp = dblChange + (dblChange * 1 / 3)
End Function

Public Sub zlRefreshViewFile()
    If Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.f保留)) <> 1 Then
        Call ExecuteCommand("显示文件内容", vfgFile.ROW)
    End If
End Sub

Public Sub StartTimer(ByVal blnStart As Boolean)
    TimFresh.Enabled = blnStart
End Sub

Private Sub vfgFile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    If Button = 2 Then
        Set cbrPopupBar = cbsMain.Add("右键菜单", xtpBarPopup)
        cbrPopupBar.Title = "右键菜单"
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Add, "新增(&A)"): cbrPopupItem.IconId = 1
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Modify, "修改(&M)"):  cbrPopupItem.IconId = 2
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Delete, "删除(&D)"): cbrPopupItem.IconId = 3
        
        cbrPopupBar.ShowPopup
    End If
End Sub
