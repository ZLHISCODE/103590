VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockInTendFile 
   BorderStyle     =   0  'None
   Caption         =   "护理记录文件"
   ClientHeight    =   6525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   60
      ScaleHeight     =   2715
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   1500
      Width           =   3675
      Begin VB.PictureBox picNote 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         ScaleHeight     =   195
         ScaleWidth      =   3825
         TabIndex        =   2
         Top             =   2190
         Width           =   3825
         Begin VB.Label lblNote 
            Caption         =   "Label1"
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   30
            TabIndex        =   13
            Top             =   0
            Width           =   2175
         End
      End
      Begin XtremeSuiteControls.TabControl tbcFile 
         Height          =   1830
         Left            =   660
         TabIndex        =   1
         Top             =   330
         Width           =   2100
         _Version        =   589884
         _ExtentX        =   3704
         _ExtentY        =   3228
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picRecord 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5835
      Left            =   3930
      ScaleHeight     =   5835
      ScaleWidth      =   7275
      TabIndex        =   3
      Top             =   300
      Width           =   7275
      Begin VB.PictureBox picSplit 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   6435
         TabIndex        =   12
         Top             =   1560
         Width           =   6435
      End
      Begin VB.PictureBox picPane 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   1560
         Index           =   0
         Left            =   0
         ScaleHeight     =   1560
         ScaleWidth      =   6690
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   6690
         Begin VB.Frame fra 
            Height          =   525
            Left            =   0
            TabIndex        =   8
            Top             =   -90
            Width           =   6015
            Begin VB.ComboBox cboBaby 
               Height          =   300
               Left            =   4170
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   150
               Width           =   1350
            End
            Begin VB.Label lblFile 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "护理记录文件:(按规范格式显示查阅护理记录)"
               Height          =   180
               Left            =   60
               TabIndex        =   10
               Top             =   210
               Width           =   3690
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vfgFile 
            Height          =   1095
            Left            =   -15
            TabIndex        =   11
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
         Begin MSComctlLib.ImageList imgData 
            Left            =   3810
            Top             =   4080
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
                  Picture         =   "frmDockInTendFile.frx":0000
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInTendFile.frx":6862
                  Key             =   "体温"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInTendFile.frx":6DFC
                  Key             =   "普通"
               EndProperty
            EndProperty
         End
      End
      Begin VB.PictureBox picPane 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   2700
         Index           =   1
         Left            =   0
         ScaleHeight     =   2700
         ScaleWidth      =   4680
         TabIndex        =   5
         Top             =   2505
         Width           =   4680
         Begin XtremeSuiteControls.TabControl tbcSub 
            Height          =   2490
            Left            =   120
            TabIndex        =   6
            Top             =   150
            Width           =   3450
            _Version        =   589884
            _ExtentX        =   6085
            _ExtentY        =   4392
            _StockProps     =   64
         End
      End
      Begin VB.PictureBox picPane 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   1455
         Index           =   2
         Left            =   5310
         ScaleHeight     =   1455
         ScaleWidth      =   1410
         TabIndex        =   4
         Top             =   3330
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmDockInTendFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnMouseMove As Boolean

'######################################################################################################################

Private Enum mCol
    r标志 = 0: rID: r发生时间: r记录项目: r记录数据: r分组: r护士: r登记时间: r病区ID: r病区名:: r签名人: r签名时间: r项目序号: r开始版本: r未记说明
    f标志 = 0: fID: f编号: f文件: f日期范围: f病区id: f病区名: f护理级别: f文件级别: f保留
    w标志 = 0: wID: w页面编号: w页面名称: w病历名称: w创建人: w创建时间: w保存人: w完成时间: w当前版本: w签名级别: w当前情况: w归档人: w归档日期: w病区ID: w病区名: w病人状态
End Enum

Private Enum mColWidth
        c标志 = 270: cID = 0: c编号 = 600: c文件 = 2000: c日期范围 = 3500: c病区id = 0: c病区名 = 1200: c护理级别 = 810: c文件级别 = 0: c保留 = 0
End Enum

Private WithEvents mclsDockAduits As zlRichEPR.clsDockAduits
Attribute mclsDockAduits.VB_VarHelpID = -1
Private mfrmCaseTendEditForBatch As frmCaseTendEditForBatch
Private mblnNoRefresh As Boolean
Private mstrPrivs As String                             '当前使用者对本程序(1255)的权限串
Private mlngPatiId As Long                              '病人id
Private mlngPageId As Long                              '主页id
Private mlngDeptId As Long                              '当前操作科室id，如病人科室和当前科室不一致，则不能操作归档外的功能
Private mblnEdit As Boolean                             '是否允许操作，通常由上级程序根据当前操作科室是否当前病人病区决定。
Private mblnDoctorStation As Boolean
Private mbytFontSize As Byte                            '字体显示大小0-9号字体,1-12号字体
Private mblnRefreshFontSize As Boolean                  '记录是否刷新字体信息
Private rsTemp As New ADODB.Recordset
Private mintBaby As Integer
Private mfrmMain As Object
Private mblnTendArchive As Boolean

Public Event AfterDataChanged()
Public Event Activate()

Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

''######################################################################################################################

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-18 15:16
    '问题:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-20 15:15:00
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytSize As Byte
    Dim PATI_COLWIDTH As Variant
    Dim lngCol As Long, lngReDraw As Long
    bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
    
    Me.FontSize = mbytFontSize
    Me.FontName = "宋体"
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Label")
                objCtrl.FontSize = mbytFontSize
                objCtrl.Height = TextHeight("刘") + 20
        Case UCase("VsFlexGrid")
            objCtrl.FontSize = mbytFontSize
        Case UCase("ComboBox")
            objCtrl.FontSize = mbytFontSize
        Case UCase("TabControl")
            Set CtlFont = objCtrl.PaintManager.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.Font = CtlFont
            objCtrl.PaintManager.Layout = xtpTabLayoutAutoSize
        End Select
    Next
    '调整控件位置
    PATI_COLWIDTH = Array(c标志, cID, c编号, c文件, c日期范围, c病区id, c病区名, c护理级别, c文件级别, c保留)
    With vfgFile
        lngReDraw = .Redraw
        .Redraw = flexRDNone
        For lngCol = fID To .Cols - 1
            .ColWidth(lngCol) = BlowUp(CDbl(PATI_COLWIDTH(lngCol)))
        Next lngCol
        .Redraw = lngReDraw
    End With
    
    lblNote.Top = 0: lblNote.Left = 30
    picNote.Height = lblNote.Height
    lblFile.Left = 60
    cboBaby.Top = 150
    cboBaby.Width = BlowUp(1350)
    lblFile.Top = cboBaby.Top + (cboBaby.Height - lblFile.Height) \ 2
    fra.Height = cboBaby.Top + cboBaby.Height + 75
    picSplit.Top = vfgFile.Rows * vfgFile.RowHeightMin + vfgFile.Top + 100
    Call Form_Resize
    
    '刷新字体
    Call ExecuteCommand("设置字体")
End Sub

Private Function BlowUp(ByRef dblChange As Double) As Double
    '放大：字体，单元格宽度
    BlowUp = dblChange + (dblChange * IIf(mbytFontSize = 12, 1, 0) / 3)
End Function

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strInfo As String
    Dim byt护理等级 As Byte
    Dim objFrmBody As Object
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Open

        With vfgFile
        
            strInfo = Val(.TextMatrix(.Row, mCol.f病区id))
            
            If Val(.TextMatrix(.Row, mCol.f保留)) = -1 Then
                '体温单查看：病人ID;主页ID;病区ID;出院;编辑;婴儿
                If Not CreateBodyEditor Then Exit Sub
                Set objFrmBody = gobjBodyEditor.GetTendBody
                On Error Resume Next
                objFrmBody.Resize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
                If Err <> 0 Then Err.Clear
                On Error GoTo errHand
                Call gobjBodyEditor.GetTendBody.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";0;0;" & mintBaby, 1, mstrPrivs)
            Else
                                    
                Call frmTendFileOpen.ShowMe(Me, Val(.TextMatrix(.Row, mCol.fID)), mlngPatiId, mlngPageId, Val(strInfo), mintBaby, .TextMatrix(.Row, mCol.f日期范围), , Val(.TextMatrix(.Row, mCol.f护理级别)), mblnMoved_HL, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
                
            End If
        End With

        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview
        
        ''1-预览,2-打印
        
        With vfgFile

            If Val(.TextMatrix(.Row, mCol.f保留)) = -1 Then
                
                Call mclsDockAduits.zlPrintDocument(1, 1)

            ElseIf .TextMatrix(.Row, mCol.f文件) <> "文件" Then
                
                Call mclsDockAduits.zlPrintDocument(2, 1)
                
            End If
            
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print
            
        With vfgFile

            If Val(.TextMatrix(.Row, mCol.f保留)) = -1 Then
                
                Call mclsDockAduits.zlPrintDocument(1, 2)

            ElseIf .TextMatrix(.Row, mCol.f文件) <> "文件" Then
                
                Call mclsDockAduits.zlPrintDocument(2, 2)
                
            End If
            
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel
            
        With vfgFile

            If Val(.TextMatrix(.Row, mCol.f保留)) = -1 Then
                
                ShowSimpleMsg "对不起，体温单不支持输出到Excel！"

            ElseIf .TextMatrix(.Row, mCol.f文件) <> "文件" Then
                
                Call mclsDockAduits.zlPrintDocument(2, 3)
                
            End If
            
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MarkMap
        
        '体温作图：病人ID;主页ID;病区ID;出院;编辑;婴儿
        If Not CreateBodyEditor Then Exit Sub
        Set objFrmBody = gobjBodyEditor.GetTendBody
        On Error Resume Next
        objFrmBody.Resize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
        If Err <> 0 Then Err.Clear
        On Error GoTo errHand
        If objFrmBody.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";0;1;" & mintBaby, 2, mstrPrivs) Then
            
            Call ExecuteCommand("刷新数据")

            RaiseEvent AfterDataChanged

        End If
        
    Case conMenu_File_PrintDayDetail        '批量录入
        If mfrmCaseTendEditForBatch Is Nothing Then Set mfrmCaseTendEditForBatch = New frmCaseTendEditForBatch
        Call mfrmCaseTendEditForBatch.ShowMe(Me, mlngDeptId, mstrPrivs)
    Case conMenu_Tool_Sign
        Call mclsDockAduits.zlGetFormTendEdit.SignMe
        RaiseEvent AfterDataChanged
    Case conMenu_Tool_SignEarse
        Call mclsDockAduits.zlGetFormTendEdit.UnSignMe
        RaiseEvent AfterDataChanged
    Case conMenu_Edit_Archive * 10
        Call mclsDockAduits.zlGetFormTendEdit.ArchiveMe
        RaiseEvent AfterDataChanged
    Case conMenu_Edit_UnArchive
        Call mclsDockAduits.zlGetFormTendEdit.UnArchiveMe
        RaiseEvent AfterDataChanged
    Case conMenu_Tool_SignVerify
        Call mclsDockAduits.SignMarker
    Case conMenu_Edit_Save
        If mclsDockAduits.zlGetFormTendEdit.SaveME Then RaiseEvent AfterDataChanged
    Case conMenu_Edit_Transf_Cancle
        Call mclsDockAduits.CancelMe
    End Select
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Sub

Public Property Get TendArchive() As Boolean
    TendArchive = mblnTendArchive
End Property

Public Property Let TendArchive(ByVal vData As Boolean)
    mblnTendArchive = vData
End Property

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Open
        Control.Visible = True
        Control.Enabled = tbcFile.Item(1).Selected And (Val(vfgFile.TextMatrix(Me.vfgFile.Row, mCol.fID)) <> 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print
        Control.Enabled = tbcFile.Item(1).Selected And (Val(vfgFile.TextMatrix(Me.vfgFile.Row, mCol.fID)) <> 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel
        Control.Enabled = tbcFile.Item(1).Selected And (vfgFile.Rows > 1 And Val(vfgFile.TextMatrix(vfgFile.Row, mCol.f保留)) <> -1)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_EditPopup
        Control.Visible = (mblnDoctorStation = False And InStr(1, mstrPrivs, "体温单作图") > 0)
        Control.Enabled = Control.Visible
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MarkMap
        Control.Visible = (InStr(1, mstrPrivs, "体温单作图") > 0 And mblnDoctorStation = False)
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And Control.Visible And TendArchive = False And Not mblnMoved_HL) 'And Val(vfgFile.TextMatrix(vfgFile.Row, mCol.f保留)) = -1
    Case conMenu_File_PrintDayDetail
        Control.Enabled = (mblnEdit And mlngPatiId > 0)  'And (Not mclsDockAduits.zlIsPigeonhole))

        Control.Visible = (InStr(1, mstrPrivs, "护理记录登记") > 0 And mblnDoctorStation = False And Not mblnMoved_HL)
        If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "护理记录登记") > 0)
    Case conMenu_Tool_Sign  '签名
        Control.Visible = Not mblnDoctorStation
        Control.Enabled = (Not mclsDockAduits.zlIsCert) And (mlngPatiId > 0) And (Not mclsDockAduits.zlDataChange) And (Not mblnDoctorStation) And Not mblnMoved_HL And mblnEdit
    Case conMenu_Tool_SignEarse  '取消签名
        Control.Visible = Not mblnDoctorStation
        Control.Enabled = mclsDockAduits.zlIsCert And (mlngPatiId > 0) And (Not mclsDockAduits.zlDataChange) And (Not mblnDoctorStation) And Not mblnMoved_HL And mblnEdit
    Case conMenu_Edit_Archive * 10 '归档
        Control.Visible = Not mblnDoctorStation And mblnTendArchive = False
        Control.Enabled = (Not mclsDockAduits.zlIsPigeonhole) And (mlngPatiId > 0) And (Not mclsDockAduits.zlDataChange) And (Not mblnDoctorStation) And Not mblnMoved_HL And mblnEdit
    Case conMenu_Edit_UnArchive  '取消归档
        Control.Visible = Not mblnDoctorStation And mblnTendArchive
        Control.Enabled = mclsDockAduits.zlIsPigeonhole And (mlngPatiId > 0) And (Not mclsDockAduits.zlDataChange) And (Not mblnDoctorStation) And Not mblnMoved_HL And mblnEdit
    Case conMenu_Edit_Save  '保存
        Control.Visible = Not mblnDoctorStation
        Control.Enabled = (mclsDockAduits.zlDataChange) And (Not mblnDoctorStation) And Not mblnMoved_HL
    Case conMenu_Edit_Transf_Cancle  '取消
        Control.Visible = Not mblnDoctorStation
        Control.Enabled = (mclsDockAduits.zlDataChange) And (Not mblnDoctorStation)
    Case conMenu_Tool_SignVerify
        Control.Visible = (tbcFile.Selected.Index = 0)
        Control.Enabled = Control.Visible
    End Select
    
End Sub

Private Function zlRefData() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngCol As Long, lngRow As Long
    Dim rsMain As New ADODB.Recordset
    Dim rs As New ADODB.Recordset

    Dim strSvrKey As String
    Dim int保留 As Integer
    Dim strCode As String
    Dim strFile As String
    Dim strStart As String
    Dim strEnd As String
    Dim lng科室ID As Long
    Dim str科室 As String
    Dim str护理级别 As String
    Dim bln护理级别 As Boolean
    Dim bln一份文件 As Boolean
    Dim strTmp As String
    '按护理级别分别显示时,科室护理等级及护理文件ID未发生变化时,只显示一份文件
    Dim strStart_Cur As String
    Dim strEnd_Cur As String
    Dim strHLDate_Cur As String
    Dim strFile_Cur As String
    Dim str科室ID_Cur As String
    Dim str护理级别_Cur As String
    Dim str编号_CUR As String
    Dim str名称_CUR As String
    Dim str科室_CUR As String
    Dim str报表_CUR As String
    Dim blnExit As Boolean
    
    Err = 0
    On Error GoTo errHand
    '------------------------------------------------------------------------------------------------------------------
    '护理文件刷新
    
    bln护理级别 = (Val(zlDatabase.GetPara("按护理级别分组", glngSys, 1255, "0")) = 1)
    bln一份文件 = (Val(zlDatabase.GetPara("显示一份护理文件", glngSys, 1255, "1")) = 1)

    If bln护理级别 Then
        '--------------------------------------------------------------------------------------------------------------
        With vfgFile
            .Rows = 2
            .Cols = 10
            .FixedCols = 1
            
            .TextMatrix(0, 0) = ""
            .TextMatrix(0, 1) = "ID"
            .TextMatrix(0, 2) = "编号"
            .TextMatrix(0, 3) = "文件"
            .TextMatrix(0, 4) = "日期范围"
            .TextMatrix(0, 5) = "科室id"
            .TextMatrix(0, 6) = "科室"
            .TextMatrix(0, 7) = "护理级别"
            .TextMatrix(0, 8) = "文件级别"
            .TextMatrix(0, 9) = "保留"
            
            Set .Cell(flexcpPicture, 1, mCol.f标志) = Nothing
            .TextMatrix(1, mCol.fID) = ""
            .TextMatrix(1, mCol.f编号) = ""
            .TextMatrix(1, mCol.f文件) = ""
            .TextMatrix(1, mCol.f日期范围) = ""
            .TextMatrix(1, mCol.f病区id) = ""
            .TextMatrix(1, mCol.f病区名) = ""
            .TextMatrix(1, mCol.f护理级别) = ""
            .TextMatrix(1, mCol.f文件级别) = ""
            .TextMatrix(1, mCol.f保留) = ""
            
            .ColWidth(mCol.f标志) = mColWidth.c标志
            .ColWidth(mCol.fID) = mColWidth.cID: .ColWidth(mCol.f编号) = mColWidth.c编号: .ColWidth(mCol.f文件) = mColWidth.c文件: .ColWidth(mCol.f日期范围) = mColWidth.c日期范围
            .ColWidth(mCol.f病区id) = mColWidth.c病区id: .ColWidth(mCol.f病区名) = mColWidth.c病区名: .ColWidth(mCol.f护理级别) = mColWidth.c护理级别: .ColWidth(mCol.f保留) = mColWidth.c保留: .ColWidth(mCol.f文件级别) = mColWidth.c文件级别
    
        End With
        
        gstrSQL = "Select a.Id, a.编号, a.名称 As 文件," & _
                "        To_Char(a.开始, 'yyyy-mm-dd hh24:mi') || ' ～ ' || To_Char(a.截止, 'yyyy-mm-dd hh24:mi') As 日期范围," & _
                "        a.科室id, b.名称 As 科室, 3 As 护理级别,保留" & _
                " From (" & _
                "        Select f.Id, f.编号, f.名称, r.开始, r.截止, r.科室id, 保留" & _
                "        From ( Select Id, 编号, 名称, 3 As 护理级别, 通用, 0 As 科室id,保留 From 病历文件列表 Where 种类=3 And 保留<0 And NVL(子类,0)=0) f," & _
                "             (Select r.科室id, Nvl(Min(r.护理级别),3) As 护理级别, Min(r.发生时间) As 开始, Max(r.发生时间) As 截止" & _
                "               From 病人护理记录 r" & _
                "               Where r.病人来源 = 2 And r.病人ID = [1] And NVL(r.主页ID, 0) = [2] And Nvl(r.婴儿,0)=[3] " & _
                "               Group By r.科室id) r" & _
                "        Where f.保留<0  And f.护理级别 >= r.护理级别) a, 部门表 b" & _
                " Where a.科室ID = b.ID " & _
                " Order By a.编号, To_Char(a.开始, 'yyyy-mm-dd hh24:mi') || ' ～ ' || To_Char(a.截止, 'yyyy-mm-dd hh24:mi')"
        If mblnMoved_HL Then
            gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
            gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
        End If
        Set rsMain = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, mintBaby)
        
            
        If rsMain.BOF = False Then
            
            With vfgFile
                If Val(.TextMatrix(.Rows - 1, mCol.fID)) > 0 Then .AddItem ""
    
                Set .Cell(flexcpPicture, .Rows - 1, mCol.f标志) = imgData.ListImages("体温").Picture
    
                .TextMatrix(.Rows - 1, mCol.fID) = rsMain("ID").Value
                .TextMatrix(.Rows - 1, mCol.f编号) = rsMain("编号").Value
                .TextMatrix(.Rows - 1, mCol.f文件) = rsMain("文件").Value
                .TextMatrix(.Rows - 1, mCol.f日期范围) = rsMain("日期范围").Value
                .TextMatrix(.Rows - 1, mCol.f病区id) = rsMain("科室id").Value
                .TextMatrix(.Rows - 1, mCol.f病区名) = rsMain("科室").Value
                .TextMatrix(.Rows - 1, mCol.f护理级别) = "/"
                .TextMatrix(.Rows - 1, mCol.f文件级别) = "/"
                .TextMatrix(.Rows - 1, mCol.f保留) = -1
                
            End With
        End If
        
        '1.求时段
        gstrSQL = _
            "Select a.科室id,a.病区id, b.护理级别 As 护理级别,d.名称 As 科室, Min(a.开始时间) As 开始时间, Max(Nvl(a.终止时间,Sysdate+100)) As 终止时间" & vbNewLine & _
            "From 病人变动记录 a," & vbNewLine & _
            "        (Select Id, 护理等级,Decode(特级, 'Y', 0, Decode(一级, 'Y', 1, Decode(二级, 'Y', 2, 3))) As 护理级别" & vbNewLine & _
            "            From (Select b.Id,b.名称 As 护理等级, Decode(Sign(Instr(b.名称, '特')), 1, 'Y', Decode(Sign(Instr(b.名称, '重')), 1, 'Y', 'N')) As 特级," & vbNewLine & _
            "                                        Decode(Sign(Instr(b.名称, '一')), 1, 'Y'," & vbNewLine & _
            "                                                        Decode(Sign(Instr(b.名称, '1')), 1, 'Y'," & vbNewLine & _
            "                                                                        Decode(Sign(Instr(b.名称, 'Ⅰ')), 1, 'Y', Decode(Sign(Instr(b.名称, 'I')), 1, 'Y', 'N')))) As 一级," & vbNewLine & _
            "                                        Decode(Sign(Instr(b.名称, '二')), 1, 'Y'," & vbNewLine & _
            "                                                        Decode(Sign(Instr(b.名称, '2')), 1, 'Y'," & vbNewLine & _
            "                                                                        Decode(Sign(Instr(b.名称, 'Ⅱ')), 1, 'Y', Decode(Sign(Instr(b.名称, 'II')), 1, 'Y', 'N')))) As 二级" & vbNewLine & _
            "                           From 收费项目目录 b" & vbNewLine & _
            "                           Where b.类别 = 'H')) b,部门表 d" & vbNewLine & _
            "Where a.病人id = [1] And a.主页id = [2] And b.Id = a.护理等级id  And d.Id = a.科室id" & vbNewLine & _
            "Group By a.科室id,a.病区id, b.护理级别,d.名称 "
        gstrSQL = " Select * From (" & gstrSQL & ") Order by 开始时间"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)
        If rs.EOF = False Then
            Do While Not rs.EOF
                '2.求指定的最大护理级别的护理文件,只取第一个(参数：病区id,护理级别)
                gstrSQL = _
                    "Select l.Id, l.编号, l.名称, l.保留,a.科室id,f.报表" & vbNewLine & _
                    "From 病历文件列表 l, 病历页面格式 f, 病历应用科室 a" & vbNewLine & _
                    "Where l.种类 = 3 And l.保留 = 0 And l.种类 = f.种类 And l.编号 = f.编号 And l.Id = a.文件id(+) And" & vbNewLine & _
                    "           (l.保留 < 0 Or l.通用 = 1 Or l.通用 = 2 And a.科室id = [1]) And f.报表 >= [2]" & vbNewLine & _
                    "Order By f.报表"
                
                If IsNull(rs("病区id").Value) = False Then
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(rs("病区id").Value), Val(rs("护理级别").Value))
                Else
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(rs("科室id").Value), Val(rs("护理级别").Value))
                End If
                
                strStart_Cur = ""
                Do While Not rsTemp.EOF
                    
                    '只取第一个
                    int保留 = rsTemp("保留").Value
                    strCode = rsTemp("编号").Value
                    strFile = rsTemp("名称").Value
                    lng科室ID = rs("科室id").Value
                    str科室 = rs("科室").Value
                    str护理级别 = rs!护理级别
                    
                    strStart = Format(rs("开始时间").Value, "yyyy-MM-dd HH:mm:ss")
                    strEnd = Format(rs("终止时间").Value, "yyyy-MM-dd HH:mm:ss")
                    
                    If bln一份文件 Then
                        Call ShowFileOnly(mlngPatiId, mlngPageId, mintBaby, strStart, strEnd, lng科室ID, rs!护理级别, rsTemp!ID, strCode, strFile, str科室, int保留, Val(rsTemp("报表").Value), rs.AbsolutePosition = 1)
                        Exit Do
                    Else
                        '按护理级别分别显示时,科室护理等级及护理文件ID未发生变化时,只显示一份文件
                        If strStart_Cur = "" Then
                            strStart_Cur = strStart
                            strEnd_Cur = strEnd
                            strFile_Cur = rsTemp!ID
                            str科室ID_Cur = rs!科室ID
                            str护理级别_Cur = rs!护理级别
                            
                            str编号_CUR = strCode
                            str名称_CUR = strFile
                            str科室_CUR = str科室
                            str报表_CUR = Val(rsTemp("报表").Value)
                        End If
                        
                        If str科室ID_Cur <> rs!科室ID Or Val(str护理级别_Cur) <> rs!护理级别 Or strFile_Cur <> rsTemp!ID Then
                            Call ShowFile(mlngPatiId, mlngPageId, mintBaby, strStart_Cur, strEnd_Cur, str科室ID_Cur, str护理级别_Cur, strFile_Cur, str编号_CUR, str名称_CUR, str科室_CUR, int保留, Val(str报表_CUR), True)
                            
                            strStart_Cur = strStart
                            strFile_Cur = rsTemp!ID
                            str科室ID_Cur = rs!科室ID
                            str护理级别_Cur = rs!护理级别
                            
                            str编号_CUR = strCode
                            str名称_CUR = strFile
                            str科室_CUR = str科室
                            str报表_CUR = Val(rsTemp("报表").Value)
                        End If
                        strEnd_Cur = strEnd
                    End If
                    
                    rsTemp.MoveNext
                Loop
                '最后一条记录必须要添加
                If Not bln一份文件 Then
                    Call ShowFile(mlngPatiId, mlngPageId, mintBaby, strStart_Cur, strEnd_Cur, str科室ID_Cur, str护理级别_Cur, strFile_Cur, str编号_CUR, str名称_CUR, str科室_CUR, int保留, Val(str报表_CUR), True)
                End If
                
                rs.MoveNext
            Loop
        End If
        
    Else
        '--------------------------------------------------------------------------------------------------------------
        
        With vfgFile
            .Rows = 2
            .Cols = 10
            .FixedCols = 1
            
            .TextMatrix(0, 0) = ""
            .TextMatrix(0, 1) = "ID"
            .TextMatrix(0, 2) = "编号"
            .TextMatrix(0, 3) = "文件"
            .TextMatrix(0, 4) = "日期范围"
            .TextMatrix(0, 5) = "科室id"
            .TextMatrix(0, 6) = "科室"
            .TextMatrix(0, 7) = "护理级别"
            .TextMatrix(0, 8) = "文件级别"
            .TextMatrix(0, 9) = "保留"
            
            Set .Cell(flexcpPicture, 1, mCol.f标志) = Nothing
            .TextMatrix(1, mCol.fID) = ""
            .TextMatrix(1, mCol.f编号) = ""
            .TextMatrix(1, mCol.f文件) = ""
            .TextMatrix(1, mCol.f日期范围) = ""
            .TextMatrix(1, mCol.f病区id) = ""
            .TextMatrix(1, mCol.f病区名) = ""
            .TextMatrix(1, mCol.f护理级别) = ""
            .TextMatrix(1, mCol.f文件级别) = ""
            .TextMatrix(1, mCol.f保留) = ""
            
            .ColWidth(mCol.f标志) = mColWidth.c标志
            .ColWidth(mCol.fID) = mColWidth.cID: .ColWidth(mCol.f编号) = mColWidth.c编号: .ColWidth(mCol.f文件) = mColWidth.c文件: .ColWidth(mCol.f日期范围) = mColWidth.c日期范围
            .ColWidth(mCol.f病区id) = mColWidth.c病区id: .ColWidth(mCol.f病区名) = mColWidth.c病区名: .ColWidth(mCol.f护理级别) = mColWidth.c护理级别: .ColWidth(mCol.f保留) = mColWidth.c保留: .ColWidth(mCol.f文件级别) = mColWidth.c文件级别
    
        End With
        
        '--------------------------------------------------------------------------------------------------------------
        gstrSQL = "Select distinct a.Id, a.编号, a.名称 As 文件," & _
                "        a.开始,a.截止," & _
                "        a.科室id, b.名称 As 科室, 0 As 护理级别,a.文件级别,保留" & _
                " From (" & _
                "        Select f.Id, f.编号, f.名称, r.开始, r.截止, r.科室id, 保留,文件级别 " & _
                "        From ( Select Id, 编号, 名称, 3 As 文件级别, 通用, 0 As 科室id,保留 From 病历文件列表 Where 种类=3 And 保留<0 And NVL(子类,0)=0" & _
                "               Union All " & _
                "               Select l.Id, l.编号, l.名称, f.报表 As 文件级别, l.通用, a.科室id,l.保留 " & _
                "               From 病历文件列表 l, 病历页面格式 f, 病历应用科室 a" & _
                "               Where l.种类 = 3 And l.保留 = 0 And l.种类 = f.种类 And l.编号 = f.编号 And l.Id = a.文件id(+)) f," & _
                "             (Select r.科室id, Nvl(Min(r.护理级别),3) As 护理级别, Min(r.发生时间) As 开始, Max(r.发生时间) As 截止" & _
                "               From 病人护理记录 r" & _
                "               Where r.病人来源 = 2 And r.病人ID = [1] And NVL(r.主页ID, 0) = [2] And Nvl(r.婴儿,0)=[3] " & _
                "               Group By r.科室id) r" & _
                "        Where (f.保留<0 Or f.通用 = 1 Or f.通用 = 2 And r.科室id In (Select t.科室id From 病区科室对应 t Where t.病区id=f.科室id)) And f.文件级别 >= r.护理级别) a, 部门表 b" & _
                " Where a.科室ID = b.ID " & _
                " Order By a.保留,A.文件级别,A.编号 desc, To_Char(a.开始, 'yyyy-mm-dd hh24:mi') || ' ～ ' || To_Char(a.截止, 'yyyy-mm-dd hh24:mi')"
        If mblnMoved_HL Then
            gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
            gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
        End If
        Set rsMain = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, mintBaby)
        
        '需要先行进行数据处理(可能出现内一科A文件,内二科A文件,内一科B文件,内二科B文件,但选择只显示一份文件时,后两条记录不应该显示出来)
        Dim rsData As New ADODB.Recordset
        Set rsData = DataProcess(rsMain, bln一份文件)
        
        With Me.vfgFile
            If rsData.RecordCount <> 0 Then rsData.MoveFirst
            Do While Not rsData.EOF
                If rsData!删除 = 0 Then
                    int保留 = rsData("保留").Value
                    strCode = rsData("编号").Value
                    strFile = rsData("名称").Value
                    lng科室ID = rsData("科室id").Value
                    str科室 = rsData("科室名称").Value
    
                    strStart = Format(rsData("开始时间").Value, "yyyy-MM-dd HH:mm:ss")
                    strEnd = Format(rsData("结束时间").Value, "yyyy-MM-dd HH:mm:ss")
                    
                    Call ShowFile(mlngPatiId, mlngPageId, mintBaby, strStart, strEnd, lng科室ID, 0, Val(rsData("文件ID").Value), strCode, strFile, str科室, int保留, Val(rsData("文件级别").Value), (rsData.AbsolutePosition = 1))
                End If
                
                rsData.MoveNext
            Loop

            For lngRow = .FixedRows To .Rows - 1
                If Val(.TextMatrix(lngRow, mCol.f保留)) = -1 Then
                    Set .Cell(flexcpPicture, lngRow, mCol.f标志) = Me.imgData.ListImages("体温").Picture
                Else
                    Set .Cell(flexcpPicture, lngRow, mCol.f标志) = Me.imgData.ListImages("普通").Picture
                End If
            Next
        End With
    End If
    
    If mblnEdit = True Then
        '41778,刘鹏飞,2012-09-06
        '如果病人老板和新版数据都已经存在，不做任何限制。如果只有新板数据，没有老版。则老板不能添加文件。
        '婴儿应该和母亲使用同一套系统。
        gstrSQL = "Select 1 From 病人护理文件 A Where a.病人id = [1] And a.主页id = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)
        If rsTemp.RecordCount > 0 And Val(vfgFile.TextMatrix(vfgFile.Rows - 1, mCol.fID)) = 0 Then
            mblnEdit = False
        End If
    End If
    
    zlRefData = True

    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DataProcess(ByVal rsMain As ADODB.Recordset, ByVal bln一份文件 As Boolean) As ADODB.Recordset
    Dim blnAdd As Boolean           '未分娩则不显示分娩后的护理记录单
    Dim arrFormat, intFormat As Integer
    Dim strField As String, strValue As String, str开始 As String, str截止 As String
    Dim intLocal As Integer '当前指针位置
    Dim intCount As Integer
        Dim intRecords As Integer
    Dim int护理级别 As Integer, lng科室ID As Long, int保留 As Integer
    Dim rsData As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '需要先行进行数据处理(可能出现内一科A文件,内二科A文件,内一科B文件,内二科B文件,但选择只显示一份文件时,后两条记录不应该显示出来)
        
    strField = "ID," & adDouble & ",5|保留," & adDouble & ",18|编号," & adLongVarChar & ",50|名称," & adLongVarChar & ",200|" & _
               "科室ID," & adDouble & ",18|科室名称," & adLongVarChar & ",200|开始时间," & adLongVarChar & ",20|" & _
               "结束时间," & adLongVarChar & ",20|文件ID," & adDouble & ",18|文件级别," & adDouble & ",18|删除," & adDouble & ",1"
    Set rsData = New ADODB.Recordset
    Call Record_Init(rsData, strField)
    
    strField = "ID|保留|编号|名称|科室ID|科室名称|开始时间|结束时间|文件ID|文件级别|删除"
    If rsMain.RecordCount <> 0 Then rsMain.MoveFirst
    Do While Not rsMain.EOF
        str开始 = Format(rsMain("开始").Value, "yyyy-MM-dd HH:mm:ss")
        str截止 = Format(rsMain("截止").Value, "yyyy-MM-dd HH:mm:ss")
        blnAdd = True
        
        '检查产科护理记录单
        If rsMain!保留 <> -1 Then
            gstrSQL = " Select 格式 From 病历页面格式 Where 种类=3 And 编号=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取文件格式", CStr(rsMain!编号))
            intFormat = 0
            arrFormat = Split(NVL(rsTemp!格式, ";;;;;;;;"), ";")
            If UBound(arrFormat) >= 8 Then intFormat = Val(arrFormat(8))
            
            If intFormat <> 0 Then
                '1-分娩前;2-分娩后
                gstrSQL = " Select MAX(出生时间) AS 出生时间 From 病人新生儿记录 Where 病人ID=[1] And 主页ID=[2] "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取新生儿出生时间", mlngPatiId, mlngPageId)
                If Not IsNull(rsTemp!出生时间) Then
                    If intFormat = 1 Then
                        str截止 = Format(rsTemp!出生时间, "yyyy-MM-dd HH:mm:ss")
                    Else
                        str开始 = Format(rsTemp!出生时间, "yyyy-MM-dd HH:mm:ss")
                    End If
                Else
                    blnAdd = (intFormat = 1)
                End If
            End If
        End If
        
        If blnAdd Then
                        intRecords = intRecords + 1
            strValue = intRecords & "|" & rsMain!保留 & "|" & rsMain!编号 & "|" & rsMain!文件 & "|" & rsMain!科室ID & "|" & _
                    rsMain!科室 & "|" & str开始 & "|" & str截止 & "|" & rsMain!ID & "|" & Val(rsMain!文件级别) & "|0"
            'Debug.Print strValue
            Call Record_Update(rsData, strField, strValue, "ID|" & intRecords)
        End If
        rsMain.MoveNext
    Loop
    
    If Not bln一份文件 Then
        Set DataProcess = rsData
        Exit Function
    End If
    
    '依次循环检查,之后存在护理等级\科室ID相同的,把记录删掉
    intCount = rsData.RecordCount
    If intCount > 0 Then
        For intLocal = 1 To intCount
            rsData.MoveFirst
            rsData.Move intLocal - 1
            
            If rsData!删除 = 0 Then
                int保留 = rsData!保留
                int护理级别 = rsData!文件级别
                lng科室ID = rsData!科室ID
                
                rsData.MoveFirst
                Do While Not rsData.EOF
                    If rsData.AbsolutePosition <> intLocal Then
                        If rsData!删除 = 0 And rsData!文件级别 = int护理级别 And rsData!科室ID = lng科室ID And int保留 = rsData!保留 Then
                            Call Record_Update(rsData, "删除", 1, "ID|" & rsData.AbsolutePosition)
                        End If
                    End If
                    rsData.MoveNext
                Loop
            End If
        Next
    End If
    Set DataProcess = rsData
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowFile(ByVal lngPatiID As Long, _
                        ByVal lngPageId As Long, _
                        ByVal intBaby As Integer, _
                        ByVal strStart As String, _
                        ByVal strEnd As String, _
                        ByVal lng科室ID As Long, _
                        ByVal byt护理级别 As Byte, _
                        ByVal lngId As Long, _
                        ByVal strCode As String, _
                        ByVal strFile As String, _
                        ByVal str科室 As String, _
                        ByVal int保留 As Integer, _
                        ByVal byt文件级别 As Byte, _
                        Optional ByVal blnFirst As Boolean = False) As Boolean
    '******************************************************************************************************************
    '功能：检查指定时间段内有没有护理记录数据，有数据则填写并更新实际的日期范围
    '参数：blnShow=False,表示无数据不显示;True,人工做的数据,强制显示
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    
    With vfgFile
        gstrSQL = "Select Min(r.发生时间) As 开始, Max(r.发生时间) As 截止 From 病人护理记录 r Where r.病人ID = [1] And NVL(r.主页ID, 0) = [2] And Nvl(r.婴儿,0)=[3] And r.发生时间 between [4] And [5] And r.科室id=[6] And r.护理级别<=[7]"
        If mblnMoved_HL Then
            gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
            gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
        End If
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngPageId, intBaby, CDate(strStart), CDate(strEnd), lng科室ID, byt文件级别)
        
        If rs.EOF = False Then
            
            If strEnd >= Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") Then strEnd = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            If zlCommFun.NVL(rs("开始").Value, "") <> "" Then

                strStart = Format(rs("开始").Value, "yyyy-MM-dd HH:mm")
                strEnd = Format(rs("截止").Value, "yyyy-MM-dd HH:mm")
                
                If Val(.TextMatrix(.Rows - 1, mCol.fID)) > 0 Then .AddItem ""

                If int保留 = -1 Then
                    Set .Cell(flexcpPicture, .Rows - 1, mCol.f标志) = imgData.ListImages("体温").Picture
                Else
                    Set .Cell(flexcpPicture, .Rows - 1, mCol.f标志) = imgData.ListImages("普通").Picture
                End If

                .TextMatrix(.Rows - 1, mCol.fID) = lngId
                .TextMatrix(.Rows - 1, mCol.f编号) = strCode
                .TextMatrix(.Rows - 1, mCol.f文件) = strFile
                .TextMatrix(.Rows - 1, mCol.f日期范围) = Format(strStart, "yyyy-MM-dd HH:mm") & " ～ " & Format(strEnd, "yyyy-MM-dd HH:mm")
                .TextMatrix(.Rows - 1, mCol.f病区id) = lng科室ID
                .TextMatrix(.Rows - 1, mCol.f病区名) = str科室
                .TextMatrix(.Rows - 1, mCol.f护理级别) = IIf(int保留 = -1, "/", byt护理级别)
                .TextMatrix(.Rows - 1, mCol.f文件级别) = IIf(int保留 = -1, "/", byt文件级别)
                .TextMatrix(.Rows - 1, mCol.f保留) = int保留
            End If
        End If
    End With

    ShowFile = True

End Function

Private Function ShowFileOnly(ByVal lngPatiID As Long, _
                        ByVal lngPageId As Long, _
                        ByVal intBaby As Integer, _
                        ByVal strStart As String, _
                        ByVal strEnd As String, _
                        ByVal lng科室ID As Long, _
                        ByVal byt护理级别 As Byte, _
                        ByVal lngId As Long, _
                        ByVal strCode As String, _
                        ByVal strFile As String, _
                        ByVal str科室 As String, _
                        ByVal int保留 As Integer, _
                        ByVal byt文件级别 As Byte, _
                        Optional ByVal blnFirst As Boolean = False) As Boolean
    '******************************************************************************************************************
    '功能：检查指定时间段内有没有护理记录数据，有数据则填写并更新实际的日期范围
    '参数：blnShow=False,表示无数据不显示;True,人工做的数据,强制显示
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Static sbyt护理级别  As Byte
    Static slng科室ID As Long
    Static slng病人ID As Long
    Static sintBaby As Integer
    Static sbyt文件级别 As Byte

    If slng病人ID <> lngPatiID Or sintBaby <> intBaby Or blnFirst Then
        '如果病人发生变化,清零
        slng病人ID = lngPatiID
        sintBaby = intBaby
        sbyt护理级别 = 0
        slng科室ID = 0
        sbyt文件级别 = 0
    End If
    
    With vfgFile
        gstrSQL = "Select Min(r.发生时间) As 开始, Max(r.发生时间) As 截止 From 病人护理记录 r Where r.病人ID = [1] And NVL(r.主页ID, 0) = [2] And Nvl(r.婴儿,0)=[3] And r.发生时间 between [4] And [5] And r.科室id=[6] And r.护理级别<=[7]"
        If mblnMoved_HL Then
            gstrSQL = Replace(gstrSQL, "病人护理记录", "H病人护理记录")
            gstrSQL = Replace(gstrSQL, "病人护理内容", "H病人护理内容")
        End If
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngPageId, intBaby, CDate(strStart), CDate(strEnd), lng科室ID, byt文件级别)
        
        If rs.EOF = False Then
            
            If strEnd >= Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") Then strEnd = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            If zlCommFun.NVL(rs("开始").Value, "") <> "" Then
                If (sbyt护理级别 <> byt护理级别 Or slng科室ID <> lng科室ID Or sbyt文件级别 <> byt文件级别) Then
                    
                    If Val(.TextMatrix(.Rows - 1, mCol.fID)) > 0 Then .AddItem ""
    
                    If int保留 = -1 Then
                        Set .Cell(flexcpPicture, .Rows - 1, mCol.f标志) = imgData.ListImages("体温").Picture
                    Else
                        Set .Cell(flexcpPicture, .Rows - 1, mCol.f标志) = imgData.ListImages("普通").Picture
                    End If
    
                    .TextMatrix(.Rows - 1, mCol.fID) = lngId
                    .TextMatrix(.Rows - 1, mCol.f编号) = strCode
                    .TextMatrix(.Rows - 1, mCol.f文件) = strFile
                    .TextMatrix(.Rows - 1, mCol.f日期范围) = Format(strStart, "yyyy-MM-dd HH:mm") & " ～ " & Format(strEnd, "yyyy-MM-dd HH:mm")
                    .TextMatrix(.Rows - 1, mCol.f病区id) = lng科室ID
                    .TextMatrix(.Rows - 1, mCol.f病区名) = str科室
                    .TextMatrix(.Rows - 1, mCol.f护理级别) = IIf(int保留 = -1, "/", byt护理级别)
                    .TextMatrix(.Rows - 1, mCol.f文件级别) = IIf(int保留 = -1, "/", byt文件级别)
                    .TextMatrix(.Rows - 1, mCol.f保留) = int保留
    '
                    sbyt护理级别 = byt护理级别
                    slng科室ID = lng科室ID
                    sbyt文件级别 = byt文件级别
                Else
                    .TextMatrix(.Rows - 1, mCol.f日期范围) = Split(.TextMatrix(.Rows - 1, mCol.f日期范围), " ～ ")(0) & " ～ " & Format(strEnd, "yyyy-MM-dd HH:mm")
                End If
            End If
    
            If int保留 = -1 Then
                sbyt护理级别 = 0
                slng科室ID = 0
                sbyt文件级别 = 0
            End If
        End If
    End With

    ShowFileOnly = True

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

Public Function RefreshData(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lngDeptId As Long, ByVal blnDoctorStation As Boolean, ByVal blnEdit As Boolean) As Boolean
    '******************************************************************************************************************
    '功能：刷新数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    
    mlngPatiId = lng病人ID
    mlngPageId = lng主页ID
    mlngDeptId = lngDeptId
    mblnEdit = blnEdit And Not mblnMoved_HL
    
    mblnDoctorStation = blnDoctorStation
    mblnRefreshFontSize = False
    Call ExecuteCommand("刷新数据")
    
    If mblnDoctorStation Then
        tbcFile.Item(1).Selected = True
        tbcFile.Item(0).Visible = False
    End If
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
    Dim byt护理等级 As Byte, bytSize As Byte
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        
        Set mclsDockAduits = New zlRichEPR.clsDockAduits
        Call FormSetCaption(mclsDockAduits.zlGetFormTendBody, False, False)
        
        With tbcSub
            With .PaintManager
                .Appearance = xtpTabAppearancePropertyPage2003
                .BoldSelected = True
                .ClientFrame = xtpTabFrameSingleLine
                .ShowIcons = True
                .DisableLunaColors = False
                .Position = xtpTabPositionTop
            End With

            .InsertItem 0, "", picPane(2).hWnd, 0
            .InsertItem 1, "体温记录单", mclsDockAduits.zlGetFormTendBody.hWnd, 0
            .InsertItem 2, "护理记录单", mclsDockAduits.zlGetFormTendFile.hWnd, 0
            .Item(0).Selected = True
            Call SetTabVisible(0)
        End With
        
        With tbcFile
            With .PaintManager
                .Appearance = xtpTabAppearancePropertyPage2003
                .BoldSelected = True
                .ClientFrame = xtpTabFrameSingleLine
                .ShowIcons = True
                .DisableLunaColors = False
                .Position = xtpTabPositionBottom
            End With

            .InsertItem 0, "快速录入", mclsDockAduits.zlGetFormTendEdit.hWnd, 0
            .InsertItem 1, "护理记录单", picRecord.hWnd, 0
            .Item(0).Selected = True
        End With
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
        
        mblnNoRefresh = True
        cboBaby.Clear
        cboBaby.AddItem "病人本人"
        gstrSQL = "Select a.序号,Decode(a.婴儿姓名,Null,NVL(c.姓名,b.姓名) ||'之子'||Trim(To_Char(a.序号,'9')),a.婴儿姓名) As 婴儿姓名" & _
            " From 病人信息 b,病案主页 c,病人新生儿记录 a Where b.病人id=c.病人id And a.病人id=c.病人id And a.主页id=c.主页id And c.病人id=[1] And c.主页id=[2]  Order By a.序号"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngPatiId, mlngPageId)
        If rs.BOF = False Then
            Do While Not rs.EOF
                cboBaby.AddItem rs("婴儿姓名").Value
                rs.MoveNext
            Loop
        End If
        cboBaby.ListIndex = 0
        cboBaby.Visible = (cboBaby.ListCount > 1)
        
        Call zlRefData
        mblnNoRefresh = False
        Call ExecuteCommand("显示文件内容", vfgFile.Row)
        
    '------------------------------------------------------------------------------------------------------------------
    Case "显示文件内容"
        If tbcFile.Item(0).Selected Then
            '提取该病人当时的护理等级
            gstrSQL = "select Zl_Patittendgrade([1],[2]) from dual"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngPatiId, mlngPageId)
            byt护理等级 = rs.Fields(0).Value
            Call mclsDockAduits.zlRefreshTendEdit(mlngPatiId, mlngPageId, mlngDeptId, byt护理等级, 0, mstrPrivs, False, mblnEdit)
        Else
            With vfgFile
                Call mclsDockAduits.zlRefreshTendBody(mlngPatiId, mlngPageId, mlngDeptId, mintBaby)
                If Val(.TextMatrix(.Row, mCol.f保留)) = -1 Then
                    '体温单查看：病人ID;主页ID;病区ID;出院;编辑;婴儿
                    
                    Call SetTabVisible(1)
                    tbcSub.Item(1).Caption = .TextMatrix(.Row, mCol.f文件) & "(" & .TextMatrix(.Row, mCol.f日期范围) & ")"
                
                ElseIf .TextMatrix(.Row, mCol.f文件) <> "文件" And .TextMatrix(.Row, mCol.f文件) <> "" Then
                    
                    Call mclsDockAduits.zlRefresh(3, Val(.TextMatrix(.Row, mCol.fID)), mlngPatiId, mlngPageId, Val(.TextMatrix(.Row, mCol.f病区id)), .TextMatrix(.Row, mCol.f日期范围), Val(.TextMatrix(.Row, mCol.f文件级别)), mintBaby)
                    Call SetTabVisible(2)
                    tbcSub.Item(2).Caption = .TextMatrix(.Row, mCol.f文件) & "(" & .TextMatrix(.Row, mCol.f日期范围) & ")"
                    tbcSub.Item(1).Selected = True
                    tbcSub.Item(2).Selected = True
                Else
                    Call SetTabVisible(0)
                    tbcSub.Item(0).Caption = "无可显示的护理文件"
                End If
                tbcSub.PaintManager.Layout = xtpTabLayoutAutoSize
            End With
        End If
        '进行字体设置
        If mblnRefreshFontSize = True Then Call ExecuteCommand("设置字体")
        mblnRefreshFontSize = True
    Case "设置字体"
        bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
        If tbcFile.Item(0).Selected Then
            '快速录入
            Call mclsDockAduits.SetFontSize(0, bytSize)
        Else
            With vfgFile
                If Val(.TextMatrix(.Row, mCol.f保留)) = -1 Then
                    '体温单记录
                    Call mclsDockAduits.SetFontSize(1, bytSize)
                ElseIf .TextMatrix(.Row, mCol.f文件) <> "文件" And .TextMatrix(.Row, mCol.f文件) <> "" Then
                    '记录单
                    Call mclsDockAduits.SetFontSize(2, bytSize)
                End If
            End With
        End If
    End Select

    ExecuteCommand = True

    GoTo EndHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
EndHand:

End Function

Private Function SetTabVisible(ByVal intIndex As Integer) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    
    If tbcSub.Item(intIndex).Visible = False Then
        tbcSub.Item(intIndex).Visible = True
        tbcSub.Item(intIndex).Selected = True
    End If

    For intLoop = 0 To tbcSub.ItemCount - 1
        If intLoop <> intIndex Then
            If tbcSub.Item(intLoop).Visible = True Then tbcSub.Item(intLoop).Visible = False
        End If
    Next
    SetTabVisible = True
End Function

Private Sub cboBaby_Click()
    If mintBaby = cboBaby.ListIndex Then Exit Sub
    mintBaby = cboBaby.ListIndex
    If mblnNoRefresh = True Then Exit Sub
    Call zlRefData
End Sub

Private Sub Form_Load()
    lblNote.Caption = ""
    mblnMouseMove = False
End Sub

Private Sub Form_Resize()
    Dim intSel As Integer
    On Error Resume Next
    
    picFile.Move 0, 0, Me.ScaleWidth + 500, Me.ScaleHeight + 500
    picRecord.Move 0, 0, picFile.ScaleWidth, picFile.ScaleHeight
    picNote.Move 3000, tbcFile.Top + tbcFile.Height - picNote.Height - 50, picFile.Width - picNote.Left
    
    '重新选择当前页面,如果不这样,那么页头就看不见,怪的很,估计和该控件嵌套有关
    intSel = tbcFile.Selected.Index
    tbcFile.Item(0).Selected = True
    If intSel <> 0 Then tbcFile.Item(intSel).Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnNoRefresh = False
    Set mclsDockAduits = Nothing
    If Not mfrmCaseTendEditForBatch Is Nothing Then Unload mfrmCaseTendEditForBatch
    Set mfrmCaseTendEditForBatch = Nothing
End Sub

Private Sub mclsDockAduits_ShowItemInfo(ByVal strInfo As String)
    lblNote.Width = picNote.Width
    lblNote.Caption = strInfo
End Sub

Private Sub PicFile_Resize()
    On Error Resume Next
    
    tbcFile.Move 0, 0, picFile.ScaleWidth - 500, picFile.ScaleHeight - 500
End Sub


Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        fra.Move 0, -90, picPane(Index).Width
        
        cboBaby.Move fra.Width - cboBaby.Width, cboBaby.Top
        vfgFile.Move 15, fra.Top + fra.Height + 15, picPane(Index).Width - 30, picPane(Index).Height - (fra.Top + fra.Height + 15) - 15
    Case 1
        tbcSub.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    End Select
End Sub

Private Sub picRecord_Resize()
    On Error Resume Next
    If picSplit.Top < 1000 Then picSplit.Top = 1000
    If picSplit.Top > picRecord.Height - 2000 Then picSplit.Top = picRecord.Height - 2000
    
    With picSplit
        .Left = 0
        .Width = picRecord.Width
    End With
    
    With picPane(0)
        .Height = picSplit.Top
        .Width = picSplit.Width
    End With
    
    With picPane(1)
        .Top = picSplit.Top + picSplit.Height
        .Height = picRecord.Height - .Top
        .Width = picSplit.Width
    End With
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnMouseMove = (Button = 1)
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnMouseMove = False Then Exit Sub
    
    If picSplit.Top < 1000 Then picSplit.Top = 1000
    If picSplit.Top > picRecord.Height - 2000 Then picSplit.Top = picRecord.Height - 2000
    picSplit.Move 0, picSplit.Top + Y
    Me.Refresh
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnMouseMove = False
    
    Call picRecord_Resize
End Sub

Private Sub tbcFile_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnNoRefresh = True Then Exit Sub
    lblNote.Caption = ""
    Call ExecuteCommand("显示文件内容")
End Sub

Private Sub vfgFile_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnNoRefresh = True Then Exit Sub
    If OldRow <> NewRow Then
        
        Call ExecuteCommand("显示文件内容", NewRow)
        DoEvents
        
        On Error Resume Next
        vfgFile.SetFocus
    End If

    
End Sub

Private Sub vfgFile_DblClick()
    Dim strInfo As String
    Dim intEdit As Integer
    Dim objFrmBody As Object
    
    On Error GoTo errHand
    
    strInfo = Val(Me.vfgFile.TextMatrix(vfgFile.Row, mCol.f病区id))

    If Val(vfgFile.TextMatrix(vfgFile.Row, mCol.f保留)) = -1 Then
        '体温单查看：病人ID;主页ID;病区ID;出院;编辑;婴儿

        intEdit = 0
        If (InStr(1, mstrPrivs, "体温单作图") > 0 And mblnDoctorStation = False) Then
            If (mblnEdit And mlngPatiId > 0 And TendArchive = False) Then
                intEdit = 1
            End If
        End If
        
        If Not CreateBodyEditor Then Exit Sub
        Set objFrmBody = gobjBodyEditor.GetTendBody
        On Error Resume Next
        objFrmBody.Resize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
        If Err <> 0 Then Err.Clear
        On Error GoTo errHand
        Call objFrmBody.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";0;" & intEdit & ";" & mintBaby, 1, mstrPrivs)

    Else
        With vfgFile
            Call frmTendFileOpen.ShowMe(Me, Val(.TextMatrix(.Row, mCol.fID)), mlngPatiId, mlngPageId, Val(strInfo), mintBaby, .TextMatrix(.Row, mCol.f日期范围), , Val(.TextMatrix(.Row, mCol.f护理级别)), mblnMoved_HL, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
        End With
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vfgFile_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call vfgFile_DblClick
End Sub

'---------------------------------------------------------------------------------
'以下是基础函数或过程
'---------------------------------------------------------------------------------
Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False)
    Dim arrFields, arrValues, intField As Integer
    '更新记录,如果不存在,则新增
    'strPrimary:字段名|值
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'strPrimary = "RecordID,5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '定位到指定记录
    'strPrimary:主健,值
    'blnDelete=True,则该记录集存在"删除"字段
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !删除 = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub
