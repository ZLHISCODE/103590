VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLocalSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "本地打印设置"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   Icon            =   "frmLocalSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picOperateFormat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   720
      Picture         =   "frmLocalSet.frx":27A2
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picOperateFormat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   240
      Picture         =   "frmLocalSet.frx":2B2C
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame fraReport 
      Caption         =   " 报表 "
      Height          =   2670
      Left            =   135
      TabIndex        =   17
      Top             =   105
      Width           =   6120
      Begin VB.CheckBox chkAllFormat 
         Caption         =   "程序内部调用打印时自动打印所有的格式"
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   2280
         Width           =   3540
      End
      Begin VB.ComboBox cboFormat 
         Height          =   300
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1845
         Width           =   4230
      End
      Begin MSComctlLib.ImageList img32 
         Left            =   135
         Top             =   1155
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLocalSet.frx":2EB6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblFormat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缺省报表格式"
         Height          =   180
         Left            =   360
         TabIndex        =   0
         Top             =   1905
         Width           =   1080
      End
      Begin VB.Label lblLoc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "位置:"
         Height          =   180
         Left            =   975
         TabIndex        =   23
         Top             =   975
         Width           =   690
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "尺寸:"
         Height          =   180
         Left            =   975
         TabIndex        =   21
         Top             =   1515
         Width           =   450
      End
      Begin VB.Label lblPaper 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "纸张:"
         Height          =   180
         Left            =   975
         TabIndex        =   20
         Top             =   1245
         Width           =   450
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "报表说明:"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   975
         TabIndex        =   19
         Top             =   540
         Width           =   4200
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "报表编号名称:"
         Height          =   180
         Left            =   975
         TabIndex        =   18
         Top             =   270
         Width           =   1170
      End
      Begin VB.Image imgReport 
         Height          =   480
         Left            =   135
         Picture         =   "frmLocalSet.frx":3790
         Top             =   390
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5040
      TabIndex        =   15
      Top             =   6480
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3840
      TabIndex        =   14
      Top             =   6480
      Width           =   1100
   End
   Begin VB.Frame fraPrinter 
      Caption         =   " 输出 "
      Height          =   3540
      Left            =   135
      TabIndex        =   16
      Top             =   2835
      Width           =   6120
      Begin VSFlex8Ctl.VSFlexGrid vsfReportFormat 
         Height          =   1335
         Left            =   1590
         TabIndex        =   4
         Top             =   240
         Width           =   4230
         _cx             =   7461
         _cy             =   2355
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
         RowHeightMin    =   300
         RowHeightMax    =   300
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
      Begin VB.ComboBox cboForm 
         Height          =   300
         ItemData        =   "frmLocalSet.frx":3A9A
         Left            =   1590
         List            =   "frmLocalSet.frx":3A9C
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3045
         Width           =   4230
      End
      Begin VB.CheckBox chkForm 
         Caption         =   "自定义纸张通过打印服务器的格式来设置"
         Height          =   195
         Left            =   1590
         TabIndex        =   12
         ToolTipText     =   "以前老的打印方式，当打印出现问题时才尝试使用。需要去掉打印机属性中的高级打印功能。"
         Top             =   2775
         Width           =   3540
      End
      Begin VB.TextBox txtCopy 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   5250
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "1"
         Top             =   2340
         Width           =   315
      End
      Begin VB.ComboBox cboBin 
         Height          =   300
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2340
         Width           =   2340
      End
      Begin VB.ComboBox cboPrinter 
         Height          =   300
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1680
         Width           =   4230
      End
      Begin MSComCtl2.UpDown udCopy 
         Height          =   300
         Left            =   5565
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2340
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtCopy"
         BuddyDispid     =   196625
         OrigLeft        =   1935
         OrigTop         =   240
         OrigRight       =   2175
         OrigBottom      =   585
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "报表格式"
         Height          =   180
         Left            =   795
         TabIndex        =   3
         Top             =   285
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "打印份数"
         Height          =   180
         Left            =   4500
         TabIndex        =   9
         Top             =   2400
         Width           =   720
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   210
         Picture         =   "frmLocalSet.frx":3A9E
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "纸张来源"
         Height          =   180
         Left            =   795
         TabIndex        =   7
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label lblLocal 
         AutoSize        =   -1  'True
         Caption         =   "位置"
         Height          =   180
         Left            =   1605
         TabIndex        =   22
         Top             =   2055
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "打印机"
         Height          =   180
         Left            =   795
         TabIndex        =   5
         Top             =   1740
         Width           =   540
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   1320
      Top             =   6480
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmLocalSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rsInfo As ADODB.Recordset    'IN
Public mblnOutCall As Boolean       'IN:是否外部通过接口在调用
Public mintFormat As Integer        'IN:指定要设置的格式，为0表示不指定

Private Const MLNG_POPMENU_ID As Long = 10000

Private mblnStartUp As Boolean
Private mrsFormat As ADODB.Recordset '记录格式信息

Private Sub CboFormat_Click()
    If cboFormat.ListIndex = -1 Then Exit Sub
    mrsFormat.Filter = "序号=" & cboFormat.ListIndex + 1
    If mrsFormat.EOF Then Exit Sub
    
    '自定义纸张处理方式
    chkForm.Enabled = Nvl(mrsFormat!纸张, 0) = 256
    
    '格式的纸张信息
    lblPaper.Caption = "纸张:" & GetPaperName(Nvl(mrsFormat!纸张, 9), Nvl(mrsFormat!W, INIT_WIDTH), Nvl(mrsFormat!H, INIT_HEIGHT))
    lblSize.Caption = "尺寸:" & CInt(Nvl(mrsFormat!W, INIT_WIDTH) / Twip_mm) & "mm(宽) × " & _
        CInt(Nvl(mrsFormat!H, INIT_HEIGHT) / Twip_mm) & "mm(高)   " & _
        Switch(IsNull(mrsFormat!纸向), "纵向", mrsFormat!纸向 = 1, "纵向", mrsFormat!纸向 = 2, "横向")
End Sub

Private Sub cboPrinter_Click()
    Dim i As Integer, j As Integer
    Dim k As Integer, strTmp As String
    Dim lngCount As Long, intCur As Integer
    Dim strPaperBin As String * 100
    Dim strPaperBinName As String * 1000
    
    If cboPrinter.ListIndex = -1 Or cboPrinter.Tag = "1" Then Exit Sub

    Set Printer = Printers(cboPrinter.ListIndex)
    lblLocal.Caption = "位置:" & Printer.Port
    
    '设置可用进纸方式
    cboBin.Clear
    '--------------------------------------------------------------------------------------------
    Call DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINNAMES, strPaperBinName, 0)
    lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINS, strPaperBin, 0)
    
    'GetSetting函数语句放在API调用之前会崩溃(？？)
    
    j = 1
    For i = 1 To lngCount
        k = 0
        '进纸名称
        Do
            If Mid(strPaperBinName, j, 1) = Chr(0) Then
                If Trim(strTmp) <> "" Then
                    cboBin.AddItem Trim(strTmp)
                    
                    '进纸编号
                    cboBin.ItemData(cboBin.ListCount - 1) = Asc(Mid(strPaperBin, i * 2, 1)) * 256# + Asc(Mid(strPaperBin, i * 2 - 1, 1))
                    If cboBin.ItemData(cboBin.ListCount - 1) = intCur Then
                        cboBin.ListIndex = cboBin.ListCount - 1 '定位在原设置上
                    End If
                    If cboBin.ListIndex = -1 And cboBin.ItemData(cboBin.ListCount - 1) = Printer.PaperBin Then
                        cboBin.ListIndex = cboBin.ListCount - 1 '定位在打印机缺省设置上
                    End If
                End If
                
                j = 24 + j - LenB(StrConv(strTmp, vbFromUnicode))
                strTmp = ""
                Exit Do
            Else
                strTmp = strTmp & Mid(strPaperBinName, j, 1)
                j = j + 1
                k = k + 1
                If k > 24 Then Exit Do
            End If
        Loop
    Next
    
    If cboBin.ListIndex = -1 And cboBin.ListCount > 0 Then cboBin.ListIndex = 0
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
    Case MLNG_POPMENU_ID + 1
    Case MLNG_POPMENU_ID + 2
    End Select
    Dim i As Integer
    Dim blnFind As Boolean
    
    '检查VSF中是否存在
    With vsfReportFormat
        For i = 0 To .Rows - 1
            If Trim(Control.Caption) = Trim(.TextMatrix(i, .ColIndex("格式"))) Then
                blnFind = True
                .Col = .ColIndex("格式")
                .Row = i
                Exit For
            End If
        Next

        '添加单独报表格式的打印设置
        If blnFind = False Then
            '将注册表信息写到CellData缓存
            .Redraw = False
            .Rows = .Rows + 1
            .Row = .Rows - 1
            i = .Row
            .TextMatrix(i, .ColIndex("格式")) = Trim(Control.Caption)
            Call PrinterInfo2CellData(i, rsInfo!编号, Trim(Control.Caption))
            .Redraw = True

            '控件刷新
            Call SetPrintInfo(.Cell(flexcpData, i, .ColIndex("格式")))
        End If
    End With

End Sub

Private Sub chkForm_Click()
    If chkForm.Value = 1 And chkForm.Enabled = True Then
        cboForm.Enabled = True
    Else
        cboForm.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngW As Long, lngH As Long
    Dim strSection As String, strZLHIS As String, strInfo As String
    Dim arrItems As Variant, arrTmp As Variant, arrK As Variant
    Dim i As Integer, j As Integer
    
    arrItems = Array()
    
    '检查设置Form和服务器误差不能大于80%
    If chkForm.Value = 1 Then
        If cboForm.ListIndex <> -1 Then
            lngW = Val(Split(Split(cboForm.List(cboForm.ListIndex), " ")(1), "×")(0))
            lngH = Val(Split(Split(cboForm.List(cboForm.ListIndex), " ")(1), "×")(1))
            If Abs(CInt(Nvl(mrsFormat!W, INIT_WIDTH) / Twip_mm) - lngW) / CInt(Nvl(mrsFormat!W, INIT_WIDTH) / Twip_mm) > 4 / 5 _
                Or Abs(CInt(Nvl(mrsFormat!H, INIT_HEIGHT) / Twip_mm) - lngH) / CInt(Nvl(mrsFormat!H, INIT_HEIGHT) / Twip_mm) > 4 / 5 Then
                MsgBox "自定义纸张的大小不能超过报表缺省大小的80%", vbInformation, App.Title
                Exit Sub
            End If
        End If
    End If
    
    If cboFormat.Enabled And cboFormat.Visible Then
        SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\LocalSet\" & rsInfo!编号, "Format", cboFormat.ListIndex + 1
    End If
    
    '正确性检查
    If cboPrinter.ListIndex = -1 Then
        MsgBox "请选择一个打印机！", vbInformation, App.Title
        cboPrinter.SetFocus: Exit Sub
    End If
    
    SaveSetting "ZLSOFT", "私有模块\" & App.ProductName & "\LocalSet\" & rsInfo!编号, "AllFormat", chkAllFormat.Value
    
    '保存设置
    ''当前的设置更新至CellData
    strInfo = GetPrintInfo("", True)
    vsfReportFormat.Cell(flexcpData, vsfReportFormat.Row, 0) = strInfo
    
    ''获取报表的所有格式
    strSection = "Software\VB and VBA Program Settings\ZLSOFT\私有模块\" & App.ProductName & "\LocalSet\" & rsInfo!编号
    strZLHIS = "私有模块\" & App.ProductName & "\LocalSet\" & rsInfo!编号
    
    '清除注册表的全部格式
    arrItems = mdlPublic.GetAllSubKey(HKEY_CURRENT_USER, strSection)
    For i = LBound(arrItems) To UBound(arrItems)
        Call SHDeleteKey(HKEY_CURRENT_USER, strSection & "\" & arrItems(i))
    Next
        
    '保存新信息
    With vsfReportFormat
        For i = 0 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("格式"))) <> "" Then
                'Cell内容格式：Key1=Value1|Key2=Value2|...
                strInfo = .Cell(flexcpData, i, .ColIndex("格式"))
                arrTmp = Split(strInfo, "|")
                For j = LBound(arrTmp) To UBound(arrTmp)
                    '保存
                    arrK = Split(arrTmp(j) & "=", "=")
                    If Trim(arrK(0)) <> "" Then
                        SaveSetting "ZLSOFT", _
                                    strZLHIS & "\" & Trim(.TextMatrix(i, .ColIndex("格式"))), _
                                    Trim(arrK(0)), _
                                    Trim(arrK(1))
                    End If
                Next
            End If
        Next
    End With
    
    gblnOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    DoEvents
    
    cboPrinter.SetFocus
    Err.Clear: On Error GoTo 0
End Sub

Private Sub Form_Load()
    Dim strCur As String, i As Integer
    Dim intFormat As Integer, strSQL As String
    Dim lngW As Long, lngH As Long, intOrient As Integer
    Dim strFormName As String
    Dim strTmp As String
    Dim strDefault As String
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControlMain As CommandBarControl
    
    mblnStartUp = True
    gblnOK = False
            
    SetForegroundWindow hwnd
    SetActiveWindow hwnd
    
    '票据的标识
    lblName.Caption = IIF(Nvl(rsInfo!票据, 0) = 1, "票据", "报表") & ":[" & rsInfo!编号 & "]" & rsInfo!名称
    fraReport.Caption = IIF(Nvl(rsInfo!票据, 0) = 1, " 票据 ", " 报表 ")
    If Nvl(rsInfo!票据, 0) = 1 Then Set imgReport.Picture = img32.ListImages(1).Picture
    
    '缺省显示的格式
    If mintFormat <> 0 Then
        intFormat = mintFormat
    Else
        '缺省为第一种格式
        strTmp = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\LocalSet\" & rsInfo!编号, "Format", "")
        If strTmp = "" Then
            intFormat = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\LocalSet\" & rsInfo!编号, "Format", 1))
        Else
            intFormat = Val(strTmp)
        End If
    End If
    
    '报表格式的弹出菜单初始化
    Call InitCommandBars
    
    '报表可用的格式
    On Error GoTo errH
    strSQL = "Select 报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样 From zlRPTFMTs Where 报表ID=[1] Order by 序号"
    Set mrsFormat = OpenSQLRecord(strSQL, Me.Caption, Val(rsInfo!id))
    On Error GoTo 0
    cboFormat.Clear
    
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, MLNG_POPMENU_ID, "", -1, False)
    cbrMenuBar.id = MLNG_POPMENU_ID
    
    For i = 1 To mrsFormat.RecordCount
        cboFormat.AddItem mrsFormat!说明
        If mrsFormat!序号 = intFormat Then
            CboSetIndex cboFormat.hwnd, cboFormat.NewIndex
        End If
        
        '加载报表格式至弹出菜单
        Set cbrControlMain = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, MLNG_POPMENU_ID + i, Trim(mrsFormat!说明))
        
        mrsFormat.MoveNext
    Next
    If cboFormat.ListIndex = -1 And cboFormat.ListCount > 0 Then CboSetIndex cboFormat.hwnd, 0
    Call CboFormat_Click
    cboFormat.Enabled = mblnOutCall
    
    '报表说明
    lblNote.Caption = "说明:" & IIF(IsNull(rsInfo!说明), "", rsInfo!说明)
    If Not IsNull(rsInfo!程序ID) Then lblLoc.Caption = "位置:" & GetMenuPath(rsInfo!id)
    lblLoc.ToolTipText = lblLoc.Caption
    
    '打印机信息
    cboPrinter.Tag = "1"
    For i = 0 To Printers.count - 1
        cboPrinter.AddItem Printers(i).DeviceName
    Next
    cboPrinter.Tag = ""
    
    '打印Form信息
    On Error Resume Next
    SetNTPrinterPaper_Form Me.hwnd, lngW, lngH, 0, 0, cboForm
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo errH
    
    strTmp = GetSetting("ZLSOFT", "私有模块\" & App.ProductName & "\LocalSet\" & rsInfo!编号, "AllFormat", "")
    If strTmp = "" Then strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\LocalSet\" & rsInfo!编号, "AllFormat", 0)
    chkAllFormat.Value = Val(strTmp)
    
    '如果是票据，则只能打印1份
    If Nvl(rsInfo!票据, 0) = 1 Then
        txtCopy.Enabled = False
        udCopy.Enabled = False
        txtCopy.Text = 1
    End If
    
    '窗体控件设置
    If mblnOutCall Then
        chkAllFormat.Visible = True
    Else
        chkAllFormat.Visible = False
        lblFormat.Top = lblFormat.Top + chkAllFormat.Height
        cboFormat.Top = cboFormat.Top + chkAllFormat.Height
    End If
    If IsWindowsNT Then
        chkForm.Visible = True
    Else
        chkForm.Visible = False
        fraPrinter.Height = fraPrinter.Height - chkForm.Height - 60
        cboForm.Visible = False
        fraPrinter.Height = fraPrinter.Height - cboForm.Height - 60
    End If
    
    Call IniReportFormat(rsInfo!编号)
    Call chkForm_Click
    
    fraPrinter.Top = fraReport.Top + fraReport.Height + 60
    cmdOK.Top = fraPrinter.Top + fraPrinter.Height + 150
    cmdCancel.Top = cmdOK.Top
    Me.Height = cmdOK.Top + cmdOK.Height + 150 + (Me.Height - Me.ScaleHeight)
    mblnStartUp = False
    Exit Sub
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsInfo = Nothing
    mblnOutCall = False
    mintFormat = 0
    Set mrsFormat = Nothing
End Sub

Private Sub txtCopy_GotFocus()
    Call SelAll(txtCopy)
End Sub

Private Sub IniReportFormat(ByVal strRPTCode As String)
    Dim strSection As String, strPathKey As String, strTmp As String
    Dim arrItems As Variant
    Dim i As Integer, intRow As Integer
    
    arrItems = Array()

    '界面
    With vsfReportFormat
        '.Appearance = flexFlat
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
        .FixedCols = 0
        .FixedRows = 0
        .SelectionMode = flexSelectionByRow
        .SheetBorder = .BackColor
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .ExplorerBar = flexExNone
        .AutoResize = True
        .SheetBorder = &H8000000F
        .BackColorBkg = &H80000005
        .FocusRect = flexFocusHeavy
        .ScrollBars = flexScrollBarVertical
    
        .Cols = 3
        .Rows = 1
        
        .ColKey(0) = "格式"
        .ColKey(1) = "新增"
        .ColKey(2) = "删除"
        
        .TextMatrix(0, .ColIndex("格式")) = "所有格式"
        .Cell(flexcpPicture, 0, .ColIndex("新增")) = picOperateFormat(0).Picture
        .Cell(flexcpPicture, 0, .ColIndex("删除")) = Nothing
        .Cell(flexcpPictureAlignment, 0, .ColIndex("新增")) = flexAlignCenterCenter
        .Cell(flexcpPictureAlignment, 0, .ColIndex("删除")) = flexAlignCenterCenter
        .Cell(flexcpData, 0, .ColIndex("格式"), 0, .ColIndex("格式")) = ""
    End With
    
    '数据项（报表格式）
    ''本机
    strSection = "Software\VB and VBA Program Settings\ZLSOFT\私有模块\" & App.ProductName & "\LocalSet\" & strRPTCode
    arrItems = mdlPublic.GetAllSubKey(HKEY_CURRENT_USER, strSection)    '所有子结点
    With vsfReportFormat
        .Redraw = False
        For i = LBound(arrItems) To UBound(arrItems)
            If Trim(arrItems(i)) = "所有格式" Then
                intRow = 0
            Else
                .Rows = .Rows + 1
                intRow = .Rows - 1
            End If
            .TextMatrix(intRow, .ColIndex("格式")) = Trim(arrItems(i))
            
            '将注册表信息写到CellData缓存
            Call PrinterInfo2CellData(intRow, strRPTCode, Trim(arrItems(i)))
        Next
        
        '更新所有格式对应的打印设置界面
        For i = 0 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("格式"))) = "所有格式" Then
                .Row = i
                '界面控件更新
                Call SetPrintInfo(.Cell(flexcpData, .Row, .ColIndex("格式"), .Row, .ColIndex("格式")))
                Exit For
            End If
        Next
        .Redraw = True
    End With
    
    '调整列宽
    With vsfReportFormat
        If .Rows > 4 Then
            .ColWidth(0) = .Width - 15 * 8 * 3 * 2 - 60 - 240
        Else
            .ColWidth(0) = .Width - 15 * 8 * 3 * 2 - 60
        End If
        .ColWidth(1) = 15 * 8 * 3
        .ColWidth(2) = .ColWidth(1)
    End With
End Sub

Private Function GetPrintInfo(ByVal strInfo As String, Optional ByVal blnFromInterface As Boolean = False) As String
    Const STR_ITEMS As String = "PaperBin|PaperCopy|PaperForm|Printer|PaperFormName"
    Dim i As Integer
    Dim arrItems As Variant
    Dim strKey As String, strValue As String, strResult As String
    
    arrItems = Split(STR_ITEMS, "|")
    
    strResult = ""
    If blnFromInterface Then
        If cboBin.ListIndex < 0 Then
            strResult = strResult & "|" & arrItems(0) & "="
        Else
            strResult = strResult & "|" & arrItems(0) & "=" & cboBin.ItemData(cboBin.ListIndex)
        End If
        strResult = strResult & "|" & arrItems(1) & "=" & Trim(txtCopy.Text)
        strResult = strResult & "|" & arrItems(2) & "=" & chkForm.Value
        strResult = strResult & "|" & arrItems(3) & "=" & Trim(cboPrinter.Text)
        strResult = strResult & "|" & arrItems(4) & "=" & Trim(cboForm.Text)
    Else
        For i = LBound(arrItems) To UBound(arrItems)
            strKey = arrItems(i)
            strValue = GetSetting("ZLSOFT", strInfo, strKey, "")
            strResult = strResult & "|" & strKey & "=" & strValue
        Next
    End If
    If Left(strResult, 1) = "|" Then strResult = Mid(strResult, 2)
    
    GetPrintInfo = strResult
End Function

Private Sub SetPrintInfo(ByVal strInfo As String)
    Dim i As Integer, j As Integer
    Dim arrItems As Variant, arrKey As Variant
    
    arrItems = Split(strInfo, "|")
    For i = LBound(arrItems) To UBound(arrItems)
        arrKey = Split(arrItems(i), "=")
        Select Case LCase(arrKey(0))
        Case "papercopy"
            txtCopy.Text = "" & IIF(Val(arrKey(1)) <= 0, 1, Val(arrKey(1)))
        Case "paperform"
            chkForm.Value = Val(arrKey(1))
        Case "printer"
            For j = 0 To cboPrinter.ListCount - 1
                If LCase(cboPrinter.List(j)) = LCase(Trim(arrKey(1))) Then
                    cboPrinter.ListIndex = j
                    Exit For
                End If
            Next
        Case "paperformname"
            For j = 0 To cboForm.ListCount - 1
                If LCase(cboForm.List(j)) = LCase(Trim(arrKey(1))) Then
                    cboForm.ListIndex = j
                    Exit For
                End If
            Next
        End Select
    Next
    
    '最后处理PaperBin
    For i = LBound(arrItems) To UBound(arrItems)
        arrKey = Split(arrItems(i), "=")
        If LCase(arrKey(0)) = "paperbin" Then
            For j = 0 To cboBin.ListCount - 1
                If cboBin.ItemData(j) = Val(arrKey(1)) Then
                    cboBin.ListIndex = j
                    Exit For
                End If
            Next
            
            Exit For
        End If
    Next
    
    cboForm.Enabled = chkForm.Value = 1
    If cboForm.Enabled = False Then cboForm.ListIndex = -1
End Sub

Private Sub PrinterInfo2CellData(ByVal lngRow As Long, ByVal strCode As String, ByVal strFormat As String)
    Dim strTmp As String

    With vsfReportFormat
        If lngRow <> 0 Then
            .Cell(flexcpPicture, lngRow, .ColIndex("新增")) = picOperateFormat(0).Picture
            .Cell(flexcpPicture, lngRow, .ColIndex("删除")) = picOperateFormat(1).Picture
        End If
        .Cell(flexcpPictureAlignment, lngRow, 0, lngRow, 2) = flexAlignCenterCenter
        
        strTmp = "私有模块\" & App.ProductName & "\LocalSet\" & strCode & "\" & strFormat
        If strTmp = "" Then
            '兼容旧处理
            strTmp = "私有模块\" & App.ProductName & "\LocalSet\" & strCode
            If strTmp = "" Then
                strTmp = "私有模块\" & gstrDBUser & "\" & App.ProductName & "\LocalSet\" & strCode
            End If
        End If
        
        '保存
        .Cell(flexcpData, lngRow, .ColIndex("格式"), lngRow, .ColIndex("格式")) = GetPrintInfo(strTmp, False)
    End With
End Sub

Private Sub vsfReportFormat_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Me.Visible = False Then Exit Sub
    
    If OldRow <> NewRow Then
        Call SetPrintInfo(vsfReportFormat.Cell(flexcpData, NewRow, vsfReportFormat.ColIndex("格式")))
    End If
End Sub

Private Sub vsfReportFormat_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim strInfo As String
    
    If Me.Visible = False Then Exit Sub
    
    If cboPrinter.ListIndex < 0 Then
        MsgBox "请完成“打印机”的选择！", vbInformation, App.Title
        Cancel = True
        If cboPrinter.Visible And cboPrinter.Enabled Then cboPrinter.SetFocus
        Exit Sub
    End If

    If OldRow <> NewRow Then
        '保存原记录行的打印设置
        strInfo = GetPrintInfo("", True)
        If vsfReportFormat.Rows - 1 >= OldRow Then
            With vsfReportFormat
                .Redraw = False
                .Cell(flexcpData, OldRow, 0, OldRow, 0) = strInfo
                .Redraw = True
            End With
        End If
    End If
End Sub

Private Sub vsfReportFormat_Click()
    Dim objPopup As CommandBarPopup
    
    With vsfReportFormat
        Select Case .Col
        Case .ColIndex("新增")
            If cboFormat.ListCount > 1 Then
                Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, MLNG_POPMENU_ID)
                If Not objPopup Is Nothing Then
                    objPopup.CommandBar.ShowPopup
                End If
            Else
                MsgBox "该报表不存在多格式，新增无意义！", vbInformation, App.Title
            End If
            .Col = .ColIndex("格式")
        Case .ColIndex("删除")
            If .Row > 0 And Not .Cell(flexcpPicture, .Row, .Col) Is Nothing Then
                If MsgBox("确定删除该报表格式的打印设置？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
                    .RemoveItem .Row
                End If
            End If
            .Col = .ColIndex("格式")
        End Select
    End With
End Sub

Private Sub InitCommandBars()
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsMain.VisualTheme = xtpThemeOffice2003 'xtpthemeoffice2000有凹凸感
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    With cbsMain
        .EnableCustomization False
        'Set .Icons = zlCommFun.GetPubIcons
        .ActiveMenuBar.Title = "菜单"
        .ActiveMenuBar.EnableDocking xtpFlagHideWrap Or xtpFlagStretched
        .ActiveMenuBar.Visible = False
    End With
End Sub
