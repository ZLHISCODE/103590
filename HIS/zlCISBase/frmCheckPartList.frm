VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmCheckPartList 
   Caption         =   "检查部位设置"
   ClientHeight    =   6435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8970
   Icon            =   "frmCheckPartList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox cboKind 
      Height          =   300
      Left            =   6120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   240
      Width           =   1800
   End
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   75
      ScaleHeight     =   5295
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   450
      Width           =   4425
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   4410
         Left            =   15
         TabIndex        =   1
         Top             =   30
         Width           =   4395
         _Version        =   589884
         _ExtentX        =   7752
         _ExtentY        =   7779
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   90
         Top             =   4800
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
               Picture         =   "frmCheckPartList.frx":058A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6060
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCheckPartList.frx":0B24
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10742
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   900
      Left            =   180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5625
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmCheckPartList.frx":13B6
      Left            =   945
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmCheckPartList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    图标 = 0: 分组: 编码: 名称: 方法: 组织: 备注
End Enum

Const conPane_List = 201
Const conPane_Edit = 202

'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者权限串

Private mfrmEdit As frmCheckPartEdit
Private mintEditState As Integer    '当前编辑状态：0-非编辑状态,1-编辑状态
Private mblnPACSInterface As Boolean        '启用影像信息系统接口
Private mstrKind As String          '当前类型
Private mstrPart As String          '当前部位
Private mstrCopy As String          '待粘贴的方法
Private mLngEditWidth As Long       '为适应大字体情况下窗体变大.先读入窗体大小.
'-----------------------------------------------------
'临时变量
'-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrCustom As CommandBarControlCustom
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

Dim lngCount As Long

'-----------------------------------------------------
'以下为内部公共程序
'-----------------------------------------------------
Private Function RefKindList() As Long
    '功能：装入诊疗检查类型
    '返回：类型数量
    Dim rsTemp As New ADODB.Recordset
    gstrSql = "Select 编码, 名称 From 诊疗检查类型"
    
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.cboKind.Clear
    With rsTemp
        Do While Not .EOF
            Me.cboKind.AddItem !编码 & "-" & !名称
            .MoveNext
        Loop
    End With
    If Me.cboKind.ListCount > 0 Then Me.cboKind.ListIndex = 0
    RefKindList = Me.cboKind.ListCount
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    RefKindList = 0
End Function

Private Function zlRefList(Optional strPart As String) As Long
    '功能：刷新装入清单，并定位到指定的记录
    Dim rsTemp As New ADODB.Recordset
    Dim aryItem() As String, strItems As String, strTemp As String
    Dim aryChild() As String, lngChild As Long
    gstrSql = "Select 编码, 名称, 分组, 方法, 备注 From 诊疗检查部位 Where 类型 = [1]"
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrKind)
    Me.rptList.Records.DeleteAll
    With rsTemp
        Do While Not .EOF
            Set rptRcd = Me.rptList.Records.Add()
            Set rptItem = rptRcd.AddItem(CStr("0")): rptItem.Icon = 0
            rptRcd.AddItem CStr(IIf(IsNull(!分组), "未分组", !分组))
            Set rptItem = rptRcd.AddItem(CStr(!编码)): rptItem.SortPriority = Val(("" & !编码))
            rptRcd.AddItem CStr("" & !名称)
            rptRcd.AddItem CStr("" & !方法)
            
            strItems = ""
            strTemp = ""
            For lngCount = 1 To Len(!方法)
                If Mid(!方法, lngCount, 1) = vbTab And lngCount > 2 Then
                    If Mid(strTemp, Len(strTemp), 1) <> ";" Then strTemp = strTemp & ";"
                End If
                strTemp = strTemp & Mid(!方法, lngCount, 1)
            Next
        
            aryItem() = IIf(Mid(strTemp, 1, 1) = ";", Split(Mid(strTemp, 2), ";"), Split(strTemp, ";"))
            strTemp = ""
            For lngCount = 0 To UBound(aryItem)
                If InStr(1, aryItem(lngCount), vbTab) = 0 Then
                    strTemp = Mid(aryItem(lngCount), 2)
                Else
                    strTemp = Mid(aryItem(lngCount), 3)
                End If
                If InStr(1, strTemp, ",") > 0 Then
                    aryChild = Split(strTemp, ",")
                    strTemp = ""
                    For lngChild = 1 To UBound(aryChild)
                        strTemp = strTemp & " □" & Mid(aryChild(lngChild), 2)
                    Next
                    strTemp = aryChild(0) & "(" & Trim(strTemp) & ")"
                End If
                
                
                If InStr(1, aryItem(lngCount), vbTab) = 0 Then
                    strItems = strItems & " ○" & strTemp
                Else
                    strItems = strItems & " □" & strTemp
                End If
            Next
            rptRcd.AddItem CStr(Trim(strItems))
            rptRcd.AddItem CStr("" & !备注)
            .MoveNext
        Loop
    End With
    With Me.rptList
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(mCol.分组)
        .GroupsOrder(0).SortAscending = True
        .Populate
    End With
    If strPart <> "" Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mCol.编码).Value) = strPart Then
                    Set Me.rptList.FocusedRow = rptRow
                    Exit For
                End If
            End If
        Next
    End If
    If Me.rptList.FocusedRow Is Nothing And Me.rptList.Rows.Count > 0 Then
        If Me.rptList.Rows(0).GroupRow Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0).Childs(0)
        Else
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
    End If
    Call rptList_SelectionChanged
    
    zlRefList = Me.rptList.Records.Count
    Me.stbThis.Panels(2).Text = "共有" & Me.rptList.Records.Count & "个部位"
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
End Function

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '复制数据表格
    If zlControl.RPTCopyToVSF(Me.rptList, Me.vfgList) Is Nothing Then Exit Sub
     
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgList
    objPrint.Title.Text = mstrKind & "检查部位与方法清单"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cboKind_Click()
    mstrKind = Split(Me.cboKind.Text, "-")(1)
    Call zlRefList(mstrPart)
    If Me.Visible Then Me.rptList.SetFocus
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strReturn As String
    Dim blnRisTrans As Boolean
    Dim blnTrans As Boolean
    
    '------------------------------------
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
    
    Case conMenu_Edit_Save:
        strReturn = mfrmEdit.zlEditSave()
        If strReturn <> "" Then
            mstrPart = strReturn: Call zlRefList(mstrPart)
            mintEditState = 0: Me.cboKind.Enabled = True: Me.picList.Enabled = True: Me.rptList.SetFocus
        End If
    Case conMenu_Edit_Untread:
        Call mfrmEdit.zlEditCancel
        mintEditState = 0: Me.cboKind.Enabled = True: Me.picList.Enabled = True: Me.rptList.SetFocus
    Case conMenu_Edit_NewItem
        If mfrmEdit.zlEditStart(True, mstrPart) = False Then Exit Sub
        mintEditState = 1: Me.cboKind.Enabled = False: Me.picList.Enabled = False
        Me.dkpMan.FindPane(conPane_Edit).Select
    Case conMenu_Edit_Modify
        If mstrPart = "" Then Exit Sub
        If mfrmEdit.zlEditStart(False, mstrPart) = False Then Exit Sub
        mintEditState = 1: Me.cboKind.Enabled = False: Me.picList.Enabled = False
        Me.dkpMan.FindPane(conPane_Edit).Select

    Case conMenu_Edit_Delete
        Dim strMsg As String
        With Me.rptList
            strMsg = "真的删除该检查部位吗？"
            strMsg = strMsg & vbCrLf & "——" & mstrKind & "：" & .FocusedRow.Record(mCol.名称).Value
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            Err = 0: On Error GoTo ErrHand
            
            '新网RIS接口，检查部位删除时，删除对应的诊疗项目部位；启用参数，接口部件有效的前提下
            If mblnPACSInterface = True Then
                If Not gobjRIS Is Nothing Then
                    '传入部位类型和部位名称
                    If gobjRIS.HISBasicDictTable(RISBaseItemType.ClinicItemPart, RISBaseItemOper.Delete, mstrKind & "|" & Me.rptList.FocusedRow.Record(mCol.名称).Value) <> 1 Then
                        '出错时提示接口错误信息
                        If gobjRIS.LastErrorInfo <> "" Then
                            MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                        Else
                            MsgBox "调用RIS接口错误，不能继续当前操作！请与系统管理员联系", vbInformation, gstrSysName
                        End If
                        
                        Exit Sub
                    End If
                    
                    blnRisTrans = True
                Else
                   '接口部件无效时禁止并提示
                    MsgBox "RIS接口创建失败，不能继续当前操作！可能是接口文件安装或注册不正常，请与系统管理员联系。", vbInformation, gstrSysName
                    
                    Exit Sub
                End If
            End If
            
            gstrSql = "Zl_诊疗检查部位_Edit(3,'" & mstrKind & "','" & mstrPart & "')"
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)

            mstrPart = "": lngCount = .FocusedRow.Index
            If .Rows.Count > lngCount + 1 Then
                lngCount = lngCount + 1
            ElseIf lngCount > 0 Then
                lngCount = lngCount - 1
            End If
            If .Rows(lngCount).GroupRow = False Then mstrPart = .Rows(lngCount - 1).Record(mCol.编码).Value
            Call zlRefList(mstrPart)
        End With
    Case conMenu_Edit_MarkMap
        If Me.rptList.FocusedRow.GroupRow Then Exit Sub
        mstrCopy = Trim(Me.rptList.FocusedRow.Record(mCol.方法).Value)
    Case conMenu_Edit_Import
        If Me.rptList.FocusedRow.GroupRow Then Exit Sub

        Err = 0: On Error GoTo ErrHand
        
        '新网RIS接口，检查部位修改时，先删除原部位对应的诊疗项目部位；启用参数，接口部件有效的前提下
        '放到HIS执行过程之前
        If mblnPACSInterface = True Then
            If Not gobjRIS Is Nothing Then
                '传入部位类型和原部位名称
                If gobjRIS.HISBasicDictTable(RISBaseItemType.ClinicItemPart, RISBaseItemOper.Delete, mstrKind & "|" & Me.rptList.FocusedRow.Record(mCol.名称).Value) <> 1 Then
                    '出错时提示接口错误信息
                    If gobjRIS.LastErrorInfo <> "" Then
                        MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                    Else
                        MsgBox "调用RIS接口错误，不能继续当前操作！请与系统管理员联系", vbInformation, gstrSysName
                    End If
                    
                    Exit Sub
                End If
                    
                blnRisTrans = True
            Else
               '接口部件无效时禁止并提示
                MsgBox "RIS接口创建失败，不能继续当前操作！可能是接口文件安装或注册不正常，请与系统管理员联系。", vbInformation, gstrSysName
                
                Exit Sub
            End If
        End If
        
        gcnOracle.BeginTrans
        blnTrans = True
        
        With Me.rptList.FocusedRow
            gstrSql = "Zl_诊疗检查部位_Edit(2,'" & mstrKind & "','" & .Record(mCol.编码).Value & "','" & .Record(mCol.编码).Value & "'"
            gstrSql = gstrSql & ",'" & .Record(mCol.名称).Value & "','" & .Record(mCol.分组).Value & "','" & .Record(mCol.备注).Value & "','" & mstrCopy & "')"
        End With
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        
        '新网RIS接口，检查部位修改时，在删除了原部位对应的诊疗项目部位的前提下再增加新部位对应的方法；启用参数，接口部件有效的前提下
        '放到HIS执行过程之后
        If mblnPACSInterface = True Then
            If Not gobjRIS Is Nothing Then
                '传入部位类型和新部位名称
                If gobjRIS.HISBasicDictTable(RISBaseItemType.ClinicItemPart, RISBaseItemOper.AddNew, mstrKind & "|" & Me.rptList.FocusedRow.Record(mCol.名称).Value) <> 1 Then
                    gcnOracle.RollbackTrans
                    
                    '出错时提示接口错误信息
                    If gobjRIS.LastErrorInfo <> "" Then
                        MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                    Else
                        MsgBox "调用RIS接口错误，不能继续当前操作！请与系统管理员联系", vbInformation, gstrSysName
                    End If
                    
                    Exit Sub
                End If
                    
                blnRisTrans = True
            Else
                gcnOracle.RollbackTrans
                
               '接口部件无效时禁止并提示
                MsgBox "RIS接口创建失败，不能继续当前操作！可能是接口文件安装或注册不正常，请与系统管理员联系。", vbInformation, gstrSysName
                
                Exit Sub
            End If
        End If
        
        gcnOracle.CommitTrans
        blnTrans = False
        blnRisTrans = False
        
        Call zlRefList(mstrPart)
        
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh
        Call zlRefList(mstrPart)
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    End Select
    Exit Sub

ErrHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    
    'Ris接口和HIS不同步时，写错误日志
    If blnRisTrans = True And Not gobjRIS Is Nothing Then
        MsgBox "HIS新增或删除检查部位错误，RIS接口和HIS数据不同步，请与系统管理员联系。", vbInformation, gstrSysName
        
        On Error Resume Next
        Call gobjRIS.WriteCommLog("frmCheckPartList：cbsThis_Execute", "HIS新增或删除检查部位错误，RIS接口和HIS数据不同步", "类型=" & mstrKind & " " & "部位名称=" & Me.rptList.FocusedRow.Record(mCol.名称).Value, 0)
    End If
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.rptList.Records.Count <> 0 And mintEditState = 0)
    Case conMenu_Edit_Save, conMenu_Edit_Untread
        Control.Enabled = (mintEditState <> 0)
    Case conMenu_Edit_NewItem
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0)
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0)
        If Control.Enabled Then Control.Enabled = (mstrPart <> "")
        If Control.Enabled Then Control.Enabled = Not Me.rptList.FocusedRow.GroupRow
    Case conMenu_Edit_MarkMap
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0)
        If Control.Enabled Then Control.Enabled = (mstrPart <> "")
        If Control.Enabled Then Control.Enabled = Not Me.rptList.FocusedRow.GroupRow
        If Control.Enabled Then Control.Enabled = (Trim(Me.rptList.FocusedRow.Record(mCol.方法).Value) <> "")
    Case conMenu_Edit_Import
        Control.Enabled = (InStr(1, mstrPrivs, "增删改") > 0 And mintEditState = 0)
        If Control.Enabled Then Control.Enabled = (mstrPart <> "")
        If Control.Enabled Then Control.Enabled = Not Me.rptList.FocusedRow.GroupRow
        If Control.Enabled Then Control.Enabled = (mstrCopy <> "")
    
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Find, conMenu_View_Refresh, conMenu_View_Option: Control.Enabled = (mintEditState = 0)
    End Select
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_List
        Item.Handle = Me.picList.hWnd
    Case conPane_Edit
        If mfrmEdit Is Nothing Then Set mfrmEdit = New frmCheckPartEdit
        Item.Handle = mfrmEdit.hWnd
    End Select
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs
    mLngEditWidth = frmCheckPartEdit.ScaleWidth
    mintEditState = 0
    mstrPart = ""
    mstrCopy = ""
    mblnPACSInterface = (Val(zlDatabase.GetPara(255, glngSys, , "0")) = 1)
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
'    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "方法复制(&C)"): cbrControl.BeginGroup = True: cbrControl.Style = xtpButtonCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Import, "方法粘贴(&V)"): cbrControl.Style = xtpButtonCaption
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "检查类型")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "检查类型")
    cbrCustom.Handle = Me.cboKind.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Untread
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F3, conMenu_Edit_MarkMap
        .Add 0, VK_F4, conMenu_Edit_Import
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '设置词句显示停靠窗格
    Dim panThis As Pane
    If mfrmEdit Is Nothing Then Set mfrmEdit = New frmCheckPartEdit
    
    Set panThis = dkpMan.CreatePane(conPane_List, 450, 580, DockLeftOf, Nothing)
    panThis.Title = "部位列表"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set panThis = dkpMan.CreatePane(conPane_Edit, 550, 580, DockRightOf, Nothing)
    panThis.Title = "部位编辑"
    panThis.Options = PaneNoCaption

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    With Me.rptList
        .AutoColumnSizing = (Screen.Width / Screen.TwipsPerPixelX > 800)   '必须在列设置之前设置，才能生效
        Set rptCol = .Columns.Add(mCol.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.分组, "分组", 120, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.编码, "编码", 50, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.名称, "名称", 120, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.方法, "方法", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.组织, "组织", 150, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.备注, "备注", 100, True): rptCol.Editable = False: rptCol.Groupable = False
        
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列性质..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
    
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    '数据装入
    If RefKindList() = 0 Then
        MsgBox "诊疗检查类型字典未初始化！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    '初始化新网RIS接口
    If mblnPACSInterface Then
        Call IniRIS
    End If
    
End Sub

Private Sub Form_Resize()
    Dim panThis As Pane
    If Me.WindowState = vbMinimized Then Exit Sub
    Set panThis = Me.dkpMan.FindPane(conPane_Edit)
    panThis.MinTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, panThis.MinTrackSize.Height
    panThis.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, panThis.MaxTrackSize.Height
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
    panThis.MinTrackSize.SetSize 0, panThis.MinTrackSize.Height
    panThis.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, panThis.MaxTrackSize.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmEdit
    Set mfrmEdit = Nothing
    Call SaveWinState(Me, App.ProductName)
    
    If Not gobjRIS Is Nothing Then
        Set gobjRIS = Nothing
    End If
End Sub

Private Sub picList_Resize()
    With Me.rptList
        .Left = Me.picList.ScaleLeft: .Width = Me.picList.ScaleWidth - .Left
        .Top = Me.picList.ScaleTop: .Height = Me.picList.ScaleHeight - .Top
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.GroupRow Then Exit Sub
    Call rptList_RowDblClick(Me.rptList.FocusedRow, Me.rptList.FocusedRow.Record.Item(mCol.编码))
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    
    If Button <> vbRightButton Then Exit Sub
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If mstrPart = "" Then Exit Sub
    
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)

End Sub

Private Sub rptList_SelectionChanged()
    With Me.rptList
        If .FocusedRow Is Nothing Then
            mstrPart = ""
        ElseIf .FocusedRow.GroupRow = True Then
            mstrPart = ""
        Else
            mstrPart = .FocusedRow.Record.Item(mCol.编码).Value
        End If
        Call mfrmEdit.zlRefresh(mstrKind, mstrPart)
    End With
End Sub


