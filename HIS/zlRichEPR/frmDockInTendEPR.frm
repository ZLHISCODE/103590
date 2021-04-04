VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockInTendEPR 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3420
      Index           =   0
      Left            =   1035
      ScaleHeight     =   3420
      ScaleWidth      =   4650
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   390
      Width           =   4650
      Begin VB.Frame fraColSel 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   195
         Begin VB.Image imgColSel 
            Height          =   195
            Left            =   0
            Picture         =   "frmDockInTendEPR.frx":0000
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsColumn 
         Height          =   3480
         Left            =   135
         TabIndex        =   1
         Top             =   945
         Visible         =   0   'False
         Width           =   1470
         _cx             =   2593
         _cy             =   6138
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
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   14737632
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDockInTendEPR.frx":054E
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
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
      Begin VSFlex8Ctl.VSFlexGrid vfgWrit 
         Height          =   2310
         Left            =   45
         TabIndex        =   3
         Top             =   75
         Width           =   3735
         _cx             =   6588
         _cy             =   4075
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
         Cols            =   18
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
         Begin VB.PictureBox picInfo 
            BackColor       =   &H00FFEBD7&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   0
            Picture         =   "frmDockInTendEPR.frx":059C
            ScaleHeight     =   225
            ScaleMode       =   0  'User
            ScaleWidth      =   283.333
            TabIndex        =   4
            Top             =   0
            Width           =   250
         End
         Begin MSComctlLib.ImageList imgWrit 
            Left            =   1860
            Top             =   1005
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInTendEPR.frx":6DEE
                  Key             =   "书写"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInTendEPR.frx":7388
                  Key             =   "修订"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInTendEPR.frx":7922
                  Key             =   "归档"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInTendEPR.frx":7EBC
                  Key             =   "转交"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInTendEPR.frx":8256
                  Key             =   "打印"
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   720
      Top             =   -15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   195
      Top             =   540
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmDockInTendEPR.frx":EAB8
      Left            =   120
      Top             =   165
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDockInTendEPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private Enum mCol
    w标志 = 0: wID: w页面编号: w页面名称: w病历名称: w创建人: w创建时间: w保存人: w完成时间: w当前版本: w签名级别: w当前情况: w归档人: w归档日期: w病区ID: w病区名: w病人状态: w编辑方式: w婴儿: w打印
End Enum

Private mstrColWidthConfig As String
'
Private mstrPrivs As String                             '当前使用者对本程序(1255)的权限串
Private mblnSearch As Boolean                           '当前使用者是否具备病历检索(1273)权
Private mlngPatiId As Long                              '病人id
Private mlngPageId As Long                              '主页id
Private mlngDeptId As Long                              '当前操作科室id，如病人科室和当前科室不一致，则不能操作归档外的功能
Private mblnEdit As Boolean                             '是否允许操作，通常由上级程序根据当前操作科室是否当前病人病区决定。
Private mblnDoctorStation As Boolean                    '是否医生站调用
Private mblnMoved As Boolean                            '是否转储
Private mblnInsideTools As Boolean
Private WithEvents mfrmNew As frmDockEPRNew
Attribute mfrmNew.VB_VarHelpID = -1
Private WithEvents mfrmContent As frmDockEPRContent
Attribute mfrmContent.VB_VarHelpID = -1
Private mfrmMonitor As New frmDockEPRMonitor
Private mObjTabEpr As cTableEPR
Attribute mObjTabEpr.VB_VarHelpID = -1
Private mObjTabEprView As cTableEPR
Public Event Activate()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Private mfrmTipInfo As New frmTipInfo
Private mblnViewTag As Boolean   'vfgWrit行列变换事件执行标志，true正在执行，false没有执行
Private mblnViewNow As Boolean  'vfgWrit双击事件标志，ture正在执行，false没有执行
Private mlngCurId As Long
Public Function GetFormOperation() As String
'记录界面选定信息，因为工作站在切换页卡时是释放了对象，换回来时重新初始化刷新的。
    GetFormOperation = mlngCurId
End Function

Public Sub RestoreFormOperation(ByVal strValue As String)
'恢复界面选定信息，工作站在刷新之前调用
    mlngCurId = Val(strValue)
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
'-0-小(缺省)，1-大
Dim bytFontSize As Byte

    bytFontSize = Decode(bytSize, 0, 9, 1, 12, bytSize)
    Call mPublic.SetFontSize(Me, bytFontSize)
    Call mPublic.SetFontSize(mfrmNew, bytFontSize)
End Sub
Public Function InitData(ByVal strPrivs As String) As Boolean
    mstrPrivs = strPrivs
End Function

Public Function RefreshData(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long, ByVal blnDoctorStation As Boolean, _
                            ByVal blnEdit As Boolean, Optional ByVal blnForce As Boolean, Optional ByVal blnMoved As Boolean) As Boolean
    '******************************************************************************************************************
    '功能：刷新数据
    '参数：
    '返回：
    '******************************************************************************************************************
    If mlngPatiId = lngPatiID And mlngPageId = lngPageId And blnForce = False Then Exit Function '非强制刷新，两次相同不刷新
    
    If mlngDeptId <> lngDeptId Or gstrESign = "" Then '提取是否本部门启用电子签名,科室变更或没取过时提取
        gstrESign = getPassESign(4, lngDeptId)
    End If
    
    mlngPatiId = lngPatiID: mlngPageId = lngPageId: mblnEdit = blnEdit: mlngDeptId = lngDeptId
    
    mblnDoctorStation = blnDoctorStation: mblnMoved = blnMoved
    Call zlRefWrit
    
End Function
Public Sub zlDefCommandBars(ByVal cbsThis As Object, ByVal blnInsideTools As Boolean)
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

    mblnInsideTools = blnInsideTools
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    '文件菜单
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '特殊情况:放在第一个
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "打开(&O)…", 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
        
        '放在输出到Excel之后
        Set cbrControl = .Find(, conMenu_File_Excel)
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "导出为XML文件(&L)…", cbrControl.Index + 1)
    End With

    '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "病历(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审阅(&U)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10 + 1, "归档(&I)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "验证签名(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Sort, "病历排序(&S)"): cbrControl.BeginGroup = True
    End With

    '工具菜单:主窗体可能没有,放在帮助菜单前面
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", cbrMenuBar.Index, False)
        cbrMenuBar.ID = conMenu_ToolPopup
    End If
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Monitor, "病历质量监测(&M)"): cbrControl.BeginGroup = True
    End With
    
    '工具栏定义
    cbrMain.DeleteAll
    If mblnInsideTools Then
        Set cbrToolBar = cbrMain.Add("工具栏", xtpBarTop)
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        cbrToolBar.ContextMenuPresent = False
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "打开"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审阅"): cbrControl.STYLE = xtpButtonIconAndCaption
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10 + 1, "归档"): cbrControl.STYLE = xtpButtonIconAndCaption
        End With
    Else
        Set cbrToolBar = cbsThis(2)
        For Each cbrControl In cbrToolBar.Controls '先求出前面的最后一个Control
            If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
                Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
            End If
        Next
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增", cbrControl.Index + 1): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审阅", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10 + 1, "归档", cbrControl.Index + 1)
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "打开", 1)
            .Item(cbrControl.Index + 1).BeginGroup = True
        End With
    End If
    
    '命令的快键绑定
    '-----------------------------------------------------
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("O"), conMenu_File_Open
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("U"), conMenu_Edit_Audit
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
    End With

    '-----------------------------------------------------
    '根据权限状态，显示增加窗格
    '-----------------------------------------------------
    If (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "护理病历书写") > 0) Then
        Me.dkpMain.Panes(3).Select
        Call mfrmNew.zlRefList(3, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs)
    End If
End Sub
Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim strInfo As String, lFileId As Long
Dim bFinded As Boolean, frmThis As Form, bEditor As Byte
    If mblnMoved And (Control.ID = conMenu_Edit_Modify Or Control.ID = conMenu_Edit_Delete Or _
                        Control.ID = conMenu_Edit_Audit Or Control.ID = conMenu_Edit_Archive * 10 + 1 Or _
                        Control.ID = conMenu_File_Open Or Control.ID = conMenu_File_ExportToXML) Then  '已转储病人,修改,删除,审核,归档,打开不允许操作
        MsgBox "该病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
                        "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    lFileId = Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID))
    bEditor = Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.w编辑方式))
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Open        '病历阅读
        If GetCurrentGdi > 8000 Then Call MsgBox("当前系统资源占用过多，请先关闭一些病历编辑窗口后再重试！", vbInformation, gstrSysName): Exit Sub
        
        If bEditor = 0 Then
            Dim fViewDoc As New frmEPRView, blnCanPrint As Boolean
            blnCanPrint = (InStr(1, mstrPrivs, "护理病历打印") > 0) And (Trim(vfgWrit.TextMatrix(vfgWrit.Row, mCol.w归档人)) = "" Or InStr(1, mstrPrivs, "归档病历输出") > 0)
            fViewDoc.ShowMe Me, lFileId, , blnCanPrint
        Else
            If Not mObjTabEprView Is Nothing Then
                bFinded = mObjTabEprView.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_住院, mlngDeptId)
            End If
            If Not bFinded Then
                mObjTabEprView.InitOpenEPR Me, cprEM_修改, cprET_单病历编辑, lFileId, True, 0, cprPF_住院, mlngPatiId, mlngPageId, , mlngDeptId, 0, mstrPrivs, , InStr(mstrPrivs, "护理病历打印") > 0, Val(gstrESign)
            End If
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview
        If GetCurrentGdi > 8000 Then Call MsgBox("当前系统资源占用过多，请先关闭一些病历编辑窗口后再重试！", vbInformation, gstrSysName): Exit Sub
        
        If EprPrinted(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) And InStr(mstrPrivs, "取消打印") = 0 Then '已经打印过且没有取消打印权限,不允许重复打印
            MsgBox "当前病历已打印，不允许重复打印！", vbInformation, gstrSysName
            Exit Sub
        End If
        Call zlEPRPrint(True)
        Call zlRefWrit
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print
        If GetCurrentGdi > 8000 Then Call MsgBox("当前系统资源占用过多，请先关闭一些病历编辑窗口后再重试！", vbInformation, gstrSysName): Exit Sub
        
        If EprPrinted(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) And InStr(mstrPrivs, "取消打印") = 0 Then '已经打印过且没有取消打印权限,不允许重复打印
            MsgBox "当前病历已打印，不允许重复打印！", vbInformation, gstrSysName
            Exit Sub
        End If
        Call zlEPRPrint(False)
        Call zlRefWrit
    Case conMenu_Edit_NoPrint '取消打印标记
        If Split(EprIsCommit, "|")(0) = 0 Then
            MsgBox "该病人病案已提交审查，不能撤消打印，请取消审查后再试！", vbInformation, gstrSysName
            Exit Sub
        End If
        Call PrintCancel(CLng(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)))
        Call zlRefWrit
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_ExportToXML
        If GetCurrentGdi > 8000 Then Call MsgBox("当前系统资源占用过多，请先关闭一些病历编辑窗口后再重试！", vbInformation, gstrSysName): Exit Sub
        
        '导出到XML文件
        Dim strF As String
        dlgThis.Filename = "病历_" & Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.w病历名称) & _
            "(" & Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.wID) & ").xml"
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        On Error Resume Next
        dlgThis.ShowSave
        If Err.Number <> 0 Then Err.Clear: Exit Sub
        strF = dlgThis.Filename
        On Error GoTo errHand
        If gobjFSO.FileExists(strF) Then
            DoEvents
            If MsgBox("该文件已经存在，是否覆盖？", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
        End If
        
        If bEditor = 1 Then
                '表格式病历
            mObjTabEprView.InitOpenEPR Me, cprEM_修改, cprET_单病历编辑, lFileId, False, 0, cprPF_住院, _
                    mlngPatiId, mlngPageId, , mlngDeptId, 0, mstrPrivs
            If mObjTabEprView.zlExportXML(strF) Then
                MsgBox "成功导出为XML文件！" & vbCrLf & "文件名:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        Else
            Dim DocXML As New cEPRDocument '普通住院病历
            DocXML.InitAndOpenEPR Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.wID), 0, , True
            If DocXML.ExportToXMLFile(DocXML.frmEditor.Editor1, strF) Then
                DoEvents
                MsgBox "成功导出为XML文件！" & vbCrLf & "文件名:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem

        dkpMain.Panes(3).Select
        Call mfrmNew.zlRefList(3, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify                    '修改护理记录数据
        If GetCurrentGdi > 8000 Then Call MsgBox("当前系统资源占用过多，请先关闭一些病历编辑窗口后再重试！", vbInformation, gstrSysName): Exit Sub
        
        If EprPrinted(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) Then MsgBox "当前病历已打印，不允许操作，若确需再次操作请取消打印后再进行！", vbInformation, gstrSysName: Exit Sub
        lFileId = CLng(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID))
        If vfgWrit.TextMatrix(vfgWrit.Row, mCol.w编辑方式) = 1 Then
            '表格式病历
            If Not mObjTabEpr Is Nothing Then
                bFinded = mObjTabEpr.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_住院, mlngDeptId)
            End If
            If bFinded = False Then
                Set mObjTabEpr = New cTableEPR
                mObjTabEpr.InitOpenEPR Me, cprEM_修改, cprET_单病历编辑, lFileId, True, 0, cprPF_住院, _
                    mlngPatiId, mlngPageId, , mlngDeptId, 0, mstrPrivs, , InStr(mstrPrivs, "护理病历打印") > 0, Val(gstrESign)
                    mObjTabEpr.EPRPatiRecInfo.婴儿 = CLng(Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.w婴儿)))
            End If
        Else
            '单病历编辑模式
            Dim Doc As New cEPRDocument
            With Me.vfgWrit
                Doc.InitEPRDoc cprEM_修改, cprET_单病历编辑, .TextMatrix(.Row, mCol.wID), cprPF_住院, mlngPatiId, CStr(mlngPageId)
                Doc.EPRPatiRecInfo.婴儿 = CLng(Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.w婴儿)))
                Doc.ShowEPREditor Me
            End With
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        If Split(EprIsCommit, "|")(1) = 0 Then
            MsgBox "该病人病案已提交审查，不能删除，请取消审查后再试！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If EprPrinted(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) Then MsgBox "当前病历已打印，不允许操作，若确需再次操作请取消打印后再进行！", vbInformation, gstrSysName: Exit Sub
        With Me.vfgWrit
            strInfo = "真的删除这份“" & .TextMatrix(.Row, mCol.w病历名称) & "”吗？"
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSQL = "Zl_电子病历记录_Delete(" & .TextMatrix(.Row, mCol.wID) & ")"
            Err = 0: On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Err = 0: On Error GoTo 0
            Call zlRefWrit
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Audit
        If GetCurrentGdi > 8000 Then Call MsgBox("当前系统资源占用过多，请先关闭一些病历编辑窗口后再重试！", vbInformation, gstrSysName): Exit Sub
        
        If EprPrinted(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) Then MsgBox "当前病历已打印，不允许操作，若确需再次操作请取消打印后再进行！", vbInformation, gstrSysName: Exit Sub
        If bEditor = 1 Then
            '表格式病历
            If Not mObjTabEpr Is Nothing Then
                bFinded = mObjTabEpr.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_住院, mlngDeptId)
            End If
            If bFinded = False Then
                Set mObjTabEpr = New cTableEPR
                mObjTabEpr.InitOpenEPR Me, cprEM_修改, cprET_单病历审核, lFileId, True, 0, cprPF_住院, _
                    mlngPatiId, mlngPageId, , mlngDeptId, 0, mstrPrivs, , InStr(mstrPrivs, "护理病历打印") > 0, Val(gstrESign)
            End If
        Else
            '单病历审核模式
            Dim frmAudit As Form, bFindedAudit As Boolean
            For Each frmAudit In Forms
                If frmAudit.Name = "frmMain" Then
                    If frmAudit.Document.EPRPatiRecInfo.ID = Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, 1) _
                        And frmAudit.Document.EPRPatiRecInfo.病人来源 = cprPF_住院 And frmAudit.Document.EPRPatiRecInfo.病人ID = mlngPatiId _
                        And frmAudit.Document.EPRPatiRecInfo.主页ID = mlngPageId And frmAudit.ChildMode = False Then
                        frmAudit.Show
                        bFindedAudit = True
                    End If
                End If
            Next
            If bFindedAudit = False Then
                '首次审核
                Dim DocAudit As New cEPRDocument
                DocAudit.InitEPRDoc cprEM_修改, cprET_单病历审核, Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, 1), cprPF_住院, mlngPatiId, CStr(mlngPageId)
                DocAudit.ShowEPREditor Me
            End If
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Archive * 10 + 1

        With vfgWrit
            If Trim(.TextMatrix(.Row, mCol.w归档人)) = "" Then
                If Trim(.TextMatrix(.Row, mCol.w病人状态)) = "在院" Then
                    strInfo = "真的将该份“" & .TextMatrix(.Row, mCol.w病历名称) & "”归档吗？"
                    If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    gstrSQL = "Zl_电子病历记录_Archive(" & lFileId & ",0)"
                Else
                    strInfo = "病人已经" & Trim(.TextMatrix(.Row, mCol.w病人状态)) & "，要将病人本次住院全部护理病历归档吗？" & vbCrLf _
                            & "  选择“是”，归档病人本次全部护理病历；" & vbCrLf _
                            & "  选择“否”，仅归档该份“" & .TextMatrix(.Row, mCol.w病历名称) & "”。"
                    Select Case MsgBox(strInfo, vbQuestion + vbYesNoCancel + vbDefaultButton3, gstrSysName)
                    Case vbYes: gstrSQL = "Zl_电子病历记录_Archive(" & lFileId & ",0,1)"
                    Case vbNo: gstrSQL = "Zl_电子病历记录_Archive(" & lFileId & ",0)"
                    Case Else: Exit Sub
                    End Select
                End If
            Else
                strInfo = "需要撤销该病人本次住院所有已归档护理病历吗？" & vbCrLf _
                        & "  选择“是”，撤销该病人本次住院所有已归档护理病历；" & vbCrLf _
                        & "  选择“否”，仅撤消该份“" & .TextMatrix(.Row, mCol.w病历名称) & "”的归档。"
                Select Case MsgBox(strInfo, vbQuestion + vbYesNoCancel + vbDefaultButton3, gstrSysName)
                Case vbYes: gstrSQL = "Zl_电子病历记录_Archive(" & lFileId & ",1,1)"
                Case vbNo: gstrSQL = "Zl_电子病历记录_Archive(" & lFileId & ",1)"
                Case Else: Exit Sub
                End Select
            End If
            Err = 0: On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Err = 0: On Error GoTo 0
            Call zlRefWrit
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Sort
        '排序
        Dim frmSort As New frmEPRSort
        If frmSort.ShowMe(Me, mlngPatiId, mlngPageId, cpr护理病历, Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.w页面编号)) = True Then
            '刷新显示
            Call zlRefWrit
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Refresh

        Call zlRefWrit

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Monitor
        If mfrmMonitor.Visible = False Then mfrmMonitor.Show vbModeless, Me
        Call mfrmMonitor.zlRefList(mlngPatiId, mlngPageId, 4, mlngDeptId, 1, 1)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Search
        Call frmEPRSearchMan.ShowSearchClinic(Me, mlngDeptId)
    Case conMenu_Tool_SignVerify
        If bEditor = 0 Then
            Call VerifySignature(Me, lFileId, mblnMoved)
        Else '表格式病历，28未处理数字签名情况
            'call
        End If
    End Select
    
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
LL:
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    Dim lngCount As Long, blnFinished As Boolean, lngMaxVersion As Long, eSignLevel As EPRSignLevelEnum

    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Open
        Control.Visible = True
        Control.Enabled = (Val(Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.wID)) <> 0 And mblnEdit)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_ExportToXML

        Control.Enabled = (Val(Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.wID)) <> 0)
        Control.Enabled = (Val(Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.wID)) <> 0 And InStr(1, mstrPrivs, "护理病历打印") > 0)
        If Control.Enabled Then Control.Enabled = (Trim(Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.w归档人)) = "" Or InStr(1, mstrPrivs, "归档病历输出") > 0)

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NoPrint
        Control.Enabled = InStr(mstrPrivs, "取消打印") > 0 And (Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) <> 0)
        If Control.Enabled Then Control.Enabled = Trim(vfgWrit.TextMatrix(vfgWrit.Row, mCol.w打印)) <> ""
        If Control.Enabled Then Control.Enabled = mblnEdit
    Case conMenu_File_Excel

        Control.Enabled = (Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) <> 0)

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem
        Control.Enabled = (mblnEdit And mlngPatiId > 0)
        Control.Visible = (InStr(1, mstrPrivs, "护理病历书写") > 0 And mblnDoctorStation = False)
        If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "护理病历书写") > 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify
    
        Control.Enabled = (mblnEdit And mlngPatiId > 0)

        With Me.vfgWrit
            Control.Visible = (InStr(1, mstrPrivs, "护理病历书写") > 0 Or InStr(1, mstrPrivs, "他人护理病历") > 0) And mblnDoctorStation = False
            If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "护理病历书写") > 0)
            If Control.Enabled Then Control.Enabled = (mlngDeptId = Val(.TextMatrix(.Row, mCol.w病区ID)))   '本科病历才可以改
            If Control.Enabled Then
                If Trim(.TextMatrix(.Row, mCol.w完成时间)) = "" Then
                    Control.Enabled = (InStr(1, mstrPrivs, "他人护理病历") > 0 Or Trim(.TextMatrix(.Row, mCol.w创建人)) = Trim(gstrUserName))
                ElseIf Trim(.TextMatrix(.Row, mCol.w归档人)) = "" And Val(.TextMatrix(.Row, mCol.w当前版本)) <= 1 And InStr(1, ",1,2,4,", Val(.TextMatrix(.Row, mCol.w签名级别))) > 0 Then
                    Control.Enabled = (InStr(1, mstrPrivs, "他人护理病历") > 0 Or InStr(1, .TextMatrix(.Row, mCol.w保存人), Trim(gstrUserName)) > 0)
                Else
                    Control.Enabled = False
                End If
            End If
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        Control.Enabled = (mblnEdit And mlngPatiId > 0)

        With vfgWrit

            Control.Visible = (InStr(1, mstrPrivs, "强制删除病历") > 0 Or InStr(1, mstrPrivs, "护理病历书写") > 0 Or InStr(1, mstrPrivs, "他人护理病历") > 0) And mblnDoctorStation = False

            Control.Enabled = (Val(Me.vfgWrit.TextMatrix(Me.vfgWrit.Row, mCol.wID)) <> 0)
            If Control.Enabled And InStr(1, mstrPrivs, "强制删除病历") > 0 Then Exit Sub '具备强制删除权限，则不进行后续的判断
            If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "护理病历书写") > 0)
            If Control.Enabled Then Control.Enabled = (mlngDeptId = Val(.TextMatrix(.Row, mCol.w病区ID)))   '本科病历才可以删
            If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.w完成时间)) = "")        '未完成病历可以删
            If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "他人护理病历") > 0 Or Trim(.TextMatrix(.Row, mCol.w创建人)) = Trim(gstrUserName))
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Audit

        Control.Visible = (InStr(1, mstrPrivs, "护理病历审阅") > 0 And mblnDoctorStation = False)
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And Control.Visible)
        With vfgWrit
'                If Control.Enabled Then Control.Enabled = (mlngDeptId = Val(.TextMatrix(.Row, mCol.w病区ID)))   '本科病历才可以审
            If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.w完成时间)) <> "")       '完成病历才可以审
            If Control.Enabled Then Control.Enabled = (Trim(.TextMatrix(.Row, mCol.w归档人)) = "")          '未归档病历可以审
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Archive * 10 + 1

        Control.Visible = (InStr(1, mstrPrivs, "护理病历归档") > 0 And mblnDoctorStation = False)

        '只有已经完成的未归档的病历,才能进行归档操作
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And Control.Visible)
        With Me.vfgWrit
            If Control.Enabled Then Control.Enabled = (Val(.TextMatrix(.Row, mCol.w签名级别)) <> 0)         '当前版本已经签名完成才可以归档
            If Trim(.TextMatrix(.Row, mCol.w归档人)) = "" Then
                Control.Caption = "病历归档": Control.Checked = False
            Else
                Control.Caption = "病历撤档": Control.Checked = True
            End If
        End With

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_EditPopup

        Control.Visible = (mblnDoctorStation = False And (InStr(1, mstrPrivs, "护理病历归档") > 0 _
                                                        Or InStr(1, mstrPrivs, "护理病历审阅") > 0 _
                                                        Or InStr(1, mstrPrivs, "护理病历书写") > 0 _
                                                        Or InStr(1, mstrPrivs, "强制删除病历") > 0))
        Control.Enabled = Control.Visible
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Monitor
        Control.Visible = True
        Control.Enabled = (mlngPatiId > 0 And InStr(1, mstrPrivs, "护理病历监测") > 0)

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Search
        Control.Visible = True
        Control.Enabled = mblnSearch
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Sort
        '排序（只有多文档共用页面时才可以调整序号）
        Dim R1&, C1&, R2&, C2&
        vfgWrit.GetMergedRange vfgWrit.Row, mCol.w页面名称, R1, C1, R2, C2
        Control.Enabled = (R1 <> R2)
        Control.Visible = Control.Enabled
    Case conMenu_Tool_SignVerify
        Control.Enabled = Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) <> 0 And Trim(vfgWrit.TextMatrix(vfgWrit.Row, mCol.w完成时间)) <> ""
    End Select
End Sub

Public Sub RefreshList()
    Call zlRefWrit
End Sub
Private Function InitColumnSelect() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    On Error Resume Next
    '功能：根据原始列显示状态初始化列选择器
    Dim lngRow As Long, i As Long

    vsColumn.Rows = vsColumn.FixedRows
    With vfgWrit
        For i = .FixedCols To .Cols - 1
            Select Case i
            Case mCol.w病历名称, mCol.w创建人, mCol.w创建时间, mCol.w保存人, mCol.w完成时间, mCol.w当前情况, mCol.w病区名
                 vsColumn.Rows = vsColumn.Rows + 1
                 lngRow = vsColumn.Rows - 1
                 vsColumn.TextMatrix(lngRow, 1) = .TextMatrix(0, i)
                 vsColumn.RowData(lngRow) = i

                 '固定显示列
                 If InStr(",页面名称,病历名称,", "," & .TextMatrix(0, i) & ",") > 0 Then
                     vsColumn.TextMatrix(lngRow, 0) = 1
                     vsColumn.Cell(flexcpForeColor, lngRow, 0, lngRow, 1) = vsColumn.BackColorFixed
                 End If
            End Select
        Next
    End With
    vsColumn.Height = vsColumn.RowHeightMin * vsColumn.Rows + 130
    vsColumn.Row = 1

    InitColumnSelect = True

End Function

Private Sub zlEPRPrint(blnPreview As Boolean)
Dim lFileId As Long
Dim frmP As New frmPrintPreview, r As String, blnOrigMode As Boolean '是否显示原始状态
    If vfgWrit.TextMatrix(vfgWrit.Row, mCol.w编辑方式) = 0 Then
        r = zlCommFun.ShowMsgBox("病历预览/打印", "请选择病历预览/打印的格式？", "!最终格式(&F),原始格式(&O),取消(&C)", Nothing)
        If r = "最终格式" Then
            blnOrigMode = False
        ElseIf r = "原始格式" Then
            blnOrigMode = True
        Else
            Exit Sub
        End If
        frmP.DoMultiDocPreview Me, cpr护理病历, mlngPatiId, mlngPageId, cpr护理病历, Me.vfgWrit.Cell(flexcpText, Me.vfgWrit.Row, mCol.w页面编号), Me.vfgWrit.Cell(flexcpText, Me.vfgWrit.Row, mCol.wID), Not blnPreview, blnOrigMode, , mblnMoved
        Unload frmP 'ByZT:窗体Load了未显示，没有人为关闭的情况下VB不会自动Unload
        Set frmP = Nothing
    Else
        lFileId = CLng(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID))
        mObjTabEprView.InitOpenEPR Me, cprEM_修改, cprET_单病历审核, lFileId, False, 0, cprPF_住院, mlngPatiId, mlngPageId, , mlngDeptId, , mstrPrivs, mblnMoved, InStr(mstrPrivs, "护理病历打印") > 0
        mObjTabEprView.zlPrintDoc Me, blnPreview
    End If
End Sub

Private Sub zlRefWrit()
'---------------------------------------------
'护理病历刷新
'---------------------------------------------
Dim lngCurId As Long    '刷新前选中的病历记录ID
Dim lngCurRow As Long
Dim lngCol As Long
Dim lngRow As Long
Dim rsTemp As New ADODB.Recordset
    
    vsColumn.Visible = False
    vfgWrit.Tag = ""
    Call mfrmContent.Clear
    
    gstrSQL = "Select r.Id, f.编号, Decode(f.页面, Null, r.病历名称, f.页面) As 页面, r.病历名称, r.创建人 As 创建人," & _
            "        To_Char(r.创建时间, 'yyyy-mm-dd hh24:mi') As 创建时间, r.保存人," & _
            "        To_Char(r.完成时间, 'yyyy-mm-dd hh24:mi') As 完成时间, r.最后版本 As 当前版本, r.签名级别," & _
            "        Decode(r.最后版本, 1, '书写：', '修订：') || r.保存人 || '在' || To_Char(r.保存时间, 'yyyy-mm-dd hh24:mi') ||" & _
            "         Decode(Nvl(r.签名级别, 0), 0, '保存(未完成)', 1, '完成', '审签') As 当前情况, r.归档人, r.归档日期," & _
            "        r.科室id As 病区id, d.名称 As 病区, p.病人状态,r.编辑方式,r.婴儿,r.打印人 as 打印" & _
            " From 电子病历记录 r, 部门表 d," & _
            "      (Select Decode(出院日期, Null, Decode(状态, 3, '预出院', '在院'), '出院') As 病人状态" & vbNewLine & _
            "        From 病案主页" & vbNewLine & _
            "        Where 病人id = [1] And 主页id = [2]) p," & _
            "      (Select d.Id As 文件id, f.种类, f.编号, f.名称 As 页面, d.保留" & _
            "        From 病历文件列表 d, 病历页面格式 f" & _
            "        Where d.种类 = 4 And d.种类 = f.种类 And d.页面 = f.编号) f" & _
            " Where r.文件id = f.文件id(+) And r.病人来源 = 2 And r.病历种类 = 4 And r.科室id = d.Id And r.病人id = [1] And r.主页id = [2]" & _
            " Order By r.病历种类, f.编号, r.序号, r.创建时间"
    Err = 0: On Error GoTo errHand
    If mblnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)
    
    With Me.vfgWrit
        Err = 0: On Error Resume Next
        lngCurId = Val(.TextMatrix(.Row, mCol.wID))
        If lngCurId = 0 Then lngCurId = mlngCurId
        .Clear
        Set .DataSource = rsTemp

        .MergeCells = flexMergeFree: .MergeCellsFixed = flexMergeFree
        .MergeCol(mCol.w页面名称) = True

        Dim T As Variant, i As Long
        On Error Resume Next
        T = Split(mstrColWidthConfig, ";")
        If UBound(T) < 18 Then
            mstrColWidthConfig = "270;0;0;1200;2000;800;1600;800;0;800;0;3300;0;0;0;1200;0;0;0"
        Else
            For i = 0 To .Cols - 1
                .ColWidth(i) = T(i)
                .ColHidden(i) = (.ColWidth(i) = 0)
            Next
        End If
        .TextMatrix(0, mCol.w页面名称) = .TextMatrix(0, mCol.w病历名称)
        .MergeRow(0) = True
        For lngCol = .FixedCols To .Cols - 1
            .FixedAlignment(lngCol) = flexAlignCenterCenter
        Next
        For lngRow = .FixedRows To .Rows - 1
            .MergeRow(lngRow) = True
            If Trim(.TextMatrix(lngRow, mCol.w归档人)) <> "" Then
                Set .Cell(flexcpPicture, lngRow, mCol.w标志) = imgWrit.ListImages("归档").Picture
            ElseIf Val(.TextMatrix(lngRow, mCol.w当前版本)) <= 1 Then
                Set .Cell(flexcpPicture, lngRow, mCol.w标志) = imgWrit.ListImages("书写").Picture
            Else
                Set .Cell(flexcpPicture, lngRow, mCol.w标志) = imgWrit.ListImages("修订").Picture
            End If
            If Trim(.TextMatrix(lngRow, mCol.w打印)) <> "" Then
                Set .Cell(flexcpPicture, lngRow, mCol.w页面名称) = imgWrit.ListImages("打印").Picture
            End If
            If lngCurId = Val(.TextMatrix(lngRow, mCol.wID)) Then lngCurRow = lngRow
        Next
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        If lngCurRow = 0 Then
            .Row = 0 '促使vfgthis不选中任何行，不显示任何内容，仅当选中某行时才刷新
        Else
            .Row = lngCurRow
        End If
        Call vfgWrit_RowColChange
    End With

    Call InitColumnSelect '列选择器
    
    If (mblnEdit And mlngPatiId > 0 And InStr(1, mstrPrivs, "护理病历书写") > 0) Then
        Me.dkpMain.Panes(3).Select
        Call mfrmNew.zlRefList(3, mlngPatiId, mlngPageId, mlngDeptId, mstrPrivs)
    End If
    'vfgWrit.Cell(flexcpWidth, mCol.w打印) = 0
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    zlExecuteCommandBars Control
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    zlUpdateCommandBars Control
End Sub

'######################################################################################################################
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hwnd
    Case 2
        If mfrmContent Is Nothing Then Set mfrmContent = New frmDockEPRContent
        Item.Handle = mfrmContent.hwnd
    Case 3
        If mfrmNew Is Nothing Then Set mfrmNew = New frmDockEPRNew
        Item.Handle = mfrmNew.hwnd
    End Select
End Sub


Private Sub Form_Activate()
    On Error Resume Next
    If vsColumn.Visible Then
        vsColumn.SetFocus '列选择器
    Else
        If Me.vfgWrit.Visible Then Me.vfgWrit.SetFocus
    End If
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    vsColumn.Visible = False '列选择器
End Sub

Private Sub Form_Load()
 Dim objPane As Pane, lngFontSize As Long
    On Error GoTo errHand
    
    mblnSearch = (InStr(1, GetPrivFunc(glngSys, 1273), "基本") > 0)

    mstrColWidthConfig = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ColWidthConfig", _
        "270;0;0;1200;2000;800;1600;800;0;800;0;3300;0;0;0;1200;0;0;0")
    
    lngFontSize = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name & "\" & vfgWrit.Name, "FontSize", 9)
    vfgWrit.FontSize = lngFontSize
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    With Me.cbrMain
        .VisualTheme = xtpThemeOffice2003
        Set .Icons = zlCommFun.GetPubIcons
        .ActiveMenuBar.Visible = False
        .EnableCustomization False
        With .Options
            .ShowExpandButtonAlways = False
            .ToolBarAccelTips = True
            .AlwaysShowFullMenus = False
            .IconsWithShadow = True '放在VisualTheme后有效
            .UseDisabledIcons = True
            .LargeIcons = True
            .SetIconSize True, 24, 24
        End With
    End With
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
    dkpMain.SetCommandBars cbrMain
    
    Set objPane = dkpMain.CreatePane(1, 100, 100, DockTopOf, Nothing): objPane.Title = "病历列表": objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(2, 100, 500, DockBottomOf, objPane): objPane.Title = "病历预览": objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(3, 100, 100, DockRightOf, Nothing): objPane.Title = "新建病历": objPane.Options = PaneNoCaption

    Set mObjTabEprView = New cTableEPR
    If mfrmContent Is Nothing Then Set mfrmContent = New frmDockEPRContent
    mObjTabEprView.InitTableEPR gcnOracle, glngSys, gstrDbOwner
    Call RestoreWinState(Me, App.ProductName)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim strCols As String, i As Long
    On Error Resume Next
    For i = 0 To vfgWrit.Cols - 1
        strCols = strCols & IIf(i = 0, "", ";") & vfgWrit.ColWidth(i)
    Next

    mstrColWidthConfig = strCols
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "ColWidthConfig", mstrColWidthConfig
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name & "\" & vfgWrit.Name, "FontSize", vfgWrit.FontSize
    If Not mfrmContent Is Nothing Then Unload mfrmContent
    If Not mfrmNew Is Nothing Then Unload mfrmNew
    If Not mfrmMonitor Is Nothing Then Unload mfrmMonitor
    If Not mfrmTipInfo Is Nothing Then Unload mfrmTipInfo
    Set mfrmContent = Nothing
    Set mfrmNew = Nothing
    Set mfrmMonitor = Nothing
    Set mObjTabEpr = Nothing
    Set mObjTabEprView = Nothing
    Set mfrmTipInfo = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mfrmNew_NewClick(ByVal FileId As Long, ByVal babyNum As Long)
Dim frmThis As Form, bFinded As Boolean, strTmp As String
Dim rs As New ADODB.Recordset, strSQL As String
    
    If GetCurrentGdi > 8000 Then Call MsgBox("当前系统资源占用过多，请先关闭一些病历编辑窗口后再重试！", vbInformation, gstrSysName): Exit Sub
        
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        If Not gobjPlugIn.AddEMRBefore(glngSys, 1255, mlngPatiId, mlngPageId, FileId) Then Exit Sub
        Err.Clear: On Error GoTo 0
    End If
    
    On Error GoTo errHand
    If gstrPrivsEpr = ";;" Then
        MsgBox "您不具备病历编辑相应权限，请与系统管理员联系。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If Split(EprIsCommit, "|")(0) = 0 Then
        MsgBox "该病人病案已提交审查，不能新增病历，请取消审查后再试！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If TimeLimitOut Then Exit Sub
    
    strSQL = "Select 保留 From 病历文件列表 Where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, FileId)
    If rs!保留 < 0 Then
        '特殊病历，手术麻醉单
        Exit Sub
    ElseIf rs!保留 = 2 Then '表格式编辑器
        If Not mObjTabEpr Is Nothing Then
            bFinded = mObjTabEpr.Showfrm(FileId, mlngPatiId, mlngPageId, cprPF_住院, mlngDeptId)
        End If
        If Not bFinded Then
            Set mObjTabEpr = New cTableEPR
            mObjTabEpr.InitOpenEPR Me, cprEM_新增, cprET_单病历编辑, FileId, True, 0, cprPF_住院, mlngPatiId, mlngPageId, , mlngDeptId, 0, mstrPrivs, , InStr(mstrPrivs, "护理病历打印") > 0, Val(gstrESign)
            dkpMain.Panes(3).Close
        End If
    Else
        'RichEPR病历
        '判断共享文档是否已经书写过
        gstrSQL = "Select ID From 病历文件列表 Where 编号 <> NVL(页面,编号) And ID =[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, FileId)
        If rs.EOF = False Then '是共享文档
            gstrSQL = "Select M.ID,M.名称" & vbNewLine & _
                        "       From 病历文件列表 L, 病历文件列表 M" & vbNewLine & _
                        "       Where M.种类 = L.种类 And M.编号 = L.页面 And L.ID =[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, FileId)
            If rs.EOF Then MsgBox "该病历的共享病历定义失效，请联系系统管理员。", vbInformation, gstrSysName: Exit Sub
            strTmp = rs!ID & "|" & rs!名称
            gstrSQL = "Select ID" & vbNewLine & _
                        "From 电子病历记录" & vbNewLine & _
                        "Where 病人id = [1] And 主页id =[2] And 文件id+0 =[3]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, Val(Split(strTmp, "|")(0)))
            If rs.EOF Then
                MsgBox "该病历的共享病历 [" & Split(strTmp, "|")(1) & "] 尚未书写，请检查。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    
        For Each frmThis In Forms
            If frmThis.Name = "frmMain" Then
                With frmThis.Document
                    If .EPRFileInfo.ID = FileId And .EPRPatiRecInfo.病人ID = mlngPatiId _
                        And .EPRPatiRecInfo.病人来源 = cprPF_住院 And .EPRPatiRecInfo.主页ID = mlngPageId _
                        And .EPRPatiRecInfo.科室ID = mlngDeptId And frmThis.ChildMode = False Then
                        frmThis.Show
                        bFinded = True
                    End If
                End With
            End If
        Next
        If bFinded = False Then
            Dim Doc As New cEPRDocument
            
            Doc.InitEPRDoc cprEM_新增, cprET_单病历编辑, FileId, cprPF_住院, mlngPatiId, CStr(mlngPageId), , mlngDeptId
            Doc.EPRPatiRecInfo.婴儿 = babyNum
            Doc.ShowEPREditor Me
            
            dkpMain.Panes(3).Close
        End If
    End If
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'显示指定病历列表行的历史签名记录
Dim strTipInfo As String, lngRow As Long, strPrint As String
    If picInfo.Visible = False Then Exit Sub
    lngRow = vfgWrit.MouseRow
    If lngRow <= 0 Then Exit Sub
    
    strTipInfo = vfgWrit.Cell(flexcpData, lngRow, mCol.w当前情况)
    If strTipInfo = "" Then '如果没有获取过，则立即获取并记录在列表中
        strTipInfo = GetEprSign(vfgWrit.TextMatrix(lngRow, mCol.wID))   '提取签名
        Call EprPrinted(vfgWrit.TextMatrix(lngRow, mCol.wID), strPrint) '提取打印记录
        strTipInfo = "由 " & Rpad(vfgWrit.TextMatrix(lngRow, mCol.w创建人), 8) & _
                     "于 " & Rpad(vfgWrit.TextMatrix(lngRow, mCol.w创建时间), 19) & " 创建" & vbCrLf & strTipInfo
        strTipInfo = strTipInfo & vbCrLf & strPrint
        vfgWrit.Cell(flexcpData, lngRow, mCol.w当前情况) = strTipInfo
    End If
    
    mfrmTipInfo.ShowTipInfo picInfo.hwnd, strTipInfo, True
End Sub
Private Function GetEprSign(ByVal lngFileID As Long)
'提取病历历史签名记录
Dim rsTemp As ADODB.Recordset, strSign As String
    gstrSQL = "Select 开始版 As 版本, Decode(要素表示, 3, '主任医师', 2, '主治医师', '经治医师') || '身份' || Decode(开始版, 1, '签名', '修订') As 操作," & vbNewLine & _
                "       Decode(Nvl(Instr(内容文本, ';'), 0), 0, 内容文本, Substr(内容文本, 1, Instr(内容文本, ';') - 1)) As 人员," & vbNewLine & _
                "       RTrim(Substr(对象属性, Instr(对象属性, ';', 1, 4) + 1)) As 时间" & vbNewLine & _
                "From 电子病历内容" & vbNewLine & _
                "Where 文件id = [1] And 对象类型 = 8 Order By 对象标记"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取签名记录", lngFileID)
    Do Until rsTemp.EOF
        strSign = strSign & "由 " & Rpad(NVL(rsTemp!人员), 8) & "于 " & Rpad(NVL(rsTemp!时间), 19) & " 以" & NVL(rsTemp!操作) & vbCrLf
        rsTemp.MoveNext
    Loop
    GetEprSign = strSign
End Function
Private Function EprPrinted(ByVal lngRecordId As Long, Optional strPrintInfo As String) As Boolean
'检查当前病历记录是否已经打印过
Dim rsTemp As ADODB.Recordset
On Error GoTo errHand
    '因要求保留电子病历记录（打印人，打印时间），所以历史数据不转移，记录进行联合查询
    gstrSQL = "Select 打印人, 打印时间 From 电子病历打印 Where 文件id = [1]" & vbNewLine & _
            "Union" & vbNewLine & _
            "Select 打印人, 打印时间 From 电子病历记录 Where ID = [1] And 打印人 is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
    If rsTemp.EOF Then Exit Function
    
    Do Until rsTemp.EOF
        strPrintInfo = strPrintInfo & vbCrLf & "打印人：" & Rpad(rsTemp!打印人, 8) & "打印时间：" & Format(rsTemp!打印时间, "yyyy-MM-dd hh:mm")
        rsTemp.MoveNext
    Loop
    strPrintInfo = Mid(strPrintInfo, 3)
    EprPrinted = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function EprIsCommit() As String
'以|分隔方式返回,状态为0 不允许 1 允许，分别控制 新增|删除|撤档

Dim rsTemp As ADODB.Recordset, intNew As Integer, intDel As Integer, intMod As Integer
    gstrSQL = "Select 病案状态 From 病案主页 Where 病人id = [1] And 主页id = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)

    Select Case NVL(rsTemp!病案状态, 0)
        Case 0
            intNew = 1: intDel = 1: intMod = 1
        Case 1 '等待审查
            intNew = 0: intDel = 0: intMod = 0
        Case 2 '拒绝审查
            intNew = 1: intDel = 1: intMod = 1
        Case 3 '正在审查
            intNew = 0: intDel = 0: intMod = 0
        Case 4 '审查反馈
            intNew = 0: intDel = 0: intMod = 1
        Case 5 '审查归档
            intNew = 0: intDel = 0: intMod = 0
        Case 6 '审查整改
            intNew = 0: intDel = 0: intMod = 1
        Case 13 '正在抽查
            intNew = 1: intDel = 1: intMod = 1
        Case 14 '抽查反馈
            intNew = 1: intDel = 1: intMod = 1
        Case 16 '抽查整改
            intNew = 1: intDel = 1: intMod = 1
        Case Else
            intNew = 0: intDel = 0: intMod = 0
    End Select
    EprIsCommit = CStr(intNew) & "|" & CStr(intDel) & "|" & CStr(intMod)
End Function
Private Sub PrintCancel(ByVal lngRecordId As Long)
'取消标记打印
On Error GoTo errHand
    gstrSQL = "Zl_电子病历打印_Cancel(" & lngRecordId & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        vfgWrit.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        fraColSel.Move Me.vfgWrit.Left + 50, Me.vfgWrit.Top + 50
        fraColSel.ZOrder 0
        vsColumn.Move fraColSel.Left, fraColSel.Top + fraColSel.Height
        vsColumn.ZOrder 0
        
    End Select
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim i As Long

    If Button = 1 Then '列选择器
        '根据当前状态直接确定勾选状态
        With vsColumn
            If .Visible Then
                .Visible = False
                vfgWrit.SetFocus
            Else
                For i = .FixedRows To .Rows - 1
                    If vfgWrit.ColHidden(.RowData(i)) Or vfgWrit.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next

                .Left = fraColSel.Left
                .Top = fraColSel.Top + fraColSel.Height
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub vfgWrit_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
If picInfo.Visible Then
    picInfo.Move vfgWrit.Cell(flexcpLeft, NewTopRow, mCol.w当前情况) + vfgWrit.Cell(flexcpWidth, NewTopRow, mCol.w当前情况) - picInfo.Width - 30
End If
End Sub

Private Sub vfgWrit_DblClick()
Dim lFileId As Long, bFinded As Boolean
    '当vfgWrit行变换时不允许双击事件，当双击事件执行时不允许双击事件重载
    If mblnViewTag = True Or mblnViewNow = True Then Exit Sub
    '双击事件执行标志，true正在执行，false没有执行
    mblnViewNow = True
    If Not mblnEdit Then Exit Sub
    lFileId = Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID))
    If lFileId = 0 Then Exit Sub
    If vfgWrit.TextMatrix(vfgWrit.Row, mCol.w编辑方式) = 0 Then
        Dim fViewDoc As New frmEPRView, blnCanPrint As Boolean
        blnCanPrint = (InStr(1, mstrPrivs, "护理病历打印") > 0) And (Trim(vfgWrit.TextMatrix(vfgWrit.Row, mCol.w归档人)) = "" Or InStr(1, mstrPrivs, "归档病历输出") > 0)
        fViewDoc.ShowMe Me, CLng(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)), , blnCanPrint
    Else
        If Not mObjTabEprView Is Nothing Then
            bFinded = mObjTabEprView.Showfrm(lFileId, mlngPatiId, mlngPageId, cprPF_住院, mlngDeptId)
        End If
        If Not bFinded And mblnEdit Then
            mObjTabEprView.InitOpenEPR Me, cprEM_修改, cprET_单病历编辑, lFileId, True, 0, cprPF_住院, mlngPatiId, mlngPageId, , mlngDeptId, 0, mstrPrivs, , InStr(mstrPrivs, "护理病历打印") > 0, Val(gstrESign)
        End If
    End If
    mblnViewNow = False
    Call vfgWrit_RowColChange '防止选中行数据没有刷新，手动刷新
End Sub

Private Sub vfgWrit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngCol As Long, lngRow As Long
    lngCol = vfgWrit.MouseCol: lngRow = vfgWrit.MouseRow
    If lngRow <= 0 Then picInfo.Visible = False: Exit Sub
    
    If Not Me.ActiveControl Is Nothing Then
        If Me.ActiveControl.Name <> "vfgWrit" Then
            vfgWrit.SetFocus
        Else
            vfgWrit.SetFocus
        End If
    Else
        vfgWrit.SetFocus
    End If
    
    If Val(vfgWrit.TextMatrix(lngRow, mCol.wID)) <> 0 Then
        If Val(picInfo.Tag) = lngRow And picInfo.Visible Then Exit Sub
        picInfo.Tag = lngRow
        picInfo.Move vfgWrit.Cell(flexcpLeft, lngRow, mCol.w当前情况) + vfgWrit.Cell(flexcpWidth, lngRow, mCol.w当前情况) - picInfo.Width - 30, vfgWrit.Cell(flexcpTop, lngRow, mCol.w当前情况) + 15
        If vfgWrit.RowSel = lngRow Then
            picInfo.BackColor = vfgWrit.BackColorSel
        Else
            picInfo.BackColor = &H80000005
        End If
        picInfo.Visible = True
    Else
        picInfo.Visible = False
    End If
End Sub

Private Sub vfgWrit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If vfgWrit.MouseRow = -1 Then vfgWrit.Row = vfgWrit.Rows - 1
    
    RaiseEvent MouseUp(Button, Shift, X, Y)
    'mblnInsideTools = True
    If Button = vbRightButton And mblnInsideTools Then
        Dim Popup As CommandBar
        Dim cbrControl As CommandBarControl
        
        Set Popup = cbrMain.Add("Popup", xtpBarPopup)
        With Popup.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审阅(&U)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10 + 1, "归档(&I)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "取消打印(&P)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "验证签名(&V)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Sort, "病历排序(&S)")
            Popup.ShowPopup
        End With
    End If
End Sub

Private Sub vfgWrit_RowColChange()
    If vfgWrit.Cols < mCol.wID + 1 Then Exit Sub '未初始化
    
    '当双击事件执行时，不执行变换读取内容操作，当前行与记录行不相等时，执行刷新
    If mblnViewNow = True Then Exit Sub
    mblnViewTag = True
    
    If Not mfrmNew Is Nothing Then dkpMain.Panes(3).Close

    Err = 0
    On Error Resume Next
    mlngCurId = Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.wID)) '选中行的ID
    If Val(vfgWrit.Tag) = mlngCurId Then Exit Sub '未切换行

    If Not mfrmContent Is Nothing Then Call mfrmContent.zlRefresh(mlngCurId, IIf(mblnEdit = False, "", mstrPrivs), , mblnMoved, , Val(vfgWrit.TextMatrix(vfgWrit.Row, mCol.w编辑方式)), True)
    vfgWrit.Tag = mlngCurId '刷新完后记录当前行的ID
    mblnViewTag = False
    
End Sub

Private Sub vsColumn_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
    Dim lngCol As Long, T As Variant, i As Long

    If Col = 0 Then
        lngCol = vsColumn.RowData(Row)
        If Val(vsColumn.TextMatrix(Row, 0)) <> 0 Then
            T = Split("270;0;0;1200;2000;800;1600;800;1600;800;0;3300;0;0;0;1200;0;0", ";")
            vfgWrit.ColWidth(lngCol) = T(lngCol)
            vfgWrit.ColHidden(lngCol) = False
        Else
            vfgWrit.ColWidth(lngCol) = 0
            vfgWrit.ColHidden(lngCol) = True
        End If
    End If
    Dim strCols As String
    For i = 0 To vfgWrit.Cols - 1
        strCols = strCols & IIf(i = 0, "", ";") & vfgWrit.ColWidth(i)
    Next
    mstrColWidthConfig = strCols
End Sub

Private Sub vsColumn_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    With vsColumn
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub

Private Sub vsColumn_LostFocus()
    On Error Resume Next
    vsColumn.Visible = False
End Sub

Private Sub vsColumn_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error Resume Next
    If Col <> 0 Or vsColumn.Cell(flexcpForeColor, Row, 1) = vsColumn.BackColorFixed Then Cancel = True
End Sub

Private Sub vsColumn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then '关闭列选择器
        If vsColumn.Visible Then
            vsColumn.Visible = False
            vfgWrit.SetFocus
        End If
    ElseIf Shift = vbAltMask And KeyCode = vbKeyC Then '打开列选择器
        Call imgColSel_MouseUp(1, 0, 0, 0)
    End If
End Sub

Private Sub vfgWrit_KeyDown(KeyCode As Integer, Shift As Integer)
    vsColumn_KeyDown KeyCode, Shift
End Sub
Private Function TimeLimitOut() As Boolean
'功能:检查是否有转科，出院，预出院情况，有则给出事件和补录时限
Dim rsTemp As New ADODB.Recordset, lngTimeLimit As Long, strReturn As String
    
    gstrSQL = "Select Decode(终止原因, 1, '出院', 3, '转科', 10, '预出院') 事件, 终止时间,Trunc((Sysdate - 终止时间) * 24, 5) 当前时差" & vbNewLine & _
                "From 病人变动记录" & vbNewLine & _
                "Where ID = (Select Nvl(Max(ID), 0)" & vbNewLine & _
                "            From 病人变动记录" & vbNewLine & _
                "            Where 病人id = [1] And 主页id = [2] And 终止时间 Is Not Null And 终止原因 In (1, 3, 10))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提出变动记录", mlngPatiId, mlngPageId)
    If rsTemp.EOF Then Exit Function
    
    lngTimeLimit = Val(zlDatabase.GetPara("数据补录时限", 100))
    
    If rsTemp!当前时差 > lngTimeLimit Then
        If rsTemp!事件 = "转科" Then
            strReturn = rsTemp!事件 & "|" & lngTimeLimit
            gstrSQL = "Select 出院科室id From 病案主页 Where 病人id = [1] And 主页id = [2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取出院科室", mlngPatiId, mlngPageId)
            If mlngDeptId = rsTemp!出院科室ID Then strReturn = "" '转科后，在转入科室新增病历不受时限限制
        Else
            strReturn = rsTemp!事件 & "|" & lngTimeLimit
        End If
    End If
    
    If strReturn <> "" Then
        MsgBox "该病人已经" & Split(strReturn, "|")(0) & ",并且超过设定的" & Split(strReturn, "|")(1) & "小时补录时限,不允许变动病历。", vbInformation, gstrSysName
        TimeLimitOut = True
    End If
End Function
