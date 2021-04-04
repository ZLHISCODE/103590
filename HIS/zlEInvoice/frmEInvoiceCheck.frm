VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEInvoiceCheck 
   BorderStyle     =   0  'None
   Caption         =   "电子票据核对"
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picMain 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   8748
      Left            =   600
      ScaleHeight     =   8745
      ScaleWidth      =   10335
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   576
      Width           =   10332
      Begin VB.Frame fraSplit 
         BorderStyle     =   0  'None
         Height          =   50
         Left            =   2088
         MousePointer    =   7  'Size N S
         TabIndex        =   13
         Top             =   3624
         Width           =   1005
      End
      Begin VB.PictureBox picFilter 
         BorderStyle     =   0  'None
         Height          =   444
         Left            =   72
         ScaleHeight     =   450
         ScaleWidth      =   9990
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   144
         Width           =   9996
         Begin VB.OptionButton opt退票 
            Caption         =   "退票"
            Height          =   285
            Left            =   8190
            TabIndex        =   10
            Top             =   92
            Width           =   705
         End
         Begin VB.OptionButton opt开票和退票 
            Caption         =   "开票和退票"
            Height          =   285
            Left            =   6810
            TabIndex        =   9
            Top             =   92
            Value           =   -1  'True
            Width           =   1245
         End
         Begin VB.CommandButton cmdCheck 
            Caption         =   "核对(&C)"
            Height          =   300
            Left            =   9015
            TabIndex        =   11
            Top             =   84
            Width           =   1000
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   276
            Left            =   3576
            TabIndex        =   6
            Top             =   96
            Width           =   1356
            _ExtentX        =   2381
            _ExtentY        =   476
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   115539971
            CurrentDate     =   43941
         End
         Begin VB.ComboBox cbo开票点 
            Height          =   276
            Left            =   672
            TabIndex        =   4
            Top             =   96
            Width           =   1812
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   276
            Left            =   5232
            TabIndex        =   8
            Top             =   96
            Width           =   1356
            _ExtentX        =   2381
            _ExtentY        =   476
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   115539971
            CurrentDate     =   43941
         End
         Begin VB.Label lbl业务时间_ 
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Left            =   4992
            TabIndex        =   7
            Top             =   144
            Width           =   180
         End
         Begin VB.Label lbl业务日期 
            AutoSize        =   -1  'True
            Caption         =   "收费日期"
            Height          =   180
            Left            =   2760
            TabIndex        =   5
            Top             =   144
            Width           =   720
         End
         Begin VB.Label lbl开票点 
            AutoSize        =   -1  'True
            Caption         =   "开票点"
            Height          =   180
            Left            =   72
            TabIndex        =   3
            Top             =   144
            Width           =   540
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfTotalCheck 
         Height          =   1356
         Left            =   192
         TabIndex        =   12
         Top             =   864
         Width           =   6108
         _cx             =   1983064598
         _cy             =   1983056216
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
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
      Begin VSFlex8Ctl.VSFlexGrid vsfDetailCheck 
         Height          =   1404
         Left            =   816
         TabIndex        =   14
         Top             =   4104
         Width           =   4404
         _cx             =   1983061592
         _cy             =   1983056300
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
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
   Begin VB.Shape shpBorder 
      BackColor       =   &H8000000D&
      BorderColor     =   &H8000000C&
      Height          =   1032
      Left            =   0
      Top             =   888
      Width           =   528
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   10848
      _Version        =   589884
      _ExtentX        =   19135
      _ExtentY        =   529
      _StockProps     =   6
      Caption         =   "电子票据核对"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmEInvoiceCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Object, mlngSys As Long, mlngModule As Long, mstrDBUser As String
Private mcbsMain   As Object          'CommandBar控件
Private mobjEInvoice As clsEInvoiceModule
Private mblnPrinting As Boolean
Private mrs开票点 As ADODB.Recordset
Private mrs收费员 As ADODB.Recordset
Private mbyt票据核对时间类型 As Byte '0-票据开具时间，1-费用业务发生时间
Private mstrEInvoiceNodeCode As String '开票点

Public Event ShowPopupMenu(ByVal blnAddOutPutExcel As Boolean)
Public Event ShowInfo(ByVal strInfo As String)

Public Sub InitCommVariable(frmParent As Object, cbsThis As Object, ByVal lngSys As Long, lngModule As Long, ByVal strDBUser As String, _
    objEInvoice As Object)
    '初始化变量
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    mstrDBUser = strDBUser
    mlngSys = lngSys: mlngModule = lngModule
    Set mobjEInvoice = objEInvoice
    mbyt票据核对时间类型 = mobjEInvoice.ZLCheckTimeMode
End Sub

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom

    '文件菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '放在输出到Excel之后
        Set cbrControl = .Find(, conMenu_File_Excel)
    End With

    '编辑菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If

    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", cbrMenuBar.index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "核对明细(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_FlatAccount, "平账修正(&M)"): cbrControl.BeginGroup = True
    End With

    '查看菜单
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) '刷新项前(多个时注意反序)
    End With

    '工具栏定义
    '-----------------------------------------------------
    Set cbrToolBar = mcbsMain(2)
    For Each cbrControl In cbrToolBar.Controls '先求出前面的最后一个Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "核对明细", cbrControl.index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_FlatAccount, "平账修正", cbrControl.index + 1): cbrControl.BeginGroup = True
    End With

    '命令的快键绑定
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        '.Add FCONTROL, Asc("N"), conMenu_Edit_NewItem
    End With
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Select Case Control.ID
    Case conMenu_File_Preview '预览
        Call OutputList(2)
    Case conMenu_File_Print '打印
        Call OutputList(1)
    Case conMenu_File_Excel '输出到Excel…
        Call OutputList(3)
    Case conMenu_Edit_Audit '核对
        Call DetailCheck
    Case conMenu_Edit_FlatAccount '平账修正
        Call FlatAccount
    End Select
End Sub

Private Sub FlatAccount()
    '做平帐处理
    Dim i As Long, byt场合 As Byte
    Dim lngEInvoiceID As Long, lng结算ID As Long, bln已换开纸质 As Boolean
    Dim rsEInvoice As ADODB.Recordset, strErrMsg As String
    Dim strSQL As String, blnTrans As Boolean
    Dim str业务日期 As String, strDate As String
    Dim bln修正成功 As Boolean, strMsg As String
    Dim cllPro As New Collection, lng冲销ID As Long
    
    Dim strSysSouceName_Out As String, strExtend As String
    Dim strEInvoiceCode_out As String, strEInvoiceNo_Out As String
    Dim strCheckCode_out As String, strCreateTime_Out As String, strEInvQRCode_Out As String, strEInvUrl_Out As String, strEInvUrl1_Out As String
    Dim strEinvRemark_Out As String
    
    On Error GoTo ErrHandler
    str业务日期 = vsfTotalCheck.TextMatrix(vsfTotalCheck.Row, vsfTotalCheck.ColIndex("业务日期"))
    If str业务日期 = "" Then Exit Sub
    
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    With vsfDetailCheck
        For i = .FixedRows To .Rows - 1
            '.RowData(i) = "":平台有his无的票据，这类票据在his无记录，不能修正
            If .TextMatrix(i, .ColIndex("核对结果")) = "核对失败" And InStr(1, .RowData(i), "_") > 0 Then
                Set cllPro = New Collection
                lng结算ID = Split(.RowData(i), "_")(1): lngEInvoiceID = Split(.RowData(i), "_")(2)
                
                Select Case Val(zlStr.NeedCode(.TextMatrix(i, .ColIndex("修正方式"))))
                Case 1 '1-作废HIS数据
                    Set rsEInvoice = GetEInvoiceInfo(lngEInvoiceID, strErrMsg)
                    If Val(Nvl(rsEInvoice!是否换开)) = 1 Then '已换开纸质票据，无法作废收回
                        bln修正成功 = False
                        strMsg = "已换开纸质票据，无法进行作废收回"
                    Else
                        bln修正成功 = True
                        lng冲销ID = zlDatabase.GetNextId("电子票据使用记录")
                        'Zl_电子票据使用记录_Delete
                        strSQL = "Zl_电子票据使用记录_Delete("
                        '  Id_In           In 电子票据使用记录.Id%Type,
                        strSQL = strSQL & "" & lng冲销ID & ","
                        '  开票点_In       In 电子票据使用记录.开票点%Type,
                        strSQL = strSQL & "'" & .Cell(flexcpData, i, .ColIndex("HIS数据-开票点")) & "',"
                        '  系统来源_In     In 电子票据使用记录.系统来源%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  生成时间_In     In 电子票据使用记录.生成时间%Type,
                        strSQL = strSQL & "'" & Format(strDate, "yyyyMMddHHmmss000") & "',"
                        '  备注_In         In 电子票据使用记录.备注%Type,
                        strSQL = strSQL & "'" & "平账修正，作废HIS数据" & "',"
                        '  操作员编号_In   In 电子票据使用记录.操作员编号%Type,
                        strSQL = strSQL & "'" & UserInfo.编号 & "',"
                        '  操作员姓名_In   In 电子票据使用记录.操作员姓名%Type,
                        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                        '  登记时间_In     In 电子票据使用记录.登记时间%Type,
                        strSQL = strSQL & "to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
                        '  原电子票据id_In In 电子票据使用记录.Id%Type
                        strSQL = strSQL & "" & lngEInvoiceID & ")"
                        cllPro.Add strSQL
                        
                        '作废记录加入修正记录
                        'Zl_电子票据修正记录_Update
                        strSQL = "Zl_电子票据修正记录_Update("
                        '  业务日期_In   电子票据修正记录.业务日期%Type,
                        strSQL = strSQL & "To_Date('" & str业务日期 & "','yyyy-mm-dd'),"
                        '  电子票据id_In 电子票据修正记录.电子票据id%Type,
                        strSQL = strSQL & "" & lng冲销ID & ","
                        '  业务流水号_In 电子票据修正记录.业务流水号%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  His开票点_In    电子票据修正记录.His开票点%Type,
                        strSQL = strSQL & "'" & .Cell(flexcpData, i, .ColIndex("HIS数据-开票点")) & "',"
                        '  His开票金额_In  电子票据修正记录.His开票金额%Type,
                        strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("HIS数据-开票金额"))) & ","
                        '  His票据状态_In  电子票据修正记录.His票据状态%Type,
                        strSQL = strSQL & "" & 3 & ","
                        '  平台开票点_In   电子票据修正记录.平台开票点%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  平台开票金额_In 电子票据修正记录.平台开票金额%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  平台票据状态_In 电子票据修正记录.平台票据状态%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  修正方式_In   电子票据修正记录.修正方式%Type,
                        strSQL = strSQL & "" & 4 & ","
                        '  修正人_In     电子票据修正记录.修正人%Type,
                        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                        '  修正时间_In   电子票据修正记录.修正时间%Type,
                        strSQL = strSQL & "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
                        '  修正结果_In   电子票据修正记录.修正结果%Type,
                        strSQL = strSQL & "" & 1 & ","
                        '  修正说明_In   电子票据修正记录.修正说明%Type
                        strSQL = strSQL & "'" & "平账修正，作废HIS数据时产生的作废记录" & "')"
                        cllPro.Add strSQL
                    End If
                    
                Case 2 '2-作废平台数据
                    strExtend = GetJsonNodeString("bustype", .Cell(flexcpData, i, .ColIndex("票据平台数据-票据代码")), Json_Text)
                    strExtend = strExtend & "," & GetJsonNodeString("billbatchcode", .TextMatrix(i, .ColIndex("票据平台数据-票据代码")), Json_Text)
                    strExtend = strExtend & "," & GetJsonNodeString("billno", .TextMatrix(i, .ColIndex("票据平台数据-票据号码")), Json_Text)
                    strExtend = "{" & strExtend & "}"
                    
                    bln修正成功 = mobjEInvoice.zlCancelEInvoice(Me, lngEInvoiceID, mstrEInvoiceNodeCode, strSysSouceName_Out, _
                        strEInvoiceCode_out, strEInvoiceNo_Out, strCheckCode_out, strCreateTime_Out, strEInvQRCode_Out, strEInvUrl_Out, strEInvUrl1_Out, strEinvRemark_Out, strMsg, strExtend)
                    
                    If bln修正成功 Then '冲红记录加入修正记录
                         'Zl_电子票据修正记录_Update
                        strSQL = "Zl_电子票据修正记录_Update("
                        '  业务日期_In   电子票据修正记录.业务日期%Type,
                        strSQL = strSQL & "To_Date('" & str业务日期 & "','yyyy-mm-dd'),"
                        '  电子票据id_In 电子票据修正记录.电子票据id%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  业务流水号_In 电子票据修正记录.业务流水号%Type,
                        strSQL = strSQL & "'" & .TextMatrix(i, .ColIndex("票据平台数据-业务流水号")) & "',"
                        '  His开票点_In    电子票据修正记录.His开票点%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  His开票金额_In  电子票据修正记录.His开票金额%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  His票据状态_In  电子票据修正记录.His票据状态%Type,
                        strSQL = strSQL & "" & "Null" & ","
                        '  平台开票点_In   电子票据修正记录.平台开票点%Type,
                        strSQL = strSQL & "'" & mstrEInvoiceNodeCode & "',"
                        '  平台开票金额_In 电子票据修正记录.平台开票金额%Type,
                        strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("票据平台数据-开票金额"))) & ","
                        '  平台票据状态_In 电子票据修正记录.平台票据状态%Type,
                        strSQL = strSQL & "" & 3 & ","
                        '  修正方式_In   电子票据修正记录.修正方式%Type,
                        strSQL = strSQL & "" & 4 & ","
                        '  修正人_In     电子票据修正记录.修正人%Type,
                        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                        '  修正时间_In   电子票据修正记录.修正时间%Type,
                        strSQL = strSQL & "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
                        '  修正结果_In   电子票据修正记录.修正结果%Type,
                        strSQL = strSQL & "" & 1 & ","
                        '  修正说明_In   电子票据修正记录.修正说明%Type
                        strSQL = strSQL & "'" & "平账修正，作废平台数据时产生的冲红记录" & "')"
                        cllPro.Add strSQL
                    End If
                    
                Case 3 '3-作废HIS和平台数据重开票据
                    '暂不处理，应该不会出现这种情况
                    bln修正成功 = False
                    strMsg = "暂不支持作废HIS和平台数据重开票据"
                    
                Case 4 '4-不修正仅标记
                    bln修正成功 = True
                End Select
                
                'Zl_电子票据修正记录_Update
                strSQL = "Zl_电子票据修正记录_Update("
                '  业务日期_In   电子票据修正记录.业务日期%Type,
                strSQL = strSQL & "To_Date('" & str业务日期 & "','yyyy-mm-dd'),"
                '  电子票据id_In 电子票据修正记录.电子票据id%Type,
                strSQL = strSQL & "" & ZVal(lngEInvoiceID) & ","
                '  业务流水号_In 电子票据修正记录.业务流水号%Type,
                strSQL = strSQL & "'" & .TextMatrix(i, .ColIndex("票据平台数据-业务流水号")) & "',"
                '  His开票点_In    电子票据修正记录.His开票点%Type,
                strSQL = strSQL & "'" & .Cell(flexcpData, i, .ColIndex("HIS数据-开票点")) & "',"
                '  His开票金额_In  电子票据修正记录.His开票金额%Type,
                strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("HIS数据-开票金额"))) & ","
                '  His票据状态_In  电子票据修正记录.His票据状态%Type,
                strSQL = strSQL & "" & Val(.Cell(flexcpData, i, .ColIndex("HIS数据-票据状态"))) & ","
                '  平台开票点_In   电子票据修正记录.平台开票点%Type,
                strSQL = strSQL & "'" & .Cell(flexcpData, i, .ColIndex("票据平台数据-开票点")) & "',"
                '  平台开票金额_In 电子票据修正记录.平台开票金额%Type,
                strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("票据平台数据-开票金额"))) & ","
                '  平台票据状态_In 电子票据修正记录.平台票据状态%Type,
                strSQL = strSQL & "" & Val(.Cell(flexcpData, i, .ColIndex("票据平台数据-票据状态"))) & ","
                '  修正方式_In   电子票据修正记录.修正方式%Type,
                strSQL = strSQL & "" & Val(zlStr.NeedCode(.TextMatrix(i, .ColIndex("修正方式")))) & ","
                '  修正人_In     电子票据修正记录.修正人%Type,
                strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                '  修正时间_In   电子票据修正记录.修正时间%Type,
                strSQL = strSQL & "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
                '  修正结果_In   电子票据修正记录.修正结果%Type,
                strSQL = strSQL & "" & IIf(bln修正成功, 1, 0) & ","
                '  修正说明_In   电子票据修正记录.修正说明%Type
                strSQL = strSQL & "'" & strMsg & "')"
                cllPro.Add strSQL
                
                gcnOracle.BeginTrans: blnTrans = True
                ExecuteProcedureArrAy cllPro, Me.Caption, True, True
                gcnOracle.CommitTrans: blnTrans = False
                
                .TextMatrix(i, .ColIndex("修正人")) = UserInfo.姓名
                .TextMatrix(i, .ColIndex("修正时间")) = Format(strDate, "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(i, .ColIndex("修正结果")) = IIf(bln修正成功, "修正成功", "修正失败")
                .TextMatrix(i, .ColIndex("修正说明")) = strMsg
                
                If bln修正成功 Then
                    .TextMatrix(i, .ColIndex("核对结果")) = "核对成功"
                    .TextMatrix(i, .ColIndex("核对说明")) = ""
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
                Else
                    .TextMatrix(i, .ColIndex("核对结果")) = "核对失败"
                    .TextMatrix(i, .ColIndex("核对说明")) = strMsg
                End If
            End If
        Next
    End With
    
    Call DetailCheck
    Exit Sub
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub OutputList(bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte
    Dim intCurrentRow As Integer, vsfGrid As VSFlexGrid
    
    On Error GoTo ErrHandler
    '表头
    Set objOut = New zlPrint1Grd
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    If Me.ActiveControl Is vsfDetailCheck Then
        Set vsfGrid = vsfDetailCheck
        objOut.Title.Text = "电子票据明细核对清单"
    Else
        Set vsfGrid = vsfTotalCheck
        objOut.Title.Text = "电子票据汇总核对清单"
    End If
    
    '表项
    If Me.ActiveControl Is vsfDetailCheck Then
        Set objRow = New zlTabAppRow
        objRow.Add "开票点：" & cbo开票点.Text
        objRow.Add "业务日期：" & vsfTotalCheck.TextMatrix(vsfTotalCheck.Row, vsfTotalCheck.ColIndex("业务日期"))
        objOut.UnderAppRows.Add objRow
    Else
        Set objRow = New zlTabAppRow
        objRow.Add "开票点：" & cbo开票点.Text
        objRow.Add "业务时间：" & Format(dtp开始时间, "yyyy-mm-dd") & " 至 " & Format(dtp结束时间, "yyyy-mm-dd")
        objOut.UnderAppRows.Add objRow
    End If
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    vsfGrid.Redraw = False
    intCurrentRow = vsfGrid.Row
    mblnPrinting = True
    
    '表体
    Set objOut.Body = vsfGrid
    '输出
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mblnPrinting = False
    vsfGrid.Row = intCurrentRow
    vsfGrid.Redraw = True
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    mblnPrinting = False
    vsfGrid.Row = intCurrentRow
    vsfGrid.Redraw = True
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel '预览,打印,输出到Excel…
        If Me.ActiveControl Is vsfDetailCheck Then
            Control.Enabled = vsfDetailCheck.TextMatrix(2, 0) <> ""
        Else
            Control.Enabled = vsfTotalCheck.TextMatrix(2, 0) <> ""
        End If
    
    Case conMenu_Edit_Audit '核对明细
        Control.Enabled = vsfTotalCheck.TextMatrix(2, 0) <> ""
    Case conMenu_Edit_FlatAccount '平账修正
        If vsfTotalCheck.Row > 0 Then
            Control.Enabled = vsfDetailCheck.TextMatrix(2, 0) <> "" And vsfTotalCheck.TextMatrix(vsfTotalCheck.Row, vsfTotalCheck.ColIndex("核对结果")) = "核对失败"
        Else
            Control.Enabled = False
        End If
    
    Case conMenu_View_Refresh '刷新
        Control.Visible = False
        Control.Enabled = Control.Visible
    End Select
End Sub

Private Sub cbo开票点_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
     
    If cbo开票点.ListIndex <> -1 Then
        '弹出列表时,又在文本框输入了内容
        If UCase(cbo开票点.Text) <> UCase(cbo开票点.List(cbo开票点.ListIndex)) Then Call zlControl.CboSetIndex(cbo开票点.hWnd, -1)
    End If
    
    If cbo开票点.Text = "" Then
        cbo开票点.ListIndex = -1
    ElseIf cbo开票点.ListIndex = -1 Then
        If mrs收费员 Is Nothing Then
            If Select开票点(Me, mlngSys, mlngModule, cbo开票点, mrs开票点) = False Then
                KeyAscii = 0: zlControl.TxtSelAll cbo开票点: Exit Sub
            End If
        Else
            If Select收费员(Me, mlngSys, mlngModule, cbo开票点, mrs收费员) = False Then
                KeyAscii = 0: zlControl.TxtSelAll cbo开票点: Exit Sub
            End If
        End If
    End If
    
    If cbo开票点.ListIndex = -1 Then cbo开票点.Text = ""
End Sub

Private Sub cbo开票点_LostFocus()
    If cbo开票点.Text <> "" And cbo开票点.ListIndex < 0 Then cbo开票点.Text = ""
End Sub

Private Sub cmdCheck_Click()
    Call TotalCheck
End Sub

Private Sub TotalCheck()
    '汇总核对
    Dim dtBegin As Date, dtEnd As Date, strErrMsg As String
    Dim str开票点 As String, bytMode As Byte '1-核对开票和退票，2-仅核对退票
    
    '1.数据检查
    On Error GoTo ErrHandler
    dtBegin = Format(dtp开始时间.Value, "yyyy-MM-dd"): dtEnd = Format(dtp结束时间.Value, "yyyy-MM-dd 23:59:59")
    If dtp开始时间 > dtp结束时间 Then
        MsgBox "开始时间不能大于结束时间！", vbInformation, gstrSysName
        zlControl.ControlSetFocus dtp结束时间:  Exit Sub
    End If
    
    If DateDiff("m", dtp开始时间, dtp结束时间) > 6 Then
        If MsgBox("对当前时间范围内的数据进行核对可能需要较长时间，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
     
    bytMode = IIf(opt开票和退票.Value, 1, 2)
    str开票点 = zlStr.NeedCode(cbo开票点.Text)
    
    '2.获取数据
    Dim strSQL As String, rsHISEInvoice As ADODB.Recordset, strWhere As String, strSqlSub As String
    If mbyt票据核对时间类型 = 1 Then
        If bytMode = 2 Then strWhere = " And a.记录状态 = 2"
        If str开票点 <> "" Then strWhere = strWhere & " And a.开票点 = [3]"
        
        '1)预交款
        strSQL = _
            " Select Distinct a.ID" & _
            " From 电子票据使用记录 A,病人预交记录 B" & _
            " Where a.结算ID =b.ID And a.票种=2 And b.记录性质=1 And b.收款时间 Between [1] And [2]" & strWhere
        '余额退款
        strSQL = strSQL & " Union All " & _
            " Select Distinct a.ID" & _
            " From 电子票据使用记录 A,病人预交记录 B" & _
            " Where a.退款ID =b.ID And a.票种=2 And b.记录性质=11 And b.收款时间 Between [1] And [2]" & strWhere
        '2)就诊卡
        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            " Select Distinct a.ID " & _
            " From 电子票据使用记录 A,住院费用记录 B" & _
            " Where a.结算ID =b.结帐ID And a.票种=5 And b.记录性质=5 And b.记录状态 In(1,3) And b.登记时间 Between [1] And [2]" & strWhere
        '3)结帐
            strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
                " Select Distinct a.ID" & _
                " From 电子票据使用记录 A,病人结帐记录 B" & _
                " Where a.结算ID =b.ID And a.票种=3 And b.记录状态 In(1,3) And b.收费时间 Between [1] And [2]" & strWhere
        '4)挂号、收费
        strSqlSub = _
            " Select Distinct a.ID" & _
            " From 电子票据使用记录 A,门诊费用记录 B" & _
            " Where a.结算ID =b.结帐ID And a.票种=[票种] And b.记录性质=[记录性质] And b.记录状态 In(1,3) And b.登记时间 Between [1] And [2]" & strWhere
        
        '保险补充结算
        strSqlSub = strSqlSub & " Union All " & _
            " Select Distinct a.ID" & _
            " From 电子票据使用记录 A,费用补充记录 B" & _
            " Where a.结算ID =b.结算ID And a.票种=[票种]  And b.记录性质=[记录性质] And Nvl(b.附加标志,0)=[附加标志]" & _
            "           And b.记录状态 In(1,3) And b.登记时间 Between [1] And [2]" & strWhere
            
        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            Replace(Replace(Replace(strSqlSub, "[记录性质]", 1), "[附加标志]", 0), "[票种]", 1)

        strSQL = IIf(strSQL = "", "", strSQL & " Union All ") & _
            Replace(Replace(Replace(strSqlSub, "[记录性质]", 4), "[附加标志]", 1), "[票种]", 4)
        
        strSQL = _
            " Select To_Char(To_Date(Substr(a.生成时间, 1, 8), 'yyyymmdd'),'yyyy-mm-dd') As 业务日期, Count(1) As 开票数, Sum(Decode(a.记录状态, 2, -1, 1) * a.票据金额) As 开票金额" & _
            " From 电子票据使用记录 A,(" & strSQL & ") B" & _
            " Where a.ID =b.ID" & _
            " Group By To_Date(Substr(a.生成时间, 1, 8), 'yyyymmdd')"
        
    Else
        strSQL = _
            " Select To_Char(To_Date(Substr(a.生成时间, 1, 8), 'yyyymmdd'),'yyyy-mm-dd') As 业务日期, Count(1) As 开票数, Sum(Decode(a.记录状态, 2, -1, 1) * a.票据金额) As 开票金额" & _
            " From 电子票据使用记录 A" & _
            " Where To_Date(Substr(a.生成时间, 1, 8), 'yyyymmdd') Between [1] And [2]" & _
                        IIf(bytMode = 2, " And a.记录状态 = 2", " And a.记录状态 In(1,2,3)") & _
                        IIf(str开票点 = "", "", " And a.开票点=[3]") & _
            " Group By To_Date(Substr(a.生成时间, 1, 8), 'yyyymmdd')"
    End If
    Set rsHISEInvoice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtBegin, dtEnd, str开票点)
            
    Dim clldata As Collection, cllDatas As Collection '集合(业务日期,总笔数,开票数,开票金额,返回结果,错误原因),Key=_业务日期
    If mobjEInvoice.ZlGetTotalCheckData(dtBegin, dtEnd, cllDatas, bytMode, str开票点, strErrMsg) Then
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Sub
    End If
    
    Dim rsPreCheck As ADODB.Recordset '上次核对记录
    strSQL = _
        " Select To_Char(业务日期,'yyyy-mm-dd') As 业务日期, 核对人, 核对时间, 核对结果, 核对说明" & _
        " From (Select a.业务日期, a.核对人, a.核对时间, a.核对结果, a.核对说明, Row_Number() Over(Partition By a.业务日期 Order By a.核对时间 Desc) As 组号" & _
        "           From 电子票据核对记录 A" & _
        "           Where 业务日期 Between [1] And [2] And 核对类型=[3]" & _
                                IIf(str开票点 = "", "", " And a.开票点=[4]") & _
        "           )" & _
        " Where 组号 = 1"
    Set rsPreCheck = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtBegin, dtEnd, bytMode, str开票点)
    
    Dim rsUpdate As ADODB.Recordset '修正记录
    strSQL = _
        " Select 1 As 类型, To_Char(业务日期,'yyyy-mm-dd') As 业务日期, Count(1) As 开票数, Sum(HIS开票金额) As 开票金额" & _
        " From 电子票据修正记录" & _
        " Where 业务日期 Between [1] And [2] And 修正结果 = 1 And 电子票据id Is Not Null" & _
                    IIf(bytMode = 2, " And HIS票据状态 = 2", " And HIS票据状态 In(1,2,3)") & _
                    IIf(str开票点 = "", "", " And HIS开票点=[3]") & _
        " Group By 业务日期" & _
        " Union All" & _
        " Select 2 As 类型, To_Char(业务日期,'yyyy-mm-dd') As 业务日期, Count(1) As 开票数, Sum(平台开票金额) As 开票金额" & _
        " From 电子票据修正记录" & _
        " Where 业务日期 Between [1] And [2] And 修正结果 = 1 And 业务流水号 Is Not Null" & _
                    IIf(bytMode = 2, " And 平台票据状态 = 2", " And 平台票据状态 In(1,2,3)") & _
                    IIf(str开票点 = "", "", " And 平台开票点=[3]") & _
        " Group By 业务日期"
    Set rsUpdate = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtBegin, dtEnd, str开票点)

    '3.开始核对
    Dim dtCurrent As Date, blnChecked As Boolean, strCheckMsg As String
    Dim lngHIS开票数 As Long, lng平台开票数 As Long, dblHIS开票额 As Double, dbl平台开票额 As Double
    Dim lngOldRow As Long, lngOldCol As Long
    
    lngOldRow = vsfTotalCheck.Row: lngOldCol = vsfTotalCheck.Col
    vsfTotalCheck.Clear 1
    vsfTotalCheck.Rows = vsfTotalCheck.FixedRows + 1
    
    vsfDetailCheck.Clear 1
    vsfDetailCheck.Rows = vsfDetailCheck.FixedRows + 1
    
    '业务日期,HIS数据-开票数,HIS数据-开票金额,票据平台数据-开票数,票据平台数据-开票金额,票据平台数据-总笔数,
    '核对结果,核对说明,上次核对人,上次核对时间,上次核对结果,上次核对说明
    With vsfTotalCheck
        .Redraw = flexRDNone
        
        dtCurrent = dtBegin
        Do While dtCurrent <= dtEnd
            
            If .TextMatrix(.Rows - 1, .ColIndex("业务日期")) <> "" Then .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("业务日期")) = Format(dtCurrent, "yyyy-MM-dd")
            
            lngHIS开票数 = 0: dblHIS开票额 = 0
            rsHISEInvoice.Filter = "业务日期='" & Format(dtCurrent, "yyyy-MM-dd") & "'"
            If Not rsHISEInvoice.EOF Then
                lngHIS开票数 = Val(Nvl(rsHISEInvoice!开票数)): dblHIS开票额 = Val(Nvl(rsHISEInvoice!开票金额))
                rsUpdate.Filter = "类型=1 And 业务日期='" & Format(dtCurrent, "yyyy-MM-dd") & "'"
                If Not rsUpdate.EOF Then '排除已修正处理记录
                    lngHIS开票数 = lngHIS开票数 - Val(Nvl(rsUpdate!开票数))
                    dblHIS开票额 = dblHIS开票额 - Val(Nvl(rsUpdate!开票金额))
                End If
                
                .TextMatrix(.Rows - 1, .ColIndex("HIS数据-开票数")) = lngHIS开票数
                .TextMatrix(.Rows - 1, .ColIndex("HIS数据-开票金额")) = FormatEx(dblHIS开票额, 6, , , 2)
            End If
            
            lng平台开票数 = 0: dbl平台开票额 = 0
            Set clldata = Nothing
            If CollectionExitsValue(cllDatas, "_" & Format(dtCurrent, "yyyy-MM-dd")) Then
                Set clldata = cllDatas("_" & Format(dtCurrent, "yyyy-MM-dd"))
            End If
            
            blnChecked = False
            If Not clldata Is Nothing Then
                lng平台开票数 = Val(Nvl(clldata("开票数"))): dbl平台开票额 = Val(Nvl(clldata("开票数")))
                rsUpdate.Filter = "类型=2 And 业务日期='" & Format(dtCurrent, "yyyy-MM-dd") & "'"
                If Not rsUpdate.EOF Then '排除已修正处理记录
                    lngHIS开票数 = lngHIS开票数 - Val(Nvl(rsUpdate!开票数))
                    dblHIS开票额 = dblHIS开票额 - Val(Nvl(rsUpdate!开票金额))
                End If
                
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-开票数")) = lng平台开票数
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-开票金额")) = FormatEx(dbl平台开票额, 6, , , 2)
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-总笔数")) = Nvl(clldata("总笔数"))
                If Nvl(clldata("返回结果")) = "失败" Then
                    blnChecked = True
                    .TextMatrix(.Rows - 1, .ColIndex("核对结果")) = "核对失败"
                    .TextMatrix(.Rows - 1, .ColIndex("核对说明")) = Nvl(clldata("错误原因"))
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                End If
            End If
            
            If Not blnChecked Then
                blnChecked = True: strCheckMsg = ""
                '核对规则：HIS开票数 = 平台开票数 And HIS开票金额 = 平台开票金额
                If lngHIS开票数 <> lng平台开票数 Then
                    strCheckMsg = strCheckMsg & "  HIS开票数：" & lngHIS开票数 & "张/平台开票数：" & lng平台开票数 & "张"
                End If
                If dblHIS开票额 <> dbl平台开票额 Then
                    strCheckMsg = strCheckMsg & "  HIS开票金额：" & FormatEx(dblHIS开票额, 6, , , 2) & "/平台开票金额：" & FormatEx(dbl平台开票额, 6, , , 2)
                End If
                If strCheckMsg <> "" Then strCheckMsg = Mid(strCheckMsg, 2)
                blnChecked = strCheckMsg = ""
                
                If blnChecked Then
                    .TextMatrix(.Rows - 1, .ColIndex("核对结果")) = "核对成功"
                    .TextMatrix(.Rows - 1, .ColIndex("核对说明")) = ""
                Else
                    .TextMatrix(.Rows - 1, .ColIndex("核对结果")) = "核对失败"
                    .TextMatrix(.Rows - 1, .ColIndex("核对说明")) = strCheckMsg
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                End If
            End If
            
            rsPreCheck.Filter = "业务日期='" & Format(dtCurrent, "yyyy-MM-dd") & "'"
            If Not rsPreCheck.EOF Then
                .TextMatrix(.Rows - 1, .ColIndex("上次核对人")) = Nvl(rsPreCheck!核对人)
                .TextMatrix(.Rows - 1, .ColIndex("上次核对时间")) = Format(Nvl(rsPreCheck!核对时间), "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(.Rows - 1, .ColIndex("上次核对结果")) = IIf(Val(Nvl(rsPreCheck!核对结果)) = 1, "核对成功", "核对失败")
                .TextMatrix(.Rows - 1, .ColIndex("上次核对说明")) = Nvl(rsPreCheck!核对说明)
            End If
        
            dtCurrent = DateAdd("d", 1, dtCurrent)
        Loop
        
        If .Rows > .FixedRows And .Cols > .FixedCols Then     '缺省定位行
            .Row = -1 '保证在选择行不变的情况下也触发RowColChange事件
            .Row = IIf(lngOldRow < .FixedRows Or lngOldRow > .Rows - 1, IIf(.Rows - 1 > .FixedRows, .FixedRows + 1, .FixedRows), lngOldRow)
            .Col = IIf(lngOldCol = 0 Or lngOldCol > .Cols - 1, .FixedCols, lngOldCol)
            .ShowCell .Row, .Col  '立刻显示到指定单元
        End If
        .Redraw = flexRDBuffered
    End With
    
    If SaveTotalCheckData(bytMode, str开票点) = False Then
        vsfTotalCheck.Clear 1
        vsfTotalCheck.Rows = vsfTotalCheck.FixedRows + 1
    End If
    Call ShowTotalRow
    Exit Sub
ErrHandler:
    vsfTotalCheck.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowTotalRow()
    '显示汇总行
    Dim i As Long
    
    On Error GoTo ErrHandler
    With vsfTotalCheck
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, .ColIndex("业务日期")) = "合计"
        For i = .FixedRows To .Rows - 1
            .TextMatrix(.Rows - 1, .ColIndex("HIS数据-开票数")) = Val(.TextMatrix(.Rows - 1, .ColIndex("HIS数据-开票数"))) + Val(.TextMatrix(i, .ColIndex("HIS数据-开票数")))
            .TextMatrix(.Rows - 1, .ColIndex("HIS数据-开票金额")) = FormatEx(Val(.TextMatrix(.Rows - 1, .ColIndex("HIS数据-开票金额"))) + Val(.TextMatrix(i, .ColIndex("HIS数据-开票金额"))), 6, , , 2)
            .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-开票数")) = Val(.TextMatrix(.Rows - 1, .ColIndex("票据平台数据-开票数"))) + Val(.TextMatrix(i, .ColIndex("票据平台数据-开票数")))
            .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-开票金额")) = FormatEx(Val(.TextMatrix(.Rows - 1, .ColIndex("票据平台数据-开票金额"))) + Val(.TextMatrix(i, .ColIndex("票据平台数据-开票金额"))), 6, , , 2)
            .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-总笔数")) = Val(.TextMatrix(.Rows - 1, .ColIndex("票据平台数据-总笔数"))) + Val(.TextMatrix(i, .ColIndex("票据平台数据-总笔数")))
        Next
        .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SaveTotalCheckData(ByVal byt核对类型 As Byte, ByVal str开票点 As String) As Boolean
    '保存汇总核对数据
    '入参：
    '   byt核对类型 1-核对开票和退票，2-仅核对退票
    Dim strSQL As String, cllPro As New Collection
    Dim blnTran As Boolean, i As Long, strDate As String
    
    On Error GoTo ErrHandler
    
    '业务日期,HIS数据-开票数,HIS数据-开票金额,票据平台数据-开票数,票据平台数据-开票金额,票据平台数据-总笔数,
    '核对结果,核对说明,上次核对人,上次核对时间,上次核对结果,上次核对说明
    With vsfTotalCheck
        If .TextMatrix(.FixedRows, .ColIndex("业务日期")) = "" Then SaveTotalCheckData = True: Exit Function
        
        strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        For i = .FixedRows To .Rows - 1
            'Zl_电子票据核对记录_Update
            strSQL = "Zl_电子票据核对记录_Update("
            '  核对类型_In     电子票据核对记录.核对类型%Type,
            strSQL = strSQL & "" & byt核对类型 & ","
            '  业务日期_In     电子票据核对记录.业务日期%Type,
            strSQL = strSQL & "To_Date('" & Format(.TextMatrix(i, .ColIndex("业务日期")), "yyyy-MM-dd") & "','yyyy-mm-dd'),"
            '  开票点_In       电子票据核对记录.开票点%Type,
            strSQL = strSQL & "'" & str开票点 & "',"
            '  His开票数_In    电子票据核对记录.His开票数%Type,
            strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("HIS数据-开票数"))) & ","
            '  His开票金额_In  电子票据核对记录.His开票金额%Type,
            strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("HIS数据-开票金额"))) & ","
            '  平台开票数_In   电子票据核对记录.平台开票数%Type,
            strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("票据平台数据-开票数"))) & ","
            '  平台开票金额_In 电子票据核对记录.平台开票金额%Type,
            strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("票据平台数据-开票金额"))) & ","
            '  核对人_In       电子票据核对记录.核对人%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '  核对时间_In     电子票据核对记录.核对时间%Type,
            strSQL = strSQL & "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
            '  核对结果_In     电子票据核对记录.核对结果%Type,
            strSQL = strSQL & "" & IIf(.TextMatrix(i, .ColIndex("核对结果")) = "核对成功", 1, 0) & ","
            '  核对说明_In     电子票据核对记录.核对说明%Type
            strSQL = strSQL & "'" & .TextMatrix(i, .ColIndex("核对说明")) & "')"
            cllPro.Add strSQL
        Next
    End With
    
    gcnOracle.BeginTrans: blnTran = True
    ExecuteProcedureArrAy cllPro, Me.Caption, True, True
    gcnOracle.CommitTrans: blnTran = False
    
    SaveTotalCheckData = True
    Exit Function
ErrHandler:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub DetailCheck()
    '明细核对
    Dim dtBegin As Date, dtEnd As Date, strErrMsg As String
    Dim str开票点 As String, bytMode As Byte '1-核对开票和退票，2-仅核对退票
    Dim str业务日期 As String, strSQL As String
    
    On Error GoTo ErrHandler
    '1.数据检查
    If vsfDetailCheck.Row < vsfDetailCheck.FixedRows Or vsfDetailCheck.Row > vsfDetailCheck.Rows - 1 Then Exit Sub
    
    str业务日期 = vsfTotalCheck.TextMatrix(vsfTotalCheck.Row, vsfTotalCheck.ColIndex("业务日期"))
    If str业务日期 = "" Then Exit Sub
    
    dtBegin = str业务日期: dtEnd = Format(str业务日期, "yyyy-MM-dd 23:59:59")
     
    bytMode = IIf(opt开票和退票.Value, 1, 2)
    str开票点 = zlStr.NeedCode(cbo开票点.Text)
    
    '2.获取数据
    Dim rsHISEInvoice As ADODB.Recordset
    If GetEInvoiceData(0, dtBegin, dtEnd, rsHISEInvoice, IIf(bytMode = 2, 2, 0), mbyt票据核对时间类型, 0, "", str开票点) = False Then Exit Sub
    
    Dim clldata As Collection, cllDatas As Collection '集合(业务日期,业务流水号,开票点,票据种类名称,票据代码,票据号码,开票金额,开票时间,数据类型,关联票据代码,关联票据号码),Key=_业务流水号
    If Not mobjEInvoice.ZlGetDetailCheckData(dtBegin, dtEnd, cllDatas, bytMode, str开票点, strErrMsg) Then
        MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Sub
    End If

    Dim rsUpdate As ADODB.Recordset '修正记录
    strSQL = _
        " Select 电子票据id, 业务流水号, 修正方式, 修正人, 修正时间, 修正结果, 修正说明, 平台票据状态 as 票据状态" & _
        " From 电子票据修正记录" & _
        " Where 业务日期 = [1]" & IIf(str开票点 = "", "", " And HIS开票点=[2]")
    Set rsUpdate = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtBegin, str开票点)
    
    '3.开始核对
    Dim str业务流水号 As String, blnChecked As Boolean, strCheckMsg As String
    Dim strHIS票据号码 As String, str平台票据号码 As String, dblHIS开票额 As Double, dbl平台开票额 As Double
    vsfDetailCheck.Clear 1
    vsfDetailCheck.Rows = 2
    vsfDetailCheck.Rows = vsfDetailCheck.FixedRows + 1
    
    With vsfDetailCheck
        .Redraw = flexRDNone
        
        Do While Not rsHISEInvoice.EOF
            'HIS数据-收费时间,HIS数据-票据类型,HIS数据-票据状态,HIS数据-单据,HIS数据-票据代码,HIS数据-票据号码,HIS数据-开票金额,HIS数据-开票时间,
            '票据平台数据-业务流水号,票据平台数据-开票点,票据平台数据-票据代码,票据平台数据-票据号码,票据平台数据-开票金额,
            '票据平台数据-开票时间,票据平台数据-数据类型,票据平台数据-票据状态,票据平台数据-关联票据代码,票据平台数据-关联票据号码,
            '核对结果,核对说明,修正方式,修正人,修正时间,修正结果,修正说明
            
            str业务流水号 = Nvl(rsHISEInvoice!结算ID) & "_" & IIf(Val(Nvl(rsHISEInvoice!原票据ID)) = 0, Nvl(rsHISEInvoice!ID), Nvl(rsHISEInvoice!原票据ID))  '电子票据使用记录.结算ID_电子票据使用记录.ID
            If .TextMatrix(.Rows - 1, .ColIndex("HIS数据-收费时间")) <> "" Then .Rows = .Rows + 1
            .RowData(.Rows - 1) = str业务流水号
            .TextMatrix(.Rows - 1, .ColIndex("HIS数据-收费时间")) = Format(Nvl(rsHISEInvoice!收费时间), "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(.Rows - 1, .ColIndex("HIS数据-开票点")) = Nvl(rsHISEInvoice!开票点)
            .Cell(flexcpData, .Rows - 1, .ColIndex("HIS数据-开票点")) = Nvl(rsHISEInvoice!开票点)
            .TextMatrix(.Rows - 1, .ColIndex("HIS数据-票据类型")) = Decode(Val(Nvl(rsHISEInvoice!票种)), 1, "收费", 2, "预交", 3, "结帐", 4, "挂号", 5, "就诊卡")
            .Cell(flexcpData, .Rows - 1, .ColIndex("HIS数据-票据类型")) = Val(Nvl(rsHISEInvoice!票种))
            .TextMatrix(.Rows - 1, .ColIndex("HIS数据-票据状态")) = Decode(Val(Nvl(rsHISEInvoice!票据状态)), 1, "正常", 2, "冲红", 3, "作废")
            .Cell(flexcpData, .Rows - 1, .ColIndex("HIS数据-票据状态")) = Val(Nvl(rsHISEInvoice!票据状态))
            .TextMatrix(.Rows - 1, .ColIndex("HIS数据-单据号")) = Nvl(rsHISEInvoice!No)
            .TextMatrix(.Rows - 1, .ColIndex("HIS数据-票据代码")) = Nvl(rsHISEInvoice!票据代码)
            .TextMatrix(.Rows - 1, .ColIndex("HIS数据-票据号码")) = Nvl(rsHISEInvoice!票据号码)
            .TextMatrix(.Rows - 1, .ColIndex("HIS数据-开票金额")) = FormatEx(Val(Nvl(rsHISEInvoice!票据金额)), 6, , , 2)
            .TextMatrix(.Rows - 1, .ColIndex("HIS数据-开票时间")) = Format(Nvl(rsHISEInvoice!开票时间), "yyyy-MM-dd HH:mm:ss")
            strHIS票据号码 = Nvl(rsHISEInvoice!票据号码): dblHIS开票额 = Val(Nvl(rsHISEInvoice!票据金额))
            
            str业务流水号 = IIf(Val(Nvl(rsHISEInvoice!票据状态)) = 2, 2, 1) & "_" & str业务流水号
            If CollectionExitsValue(cllDatas, str业务流水号) Then
                Set clldata = cllDatas(str业务流水号)
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-业务流水号")) = Nvl(clldata("业务流水号"))
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-开票点")) = Nvl(clldata("开票点"))
                .Cell(flexcpData, .Rows - 1, .ColIndex("票据平台数据-开票点")) = Nvl(clldata("开票点"))
                '.TextMatrix(.Rows - 1, .ColIndex("票据平台数据-票据种类名称")) = Nvl(clldata("票据种类名称"))
                .Cell(flexcpData, .Rows - 1, .ColIndex("票据平台数据-票据代码")) = Nvl(clldata("业务标识"))
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-票据代码")) = Nvl(clldata("票据代码"))
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-票据号码")) = Nvl(clldata("票据号码"))
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-开票金额")) = FormatEx(Val(Nvl(clldata("开票金额"))), 6, , , 2)
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-开票时间")) = Format(CDateEx(Nvl(clldata("开票时间"))), "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-数据类型")) = Decode(Val(Nvl(clldata("数据类型"))), 1, "正常电子", 2, "电子红票", 3, "换开纸质", 4, "换开纸质红票", 5, "空白纸质")
                .Cell(flexcpData, .Rows - 1, .ColIndex("票据平台数据-数据类型")) = Val(Nvl(clldata("数据类型")))
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-票据状态")) = Decode(Val(Nvl(clldata("票据状态"))), 1, "正常", 2, "作废", 3, "冲红")
                .Cell(flexcpData, .Rows - 1, .ColIndex("票据平台数据-票据状态")) = Val(Nvl(clldata("票据状态")))
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-关联票据代码")) = Nvl(clldata("关联票据代码"))
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-关联票据号码")) = Nvl(clldata("关联票据号码"))
                str平台票据号码 = Nvl(clldata("票据号码")): dbl平台开票额 = Val(Nvl(clldata("开票金额")))
                
                blnChecked = False
                rsUpdate.Filter = "电子票据id=" & Val(Nvl(rsHISEInvoice!ID)) & " And 业务流水号='" & str业务流水号 & "'"
                If Not rsUpdate.EOF Then
                    blnChecked = Val(Nvl(rsUpdate!修正结果)) = 1
                    .TextMatrix(.Rows - 1, .ColIndex("修正方式")) = Decode(Val(Nvl(rsUpdate!修正方式)), 1, "1-作废HIS数据", 2, "2-作废平台数据", 3, "3-作废数据重开票据", 4, "4-不修正仅标记")
                    .TextMatrix(.Rows - 1, .ColIndex("修正人")) = Nvl(rsUpdate!修正人)
                    .TextMatrix(.Rows - 1, .ColIndex("修正时间")) = Format(Nvl(rsUpdate!修正时间), "yyyy-MM-dd HH:mm:ss")
                    .TextMatrix(.Rows - 1, .ColIndex("修正结果")) = IIf(Val(Nvl(rsUpdate!修正结果)) = 1, "修正成功", "修正失败")
                    .TextMatrix(.Rows - 1, .ColIndex("修正说明")) = Nvl(rsUpdate!修正说明)
                End If
                
                strCheckMsg = ""
                If Not blnChecked Then
                    '核对规则：HIS票据号码 = 平台票据号码 And HIS开票金额 = 平台开票金额
                    If strHIS票据号码 <> str平台票据号码 Then
                        strCheckMsg = strCheckMsg & "  HIS票据号码：" & strHIS票据号码 & "/平台开票数：" & str平台票据号码
                    End If
                    If dblHIS开票额 <> dbl平台开票额 Then
                        strCheckMsg = strCheckMsg & "  HIS开票金额：" & FormatEx(dblHIS开票额, 6, , , 2) & "/平台开票金额：" & FormatEx(dbl平台开票额, 6, , , 2)
                    End If
                    If strCheckMsg <> "" Then strCheckMsg = Mid(strCheckMsg, 2)
                    blnChecked = strCheckMsg = ""
                End If
                
                If blnChecked Then
                    .TextMatrix(.Rows - 1, .ColIndex("核对结果")) = "核对成功"
                    .TextMatrix(.Rows - 1, .ColIndex("核对说明")) = ""
                Else
                    .TextMatrix(.Rows - 1, .ColIndex("核对结果")) = "核对失败"
                    .TextMatrix(.Rows - 1, .ColIndex("核对说明")) = strCheckMsg
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                    .TextMatrix(.Rows - 1, .ColIndex("修正方式")) = IIf(Val(Nvl(rsHISEInvoice!票据状态)) = 1, "3-作废数据重开票据", "4-不修正仅标记")
                End If
                
                cllDatas.Remove str业务流水号
            Else
                blnChecked = False
                rsUpdate.Filter = "电子票据id=" & Val(Nvl(rsHISEInvoice!ID))
                If Not rsUpdate.EOF Then
                    blnChecked = Val(Nvl(rsUpdate!修正结果)) = 1
                    .TextMatrix(.Rows - 1, .ColIndex("修正方式")) = Decode(Val(Nvl(rsUpdate!修正方式)), 1, "1-作废HIS数据", 2, "2-作废平台数据", 3, "3-作废数据重开票据", 4, "4-不修正仅标记")
                    .TextMatrix(.Rows - 1, .ColIndex("修正人")) = Nvl(rsUpdate!修正人)
                    .TextMatrix(.Rows - 1, .ColIndex("修正时间")) = Format(Nvl(rsUpdate!修正时间), "yyyy-MM-dd HH:mm:ss")
                    .TextMatrix(.Rows - 1, .ColIndex("修正结果")) = IIf(Val(Nvl(rsUpdate!修正结果)) = 1, "修正成功", "修正失败")
                    .TextMatrix(.Rows - 1, .ColIndex("修正说明")) = Nvl(rsUpdate!修正说明)
                End If
                
                If blnChecked Then
                    .TextMatrix(.Rows - 1, .ColIndex("核对结果")) = "核对成功"
                    .TextMatrix(.Rows - 1, .ColIndex("核对说明")) = ""
                Else
                    .TextMatrix(.Rows - 1, .ColIndex("核对结果")) = "核对失败"
                    .TextMatrix(.Rows - 1, .ColIndex("核对说明")) = "HIS票据号码：" & strHIS票据号码 & "/平台票据号码：无"
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                    .TextMatrix(.Rows - 1, .ColIndex("修正方式")) = IIf(Val(Nvl(rsHISEInvoice!票据状态)) = 1, "1-作废HIS数据", "4-不修正仅标记")
                End If
            End If
             
            rsHISEInvoice.MoveNext
        Loop
        
        If Not cllDatas Is Nothing Then
            For Each clldata In cllDatas
                If .TextMatrix(.Rows - 1, .ColIndex("HIS数据-收费时间")) <> "" Or .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-业务流水号")) <> "" Then .Rows = .Rows + 1
                .RowData(.Rows - 1) = Nvl(clldata("业务流水号"))
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-业务流水号")) = Nvl(clldata("业务流水号"))
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-开票点")) = Nvl(clldata("开票点"))
                .Cell(flexcpData, .Rows - 1, .ColIndex("票据平台数据-开票点")) = Nvl(clldata("开票点"))
                '.TextMatrix(.Rows - 1, .ColIndex("票据平台数据-票据种类名称")) = Nvl(clldata("票据种类名称"))
                .Cell(flexcpData, .Rows - 1, .ColIndex("票据平台数据-票据代码")) = Nvl(clldata("业务标识"))
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-票据代码")) = Nvl(clldata("票据代码"))
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-票据号码")) = Nvl(clldata("票据号码"))
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-开票金额")) = FormatEx(Val(Nvl(clldata("开票金额"))), 6, , , 2)
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-开票时间")) = Format(CDateEx(Nvl(clldata("开票时间"))), "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-数据类型")) = Decode(Val(Nvl(clldata("数据类型"))), 1, "正常电子", 2, "电子红票", 3, "换开纸质", 4, "换开纸质红票", 5, "空白纸质")
                .Cell(flexcpData, .Rows - 1, .ColIndex("票据平台数据-数据类型")) = Val(Nvl(clldata("数据类型")))
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-票据状态")) = Decode(Val(Nvl(clldata("票据状态"))), 1, "正常", 2, "作废", 3, "冲红")
                .Cell(flexcpData, .Rows - 1, .ColIndex("票据平台数据-票据状态")) = Val(Nvl(clldata("票据状态")))
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-关联票据代码")) = Nvl(clldata("关联票据代码"))
                .TextMatrix(.Rows - 1, .ColIndex("票据平台数据-关联票据号码")) = Nvl(clldata("关联票据号码"))
                
                blnChecked = False
                rsUpdate.Filter = "业务流水号='" & Nvl(clldata("业务流水号")) & "' And 票据状态=" & Val(Nvl(clldata("票据状态")))
                If Not rsUpdate.EOF Then
                    blnChecked = Val(Nvl(rsUpdate!修正结果)) = 1
                    .TextMatrix(.Rows - 1, .ColIndex("修正方式")) = Decode(Val(Nvl(rsUpdate!修正方式)), 1, "1-作废HIS数据", 2, "2-作废平台数据", 3, "3-作废数据重开票据", 4, "4-不修正仅标记")
                    .TextMatrix(.Rows - 1, .ColIndex("修正人")) = Nvl(rsUpdate!修正人)
                    .TextMatrix(.Rows - 1, .ColIndex("修正时间")) = Format(Nvl(rsUpdate!修正时间), "yyyy-MM-dd HH:mm:ss")
                    .TextMatrix(.Rows - 1, .ColIndex("修正结果")) = IIf(Val(Nvl(rsUpdate!修正结果)) = 1, "修正成功", "修正失败")
                    .TextMatrix(.Rows - 1, .ColIndex("修正说明")) = Nvl(rsUpdate!修正说明)
                End If
                
                If blnChecked Then
                    .TextMatrix(.Rows - 1, .ColIndex("核对结果")) = "核对成功"
                    .TextMatrix(.Rows - 1, .ColIndex("核对说明")) = ""
                Else
                    .TextMatrix(.Rows - 1, .ColIndex("核对结果")) = "核对失败"
                    .TextMatrix(.Rows - 1, .ColIndex("核对说明")) = "HIS票据号码：无/平台票据号码：" & Nvl(clldata("票据号码"))
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
                    .TextMatrix(.Rows - 1, .ColIndex("修正方式")) = IIf(Val(Nvl(clldata("票据状态"))) = 1 And Val(Nvl(clldata("数据类型"))) = 1, "2-作废平台数据", "4-不修正仅标记")
                End If
            Next
        End If
        
        If .Rows > .FixedRows And .Cols > .FixedCols Then     '缺省定位行
            .Row = -1 '保证在选择行不变的情况下也触发RowColChange事件
            .Row = .FixedRows
        End If
            
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHandler:
    vsfDetailCheck.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    lbl业务日期.Caption = IIf(mbyt票据核对时间类型 = 0, "开票日期", "收费日期")
    
    Call InitTotalCheckGrid
    Call InitDetailCheckGrid
    
    Call load开票点(cbo开票点, mrs开票点, mrs收费员)
    Call Get开票点编码(UserInfo.ID, OS.ComputerName, mstrEInvoiceNodeCode)
    
    dtp结束时间.Value = zlDatabase.Currentdate
    dtp开始时间.Value = DateAdd("d", -7, dtp结束时间.Value)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    sccTitle.Move 8, 8, shpBorder.Width - 20
    picMain.Move sccTitle.Left, sccTitle.Top + sccTitle.Height, Me.ScaleWidth - 2 * sccTitle.Left, Me.ScaleHeight - (2 * sccTitle.Top + sccTitle.Height)
End Sub

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    If Button <> vbLeftButton Then Exit Sub
    If vsfTotalCheck.Height + Y < 1200 Or vsfDetailCheck.Height - Y < 1200 Then Exit Sub

    fraSplit.Top = fraSplit.Top + Y
    
    vsfTotalCheck.Height = vsfTotalCheck.Height + Y
    vsfDetailCheck.Top = vsfDetailCheck.Top + Y
    vsfDetailCheck.Height = vsfDetailCheck.Height - Y
    Me.Refresh
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    picFilter.Move 0, 0, picMain.ScaleWidth
    vsfTotalCheck.Move 0, picFilter.Top + picFilter.Height, picMain.ScaleWidth + 20, picMain.ScaleHeight * 2 / 3
    fraSplit.Move 0, vsfTotalCheck.Top + vsfTotalCheck.Height, picMain.ScaleWidth + 20
    vsfDetailCheck.Move 0, fraSplit.Top + fraSplit.Height, picMain.ScaleWidth + 20, picMain.ScaleHeight - (fraSplit.Top + fraSplit.Height) + 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmMain = Nothing
    Set mcbsMain = Nothing
    Set mobjEInvoice = Nothing
    
    Set mrs开票点 = Nothing
    Set mrs收费员 = Nothing
End Sub

Private Function InitTotalCheckGrid() As Boolean
    '初始化VSFGrid表格控件
    Dim strHead As String, varData As Variant
    Dim strHead0 As String, varData0 As Variant
    Dim i As Integer

    On Error GoTo ErrHandler
    '列名1,对齐方式1,列宽1|列名2,对齐方式2,列宽2|...
    strHead = "业务日期,4,1000|开票数,7,1000|开票金额,7,1200" & _
                    "|开票数,7,1000|开票金额,7,1200|总笔数,7,1000" & _
                    "|核对结果,1,1000|核对说明,1,6000" & _
                    "|上次核对人,1,1000|上次核对时间,1,2000|上次核对结果,1,1200|上次核对说明,1,6000"
    strHead0 = "*|HIS数据|HIS数据|票据平台数据|票据平台数据|票据平台数据|*|*|*|*|*|*"
    With vsfTotalCheck
        .Redraw = flexRDNone '暂停表格显示刷新
        .Clear
        .Rows = 3
        .FixedRows = 2: .FixedCols = 0

        varData = Split(strHead, "|"): varData0 = Split(strHead0, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = IIf(varData0(i) = "*", Split(varData(i), ",")(0), varData0(i))
            .TextMatrix(1, i) = Split(varData(i), ",")(0)
            .ColKey(i) = IIf(varData0(i) = "*", "", varData0(i) & "-") & Split(varData(i), ",")(0) '设置Key值,用于根据 ColIndex() 确定列
            .ColWidth(i) = Split(varData(i), ",")(2)
            If .ColWidth(i) = 0 Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = Split(varData(i), ",")(1)
        Next
        .Cell(flexcpText, 0, .ColIndex("业务日期"), 1, .ColIndex("业务日期")) = IIf(mbyt票据核对时间类型 = 0, "开票日期", "收费日期")

        .AllowSelection = False '不允许多选
        .AllowBigSelection = False '不允许点击固定行/列选择整行/整列
        .SelectionMode = flexSelectionByRow '按行选择
        .AllowUserResizing = flexResizeColumns '允许用户调整列宽
        .BackColorSel = &HE0E0E0
        .ForeColorSel = vbBlack
        
        .MergeCellsFixed = flexMergeFree
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .MergeRow(1) = True
        
        .MergeCol(.ColIndex("业务日期")) = True
        .MergeCol(.ColIndex("核对结果")) = True
        .MergeCol(.ColIndex("核对说明")) = True
        .MergeCol(.ColIndex("上次核对人")) = True
        .MergeCol(.ColIndex("上次核对时间")) = True
        .MergeCol(.ColIndex("上次核对结果")) = True
        .MergeCol(.ColIndex("上次核对说明")) = True

        .RowHeightMin = 300
        .Cell(flexcpBackColor, .FixedRows, 0, .Rows - 1, .ColIndex("HIS数据-开票金额")) = .BackColorFixed
        
        Call ShowTotalRow
        
        .Redraw = flexRDBuffered '刷新表格显示
    End With
    InitTotalCheckGrid = True
    Exit Function
ErrHandler:
    vsfTotalCheck.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitDetailCheckGrid() As Boolean
    '初始化VSFGrid表格控件
    Dim strHead As String, varData As Variant
    Dim strHead0 As String, varData0 As Variant
    Dim i As Integer

    On Error GoTo ErrHandler
    '列名1,对齐方式1,列宽1|列名2,对齐方式2,列宽2|...
    strHead = "收费时间,4,1900|开票点,1,1000|票据类型,4,1000|票据状态,4,1000|单据号,1,1000|票据代码,1,2000|票据号码,1,2000|开票金额,7,1200|开票时间,4,1900" & _
                    "|业务流水号,1,1000|开票点,1,1000|票据代码,1,2000|票据号码,1,2000|开票金额,7,1200" & _
                    "|开票时间,4,1000|数据类型,1,1000|票据状态,1,1000|关联票据代码,1,2000|关联票据号码,1,2000" & _
                    "|核对结果,1,1000|核对说明,1,6000|修正方式,1,2000|修正人,1,1000|修正时间,4,1900|修正结果,1,2000|修正说明,1,6000"
    strHead0 = "HIS数据|HIS数据|HIS数据|HIS数据|HIS数据|HIS数据|HIS数据|HIS数据|HIS数据" & _
                    "|票据平台数据|票据平台数据|票据平台数据|票据平台数据|票据平台数据" & _
                    "|票据平台数据|票据平台数据|票据平台数据|票据平台数据|票据平台数据|*|*|*|*|*|*|*"
    With vsfDetailCheck
        .Redraw = flexRDNone '暂停表格显示刷新
        .Clear
        .Rows = 3
        .FixedRows = 2: .FixedCols = 0

        varData = Split(strHead, "|"): varData0 = Split(strHead0, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = IIf(varData0(i) = "*", Split(varData(i), ",")(0), varData0(i))
            .TextMatrix(1, i) = Split(varData(i), ",")(0)
            .ColKey(i) = IIf(varData0(i) = "*", "", varData0(i) & "-") & Split(varData(i), ",")(0) '设置Key值,用于根据 ColIndex() 确定列
            .ColWidth(i) = Split(varData(i), ",")(2)
            If .ColWidth(i) = 0 Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = Split(varData(i), ",")(1)
        Next

        .AllowSelection = False '不允许多选
        .AllowBigSelection = False '不允许点击固定行/列选择整行/整列
        .SelectionMode = flexSelectionByRow '按行选择
        .AllowUserResizing = flexResizeColumns '允许用户调整列宽
        .BackColorSel = &HE0E0E0
        .ForeColorSel = vbBlack

        .MergeCellsFixed = flexMergeFree
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(.ColIndex("核对结果")) = True
        .MergeCol(.ColIndex("核对说明")) = True
        .MergeCol(.ColIndex("修正方式")) = True
        .MergeCol(.ColIndex("修正人")) = True
        .MergeCol(.ColIndex("修正时间")) = True
        .MergeCol(.ColIndex("修正结果")) = True
        .MergeCol(.ColIndex("修正说明")) = True

        .RowHeightMin = 300
        .Cell(flexcpBackColor, .FixedRows, 0, .Rows - 1, .ColIndex("HIS数据-开票金额")) = .BackColorFixed

        .Redraw = flexRDBuffered '刷新表格显示
    End With
    InitDetailCheckGrid = True
    Exit Function
ErrHandler:
    vsfDetailCheck.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfDetailCheck_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnPrinting Then Exit Sub
    If OldRow = NewRow Then Exit Sub
    
    On Error Resume Next
    vsfDetailCheck.ForeColorSel = vsfDetailCheck.CellForeColor
End Sub

Private Sub vsfDetailCheck_GotFocus()
    Call SetActiveList(vsfDetailCheck)
End Sub

Private Sub vsfDetailCheck_LostFocus()
    Call SetActiveList(vsfDetailCheck, False)
End Sub

Private Sub vsfDetailCheck_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not (Me.ActiveControl Is vsfDetailCheck And Button = vbRightButton) Then Exit Sub
    RaiseEvent ShowPopupMenu(True)
End Sub

Private Sub vsfTotalCheck_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnPrinting Then Exit Sub
    If OldRow = NewRow Or NewRow < vsfTotalCheck.FixedRows Then Exit Sub
    vsfDetailCheck.Clear 1
    vsfDetailCheck.Rows = vsfDetailCheck.FixedRows + 1
    
    On Error Resume Next
    vsfTotalCheck.ForeColorSel = vsfTotalCheck.CellForeColor
End Sub

Private Sub vsfTotalCheck_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not (Me.ActiveControl Is vsfTotalCheck And Button = vbRightButton) Then Exit Sub
    RaiseEvent ShowPopupMenu(True)
End Sub

Private Sub vsfTotalCheck_GotFocus()
    Call SetActiveList(vsfTotalCheck)
End Sub

Private Sub vsfTotalCheck_LostFocus()
    Call SetActiveList(vsfTotalCheck, False)
End Sub

Private Sub SetActiveList(vsfGrid As VSFlexGrid, Optional ByVal blnGetFocus As Boolean = True)
    '设置控件选择行背景高亮色
    If blnGetFocus Then
        vsfTotalCheck.BackColorSel = &HE0E0E0
        vsfDetailCheck.BackColorSel = &HE0E0E0

        If vsfGrid Is Nothing Then Exit Sub
        vsfGrid.BackColorSel = &H8000000D '&HC0C0C0
    Else
        If vsfGrid Is Nothing Then Exit Sub
        vsfGrid.BackColorSel = &HE0E0E0
    End If
End Sub

Private Function Get开票点编码(ByVal lng人员id As Long, ByVal str客户端 As String, ByRef str开票点_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取开票点编号
    '入参:
    '出参:str开票点_Out-返回开票点编号
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2020-03-20 15:13:36
    '说明：如果开票点未设置对码，则以操作员编码为准
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp  As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select 1 From 票据开票点对照 where Rownum <2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取开票点对照信息", lng人员id, str客户端)
     
    If rsTemp.RecordCount = 0 Then  '未配置，缺省为当前操作员编号
        str开票点_Out = UserInfo.编号
        Get开票点编码 = UserInfo.编号 <> "": Exit Function
    End If
    

    strSQL = "" & _
    "   Select  nvl(A.人员ID,0) as 人员ID,nvl(A.客户端,'-')  as 客户端,A.开票点ID,b.编码 as 开票点编码,B.名称  " & _
    "   From 票据开票点对照 A,电子票据开票点 B " & _
    "   Where A.开票点ID=B.ID And nvl(B.撤档时间,sysdate+1)>=SysDate And (a.人员ID=[1] Or a.客户端=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取开票点对照信息", UserInfo.ID, str客户端)
    
    '人员ID+客户端
    rsTemp.Filter = "人员ID=" & lng人员id & " And 客户端='" & str客户端 & "'"
    If rsTemp.EOF = False Then
        str开票点_Out = Nvl(rsTemp!开票点编码)
        Get开票点编码 = str开票点_Out <> ""
        Exit Function
    End If

    '仅收费员
    rsTemp.Filter = "人员ID=" & lng人员id & " And 客户端='-'"
    If rsTemp.EOF = False Then
        str开票点_Out = Nvl(rsTemp!开票点编码)
        Get开票点编码 = str开票点_Out <> ""
        Exit Function
    End If
    
    '客户端
    rsTemp.Filter = "客户端='" & str客户端 & "' And 人员ID=0"
    If rsTemp.EOF = False Then
        str开票点_Out = Nvl(rsTemp!开票点编码)
        Get开票点编码 = str开票点_Out <> ""
        Exit Function
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

