VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInAdviceSend 
   AutoRedraw      =   -1  'True
   Caption         =   "住院临嘱发送"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "frmInAdviceSend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9615
   Begin VB.Frame fraSetup 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   9315
      Begin VB.ComboBox cboDrugType 
         Height          =   300
         Left            =   3045
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.Frame fraBaby 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6120
         TabIndex        =   10
         Top             =   50
         Width           =   3195
         Begin VB.OptionButton optBaby 
            Caption         =   "病人医嘱"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   13
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "所有医嘱"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "婴儿医嘱"
            Height          =   180
            Index           =   2
            Left            =   2175
            TabIndex        =   11
            Top             =   0
            Width           =   1020
         End
      End
      Begin VB.Label lblDrugType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "药嘱分类"
         Height          =   180
         Left            =   2280
         TabIndex        =   15
         Top             =   45
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   6615
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "100%"
      Top             =   6255
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.ProgressBar psb 
      Height          =   270
      Left            =   2115
      TabIndex        =   3
      Top             =   6210
      Visible         =   0   'False
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   6150
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmInAdviceSend.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12091
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   25
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   25
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmInAdviceSend.frx":0E1E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmInAdviceSend.frx":1458
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
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
   Begin VB.Frame fraUD 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   5
      Top             =   4605
      Width           =   9495
   End
   Begin VB.Frame fraInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   60
      TabIndex        =   6
      Top             =   525
      Width           =   9435
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0FFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   60
         Width           =   90
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   3405
      Left            =   0
      TabIndex        =   0
      Top             =   1185
      Width           =   9540
      _cx             =   16828
      _cy             =   6006
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
      BackColorSel    =   16771802
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInAdviceSend.frx":1A92
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
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrice 
      Height          =   1470
      Left            =   0
      TabIndex        =   1
      Top             =   4665
      Width           =   9525
      _cx             =   16801
      _cy             =   2593
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
      BackColorSel    =   4210752
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
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
   Begin MSComCtl2.DTPicker dkpExecTime 
      Height          =   300
      Left            =   2880
      TabIndex        =   8
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   216727555
      CurrentDate     =   40976
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   105
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmInAdviceSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng病人ID As Long 'IN
Private mlng主页ID As Long 'IN
Private mstr前提IDs As String 'IN
Private mlng病人病区ID As Long 'IN
Private mlng病人科室id As Long 'IN
Private mlng医技科室ID As Long 'IN
Private mblnSend As Boolean 'OUT:是否成功发送过。
Private mblnRefresh As Boolean 'OUT'是否需要刷新主界面
Private mstr姓名 As String
Private mstr住院号 As String
Private mstr床号 As String
Private mstr性别 As String

Private mlngNOSequence As Long
Private mcolStock1 As Collection '存放各个药品库房的出库检查方式
Private mcolStock2 As Collection '存放各个卫材库房的出库检查方式
Private mrsPati As ADODB.Recordset '包含病人信息
Private mrsPrice As ADODB.Recordset '包含计价关系
Private mrsBill As ADODB.Recordset
Private mrsRXKey As ADODB.Recordset

Private mstrLike As String
Private mblnFirst As Boolean
Private mint简码 As Integer
Private mint险类 As Integer
Private mbln领药号 As Boolean
Private mstr领药号 As String
Private mlng领药部门ID As Long
Private mblnAutoExe As Boolean
Private mbln一并给药发送为一张 As Boolean '一并给药中药品对应的处方笺不同时，是否仍发送为一张单据
Private mlngRefModld As Long        '0常规医嘱，1=备用医嘱
Private mobjCustom As CommandBarControlCustom
Private mobjlblMsg As CommandBarControl
Private mstr类别s As String
Private mstr药品 As String
Private mbln发送到配置中心 As Boolean  '启用了发送到配置中心的科室，才产生配药记录
Private mbln使用预交 As Boolean '参数：诊间支付允许使用预交款
Private mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mbln检验单独产生单据 As Boolean  '检验医嘱单独产生单据
Private mlng西药房 As Long
Private mlng成药房 As Long
Private mlng中药房 As Long
Private mlng发料部门 As Long
Private mstr药品价格等级 As String '病人的药品价格等级
Private mstr卫材价格等级 As String '病人的卫材价格等级
Private mstr普通项目价格等级 As String '病人的普通项目价格等级
Private mbln阳性用药 As Boolean  '皮试阳性用药 参数，当启用此参数后不用判断皮试结果，但需要填写皮试阳性用药说明
Private mstrAdDrugIDs As String '需进一步添加阳性说明的药品行医嘱ID串儿
Private mbln医嘱处理范围 As Boolean '是否可以分开处理婴儿和病人医嘱

'----------------------------------------------
Private Const COL_选择 = 0
Private Const COL_婴儿 = 1
Private Const col_医嘱内容 = 2
Private Const COL_总量 = 3
Private Const COL_总量单位 = 4
Private Const COL_单量 = 5
Private Const COL_单量单位 = 6
Private Const COL_金额 = 7
Private Const COL_频率 = 8
Private Const COL_用法 = 9
Private Const COL_医生嘱托 = 10 'Data用于存放摘要(医保)
Private Const COL_执行时间 = 11
Private Const COL_执行科室 = 12
Private Const COL_执行性质 = 13
Private Const COL_ID = 14 '隐藏列
Private Const COL_相关ID = 15
Private Const COL_医嘱状态 = 16
Private Const COL_病人科室ID = 17
Private Const COL_开嘱科室ID = 18
Private Const COL_开嘱医生 = 19
Private Const COL_开嘱时间 = 20
Private Const COL_诊疗类别 = 21
Private Const COL_诊疗项目ID = 22
Private Const COL_标本部位 = 23
Private Const COL_检查方法 = 24
Private Const COL_执行标记 = 25
Private Const COL_计价特性 = 26
Private Const COL_执行性质ID = 27
Private Const COL_执行科室ID = 28
Private Const COL_操作类型 = 29
Private Const COL_试管编码 = 30
Private Const COL_收费细目ID = 31
Private Const COL_剂量系数 = 32
Private Const COL_住院包装 = 33
Private Const COL_住院单位 = 34
Private Const COL_可否分零 = 35
Private Const COL_库存 = 36
Private Const COL_次数 = 37
Private Const COL_分解时间 = 38
Private Const COL_首次时间 = 39
Private Const COL_末次时间 = 40
Private Const COL_签名ID = 41
Private Const COL_紧急标志 = 42
Private Const COL_计算方式 = 43
Private Const COL_执行安排 = 44
Private Const COL_开始时间 = 45
Private Const COL_申请序号 = 46
Private Const COL_执行分类 = 47
Private Const COL_跟踪在用 = 48
Private Const COL_毒理分类 = 49
Private Const COL_用药理由 = 50


'-------------------------------------------------
Private Const COLP_行号 = 0
Private Const COLP_收费细目ID = 1
Private Const COLP_固定 = 2
Private Const COLP_变价 = 3
Private Const COLP_计价医嘱 = 4 '可见列
Private Const COLP_类别 = 5
Private Const COLP_收费项目 = 6
Private Const COLP_计价数量 = 7
Private Const COLP_数量 = 8
Private Const COLP_单位 = 9
Private Const COLP_单价 = 10
Private Const COLP_应收金额 = 11
Private Const COLP_实收金额 = 12
Private Const COLP_执行科室 = 13
Private Const COLP_费用类型 = 14
Private Const COLP_从项 = 15
Private Const COLP_收费方式 = 16
Private Const COLP_收费类别 = 17 '隐藏列
Private Const COLP_执行科室ID = 18
Private Const COLP_跟踪在用 = 19
Private Const COLP_费用性质 = 20

Private Enum ESend
    EInBilling = 0  '住院记帐单
    EOutCharge = 1  '门诊收费单
    EOutBilling = 2 '门诊记帐单
End Enum
Private mbytSendKind As ESend 'IN:0-发送住院记帐，1=发送门诊收费,2=发送门诊记帐
Private mlng病人性质 As Long  'IN:0-普通住院病人,1-门诊留观病人,2-住院留观病人
Private mint医嘱处理范围 As Integer    '医嘱处理范围   0-所有医嘱,1-病人医嘱,2-婴儿医嘱
Private mlng医护科室ID As Long

Private Property Let Progress(ByVal vNewValue As Single)
'vNewValue=0-100
    If vNewValue = 0 Then
        psb.value = 0: txtPer.Text = ""
        psb.Visible = False: txtPer.Visible = False
    Else
        psb.value = vNewValue
        txtPer.Text = CInt(psb.value) & "%"
        psb.Visible = True: txtPer.Visible = True
        txtPer.Refresh
    End If
End Property

Public Function ShowMe(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str前提IDs As String, _
    ByVal lng病人病区ID As Long, ByVal lng病人科室ID As Long, ByVal lng医技科室ID As Long, blnRefresh As Boolean, _
    ByVal bytSendKind As Byte, ByVal lng病人性质 As Long, Optional ByVal lng医护科室ID As Long, Optional ByRef objMip As Object) As Boolean
'功能：发送医嘱
'参数：blnRefresh=是否刷新整个主界面
'      bytSendKind=0-发送住院记帐，1=发送门诊收费,2=发送门诊记帐
'      str前提IDs 医技站下达医嘱的前提ID
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstr前提IDs = str前提IDs
    mlng病人病区ID = lng病人病区ID
    mlng病人科室id = lng病人科室ID
    mlng医技科室ID = lng医技科室ID
    mlng医护科室ID = lng医护科室ID
    mbytSendKind = bytSendKind
    mlng病人性质 = lng病人性质
    If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, mlng病人ID, mlng主页ID, "", mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    On Error Resume Next
    Me.Show 1, frmParent
    err.Clear: On Error GoTo 0
    blnRefresh = mblnRefresh
    ShowMe = mblnSend
End Function

Private Sub cboDrugType_Click()
'读取数据
    If Val(cboDrugType.Tag) <> cboDrugType.ListIndex Then
        Call LoadAdviceSend(mstr类别s, mstr药品, 0)
        cboDrugType.Tag = cboDrugType.ListIndex
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Long
    
    If Not Control.Visible Then Exit Sub
    
    Select Case Control.ID
    Case conMenu_Edit_Send
        Call FuncSendAdvice(Control)
    
    Case conMenu_View_Refresh
        mobjCustom.Visible = False
        mobjlblMsg.Visible = False
        Call LoadAdviceSend(mstr类别s, mstr药品, 0)
    Case conMenu_View_RefreshSpare
        mobjCustom.Visible = True
        mobjlblMsg.Visible = True
        Call LoadAdviceSend(mstr类别s, mstr药品, 1)
    Case conMenu_Tool_Option
        With frmInAdviceSendCond
            .Show 1, Me
            If .mblnOK Then
                Call LoadAdviceSend(.mstr类别s, .mstr药品)
                If mbln领药号 Then Call Refresh领药号
            End If
        End With
    Case conMenu_Edit_SelAll
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_选择) = 0 Then
                    Set .Cell(flexcpPicture, i, COL_选择) = frmIcons.imgTrueFalse.ListImages("T").Picture
                End If
            Next
        End With
        Call ShowSendTotal
    Case conMenu_Edit_ClsAll
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_选择) = 0 Then
                    Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                End If
            Next
        End With
        Call ShowSendTotal
    Case conMenu_Help_Help
        ShowHelp App.ProductName, Me.hwnd, Me.Name
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    
    fraInfo.Top = lngTop
    fraInfo.Left = lngLeft
    fraInfo.Width = lngRight - lngLeft
    
    fraSetup.Top = fraInfo.Top + fraInfo.Height
    fraSetup.Left = lngLeft
    fraSetup.Width = lngRight - lngLeft
    
    If fraSetup.Visible Then
        'mbln医嘱处理范围 Or gbln特殊药品分开发送
        If gbln特殊药品分开发送 Then
            fraSetup.Height = 400
            fraBaby.Top = 100
            lblDrugType.Top = 100
            cboDrugType.Top = 50
            cboDrugType.Left = fraInfo.Width - cboDrugType.Width - 50 - IIF(mbln医嘱处理范围, fraBaby.Width, 0)
            lblDrugType.Left = cboDrugType.Left - lblDrugType.Width - 50
        End If
    End If
    
    fraBaby.Left = fraSetup.Width - fraBaby.Width
    
    vsAdvice.Left = lngLeft
    vsAdvice.Top = IIF(fraSetup.Visible, fraSetup.Top + fraSetup.Height, fraInfo.Top + fraInfo.Height)
    vsAdvice.Width = lngRight - lngLeft
    vsAdvice.Height = lngBottom - lngTop - fraInfo.Height - vsPrice.Height - fraUD.Height - stbThis.Height
    
    fraUD.Top = vsAdvice.Top + vsAdvice.Height
    fraUD.Left = lngLeft
    fraUD.Width = vsAdvice.Width
    
    vsPrice.Left = lngLeft
    vsPrice.Top = fraUD.Top + fraUD.Height
    vsPrice.Width = vsAdvice.Width
    
    psb.Top = stbThis.Top + Screen.TwipsPerPixelY * 4
    psb.Width = stbThis.Panels(2).Width - txtPer.Width - Screen.TwipsPerPixelX * 6
    psb.Left = stbThis.Panels(2).Left + Screen.TwipsPerPixelX * 2
    
    txtPer.Left = psb.Left + psb.Width
    txtPer.Top = psb.Top + (psb.Height - txtPer.Height) / 2
 
    Me.Refresh
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_Find
        Control.Visible = mbln领药号
    End Select
End Sub

Private Sub dkpExecTime_Change()
    Call LoadAdviceSend(frmInAdviceSendCond.mstr类别s, frmInAdviceSendCond.mstr药品, 1)
End Sub

Private Sub Form_Activate()
    If mblnFirst Then
        mblnFirst = False
        
        '读取发送清单
        Me.Refresh
        mstr类别s = zlDatabase.GetPara("住院临嘱发送类别", glngSys, p住院医嘱下达)
        mstr药品 = zlDatabase.GetPara("住院临嘱药品发送类别", glngSys, p住院医嘱下达)
        If LoadAdviceSend(mstr类别s, mstr药品) Then
            If mbln领药号 Then Call Refresh领药号
        Else
            Unload Me: Exit Sub
        End If
    End If
End Sub

Private Function GetPatiInfo() As Boolean
'功能：读取病人信息
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = _
        " Select 病人ID,预交余额,费用余额,0 as 预结费用 From 病人余额 Where 性质=1 And 病人ID=[1] And 类型 = " & IIF(mlng病人性质 = 1, 1, 2) & _
        " Union ALL" & _
        " Select A.病人ID,0,0,Sum(金额) From 保险模拟结算 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.险类 Is Not Null And A.病人ID=[1] And A.主页ID=[2] Group by A.病人ID"
    strSQL = "Select 病人ID,Nvl(Sum(预交余额),0)-Nvl(Sum(费用余额),0)+Nvl(Sum(预结费用),0) as 剩余款 From (" & strSQL & ") Group by 病人ID"
    
    strSQL = "Select A.门诊号,B.住院号,Nvl(B.姓名,A.姓名) 姓名,Nvl(B.性别,A.性别) 性别,Nvl(B.年龄,A.年龄) 年龄,B.出院病床 as 床号," & _
        " B.费别,B.医疗付款方式,B.险类,C.剩余款," & _
        " B.状态,B.病人类型,zl_PatiWarnScheme(A.病人ID,B.主页ID) as 适用病人," & _
        " Decode(A.担保额,Null,Null,zl_PatientSurety(A.病人ID,B.主页ID)) as 担保额,a.家庭电话 as PhoneNO," & _
        "To_Char(A.出生日期,'YYYY-MM-DD HH24:MI:SS') as Birthdate,a.家庭地址 as Address" & _
        " From 病人信息 A,病案主页 B,(" & strSQL & ") C" & _
        " Where A.病人ID=B.病人ID And A.病人ID=C.病人ID(+)" & _
        " And B.病人ID=[1] And B.主页ID=[2]"
    Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    
    lblPati.Caption = _
        IIF(mlng病人性质 = 1, "门诊号:" & NVL(mrsPati!门诊号), "住院号:" & NVL(mrsPati!住院号)) & "　姓名:" & mrsPati!姓名 & "　性别:" & NVL(mrsPati!性别) & "　年龄:" & NVL(mrsPati!年龄) & _
        "　床号:" & NVL(mrsPati!床号) & "　费别:" & NVL(mrsPati!费别) & "　类型:" & NVL(mrsPati!病人类型) & "　付款方式:" & NVL(mrsPati!医疗付款方式) & _
        "　剩余款:" & Format(NVL(mrsPati!剩余款, 0), "0.00")
    mint险类 = NVL(mrsPati!险类, 0)
    mstr姓名 = mrsPati!姓名 & ""
    mstr住院号 = NVL(mrsPati!住院号)
    mstr床号 = NVL(mrsPati!床号)
    mstr性别 = mrsPati!姓名 & ""
    '保险病人用红色显示
    If Not IsNull(mrsPati!险类) Then lblPati.ForeColor = zlDatabase.GetPatiColor(NVL(mrsPati!病人类型))
    GetPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then '切换输入法
        If stbThis.Panels("WB").Visible And stbThis.Panels("PY").Visible Then
            If stbThis.Panels("WB").Bevel = sbrRaised Then
                Call stbThis_PanelClick(stbThis.Panels("WB"))
            Else
                Call stbThis_PanelClick(stbThis.Panels("PY"))
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox
    Dim strSQL As String
    Dim strPar As String
    Dim blnDo As Boolean
    
    If Not PatiFeeUsable(mlng病人ID, mlng主页ID) Then Unload Me: Exit Sub
    '工具栏----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlcommfun.GetPubIcons
    
    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "发送")
        If mbytSendKind = EInBilling Then
            objControl.Caption = "发送住院记帐"
        ElseIf mbytSendKind = EOutCharge Then
            objControl.Caption = "发送门诊收费"
        Else
            objControl.Caption = "发送门诊记帐"
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "读取常规医嘱")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_RefreshSpare, "读取备用医嘱")
            
        Set mobjlblMsg = .Add(xtpControlLabel, conMenu_View_RefreshSpare * 100# + 1, "【执行时间】:")
        mobjlblMsg.Visible = False
        Set mobjCustom = .Add(xtpControlCustom, conMenu_View_RefreshSpare * 100# + 2, "")
        mobjCustom.ToolTipText = "请输入该备用医嘱执行的时间。"
        dkpExecTime.value = zlDatabase.Currentdate
        mobjCustom.Handle = dkpExecTime.hwnd
        mobjCustom.Visible = False
        
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "选项")
            objControl.BeginGroup = True
            objControl.IconId = conMenu_File_Parameter
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SelAll, "全选")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ClsAll, "全清")
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKey1, conMenu_Edit_Send
        .Add 0, vbKeyF2, conMenu_Edit_Send
        .Add 0, vbKeyF12, conMenu_Tool_Option
        .Add FCONTROL, vbKeyA, conMenu_Edit_SelAll
        .Add FCONTROL, vbKeyR, conMenu_Edit_ClsAll
        .Add 0, vbKeyF1, conMenu_Help_Help
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add 0, vbKeyF5, conMenu_View_Refresh
    End With
        
    '主菜单右侧的领药号
    objBar.EnableDocking xtpFlagStretched
    With objBar.Controls
        Set objCbo = .Add(xtpControlComboBox, conMenu_View_Find, "领药号")
            objCbo.BeginGroup = True
            objCbo.Flags = xtpFlagRightAlign
            objCbo.Style = xtpComboLabel
            objCbo.Width = 200
    End With
    '-----------------------------------------------------
    Call InitAdviceTable
    Call InitPriceTable
    Call RestoreWinState(Me, App.ProductName)
    
    mbln领药号 = Val(zlDatabase.GetPara(27, glngSys)) <> 0
    If mstr前提IDs <> "" Then
        mlng领药部门ID = mlng医技科室ID
    Else
        mlng领药部门ID = GetICUDeptID
        If mlng领药部门ID = 0 Then mlng领药部门ID = IIF(mlng病人病区ID <> 0, mlng病人病区ID, mlng病人科室id)
    End If
    
    mbln一并给药发送为一张 = Val(zlDatabase.GetPara("一并给药发送为一张", glngSys, p门诊医嘱下达, 1)) = 1
    mblnAutoExe = Val(Mid(zlDatabase.GetPara("本科执行自动完成", glngSys, p住院医嘱发送), 2, 1)) <> 0
    mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    mint简码 = Val(zlDatabase.GetPara("简码方式")) '简码匹配方式：0-拼音,1-五笔
    mbln使用预交 = Val(zlDatabase.GetPara("诊间支付允许使用预交款", glngSys, p门诊医嘱下达, "1"))
    
    mbln发送到配置中心 = gstr输液配置中心 <> ""
    If mbln发送到配置中心 Then
        strPar = zlDatabase.GetPara("来源病区", glngSys, p输液配置中心, "")
        If strPar <> "" Then
            If InStr("," & strPar & ",", "," & mlng病人科室id & ",") = 0 Then mbln发送到配置中心 = False
        End If
        strPar = Val(zlDatabase.GetPara("医嘱类型", glngSys, p输液配置中心, "1"))
        If Val(strPar) = 1 Then
            mbln发送到配置中心 = False
        End If
    End If
    
    mbln检验单独产生单据 = Val(zlDatabase.GetPara("检验医嘱单独产生单据", glngSys, p住院医嘱发送, "0")) = 1
    mlng中药房 = Val(zlDatabase.GetPara("住院缺省中药房", glngSys, p住院医嘱下达, , , , , mlng病人科室id))
    mlng西药房 = Val(zlDatabase.GetPara("住院缺省西药房", glngSys, p住院医嘱下达, , , , , mlng病人科室id))
    mlng成药房 = Val(zlDatabase.GetPara("住院缺省成药房", glngSys, p住院医嘱下达, , , , , mlng病人科室id))
    mlng发料部门 = Val(zlDatabase.GetPara("住院缺省发料部门", glngSys, p住院医嘱下达, , , , , mlng病人科室id))
    
    mbln阳性用药 = Val(zlDatabase.GetPara("皮试阳性用药", glngSys, p住院医嘱下达)) <> 0
    
    cboDrugType.AddItem "0-全部"
    cboDrugType.AddItem "1-毒品类"
    cboDrugType.AddItem "2-麻醉和精神I类"
    cboDrugType.AddItem "3-其它(非1和2类)"
    Call Cbo.SetIndex(cboDrugType.hwnd, 0)
    cboDrugType.Tag = "0"
  
    cboDrugType.Visible = gbln特殊药品分开发送
    lblDrugType.Visible = gbln特殊药品分开发送
    
    mbln医嘱处理范围 = DeptIsWoman(0, Get科室IDs(mlng病人病区ID))
    If mbln医嘱处理范围 Then
        fraSetup.Visible = True
        '医嘱处理范围
        mint医嘱处理范围 = Val(zlDatabase.GetPara("医嘱处理范围", glngSys, p住院医嘱发送, "0"))
        optBaby(mint医嘱处理范围).value = True
    End If
 
    fraSetup.Visible = (mbln医嘱处理范围 Or gbln特殊药品分开发送)
    fraBaby.Visible = mbln医嘱处理范围
    Select Case mint简码
        Case 0
            stbThis.Panels("PY").Bevel = sbrInset
            stbThis.Panels("WB").Bevel = sbrRaised
        Case 1
            stbThis.Panels("PY").Bevel = sbrRaised
            stbThis.Panels("WB").Bevel = sbrInset
        Case Else
            stbThis.Panels("PY").Bevel = sbrInset
            stbThis.Panels("WB").Bevel = sbrInset
    End Select
    If Not gbln简码匹配方式切换 Then
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
   
    mblnSend = False
    mblnRefresh = False
    mblnFirst = True
    
    '各个库房药品出库检查方式
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    
    '显示病人信息
    If Not GetPatiInfo Then Unload Me: Exit Sub
    
    '一卡通结算部件
    If mblnAutoExe And gbln执行前先结算 And Not mbytSendKind = EInBilling Then
        If gobjSquareCard Is Nothing Then
            On Error Resume Next
            Set gobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
            err.Clear: On Error GoTo 0
            
            If Not gobjSquareCard Is Nothing Then
                If gobjSquareCard.zlInitComponents(Me, p住院医嘱下达, glngSys, gstrDBUser, gcnOracle, False) = False Then
                    Set gobjSquareCard = Nothing
                    MsgBox "卡结算部件（zl9CardSquare）初始化失败!本科自动执行的医嘱不会被自动执行。", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
End Sub

Private Function TheStockCheck(ByVal lng库房ID As Long, ByVal str类别 As String) As Integer
'功能：获取指定库房的出库库存检查方式
    Dim intStyle As Integer
    On Error Resume Next
    If InStr(",5,6,7,", str类别) > 0 Then
        intStyle = mcolStock1("_" & lng库房ID)
    ElseIf str类别 = "4" Then
        intStyle = mcolStock2("_" & lng库房ID)
    End If
    err.Clear: On Error GoTo 0
    TheStockCheck = intStyle
End Function

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    '释放私有及IN变量
    mlng主页ID = 0
    mlng病人ID = 0
    Set mrsPati = Nothing
    Set mrsPrice = Nothing
    Set mrsBill = Nothing
    Set mcolStock1 = Nothing
    Set mcolStock2 = Nothing
    Set mrsRXKey = Nothing
    Set mobjCustom = Nothing
    Set mobjlblMsg = Nothing
    Set mclsMipModule = Nothing
    gbln加班加价 = False
End Sub

Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsAdvice.Height + Y < 1000 Or vsPrice.Height - Y < 500 Then Exit Sub
        fraUD.Top = fraUD.Top + Y
        vsAdvice.Height = vsAdvice.Height + Y
        vsPrice.Top = vsPrice.Top + Y
        vsPrice.Height = vsPrice.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub optBaby_Click(Index As Integer)
    mint医嘱处理范围 = Index
    Call LoadAdviceSend(mstr类别s, mstr药品, 0)
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '切换并保存简码匹配方式
        Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            stbThis.Panels("WB").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            stbThis.Panels("PY").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        Call zlDatabase.SetPara("简码方式", IIF(stbThis.Panels("PY").Bevel = sbrInset And stbThis.Panels("WB").Bevel = sbrInset, 2, IIF(stbThis.Panels("WB").Bevel = sbrInset, 1, 0)))
        mint简码 = Val(zlDatabase.GetPara("简码方式")) '简码匹配方式：0-拼音,1-五笔
    End If
End Sub

Private Sub FuncSendAdvice(ByVal Control As CommandBarControl)
    Dim lng发送号 As Long, strMsg As String, i As Long
    Dim objCbo As CommandBarComboBox
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_ID)) <> 0 And .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                Exit For
            End If
        Next
        If i > .Rows - 1 Then
            MsgBox "当前没有可以发送的医嘱。", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    If mbytSendKind = EInBilling Then
        strMsg = "本次医嘱发送的费用将产生为住院记帐单据，你确定吗？"
    ElseIf mbytSendKind = EOutCharge Then
        strMsg = "本次医嘱发送的费用将产生为门诊收费单据，你确定吗？"
    Else
        strMsg = "本次医嘱发送的费用将产生为门诊记帐单据，你确定吗？"
    End If
    If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
    
    lng发送号 = SendAdvice()
    If lng发送号 <> 0 Then
        mblnSend = True
        
        '使用了新领药号的处理
        If mstr领药号 <> "" Then
            Set objCbo = cbsMain.FindControl(, conMenu_View_Find)
            i = objCbo.FindItem(mstr领药号)
            If i = 0 Then
                objCbo.AddItem mstr领药号, 2
                objCbo.ListIndex = 2
            End If
        End If
            
        '打印诊疗单据
        Call frmSendBillPrint.ShowMe(lng发送号, 2, Me, mstr前提IDs)
        
        '如果全部发送完毕,则退出
        If vsAdvice.Rows = 2 Then
            If Val(vsAdvice.TextMatrix(1, COL_ID)) = 0 Then
                Unload Me: Exit Sub
            End If
        End If
        Call GetPatiInfo
    End If
End Sub

Private Sub RowSelectSame(ByVal lngRow As Long, ByVal lngCol As Long, _
    Optional rsSQL As ADODB.Recordset, Optional rsTotal As ADODB.Recordset, _
    Optional rsUpload As ADODB.Recordset, Optional str医嘱IDs As String)
'功能：根据可见行的选择状态,将相关医嘱一并选择
    Dim i As Long
    
    With vsAdvice
        If lngCol = COL_选择 Then
            For i = lngRow + 1 To .Rows - 1
                If IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID))) _
                    = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) <> 0, Val(.TextMatrix(lngRow, COL_相关ID)), Val(.TextMatrix(lngRow, COL_ID))) Then
                    .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                    Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                Else
                    Exit For
                End If
            Next
            
            For i = lngRow - 1 To .FixedRows Step -1
                If IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID))) _
                    = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) <> 0, Val(.TextMatrix(lngRow, COL_相关ID)), Val(.TextMatrix(lngRow, COL_ID))) Then
                    .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                    Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                Else
                    Exit For
                End If
            Next
            
            '一组申请序号的会诊医嘱
            If Val(.TextMatrix(lngRow, COL_申请序号)) <> 0 And .TextMatrix(lngRow, COL_诊疗类别) = "Z" Then
                For i = lngRow - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(i, COL_申请序号)) = Val(.TextMatrix(lngRow, COL_申请序号)) Then
                        .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                        Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                    Else
                        Exit For
                    End If
                Next
                
                For i = lngRow + 1 To .Rows - 1
                    If Val(.TextMatrix(i, COL_申请序号)) = Val(.TextMatrix(lngRow, COL_申请序号)) Then
                        .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                        Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                    Else
                        Exit For
                    End If
                Next
            End If
            
            '取消选择时
            If Not (.Cell(flexcpData, lngRow, COL_选择) = 0 And Not .Cell(flexcpPicture, lngRow, COL_选择) Is Nothing) Then
                i = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) = 0, Val(.TextMatrix(lngRow, COL_ID)), Val(.TextMatrix(lngRow, COL_相关ID)))
                '1.清除对应的费用及发送记录填写
                If Not rsSQL Is Nothing Then
                    rsSQL.Filter = "医嘱ID=" & i
                    Do While Not rsSQL.EOF
                        rsSQL.Delete
                        rsSQL.Update
                        rsSQL.MoveNext
                    Loop
                    rsSQL.Filter = 0 '因为要使用BookMark，因此恢复
                End If
                '2.清除对应的发送计价数量累计
                If Not rsTotal Is Nothing Then
                    rsTotal.Filter = "医嘱ID=" & i
                    Do While Not rsTotal.EOF
                        rsTotal.Delete
                        rsTotal.Update
                        rsTotal.MoveNext
                    Loop
                End If
                '3.清除对应的医保上传单据号
                If Not rsUpload Is Nothing Then
                    rsUpload.Filter = "医嘱ID=" & i
                    Do While Not rsUpload.EOF
                        rsUpload.Delete
                        rsUpload.Update
                        rsUpload.MoveNext
                    Loop
                End If
                '4.清除不发送的签名医嘱组ID
                If str医嘱IDs <> "" Then
                    str医嘱IDs = str医嘱IDs & ","
                    str医嘱IDs = Replace(str医嘱IDs, "," & i & ",", ",")
                    If str医嘱IDs <> "" Then
                        str医嘱IDs = Left(str医嘱IDs, Len(str医嘱IDs) - 1)
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function GetVisibleRow(ByVal lngRow As Long, Optional ByVal blnFirst As Boolean) As Long
'功能：根据指定医嘱行，返回该医嘱中可见的行
    Dim lng组ID As Long, i As Long
    
    GetVisibleRow = lngRow
    
    With vsAdvice
        If Not .RowHidden(lngRow) Then Exit Function
        
        '一并给药的定位到第一药品行
        If blnFirst Then
            If .TextMatrix(lngRow, COL_诊疗类别) = "E" And InStr(",5,6,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 _
                And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 And Val(.TextMatrix(lngRow, COL_ID)) = Val(.TextMatrix(lngRow - 1, COL_相关ID)) Then
                i = .FindRow(.TextMatrix(lngRow, COL_ID), , COL_相关ID)
                If i <> -1 Then GetVisibleRow = i: Exit Function
            End If
        End If
        
        lng组ID = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) <> 0, Val(.TextMatrix(lngRow, COL_相关ID)), Val(.TextMatrix(lngRow, COL_ID)))
        For i = lngRow - 1 To .FixedRows Step -1
            If lng组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID))) Then
                If Not .RowHidden(i) Then GetVisibleRow = i: Exit Function
            Else
                Exit For
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            If lng组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID))) Then
                If Not .RowHidden(i) Then GetVisibleRow = i: Exit Function
            Else
                Exit For
            End If
        Next
    End With
End Function

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsAdvice
        If OldRow <> NewRow And .Redraw <> flexRDNone And Not .RowHidden(NewRow) Then
            If Val(.TextMatrix(NewRow, COL_ID)) <> 0 Then
                Call ShowAdvicePrice(NewRow)
                
                '缺省选择计价医嘱(如果可能)
                Call ShowDefaultRow
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserFreeze()
    With vsAdvice
        If .FrozenCols < COL_选择 + 1 - .FixedCols Then
            .FrozenCols = COL_选择 + 1 - .FixedCols
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    With vsAdvice
        If Col = col_医嘱内容 Then
            .AutoSize col_医嘱内容
            .RowHeight(0) = 320
        ElseIf Row = -1 Then
            lngW = Me.TextWidth(.TextMatrix(.FixedRows - 1, Col) & "A")
            If .ColWidth(Col) < lngW Then
                .ColWidth(Col) = lngW
            ElseIf .ColWidth(Col) > .Width * 0.5 Then
                .ColWidth(Col) = .Width * 0.5
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_选择 Then Cancel = True
End Sub

Private Sub vsAdvice_DblClick()
    With vsAdvice
        If .MouseCol = COL_选择 And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsAdvice_KeyPress(32)
        End If
    End With
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        lngLeft = COL_频率: lngRight = COL_用法
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        
        If Not RowIn一并给药(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '底行保留下边线
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim i As Long
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i: Exit For
                End If
            Next
            If i > .Rows - 1 Then .Row = .FixedRows
            Call .ShowCell(.Row, .Col)
        ElseIf KeyAscii = 32 And .Col = COL_选择 Then
            KeyAscii = 0
            If .Cell(flexcpData, .Row, COL_选择) = 0 Then
                If .Cell(flexcpPicture, .Row, COL_选择) Is Nothing Then
                    Set .Cell(flexcpPicture, .Row, COL_选择) = frmIcons.imgTrueFalse.ListImages("T").Picture
                Else
                    Set .Cell(flexcpPicture, .Row, COL_选择) = Nothing
                End If
                Call RowSelectSame(.Row, .Col)
                Call ShowSendTotal
            End If
        End If
    End With
End Sub

Private Sub ShowDefaultRow()
'功能：对于可以计价的医嘱,缺省增加一行并设置缺省计价医嘱
'说明：ComboList="#医嘱ID1;计价医嘱1|#医嘱ID2;计价医嘱2|..."
'      仅在第一次显示计价表和回车新增行时调用
    Dim arrCombo As Variant, lngRow As Long, i As Long
    Dim lng医嘱ID As Long, lng行号 As Long, str计价医嘱 As String
    Dim blnFirst As Boolean, blnHave As Boolean
    
    With vsPrice
        If .ColData(COLP_计价医嘱) <> "" And .Editable <> flexEDNone Then
            arrCombo = Split(.ColData(COLP_计价医嘱), "|")
            
            If Val(.TextMatrix(.Rows - 1, COLP_行号)) <> 0 _
                And Val(.TextMatrix(.Rows - 1, COLP_收费细目ID)) <> 0 Then
                '第一次显示时缺省增加一行
                blnFirst = True
                .AddItem "", .Rows
                .Row = .Rows - 1
            End If
            lngRow = .Rows - 1
            
            '不是第一次显示时缺省计价医嘱与上一行相同
            If lngRow > 1 And Not blnFirst Then
                If Val(.TextMatrix(lngRow - 1, COLP_固定)) = 0 _
                    And Val(.TextMatrix(lngRow - 1, COLP_行号)) <> 0 Then
                    blnHave = True
                End If
            End If
            For i = 0 To UBound(arrCombo)
                lng医嘱ID = Val(Mid(Mid(arrCombo(i), 1, InStr(arrCombo(i), ";") - 1), 2))
                str计价医嘱 = Replace(arrCombo(i), "#" & lng医嘱ID & ";", "")
                lng行号 = vsAdvice.FindRow(CStr(lng医嘱ID), , COL_ID)
                If blnHave Then
                    If lng行号 = Val(.TextMatrix(lngRow - 1, COLP_行号)) Then
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
                        
            '模拟选中这个计价医嘱
            .TextMatrix(lngRow, COLP_行号) = lng行号
            .TextMatrix(lngRow, COLP_计价医嘱) = str计价医嘱
            .Cell(flexcpData, lngRow, COLP_计价医嘱) = .TextMatrix(lngRow, COLP_计价医嘱)
            
            '只有一个计价医嘱时不必停留
            If UBound(arrCombo) = 0 Then
                .Col = COLP_收费项目
            Else
                .Col = COLP_计价医嘱
            End If
        End If
        Call .ShowCell(.Row, .Col)
    End With
End Sub

Private Sub vsPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lng原嘱ID As Long, lng医嘱ID As Long
    Dim int费用性质 As Integer, int原费用性质 As Integer
    Dim lng收费细目ID As Long, i As Long
    Dim blnHaveSub As Boolean
    
    On Error GoTo errH
    
    With vsPrice
        If Col = COLP_计价医嘱 Then
            '如果绑定了ComboData,TextMatrix取值就为ComboData
            If .Cell(flexcpTextDisplay, Row, Col) <> .Cell(flexcpData, Row, Col) Then
                lng医嘱ID = .ComboData
                If lng医嘱ID < 0 Then
                    int费用性质 = Val(Left(Abs(lng医嘱ID), 1))
                    lng医嘱ID = Val(Mid(Abs(lng医嘱ID), 2))
                End If
                lng原嘱ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_行号)), COL_ID))
                int原费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                lng收费细目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                                
                '检查该计价医嘱是否已有相同收费细目
                If lng收费细目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng收费细目ID
                    If Not mrsPrice.EOF Then
                        MsgBox """" & .Cell(flexcpTextDisplay, Row, Col) & """已经设置了收费项目""" & .TextMatrix(Row, COLP_收费项目) & """。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                
                '原来的医嘱如果有从项至少要保留一个(主项是固定不可动的)
                If lng原嘱ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng原嘱ID & " And 费用性质=" & int原费用性质 & " And 从项=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(Row, COLP_从项) <> "" Then
                        MsgBox """" & .Cell(flexcpData, Row, Col) & """至少要保留一个从属计价项目。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col): Exit Sub
                    End If
                End If
                
                '标明输入了的计价医嘱部份
                i = vsAdvice.FindRow(CStr(lng医嘱ID), , COL_ID)
                .TextMatrix(Row, COLP_行号) = i
                .TextMatrix(Row, COLP_费用性质) = int费用性质
                .TextMatrix(Row, Col) = .Cell(flexcpTextDisplay, Row, Col)
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                If lng收费细目ID <> 0 Then
                    '新选择的医嘱是否有从项决定修改后的项目是否从项
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " 费用性质=" & int费用性质 & " And 从项=1"
                    If Not mrsPrice.EOF Then blnHaveSub = True
                    .TextMatrix(Row, COLP_从项) = IIF(blnHaveSub, "√", "")
                
                    '更新或增加记录集内容
                    If lng原嘱ID = 0 Then
                        mrsPrice.AddNew '加入
                    Else '更新
                        mrsPrice.Filter = "医嘱ID=" & lng原嘱ID & " And 费用性质=" & int原费用性质 & " And 收费细目ID=" & lng收费细目ID
                    End If
                    mrsPrice!医嘱ID = lng医嘱ID
                    If Val(vsAdvice.TextMatrix(i, COL_相关ID)) <> 0 Then
                        mrsPrice!相关ID = vsAdvice.TextMatrix(i, COL_相关ID)
                    Else
                        mrsPrice!相关ID = Null
                    End If
                    mrsPrice!费用性质 = int费用性质
                    mrsPrice!收费方式 = 0
                    If lng原嘱ID = 0 Then
                        mrsPrice!收费细目ID = lng收费细目ID
                        mrsPrice!数量 = Val(.TextMatrix(Row, COLP_计价数量))
                        mrsPrice!单价 = Val(.TextMatrix(Row, COLP_单价))
                        mrsPrice!在用 = Val(.TextMatrix(Row, COLP_跟踪在用))
                        mrsPrice!变价 = Val(.TextMatrix(Row, COLP_变价))
                        mrsPrice!固定 = 0
                    End If
                    mrsPrice!从项 = IIF(blnHaveSub, 1, 0)
                    mrsPrice.Update
                    
                    Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                End If
            End If
        ElseIf Col = COLP_收费项目 Or Col = COLP_执行科室 Then
            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
        ElseIf Col = COLP_计价数量 Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '更新记录集
            lng医嘱ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_行号)), COL_ID))
            int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
            lng收费细目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
            If lng医嘱ID <> 0 And lng收费细目ID <> 0 Then
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng收费细目ID
                mrsPrice!数量 = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                
                Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
            End If
        ElseIf Col = COLP_单价 Then
            If Not IsNumeric(.TextMatrix(Row, Col)) _
                Or Val(.TextMatrix(Row, Col)) <= 0 _
                Or Val(.TextMatrix(Row, Col)) > LONG_MAX Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            If CheckScope(.Cell(flexcpData, Row, COLP_应收金额), .Cell(flexcpData, Row, COLP_实收金额), .TextMatrix(Row, Col)) <> "" Then
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gstrDecPrice)
            .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
            
            '更新记录集
            lng医嘱ID = Val(vsAdvice.TextMatrix(Val(.TextMatrix(Row, COLP_行号)), COL_ID))
            int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
            lng收费细目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
            If lng医嘱ID <> 0 And lng收费细目ID <> 0 Then
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng收费细目ID
                mrsPrice!单价 = Val(.TextMatrix(Row, Col))
                mrsPrice.Update
                
                Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngRow As Long
    
    '根据可否编辑设置
    If Not CellEditable(NewRow, NewCol) Then
        vsPrice.ComboList = ""
        vsPrice.FocusRect = flexFocusLight
    Else
        vsPrice.FocusRect = flexFocusSolid
        If NewCol = COLP_计价医嘱 Then
            vsPrice.ComboList = vsPrice.ColData(NewCol)
        ElseIf NewCol = COLP_收费项目 Or NewCol = COLP_执行科室 Then
            vsPrice.ComboList = "..."
        Else
            vsPrice.ComboList = ""
        End If
    End If
        
    If NewRow <> OldRow Then
        '显示药品库存
        With vsPrice
            stbThis.Panels(2).Text = ""
            lngRow = Val(.TextMatrix(NewRow, COLP_行号))
            If lngRow <> 0 And .TextMatrix(NewRow, COLP_收费类别) <> "" Then
                If InStr(",5,6,7,", .TextMatrix(NewRow, COLP_收费类别)) > 0 _
                    Or .TextMatrix(NewRow, COLP_收费类别) = "4" And Val(.TextMatrix(NewRow, COLP_跟踪在用)) = 1 Then
                    '显示药品及跟踪卫材的库存
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                        If InStr(GetInsidePrivs(p住院医嘱下达), "显示药品库存") = 0 Then
                            stbThis.Panels(2).Text = vsAdvice.TextMatrix(lngRow, col_医嘱内容) & "," & vsAdvice.TextMatrix(lngRow, COL_执行科室) & IIF(Val(vsAdvice.TextMatrix(lngRow, COL_库存)) > 0, "有库存", "无库存")
                        Else
                            stbThis.Panels(2).Text = vsAdvice.TextMatrix(lngRow, col_医嘱内容) & "," & vsAdvice.TextMatrix(lngRow, COL_执行科室) & "可用库存:" & FormatEx(Val(vsAdvice.TextMatrix(lngRow, COL_库存)), 5) & vsAdvice.TextMatrix(lngRow, COL_住院单位)
                        End If
                    Else
                        '同一个函数取:药品按住院单位,卫材按售价单位
                        If InStr(GetInsidePrivs(p住院医嘱下达), "显示药品库存") = 0 Then
                            If GetStock(Val(.TextMatrix(NewRow, COLP_收费细目ID)), Val(.TextMatrix(NewRow, COLP_执行科室ID))) > 0 Then
                                stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_收费项目) & "," & .TextMatrix(NewRow, COLP_执行科室) & "有库存"
                            Else
                                stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_收费项目) & "," & .TextMatrix(NewRow, COLP_执行科室) & "无库存"
                            End If
                        Else
                            stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_收费项目) & "," & .TextMatrix(NewRow, COLP_执行科室) & "可用库存:" & _
                                FormatEx(GetStock(Val(.TextMatrix(NewRow, COLP_收费细目ID)), Val(.TextMatrix(NewRow, COLP_执行科室ID))), 5) & .TextMatrix(NewRow, COLP_单位)
                        End If
                    End If
                End If
            End If
        End With
        
        '显示医保大类
        stbThis.Panels(3).Text = Get医保大类(NewRow)
    End If
End Sub

Private Function Get医保大类(ByVal lngRow As Long) As String
'功能：获取指定行的费用类型
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, str大类 As String
    
    With vsPrice
        If Val(.TextMatrix(lngRow, COLP_收费细目ID)) <> 0 Then
            strSQL = "Select N.名称 From 保险支付项目 M,保险支付大类 N Where M.收费细目ID=[1] And M.大类ID=N.ID And M.险类=[2]"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COLP_收费细目ID)), mint险类)
            If Not rsTmp.EOF Then str大类 = NVL(rsTmp!名称)
        End If
    End With
    Get医保大类 = IIF(str大类 <> "", "医保大类:" & str大类, "")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
'说明：返回的行号范围不包括给药途径的行号
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_诊疗类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Sub InitAdviceTable()
'功能：初始化清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = ",300,4;" & _
        "婴儿,550,1;医嘱内容,3000,1;总量,600,7;单位,450,1;单量,600,7;单位,450,1;金额,850,7;" & _
        "频率,1000,1;用法,1000,1;医生嘱托,1500,1;执行时间,1000,1;执行科室,850,1;执行性质,850,1;" & _
        "ID;相关ID;医嘱状态;病人科室ID;开嘱科室ID;开嘱医生;开嘱时间;诊疗类别;诊疗项目ID;标本部位;检查方法;执行标记;计价特性;" & _
        "执行性质ID;执行科室ID;操作类型;试管编码;收费细目ID;剂量系数;住院包装;住院单位;可否分零;库存;次数;分解时间;首次时间;" & _
        "末次时间;签名否;紧急标志;计算方式;执行安排;开始时间;申请序号;执行分类;跟踪在用;毒理分类;用药理由"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .FrozenCols = COL_选择 + 1 - .FixedCols
        .RowHeight(0) = 320
    End With
End Sub

Private Sub InitPriceTable()
'功能：初始化计价清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "行号;收费细目ID;固定;变价;计价医嘱,2000,1;类别,650,1;收费项目,2000,1;计价数量,900,7;" & _
        "数量,800,7;单位,500,1;单价,1000,7;应收金额,1050,7;实收金额,1050,7;执行科室,1000,1;费用类型,850,1;" & _
        "从项,450,4;收费方式,1500,1;收费类别;执行科室ID;跟踪在用;费用性质"
    arrHead = Split(strHead, ";")
    With vsPrice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub DeleteCurRow(ByVal lngRow As Long, Optional ByVal blnDelCur As Boolean = True)
'功能：在处理待发送清单的过程中删除最近加入的行(含药疗或非药)
'参数：blnDelCur=是否删除当前行
    Dim lng医嘱ID As Long, lng相关ID As Long, i As Long
    
    With vsAdvice
        lng医嘱ID = Val(.TextMatrix(lngRow, COL_ID))
        lng相关ID = Val(.TextMatrix(lngRow, COL_相关ID))
                
        '删除当前行
        If blnDelCur Then .RemoveItem lngRow
        
        '删除相关行
        If lng相关ID <> 0 Then
            For i = .Rows - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = lng相关ID _
                    Or Val(.TextMatrix(i, COL_ID)) = lng相关ID Then
                    .RemoveItem i
                End If
            Next
        Else
            For i = .Rows - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = lng医嘱ID Then
                    .RemoveItem i
                End If
            Next
        End If
    End With
End Sub

Private Sub InitSeekSet(rsSeek As ADODB.Recordset)
'功能：初始化用于汇总计算折扣的临时记录集
    Set rsSeek = New ADODB.Recordset
    rsSeek.Fields.Append "费用性质", adInteger
    rsSeek.Fields.Append "主项标签", adVariant
    rsSeek.Fields.Append "主收入ID", adBigInt
    rsSeek.Fields.Append "合计", adCurrency, , adFldIsNullable
    rsSeek.CursorLocation = adUseClient
    rsSeek.LockType = adLockOptimistic
    rsSeek.CursorType = adOpenStatic
    rsSeek.Open
End Sub

Private Sub InitPriceRecordset()
'功能：初始化医嘱计价记录集
    Set mrsPrice = New ADODB.Recordset
    
    mrsPrice.Fields.Append "医嘱ID", adBigInt
    mrsPrice.Fields.Append "相关ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "费用性质", adInteger, , adFldIsNullable
    mrsPrice.Fields.Append "收费方式", adInteger, , adFldIsNullable
    mrsPrice.Fields.Append "收费类别", adVarChar, 1
    mrsPrice.Fields.Append "收费细目ID", adBigInt
    mrsPrice.Fields.Append "执行科室ID", adBigInt, , adFldIsNullable
    mrsPrice.Fields.Append "数量", adDouble
    mrsPrice.Fields.Append "单价", adDouble, , adFldIsNullable '变价价格
    mrsPrice.Fields.Append "在用", adInteger '卫材是否跟踪在用
    mrsPrice.Fields.Append "变价", adInteger
    mrsPrice.Fields.Append "从项", adInteger
    mrsPrice.Fields.Append "固定", adInteger
    
    mrsPrice.CursorLocation = adUseClient
    mrsPrice.LockType = adLockOptimistic
    mrsPrice.CursorType = adOpenStatic
    mrsPrice.Open
End Sub

Private Sub InitRecordSet(rsSQL As ADODB.Recordset, rsTotal As ADODB.Recordset, rsUpload As ADODB.Recordset, _
    rsNumber As ADODB.Recordset, rsMoneyNow As ADODB.Recordset, rsItems As ADODB.Recordset)
'初始化记录集
    'SQL记录集
    Set rsSQL = New ADODB.Recordset
    rsSQL.Fields.Append "类型", adInteger '1-计价,2-签名,3-校对,4-发送,5-费用,6-发料
    rsSQL.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsSQL.Fields.Append "项目ID", adBigInt '收费细目ID
    rsSQL.Fields.Append "序号", adBigInt '用于排序
    rsSQL.Fields.Append "SQL", adVarChar, 5000 'SQL
    rsSQL.Fields.Append "NO", adVarChar, 30, adFldIsNullable '用于NO替换处理时排序
    rsSQL.CursorLocation = adUseClient
    rsSQL.LockType = adLockOptimistic
    rsSQL.CursorType = adOpenStatic
    rsSQL.Open
    
    '计价数量累计记录集
    Set rsTotal = New ADODB.Recordset
    rsTotal.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsTotal.Fields.Append "项目ID", adBigInt
    rsTotal.Fields.Append "库房ID", adBigInt
    rsTotal.Fields.Append "数量", adDouble
    rsTotal.CursorLocation = adUseClient
    rsTotal.LockType = adLockOptimistic
    rsTotal.CursorType = adOpenStatic
    rsTotal.Open
    
    '医保上传记帐单
    Set rsUpload = New ADODB.Recordset
    rsUpload.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsUpload.Fields.Append "NO", adVarChar, 30
    rsUpload.CursorLocation = adUseClient
    rsUpload.LockType = adLockOptimistic
    rsUpload.CursorType = adOpenStatic
    rsUpload.Open
    
    '计录试管编码
    Set rsNumber = New ADODB.Recordset
    rsNumber.Fields.Append "管码", adVarChar, 18
    rsNumber.Fields.Append "相关ID", adBigInt
    rsNumber.Fields.Append "样本条码", adVarChar, 18
    rsNumber.Fields.Append "执行科室ID", adVarChar, 18
    rsNumber.Fields.Append "诊疗项目ID", adVarChar, 18
    rsNumber.Fields.Append "婴儿", adBigInt
    rsNumber.Fields.Append "紧急标志", adBigInt
    rsNumber.Fields.Append "标本", adVarChar, 18
    rsNumber.Fields.Append "采集科室ID", adBigInt
    rsNumber.CursorLocation = adUseClient
    rsNumber.LockType = adLockOptimistic
    rsNumber.CursorType = adOpenStatic
    rsNumber.Open
    
    '当前病人本次要发送的费用
    Set rsMoneyNow = New ADODB.Recordset
    rsMoneyNow.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsMoneyNow.Fields.Append "诊疗项目ID", adBigInt
    rsMoneyNow.Fields.Append "收费项目ID", adBigInt
    rsMoneyNow.Fields.Append "试管编码", adVarChar, 18, adFldIsNullable
    rsMoneyNow.Fields.Append "样本条码", adVarChar, 50, adFldIsNullable
    rsMoneyNow.Fields.Append "收费方式", adInteger
    rsMoneyNow.Fields.Append "收费时间", adVarChar, 10
    rsMoneyNow.Fields.Append "执行部门ID", adBigInt
    rsMoneyNow.Fields.Append "子医嘱ID", adBigInt '相关ID不为空的医嘱行的医嘱ID
    rsMoneyNow.Fields.Append "检查部位", adVarChar, 100
    rsMoneyNow.Fields.Append "检查方法", adVarChar, 100
    rsMoneyNow.Fields.Append "数量", adDouble '收费数量
    
    rsMoneyNow.CursorLocation = adUseClient
    rsMoneyNow.LockType = adLockOptimistic
    rsMoneyNow.CursorType = adOpenStatic
    rsMoneyNow.Open
    
    '当前病人本次发送的费用项目汇总
    Set rsItems = New ADODB.Recordset
    rsItems.Fields.Append "病人ID", adBigInt
    rsItems.Fields.Append "主页ID", adBigInt, , adFldIsNullable
    rsItems.Fields.Append "医嘱ID", adBigInt
    rsItems.Fields.Append "收费类别", adVarChar, 1
    rsItems.Fields.Append "收费细目ID", adBigInt
    rsItems.Fields.Append "数量", adDouble
    rsItems.Fields.Append "单价", adDouble
    rsItems.Fields.Append "实收金额", adDouble
    rsItems.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
    rsItems.Fields.Append "开单科室", adVarChar, 100, adFldIsNullable
    rsItems.CursorLocation = adUseClient
    rsItems.LockType = adLockOptimistic
    rsItems.CursorType = adOpenStatic
    rsItems.Open
End Sub

Private Function LoadAdvicePrice(ByVal lngRow As Long, rsSend As ADODB.Recordset, cur合计 As Currency) As Boolean
'功能：读取指定医嘱(仅当前行)的计价关系到临时记录集,并计算缺省发送金额(按费别打折)
'返回：cur合计=计算出的医嘱发送金额(非药变价未算,需要输入价格后才行)
    Dim rsTmp As New ADODB.Recordset
    Dim rsCur As New ADODB.Recordset
    Dim strSQL As String, strPrice As String
    Dim str费用性质 As String, arr费用性质 As Variant
    Dim blnDo As Boolean, i As Long, k As Long
    Dim dbl数量 As Double, dbl单价 As Double, dbl应收 As Double
    Dim cur应收 As Currency, cur实收 As Currency
    Dim bln附加手术 As Boolean, lng项目ID As Long
    Dim lng主收入ID As Long, blnHaveSub As Boolean
    Dim lng执行科室ID As Long, cur金额 As Currency
    Dim lng材料ID As Long
    
    On Error GoTo errH
    
    cur金额 = 0
    With vsAdvice
        If InStr(",4,5,6,7,", rsSend!诊疗类别) > 0 Then
            '不为院外执行(自备药),药品不可能为叮嘱,且固定正常计价
            If NVL(rsSend!执行性质, 0) <> 5 Then
                mrsPrice.AddNew
                mrsPrice!医嘱ID = rsSend!ID
                mrsPrice!相关ID = rsSend!相关ID
                mrsPrice!费用性质 = 0
                mrsPrice!收费方式 = 0
                mrsPrice!收费类别 = rsSend!诊疗类别
                mrsPrice!收费细目ID = rsSend!收费细目ID
                mrsPrice!执行科室ID = rsSend!执行科室ID
                mrsPrice!数量 = 1
                mrsPrice!在用 = NVL(rsSend!跟踪在用, 0)
                mrsPrice!变价 = NVL(rsSend!是否变价, 0)
                mrsPrice!固定 = 1
                mrsPrice!从项 = 0
                                
                '发送的零售数量
                If rsSend!诊疗类别 = "7" Then
                    '中药药房单位按不可分零处理:每付
                    If NVL(rsSend!可否分零, 0) = 0 Then
                        dbl数量 = Val(.TextMatrix(lngRow, COL_总量)) * Val(.TextMatrix(lngRow, COL_单量)) / NVL(rsSend!剂量系数, 1)
                    Else
                        dbl数量 = Val(.TextMatrix(lngRow, COL_总量)) _
                            * IntEx(Val(.TextMatrix(lngRow, COL_单量)) / NVL(rsSend!剂量系数, 1) / NVL(rsSend!住院包装, 1)) * NVL(rsSend!住院包装, 1)
                    End If
                Else
                    dbl数量 = Val(.TextMatrix(lngRow, COL_总量)) * NVL(rsSend!住院包装, 1)
                End If
                dbl数量 = Format(dbl数量, "0.00000")
                                
                '记录售价单价
                If NVL(rsSend!是否变价, 0) = 0 Or rsSend!诊疗类别 = "4" And NVL(rsSend!跟踪在用, 0) = 0 Then
                    mrsPrice!单价 = Format(CalcPrice(rsSend!收费细目ID, , , True, , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                Else '以售价计算药品时价,自备药时无对应药房
                    mrsPrice!单价 = Format(CalcDrugPrice(rsSend!收费细目ID, NVL(rsSend!执行科室ID, 0), dbl数量, , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                End If
                mrsPrice.Update
                                
                '计算医嘱发送金额(按费别打折的实收金额)
                If Not IsNull(mrsPati!费别) Then
                    If NVL(rsSend!是否变价, 0) = 0 Or rsSend!诊疗类别 = "4" And NVL(rsSend!跟踪在用, 0) = 0 Then
                        cur金额 = Format(CalcPrice(rsSend!收费细目ID, mrsPati!费别, dbl数量, , NVL(rsSend!执行科室ID, 0), , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDec)
                    Else
                        cur金额 = Format(CalcDrugPrice(rsSend!收费细目ID, NVL(rsSend!执行科室ID, 0), dbl数量, mrsPati!费别, , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), "0.00000")
                    End If
                Else
                    If gbln加班加价 Then
                        '处理加班加价
                        If NVL(rsSend!是否变价, 0) = 0 Or rsSend!诊疗类别 = "4" And NVL(rsSend!跟踪在用, 0) = 0 Then
                            dbl单价 = Format(CalcPrice(rsSend!收费细目ID, , , , , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                        Else '以售价计算药品时价,自备药时无对应药房
                            dbl单价 = Format(CalcDrugPrice(rsSend!收费细目ID, NVL(rsSend!执行科室ID, 0), dbl数量, , , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                        End If
                        cur金额 = Format(mrsPrice!数量 * dbl数量 * dbl单价, gstrDec)
                    Else
                        cur金额 = Format(mrsPrice!数量 * dbl数量 * mrsPrice!单价, gstrDec)
                    End If
                End If
            End If
            
            cur合计 = cur金额
        Else
            '取诊疗收费 关系中的对照(发送时才定计价):正常计价,不为叮嘱、院外执行
            If NVL(rsSend!计价特性, 0) = 0 And InStr(",0,5,", NVL(rsSend!执行性质, 0)) = 0 Then
                dbl数量 = Format(Val(.TextMatrix(lngRow, COL_总量)), "0.00000")
                bln附加手术 = (rsSend!诊疗类别 = "F" And Not IsNull(rsSend!相关ID))
                
                '几种对应的计价情况
                If Not IsNull(rsSend!标本部位) And Not IsNull(rsSend!检查方法) Then
                    strPrice = " And 检查部位=[4] And 检查方法=[5] And Nvl(费用性质,0)=0"
                ElseIf NVL(rsSend!执行标记, 0) = 0 Then
                    strPrice = " And 检查部位 Is Null And 检查方法 is Null And Nvl(费用性质,0)=0"
                Else '目前包含床旁或术中加收的情况
                    strPrice = " And 检查部位 Is Null And 检查方法 is Null And Nvl(费用性质,0) IN(0,1)"
                End If
                                
                strPrice = "Select 收费项目ID,固有对照 From (" & _
                    " Select c.收费项目ID, c.固有对照, c.适用科室id" & _
                    "   ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
                    " From 诊疗收费关系 C Where C.诊疗项目ID=[2]" & strPrice & _
                    "       And (C.适用科室ID is Null And C.病人来源 = 0 or C.适用科室ID = Decode([3],0,[6],[3]) And C.病人来源 = " & IIF(mbytSendKind = EInBilling, 2, 1) & ")" & _
                    " ) Where Nvl(适用科室id, 0) = Top"
                    
                '先读取已有的计价
                strSQL = _
                    " Select C.类别,A.收费细目ID as 收费项目ID,A.数量 as 收费数量,Nvl(E.固有对照,0) as 固有对照," & _
                    " B.收入项目ID,C.加班加价,B.加班加价率,Decode(C.是否变价,1,A.单价,B.现价)" & IIF(bln附加手术, "*Nvl(B.附术收费率,100)/100", "") & " as 单价," & _
                    " C.是否变价,Nvl(A.从项,0) as 从项,D.跟踪在用,Nvl(A.执行科室ID,[3]) as 执行科室ID,C.屏蔽费别," & _
                    " Nvl(A.费用性质,0) as 费用性质,Nvl(A.收费方式,0) as 收费方式" & _
                    " From 病人医嘱计价 A,收费价目 B,收费项目目录 C,材料特性 D,(" & strPrice & ")  E" & _
                    " Where A.医嘱ID=[1] And A.收费细目ID=0+E.收费项目ID(+)" & _
                    " And A.收费细目ID=B.收费细目ID And A.收费细目ID=C.ID And A.收费细目ID=D.材料ID(+)" & _
                    GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "B", "7", "8", "9") & _
                    " And C.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3) And (C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                    " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                    " Order by 费用性质,从项,A.收费细目ID"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsSend!ID), Val(rsSend!诊疗项目ID), _
                    Val(NVL(rsSend!执行科室ID, 0)), CStr(NVL(rsSend!标本部位)), CStr(NVL(rsSend!检查方法)), mlng病人病区ID, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                
                '没有则读取默认的计价，只针对尚未校对的医嘱
                If rsTmp.EOF And rsSend!医嘱状态 = 1 Then
                    lng材料ID = 0 '检验试管费用,只收取试管对应的卫材费用
                    If .TextMatrix(lngRow, COL_试管编码) <> "" Then
                        lng材料ID = GetTubeMaterial(.TextMatrix(lngRow, COL_试管编码))
                    End If
                
                    '几种对应的计价情况
                    If Not IsNull(rsSend!标本部位) And Not IsNull(rsSend!检查方法) Then
                        strPrice = " And c.检查部位=[3] And c.检查方法=[4] And Nvl(c.费用性质,0)=0"
                    ElseIf NVL(rsSend!执行标记, 0) = 0 Then
                        strPrice = " And c.检查部位 Is Null And c.检查方法 is Null And Nvl(c.费用性质,0)=0"
                    Else '目前包含床旁或术中加收的情况
                        strPrice = " And c.检查部位 Is Null And c.检查方法 is Null And Nvl(c.费用性质,0) IN(0,1)"
                    End If
                    
                    strPrice = "Select * From (" & _
                        "Select C.诊疗项目ID,C.收费项目ID,C.检查部位,C.检查方法,C.费用性质,C.收费数量,C.固有对照,C.从属项目,C.收费方式,c.适用科室id" & _
                        " ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
                        " From 诊疗收费关系 C Where C.诊疗项目ID=[1]" & strPrice & _
                        "      And (C.适用科室ID is Null And C.病人来源 = 0 or C.适用科室ID = Decode([2],0,[6],[2]) And C.病人来源 = " & IIF(mbytSendKind = EInBilling, 2, 1) & ")" & _
                        " ) Where Nvl(适用科室id, 0) = Top"
                        
                    strSQL = _
                        " Select C.类别,A.收费项目ID,A.收费数量,A.固有对照,B.收入项目ID," & _
                        " C.加班加价,B.加班加价率,Decode(C.是否变价,1,B.缺省价格,B.现价)" & IIF(bln附加手术, "*Nvl(B.附术收费率,100)/100", "") & " as 单价," & _
                        " C.是否变价,Nvl(A.从属项目,0) as 从项,D.跟踪在用,[2] as 执行科室ID,C.屏蔽费别," & _
                        " Nvl(A.费用性质,0) as 费用性质,Nvl(A.收费方式,0) as 收费方式" & _
                        " From (" & strPrice & ") A,收费价目 B,收费项目目录 C,材料特性 D" & _
                        " Where A.收费项目ID=B.收费细目ID And A.收费项目ID=C.ID And A.收费项目ID=D.材料ID(+)" & _
                        GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "B", "7", "8", "9") & _
                        " And C.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3) And (C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                        " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                        " And (Nvl(A.收费方式,0)=1 And C.类别='4' And A.收费项目ID=[5] Or Not(Nvl(A.收费方式,0)=1 And C.类别='4' And [5]<>0))" & _
                        " Order by 费用性质,从项,A.收费项目ID"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsSend!诊疗项目ID), _
                        Val(NVL(rsSend!执行科室ID, 0)), CStr(NVL(rsSend!标本部位)), CStr(NVL(rsSend!检查方法)), lng材料ID, mlng病人病区ID, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                End If
                
                '确定计价之中是否包含从项以及主项收入ID
                arr费用性质 = Array()
                If Not rsTmp.EOF Then
                    Do While Not rsTmp.EOF
                        If InStr(str费用性质 & ",", "," & rsTmp!费用性质 & ",") = 0 Then
                            str费用性质 = str费用性质 & "," & rsTmp!费用性质
                        End If
                        rsTmp.MoveNext
                    Loop
                    arr费用性质 = Split(Mid(str费用性质, 2), ",")
                End If
                                
                For k = 0 To UBound(arr费用性质)
                    rsTmp.Filter = "费用性质=" & arr费用性质(k)
                    
                    lng项目ID = 0: cur金额 = 0
                    lng主收入ID = 0: blnHaveSub = False
                    If Not rsTmp.EOF And gbln从项汇总折扣 Then
                        Do While Not rsTmp.EOF
                            If NVL(rsTmp!从项, 0) = 0 Then
                                'SQL中主项排在前面,只取主项目的第一个收入
                                If lng主收入ID = 0 Then lng主收入ID = rsTmp!收入项目ID
                            ElseIf NVL(rsTmp!从项, 0) = 1 Then
                                blnHaveSub = True: Exit Do
                            End If
                            rsTmp.MoveNext
                        Loop
                        rsTmp.MoveFirst
                    End If
                    
                    Do While True
                        blnDo = False
                        If rsTmp.EOF Then
                            If lng项目ID <> 0 Then blnDo = True
                        Else
                            If rsTmp!收费项目ID <> lng项目ID And lng项目ID <> 0 Then blnDo = True
                        End If
                        If blnDo Then
                            If Not IsNull(mrsPrice!单价) Then
                                mrsPrice!单价 = Format(mrsPrice!单价, gstrDecPrice)
                            End If
                            mrsPrice.Update
                            
                            '医嘱发送金额
                            cur金额 = cur金额 + Format(cur实收, gstrDec)
                        End If
                        If rsTmp.EOF Then Exit Do
                        
                        '------------------------------------
                        If rsTmp!收费项目ID <> lng项目ID Then
                            cur实收 = 0
                            mrsPrice.AddNew
                            mrsPrice!医嘱ID = rsSend!ID
                            mrsPrice!相关ID = rsSend!相关ID
                            mrsPrice!费用性质 = NVL(rsTmp!费用性质, 0)
                            mrsPrice!收费方式 = NVL(rsTmp!收费方式, 0)
                            mrsPrice!收费类别 = rsTmp!类别
                            mrsPrice!收费细目ID = rsTmp!收费项目ID
                            mrsPrice!数量 = NVL(rsTmp!收费数量, 0)
                            mrsPrice!在用 = NVL(rsTmp!跟踪在用, 0)
                            mrsPrice!变价 = NVL(rsTmp!是否变价, 0)
                            mrsPrice!固定 = NVL(rsTmp!固有对照, 0)
                            mrsPrice!从项 = NVL(rsTmp!从项, 0)
                            
                            If rsSend!诊疗类别 = "E" And rsSend!操作类型 = "1" And rsSend!执行分类 = 5 And InStr(",5,6,", rsTmp!类别) > 0 Then
                                '原液皮试问题。绑定的药品费用如果没有指定科室则按原来逻辑
                                If Val(rsSend!用药理由 & "") <> 0 Then
                                    lng执行科室ID = Val(rsSend!用药理由 & "")
                                Else
                                    lng执行科室ID = NVL(rsTmp!执行科室ID, 0)
                                End If
                                lng执行科室ID = Get收费执行科室ID(mlng病人ID, 0, rsTmp!类别, rsTmp!收费项目ID, 4, NVL(rsSend!病人科室id, 0), 0, 2, lng执行科室ID, , , 2)
                            Else
                                '执行科室:非药嘱药品及跟踪卫材的专门取
                                lng执行科室ID = NVL(rsTmp!执行科室ID, 0)
                                If rsTmp!类别 = "4" And NVL(rsTmp!跟踪在用, 0) = 1 Or InStr(",5,6,7,", rsTmp!类别) > 0 Then
                                    lng执行科室ID = Get收费执行科室ID(mlng病人ID, 0, rsTmp!类别, rsTmp!收费项目ID, 4, NVL(rsSend!病人科室id, 0), 0, 2, lng执行科室ID, , , 2)
                                End If
                            End If
                            If lng执行科室ID <> 0 Then
                                mrsPrice!执行科室ID = lng执行科室ID
                            Else
                                mrsPrice!执行科室ID = Null
                            End If
                        End If
                        lng项目ID = rsTmp!收费项目ID
                        
                        '计算单价和实收
                        If NVL(rsTmp!是否变价, 0) = 1 And InStr(",5,6,7,", rsTmp!类别) > 0 Then
                            '非药嘱药品计价按时价计算(仅一个收入),其它变价需要由医生输入
                            mrsPrice!单价 = CalcDrugPrice(rsTmp!收费项目ID, NVL(mrsPrice!执行科室ID, 0), dbl数量 * NVL(rsTmp!收费数量, 0), , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                            
                            dbl应收 = Format(mrsPrice!数量 * dbl数量, "0.00000") * Format(mrsPrice!单价, gstrDecPrice)
                            
                            '处理加班加价
                            If gbln加班加价 And NVL(rsTmp!加班加价, 0) = 1 Then
                                dbl应收 = dbl应收 * (1 + NVL(rsTmp!加班加价率, 0) / 100)
                            End If
    
                            cur应收 = Format(dbl应收, gstrDec)
                            
                            If Not IsNull(mrsPati!费别) And Not (gbln从项汇总折扣 And blnHaveSub) And NVL(rsTmp!屏蔽费别, 0) = 0 Then
                                cur实收 = cur实收 + Format(ActualMoney(mrsPati!费别, rsTmp!收入项目ID, cur应收, rsTmp!收费项目ID, lng执行科室ID, _
                                    mrsPrice!数量 * dbl数量, IIF(gbln加班加价 And NVL(rsTmp!加班加价, 0) = 1, NVL(rsTmp!加班加价率, 0) / 100, 0)), gstrDec)
                            Else
                                cur实收 = cur实收 + cur应收
                            End If
                        ElseIf NVL(rsTmp!是否变价, 0) = 1 And rsTmp!类别 = "4" And NVL(rsTmp!跟踪在用, 0) = 1 Then
                            '跟踪在用的时价卫材和药品一样计算
                            mrsPrice!单价 = CalcDrugPrice(rsTmp!收费项目ID, NVL(mrsPrice!执行科室ID, 0), dbl数量 * NVL(rsTmp!收费数量, 0), , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                            
                            dbl应收 = Format(mrsPrice!数量 * dbl数量, "0.00000") * Format(mrsPrice!单价, gstrDecPrice)
                            
                            '处理加班加价
                            If gbln加班加价 And NVL(rsTmp!加班加价, 0) = 1 Then
                                dbl应收 = dbl应收 * (1 + NVL(rsTmp!加班加价率, 0) / 100)
                            End If
    
                            cur应收 = Format(dbl应收, gstrDec)
                            
                            If Not IsNull(mrsPati!费别) And Not (gbln从项汇总折扣 And blnHaveSub) And NVL(rsTmp!屏蔽费别, 0) = 0 Then
                                cur实收 = cur实收 + Format(ActualMoney(mrsPati!费别, rsTmp!收入项目ID, cur应收, rsTmp!收费项目ID, lng执行科室ID, _
                                    mrsPrice!数量 * dbl数量, IIF(gbln加班加价 And NVL(rsTmp!加班加价, 0) = 1, NVL(rsTmp!加班加价率, 0) / 100, 0)), gstrDec)
                            Else
                                cur实收 = cur实收 + cur应收
                            End If
                        Else '固定价格或普通变价(只有一个收入项目)
                            mrsPrice!单价 = NVL(mrsPrice!单价, 0) + NVL(rsTmp!单价, 0)
                            
                            dbl应收 = Format(mrsPrice!数量 * dbl数量, "0.00000") * Format(NVL(rsTmp!单价, 0), gstrDecPrice)
                            
                            '处理加班加价
                            If gbln加班加价 And NVL(rsTmp!加班加价, 0) = 1 Then
                                dbl应收 = dbl应收 * (1 + NVL(rsTmp!加班加价率, 0) / 100)
                            End If
                            
                            cur应收 = Format(dbl应收, gstrDec)
                            
                            If Not IsNull(mrsPati!费别) And Not (gbln从项汇总折扣 And blnHaveSub) And NVL(rsTmp!屏蔽费别, 0) = 0 Then
                                cur实收 = cur实收 + Format(ActualMoney(mrsPati!费别, rsTmp!收入项目ID, cur应收, rsTmp!收费项目ID, lng执行科室ID, _
                                    mrsPrice!数量 * dbl数量, IIF(gbln加班加价 And NVL(rsTmp!加班加价, 0) = 1, NVL(rsTmp!加班加价率, 0) / 100, 0)), gstrDec)
                            Else
                                cur实收 = cur实收 + cur应收
                            End If
                        End If
                        
                        rsTmp.MoveNext
                    Loop
                    
                    '从属项目汇总计算折扣
                    If gbln从项汇总折扣 And blnHaveSub And lng主收入ID <> 0 Then
                        cur金额 = Format(ActualMoney(NVL(mrsPati!费别), lng主收入ID, cur金额), gstrDec)
                    End If
                    
                    cur合计 = cur合计 + cur金额
                Next
            End If
        End If
    End With
    LoadAdvicePrice = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetComboList(ByVal lngRow As Long) As String
'功能：根据当前医嘱行获取可选择的计价医嘱内容
'参数：lngRow=可见行(药疗或非药)
'说明：注意这里是根据具体医嘱在取
    Dim strCombo As String
    Dim strTmp As String, lngTmp As Long
    Dim i As Long, j As Long
    
    With vsAdvice
        If .Cell(flexcpData, lngRow, COL_ID) = 3 Then
            '中药用法：中药用法,中药煎法
            lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_相关ID)
            For i = lngTmp To lngRow
                If InStr(",2,3,", CLng(.Cell(flexcpData, i, COL_ID))) > 0 Then
                    If Val(.TextMatrix(i, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                        mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                        If Not mrsPrice.EOF Then
                            For j = 1 To mrsPrice.RecordCount
                                If NVL(mrsPrice!固定, 0) = 0 Then
                                    If .Cell(flexcpData, i, COL_ID) = 2 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";中药煎法-" & .Cell(flexcpData, i, col_医嘱内容)
                                    ElseIf .Cell(flexcpData, i, COL_ID) = 3 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";中药用法-" & .Cell(flexcpData, i, col_医嘱内容)
                                    End If
                                    If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                        strCombo = strCombo & "|#" & strTmp
                                    End If
                                End If
                                mrsPrice.MoveNext
                            Next
                        Else
                            If .Cell(flexcpData, i, COL_ID) = 2 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";中药煎法-" & .Cell(flexcpData, i, col_医嘱内容)
                            ElseIf .Cell(flexcpData, i, COL_ID) = 3 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";中药用法-" & .Cell(flexcpData, i, col_医嘱内容)
                            End If
                            If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                strCombo = strCombo & "|#" & strTmp
                            End If
                        End If
                    End If
                End If
            Next
        ElseIf .Cell(flexcpData, lngRow, COL_ID) = 4 Then
            '采集方法行
            lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_相关ID)
            For i = lngTmp To lngRow
                If Val(.TextMatrix(i, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                    mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                    If Not mrsPrice.EOF Then
                        For j = 1 To mrsPrice.RecordCount
                            If NVL(mrsPrice!固定, 0) = 0 Then
                                If .TextMatrix(i, COL_诊疗类别) = "C" Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";检验项目-" & .Cell(flexcpData, i, col_医嘱内容)
                                ElseIf .TextMatrix(i, COL_诊疗类别) = "E" Then
                                    strTmp = Val(.TextMatrix(i, COL_ID)) & ";采集方法-" & .Cell(flexcpData, i, col_医嘱内容)
                                End If
                                If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                    strCombo = strCombo & "|#" & strTmp
                                End If
                            End If
                            mrsPrice.MoveNext
                        Next
                    Else
                        If .TextMatrix(i, COL_诊疗类别) = "C" Then
                            strTmp = Val(.TextMatrix(i, COL_ID)) & ";检验项目-" & .Cell(flexcpData, i, col_医嘱内容)
                        ElseIf .TextMatrix(i, COL_诊疗类别) = "E" Then
                            strTmp = Val(.TextMatrix(i, COL_ID)) & ";采集方法-" & .Cell(flexcpData, i, col_医嘱内容)
                        End If
                        If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                            strCombo = strCombo & "|#" & strTmp
                        End If
                    End If
                End If
            Next
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            '首行成药：给药途径
            If Val(.TextMatrix(lngRow - 1, COL_相关ID)) <> Val(.TextMatrix(lngRow, COL_相关ID)) Then
                lngTmp = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_相关ID))), lngRow + 1, COL_ID)
                If Val(.TextMatrix(lngTmp, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(lngTmp, COL_执行性质ID))) = 0 Then
                    mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(lngTmp, COL_ID))
                    If Not mrsPrice.EOF Then
                        For j = 1 To mrsPrice.RecordCount
                            If NVL(mrsPrice!固定, 0) = 0 Then
                                strCombo = "|#" & Val(.TextMatrix(lngTmp, COL_ID)) & ";给药途径-" & .Cell(flexcpData, lngTmp, col_医嘱内容)
                                Exit For
                            End If
                            mrsPrice.MoveNext
                        Next
                    Else
                        strCombo = "|#" & Val(.TextMatrix(lngTmp, COL_ID)) & ";给药途径-" & .Cell(flexcpData, lngTmp, col_医嘱内容)
                    End If
                End If
            End If
        Else
            '一组手术或检查，或输血医嘱，或独立医嘱
            For i = lngRow To .Rows - 1
                If i = lngRow Or Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If Val(.TextMatrix(i, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                        mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                        If Not mrsPrice.EOF Then
                            For j = 1 To mrsPrice.RecordCount
                                If NVL(mrsPrice!固定, 0) = 0 Then
                                    If .TextMatrix(i, COL_诊疗类别) = "F" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";附加手术-" & .Cell(flexcpData, i, col_医嘱内容)
                                    ElseIf .TextMatrix(i, COL_诊疗类别) = "G" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";手术麻醉-" & .Cell(flexcpData, i, col_医嘱内容)
                                    ElseIf .TextMatrix(i, COL_诊疗类别) = "D" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";检查部位-" & .TextMatrix(i, COL_标本部位) & "(" & .TextMatrix(i, COL_检查方法) & ")"
                                    ElseIf .TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(lngRow, COL_诊疗类别) = "K" Then
                                        strTmp = Val(.TextMatrix(i, COL_ID)) & ";输血途径-" & .Cell(flexcpData, i, col_医嘱内容)
                                    Else
                                        If mrsPrice!费用性质 <> 0 Then
                                            '加收费用：目前包含检查的床旁和术中加收
                                            lngTmp = -1 * Val(mrsPrice!费用性质 & Val(.TextMatrix(i, COL_ID)))
                                            strTmp = lngTmp & ";" & .Cell(flexcpData, i, COL_诊疗类别) & "医嘱-" & .Cell(flexcpData, i, col_医嘱内容) & _
                                                "(" & decode(Val(.TextMatrix(i, COL_执行标记)), 1, "床旁", 2, "术中", "") & "加收)"
                                        Else
                                            strTmp = Val(.TextMatrix(i, COL_ID)) & ";" & .Cell(flexcpData, i, COL_诊疗类别) & "医嘱-" & .Cell(flexcpData, i, col_医嘱内容)
                                        End If
                                    End If
                                    If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                        strCombo = strCombo & "|#" & strTmp
                                    End If
                                End If
                                mrsPrice.MoveNext
                            Next
                        Else
                            '未设置计价的，可能选择添加计价项目
                            If .TextMatrix(i, COL_诊疗类别) = "F" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";附加手术-" & .Cell(flexcpData, i, col_医嘱内容)
                            ElseIf .TextMatrix(i, COL_诊疗类别) = "G" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";手术麻醉-" & .Cell(flexcpData, i, col_医嘱内容)
                            ElseIf .TextMatrix(i, COL_诊疗类别) = "D" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";检查部位-" & .TextMatrix(i, COL_标本部位) & "(" & .TextMatrix(i, COL_检查方法) & ")"
                            ElseIf .TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(lngRow, COL_诊疗类别) = "K" Then
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";输血途径-" & .Cell(flexcpData, i, col_医嘱内容)
                            Else
                                strTmp = Val(.TextMatrix(i, COL_ID)) & ";" & .Cell(flexcpData, i, COL_诊疗类别) & "医嘱-" & .Cell(flexcpData, i, col_医嘱内容)
                            End If
                            If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                strCombo = strCombo & "|#" & strTmp
                            End If
                            
                            '加收费用：目前包含检查的床旁或术中加收
                            If .TextMatrix(i, COL_诊疗类别) = "D" And Val(.TextMatrix(i, COL_相关ID)) = 0 _
                                And (Val(.TextMatrix(i, COL_执行标记)) = 1 Or Val(.TextMatrix(i, COL_执行标记)) = 2) Then
                                lngTmp = -1 * Val(1 & Val(.TextMatrix(i, COL_ID)))
                                strTmp = lngTmp & ";" & .Cell(flexcpData, i, COL_诊疗类别) & "医嘱-" & .Cell(flexcpData, i, col_医嘱内容) & _
                                    "(" & decode(Val(.TextMatrix(i, COL_执行标记)), 1, "床旁", 2, "术中", "") & "加收)"
                                If InStr(strCombo & "|", "|#" & strTmp & "|") = 0 Then
                                    strCombo = strCombo & "|#" & strTmp
                                End If
                            End If
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    
    GetComboList = Mid(strCombo, 2)
End Function

Private Function ShowAdvicePrice(ByVal lngRow As Long) As Boolean
'功能：根据医嘱计价关系，计算并显示指定医嘱的费用(整个医嘱，可能多行)
'参数：lngRow=可见行(药疗或非药)
    Dim rsTmp As New ADODB.Recordset
    Dim rsExeDays As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngTopRow As Long, lngLeftCol As Long
    Dim lngPreRow As Long, lngPreCol As Long
    Dim blnFirst As Boolean, str计价医嘱 As String
    Dim str单位 As String, dbl数量 As Double
    Dim bln附加手术 As Boolean, strCombo As String, str行号 As String, str分解时间 As String
    Dim dbl单价 As Double, cur应收 As Currency, cur实收 As Currency
    Dim dbl当前单价 As Double, dbl当前应收 As Double, cur当前应收 As Currency, cur当前实收 As Currency
    Dim lng行号 As Long, cur合计 As Currency
    
    Dim rsMain As New ADODB.Recordset
    Dim rsClone As New ADODB.Recordset
    Dim strHaveSub As String, strNoneSub As String
    Dim strPriceType As String
        
    On Error GoTo errH
    
    '用于汇总计算折扣的临时记录集
    rsMain.Fields.Append "医嘱行号", adBigInt
    rsMain.Fields.Append "费用性质", adInteger
    rsMain.Fields.Append "主项行号", adBigInt
    rsMain.Fields.Append "主收入ID", adBigInt
    rsMain.Fields.Append "医嘱合计", adCurrency, , adFldIsNullable
    rsMain.CursorLocation = adUseClient
    rsMain.LockType = adLockOptimistic
    rsMain.CursorType = adOpenStatic
    rsMain.Open
    
    With vsAdvice
        blnFirst = True
        If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            If Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                blnFirst = False '一并给药中是否第一药品行
            End If
        End If
        
        If Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
            If blnFirst Then
                mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(lngRow, COL_ID)) & _
                    " Or 医嘱ID=" & Val(.TextMatrix(lngRow, COL_相关ID))
            Else
                mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(lngRow, COL_ID))
            End If
        Else
            mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(lngRow, COL_ID)) & _
                " Or 相关ID=" & Val(.TextMatrix(lngRow, COL_ID))
        End If
        
        For i = 1 To mrsPrice.RecordCount
            '计价医嘱
            bln附加手术 = False
            lng行号 = .FindRow(CStr(mrsPrice!医嘱ID), , COL_ID)
            If .TextMatrix(lng行号, COL_诊疗类别) = "4" Then
                str计价医嘱 = "卫生材料-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf InStr(",5,6,7", .TextMatrix(lng行号, COL_诊疗类别)) > 0 Then
                str计价医嘱 = "药品医嘱-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf CLng(.Cell(flexcpData, lng行号, COL_ID)) = 1 Then
                str计价医嘱 = "给药途径-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf CLng(.Cell(flexcpData, lng行号, COL_ID)) = 2 Then
                str计价医嘱 = "中药煎法-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf CLng(.Cell(flexcpData, lng行号, COL_ID)) = 3 Then
                str计价医嘱 = "中药用法-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf CLng(.Cell(flexcpData, lng行号, COL_ID)) = 4 Then
                str计价医嘱 = "采集方法-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf CLng(.Cell(flexcpData, lng行号, COL_ID)) = 5 Then
                str计价医嘱 = "输血途径-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf .TextMatrix(lng行号, COL_诊疗类别) = "C" And Val(.TextMatrix(lng行号, COL_相关ID)) <> 0 Then
                str计价医嘱 = "检验项目-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf .TextMatrix(lng行号, COL_诊疗类别) = "F" And Val(.TextMatrix(lng行号, COL_相关ID)) <> 0 Then
                bln附加手术 = True
                str计价医嘱 = "附加手术-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf .TextMatrix(lng行号, COL_诊疗类别) = "G" And Val(.TextMatrix(lng行号, COL_相关ID)) <> 0 Then
                str计价医嘱 = "手术麻醉-" & .Cell(flexcpData, lng行号, col_医嘱内容)
            ElseIf .TextMatrix(lng行号, COL_诊疗类别) = "D" And Val(.TextMatrix(lng行号, COL_相关ID)) <> 0 Then
                str计价医嘱 = "检查部位-" & .TextMatrix(lng行号, COL_标本部位) & "(" & .TextMatrix(lng行号, COL_检查方法) & ")"
            Else
                If NVL(mrsPrice!费用性质, 0) = 1 Then
                    '床旁或术中加收费用
                    str计价医嘱 = .Cell(flexcpData, lng行号, COL_诊疗类别) & "医嘱-" & .Cell(flexcpData, lng行号, col_医嘱内容) & _
                        "(" & decode(Val(.TextMatrix(lng行号, COL_执行标记)), 1, "床旁", 2, "术中", "") & "加收)"
                Else
                    str计价医嘱 = .Cell(flexcpData, lng行号, COL_诊疗类别) & "医嘱-" & .Cell(flexcpData, lng行号, col_医嘱内容)
                End If
            End If
            str计价医嘱 = Replace(str计价医嘱, "'", "''")
            
            '数量:药品按住院单位的数量,其它按零售数量
            If InStr(",5,6,", .TextMatrix(lng行号, COL_诊疗类别)) > 0 Then
                dbl数量 = Val(.TextMatrix(lng行号, COL_总量))
            ElseIf .TextMatrix(lng行号, COL_诊疗类别) = "7" Then
                '中药药房单位按不可分零处理:每付
                If Val(.TextMatrix(lng行号, COL_可否分零)) = 0 Then
                    dbl数量 = Val(.TextMatrix(lng行号, COL_总量)) * Val(.TextMatrix(lng行号, COL_单量)) _
                        / Val(.TextMatrix(lng行号, COL_剂量系数)) / Val(.TextMatrix(lng行号, COL_住院包装))
                Else
                    dbl数量 = Val(.TextMatrix(lng行号, COL_总量)) _
                        * IntEx(Val(.TextMatrix(lng行号, COL_单量)) / Val(.TextMatrix(lng行号, COL_剂量系数)) / Val(.TextMatrix(lng行号, COL_住院包装)))
                End If
            Else
                If InStr(",3,4,5,6,", Val("" & mrsPrice!收费方式)) > 0 Then '一天只收一次的
                     '分解时间
                    If .TextMatrix(lng行号, COL_分解时间) <> "" Then
                        str分解时间 = .TextMatrix(lng行号, COL_分解时间)
                    Else
                        str分解时间 = .Cell(flexcpData, lng行号, COL_分解时间)    '开始执行时间
                    End If
                    
                    Set rsExeDays = GetExecDays(str分解时间)
                    dbl数量 = rsExeDays.RecordCount
                ElseIf InStr(",1,2,", Val("" & mrsPrice!收费方式)) > 0 Then '一次发送只收一次
                    dbl数量 = 1
                Else
                    dbl数量 = Val(.TextMatrix(lng行号, COL_总量))
                End If
            End If
            dbl数量 = Format(dbl数量 * NVL(mrsPrice!数量, 0), "0.00000")
                        
            '组合SQL
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                " Select " & i & " as 序号," & mrsPrice!医嘱ID & " as 医嘱ID,ID," & _
                NVL(mrsPrice!固定, 0) & " as 固定,'" & str计价医嘱 & "' as 计价医嘱,类别,名称,产地,规格," & _
                "计算单位 as 单位," & NVL(mrsPrice!数量, 0) & " as 计价数量," & dbl数量 & " as 数量," & _
                Format(NVL(mrsPrice!单价, 0), gstrDecPrice) & " as 单价,费用类型," & lng行号 & " as 行号," & _
                " 是否变价,加班加价," & IIF(bln附加手术, 1, 0) & " as 附加手术," & mrsPrice!从项 & " as 从项," & _
                NVL(mrsPrice!执行科室ID, 0) & " as 执行科室ID,屏蔽费别," & mrsPrice!费用性质 & " as 费用性质," & _
                mrsPrice!收费方式 & " as 收费方式 From 收费项目目录 Where ID=" & mrsPrice!收费细目ID
            mrsPrice.MoveNext
        Next
    End With
    
    With vsPrice
        lngPreRow = .Row: lngPreCol = .Col
        lngTopRow = .TopRow: lngLeftCol = .LeftCol
        .Editable = flexEDNone
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        '需要计价的医嘱选择
        '根据待发送医嘱取可计价医嘱(不能从mrsPrice取,因为可能无收费关系或已删除,而且也许现在计价已全部删除)
        strCombo = GetComboList(lngRow)
        If strCombo <> "" Then
            .ColData(COLP_计价医嘱) = strCombo
            .Editable = flexEDKbdMouse '可以选择则可以编辑
        Else
            .ColData(COLP_计价医嘱) = ""
        End If
        
        '显示已有的计价项目
        If strSQL <> "" Then
            strSQL = "Select A.行号,A.ID AS 收费细目ID,A.固定,A.从项,A.计价医嘱,A.类别,C.名称 as 类别名称,A.执行科室ID,G.名称 as 执行科室," & _
                " Nvl(E.名称,A.名称)||Decode(A.产地,NULL,NULL,'('||A.产地||')')||Decode(A.规格,NULL,NULL,' '||A.规格) as 名称," & _
                " A.单位,A.计价数量,A.数量,D.住院包装,D.住院单位,Decode(A.是否变价,1,A.单价,B.现价) as 单价,F.跟踪在用," & _
                " A.费用性质,A.收费方式,A.费用类型,A.屏蔽费别,A.是否变价,A.加班加价,B.加班加价率,B.原价,B.现价,A.附加手术,B.附术收费率,B.收入项目ID" & _
                " From (" & strSQL & ") A,收费价目 B,收费项目类别 C,药品规格 D,收费项目别名 E,材料特性 F,部门表 G" & _
                " Where A.ID=B.收费细目ID And A.类别=C.编码 And A.ID=D.药品ID(+)" & _
                " And A.ID=F.材料ID(+) And A.执行科室ID=G.ID(+)" & _
                GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "1", "2", "3") & _
                " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                " And A.ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIF(gbyt药品名称显示 = 0, 1, 3) & _
                " Order by A.序号"
                '因为输入后是调用本函数刷新,要保持动态记录集中记录顺序
                '要保证主项排在前面,LoadAdvicePrice时，主项是排在前面，而且编辑后只可能加了从项
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级) '没法
            
            If Not rsTmp.EOF And gbln从项汇总折扣 Then
                Set rsClone = rsTmp.Clone
            End If
            
            For i = 1 To rsTmp.RecordCount
                If str行号 <> rsTmp!行号 & "_" & rsTmp!费用性质 & "_" & rsTmp!收费细目ID Then
                    If str行号 <> "" Then
                        If Not (Val(.TextMatrix(.Rows - 1, COLP_变价)) = 1 And dbl单价 = 0) Then
                            .TextMatrix(.Rows - 1, COLP_单价) = Format(dbl单价, gstrDecPrice)
                            .Cell(flexcpData, .Rows - 1, COLP_单价) = .TextMatrix(.Rows - 1, COLP_单价) '记录用于恢复输入
                            .TextMatrix(.Rows - 1, COLP_应收金额) = Format(cur应收, gstrDec)
                            .TextMatrix(.Rows - 1, COLP_实收金额) = Format(cur实收, gstrDec)
                        End If
                        cur合计 = cur合计 + Format(cur实收, gstrDec)
                    End If
                    str行号 = rsTmp!行号 & "_" & rsTmp!费用性质 & "_" & rsTmp!收费细目ID
                    dbl单价 = 0: cur应收 = 0: cur实收 = 0
                    .Rows = .Rows + 1
                    
                    '标识固定对照为灰色
                    If rsTmp!固定 <> 0 Then
                        .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HE0E0E0
                    End If

                    .TextMatrix(.Rows - 1, COLP_行号) = rsTmp!行号
                    .TextMatrix(.Rows - 1, COLP_收费细目ID) = rsTmp!收费细目ID
                    .TextMatrix(.Rows - 1, COLP_固定) = rsTmp!固定
                    .TextMatrix(.Rows - 1, COLP_计价医嘱) = rsTmp!计价医嘱
                    .TextMatrix(.Rows - 1, COLP_费用性质) = rsTmp!费用性质
                    .TextMatrix(.Rows - 1, COLP_收费方式) = getChargeMode(Val(NVL(rsTmp!收费方式, 0)))
                        .Cell(flexcpData, .Rows - 1, COLP_收费方式) = Val(NVL(rsTmp!收费方式, 0))
                    .TextMatrix(.Rows - 1, COLP_类别) = rsTmp!类别名称
                    .TextMatrix(.Rows - 1, COLP_收费类别) = rsTmp!类别
                    .TextMatrix(.Rows - 1, COLP_收费项目) = rsTmp!名称
                    .TextMatrix(.Rows - 1, COLP_计价数量) = NVL(rsTmp!计价数量, 0) '相对数量
                    
                    dbl数量 = NVL(rsTmp!数量, 0) '售价数量用于后面按成本打折计算
                    If InStr(",5,6,7,", rsTmp!类别) > 0 Then '住院包装
                        .TextMatrix(.Rows - 1, COLP_单位) = NVL(rsTmp!住院单位)
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!行号, COL_诊疗类别)) > 0 Then
                            .TextMatrix(.Rows - 1, COLP_数量) = FormatEx(NVL(rsTmp!数量, 0), 5)
                            dbl数量 = dbl数量 * NVL(rsTmp!住院包装, 1)
                        Else
                            '中药药房单位按不可分零处理:每付
                            '非药嘱药品计价:因为这里预定了售价数量,因此转换为药房单位显示时不作不分零处理
                            .TextMatrix(.Rows - 1, COLP_数量) = FormatEx(NVL(rsTmp!数量, 0) / NVL(rsTmp!住院包装, 1), 5)
                        End If
                    Else
                        .TextMatrix(.Rows - 1, COLP_单位) = NVL(rsTmp!单位)
                        .TextMatrix(.Rows - 1, COLP_数量) = FormatEx(NVL(rsTmp!数量, 0), 5)
                    End If
                    
                    .TextMatrix(.Rows - 1, COLP_执行科室) = NVL(rsTmp!执行科室)
                    .TextMatrix(.Rows - 1, COLP_执行科室ID) = NVL(rsTmp!执行科室ID, 0)
                    
                    '显示医保费用类型，医生站发送只能发送临嘱
                    If Val(rsTmp!收费细目ID) <> 0 Then
                        strPriceType = GetPriceType(Val(mlng病人ID), Val(rsTmp!收费细目ID & ""), Val(mint险类), mlng病人性质 = 1)
                    End If
                    '费用类型
                    If strPriceType = "" Then
                        .TextMatrix(.Rows - 1, COLP_费用类型) = NVL(rsTmp!费用类型)
                    Else
                        .TextMatrix(.Rows - 1, COLP_费用类型) = strPriceType
                    End If
                    
                    .TextMatrix(.Rows - 1, COLP_从项) = IIF(NVL(rsTmp!从项, 0) = 0, "", "√")
                    .TextMatrix(.Rows - 1, COLP_跟踪在用) = NVL(rsTmp!跟踪在用, 0)
                    
                    '记录用于输入恢复
                    .Cell(flexcpData, .Rows - 1, COLP_计价医嘱) = .TextMatrix(.Rows - 1, COLP_计价医嘱)
                    .Cell(flexcpData, .Rows - 1, COLP_收费项目) = .TextMatrix(.Rows - 1, COLP_收费项目)
                    .Cell(flexcpData, .Rows - 1, COLP_计价数量) = .TextMatrix(.Rows - 1, COLP_计价数量)
                    .Cell(flexcpData, .Rows - 1, COLP_执行科室) = .TextMatrix(.Rows - 1, COLP_执行科室)
                    
                    '记录从属主项的信息，以便计算
                    If gbln从项汇总折扣 And rsTmp!从项 = 0 Then
                        If InStr(strHaveSub & ",", "," & rsTmp!行号 & "_" & rsTmp!费用性质 & ",") = 0 _
                            And InStr(strNoneSub & ",", "," & rsTmp!行号 & "_" & rsTmp!费用性质 & ",") = 0 Then
                            rsClone.Filter = "行号=" & rsTmp!行号 & " And 费用性质=" & rsTmp!费用性质 & " And 从项=1"
                            If Not rsClone.EOF Then
                                rsMain.AddNew
                                rsMain!医嘱行号 = rsTmp!行号
                                rsMain!费用性质 = rsTmp!费用性质
                                rsMain!主项行号 = .Rows - 1
                                rsMain!主收入ID = rsTmp!收入项目ID
                                rsMain.Update
                                strHaveSub = strHaveSub & "," & rsTmp!行号 & "_" & rsTmp!费用性质
                            Else
                                strNoneSub = strNoneSub & "," & rsTmp!行号 & "_" & rsTmp!费用性质
                            End If
                        End If
                    End If
                    
                    '非药品、卫材医嘱的药品和跟踪卫材计价：即使固定也可以修改执行科室
                    If InStr(",5,6,7,", rsTmp!类别) > 0 _
                        Or rsTmp!类别 = "4" And NVL(rsTmp!跟踪在用, 0) = 1 Then
                        .Editable = flexEDKbdMouse
                    End If
                End If
                
                '单价计算处理
                If InStr(",5,6,7,", rsTmp!类别) > 0 Then
                    If NVL(rsTmp!是否变价, 0) = 0 Then
                        dbl当前单价 = NVL(rsTmp!单价, 0)
                    Else
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!行号, COL_诊疗类别)) > 0 Then
                            dbl当前单价 = CalcDrugPrice(rsTmp!收费细目ID, NVL(rsTmp!执行科室ID, 0), Format(NVL(rsTmp!数量, 0) * NVL(rsTmp!住院包装, 1), "0.00000"), , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                        Else
                            dbl当前单价 = CalcDrugPrice(rsTmp!收费细目ID, NVL(rsTmp!执行科室ID, 0), Format(NVL(rsTmp!数量, 0), "0.00000"), , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                        End If
                    End If
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!行号, COL_诊疗类别)) > 0 Then
                        dbl当前单价 = dbl当前单价 * NVL(rsTmp!住院包装, 1)
                        dbl当前应收 = Format(NVL(rsTmp!数量, 0), "0.00000") * dbl当前单价
                    Else
                        dbl当前应收 = Format(NVL(rsTmp!数量, 0), "0.00000") * dbl当前单价
                        dbl当前单价 = dbl当前单价 * NVL(rsTmp!住院包装, 1)
                    End If
                ElseIf rsTmp!类别 = "4" And NVL(rsTmp!跟踪在用, 0) = 1 And NVL(rsTmp!是否变价, 0) = 1 Then
                    '跟踪在用的时价卫材和药品一样计算
                    dbl当前单价 = CalcDrugPrice(rsTmp!收费细目ID, NVL(rsTmp!执行科室ID, 0), Format(NVL(rsTmp!数量, 0), "0.00000"), , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                    dbl当前应收 = Format(NVL(rsTmp!数量, 0), "0.00000") * dbl当前单价
                Else
                    dbl当前单价 = NVL(rsTmp!单价, 0) '其它如果为变价则是用户输入的
                    dbl当前应收 = Format(NVL(rsTmp!数量, 0), "0.00000") * dbl当前单价
                    If NVL(rsTmp!是否变价, 0) = 1 Then '记录非药变价范围
                        .TextMatrix(.Rows - 1, COLP_变价) = 1
                        .Cell(flexcpData, .Rows - 1, COLP_应收金额) = CCur(NVL(rsTmp!原价, 0))
                        .Cell(flexcpData, .Rows - 1, COLP_实收金额) = CCur(NVL(rsTmp!现价, 0))
                        .Editable = flexEDKbdMouse '非药品变价,即使固定也可以定价
                    End If
                End If
                '应收
                If rsTmp!附加手术 = 1 Then
                    dbl当前应收 = dbl当前应收 * NVL(rsTmp!附术收费率, 100) / 100
                End If
                '处理加班加价
                If gbln加班加价 And NVL(rsTmp!加班加价, 0) = 1 Then
                    dbl当前应收 = dbl当前应收 * (1 + NVL(rsTmp!加班加价率, 0) / 100)
                End If
                cur当前应收 = Format(dbl当前应收, gstrDec)
                
                '实收
                If gbln从项汇总折扣 And (rsTmp!从项 = 1 Or InStr(strHaveSub & ",", "," & rsTmp!行号 & "_" & rsTmp!费用性质 & ",") > 0) Then
                    cur当前实收 = Format(cur当前应收, gstrDec)
                    '累计医嘱合计来计算折扣
                    rsMain.Filter = "医嘱行号=" & rsTmp!行号 & " And 费用性质=" & rsTmp!费用性质
                    rsMain!医嘱合计 = NVL(rsMain!医嘱合计, 0) + cur当前实收
                    rsMain.Update
                ElseIf NVL(rsTmp!屏蔽费别, 0) = 0 And Not IsNull(mrsPati!费别) Then
                    cur当前实收 = Format(ActualMoney(mrsPati!费别, rsTmp!收入项目ID, cur当前应收, rsTmp!收费细目ID, NVL(rsTmp!执行科室ID, 0), _
                        dbl数量, IIF(gbln加班加价 And NVL(rsTmp!加班加价, 0) = 1, NVL(rsTmp!加班加价率, 0) / 100, 0)), gstrDec)
                Else
                    cur当前实收 = Format(cur当前应收, gstrDec)
                End If
                
                dbl单价 = dbl单价 + dbl当前单价
                cur应收 = cur应收 + cur当前应收
                cur实收 = cur实收 + cur当前实收
                
                rsTmp.MoveNext
            Next
            If str行号 <> "" Then
                If Not (Val(.TextMatrix(.Rows - 1, COLP_变价)) = 1 And dbl单价 = 0) Then
                    .TextMatrix(.Rows - 1, COLP_单价) = Format(dbl单价, gstrDecPrice)
                    .Cell(flexcpData, .Rows - 1, COLP_单价) = .TextMatrix(.Rows - 1, COLP_单价) '记录用于恢复输入
                    .TextMatrix(.Rows - 1, COLP_应收金额) = Format(cur应收, gstrDec)
                    .TextMatrix(.Rows - 1, COLP_实收金额) = Format(cur实收, gstrDec)
                End If
                cur合计 = cur合计 + Format(cur实收, gstrDec)
            End If
        End If
        
        '汇总计算折扣
        If gbln从项汇总折扣 And strHaveSub <> "" Then
            rsMain.Filter = 0
            Do While Not rsMain.EOF
                cur当前实收 = Format(ActualMoney(NVL(mrsPati!费别), rsMain!主收入ID, rsMain!医嘱合计), gstrDec)
                cur合计 = cur合计 - Val(.TextMatrix(rsMain!主项行号, COLP_实收金额))
                .TextMatrix(rsMain!主项行号, COLP_实收金额) = Format(Val(.TextMatrix(rsMain!主项行号, COLP_实收金额)) + (cur当前实收 - rsMain!医嘱合计), gstrDec)
                cur合计 = cur合计 + Val(.TextMatrix(rsMain!主项行号, COLP_实收金额))
                rsMain.MoveNext
            Loop
        End If
        
        '------------------------------------------------
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        '定位缺省单元
        If lngPreRow >= .FixedRows And lngPreRow <= .Rows - 1 Then
            .Row = lngPreRow
        Else
            .Row = .FixedRows
        End If
        If lngPreCol >= COLP_计价医嘱 And lngPreCol <= .Cols - 1 Then
            .Col = lngPreCol
        Else
            .Col = COLP_计价医嘱
        End If
        '定位表格输入位置
        If lngTopRow >= .FixedRows And lngTopRow <= .Rows - 1 Then
            .TopRow = lngTopRow
        End If
        If lngLeftCol >= COLP_计价医嘱 And lngLeftCol <= .Cols - 1 Then
            .LeftCol = lngLeftCol
        End If
        .Redraw = flexRDDirect
    End With
    
    '重新汇总显示可见行的发送医嘱金额
    vsAdvice.TextMatrix(lngRow, COL_金额) = Format(cur合计, gstrDec)
    ShowAdvicePrice = True
    
    Call ShowSendTotal
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CellEditable(ByVal lngRow As Long, ByVal lngCol As Long, Optional bln非本科 As Boolean) As Boolean
'功能：判断价表中单元格是否可以编辑
    Dim lng行号 As Long
    
    With vsPrice
        bln非本科 = False
        CellEditable = .Editable
        lng行号 = Val(.TextMatrix(lngRow, COLP_行号))
        If lngCol = COLP_执行科室 Then
            '跟踪在用的卫材,非药嘱药品计价的执行科室可以修改
            If Not ((.TextMatrix(lngRow, COLP_收费类别) = "4" And Val(.TextMatrix(lngRow, COLP_跟踪在用)) = 1 _
                Or InStr(",5,6,7,", .TextMatrix(lngRow, COLP_收费类别)) > 0) And InStr(",4,5,6,7,", vsAdvice.TextMatrix(lng行号, COL_诊疗类别)) = 0) Then
                CellEditable = False
            End If
            If .TextMatrix(lngRow, COLP_收费项目) = "" Or .TextMatrix(lngRow, COLP_行号) = "" Then
                CellEditable = False
            End If
        ElseIf Val(.TextMatrix(lngRow, COLP_固定)) <> 0 Then
            '固定对照行仅可以修改变价
            If Not (Val(.TextMatrix(lngRow, COLP_变价)) = 1 And lngCol = COLP_单价) Then
                CellEditable = False
            End If
        Else
            If lngCol = COLP_单价 Then
                If Val(.TextMatrix(lngRow, COLP_变价)) <> 1 Then
                    CellEditable = False
                Else
                    '非本科执行的变价项目不允许定价格
                    If lng行号 <> 0 Then
                        If Not Check本科执行(Val(vsAdvice.TextMatrix(lng行号, COL_执行科室ID))) Then
                            bln非本科 = True: CellEditable = False
                        End If
                    End If
                End If
            ElseIf lngCol <> COLP_计价医嘱 And lngCol <> COLP_计价数量 And lngCol <> COLP_收费项目 Then
                CellEditable = False
            End If
        End If
    End With
End Function

Private Sub Refresh领药号()
    Dim objCbo As CommandBarComboBox
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strPre As String
    
    On Error GoTo errH
    
    Set objCbo = cbsMain.FindControl(, conMenu_View_Find)
    
    If objCbo.ListIndex > 0 Then strPre = objCbo.List(objCbo.ListIndex)
    
    objCbo.Clear
    objCbo.AddItem "<使用新的领药号>"
    objCbo.ListIndex = 1
    
    strSQL = "Select Distinct 领药号 From 未发药品记录 Where 填制日期>=Trunc(Sysdate) And 单据=9 And 对方部门ID=[1] And 领药号 is Not NULL Order by 领药号 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng领药部门ID)
    Do While Not rsTmp.EOF
        objCbo.AddItem rsTmp!领药号
        If rsTmp!领药号 = strPre Then
            objCbo.ListIndex = objCbo.ListCount
        End If
        rsTmp.MoveNext
    Loop

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Get领药号() As String
    Dim objCbo As CommandBarComboBox
    
    Set objCbo = cbsMain.FindControl(, conMenu_View_Find)
    If objCbo.ListIndex = 1 Then
        Get领药号 = zlDatabase.GetNextNo(122, mlng领药部门ID)
    ElseIf objCbo.ListIndex > 1 Then
        Get领药号 = objCbo.List(objCbo.ListIndex)
    End If
End Function

Private Function LoadAdviceSend(ByVal str类别s As String, ByVal str药品 As String, Optional ByVal lngModle As Long) As Boolean
'功能：根据条件读取并显示要发送的药品医嘱清单
'参数：str类别s,str药品：对应发送选项中的医嘱发送类别及药品发送类别,lngModle=0 读取常规医嘱，=1读取备用医嘱
'说明：注意CellData中存放得有附加数据
'   RowData：0-未发送的,-1-已成功发送的
'   COL_选择：0-可自由选择的,1-禁止改变选择状态的
'   COL_ID：1-给药途径，2-中药煎法，3-中药用法，4-采集方法，5-输血途径
'   COL_婴儿：存放婴儿编号
'   COL_诊疗类别：存放诊疗类别名称，用于显示计价医嘱
'   COL_医嘱内容：存放诊疗项目名称或标本部位，用于显示计价医嘱
'   COL_分解时间：存放费用的发生时间(无分解时间时)
'   COL_频率：1-"一次性"临嘱
'   COL_金额：原始的金额而不是累计显示的
    Dim rsSend As New ADODB.Recordset
    Dim strSQL As String, lngTmp As Long, strTmp As String
    Dim lngRow As Long, lngDel组ID As Long
    Dim bln分批时价 As Boolean, lng次数 As Long, lng最小次数 As Long
    Dim str分解时间 As String, dbl总量 As Double, cur金额 As Currency
    
    Dim vMsg As VbMsgBoxResult
    Dim bln药品时价提示 As Boolean, bln药品库存提示 As Boolean, bln药品默认发送 As Boolean
    Dim bln卫材时价提示 As Boolean, bln卫材库存提示 As Boolean, bln卫材默认发送 As Boolean
    Dim str用法 As String, i As Long, j As Long
    Dim str停用 As String
    Dim str备用医嘱 As String
    Dim strNoneIDs As String
    Dim str输液药品排除 As String '是否可以发送输液药品
    Dim bln药品零差价提示 As Boolean
    Dim str特殊医嘱排除 As String
    
    Screen.MousePointer = 11
    mlngRefModld = lngModle
    stbThis.Panels(3).Text = "": stbThis.Panels(4).Text = "": Call Form_Resize
    
    vsPrice.Rows = vsPrice.FixedRows
    vsPrice.Rows = vsPrice.FixedRows + 1
    vsAdvice.Rows = vsAdvice.FixedRows '有删除行功能
    
    vsAdvice.ColHidden(COL_婴儿) = True
    Me.Refresh
    
    Call InitPriceRecordset '计价关系表
 
    mstrAdDrugIDs = ""
    If mstr前提IDs = "" Then
        strNoneIDs = GetNoneSendID(mlng病人ID, mlng主页ID, 2, , , mstrAdDrugIDs)
    End If
    '获取发送清单:新开和已校的每条医嘱记录(药品和非药品),开嘱医生为自已
    '----------------------------------------------------------------------------------------------------------
    '叮嘱、护理等级、术前术后医嘱不发送,但这里先读取叮嘱(给药途径,用法,煎法,采集方法,输血途径)
    '备用医嘱
    If lngModle = 1 Then
        str备用医嘱 = " And NVL(a.执行频次,'无')='需要时' And to_date([5],'yyyy-mm-dd hh24:mi') - a.开始执行时间<0.5 "
    Else
        str备用医嘱 = " And NVL(a.执行频次,'无')<>'需要时' "
    End If

    '排除输液药品,如果一组药品中有一个不是发送要输液配置中心的，就可以在这里发送
    If lngModle = 0 And mbln发送到配置中心 Then
        str特殊医嘱排除 = Get输液类医嘱(mlng病人ID, mlng主页ID, 0)
        str输液药品排除 = " and instr(','||[8]|| ',',','||Nvl(A.相关ID,A.ID)||',')=0"
    Else
        If gstr输液配置中心 <> "" And mbln发送到配置中心 = False Then
            '如果科室启用了配置中心，则排除所有的静配流程的药品，如果科室没有启用，但全院启用了，则只排除静脉营养医嘱
            str输液药品排除 = " And NVL(B.执行标记,0)<>2 And (Not Exists(Select 1 From 诊疗项目目录 Y Where X.诊疗项目id = y.Id And NVL(Y.执行标记,0)=2) OR x.诊疗项目id is null)"
        End If
    End If
    strSQL = _
        " Select A.ID,A.相关ID,Nvl(A.相关ID,A.ID) as 组ID,Nvl(X.序号,A.序号) as 组号,A.医嘱状态," & _
        " A.诊疗类别,F.名称 as 类别名称,A.诊疗项目ID,B.名称 as 诊疗项目,A.收费细目ID,A.婴儿," & _
        " A.医嘱内容,A.标本部位,A.检查方法,A.执行标记,A.天数,A.总给予量,D.住院单位,A.单次用量," & _
        " Decode(A.诊疗类别,'4',C.计算单位,B.计算单位) as 计算单位,D.剂量系数,D.住院包装," & _
        " A.开始执行时间,A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,A.医生嘱托,A.执行时间方案,a.申请序号,b.执行分类," & _
        " A.病人科室ID,A.开嘱科室ID,A.开嘱医生,A.开嘱时间,A.计价特性,A.执行性质,A.执行科室ID,Nvl(E.名称,Decode(Nvl(A.执行性质,0),5,'-')) as 执行科室," & _
        " B.操作类型,B.试管编码,Nvl(A.可否分零,D.住院可否分零) as 可否分零,Decode(A.诊疗类别,'4',G.在用分批,D.药房分批) as 分批," & _
        " G.跟踪在用,C.是否变价,C.撤档时间,C.服务对象,A.新开签名ID as 签名ID,A.摘要,a.紧急标志,c.撤档时间,B.计算方式,b.执行安排,h.毒理分类,a.用药理由" & _
        " From 病人医嘱记录 A,诊疗项目目录 B,收费项目目录 C,药品规格 D,病案主页 H,部门表 E,诊疗项目类别 F,材料特性 G,药品特性 H,病人医嘱记录 X" & _
        " Where A.病人ID=[1] And A.主页ID=[2] And h.病人ID=A.病人ID And h.主页ID = a.主页ID And Nvl(A.前提ID,0) in (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist)) X)" & _
        " And A.医嘱状态 IN(1,3,5) And A.医嘱期效=1 And A.相关ID=X.ID(+) And B.类别=F.编码" & _
        " And A.诊疗项目ID=B.ID And A.收费细目ID=C.ID(+) And A.收费细目ID=D.药品ID(+) And A.收费细目ID=G.材料ID(+)" & _
        " And A.执行科室ID=E.ID(+) And Nvl(A.执行标记,0)<>-1 And A.病人来源<>3" & _
        " And A.开始执行时间 is Not NULL And Nvl(A.医嘱状态,0)<>-1 and b.id=h.药名id(+)" & _
        IIF(gblnKSSStrict Or gbln手术分级管理 Or gbln输血分级管理 Or gbln血库系统, " And Nvl(A.审核状态,0) Not in " & IIF(gbln血库系统 = True, " (1,3,7)", " (1,3,4,5,7)"), "") & _
        IIF(strNoneIDs <> "" And Not mbln阳性用药, " And Instr([6],','||A.ID||',')=0", "") & _
        IIF(InStr(UserInfo.性质, "护士") > 0, "", " And Decode(A.审核标记,1,Substr(A.开嘱医生,1,Instr(A.开嘱医生,'/')-1),Substr(A.开嘱医生,Instr(A.开嘱医生,'/')+1))=[4]") & _
        " And Exists(Select M.姓名 From 人员表 M,执业类别 N" & _
        " Where M.姓名=Decode(A.审核标记,1,Substr(A.开嘱医生,1,Instr(A.开嘱医生,'/')-1),Substr(A.开嘱医生,Instr(A.开嘱医生,'/')+1))" & _
        " And M.执业类别=N.编码 And N.分类 IN('执业医师','执业助理医师'))" & _
        " And Not(A.诊疗类别='H' And B.操作类型='1' And B.执行频率=2) And Not(A.诊疗类别='Z' And B.操作类型 In('4','14'))" & _
         str备用医嘱 & str输液药品排除 & " And (h.婴儿科室ID is null or h.婴儿科室ID is not null and (h.婴儿病区ID=[7] or h.婴儿科室ID=[7]) and NVL(A.婴儿,0)<>0 or h.婴儿科室ID is not null and (h.婴儿病区ID<>[7] and h.婴儿科室ID<>[7]) and NVL(A.婴儿,0)=0) " & _
        decode(mint医嘱处理范围, 1, " And nvl(a.婴儿,0) = 0 ", 2, " And nvl(a.婴儿,0) <> 0 ", "") & " Order by A.婴儿,组号,组ID,A.序号"

    On Error GoTo errH
    Set rsSend = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, IIF(mstr前提IDs = "", "0", mstr前提IDs), UserInfo.姓名, Format(dkpExecTime.value, "YYYY-MM-DD HH:mm"), _
        "," & strNoneIDs & ",", mlng医护科室ID, gstr输液配置中心, str特殊医嘱排除)
    
    '计算并显示发送清单
    '----------------------------------------------------------------------------------------------------------
    If Not rsSend.EOF Then
        With vsAdvice
            bln药品时价提示 = True: bln药品库存提示 = True: bln药品默认发送 = True
            bln卫材时价提示 = True: bln卫材库存提示 = True: bln卫材默认发送 = True
            bln药品零差价提示 = True
            If str药品 = "" Then str药品 = "111" '药品类别：院内用药，离院带药，自取药
            .Redraw = flexRDNone
            For i = 1 To rsSend.RecordCount
                '一组医嘱中的一个不能发送,则整组不能发送
                If lngDel组ID <> 0 Then
                    If NVL(rsSend!相关ID, rsSend!ID) = lngDel组ID Then
                        GoTo NextLoop
                    Else
                        lngDel组ID = 0
                    End If
                End If
                
                
                '检查不允许发送的诊疗类别
                '一组医嘱检查第一个医嘱行,其它相关行不检查
                If str类别s <> "" And lngDel组ID = 0 Then
                    If rsSend!诊疗类别 = "7" Then
                        '中药配方
                        If InStr(str类别s, "'8'") = 0 Then lngDel组ID = rsSend!相关ID
                    ElseIf InStr(",5,6,", rsSend!诊疗类别) > 0 Then
                        '中西成药(包括西成药，中成药混合一并给药的情况)
                        If InStr(str类别s, "'" & rsSend!诊疗类别 & "'") = 0 Then
                            lngDel组ID = rsSend!相关ID: lng最小次数 = 0
                        End If
                    ElseIf rsSend!诊疗类别 = "D" Then
                        '检查组合(或独立的检查)
                        If InStr(str类别s, "'D'") = 0 Then lngDel组ID = rsSend!ID
                    ElseIf rsSend!诊疗类别 = "F" Then
                        '手术组合(或独立的手术)
                        If InStr(str类别s, "'F'") = 0 Then lngDel组ID = rsSend!ID
                    ElseIf rsSend!诊疗类别 = "C" Then
                        '检验组合(或独立的检验)
                        If InStr(str类别s, "'C'") = 0 Then lngDel组ID = NVL(rsSend!相关ID, rsSend!ID)
                    ElseIf rsSend!诊疗类别 = "K" Then
                        '输血项目和途径(或独立的输血)
                        If InStr(str类别s, "'K'") = 0 Then lngDel组ID = rsSend!ID
                    ElseIf IsNull(rsSend!相关ID) And rsSend!ID <> Val(.TextMatrix(.Rows - 1, COL_相关ID)) Then
                        '其它独立项目
                        If InStr(str类别s, "'" & rsSend!诊疗类别 & "'") = 0 Then lngDel组ID = rsSend!ID
                    End If
                    If lngDel组ID <> 0 Then GoTo NextLoop
                End If
                                                
                '加入当前行
                .Rows = .Rows + 1: lngRow = .Rows - 1
                .Cell(flexcpPictureAlignment, lngRow, COL_选择) = 4
                Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("T").Picture
                
                '如果是停用的，则提示不能发送
                If Format(NVL(rsSend!撤档时间, "3000-1-1"), "YYYY-MM-DD") <> Format("3000-1-1", "YYYY-MM-DD") Then
                    .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                    Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                    If InStr(str停用 & ",", "," & rsSend!医嘱内容 & ",") = 0 Then str停用 = str停用 & "," & rsSend!医嘱内容
                End If
                
                '隐藏相关行
                If rsSend!诊疗类别 = "7" Then
                    .RowHidden(lngRow) = True '中草药
                ElseIf rsSend!诊疗类别 = "E" Then
                    If Not IsNull(rsSend!相关ID) Then
                        .RowHidden(lngRow) = True
                        If .TextMatrix(lngRow - 1, COL_诊疗类别) = "K" Then
                            .Cell(flexcpData, lngRow, COL_ID) = 5 '输血途径
                        Else
                            .Cell(flexcpData, lngRow, COL_ID) = 2 '中药煎法
                        End If
                    ElseIf Val(.TextMatrix(lngRow - 1, COL_相关ID)) = rsSend!ID Then
                        If InStr(",5,6,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 Then
                            .RowHidden(lngRow) = True
                            .Cell(flexcpData, lngRow, COL_ID) = 1 '给药途径
                        ElseIf .TextMatrix(lngRow - 1, COL_诊疗类别) = "C" Then
                            .Cell(flexcpData, lngRow, COL_ID) = 4 '采集方法
                        Else
                            .Cell(flexcpData, lngRow, COL_ID) = 3 '中药用法
                        End If
                    End If
                ElseIf InStr(",5,6,", rsSend!诊疗类别) = 0 And Not IsNull(rsSend!相关ID) Then
                    '附加手术,手术麻醉,检查部位,一并采集的检验项目
                    .RowHidden(lngRow) = True
                End If
                
                '排开一般的叮嘱(不含给药途径,中药煎法,用法,采集方法,输血途径)
                If NVL(rsSend!执行性质, 0) = 0 Then
                    If InStr(",1,2,3,4,5,", CLng(.Cell(flexcpData, lngRow, COL_ID))) = 0 _
                        And InStr(",5,6,7,", rsSend!诊疗类别) = 0 Then
                        Call .RemoveItem(lngRow): GoTo NextLoop
                    End If
                End If
                
                '一般列赋值
                '---------------------------------------------------------------
                .Cell(flexcpData, lngRow, COL_婴儿) = CLng(NVL(rsSend!婴儿, 0))
                If NVL(rsSend!婴儿, 0) = 0 Then
                    .TextMatrix(lngRow, COL_婴儿) = "病人"
                Else
                    .TextMatrix(lngRow, COL_婴儿) = "婴儿" & rsSend!婴儿
                    .ColHidden(COL_婴儿) = False '有婴儿医嘱时才显示
                End If
                
                .TextMatrix(lngRow, COL_ID) = rsSend!ID
                .TextMatrix(lngRow, COL_相关ID) = NVL(rsSend!相关ID)
                .TextMatrix(lngRow, COL_医嘱状态) = rsSend!医嘱状态
                .TextMatrix(lngRow, COL_诊疗类别) = rsSend!诊疗类别
                .TextMatrix(lngRow, COL_诊疗项目ID) = rsSend!诊疗项目ID
                .TextMatrix(lngRow, col_医嘱内容) = NVL(rsSend!医嘱内容)
                
                .TextMatrix(lngRow, COL_标本部位) = NVL(rsSend!标本部位)
                .TextMatrix(lngRow, COL_检查方法) = NVL(rsSend!检查方法)
                .TextMatrix(lngRow, COL_执行标记) = NVL(rsSend!执行标记, 0)
                .TextMatrix(lngRow, COL_紧急标志) = NVL(rsSend!紧急标志, 0)
                If InStr(",4,5,6,7,", "," & rsSend!诊疗类别 & ",") = 0 Then .TextMatrix(lngRow, COL_计算方式) = NVL(rsSend!计算方式, 0)
                .TextMatrix(lngRow, COL_执行安排) = NVL(rsSend!执行安排, 0)
                .TextMatrix(lngRow, COL_开始时间) = Format(NVL(rsSend!开始执行时间), "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(lngRow, COL_申请序号) = NVL(rsSend!申请序号, 0)
                .TextMatrix(lngRow, COL_执行分类) = NVL(rsSend!执行分类, 0)
                .TextMatrix(lngRow, COL_跟踪在用) = NVL(rsSend!跟踪在用, 0)
                .TextMatrix(lngRow, COL_用药理由) = NVL(rsSend!用药理由)
                '电子签名标识
                .TextMatrix(lngRow, COL_签名ID) = NVL(rsSend!签名ID)
                If Val(.TextMatrix(lngRow, COL_签名ID)) <> 0 Then
                    Set .Cell(flexcpPicture, lngRow, col_医嘱内容) = frmIcons.imgSign.ListImages("签名").Picture
                End If
                
                '用于显示计价医嘱
                .Cell(flexcpData, lngRow, COL_诊疗类别) = CStr(NVL(rsSend!类别名称))
                .Cell(flexcpData, lngRow, col_医嘱内容) = CStr(NVL(rsSend!诊疗项目))
                
                .TextMatrix(lngRow, COL_医生嘱托) = NVL(rsSend!医生嘱托)
                .Cell(flexcpData, lngRow, COL_医生嘱托) = CStr(NVL(rsSend!摘要))
                
                .TextMatrix(lngRow, COL_执行时间) = NVL(rsSend!执行时间方案)
                .TextMatrix(lngRow, COL_频率) = NVL(rsSend!执行频次)
                
                .TextMatrix(lngRow, COL_病人科室ID) = NVL(rsSend!病人科室id)
                .TextMatrix(lngRow, COL_开嘱科室ID) = NVL(rsSend!开嘱科室id)
                .TextMatrix(lngRow, COL_开嘱医生) = NVL(rsSend!开嘱医生)
                .TextMatrix(lngRow, COL_开嘱时间) = Format(NVL(rsSend!开嘱时间), "yyyy-MM-dd HH:mm:ss")
                                
                '采集方法行显示检验项目的执行科室
                If Val(.Cell(flexcpData, lngRow, COL_ID)) = 4 Then
                    .TextMatrix(lngRow, COL_执行科室) = .TextMatrix(lngRow - 1, COL_执行科室)
                Else
                    .TextMatrix(lngRow, COL_执行科室) = NVL(rsSend!执行科室)
                End If
                .TextMatrix(lngRow, COL_执行科室ID) = NVL(rsSend!执行科室ID)
                
                .TextMatrix(lngRow, COL_计价特性) = NVL(rsSend!计价特性, 0)
                .TextMatrix(lngRow, COL_执行性质ID) = NVL(rsSend!执行性质, 0)
                .TextMatrix(lngRow, COL_操作类型) = NVL(rsSend!操作类型)
                                
                '采集方式的管码与一并的第一个检验相同
                If Val(.Cell(flexcpData, lngRow, COL_ID)) = 4 Then
                    j = .FindRow(CStr(rsSend!ID), .FixedRows, COL_相关ID)
                    If j <> -1 Then
                        .TextMatrix(lngRow, COL_试管编码) = .TextMatrix(j, COL_试管编码)
                    End If
                Else
                    .TextMatrix(lngRow, COL_试管编码) = NVL(rsSend!试管编码)
                End If
                                
                '药品相关信息
                If InStr(",5,6,7", rsSend!诊疗类别) > 0 Then
                    '药品对应的规格已撤档则不允许发送(诊疗项目本身也可以相同处理,目前暂未处理)
                    If Format(NVL(rsSend!撤档时间, "3000-01-01"), "yyyy-MM-dd") <> "3000-01-01" Or （InStr(",2,3,", NVL(rsSend!服务对象, 0)) = 0 And mlng病人性质 <> 1) Then
                        If rsSend!诊疗类别 = "7" Then
                            strTmp = "该中草药对应的中药配方无法发送：" & vbCrLf & vbCrLf & "　　" & NVL(rsSend!医嘱内容)
                        Else
                            strTmp = "该药品(及一并给药的其他药品)无法发送：" & vbCrLf & vbCrLf & "　　" & NVL(rsSend!医嘱内容)
                        End If
                        strTmp = strTmp & vbCrLf & vbCrLf & "没有发现有效的药品规格信息，该药品可能已经被停用或不能用于住院病人。"
                        strTmp = strTmp & vbCrLf & "请先到药品目录管理中处理，按[确定]继续处理其他医嘱。"
                        
                        .Redraw = flexRDDirect
                        Call .ShowCell(lngRow, COL_选择)
                        Screen.MousePointer = 0
                        MsgBox strTmp, vbInformation, gstrSysName
                        
                        '删除当前行(及相关行),及处理下一医嘱
                        Screen.MousePointer = 11
                        lngDel组ID = NVL(rsSend!相关ID, rsSend!ID)
                        Call DeleteCurRow(lngRow)
                        .Refresh: .Redraw = flexRDNone
                        lng最小次数 = 0: GoTo NextLoop
                    End If
                    
                    '毒理分类判断
                    If gbln特殊药品分开发送 Then
                        strTmp = ""
                        Select Case cboDrugType.ListIndex
                        Case 1
                            If rsSend!毒理分类 & "" <> "毒性药" Then strTmp = "1"
                        Case 2
                            If InStr(",麻醉药,精神I类,", "," & rsSend!毒理分类 & ",") = 0 Then strTmp = "1"
                        Case 3
                            If InStr(",毒性药,麻醉药,精神I类,", "," & rsSend!毒理分类 & ",") > 0 Then strTmp = "1"
                        End Select
                        
                        If strTmp <> "" Then
                            lngDel组ID = NVL(rsSend!相关ID, 0)
                            Call DeleteCurRow(lngRow, rsSend!相关ID)
                            lng最小次数 = 0: GoTo NextLoop
                        End If
                        .TextMatrix(lngRow, COL_毒理分类) = NVL(rsSend!毒理分类, "空")
                    End If
                
                    .TextMatrix(lngRow, COL_收费细目ID) = rsSend!收费细目ID
                    .TextMatrix(lngRow, COL_剂量系数) = NVL(rsSend!剂量系数, 1)
                    .TextMatrix(lngRow, COL_住院包装) = NVL(rsSend!住院包装, 1)
                    .TextMatrix(lngRow, COL_住院单位) = NVL(rsSend!住院单位)
                    .TextMatrix(lngRow, COL_可否分零) = NVL(rsSend!可否分零, 0)
                    .TextMatrix(lngRow, COL_库存) = GetStock(rsSend!收费细目ID, NVL(rsSend!执行科室ID, 0), 2) '按住院包装
                ElseIf rsSend!诊疗类别 = "4" Then
                    .TextMatrix(lngRow, COL_收费细目ID) = rsSend!收费细目ID
                    .TextMatrix(lngRow, COL_剂量系数) = 1
                    .TextMatrix(lngRow, COL_住院包装) = 1
                    .TextMatrix(lngRow, COL_住院单位) = NVL(rsSend!计算单位)
                    .TextMatrix(lngRow, COL_库存) = GetStock(rsSend!收费细目ID, NVL(rsSend!执行科室ID, 0), 2)
                End If
                                                                        
                '计算发送次数，执行的分解时间等
                '---------------------------------------------------------------
                If rsSend!诊疗类别 = "7" Then
                    .TextMatrix(lngRow, COL_次数) = NVL(rsSend!总给予量, 0)
                    If Not IsNull(rsSend!执行时间方案) Or NVL(rsSend!间隔单位) = "分钟" Then
                        .TextMatrix(lngRow, COL_分解时间) = Calc次数分解时间(rsSend!总给予量, rsSend!开始执行时间, CDate("3000-01-01"), "", NVL(rsSend!执行时间方案), rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位)
                        .TextMatrix(lngRow, COL_首次时间) = Format(Split(.TextMatrix(lngRow, COL_分解时间), ",")(0), "yyyy-MM-dd HH:mm")
                        .TextMatrix(lngRow, COL_末次时间) = Format(Split(.TextMatrix(lngRow, COL_分解时间), ",")(rsSend!总给予量 - 1), "yyyy-MM-dd HH:mm")
                    Else
                        '无分解时间(一次性临嘱未输入执行时间而无法分解)
                        '记录费用发生时间(以医嘱开始执行时间)
                        .Cell(flexcpData, lngRow, COL_分解时间) = Format(rsSend!开始执行时间, "yyyy-MM-dd HH:mm:ss")
                    End If
                    
                    .TextMatrix(lngRow, COL_单量) = NVL(rsSend!单次用量) '单量
                    .TextMatrix(lngRow, COL_单量单位) = NVL(rsSend!计算单位)
                    .TextMatrix(lngRow, COL_总量) = NVL(rsSend!总给予量, 0) '付数
                    .TextMatrix(lngRow, COL_总量单位) = "付"
                ElseIf InStr(",5,6,", rsSend!诊疗类别) > 0 Then
                    '计算临嘱用药次数
                    If NVL(rsSend!频率次数, 0) = 0 Or NVL(rsSend!频率间隔, 0) = 0 Then
                        lng次数 = 1 '设置为一次性的临嘱药品
                    ElseIf NVL(rsSend!天数, 0) <> 0 And Not IsNull(rsSend!执行频次) Then
                        '一个频率周期的次数
                        If rsSend!间隔单位 = "周" Then
                            lng次数 = IntEx(rsSend!天数 * (rsSend!频率次数 / 7))
                        ElseIf rsSend!间隔单位 = "天" Then
                            lng次数 = IntEx(rsSend!天数 * (rsSend!频率次数 / rsSend!频率间隔))
                        ElseIf rsSend!间隔单位 = "小时" Then
                            lng次数 = IntEx(rsSend!天数 * (rsSend!频率次数 / rsSend!频率间隔) * 24)
                        ElseIf rsSend!间隔单位 = "分钟" Then
                            lng次数 = IntEx(rsSend!天数 * (rsSend!频率次数 / rsSend!频率间隔) * (24 * 60))
                        End If
                    Else
                         '可分零药品时,按总量对单量的倍数计算给药途径的次数,不可分零与一次性使用药品时，按总量对（单量与剂量系数比值取整）的倍数计算给药途径的次数，
                         '否则按一个频率周期的次数计算
                        If NVL(rsSend!可否分零, 0) = 0 And NVL(rsSend!单次用量, 0) <> 0 Then
                            lng次数 = IntEx(rsSend!总给予量 * rsSend!剂量系数 / rsSend!单次用量)
                        ElseIf (NVL(rsSend!可否分零, 0) = 1 Or NVL(rsSend!可否分零, 0) = 2) And NVL(rsSend!单次用量, 0) <> 0 Then
                            lng次数 = IntEx(rsSend!总给予量 / IntEx(rsSend!单次用量 / rsSend!剂量系数))
                        Else
                            lng次数 = NVL(rsSend!频率次数, 0)
                        End If
                    End If
                    If Not IsNull(rsSend!执行时间方案) Or NVL(rsSend!间隔单位) = "分钟" Then
                        str分解时间 = Calc次数分解时间(lng次数, rsSend!开始执行时间, CDate("3000-01-01"), "", NVL(rsSend!执行时间方案), rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位)
                        If str分解时间 <> "" Then
                            .TextMatrix(lngRow, COL_分解时间) = str分解时间
                            .TextMatrix(lngRow, COL_首次时间) = Format(Split(str分解时间, ",")(0), "yyyy-MM-dd HH:mm")
                            .TextMatrix(lngRow, COL_末次时间) = Format(Split(str分解时间, ",")(lng次数 - 1), "yyyy-MM-dd HH:mm")
                        End If
                    Else
                        '无分解时间(一次性临嘱未输入执行时间而无法分解)
                        '记录费用发生时间(以医嘱开始执行时间)
                        .Cell(flexcpData, lngRow, COL_分解时间) = Format(rsSend!开始执行时间, "yyyy-MM-dd HH:mm:ss")
                    End If
                    .TextMatrix(lngRow, COL_次数) = lng次数
                    .TextMatrix(lngRow, COL_单量) = FormatEx(NVL(rsSend!单次用量), 5)
                    .TextMatrix(lngRow, COL_单量单位) = NVL(rsSend!计算单位)
                    .TextMatrix(lngRow, COL_总量) = FormatEx(rsSend!总给予量 / rsSend!住院包装, 5) '以住院单位显示
                    .TextMatrix(lngRow, COL_总量单位) = NVL(rsSend!住院单位)
                    
                    If lng次数 < lng最小次数 Or lng最小次数 = 0 Then lng最小次数 = lng次数
                ElseIf rsSend!诊疗类别 = "E" And CLng(.Cell(flexcpData, lngRow, COL_ID)) <> 0 Then
                    '给药途径,中药煎法,中药用法,采集方法,输血途径
                    '一并给药的按最小次数发送(影响给药途径计费)
                    If .Cell(flexcpData, lngRow, COL_ID) = 1 Then '给药途径
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_相关ID)) = rsSend!ID Then
                                If Val(.TextMatrix(j, COL_次数)) > lng最小次数 Then
                                    .TextMatrix(j, COL_次数) = lng最小次数
                                    If .TextMatrix(j, COL_分解时间) <> "" Then
                                        .TextMatrix(j, COL_分解时间) = Trim分解时间(lng最小次数, .TextMatrix(j, COL_分解时间))
                                        .TextMatrix(j, COL_首次时间) = Format(Split(.TextMatrix(j, COL_分解时间), ",")(0), "yyyy-MM-dd HH:mm")
                                        .TextMatrix(j, COL_末次时间) = Format(Split(.TextMatrix(j, COL_分解时间), ",")(lng最小次数 - 1), "yyyy-MM-dd HH:mm")
                                    End If
                                End If
                            Else
                                Exit For
                            End If
                        Next
                        lng最小次数 = 0
                    End If
                    
                    .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngRow - 1, COL_次数) '付数或次数
                    .TextMatrix(lngRow, COL_次数) = .TextMatrix(lngRow - 1, COL_次数)
                    If .Cell(flexcpData, lngRow, COL_ID) = 3 Then '中药用法
                        .TextMatrix(lngRow, COL_总量单位) = "付"
                    Else
                        .TextMatrix(lngRow, COL_总量单位) = NVL(rsSend!计算单位)
                    End If
                    
                    .TextMatrix(lngRow, COL_分解时间) = .TextMatrix(lngRow - 1, COL_分解时间)
                    .Cell(flexcpData, lngRow, COL_分解时间) = .Cell(flexcpData, lngRow - 1, COL_分解时间)
                    If .TextMatrix(lngRow, COL_诊疗类别) = "E" And .TextMatrix(lngRow, COL_操作类型) = "6" And .Cell(flexcpData, lngRow - 1, COL_分解时间) <> .TextMatrix(lngRow, COL_开始时间) Then
                        .Cell(flexcpData, lngRow, COL_分解时间) = .TextMatrix(lngRow, COL_开始时间)
                    End If
                    .TextMatrix(lngRow, COL_首次时间) = .TextMatrix(lngRow - 1, COL_首次时间)
                    .TextMatrix(lngRow, COL_末次时间) = .TextMatrix(lngRow - 1, COL_末次时间)
                Else
                    '其它非药临嘱:采集方法在上面的分支中已作处理
                    If IsNull(rsSend!相关ID) Or (Not IsNull(rsSend!相关ID) And rsSend!诊疗类别 = "C") Then '主要医嘱,包括检验组合
                        If rsSend!诊疗类别 = "K" Then
                            '输血途径的执行次数
                            dbl总量 = NVL(rsSend!总给予量, 0)
                            If IsNull(rsSend!执行时间方案) And (NVL(rsSend!频率次数, 0) = 0 Or NVL(rsSend!频率间隔, 0) = 0 Or IsNull(rsSend!间隔单位)) Then
                                lng次数 = 1
                            Else
                                lng次数 = NVL(rsSend!频率次数, 1)
                            End If
                        Else
                            dbl总量 = NVL(rsSend!总给予量, 1)
                            lng次数 = IntEx(dbl总量 / NVL(rsSend!单次用量, 1))
                        End If
                        
                        If IsNull(rsSend!执行时间方案) And (NVL(rsSend!频率次数, 0) = 0 Or NVL(rsSend!频率间隔, 0) = 0 Or IsNull(rsSend!间隔单位)) Then
                            '执行频率为"一次性"的项目
                            str分解时间 = ""
                            .Cell(flexcpData, lngRow, COL_频率) = 1
                        Else
                            '执行频率为"可选频率"的项目:下医嘱时应输入了总量
                            If Not IsNull(rsSend!执行时间方案) Or NVL(rsSend!间隔单位) = "分钟" Then
                                str分解时间 = Calc次数分解时间(lng次数, rsSend!开始执行时间, CDate("3000-01-01"), "", NVL(rsSend!执行时间方案), rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位)
                            Else
                                str分解时间 = "" '临嘱也许未输入执行时间,无法分解
                            End If
                        End If
                        .TextMatrix(lngRow, COL_次数) = lng次数
                        .TextMatrix(lngRow, COL_分解时间) = str分解时间
                        If str分解时间 <> "" Then
                            .TextMatrix(lngRow, COL_首次时间) = Format(Split(str分解时间, ",")(0), "yyyy-MM-dd HH:mm")
                            .TextMatrix(lngRow, COL_末次时间) = Format(Split(str分解时间, ",")(lng次数 - 1), "yyyy-MM-dd HH:mm")
                        Else
                            '记录费用发生时间(当无分解时间时),以医嘱的开始执行时间
                            .Cell(flexcpData, lngRow, COL_分解时间) = CStr(Format(rsSend!开始执行时间, "yyyy-MM-dd HH:mm:ss"))
                        End If
                        
                        .TextMatrix(lngRow, COL_单量) = FormatEx(NVL(rsSend!单次用量), 5)
                        If Not IsNull(rsSend!单次用量) Then
                            .TextMatrix(lngRow, COL_单量单位) = NVL(rsSend!计算单位)
                        End If
                        .TextMatrix(lngRow, COL_总量) = IIF(dbl总量 = 0, "", FormatEx(dbl总量, 5))
                        .TextMatrix(lngRow, COL_总量单位) = NVL(rsSend!计算单位)
                    Else
                        .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngRow - 1, COL_总量)
                        .TextMatrix(lngRow, COL_次数) = .TextMatrix(lngRow - 1, COL_次数)
                        .TextMatrix(lngRow, COL_分解时间) = .TextMatrix(lngRow - 1, COL_分解时间)
                        .Cell(flexcpData, lngRow, COL_分解时间) = .Cell(flexcpData, lngRow - 1, COL_分解时间)
                        .TextMatrix(lngRow, COL_首次时间) = .TextMatrix(lngRow - 1, COL_首次时间)
                        .TextMatrix(lngRow, COL_末次时间) = .TextMatrix(lngRow - 1, COL_末次时间)
                    End If
                End If
                '计算项目发送金额
                cur金额 = 0
                If Not LoadAdvicePrice(lngRow, rsSend, cur金额) Then
                    lngDel组ID = NVL(rsSend!相关ID, rsSend!ID)
                    Call DeleteCurRow(lngRow)
                    lng最小次数 = 0: GoTo NextLoop
                End If
                .TextMatrix(lngRow, COL_金额) = Format(cur金额, gstrDec)
                .Cell(flexcpData, lngRow, COL_金额) = CCur(.TextMatrix(lngRow, COL_金额))
                
                '相关行时的一些处理：累计显示金额,给药途径,用法,执行科室,执行性质
                '---------------------------------------------------------------
                If rsSend!诊疗类别 = "E" And InStr(",1,3,", Val(.Cell(flexcpData, lngRow, COL_ID))) > 0 Then '给药途径或中药用法
                    cur金额 = 0
                    lngTmp = .FindRow(CStr(rsSend!ID), , COL_相关ID)
                    
                    If .Cell(flexcpData, lngRow, COL_ID) = 1 Then '给药途径
                        '一并给药时,给药途径的金额累加显示在第一个成药中
                        .TextMatrix(lngTmp, COL_金额) = Format(Val(.TextMatrix(lngTmp, COL_金额)) + Val(.TextMatrix(lngRow, COL_金额)), gstrDec)
                        '显示给药途径,执行性质
                        For j = lngTmp To lngRow - 1
                            strTmp = ""
                            If Val(.TextMatrix(j, COL_执行性质ID)) = 5 And Val(.TextMatrix(lngRow, COL_执行性质ID)) <> 5 Then
                                If Val(.TextMatrix(j, COL_执行标记)) = 2 Then
                                    strTmp = "不取药"
                                Else
                                    strTmp = "自备药"
                                End If
                            ElseIf Val(.TextMatrix(j, COL_执行性质ID)) <> 5 And Val(.TextMatrix(lngRow, COL_执行性质ID)) = 5 Then
                                strTmp = "离院带药"
                            Else
                                strTmp = IIF(Val(.TextMatrix(j, COL_执行标记)) = 1, "自取药", "")
                            End If
                            .TextMatrix(j, COL_执行性质) = strTmp
                            .TextMatrix(j, COL_用法) = rsSend!诊疗项目
                            
                            '药品发送类别限制
                            If str药品 <> "111" Then
                                If Val(Mid(str药品, 2, 1)) = 0 And strTmp = "离院带药" _
                                    Or Val(Mid(str药品, 3, 1)) = 0 And strTmp = "自取药" _
                                    Or Val(Mid(str药品, 1, 1)) = 0 And strTmp <> "离院带药" And strTmp <> "自取药" Then
                                    lngDel组ID = NVL(rsSend!相关ID, rsSend!ID)
                                    Call DeleteCurRow(lngRow)
                                    lng最小次数 = 0: GoTo NextLoop
                                End If
                            End If
                        Next
                    Else
                        '药品的执行性质
                        strTmp = ""
                        If Val(.TextMatrix(lngTmp, COL_执行性质ID)) = 5 And Val(.TextMatrix(lngRow, COL_执行性质ID)) <> 5 Then
                            If Val(.TextMatrix(lngTmp, COL_执行标记)) = 2 Then
                                strTmp = "不取药"
                            Else
                                strTmp = "自备药"
                            End If
                        ElseIf Val(.TextMatrix(lngTmp, COL_执行性质ID)) <> 5 And Val(.TextMatrix(lngRow, COL_执行性质ID)) = 5 Then
                            strTmp = "离院带药"
                        Else
                            strTmp = IIF(Val(.TextMatrix(lngTmp, COL_执行标记)) = 1, "自取药", "")
                        End If
                    
                        '药品发送类别限制
                        If str药品 <> "111" Then
                            If Val(Mid(str药品, 2, 1)) = 0 And strTmp = "离院带药" _
                                Or Val(Mid(str药品, 3, 1)) = 0 And strTmp = "自取药" _
                                Or Val(Mid(str药品, 1, 1)) = 0 And strTmp <> "离院带药" And strTmp <> "自取药" Then
                                lngDel组ID = NVL(rsSend!相关ID, rsSend!ID)
                                Call DeleteCurRow(lngRow)
                                lng最小次数 = 0: GoTo NextLoop
                            End If
                        End If
                    
                        '中药用法,煎法
                        str用法 = rsSend!诊疗项目
                        If Val(.Cell(flexcpData, lngRow - 1, COL_ID)) = 2 Then
                            str用法 = str用法 & "|" & sys.RowValue("诊疗项目目录", Val(.TextMatrix(lngRow - 1, COL_诊疗项目ID)), "名称")
                        End If
                        For j = lngTmp To lngRow
                            .TextMatrix(j, COL_用法) = str用法 '用于填写收发记录
                            cur金额 = cur金额 + Val(.TextMatrix(j, COL_金额))
                        Next
                        .TextMatrix(lngRow, COL_金额) = Format(cur金额, gstrDec)
                        '显示执行性质
                        .TextMatrix(lngRow, COL_执行性质) = strTmp
                        '显示配方执行科室
                        .TextMatrix(lngRow, COL_执行科室) = .TextMatrix(lngTmp, COL_执行科室)
                    End If
                    
                    '使相关医嘱选择状态相同(固为库存的原因；非药医嘱不用)
                    For j = lngTmp To lngRow
                        If .Cell(flexcpData, j, COL_选择) <> 0 Then
                            Call RowSelectSame(j, COL_选择)
                            Exit For '一个禁止,全部禁止
                        End If
                    Next
                    If j > lngRow Then
                        For j = lngRow To lngTmp Step -1
                            If InStr(",5,6,7,", .TextMatrix(j, COL_诊疗类别)) > 0 Then
                                If .Cell(flexcpPicture, j, COL_选择) Is Nothing Then
                                    Call RowSelectSame(j, COL_选择)
                                    Exit For '最后不选,全部不选
                                End If
                            End If
                        Next
                    End If
                ElseIf InStr(",5,6,7,", rsSend!诊疗类别) = 0 Then
                    If Not IsNull(rsSend!相关ID) And rsSend!诊疗类别 <> "C" Then
                        '其它非药医嘱
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_ID)) = rsSend!相关ID Then
                                .TextMatrix(j, COL_金额) = Format(Val(.TextMatrix(j, COL_金额)) + Val(.TextMatrix(lngRow, COL_金额)), gstrDec)
                                Exit For
                            End If
                        Next
                        
                        '输血途径
                        If rsSend!诊疗类别 = "E" And Val(.Cell(flexcpData, lngRow, COL_ID)) = 5 Then
                            .TextMatrix(lngRow - 1, COL_用法) = rsSend!诊疗项目
                        End If
                    ElseIf Val(.Cell(flexcpData, lngRow, COL_ID)) = 4 Then
                        '检验标本采集方法为显示行
                        .TextMatrix(lngRow, COL_用法) = rsSend!诊疗项目
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_相关ID)) = rsSend!ID Then
                                .TextMatrix(lngRow, COL_金额) = Format(Val(.TextMatrix(lngRow, COL_金额)) + Val(.TextMatrix(j, COL_金额)), gstrDec)
                            Else
                                Exit For
                            End If
                        Next
                    End If
                End If

                '药品、卫材库存检查(0-不检查;1-检查,不足提醒;2-检查，不足禁止),自备药不检查
                '---------------------------------------------------------------
                If InStr(",5,6,7,", rsSend!诊疗类别) > 0 And NVL(rsSend!执行性质, 0) <> 5 Then
                    Call CheckStock(lngRow, rsSend, bln药品库存提示, bln药品时价提示, bln药品默认发送)
                    Call CheckDrug零差价(lngRow, bln药品零差价提示)
                ElseIf rsSend!诊疗类别 = "4" And NVL(rsSend!跟踪在用, 0) = 1 Then
                    Call CheckStock(lngRow, rsSend, bln卫材库存提示, bln卫材时价提示, bln卫材默认发送)
                End If
                
NextLoop:       '---------------------------------------------------------------
                Progress = i / rsSend.RecordCount * 100
                txtPer.Text = CInt(psb.value) & "%"
                txtPer.Refresh
                rsSend.MoveNext
            Next
        End With
    End If
    With vsAdvice
        .AutoSize col_医嘱内容
        .RowHeight(0) = 320
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        
        '电子签名图标对齐
        .Cell(flexcpPictureAlignment, .FixedRows, col_医嘱内容, .Rows - 1, col_医嘱内容) = 0
        
        .Col = .FixedCols
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) Then
                .Row = i: Exit For
            End If
        Next
        
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        
        '如果有停用的项目，则提示
        If str停用 <> "" Then
            Call MsgBox("诊疗项目：" & Mid(str停用, 2) & " 已经停用，不能发送。", vbInformation, Me.Caption)
        End If
        
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    
    Call ShowSendTotal
    Progress = 0: Screen.MousePointer = 0
    LoadAdviceSend = True
    Exit Function
errH:
    vsAdvice.Redraw = flexRDDirect
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        vsAdvice.Redraw = flexRDNone
        Resume
    End If
    Call SaveErrLog
    Progress = 0
End Function

Private Sub CheckDrug零差价(ByVal lngRow As Long, ByRef bln提示 As Boolean)
'功能：发送过程中对零差价药品进行检查禁止
    Dim strTmp As String
    Dim blnTmp As Boolean
    Dim vMsg As VbMsgBoxResult
    
    With vsAdvice
        If 0 <> Val(.TextMatrix(lngRow, COL_收费细目ID)) And 0 <> Val(.TextMatrix(lngRow, COL_执行科室ID)) And .Cell(flexcpData, lngRow, COL_选择) <> 1 Then
            If InitObjPublicDrug Then
                blnTmp = gobjPublicDrug.zlCheckPriceAdjustBySell(Val(.TextMatrix(lngRow, COL_收费细目ID)), Val(.TextMatrix(lngRow, COL_执行科室ID)), False)
                If Not blnTmp Then
                    strTmp = "在(" & .TextMatrix(lngRow, COL_执行科室) & ")中药品""" & .TextMatrix(lngRow, col_医嘱内容) & """" & vbCrLf & vbCrLf & _
                        "不满足零差价管理的要求：成本价和售价不一致，不能销售出库。" & vbCrLf & vbCrLf & _
                        "请联系药房或药剂科进行调价处理。"
                    
                    If bln提示 Then
                        .Redraw = flexRDDirect:
                        Call .ShowCell(lngRow, COL_选择)
                        Screen.MousePointer = 0
                        vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, True)
                        If vMsg = vbIgnore Then bln提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        Screen.MousePointer = 11
                        .Refresh: .Redraw = flexRDNone
                    Else
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub CheckStock(ByVal lngRow As Long, rsSend As ADODB.Recordset, Optional bln库存提示 As Boolean, Optional bln时价提示 As Boolean, Optional bln默认发送 As Boolean)
'功能：根据库存检查参数检查发送药品、跟踪卫材的库存
'参数：lngRow=医嘱行号,rsSend=当前发送医嘱信息
'      bln库存提示,bln时价提示,bln默认发送=用于提示框相关显示控制
'返回：根据提示，是否对选择状态进行了处理
    Dim int库存检查 As Integer, dbl总量 As Double
    Dim dbl可用库存 As Double, dbl已发库存 As Double
    Dim bln分批时价 As Boolean, bln分批 As Boolean, bln时价 As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        '药品库存检查(0-不检查;1-检查,不足提醒;2-检查，不足禁止)
        int库存检查 = TheStockCheck(Val(.TextMatrix(lngRow, COL_执行科室ID)), .TextMatrix(lngRow, COL_诊疗类别))
        bln分批 = NVL(rsSend!分批, 0) = 1
        bln时价 = NVL(rsSend!是否变价, 0) = 1
        
        '分批或时价药品必须要有足够的库存,其它根据库存检查参数决定
        If int库存检查 <> 0 Or bln分批 Or bln时价 Then
            strTmp = .TextMatrix(lngRow, COL_住院单位) '卫材是散装单位
            
            '当本身就不足禁止时,分批时间就不必单独处理
            bln分批时价 = int库存检查 <> 2 And (bln分批 Or bln时价)
            
            '当前药品总量:住院包装
            If .TextMatrix(lngRow, COL_诊疗类别) = "7" Then
                '中药药房单位按不可分零处理:每付
                If Val(.TextMatrix(lngRow, COL_可否分零)) = 0 Then
                    dbl总量 = Val(.TextMatrix(lngRow, COL_总量)) * Val(.TextMatrix(lngRow, COL_单量))
                    dbl总量 = dbl总量 / Val(.TextMatrix(lngRow, COL_剂量系数)) / Val(.TextMatrix(lngRow, COL_住院包装))
                Else
                    dbl总量 = IntEx(Val(.TextMatrix(lngRow, COL_单量)) / Val(.TextMatrix(lngRow, COL_剂量系数)) / Val(.TextMatrix(lngRow, COL_住院包装)))
                    dbl总量 = dbl总量 * Val(.TextMatrix(lngRow, COL_总量))
                End If
            Else
                dbl总量 = Val(.TextMatrix(lngRow, COL_总量))
            End If
            
            '当前可用库存:住院包装,减去前面相同药品要发送的库存
            For i = lngRow - 1 To .FixedRows Step -1
                If rsSend!诊疗类别 = "4" Then
                    blnDo = .TextMatrix(i, COL_诊疗类别) = "4"
                Else
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0
                End If
                If blnDo Then
                    blnDo = Val(.TextMatrix(i, COL_收费细目ID)) = Val(.TextMatrix(lngRow, COL_收费细目ID)) _
                        And Val(.TextMatrix(i, COL_执行科室ID)) = Val(.TextMatrix(lngRow, COL_执行科室ID))
                End If
                If blnDo Then
                    blnDo = .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing
                End If
                If blnDo Then
                    If .TextMatrix(i, COL_诊疗类别) = "7" Then
                        '中药药房单位按不可分零处理:每付
                        If Val(.TextMatrix(i, COL_可否分零)) = 0 Then
                            dbl已发库存 = dbl已发库存 + _
                                Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_单量)) _
                                / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_住院包装))
                        Else
                            dbl已发库存 = dbl已发库存 + Val(.TextMatrix(i, COL_总量)) _
                                * IntEx(Val(.TextMatrix(i, COL_单量)) / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_住院包装)))
                        End If
                    Else
                        dbl已发库存 = dbl已发库存 + Val(.TextMatrix(i, COL_总量))
                    End If
                End If
            Next
            dbl可用库存 = Val(.TextMatrix(lngRow, COL_库存))
            dbl可用库存 = dbl可用库存 - dbl已发库存
            
            If dbl总量 > dbl可用库存 Then
                If (Not bln分批时价 And int库存检查 <> 0 And bln库存提示) Or (bln分批时价 And bln时价提示) Then
                    '上一次没有选择不再提示,则提示
                    If bln分批时价 Then
                        If InStr(GetInsidePrivs(p住院医嘱下达), "显示药品库存") = 0 Then
                            strTmp = "分批或时价药品""" & .TextMatrix(lngRow, col_医嘱内容) & """：" & vbCrLf & vbCrLf & _
                                "在" & .TextMatrix(lngRow, COL_执行科室) & "库存不足" & _
                                IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，" & _
                                "本次发送量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                        Else
                            strTmp = "分批或时价药品""" & .TextMatrix(lngRow, col_医嘱内容) & """库存不足：" & vbCrLf & vbCrLf & _
                                .TextMatrix(lngRow, COL_执行科室) & "可用库存：" & FormatEx(dbl可用库存, 5) & strTmp & _
                                IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，" & _
                                "本次发送量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                        End If
                    Else
                        If InStr(GetInsidePrivs(p住院医嘱下达), "显示药品库存") = 0 Then
                            strTmp = "药品""" & .TextMatrix(lngRow, col_医嘱内容) & """：" & vbCrLf & vbCrLf & _
                                "在" & .TextMatrix(lngRow, COL_执行科室) & "库存不足" & _
                                IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，" & _
                                "本次发送量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                        Else
                            strTmp = "药品""" & .TextMatrix(lngRow, col_医嘱内容) & """库存不足：" & vbCrLf & vbCrLf & _
                                .TextMatrix(lngRow, COL_执行科室) & "可用库存：" & FormatEx(dbl可用库存, 5) & strTmp & _
                                IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，" & _
                                "本次发送量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                        End If
                    End If
                    If int库存检查 = 1 And Not bln分批时价 Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "要发送该药品吗？"
                    End If
                    If rsSend!诊疗类别 = "4" Then
                        strTmp = Replace(strTmp, "药品", "卫材")
                    End If
                    
                    .Redraw = flexRDDirect:
                    Call .ShowCell(lngRow, COL_选择)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int库存检查 = 2 Or bln分批时价)
                    
                    If bln分批时价 Then
                        If vMsg = vbIgnore Then bln时价提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                    ElseIf int库存检查 = 2 Then '库存禁止
                        If vMsg = vbIgnore Then bln库存提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                    ElseIf int库存检查 = 1 Then '库存提醒
                        If vMsg = vbYes Or vMsg = vbIgnore Then
                            If vMsg = vbIgnore Then bln库存提示 = False
                            bln默认发送 = True
                        ElseIf vMsg = vbNo Or vMsg = vbCancel Then
                            If vMsg = vbCancel Then bln库存提示 = False
                            bln默认发送 = False
                            Set .Cell(flexcpPicture, lngRow, COL_选择) = Nothing '缺省不发送
                        End If
                    End If
                    
                    Screen.MousePointer = 11
                    .Refresh: .Redraw = flexRDNone
                Else
                    '上一次选择了不再提示
                    If int库存检查 = 2 Or bln分批 Or bln时价 Then
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                    ElseIf int库存检查 = 1 Then
                        '根据上一次的结果处理
                        If Not bln默认发送 Then
                            Set .Cell(flexcpPicture, lngRow, COL_选择) = Nothing '缺省不发送
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function CheckPriceStock(ByVal lngRow As Long, rsPrice As ADODB.Recordset, ByVal lng库房ID As Long, ByVal dbl数量 As Double, _
    rsTotal As ADODB.Recordset, Optional bln库存提示 As Boolean, Optional bln时价提示 As Boolean, Optional bln默认发送 As Boolean) As Boolean
'功能：发送过程中时，对非药嘱药品及跟踪在用的卫材计价进行库存检查(累计检查)
'参数：lngRow=医嘱行号
'      dbl数量=已计算好的计价数量(售价单位)
'      rsTotal=当前病人前面已累计发送的计价药品或卫材数量(售价单位)
'      bln库存提示,bln时价提示,bln默认发送=用于提示框相关显示控制
'返回：根据提示，是否对选择状态进行了处理
    Dim int库存检查 As Integer, dbl总量 As Double
    Dim dbl可用库存 As Double, dbl已发库存 As Double
    Dim bln分批时价 As Boolean, bln分批 As Boolean, bln时价 As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        '药品库存检查(0-不检查;1-检查,不足提醒;2-检查，不足禁止)
        int库存检查 = TheStockCheck(lng库房ID, rsPrice!类别)
        bln分批 = NVL(rsPrice!分批, 0) = 1
        bln时价 = NVL(rsPrice!是否变价, 0) = 1
        
        '分批或时价药品必须要有足够的库存,其它根据库存检查参数决定
        If int库存检查 <> 0 Or bln分批 Or bln时价 Then
            strTmp = NVL(rsPrice!住院单位, NVL(rsPrice!计算单位)) '用于提示
            
            '当本身就不足禁止时,分批时间就不必单独处理
            bln分批时价 = int库存检查 <> 2 And (bln分批 Or bln时价)
            
            '当前药品或卫材总量:住院包装
            dbl总量 = Format(dbl数量 / NVL(rsPrice!住院包装, 1), "0.00000")
            
            '当前可用库存:住院包装,减去前面相同药品医嘱要发送的库存
            If InStr(",5,6,7,", rsPrice!类别) > 0 Then
                For i = lngRow - 1 To .FixedRows Step -1
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0
                    If blnDo Then
                        blnDo = Val(.TextMatrix(i, COL_收费细目ID)) = rsPrice!ID And Val(.TextMatrix(i, COL_执行科室ID)) = lng库房ID
                    End If
                    If blnDo Then
                        blnDo = .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing
                    End If
                    If blnDo Then
                        If .TextMatrix(i, COL_诊疗类别) = "7" Then
                            '中药药房单位按不可分零处理:每付
                            If Val(.TextMatrix(i, COL_可否分零)) = 0 Then
                                dbl已发库存 = dbl已发库存 + _
                                    Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_单量)) _
                                    / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_住院包装))
                            Else
                                dbl已发库存 = dbl已发库存 + Val(.TextMatrix(i, COL_总量)) _
                                    * IntEx(Val(.TextMatrix(i, COL_单量)) / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_住院包装)))
                            End If
                        Else
                            dbl已发库存 = dbl已发库存 + Val(.TextMatrix(i, COL_总量))
                        End If
                    End If
                Next
            End If
            '计价部份要发送的累计数量
            rsTotal.Filter = "项目ID=" & rsPrice!ID & " And 库房ID=" & lng库房ID
            Do While Not rsTotal.EOF
                dbl已发库存 = dbl已发库存 + Format(rsTotal!数量 / NVL(rsPrice!住院包装, 1), "0.00000")
                rsTotal.MoveNext
            Loop
            
            dbl可用库存 = Format(GetStock(rsPrice!ID, lng库房ID, 2), "0.00000")
            dbl可用库存 = dbl可用库存 - dbl已发库存
            
            If dbl总量 > dbl可用库存 Then
                If (Not bln分批时价 And int库存检查 <> 0 And bln库存提示) Or (bln分批时价 And bln时价提示) Then
                    '上一次没有选择不再提示,则提示
                    If bln分批时价 Then
                        If InStr(GetInsidePrivs(p住院医嘱下达), "显示药品库存") = 0 Then
                            strTmp = "医嘱""" & .TextMatrix(lngRow, col_医嘱内容) & """的分批或时价计价项目：" & vbCrLf & vbCrLf & _
                                """" & rsPrice!名称 & """在" & sys.RowValue("部门表", lng库房ID, "名称") & "库存不足" & _
                                IIF(dbl已发库存 <> 0, "(排开前面相同项目所需库存)", "") & "，本次发送数量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                        Else
                            strTmp = "医嘱""" & .TextMatrix(lngRow, col_医嘱内容) & """的分批或时价计价项目""" & rsPrice!名称 & """库存不足：" & _
                                vbCrLf & vbCrLf & sys.RowValue("部门表", lng库房ID, "名称") & "可用库存：" & FormatEx(dbl可用库存, 5) & strTmp & _
                                IIF(dbl已发库存 <> 0, "(排开前面相同项目所需库存)", "") & "，本次发送数量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                        End If
                    Else
                        If InStr(GetInsidePrivs(p住院医嘱下达), "显示药品库存") = 0 Then
                            strTmp = "医嘱""" & .TextMatrix(lngRow, col_医嘱内容) & """的计价项目：" & vbCrLf & vbCrLf & _
                                """" & rsPrice!名称 & """在" & sys.RowValue("部门表", lng库房ID, "名称") & "库存不足" & _
                                IIF(dbl已发库存 <> 0, "(排开前面相同项目所需库存)", "") & "，本次发送数量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                        Else
                            strTmp = "医嘱""" & .TextMatrix(lngRow, col_医嘱内容) & """的计价项目""" & rsPrice!名称 & """库存不足：" & _
                                vbCrLf & vbCrLf & sys.RowValue("部门表", lng库房ID, "名称") & "可用库存：" & FormatEx(dbl可用库存, 5) & strTmp & _
                                IIF(dbl已发库存 <> 0, "(排开前面相同项目所需库存)", "") & "，本次发送数量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                        End If
                    End If
                    If int库存检查 = 1 And Not bln分批时价 Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "要发送该医嘱吗？"
                    End If
                    
                    .Redraw = flexRDDirect
                    .Row = GetVisibleRow(lngRow, True)
                    Call .ShowCell(.Row, COL_选择)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int库存检查 = 2 Or bln分批时价)
                    
                    If bln分批时价 Then
                        If vMsg = vbIgnore Then bln时价提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int库存检查 = 2 Then '库存禁止
                        If vMsg = vbIgnore Then bln库存提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int库存检查 = 1 Then '库存提醒
                        If vMsg = vbYes Or vMsg = vbIgnore Then
                            If vMsg = vbIgnore Then bln库存提示 = False
                            bln默认发送 = True
                        ElseIf vMsg = vbNo Or vMsg = vbCancel Then
                            If vMsg = vbCancel Then bln库存提示 = False
                            bln默认发送 = False
                            Set .Cell(flexcpPicture, lngRow, COL_选择) = Nothing '缺省不发送
                            CheckPriceStock = True
                        End If
                    End If
                    Screen.MousePointer = 11
                    .Refresh: .Redraw = flexRDNone
                Else
                    '上一次选择了不再提示
                    If int库存检查 = 2 Or bln分批 Or bln时价 Then
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int库存检查 = 1 Then
                        '根据上一次的结果处理
                        If Not bln默认发送 Then
                            Set .Cell(flexcpPicture, lngRow, COL_选择) = Nothing '缺省不发送
                            CheckPriceStock = True
                        End If
                    End If
                End If
            End If
        End If
        
        '如果未提示或要发送,加入累计发送数量
        If Not CheckPriceStock Then
            rsTotal.AddNew
            If Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
                rsTotal!医嘱ID = Val(.TextMatrix(lngRow, COL_相关ID))
            Else
                rsTotal!医嘱ID = Val(.TextMatrix(lngRow, COL_ID))
            End If
            rsTotal!项目ID = rsPrice!ID
            rsTotal!库房ID = lng库房ID
            rsTotal!数量 = dbl数量
            rsTotal.Update
        End If
    End With
End Function

Private Sub vsPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng行号 As Long, i As Long
    Dim str项目IDs As String, blnCancel As Boolean
    Dim lng医嘱ID As Long, lng原项目ID As Long
    Dim int费用性质 As Integer, vPoint As PointAPI
    Dim strSQL2 As String
    
    With vsPrice
        lng行号 = Val(.TextMatrix(Row, COLP_行号))
        If Col = COLP_收费项目 Then
            '不能选择已有的项目
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, COLP_行号)) = lng行号 And lng行号 <> 0 And i <> Row Then
                    str项目IDs = str项目IDs & "," & Val(.TextMatrix(i, COLP_收费细目ID))
                End If
            Next
            str项目IDs = Mid(str项目IDs, 2)
            
            strSQL = _
                " Select Distinct 0 as 末级,To_Number('999999999'||类型) as ID,-NULL as 上级ID," & _
                " CHR(13)||类型 as 编码,Decode(类型,1,'西成药',2,'中成药',3,'中草药',7,'卫生材料') as 名称," & _
                " NULL as 单位,NULL as 规格,NULL as 产地,NULL as 类别,NULL as 费用类型,NULL as 医保大类,NULL as 说明,NULL as 价格," & _
                " -NULL as 原价ID,-NULL as 现价ID,-NULL as 缺省价格ID,-NULL as 是否变价ID,Null as 类别ID,-Null as 跟踪在用ID" & _
                " From 诊疗分类目录 Where 类型 in (1,2,3,7) And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as 末级,-ID as ID,Nvl(-上级ID,To_Number('999999999'||类型)) as 上级ID,编码,名称," & _
                " NULL as 单位,NULL as 规格,NULL as 产地,NULL as 类别,NULL as 费用类型,NULL as 医保大类,NULL as 说明,NULL as 价格," & _
                " -NULL as 原价ID,-NULL as 现价ID,-NULL as 缺省价格ID,-NULL as 是否变价ID,Null as 类别ID,-Null as 跟踪在用ID" & _
                " From 诊疗分类目录 Where 类型 in (1,2,3,7) And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With 上级ID is NULL Connect by Prior ID=上级ID"
            strSQL = strSQL & " Union ALL " & _
                " Select 0 as 末级,ID,上级ID,编码,名称,NULL as 单位,NULL as 规格,NULL as 产地,NULL as 类别,NULL as 费用类型,NULL as 医保大类," & _
                " NULL as 说明,NULL as 价格,-NULL as 原价ID,-NULL as 现价ID,-NULL as 缺省价格ID,-NULL as 是否变价ID,Null as 类别ID,-Null as 跟踪在用ID" & _
                " From 收费分类目录 Start With 上级ID is NULL Connect by Prior ID=上级ID"
            strSQL2 = _
                " Select 末级,ID,上级ID,编码,名称,单位,规格,产地,类别,费用类型,医保大类,说明," & _
                " Decode(Nvl(是否变价,0),1,Decode(Instr('567',类别ID),0,Sum(Nvl(原价,0))||'-'||Sum(Nvl(现价,0)),'时价'),Sum(现价)) as 价格," & _
                " Sum(原价) as 原价ID,Sum(现价) as 现价ID,Sum(缺省价格) as 缺省价格ID,是否变价 as 是否变价ID,类别ID,跟踪在用ID" & _
                " From (" & _
                " Select Distinct 1 as 末级,A.ID,Decode(Instr('567',A.类别),0,A.分类ID,-E.分类ID) as 上级ID,A.编码,A.名称," & _
                " A.计算单位 as 单位,A.规格,A.产地,C.名称 as 类别,A.费用类型,N.名称 as 医保大类,A.说明,B.原价,B.现价,B.缺省价格,A.是否变价," & _
                " A.类别 as 类别ID,-Null as 跟踪在用ID" & _
                " From 收费项目目录 A,收费价目 B,收费项目类别 C,药品规格 D,诊疗项目目录 E,保险支付项目 M,保险支付大类 N" & _
                " Where A.ID=B.收费细目ID  [选择替换的过条件1] And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "4", "5", "6") & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                " And A.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3)" & IIF(str项目IDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                " And A.类别 Not IN('4','J','1') And A.类别=C.编码 And A.ID=D.药品ID(+) And D.药名ID=E.ID(+)" & _
                " And A.ID=M.收费细目ID(+) And M.大类ID=N.ID(+) And M.险类(+)=[2]" & _
                " And (Nvl(a.执行科室,0) <> 4 Or Exists (Select 1 From 收费执行科室 W Where w.收费细目id = a.Id And (w.病人来源=2 or (w.病人来源 is Null And Nvl(w.开单科室id,[3]) = [3]))))" & _
                " And (a.类别 Not in ('5','6','7') Or Exists(Select 1 From 收费执行科室 W Where w.收费细目id=a.Id And Nvl(w.开单科室id,[3])=[3]))"
            If DeptExist("发料部门", 2) Then
                strSQL2 = strSQL2 & " Union ALL " & _
                    " Select Distinct 1 as 末级,A.ID,-E.分类ID as 上级ID,A.编码,A.名称," & _
                    " A.计算单位 as 单位,A.规格,A.产地,C.名称 as 类别,A.费用类型,N.名称 as 医保大类,A.说明," & _
                    " B.原价,B.现价,B.缺省价格,A.是否变价,A.类别 as 类别ID,D.跟踪在用 as 跟踪在用ID" & _
                    " From 收费项目目录 A,收费价目 B,收费项目类别 C,材料特性 D,诊疗项目目录 E,保险支付项目 M,保险支付大类 N" & _
                    " Where A.ID=B.收费细目ID [选择替换的过条件2]  And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "4", "5", "6") & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                    " And A.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3)" & IIF(str项目IDs <> "", " And Instr([1],','||A.ID||',')=0", "") & _
                    " And A.类别='4' And A.类别=C.编码 And A.ID=D.材料ID And D.诊疗ID=E.ID" & _
                    " And A.ID=M.收费细目ID(+) And M.大类ID=N.ID(+) And M.险类(+)=[2]" & _
                    " And Exists(Select 1 From 收费执行科室 W Where w.收费细目id=a.Id And Nvl(w.开单科室id,[3])=[3])"
            End If
            strSQL2 = strSQL2 & " ) Group by 末级,ID,上级ID,类别,编码,名称,单位,规格,产地,费用类型,医保大类,说明,是否变价,类别ID,跟踪在用ID"
            '[选择替换的过条件1],[选择替换的过条件2],这两个串在选器中处理的
            '要确保 "占位参数" 在最后一位，该参数在选择器中拼接，要解决4000长度的问题
            Set rsTmp = ShowSQLSelectCIS(Me, strSQL, strSQL2, 2, "收费项目", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, "," & str项目IDs & ",", mint险类, mlng病人科室id, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "占位参数")
            If Not rsTmp Is Nothing Then
                '非本科执行的医嘱不允许输入变价项目
                If lng行号 <> 0 Then
                    If NVL(rsTmp!是否变价ID, 0) = 1 And Not (InStr(",5,6,7,", rsTmp!类别ID) > 0 Or rsTmp!类别ID = "4" And NVL(rsTmp!跟踪在用ID, 0) = 1) Then
                        If Not Check本科执行(Val(vsAdvice.TextMatrix(lng行号, COL_执行科室ID))) Then
                            MsgBox "该医嘱非本科执行，不允许对变价项目""" & rsTmp!名称 & """定价。该计价项目需要手工计价。", vbInformation, gstrSysName
                            .SetFocus: Exit Sub
                        End If
                    End If
                End If
                
                '医保对码检查
                If CheckItemInsure(rsTmp) Then
                    .SetFocus: Exit Sub
                End If
                
                lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                Call SetItemInput(Row, rsTmp, lng医嘱ID, int费用性质, lng原项目ID)
                If lng行号 <> 0 Then
                    Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                End If
                Call EnterNextCell(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "没有可用的收费项目，请先到收费项目管理中设置！", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        ElseIf Col = COLP_执行科室 Then
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
            If .TextMatrix(Row, COLP_收费类别) = "4" Then
                '跟踪在用的卫材
                strSQL = _
                    " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                    " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                    " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
                    " And B.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3) And B.部门ID=C.ID" & _
                    " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                    " And (A.病人来源 is NULL Or A.病人来源=" & IIF(mbytSendKind = EInBilling, 2, 1) & ")" & _
                    " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                    " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                    " And A.收费细目ID=[1]" & _
                    " Order by B.服务对象,C.编码"
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "发料部门", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_收费细目ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)))
            ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_收费类别)) > 0 Then
                '药品
                '药品从系统指定的储备药房中找
                If Not Check上班安排(True) Then
                    strSQL = _
                        " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                        " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                        " Where A.执行科室ID+0=B.部门ID And B.工作性质=[3]" & _
                        " And B.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3) And B.部门ID=C.ID" & _
                        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " And (A.病人来源 is NULL Or A.病人来源=" & IIF(mbytSendKind = EInBilling, 2, 1) & ")" & _
                        " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                        " And A.收费细目ID=[1]" & _
                        " Order by B.服务对象,C.编码"
                Else
                    strSQL = _
                        " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                        " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
                        " Where A.执行科室ID+0=B.部门ID And B.工作性质=[3]" & _
                        " And B.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3) And B.部门ID=C.ID" & _
                        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " And D.部门ID=C.ID And D.星期=To_Number(To_Char(Sysdate,'D'))-1" & _
                        " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                        " And (A.病人来源 is NULL Or A.病人来源=" & IIF(mbytSendKind = EInBilling, 2, 1) & ")" & _
                        " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                        " And A.收费细目ID=[1]" & _
                        " Order by B.服务对象,C.编码"
                End If
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "药房", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    Val(.TextMatrix(Row, COLP_收费细目ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)), _
                    decode(.TextMatrix(Row, COLP_收费类别), "5", "西药房", "6", "成药房", "7", "中药房"))
            End If
            If Not rsTmp Is Nothing Then
                .TextMatrix(Row, COLP_执行科室ID) = rsTmp!ID
                .TextMatrix(Row, Col) = rsTmp!名称
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '更新记录集
                lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng原项目ID
                    mrsPrice!执行科室ID = rsTmp!ID
                    mrsPrice.Update
                    Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                End If
                Call EnterNextCell(Row, Col)
            Else
                If Not blnCancel Then
                    MsgBox "没有找到可用的科室。", vbInformation, gstrSysName
                End If
                .SetFocus
            End If
        End If
    End With
End Sub

Private Function CheckItemInsure(rsInput As ADODB.Recordset) As Boolean
'功能：检查输入(选择)计价项目是否医保对码
'返回：如果未对码，并且提示选择不继续，则返回真。
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, int险类 As Integer
    
    If gint医保对码 = 0 Then Exit Function
    
    On Error GoTo errH
    
    strSQL = "Select 险类 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckItemInsure", mlng病人ID, mlng主页ID)
    If Not rsTmp.EOF Then int险类 = NVL(rsTmp!险类, 0)
    If int险类 <> 0 Then
        If Not ItemExistInsure(mlng病人ID, rsInput!ID, int险类) Then
            If gint医保对码 = 1 Then
                If MsgBox("项目""" & rsInput!名称 & """没有设置对应的保险项目，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    CheckItemInsure = True
                End If
            ElseIf gint医保对码 = 2 Then
                MsgBox "项目""" & rsInput!名称 & """没有设置对应的保险项目。", vbInformation, gstrSysName
                CheckItemInsure = True
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsPrice_DblClick()
    Call vsPrice_KeyPress(32)
End Sub

Private Sub vsPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsPrice
        If KeyCode = vbKeyF4 Then
            If CellEditable(.Row, .Col) And .Col = COLP_计价医嘱 Then
                Call zlcommfun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Editable And Val(.TextMatrix(.Row, COLP_固定)) = 0 Then
                If Val(.TextMatrix(.Row, COLP_行号)) <> 0 And Val(.TextMatrix(.Row, COLP_收费细目ID)) <> 0 Then
                    '医嘱如果有从项至少要保留一个(主项是固定不可动的)
                    mrsPrice.Filter = "医嘱ID=" & Val(vsAdvice.TextMatrix(Val(.TextMatrix(.Row, COLP_行号)), COL_ID)) & _
                        " And 费用性质=" & Val(.TextMatrix(.Row, COLP_费用性质)) & " And 从项=1"
                    If mrsPrice.RecordCount = 1 And .TextMatrix(.Row, COLP_从项) <> "" Then
                        MsgBox """" & .Cell(flexcpData, .Row, COLP_计价医嘱) & """至少要保留一个从属计价项目。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                
                    If MsgBox("确实要删除当前计价行吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    mrsPrice.Filter = "医嘱ID=" & Val(vsAdvice.TextMatrix(Val(.TextMatrix(.Row, COLP_行号)), COL_ID)) & _
                        " And 费用性质=" & Val(.TextMatrix(.Row, COLP_费用性质)) & " And 收费细目ID=" & Val(.TextMatrix(.Row, COLP_收费细目ID))
                    mrsPrice.Delete
                End If
                
                .RemoveItem .Row
                If .Rows = .FixedRows Then
                    .Rows = .FixedRows + 1
                    .Row = .FixedRows: .Col = COLP_计价医嘱
                End If
                
                Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
            End If
        ElseIf KeyCode > 127 Then
            '解决直接输入汉字的问题
            Call vsPrice_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsPrice_KeyPress(KeyAscii As Integer)
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterNextCell(.Row, .Col)
        Else
            If CellEditable(.Row, .Col) And (.Col = COLP_收费项目 Or .Col = COLP_执行科室) Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsPrice_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'功能：定位到价表中下一个可以输入的单元格
    Dim i As Long, j As Long
    
    With vsPrice
        '当前单元格如果未输入完整,则退出
        If CellEditable(lngRow, lngCol) Then
            If lngCol = COLP_单价 And Val(.TextMatrix(lngRow, lngCol)) = 0 Then
                Exit Sub
            ElseIf .TextMatrix(lngRow, lngCol) = "" Then
                Exit Sub
            End If
        End If
        
        '从下一单元开始循环搜索
        For i = lngRow To .Rows - 1
            For j = IIF(i = lngRow, lngCol + 1, COLP_计价医嘱) To .Cols - 1
                If CellEditable(i, j) Then Exit For
            Next
            If j <= .Cols - 1 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
        Else
            '当前表格内没有找到下一个可编辑单元,如果有需计价医嘱,则增加一新行
            If CStr(.ColData(COLP_计价医嘱)) <> "" Then
                '当前行未输入完整,则定位到不完整单元
                If .TextMatrix(lngRow, COLP_计价医嘱) = "" Then
                    .Col = COLP_计价医嘱
                ElseIf .TextMatrix(lngRow, COLP_计价数量) = "" Then
                    .Col = COLP_计价数量
                ElseIf .TextMatrix(lngRow, COLP_收费项目) = "" Then
                    .Col = COLP_收费项目
                ElseIf Val(.TextMatrix(lngRow, COLP_变价)) = 1 _
                    And Val(.TextMatrix(lngRow, COLP_单价)) = 0 _
                    And CellEditable(lngRow, COLP_单价) Then
                    .Col = COLP_单价
                Else
                    .AddItem "", .Rows
                    .Row = .Rows - 1: .Col = COLP_计价医嘱
                    
                    '缺省选择计价医嘱(如果可能)
                    Call ShowDefaultRow
                End If
            Else
                If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1 '不可编辑时随意定一个
            End If
        End If
        .ShowCell .Row, .Col
    End With
End Sub

Private Sub vsPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng行号 As Long, i As Long
    Dim str项目IDs As String, int费用性质 As Integer
    Dim lng医嘱ID As Long, lng原项目ID As Long
    Dim strTmp As String, blnCancel As Boolean
    Dim strInput As String, strMatch As String
    Dim vPoint As PointAPI
    Dim strStock As String
    
    With vsPrice
        If KeyAscii = 13 Then
            KeyAscii = 0
            lng行号 = Val(.TextMatrix(Row, COLP_行号))
            If Col = COLP_计价医嘱 Then
                '下拉时回车
                If .ComboIndex <> -1 Then
                    .TextMatrix(.Row, .Col) = .ComboItem(.ComboIndex) '不然EnterNextCell函数要退出
                    Call EnterNextCell(Row, Col)
                End If
            ElseIf Col = COLP_计价数量 Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "计价数量输入错误，不是大于零的数字或输入数值过大！", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '更新记录集
                lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng原项目ID
                    mrsPrice!数量 = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                End If
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_单价 Then
                If Not IsNumeric(.EditText) Or Val(.EditText) <= 0 Or Val(.EditText) > LONG_MAX Then
                    MsgBox "收费单价输入错误，不是大于零的数字或输入数值过大！", vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                '检查变价输入范围
                strTmp = CheckScope(.Cell(flexcpData, Row, COLP_应收金额), .Cell(flexcpData, Row, COLP_实收金额), .EditText)
                If strTmp <> "" Then
                    MsgBox strTmp, vbInformation, gstrSysName
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                .EditText = Format(.EditText, gstrDecPrice)
                .TextMatrix(Row, Col) = .EditText
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                
                '更新记录集
                lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                    mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng原项目ID
                    mrsPrice!单价 = Val(.TextMatrix(Row, Col))
                    mrsPrice.Update
                    Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                End If
                Call EnterNextCell(Row, Col)
            ElseIf Col = COLP_收费项目 And .EditText <> "" Then
                '不能选择已有的项目
                For i = .FixedRows To .Rows - 1
                    If Val(vsAdvice.TextMatrix(Val(.TextMatrix(i, COLP_行号)), COL_ID)) = Val(vsAdvice.TextMatrix(lng行号, COL_ID)) _
                        And Val(vsAdvice.TextMatrix(lng行号, COL_ID)) <> 0 And i <> Row Then
                        str项目IDs = str项目IDs & "," & Val(.TextMatrix(i, COLP_收费细目ID))
                    End If
                Next
                str项目IDs = Mid(str项目IDs, 2)
                
                If mlng西药房 <> 0 Or mlng成药房 <> 0 Or mlng中药房 <> 0 Or mlng发料部门 <> 0 Then
                    strStock = _
                        "Select A.药品ID,Sum(Nvl(A.可用数量,0)) as 库存" & _
                        " From 药品库存 A,收费项目目录 B" & _
                        " Where A.性质 = 1 And (Nvl(A.批次,0)=0 Or A.效期 Is Null Or A.效期>Trunc(Sysdate))" & _
                        " And A.库房ID=Decode(B.类别,'5',[7],'6',[8],'7',[9],'4',[10],Null)" & _
                        " And A.药品ID=B.ID And B.类别 IN('4','5','6','7')" & _
                        " Group by A.药品ID Having Sum(Nvl(A.可用数量,0))<>0"
                Else
                    strStock = "Select Null as 药品ID,Null as 库存 From Dual"
                End If
                
                '不同的输入匹配方式
                strInput = UCase(.EditText)
                strMatch = " And (A.编码 Like [1] And C.码类=[3] Or C.名称 Like [2] And C.码类=[3] Or C.简码 Like [2] And C.码类 IN([3],3))"
                If IsNumeric(strInput) Then                         '10,11.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strMatch = " And (A.编码 Like [1] And C.码类=[3] Or C.简码 Like [2] And C.码类=3)"
                ElseIf zlcommfun.IsCharAlpha(strInput) Then         '01,11.输入全是字母时只匹配简码
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strMatch = " And C.简码 Like [2] And C.码类=[3]"
                ElseIf zlcommfun.IsCharChinese(strInput) Then
                    strMatch = " And C.名称 Like [2] And C.码类=[3]"
                End If
                
                strSQL = ""
                If Not DeptExist("发料部门", 2) Then strSQL = " And A.类别<>'4'"
                strSQL = _
                    " Select A.末级,A.ID,A.类别,A.编码,A.名称,A.单位,A.规格,A.产地," & _
                    " Decode(Instr('4567',A.类别ID),0,NULL,1," & _
                    "   Decode(S.库存,NULL,NULL,LTrim(To_Char(S.库存,'999990.0000'))||A.单位)," & _
                    "   Decode(S.库存,NULL,NULL,LTrim(To_Char(S.库存/Nvl(C.住院包装,1),'999990.0000'))||C.住院单位)) as 库存," & _
                    "   A.费用类型,N.名称 as 医保大类,A.说明," & _
                    " Decode(Nvl(A.是否变价,0),1,Decode(Instr('567',A.类别ID),0,Sum(Nvl(A.原价,0))||'-'||Sum(Nvl(A.现价,0)),'时价'),Sum(A.现价)) as 价格," & _
                    " Sum(A.原价) as 原价ID,Sum(A.现价) as 现价ID,Sum(A.缺省价格) as 缺省价格ID,A.是否变价 as 是否变价ID,A.类别ID,B.跟踪在用 as 跟踪在用ID" & _
                    " From (" & _
                    " Select Distinct 1 as 末级,A.ID,a.执行科室,A.类别 as 类别ID,D.名称 as 类别,A.编码,A.名称,A.计算单位 as 单位," & _
                    " A.规格,A.产地,A.费用类型,A.说明,B.原价,B.现价,B.缺省价格,A.是否变价" & _
                    " From 收费项目目录 A,收费价目 B,收费项目别名 C,收费项目类别 D" & _
                    " Where A.ID=B.收费细目ID And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "11", "12", "13") & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                    " And A.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3)" & IIF(str项目IDs <> "", " And Instr([4],','||A.ID||',')=0", "") & _
                    " And A.ID=C.收费细目ID And A.类别=D.编码 And A.类别 Not IN('J','1')" & strSQL & strMatch & _
                    " ) A,材料特性 B,药品规格 C,保险支付项目 M,保险支付大类 N,(" & strStock & ") S" & _
                    " Where A.ID=B.材料ID(+) And A.ID=M.收费细目ID(+) And M.大类ID=N.ID(+) And M.险类(+)=[5]  And A.ID=C.药品ID(+) And A.ID=S.药品ID(+)" & _
                    " And (Nvl(a.执行科室,0) <> 4 Or Exists (Select 1 From 收费执行科室 W Where w.收费细目id = a.Id And (w.病人来源=2 or (w.病人来源 is Null And Nvl(w.开单科室id,[6]) = [6]))))" & _
                    " And (a.类别id not in ('4','5','6','7') Or Exists(Select 1 From 收费执行科室 W Where w.收费细目id=a.Id And Nvl(w.开单科室id,[6])=[6]))" & _
                    " Group by A.末级,A.ID,A.类别,A.编码,A.名称,A.单位,A.规格,A.产地,A.费用类型,C.住院单位,C.住院包装,S.库存,N.名称,A.说明,A.是否变价,A.类别ID,B.跟踪在用" & _
                    " Order by A.类别,A.编码"
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "收费项目", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", mstrLike & strInput & "%", mint简码 + 1, "," & str项目IDs & ",", mint险类, mlng病人科室id, mlng西药房, mlng成药房, mlng中药房, mlng发料部门, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                If Not rsTmp Is Nothing Then
                    '非本科执行的医嘱不允许输入变价项目
                    If lng行号 <> 0 Then
                        If NVL(rsTmp!是否变价ID, 0) = 1 And Not (InStr(",5,6,7,", rsTmp!类别ID) > 0 Or rsTmp!类别ID = "4" And NVL(rsTmp!跟踪在用ID, 0) = 1) Then
                            If Not Check本科执行(Val(vsAdvice.TextMatrix(lng行号, COL_执行科室ID))) Then
                                MsgBox "该医嘱非本科执行，不允许对变价项目""" & rsTmp!名称 & """定价。该计价项目需要手工计价。", vbInformation, gstrSysName
                                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                                Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                                .SetFocus: Exit Sub
                            End If
                        End If
                    End If
                
                    '医保对码检查
                    If CheckItemInsure(rsTmp) Then
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                        Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                        .SetFocus: Exit Sub
                    End If
                
                    lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                    int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                    lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                    Call SetItemInput(Row, rsTmp, lng医嘱ID, int费用性质, lng原项目ID)
                    .EditText = .TextMatrix(Row, Col) '直接输入匹配需要
                    If lng行号 <> 0 Then
                        Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                    End If
                    Call EnterNextCell(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "没有找到可用的收费项目！", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                    .SetFocus
                End If
            ElseIf Col = COLP_执行科室 And .EditText <> "" Then '执行科室
                vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
                If .TextMatrix(Row, COLP_收费类别) = "4" Then
                    '跟踪在用的卫材
                    strSQL = _
                        " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                        " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                        " Where A.执行科室ID+0=B.部门ID And B.工作性质='发料部门'" & _
                        " And B.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3) And B.部门ID=C.ID" & _
                        " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                        " And (A.病人来源 is NULL Or A.病人来源=" & IIF(mbytSendKind = EInBilling, 2, 1) & ")" & _
                        " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                        " And A.收费细目ID=[1] And (C.编码 Like [3] Or C.名称 Like [4] Or C.简码 Like [4])" & _
                        " Order by B.服务对象,C.编码"
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "发料部门", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_收费细目ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)), UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                ElseIf InStr(",5,6,7,", .TextMatrix(Row, COLP_收费类别)) > 0 Then
                    '药品从系统指定的储备药房中找
                    If Not Check上班安排(True) Then
                        strSQL = _
                            " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                            " From 收费执行科室 A,部门性质说明 B,部门表 C" & _
                            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[3]" & _
                            " And B.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3) And B.部门ID=C.ID" & _
                            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                            " And (A.病人来源 is NULL Or A.病人来源=" & IIF(mbytSendKind = EInBilling, 2, 1) & ")" & _
                            " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                            " And A.收费细目ID=[1] And (C.编码 Like [4] Or C.名称 Like [5] Or C.简码 Like [5])" & _
                            " Order by B.服务对象,C.编码"
                    Else
                        strSQL = _
                            " Select Distinct C.ID,C.编码,C.名称,C.简码,B.服务对象 as 范围ID" & _
                            " From 收费执行科室 A,部门性质说明 B,部门表 C,部门安排 D" & _
                            " Where A.执行科室ID+0=B.部门ID And B.工作性质=[3]" & _
                            " And B.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3) And B.部门ID=C.ID" & _
                            " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                            " And D.部门ID=C.ID And D.星期=To_Number(To_Char(Sysdate,'D'))-1" & _
                            " And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS') " & _
                            " And (A.病人来源 is NULL Or A.病人来源=" & IIF(mbytSendKind = EInBilling, 2, 1) & ")" & _
                            " And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
                            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
                            " And A.收费细目ID=[1] And (C.编码 Like [4] Or C.名称 Like [5] Or C.简码 Like [5])" & _
                            " Order by B.服务对象,C.编码"
                    End If
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "药房", False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        Val(.TextMatrix(Row, COLP_收费细目ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_病人科室ID)), _
                        decode(.TextMatrix(Row, COLP_收费类别), "5", "西药房", "6", "成药房", "7", "中药房"), _
                        UCase(.EditText) & "%", mstrLike & UCase(.EditText) & "%")
                End If
                If Not rsTmp Is Nothing Then
                    .TextMatrix(Row, COLP_执行科室ID) = rsTmp!ID
                    .TextMatrix(Row, Col) = rsTmp!名称
                    .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                    .EditText = .TextMatrix(Row, Col) '直接输入匹配需要
                    
                    '更新记录集
                    lng医嘱ID = Val(vsAdvice.TextMatrix(lng行号, COL_ID))
                    int费用性质 = Val(.TextMatrix(Row, COLP_费用性质))
                    lng原项目ID = Val(.TextMatrix(Row, COLP_收费细目ID))
                    If lng医嘱ID <> 0 And lng原项目ID <> 0 Then
                        mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng原项目ID
                        mrsPrice!执行科室ID = rsTmp!ID
                        mrsPrice.Update
                        Call ShowAdvicePrice(vsAdvice.Row) '重新计算显示
                    End If
                    Call EnterNextCell(Row, Col)
                Else
                    If Not blnCancel Then
                        MsgBox "没有找到可用的科室。", vbInformation, gstrSysName
                    End If
                    .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                    Call vsPrice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                    .SetFocus
                End If
            End If
        Else
            If Col = COLP_计价数量 Or Col = COLP_单价 Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub SetItemInput(lngRow As Long, rsInput As ADODB.Recordset, ByVal lng医嘱ID As Long, ByVal int费用性质 As Integer, ByVal lng原项目ID As Long)
    Dim lng执行科室ID As Long, lng病人科室ID As Long
    Dim lng行号 As Long, dbl单价 As Double
    Dim blnHaveSub As Boolean
    
    With vsPrice
        '记录集内容
        '表格内容:仅临时显示标记输入了项目,也可以处理为未定计价医嘱不允许输入项目
        .TextMatrix(lngRow, COLP_类别) = rsInput!类别
        .TextMatrix(lngRow, COLP_收费类别) = rsInput!类别ID
        .TextMatrix(lngRow, COLP_收费细目ID) = rsInput!ID
        .TextMatrix(lngRow, COLP_收费项目) = rsInput!名称
        If Not IsNull(rsInput!产地) Then
            .TextMatrix(lngRow, COLP_收费项目) = .TextMatrix(lngRow, COLP_收费项目) & "(" & rsInput!产地 & ")"
        End If
        If Not IsNull(rsInput!规格) Then
            .TextMatrix(lngRow, COLP_收费项目) = .TextMatrix(lngRow, COLP_收费项目) & " " & rsInput!规格
        End If
        .TextMatrix(lngRow, COLP_单位) = NVL(rsInput!单位) '都按零售单位(包括非药嘱药品计价)
        .TextMatrix(lngRow, COLP_计价数量) = 1 '缺省相对计价1,药品为计1个零售单位
        
        '执行科室
        lng行号 = Val(.TextMatrix(lngRow, COLP_行号))
        If lng行号 <> 0 Then
            lng执行科室ID = Val(vsAdvice.TextMatrix(lng行号, COL_执行科室ID))
            '非药嘱药品和跟踪在用的卫材专门求执行科室
            If rsInput!类别ID = "4" And NVL(rsInput!跟踪在用ID, 0) = 1 Or InStr(",5,6,7,", rsInput!类别ID) > 0 Then
                lng病人科室ID = Val(vsAdvice.TextMatrix(lng行号, COL_病人科室ID))
                lng执行科室ID = Get收费执行科室ID(mlng病人ID, 0, rsInput!类别ID, rsInput!ID, 4, lng病人科室ID, 0, 2, lng执行科室ID, , , 2)
            End If
        End If
        .TextMatrix(lngRow, COLP_执行科室) = sys.RowValue("部门表", lng执行科室ID, "名称")
        .TextMatrix(lngRow, COLP_执行科室ID) = lng执行科室ID
        
        '单价计算处理:药嘱的药品计价不可能在这里处理
        If InStr(",5,6,7,", rsInput!类别ID) > 0 Then
            If NVL(rsInput!是否变价ID, 0) = 0 Then
                dbl单价 = NVL(rsInput!现价ID, 0)
            ElseIf lng行号 <> 0 Then
                '按每次缺省一个零售单位,当前发送数次计算
                dbl单价 = CalcDrugPrice(rsInput!ID, lng执行科室ID, Val(vsAdvice.TextMatrix(lng行号, COL_总量)), , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
            End If
            .TextMatrix(lngRow, COLP_单价) = Format(dbl单价, gstrDecPrice)
                        
            '时价药品不输入价格
            .TextMatrix(lngRow, COLP_变价) = 0
            .Cell(flexcpData, lngRow, COLP_应收金额) = 0
            .Cell(flexcpData, lngRow, COLP_实收金额) = 0
        ElseIf rsInput!类别ID = "4" And NVL(rsInput!跟踪在用ID, 0) = 1 And NVL(rsInput!是否变价ID, 0) = 1 Then
            '跟踪在用的时价卫材和药品一样计算
            dbl单价 = 0
            If lng行号 <> 0 Then
                dbl单价 = CalcDrugPrice(rsInput!ID, lng执行科室ID, Val(vsAdvice.TextMatrix(lng行号, COL_总量)), , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
            End If
            .TextMatrix(lngRow, COLP_变价) = 0
            .TextMatrix(lngRow, COLP_单价) = Format(dbl单价, gstrDecPrice)
            .Cell(flexcpData, lngRow, COLP_应收金额) = 0
            .Cell(flexcpData, lngRow, COLP_实收金额) = 0
        Else
            If NVL(rsInput!是否变价ID, 0) = 0 Then
                .TextMatrix(lngRow, COLP_变价) = 0
                .TextMatrix(lngRow, COLP_单价) = Format(NVL(rsInput!现价ID, 0), gstrDecPrice)
                .Cell(flexcpData, lngRow, COLP_应收金额) = 0
                .Cell(flexcpData, lngRow, COLP_实收金额) = 0
            Else
                .TextMatrix(lngRow, COLP_变价) = 1
                .TextMatrix(lngRow, COLP_单价) = Format(NVL(rsInput!缺省价格ID), gstrDecPrice)
                .Cell(flexcpData, lngRow, COLP_应收金额) = NVL(rsInput!原价ID, 0)
                .Cell(flexcpData, lngRow, COLP_实收金额) = NVL(rsInput!现价ID, 0)
            End If
        End If
        
        .TextMatrix(lngRow, COLP_费用类型) = NVL(rsInput!费用类型)
        .TextMatrix(lngRow, COLP_固定) = 0
        
        '用于输入恢复
        .Cell(flexcpData, lngRow, COLP_收费项目) = .TextMatrix(lngRow, COLP_收费项目)
        .Cell(flexcpData, lngRow, COLP_计价数量) = .TextMatrix(lngRow, COLP_计价数量)
        .Cell(flexcpData, lngRow, COLP_单价) = .TextMatrix(lngRow, COLP_单价)
        .Cell(flexcpData, lngRow, COLP_执行科室) = .TextMatrix(lngRow, COLP_执行科室)
        
        '记录集内容
        If lng医嘱ID <> 0 Then
            If lng原项目ID = 0 Then
                '当前医嘱是否有从项决定新增的项目是否从项
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 从项=1"
                If Not mrsPrice.EOF Then blnHaveSub = True
                .TextMatrix(lngRow, COLP_从项) = IIF(blnHaveSub, "√", "")
            
                mrsPrice.AddNew '加入
            Else '更新
                mrsPrice.Filter = "医嘱ID=" & lng医嘱ID & " And 费用性质=" & int费用性质 & " And 收费细目ID=" & lng原项目ID
            End If
            If lng原项目ID = 0 Then
                mrsPrice!医嘱ID = lng医嘱ID
                lng行号 = Val(.TextMatrix(lngRow, COLP_行号))
                If Val(vsAdvice.TextMatrix(lng行号, COL_相关ID)) <> 0 Then
                    mrsPrice!相关ID = Val(vsAdvice.TextMatrix(lng行号, COL_相关ID))
                Else
                    mrsPrice!相关ID = Null
                End If
                mrsPrice!费用性质 = int费用性质
                mrsPrice!从项 = IIF(blnHaveSub, 1, 0)
            End If
            mrsPrice!收费方式 = 0
            mrsPrice!收费类别 = rsInput!类别ID
            mrsPrice!收费细目ID = rsInput!ID
            If lng执行科室ID <> 0 Then
                mrsPrice!执行科室ID = lng执行科室ID
            Else
                mrsPrice!执行科室ID = Null
            End If
            mrsPrice!在用 = NVL(rsInput!跟踪在用ID, 0)
            mrsPrice!变价 = NVL(rsInput!是否变价ID, 0)
            mrsPrice!单价 = Val(.TextMatrix(lngRow, COLP_单价))
            mrsPrice!数量 = 1
            mrsPrice!固定 = 0
            mrsPrice.Update
        End If
    End With
End Sub

Private Sub vsPrice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsPrice.EditSelStart = 0
    vsPrice.EditSelLength = zlcommfun.ActualLen(vsPrice.EditText)
End Sub

Private Sub vsPrice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim bln非本科 As Boolean
    
    If Not CellEditable(Row, Col, bln非本科) Then
        '非本科执行的变价项目不允许定价格
        If bln非本科 Then
            MsgBox "该医嘱非本科执行，不允许对变价项目定价。该计价项目需要手工计价。", vbInformation, gstrSysName
        End If
        Cancel = True
    Else
        If Col = COLP_计价数量 Or Col = COLP_单价 Or Col = COLP_执行科室 Then
            '必须先确定收费项目
            If vsPrice.TextMatrix(Row, COLP_收费项目) = "" Then Cancel = True
        End If
        If Col = COLP_单价 Then
            '输入变价前必须先确定计价医嘱,以决定是否可以输入(本科执行)
            If vsPrice.TextMatrix(Row, COLP_计价医嘱) = "" Then Cancel = True
        End If
    End If
    
    If Col = COLP_计价数量 Or Col = COLP_单价 Then
        vsPrice.EditMaxLength = 10
    Else
        vsPrice.EditMaxLength = 0
    End If
End Sub

Private Sub InitBillSet()
'功能：初始化医嘱记帐单据生成记录集
    Set mrsBill = New ADODB.Recordset
    
    mrsBill.Fields.Append "Key", adVarChar, 100
    mrsBill.Fields.Append "NO", adVarChar, 30
    mrsBill.Fields.Append "费用序号", adBigInt
    mrsBill.Fields.Append "发送序号", adBigInt
    mrsBill.CursorLocation = adUseClient
    mrsBill.LockType = adLockOptimistic
    mrsBill.CursorType = adOpenStatic
    mrsBill.Open
        
    Set mrsRXKey = New ADODB.Recordset
    mrsRXKey.Fields.Append "Key", adVarChar, 200
    mrsRXKey.Fields.Append "医嘱ID", adVarChar, 200
    mrsRXKey.Fields.Append "条数", adBigInt
    mrsRXKey.Fields.Append "张数", adBigInt
    mrsRXKey.CursorLocation = adUseClient
    mrsRXKey.LockType = adLockOptimistic
    mrsRXKey.CursorType = adOpenStatic
    mrsRXKey.Open
End Sub

Private Sub GetCurBillSet(ByVal strKey As String, strNO As String, lng费用序号 As Long, lng发送序号 As Long, bln收费单 As Boolean)
'功能：获取当前费用单据的NO及序号
'参数：lng费用序号=费用记录中的序号,为-1时表示不取费用序号
'      lng发送序号=发送记录中的序号,为-1时表示不取发送序号
'说明：strKey=根据记帐单据生成规则定的唯一关键字
'1.中西成药按"病人(病人ID,挂号单)_病人科室ID_开嘱科室ID_开嘱医生_执行科室ID"分号。
'2.一个配方中的所有草药分配一个独立单据号
'3.材料医嘱与成药分号规则相同。
'4.其它非药医嘱每条医嘱一个独立单据号(包括给药途径，配方煎法、用法)
'5.检查部位和附加手术与主要医嘱分配相同单据号，手术麻醉分配单独的单据号。
'6.一并采集的检验组合分配相同的单据号，标本采集方法分配单独的单据号
    mrsBill.Filter = "Key='" & strKey & "'"
    If mrsBill.EOF Then
        mrsBill.AddNew
        mrsBill!Key = strKey
        
        '取单据号
        'mrsBill!NO = zlDatabase.GetNextNo(IIF(bln收费单, 13, 14)),门诊记帐单也是14
        mlngNOSequence = mlngNOSequence + 1
        mrsBill!NO = "TemporaryNO=" & IIF(bln收费单, 13, 14) & Format(mlngNOSequence, "00000")
        
        mrsBill!费用序号 = IIF(lng费用序号 = -1, 0, 1)
        mrsBill!发送序号 = IIF(lng发送序号 = -1, 0, 1)
        mrsBill.Update
    Else
        If lng费用序号 <> -1 Then
            mrsBill!费用序号 = mrsBill!费用序号 + 1
        End If
        If lng发送序号 <> -1 Then
            mrsBill!发送序号 = mrsBill!发送序号 + 1
        End If
        mrsBill.Update
    End If
    strNO = mrsBill!NO
    If lng费用序号 <> -1 Then lng费用序号 = mrsBill!费用序号
    If lng发送序号 <> -1 Then lng发送序号 = mrsBill!发送序号
End Sub

Private Sub ReplaceTrueNO(rsSQL As ADODB.Recordset, rsUpload As ADODB.Recordset)
'功能：将临时产生的NO替换成最终保存的真实NO
    Dim strNO As String, strCur As String, strPre As String
    
    rsSQL.Filter = 0
    rsSQL.Sort = "NO"
    Do While Not rsSQL.EOF
        If Not IsNull(rsSQL!NO) Then
            strCur = Split(rsSQL!NO, "=")(1)
            If strCur <> strPre Then
                strPre = strCur
                strNO = zlDatabase.GetNextNo(Val(Left(strCur, 2)))
                            
                'rsUpload中一个NO只有一条记录
                rsUpload.Filter = "NO='" & rsSQL!NO & "'"
                If Not rsUpload.EOF Then
                    rsUpload!NO = strNO
                    rsUpload.Update
                End If
            End If
            
            rsSQL!Sql = Replace(rsSQL!Sql, rsSQL!NO, strNO)
            'rsSQL!NO = strNO '这个不更新，避免导致Sort后顺序紊乱
            rsSQL.Update
        End If
        rsSQL.MoveNext
    Loop
End Sub

Private Sub DeleteSendRow()
'功能：将待发送医嘱清单中已发送成功的的行删除
    Dim i As Long, blnDel As Boolean
    
    With vsAdvice
        .Redraw = flexRDNone
        For i = .Rows - 1 To .FixedRows Step -1
            If .RowData(i) = -1 Then .RemoveItem i: blnDel = True
        Next
        .Redraw = flexRDDirect
        
        If blnDel Then
            If .Rows = .FixedRows Then .Rows = .FixedRows + 1
            For i = .FixedRows To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i: .Col = COL_选择
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            
            vsPrice.Rows = vsPrice.FixedRows
            vsPrice.Rows = vsPrice.FixedRows + 1
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
        End If
    End With
End Sub

Private Function Get实收金额(ByVal strSQL As String) As Currency
    Dim lngPos As Long, strMatch As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    strSQL = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    strMatch = "End" & Chr(0) & Chr(1)
    strSQL = Left(strSQL, InStr(strSQL, strMatch) - 1)
    Get实收金额 = CCur(strSQL)
End Function

Private Function Set实收金额(ByVal strSQL As String, ByVal cur金额 As Currency) As String
    Dim strLeft As String, strRight As String
    Dim strMatch As String, strVal As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    strLeft = Mid(strSQL, 1, InStr(strSQL, strMatch) - 1)
    strMatch = "End" & Chr(0) & Chr(1)
    strRight = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    
    Set实收金额 = strLeft & cur金额 & strRight
End Function

Private Function CheckSignSend() As Boolean
'功能：检查一起签名的医嘱是否一起发送的
'说明：这里只检查新开的医嘱，已校对的医嘱发送不再受影响(不同于门诊没有校对)
    Dim col签名ID As New Collection, str签名ID As String
    Dim lng签名id As Long, strTmp As String
    Dim i As Long, j As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_医嘱状态)) = 1 Then
                '收集已签名医嘱的发送状态
                lng签名id = Val(.TextMatrix(i, COL_签名ID))
                If lng签名id <> 0 Then
                    If InStr(str签名ID & ",", "," & lng签名id & ",") > 0 Then
                        strTmp = Split(col签名ID("_" & lng签名id), "=")(1)
                        If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                            If InStr(strTmp, "1") = 0 Then
                                col签名ID.Remove "_" & lng签名id
                                col签名ID.Add lng签名id & "=" & strTmp & "1", "_" & lng签名id
                            End If
                        Else
                            If InStr(strTmp, "0") = 0 Then
                                col签名ID.Remove "_" & lng签名id
                                col签名ID.Add lng签名id & "=" & strTmp & "0", "_" & lng签名id
                            End If
                        End If
                    Else
                        str签名ID = str签名ID & "," & lng签名id
                        If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                            col签名ID.Add lng签名id & "=1", "_" & lng签名id
                        Else
                            col签名ID.Add lng签名id & "=0", "_" & lng签名id
                        End If
                    End If
                End If
            End If
        Next
            
        '检查签名情况(一次签名的医嘱必须一起发送)
        strTmp = ""
        For i = 1 To col签名ID.Count
            lng签名id = Split(col签名ID(i), "=")(0)
            str签名ID = Split(col签名ID(i), "=")(1)
            If Not (str签名ID = "1" Or str签名ID = "0") Then
                '这次签名的内容不是"都要发送或都不发送"的情况
                j = .FindRow(CStr(lng签名id), , COL_签名ID)
                Do While j <> -1
                    If Not .RowHidden(j) Then
                        If .Cell(flexcpData, j, COL_选择) = 1 Or .Cell(flexcpPicture, j, COL_选择) Is Nothing Then
                            strTmp = strTmp & vbCrLf & "●" & .TextMatrix(j, col_医嘱内容)
                        End If
                    End If
                    j = .FindRow(CStr(lng签名id), j + 1, COL_签名ID)
                Loop
                Exit For '暂只提示第一组
            End If
        Next
    End With
    
    If strTmp <> "" Then
        MsgBox "以下医嘱与其他本次要发送的医嘱一起签名，但当前处理为不发送：" & vbCrLf & strTmp & _
            vbCrLf & vbCrLf & "一起签名的医嘱必须一起发送，请调整相关医嘱的发送状态。", vbInformation, gstrSysName
        Exit Function
    End If
    CheckSignSend = True
End Function

Private Sub SeekPriceRow(ByVal lngRow As Long, ByVal lng项目ID As Long, ByVal int费用性质 As Integer, ByVal lngCol As Long)
'功能：定位到并显示指定医嘱的指定计价行
'参数：lngRow=医嘱行号
'      lng项目ID=计价项目ID
'      lngCol=计价表格显示列
    Dim k As Long
    
    With vsAdvice
        .Col = col_医嘱内容 '进入行自动ShowPrice,mrsPrice发生变化
        If Not .RowHidden(lngRow) Then
            .Row = lngRow
        Else
            If InStr(",F,D,G,C,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 And Val(.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
                '附加手术,手术麻醉,检查部位,检验组合项目
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_相关ID))), , COL_ID)
            ElseIf CLng(.Cell(flexcpData, lngRow, COL_ID)) = 1 Then
                '给药途径
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_ID))), , COL_相关ID)
            ElseIf CLng(.Cell(flexcpData, lngRow, COL_ID)) = 2 Then
                '中药煎法
                .Row = .FindRow(CStr(Val(.TextMatrix(lngRow, COL_相关ID))), lngRow + 1, COL_ID)
            End If
        End If
        For k = vsPrice.FixedRows To vsPrice.Rows - 1
            If Val(vsPrice.TextMatrix(k, COLP_行号)) = lngRow _
                And Val(vsPrice.TextMatrix(k, COLP_费用性质)) = int费用性质 _
                And Val(vsPrice.TextMatrix(k, COLP_收费细目ID)) = lng项目ID Then
                vsPrice.Row = k: vsPrice.Col = lngCol: Exit For
            End If
        Next
        Call .ShowCell(.Row, .Col)
        Call vsPrice.ShowCell(vsPrice.Row, vsPrice.Col)
    End With
End Sub

Private Function GetMergeDrugStore(ByVal lngRow As Long) As Long
'功能：获取一并给药的基准药房，用于生成发送NO的Key值
'说明：一并给药的药品发送到一起，包括自备药和不同药房的情况
    Dim lng药房ID As Long, lngBegin As Long, i As Long
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_相关ID)) <> Val(.TextMatrix(lngRow - 1, COL_相关ID)) And Val(.TextMatrix(lngRow, COL_执行科室ID)) <> 0 Then
            lng药房ID = Val(.TextMatrix(lngRow, COL_执行科室ID))
        Else
            lngBegin = .FindRow(.TextMatrix(lngRow, COL_相关ID), , COL_相关ID)
            For i = lngBegin To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    If Val(.TextMatrix(i, COL_执行科室ID)) <> 0 Then
                        lng药房ID = Val(.TextMatrix(i, COL_执行科室ID)): Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    
    GetMergeDrugStore = lng药房ID
End Function

Private Function CheckWaitExecute(ByVal lngRow As Long, ByVal byt项目检查方式 As Byte, ByVal byt药品检查方式 As Byte) As Boolean
'功能：按照指定的检查方式，对病人未执行的项圈或未发药品进行检查
'参数：byt检查方式=0-不检查,1-检查并提示,2-检查并禁止
'返回：是否继续
    Dim strTmp As String
        
    With vsAdvice
        If byt项目检查方式 <> 0 Then
            strTmp = ExistWaitExe(mlng病人ID, mlng主页ID, -1)
            If strTmp <> "" Then
                Call .ShowCell(lngRow, col_医嘱内容): .Refresh
                If byt项目检查方式 = 1 Then
                    If MsgBox("发现病人""" & mrsPati!姓名 & """存在尚未执行完成的内容：" & _
                        vbCrLf & vbCrLf & strTmp & vbCrLf & vbCrLf & "确实要发送""" & .TextMatrix(lngRow, col_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    MsgBox "发现病人""" & mrsPati!姓名 & """存在尚未执行完成的内容：" & _
                        vbCrLf & vbCrLf & strTmp & vbCrLf & vbCrLf & "医嘱""" & .TextMatrix(lngRow, col_医嘱内容) & """将不被发送。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        If byt药品检查方式 <> 0 Then
            strTmp = ExistWaitDrug(mlng病人ID, mlng主页ID, -1)
            If strTmp <> "" Then
                Call .ShowCell(lngRow, col_医嘱内容): .Refresh
                If byt药品检查方式 = 1 Then
                    If MsgBox("发现病人""" & mrsPati!姓名 & """" & _
                        strTmp & vbCrLf & vbCrLf & "确实要发送""" & .TextMatrix(lngRow, col_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Else
                    MsgBox "发现病人""" & mrsPati!姓名 & """" & _
                        strTmp & vbCrLf & vbCrLf & "医嘱""" & .TextMatrix(lngRow, col_医嘱内容) & """将不被发送", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End With
    
    CheckWaitExecute = True
End Function

Private Function SendAdvice() As Long
'功能：处理医嘱发送(这个过程中记帐报警)
'说明：逐个病人发送提交
'返回：如果成功则返回发送号
    Dim rsSQL As ADODB.Recordset
    Dim rsTotal As ADODB.Recordset
    Dim rsUpload As ADODB.Recordset
    Dim rsNumber As ADODB.Recordset
    Dim rsItems As ADODB.Recordset '用于医保管控的费用记录集,动态记录集
    Dim rsMoneyNow As ADODB.Recordset '当前病人本次要发送的费用,动态记录集
    Dim rsMoneyDay As ADODB.Recordset '当前病人当天已发送的费用,静态记录集
    Dim rsTmp As ADODB.Recordset
    Dim rsMoney As New ADODB.Recordset

    Dim i As Long, j As Long
    Dim strSQL As String, curDate As Date, strCurDate As String
    Dim blnTran As Boolean, blnBool As Boolean, strTmp As String
    Dim str类别 As String, str毒理 As String
    
    Dim lng发送号 As Long, int计费状态 As Integer, bln划价 As Boolean, int划价 As Integer, strNO As String, int执行状态 As Integer
    Dim str收费项目 As String, lng费用序号 As Long, lng费用父号 As Long, lng发送序号 As Long
    Dim int付数 As Integer, dbl数量 As Double, cur合计 As Currency, cur记帐合计 As Currency
    Dim dbl单价 As Double, dbl应收 As Double, cur应收 As Currency, cur实收 As Currency
    Dim str分解时间 As String, str首次时间 As String, str末次时间 As String, strPre诊疗单据ID As String
    Dim int配方数 As Integer, strNOKey As String, str自动发料 As String
    Dim str发生时间 As String, str登记时间 As String
    Dim dbl发送数次 As Double, blnFirst As Boolean '配方数及分号关键字
    Dim lng药品类别ID As Long, lng卫材类别ID As Long
    Dim lng执行科室ID As Long, lng病人科室ID As Long
    Dim int药品性质 As Integer, bln附加手术 As Boolean
    
    Dim rsClone As ADODB.Recordset
    Dim rsSeek As ADODB.Recordset
    Dim rsExec As ADODB.Recordset  '医嘱执行计价
    Dim strNoneSub As String, strHaveSub As String
    Dim int父序号 As Integer, lng父项目ID As Long, str实收 As String
    
    Dim bln药品时价提示 As Boolean, bln药品库存提示 As Boolean, bln药品默认发送 As Boolean
    Dim bln卫材时价提示 As Boolean, bln卫材库存提示 As Boolean, bln卫材默认发送 As Boolean
    Dim bln保险项目否 As Boolean, lng保险大类ID As Long, cur统筹金额 As Currency, str保险编码 As String, str费用类型 As String
    Dim str结算医嘱IDs As String
    Dim rsAudit As ADODB.Recordset, strAudit As String
    '电子签名
    Dim lng组ID As Long, str医嘱IDs As String, strSource As String
    Dim intRule As Integer, strSign As String, strTimeStamp As String, strTimeStampCode As String
    Dim lng证书ID As Long, lng签名id As Long
    
    Dim str领药号 As String, bln实时监控 As Boolean, strCuvetteNumber As String '生成条码
    Dim lng费用次数 As Long '一天只收一次时，本次发送应收取的费用次数
    Dim blnCheckAdvice As Boolean, strMsg As String, lngSpecialAdviceID, lngBabyNum As Long
    Dim str特殊医嘱 As String
    Dim lng主医嘱行 As Long
    Dim lng采集科室ID As Long
    Dim str给药IDs As String, str不发给药IDs As String, str医嘱内容 As String
    Dim bln自动执行 As Boolean
    Dim str部位方法 As String '检查项目的部位方法，固定格式，检查部位<sTab>检查方法，如："头部<sTab>平扫"
    Dim dblOther数量 As Double '费用项目收费次数
    Dim str关联药行  As String '关联的药品行医嘱 ,"皮试医嘱ID,药品行医嘱ID"
    Dim rs皮试 As ADODB.Recordset
    Dim strMinDate As String
    
    On Error GoTo errH
    
    '调用中联合理用药审方结果判断
    Call Check处方审查
    
    '检查一起签名的医嘱是否一起发送
    If Not CheckSignSend Then Exit Function
    
    'RIS预约检查判断提示
    Call CheckRISScheduling
    
    blnCheckAdvice = Val(zlDatabase.GetPara("特殊医嘱发送前检查未生效医嘱", glngSys, p住院医嘱发送, 0)) = 1
    
    With vsAdvice
        If mbytSendKind <> EOutCharge And mbytSendKind <> EOutBilling Then
        '暂不支持门诊留观
            If InitObjRecipeAudit(p住院医嘱下达) Then
                '处方审查系统产生待审数据
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                        If .TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "2" Then
                            str给药IDs = str给药IDs & "," & .TextMatrix(i, COL_ID)
                        End If
                    End If
                Next
                If Mid(str给药IDs, 2) <> "" Then
                    Call gobjRecipeAudit.BuildData(Mid(str给药IDs, 2), mlng病人科室id, 1, mlng病人ID, mlng主页ID, str不发给药IDs)
                End If
            End If
        End If
        
        '先检查并提示特殊医嘱:3-转科;4-术后;5-出院;6-转院,11-死亡,14-术前
        strTmp = ""
        strMinDate = "3000-01-01 00:00"
        Call InitExecRecordset(rsExec)   '医嘱执行计价
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                If .TextMatrix(i, COL_诊疗类别) = "Z" And InStr(",3,5,6,11,", Val(.TextMatrix(i, COL_操作类型))) > 0 Then
                    lngSpecialAdviceID = .TextMatrix(i, COL_ID) '只取其中一个，特殊医嘱不允许一起发
                    lngBabyNum = Val(.Cell(flexcpData, i, COL_婴儿))
                    strTmp = strTmp & vbCrLf & mrsPati!姓名 & IIF(.Cell(flexcpData, i, COL_婴儿) <> 0, "(婴儿" & .Cell(flexcpData, i, COL_婴儿) & ")", "") & "：" & .TextMatrix(i, col_医嘱内容)
                    
                    '转科医嘱处理时判断长嘱可以继续发送
                    If Val(.TextMatrix(i, COL_操作类型)) = 3 Then
                        If CheckCanSendAdvice(mlng病人ID, mlng主页ID, lngSpecialAdviceID, lngBabyNum) Then
                            Call MsgBox("发现转科医嘱：" & vbCrLf & mrsPati!姓名 & IIF(.Cell(flexcpData, i, COL_婴儿) <> 0, "(婴儿" & .Cell(flexcpData, i, COL_婴儿) & ")", "") & "：" & .TextMatrix(i, col_医嘱内容) & vbCrLf & vbCrLf & "必须将可以发送的长期医嘱处理后才能发送。", vbInformation, gstrSysName)
                            Exit Function
                        End If
                    End If

                    '转科时未审核销帐单据检查
                    If Val(.TextMatrix(i, COL_操作类型)) = 3 Then
                        If CheckWaitQuittance(mlng病人ID, mlng主页ID) Then Exit Function
                    End If
                    
                End If
                
                If Mid(gstrESign, 2, 1) = "1" Then '住院医生站启用了电子签名才检查
                    If .TextMatrix(i, COL_诊疗类别) = "Z" And InStr(",3,5,6,11,4,14,", Val(.TextMatrix(i, COL_操作类型))) > 0 Then
                        str特殊医嘱 = str特殊医嘱 & "," & .TextMatrix(i, col_医嘱内容)
                    End If
                End If
                '毒理分类判断信息收集
                If gbln特殊药品分开发送 Then
                    If cboDrugType.ListIndex = 0 Then
                        If InStr("," & str毒理 & ",", "," & .TextMatrix(i, COL_毒理分类) & ",") = 0 Then
                            str毒理 = str毒理 & "," & .TextMatrix(i, COL_毒理分类)
                        End If
                    ElseIf cboDrugType.ListIndex = 3 Then
                        str毒理 = ""
                    Else
                        str毒理 = ",毒性药"
                    End If
                End If
                If .TextMatrix(i, COL_首次时间) < strMinDate Then
                    strMinDate = .TextMatrix(i, COL_首次时间)
                End If
            End If
        Next
        If strMinDate = "3000-01-01 00:00" Then strMinDate = ""
        
        If str毒理 <> "" And cboDrugType.ListIndex = 0 Then
            If Not (str毒理 = ",毒性药" Or str毒理 = ",精神I类" Or str毒理 = ",麻醉药" Or str毒理 = ",麻醉药,精神I类" Or str毒理 = ",精神I类,麻醉药") Then
                If Not (InStr(str毒理 & ",", ",毒性药,") = 0 And InStr(str毒理 & ",", ",麻醉药,") = 0 And InStr(str毒理 & ",", ",精神I类,") = 0) Then
                    MsgBox "本次发送的医嘱中可能包含毒麻精类药品，需分别发送，请修改过滤条件重新读取医嘱后再发送。", vbInformation, gstrSysName
                    Exit Function
                Else
                    str毒理 = ""
                End If
            End If
        End If
        
        If strTmp <> "" Then
            If blnCheckAdvice Then
                strMsg = CheckUnExecutedAdvice(mlng病人ID, mlng主页ID, lngSpecialAdviceID, lngBabyNum)
            End If
            
            If strMsg <> "" Then
                Call MsgBox("发现以下特殊医嘱：" & vbCrLf & strTmp & vbCrLf & vbCrLf & "必须将" & strMsg & "处理后才能发送。", vbInformation, gstrSysName)
                Exit Function
                
            ElseIf MsgBox("要发送的医嘱中包含下列特殊医嘱：" & vbCrLf & strTmp & vbCrLf & vbCrLf & "确实要发送当前选择的医嘱吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
        
        '如果启用了电子签名，检查存在"已停止但未确认停止"的医嘱，提示护士先进行确认停止
        '因为特殊医嘱校对时会将"已停止但未确认停止"的医嘱的"执行终止时间"调整为特殊医嘱的开始执行时间，医嘱停止的签名源文包含了"执行终止时间"，这会导致签名验证无法通过
        If str特殊医嘱 <> "" Then
            str特殊医嘱 = Mid(str特殊医嘱, 2)
            If CheckStopedUnAffirm(mlng病人ID & ":" & mlng主页ID, "") Then
                MsgBox "要发送的医嘱中包含特殊医嘱：" & vbCrLf & str特殊医嘱 & _
                    vbCrLf & vbCrLf & "发送后会将未确认停止的医嘱重新停止，为了不影响签名验证，请先进行确认停止操作。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '阳性用药
        If mbln阳性用药 Then
            blnBool = Set阳性用药()
            If Not blnBool Then
                GoTo FuncEnd
            End If
        End If
        
        If Not zlPluginAdviceBeforeSend Then
            Exit Function
        End If
    End With
    
    '提取当前病人的审批项目清单
    strAudit = ""
    If Not IsNull(mrsPati!险类) Then
        bln实时监控 = gclsInsure.GetCapability(support实时监控, mlng病人ID, mrsPati!险类)
        If Val(zlDatabase.GetPara("检查医保审批", glngSys, p住院医嘱发送, "1")) = 1 Then
            Set rsAudit = GetAuditRecord(mlng病人ID, mlng主页ID)
        Else
            Set rsAudit = Nothing
        End If
    Else
        bln实时监控 = False
        Set rsAudit = Nothing '以Nothing为标志该病人不需要判断
    End If
    
    '读取药品/卫材入出类别
    lng药品类别ID = ExistIOClass(IIF(mbytSendKind = EOutCharge, 8, 9))    '8-收费处方发药；9-记帐单处方发药
    lng卫材类别ID = ExistIOClass(IIF(mbytSendKind = EOutCharge, 40, 41))        '40-收费处方发料,41-记帐单处方发料
    
    Screen.MousePointer = 11
    
    mstr领药号 = ""
    bln药品时价提示 = True: bln药品库存提示 = True: bln药品默认发送 = True
    bln卫材时价提示 = True: bln卫材库存提示 = True: bln卫材默认发送 = True
    
    Call InitBillSet
    Call InitRecordSet(rsSQL, rsTotal, rsUpload, rsNumber, rsMoneyNow, rsItems)
    lng发送号 = zlDatabase.GetNextNo(10)
    mlngNOSequence = 0 '单据号序列重新初始
    
    '这个时间发送过程中未用于停止时间,为避免与校对时间重复(取的Sysdate)
    curDate = zlDatabase.Currentdate
    strCurDate = "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    bln划价 = True '初始全部是划价
    int配方数 = 1 '表示发送的第几付配方,用于分单据号
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                '特殊医嘱：3-转科;5-出院;6-转院,11-死亡
                If .TextMatrix(i, COL_诊疗类别) = "Z" Then
                    '转科,出院,转院,死亡医嘱发送时，病人要处于正常状态
                    If .Cell(flexcpData, i, COL_婴儿) = 0 Then
                        If InStr(",3,5,6,11,", .TextMatrix(i, COL_操作类型)) > 0 And NVL(mrsPati!状态, 0) <> 0 Then
                            MsgBox "病人""" & mrsPati!姓名 & """当前处于""" & decode(NVL(mrsPati!状态, 0), 1, "等待入科", 2, "正在转科", 3, "已预出院") & """状态，" & _
                                "不能发送""" & .TextMatrix(i, col_医嘱内容) & """医嘱。", vbInformation, gstrSysName
                            Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                            Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                            GoTo NextAdvice
                        End If
                    End If
                    
                    '如果是转科、出院、转院医嘱,检查病人是否有未执行的医技项目及未发药品
                    If InStr(",3,", .TextMatrix(i, COL_操作类型)) > 0 Then
                        If Not CheckWaitExecute(i, gbyt转科检查未执行, gbyt转科检查未发药) Then
                            Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                            Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                            GoTo NextAdvice
                        End If
                    End If
                    If InStr(",5,6,", .TextMatrix(i, COL_操作类型)) > 0 Then
                        If Not CheckWaitExecute(i, gbyt出院检查未执行, gbyt出院检查未发药) Then
                            Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                            Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                            GoTo NextAdvice
                        End If
                    End If
                End If
            
                '产生单据号分配关键字
                '-----------------------------------------------------------------------------------------
                If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                    '启用参数：特殊药品分开发送 时，特殊药品医嘱的药品行单独生成单据号，一组医嘱分配一个号
                    If str毒理 <> "" Then
                        strNOKey = "中西成药_" & .TextMatrix(i, COL_相关ID)
                    Else
                        '中西成药按"病人(病人ID,挂号单)_病人科室ID_开嘱科室ID_开嘱医生_执行科室ID"分号。
                        '一并给药的，发送到一起：包括自备药和不同药房的情况
                        strNOKey = "中西成药_" & mlng病人ID & "_" & mlng主页ID & "_" & _
                            Val(.TextMatrix(i, COL_病人科室ID)) & "_" & Val(.TextMatrix(i, COL_开嘱科室ID)) & "_" & _
                            .TextMatrix(i, COL_开嘱医生) & "_" & GetMergeDrugStore(i)
                        
                        If mbln一并给药发送为一张 Then
                            If Val(.TextMatrix(i, COL_相关ID)) <> Val(.TextMatrix(i - 1, COL_相关ID)) Then
                                '再按要打印的诊疗单据分号(一并给药的，只取第一个药品的诊疗单据ID)
                                strPre诊疗单据ID = GetClinicBillID(Val(.TextMatrix(i, COL_诊疗项目ID)), IIF(mlng病人性质 = 1, 1, 2))
                            End If
                            strNOKey = strNOKey & "_" & strPre诊疗单据ID
                        Else
                            strNOKey = strNOKey & "_" & GetClinicBillID(Val(.TextMatrix(i, COL_诊疗项目ID)), IIF(mlng病人性质 = 1, 1, 2))
                        End If
                        
                        '再按处方条数限制进行分号
                        If gintRXCount > 0 And mlng病人性质 = 1 Then
                            strTmp = ""
                            If Val(.TextMatrix(i, COL_相关ID)) <> Val(.TextMatrix(i - 1, COL_相关ID)) Then
                                strTmp = GetMergeIDs(vsAdvice, i, COL_相关ID, COL_ID) '一并给药开始行或独立药品行才取值
                            End If
                            strNOKey = strNOKey & "_" & GetRXKey(mrsRXKey, strNOKey, strTmp)
                        End If
                        '给药执行科室不相同，则分配不同的NO号
                        j = .FindRow(CStr(.TextMatrix(i, COL_相关ID)), i + 1, COL_ID)
                        If j > 0 Then strNOKey = strNOKey & "_" & Val(.TextMatrix(j, COL_执行科室ID))
                    End If
                ElseIf InStr(",4,M,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                    '材料按"病人(病人ID,挂号单)_病人科室ID_开嘱科室ID_开嘱医生_执行科室ID"分号。
                    strNOKey = "材料医嘱_" & mlng病人ID & "_" & mlng主页ID & "_" & _
                        Val(.TextMatrix(i, COL_病人科室ID)) & "_" & Val(.TextMatrix(i, COL_开嘱科室ID)) & "_" & _
                        .TextMatrix(i, COL_开嘱医生) & "_" & Val(.TextMatrix(i, COL_执行科室ID))
                    '再按要打印的诊疗单据分号
                    strNOKey = strNOKey & "_" & GetClinicBillID(Val(.TextMatrix(i, COL_诊疗项目ID)), 2)
                ElseIf .TextMatrix(i, COL_诊疗类别) = "7" Then
                    '一个配方中的所有草药分配一个独立单据号
                    strNOKey = "中药配方_" & mlng病人ID & "_" & mlng主页ID & "_" & int配方数
                ElseIf Val(.TextMatrix(i, COL_相关ID)) <> 0 And .TextMatrix(i, COL_诊疗类别) = "C" Then
                    '一并采集的检验组合分配相同的单据号，标本采集方法分配单独的单据号
                    '同一个类检验型，同一个检验执行科室，同一采集管，同一个采集方式，同一个采集执行科室的检验分配相同的单据号
                    If mbln检验单独产生单据 Then
                        strNOKey = "一并采集_" & Val(.TextMatrix(i, COL_相关ID))
                    Else
                        lng主医嘱行 = .FindRow(CStr(.TextMatrix(i, COL_相关ID)), i + 1, COL_ID)
                        strNOKey = "一并采集_" & mlng病人ID & "_" & mlng主页ID & "_" & .TextMatrix(i, COL_标本部位) & "_" & _
                            .TextMatrix(i, COL_执行科室ID) & "_" & .TextMatrix(i, COL_操作类型) & "_" & .TextMatrix(i, COL_试管编码) & "_" & _
                            .TextMatrix(lng主医嘱行, COL_诊疗项目ID) & "_" & .TextMatrix(lng主医嘱行, COL_执行科室ID)
                    End If
                ElseIf Val(.TextMatrix(i, COL_相关ID)) <> 0 And InStr(",F,D,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                    '检查部位和附加手术与主要医嘱分配相同单据号，手术麻醉分配单独的单据号。
                    strNOKey = "非药医嘱_" & Val(.TextMatrix(i, COL_相关ID))
                Else
                    '其它非药医嘱每条医嘱一个独立单据号(包括给药途径，配方煎法、用法，采集方式，麻醉方式，输血医嘱/输血途径)
                    strNOKey = "非药医嘱_" & Val(.TextMatrix(i, COL_ID))
                End If
                
                '是否离院带药
                int药品性质 = 0
                If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                    int药品性质 = decode(.TextMatrix(i, COL_执行性质), "离院带药", 3, "自取药", 4, int药品性质)
                ElseIf .TextMatrix(i, COL_诊疗类别) = "7" Then
                    j = .FindRow(CStr(.TextMatrix(i, COL_相关ID)), i + 1, COL_ID)
                    If j <> -1 Then
                        int药品性质 = decode(.TextMatrix(j, COL_执行性质), "离院带药", 3, "自取药", 4, int药品性质)
                    End If
                End If
                
                '产生医嘱记帐费用:以最新价格计算
                '-----------------------------------------------------------------------------------------
                strSQL = "": str收费项目 = ""
                If InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                    '药品缺省固定为正常计价,但下医嘱时指定了为自备药(院外执行)的不读取;药品不可能为叮嘱
                    If Val(.TextMatrix(i, COL_执行性质ID)) <> 5 Then
                        strSQL = _
                            " Select A.ID,A.类别,D.名称 as 类别名称,RTrim(A.名称||' '||A.规格) as 名称," & _
                            " A.计算单位,A.是否变价,A.屏蔽费别,A.费用确认,A.加班加价,B.加班加价率,100 as 附术收费率," & _
                            " Y.住院单位,Y.住院包装,Y.剂量系数,Y.药房分批 as 分批,0 as 跟踪在用,B.收入项目ID," & _
                            " C.收据费目,1 as 数量,B.现价 as 单价,[2] as 执行科室ID,0 as 从项,0 as 费用性质,0 as 收费方式,I.要求审批" & _
                            " From 收费项目目录 A,收费价目 B,收入项目 C,收费项目类别 D,药品规格 Y,保险支付项目 I" & _
                            " Where A.ID=B.收费细目ID And B.收入项目ID=C.ID And A.类别=D.编码" & _
                            GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "4", "5", "6") & _
                            " And A.ID=Y.药品ID(+) And A.ID=[1] And A.ID=I.收费细目ID(+) And I.险类(+)=[3]" & _
                            " And ((Sysdate Between B.执行日期 and B.终止日期) Or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                            " Order by A.编码"
                    End If
                Else
                    '先删除原非药医嘱的计价(应该没有)
                    rsSQL.AddNew
                    rsSQL!类型 = 1: rsSQL!项目ID = 0: rsSQL!序号 = i
                    rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                    rsSQL!Sql = "ZL_病人医嘱计价_Delete(" & Val(.TextMatrix(i, COL_ID)) & ",1)"
                    rsSQL.Update
                    
                    '不计价,手工计价；叮嘱,院外执行的医嘱不读取
                    If Val(.TextMatrix(i, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                        mrsPrice.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                        If Not mrsPrice.EOF Then
                            For j = 1 To mrsPrice.RecordCount
                                If NVL(mrsPrice!收费细目ID, 0) <> 0 And NVL(mrsPrice!数量, 0) <> 0 Then '对照数量为0的自动过滤掉
                                    '普通项目的变价单价要求输入，包括非跟踪在用的时价卫材医嘱
                                    If NVL(mrsPrice!单价, 0) = 0 And NVL(mrsPrice!变价, 0) = 1 _
                                        And Not (InStr(",5,6,7,", mrsPrice!收费类别) > 0 Or mrsPrice!收费类别 = "4" And NVL(mrsPrice!在用, 0) = 1) Then
                                        Call SeekPriceRow(i, mrsPrice!收费细目ID, mrsPrice!费用性质, COLP_单价)
                                        Screen.MousePointer = 0
                                        MsgBox "必须为变价的收费项目确定一个收费价格。", vbInformation, gstrSysName
                                        vsPrice.SetFocus: GoTo FuncEnd
                                    End If
                                    
                                    '计价执行科室:只保存非药品及卫材医嘱的，药品和卫材计价的执行科室
                                    If InStr(",4,5,6,7,", .TextMatrix(i, COL_诊疗类别)) = 0 _
                                        And (InStr(",5,6,7,", mrsPrice!收费类别) > 0 Or mrsPrice!收费类别 = "4" And NVL(mrsPrice!在用, 0) = 1) Then
                                        lng执行科室ID = NVL(mrsPrice!执行科室ID, 0)
                                        
                                        '卫材必须设置执行科室
                                        If lng执行科室ID = 0 And mrsPrice!收费类别 = "4" Then
                                            Call SeekPriceRow(i, mrsPrice!收费细目ID, mrsPrice!费用性质, COLP_执行科室)
                                            Screen.MousePointer = 0
                                            MsgBox "卫材""" & vsPrice.TextMatrix(vsPrice.Row, COLP_收费项目) & """没有确定执行科室，请手工输入正确的执行科室。" & vbCrLf & _
                                                "如果不能确定正确的执行科室，请到""卫材目录管理""中检查存储库房设置是否正确。", vbInformation, gstrSysName
                                            vsPrice.SetFocus: GoTo FuncEnd
                                        End If
                                    Else
                                        lng执行科室ID = 0
                                    End If
                                    
                                    '药品、卫材医嘱的计价固定对应不保存；非跟踪在用的时价卫材的变价需要输入，因此要保存到计价表中
                                    If InStr(",4,5,6,7,", .TextMatrix(i, COL_诊疗类别)) = 0 _
                                        Or .TextMatrix(i, COL_诊疗类别) = "4" And NVL(mrsPrice!在用, 0) = 0 And NVL(mrsPrice!变价, 0) = 1 Then
                                        rsSQL.AddNew
                                        rsSQL!类型 = 1: rsSQL!项目ID = mrsPrice!收费细目ID: rsSQL!序号 = i
                                        rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                                        rsSQL!Sql = "ZL_病人医嘱计价_INSERT(" & _
                                            mrsPrice!医嘱ID & "," & mrsPrice!收费细目ID & "," & _
                                            NVL(mrsPrice!数量, 0) & "," & NVL(mrsPrice!单价, 0) & "," & _
                                            NVL(mrsPrice!从项, 0) & "," & ZVal(lng执行科室ID) & "," & _
                                            NVL(mrsPrice!费用性质, 0) & "," & NVL(mrsPrice!收费方式, 0) & ")"
                                        rsSQL.Update
                                    End If
                                    
                                    '临时病人医嘱计价表
                                    If Val(.TextMatrix(i, COL_总量)) <> 0 Then
                                        strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                                            "Select " & mrsPrice!收费细目ID & " as 收费细目ID," & _
                                            NVL(mrsPrice!执行科室ID, 0) & " as 执行科室ID," & _
                                            NVL(mrsPrice!数量, 0) & " as 数量," & Format(NVL(mrsPrice!单价, 0), gstrDecPrice) & " as 单价," & _
                                            NVL(mrsPrice!从项, 0) & " as 从项," & NVL(mrsPrice!费用性质, 0) & " as 费用性质," & _
                                            NVL(mrsPrice!收费方式, 0) & " as 收费方式 From Dual"
                                    End If
                                End If
                                
                                mrsPrice.MoveNext
                            Next
                        End If
                    End If
                    
                    If strSQL <> "" Then
                        strSQL = _
                            " Select A.ID,A.类别,D.名称 as 类别名称,A.名称,A.计算单位,A.是否变价," & _
                            " A.屏蔽费别,A.费用确认,A.加班加价,B.加班加价率,B.附术收费率,Y.住院单位,Y.住院包装,Y.剂量系数," & _
                            " Decode(A.类别,'4',E.在用分批,Y.药房分批) as 分批,E.跟踪在用,B.收入项目ID," & _
                            " C.收据费目,X.数量,Decode(A.是否变价,1,X.单价,B.现价) as 单价,X.执行科室ID," & _
                            " X.从项,X.费用性质,X.收费方式,I.要求审批" & _
                            " From 收费项目目录 A,收费价目 B,收入项目 C,收费项目类别 D,材料特性 E,(" & strSQL & ") X,药品规格 Y,保险支付项目 I" & _
                            " Where A.ID=B.收费细目ID And B.收入项目ID=C.ID And A.ID=E.材料ID(+)" & _
                            GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "4", "5", "6") & _
                            " And A.类别=D.编码 And X.收费细目ID=A.ID And A.ID=Y.药品ID(+)" & _
                            " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                            " And A.ID=I.收费细目ID(+) And I.险类(+)=[3]" & _
                            " Order by X.费用性质,X.从项,X.收费方式 Desc,A.ID"
                            '一定要把主项排在前面,以便于计算和在费用记录中保持主从关系
                    End If
                End If
                                
                '汇总折扣变量初始
                int父序号 = 0: lng父项目ID = 0
                strHaveSub = "": strNoneSub = ""
                Call InitSeekSet(rsSeek)
                
                '提前生成样本条码(参数"医嘱发送生成条形码"没有启用时也产生一个虚拟的条码，用于判断是否收采血管费用)
                strCuvetteNumber = ""
                If Val(.TextMatrix(i, COL_执行性质ID)) <> 0 Then
                    j = .FindRow(CStr(.TextMatrix(i, COL_相关ID)), i + 1, COL_ID)
                    If j > 0 Then lng采集科室ID = Val(.TextMatrix(j, COL_执行科室ID))
                    strCuvetteNumber = GetCuvetteNumber(rsNumber, .TextMatrix(i, COL_试管编码), _
                        Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)), .TextMatrix(i, COL_诊疗类别), Val(.TextMatrix(i, COL_操作类型)), _
                        Val(.TextMatrix(i, COL_执行科室ID)), Val(.TextMatrix(i, COL_婴儿)), Val(.TextMatrix(i, COL_诊疗项目ID)), _
                        Val(.TextMatrix(i, COL_紧急标志)), .TextMatrix(i, COL_标本部位), lng采集科室ID)
                End If
                
                '本科执行的自动执行(特殊医嘱不用处理)
                int执行状态 = 0
                bln自动执行 = False
                If mblnAutoExe Then
                    If Not (.TextMatrix(i, COL_诊疗类别) = "Z" And Val(.TextMatrix(i, COL_操作类型)) <> 0) _
                        And (mstr前提IDs = "" And (Val(.TextMatrix(i, COL_执行科室ID)) = Val(.TextMatrix(i, COL_病人科室ID)) Or _
                        Val(.TextMatrix(i, COL_执行科室ID)) = mlng病人病区ID) Or _
                            mstr前提IDs <> "" And Val(.TextMatrix(i, COL_执行科室ID)) = mlng医技科室ID) Then
                        bln自动执行 = True
                    End If
                    If bln自动执行 Then
                        bln自动执行 = CanAutoExeItem(Val(.TextMatrix(i, COL_执行科室ID)), .TextMatrix(i, COL_诊疗类别), .TextMatrix(i, COL_操作类型), Val(.TextMatrix(i, COL_执行分类)))
                    End If
                    If bln自动执行 Then
                        '执行前先结算时，优先于“执行后自动审核记帐划价单”
                        If gbln执行前先结算 And (mbytSendKind = EOutBilling Or mbytSendKind = EOutCharge) Then
                            If Not gobjSquareCard Is Nothing Then
                                str结算医嘱IDs = str结算医嘱IDs & "," & .TextMatrix(i, COL_ID)
                            End If
                        Else
                            int执行状态 = 1
                            '血库相关特殊处理
                            If gbln血库系统 Then
                                strTmp = .TextMatrix(i, COL_诊疗类别) & .TextMatrix(i, COL_操作类型)
                                    If strTmp = "E8" Or strTmp = "E9" Then
                                        strTmp = "Select 1 From 诊疗项目目录 a where a.id=[1] and nvl(a.执行分类,0)  in (0,1)"
                                        Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, Me.Caption, Val(.TextMatrix(i, COL_诊疗项目ID)))
                                        If Not rsTmp.EOF Then
                                            int执行状态 = 0
                                        End If
                                    End If
                            End If
                        End If
                    End If
                End If
                
                If Val(.TextMatrix(i, COL_相关ID)) <> 0 And .TextMatrix(i, COL_诊疗类别) = "D" Then
                    str部位方法 = .TextMatrix(i, COL_标本部位) & "<sTab>" & .TextMatrix(i, COL_检查方法)
                Else
                    str部位方法 = ""
                End If
                                
                int计费状态 = IIF(Val(.TextMatrix(i, COL_计价特性)) = 1, -1, 0) '无需计费或未计费
                        
                '分解时间
                If .TextMatrix(i, COL_分解时间) <> "" Then
                    str分解时间 = .TextMatrix(i, COL_分解时间)
                Else
                    str分解时间 = .Cell(flexcpData, i, COL_分解时间)    '开始执行时间
                End If
                
                If Len(str分解时间) > 4000 Then
                    Screen.MousePointer = 0
                    MsgBox "当前发送的医嘱时间范围太长,共需执行" & CStr(UBound(Split(str分解时间, ",")) + 1) & "次。" & vbCrLf & _
                        "超过了支持的最大次数" & CStr(UBound(Split(Mid(str分解时间, 1, 4000), ",")) + 1) & "次,请调整总量或天数后重新发送！", vbInformation, gstrSysName
                    GoTo FuncEnd
                End If
                
                If strSQL <> "" Then
                    Set rsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_收费细目ID)), Val(.TextMatrix(i, COL_执行科室ID)), Val(NVL(mrsPati!险类, 0)), mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                    If Not rsMoney.EOF Then
                        int计费状态 = 1 '已计费
                        Set rsClone = rsMoney.Clone
                    End If
                    
                    '处理收入项目级的费用明细
                    bln附加手术 = .TextMatrix(i, COL_诊疗类别) = "F" And Val(.TextMatrix(i, COL_相关ID)) <> 0
                    Do While Not rsMoney.EOF
MoneyItemBegin:
                        '执行科室ID
                        lng执行科室ID = NVL(rsMoney!执行科室ID, 0)
                        '在原值基础上取有效的非药嘱药品及跟踪卫材的执行科室
                        If InStr(",4,5,6,7", .TextMatrix(i, COL_诊疗类别)) = 0 _
                            And (rsMoney!类别 = "4" And NVL(rsMoney!跟踪在用, 0) = 1 Or InStr(",5,6,7", rsMoney!类别) > 0) Then
                            lng病人科室ID = Val(.TextMatrix(i, COL_病人科室ID))
                            lng执行科室ID = Get收费执行科室ID(mlng病人ID, 0, rsMoney!类别, rsMoney!ID, 4, lng病人科室ID, 0, 2, lng执行科室ID, , , 2)
                        End If
                        
                        '----------------------------------------
                        '根据收费方式，确定当前收费项目是否应收费
                        If rsMoney!费用性质 & "_" & rsMoney!ID <> str收费项目 Then
                            If Not AdviceMoneyMake(mlng病人ID, mlng主页ID, rsMoneyNow, rsMoneyDay, _
                                IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID))), _
                                Val(.TextMatrix(i, COL_诊疗项目ID)), rsMoney!ID, lng执行科室ID, .TextMatrix(i, COL_试管编码), _
                                rsMoney!类别, NVL(rsMoney!收费方式, 0), str分解时间, IIF(mbytSendKind = EInBilling, 2, 1), lng费用次数, Val(.TextMatrix(i, COL_总量)), _
                                Val(.TextMatrix(i, COL_ID)), lng发送号, Val(rsMoney!数量 & ""), rsExec, Val(.TextMatrix(i, COL_计算方式)), _
                                        .TextMatrix(i, COL_频率), Val(.TextMatrix(i, COL_单量)), 1, rsMoney!费用性质, .TextMatrix(i, COL_诊疗类别), strCuvetteNumber, str部位方法, dblOther数量, strMinDate) Then
                                '跳过当前收费项目(多个收入项目)
                                str收费项目 = rsMoney!费用性质 & "_" & rsMoney!ID
                                Do While rsMoney!费用性质 & "_" & rsMoney!ID = str收费项目
                                    rsMoney.MoveNext
                                    If rsMoney.EOF Then Exit Do
                                Loop
                                If rsMoney.EOF Then Exit Do
                                GoTo MoneyItemBegin
                            End If
                        End If
                        '----------------------------------------
                        
                        '检查是否需要和已经审批
                        If NVL(rsMoney!要求审批, 0) = 1 And Not rsAudit Is Nothing Then
                            rsAudit.Filter = "项目ID=" & rsMoney!ID
                            If rsAudit.EOF Then
                                If UBound(Split(strAudit, vbCrLf)) < 10 Then
                                    If InStr(strAudit, "●" & rsMoney!名称) = 0 Then
                                        strAudit = strAudit & vbCrLf & "●" & rsMoney!名称
                                    End If
                                ElseIf UBound(Split(strAudit, vbCrLf)) = 10 Then
                                    strAudit = strAudit & vbCrLf & "… …"
                                End If
                            End If
                        End If
                    
                        If InStr(",5,6,7", rsMoney!类别) > 0 Then
                            If lng药品类别ID = 0 Then
                                MsgBox "不能确定药品处方单据的入出类别,请先到入出类别管理中设置！", vbInformation, gstrSysName
                                GoTo FuncEnd
                            End If
                        
                            If InStr(",5,6,7", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                                If .TextMatrix(i, COL_诊疗类别) = "7" Then
                                    int付数 = Val(.TextMatrix(i, COL_总量))
                                    '中药药房单位按不可分零处理:每付
                                    If Val(.TextMatrix(i, COL_可否分零)) = 0 Then
                                        dbl数量 = Val(.TextMatrix(i, COL_单量)) / NVL(rsMoney!剂量系数, 1)
                                    Else
                                        dbl数量 = IntEx(Val(.TextMatrix(i, COL_单量)) / NVL(rsMoney!剂量系数, 1) / NVL(rsMoney!住院包装, 1)) * NVL(rsMoney!住院包装, 1)
                                    End If
                                Else
                                    int付数 = 1
                                    dbl数量 = Val(.TextMatrix(i, COL_总量)) * NVL(rsMoney!住院包装, 1)
                                    If rs皮试 Is Nothing Then
                                        Set rs皮试 = Get原液皮试(mlng病人ID, mlng主页ID, "")
                                    End If
                                    rs皮试.Filter = "药品ID=" & Val(rsMoney!ID & "")
                                    If Not rs皮试.EOF Then
                                        If Val(rs皮试!标号 & "") = 0 Then
                                            '进行减总量计算
                                            dbl数量 = (Val(.TextMatrix(i, COL_总量)) - 1) * NVL(rsMoney!住院包装, 1)
                                            rs皮试!标号 = Val(.TextMatrix(i, COL_ID))
                                            str关联药行 = "'" & rs皮试!皮试医嘱ID & "," & rs皮试!标号 & "'"
                                            rs皮试.Update
                                            If dbl数量 <= 0 Then
                                                rsMoney.MoveNext
                                                If rsMoney.EOF Then Exit Do
                                                GoTo MoneyItemBegin
                                            End If
                                        End If
                                    End If
                                    
                                End If
                            Else
                                int付数 = 1
                                '中药药房单位按不可分零处理:每付
                                '非药嘱药品计价:因为这里预定了售价数量,因此不作不分零处理
                                '对于收费对照中的药品，且为当天只收取一次，数量为费用次数*对照数量
                                If InStr(",2,3,4,5,6,7,9,", Val("" & rsMoney!收费方式)) > 0 Then
                                    If dblOther数量 > 0 Then
                                        dbl数量 = Format(dblOther数量, "0.00000")
                                    Else
                                        dbl数量 = Format(lng费用次数 * NVL(rsMoney!数量, 0), "0.00000")
                                    End If
                                Else
                                    dbl数量 = Val(.TextMatrix(i, COL_总量)) * NVL(rsMoney!数量, 0)
                                End If
                            End If
                            dbl数量 = Format(dbl数量, "0.00000")
                            
                            If NVL(rsMoney!是否变价, 0) = 1 Then
                                dbl单价 = Format(CalcDrugPrice(rsMoney!ID, lng执行科室ID, int付数 * dbl数量, , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                            Else
                                dbl单价 = Format(NVL(rsMoney!单价, 0), gstrDecPrice)
                            End If
                        ElseIf rsMoney!类别 = "4" And NVL(rsMoney!跟踪在用, 0) = 1 Then
                            '检查卫生材料入出类别
                            If lng卫材类别ID = 0 Then
                                Screen.MousePointer = 0
                                MsgBox "不能确定卫生材料单据的入出类别,请先到入出类别管理中设置！", vbInformation, gstrSysName
                                GoTo FuncEnd
                            End If
                            
                            int付数 = 1
                            If InStr(",1,2,3,4,5,6,7,9,", Val("" & rsMoney!收费方式)) > 0 Then
                                If dblOther数量 > 0 Then
                                    dbl数量 = Format(dblOther数量, "0.00000")
                                Else
                                    dbl数量 = Format(lng费用次数 * NVL(rsMoney!数量, 0), "0.00000")
                                End If
                            Else
                                dbl数量 = Format(Val(.TextMatrix(i, COL_总量)) * NVL(rsMoney!数量, 0), "0.00000")
                            End If
                            
                            '确定时价卫材价格
                            If NVL(rsMoney!是否变价, 0) = 1 Then
                                dbl单价 = Format(CalcDrugPrice(rsMoney!ID, lng执行科室ID, dbl数量, , True, , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                            Else
                                dbl单价 = Format(NVL(rsMoney!单价, 0), gstrDecPrice)
                            End If
                        Else
                            '总量等于单次用量乘数次。一天只收一次时，有多少天要执行，就收多少次，不管单次用量（例如：每天两次）,但要管收费对照的次数
                            int付数 = 1
                            If InStr(",1,2,3,4,5,6,7,9,", Val("" & rsMoney!收费方式)) > 0 Then
                                If dblOther数量 > 0 Then
                                    dbl数量 = Format(dblOther数量, "0.00000")
                                Else
                                    dbl数量 = Format(lng费用次数 * NVL(rsMoney!数量, 0), "0.00000")
                                End If
                            Else
                                dbl数量 = Format(Val(.TextMatrix(i, COL_总量)) * NVL(rsMoney!数量, 0), "0.00000")
                            End If
                            dbl单价 = Format(NVL(rsMoney!单价, 0), gstrDecPrice)
                        End If
                        
                        '非药嘱药品及跟踪卫材的库存检查
                        If InStr(",4,5,6,7", .TextMatrix(i, COL_诊疗类别)) = 0 _
                            And (rsMoney!类别 = "4" And NVL(rsMoney!跟踪在用, 0) = 1 Or InStr(",5,6,7", rsMoney!类别) > 0) Then
                            If TheStockCheck(lng执行科室ID, rsMoney!类别) <> 0 Or NVL(rsMoney!是否变价, 0) = 1 Or NVL(rsMoney!分批, 0) = 1 Then
                                If rsMoney!类别 = "4" Then
                                    blnBool = CheckPriceStock(i, rsMoney, lng执行科室ID, int付数 * dbl数量, rsTotal, bln卫材库存提示, bln卫材时价提示, bln卫材默认发送)
                                Else
                                    blnBool = CheckPriceStock(i, rsMoney, lng执行科室ID, int付数 * dbl数量, rsTotal, bln药品库存提示, bln药品时价提示, bln药品默认发送)
                                End If
                                If blnBool Then
                                    Call RowSelectSame(i, COL_选择, rsSQL, rsTotal, rsUpload, str医嘱IDs)
                                    '如果是签名医嘱，检查是否一同签名的医嘱必须一起发送
                                    If Val(.TextMatrix(i, COL_签名ID)) <> 0 Then
                                        If Not CheckSignSend Then
                                            GoTo FuncEnd
                                        Else
                                            Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                                            GoTo NextAdvice
                                        End If
                                    Else
                                        Call DeleteRsExec(rsExec, Val(.TextMatrix(i, COL_ID)))
                                        GoTo NextAdvice
                                    End If
                                End If
                            End If
                        End If
                            
                        '发送金额
                        dbl应收 = int付数 * dbl数量 * dbl单价
                        If bln附加手术 Then
                            dbl应收 = dbl应收 * NVL(rsMoney!附术收费率, 100) / 100
                        End If
                        
                        '处理加班加价
                        If gbln加班加价 And NVL(rsMoney!加班加价, 0) = 1 Then
                            dbl应收 = dbl应收 * (1 + NVL(rsMoney!加班加价率, 0) / 100)
                        End If
                        
                        cur应收 = Format(dbl应收, gstrDec)

                        'NO,序号
                        Call GetCurBillSet(strNOKey, strNO, lng费用序号, -1, mbytSendKind = EOutCharge)
                        rsSQL.AddNew: blnBool = False
                        If rsMoney!费用性质 & "_" & rsMoney!ID <> str收费项目 Then
                            lng费用父号 = lng费用序号
                            If rsMoney!从项 = 0 Then
                                '记录主项信息，主项肯定在从项前
                                '即使不汇总折扣，也要记录主从项关系
                                If InStr(strHaveSub & ",", "," & rsMoney!费用性质 & ",") = 0 _
                                    And InStr(strNoneSub & ",", "," & rsMoney!费用性质 & ",") = 0 Then
                                    rsClone.Filter = "费用性质=" & rsMoney!费用性质 & " And 从项=1"
                                    If Not rsClone.EOF Then
                                        int父序号 = lng费用序号
                                        lng父项目ID = rsMoney!ID
                                        
                                        rsSeek.AddNew
                                        rsSeek!费用性质 = rsMoney!费用性质
                                        rsSeek!主项标签 = rsSQL.Bookmark 'Variant(Double)
                                        rsSeek!主收入ID = rsMoney!收入项目ID
                                        rsSeek.Update
                                        strHaveSub = strHaveSub & "," & rsMoney!费用性质
                                        
                                        blnBool = True
                                    Else
                                        strNoneSub = strNoneSub & "," & rsMoney!费用性质
                                    End If
                                End If
                            End If
                        End If
                        
                        '计算汇总折扣合计
                        If gbln从项汇总折扣 And (rsMoney!从项 = 1 Or InStr(strHaveSub & ",", "," & rsMoney!费用性质 & ",") > 0) Then
                            cur实收 = cur应收
                            
                            '累计医嘱合计来计算折扣
                            rsSeek.Filter = "费用性质=" & rsMoney!费用性质
                            rsSeek!合计 = NVL(rsSeek!合计, 0) + cur实收
                            rsSeek.Update
                        ElseIf NVL(rsMoney!屏蔽费别, 0) = 0 Then
                            cur实收 = Format(ActualMoney(NVL(mrsPati!费别), rsMoney!收入项目ID, cur应收, rsMoney!ID, lng执行科室ID, int付数 * dbl数量, _
                                IIF(gbln加班加价 And NVL(rsMoney!加班加价, 0) = 1, NVL(rsMoney!加班加价率, 0) / 100, 0)), gstrDec)
                        Else
                            cur实收 = cur应收
                        End If
                        '汇总折扣时，对主项的实收金额作特殊处理
                        If gbln从项汇总折扣 And blnBool Then
                            str实收 = Chr(0) & Chr(1) & "Begin" & cur实收 & "End" & Chr(0) & Chr(1)
                        Else
                            str实收 = cur实收
                        End If
                        
                        '医保相关字段
                        bln保险项目否 = False: lng保险大类ID = 0: cur统筹金额 = 0: str保险编码 = "": str费用类型 = ""
                        If Not IsNull(mrsPati!险类) Then
                            strTmp = gclsInsure.GetItemInsure(mlng病人ID, rsMoney!ID, cur实收, False, mrsPati!险类, .Cell(flexcpData, i, COL_医生嘱托) & "||" & int付数 * dbl数量)
                            If strTmp <> "" Then
                                bln保险项目否 = Val(Split(strTmp, ";")(0)) <> 0
                                lng保险大类ID = Val(Split(strTmp, ";")(1))
                                cur统筹金额 = Format(Val(Split(strTmp, ";")(2)), gstrDec)
                                str保险编码 = CStr(Split(strTmp, ";")(3))
                                If UBound(Split(strTmp, ";")) >= 5 Then
                                    If Split(strTmp, ";")(5) <> "" Then
                                        str费用类型 = Split(strTmp, ";")(5)
                                    End If
                                End If
                            End If
                        End If
                        
                        '收集记帐报警类别
                        cur合计 = cur合计 + cur实收
                        If InStr(str类别, rsMoney!类别) = 0 Then
                            str类别 = str类别 & rsMoney!类别
                        End If
                        
                        '发生时间
                        If int药品性质 = 3 Then
                            str发生时间 = strCurDate
                        ElseIf .TextMatrix(i, COL_分解时间) <> "" Then
                            str发生时间 = "To_Date('" & Split(.TextMatrix(i, COL_分解时间), ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                        Else
                            str发生时间 = "To_Date('" & .Cell(flexcpData, i, COL_分解时间) & "','YYYY-MM-DD HH24:MI:SS')"
                        End If
                                                
                        '因为现在不计价的医嘱不产生费用,所以传入的计价特性都为(0-正常计价)
                        rsSQL!类型 = 4: rsSQL!项目ID = rsMoney!ID: rsSQL!序号 = i
                        rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                        rsSQL!NO = strNO
                        If mbytSendKind = EOutCharge Then
                            '暂未取发药窗口
                            rsSQL!Sql = "ZL_门诊划价记录_INSERT(" & _
                                "'" & strNO & "'," & lng费用序号 & "," & mlng病人ID & "," & ZVal(mlng主页ID) & ",'" & _
                                IIF(mlng病人性质 = 1, NVL(mrsPati!门诊号), NVL(mrsPati!住院号)) & "','" & IIF(mlng病人性质 = 1, "", NVL(mrsPati!床号)) & "','" & NVL(mrsPati!姓名) & "'," & _
                                "'" & NVL(mrsPati!性别) & "','" & NVL(mrsPati!年龄) & "'," & _
                                "'" & NVL(mrsPati!费别) & "',NULL," & _
                                ZVal(.TextMatrix(i, COL_病人科室ID)) & "," & ZVal(.TextMatrix(i, COL_开嘱科室ID)) & "," & _
                                "'" & .TextMatrix(i, COL_开嘱医生) & "'," & IIF(rsMoney!从项 = 1, ZVal(int父序号), "NULL") & "," & _
                                rsMoney!ID & ",'" & rsMoney!类别 & "','" & NVL(rsMoney!计算单位) & "',NULL," & _
                                int付数 & "," & dbl数量 & "," & IIF(bln附加手术, 1, 0) & "," & ZVal(lng执行科室ID) & "," & _
                                IIF(lng费用父号 = lng费用序号, "NULL", lng费用父号) & "," & rsMoney!收入项目ID & "," & _
                                "'" & NVL(rsMoney!收据费目) & "'," & dbl单价 & "," & cur应收 & "," & str实收 & "," & _
                                str发生时间 & "," & strCurDate & "," & _
                                "'医嘱发送','" & UserInfo.姓名 & "'," & _
                                "'" & .TextMatrix(i, col_医嘱内容) & "'," & Val(.TextMatrix(i, COL_ID)) & ",'" & .TextMatrix(i, COL_频率) & "'," & _
                                ZVal(.TextMatrix(i, COL_单量)) & ",'" & .TextMatrix(i, COL_用法) & "',1," & _
                                IIF(int药品性质 <> 0, int药品性质, Val(.TextMatrix(i, COL_计价特性))) & "," & IIF(mlng病人性质 = 1, "1", "2") & "," & _
                                "'" & str保险编码 & "','" & str费用类型 & "'," & IIF(bln保险项目否, 1, 0) & "," & ZVal(lng保险大类ID) & ",NULL,0," & ZVal(Val(.TextMatrix(i, COL_检查方法))) & "," & ZVal(mlng主页ID) & "," & ZVal(mlng病人病区ID) & ")"
                                
                        ElseIf mbytSendKind = EOutBilling Then
                            '是否划价费用
                            If InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                                int划价 = IIF(InStr(gstr门诊发送划价单, "5") > 0, 1, 0)
                            Else
                                int划价 = IIF(InStr(gstr门诊发送划价单, .TextMatrix(i, COL_诊疗类别)) > 0, 1, 0)
                            End If
                            If int划价 = 0 Then int划价 = IIF(NVL(rsMoney!费用确认, 0) = 1, 1, 0)
                            
                            If int划价 = 0 Or int执行状态 = 1 Then
                                bln划价 = False
                                If gdbl预存款消费验卡 <> 0 Then cur记帐合计 = cur记帐合计 + cur实收
                            End If
                            
                            '登记时间
                            If int划价 = 1 Then '与非划价的时间上区分开
                                str登记时间 = "To_Date('" & Format(DateAdd("s", 1, curDate), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                str登记时间 = strCurDate
                            End If
                        
                            rsSQL!Sql = "zl_门诊记帐记录_INSERT(" & _
                                "'" & strNO & "'," & lng费用序号 & "," & mlng病人ID & ",'" & _
                                IIF(mlng病人性质 = 1, NVL(mrsPati!门诊号), NVL(mrsPati!住院号)) & "','" & mrsPati!姓名 & "'," & _
                                "'" & mrsPati!性别 & "','" & mrsPati!年龄 & "'," & "'" & mrsPati!费别 & "',0," & Val(.Cell(flexcpData, i, COL_婴儿)) & "," & _
                                ZVal(.TextMatrix(i, COL_病人科室ID)) & "," & ZVal(.TextMatrix(i, COL_开嘱科室ID)) & "," & _
                                "'" & .TextMatrix(i, COL_开嘱医生) & "'," & IIF(rsMoney!从项 = 1, ZVal(int父序号), "NULL") & "," & _
                                rsMoney!ID & ",'" & rsMoney!类别 & "','" & rsMoney!计算单位 & "'," & _
                                int付数 & "," & dbl数量 & "," & IIF(bln附加手术, 1, 0) & "," & ZVal(lng执行科室ID) & "," & _
                                IIF(lng费用父号 = lng费用序号, "NULL", lng费用父号) & "," & rsMoney!收入项目ID & "," & _
                                "'" & rsMoney!收据费目 & "'," & dbl单价 & "," & cur应收 & "," & str实收 & "," & _
                                str发生时间 & "," & str登记时间 & "," & _
                                "'医嘱发送'," & int划价 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                                "Null,'" & .TextMatrix(i, col_医嘱内容) & "'," & Val(.TextMatrix(i, COL_ID)) & ",'" & .TextMatrix(i, COL_频率) & "'," & _
                                ZVal(.TextMatrix(i, COL_单量)) & ",'" & .TextMatrix(i, COL_用法) & "',1," & _
                                IIF(int药品性质 <> 0, int药品性质, Val(.TextMatrix(i, COL_计价特性))) & ",1,Null,0," & ZVal(Val(.TextMatrix(i, COL_检查方法))) & "," & ZVal(mlng主页ID) & "," & ZVal(mlng病人病区ID) & ")"
                            '门诊标志要填1才是门诊记帐费用，填2就是住院记帐费用了
                        Else
                            '是否划价费用
                            strTmp = mlng领药部门ID
                            If InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                                int划价 = IIF(InStr(gstr住院发送划价单, "5") > 0, 1, 0)
                                
                                j = .FindRow(CStr(.TextMatrix(i, COL_相关ID)), i + 1, COL_ID)
                                If Val(.TextMatrix(j, COL_执行科室ID)) <> 0 Then strTmp = Val(.TextMatrix(j, COL_执行科室ID))
                            Else
                                int划价 = IIF(InStr(gstr住院发送划价单, .TextMatrix(i, COL_诊疗类别)) > 0, 1, 0)
                            End If
                            If int划价 = 0 Then int划价 = IIF(NVL(rsMoney!费用确认, 0) = 1, 1, 0)
                            
                            If int划价 = 0 Or int执行状态 = 1 Then
                                bln划价 = False
                                cur记帐合计 = cur记帐合计 + cur实收
                            End If
                            
                            '登记时间
                            If int划价 = 1 Then '与非划价的时间上区分开
                                str登记时间 = "To_Date('" & Format(DateAdd("s", 1, curDate), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                str登记时间 = strCurDate
                            End If
                            
                            '收集医保上传单据号:mrsBill中的不一定产生了费用
                            If int划价 = 0 Then
                                rsUpload.Filter = "NO='" & strNO & "'"
                                If rsUpload.EOF Then
                                    rsUpload.AddNew
                                    rsUpload!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                                    rsUpload!NO = strNO
                                    rsUpload.Update
                                End If
                            End If
                            
                            rsSQL!Sql = "ZL_住院记帐记录_Insert(" & _
                                "'" & strNO & "'," & lng费用序号 & "," & mlng病人ID & "," & ZVal(mlng主页ID) & "," & _
                                IIF(IsNull(mrsPati!住院号), "NULL", "'" & mrsPati!住院号 & "'") & ",'" & NVL(mrsPati!姓名) & "'," & _
                                "'" & NVL(mrsPati!性别) & "','" & NVL(mrsPati!年龄) & "'," & _
                                "'" & NVL(mrsPati!床号) & "','" & NVL(mrsPati!费别) & "'," & _
                                ZVal(mlng病人病区ID) & "," & ZVal(.TextMatrix(i, COL_病人科室ID)) & ",0," & _
                                Val(.Cell(flexcpData, i, COL_婴儿)) & "," & _
                                ZVal(.TextMatrix(i, COL_开嘱科室ID)) & ",'" & .TextMatrix(i, COL_开嘱医生) & "'," & _
                                IIF(rsMoney!从项 = 1, ZVal(int父序号), "NULL") & "," & rsMoney!ID & "," & _
                                "'" & rsMoney!类别 & "','" & NVL(rsMoney!计算单位) & "'," & _
                                IIF(bln保险项目否, 1, 0) & "," & ZVal(lng保险大类ID) & ",'" & str保险编码 & "'," & _
                                int付数 & "," & dbl数量 & ",NULL," & ZVal(lng执行科室ID) & "," & _
                                IIF(lng费用父号 = lng费用序号, "NULL", lng费用父号) & "," & rsMoney!收入项目ID & "," & _
                                "'" & NVL(rsMoney!收据费目) & "'," & dbl单价 & "," & cur应收 & "," & str实收 & "," & _
                                cur统筹金额 & "," & str发生时间 & "," & str登记时间 & "," & _
                                "'医嘱发送'," & int划价 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "',0," & _
                                IIF(rsMoney!类别 = "4", lng卫材类别ID, lng药品类别ID) & "," & _
                                "NULL,'" & .TextMatrix(i, col_医嘱内容) & "',NULL," & Val(.TextMatrix(i, COL_ID)) & "," & _
                                "'" & .TextMatrix(i, COL_频率) & "'," & ZVal(.TextMatrix(i, COL_单量)) & "," & _
                                "'" & .TextMatrix(i, COL_用法) & "',1," & _
                                IIF(int药品性质 <> 0, int药品性质, Val(.TextMatrix(i, COL_计价特性))) & "," & _
                                "Null,'" & str费用类型 & "',Null," & strTmp & ",NULL,-1,0," & ZVal(Val(.TextMatrix(i, COL_检查方法))) & ")"
                        End If
                        rsSQL.Update
                        
                        '记录自动发料的SQL
                        If ((gbyt住院自动发料 = 1 Or gbyt住院自动发料 = 2 And lng执行科室ID = Val(.TextMatrix(i, COL_开嘱科室ID))) And mbytSendKind = EInBilling Or gbln门诊自动发料 And mbytSendKind = EOutBilling) _
                            And int划价 = 0 And lng执行科室ID <> 0 And rsMoney!类别 = "4" And NVL(rsMoney!跟踪在用, 0) = 1 Then
                            If InStr(str自动发料 & ";", ";" & strNO & "," & lng执行科室ID & ";") = 0 Then
                                rsSQL.AddNew
                                rsSQL!类型 = 6
                                rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                                rsSQL!项目ID = 0
                                rsSQL!序号 = i
                                rsSQL!NO = strNO
                                rsSQL!Sql = "zl_材料收发记录_处方发料(" & lng执行科室ID & ",25,'" & strNO & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "',1,Sysdate)"
                                rsSQL.Update
                                str自动发料 = str自动发料 & ";" & strNO & "," & lng执行科室ID
                            End If
                        End If
                        
                        '医保管控实时监测：生成费用项目记录集,以收费细目汇总
                        If Not IsNull(mrsPati!险类) And bln实时监控 Then
                            rsItems.Filter = "收费细目ID=" & rsMoney!ID
                            If rsItems.EOF Then
                                '加入收费项目对应的原始信息
                                rsItems.AddNew
                                rsItems!病人ID = mlng病人ID
                                rsItems!主页ID = mlng主页ID
                                rsItems!医嘱ID = Val(.TextMatrix(i, COL_ID))
                                rsItems!收费类别 = rsMoney!类别
                                rsItems!收费细目ID = rsMoney!ID
                                rsItems!开单人 = .TextMatrix(i, COL_开嘱医生)
                                rsItems!开单科室 = CStr(sys.RowValue("部门表", Val(.TextMatrix(i, COL_开嘱科室ID)), "名称"))
                                
                                rsItems!数量 = int付数 * dbl数量
                                rsItems!单价 = dbl单价
                            Else
                                '基于一个医嘱(诊疗项目)的收费对照不会有重复的收费细目
                                '数量：同一收费项目的不同收入项目记录相同
                                If rsMoney!费用性质 & "_" & rsMoney!ID <> str收费项目 Then
                                    rsItems!数量 = NVL(rsItems!数量, 0) + int付数 * dbl数量
                                End If
                                '单价：同一收费项目的不同收入项目累加
                                If Val(.TextMatrix(i, COL_ID)) = rsItems!医嘱ID Then
                                    rsItems!单价 = NVL(rsItems!单价, 0) + dbl单价
                                End If
                            End If
                            rsItems!实收金额 = NVL(rsItems!实收金额, 0) + cur实收
                            rsItems.Update
                        End If
                            
                        str收费项目 = rsMoney!费用性质 & "_" & rsMoney!ID
                        rsMoney.MoveNext
                    Loop
                End If
                
                '对医嘱金额进行汇总折扣处理
                If gbln从项汇总折扣 And strHaveSub <> "" Then
                    rsSeek.Filter = 0
                    Do While Not rsSeek.EOF
                        rsSQL.Bookmark = rsSeek!主项标签
                        cur实收 = Format(ActualMoney(NVL(mrsPati!费别), rsSeek!主收入ID, rsSeek!合计), gstrDec)
                        cur实收 = cur实收 - rsSeek!合计 '打折差额
                        
                        '医保管控实时监测：费用项目金额替换
                        If Not IsNull(mrsPati!险类) And bln实时监控 Then
                            rsItems.Filter = "收费细目ID=" & lng父项目ID
                            If Not rsItems.EOF Then
                                rsItems!实收金额 = NVL(rsItems!实收金额, 0) + cur实收
                                rsItems.Update
                            End If
                        End If
                        
                        '费用SQL生成替换
                        cur实收 = Get实收金额(rsSQL!Sql) + cur实收
                        rsSQL!Sql = Set实收金额(rsSQL!Sql, cur实收)
                        rsSQL.Update
                    
                        rsSeek.MoveNext
                    Loop
                End If
                
                '选择要发送的医嘱自动进行校对(实际可能因为叮嘱不发送)
                If Val(.TextMatrix(i, COL_医嘱状态)) = 1 And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                    rsSQL.AddNew
                    rsSQL!类型 = 3: rsSQL!项目ID = 0: rsSQL!序号 = i
                    rsSQL!医嘱ID = Val(.TextMatrix(i, COL_ID))
                    rsSQL!Sql = "ZL_病人医嘱记录_校对(" & Val(.TextMatrix(i, COL_ID)) & ",3," & _
                        "To_Date('" & Format(.TextMatrix(i, COL_开嘱时间), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                        "NULL,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                End If
                
                
                '产生医嘱发送记录
                '-----------------------------------------------------------------------------------------
                If Val(.TextMatrix(i, COL_执行性质ID)) <> 0 Then '叮嘱不发送(给药途径，配方煎法、用法,采集方法、输血途径可能为)
                    '发送了出院,转院,死亡医嘱
                    If .TextMatrix(i, COL_诊疗类别) = "Z" _
                        And InStr(",5,6,11,", Val(.TextMatrix(i, COL_操作类型))) > 0 Then
                        mblnRefresh = True
                    End If
                    
                    '一样要产生费用NO
                    Call GetCurBillSet(strNOKey, strNO, -1, lng发送序号, mbytSendKind = EOutCharge)
                                                            
                    '是否一组医嘱的第一医嘱行:药疗的第一药品行为第一医嘱行
                    blnFirst = False
                    If InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                        If Val(.TextMatrix(i, COL_相关ID)) <> Val(.TextMatrix(i - 1, COL_相关ID)) Then
                            blnFirst = True
                        End If
                    ElseIf .TextMatrix(i, COL_诊疗类别) = "C" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                        If Val(.TextMatrix(i, COL_相关ID)) <> Val(.TextMatrix(i - 1, COL_相关ID)) Then
                            blnFirst = True '检验组合中的第一检验行
                        End If
                    ElseIf InStr(",1,2,3,4,5,", CLng(.Cell(flexcpData, i, COL_ID))) = 0 Then '排开给药途径、中药煎法、中药用法、采集方法、输血途径
                        If Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                            blnFirst = True
                        End If
                    End If
                                        
                    '发送数次:药品为剂量单位的总量,其它为次数
                    If .TextMatrix(i, COL_诊疗类别) = "7" Then
                        dbl发送数次 = Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_单量))
                    ElseIf InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                        dbl发送数次 = Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_住院包装)) * Val(.TextMatrix(i, COL_剂量系数))
                    Else
                        dbl发送数次 = Val(.TextMatrix(i, COL_总量))
                    End If
                    dbl发送数次 = Format(dbl发送数次, "0.00000")
                                                            
                    '首末时间(不能用“str分解时间”判断，因为一次性临嘱记录的是首次时间)
                    If .TextMatrix(i, COL_分解时间) <> "" Or mlngRefModld = 1 Then
                        str首次时间 = "To_Date('" & Split(str分解时间, ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                        str末次时间 = "To_Date('" & Split(str分解时间, ",")(UBound(Split(str分解时间, ","))) & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        '无法分解或为"一次性"临嘱，填为开始执行时间（74366）
                        str首次时间 = "To_Date('" & .TextMatrix(i, COL_开始时间) & "','YYYY-MM-DD HH24:MI:SS')"
                        str末次时间 = "To_Date('" & .TextMatrix(i, COL_开始时间) & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    If str分解时间 = "" Then str分解时间 = .TextMatrix(i, COL_开始时间)
                        
                    '领药号
                    str领药号 = ""
                    If mbln领药号 And InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                        If mstr领药号 = "" Then mstr领药号 = Get领药号
                        str领药号 = mstr领药号
                    End If
                                       
                    If Not gbln发送生成条形码 Then strCuvetteNumber = ""
                    rsSQL.AddNew
                    rsSQL!类型 = 5: rsSQL!项目ID = 0: rsSQL!序号 = i
                    rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                    rsSQL!NO = strNO
                    rsSQL!Sql = "ZL_病人医嘱发送_Insert(" & _
                        Val(.TextMatrix(i, COL_ID)) & "," & lng发送号 & "," & IIF(mbytSendKind = EOutCharge, 1, 2) & ",'" & strNO & "'," & _
                        lng发送序号 & "," & ZVal(dbl发送数次) & "," & str首次时间 & "," & str末次时间 & "," & strCurDate & "," & _
                        int执行状态 & "," & ZVal(.TextMatrix(i, COL_执行科室ID)) & "," & int计费状态 & "," & _
                        IIF(blnFirst, 1, 0) & ",'" & strCuvetteNumber & "','" & UserInfo.编号 & "'," & _
                        "'" & UserInfo.姓名 & "','" & str领药号 & "'," & IIF(mbytSendKind = EOutBilling, 1, "Null") & ",'" & str分解时间 & "'," & IIF(InStr(str关联药行, "," & Val(.TextMatrix(i, COL_ID)) & "'") > 0, str关联药行, "Null") & ")"
                    rsSQL.Update
                    
                    str关联药行 = "''"
                    
                    If gbln血库系统 And .TextMatrix(i, COL_诊疗类别) = "K" Then
                        rsSQL.AddNew
                        rsSQL!类型 = 9 '血库配血费用
                        rsSQL!项目ID = 0
                        rsSQL!序号 = 0
                        rsSQL!Sql = "Zl_血液配血费用_Insert(" & Val(.TextMatrix(i, COL_ID)) & ")"
                        rsSQL.Update
                    End If
                    
                    '医嘱执行计价
                    If rsExec.RecordCount > 0 Then
                        rsExec.Filter = "医嘱ID=" & Val(.TextMatrix(i, COL_ID)) & " And 发送号=" & lng发送号
                        If rsExec.RecordCount > 0 Then rsExec.MoveFirst
                        Do While Not rsExec.EOF
                            rsSQL.AddNew
                            rsSQL!类型 = 8
                            rsSQL!项目ID = 0
                            rsSQL!序号 = 0
                            rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                            rsSQL!Sql = "Zl_医嘱执行计价_Insert(" & rsExec!医嘱ID & "," & rsExec!发送号 & ",To_date('" & _
                            rsExec!要求时间 & "','yyyy-MM-dd HH24:mi:ss')," & ZVal(Val(rsExec!收费细目ID & "")) & "," & rsExec!数量 & "," & rsExec!费用性质 & ")"
                            rsSQL.Update
                            rsExec.MoveNext
                        Loop
                        rsExec.Filter = 0
                    End If
                    
                    '要发送的尚未签名的新开医嘱ID(组ID,一组中的叮嘱也会被签名)
                    If Val(.TextMatrix(i, COL_签名ID)) = 0 And Val(.TextMatrix(i, COL_医嘱状态)) = 1 Then
                        If Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                            lng组ID = Val(.TextMatrix(i, COL_相关ID))
                        Else
                            lng组ID = Val(.TextMatrix(i, COL_ID))
                        End If
                        If InStr(str医嘱IDs & ",", "," & lng组ID & ",") = 0 Then
                            str医嘱IDs = str医嘱IDs & "," & lng组ID
                        End If
                    End If
                End If
                
                '计算中药配方数
                If .Cell(flexcpData, i, COL_ID) = 3 Then '中药用法
                    int配方数 = int配方数 + 1
                End If
            End If
NextAdvice:
            '----------------------------------------
            Progress = (i - .FixedRows + 1) / (.Rows - .FixedRows) * 100
            txtPer.Text = CInt(psb.value) & "%"
            txtPer.Refresh
        Next
                
        '提示未审核项目
        If strAudit <> "" Then
            MsgBox "病人""" & mrsPati!姓名 & """以下费用项目还没有经过审批，对应的医嘱不能发送：" & vbCrLf & strAudit, vbInformation, gstrSysName
            GoTo errH
        End If
        
        '自动进行电子签名(未签名部份)
        '-----------------------------------------------------------------------------------------
        If Not gobjESign Is Nothing And CheckSign(IIF(mlng医技科室ID <> 0, 3, 1), 0, mlng医技科室ID, mlng病人科室id, 2, , gobjESign) And str医嘱IDs <> "" Then
            str医嘱IDs = Mid(str医嘱IDs, 2) '这里是组ID,返回为明细的ID
            intRule = ReadAdviceSignSource(1, mlng病人ID, mlng主页ID, str医嘱IDs, 0, False, strSource, mstr前提IDs)
            If intRule = 0 Then GoTo FuncEnd
            If strSource = "" Then
                Screen.MousePointer = 0
                MsgBox "不能读取要签名的医嘱源文。", vbInformation, gstrSysName
                GoTo FuncEnd
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng证书ID, strTimeStamp, Nothing, strTimeStampCode)
            If strSign = "" Then GoTo FuncEnd
            If strTimeStamp <> "" Then
                strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
            Else
                strTimeStamp = "NULL"
            End If
            lng签名id = zlDatabase.GetNextID("医嘱签名记录")
            rsSQL.AddNew
            rsSQL!类型 = 2: rsSQL!医嘱ID = 0: rsSQL!项目ID = 0: rsSQL!序号 = 0
            rsSQL!Sql = "zl_医嘱签名记录_Insert(" & lng签名id & ",1," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & str医嘱IDs & "'," & strTimeStamp & ",'" & UserInfo.姓名 & "','" & strTimeStampCode & "')"
            rsSQL.Update
        End If
        
        '医保管控实时监测
        If Not IsNull(mrsPati!险类) And bln实时监控 Then
            rsItems.Filter = 0
            If Not rsItems.EOF Then
                If Not gclsInsure.CheckItem(mrsPati!险类, 1, 2, rsItems) Then GoTo FuncEnd
            End If
        End If
        
        '普通发送页面都不产生配药记录（64615）
        
        str结算医嘱IDs = Mid(str结算医嘱IDs, 2)
        
        '提交病人数据
        '-----------------------------------------------------------------------------------------
        If Not CompletePatiSend(rsSQL, rsUpload, cur合计, cur记帐合计, str类别, bln划价, blnTran, lng发送号, str结算医嘱IDs) Then GoTo errH
    End With
    SendAdvice = lng发送号
    '调用外挂接口
    If CreatePlugInOK(p住院医嘱发送) Then
        On Error Resume Next
        Call gobjPlugIn.AdviceSendEnd(glngSys, p住院医嘱发送, lng发送号 & "")
        Call zlPlugInErrH(err, "AdviceSendEnd")
        On Error GoTo 0
    End If
    Call Make待执行消息(Format(curDate, "yyyy-MM-dd HH:mm:ss"))
FuncEnd:
    '删除所有已成功发送的行
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0: Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTran Then
        gcnOracle.RollbackTrans
    End If
    If err.Number <> 0 Then
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0
End Function

Private Function CompletePatiSend(rsSQL As ADODB.Recordset, _
    rsUpload As ADODB.Recordset, ByVal cur合计 As Currency, ByVal cur记帐合计 As Currency, ByVal str类别 As String, _
    ByVal bln划价 As Boolean, blnTran As Boolean, ByVal lng发送号 As Long, ByVal str结算医嘱IDs As String) As Boolean
'功能：提交一个病人的医嘱发送数据,在这之前处理记帐报警
'参数：
'      bln划价=是否全部费用都是划价模式，用于报警的特殊处理
'      rsSQL=包含所有要执行的SQL
'      rsUpload=用于医保上传的记帐单据号
'      cur合计=病人本次要发送医嘱的记帐金额合计,用于记帐报警
'      cur记帐合计=病人本次要发送医嘱的记帐金额合计，包括本科执行后自动审核的划价费用，不含其它划价费用
'      str类别=病人本次发送记帐费用的收费类别,用于记帐报警
'      lng发送号=本次发送的主关键字
'      str结算医嘱IDs=一卡通结算的医嘱ID串
'说明：如果出错,则在调用函数中处理,blnTran返回是否启用了事务
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intR As Integer, lng组ID As Long, str医嘱IDs As String
    Dim cur当日 As Currency, i As Long, j As Long
    Dim arrNOs() As String, strMsg As String, cur余额 As Currency
    Dim strAllmsg As String, strDiag As String, strAdviceInfo As String
    Dim arrSQL As Variant, arrAdviceID As Variant
    Dim strErr As String
    Dim str当前病情 As String
    Dim blnClearPatiCache As Boolean
    Dim blnPlugIn As Boolean, int场合 As Integer
    Dim rsAdviceRis As ADODB.Recordset
    Dim strAdvices输血 As String
    Dim var输血 As Variant
    
    int场合 = IIF("" = mstr前提IDs, 0, 2)
'    调用外挂接口发送前检查医嘱费用
    If CreatePlugInOK(p住院医嘱发送, int场合) Then
        blnPlugIn = True
        On Error Resume Next
        blnPlugIn = gobjPlugIn.AdviceCheckSendFee(glngSys, p住院医嘱发送, mlng病人ID, mlng主页ID, cur合计, int场合)
        If Not blnPlugIn And err.Number <> 0 Then blnPlugIn = True
        Call zlPlugInErrH(err, "AdviceCheckSendFee")
        err.Clear: On Error GoTo 0
        If Not blnPlugIn Then
            Exit Function
        End If
    End If
    
    '病人费用报警
    blnClearPatiCache = True
    If Not (mbytSendKind = EOutCharge) And cur合计 > 0 Then
        If InitObjPublicExpense Then
            For i = 1 To Len(str类别)
                Call gobjPublicExpense.zlBillingWarn.zlBillingWarnCheck(Me, IIF(mbytSendKind = EInBilling, 1, 0), IIF(bln划价, 1, 0), mlng病人ID, IIF(mlng病人性质 = 1, 0, mlng主页ID), mlng病人病区ID, Mid(str类别, i, 1), IIF(gbln报警包含划价费用, cur合计, cur记帐合计), InStr(";" & GetInsidePrivs(p住院医嘱下达) & ";", ";欠费强制记帐;") > 0, False, blnClearPatiCache, intR, , , , True)
                blnClearPatiCache = False
                If InStr(",2,3,", intR) > 0 Then Exit Function
            Next
        End If
    End If
    
    If mbytSendKind = EOutBilling And gdbl预存款消费验卡 <> 0 And cur记帐合计 > 0 Then
        If Not zlDatabase.PatiIdentify(Me, glngSys, mlng病人ID, cur记帐合计, , , , IIF(-1 * gdbl预存款消费验卡 >= Val(cur记帐合计), False, True), , , (gdbl预存款消费验卡 <> 0), (2 = gdbl预存款消费验卡)) Then Exit Function
    End If
    
    Call InitObjLis(p住院医生站)
    '调用LIS申请接口
    If Not gobjLIS Is Nothing Then
        strAdviceInfo = Get检验医嘱信息
        If strAdviceInfo <> "" Then
            Set rsTmp = Get病人诊断记录(mlng病人ID, mlng主页ID, "2")
            If rsTmp.RecordCount > 0 Then strDiag = rsTmp!诊断描述
        End If
    End If
    
    If gbln血库系统 Then
        If InitObjBlood(True) Then
            strAdvices输血 = Get输血医嘱信息
            If strAdvices输血 <> "" Then
                var输血 = Split(strAdvices输血, ",")
            End If
        End If
    End If
    
    Call ReplaceTrueNO(rsSQL, rsUpload)
    
    '执行顺序:1-计价,2-签名,3-校对,4-费用,5-发送,6-发料
    '1.先填写费用,因为发送时可能处理费用
    '2.对费用记录按收费细目ID排序插入
    rsSQL.Filter = 0 '上层函数可能使用过,即使没用过也MoveFirst
    rsSQL.Sort = "类型,项目ID,序号"
    rsUpload.Filter = 0 '上层函数可能使用过,即使没用过也MoveFirst

    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                If .TextMatrix(i, COL_诊疗类别) = "Z" And (Val(.TextMatrix(i, COL_操作类型)) = 9 Or Val(.TextMatrix(i, COL_操作类型)) = 10) Then
                    str当前病情 = Get病人当前病情(mlng病人ID, mlng主页ID)
                    Exit For
                End If
            End If
        Next
    End With
 
    gcnOracle.BeginTrans: blnTran = True
    Do While Not rsSQL.EOF
        Call zlDatabase.ExecuteProcedure(rsSQL!Sql, Me.Caption)
        rsSQL.MoveNext
    Loop
    
    '调用LIS申请接口
    If strAdviceInfo <> "" Then
        If gobjLIS.SendLisApplicationForm(strAdviceInfo, strDiag) = False Then
            gcnOracle.RollbackTrans: blnTran = False
            Screen.MousePointer = 0
            Call Del检验申请
            MsgBox "检验接口调用失败，不能发送检验医嘱。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
            
    '医保数据上传
    If Not IsNull(mrsPati!险类) Then
        If gclsInsure.GetCapability(support医嘱上传, mlng病人ID, mrsPati!险类) And Not gclsInsure.GetCapability(support记帐完成后上传, mlng病人ID, mrsPati!险类) Then
            Do While Not rsUpload.EOF
                strMsg = "" '因为现在一张NO内肯定为一个病人的,所以最后病人参数可以不传
                If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 1, strMsg, , mrsPati!险类, rsUpload.RecordCount & "|" & rsUpload.AbsolutePosition) Then
                    '未提交前上传失败则回滚并中止发送
                    gcnOracle.RollbackTrans: blnTran = False
                    Screen.MousePointer = 0
                    If strMsg <> "" Then
                        MsgBox strMsg, vbInformation, gstrSysName '每张提示
                    Else
                        MsgBox mrsPati!姓名 & "的费用上传失败，发送操作将被中止。", vbExclamation, gstrSysName
                    End If
                    Exit Function
                Else
                    If strMsg <> "" Then strAllmsg = strAllmsg & rsUpload!NO & ":" & strMsg & vbCrLf
                End If
                rsUpload.MoveNext
            Loop
        End If
        
        
        '医保档案上传接口(事务内以限制)
        If gclsInsure.GetCapability(support上传住院档案, mlng病人ID, mrsPati!险类) Then
            If Not gclsInsure.TranElecDossier(2, mlng病人ID, mlng主页ID, mrsPati!险类) Then Exit Function
        End If
    End If
    If strAdvices输血 <> "" Then
        For i = 0 To UBound(var输血)
            If gobjPublicBlood.AdviceOperation(p住院医嘱下达, Val(var输血(i)), 5, False, strErr) = False Then
                gcnOracle.RollbackTrans: blnTran = False
                Screen.MousePointer = 0
                MsgBox "血库系统接口调用失败：" & strErr, vbInformation, gstrSysName
                Exit Function
            End If
        Next
    End If
    gcnOracle.CommitTrans: blnTran = False
    Screen.MousePointer = 0
    If strAllmsg <> "" Then
        MsgBox strAllmsg, vbInformation, gstrSysName
    End If
    
    '一卡通结算(发送完成后调用结算，结算成功后再调用执行，取消结算或结算失败，不调执行)
    If str结算医嘱IDs <> "" Then
        If gobjSquareCard.zlSquareAffirm(Me, p住院医嘱下达, GetInsidePrivs(p住院医嘱下达), mlng病人ID, 0, False, , , str结算医嘱IDs, , , mbln使用预交) Then
            arrSQL = Array()
            arrAdviceID = Split(str结算医嘱IDs, ",")
            
            For i = 0 To UBound(arrAdviceID)
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人医嘱执行_Finish(" & arrAdviceID(i) & "," & lng发送号 & ",Null,0,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & IIF(mlng医技科室ID <> 0, mlng医技科室ID, mlng病人科室id) & ")"
            Next
                            
            gcnOracle.BeginTrans: blnTran = True
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
            Next
            gcnOracle.CommitTrans: blnTran = False
        End If
    End If
    
    '医保数据上传
    If Not IsNull(mrsPati!险类) Then
        If gclsInsure.GetCapability(support医嘱上传, mlng病人ID, mrsPati!险类) And gclsInsure.GetCapability(support记帐完成后上传, mlng病人ID, mrsPati!险类) Then
            Do While Not rsUpload.EOF
                strMsg = ""
                If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 1, strMsg, , mrsPati!险类, rsUpload.RecordCount & "|" & rsUpload.AbsolutePosition) Then
                    '提交后上传失败,仅提示
                    If strMsg <> "" Then
                        MsgBox strMsg, vbInformation, gstrSysName
                    Else
                        MsgBox mrsPati!姓名 & "的记帐单""" & rsUpload!NO & """上传失败，HIS端数据已提交，按确定继续发送。", vbExclamation, gstrSysName
                    End If
                Else
                    If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                End If
                rsUpload.MoveNext
            Loop
        End If
    End If
 
    Call SendMsg发送(lng发送号, IIF(mlng病人性质 = 1, 1, 2), IIF(bln划价, 1, 2), str当前病情)
    
    'RIS接口
    If HaveRIS Then
        If GetAdviceRis(rsAdviceRis) Then
            On Error Resume Next
            If gobjRis.HISSendAdvice(rsAdviceRis, 2, mlng病人ID, mlng主页ID, "", lng发送号) <> 1 Then
                MsgBox "当前启用了影像信息系统接口， 但由于影像信息系统接口(HISSendAdvice)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
            End If
            err.Clear: On Error GoTo 0
        End If
    ElseIf gbln启用影像信息系统接口 = True Then
        MsgBox "当前启用了影像信息系统接口， 但于由RIS接口创建失败未调用(HISSendAdvice)接口，请与系统管理员联系。", vbInformation, gstrSysName
    End If
    
    '提交成功,将病人医嘱行标记为可删除
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                .RowData(i) = -1
            End If
        Next
        
        '调用外挂接口
        If CreatePlugInOK(p住院医嘱发送) Then
            On Error Resume Next
            Call gobjPlugIn.AdviceSend(glngSys, p住院医嘱发送, mlng病人ID, mlng主页ID, lng发送号)
            Call zlPlugInErrH(err, "AdviceSend")
            On Error GoTo 0
        End If
        If gobjExchange Is Nothing Then
            On Error Resume Next
            Set gobjExchange = CreateObject("zlExchange.clsExchange")
            If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
            err.Clear: On Error GoTo 0
        End If
        '调用数据交换平台，向LIS,PACS传递申请单
        If Not gobjExchange Is Nothing Then
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                    'c-检验,d-检查
                    If .TextMatrix(i, COL_诊疗类别) = "C" Or .TextMatrix(i, COL_诊疗类别) = "D" Then
                        If Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                            lng组ID = Val(.TextMatrix(i, COL_相关ID))
                        Else
                            lng组ID = Val(.TextMatrix(i, COL_ID))
                        End If
                        If InStr(str医嘱IDs & ",", "," & lng组ID & ",") = 0 Then
                            str医嘱IDs = str医嘱IDs & "," & lng组ID
                            Call gobjExchange.SendMsg(IIF(.TextMatrix(i, COL_诊疗类别) = "C", 1, 2), "病人ID::" & mlng病人ID & "||主页ID::" & mlng主页ID & "||医嘱ID::" & lng组ID & "||操作类型::1")
                        End If
                    End If
                End If
            Next
        End If
    End With
        
    CompletePatiSend = True
End Function

Private Sub SendMsg发送(ByVal lng发送号 As Long, ByVal int病人性质 As Integer, ByVal int单据性质 As Integer, ByVal str当前病情 As String)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strIDs As String
    Dim lngTmp As Long
    Dim strTmp1 As String
    Dim strTmp2 As String
    Dim str开嘱科室 As String
    Dim i As Long
    Dim j As Long
    On Error GoTo errH
    strSQL = "select 名称 from 部门表 where id=[1]"
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                '申请安排
                If Val(.TextMatrix(i, COL_执行安排)) = 1 Then
                    Call ZLHIS_CIS_004(mclsMipModule, mlng病人ID, mstr姓名, mstr住院号, , int病人性质, _
                        mlng主页ID, .TextMatrix(i, COL_病人科室ID), "", , mstr床号, Val(.TextMatrix(i, COL_ID)), 1, .TextMatrix(i, COL_诊疗类别), .TextMatrix(i, COL_操作类型), _
                        lng发送号, .TextMatrix(i, COL_执行科室ID))
                End If
                '检验医嘱
                If .TextMatrix(i, COL_诊疗类别) = "E" And Val(.TextMatrix(i, COL_操作类型)) = 6 Then
                    strIDs = ""
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, COL_相关ID)) <> Val(.TextMatrix(i, COL_ID)) Then
                            Exit For
                        Else
                            If .TextMatrix(j, COL_诊疗类别) = "C" Then
                                strIDs = strIDs & "," & Val(.TextMatrix(j, COL_ID))
                                lngTmp = Val(.TextMatrix(j, COL_执行科室ID))
                            End If
                        End If
                    Next
                    strIDs = Mid(strIDs, 2)
                    If strIDs <> "" Then
                        Call ZLHIS_CIS_016(mclsMipModule, mlng病人ID, mstr姓名, mstr住院号, , int病人性质, mlng主页ID, mlng病人科室id, , Val(.TextMatrix(i, COL_ID)), _
                            .TextMatrix(i, COL_标本部位), .TextMatrix(i, COL_诊疗项目ID), , .TextMatrix(i, COL_执行科室ID), , strIDs, , lngTmp, , lng发送号, "", _
                            int单据性质, .TextMatrix(i, COL_开嘱医生), .TextMatrix(i, COL_开嘱时间), .TextMatrix(i, COL_开嘱科室ID), , "")
                    End If
                '检查申请
                ElseIf .TextMatrix(i, COL_诊疗类别) = "D" And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                    strTmp1 = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, COL_相关ID)) <> Val(.TextMatrix(i, COL_ID)) Then
                            Exit For
                        Else
                            If .TextMatrix(j, COL_诊疗类别) = "D" Then
                                strTmp1 = strTmp1 & "," & .TextMatrix(j, COL_标本部位)
                            End If
                        End If
                    Next
                    strTmp1 = Mid(strTmp1, 2)
                    Call ZLHIS_CIS_017(mclsMipModule, mlng病人ID, mstr姓名, mstr住院号, , int病人性质, mlng主页ID, Val(.TextMatrix(i, COL_病人科室ID)), "", Val(.TextMatrix(i, COL_ID)), _
                        .TextMatrix(i, COL_诊疗项目ID), .TextMatrix(i, col_医嘱内容), strTmp1, .TextMatrix(i, COL_执行科室ID), , lng发送号, _
                        "", int单据性质, .TextMatrix(i, COL_开嘱医生), .TextMatrix(i, COL_开嘱时间), .TextMatrix(i, COL_开嘱科室ID), , "")
                '手术申请
                ElseIf .TextMatrix(i, COL_诊疗类别) = "F" And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                    strTmp1 = Get医嘱附项内容(Val(.TextMatrix(i, COL_ID)), "主刀医生")
                    strTmp2 = Get医嘱附项内容(Val(.TextMatrix(i, COL_ID)), "助手医生")
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, COL_相关ID)) <> Val(.TextMatrix(i, COL_ID)) Then
                            Exit For
                        Else
                            If .TextMatrix(j, COL_诊疗类别) = "F" Then
                                strIDs = strIDs & "," & .TextMatrix(j, COL_ID)
                            ElseIf .TextMatrix(j, COL_诊疗类别) = "G" Then
                                lngTmp = Val(.TextMatrix(j, COL_ID))
                            End If
                        End If
                    Next
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_开嘱科室ID)))
                    If Not rsTmp.EOF Then str开嘱科室 = rsTmp!名称 & ""
                    strIDs = Mid(strIDs, 2)
                    Call ZLHIS_CIS_018(mclsMipModule, mlng病人ID, mstr姓名, mstr住院号, , int病人性质, _
                        mlng主页ID, mlng病人科室id, "", Val(.TextMatrix(i, COL_ID)), strIDs, , lngTmp, , strTmp1, strTmp2, .TextMatrix(i, COL_执行科室ID), .TextMatrix(i, COL_执行科室), lng发送号, _
                        "", int单据性质, .TextMatrix(i, COL_开嘱医生), .TextMatrix(i, COL_开嘱时间), .TextMatrix(i, COL_开嘱科室ID), str开嘱科室)
                '输血申请
                ElseIf .TextMatrix(i, COL_诊疗类别) = "K" Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_开嘱科室ID)))
                    If Not rsTmp.EOF Then str开嘱科室 = rsTmp!名称 & ""
                    Call ZLHIS_CIS_019(mclsMipModule, mlng病人ID, mstr姓名, mstr住院号, , int病人性质, _
                        mlng主页ID, mlng病人科室id, "", Val(.TextMatrix(i, COL_ID)), .TextMatrix(i, COL_执行科室ID), .TextMatrix(i, COL_执行科室), lng发送号, _
                        "", int单据性质, .TextMatrix(i, COL_开嘱医生), .TextMatrix(i, COL_开嘱时间), .TextMatrix(i, COL_开嘱科室ID), str开嘱科室)
                ElseIf .TextMatrix(i, COL_诊疗类别) = "Z" And InStr(",7,8,11,", "," & .TextMatrix(i, COL_操作类型) & ",") > 0 _
                    Or .TextMatrix(i, COL_诊疗类别) = "E" And Val(.TextMatrix(i, COL_操作类型)) = 5 Then
                    If .TextMatrix(i, COL_操作类型) = "7" Then
                        strTmp1 = "ZLHIS_CIS_020" '会诊申请
                    ElseIf .TextMatrix(i, COL_操作类型) = "8" Then
                        strTmp1 = "ZLHIS_CIS_021"  '患者抢救医嘱
                    ElseIf .TextMatrix(i, COL_操作类型) = "11" Then
                        strTmp1 = "ZLHIS_CIS_022" '患者死亡医嘱
                    ElseIf .TextMatrix(i, COL_诊疗类别) = "E" And Val(.TextMatrix(i, COL_操作类型)) = 5 Then
                        strTmp1 = "ZLHIS_CIS_023"  '患者特殊治疗医嘱
                    End If
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_开嘱科室ID)))
                    If Not rsTmp.EOF Then str开嘱科室 = rsTmp!名称 & ""
                    Call SendMsg(strTmp1, mclsMipModule, mlng病人ID, mstr姓名, mstr住院号, , int病人性质, _
                        mlng主页ID, Val(.TextMatrix(i, COL_病人科室ID)), "", Val(.TextMatrix(i, COL_ID)), .TextMatrix(i, COL_执行科室ID), .TextMatrix(i, COL_执行科室), lng发送号, _
                        "", int单据性质, .TextMatrix(i, COL_开嘱医生), .TextMatrix(i, COL_开嘱时间), .TextMatrix(i, COL_开嘱科室ID), str开嘱科室)
                
                '住院患者预出院
                ElseIf .TextMatrix(i, COL_诊疗类别) = "Z" And Val(.TextMatrix(i, COL_操作类型)) = 5 Then
                    Call GetPatChange(Val(.TextMatrix(i, COL_ID)), 10, lngTmp, strTmp1)
                    Call ZLHIS_PATIENT_009(mclsMipModule, mlng病人ID, mlng主页ID, mstr姓名, mstr性别, mstr住院号, _
                        lngTmp, .TextMatrix(i, COL_开嘱时间), mlng病人病区ID, , mlng病人科室id, "", , mstr床号, Val(.TextMatrix(i, COL_ID)))
               
                '住院患者病情变更
                ElseIf .TextMatrix(i, COL_诊疗类别) = "Z" And (Val(.TextMatrix(i, COL_操作类型)) = 9 Or Val(.TextMatrix(i, COL_操作类型)) = 10) Then
                    Call GetPatChange(Val(.TextMatrix(i, COL_ID)), 13, lngTmp, strTmp1)
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人科室id)
                    strTmp2 = "": If Not rsTmp.EOF Then strTmp2 = rsTmp!名称 & ""
                    Call ZLHIS_PATIENT_005(mclsMipModule, mlng病人ID, mlng主页ID, mstr姓名, mstr性别, mstr住院号, _
                        mlng病人病区ID, , mlng病人科室id, strTmp2, str当前病情, lngTmp, .TextMatrix(i, COL_开嘱时间), strTmp1, .TextMatrix(i, COL_开嘱医生), Val(.TextMatrix(i, COL_ID)))
                 
                '住院患者转出科室
                ElseIf .TextMatrix(i, COL_诊疗类别) = "Z" And Val(.TextMatrix(i, COL_操作类型)) = 3 Then
                    Call GetPatChange(Val(.TextMatrix(i, COL_ID)), 3, lngTmp, strTmp1)
                    Call ZLHIS_PATIENT_003(mclsMipModule, mlng病人ID, mlng主页ID, mstr姓名, mstr性别, mstr住院号, _
                        mlng病人病区ID, , mlng病人科室id, "", , mstr住院号, _
                        lngTmp, .TextMatrix(i, COL_开嘱时间), , , Val(.TextMatrix(i, COL_执行科室ID)), , Val(.TextMatrix(i, COL_ID)))
                End If
            End If
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowSendTotal()
'功能：根据当前选择要发送的医嘱，计算并显示发送的医嘱合计
    Dim cur金额 As Currency, cur药品金额 As Currency, i As Long
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                '可见行的金额:是一组的汇总金额
                If Not .RowHidden(i) Then
                    cur金额 = cur金额 + Val(.TextMatrix(i, COL_金额))
                End If
                '药品的金额,取原始金额
                If InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                    cur药品金额 = cur药品金额 + Val(.Cell(flexcpData, i, COL_金额))
                End If
            End If
        Next
    End With
    stbThis.Panels(4).Text = "金额:" & FormatEx(cur金额, gbytDec) & "(药" & FormatEx(cur药品金额, gbytDec) & ")"
    Call Form_Resize
End Sub


Private Function GetICUDeptID() As Long
'功能：获取当前医生所在的ICU部门ID（不含临床部门时）
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Long, arrTmp As Variant
    
    strSQL = "Select a.部门id, b.工作性质" & vbNewLine & _
            "From 部门人员 A, 部门性质说明 B" & vbNewLine & _
            "Where a.人员id = [1] And a.部门id = b.部门id And b.工作性质 In ('ICU', '临床')"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    If rsTmp.RecordCount > 0 Then
        rsTmp.Filter = "工作性质='临床'"
        If rsTmp.RecordCount = 0 Then
            rsTmp.Filter = "工作性质='ICU'"
                If rsTmp.RecordCount > 0 Then
                    GetICUDeptID = Val(rsTmp!部门ID)
                End If
        Else
            strSQL = ""
            For i = 1 To rsTmp.RecordCount
                strSQL = strSQL & "," & rsTmp!部门ID
                rsTmp.MoveNext
            Next
            arrTmp = Split(Mid(strSQL, 2), ",")
            For i = 0 To UBound(arrTmp)
                rsTmp.Filter = "工作性质='ICU' And 部门ID=" & arrTmp(i)
                If rsTmp.RecordCount = 0 Then
                    GetICUDeptID = 0: Exit For
                Else
                    If i = 0 Then GetICUDeptID = arrTmp(i)
                End If
            Next
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Del检验申请()
'功能：医嘱发送失败，事务回退后，调用检验申请删除接口
    Dim i As Long, str医嘱IDs As String, strErr As String
        
    '收集采集方法
    With vsAdvice
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_操作类型)) = 6 And .TextMatrix(i, COL_诊疗类别) = "E" Then
                If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                    str医嘱IDs = str医嘱IDs & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
    End With
    
    If str医嘱IDs <> "" Then
        str医嘱IDs = Mid(str医嘱IDs, 2)
        Call InitObjLis(p住院医生站)
        If Not gobjLIS Is Nothing Then
            If gobjLIS.DelLisApplicationForm(str医嘱IDs, strErr) = False Then
                MsgBox "删除检验申请失败：" & strErr, vbInformation, gstrSysName
            End If
        End If
    End If
End Sub

Private Function Get检验医嘱信息() As String
'功能：获取检验医嘱信息，传递给检验接口程序
    Dim i As Long, strInfo As String
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COL_操作类型)) = 6 And .TextMatrix(i, COL_诊疗类别) = "E" Then
                '检验医嘱ID1,采集医嘱ID1,执行科室ID1,标本1;.....
                'LIS接口产生的检验，一个采集方式只有一条检验医嘱（没有一并采集的情况）
                If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                    strInfo = strInfo & ";" & .TextMatrix(i - 1, COL_ID) & "," & .TextMatrix(i, COL_ID) & "," & .TextMatrix(i - 1, COL_执行科室ID) & "," & .TextMatrix(i - 1, COL_标本部位)
                End If
            End If
        Next
    End With
    Get检验医嘱信息 = Mid(strInfo, 2)
End Function

Private Function Get输血医嘱信息() As String
'功能：获取输血医嘱信息，传递给接口程序，仅取主医嘱ID
    Dim i As Long, strInfo As String
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, COL_诊疗类别) = "K" Then
                If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                    strInfo = strInfo & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
    End With
    Get输血医嘱信息 = Mid(strInfo, 2)
End Function


Private Sub InitExecRecordset(rsExec As Recordset)
'功能：初始化医嘱计价记录集
    Set rsExec = New ADODB.Recordset
    
    rsExec.Fields.Append "医嘱ID", adBigInt
    rsExec.Fields.Append "发送号", adBigInt, , adFldIsNullable
    rsExec.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsExec.Fields.Append "要求时间", adDate, , adFldIsNullable
    rsExec.Fields.Append "数量", adDouble, , adFldIsNullable
    rsExec.Fields.Append "费用性质", adInteger, , adFldIsNullable
    
    rsExec.CursorLocation = adUseClient
    rsExec.LockType = adLockOptimistic
    rsExec.CursorType = adOpenStatic
    rsExec.Open
End Sub

Private Function zlPluginAdviceBeforeSend() As Boolean
'功能：医嘱发送前调用外挂号
    Dim i As Long, j As Long
    Dim strAdviceIDs As String, strMsg  As String
    Dim rsDataPlugIn As ADODB.Recordset
    Dim lng数量 As Long
    Dim str分解时间 As String, strTmp As String
    Dim int场合 As Integer
    
    zlPluginAdviceBeforeSend = True
    
    '调用外挂接口，医嘱发送前的检查
    If CreatePlugInOK(p住院医嘱发送) Then
        Call InitPlugInRs(rsDataPlugIn)
        int场合 = IIF("" = mstr前提IDs, 0, 2)
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                    If .TextMatrix(i, COL_分解时间) <> "" Then
                        str分解时间 = .TextMatrix(i, COL_分解时间)
                    Else
                        str分解时间 = .Cell(flexcpData, i, COL_分解时间)    '开始执行时间
                    End If
                    rsDataPlugIn.AddNew
                    rsDataPlugIn!病人ID = mlng病人ID
                    rsDataPlugIn!就诊ID = mlng主页ID
                    rsDataPlugIn!医嘱ID = Val(.TextMatrix(i, COL_ID))
                    rsDataPlugIn!相关ID = Val(.TextMatrix(i, COL_相关ID))
                    rsDataPlugIn!收费细目ID = Val(.TextMatrix(i, COL_收费细目ID))
                    rsDataPlugIn!分解时间 = str分解时间
                    rsDataPlugIn!次数 = Val(.TextMatrix(i, COL_次数))
                    rsDataPlugIn!单量 = Val(.TextMatrix(i, COL_单量))
                    rsDataPlugIn!单量单位 = .TextMatrix(i, COL_单量单位)
                    rsDataPlugIn!总量 = Val(.TextMatrix(i, COL_总量))
                    rsDataPlugIn!总量单位 = .TextMatrix(i, COL_总量单位)
                    rsDataPlugIn!场合 = int场合
                    rsDataPlugIn.Update
                End If
            Next
            If rsDataPlugIn.RecordCount > 0 Then rsDataPlugIn.MoveFirst
            strAdviceIDs = "": strMsg = ""
            On Error Resume Next
            Call gobjPlugIn.AdviceBeforeSend("", rsDataPlugIn, strAdviceIDs, strMsg)
            Call zlPlugInErrH(err, "AdviceBeforeSend")
            err.Clear
            On Error GoTo 0
             
            If strAdviceIDs <> "" Then
                strTmp = ""
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                        If InStr("," & strAdviceIDs & ",", "," & Val(.TextMatrix(i, COL_ID)) & ",") > 0 Then
                            If Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                                j = Val(.TextMatrix(i, COL_ID))
                            Else
                                j = Val(.TextMatrix(i, COL_相关ID))
                            End If
                            
                            If InStr("," & strTmp & ",", "," & j & ",") = 0 Then
                                strTmp = strTmp & "," & j
                            End If
                        End If
                    End If
                Next
                strAdviceIDs = Mid(strTmp, 2)
                lng数量 = 0
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                        If Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                            j = Val(.TextMatrix(i, COL_ID))
                        Else
                            j = Val(.TextMatrix(i, COL_相关ID))
                        End If
                        lng数量 = lng数量 + 1
                        If InStr("," & strAdviceIDs & ",", "," & j & ",") > 0 Then
                            .Cell(flexcpData, i, COL_选择) = 1
                            Set .Cell(flexcpPicture, i, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                            lng数量 = lng数量 - 1
                        End If
                    End If
                Next
                
                If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                If lng数量 = 0 Then
                    MsgBox "当前没有可以发送的医嘱。", vbInformation, gstrSysName
                    zlPluginAdviceBeforeSend = False
                End If
            End If
        End With
    End If
End Function
 
Private Function GetAdviceRis(ByRef rsData As ADODB.Recordset) As Boolean
'功能：获取发送到RIS的医嘱信息
    Dim i As Long
    
    On Error GoTo errH
    
    Set rsData = New ADODB.Recordset
    
    rsData.Fields.Append "医嘱ID", adBigInt
    rsData.Fields.Append "开嘱科室ID", adBigInt
    rsData.Fields.Append "执行科室ID", adBigInt
    rsData.Fields.Append "诊疗项目ID", adBigInt
    rsData.Fields.Append "病人来源", adInteger '1-门诊;2-住院;
    rsData.Fields.Append "类别", adVarChar, 10
    rsData.CursorLocation = adUseClient
    rsData.LockType = adLockOptimistic
    rsData.CursorType = adOpenStatic
    rsData.Open
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                If InStr(",D,F,", .TextMatrix(i, COL_诊疗类别)) > 0 Or InStr(",0,5,", Val(.TextMatrix(i, COL_操作类型))) > 0 And .TextMatrix(i, COL_诊疗类别) = "E" Then
                    If Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                        rsData.AddNew
                        rsData!医嘱ID = Val(.TextMatrix(i, COL_ID))
                        rsData!开嘱科室id = Val(.TextMatrix(i, COL_开嘱科室ID))
                        rsData!执行科室ID = Val(.TextMatrix(i, COL_执行科室ID))
                        rsData!诊疗项目ID = Val(.TextMatrix(i, COL_诊疗项目ID))
                        rsData!病人来源 = 2
                        rsData!类别 = .TextMatrix(i, COL_诊疗类别)
                        rsData.Update
                    End If
                End If
            End If
        Next
    End With
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
        GetAdviceRis = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckRISScheduling() As Boolean
'功能：检查项目是否是必须预约
    Dim i As Long
    Dim blnDo As Boolean
    Dim lng医嘱ID As Long
    Dim lng诊疗项目ID As Long
    Dim lngRst As Long
    Dim strMsg As String
    
    CheckRISScheduling = True
    
    If HaveRIS Then
        If gbln启用影像信息系统预约 Then
            blnDo = True
        End If
    End If
    
    If Not blnDo Then Exit Function
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                If Val(.TextMatrix(i, COL_紧急标志)) <> 1 Then
                    If InStr(",D,F,", .TextMatrix(i, COL_诊疗类别)) > 0 Or InStr(",0,5,", Val(.TextMatrix(i, COL_操作类型))) > 0 And .TextMatrix(i, COL_诊疗类别) = "E" Then
                        If Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                            lng医嘱ID = Val(.TextMatrix(i, COL_ID))
                            lng诊疗项目ID = Val(.TextMatrix(i, COL_诊疗项目ID))
                            lngRst = -1
                            lngRst = gobjRis.HISScheduling(2, lng医嘱ID, lng诊疗项目ID, False)
                            If lngRst <> 0 Then
                            '接口返回失败给出提示
                                .Cell(flexcpData, i, COL_选择) = 1 '当前禁止选择
                                Set .Cell(flexcpPicture, i, COL_选择) = frmIcons.imgTrueFalse.ListImages("F").Picture
                                Call RowSelectSame(i, COL_选择)
                                strMsg = IIF("" = strMsg, "", strMsg & "、") & .TextMatrix(i, col_医嘱内容)
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End With
    
    If strMsg <> "" Then
        MsgBox "由于启用了RIS系统预约流程：" & vbCrLf & "【" & strMsg & "】" & _
                vbCrLf & "医嘱没有预约，预约成功后才能发送。", vbInformation, gstrSysName
        CheckRISScheduling = False
    End If
End Function

Private Function Set阳性用药() As Boolean
'功能：设置药品医嘱行的阳性用药说明
    Dim i As Long
    Dim strMsg As String
    Dim str阳性用药 As String
    Dim strSQL As String
    Dim str医嘱IDs As String
    
    On Error GoTo errH
    If mstrAdDrugIDs = "" Then
        Set阳性用药 = True
        Exit Function
    End If
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                If InStr("," & mstrAdDrugIDs & ",", "," & .TextMatrix(i, COL_ID) & ",") > 0 Then
                    strMsg = strMsg & "," & .TextMatrix(i, col_医嘱内容)
                    str医嘱IDs = str医嘱IDs & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
    End With
    If strMsg = "" Then
        Set阳性用药 = True
        Exit Function
    End If
    Call frmMsgDruExcess.ShowMe(Me, 1, Mid(strMsg, 2), str阳性用药)
    If str阳性用药 = "*NULL*" Then
        Exit Function
    End If
    strSQL = "Zl_病人医嘱记录_阳性用药('" & Mid(str医嘱IDs, 2) & "','" & str阳性用药 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Set阳性用药 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Check处方审查()
'功能：调用审方接口判断当前医嘱是不是允许发送
    Dim i As Long
    Dim str给药IDs As String '传入到接口中的参数
    Dim strOut医嘱IDs As String '不能够发送的主医嘱ID
    Dim str药行医嘱IDs As String
    Dim strErr As String
    Dim lng医嘱ID As Long
    Dim str医嘱内容 As String
    Dim rsTmp As ADODB.Recordset
    Dim blnTmp As Boolean
    
    On Error GoTo errH
    
    If Not gbln审方系统 Then Exit Sub
    
    With vsAdvice
        '调用合理用部件中提供的接口
        blnTmp = gobjPass.ZLPharmReviewResultView(mlng病人ID, mlng主页ID, rsTmp, strErr)
        '接口没出错
        If blnTmp Then
            If strErr = "" Then
                If Not rsTmp Is Nothing Then
                    If Not rsTmp.EOF Then
                        For i = 1 To rsTmp.RecordCount
                            If InStr("," & strOut医嘱IDs & ",", "," & rsTmp!相关ID & ",") = 0 Then
                                strOut医嘱IDs = strOut医嘱IDs & "," & rsTmp!相关ID
                            End If
                            str药行医嘱IDs = str药行医嘱IDs & "," & rsTmp!医嘱ID
                            rsTmp.MoveNext
                        Next
                    End If
                End If
            End If
        End If
            
        If strOut医嘱IDs <> "" Then
            '取消选择
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                    If Val(.TextMatrix(i, COL_医嘱状态)) = 1 Then
                        lng医嘱ID = IIF(0 = Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                        If InStr("," & strOut医嘱IDs & ",", "," & lng医嘱ID & ",") > 0 Then
                            Set .Cell(flexcpPicture, i, COL_选择) = Nothing '缺省不发送
                            If Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                If InStr("," & str药行医嘱IDs & ",", "," & Val(.TextMatrix(i, COL_ID)) & ",") > 0 Then
                                    str医嘱内容 = str医嘱内容 & vbCrLf & .TextMatrix(i, col_医嘱内容)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
        If str医嘱内容 <> "" Then
            Call MsgBox("以下医嘱未通过处方审查，不能发送：" & str医嘱内容, vbInformation, Me.Caption)
        End If
            
       
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
