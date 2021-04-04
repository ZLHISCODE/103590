VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmAdviceSendDrug 
   AutoRedraw      =   -1  'True
   Caption         =   "药疗医嘱发送"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   Icon            =   "frmAdviceSendDrug.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmAdviceSendDrug.frx":058A
   ScaleHeight     =   6510
   ScaleWidth      =   9540
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   7290
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "100%"
      Top             =   6255
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.ProgressBar psb 
      Height          =   270
      Left            =   2115
      TabIndex        =   5
      Top             =   6210
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
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
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAdviceSendDrug.frx":0B14
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11430
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
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
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   900
      BandCount       =   1
      _CBWidth        =   9540
      _CBHeight       =   510
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   450
      Width1          =   3525
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   450
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   794
         ButtonWidth     =   1561
         ButtonHeight    =   794
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "全选"
               Key             =   "全选"
               Description     =   "全选"
               Object.ToolTipText     =   "全选(Ctrl+A)"
               Object.Tag             =   "全选"
               ImageKey        =   "全选"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "全清"
               Key             =   "全清"
               Description     =   "全清"
               Object.ToolTipText     =   "全清(Ctrl+R)"
               Object.Tag             =   "全清"
               ImageKey        =   "全清"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "发送"
               Key             =   "发送"
               Description     =   "发送"
               Object.ToolTipText     =   "发送选择的医嘱(Ctrl+E)"
               Object.Tag             =   "发送"
               ImageKey        =   "发送"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "重置"
               Key             =   "重置"
               Description     =   "重置"
               Object.ToolTipText     =   "重新设置条件并产生发送清单(F12)"
               Object.Tag             =   "重置"
               ImageKey        =   "重置"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Description     =   "帮助"
               Object.ToolTipText     =   "帮助(F1)"
               Object.Tag             =   "帮助"
               ImageKey        =   "帮助"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Description     =   "退出"
               Object.ToolTipText     =   "退出(ALT+X)"
               Object.Tag             =   "退出"
               ImageKey        =   "退出"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraUD 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   8
      Top             =   4605
      Width           =   9495
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrice 
      Height          =   1425
      Left            =   0
      TabIndex        =   1
      Top             =   4725
      Width           =   9540
      _cx             =   16828
      _cy             =   2514
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
      FormatString    =   $"frmAdviceSendDrug.frx":13A8
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
   Begin VB.Frame fraInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   60
      TabIndex        =   6
      Top             =   525
      Width           =   9435
      Begin VB.Label lblInfo 
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
      Height          =   3765
      Left            =   0
      TabIndex        =   0
      Top             =   825
      Width           =   9540
      _cx             =   16828
      _cy             =   6641
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
      BackColorSel    =   16764057
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
      FormatString    =   $"frmAdviceSendDrug.frx":1443
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
      Editable        =   2
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
      Begin MSComctlLib.ImageList img16 
         Left            =   3435
         Top             =   1905
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceSendDrug.frx":14DE
               Key             =   "T"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceSendDrug.frx":1A78
               Key             =   "F"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   360
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":2012
            Key             =   "全选"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":222C
            Key             =   "全清"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":2446
            Key             =   "发送"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":2660
            Key             =   "重置"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":287A
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":2A94
            Key             =   "退出"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   960
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":2CAE
            Key             =   "全选"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":2EC8
            Key             =   "全清"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":30E2
            Key             =   "发送"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":32FC
            Key             =   "重置"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":3516
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":3730
            Key             =   "退出"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAdviceSendDrug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String 'IN
Public mlng病区ID As Long 'IN:用于记录主界面的病区及上次发送病区
Public mlng病人ID As Long 'IN
Public mblnSend As Boolean 'OUT:是否成功发送过。

Private mcolStock As Collection '存放各个药品库房的出库检查方式
Private mrs药房 As ADODB.Recordset
Private mrsBill As ADODB.Recordset
Private mstrEnd As String '本次发送的结束时间
Private mint期效 As Integer '本次发送的医嘱期效
Private mblnLimit As Boolean '本次发送给药途径计算是否以结束时间限制
Private mlng药品类别ID As Long '药品入出类别ID
Private mlng卫材类别ID As Long
Private mblnAutoExe As Boolean '本科执行自动完成
Private mstrLike As String
Private mblnFirst As Boolean

'----------------------------------------------
Private Const COL_选择 = 0
Private Const COL_科室 = 1
Private Const COL_姓名 = 2
Private Const COL_住院号 = 3
Private Const COL_床号 = 4
Private Const COL_费别 = 5
Private Const COL_婴儿 = 6
Private Const COL_医嘱内容 = 7
Private Const COL_规格 = 8
Private Const COL_总量 = 9
Private Const COL_总量单位 = 10
Private Const COL_单量 = 11
Private Const COL_单量单位 = 12
Private Const COL_金额 = 13
Private Const COL_频率 = 14
Private Const COL_用法 = 15
Private Const COL_医生嘱托 = 16
Private Const COL_执行时间 = 17
Private Const COL_首次时间 = 18
Private Const COL_末次时间 = 19
Private Const COL_执行科室 = 20
Private Const COL_附加执行 = 21
Private Const COL_执行性质 = 22
Private Const COL_病人ID = 23 '隐藏列
Private Const COL_主页ID = 24
Private Const COL_性别 = 25
Private Const COL_年龄 = 26
Private Const COL_ID = 27
Private Const COL_相关ID = 28
Private Const COL_病人科室ID = 29
Private Const COL_开嘱科室ID = 30
Private Const COL_开嘱医生 = 31
Private Const COL_诊疗类别 = 32
Private Const COL_诊疗项目ID = 33
Private Const COL_计价特性 = 34
Private Const COL_执行性质ID = 35
Private Const COL_执行科室ID = 36
Private Const COL_药品ID = 37
Private Const COL_剂量系数 = 38
Private Const COL_住院包装 = 39
Private Const COL_住院单位 = 40
Private Const COL_可否分零 = 41
Private Const COL_药房分批 = 42
Private Const COL_是否变价 = 43
Private Const COL_库存 = 44
Private Const COL_次数 = 45
Private Const COL_分解时间 = 46
'-------------------------------------------------
Private Const COLP_计价医嘱 = 0
Private Const COLP_类别 = 1
Private Const COLP_收费项目 = 2
Private Const COLP_付数 = 3
Private Const COLP_数量 = 4
Private Const COLP_单位 = 5
Private Const COLP_单价 = 6
Private Const COLP_应收金额 = 7
Private Const COLP_实收金额 = 8
Private Const COLP_执行科室 = 9
Private Const COLP_费用类型 = 10
Private Const COLP_从项 = 11
Private Const COLP_行号 = 12

Private Property Let Progress(ByVal vNewValue As Single)
'vNewValue=0-100
    If vNewValue = 0 Then
        psb.Value = 0: txtPer.Text = ""
        psb.Visible = False: txtPer.Visible = False
    Else
        psb.Value = vNewValue
        txtPer.Text = CInt(psb.Value) & "%"
        psb.Visible = True: txtPer.Visible = True
        txtPer.Refresh
    End If
End Property

Private Sub Form_Activate()
    If mblnFirst Then
        mblnFirst = False
        
        mlng药品类别ID = ExistIOClass(9)
        If mlng药品类别ID = 0 Then
            MsgBox "不能确定药品处方单据的入出类别,请先到入出类别管理中设置！", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        mlng卫材类别ID = ExistIOClass(41) '不能确定是否使用了卫材收费,后面再判断
        
        If Not ResetSend Then Unload Me: Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call tbr_ButtonClick(tbr.Buttons("帮助"))
    ElseIf KeyCode = vbKeyX And Shift = vbAltMask Then
        Call tbr_ButtonClick(tbr.Buttons("退出"))
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("全选"))
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("全清"))
    ElseIf KeyCode = vbKeyE And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("发送"))
    ElseIf KeyCode = vbKeyF12 Then
        Call tbr_ButtonClick(tbr.Buttons("重置"))
    End If
End Sub

Private Sub Form_Load()
    Call InitAdviceTable
    Call InitPriceTable
    Call RestoreWinState(Me, App.ProductName)
        
    mblnSend = False
    mblnFirst = True
    mstrLike = IIF(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
    mblnAutoExe = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "本科执行自动完成", 0)) <> 0
    
    '初始读取一些数据
    '各个库房药品出库检查方式,包括发料部门
    Set mcolStock = InitStockCheck(2, True)
    Call Init药房置换
End Sub

Private Function GetStockCheck(ByVal lng库房ID As Long) As Integer
'功能：获取指定库房的出库库存检查方式
    Dim intStyle As Integer
    On Error Resume Next
    intStyle = mcolStock("_" & lng库房ID)
    Err.Clear: On Error GoTo 0
    GetStockCheck = intStyle
End Function

Private Function Init药房置换() As Boolean
'功能：'初始读取一些数据
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    '读取可用药房到集合中:用于药房置换
    Set mrs药房 = New ADODB.Recordset
    mrs药房.Fields.Append "ID", adBigInt
    mrs药房.Fields.Append "编码", adVarChar, 10
    mrs药房.Fields.Append "名称", adVarChar, 20
    mrs药房.Fields.Append "现ID", adBigInt
    mrs药房.CursorLocation = adUseClient
    mrs药房.LockType = adLockOptimistic
    mrs药房.CursorType = adOpenStatic
    mrs药房.Open
    
    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " AND B.部门ID=A.ID And B.服务对象 IN(2,3) and B.工作性质 in('中药房','西药房','成药房')" & _
        " Order by A.编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        mrs药房.AddNew
        mrs药房!ID = rsTmp!ID
        mrs药房!编码 = rsTmp!编码
        mrs药房!名称 = rsTmp!名称
        mrs药房!现ID = rsTmp!ID
        mrs药房.Update
        rsTmp.MoveNext
    Next
    
    Init药房置换 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    fraInfo.Top = cbr.Height
    fraInfo.Left = 0
    fraInfo.Width = Me.ScaleWidth
    
    vsAdvice.Left = 0
    vsAdvice.Top = fraInfo.Top + fraInfo.Height
    vsAdvice.Width = Me.ScaleWidth
    vsAdvice.Height = Me.ScaleHeight - fraInfo.Height - vsPrice.Height - fraUD.Height - cbr.Height - stbThis.Height
    
    fraUD.Top = vsAdvice.Top + vsAdvice.Height
    fraUD.Left = 0
    fraUD.Width = Me.ScaleWidth
    
    vsPrice.Left = 0
    vsPrice.Top = fraUD.Top + fraUD.Height
    vsPrice.Width = Me.ScaleWidth
    
    psb.Top = stbThis.Top + 60
    psb.Width = stbThis.Panels(2).Width - txtPer.Width - 100
    psb.Left = stbThis.Panels(2).Left + 30
    
    txtPer.Left = psb.Left + psb.Width
    txtPer.Top = psb.Top + (psb.Height - txtPer.Height) / 2
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    '释放私有及IN变量
    mstrPrivs = ""
    mlng病区ID = 0
    mlng病人ID = 0
    mstrEnd = ""
    mint期效 = 0
    mblnLimit = False
    mlng药品类别ID = 0
    mlng卫材类别ID = 0
    Set mrs药房 = Nothing
    Set mrsBill = Nothing
    Set mcolStock = Nothing
    
    gbln加班加价 = False
End Sub

Private Function ResetSend() As Boolean
'功能：重置发送条件
    With frmAdviceSendDrugCond
        .mstrPrivs = mstrPrivs
        .mlng病区ID = mlng病区ID
        .mlng病人ID = mlng病人ID
        Set .mrs药房 = mrs药房
        .Show 1, Me
        If .mblnOK Then
            mlng病区ID = .mlng病区ID
            mstrEnd = .mstrEnd
            mint期效 = .mint期效
            mblnLimit = .mblnLimit
            Call LoadAdviceSend(.mstrEnd, .mint期效, .mlng药房ID, .mstr病人IDs, .mstr给药IDs)
        End If
        ResetSend = .mblnOK
    End With
End Function

Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsAdvice.Height + y < 1000 Or vsPrice.Height - y < 500 Then Exit Sub
        fraUD.Top = fraUD.Top + y
        vsAdvice.Height = vsAdvice.Height + y
        vsPrice.Top = vsPrice.Top + y
        vsPrice.Height = vsPrice.Height - y
        Me.Refresh
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lng发送号 As Long, i As Long
    
    Select Case Button.Key
        Case "全选"
            With vsAdvice
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_选择) = 0 Then
                        Set .Cell(flexcpPicture, i, COL_选择) = img16.ListImages("T").Picture
                    End If
                Next
            End With
            Call ShowSendTotal
        Case "全清"
            With vsAdvice
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_选择) = 0 Then
                        Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                    End If
                Next
            End With
            Call ShowSendTotal
        Case "发送"
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
            If MsgBox("确实要发送当前选择的医嘱吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                lng发送号 = SendAdvice
                If lng发送号 <> 0 Then
                    mblnSend = True
                    '打印诊疗单据
                    Call frmSendBillPrint.ShowMe(lng发送号, 2, Me)
                End If
            End If
        Case "重置"
            Call ResetSend
        Case "帮助"
            ShowHelp App.ProductName, Me.Hwnd, Me.Name
        Case "退出"
            Unload Me
    End Select
End Sub

Private Sub RowSelectSame(ByVal lngRow As Long, ByVal lngCol As Long, _
    Optional rsSQL As ADODB.Recordset, Optional rsTotal As ADODB.Recordset, Optional rsUpload As ADODB.Recordset)
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

Private Function ShowSendPrice(ByVal lngRow As Long) As Boolean
'功能：显示当前发送医嘱行的记帐费用信息(以住院单位,按费别打折；整个医嘱，可能多行)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngTmp As Long, strTmp As String
    Dim str费别 As String, str行号 As String, dbl数量 As Double
    Dim dbl单价 As Double, cur应收 As Currency, cur实收 As Currency
    Dim dbl当前单价 As Double, cur当前应收 As Currency, cur当前实收 As Currency
    Dim lng执行科室ID As Long, lng病人科室ID As Long
    Dim lng病人ID As Long, lng主页ID As Long
    
    Dim rsMain As New ADODB.Recordset
    Dim rsClone As New ADODB.Recordset
    Dim strHaveSub As String, strNoneSub As String
    
    On Error GoTo errH
    
    '用于汇总计算折扣的临时记录集
    rsMain.Fields.Append "医嘱行号", adBigInt
    rsMain.Fields.Append "主项行号", adBigInt
    rsMain.Fields.Append "主收入ID", adBigInt
    rsMain.Fields.Append "医嘱合计", adCurrency, , adFldIsNullable
    rsMain.CursorLocation = adUseClient
    rsMain.LockType = adLockOptimistic
    rsMain.CursorType = adOpenStatic
    rsMain.Open
    
    With vsAdvice
        lng病人ID = Val(.TextMatrix(lngRow, COL_病人ID))
        lng主页ID = Val(.TextMatrix(lngRow, COL_主页ID))
        str费别 = .TextMatrix(lngRow, COL_费别)
        
        If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            '成药品计价部分:药品缺省固定为正常计价,但下医嘱时指定了为自备药(院外执行)的不读取;药品不可能为叮嘱
            If Val(.TextMatrix(lngRow, COL_执行性质ID)) <> 5 Then
                strSQL = "Select '药品医嘱-" & Replace(.Cell(flexcpData, lngRow, COL_医嘱内容), "'", "''") & "' as 计价医嘱," & _
                    lngRow & " as 行号,A.ID,A.类别,Nvl(B.名称,A.名称)||Decode(A.产地,NULL,NULL,'('||A.产地||')')||Decode(A.规格,NULL,NULL,' '||A.规格) as 名称," & _
                    "'" & .TextMatrix(lngRow, COL_总量单位) & "' as 单位,A.是否变价,A.加班加价,0 as 单价,1 as 付数," & _
                    Val(.TextMatrix(lngRow, COL_总量)) & " as 数量," & Val(.TextMatrix(lngRow, COL_执行科室ID)) & " as 执行科室ID," & _
                    " A.屏蔽费别,A.费用类型,0 as 从项" & _
                    " From 收费项目目录 A,收费项目别名 B" & _
                    " Where A.ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIF(gbln商品名, 3, 1) & _
                    " And A.ID=" & Val(.TextMatrix(lngRow, COL_药品ID))
            End If
            '给药途径计价部分:一并给药(如果是)中的第一成药行,才显示给药途径费用
            If Val(.TextMatrix(lngRow - 1, COL_相关ID)) <> Val(.TextMatrix(lngRow, COL_相关ID)) Then
                '不计价,手工计价；叮嘱,院外执行；的医嘱不读取
                lngTmp = .FindRow(.TextMatrix(lngRow, COL_相关ID), lngRow, COL_ID)
                If Val(.TextMatrix(lngTmp, COL_计价特性)) = 0 _
                    And InStr(",0,5,", Val(.TextMatrix(lngTmp, COL_执行性质ID))) = 0 Then
                    strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                        "Select '给药途径-" & Replace(.Cell(flexcpData, lngTmp, COL_医嘱内容), "'", "''") & "' as 计价医嘱," & _
                        lngTmp & " as 行号,B.ID,B.类别,B.名称,B.计算单位 as 单位,B.是否变价,B.加班加价,A.单价,1 as 付数," & _
                        " Nvl(A.数量,0)*" & Val(.TextMatrix(lngTmp, COL_总量)) & " as 数量," & _
                        " Nvl(A.执行科室ID," & Val(.TextMatrix(lngTmp, COL_执行科室ID)) & ") as 执行科室ID," & _
                        " B.屏蔽费别,B.费用类型,Nvl(A.从项,0) as 从项" & _
                        " From 病人医嘱计价 A,收费项目目录 B" & _
                        " Where A.收费细目ID=B.ID And Nvl(A.数量,0)<>0 And A.医嘱ID=" & Val(.TextMatrix(lngTmp, COL_ID))
                End If
            End If
            vsPrice.ColHidden(COLP_付数) = True
        ElseIf .TextMatrix(lngRow, COL_诊疗类别) = "E" Then
            lngTmp = .FindRow(.TextMatrix(lngRow, COL_ID), , COL_相关ID)
            For i = lngTmp To lngRow
                If .TextMatrix(i, COL_诊疗类别) = "7" Then
                    '组成中药
                    If Val(.TextMatrix(i, COL_执行性质ID)) <> 5 Then
                        '中药药房单位按不可分零处理:每付
                        If Val(.TextMatrix(i, COL_可否分零)) = 0 Then
                            dbl数量 = Format(Val(.TextMatrix(i, COL_单量)) / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_住院包装)), "0.00000")
                        Else
                            dbl数量 = IntEx(Val(.TextMatrix(i, COL_单量)) / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_住院包装)))
                        End If
                        strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                            "Select '药品医嘱-" & Replace(.Cell(flexcpData, i, COL_医嘱内容), "'", "''") & "' as 计价医嘱," & i & " as 行号," & _
                            " A.ID,A.类别,A.名称||Decode(A.产地,NULL,NULL,'('||A.产地||')')||Decode(A.规格,NULL,NULL,' '||A.规格) as 名称," & _
                            " B.住院单位 as 单位,A.是否变价,A.加班加价,0 as 单价," & Val(.TextMatrix(i, COL_总量)) & " as 付数," & _
                            dbl数量 & " as 数量," & Val(.TextMatrix(i, COL_执行科室ID)) & " as 执行科室ID," & _
                            " A.屏蔽费别,A.费用类型,0 as 从项" & _
                            " From 收费项目目录 A,药品规格 B Where A.ID=B.药品ID And A.ID=" & Val(.TextMatrix(i, COL_药品ID))
                    End If
                Else
                    '中药煎法,用法
                    If Val(.TextMatrix(i, COL_计价特性)) = 0 _
                        And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                        If Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                            strTmp = "中药用法-" & Replace(.Cell(flexcpData, i, COL_医嘱内容), "'", "''")
                        Else
                            strTmp = "中药煎法-" & Replace(.Cell(flexcpData, i, COL_医嘱内容), "'", "''")
                        End If
                        strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                            "Select '" & strTmp & "' as 计价医嘱," & _
                            i & " as 行号,B.ID,B.类别,B.名称,B.计算单位 as 单位,B.是否变价,B.加班加价,A.单价,1 as 付数," & _
                            " Nvl(A.数量,0)*" & Val(.TextMatrix(i, COL_总量)) & " as 数量," & _
                            " Nvl(A.执行科室ID," & Val(.TextMatrix(i, COL_执行科室ID)) & ") as 执行科室ID," & _
                            " B.屏蔽费别,B.费用类型,Nvl(A.从项,0) as 从项" & _
                            " From 病人医嘱计价 A,收费项目目录 B" & _
                            " Where A.收费细目ID=B.ID And Nvl(A.数量,0)<>0 And A.医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                    End If
                End If
            Next
            vsPrice.ColHidden(COLP_付数) = False
        End If
    End With
    
    With vsPrice
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If strSQL <> "" Then
            '以最新价格计算
            strSQL = _
                " Select A.行号,B.收入项目ID,A.计价医嘱,A.ID,A.类别,C.名称 as 类别名称," & _
                " A.名称,A.单位,D.住院单位,A.付数,A.数量,A.执行科室ID,F.名称 as 执行科室,A.屏蔽费别,A.费用类型,D.住院包装," & _
                " A.是否变价,E.跟踪在用,A.从项,A.加班加价,B.加班加价率,Decode(Nvl(A.是否变价,0),1,A.单价,B.现价) as 单价" & _
                " From (" & strSQL & ") A,收费价目 B,收费项目类别 C,药品规格 D,材料特性 E,部门表 F" & _
                " Where A.ID=B.收费细目ID And A.类别=C.编码 And A.ID=D.药品ID(+) And A.ID=E.材料ID(+) And A.执行科室ID=F.ID(+)" & _
                " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                " Order by A.行号,A.从项,A.ID"
            Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
            
            If Not rsTmp.EOF And gbln从项汇总折扣 Then
                Set rsClone = rsTmp.Clone
            End If
            
            For i = 1 To rsTmp.RecordCount
                If str行号 <> rsTmp!行号 & "_" & rsTmp!ID Then
                    If str行号 <> "" Then
                        .TextMatrix(.Rows - 1, COLP_单价) = Format(dbl单价, "0.00000")
                        .TextMatrix(.Rows - 1, COLP_应收金额) = Format(cur应收, gstrDec)
                        .TextMatrix(.Rows - 1, COLP_实收金额) = Format(cur实收, gstrDec)
                    End If
                    str行号 = rsTmp!行号 & "_" & rsTmp!ID
                    dbl单价 = 0: cur应收 = 0: cur实收 = 0
                    .Rows = .Rows + 1
                    
                    .TextMatrix(.Rows - 1, COLP_行号) = rsTmp!行号
                
                    .TextMatrix(.Rows - 1, COLP_计价医嘱) = rsTmp!计价医嘱
                    .TextMatrix(.Rows - 1, COLP_类别) = rsTmp!类别名称
                    .TextMatrix(.Rows - 1, COLP_收费项目) = rsTmp!名称
                    .TextMatrix(.Rows - 1, COLP_费用类型) = Nvl(rsTmp!费用类型)
                    .TextMatrix(.Rows - 1, COLP_从项) = IIF(Nvl(rsTmp!从项, 0) = 0, "", "√")
                
                    If rsTmp!类别 = "7" Then
                        .TextMatrix(.Rows - 1, COLP_付数) = Nvl(rsTmp!付数, 1)
                    End If
                    
                    '药品计价的数量按住院单位显示
                    dbl数量 = Nvl(rsTmp!数量, 0) '取售价数量计算打折用
                    If InStr(",5,6,7,", rsTmp!类别) > 0 Then
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!行号, COL_诊疗类别)) > 0 Then
                            .TextMatrix(.Rows - 1, COLP_单位) = Nvl(rsTmp!单位)
                            .TextMatrix(.Rows - 1, COLP_数量) = FormatEx(Nvl(rsTmp!数量, 0), 5)
                            dbl数量 = dbl数量 * Nvl(rsTmp!住院包装, 1)
                        Else
                            .TextMatrix(.Rows - 1, COLP_单位) = Nvl(rsTmp!住院单位)
                            '中药药房单位按不可分零处理:每付
                            '非药嘱药品计价:因为这里预定了售价数量,因此转换为药房单位显示时不作不分零处理
                            .TextMatrix(.Rows - 1, COLP_数量) = FormatEx(Nvl(rsTmp!数量, 0) / Nvl(rsTmp!住院包装, 1), 5)
                        End If
                    Else
                        .TextMatrix(.Rows - 1, COLP_单位) = Nvl(rsTmp!单位)
                        .TextMatrix(.Rows - 1, COLP_数量) = FormatEx(Nvl(rsTmp!数量, 0), 5)
                    End If
                    
                    '执行科室:及相关附加数据
                    .TextMatrix(.Rows - 1, COLP_执行科室) = Nvl(rsTmp!执行科室)
                    .Cell(flexcpData, .Rows - 1, COLP_类别) = CStr(rsTmp!类别) '收费类别
                    .Cell(flexcpData, .Rows - 1, COLP_收费项目) = Val(rsTmp!ID) '项目ID
                    .Cell(flexcpData, .Rows - 1, COLP_执行科室) = Val(Nvl(rsTmp!执行科室ID, 0)) '执行科室ID
                    .Cell(flexcpData, .Rows - 1, COLP_费用类型) = Val(Nvl(rsTmp!跟踪在用, 0)) '跟踪在用
                    
                    '重新检查非药嘱药品及跟踪卫材的有效执行科室
                    lng执行科室ID = Nvl(rsTmp!执行科室ID, 0)
                    If rsTmp!类别 = "4" And Nvl(rsTmp!跟踪在用, 0) = 1 _
                        Or InStr(",5,6,7,", rsTmp!类别) > 0 And InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!行号, COL_诊疗类别)) = 0 Then
                        lng病人科室ID = Val(vsAdvice.TextMatrix(rsTmp!行号, COL_病人科室ID))
                        lng执行科室ID = Get收费执行科室ID(lng病人ID, lng主页ID, rsTmp!类别, rsTmp!ID, 4, lng病人科室ID, 0, 2, lng执行科室ID)
                        If lng执行科室ID <> Val(Nvl(rsTmp!执行科室ID, 0)) Then
                            .TextMatrix(.Rows - 1, COLP_执行科室) = Get部门名称(lng执行科室ID)
                            .Cell(flexcpData, .Rows - 1, COLP_执行科室) = lng执行科室ID
                        End If
                    End If
                    
                    '记录从属主项的信息，以便计算
                    If gbln从项汇总折扣 And rsTmp!从项 = 0 Then
                        If InStr(strHaveSub & ",", "," & rsTmp!行号 & ",") = 0 _
                            And InStr(strNoneSub & ",", "," & rsTmp!行号 & ",") = 0 Then
                            rsClone.Filter = "行号=" & rsTmp!行号 & " And 从项=1"
                            If Not rsClone.EOF Then
                                rsMain.AddNew
                                rsMain!医嘱行号 = rsTmp!行号
                                rsMain!主项行号 = .Rows - 1
                                rsMain!主收入ID = rsTmp!收入项目ID
                                rsMain.Update
                                strHaveSub = strHaveSub & "," & rsTmp!行号
                            Else
                                strNoneSub = strNoneSub & "," & rsTmp!行号
                            End If
                        End If
                    End If
                End If
                
                '单价及金额计算处理:非药嘱药品计价以住院单位处理(与门诊及计价设置处不同)
                If InStr(",5,6,7,", rsTmp!类别) > 0 Then
                    If Nvl(rsTmp!是否变价, 0) = 0 Then
                        dbl当前单价 = Nvl(rsTmp!单价, 0)
                    Else
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!行号, COL_诊疗类别)) > 0 Then
                            dbl当前单价 = CalcDrugPrice(rsTmp!ID, Val(.Cell(flexcpData, .Rows - 1, COLP_执行科室)), _
                                Format(Nvl(rsTmp!付数, 1) * Nvl(rsTmp!数量, 0) * Nvl(rsTmp!住院包装, 1), "0.00000"), , True)
                        Else
                            dbl当前单价 = CalcDrugPrice(rsTmp!ID, Val(.Cell(flexcpData, .Rows - 1, COLP_执行科室)), _
                                Format(Nvl(rsTmp!付数, 1) * Nvl(rsTmp!数量, 0), "0.00000"), , True)
                        End If
                    End If
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!行号, COL_诊疗类别)) > 0 Then
                        dbl当前单价 = Format(dbl当前单价 * Nvl(rsTmp!住院包装, 1), "0.00000")
                        cur当前应收 = Nvl(rsTmp!付数, 1) * Format(Nvl(rsTmp!数量, 0), "0.00000") * dbl当前单价
                    Else
                        cur当前应收 = Nvl(rsTmp!付数, 1) * Format(Nvl(rsTmp!数量, 0), "0.00000") * dbl当前单价
                        dbl当前单价 = Format(dbl当前单价 * Nvl(rsTmp!住院包装, 1), "0.00000")
                    End If
                ElseIf rsTmp!类别 = "4" And Nvl(rsTmp!跟踪在用, 0) = 1 And Nvl(rsTmp!是否变价, 0) = 1 Then
                    '时价卫材的单价和药品一样计算
                    dbl当前单价 = CalcDrugPrice(rsTmp!ID, Val(.Cell(flexcpData, .Rows - 1, COLP_执行科室)), Format(Nvl(rsTmp!数量, 0), "0.00000"), , True)
                    cur当前应收 = Format(Nvl(rsTmp!数量, 0), "0.00000") * dbl当前单价
                Else
                    dbl当前单价 = Format(Nvl(rsTmp!单价, 0), "0.00000")
                    cur当前应收 = Nvl(rsTmp!付数, 1) * Format(Nvl(rsTmp!数量, 0), "0.00000") * dbl当前单价
                End If
                
                '处理加班加价
                If gbln加班加价 And Nvl(rsTmp!加班加价, 0) = 1 Then
                    cur当前应收 = Format(cur当前应收 * (1 + Nvl(rsTmp!加班加价率, 0) / 100), gstrDec)
                Else
                    cur当前应收 = Format(cur当前应收, gstrDec)
                End If
                
                '实收
                If gbln从项汇总折扣 And (rsTmp!从项 = 1 Or InStr(strHaveSub & ",", "," & rsTmp!行号 & ",") > 0) Then
                    cur当前实收 = Format(cur当前应收, gstrDec)
                    '累计医嘱合计来计算折扣
                    rsMain.Filter = "医嘱行号=" & rsTmp!行号
                    rsMain!医嘱合计 = Nvl(rsMain!医嘱合计, 0) + cur当前实收
                    rsMain.Update
                ElseIf Nvl(rsTmp!屏蔽费别, 0) = 0 Then
                    cur当前实收 = Format(ActualMoney(str费别, rsTmp!收入项目ID, cur当前应收, rsTmp!ID, lng执行科室ID, dbl数量, _
                        IIF(gbln加班加价 And Nvl(rsTmp!加班加价, 0) = 1, Nvl(rsTmp!加班加价率, 0) / 100, 0)), gstrDec)
                Else
                    cur当前实收 = Format(cur当前应收, gstrDec)
                End If
                
                dbl单价 = dbl单价 + dbl当前单价
                cur应收 = cur应收 + cur当前应收
                cur实收 = cur实收 + cur当前实收
                rsTmp.MoveNext
            Next
            If str行号 <> "" Then
                .TextMatrix(.Rows - 1, COLP_单价) = Format(dbl单价, "0.00000")
                .TextMatrix(.Rows - 1, COLP_应收金额) = Format(cur应收, gstrDec)
                .TextMatrix(.Rows - 1, COLP_实收金额) = Format(cur实收, gstrDec)
            End If
        End If
        
        '汇总计算折扣:主项汇总打折不支持按成本加收打折
        If gbln从项汇总折扣 And strHaveSub <> "" Then
            rsMain.Filter = 0
            Do While Not rsMain.EOF
                cur当前实收 = Format(ActualMoney(str费别, rsMain!主收入ID, rsMain!医嘱合计), gstrDec)
                .TextMatrix(rsMain!主项行号, COLP_实收金额) = Format(Val(.TextMatrix(rsMain!主项行号, COLP_实收金额)) + (cur当前实收 - rsMain!医嘱合计), gstrDec)
                rsMain.MoveNext
            Loop
        End If
        
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        .Row = .FixedRows: .Col = .FixedCols
        .ShowCell .Row, .Col
        .Redraw = flexRDDirect
        Call vsPrice_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    ShowSendPrice = True
    Exit Function
errH:
    vsPrice.Redraw = flexRDDirect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'功能：更改成药规格
    Dim rsDrug As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng次数 As Long, lng最小次数 As Long
    Dim dbl总量 As Double, str分解时间 As String
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim cur金额 As Currency
    
    If Col = COL_附加执行 Then
        With vsAdvice
            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Call vsAdvice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
        End With
    ElseIf Col = COL_规格 Then
        With vsAdvice
            If Val(.TextMatrix(Row, COL_药品ID)) = .ComboData Then Exit Sub
            '药品相关信息
            .TextMatrix(Row, COL_药品ID) = .ComboData
            Set rsDrug = GetDrugInfo(Val(.TextMatrix(Row, COL_诊疗项目ID)), Val(.TextMatrix(Row, COL_药品ID)), Val(.TextMatrix(Row, COL_执行科室ID)))
            .TextMatrix(Row, COL_规格) = rsDrug!名称 & IIF(Not IsNull(rsDrug!产地), "(" & rsDrug!产地 & ")", "") & IIF(Not IsNull(rsDrug!规格), " " & rsDrug!规格, "")
            .TextMatrix(Row, COL_剂量系数) = rsDrug!剂量系数
            .TextMatrix(Row, COL_住院包装) = rsDrug!住院包装
            .TextMatrix(Row, COL_是否变价) = rsDrug!是否变价
            .TextMatrix(Row, COL_药房分批) = rsDrug!药房分批
            
            .TextMatrix(Row, COL_总量单位) = rsDrug!住院单位
            .TextMatrix(Row, COL_库存) = Format(Nvl(rsDrug!库存, 0), "0.00000")
            
            '医嘱相关信息
            strSQL = "Select ID,开始执行时间,上次执行时间,执行终止时间,执行时间方案," & _
                " 频率次数,频率间隔,间隔单位,单次用量 From 病人医嘱记录 Where ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(Row, COL_ID)))
            
            '重新计算总量,次数,分解时间
            Call Calc总量次数时间(dbl总量, lng次数, str分解时间, mstrEnd, rsTmp, rsDrug)
            
            .TextMatrix(Row, COL_总量) = FormatEx(dbl总量, 5)
            .TextMatrix(Row, COL_总量单位) = Nvl(rsDrug!住院单位)
            
            .TextMatrix(Row, COL_次数) = lng次数
            .TextMatrix(Row, COL_分解时间) = str分解时间
            .TextMatrix(Row, COL_首次时间) = Format(Split(str分解时间, ",")(0), "MM-dd HH:mm")
            .TextMatrix(Row, COL_末次时间) = Format(Split(str分解时间, ",")(lng次数 - 1), "MM-dd HH:mm")
                        
            '同步更改给药途径的次数
            i = .FindRow(.TextMatrix(Row, COL_相关ID), , COL_ID)
            .TextMatrix(i, COL_次数) = .TextMatrix(Row, COL_次数)
            .TextMatrix(i, COL_分解时间) = .TextMatrix(Row, COL_分解时间)
            .TextMatrix(i, COL_首次时间) = .TextMatrix(Row, COL_首次时间)
            .TextMatrix(i, COL_末次时间) = .TextMatrix(Row, COL_末次时间)
                                    
            '一并给药的按最小次数计算
            If RowIn一并给药(Row, lngBegin, lngEnd) Then
                For i = lngBegin To lngEnd
                    If Val(.TextMatrix(i, COL_次数)) < lng最小次数 Or lng最小次数 = 0 Then
                        lng最小次数 = Val(.TextMatrix(i, COL_次数))
                    End If
                Next
                For i = lngBegin To lngEnd + 1
                    If Val(.TextMatrix(i, COL_次数)) > lng最小次数 Then
                        .TextMatrix(i, COL_次数) = lng最小次数
                        .TextMatrix(i, COL_分解时间) = Trim分解时间(lng最小次数, .TextMatrix(i, COL_分解时间))
                        .TextMatrix(i, COL_首次时间) = Format(Split(.TextMatrix(i, COL_分解时间), ",")(0), "MM-dd HH:mm")
                        .TextMatrix(i, COL_末次时间) = Format(Split(.TextMatrix(i, COL_分解时间), ",")(lng最小次数 - 1), "MM-dd HH:mm")
                    End If
                Next
            Else
                lngBegin = Row: lngEnd = Row
            End If
            
            '重新计算并显示金额
            For i = lngBegin To lngEnd + 1
                .TextMatrix(i, COL_金额) = Format(Calc医嘱记帐金额(i), gstrDec)
                If .TextMatrix(i, COL_诊疗类别) = "E" And Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                    cur金额 = Val(.TextMatrix(i, COL_金额))
                End If
            Next
            '一并给药的第一行(如果是)才显示包含给药途径的金额
            .TextMatrix(lngBegin, COL_金额) = Format(Val(.TextMatrix(lngBegin, COL_金额)) + cur金额, gstrDec)
            
            '根据库存决定选择状态
            If Val(.TextMatrix(Row, COL_总量)) > Val(.TextMatrix(Row, COL_库存)) Then
                If GetStockCheck(Val(.TextMatrix(Row, COL_执行科室ID))) = 2 _
                    Or Val(.TextMatrix(Row, COL_药房分批)) = 1 Or Val(.TextMatrix(Row, COL_是否变价)) = 1 Then
                    .Cell(flexcpData, Row, COL_选择) = 1
                    Set .Cell(flexcpPicture, Row, COL_选择) = img16.ListImages("F").Picture
                ElseIf GetStockCheck(Val(.TextMatrix(Row, COL_执行科室ID))) = 1 Then
                    .Cell(flexcpData, Row, COL_选择) = 0
                    Set .Cell(flexcpPicture, Row, COL_选择) = Nothing
                ElseIf GetStockCheck(Val(.TextMatrix(Row, COL_执行科室ID))) = 0 Then
                    .Cell(flexcpData, Row, COL_选择) = 0
                    Set .Cell(flexcpPicture, Row, COL_选择) = img16.ListImages("T").Picture
                End If
            ElseIf Val(.TextMatrix(Row, COL_总量)) <= Val(.TextMatrix(Row, COL_库存)) Then
                .Cell(flexcpData, Row, COL_选择) = 0
                Set .Cell(flexcpPicture, Row, COL_选择) = img16.ListImages("T").Picture
            End If
            Call RowSelectSame(Row, COL_选择)
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
            Call ShowSendTotal
        End With
    End If
End Sub

Private Function Calc医嘱记帐金额(ByVal lngRow As Long) As Currency
'功能：计算指定医嘱行的记帐金额(显示供查看及记帐报警),以最新价格计算
'说明：药品计算过程按住院包装处理,但最终金额都是相同的
'返回：str类别=计价类别
    Dim str费别 As String, dbl数量 As Double
    Dim dbl单价 As Double, cur金额 As Currency
    
    With vsAdvice
        str费别 = .TextMatrix(lngRow, COL_费别)
        If InStr(",5,6,7,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            '药品缺省固定为正常计价,但下医嘱时指定了为自备药(院外执行)的不读取;药品不可能为叮嘱
            If Val(.TextMatrix(lngRow, COL_执行性质ID)) <> 5 Then
                '按售价单位的数量
                If .TextMatrix(lngRow, COL_诊疗类别) = "7" Then
                    '中药药房单位按不可分零处理:每付
                    If Val(.TextMatrix(lngRow, COL_可否分零)) = 0 Then
                        dbl数量 = Val(.TextMatrix(lngRow, COL_总量)) * Val(.TextMatrix(lngRow, COL_单量)) '单味剂量
                        dbl数量 = dbl数量 / Val(.TextMatrix(lngRow, COL_剂量系数))
                    Else
                        dbl数量 = IntEx(Val(.TextMatrix(lngRow, COL_单量)) / Val(.TextMatrix(lngRow, COL_剂量系数)) / Val(.TextMatrix(lngRow, COL_住院包装)))
                        dbl数量 = dbl数量 * Val(.TextMatrix(lngRow, COL_住院包装)) * Val(.TextMatrix(lngRow, COL_总量))
                    End If
                Else
                    dbl数量 = Val(.TextMatrix(lngRow, COL_总量)) * Val(.TextMatrix(lngRow, COL_住院包装))
                End If
                dbl数量 = Format(dbl数量, "0.00000")
                
                '按售价单位的价格
                If str费别 = "" Then
                    If Val(.TextMatrix(lngRow, COL_是否变价)) = 1 Then
                        '药品实价按售价数量算
                        dbl单价 = CalcDrugPrice(Val(.TextMatrix(lngRow, COL_药品ID)), Val(.TextMatrix(lngRow, COL_执行科室ID)), dbl数量)
                    Else
                        dbl单价 = CalcPrice(Val(.TextMatrix(lngRow, COL_药品ID)))
                    End If
                    cur金额 = Format(dbl数量 * Format(dbl单价, "0.00000"), gstrDec)
                Else
                    If Val(.TextMatrix(lngRow, COL_是否变价)) = 1 Then
                        '药品实价按售价数量算
                        cur金额 = Format(CalcDrugPrice(Val(.TextMatrix(lngRow, COL_药品ID)), Val(.TextMatrix(lngRow, COL_执行科室ID)), dbl数量, str费别), gstrDec)
                    Else
                        cur金额 = Format(CalcPrice(Val(.TextMatrix(lngRow, COL_药品ID)), str费别, dbl数量, , Val(.TextMatrix(lngRow, COL_执行科室ID))), gstrDec)
                    End If
                End If
            End If
        Else
            '不计价,手工计价；叮嘱,院外执行；的医嘱不读取
            If Val(.TextMatrix(lngRow, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质ID))) = 0 Then
                dbl数量 = Format(Val(.TextMatrix(lngRow, COL_总量)), "0.00000")
                If str费别 = "" Then
                    dbl单价 = Format(CalcAdvicePrice(Val(.TextMatrix(lngRow, COL_ID))), "0.00000")
                    cur金额 = Format(dbl数量 * dbl单价, gstrDec)
                Else
                    cur金额 = Format(CalcAdvicePrice(Val(.TextMatrix(lngRow, COL_ID)), str费别, , dbl数量), gstrDec)
                End If
            End If
        End If
    End With
    Calc医嘱记帐金额 = cur金额
End Function

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsAdvice
        '根据可否编辑设置编辑特性及光标特性
        If NewCol = COL_规格 Then
            .ComboList = .Cell(flexcpData, NewRow, NewCol)
            .FocusRect = flexFocusLight
        ElseIf NewCol = COL_附加执行 Then
            If Should附加执行(NewRow) Then
                .ComboList = "..."
                Set .CellButtonPicture = Me.Picture
                .FocusRect = flexFocusHeavy
            Else
                .ComboList = ""
                .FocusRect = flexFocusLight
            End If
        Else
            .ComboList = ""
            .FocusRect = flexFocusLight
        End If
        
        If OldRow <> NewRow And .Redraw <> flexRDNone And Not .RowHidden(NewRow) Then
            If Val(.TextMatrix(NewRow, COL_ID)) <> 0 Then
                Call ShowSendPrice(NewRow)
            End If
        End If
    End With
End Sub

Private Function Should附加执行(ByVal lngRow As Long) As Boolean
'功能：判断指定的医嘱行(可见行)是否可以设置附加的执行科室
    Dim lngRow2 As Long, i As Long
    
    If lngRow = 0 Then Exit Function
    
    lngRow2 = -1
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_ID)) = 0 Then Exit Function
        If .TextMatrix(lngRow, COL_诊疗类别) = "E" And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 _
            And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) _
            And InStr(",7,E,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 Then
            '中药用法
            lngRow2 = lngRow
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            '给药途径
            lngRow2 = .FindRow(CLng(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1, COL_ID)
        End If
        
        '叮嘱或院外执行
        If lngRow2 <> -1 Then
            If InStr(",0,5,", Val(.TextMatrix(lngRow2, COL_执行性质ID))) = 0 Then
                Should附加执行 = True
            End If
        End If
    End With
End Function

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
        If Col = COL_医嘱内容 Or Col = COL_规格 Then
            If Not .ColHidden(COL_规格) Then
                .AutoSize COL_医嘱内容, COL_规格
            Else
                .AutoSize COL_医嘱内容
            End If
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

Private Sub vsAdvice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim vPoint As POINTAPI, blnCancel As Boolean
    
    strSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " Order by A.编码"
    With vsAdvice
        vPoint = GetCoordPos(.Hwnd, .CellLeft, .CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "执行科室", , , , , , True, vPoint.x, vPoint.y, .CellHeight, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            Call SetDeptInput(Row, Col, rsTmp)
            Call vsAdvice_AfterRowColChange(-1, -1, Row, Col) '重新显示计价执行科室
        Else
            If Not blnCancel Then
                MsgBox "没有可用的科室数据，请先到部门管理中设置。", vbInformation, gstrSysName
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_ChangeEdit()
    If vsAdvice.Col = COL_规格 Then
        Call vsAdvice_AfterEdit(vsAdvice.Row, vsAdvice.Col)
    End If
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
                vRect.Bottom = Bottom - 2 '底行保留下边线(本窗体中用到下边线粗为2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode > 127 Then '解决直接输入汉字的问题
        Call vsAdvice_KeyPress(KeyCode)
    End If
End Sub

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
    Dim i As Long
    
    With vsAdvice
        For i = lngRow + 1 To .Rows - 1
            If Not .RowHidden(i) Then
                .Row = i: Exit For
            End If
        Next
        If i > .Rows - 1 Then .Row = .FixedRows
        If .RowHidden(.Row) Then .Row = lngRow
        Call .ShowCell(.Row, .Col)
    End With
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterNextCell(.Row, .Col)
        ElseIf KeyAscii = 32 And .Col = COL_选择 Then
            KeyAscii = 0
            If .Cell(flexcpData, .Row, COL_选择) = 0 Then
                If .Cell(flexcpPicture, .Row, COL_选择) Is Nothing Then
                    Set .Cell(flexcpPicture, .Row, COL_选择) = img16.ListImages("T").Picture
                Else
                    Set .Cell(flexcpPicture, .Row, COL_选择) = Nothing
                End If
                Call RowSelectSame(.Row, .Col)
                Call ShowSendTotal
            End If
        Else
            If .Col = COL_附加执行 And .ComboList = "..." Then
                If Should附加执行(.Row) Then blnDo = True
            End If
            If blnDo Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsAdvice_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, StrInput As String
    Dim vPoint As POINTAPI, blnCancel As Boolean
    
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Col = COL_规格 Then
                Call vsAdvice_KeyPress(13)
            ElseIf Col = COL_附加执行 And .EditText <> "" Then
                StrInput = UCase(.EditText)
                strSQL = "Select Distinct A.ID,A.编码,A.名称,A.简码" & _
                    " From 部门表 A,部门性质说明 B" & _
                    " Where A.ID=B.部门ID And B.服务对象 IN(2,3)" & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (Upper(A.编码) Like [1] Or Upper(A.名称) Like [2] Or Upper(A.简码) Like [2])" & _
                    " Order by A.编码"
                With vsAdvice
                    vPoint = GetCoordPos(.Hwnd, .CellLeft, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "执行科室", False, "", "", False, False, True, _
                        vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, StrInput & "%", mstrLike & StrInput & "%")
                    If Not rsTmp Is Nothing Then
                        Call SetDeptInput(Row, Col, rsTmp)
                        .EditText = .TextMatrix(Row, Col) '直接输入匹配需要
                        Call EnterNextCell(Row, Col)
                    Else
                        If Not blnCancel Then
                            MsgBox "没有找到匹配的科室。", vbInformation, gstrSysName
                        End If
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                        Call vsAdvice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
                    End If
                End With
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsAdvice.EditSelStart = 0
    vsAdvice.EditSelLength = zlCommFun.ActualLen(vsAdvice.EditText)
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsAdvice
        If Col = COL_规格 Then
            If .ComboList = "" Then Cancel = True
        ElseIf Col = COL_附加执行 Then
            If Not Should附加执行(Row) Then Cancel = True
        Else
            Cancel = True
        End If
    End With
End Sub

Private Sub vsPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngRow As Long
    
    If NewRow <> OldRow Then
        With vsPrice
            stbThis.Panels(2).Text = ""
            lngRow = Val(.TextMatrix(NewRow, COLP_行号))
            If lngRow <> 0 And .Cell(flexcpData, NewRow, COLP_类别) <> "" Then
                If InStr(",5,6,7,", .Cell(flexcpData, NewRow, COLP_类别)) > 0 _
                    Or .Cell(flexcpData, NewRow, COLP_类别) = "4" And Val(.Cell(flexcpData, NewRow, COLP_费用类型)) = 1 Then
                    '显示药品及跟踪卫材的库存
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                        stbThis.Panels(2).Text = vsAdvice.TextMatrix(lngRow, COL_规格) & "，" & vsAdvice.TextMatrix(lngRow, COL_执行科室) & "可用库存：" & _
                            FormatEx(Val(vsAdvice.TextMatrix(lngRow, COL_库存)), 5) & vsAdvice.TextMatrix(lngRow, COL_住院单位)
                    Else
                        '同一个函数取:药品按住院单位,卫材按售价单位
                        stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_收费项目) & "，" & .TextMatrix(NewRow, COLP_执行科室) & "可用库存：" & _
                            FormatEx(GetStock(Val(.Cell(flexcpData, NewRow, COLP_收费项目)), Val(.Cell(flexcpData, NewRow, COLP_执行科室))), 5) & .TextMatrix(NewRow, COLP_单位)
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsPrice_GotFocus()
    vsPrice.BackColorSel = &HFFCC99
End Sub

Private Sub vsPrice_LostFocus()
    vsPrice.BackColorSel = &HFFEBD7
End Sub

Private Sub vsAdvice_GotFocus()
    vsAdvice.BackColorSel = &HFFCC99
End Sub

Private Sub vsAdvice_LostFocus()
    vsAdvice.BackColorSel = &HFFEBD7
End Sub

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
    
    strHead = ",300,4;科室,850,1;姓名,750,1;住院号,750,1;床号,500,4;费别,750,1;" & _
        "婴儿,550,1;医嘱内容,2000,1;规格,2000,1;总量,600,7;单位,450,1;单量,600,7;单位,450,1;金额,850,7;" & _
        "频率,1000,1;用法,1000,1;医生嘱托,1500,1;执行时间,1000,1;首次时间,1080,1;末次时间,1080,1;执行科室,850,1;附加执行,850,1;执行性质,850,1;" & _
        "病人ID;主页ID;性别;年龄;ID;相关ID;病人科室ID;开嘱科室ID;开嘱医生;诊疗类别;诊疗项目ID;计价特性;执行性质ID;执行科室ID;" & _
        "药品ID;剂量系数;住院包装;住院单位;可否分零;药房分批;是否变价;库存;次数;分解时间"
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
    
    strHead = "计价医嘱,2000,1;类别,650,1;收费项目,2500,1;付数,450,4;数量,900,7;" & _
        "单位,500,1;单价,1000,7;应收金额,1200,7;实收金额,1200,7;执行科室,1000,1;费用类型,850,1;从项,450,4;行号"
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

Private Function Decode药房置换() As String
'功能：根据药房置换设置,产生字段Decode语句
'说明：SQL查询中,"病人医嘱记录"别名为"A"
    Dim i As Long, strTmp As String
    
    mrs药房.Filter = 0
    If Not mrs药房.EOF Then
        For i = 1 To mrs药房.RecordCount
            If mrs药房!ID <> mrs药房!现ID Then
                strTmp = strTmp & "," & mrs药房!ID & "," & mrs药房!现ID
            End If
            mrs药房.MoveNext
        Next
    End If
    If strTmp <> "" Then
        Decode药房置换 = "Decode(A.执行科室ID" & strTmp & ",A.执行科室ID)"
    Else
        Decode药房置换 = "A.执行科室ID"
    End If
End Function

Private Sub DeleteCurRow(ByVal lngRow As Long, ByVal lng相关ID As Long)
'功能：在处理待发送清单的过程中删除最近加入的行
    Dim i As Long
    With vsAdvice
        '删除当前行
        .RemoveItem lngRow
        
        '删除配方或一并给药中已经加入的行
        If lng相关ID <> 0 Then
            For i = .Rows - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = lng相关ID Then
                    .RemoveItem i
                End If
            Next
        End If
    End With
End Sub

Private Function Calc总量次数时间(dbl总量 As Double, lng次数 As Long, str分解时间 As String, ByVal strEnd As String, rsSend As ADODB.Recordset, rsDrug As ADODB.Recordset) As Boolean
'功能：对长期成药医嘱计算总量,执行次数,执行时间分解
'参数：rsDrug=包含药品规格的相关信息
'      rsSend=包含当前药品医嘱的相关信息
'      strEnd=本次发送的结束时间
'返回：dbl总量=住院包装
'      lng次数=执行次数(即为给药途径的执行次数)
'      str分解时间=具体的执行时间分解
    Dim datBegin As Date, datEnd As Date, strPause As String
    
    '当前医嘱的暂停时间段:"暂停时间,开始时间;...."
    strPause = GetAdvicePause(rsSend!ID)
    
    '当前医嘱的发送计算时间段
    datBegin = rsSend!开始执行时间
    If Not IsNull(rsSend!上次执行时间) Then
        datBegin = Calc本周期开始时间(rsSend!开始执行时间, rsSend!上次执行时间, rsSend!频率间隔, rsSend!间隔单位)
        
        '本周期内已执行的时间不再计算,这里通过暂停方式来处理
        strPause = strPause & ";" & Format(datBegin, "yyyy-MM-dd HH:mm:ss") & "," & Format(rsSend!上次执行时间, "yyyy-MM-dd HH:mm:ss")
        If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
    End If
    datEnd = CDate(strEnd)
    If Not IsNull(rsSend!执行终止时间) Then
        If rsSend!执行终止时间 < CDate(strEnd) Then
            datEnd = rsSend!执行终止时间
        End If
    End If
    
    '先按正常发送时间段计算分解时间及次数
    str分解时间 = Calc段内分解时间(datBegin, datEnd, strPause, rsSend!执行时间方案, rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位)
    If str分解时间 = "" Then Calc总量次数时间 = True: Exit Function

    lng次数 = UBound(Split(str分解时间, ",")) + 1
    
    '再按药品分零特性计算总量(按住院单位),这时次数和分解时间可能增加
    dbl总量 = Calc发送药品总量( _
        rsSend!开始执行时间, lng次数, str分解时间, rsSend!单次用量, _
        rsDrug!剂量系数, rsDrug!住院包装, Nvl(rsDrug!可否分零, 0), _
        Nvl(rsSend!执行终止时间, CDate("3000-01-01")), strPause, Nvl(rsSend!执行时间方案), _
        rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位, mblnLimit)
    
    Calc总量次数时间 = True
End Function

Private Function CheckStock(ByVal lngRow As Long, Optional bln库存提示 As Boolean, Optional bln时价提示 As Boolean, Optional bln默认发送 As Boolean, Optional ByVal blnCurPati As Boolean) As Boolean
'功能：根据库存检查参数检查发送药品的库存
'参数：lngRow=医嘱行号
'      blnCurPati=是否只对当前病人进行汇总检查,用于发送过程中,因为是按病人提交,这时重新提取的库存是准确的
'      bln库存提示,bln时价提示,bln默认发送=用于提示框相关显示控制
'返回：根据提示，是否对选择状态进行了处理
    Dim int库存检查 As Integer, dbl总量 As Double
    Dim dbl可用库存 As Double, dbl已发库存 As Double
    Dim bln分批时价 As Boolean, bln分批 As Boolean, bln时价 As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        '药品库存检查(0-不检查;1-检查,不足提醒;2-检查，不足禁止)
        int库存检查 = GetStockCheck(Val(.TextMatrix(lngRow, COL_执行科室ID)))
        bln分批 = Val(.TextMatrix(lngRow, COL_药房分批)) = 1
        bln时价 = Val(.TextMatrix(lngRow, COL_是否变价)) = 1
        
        '分批或时价药品必须要有足够的库存,其它根据库存检查参数决定
        If int库存检查 <> 0 Or bln分批 Or bln时价 Then
            strTmp = .TextMatrix(lngRow, COL_住院单位) '用于提示
            
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
                If blnCurPati And Val(.TextMatrix(i, COL_病人ID)) = Val(.TextMatrix(lngRow, COL_病人ID)) Or Not blnCurPati Then
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0
                    If blnDo Then
                        blnDo = Val(.TextMatrix(i, COL_药品ID)) = Val(.TextMatrix(lngRow, COL_药品ID)) _
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
                Else
                    Exit For
                End If
            Next
            dbl可用库存 = Val(.TextMatrix(lngRow, COL_库存))
            dbl可用库存 = dbl可用库存 - dbl已发库存
            
            If dbl总量 > dbl可用库存 Then
                If (Not bln分批时价 And int库存检查 <> 0 And bln库存提示) Or (bln分批时价 And bln时价提示) Then
                    '上一次没有选择不再提示,则提示
                    If bln分批时价 Then
                        strTmp = "药房分批或时价药品""" & .TextMatrix(lngRow, COL_规格) & """库存不足：" & vbCrLf & vbCrLf & _
                            .TextMatrix(lngRow, COL_执行科室) & "可用库存：" & FormatEx(dbl可用库存, 5) & strTmp & _
                            IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，" & _
                            "本次发送量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                    Else
                        strTmp = """" & .TextMatrix(lngRow, COL_规格) & """库存不足：" & vbCrLf & vbCrLf & _
                            .TextMatrix(lngRow, COL_执行科室) & "可用库存：" & FormatEx(dbl可用库存, 5) & strTmp & _
                            IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，" & _
                            "本次发送量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                    End If
                    If .Cell(flexcpData, lngRow, COL_规格) <> "" Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "你可以在发送清单中选择该药品具有足够库存的其它规格。"
                    End If
                    If int库存检查 = 1 And Not bln分批时价 Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "要发送该药品吗？"
                    End If
                    
                    strTmp = "病人" & .TextMatrix(lngRow, COL_姓名) & "：" & vbCrLf & vbCrLf & strTmp
                    
                    .Redraw = flexRDDirect:
                    Call .ShowCell(lngRow, COL_选择)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int库存检查 = 2 Or bln分批时价)
                    
                    If bln分批时价 Then
                        If vMsg = vbIgnore Then bln时价提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = img16.ListImages("F").Picture
                        CheckStock = True
                    ElseIf int库存检查 = 2 Then '库存禁止
                        If vMsg = vbIgnore Then bln库存提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = img16.ListImages("F").Picture
                        CheckStock = True
                    ElseIf int库存检查 = 1 Then '库存提醒
                        If vMsg = vbYes Or vMsg = vbIgnore Then
                            If vMsg = vbIgnore Then bln库存提示 = False
                            bln默认发送 = True
                        ElseIf vMsg = vbNo Or vMsg = vbCancel Then
                            If vMsg = vbCancel Then bln库存提示 = False
                            bln默认发送 = False
                            Set .Cell(flexcpPicture, lngRow, COL_选择) = Nothing '缺省不发送
                            CheckStock = True
                        End If
                    End If
                    Screen.MousePointer = 11
                    .Refresh: .Redraw = flexRDNone
                Else
                    '上一次选择了不再提示
                    If int库存检查 = 2 Or bln分批 Or bln时价 Then
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = img16.ListImages("F").Picture
                        CheckStock = True
                    ElseIf int库存检查 = 1 Then
                        '根据上一次的结果处理
                        If Not bln默认发送 Then
                            Set .Cell(flexcpPicture, lngRow, COL_选择) = Nothing '缺省不发送
                            CheckStock = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

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
        int库存检查 = GetStockCheck(lng库房ID)
        bln分批 = Nvl(rsPrice!分批, 0) = 1
        bln时价 = Nvl(rsPrice!是否变价, 0) = 1
        
        '分批或时价药品必须要有足够的库存,其它根据库存检查参数决定
        If int库存检查 <> 0 Or bln分批 Or bln时价 Then
            strTmp = Nvl(rsPrice!住院单位, Nvl(rsPrice!计算单位)) '用于提示
            
            '当本身就不足禁止时,分批时间就不必单独处理
            bln分批时价 = int库存检查 <> 2 And (bln分批 Or bln时价)
            
            '当前药品或卫材总量:住院包装
            dbl总量 = Format(dbl数量 / Nvl(rsPrice!住院包装, 1), "0.00000")
            
            '当前可用库存:住院包装,减去前面相同药品医嘱要发送的库存
            If InStr(",5,6,7,", rsPrice!类别) > 0 Then
                For i = lngRow - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(i, COL_病人ID)) = Val(.TextMatrix(lngRow, COL_病人ID)) Then
                        blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0
                        If blnDo Then
                            blnDo = Val(.TextMatrix(i, COL_药品ID)) = rsPrice!ID And Val(.TextMatrix(i, COL_执行科室ID)) = lng库房ID
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
                    Else
                        Exit For
                    End If
                Next
            End If
            '计价部份要发送的累计数量
            rsTotal.Filter = "项目ID=" & rsPrice!ID & " And 库房ID=" & lng库房ID
            Do While Not rsTotal.EOF
                dbl已发库存 = dbl已发库存 + Format(rsTotal!数量 / Nvl(rsPrice!住院包装, 1), "0.00000")
                rsTotal.MoveNext
            Loop
            
            dbl可用库存 = Format(GetStock(rsPrice!ID, lng库房ID, 2), "0.00000")
            dbl可用库存 = dbl可用库存 - dbl已发库存
            
            If dbl总量 > dbl可用库存 Then
                If (Not bln分批时价 And int库存检查 <> 0 And bln库存提示) Or (bln分批时价 And bln时价提示) Then
                    '上一次没有选择不再提示,则提示
                    If bln分批时价 Then
                        strTmp = "医嘱""" & .TextMatrix(lngRow, COL_医嘱内容) & """的分批或时价计价项目""" & rsPrice!名称 & """库存不足：" & _
                            vbCrLf & vbCrLf & Get部门名称(lng库房ID) & "可用库存：" & FormatEx(dbl可用库存, 5) & strTmp & _
                            IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，本次发送数量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                    Else
                        strTmp = "医嘱""" & .TextMatrix(lngRow, COL_医嘱内容) & """的计价项目""" & rsPrice!名称 & """库存不足：" & _
                            vbCrLf & vbCrLf & Get部门名称(lng库房ID) & "可用库存：" & FormatEx(dbl可用库存, 5) & strTmp & _
                            IIF(dbl已发库存 <> 0, "(排开前面相同药品所需库存)", "") & "，本次发送数量：" & FormatEx(dbl总量, 5) & strTmp & "。"
                    End If
                    If int库存检查 = 1 And Not bln分批时价 Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "要发送该医嘱吗？"
                    End If
                    strTmp = "病人" & .TextMatrix(lngRow, COL_姓名) & "：" & vbCrLf & vbCrLf & strTmp
                    
                    .Redraw = flexRDDirect
                    .Row = GetVisibleRow(lngRow, True)
                    Call .ShowCell(.Row, COL_选择)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int库存检查 = 2 Or bln分批时价)
                    
                    If bln分批时价 Then
                        If vMsg = vbIgnore Then bln时价提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = img16.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int库存检查 = 2 Then '库存禁止
                        If vMsg = vbIgnore Then bln库存提示 = False
                        .Cell(flexcpData, lngRow, COL_选择) = 1 '当前规格禁止选择
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = img16.ListImages("F").Picture
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
                        Set .Cell(flexcpPicture, lngRow, COL_选择) = img16.ListImages("F").Picture
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

Private Sub DeleteDrugRow(rsSend As ADODB.Recordset, ByVal lngRow As Long, lngDel相关ID As Long)
'功能：删除对应的药品行,用于药品停用或因其它原因找不到有效规格时
'返回：lngDel相关ID-需要同时删除的其它相关医嘱标识
    Dim strMsg As String
    
    With vsAdvice
        If rsSend!诊疗类别 = "7" Then
            strMsg = "该中草药对应的中药配方无法发送：" & vbCrLf & vbCrLf & "　　" & Nvl(rsSend!医嘱内容)
        Else
            strMsg = "该药品(及一并给药的其他药品)无法发送：" & vbCrLf & vbCrLf & "　　" & Nvl(rsSend!医嘱内容)
        End If
        strMsg = strMsg & vbCrLf & vbCrLf & "没有发现有效的药品规格信息，该药品可能已经被停用或不能用于住院病人。"
        strMsg = strMsg & vbCrLf & "请先到药品目录管理中处理，按[确定]继续处理其他医嘱。"
        .Redraw = flexRDDirect
        Call .ShowCell(lngRow, COL_选择)
        Screen.MousePointer = 0
        MsgBox strMsg, vbInformation, gstrSysName
        
        Screen.MousePointer = 11
        lngDel相关ID = Nvl(rsSend!相关ID, 0)
        Call DeleteCurRow(lngRow, rsSend!相关ID)
        .Refresh: .Redraw = flexRDNone
    End With
End Sub

Private Sub SeekMatchDrug(rsSend As ADODB.Recordset, rsDrug As ADODB.Recordset, ByVal dbl总量 As Double, vBookMark As Variant, Optional strList As String)
'功能：根据药品的多个规格定位缺省合适的规格,并设置相关药品信息到表格中
'参数：rsSend=要发送的医嘱信息
'      rsDrug=药品信息
'      dbl总量=要发送的药品总量,为0时表示还未计算出来
'      vBookMark=返回用于定位规格位置的书签
'      strList=返回有效可供选择的规格,用于设置下拉框数据
    Dim vPreBookMark As Variant
    Dim lng倍数 As Long
        
    vPreBookMark = 0
    If Not rsDrug.EOF And Not rsDrug.BOF Then
        vPreBookMark = rsDrug.Bookmark
    End If
    
    rsDrug.MoveFirst
    vBookMark = 0: strList = ""
    Do While Not rsDrug.EOF
        '排开停用的药品
        If Nvl(rsDrug!撤档时间, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", Nvl(rsDrug!服务对象, 0)) > 0 Then
            If CInt(Nvl(rsSend!单次用量, 0)) <> 0 And (Nvl(rsDrug!库存, 0) > dbl总量 Or Nvl(rsDrug!库存, 0) = dbl总量 And dbl总量 <> 0) Then
                '寻找剂量单位为单量的最小倍数的规格
                If rsDrug!剂量系数 / rsSend!单次用量 = Int(rsDrug!剂量系数 / rsSend!单次用量) Then
                    If rsDrug!剂量系数 / rsSend!单次用量 < lng倍数 Or lng倍数 = 0 Then
                        vBookMark = rsDrug.Bookmark
                        lng倍数 = rsDrug!剂量系数 / rsSend!单次用量
                    End If
                End If
            End If
            strList = strList & "|#" & rsDrug!药品ID & ";" & rsDrug!名称 & IIF(Not IsNull(rsDrug!产地), "(" & rsDrug!产地 & ")", "") & IIF(Not IsNull(rsDrug!规格), " " & rsDrug!规格, "") & vbTab & "库存:" & Nvl(rsDrug!库存, 0) & rsDrug!住院单位
        End If
        rsDrug.MoveNext
    Loop
    If vBookMark = 0 Then
        rsDrug.MoveFirst
        Do While Not rsDrug.EOF
            If Nvl(rsDrug!撤档时间, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", Nvl(rsDrug!服务对象, 0)) > 0 Then
                If Nvl(rsDrug!库存, 0) > dbl总量 Or Nvl(rsDrug!库存, 0) = dbl总量 And dbl总量 <> 0 Then
                    vBookMark = rsDrug.Bookmark: Exit Do
                End If
            End If
            rsDrug.MoveNext
        Loop
    End If
    strList = Mid(strList, 2)
    
    If vBookMark = 0 And vPreBookMark <> 0 Then '没找到时恢复原有位置
        rsDrug.Bookmark = vPreBookMark
    End If
End Sub

Private Function LoadAdviceSend(ByVal strEnd As String, ByVal int期效 As Integer, _
    ByVal lng药房ID As Long, ByVal str病人IDs As String, ByVal str给药IDs As String) As Boolean
'功能：根据条件读取并显示要发送的药品医嘱清单
'参数：strEnd=发送到的结束时间(yyyy-MM-dd HH:mm:ss),临嘱没有
'      int期效=0-长嘱,1-临嘱
'      lng药房ID=要发送药品的药房ID,0表示不限制
'      str病人IDs=要发送医嘱病人ID串(12,23,34....)
'      str给药IDs=要发送的给药途径为指定ID串的成药
'      mrs药房=包含药品置换信息的记录集
'说明：注意CellData中存放得有附加数据
'   RowData：0-未发送的,-1-已成功发送的
'   COL_选择：0-可自由选择的,1-禁止改变选择状态的
'   COL_诊疗类别：1-给药途径，2-中药煎法，3-中药用法
'   COL_婴儿：存放婴儿编号
'   COL_医嘱内容：存放诊疗项目名称,用于显示计价医嘱
'   COL_规格：存放成药可选择的规格下拉数据(ComboList)
'   COL_分解时间:临嘱无分解时间时,存放费用发生时间

    Dim rsSend As New ADODB.Recordset
    Dim rsDrug As New ADODB.Recordset
    Dim strSQL As String, str期效条件 As String
    Dim str药房条件 As String, str给药途径 As String
    Dim i As Long, j As Long, k As Long, lngTmp As Long, strTmp As String
    Dim lng病人ID As Long, lng倍数 As Long, lng病人数 As Long
    Dim lngRow As Long, lngDel相关ID As Long, vBookMark As Variant
    Dim str药房置换 As String, str科室 As String, bln分批时价 As Boolean
    Dim lng次数 As Long, lng最小次数 As Long, str用法 As String
    Dim str分解时间 As String, dbl总量 As Double, cur金额 As Currency
    Dim bln时价提示 As Boolean, bln库存提示 As Boolean, bln默认发送 As Boolean
    Dim bln品种药品 As Boolean
        
    Screen.MousePointer = 11
    
    stbThis.Panels(3).Text = "": Call Form_Resize
    If int期效 = 0 Then
        lblInfo.Caption = "本次发送：长期药品医嘱，结束时间：" & strEnd
    Else
        lblInfo.Caption = "本次发送：临时药品医嘱"
    End If
    
    vsPrice.Rows = vsPrice.FixedRows
    vsPrice.Rows = vsPrice.FixedRows + 1
    vsAdvice.Rows = vsAdvice.FixedRows '有删除行功能
    Me.Refresh
        
    bln时价提示 = True: bln库存提示 = True: bln默认发送 = True
    
    '获取发送清单:每条医嘱记录(作废的医嘱不管作废时间,作废后即无效)
    '----------------------------------------------------------------------------------------------------------
    '置换的药房ID
    str药房置换 = Decode药房置换
    
    '不同期效的条件
    If int期效 = 0 Then
        str期效条件 = _
            " And A.开始执行时间<=[1] And (A.上次执行时间<[1] Or A.上次执行时间 is NULL)" & _
            " And (A.执行终止时间>A.上次执行时间 Or A.执行终止时间 is NULL Or A.上次执行时间 Is NULL)" & _
            " And (A.执行终止时间>A.开始执行时间 Or A.执行终止时间 is NULL)" & _
            " And Nvl(A.医嘱状态,0) Not IN(1,2,4) And Nvl(A.医嘱期效,0)=0"
    Else
        str期效条件 = " And Nvl(A.医嘱状态,0) Not IN(1,2,4,8,9) And Nvl(A.医嘱期效,0)=1"
    End If
    '发送的医嘱范围限制
    If InStr(mstrPrivs, "全院医嘱发送") = 0 Then
        str期效条件 = str期效条件 & " And A.开嘱医生 In(" & _
            " Select Distinct B.姓名" & _
            " From 部门人员 A,人员表 B,人员性质说明 C" & _
            " Where A.人员ID=B.ID And B.ID=C.人员ID And C.人员性质='医生'" & _
            "   And A.部门ID In(" & _
            "     Select Distinct B.科室ID From 部门人员 A,床位状况记录 B" & _
            "     Where A.人员ID=(Select 人员ID From 上机人员表 Where 用户名=User)" & _
            "       And A.部门ID=B.病区ID)" & _
            ")"
    End If
    
    For k = 0 To UBound(Split(str病人IDs, ","))
        '只发送指定药房的药品:药房置换之后的为准
        If lng药房ID <> 0 Then
            str药房条件 = "Select ID From 病人医嘱记录 X" & _
                " Where 诊疗类别 IN('5','6','7') And (X.相关ID=A.相关ID Or X.相关ID=A.ID)" & _
                " And " & Replace(str药房置换, "A.执行科室ID", "执行科室ID") & "+0=[3] And 病人ID=[2]"
            str药房条件 = " And Exists(" & str药房条件 & ")"
        End If
        
        '允许的给药途径部份(关联对应的成药)
        If str给药IDs <> "" Then
            str给药途径 = "Select ID From 病人医嘱记录 X" & _
                " Where 诊疗类别='E' And (X.ID=A.ID Or X.ID=A.相关ID)" & _
                " And 病人ID=[2] And 诊疗项目ID+0 IN(" & str给药IDs & ")"
            str给药途径 = " And Exists(" & str给药途径 & ")"
        End If
        
        '读取发送明细:(未排除正常的治疗医嘱)
        '叮嘱不发送(给药途径,用法,煎法可能为),但这里先读取出来
        strSQL = "Select A.ID,A.相关ID,Nvl(A.相关ID,A.ID) as 组ID,Nvl(X.序号,A.序号) as 组号," & _
            " A.诊疗类别,A.诊疗项目ID,E.名称 as 诊疗项目,A.收费细目ID,A.婴儿," & _
            " A.病人ID,A.主页ID,C.住院号,B.出院病床 as 床号,D.名称 as 科室,C.姓名,C.性别,C.年龄,B.费别,B.险类," & _
            " A.开始执行时间,A.上次执行时间,A.医嘱内容,A.天数,A.总给予量,A.单次用量,E.计算单位,A.执行终止时间," & _
            " A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,A.医生嘱托,A.执行时间方案,A.病人科室ID,A.开嘱科室ID,A.开嘱医生," & _
            " A.计价特性,A.执行性质," & str药房置换 & " as 执行科室ID,F.名称 as 执行科室" & _
            " From 病人医嘱记录 A,病案主页 B,病人信息 C,部门表 D,诊疗项目目录 E,部门表 F,病人医嘱记录 X" & _
            " Where A.病人ID=[2] And A.病人ID=C.病人ID And B.出院科室ID=D.ID" & _
            " And A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.出院日期 is NULL" & _
            " And A.相关ID=X.ID(+) And A.诊疗项目ID=E.ID And " & str药房置换 & "=F.ID(+)" & _
            " And A.诊疗类别 IN('5','6','7','E') And A.开始执行时间 is Not NULL And A.病人来源<>3" & str期效条件 & str药房条件 & str给药途径 & _
            " Order by D.编码,LPAD(B.出院病床,10,' '),A.婴儿,组号,组ID,A.序号"
        
        On Error GoTo errH
        Set rsSend = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(IIF(strEnd = "", "1990-01-01", strEnd)), Val(Split(str病人IDs, ",")(k)), lng药房ID)
        
        '计算并显示发送清单
        '----------------------------------------------------------------------------------------------------------
        If Not rsSend.EOF Then
            With vsAdvice
                .Redraw = flexRDNone
                For i = 1 To rsSend.RecordCount
                    If rsSend!诊疗类别 = "E" And IsNull(rsSend!相关ID) And rsSend!ID <> Val(.TextMatrix(.Rows - 1, COL_相关ID)) Then
                        GoTo NextLoop '跳过非药正常的治疗医嘱或检验采集方法
                    ElseIf (rsSend!ID = lngDel相关ID Or Nvl(rsSend!相关ID, 0) = lngDel相关ID) And lngDel相关ID <> 0 Then
                        GoTo NextLoop '一并给药或配方中的一个可能已经不能发送,则整组不能发送
                    Else
                        lngDel相关ID = 0
                    End If
                                                    
                    '加入当前行
                    .Rows = .Rows + 1: lngRow = .Rows - 1
                    .Cell(flexcpPictureAlignment, lngRow, COL_选择) = 4
                    Set .Cell(flexcpPicture, lngRow, COL_选择) = img16.ListImages("T").Picture
                    
                    '隐藏附加行
                    If rsSend!诊疗类别 = "7" Then
                        .RowHidden(lngRow) = True '中草药
                    ElseIf rsSend!诊疗类别 = "E" Then
                        If Not IsNull(rsSend!相关ID) Then
                            .RowHidden(lngRow) = True
                            .Cell(flexcpData, lngRow, COL_诊疗类别) = 2 '中药煎法
                        ElseIf Val(.TextMatrix(lngRow - 1, COL_相关ID)) = rsSend!ID _
                            And InStr(",5,6,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 Then
                            .RowHidden(lngRow) = True
                            .Cell(flexcpData, lngRow, COL_诊疗类别) = 1 '给药途径
                        Else
                            .Cell(flexcpData, lngRow, COL_诊疗类别) = 3 '中药用法
                        End If
                    End If
                    
                    '一般列赋值
                    '---------------------------------------------------------------
                    .Cell(flexcpData, lngRow, COL_婴儿) = CLng(Nvl(rsSend!婴儿, 0))
                    If Nvl(rsSend!婴儿, 0) = 0 Then
                        .TextMatrix(lngRow, COL_婴儿) = "病人"
                    Else
                        .TextMatrix(lngRow, COL_婴儿) = "婴儿" & rsSend!婴儿
                        .ColHidden(COL_婴儿) = False '有婴儿医嘱时才显示
                    End If
                    .TextMatrix(lngRow, COL_科室) = rsSend!科室
                    If InStr(str科室 & ",", "," & rsSend!科室 & ",") = 0 Then
                        If str科室 <> "" Then .ColHidden(COL_科室) = False
                        str科室 = str科室 & "," & rsSend!科室
                    End If
                    
                    .TextMatrix(lngRow, COL_病人ID) = rsSend!病人ID
                    .TextMatrix(lngRow, COL_主页ID) = rsSend!主页ID
                    .TextMatrix(lngRow, COL_姓名) = rsSend!姓名
                    .TextMatrix(lngRow, COL_性别) = Nvl(rsSend!性别)
                    .TextMatrix(lngRow, COL_年龄) = Nvl(rsSend!年龄)
                    .TextMatrix(lngRow, COL_住院号) = Nvl(rsSend!住院号)
                    .TextMatrix(lngRow, COL_床号) = Nvl(rsSend!床号)
                    .TextMatrix(lngRow, COL_费别) = Nvl(rsSend!费别)
                    
                    .TextMatrix(lngRow, COL_ID) = rsSend!ID
                    .TextMatrix(lngRow, COL_相关ID) = Nvl(rsSend!相关ID)
                    .TextMatrix(lngRow, COL_诊疗类别) = rsSend!诊疗类别
                    .TextMatrix(lngRow, COL_诊疗项目ID) = rsSend!诊疗项目ID
                    
                    .TextMatrix(lngRow, COL_医嘱内容) = Nvl(rsSend!医嘱内容)
                    .Cell(flexcpData, lngRow, COL_医嘱内容) = CStr(Nvl(rsSend!诊疗项目)) '用于显示计价医嘱
                    
                    .TextMatrix(lngRow, COL_医生嘱托) = Nvl(rsSend!医生嘱托)
                    .TextMatrix(lngRow, COL_执行时间) = Nvl(rsSend!执行时间方案)
                    .TextMatrix(lngRow, COL_频率) = Nvl(rsSend!执行频次)
                    
                    .TextMatrix(lngRow, COL_病人科室ID) = Nvl(rsSend!病人科室ID)
                    .TextMatrix(lngRow, COL_开嘱科室ID) = Nvl(rsSend!开嘱科室ID)
                    .TextMatrix(lngRow, COL_开嘱医生) = Nvl(rsSend!开嘱医生)
                    
                    .TextMatrix(lngRow, COL_计价特性) = Nvl(rsSend!计价特性, 0)
                    .TextMatrix(lngRow, COL_执行性质ID) = Nvl(rsSend!执行性质, 0)
                                                        
                    '显示主要执行科室
                    .TextMatrix(lngRow, COL_执行科室) = Nvl(rsSend!执行科室)
                    
                    '显示附加执行科室
                    If rsSend!诊疗类别 = "E" And IsNull(rsSend!相关ID) Then
                        If InStr(",7,E,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 Then
                            '中药用法
                            .TextMatrix(lngRow, COL_附加执行) = Nvl(rsSend!执行科室)
                            .Cell(flexcpData, lngRow, COL_附加执行) = CStr(Nvl(rsSend!执行科室))
                        ElseIf InStr(",5,6,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 Then
                            '给药途径
                            For j = lngRow - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_相关ID)) = rsSend!ID Then
                                    .TextMatrix(j, COL_附加执行) = Nvl(rsSend!执行科室)
                                    .Cell(flexcpData, j, COL_附加执行) = CStr(Nvl(rsSend!执行科室))
                                Else
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    
                    .TextMatrix(lngRow, COL_执行科室ID) = Nvl(rsSend!执行科室ID)
                                                        
                    '读取药品相关信息
                    '---------------------------------------------------------------
                    If InStr(",5,6,7", rsSend!诊疗类别) > 0 Then
                        Set rsDrug = New ADODB.Recordset
                        '先包括停用药品,待确认要发送的医嘱再检查停用
                        Set rsDrug = GetDrugInfo(rsSend!诊疗项目ID, Nvl(rsSend!收费细目ID, 0), Nvl(rsSend!执行科室ID, 0), 2, False)
                        If rsDrug.EOF Then
                            '药品没有对应的规格信息
                            '删除当前行(及相关行),及处理下一医嘱
                            Call DeleteDrugRow(rsSend, lngRow, lngDel相关ID)
                            lng最小次数 = 0: GoTo NextLoop
                        ElseIf rsDrug.RecordCount > 1 Then
                            '寻找合适的规格
                            Call SeekMatchDrug(rsSend, rsDrug, 0, vBookMark, strTmp)
                            If vBookMark <> 0 Then
                                rsDrug.Bookmark = vBookMark
                            Else
                                rsDrug.MoveFirst
                            End If
                            .Cell(flexcpData, lngRow, COL_规格) = strTmp '可选择的规格
                            '检查全部(指定)规格都停用的药品
                            If .Cell(flexcpData, lngRow, COL_规格) = "" Then
                                Call DeleteDrugRow(rsSend, lngRow, lngDel相关ID)
                                lng最小次数 = 0: GoTo NextLoop
                            End If
                        Else
                            '检查全部(指定)规格都停用的药品
                            If Not (Nvl(rsDrug!撤档时间, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", Nvl(rsDrug!服务对象, 0)) > 0) Then
                                Call DeleteDrugRow(rsSend, lngRow, lngDel相关ID)
                                lng最小次数 = 0: GoTo NextLoop
                            End If
                        End If
                        .TextMatrix(lngRow, COL_规格) = rsDrug!名称 & IIF(Not IsNull(rsDrug!产地), "(" & rsDrug!产地 & ")", "") & IIF(Not IsNull(rsDrug!规格), " " & rsDrug!规格, "")
                        .TextMatrix(lngRow, COL_药品ID) = rsDrug!药品ID
                        .TextMatrix(lngRow, COL_库存) = Format(Nvl(rsDrug!库存, 0), "0.00000") '按住院包装
                        .TextMatrix(lngRow, COL_剂量系数) = Nvl(rsDrug!剂量系数, 1)
                        .TextMatrix(lngRow, COL_住院包装) = Nvl(rsDrug!住院包装, 1)
                        .TextMatrix(lngRow, COL_住院单位) = Nvl(rsDrug!住院单位)
                        .TextMatrix(lngRow, COL_可否分零) = Nvl(rsDrug!可否分零, 0)
                        .TextMatrix(lngRow, COL_药房分批) = Nvl(rsDrug!药房分批, 0)
                        .TextMatrix(lngRow, COL_是否变价) = Nvl(rsDrug!是否变价, 0)
                        
                        '是否存在未确定规格的品种药品
                        If .Cell(flexcpData, lngRow, COL_规格) <> "" Then
                            .Cell(flexcpForeColor, lngRow, COL_规格) = vbBlue '突出显示
                            bln品种药品 = True
                        End If
                    End If
                                                                            
                    '计算发送次数，执行的分解时间，总量
                    '---------------------------------------------------------------
                    If int期效 = 0 Then
                        '长嘱---------------------------------------------
                        If rsSend!诊疗类别 = "7" Then
                            '中药配方不能暂停,长嘱必然有终止时间。配方的末次执行时间应该与终止时间相同
                            .TextMatrix(lngRow, COL_次数) = rsSend!总给予量
                            .TextMatrix(lngRow, COL_分解时间) = Calc次数分解时间(rsSend!总给予量, rsSend!开始执行时间, rsSend!执行终止时间, "", rsSend!执行时间方案, rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位)
                            .TextMatrix(lngRow, COL_首次时间) = Format(Split(.TextMatrix(lngRow, COL_分解时间), ",")(0), "MM-dd HH:mm")
                            .TextMatrix(lngRow, COL_末次时间) = Format(Split(.TextMatrix(lngRow, COL_分解时间), ",")(rsSend!总给予量 - 1), "MM-dd HH:mm")
                            
                            .TextMatrix(lngRow, COL_单量) = Nvl(rsSend!单次用量) '单量
                            .TextMatrix(lngRow, COL_单量单位) = Nvl(rsSend!计算单位)
                            .TextMatrix(lngRow, COL_总量) = rsSend!总给予量 '付数
                            .TextMatrix(lngRow, COL_总量单位) = "付"
                        ElseIf InStr(",5,6,", rsSend!诊疗类别) > 0 Then
                            '当前医嘱的发送计算时间段
                            Call Calc总量次数时间(dbl总量, lng次数, str分解时间, strEnd, rsSend, rsDrug)
                            If str分解时间 = "" Then
                                '无法分解时间(如被暂停的)
                                lngDel相关ID = rsSend!相关ID
                                Call DeleteCurRow(lngRow, rsSend!相关ID)
                                lng最小次数 = 0: GoTo NextLoop
                            End If
                            .TextMatrix(lngRow, COL_次数) = lng次数
                            .TextMatrix(lngRow, COL_分解时间) = str分解时间
                            .TextMatrix(lngRow, COL_首次时间) = Format(Split(str分解时间, ",")(0), "MM-dd HH:mm")
                            .TextMatrix(lngRow, COL_末次时间) = Format(Split(str分解时间, ",")(lng次数 - 1), "MM-dd HH:mm")
                            
                            .TextMatrix(lngRow, COL_单量) = FormatEx(Nvl(rsSend!单次用量), 5)
                            .TextMatrix(lngRow, COL_单量单位) = Nvl(rsSend!计算单位)
                            .TextMatrix(lngRow, COL_总量) = FormatEx(dbl总量, 5)
                            .TextMatrix(lngRow, COL_总量单位) = Nvl(rsDrug!住院单位)
                            
                            If lng次数 < lng最小次数 Or lng最小次数 = 0 Then lng最小次数 = lng次数
                            
                            '当有多个规格可选择时，根据库存是否足够再次定位规格
                            If .Cell(flexcpData, lngRow, COL_规格) <> "" _
                                And Val(.TextMatrix(lngRow, COL_总量)) > Val(.TextMatrix(lngRow, COL_库存)) Then
                                Call SeekMatchDrug(rsSend, rsDrug, Val(.TextMatrix(lngRow, COL_总量)), vBookMark)
                                If vBookMark <> 0 Then
                                    rsDrug.Bookmark = vBookMark
                                    .TextMatrix(lngRow, COL_规格) = rsDrug!名称 & IIF(Not IsNull(rsDrug!产地), "(" & rsDrug!产地 & ")", "") & IIF(Not IsNull(rsDrug!规格), " " & rsDrug!规格, "")
                                    .TextMatrix(lngRow, COL_药品ID) = rsDrug!药品ID
                                    .TextMatrix(lngRow, COL_库存) = Format(Nvl(rsDrug!库存, 0), "0.00000") '按住院包装
                                    .TextMatrix(lngRow, COL_剂量系数) = Nvl(rsDrug!剂量系数, 1)
                                    .TextMatrix(lngRow, COL_住院包装) = Nvl(rsDrug!住院包装, 1)
                                    .TextMatrix(lngRow, COL_住院单位) = Nvl(rsDrug!住院单位)
                                    .TextMatrix(lngRow, COL_药房分批) = Nvl(rsDrug!药房分批, 0)
                                    .TextMatrix(lngRow, COL_是否变价) = Nvl(rsDrug!是否变价, 0)
                                End If
                            End If
                        Else
                            '一并给药的按最小次数发送(影响给药途径计费及上次执行时间)(不分零的可能浪废)
                            If .Cell(flexcpData, lngRow, COL_诊疗类别) = 1 Then '给药途径
                                For j = lngRow - 1 To .FixedRows Step -1
                                    If Val(.TextMatrix(j, COL_相关ID)) = rsSend!ID Then
                                        If Val(.TextMatrix(j, COL_次数)) > lng最小次数 Then
                                            .TextMatrix(j, COL_次数) = lng最小次数
                                            .TextMatrix(j, COL_分解时间) = Trim分解时间(lng最小次数, .TextMatrix(j, COL_分解时间))
                                            .TextMatrix(j, COL_首次时间) = Format(Split(.TextMatrix(j, COL_分解时间), ",")(0), "MM-dd HH:mm")
                                            .TextMatrix(j, COL_末次时间) = Format(Split(.TextMatrix(j, COL_分解时间), ",")(lng最小次数 - 1), "MM-dd HH:mm")
                                        End If
                                    Else
                                        Exit For
                                    End If
                                Next
                                lng最小次数 = 0
                            End If
                            
                            .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngRow - 1, COL_次数) '付数或次数
                            .TextMatrix(lngRow, COL_次数) = .TextMatrix(lngRow - 1, COL_次数)
                            If .Cell(flexcpData, lngRow, COL_诊疗类别) = 3 Then '中药用法
                                .TextMatrix(lngRow, COL_总量单位) = "付"
                            End If
                            
                            .TextMatrix(lngRow, COL_分解时间) = .TextMatrix(lngRow - 1, COL_分解时间)
                            .TextMatrix(lngRow, COL_首次时间) = .TextMatrix(lngRow - 1, COL_首次时间)
                            .TextMatrix(lngRow, COL_末次时间) = .TextMatrix(lngRow - 1, COL_末次时间)
                        End If
                    Else
                        '临嘱---------------------------------------------
                        If rsSend!诊疗类别 = "7" Then
                            .TextMatrix(lngRow, COL_次数) = rsSend!总给予量
                            If Not IsNull(rsSend!执行时间方案) Then
                                .TextMatrix(lngRow, COL_分解时间) = Calc次数分解时间(rsSend!总给予量, rsSend!开始执行时间, CDate("3000-01-01"), "", rsSend!执行时间方案, rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位)
                                .TextMatrix(lngRow, COL_首次时间) = Format(Split(.TextMatrix(lngRow, COL_分解时间), ",")(0), "MM-dd HH:mm")
                                .TextMatrix(lngRow, COL_末次时间) = Format(Split(.TextMatrix(lngRow, COL_分解时间), ",")(rsSend!总给予量 - 1), "MM-dd HH:mm")
                            Else
                                '无分解时间(临嘱可能未输入执行时间而无法分解)
                                '记录费用发生时间(以医嘱开始执行时间)
                                .Cell(flexcpData, lngRow, COL_分解时间) = Format(rsSend!开始执行时间, "yyyy-MM-dd HH:mm:ss")
                            End If
                            .TextMatrix(lngRow, COL_单量) = Nvl(rsSend!单次用量) '单量
                            .TextMatrix(lngRow, COL_单量单位) = Nvl(rsSend!计算单位)
                            .TextMatrix(lngRow, COL_总量) = rsSend!总给予量 '付数
                            .TextMatrix(lngRow, COL_总量单位) = "付"
                        ElseIf InStr(",5,6,", rsSend!诊疗类别) > 0 Then
                            '计算临嘱用药次数
                            If Nvl(rsSend!频率次数, 0) = 0 Or Nvl(rsSend!频率间隔, 0) = 0 Then
                                lng次数 = 1 '设置为一次性的临嘱药品
                            ElseIf Nvl(rsSend!天数, 0) <> 0 And Not IsNull(rsSend!执行频次) Then
                                '一个频率周期的次数
                                If rsSend!间隔单位 = "周" Then
                                    lng次数 = IntEx(rsSend!天数 * (rsSend!频率次数 / 7))
                                ElseIf rsSend!间隔单位 = "天" Then
                                    lng次数 = IntEx(rsSend!天数 * (rsSend!频率次数 / rsSend!频率间隔))
                                ElseIf rsSend!间隔单位 = "小时" Then
                                    lng次数 = IntEx(rsSend!天数 * (rsSend!频率次数 / rsSend!频率间隔) * 24)
                                End If
                            Else
                                '可分零药品时,按总量对单量的倍数计算给药途径的次数,否则按一个频率周期的次数计算
                                If Nvl(rsDrug!可否分零, 0) = 0 And Nvl(rsSend!单次用量, 0) <> 0 Then
                                    lng次数 = IntEx(rsSend!总给予量 * rsDrug!剂量系数 / rsSend!单次用量)
                                Else
                                    lng次数 = Nvl(rsSend!频率次数, 0)
                                End If
                            End If
                            If Not IsNull(rsSend!执行时间方案) Then
                                str分解时间 = Calc次数分解时间(lng次数, rsSend!开始执行时间, CDate("3000-01-01"), "", rsSend!执行时间方案, rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位)
                                If str分解时间 <> "" Then
                                    .TextMatrix(lngRow, COL_分解时间) = str分解时间
                                    .TextMatrix(lngRow, COL_首次时间) = Format(Split(str分解时间, ",")(0), "MM-dd HH:mm")
                                    .TextMatrix(lngRow, COL_末次时间) = Format(Split(str分解时间, ",")(lng次数 - 1), "MM-dd HH:mm")
                                End If
                            Else
                                '无分解时间(临嘱可能未输入执行时间而无法分解)
                                '记录费用发生时间(以医嘱开始执行时间)
                                .Cell(flexcpData, lngRow, COL_分解时间) = Format(rsSend!开始执行时间, "yyyy-MM-dd HH:mm:ss")
                            End If
                            .TextMatrix(lngRow, COL_次数) = lng次数
                            .TextMatrix(lngRow, COL_单量) = FormatEx(Nvl(rsSend!单次用量), 5)
                            .TextMatrix(lngRow, COL_单量单位) = Nvl(rsSend!计算单位)
                            .TextMatrix(lngRow, COL_总量) = FormatEx(rsSend!总给予量 / rsDrug!住院包装, 5) '以住院单位显示
                            .TextMatrix(lngRow, COL_总量单位) = Nvl(rsDrug!住院单位)
                            
                            If lng次数 < lng最小次数 Or lng最小次数 = 0 Then lng最小次数 = lng次数
                        Else
                            '临嘱：一并给药的按最小次数发送(影响给药途径计费)
                            If .Cell(flexcpData, lngRow, COL_诊疗类别) = 1 Then '给药途径
                                For j = lngRow - 1 To .FixedRows Step -1
                                    If Val(.TextMatrix(j, COL_相关ID)) = rsSend!ID Then
                                        If Val(.TextMatrix(j, COL_次数)) > lng最小次数 Then
                                            .TextMatrix(j, COL_次数) = lng最小次数
                                            If .TextMatrix(j, COL_分解时间) <> "" Then
                                                .TextMatrix(j, COL_分解时间) = Trim分解时间(lng最小次数, .TextMatrix(j, COL_分解时间))
                                                .TextMatrix(j, COL_首次时间) = Format(Split(.TextMatrix(j, COL_分解时间), ",")(0), "MM-dd HH:mm")
                                                .TextMatrix(j, COL_末次时间) = Format(Split(.TextMatrix(j, COL_分解时间), ",")(lng最小次数 - 1), "MM-dd HH:mm")
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
                            If .Cell(flexcpData, lngRow, COL_诊疗类别) = 3 Then '中药用法
                                .TextMatrix(lngRow, COL_总量单位) = "付"
                            End If
                            
                            .TextMatrix(lngRow, COL_分解时间) = .TextMatrix(lngRow - 1, COL_分解时间)
                            .Cell(flexcpData, lngRow, COL_分解时间) = .Cell(flexcpData, lngRow - 1, COL_分解时间)
                            .TextMatrix(lngRow, COL_首次时间) = .TextMatrix(lngRow - 1, COL_首次时间)
                            .TextMatrix(lngRow, COL_末次时间) = .TextMatrix(lngRow - 1, COL_末次时间)
                        End If
                    End If
                    
                    '计算项目的金额:用于查看及记帐报警
                    '---------------------------------------------------------------
                    .TextMatrix(lngRow, COL_金额) = Format(Calc医嘱记帐金额(lngRow), gstrDec)
                    
                    '相关行时的一些处理：累计显示金额,给药途径,用法,执行科室,执行性质
                    '---------------------------------------------------------------
                    If InStr(",1,3,", Val(.Cell(flexcpData, lngRow, COL_诊疗类别))) > 0 Then '给药途径或中药用法
                        cur金额 = 0
                        lngTmp = .FindRow(CStr(rsSend!ID), , COL_相关ID)
                        
                        If .Cell(flexcpData, lngRow, COL_诊疗类别) = 1 Then '给药途径
                            '一并给药时,给药途径的金额累加显示在第一个成药中
                            .TextMatrix(lngTmp, COL_金额) = Format(Val(.TextMatrix(lngTmp, COL_金额)) + Val(.TextMatrix(lngRow, COL_金额)), gstrDec)
                            
                            '显示给药途径,执行性质
                            For j = lngTmp To lngRow - 1
                                strTmp = ""
                                If Val(.TextMatrix(j, COL_执行性质ID)) = 5 And Val(.TextMatrix(lngRow, COL_执行性质ID)) <> 5 Then
                                    strTmp = "自备药"
                                ElseIf Val(.TextMatrix(j, COL_执行性质ID)) <> 5 And Val(.TextMatrix(lngRow, COL_执行性质ID)) = 5 Then
                                    strTmp = "离院带药"
                                End If
                                .TextMatrix(j, COL_执行性质) = strTmp
                                .TextMatrix(j, COL_用法) = rsSend!诊疗项目
                            Next
                        Else
                            '药品的执行性质
                            strTmp = ""
                            If Val(.TextMatrix(lngTmp, COL_执行性质ID)) = 5 And Val(.TextMatrix(lngRow, COL_执行性质ID)) <> 5 Then
                                strTmp = "自备药"
                            ElseIf Val(.TextMatrix(lngTmp, COL_执行性质ID)) <> 5 And Val(.TextMatrix(lngRow, COL_执行性质ID)) = 5 Then
                                strTmp = "离院带药"
                            End If
                            
                            '中药用法,煎法
                            str用法 = rsSend!诊疗项目
                            If Val(.Cell(flexcpData, lngRow - 1, COL_诊疗类别)) = 2 Then
                                str用法 = str用法 & "|" & GetItemField("诊疗项目目录", Val(.TextMatrix(lngRow - 1, COL_诊疗项目ID)), "名称")
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
                        
                        '使相关医嘱选择状态相同(固为库存的原因)
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
                    End If
                    
                    '药品库存检查:自备药不检查
                    '---------------------------------------------------------------
                    If InStr(",5,6,7,", rsSend!诊疗类别) > 0 And Nvl(rsSend!执行性质, 0) <> 5 Then
                        Call CheckStock(lngRow, bln库存提示, bln时价提示, bln默认发送)
                    End If
                    
                    '其它处理
                    '---------------------------------------------------------------
                    '病人计数及分隔
                    If rsSend!病人ID <> lng病人ID Then
                        lng病人数 = lng病人数 + 1
                        If lng病人ID <> 0 Then
                            For j = lngRow - 1 To .FixedRows Step -1
                                If Not .RowHidden(j) Then
                                    .CellBorderRange j, .FixedCols, j, .Cols - 1, vbBlack, 0, 0, 0, 2, 0, 0
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    lng病人ID = rsSend!病人ID

NextLoop:           '---------------------------------------------------------------
                    Progress = i / rsSend.RecordCount * 100
                    rsSend.MoveNext
                Next
            End With
        End If
    Next
    
    lblInfo.Caption = lblInfo.Caption & "，共有" & IIF(str科室 = "", " ", "(" & Mid(str科室, 2) & ") ") & lng病人数 & " 个病人的医嘱"
    With vsAdvice
        If Not .ColHidden(COL_规格) Then
            .AutoSize COL_医嘱内容, COL_规格
        Else
            .AutoSize COL_医嘱内容
        End If
        .RowHeight(0) = 320
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        
        .Col = .FixedCols
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) Then
                .Row = i: Exit For
            End If
        Next
        
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    vsAdvice.ColHidden(COL_科室) = True
    vsAdvice.ColHidden(COL_婴儿) = True
    vsAdvice.ColHidden(COL_规格) = Not bln品种药品 Or int期效 = 1
    vsAdvice.ColHidden(COL_首次时间) = int期效 = 1
    vsAdvice.ColHidden(COL_末次时间) = int期效 = 1
    vsAdvice.SetFocus: Call vsAdvice_GotFocus
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

Private Sub InitBillSet()
'功能：初始化医嘱记帐单据生成记录集
    Set mrsBill = New ADODB.Recordset
    mrsBill.Fields.Append "Key", adVarChar, 100
    mrsBill.Fields.Append "NO", adVarChar, 8
    mrsBill.Fields.Append "费用序号", adBigInt
    mrsBill.Fields.Append "发送序号", adBigInt
    mrsBill.CursorLocation = adUseClient
    mrsBill.LockType = adLockOptimistic
    mrsBill.CursorType = adOpenStatic
    mrsBill.Open
End Sub

Private Sub InitRecordSet(rsSQL As ADODB.Recordset, rsTotal As ADODB.Recordset, rsUpload As ADODB.Recordset)
'初始化记录集
    'SQL记录集
    Set rsSQL = New ADODB.Recordset
    rsSQL.Fields.Append "类型", adInteger '1-费用记录,2-医嘱记录,3-发送记录,4-发料记录
    rsSQL.Fields.Append "医嘱ID", adBigInt '一组医嘱的ID
    rsSQL.Fields.Append "项目ID", adBigInt '收费细目ID
    rsSQL.Fields.Append "序号", adBigInt '用于排序
    rsSQL.Fields.Append "SQL", adVarChar, 5000 'SQL
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
    rsUpload.Fields.Append "NO", adVarChar, 10
    rsUpload.CursorLocation = adUseClient
    rsUpload.LockType = adLockOptimistic
    rsUpload.CursorType = adOpenStatic
    rsUpload.Open
End Sub

Private Sub GetCurBillSet(ByVal strKey As String, strNO As String, lng费用序号 As Long, lng发送序号 As Long)
'功能：获取当前记帐单据的NO及序号
'参数：lng费用序号=费用记录中的序号,为-1时表示不取费用序号
'      lng发送序号=发送记录中的序号,为-1时表示不取发送序号
'说明：strKey=根据记帐单据生成规则定的唯一关键字
'1.中西成药按"病人(病人ID,主页ID)_病人科室ID_开嘱科室ID_开嘱医生_执行科室ID"分号。
'2.一个配方中的所有草药分配一个独立单据号
'3.材料医嘱与成药分号规则相同。
'4.其它非药医嘱每条医嘱一个独立单据号(包括给药途径，配方煎法、用法)
'5.检查部位和附加手术,手术麻醉与主要医嘱分配相同单据号。
'6.一并采集的检验组合分配相同的单据号，标本采集方法分配单独的单据号
    mrsBill.Filter = "Key='" & strKey & "'"
    If mrsBill.EOF Then
        mrsBill.AddNew
        mrsBill!Key = strKey
        mrsBill!NO = zlDatabase.GetNextNO(14)
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

Private Function CompletePatiSend(rsPati As ADODB.Recordset, rsSQL As ADODB.Recordset, _
    rsUpload As ADODB.Recordset, ByVal cur合计 As Currency, ByVal str类别 As String, ByVal str类别名称 As String, _
    strWarn As String, intWarn As Integer, blnTran As Boolean) As Boolean
'功能：提交一个病人的医嘱发送数据,在这之前处理记帐报警
'参数：rsPati=包含病人信息的记录集,用于记帐报警
'      rsSQL=包含所有要执行的SQL
'      rsUpload=用于医保上传的记帐单据号
'      cur合计=病人本次要发送医嘱的记帐金额合计,用于记帐报警
'      str类别=病人本次发送记帐费用的收费类别,用于记帐报警
'      str类别=病人本次发送记帐费用的收费类别名称,用于记帐报警
'      strWarn(I/O)=用于记录当前病人已报警类别
'      intWarn(I/O)=用于记录本次发送报警提示时的选择项
'说明：如果出错,则在调用函数中处理,blnTran返回是否启用了事务
    Dim rsWarn As New ADODB.Recordset
    Dim strSQL As String, intR As Integer
    Dim cur当日 As Currency, i As Long
    Dim strMsg As String
    
    '病人费用报警
    If cur合计 > 0 Then
        strSQL = "Select Nvl(适用病人,1) as 适用病人,Nvl(报警方法,1) as 报警方法," & _
            " 报警值,报警标志1,报警标志2,报警标志3 From 记帐报警线" & _
            " Where 病区ID=[1] And Nvl(适用病人,1)=[2]"
        Set rsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsPati!当前病区ID), IIF(Nvl(rsPati!医保, 0) = 1, 2, 1))
        If Not rsWarn.EOF Then
            If rsWarn!报警方法 = 2 Then cur当日 = GetPatiDayMoney(rsPati!病人ID)
            str类别名称 = Mid(str类别名称, 2)
            For i = 1 To Len(str类别)
                intR = BillingWarn(Me, mstrPrivs, rsWarn, rsPati!姓名, Nvl(rsPati!剩余款, 0), cur当日, cur合计, Nvl(rsPati!担保额, 0), Mid(str类别, i, 1), Split(str类别名称, ",")(i - 1), strWarn, intWarn, Nvl(rsPati!医保, 0) = 1)
                If InStr(",2,3,", intR) > 0 Then Exit For
            Next
        End If
    End If
    
    If InStr(",2,3,", intR) = 0 Then
        '执行顺序:费用,医嘱,发送
        '1.先填写费用,因为发送时可能处理费用
        '2.对费用记录按收费细目ID排序插入
        rsSQL.Filter = 0 '上层函数可能使用过,即使没用过也MoveFirst
        rsSQL.Sort = "类型,项目ID,序号"
        rsUpload.Filter = 0 '上层函数可能使用过,即使没用过也MoveFirst
        
        gcnOracle.BeginTrans: blnTran = True
        Do While Not rsSQL.EOF
            Call zlDatabase.ExecuteProcedure(rsSQL!SQL, Me.Caption)
            rsSQL.MoveNext
        Loop
        
        '医保数据上传
        If Not IsNull(rsPati!险类) Then
            If gclsInsure.GetCapability(support医嘱上传, , rsPati!险类) And Not gclsInsure.GetCapability(support记帐完成后上传, , rsPati!险类) Then
                Do While Not rsUpload.EOF
                    strMsg = "" '因为现在一张NO内肯定为一个病人的,所以最后病人参数可以不传
                    If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 1, strMsg, , rsPati!险类) Then
                        '未提交前上传失败则回滚并中止发送
                        If strMsg <> "" Then
                            MsgBox strMsg, vbInformation, gstrSysName '每张提示
                        Else
                            MsgBox rsPati!姓名 & "的费用上传失败，发送操作将被中止。", vbExclamation, gstrSysName
                        End If
                        Exit Function
                    Else
                        If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName '每张提示
                    End If
                    rsUpload.MoveNext
                Loop
            End If
        End If
        gcnOracle.CommitTrans: blnTran = False
        
        '医保数据上传
        If Not IsNull(rsPati!险类) Then
            If gclsInsure.GetCapability(support医嘱上传, , rsPati!险类) And gclsInsure.GetCapability(support记帐完成后上传, , rsPati!险类) Then
                Do While Not rsUpload.EOF
                    strMsg = ""
                    If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 1, strMsg, , rsPati!险类) Then
                        '提交后上传失败,仅提示
                        If strMsg <> "" Then
                            MsgBox strMsg, vbInformation, gstrSysName
                        Else
                            MsgBox rsPati!姓名 & "的记帐单""" & rsUpload!NO & """上传失败，HIS端数据已提交，按确定继续发送。", vbExclamation, gstrSysName
                        End If
                    Else
                        If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                    End If
                    rsUpload.MoveNext
                Loop
            End If
        End If
            
        '提交成功,将病人医嘱行标记为可删除
        With vsAdvice
            i = .FindRow(CStr(rsPati!病人ID), , COL_病人ID)
            For i = i To .Rows - 1
                If Val(.TextMatrix(i, COL_病人ID)) = rsPati!病人ID Then
                    If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                        .RowData(i) = -1
                    End If
                Else
                    Exit For
                End If
            Next
        End With
    End If
    
    CompletePatiSend = True
End Function

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

Public Function SendAdvice() As Long
'功能：处理医嘱发送(这个过程中记帐报警)
'说明：逐个病人发送提交
'返回：如果成功，则返回发送号
    Dim rsPati As New ADODB.Recordset
    Dim rsPrice As New ADODB.Recordset
    Dim rsSQL As ADODB.Recordset
    Dim rsTotal As ADODB.Recordset
    Dim rsUpload As ADODB.Recordset
    
    Dim i As Long, j As Long
    Dim strSQL As String, strTmp As String
    Dim curDate As Date, blnTran As Boolean
    Dim strWarn As String, intWarn As Integer, str类别 As String, str类别名称 As String
    
    Dim lng病人ID As Long, lng发送号 As Long, int计费状态 As Integer, int划价 As Integer, strNO As String
    Dim lng细目ID As Long, lng费用序号 As Long, lng费用父号 As Long, lng发送序号 As Long
    Dim int付数 As Integer, dbl数量 As Double, cur合计 As Currency
    Dim dbl单价 As Double, cur应收 As Currency, cur实收 As Currency
    Dim bln保险项目否 As Boolean, lng保险大类ID As Long, cur统筹金额 As Currency, str保险编码 As String, str费用类型 As String
    Dim str分解时间 As String, str首次时间 As String, str末次时间 As String
    Dim int配方数 As Integer, strNOKey As String, str自动发料 As String
    Dim str发生时间 As String, str登记时间 As String
    Dim dbl发送数次 As Double, blnFirst As Boolean '配方数及分号关键字
    Dim lng病人科室ID As Long, lng执行科室ID As Long, int执行状态 As Integer
    Dim bln离院带药 As Boolean, blnBool As Boolean
    
    Dim blnHaveSub As Boolean, cur医嘱合计 As Currency
    Dim int父序号 As Integer, var父索引 As Variant
    Dim lng父收入ID As Long, str实收 As String
    
    Dim bln药品时价提示 As Boolean, bln药品库存提示 As Boolean, bln药品默认发送 As Boolean
    Dim bln卫材时价提示 As Boolean, bln卫材库存提示 As Boolean, bln卫材默认发送 As Boolean
    
    Dim rsAudit As ADODB.Recordset
    Dim strAudit As String
    
    On Error GoTo errH
    Screen.MousePointer = 11
        
    bln药品时价提示 = True: bln药品库存提示 = True: bln药品默认发送 = True
    bln卫材时价提示 = True: bln卫材库存提示 = True: bln卫材默认发送 = True
    
    Call InitBillSet
    lng发送号 = zlDatabase.GetNextNO(10)
    curDate = zlDatabase.Currentdate
    intWarn = -1 '记帐报警时缺省要提示,与病人无关
    int配方数 = 1 '表示发送的第几付配方,用于分单据号
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                '提交当前病人的数据
                '-----------------------------------------------------------------------------------------
                If Val(.TextMatrix(i, COL_病人ID)) <> lng病人ID Then
                    '提交当前病人数据
                    If lng病人ID <> 0 Then
                        If strAudit <> "" Then
                            MsgBox "病人""" & rsPati!姓名 & """以下费用项目还没有经过审批，对应的医嘱不能发送：" & vbCrLf & strAudit, vbInformation, gstrSysName
                            GoTo errH
                        End If
                    
                        If Not CompletePatiSend(rsPati, rsSQL, rsUpload, cur合计, str类别, str类别名称, strWarn, intWarn, blnTran) Then GoTo errH
                        SendAdvice = lng发送号 '只要提交成功则标注
                    End If
                    
                    '重置病人相关变量
                    str自动发料 = ""
                    lng病人ID = Val(.TextMatrix(i, COL_病人ID))
                    Call InitRecordSet(rsSQL, rsTotal, rsUpload) '重置SQL数组
                    cur合计 = 0:  str类别 = "": str类别名称 = "": strWarn = "" '重置报警变量
                    
                    '获取当前病人信息
                    strSQL = _
                        " Select 病人ID,预交余额,费用余额,0 as 预结费用 From 病人余额 Where 性质=1 And 病人ID=[1]" & _
                        " Union ALL" & _
                        " Select A.病人ID,0,0,Sum(金额) From 保险模拟结算 A,病案主页 B" & _
                        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.险类 Is Not Null And A.病人ID=[1] And A.主页ID=[2] Group by A.病人ID"
                    strSQL = "Select 病人ID,Nvl(Sum(预交余额),0)-Nvl(Sum(费用余额),0)+Nvl(Sum(预结费用),0) as 剩余款 From (" & strSQL & ") Group by 病人ID"
                    
                    strSQL = "Select A.病人ID,B.主页ID,A.姓名,B.险类,B.当前病区ID,B.出院科室ID," & _
                        " D.编码 as 付款码,Decode(D.编码,'1',1,Decode(Nvl(B.险类,0),0,0,1)) as 医保,C.剩余款," & _
                        " Decode(A.担保额,Null,Null,zl_PatientSurety(A.病人ID,B.主页ID)) as 担保额" & _
                        " From 病人信息 A,病案主页 B,(" & strSQL & ") C,医疗付款方式 D" & _
                        " Where A.病人ID=B.病人ID And A.病人ID=C.病人ID(+) And B.医疗付款方式=D.名称(+)" & _
                        " And A.病人ID=[1] And B.主页ID=[2]"
                    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, Val(.TextMatrix(i, COL_主页ID)))
                    
                    '提取当前病人的审批项目清单
                    strAudit = ""
                    If Not IsNull(rsPati!险类) Then
                        Set rsAudit = GetAuditRecord(lng病人ID, Val(.TextMatrix(i, COL_主页ID)))
                    Else
                        Set rsAudit = Nothing '以Nothing为标志该病人不需要判断
                    End If
                    
                    '检查更新并检查当前病人医嘱的药品库存,自备药不检查
                    '虽然提取时已汇总检查，但按品种下时如果改了规格可能发生变化
                    For j = i To .Rows - 1
                        If Val(.TextMatrix(j, COL_病人ID)) = lng病人ID Then
                            '可能根据前面库存检查提示的结果现在已不可用
                            If .Cell(flexcpData, j, COL_选择) = 0 And Not .Cell(flexcpPicture, j, COL_选择) Is Nothing Then
                                If InStr(",5,6,7,", .TextMatrix(j, COL_诊疗类别)) > 0 And Val(.TextMatrix(j, COL_执行性质ID)) <> 5 Then
                                    '在不足禁止的情况下,包括分批或时价药品
                                    If GetStockCheck(Val(.TextMatrix(j, COL_执行科室ID))) = 2 _
                                        Or Val(.TextMatrix(j, COL_药房分批)) = 1 Or Val(.TextMatrix(j, COL_是否变价)) = 1 Then
                                        .TextMatrix(j, COL_库存) = Format(GetStock(Val(.TextMatrix(j, COL_药品ID)), Val(.TextMatrix(j, COL_执行科室ID)), 2), "0.00000")
                                        If CheckStock(j, bln药品库存提示, bln药品时价提示, bln药品默认发送, True) Then
                                            Call RowSelectSame(j, COL_选择)
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
                                    
                '可能根据前面库存检查提示的结果现在已不可用
                If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                    '产生单据号分配关键字
                    '-----------------------------------------------------------------------------------------
                    If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                        '中西成药按"病人(病人ID,主页ID)_病人科室ID_开嘱科室ID_开嘱医生_执行科室ID"分号。
                        strNOKey = "中西成药_" & lng病人ID & "_" & Val(.TextMatrix(i, COL_主页ID)) & "_" & _
                            Val(.TextMatrix(i, COL_病人科室ID)) & "_" & Val(.TextMatrix(i, COL_开嘱科室ID)) & "_" & _
                            .TextMatrix(i, COL_开嘱医生) & "_" & Val(.TextMatrix(i, COL_执行科室ID))
                        '再按要打印的诊疗单据分号
                        strNOKey = strNOKey & "_" & GetClinicBillID(Val(.TextMatrix(i, COL_诊疗项目ID)), 2)
                    ElseIf .TextMatrix(i, COL_诊疗类别) = "7" Then
                        '一个配方中的所有草药分配一个独立单据号
                        strNOKey = "中药配方_" & lng病人ID & "_" & Val(.TextMatrix(i, COL_主页ID)) & "_" & int配方数
                    Else
                        '其它非药医嘱每条医嘱一个独立单据号(包括给药途径，配方煎法、用法)
                        strNOKey = "非药医嘱_" & Val(.TextMatrix(i, COL_ID))
                    End If
                    
                    '是否离院带药
                    bln离院带药 = False
                    If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                        If .TextMatrix(i, COL_执行性质) = "离院带药" Then bln离院带药 = True
                    ElseIf .TextMatrix(i, COL_诊疗类别) = "7" Then
                        j = .FindRow(CStr(.TextMatrix(i, COL_相关ID)), i + 1, COL_ID)
                        If j <> -1 Then
                            If .TextMatrix(j, COL_执行性质) = "离院带药" Then bln离院带药 = True
                        End If
                    End If
                    
                    '产生医嘱记帐费用:以最新价格计算
                    '-----------------------------------------------------------------------------------------
                    strSQL = "": lng细目ID = 0
                    If InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                        '药品缺省固定为正常计价,但下医嘱时指定了为自备药(院外执行)的不读取;药品不可能为叮嘱
                        If Val(.TextMatrix(i, COL_执行性质ID)) <> 5 Then
                            strSQL = _
                                " Select A.ID,A.类别,D.名称 as 类别名称,A.名称,A.计算单位,B.收入项目ID," & _
                                " C.收据费目,Y.住院单位,Y.住院包装,Y.剂量系数,1 as 数量,B.现价 as 单价," & _
                                " A.加班加价,B.加班加价率,A.是否变价,Y.药房分批 as 分批,0 as 跟踪在用," & _
                                " 0 as 从项,[3] as 执行科室ID,A.屏蔽费别,I.要求审批" & _
                                " From 收费项目目录 A,收费价目 B,收入项目 C,收费项目类别 D,药品规格 Y,保险支付项目 I" & _
                                " Where A.ID=B.收费细目ID And B.收入项目ID=C.ID And A.类别=D.编码" & _
                                " And A.ID=Y.药品ID(+) And A.ID=[1] And A.ID=I.收费细目ID(+) And I.险类(+)=[4]" & _
                                " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                                " Order by A.编码"
                        End If
                    ElseIf Val(.TextMatrix(i, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                        '不计价,手工计价；叮嘱,院外执行的医嘱不读取
                        strSQL = _
                            " Select A.ID,A.类别,D.名称 as 类别名称,RTrim(A.名称||' '||A.规格) as 名称," & _
                            " A.计算单位,B.收入项目ID,C.收据费目,Y.住院单位,Y.住院包装,Y.剂量系数,X.数量," & _
                            " Decode(A.是否变价,1,X.单价,B.现价) as 单价,A.加班加价,B.加班加价率,A.是否变价," & _
                            " Decode(A.类别,'4',E.在用分批,Y.药房分批) as 分批,E.跟踪在用,Nvl(X.从项,0) as 从项," & _
                            " Nvl(X.执行科室ID,[3]) as 执行科室ID,A.屏蔽费别,I.要求审批" & _
                            " From 收费项目目录 A,收费价目 B,收入项目 C,收费项目类别 D,材料特性 E,病人医嘱计价 X,药品规格 Y,保险支付项目 I" & _
                            " Where A.ID=B.收费细目ID And B.收入项目ID=C.ID And A.类别=D.编码 And A.ID=E.材料ID(+)" & _
                            " And A.ID=Y.药品ID(+) And X.收费细目ID=A.ID And Nvl(X.数量,0)<>0 And X.医嘱ID=[2]" & _
                            " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                            " And A.ID=I.收费细目ID(+) And I.险类(+)=[4]" & _
                            " Order by 从项,A.ID"
                            '一定要把主项排在前面,以便于计算和在费用记录中保持主从关系
                    End If
                    
                    '汇总折扣变量初始
                    blnHaveSub = False
                    var父索引 = Empty: int父序号 = 0
                    cur医嘱合计 = 0: lng父收入ID = 0
                    
                    int计费状态 = IIF(Val(.TextMatrix(i, COL_计价特性)) = 1, -1, 0) '无需计费或未计费
                    If strSQL <> "" Then
                        Set rsPrice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_药品ID)), Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_执行科室ID)), Val(Nvl(rsPati!险类, 0)))
                        If Not rsPrice.EOF Then
                            int计费状态 = 1 '已计费
                            
                            '确定是否主从关系:即使不汇总折扣,也要记录
                            rsPrice.Filter = "从项=1"
                            If Not rsPrice.EOF Then blnHaveSub = True
                            rsPrice.Filter = 0
                        End If
                        
                        '处理收入项目级的费用明细
                        For j = 1 To rsPrice.RecordCount
                            '检查是否需要和已经审批
                            If Nvl(rsPrice!要求审批, 0) = 1 And Not rsAudit Is Nothing Then
                                rsAudit.Filter = "项目ID=" & rsPrice!ID
                                If rsAudit.EOF Then
                                    If UBound(Split(strAudit, vbCrLf)) < 10 Then
                                        If InStr(strAudit, "●" & rsPrice!名称) = 0 Then
                                            strAudit = strAudit & vbCrLf & "●" & rsPrice!名称
                                        End If
                                    ElseIf UBound(Split(strAudit, vbCrLf)) = 10 Then
                                        strAudit = strAudit & vbCrLf & "… …"
                                    End If
                                End If
                            End If
                        
                            '执行科室ID
                            lng执行科室ID = Nvl(rsPrice!执行科室ID, 0)
                            '在原值基础上取有效的非药嘱药品及跟踪卫材的执行科室
                            If rsPrice!类别 = "4" And Nvl(rsPrice!跟踪在用, 0) = 1 _
                                Or InStr(",5,6,7", rsPrice!类别) > 0 And InStr(",5,6,7", .TextMatrix(i, COL_诊疗类别)) = 0 Then
                                lng病人科室ID = Val(.TextMatrix(i, COL_病人科室ID))
                                lng执行科室ID = Get收费执行科室ID(rsPati!病人ID, rsPati!主页ID, rsPrice!类别, rsPrice!ID, 4, lng病人科室ID, 0, 2, lng执行科室ID)
                                
                                '卫材必须设置执行科室
                                If lng执行科室ID = 0 And rsPrice!类别 = "4" Then
                                    .Row = GetVisibleRow(i, True)
                                    Call .ShowCell(.Row, .Col)
                                    Screen.MousePointer = 0
                                    MsgBox "系统不能为计价卫材""" & rsPrice!名称 & """确定合适的执行科室。" & vbCrLf & _
                                        "请使用计价调整功能人为确定，或到""卫材目录管理""中检查存储库房设置是否正确。", vbInformation, gstrSysName
                                    Call DeleteSendRow: Call ShowSendTotal
                                    Progress = 0: Exit Function
                                End If
                            End If
                            
                            If InStr(",5,6,7", rsPrice!类别) > 0 Then
                                If InStr(",5,6,7", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                                    If .TextMatrix(i, COL_诊疗类别) = "7" Then
                                        int付数 = Val(.TextMatrix(i, COL_总量))
                                        '中药药房单位按不可分零处理:每付
                                        If Val(.TextMatrix(i, COL_可否分零)) = 0 Then
                                            dbl数量 = Val(.TextMatrix(i, COL_单量)) / Nvl(rsPrice!剂量系数, 1)
                                        Else
                                            dbl数量 = IntEx(Val(.TextMatrix(i, COL_单量)) / Nvl(rsPrice!剂量系数, 1) / Nvl(rsPrice!住院包装, 1)) * Nvl(rsPrice!住院包装, 1)
                                        End If
                                    Else
                                        int付数 = 1
                                        dbl数量 = Val(.TextMatrix(i, COL_总量)) * Nvl(rsPrice!住院包装, 1)
                                    End If
                                Else
                                    int付数 = 1
                                    '中药药房单位按不可分零处理:每付
                                    '非药嘱药品计价:因为这里预定了售价数量,因此不作不分零处理
                                    dbl数量 = Val(.TextMatrix(i, COL_总量)) * Nvl(rsPrice!数量, 0)
                                End If
                                dbl数量 = Format(dbl数量, "0.00000")
                                
                                If Nvl(rsPrice!是否变价, 0) = 1 Then
                                    dbl单价 = Format(CalcDrugPrice(rsPrice!ID, lng执行科室ID, int付数 * dbl数量, , True), "0.00000")
                                Else
                                    dbl单价 = Format(Nvl(rsPrice!单价, 0), "0.00000")
                                End If
                            ElseIf rsPrice!类别 = "4" And Nvl(rsPrice!跟踪在用, 0) = 1 Then
                                '检查卫生材料入出类别
                                If mlng卫材类别ID = 0 Then
                                    Screen.MousePointer = 0
                                    MsgBox "不能确定卫生材料单据的入出类别,请先到入出类别管理中设置！", vbInformation, gstrSysName
                                    Call DeleteSendRow: Call ShowSendTotal
                                    Progress = 0: Exit Function
                                End If
                                
                                int付数 = 1
                                dbl数量 = Format(Val(.TextMatrix(i, COL_总量)) * Nvl(rsPrice!数量, 0), "0.00000")
                                
                                '计算时价卫材单价
                                If Nvl(rsPrice!是否变价, 0) = 1 Then
                                    dbl单价 = Format(CalcDrugPrice(rsPrice!ID, lng执行科室ID, dbl数量, , True), "0.00000")
                                Else
                                    dbl单价 = Format(Nvl(rsPrice!单价, 0), "0.00000")
                                End If
                            Else
                                int付数 = 1
                                dbl数量 = Format(Val(.TextMatrix(i, COL_总量)) * Nvl(rsPrice!数量, 0), "0.00000")
                                dbl单价 = Format(Nvl(rsPrice!单价, 0), "0.00000")
                            End If
                            
                            '非药嘱药品及跟踪卫材的库存检查
                            If rsPrice!类别 = "4" And Nvl(rsPrice!跟踪在用, 0) = 1 _
                                Or InStr(",5,6,7", rsPrice!类别) > 0 And InStr(",5,6,7", .TextMatrix(i, COL_诊疗类别)) = 0 Then
                                If GetStockCheck(lng执行科室ID) <> 0 Or Nvl(rsPrice!是否变价, 0) = 1 Or Nvl(rsPrice!分批, 0) = 1 Then
                                    If rsPrice!类别 = "4" Then
                                        blnBool = CheckPriceStock(i, rsPrice, lng执行科室ID, int付数 * dbl数量, rsTotal, bln卫材库存提示, bln卫材时价提示, bln卫材默认发送)
                                    Else
                                        blnBool = CheckPriceStock(i, rsPrice, lng执行科室ID, int付数 * dbl数量, rsTotal, bln药品库存提示, bln药品时价提示, bln药品默认发送)
                                    End If
                                    If blnBool Then
                                        Call RowSelectSame(i, COL_选择, rsSQL, rsTotal, rsUpload)
                                        GoTo NextAdvice
                                    End If
                                End If
                            End If
                            
                            '发送金额
                            cur应收 = int付数 * dbl数量 * dbl单价
                            
                            '处理加班加价
                            If gbln加班加价 And Nvl(rsPrice!加班加价, 0) = 1 Then
                                cur应收 = Format(cur应收 * (1 + Nvl(rsPrice!加班加价率, 0) / 100), gstrDec)
                            Else
                                cur应收 = Format(cur应收, gstrDec)
                            End If
                            
                            '计算汇总折扣合计
                            If gbln从项汇总折扣 And blnHaveSub Then
                                cur实收 = cur应收
                                cur医嘱合计 = cur医嘱合计 + cur实收
                            ElseIf Nvl(rsPrice!屏蔽费别, 0) = 0 Then
                                cur实收 = Format(ActualMoney(.TextMatrix(i, COL_费别), rsPrice!收入项目ID, cur应收, rsPrice!ID, lng执行科室ID, _
                                    int付数 * dbl数量, IIF(gbln加班加价 And Nvl(rsPrice!加班加价, 0) = 1, Nvl(rsPrice!加班加价率, 0) / 100, 0)), gstrDec)
                            Else
                                cur实收 = cur应收
                            End If
                                
                            '医保相关字段
                            bln保险项目否 = False: lng保险大类ID = 0: cur统筹金额 = 0: str保险编码 = "": str费用类型 = ""
                            If Not IsNull(rsPati!险类) Then
                                strTmp = gclsInsure.GetItemInsure(lng病人ID, rsPrice!ID, cur实收, False, rsPati!险类)
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
                            If InStr(str类别, rsPrice!类别) = 0 Then
                                str类别 = str类别 & rsPrice!类别
                                str类别名称 = str类别名称 & "," & rsPrice!类别名称
                            End If
                                                        
                            'NO,序号
                            Call GetCurBillSet(strNOKey, strNO, lng费用序号, -1)
                            rsSQL.AddNew: blnBool = False
                            If rsPrice!ID <> lng细目ID Then
                                lng费用父号 = lng费用序号
                                '主从关系时，记录主项信息
                                If rsPrice!从项 = 0 And blnHaveSub Then
                                    int父序号 = lng费用序号
                                    lng父收入ID = rsPrice!收入项目ID
                                    var父索引 = rsSQL.Bookmark
                                    blnBool = True
                                End If
                            End If
                            lng细目ID = rsPrice!ID
                            
                            '汇总折扣时，对主项的实收金额作特殊处理
                            If gbln从项汇总折扣 And blnHaveSub And blnBool Then
                                str实收 = Chr(0) & Chr(1) & "Begin" & cur实收 & "End" & Chr(0) & Chr(1)
                            Else
                                str实收 = cur实收
                            End If
                            
                            '是否划价
                            If InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                                int划价 = IIF(InStr(gstr发送划价单, "5") > 0, 1, 0)
                            Else
                                int划价 = IIF(InStr(gstr发送划价单, .TextMatrix(i, COL_诊疗类别)) > 0, 1, 0)
                            End If
                            
                            '发生时间
                            If .TextMatrix(i, COL_分解时间) <> "" Then
                                str发生时间 = "To_Date('" & Split(.TextMatrix(i, COL_分解时间), ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                str发生时间 = "To_Date('" & .Cell(flexcpData, i, COL_分解时间) & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            
                            '登记时间
                            If int划价 = 1 Then '与非划价的时间上区分开
                                str登记时间 = "To_Date('" & Format(DateAdd("s", 1, curDate), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                str登记时间 = "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
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
                            
                            '因为现在不计价的医嘱不产生费用,所以传入的计价特性都为(0-正常计价)
                            rsSQL!类型 = 1
                            rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                            rsSQL!项目ID = rsPrice!ID
                            rsSQL!序号 = i
                            rsSQL!SQL = "ZL_住院记帐记录_Insert(" & _
                                "'" & strNO & "'," & lng费用序号 & "," & lng病人ID & "," & ZVal(.TextMatrix(i, COL_主页ID)) & "," & _
                                ZVal(.TextMatrix(i, COL_住院号)) & ",'" & .TextMatrix(i, COL_姓名) & "'," & _
                                "'" & .TextMatrix(i, COL_性别) & "','" & .TextMatrix(i, COL_年龄) & "'," & _
                                "'" & .TextMatrix(i, COL_床号) & "','" & .TextMatrix(i, COL_费别) & "'," & _
                                rsPati!当前病区ID & "," & rsPati!出院科室ID & ",0," & Val(.Cell(flexcpData, i, COL_婴儿)) & "," & _
                                ZVal(.TextMatrix(i, COL_开嘱科室ID)) & ",'" & .TextMatrix(i, COL_开嘱医生) & "'," & _
                                IIF(rsPrice!从项 = 1, ZVal(int父序号), "NULL") & "," & rsPrice!ID & "," & _
                                "'" & rsPrice!类别 & "','" & Nvl(rsPrice!计算单位) & "'," & _
                                IIF(bln保险项目否, 1, 0) & "," & ZVal(lng保险大类ID) & ",'" & str保险编码 & "'," & _
                                int付数 & "," & dbl数量 & ",NULL," & ZVal(lng执行科室ID) & "," & _
                                IIF(lng费用父号 = lng费用序号, "NULL", lng费用父号) & "," & rsPrice!收入项目ID & "," & _
                                "'" & Nvl(rsPrice!收据费目) & "'," & dbl单价 & "," & cur应收 & "," & str实收 & "," & _
                                cur统筹金额 & "," & str发生时间 & "," & str登记时间 & "," & _
                                "'医嘱发送'," & int划价 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "',0," & _
                                IIF(rsPrice!类别 = "4", mlng卫材类别ID, mlng药品类别ID) & "," & _
                                "NULL,'" & .TextMatrix(i, COL_医嘱内容) & "',NULL," & Val(.TextMatrix(i, COL_ID)) & "," & _
                                "'" & .TextMatrix(i, COL_频率) & "'," & ZVal(.TextMatrix(i, COL_单量)) & "," & _
                                "'" & .TextMatrix(i, COL_用法) & "'," & mint期效 & "," & _
                                IIF(bln离院带药, 3, Val(.TextMatrix(i, COL_计价特性))) & ",Null,'" & str费用类型 & "')"
                            rsSQL.Update
                            
                            '记录自动发料的SQL
                            If gbln住院自动发料 And int划价 = 0 And lng执行科室ID <> 0 And rsPrice!类别 = "4" And Nvl(rsPrice!跟踪在用, 0) = 1 Then
                                If InStr(str自动发料 & ";", ";" & strNO & "," & lng执行科室ID & ";") = 0 Then
                                    rsSQL.AddNew
                                    rsSQL!类型 = 4
                                    rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                                    rsSQL!项目ID = 0
                                    rsSQL!序号 = i
                                    rsSQL!SQL = "zl_材料收发记录_处方发料(" & lng执行科室ID & ",25,'" & strNO & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "','" & UserInfo.姓名 & "',1,Sysdate)"
                                    rsSQL.Update
                                    str自动发料 = str自动发料 & ";" & strNO & "," & lng执行科室ID
                                End If
                            End If
                            
                            rsPrice.MoveNext
                        Next
                    End If
                    
                    '对医嘱金额进行汇总折扣处理:主项汇总打折不支持按成本加收打折
                    If gbln从项汇总折扣 And blnHaveSub And var父索引 <> Empty And lng父收入ID <> 0 Then
                        rsSQL.Bookmark = var父索引
                        cur实收 = Format(ActualMoney(.TextMatrix(i, COL_费别), lng父收入ID, cur医嘱合计), gstrDec)
                        cur实收 = cur实收 - cur医嘱合计 '打折差额
                        cur实收 = Get实收金额(rsSQL!SQL) + cur实收
                        rsSQL!SQL = Set实收金额(rsSQL!SQL, cur实收)
                        rsSQL.Update
                    End If
                    
                    '更改医嘱的执行科室
                    If .Cell(flexcpData, i, COL_执行科室ID) = 1 Then
                        rsSQL.AddNew
                        rsSQL!类型 = 2
                        rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                        rsSQL!项目ID = 0
                        rsSQL!序号 = i
                        rsSQL!SQL = "ZL_医嘱执行科室_Update(" & Val(.TextMatrix(i, COL_ID)) & "," & ZVal(.TextMatrix(i, COL_执行科室ID)) & ")"
                        rsSQL.Update
                    End If
                    
                    '产生医嘱发送记录
                    '-----------------------------------------------------------------------------------------
                    If Val(.TextMatrix(i, COL_执行性质ID)) <> 0 Then '叮嘱不发送(给药途径，配方煎法、用法可能为)
                        '一样要产生费用NO
                        Call GetCurBillSet(strNOKey, strNO, -1, lng发送序号)
                                                                
                        '是否一组医嘱的第一医嘱行
                        blnFirst = False
                        If InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                            If Val(.TextMatrix(i, COL_相关ID)) <> Val(.TextMatrix(i - 1, COL_相关ID)) Then
                                blnFirst = True '药疗发送时,只有第一药品行才为第一医嘱行
                            End If
                        End If
                        
                        '执行状态
                        int执行状态 = 0
                        If mblnAutoExe And InStr(",5,6,7,", .TextMatrix(i, COL_诊疗类别)) = 0 Then
                            If Val(.TextMatrix(i, COL_开嘱科室ID)) = Val(.TextMatrix(i, COL_执行科室ID)) Then
                                int执行状态 = 1
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
                                                                
                        '首末时间
                        str分解时间 = .TextMatrix(i, COL_分解时间)
                        If str分解时间 <> "" Then
                            str首次时间 = "To_Date('" & Split(str分解时间, ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                            str末次时间 = "To_Date('" & Split(str分解时间, ",")(Val(.TextMatrix(i, COL_次数)) - 1) & "','YYYY-MM-DD HH24:MI:SS')"
                        Else
                            str首次时间 = "NULL"
                            str末次时间 = "NULL"
                        End If
                        
                        rsSQL.AddNew
                        rsSQL!类型 = 3
                        rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                        rsSQL!项目ID = 0
                        rsSQL!序号 = i
                        rsSQL!SQL = "ZL_病人医嘱发送_Insert(" & _
                            Val(.TextMatrix(i, COL_ID)) & "," & lng发送号 & ",2,'" & strNO & "'," & _
                            lng发送序号 & "," & dbl发送数次 & "," & str首次时间 & "," & str末次时间 & "," & _
                            "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                            int执行状态 & "," & ZVal(.TextMatrix(i, COL_执行科室ID)) & "," & int计费状态 & "," & IIF(blnFirst, 1, 0) & ")"
                        rsSQL.Update
                    End If
                    
                    '计算中药配方数
                    If .Cell(flexcpData, i, COL_诊疗类别) = 3 Then '中药用法
                        int配方数 = int配方数 + 1
                    End If
                End If
            End If
NextAdvice:
            '----------------------------------------
            Progress = (i - .FixedRows + 1) / (.Rows - .FixedRows) * 100
        Next
        
        '提交最后一个病人的数据
        '-----------------------------------------------------------------------------------------
        If lng病人ID <> 0 Then
            If strAudit <> "" Then
                MsgBox "病人""" & rsPati!姓名 & """以下费用项目还没有经过审批，对应的医嘱不能发送：" & vbCrLf & strAudit, vbInformation, gstrSysName
                GoTo errH
            End If
        
            If Not CompletePatiSend(rsPati, rsSQL, rsUpload, cur合计, str类别, str类别名称, strWarn, intWarn, blnTran) Then GoTo errH
            SendAdvice = lng发送号 '只要提交成功则标注
        End If
        
    End With
    '删除所有已成功发送的行
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0: Screen.MousePointer = 0
    SendAdvice = lng发送号
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTran Then gcnOracle.RollbackTrans
    If Err.Number <> 0 Then '如医保上传失败退出没有错误
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0
End Function

Private Sub ShowSendTotal()
'功能：根据当前选择要发送的医嘱，计算并显示发送的医嘱合计
    Dim curTotal As Currency, i As Long
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If Not .RowHidden(i) And .Cell(flexcpData, i, COL_选择) = 0 _
                And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                curTotal = curTotal + Val(.TextMatrix(i, COL_金额))
            End If
        Next
    End With
    stbThis.Panels(3).Text = "发送费用：" & Format(curTotal, gstrDec)
    Call Form_Resize
End Sub

Private Sub SetDeptInput(ByVal lngRow As Long, ByVal lngCol As Long, rsInput As ADODB.Recordset)
'功能：设置执行科室输入的的值
    Dim i As Long
        
    With vsAdvice
        If lngCol = COL_附加执行 Then
            '更改显示行的附加执行科室显示
            .TextMatrix(lngRow, COL_附加执行) = rsInput!名称
            .Cell(flexcpData, lngRow, COL_附加执行) = .TextMatrix(lngRow, COL_附加执行)
            
            '更改附加项目行的执行科室
            If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                '给药途径
                i = .FindRow(CStr(.TextMatrix(lngRow, COL_相关ID)), lngRow + 1, COL_ID)
                .TextMatrix(i, COL_执行科室ID) = rsInput!ID
                .Cell(flexcpData, i, COL_执行科室ID) = 1
                
                For i = lngRow + 1 To .Rows - 1
                    If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                        .TextMatrix(i, COL_附加执行) = rsInput!名称
                        .Cell(flexcpData, i, COL_附加执行) = .TextMatrix(lngRow, COL_附加执行)
                    Else
                        Exit For
                    End If
                Next
            ElseIf .TextMatrix(lngRow, COL_诊疗类别) = "E" And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 _
                And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) _
                And InStr(",7,E,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 Then
                '中药用法
                .TextMatrix(lngRow, COL_执行科室ID) = rsInput!ID
                .Cell(flexcpData, lngRow, COL_执行科室ID) = 1
            End If
        End If
    End With
End Sub
