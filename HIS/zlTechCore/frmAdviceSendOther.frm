VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmAdviceSendOther 
   AutoRedraw      =   -1  'True
   Caption         =   "其它医嘱发送"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   Icon            =   "frmAdviceSendOther.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmAdviceSendOther.frx":038A
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
            Picture         =   "frmAdviceSendOther.frx":0914
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
      FormatString    =   $"frmAdviceSendOther.frx":11A8
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
      FormatString    =   $"frmAdviceSendOther.frx":1243
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
               Picture         =   "frmAdviceSendOther.frx":12DE
               Key             =   "T"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceSendOther.frx":1878
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
            Picture         =   "frmAdviceSendOther.frx":1E12
            Key             =   "全选"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendOther.frx":202C
            Key             =   "全清"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendOther.frx":2246
            Key             =   "发送"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendOther.frx":2460
            Key             =   "重置"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendOther.frx":267A
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendOther.frx":2894
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
            Picture         =   "frmAdviceSendOther.frx":2AAE
            Key             =   "全选"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendOther.frx":2CC8
            Key             =   "全清"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendOther.frx":2EE2
            Key             =   "发送"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendOther.frx":30FC
            Key             =   "重置"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendOther.frx":3316
            Key             =   "帮助"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendOther.frx":3530
            Key             =   "退出"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAdviceSendOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String 'IN
Public mlng病区ID As Long 'IN:用于记录主界面的病区及上次发送病区
Public mlng病人ID As Long 'IN
Public mblnSend As Boolean 'OUT:是否成功发送过。
Public mblnRefresh As Boolean 'OUT:发送后是否需要刷新主界面

Private mcolStock As Collection '存放各个药品库房的出库检查方式
Private mrsBill As ADODB.Recordset
Private mstrEnd As String '本次发送的结束时间
Private mint期效 As Integer '本次发送的医嘱期效
Private mlng药品类别ID As Long '药品入出类别ID
Private mlng卫材类别ID As Long
Private mblnAutoExe As Boolean
Private mstrLike As String
Private mblnFirst As Boolean
Private mstrRollNotify As String
'----------------------------------------------
Private Const COL_选择 = 0
Private Const COL_科室 = 1
Private Const COL_姓名 = 2
Private Const COL_住院号 = 3
Private Const COL_床号 = 4
Private Const COL_费别 = 5
Private Const COL_婴儿 = 6
Private Const COL_医嘱内容 = 7
Private Const COL_总量 = 8
Private Const COL_总量单位 = 9
Private Const COL_单量 = 10
Private Const COL_单量单位 = 11
Private Const COL_金额 = 12
Private Const COL_频率 = 13
Private Const COL_医生嘱托 = 14
Private Const COL_执行科室 = 15
Private Const COL_附加执行 = 16
Private Const COL_执行时间 = 17
Private Const COL_首次时间 = 18
Private Const COL_末次时间 = 19
Private Const COL_病人ID = 20 '隐藏列
Private Const COL_主页ID = 21
Private Const COL_性别 = 22
Private Const COL_年龄 = 23
Private Const COL_ID = 24
Private Const COL_相关ID = 25
Private Const COL_病人科室ID = 26
Private Const COL_开嘱科室ID = 27
Private Const COL_开嘱医生 = 28
Private Const COL_诊疗类别 = 29
Private Const COL_诊疗项目ID = 30
Private Const COL_计价特性 = 31
Private Const COL_操作类型 = 32
Private Const COL_执行性质ID = 33
Private Const COL_执行科室ID = 34
Private Const COL_次数 = 35
Private Const COL_分解时间 = 36
'-------------------------------------------------
Private Const COLP_计价医嘱 = 0
Private Const COLP_类别 = 1
Private Const COLP_收费项目 = 2
Private Const COLP_数量 = 3
Private Const COLP_单位 = 4
Private Const COLP_单价 = 5
Private Const COLP_应收金额 = 6
Private Const COLP_实收金额 = 7
Private Const COLP_执行科室 = 8
Private Const COLP_费用类型 = 9
Private Const COLP_从项 = 10

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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call InitAdviceTable
    Call InitPriceTable
    Call RestoreWinState(Me, App.ProductName)
    
    mstrLike = IIF(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
    mblnAutoExe = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "本科执行自动完成", 0)) <> 0
    mblnSend = False
    mblnRefresh = False
    mblnFirst = True
    
    '各个库房药品出库检查方式,包括发料部门
    Set mcolStock = InitStockCheck(2, True)
End Sub

Private Function GetStockCheck(ByVal lng库房ID As Long) As Integer
'功能：获取指定库房的出库库存检查方式
    Dim intStyle As Integer
    On Error Resume Next
    intStyle = mcolStock("_" & lng库房ID)
    Err.Clear: On Error GoTo 0
    GetStockCheck = intStyle
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
    mlng药品类别ID = 0
    mlng卫材类别ID = 0
    Set mrsBill = Nothing
    Set mcolStock = Nothing
    
    gbln加班加价 = False
End Sub

Private Function ResetSend() As Boolean
'功能：重置发送条件
    With frmAdviceSendOtherCond
        .mstrPrivs = mstrPrivs
        .mlng病区ID = mlng病区ID
        .mlng病人ID = mlng病人ID
        .Show 1, Me
        If .mblnOK Then
            mlng病区ID = .mlng病区ID
            mstrEnd = .mstrEnd
            mint期效 = .mint期效
            Call LoadAdviceSend(.mstrEnd, .mint期效, .mlng执行科室ID, .mstr病人IDs, .mstr类别s)
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
            
            lng发送号 = SendAdvice
            If lng发送号 <> 0 Then
                '发送了特殊医嘱时检查并提醒超期收回(自动)停止的医嘱
                If mstrRollNotify <> "" Then
                    Call ShowRollNotify
                End If
                
                mblnSend = True
                '打印诊疗单据
                Call frmSendBillPrint.ShowMe(lng发送号, 2, Me)
            End If
        Case "重置"
            Call ResetSend
        Case "帮助"
            ShowHelp App.ProductName, Me.Hwnd, "frmAdviceSendDrug"
        Case "退出"
            Unload Me
    End Select
End Sub

Private Sub ShowRollNotify()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String

    On Error GoTo errH
    
    '条件与超期收回中一致，但只包含当前状态为(自动)停止的。
    strSQL = "(A.执行时间方案 is NULL And (Nvl(A.频率次数,0)=0 Or Nvl(A.频率间隔,0)=0 Or A.频率间隔 is NULL))"
    strSQL = _
        " Select C.姓名,A.医嘱内容 From 病人医嘱记录 A,病人信息 C,诊疗项目目录 E" & _
        " Where A.诊疗项目ID=E.ID And A.病人ID=C.病人ID" & _
        " And (A.病人ID,A.主页ID) IN(" & Mid(mstrRollNotify, 2) & ")" & _
        " And Not(A.诊疗类别='H' And E.操作类型='1') And Not(A.诊疗类别='Z' And E.操作类型='4')" & _
        " And Nvl(A.执行性质,0)<>0 And A.总给予量 is NULL And Nvl(A.医嘱期效,0)=0" & _
        " And ((Not " & strSQL & " And A.执行终止时间<A.上次执行时间)" & _
        " Or (" & strSQL & " And Trunc(A.执行终止时间)<Trunc(A.上次执行时间)+1))" & _
        " And A.医嘱状态=8 And (A.相关ID is Null Or A.诊疗类别 IN('5','6'))" & _
        " And A.开始执行时间 is Not NULL And A.病人来源<>3" & _
        " And Not Exists(" & _
            " Select ID From 病人医嘱记录 X" & _
            " Where 诊疗类别 IN('5','6') And X.相关ID=A.ID" & _
            " And (病人ID,主页ID) IN(" & Mid(mstrRollNotify, 2) & "))" & _
        " Order by A.病人ID,A.序号"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        strMsg = strMsg & vbCrLf & "●　病人：" & rsTmp!姓名 & "　医嘱：" & rsTmp!医嘱内容
        rsTmp.MoveNext
    Loop
    If strMsg <> "" Then
        MsgBox "下列已停止的医嘱被超期发送：" & vbCrLf & strMsg & vbCrLf & vbCrLf & "该类医嘱可以在护士工作站中使用""超期发送收回""进行处理。", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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

Private Function GetVisibleRow(ByVal lngRow As Long) As Long
'功能：根据指定医嘱行，返回该医嘱中可见的行
    Dim lng组ID As Long, i As Long
    
    GetVisibleRow = lngRow
    
    With vsAdvice
        If Not .RowHidden(lngRow) Then Exit Function
        
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
'功能：显示当前发送医嘱行的记帐费用信息(按费别打折)
    Dim rsTmp As New ADODB.Recordset
    Dim bln附加手术 As Boolean, strSQL As String, i As Long
    Dim str费别 As String, str行号 As String, strTmp As String
    Dim dbl单价 As Double, cur应收 As Currency, cur实收 As Currency
    Dim dbl当前单价 As Double, cur当前应收 As Currency, cur当前实收 As Currency
    Dim lng病人科室ID As Long, lng执行科室ID As Long
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
        
        '不计价,手工计价,叮嘱,院外执行的不读取计价
        If .TextMatrix(lngRow, COL_诊疗类别) = "E" And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
            '检验组合
            If Val(.TextMatrix(lngRow, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质ID))) = 0 Then
                strTmp = "采集方法-" & Replace(.Cell(flexcpData, lngRow, COL_医嘱内容), "'", "''")
                strSQL = _
                    "Select " & lngRow & " as 行号,'" & strTmp & "' as 计价医嘱," & _
                    " B.ID,B.类别,B.名称,B.计算单位 as 单位,0 as 附加手术,B.是否变价,B.加班加价,A.单价," & _
                    " Nvl(A.数量,0)*" & Val(.TextMatrix(lngRow, COL_总量)) & " as 数量," & _
                    " Nvl(A.执行科室ID," & Val(.TextMatrix(lngRow, COL_执行科室ID)) & ") as 执行科室ID," & _
                    " B.费用类型,B.屏蔽费别,Nvl(A.从项,0) as 从项" & _
                    " From 病人医嘱计价 A,收费项目目录 B" & _
                    " Where A.收费细目ID=B.ID And Nvl(A.数量,0)<>0 And A.医嘱ID=" & Val(.TextMatrix(lngRow, COL_ID))
            End If
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If Val(.TextMatrix(i, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                        strTmp = "检验项目-" & Replace(.Cell(flexcpData, i, COL_医嘱内容), "'", "''")
                        strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                            "Select " & i & " as 行号,'" & strTmp & "' as 计价医嘱,B.ID,B.类别,B.名称," & _
                            " B.计算单位 as 单位,0 as 附加手术,B.是否变价,B.加班加价,A.单价," & _
                            " Nvl(A.数量,0)*" & Val(.TextMatrix(i, COL_总量)) & " as 数量," & _
                            " Nvl(A.执行科室ID," & Val(.TextMatrix(i, COL_执行科室ID)) & ") as 执行科室ID," & _
                            " B.费用类型,B.屏蔽费别,Nvl(A.从项,0) as 从项" & _
                            " From 病人医嘱计价 A,收费项目目录 B" & _
                            " Where A.收费细目ID=B.ID And Nvl(A.数量,0)<>0 And A.医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                    End If
                Else
                    Exit For
                End If
            Next
        Else
            If Val(.TextMatrix(lngRow, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质ID))) = 0 Then
                strTmp = .Cell(flexcpData, lngRow, COL_诊疗类别) & "医嘱-" & Replace(.Cell(flexcpData, lngRow, COL_医嘱内容), "'", "''")
                strSQL = _
                    "Select " & lngRow & " as 行号,'" & strTmp & "' as 计价医嘱," & _
                    " B.ID,B.类别,B.名称,B.计算单位 as 单位,0 as 附加手术,B.是否变价,B.加班加价,A.单价," & _
                    " Nvl(A.数量,0)*" & Val(.TextMatrix(lngRow, COL_总量)) & " as 数量," & _
                    " Nvl(A.执行科室ID," & Val(.TextMatrix(lngRow, COL_执行科室ID)) & ") as 执行科室ID," & _
                    " B.费用类型,B.屏蔽费别,Nvl(A.从项,0) as 从项" & _
                    " From 病人医嘱计价 A,收费项目目录 B" & _
                    " Where A.收费细目ID=B.ID And Nvl(A.数量,0)<>0 And A.医嘱ID=" & Val(.TextMatrix(lngRow, COL_ID))
            End If
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If Val(.TextMatrix(i, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                        bln附加手术 = False
                        If .TextMatrix(i, COL_诊疗类别) = "F" Then
                            bln附加手术 = True
                            strTmp = "附加手术-" & Replace(.Cell(flexcpData, i, COL_医嘱内容), "'", "''")
                        ElseIf .TextMatrix(i, COL_诊疗类别) = "G" Then
                            strTmp = "手术麻醉-" & Replace(.Cell(flexcpData, i, COL_医嘱内容), "'", "''")
                        ElseIf .TextMatrix(i, COL_诊疗类别) = "D" Then
                            strTmp = "检查部位-" & Replace(.Cell(flexcpData, i, COL_医嘱内容), "'", "''")
                        End If
                        
                        strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                            "Select " & i & " as 行号,'" & strTmp & "' as 计价医嘱," & _
                            " B.ID,B.类别,B.名称,B.计算单位 as 单位," & IIF(bln附加手术, 1, 0) & " as 附加手术," & _
                            " B.是否变价,B.加班加价,A.单价,Nvl(A.数量,0)*" & Val(.TextMatrix(i, COL_总量)) & " as 数量," & _
                            " Nvl(A.执行科室ID," & Val(.TextMatrix(i, COL_执行科室ID)) & ") as 执行科室ID," & _
                            " B.费用类型,B.屏蔽费别,Nvl(A.从项,0) as 从项" & _
                            " From 病人医嘱计价 A,收费项目目录 B" & _
                            " Where A.收费细目ID=B.ID And Nvl(A.数量,0)<>0 And A.医嘱ID=" & Val(.TextMatrix(i, COL_ID))
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    
    With vsPrice
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If strSQL <> "" Then
            '以最新价格计算
            strSQL = "Select A.行号," & _
                " B.收入项目ID,A.计价医嘱,A.ID,A.从项,A.类别,C.名称 as 类别名称,A.名称,A.单位,A.屏蔽费别," & _
                " A.执行科室ID,F.名称 as 执行科室,A.费用类型,E.跟踪在用,D.住院单位,A.数量,A.附加手术,B.附术收费率," & _
                " D.住院包装,A.是否变价,A.加班加价,B.加班加价率,Decode(Nvl(A.是否变价,0),1,A.单价,B.现价) as 单价" & _
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
                    
                    .TextMatrix(.Rows - 1, COLP_计价医嘱) = rsTmp!计价医嘱
                    .TextMatrix(.Rows - 1, COLP_类别) = rsTmp!类别名称
                    .TextMatrix(.Rows - 1, COLP_收费项目) = rsTmp!名称
                    .TextMatrix(.Rows - 1, COLP_费用类型) = Nvl(rsTmp!费用类型)
                    .TextMatrix(.Rows - 1, COLP_从项) = IIF(Nvl(rsTmp!从项, 0) = 0, "", "√")
                    
                    If InStr(",5,6,7,", rsTmp!类别) > 0 Then
                        .TextMatrix(.Rows - 1, COLP_单位) = Nvl(rsTmp!住院单位)
                        .TextMatrix(.Rows - 1, COLP_数量) = FormatEx(Nvl(rsTmp!数量, 0) / Nvl(rsTmp!住院包装, 1), 5)
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
                    If rsTmp!类别 = "4" And Nvl(rsTmp!跟踪在用, 0) = 1 Or InStr(",5,6,7,", rsTmp!类别) > 0 Then
                        lng病人科室ID = Val(vsAdvice.TextMatrix(lngRow, COL_病人科室ID))
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

                '单价计算处理
                If InStr(",5,6,7,", rsTmp!类别) > 0 Then
                    If Nvl(rsTmp!是否变价, 0) = 0 Then
                        dbl当前单价 = Nvl(rsTmp!单价, 0)
                    Else
                        '非药医嘱对应的药品时价计价
                        dbl当前单价 = CalcDrugPrice(rsTmp!ID, Val(.Cell(flexcpData, .Rows - 1, COLP_执行科室)), Nvl(rsTmp!数量, 0), , True)
                    End If
                    cur当前应收 = Format(Nvl(rsTmp!数量, 0), "0.00000") * Format(dbl当前单价, "0.00000")
                    dbl当前单价 = Format(dbl当前单价 * Nvl(rsTmp!住院包装, 1), "0.00000")
                ElseIf rsTmp!类别 = "4" And Nvl(rsTmp!跟踪在用, 0) = 1 And Nvl(rsTmp!是否变价, 0) = 1 Then
                    '时价卫材单价和药品一样计算
                    dbl当前单价 = CalcDrugPrice(rsTmp!ID, Val(.Cell(flexcpData, .Rows - 1, COLP_执行科室)), Nvl(rsTmp!数量, 0), , True)
                    cur当前应收 = Format(Nvl(rsTmp!数量, 0), "0.00000") * dbl当前单价
                Else
                    dbl当前单价 = Format(Nvl(rsTmp!单价, 0), "0.00000")
                    cur当前应收 = Format(Nvl(rsTmp!数量, 0), "0.00000") * dbl当前单价
                End If
                
                If rsTmp!附加手术 = 1 Then
                    cur当前应收 = cur当前应收 * Nvl(rsTmp!附术收费率, 100) / 100
                End If
                
                '处理加班加价
                If gbln加班加价 And Nvl(rsTmp!加班加价, 0) = 1 Then
                    cur当前应收 = cur当前应收 * (1 + Nvl(rsTmp!加班加价率, 0) / 100)
                End If
                
                cur当前应收 = Format(cur当前应收, gstrDec)
                
                '实收
                If gbln从项汇总折扣 And (rsTmp!从项 = 1 Or InStr(strHaveSub & ",", "," & rsTmp!行号 & ",") > 0) Then
                    cur当前实收 = Format(cur当前应收, gstrDec)
                    '累计医嘱合计来计算折扣
                    rsMain.Filter = "医嘱行号=" & rsTmp!行号
                    rsMain!医嘱合计 = Nvl(rsMain!医嘱合计, 0) + cur当前实收
                    rsMain.Update
                ElseIf Nvl(rsTmp!屏蔽费别, 0) = 0 Then
                    cur当前实收 = Format(ActualMoney(str费别, rsTmp!收入项目ID, cur当前应收, rsTmp!ID, lng执行科室ID, Nvl(rsTmp!数量, 0), _
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
        
        '汇总计算折扣
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
    End With
    ShowSendPrice = True
    Exit Function
errH:
    vsPrice.Redraw = flexRDDirect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Calc医嘱记帐金额(ByVal lngRow As Long) As Currency
'功能：计算指定医嘱行的记帐金额(显示供查看及记帐报警),以最新价格计算
'返回：str类别=计价类别
    Dim str费别 As String, dbl数量 As Double
    Dim dbl单价 As Double, cur金额 As Currency
    Dim bln附加手术 As Boolean
    
    With vsAdvice
        str费别 = .TextMatrix(lngRow, COL_费别)
        '不计价,手工计价；叮嘱,院外执行；的医嘱不读取
        If Val(.TextMatrix(lngRow, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质ID))) = 0 Then
            bln附加手术 = .TextMatrix(lngRow, COL_诊疗类别) = "F" And .RowHidden(lngRow)
            If str费别 = "" Then
                dbl数量 = Format(Val(.TextMatrix(lngRow, COL_总量)), "0.00000")
                dbl单价 = Format(CalcAdvicePrice(Val(.TextMatrix(lngRow, COL_ID)), , bln附加手术), "0.00000")
                cur金额 = Format(dbl数量 * dbl单价, gstrDec)
            Else
                dbl数量 = Format(Val(.TextMatrix(lngRow, COL_总量)), "0.00000")
                cur金额 = Format(CalcAdvicePrice(Val(.TextMatrix(lngRow, COL_ID)), str费别, bln附加手术, dbl数量), gstrDec)
            End If
        End If
    End With
    Calc医嘱记帐金额 = cur金额
End Function

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsAdvice
        If Col = COL_执行科室 Or Col = COL_附加执行 Then
            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Call vsAdvice_AfterRowColChange(Row, Col, Row, Col) '重新显示按钮
        End If
    End With
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsAdvice
        If OldRow <> NewRow And .Redraw <> flexRDNone And Not .RowHidden(NewRow) Then
            If Val(.TextMatrix(NewRow, COL_ID)) <> 0 Then
                Call ShowSendPrice(NewRow)
            End If
        End If
                
        '根据可否编辑设置编辑特性及光标特性
        If Not CellEditable(NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .ComboList = "..."
            Set .CellButtonPicture = Me.Picture
            .FocusRect = flexFocusHeavy
        End If
    End With
End Sub

Private Function CellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'功能：判断发送医嘱清单中单元格是否可以编辑
    Dim bln采集 As Boolean, blnDo As Boolean, i As Long
    
    If lngRow = 0 Then Exit Function
    
    With vsAdvice
        CellEditable = .Editable
        If lngCol = COL_执行科室 Then
            '检验组合中只有有一个可以设置，就允许选择
            If Val(.TextMatrix(lngRow, COL_相关ID)) = 0 And .TextMatrix(lngRow, COL_诊疗类别) = "E" _
                And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then bln采集 = True
            
            If bln采集 Then
                blnDo = False
                For i = lngRow - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                        If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                            blnDo = True: Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next
            Else
                blnDo = InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质ID))) = 0
            End If
            If Not blnDo Then CellEditable = False
        ElseIf lngCol = COL_附加执行 Then
            CellEditable = Should附加执行(lngRow)
        Else
            CellEditable = False
        End If
    End With
End Function

Private Function Should附加执行(ByVal lngRow As Long) As Boolean
'功能：判断指定的医嘱行(可见行)是否可以设置附加的执行科室
    Dim lngRow2 As Long, i As Long
        
    If lngRow = 0 Then Exit Function
    
    lngRow2 = -1
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_ID)) = 0 Then Exit Function
        If .TextMatrix(lngRow, COL_诊疗类别) = "F" Then
            '手术麻醉
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If .TextMatrix(i, COL_诊疗类别) = "G" Then
                        lngRow2 = i: Exit For
                    End If
                Else
                    Exit For
                End If
            Next
        ElseIf .TextMatrix(lngRow, COL_诊疗类别) = "E" _
            And .TextMatrix(lngRow - 1, COL_诊疗类别) = "C" _
            And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
            '采集方式
            lngRow2 = lngRow
        End If
        
        '检查叮嘱或院外执行
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
        If Col = COL_医嘱内容 Then
            .AutoSize COL_医嘱内容
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

Private Sub vsAdvice_DblClick()
    With vsAdvice
        If .MouseCol = COL_选择 And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsAdvice_KeyPress(32)
        End If
    End With
End Sub

Private Sub vsAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode > 127 Then '解决直接输入汉字的问题
        Call vsAdvice_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim i As Long
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
            If CellEditable(.Row, .Col) And .ComboList = "..." Then
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
            If (Col = COL_执行科室 Or Col = COL_附加执行) And .EditText <> "" Then
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

Private Sub SetDeptInput(ByVal lngRow As Long, ByVal lngCol As Long, rsInput As ADODB.Recordset)
'功能：设置执行科室输入的的值
    Dim i As Long
        
    With vsAdvice
        If lngCol = COL_执行科室 Then
            '更改显示行的执行科室显示
            .TextMatrix(lngRow, COL_执行科室) = rsInput!名称
            .Cell(flexcpData, lngRow, COL_执行科室) = .TextMatrix(lngRow, COL_执行科室)
            
            '更改主项目的执行科室(排开当前显示行为采集方式的行)
            If Not (.TextMatrix(lngRow, COL_诊疗类别) = "E" And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 _
                And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID))) Then
                .TextMatrix(lngRow, COL_执行科室ID) = rsInput!ID
                .Cell(flexcpData, lngRow, COL_执行科室ID) = 1
            End If
            
            '手术或检查组合的附加内容
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If .TextMatrix(i, COL_诊疗类别) <> "G" _
                        And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then  '不更改手术麻醉的执行科室
                        .TextMatrix(i, COL_执行科室ID) = rsInput!ID
                        .Cell(flexcpData, i, COL_执行科室ID) = 1
                    End If
                Else
                    Exit For
                End If
            Next
            
            '检验组合的内容
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                    If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                        .TextMatrix(i, COL_执行科室) = rsInput!名称
                        .Cell(flexcpData, i, COL_执行科室) = .TextMatrix(i, COL_执行科室)
                        .TextMatrix(i, COL_执行科室ID) = rsInput!ID
                        .Cell(flexcpData, i, COL_执行科室ID) = 1
                    End If
                Else
                    Exit For
                End If
            Next
        ElseIf lngCol = COL_附加执行 Then
            '更改显示行的附加执行科室显示
            .TextMatrix(lngRow, COL_附加执行) = rsInput!名称
            .Cell(flexcpData, lngRow, COL_附加执行) = .TextMatrix(lngRow, COL_附加执行)
            
            '更改附加项目行的执行科室
            If .TextMatrix(lngRow, COL_诊疗类别) = "F" Then
                '手术麻醉
                For i = lngRow + 1 To .Rows - 1
                    If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                        If .TextMatrix(i, COL_诊疗类别) = "G" Then
                            If InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                                .TextMatrix(i, COL_执行科室ID) = rsInput!ID
                                .Cell(flexcpData, i, COL_执行科室ID) = 1
                            End If
                            Exit For '只有一个麻醉
                        End If
                    Else
                        Exit For
                    End If
                Next
            ElseIf .TextMatrix(lngRow, COL_诊疗类别) = "E" And Val(.TextMatrix(lngRow, COL_相关ID)) = 0 _
                And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_ID)) Then
                '采集方式
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_执行性质ID))) = 0 Then
                    .TextMatrix(lngRow, COL_执行科室ID) = rsInput!ID
                    .Cell(flexcpData, lngRow, COL_执行科室ID) = 1
                End If
            End If
        End If
    End With
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
        Call .ShowCell(.Row, .Col)
    End With
End Sub

Private Sub vsAdvice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsAdvice.EditSelStart = 0
    vsAdvice.EditSelLength = zlCommFun.ActualLen(vsAdvice.EditText)
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsAdvice
        If Not CellEditable(Row, Col) Then Cancel = True
    End With
End Sub

Private Sub vsPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <> OldRow Then
        With vsPrice
            stbThis.Panels(2).Text = ""
            If .Cell(flexcpData, NewRow, COLP_类别) <> "" Then
                If InStr(",5,6,7,", .Cell(flexcpData, NewRow, COLP_类别)) > 0 _
                    Or .Cell(flexcpData, NewRow, COLP_类别) = "4" And Val(.Cell(flexcpData, NewRow, COLP_费用类型)) = 1 Then
                    '显示药品及跟踪卫材的库存:药品按住院单位,卫材按售价单位
                    stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_收费项目) & "，" & .TextMatrix(NewRow, COLP_执行科室) & "可用库存：" & _
                        FormatEx(GetStock(Val(.Cell(flexcpData, NewRow, COLP_收费项目)), Val(.Cell(flexcpData, NewRow, COLP_执行科室))), 5) & .TextMatrix(NewRow, COLP_单位)
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

Private Sub InitAdviceTable()
'功能：初始化清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = ",300,4;科室,850,1;姓名,750,1;住院号,750,1;床号,500,4;费别,750,1;" & _
        "婴儿,550,1;医嘱内容,2000,1;总量,600,7;单位,450,1;单量,600,7;单位,450,1;金额,850,7;" & _
        "频率,1000,1;医生嘱托,1500,1;执行科室,1500,1;附加执行,1500,1;执行时间,1000,1;首次时间,1080,1;末次时间,1080,1;" & _
        "病人ID;主页ID;性别;年龄;ID;相关ID;病人科室ID;开嘱科室ID;开嘱医生;诊疗类别;诊疗项目ID;" & _
        "计价特性;操作类型;执行性质ID;执行科室ID;次数;分解时间"
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
    
    strHead = "计价医嘱,2000,1;类别,650,1;收费项目,2000,1;数量,900,7;单位,500,1;单价,1000,7;" & _
        "应收金额,1200,7;实收金额,1200,7;执行科室,1000,1;费用类型,850,1;从项,450,4"
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

Private Function LoadAdviceSend(ByVal strEnd As String, ByVal int期效 As Integer, _
    ByVal lng执行科室ID As Long, ByVal str病人IDs As String, ByVal str类别s As String) As Boolean
'功能：根据条件读取并显示要发送的医嘱清单
'参数：strEnd=发送到的结束时间(yyyy-MM-dd HH:mm:ss),临嘱没有
'      int期效=0-长嘱,1-临嘱
'      lng执行科室ID=要发送医嘱的执行科室ID,0表示不限制
'      str病人IDs=要发送医嘱病人ID串(12,23,34....)
'      str类别s=要发送的诊疗类别串"'5','6','7'..."
'说明：注意CellData中存放得有附加数据
'   RowData：0-未发送的,-1-已成功发送的
'   COL_选择：0-可自由选择的,1-禁止改变选择状态的
'   COL_婴儿：存放婴儿编号
'   COL_诊疗类别：存放诊疗类别名称，用于显示计价医嘱
'   COL_医嘱内容：存放诊疗项目名称或标本部位，用于显示计价医嘱
'   COL_首次时间,COL_末次时间：存放持续性长嘱的首末次执行时间
'   COL_分解时间：存放费用的发生时间(无分解时间时)
'   COL_频率：1-"一次性"临嘱，2-"持续性"长嘱
'   COL_执行科室：存放原执行科室名称
'   COL_执行科室ID：是否更改了执行科室

    Dim rsSend As New ADODB.Recordset
    Dim strSQL As String, str期效条件 As String
    Dim str执行科室 As String, str诊疗类别 As String
    Dim strTmp As String, i As Long, j As Long, k As Long
    Dim datBegin As Date, datEnd As Date, strPause As String
    Dim lng次数 As Long, dbl总量 As Double, bln采集方法 As Boolean
    Dim str分解时间 As String, str首次时间 As String, str末次时间 As String
    Dim lng病人数 As Long, str科室 As String, lng单量数 As Long
    Dim lng病人ID As Long, lngDel医嘱ID As Long, lngRow As Long
        
    Screen.MousePointer = 11
    
    stbThis.Panels(3).Text = "": Call Form_Resize
    If int期效 = 0 Then
        lblInfo.Caption = "本次发送：长期医嘱，结束时间：" & strEnd
    Else
        lblInfo.Caption = "本次发送：临时医嘱"
    End If
    
    vsPrice.Rows = vsPrice.FixedRows
    vsPrice.Rows = vsPrice.FixedRows + 1
    vsAdvice.Rows = vsAdvice.FixedRows '有删除行功能
    
    vsAdvice.ColHidden(COL_科室) = True
    vsAdvice.ColHidden(COL_婴儿) = True
    vsAdvice.ColHidden(COL_首次时间) = int期效 = 1
    vsAdvice.ColHidden(COL_末次时间) = int期效 = 1
    Me.Refresh
    
    '获取发送清单:每条医嘱记录(作废的医嘱不管作废时间,作废后即无效)
    '----------------------------------------------------------------------------------------------------------
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
        '执行科室(以主要医嘱为准)
        If lng执行科室ID <> 0 Then
            '一般项目以及手术组合,检查组合;检验项目(组合)
            str执行科室 = _
                " And Exists(" & _
                " Select ID From 病人医嘱记录 X" & _
                " Where 相关ID is Null And (X.ID=A.ID Or X.ID=A.相关ID)" & _
                " And 病人ID=[2] And 执行科室ID+0=[3]" & _
                " Union ALL " & _
                " Select ID From 病人医嘱记录 X" & _
                " Where 相关ID is Not Null And 诊疗类别='C' And (X.相关ID=A.相关ID Or X.相关ID=A.ID)" & _
                " And 病人ID=[2] And 执行科室ID+0=[3])"
        End If
        
        '允许的诊疗类别部份(以主要医嘱为准)
        If str类别s <> "" Then
            '一般项目以及手术组合,检查组合;检验项目(组合)
            str诊疗类别 = _
                " And Exists(" & _
                " Select ID From 病人医嘱记录 X" & _
                " Where 相关ID is Null And (X.ID=A.ID Or X.ID=A.相关ID)" & _
                " And 病人ID=[2] And 诊疗类别 IN(" & str类别s & ")" & _
                " Union ALL" & _
                " Select ID From 病人医嘱记录 X" & _
                " Where 相关ID is Not Null And 诊疗类别='C' And (X.相关ID=A.相关ID Or X.相关ID=A.ID)" & _
                " And 病人ID=[2] And 诊疗类别 IN(" & str类别s & "))"
        End If
        
        '排开给药途径，中药煎法、用法
        strSQL = _
            " And Not(A.诊疗类别='E' And A.相关ID is Not NULL)" & _
            " And Not Exists(Select ID From 病人医嘱记录 X" & _
            " Where 诊疗类别 IN('5','6','7') And X.相关ID=A.ID" & _
            " And 病人ID=[2])"
        
        '读取发送明细:叮嘱不发送(手术,检查,检验不允许为叮嘱,采集方法可能为叮嘱),护理等级,术后医嘱不发送
        strSQL = "Select A.ID,A.相关ID,Nvl(A.相关ID,A.ID) as 组ID,Nvl(X.序号,A.序号) as 组号," & _
            " A.诊疗类别,G.名称 as 类别名称,A.诊疗项目ID,E.名称 as 诊疗项目,A.收费细目ID," & _
            " A.婴儿,A.病人ID,A.主页ID,C.住院号,B.出院病床 as 床号,D.名称 as 科室,C.姓名,C.性别,C.年龄,B.费别,B.险类," & _
            " A.开始执行时间,A.上次执行时间,A.医嘱内容,A.总给予量,A.单次用量,E.计算单位,A.执行终止时间," & _
            " A.执行频次,A.频率次数,A.频率间隔,A.间隔单位,A.医生嘱托,A.执行时间方案,A.病人科室ID,A.开嘱科室ID,A.开嘱医生," & _
            " A.标本部位,A.计价特性,E.操作类型,A.执行性质,A.执行科室ID,F.名称 as 执行科室" & _
            " From 病人医嘱记录 A,病案主页 B,病人信息 C,部门表 D,诊疗项目目录 E,部门表 F,诊疗项目类别 G,病人医嘱记录 X" & _
            " Where A.病人ID=[2] And A.病人ID=C.病人ID And B.出院科室ID=D.ID" & _
            " And A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.出院日期 is NULL And A.相关ID=X.ID(+)" & _
            " And A.诊疗项目ID=E.ID And E.类别=G.编码 And A.执行科室ID=F.ID(+)" & strSQL & _
            " And A.诊疗类别 Not IN('5','6','7')" & str期效条件 & str执行科室 & str诊疗类别 & _
            " And (Nvl(A.执行性质,0)<>0 Or A.诊疗类别='E' And E.操作类型='6')" & _
            " And Not(A.诊疗类别='H' And E.操作类型='1') And Not(A.诊疗类别='Z' And E.操作类型='4')" & _
            " And A.开始执行时间 is Not NULL And A.病人来源<>3" & _
            " Order by D.编码,LPAD(B.出院病床,10,' '),A.婴儿,组号,组ID,A.序号"
        
        On Error GoTo errH
        Set rsSend = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(IIF(strEnd = "", "1990-01-01", strEnd)), Val(Split(str病人IDs, ",")(k)), lng执行科室ID)
        
        '计算并显示发送清单
        '----------------------------------------------------------------------------------------------------------
        If Not rsSend.EOF Then
            With vsAdvice
                .Redraw = flexRDNone
                For i = 1 To rsSend.RecordCount
                    If Nvl(rsSend!相关ID, 0) = lngDel医嘱ID And lngDel医嘱ID <> 0 Then
                        GoTo NextLoop '检查组合或手术组合中的一个可能已经不能发送,则整组不能发送
                    Else
                        lngDel医嘱ID = 0
                    End If
                    
                    bln采集方法 = False
                    
                    '加入当前行
                    .Rows = .Rows + 1: lngRow = .Rows - 1
                    .Cell(flexcpPictureAlignment, lngRow, COL_选择) = 4
                    Set .Cell(flexcpPicture, lngRow, COL_选择) = img16.ListImages("T").Picture
                    
                    '隐藏:附加手术,手术麻醉,检查部位,采集方法
                    .RowHidden(lngRow) = Not IsNull(rsSend!相关ID)
                    
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
                     
                    '用于显示计价医嘱
                    .Cell(flexcpData, lngRow, COL_诊疗类别) = CStr(Nvl(rsSend!类别名称))
                    If Not IsNull(rsSend!相关ID) And rsSend!诊疗类别 = "D" Then
                        .Cell(flexcpData, lngRow, COL_医嘱内容) = CStr(Nvl(rsSend!标本部位)) '记录检查部位名
                    Else
                        .Cell(flexcpData, lngRow, COL_医嘱内容) = CStr(Nvl(rsSend!诊疗项目)) '记录诊疗项目名
                    End If
                    
                    .TextMatrix(lngRow, COL_医生嘱托) = Nvl(rsSend!医生嘱托)
                    .TextMatrix(lngRow, COL_执行时间) = Nvl(rsSend!执行时间方案)
                    .TextMatrix(lngRow, COL_频率) = Nvl(rsSend!执行频次)
                    
                    .TextMatrix(lngRow, COL_病人科室ID) = Nvl(rsSend!病人科室ID)
                    .TextMatrix(lngRow, COL_开嘱科室ID) = Nvl(rsSend!开嘱科室ID)
                    .TextMatrix(lngRow, COL_开嘱医生) = Nvl(rsSend!开嘱医生)
                    
                    .TextMatrix(lngRow, COL_计价特性) = Nvl(rsSend!计价特性, 0)
                    .TextMatrix(lngRow, COL_操作类型) = Nvl(rsSend!操作类型)
                    .TextMatrix(lngRow, COL_执行性质ID) = Nvl(rsSend!执行性质, 0)
                    
                    '主项目执行科室显示
                    If IsNull(rsSend!相关ID) And rsSend!诊疗类别 = "E" _
                        And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = rsSend!ID Then
                        '采集方法显示为检验项目的执行科室
                        bln采集方法 = True
                        .TextMatrix(lngRow, COL_执行科室) = .TextMatrix(lngRow - 1, COL_执行科室)
                        .Cell(flexcpData, lngRow, COL_执行科室) = .Cell(flexcpData, lngRow - 1, COL_执行科室)
                    Else
                        .TextMatrix(lngRow, COL_执行科室) = Nvl(rsSend!执行科室)
                        .Cell(flexcpData, lngRow, COL_执行科室) = CStr(Nvl(rsSend!执行科室))
                    End If
                    
                    '附加项目执行科室显示
                    If Nvl(rsSend!诊疗类别) = "E" And IsNull(rsSend!相关ID) _
                        And .TextMatrix(lngRow - 1, COL_诊疗类别) = "C" _
                        And Val(.TextMatrix(lngRow - 1, COL_相关ID)) = rsSend!ID Then
                        '采集方式在当前行显示附加执行科室
                        .TextMatrix(lngRow, COL_附加执行) = Nvl(rsSend!执行科室)
                        .Cell(flexcpData, lngRow, COL_附加执行) = CStr(Nvl(rsSend!执行科室))
                    ElseIf Nvl(rsSend!诊疗类别) = "G" And Not IsNull(rsSend!相关ID) Then
                        '手术麻醉主手术行显示附加执行科室
                        j = .FindRow(CStr(rsSend!相关ID), .FixedRows, COL_ID)
                        If j <> -1 Then
                            .TextMatrix(j, COL_附加执行) = Nvl(rsSend!执行科室)
                            .Cell(flexcpData, j, COL_附加执行) = CStr(Nvl(rsSend!执行科室))
                        End If
                    End If
                    
                    .TextMatrix(lngRow, COL_执行科室ID) = Nvl(rsSend!执行科室ID)
                                    
                    '计算发送次数，执行的分解时间，总量
                    '---------------------------------------------------------------
                    If int期效 = 0 Then
                        '长嘱---------------------------------------------
                        If (IsNull(rsSend!相关ID) And Not bln采集方法) _
                            Or (Not IsNull(rsSend!相关ID) And rsSend!诊疗类别 = "C") Then '主要医嘱或一并采集的检验项目
                        
                            '当前医嘱的暂停时间段:"暂停时间,开始时间;...."
                            strPause = GetAdvicePause(rsSend!ID)
                            
                            '当前医嘱的发送计算时间段
                            datBegin = rsSend!开始执行时间
                            If Not IsNull(rsSend!上次执行时间) Then
                                If IsNull(rsSend!执行时间方案) And (Nvl(rsSend!频率次数, 0) = 0 Or Nvl(rsSend!频率间隔, 0) = 0 Or IsNull(rsSend!间隔单位)) Then
                                    datBegin = DateAdd("s", 1, rsSend!上次执行时间) '"持续性"的项目
                                Else
                                    datBegin = Calc本周期开始时间(rsSend!开始执行时间, rsSend!上次执行时间, rsSend!频率间隔, rsSend!间隔单位)
                                    
                                    '本周期内已执行的时间不再计算,这里通过暂停方式来处理
                                    strPause = strPause & ";" & Format(datBegin, "yyyy-MM-dd HH:mm:ss") & "," & Format(rsSend!上次执行时间, "yyyy-MM-dd HH:mm:ss")
                                    If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
                                End If
                            End If
                            datEnd = CDate(strEnd)
                            If Not IsNull(rsSend!执行终止时间) Then
                                If rsSend!执行终止时间 < CDate(strEnd) Then
                                    datEnd = rsSend!执行终止时间
                                End If
                            End If
                            
                            '计算分解时间及次数
                            If IsNull(rsSend!执行时间方案) And (Nvl(rsSend!频率次数, 0) = 0 Or Nvl(rsSend!频率间隔, 0) = 0 Or IsNull(rsSend!间隔单位)) Then
                                '执行频率为"持续性"的项目,每天发送一次(00:00)
                                lng次数 = Calc持续性长嘱次数(datBegin, datEnd, _
                                    Format(Nvl(rsSend!上次执行时间), "yyyy-MM-dd HH:mm:ss"), _
                                    Format(Nvl(rsSend!执行终止时间), "yyyy-MM-dd HH:mm:ss"), _
                                    strPause, str首次时间, str末次时间)
                                If lng次数 = 0 Then '本次无需发送
                                    lngDel医嘱ID = Nvl(rsSend!ID, 0)
                                    .RemoveItem lngRow
                                    GoTo NextLoop
                                End If
                                
                                '记录本条医嘱发送的首次,末次时间(仅持续性长嘱)
                                str分解时间 = "" '不需要
                                .Cell(flexcpData, lngRow, COL_首次时间) = str首次时间
                                .Cell(flexcpData, lngRow, COL_末次时间) = str末次时间
                                
                                '记录费用发生时间(当无分解时间时),以本次发送首次时间
                                .Cell(flexcpData, lngRow, COL_分解时间) = str首次时间
                                
                                '标记为"持续性"长嘱
                                .Cell(flexcpData, lngRow, COL_频率) = 2
                            Else
                                '执行频率为"可选频率"的项目
                                str分解时间 = Calc段内分解时间(datBegin, datEnd, strPause, rsSend!执行时间方案, rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位)
                                If str分解时间 = "" Then '无法分解时间(如被暂停的)
                                    lngDel医嘱ID = Nvl(rsSend!ID, 0)
                                    .RemoveItem lngRow
                                    GoTo NextLoop
                                End If
                                lng次数 = UBound(Split(str分解时间, ",")) + 1
                            End If
                            dbl总量 = Nvl(rsSend!单次用量, 1) * lng次数
    
                            .TextMatrix(lngRow, COL_次数) = lng次数
                            .TextMatrix(lngRow, COL_分解时间) = str分解时间
                            If str分解时间 <> "" Then
                                .TextMatrix(lngRow, COL_首次时间) = Format(Split(str分解时间, ",")(0), "MM-dd HH:mm")
                                .TextMatrix(lngRow, COL_末次时间) = Format(Split(str分解时间, ",")(lng次数 - 1), "MM-dd HH:mm")
                            End If
                            
                            .TextMatrix(lngRow, COL_单量) = FormatEx(Nvl(rsSend!单次用量), 5)
                            If Not IsNull(rsSend!单次用量) Then
                                .TextMatrix(lngRow, COL_单量单位) = Nvl(rsSend!计算单位)
                            End If
                            .TextMatrix(lngRow, COL_总量) = FormatEx(dbl总量, 5)
                            .TextMatrix(lngRow, COL_总量单位) = Nvl(rsSend!计算单位)
                        ElseIf Not IsNull(rsSend!相关ID) Or bln采集方法 Then '附加医嘱或标本采集方法
                            '检查组合和手术组合不可能为长嘱,所以此段不会执行
                            .TextMatrix(lngRow, COL_单量) = FormatEx(Nvl(rsSend!单次用量), 5)
                            If Not IsNull(rsSend!单次用量) Then
                                .TextMatrix(lngRow, COL_单量单位) = Nvl(rsSend!计算单位)
                            End If
                            .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngRow - 1, COL_总量)
                            .TextMatrix(lngRow, COL_总量单位) = Nvl(rsSend!计算单位)
                            .TextMatrix(lngRow, COL_次数) = .TextMatrix(lngRow - 1, COL_次数)
                            .TextMatrix(lngRow, COL_分解时间) = .TextMatrix(lngRow - 1, COL_分解时间)
                            .Cell(flexcpData, lngRow, COL_分解时间) = .Cell(flexcpData, lngRow - 1, COL_分解时间)
                            .TextMatrix(lngRow, COL_首次时间) = .TextMatrix(lngRow - 1, COL_首次时间)
                            .TextMatrix(lngRow, COL_末次时间) = .TextMatrix(lngRow - 1, COL_末次时间)
                        End If
                    Else
                        '临嘱---------------------------------------------
                        If (IsNull(rsSend!相关ID) And Not bln采集方法) _
                            Or (Not IsNull(rsSend!相关ID) And rsSend!诊疗类别 = "C") Then '主要医嘱或一并采集的检验项目
                            
                            dbl总量 = Nvl(rsSend!总给予量, 1)
                            lng次数 = IntEx(dbl总量 / Nvl(rsSend!单次用量, 1))
                            
                            If IsNull(rsSend!执行时间方案) And (Nvl(rsSend!频率次数, 0) = 0 Or Nvl(rsSend!频率间隔, 0) = 0 Or IsNull(rsSend!间隔单位)) Then
                                '执行频率为"一次性"的项目
                                str分解时间 = "" '不需要
                                .Cell(flexcpData, lngRow, COL_频率) = 1
                            Else
                                '执行频率为"可选频率"的项目
                                If Not IsNull(rsSend!执行时间方案) Then
                                    str分解时间 = Calc次数分解时间(lng次数, rsSend!开始执行时间, CDate("3000-01-01"), "", rsSend!执行时间方案, rsSend!频率次数, rsSend!频率间隔, rsSend!间隔单位)
                                Else
                                    str分解时间 = "" '临嘱也许未输入执行时间,无法分解
                                End If
                            End If
                            .TextMatrix(lngRow, COL_次数) = lng次数
                            .TextMatrix(lngRow, COL_分解时间) = str分解时间
                            If str分解时间 <> "" Then
                                .TextMatrix(lngRow, COL_首次时间) = Format(Split(str分解时间, ",")(0), "MM-dd HH:mm")
                                .TextMatrix(lngRow, COL_末次时间) = Format(Split(str分解时间, ",")(lng次数 - 1), "MM-dd HH:mm")
                            Else
                                '记录费用发生时间(当无分解时间时),以医嘱的开始执行时间
                                .Cell(flexcpData, lngRow, COL_分解时间) = CStr(Format(rsSend!开始执行时间, "yyyy-MM-dd HH:mm:ss"))
                            End If
                            
                            .TextMatrix(lngRow, COL_单量) = FormatEx(Nvl(rsSend!单次用量), 5)
                            If Not IsNull(rsSend!单次用量) Then
                                .TextMatrix(lngRow, COL_单量单位) = Nvl(rsSend!计算单位)
                            End If
                            .TextMatrix(lngRow, COL_总量) = FormatEx(dbl总量, 5)
                            .TextMatrix(lngRow, COL_总量单位) = Nvl(rsSend!计算单位)
                        ElseIf Not IsNull(rsSend!相关ID) Or bln采集方法 Then '附加医嘱或标本采集方法
                            .TextMatrix(lngRow, COL_单量) = FormatEx(Nvl(rsSend!单次用量), 5)
                            If Not IsNull(rsSend!单次用量) Then
                                .TextMatrix(lngRow, COL_单量单位) = Nvl(rsSend!计算单位)
                            End If
                            .TextMatrix(lngRow, COL_总量) = .TextMatrix(lngRow - 1, COL_总量)
                            .TextMatrix(lngRow, COL_总量单位) = Nvl(rsSend!计算单位)
                            .TextMatrix(lngRow, COL_次数) = .TextMatrix(lngRow - 1, COL_次数)
                            .TextMatrix(lngRow, COL_分解时间) = .TextMatrix(lngRow - 1, COL_分解时间)
                            .Cell(flexcpData, lngRow, COL_分解时间) = .Cell(flexcpData, lngRow - 1, COL_分解时间)
                            .TextMatrix(lngRow, COL_首次时间) = .TextMatrix(lngRow - 1, COL_首次时间)
                            .TextMatrix(lngRow, COL_末次时间) = .TextMatrix(lngRow - 1, COL_末次时间)
                        End If
                    End If
                    If Not IsNull(rsSend!单次用量) Then
                        lng单量数 = lng单量数 + 1 '决定是否显示单量列
                    End If
                    
                    '计算项目的金额:用于查看及记帐报警
                    '---------------------------------------------------------------
                    .TextMatrix(lngRow, COL_金额) = Format(Calc医嘱记帐金额(lngRow), gstrDec)
                    
                    '相关行时的一些处理：累计显示一组医嘱的金额
                    '---------------------------------------------------------------
                    If Not IsNull(rsSend!相关ID) And rsSend!诊疗类别 <> "C" Then
                        '其它附加医嘱
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_ID)) = rsSend!相关ID Then
                                .TextMatrix(j, COL_金额) = Format(Val(.TextMatrix(j, COL_金额)) + Val(.TextMatrix(lngRow, COL_金额)), gstrDec)
                                Exit For
                            End If
                        Next
                    ElseIf bln采集方法 Then
                        '检验标本采集方法为显示行
                        For j = lngRow - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(j, COL_相关ID)) = rsSend!ID Then
                                .TextMatrix(lngRow, COL_金额) = Format(Val(.TextMatrix(lngRow, COL_金额)) + Val(.TextMatrix(j, COL_金额)), gstrDec)
                            Else
                                Exit For
                            End If
                        Next
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
        .ColHidden(COL_单量) = lng单量数 = 0
        .ColHidden(COL_单量单位) = .ColHidden(COL_单量)
        
        .AutoSize COL_医嘱内容
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
    vsAdvice.SetFocus: Call vsAdvice_GotFocus
    Call ShowSendTotal
    Progress = 0: Screen.MousePointer = 0
    
    LoadAdviceSend = True
    Exit Function
errH:
    vsAdvice.Redraw = flexRDDirect
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        vsAdvice.Redraw = flexRDNone: Resume
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
'5.检查部位和附加手术与主要医嘱分配相同单据号，手术麻醉分配单独的单据号。
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
'返回：如果成功则返回发送号
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
    Dim dbl数量 As Double, dbl单价 As Double, cur应收 As Currency, cur实收 As Currency, cur合计 As Currency
    Dim bln保险项目否 As Boolean, lng保险大类ID As Long, cur统筹金额 As Currency, str保险编码 As String, str费用类型 As String
    Dim str发送数次 As String, str分解时间 As String, str首次时间 As String, str末次时间 As String
    Dim bln附加手术 As Boolean, strNOKey As String, blnFirst As Boolean '配方数及分号关键字
    Dim lng病人科室ID As Long, lng执行科室ID As Long, str自动发料 As String
    Dim str发生时间 As String, str登记时间 As String
    Dim int执行状态 As Integer, blnBool As Boolean
    
    Dim bln药品时价提示 As Boolean, bln药品库存提示 As Boolean, bln药品默认发送 As Boolean
    Dim bln卫材时价提示 As Boolean, bln卫材库存提示 As Boolean, bln卫材默认发送 As Boolean
     
    Dim blnHaveSub As Boolean, cur医嘱合计 As Currency
    Dim int父序号 As Integer, var父索引 As Variant
    Dim lng父收入ID As Long, str实收 As String
            
    Dim rsAudit As ADODB.Recordset
    Dim strAudit As String
    
    mstrRollNotify = ""
    
    With vsAdvice
        '先检查并提示特殊医嘱:3-转科,5-出院,6-转院,11-死亡
        strTmp = ""
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_选择) = 0 And Not .Cell(flexcpPicture, i, COL_选择) Is Nothing Then
                If .TextMatrix(i, COL_诊疗类别) = "Z" And InStr(",3,5,6,11,", Val(.TextMatrix(i, COL_操作类型))) > 0 Then
                    strTmp = strTmp & vbCrLf & .TextMatrix(i, COL_姓名) & IIF(.Cell(flexcpData, i, COL_婴儿) <> 0, "(婴儿" & .Cell(flexcpData, i, COL_婴儿) & ")", "") & "：" & .TextMatrix(i, COL_医嘱内容)
                End If
            End If
        Next
        If strTmp <> "" Then
            If MsgBox("要发送的医嘱中包含下列特殊医嘱：" & vbCrLf & strTmp & vbCrLf & vbCrLf & "确实要发送当前选择的医嘱吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            If MsgBox("确实要发送当前选择的医嘱吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End With
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    bln药品时价提示 = True: bln药品库存提示 = True: bln药品默认发送 = True
    bln卫材时价提示 = True: bln卫材库存提示 = True: bln卫材默认发送 = True
    
    intWarn = -1 '记帐报警时缺省要提示,与病人无关
    lng发送号 = zlDatabase.GetNextNO(10)
    curDate = zlDatabase.Currentdate
    Call InitBillSet
    
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
                    
                    strSQL = _
                        " Select 病人ID,预交余额,费用余额,0 as 预结费用 From 病人余额 Where 性质=1 And 病人ID=[1]" & _
                        " Union ALL" & _
                        " Select A.病人ID,0,0,Sum(金额) From 保险模拟结算 A,病案主页 B" & _
                        " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.险类 Is Not Null And A.病人ID=[1] And A.主页ID=[2] Group by A.病人ID"
                    strSQL = "Select 病人ID,Nvl(Sum(预交余额),0)-Nvl(Sum(费用余额),0)+Nvl(Sum(预结费用),0) as 剩余款 From (" & strSQL & ") Group by 病人ID"
                    
                    '获取当前病人信息,状态:0-正常；1-尚未入科；2-正在转科；3-已预出院
                    strSQL = "Select A.病人ID,B.主页ID,A.姓名,B.险类,B.状态,B.当前病区ID,B.出院科室ID," & _
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
                End If
                
                '特殊医嘱：3-转科;5-出院;6-转院,11-死亡
                If .TextMatrix(i, COL_诊疗类别) = "Z" Then
                    '转科,出院,转院,死亡医嘱发送时，病人要处于正常状态
                    If .Cell(flexcpData, i, COL_婴儿) = 0 Then
                        If InStr(",3,5,6,11,", .TextMatrix(i, COL_操作类型)) > 0 And Nvl(rsPati!状态, 0) <> 0 Then
                            MsgBox "病人""" & rsPati!姓名 & """当前处于""" & Decode(Nvl(rsPati!状态, 0), 1, "等待入科", 2, "正在转科", 3, "已预出院") & """状态，" & _
                                "不能发送""" & .TextMatrix(i, COL_医嘱内容) & """医嘱。", vbInformation, gstrSysName
                            Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                            GoTo NextLoop
                        End If
                    End If
                    
                    '如果是转科、出院、转院医嘱,检查病人是否有未执行的医技项目及未发药品
                    If InStr(",3,5,6,", .TextMatrix(i, COL_操作类型)) > 0 And gbyt检查未执行 <> 0 Then
                        strTmp = ExistWaitExe(lng病人ID, Val(.TextMatrix(i, COL_主页ID)), .Cell(flexcpData, i, COL_婴儿))
                        If strTmp <> "" Then
                            Call .ShowCell(i, COL_医嘱内容): .Refresh
                            If gbyt检查未执行 = 1 Then
                                If MsgBox("发现病人""" & rsPati!姓名 & IIF(.Cell(flexcpData, i, COL_婴儿) <> 0, "(婴儿" & .Cell(flexcpData, i, COL_婴儿) & ")", "") & """存在尚未执行完成的内容：" & _
                                    vbCrLf & vbCrLf & strTmp & vbCrLf & vbCrLf & "确实要发送""" & .TextMatrix(i, COL_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                                    GoTo NextLoop
                                End If
                            Else
                                MsgBox "发现病人""" & rsPati!姓名 & IIF(.Cell(flexcpData, i, COL_婴儿) <> 0, "(婴儿" & .Cell(flexcpData, i, COL_婴儿) & ")", "") & """存在尚未执行完成的内容：" & _
                                    vbCrLf & vbCrLf & strTmp & vbCrLf & vbCrLf & "医嘱""" & .TextMatrix(i, COL_医嘱内容) & """将不被发送。", vbInformation, gstrSysName
                                Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                                GoTo NextLoop
                            End If
                        End If
                        
                        strTmp = ExistWaitDrug(lng病人ID, Val(.TextMatrix(i, COL_主页ID)), .Cell(flexcpData, i, COL_婴儿))
                        If strTmp <> "" Then
                            Call .ShowCell(i, COL_医嘱内容): .Refresh
                            If gbyt检查未执行 = 1 Then
                                If MsgBox("发现病人""" & rsPati!姓名 & IIF(.Cell(flexcpData, i, COL_婴儿) <> 0, "(婴儿" & .Cell(flexcpData, i, COL_婴儿) & ")", "") & """" & _
                                    strTmp & vbCrLf & vbCrLf & "确实要发送""" & .TextMatrix(i, COL_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                                    GoTo NextLoop
                                End If
                            Else
                                MsgBox "发现病人""" & rsPati!姓名 & IIF(.Cell(flexcpData, i, COL_婴儿) <> 0, "(婴儿" & .Cell(flexcpData, i, COL_婴儿) & ")", "") & """" & _
                                    strTmp & vbCrLf & vbCrLf & "医嘱""" & .TextMatrix(i, COL_医嘱内容) & """将不被发送", vbInformation, gstrSysName
                                Set .Cell(flexcpPicture, i, COL_选择) = Nothing
                                GoTo NextLoop
                            End If
                        End If
                    End If
                    
                    '因为自动停止医嘱，需要进行超期收回提醒
                    If InStr(",3,5,6,11,", .TextMatrix(i, COL_操作类型)) > 0 Then
                        If InStr(mstrRollNotify, "(" & lng病人ID & "," & Val(.TextMatrix(i, COL_主页ID)) & ")") = 0 Then
                            mstrRollNotify = mstrRollNotify & ",(" & lng病人ID & "," & Val(.TextMatrix(i, COL_主页ID)) & ")"
                        End If
                    End If
                End If
                
                '产生单据号分配关键字
                '-----------------------------------------------------------------------------------------
                If .TextMatrix(i, COL_诊疗类别) = "M" Then
                    '材料按"病人(病人ID,主页ID)_病人科室ID_开嘱科室ID_开嘱医生_执行科室ID"分号。
                    strNOKey = "材料医嘱_" & lng病人ID & "_" & Val(.TextMatrix(i, COL_主页ID)) & "_" & _
                        Val(.TextMatrix(i, COL_病人科室ID)) & "_" & Val(.TextMatrix(i, COL_开嘱科室ID)) & "_" & _
                        .TextMatrix(i, COL_开嘱医生) & "_" & Val(.TextMatrix(i, COL_执行科室ID))
                    '再按要打印的诊疗单据分号
                    strNOKey = strNOKey & "_" & GetClinicBillID(Val(.TextMatrix(i, COL_诊疗项目ID)), 2)
                ElseIf Val(.TextMatrix(i, COL_相关ID)) <> 0 And .TextMatrix(i, COL_诊疗类别) = "C" Then
                    '一并采集的检验组合分配相同的单据号，标本采集方法分配单独的单据号
                    strNOKey = "一并采集_" & Val(.TextMatrix(i, COL_相关ID))
                ElseIf Val(.TextMatrix(i, COL_相关ID)) <> 0 And .TextMatrix(i, COL_诊疗类别) <> "G" Then
                    '检查部位和附加手术与主要医嘱分配相同单据号，手术麻醉分配单独的单据号。
                    strNOKey = "非药医嘱_" & Val(.TextMatrix(i, COL_相关ID))
                Else
                    '其它非药医嘱每条医嘱一个独立单据号
                    strNOKey = "非药医嘱_" & Val(.TextMatrix(i, COL_ID))
                End If
                
                '产生医嘱记帐费用:以最新价格计算
                '-----------------------------------------------------------------------------------------
                strSQL = "": lng细目ID = 0
                If Val(.TextMatrix(i, COL_计价特性)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_执行性质ID))) = 0 Then
                    '不计价,手工计价；叮嘱,院外执行的医嘱不读取；
                    strSQL = _
                        " Select A.ID,A.类别,D.名称 as 类别名称,A.名称,A.计算单位,B.收入项目ID,C.收据费目," & _
                        " Y.住院单位,Y.住院包装,X.数量,Decode(A.是否变价,1,X.单价,B.现价) as 单价,A.加班加价," & _
                        " B.加班加价率,B.附术收费率,A.是否变价,Decode(A.类别,'4',E.在用分批,Y.药房分批) as 分批," & _
                        " E.跟踪在用,Nvl(X.从项,0) as 从项,Nvl(X.执行科室ID,[2]) as 执行科室ID,A.屏蔽费别,I.要求审批" & _
                        " From 收费项目目录 A,收费价目 B,收入项目 C,收费项目类别 D,材料特性 E,病人医嘱计价 X,药品规格 Y,保险支付项目 I" & _
                        " Where A.ID=B.收费细目ID And B.收入项目ID=C.ID And A.类别=D.编码 And A.ID=E.材料ID(+)" & _
                        " And A.ID=Y.药品ID(+) And X.收费细目ID=A.ID And Nvl(X.数量,0)<>0 And X.医嘱ID=[1]" & _
                        " And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
                        " And A.ID=I.收费细目ID(+) And I.险类(+)=[3]" & _
                        " Order by 从项,A.ID"
                        '一定要把主项排在前面,以便于计算和在费用记录中保持主从关系
                End If
                
                '汇总折扣变量初始
                blnHaveSub = False
                var父索引 = Empty: int父序号 = 0
                cur医嘱合计 = 0: lng父收入ID = 0
                
                int计费状态 = IIF(Val(.TextMatrix(i, COL_计价特性)) = 1, -1, 0) '无需计费或未计费
                If CLng(.Cell(flexcpData, i, COL_频率)) = 2 Then
                    If strSQL <> "" Then strSQL = "" '"持续性"长嘱不产生费用
                End If
                If strSQL <> "" Then
                    Set rsPrice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_执行科室ID)), Val(Nvl(rsPati!险类, 0)))
                    If Not rsPrice.EOF Then
                        int计费状态 = 1 '已计费
                        '确定是否主从关系:即使不汇总折扣,也要记录
                        rsPrice.Filter = "从项=1"
                        If Not rsPrice.EOF Then blnHaveSub = True
                        rsPrice.Filter = 0
                    End If
                    '处理收入项目级的费用明细
                    bln附加手术 = .TextMatrix(i, COL_诊疗类别) = "F" And Val(.TextMatrix(i, COL_相关ID)) <> 0
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
                        If rsPrice!类别 = "4" And Nvl(rsPrice!跟踪在用, 0) = 1 Or InStr(",5,6,7", rsPrice!类别) > 0 Then
                            lng病人科室ID = Val(.TextMatrix(i, COL_病人科室ID))
                            lng执行科室ID = Get收费执行科室ID(rsPati!病人ID, rsPati!主页ID, rsPrice!类别, rsPrice!ID, 4, lng病人科室ID, 0, 2, lng执行科室ID)
                            
                            '卫材必须设置执行科室
                            If lng执行科室ID = 0 And rsPrice!类别 = "4" Then
                                .Row = GetVisibleRow(i)
                                Call .ShowCell(.Row, .Col)
                                Screen.MousePointer = 0
                                MsgBox "系统不能为计价卫材""" & rsPrice!名称 & """确定合适的执行科室。" & vbCrLf & _
                                    "请使用计价调整功能人为确定，或到""卫材目录管理""中检查存储库房设置是否正确。", vbInformation, gstrSysName
                                Call DeleteSendRow: Call ShowSendTotal
                                Progress = 0: Exit Function
                            End If
                        End If
                        
                        '数量
                        dbl数量 = Format(Val(.TextMatrix(i, COL_总量)) * Nvl(rsPrice!数量, 0), "0.00000")
                        
                        '非药医嘱对应的药品计价
                        If InStr(",5,6,7,", rsPrice!类别) > 0 Then
                            If Nvl(rsPrice!是否变价, 0) = 0 Then
                                dbl单价 = Format(Nvl(rsPrice!单价, 0), "0.00000")
                            Else
                                dbl单价 = Format(CalcDrugPrice(rsPrice!ID, lng执行科室ID, dbl数量, , True), "0.00000")
                            End If
                        ElseIf rsPrice!类别 = "4" And Nvl(rsPrice!跟踪在用, 0) = 1 Then
                            '检查卫生材料入出类别
                            If mlng卫材类别ID = 0 Then
                                Screen.MousePointer = 0
                                MsgBox "不能确定卫生材料单据的入出类别,请先到入出类别管理中设置！", vbInformation, gstrSysName
                                Call DeleteSendRow: Call ShowSendTotal
                                Progress = 0: Exit Function
                            End If
                            
                            If Nvl(rsPrice!是否变价, 0) = 0 Then
                                dbl单价 = Format(Nvl(rsPrice!单价, 0), "0.00000")
                            Else
                                dbl单价 = Format(CalcDrugPrice(rsPrice!ID, lng执行科室ID, dbl数量, , True), "0.00000")
                            End If
                        Else
                            dbl单价 = Format(Nvl(rsPrice!单价, 0), "0.00000")
                        End If
                        
                        '非药嘱药品及跟踪卫材的库存检查
                        If rsPrice!类别 = "4" And Nvl(rsPrice!跟踪在用, 0) = 1 Or InStr(",5,6,7", rsPrice!类别) > 0 Then
                            If GetStockCheck(lng执行科室ID) <> 0 Or Nvl(rsPrice!是否变价, 0) = 1 Or Nvl(rsPrice!分批, 0) = 1 Then
                                If rsPrice!类别 = "4" Then
                                    blnBool = CheckPriceStock(i, rsPrice, lng执行科室ID, dbl数量, rsTotal, bln卫材库存提示, bln卫材时价提示, bln卫材默认发送)
                                Else
                                    blnBool = CheckPriceStock(i, rsPrice, lng执行科室ID, dbl数量, rsTotal, bln药品库存提示, bln药品时价提示, bln药品默认发送)
                                End If
                                If blnBool Then
                                    Call RowSelectSame(i, COL_选择, rsSQL, rsTotal, rsUpload)
                                    GoTo NextLoop
                                End If
                            End If
                        End If
                        
                        '发送金额
                        cur应收 = dbl数量 * dbl单价
                        If bln附加手术 Then
                            cur应收 = cur应收 * Nvl(rsPrice!附术收费率, 100) / 100
                        End If
                        
                        '处理加班加价
                        If gbln加班加价 And Nvl(rsPrice!加班加价, 0) = 1 Then
                            cur应收 = cur应收 * (1 + Nvl(rsPrice!加班加价率, 0) / 100)
                        End If
                        
                        cur应收 = Format(cur应收, gstrDec)
                        
                        '计算汇总折扣合计
                        If gbln从项汇总折扣 And blnHaveSub Then
                            cur实收 = cur应收
                            cur医嘱合计 = cur医嘱合计 + cur实收
                        ElseIf Nvl(rsPrice!屏蔽费别, 0) = 0 Then
                            cur实收 = Format(ActualMoney(.TextMatrix(i, COL_费别), rsPrice!收入项目ID, cur应收, rsPrice!ID, lng执行科室ID, dbl数量, _
                                IIF(gbln加班加价 And Nvl(rsPrice!加班加价, 0) = 1, Nvl(rsPrice!加班加价率, 0) / 100, 0)), gstrDec)
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
                        
                        '因为现在不计价的医嘱不产生费用,所以传入的计价特性都为(0-正常计价)
                        '是否划价费用
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
                            "1," & dbl数量 & "," & IIF(bln附加手术, 1, 0) & "," & ZVal(lng执行科室ID) & "," & _
                            IIF(lng费用父号 = lng费用序号, "NULL", lng费用父号) & "," & rsPrice!收入项目ID & "," & _
                            "'" & Nvl(rsPrice!收据费目) & "'," & dbl单价 & "," & cur应收 & "," & str实收 & "," & _
                            cur统筹金额 & "," & str发生时间 & "," & str登记时间 & "," & _
                            "NULL," & int划价 & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "',0," & _
                            IIF(rsPrice!类别 = "4", mlng卫材类别ID, mlng药品类别ID) & "," & _
                            "NULL,'" & .TextMatrix(i, COL_医嘱内容) & "',NULL," & Val(.TextMatrix(i, COL_ID)) & "," & _
                            "Null,Null,Null,Null,Null,Null,'" & str费用类型 & "')"
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
                
                '对医嘱金额进行汇总折扣处理
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
                
                '产生医嘱发送记录:一样要产生费用NO
                '-----------------------------------------------------------------------------------------
                If Val(.TextMatrix(i, COL_执行性质ID)) <> 0 Then '叮嘱不发送(采集方法可能为)
                    '发送了出院,转院,死亡医嘱
                    If .TextMatrix(i, COL_诊疗类别) = "Z" _
                        And InStr(",5,6,11,", Val(.TextMatrix(i, COL_操作类型))) > 0 Then
                        mblnRefresh = True
                    End If
                    
                    Call GetCurBillSet(strNOKey, strNO, -1, lng发送序号)
                    
                    str分解时间 = .TextMatrix(i, COL_分解时间)
                    If str分解时间 <> "" Then
                        str发送数次 = Format(Val(.TextMatrix(i, COL_总量)), "0.00000")
                        str首次时间 = "To_Date('" & Split(str分解时间, ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                        str末次时间 = "To_Date('" & Split(str分解时间, ",")(Val(.TextMatrix(i, COL_次数)) - 1) & "','YYYY-MM-DD HH24:MI:SS')"
                    ElseIf CLng(.Cell(flexcpData, i, COL_频率)) = 2 Then
                        '"持续性"长嘱:不填写发送数次
                        str发送数次 = "NULL"
                        str首次时间 = "To_Date('" & .Cell(flexcpData, i, COL_首次时间) & "','YYYY-MM-DD HH24:MI:SS')"
                        str末次时间 = "To_Date('" & .Cell(flexcpData, i, COL_末次时间) & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        '否则应为"一次性"临嘱
                        str发送数次 = Format(Val(.TextMatrix(i, COL_总量)), "0.00000")
                        str首次时间 = "NULL"
                        str末次时间 = "NULL"
                    End If
                    
                    '执行状态
                    int执行状态 = 0
                    If mblnAutoExe And Val(.TextMatrix(i, COL_开嘱科室ID)) = Val(.TextMatrix(i, COL_执行科室ID)) Then
                        '本科执行的自动执行,特殊医嘱不处理
                        If Not (.TextMatrix(i, COL_诊疗类别) = "Z" And Val(.TextMatrix(i, COL_操作类型)) <> 0) Then
                            int执行状态 = 1
                        End If
                    End If
                    
                    '是否一组医嘱的第一行
                    blnFirst = False
                    If .TextMatrix(i, COL_诊疗类别) = "C" And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                        If Val(.TextMatrix(i, COL_相关ID)) <> Val(.TextMatrix(i - 1, COL_相关ID)) Then
                            blnFirst = True '检验组合中的第一检验行
                        End If
                    ElseIf Val(.TextMatrix(i, COL_相关ID)) = 0 Then
                        If Not (.TextMatrix(i, COL_诊疗类别) = "E" _
                            And Val(.TextMatrix(i, COL_ID)) = Val(.TextMatrix(i - 1, COL_相关ID))) Then '排开采集方法
                            blnFirst = True
                        End If
                    End If
                    
                    rsSQL.AddNew
                    rsSQL!类型 = 3
                    rsSQL!医嘱ID = IIF(Val(.TextMatrix(i, COL_相关ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_相关ID)))
                    rsSQL!项目ID = 0
                    rsSQL!序号 = i
                    rsSQL!SQL = "ZL_病人医嘱发送_Insert(" & _
                        Val(.TextMatrix(i, COL_ID)) & "," & lng发送号 & ",2,'" & strNO & "'," & _
                        lng发送序号 & "," & str发送数次 & "," & str首次时间 & "," & str末次时间 & "," & _
                        "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                        int执行状态 & "," & ZVal(.TextMatrix(i, COL_执行科室ID)) & "," & int计费状态 & "," & IIF(blnFirst, 1, 0) & ")"
                    rsSQL.Update
                End If
            End If
            
            '----------------------------------------
NextLoop:
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
        If ErrCenter() = 1 Then Resume
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
            
            '当前可用库存:住院包装,减去前面计价部份要发送的累计数量
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
                    .Row = GetVisibleRow(lngRow)
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
