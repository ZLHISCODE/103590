VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmMedicalStationPrintRpt 
   Caption         =   "#"
   ClientHeight    =   6015
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   9780
   Icon            =   "frmMedicalStationPrintRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   9780
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   3570
      Left            =   390
      TabIndex        =   0
      Top             =   765
      Width           =   7170
      _cx             =   12647
      _cy             =   6297
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
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   435
         Y2              =   1650
      End
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   630
      Left            =   360
      TabIndex        =   4
      Top             =   4920
      Width           =   3645
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1980
         TabIndex        =   2
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton cmdMenu 
         Height          =   270
         Left            =   675
         Picture         =   "frmMedicalStationPrintRpt.frx":020A
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   285
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "查找"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Tag             =   "姓名"
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&6.姓名"
         Height          =   180
         Index           =   1
         Left            =   1020
         TabIndex        =   1
         Tag             =   "姓名"
         Top             =   285
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   8370
      Top             =   4665
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":0490
            Key             =   "公共"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":082A
            Key             =   "报告"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":0AC0
            Key             =   "状态"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":0E5A
            Key             =   "住院"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":11F4
            Key             =   "单据"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":158E
            Key             =   "附加"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":1928
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":1BBE
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":1E54
            Key             =   "GChecked"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":20EA
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":2380
            Key             =   "Checked"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5655
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalStationPrintRpt.frx":2616
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12171
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   7950
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":2EAA
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":3624
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":3D9E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":3FB8
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":41D2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":43F2
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":4612
            Key             =   "mail"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   8625
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":4D8C
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":5506
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":5C80
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":5E9A
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":60B4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":62D4
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationPrintRpt.frx":64F4
            Key             =   "mail"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9780
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   9660
         _ExtentX        =   17039
         _ExtentY        =   1270
         ButtonWidth     =   1402
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&V.预览"
               Key             =   "预览"
               Object.ToolTipText     =   "预览(Alt+V)"
               Object.Tag             =   "&V.预览"
               ImageKey        =   "PrintView"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&P.打印"
               Key             =   "打印"
               Object.ToolTipText     =   "打印(Alt+P)"
               Object.Tag             =   "&P.打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&A.全选"
               Key             =   "全选"
               Object.ToolTipText     =   "全选(Alt+A)"
               Object.Tag             =   "&A.全选"
               ImageKey        =   "SelectAll"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&C.全清"
               Key             =   "全清"
               Object.ToolTipText     =   "全清(Alt+C)"
               Object.Tag             =   "&C.全清"
               ImageKey        =   "ClearAll"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助(Alt+H)"
               Object.Tag             =   "&H.帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出(Alt+X)"
               Object.Tag             =   "&X.退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileOutExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRptGroup 
         Caption         =   "团体报告单"
         Begin VB.Menu mnuFileRptGroupPrintView 
            Caption         =   "预览(&V)"
         End
         Begin VB.Menu mnuFileRptGroupPrint 
            Caption         =   "打印(&P)"
         End
         Begin VB.Menu mnuFileRptGroupExcel 
            Caption         =   "输出到&Excel"
         End
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSelectAll 
         Caption         =   "全选(&A)"
      End
      Begin VB.Menu mnuFileClearAll 
         Caption         =   "全清(&C)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPrintOption 
         Caption         =   "打印选项(&O)"
         Begin VB.Menu mnuViewPrint 
            Caption         =   "打印封面(&1)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewPrint 
            Caption         =   "打印报告(&2)"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuViewPrint 
            Caption         =   "打印总检(&3)"
            Checked         =   -1  'True
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "帮助主题(&T)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmMedicalStationPrintRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean
Private mblnStarted As Boolean
Private mstrMenu As String
Private mlng病人id As Long

Private WithEvents mobjPopMenu As clsPopMenu                '自定义弹出菜单对象
Attribute mobjPopMenu.VB_VarHelpID = -1
Private mbytPopMenu As Byte

'（２）自定义过程或函数************************************************************************************************
Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '值域:
    '------------------------------------------------------------------------------------------------------------------
    
    
    mnuFilePrint.Enabled = True
    mnuFilePrintView.Enabled = True
    mnuFileOutExcel.Enabled = True
        
    If vData = False Then
        mnuFilePrint.Enabled = False
        mnuFilePrintView.Enabled = False
        mnuFileOutExcel.Enabled = False
    End If
    
    
    tbrThis.Buttons("打印").Enabled = mnuFilePrint.Enabled
    tbrThis.Buttons("预览").Enabled = mnuFilePrintView.Enabled
'
    
    
End Property

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    On Error Resume Next



    On Error GoTo 0

    Call InitData

    EditChanged = True


End Function

Public Function ShowEdit(ByVal frmMain As Object, ByRef lngKey As Long, Optional lng病人id As Long = 0, Optional ByVal strMenu As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:  frmMain         调用窗体对象
    '       lngKey          预约登记id
    '返回:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
    mstrMenu = strMenu
    mlng病人id = lng病人id
    mlngKey = lngKey
    Set mfrmMain = frmMain
        
    If InitData = False Then Exit Function
    If ReadData(mlngKey, lng病人id) = False Then Exit Function
    
    
    EditChanged = (Val(vsf.RowData(1)) > 0)

    Me.Show 1, frmMain
    
    ShowEdit = mblnOK

End Function

Private Function ReadData(ByVal lngKey As Long, ByVal lng病人id As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  读取数据
    '参数:  lngKey      体检类型序号
    '返回:  True        读取成功
    '       False       读取失败
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand
            
    Select Case mstrMenu
    Case "体检指引单"
    
        gstrSQL = "SELECT 1 AS 选择,A.病人id AS ID,A.姓名,B.性别,B.门诊号,B.健康号,a.体检编号,b.就诊卡号,b.身份证号,B.婚姻状况,TO_CHAR(B.出生日期,'yyyy-mm-dd') AS 出生日期,A.组别名称 AS 组别,'' AS 未到原因 " & _
                    "FROM 体检人员档案 A,病人信息 B " & _
                    "WHERE A.体检状态>0 AND A.病人id=B.病人id and A.登记id=[1] "
        If lng病人id > 0 Then gstrSQL = gstrSQL & " AND B.病人id=[2]"
        gstrSQL = gstrSQL & " Order By B.门诊号"
    
        
    Case "项目申请单", "体检报告单"

        gstrSQL = " Select x.*, t.姓名,t.门诊号,t.健康号,t.就诊卡号,t.身份证号,t.病人id," & _
                              "y.名称 As 执行科室, " & _
                              "z.名称 As 项目, " & _
                              "DECODE(x.报告id,NULL,DECODE(d.病历文件id, NULL, '', '单据'),Decode(h.书写人, NULL, '单据', '报告')) AS 状态, " & _
                              "d.病历文件id as 单据id, " & _
                              "h.书写人 AS 报告人, " & _
                              "TO_CHAR(h.书写日期, 'yyyy-mm-dd hh24:mi') AS 时间 " & _
                         "From (Select e.id,c.病人id, " & _
                                      "a.执行科室id, " & _
                                      "a.诊疗项目id, " & _
                                      "a.结算途径, " & _
                                      "DECODE(g.执行状态,1,'完全执行',2,'取消执行',3,'正在执行','') As 执行状态, " & _
                                      "g.报告id, " & _
                                      "g.NO, " & _
                                      "Decode(a.病人id, Null, '', '附加') As 公共 " & _
                                 "From 体检项目医嘱 b, " & _
                                      "体检项目清单 a, " & _
                                      "体检人员档案 c, " & _
                                      "病人医嘱记录 e, " & _
                                      "病人医嘱发送 g " & _
                                "Where a.ID = b.清单id " & _
                                      "and b.病人id = c.病人id " & _
                                      "and c.登记id = a.登记id " & _
                                      "and e.id = b.医嘱id " & _
                                      "and e.诊疗类别 In ('C', 'D') "
            gstrSQL = gstrSQL & _
                                      "and g.医嘱id = e.id " & _
                                       "and c.登记ID = [1] " & IIf(mlng病人id > 0, " And c.病人id=[2] ", "") & _
                               " Union All " & _
                                 "Select f.id,c.病人id, " & _
                                        "a.执行科室id, " & _
                                        "a.诊疗项目id, " & _
                                        "a.结算途径, " & _
                                        "DECODE(g.执行状态,1,'完全执行',2,'取消执行',3,'正在执行','') As 执行状态, " & _
                                        "g.报告id, " & _
                                        "g.NO, " & _
                                        "Decode(a.病人id, Null, '', '附加') As 公共 " & _
                                   "From 体检项目医嘱 b, " & _
                                        "体检项目清单 a, " & _
                                        "体检人员档案 c, " & _
                                        "病人医嘱记录 e, " & _
                                        "病人医嘱记录 f, " & _
                                        "病人医嘱发送 g " & _
                                  "Where a.ID = b.清单id " & _
                                        "and b.病人id = c.病人id " & _
                                        "and c.登记id = a.登记id " & _
                                        "and e.id = b.医嘱id " & _
                                        "and e.诊疗类别 = 'E' " & _
                                        "and e.id = f.相关id " & _
                                        "and g.医嘱id = f.id "
            gstrSQL = gstrSQL & _
                                        "and c.登记ID = [1] " & IIf(mlng病人id > 0, " And c.病人id=[2] ", "") & _
                               ") x, " & _
                              "部门表 y, " & _
                              "诊疗项目目录 z, " & _
                              "诊疗单据应用 d, " & _
                              "病人病历记录 h, " & _
                              "病人信息 t " & _
                        "Where x.执行科室id = y.ID " & _
                              "and z.id = x.诊疗项目id " & _
                              "and x.报告id = h.id(+) " & _
                              "and d.应用场合(+)=4 " & _
                              "and x.诊疗项目id = d.诊疗项目id(+) and t.病人id=x.病人id Order By t.门诊号,y.名称"


    End Select
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, mlng病人id)
    If rs.BOF = False Then
    
        vsf.TextMatrix(0, GetCol(vsf, "选择")) = "选择"
        Call LoadGrid(vsf, rs, , , ils13)
        Call AppendRows(vsf, lnX, lnY)
        vsf.TextMatrix(0, GetCol(vsf, "选择")) = ""
        
    End If
    ReadData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume

End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  初始化设置
    '返回:  True        初始化成功
    '       False       初始化失败
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    
    On Error GoTo errHand
    
    Me.Caption = mstrMenu
    
    vsf.MergeCells = flexMergeFree
    
    mnuFileRptGroup.Visible = False
    mnuFile_2.Visible = False
    mnuViewPrintOption.Visible = False
    mnuView_1.Visible = False
    
    Select Case mstrMenu
    Case "体检指引单"
        
        strVsf = ",240,1,1,1,选择;姓名,1500,1,1,1,;门诊号,810,7,1,1,;健康号,810,7,1,1,;就诊卡号,0,1,1,0,;体检编号,990,1,1,1,;身份证号,900,1,1,0,;性别,810,1,1,1,;出生日期,1080,1,1,1,;婚姻状况,1200,1,1,1,;组别,1200,1,1,1,"
        
    Case "项目申请单", "体检报告单"
        
        strVsf = "姓名,750,1,1,1,;门诊号,810,7,1,1,;健康号,810,7,1,1,;就诊卡号,0,1,1,0,;体检编号,990,1,1,1,;身份证号,900,1,1,0,;,240,1,1,1,选择;项目,2400,1,1,1,;执行科室,1080,1,1,1,;执行状态,900,1,1,1,;报告id,0,1,1,1,;单据id,0,1,1,1,;No,0,1,1,1,;报告来源,0,1,1,1,;病人id,0,1,1,0,;,255,4,1,1,[公共];,255,4,1,1,[状态]"
        
        If mstrMenu = "体检报告单" Then
        
            mnuFileRptGroup.Visible = (mlng病人id = 0)
            mnuFile_2.Visible = (mlng病人id = 0)
            mnuViewPrintOption.Visible = True
            mnuView_1.Visible = True
            
        End If
        
    End Select
    
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(GetCol(vsf, "选择")) = flexDTBoolean
    vsf.Editable = flexEDKbdMouse
    
    Select Case mstrMenu
    Case "项目申请单", "体检报告单"
        Set vsf.Cell(flexcpPicture, 0, GetCol(vsf, "[公共]")) = ils13.ListImages("公共").Picture
        Set vsf.Cell(flexcpPicture, 0, GetCol(vsf, "[状态]")) = ils13.ListImages("状态").Picture
        vsf.MergeCol(GetCol(vsf, "姓名")) = True
        vsf.MergeCol(GetCol(vsf, "门诊号")) = True
        vsf.MergeCol(GetCol(vsf, "健康号")) = True
        vsf.MergeCol(GetCol(vsf, "就诊卡号")) = True
    End Select
    
    Call AppendRows(vsf, lnX, lnY)
    
    InitData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  校验数据的有效性
    '返回:  True        数据有效
    '       False       数据无效
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long


    ValidEdit = True

End Function

Private Function GetReportCode(ByVal lngKey As Long, ByRef strCode As String, ByRef strNo As String, ByRef bytMode As Byte) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能;
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    If lngKey = 0 Then Exit Function
    

        strSQL = "SELECT DISTINCT 'ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-1' AS 报表编号," & _
                           "D.NO," & _
                           "D.记录性质 " & _
                    "FROM 病历文件目录 C,(SELECT A.NO,A.记录性质,E.病历文件id FROM 病人医嘱发送 A,病人医嘱记录 B,诊疗单据应用 E WHERE E.应用场合=4 AND E.诊疗项目id=B.诊疗项目id AND B.诊疗类别 IN ('C','D') AND A.医嘱id=B.ID AND (B.相关id=[1] OR B.ID=[1]) AND ROWNUM<2) D " & _
                    "Where C.ID=D.病历文件id"
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlLisWork", lngKey)
    If rs.BOF = False Then
        strCode = zlCommFun.NVL(rs("报表编号"))
        strNo = zlCommFun.NVL(rs("NO"))
        bytMode = zlCommFun.NVL(rs("记录性质"), 1)
    End If
    
    GetReportCode = True
    
End Function

Private Function PrintData(ByVal bytMode As Byte, Optional ByVal blnGroup As Boolean) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  保存数据
    '返回:  True        保存成功
    '       False       保存失败
    '------------------------------------------------------------------------------------------------------------------
    Dim strReportCode As String
    Dim lngLoop As Long
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim int选择 As Integer
    Dim strSQL As String
    Dim int报告id As Integer
    Dim int门诊号 As Integer
    Dim int病人id As Integer
    Dim rs As New ADODB.Recordset
    Dim strSvr体检编号 As String
    On Error GoTo errHand
    
    Select Case mstrMenu
    Case "体检指引单"
        strReportCode = "ZL1_BILL_1861"
    Case "体检报告单"
        strReportCode = "ZL1_BILL_1861_2"
    End Select
    
    int选择 = GetCol(vsf, "选择")
    
    If blnGroup Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1861_3", Me, "登记id=" & mlngKey, bytMode)
    Else
        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.RowData(lngLoop)) > 0 And Abs(Val(vsf.TextMatrix(lngLoop, int选择))) = 1 Then
                
                Select Case mstrMenu
                Case "体检指引单"
                
                    Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "登记id=" & mlngKey, "病人id=" & Val(vsf.RowData(lngLoop)), bytMode)
                    
                Case "项目申请单"
                    
                    If GetReportCode(Val(vsf.RowData(lngLoop)), strReportCode, strReportParaNo, bytReportParaMode) Then
                        Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, bytMode)
                    End If
                
                Case "体检报告单"
                    
                    If int报告id = 0 Then int报告id = GetCol(vsf, "报告id")
                    If int门诊号 = 0 Then int门诊号 = GetCol(vsf, "门诊号")
                    If int病人id = 0 Then int病人id = GetCol(vsf, "病人id")
                    
                    If strSvr体检编号 <> vsf.TextMatrix(lngLoop, int门诊号) Then
                    
                        strSvr体检编号 = vsf.TextMatrix(lngLoop, int门诊号)
                        
                        '1.调用"报告封面"
                        If bytMode <> 1 And mnuViewPrint(0).Checked Then
                            Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "登记id=" & mlngKey, "病人id=" & Val(vsf.TextMatrix(lngLoop, int病人id)), "报告id=0", "ReportFormat=2", bytMode)
                        End If
                    End If
                    
                    '2.调用"项目报告",缺省调用
                    If mnuViewPrint(1).Checked Or bytMode = 1 Then
                        Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "登记id=" & mlngKey, "病人id=" & Val(vsf.TextMatrix(lngLoop, int病人id)), "报告id=" & Val(vsf.TextMatrix(lngLoop, int报告id)), "ReportFormat=1", bytMode)
                    End If
                                                                               
                    If lngLoop < vsf.Rows - 1 Then
                        If strSvr体检编号 <> vsf.TextMatrix(lngLoop + 1, int门诊号) Then
                            '3.调用"体检总检"
                            If bytMode <> 1 And mnuViewPrint(2).Checked Then
                                Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "登记id=" & mlngKey, "病人id=" & Val(vsf.TextMatrix(lngLoop, int病人id)), "报告id=0", "ReportFormat=3", bytMode)
                            End If
                            
                        End If
                    Else
                        '3.调用"体检总检"
                        If bytMode <> 1 And mnuViewPrint(2).Checked Then
                            Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "登记id=" & mlngKey, "病人id=" & Val(vsf.TextMatrix(lngLoop, int病人id)), "报告id=0", "ReportFormat=3", bytMode)
                        End If
                    End If
                    
                End Select
                
                '如果是预览，只一次预览
                If bytMode = 1 Then Exit For
                
            End If
        Next
    End If
      
    PrintData = True

    Exit Function

errHand:

    If ErrCenter = 1 Then
        Resume
    End If

End Function


Private Sub cmdMenu_Click()
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(cmdMenu.hWnd, objPoint)
    
    mbytPopMenu = 3
    Set mobjPopMenu = New clsPopMenu
    Call mobjPopMenu.ShowPopupMenu(objPoint.X * Screen.TwipsPerPixelX, objPoint.Y * Screen.TwipsPerPixelY - 255 * 8)
    
    txt(1).Text = ""
    LocationObj txt(1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyA
            If tbrThis.Buttons("全选").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("全选"))
        Case vbKeyC
            If tbrThis.Buttons("全清").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("全清"))
        Case vbKeyM
            If tbrThis.Buttons("邮件").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("邮件"))
        Case vbKeyV
            If tbrThis.Buttons("预览").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("预览"))
        Case vbKeyP
            If tbrThis.Buttons("打印").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("打印"))
        Case vbKeyH
            If tbrThis.Buttons("帮助").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("帮助"))
        Case vbKeyX
            If tbrThis.Buttons("退出").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("退出"))
        End Select
    ElseIf Shift = 0 Then
        If KeyCode = vbKeyEscape Then
            If tbrThis.Buttons("退出").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("退出"))
        End If
    End If
End Sub

'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub Form_Load()

    Call RestoreWinState(Me, App.ProductName)
    
    If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", "0")) = 1 Then
        '使用个性化设置
      
        lbl(1).Caption = "&6." & (GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "查找信息", "姓名"))
        lbl(1).Tag = Mid(lbl(1).Caption, 4)

    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    With vsf
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0) - fraInfo.Height + 90
    End With
    With fraInfo
        .Left = 0
        .Top = vsf.Top + vsf.Height - 75
        .Width = vsf.Width
    End With
    
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "查找信息", lbl(1).Tag)
End Sub


Private Sub mnuFileClearAll_Click()
    Dim lngLoop As Long
    Dim int选择 As Integer
    
    int选择 = GetCol(vsf, "选择")
    If int选择 >= 0 Then
    
        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.RowData(lngLoop)) > 0 Then
                vsf.TextMatrix(lngLoop, int选择) = 0
            End If
        Next
        
        EditChanged = False
        
    End If
    
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub


Private Sub mnuFileOutExcel_Click()
    Call PrintData(3)
End Sub

Private Sub mnuFilePrint_Click()
    
    Call PrintData(2)

End Sub

Private Sub mnuFilePrintView_Click()
    
    Call PrintData(1)
    
End Sub

Private Sub mnuFileRptGroupExcel_Click()
    Call PrintData(3, True)
End Sub

Private Sub mnuFileRptGroupPrint_Click()
    Call PrintData(2, True)
End Sub

Private Sub mnuFileRptGroupPrintView_Click()
    Call PrintData(1, True)
End Sub

Private Sub mnuFileSelectAll_Click()
    Dim lngLoop As Long
    Dim int选择 As Integer
    
    int选择 = GetCol(vsf, "选择")
    If int选择 >= 0 Then
        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.RowData(lngLoop)) > 0 Then
                vsf.TextMatrix(lngLoop, int选择) = 1
                EditChanged = True
            End If
        Next
    End If
    
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewPrint_Click(Index As Integer)
    mnuViewPrint(Index).Checked = Not mnuViewPrint(Index).Checked
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intLoop As Integer

    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize

End Sub

Private Sub mobjPopMenu_MenuBeforeShow(Cancel As Boolean)
    
    Select Case mbytPopMenu

    Case 3
        
        mobjPopMenu.Add 1, "&1.姓名", , , True, , (lbl(1).Tag = "姓名")
        mobjPopMenu.Add 2, "&2.门诊号", , , True, , (lbl(1).Tag = "门诊号")
        mobjPopMenu.Add 3, "&3.健康号", , , True, , (lbl(1).Tag = "健康号")
        mobjPopMenu.Add 4, "&4.就诊卡号", , , True, , (lbl(1).Tag = "就诊卡号")
        mobjPopMenu.Add 5, "&5.姓名拼音", , , True, , (lbl(1).Tag = "姓名拼音")
        mobjPopMenu.Add 6, "&6.姓名五笔", , , True, , (lbl(1).Tag = "姓名五笔")
        mobjPopMenu.Add 7, "&7.身份证号", , , True, , (lbl(1).Tag = "身份证号")
        mobjPopMenu.Add 8, "&8.体检编号", , , True, , (lbl(1).Tag = "体检编号")
        
    End Select
    
End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    Select Case mbytPopMenu

    Case 3
    
        Caption = Mid(Caption, 4)
        
        lbl(1).Caption = "&6." & Left(Trim(Caption), Len(Trim(Caption)) - 1)
        lbl(1).Tag = Left(Trim(Caption), Len(Trim(Caption)) - 1)
        
    End Select
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "全选"
        Call mnuFileSelectAll_Click
    Case "全清"
        Call mnuFileClearAll_Click
    Case "预览"
        Call mnuFilePrintView_Click
    Case "打印"
        Call mnuFilePrint_Click
    Case "邮件"
        
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Call tbrThis_ButtonClick(Button)
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngLoop As Long
    Dim strCol As String
    Dim lngCol As Long
    Dim lngRow As Long
    Dim blnCard As Boolean
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
    strCol = Mid(lbl(1).Caption, 4)
    lngCol = GetCol(vsf, strCol)
            
    If strCol = "就诊卡号" And KeyAscii <> vbKeyReturn Then
        '就诊卡号，自动识别

        blnCard = InputIsCard(txt(Index).Text, KeyAscii)

        If blnCard And Len(txt(Index).Text) = ParamInfo.就诊卡号码长度 - 1 And KeyAscii <> 8 And txt(Index).Text <> "" Then
            If KeyAscii <> 13 Then
                txt(Index).Text = txt(Index).Text & Chr(KeyAscii)
                txt(Index).SelStart = Len(txt(Index).Text)
            End If
            KeyAscii = vbKeyReturn
        End If

    End If
    
    If KeyAscii = vbKeyReturn Then
        
        If Index = 1 And Trim(txt(Index).Text) <> "" Then
            
            strCol = Mid(lbl(1).Caption, 4)
            Select Case strCol
            Case "姓名拼音"
                lngCol = GetCol(vsf, "姓名")
            Case "姓名五笔"
                lngCol = GetCol(vsf, "姓名")
            Case Else
                lngCol = GetCol(vsf, strCol)
            End Select
            If lngCol < 0 Then Exit Sub
            
            lngRow = 0
            If vsf.Row + 1 <= vsf.Rows - 1 Then
                For lngLoop = vsf.Row + 1 To vsf.Rows - 1
                    
                    lngRow = 0
                    Select Case strCol
                    Case "门诊号"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "健康号"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "就诊卡号"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "身份证号"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名拼音"
                        If zlGetSymbol(UCase(vsf.TextMatrix(lngLoop, lngCol))) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名五笔"
                        If zlGetSymbol(UCase(vsf.TextMatrix(lngLoop, lngCol)), 1) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case Else
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    End Select
            
                    If lngRow > 0 Then Exit For

                Next
            End If
            
            If lngRow = 0 Then
                For lngLoop = 1 To vsf.Row
                
                    lngRow = 0
                    Select Case strCol
                    Case "门诊号"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "健康号"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "就诊卡号"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "身份证号"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名"
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名拼音"
                        If zlGetSymbol(UCase(vsf.TextMatrix(lngLoop, lngCol))) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名五笔"
                        If zlGetSymbol(UCase(vsf.TextMatrix(lngLoop, lngCol)), 1) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case Else
                        If UCase(vsf.TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    End Select
                    
                    If lngRow > 0 Then Exit For
                Next
            End If
            
            If lngRow <= 0 Then
                ShowSimpleMsg "没有找到符合要求的信息！"
                txt(Index).Text = ""
            Else
                vsf.ShowCell lngRow, vsf.Col
                vsf.Row = lngRow
            End If
        End If
        
        txt(Index).SetFocus
        zlControl.TxtSelAll txt(Index)
    End If
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngLoop As Long
    Dim int选择 As Integer
    
    int选择 = GetCol(vsf, "选择")
    
    If int选择 >= 0 Then
        If Abs(Val(vsf.TextMatrix(Row, int选择))) = 1 Then
            EditChanged = True
            Exit Sub
        End If
            
        For lngLoop = 1 To vsf.Rows - 1
            If Abs(Val(vsf.TextMatrix(lngLoop, int选择))) = 1 Then
                EditChanged = True
                Exit Sub
            End If
        Next
        
        If lngLoop = vsf.Rows Then EditChanged = False
    End If
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> GetCol(vsf, "选择") Or Val(vsf.RowData(Row)) <= 0 Then
        Cancel = True
    End If
End Sub



Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

