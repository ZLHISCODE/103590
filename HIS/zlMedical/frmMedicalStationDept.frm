VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmMedicalStationDept 
   Caption         =   "执行科室调整"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10350
   Icon            =   "frmMedicalStationDept.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   10350
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   0
      Left            =   90
      ScaleHeight     =   435
      ScaleWidth      =   5565
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1305
      Width           =   5565
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   90
         Width           =   3720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&6.体检项目"
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   150
         Width           =   900
      End
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   1
      Left            =   345
      ScaleHeight     =   435
      ScaleWidth      =   5565
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4395
      Width           =   5565
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   1815
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   105
         Width           =   2280
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&7.调整为新执行科室"
         Height          =   180
         Index           =   5
         Left            =   60
         TabIndex        =   0
         Top             =   150
         Width           =   1620
      End
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   10350
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
         TabIndex        =   3
         Top             =   30
         Width           =   10230
         _ExtentX        =   18045
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
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&A.全选"
               Key             =   "全选"
               Object.ToolTipText     =   "全选(Alt+A)"
               Object.Tag             =   "&A.全选"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&C.全清"
               Key             =   "全清"
               Object.ToolTipText     =   "全清(Alt+C)"
               Object.Tag             =   "&C.全清"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&D.调整"
               Key             =   "调整"
               Object.ToolTipText     =   "调整(Alt+D)"
               Object.Tag             =   "&D.调整"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助(Alt+H)"
               Object.Tag             =   "&H.帮助"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出(Alt+X)"
               Object.Tag             =   "&X.退出"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   4470
      Top             =   750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":19EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":2166
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":28E0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":2B00
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   5460
      Top             =   750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":2D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":349A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":3C14
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":438E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":45AE
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   5865
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalStationDept.frx":47CE
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13176
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
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1980
      Left            =   165
      TabIndex        =   1
      Top             =   1905
      Width           =   3795
      _cx             =   6694
      _cy             =   3492
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
      Cols            =   2
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
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   825
         X2              =   825
         Y1              =   0
         Y2              =   1215
      End
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   1785
         Y1              =   135
         Y2              =   135
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   8430
      Top             =   3855
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
            Picture         =   "frmMedicalStationDept.frx":5062
            Key             =   "报告"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":53FC
            Key             =   "状态"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationDept.frx":5796
            Key             =   "单据"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Height          =   4875
      Left            =   7020
      TabIndex        =   10
      Top             =   855
      Width           =   3285
      Begin VB.OptionButton opt 
         Caption         =   "找到后处理为未选中(&5)"
         Height          =   210
         Index           =   1
         Left            =   1005
         TabIndex        =   20
         Top             =   2325
         Width           =   2205
      End
      Begin VB.OptionButton opt 
         Caption         =   "找到后处理为选中(&4)"
         Height          =   210
         Index           =   0
         Left            =   1005
         TabIndex        =   19
         Top             =   1965
         Value           =   -1  'True
         Width           =   2205
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   3
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   555
         Width           =   1995
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&3.按姓名定位"
         Height          =   240
         Left            =   4710
         TabIndex        =   14
         Top             =   210
         Width           =   1425
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "查找(&F)"
         Height          =   350
         Left            =   1125
         TabIndex        =   13
         Top             =   1365
         Width           =   1470
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1125
         TabIndex        =   12
         Top             =   930
         Width           =   1995
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   2
         Left            =   1110
         TabIndex        =   11
         Text            =   "cbo"
         Top             =   195
         Width           =   1995
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&1.工作单位"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   255
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&2.组    别"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   630
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&3.门 诊 号"
         Height          =   180
         Index           =   3
         Left            =   135
         TabIndex        =   16
         Tag             =   "门诊号"
         Top             =   1005
         Width           =   900
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileSelectAll 
         Caption         =   "全选(&A)"
      End
      Begin VB.Menu mnuFileClearAll 
         Caption         =   "全清(&C)"
      End
      Begin VB.Menu mnuFile_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "调整(&D)"
         Shortcut        =   ^D
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
Attribute VB_Name = "frmMedicalStationDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Private mblnOK As Boolean
Private mblnStartUp As Boolean
Private mfrmMain As Form
Private mlngLoop As Long
Private mstrSQL As String
Private mblnChangeEdit As Boolean

Private mlngKey As Long

Private Enum mCol
    选择 = 0
    门诊号
    姓名
    执行科室
    工作单位
    组别
    病人id
    执行科室id
End Enum

Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '值域:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long

    mnuFileSave.Enabled = True
        
    If vData = False Then
        mnuFileSave.Enabled = False
    End If

        tbrThis.Buttons("调整").Enabled = mnuFileSave.Enabled

    
End Property

Private Sub RefreshState()
    
    Dim lngLoop As Long
    Dim intCount As Integer
    
    intCount = 0
    For lngLoop = 1 To vsf.Rows - 1
        If Abs(Val(vsf.TextMatrix(lngLoop, 0))) = 1 Then
            intCount = intCount + 1
        End If
    Next
    
    stbThis.Panels(2).Text = "当前选中 " & intCount & " 人"
End Sub

Private Function SearchSelect(ByVal blnSel As Boolean) As Boolean
    '==================================================================================================================
    
    Dim lngRow As Long
    Dim lngCol As Long
    Dim blnFind1 As Boolean
    Dim blnFind2 As Boolean
    Dim blnFind3 As Boolean
    Dim blnFind4 As Boolean
    Dim lngStartRow As Long

    lngStartRow = vsf.Row
    For lngRow = lngStartRow To vsf.Rows - 1
        
        blnFind1 = True
        blnFind2 = True
        blnFind3 = True
        blnFind4 = True
        
        If cbo(2).Text <> "" Then
            blnFind1 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.工作单位)), UCase(cbo(2).Text)) > 0 Then
                blnFind1 = True
            End If
        End If
        
        If txt(1).Text <> "" Then
            blnFind2 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.门诊号)), UCase(txt(1).Text)) > 0 Then
                blnFind2 = True
            End If
        End If
        
        If cbo(3).Text <> "" Then
            blnFind4 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.组别)), UCase(cbo(3).Text)) > 0 Then
                blnFind4 = True
            End If
        End If
        
        If blnFind1 And blnFind2 And blnFind3 And blnFind4 Then
            '找到

            vsf.TextMatrix(lngRow, mCol.选择) = IIf(blnSel, 1, 0)
            SearchSelect = True
            vsf.Row = lngRow

        End If
    Next
    
    For lngRow = 1 To lngStartRow
        
        blnFind1 = True
        blnFind2 = True
        blnFind3 = True
        blnFind4 = True
        
        If cbo(1).Text <> "" Then
            blnFind1 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.工作单位)), UCase(cbo(1).Text)) > 0 Then
                blnFind1 = True
            End If
        End If
        
        If txt(1).Text <> "" Then
            blnFind2 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.门诊号)), UCase(txt(1).Text)) > 0 Then
                blnFind2 = True
            End If
        End If
        
        If cbo(0).Text <> "" Then
            blnFind4 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.组别)), UCase(cbo(0).Text)) > 0 Then
                blnFind4 = True
            End If
        End If
        
        If blnFind1 And blnFind2 And blnFind3 And blnFind4 Then
        
            vsf.TextMatrix(lngRow, mCol.选择) = IIf(blnSel, 1, 0)
            SearchSelect = True
            
            vsf.Row = lngRow
            
        End If
    Next
    
    If SearchSelect Then
        vsf.ShowCell vsf.Row, vsf.Col
        vsf.SetFocus
    End If
    
End Function

Private Sub AdjustEnableState()
    '-----------------------------------------------------------------------------------------
    '功能:根据修改状态设置按钮、菜单等的可用状态
    '-----------------------------------------------------------------------------------------
    
    mnuFileSave.Enabled = True
        
    If mblnChangeEdit = False Then mnuFileSave.Enabled = False
        
    tbrThis.Buttons("调整").Enabled = mnuFileSave.Enabled
        
End Sub

Private Sub RefreshStatus()
    '-----------------------------------------------------------------------------------------
    '功能:
    '-----------------------------------------------------------------------------------------
    If vsf.Rows = 2 And Trim(vsf.TextMatrix(1, 1)) = "" Then
        stbThis.Panels(2).Text = "没有信息。"
    Else
        stbThis.Panels(2).Text = "共找到 " & vsf.Rows - 1 & " 个信息。"
    End If
    
End Sub

Public Function ShowEdit(ByVal frmMain As Form, ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：显示本编辑窗体
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
    
    mlngKey = lngKey
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
                    
    mblnChangeEdit = False
    Call AdjustEnableState
    mblnStartUp = False
    
    Call cbo_Click(0)
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strVsf As String
    
    On Error GoTo errHand
    
    strVsf = "选择,540,4,1,1,;门诊号,900,7,1,1,;姓名,900,1,1,1,;执行科室,1200,1,1,1,;工作单位,1500,1,1,1,;组别,1500,1,1,1,;病人id,0,1,1,0,;执行科室id,0,1,1,0,"
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(mCol.选择) = flexDTBoolean
    vsf.Editable = flexEDKbdMouse

    '读取项目清单
    gstrSQL = "Select Distinct b.名称,b.ID from 体检项目清单 a,诊疗项目目录 b where a.诊疗项目id=b.id and a.登记id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
    Call AddComboData(cbo(0), rs)
    If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
    
    '读取医技科室
    gstrSQL = "Select Distinct b.名称,b.ID from 部门表 b where b.id in (select 部门id From 部门性质说明 Where 工作性质 In ('检查','检验'))"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    Call AddComboData(cbo(1), rs)
    If cbo(1).ListCount > 0 And cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0


    gstrSQL = "SELECT Distinct B.工作单位 " & _
                "FROM 体检人员档案 A,病人信息 B " & _
                "WHERE B.工作单位 Is Not Null And A.体检状态 IN (1,4) AND A.体检报到 In ([2],[3]) AND A.病人id=B.病人id and A.登记id=[1]"
                
    cbo(2).Clear
    cbo(2).AddItem ""
    cbo(2).ListIndex = 0
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, 0, 1)
    If rs.BOF = False Then
        Do While Not rs.EOF
            cbo(2).AddItem rs("工作单位").Value
            rs.MoveNext
        Loop
    End If
    
    gstrSQL = "SELECT Distinct A.组别名称 " & _
                "FROM 体检人员档案 A " & _
                "WHERE A.组别名称 Is Not Null And A.体检状态 IN (1,4) AND A.登记id=[1]"
                    
    cbo(3).Clear
    cbo(3).AddItem ""
    cbo(3).ListIndex = 0
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
    If rs.BOF = False Then
        Do While Not rs.EOF
            cbo(3).AddItem rs("组别名称").Value
            rs.MoveNext
        Loop
    End If
    
    InitData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strWhere As String
    Dim varTmp2 As Variant
    Dim strTmp As String
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    Call ResetVsf(vsf)
    Call AppendRows(vsf, lnX, lnY)
    
    mstrSQL = "Select 0 As 选择,y.病人id,y.病人id AS ID,y.姓名,y.门诊号,y.健康号,y.工作单位,x.报告id,x.执行科室id,z.名称 As 执行科室,x.组别名称 As 组别 From " & _
                "( " & _
                "Select Distinct  b.病人id,f.报告id,a.执行科室id,b.组别名称 " & _
                "from 体检项目医嘱 a,体检人员档案 b,体检项目清单 c,体检登记记录 d,病人医嘱记录 e,病人医嘱发送 f " & _
                "Where c.诊疗项目id = [1] and c.登记id=[2] " & _
                      "and a.清单id=c.id " & _
                      "and c.登记id=b.登记id " & _
                      "and a.病人id=b.病人id " & _
                      "and d.id=b.登记id " & _
                      "and e.挂号单=d.体检号 " & _
                      "and e.病人来源=4 and b.体检报到=1 And b.体检状态=4 " & _
                      "and e.医嘱状态<>4 " & _
                      "and e.诊疗类别 In ('C','D') and e.诊疗项目id=c.诊疗项目id and e.病人id=b.病人id " & _
                      "and f.医嘱id(+)=e.id " & _
                ") x,病人信息 y,部门表 z " & _
                "where x.病人id=y.病人id and z.id=x.执行科室id and x.报告id Is Null"
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngKey, mlngKey)
    If rs.BOF = False Then
        Call LoadGrid(vsf, rs)
        Call AppendRows(vsf, lnX, lnY)
    End If
    
    ReadData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ValidData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim strError As String
    Dim rs As New ADODB.Recordset
    
    '检验输入的新值是否正确,主要是检验公式
    
    On Error GoTo errHand
    
    
            
    ValidData = True
    
    Exit Function
errHand:
    LocationObj txt(1)
    strError = "调整结果值或公式不合法！"
    MsgBox strError, vbInformation, gstrSysName
End Function

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim str采集No As String
    Dim strNo As String
    Dim lngSendNo As Long
    Dim lngDept As Long
    Dim lngTotal As Long
    Dim lngCount As Long
    
    On Error GoTo errHand
    
    Me.Enabled = False
    Call frmWait.OpenWait(Me, "调整执行地点")
    frmWait.WaitInfo = "正在调整执行地点..."
    
    lngSendNo = GetNextNo(10)
    lngDept = mfrmMain.cboDept.ItemData(mfrmMain.cboDept.ListIndex)
    
    gstrSQL = "Select a.*,b.病人id As 病人 From 体检项目清单 a,体检项目医嘱 b Where a.登记id=[1] and a.诊疗项目id=[2] And a.id=b.清单id and b.执行科室id<> [3] "
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, cbo(0).ItemData(cbo(0).ListIndex), cbo(1).ItemData(cbo(1).ListIndex))
            
    lngTotal = vsf.Rows - 1
    For mlngLoop = 1 To lngTotal
        If Abs(Val(vsf.TextMatrix(mlngLoop, mCol.选择))) = 1 And Val(vsf.RowData(mlngLoop)) > 0 Then

            frmWait.WaitInfo = "正在调整执行地点“" & vsf.TextMatrix(mlngLoop, mCol.姓名) & " ”..." & Format(100 * mlngLoop / lngTotal, "0.00") & "%"
            
            rs.Filter = ""
            rs.Filter = "病人=" & Val(vsf.RowData(mlngLoop))
            If rs.RecordCount > 0 Then
                
                Call SQLRecord(rsSQL)
                
                strSQL = "zl_体检项目医嘱_Modify(" & mlngKey & "," & cbo(0).ItemData(cbo(0).ListIndex) & "," & Val(vsf.RowData(mlngLoop)) & "," & cbo(1).ItemData(cbo(1).ListIndex) & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
                
                '产生费用单据号
                str采集No = ""
                strNo = ""
                If Val(zlCommFun.NVL(rs("结算途径").Value, 1)) = 1 Then
                    '记帐
                    strNo = GetNextNo(14)
                Else
                    strNo = GetNextNo(13)
                End If
                
                If Val(zlCommFun.NVL(rs("采集方式id").Value, 0)) > 0 Then
                    '采集
                    If Val(zlCommFun.NVL(rs("结算途径").Value, 1)) = 1 Then
                        '记帐
                        str采集No = GetNextNo(14)
                    Else
                        str采集No = GetNextNo(13)
                    End If
                End If
                
                
                strSQL = "ZL_体检项目医嘱_NO(" & rs("ID").Value & "," & Val(vsf.RowData(mlngLoop)) & ",'" & strNo & "','" & str采集No & "')"
'                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                Call SQLRecordAdd(rsSQL, strSQL)
                
                strSQL = "zl_体检人员档案_Accept(" & mlngKey & "," & lngSendNo & "," & Val(vsf.RowData(mlngLoop)) & "," & lngDept & "," & rs("ID").Value & ",1)"
'                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                Call SQLRecordAdd(rsSQL, strSQL)
                
                blnTran = True
                gcnOracle.BeginTrans
                
                If rsSQL.RecordCount > 0 Then rsSQL.MoveFirst
                For lngCount = 1 To rsSQL.RecordCount
                    Call zlDatabase.ExecuteProcedure(CStr(rsSQL("SQL").Value), Me.Caption)
                    rsSQL.MoveNext
                Next
                
                
                '产生相关费用
                If MakeMedicalCharge(rsSQL, mlngKey) = False Then
                    gcnOracle.RollbackTrans
                    blnTran = False
                    Exit Function
                End If
                
                strSQL = "zl_体检人员档案_Accept(" & mlngKey & "," & lngSendNo & "," & Val(vsf.RowData(mlngLoop)) & "," & lngDept & "," & rs("ID").Value & ",2)"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                
                gcnOracle.CommitTrans
                blnTran = False
            End If
        End If
    Next
    
    frmWait.CloseWait
    Me.Enabled = True
    
    SaveData = True

    Exit Function

errHand:
    frmWait.CloseWait
    Me.Enabled = True
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTran Then
        gcnOracle.RollbackTrans
        ShowSimpleMsg "未完全调整成功或部份调整成功！"
    End If
End Function

Private Sub cbo_Click(Index As Integer)
    If mblnStartUp Then Exit Sub
    
    If Index = 0 Then
        If cbo(Index).ListIndex >= 0 Then
            Call ReadData(cbo(Index).ItemData(cbo(Index).ListIndex))
        End If
    ElseIf Index > 1 Then

    End If
    
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_Click(Index As Integer)
    zlControl.TxtSelAll txt(1)
    txt(1).SetFocus
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdSelect_Click()
    If SearchSelect(opt(0).Value) Then

        EditChanged = True
        Call RefreshState
        
    End If
    zlControl.TxtSelAll txt(1)
    txt(1).SetFocus
End Sub

Private Sub Form_Activate()
    Dim rs As New ADODB.Recordset
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    '检验部门
'    cbo(1).Clear
'    For mlngLoop = 0 To mfrmMain.cboDept.ListCount - 1
'        cbo(1).AddItem mfrmMain.cboDept.List(mlngLoop)
'        cbo(1).ItemData(cbo(1).NewIndex) = mfrmMain.cboDept.ItemData(mlngLoop)
'    Next
'    cbo(1).ListIndex = mfrmMain.cboDept.ListIndex
'
'
'    If cbo(1).ListIndex = -1 Then
'        zlControl.CboLocate cbo(1), UserInfo.部门ID, True
'        If cbo(1).ListCount > 0 And cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
'    End If
    
'    txt(0).Text = ""
'    txt(2).Text = ""

    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyA
            If tbrThis.Buttons("全选").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("全选"))
        Case vbKeyC
            If tbrThis.Buttons("全清").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("全清"))
        Case vbKeyB
            If tbrThis.Buttons("调整").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("调整"))
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

Private Sub Form_Load()
    
    Call RestoreWinState(Me, App.ProductName)
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picBack(0).Move 0, IIf(cbrThis.Visible, cbrThis.Height, 0), Me.ScaleWidth - fra.Width
    vsf.Move 0, picBack(0).Top + picBack(0).Height, picBack(0).Width, Me.ScaleHeight - picBack(0).Top - picBack(0).Height - IIf(stbThis.Visible, stbThis.Height, 0) - picBack(1).Height
    picBack(1).Move 0, vsf.Top + vsf.Height, vsf.Width
    
    fra.Move vsf.Left + vsf.Width, picBack(0).Top - 90, fra.Width, Me.ScaleHeight - fra.Top - IIf(stbThis.Visible, stbThis.Height, 0) + 90

    Call AppendRows(vsf, lnX, lnY)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If mblnChangeEdit Then
        Cancel = (MsgBox("新增或修改的数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mnuFileClearAll_Click()
    For mlngLoop = 1 To vsf.Rows - 1
        vsf.TextMatrix(mlngLoop, mCol.选择) = 0
    Next
    mblnChangeEdit = False
    Call AdjustEnableState
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileSave_Click()
    If mblnChangeEdit Then
        
        If MsgBox("真的要将选中的项目调整为新执行科室吗？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        If SaveData() = False Then Exit Sub
        
        mblnOK = True
        
        mblnChangeEdit = False
        Call AdjustEnableState

        ShowSimpleMsg "体检执行科室调整成功！"
        
        Call ResetVsf(vsf)
        txt(1).Text = ""
        Call AppendRows(vsf, lnX, lnY)
        
        Exit Sub
    End If
End Sub

Private Sub mnuFileSelectAll_Click()
    For mlngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(mlngLoop)) > 0 Then
            vsf.TextMatrix(mlngLoop, mCol.选择) = 1
        End If
    Next
    
    mblnChangeEdit = True
    Call AdjustEnableState
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

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "全选"
        Call mnuFileSelectAll_Click
    Case "全清"
        Call mnuFileClearAll_Click
    Case "调整"
        Call mnuFileSave_Click
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub


Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
    
    cmdSelect.Default = True
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 0
        zlCommFun.OpenIme False
    End Select
    
    cmdSelect.Default = False
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub


Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    If Abs(Val(vsf.TextMatrix(Row, mCol.选择))) = 1 Then
        mblnChangeEdit = True
        Call AdjustEnableState
        Exit Sub
    End If
        
    For mlngLoop = 1 To vsf.Rows - 1
        If Abs(Val(vsf.TextMatrix(mlngLoop, mCol.选择))) = 1 Then
            mblnChangeEdit = True
            Call AdjustEnableState
            Exit Sub
        End If
    Next
    
    If mlngLoop = vsf.Rows Then
        mblnChangeEdit = False
        Call AdjustEnableState
    End If
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(vsf.RowData(Row)) = 0 Then Cancel = True
    If Col <> 0 Then Cancel = True
    
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

