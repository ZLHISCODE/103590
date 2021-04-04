VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmMedicalStationRptBook 
   Caption         =   "#"
   ClientHeight    =   7800
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   11820
   Icon            =   "frmMedicalStationRptBook.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11820
   Begin zl9Medical.VsfGrid vsf 
      Height          =   2685
      Index           =   1
      Left            =   3120
      TabIndex        =   18
      Top             =   3000
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   4736
   End
   Begin zl9Medical.VsfGrid vsf 
      Height          =   1470
      Index           =   0
      Left            =   2850
      TabIndex        =   17
      Top             =   1140
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   2593
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
            Picture         =   "frmMedicalStationRptBook.frx":020A
            Key             =   "公共"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":05A4
            Key             =   "报告"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":083A
            Key             =   "状态"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":0BD4
            Key             =   "住院"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":0F6E
            Key             =   "单据"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":1308
            Key             =   "附加"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":16A2
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":1938
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":1BCE
            Key             =   "GChecked"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":1E64
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":20FA
            Key             =   "Checked"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   7440
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalStationRptBook.frx":2390
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15769
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
            Picture         =   "frmMedicalStationRptBook.frx":2C24
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":339E
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":3B18
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":3D32
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":3F4C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":416C
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":438C
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
            Picture         =   "frmMedicalStationRptBook.frx":4B06
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":5280
            Key             =   "ClearAll"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":59FA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":5C14
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":5E2E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":604E
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationRptBook.frx":626E
            Key             =   "mail"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   11820
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   1138
         ButtonWidth     =   1296
         ButtonHeight    =   1138
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
   Begin VB.Frame fra1 
      Height          =   1770
      Left            =   90
      TabIndex        =   8
      Top             =   765
      Width           =   2580
      Begin VB.CheckBox chk 
         Caption         =   "&4.打印空项"
         Height          =   240
         Index           =   3
         Left            =   255
         TabIndex        =   15
         Top             =   1395
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.CheckBox chk 
         Caption         =   "&2.打印总检"
         Height          =   240
         Index           =   2
         Left            =   255
         TabIndex        =   10
         Top             =   570
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.CheckBox chk 
         Caption         =   "&3.打印项目"
         Height          =   240
         Index           =   1
         Left            =   255
         TabIndex        =   11
         Top             =   900
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.CheckBox chk 
         Caption         =   "&1.打印封面"
         Height          =   240
         Index           =   0
         Left            =   255
         TabIndex        =   9
         Top             =   255
         Value           =   1  'Checked
         Width           =   1410
      End
   End
   Begin VB.Frame fra2 
      Height          =   3885
      Left            =   75
      TabIndex        =   12
      Top             =   2475
      Width           =   2595
      Begin VB.OptionButton opt 
         Caption         =   "&6.项目报告格式"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   2295
         Width           =   1590
      End
      Begin VB.OptionButton opt 
         Caption         =   "&5.项目统一格式"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   345
         Value           =   -1  'True
         Width           =   1620
      End
      Begin VB.ListBox lstStyle 
         Height          =   1530
         Left            =   480
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   675
         Width           =   2025
      End
   End
   Begin VB.Frame fra3 
      Height          =   630
      Left            =   4110
      TabIndex        =   3
      Top             =   5985
      Width           =   3315
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1980
         TabIndex        =   1
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton cmdMenu 
         Height          =   270
         Left            =   675
         Picture         =   "frmMedicalStationRptBook.frx":69E8
         Style           =   1  'Graphical
         TabIndex        =   4
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
         TabIndex        =   5
         Tag             =   "姓名"
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&7.姓名"
         Height          =   180
         Index           =   1
         Left            =   1020
         TabIndex        =   0
         Tag             =   "姓名"
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
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
Attribute VB_Name = "frmMedicalStationRptBook"
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
Private mlng病人id As Long
Private mblnDataMoved As Boolean

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

Public Function ShowEdit(ByVal frmMain As Object, ByRef lngKey As Long, Optional lng病人id As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:  frmMain         调用窗体对象
    '       lngKey          预约登记id
    '返回:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
    mlng病人id = lng病人id
    mlngKey = lngKey
    Set mfrmMain = frmMain
        
    If InitData = False Then Exit Function
    If ReadPersonData(mlngKey, lng病人id) = False Then Exit Function
    
    If Val(vsf(0).RowData(vsf(0).Row)) > 0 Then
        If ReadData(mlngKey, Val(vsf(0).RowData(vsf(0).Row))) = False Then Exit Function
    End If
    
    If opt(0).Value Then
        Call opt_Click(0)
    Else
        Call opt_Click(1)
    End If
    
    Dim lngLoop As Long
    Dim strTmp As String
    
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "报告书输出格式", "a")
    If strTmp <> "a" Then
        strTmp = ";" & strTmp & ";"
        For lngLoop = 0 To lstStyle.ListCount - 1
            If InStr(strTmp, ";" & lstStyle.List(lngLoop) & ";") > 0 Then
                lstStyle.Selected(lngLoop) = True
            Else
                lstStyle.Selected(lngLoop) = False
            End If
        Next
    End If
    
    
'    EditChanged = (Val(vsf(0).RowData(1)) > 0)

    Me.Show 1, frmMain
    
    ShowEdit = mblnOK

End Function

Private Function ReadPersonData(ByVal lng登记id As Long, Optional ByVal lng病人id As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  读取数据
    '参数:  lngKey      体检类型序号
    '返回:  True        读取成功
    '       False       读取失败
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand
            
          
    gstrSQL = _
            "Select 0 As 选择, B.姓名, B.门诊号, B.健康号,b.就诊卡号, b.身份证号,a.体检编号,B.病人id As ID,b.性别,'' As 清单 " & vbNewLine & _
            "From 体检人员档案 A, 病人信息 B" & vbNewLine & _
            "Where 体检报到 = 1 And A.病人id = B.病人id And A.登记id = [1] "
    
    If lng病人id > 0 Then
        gstrSQL = gstrSQL & " And a.病人id=[2] "
    End If
    
    Call ClearGrid(vsf(0))
    
    mblnDataMoved = DataMove(lng登记id)
    If mblnDataMoved Then
        gstrSQL = Replace(gstrSQL, "体检人员档案", "H体检人员档案")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng登记id, lng病人id)
    If rs.BOF = False Then
        Call FillGrid(vsf(0), rs)
'        Call LoadGrid(vsf(0), rs, , , ils13)
        
    End If
    vsf(0).AppendRow = True
    ReadPersonData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume

End Function

Private Function ShowItemSelect(ByVal lngRow As Long)
    Dim lngLoop As Long
    
    If vsf(0).TextMatrix(lngRow, 5) = "-1" Then
        '全部
        vsf(1).Cell(flexcpText, 1, 0, vsf(1).Rows - 1, 0) = 1
    ElseIf vsf(0).TextMatrix(lngRow, 5) = "" Then
        vsf(1).Cell(flexcpText, 1, 0, vsf(1).Rows - 1, 0) = 0
    Else
        For lngLoop = 1 To vsf(1).Rows - 1
            If InStr(vsf(0).TextMatrix(lngRow, 5), "," & vsf(1).RowData(lngLoop) & ",") > 0 Then
                vsf(1).TextMatrix(lngLoop, 0) = 1
            End If
        Next
    End If
        
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
            
    Call ClearGrid(vsf(1))
    
    gstrSQL = " Select x.*,0 As 选择," & _
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
                               "and c.登记ID =[1]  And c.病人id=[2] " & _
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
                                "and c.登记ID =[1] And c.病人id=[2] " & _
                       ") x, " & _
                      "部门表 y, " & _
                      "诊疗项目目录 z, " & _
                      "诊疗单据应用 d, " & _
                      "病人病历记录 h " & _
                "Where x.执行科室id = y.ID " & _
                      "and z.id = x.诊疗项目id " & _
                      "and x.报告id = h.id(+) " & _
                      "and d.应用场合(+)=4 " & _
                      "and x.诊疗项目id = d.诊疗项目id(+) Order By y.名称"
                      
    If mblnDataMoved Then
        gstrSQL = Replace(gstrSQL, "体检人员档案", "H体检人员档案")
        gstrSQL = Replace(gstrSQL, "体检项目清单", "H体检项目清单")
        gstrSQL = Replace(gstrSQL, "体检项目医嘱", "H体检项目医嘱")
        gstrSQL = Replace(gstrSQL, "病人医嘱记录", "H病人医嘱记录")
        gstrSQL = Replace(gstrSQL, "病人医嘱发送", "H病人医嘱发送")
        gstrSQL = Replace(gstrSQL, "病人病历记录", "H病人病历记录")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, lng病人id)
    If rs.BOF = False Then
        
        Call FillGrid(vsf(1), rs)
'        Call LoadGrid(vsf(1), rs, , , ils13)
        
    End If
    
    vsf(1).AppendRow = True
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
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    Me.Caption = "体检报告书"
    
    mnuFileRptGroup.Visible = False
    mnuFile_2.Visible = False
    
    With vsf(0)
        .Cols = 0
'        .NewColumn "", 240
        .NewColumn "选择", 450, 4, , 1, , flexDTBoolean
        .NewColumn "姓名", 900
        .NewColumn "性别", 600, 1
        .NewColumn "门诊号", 900, 7
        .NewColumn "健康号", 900, 7
        .NewColumn "就诊卡号", 0, 1
        .NewColumn "身份证号", 0, 1
        .NewColumn "体检编号", 990, 1
                
        .NewColumn "清单", 0
'        .FixedCols = 1
        .NewColumn "", 15
        .ExtendLastCol = True
        .SelectMode = True
        .Body.GridColor = COLOR.浅灰色
        .Body.GridColorFixed = COLOR.浅灰色
        .AppendRow = True
    End With
    
    With vsf(1)
        .Cols = 0
'        .NewColumn "", 240
        .NewColumn "选择", 450, 4, , 1, , flexDTBoolean
        .NewColumn "项目", 2400
        .NewColumn "执行科室", 1080
        .NewColumn "执行状态", 900
        .NewColumn "报告id", 0
        .NewColumn "单据id", 0
        .NewColumn "No", 0
        .NewColumn "", 15
'        .FixedCols = 1
        .ExtendLastCol = True
        .Body.GridColor = COLOR.浅灰色
        .Body.GridColorFixed = COLOR.浅灰色
        .SelectMode = True
        .AppendRow = True
    End With

    mnuFileRptGroup.Visible = (mlng病人id = 0)
    mnuFile_2.Visible = (mlng病人id = 0)

    
    lstStyle.Clear
    gstrSQL = "select a.说明 As 格式 from zlRPTFMTs a,zlreports b where a.报表id=b.id and  b.编号=[1] Order By a.序号"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "ZL1_BILL_1861_2")
    If rs.BOF = False Then
        Do While Not rs.EOF
            lstStyle.AddItem zlCommFun.NVL(rs("格式"))
            lstStyle.Selected(lstStyle.NewIndex) = True
            rs.MoveNext
        Loop
    End If
    
    
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
    
    
    strSQL = "SELECT DISTINCT 'ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2' AS 报表编号," & _
                       "D.NO," & _
                       "D.记录性质 " & _
                "FROM 病历文件目录 C,(SELECT A.NO,A.记录性质,E.病历文件id FROM 病人医嘱发送 A,病人医嘱记录 B,诊疗单据应用 E WHERE E.应用场合=4 AND E.诊疗项目id=B.诊疗项目id AND B.诊疗类别 IN ('C','D') AND A.医嘱id=B.ID AND (B.相关id=[1] OR B.ID=[1]) AND ROWNUM<2) D " & _
                "Where C.ID=D.病历文件id"

    If mblnDataMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
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
    Dim lng病人id As Long
    Dim intCount As Integer
    Dim lngRow As Long
    
    On Error GoTo errHand
    
    
    int选择 = 0
    
    If blnGroup Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1861_3", Me, "登记id=" & mlngKey, bytMode)
    Else
        
        For lngLoop = 1 To vsf(0).Rows - 1
        
            If Val(vsf(0).RowData(lngLoop)) > 0 And Abs(Val(vsf(0).TextMatrix(lngLoop, int选择))) = 1 Then
                
                 '1.调用"报告封面"
                If chk(0).Value = 1 Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1861_2_1", Me, "登记id=" & mlngKey, "病人id=" & Val(vsf(0).RowData(lngLoop)), bytMode)
                End If
                
                '2.调用"体检总检"
                If chk(2).Value = 1 Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1861_2_2", Me, "登记id=" & mlngKey, "病人id=" & Val(vsf(0).RowData(lngLoop)), bytMode)
                End If
                
                '3.统一报告书
                If opt(0).Value And chk(1).Value = 1 Then

                    gstrSQL = "zl_体检项目医嘱_Update(" & mlngKey & "," & Val(vsf(0).RowData(lngLoop)) & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                    
                    For intCount = 0 To lstStyle.ListCount - 1
                        If lstStyle.Selected(intCount) Then
                            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1861_2", Me, "登记id=" & mlngKey, "病人id=" & Val(vsf(0).RowData(lngLoop)), "空项=" & chk(3).Value, "REPORTFORMAT=" & (intCount + 1), bytMode)
                        End If
                    Next
                End If
                
                '4.调用"项目报告",缺省调用
                If opt(1).Value And chk(1).Value = 1 Then
                    
                    vsf(0).Row = lngLoop
                    Call vsf_AfterRowColChange(0, -1, -1, vsf(0).Row, vsf(0).Col)
                    
                    For lngRow = 1 To vsf(1).Rows - 1
                        
                        If Val(vsf(1).RowData(lngRow)) > 0 And Abs(Val(vsf(1).TextMatrix(lngRow, int选择))) = 1 Then
                        
                            If GetReportCode(Val(vsf(1).RowData(lngRow)), strReportCode, strReportParaNo, bytReportParaMode) Then
                                Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, bytMode)
                            End If
                            
                        End If
                    Next
                    
                End If
                
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
    
    On Error Resume Next
    opt(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "报告书格式", 0))).Value = True

        
    chk(0).Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "打印封面", 1))
    chk(1).Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "打印项目", 1))
    chk(2).Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "打印总检", 1))
    chk(3).Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "打印空项", 1))
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With fra1
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0) - 90
    End With
    
    With fra2
        .Left = fra1.Left
        .Top = fra1.Top + fra1.Height - 90
        .Width = fra1.Width
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With

    With vsf(0)
        .Left = fra1.Left + fra1.Width
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0) - fra3.Height + 90 - IIf(vsf(1).Visible, vsf(1).Height + 30, 0)
    End With
    
    vsf(1).Move vsf(0).Left, vsf(0).Top + vsf(0).Height + 30, vsf(0).Width
    
    If vsf(1).Visible Then
        fra3.Move vsf(1).Left, vsf(1).Top + vsf(1).Height - 90, vsf(1).Width
    Else
        fra3.Move vsf(0).Left, vsf(0).Top + vsf(0).Height - 90, vsf(0).Width
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lngLoop As Long
    Dim strTmp As String
    
    Call SaveWinState(Me, App.ProductName)
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "查找信息", lbl(1).Tag)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "打印封面", chk(0).Value)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "打印项目", chk(1).Value)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "打印总检", chk(2).Value)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "打印空项", chk(3).Value)
    
    If opt(0).Value Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "报告书格式", 0)
        
        For lngLoop = 0 To lstStyle.ListCount - 1
            If lstStyle.Selected(lngLoop) Then strTmp = strTmp & ";" & lstStyle.List(lngLoop)
        Next
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "报告书输出格式", strTmp)
        
    Else
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "报告书格式", 1)
    End If
    
End Sub


Private Sub mnuFileClearAll_Click()
    Dim lngLoop As Long
    Dim int选择 As Integer
    
    int选择 = 0
    If int选择 >= 0 Then
    
        For lngLoop = 1 To vsf(0).Rows - 1
            If Val(vsf(0).RowData(lngLoop)) > 0 Then
                vsf(0).TextMatrix(lngLoop, int选择) = 0
                vsf(0).TextMatrix(lngLoop, 5) = ""
            End If
        Next
        
        EditChanged = False
        
    End If
    Call ShowItemSelect(vsf(0).Row)
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


'Private Sub mnuFilePrintSetRpt_Click(Index As Integer)
'
'    Select Case Index
'    Case 0
'        ReportPrintSet gcnOracle, glngSys, "ZL1_BILL_1861_2_1"
'    Case 1
'        ReportPrintSet gcnOracle, glngSys, "ZL1_BILL_1861_2"
'    Case 2
'        ReportPrintSet gcnOracle, glngSys, "ZL1_BILL_1861_2_2"
'    End Select
'
'End Sub

Private Sub mnuFilePrintSet_Click()

    Call frmMedicalStationRptPrintSet.ShowEdit(Me, mlngKey, mlng病人id)
    
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
    
    int选择 = 0
    If int选择 >= 0 Then
        For lngLoop = 1 To vsf(0).Rows - 1
            If Val(vsf(0).RowData(lngLoop)) > 0 Then
                vsf(0).TextMatrix(lngLoop, int选择) = 1
                vsf(0).TextMatrix(lngLoop, 5) = "-1"
                EditChanged = True
            End If
        Next
    End If
    Call ShowItemSelect(vsf(0).Row)
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
        
        lbl(1).Caption = "&7." & Left(Trim(Caption), Len(Trim(Caption)) - 1)
        lbl(1).Tag = Left(Trim(Caption), Len(Trim(Caption)) - 1)
        
    End Select
End Sub

Private Sub opt_Click(Index As Integer)
    If Index = 0 Then
        vsf(1).Visible = False
    Else
        vsf(1).Visible = True
    End If
    
    Call Form_Resize
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
    lngCol = GetCol(vsf(0), strCol)
            
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
            lngCol = -1
            Select Case Mid(lbl(1).Caption, 4)
            Case "姓名", "姓名拼音", "姓名五笔"
                lngCol = 1
            Case "门诊号"
                lngCol = 3
            Case "健康号"
                lngCol = 4
            Case "就诊卡号"
                lngCol = 5
            Case "身份证号"
                lngCol = 6
            Case "体检编号"
                lngCol = 7
            End Select
            If lngCol < 0 Then Exit Sub
            
            lngRow = 0
            If vsf(0).Row + 1 <= vsf(0).Rows - 1 Then
                For lngLoop = vsf(0).Row + 1 To vsf(0).Rows - 1
                
                    lngRow = 0
                    Select Case strCol
                    Case "门诊号"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "健康号"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "就诊卡号"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "身份证号"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名拼音"
                        If zlGetSymbol(UCase(vsf(0).TextMatrix(lngLoop, lngCol))) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名五笔"
                        If zlGetSymbol(UCase(vsf(0).TextMatrix(lngLoop, lngCol)), 1) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case Else
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    End Select
            
                    If lngRow > 0 Then Exit For

                Next
            End If
            
            If lngRow = 0 Then
                For lngLoop = 1 To vsf(0).Row
                    lngRow = 0
                    Select Case strCol
                    Case "门诊号"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "健康号"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "就诊卡号"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "身份证号"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名"
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名拼音"
                        If zlGetSymbol(UCase(vsf(0).TextMatrix(lngLoop, lngCol))) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case "姓名五笔"
                        If zlGetSymbol(UCase(vsf(0).TextMatrix(lngLoop, lngCol)), 1) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    Case Else
                        If UCase(vsf(0).TextMatrix(lngLoop, lngCol)) = UCase(txt(Index).Text) Then lngRow = lngLoop
                    End Select
            
                    If lngRow > 0 Then Exit For
                Next
            End If
            
            If lngRow <= 0 Then
                ShowSimpleMsg "没有找到符合要求的信息！"
                txt(Index).Text = ""
            Else
                vsf(0).ShowCell lngRow, vsf(0).Col
                vsf(0).Row = lngRow
            End If
        End If
        
        txt(Index).SetFocus
        zlControl.TxtSelAll txt(Index)
    End If
End Sub


Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim lngLoop As Long
    
    If Index = 1 Then
        
        vsf(0).TextMatrix(vsf(0).Row, 5) = ""
        For lngLoop = 1 To vsf(1).Rows - 1
            If Abs(Val(vsf(1).TextMatrix(lngLoop, 0))) = 1 Then
                vsf(0).TextMatrix(vsf(0).Row, 5) = vsf(0).TextMatrix(vsf(0).Row, 5) & "," & Val(vsf(1).RowData(lngLoop))
            End If
        Next
        If vsf(0).TextMatrix(vsf(0).Row, 5) <> "" Then vsf(0).TextMatrix(vsf(0).Row, 5) = vsf(0).TextMatrix(vsf(0).Row, 5) & ","
    Else
        If Abs(Val(vsf(0).TextMatrix(Row, 0))) = 1 Then
            vsf(0).TextMatrix(Row, 5) = "-1"
        Else
            vsf(0).TextMatrix(Row, 5) = ""
        End If
        Call ShowItemSelect(Row)
    End If
    
    EditChanged = True
    
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngLoop As Long
    
    On Error Resume Next
    
    If Index = 0 Then
        Call ReadData(mlngKey, Val(vsf(0).RowData(vsf(0).Row)))
        Call ShowItemSelect(NewRow)
                
    End If
End Sub

Private Sub vsf_BeforeDeleteCell(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_BeforeDeleteRow(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(Index As Integer, ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Abs(Val(vsf(0).TextMatrix(vsf(0).Row, 0))) = 0 And Index = 1 Then
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

