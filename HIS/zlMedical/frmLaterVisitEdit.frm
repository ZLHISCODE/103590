VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLaterVisitEdit 
   Caption         =   "体检随访记录"
   ClientHeight    =   6645
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   10500
   Icon            =   "frmLaterVisitEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10500
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1410
      Left            =   0
      TabIndex        =   5
      Top             =   1275
      Width           =   6645
      _cx             =   11721
      _cy             =   2487
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
      BackColorSel    =   8388608
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12698049
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   20
      Cols            =   1
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
      Begin VB.Line lnY2 
         Index           =   0
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
      Begin VB.Line lnX2 
         Index           =   0
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
   End
   Begin VB.Frame fra1 
      Height          =   525
      Left            =   0
      TabIndex        =   6
      Top             =   2625
      Width           =   8385
      Begin VB.OptionButton opt 
         Caption         =   "&1.正常"
         Height          =   210
         Index           =   0
         Left            =   945
         TabIndex        =   8
         Top             =   210
         Value           =   -1  'True
         Width           =   840
      End
      Begin VB.OptionButton opt 
         Caption         =   "&2.观察"
         Height          =   210
         Index           =   1
         Left            =   1905
         TabIndex        =   9
         Top             =   210
         Width           =   840
      End
      Begin VB.OptionButton opt 
         Caption         =   "&3.复查"
         Height          =   210
         Index           =   2
         Left            =   2805
         TabIndex        =   10
         Top             =   210
         Width           =   840
      End
      Begin VB.OptionButton opt 
         Caption         =   "&4.治疗"
         Height          =   210
         Index           =   3
         Left            =   3690
         TabIndex        =   11
         Top             =   210
         Width           =   840
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "结果(&L)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   225
         Width           =   705
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   6285
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLaterVisitEdit.frx":076A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13467
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
      Left            =   9225
      Top             =   1545
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
            Picture         =   "frmLaterVisitEdit.frx":0FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisitEdit.frx":1218
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisitEdit.frx":1438
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisitEdit.frx":1652
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisitEdit.frx":1872
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   8625
      Top             =   1545
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
            Picture         =   "frmLaterVisitEdit.frx":1A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisitEdit.frx":1CAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisitEdit.frx":1EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisitEdit.frx":2218
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisitEdit.frx":2438
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   10500
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
         TabIndex        =   17
         Top             =   30
         Width           =   10380
         _ExtentX        =   18309
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
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&S.保存"
               Key             =   "保存"
               Object.ToolTipText     =   "保存(Alt+S)"
               Object.Tag             =   "&S.保存"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&R.重填"
               Key             =   "重填"
               Object.ToolTipText     =   "重填(Alt+R)"
               Object.Tag             =   "&R.重填"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助(Alt+H)"
               Object.Tag             =   "&H.帮助"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出(Alt+X)"
               Object.Tag             =   "&X.退出"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fra2 
      Height          =   2700
      Left            =   0
      TabIndex        =   12
      Top             =   3075
      Width           =   8385
      Begin RichTextLib.RichTextBox rtb 
         Height          =   2445
         Left            =   915
         TabIndex        =   14
         Top             =   165
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   4313
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         MaxLength       =   4000
         TextRTF         =   $"frmLaterVisitEdit.frx":2658
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "描述(&N)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   13
         Top             =   210
         Width           =   705
      End
   End
   Begin VB.Frame fra0 
      Height          =   570
      Left            =   405
      TabIndex        =   0
      Top             =   720
      Width           =   9555
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   3705
         MaxLength       =   20
         TabIndex        =   4
         Top             =   165
         Width           =   1305
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Left            =   1215
         TabIndex        =   2
         Top             =   165
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   75431939
         CurrentDate     =   38406
      End
      Begin VB.Label lblNo 
         AutoSize        =   -1  'True
         Caption         =   "12345678"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5520
         TabIndex        =   19
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "NO:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5190
         TabIndex        =   18
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "随访人(&M)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2775
         TabIndex        =   3
         Top             =   225
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "随访日期(&D)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   1095
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "重填(&R)"
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
Attribute VB_Name = "frmLaterVisitEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mblnOK As Boolean
Private mfrmMain As Object
Private mstrNo As String
Private mlng病人id As Long
Private mvarParam As Variant

Private Enum mCol
    类别 = 1
    项目
    编码
    单位
End Enum

'（２）自定义过程或函数************************************************************************************************
Private Function ShowOpenList(Optional strText As String) As Byte
    '------------------------------------------------------------------------------------------------------------------
    '功能:打开列表结构的诊疗检验标本数据
    '返回:出错返回2;成功返回1;取消返回0
    '------------------------------------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    strLvw = "编码,900,0,1;名称,1800,0,0;是否疾病,900,0,0"

    ShowOpenList = 2
    
    strSQL = _
                "SELECT 序号 AS ID, " & _
                        "A.编码, " & _
                        "A.名称, " & _
                        "Decode(是否疾病,1,'√','') As 是否疾病 " & _
                "FROM 体检诊断建议 A " & _
                "WHERE NVL(末级,0)=1 "
    strSQL = strSQL & " AND (A.编码 Like [1] OR A.名称 Like [1] OR A.简码 Like [1])"
            
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "%" & UCase(strText) & "%")
    If rs.BOF Then
        ShowOpenList = 0
        Exit Function
    End If
    
    If rs.RecordCount = 1 Then GoTo Over
        
    Call CalcPosition(sglX, sglY, vsf)
    
    If frmSelectDialog.ShowSelect(Me, 2, rs, strLvw, "请从下表中选择一个诊断", sglX + 60, sglY, 9000, 5100, 300, , Me.Name & "\体检结论过滤选择", , False) Then GoTo Over
    
    Exit Function
    
Over:
    vsf.RowData(vsf.Row) = 1
    vsf.EditText = zlCommFun.NVL(rs("名称").Value)
    vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
    vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
    
    ShowOpenList = 1
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function


Private Sub GoNextCell()
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngCol As Long
    Dim blnCancel As Boolean
    
    If GetAllowCol(vsf.Col + 1) > vsf.Cols - 1 Then
        '换行之前，先检查是否允许换行，即是否有必输的项目没有输入
                
        If vsf.Row = vsf.Rows - 1 Then
            blnCancel = False
            
            lngCol = 1
            
            If blnCancel Then
                vsf.Col = lngCol
                vsf.ShowCell vsf.Row, vsf.Col
                Exit Sub
            End If
            
            Call InsertNewRow
        Else
            vsf.Row = vsf.Row + 1
        End If
        
        '找第一个可以编辑的列
        vsf.Col = GetAllowCol(1)
    Else
        '找下一个可以编辑的列
        vsf.Col = GetAllowCol(vsf.Col + 1)
    End If
    
    vsf.ShowCell vsf.Row, vsf.Col
    
End Sub

Private Sub InsertNewRow()
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '------------------------------------------------------------------------------------------------------------------
    If vsf.Editable <> flexEDNone Then
        vsf.AddItem "", vsf.Rows
        vsf.Row = vsf.Rows - 1
    Else
        vsf.Row = vsf.Rows - 1
    End If
    
    Call AdjustRowFlag(vsf)
    
End Sub

Private Function GetAllowCol(ByVal lngFromCol As Long) As Long
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim lngLoop As Long
    
    lngRow = vsf.Row
    
    For lngLoop = lngFromCol To vsf.Cols - 1
        If lngLoop = 3 Then Exit For
    Next
    
    GetAllowCol = lngLoop
End Function

Private Sub AdjustRowFlag(ByRef objVsf As Object, Optional ByVal intRow As Integer)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '------------------------------------------------------------------------------------------------------------------
    If objVsf.FixedCols = 0 Then Exit Sub
    
    Dim lngLoop As Long
    
    For lngLoop = 0 To vsf.Rows - 1
        vsf.TextMatrix(lngLoop, 0) = lngLoop + 1 & "、"
    Next
End Sub

Private Sub ShowSelectRow(ByRef objVsf As Object, Optional ByVal intRow As Integer)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '------------------------------------------------------------------------------------------------------------------
    vsf.Cell(flexcpFontBold, 0, 0, vsf.Rows - 1, vsf.Cols - 1) = False
    vsf.Cell(flexcpFontBold, intRow, 0, intRow, vsf.Cols - 1) = True
    
End Sub

Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '值域:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
    
    mnuFileSave.Enabled = True
    mnuFileRestore.Enabled = True

    If vData = False Then
        mnuFileSave.Enabled = False
        mnuFileRestore.Enabled = False

    End If
    
    tbrThis.Buttons("保存").Enabled = mnuFileSave.Enabled
    tbrThis.Buttons("重填").Enabled = mnuFileRestore.Enabled
        
End Property

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    On Error Resume Next
    
    vsf.Rows = 1
    vsf.Cell(flexcpText, 0, 0, vsf.Rows - 1, vsf.Cols - 1) = ""
    vsf.RowData(0) = 0
    vsf.TextMatrix(0, 0) = "1、"
    
    Call ReadRow(vsf.Row)
    
    On Error GoTo 0
    
    EditChanged = True
    
        
End Function

Public Function ShowEdit(ByVal frmMain As Object, ByVal strParam As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:  frmMain         调用窗体对象
    '       lngKey          预约登记id
    '返回:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    Dim varGroup As Variant
    Dim lngLoop As Long
    
    mblnStartUp = True
    mblnOK = False
    
    '病人id,体检单号,随访单号
    mvarParam = Split(strParam, "'")
    
    mstrNo = mvarParam(2)
    mlng病人id = Val(mvarParam(0))
    
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    EditChanged = False
    
    If ReadData = False Then Exit Function
    If mstrNo <> "" Then EditChanged = False
    Call ShowSelectRow(vsf, 0)
    
    'stbThis.Panels(2).Text = "填写/选择体检人员资料。"
                
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    If mstrNo = "" Then
                        
        lblNo.Caption = GetNextNo(79)
        
        '首先从上次随访收集
        gstrSQL = "select A.随访时间," & _
                        "'" & UserInfo.姓名 & "' as 随访人," & _
                        "'" & lblNo.Caption & "' as no," & _
                        "A.诊断描述," & _
                        "0 AS 随访结果," & _
                        "'' AS 随访情况 " & _
                    "from 体检随访记录 A,体检人员档案 B,体检登记记录 C " & _
                    "Where A.病人ID = B.病人ID " & _
                        "AND A.体检单号=C.体检号 " & _
                        "AND A.随访时间=B.随访时间 " & _
                        "AND A.随访结果<>1 " & _
                        "AND B.登记id=C.ID " & _
                        "AND B.病人id=[1] " & _
                        "AND C.体检号=[2] " & _
                    "order by A.序号"
                        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病人id, CStr(mvarParam(1)))
        If rs.BOF = False Then
                
        Else
            
            '再从体检结果收集
                        
            gstrSQL = _
                "select TRUNC(sysdate) as 随访时间," & _
                       "'" & UserInfo.姓名 & "' as 随访人," & _
                       "'" & lblNo.Caption & "' as no," & _
                       "A.结论描述 as 诊断描述," & _
                       "0 as 随访结果," & _
                       "'' as 随访情况 " & _
                "from 体检人员结论 A," & _
                     "(SELECT 体检病历ID FROM 体检人员档案 U,体检登记记录 T WHERE ROWNUM<2 AND U.登记id=T.ID AND U.体检状态=5 AND U.病人id=" & mlng病人id & " AND T.体检号='" & mvarParam(1) & "' AND U.完成时间=(SELECT MAX(X.完成时间) FROM  体检人员档案 X,体检登记记录 Y WHERE X.登记id=Y.ID AND X.病人id=" & mlng病人id & " AND Y.体检号='" & mvarParam(1) & "')) B," & _
                     "病人病历内容 C " & _
                "Where B.体检病历ID = C.病历记录ID " & _
                      "AND C.ID=A.病历id " & _
                      "AND A.记录性质=0 " & _
                "ORDER BY A.记录序号"
        End If
    Else
        '从本单据收集
        
        gstrSQL = "SELECT * FROM 体检随访记录 WHERE NO=[1] Order by 序号"
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrNo)
    If rs.BOF = False Then
        
        dtp.Value = Format(zlCommFun.NVL(rs("随访时间")), dtp.CustomFormat)
        txt.Text = zlCommFun.NVL(rs("随访人"))
        lblNo.Caption = zlCommFun.NVL(rs("NO"))
        
        Do While Not rs.EOF
            
            If Val(vsf.RowData(vsf.Rows - 1)) = 1 Then
                vsf.Rows = vsf.Rows + 1
            End If
            
            vsf.RowData(vsf.Rows - 1) = 1
            vsf.TextMatrix(vsf.Rows - 1, 0) = vsf.Rows & "、"
            vsf.TextMatrix(vsf.Rows - 1, 1) = zlCommFun.NVL(rs("随访结果"), 0)
            vsf.TextMatrix(vsf.Rows - 1, 2) = zlCommFun.NVL(rs("随访情况"))
            vsf.TextMatrix(vsf.Rows - 1, 3) = zlCommFun.NVL(rs("诊断描述"))
            
            rs.MoveNext
        Loop
    End If
    
    Call ReadRow(vsf.Row)

            
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
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    vsf.FixedRows = 0
    vsf.FixedCols = 1
    vsf.Cols = 4
    vsf.Rows = 1
    vsf.ColWidth(0) = 300
    
    vsf.ComboList = "..."
    
    vsf.ColHidden(1) = True         '保存结果
    vsf.ColHidden(2) = True         '保存描述
    
    vsf.TextMatrix(0, 0) = "1、"
    vsf.BackColorFixed = vsf.BackColor
    vsf.GridLines = flexGridNone
    vsf.GridLinesFixed = flexGridNone
                        
    vsf.Editable = flexEDKbdMouse
    
    dtp.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    txt.Text = UserInfo.姓名
    
    lblNo.Caption = ""
    
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
    Dim rs As New ADODB.Recordset
    
    For lngLoop = 0 To vsf.Rows - 1
        
        If StrIsValid(vsf.TextMatrix(lngLoop, 3), 100) = False Then
            vsf.Row = lngLoop
            vsf.Col = 3
            Call vsf.ShowCell(vsf.Row, vsf.Col)
            Exit Function
        End If
                
    Next
    
    If StrIsValid(rtb.Text, 4000) = False Then
        rtb.SetFocus
        Exit Function
    End If
                                                                
    ValidEdit = True
    
End Function

Private Function SaveEdit(ByRef lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  保存数据
    '返回:  True        保存成功
    '       False       保存失败
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim strNow As String
    Dim rsPati As New ADODB.Recordset
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    strSQL(ReDimArray(strSQL)) = "ZL_体检随访记录_DELETE('" & lblNo.Caption & "')"
    For lngLoop = 0 To vsf.Rows - 1
        If Trim(vsf.TextMatrix(lngLoop, 3)) <> "" Then
            strSQL(ReDimArray(strSQL)) = "ZL_体检随访记录_INSERT(" & mlng病人id & ",'" & mvarParam(1) & "','" & lblNo.Caption & "'," & lngLoop + 1 & ",'" & Trim(vsf.TextMatrix(lngLoop, 3)) & "'," & Val(vsf.TextMatrix(lngLoop, 1)) & ",'" & vsf.TextMatrix(lngLoop, 2) & "','" & txt.Text & "',to_date('" & Format(dtp.Value, "yyyy-MM-dd") & "','yyyy-mm-dd hh24:mi:ss'),to_date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
        End If
    Next
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveEdit = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    
End Function


Private Sub dtp_Change()
    EditChanged = True
End Sub

Private Sub dtp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyS
            If tbrThis.Buttons("保存").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("保存"))
        Case vbKeyR
            If tbrThis.Buttons("重填").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("重填"))
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
    
    
    With fra0
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0) - 75
        .Width = Me.ScaleWidth - .Left
    End With
    
    With vsf
        .Left = 0
        .Top = fra0.Top + fra0.Height + 15
        .Width = Me.ScaleWidth - .Left
    End With
    
    With fra1
        .Left = 0
        .Top = vsf.Top + vsf.Height - 60
        .Width = vsf.Width
    End With
    
    With fra2
        .Left = 0
        .Top = fra1.Top + fra1.Height - 90
        .Width = vsf.Width
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top
    End With
    
    With rtb
        .Top = 150
        .Width = fra2.Width - .Left - 75
        .Height = fra2.Height - .Top - 75
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mnuFileSave.Enabled Then
        Cancel = (MsgBox("数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileRestore_Click()
        
    If MsgBox("确实要恢复保存前的内容吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Call ClearData
    Call ReadData
    
    EditChanged = False
    
End Sub

Private Sub mnuFileSave_Click()
    Dim lngKey As Long
    
    Call WriteRow(vsf.Row)
        
    If ValidEdit = False Then Exit Sub
    
    If SaveEdit(lngKey) Then
        EditChanged = False
        mblnOK = True
        
        mstrNo = lblNo.Caption
        
        On Error Resume Next
        Call mfrmMain.EditRefresh("随访记录", lblNo.Caption)
        
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

Private Sub opt_Click(Index As Integer)
    EditChanged = True
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub rtb_Change()
    EditChanged = True
End Sub

Private Sub rtb_GotFocus()
    zlCommFun.OpenIme True
End Sub

Private Sub rtb_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub rtb_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(rtb.Text, rtb.MaxLength)
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "保存"
        Call mnuFileSave_Click
    Case "重填"
        Call mnuFileRestore_Click
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub txt_Change()
    EditChanged = True
End Sub

Private Sub WriteRow(ByVal lngRow As Long)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    Dim blnSvr As Boolean
    
    blnSvr = mnuFileSave.Enabled
    
    If opt(0).Value Then
        vsf.TextMatrix(lngRow, 1) = "1"
    ElseIf opt(1).Value Then
        vsf.TextMatrix(lngRow, 1) = "2"
    ElseIf opt(2).Value Then
        vsf.TextMatrix(lngRow, 1) = "3"
    ElseIf opt(3).Value Then
        vsf.TextMatrix(lngRow, 1) = "4"
    Else
        vsf.TextMatrix(lngRow, 1) = ""
    End If
    
    vsf.TextMatrix(lngRow, 2) = rtb.Text
    
    EditChanged = blnSvr
End Sub

Private Sub ReadRow(ByVal lngRow As Long)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    Dim blnSvr As Boolean
    
    blnSvr = mnuFileSave.Enabled
    
    If Val(vsf.TextMatrix(lngRow, 1)) >= 1 And Val(vsf.TextMatrix(lngRow, 1)) <= 4 Then
        opt(Val(vsf.TextMatrix(lngRow, 1)) - 1).Value = True
    Else
        opt(0).Value = True
    End If
    
    rtb.Text = vsf.TextMatrix(lngRow, 2)
    
    EditChanged = blnSvr
End Sub

Private Sub txt_GotFocus()
    
    zlCommFun.OpenIme True
    
    zlControl.TxtSelAll txt
    
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        vsf.Col = 3
        vsf.SetFocus
    End If
End Sub

Private Sub txt_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    EditChanged = True
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    
    If OldRow <> NewRow Then
        Call ShowSelectRow(vsf, NewRow)
    End If
    vsf.ComboList = "..."
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
     
    On Error Resume Next
    
    If OldRow <> NewRow Then
        Call WriteRow(OldRow)
        Call ReadRow(NewRow)
    End If
    
    
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    gstrSQL = "SELECT -1 AS ID," & _
                        "0 AS 上级ID," & _
                        "0 AS 末级," & _
                        "'' AS 编码," & _
                        "'所有分类' AS 名称, " & _
                        "'' AS 疾病 " & _
                "FROM dual "
                
    gstrSQL = gstrSQL & _
            " UNION ALL " & _
            "SELECT 序号 AS ID," & _
                        "DECODE(上级序号,NULL,-1,上级序号) AS 上级ID," & _
                        "0 AS 末级," & _
                        "编码," & _
                        "名称, " & _
                        "'' AS 疾病 " & _
                "FROM 体检诊断建议 " & _
                "WHERE NVL(末级,0)=0 " & _
                "START WITH 上级序号 is NULL CONNECT BY PRIOR 序号 = 上级序号 "
    
    gstrSQL = gstrSQL & _
                "UNION ALL " & _
                "SELECT 序号 AS ID, " & _
                        "DECODE(上级序号,NULL,-1,上级序号) AS 上级ID, " & _
                        "1 AS 末级, " & _
                        "A.编码, " & _
                        "A.名称, " & _
                        "DECODE(是否疾病,1,'√','是') AS 疾病 " & _
                "FROM 体检诊断建议 A " & _
                "WHERE NVL(末级,0)=1 "
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If ShowGrdSelect(Me, vsf, "编码,900,0,1;名称,1800,0,0;疾病,900,0,0", Me.Name & "\体检结论选择", "请从列表中选择一个结论。", rsData, rs, 9000, 5100) Then

        
        vsf.RowData(vsf.Row) = 1
        vsf.EditText = zlCommFun.NVL(rs("名称").Value)
        vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
        vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
        
        EditChanged = True
        
    End If
        
End Sub

Private Sub vsf_DblClick()

    Call vsf_KeyPress(32)
    
End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngLoop As Long
    Dim blnCancel As Boolean
    
    On Error GoTo errHand
    
    Select Case KeyCode
    Case vbKeyDelete
        
        If Shift = 0 And vsf.Editable <> flexEDNone Then
            '删除整行及内容
            
            If vsf.Rows > 0 Then
                If vsf.Rows = 1 And vsf.Row = 0 Then
                    For lngLoop = 0 To vsf.Cols - 1
                        vsf.TextMatrix(0, lngLoop) = ""
                    Next
                    vsf.RowData(0) = ""
                Else
                    vsf.RemoveItem vsf.Row
                    
                End If
                Call AdjustRowFlag(vsf, vsf.Row)
                
            End If
            
        End If
        
        If Shift = 1 And vsf.Editable <> flexEDNone And vsf.Col = 3 Then
            '删除当前单元格的内容
            
            vsf.TextMatrix(vsf.Row, vsf.Col) = ""
            
        End If
    End Select
    
    Exit Sub
    
errHand:
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim strSvrText As String
    
    If KeyCode = vbKeyReturn Then
        '对于2-文字型的情况
        
        If InStr(vsf.EditText, "'") > 0 Then
            KeyCode = 0
            Exit Sub
        End If

        strSvrText = vsf.EditText
        Select Case ShowOpenList(vsf.EditText)
        Case 2
            '取消了本次选择
            KeyCode = 0
            
            vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
            vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)

        End Select
    End If
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If Trim(vsf.TextMatrix(vsf.Row, vsf.Col)) = "" Then
            zlCommFun.PressKey vbKeyTab
        Else
            Call GoNextCell
        End If
    Else
        If vsf.ComboList = "..." Then vsf.ComboList = ""
    End If
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call GoNextCell
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

